const statusBadge = document.getElementById('statusBadge');
const statusText = document.getElementById('statusText');
const startButton = document.getElementById('startButton');
const stopButton = document.getElementById('stopButton');
const clearButton = document.getElementById('clearButton');
const stepList = document.getElementById('stepList');

let mediaRecorder = null;
let mediaStream = null;
let recordedChunks = [];

startButton.addEventListener('click', startRecording);
stopButton.addEventListener('click', stopRecording);
clearButton.addEventListener('click', clearRecording);

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== 'local') return;
  if (changes.recording || changes.steps || changes.last_error) refresh();
});

refresh();

async function refresh() {
  const state = await chrome.storage.local.get({ recording: false, steps: [], last_error: '' });
  statusBadge.textContent = state.recording ? '記録中' : state.last_error ? 'エラー' : '停止中';
  statusBadge.classList.toggle('recording', state.recording);
  statusText.textContent = state.last_error || `現在 ${state.steps.length} ステップです。`;
  startButton.disabled = state.recording;
  stopButton.disabled = !state.recording;
  clearButton.disabled = false;
  renderSteps(state.steps);
}

async function startRecording() {
  setBusy(true);
  try {
    const tab = await getActiveTab();
    validateTab(tab);
    recordedChunks = [];
    mediaStream = await navigator.mediaDevices.getDisplayMedia({
      video: true,
      audio: false
    });

    const mimeType = getSupportedMimeType();
    mediaRecorder = new MediaRecorder(mediaStream, mimeType ? { mimeType } : undefined);
    mediaRecorder.addEventListener('dataavailable', (event) => {
      if (event.data && event.data.size > 0) {
        recordedChunks.push(event.data);
      }
    });
    mediaStream.getVideoTracks()[0].addEventListener('ended', () => {
      if (mediaRecorder && mediaRecorder.state === 'recording') {
        stopRecording();
      }
    });
    mediaRecorder.start(250);

    const response = await chrome.runtime.sendMessage({
      type: 'manual-begin-recording',
      tabId: tab.id
    });

    if (!response || !response.ok) {
      stopLocalRecording();
      showError(response && response.error ? response.error : '記録を開始できませんでした。');
    }
  } catch (error) {
    stopLocalRecording();
    showError(error.message || '記録を開始できませんでした。');
  } finally {
    await refresh();
  }
}

async function stopRecording() {
  setBusy(true);
  try {
    const dataUrl = await stopLocalRecording();
    const response = await chrome.runtime.sendMessage({
      type: 'manual-finish-recording',
      dataUrl
    });

    if (!response || !response.ok) {
      showError(response && response.error ? response.error : '記録を保存できませんでした。');
    }
  } catch (error) {
    showError(error.message || '記録を保存できませんでした。');
  } finally {
    await refresh();
  }
}

async function clearRecording() {
  stopLocalRecording();
  await chrome.storage.local.set({
    recording: false,
    steps: [],
    recording_started_at: null,
    recording_tab_id: null,
    last_error: ''
  });
  await refresh();
}

function stopLocalRecording() {
  return new Promise((resolve) => {
    if (!mediaRecorder || mediaRecorder.state === 'inactive') {
      cleanupStream();
      resolve('');
      return;
    }

    mediaRecorder.addEventListener(
      'stop',
      async () => {
        const blob = new Blob(recordedChunks, { type: mediaRecorder.mimeType || 'video/webm' });
        const dataUrl = blob.size > 0 ? await blobToDataUrl(blob) : '';
        cleanupStream();
        resolve(dataUrl);
      },
      { once: true }
    );
    mediaRecorder.stop();
  });
}

function cleanupStream() {
  if (mediaStream) {
    mediaStream.getTracks().forEach((track) => track.stop());
  }
  mediaRecorder = null;
  mediaStream = null;
  recordedChunks = [];
}

function getActiveTab() {
  return new Promise((resolve, reject) => {
    chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
      const error = chrome.runtime.lastError;
      if (error) {
        reject(new Error(error.message));
        return;
      }
      resolve(tabs[0]);
    });
  });
}

function validateTab(tab) {
  if (!tab || typeof tab.id !== 'number') {
    throw new Error('記録対象のタブを取得できませんでした。');
  }
  if (!tab.url || /^chrome:|^edge:|^chrome-extension:/.test(tab.url)) {
    throw new Error('Chrome内部ページや拡張機能ページは記録できません。通常のWebページで開始してください。');
  }
}

function getSupportedMimeType() {
  const types = ['video/webm;codecs=vp9', 'video/webm;codecs=vp8', 'video/webm'];
  return types.find((type) => MediaRecorder.isTypeSupported(type)) || '';
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener('load', () => resolve(reader.result));
    reader.addEventListener('error', () => reject(reader.error));
    reader.readAsDataURL(blob);
  });
}

function setBusy(busy) {
  startButton.disabled = busy;
  stopButton.disabled = busy;
  clearButton.disabled = busy;
}

function showError(message) {
  statusBadge.textContent = 'エラー';
  statusBadge.classList.remove('recording');
  statusText.textContent = message;
}

function renderSteps(steps) {
  stepList.replaceChildren();

  for (const step of steps) {
    const item = document.createElement('li');
    const title = document.createElement('span');
    const meta = document.createElement('span');

    title.className = 'step-title';
    title.textContent = `Step ${step.step_no}: ${step.element_text || 'クリック'}`;
    meta.className = 'step-meta';
    meta.textContent = step.page_title || step.url || '';

    item.append(title, meta);
    stepList.appendChild(item);
  }
}
