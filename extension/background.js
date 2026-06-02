const STORAGE_DEFAULTS = {
  recording: false,
  steps: [],
  recording_started_at: null,
  recording_tab_id: null,
  last_error: ''
};

chrome.runtime.onInstalled.addListener(() => {
  if (!chrome.sidePanel) return;
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || !message.type) return false;

  if (message.type === 'manual-begin-recording') {
    beginRecording(message)
      .then((result) => sendResponse({ ok: true, ...result }))
      .catch(async (error) => {
        await chrome.storage.local.set({ recording: false, last_error: error.message });
        sendResponse({ ok: false, error: error.message });
      });
    return true;
  }

  if (message.type === 'manual-finish-recording') {
    finishRecording(message)
      .then((result) => sendResponse({ ok: true, ...result }))
      .catch(async (error) => {
        await chrome.storage.local.set({ recording: false, last_error: error.message });
        sendResponse({ ok: false, error: error.message });
      });
    return true;
  }

  if (message.type === 'manual-click') {
    saveStep(message.payload, sender.tab)
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  return false;
});

async function beginRecording(message) {
  if (typeof message.tabId !== 'number') {
    throw new Error('記録対象のタブを取得できませんでした。');
  }

  const startedAt = Date.now();
  await chrome.storage.local.set({
    recording: true,
    steps: [],
    recording_started_at: startedAt,
    recording_tab_id: message.tabId,
    last_error: ''
  });

  return { startedAt };
}

async function finishRecording(message) {
  const state = await chrome.storage.local.get(STORAGE_DEFAULTS);
  await chrome.storage.local.set({ recording: false });

  const fileId = createFileId(state.recording_started_at || Date.now());
  const eventsFileName = `events-${fileId}.json`;
  const recordingFileName = `recording-${fileId}${getVideoExtension(message.mimeType || message.dataUrl)}`;
  const steps = state.steps.map((step, index) => ({
    ...step,
    step_no: index + 1,
    image: `step_${String(index + 1).padStart(3, '0')}.png`
  }));

  const project = {
    title: '作業手順書',
    capture_mode: 'video',
    recording_started_at: state.recording_started_at,
    recording_file: recordingFileName,
    recording_mime_type: message.mimeType || '',
    steps
  };

  await downloadText(eventsFileName, JSON.stringify(project, null, 2), 'application/json');

  if (message.dataUrl) {
    await chrome.downloads.download({
      url: message.dataUrl,
      filename: `manual-project/${recordingFileName}`,
      saveAs: false
    });
  }

  await chrome.storage.local.set({
    steps,
    recording_started_at: null,
    recording_tab_id: null,
    last_error: ''
  });

  return {
    stepCount: steps.length,
    recordingSaved: Boolean(message.dataUrl),
    eventsFileName,
    recordingFileName
  };
}

async function saveStep(payload, tab) {
  const state = await chrome.storage.local.get(STORAGE_DEFAULTS);
  if (!state.recording) return;
  if (state.recording_tab_id && tab && tab.id !== state.recording_tab_id) return;

  const stepNo = state.steps.length + 1;
  const step = {
    step_no: stepNo,
    description: createDefaultDescription(payload),
    image: `step_${String(stepNo).padStart(3, '0')}.png`,
    url: payload.url,
    page_title: payload.page_title,
    element_text: payload.element_text,
    tag_name: payload.tag_name,
    aria_label: payload.aria_label,
    placeholder: payload.placeholder,
    role: payload.role,
    column_header: payload.column_header,
    x: payload.x,
    y: payload.y,
    viewport_width: payload.viewport_width,
    viewport_height: payload.viewport_height,
    timestamp_ms: payload.timestamp_ms,
    captured_at: payload.captured_at
  };

  await chrome.storage.local.set({ steps: [...state.steps, step] });
}

function downloadText(filename, text, mimeType) {
  const dataUrl = `data:${mimeType};charset=utf-8,${encodeURIComponent(text)}`;
  return chrome.downloads.download({
    url: dataUrl,
    filename: `manual-project/${filename}`,
    saveAs: false
  });
}

function createFileId(timestamp) {
  const date = new Date(Number(timestamp));
  const parts = [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, '0'),
    String(date.getDate()).padStart(2, '0'),
    '-',
    String(date.getHours()).padStart(2, '0'),
    String(date.getMinutes()).padStart(2, '0'),
    String(date.getSeconds()).padStart(2, '0')
  ];
  return parts.join('');
}

function getVideoExtension(value) {
  const text = String(value || '').toLowerCase();
  if (text.includes('video/mp4')) return '.mp4';
  return '.webm';
}

function createDefaultDescription(payload) {
  const label = payload.aria_label || payload.element_text || payload.placeholder || payload.role || '対象';
  const suffix = payload.role === 'button' || payload.tag_name === 'BUTTON' ? 'をクリックします' : 'を選択します';
  return `${label}${suffix}`;
}
