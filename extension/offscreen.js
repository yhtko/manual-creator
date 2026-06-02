let recorder = null;
let stream = null;
let chunks = [];

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || !message.type) return false;

  if (message.type === 'offscreen-start-recording') {
    startRecording(message.streamId, message.source || 'desktop')
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  if (message.type === 'offscreen-stop-recording') {
    stopRecording()
      .then((result) => sendResponse({ ok: true, ...result }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  return false;
});

async function startRecording(streamId, source) {
  if (!streamId) {
    throw new Error('録画用ストリームIDが空です。');
  }

  if (recorder && recorder.state !== 'inactive') {
    recorder.stop();
  }

  chunks = [];
  stream = await navigator.mediaDevices.getUserMedia({
    audio: false,
    video: {
      mandatory: {
        chromeMediaSource: source,
        chromeMediaSourceId: streamId
      }
    }
  });

  const mimeType = getSupportedMimeType();
  recorder = new MediaRecorder(stream, mimeType ? { mimeType } : undefined);
  recorder.addEventListener('dataavailable', (event) => {
    if (event.data && event.data.size > 0) {
      chunks.push(event.data);
    }
  });
  recorder.start(250);
}

function stopRecording() {
  return new Promise((resolve, reject) => {
    if (!recorder || recorder.state === 'inactive') {
      cleanup();
      resolve({ dataUrl: null, mimeType: '' });
      return;
    }

    recorder.addEventListener(
      'stop',
      async () => {
        try {
          const mimeType = recorder.mimeType || 'video/webm';
          const blob = new Blob(chunks, { type: mimeType });
          const dataUrl = await blobToDataUrl(blob);
          cleanup();
          resolve({ dataUrl, mimeType });
        } catch (error) {
          cleanup();
          reject(error);
        }
      },
      { once: true }
    );

    recorder.stop();
  });
}

function cleanup() {
  if (stream) {
    stream.getTracks().forEach((track) => track.stop());
  }
  recorder = null;
  stream = null;
  chunks = [];
}

function getSupportedMimeType() {
  const types = [
    'video/mp4;codecs=avc1.42E01E',
    'video/mp4;codecs=h264',
    'video/mp4',
    'video/webm;codecs=vp9',
    'video/webm;codecs=vp8',
    'video/webm'
  ];
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
