(() => {
  'use strict';

  let recording = false;

  chrome.storage.local.get({ recording: false }, (state) => {
    recording = Boolean(state.recording);
  });

  chrome.storage.onChanged.addListener((changes, areaName) => {
    if (areaName !== 'local' || !changes.recording) return;
    recording = Boolean(changes.recording.newValue);
  });

  document.addEventListener(
    'mousedown',
    (event) => {
      if (!recording || event.button !== 0) return;

      const target = event.target;
      const element = getElement(target);
      const metadata = getElementMetadata(element);

      chrome.runtime.sendMessage({
        type: 'manual-click',
        payload: {
          ...metadata,
          url: location.href,
          page_title: document.title,
          x: Math.round(event.clientX),
          y: Math.round(event.clientY),
          viewport_width: window.innerWidth,
          viewport_height: window.innerHeight,
          timestamp_ms: Math.round(performance.timeOrigin + event.timeStamp),
          captured_at: new Date().toISOString()
        }
      });
    },
    true
  );

  function getElement(target) {
    if (!target || target === document) return null;
    return target.nodeType === Node.ELEMENT_NODE ? target : target.parentElement;
  }

  function getElementMetadata(element) {
    if (!element) {
      return {
        tag_name: '',
        element_text: '',
        aria_label: '',
        placeholder: '',
        role: '',
        column_header: ''
      };
    }

    return {
      tag_name: element.tagName,
      element_text: getElementText(element),
      aria_label: element.getAttribute('aria-label') || '',
      placeholder: element.getAttribute('placeholder') || '',
      role: element.getAttribute('role') || inferRole(element),
      column_header: getColumnHeader(element)
    };
  }

  function getElementText(element) {
    const candidates = [
      element.getAttribute('aria-label'),
      element.getAttribute('title'),
      element.getAttribute('alt'),
      element.value,
      element.textContent
    ];

    const text = candidates.find((value) => typeof value === 'string' && value.trim());
    return text ? text.trim().replace(/\s+/g, ' ').slice(0, 80) : '';
  }

  function inferRole(element) {
    if (element.tagName === 'BUTTON') return 'button';
    if (element.tagName === 'A') return 'link';
    if (element.tagName === 'INPUT' || element.tagName === 'TEXTAREA' || element.tagName === 'SELECT') return 'input';
    return '';
  }

  function getColumnHeader(element) {
    const cell = element.closest('td,th,[role="gridcell"],[role="columnheader"]');
    if (!cell) return '';

    if (cell.getAttribute('role') === 'columnheader' || cell.tagName === 'TH') {
      return (cell.textContent || '').trim().replace(/\s+/g, ' ').slice(0, 80);
    }

    const table = cell.closest('table');
    if (!table || typeof cell.cellIndex !== 'number' || cell.cellIndex < 0) return '';

    const header = table.querySelector(`thead th:nth-child(${cell.cellIndex + 1})`);
    return header ? header.textContent.trim().replace(/\s+/g, ' ').slice(0, 80) : '';
  }
})();
