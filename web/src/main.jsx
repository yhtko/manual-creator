import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { Document, HeadingLevel, ImageRun, Packer, Paragraph, TextRun } from 'docx';
import { saveAs } from 'file-saver';
import pptxgen from 'pptxgenjs';
import './styles.css';

const DEFAULT_PROJECT = {
  title: '作業手順書',
  purpose: '',
  audience: '',
  prerequisites: '',
  completion: '',
  capture_mode: '',
  recording_started_at: null,
  recording_file: '',
  steps: []
};

function App() {
  const [project, setProject] = useState(DEFAULT_PROJECT);
  const [images, setImages] = useState({});
  const [recordingDataUrl, setRecordingDataUrl] = useState('');
  const [recordingDataUrls, setRecordingDataUrls] = useState({});
  const [videoStatus, setVideoStatus] = useState('');
  const [slidesPerPage, setSlidesPerPage] = useState(1);

  const orderedSteps = useMemo(
    () => project.steps.map((step, index) => ({ ...step, step_no: index + 1 })),
    [project.steps]
  );

  function updateProjectTitle(value) {
    setProject((current) => ({ ...current, title: value }));
  }

  function updateProjectInfo(field, value) {
    setProject((current) => ({ ...current, [field]: value }));
  }

  async function importProject(file) {
    if (!file) return;
    const text = await file.text();
    const data = JSON.parse(text);
    const normalized = normalizeProject(data, 'recording_1');
    setProject(normalized);
    setRecordingDataUrl('');
    setRecordingDataUrls({});
    setVideoStatus(
      normalized.recording_started_at
        ? `続けて${normalized.recording_file || 'recording-*.webm'}を読み込んでください。`
        : ''
    );
  }

  async function appendProject(file) {
    if (!file) return;
    const text = await file.text();
    const data = JSON.parse(text);
    const recordingId = `recording_${getRecordingIds(project.steps).length + 1}`;
    const appended = normalizeProject(data, recordingId);
    setProject((current) => ({
      ...current,
      steps: [...current.steps, ...appended.steps],
      purpose: current.purpose || appended.purpose,
      audience: current.audience || appended.audience,
      prerequisites: current.prerequisites || appended.prerequisites,
      completion: current.completion || appended.completion
    }));
    setVideoStatus('追加した記録のrecording.webmを読み込んでください。');
  }

  async function importImages(fileList) {
    const entries = await Promise.all(
      Array.from(fileList || []).map(async (file) => ({
        fileName: file.name,
        dataUrl: await fileToDataUrl(file)
      }))
    );
    mergeImages(entries);
  }

  async function importVideo(file) {
    if (!file) return;
    if (!project.recording_started_at) {
      setVideoStatus('先にevents.jsonを読み込んでください。');
      return;
    }
    if (orderedSteps.length === 0) {
      setVideoStatus('切り出すステップがありません。');
      return;
    }

    setVideoStatus('録画から画像を切り出しています。');
    try {
      const targetRecordingId = getPendingRecordingId(orderedSteps, recordingDataUrls);
      const targetSteps = orderedSteps.filter((step) => (step.recording_id || 'recording_1') === targetRecordingId);
      const recordingStartedAt = targetSteps[0]?.recording_started_at || project.recording_started_at;
      const dataUrl = await fileToDataUrl(file);
      setRecordingDataUrl(dataUrl);
      setRecordingDataUrls((current) => ({ ...current, [targetRecordingId]: dataUrl }));
      const entries = await extractStepImagesFromVideo(file, targetSteps, recordingStartedAt);
      mergeImages(entries);
      setVideoStatus(`${entries.length}件の画像を生成しました。`);
    } catch (error) {
      setVideoStatus(error.message || '録画から画像を生成できませんでした。');
    }
  }

  function mergeImages(entries) {
    setImages((current) => {
      const next = { ...current };
      for (const entry of entries) {
        for (const key of getImageKeys(entry.fileName)) {
          next[key] = entry.dataUrl;
        }
      }
      return next;
    });
  }

  function updateStep(index, patch) {
    setProject((current) => ({
      ...current,
      steps: current.steps.map((step, stepIndex) => (stepIndex === index ? { ...step, ...patch } : step))
    }));
  }

  function deleteStep(index) {
    setProject((current) => ({
      ...current,
      steps: current.steps.filter((_, stepIndex) => stepIndex !== index)
    }));
  }

  function moveStep(index, direction) {
    const nextIndex = index + direction;
    if (nextIndex < 0 || nextIndex >= project.steps.length) return;

    setProject((current) => {
      const steps = [...current.steps];
      const [step] = steps.splice(index, 1);
      steps.splice(nextIndex, 0, step);
      return { ...current, steps };
    });
  }

  async function exportWord() {
    const children = [
      new Paragraph({
        text: project.title || '作業手順書',
        heading: HeadingLevel.TITLE
      })
    ];

    for (const step of orderedSteps) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: `Step ${step.step_no}`, bold: true, size: 28 })]
        })
      );

      const markedImage = await createMarkedImage(step, getStepImage(images, step));
      if (markedImage) {
        const wordSize = containSize(markedImage.width, markedImage.height, 520, 320);
        children.push(
          new Paragraph({
            children: [
              new ImageRun({
                type: 'png',
                data: dataUrlToUint8Array(markedImage.dataUrl),
                transformation: { width: wordSize.w, height: wordSize.h }
              })
            ]
          })
        );
      }

      children.push(new Paragraph(step.description || step.element_text || ''));
    }

    const doc = new Document({ sections: [{ children }] });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${safeFileName(project.title || 'manual')}.docx`);
  }

  async function exportPowerPoint() {
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = '業務手順書自動作成ツール';
    pptx.subject = project.title;
    pptx.title = project.title;

    const chunkSize = Number(slidesPerPage);
    for (let i = 0; i < orderedSteps.length; i += chunkSize) {
      const slide = pptx.addSlide();
      slide.background = { color: 'FFFFFF' };
      slide.addText(project.title || '作業手順書', {
        x: 0.4,
        y: 0.2,
        w: 12.5,
        h: 0.35,
        fontSize: 15,
        bold: true,
        color: '111827'
      });

      const group = orderedSteps.slice(i, i + chunkSize);
      const layout = getSlideLayout(group.length);

      for (let j = 0; j < group.length; j += 1) {
        const step = group[j];
        const box = layout[j];
        const markedImage = await createMarkedImage(step, getStepImage(images, step));

        slide.addText(`Step ${step.step_no}`, {
          x: box.x,
          y: box.y,
          w: box.w,
          h: 0.25,
          fontSize: 11,
          bold: true,
          color: '2563EB'
        });

        if (markedImage) {
          addContainedImage(slide, markedImage, {
            x: box.x,
            y: box.y + 0.32,
            w: box.w,
            h: box.imageH
          });
        }

        slide.addText(step.description || step.element_text || '', {
          x: box.x,
          y: box.y + 0.38 + box.imageH,
          w: box.w,
          h: 0.45,
          fontSize: 9,
          color: '111827',
          fit: 'shrink'
        });
      }
    }

    await pptx.writeFile({ fileName: `${safeFileName(project.title || 'manual')}.pptx` });
  }

  async function exportHtml() {
    const htmlSteps = [];

    for (const step of orderedSteps) {
      const markedImage = await createMarkedImage(step, getStepImage(images, step));
      const imageHtml = markedImage
        ? `<img class="step-image" src="${markedImage.dataUrl}" alt="Step ${step.step_no}">`
        : '<div class="missing-image">画像がありません</div>';
      const description = step.description || createDefaultDescription(step);

      htmlSteps.push(`
        <section class="step">
          <div class="step-media">${imageHtml}</div>
          <div class="step-content">
            <p class="step-number">Step ${step.step_no}</p>
            <h2>${escapeHtml(getStepTitle(step))}</h2>
            <p class="description">${escapeHtml(description)}</p>
            <dl>
              <div><dt>画面</dt><dd>${escapeHtml(step.page_title || '-')}</dd></div>
              <div><dt>対象</dt><dd>${escapeHtml(getStepTitle(step))}</dd></div>
            </dl>
          </div>
        </section>
      `);
    }

    const htmlRecordingDataUrls =
      Object.keys(recordingDataUrls).length > 0 ? recordingDataUrls : recordingDataUrl ? { recording_1: recordingDataUrl } : {};
    const groupedHtmlSteps = Object.keys(htmlRecordingDataUrls).length > 0
      ? await createGroupedHtmlSteps(orderedSteps, images, htmlRecordingDataUrls)
      : htmlSteps;

    const html = `<!doctype html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(project.title || '作業手順書')}</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; color: #111827; font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: #f8fafc; }
    main { max-width: 1040px; margin: 0 auto; padding: 32px 20px 56px; }
    header { margin-bottom: 24px; }
    h1 { margin: 0; font-size: 28px; line-height: 1.35; }
    .subtitle { margin: 8px 0 0; color: #64748b; font-size: 14px; }
    .overview { display: grid; grid-template-columns: repeat(2, minmax(220px, 1fr)); gap: 12px; margin: 20px 0 24px; }
    .overview div { display: grid; grid-template-columns: 84px 1fr; gap: 10px; align-items: start; border: 1px solid #dbe3ef; border-radius: 8px; padding: 12px; background: #ffffff; }
    .overview dt { margin: 0; color: #2563eb; font-size: 13px; font-weight: 800; white-space: nowrap; }
    .overview dd { margin: 0; color: #111827; font-size: 14px; line-height: 1.7; }
    .step { display: grid; grid-template-columns: minmax(280px, 54%) 1fr; gap: 20px; align-items: start; margin-top: 18px; border: 1px solid #dbe3ef; border-radius: 8px; padding: 16px; background: #ffffff; break-inside: avoid; }
    .step-image { display: block; width: 100%; height: auto; border: 1px solid #e5e7eb; border-radius: 6px; }
    .step-video { display: block; width: 100%; height: auto; border: 1px solid #e5e7eb; border-radius: 6px; background: #000000; }
    .missing-image { display: grid; place-items: center; min-height: 220px; border: 1px dashed #cbd5e1; border-radius: 6px; color: #64748b; background: #f8fafc; }
    .step-number { margin: 0 0 8px; color: #2563eb; font-size: 13px; font-weight: 800; }
    h2 { margin: 0 0 12px; font-size: 20px; line-height: 1.45; }
    .description { margin: 0 0 16px; font-size: 16px; line-height: 1.8; }
    dl { display: grid; gap: 8px; margin: 0; color: #475569; font-size: 13px; }
    .step-content dl div { display: grid; grid-template-columns: 48px 1fr; gap: 8px; }
    dt { font-weight: 800; }
    dd { margin: 0; overflow-wrap: anywhere; }
    .operation-list { margin: 0; padding-left: 1.25rem; line-height: 1.8; }
    .video-step { border-color: #bfdbfe; background: #f8fbff; }
    @media (max-width: 760px) { .step { grid-template-columns: 1fr; } }
    @media print { body { background: #ffffff; } main { max-width: none; padding: 0; } .step { page-break-inside: avoid; } }
  </style>
</head>
<body>
  <main>
    <header>
      <h1>${escapeHtml(project.title || '作業手順書')}</h1>
      <p class="subtitle">${orderedSteps.length} ステップ</p>
    </header>
    <dl class="overview">
      <div><dt>目的</dt><dd>${escapeHtml(project.purpose || '-')}</dd></div>
      <div><dt>対象者</dt><dd>${escapeHtml(project.audience || '-')}</dd></div>
      <div><dt>前提条件</dt><dd>${escapeHtml(project.prerequisites || '-')}</dd></div>
      <div><dt>完了条件</dt><dd>${escapeHtml(project.completion || '-')}</dd></div>
    </dl>
    ${groupedHtmlSteps.join('\n')}
  </main>
  <script>
    const recordingDataUrls = JSON.parse('${escapeScriptString(JSON.stringify(htmlRecordingDataUrls))}');
    document.querySelectorAll('video[data-start][data-end]').forEach((video) => {
      const recordingDataUrl = recordingDataUrls[video.dataset.recordingId || 'recording_1'];
      if (recordingDataUrl) video.src = recordingDataUrl;
      const start = Number(video.dataset.start || 0);
      const end = Number(video.dataset.end || 0);
      video.addEventListener('loadedmetadata', () => { video.currentTime = start; });
      video.addEventListener('play', () => {
        if (video.currentTime < start || video.currentTime > end) video.currentTime = start;
      });
      video.addEventListener('timeupdate', () => {
        if (video.currentTime >= end) {
          video.pause();
          video.currentTime = start;
        }
      });
    });
  </script>
</body>
</html>`;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    saveAs(blob, `${safeFileName(project.title || 'manual')}.html`);
    exportProjectJson();
  }

  function exportProjectJson() {
    const editableProject = createEditableProject(project, orderedSteps);
    const blob = new Blob([JSON.stringify(editableProject, null, 2)], { type: 'application/json;charset=utf-8' });
    saveAs(blob, `${safeFileName(project.title || 'manual')}-project-${createFileStamp()}.json`);
  }

  return (
    <main className="app-shell">
      <section className="workspace">
        <header className="topbar">
          <div>
            <label className="eyebrow" htmlFor="projectTitle">
              タイトル
            </label>
            <input
              id="projectTitle"
              className="title-input"
              value={project.title}
              onChange={(event) => updateProjectTitle(event.target.value)}
            />
          </div>
          <div className="export-actions">
            <select value={slidesPerPage} onChange={(event) => setSlidesPerPage(event.target.value)}>
              <option value="1">1ステップ/スライド</option>
              <option value="2">2ステップ/スライド</option>
              <option value="4">4ステップ/スライド</option>
            </select>
            <button type="button" onClick={exportWord} disabled={orderedSteps.length === 0}>
              Word出力
            </button>
            <button type="button" onClick={exportHtml} disabled={orderedSteps.length === 0}>
              HTML出力
            </button>
            <button type="button" onClick={exportProjectJson} disabled={orderedSteps.length === 0}>
              編集JSON保存
            </button>
            <button type="button" onClick={exportPowerPoint} disabled={orderedSteps.length === 0}>
              PowerPoint出力
            </button>
          </div>
        </header>

        <section className="document-info">
          <label>
            目的
            <textarea
              value={project.purpose || ''}
              onChange={(event) => updateProjectInfo('purpose', event.target.value)}
            />
          </label>
          <label>
            対象者
            <textarea
              value={project.audience || ''}
              onChange={(event) => updateProjectInfo('audience', event.target.value)}
            />
          </label>
          <label>
            前提条件
            <textarea
              value={project.prerequisites || ''}
              onChange={(event) => updateProjectInfo('prerequisites', event.target.value)}
            />
          </label>
          <label>
            完了条件
            <textarea
              value={project.completion || ''}
              onChange={(event) => updateProjectInfo('completion', event.target.value)}
            />
          </label>
        </section>

        <section className="import-panel">
          <label>
            JSON読込
            <input type="file" accept="application/json,.json" onChange={(event) => importProject(event.target.files[0])} />
          </label>
          <label>
            JSON追記
            <input type="file" accept="application/json,.json" onChange={(event) => appendProject(event.target.files[0])} />
          </label>
          <label>
            WebM読込
            <input type="file" accept="video/webm,.webm" onChange={(event) => importVideo(event.target.files[0])} />
          </label>
          <label>
            PNG読込
            <input type="file" accept="image/png" multiple onChange={(event) => importImages(event.target.files)} />
          </label>
          {videoStatus ? <p className="import-status">{videoStatus}</p> : null}
        </section>

        <section className="steps">
          {orderedSteps.length === 0 ? (
            <div className="empty">events.json と recording.webm、または従来のJSONとPNGを読み込んでください。</div>
          ) : (
            orderedSteps.map((step, index) => (
              <StepEditor
                key={`${step.image}-${index}`}
                step={step}
                image={getStepImage(images, step)}
                index={index}
                total={orderedSteps.length}
                onUpdate={(patch) => updateStep(index, patch)}
                onDelete={() => deleteStep(index)}
                onMove={(direction) => moveStep(index, direction)}
              />
            ))
          )}
        </section>
      </section>
    </main>
  );
}

function StepEditor({ step, image, index, total, onUpdate, onDelete, onMove }) {
  return (
    <article className="step-card">
      <div className="step-preview">
        {image ? <MarkedPreview step={step} image={image} /> : <span>{step.image} を生成または読み込んでください</span>}
      </div>
      <div className="step-body">
        <div className="step-heading">
          <strong>Step {step.step_no}</strong>
          <div className="step-actions">
            <button type="button" onClick={() => onMove(-1)} disabled={index === 0}>
              上へ
            </button>
            <button type="button" onClick={() => onMove(1)} disabled={index === total - 1}>
              下へ
            </button>
            <button type="button" className="danger" onClick={onDelete}>
              削除
            </button>
          </div>
        </div>
        <dl className="meta">
          <div>
            <dt>ページ</dt>
            <dd>{step.page_title || '-'}</dd>
          </div>
          <div>
            <dt>対象</dt>
            <dd>{step.element_text || step.aria_label || '-'}</dd>
          </div>
          <div>
            <dt>URL</dt>
            <dd>{step.url || '-'}</dd>
          </div>
        </dl>
        <label className="description-label">
          説明文
          <textarea
            value={step.description || ''}
            placeholder={createDefaultDescription(step)}
            onChange={(event) => onUpdate({ description: event.target.value })}
          />
        </label>
      </div>
    </article>
  );
}

function MarkedPreview({ step, image }) {
  const markerStyle = {
    left: `${toPercent(step.x, step.viewport_width)}%`,
    top: `${toPercent(step.y, step.viewport_height)}%`
  };

  return (
    <div className="image-wrap">
      <img src={image} alt={`Step ${step.step_no}`} />
      <span className="marker" style={markerStyle}>
        {step.step_no}
      </span>
    </div>
  );
}

function normalizeProject(data, recordingId = 'recording_1') {
  const recordingFileMap = getRecordingFileMap(data);
  const steps = Array.isArray(data.steps)
    ? data.steps.map((step, index) => ({
        ...step,
        step_no: index + 1,
        recording_id: step.recording_id || recordingId,
        recording_started_at: step.recording_started_at || data.recording_started_at || null,
        recording_file: step.recording_file || recordingFileMap[step.recording_id || recordingId] || data.recording_file || '',
        image: step.image || `step_${String(index + 1).padStart(3, '0')}.png`,
        description: step.description || createDefaultDescription(step)
      }))
    : [];
  const inferredInfo = inferProjectInfo(data, steps);

  return {
    ...DEFAULT_PROJECT,
    title: data.title || DEFAULT_PROJECT.title,
    purpose: data.purpose || inferredInfo.purpose,
    audience: data.audience || inferredInfo.audience,
    prerequisites: data.prerequisites || inferredInfo.prerequisites,
    completion: data.completion || inferredInfo.completion,
    capture_mode: data.capture_mode || '',
    recording_started_at: data.recording_started_at || null,
    recording_file: data.recording_file || '',
    steps
  };
}

function createEditableProject(project, orderedSteps) {
  const recordings = getRecordingIds(orderedSteps).map((id) => {
    const step = orderedSteps.find((item) => (item.recording_id || 'recording_1') === id) || {};
    return {
      id,
      file: step.recording_file || (id === 'recording_1' ? project.recording_file : ''),
      recording_started_at: step.recording_started_at || project.recording_started_at || null
    };
  });

  return {
    schema: 'manual-creator-project',
    schema_version: 1,
    saved_at: new Date().toISOString(),
    title: project.title || DEFAULT_PROJECT.title,
    purpose: project.purpose || '',
    audience: project.audience || '',
    prerequisites: project.prerequisites || '',
    completion: project.completion || '',
    capture_mode: project.capture_mode || 'video',
    recording_started_at: project.recording_started_at || recordings[0]?.recording_started_at || null,
    recording_file: project.recording_file || recordings[0]?.file || '',
    recordings,
    steps: orderedSteps.map((step, index) => ({
      ...step,
      step_no: index + 1,
      recording_id: step.recording_id || 'recording_1',
      recording_file: step.recording_file || recordings.find((recording) => recording.id === (step.recording_id || 'recording_1'))?.file || '',
      recording_started_at: step.recording_started_at || project.recording_started_at || null,
      description: step.description || createDefaultDescription(step)
    }))
  };
}

function getRecordingFileMap(data) {
  const map = {};
  if (Array.isArray(data.recordings)) {
    for (const recording of data.recordings) {
      if (recording?.id && recording?.file) {
        map[recording.id] = recording.file;
      }
    }
  }
  if (data.recording_file) {
    map.recording_1 = map.recording_1 || data.recording_file;
  }
  return map;
}

function inferProjectInfo(data, steps) {
  const title = data.title || DEFAULT_PROJECT.title;
  const firstStep = steps[0] || {};
  const lastStep = steps[steps.length - 1] || {};
  const pageTitle = firstStep.page_title || lastStep.page_title || '';
  const url = firstStep.url || lastStep.url || '';
  const isKintone = /kintone\.com|cybozu\.com/.test(url);
  const finalAction = lastStep.description || createDefaultDescription(lastStep);

  return {
    purpose: `${title}の作業を、画面操作に沿って正確に実施するため。`,
    audience: inferAudience(pageTitle, url),
    prerequisites: isKintone
      ? 'kintoneにログイン済みであり、対象アプリを閲覧・操作できる権限があること。'
      : '対象システムにログイン済みであり、操作に必要な権限があること。',
    completion: finalAction ? `${finalAction}まで完了していること。` : '一連の操作が完了していること。'
  };
}

function inferAudience(pageTitle, url) {
  const text = `${pageTitle} ${url}`.toLowerCase();
  if (text.includes('logistics') || text.includes('receiving') || text.includes('shipping')) {
    return '物流担当者、業務担当者';
  }
  if (text.includes('sales') || text.includes('order')) {
    return '営業担当者、営業事務担当者';
  }
  if (text.includes('kintone')) {
    return 'kintone利用者、業務担当者';
  }
  return '業務担当者';
}

function getRecordingIds(steps) {
  return Array.from(new Set((steps || []).map((step) => step.recording_id || 'recording_1')));
}

function getPendingRecordingId(steps, recordingMap) {
  const ids = getRecordingIds(steps);
  return ids.find((id) => !recordingMap[id]) || ids[ids.length - 1] || 'recording_1';
}

async function extractStepImagesFromVideo(file, steps, recordingStartedAt) {
  const video = document.createElement('video');
  const objectUrl = URL.createObjectURL(file);
  video.src = objectUrl;
  video.muted = true;
  video.preload = 'metadata';

  try {
    await waitForEvent(video, 'loadedmetadata');
    const canvas = document.createElement('canvas');
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    const context = canvas.getContext('2d');
    const entries = [];

    for (const step of steps) {
      const time = getFrameTime(step, recordingStartedAt, video.duration);
      video.currentTime = time;
      await waitForEvent(video, 'seeked');
      context.drawImage(video, 0, 0, canvas.width, canvas.height);
      entries.push({
        fileName: step.image,
        dataUrl: canvas.toDataURL('image/png')
      });
    }

    return entries;
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
}

function getFrameTime(step, recordingStartedAt, duration) {
  const rawTime = (Number(step.timestamp_ms) - Number(recordingStartedAt)) / 1000;
  const afterClickOffset = 0.25;
  const maxTime = Number.isFinite(duration) ? Math.max(0, duration - 0.05) : rawTime + afterClickOffset;
  return Math.max(0, Math.min(rawTime + afterClickOffset, maxTime));
}

async function createGroupedHtmlSteps(steps, images, recordingDataUrls) {
  const groups = createHtmlStepGroups(steps);
  const html = [];

  for (const group of groups) {
    const recordingId = group.steps[0]?.recording_id || 'recording_1';
    if (group.type === 'video' && recordingDataUrls[recordingId]) {
      html.push(createVideoGroupHtml(group));
      continue;
    }

    const step = group.steps[0];
    const markedImage = await createMarkedImage(step, getStepImage(images, step));
    html.push(createImageStepHtml(step, markedImage));
  }

  return html;
}

function createHtmlStepGroups(steps) {
  const groups = [];
  let current = [];

  for (const step of steps) {
    const previous = current[current.length - 1];
    const canJoin =
      previous &&
      (step.recording_id || 'recording_1') === (previous.recording_id || 'recording_1') &&
      Math.abs(Number(step.timestamp_ms) - Number(previous.timestamp_ms)) <= 3000;

    if (!canJoin) {
      pushHtmlStepGroup(groups, current);
      current = [step];
      continue;
    }

    current.push(step);
  }

  pushHtmlStepGroup(groups, current);
  return groups;
}

function pushHtmlStepGroup(groups, steps) {
  if (steps.length === 0) return;
  groups.push({
    type: steps.length >= 2 ? 'video' : 'image',
    steps
  });
}

function createImageStepHtml(step, markedImage) {
  const imageHtml = markedImage
    ? `<img class="step-image" src="${markedImage.dataUrl}" alt="Step ${step.step_no}">`
    : '<div class="missing-image">画像がありません</div>';
  const description = step.description || createDefaultDescription(step);

  return `
    <section class="step">
      <div class="step-media">${imageHtml}</div>
      <div class="step-content">
        <p class="step-number">Step ${step.step_no}</p>
        <h2>${escapeHtml(getStepTitle(step))}</h2>
        <p class="description">${escapeHtml(description)}</p>
        <dl>
          <div><dt>画面</dt><dd>${escapeHtml(step.page_title || '-')}</dd></div>
          <div><dt>対象</dt><dd>${escapeHtml(getStepTitle(step))}</dd></div>
        </dl>
      </div>
    </section>
  `;
}

function createVideoGroupHtml(group) {
  const first = group.steps[0];
  const last = group.steps[group.steps.length - 1];
  const recordingStartedAt = first.recording_started_at;
  const recordingId = first.recording_id || 'recording_1';
  const start = Math.max(0, (Number(first.timestamp_ms) - Number(recordingStartedAt)) / 1000 - 0.35);
  const end = Math.max(start + 0.8, (Number(last.timestamp_ms) - Number(recordingStartedAt)) / 1000 + 0.75);
  const title = `${getStepTitle(first)}から${getStepTitle(last)}まで`;
  const operations = group.steps
    .map((step) => `<li>Step ${step.step_no}: ${escapeHtml(step.description || createDefaultDescription(step))}</li>`)
    .join('');

  return `
    <section class="step video-step">
      <div class="step-media">
        <video class="step-video" controls preload="metadata" data-recording-id="${escapeHtml(recordingId)}" data-start="${start.toFixed(2)}" data-end="${end.toFixed(2)}"></video>
      </div>
      <div class="step-content">
        <p class="step-number">Step ${first.step_no}-${last.step_no}</p>
        <h2>${escapeHtml(title)}</h2>
        <p class="description">以下の順に操作します。</p>
        <ol class="operation-list">${operations}</ol>
      </div>
    </section>
  `;
}

async function createMarkedImage(step, imageDataUrl) {
  if (!imageDataUrl) return null;

  const image = await loadImage(imageDataUrl);
  const maxWidth = 1600;
  const maxHeight = 1000;
  const scale = Math.min(1, maxWidth / image.width, maxHeight / image.height);
  const width = Math.round(image.width * scale);
  const height = Math.round(image.height * scale);
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext('2d');

  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);

  const markerX = (Number(step.x) / Number(step.viewport_width || image.width)) * width;
  const markerY = (Number(step.y) / Number(step.viewport_height || image.height)) * height;

  drawMarker(context, markerX, markerY, step.step_no);
  return {
    dataUrl: canvas.toDataURL('image/png'),
    width,
    height
  };
}

function drawMarker(context, x, y, stepNo) {
  context.save();
  context.strokeStyle = '#dc2626';
  context.fillStyle = '#dc2626';
  context.lineWidth = 5;

  context.beginPath();
  context.arc(x, y, 26, 0, Math.PI * 2);
  context.stroke();

  const labelX = Math.min(x + 58, context.canvas.width - 36);
  const labelY = Math.max(y - 48, 36);
  context.beginPath();
  context.moveTo(x + 22, y - 12);
  context.lineTo(labelX - 16, labelY + 12);
  context.stroke();

  context.beginPath();
  context.arc(labelX, labelY, 22, 0, Math.PI * 2);
  context.fill();
  context.fillStyle = '#ffffff';
  context.font = 'bold 22px system-ui, sans-serif';
  context.textAlign = 'center';
  context.textBaseline = 'middle';
  context.fillText(String(stepNo), labelX, labelY + 1);
  context.restore();
}

function getSlideLayout(count) {
  if (count === 1) return [{ x: 0.75, y: 0.8, w: 12.0, imageH: 5.6 }];
  if (count === 2) {
    return [
      { x: 0.55, y: 0.85, w: 6.0, imageH: 4.15 },
      { x: 6.85, y: 0.85, w: 6.0, imageH: 4.15 }
    ];
  }
  return [
    { x: 0.55, y: 0.75, w: 6.0, imageH: 2.35 },
    { x: 6.85, y: 0.75, w: 6.0, imageH: 2.35 },
    { x: 0.55, y: 3.95, w: 6.0, imageH: 2.35 },
    { x: 6.85, y: 3.95, w: 6.0, imageH: 2.35 }
  ];
}

function addContainedImage(slide, markedImage, box) {
  const size = containSize(markedImage.width, markedImage.height, box.w, box.h);
  slide.addImage({
    data: markedImage.dataUrl,
    x: box.x + (box.w - size.w) / 2,
    y: box.y + (box.h - size.h) / 2,
    w: size.w,
    h: size.h
  });
}

function containSize(sourceWidth, sourceHeight, maxWidth, maxHeight) {
  const sourceRatio = sourceWidth / sourceHeight;
  const boxRatio = maxWidth / maxHeight;

  if (sourceRatio > boxRatio) {
    return {
      w: maxWidth,
      h: maxWidth / sourceRatio
    };
  }

  return {
    w: maxHeight * sourceRatio,
    h: maxHeight
  };
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener('load', () => resolve(reader.result));
    reader.addEventListener('error', () => reject(reader.error));
    reader.readAsDataURL(file);
  });
}

function getStepImage(images, step) {
  for (const key of getImageKeys(step.image || '')) {
    if (images[key]) return images[key];
  }
  return null;
}

function getImageKeys(fileName) {
  const baseName = String(fileName).split(/[\\/]/).pop();
  const normalized = normalizeImageName(baseName);
  const keys = new Set([fileName, baseName, normalized]);

  if (normalized) {
    keys.add(normalized.toLowerCase());
  }

  const stepMatch = normalized.match(/step[_-]?0*(\d+)\.png$/i);
  if (stepMatch) {
    keys.add(`step_${String(Number(stepMatch[1])).padStart(3, '0')}.png`);
  }

  return Array.from(keys).filter(Boolean);
}

function normalizeImageName(fileName) {
  return String(fileName)
    .trim()
    .replace(/\s+\(\d+\)(?=\.png$)/i, '')
    .replace(/\s+-\s+コピー(?=\.png$)/i, '')
    .replace(/\s+copy(?=\.png$)/i, '');
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.addEventListener('load', () => resolve(image));
    image.addEventListener('error', () => reject(new Error('画像を読み込めませんでした。')));
    image.src = src;
  });
}

function waitForEvent(target, eventName) {
  return new Promise((resolve, reject) => {
    const cleanup = () => {
      target.removeEventListener(eventName, handleEvent);
      target.removeEventListener('error', handleError);
    };
    const handleEvent = () => {
      cleanup();
      resolve();
    };
    const handleError = () => {
      cleanup();
      reject(new Error('動画を読み込めませんでした。'));
    };
    target.addEventListener(eventName, handleEvent, { once: true });
    target.addEventListener('error', handleError, { once: true });
  });
}

function dataUrlToUint8Array(dataUrl) {
  const base64 = dataUrl.split(',')[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function toPercent(value, base) {
  const numericValue = Number(value);
  const numericBase = Number(base);
  if (!Number.isFinite(numericValue) || !Number.isFinite(numericBase) || numericBase <= 0) return 50;
  return Math.max(0, Math.min(100, (numericValue / numericBase) * 100));
}

function legacyCreateDefaultDescription(step) {
  const label = step.aria_label || step.element_text || step.placeholder || step.column_header || '対象';
  const suffix = step.role === 'button' || step.tag_name === 'BUTTON' ? 'をクリックします' : 'を選択します';
  return `${label}${suffix}`;
}

function safeFileName(value) {
  return value.replace(/[\\/:*?"<>|]/g, '_').trim() || 'manual';
}

function createFileStamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, '0');
  return [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
    '-',
    pad(now.getHours()),
    pad(now.getMinutes()),
    pad(now.getSeconds())
  ].join('');
}

function createDefaultDescription(step) {
  const label = getStepTitle(step);
  const normalized = normalizeLabel(label);
  const pageTitle = normalizeLabel(step.page_title || '');

  if (isButtonLike(step)) {
    if (normalized.includes('保存')) return '保存ボタンをクリックします。';
    if (normalized.includes('検索')) return '検索を実行します。';
    if (normalized.includes('追加')) return '追加ボタンをクリックします。';
    if (normalized.includes('編集')) return '編集ボタンをクリックします。';
    if (normalized.includes('削除')) return '削除ボタンをクリックします。';
    if (normalized.includes('キャンセル')) return 'キャンセルボタンをクリックします。';
    if (normalized.includes('レコード')) return 'レコード操作ボタンをクリックします。';
    return `${label}をクリックします。`;
  }

  if (step.role === 'link' || step.tag_name === 'A') {
    return `${label}を開きます。`;
  }

  if (step.role === 'input' || ['INPUT', 'TEXTAREA', 'SELECT'].includes(step.tag_name)) {
    if (step.placeholder) return `${step.placeholder}に入力します。`;
    if (step.column_header) return `${step.column_header}の項目を入力します。`;
    return `${label}を入力します。`;
  }

  if (step.column_header) {
    return `${step.column_header}の項目を選択します。`;
  }

  if (pageTitle.includes('レコードの一覧')) {
    return `${label}を選択します。`;
  }

  if (pageTitle.includes('レコードの詳細')) {
    return `${label}を確認します。`;
  }

  return `${label}をクリックします。`;
}

function getStepTitle(step) {
  return sanitizeLabel(step.aria_label || step.element_text || step.placeholder || step.column_header || '対象');
}

function sanitizeLabel(value) {
  const label = String(value || '').replace(/\s+/g, ' ').trim();
  if (!label) return '対象';
  if (/^\d{4,}$/.test(label)) return '対象レコード';
  if (/^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(label)) return 'メールアドレス項目';
  if (/^\d{4}[/-]\d{1,2}[/-]\d{1,2}/.test(label)) return '日付項目';
  if (label.length > 40) return `${label.slice(0, 40)}...`;
  return label;
}

function normalizeLabel(value) {
  return String(value || '').toLowerCase();
}

function isButtonLike(step) {
  return step.role === 'button' || step.tag_name === 'BUTTON' || step.tag_name === 'INPUT';
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeScriptString(value) {
  return String(value || '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/\r/g, '\\r')
    .replace(/\n/g, '\\n')
    .replace(/<\/script/gi, '<\\/script');
}

createRoot(document.getElementById('root')).render(<App />);
