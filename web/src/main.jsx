import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType
} from 'docx';
import { saveAs } from 'file-saver';
import pptxgen from 'pptxgenjs';
import './styles.css';

const DEFAULT_PROJECT = {
  title: '作業手順書',
  language: 'ja',
  created_at: '',
  author: '',
  purpose: '',
  audience: '',
  prerequisites: '',
  completion: '',
  capture_mode: '',
  recording_started_at: null,
  recording_file: '',
  steps: []
};

const WORD_CELL_MARGINS = {
  top: 140,
  bottom: 140,
  left: 160,
  right: 160
};

const WORD_META_CELL_MARGINS = {
  top: 80,
  bottom: 80,
  left: 100,
  right: 100
};

const WORD_TABLE_BORDERS = {
  top: { style: BorderStyle.SINGLE, size: 1, color: 'cbd5e1' },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'cbd5e1' },
  left: { style: BorderStyle.SINGLE, size: 1, color: 'cbd5e1' },
  right: { style: BorderStyle.SINGLE, size: 1, color: 'cbd5e1' },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'e2e8f0' },
  insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'e2e8f0' }
};

const WORD_INNER_TABLE_BORDERS = {
  top: { style: BorderStyle.SINGLE, size: 1, color: 'e2e8f0' },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'e2e8f0' },
  left: { style: BorderStyle.SINGLE, size: 1, color: 'e2e8f0' },
  right: { style: BorderStyle.SINGLE, size: 1, color: 'e2e8f0' },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'e2e8f0' },
  insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'e2e8f0' }
};

const TEXT = {
  ja: {
    title: 'タイトル',
    language: '言語',
    japanese: '日本語',
    english: '英語',
    createdAt: '作成日',
    author: '作成者',
    purpose: '目的',
    audience: '対象者',
    prerequisites: '前提条件',
    completion: '完了条件',
    steps: 'ステップ',
    stepList: '手順一覧',
    operationTarget: '操作対象',
    description: '説明文',
    screen: '画面',
    target: '対象',
    page: 'ページ',
    url: 'URL',
    wordExport: 'Word出力',
    jsonSave: '編集JSON保存',
    pptExport: 'PowerPoint出力',
    slideOne: '1ステップ/スライド',
    slideTwo: '2ステップ/スライド',
    slideFour: '4ステップ/スライド',
    jsonImport: 'JSON読込',
    jsonAppend: 'JSON追記',
    videoImport: '動画読込',
    pngImport: 'PNG読込',
    moveUp: '上へ',
    moveDown: '下へ',
    delete: '削除',
    imageMissing: '画像がありません',
    needImage: (image) => `${image} を生成または読み込んでください`,
    empty: 'events.json と recording.webm、または従来のJSONとPNGを読み込んでください。',
    continueVideo: (file) => `続けて${file || 'recording-*.webm'}を読み込んでください。`,
    appendVideo: '追加した記録のrecording-*.webmを読み込んでください。',
    importJsonFirst: '先にevents.jsonを読み込んでください。',
    noSteps: '切り出すステップがありません。',
    extracting: '録画から画像を切り出しています。',
    extracted: (count) => `${count}件の画像を生成しました。`,
    extractFailed: '録画から画像を生成できませんでした。',
    creator: '業務手順書自動作成ツール',
    defaultTitle: '作業手順書'
  },
  en: {
    title: 'Title',
    language: 'Language',
    japanese: 'Japanese',
    english: 'English',
    createdAt: 'Created Date',
    author: 'Author',
    purpose: 'Purpose',
    audience: 'Audience',
    prerequisites: 'Prerequisites',
    completion: 'Completion Criteria',
    steps: 'steps',
    stepList: 'Procedure List',
    operationTarget: 'Target',
    description: 'Description',
    screen: 'Screen',
    target: 'Target',
    page: 'Page',
    url: 'URL',
    wordExport: 'Export Word',
    jsonSave: 'Save JSON',
    pptExport: 'Export PowerPoint',
    slideOne: '1 step / slide',
    slideTwo: '2 steps / slide',
    slideFour: '4 steps / slide',
    jsonImport: 'Import JSON',
    jsonAppend: 'Append JSON',
    videoImport: 'Import Video',
    pngImport: 'Import PNG',
    moveUp: 'Up',
    moveDown: 'Down',
    delete: 'Delete',
    imageMissing: 'No image',
    needImage: (image) => `Generate or import ${image}`,
    empty: 'Import events.json and recording.webm, or legacy JSON and PNG files.',
    continueVideo: (file) => `Next, import ${file || 'recording-*.webm'}.`,
    appendVideo: 'Import the recording-*.webm file for the appended recording.',
    importJsonFirst: 'Import events.json first.',
    noSteps: 'There are no steps to extract.',
    extracting: 'Extracting images from the recording.',
    extracted: (count) => `Generated ${count} images.`,
    extractFailed: 'Could not generate images from the recording.',
    creator: 'Manual Creator',
    defaultTitle: 'Procedure Manual'
  }
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
  const language = project.language === 'en' ? 'en' : 'ja';
  const text = TEXT[language];

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
        ? text.continueVideo(normalized.recording_file)
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
    setVideoStatus(text.appendVideo);
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
      setVideoStatus(text.importJsonFirst);
      return;
    }
    if (orderedSteps.length === 0) {
      setVideoStatus(text.noSteps);
      return;
    }

    setVideoStatus(text.extracting);
    try {
      const targetRecordingId = getPendingRecordingId(orderedSteps, recordingDataUrls);
      const targetSteps = orderedSteps.filter((step) => (step.recording_id || 'recording_1') === targetRecordingId);
      const recordingStartedAt = targetSteps[0]?.recording_started_at || project.recording_started_at;
      const dataUrl = await fileToDataUrl(file);
      setRecordingDataUrl(dataUrl);
      setRecordingDataUrls((current) => ({ ...current, [targetRecordingId]: dataUrl }));
      setProject((current) => ({
        ...current,
        recording_file: targetRecordingId === 'recording_1' ? file.name : current.recording_file,
        steps: current.steps.map((step) =>
          (step.recording_id || 'recording_1') === targetRecordingId ? { ...step, recording_file: file.name } : step
        )
      }));
      const entries = await extractStepImagesFromVideo(file, targetSteps, recordingStartedAt);
      mergeImages(entries);
      setVideoStatus(text.extracted(entries.length));
    } catch (error) {
      setVideoStatus(error.message || text.extractFailed);
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
    const children = createWordCover(project, orderedSteps, text);

    for (const step of orderedSteps) {
      const markedImage = await createMarkedImage(step, getStepImage(images, step));
      children.push(createWordStepTable(step, markedImage, text));
      children.push(spacerParagraph(120));
    }

    const doc = new Document({
      creator: text.creator,
      title: project.title || text.defaultTitle,
      description: project.purpose || '',
      styles: {
        default: {
          document: {
            run: {
              font: 'Yu Gothic',
              size: 21,
              color: '111827'
            },
            paragraph: {
              spacing: { after: 80 }
            }
          }
        },
        paragraphStyles: [
          {
            id: 'Title',
            name: 'Title',
            basedOn: 'Normal',
            next: 'Normal',
            run: {
              size: 42,
              bold: true,
              color: '111827',
              font: 'Yu Gothic'
            },
            paragraph: {
              spacing: { after: 180 }
            }
          },
          {
            id: 'Heading1',
            name: 'Heading 1',
            basedOn: 'Normal',
            next: 'Normal',
            run: {
              size: 28,
              bold: true,
              color: '1d4ed8',
              font: 'Yu Gothic'
            },
            paragraph: {
              spacing: { before: 220, after: 120 }
            }
          }
        ]
      },
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 720,
                right: 720,
                bottom: 720,
                left: 720
              }
            }
          },
          children
        }
      ]
    });
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${safeFileName(project.title || 'manual')}.docx`);
  }

  async function exportPowerPoint() {
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = project.author || text.creator;
    pptx.subject = project.title;
    pptx.title = project.title;
    pptx.company = 'Yahata';
    pptx.lang = language === 'en' ? 'en-US' : 'ja-JP';

    addPptCoverSlide(pptx, project, orderedSteps, text);
    addPptOverviewSlide(pptx, project, orderedSteps, text);

    const chunkSize = Number(slidesPerPage);
    for (let i = 0; i < orderedSteps.length; i += chunkSize) {
      const slide = pptx.addSlide();
      const group = orderedSteps.slice(i, i + chunkSize);
      addPptSlideHeader(slide, project.title || text.defaultTitle, `${i + 1}-${i + group.length} / ${orderedSteps.length}`);

      if (group.length === 1) {
        const step = group[0];
        const markedImage = await createMarkedImage(step, getStepImage(images, step));
        addPptSingleStep(slide, step, markedImage, text);
        addPptFooter(slide, project, text);
        continue;
      }

      const layout = getPptCardLayout(group.length);

      for (let j = 0; j < group.length; j += 1) {
        const step = group[j];
        const box = layout[j];
        const markedImage = await createMarkedImage(step, getStepImage(images, step));
        addPptStepCard(slide, step, markedImage, box, text);
      }

      addPptFooter(slide, project, text);
    }

    await pptx.writeFile({ fileName: `${safeFileName(project.title || 'manual')}.pptx` });
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
              {text.title}
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
              <option value="1">{text.slideOne}</option>
              <option value="2">{text.slideTwo}</option>
              <option value="4">{text.slideFour}</option>
            </select>
            <button type="button" onClick={exportWord} disabled={orderedSteps.length === 0}>
              {text.wordExport}
            </button>
            <button type="button" onClick={exportProjectJson} disabled={orderedSteps.length === 0}>
              {text.jsonSave}
            </button>
            <button type="button" onClick={exportPowerPoint} disabled={orderedSteps.length === 0}>
              {text.pptExport}
            </button>
          </div>
        </header>

        <section className="document-info">
          <label>
            {text.language}
            <select value={language} onChange={(event) => updateProjectInfo('language', event.target.value)}>
              <option value="ja">{text.japanese}</option>
              <option value="en">{text.english}</option>
            </select>
          </label>
          <label>
            {text.createdAt}
            <input
              type="date"
              value={project.created_at || ''}
              onChange={(event) => updateProjectInfo('created_at', event.target.value)}
            />
          </label>
          <label>
            {text.author}
            <input value={project.author || ''} onChange={(event) => updateProjectInfo('author', event.target.value)} />
          </label>
          <label>
            {text.purpose}
            <textarea
              value={project.purpose || ''}
              onChange={(event) => updateProjectInfo('purpose', event.target.value)}
            />
          </label>
          <label>
            {text.audience}
            <textarea
              value={project.audience || ''}
              onChange={(event) => updateProjectInfo('audience', event.target.value)}
            />
          </label>
          <label>
            {text.prerequisites}
            <textarea
              value={project.prerequisites || ''}
              onChange={(event) => updateProjectInfo('prerequisites', event.target.value)}
            />
          </label>
          <label>
            {text.completion}
            <textarea
              value={project.completion || ''}
              onChange={(event) => updateProjectInfo('completion', event.target.value)}
            />
          </label>
        </section>

        <section className="import-panel">
          <label>
            {text.jsonImport}
            <input type="file" accept="application/json,.json" onChange={(event) => importProject(event.target.files[0])} />
          </label>
          <label>
            {text.jsonAppend}
            <input type="file" accept="application/json,.json" onChange={(event) => appendProject(event.target.files[0])} />
          </label>
          <label>
            {text.videoImport}
            <input type="file" accept="video/webm,.webm" onChange={(event) => importVideo(event.target.files[0])} />
          </label>
          <label>
            {text.pngImport}
            <input type="file" accept="image/png" multiple onChange={(event) => importImages(event.target.files)} />
          </label>
          {videoStatus ? <p className="import-status">{videoStatus}</p> : null}
        </section>

        <section className="steps">
          {orderedSteps.length === 0 ? (
            <div className="empty">{text.empty}</div>
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
                text={text}
              />
            ))
          )}
        </section>
      </section>
    </main>
  );
}

function StepEditor({ step, image, index, total, onUpdate, onDelete, onMove, text }) {
  return (
    <article className="step-card">
      <div className="step-preview">
        {image ? <MarkedPreview step={step} image={image} /> : <span>{text.needImage(step.image)}</span>}
      </div>
      <div className="step-body">
        <div className="step-heading">
          <strong>Step {step.step_no}</strong>
          <div className="step-actions">
            <button type="button" onClick={() => onMove(-1)} disabled={index === 0}>
              {text.moveUp}
            </button>
            <button type="button" onClick={() => onMove(1)} disabled={index === total - 1}>
              {text.moveDown}
            </button>
            <button type="button" className="danger" onClick={onDelete}>
              {text.delete}
            </button>
          </div>
        </div>
        <dl className="meta">
          <div>
            <dt>{text.page}</dt>
            <dd>{step.page_title || '-'}</dd>
          </div>
          <div>
            <dt>{text.target}</dt>
            <dd>{step.element_text || step.aria_label || '-'}</dd>
          </div>
          <div>
            <dt>URL</dt>
            <dd>{step.url || '-'}</dd>
          </div>
        </dl>
        <label className="description-label">
          {text.description}
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

function createWordCover(project, orderedSteps, text) {
  return [
    new Paragraph({
      text: project.title || text.defaultTitle,
      heading: HeadingLevel.TITLE
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `${orderedSteps.length} ${text.steps}`,
          color: '64748b',
          size: 22
        })
      ],
      spacing: { after: 180 }
    }),
    createWordOverviewTable(project, text),
    spacerParagraph(120),
    new Paragraph({
      text: text.stepList,
      heading: HeadingLevel.HEADING_1
    }),
    createWordStepListTable(orderedSteps, text),
    new Paragraph({
      children: [new PageBreak()]
    })
  ];
}

function createWordOverviewTable(project, text) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: [1600, 7600],
    borders: WORD_TABLE_BORDERS,
    rows: [
      createWordInfoRow(text.createdAt, project.created_at || createDateStamp()),
      createWordInfoRow(text.author, project.author || '-'),
      createWordInfoRow(text.purpose, project.purpose || '-'),
      createWordInfoRow(text.audience, project.audience || '-'),
      createWordInfoRow(text.prerequisites, project.prerequisites || '-'),
      createWordInfoRow(text.completion, project.completion || '-')
    ]
  });
}

function createWordInfoRow(label, value) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 18, type: WidthType.PERCENTAGE },
        shading: { fill: 'eff6ff' },
        margins: WORD_CELL_MARGINS,
        verticalAlign: VerticalAlign.CENTER,
        children: [wordText(label, { bold: true, color: '1d4ed8' })]
      }),
      new TableCell({
        width: { size: 82, type: WidthType.PERCENTAGE },
        margins: WORD_CELL_MARGINS,
        children: [wordText(value)]
      })
    ]
  });
}

function createWordStepListTable(steps, text) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: [900, 3000, 5300],
    borders: WORD_TABLE_BORDERS,
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          wordHeaderCell('No.'),
          wordHeaderCell(text.operationTarget),
          wordHeaderCell(text.description)
        ]
      }),
      ...steps.map((step) =>
        new TableRow({
          children: [
            wordCell(`Step ${step.step_no}`, { width: 10, bold: true, color: '1d4ed8' }),
            wordCell(getStepTitle(step), { width: 32 }),
            wordCell(step.description || createDefaultDescription(step), { width: 58 })
          ]
        })
      )
    ]
  });
}

function createWordStepTable(step, markedImage, text) {
  const wordImageSize = markedImage ? containWordImageSize(markedImage.width, markedImage.height, 360, 230) : null;
  const imageParagraph = markedImage
    ? new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new ImageRun({
            type: 'png',
            data: dataUrlToUint8Array(markedImage.dataUrl),
            transformation: wordImageSize
          })
        ]
      })
    : wordText(text.imageMissing, { color: '64748b', alignment: AlignmentType.CENTER });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    columnWidths: [4300, 4900],
    borders: WORD_TABLE_BORDERS,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 18, type: WidthType.PERCENTAGE },
            shading: { fill: 'eff6ff' },
            margins: WORD_CELL_MARGINS,
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: `Step ${step.step_no}`, bold: true, color: '1d4ed8', size: 24 })
                ]
              })
            ]
          }),
          new TableCell({
            width: { size: 82, type: WidthType.PERCENTAGE },
            shading: { fill: 'eff6ff' },
            margins: WORD_CELL_MARGINS,
            children: [
              new Paragraph({
                children: [new TextRun({ text: getStepTitle(step), bold: true, color: '111827', size: 24 })]
              })
            ]
          })
        ]
      }),
      new TableRow({
        children: [
          new TableCell({
            width: { size: 47, type: WidthType.PERCENTAGE },
            margins: WORD_CELL_MARGINS,
            verticalAlign: VerticalAlign.CENTER,
            children: [imageParagraph]
          }),
          new TableCell({
            width: { size: 53, type: WidthType.PERCENTAGE },
            margins: WORD_CELL_MARGINS,
            children: [
              wordLabel(text.description),
              wordText(step.description || createDefaultDescription(step), { size: 22 }),
              spacerParagraph(80),
              createWordMetaTable(step, text)
            ]
          })
        ]
      })
    ]
  });
}

function createWordMetaTable(step, text) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: WORD_INNER_TABLE_BORDERS,
    rows: [
      createWordMetaRow(text.screen, step.page_title || '-'),
      createWordMetaRow(text.target, step.element_text || step.aria_label || getStepTitle(step)),
      createWordMetaRow(text.url, step.url || '-')
    ]
  });
}

function createWordMetaRow(label, value) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 20, type: WidthType.PERCENTAGE },
        shading: { fill: 'f8fafc' },
        margins: WORD_META_CELL_MARGINS,
        children: [wordText(label, { bold: true, color: '475569', size: 18 })]
      }),
      new TableCell({
        width: { size: 80, type: WidthType.PERCENTAGE },
        margins: WORD_META_CELL_MARGINS,
        children: [wordText(value, { size: 18, color: '475569' })]
      })
    ]
  });
}

function wordHeaderCell(text) {
  return new TableCell({
    shading: { fill: '1d4ed8' },
    margins: WORD_CELL_MARGINS,
    children: [wordText(text, { bold: true, color: 'ffffff' })]
  });
}

function wordCell(text, options = {}) {
  const cellOptions = {
    margins: WORD_CELL_MARGINS,
    children: [wordText(text, options)]
  };
  if (options.width) {
    cellOptions.width = { size: options.width, type: WidthType.PERCENTAGE };
  }
  return new TableCell(cellOptions);
}

function wordLabel(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, color: '1d4ed8', size: 18 })],
    spacing: { after: 50 }
  });
}

function wordText(text, options = {}) {
  const paragraphOptions = {
    children: [
      new TextRun({
        text: String(text || ''),
        bold: Boolean(options.bold),
        color: options.color || '111827',
        size: options.size || 20
      })
    ],
    spacing: { after: options.after ?? 60 }
  };
  if (options.alignment) {
    paragraphOptions.alignment = options.alignment;
  }
  return new Paragraph(paragraphOptions);
}

function spacerParagraph(after = 120) {
  return new Paragraph({
    children: [new TextRun({ text: '' })],
    spacing: { after }
  });
}

function containWordImageSize(sourceWidth, sourceHeight, maxWidth, maxHeight) {
  const size = containSize(sourceWidth, sourceHeight, maxWidth, maxHeight);
  return {
    width: Math.max(1, Math.round(size.w)),
    height: Math.max(1, Math.round(size.h))
  };
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
    language: data.language === 'en' ? 'en' : 'ja',
    created_at: data.created_at || createDateStamp(),
    author: data.author || '',
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
    language: project.language === 'en' ? 'en' : 'ja',
    created_at: project.created_at || createDateStamp(),
    author: project.author || '',
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

async function createExternalAssetHtmlSteps(steps, images, assetsFolder, recordingFileNames) {
  const groups = createHtmlStepGroups(steps);
  const html = [];

  for (const group of groups) {
    const recordingId = group.steps[0]?.recording_id || 'recording_1';
    if (group.type === 'video' && recordingFileNames[recordingId]) {
      const posterAsset = await addMarkedImageAsset(
        assetsFolder,
        group.steps[0],
        images,
        `group_${String(group.steps[0].step_no).padStart(3, '0')}_${String(group.steps[group.steps.length - 1].step_no).padStart(3, '0')}.png`
      );
      html.push(createExternalVideoGroupHtml(group, recordingFileNames[recordingId], posterAsset));
      continue;
    }

    const step = group.steps[0];
    const imageAsset = await addMarkedImageAsset(assetsFolder, step, images, `step_${String(step.step_no).padStart(3, '0')}.png`);
    html.push(createExternalImageStepHtml(step, imageAsset));
  }

  return html;
}

async function addMarkedImageAsset(assetsFolder, step, images, fileName) {
  const markedImage = await createMarkedImage(step, getStepImage(images, step));
  if (!markedImage) return null;

  const safeName = safeAssetFileName(fileName, `step_${String(step.step_no).padStart(3, '0')}.png`);
  assetsFolder.file(safeName, dataUrlToUint8Array(markedImage.dataUrl));
  return `assets/${safeName}`;
}

async function createSharePointPastePackage(project, steps, images, zip) {
  const procedureSteps = [];

  for (const step of steps) {
    const imageName = `step_${String(step.step_no).padStart(2, '0')}.png`;
    const markedImage = await createMarkedImage(step, getStepImage(images, step));

    if (markedImage) {
      zip.file(imageName, dataUrlToUint8Array(markedImage.dataUrl));
    }

    procedureSteps.push({
      step_no: step.step_no,
      title: getStepTitle(step),
      description: step.description || createDefaultDescription(step),
      image: imageName,
      video: null,
      video_required: false,
      video_status: 'not_used',
      clip_source: null,
      clip_start_sec: null,
      clip_end_sec: null,
      source_element: {
        tag: step.tag_name || '',
        text: step.element_text || '',
        aria_label: step.aria_label || '',
        selector: step.selector || ''
      }
    });
  }

  const procedure = {
    title: project.title || DEFAULT_PROJECT.title,
    language: project.language === 'en' ? 'en' : 'ja',
    created_at: project.created_at || createDateStamp(),
    author: project.author || '',
    category: '事務手順書',
    purpose: project.purpose || '',
    audience: project.audience || '',
    prerequisites: project.prerequisites || '',
    completion: project.completion || '',
    notes: [
      'SharePointページへ貼り付けるための半自動出力です。',
      '動画は出力しません。録画ファイルはクリック時点のスクリーンショット抽出素材として使用します。'
    ],
    steps: procedureSteps
  };

  return {
    procedure,
    markdown: createSharePointPasteMarkdown(procedure)
  };
}

function createSharePointPasteMarkdown(procedure) {
  const lines = [
    `# ${procedure.title}`,
    '',
    '## 概要',
    procedure.purpose || 'このページは、事務作業の手順を説明します。',
    '',
    `対象者：${procedure.audience || '-'}`,
    '',
    `前提条件：${procedure.prerequisites || '-'}`,
    '',
    `完了条件：${procedure.completion || '-'}`,
    ''
  ];

  for (const step of procedure.steps) {
    lines.push(`## 手順${step.step_no}：${step.title}`);
    lines.push(step.description || '');
    lines.push('');
    lines.push(`画像：${step.image}`);
    lines.push('動画：なし');
    lines.push('');
  }

  lines.push('## SharePoint配置メモ');
  lines.push('1. このMarkdown本文をSharePointページへ貼り付けます。');
  lines.push('2. 各手順の画像ファイルを、該当箇所へ手動で配置します。');
  lines.push('');

  return `${lines.join('\n')}\n`;
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

function createExternalImageStepHtml(step, imagePath) {
  const imageHtml = imagePath
    ? `<img class="step-image" src="${escapeHtml(imagePath)}" alt="Step ${step.step_no}">`
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

function createExternalVideoGroupHtml(group, videoPath, posterPath) {
  const first = group.steps[0];
  const last = group.steps[group.steps.length - 1];
  const recordingStartedAt = first.recording_started_at;
  const start = Math.max(0, (Number(first.timestamp_ms) - Number(recordingStartedAt)) / 1000 - 0.35);
  const end = Math.max(start + 0.8, (Number(last.timestamp_ms) - Number(recordingStartedAt)) / 1000 + 0.75);
  const title = `${getStepTitle(first)}から${getStepTitle(last)}まで`;
  const operations = group.steps
    .map((step) => `<li>Step ${step.step_no}: ${escapeHtml(step.description || createDefaultDescription(step))}</li>`)
    .join('');
  const poster = posterPath ? ` poster="${escapeHtml(posterPath)}"` : '';

  return `
    <section class="step video-step">
      <div class="step-media">
        <video class="step-video" controls preload="metadata" src="${escapeHtml(videoPath)}#t=${start.toFixed(2)},${end.toFixed(2)}"${poster}></video>
        ${
          posterPath
            ? `<p class="asset-note"><a href="${escapeHtml(videoPath)}">動画ファイルを開く</a> / <a href="${escapeHtml(posterPath)}">静止画を開く</a></p>`
            : `<p class="asset-note"><a href="${escapeHtml(videoPath)}">動画ファイルを開く</a></p>`
        }
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

function createManualHtml(project, orderedSteps, groupedHtmlSteps, script = '') {
  return `<!doctype html>
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
    .asset-note { margin: 8px 0 0; color: #64748b; font-size: 12px; line-height: 1.6; }
    .asset-note a { color: #2563eb; }
    @media (max-width: 760px) { .step, .overview { grid-template-columns: 1fr; } }
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
  ${script}
</body>
</html>`;
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

function addPptCoverSlide(pptx, project, steps, text) {
  const slide = pptx.addSlide();
  slide.background = { color: 'F8FAFC' };
  slide.addShape('rect', {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.18,
    fill: { color: '2563EB' },
    line: { color: '2563EB' }
  });
  slide.addText(project.title || text.defaultTitle, {
    x: 0.72,
    y: 1.25,
    w: 9.6,
    h: 0.65,
    fontFace: 'Yu Gothic',
    fontSize: 30,
    bold: true,
    color: '111827',
    margin: 0
  });
  slide.addText(`${steps.length} ${text.steps}`, {
    x: 0.76,
    y: 2.08,
    w: 3.2,
    h: 0.28,
    fontFace: 'Yu Gothic',
    fontSize: 12,
    color: '64748B',
    margin: 0
  });
  addPptInfoBand(slide, [
    [text.createdAt, project.created_at || createDateStamp()],
    [text.author, project.author || '-'],
    [text.purpose, project.purpose || '-']
  ]);
}

function addPptOverviewSlide(pptx, project, steps, text) {
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addPptSlideHeader(slide, text.stepList, `${steps.length} ${text.steps}`);

  const rows = [
    [text.purpose, project.purpose || '-'],
    [text.audience, project.audience || '-'],
    [text.prerequisites, project.prerequisites || '-'],
    [text.completion, project.completion || '-']
  ];
  rows.forEach(([label, value], index) => {
    const y = 0.9 + index * 0.68;
    slide.addShape('roundRect', {
      x: 0.65,
      y,
      w: 12.0,
      h: 0.52,
      rectRadius: 0.04,
      fill: { color: index % 2 === 0 ? 'F8FAFC' : 'FFFFFF' },
      line: { color: 'DBE3EF', width: 0.8 }
    });
    slide.addText(label, {
      x: 0.9,
      y: y + 0.14,
      w: 1.4,
      h: 0.18,
      fontFace: 'Yu Gothic',
      fontSize: 9,
      bold: true,
      color: '2563EB',
      margin: 0
    });
    slide.addText(value, {
      x: 2.3,
      y: y + 0.11,
      w: 9.9,
      h: 0.25,
      fontFace: 'Yu Gothic',
      fontSize: 10,
      color: '111827',
      fit: 'shrink',
      margin: 0
    });
  });

  const tableTop = 3.85;
  slide.addText(text.stepList, {
    x: 0.7,
    y: tableTop,
    w: 3.0,
    h: 0.24,
    fontFace: 'Yu Gothic',
    fontSize: 13,
    bold: true,
    color: '111827',
    margin: 0
  });

  steps.slice(0, 10).forEach((step, index) => {
    const y = tableTop + 0.42 + index * 0.27;
    slide.addText(`Step ${step.step_no}`, {
      x: 0.75,
      y,
      w: 0.9,
      h: 0.18,
      fontFace: 'Yu Gothic',
      fontSize: 7.5,
      bold: true,
      color: '2563EB',
      margin: 0
    });
    slide.addText(getStepTitle(step), {
      x: 1.72,
      y,
      w: 2.5,
      h: 0.18,
      fontFace: 'Yu Gothic',
      fontSize: 7.5,
      color: '111827',
      fit: 'shrink',
      margin: 0
    });
    slide.addText(step.description || createDefaultDescription(step), {
      x: 4.25,
      y,
      w: 8.0,
      h: 0.18,
      fontFace: 'Yu Gothic',
      fontSize: 7.5,
      color: '475569',
      fit: 'shrink',
      margin: 0
    });
  });
  addPptFooter(slide, project, text);
}

function addPptSingleStep(slide, step, markedImage, text) {
  slide.addShape('roundRect', {
    x: 0.55,
    y: 0.88,
    w: 7.35,
    h: 5.65,
    rectRadius: 0.04,
    fill: { color: 'F8FAFC' },
    line: { color: 'DBE3EF', width: 0.8 }
  });
  if (markedImage) {
    addContainedImage(slide, markedImage, { x: 0.75, y: 1.08, w: 6.95, h: 5.25 });
  } else {
    slide.addText(text.imageMissing, {
      x: 0.75,
      y: 3.4,
      w: 6.95,
      h: 0.3,
      align: 'center',
      fontFace: 'Yu Gothic',
      fontSize: 12,
      color: '64748B'
    });
  }

  slide.addShape('roundRect', {
    x: 8.15,
    y: 0.88,
    w: 4.65,
    h: 5.65,
    rectRadius: 0.04,
    fill: { color: 'FFFFFF' },
    line: { color: 'DBE3EF', width: 0.8 }
  });
  slide.addText(`Step ${step.step_no}`, {
    x: 8.45,
    y: 1.2,
    w: 1.2,
    h: 0.24,
    fontFace: 'Yu Gothic',
    fontSize: 12,
    bold: true,
    color: '2563EB',
    margin: 0
  });
  slide.addText(getStepTitle(step), {
    x: 8.45,
    y: 1.58,
    w: 3.9,
    h: 0.45,
    fontFace: 'Yu Gothic',
    fontSize: 17,
    bold: true,
    color: '111827',
    fit: 'shrink',
    margin: 0
  });
  slide.addText(step.description || createDefaultDescription(step), {
    x: 8.45,
    y: 2.28,
    w: 3.95,
    h: 1.05,
    fontFace: 'Yu Gothic',
    fontSize: 13,
    color: '111827',
    fit: 'shrink',
    margin: 0.04,
    breakLine: false
  });
  addPptMeta(slide, step, text, 8.45, 3.65, 3.95);
}

function addPptStepCard(slide, step, markedImage, box, text) {
  slide.addShape('roundRect', {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    rectRadius: 0.04,
    fill: { color: 'FFFFFF' },
    line: { color: 'DBE3EF', width: 0.8 }
  });
  slide.addText(`Step ${step.step_no}`, {
    x: box.x + 0.18,
    y: box.y + 0.14,
    w: 0.9,
    h: 0.18,
    fontFace: 'Yu Gothic',
    fontSize: 8,
    bold: true,
    color: '2563EB',
    margin: 0
  });
  slide.addText(getStepTitle(step), {
    x: box.x + 1.0,
    y: box.y + 0.13,
    w: box.w - 1.25,
    h: 0.2,
    fontFace: 'Yu Gothic',
    fontSize: 9,
    bold: true,
    color: '111827',
    fit: 'shrink',
    margin: 0
  });
  if (markedImage) {
    addContainedImage(slide, markedImage, {
      x: box.x + 0.18,
      y: box.y + 0.48,
      w: box.imageW,
      h: box.imageH
    });
  }
  slide.addText(step.description || createDefaultDescription(step), {
    x: box.x + box.imageW + 0.38,
    y: box.y + 0.5,
    w: box.w - box.imageW - 0.58,
    h: box.imageH,
    fontFace: 'Yu Gothic',
    fontSize: 8.5,
    color: '111827',
    fit: 'shrink',
    margin: 0.03
  });
}

function addPptSlideHeader(slide, title, meta) {
  slide.addShape('rect', {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.12,
    fill: { color: '2563EB' },
    line: { color: '2563EB' }
  });
  slide.addText(title, {
    x: 0.55,
    y: 0.28,
    w: 9.4,
    h: 0.26,
    fontFace: 'Yu Gothic',
    fontSize: 15,
    bold: true,
    color: '111827',
    fit: 'shrink',
    margin: 0
  });
  slide.addText(meta, {
    x: 10.7,
    y: 0.31,
    w: 2.1,
    h: 0.18,
    align: 'right',
    fontFace: 'Yu Gothic',
    fontSize: 8,
    color: '64748B',
    margin: 0
  });
}

function addPptFooter(slide, project, text) {
  slide.addText(`${text.createdAt}: ${project.created_at || createDateStamp()}   ${text.author}: ${project.author || '-'}`, {
    x: 0.55,
    y: 7.08,
    w: 7.2,
    h: 0.16,
    fontFace: 'Yu Gothic',
    fontSize: 7,
    color: '94A3B8',
    margin: 0
  });
}

function addPptInfoBand(slide, rows) {
  rows.forEach(([label, value], index) => {
    const y = 3.05 + index * 0.5;
    slide.addText(label, {
      x: 0.78,
      y,
      w: 1.2,
      h: 0.18,
      fontFace: 'Yu Gothic',
      fontSize: 8.5,
      bold: true,
      color: '2563EB',
      margin: 0
    });
    slide.addText(value, {
      x: 2.05,
      y,
      w: 8.8,
      h: 0.2,
      fontFace: 'Yu Gothic',
      fontSize: 9.5,
      color: '111827',
      fit: 'shrink',
      margin: 0
    });
  });
}

function addPptMeta(slide, step, text, x, y, w) {
  const rows = [
    [text.screen, step.page_title || '-'],
    [text.target, step.element_text || step.aria_label || getStepTitle(step)],
    [text.url, step.url || '-']
  ];
  rows.forEach(([label, value], index) => {
    const rowY = y + index * 0.36;
    slide.addText(label, {
      x,
      y: rowY,
      w: 0.75,
      h: 0.18,
      fontFace: 'Yu Gothic',
      fontSize: 7,
      bold: true,
      color: '64748B',
      margin: 0
    });
    slide.addText(value, {
      x: x + 0.8,
      y: rowY,
      w: w - 0.85,
      h: 0.18,
      fontFace: 'Yu Gothic',
      fontSize: 7,
      color: '475569',
      fit: 'shrink',
      margin: 0
    });
  });
}

function getPptCardLayout(count) {
  if (count === 2) {
    return [
      { x: 0.6, y: 0.95, w: 6.05, h: 5.6, imageW: 3.35, imageH: 4.55 },
      { x: 6.85, y: 0.95, w: 6.05, h: 5.6, imageW: 3.35, imageH: 4.55 }
    ];
  }
  return [
    { x: 0.55, y: 0.9, w: 6.1, h: 2.85, imageW: 2.95, imageH: 2.05 },
    { x: 6.85, y: 0.9, w: 6.1, h: 2.85, imageW: 2.95, imageH: 2.05 },
    { x: 0.55, y: 3.9, w: 6.1, h: 2.85, imageW: 2.95, imageH: 2.05 },
    { x: 6.85, y: 3.9, w: 6.1, h: 2.85, imageW: 2.95, imageH: 2.05 }
  ];
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

function safeAssetFileName(value, fallback) {
  const name = String(value || fallback || 'asset')
    .split(/[\\/]/)
    .pop()
    .replace(/[\\/:*?"<>|#%&{}$!'@+=`]/g, '_')
    .trim();
  return name || fallback || 'asset';
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

function createDateStamp() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, '0');
  return [now.getFullYear(), pad(now.getMonth() + 1), pad(now.getDate())].join('-');
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
