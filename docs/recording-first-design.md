# Scribe代替ツール 設計メモ

## 概要

Chrome拡張とWebアプリで構成する、社内向け操作手順書自動生成ツール。

Scribeのような共有・クラウド管理を主目的にせず、日本企業の業務手順書として「さっと見てわかる」HTML / Word / PowerPointを出力する。

現行MVPの `captureVisibleTab()` をクリックごとに呼ぶ方式は、モーダル消失・画面遷移・描画完了待ちのズレが避けづらい。次フェーズでは録画を一次ソースにし、クリック時刻から最適なフレームを切り出す方式へ移行する。

## アーキテクチャ

```text
Chrome拡張
├── content.js       mousedown検知・DOM情報取得
├── background.js    Service Worker。記録状態、イベント蓄積、Offscreen制御
└── offscreen.js     MediaRecorderによる録画処理

Webアプリ
├── JSON + PNG / WebM受け取り
├── GIF生成 / スクショ判定
├── HTML / Word / PowerPoint出力
└── AI説明文生成（オプション）
```

## キャプチャ方式

- `chrome.tabCapture` または `getDisplayMedia` で対象タブを録画する。
- `MediaRecorder` はService Workerで動かせないため、Offscreen Documentで実行する。
- `content.js` は `mousedown` の時点で座標・タイムスタンプ・DOM情報を記録する。
- 録画フレームから、クリック時刻の直後または直前の最適フレームを切り出す。
- モーダル消失・メニュー消失・画面遷移直後のズレは、録画フレーム選択で吸収する。

## ステップ検知

クリックイベントで以下を取得する。

```json
{
  "tagName": "BUTTON",
  "text": "保存",
  "ariaLabel": "",
  "placeholder": "",
  "role": "button",
  "columnHeader": "",
  "x": 520,
  "y": 340,
  "timestamp": 123456.78
}
```

方針:

- 具体的な値ではなく、操作の種類・役割を優先して取得する。
- 人名、ID、金額などの値は説明文に直接使わない。
- 例: `山田太郎` ではなく `担当者カードをクリックします`。

## 出力形式の自動判別

| 条件 | 出力 |
| --- | --- |
| クリック間隔1秒以内の連続操作 | GIF（3〜5秒） |
| 単発クリック・画面遷移 | スクショ + 赤枠 |
| 複数画面またぐフロー | 短い動画（検討中） |

判定閾値は実データで調整する。

## 出力フォーマット

### HTML

優先度高。SharePoint、Notion、社内Wikiへの貼り付けや配布を想定する。

- 縦スクロール型
- 1ステップ1ブロック
- GIFまたはスクショ + 説明文
- SaaSのヘルプページに近い読みやすさを目指す

### Word

社内配布・編集用途向け。

- 表紙
- 1ページ3ステップ
- 左: 画像（赤枠合成済み）
- 右: 説明文 + 補足
- 将来的に会社テンプレート対応

### PowerPoint

教育・説明会用途向け。

- 1ステップ1スライド
- 2〜4ステップ/スライド
- 画像の縦横比は維持する

## 説明文生成

### DOM情報ベース

APIキーなしで動作する標準機能。

- `role`
- `ariaLabel`
- `text`
- `placeholder`
- `columnHeader`
- ページタイトル

これらから説明文を生成する。

例:

- `保存` + `button` → `保存ボタンをクリックします`
- `検索` + `button` → `検索を実行します`
- `担当者` + `columnHeader` → `担当者の項目を選択します`

### AIオプション

未確定。ユーザーがAPIキーを設定した場合のみ有効にする。

- スクショ + DOM情報 + 前後ステップ → 説明文生成
- キー未設定時はDOM情報ベースへフォールバック
- 目標は9割自動、1割手修正

## 現行MVPからの移行方針

1. 現行のJSON + PNG読み込み、編集、Word/PPT出力は維持する。
2. Chrome拡張にOffscreen Documentを追加する。
3. 記録開始時にタブ録画を開始する。
4. クリックごとの `captureVisibleTab()` は廃止し、クリックイベントJSONのみ記録する。
5. 記録停止時に `events.json` と `recording.webm` を保存する。
6. Webアプリ側でWebMからフレーム切り出し、PNG/GIF生成を行う。
7. 安定後、従来のPNG直接読込は互換機能として残す。

## 未確定事項

- GIF/動画の切り出し精度とクリック時刻の許容範囲
- HTML出力を単一HTMLファイルにするか、画像フォルダ付きZIPにするか
- DOM情報だけで説明文品質が足りるか
- Word以外にPDF出力需要があるか
- `chrome.tabCapture` と `getDisplayMedia` のどちらを標準にするか
