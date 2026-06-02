# 業務手順書自動作成ツール

Scribe のような共有機能やクラウド管理機能は持たず、日本企業向けの業務手順書を Word / PowerPoint に最短で出力するためのツールです。

## 構成

- `extension/`: Manifest V3 Chrome 拡張。サイドパネルから記録を開始し、タブ録画とクリックイベントを保存します。
- `web/`: React 製の編集・出力アプリ。記録データを読み込み、手順編集、画像マーキング、Word / PowerPoint 出力を行います。

## 設計メモ

- [録画ファースト方式の設計](docs/recording-first-design.md)

現行版は録画を一次ソースにし、クリック時刻からスクリーンショットを切り出す方式です。モーダル消失や画面遷移タイミングのズレを、クリックごとの静止画撮影ではなく録画フレーム選択で吸収します。

## 開発

```bash
npm install
npm run dev
```

## GitHub Pages

WebアプリはGitHub Pagesで公開します。ローカルホストは開発時だけ必要です。

```text
https://yhtko.github.io/manual-creator/
```

GitHubリポジトリのPages設定は `GitHub Actions` を選択してください。`main` ブランチへpushすると、`.github/workflows/pages.yml` が自動でビルドと公開を行います。

## Chrome 拡張の読み込み

1. Chrome で `chrome://extensions` を開きます。
2. デベロッパーモードを有効にします。
3. 「パッケージ化されていない拡張機能を読み込む」から `extension/` を選択します。
4. 拡張アイコンをクリックすると、右側にサイドパネルが開きます。

## 記録データの使い方

1. サイドパネルで「記録開始」を押します。
2. 対象ページでクリック操作を行います。サイドパネルは開いたまま記録を継続できます。
3. サイドパネルで「記録停止して保存」を押します。
4. `manual-project/` に `events-YYYYMMDD-HHMMSS.json` と `recording-YYYYMMDD-HHMMSS.webm` がダウンロードされます。
5. GitHub PagesのWebアプリを開き、`events-YYYYMMDD-HHMMSS.json` を読み込みます。
6. 続けて `recording-YYYYMMDD-HHMMSS.webm` を読み込むと、各クリック時刻のスクリーンショットが自動生成されます。
7. 説明文を編集し、HTML / Word / PowerPoint を出力します。

## HTML出力

Webアプリの「HTML出力」から、画像と説明文を含む単体HTMLファイルを出力できます。

- 目的、対象者、前提条件、完了条件の概要付き
- 縦スクロール型
- 1ステップ1ブロック
- 赤丸・ステップ番号付き画像
- 3秒以内の連続クリックはWebM動画ブロックとして表示
- SharePoint、Notion、社内Wikiへの貼り付けや配布を想定

GIF変換はブラウザ負荷とファイルサイズが大きいため、現時点では行いません。録画済みの `recording.webm` をHTML内に1回だけ埋め込み、連続操作ブロックごとに再生区間を指定します。

## 説明文生成

DOM情報から初期説明文を生成します。

- 録画内容から、目的、対象者、前提条件、完了条件の初期値を推定
- 保存、検索、追加、編集、削除などのボタン文言を判定
- kintoneの「レコードの一覧」「レコードの詳細」画面を考慮
- 数字ID、メールアドレス、日付などの具体値は汎用表現へ寄せる

## 互換入力

以前の方式で保存した `manual-project.json` と `step_XXX.png` も読み込めます。その場合は、WebアプリでJSONとPNGを選択してください。

## 複数記録の結合

出力済みHTMLを手作業で結合するのではなく、Webアプリ上で記録データを追加してから1つのHTMLとして出力します。

1. 1つ目の `events-*.json` を「JSON読込」で読み込みます。
2. 1つ目の `recording-*.webm` を「WebM読込」で読み込みます。
3. 2つ目以降の `events-*.json` を「JSON追記」で追加します。
4. 追加した記録の `recording-*.webm` を「WebM読込」で読み込みます。
5. 最後に「HTML出力」を押します。
