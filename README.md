# 配送業務日報

スマホとPCからURLで開くための静的Webアプリです。

## 公開方法

条件なしでスマホから開くには、GitHub Pages、Netlify、Vercelなどの静的ホスティングに置きます。

このフォルダはそのまま公開できます。

- `index.html`: アプリ本体
- `manifest.webmanifest`: スマホのホーム画面追加用設定
- `sw.js`: 公開URLで開いた後のオフライン用キャッシュ
- `icon.svg`: アプリアイコン

## Google Sheets連携

端末内の自動保存に加えて、確定保存した日報とメモをGoogle Sheetsへ同期できます。

1. `google-apps-script/Code.gs` をApps Scriptへ登録します。
2. スクリプトプロパティとWeb Appを設定します。
3. アプリの「保存管理 → Google Sheets連携」にWeb App URLと端末キーを入力します。

詳しい導入手順は `google-apps-script/README.md` を参照してください。通信できない場合も端末内の保存は継続し、未同期データは後から再送できます。
