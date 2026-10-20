# ビジラボ 仕事管理（app/annken-kanri-app）

**アプリはここ（BUSILABO-AI-Team の app/annken-kanri-app）で直す。** main にプッシュすると、GitHub Actions（.github/workflows/publish-annken-kanri-app.yml）がこのフォルダの中身を公開リポジトリ busilabo/annken-kanri-app に写し、https://busilabo.github.io/annken-kanri-app/ に出る（2026-10-20 から）。公開リポジトリを直接直しても、次の公開で上書きされる。

直す前に `docs/案件管理アプリ/設計書.md` を必ず読むこと（仕事が入る道・画面・AI の受け持ち・直すときの決まり）。変えたら設計書も直す。公開する前に `tools/app-check` で本物のデータを使って確かめる。

ここで守ること（公開されるファイル）：
- API キー・トークン・パスワード、お客様の名前や連絡先をこのフォルダのファイルに入れない（そのまま公開される）
- 画面の言葉はやさしい日本語。呼び方は「スタッフ」
- スタッフの手間がいちばん少ない形を選ぶ
- ファイル：`index.html`（画面）／`sync.js`（SharePoint とのやり取り・自動で仕事を作る）／`chat.js`（マネージャー）
