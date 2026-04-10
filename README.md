# Inputto Firestore

Google Apps Script ベースだった入力画面を、GitHub Pages で公開できる Firestore 版へ切り替えるための土台です。

## ねらい

- `1符号 = 1ドキュメント` にして保存単位を小さくする
- 工事一覧では `projects` だけ読む
- 図面を開いた時だけ、その図面の `symbols` を読む
- 保存は開いている図面分だけ差分更新する
- ラベルは `略称` を必ず使い、`焼付色` を出せるようにする

## Firestore モデル

- `environments/{env}/projects/{c2}`
  - 工事番号、現場名、略称、担当、図面数、符号数
- `environments/{env}/projects/{c2}/drawings/{drawingId}`
  - 図面番号、図面状態、担当、件数、プレビュー用符号
- `environments/{env}/projects/{c2}/symbols/{symbolId}`
  - 符号、品名、フロア、L/R、両開き、勝手なし、W/H、枠見込、DW、DH、内外、ラベル枚数、焼付色、断熱、バラ図日付、出荷日

## 画像から拾った登録対象

- 工事番号
- 現場名
- 略称
- 担当
- 図面番号
- 図面状態
- 符号
- 品名
- フロア
- L / R / 両開き_L親 / 両開き_R親 / 勝手なし
- W / H / 枠見込 / DW(L) / DW(R) / DH
- 内外
- ラベル枚数 / R / 両 / 勝手なし
- 焼付色
- フロアごとの数量
- GW密度 / GW厚み
- RW密度 / RW厚み
- バラ図担当 / バラ図_枠 / バラ図_扉
- 組立完了日_枠 / 組立完了日_扉
- 枠出荷日 / 扉出荷日

## セットアップ

```bash
npm install
npm run dev
```

## GitHub Pages

```bash
npm run build
```

Vite の出力先は `docs/` です。GitHub Pages は `main` ブランチの `docs` フォルダを公開元にするとそのまま出せます。

## Firebase プロジェクト

- projectId: `kategu-sys-v15`
- authDomain: `kategu-sys-v15.firebaseapp.com`

## 重要

- この状態で Firestore を本当に使うには、Firebase 側で Firestore Database を有効化する必要があります。
- Security Rules はまだ安全側に締めていません。GitHub Pages で公開運用する前に、認証方針に合わせて必ず調整してください。
