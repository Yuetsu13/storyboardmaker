# 絵コンテメーカー（16:9） (Ekonte Maker)

Mac デスクトップ風の Web ベースの絵コンテ（ストーリーボード）ツールです。

## 必要な環境

- Node.js 18+
- npm または yarn

## セットアップと起動

```bash
# 依存関係のインストール
npm install

# 開発サーバー起動（http://localhost:5173 で開きます）
npm run dev
```

ビルドしてプレビューする場合:

```bash
npm run build
npm run preview
```

## 主な機能

- **固定 Cut# 列**: 左端（ドラッグ列の右）に常に「Cut#」を表示（1,2,3…は行の並び順で自動計算。状態には保存しない）
- **列の選択（Column Picker）**: 画像・内容・シーン・ロケーション・動き・尺(s)・備考から選び、右パネルの順が表の左→右の列順になります
- **動的グリッド**: 固定カット番号列の右に、選択した列だけを表示
- **行のドラッグ＆ドロップ**: 左端の ⋮⋮ 列をドラッグして行の並び替え（dnd-kit）
- **編集可能なセル**・**画像ドロップ**（16:9 加工）・**行を追加**
- **Excel 出力**: 選択した列順で .xlsx をダウンロード（画像は埋め込み）

## 技術スタック

- React 18 + Vite
- 関数コンポーネントと hooks
- @dnd-kit（ドラッグ＆ドロップ）
- exceljs（Excel エクスポート）
- プレーン CSS（Tailwind なし）

## プロジェクト構成

```
ekonte/
├── index.html
├── package.json
├── vite.config.js
├── README.md
└── src/
    ├── main.jsx
    ├── App.jsx
    ├── App.css
    ├── columnConfig.js
    └── ColumnPicker.jsx
```

すべてクライアント側で完結する MVP です。バックエンドは不要です。
