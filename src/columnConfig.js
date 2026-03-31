import { getElapsedTimeDisplay } from './elapsedTimeUtils.js'

/** 固定列（列ピッカーには含めない）— 表示は行順で 1,2,3… を算出 */
export const FIXED_CUT_NUMBER_COLUMN = { id: 'cutNumber', label: 'Cut#' }

/** 利用可能な全列（id は行データのキーと対応） */
export const ALL_COLUMNS = [
  { id: 'image', label: '画像' },
  { id: 'content', label: '内容' },
  { id: 'scene', label: 'シーン' },
  { id: 'location', label: 'ロケーション' },
  { id: 'onscreenText', label: '画面上テキスト' },
  { id: 'narration', label: 'ナレーション' },
  { id: 'action', label: '動き' },
  { id: 'duration', label: '尺(s)' },
  { id: 'shootingDate', label: '撮影日' },
  { id: 'model', label: 'モデル' },
  { id: 'costume', label: '衣装' },
  { id: 'elapsedTime', label: '時間経過' },
  { id: 'note', label: '備考' },
]

export function getColumnMeta(id) {
  return ALL_COLUMNS.find((c) => c.id === id)
}

/** 列の相対幅（画像列を含む — レイアウト・Excel 幅の配分に使用） */
export const COLUMN_WIDTH_RATIOS = {
  image: 2,
  content: 3,
  /** ロケーション列と同程度の幅 */
  scene: 1.5,
  location: 1.5,
  onscreenText: 1.5,
  narration: 2.25,
  action: 1.5,
  /** ロケーション列と同程度の幅 */
  duration: 1.5,
  /** シーン列と同程度 */
  shootingDate: 1.5,
  model: 1,
  costume: 1,
  elapsedTime: 1,
  note: 1,
}

/** テーブル想定のコンテナ幅（px）— Web 編集画面の参考（Excel 出力とは無関係） */
export const LAYOUT_TOTAL_WIDTH_PX = 1470

/** Excel 列幅換算の基準 px（width = (basePx * ratio) / 7） */
export const EXCEL_BASE_PX = 100

/** Cut# 列の比率（100×0.7/7 ≈ 10 文字幅） */
export const CUT_COLUMN_EXCEL_RATIO = 0.7

/** Excel 列幅（文字単位）と px の換算に使用 */
export const EXCEL_CHAR_WIDTH_FACTOR = 7

/**
 * 画像列の Excel 幅（文字単位）— 固定。画像の物理サイズとは別単位。
 */
export const EXCEL_IMAGE_COLUMN_WIDTH_FIXED = 39.5

/** 列の概算幅（px）= 固定列幅 × 換算（参考・列は画像サイズで変えない） */
export const EXCEL_IMAGE_COLUMN_WIDTH_PX =
  EXCEL_IMAGE_COLUMN_WIDTH_FIXED * EXCEL_CHAR_WIDTH_FACTOR

/**
 * ExcelJS の addImage `ext` は 96dpi ベースのピクセル（lib/xlsx/xform/drawing/ext-xform.js）
 */
export const EXCEL_IMAGE_EXT_DPI = 96

/** 埋め込み画像の物理サイズ（インチ）— 動的スケールしない */
export const EXCEL_IMAGE_WIDTH_INCHES = 3.28
export const EXCEL_IMAGE_HEIGHT_INCHES = 1.85

/** キャンバス論理サイズ / addImage ext（px @ 96dpi）= 物理インチに一致 */
export const EXCEL_IMAGE_DISPLAY_WIDTH_PX = Math.round(
  EXCEL_IMAGE_WIDTH_INCHES * EXCEL_IMAGE_EXT_DPI
)
export const EXCEL_IMAGE_DISPLAY_HEIGHT_PX = Math.round(
  EXCEL_IMAGE_HEIGHT_INCHES * EXCEL_IMAGE_EXT_DPI
)

/**
 * 行高（pt）= 画像の物理の高さ（1" = 72pt）。画像の縦と一致、余白なし。
 */
export const EXCEL_IMAGE_ROW_HEIGHT_POINTS = EXCEL_IMAGE_HEIGHT_INCHES * 72

/**
 * ratio に基づく Excel 列幅（文字単位）
 * width = (EXCEL_BASE_PX * ratio) / EXCEL_CHAR_WIDTH_FACTOR
 */
export function excelColumnWidthFromRatio(ratio) {
  return Math.max(4, (EXCEL_BASE_PX * ratio) / EXCEL_CHAR_WIDTH_FACTOR)
}

export function getColumnRatio(id) {
  return COLUMN_WIDTH_RATIOS[id] ?? 1
}

/**
 * 選択列ごとの Excel 列幅（文字単位）
 * 画像列のみ常に 39.5 固定（動的拡張なし）、それ以外は basePx×ratio/7
 */
export function computeExcelColumnWidths(selectedColumnIds) {
  return selectedColumnIds.map((id) => {
    if (id === 'image') return EXCEL_IMAGE_COLUMN_WIDTH_FIXED
    return excelColumnWidthFromRatio(getColumnRatio(id))
  })
}

/**
 * Excel / 表示用のセル値（image はオブジェクト）
 * @param {{ rows?: unknown[], rowIndex?: number }} [context] elapsedTime 自動計算に必要
 */
export function getRowFieldValue(row, columnId, context) {
  switch (columnId) {
    case 'image':
      return row.image
    case 'content':
      return row.content ?? ''
    case 'scene':
      return row.scene ?? ''
    case 'location':
      return row.location ?? ''
    case 'onscreenText':
      return row.onscreenText ?? ''
    case 'narration':
      return row.narration ?? ''
    case 'action':
      return row.action ?? ''
    case 'duration':
      return row.duration ?? ''
    case 'shootingDate':
      return row.shootingDate ?? ''
    case 'model':
      return row.model ?? ''
    case 'costume':
      return row.costume ?? ''
    case 'elapsedTime':
      if (
        context?.rows &&
        typeof context.rowIndex === 'number'
      ) {
        return getElapsedTimeDisplay(row, context.rowIndex, context.rows)
      }
      return row.elapsedTime ?? ''
    case 'note':
      return row.note ?? ''
    default:
      return ''
  }
}
