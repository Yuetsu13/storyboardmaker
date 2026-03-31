import { useState, useCallback, useMemo, useRef } from 'react'
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
} from '@dnd-kit/core'
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy,
} from '@dnd-kit/sortable'
import { CSS } from '@dnd-kit/utilities'
import ExcelJS from 'exceljs'
import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'
import {
  cropStateToExportImage,
  DISPLAY_CROP_HEIGHT,
  DISPLAY_CROP_WIDTH,
  loadImageFromUrl,
} from './imageCropUtils.js'
import ImageCropCell from './ImageCropCell.jsx'
import {
  ALL_COLUMNS,
  CUT_COLUMN_EXCEL_RATIO,
  EXCEL_CHAR_WIDTH_FACTOR,
  EXCEL_IMAGE_DISPLAY_HEIGHT_PX,
  EXCEL_IMAGE_DISPLAY_WIDTH_PX,
  EXCEL_IMAGE_ROW_HEIGHT_POINTS,
  FIXED_CUT_NUMBER_COLUMN,
  computeExcelColumnWidths,
  excelColumnWidthFromRatio,
  getColumnMeta,
  getColumnRatio,
  getRowFieldValue,
} from './columnConfig.js'
import { getElapsedTimeDisplay } from './elapsedTimeUtils.js'
import ColumnPicker from './ColumnPicker.jsx'
import HistoryTextareaCell from './HistoryTextareaCell.jsx'
import DurationCell, { hasDurationNumericValue } from './DurationCell.jsx'
import ShootingDateCell from './ShootingDateCell.jsx'
import ScheduleComponent from './ScheduleComponent.jsx'

function AppHeading() {
  return (
    <h1 className="app-heading">
      絵コンテメーカー
      <span className="app-heading-aspect">16:9</span>
    </h1>
  )
}

function formatTime(sec) {
  const total = Math.max(0, Math.floor(Number(sec) || 0))
  const m = Math.floor(total / 60)
  const s = total % 60
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`
}

function parseNaPptxSlideXml(xmlText) {
  const parser = new DOMParser()
  const doc = parser.parseFromString(xmlText, 'application/xml')
  const spNodes = Array.from(doc.getElementsByTagName('p:sp'))
  const shapes = spNodes
    .map((sp) => {
      const text = Array.from(sp.getElementsByTagName('a:t'))
        .map((n) => n.textContent ?? '')
        .join('')
        .trim()
      if (!text) return null
      const off = sp.getElementsByTagName('a:off')[0]
      const x = Number(off?.getAttribute('x') ?? 0)
      const y = Number(off?.getAttribute('y') ?? 0)
      return { text, x, y }
    })
    .filter(Boolean)

  const indexShapes = shapes.filter((s) => /^\d+$/.test(s.text) && s.y > 5_400_000) // 下部の列番号
  const timeShapes = shapes.filter((s) => /^\d{2}:\d{2}$/.test(s.text))
  const narrationShapes = shapes.filter(
    (s) => !/^\d+$/.test(s.text) && !/^\d{2}:\d{2}$/.test(s.text)
  )

  const out = []
  indexShapes.forEach((idxShape) => {
    const idxNo = Number(idxShape.text)
    if (!Number.isFinite(idxNo) || idxNo <= 0) return

    const nearestTime = timeShapes
      .slice()
      .sort((a, b) => Math.abs(a.x - idxShape.x) - Math.abs(b.x - idxShape.x))[0]
    const nearestNarration = narrationShapes
      .slice()
      .sort((a, b) => Math.abs(a.x - idxShape.x) - Math.abs(b.x - idxShape.x))[0]

    const timecode = nearestTime?.text ?? formatTime(0)
    const text = nearestNarration?.text ?? ''
    out.push({ index: idxNo - 1, row: text ? { timecode, text } : null })
  })

  return out
}

function readNaDataFromStorage() {
  try {
    const raw = localStorage.getItem('naData')
    const parsed = JSON.parse(raw || '[]')
    return Array.isArray(parsed) ? parsed : []
  } catch (_e) {
    return []
  }
}

function chunkIntoPages(items, size = 5) {
  const out = []
  for (let i = 0; i < items.length; i += size) {
    out.push(items.slice(i, i + size))
  }
  return out
}

/** NA 原稿：未クリックの空欄は null。表示ページ数（最低1ページ・5枠） */
function getNaPageCount(rows) {
  if (!Array.isArray(rows)) return 1
  let maxIdx = -1
  for (let i = 0; i < rows.length; i++) {
    if (rows[i] != null) maxIdx = i
  }
  const n = Math.max(maxIdx + 1, rows.length, 1)
  return Math.max(1, Math.ceil(n / 5))
}

/** 次のページ分（5枠）ぶん配列を延ばす（未入力スロットは undefined） */
function addNaPageSlots(rows) {
  const next = rows.slice()
  const pageCount = getNaPageCount(rows)
  next.length = (pageCount + 1) * 5
  return next
}

function generateId() {
  return `row-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`
}

function emptyRow() {
  return {
    id: generateId(),
    image: null,
    content: '',
    scene: '',
    location: '',
    onscreenText: '',
    narration: '',
    action: '',
    duration: '',
    shootingDate: '',
    model: '',
    costume: '',
    elapsedTime: '',
    /** true のときは elapsedTime をそのまま表示（尺からの自動計算はしない） */
    elapsedTimeManual: false,
    note: '',
  }
}

const initialRows = [emptyRow(), emptyRow()]

function emptyCallSheetRow() {
  return {
    id: generateId(),
    time: '',
    schedule: '',
    note: '',
    // 画像は最大3列まで（画像1〜3）
    images: [null, null, null, null, null, null, null, null],
    // Cut# は画像列ごとに保持（最大8）
    cutNos: ['', '', '', '', '', '', '', ''],
  }
}

const initialCallSheetRows = [emptyCallSheetRow()]

/** 絵コンテ行から香盤表「カットを選択」用の一覧（Web のみ） */
function buildCallsheetStoryboardSourceItems(storyboardRows) {
  if (!Array.isArray(storyboardRows)) return []
  return storyboardRows
    .map((row, idx) => ({ row, cutNumber: idx + 1 }))
    .filter(({ row }) => row.image)
    .map(({ row, cutNumber }) => ({
      key: `${row.id}-${cutNumber}`,
      cutNumber,
      image: row.image,
    }))
}

function cloneCallSheetRowsForUndo(rows) {
  if (!Array.isArray(rows)) return []
  return rows.map((r) => ({
    ...r,
    images: Array.isArray(r.images) ? [...r.images] : [],
    cutNos: Array.isArray(r.cutNos) ? [...r.cutNos] : [],
  }))
}

/** 香盤表へドロップする用に objectUrl を複製（片方の revoke で他が壊れないようにする） */
async function cloneImageStateForCallSheet(img) {
  if (!img?.objectUrl) return null
  try {
    const res = await fetch(img.objectUrl)
    const blob = await res.blob()
    const newUrl = URL.createObjectURL(blob)
    return {
      ...img,
      objectUrl: newUrl,
      file:
        img.file && img.file instanceof File
          ? new File([blob], img.file.name || 'image.png', { type: blob.type || img.file.type })
          : undefined,
    }
  } catch (e) {
    console.error('cloneImageStateForCallSheet', e)
    return { ...img }
  }
}

/** Excel 描画の colOff/rowOff（EMU）: 96dpi では 1px ≈ 9525 EMU */
const EXCEL_EMU_PER_PX = 9525

/** 画像・各セルがすべて空なら true（Excel 保存の disabled 判定用） */
function isRowEmptyForExport(row) {
  if (row.image) return false
  const keys = [
    'content',
    'scene',
    'location',
    'onscreenText',
    'narration',
    'action',
    'shootingDate',
    'model',
    'costume',
    'elapsedTime',
    'note',
  ]
  for (const k of keys) {
    if (String(row[k] ?? '').trim() !== '') return false
  }
  if (hasDurationNumericValue(row.duration)) return false
  return true
}

function storyboardHasExportableRow(rows) {
  return Array.isArray(rows) && rows.some((row) => !isRowEmptyForExport(row))
}

function excelCellValueToString(value) {
  if (value == null) return ''
  if (typeof value === 'string') return value
  if (typeof value === 'number') return String(value)
  // exceljs: RichText / Hyperlink / その他オブジェクト系の可能性
  if (typeof value === 'object') {
    if (typeof value.text === 'string') return value.text
    if (Array.isArray(value.richText)) return value.richText.map((t) => t.text).join('')
    if (value.result != null) return String(value.result)
  }
  return String(value)
}

function parseHHmmToMinutesForCallSheet(s) {
  const t = String(s ?? '').trim()
  const m = t.match(/^(\d{1,2}):(\d{2})$/)
  if (!m) return null
  const h = parseInt(m[1], 10)
  const min = parseInt(m[2], 10)
  if (Number.isNaN(h) || Number.isNaN(min) || h < 0 || h > 23 || min < 0 || min > 59) return null
  return h * 60 + min
}

function durationLabelFromTimesForCallSheet(curTime, nextTime) {
  const cur = parseHHmmToMinutesForCallSheet(curTime)
  const next = parseHHmmToMinutesForCallSheet(nextTime)
  if (cur == null || next == null) return ''

  let diff = next - cur
  let dayPrefix = ''
  if (diff < 0) {
    diff += 24 * 60
    dayPrefix = '1day '
  }
  if (diff <= 0) return ''

  const h = Math.floor(diff / 60)
  const m = diff % 60
  let base = ''
  if (h > 0 && m > 0) base = `${h}h ${m}min`
  else if (h > 0) base = `${h}h`
  else base = `${m}min`
  return `${dayPrefix}${base}`
}

function stripCallSheetDurationSuffixText(s) {
  const text = String(s ?? '')
  return text
    .replace(/\s*(?:1day\s+)?\d+h\s*\d+min\s*$/i, '')
    .replace(/\s*(?:1day\s+)?\d+h\s*$/i, '')
    .replace(/\s*(?:1day\s+)?\d+min\s*$/i, '')
    .trimEnd()
}

function buildGridTemplateColumns(selectedColumnIds) {
  const parts = ['44px', '56px']
  selectedColumnIds.forEach((id) => {
    if (id === 'image') {
      /* 内容列と同じ 2fr。minmax(0,…) で画像の min-content が列を押し広げない */
      parts.push('minmax(0, 2fr)')
    } else {
      const r = getColumnRatio(id)
      parts.push(`minmax(0,${r}fr)`)
    }
  })
  return parts.join(' ')
}

function SortableRow({
  row,
  rowIndex,
  rows,
  firstEmptyImageRowIndex,
  displayIndex,
  selectedColumnIds,
  gridTemplateColumns,
  sceneOptions,
  locationOptions,
  modelOptions,
  costumeOptions,
  onUpdate,
  onDelete,
  onImageDrop,
  onImageCropUpdate,
}) {
  const showImagePlaceholderHint =
    !row.image &&
    firstEmptyImageRowIndex >= 0 &&
    rowIndex === firstEmptyImageRowIndex

  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: row.id })

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  }

  const handleContentChange = (e) => onUpdate(row.id, { content: e.target.value })
  const handleActionChange = (e) => onUpdate(row.id, { action: e.target.value })
  const handleNoteChange = (e) => onUpdate(row.id, { note: e.target.value })

  const handleImageDrop = async (e) => {
    e.preventDefault()
    const file = e.dataTransfer?.files?.[0]
    if (!file || !file.type.startsWith('image/')) return
    const objectUrl = URL.createObjectURL(file)
    try {
      const img = await loadImageFromUrl(objectUrl)
      onImageDrop(row.id, {
        file,
        objectUrl,
        naturalWidth: img.naturalWidth,
        naturalHeight: img.naturalHeight,
        initialScale: 0,
        scale: 0,
        offsetX: 0,
        offsetY: 0,
      })
    } catch (err) {
      console.error('Image load failed:', err)
      URL.revokeObjectURL(objectUrl)
    }
  }
  const handleImageDragOver = (e) => e.preventDefault()
  const handleImageClear = () => onImageDrop(row.id, null)

  const renderCell = (columnId) => {
    switch (columnId) {
      case 'image':
        return (
          <div
            key="image"
            className="cell cell-image"
            onDrop={handleImageDrop}
            onDragOver={handleImageDragOver}
            aria-label={
              row.image
                ? undefined
                : showImagePlaceholderHint
                  ? '画像をドロップ。選択してスクロールまたはドラッグで調整。'
                  : '画像をドロップ'
            }
          >
            {row.image ? (
              <ImageCropCell
                image={row.image}
                onChange={(patch) => onImageCropUpdate(row.id, patch)}
                onClear={handleImageClear}
              />
            ) : showImagePlaceholderHint ? (
              <div className="image-placeholder" aria-hidden="true">
                <span className="image-placeholder__line image-placeholder__line--primary">
                  画像をドロップ
                </span>
                <span className="image-placeholder__line">選択してスクロール/ドラッグで調整</span>
              </div>
            ) : null}
          </div>
        )
      case 'content':
        return (
          <div className="cell cell-description" key="content">
            <textarea
              className="cell-textarea"
              value={row.content ?? ''}
              onChange={handleContentChange}
            />
          </div>
        )
      case 'scene':
        return (
          <div className="cell cell-description" key="scene">
            <HistoryTextareaCell
              value={row.scene ?? ''}
              onChange={(v) => onUpdate(row.id, { scene: v })}
              options={sceneOptions}
            />
          </div>
        )
      case 'location':
        return (
          <div className="cell cell-description" key="location">
            <HistoryTextareaCell
              value={row.location ?? ''}
              onChange={(v) => onUpdate(row.id, { location: v })}
              options={locationOptions}
            />
          </div>
        )
      case 'onscreenText':
        return (
          <div className="cell cell-description" key="onscreenText">
            <textarea
              className="cell-textarea"
              value={row.onscreenText ?? ''}
              onChange={(e) => onUpdate(row.id, { onscreenText: e.target.value })}
            />
          </div>
        )
      case 'narration':
        return (
          <div className="cell cell-description" key="narration">
            <textarea
              className="cell-textarea"
              value={row.narration ?? ''}
              onChange={(e) => onUpdate(row.id, { narration: e.target.value })}
            />
          </div>
        )
      case 'action':
        return (
          <div className="cell cell-description" key="action">
            <textarea
              className="cell-textarea"
              value={row.action ?? ''}
              onChange={handleActionChange}
            />
          </div>
        )
      case 'duration':
        return (
          <div
            className={`cell cell-description cell-duration${
              hasDurationNumericValue(row.duration) ? ' cell-duration--has-value' : ''
            }`}
            key="duration"
          >
            <DurationCell
              value={row.duration ?? ''}
              onChange={(v) => onUpdate(row.id, { duration: v })}
            />
          </div>
        )
      case 'shootingDate':
        return (
          <div className="cell cell-description" key="shootingDate">
            <ShootingDateCell
              value={row.shootingDate ?? ''}
              onChange={(v) => onUpdate(row.id, { shootingDate: v })}
            />
          </div>
        )
      case 'model':
        return (
          <div className="cell cell-description" key="model">
            <HistoryTextareaCell
              value={row.model ?? ''}
              onChange={(v) => onUpdate(row.id, { model: v })}
              options={modelOptions}
            />
          </div>
        )
      case 'costume':
        return (
          <div className="cell cell-description" key="costume">
            <HistoryTextareaCell
              value={row.costume ?? ''}
              onChange={(v) => onUpdate(row.id, { costume: v })}
              options={costumeOptions}
            />
          </div>
        )
      case 'elapsedTime':
        return (
          <div className="cell cell-description" key="elapsedTime">
            <textarea
              className="cell-textarea"
              value={getElapsedTimeDisplay(row, rowIndex, rows)}
              onChange={(e) => {
                const v = e.target.value
                onUpdate(row.id, {
                  elapsedTime: v,
                  elapsedTimeManual: v.trim() !== '',
                })
              }}
            />
          </div>
        )
      case 'note':
        return (
          <div className="cell cell-description" key="note">
            <textarea
              className="cell-textarea"
              value={row.note ?? ''}
              onChange={handleNoteChange}
            />
          </div>
        )
      default:
        return null
    }
  }

  return (
    <div
      ref={setNodeRef}
      style={{ ...style, gridTemplateColumns }}
      className={`storyboard-row ${isDragging ? 'dragging' : ''}`}
    >
      <div className="cell cell-drag-only" aria-hidden>
        <span className="drag-handle" {...attributes} {...listeners}>⋮⋮</span>
      </div>
      <div className="cell cell-cut-number" title="Cut#（行順）">
        <span>{displayIndex}</span>
        <button
          type="button"
          className="row-delete-btn"
          onClick={() => onDelete(row.id)}
          aria-label={`${displayIndex}行目を削除`}
          title="この行を削除"
        >
          ×
        </button>
      </div>
      {selectedColumnIds.map((id) => renderCell(id))}
    </div>
  )
}

function CallSheetRow({ row, index, onUpdate, onDelete, onImageDrop, onImageCropUpdate }) {
  const [editingCutIndex, setEditingCutIndex] = useState(null)
  const scheduleImageCount = 1

  const handleImageDrop = async (imageIndex, e) => {
    e.preventDefault()
    const file = e.dataTransfer?.files?.[0]
    if (!file || !file.type.startsWith('image/')) return
    const objectUrl = URL.createObjectURL(file)
    try {
      const img = await loadImageFromUrl(objectUrl)
      onImageDrop(row.id, imageIndex, {
        file,
        objectUrl,
        naturalWidth: img.naturalWidth,
        naturalHeight: img.naturalHeight,
        initialScale: 0,
        scale: 0,
        offsetX: 0,
        offsetY: 0,
      })
    } catch (err) {
      console.error('Image load failed:', err)
      URL.revokeObjectURL(objectUrl)
    }
  }

  return (
    <div className="callsheet-row">
      <button
        type="button"
        className="row-delete-btn callsheet-delete-btn"
        onClick={() => onDelete(row.id)}
        aria-label={`${index + 1}行目を削除`}
        title="この行を削除"
      >
        ×
      </button>

      <input
        className="callsheet-input callsheet-time"
        value={row.time}
        onChange={(e) => onUpdate(row.id, { time: e.target.value })}
      />

      <div className="callsheet-schedule-cell">
        <textarea
          className="callsheet-input callsheet-schedule-text"
          value={row.schedule}
          onChange={(e) => onUpdate(row.id, { schedule: e.target.value })}
        />

        <div className="callsheet-schedule-images">
          {Array.from({ length: scheduleImageCount }).map((_, imageIndex) => {
            const img = row.images?.[imageIndex] ?? null
            const cutStr = String(row.cutNos?.[imageIndex] ?? '')
            const cutDigits = Math.max(2, cutStr.trim().length)

            return (
              <div
                key={`callsheet-schedule-img-${index}-${imageIndex}`}
                className="callsheet-image-cell callsheet-schedule-image-cell"
                onDrop={(e) => handleImageDrop(imageIndex, e)}
                onDragOver={(e) => e.preventDefault()}
              >
                {img ? (
                  <div className="callsheet-image-wrap">
                    <ImageCropCell
                      image={img}
                      onChange={(patch) => onImageCropUpdate(row.id, imageIndex, patch)}
                      onClear={() => {
                        setEditingCutIndex(null)
                        onImageDrop(row.id, imageIndex, null)
                      }}
                    />
                    <div
                      className="callsheet-cut-overlay"
                      onClick={(e) => {
                        e.stopPropagation()
                        setEditingCutIndex(imageIndex)
                      }}
                      role="button"
                      tabIndex={0}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' || e.key === ' ') {
                          e.preventDefault()
                          setEditingCutIndex(imageIndex)
                        }
                      }}
                      aria-label={`Cut#${imageIndex + 1}を編集`}
                      title="クリックでCut番号を編集"
                    >
                      <span className="callsheet-cut-label">Cut# </span>
                      {editingCutIndex === imageIndex ? (
                        <input
                          className="callsheet-cut-input"
                          autoFocus
                          value={row.cutNos?.[imageIndex] ?? ''}
                          size={cutDigits}
                          onChange={(e) => {
                            const next = (row.cutNos ?? ['', '', '', '', '', '', '', '']).slice()
                            next[imageIndex] = e.target.value
                            onUpdate(row.id, { cutNos: next })
                          }}
                          onClick={(e) => e.stopPropagation()}
                          onBlur={() => setEditingCutIndex(null)}
                        />
                      ) : (
                        <span className="callsheet-cut-value">{row.cutNos?.[imageIndex] ?? ''}</span>
                      )}
                    </div>
                  </div>
                ) : (
                  <span className="callsheet-image-placeholder">画像をドロップ</span>
                )}
              </div>
            )
          })}
        </div>
      </div>

      <textarea
        className="callsheet-input callsheet-note"
        value={row.note}
        onChange={(e) => onUpdate(row.id, { note: e.target.value })}
      />
    </div>
  )
}

export default function App() {
  const [page, setPage] = useState('home')
  const [naMode, setNaMode] = useState('new')
  const [manualNaRows, setManualNaRows] = useState([])
  const [excelNaRows, setExcelNaRows] = useState([])
  const [manualNaUndoStack, setManualNaUndoStack] = useState([])
  const [naDataTick, setNaDataTick] = useState(0)
  const [screen, setScreen] = useState('picker')
  const [selectedColumnIds, setSelectedColumnIds] = useState([])
  const [rows, setRows] = useState(initialRows)
  const [callSheetRows, setCallSheetRows] = useState(initialCallSheetRows)
  const [callSheetDurationHiddenByRow, setCallSheetDurationHiddenByRow] = useState({})
  const [callSheetUndoStack, setCallSheetUndoStack] = useState([])
  const naPptImportInputRef = useRef(null)
  /** 絵コンテから遷移したときのみセット（ホームからは null）。Web のカットプール表示用 */
  const [callsheetStoryboardSourceItems, setCallsheetStoryboardSourceItems] = useState(null)
  const [callSheetImageColumnCount, setCallSheetImageColumnCount] = useState(1)
  const callSheetGridTemplateColumns = useMemo(() => {
    // 列順: [削除] [TIME] [スケジュール] [画像1..N] [備考]
    // - 画像列が 1〜5 の間は「画像列幅は不変」、その分備考が縮む
    // - 画像列が 6 以上になったら、画像列側も縮む
    const n = callSheetImageColumnCount

    const baseLeft = '36px 84px'
    const timeCol = 'minmax(210px, 1fr)'
    const scheduleCol = 'minmax(105px, 1fr)'
    const noteCol = 'minmax(180px, 1fr)'

    if (n <= 5) {
      const fixedImgW = 'minmax(160px, 160px)'
      const imgCols = Array.from({ length: n }).map(() => fixedImgW).join(' ')
      return `${baseLeft} ${timeCol} ${scheduleCol} ${imgCols} ${noteCol}`
    }

    // 6列目以降: 画像列も可変で縮む（fr配分）
    const imgCols = Array.from({ length: n }).map(() => 'minmax(80px, 1fr)').join(' ')
    return `${baseLeft} ${timeCol} ${scheduleCol} ${imgCols} ${noteCol}`
  }, [callSheetImageColumnCount])
  /** 行削除の「戻す」用（直近から LIFO）。画像は戻すまで objectUrl を維持する */
  const [deleteUndoStack, setDeleteUndoStack] = useState([])

  const gridTemplate = useMemo(
    () => buildGridTemplateColumns(selectedColumnIds),
    [selectedColumnIds]
  )

  /** 画像列で上から最初の空の画像セル（説明文はこの1行だけに表示） */
  const firstEmptyImageRowIndex = useMemo(
    () => rows.findIndex((r) => !r.image),
    [rows]
  )

  /** いずれかの行に入力済みのシーン名（重複除き・ソート）— 候補用 */
  const sceneHistoryOptions = useMemo(() => {
    const seen = new Set()
    const out = []
    for (const r of rows) {
      const s = (r.scene ?? '').trim()
      if (s && !seen.has(s)) {
        seen.add(s)
        out.push(s)
      }
    }
    out.sort((a, b) => a.localeCompare(b, 'ja'))
    return out
  }, [rows])

  /** いずれかの行に入力済みのロケーション（重複除き・ソート）— 候補用 */
  const locationHistoryOptions = useMemo(() => {
    const seen = new Set()
    const out = []
    for (const r of rows) {
      const s = (r.location ?? '').trim()
      if (s && !seen.has(s)) {
        seen.add(s)
        out.push(s)
      }
    }
    out.sort((a, b) => a.localeCompare(b, 'ja'))
    return out
  }, [rows])

  /** いずれかの行に入力済みのモデル名（重複除き・ソート）— 候補用 */
  const modelHistoryOptions = useMemo(() => {
    const seen = new Set()
    const out = []
    for (const r of rows) {
      const s = (r.model ?? '').trim()
      if (s && !seen.has(s)) {
        seen.add(s)
        out.push(s)
      }
    }
    out.sort((a, b) => a.localeCompare(b, 'ja'))
    return out
  }, [rows])

  /** いずれかの行に入力済みの衣装（重複除き・ソート）— 候補用 */
  const costumeHistoryOptions = useMemo(() => {
    const seen = new Set()
    const out = []
    for (const r of rows) {
      const s = (r.costume ?? '').trim()
      if (s && !seen.has(s)) {
        seen.add(s)
        out.push(s)
      }
    }
    out.sort((a, b) => a.localeCompare(b, 'ja'))
    return out
  }, [rows])

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: { distance: 8 },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  )

  const handleAddColumn = useCallback((id) => {
    setSelectedColumnIds((prev) => [...prev, id])
  }, [])

  const handleRemoveColumn = useCallback((id) => {
    setSelectedColumnIds((prev) => prev.filter((x) => x !== id))
  }, [])

  const handleReorderSelectedColumns = useCallback((newIds) => {
    setSelectedColumnIds(newIds)
  }, [])

  const handleImportExcel = useCallback(
    async (file) => {
      if (!file) return
      try {
        const arrayBuffer = await file.arrayBuffer()
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(arrayBuffer)

        const worksheet =
          workbook.getWorksheet('絵コンテ') ?? workbook.worksheets?.[0] ?? null
        if (!worksheet) {
          throw new Error('対応するシート（絵コンテ）が見つかりませんでした')
        }

        // 1行目: ヘッダー
        const headerRow = worksheet.getRow(1)
        const headerValues = headerRow.values ?? []

        const labelToId = new Map(ALL_COLUMNS.map((c) => [c.label, c.id]))
        const importedSelectedColumnIds = []

        // A列(Cut#)の次（B列=2）から左→右の列を復元
        for (let sheetCol = 2; sheetCol < headerValues.length; sheetCol++) {
          const raw = headerValues[sheetCol]
          if (raw == null || raw === '') continue
          const label = excelCellValueToString(raw).trim()
          const id = labelToId.get(label)
          if (!id) continue
          importedSelectedColumnIds.push(id)
        }

        if (importedSelectedColumnIds.length === 0) {
          throw new Error('列ヘッダーを読み取れませんでした')
        }

        // データ行数（2行目〜）
        const lastRowNumber = worksheet.lastRow?.number ?? worksheet.rowCount ?? 1
        const dataRowCount = Math.max(0, lastRowNumber - 1)

        const newRows = Array.from({ length: dataRowCount }, () => emptyRow())
        for (let dataIdx = 0; dataIdx < dataRowCount; dataIdx++) {
          const excelRowNumber = dataIdx + 2 // 2行目からデータ
          const r = newRows[dataIdx]

          importedSelectedColumnIds.forEach((colId, colIdx) => {
            if (colId === 'image') return
            // sheet: 0-basedで Cut# が 0列目、excel getCell は 1-based
            const excelColNumber = 2 + colIdx // B列(2)が selectedColumnIds[0]
            const cell = worksheet.getCell(excelRowNumber, excelColNumber)
            r[colId] = excelCellValueToString(cell?.value)
          })
        }

        // 埋め込み画像を復元
        const imageColIdx = importedSelectedColumnIds.indexOf('image')
        if (imageColIdx >= 0) {
          const expectedImageCol0Based = imageColIdx + 1 // Cut# が 0列目なので +1
          const images = worksheet.getImages ? worksheet.getImages() : []
          const imagesInCol = images.filter((im) => {
            const nativeCol = Number(im?.range?.tl?.nativeCol)
            return Number.isFinite(nativeCol) && nativeCol === expectedImageCol0Based
          })

          // ExcelJS の画像アンカー基準（row の 0/1 始まり等）がファイルによってズレることがあるため、
          // 次の優先順位で offset を推定します。
          // 1) dataIdx=0（先頭行）に画像が載る（placeholder にならない）ものを最優先
          // 2) 次に、データ行内に入る画像数が最大のもの
          const candidateOffsets = [0, 1, 2, 3]
          let bestOffset = 1
          let bestCount0 = -1
          let bestCountInRange = -1
          for (const offset of candidateOffsets) {
            let count0 = 0
            let countInRange = 0
            for (const im of imagesInCol) {
              const nativeRow = Number(im?.range?.tl?.nativeRow)
              if (!Number.isFinite(nativeRow)) continue
              const dataIdx = nativeRow - offset
              if (dataIdx === 0) count0 += 1
              if (dataIdx >= 0 && dataIdx < newRows.length) countInRange += 1
            }
            if (count0 > bestCount0) {
              bestCount0 = count0
              bestCountInRange = countInRange
              bestOffset = offset
            } else if (count0 === bestCount0 && countInRange > bestCountInRange) {
              bestCountInRange = countInRange
              bestOffset = offset
            }
          }

          for (const im of imagesInCol) {
            const nativeRow = Number(im?.range?.tl?.nativeRow)
            if (!Number.isFinite(nativeRow)) continue

            const dataIdx = nativeRow - bestOffset
            if (dataIdx < 0 || dataIdx >= newRows.length) continue

            // 同じ行に複数画像が割り当たる場合は、先に入ったものを優先（後続で上書しない）
            if (newRows[dataIdx].image) continue

            const imageIdNum = Number(im.imageId)
            const medium = workbook.getImage(imageIdNum)
            if (!medium?.buffer) continue

            const extension = medium.extension || 'png'
            const blob = new Blob([medium.buffer], { type: `image/${extension}` })
            const objectUrl = URL.createObjectURL(blob)

            const img = await loadImageFromUrl(objectUrl)
            newRows[dataIdx].image = {
              file: null,
              objectUrl,
              naturalWidth: img.naturalWidth,
              naturalHeight: img.naturalHeight,
              initialScale: 0,
              scale: 0,
              offsetX: 0,
              offsetY: 0,
            }
          }
        }

        // 既存の objectUrl を解放
        rows.forEach((r) => {
          if (r.image?.objectUrl) URL.revokeObjectURL(r.image.objectUrl)
        })

        setSelectedColumnIds(importedSelectedColumnIds)
        setRows(newRows)
        setDeleteUndoStack([])
        setScreen('editor')
      } catch (e) {
        // eslint-disable-next-line no-console
        console.error('Excel import failed:', e)
        alert(e?.message ?? 'Excelの読み込みに失敗しました')
      }
    },
    [rows]
  )

  const handleConfirmColumns = useCallback(() => {
    setScreen('editor')
  }, [])

  const handleImportCallSheetExcel = useCallback(
    async (file) => {
      if (!file) return
      try {
        const arrayBuffer = await file.arrayBuffer()
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(arrayBuffer)

        const worksheet = workbook.getWorksheet('香盤表') ?? workbook.worksheets?.[0] ?? null
        if (!worksheet) throw new Error('対応するシート（香盤表）が見つかりませんでした')

        const lastRowNumber = worksheet.lastRow?.number ?? worksheet.rowCount ?? 1
        const parsed = []
        for (let titleRow = 2; titleRow <= lastRowNumber; titleRow += 3) {
          const cutRow = titleRow + 1
          const imageRow = titleRow + 2
          const row = emptyCallSheetRow()

          row.time = excelCellValueToString(worksheet.getCell(titleRow, 1)?.value).trim()
          row.schedule = stripCallSheetDurationSuffixText(
            excelCellValueToString(worksheet.getCell(titleRow, 2)?.value)
          )
          row.note = excelCellValueToString(worksheet.getCell(titleRow, 3)?.value)

          const cutText = excelCellValueToString(worksheet.getCell(cutRow, 2)?.value)
          const cutTokens = cutText
            .split(',')
            .map((v) => String(v ?? '').trim())
            .filter(Boolean)
            .map((v) => v.replace(/^cut#\s*/i, ''))

          parsed.push({ row, imageRow, cutTokens })
        }

        const rowsWithAtLeastOne = parsed.length > 0 ? parsed : [{ row: emptyCallSheetRow(), imageRow: 4, cutTokens: [] }]

        const imageGroups = new Map()
        const nativeRowToParsedIdx = new Map(
          rowsWithAtLeastOne.map((item, idx) => [item.imageRow - 1, idx])
        )
        const images = worksheet.getImages ? worksheet.getImages() : []
        images.forEach((im) => {
          const nativeRow = Number(im?.range?.tl?.nativeRow)
          if (!Number.isFinite(nativeRow)) return
          const parsedIdx = nativeRowToParsedIdx.get(nativeRow)
          if (parsedIdx == null) return
          if (!imageGroups.has(parsedIdx)) imageGroups.set(parsedIdx, [])
          imageGroups.get(parsedIdx).push(im)
        })

        for (const [parsedIdx, group] of imageGroups.entries()) {
          const ordered = group
            .slice()
            .sort((a, b) => Number(a?.range?.tl?.nativeColOff || 0) - Number(b?.range?.tl?.nativeColOff || 0))
          for (let slot = 0; slot < ordered.length && slot < 8; slot++) {
            const im = ordered[slot]
            const imageIdNum = Number(im.imageId)
            const medium = workbook.getImage(imageIdNum)
            if (!medium?.buffer) continue
            const extension = medium.extension || 'png'
            const blob = new Blob([medium.buffer], { type: `image/${extension}` })
            const objectUrl = URL.createObjectURL(blob)
            const img = await loadImageFromUrl(objectUrl)
            rowsWithAtLeastOne[parsedIdx].row.images[slot] = {
              file: null,
              objectUrl,
              naturalWidth: img.naturalWidth,
              naturalHeight: img.naturalHeight,
              containerWidth: DISPLAY_CROP_WIDTH,
              containerHeight: DISPLAY_CROP_HEIGHT,
              initialScale: 0,
              scale: 0,
              offsetX: 0,
              offsetY: 0,
              crop: null,
            }
            rowsWithAtLeastOne[parsedIdx].row.cutNos[slot] = rowsWithAtLeastOne[parsedIdx].cutTokens[slot] ?? ''
          }
        }

        callSheetRows.forEach((r) => {
          ;(r.images ?? []).forEach((im) => {
            if (im?.objectUrl) URL.revokeObjectURL(im.objectUrl)
          })
        })

        setCallSheetRows(rowsWithAtLeastOne.map((x) => x.row))
        setCallSheetUndoStack([])
        setCallSheetDurationHiddenByRow({})
      } catch (e) {
        console.error('CallSheet Excel import failed:', e)
        alert(e?.message ?? '香盤表Excelの読み込みに失敗しました')
      }
    },
    [callSheetRows]
  )

  const handleDragEnd = useCallback((event) => {
    const { active, over } = event
    if (over && active.id !== over.id) {
      setRows((prev) => {
        const oldIndex = prev.findIndex((r) => r.id === active.id)
        const newIndex = prev.findIndex((r) => r.id === over.id)
        if (oldIndex === -1 || newIndex === -1) return prev
        return arrayMove(prev, oldIndex, newIndex)
      })
    }
  }, [])

  const handleUpdate = useCallback((id, patch) => {
    setRows((prev) =>
      prev.map((r) => (r.id === id ? { ...r, ...patch } : r))
    )
  }, [])

  const handleImageDrop = useCallback((id, imageData) => {
    setRows((prev) =>
      prev.map((r) => {
        if (r.id !== id) return r
        if (r.image?.objectUrl && imageData?.objectUrl !== r.image.objectUrl) {
          URL.revokeObjectURL(r.image.objectUrl)
        }
        if (imageData === null && r.image?.objectUrl) {
          URL.revokeObjectURL(r.image.objectUrl)
        }
        return { ...r, image: imageData }
      })
    )
  }, [])

  const handleImageCropUpdate = useCallback((id, patch) => {
    setRows((prev) =>
      prev.map((r) =>
        r.id === id && r.image ? { ...r, image: { ...r.image, ...patch } } : r
      )
    )
  }, [])

  const handleCallSheetUpdate = useCallback((id, patch) => {
    setCallSheetRows((prev) => {
      setCallSheetUndoStack((stack) => [
        ...stack.slice(-49),
        {
          rows: cloneCallSheetRowsForUndo(prev),
          durationHiddenByRow: { ...callSheetDurationHiddenByRow },
        },
      ])
      return prev.map((r) => (r.id === id ? { ...r, ...patch } : r))
    })
  }, [callSheetDurationHiddenByRow])

  const handleCallSheetReorderRows = useCallback((activeId, overId) => {
    if (!activeId || !overId || activeId === overId) return
    setCallSheetRows((prev) => {
      const oldIndex = prev.findIndex((r) => r.id === activeId)
      const newIndex = prev.findIndex((r) => r.id === overId)
      if (oldIndex === -1 || newIndex === -1) return prev
      setCallSheetUndoStack((stack) => [
        ...stack.slice(-49),
        {
          rows: cloneCallSheetRowsForUndo(prev),
          durationHiddenByRow: { ...callSheetDurationHiddenByRow },
        },
      ])
      const next = arrayMove(prev, oldIndex, newIndex)
      const changedFrom = Math.min(oldIndex, newIndex)
      // 並び替えで影響が出る起点行を含めて TIME をクリアする
      return next.map((row, idx) => (idx >= changedFrom ? { ...row, time: '' } : row))
    })
  }, [callSheetDurationHiddenByRow])

  const handleUndoCallSheet = useCallback(() => {
    setCallSheetUndoStack((stack) => {
      if (stack.length === 0) return stack
      const last = stack[stack.length - 1]
      setCallSheetRows(cloneCallSheetRowsForUndo(last.rows))
      setCallSheetDurationHiddenByRow({ ...(last.durationHiddenByRow ?? {}) })
      return stack.slice(0, -1)
    })
  }, [])

  const handleCallSheetDurationHiddenChange = useCallback((rowId, hiddenLabel) => {
    setCallSheetDurationHiddenByRow((prev) => {
      const next = { ...prev }
      if (!hiddenLabel) delete next[rowId]
      else next[rowId] = hiddenLabel
      return next
    })
  }, [])

  const handleCallSheetImageDrop = useCallback((id, imageIndex, imageData) => {
    setCallSheetRows((prev) =>
      {
        setCallSheetUndoStack((stack) => [
          ...stack.slice(-49),
          {
            rows: cloneCallSheetRowsForUndo(prev),
            durationHiddenByRow: { ...callSheetDurationHiddenByRow },
          },
        ])
        return prev.map((r) => {
        if (r.id !== id) return r
        const nextImages = (r.images ?? [null, null, null, null, null, null, null, null]).slice()
        const nextCutNos = (r.cutNos ?? ['', '', '', '', '', '', '', '']).slice()

        const prevImg = nextImages[imageIndex]
        if (prevImg?.objectUrl && imageData?.objectUrl !== prevImg.objectUrl) {
          URL.revokeObjectURL(prevImg.objectUrl)
        }

        nextImages[imageIndex] = imageData
        if (imageData === null) nextCutNos[imageIndex] = ''

        return { ...r, images: nextImages, cutNos: nextCutNos }
      })
      }
    )
  }, [callSheetDurationHiddenByRow])

  // 複数画像を「置き換え」ではなく「配列へ追加」する（未使用スロットに順に詰める）
  const handleCallSheetAddImages = useCallback((id, imageStates, maxImages = 8) => {
    if (!Array.isArray(imageStates) || imageStates.length === 0) return
    setCallSheetRows((prev) =>
      {
        setCallSheetUndoStack((stack) => [
          ...stack.slice(-49),
          {
            rows: cloneCallSheetRowsForUndo(prev),
            durationHiddenByRow: { ...callSheetDurationHiddenByRow },
          },
        ])
        return prev.map((r) => {
        if (r.id !== id) return r

        const nextImages = (r.images ?? []).slice(0, maxImages)
        const nextCutNos = (r.cutNos ?? []).slice(0, maxImages)

        // 最大数までのスロットを確保（既存は 8 スロット前提）
        while (nextImages.length < maxImages) nextImages.push(null)
        while (nextCutNos.length < maxImages) nextCutNos.push('')

        const freeIndices = []
        for (let i = 0; i < maxImages; i++) {
          if (!nextImages[i]) freeIndices.push(i)
        }

        const additions = imageStates.slice(0, freeIndices.length)
        additions.forEach((imgState, j) => {
          const targetIdx = freeIndices[j]
          nextImages[targetIdx] = imgState
          nextCutNos[targetIdx] = ''
        })

        return { ...r, images: nextImages, cutNos: nextCutNos }
      })
      }
    )
  }, [callSheetDurationHiddenByRow])

  const handleCallSheetImageCropUpdate = useCallback((id, imageIndex, patch) => {
    setCallSheetRows((prev) =>
      {
        setCallSheetUndoStack((stack) => [
          ...stack.slice(-49),
          {
            rows: cloneCallSheetRowsForUndo(prev),
            durationHiddenByRow: { ...callSheetDurationHiddenByRow },
          },
        ])
        return prev.map((r) => {
        if (r.id !== id) return r
        const nextImages = (r.images ?? [null, null, null, null, null, null, null, null]).slice()
        const cur = nextImages[imageIndex]
        if (!cur) return r
        nextImages[imageIndex] = { ...cur, ...patch }
        return { ...r, images: nextImages }
      })
      }
    )
  }, [callSheetDurationHiddenByRow])

  const handleCallSheetDropFromStoryboard = useCallback(async (callSheetRowId, payload) => {
    const { image, cutNumber } = payload
    const cloned = await cloneImageStateForCallSheet(image)
    if (!cloned) return
    setCallSheetRows((prev) =>
      {
        setCallSheetUndoStack((stack) => [
          ...stack.slice(-49),
          {
            rows: cloneCallSheetRowsForUndo(prev),
            durationHiddenByRow: { ...callSheetDurationHiddenByRow },
          },
        ])
        return prev.map((r) => {
        if (r.id !== callSheetRowId) return r
        const nextImages = [...(r.images ?? [])]
        const nextCutNos = [...(r.cutNos ?? [])]
        while (nextImages.length < 8) nextImages.push(null)
        while (nextCutNos.length < 8) nextCutNos.push('')
        const idx = nextImages.findIndex((x) => !x)
        if (idx === -1) return r
        nextImages[idx] = cloned
        nextCutNos[idx] = String(cutNumber)
        return { ...r, images: nextImages, cutNos: nextCutNos }
      })
      }
    )
  }, [callSheetDurationHiddenByRow])

  const goToCallSheetFromStoryboard = useCallback(() => {
    setCallsheetStoryboardSourceItems(buildCallsheetStoryboardSourceItems(rows))
    setPage('callsheet')
  }, [rows])

  const handleAddCallSheetRow = useCallback(() => {
    setCallSheetRows((prev) => {
      setCallSheetUndoStack((stack) => [
        ...stack.slice(-49),
        {
          rows: cloneCallSheetRowsForUndo(prev),
          durationHiddenByRow: { ...callSheetDurationHiddenByRow },
        },
      ])
      return [...prev, emptyCallSheetRow()]
    })
  }, [callSheetDurationHiddenByRow])

  const handleDeleteCallSheetRow = useCallback((id) => {
    setCallSheetRows((prev) => {
      setCallSheetUndoStack((stack) => [
        ...stack.slice(-49),
        {
          rows: cloneCallSheetRowsForUndo(prev),
          durationHiddenByRow: { ...callSheetDurationHiddenByRow },
        },
      ])
      const target = prev.find((r) => r.id === id)
      ;(target?.images ?? []).forEach((img) => {
        if (img?.objectUrl) URL.revokeObjectURL(img.objectUrl)
      })
      return prev.filter((r) => r.id !== id)
    })
  }, [callSheetDurationHiddenByRow])

  const canExportCallSheetExcel = useMemo(
    () =>
      callSheetRows.some((r) => {
        const hasImage = (r.images ?? []).some(Boolean)
        return (
          hasImage ||
          String(r.time ?? '').trim() !== '' ||
          String(r.schedule ?? '').trim() !== '' ||
          String(r.note ?? '').trim() !== ''
        )
      }),
    [callSheetRows]
  )

  const handleAddRow = useCallback(() => {
    setRows((prev) => [...prev, emptyRow()])
  }, [])

  const handleDeleteRow = useCallback((id) => {
    setRows((prev) => {
      const idx = prev.findIndex((r) => r.id === id)
      if (idx === -1) return prev
      const removed = prev[idx]
      const snapshot = {
        row: {
          ...removed,
          image: removed.image ? { ...removed.image } : null,
        },
        index: idx,
      }
      setDeleteUndoStack((stack) => [...stack, snapshot])
      return prev.filter((r) => r.id !== id)
    })
  }, [])

  const handleUndoDelete = useCallback(() => {
    setDeleteUndoStack((stack) => {
      if (stack.length === 0) return stack
      const snap = stack[stack.length - 1]
      const rest = stack.slice(0, -1)
      setRows((prev) => {
        const { row, index } = snap
        const next = [...prev]
        const safeIndex = Math.min(Math.max(0, index), next.length)
        next.splice(safeIndex, 0, row)
        return next
      })
      return rest
    })
  }, [])

  const handleExportExcel = useCallback(async () => {
    if (!storyboardHasExportableRow(rows)) return

    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('絵コンテ', { views: [{ state: 'frozen', ySplit: 1 }] })
    const IMAGE_WIDTH = EXCEL_IMAGE_DISPLAY_WIDTH_PX
    const IMAGE_HEIGHT = EXCEL_IMAGE_DISPLAY_HEIGHT_PX
    const EXCEL_DATA_ROW_HEIGHT = EXCEL_IMAGE_ROW_HEIGHT_POINTS
    /** ヘッダー行のみ：既定より少し高く（pt） */
    const EXCEL_HEADER_ROW_HEIGHT = 22
    const EXCEL_FONT_SIZE = 12
    const cellAlignCenter = {
      horizontal: 'center',
      vertical: 'middle',
    }

    const excelWidths = computeExcelColumnWidths(selectedColumnIds)
    const cols = [
      {
        header: FIXED_CUT_NUMBER_COLUMN.label,
        key: 'c0',
        width: excelColumnWidthFromRatio(CUT_COLUMN_EXCEL_RATIO),
      },
      ...selectedColumnIds.map((id, i) => ({
        header: getColumnMeta(id)?.label ?? id,
        key: `c${i + 1}`,
        width: excelWidths[i],
      })),
    ]
    sheet.columns = cols

    const headerRow = sheet.getRow(1)
    headerRow.height = EXCEL_HEADER_ROW_HEIGHT
    headerRow.font = { size: EXCEL_FONT_SIZE, bold: false }
    headerRow.eachCell((cell) => {
      cell.font = { size: EXCEL_FONT_SIZE, bold: false }
      cell.alignment = {
        ...cellAlignCenter,
        wrapText: true,
      }
    })

    const imageColIndexInSelection = selectedColumnIds.indexOf('image')
    const excelImageCol =
      imageColIndexInSelection >= 0 ? imageColIndexInSelection + 1 : -1

    rows.forEach((row, index) => {
      const excelRowIndex = index + 2
      const data = { c0: index + 1 }
      selectedColumnIds.forEach((id, i) => {
        const key = `c${i + 1}`
        if (id === 'image') data[key] = ''
        else data[key] = getRowFieldValue(row, id, { rows, rowIndex: index })
      })
      sheet.addRow(data)
      if (excelImageCol >= 0 && row.image) {
        sheet.getRow(excelRowIndex).height = EXCEL_DATA_ROW_HEIGHT
      }
    })

    /* 全セル: フォント12・上下左右中央（テキスト列は折り返し）。画像列の幅・高さロジックは変更しない */
    rows.forEach((row, index) => {
      const excelRowIndex = index + 2
      const sheetRow = sheet.getRow(excelRowIndex)
      sheetRow.font = { size: EXCEL_FONT_SIZE }
      sheetRow.getCell(1).alignment = cellAlignCenter
      selectedColumnIds.forEach((id, i) => {
        const colIdx = i + 2
        if (id === 'image') {
          sheetRow.getCell(colIdx).alignment = { ...cellAlignCenter }
        } else {
          sheetRow.getCell(colIdx).alignment = {
            ...cellAlignCenter,
            wrapText: true,
          }
        }
      })
      if (!row.image) {
        sheetRow.height = undefined
      }
    })

    if (excelImageCol >= 0) {
      for (let index = 0; index < rows.length; index++) {
        const row = rows[index]
        if (!row.image) continue
        const im = row.image
        let objectUrl = im.objectUrl
        let revokeTemp = false
        if (!objectUrl && im.file) {
          objectUrl = URL.createObjectURL(im.file)
          revokeTemp = true
        }
        if (!objectUrl) continue

        let base64
        let extension = 'png'
        try {
          const out = await cropStateToExportImage(
            {
              objectUrl,
              naturalWidth: im.naturalWidth,
              naturalHeight: im.naturalHeight,
              containerWidth: im.containerWidth ?? DISPLAY_CROP_WIDTH,
              containerHeight: im.containerHeight ?? DISPLAY_CROP_HEIGHT,
              scale: im.scale,
              initialScale: im.initialScale,
              offsetX: im.offsetX,
              offsetY: im.offsetY,
              crop: im.crop,
            },
            {
              format: 'png',
              logicalWidth: EXCEL_IMAGE_DISPLAY_WIDTH_PX,
              logicalHeight: EXCEL_IMAGE_DISPLAY_HEIGHT_PX,
            }
          )
          base64 = out.base64
          extension = out.extension
        } catch (e) {
          console.error('Excel image export failed:', e)
          continue
        } finally {
          if (revokeTemp) URL.revokeObjectURL(objectUrl)
        }
        const imageId = workbook.addImage({
          base64,
          extension,
        })
        const r = index + 1
        sheet.addImage(imageId, {
          tl: { col: excelImageCol, row: r },
          ext: { width: IMAGE_WIDTH, height: IMAGE_HEIGHT },
          editAs: 'oneCell',
        })
      }
    }

    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `絵コンテ_${new Date().toISOString().slice(0, 10)}.xlsx`
    a.click()
    URL.revokeObjectURL(url)
  }, [rows, selectedColumnIds])

  const handleExportCallSheetExcel = useCallback(async () => {
    if (!canExportCallSheetExcel) return

    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('香盤表', { views: [{ state: 'frozen', ySplit: 1 }] })

    const imgW = 240
    const imgH = 135
    const gap = 5

    const colIndex = 1
    const TEXT_ROW_HEIGHT_PTS = 40
    const CUT_ROW_HEIGHT_PTS = 20

    sheet.columns = [
      { header: 'TIME', key: 'time' },
      { header: 'スケジュール', key: 'schedule', width: 110 },
      { header: '備考', key: 'note', width: 20 },
    ]

    const header = sheet.getRow(1)
    header.height = 22
    header.eachCell((cell) => {
      cell.font = { size: 12, bold: false }
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E8F6' } }
      cell.alignment = { vertical: 'middle', horizontal: 'center' }
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF7F7F7F' } },
        left: { style: 'thin', color: { argb: 'FF7F7F7F' } },
        bottom: { style: 'thin', color: { argb: 'FF7F7F7F' } },
        right: { style: 'thin', color: { argb: 'FF7F7F7F' } },
      }
    })

    const thinBorder = { style: 'thin', color: { argb: 'FF7F7F7F' } }
    const rowLayout = []

    let currentRow = 2
    callSheetRows.forEach((row, rowIdx) => {
      const baseRow = currentRow
      const titleRow = baseRow
      const cutRow = baseRow + 1
      const imageRow = baseRow + 2

      const imagesWithSlotIdx = (row.images ?? [])
        .map((img, slotIdx) => (img ? { img, slotIdx } : null))
        .filter(Boolean)

      const imageCount = imagesWithSlotIdx.length
      const durationLabel = durationLabelFromTimesForCallSheet(row.time ?? '', callSheetRows[rowIdx + 1]?.time ?? '')
      const scheduleBase = String(row.schedule ?? '')
      const titleText = durationLabel
        ? `${scheduleBase}${scheduleBase ? ' ' : ''}${durationLabel}`
        : scheduleBase
      const cutText = (row.cutNos ?? [])
        .map((v) => String(v ?? '').trim())
        .filter((v) => v !== '')
        .map((v) => `Cut#${v.replace(/^cut#?/i, '')}`)
        .join(', ')

      sheet.addRow({ time: '', schedule: '', note: '' })
      sheet.addRow({ time: '', schedule: '', note: '' })
      sheet.addRow({ time: '', schedule: '', note: '' })

      currentRow += 3

      rowLayout.push({
        titleRow,
        cutRow,
        imageRow,
        imageCount,
        timeText: row.time ?? '',
        noteText: row.note ?? '',
        titleText,
        cutText,
      })
    })

    rowLayout.forEach(
      ({
        titleRow,
        cutRow,
        imageRow,
        imageCount,
        timeText,
        noteText,
        titleText,
        cutText,
      }) => {
        sheet.mergeCells(`A${titleRow}:A${imageRow}`)
        sheet.mergeCells(`C${titleRow}:C${imageRow}`)

        sheet.getCell(titleRow, 1).value = timeText
        sheet.getCell(titleRow, 3).value = noteText
        sheet.getCell(titleRow, 2).value = titleText
        sheet.getCell(cutRow, 2).value = cutText

        const topRow = sheet.getRow(titleRow)
        topRow.height = TEXT_ROW_HEIGHT_PTS
        const cutRowSheet = sheet.getRow(cutRow)
        cutRowSheet.height = CUT_ROW_HEIGHT_PTS
        const bottomRow = sheet.getRow(imageRow)
        bottomRow.height =
          imageCount > 0 ? imgH * 0.75 : 1

        const aMaster = sheet.getCell(titleRow, 1)
        aMaster.font = { size: 11 }
        aMaster.alignment = {
          vertical: 'middle',
          horizontal: 'center',
          wrapText: true,
        }
        aMaster.border = {
          top: thinBorder,
          left: thinBorder,
          bottom: thinBorder,
          right: thinBorder,
        }

        const bTitle = sheet.getCell(titleRow, 2)
        bTitle.font = { size: 11 }
        bTitle.alignment = {
          vertical: 'top',
          horizontal: 'left',
          wrapText: true,
        }
        bTitle.border = {
          top: thinBorder,
          left: thinBorder,
          bottom: { style: 'none' },
          right: thinBorder,
        }

        const bCut = sheet.getCell(cutRow, 2)
        bCut.font = { size: 11 }
        bCut.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        }
        bCut.border = {
          top: { style: 'none' },
          left: thinBorder,
          bottom: { style: 'none' },
          right: thinBorder,
        }

        const bImg = sheet.getCell(imageRow, 2)
        bImg.font = { size: 11 }
        bImg.alignment = {
          vertical: 'top',
          horizontal: 'left',
          wrapText: true,
        }
        bImg.border = {
          top: { style: 'none' },
          left: thinBorder,
          bottom: thinBorder,
          right: thinBorder,
        }

        const cMaster = sheet.getCell(titleRow, 3)
        cMaster.font = { size: 11 }
        cMaster.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: true,
        }
        cMaster.border = {
          top: thinBorder,
          left: thinBorder,
          bottom: thinBorder,
          right: thinBorder,
        }
      }
    )

    for (let i = 0; i < callSheetRows.length; i++) {
      const row = callSheetRows[i]
      const { imageRow } = rowLayout[i]

      const imagesWithSlotIdx = (row.images ?? [])
        .map((img, slotIdx) => (img ? { img, slotIdx } : null))
        .filter(Boolean)

      for (let index = 0; index < imagesWithSlotIdx.length; index++) {
        const { img: im } = imagesWithSlotIdx[index]

        let objectUrl = im.objectUrl
        let revokeTemp = false
        if (!objectUrl && im.file) {
          objectUrl = URL.createObjectURL(im.file)
          revokeTemp = true
        }
        if (!objectUrl) continue

        let base64
        let extension = 'png'
        try {
          const out = await cropStateToExportImage(
            {
              objectUrl,
              naturalWidth: im.naturalWidth,
              naturalHeight: im.naturalHeight,
              containerWidth: im.containerWidth ?? DISPLAY_CROP_WIDTH,
              containerHeight: im.containerHeight ?? DISPLAY_CROP_HEIGHT,
              scale: im.scale,
              initialScale: im.initialScale,
              offsetX: im.offsetX,
              offsetY: im.offsetY,
              crop: im.crop,
            },
            {
              format: 'png',
              logicalWidth: imgW,
              logicalHeight: imgH,
            }
          )
          base64 = out.base64
          extension = out.extension
        } catch (e) {
          console.error('Call sheet image export failed:', e)
          continue
        } finally {
          if (revokeTemp) URL.revokeObjectURL(objectUrl)
        }

        const imageId = workbook.addImage({ base64, extension })
        sheet.addImage(imageId, {
          tl: {
            col: colIndex,
            row: imageRow - 1,
          },
          ext: { width: imgW, height: imgH },
        })

        const last = sheet.getImages().slice(-1)[0]
        if (last?.range?.tl) {
          last.range.tl.nativeColOff = index * (imgW + gap) * EXCEL_EMU_PER_PX
        }
      }
    }

    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `香盤表_${new Date().toISOString().slice(0, 10)}.xlsx`
    a.click()
    URL.revokeObjectURL(url)
  }, [callSheetRows, canExportCallSheetExcel])

  const canExportExcel = useMemo(() => storyboardHasExportableRow(rows), [rows])

  const canGenerateNaScript = useMemo(() => {
    const hasNarrationColumn = selectedColumnIds.includes('narration')
    const hasNarrationText = rows.some((scene) => (scene.narration ?? '').trim() !== '')
    return hasNarrationColumn && hasNarrationText
  }, [rows, selectedColumnIds])

  const handleGenerateNaScript = useCallback(() => {
    const naData = rows
      .filter((scene) => (scene.narration ?? '').trim() !== '')
      .map((scene) => ({
        text: scene.narration,
        duration: Number(scene.duration) || 0,
        timecode: formatTime(Number(scene.duration) || 0),
      }))
    localStorage.setItem('naData', JSON.stringify(naData))
    setExcelNaRows(naData)
    setNaDataTick((n) => n + 1)
    setNaMode('excel')
    setPage('na')
  }, [rows])

  const handleDownloadNaPpt = useCallback(async (naData) => {
    if (!Array.isArray(naData) || naData.length === 0) return

    const pptx = new PptxGenJS()
    pptx.layout = 'LAYOUT_WIDE'
    const slideWidth = 13.333
    const marginX = 0.25
    const topY = 0.25
    const timeRowHeight = 0.9
    const bodyHeight = 5.9
    const bottomFooterHeight = 0.35
    const lineColor = '5B8AA6'
    const slotsPerPage = 5
    const pages = chunkIntoPages(naData, slotsPerPage)

    pages.forEach((pageItems, pageIdx) => {
      const slide = pptx.addSlide()
      const areaWidth = slideWidth - marginX * 2
      const colWidth = areaWidth / slotsPerPage
      const gridBottomY = topY + timeRowHeight + bodyHeight

      // 区切り線（上の水平線 + 各列の縦線）で読みやすくする
      slide.addShape(pptx.ShapeType.line, {
        x: marginX,
        y: topY + timeRowHeight,
        w: areaWidth,
        h: 0,
        line: { color: lineColor, pt: 1.2 },
      })

      // 縦線は列間のみ（左右端は描かない）
      for (let i = 1; i < slotsPerPage; i++) {
        const x = marginX + colWidth * i
        slide.addShape(pptx.ShapeType.line, {
          x,
          y: topY,
          w: 0,
          h: timeRowHeight + bodyHeight + bottomFooterHeight,
          line: { color: lineColor, pt: 1.2 },
        })
      }

      for (let i = 0; i < slotsPerPage; i++) {
        const colX = marginX + i * colWidth
        // 右端から順にデータを入れる（未入力スロットは左側に残る）
        const sourceIndex = slotsPerPage - 1 - i
        const item = pageItems[sourceIndex]

        // 下部の列番号（右から左で 1,2,3...）
        slide.addText(String(pageIdx * slotsPerPage + (slotsPerPage - i)), {
          x: colX,
          y: gridBottomY + 0.02,
          w: colWidth,
          h: 0.28,
          align: 'right',
          valign: 'mid',
          fontSize: 11,
          bold: true,
          color: '222222',
        })

        // 未入力のブロックは空のまま（スペースだけ確保）
        if (!item) continue

        // タイムコード（横書き）— 枠は図形ではなくテキストボックスの線
        slide.addText(item.timecode || formatTime(item.duration || 0), {
          x: colX + colWidth * 0.22,
          y: topY + 0.24,
          w: colWidth * 0.56,
          h: 0.3,
          align: 'center',
          valign: 'mid',
          fontSize: 11,
          isTextBox: true,
          line: { color: '333333', width: 1 },
        })

        // ナレーション本文（縦書き）
        slide.addText(item.text || '', {
          x: colX + colWidth * 0.2,
          y: topY + timeRowHeight + 0.16,
          w: colWidth * 0.6,
          h: bodyHeight - 0.17,
          vert: 'eaVert',
          align: 'left',
          valign: 'top',
          margin: [0, 0, 0, 0],
          fontSize: 17,
          bold: true,
        })
      }
    })

    await pptx.writeFile({ fileName: 'na_script.pptx' })
  }, [])

  const handleImportNaPpt = useCallback(async (file) => {
    if (!file) return
    try {
      const arrayBuffer = await file.arrayBuffer()
      const zip = await JSZip.loadAsync(arrayBuffer)
      const slideFiles = Object.keys(zip.files)
        .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
        .sort((a, b) => {
          const an = Number((a.match(/slide(\d+)\.xml/i) || [])[1] || 0)
          const bn = Number((b.match(/slide(\d+)\.xml/i) || [])[1] || 0)
          return an - bn
        })

      const parsedRows = []
      for (const slidePath of slideFiles) {
        const xml = await zip.file(slidePath)?.async('text')
        if (!xml) continue
        const slots = parseNaPptxSlideXml(xml)
        slots.forEach(({ index, row }) => {
          if (parsedRows.length <= index) parsedRows.length = index + 1
          parsedRows[index] = row
        })
      }

      localStorage.setItem('naData', JSON.stringify(parsedRows))
      setExcelNaRows(parsedRows)
      setNaDataTick((n) => n + 1)
      setNaMode('excel')
      setPage('na')
    } catch (e) {
      console.error('NA PPT import failed:', e)
      alert(e?.message ?? 'PowerPointの読み込みに失敗しました')
    }
  }, [])

  const pushManualUndo = useCallback((prevRows) => {
    setManualNaUndoStack((stack) => [...stack, prevRows])
  }, [])

  const handleManualUndo = useCallback(() => {
    setManualNaUndoStack((stack) => {
      if (stack.length === 0) return stack
      const prevRows = stack[stack.length - 1]
      setManualNaRows(prevRows)
      return stack.slice(0, -1)
    })
  }, [])

  const updateExcelNaRows = useCallback((updater) => {
    setExcelNaRows((prev) => {
      const base = prev.length > 0 ? prev : readNaDataFromStorage()
      const next = typeof updater === 'function' ? updater(base) : updater
      localStorage.setItem('naData', JSON.stringify(next))
      return next
    })
    setNaDataTick((n) => n + 1)
  }, [])

  const handleClearManualNaBlock = useCallback(
    (globalIdx) => {
      setManualNaRows((prev) => {
        pushManualUndo(prev)
        const next = prev.slice()
        next[globalIdx] = null
        return next
      })
    },
    [pushManualUndo]
  )

  const handleClearExcelNaBlock = useCallback(
    (globalIdx) => {
      updateExcelNaRows((prev) => {
        const next = prev.slice()
        next[globalIdx] = null
        return next
      })
    },
    [updateExcelNaRows]
  )

  const handleActivateManualNaSlot = useCallback(
    (globalIdx) => {
      setManualNaRows((prev) => {
        pushManualUndo(prev)
        const next = prev.slice()
        if (next.length <= globalIdx) next.length = globalIdx + 1
        next[globalIdx] = { timecode: formatTime(0), text: '' }
        return next
      })
    },
    [pushManualUndo]
  )

  const handleActivateExcelNaSlot = useCallback(
    (globalIdx) => {
      updateExcelNaRows((prev) => {
        const next = prev.slice()
        if (next.length <= globalIdx) next.length = globalIdx + 1
        next[globalIdx] = { timecode: formatTime(0), text: '' }
        return next
      })
    },
    [updateExcelNaRows]
  )

  const handleAddManualNaPage = useCallback(() => {
    setManualNaRows((prev) => {
      pushManualUndo(prev)
      return addNaPageSlots(prev)
    })
  }, [pushManualUndo])

  const handleAddExcelNaPage = useCallback(() => {
    updateExcelNaRows((prev) => addNaPageSlots(prev))
  }, [updateExcelNaRows])

  if (page === 'home') {
    return (
      <div className="app app--home">
        <div className="home-card">
          <h1 className="home-title">映像制作Toolbox</h1>
          <div className="home-actions">
            <button type="button" className="btn btn-secondary btn-home-action" onClick={() => setPage('storyboard')}>
              絵コンテメーカー
            </button>
            <button
              type="button"
              className="btn btn-secondary btn-home-action"
              onClick={() => {
                setCallsheetStoryboardSourceItems(null)
                setPage('callsheet')
              }}
            >
              香盤表メーカー
            </button>
            <button
              type="button"
              className="btn btn-secondary btn-home-action"
              onClick={() => {
                setNaMode('new')
                setPage('na')
              }}
            >
              NA原稿メーカー
            </button>
          </div>
        </div>
      </div>
    )
  }

  if (page === 'callsheet') {
    return (
      <ScheduleComponent
        rows={callSheetRows}
        onUpdateRow={handleCallSheetUpdate}
        onDeleteRow={handleDeleteCallSheetRow}
        onAddRow={handleAddCallSheetRow}
        onImportExcel={handleImportCallSheetExcel}
        onExportExcel={handleExportCallSheetExcel}
        canExportExcel={canExportCallSheetExcel}
        onBackHome={() => {
          setCallsheetStoryboardSourceItems(null)
          setPage('home')
        }}
        onAddImages={handleCallSheetAddImages}
        storyboardSourceItems={callsheetStoryboardSourceItems}
        onDropStoryboardCut={handleCallSheetDropFromStoryboard}
        onReorderRows={handleCallSheetReorderRows}
        onUndo={handleUndoCallSheet}
        canUndo={callSheetUndoStack.length > 0}
        durationHiddenByRow={callSheetDurationHiddenByRow}
        onDurationHiddenChange={handleCallSheetDurationHiddenChange}
      />
    )
  }

  if (page === 'na') {
    void naDataTick
    const naDataFromStorage = readNaDataFromStorage()
    const effectiveExcelRows = excelNaRows.length > 0 ? excelNaRows : naDataFromStorage
    const sourceRows = naMode === 'new' ? manualNaRows : effectiveExcelRows
    const naData = sourceRows
      .filter((row) => row != null && (row.text ?? '').trim() !== '')
      .map((row) => ({ text: row.text, duration: 0, timecode: row.timecode || formatTime(0) }))
    const manualPageCount = getNaPageCount(manualNaRows)
    const excelPageCount = getNaPageCount(effectiveExcelRows)
    return (
      <div className="app na-page">
        <header className="header">
          <h1 className="app-heading">NA原稿メーカー</h1>
          <div className="toolbar">
            <button type="button" className="btn btn-secondary" onClick={() => setPage('home')}>
              ホーム
            </button>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={handleManualUndo}
              disabled={naMode !== 'new' || manualNaUndoStack.length === 0}
            >
              戻す
            </button>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => setPage('storyboard')}
            >
              絵コンテから作成
            </button>
            <button
              type="button"
              className="btn btn-secondary"
              onClick={() => naPptImportInputRef.current?.click()}
            >
              保存されたPowerPointで継続
            </button>
            <input
              ref={naPptImportInputRef}
              type="file"
              accept=".pptx"
              style={{ display: 'none' }}
              onChange={(e) => {
                const file = e.target.files?.[0]
                if (file) handleImportNaPpt(file)
                e.target.value = ''
              }}
            />
            <button
              type="button"
              className="btn btn-powerpoint"
              onClick={() => handleDownloadNaPpt(naData)}
              disabled={naData.length === 0}
            >
              PowerPointで保存
            </button>
          </div>
        </header>

        {naMode === 'new' ? (
          <div className="na-editor">
            <div className="na-pages">
              {Array.from({ length: manualPageCount }).map((_, pageIdx) => (
                <section className="na-sheet" key={`manual-page-${pageIdx}`}>
                  <div className="na-sheet__columns na-columns na-columns--rtl">
                    {Array.from({ length: 5 }).map((_, slotIdx) => {
                      const globalIdx = pageIdx * 5 + slotIdx
                      const row = manualNaRows[globalIdx]
                      const blockNo = pageIdx * 5 + (slotIdx + 1)
                      if (row == null) {
                        return (
                          <button
                            type="button"
                            key={`manual-empty-${pageIdx}-${slotIdx}`}
                            className="na-column na-column--empty na-column--add-slot"
                            onClick={() => handleActivateManualNaSlot(globalIdx)}
                            aria-label={`列${blockNo}を追加`}
                          >
                            <div className="na-column-footer na-column-footer--empty">
                              <div className="na-index">{blockNo}</div>
                            </div>
                          </button>
                        )
                      }
                      return (
                        <div className="na-column" key={`manual-${globalIdx}`}>
                          <input
                            className="na-time-input"
                            value={row.timecode}
                            onChange={(e) =>
                              setManualNaRows((prev) => {
                                pushManualUndo(prev)
                                const next = prev.slice()
                                const cur = next[globalIdx]
                                if (!cur) return prev
                                next[globalIdx] = { ...cur, timecode: e.target.value }
                                return next
                              })
                            }
                            placeholder="00:00"
                          />
                          <textarea
                            className="na-vertical na-vertical-input"
                            value={row.text}
                            onChange={(e) =>
                              setManualNaRows((prev) => {
                                pushManualUndo(prev)
                                const next = prev.slice()
                                const cur = next[globalIdx]
                                if (!cur) return prev
                                next[globalIdx] = { ...cur, text: e.target.value }
                                return next
                              })
                            }
                          />
                          <div className="na-column-footer">
                            <button
                              type="button"
                              className="row-delete-btn"
                              onClick={() => handleClearManualNaBlock(globalIdx)}
                              aria-label={`列${blockNo}を空欄に戻す`}
                              title="この列を空欄に戻す"
                            >
                              ×
                            </button>
                            <div className="na-index">{blockNo}</div>
                          </div>
                        </div>
                      )
                    })}
                  </div>
                </section>
              ))}
              <button
                type="button"
                className="grid-add-row-zone na-add-page-bar"
                onClick={handleAddManualNaPage}
                aria-label="ページを追加"
              >
                <span className="grid-add-row-zone__content">
                  <span className="grid-add-row-zone__icon" aria-hidden>
                    +
                  </span>
                  <span className="grid-add-row-zone__label">ページを追加</span>
                </span>
              </button>
            </div>
          </div>
        ) : effectiveExcelRows.length === 0 ? (
          <div className="na-empty">
            NAデータがありません。「絵コンテから作成」で絵コンテメーカーを開き、「NA原稿を生成」してください。
          </div>
        ) : (
          <div className="na-editor">
            <div className="na-pages">
              {Array.from({ length: excelPageCount }).map((_, pageIdx) => (
                <section className="na-sheet" key={`excel-page-${pageIdx}`}>
                  <div className="na-sheet__columns na-columns na-columns--rtl">
                    {Array.from({ length: 5 }).map((_, slotIdx) => {
                      const globalIdx = pageIdx * 5 + slotIdx
                      const row = effectiveExcelRows[globalIdx]
                      const blockNo = pageIdx * 5 + (slotIdx + 1)
                      if (row == null) {
                        return (
                          <button
                            type="button"
                            key={`excel-empty-${pageIdx}-${slotIdx}`}
                            className="na-column na-column--empty na-column--add-slot"
                            onClick={() => handleActivateExcelNaSlot(globalIdx)}
                            aria-label={`列${blockNo}を追加`}
                          >
                            <div className="na-column-footer na-column-footer--empty">
                              <div className="na-index">{blockNo}</div>
                            </div>
                          </button>
                        )
                      }
                      return (
                        <div className="na-column" key={`excel-${globalIdx}`}>
                          <input
                            className="na-time-input"
                            value={row.timecode}
                            onChange={(e) =>
                              updateExcelNaRows((prev) => {
                                const next = prev.slice()
                                const cur = next[globalIdx]
                                if (!cur) return prev
                                next[globalIdx] = { ...cur, timecode: e.target.value }
                                return next
                              })
                            }
                            placeholder="00:00"
                          />
                          <textarea
                            className="na-vertical na-vertical-input"
                            value={row.text}
                            onChange={(e) =>
                              updateExcelNaRows((prev) => {
                                const next = prev.slice()
                                const cur = next[globalIdx]
                                if (!cur) return prev
                                next[globalIdx] = { ...cur, text: e.target.value }
                                return next
                              })
                            }
                          />
                          <div className="na-column-footer">
                            <button
                              type="button"
                              className="row-delete-btn"
                              onClick={() => handleClearExcelNaBlock(globalIdx)}
                              aria-label={`列${blockNo}を空欄に戻す`}
                              title="この列を空欄に戻す"
                            >
                              ×
                            </button>
                            <div className="na-index">{blockNo}</div>
                          </div>
                        </div>
                      )
                    })}
                  </div>
                </section>
              ))}
              <button
                type="button"
                className="grid-add-row-zone na-add-page-bar"
                onClick={handleAddExcelNaPage}
                aria-label="ページを追加"
              >
                <span className="grid-add-row-zone__content">
                  <span className="grid-add-row-zone__icon" aria-hidden>
                    +
                  </span>
                  <span className="grid-add-row-zone__label">ページを追加</span>
                </span>
              </button>
            </div>
          </div>
        )}

        <section className="editor-tips" aria-label="注意事項">
          <div className="editor-tips__title">💡 Tips:</div>
          <ul className="editor-tips__list">
            <li>こちらのウェブサイトで入力内容は保存されません。</li>
            <li>ページを再読み込みするとすべての内容がクリアされます。</li>
            <li>PowerPointで保存して、またアップロードすることで継続して作業できます。</li>
          </ul>
        </section>
      </div>
    )
  }

  if (screen === 'picker') {
    return (
      <div className="app app--picker">
        <header className="header">
          <AppHeading />
          <div className="toolbar">
            <button type="button" className="btn btn-secondary" onClick={() => setPage('home')}>
              ホーム
            </button>
          </div>
        </header>
        <ColumnPicker
          selectedColumnIds={selectedColumnIds}
          onAddColumn={handleAddColumn}
          onRemoveColumn={handleRemoveColumn}
          onReorderSelectedColumns={handleReorderSelectedColumns}
          onImportExcel={handleImportExcel}
          onConfirm={handleConfirmColumns}
        />
      </div>
    )
  }

  return (
    <div className="app">
      <header className="header">
        <AppHeading />
        <div className="toolbar">
          <button
            type="button"
            className="btn btn-export"
            onClick={handleExportExcel}
            disabled={!canExportExcel}
          >
            Excelで保存
          </button>
          <button type="button" className="btn btn-secondary" onClick={() => setPage('home')}>
            ホーム
          </button>
          <button type="button" className="btn btn-secondary" onClick={() => setScreen('picker')}>
            列の設定
          </button>
          <button
            type="button"
            className="btn btn-secondary"
            onClick={handleUndoDelete}
            disabled={deleteUndoStack.length === 0}
            title="直前に削除した行を元の位置に戻します"
          >
            戻す
          </button>
          <button
            type="button"
            className="btn btn-callsheet-create"
            onClick={goToCallSheetFromStoryboard}
            disabled={!rows.some((r) => r.image)}
            title="香盤表メーカーへ。絵コンテの画像をドラッグして割り当てられます。"
          >
            香盤表を作成
          </button>
          <button
            type="button"
            className="btn btn-add"
            onClick={handleGenerateNaScript}
            disabled={!canGenerateNaScript}
          >
            NA原稿を生成
          </button>
        </div>
      </header>

      <div className="grid-container">
        <div className="grid-header" style={{ gridTemplateColumns: gridTemplate }}>
          <div className="cell cell-drag-only cell-drag-header" />
          <div className="cell cell-cut-number cell-cut-number-header">
            {FIXED_CUT_NUMBER_COLUMN.label}
          </div>
          {selectedColumnIds.map((id) => (
            <div
              key={id}
              className={`cell grid-header-column ${id === 'image' ? 'cell-image-header' : ''}`}
            >
              {getColumnMeta(id)?.label ?? id}
            </div>
          ))}
        </div>

        <DndContext
          sensors={sensors}
          collisionDetection={closestCenter}
          onDragEnd={handleDragEnd}
        >
          <SortableContext
            items={rows.map((r) => r.id)}
            strategy={verticalListSortingStrategy}
          >
            <div className="grid-body">
              {rows.map((row, rowIndex) => (
                <SortableRow
                  key={row.id}
                  row={row}
                  rowIndex={rowIndex}
                  rows={rows}
                  firstEmptyImageRowIndex={firstEmptyImageRowIndex}
                  displayIndex={rowIndex + 1}
                  selectedColumnIds={selectedColumnIds}
                  gridTemplateColumns={gridTemplate}
                  sceneOptions={sceneHistoryOptions}
                  locationOptions={locationHistoryOptions}
                  modelOptions={modelHistoryOptions}
                  costumeOptions={costumeHistoryOptions}
                  onUpdate={handleUpdate}
                  onDelete={handleDeleteRow}
                  onImageDrop={handleImageDrop}
                  onImageCropUpdate={handleImageCropUpdate}
                />
              ))}
            </div>
          </SortableContext>
        </DndContext>

        <button
          type="button"
          className="grid-add-row-zone"
          onClick={handleAddRow}
          aria-label="行を追加"
        >
          <span className="grid-add-row-zone__content">
            <span className="grid-add-row-zone__icon" aria-hidden>
              +
            </span>
            <span className="grid-add-row-zone__label">行を追加</span>
          </span>
        </button>
      </div>

      <section className="editor-tips" aria-label="注意事項">
        <div className="editor-tips__title">💡 Tips:</div>
        <ul className="editor-tips__list">
          <li>こちらのウェブサイトで入力内容は保存されません。</li>
          <li>ページを再読み込みするとすべての内容がクリアされます。</li>
          <li>Excelで保存して、またアップロードすることで継続して作業できます。</li>
        </ul>
      </section>
    </div>
  )
}
