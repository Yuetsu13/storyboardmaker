import { useState, useCallback, useMemo } from 'react'
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

function AppHeading() {
  return (
    <h1 className="app-heading">
      絵コンテメーカー
      <span className="app-heading-aspect">16:9</span>
    </h1>
  )
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
    elapsedTime: '',
    /** true のときは elapsedTime をそのまま表示（尺からの自動計算はしない） */
    elapsedTimeManual: false,
    note: '',
  }
}

const initialRows = [emptyRow(), emptyRow()]

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

export default function App() {
  const [screen, setScreen] = useState('picker')
  const [selectedColumnIds, setSelectedColumnIds] = useState([])
  const [rows, setRows] = useState(initialRows)
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

  if (screen === 'picker') {
    return (
      <div className="app app--picker">
        <header className="header">
          <AppHeading />
          {/* 編集画面のツールバー幅と揃え、見出し位置が切り替えで動かないようにする */}
          <div className="toolbar toolbar--spacer" aria-hidden="true" />
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
          <button type="button" className="btn btn-export" onClick={handleExportExcel}>
            Excelで保存
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
    </div>
  )
}
