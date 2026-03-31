import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import { createPortal } from 'react-dom'
import { DndContext, KeyboardSensor, PointerSensor, closestCenter, useSensor, useSensors } from '@dnd-kit/core'
import { SortableContext, sortableKeyboardCoordinates, useSortable, verticalListSortingStrategy } from '@dnd-kit/sortable'
import { CSS } from '@dnd-kit/utilities'
import { DISPLAY_CROP_HEIGHT, DISPLAY_CROP_WIDTH, loadImageFromUrl } from './imageCropUtils.js'

import './ScheduleComponent.css'

/** 24h・15分刻み（00:00 … 23:45） */
const TIME_OPTIONS_15MIN = (() => {
  const out = []
  for (let h = 0; h < 24; h++) {
    for (const m of [0, 15, 30, 45]) {
      out.push(`${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`)
    }
  }
  return out
})()

/** 既存テキストを最寄りの15分スロットへ（一覧に無い場合は ''） */
function normalizeTimeToSlot15(s) {
  const t = String(s ?? '').trim()
  if (!t) return ''
  const m = t.match(/^(\d{1,2}):(\d{2})$/)
  if (!m) return ''
  let h = parseInt(m[1], 10)
  let min = parseInt(m[2], 10)
  if (Number.isNaN(h) || Number.isNaN(min) || h < 0 || h > 23 || min < 0 || min > 59) return ''
  const total = h * 60 + min
  const snapped = Math.round(total / 15) * 15
  const maxMin = 23 * 60 + 45
  const clamped = Math.min(maxMin, Math.max(0, snapped))
  const nh = Math.floor(clamped / 60)
  const nm = clamped % 60
  const slot = `${String(nh).padStart(2, '0')}:${String(nm).padStart(2, '0')}`
  return TIME_OPTIONS_15MIN.includes(slot) ? slot : ''
}

function parseHHmmToMinutes(s) {
  const t = String(s ?? '').trim()
  const m = t.match(/^(\d{1,2}):(\d{2})$/)
  if (!m) return null
  const h = parseInt(m[1], 10)
  const min = parseInt(m[2], 10)
  if (Number.isNaN(h) || Number.isNaN(min) || h < 0 || h > 23 || min < 0 || min > 59) return null
  return h * 60 + min
}

function durationLabelFromTimes(prevTime, curTime) {
  const prev = parseHHmmToMinutes(prevTime)
  const cur = parseHHmmToMinutes(curTime)
  if (prev == null || cur == null) return ''

  let diff = cur - prev
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

const SB_DRAG_MIME = 'application/x-ekonte-storyboard-cut'
const storyboardDragPayloads = new Map()

/** スケジュール列コンボボックスの候補（自由入力も可） */
const SCHEDULE_PRESET_OPTIONS = [
  'スタッフ集合・機材準備',
  'スタッフ・キャストIN',
  'モデルIN',
  'メイク・着替え開始',
  'SHOOT',
  '移動',
  'セットチェンジ・機材準備',
  '昼食・休憩',
  '撮影終了　おつかれさまでした！',
]

/**
 * 絵コンテ HistoryTextareaCell と同様：先頭行の直下を viewport 固定で表示し、親の overflow では切れない。
 */
function getScheduleDropdownPosition(textareaEl) {
  if (!textareaEl) return null
  const rect = textareaEl.getBoundingClientRect()
  const style = window.getComputedStyle(textareaEl)
  const padTop = parseFloat(style.paddingTop) || 0
  const fs = parseFloat(style.fontSize) || 14
  const lhRaw = style.lineHeight
  let lineHeight = fs * 1.4
  if (lhRaw && lhRaw !== 'normal') {
    const parsed = parseFloat(lhRaw)
    if (Number.isFinite(parsed)) lineHeight = parsed
  }
  const gap = 4
  const top = rect.top + padTop + lineHeight + gap
  const spaceBelow = window.innerHeight - top - 16
  const maxHeight = Math.min(280, Math.max(80, spaceBelow))
  return {
    top,
    left: rect.left,
    width: rect.width,
    maxHeight,
  }
}

function ScheduleScheduleCombobox({ value, onChange }) {
  const [open, setOpen] = useState(false)
  const [fixedPos, setFixedPos] = useState(null)
  const textareaRef = useRef(null)
  const dropdownRef = useRef(null)

  const close = useCallback(() => {
    setOpen(false)
    setFixedPos(null)
  }, [])

  const updatePosition = useCallback(() => {
    const pos = getScheduleDropdownPosition(textareaRef.current)
    setFixedPos(pos)
  }, [])

  useLayoutEffect(() => {
    if (open) {
      updatePosition()
    }
  }, [open, updatePosition, value])

  useEffect(() => {
    if (!open) return
    const onScrollOrResize = () => updatePosition()
    window.addEventListener('scroll', onScrollOrResize, true)
    window.addEventListener('resize', onScrollOrResize)
    return () => {
      window.removeEventListener('scroll', onScrollOrResize, true)
      window.removeEventListener('resize', onScrollOrResize)
    }
  }, [open, updatePosition])

  useEffect(() => {
    if (!open) return
    const onDoc = (e) => {
      const t = e.target
      if (textareaRef.current?.contains(t)) return
      if (dropdownRef.current?.contains(t)) return
      close()
    }
    document.addEventListener('mousedown', onDoc)
    return () => document.removeEventListener('mousedown', onDoc)
  }, [open, close])

  useEffect(() => {
    const onKey = (e) => {
      if (e.key === 'Escape') close()
    }
    if (open) {
      document.addEventListener('keydown', onKey)
      return () => document.removeEventListener('keydown', onKey)
    }
  }, [open, close])

  const handleBlur = useCallback(() => {
    requestAnimationFrame(() => {
      const ae = document.activeElement
      if (dropdownRef.current?.contains(ae) || textareaRef.current === ae) return
      close()
    })
  }, [close])

  const selectOption = useCallback(
    (text) => {
      onChange(text)
      close()
    },
    [close, onChange]
  )

  const dropdownContent =
    open && fixedPos ? (
      <div
        ref={dropdownRef}
        className="history-dropdown history-dropdown--portal"
        role="listbox"
        aria-label="スケジュール候補"
        style={{
          position: 'fixed',
          top: fixedPos.top,
          left: fixedPos.left,
          width: fixedPos.width,
          maxHeight: fixedPos.maxHeight,
          zIndex: 10050,
        }}
      >
        <div className="history-dropdown-list">
          {SCHEDULE_PRESET_OPTIONS.map((opt) => (
            <button
              type="button"
              key={opt}
              className="history-dropdown-item"
              role="option"
              onMouseDown={(e) => {
                e.preventDefault()
                selectOption(opt)
              }}
            >
              {opt}
            </button>
          ))}
        </div>
      </div>
    ) : null

  return (
    <div className="schedule-schedule-combobox">
      <textarea
        ref={textareaRef}
        className="schedule-textarea"
        value={value ?? ''}
        onChange={(e) => onChange(e.target.value)}
        onFocus={() => {
          setOpen(true)
          requestAnimationFrame(() => updatePosition())
        }}
        onBlur={handleBlur}
        spellCheck={false}
        aria-autocomplete="list"
        aria-expanded={open}
        aria-haspopup="listbox"
        autoComplete="off"
      />
      {dropdownContent && typeof document !== 'undefined' ? createPortal(dropdownContent, document.body) : null}
    </div>
  )
}

function SortableScheduleRow({
  row,
  durationLabel,
  onDeleteRow,
  onUpdateRow,
  handleDropImages,
  editingCut,
  setEditingCut,
}) {
  const { attributes, listeners, setNodeRef, setActivatorNodeRef, transform, transition, isDragging } = useSortable({
    id: row.id,
  })
  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  }

  const imageEntries = (row?.images ?? [])
    .map((img, idx) => (img ? { img, idx } : null))
    .filter(Boolean)
  const timeSlot = normalizeTimeToSlot15(row.time ?? '')

  return (
    <tr ref={setNodeRef} style={style} className={isDragging ? 'schedule-row-dragging' : ''}>
      <td className="schedule-td schedule-td--drag">
        <button
          type="button"
          className="schedule-row-drag-handle drag-handle"
          ref={setActivatorNodeRef}
          {...attributes}
          {...listeners}
          aria-label="行をドラッグして並び替え"
          title="ドラッグして並び替え"
        >
          ⋮⋮
        </button>
      </td>
      <td className="schedule-td schedule-td--time">
        <div className="schedule-time-cell">
          <button
            type="button"
            className="row-delete-btn"
            onClick={() => onDeleteRow(row.id)}
            aria-label="行を削除"
            title="この行を削除"
          >
            ×
          </button>
          <select
            className="schedule-time-select"
            value={timeSlot}
            onChange={(e) => onUpdateRow(row.id, { time: e.target.value })}
            aria-label="TIME"
          >
            <option value="" />
            {TIME_OPTIONS_15MIN.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
          {durationLabel && <div className="schedule-time-duration">{durationLabel}</div>}
        </div>
      </td>

      <td className="schedule-td schedule-td--schedule">
        <ScheduleScheduleCombobox
          value={row.schedule ?? ''}
          onChange={(next) => onUpdateRow(row.id, { schedule: next })}
        />
        <div
          className="image-container"
          onDrop={(e) => handleDropImages(row.id, e)}
          onDragOver={(e) => {
            e.preventDefault()
            e.dataTransfer.dropEffect = 'copy'
          }}
          aria-label="画像をドロップ"
        >
          {imageEntries.map(({ img, idx }) => {
            const cutStr = String(row?.cutNos?.[idx] ?? '')
            const cutDigits = Math.max(2, cutStr.trim().length)
            const isEditing = editingCut?.rowId === row.id && editingCut?.imageIndex === idx

            return (
              <div key={`${row.id}-img-${idx}`} className="image-item">
                <img className="image-item__img" src={img.objectUrl} alt="" draggable={false} />
                <button
                  type="button"
                  className="schedule-image-remove-btn"
                  onClick={(e) => {
                    e.preventDefault()
                    e.stopPropagation()
                    const nextImages = Array.isArray(row.images) ? row.images.slice() : []
                    const nextCutNos = Array.isArray(row.cutNos) ? row.cutNos.slice() : []
                    nextImages[idx] = null
                    nextCutNos[idx] = ''
                    onUpdateRow(row.id, { images: nextImages, cutNos: nextCutNos })
                  }}
                  aria-label="画像を削除"
                  title="この画像を削除"
                >
                  ×
                </button>

                <button
                  type="button"
                  className="schedule-cut-overlay schedule-cut-overlay--image"
                  onClick={(e) => {
                    e.preventDefault()
                    e.stopPropagation()
                    setEditingCut({ rowId: row.id, imageIndex: idx })
                  }}
                  aria-label="Cut番号を編集"
                  title="クリックでCut番号を編集"
                >
                  <span className="schedule-cut-overlay__label">Cut#</span>
                  {isEditing ? (
                    <input
                      className="schedule-cut-input"
                      autoFocus
                      value={row.cutNos?.[idx] ?? ''}
                      size={cutDigits}
                      onChange={(e) => {
                        const next = Array.isArray(row.cutNos) ? row.cutNos.slice() : []
                        next[idx] = e.target.value
                        onUpdateRow(row.id, { cutNos: next })
                      }}
                      onClick={(e) => e.stopPropagation()}
                      onBlur={() => setEditingCut(null)}
                    />
                  ) : (
                    <span className="schedule-cut-overlay__value">{row.cutNos?.[idx] ?? ''}</span>
                  )}
                </button>
              </div>
            )
          })}
        </div>
      </td>

      <td className="schedule-td schedule-td--note">
        <textarea
          className="schedule-textarea schedule-textarea--note"
          value={row.note ?? ''}
          onChange={(e) => onUpdateRow(row.id, { note: e.target.value })}
        />
      </td>
    </tr>
  )
}

export default function ScheduleComponent({
  rows,
  onUpdateRow,
  onDeleteRow,
  onAddRow,
  onImportExcel,
  onExportExcel,
  canExportExcel,
  onBackHome,
  onAddImages,
  storyboardSourceItems = null,
  onDropStoryboardCut,
  onReorderRows,
  onUndo,
  canUndo = false,
}) {
  const [editingCut, setEditingCut] = useState(null) // { rowId, imageIndex }
  const lastStoryboardDragTokenRef = useRef(null)
  const importInputRef = useRef(null)

  const scheduleItems = useMemo(() => {
    return Array.isArray(rows) ? rows : []
  }, [rows])

  useEffect(() => {
    scheduleItems.forEach((row) => {
      const raw = String(row.time ?? '').trim()
      if (!raw) return
      const n = normalizeTimeToSlot15(row.time ?? '')
      if (n === '') {
        onUpdateRow(row.id, { time: '' })
        return
      }
      if (n !== raw) onUpdateRow(row.id, { time: n })
    })
  }, [scheduleItems, onUpdateRow])

  const handleDropImages = useCallback(
    async (rowId, e) => {
      e.preventDefault()
      let sbToken = e.dataTransfer?.getData(SB_DRAG_MIME)
      if (!sbToken) sbToken = e.dataTransfer?.getData('text/plain')
      if (sbToken && typeof onDropStoryboardCut === 'function') {
        const payload = storyboardDragPayloads.get(sbToken)
        if (payload) {
          storyboardDragPayloads.delete(sbToken)
          await onDropStoryboardCut(rowId, payload)
          return
        }
      }

      const files = Array.from(e.dataTransfer?.files ?? [])
      const imageFiles = files.filter((f) => f && f.type && f.type.startsWith('image/'))
      if (imageFiles.length === 0) return

      const cw = DISPLAY_CROP_WIDTH
      const ch = DISPLAY_CROP_HEIGHT

      const imageStates = await Promise.all(
        imageFiles.map(async (file) => {
          const objectUrl = URL.createObjectURL(file)
          const img = await loadImageFromUrl(objectUrl)
          return {
            file,
            objectUrl,
            naturalWidth: img.naturalWidth,
            naturalHeight: img.naturalHeight,
            containerWidth: cw,
            containerHeight: ch,
            initialScale: 0,
            scale: 0,
            offsetX: 0,
            offsetY: 0,
            crop: null,
          }
        })
      )

      onAddImages(rowId, imageStates)
    },
    [onAddImages, onDropStoryboardCut]
  )

  const handleStoryboardPoolDragStart = useCallback((item, e) => {
    const token = `sb-${item.key}-${Date.now()}`
    lastStoryboardDragTokenRef.current = token
    storyboardDragPayloads.set(token, {
      image: item.image,
      cutNumber: item.cutNumber,
    })
    e.dataTransfer.setData(SB_DRAG_MIME, token)
    e.dataTransfer.setData('text/plain', token)
    e.dataTransfer.effectAllowed = 'copy'
  }, [])

  const handleStoryboardPoolDragEnd = useCallback(() => {
    const t = lastStoryboardDragTokenRef.current
    lastStoryboardDragTokenRef.current = null
    if (t && storyboardDragPayloads.has(t)) {
      storyboardDragPayloads.delete(t)
    }
  }, [])

  const scheduleSensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 8 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates })
  )

  const handleScheduleDragEnd = useCallback(
    (event) => {
      const { active, over } = event
      if (!over || active.id === over.id) return
      if (typeof onReorderRows === 'function') {
        onReorderRows(active.id, over.id)
      }
    },
    [onReorderRows]
  )

  return (
    <div className="app schedule-page">
      <header className="header">
        <h1 className="app-heading">香盤表メーカー</h1>
        <div className="toolbar">
          <button type="button" className="btn btn-secondary" onClick={onBackHome}>
            ホーム
          </button>
          <button type="button" className="btn btn-secondary" onClick={() => importInputRef.current?.click()}>
            保存されたExcelで継続
          </button>
          <input
            ref={importInputRef}
            type="file"
            accept=".xlsx,.xlsm,.xls"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0]
              if (file && typeof onImportExcel === 'function') onImportExcel(file)
              e.target.value = ''
            }}
          />
          <button type="button" className="btn btn-secondary" onClick={onUndo} disabled={!canUndo}>
            戻す
          </button>
          <button
            type="button"
            className="btn btn-export"
            onClick={onExportExcel}
            disabled={!canExportExcel}
          >
            Excelで保存
          </button>
        </div>
      </header>

      {Array.isArray(storyboardSourceItems) && storyboardSourceItems.length > 0 && (
        <section className="schedule-storyboard-source" aria-label="絵コンテのカット">
          <h2 className="schedule-storyboard-source__title">カットを選択して入れる</h2>
          <p className="schedule-storyboard-source__hint">
            カットをドラッグして、下の表のスケジュール欄の画像エリアにドロップしてください。
          </p>
          <div className="schedule-storyboard-source__grid">
            {storyboardSourceItems.map((item) => (
              <div
                key={item.key}
                className="schedule-storyboard-source-item"
                draggable
                onDragStart={(e) => handleStoryboardPoolDragStart(item, e)}
                onDragEnd={handleStoryboardPoolDragEnd}
              >
                <img className="schedule-storyboard-source-item__img" src={item.image.objectUrl} alt="" draggable={false} />
                <div className="schedule-cut-overlay schedule-cut-overlay--storyboard-pool" aria-hidden>
                  <span className="schedule-cut-overlay__label">Cut#</span>
                  <span className="schedule-cut-overlay__value">{item.cutNumber}</span>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      <div className="schedule-editor-panel">
        <table className="schedule-table">
          <colgroup>
            <col className="schedule-col schedule-col--drag" />
            <col className="schedule-col schedule-col--time" />
            <col className="schedule-col schedule-col--schedule" />
            <col className="schedule-col schedule-col--note" />
          </colgroup>
          <thead>
            <tr>
              <th className="schedule-th schedule-th--drag" />
              <th>TIME</th>
              <th>スケジュール</th>
              <th>備考</th>
            </tr>
          </thead>
          <DndContext sensors={scheduleSensors} collisionDetection={closestCenter} onDragEnd={handleScheduleDragEnd}>
            <SortableContext items={scheduleItems.map((r) => r.id)} strategy={verticalListSortingStrategy}>
              <tbody>
                {scheduleItems.map((row, rowIndex) => (
                  <SortableScheduleRow
                    key={row.id}
                    row={row}
                    durationLabel={
                      rowIndex < scheduleItems.length - 1
                        ? durationLabelFromTimes(row.time, scheduleItems[rowIndex + 1]?.time)
                        : ''
                    }
                    onDeleteRow={onDeleteRow}
                    onUpdateRow={onUpdateRow}
                    handleDropImages={handleDropImages}
                    editingCut={editingCut}
                    setEditingCut={setEditingCut}
                  />
                ))}
              </tbody>
            </SortableContext>
          </DndContext>
        </table>

        <button type="button" className="grid-add-row-zone" onClick={onAddRow} aria-label="行を追加">
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

