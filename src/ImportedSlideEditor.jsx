import { useState, useRef, useCallback, useEffect } from 'react'

// ── Helpers ───────────────────────────────────────────────────────────────────

const readFileAsDataUrl = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result)
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })

// ── Editable Image Slot ───────────────────────────────────────────────────────

function EditableImage({ el, onReplace }) {
  const [isDragging, setIsDragging] = useState(false)
  const inputRef = useRef(null)

  const handleDrop = useCallback(async (e) => {
    e.preventDefault()
    setIsDragging(false)
    const file = e.dataTransfer.files?.[0]
    if (file && file.type.startsWith('image/')) {
      const dataUrl = await readFileAsDataUrl(file)
      onReplace(dataUrl)
    }
  }, [onReplace])

  const handleFileChange = useCallback(async (e) => {
    const file = e.target.files?.[0]
    if (file) {
      const dataUrl = await readFileAsDataUrl(file)
      onReplace(dataUrl)
      e.target.value = ''
    }
  }, [onReplace])

  return (
    <div
      style={{
        position: 'absolute',
        left: `${el.xPct}%`,
        top: `${el.yPct}%`,
        width: `${el.wPct}%`,
        height: `${el.hPct}%`,
        boxSizing: 'border-box',
        border: isDragging
          ? '2px solid #0b7a38'
          : '1.5px dashed rgba(11,122,56,0.55)',
        borderRadius: '3px',
        overflow: 'hidden',
        cursor: 'pointer',
        background: el.dataUrl ? 'transparent' : 'rgba(11,122,56,0.06)',
      }}
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true) }}
      onDragEnter={(e) => { e.preventDefault(); setIsDragging(true) }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      title="Click or drag an image to replace"
    >
      {el.dataUrl && (
        <img
          src={el.dataUrl}
          alt=""
          style={{ width: '100%', height: '100%', objectFit: 'contain', display: 'block', pointerEvents: 'none' }}
        />
      )}

      <div style={{
        position: 'absolute',
        bottom: 0, left: 0, right: 0,
        background: 'rgba(0,0,0,0.5)',
        color: '#fff',
        fontSize: '0.7cqw',
        textAlign: 'center',
        padding: '1px 0',
        letterSpacing: '0.04em',
        opacity: el.dataUrl ? 0 : 1,
        transition: 'opacity 0.18s',
      }} className="ise-img-hint">
        {el.dataUrl ? '🔄 Replace' : '🖼 Drop image'}
      </div>

      <input
        ref={inputRef}
        type="file"
        accept="image/*"
        style={{ display: 'none' }}
        onChange={handleFileChange}
      />
    </div>
  )
}

// ── Editable Text Field ───────────────────────────────────────────────────────

function EditableText({ el, onChange }) {
  const multiline = el.hPct > 8

  const baseStyle = {
    width: '100%',
    border: '1.5px dashed rgba(225,140,0,0.55)',
    borderRadius: '3px',
    background: 'rgba(255,255,240,0.82)',
    fontFamily: el.fontFace || 'Calibri',
    fontSize: `${Math.max(el.fontSizePct || 1.5, 0.8)}cqw`,
    fontWeight: el.bold ? 'bold' : 'normal',
    color: el.color ? `#${el.color}` : '#111111',
    textAlign: el.align || 'left',
    padding: '1px 4px',
    boxSizing: 'border-box',
    outline: 'none',
    lineHeight: '1.25',
    resize: 'none',
  }

  return (
    <div style={{
      position: 'absolute',
      left: `${el.xPct}%`,
      top: `${el.yPct}%`,
      width: `${el.wPct}%`,
      height: `${el.hPct}%`,
      boxSizing: 'border-box',
      display: 'flex',
      alignItems: 'center',
    }}>
      {multiline ? (
        <textarea
          value={el.text}
          onChange={(e) => onChange(e.target.value)}
          style={{ ...baseStyle, height: '100%' }}
        />
      ) : (
        <input
          type="text"
          value={el.text}
          onChange={(e) => onChange(e.target.value)}
          style={baseStyle}
        />
      )}
    </div>
  )
}

// ── Editable Table ────────────────────────────────────────────────────────────

function EditableTable({ el, onCellChange }) {
  return (
    <div style={{
      position: 'absolute',
      left: `${el.xPct}%`,
      top: `${el.yPct}%`,
      width: `${el.wPct}%`,
      height: `${el.hPct}%`,
      boxSizing: 'border-box',
      overflow: 'hidden',
      border: '1px solid rgba(0,70,200,0.35)',
      borderRadius: '2px',
    }}>
      <table style={{
        width: '100%',
        height: '100%',
        borderCollapse: 'collapse',
        tableLayout: 'fixed',
      }}>
        <tbody>
          {el.rows.map((row, ri) => (
            <tr key={ri}>
              {row.map((cell, ci) => (
                <td
                  key={ci}
                  style={{
                    border: `1px solid ${cell.borderColor || '#bbbbbb'}`,
                    background: cell.fillColor ? `#${cell.fillColor}` : 'transparent',
                    padding: '1px 3px',
                    verticalAlign: 'middle',
                    overflow: 'hidden',
                  }}
                >
                  <input
                    type="text"
                    value={cell.text}
                    onChange={(e) => onCellChange(ri, ci, e.target.value)}
                    style={{
                      width: '100%',
                      border: 'none',
                      background: 'transparent',
                      fontFamily: cell.fontFace || 'Calibri',
                      fontSize: `${Math.max(el.fontSizePct || 1.2, 0.7)}cqw`,
                      fontWeight: cell.bold ? 'bold' : 'normal',
                      color: cell.color ? `#${cell.color}` : '#111111',
                      textAlign: cell.align || 'left',
                      outline: 'none',
                      padding: 0,
                      minWidth: 0,
                    }}
                  />
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

// ── Main component ────────────────────────────────────────────────────────────

/**
 * Renders an editable preview of an imported PPTX slide (first or last).
 *
 * Props:
 *  title         – label shown in the panel header
 *  slideData     – object from importFirstLastSlideData():
 *                  { images, texts, tables, backgroundDataUrl, rawXml, rawRelsXml, mediaFiles, fullRelMap }
 *  templateUrl   – URL/dataUrl of the template slide PNG (shown if user clicks Use Template)
 *  onUseTemplate – () => void — called when user discards imported content
 *  onChange      – (updatedSlideData) => void — called on any edit
 */
export function ImportedSlideEditor({ title, slideData, templateUrl, onUseTemplate, onChange }) {
  // Local editable state (mirrors slideData but is mutable)
  const [images, setImages] = useState(() => slideData?.images || [])
  const [texts, setTexts] = useState(() => slideData?.texts || [])
  const [tables, setTables] = useState(() => slideData?.tables || [])

  // Re-sync when the parent passes new slideData (re-import)
  useEffect(() => {
    if (slideData) {
      setImages(slideData.images || [])
      setTexts(slideData.texts || [])
      setTables(slideData.tables || [])
    }
  }, [slideData])

  // Notify parent of any change
  const notifyChange = useCallback((nextImages, nextTexts, nextTables) => {
    onChange?.({
      ...slideData,
      images: nextImages,
      texts: nextTexts,
      tables: nextTables,
    })
  }, [onChange, slideData])

  const handleReplaceImage = useCallback((idx, newDataUrl) => {
    setImages((prev) => {
      const next = prev.map((img, i) =>
        i === idx
          ? { ...img, dataUrl: newDataUrl, replaced: true }
          : img
      )
      notifyChange(next, texts, tables)
      return next
    })
  }, [notifyChange, texts, tables])

  const handleTextChange = useCallback((idx, newText) => {
    setTexts((prev) => {
      const next = prev.map((t, i) => (i === idx ? { ...t, text: newText } : t))
      notifyChange(images, next, tables)
      return next
    })
  }, [notifyChange, images, tables])

  const handleCellChange = useCallback((tblIdx, ri, ci, newText) => {
    setTables((prev) => {
      const next = prev.map((tbl, ti) => {
        if (ti !== tblIdx) return tbl
        return {
          ...tbl,
          rows: tbl.rows.map((row, rowI) =>
            rowI === ri
              ? row.map((cell, colI) => (colI === ci ? { ...cell, text: newText } : cell))
              : row
          ),
        }
      })
      notifyChange(images, texts, next)
      return next
    })
  }, [notifyChange, images, texts])

  if (!slideData) return null

  const hasElements = images.length > 0 || texts.length > 0 || tables.length > 0

  return (
    <article className="pair imported-slide-editor">
      {/* Header */}
      <div className="pair__header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
          <p className="label" style={{ margin: 0, color: '#7c3aed', fontWeight: 700 }}>
            📎 {title}
          </p>
          <span
            className="pair__status is-complete"
            style={{
              background: 'rgba(124,58,237,0.1)',
              borderColor: 'rgba(124,58,237,0.4)',
              color: '#7c3aed',
            }}
          >
            Imported from PPTX
          </span>
          {images.some(img => img.replaced) && (
            <span className="pair__status is-complete">✏️ Image replaced</span>
          )}
        </div>
        <button
          type="button"
          className="ghost"
          onClick={onUseTemplate}
          style={{ borderColor: '#7c3aed', color: '#7c3aed' }}
          title="Discard imported content and use the template slide instead"
        >
          🔄 Use Template Slide
        </button>
      </div>

      {/* Hint */}
      <p style={{ margin: 0, fontSize: '12px', color: 'var(--muted)' }}>
        ✏️ Click any <strong>text</strong> to edit inline ·
        Drop an image on a <strong>picture slot</strong> to replace it ·
        <strong> Table cells</strong> are individually editable ·
        All edits are embedded as valid PPTX XML
      </p>

      {/* Slide preview canvas */}
      <div
        className="slide-canvas imported-slide-canvas"
        style={{
          backgroundImage: slideData.backgroundDataUrl
            ? `url(${slideData.backgroundDataUrl})`
            : templateUrl
              ? `url(${templateUrl})`
              : 'none',
        }}
      >
        {!hasElements && (
          <div style={{
            position: 'absolute', inset: 0,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            color: 'var(--muted)', fontSize: '1.2cqw', pointerEvents: 'none',
          }}>
            No editable elements found on this slide
          </div>
        )}

        {/* Images */}
        {images.map((img, idx) => (
          <EditableImage
            key={img.id}
            el={img}
            onReplace={(dataUrl) => handleReplaceImage(idx, dataUrl)}
          />
        ))}

        {/* Texts */}
        {texts.map((txt, idx) => (
          <EditableText
            key={txt.id}
            el={txt}
            onChange={(val) => handleTextChange(idx, val)}
          />
        ))}

        {/* Tables */}
        {tables.map((tbl, tblIdx) => (
          <EditableTable
            key={tbl.id}
            el={tbl}
            onCellChange={(ri, ci, val) => handleCellChange(tblIdx, ri, ci, val)}
          />
        ))}
      </div>

      <p style={{ margin: 0, fontSize: '11px', color: 'var(--muted)', fontStyle: 'italic' }}>
        💡 Your edits are written into the downloaded PPTX as valid XML — images, text shapes and table cells are fully preserved including formatting.
      </p>
    </article>
  )
}
