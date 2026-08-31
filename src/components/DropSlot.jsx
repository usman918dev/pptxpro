import { useState, useRef, useEffect } from 'react'
import { readFileAsDataUrl } from '../utils/pairUtils'

const UNDO_DURATION_MS = 5000

export function DropSlot({ label, value, onChange, className = '', urlMode = 'inline', style }) {
  const [isDragging, setIsDragging] = useState(false)
  const [urlInput, setUrlInput] = useState('')
  const [pendingDelete, setPendingDelete] = useState(null) // stores the image url pending deletion
  const [undoProgress, setUndoProgress] = useState(100) // 100 → 0 countdown
  const fileInputRef = useRef(null)
  const undoTimerRef = useRef(null)
  const undoRafRef = useRef(null)

  const clearUndoTimer = () => {
    if (undoTimerRef.current) {
      clearTimeout(undoTimerRef.current)
      undoTimerRef.current = null
    }
    if (undoRafRef.current) {
      cancelAnimationFrame(undoRafRef.current)
      undoRafRef.current = null
    }
  }

  const commitDelete = () => {
    clearUndoTimer()
    setPendingDelete(null)
    setUndoProgress(100)
    onChange('')
  }

  const handleDeleteClick = () => {
    clearUndoTimer()
    const imageToDelete = value
    setPendingDelete(imageToDelete)
    setUndoProgress(100)

    // Animate the countdown ring
    const startTime = performance.now()
    const tick = (now) => {
      const elapsed = now - startTime
      const remaining = Math.max(0, 100 - (elapsed / UNDO_DURATION_MS) * 100)
      setUndoProgress(remaining)
      if (remaining > 0) {
        undoRafRef.current = requestAnimationFrame(tick)
      }
    }
    undoRafRef.current = requestAnimationFrame(tick)

    undoTimerRef.current = setTimeout(() => {
      commitDelete()
    }, UNDO_DURATION_MS)
  }

  const handleUndo = () => {
    clearUndoTimer()
    setPendingDelete(null)
    setUndoProgress(100)
    // value is already set (we never called onChange) so nothing more needed
  }

  const handleFile = async (file) => {
    if (!file) {
      return
    }
    clearUndoTimer()
    setPendingDelete(null)
    const dataUrl = await readFileAsDataUrl(file)
    onChange(dataUrl)
  }

  const handleDrop = (event) => {
    event.preventDefault()
    setIsDragging(false)
    const file = event.dataTransfer.files?.[0]
    if (file) {
      handleFile(file)
    }
  }

  const handleDragOver = (event) => {
    event.preventDefault()
  }

  const handleDragEnter = (event) => {
    event.preventDefault()
    setIsDragging(true)
  }

  const handleDragLeave = () => {
    setIsDragging(false)
  }

  const handleBrowse = (event) => {
    const file = event.target.files?.[0]
    if (file) {
      handleFile(file)
    }
    event.target.value = ''
  }

  const handleUrlAdd = (nextUrl) => {
    const trimmed = (nextUrl ?? urlInput).trim()
    if (!trimmed) {
      return
    }
    clearUndoTimer()
    setPendingDelete(null)
    onChange(trimmed)
    if (nextUrl === undefined) {
      setUrlInput('')
    }
  }

  const handleUrlPrompt = () => {
    const response = window.prompt('Paste image URL')
    if (response) {
      handleUrlAdd(response)
    }
  }

  const handleUrlKeyDown = (event) => {
    if (event.key === 'Enter') {
      event.preventDefault()
      handleUrlAdd()
    }
  }

  // Cleanup on unmount
  useEffect(() => {
    return () => clearUndoTimer()
  }, [])

  // Determine what image to show: if pending delete, show the pending image (greyed)
  const displayValue = pendingDelete !== null ? pendingDelete : value
  const isPendingDelete = pendingDelete !== null
  const slotClassName = `drop-slot${displayValue ? ' has-image' : ''}${
    isDragging ? ' is-dragging' : ''
  }${isPendingDelete ? ' is-pending-delete' : ''} ${className}`

  // SVG ring circumference for the countdown
  const RING_R = 14
  const RING_CIRC = 2 * Math.PI * RING_R
  const ringDash = (undoProgress / 100) * RING_CIRC

  return (
    <div
      className={slotClassName.trim()}
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      onDragEnter={handleDragEnter}
      onDragLeave={handleDragLeave}
      style={style}
    >
      <div className="drop-slot__label">{label}</div>
      {displayValue ? (
        <div className="drop-slot__preview-wrap">
          <img
            className={`drop-slot__preview${isPendingDelete ? ' is-fading' : ''}`}
            src={displayValue}
            alt={`${label} preview`}
            style={style?.borderRadius !== undefined ? { borderRadius: style.borderRadius } : {}}
          />
          {isPendingDelete ? (
            <button
              type="button"
              className="drop-slot__undo"
              onClick={handleUndo}
              aria-label={`Undo remove ${label} image`}
            >
              <svg className="drop-slot__undo-ring" viewBox="0 0 36 36" aria-hidden="true">
                <circle
                  cx="18" cy="18" r={RING_R}
                  fill="none"
                  stroke="rgba(255,255,255,0.25)"
                  strokeWidth="3"
                />
                <circle
                  cx="18" cy="18" r={RING_R}
                  fill="none"
                  stroke="#fff"
                  strokeWidth="3"
                  strokeDasharray={`${ringDash} ${RING_CIRC}`}
                  strokeLinecap="round"
                  transform="rotate(-90 18 18)"
                />
              </svg>
              <span>↩ Undo</span>
            </button>
          ) : (
            <button
              type="button"
              className="drop-slot__delete"
              onClick={handleDeleteClick}
              aria-label={`Remove ${label} image`}
            >
              Delete
            </button>
          )}
        </div>
      ) : (
        <div className="drop-slot__placeholder">
          <p>Drag & drop an image</p>
          <span>or click Browse</span>
        </div>
      )}
      {!displayValue && (
        <>
          <div className="drop-slot__actions">
            <button
              type="button"
              className="ghost"
              onClick={() => fileInputRef.current?.click()}
            >
              Browse
            </button>
            {urlMode === 'prompt' && (
              <button type="button" className="ghost" onClick={handleUrlPrompt}>
                Paste URL
              </button>
            )}
          </div>
          {urlMode === 'inline' && (
            <div className="drop-slot__url">
              <input
                type="text"
                placeholder="Paste image URL"
                value={urlInput}
                onChange={(event) => setUrlInput(event.target.value)}
                onKeyDown={handleUrlKeyDown}
              />
              <button type="button" className="ghost" onClick={handleUrlAdd}>
                Add URL
              </button>
            </div>
          )}
        </>
      )}
      <input
        ref={fileInputRef}
        className="drop-slot__file"
        type="file"
        accept="image/*"
        onChange={handleBrowse}
      />
    </div>
  )
}
