import React, { useState, useRef, useEffect, useCallback } from 'react'

// ─── Preset Library Panel ────────────────────────────────────────────────────

function PresetLibrary({ presets, activePresetId, onLoad, onDelete, onRename, onSaveNew, onExportPreset, onExportAll, onImportPresets }) {
  const [newName, setNewName] = useState('')
  const [renamingId, setRenamingId] = useState(null)
  const [renameValue, setRenameValue] = useState('')
  const importInputRef = useRef(null)

  const handleSaveNew = () => {
    const name = newName.trim() || `Preset ${presets.length + 1}`
    onSaveNew(name)
    setNewName('')
  }

  const startRename = (preset) => {
    setRenamingId(preset.id)
    setRenameValue(preset.name)
  }

  const commitRename = (id) => {
    if (renameValue.trim()) onRename(id, renameValue.trim())
    setRenamingId(null)
  }

  return (
    <div className="preset-library card">
      <div className="preset-library__header">
        <div>
          <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700 }}>Master Slide Library</h3>
          <p style={{ margin: '3px 0 0', fontSize: '12px', color: 'var(--muted)' }}>
            Save unlimited named presets — switch between them instantly
          </p>
        </div>
        <div className="preset-library__save-row">
          <input
            type="text"
            placeholder="Preset name…"
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && handleSaveNew()}
            style={{
              borderRadius: '999px',
              border: '1px solid var(--border)',
              padding: '8px 14px',
              fontSize: '13px',
              fontFamily: 'inherit',
              background: '#fff',
              color: 'var(--ink)',
              width: '180px',
              outline: 'none',
            }}
          />
          <button type="button" className="button" style={{ padding: '8px 18px', fontSize: '12px' }} onClick={handleSaveNew}>
            💾 Save as New
          </button>
        </div>
      </div>

      {/* ── Export All / Import row ── */}
      <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center', marginTop: '10px', paddingTop: '10px', borderTop: '1px solid var(--border)' }}>
        <button
          type="button"
          className="ghost"
          style={{ padding: '6px 14px', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '5px' }}
          onClick={onExportAll}
          title="Download all templates as a JSON file"
        >
          📤 Export All Templates
        </button>
        <button
          type="button"
          className="ghost"
          style={{ padding: '6px 14px', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '5px' }}
          onClick={() => importInputRef.current?.click()}
          title="Import templates from a JSON file"
        >
          📥 Import Templates
        </button>
        <input
          ref={importInputRef}
          type="file"
          accept=".json,application/json"
          style={{ display: 'none' }}
          onChange={(e) => {
            const file = e.target.files?.[0]
            if (file) onImportPresets(file)
            e.target.value = ''
          }}
        />
        {presets.length > 0 && (
          <span style={{ fontSize: '11px', color: 'var(--muted)', marginLeft: 'auto' }}>
            {presets.length} template{presets.length !== 1 ? 's' : ''} saved
          </span>
        )}
      </div>

      {presets.length === 0 ? (
        <p style={{ margin: 0, fontSize: '13px', color: 'var(--muted)', fontStyle: 'italic' }}>
          No presets yet. Design your layout and save it above.
        </p>
      ) : (
        <div className="preset-library__grid">
          {presets.map((preset) => {
            const isActive = preset.id === activePresetId
            return (
              <div
                key={preset.id}
                className={`preset-card${isActive ? ' is-active' : ''}`}
              >
                <div className="preset-card__thumb">
                  {preset.layout?.masterBgUrl ? (
                    <img src={preset.layout.masterBgUrl} alt={preset.name} />
                  ) : (
                    <div className="preset-card__thumb-empty">
                      <span>🖼️</span>
                    </div>
                  )}
                  {isActive && <div className="preset-card__active-badge">Active</div>}
                </div>
                <div className="preset-card__body">
                  {renamingId === preset.id ? (
                    <input
                      className="preset-card__rename-input"
                      value={renameValue}
                      autoFocus
                      onChange={(e) => setRenameValue(e.target.value)}
                      onBlur={() => commitRename(preset.id)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') commitRename(preset.id)
                        if (e.key === 'Escape') setRenamingId(null)
                      }}
                    />
                  ) : (
                    <p className="preset-card__name" onDoubleClick={() => startRename(preset)}>
                      {preset.name}
                    </p>
                  )}
                  <p className="preset-card__meta">
                    {preset.layout?.placeholders?.length || 0} img · {preset.layout?.textboxes?.length || 0} txt
                  </p>
                  <div className="preset-card__actions">
                    <button
                      type="button"
                      className="ghost"
                      style={{ padding: '4px 10px', fontSize: '10px' }}
                      onClick={() => onLoad(preset.id)}
                      disabled={isActive}
                    >
                      {isActive ? '✓ Loaded' : 'Load'}
                    </button>
                    <button
                      type="button"
                      className="ghost"
                      style={{ padding: '4px 10px', fontSize: '10px' }}
                      onClick={() => startRename(preset)}
                    >
                      Rename
                    </button>
                    <button
                      type="button"
                      className="ghost"
                      style={{ padding: '4px 10px', fontSize: '10px', display: 'flex', alignItems: 'center', gap: '3px' }}
                      onClick={() => onExportPreset(preset.id)}
                      title="Download this template as a JSON file"
                    >
                      📥 Download
                    </button>
                    <button
                      type="button"
                      style={{
                        padding: '4px 10px',
                        fontSize: '10px',
                        background: 'rgba(225,43,43,0.08)',
                        border: '1px solid rgba(225,43,43,0.3)',
                        borderRadius: '999px',
                        color: '#b51d1d',
                        cursor: 'pointer',
                        letterSpacing: '0.1em',
                        textTransform: 'uppercase',
                      }}
                      onClick={() => {
                        if (window.confirm(`Delete preset "${preset.name}"?`)) onDelete(preset.id)
                      }}
                    >
                      Delete
                    </button>
                  </div>
                </div>
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}

// ─── Main MasterDesigner ─────────────────────────────────────────────────────

export function MasterDesigner({
  customLayout,
  onSave,
  presets = [],
  activePresetId = null,
  onSavePreset,
  onLoadPreset,
  onDeletePreset,
  onRenamePreset,
  onExportPreset,
  onExportAll,
  onImportPresets,
}) {
  const [firstSlideUrl, setFirstSlideUrl] = useState(customLayout?.firstSlideUrl || '')
  const [lastSlideUrl, setLastSlideUrl] = useState(customLayout?.lastSlideUrl || '')
  const [masterBgUrl, setMasterBgUrl] = useState(customLayout?.masterBgUrl || '')

  const [placeholders, setPlaceholders] = useState(customLayout?.placeholders || [])
  const [textboxes, setTextboxes] = useState(customLayout?.textboxes || [])

  const [firstSlidePlaceholders, setFirstSlidePlaceholders] = useState(customLayout?.firstSlidePlaceholders || [])
  const [firstSlideTextboxes, setFirstSlideTextboxes] = useState(customLayout?.firstSlideTextboxes || [])

  const [lastSlidePlaceholders, setLastSlidePlaceholders] = useState(customLayout?.lastSlidePlaceholders || [])
  const [lastSlideTextboxes, setLastSlideTextboxes] = useState(customLayout?.lastSlideTextboxes || [])

  const [editSlideType, setEditSlideType] = useState('master')
  const [selectedId, setSelectedId] = useState(null)
  const [dragState, setDragState] = useState(null)

  // Extract mode: draw-to-create textbox
  const [extractMode, setExtractMode] = useState(false)
  const [drawState, setDrawState] = useState(null) // { startXPct, startYPct }
  const [drawRect, setDrawRect] = useState(null)   // { xPct, yPct, wPct, hPct }

  const canvasRef = useRef(null)

  // ── Sync when customLayout prop changes (preset loaded) ──────────────────
  useEffect(() => {
    if (customLayout) {
      setFirstSlideUrl(customLayout.firstSlideUrl || '')
      setLastSlideUrl(customLayout.lastSlideUrl || '')
      setMasterBgUrl(customLayout.masterBgUrl || '')
      setPlaceholders(customLayout.placeholders || [])
      setTextboxes(customLayout.textboxes || [])
      setFirstSlidePlaceholders(customLayout.firstSlidePlaceholders || [])
      setFirstSlideTextboxes(customLayout.firstSlideTextboxes || [])
      setLastSlidePlaceholders(customLayout.lastSlidePlaceholders || [])
      setLastSlideTextboxes(customLayout.lastSlideTextboxes || [])
      setSelectedId(null)
    }
  }, [customLayout])

  // ── Active list helpers ──────────────────────────────────────────────────
  const getActiveList = useCallback(() => {
    if (editSlideType === 'first') return { placeholders: firstSlidePlaceholders, textboxes: firstSlideTextboxes }
    if (editSlideType === 'last') return { placeholders: lastSlidePlaceholders, textboxes: lastSlideTextboxes }
    return { placeholders, textboxes }
  }, [editSlideType, placeholders, textboxes, firstSlidePlaceholders, firstSlideTextboxes, lastSlidePlaceholders, lastSlideTextboxes])

  const setActivePlaceholders = (val) => {
    if (editSlideType === 'first') setFirstSlidePlaceholders(val)
    else if (editSlideType === 'last') setLastSlidePlaceholders(val)
    else setPlaceholders(val)
  }

  const setActiveTextboxes = (val) => {
    if (editSlideType === 'first') setFirstSlideTextboxes(val)
    else if (editSlideType === 'last') setLastSlideTextboxes(val)
    else setTextboxes(val)
  }

  const handleFileChange = (e, setter) => {
    const file = e.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => setter(reader.result)
    reader.readAsDataURL(file)
  }

  const addPlaceholder = () => {
    const id = `placeholder_${Date.now()}`
    const { placeholders: activeP } = getActiveList()
    const count = activeP.length
    setActivePlaceholders([...activeP, {
      id,
      key: `${editSlideType}_image_${count}`,
      label: `Image Placeholder ${count + 1}`,
      x: 1.0 + (count * 0.5) % 8,
      y: 2.0,
      w: 4.0,
      h: 3.5,
    }])
    setSelectedId(id)
  }

  const addTextbox = (overrides = {}) => {
    const id = `textbox_${Date.now()}`
    const { textboxes: activeT } = getActiveList()
    const count = activeT.length
    setActiveTextboxes([...activeT, {
      id,
      key: `${editSlideType}_text_${count}`,
      textDefault: 'Enter description text',
      x: 1.0 + (count * 0.5) % 8,
      y: 5.8,
      w: 8.0,
      h: 0.8,
      fontSize: 20,
      fontColor: '111111',
      fontFace: 'Calibri',
      bold: true,
      align: 'center',
      ...overrides,
    }])
    setSelectedId(id)
  }

  const deleteElement = (id) => {
    const { placeholders: activeP, textboxes: activeT } = getActiveList()
    setActivePlaceholders(activeP.filter((p) => p.id !== id))
    setActiveTextboxes(activeT.filter((t) => t.id !== id))
    if (selectedId === id) setSelectedId(null)
  }

  const getElement = useCallback((id) => {
    const { placeholders: activeP, textboxes: activeT } = getActiveList()
    return activeP.find((p) => p.id === id) || activeT.find((t) => t.id === id)
  }, [getActiveList])

  const updateElement = useCallback((id, fields) => {
    const { placeholders: activeP, textboxes: activeT } = getActiveList()
    setActivePlaceholders(activeP.map((p) => (p.id === id ? { ...p, ...fields } : p)))
    setActiveTextboxes(activeT.map((t) => (t.id === id ? { ...t, ...fields } : t)))
  }, [getActiveList])

  const { placeholders: activePlaceholders, textboxes: activeTextboxes } = getActiveList()
  const selectedElement = getElement(selectedId)

  // ── Drag/Resize mouse handling ───────────────────────────────────────────
  const handleMouseDown = (e, id, type) => {
    e.stopPropagation()
    e.preventDefault()
    setSelectedId(id)
    if (!canvasRef.current) return
    const canvasRect = canvasRef.current.getBoundingClientRect()
    const element = getElement(id)
    if (!element) return
    setDragState({
      id,
      type,
      startX: e.clientX,
      startY: e.clientY,
      startLeft: (element.x / 13.333) * canvasRect.width,
      startTop: (element.y / 7.5) * canvasRect.height,
      startWidth: (element.w / 13.333) * canvasRect.width,
      startHeight: (element.h / 7.5) * canvasRect.height,
    })
  }

  useEffect(() => {
    const handleMouseMove = (e) => {
      if (!dragState || !canvasRef.current) return
      const canvasRect = canvasRef.current.getBoundingClientRect()
      const deltaX = e.clientX - dragState.startX
      const deltaY = e.clientY - dragState.startY
      const element = getElement(dragState.id)
      if (!element) return

      if (dragState.type === 'move') {
        let newX = ((dragState.startLeft + deltaX) / canvasRect.width) * 13.333
        let newY = ((dragState.startTop + deltaY) / canvasRect.height) * 7.5
        newX = Math.max(0, Math.min(13.333 - element.w, newX))
        newY = Math.max(0, Math.min(7.5 - element.h, newY))
        updateElement(dragState.id, { x: parseFloat(newX.toFixed(2)), y: parseFloat(newY.toFixed(2)) })
      } else if (dragState.type === 'resize') {
        let newW = ((dragState.startWidth + deltaX) / canvasRect.width) * 13.333
        let newH = ((dragState.startHeight + deltaY) / canvasRect.height) * 7.5
        newW = Math.max(0.5, Math.min(13.333 - element.x, newW))
        newH = Math.max(0.2, Math.min(7.5 - element.y, newH))
        updateElement(dragState.id, { w: parseFloat(newW.toFixed(2)), h: parseFloat(newH.toFixed(2)) })
      }
    }
    const handleMouseUp = () => setDragState(null)

    if (dragState) {
      window.addEventListener('mousemove', handleMouseMove)
      window.addEventListener('mouseup', handleMouseUp)
    }
    return () => {
      window.removeEventListener('mousemove', handleMouseMove)
      window.removeEventListener('mouseup', handleMouseUp)
    }
  }, [dragState, getElement, updateElement])

  // ── Extract mode: draw-to-create textbox ─────────────────────────────────
  const handleCanvasMouseDown = (e) => {
    if (!extractMode) {
      setSelectedId(null)
      return
    }
    if (!canvasRef.current) return
    e.preventDefault()
    e.stopPropagation()
    const rect = canvasRef.current.getBoundingClientRect()
    const xPct = ((e.clientX - rect.left) / rect.width) * 100
    const yPct = ((e.clientY - rect.top) / rect.height) * 100
    setDrawState({ startXPct: xPct, startYPct: yPct })
    setDrawRect({ xPct, yPct, wPct: 0, hPct: 0 })
  }

  useEffect(() => {
    if (!drawState) return

    const handleMove = (e) => {
      if (!canvasRef.current) return
      const rect = canvasRef.current.getBoundingClientRect()
      const curXPct = Math.max(0, Math.min(100, ((e.clientX - rect.left) / rect.width) * 100))
      const curYPct = Math.max(0, Math.min(100, ((e.clientY - rect.top) / rect.height) * 100))
      const x = Math.min(drawState.startXPct, curXPct)
      const y = Math.min(drawState.startYPct, curYPct)
      const w = Math.abs(curXPct - drawState.startXPct)
      const h = Math.abs(curYPct - drawState.startYPct)
      setDrawRect({ xPct: x, yPct: y, wPct: w, hPct: h })
    }

    const handleUp = (e) => {
      if (!canvasRef.current || !drawRect) {
        setDrawState(null)
        setDrawRect(null)
        return
      }
      const rect = canvasRef.current.getBoundingClientRect()
      const { xPct, yPct, wPct, hPct } = drawRect

      // Convert % → PptxGenJS inches
      const x = parseFloat(((xPct / 100) * 13.333).toFixed(2))
      const y = parseFloat(((yPct / 100) * 7.5).toFixed(2))
      const w = parseFloat(((wPct / 100) * 13.333).toFixed(2))
      const h = parseFloat(((hPct / 100) * 7.5).toFixed(2))

      // Only create if box is big enough (> 0.3" in either dimension)
      if (w > 0.3 && h > 0.1) {
        addTextbox({ x, y, w, h, textDefault: 'Extracted Text' })
      }

      setDrawState(null)
      setDrawRect(null)
    }

    window.addEventListener('mousemove', handleMove)
    window.addEventListener('mouseup', handleUp)
    return () => {
      window.removeEventListener('mousemove', handleMove)
      window.removeEventListener('mouseup', handleUp)
    }
  }, [drawState, drawRect])

  // ── Save handler ─────────────────────────────────────────────────────────
  const handleSave = () => {
    const sortAndMapPlaceholders = (list, prefix) =>
      [...list].sort((a, b) => a.x - b.x).map((p, idx) => ({ ...p, key: `${prefix}_${idx}` }))
    const sortAndMapTextboxes = (list, prefix) =>
      [...list].sort((a, b) => a.x - b.x).map((t, idx) => ({ ...t, key: `${prefix}_${idx}` }))

    const updatedLayout = {
      firstSlideUrl,
      lastSlideUrl,
      masterBgUrl,
      placeholders: sortAndMapPlaceholders(placeholders, 'image'),
      textboxes: sortAndMapTextboxes(textboxes, 'text'),
      firstSlidePlaceholders: sortAndMapPlaceholders(firstSlidePlaceholders, 'first_image'),
      firstSlideTextboxes: sortAndMapTextboxes(firstSlideTextboxes, 'first_text'),
      lastSlidePlaceholders: sortAndMapPlaceholders(lastSlidePlaceholders, 'last_image'),
      lastSlideTextboxes: sortAndMapTextboxes(lastSlideTextboxes, 'last_text'),
    }
    onSave(updatedLayout)
  }

  const handleOverwritePreset = () => {
    if (!activePresetId) {
      alert('No active preset to overwrite. Use "Save as New" in the library above.')
      return
    }
    handleSave()
  }

  const canvasBackgroundUrl =
    editSlideType === 'first' ? firstSlideUrl
    : editSlideType === 'last' ? lastSlideUrl
    : masterBgUrl

  const hasBackground = Boolean(canvasBackgroundUrl)

  return (
    <div className="master-designer">
      {/* ── Preset Library ── */}
      <PresetLibrary
        presets={presets}
        activePresetId={activePresetId}
        onLoad={onLoadPreset}
        onDelete={onDeletePreset}
        onRename={onRenamePreset}
        onExportPreset={onExportPreset}
        onExportAll={onExportAll}
        onImportPresets={onImportPresets}
        onSaveNew={(name) => {
          // Collect current canvas layout and save as new preset
          const sortAndMapPlaceholders = (list, prefix) =>
            [...list].sort((a, b) => a.x - b.x).map((p, idx) => ({ ...p, key: `${prefix}_${idx}` }))
          const sortAndMapTextboxes = (list, prefix) =>
            [...list].sort((a, b) => a.x - b.x).map((t, idx) => ({ ...t, key: `${prefix}_${idx}` }))
          const layout = {
            firstSlideUrl,
            lastSlideUrl,
            masterBgUrl,
            placeholders: sortAndMapPlaceholders(placeholders, 'image'),
            textboxes: sortAndMapTextboxes(textboxes, 'text'),
            firstSlidePlaceholders: sortAndMapPlaceholders(firstSlidePlaceholders, 'first_image'),
            firstSlideTextboxes: sortAndMapTextboxes(firstSlideTextboxes, 'first_text'),
            lastSlidePlaceholders: sortAndMapPlaceholders(lastSlidePlaceholders, 'last_image'),
            lastSlideTextboxes: sortAndMapTextboxes(lastSlideTextboxes, 'last_text'),
          }
          onSavePreset(name, layout)
        }}
      />

      {/* ── Header panel ── */}
      <div className="master-designer__header-panel card">
        <div>
          <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 600 }}>Layout Canvas</h2>
          <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--muted)' }}>
            Add elements, or switch to <strong>Draw Mode</strong> to extract text boxes by drawing on the slide.
          </p>
        </div>
        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
          {/* Extract mode toggle */}
          <button
            type="button"
            className={`ghost${extractMode ? ' is-active' : ''}`}
            style={extractMode ? { borderColor: '#7c3aed', color: '#7c3aed', background: 'rgba(124,58,237,0.08)' } : {}}
            onClick={() => {
              setExtractMode((v) => !v)
              setDrawState(null)
              setDrawRect(null)
            }}
            title="Click and drag on the canvas to draw text box regions"
          >
            {extractMode ? '✏️ Drawing Mode ON' : '✏️ Draw Text Box'}
          </button>
          <button type="button" className="ghost" onClick={addPlaceholder} disabled={extractMode}>
            + Image Box
          </button>
          <button type="button" className="ghost" onClick={() => addTextbox()} disabled={extractMode}>
            + Text Box
          </button>
          {activePresetId && (
            <button type="button" className="ghost" onClick={handleOverwritePreset}>
              ↩ Overwrite Active
            </button>
          )}
          <button type="button" className="button" onClick={handleSave}>
            Save Layout
          </button>
        </div>
      </div>

      {/* ── Slide tab selector ── */}
      <div style={{ display: 'flex', gap: '10px' }}>
        {['first', 'master', 'last'].map((type) => (
          <button
            key={type}
            type="button"
            className={`ghost${editSlideType === type ? ' is-active' : ''}`}
            onClick={() => { setEditSlideType(type); setSelectedId(null) }}
          >
            {type === 'first' ? 'First Slide' : type === 'master' ? 'Master Slide' : 'Last Slide'}
          </button>
        ))}
      </div>

      {/* ── Extract mode hint ── */}
      {extractMode && (
        <div style={{
          padding: '10px 16px',
          borderRadius: '12px',
          background: 'rgba(124,58,237,0.08)',
          border: '1px solid rgba(124,58,237,0.3)',
          fontSize: '13px',
          color: '#5b21b6',
          display: 'flex',
          alignItems: 'center',
          gap: '8px',
        }}>
          <span style={{ fontSize: '18px' }}>✏️</span>
          <span>
            <strong>Draw Mode Active</strong> — Click and drag on the slide canvas to extract a text box at that exact position.
            {!hasBackground && <span style={{ color: '#b91c1c', marginLeft: '8px' }}>⚠ Upload a background image first to see accurate positioning.</span>}
          </span>
        </div>
      )}

      <div className="master-designer__workspace">
        <div className="master-designer__main">
          {/* ── Canvas ── */}
          <div
            ref={canvasRef}
            className="master-designer__canvas"
            style={{
              backgroundImage: canvasBackgroundUrl ? `url(${canvasBackgroundUrl})` : 'none',
              backgroundColor: canvasBackgroundUrl ? 'transparent' : '#f0ece3',
              position: 'relative',
              width: '100%',
              aspectRatio: '16/9',
              border: extractMode ? '2px solid #7c3aed' : '2px dashed #b5b5b5',
              borderRadius: '12px',
              overflow: 'hidden',
              boxShadow: 'inset 0 4px 10px rgba(0,0,0,0.05)',
              cursor: extractMode ? 'crosshair' : 'default',
              userSelect: 'none',
            }}
            onMouseDown={handleCanvasMouseDown}
          >
            {!canvasBackgroundUrl && (
              <div className="master-designer__grid-indicator">
                <p>Upload a background image below to start designing.</p>
                <p style={{ marginTop: '6px', fontSize: '12px' }}>You can also use Draw Mode without a background.</p>
              </div>
            )}

            {/* Draw selection overlay */}
            {drawRect && drawRect.wPct > 0 && drawRect.hPct > 0 && (
              <div
                style={{
                  position: 'absolute',
                  left: `${drawRect.xPct}%`,
                  top: `${drawRect.yPct}%`,
                  width: `${drawRect.wPct}%`,
                  height: `${drawRect.hPct}%`,
                  border: '2px solid #7c3aed',
                  background: 'rgba(124,58,237,0.15)',
                  borderRadius: '4px',
                  pointerEvents: 'none',
                  zIndex: 100,
                  boxShadow: '0 0 0 1px rgba(124,58,237,0.4)',
                }}
              >
                <div style={{
                  position: 'absolute',
                  top: '2px',
                  left: '4px',
                  fontSize: '10px',
                  color: '#7c3aed',
                  fontWeight: 'bold',
                  whiteSpace: 'nowrap',
                  background: 'rgba(255,255,255,0.85)',
                  borderRadius: '3px',
                  padding: '1px 4px',
                }}>
                  Text Box
                </div>
              </div>
            )}

            {/* Placeholders */}
            {activePlaceholders.map((p) => {
              const isSelected = selectedId === p.id
              return (
                <div
                  key={p.id}
                  style={{
                    position: 'absolute',
                    left: `${(p.x / 13.333) * 100}%`,
                    top: `${(p.y / 7.5) * 100}%`,
                    width: `${(p.w / 13.333) * 100}%`,
                    height: `${(p.h / 7.5) * 100}%`,
                    border: isSelected ? '2px solid #0b7a38' : '2px dashed #666',
                    backgroundColor: isSelected ? 'rgba(11, 122, 56, 0.15)' : 'rgba(255, 255, 255, 0.75)',
                    borderRadius: '8px',
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'center',
                    justifyContent: 'center',
                    cursor: extractMode ? 'crosshair' : 'move',
                    padding: '8px',
                    boxSizing: 'border-box',
                    userSelect: 'none',
                    zIndex: isSelected ? 10 : 1,
                    pointerEvents: extractMode ? 'none' : 'auto',
                  }}
                  onMouseDown={(e) => !extractMode && handleMouseDown(e, p.id, 'move')}
                  onClick={(e) => { if (!extractMode) { e.stopPropagation(); setSelectedId(p.id) } }}
                >
                  <span style={{ fontSize: '11px', fontWeight: 'bold', color: '#333', textAlign: 'center' }}>
                    🖼 {p.label}
                  </span>
                  <span style={{ fontSize: '9px', color: '#666', marginTop: '3px' }}>
                    {p.w}" × {p.h}"
                  </span>
                  {!extractMode && (
                    <div
                      style={{
                        position: 'absolute', right: 0, bottom: 0,
                        width: '14px', height: '14px',
                        backgroundColor: isSelected ? '#0b7a38' : '#666',
                        borderRadius: '50%', cursor: 'se-resize',
                        transform: 'translate(4px, 4px)', border: '2px solid #fff',
                      }}
                      onMouseDown={(e) => handleMouseDown(e, p.id, 'resize')}
                      onClick={(e) => e.stopPropagation()}
                    />
                  )}
                </div>
              )
            })}

            {/* Textboxes */}
            {activeTextboxes.map((t) => {
              const isSelected = selectedId === t.id
              const isExtracted = t.textDefault === 'Extracted Text' || t._extracted
              return (
                <div
                  key={t.id}
                  style={{
                    position: 'absolute',
                    left: `${(t.x / 13.333) * 100}%`,
                    top: `${(t.y / 7.5) * 100}%`,
                    width: `${(t.w / 13.333) * 100}%`,
                    height: `${(t.h / 7.5) * 100}%`,
                    border: isSelected
                      ? '2px solid #0b7a38'
                      : isExtracted
                      ? '1.5px dashed #7c3aed'
                      : '1px dashed #e12b2b',
                    backgroundColor: isSelected
                      ? 'rgba(11, 122, 56, 0.1)'
                      : isExtracted
                      ? 'rgba(124, 58, 237, 0.08)'
                      : 'rgba(255, 255, 255, 0.85)',
                    borderRadius: '6px',
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'center',
                    justifyContent: 'center',
                    cursor: extractMode ? 'crosshair' : 'move',
                    padding: '4px',
                    boxSizing: 'border-box',
                    userSelect: 'none',
                    zIndex: isSelected ? 10 : 2,
                    pointerEvents: extractMode ? 'none' : 'auto',
                  }}
                  onMouseDown={(e) => !extractMode && handleMouseDown(e, t.id, 'move')}
                  onClick={(e) => { if (!extractMode) { e.stopPropagation(); setSelectedId(t.id) } }}
                >
                  <span
                    style={{
                      fontFamily: t.fontFace || 'Calibri',
                      fontSize: `${(t.fontSize || 20) * 0.75}px`,
                      color: `#${t.fontColor || '111111'}`,
                      fontWeight: t.bold ? 'bold' : 'normal',
                      textAlign: t.align || 'center',
                      width: '100%',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {t.textDefault || 'Text Field'}
                  </span>
                  {isExtracted && (
                    <span style={{ fontSize: '8px', color: '#7c3aed', marginTop: '2px', opacity: 0.8 }}>
                      ✏️ Extracted
                    </span>
                  )}
                  {!extractMode && (
                    <div
                      style={{
                        position: 'absolute', right: 0, bottom: 0,
                        width: '14px', height: '14px',
                        backgroundColor: isSelected ? '#0b7a38' : isExtracted ? '#7c3aed' : '#e12b2b',
                        borderRadius: '50%', cursor: 'se-resize',
                        transform: 'translate(4px, 4px)', border: '2px solid #fff',
                      }}
                      onMouseDown={(e) => handleMouseDown(e, t.id, 'resize')}
                      onClick={(e) => e.stopPropagation()}
                    />
                  )}
                </div>
              )
            })}
          </div>

          {/* ── Slide asset uploads ── */}
          <div className="master-designer__assets">
            {[
              { label: 'First Slide Image', url: firstSlideUrl, setter: setFirstSlideUrl },
              { label: 'Master Slide Background', url: masterBgUrl, setter: setMasterBgUrl },
              { label: 'Last Slide Image', url: lastSlideUrl, setter: setLastSlideUrl },
            ].map(({ label, url, setter }) => (
              <div key={label} className="asset-upload-card">
                <span className="label">{label}</span>
                {url ? (
                  <div className="asset-preview-wrap">
                    <img src={url} alt={label} />
                    <button type="button" onClick={() => setter('')}>Remove</button>
                  </div>
                ) : (
                  <div className="asset-dropzone">
                    <input type="file" accept="image/*" onChange={(e) => handleFileChange(e, setter)} />
                    <span>Upload PNG / JPEG</span>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>

        {/* ── Sidebar ── */}
        <div className="master-designer__sidebar">
          {selectedElement ? (
            <div className="sidebar-editor">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h3 style={{ margin: 0, fontSize: '15px' }}>
                  {selectedElement.id.startsWith('placeholder_') ? '🖼 Image Box' : '📝 Text Box'}
                </h3>
                <button
                  type="button"
                  style={{
                    border: 'none', background: '#e12b2b', color: '#fff',
                    borderRadius: '6px', padding: '4px 8px', fontSize: '11px', cursor: 'pointer',
                  }}
                  onClick={() => deleteElement(selectedElement.id)}
                >
                  Delete
                </button>
              </div>

              <hr style={{ margin: '10px 0', border: 'none', borderBottom: '1px solid var(--border)' }} />

              <div className="sidebar-group">
                <label className="label">Component ID</label>
                <input type="text" value={selectedElement.key} disabled />
              </div>

              {selectedElement.id.startsWith('placeholder_') ? (
                <div className="sidebar-group">
                  <label className="label">Placeholder Label</label>
                  <input
                    type="text"
                    value={selectedElement.label}
                    onChange={(e) => updateElement(selectedElement.id, { label: e.target.value })}
                  />
                </div>
              ) : (
                <>
                  <div className="sidebar-group">
                    <label className="label">Default Text</label>
                    <input
                      type="text"
                      value={selectedElement.textDefault}
                      onChange={(e) => updateElement(selectedElement.id, { textDefault: e.target.value })}
                    />
                  </div>
                  <div className="sidebar-group">
                    <label className="label">Font Size (pt)</label>
                    <input
                      type="number"
                      value={selectedElement.fontSize}
                      onChange={(e) => updateElement(selectedElement.id, { fontSize: parseInt(e.target.value) || 12 })}
                    />
                  </div>
                  <div className="sidebar-group">
                    <label className="label">Font Color</label>
                    <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                      <input
                        type="color"
                        value={`#${selectedElement.fontColor || '111111'}`}
                        onChange={(e) => updateElement(selectedElement.id, { fontColor: e.target.value.replace('#', '') })}
                        style={{ padding: 0, width: '36px', height: '36px', border: '1px solid var(--border)', borderRadius: '6px', cursor: 'pointer' }}
                      />
                      <input
                        type="text"
                        value={selectedElement.fontColor}
                        onChange={(e) => updateElement(selectedElement.id, { fontColor: e.target.value.replace('#', '') })}
                        style={{ flex: 1, margin: 0 }}
                        placeholder="111111"
                      />
                    </div>
                  </div>
                  <div className="sidebar-group">
                    <label className="label">Font Family</label>
                    <select
                      value={selectedElement.fontFace || 'Calibri'}
                      onChange={(e) => updateElement(selectedElement.id, { fontFace: e.target.value })}
                    >
                      {['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Courier New', 'Impact', 'Trebuchet MS', 'Verdana', 'Garamond', 'Arial Black'].map((f) => (
                        <option key={f} value={f}>{f}</option>
                      ))}
                    </select>
                  </div>
                  <div className="sidebar-group checkbox-row">
                    <input
                      type="checkbox"
                      id="bold-check"
                      checked={!!selectedElement.bold}
                      onChange={(e) => updateElement(selectedElement.id, { bold: e.target.checked })}
                    />
                    <label htmlFor="bold-check">Bold</label>
                  </div>
                  <div className="sidebar-group">
                    <label className="label">Text Align</label>
                    <select
                      value={selectedElement.align || 'center'}
                      onChange={(e) => updateElement(selectedElement.id, { align: e.target.value })}
                    >
                      <option value="left">Left</option>
                      <option value="center">Center</option>
                      <option value="right">Right</option>
                    </select>
                  </div>
                </>
              )}

              <hr style={{ margin: '10px 0', border: 'none', borderBottom: '1px solid var(--border)' }} />

              <h4 style={{ margin: '0 0 10px', fontSize: '12px', color: 'var(--muted)' }}>Position &amp; Size (inches)</h4>
              <div className="sidebar-grid">
                {[
                  { label: 'X', key: 'x', min: 0 },
                  { label: 'Y', key: 'y', min: 0 },
                  { label: 'W', key: 'w', min: 0.2 },
                  { label: 'H', key: 'h', min: 0.1 },
                ].map(({ label, key, min }) => (
                  <div key={key}>
                    <label className="label">{label}</label>
                    <input
                      type="number"
                      step="0.05"
                      value={selectedElement[key]}
                      onChange={(e) => updateElement(selectedElement.id, { [key]: Math.max(min, parseFloat(e.target.value) || min) })}
                    />
                  </div>
                ))}
              </div>
            </div>
          ) : (
            <div className="sidebar-empty">
              {extractMode ? (
                <>
                  <p style={{ fontSize: '28px', margin: '0 0 10px' }}>✏️</p>
                  <p><strong>Draw Mode Active</strong></p>
                  <p>Click and drag on the canvas to place a text box at the exact position on the slide.</p>
                </>
              ) : (
                <p>Click any element on the canvas to configure it, or use <strong>Draw Mode</strong> to extract text boxes.</p>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
