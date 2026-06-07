import React, { useState, useRef, useEffect } from 'react'

export function MasterDesigner({ customLayout, onSave }) {
  const [firstSlideUrl, setFirstSlideUrl] = useState(customLayout?.firstSlideUrl || '')
  const [lastSlideUrl, setLastSlideUrl] = useState(customLayout?.lastSlideUrl || '')
  const [masterBgUrl, setMasterBgUrl] = useState(customLayout?.masterBgUrl || '')

  // Slides components states
  const [placeholders, setPlaceholders] = useState(customLayout?.placeholders || [])
  const [textboxes, setTextboxes] = useState(customLayout?.textboxes || [])

  const [firstSlidePlaceholders, setFirstSlidePlaceholders] = useState(customLayout?.firstSlidePlaceholders || [])
  const [firstSlideTextboxes, setFirstSlideTextboxes] = useState(customLayout?.firstSlideTextboxes || [])

  const [lastSlidePlaceholders, setLastSlidePlaceholders] = useState(customLayout?.lastSlidePlaceholders || [])
  const [lastSlideTextboxes, setLastSlideTextboxes] = useState(customLayout?.lastSlideTextboxes || [])

  // Active slide editor tab
  const [editSlideType, setEditSlideType] = useState('master') // 'first', 'master', 'last'

  const [selectedId, setSelectedId] = useState(null)
  const [dragState, setDragState] = useState(null)

  const canvasRef = useRef(null)

  // Sync state when props change
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
    }
  }, [customLayout])

  // Get active slide lists
  const getActiveList = () => {
    if (editSlideType === 'first') return { placeholders: firstSlidePlaceholders, textboxes: firstSlideTextboxes }
    if (editSlideType === 'last') return { placeholders: lastSlidePlaceholders, textboxes: lastSlideTextboxes }
    return { placeholders, textboxes }
  }

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
    const newPlaceholder = {
      id,
      key: `${editSlideType}_image_${count}`,
      label: `Image Placeholder ${count + 1}`,
      x: 1.0 + (count * 0.5) % 8,
      y: 2.0,
      w: 4.0,
      h: 3.5,
    }
    setActivePlaceholders([...activeP, newPlaceholder])
    setSelectedId(id)
  }

  const addTextbox = () => {
    const id = `textbox_${Date.now()}`
    const { textboxes: activeT } = getActiveList()
    const count = activeT.length
    const newTextbox = {
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
    }
    setActiveTextboxes([...activeT, newTextbox])
    setSelectedId(id)
  }

  const deleteElement = (id) => {
    const { placeholders: activeP, textboxes: activeT } = getActiveList()
    setActivePlaceholders(activeP.filter((p) => p.id !== id))
    setActiveTextboxes(activeT.filter((t) => t.id !== id))
    if (selectedId === id) {
      setSelectedId(null)
    }
  }

  const getElement = (id) => {
    const { placeholders: activeP, textboxes: activeT } = getActiveList()
    return activeP.find((p) => p.id === id) || activeT.find((t) => t.id === id)
  }

  const updateElement = (id, fields) => {
    const { placeholders: activeP, textboxes: activeT } = getActiveList()
    setActivePlaceholders(
      activeP.map((p) => (p.id === id ? { ...p, ...fields } : p))
    )
    setActiveTextboxes(
      activeT.map((t) => (t.id === id ? { ...t, ...fields } : t))
    )
  }

  const { placeholders: activePlaceholders, textboxes: activeTextboxes } = getActiveList()
  const selectedElement = getElement(selectedId)

  // Mouse drag and resize implementation
  const handleMouseDown = (e, id, type) => {
    e.stopPropagation()
    e.preventDefault()
    setSelectedId(id)

    if (!canvasRef.current) return

    const canvasRect = canvasRef.current.getBoundingClientRect()
    const element = getElement(id)
    if (!element) return

    const elLeft = (element.x / 13.333) * canvasRect.width
    const elTop = (element.y / 7.5) * canvasRect.height
    const elWidth = (element.w / 13.333) * canvasRect.width
    const elHeight = (element.h / 7.5) * canvasRect.height

    setDragState({
      id,
      type,
      startX: e.clientX,
      startY: e.clientY,
      startLeft: elLeft,
      startTop: elTop,
      startWidth: elWidth,
      startHeight: elHeight,
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
        const newLeftPx = dragState.startLeft + deltaX
        const newTopPx = dragState.startTop + deltaY

        let newX = (newLeftPx / canvasRect.width) * 13.333
        let newY = (newTopPx / canvasRect.height) * 7.5

        newX = Math.max(0, Math.min(13.333 - element.w, newX))
        newY = Math.max(0, Math.min(7.5 - element.h, newY))

        updateElement(dragState.id, {
          x: parseFloat(newX.toFixed(2)),
          y: parseFloat(newY.toFixed(2)),
        })
      } else if (dragState.type === 'resize') {
        const newWidthPx = dragState.startWidth + deltaX
        const newHeightPx = dragState.startHeight + deltaY

        let newW = (newWidthPx / canvasRect.width) * 13.333
        let newH = (newHeightPx / canvasRect.height) * 7.5

        newW = Math.max(0.5, Math.min(13.333 - element.x, newW))
        newH = Math.max(0.5, Math.min(7.5 - element.y, newH))

        updateElement(dragState.id, {
          w: parseFloat(newW.toFixed(2)),
          h: parseFloat(newH.toFixed(2)),
        })
      }
    }

    const handleMouseUp = () => {
      setDragState(null)
    }

    if (dragState) {
      window.addEventListener('mousemove', handleMouseMove)
      window.addEventListener('mouseup', handleMouseUp)
    }

    return () => {
      window.removeEventListener('mousemove', handleMouseMove)
      window.removeEventListener('mouseup', handleMouseUp)
    }
  }, [dragState])

  const handleSave = () => {
    const sortAndMapPlaceholders = (list, prefix) => [...list]
      .sort((a, b) => a.x - b.x)
      .map((p, idx) => ({ ...p, key: `${prefix}_${idx}` }))

    const sortAndMapTextboxes = (list, prefix) => [...list]
      .sort((a, b) => a.x - b.x)
      .map((t, idx) => ({ ...t, key: `${prefix}_${idx}` }))

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

  const canvasBackgroundUrl = 
    editSlideType === 'first' 
      ? firstSlideUrl 
      : editSlideType === 'last' 
      ? lastSlideUrl 
      : masterBgUrl

  return (
    <div className="master-designer">
      <div className="master-designer__header-panel card">
        <div>
          <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 600 }}>Master Layout Editor</h2>
          <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--muted)' }}>
            Upload slides, add placeholders and text fields, drag and resize them, and save.
          </p>
        </div>
        <div style={{ display: 'flex', gap: '12px' }}>
          <button type="button" className="ghost" onClick={addPlaceholder}>
            + Add Image Box
          </button>
          <button type="button" className="ghost" onClick={addTextbox}>
            + Add Text Box
          </button>
          <button type="button" className="button" onClick={handleSave}>
            Save Layout
          </button>
        </div>
      </div>

      {/* Slide Selector Control */}
      <div style={{ display: 'flex', gap: '10px' }}>
        <button
          type="button"
          className={`ghost${editSlideType === 'first' ? ' is-active' : ''}`}
          onClick={() => { setEditSlideType('first'); setSelectedId(null); }}
        >
          First Slide Layout
        </button>
        <button
          type="button"
          className={`ghost${editSlideType === 'master' ? ' is-active' : ''}`}
          onClick={() => { setEditSlideType('master'); setSelectedId(null); }}
        >
          Master Slide Layout
        </button>
        <button
          type="button"
          className={`ghost${editSlideType === 'last' ? ' is-active' : ''}`}
          onClick={() => { setEditSlideType('last'); setSelectedId(null); }}
        >
          Last Slide Layout
        </button>
      </div>

      <div className="master-designer__workspace">
        <div className="master-designer__main">
          {/* Canvas */}
          <div
            ref={canvasRef}
            className="master-designer__canvas"
            style={{
              backgroundImage: canvasBackgroundUrl ? `url(${canvasBackgroundUrl})` : 'none',
              backgroundColor: canvasBackgroundUrl ? 'transparent' : '#f0ece3',
              position: 'relative',
              width: '100%',
              aspectRatio: '16/9',
              border: '2px dashed #b5b5b5',
              borderRadius: '12px',
              overflow: 'hidden',
              boxShadow: 'inset 0 4px 10px rgba(0,0,0,0.05)',
            }}
            onClick={() => setSelectedId(null)}
          >
            {!canvasBackgroundUrl && (
              <div className="master-designer__grid-indicator">
                <p>Upload a background image below to start designing the {editSlideType} slide components.</p>
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
                    cursor: 'move',
                    padding: '8px',
                    boxSizing: 'border-box',
                    userSelect: 'none',
                    zIndex: isSelected ? 10 : 1,
                  }}
                  onMouseDown={(e) => handleMouseDown(e, p.id, 'move')}
                  onClick={(e) => {
                    e.stopPropagation()
                    setSelectedId(p.id)
                  }}
                >
                  <span style={{ fontSize: '12px', fontWeight: 'bold', color: '#333', textAlign: 'center' }}>
                    {p.label}
                  </span>
                  <span style={{ fontSize: '10px', color: '#666', marginTop: '4px' }}>
                    {p.w}" x {p.h}"
                  </span>

                  <div
                    style={{
                      position: 'absolute',
                      right: '0',
                      bottom: '0',
                      width: '14px',
                      height: '14px',
                      backgroundColor: isSelected ? '#0b7a38' : '#666',
                      borderRadius: '50%',
                      cursor: 'se-resize',
                      transform: 'translate(4px, 4px)',
                      border: '2px solid #fff',
                    }}
                    onMouseDown={(e) => handleMouseDown(e, p.id, 'resize')}
                    onClick={(e) => e.stopPropagation()}
                  />
                </div>
              )
            })}

            {/* Textboxes */}
            {activeTextboxes.map((t) => {
              const isSelected = selectedId === t.id
              return (
                <div
                  key={t.id}
                  style={{
                    position: 'absolute',
                    left: `${(t.x / 13.333) * 100}%`,
                    top: `${(t.y / 7.5) * 100}%`,
                    width: `${(t.w / 13.333) * 100}%`,
                    height: `${(t.h / 7.5) * 100}%`,
                    border: isSelected ? '2px solid #0b7a38' : '1px dashed #e12b2b',
                    backgroundColor: isSelected ? 'rgba(11, 122, 56, 0.1)' : 'rgba(255, 255, 255, 0.85)',
                    borderRadius: '6px',
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'center',
                    justifyContent: 'center',
                    cursor: 'move',
                    padding: '4px',
                    boxSizing: 'border-box',
                    userSelect: 'none',
                    zIndex: isSelected ? 10 : 2,
                  }}
                  onMouseDown={(e) => handleMouseDown(e, t.id, 'move')}
                  onClick={(e) => {
                    e.stopPropagation()
                    setSelectedId(t.id)
                  }}
                >
                  <span
                    style={{
                      fontFamily: t.fontFace || 'Calibri',
                      fontSize: `${t.fontSize * 0.75}px`,
                      color: `#${t.fontColor}`,
                      fontWeight: t.bold ? 'bold' : 'normal',
                      textAlign: t.align,
                      width: '100%',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {t.textDefault || 'Text Field'}
                  </span>

                  <div
                    style={{
                      position: 'absolute',
                      right: '0',
                      bottom: '0',
                      width: '14px',
                      height: '14px',
                      backgroundColor: isSelected ? '#0b7a38' : '#e12b2b',
                      borderRadius: '50%',
                      cursor: 'se-resize',
                      transform: 'translate(4px, 4px)',
                      border: '2px solid #fff',
                    }}
                    onMouseDown={(e) => handleMouseDown(e, t.id, 'resize')}
                    onClick={(e) => e.stopPropagation()}
                  />
                </div>
              )
            })}
          </div>

          {/* Slide Assets Upload Row */}
          <div className="master-designer__assets">
            <div className="asset-upload-card">
              <span className="label">First Slide Image</span>
              {firstSlideUrl ? (
                <div className="asset-preview-wrap">
                  <img src={firstSlideUrl} alt="First slide preview" />
                  <button type="button" onClick={() => setFirstSlideUrl('')}>Remove</button>
                </div>
              ) : (
                <div className="asset-dropzone">
                  <input type="file" accept="image/*" onChange={(e) => handleFileChange(e, setFirstSlideUrl)} />
                  <span>Upload PNG / JPEG</span>
                </div>
              )}
            </div>

            <div className="asset-upload-card">
              <span className="label">Master Slide Background</span>
              {masterBgUrl ? (
                <div className="asset-preview-wrap">
                  <img src={masterBgUrl} alt="Master Bg preview" />
                  <button type="button" onClick={() => setMasterBgUrl('')}>Remove</button>
                </div>
              ) : (
                <div className="asset-dropzone">
                  <input type="file" accept="image/*" onChange={(e) => handleFileChange(e, setMasterBgUrl)} />
                  <span>Upload PNG / JPEG</span>
                </div>
              )}
            </div>

            <div className="asset-upload-card">
              <span className="label">Last Slide Image</span>
              {lastSlideUrl ? (
                <div className="asset-preview-wrap">
                  <img src={lastSlideUrl} alt="Last slide preview" />
                  <button type="button" onClick={() => setLastSlideUrl('')}>Remove</button>
                </div>
              ) : (
                <div className="asset-dropzone">
                  <input type="file" accept="image/*" onChange={(e) => handleFileChange(e, setLastSlideUrl)} />
                  <span>Upload PNG / JPEG</span>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Sidebar settings */}
        <div className="master-designer__sidebar">
          {selectedElement ? (
            <div className="sidebar-editor">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h3 style={{ margin: 0, fontSize: '15px' }}>Modify Component</h3>
                <button
                  type="button"
                  style={{
                    border: 'none',
                    background: '#e12b2b',
                    color: '#fff',
                    borderRadius: '6px',
                    padding: '4px 8px',
                    fontSize: '11px',
                    cursor: 'pointer',
                  }}
                  onClick={() => deleteElement(selectedElement.id)}
                >
                  Delete
                </button>
              </div>

              <hr style={{ margin: '12px 0', border: 'none', borderBottom: '1px solid var(--border)' }} />

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
                    <label className="label">Font Size (px)</label>
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
                        style={{ padding: '0', width: '36px', height: '36px', border: '1px solid var(--border)', borderRadius: '6px', cursor: 'pointer' }}
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
                      <option value="Calibri">Calibri</option>
                      <option value="Arial">Arial</option>
                      <option value="Times New Roman">Times New Roman</option>
                      <option value="Georgia">Georgia</option>
                      <option value="Courier New">Courier New</option>
                      <option value="Impact">Impact</option>
                      <option value="Trebuchet MS">Trebuchet MS</option>
                      <option value="Verdana">Verdana</option>
                      <option value="Garamond">Garamond</option>
                      <option value="Arial Black">Arial Black</option>
                    </select>
                  </div>
                  <div className="sidebar-group checkbox-row">
                    <input
                      type="checkbox"
                      id="bold-check"
                      checked={selectedElement.bold}
                      onChange={(e) => updateElement(selectedElement.id, { bold: e.target.checked })}
                    />
                    <label htmlFor="bold-check">Bold</label>
                  </div>
                  <div className="sidebar-group">
                    <label className="label">Text Align</label>
                    <select
                      value={selectedElement.align}
                      onChange={(e) => updateElement(selectedElement.id, { align: e.target.value })}
                    >
                      <option value="left">Left</option>
                      <option value="center">Center</option>
                      <option value="right">Right</option>
                    </select>
                  </div>
                </>
              )}

              <hr style={{ margin: '12px 0', border: 'none', borderBottom: '1px solid var(--border)' }} />

              <h4 style={{ margin: '0 0 10px', fontSize: '12px', color: 'var(--muted)' }}>Dimensions (Inches)</h4>
              <div className="sidebar-grid">
                <div>
                  <label className="label">X</label>
                  <input
                    type="number"
                    step="0.05"
                    value={selectedElement.x}
                    onChange={(e) => updateElement(selectedElement.id, { x: Math.max(0, parseFloat(e.target.value) || 0) })}
                  />
                </div>
                <div>
                  <label className="label">Y</label>
                  <input
                    type="number"
                    step="0.05"
                    value={selectedElement.y}
                    onChange={(e) => updateElement(selectedElement.id, { y: Math.max(0, parseFloat(e.target.value) || 0) })}
                  />
                </div>
                <div>
                  <label className="label">W</label>
                  <input
                    type="number"
                    step="0.05"
                    value={selectedElement.w}
                    onChange={(e) => updateElement(selectedElement.id, { w: Math.max(0.5, parseFloat(e.target.value) || 0.5) })}
                  />
                </div>
                <div>
                  <label className="label">H</label>
                  <input
                    type="number"
                    step="0.05"
                    value={selectedElement.h}
                    onChange={(e) => updateElement(selectedElement.id, { h: Math.max(0.5, parseFloat(e.target.value) || 0.5) })}
                  />
                </div>
              </div>
            </div>
          ) : (
            <div className="sidebar-empty">
              <p>Click on any component on the canvas to configure its properties.</p>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
