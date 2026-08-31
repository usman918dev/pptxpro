import React, { useState, useRef, useEffect } from 'react'
import { parsePptxForEditing, exportEditedPptx } from './report/pptxEditorUtils'
import { loadPptxEditorState, savePptxEditorState, clearPptxEditorState } from './utils/storage'
import './PptxEditor.css'

export function PptxEditor() {
  const [file, setFile] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [parsedData, setParsedData] = useState(null)
  const [activeSlideIndex, setActiveSlideIndex] = useState(0)
  const [selectedElementId, setSelectedElementId] = useState(null)
  const [isExporting, setIsExporting] = useState(false)
  const [isDraggingUpload, setIsDraggingUpload] = useState(false)

  // Drag-and-drop image replace state
  const [dragOverElemId, setDragOverElemId] = useState(null)
  const imageInputRef = useRef(null)
  const [replacingElemId, setReplacingElemId] = useState(null)
  const fileBufferRef = useRef(null)
  const fileHandleRef = useRef(null)
  const isHydratedRef = useRef(false)
  const [toastMessage, setToastMessage] = useState('')
  const [deferredPrompt, setDeferredPrompt] = useState(null)

  // Listen for PWA desktop install prompt
  useEffect(() => {
    const handleBeforeInstall = (e) => {
      e.preventDefault()
      setDeferredPrompt(e)
    }
    window.addEventListener('beforeinstallprompt', handleBeforeInstall)
    return () => window.removeEventListener('beforeinstallprompt', handleBeforeInstall)
  }, [])

  const showToast = (msg) => {
    setToastMessage(msg)
    setTimeout(() => setToastMessage(''), 4000)
  }

  // Hydrate saved PPTX Editor state on mount
  useEffect(() => {
    const hydrate = async () => {
      try {
        const saved = await loadPptxEditorState()
        if (saved && saved.parsedData) {
          if (saved.file) {
            setFile(saved.file)
            fileBufferRef.current = await saved.file.arrayBuffer()
          }
          setParsedData(saved.parsedData)
          setActiveSlideIndex(saved.activeSlideIndex || 0)
        }
      } catch (err) {
        console.warn('Failed to restore PPTX Editor state:', err)
      } finally {
        isHydratedRef.current = true
      }
    }
    hydrate()
  }, [])

  // Auto-save changes to IndexedDB whenever parsedData or activeSlideIndex changes
  useEffect(() => {
    if (!isHydratedRef.current || !parsedData) return
    savePptxEditorState({
      fileBuffer: fileBufferRef.current,
      filename: file?.name || parsedData.filename,
      parsedData,
      activeSlideIndex,
    })
  }, [parsedData, activeSlideIndex, file])

  const handleFileUpload = async (uploadedFile, fileHandle = null) => {
    if (!uploadedFile) return
    if (!uploadedFile.name.toLowerCase().endsWith('.pptx')) {
      setError('Please select a valid PowerPoint (.pptx) file.')
      return
    }

    setLoading(true)
    setError('')
    setFile(uploadedFile)
    if (fileHandle) fileHandleRef.current = fileHandle

    try {
      const buffer = await uploadedFile.arrayBuffer()
      fileBufferRef.current = buffer
      const data = await parsePptxForEditing(uploadedFile)
      setParsedData(data)
      setActiveSlideIndex(0)
      setSelectedElementId(null)
      savePptxEditorState({
        fileBuffer: buffer,
        filename: uploadedFile.name,
        parsedData: data,
        activeSlideIndex: 0,
      })
    } catch (err) {
      console.error(err)
      setError(err.message || 'Failed to parse PPTX file structure.')
      setParsedData(null)
    } finally {
      setLoading(false)
    }
  }

  // Open PPTX using Native File Picker (retains file handle for 1-click overwrite!)
  const handleNativeOpen = async () => {
    if (typeof window.showOpenFilePicker === 'function') {
      try {
        const [handle] = await window.showOpenFilePicker({
          types: [
            {
              description: 'PowerPoint Presentation',
              accept: {
                'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'],
              },
            },
          ],
        })
        const selectedFile = await handle.getFile()
        handleFileUpload(selectedFile, handle)
      } catch (err) {
        if (err.name !== 'AbortError') console.error('Native open failed:', err)
      }
    } else {
      // Fallback input trigger
      document.getElementById('pptx-file-input-fallback')?.click()
    }
  }

  // Save PPTX directly to original file on disk (1-click overwrite!)
  const handleSaveDirect = async () => {
    const exportFile =
      file ||
      (fileBufferRef.current
        ? new File([fileBufferRef.current], parsedData?.filename || 'Presentation.pptx', {
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          })
        : null)

    if (!exportFile || !parsedData) return

    setIsExporting(true)
    try {
      const res = await exportEditedPptx(exportFile, parsedData.slides, {
        fileHandle: fileHandleRef.current,
        saveAs: false,
        download: !fileHandleRef.current,
        outputFileName: exportFile.name,
      })
      if (res?.savedDirectly) {
        showToast(`✓ Saved directly to "${res.fileName}" on disk!`)
      } else {
        showToast(`✓ Exported "${exportFile.name}"!`)
      }
    } catch (err) {
      console.error('Save error:', err)
      alert('Error saving PPTX: ' + err.message)
    } finally {
      setIsExporting(false)
    }
  }

  // Save As (pick new file location on disk)
  const handleSaveAs = async () => {
    const exportFile =
      file ||
      (fileBufferRef.current
        ? new File([fileBufferRef.current], parsedData?.filename || 'Presentation.pptx', {
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          })
        : null)

    if (!exportFile || !parsedData) return

    setIsExporting(true)
    try {
      const res = await exportEditedPptx(exportFile, parsedData.slides, {
        saveAs: true,
        download: false,
        outputFileName: exportFile.name,
      })
      if (res?.savedDirectly) {
        if (res.fileHandle) fileHandleRef.current = res.fileHandle
        showToast(`✓ Saved to "${res.fileName}"!`)
      }
    } catch (err) {
      console.error('Save As error:', err)
      alert('Error saving PPTX: ' + err.message)
    } finally {
      setIsExporting(false)
    }
  }

  const handleInstallApp = () => {
    if (deferredPrompt) {
      deferredPrompt.prompt()
      deferredPrompt.userChoice.then(() => setDeferredPrompt(null))
    }
  }

  const handleUploadDrop = (e) => {
    e.preventDefault()
    setIsDraggingUpload(false)
    const droppedFile = e.dataTransfer.files?.[0]
    if (droppedFile) {
      handleFileUpload(droppedFile)
    }
  }

  const activeSlide = parsedData?.slides[activeSlideIndex]

  // Element Updates
  const updateElement = (elemId, updates) => {
    setParsedData((prev) => {
      if (!prev) return prev
      const newSlides = [...prev.slides]
      const slide = { ...newSlides[activeSlideIndex] }
      slide.elements = slide.elements.map((elem) => {
        if (elem.id === elemId) {
          return { ...elem, ...updates }
        }
        return elem
      })
      newSlides[activeSlideIndex] = slide
      return { ...prev, slides: newSlides }
    })
  }

  // Handle image replacement via file drop directly onto canvas element
  const handleElementImageDrop = (e, elemId) => {
    e.preventDefault()
    e.stopPropagation()
    setDragOverElemId(null)
    const droppedFile = e.dataTransfer.files?.[0]
    if (!droppedFile || !droppedFile.type.startsWith('image/')) return

    const reader = new FileReader()
    reader.onload = (evt) => {
      updateElement(elemId, { dataUrl: evt.target.result })
    }
    reader.readAsDataURL(droppedFile)
  }

  // Trigger file browse for image element replacement
  const triggerImageReplacement = (elemId) => {
    setReplacingElemId(elemId)
    if (imageInputRef.current) {
      imageInputRef.current.value = ''
      imageInputRef.current.click()
    }
  }

  const handleImageFileChange = (e) => {
    const selectedFile = e.target.files?.[0]
    if (!selectedFile || !replacingElemId) return

    const reader = new FileReader()
    reader.onload = (evt) => {
      updateElement(replacingElemId, { dataUrl: evt.target.result })
      setReplacingElemId(null)
    }
    reader.readAsDataURL(selectedFile)
  }

  // Table Cell Editing
  const handleTableCellChange = (elemId, rowIndex, colIndex, newText) => {
    setParsedData((prev) => {
      if (!prev) return prev
      const newSlides = [...prev.slides]
      const slide = { ...newSlides[activeSlideIndex] }
      slide.elements = slide.elements.map((elem) => {
        if (elem.id === elemId && elem.type === 'table') {
          const newRows = elem.rows.map((row, rIdx) => {
            if (rIdx === rowIndex) {
              return row.map((cell, cIdx) => (cIdx === colIndex ? { ...cell, text: newText } : cell))
            }
            return row
          })
          return { ...elem, rows: newRows }
        }
        return elem
      })
      newSlides[activeSlideIndex] = slide
      return { ...prev, slides: newSlides }
    })
  }

  // Element Insertion Handlers
  const handleAddTextBox = () => {
    if (!parsedData) return
    const newId = `inserted_text_${Date.now()}`
    const newElem = {
      id: newId,
      type: 'text',
      tagName: '<p:sp> Text Box',
      text: 'New Text Box',
      originalText: '',
      fontFace: 'Calibri',
      fontSizePct: 2.5,
      bold: false,
      color: '111111',
      align: 'left',
      xPct: 35,
      yPct: 40,
      wPct: 30,
      hPct: 15,
      xEmu: 4267200,
      yEmu: 2743200,
    }
    setParsedData((prev) => {
      const newSlides = [...prev.slides]
      const slide = { ...newSlides[activeSlideIndex] }
      slide.elements = [...slide.elements, newElem]
      newSlides[activeSlideIndex] = slide
      return { ...prev, slides: newSlides }
    })
    setSelectedElementId(newId)
  }

  const handleAddImageBox = () => {
    if (!parsedData) return
    const newId = `inserted_pic_${Date.now()}`
    const newElem = {
      id: newId,
      type: 'image',
      tagName: '<p:pic> Image Box',
      dataUrl: '',
      originalDataUrl: '',
      xPct: 35,
      yPct: 30,
      wPct: 30,
      hPct: 35,
      xEmu: 4267200,
      yEmu: 2057400,
    }
    setParsedData((prev) => {
      const newSlides = [...prev.slides]
      const slide = { ...newSlides[activeSlideIndex] }
      slide.elements = [...slide.elements, newElem]
      newSlides[activeSlideIndex] = slide
      return { ...prev, slides: newSlides }
    })
    setSelectedElementId(newId)
  }

  const handleAddTable = () => {
    if (!parsedData) return
    const newId = `inserted_tbl_${Date.now()}`
    const newElem = {
      id: newId,
      type: 'table',
      tagName: '<a:tbl> Table',
      xPct: 25,
      yPct: 30,
      wPct: 50,
      hPct: 35,
      xEmu: 3048000,
      yEmu: 2057400,
      rows: [
        [
          { text: 'Header 1', bold: true, fillColor: '4472C4', color: 'FFFFFF', align: 'center', borderColor: 'FFFFFF' },
          { text: 'Header 2', bold: true, fillColor: '4472C4', color: 'FFFFFF', align: 'center', borderColor: 'FFFFFF' },
        ],
        [
          { text: 'Row 1 Data', bold: false, fillColor: 'F2F2F2', color: '111111', align: 'left', borderColor: 'CCCCCC' },
          { text: 'Row 1 Value', bold: false, fillColor: 'F2F2F2', color: '111111', align: 'left', borderColor: 'CCCCCC' },
        ],
      ],
    }
    setParsedData((prev) => {
      const newSlides = [...prev.slides]
      const slide = { ...newSlides[activeSlideIndex] }
      slide.elements = [...slide.elements, newElem]
      newSlides[activeSlideIndex] = slide
      return { ...prev, slides: newSlides }
    })
    setSelectedElementId(newId)
  }

  const handleAddSlide = () => {
    if (!parsedData) return
    const newSlideNum = parsedData.slides.length + 1
    const newSlide = {
      id: `slide_${newSlideNum}_${Date.now()}`,
      slideNumber: newSlideNum,
      title: `Slide ${newSlideNum}`,
      xmlPath: `ppt/slides/slide${newSlideNum}.xml`,
      relsPath: `ppt/slides/_rels/slide${newSlideNum}.xml.rels`,
      elements: [
        {
          id: `slide_${newSlideNum}_title_${Date.now()}`,
          type: 'text',
          tagName: '<p:sp> Title',
          text: `Slide ${newSlideNum} Title`,
          originalText: '',
          fontFace: 'Calibri',
          fontSizePct: 3.5,
          bold: true,
          color: '111111',
          align: 'center',
          isTitle: true,
          xPct: 10,
          yPct: 10,
          wPct: 80,
          hPct: 15,
          xEmu: 1219200,
          yEmu: 685800,
        },
      ],
      backgroundDataUrl: '',
      rawXml: '',
    }
    setParsedData((prev) => {
      const slides = [...prev.slides, newSlide]
      return { ...prev, slides }
    })
    setActiveSlideIndex(parsedData.slides.length)
  }

  // Slide Deck Management
  const moveSlide = (fromIndex, toIndex) => {
    if (toIndex < 0 || toIndex >= parsedData.slides.length) return
    setParsedData((prev) => {
      const slides = [...prev.slides]
      const [moved] = slides.splice(fromIndex, 1)
      slides.splice(toIndex, 0, moved)
      const renumbered = slides.map((s, idx) => ({ ...s, slideNumber: idx + 1 }))
      return { ...prev, slides: renumbered }
    })
    setActiveSlideIndex(toIndex)
  }

  const duplicateSlide = (index) => {
    setParsedData((prev) => {
      const slides = [...prev.slides]
      const orig = slides[index]
      const copy = {
        ...orig,
        id: `slide_${Date.now()}_copy`,
        slideNumber: index + 2,
        title: `${orig.title} (Copy)`,
        elements: orig.elements.map((e) => ({ ...e, id: `${e.id}_copy_${Date.now()}` })),
      }
      slides.splice(index + 1, 0, copy)
      const renumbered = slides.map((s, idx) => ({ ...s, slideNumber: idx + 1 }))
      return { ...prev, slides: renumbered }
    })
    setActiveSlideIndex(index + 1)
  }

  const deleteSlide = (index) => {
    if (parsedData.slides.length <= 1) {
      alert('Cannot delete the last slide in the presentation.')
      return
    }
    setParsedData((prev) => {
      const slides = prev.slides.filter((_, i) => i !== index)
      const renumbered = slides.map((s, idx) => ({ ...s, slideNumber: idx + 1 }))
      return { ...prev, slides: renumbered }
    })
    setActiveSlideIndex((prev) => Math.max(0, Math.min(prev, parsedData.slides.length - 2)))
  }

  const deleteSelectedElement = () => {
    if (!selectedElementId || !activeSlide) return
    setParsedData((prev) => {
      const newSlides = [...prev.slides]
      const slide = { ...newSlides[activeSlideIndex] }
      slide.elements = slide.elements.filter((e) => e.id !== selectedElementId)
      newSlides[activeSlideIndex] = slide
      return { ...prev, slides: newSlides }
    })
    setSelectedElementId(null)
  }

  // Handle PPTX export
  const handleExport = async () => {
    const exportFile =
      file ||
      (fileBufferRef.current
        ? new File([fileBufferRef.current], parsedData?.filename || 'Presentation.pptx', {
            type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          })
        : null)

    if (!exportFile || !parsedData) {
      alert('No presentation loaded to export.')
      return
    }

    setIsExporting(true)
    try {
      await exportEditedPptx(exportFile, parsedData.slides, `Edited_${exportFile.name}`)
    } catch (err) {
      console.error('Export error:', err)
      alert('Error exporting PPTX: ' + err.message)
    } finally {
      setIsExporting(false)
    }
  }

  const selectedElem = activeSlide?.elements.find((e) => e.id === selectedElementId)

  return (
    <div className="pptx-studio">
      <input
        type="file"
        ref={imageInputRef}
        onChange={handleImageFileChange}
        accept="image/png, image/jpeg, image/webp, image/gif, image/svg+xml"
        style={{ display: 'none' }}
      />

      {/* ── Ribbon Toolbar Header ── */}
      <header className="pptx-ribbon">
        <div className="pptx-ribbon__brand">
          <div className="pptx-ribbon__logo">📊</div>
          <div>
            <h1 className="pptx-ribbon__title">PowerPoint Studio Editor</h1>
            <span className="pptx-ribbon__subtitle">
              {parsedData ? parsedData.filename : 'No presentation loaded'}
            </span>
          </div>
        </div>

        {parsedData && (
          <div className="pptx-ribbon__actions">
            <div className="pptx-ribbon__group">
              <button
                type="button"
                className="pptx-ribbon__btn"
                onClick={handleAddSlide}
                title="Add a new blank slide to the presentation"
              >
                <span className="pptx-ribbon__icon">➕</span>
                <span>New Slide</span>
              </button>
            </div>

            <div className="pptx-ribbon__divider" />

            <div className="pptx-ribbon__group">
              <button
                type="button"
                className="pptx-ribbon__btn"
                onClick={handleAddTextBox}
                title="Insert a new text box onto the current slide"
              >
                <span className="pptx-ribbon__icon">📝</span>
                <span>Add Text Box</span>
              </button>

              <button
                type="button"
                className="pptx-ribbon__btn"
                onClick={handleAddImageBox}
                title="Insert a new image slot onto the current slide"
              >
                <span className="pptx-ribbon__icon">🖼️</span>
                <span>Add Image Slot</span>
              </button>

              <button
                type="button"
                className="pptx-ribbon__btn"
                onClick={handleAddTable}
                title="Insert a new table onto the current slide"
              >
                <span className="pptx-ribbon__icon">📊</span>
                <span>Add Table</span>
              </button>
            </div>

            <div className="pptx-ribbon__divider" />

            <div className="pptx-ribbon__group">
              <button
                type="button"
                className="pptx-ribbon__btn pptx-ribbon__btn--ghost"
                onClick={async () => {
                  if (confirm('Open a new presentation? Unsaved changes will be lost.')) {
                    await clearPptxEditorState()
                    setParsedData(null)
                    setFile(null)
                    fileBufferRef.current = null
                    fileHandleRef.current = null
                  }
                }}
              >
                <span className="pptx-ribbon__icon">📂</span>
                <span>Open PPTX</span>
              </button>

              <button
                type="button"
                className="pptx-ribbon__btn pptx-ribbon__btn--primary"
                onClick={handleSaveDirect}
                disabled={isExporting}
                title="Overwrite and save directly to original file on disk without downloading again"
              >
                <span className="pptx-ribbon__icon">{isExporting ? '⏳' : '💾'}</span>
                <span>{isExporting ? 'Saving...' : 'Save File'}</span>
              </button>

              <button
                type="button"
                className="pptx-ribbon__btn pptx-ribbon__btn--ghost"
                onClick={handleSaveAs}
                disabled={isExporting}
                title="Choose a new file location on disk to save"
              >
                <span className="pptx-ribbon__icon">💾</span>
                <span>Save As...</span>
              </button>

              <button
                type="button"
                className="pptx-ribbon__btn pptx-ribbon__btn--ghost"
                onClick={handleExport}
                disabled={isExporting}
                title="Download a separate copy file via browser"
              >
                <span className="pptx-ribbon__icon">⬇️</span>
                <span>Download Copy</span>
              </button>

              {deferredPrompt && (
                <button
                  type="button"
                  className="pptx-ribbon__btn pptx-ribbon__btn--primary"
                  onClick={handleInstallApp}
                  title="Install PPTXPro as a desktop application on your PC"
                >
                  <span className="pptx-ribbon__icon">💻</span>
                  <span>Install Desktop App</span>
                </button>
              )}
            </div>
          </div>
        )}
      </header>

      {/* Toast Save Banner */}
      {toastMessage && (
        <div className="pptx-editor__toast">
          <span>{toastMessage}</span>
        </div>
      )}

      {/* Error Alert */}
      {error && (
        <div className="pptx-editor__alert pptx-editor__alert--error">
          <span>⚠️ {error}</span>
          <button type="button" onClick={() => setError('')}>✕</button>
        </div>
      )}

      {/* ── 1. Upload View ── */}
      {!parsedData && !loading && (
        <div className="pptx-editor__upload-wrapper">
          <div
            className={`pptx-editor__dropzone${isDraggingUpload ? ' is-dragging' : ''}`}
            onDrop={handleUploadDrop}
            onDragOver={(e) => { e.preventDefault(); setIsDraggingUpload(true) }}
            onDragLeave={() => setIsDraggingUpload(false)}
          >
            <div className="pptx-editor__dropzone-icon">📊</div>
            <h3>Drag & Drop your PowerPoint (.pptx) file here</h3>
            <p>Full web presentation editor with interactive slide canvas, direct text/table editing, and drag-drop image replacement</p>

            <button type="button" className="pptx-editor__upload-btn" onClick={handleNativeOpen}>
              <span>Choose PPTX File</span>
            </button>
            <input
              id="pptx-file-input-fallback"
              type="file"
              accept=".pptx"
              style={{ display: 'none' }}
              onChange={(e) => handleFileUpload(e.target.files?.[0])}
            />
          </div>
        </div>
      )}

      {/* Loading Spinner */}
      {loading && (
        <div className="pptx-editor__loading">
          <div className="pptx-editor__spinner" />
          <p>Extracting slide elements, shapes, images, and tables...</p>
        </div>
      )}

      {/* ── 2. Studio Workbench ── */}
      {parsedData && (
        <div className="pptx-workbench">
          {/* Left Panel: Slide Deck Sidebar */}
          <aside className="pptx-deck">
            <div className="pptx-deck__header">
              <h3>Slides ({parsedData.slides.length})</h3>
              <button
                type="button"
                className="pptx-mini-btn"
                onClick={handleAddSlide}
                title="Add new slide"
              >
                ➕ New
              </button>
            </div>

            <div className="pptx-deck__list">
              {parsedData.slides.map((slide, idx) => {
                const isActive = idx === activeSlideIndex
                const textCount = slide.elements.filter((e) => e.type === 'text').length
                const imgCount = slide.elements.filter((e) => e.type === 'image').length
                const tblCount = slide.elements.filter((e) => e.type === 'table').length

                return (
                  <div
                    key={slide.id}
                    className={`pptx-deck__card${isActive ? ' is-active' : ''}`}
                    onClick={() => {
                      setActiveSlideIndex(idx)
                      setSelectedElementId(null)
                    }}
                  >
                    <div className="pptx-deck__num">{slide.slideNumber}</div>
                    <div className="pptx-deck__thumb">
                      {slide.backgroundDataUrl ? (
                        <img src={slide.backgroundDataUrl} alt="Slide thumb" className="pptx-deck__thumb-bg" />
                      ) : (
                        <div className="pptx-deck__thumb-empty" />
                      )}
                      <span className="pptx-deck__title">{slide.title}</span>
                    </div>

                    <div className="pptx-deck__badges">
                      {textCount > 0 && <span className="pptx-chip">📝 {textCount}</span>}
                      {imgCount > 0 && <span className="pptx-chip">🖼️ {imgCount}</span>}
                      {tblCount > 0 && <span className="pptx-chip">📊 {tblCount}</span>}
                    </div>

                    <div className="pptx-deck__controls" onClick={(e) => e.stopPropagation()}>
                      <button
                        type="button"
                        className="pptx-icon-btn"
                        disabled={idx === 0}
                        onClick={() => moveSlide(idx, idx - 1)}
                        title="Move Up"
                      >
                        ▲
                      </button>
                      <button
                        type="button"
                        className="pptx-icon-btn"
                        disabled={idx === parsedData.slides.length - 1}
                        onClick={() => moveSlide(idx, idx + 1)}
                        title="Move Down"
                      >
                        ▼
                      </button>
                      <button
                        type="button"
                        className="pptx-icon-btn"
                        onClick={() => duplicateSlide(idx)}
                        title="Duplicate"
                      >
                        📋
                      </button>
                      <button
                        type="button"
                        className="pptx-icon-btn pptx-icon-btn--danger"
                        onClick={() => deleteSlide(idx)}
                        title="Delete"
                      >
                        🗑️
                      </button>
                    </div>
                  </div>
                )
              })}
            </div>
          </aside>

          {/* Center Stage: Interactive 16:9 Slide Canvas */}
          <main className="pptx-stage">
            <div className="pptx-stage__bar">
              <div className="pptx-stage__info">
                <span>Slide {activeSlide?.slideNumber} of {parsedData.slides.length}</span>
                <span className="pptx-stage__divider">•</span>
                <span>{activeSlide?.title}</span>
              </div>
              <span className="pptx-stage__hint">
                💡 Double-click text to edit • Drag & drop images onto slots to replace
              </span>
            </div>

            <div className="pptx-stage__viewport">
              {activeSlide ? (
                <div
                  className="pptx-canvas"
                  style={{
                    backgroundImage: activeSlide.backgroundDataUrl ? `url(${activeSlide.backgroundDataUrl})` : undefined,
                  }}
                  onClick={() => setSelectedElementId(null)}
                >
                  {/* Render Elements */}
                  {activeSlide.elements.map((elem) => {
                    const isSelected = selectedElementId === elem.id
                    const isDragOver = dragOverElemId === elem.id

                    // 1. Text Element
                    if (elem.type === 'text') {
                      return (
                        <div
                          key={elem.id}
                          className={`pptx-canvas__elem pptx-canvas__text${isSelected ? ' is-selected' : ''}`}
                          style={{
                            left: `${elem.xPct}%`,
                            top: `${elem.yPct}%`,
                            width: `${elem.wPct}%`,
                            height: `${elem.hPct}%`,
                            fontSize: `${elem.fontSizePct || 1.8}cqw`,
                            fontFamily: elem.fontFace || 'Calibri',
                            color: elem.color ? `#${elem.color}` : 'inherit',
                            fontWeight: elem.bold ? 'bold' : 'normal',
                            textAlign: elem.align || 'left',
                          }}
                          onClick={(e) => {
                            e.stopPropagation()
                            setSelectedElementId(elem.id)
                          }}
                        >
                          <textarea
                            className="pptx-canvas__text-input"
                            value={elem.text}
                            onChange={(e) => updateElement(elem.id, { text: e.target.value })}
                            placeholder="Enter text..."
                            rows={1}
                            style={{
                              fontSize: 'inherit',
                              fontFamily: 'inherit',
                              color: 'inherit',
                              fontWeight: 'inherit',
                              textAlign: 'inherit',
                            }}
                          />
                        </div>
                      )
                    }

                    // 2. Image Element
                    if (elem.type === 'image') {
                      return (
                        <div
                          key={elem.id}
                          className={`pptx-canvas__elem pptx-canvas__img${isSelected ? ' is-selected' : ''}${isDragOver ? ' is-drag-over' : ''}`}
                          style={{
                            left: `${elem.xPct}%`,
                            top: `${elem.yPct}%`,
                            width: `${elem.wPct}%`,
                            height: `${elem.hPct}%`,
                          }}
                          onClick={(e) => {
                            e.stopPropagation()
                            setSelectedElementId(elem.id)
                          }}
                          onDragOver={(e) => {
                            e.preventDefault()
                            e.stopPropagation()
                            setDragOverElemId(elem.id)
                          }}
                          onDragLeave={() => setDragOverElemId(null)}
                          onDrop={(e) => handleElementImageDrop(e, elem.id)}
                        >
                          {elem.dataUrl ? (
                            <img src={elem.dataUrl} alt="Slide Asset" className="pptx-canvas__img-asset" />
                          ) : (
                            <div className="pptx-canvas__img-placeholder">
                              <span>🖼️ Image Slot</span>
                              <small>Drag & drop or click replace</small>
                            </div>
                          )}

                          <button
                            type="button"
                            className="pptx-canvas__img-btn"
                            onClick={(e) => {
                              e.stopPropagation()
                              triggerImageReplacement(elem.id)
                            }}
                            title="Replace Image"
                          >
                            📷 Replace
                          </button>
                        </div>
                      )
                    }

                    // 3. Table Element
                    if (elem.type === 'table') {
                      return (
                        <div
                          key={elem.id}
                          className={`pptx-canvas__elem pptx-canvas__table${isSelected ? ' is-selected' : ''}`}
                          style={{
                            left: `${elem.xPct}%`,
                            top: `${elem.yPct}%`,
                            width: `${elem.wPct}%`,
                            height: `${elem.hPct}%`,
                          }}
                          onClick={(e) => {
                            e.stopPropagation()
                            setSelectedElementId(elem.id)
                          }}
                        >
                          <table className="pptx-canvas__tbl-grid">
                            <tbody>
                              {(elem.rows || []).map((row, rIdx) => (
                                <tr key={`r_${rIdx}`}>
                                  {row.map((cell, cIdx) => (
                                    <td
                                      key={`c_${cIdx}`}
                                      style={{
                                        backgroundColor: cell.fillColor ? `#${cell.fillColor}` : undefined,
                                        borderColor: cell.borderColor ? `#${cell.borderColor}` : '#cccccc',
                                        color: cell.color ? `#${cell.color}` : '#111111',
                                        fontWeight: cell.bold ? 'bold' : 'normal',
                                        fontSize: `${cell.fontSizePct || 1.4}cqw`,
                                        fontFamily: cell.fontFace || 'Calibri',
                                        textAlign: cell.align || 'left',
                                      }}
                                    >
                                      <input
                                        type="text"
                                        className="pptx-canvas__tbl-cell-input"
                                        value={cell.text || ''}
                                        onChange={(e) =>
                                          handleTableCellChange(elem.id, rIdx, cIdx, e.target.value)
                                        }
                                        placeholder="..."
                                        style={{
                                          fontSize: 'inherit',
                                          fontFamily: 'inherit',
                                          color: 'inherit',
                                          fontWeight: 'inherit',
                                          textAlign: 'inherit',
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

                    return null
                  })}
                </div>
              ) : (
                <div className="pptx-stage__empty">Select a slide to edit</div>
              )}
            </div>
          </main>

          {/* Right Panel: Property Inspector */}
          <aside className="pptx-inspector">
            <div className="pptx-inspector__header">
              <h3>Element Inspector</h3>
              {selectedElem && (
                <button
                  type="button"
                  className="pptx-mini-btn pptx-mini-btn--danger"
                  onClick={deleteSelectedElement}
                  title="Delete Selected Element"
                >
                  🗑️ Delete
                </button>
              )}
            </div>

            {selectedElem ? (
              <div className="pptx-inspector__body">
                <div className="pptx-inspector__section">
                  <span className="pptx-inspector__tag">{selectedElem.tagName}</span>
                  <span className="pptx-inspector__id">ID: {selectedElem.id}</span>
                </div>

                {/* Position Controls */}
                <div className="pptx-inspector__section">
                  <h4 className="pptx-inspector__title">Position & Size (%)</h4>
                  <div className="pptx-inspector__grid">
                    <label>
                      <span>Left X</span>
                      <input
                        type="number"
                        value={Math.round(selectedElem.xPct || 0)}
                        onChange={(e) =>
                          updateElement(selectedElem.id, { xPct: Number(e.target.value) })
                        }
                      />
                    </label>
                    <label>
                      <span>Top Y</span>
                      <input
                        type="number"
                        value={Math.round(selectedElem.yPct || 0)}
                        onChange={(e) =>
                          updateElement(selectedElem.id, { yPct: Number(e.target.value) })
                        }
                      />
                    </label>
                    <label>
                      <span>Width</span>
                      <input
                        type="number"
                        value={Math.round(selectedElem.wPct || 0)}
                        onChange={(e) =>
                          updateElement(selectedElem.id, { wPct: Number(e.target.value) })
                        }
                      />
                    </label>
                    <label>
                      <span>Height</span>
                      <input
                        type="number"
                        value={Math.round(selectedElem.hPct || 0)}
                        onChange={(e) =>
                          updateElement(selectedElem.id, { hPct: Number(e.target.value) })
                        }
                      />
                    </label>
                  </div>
                </div>

                {/* Text Controls */}
                {selectedElem.type === 'text' && (
                  <div className="pptx-inspector__section">
                    <h4 className="pptx-inspector__title">Text Content</h4>
                    <textarea
                      className="pptx-inspector__textarea"
                      rows={5}
                      value={selectedElem.text}
                      onChange={(e) =>
                        updateElement(selectedElem.id, { text: e.target.value })
                      }
                    />

                    <div className="pptx-inspector__row">
                      <label className="pptx-inspector__checkbox">
                        <input
                          type="checkbox"
                          checked={Boolean(selectedElem.bold)}
                          onChange={(e) =>
                            updateElement(selectedElem.id, { bold: e.target.checked })
                          }
                        />
                        <span>Bold</span>
                      </label>

                      <label className="pptx-inspector__field">
                        <span>Align:</span>
                        <select
                          value={selectedElem.align || 'left'}
                          onChange={(e) =>
                            updateElement(selectedElem.id, { align: e.target.value })
                          }
                        >
                          <option value="left">Left</option>
                          <option value="center">Center</option>
                          <option value="right">Right</option>
                        </select>
                      </label>
                    </div>
                  </div>
                )}

                {/* Image Controls */}
                {selectedElem.type === 'image' && (
                  <div className="pptx-inspector__section">
                    <h4 className="pptx-inspector__title">Image Asset</h4>
                    <div className="pptx-inspector__img-box">
                      {selectedElem.dataUrl ? (
                        <img src={selectedElem.dataUrl} alt="Inspector Preview" />
                      ) : (
                        <span>No Image Set</span>
                      )}
                    </div>
                    <button
                      type="button"
                      className="pptx-editor__btn pptx-editor__btn--small"
                      onClick={() => triggerImageReplacement(selectedElem.id)}
                    >
                      📷 Replace Image File
                    </button>
                  </div>
                )}
              </div>
            ) : (
              <div className="pptx-inspector__none">
                Select an element on the slide canvas to inspect and edit properties.
              </div>
            )}
          </aside>
        </div>
      )}
    </div>
  )
}
