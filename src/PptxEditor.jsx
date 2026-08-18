import React, { useState, useRef, useEffect } from 'react'
import { parsePptxForEditing, exportEditedPptx } from './report/pptxEditorUtils'
import './PptxEditor.css'

export function PptxEditor() {
  const [file, setFile] = useState(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [parsedData, setParsedData] = useState(null)
  const [activeSlideIndex, setActiveSlideIndex] = useState(0)
  const [selectedElementId, setSelectedElementId] = useState(null)
  const [tagFilter, setTagFilter] = useState('all') // 'all', 'text', 'image'
  const [isExporting, setIsExporting] = useState(false)
  const [isDragging, setIsDragging] = useState(false)

  const imageInputRef = useRef(null)
  const [replacingElemId, setReplacingElemId] = useState(null)

  const handleFileUpload = async (uploadedFile) => {
    if (!uploadedFile) return
    if (!uploadedFile.name.toLowerCase().endsWith('.pptx')) {
      setError('Please select a valid PowerPoint (.pptx) file.')
      return
    }

    setLoading(true)
    setError('')
    setFile(uploadedFile)

    try {
      const data = await parsePptxForEditing(uploadedFile)
      setParsedData(data)
      setActiveSlideIndex(0)
      setSelectedElementId(null)
    } catch (err) {
      console.error(err)
      setError(err.message || 'Failed to parse PPTX file structure.')
      setParsedData(null)
    } finally {
      setLoading(false)
    }
  }

  const handleDrop = (e) => {
    e.preventDefault()
    setIsDragging(false)
    const droppedFile = e.dataTransfer.files?.[0]
    if (droppedFile) {
      handleFileUpload(droppedFile)
    }
  }

  const handleDragOver = (e) => {
    e.preventDefault()
    setIsDragging(true)
  }

  const handleDragLeave = () => {
    setIsDragging(false)
  }

  // Update text element in state
  const handleTextChange = (slideIndex, elementId, newText) => {
    setParsedData((prev) => {
      if (!prev) return prev
      const newSlides = [...prev.slides]
      const slide = { ...newSlides[slideIndex] }
      slide.elements = slide.elements.map((elem) => {
        if (elem.id === elementId) {
          return { ...elem, text: newText }
        }
        return elem
      })
      // Update slide title if main title changed
      const titleElem = slide.elements.find((e) => e.type === 'text' && e.isTitle)
      if (titleElem) {
        slide.title = titleElem.text.slice(0, 40) || `Slide ${slideIndex + 1}`
      }
      newSlides[slideIndex] = slide
      return { ...prev, slides: newSlides }
    })
  }

  // Trigger image replacement file selector
  const triggerImageReplacement = (elemId) => {
    setReplacingElemId(elemId)
    if (imageInputRef.current) {
      imageInputRef.current.value = ''
      imageInputRef.current.click()
    }
  }

  // Handle new image file selection
  const handleImageFileChange = (e) => {
    const selectedFile = e.target.files?.[0]
    if (!selectedFile || !replacingElemId) return

    const reader = new FileReader()
    reader.onload = (event) => {
      const dataUrl = event.target.result
      setParsedData((prev) => {
        if (!prev) return prev
        const newSlides = [...prev.slides]
        const slide = { ...newSlides[activeSlideIndex] }
        slide.elements = slide.elements.map((elem) => {
          if (elem.id === replacingElemId) {
            return { ...elem, dataUrl }
          }
          return elem
        })
        newSlides[activeSlideIndex] = slide
        return { ...prev, slides: newSlides }
      })
      setReplacingElemId(null)
    }
    reader.readAsDataURL(selectedFile)
  }

  // Reset a specific element back to original
  const handleResetElement = (slideIndex, elementId) => {
    setParsedData((prev) => {
      if (!prev) return prev
      const newSlides = [...prev.slides]
      const slide = { ...newSlides[slideIndex] }
      slide.elements = slide.elements.map((elem) => {
        if (elem.id === elementId) {
          if (elem.type === 'text') {
            return { ...elem, text: elem.originalText }
          } else if (elem.type === 'image') {
            return { ...elem, dataUrl: elem.originalDataUrl }
          }
        }
        return elem
      })
      newSlides[slideIndex] = slide
      return { ...prev, slides: newSlides }
    })
  }

  // Move slide up or down
  const moveSlide = (fromIndex, toIndex) => {
    if (toIndex < 0 || toIndex >= parsedData.slides.length) return
    setParsedData((prev) => {
      const slides = [...prev.slides]
      const [moved] = slides.splice(fromIndex, 1)
      slides.splice(toIndex, 0, moved)
      // re-number slides
      const renumbered = slides.map((s, idx) => ({ ...s, slideNumber: idx + 1 }))
      return { ...prev, slides: renumbered }
    })
    setActiveSlideIndex(toIndex)
  }

  // Duplicate slide
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

  // Delete slide
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

  // Handle PPTX export
  const handleExport = async () => {
    if (!file || !parsedData) return
    setIsExporting(true)
    try {
      await exportEditedPptx(file, parsedData.slides, `Edited_${file.name}`)
    } catch (err) {
      console.error('Export error:', err)
      alert('Error exporting PPTX: ' + err.message)
    } finally {
      setIsExporting(false)
    }
  }

  const activeSlide = parsedData?.slides[activeSlideIndex]

  // Filter elements for current slide
  const filteredElements = activeSlide?.elements.filter((elem) => {
    if (tagFilter === 'text') return elem.type === 'text'
    if (tagFilter === 'image') return elem.type === 'image'
    return true
  }) || []

  // Count total text & image tags across all slides
  const totalStats = React.useMemo(() => {
    if (!parsedData) return { text: 0, image: 0, total: 0 }
    let text = 0
    let image = 0
    parsedData.slides.forEach((s) => {
      s.elements.forEach((e) => {
        if (e.type === 'text') text++
        if (e.type === 'image') image++
      })
    })
    return { text, image, total: text + image }
  }, [parsedData])

  return (
    <div className="pptx-editor-container">
      {/* Hidden File Input for Image Replacement */}
      <input
        type="file"
        ref={imageInputRef}
        onChange={handleImageFileChange}
        accept="image/png, image/jpeg, image/webp, image/gif, image/svg+xml"
        style={{ display: 'none' }}
      />

      {/* Main Hero Header */}
      <header className="pptx-editor__header">
        <div className="pptx-editor__header-title-group">
          <div className="pptx-editor__logo-badge">✏️</div>
          <div>
            <h1 className="pptx-editor__title">Basic PPTX Editor</h1>
            <p className="pptx-editor__subtitle">
              Edit basic PowerPoint tags (text & images) while preserving slide structure, designs & master layouts
            </p>
          </div>
        </div>

        {parsedData && (
          <div className="pptx-editor__header-actions">
            <button
              type="button"
              className="pptx-editor__btn pptx-editor__btn--ghost"
              onClick={() => {
                if (confirm('Are you sure you want to upload a different presentation? Unsaved changes will be lost.')) {
                  setParsedData(null)
                  setFile(null)
                }
              }}
            >
              📂 Open New PPTX
            </button>

            <button
              type="button"
              className="pptx-editor__btn pptx-editor__btn--primary"
              onClick={handleExport}
              disabled={isExporting}
            >
              {isExporting ? '⏳ Exporting...' : '💾 Export & Download PPTX'}
            </button>
          </div>
        )}
      </header>

      {/* Error Banner */}
      {error && (
        <div className="pptx-editor__alert pptx-editor__alert--error">
          <span>⚠️ {error}</span>
          <button type="button" onClick={() => setError('')}>✕</button>
        </div>
      )}

      {/* 1. Upload State */}
      {!parsedData && !loading && (
        <div className="pptx-editor__upload-wrapper">
          <div
            className={`pptx-editor__dropzone${isDragging ? ' is-dragging' : ''}`}
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
          >
            <div className="pptx-editor__dropzone-icon">📊</div>
            <h3>Drag & Drop your PowerPoint (.pptx) file here</h3>
            <p>Supports all standard PPTX decks with preserved slide structure and XML tags</p>

            <label className="pptx-editor__upload-btn">
              <span>Choose PPTX File</span>
              <input
                type="file"
                accept=".pptx"
                onChange={(e) => handleFileUpload(e.target.files?.[0])}
              />
            </label>
          </div>

          {/* Feature Highlights Grid */}
          <div className="pptx-editor__features">
            <div className="pptx-editor__feature-card">
              <div className="pptx-editor__feature-icon">🏷️</div>
              <h4>Preserves Slide Structure</h4>
              <p>Edits text & images directly inside OpenXML tags (<kbd>&lt;p:sp&gt;</kbd> and <kbd>&lt;p:pic&gt;</kbd>) without messing up master themes.</p>
            </div>
            <div className="pptx-editor__feature-card">
              <div className="pptx-editor__feature-icon">📝</div>
              <h4>Live Text Tag Editing</h4>
              <p>Instantly modify titles, body text, subheadings, and bullet points on any slide deck.</p>
            </div>
            <div className="pptx-editor__feature-card">
              <div className="pptx-editor__feature-icon">🖼️</div>
              <h4>Image Tag Replacement</h4>
              <p>Swap slide pictures and diagrams effortlessly with new high-resolution images.</p>
            </div>
            <div className="pptx-editor__feature-card">
              <div className="pptx-editor__feature-icon">📑</div>
              <h4>Slide Deck Management</h4>
              <p>Reorder, duplicate, or delete slides while retaining all underlying slide layouts.</p>
            </div>
          </div>
        </div>
      )}

      {/* Loading Spinner */}
      {loading && (
        <div className="pptx-editor__loading">
          <div className="pptx-editor__spinner" />
          <p>Parsing presentation structure and extracting slide tags...</p>
        </div>
      )}

      {/* 2. Main Editor Interface */}
      {parsedData && (
        <div className="pptx-editor__workspace">
          {/* Top Metadata Strip */}
          <div className="pptx-editor__meta-strip">
            <div className="pptx-editor__meta-info">
              <span className="pptx-editor__file-badge">📄 {parsedData.filename}</span>
              <span className="pptx-editor__stat-chip">📑 {parsedData.slides.length} Slides</span>
              <span className="pptx-editor__stat-chip">🏷️ {totalStats.text} Text Tags</span>
              <span className="pptx-editor__stat-chip">🖼️ {totalStats.image} Image Tags</span>
            </div>

            <div className="pptx-editor__tag-filters">
              <span className="pptx-editor__filter-label">Filter Tags:</span>
              <button
                type="button"
                className={`pptx-editor__filter-btn${tagFilter === 'all' ? ' is-active' : ''}`}
                onClick={() => setTagFilter('all')}
              >
                All ({activeSlide?.elements.length || 0})
              </button>
              <button
                type="button"
                className={`pptx-editor__filter-btn${tagFilter === 'text' ? ' is-active' : ''}`}
                onClick={() => setTagFilter('text')}
              >
                Text Tags ({activeSlide?.elements.filter(e => e.type === 'text').length || 0})
              </button>
              <button
                type="button"
                className={`pptx-editor__filter-btn${tagFilter === 'image' ? ' is-active' : ''}`}
                onClick={() => setTagFilter('image')}
              >
                Image Tags ({activeSlide?.elements.filter(e => e.type === 'image').length || 0})
              </button>
            </div>
          </div>

          {/* Main 3-Column Studio Grid */}
          <div className="pptx-editor__studio-grid">
            {/* Left Sidebar: Slide Deck List */}
            <aside className="pptx-editor__deck-sidebar">
              <div className="pptx-editor__deck-header">
                <h3>Slide Deck</h3>
                <span className="pptx-editor__deck-count">{parsedData.slides.length} slides</span>
              </div>

              <div className="pptx-editor__slide-list">
                {parsedData.slides.map((slide, idx) => {
                  const isActive = idx === activeSlideIndex
                  const textCount = slide.elements.filter((e) => e.type === 'text').length
                  const imgCount = slide.elements.filter((e) => e.type === 'image').length

                  return (
                    <div
                      key={slide.id}
                      className={`pptx-editor__slide-card${isActive ? ' is-active' : ''}`}
                      onClick={() => setActiveSlideIndex(idx)}
                    >
                      <div className="pptx-editor__slide-num">{slide.slideNumber}</div>
                      <div className="pptx-editor__slide-info">
                        <span className="pptx-editor__slide-title">{slide.title}</span>
                        <div className="pptx-editor__slide-tags-mini">
                          {textCount > 0 && <span className="pptx-mini-tag">📝 {textCount}</span>}
                          {imgCount > 0 && <span className="pptx-mini-tag">🖼️ {imgCount}</span>}
                        </div>
                      </div>

                      {/* Controls */}
                      <div className="pptx-editor__slide-actions" onClick={(e) => e.stopPropagation()}>
                        <button
                          type="button"
                          className="pptx-icon-btn"
                          disabled={idx === 0}
                          onClick={() => moveSlide(idx, idx - 1)}
                          title="Move Slide Up"
                        >
                          ▲
                        </button>
                        <button
                          type="button"
                          className="pptx-icon-btn"
                          disabled={idx === parsedData.slides.length - 1}
                          onClick={() => moveSlide(idx, idx + 1)}
                          title="Move Slide Down"
                        >
                          ▼
                        </button>
                        <button
                          type="button"
                          className="pptx-icon-btn"
                          onClick={() => duplicateSlide(idx)}
                          title="Duplicate Slide"
                        >
                          📋
                        </button>
                        <button
                          type="button"
                          className="pptx-icon-btn pptx-icon-btn--danger"
                          onClick={() => deleteSlide(idx)}
                          title="Delete Slide"
                        >
                          🗑️
                        </button>
                      </div>
                    </div>
                  )
                })}
              </div>
            </aside>

            {/* Center Canvas Preview */}
            <main className="pptx-editor__preview-stage">
              <div className="pptx-editor__preview-header">
                <h3>Slide {activeSlide?.slideNumber}: {activeSlide?.title}</h3>
                <span className="pptx-editor__preview-dims">
                  Aspect Ratio 16:9 ({parsedData.slideSize.wInch.toFixed(1)}" × {parsedData.slideSize.hInch.toFixed(1)}")
                </span>
              </div>

              <div className="pptx-editor__canvas-frame">
                {activeSlide ? (
                  <div
                    className="pptx-editor__canvas"
                    style={{
                      aspectRatio: `${parsedData.slideSize.wInch} / ${parsedData.slideSize.hInch}`,
                    }}
                  >
                    {/* Render visual tag overlays */}
                    {activeSlide.elements.map((elem) => {
                      const isSelected = selectedElementId === elem.id
                      const leftPct = (elem.xInch / parsedData.slideSize.wInch) * 100
                      const topPct = (elem.yInch / parsedData.slideSize.hInch) * 100
                      const widthPct = (elem.wInch / parsedData.slideSize.wInch) * 100
                      const heightPct = (elem.hInch / parsedData.slideSize.hInch) * 100

                      if (elem.type === 'text') {
                        return (
                          <div
                            key={elem.id}
                            className={`pptx-canvas-elem pptx-canvas-elem--text${isSelected ? ' is-selected' : ''}`}
                            style={{
                              left: `${leftPct}%`,
                              top: `${topPct}%`,
                              width: `${Math.max(10, widthPct)}%`,
                              height: `${Math.max(5, heightPct)}%`,
                            }}
                            onClick={() => setSelectedElementId(elem.id)}
                            title={`Click to edit: ${elem.text}`}
                          >
                            <div className="pptx-canvas-elem__badge">{elem.tagName}</div>
                            <div className="pptx-canvas-elem__text">{elem.text}</div>
                          </div>
                        )
                      } else if (elem.type === 'image') {
                        return (
                          <div
                            key={elem.id}
                            className={`pptx-canvas-elem pptx-canvas-elem--image${isSelected ? ' is-selected' : ''}`}
                            style={{
                              left: `${leftPct}%`,
                              top: `${topPct}%`,
                              width: `${Math.max(10, widthPct)}%`,
                              height: `${Math.max(10, heightPct)}%`,
                            }}
                            onClick={() => setSelectedElementId(elem.id)}
                            title="Click to manage image tag"
                          >
                            <div className="pptx-canvas-elem__badge">{elem.tagName}</div>
                            {elem.dataUrl ? (
                              <img src={elem.dataUrl} alt="Slide Tag Asset" className="pptx-canvas-elem__img" />
                            ) : (
                              <div className="pptx-canvas-elem__img-placeholder">Image Placeholder</div>
                            )}
                          </div>
                        )
                      }
                      return null
                    })}
                  </div>
                ) : (
                  <div className="pptx-editor__empty-slide">No slide selected</div>
                )}
              </div>
            </main>

            {/* Right Inspector: Slide Tags & Content Editor */}
            <aside className="pptx-editor__inspector-sidebar">
              <div className="pptx-editor__inspector-header">
                <h3>Editable PPTX Tags</h3>
                <span className="pptx-editor__tag-count">
                  {filteredElements.length} tags on Slide {activeSlideIndex + 1}
                </span>
              </div>

              <div className="pptx-editor__tags-scroll">
                {filteredElements.length === 0 ? (
                  <div className="pptx-editor__no-tags">
                    No matching tags found on this slide.
                  </div>
                ) : (
                  filteredElements.map((elem) => {
                    const isSelected = selectedElementId === elem.id
                    const isModified =
                      elem.type === 'text'
                        ? elem.text !== elem.originalText
                        : elem.dataUrl !== elem.originalDataUrl

                    return (
                      <div
                        key={elem.id}
                        className={`pptx-tag-card${isSelected ? ' is-focused' : ''}${isModified ? ' is-modified' : ''}`}
                        onClick={() => setSelectedElementId(elem.id)}
                      >
                        <div className="pptx-tag-card__header">
                          <div className="pptx-tag-card__badge-row">
                            <span className={`pptx-tag-badge${elem.type === 'image' ? ' pptx-tag-badge--image' : ''}`}>
                              {elem.tagName}
                            </span>
                            {isModified && <span className="pptx-mod-badge">Modified</span>}
                          </div>

                          {isModified && (
                            <button
                              type="button"
                              className="pptx-tag-reset-btn"
                              onClick={(e) => {
                                e.stopPropagation()
                                handleResetElement(activeSlideIndex, elem.id)
                              }}
                              title="Reset element to original"
                            >
                              ↺ Reset
                            </button>
                          )}
                        </div>

                        {/* Text Tag Editor */}
                        {elem.type === 'text' && (
                          <div className="pptx-tag-card__body">
                            <textarea
                              className="pptx-tag-textarea"
                              rows={Math.max(2, Math.min(6, Math.ceil(elem.text.length / 35)))}
                              value={elem.text}
                              onChange={(e) =>
                                handleTextChange(activeSlideIndex, elem.id, e.target.value)
                              }
                              placeholder="Enter slide text..."
                            />
                            <div className="pptx-tag-meta">
                              <span>{elem.text.length} characters</span>
                              <span>Pos: X:{elem.xInch.toFixed(1)}" Y:{elem.yInch.toFixed(1)}"</span>
                            </div>
                          </div>
                        )}

                        {/* Image Tag Editor */}
                        {elem.type === 'image' && (
                          <div className="pptx-tag-card__body">
                            <div className="pptx-image-preview-box">
                              {elem.dataUrl ? (
                                <img src={elem.dataUrl} alt="Tag asset" className="pptx-tag-img" />
                              ) : (
                                <div className="pptx-no-img-text">No Image Data</div>
                              )}
                            </div>
                            <div className="pptx-image-actions">
                              <button
                                type="button"
                                className="pptx-editor__btn pptx-editor__btn--small"
                                onClick={(e) => {
                                  e.stopPropagation()
                                  triggerImageReplacement(elem.id)
                                }}
                              >
                                📷 Replace Image
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    )
                  })
                )}
              </div>
            </aside>
          </div>
        </div>
      )}
    </div>
  )
}
