import { useRef, useState, useCallback } from 'react'
import { PDFDocument } from 'pdf-lib'

// Helper to format file size in a human readable way
const formatSize = (bytes) => {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(2))} ${sizes[i]}`
}

// Helper to get the number of pages from a PDF file using pdf-lib
const getPdfPageCount = async (file) => {
  try {
    const arrayBuffer = await file.arrayBuffer()
    const pdfDoc = await PDFDocument.load(arrayBuffer, { ignoreEncryption: true })
    return pdfDoc.getPageCount()
  } catch (err) {
    console.error('Failed to parse page count for', file.name, err)
    return 0
  }
}

function UploadZone({ onFiles, isLoading }) {
  const inputRef = useRef(null)
  const [isDragging, setIsDragging] = useState(false)

  const handleFiles = (fileList) => {
    if (!fileList || fileList.length === 0) return
    const pdfFiles = Array.from(fileList).filter((file) =>
      file.name.toLowerCase().endsWith('.pdf')
    )
    if (pdfFiles.length === 0) {
      alert('Please select valid .pdf files.')
      return
    }
    onFiles(pdfFiles)
  }

  return (
    <div
      className={`merger-upload${isDragging ? ' is-dragging' : ''}${isLoading ? ' is-loading' : ''}`}
      onClick={() => !isLoading && inputRef.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true) }}
      onDragEnter={(e) => { e.preventDefault(); setIsDragging(true) }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={(e) => {
        e.preventDefault()
        setIsDragging(false)
        if (e.dataTransfer.files) handleFiles(e.dataTransfer.files)
      }}
    >
      <div className="merger-upload__inner">
        <span className="merger-upload__icon" aria-hidden="true">📕</span>
        <p className="merger-upload__label">Drop your PDF files here or click to browse</p>
        <p className="merger-upload__hint">Select multiple PDF files to combine them into one document</p>
      </div>
      <input
        ref={inputRef}
        type="file"
        multiple
        accept=".pdf,application/pdf"
        className="merger-upload__input"
        onChange={(e) => {
          if (e.target.files) handleFiles(e.target.files)
          e.target.value = ''
        }}
      />
    </div>
  )
}

function MergerItem({
  fileInfo,
  index,
  totalItems,
  onMoveUp,
  onMoveDown,
  onRemove,
  onDragStart,
  onDragOver,
  onDrop,
  onDragEnd,
  isDragging,
  isDragOver
}) {
  return (
    <div
      className={`merger-item${isDragging ? ' is-dragging' : ''}${isDragOver ? ' is-drag-over' : ''}`}
      draggable
      onDragStart={onDragStart}
      onDragOver={(e) => { e.preventDefault(); onDragOver() }}
      onDrop={(e) => { e.preventDefault(); onDrop() }}
      onDragEnd={onDragEnd}
    >
      <div className="merger-item__left">
        <span className="merger-item__drag-handle" title="Drag to reorder">⠿</span>
        <span className="merger-item__index">{index + 1}</span>
        <div className="merger-item__details">
          <span className="merger-item__name" title={fileInfo.file.name}>{fileInfo.file.name}</span>
          <span className="merger-item__meta">
            {formatSize(fileInfo.file.size)} · {fileInfo.pageCount} page{fileInfo.pageCount !== 1 ? 's' : ''}
          </span>
        </div>
      </div>
      <div className="merger-item__right">
        <button
          type="button"
          className="merger-item__btn-move"
          onClick={(e) => { e.stopPropagation(); onMoveUp() }}
          disabled={index === 0}
          title="Move Up"
        >
          ▲
        </button>
        <button
          type="button"
          className="merger-item__btn-move"
          onClick={(e) => { e.stopPropagation(); onMoveDown() }}
          disabled={index === totalItems - 1}
          title="Move Down"
        >
          ▼
        </button>
        <button
          type="button"
          className="merger-item__btn-delete"
          onClick={(e) => { e.stopPropagation(); onRemove() }}
          title="Remove from list"
        >
          Remove
        </button>
      </div>
    </div>
  )
}

export function PdfMerger() {
  const [files, setFiles] = useState([])
  const [isProcessing, setIsProcessing] = useState(false)
  const [status, setStatus] = useState({ type: 'idle', message: '' })
  const [dragIndex, setDragIndex] = useState(null)
  const [dragOverIndex, setDragOverIndex] = useState(null)
  const inputRef = useRef(null)

  const handleAddFiles = useCallback(async (newFiles) => {
    setIsProcessing(true)
    setStatus({ type: 'working', message: 'Analyzing PDF files...' })
    try {
      const addedInfos = []
      for (const file of newFiles) {
        const pageCount = await getPdfPageCount(file)
        addedInfos.push({
          id: `${file.name}-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
          file,
          pageCount
        })
      }
      setFiles((prev) => [...prev, ...addedInfos])
      setStatus({ type: 'success', message: `Added ${newFiles.length} PDF file(s).` })
    } catch (err) {
      console.error('Failed to add files:', err)
      setStatus({ type: 'error', message: 'Failed to analyze some PDF files.' })
    } finally {
      setIsProcessing(false)
    }
  }, [])

  const handleRemove = (index) => {
    setFiles((prev) => prev.filter((_, i) => i !== index))
    setStatus({ type: 'idle', message: '' })
  }

  const handleMoveUp = (index) => {
    if (index === 0) return
    setFiles((prev) => {
      const next = [...prev]
      const temp = next[index]
      next[index] = next[index - 1]
      next[index - 1] = temp
      return next
    })
  }

  const handleMoveDown = (index) => {
    if (index === files.length - 1) return
    setFiles((prev) => {
      const next = [...prev]
      const temp = next[index]
      next[index] = next[index + 1]
      next[index + 1] = temp
      return next
    })
  }

  const handleMerge = async () => {
    if (files.length < 2) {
      alert('Please add at least 2 PDF files to merge.')
      return
    }

    setIsProcessing(true)
    setStatus({ type: 'working', message: 'Merging PDF documents...' })

    try {
      const mergedPdf = await PDFDocument.create()

      for (let i = 0; i < files.length; i++) {
        const fileInfo = files[i]
        setStatus({
          type: 'working',
          message: `Processing "${fileInfo.file.name}" (${i + 1} of ${files.length})...`
        })

        const arrayBuffer = await fileInfo.file.arrayBuffer()
        const pdfDoc = await PDFDocument.load(arrayBuffer, { ignoreEncryption: true })
        const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices())
        copiedPages.forEach((page) => mergedPdf.addPage(page))
      }

      const totalMergedPages = mergedPdf.getPageCount()
      setStatus({ type: 'working', message: 'Generating merged PDF file...' })

      const pdfBytes = await mergedPdf.save()
      const blob = new Blob([pdfBytes], { type: 'application/pdf' })

      // Download
      const dateStr = new Date().toISOString().slice(0, 10)
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `Merged_PDF_${dateStr}.pdf`
      a.click()
      URL.revokeObjectURL(url)

      setStatus({
        type: 'success',
        message: `Successfully merged ${files.length} PDFs into a single document with ${totalMergedPages} total pages! Download started.`
      })
    } catch (err) {
      console.error(err)
      setStatus({
        type: 'error',
        message: `Merge failed: ${err.message || 'Unknown error'}`
      })
    } finally {
      setIsProcessing(false)
    }
  }

  const handleClearAll = () => {
    if (window.confirm('Clear all files from the list?')) {
      setFiles([])
      setStatus({ type: 'idle', message: '' })
    }
  }

  // Drag and Drop
  const handleDragStart = (index) => {
    setDragIndex(index)
    setDragOverIndex(null)
  }

  const handleDragOver = (index) => {
    if (dragIndex === null || dragIndex === index) return
    setDragOverIndex(index)
  }

  const handleDrop = (index) => {
    if (dragIndex === null || dragIndex === index) return
    setFiles((prev) => {
      const next = [...prev]
      const [moved] = next.splice(dragIndex, 1)
      next.splice(index, 0, moved)
      return next
    })
    setDragIndex(null)
    setDragOverIndex(null)
  }

  const handleDragEnd = () => {
    setDragIndex(null)
    setDragOverIndex(null)
  }

  const totalPages = files.reduce((acc, f) => acc + f.pageCount, 0)
  const canMerge = files.length >= 2 && !isProcessing

  return (
    <div className="merger">
      {files.length === 0 ? (
        <UploadZone onFiles={handleAddFiles} isLoading={isProcessing} />
      ) : (
        <>
          <div className="merger__toolbar">
            <div className="merger__toolbar-left">
              <span className="merger__count" style={{ fontSize: '14px', fontWeight: 'bold', background: 'rgba(11, 122, 56, 0.1)', color: '#0b7a38', padding: '6px 14px', borderRadius: '20px' }}>
                📄 {files.length} PDF{files.length !== 1 ? 's' : ''} · {totalPages} Total Page{totalPages !== 1 ? 's' : ''}
              </span>
              <button
                type="button"
                className="ghost"
                onClick={handleClearAll}
                disabled={isProcessing}
              >
                Clear All
              </button>
              <button
                type="button"
                className="ghost"
                onClick={() => inputRef.current?.click()}
                disabled={isProcessing}
              >
                Add More PDFs
              </button>
              <input
                ref={inputRef}
                type="file"
                multiple
                accept=".pdf,application/pdf"
                className="merger-upload__input"
                onChange={(e) => {
                  if (e.target.files) handleAddFiles(e.target.files)
                  e.target.value = ''
                }}
              />
            </div>
            <div className="merger__toolbar-right">
              <button
                type="button"
                className="button"
                onClick={handleMerge}
                disabled={!canMerge}
              >
                {isProcessing && status.type === 'working' ? 'Merging PDFs...' : 'Merge PDF Files'}
              </button>
            </div>
          </div>

          {status.message && (
            <p className={`app__note app__note--${status.type}`}>
              {status.type === 'working' && <span className="extractor-spinner" style={{ display: 'inline-block', verticalAlign: 'middle', marginRight: '10px', width: '16px', height: '16px', borderWidth: '2px' }} />}
              {status.message}
            </p>
          )}

          <p className="extractor__hint">
            Drag items by the handle ⠿ to reorder them · Total pages after merging will be {totalPages}.
          </p>

          <div className="merger__list">
            {files.map((fileInfo, index) => (
              <MergerItem
                key={fileInfo.id}
                fileInfo={fileInfo}
                index={index}
                totalItems={files.length}
                onMoveUp={() => handleMoveUp(index)}
                onMoveDown={() => handleMoveDown(index)}
                onRemove={() => handleRemove(index)}
                onDragStart={() => handleDragStart(index)}
                onDragOver={() => handleDragOver(index)}
                onDrop={() => handleDrop(index)}
                onDragEnd={handleDragEnd}
                isDragging={dragIndex === index}
                isDragOver={dragOverIndex === index}
              />
            ))}
          </div>
        </>
      )}
    </div>
  )
}
