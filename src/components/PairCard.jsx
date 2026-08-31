import { SlideCanvas } from './SlideCanvas'
import { isPairComplete, hasPairContent } from '../utils/pairUtils'

export function PairCard({
  pair,
  index,
  template,
  slotKeys,
  requiresText,
  movableCount,
  dragIndex,
  dragOverIndex,
  moveMenuIndex,
  pairRef,
  moveMenuRef,
  onUpdatePair,
  onSwapImages,
  onDragStart,
  onDragEnd,
  onDragOver,
  onDrop,
  onMovePair,
  onSetMoveMenuIndex,
}) {
  const canDrag = hasPairContent(pair, { slotKeys, requiresText })
  const isDragging = dragIndex === index
  const isDragOver = dragOverIndex === index
  const isMenuOpen = moveMenuIndex === index
  const canMove = canDrag && movableCount > 1
  const isComplete = isPairComplete(pair, { slotKeys, requiresText })
  const statusLabel = isComplete
    ? 'Complete'
    : hasPairContent(pair, { slotKeys, requiresText })
      ? 'In progress'
      : 'Empty'

  return (
    <article
      className={`pair${isDragging ? ' is-dragging' : ''}${
        isDragOver ? ' is-drag-over' : ''
      }`}
      ref={pairRef}
      onDragOver={(event) => onDragOver(event, index)}
      onDrop={(event) => onDrop(event, index)}
    >
      <div className="pair__header">
        <p className="label">Pair {index + 1}</p>
        <div className="pair__controls">
          <span className={`pair__status${isComplete ? ' is-complete' : ''}`}>
            {statusLabel}
          </span>
          <button
            type="button"
            className="ghost"
            onClick={() => onSwapImages(index)}
            disabled={!slotKeys.some((key) => pair?.[key])}
            title="Swap before and after images"
          >
            Swap
          </button>
          <div
            className="pair__move"
            ref={isMenuOpen ? moveMenuRef : null}
          >
            <button
              type="button"
              className="pair__drag"
              draggable={canDrag}
              disabled={!canMove}
              aria-label="Drag to reorder or click to move"
              aria-expanded={isMenuOpen}
              aria-haspopup="menu"
              title={
                canMove
                  ? 'Drag to reorder or click to move'
                  : 'Add another slide to enable moving'
              }
              onDragStart={(event) => onDragStart(event, index)}
              onDragEnd={onDragEnd}
              onClick={(event) => {
                if (!canMove) {
                  return
                }
                event.stopPropagation()
                onSetMoveMenuIndex((prev) => (prev === index ? null : index))
              }}
            >
              Drag
            </button>
            {isMenuOpen && (
              <div className="pair__move-menu" role="menu">
                <p className="pair__move-title">Move to position</p>
                <div className="pair__move-list">
                  {Array.from({ length: movableCount }, (_, target) => {
                    const position = target + 1
                    const isCurrent = target === index
                    return (
                      <button
                        type="button"
                        key={`move-${index}-to-${position}`}
                        className={`pair__move-option${
                          isCurrent ? ' is-current' : ''
                        }`}
                        disabled={isCurrent}
                        onClick={() => {
                          onMovePair(index, target)
                          onSetMoveMenuIndex(null)
                        }}
                      >
                        {position}
                      </button>
                    )
                  })}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
      {template.textBox && (
        <div className="pair__text">
          <label className="label" htmlFor={`pair-text-${index}`}>
            {template.textLabel || 'Slide text'}
          </label>
          <input
            id={`pair-text-${index}`}
            type="text"
            value={pair.slideText || ''}
            onChange={(event) =>
              onUpdatePair(index, 'slideText', event.target.value)
            }
            placeholder={template.textDefault || 'Enter text'}
          />
        </div>
      )}
      <SlideCanvas
        pair={pair}
        slots={template.slots}
        onChange={(key, value) => onUpdatePair(index, key, value)}
        backgroundUrl={template.masterBgUrl}
        showText={Boolean(template.textBox)}
        textValue={pair.slideText || ''}
        textPlaceholder={template.textDefault || ''}
        textBoxes={template.textBoxes || []}
        onTextChange={(key, value) => onUpdatePair(index, key, value)}
      />
    </article>
  )
}
