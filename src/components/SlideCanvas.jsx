import { DropSlot } from './DropSlot'
import { DEFAULT_SLOTS } from '../config/templates'

export function SlideCanvas({
  pair,
  slots = DEFAULT_SLOTS,
  onChange,
  backgroundUrl,
  showText,
  textValue,
  textPlaceholder,
  textBoxes = [],
  onTextChange,
}) {
  const displayText = textValue?.trim() || textPlaceholder || ''
  return (
    <div
      className="slide-canvas"
      style={{ backgroundImage: `url(${backgroundUrl})` }}
    >
      {showText && displayText && (
        <div className="slide-text" aria-hidden="true">
          {displayText}
        </div>
      )}
      {slots.map((slot) => (
        <DropSlot
          key={slot.key}
          label={slot.label}
          value={pair?.[slot.key]}
          onChange={(value) => onChange(slot.key, value)}
          className={slot.className}
          style={slot.style}
          urlMode={slot.urlMode || 'inline'}
        />
      ))}
      {textBoxes.map((box) => {
        const textVal = pair?.[box.key] ?? ''
        return (
          <input
            key={box.key}
            type="text"
            className="custom-slide-textbox-input"
            value={textVal}
            onChange={(e) => onTextChange?.(box.key, e.target.value)}
            placeholder={box.textDefault || 'Enter text'}
            style={{
              position: 'absolute',
              left: `${(box.x / 13.333) * 100}%`,
              top: `${(box.y / 7.5) * 100}%`,
              width: `${(box.w / 13.333) * 100}%`,
              height: `${(box.h / 7.5) * 100}%`,
              fontSize: `${((box.fontSize || 20) * 0.104).toFixed(3)}cqw`,
              fontFamily: box.fontFace || 'Calibri',
              color: box.fontColor ? `#${box.fontColor}` : '#111111',
              fontWeight: box.bold ? 'bold' : 'normal',
              textAlign: box.align || 'center',
              background: 'rgba(255, 255, 255, 0.75)',
              border: '1px dashed rgba(11, 122, 56, 0.4)',
              borderRadius: '8px',
              padding: '4px 8px',
              outline: 'none',
              boxSizing: 'border-box',
            }}
          />
        )
      })}
    </div>
  )
}
