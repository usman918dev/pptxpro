import React, { useState, useEffect, useRef } from 'react'
import './Navbar.css'

// Inline SVG Icon Helper Components for Zero Dependencies & Instant Load
const Icons = {
  Sparkles: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="m12 3-1.9 5.8a2 2 0 0 1-1.3 1.3L3 12l5.8 1.9a2 2 0 0 1 1.3 1.3L12 21l1.9-5.8a2 2 0 0 1 1.3-1.3L21 12l-5.8-1.9a2 2 0 0 1-1.3-1.3z"/>
    </svg>
  ),
  ShieldCheck: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.8 17 5 19 5a1 1 0 0 1 1 1z"/>
      <path d="m9 12 2 2 4-4"/>
    </svg>
  ),
  Droplets: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M7 16.3c2.2 0 4-1.83 4-4.05 0-1.16-.57-2.26-1.71-3.19S7.29 6.75 7 5.3c-.29 1.45-1.42 2.84-2.41 3.76S3 11.09 3 12.25c0 2.22 1.8 4.05 4 4.05z"/>
      <path d="M12.56 6.6A10.97 10.97 0 0 0 14 3c.5 2.5 2 4.9 4 6.5s3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.2.27-2.3.75-3.3"/>
    </svg>
  ),
  MapPin: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M20 10c0 4.993-5.539 10.193-7.399 11.799a1 1 0 0 1-1.202 0C9.539 20.193 4 14.993 4 10a8 8 0 0 1 16 0"/>
      <circle cx="12" cy="10" r="3"/>
    </svg>
  ),
  Palette: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="13.5" cy="6.5" r=".5" fill="currentColor"/>
      <circle cx="17.5" cy="10.5" r=".5" fill="currentColor"/>
      <circle cx="8.5" cy="7.5" r=".5" fill="currentColor"/>
      <circle cx="6.5" cy="12.5" r=".5" fill="currentColor"/>
      <path d="M12 2C6.5 2 2 6.5 2 12s4.5 10 10 10c.926 0 1.648-.746 1.648-1.688 0-.437-.18-.835-.437-1.125-.29-.289-.438-.652-.438-1.125a1.64 1.64 0 0 1 1.668-1.668h1.996c3.051 0 5.555-2.503 5.555-5.554C21.965 6.012 17.461 2 12 2z"/>
    </svg>
  ),
  LayoutGrid: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect width="7" height="7" x="3" y="3" rx="1"/>
      <rect width="7" height="7" x="14" y="3" rx="1"/>
      <rect width="7" height="7" x="14" y="14" rx="1"/>
      <rect width="7" height="7" x="3" y="14" rx="1"/>
    </svg>
  ),
  FileSpreadsheet: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/>
      <path d="M14 2v4a1 1 0 0 0 1 1h4"/>
      <path d="M8 13h8"/>
      <path d="M8 17h8"/>
      <path d="M10 9h4"/>
    </svg>
  ),
  Files: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M20 7h-3a2 2 0 0 1-2-2V2"/>
      <path d="M9 18a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h6l5 5v9a2 2 0 0 1-2 2Z"/>
      <path d="M3 12v8a2 2 0 0 0 2 2h11"/>
    </svg>
  ),
  Presentation: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M2 3h20"/>
      <path d="M21 3v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V3"/>
      <path d="m7 21 5-5 5 5"/>
    </svg>
  ),
  ImageDown: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M10.3 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v10l-3.1-3.1a2 2 0 0 0-2.814.014L6 21"/>
      <circle cx="9" cy="9" r="2"/>
      <path d="M14 19l3 3 3-3"/>
      <path d="M17 14v8"/>
    </svg>
  ),
  Search: () => (
    <svg className="nav-icon nav-icon--small" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <circle cx="11" cy="11" r="8"/>
      <path d="m21 21-4.3-4.3"/>
    </svg>
  ),
  ChevronDown: () => (
    <svg className="nav-icon nav-icon--xs" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="m6 9 6 6 6-6"/>
    </svg>
  ),
  Menu: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="4" x2="20" y1="12" y2="12"/>
      <line x1="4" x2="20" y1="6" y2="6"/>
      <line x1="4" x2="20" y1="18" y2="18"/>
    </svg>
  ),
  FileEdit: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M12 20h9"/>
      <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/>
    </svg>
  ),
  X: () => (
    <svg className="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M18 6 6 18"/>
      <path d="m6 6 12 12"/>
    </svg>
  )
}

export const TOOL_CATEGORIES = [
  {
    id: 'reports',
    title: 'Presentation Reports',
    shortTitle: 'Reports',
    iconName: 'Sparkles',
    tools: [
      {
        route: '/',
        label: 'Clean Punjab',
        iconName: 'Sparkles',
        description: 'Plots Cleaning-Activity report with automated before/after slides',
      },
      {
        route: '/compliance',
        label: 'Compliance Report',
        iconName: 'ShieldCheck',
        description: 'Suthra Punjab compliance tracking & verification report',
      },
      {
        route: '/desilting',
        label: 'Desilting Report',
        iconName: 'Droplets',
        description: '3-stage Sector Desilting presentation generator',
      },
      {
        route: '/daily-plot',
        label: 'OTC Plot Report',
        iconName: 'MapPin',
        description: 'Daily Plot clearance reports (Urban & Rural variants)',
      },
    ],
  },
  {
    id: 'studio',
    title: 'Custom Template Studio',
    shortTitle: 'Studio',
    iconName: 'Palette',
    tools: [
      {
        route: '/master',
        label: 'Master Creator',
        iconName: 'Palette',
        badge: 'STUDIO',
        description: 'Design custom master slide layouts, placeholders & presets',
      },
    ],
  },
  {
    id: 'tools',
    title: 'Media & Document Tools',
    shortTitle: 'Tools & Converters',
    iconName: 'LayoutGrid',
    tools: [
      {
        route: '/gps-pdf',
        label: 'PPTX to GPS PDF',
        iconName: 'MapPin',
        badge: 'NEW',
        description: 'Convert PPTX to PDF with bottom 30% GPS OCR & Google Maps links',
      },
      {
        route: '/collage',
        label: 'Collage Maker',
        iconName: 'LayoutGrid',
        badge: 'POPULAR',
        description: 'Create custom photo grids & layout collages effortlessly',
      },
      {
        route: '/pdf',
        label: 'PDF to PPTX',
        iconName: 'FileSpreadsheet',
        description: 'Convert PDF slides directly into edit-ready PPTX decks',
      },
      {
        route: '/merge-pdf',
        label: 'Merge PDF',
        iconName: 'Files',
        description: 'Combine multiple PDF documents into a single document',
      },
      {
        route: '/merge',
        label: 'Merge Presentations',
        iconName: 'Presentation',
        description: 'Merge multiple PPTX decks into one cohesive presentation',
      },
      {
        route: '/pptx-editor',
        label: 'PPTX Editor',
        iconName: 'FileEdit',
        badge: 'NEW',
        description: 'Edit basic text & image tags preserving slide structure',
      },
      {
        route: '/extract',
        label: 'Extract Images',
        iconName: 'ImageDown',
        description: 'Extract high-resolution image assets from PPTX files',
      },
    ],
  },
]

// Flat list for search palette
const ALL_TOOLS = TOOL_CATEGORIES.flatMap((cat) =>
  cat.tools.map((t) => ({ ...t, categoryTitle: cat.title })),
)

export default function Navbar({
  currentRoute,
  dailyVariant,
  setDailyVariant,
  designerMode,
  setDesignerMode,
  customLayout,
  ROUTES,
}) {
  const [activeDropdown, setActiveDropdown] = useState(null)
  const [searchOpen, setSearchOpen] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const [selectedIndex, setSelectedIndex] = useState(0)
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false)
  const searchInputRef = useRef(null)
  const dropdownRef = useRef(null)

  // Filter tools for search palette
  const filteredTools = ALL_TOOLS.filter((t) => {
    const q = searchQuery.toLowerCase().trim()
    if (!q) return true
    return (
      t.label.toLowerCase().includes(q) ||
      t.description.toLowerCase().includes(q) ||
      t.categoryTitle.toLowerCase().includes(q)
    )
  })

  // Keyboard shortcut listener (Ctrl+K / Cmd+K / Slash)
  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault()
        setSearchOpen((prev) => !prev)
      } else if (e.key === 'Escape' && searchOpen) {
        setSearchOpen(false)
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [searchOpen])

  // Focus search input when modal opens
  useEffect(() => {
    if (searchOpen) {
      setTimeout(() => searchInputRef.current?.focus(), 50)
      setSelectedIndex(0)
    } else {
      setSearchQuery('')
    }
  }, [searchOpen])

  // Close dropdown on outside click
  useEffect(() => {
    const handleOutsideClick = (e) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target)) {
        setActiveDropdown(null)
      }
    }
    window.addEventListener('click', handleOutsideClick)
    return () => window.removeEventListener('click', handleOutsideClick)
  }, [])

  // Handle navigation
  const navigateTo = (route) => {
    setActiveDropdown(null)
    setSearchOpen(false)
    setMobileMenuOpen(false)
    if (currentRoute !== route) {
      window.location.pathname = route
    }
  }

  // Handle keyboard navigation inside search modal
  const handleSearchKeyDown = (e) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault()
      setSelectedIndex((prev) => (prev + 1) % (filteredTools.length || 1))
    } else if (e.key === 'ArrowUp') {
      e.preventDefault()
      setSelectedIndex((prev) => (prev - 1 + filteredTools.length) % (filteredTools.length || 1))
    } else if (e.key === 'Enter' && filteredTools[selectedIndex]) {
      e.preventDefault()
      navigateTo(filteredTools[selectedIndex].route)
    }
  }

  const renderIcon = (iconName) => {
    const Component = Icons[iconName]
    return Component ? <Component /> : null
  }

  // Find active category
  const activeCategory = TOOL_CATEGORIES.find((cat) =>
    cat.tools.some((t) => t.route === currentRoute),
  )

  return (
    <nav className="pptx-navbar">
      <div className="pptx-navbar__container" ref={dropdownRef}>
        {/* Brand & Logo */}
        <div className="pptx-navbar__brand" onClick={() => navigateTo('/')} role="button" tabIndex={0}>
          <div className="pptx-navbar__logo-badge">
            <Icons.Sparkles />
          </div>
          <div className="pptx-navbar__title-group">
            <span className="pptx-navbar__title">PPTX Pro</span>
            <span className="pptx-navbar__tagline">Presentation Suite</span>
          </div>
        </div>

        {/* Desktop Categorized Menu & Quick Launcher Bar */}
        <div className="pptx-navbar__desktop-menu">
          {TOOL_CATEGORIES.map((category) => {
            const hasActiveChild = category.tools.some((t) => t.route === currentRoute)
            const isOpen = activeDropdown === category.id

            return (
              <div key={category.id} className="pptx-navbar__dropdown-wrapper">
                <button
                  type="button"
                  className={`pptx-navbar__cat-btn${hasActiveChild ? ' is-active' : ''}${isOpen ? ' is-open' : ''}`}
                  onClick={(e) => {
                    e.stopPropagation()
                    setActiveDropdown(isOpen ? null : category.id)
                  }}
                >
                  <span className="pptx-navbar__cat-icon">{renderIcon(category.iconName)}</span>
                  <span>{category.shortTitle}</span>
                  <span className="pptx-navbar__cat-arrow">
                    <Icons.ChevronDown />
                  </span>
                </button>

                {/* Dropdown Megamenu */}
                {isOpen && (
                  <div className="pptx-navbar__megamenu">
                    <div className="pptx-navbar__megamenu-header">
                      <span>{category.title}</span>
                    </div>
                    <div className="pptx-navbar__megamenu-grid">
                      {category.tools.map((tool) => {
                        const isCurrent = tool.route === currentRoute
                        return (
                          <button
                            key={tool.route}
                            type="button"
                            className={`pptx-navbar__item-card${isCurrent ? ' is-active' : ''}`}
                            onClick={() => navigateTo(tool.route)}
                          >
                            <div className="pptx-navbar__item-icon-wrap">{renderIcon(tool.iconName)}</div>
                            <div className="pptx-navbar__item-details">
                              <div className="pptx-navbar__item-title-row">
                                <span className="pptx-navbar__item-title">{tool.label}</span>
                                {tool.badge && <span className="pptx-navbar__badge">{tool.badge}</span>}
                              </div>
                              <p className="pptx-navbar__item-desc">{tool.description}</p>
                            </div>
                          </button>
                        )
                      })}
                    </div>
                  </div>
                )}
              </div>
            )
          })}

          {/* Direct Quick Shortcuts Strip for Most Requested Tools */}
          <div className="pptx-navbar__shortcuts">
            <span className="pptx-navbar__shortcut-divider" />
            
            {/* Quick Master Creator Link */}
            <button
              type="button"
              className={`pptx-navbar__quick-link${currentRoute === ROUTES.master ? ' is-active' : ''}`}
              onClick={() => navigateTo(ROUTES.master)}
              title="Design Master Templates"
            >
              <Icons.Palette />
              <span>Master Creator</span>
              <span className="pptx-navbar__mini-badge">STUDIO</span>
            </button>

            {/* Quick Collage Maker Link */}
            <button
              type="button"
              className={`pptx-navbar__quick-link${currentRoute === ROUTES.collage ? ' is-active' : ''}`}
              onClick={() => navigateTo(ROUTES.collage)}
              title="Create Image Collages"
            >
              <Icons.LayoutGrid />
              <span>Collage Maker</span>
            </button>

            {/* Quick PPTX Editor Link */}
            <button
              type="button"
              className={`pptx-navbar__quick-link${currentRoute === ROUTES.pptxEditor ? ' is-active' : ''}`}
              onClick={() => navigateTo(ROUTES.pptxEditor)}
              title="Edit Basic PPTX Tags"
            >
              <Icons.FileEdit />
              <span>PPTX Editor</span>
              <span className="pptx-navbar__mini-badge">NEW</span>
            </button>

            {/* Quick PDF Converter Link */}
            <button
              type="button"
              className={`pptx-navbar__quick-link${currentRoute === ROUTES.pdf ? ' is-active' : ''}`}
              onClick={() => navigateTo(ROUTES.pdf)}
              title="Convert PDF to PPTX"
            >
              <Icons.FileSpreadsheet />
              <span>PDF to PPTX</span>
            </button>
          </div>
        </div>

        {/* Right Action Utilities: Quick Search Trigger & Mobile Toggle */}
        <div className="pptx-navbar__right">
          <button
            type="button"
            className="pptx-navbar__search-trigger"
            onClick={() => setSearchOpen(true)}
            title="Search tools (Ctrl+K)"
          >
            <Icons.Search />
            <span className="pptx-navbar__search-label">Quick Search...</span>
            <kbd className="pptx-navbar__search-kbd">Ctrl K</kbd>
          </button>

          <button
            type="button"
            className="pptx-navbar__mobile-toggle"
            onClick={() => setMobileMenuOpen((prev) => !prev)}
            aria-label="Toggle Navigation Menu"
          >
            {mobileMenuOpen ? <Icons.X /> : <Icons.Menu />}
          </button>
        </div>
      </div>

      {/* Contextual Subnav Switcher (Urban/Rural or Design/Use Template) */}
      {(currentRoute === ROUTES.dailyPlot || currentRoute === ROUTES.master) && (
        <div className="pptx-navbar__subnav">
          <div className="pptx-navbar__subnav-container">
            {currentRoute === ROUTES.dailyPlot && (
              <div className="pptx-navbar__subnav-group">
                <span className="pptx-navbar__subnav-label">Report Variant:</span>
                <button
                  type="button"
                  className={`pptx-navbar__subnav-btn${dailyVariant === 'urban' ? ' is-active' : ''}`}
                  onClick={() => setDailyVariant('urban')}
                >
                  🏙️ Urban
                </button>
                <button
                  type="button"
                  className={`pptx-navbar__subnav-btn${dailyVariant === 'rural' ? ' is-active' : ''}`}
                  onClick={() => setDailyVariant('rural')}
                >
                  🏞️ Rural
                </button>
              </div>
            )}

            {currentRoute === ROUTES.master && (
              <div className="pptx-navbar__subnav-group">
                <span className="pptx-navbar__subnav-label">Master Designer Mode:</span>
                <button
                  type="button"
                  className={`pptx-navbar__subnav-btn${designerMode === 'design' ? ' is-active' : ''}`}
                  onClick={() => setDesignerMode('design')}
                >
                  ✏️ Design Layout
                </button>
                <button
                  type="button"
                  className={`pptx-navbar__subnav-btn${designerMode === 'use' ? ' is-active' : ''}`}
                  onClick={() => {
                    if (customLayout?.placeholders?.length) {
                      setDesignerMode('use')
                    } else {
                      alert('Please add at least one Image Placeholder to your master slide layout before switching!')
                    }
                  }}
                >
                  🚀 Use Template
                </button>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Mobile Drawer Menu */}
      {mobileMenuOpen && (
        <div className="pptx-navbar__mobile-drawer">
          <div className="pptx-navbar__mobile-search">
            <button
              type="button"
              className="pptx-navbar__search-trigger is-full"
              onClick={() => {
                setMobileMenuOpen(false)
                setSearchOpen(true)
              }}
            >
              <Icons.Search />
              <span>Search all 10 tools...</span>
            </button>
          </div>

          <div className="pptx-navbar__mobile-categories">
            {TOOL_CATEGORIES.map((cat) => (
              <div key={cat.id} className="pptx-navbar__mobile-cat-group">
                <div className="pptx-navbar__mobile-cat-header">
                  {renderIcon(cat.iconName)}
                  <span>{cat.title}</span>
                </div>
                <div className="pptx-navbar__mobile-items">
                  {cat.tools.map((tool) => (
                    <button
                      key={tool.route}
                      type="button"
                      className={`pptx-navbar__mobile-item${tool.route === currentRoute ? ' is-active' : ''}`}
                      onClick={() => navigateTo(tool.route)}
                    >
                      <span className="pptx-navbar__mobile-item-icon">{renderIcon(tool.iconName)}</span>
                      <div className="pptx-navbar__mobile-item-text">
                        <span className="pptx-navbar__mobile-item-label">{tool.label}</span>
                        <span className="pptx-navbar__mobile-item-sub">{tool.description}</span>
                      </div>
                      {tool.badge && <span className="pptx-navbar__badge">{tool.badge}</span>}
                    </button>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Quick Search Palette Modal (`Ctrl + K`) */}
      {searchOpen && (
        <div className="pptx-search-modal__overlay" onClick={() => setSearchOpen(false)}>
          <div
            className="pptx-search-modal__content"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="pptx-search-modal__input-wrap">
              <Icons.Search />
              <input
                ref={searchInputRef}
                type="text"
                className="pptx-search-modal__input"
                placeholder="Search tools (e.g. Master Creator, Collage Maker, PDF...)"
                value={searchQuery}
                onChange={(e) => {
                  setSearchQuery(e.target.value)
                  setSelectedIndex(0)
                }}
                onKeyDown={handleSearchKeyDown}
              />
              <button
                type="button"
                className="pptx-search-modal__close"
                onClick={() => setSearchOpen(false)}
              >
                <Icons.X />
              </button>
            </div>

            <div className="pptx-search-modal__results">
              {filteredTools.length === 0 ? (
                <div className="pptx-search-modal__empty">
                  No matching tools found for "{searchQuery}"
                </div>
              ) : (
                filteredTools.map((tool, idx) => {
                  const isSelected = idx === selectedIndex
                  const isCurrent = tool.route === currentRoute
                  return (
                    <div
                      key={tool.route}
                      className={`pptx-search-modal__item${isSelected ? ' is-selected' : ''}${isCurrent ? ' is-current' : ''}`}
                      onClick={() => navigateTo(tool.route)}
                      onMouseEnter={() => setSelectedIndex(idx)}
                    >
                      <div className="pptx-search-modal__item-icon">{renderIcon(tool.iconName)}</div>
                      <div className="pptx-search-modal__item-info">
                        <div className="pptx-search-modal__item-title-row">
                          <span className="pptx-search-modal__item-title">{tool.label}</span>
                          <span className="pptx-search-modal__item-cat">{tool.categoryTitle}</span>
                          {tool.badge && <span className="pptx-navbar__badge">{tool.badge}</span>}
                        </div>
                        <span className="pptx-search-modal__item-desc">{tool.description}</span>
                      </div>
                      <span className="pptx-search-modal__enter-hint">↵ Jump</span>
                    </div>
                  )
                })
              )}
            </div>

            <div className="pptx-search-modal__footer">
              <span>Use <kbd>↑</kbd> <kbd>↓</kbd> to navigate, <kbd>Enter</kbd> to select, <kbd>Esc</kbd> to dismiss</span>
            </div>
          </div>
        </div>
      )}
    </nav>
  )
}
