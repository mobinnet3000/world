/* ==========================================================================
   Word Online Main Application Controller & UI Event Handlers
   ========================================================================== */

const App = {
  currentZoom: 100,
  saveTimeout: null,

  init() {
    this.bindRibbonTabs();
    this.bindBackstage();
    this.bindToolbarButtons();
    this.bindModals();
    this.bindZoom();
    this.bindKeyboardShortcuts();
    this.bindThemeToggle();
    this.loadInitialDocument();
    
    Editor.init();
    this.updateStats();

    // Render Lucide icons if available
    if (typeof lucide !== 'undefined') {
      lucide.createIcons();
    }
  },

  /**
   * Ribbon Tabs switching
   */
  bindRibbonTabs() {
    const tabs = document.querySelectorAll('.ribbon-tab:not(.file-tab)');
    const panels = document.querySelectorAll('.ribbon-panel');

    tabs.forEach(tab => {
      tab.addEventListener('click', () => {
        tabs.forEach(t => t.classList.remove('active'));
        panels.forEach(p => p.classList.remove('active'));

        tab.classList.add('active');
        const targetId = tab.getAttribute('data-tab');
        const targetPanel = document.getElementById(targetId);
        if (targetPanel) {
          targetPanel.classList.add('active');
        }
      });
    });
  },

  /**
   * Backstage (File Menu) Open / Close and Sections
   */
  bindBackstage() {
    const fileTab = document.getElementById('tabFile');
    const backstage = document.getElementById('backstageOverlay');
    const backBtn = document.getElementById('backstageBackBtn');
    const navItems = document.querySelectorAll('.backstage-nav-item');
    const sections = document.querySelectorAll('.backstage-section');

    if (fileTab && backstage) {
      fileTab.addEventListener('click', () => {
        backstage.classList.add('active');
      });
    }

    if (backBtn && backstage) {
      backBtn.addEventListener('click', () => {
        backstage.classList.remove('active');
      });
    }

    navItems.forEach(item => {
      item.addEventListener('click', () => {
        navItems.forEach(n => n.classList.remove('active'));
        sections.forEach(s => s.style.display = 'none');

        item.classList.add('active');
        const sectionId = item.getAttribute('data-section');
        const sectionEl = document.getElementById(sectionId);
        if (sectionEl) {
          sectionEl.style.display = 'block';
        }
      });
    });

    // File Input for DOCX Import
    const fileInput = document.getElementById('docxFileInput');
    if (fileInput) {
      fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          DocxManager.importDocx(e.target.files[0]);
          if (backstage) backstage.classList.remove('active');
        }
      });
    }
  },

  /**
   * Bind All Ribbon Tools & Dropdowns
   */
  bindToolbarButtons() {
    // Quick Actions
    document.getElementById('btnQuickSave')?.addEventListener('click', () => this.saveToLocalStorage(true));
    document.getElementById('btnQuickUndo')?.addEventListener('click', () => Editor.exec('undo'));
    document.getElementById('btnQuickRedo')?.addEventListener('click', () => Editor.exec('redo'));
    document.getElementById('btnQuickPdf')?.addEventListener('click', () => {
      const name = document.getElementById('docTitleInput')?.value || 'document';
      PdfExporter.exportToPdf(name);
    });
    document.getElementById('btnQuickDocx')?.addEventListener('click', () => {
      const name = document.getElementById('docTitleInput')?.value || 'document';
      DocxManager.exportToDocx(name);
    });

    // Clipboard
    document.getElementById('btnUndo')?.addEventListener('click', () => Editor.exec('undo'));
    document.getElementById('btnRedo')?.addEventListener('click', () => Editor.exec('redo'));
    document.getElementById('btnCopy')?.addEventListener('click', () => Editor.exec('copy'));
    document.getElementById('btnCut')?.addEventListener('click', () => Editor.exec('cut'));

    // Font family & size
    const fontSelect = document.getElementById('fontFamilySelect');
    if (fontSelect) {
      fontSelect.addEventListener('change', (e) => {
        Editor.setFontFamily(e.target.value);
      });
    }

    const fontSizeSelect = document.getElementById('fontSizeSelect');
    if (fontSizeSelect) {
      fontSizeSelect.addEventListener('change', (e) => {
        Editor.setFontSize(e.target.value);
      });
    }

    document.getElementById('btnGrowFont')?.addEventListener('click', () => {
      const cur = parseInt(fontSizeSelect.value) || 14;
      fontSizeSelect.value = Math.min(cur + 2, 72);
      Editor.setFontSize(fontSizeSelect.value);
    });

    document.getElementById('btnShrinkFont')?.addEventListener('click', () => {
      const cur = parseInt(fontSizeSelect.value) || 14;
      fontSizeSelect.value = Math.max(cur - 2, 8);
      Editor.setFontSize(fontSizeSelect.value);
    });

    document.getElementById('btnClearFormat')?.addEventListener('click', () => Editor.clearFormatting());

    // Font styles
    document.getElementById('btnBold')?.addEventListener('click', () => Editor.exec('bold'));
    document.getElementById('btnItalic')?.addEventListener('click', () => Editor.exec('italic'));
    document.getElementById('btnUnderline')?.addEventListener('click', () => Editor.exec('underline'));
    document.getElementById('btnStrike')?.addEventListener('click', () => Editor.exec('strikeThrough'));
    document.getElementById('btnSubscript')?.addEventListener('click', () => Editor.exec('subscript'));
    document.getElementById('btnSuperscript')?.addEventListener('click', () => Editor.exec('superscript'));

    // Colors
    const textColorInput = document.getElementById('textColorInput');
    if (textColorInput) {
      textColorInput.addEventListener('input', (e) => Editor.setTextColor(e.target.value));
    }
    const highlightColorInput = document.getElementById('highlightColorInput');
    if (highlightColorInput) {
      highlightColorInput.addEventListener('input', (e) => Editor.setHighlightColor(e.target.value));
    }

    // Paragraph Alignment
    document.getElementById('btnAlignRight')?.addEventListener('click', () => Editor.exec('justifyRight'));
    document.getElementById('btnAlignCenter')?.addEventListener('click', () => Editor.exec('justifyCenter'));
    document.getElementById('btnAlignLeft')?.addEventListener('click', () => Editor.exec('justifyLeft'));
    document.getElementById('btnAlignJustify')?.addEventListener('click', () => Editor.exec('justifyFull'));

    // Direction
    document.getElementById('btnDirRtl')?.addEventListener('click', () => Editor.setTextDirection('rtl'));
    document.getElementById('btnDirLtr')?.addEventListener('click', () => Editor.setTextDirection('ltr'));

    // Lists & Indent
    document.getElementById('btnBulletList')?.addEventListener('click', () => Editor.exec('insertUnorderedList'));
    document.getElementById('btnNumberedList')?.addEventListener('click', () => Editor.exec('insertOrderedList'));
    document.getElementById('btnIndentInc')?.addEventListener('click', () => Editor.exec('indent'));
    document.getElementById('btnIndentDec')?.addEventListener('click', () => Editor.exec('outdent'));

    // Line Spacing
    const lineSpacingSelect = document.getElementById('lineSpacingSelect');
    if (lineSpacingSelect) {
      lineSpacingSelect.addEventListener('change', (e) => Editor.setLineSpacing(e.target.value));
    }

    // Heading Styles
    document.querySelectorAll('.style-card').forEach(card => {
      card.addEventListener('click', () => {
        const tag = card.getAttribute('data-tag');
        if (tag === 'p') {
          Editor.formatBlock('p');
        } else {
          Editor.formatBlock(tag);
        }
      });
    });

    // Insert Panel tools
    document.getElementById('btnInsertTableModal')?.addEventListener('click', () => {
      this.openModal('modalTable');
    });
    document.getElementById('btnInsertImage')?.addEventListener('click', () => {
      document.getElementById('imgUploadInput')?.click();
    });
    document.getElementById('imgUploadInput')?.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        Editor.insertImageFile(e.target.files[0]);
      }
    });

    document.getElementById('btnInsertPageBreak')?.addEventListener('click', () => Editor.insertPageBreak());
    document.getElementById('btnInsertHr')?.addEventListener('click', () => Editor.exec('insertHorizontalRule'));
    document.getElementById('btnInsertCalloutInfo')?.addEventListener('click', () => Editor.insertCallout('info'));
    document.getElementById('btnInsertCalloutWarn')?.addEventListener('click', () => Editor.insertCallout('warning'));

    // Insert Date & Time
    document.getElementById('btnInsertDateTime')?.addEventListener('click', () => {
      const now = new Date();
      const persianDate = new Intl.DateTimeFormat('fa-IR', { dateStyle: 'full', timeStyle: 'short' }).format(now);
      document.execCommand('insertHTML', false, `<span>${persianDate}</span>`);
    });

    // Layout Panel tools
    document.getElementById('btnMarginsNormal')?.addEventListener('click', () => Editor.setMargins('normal'));
    document.getElementById('btnMarginsNarrow')?.addEventListener('click', () => Editor.setMargins('narrow'));
    document.getElementById('btnMarginsWide')?.addEventListener('click', () => Editor.setMargins('wide'));
    document.getElementById('btnMarginsCustom')?.addEventListener('click', () => this.openModal('modalMargins'));

    document.getElementById('btnOrientPortrait')?.addEventListener('click', () => Editor.setOrientation('portrait'));
    document.getElementById('btnOrientLandscape')?.addEventListener('click', () => Editor.setOrientation('landscape'));

    document.getElementById('btnWatermarkModal')?.addEventListener('click', () => this.openModal('modalWatermark'));
    document.getElementById('btnPageBorderModal')?.addEventListener('click', () => this.openModal('modalBorder'));

    // View Panel tools
    document.getElementById('btnToggleRuler')?.addEventListener('click', () => {
      const ruler = document.querySelector('.ruler-top-container');
      if (ruler) {
        ruler.style.display = ruler.style.display === 'none' ? 'flex' : 'none';
      }
    });

    document.getElementById('btnFullscreen')?.addEventListener('click', () => {
      if (!document.fullscreenElement) {
        document.documentElement.requestFullscreen().catch(() => {});
      } else {
        document.exitFullscreen().catch(() => {});
      }
    });

    // Find & Replace
    document.getElementById('btnFindReplace')?.addEventListener('click', () => {
      const bar = document.getElementById('findReplaceBar');
      if (bar) bar.classList.toggle('active');
    });
    document.getElementById('btnFindClose')?.addEventListener('click', () => {
      document.getElementById('findReplaceBar')?.classList.remove('active');
    });
    document.getElementById('btnFindNext')?.addEventListener('click', () => {
      const q = document.getElementById('findInput')?.value;
      if (q) window.find(q);
    });
    document.getElementById('btnReplaceOne')?.addEventListener('click', () => {
      const q = document.getElementById('findInput')?.value;
      const r = document.getElementById('replaceInput')?.value || '';
      Editor.findAndReplace(q, r, false);
    });
    document.getElementById('btnReplaceAll')?.addEventListener('click', () => {
      const q = document.getElementById('findInput')?.value;
      const r = document.getElementById('replaceInput')?.value || '';
      Editor.findAndReplace(q, r, true);
    });

    // Document Title Rename
    const titleInput = document.getElementById('docTitleInput');
    if (titleInput) {
      titleInput.addEventListener('change', () => {
        document.title = `${titleInput.value} - Word آنلاین`;
        this.saveToLocalStorage(true);
      });
    }
  },

  /**
   * Modals Handling
   */
  bindModals() {
    // Close modal on close buttons
    document.querySelectorAll('.modal-close-btn, .modal-cancel-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('active'));
      });
    });

    // Table modal submit
    document.getElementById('btnConfirmTable')?.addEventListener('click', () => {
      const rows = parseInt(document.getElementById('tableRowsInput')?.value) || 3;
      const cols = parseInt(document.getElementById('tableColsInput')?.value) || 3;
      const styled = document.getElementById('tableStyleCheck')?.checked ?? true;
      Editor.insertTable(rows, cols, styled);
      this.closeModal('modalTable');
    });

    // Margins modal submit
    document.getElementById('btnConfirmMargins')?.addEventListener('click', () => {
      const top = parseFloat(document.getElementById('marginTopInput')?.value) || 2.54;
      const bottom = parseFloat(document.getElementById('marginBottomInput')?.value) || 2.54;
      const left = parseFloat(document.getElementById('marginLeftInput')?.value) || 2.54;
      const right = parseFloat(document.getElementById('marginRightInput')?.value) || 2.54;
      Editor.setMargins('custom', top, bottom, left, right);
      this.closeModal('modalMargins');
    });

    // Watermark modal submit
    document.getElementById('btnConfirmWatermark')?.addEventListener('click', () => {
      const text = document.getElementById('watermarkTextInput')?.value || 'محرمانه';
      const enable = document.getElementById('watermarkEnableCheck')?.checked ?? true;
      Editor.setWatermark(text, enable);
      this.closeModal('modalWatermark');
    });

    // Border modal submit
    document.getElementById('btnConfirmBorder')?.addEventListener('click', () => {
      const style = document.getElementById('borderStyleSelect')?.value || 'solid';
      const width = document.getElementById('borderWidthSelect')?.value || '2px';
      const color = document.getElementById('borderColorInput')?.value || '#185abd';
      Editor.setPageBorder(style, width, color);
      this.closeModal('modalBorder');
    });
  },

  openModal(id) {
    const el = document.getElementById(id);
    if (el) el.classList.add('active');
  },

  closeModal(id) {
    const el = document.getElementById(id);
    if (el) el.classList.remove('active');
  },

  /**
   * Zoom Slider
   */
  bindZoom() {
    const slider = document.getElementById('zoomSlider');
    const label = document.getElementById('zoomLabel');
    const container = document.querySelector('.page-container');

    if (slider && label && container) {
      slider.addEventListener('input', (e) => {
        const val = e.target.value;
        this.currentZoom = val;
        label.textContent = `${val}%`;
        container.style.transform = `scale(${val / 100})`;
      });

      label.addEventListener('click', () => {
        slider.value = 100;
        this.currentZoom = 100;
        label.textContent = '100%';
        container.style.transform = 'scale(1)';
      });
    }
  },

  /**
   * Keyboard Shortcuts
   */
  bindKeyboardShortcuts() {
    window.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey)) {
        if (e.key === 's' || e.key === 'S') {
          e.preventDefault();
          this.saveToLocalStorage(true);
        } else if (e.key === 'p' || e.key === 'P') {
          e.preventDefault();
          PdfExporter.printDocument();
        } else if (e.key === 'f' || e.key === 'F') {
          e.preventDefault();
          const bar = document.getElementById('findReplaceBar');
          if (bar) {
            bar.classList.add('active');
            document.getElementById('findInput')?.focus();
          }
        }
      }
    });
  },

  /**
   * Dark Mode Toggle
   */
  bindThemeToggle() {
    const btn = document.getElementById('btnThemeToggle');
    if (btn) {
      btn.addEventListener('click', () => {
        document.body.classList.toggle('dark-mode');
        const isDark = document.body.classList.contains('dark-mode');
        btn.innerHTML = isDark ? '<i data-lucide="sun"></i>' : '<i data-lucide="moon"></i>';
        if (typeof lucide !== 'undefined') lucide.createIcons();
      });
    }
  },

  /**
   * Update Document Statistics (Word count, Characters, Pages)
   */
  updateStats() {
    const body = document.querySelector('.page-body');
    const text = body ? (body.innerText || body.textContent || '') : '';

    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;

    const statWords = document.getElementById('statWords');
    const statChars = document.getElementById('statChars');
    const statPages = document.getElementById('statPages');

    if (statWords) statWords.textContent = `${words} کلمه`;
    if (statChars) statChars.textContent = `${chars} نویسه`;
    
    const pagesCount = document.querySelectorAll('.word-page').length;
    if (statPages) statPages.textContent = `صفحه ۱ از ${pagesCount}`;
  },

  /**
   * LocalStorage Auto-Save
   */
  debounceSave() {
    clearTimeout(this.saveTimeout);
    this.saveTimeout = setTimeout(() => {
      this.saveToLocalStorage(false);
    }, 1500);
  },

  saveToLocalStorage(notify = false) {
    const body = document.querySelector('.page-body');
    const title = document.getElementById('docTitleInput')?.value || 'سند بدون نام';
    const header = document.querySelector('.page-header-content')?.innerHTML || '';
    const footer = document.querySelector('.page-footer-content')?.innerHTML || '';

    const data = {
      title,
      header,
      footer,
      body: body ? body.innerHTML : '',
      savedAt: new Date().toISOString()
    };

    try {
      localStorage.setItem('word_online_document', JSON.stringify(data));
      const badge = document.getElementById('saveStatusBadge');
      if (badge) {
        badge.innerHTML = '<span>ذخیره شد ✓</span>';
      }
      if (notify) {
        alert('سند با موفقیت در مرورگر شما ذخیره شد.');
      }
    } catch (e) {
      console.warn('Storage quota exceeded or error:', e);
    }
  },

  loadInitialDocument() {
    const saved = localStorage.getItem('word_online_document');
    if (saved) {
      try {
        const data = JSON.parse(saved);
        if (data.title) {
          const tInput = document.getElementById('docTitleInput');
          if (tInput) tInput.value = data.title;
          document.title = `${data.title} - Word آنلاین`;
        }
        if (data.body) {
          const body = document.querySelector('.page-body');
          if (body) body.innerHTML = data.body;
        }
        if (data.header) {
          const h = document.querySelector('.page-header-content');
          if (h) h.innerHTML = data.header;
        }
        if (data.footer) {
          const f = document.querySelector('.page-footer-content');
          if (f) f.innerHTML = data.footer;
        }
        return;
      } catch (e) {}
    }

    // Default template: Official letter or blank
    this.applyTemplate('officialLetter');
  },

  applyTemplate(key) {
    const t = DocumentTemplates[key];
    if (!t) return;

    const tInput = document.getElementById('docTitleInput');
    if (tInput) tInput.value = t.title;
    document.title = `${t.title} - Word آنلاین`;

    const body = document.querySelector('.page-body');
    if (body) body.innerHTML = t.content;

    const h = document.querySelector('.page-header-content');
    if (h && t.header) h.innerHTML = t.header;

    const f = document.querySelector('.page-footer-content');
    if (f && t.footer) f.innerHTML = t.footer;

    const backstage = document.getElementById('backstageOverlay');
    if (backstage) backstage.classList.remove('active');

    this.updateStats();
    this.saveToLocalStorage(false);
  }
};

window.App = App;
window.addEventListener('DOMContentLoaded', () => {
  App.init();
});
