/* ==========================================================================
   Word Online Rich Text Editing Engine & Page Management
   ========================================================================== */

const Editor = {
  activePage: 1,
  totalPages: 1,

  init() {
    this.bindEvents();
    this.renderRulerTicks();
    this.updateRulerMargins(2.54, 2.54); // default normal margins 2.54cm
  },

  bindEvents() {
    // Listen to selection changes to update ribbon active buttons
    document.addEventListener('selectionchange', () => {
      this.updateRibbonState();
    });

    // Handle clicks inside page body
    const pageBody = document.querySelector('.page-body');
    if (pageBody) {
      pageBody.addEventListener('keyup', () => {
        if (window.App) {
          window.App.updateStats();
          window.App.debounceSave();
        }
      });
      pageBody.addEventListener('input', () => {
        if (window.App) {
          window.App.updateStats();
          window.App.debounceSave();
        }
      });
    }
  },

  /**
   * Execute standard rich text formatting command
   */
  exec(command, value = null) {
    document.execCommand(command, false, value);
    const pageBody = document.querySelector('.page-body');
    if (pageBody) {
      pageBody.focus();
    }
    this.updateRibbonState();
    if (window.App) {
      window.App.updateStats();
      window.App.debounceSave();
    }
  },

  /**
   * Format Block (Paragraph, Heading 1, 2, 3, Blockquote, Pre)
   */
  formatBlock(tag) {
    document.execCommand('formatBlock', false, `<${tag}>`);
    const pageBody = document.querySelector('.page-body');
    if (pageBody) {
      pageBody.focus();
    }
    this.updateRibbonState();
    if (window.App) {
      window.App.debounceSave();
    }
  },

  /**
   * Set Font Family
   */
  setFontFamily(font) {
    document.execCommand('fontName', false, font);
    const selection = window.getSelection();
    if (selection && selection.rangeCount > 0) {
      const range = selection.getRangeAt(0);
      const span = document.createElement('span');
      span.style.fontFamily = font;
    }
    if (window.App) window.App.debounceSave();
  },

  /**
   * Set Font Size in pt / px
   */
  setFontSize(pt) {
    // Use span with explicit font-size style for exact Word pt sizing
    const selection = window.getSelection();
    if (!selection.rangeCount) return;

    const range = selection.getRangeAt(0);
    if (range.collapsed) {
      // If nothing selected, format surrounding block or insert styled span
      document.execCommand('fontSize', false, '3');
      return;
    }

    const span = document.createElement('span');
    span.style.fontSize = `${pt}pt`;
    span.appendChild(range.extractContents());
    range.insertNode(span);

    // Re-select
    selection.removeAllRanges();
    const newRange = document.createRange();
    newRange.selectNodeContents(span);
    selection.addRange(newRange);

    if (window.App) window.App.debounceSave();
  },

  /**
   * Set Text Color
   */
  setTextColor(color) {
    document.execCommand('foreColor', false, color);
    const indicator = document.getElementById('textColorIndicator');
    if (indicator) indicator.style.backgroundColor = color;
    if (window.App) window.App.debounceSave();
  },

  /**
   * Set Text Highlight Background Color
   */
  setHighlightColor(color) {
    document.execCommand('hiliteColor', false, color);
    const indicator = document.getElementById('highlightColorIndicator');
    if (indicator) indicator.style.backgroundColor = color;
    if (window.App) window.App.debounceSave();
  },

  /**
   * Toggle Text Direction (LTR / RTL)
   */
  setTextDirection(dir) {
    const selection = window.getSelection();
    if (!selection.rangeCount) return;

    let node = selection.anchorNode;
    while (node && node !== document.querySelector('.page-body')) {
      if (node.nodeType === Node.ELEMENT_NODE && (node.tagName === 'P' || node.tagName.startsWith('H') || node.tagName === 'DIV' || node.tagName === 'LI')) {
        node.style.direction = dir;
        node.style.textAlign = dir === 'rtl' ? 'right' : 'left';
        if (window.App) window.App.debounceSave();
        return;
      }
      node = node.parentNode;
    }

    // Default to whole page body if no specific block found
    const body = document.querySelector('.page-body');
    if (body) {
      body.style.direction = dir;
      body.style.textAlign = dir === 'rtl' ? 'right' : 'left';
      if (window.App) window.App.debounceSave();
    }
  },

  /**
   * Set Line Height / Line Spacing
   */
  setLineSpacing(lineHeight) {
    const selection = window.getSelection();
    let node = selection.anchorNode;
    while (node && node !== document.querySelector('.page-body')) {
      if (node.nodeType === Node.ELEMENT_NODE && (node.tagName === 'P' || node.tagName.startsWith('H') || node.tagName === 'DIV')) {
        node.style.lineHeight = lineHeight;
        if (window.App) window.App.debounceSave();
        return;
      }
      node = node.parentNode;
    }
    const pageBody = document.querySelector('.page-body');
    if (pageBody) {
      pageBody.style.lineHeight = lineHeight;
      if (window.App) window.App.debounceSave();
    }
  },

  /**
   * Clear Formatting
   */
  clearFormatting() {
    document.execCommand('removeFormat', false, null);
    if (window.App) window.App.debounceSave();
  },

  /**
   * Insert Table with given rows and cols
   */
  insertTable(rows = 3, cols = 3, styled = true) {
    let html = `<table class="${styled ? 'word-styled-table' : ''}"><tbody>`;
    for (let r = 0; r < rows; r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) {
        if (r === 0 && styled) {
          html += `<th>سرستون ${c + 1}</th>`;
        } else {
          html += `<td>سلول ${r + 1}-${c + 1}</td>`;
        }
      }
      html += '</tr>';
    }
    html += '</tbody></table><p><br></p>';

    document.execCommand('insertHTML', false, html);
    if (window.App) window.App.debounceSave();
  },

  /**
   * Table Operations
   */
  addTableRow() {
    const sel = window.getSelection();
    let td = sel.anchorNode;
    while (td && td.tagName !== 'TD' && td.tagName !== 'TH') {
      td = td.parentNode;
    }
    if (!td) {
      alert('لطفاً ابتدا داخل یکی از خانه‌های جدول کلیک کنید.');
      return;
    }
    const tr = td.parentNode;
    const table = tr.parentNode;
    const colCount = tr.children.length;
    const newTr = document.createElement('tr');
    for (let i = 0; i < colCount; i++) {
      const newTd = document.createElement('td');
      newTd.innerHTML = '&nbsp;';
      newTr.appendChild(newTd);
    }
    tr.after(newTr);
    if (window.App) window.App.debounceSave();
  },

  deleteTableRow() {
    const sel = window.getSelection();
    let td = sel.anchorNode;
    while (td && td.tagName !== 'TD' && td.tagName !== 'TH') {
      td = td.parentNode;
    }
    if (!td) {
      alert('لطفاً ابتدا داخل یکی از خانه‌های جدول کلیک کنید.');
      return;
    }
    const tr = td.parentNode;
    tr.remove();
    if (window.App) window.App.debounceSave();
  },

  addTableColumn() {
    const sel = window.getSelection();
    let td = sel.anchorNode;
    while (td && td.tagName !== 'TD' && td.tagName !== 'TH') {
      td = td.parentNode;
    }
    if (!td) {
      alert('لطفاً ابتدا داخل یکی از خانه‌های جدول کلیک کنید.');
      return;
    }
    const colIndex = Array.from(td.parentNode.children).indexOf(td);
    const table = td.closest('table');
    table.querySelectorAll('tr').forEach(tr => {
      const isHeader = tr.children[0]?.tagName === 'TH';
      const newCell = document.createElement(isHeader ? 'th' : 'td');
      newCell.innerHTML = isHeader ? 'ستون جدید' : '&nbsp;';
      if (tr.children[colIndex]) {
        tr.children[colIndex].after(newCell);
      } else {
        tr.appendChild(newCell);
      }
    });
    if (window.App) window.App.debounceSave();
  },

  deleteTableColumn() {
    const sel = window.getSelection();
    let td = sel.anchorNode;
    while (td && td.tagName !== 'TD' && td.tagName !== 'TH') {
      td = td.parentNode;
    }
    if (!td) {
      alert('لطفاً ابتدا داخل یکی از خانه‌های جدول کلیک کنید.');
      return;
    }
    const colIndex = Array.from(td.parentNode.children).indexOf(td);
    const table = td.closest('table');
    table.querySelectorAll('tr').forEach(tr => {
      if (tr.children[colIndex]) {
        tr.children[colIndex].remove();
      }
    });
    if (window.App) window.App.debounceSave();
  },

  /**
   * Insert Callout Box
   */
  insertCallout(type = 'info') {
    const icons = { info: '💡', warning: '⚠️', success: '✅' };
    const titles = { info: 'یادداشت مهم:', warning: 'هشدار و توجه:', success: 'تأییدیه و نکته:' };
    const html = `
      <div class="word-callout ${type}">
        <span style="font-size: 18px;">${icons[type] || '💡'}</span>
        <div><strong>${titles[type] || ''}</strong> متن یادداشت خود را در اینجا بنویسید...</div>
      </div>
      <p><br></p>
    `;
    document.execCommand('insertHTML', false, html);
    if (window.App) window.App.debounceSave();
  },

  /**
   * Insert Image from File
   */
  insertImageFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      const imgHtml = `
        <div class="doc-img-wrapper" style="text-align: center;">
          <img src="${e.target.result}" style="max-width: 80%; border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.15);" alt="تصویر درج شده">
        </div>
        <p><br></p>
      `;
      document.execCommand('insertHTML', false, imgHtml);
      if (window.App) window.App.debounceSave();
    };
    reader.readAsDataURL(file);
  },

  /**
   * Insert Page Break
   */
  insertPageBreak() {
    // Create new physical word page in container
    this.addNewPage();
    if (window.App) {
      window.App.updateStats();
      window.App.debounceSave();
    }
  },

  /**
   * Add a new A4 physical page to the document
   */
  addNewPage() {
    const pageContainer = document.querySelector('.page-container');
    if (!pageContainer) return;

    this.totalPages++;
    const newPage = document.createElement('div');
    newPage.className = 'word-page';
    newPage.setAttribute('data-page', this.totalPages);

    // Copy orientation from first page
    const firstPage = pageContainer.querySelector('.word-page');
    if (firstPage && firstPage.classList.contains('landscape')) {
      newPage.classList.add('landscape');
    }

    const headerContent = document.querySelector('.page-header-content')?.innerHTML || '<span>مایکروسافت ورد آنلاین</span>';
    const footerContent = `
      <div class="page-footer-content">
        <span>سند جاری</span>
        <span>صفحه ${this.totalPages}</span>
      </div>
    `;

    newPage.innerHTML = `
      <div class="page-header-zone" contenteditable="true">
        <div class="page-header-content">${headerContent}</div>
      </div>
      <div class="page-watermark ${document.querySelector('.page-watermark.active') ? 'active' : ''}">
        ${document.querySelector('.page-watermark')?.textContent || 'محرمانه'}
      </div>
      <div class="page-body" contenteditable="true">
        <p><br></p>
      </div>
      <div class="page-footer-zone" contenteditable="true">
        ${footerContent}
      </div>
    `;

    pageContainer.appendChild(newPage);

    // Focus new page
    const newBody = newPage.querySelector('.page-body');
    if (newBody) {
      newBody.focus();
      newBody.addEventListener('input', () => {
        if (window.App) {
          window.App.updateStats();
          window.App.debounceSave();
        }
      });
    }

    if (window.App) window.App.updateStats();
  },

  /**
   * Page Margins Setup
   */
  setMargins(type, topCm = 2.54, bottomCm = 2.54, leftCm = 2.54, rightCm = 2.54) {
    const marginPresets = {
      normal: { top: 2.54, bottom: 2.54, left: 2.54, right: 2.54 },
      narrow: { top: 1.27, bottom: 1.27, left: 1.27, right: 1.27 },
      moderate: { top: 2.54, bottom: 2.54, left: 1.9, right: 1.9 },
      wide: { top: 2.54, bottom: 2.54, left: 3.18, right: 3.18 }
    };

    let m = marginPresets[type] || { top: topCm, bottom: bottomCm, left: leftCm, right: rightCm };

    const cmToPx = 37.795; // 96 DPI / 2.54 cm

    const pages = document.querySelectorAll('.word-page');
    pages.forEach(page => {
      const headerZone = page.querySelector('.page-header-zone');
      const footerZone = page.querySelector('.page-footer-zone');
      const bodyZone = page.querySelector('.page-body');

      if (headerZone) {
        headerZone.style.paddingLeft = `${m.left * cmToPx}px`;
        headerZone.style.paddingRight = `${m.right * cmToPx}px`;
        headerZone.style.paddingTop = `${m.top * cmToPx * 0.6}px`;
      }
      if (footerZone) {
        footerZone.style.paddingLeft = `${m.left * cmToPx}px`;
        footerZone.style.paddingRight = `${m.right * cmToPx}px`;
        footerZone.style.paddingBottom = `${m.bottom * cmToPx * 0.6}px`;
      }
      if (bodyZone) {
        bodyZone.style.paddingLeft = `${m.left * cmToPx}px`;
        bodyZone.style.paddingRight = `${m.right * cmToPx}px`;
      }
    });

    this.updateRulerMargins(m.left, m.right);
    if (window.App) window.App.debounceSave();
  },

  /**
   * Page Orientation
   */
  setOrientation(orientation) {
    const pages = document.querySelectorAll('.word-page');
    pages.forEach(p => {
      if (orientation === 'landscape') {
        p.classList.add('landscape');
      } else {
        p.classList.remove('landscape');
      }
    });
    this.renderRulerTicks();
    if (window.App) window.App.debounceSave();
  },

  /**
   * Watermark
   */
  setWatermark(text, active = true) {
    const watermarks = document.querySelectorAll('.page-watermark');
    watermarks.forEach(wm => {
      wm.textContent = text || 'محرمانه';
      if (active && text) {
        wm.classList.add('active');
      } else {
        wm.classList.remove('active');
      }
    });
    if (window.App) window.App.debounceSave();
  },

  /**
   * Page Border
   */
  setPageBorder(style = 'solid', width = '2px', color = '#185abd') {
    const pages = document.querySelectorAll('.word-page');
    pages.forEach(p => {
      if (style === 'none') {
        p.style.outline = 'none';
      } else {
        p.style.outline = `${width} ${style} ${color}`;
        p.style.outlineOffset = '-12px';
      }
    });
    if (window.App) window.App.debounceSave();
  },

  /**
   * Find & Replace
   */
  findAndReplace(findText, replaceText, replaceAll = false) {
    if (!findText) return;
    const body = document.querySelector('.page-body');
    if (!body) return;

    if (replaceAll) {
      const regex = new RegExp(findText, 'g');
      body.innerHTML = body.innerHTML.replace(regex, replaceText);
      alert(`تمام موارد "${findText}" با "${replaceText}" جایگزین شدند.`);
    } else {
      const found = window.find(findText, false, false, true, false, false, false);
      if (found) {
        document.execCommand('insertHTML', false, replaceText);
      } else {
        alert(`عبارت "${findText}" پیدا نشد.`);
      }
    }
    if (window.App) window.App.debounceSave();
  },

  /**
   * Render Interactive Ruler Centimeter Ticks
   */
  renderRulerTicks() {
    const ruler = document.getElementById('rulerTop');
    if (!ruler) return;
    ruler.innerHTML = '';

    const firstPage = document.querySelector('.word-page');
    const isLandscape = firstPage && firstPage.classList.contains('landscape');
    const maxCm = isLandscape ? 29.7 : 21.0; // A4 dimensions
    const totalTicks = Math.floor(maxCm);

    const ticksContainer = document.createElement('div');
    ticksContainer.className = 'ruler-ticks';
    ticksContainer.style.width = '100%';

    const cmPx = 37.795; // 1cm in px

    for (let cm = 0; cm <= totalTicks; cm++) {
      const tick = document.createElement('div');
      tick.className = 'ruler-tick major';
      tick.style.right = `${cm * cmPx}px`;

      const label = document.createElement('span');
      label.className = 'ruler-tick-label';
      label.textContent = cm;
      tick.appendChild(label);

      ticksContainer.appendChild(tick);

      // Half cm ticks
      if (cm < totalTicks) {
        const halfTick = document.createElement('div');
        halfTick.className = 'ruler-tick';
        halfTick.style.right = `${(cm + 0.5) * cmPx}px`;
        ticksContainer.appendChild(halfTick);
      }
    }

    ruler.appendChild(ticksContainer);
  },

  /**
   * Update Ruler Margin Highlight Zones
   */
  updateRulerMargins(leftCm, rightCm) {
    const ruler = document.getElementById('rulerTop');
    if (!ruler) return;

    // Remove existing margin overlays
    ruler.querySelectorAll('.ruler-margin-zone').forEach(z => z.remove());

    const cmPx = 37.795;

    // Right margin zone (in RTL, starts at 0 to rightCm)
    const rightZone = document.createElement('div');
    rightZone.className = 'ruler-margin-zone';
    rightZone.style.right = '0';
    rightZone.style.width = `${rightCm * cmPx}px`;

    // Left margin zone
    const leftZone = document.createElement('div');
    leftZone.className = 'ruler-margin-zone';
    leftZone.style.left = '0';
    leftZone.style.width = `${leftCm * cmPx}px`;

    ruler.appendChild(rightZone);
    ruler.appendChild(leftZone);
  },

  /**
   * Update Ribbon active button state based on selection
   */
  updateRibbonState() {
    const btnBold = document.getElementById('btnBold');
    const btnItalic = document.getElementById('btnItalic');
    const btnUnderline = document.getElementById('btnUnderline');
    const btnStrike = document.getElementById('btnStrike');

    if (btnBold) btnBold.classList.toggle('active', document.queryCommandState('bold'));
    if (btnItalic) btnItalic.classList.toggle('active', document.queryCommandState('italic'));
    if (btnUnderline) btnUnderline.classList.toggle('active', document.queryCommandState('underline'));
    if (btnStrike) btnStrike.classList.toggle('active', document.queryCommandState('strikeThrough'));
  }
};
