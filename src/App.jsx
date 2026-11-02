/* ==========================================================================
   Microsoft Word Online 365 (Vite + React Complete Home Tab & Context Menu)
   ========================================================================== */

import React, { useState, useEffect, useRef, useCallback } from 'react';
import { DocumentTemplates } from './templates.js';

export default function WordApp() {
  // Navigation & View State
  const [activeTab, setActiveTab] = useState('home');
  const [docTitle, setDocTitle] = useState('نامه اداری رسمی');
  const [saveStatus, setSaveStatus] = useState('ذخیره شده ✓');
  const [zoom, setZoom] = useState(100);
  const [isDark, setIsDark] = useState(false);
  const [showOutline, setShowOutline] = useState(false);
  const [headings, setHeadings] = useState([]);

  // Multi-Page System (True Multiple A4 Pages)
  const [pages, setPages] = useState([{ id: 1, content: '' }]);
  const [activePageIndex, setActivePageIndex] = useState(0);

  // Document Page Setup State
  const [margins, setMargins] = useState({ top: 2.54, bottom: 2.54, left: 2.54, right: 2.54, type: 'normal' });
  const [orientation, setOrientation] = useState('portrait');
  const [columns, setColumns] = useState(1);
  const [paperTheme, setPaperTheme] = useState('white');
  const [watermark, setWatermark] = useState({ enabled: false, text: 'محرمانه' });
  const [pageBorder, setPageBorder] = useState({ style: 'none', width: '2px', color: '#185abd' });
  const [showPilcrow, setShowPilcrow] = useState(false);

  // Format Painter & Editor State
  const [formatPainterActive, setFormatPainterActive] = useState(false);
  const [copiedStyle, setCopiedStyle] = useState(null);
  const [activeStyles, setActiveStyles] = useState({ bold: false, italic: false, underline: false, strike: false });
  const [fontFamily, setFontFamily] = useState('Calibri');
  const [fontSize, setFontSize] = useState('11');
  const [fontColor, setFontColor] = useState('#111827');
  const [highlightColor, setHighlightColor] = useState('#fef08a');
  const [shadingColor, setShadingColor] = useState('#f1f5f9');

  // Ribbon Dropdown State
  const [openDropdown, setOpenDropdown] = useState(null); // 'underline' | 'case' | 'textEffects' | 'fontColor' | 'highlight' | 'bullets' | 'numbering' | 'spacing' | 'shading' | 'borders' | 'justify' | 'paste' | 'find' | 'select'

  // Custom Word Context Menu State
  const [contextMenu, setContextMenu] = useState({ visible: false, x: 0, y: 0, targetNode: null });

  // Modals & Floating Toolbars
  const [activeModal, setActiveModal] = useState(null); // 'table' | 'margins' | 'watermark' | 'border' | 'stats' | 'symbols' | 'history' | 'goal' | 'font' | 'paragraph'
  const [showFindReplace, setShowFindReplace] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [replaceTerm, setReplaceTerm] = useState('');
  
  // Table Modal Inputs
  const [tableRows, setTableRows] = useState(3);
  const [tableCols, setTableCols] = useState(3);
  const [tableStyled, setTableStyled] = useState(true);

  // Custom Margins Inputs
  const [customTop, setCustomTop] = useState(2.54);
  const [customBottom, setCustomBottom] = useState(2.54);
  const [customLeft, setCustomLeft] = useState(2.54);
  const [customRight, setCustomRight] = useState(2.54);

  // Watermark Inputs
  const [watermarkText, setWatermarkText] = useState('محرمانه');
  const [watermarkEnabled, setWatermarkEnabled] = useState(true);

  // Border Inputs
  const [borderStyle, setBorderStyle] = useState('solid');
  const [borderWidth, setBorderWidth] = useState('2px');
  const [borderColor, setBorderColor] = useState('#185abd');

  // Word Count Goal
  const [wordGoal, setWordGoal] = useState(500);

  // Document Statistics
  const [stats, setStats] = useState({
    words: 0,
    chars: 0,
    charsNoSpaces: 0,
    paragraphs: 0,
    lines: 0,
    readTime: '۱ دقیقه',
    speakTime: '۱ دقیقه'
  });

  // Version History / Snapshots
  const [snapshots, setSnapshots] = useState([]);
  const [snapshotName, setSnapshotName] = useState('');

  // Refs
  const pageRefs = useRef({});
  const headerContentRef = useRef(null);
  const footerContentRef = useRef(null);
  const fileInputDocxRef = useRef(null);
  const imgInputRef = useRef(null);
  const stylesGalleryRef = useRef(null);

  // Close dropdowns and context menu on outside click
  useEffect(() => {
    const handleClickOutside = (e) => {
      if (!e.target.closest('.word-dropdown-menu') && !e.target.closest('.has-dropdown') && !e.target.closest('.split-btn-arrow')) {
        setOpenDropdown(null);
      }
      if (!e.target.closest('.word-context-menu')) {
        setContextMenu({ visible: false, x: 0, y: 0, targetNode: null });
      }
    };
    window.addEventListener('click', handleClickOutside);
    return () => window.removeEventListener('click', handleClickOutside);
  }, []);

  // Load Saved Document on Mount
  useEffect(() => {
    const saved = localStorage.getItem('word_react_document_v3');
    if (saved) {
      try {
        const data = JSON.parse(saved);
        if (data.title) setDocTitle(data.title);
        if (data.pages && Array.isArray(data.pages)) setPages(data.pages);
        if (data.header && headerContentRef.current) headerContentRef.current.innerHTML = data.header;
        if (data.footer && footerContentRef.current) footerContentRef.current.innerHTML = data.footer;
        if (data.margins) setMargins(data.margins);
        if (data.orientation) setOrientation(data.orientation);
        if (data.columns) setColumns(data.columns);
        if (data.paperTheme) setPaperTheme(data.paperTheme);
        if (data.watermark) setWatermark(data.watermark);
        if (data.pageBorder) setPageBorder(data.pageBorder);
        if (data.wordGoal) setWordGoal(data.wordGoal);
      } catch (e) {
        console.warn('Failed to parse saved document', e);
      }
    } else {
      applyTemplate('officialLetter');
    }

    const savedSnapshots = localStorage.getItem('word_react_snapshots');
    if (savedSnapshots) {
      try { setSnapshots(JSON.parse(savedSnapshots)); } catch (e) {}
    }

    updateDocumentStats();
    updateOutline();
  }, []);

  // Listen for DocxManager Events and Keyboard Shortcuts (Ctrl+S, Ctrl+P)
  useEffect(() => {
    const handleDocxImported = (e) => {
      const { html, title } = e.detail;
      if (title) setDocTitle(title);
      const newPageId = Date.now();
      const updatedPages = [{ id: newPageId, content: html }];
      setPages(updatedPages);
      setActivePageIndex(0);
      setTimeout(() => {
        if (pageRefs.current[newPageId]) {
          pageRefs.current[newPageId].innerHTML = html;
        }
        onEditorInput();
      }, 60);
      alert(`فایل ورد "${title}" با موفقیت بارگذاری شد و در حال حاضر آماده ویرایش است.`);
    };

    const handleProjectImported = (e) => {
      const data = e.detail;
      if (data.title) setDocTitle(data.title);
      if (data.pages && Array.isArray(data.pages)) setPages(data.pages);
      if (data.margins) setMargins(data.margins);
      if (data.orientation) setOrientation(data.orientation);
      if (data.paperTheme) setPaperTheme(data.paperTheme);
      if (data.watermark) setWatermark(data.watermark);
      if (data.columns) setColumns(data.columns);
      if (data.pageBorder) setPageBorder(data.pageBorder);
      if (data.wordGoal) setWordGoal(data.wordGoal);
      if (data.header && headerContentRef.current) headerContentRef.current.innerHTML = data.header;
      if (data.footer && footerContentRef.current) footerContentRef.current.innerHTML = data.footer;
      setTimeout(() => {
        if (data.pages && Array.isArray(data.pages)) {
          data.pages.forEach(p => {
            if (pageRefs.current[p.id]) pageRefs.current[p.id].innerHTML = p.content;
          });
        }
        onEditorInput();
      }, 60);
      alert(`پروژه "${data.title || 'سند'}" با موفقیت بازگردانی شد.`);
    };

    const handleKeyDown = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        saveToStorage(true);
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
        e.preventDefault();
        exportPDF();
      }
    };

    window.addEventListener('word-docx-imported', handleDocxImported);
    window.addEventListener('word-project-imported', handleProjectImported);
    window.addEventListener('keydown', handleKeyDown);

    return () => {
      window.removeEventListener('word-docx-imported', handleDocxImported);
      window.removeEventListener('word-project-imported', handleProjectImported);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [pages, docTitle, margins, orientation, paperTheme, watermark, columns, pageBorder, wordGoal]);

  // Update Lucide Icons when view updates
  useEffect(() => {
    if (window.lucide) {
      window.lucide.createIcons();
    }
  });

  // Update Ribbon State on Selection Change
  const handleSelectionChange = useCallback(() => {
    setActiveStyles({
      bold: document.queryCommandState('bold'),
      italic: document.queryCommandState('italic'),
      underline: document.queryCommandState('underline'),
      strike: document.queryCommandState('strikeThrough')
    });
  }, []);

  useEffect(() => {
    document.addEventListener('selectionchange', handleSelectionChange);
    return () => document.removeEventListener('selectionchange', handleSelectionChange);
  }, [handleSelectionChange]);

  // Auto-Save with Debounce
  const debounceSave = useCallback(() => {
    setSaveStatus('در حال ذخیره...');
    clearTimeout(window.__saveTimer);
    window.__saveTimer = setTimeout(() => {
      saveToStorage(false);
      setSaveStatus('ذخیره شده ✓');
    }, 1200);
  }, [docTitle, margins, orientation, columns, paperTheme, watermark, pageBorder, pages, wordGoal]);

  const saveToStorage = (notify = false) => {
    const updatedPages = pages.map(p => {
      const el = pageRefs.current[p.id];
      return { id: p.id, content: el ? el.innerHTML : p.content };
    });

    const data = {
      title: docTitle,
      pages: updatedPages,
      header: headerContentRef.current ? headerContentRef.current.innerHTML : '',
      footer: footerContentRef.current ? footerContentRef.current.innerHTML : '',
      margins,
      orientation,
      columns,
      paperTheme,
      watermark,
      pageBorder,
      wordGoal,
      updatedAt: new Date().toISOString()
    };
    localStorage.setItem('word_react_document_v3', JSON.stringify(data));
    if (notify) alert('سند با موفقیت در حافظه مرورگر ذخیره گردید.');
  };

  // Update Word / Character Statistics & Outline
  const updateDocumentStats = () => {
    let fullText = '';
    Object.values(pageRefs.current).forEach(el => {
      if (el) fullText += (el.innerText || '') + ' ';
    });

    const cleanWords = fullText.trim() ? fullText.trim().split(/\s+/).filter(Boolean) : [];
    const wordsCount = cleanWords.length;
    const charsCount = fullText.length;
    const charsNoSpaces = fullText.replace(/\s/g, '').length;
    const paragraphsCount = fullText.split(/\n+/).filter(Boolean).length || 1;
    const linesCount = Math.ceil(charsCount / 70) || 1;
    const readMinutes = Math.max(1, Math.ceil(wordsCount / 180));
    const speakMinutes = Math.max(1, Math.ceil(wordsCount / 130));

    setStats({
      words: wordsCount,
      chars: charsCount,
      charsNoSpaces,
      paragraphs: paragraphsCount,
      lines: linesCount,
      readTime: `${readMinutes} دقیقه`,
      speakTime: `${speakMinutes} دقیقه`
    });
  };

  const updateOutline = () => {
    const foundHeadings = [];
    Object.values(pageRefs.current).forEach(pageEl => {
      if (!pageEl) return;
      const elements = pageEl.querySelectorAll('h1, h2, h3');
      elements.forEach((el, idx) => {
        if (!el.id) el.id = `heading_${Date.now()}_${idx}`;
        foundHeadings.push({
          id: el.id,
          tag: el.tagName.toLowerCase(),
          text: el.innerText.trim() || 'عنوان بدون متن'
        });
      });
    });
    setHeadings(foundHeadings);
  };

  const onEditorInput = () => {
    updateDocumentStats();
    updateOutline();
    debounceSave();
  };

  // ==========================================
  // Custom Word Context Menu Event Handler
  // ==========================================
  const handleContextMenu = (e) => {
    e.preventDefault();
    const x = Math.min(e.clientX, window.innerWidth - 280);
    const y = Math.min(e.clientY, window.innerHeight - 380);
    setContextMenu({
      visible: true,
      x,
      y,
      targetNode: e.target
    });
  };

  // Add / Remove Pages
  const addNewPage = () => {
    const newId = Date.now();
    const updated = [...pages, { id: newId, content: '<p><br></p>' }];
    setPages(updated);
    setActivePageIndex(updated.length - 1);
    setTimeout(() => {
      const newEl = pageRefs.current[newId];
      if (newEl) newEl.focus();
    }, 100);
    debounceSave();
  };

  const removePage = (id) => {
    if (pages.length <= 1) {
      alert('سند باید حداقل شامل یک صفحه باشد.');
      return;
    }
    if (confirm('آیا از حذف این صفحه اطمینان دارید؟')) {
      const updated = pages.filter(p => p.id !== id);
      setPages(updated);
      delete pageRefs.current[id];
      debounceSave();
    }
  };

  // Rich Text Commands
  const execCmd = (command, value = null) => {
    document.execCommand(command, false, value);
    handleSelectionChange();
    onEditorInput();
  };

  const applyFormatBlock = (tag) => {
    document.execCommand('formatBlock', false, `<${tag}>`);
    onEditorInput();
  };

  const changeFontFamily = (font) => {
    setFontFamily(font);
    document.execCommand('fontName', false, font);
    debounceSave();
  };

  const changeFontSize = (sizePt) => {
    setFontSize(sizePt);
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    if (!range.collapsed) {
      const span = document.createElement('span');
      span.style.fontSize = `${sizePt}pt`;
      span.appendChild(range.extractContents());
      range.insertNode(span);
      sel.removeAllRanges();
      const newRange = document.createRange();
      newRange.selectNodeContents(span);
      sel.addRange(newRange);
    }
    debounceSave();
  };

  // Format Painter Feature
  const handleFormatPainter = () => {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) {
      alert('لطفاً ابتدا بخشی از متن الگو را انتخاب کنید.');
      return;
    }
    const parent = sel.anchorNode.parentElement;
    if (parent) {
      const style = window.getComputedStyle(parent);
      setCopiedStyle({
        fontFamily: style.fontFamily,
        fontSize: style.fontSize,
        color: style.color,
        fontWeight: style.fontWeight,
        fontStyle: style.fontStyle,
        textDecoration: style.textDecoration,
        backgroundColor: style.backgroundColor
      });
      setFormatPainterActive(true);
    }
  };

  const applyCopiedFormat = () => {
    if (!formatPainterActive || !copiedStyle) return;
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;

    const range = sel.getRangeAt(0);
    const span = document.createElement('span');
    span.style.fontFamily = copiedStyle.fontFamily;
    span.style.fontSize = copiedStyle.fontSize;
    span.style.color = copiedStyle.color;
    span.style.fontWeight = copiedStyle.fontWeight;
    span.style.fontStyle = copiedStyle.fontStyle;
    span.style.textDecoration = copiedStyle.textDecoration;
    if (copiedStyle.backgroundColor && copiedStyle.backgroundColor !== 'rgba(0, 0, 0, 0)') {
      span.style.backgroundColor = copiedStyle.backgroundColor;
    }

    span.appendChild(range.extractContents());
    range.insertNode(span);

    setFormatPainterActive(false);
    setCopiedStyle(null);
    onEditorInput();
  };

  // Underline Styles
  const applyUnderlineStyle = (style) => {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) {
      execCmd('underline');
      setOpenDropdown(null);
      return;
    }
    const range = sel.getRangeAt(0);
    const span = document.createElement('span');
    span.style.textDecoration = `underline ${style}`;
    span.appendChild(range.extractContents());
    range.insertNode(span);
    setOpenDropdown(null);
    onEditorInput();
  };

  // Case Conversion
  const changeCase = (type) => {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    let text = sel.toString();
    if (type === 'upper') text = text.toUpperCase();
    else if (type === 'lower') text = text.toLowerCase();
    else if (type === 'capitalize') {
      text = text.replace(/\b\w/g, l => l.toUpperCase());
    } else if (type === 'sentence') {
      text = text.charAt(0).toUpperCase() + text.slice(1).toLowerCase();
    } else if (type === 'toggle') {
      text = text.split('').map(c => c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()).join('');
    }
    document.execCommand('insertText', false, text);
    setOpenDropdown(null);
    onEditorInput();
  };

  // Text Effects (Shadow, Glow, Reflection)
  const applyTextEffect = (effect) => {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    const range = sel.getRangeAt(0);
    const span = document.createElement('span');

    if (effect === 'shadow') span.style.textShadow = '2px 2px 4px rgba(0,0,0,0.4)';
    else if (effect === 'glow') span.style.textShadow = '0 0 8px rgba(37,99,235,0.8), 0 0 12px rgba(37,99,235,0.4)';
    else if (effect === 'outline') span.style.webkitTextStroke = '1px #103f91';
    else if (effect === '3d') span.style.textShadow = '1px 1px #fff, 2px 2px #0c2d68';

    span.appendChild(range.extractContents());
    range.insertNode(span);
    setOpenDropdown(null);
    onEditorInput();
  };

  // Bullets Styles
  const applyBulletStyle = (type) => {
    execCmd('insertUnorderedList');
    const sel = window.getSelection();
    let ul = sel.anchorNode;
    while (ul && ul.tagName !== 'UL') ul = ul.parentNode;
    if (ul) ul.style.listStyleType = type;
    setOpenDropdown(null);
    onEditorInput();
  };

  // Numbering Styles
  const applyNumberingStyle = (type) => {
    execCmd('insertOrderedList');
    const sel = window.getSelection();
    let ol = sel.anchorNode;
    while (ol && ol.tagName !== 'OL') ol = ol.parentNode;
    if (ol) ol.style.listStyleType = type;
    setOpenDropdown(null);
    onEditorInput();
  };

  // Paragraph Shading Color
  const applyShadingColor = (color) => {
    setShadingColor(color);
    const sel = window.getSelection();
    let block = sel.anchorNode;
    while (block && block.tagName !== 'P' && !block.tagName?.startsWith('H') && block.tagName !== 'DIV') {
      block = block.parentNode;
    }
    if (block) {
      block.style.backgroundColor = color;
      block.style.padding = '4px 8px';
      block.style.borderRadius = '3px';
    }
    setOpenDropdown(null);
    onEditorInput();
  };

  // Paragraph Borders
  const applyParagraphBorder = (borderType) => {
    const sel = window.getSelection();
    let block = sel.anchorNode;
    while (block && block.tagName !== 'P' && !block.tagName?.startsWith('H') && block.tagName !== 'DIV') {
      block = block.parentNode;
    }
    if (block) {
      block.style.padding = '6px 10px';
      if (borderType === 'bottom') block.style.borderBottom = '1px solid #111827';
      else if (borderType === 'top') block.style.borderTop = '1px solid #111827';
      else if (borderType === 'left') block.style.borderLeft = '2px solid #185abd';
      else if (borderType === 'right') block.style.borderRight = '2px solid #185abd';
      else if (borderType === 'all') block.style.border = '1px solid #94a3b8';
      else if (borderType === 'none') block.style.border = 'none';
    }
    setOpenDropdown(null);
    onEditorInput();
  };

  // Sort Paragraphs A-Z
  const sortParagraphs = () => {
    const activeEl = pageRefs.current[pages[activePageIndex]?.id];
    if (!activeEl) return;
    const paras = Array.from(activeEl.querySelectorAll('p, li'));
    if (paras.length <= 1) return;
    paras.sort((a, b) => a.innerText.localeCompare(b.innerText, 'fa'));
    paras.forEach(p => activeEl.appendChild(p));
    alert('پاراگراف‌ها بر اساس حروف الفبا مرتب شدند.');
    onEditorInput();
  };

  // Persian Typography Normalizer
  const normalizePersianTypography = () => {
    pages.forEach(p => {
      const el = pageRefs.current[p.id];
      if (!el) return;
      let html = el.innerHTML;
      html = html.replace(/\s+(می|نمی)\s+/g, ' $1‌');
      html = html.replace(/\s+(ها|های|هایم|هایت|هایش|هایمان|هایتان|هایشان)\b/g, '‌$1');
      html = html.replace(/\s+(تر|ترین|ام|ات|اش|مان|تان|شان)\b/g, '‌$1');
      html = html.replace(/(\d+)\s+درصد/g, '$1٪');
      el.innerHTML = html;
    });
    alert('اصلاح نگارش و درج نیم‌فاصله‌های هوشمند اعمال شد.');
    onEditorInput();
  };

  // English to Persian Digits
  const convertToPersianDigits = () => {
    const persianDigits = ['۰', '۱', '۲', '۳', '۴', '۵', '۶', '۷', '۸', '۹'];
    pages.forEach(p => {
      const el = pageRefs.current[p.id];
      if (!el) return;
      let html = el.innerHTML;
      html = html.replace(/[0-9]/g, w => persianDigits[+w]);
      el.innerHTML = html;
    });
    alert('اعداد انگلیسی به اعداد اصیل فارسی تبدیل شدند.');
    onEditorInput();
  };

  // Text to Speech (Read Aloud)
  const readAloud = () => {
    if (!('speechSynthesis' in window)) {
      alert('مرورگر شما از پخش صوتی پشتیبانی نمی‌کند.');
      return;
    }
    window.speechSynthesis.cancel();
    const sel = window.getSelection();
    let textToRead = sel.toString();
    if (!textToRead) {
      const activeEl = pageRefs.current[pages[activePageIndex]?.id];
      textToRead = activeEl ? activeEl.innerText : '';
    }
    if (!textToRead.trim()) return;

    const utterance = new SpeechSynthesisUtterance(textToRead);
    utterance.lang = 'fa-IR';
    utterance.rate = 0.95;
    window.speechSynthesis.speak(utterance);
  };

  // Generate Table of Contents
  const generateTableOfContents = () => {
    if (headings.length === 0) {
      alert('برای ایجاد فهرست مطالب، ابتدا چند عنوان (عنوان ۱ یا ۲ یا ۳) در متن خود بنویسید.');
      return;
    }

    let tocHtml = `
      <div class="word-toc-box">
        <div class="word-toc-title">فهرست مطالب سند</div>
    `;

    headings.forEach((h) => {
      const indent = h.tag === 'h1' ? '0' : h.tag === 'h2' ? '18px' : '32px';
      tocHtml += `
        <div class="word-toc-row" style="padding-right: ${indent}">
          <span>${h.text}</span>
          <div class="word-toc-dots"></div>
          <span style="color: #64748b;">صفحه ۱</span>
        </div>
      `;
    });

    tocHtml += `</div><p><br></p>`;
    document.execCommand('insertHTML', false, tocHtml);
    onEditorInput();
  };

  // Insert Checklist
  const insertChecklist = () => {
    const html = `
      <ul class="word-checklist">
        <li><input type="checkbox"> <span>مورد اول فعالیت / وظیفه</span></li>
        <li><input type="checkbox"> <span>مورد دوم فعالیت / وظیفه</span></li>
      </ul>
      <p><br></p>
    `;
    document.execCommand('insertHTML', false, html);
    onEditorInput();
  };

  // Insert Table
  const insertTable = () => {
    let html = `<table class="${tableStyled ? 'word-styled-table' : ''}"><tbody>`;
    for (let r = 0; r < tableRows; r++) {
      html += '<tr>';
      for (let c = 0; c < tableCols; c++) {
        if (r === 0 && tableStyled) {
          html += `<th>سرستون ${c + 1}</th>`;
        } else {
          html += `<td>سلول ${r + 1}-${c + 1}</td>`;
        }
      }
      html += '</tr>';
    }
    html += '</tbody></table><p><br></p>';
    document.execCommand('insertHTML', false, html);
    setActiveModal(null);
    onEditorInput();
  };

  const addTableRow = () => {
    const sel = window.getSelection();
    let td = sel.anchorNode;
    while (td && td.tagName !== 'TD' && td.tagName !== 'TH') td = td.parentNode;
    if (!td) {
      alert('لطفاً ابتدا روی یکی از خانه‌های جدول کلیک کنید.');
      return;
    }
    const tr = td.parentNode;
    const newTr = document.createElement('tr');
    for (let i = 0; i < tr.children.length; i++) {
      const newTd = document.createElement('td');
      newTd.innerHTML = '&nbsp;';
      newTr.appendChild(newTd);
    }
    tr.after(newTr);
    onEditorInput();
  };

  const deleteTableRow = () => {
    const sel = window.getSelection();
    let td = sel.anchorNode;
    while (td && td.tagName !== 'TD' && td.tagName !== 'TH') td = td.parentNode;
    if (!td) return;
    td.parentNode.remove();
    onEditorInput();
  };

  const addTableColumn = () => {
    const sel = window.getSelection();
    let td = sel.anchorNode;
    while (td && td.tagName !== 'TD' && td.tagName !== 'TH') td = td.parentNode;
    if (!td) return;
    const colIndex = Array.from(td.parentNode.children).indexOf(td);
    const table = td.closest('table');
    table.querySelectorAll('tr').forEach(tr => {
      const isHeader = tr.children[0]?.tagName === 'TH';
      const newCell = document.createElement(isHeader ? 'th' : 'td');
      newCell.innerHTML = isHeader ? 'ستون جدید' : '&nbsp;';
      if (tr.children[colIndex]) tr.children[colIndex].after(newCell);
      else tr.appendChild(newCell);
    });
    onEditorInput();
  };

  const deleteTableColumn = () => {
    const sel = window.getSelection();
    let td = sel.anchorNode;
    while (td && td.tagName !== 'TD' && td.tagName !== 'TH') td = td.parentNode;
    if (!td) return;
    const colIndex = Array.from(td.parentNode.children).indexOf(td);
    const table = td.closest('table');
    table.querySelectorAll('tr').forEach(tr => {
      if (tr.children[colIndex]) tr.children[colIndex].remove();
    });
    onEditorInput();
  };

  // Image Upload
  const handleImageUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const imgHtml = `
        <div style="text-align: center; margin: 16px 0;">
          <img src="${event.target.result}" style="max-width: 80%; border-radius: 4px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);" alt="تصویر">
          <p style="font-size: 11px; color: #64748b; margin-top: 4px;">تصویر پیوست شده</p>
        </div>
        <p><br></p>
      `;
      document.execCommand('insertHTML', false, imgHtml);
      onEditorInput();
    };
    reader.readAsDataURL(file);
  };

  // Callout Box
  const insertCallout = (type) => {
    const icons = { info: '💡', warning: '⚠️', danger: '⛔', success: '✅' };
    const titles = { info: 'یادداشت مهم:', warning: 'هشدار و توجه:', danger: 'خط قرمز و نکته حساس:', success: 'تأییدیه نهایی:' };
    const html = `
      <div class="word-callout ${type}">
        <span style="font-size: 20px;">${icons[type]}</span>
        <div><strong>${titles[type]}</strong> متن نکته یا توضیح تکمیلی خود را در اینجا بنویسید...</div>
      </div>
      <p><br></p>
    `;
    document.execCommand('insertHTML', false, html);
    onEditorInput();
  };

  // Signature Box
  const insertSignatureBox = () => {
    const html = `
      <div class="signature-box-container">
        <div class="signature-stamp-card">
          <div class="signature-stamp-title">محل مهر رسمی شرکت</div>
          <div class="signature-stamp-line">تاریخ: ${new Intl.DateTimeFormat('fa-IR').format(new Date())}</div>
        </div>
        <div class="signature-stamp-card">
          <div class="signature-stamp-title">امضای مقام مجاز</div>
          <div class="signature-stamp-line">نام و سمت</div>
        </div>
      </div>
      <p><br></p>
    `;
    document.execCommand('insertHTML', false, html);
    onEditorInput();
  };

  // Symbols Picker
  const insertSymbol = (sym) => {
    document.execCommand('insertText', false, sym);
    setActiveModal(null);
    onEditorInput();
  };

  // Margins
  const applyMarginsPreset = (preset) => {
    const marginPresets = {
      normal: { top: 2.54, bottom: 2.54, left: 2.54, right: 2.54, type: 'normal' },
      narrow: { top: 1.27, bottom: 1.27, left: 1.27, right: 1.27, type: 'narrow' },
      wide: { top: 3.18, bottom: 3.18, left: 3.18, right: 3.18, type: 'wide' }
    };
    if (marginPresets[preset]) {
      setMargins(marginPresets[preset]);
      debounceSave();
    }
  };

  const applyCustomMargins = () => {
    setMargins({ top: customTop, bottom: customBottom, left: customLeft, right: customRight, type: 'custom' });
    setActiveModal(null);
    debounceSave();
  };

  // Watermark
  const applyWatermark = () => {
    setWatermark({ enabled: watermarkEnabled, text: watermarkText });
    setActiveModal(null);
    debounceSave();
  };

  // Page Border
  const applyPageBorder = () => {
    setPageBorder({ style: borderStyle, width: borderWidth, color: borderColor });
    setActiveModal(null);
    debounceSave();
  };

  // Snapshots
  const createSnapshot = () => {
    const name = snapshotName.trim() || `نسخه ${snapshots.length + 1}`;
    const newSnapshot = {
      id: Date.now(),
      name,
      date: new Intl.DateTimeFormat('fa-IR', { dateStyle: 'short', timeStyle: 'short' }).format(new Date()),
      title: docTitle,
      pages: pages.map(p => ({ id: p.id, content: pageRefs.current[p.id]?.innerHTML || p.content }))
    };
    const updated = [newSnapshot, ...snapshots];
    setSnapshots(updated);
    localStorage.setItem('word_react_snapshots', JSON.stringify(updated));
    setSnapshotName('');
    alert(`نسخه "${name}" با موفقیت ثبت شد.`);
  };

  const restoreSnapshot = (snap) => {
    if (confirm(`آیا مایلید سند را به "${snap.name}" بازگردانید؟`)) {
      setPages(snap.pages);
      setDocTitle(snap.title);
      setActiveModal(null);
      setTimeout(() => {
        snap.pages.forEach(p => {
          if (pageRefs.current[p.id]) pageRefs.current[p.id].innerHTML = p.content;
        });
        onEditorInput();
      }, 50);
    }
  };

  // Find & Replace
  const handleFindNext = () => {
    if (searchTerm) window.find(searchTerm);
  };

  const handleReplaceOne = () => {
    if (!searchTerm) return;
    const found = window.find(searchTerm, false, false, true, false, false, false);
    if (found) {
      document.execCommand('insertHTML', false, replaceTerm);
      onEditorInput();
    } else {
      alert(`موردی برای "${searchTerm}" پیدا نشد.`);
    }
  };

  const handleReplaceAll = () => {
    if (!searchTerm) return;
    const regex = new RegExp(searchTerm, 'g');
    pages.forEach(p => {
      const el = pageRefs.current[p.id];
      if (el) el.innerHTML = el.innerHTML.replace(regex, replaceTerm);
    });
    alert(`تمام موارد "${searchTerm}" با موفقیت جایگزین گردیدند.`);
    onEditorInput();
  };

  // Templates
  const applyTemplate = (key) => {
    const t = DocumentTemplates[key];
    if (!t) return;
    setDocTitle(t.title);
    setPages([{ id: 1, content: t.content }]);
    if (headerContentRef.current && t.header) headerContentRef.current.innerHTML = t.header;
    if (footerContentRef.current && t.footer) footerContentRef.current.innerHTML = t.footer;
    if (key === 'academicPaper') setColumns(2);
    else setColumns(1);
    setActiveTab('home');
    setTimeout(() => {
      if (pageRefs.current[1]) pageRefs.current[1].innerHTML = t.content;
      onEditorInput();
    }, 50);
  };

  // High-Fidelity Vector PDF Export Engine (100% Connected Persian Ligatures & Selectable Text)
  const exportPDF = () => {
    // Dynamically inject @page orientation style
    let printStyle = document.getElementById('dynamic-print-style');
    if (!printStyle) {
      printStyle = document.createElement('style');
      printStyle.id = 'dynamic-print-style';
      document.head.appendChild(printStyle);
    }
    printStyle.innerHTML = `@page { size: A4 ${orientation}; margin: 0mm; }`;

    const originalDocTitle = document.title;
    document.title = `${docTitle}`;

    const toast = document.createElement('div');
    toast.style.cssText = `
      position: fixed;
      bottom: 40px;
      left: 50%;
      transform: translateX(-50%);
      background: #103f91;
      color: #ffffff;
      padding: 14px 28px;
      border-radius: 8px;
      z-index: 999999;
      font-weight: bold;
      font-size: 14px;
      box-shadow: 0 8px 30px rgba(0,0,0,0.3);
      direction: rtl;
      text-align: center;
      font-family: 'Vazirmatn', system-ui, sans-serif;
    `;
    toast.innerHTML = `
      <div style="font-size: 15px; margin-bottom: 4px;">🚀 در حال آماده‌سازی پنجره چاپ و خروجی PDF وکتوری...</div>
      <div style="font-size: 12px; opacity: 0.9;">در پنجره باز شده، چاپگر را روی گزینه <b>«Save as PDF»</b> (ذخیره به عنوان PDF) قرار دهید تا متن با خطوط پیوسته فارسی ذخیره شود.</div>
    `;
    document.body.appendChild(toast);

    setTimeout(() => {
      toast.remove();
      window.print();
      setTimeout(() => {
        document.title = originalDocTitle;
      }, 1000);
    }, 500);
  };

  // Export to Native Microsoft Word (.docx) with Full Multi-Page & RTL Support
  const exportDOCX = () => {
    if (window.DocxManager) {
      const currentPages = pages.map(p => ({
        id: p.id,
        html: pageRefs.current[p.id]?.innerHTML || p.content,
        node: pageRefs.current[p.id]
      }));
      window.DocxManager.exportToDocx({
        filename: `${docTitle}.docx`,
        title: docTitle,
        pages: currentPages,
        orientation: orientation,
        margins: margins
      });
    } else {
      alert('موتور پردازش فایل ورد در حال بارگذاری است.');
    }
  };

  // Export Full Document Project (.word JSON) for 100% Exact Reload
  const exportProject = () => {
    const currentPages = pages.map(p => ({
      id: p.id,
      content: pageRefs.current[p.id]?.innerHTML || p.content
    }));
    const projectData = {
      title: docTitle,
      pages: currentPages,
      header: headerContentRef.current ? headerContentRef.current.innerHTML : '',
      footer: footerContentRef.current ? footerContentRef.current.innerHTML : '',
      margins,
      orientation,
      columns,
      paperTheme,
      watermark,
      pageBorder,
      wordGoal,
      savedAt: new Date().toISOString(),
      version: '1.0'
    };
    if (window.DocxManager) {
      window.DocxManager.exportProject(projectData, `${docTitle}.word`);
    }
  };

  // Export HTML Format
  const exportHTML = () => {
    let fullHtml = `<!DOCTYPE html><html lang="fa" dir="rtl"><head><meta charset="utf-8"><title>${docTitle}</title></head><body>`;
    pages.forEach(p => {
      fullHtml += pageRefs.current[p.id]?.innerHTML || '';
    });
    fullHtml += `</body></html>`;
    const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${docTitle}.html`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Universal File Opener (.docx, .word, .json)
  const handleOpenDocument = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'docx') {
      if (window.DocxManager) {
        window.DocxManager.importDocx(file);
      }
    } else if (ext === 'word' || ext === 'json') {
      if (window.DocxManager) {
        window.DocxManager.importProject(file);
      }
    } else {
      alert('فرمت فایل پشتیبانی نمی‌شود. لطفاً فایل با پسوند .docx یا .word انتخاب نمایید.');
    }
    e.target.value = '';
    setActiveTab('home');
  };

  // Scroll styles gallery left/right
  const scrollStyles = (dir) => {
    if (stylesGalleryRef.current) {
      stylesGalleryRef.current.scrollBy({ left: dir * 140, behavior: 'smooth' });
    }
  };

  // 10 Theme Colors Palette matrix
  const themeColorsMatrix = [
    ['#ffffff', '#000000', '#eeece1', '#1f497d', '#4f81bd', '#c0504d', '#9bbb59', '#8064a2', '#4bacc6', '#f79646'],
    ['#f2f2f2', '#7f7f7f', '#ddd9c3', '#c6d9f0', '#dce6f1', '#f2dcdb', '#ebf1dd', '#e5dfec', '#daeef3', '#fdeada'],
    ['#d8d8d8', '#595959', '#c4bd97', '#8db3e2', '#b8cce4', '#e5b9b7', '#d7e3bc', '#ccc1d9', '#b7dde8', '#fbd5b5'],
    ['#bfbfbf', '#3f3f3f', '#938953', '#548dd4', '#95b3d7', '#d99694', '#c3d69b', '#b2a2c7', '#92cddc', '#fac08f'],
    ['#a5a5a5', '#262626', '#494429', '#17365d', '#366092', '#953734', '#76923c', '#5f497a', '#31859b', '#e36c09'],
    ['#7f7f7f', '#0c0c0c', '#1d1b10', '#0f243e', '#244061', '#632423', '#4f6128', '#3f3151', '#205867', '#974806']
  ];

  const standardColors = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];

  const highlightColors = ['#ffff00', '#00ff00', '#00ffff', '#ff00ff', '#0000ff', '#ff0000', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#c0c0c0', '#000000'];

  const cmPx = 37.795;
  const pageStyle = {
    paddingTop: `${margins.top * cmPx * 0.4}px`,
    paddingBottom: `${margins.bottom * cmPx * 0.4}px`,
    paddingRight: `${margins.right * cmPx}px`,
    paddingLeft: `${margins.left * cmPx}px`,
    outline: pageBorder.style !== 'none' ? `${pageBorder.width} ${pageBorder.style} ${pageBorder.color}` : 'none',
    outlineOffset: '-12px'
  };

  const goalPercentage = Math.min(100, Math.round((stats.words / (wordGoal || 1)) * 100));

  return (
    <div className={`app-wrapper ${isDark ? 'dark-mode' : ''}`}>
      
      {/* 1. TOP HEADER (Microsoft Office 365 Blue) */}
      <header className="app-header">
        <div className="header-left">
          <button className="app-waffle" title="برنامه‌های مایکروسافت ۳۶۵">
            <i data-lucide="grid"></i>
          </button>
          <div className="word-branding">
            <div className="word-icon">W</div>
            <div className="doc-title-container">
              <input
                type="text"
                className="doc-title-input"
                value={docTitle}
                onChange={(e) => { setDocTitle(e.target.value); debounceSave(); }}
                title="تغییر نام سند"
              />
              <div className="save-status-badge">
                <i data-lucide="cloud-check" style={{ width: 14, height: 14 }}></i>
                <span>{saveStatus}</span>
              </div>
            </div>
          </div>
        </div>

        <div className="header-center">
          <div className="header-search">
            <input type="text" placeholder="جستجو در ابزارها، قالب‌ها و تنظیمات ورد..." />
            <i data-lucide="search" style={{ width: 16, height: 16 }}></i>
          </div>
        </div>

        <div className="header-right">
          <button className="header-btn" onClick={() => setIsDark(!isDark)} title="تغییر حالت شب و روز">
            <i data-lucide={isDark ? "sun" : "moon"} style={{ width: 16, height: 16 }}></i>
          </button>
          <button className="header-btn" onClick={() => saveToStorage(true)} title="ذخیره در مرورگر (Ctrl+S)">
            <i data-lucide="save" style={{ width: 16, height: 16 }}></i>
            <span>ذخیره</span>
          </button>
          <button className="header-btn primary" onClick={exportPDF} title="خروجی مستقیم PDF برداری (بدون بهم‌ریختگی حروف فارسی)">
            <i data-lucide="file-down" style={{ width: 16, height: 16 }}></i>
            <span>خروجی PDF</span>
          </button>
          <button className="header-btn primary" onClick={exportDOCX} title="خروجی رسمی فایل Word (.docx)">
            <i data-lucide="file-text" style={{ width: 16, height: 16 }}></i>
            <span>خروجی Word</span>
          </button>
          <button className="header-btn" onClick={exportProject} title="ذخیره کامل پروژه سند (.word)">
            <i data-lucide="hard-drive-download" style={{ width: 16, height: 16 }}></i>
            <span>پروژه Word</span>
          </button>
          <button className="header-btn" onClick={() => fileInputDocxRef.current?.click()} title="باز کردن و ویرایش فایل ورد یا پروژه (.docx, .word)">
            <i data-lucide="folder-open" style={{ width: 16, height: 16 }}></i>
            <span>باز کردن</span>
          </button>
          <div style={{ width: 32, height: 32, borderRadius: '50%', background: '#fff', color: '#103f91', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', fontSize: 13, cursor: 'pointer' }}>
            کاربر
          </div>
        </div>
      </header>

      {/* 2. RIBBON NAVIGATION */}
      <div className="ribbon-container">
        {/* Ribbon Tabs */}
        <div className="ribbon-tabs">
          <button className={`ribbon-tab file-tab ${activeTab === 'file' ? 'active' : ''}`} onClick={() => setActiveTab('file')}>فایل (File)</button>
          <button className={`ribbon-tab ${activeTab === 'home' ? 'active' : ''}`} onClick={() => setActiveTab('home')}>خانه (Home)</button>
          <button className={`ribbon-tab ${activeTab === 'insert' ? 'active' : ''}`} onClick={() => setActiveTab('insert')}>درج (Insert)</button>
          <button className={`ribbon-tab ${activeTab === 'layout' ? 'active' : ''}`} onClick={() => setActiveTab('layout')}>طرح‌بندی صفحه (Layout)</button>
          <button className={`ribbon-tab ${activeTab === 'review' ? 'active' : ''}`} onClick={() => setActiveTab('review')}>بازبینی و نگارش (Review)</button>
          <button className={`ribbon-tab ${activeTab === 'view' ? 'active' : ''}`} onClick={() => setActiveTab('view')}>نمایش (View)</button>
        </div>

        {/* Ribbon Panels */}
        <div className="ribbon-panels">
          
          {/* ========================================================
              HOME PANEL (Exact Word Ribbon Matching the Provided Image)
              ======================================================== */}
          <div className={`ribbon-panel ${activeTab === 'home' ? 'active' : ''}`}>
            
            {/* 1. CLIPBOARD GROUP */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                {/* Large Paste Button with Dropdown Arrow */}
                <div style={{ position: 'relative' }}>
                  <button className="ribbon-btn large" onClick={() => execCmd('paste')} title="جای‌گذاری (Ctrl+V)">
                    <i data-lucide="clipboard" style={{ width: 22, height: 22, color: '#f59e0b' }}></i>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 2 }}>
                      Paste <span className="ribbon-btn-arrow">▾</span>
                    </span>
                  </button>
                </div>

                <div className="ribbon-col">
                  {/* Cut */}
                  <button className="ribbon-btn" onClick={() => execCmd('cut')} title="برش (Ctrl+X)">
                    <i data-lucide="scissors" style={{ width: 14, height: 14, marginLeft: 4 }}></i>
                    <span>Cut</span>
                  </button>
                  {/* Copy */}
                  <button className="ribbon-btn" onClick={() => execCmd('copy')} title="کپی (Ctrl+C)">
                    <i data-lucide="copy" style={{ width: 14, height: 14, marginLeft: 4 }}></i>
                    <span>Copy</span>
                  </button>
                  {/* Format Painter */}
                  <button
                    className={`ribbon-btn ${formatPainterActive ? 'active' : ''}`}
                    onClick={handleFormatPainter}
                    title="کپی فرمت متن (Format Painter)"
                  >
                    <i data-lucide="paint-roller" style={{ width: 14, height: 14, marginLeft: 4, color: '#d97706' }}></i>
                    <span>Format Painter</span>
                  </button>
                </div>
              </div>
              <div className="ribbon-group-bottom">
                <span className="ribbon-group-label">Clipboard</span>
                <button className="ribbon-dialog-launcher" title="تنظیمات کلیپ‌بورد">⤢</button>
              </div>
            </div>

            {/* 2. FONT GROUP */}
            <div className="ribbon-group">
              <div className="ribbon-group-content ribbon-col">
                {/* Row 1: Font Select, Size, Grow, Shrink, Case, Clear Formatting */}
                <div className="ribbon-row">
                  <select
                    className="ribbon-select font-select"
                    value={fontFamily}
                    onChange={(e) => changeFontFamily(e.target.value)}
                    title="قلم متن (Font)"
                  >
                    <option value="Calibri">Calibri (Body)</option>
                    <option value="Arial">Arial</option>
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Vazirmatn">وزیرمتن (Vazirmatn)</option>
                    <option value="B Nazanin">بی نازنین (B Nazanin)</option>
                    <option value="Tahoma">تاهوما (Tahoma)</option>
                    <option value="Shabnam">شبنم (Shabnam)</option>
                    <option value="Sahel">ساحل (Sahel)</option>
                    <option value="Samim">صمیم (Samim)</option>
                    <option value="Courier New">Courier New</option>
                  </select>

                  <select
                    className="ribbon-select size-select"
                    value={fontSize}
                    onChange={(e) => changeFontSize(e.target.value)}
                    title="اندازه قلم (Font Size)"
                  >
                    {[8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>

                  {/* Grow Font A^ */}
                  <button className="ribbon-btn" onClick={() => changeFontSize(Math.min(+fontSize + 2, 72))} title="بزرگ‌تر کردن فونت (Increase Font Size)">
                    <span style={{ fontWeight: 'bold', fontSize: 13 }}>A<sup>▲</sup></span>
                  </button>
                  
                  {/* Shrink Font Av */}
                  <button className="ribbon-btn" onClick={() => changeFontSize(Math.max(+fontSize - 2, 8))} title="کوچک‌تر کردن فونت (Decrease Font Size)">
                    <span style={{ fontWeight: 'bold', fontSize: 11 }}>A<sup>▼</sup></span>
                  </button>

                  {/* Change Case Aa ▾ */}
                  <div style={{ position: 'relative' }}>
                    <button
                      className="ribbon-btn has-dropdown"
                      onClick={() => setOpenDropdown(openDropdown === 'case' ? null : 'case')}
                      title="تغییر حروف بزرگ و کوچک (Change Case)"
                    >
                      <span style={{ fontWeight: 'bold' }}>Aa</span>
                      <span className="ribbon-btn-arrow">▾</span>
                    </button>
                    {openDropdown === 'case' && (
                      <div className="word-dropdown-menu">
                        <div className="word-dropdown-item" onClick={() => changeCase('sentence')}>حالت جمله (Sentence case.)</div>
                        <div className="word-dropdown-item" onClick={() => changeCase('lower')}>حروف کوچک (lowercase)</div>
                        <div className="word-dropdown-item" onClick={() => changeCase('upper')}>حروف بزرگ (UPPERCASE)</div>
                        <div className="word-dropdown-item" onClick={() => changeCase('capitalize')}>حرف اول بزرگ (Capitalize Each Word)</div>
                        <div className="word-dropdown-item" onClick={() => changeCase('toggle')}>وارونه (tOGGLE cASE)</div>
                      </div>
                    )}
                  </div>

                  {/* Clear All Formatting */}
                  <button className="ribbon-btn" onClick={() => execCmd('removeFormat')} title="پاک کردن تمام فرمت‌ها (Clear All Formatting)">
                    <span style={{ position: 'relative', fontWeight: 'bold', color: '#111827' }}>
                      A<span style={{ color: '#db2777', fontSize: 10, position: 'absolute', bottom: -2, right: -4 }}>⌫</span>
                    </span>
                  </button>
                </div>

                {/* Row 2: B, I, U ▾, abc, x₂, x², Text Effects A ▾, Highlight ▾, Font Color ▾ */}
                <div className="ribbon-row">
                  {/* Bold */}
                  <button className={`ribbon-btn ${activeStyles.bold ? 'active' : ''}`} onClick={() => execCmd('bold')} title="ضخیم (Bold - Ctrl+B)">
                    <strong style={{ fontFamily: 'Georgia, serif' }}>B</strong>
                  </button>

                  {/* Italic */}
                  <button className={`ribbon-btn ${activeStyles.italic ? 'active' : ''}`} onClick={() => execCmd('italic')} title="مورب (Italic - Ctrl+I)">
                    <em style={{ fontFamily: 'Georgia, serif' }}>I</em>
                  </button>

                  {/* Underline with Dropdown U ▾ */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className={`ribbon-btn ${activeStyles.underline ? 'active' : ''}`} onClick={() => execCmd('underline')} title="زیرخط (Underline - Ctrl+U)">
                      <u style={{ fontFamily: 'Georgia, serif' }}>U</u>
                    </button>
                    <button
                      className="split-btn-arrow"
                      onClick={() => setOpenDropdown(openDropdown === 'underline' ? null : 'underline')}
                      title="انتخاب سبک خط زیر"
                    >
                      ▾
                    </button>
                    {openDropdown === 'underline' && (
                      <div className="word-dropdown-menu">
                        <div className="word-dropdown-item" onClick={() => applyUnderlineStyle('solid')}>خط ساده (Single Underline)</div>
                        <div className="word-dropdown-item" onClick={() => applyUnderlineStyle('double')}>خط دوتایی (Double Underline)</div>
                        <div className="word-dropdown-item" onClick={() => applyUnderlineStyle('dotted')}>خط‌چین نقطه‌ای (Dotted)</div>
                        <div className="word-dropdown-item" onClick={() => applyUnderlineStyle('dashed')}>خط‌چین فاصله (Dashed)</div>
                        <div className="word-dropdown-item" onClick={() => applyUnderlineStyle('wavy')}>خط موج‌دار (Wavy)</div>
                      </div>
                    )}
                  </div>

                  {/* Strikethrough abc */}
                  <button className={`ribbon-btn ${activeStyles.strike ? 'active' : ''}`} onClick={() => execCmd('strikeThrough')} title="خط‌خورده (Strikethrough)">
                    <s>abc</s>
                  </button>

                  {/* Subscript x₂ */}
                  <button className="ribbon-btn" onClick={() => execCmd('subscript')} title="زیرنویس (Subscript)">
                    x<sub>2</sub>
                  </button>

                  {/* Superscript x² */}
                  <button className="ribbon-btn" onClick={() => execCmd('superscript')} title="بالانویس (Superscript)">
                    x<sup>2</sup>
                  </button>

                  {/* Text Effects & Typography Glowing A ▾ */}
                  <div style={{ position: 'relative' }}>
                    <button
                      className="ribbon-btn has-dropdown"
                      onClick={() => setOpenDropdown(openDropdown === 'textEffects' ? null : 'textEffects')}
                      title="افکت‌های متنی و تایپوگرافی (Text Effects and Typography)"
                    >
                      <span style={{ fontWeight: 'bold', color: '#2563eb', textShadow: '0 0 3px #93c5fd' }}>A</span>
                      <span className="ribbon-btn-arrow">▾</span>
                    </button>
                    {openDropdown === 'textEffects' && (
                      <div className="word-dropdown-menu">
                        <div className="word-dropdown-item" onClick={() => applyTextEffect('shadow')}>سایه متن (Shadow)</div>
                        <div className="word-dropdown-item" onClick={() => applyTextEffect('glow')}>درخشش آبی (Glow)</div>
                        <div className="word-dropdown-item" onClick={() => applyTextEffect('outline')}>خط بیرونی (Outline)</div>
                        <div className="word-dropdown-item" onClick={() => applyTextEffect('3d')}>افکت سه‌بعدی (3D Style)</div>
                      </div>
                    )}
                  </div>

                  {/* Text Highlight Color Pen ▾ */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className="ribbon-btn" onClick={() => execCmd('hiliteColor', highlightColor)} title="هایلایت متن (Text Highlight Color)">
                      <i data-lucide="highlighter" style={{ width: 14, height: 14, color: '#ca8a04' }}></i>
                      <div className="color-indicator" style={{ backgroundColor: highlightColor }}></div>
                    </button>
                    <button
                      className="split-btn-arrow"
                      onClick={() => setOpenDropdown(openDropdown === 'highlight' ? null : 'highlight')}
                      title="انتخاب رنگ هایلایت"
                    >
                      ▾
                    </button>
                    {openDropdown === 'highlight' && (
                      <div className="word-dropdown-menu" style={{ width: 190 }}>
                        <div style={{ padding: '6px 8px' }}>
                          <div className="palette-section-title">رنگ‌های هایلایت:</div>
                          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: 4 }}>
                            {highlightColors.map((c, i) => (
                              <div
                                key={i}
                                className="theme-color-cell"
                                style={{ backgroundColor: c }}
                                onClick={() => {
                                  setHighlightColor(c);
                                  execCmd('hiliteColor', c);
                                  setOpenDropdown(null);
                                }}
                              />
                            ))}
                          </div>
                          <div className="word-dropdown-divider"></div>
                          <div className="word-dropdown-item" onClick={() => { execCmd('hiliteColor', 'transparent'); setOpenDropdown(null); }}>
                            بدون رنگ (No Color)
                          </div>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Font Color A ▾ with Palette */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className="ribbon-btn" onClick={() => execCmd('foreColor', fontColor)} title="رنگ قلم (Font Color)">
                      <span style={{ fontWeight: 'bold', fontSize: 13 }}>A</span>
                      <div className="color-indicator" style={{ backgroundColor: fontColor }}></div>
                    </button>
                    <button
                      className="split-btn-arrow"
                      onClick={() => setOpenDropdown(openDropdown === 'fontColor' ? null : 'fontColor')}
                      title="انتخاب پالت رنگ قلم"
                    >
                      ▾
                    </button>
                    {openDropdown === 'fontColor' && (
                      <div className="word-dropdown-menu color-palette-grid">
                        <div className="palette-section-title">رنگ‌های تم (Theme Colors):</div>
                        {themeColorsMatrix.map((row, ri) => (
                          <div key={ri} className="theme-colors-row">
                            {row.map((c, ci) => (
                              <div
                                key={ci}
                                className="theme-color-cell"
                                style={{ backgroundColor: c }}
                                onClick={() => {
                                  setFontColor(c);
                                  execCmd('foreColor', c);
                                  setOpenDropdown(null);
                                }}
                              />
                            ))}
                          </div>
                        ))}
                        <div className="palette-section-title" style={{ marginTop: 8 }}>رنگ‌های استاندارد (Standard Colors):</div>
                        <div className="theme-colors-row">
                          {standardColors.map((c, ci) => (
                            <div
                              key={ci}
                              className="theme-color-cell"
                              style={{ backgroundColor: c }}
                              onClick={() => {
                                setFontColor(c);
                                execCmd('foreColor', c);
                                setOpenDropdown(null);
                              }}
                            />
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="ribbon-group-bottom">
                <span className="ribbon-group-label">Font</span>
                <button className="ribbon-dialog-launcher" onClick={() => setActiveModal('font')} title="پنجره تنظیمات قلم (Font Dialog)">⤢</button>
              </div>
            </div>

            {/* 3. PARAGRAPH GROUP */}
            <div className="ribbon-group">
              <div className="ribbon-group-content ribbon-col">
                {/* Row 1: Bullets ▾, Numbering ▾, Multilevel ▾, Outdent, Indent, RTL, LTR, Sort A-Z, Show/Hide ¶ */}
                <div className="ribbon-row">
                  {/* Bullets ▾ */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className="ribbon-btn" onClick={() => execCmd('insertUnorderedList')} title="لیست نشانه‌دار (Bullets)">
                      <i data-lucide="list" style={{ width: 14, height: 14 }}></i>
                    </button>
                    <button className="split-btn-arrow" onClick={() => setOpenDropdown(openDropdown === 'bullets' ? null : 'bullets')}>▾</button>
                    {openDropdown === 'bullets' && (
                      <div className="word-dropdown-menu">
                        <div className="word-dropdown-item" onClick={() => applyBulletStyle('disc')}>● دایره توپر (Solid Circle)</div>
                        <div className="word-dropdown-item" onClick={() => applyBulletStyle('circle')}>○ دایره توخالی (Circle)</div>
                        <div className="word-dropdown-item" onClick={() => applyBulletStyle('square')}>■ مربع (Square)</div>
                        <div className="word-dropdown-item" onClick={() => applyBulletStyle('none')}>بدون نشانه (None)</div>
                      </div>
                    )}
                  </div>

                  {/* Numbering ▾ */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className="ribbon-btn" onClick={() => execCmd('insertOrderedList')} title="لیست شماره‌دار (Numbering)">
                      <i data-lucide="list-ordered" style={{ width: 14, height: 14 }}></i>
                    </button>
                    <button className="split-btn-arrow" onClick={() => setOpenDropdown(openDropdown === 'numbering' ? null : 'numbering')}>▾</button>
                    {openDropdown === 'numbering' && (
                      <div className="word-dropdown-menu">
                        <div className="word-dropdown-item" onClick={() => applyNumberingStyle('decimal')}>1. 2. 3. عددی لاتین</div>
                        <div className="word-dropdown-item" onClick={() => applyNumberingStyle('lower-alpha')}>a. b. c. حروف کوچک</div>
                        <div className="word-dropdown-item" onClick={() => applyNumberingStyle('upper-roman')}>I. II. III. یونانی</div>
                        <div className="word-dropdown-item" onClick={() => applyNumberingStyle('decimal')}>۱. ۲. ۳. عددی فارسی</div>
                      </div>
                    )}
                  </div>

                  {/* Multilevel List ▾ */}
                  <button className="ribbon-btn" onClick={generateTableOfContents} title="لیست چند سطحی و ساختار عناوین (Multilevel List)">
                    <span style={{ fontSize: 11, fontWeight: 'bold' }}>1<sup>a</sup><sub>i</sub></span>
                  </button>

                  {/* Decrease Indent */}
                  <button className="ribbon-btn" onClick={() => execCmd('outdent')} title="کاهش تورفتگی (Decrease Indent)">
                    <i data-lucide="indent-decrease"></i>
                  </button>

                  {/* Increase Indent */}
                  <button className="ribbon-btn" onClick={() => execCmd('indent')} title="افزایش تورفتگی (Increase Indent)">
                    <i data-lucide="indent-increase"></i>
                  </button>

                  {/* RTL Text Direction */}
                  <button className="ribbon-btn" onClick={() => {
                    const activeEl = pageRefs.current[pages[activePageIndex]?.id];
                    if (activeEl) { activeEl.dir = 'rtl'; onEditorInput(); }
                  }} title="جهت متن: راست به چپ (Right-to-Left Text Direction)">
                    <span style={{ fontWeight: 'bold', fontSize: 11 }}>¶▶</span>
                  </button>

                  {/* LTR Text Direction */}
                  <button className="ribbon-btn" onClick={() => {
                    const activeEl = pageRefs.current[pages[activePageIndex]?.id];
                    if (activeEl) { activeEl.dir = 'ltr'; onEditorInput(); }
                  }} title="جهت متن: چپ به راست (Left-to-Right Text Direction)">
                    <span style={{ fontWeight: 'bold', fontSize: 11 }}>◀¶</span>
                  </button>

                  {/* Sort Paragraphs A-Z */}
                  <button className="ribbon-btn" onClick={sortParagraphs} title="مرتب‌سازی الفبایی (Sort)">
                    <span style={{ fontSize: 11, fontWeight: 'bold' }}>A↓Z</span>
                  </button>

                  {/* Show/Hide Formatting Marks ¶ */}
                  <button className={`ribbon-btn ${showPilcrow ? 'active' : ''}`} onClick={() => setShowPilcrow(!showPilcrow)} title="نمایش نشانه‌های فرمت‌بندی (Show/Hide ¶)">
                    <span style={{ fontWeight: 'bold', fontSize: 13 }}>¶</span>
                  </button>
                </div>

                {/* Row 2: Align Left, Center, Align Right, Justify ▾, Line Spacing ▾, Shading ▾, Borders ▾ */}
                <div className="ribbon-row">
                  {/* Align Left */}
                  <button className="ribbon-btn" onClick={() => execCmd('justifyLeft')} title="چپ‌چین (Align Left - Ctrl+L)">
                    <i data-lucide="align-left"></i>
                  </button>

                  {/* Center */}
                  <button className="ribbon-btn" onClick={() => execCmd('justifyCenter')} title="وسط‌چین (Center - Ctrl+E)">
                    <i data-lucide="align-center"></i>
                  </button>

                  {/* Align Right */}
                  <button className="ribbon-btn" onClick={() => execCmd('justifyRight')} title="راست‌چین (Align Right - Ctrl+R)">
                    <i data-lucide="align-right"></i>
                  </button>

                  {/* Justify ▾ */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className="ribbon-btn" onClick={() => execCmd('justifyFull')} title="تراز کامل (Justify)">
                      <i data-lucide="align-justify"></i>
                    </button>
                    <button className="split-btn-arrow" onClick={() => setOpenDropdown(openDropdown === 'justify' ? null : 'justify')}>▾</button>
                    {openDropdown === 'justify' && (
                      <div className="word-dropdown-menu">
                        <div className="word-dropdown-item" onClick={() => { execCmd('justifyFull'); setOpenDropdown(null); }}>تراز کامل (Justify)</div>
                        <div className="word-dropdown-item" onClick={() => { execCmd('justifyFull'); setOpenDropdown(null); }}>تراز با کشش کم (Justify Low)</div>
                        <div className="word-dropdown-item" onClick={() => { execCmd('justifyFull'); setOpenDropdown(null); }}>تراز با کشش زیاد (Justify High)</div>
                      </div>
                    )}
                  </div>

                  {/* Line and Paragraph Spacing ▾ */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className="ribbon-btn" title="فاصله بین خطوط (Line and Paragraph Spacing)">
                      <span style={{ fontSize: 11, fontWeight: 'bold' }}>↕≡</span>
                    </button>
                    <button className="split-btn-arrow" onClick={() => setOpenDropdown(openDropdown === 'spacing' ? null : 'spacing')}>▾</button>
                    {openDropdown === 'spacing' && (
                      <div className="word-dropdown-menu">
                        {['1.0', '1.15', '1.5', '2.0', '2.5', '3.0'].map(val => (
                          <div
                            key={val}
                            className="word-dropdown-item"
                            onClick={() => {
                              const activeEl = pageRefs.current[pages[activePageIndex]?.id];
                              if (activeEl) { activeEl.style.lineHeight = val; onEditorInput(); }
                              setOpenDropdown(null);
                            }}
                          >
                            فاصله خطوط {val}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                  {/* Shading (Paint Bucket) ▾ */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className="ribbon-btn" onClick={() => applyShadingColor(shadingColor)} title="رنگ پس‌زمینه پاراگراف (Shading)">
                      <i data-lucide="paint-bucket" style={{ width: 14, height: 14 }}></i>
                      <div className="color-indicator" style={{ backgroundColor: shadingColor }}></div>
                    </button>
                    <button className="split-btn-arrow" onClick={() => setOpenDropdown(openDropdown === 'shading' ? null : 'shading')}>▾</button>
                    {openDropdown === 'shading' && (
                      <div className="word-dropdown-menu color-palette-grid">
                        <div className="palette-section-title">رنگ پس‌زمینه (Shading):</div>
                        <div className="theme-colors-row">
                          {standardColors.map((c, i) => (
                            <div key={i} className="theme-color-cell" style={{ backgroundColor: c }} onClick={() => applyShadingColor(c)} />
                          ))}
                        </div>
                        <div className="word-dropdown-divider"></div>
                        <div className="word-dropdown-item" onClick={() => applyShadingColor('transparent')}>بدون رنگ (No Color)</div>
                      </div>
                    )}
                  </div>

                  {/* Borders ▾ */}
                  <div className="split-btn" style={{ position: 'relative' }}>
                    <button className="ribbon-btn" onClick={() => applyParagraphBorder('all')} title="کادر پاراگراف و جدول (Borders)">
                      <i data-lucide="grid-2x2" style={{ width: 14, height: 14 }}></i>
                    </button>
                    <button className="split-btn-arrow" onClick={() => setOpenDropdown(openDropdown === 'borders' ? null : 'borders')}>▾</button>
                    {openDropdown === 'borders' && (
                      <div className="word-dropdown-menu">
                        <div className="word-dropdown-item" onClick={() => applyParagraphBorder('bottom')}>خط پایین (Bottom Border)</div>
                        <div className="word-dropdown-item" onClick={() => applyParagraphBorder('top')}>خط بالا (Top Border)</div>
                        <div className="word-dropdown-item" onClick={() => applyParagraphBorder('left')}>خط چپ (Left Border)</div>
                        <div className="word-dropdown-item" onClick={() => applyParagraphBorder('right')}>خط راست (Right Border)</div>
                        <div className="word-dropdown-divider"></div>
                        <div className="word-dropdown-item" onClick={() => applyParagraphBorder('none')}>بدون کادر (No Border)</div>
                        <div className="word-dropdown-item" onClick={() => applyParagraphBorder('all')}>کادر کامل (All Borders)</div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="ribbon-group-bottom">
                <span className="ribbon-group-label">Paragraph</span>
                <button className="ribbon-dialog-launcher" onClick={() => setActiveModal('paragraph')} title="پنجره تنظیمات پاراگراف">⤢</button>
              </div>
            </div>

            {/* 4. STYLES GROUP (Exact Word Visual Gallery) */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <div className="styles-gallery-container">
                  <div className="styles-gallery" ref={stylesGalleryRef}>
                    <div className="style-card active" onClick={() => applyFormatBlock('p')}>
                      <div className="style-preview" style={{ fontSize: 13 }}>AaBbCcDc</div>
                      <div className="style-name">Normal</div>
                    </div>
                    <div className="style-card" onClick={() => applyFormatBlock('p')}>
                      <div className="style-preview" style={{ fontSize: 13 }}>AaBbCcDc</div>
                      <div className="style-name">No Spacing</div>
                    </div>
                    <div className="style-card" onClick={() => applyFormatBlock('h1')}>
                      <div className="style-preview" style={{ fontSize: 15, color: '#103f91' }}>AaBbCc</div>
                      <div className="style-name">Heading 1</div>
                    </div>
                    <div className="style-card" onClick={() => applyFormatBlock('h2')}>
                      <div className="style-preview" style={{ fontSize: 14, color: '#185abd' }}>AaBbCc</div>
                      <div className="style-name">Heading 2</div>
                    </div>
                    <div className="style-card" onClick={() => applyFormatBlock('h1')}>
                      <div className="style-preview" style={{ fontSize: 16, fontWeight: 'bold' }}>Title</div>
                      <div className="style-name">Title</div>
                    </div>
                    <div className="style-card" onClick={() => applyFormatBlock('h3')}>
                      <div className="style-preview" style={{ fontSize: 12, color: '#64748b' }}>Subtitle</div>
                      <div className="style-name">Subtitle</div>
                    </div>
                    <div className="style-card" onClick={() => applyFormatBlock('p')}>
                      <div className="style-preview" style={{ fontStyle: 'italic', color: '#185abd' }}>AaBbCc</div>
                      <div className="style-name">Subtle Em...</div>
                    </div>
                    <div className="style-card" onClick={() => applyFormatBlock('p')}>
                      <div className="style-preview" style={{ fontStyle: 'italic', fontWeight: 'bold' }}>AaBbCc</div>
                      <div className="style-name">Emphasis</div>
                    </div>
                    <div className="style-card" onClick={() => applyFormatBlock('blockquote')}>
                      <div className="style-preview" style={{ fontStyle: 'italic', color: '#103f91' }}>«Quote»</div>
                      <div className="style-name">Quote</div>
                    </div>
                  </div>

                  {/* Gallery Scroll Controls */}
                  <div className="styles-scroll-controls">
                    <button className="styles-scroll-btn" onClick={() => scrollStyles(-1)}>▲</button>
                    <button className="styles-scroll-btn" onClick={() => scrollStyles(1)}>▼</button>
                    <button className="styles-scroll-btn" onClick={() => scrollStyles(1)}>▾</button>
                  </div>
                </div>
              </div>
              <div className="ribbon-group-bottom">
                <span className="ribbon-group-label">Styles</span>
                <button className="ribbon-dialog-launcher" title="مدیریت استایل‌ها">⤢</button>
              </div>
            </div>

            {/* 5. EDITING GROUP */}
            <div className="ribbon-group">
              <div className="ribbon-group-content ribbon-col">
                {/* Find ▾ */}
                <button className="ribbon-btn" onClick={() => setShowFindReplace(true)} title="یافتن در سند (Find - Ctrl+F)">
                  <i data-lucide="search" style={{ width: 14, height: 14, marginLeft: 4, color: '#103f91' }}></i>
                  <span>Find ▾</span>
                </button>
                {/* Replace */}
                <button className="ribbon-btn" onClick={() => setShowFindReplace(true)} title="جایگزینی عبارات (Replace - Ctrl+H)">
                  <i data-lucide="replace" style={{ width: 14, height: 14, marginLeft: 4, color: '#9333ea' }}></i>
                  <span>Replace</span>
                </button>
                {/* Select ▾ */}
                <button className="ribbon-btn" onClick={() => execCmd('selectAll')} title="انتخاب همه متن (Select All)">
                  <i data-lucide="mouse-pointer-2" style={{ width: 14, height: 14, marginLeft: 4 }}></i>
                  <span>Select ▾</span>
                </button>
              </div>
              <div className="ribbon-group-bottom">
                <span className="ribbon-group-label">Editing</span>
              </div>
            </div>

          </div>

          {/* INSERT PANEL */}
          <div className={`ribbon-panel ${activeTab === 'insert' ? 'active' : ''}`}>
            {/* Table */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <button className="ribbon-btn large" onClick={() => setActiveModal('table')} title="درج جدول جدید">
                  <i data-lucide="table"></i>
                  <span>درج جدول</span>
                </button>
                <div className="ribbon-col">
                  <button className="ribbon-btn" onClick={addTableRow} title="افزودن سطر"><i data-lucide="rows-4" style={{ width: 14, height: 14, marginLeft: 4 }}></i><span>+ سطر</span></button>
                  <button className="ribbon-btn" onClick={deleteTableRow} title="حذف سطر"><i data-lucide="trash-2" style={{ width: 14, height: 14, marginLeft: 4, color: '#dc2626' }}></i><span>- سطر</span></button>
                </div>
                <div className="ribbon-col">
                  <button className="ribbon-btn" onClick={addTableColumn} title="افزودن ستون"><i data-lucide="columns-4" style={{ width: 14, height: 14, marginLeft: 4 }}></i><span>+ ستون</span></button>
                  <button className="ribbon-btn" onClick={deleteTableColumn} title="حذف ستون"><i data-lucide="trash-2" style={{ width: 14, height: 14, marginLeft: 4, color: '#dc2626' }}></i><span>- ستون</span></button>
                </div>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">جداول</span></div>
            </div>

            {/* Pages & TOC */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <button className="ribbon-btn large" onClick={addNewPage} title="افزودن برگه فیزیکی جدید">
                  <i data-lucide="file-plus-2"></i>
                  <span>صفحه جدید</span>
                </button>
                <button className="ribbon-btn large" onClick={generateTableOfContents} title="ایجاد خودکار فهرست مطالب از تیترها">
                  <i data-lucide="list-ordered"></i>
                  <span>فهرست مطالب</span>
                </button>
                <button className="ribbon-btn large" onClick={insertChecklist} title="چک‌لیست تعاملی">
                  <i data-lucide="check-square"></i>
                  <span>چک‌لیست</span>
                </button>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">صفحات و فهرست</span></div>
            </div>

            {/* Media & Callouts */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <input type="file" ref={imgInputRef} accept="image/*" style={{ display: 'none' }} onChange={handleImageUpload} />
                <button className="ribbon-btn large" onClick={() => imgInputRef.current?.click()} title="درج تصویر">
                  <i data-lucide="image"></i>
                  <span>تصویر</span>
                </button>
                <button className="ribbon-btn large" onClick={insertSignatureBox} title="کادر رسمی مهر و امضا">
                  <i data-lucide="stamp"></i>
                  <span>مهر و امضا</span>
                </button>
                <button className="ribbon-btn large" onClick={() => insertCallout('info')} title="کادر پیام">
                  <i data-lucide="sticky-note"></i>
                  <span>کادر پیام</span>
                </button>
                <button className="ribbon-btn large" onClick={() => insertCallout('warning')} title="کادر هشدار">
                  <i data-lucide="alert-triangle" style={{ color: '#d97706' }}></i>
                  <span>کادر هشدار</span>
                </button>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">تصاویر و جعبه‌ها</span></div>
            </div>

            {/* Symbols */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <button className="ribbon-btn large" onClick={() => setActiveModal('symbols')} title="نمادهای ریاضی و ارزی">
                  <i data-lucide="pi"></i>
                  <span>نمادها</span>
                </button>
                <button className="ribbon-btn large" onClick={() => execCmd('insertHorizontalRule')} title="خط جداکننده">
                  <i data-lucide="minus"></i>
                  <span>خط افقی</span>
                </button>
                <button className="ribbon-btn large" onClick={() => {
                  const now = new Intl.DateTimeFormat('fa-IR', { dateStyle: 'full' }).format(new Date());
                  execCmd('insertHTML', `<span>${now}</span>`);
                }} title="درج تاریخ امروز">
                  <i data-lucide="calendar"></i>
                  <span>تاریخ امروز</span>
                </button>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">نمادها و تاریخ</span></div>
            </div>
          </div>

          {/* LAYOUT PANEL */}
          <div className={`ribbon-panel ${activeTab === 'layout' ? 'active' : ''}`}>
            {/* Margins */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <button className={`ribbon-btn large ${margins.type === 'normal' ? 'active' : ''}`} onClick={() => applyMarginsPreset('normal')} title="حاشیه ۲.۵۴cm">
                  <i data-lucide="square"></i>
                  <span>عادی (۲.۵)</span>
                </button>
                <button className={`ribbon-btn large ${margins.type === 'narrow' ? 'active' : ''}`} onClick={() => applyMarginsPreset('narrow')} title="حاشیه ۱.۲۷cm">
                  <i data-lucide="minimize"></i>
                  <span>باریک (۱.۲)</span>
                </button>
                <button className={`ribbon-btn large ${margins.type === 'wide' ? 'active' : ''}`} onClick={() => applyMarginsPreset('wide')} title="حاشیه ۳.۱۸cm">
                  <i data-lucide="maximize"></i>
                  <span>عریض (۳.۱)</span>
                </button>
                <button className="ribbon-btn large" onClick={() => setActiveModal('margins')} title="حاشیه‌های سفارشی">
                  <i data-lucide="sliders"></i>
                  <span>سفارشی...</span>
                </button>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">حاشیه‌ها</span></div>
            </div>

            {/* Orientation & Columns */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <button className={`ribbon-btn large ${orientation === 'portrait' ? 'active' : ''}`} onClick={() => { setOrientation('portrait'); debounceSave(); }} title="کاغذ عمودی">
                  <i data-lucide="file"></i>
                  <span>عمودی</span>
                </button>
                <button className={`ribbon-btn large ${orientation === 'landscape' ? 'active' : ''}`} onClick={() => { setOrientation('landscape'); debounceSave(); }} title="کاغذ افقی">
                  <i data-lucide="file-text" style={{ transform: 'rotate(90deg)' }}></i>
                  <span>افقی</span>
                </button>
                
                <div style={{ width: 1, height: 32, background: '#cbd5e1', margin: '0 6px' }}></div>
                
                <button className={`ribbon-btn large ${columns === 1 ? 'active' : ''}`} onClick={() => { setColumns(1); debounceSave(); }} title="تک ستونه">
                  <i data-lucide="align-justify"></i>
                  <span>۱ ستون</span>
                </button>
                <button className={`ribbon-btn large ${columns === 2 ? 'active' : ''}`} onClick={() => { setColumns(2); debounceSave(); }} title="دو ستونه (مقالات)">
                  <i data-lucide="columns-2"></i>
                  <span>۲ ستون</span>
                </button>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">جهت و ستون‌ها</span></div>
            </div>

            {/* Paper Color / Themes */}
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <select
                  className="ribbon-select"
                  value={paperTheme}
                  onChange={(e) => { setPaperTheme(e.target.value); debounceSave(); }}
                  title="رنگ کاغذ برگه"
                  style={{ width: 110, height: 36 }}
                >
                  <option value="white">کاغذ سفید (استاندارد)</option>
                  <option value="sepia">کاغذ سپیا (کرم محافظ چشم)</option>
                  <option value="mint">کاغذ سبز نعنایی</option>
                  <option value="parchment">کاغذ کهن (پوستی)</option>
                  <option value="dark">کاغذ تاریک (شب)</option>
                </select>
                <button className="ribbon-btn large" onClick={() => setActiveModal('watermark')} title="واترمارک">
                  <i data-lucide="stamp"></i>
                  <span>واترمارک</span>
                </button>
                <button className="ribbon-btn large" onClick={() => setActiveModal('border')} title="کادر صفحه">
                  <i data-lucide="box-select"></i>
                  <span>کادر صفحه</span>
                </button>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">تم برگه و آرایش</span></div>
            </div>
          </div>

          {/* REVIEW PANEL */}
          <div className={`ribbon-panel ${activeTab === 'review' ? 'active' : ''}`}>
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <button className="ribbon-btn large" onClick={normalizePersianTypography} title="درج خودکار نیم‌فاصله‌های صحیح">
                  <i data-lucide="sparkles" style={{ color: '#d97706' }}></i>
                  <span>نیم‌فاصله هوشمند</span>
                </button>
                <button className="ribbon-btn large" onClick={convertToPersianDigits} title="تبدیل اعداد ۱۲۳ به ۱۲۳ فارسی">
                  <i data-lucide="binary"></i>
                  <span>اعداد فارسی</span>
                </button>
                <button className="ribbon-btn large" onClick={readAloud} title="خواندن صوتی متن سند">
                  <i data-lucide="volume-2" style={{ color: '#16a34a' }}></i>
                  <span>خواندن صوتی</span>
                </button>
                <button className="ribbon-btn large" onClick={() => setActiveModal('goal')} title="تعیین هدف تعداد کلمات">
                  <i data-lucide="target" style={{ color: '#2563eb' }}></i>
                  <span>هدف کلمات</span>
                </button>
                <button className="ribbon-btn large" onClick={() => setActiveModal('history')} title="تاریخچه نسخه‌ها">
                  <i data-lucide="history"></i>
                  <span>سوابق سند</span>
                </button>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">نگارش و هوشمندی</span></div>
            </div>
          </div>

          {/* VIEW PANEL */}
          <div className={`ribbon-panel ${activeTab === 'view' ? 'active' : ''}`}>
            <div className="ribbon-group">
              <div className="ribbon-group-content">
                <button className={`ribbon-btn large ${showOutline ? 'active' : ''}`} onClick={() => setShowOutline(!showOutline)} title="سایدبار ناوبری تیترها">
                  <i data-lucide="list-tree"></i>
                  <span>ناوبری سند</span>
                </button>
                <button className="ribbon-btn large" onClick={() => {
                  const r = document.querySelector('.ruler-top-container');
                  if (r) r.style.display = r.style.display === 'none' ? 'flex' : 'none';
                }} title="نمایش خط‌کش">
                  <i data-lucide="ruler"></i>
                  <span>خط‌کش</span>
                </button>
                <button className="ribbon-btn large" onClick={() => {
                  if (!document.fullscreenElement) document.documentElement.requestFullscreen().catch(() => {});
                  else document.exitFullscreen().catch(() => {});
                }} title="تمام صفحه">
                  <i data-lucide="expand"></i>
                  <span>تمام‌صفحه</span>
                </button>
              </div>
              <div className="ribbon-group-bottom"><span className="ribbon-group-label">دیدار و ساختار</span></div>
            </div>
          </div>

        </div>
      </div>

      {/* 3. MAIN WORKSPACE WITH SCROLLABLE A4 SHEETS */}
      <div className="app-layout">
        
        {/* Document Outline Sidebar (Right side) */}
        {showOutline && (
          <aside className="nav-sidebar">
            <div className="nav-sidebar-header">
              <span>فهرست تیترهای سند</span>
              <button className="modal-close-btn" onClick={() => setShowOutline(false)}><i data-lucide="x" style={{ width: 14, height: 14 }}></i></button>
            </div>
            <div className="nav-sidebar-content">
              {headings.length === 0 ? (
                <p style={{ color: '#94a3b8', fontSize: 12, textAlign: 'center', marginTop: 20 }}>تیتری در سند درج نشده است.</p>
              ) : (
                headings.map((h, i) => (
                  <div
                    key={i}
                    className={`outline-item ${h.tag}`}
                    onClick={() => {
                      const el = document.getElementById(h.id);
                      if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    }}
                  >
                    <i data-lucide="hash" style={{ width: 12, height: 12 }}></i>
                    <span>{h.text}</span>
                  </div>
                ))
              )}
            </div>
          </aside>
        )}

        {/* Workspace: Top Centimeter Ruler & A4 Sheets */}
        <main className="workspace" onClick={applyCopiedFormat}>
          
          {/* Centimeter Top Ruler */}
          <div className="ruler-top-container">
            <div className="ruler-top">
              <div className="ruler-ticks" style={{ width: '100%' }}>
                {Array.from({ length: orientation === 'landscape' ? 30 : 22 }).map((_, cm) => (
                  <div key={cm} className="ruler-tick major" style={{ right: `${cm * cmPx}px` }}>
                    <span className="ruler-tick-label">{cm}</span>
                  </div>
                ))}
              </div>
              <div className="ruler-margin-zone" style={{ right: 0, width: `${margins.right * cmPx}px` }}></div>
              <div className="ruler-margin-zone" style={{ left: 0, width: `${margins.left * cmPx}px` }}></div>
            </div>
            <div className="ruler-corner">cm</div>
          </div>

          {/* Document Viewport - BUTTER-SMOOTH SCROLL */}
          <div className="document-viewport" onContextMenu={handleContextMenu}>
            <div
              className="page-container"
              id="pageContainer"
              style={{ transform: `scale(${zoom / 100})` }}
            >
              
              {/* Dynamic Pages Rendering */}
              {pages.map((page, index) => (
                <React.Fragment key={page.id}>
                  
                  {/* Page Separator Toolbar (for Page 2 and above) */}
                  {index > 0 && (
                    <div className="page-toolbar-divider no-print" style={{ width: orientation === 'landscape' ? 1123 : 794 }}>
                      <span className="page-badge">صفحه {index + 1} از {pages.length}</span>
                      <button className="page-del-btn" onClick={() => removePage(page.id)}>
                        <i data-lucide="trash-2" style={{ width: 12, height: 12 }}></i>
                        <span>حذف صفحه</span>
                      </button>
                    </div>
                  )}

                  {/* Physical A4 Sheet */}
                  <div
                    className={`word-page paper-${paperTheme} ${orientation === 'landscape' ? 'landscape' : ''} ${columns === 2 ? 'two-columns' : columns === 3 ? 'three-columns' : ''}`}
                    style={pageStyle}
                    onClick={() => setActivePageIndex(index)}
                  >
                    {/* Header Zone */}
                    <div className="page-header-zone" contentEditable="true" ref={index === 0 ? headerContentRef : null} onInput={debounceSave} title="دابل‌کلیک برای ویرایش سربرگ">
                      <div className="page-header-content">
                        <span>مایکروسافت ورد آنلاین ۳۶۵</span>
                        <span>سند رسمی</span>
                      </div>
                    </div>

                    {/* Watermark Overlay */}
                    <div className={`page-watermark ${watermark.enabled ? 'active' : ''}`}>
                      {watermark.text}
                    </div>

                    {/* Page Body Content */}
                    <div
                      className="page-body"
                      contentEditable="true"
                      ref={el => { pageRefs.current[page.id] = el; }}
                      onInput={onEditorInput}
                      onMouseUp={applyCopiedFormat}
                      spellCheck="false"
                      dangerouslySetInnerHTML={{ __html: page.content }}
                    />

                    {/* Footer Zone */}
                    <div className="page-footer-zone" contentEditable="true" ref={index === 0 ? footerContentRef : null} onInput={debounceSave} title="دابل‌کلیک برای ویرایش پانویس">
                      <div className="page-footer-content">
                        <span>تهیه شده در سامانه هوشمند ورد</span>
                        <span>صفحه {index + 1} از {pages.length}</span>
                      </div>
                    </div>
                  </div>

                </React.Fragment>
              ))}

              {/* Add New Page Button at the bottom */}
              <button
                className="btn btn-secondary add-page-btn-wide no-print"
                onClick={addNewPage}
                style={{ width: orientation === 'landscape' ? 1123 : 794, padding: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8, border: '2px dashed #cbd5e1', background: '#ffffff' }}
              >
                <i data-lucide="plus-circle" style={{ width: 18, height: 18, color: '#185abd' }}></i>
                <span style={{ fontWeight: 600 }}>افزودن صفحه جدید به سند (صفحه {pages.length + 1})</span>
              </button>

            </div>
          </div>

          {/* Floating Find & Replace */}
          {showFindReplace && (
            <div className="find-replace-bar">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontWeight: 'bold', fontSize: 13 }}>جستجو و جایگزینی پیشرفته</span>
                <button className="modal-close-btn" onClick={() => setShowFindReplace(false)}><i data-lucide="x" style={{ width: 14, height: 14 }}></i></button>
              </div>
              <div className="find-replace-row">
                <input
                  type="text"
                  className="form-control"
                  placeholder="جستجوی عبارت..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
                <button className="btn btn-secondary" onClick={handleFindNext} style={{ padding: '6px 12px' }}>یافتن</button>
              </div>
              <div className="find-replace-row">
                <input
                  type="text"
                  className="form-control"
                  placeholder="جایگزینی با..."
                  value={replaceTerm}
                  onChange={(e) => setReplaceTerm(e.target.value)}
                />
                <button className="btn btn-primary" onClick={handleReplaceOne} style={{ padding: '6px 10px' }}>تعویض</button>
                <button className="btn btn-secondary" onClick={handleReplaceAll} style={{ padding: '6px 10px' }}>همه</button>
              </div>
            </div>
          )}

          {/* ========================================================
              CUSTOM MICROSOFT WORD RIGHT-CLICK CONTEXT MENU & MINI TOOLBAR
              ======================================================== */}
          {contextMenu.visible && (
            <div
              className="word-context-menu"
              style={{ left: contextMenu.x, top: contextMenu.y }}
              onClick={(e) => e.stopPropagation()}
            >
              {/* Floating Mini Formatting Toolbar on Top */}
              <div className="context-mini-toolbar">
                <select
                  style={{ fontSize: 11, padding: '2px 4px', width: 90 }}
                  value={fontFamily}
                  onChange={(e) => changeFontFamily(e.target.value)}
                >
                  <option value="Calibri">Calibri</option>
                  <option value="Vazirmatn">Vazirmatn</option>
                  <option value="Arial">Arial</option>
                  <option value="Tahoma">Tahoma</option>
                  <option value="B Nazanin">B Nazanin</option>
                </select>

                <select
                  style={{ fontSize: 11, padding: '2px 4px', width: 42 }}
                  value={fontSize}
                  onChange={(e) => changeFontSize(e.target.value)}
                >
                  {[9, 10, 11, 12, 14, 16, 18, 24, 36].map(s => <option key={s} value={s}>{s}</option>)}
                </select>

                <button className={`context-mini-btn ${activeStyles.bold ? 'active' : ''}`} onClick={() => execCmd('bold')}><strong>B</strong></button>
                <button className={`context-mini-btn ${activeStyles.italic ? 'active' : ''}`} onClick={() => execCmd('italic')}><em>I</em></button>
                <button className={`context-mini-btn ${activeStyles.underline ? 'active' : ''}`} onClick={() => execCmd('underline')}><u>U</u></button>
                
                <button className="context-mini-btn" onClick={() => execCmd('hiliteColor', highlightColor)}>
                  <i data-lucide="highlighter" style={{ width: 12, height: 12, color: '#ca8a04' }}></i>
                </button>
                <button className="context-mini-btn" onClick={() => execCmd('foreColor', fontColor)}>
                  <span style={{ fontWeight: 'bold', color: fontColor }}>A</span>
                </button>

                <button className="context-mini-btn" onClick={() => execCmd('justifyRight')}><i data-lucide="align-right" style={{ width: 12, height: 12 }}></i></button>
                <button className="context-mini-btn" onClick={() => execCmd('justifyCenter')}><i data-lucide="align-center" style={{ width: 12, height: 12 }}></i></button>
                <button className="context-mini-btn" onClick={() => execCmd('insertUnorderedList')}><i data-lucide="list" style={{ width: 12, height: 12 }}></i></button>
                <button className="context-mini-btn" onClick={handleFormatPainter}><i data-lucide="paint-roller" style={{ width: 12, height: 12, color: '#d97706' }}></i></button>
              </div>

              {/* Context Actions List */}
              <div className="context-menu-item" onClick={() => { execCmd('cut'); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                <div className="context-item-left">
                  <i data-lucide="scissors" style={{ width: 14, height: 14 }}></i>
                  <span>برش (Cut)</span>
                </div>
                <span className="context-item-shortcut">Ctrl+X</span>
              </div>

              <div className="context-menu-item" onClick={() => { execCmd('copy'); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                <div className="context-item-left">
                  <i data-lucide="copy" style={{ width: 14, height: 14 }}></i>
                  <span>کپی (Copy)</span>
                </div>
                <span className="context-item-shortcut">Ctrl+C</span>
              </div>

              <div className="context-menu-item" onClick={() => { execCmd('paste'); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                <div className="context-item-left">
                  <i data-lucide="clipboard" style={{ width: 14, height: 14 }}></i>
                  <span>جای‌گذاری (Paste)</span>
                </div>
                <span className="context-item-shortcut">Ctrl+V</span>
              </div>

              <div className="context-divider"></div>

              <div className="context-menu-item" onClick={() => { setActiveModal('font'); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                <div className="context-item-left">
                  <i data-lucide="type" style={{ width: 14, height: 14 }}></i>
                  <span>قلم و استایل... (Font)</span>
                </div>
              </div>

              <div className="context-menu-item" onClick={() => { setActiveModal('paragraph'); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                <div className="context-item-left">
                  <i data-lucide="pilcrow" style={{ width: 14, height: 14 }}></i>
                  <span>پاراگراف... (Paragraph)</span>
                </div>
              </div>

              {/* If right-clicked inside a table */}
              {contextMenu.targetNode?.closest('table') && (
                <>
                  <div className="context-divider"></div>
                  <div className="context-menu-item" onClick={() => { addTableRow(); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                    <div className="context-item-left">
                      <i data-lucide="rows-4" style={{ width: 14, height: 14 }}></i>
                      <span>درج سطر در جدول</span>
                    </div>
                  </div>
                  <div className="context-menu-item" onClick={() => { addTableColumn(); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                    <div className="context-item-left">
                      <i data-lucide="columns-4" style={{ width: 14, height: 14 }}></i>
                      <span>درج ستون در جدول</span>
                    </div>
                  </div>
                  <div className="context-menu-item" onClick={() => { deleteTableRow(); setContextMenu({ visible: false, x: 0, y: 0 }); }} style={{ color: '#dc2626' }}>
                    <div className="context-item-left">
                      <i data-lucide="trash-2" style={{ width: 14, height: 14 }}></i>
                      <span>حذف سطر جاری</span>
                    </div>
                  </div>
                </>
              )}

              <div className="context-divider"></div>

              <div className="context-menu-item" onClick={() => { readAloud(); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                <div className="context-item-left">
                  <i data-lucide="volume-2" style={{ width: 14, height: 14, color: '#16a34a' }}></i>
                  <span>خواندن صوتی (Read Aloud)</span>
                </div>
              </div>

              <div className="context-menu-item" onClick={() => { insertCallout('info'); setContextMenu({ visible: false, x: 0, y: 0 }); }}>
                <div className="context-item-left">
                  <i data-lucide="message-square-plus" style={{ width: 14, height: 14 }}></i>
                  <span>درج یادداشت و حاشیه</span>
                </div>
              </div>
            </div>
          )}

        </main>
      </div>

      {/* 4. WORD STATUS BAR */}
      <footer className="status-bar">
        <div className="status-left">
          <div className="status-item" onClick={() => setActiveModal('stats')}>صفحه {activePageIndex + 1} از {pages.length} | {stats.words} کلمه</div>
          <div className="status-item">{stats.chars} نویسه</div>
          <div className="status-item">راست‌به‌چپ (فارسی)</div>
          
          {/* Goal Tracker */}
          <div className="goal-tracker" onClick={() => setActiveModal('goal')} title="هدف تعداد کلمات">
            <span>تارگت: {stats.words}/{wordGoal}</span>
            <div className="goal-bar">
              <div className="goal-fill" style={{ width: `${goalPercentage}%` }}></div>
            </div>
          </div>
        </div>

        <div className="status-right">
          <div className="zoom-slider-container">
            <span style={{ fontSize: 11, cursor: 'pointer' }} onClick={() => setZoom(100)} title="بازنشانی">
              {zoom}%
            </span>
            <input
              type="range"
              min="50"
              max="200"
              value={zoom}
              step="5"
              onChange={(e) => setZoom(+e.target.value)}
            />
          </div>
        </div>
      </footer>

      {/* 5. BACKSTAGE FULLSCREEN MENU */}
      <div className={`backstage-overlay ${activeTab === 'file' ? 'active' : ''}`}>
        <div className="backstage-sidebar">
          <button className="backstage-back-btn" onClick={() => setActiveTab('home')}>
            <i data-lucide="arrow-right"></i>
            <span>بازگشت به سند</span>
          </button>
          
          <div className="backstage-nav-item active">
            <i data-lucide="layout-template" style={{ width: 18, height: 18 }}></i>
            <span>قالب‌های رسمی و آماده</span>
          </div>
          
          <div className="backstage-nav-item" onClick={() => fileInputDocxRef.current?.click()}>
            <i data-lucide="folder-open" style={{ width: 18, height: 18 }}></i>
            <span>باز کردن و ویرایش فایل (.docx یا .word)</span>
            <input type="file" ref={fileInputDocxRef} accept=".docx,.word,.json" style={{ display: 'none' }} onChange={handleOpenDocument} />
          </div>

          <div className="backstage-nav-item" onClick={() => { saveToStorage(true); setActiveTab('home'); }}>
            <i data-lucide="save" style={{ width: 18, height: 18 }}></i>
            <span>ذخیره در حافظه سیستم (Ctrl+S)</span>
          </div>

          <div className="backstage-nav-item" onClick={() => { exportDOCX(); setActiveTab('home'); }}>
            <i data-lucide="file-text" style={{ width: 18, height: 18 }}></i>
            <span>دانلود و ذخیره فایل Word (.docx)</span>
          </div>

          <div className="backstage-nav-item" onClick={() => { exportProject(); setActiveTab('home'); }}>
            <i data-lucide="hard-drive-download" style={{ width: 18, height: 18 }}></i>
            <span>ذخیره پروژه کامل سند (.word)</span>
          </div>

          <div className="backstage-nav-item" onClick={() => { exportPDF(); setActiveTab('home'); }}>
            <i data-lucide="file-down" style={{ width: 18, height: 18 }}></i>
            <span>دانلود و چاپ PDF وکتوری (فونت متصل فارسی)</span>
          </div>

          <div className="backstage-nav-item" onClick={() => { exportHTML(); setActiveTab('home'); }}>
            <i data-lucide="code" style={{ width: 18, height: 18 }}></i>
            <span>دانلود خروجی وب (HTML)</span>
          </div>

          <div className="backstage-nav-item" onClick={() => { window.print(); setActiveTab('home'); }}>
            <i data-lucide="printer" style={{ width: 18, height: 18 }}></i>
            <span>چاپ و پرینت مستقیم کاغذ A4</span>
          </div>
        </div>

        <div className="backstage-content">
          <h2 className="backstage-section-title">قالب‌های آماده یا شروع سند جدید</h2>
          <div className="templates-grid">
            <div className="template-card" onClick={() => applyTemplate('blank')}>
              <div className="template-thumbnail">
                <div className="template-preview-mockup">
                  <div style={{ width: '40%', height: 6, background: '#cbd5e1' }}></div>
                  <div style={{ width: '85%', height: 4, background: '#e2e8f0', marginTop: 12 }}></div>
                </div>
              </div>
              <div className="template-name">سند سفید (خالی)</div>
            </div>

            <div className="template-card" onClick={() => applyTemplate('officialLetter')}>
              <div className="template-thumbnail">
                <div className="template-preview-mockup" style={{ borderTop: '3px solid #103f91' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                    <div style={{ width: '30%', height: 4, background: '#94a3b8' }}></div>
                    <div style={{ width: '20%', height: 4, background: '#94a3b8' }}></div>
                  </div>
                  <div style={{ width: '60%', height: 6, background: '#103f91', marginTop: 14 }}></div>
                </div>
              </div>
              <div className="template-name">نامه اداری رسمی</div>
            </div>

            <div className="template-card" onClick={() => applyTemplate('academicPaper')}>
              <div className="template-thumbnail">
                <div className="template-preview-mockup">
                  <div style={{ width: '60%', height: 7, background: '#185abd', margin: '0 auto' }}></div>
                  <div style={{ display: 'flex', gap: 6, marginTop: 10 }}>
                    <div style={{ flex: 1, height: 60, background: '#f8fafc', border: '1px dashed #cbd5e1' }}></div>
                    <div style={{ flex: 1, height: 60, background: '#f8fafc', border: '1px dashed #cbd5e1' }}></div>
                  </div>
                </div>
              </div>
              <div className="template-name">مقاله علمی دو ستونه</div>
            </div>

            <div className="template-card" onClick={() => applyTemplate('invoice')}>
              <div className="template-thumbnail">
                <div className="template-preview-mockup">
                  <div style={{ width: '50%', height: 6, background: '#103f91' }}></div>
                  <div style={{ width: '100%', height: 40, background: '#f1f5f9', border: '1px solid #cbd5e1', marginTop: 10 }}></div>
                </div>
              </div>
              <div className="template-name">فاکتور فروش رسمی</div>
            </div>

            <div className="template-card" onClick={() => applyTemplate('report')}>
              <div className="template-thumbnail">
                <div className="template-preview-mockup">
                  <div style={{ width: '70%', height: 8, background: '#185abd' }}></div>
                  <div style={{ width: '90%', height: 4, background: '#e2e8f0', marginTop: 10 }}></div>
                </div>
              </div>
              <div className="template-name">گزارش مدیریتی</div>
            </div>

            <div className="template-card" onClick={() => applyTemplate('resume')}>
              <div className="template-thumbnail">
                <div className="template-preview-mockup">
                  <div style={{ width: '45%', height: 8, background: '#0f172a' }}></div>
                  <div style={{ width: '35%', height: 4, background: '#185abd', marginTop: 4 }}></div>
                </div>
              </div>
              <div className="template-name">رزومه شغلی مدرن</div>
            </div>
          </div>

          <div className="export-cards-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))' }}>
            <div className="export-card pdf" onClick={() => { exportPDF(); setActiveTab('home'); }}>
              <div className="export-card-icon"><i data-lucide="file-down"></i></div>
              <h3>دانلود PDF وکتوری</h3>
              <p>خروجی دقیق بدون المان‌های اضافی و با اتصال ۱۰۰٪ حروف فارسی.</p>
            </div>
            <div className="export-card docx" onClick={() => { exportDOCX(); setActiveTab('home'); }}>
              <div className="export-card-icon"><i data-lucide="file-text"></i></div>
              <h3>دانلود Word (.docx)</h3>
              <p>فایل رسمی با تمام صفحات و قابلیت ویرایش در آفیس کامپیوتر و موبایل.</p>
            </div>
            <div className="export-card" style={{ borderTop: '3px solid #16a34a' }} onClick={() => { exportProject(); setActiveTab('home'); }}>
              <div className="export-card-icon" style={{ color: '#16a34a', background: '#dcfce7' }}><i data-lucide="hard-drive-download"></i></div>
              <h3>ذخیره پروژه (.word)</h3>
              <p>ذخیره فایل پروژه جهت ادیت مجدد و بازگردانی تمام صفحات و تنظیمات.</p>
            </div>
            <div className="export-card" style={{ borderTop: '3px solid #0284c7' }} onClick={() => { fileInputDocxRef.current?.click(); }}>
              <div className="export-card-icon" style={{ color: '#0284c7', background: '#e0f2fe' }}><i data-lucide="folder-open"></i></div>
              <h3>باز کردن سند و ویرایش</h3>
              <p>بارگذاری فایل‌های docx و word برای ادامه کار و ویرایش سریع.</p>
            </div>
          </div>
        </div>
      </div>

      {/* 6. MODALS */}
      
      {/* Font Dialog Modal */}
      {activeModal === 'font' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="type" style={{ width: 18, height: 18 }}></i> پنجره تنظیمات قلم (Font Dialog)</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label>نوع فونت (Font Family):</label>
                <select className="form-control" value={fontFamily} onChange={(e) => changeFontFamily(e.target.value)}>
                  <option value="Calibri">Calibri</option>
                  <option value="Vazirmatn">Vazirmatn (وزیرمتن)</option>
                  <option value="B Nazanin">B Nazanin (بی نازنین)</option>
                  <option value="Tahoma">Tahoma (تاهوما)</option>
                  <option value="Arial">Arial</option>
                  <option value="Times New Roman">Times New Roman</option>
                </select>
              </div>
              <div className="form-group">
                <label>اندازه فونت (Size Pt):</label>
                <input type="number" className="form-control" value={fontSize} min="8" max="72" onChange={(e) => changeFontSize(e.target.value)} />
              </div>
              <div className="form-group">
                <label>رنگ قلم:</label>
                <input type="color" className="form-control" style={{ height: 38 }} value={fontColor} onChange={(e) => { setFontColor(e.target.value); execCmd('foreColor', e.target.value); }} />
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-primary" onClick={() => setActiveModal(null)}>بستن</button>
            </div>
          </div>
        </div>
      )}

      {/* Paragraph Dialog Modal */}
      {activeModal === 'paragraph' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="pilcrow" style={{ width: 18, height: 18 }}></i> پنجره تنظیمات پاراگراف (Paragraph Dialog)</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label>ترازبندی (Alignment):</label>
                <select className="form-control" onChange={(e) => execCmd(e.target.value)}>
                  <option value="justifyRight">راست‌چین (Right)</option>
                  <option value="justifyCenter">وسط‌چین (Center)</option>
                  <option value="justifyLeft">چپ‌چین (Left)</option>
                  <option value="justifyFull">تراز کامل (Justified)</option>
                </select>
              </div>
              <div className="form-group">
                <label>جهت متن (Direction):</label>
                <select className="form-control" onChange={(e) => {
                  const activeEl = pageRefs.current[pages[activePageIndex]?.id];
                  if (activeEl) { activeEl.dir = e.target.value; onEditorInput(); }
                }}>
                  <option value="rtl">راست به چپ (فارسی - RTL)</option>
                  <option value="ltr">چپ به راست (انگلیسی - LTR)</option>
                </select>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-primary" onClick={() => setActiveModal(null)}>بستن</button>
            </div>
          </div>
        </div>
      )}

      {/* Table Modal */}
      {activeModal === 'table' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="table" style={{ width: 18, height: 18 }}></i> درج جدول جدید</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label>تعداد سطرها:</label>
                <input type="number" className="form-control" value={tableRows} min="1" max="50" onChange={(e) => setTableRows(+e.target.value)} />
              </div>
              <div className="form-group">
                <label>تعداد ستون‌ها:</label>
                <input type="number" className="form-control" value={tableCols} min="1" max="15" onChange={(e) => setTableCols(+e.target.value)} />
              </div>
              <div className="form-group" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <input type="checkbox" id="styledTbl" checked={tableStyled} onChange={(e) => setTableStyled(e.target.checked)} style={{ width: 18, height: 18 }} />
                <label htmlFor="styledTbl" style={{ marginBottom: 0, cursor: 'pointer' }}>قالب‌بندی سرستون رنگی ورد</label>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setActiveModal(null)}>انصراف</button>
              <button className="btn btn-primary" onClick={insertTable}>درج جدول</button>
            </div>
          </div>
        </div>
      )}

      {/* Margins Modal */}
      {activeModal === 'margins' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="sliders" style={{ width: 18, height: 18 }}></i> تنظیم حاشیه‌های سفارشی (سانتی‌متر)</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 14 }}>
                <div className="form-group">
                  <label>حاشیه بالا (Top):</label>
                  <input type="number" className="form-control" step="0.1" value={customTop} onChange={(e) => setCustomTop(+e.target.value)} />
                </div>
                <div className="form-group">
                  <label>حاشیه پایین (Bottom):</label>
                  <input type="number" className="form-control" step="0.1" value={customBottom} onChange={(e) => setCustomBottom(+e.target.value)} />
                </div>
                <div className="form-group">
                  <label>حاشیه راست (Right):</label>
                  <input type="number" className="form-control" step="0.1" value={customRight} onChange={(e) => setCustomRight(+e.target.value)} />
                </div>
                <div className="form-group">
                  <label>حاشیه چپ (Left):</label>
                  <input type="number" className="form-control" step="0.1" value={customLeft} onChange={(e) => setCustomLeft(+e.target.value)} />
                </div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setActiveModal(null)}>انصراف</button>
              <button className="btn btn-primary" onClick={applyCustomMargins}>اعمال حاشیه‌ها</button>
            </div>
          </div>
        </div>
      )}

      {/* Watermark Modal */}
      {activeModal === 'watermark' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="stamp" style={{ width: 18, height: 18 }}></i> واترمارک پس‌زمینه</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div className="form-group" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <input type="checkbox" id="enableWm" checked={watermarkEnabled} onChange={(e) => setWatermarkEnabled(e.target.checked)} style={{ width: 18, height: 18 }} />
                <label htmlFor="enableWm" style={{ marginBottom: 0, cursor: 'pointer' }}>نمایش واترمارک در تمام صفحات</label>
              </div>
              <div className="form-group">
                <label>متن واترمارک:</label>
                <input type="text" className="form-control" value={watermarkText} onChange={(e) => setWatermarkText(e.target.value)} />
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setActiveModal(null)}>انصراف</button>
              <button className="btn btn-primary" onClick={applyWatermark}>اعمال واترمارک</button>
            </div>
          </div>
        </div>
      )}

      {/* Border Modal */}
      {activeModal === 'border' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="box-select" style={{ width: 18, height: 18 }}></i> کادر دور صفحه (Page Border)</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label>نوع خط کادر:</label>
                <select className="form-control" value={borderStyle} onChange={(e) => setBorderStyle(e.target.value)}>
                  <option value="none">بدون کادر</option>
                  <option value="solid">خط ساده (Solid)</option>
                  <option value="double">خط دوتایی (Double)</option>
                  <option value="dashed">خط‌چین (Dashed)</option>
                </select>
              </div>
              <div className="form-group">
                <label>ضخامت:</label>
                <select className="form-control" value={borderWidth} onChange={(e) => setBorderWidth(e.target.value)}>
                  <option value="1px">۱ پوینت (نازک)</option>
                  <option value="2px">۲ پوینت (متوسط)</option>
                  <option value="4px">۴ پوینت (ضخیم)</option>
                </select>
              </div>
              <div className="form-group">
                <label>رنگ کادر:</label>
                <input type="color" className="form-control" style={{ height: 38 }} value={borderColor} onChange={(e) => setBorderColor(e.target.value)} />
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setActiveModal(null)}>انصراف</button>
              <button className="btn btn-primary" onClick={applyPageBorder}>اعمال کادر</button>
            </div>
          </div>
        </div>
      )}

      {/* Symbols Modal */}
      {activeModal === 'symbols' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="pi" style={{ width: 18, height: 18 }}></i> نمادها و کاراکترهای ریاضی</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div className="symbols-picker-grid">
                {['±', '×', '÷', '√', '∞', '≈', '≠', '≤', '≥', 'π', '∑', '∫', 'α', 'β', 'θ', 'µ', 'ریال', 'تومان', '$', '€', '£', '¥', '©', '®', '™', '§', '←', '→', '↑', '↓', '✓', '★'].map((sym, i) => (
                  <button key={i} className="symbol-btn" onClick={() => insertSymbol(sym)}>
                    {sym}
                  </button>
                ))}
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setActiveModal(null)}>بستن</button>
            </div>
          </div>
        </div>
      )}

      {/* Goal Modal */}
      {activeModal === 'goal' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="target" style={{ width: 18, height: 18 }}></i> تعیین هدف تعداد کلمات سند</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div className="form-group">
                <label>هدف تعداد کلمات (Word Goal):</label>
                <input
                  type="number"
                  className="form-control"
                  value={wordGoal}
                  min="50"
                  max="50000"
                  step="50"
                  onChange={(e) => setWordGoal(+e.target.value)}
                />
              </div>
              <p style={{ fontSize: 13, color: '#64748b' }}>
                پیشرفت شما بر اساس این تارگت در نوار وضعیت پایین صفحه به صورت نمودار سبز رنگ نمایش داده می‌شود.
              </p>
            </div>
            <div className="modal-footer">
              <button className="btn btn-primary" onClick={() => setActiveModal(null)}>ذخیره هدف</button>
            </div>
          </div>
        </div>
      )}

      {/* Statistics Modal */}
      {activeModal === 'stats' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="bar-chart-2" style={{ width: 18, height: 18 }}></i> آمار جامع کلمات و زمان مطالعه</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <table style={{ width: '100%', fontSize: 14, borderCollapse: 'collapse' }}>
                <tbody>
                  <tr style={{ borderBottom: '1px solid #e2e8f0', height: 36 }}>
                    <td>تعداد صفحات:</td>
                    <td style={{ fontWeight: 'bold' }}>{pages.length} صفحه</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e2e8f0', height: 36 }}>
                    <td>تعداد کلمات:</td>
                    <td style={{ fontWeight: 'bold', color: '#103f91' }}>{stats.words}</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e2e8f0', height: 36 }}>
                    <td>تعداد نویسه‌ها (با فاصله):</td>
                    <td style={{ fontWeight: 'bold' }}>{stats.chars}</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e2e8f0', height: 36 }}>
                    <td>تعداد نویسه‌ها (بدون فاصله):</td>
                    <td style={{ fontWeight: 'bold' }}>{stats.charsNoSpaces}</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e2e8f0', height: 36 }}>
                    <td>تعداد پاراگراف‌ها:</td>
                    <td style={{ fontWeight: 'bold' }}>{stats.paragraphs}</td>
                  </tr>
                  <tr style={{ borderBottom: '1px solid #e2e8f0', height: 36 }}>
                    <td>زمان تخمینی مطالعه:</td>
                    <td style={{ fontWeight: 'bold', color: '#16a34a' }}>{stats.readTime}</td>
                  </tr>
                  <tr style={{ height: 36 }}>
                    <td>زمان تخمینی قرائت شفاهی:</td>
                    <td style={{ fontWeight: 'bold', color: '#d97706' }}>{stats.speakTime}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="modal-footer">
              <button className="btn btn-primary" onClick={() => setActiveModal(null)}>تأیید</button>
            </div>
          </div>
        </div>
      )}

      {/* History Modal */}
      {activeModal === 'history' && (
        <div className="modal-overlay">
          <div className="modal-dialog">
            <div className="modal-header">
              <div className="modal-title"><i data-lucide="history" style={{ width: 18, height: 18 }}></i> سوابق و نسخه‌های سند</div>
              <button className="modal-close-btn" onClick={() => setActiveModal(null)}><i data-lucide="x"></i></button>
            </div>
            <div className="modal-body">
              <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
                <input
                  type="text"
                  className="form-control"
                  placeholder="عنوان نسخه جدید..."
                  value={snapshotName}
                  onChange={(e) => setSnapshotName(e.target.value)}
                />
                <button className="btn btn-primary" onClick={createSnapshot} style={{ whiteSpace: 'nowrap' }}>ثبت نسخه</button>
              </div>

              <div style={{ maxHeight: 240, overflowY: 'auto' }}>
                {snapshots.length === 0 ? (
                  <p style={{ textAlign: 'center', color: '#94a3b8', fontSize: 13 }}>هنوز نسخه‌ای ثبت نشده است.</p>
                ) : (
                  snapshots.map(snap => (
                    <div key={snap.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 12px', border: '1px solid #e2e8f0', borderRadius: 6, marginBottom: 8, background: '#f8fafc' }}>
                      <div>
                        <div style={{ fontWeight: 'bold', fontSize: 13, color: '#1e293b' }}>{snap.name}</div>
                        <div style={{ fontSize: 11, color: '#64748b' }}>{snap.date}</div>
                      </div>
                      <button className="btn btn-secondary" style={{ padding: '4px 10px', fontSize: 12 }} onClick={() => restoreSnapshot(snap)}>
                        بازیابی
                      </button>
                    </div>
                  ))
                )}
              </div>
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary" onClick={() => setActiveModal(null)}>بستن</button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
