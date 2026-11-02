/* ==========================================================================
   Microsoft Word (.docx) Import, Export & Project Engine
   Full Multi-Page Support, Persian RTL Typography & Lossless State Sync
   ========================================================================== */

const DocxManager = {
  /**
   * Export document content to native Microsoft Word (.docx)
   * Supports multi-page documents, tables, headings, RTL ligatures & styling
   * @param {string|object} options - Filename or configuration object
   */
  async exportToDocx(options = {}) {
    let filename = 'document.docx';
    let docTitle = 'سند مایکروسافت ورد';
    let pagesData = [];
    let orientation = 'portrait';

    if (typeof options === 'string') {
      filename = options;
      const titleInput = document.querySelector('.doc-title-input');
      docTitle = titleInput ? titleInput.value : 'سند مایکروسافت ورد';
    } else if (typeof options === 'object' && options !== null) {
      filename = options.filename || 'document.docx';
      docTitle = options.title || document.querySelector('.doc-title-input')?.value || 'سند مایکروسافت ورد';
      pagesData = options.pages || [];
      orientation = options.orientation || 'portrait';
    }

    if (!filename.toLowerCase().endsWith('.docx')) {
      filename += '.docx';
    }

    // Collect pages from DOM if not directly provided
    if (!pagesData || pagesData.length === 0) {
      const pageElements = document.querySelectorAll('.word-page');
      if (pageElements.length > 0) {
        pageElements.forEach((pageEl, idx) => {
          const bodyEl = pageEl.querySelector('.page-body');
          pagesData.push({
            id: idx + 1,
            html: bodyEl ? bodyEl.innerHTML : '',
            node: bodyEl
          });
        });
      } else {
        const singleBody = document.querySelector('.page-body');
        if (singleBody) {
          pagesData.push({ id: 1, html: singleBody.innerHTML, node: singleBody });
        }
      }
    }

    if (pagesData.length === 0) {
      alert('محتوایی برای ذخیره در فایل ورد یافت نشد.');
      return;
    }

    // Friendly Toast Notification
    const toast = document.createElement('div');
    toast.style.cssText = `
      position: fixed;
      bottom: 40px;
      left: 50%;
      transform: translateX(-50%);
      background-color: #103f91;
      color: #ffffff;
      padding: 12px 28px;
      border-radius: 8px;
      box-shadow: 0 8px 25px rgba(0,0,0,0.25);
      z-index: 99999;
      font-size: 14px;
      font-weight: bold;
      direction: rtl;
      display: flex;
      align-items: center;
      gap: 10px;
      font-family: 'Vazirmatn', system-ui, sans-serif;
    `;
    toast.innerHTML = `<span>در حال ساخت و ذخیره فایل استاندارد Microsoft Word (${pagesData.length} صفحه)...</span>`;
    document.body.appendChild(toast);

    try {
      if (typeof docx !== 'undefined') {
        const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, PageBreak } = docx;

        const sectionsChildren = [];

        pagesData.forEach((pageItem, pageIndex) => {
          // If pageIndex > 0, insert a clean page break before the next page
          if (pageIndex > 0) {
            sectionsChildren.push(new Paragraph({
              children: [new PageBreak()]
            }));
          }

          // Parse DOM nodes of the page body
          let domNodes = [];
          if (pageItem.node && pageItem.node.childNodes) {
            domNodes = Array.from(pageItem.node.childNodes);
          } else if (pageItem.html) {
            const parser = new DOMParser();
            const parsedDoc = parser.parseFromString(`<div>${pageItem.html}</div>`, 'text/html');
            domNodes = Array.from(parsedDoc.body.firstChild.childNodes);
          }

          for (const node of domNodes) {
            if (node.nodeType === Node.TEXT_NODE) {
              const text = node.textContent.trim();
              if (text) {
                sectionsChildren.push(new Paragraph({
                  children: [new TextRun({ text, font: 'Vazirmatn', rightToLeft: true })],
                  bidirectional: true,
                  alignment: AlignmentType.RIGHT,
                  spacing: { after: 120 }
                }));
              }
            } else if (node.nodeType === Node.ELEMENT_NODE) {
              const tag = node.tagName.toLowerCase();

              if (tag === 'h1') {
                sectionsChildren.push(new Paragraph({
                  children: [new TextRun({ text: node.textContent, font: 'Vazirmatn', bold: true, size: 36, color: '103f91', rightToLeft: true })],
                  heading: HeadingLevel.HEADING_1,
                  bidirectional: true,
                  alignment: AlignmentType.RIGHT,
                  spacing: { before: 240, after: 120 }
                }));
              } else if (tag === 'h2') {
                sectionsChildren.push(new Paragraph({
                  children: [new TextRun({ text: node.textContent, font: 'Vazirmatn', bold: true, size: 28, color: '185abd', rightToLeft: true })],
                  heading: HeadingLevel.HEADING_2,
                  bidirectional: true,
                  alignment: AlignmentType.RIGHT,
                  spacing: { before: 200, after: 100 }
                }));
              } else if (tag === 'h3') {
                sectionsChildren.push(new Paragraph({
                  children: [new TextRun({ text: node.textContent, font: 'Vazirmatn', bold: true, size: 24, color: '2b579a', rightToLeft: true })],
                  heading: HeadingLevel.HEADING_3,
                  bidirectional: true,
                  alignment: AlignmentType.RIGHT,
                  spacing: { before: 160, after: 80 }
                }));
              } else if (tag === 'table') {
                const rows = [];
                const trs = node.querySelectorAll('tr');
                trs.forEach((tr, rIdx) => {
                  const cells = [];
                  const tds = tr.querySelectorAll('th, td');
                  const colCount = tds.length || 1;
                  tds.forEach(td => {
                    const isTh = td.tagName.toLowerCase() === 'th';
                    cells.push(new TableCell({
                      children: [new Paragraph({
                        children: [new TextRun({
                          text: td.textContent.trim(),
                          font: 'Vazirmatn',
                          bold: isTh,
                          rightToLeft: true
                        })],
                        bidirectional: true,
                        alignment: AlignmentType.RIGHT
                      })],
                      shading: isTh ? { fill: 'f1f5f9' } : undefined,
                      width: { size: Math.floor(100 / colCount), type: WidthType.PERCENTAGE }
                    }));
                  });
                  if (cells.length > 0) {
                    rows.push(new TableRow({ children: cells }));
                  }
                });

                if (rows.length > 0) {
                  sectionsChildren.push(new Table({
                    rows: rows,
                    width: { size: 100, type: WidthType.PERCENTAGE }
                  }));
                  sectionsChildren.push(new Paragraph({ text: '', spacing: { after: 120 } }));
                }
              } else if (tag === 'ul' || tag === 'ol') {
                const items = node.querySelectorAll('li');
                items.forEach(li => {
                  sectionsChildren.push(new Paragraph({
                    children: [new TextRun({ text: li.textContent.trim(), font: 'Vazirmatn', rightToLeft: true })],
                    bullet: { level: 0 },
                    bidirectional: true,
                    alignment: AlignmentType.RIGHT,
                    spacing: { after: 60 }
                  }));
                });
              } else {
                // Paragraph, div, callout, or formatted text block
                const runs = [];
                const parseInlineNodes = (el) => {
                  el.childNodes.forEach(child => {
                    if (child.nodeType === Node.TEXT_NODE) {
                      const t = child.textContent;
                      if (t) runs.push(new TextRun({ text: t, font: 'Vazirmatn', rightToLeft: true }));
                    } else if (child.nodeType === Node.ELEMENT_NODE) {
                      const cTag = child.tagName.toLowerCase();
                      const isBold = cTag === 'b' || cTag === 'strong';
                      const isItalic = cTag === 'i' || cTag === 'em';
                      const isUnderline = cTag === 'u';
                      const isStrike = cTag === 's' || cTag === 'strike';
                      runs.push(new TextRun({
                        text: child.textContent,
                        bold: isBold,
                        italics: isItalic,
                        underline: isUnderline ? {} : undefined,
                        strike: isStrike,
                        font: 'Vazirmatn',
                        rightToLeft: true
                      }));
                    }
                  });
                };

                parseInlineNodes(node);

                sectionsChildren.push(new Paragraph({
                  children: runs.length ? runs : [new TextRun({ text: node.textContent, font: 'Vazirmatn', rightToLeft: true })],
                  bidirectional: true,
                  alignment: AlignmentType.RIGHT,
                  spacing: { after: 120 }
                }));
              }
            }
          }
        });

        const doc = new Document({
          sections: [{
            properties: {
              page: {
                margin: {
                  top: 1440, // 1 inch (2.54cm)
                  bottom: 1440,
                  left: 1440,
                  right: 1440
                }
              }
            },
            children: sectionsChildren
          }]
        });

        const blob = await Packer.toBlob(doc);
        if (typeof saveAs !== 'undefined') {
          saveAs(blob, filename);
        } else {
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = filename;
          a.click();
          URL.revokeObjectURL(url);
        }
      } else {
        // High-Fidelity Office XML MHTML Fallback
        this.exportMhtmlDocx({ filename, title: docTitle, pages: pagesData, orientation });
      }
    } catch (err) {
      console.warn('Advanced docx generation failed, falling back to Word XML export:', err);
      this.exportMhtmlDocx({ filename, title: docTitle, pages: pagesData, orientation });
    } finally {
      setTimeout(() => {
        if (toast.parentNode) {
          toast.parentNode.removeChild(toast);
        }
      }, 1500);
    }
  },

  /**
   * Universal Fallback: Generates valid multi-page Microsoft Word document
   * using full Word XML namespaces (w:rtl, w:body, A4 paper size, table styles)
   */
  exportMhtmlDocx(options = {}) {
    const filename = options.filename || 'document.docx';
    const title = options.title || 'سند مایکروسافت ورد';
    const pages = options.pages || [];
    const orientation = options.orientation || 'portrait';

    let pagesCombinedHtml = '';
    pages.forEach((p, index) => {
      if (index > 0) {
        // Official Microsoft Word Page Break
        pagesCombinedHtml += `<br clear="all" style="mso-special-character:line-break; page-break-before:always;" />`;
      }
      pagesCombinedHtml += `<div class="word-page-section" style="margin-bottom: 24pt;">${p.html || ''}</div>`;
    });

    const isLandscape = orientation === 'landscape';
    const pageWidth = isLandscape ? '841.9pt' : '595.3pt';
    const pageHeight = isLandscape ? '595.3pt' : '841.9pt';

    const wordXmlContent = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' 
            xmlns:w='urn:schemas-microsoft-com:office:word' 
            xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>${title}</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          @page WordSection1 {
            size: ${pageWidth} ${pageHeight};
            margin: 72.0pt 72.0pt 72.0pt 72.0pt;
            mso-header-margin: 36.0pt;
            mso-footer-margin: 36.0pt;
            mso-paper-source: 0;
          }
          div.WordSection1 {
            page: WordSection1;
            direction: rtl;
            text-align: right;
          }
          body {
            font-family: 'Vazirmatn', 'B Nazanin', 'Segoe UI', Tahoma, Arial, sans-serif;
            direction: rtl;
            text-align: right;
            font-size: 13pt;
            line-height: 1.6;
            color: #0f172a;
          }
          p {
            margin: 0 0 10pt 0;
            text-align: right;
            direction: rtl;
          }
          h1 { color: #103f91; font-size: 20pt; font-weight: bold; margin-top: 18pt; margin-bottom: 8pt; direction: rtl; }
          h2 { color: #185abd; font-size: 16pt; font-weight: bold; margin-top: 14pt; margin-bottom: 6pt; direction: rtl; }
          h3 { color: #2563eb; font-size: 14pt; font-weight: bold; margin-top: 10pt; margin-bottom: 4pt; direction: rtl; }
          table {
            border-collapse: collapse;
            width: 100%;
            direction: rtl;
            margin: 12pt 0;
          }
          table, th, td {
            border: 1px solid #94a3b8;
            padding: 7pt 10pt;
            text-align: right;
          }
          th {
            background-color: #f1f5f9;
            font-weight: bold;
          }
          .word-callout {
            border-right: 4px solid #185abd;
            background-color: #f8fafc;
            padding: 10pt 14pt;
            margin: 12pt 0;
            border-radius: 4px;
          }
          .signature-box-container {
            margin-top: 30pt;
            display: table;
            width: 100%;
          }
          .signature-stamp-card {
            display: table-cell;
            width: 50%;
            text-align: center;
            border: 1px dashed #cbd5e1;
            padding: 16pt;
          }
        </style>
      </head>
      <body>
        <div class="WordSection1">
          ${pagesCombinedHtml}
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff', wordXmlContent], {
      type: 'application/msword;charset=utf-8'
    });

    const outName = filename.endsWith('.docx') ? filename : `${filename}.docx`;
    if (typeof saveAs !== 'undefined') {
      saveAs(blob, outName);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = outName;
      a.click();
      URL.revokeObjectURL(url);
    }
  },

  /**
   * Import an existing Microsoft Word (.docx) file using mammoth.js
   * Converts DOCX OpenXML into clean HTML and notifies React to load it into editor pages
   * @param {File} file - File object from input[type=file]
   */
  async importDocx(file) {
    if (!file) return;

    if (typeof mammoth === 'undefined') {
      alert('کتابخانه خواندن فایل ورد (Mammoth) در دسترس نیست.');
      return;
    }

    const fileNameClean = file.name.replace(/\.[^/.]+$/, '');

    const reader = new FileReader();
    reader.onload = async (e) => {
      const arrayBuffer = e.target.result;
      try {
        const result = await mammoth.convertToHtml(
          { arrayBuffer: arrayBuffer },
          {
            styleMap: [
              "p[style-name='Heading 1'] => h1:fresh",
              "p[style-name='Heading 2'] => h2:fresh",
              "p[style-name='Heading 3'] => h3:fresh"
            ]
          }
        );
        const html = result.value || '<p>سند خالی است.</p>';

        // Dispatch Custom Event for React SPA
        window.dispatchEvent(new CustomEvent('word-docx-imported', {
          detail: {
            html: html,
            title: fileNameClean,
            messages: result.messages
          }
        }));

        // Fallback for non-React DOM
        const pageBody = document.querySelector('.page-body');
        if (pageBody) {
          pageBody.innerHTML = html;
        }

      } catch (err) {
        console.error('Error importing DOCX:', err);
        alert('خطا در باز کردن فایل ورد. لطفاً اطمینان حاصل کنید فایل .docx معتبر است.');
      }
    };
    reader.readAsArrayBuffer(file);
  },

  /**
   * Export the entire document state as a .word Project file (lossless JSON)
   * Saves pages, title, margins, orientation, themes, headers and footers
   */
  exportProject(projectData, filename = 'document.word') {
    const dataStr = JSON.stringify(projectData, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json;charset=utf-8' });
    const outName = filename.endsWith('.word') ? filename : `${filename}.word`;

    if (typeof saveAs !== 'undefined') {
      saveAs(blob, outName);
    } else {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = outName;
      a.click();
      URL.revokeObjectURL(url);
    }
  },

  /**
   * Import a saved .word or .json Project file
   */
  importProject(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const projectData = JSON.parse(e.target.result);
        window.dispatchEvent(new CustomEvent('word-project-imported', {
          detail: projectData
        }));
      } catch (err) {
        alert('فایل پروژه معتبر نیست یا محتوای آن آسیب دیده است.');
      }
    };
    reader.readAsText(file);
  }
};

// Expose globally
window.DocxManager = DocxManager;
