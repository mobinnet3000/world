/* ==========================================================================
   High-Fidelity PDF Export Engine (html2pdf.js + Native Vector Print)
   ========================================================================== */

const PdfExporter = {
  /**
   * Export the document directly as an A4 PDF file using html2pdf
   */
  async exportToPdf(filename = 'document.pdf', options = {}) {
    const pagesContainer = document.querySelector('.page-container');
    if (!pagesContainer) return;

    // Show export toast/loader
    const toast = document.createElement('div');
    toast.style.position = 'fixed';
    toast.style.bottom = '40px';
    toast.style.left = '50%';
    toast.style.transform = 'translateX(-50%)';
    toast.style.backgroundColor = '#103f91';
    toast.style.color = '#ffffff';
    toast.style.padding = '12px 24px';
    toast.style.borderRadius = '6px';
    toast.style.boxShadow = '0 6px 20px rgba(0,0,0,0.2)';
    toast.style.zIndex = '9999';
    toast.style.fontSize = '14px';
    toast.style.fontWeight = 'bold';
    toast.style.display = 'flex';
    toast.style.alignItems = 'center';
    toast.style.gap = '10px';
    toast.innerHTML = `<span>در حال آماده‌سازی و دانلود فایل PDF با کیفیت بالا...</span>`;
    document.body.appendChild(toast);

    try {
      // Check orientation
      const firstPage = document.querySelector('.word-page');
      const isLandscape = firstPage && firstPage.classList.contains('landscape');

      const opt = {
        margin: [0, 0, 0, 0], // exact margin handled within A4 page DOM
        filename: filename.endsWith('.pdf') ? filename : `${filename}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: {
          scale: 2, // High resolution crisp text
          useCORS: true,
          letterRendering: true,
          logging: false
        },
        jsPDF: {
          unit: 'mm',
          format: 'a4',
          orientation: isLandscape ? 'landscape' : 'portrait'
        },
        pagebreak: { mode: ['css', 'legacy'] }
      };

      // If html2pdf bundle is loaded
      if (typeof html2pdf !== 'undefined') {
        await html2pdf().set(opt).from(pagesContainer).save();
      } else {
        // Fallback to high-fidelity native print
        window.print();
      }
    } catch (err) {
      console.error('Error generating PDF:', err);
      // Fallback
      window.print();
    } finally {
      setTimeout(() => {
        if (toast.parentNode) {
          toast.parentNode.removeChild(toast);
        }
      }, 1500);
    }
  },

  /**
   * Launch native vector print dialog (Save as PDF with vector text)
   */
  printDocument() {
    window.print();
  }
};
