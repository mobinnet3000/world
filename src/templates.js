/* ==========================================================================
   Word Online Document Templates (Persian & Professional Templates)
   ========================================================================== */

export const DocumentTemplates = {
  blank: {
    id: 'blank',
    title: 'سند بدون نام',
    header: '<div class="page-header-content"><span>مایکروسافت ورد آنلاین</span><span>سند متنی</span></div>',
    footer: '<div class="page-footer-content"><span>صفحه ۱</span><span>تهیه شده در ورد آنلاین</span></div>',
    content: `<p>متن خود را از اینجا شروع کنید...</p>`
  },

  officialLetter: {
    id: 'officialLetter',
    title: 'نامه اداری رسمی',
    header: `
      <div class="page-header-content" style="border-bottom: 2px solid #103f91; padding-bottom: 6px; width: 100%;">
        <div style="text-align: right; font-weight: bold; color: #103f91;">
          جمهوری اسلامی ایران<br>
          <span style="font-size: 10px; font-weight: normal; color: #555;">سازمان توسعه فناوری و ارتباطات</span>
        </div>
        <div style="text-align: left; font-size: 11px; line-height: 1.5;">
          شماره: ۱۴۰۳/د/۱۴۸۲۰<br>
          تاریخ: ۱۴۰۳/۰۸/۱۲<br>
          پیوست: دارد
        </div>
      </div>
    `,
    footer: `
      <div class="page-footer-content" style="border-top: 1px solid #ccc; padding-top: 6px; width: 100%; font-size: 10px; color: #666;">
        <span>نشانی: تهران، خیابان ولیعصر، برج فناوری، طبقه ۸ | تلفن: ۸۸۸۸۸۸۸۸-۰۲۱</span>
        <span>صفحه ۱ از ۱</span>
      </div>
    `,
    content: `
      <p style="margin-top: 15px; font-weight: bold; font-size: 16px;">
        جناب آقای مهندس محمدی<br>
        <span style="font-size: 14px; font-weight: normal;">مدیرعامل محترم شرکت پیشگامان نرم‌افزار</span>
      </p>
      
      <p style="font-weight: bold; margin-top: 15px;">موضوع: ارسال گزارش نهایی طرح جامع فناوری اطلاعات و امنیت شبکه</p>
      
      <p style="text-align: justify; text-indent: 30px; margin-top: 15px;">
        با سلام و احترام،<br>
        پیرو توافق‌نامه شماره ۹۸۲ مورخ ۱۴۰۳/۰۶/۱۵ بدین‌وسیله به استحضار می‌رساند فاز سوم پروژه استقرار سامانه اتوماسیون سازمانی و زیرساخت ارتباطی یکپارچه با موفقیت کامل به پایان رسیده است. کلیه آزمایش‌های امنیتی و کارایی در محیط عملیاتی با بالاترین استانداردهای فنی تأیید گردیده و نتایج تفصیلی آن به پیوست این نامه ارسال می‌گردد.
      </p>
      
      <p style="text-align: justify; text-indent: 30px;">
        خواهشمند است دستور فرمایید هماهنگی‌های لازم جهت برگزاری جلسه ارزیابی نهایی و تحویل قطعی سامانه در تاریخ ۱۴۰۳/۰۸/۲۵ در محل سالن کنفرانس مرکزی صورت پذیرد.
      </p>
      
      <div class="signature-box-container">
        <div class="signature-stamp-card">
          <div class="signature-stamp-title">مهر دبیرخانه سازمان</div>
          <div class="signature-stamp-line">ثبت شد: ۱۴۰۳/۰۸/۱۲</div>
        </div>
        <div class="signature-stamp-card">
          <div class="signature-stamp-title">امضای مدیرکل پروژه</div>
          <div class="signature-stamp-line">علیرضا رضایی</div>
        </div>
      </div>
    `
  },

  academicPaper: {
    id: 'academicPaper',
    title: 'مقاله علمی و دانشگاهی دو ستونه',
    header: `
      <div class="page-header-content">
        <span style="color: #64748b; font-size: 11px;">مجله علمی پژوهشی سامانه‌های هوشمند | دوره ۱۴، شماره ۳</span>
        <span style="color: #185abd; font-weight: bold;">پاییز ۱۴۰۳</span>
      </div>
    `,
    footer: `
      <div class="page-footer-content">
        <span>نشریه علمی پردازش داده</span>
        <span>صفحه ۱</span>
      </div>
    `,
    content: `
      <h1 style="text-align: center; color: #103f91; margin-bottom: 8px;">بررسی روش‌های نوین پردازش زبان طبیعی فارسی با مدل‌های زبانی بزرگ</h1>
      <p style="text-align: center; font-size: 13px; color: #475569; margin-bottom: 20px;">
        دکتر احسان صادقی<sup>۱</sup>، مهندس سارا کریمی<sup>۲</sup><br>
        <span style="font-size: 11px; color: #64748b;">۱. دانشکده مهندسی کامپیوتر دانشگاه صنعتی شریف | ۲. پژوهشگاه علوم و ارتباطات</span>
      </p>

      <div style="background-color: #f1f5f9; padding: 14px; border-radius: 6px; margin-bottom: 20px; font-size: 13px; line-height: 1.7;">
        <strong>چکیده:</strong> در این مقاله، به ارزیابی دقت و عملکرد مدل‌های زبانی بزرگ در تحلیل متن‌های زبان فارسی، تشخیص معنایی و تصحیح هوشمند متون پرداخته شده است. آزمایش‌های تجربی بر روی بیش از ۱۰۰ هزار جمله نشان می‌دهد که بهره‌گیری از تنظیم دقیق مبتنی بر داده‌های ساختاریافته می‌تواند دقت را تا ۱۸٪ افزایش دهد.
        <br><strong>واژگان کلیدی:</strong> پردازش زبان طبیعی، مدل زبانی بزرگ، متن‌کاوی فارسی، هوش مصنوعی.
      </div>

      <h2>۱. مقدمه و بیان مسئله</h2>
      <p style="text-align: justify;">
        زبان فارسی به دلیل ویژگی‌های نگارشی نظیر اتصال حروف، تنوع پسوندها و کاربرد فراوان نیم‌فاصله، از چالش‌برانگیزترین زبان‌ها در حوزه پردازش خودکار متن محسوب می‌شود. در سال‌های اخیر استفاده از معماری‌های تبدیل‌کننده تحول شگرفی ایجاد کرده است.
      </p>

      <h2>۲. نتایج ارزیابی مدل‌ها</h2>
      <table class="word-styled-table">
        <thead>
          <tr>
            <th>مدل زبانی</th>
            <th>دقت معنایی (Accuracy)</th>
            <th>معیار F1</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>مدل فارسی‌پایه (Base)</td>
            <td>۸۴.۲٪</td>
            <td>۸۳.۹٪</td>
          </tr>
          <tr>
            <td>مدل بهینه‌شده ترکیبی</td>
            <td>۹۳.۵٪</td>
            <td>۹۲.۸٪</td>
          </tr>
        </tbody>
      </table>
    `
  },

  invoice: {
    id: 'invoice',
    title: 'فاکتور فروش رسمی شرکتی',
    header: `
      <div class="page-header-content" style="border-bottom: 2px solid #185abd; padding-bottom: 8px; width: 100%;">
        <div style="font-weight: bold; color: #103f91; font-size: 16px;">
          شرکت مهندسی داده‌پردازان پارس<br>
          <span style="font-size: 11px; color: #64748b; font-weight: normal;">شماره ثبت: ۵۸۴۳۲۹ | شناسه ملی: ۱۰۱۰۳۲۸۴۷</span>
        </div>
        <div style="text-align: left; font-size: 12px;">
          <strong>صورتحساب فروش کالا و خدمات</strong><br>
          شماره فاکتور: ف-۱۴۰۳/۸۹۱<br>
          تاریخ صدور: ۱۴۰۳/۰۸/۱۲
        </div>
      </div>
    `,
    footer: `
      <div class="page-footer-content" style="border-top: 1px solid #e2e8f0; font-size: 11px; color: #64748b;">
        <span>شرایط پرداخت: نقدی / واریز به حساب شماره ۱۲۳۴-۵۶۷۸-۹۰۱۲-۳۴۵۶ بانک سامان</span>
        <span>صفحه ۱</span>
      </div>
    `,
    content: `
      <div style="background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 6px; padding: 12px; margin: 16px 0; font-size: 13px;">
        <div style="display: flex; justify-content: space-between;">
          <div><strong>خریدار:</strong> شرکت فناوری نوین آینده</div>
          <div><strong>شناسه خریدار:</strong> ۱۰۳۸۴۷۲۹۱۰۰</div>
        </div>
        <div style="display: flex; justify-content: space-between; margin-top: 6px;">
          <div><strong>نشانی:</strong> تهران، سعادت‌آباد، میدان کاج، پلاک ۲۴</div>
          <div><strong>تلفن تماس:</strong> ۲۲۰۲۰۰۰۰-۰۲۱</div>
        </div>
      </div>

      <table class="word-styled-table">
        <thead>
          <tr>
            <th style="width: 40px;">ردیف</th>
            <th>شرح کالا یا خدمات</th>
            <th style="width: 60px;">تعداد</th>
            <th>مبلغ واحد (ریال)</th>
            <th>مبلغ کل (ریال)</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="text-align: center;">۱</td>
            <td>طراحی و پیاده‌سازی سامانه ابری ورد آنلاین</td>
            <td style="text-align: center;">۱</td>
            <td>۴۵۰,۰۰۰,۰۰۰</td>
            <td>۴۵۰,۰۰۰,۰۰۰</td>
          </tr>
          <tr>
            <td style="text-align: center;">۲</td>
            <td>ماژول خروجی مستقیم PDF و Word</td>
            <td style="text-align: center;">۱</td>
            <td>۱۲۰,۰۰۰,۰۰۰</td>
            <td>۱۲۰,۰۰۰,۰۰۰</td>
          </tr>
          <tr>
            <td style="text-align: center;">۳</td>
            <td>پشتیبانی فنی و نگهداری ۱۲ ماهه</td>
            <td style="text-align: center;">۱</td>
            <td>۱۸۰,۰۰۰,۰۰۰</td>
            <td>۱۸۰,۰۰۰,۰۰۰</td>
          </tr>
          <tr style="font-weight: bold; background-color: #f1f5f9;">
            <td colspan="4" style="text-align: left;">مجموع کل فاکتور:</td>
            <td style="color: #103f91;">۷۵۰,۰۰۰,۰۰۰ ریال</td>
          </tr>
        </tbody>
      </table>

      <div class="signature-box-container">
        <div class="signature-stamp-card">
          <div class="signature-stamp-title">مهر و امضای خریدار</div>
          <div class="signature-stamp-line">تأیید و دریافت کالا</div>
        </div>
        <div class="signature-stamp-card">
          <div class="signature-stamp-title">مهر و امضای فروشنده</div>
          <div class="signature-stamp-line">امور مالی و حسابداری</div>
        </div>
      </div>
    `
  },

  report: {
    id: 'report',
    title: 'گزارش تحلیلی عملکرد ماهانه',
    header: `
      <div class="page-header-content">
        <span style="font-weight: bold; color: #185abd;">شرکت بین‌المللی افق روشن | گزارش عملکرد فصلی</span>
        <span style="color: #64748b;">محرمانه و داخلی</span>
      </div>
    `,
    footer: `
      <div class="page-footer-content">
        <span>سیستم گزارش‌دهی سازمانی</span>
        <span>صفحه ۱</span>
      </div>
    `,
    content: `
      <h1 style="text-align: center; color: #103f91; margin-bottom: 6px;">گزارش جامع عملکرد استراتژیک و شاخص‌های کلیدی</h1>
      <p style="text-align: center; color: #64748b; font-size: 13px; margin-bottom: 25px;">دوره مالی پاییز ۱۴۰۳ | تهیه‌شده توسط: واحد هوش تجاری و تحلیل داده</p>
      
      <div class="word-callout">
        <div>💡</div>
        <div><strong>خلاصه مدیریتی:</strong> در این فصل، نرخ رشد جذب کاربر ۳۸٪ و درآمد حاصل از اشتراک‌های تجاری ۴۵٪ افزایش داشته که فراتر از اهداف مصوب مجمع سالیانه بوده است.</div>
      </div>

      <h2>۱. بررسی شاخص‌های مالی و درآمدی</h2>
      <p style="text-align: justify;">
        جدول زیر خلاصه‌ای از مقایسه اهداف پیش‌بینی‌شده در بودجه مصوب در برابر ارقام تحقق‌یافته واقعی را نمایش می‌دهد:
      </p>

      <table class="word-styled-table">
        <thead>
          <tr>
            <th>عنوان شاخص کلیدی</th>
            <th>هدف مصوب (میلیارد تومان)</th>
            <th>عملکرد واقعی</th>
            <th>درصد تحقق</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>فروش خدمات ابری</td>
            <td>۱۲.۵</td>
            <td>۱۴.۲</td>
            <td style="color: #16a34a; font-weight: bold;">۱۱۳.۶٪</td>
          </tr>
          <tr>
            <td>توسعه ماژول‌های اختصاصی</td>
            <td>۸.۰</td>
            <td>۷.۵</td>
            <td style="color: #ca8a04; font-weight: bold;">۹۳.۷٪</td>
          </tr>
        </tbody>
      </table>
    `
  },

  resume: {
    id: 'resume',
    title: 'رزومه حرفه‌ای مدرن',
    header: `
      <div class="page-header-content">
        <span style="font-weight: bold; color: #185abd;">رزومه کاری و سوابق تجربی</span>
        <span>به‌روزرسانی: ۱۴۰۳</span>
      </div>
    `,
    footer: `
      <div class="page-footer-content">
        <span>علی محمدی - مهندس ارشد نرم‌افزار</span>
        <span>صفحه ۱</span>
      </div>
    `,
    content: `
      <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #185abd; padding-bottom: 12px; margin-bottom: 18px;">
        <div>
          <h1 style="margin: 0; color: #103f91; font-size: 24px;">علیرضا شمس‌آبادی</h1>
          <p style="margin: 4px 0 0 0; color: #475569; font-size: 15px; font-weight: 600;">توسعه‌دهنده ارشد فول‌استک و معمار سیستم</p>
        </div>
        <div style="font-size: 12px; line-height: 1.6; text-align: left; direction: ltr; color: #334155;">
          📧 a.shams@example.com<br>
          📱 +98 912 000 0000<br>
          🌐 linkedin.com/in/example
        </div>
      </div>

      <h2>درباره من</h2>
      <p style="text-align: justify;">
        بیش از ۷ سال سابقه تخصصی در تحلیل، معماری و تولید نرم‌افزارهای تحت وب مقیاس‌پذیر سازمانی. مسلط به فریم‌ورک‌های مدرن، معماری مایکروسرویس و هدایت تیم‌های چابک نرم‌افزاری.
      </p>

      <h2>سوابق شغلی و تجارب کاری</h2>
      <div style="margin-bottom: 14px;">
        <div style="display: flex; justify-content: space-between; font-weight: bold; color: #1e293b;">
          <span>توسعه‌دهنده ارشد فرانت‌اند | شرکت فناوری پردازش نوین</span>
          <span style="color: #64748b; font-size: 12px;">۱۳۹۹ - اکنون</span>
        </div>
        <ul>
          <li>طراحی و استقرار پنل کاربری فوق‌پیشرفته با بیش از ۵۰ هزار کاربر فعال روزانه</li>
          <li>بهینه‌سازی سرعت بارگذاری اولیه نرم‌افزار تا ۶۵٪ و کاهش چشمگیر حجم باندل نهایی</li>
        </ul>
      </div>

      <h2>مهارت‌های تخصصی</h2>
      <p>
        <strong>زبان‌ها و فریم‌ورک‌ها:</strong> JavaScript, TypeScript, React, Next.js, Node.js, Python<br>
        <strong>پایگاه‌های داده و ابزارها:</strong> PostgreSQL, Redis, Docker, Git, REST & GraphQL APIs
      </p>
    `
  }
};

if (typeof window !== 'undefined') {
  window.DocumentTemplates = DocumentTemplates;
}
