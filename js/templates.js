/* ==========================================================================
   Document Templates for Word Online (Persian & Business Templates)
   ========================================================================== */

const DocumentTemplates = {
  blank: {
    title: 'سند بدون نام',
    header: '<div class="page-header-content"><span>مایکروسافت ورد آنلاین</span><span>سند متنی</span></div>',
    footer: '<div class="page-footer-content"><span>صفحه ۱</span><span>تهیه شده در ورد آنلاین</span></div>',
    content: `
      <p>متن خود را از اینجا شروع کنید...</p>
    `
  },

  officialLetter: {
    title: 'نامه اداری رسمی',
    header: `
      <div class="page-header-content" style="border-bottom: 2px solid #103f91; padding-bottom: 6px; width: 100%;">
        <div style="text-align: right; font-weight: bold; color: #103f91;">
          جمهوری اسلامی ایران<br>
          <span style="font-size: 10px; font-weight: normal; color: #555;">سازمان / شرکت توسعه فناوری</span>
        </div>
        <div style="text-align: left; font-size: 11px; line-height: 1.5;">
          شماره: ۱۴۰۳/د/۱۲۸۴۵<br>
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
        <span style="font-size: 14px; font-weight: normal;">مدیرعامل محترم شرکت پیشگامان ارتباطات</span>
      </p>
      
      <p style="font-weight: bold; margin-top: 15px;">موضوع: ارسال گزارش نهایی طرح جامع فناوری اطلاعات و امنیت شبکه</p>
      
      <p style="text-align: justify; text-indent: 30px; margin-top: 15px;">
        با سلام و احترام،<br>
        پیرو توافق‌نامه شماره ۹۸۲ مورخ ۱۴۰۳/۰۶/۱۵ بدین‌وسیله به استحضار می‌رساند فاز سوم پروژه استقرار سامانه اتوماسیون سازمانی و زیرساخت ارتباطی یکپارچه با موفقیت کامل به پایان رسیده است. کلیه آزمایش‌های امنیتی و کارایی در محیط عملیاتی با بالاترین استانداردهای فنی تأیید گردیده و نتایج تفصیلی آن به پیوست این نامه ارسال می‌گردد.
      </p>
      
      <p style="text-align: justify; text-indent: 30px;">
        خواهشمند است دستور فرمایید هماهنگی‌های لازم جهت برگزاری جلسه ارزیابی نهایی و تحویل قطعی سامانه در تاریخ ۱۴۰۳/۰۸/۲۵ در محل سالن کنفرانس مرکزی صورت پذیرد.
      </p>
      
      <p style="margin-top: 25px; margin-bottom: 40px;">
        پیشاپیش از بذل توجه و همکاری صمیمانه جنابعالی و همکاران محترم کمال تشکر را دارد.
      </p>
      
      <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-top: 20px;">
        <div style="font-size: 11px; color: #555;">
          رونوشت:<br>
          - معاونت محترم برنامه‌ریزی و توسعه<br>
          - واحد نظارت و تضمین کیفیت
        </div>
        <div style="text-align: center; font-weight: bold;">
          با احترام و تجدید ارادت<br>
          علیرضا رضایی<br>
          <span style="font-size: 12px; font-weight: normal;">مدیر پروژه و مشاور فنی</span>
        </div>
      </div>
    `
  },

  report: {
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
        <span>صفحه ۱ از ۲</span>
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
        جدول زیر خلاصه‌ای از مقایسه اهداف پیش‌بینی‌شده در بودجه مصوب در برابر ارقام تحقق‌یافته واقعی را در سه ماه گذشته نمایش می‌دهد:
      </p>

      <table class="word-styled-table">
        <thead>
          <tr>
            <th>عنوان شاخص کلیدی</th>
            <th>هدف مصوب (میلیارد تومان)</th>
            <th>عملکرد واقعی</th>
            <th>درصد تحقق</th>
            <th>وضعیت</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>فروش خدمات ابری</td>
            <td>۱۲.۵</td>
            <td>۱۴.۲</td>
            <td>۱۱۳.۶٪</td>
            <td style="color: #16a34a; font-weight: bold;">عالی</td>
          </tr>
          <tr>
            <td>پشتیبانی نرم‌افزار</td>
            <td>۶.۰</td>
            <td>۶.۸</td>
            <td>۱۱۳.۳٪</td>
            <td style="color: #16a34a; font-weight: bold;">عالی</td>
          </tr>
          <tr>
            <td>توسعه ماژول‌های اختصاصی</td>
            <td>۸.۰</td>
            <td>۷.۵</td>
            <td>۹۳.۷٪</td>
            <td style="color: #ca8a04; font-weight: bold;">مطلوب</td>
          </tr>
        </tbody>
      </table>

      <h2>۲. برنامه‌های اولویت‌دار دوره آینده</h2>
      <ul>
        <li>ارتقای زیرساخت سرورها به معماری ابری توزیع‌شده</li>
        <li>پیاده‌سازی هوش مصنوعی مولد در ماژول تحلیل داده مشتریان</li>
        <li>گسترش تیم فنی در حوزه‌های پردازش تصویر و امنیت شبکه</li>
      </ul>
    `
  },

  resume: {
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
        <strong>زبان‌ها و فریم‌ورک‌ها:</strong> JavaScript (ESNext), TypeScript, React, Next.js, Node.js, Python<br>
        <strong>پایگاه‌های داده و ابزارها:</strong> PostgreSQL, Redis, Docker, Git, REST & GraphQL APIs
      </p>
    `
  }
};
