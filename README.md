# 🎨 AdrkhaTypograph | أدرخا تايبوجراف

<p align="center">
  <img src="https://img.shields.io/github/v/release/adrkha/AdrkhaTypograph?style=for-the-badge&color=0078D4&label=%D8%A3%D8%AD%D8%AF%D8%AB%20%D8%A5%D8%B5%D8%AF%D8%A7%D8%B1" alt="Release">
  <img src="https://img.shields.io/github/license/adrkha/AdrkhaTypograph?style=for-the-badge&color=28A745&label=%D8%A7%D9%84%D8%AA%D8%B1%D8%AE%D9%8A%D8%B5" alt="License">
  <img src="https://img.shields.io/badge/PowerPoint-2016%20--%20365-D83B01?style=for-the-badge&logo=microsoftpowerpoint&logoColor=white" alt="PowerPoint">
  <img src="https://img.shields.io/badge/.NET-Framework%204.7.2-512BD4?style=for-the-badge&logo=dotnet&logoColor=white" alt=".NET Framework">
</p>

<p align="center">
  <b>إضافة احترافية لبرنامج Microsoft PowerPoint تمنحك تحكماً شاملاً بالخصائص التايبوجرافية للخطوط العربية، وتحويل النصوص إلى أشكال متجهة (Vector Shapes) عالية الدقة مع دعم كامل لسمات OpenType والتشكيل.</b>
</p>

---

## ✨ المميزات الرئيسية

- ✒️ **تحويل النص العربي إلى أشكال (Convert Text to Shapes):**
  تحويل أي نص عربي مُشكّل ومصمم في PowerPoint إلى مسارات متجهة (Vector Shapes / Paths) قابلة للتعديل والتقسيم والتلوين بحرية كاملة دون فقدان جودة الخط.

- 🎨 **دعم ميزات OpenType المتقدمة:**
  التحكم الكامل في الخصائص الخطية مثل:
  - **البدائل الأسلوبية (Stylistic Alternates / Swashes)**
  - **التراكيب والأشكال المتصلة (Ligatures & Contextual Alternates)**
  - **الكشيدة والمد (Kashida & Extension)**
  - **ضبط التشكيل والحركات العربية الدقيقة**

- ⚡ **محرك تشكيل خطي فائق الدقة (HarfBuzz & SkiaSharp):**
  الاعتماد على أقوى المحركات العالمية لتقطيع وتشكيل الخطوط العربية (`HarfBuzzSharp`) ورسم المنحنيات الدقيقة (`SkiaSharp`) لضمان صحة اتصال الحروف ومواقع التشكيل.

- 🖥️ **واجهة جانبية سلسة (Task Pane UI):**
  لوحة تحكم جانبية مدمجة داخل PowerPoint تتيح لك معاينة وتطبيق الخصائص التايبوجرافية بنقرة واحدة.

- 🔄 **تحديثات آليّة ومستمرة:**
  فحص تلقائي لأحدث الإصدارات من مستودع GitHub للتأكد من حصولك دائماً على الأداء الأفضل والميزات الجديدة.

---

## 📥 كيفية التثبيت والاستخدام

### 1️⃣ التحميل والتثبيت المباشر:
1. ادخل على صفحة **[أحدث الإصدارات (Releases)](https://github.com/adrkha/AdrkhaTypograph/releases/latest)**.
2. حمّل ملف التثبيت المباشر `AdrkhaTypograph_v1.0.0_Setup.exe`.
3. قم بتشغيل الملف واتبع خطوات التثبيت السريعة.

### 2️⃣ التشغيل في PowerPoint:
1. افتح برنامج **Microsoft PowerPoint**.
2. ستلاحظ ظهور تبويب أو لوحة **AdrkhaTypograph** في الشريط العلوي أو اللوحة الجانبية.
3. حدد النص المراد تشكيله أو تحويله، ثم اختر الخصائص التايبوجرافية واضغط على **تحويل إلى شكل**.

---

## 💻 متطلبات النظام

- **نظام التشغيل:** Windows 10 / Windows 11 (32-bit أو 64-bit).
- **برنامج Office:** Microsoft PowerPoint 2016 / 2019 / 2021 / Microsoft 365.
- **البيئة المساعدة:** [.NET Framework 4.7.2](https://dotnet.microsoft.com/download/dotnet-framework/net472) أو أحدث.

---

## 🛠️ البناء من الكود المصدري (Developer Setup)

إذا كنت مطوراً وترغب في تعديل أو بناء المشروع بنفسك:

### المتطلبات:
- **Visual Studio 2022** أو أحدث مع تفعيل بيئة `Office/SharePoint development` (VSTO).
- **Inno Setup 6** (لإنشاء ملف التثبيت `.exe`).

### خطوات البناء:
```bash
# 1. استنساخ المستودع
git clone https://github.com/adrkha/AdrkhaTypograph.git
cd AdrkhaTypograph

# 2. استعادة الحزم (NuGet Restore)
nuget restore AdrkhaTypograph.csproj

# 3. بناء المشروع باستخدام MSBuild
msbuild AdrkhaTypograph.csproj /p:Configuration=Release /t:Rebuild

# 4. تجميع ملف التثبيت (اختياري)
iscc AdrkhaTypograph.iss
```

---

## 🤝 المساهمة والتطوير

المساهمات مرحب بها دائماً! إذا كان لديك اقتراح، تحسين، أو واجهت أي مشكلة:
1. قم بفتح **[Issue](https://github.com/adrkha/AdrkhaTypograph/issues)** لوصف المشكلة أو الاقتراح.
2. قم بعمل **Fork** للمستودع وأرسل **Pull Request** بتعديلاتك.

---

## 📜 الترخيص (License)

هذا المشروع مرخص تحت رخصة **[MIT License](LICENSE)** - يمكنك استخدامه وتعديله وتوزيعه بحرية.

---

<p align="center">
  تم التطوير بحب للخط العربي والتايبوجرافيا ️❤️ | <b>AdrkhaTypograph</b>
</p>
