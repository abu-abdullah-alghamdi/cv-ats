# منشئ السيرة ATS

موقع يساعد أي شخص يسوي سيرة ذاتية متوافقة مع أنظمة الفرز (ATS):

1. يدخل بإيميله برابط، بدون كلمة مرور.
2. يرفع سيرته القديمة (PDF أو Word)، والموقع يعبّي الخانات عنه.
3. يشوف تحليل يقول له وش ناقص مقارنة بالسيرة النموذجية.
4. يعدّل ويشوف المعاينة مباشرة، ويحمّل السيرة PDF أو Word.

بيانات كل مستخدم تنحفظ في حسابه على Firebase، ويقدر يرجع يعدّلها من أي جهاز.

---

## محتويات المشروع

| الملف | وش فيه |
|---|---|
| `public/index.html` | الصفحة |
| `public/assets/styles.css` | التصميم |
| `public/assets/app.js` | منطق الموقع: الخانات، المعاينة، رفع السيرة، التحليل، الخطوات، الحساب |
| `public/assets/engine.js` | تنسيق السيرة وإخراجها PDF و Word |
| `public/assets/cities.js` | قائمة المدن والمحافظات |
| `public/assets/sample.js` | سيرة المثال |
| `public/assets/cloud.js` | الاتصال بـ Firebase (الدخول وحفظ البيانات) |
| `public/assets/firebase-config.js` | **إعدادات مشروعك في Firebase، لازم تعبّيها** |
| `firestore.rules` | قواعد الحماية: كل مستخدم يشوف بياناته هو بس |
| `firebase.json` | إعدادات الاستضافة |
| `.github/workflows/deploy.yml` | رفع تلقائي للموقع كل ما تعدّل في GitHub |

ما يحتاج أي تثبيت أو بناء. الملفات ترتفع مثل ما هي.

> قبل ما تعبّي إعدادات Firebase، الموقع يشتغل في **وضع التجربة**: بدون دخول، والبيانات تنحفظ في المتصفح فقط.

---

## الإعداد (مرة وحدة)

### 1. أنشئ مشروع Firebase
1. ادخل [console.firebase.google.com](https://console.firebase.google.com) واضغط **Add project**.
2. من صفحة المشروع اضغط أيقونة الويب `</>` وسجّل تطبيق ويب.
3. انسخ قيم `firebaseConfig` اللي تطلع لك، والصقها في `public/assets/firebase-config.js` بدل القيم اللي تبدأ بـ `YOUR_`.
4. في `.firebaserc` بدّل `YOUR_PROJECT_ID` بمعرّف مشروعك (Project ID).

### 2. فعّل الدخول برابط الإيميل
1. **Authentication** ← **Get started** ← **Sign-in method**.
2. اختر **Email/Password** وفعّله، وفعّل معه **Email link (passwordless sign-in)**، واحفظ.

### 3. أنشئ قاعدة البيانات وحط قواعد الحماية
1. **Firestore Database** ← **Create database** ← اختر **Production mode** وأقرب منطقة متاحة لك.
2. افتح تبويب **Rules**، والصق محتوى ملف `firestore.rules`، واضغط **Publish**.

### 4. رقّ الخطة إلى Blaze
خطة Spark المجانية ترسل **5 رسائل دخول بالإيميل في اليوم فقط**. خطة Blaze (الدفع حسب الاستخدام) ترسل لين 25,000 رسالة في اليوم.
من **Upgrade** في أسفل القائمة اختر Blaze واربط بطاقة. شيّك على الأسعار قبل ما ترقّي.
المصدر: [Firebase Authentication limits](https://firebase.google.com/docs/auth/limits)

### 5. اربط GitHub بالرفع التلقائي
1. في Firebase: **Project settings** ← **Service accounts** ← **Generate new private key**. ينزل لك ملف JSON.
2. في GitHub افتح المستودع ← **Settings** ← **Secrets and variables** ← **Actions**:
   - تبويب **Secrets** ← **New repository secret**: الاسم `FIREBASE_SERVICE_ACCOUNT`، والقيمة محتوى ملف JSON كامل.
   - تبويب **Variables** ← **New repository variable**: الاسم `FIREBASE_PROJECT_ID`، والقيمة معرّف مشروعك.
3. **احذف ملف JSON من جهازك** بعد ما تحطه في GitHub، ولا ترفعه للمستودع أبدًا.

بعدها كل تعديل تحفظه على فرع `main` ينرفع للموقع تلقائيًا خلال دقيقة أو دقيقتين. تقدر تتابعه من تبويب **Actions**.
رابط موقعك بيكون: `https://<PROJECT_ID>.web.app`

---

## ربط دومين خاص فيك

1. في Firebase: **Hosting** ← **Add custom domain** ← اكتب الدومين، مثل `cv.example.com`.
2. أضف سجلات DNS اللي يعطيك إياها Firebase في لوحة شركة الدومين. شهادة https تطلع تلقائيًا ومجانًا، وممكن تاخذ لين 24 ساعة.
3. **مهم للدخول بالإيميل:** **Authentication** ← **Settings** ← **Authorized domains** ← **Add domain** وأضف نفس الدومين. بدون هذي الخطوة، رسائل الدخول ما تنرسل من الدومين الجديد.

المصدر: [Connect a custom domain](https://firebase.google.com/docs/hosting/custom-domain)

---

## التعديل

- عدّل الملفات من GitHub مباشرة (أيقونة القلم) أو من جهازك، واحفظ على `main`، والموقع يتحدث لحاله.
- النصوص العربية في الصفحة موجودة في `public/index.html` و `public/assets/app.js`.
- الألوان والخطوط في أول `public/assets/styles.css`.

## تجربة الموقع على جهازك

```bash
npx serve public        # أو: python3 -m http.server --directory public
```
لازم تفتحه من رابط `http://localhost`، ما يشتغل لو فتحت `index.html` مباشرة من الملفات.
