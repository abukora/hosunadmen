/***********************************************************************
 * منصة "حصون الإيمان" - الخادم الخلفي (Google Apps Script)
 * قاعدة البيانات: Google Sheets (يتم إنشاؤها تلقائياً أول مرة)
 ***********************************************************************/

// ===================== إعدادات عامة =====================
const APP_NAME = 'حصون الإيمان';
const SS_NAME  = 'قاعدة بيانات منصة حصون الإيمان';

const SHEET_STUDENTS = 'Students';
const SHEET_COURSES  = 'Courses';
const SHEET_LESSONS  = 'Lessons';
const SHEET_PROGRESS = 'Progress';

const STUDENTS_HEADERS = ['ID', 'الاسم', 'العمر', 'رقم الجوال', 'البريد الإلكتروني', 'تاريخ التسجيل'];
const COURSES_HEADERS  = ['ID', 'العنوان', 'الأيقونة', 'اللون', 'الترتيب'];
const LESSONS_HEADERS  = ['ID', 'معرف الكراسة', 'الترتيب', 'عنوان الدرس', 'محتوى الدرس (HTML)', 'رابط اليوتيوب', 'رابط الكتاب'];
const PROGRESS_HEADERS = ['ID', 'معرف الطالب', 'معرف الكراسة', 'رقم الدرس', 'عنوان الكراسة', 'عنوان الدرس', 'تاريخ الإتمام'];

// ===================== نقطة الدخول =====================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/sheets_2020q4_48dp.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/***********************************************************************
 * 1) إنشاء/فتح قاعدة البيانات (الشيت) تلقائياً
 ***********************************************************************/
function getSS_() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');

  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (err) {
      // الشيت المحفوظ لم يعد موجوداً، سيتم إنشاء واحد جديد بالأسفل
    }
  }

  const ss = SpreadsheetApp.create(SS_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  setupSheets_(ss);
  return ss;
}

// إنشاء كل الأوراق المطلوبة وتعبئتها بالبيانات الأساسية أول مرة فقط
function setupSheets_(ss) {
  // الطلاب
  let sh = ss.getSheetByName(SHEET_STUDENTS) || ss.insertSheet(SHEET_STUDENTS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(STUDENTS_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, STUDENTS_HEADERS.length).setFontWeight('bold').setBackground('#1B3A4B').setFontColor('#FFFFFF');
  }

  // الكراسات
  sh = ss.getSheetByName(SHEET_COURSES) || ss.insertSheet(SHEET_COURSES);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COURSES_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COURSES_HEADERS.length).setFontWeight('bold').setBackground('#1B3A4B').setFontColor('#FFFFFF');
  }

  // الدروس
  sh = ss.getSheetByName(SHEET_LESSONS) || ss.insertSheet(SHEET_LESSONS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(LESSONS_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, LESSONS_HEADERS.length).setFontWeight('bold').setBackground('#1B3A4B').setFontColor('#FFFFFF');
  }

  // التقدم
  sh = ss.getSheetByName(SHEET_PROGRESS) || ss.insertSheet(SHEET_PROGRESS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(PROGRESS_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, PROGRESS_HEADERS.length).setFontWeight('bold').setBackground('#1B3A4B').setFontColor('#FFFFFF');
  }

  // حذف الورقة الافتراضية "Sheet1" إن وُجدت وكانت فارغة
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('ورقة1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // تعبئة الكراسات والدروس الأساسية مرة واحدة فقط (إن كانت فارغة)
  const coursesSheet = ss.getSheetByName(SHEET_COURSES);
  if (coursesSheet.getLastRow() === 1) {
    seedInitialCourses_(ss);
  }
}

/***********************************************************************
 * 2) تعبئة البيانات الأساسية (10 كراسات) أول مرة فقط
 *    يمكن للمعلم لاحقاً إضافة/تعديل أي كراسة أو درس مباشرة من الشيت
 *    وستظهر فوراً في الموقع دون أي تعديل في الكود
 ***********************************************************************/
function seedInitialCourses_(ss) {
  const seed = getSeedData_();
  const coursesSheet = ss.getSheetByName(SHEET_COURSES);
  const lessonsSheet = ss.getSheetByName(SHEET_LESSONS);

  const courseRows = [];
  const lessonRows = [];
  let lessonAutoId = 1;

  seed.forEach((course, cIdx) => {
    courseRows.push([course.id, course.title, course.icon, course.color, cIdx + 1]);
    course.lessons.forEach((lesson, lIdx) => {
      lessonRows.push([
        'L' + (lessonAutoId++),
        course.id,
        lIdx + 1,
        lesson.title,
        lesson.content,
        lesson.youtube || '',
        lesson.book || ''
      ]);
    });
  });

  if (courseRows.length) {
    coursesSheet.getRange(2, 1, courseRows.length, COURSES_HEADERS.length).setValues(courseRows);
  }
  if (lessonRows.length) {
    lessonsSheet.getRange(2, 1, lessonRows.length, LESSONS_HEADERS.length).setValues(lessonRows);
  }
}

/***********************************************************************
 * 3) دوال الطلاب: تسجيل جديد / تسجيل دخول
 ***********************************************************************/
function normalizePhone_(phone) {
  return String(phone || '').replace(/[^0-9]/g, '');
}

function registerStudent(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const name  = String(data.name || '').trim();
    const age   = String(data.age || '').trim();
    const phone = normalizePhone_(data.phone);
    const email = String(data.email || '').trim();

    if (!name || !phone) {
      return { success: false, message: 'الرجاء إدخال الاسم ورقم الجوال على الأقل.' };
    }

    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_STUDENTS);
    const rows = sh.getDataRange().getValues();

    // التحقق من عدم تكرار رقم الجوال
    for (let i = 1; i < rows.length; i++) {
      if (normalizePhone_(rows[i][3]) === phone) {
        return {
          success: false,
          message: 'رقم الجوال مسجل بالفعل. استخدم تسجيل الدخول.',
          alreadyRegistered: true,
          student: rowToStudent_(rows[i])
        };
      }
    }

    const id = 'S' + new Date().getTime();
    const now = new Date();
    sh.appendRow([id, name, age, phone, email, now]);

    return {
      success: true,
      student: { id: id, name: name, age: age, phone: phone, email: email }
    };
  } catch (err) {
    return { success: false, message: 'حدث خطأ أثناء التسجيل: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function loginStudent(phone) {
  try {
    const p = normalizePhone_(phone);
    if (!p) return { success: false, message: 'الرجاء إدخال رقم الجوال.' };

    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_STUDENTS);
    const rows = sh.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (normalizePhone_(rows[i][3]) === p) {
        return { success: true, student: rowToStudent_(rows[i]) };
      }
    }
    return { success: false, message: 'لا يوجد طالب مسجل بهذا الرقم. الرجاء التسجيل أولاً.' };
  } catch (err) {
    return { success: false, message: 'حدث خطأ أثناء تسجيل الدخول: ' + err.message };
  }
}

function getStudentById(id) {
  try {
    if (!id) return { success: false };
    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_STUDENTS);
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        return { success: true, student: rowToStudent_(rows[i]) };
      }
    }
    return { success: false, message: 'لم يتم العثور على الطالب.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function rowToStudent_(row) {
  return {
    id: row[0],
    name: row[1],
    age: row[2],
    phone: row[3],
    email: row[4]
  };
}

/***********************************************************************
 * 4) دوال الكراسات والدروس (مرنة بالكامل من الشيت)
 ***********************************************************************/
function getCoursesData() {
  try {
    const ss = getSS_();
    const coursesSheet = ss.getSheetByName(SHEET_COURSES);
    const lessonsSheet = ss.getSheetByName(SHEET_LESSONS);

    const courseRows  = coursesSheet.getDataRange().getValues();
    const lessonRows  = lessonsSheet.getDataRange().getValues();

    const courses = [];
    for (let i = 1; i < courseRows.length; i++) {
      const r = courseRows[i];
      if (!r[0] && r[0] !== 0) continue; // تجاهل الصفوف الفارغة
      courses.push({
        id: r[0],
        title: r[1],
        icon: r[2] || 'fa-book',
        color: r[3] || '#1B3A4B',
        order: r[4] || 999,
        lessons: []
      });
    }
    courses.sort((a, b) => a.order - b.order);

    const byId = {};
    courses.forEach(c => byId[String(c.id)] = c);

    const tempLessons = {}; // courseId -> [ {order, title, content, youtube, book} ]
    for (let i = 1; i < lessonRows.length; i++) {
      const r = lessonRows[i];
      const courseId = String(r[1]);
      if (!r[3]) continue; // لا يوجد عنوان درس -> تجاهل
      if (!tempLessons[courseId]) tempLessons[courseId] = [];
      tempLessons[courseId].push({
        order: r[2] || 999,
        title: r[3],
        content: r[4] || '',
        youtube: r[5] || '',
        book: r[6] || ''
      });
    }

    Object.keys(tempLessons).forEach(cid => {
      tempLessons[cid].sort((a, b) => a.order - b.order);
      if (byId[cid]) {
        byId[cid].lessons = tempLessons[cid].map(l => ({
          title: l.title,
          content: l.content,
          youtube: l.youtube,
          book: l.book
        }));
      }
    });

    return { success: true, courses: courses };
  } catch (err) {
    return { success: false, message: err.message, courses: [] };
  }
}

/***********************************************************************
 * 5) دوال التقدم (تُحفظ في الشيت لتبقى محفوظة حتى لو رجع الطالب لاحقاً)
 ***********************************************************************/
function getProgress(studentId) {
  try {
    if (!studentId) return { success: true, done: [] };
    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_PROGRESS);
    const rows = sh.getDataRange().getValues();
    const done = [];
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(studentId)) {
        done.push(rows[i][2] + '_' + rows[i][3]); // "courseId_lessonIndex"
      }
    }
    return { success: true, done: done };
  } catch (err) {
    return { success: false, message: err.message, done: [] };
  }
}

function saveProgress(studentId, courseId, lessonIndex, courseTitle, lessonTitle) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!studentId) return { success: false, message: 'الطالب غير مسجل دخول.' };

    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_PROGRESS);
    const rows = sh.getDataRange().getValues();

    // تجنب التكرار
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(studentId) &&
          String(rows[i][2]) === String(courseId) &&
          String(rows[i][3]) === String(lessonIndex)) {
        return { success: true, alreadyDone: true };
      }
    }

    const id = 'P' + new Date().getTime() + Math.floor(Math.random() * 1000);
    sh.appendRow([id, studentId, courseId, lessonIndex, courseTitle, lessonTitle, new Date()]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/***********************************************************************
 * 5.5) لوحة التحكم (إضافة/تعديل/حذف الكراسات والدروس من الموقع مباشرة)
 ***********************************************************************/
const ADMIN_PASSWORD_DEFAULT = 'admin123'; // يفضل تغييرها من خصائص المشروع (ADMIN_PASSWORD)

function getAdminPassword_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('ADMIN_PASSWORD') || ADMIN_PASSWORD_DEFAULT;
}

function adminLogin(password) {
  try {
    if (String(password || '') === getAdminPassword_()) {
      return { success: true };
    }
    return { success: false, message: 'كلمة المرور غير صحيحة.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// تغيير كلمة مرور المعلم (تتطلب كلمة المرور الحالية)
function adminChangePassword(oldPassword, newPassword) {
  try {
    if (String(oldPassword || '') !== getAdminPassword_()) {
      return { success: false, message: 'كلمة المرور الحالية غير صحيحة.' };
    }
    if (!newPassword || String(newPassword).length < 4) {
      return { success: false, message: 'كلمة المرور الجديدة قصيرة جداً (4 أحرف على الأقل).' };
    }
    PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', String(newPassword));
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// جلب كل البيانات (كراسات + دروس) لعرضها في لوحة التحكم
function getAdminBundle(password) {
  try {
    const auth = adminLogin(password);
    if (!auth.success) return auth;

    const ss = getSS_();
    const coursesSheet = ss.getSheetByName(SHEET_COURSES);
    const lessonsSheet = ss.getSheetByName(SHEET_LESSONS);
    const courseRows = coursesSheet.getDataRange().getValues();
    const lessonRows = lessonsSheet.getDataRange().getValues();

    const courses = [];
    for (let i = 1; i < courseRows.length; i++) {
      const r = courseRows[i];
      if (!r[0] && r[0] !== 0) continue;
      courses.push({ id: r[0], title: r[1], icon: r[2] || 'fa-book', color: r[3] || '#1B3A4B', order: r[4] || 999 });
    }
    courses.sort((a, b) => a.order - b.order);

    const lessons = [];
    for (let i = 1; i < lessonRows.length; i++) {
      const r = lessonRows[i];
      if (!r[3]) continue;
      lessons.push({
        rowId: r[0], courseId: String(r[1]), order: r[2] || 999,
        title: r[3], content: r[4] || '', youtube: r[5] || '', book: r[6] || ''
      });
    }
    lessons.sort((a, b) => a.order - b.order);

    return { success: true, courses: courses, lessons: lessons };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ---- الكراسات ----
function addCourse(password, course) {
  try {
    const auth = adminLogin(password);
    if (!auth.success) return auth;

    const title = String(course.title || '').trim();
    if (!title) return { success: false, message: 'عنوان الكراسة مطلوب.' };

    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_COURSES);
    const rows = sh.getDataRange().getValues();

    let maxId = 0;
    let maxOrder = 0;
    for (let i = 1; i < rows.length; i++) {
      const n = Number(rows[i][0]);
      if (!isNaN(n) && n > maxId) maxId = n;
      const o = Number(rows[i][4]);
      if (!isNaN(o) && o > maxOrder) maxOrder = o;
    }
    const newId = maxId + 1;

    sh.appendRow([newId, title, course.icon || 'fa-book', course.color || '#1B3A4B', maxOrder + 1]);
    return { success: true, course: { id: newId, title: title, icon: course.icon || 'fa-book', color: course.color || '#1B3A4B' } };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateCourse(password, courseId, course) {
  try {
    const auth = adminLogin(password);
    if (!auth.success) return auth;

    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_COURSES);
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(courseId)) {
        sh.getRange(i + 1, 2).setValue(course.title || rows[i][1]);
        sh.getRange(i + 1, 3).setValue(course.icon || rows[i][2]);
        sh.getRange(i + 1, 4).setValue(course.color || rows[i][3]);
        return { success: true };
      }
    }
    return { success: false, message: 'الكراسة غير موجودة.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteCourse(password, courseId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const auth = adminLogin(password);
    if (!auth.success) return auth;

    const ss = getSS_();
    const coursesSheet = ss.getSheetByName(SHEET_COURSES);
    const lessonsSheet = ss.getSheetByName(SHEET_LESSONS);

    const courseRows = coursesSheet.getDataRange().getValues();
    for (let i = courseRows.length - 1; i >= 1; i--) {
      if (String(courseRows[i][0]) === String(courseId)) {
        coursesSheet.deleteRow(i + 1);
      }
    }

    const lessonRows = lessonsSheet.getDataRange().getValues();
    for (let i = lessonRows.length - 1; i >= 1; i--) {
      if (String(lessonRows[i][1]) === String(courseId)) {
        lessonsSheet.deleteRow(i + 1);
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

// ---- الدروس ----
function addLesson(password, lesson) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const auth = adminLogin(password);
    if (!auth.success) return auth;

    const courseId = String(lesson.courseId || '').trim();
    const title = String(lesson.title || '').trim();
    if (!courseId || !title) return { success: false, message: 'الكراسة وعنوان الدرس مطلوبان.' };

    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_LESSONS);
    const rows = sh.getDataRange().getValues();

    let maxNum = 0;
    let maxOrder = 0;
    for (let i = 1; i < rows.length; i++) {
      const m = String(rows[i][0]).match(/^L(\d+)$/);
      if (m && Number(m[1]) > maxNum) maxNum = Number(m[1]);
      if (String(rows[i][1]) === courseId) {
        const o = Number(rows[i][2]);
        if (!isNaN(o) && o > maxOrder) maxOrder = o;
      }
    }
    const newId = 'L' + (maxNum + 1);

    sh.appendRow([newId, courseId, maxOrder + 1, title, lesson.content || '', lesson.youtube || '', lesson.book || '']);
    return { success: true, lesson: { rowId: newId, courseId: courseId, order: maxOrder + 1, title: title, content: lesson.content || '', youtube: lesson.youtube || '', book: lesson.book || '' } };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function updateLesson(password, rowId, lesson) {
  try {
    const auth = adminLogin(password);
    if (!auth.success) return auth;

    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_LESSONS);
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(rowId)) {
        sh.getRange(i + 1, 4).setValue(lesson.title || rows[i][3]);
        sh.getRange(i + 1, 5).setValue(lesson.content != null ? lesson.content : rows[i][4]);
        sh.getRange(i + 1, 6).setValue(lesson.youtube != null ? lesson.youtube : rows[i][5]);
        sh.getRange(i + 1, 7).setValue(lesson.book != null ? lesson.book : rows[i][6]);
        return { success: true };
      }
    }
    return { success: false, message: 'الدرس غير موجود.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function deleteLesson(password, rowId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const auth = adminLogin(password);
    if (!auth.success) return auth;

    const ss = getSS_();
    const sh = ss.getSheetByName(SHEET_LESSONS);
    const rows = sh.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(rowId)) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'الدرس غير موجود.' };
  } catch (err) {
    return { success: false, message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/***********************************************************************
 * 6) دالة تهيئة يدوية اختيارية (لتشغيلها مرة من محرر Apps Script
 *    إن أردت إنشاء الشيت قبل أول زيارة للموقع)
 ***********************************************************************/
function manualSetup() {
  getSS_();
  return 'تم إعداد قاعدة البيانات بنجاح. رابط الشيت: ' + getSS_().getUrl();
}

/***********************************************************************
 * 7) بيانات التأسيس (الكراسات العشر الأصلية) - تُستخدم مرة واحدة فقط
 *    عند إنشاء الشيت لأول مرة حتى لا تُفقد أي بيانات
 ***********************************************************************/
function getSeedData_() {
  const YT = 'https://youtube.com/@abukora?si=Nu3-ZETtW-QPsLW6';
  const BK = 'https://www.google.com/?hl=ar';
  return [
    {
      id: 1, title: 'أركان الإيمان', icon: 'fa-star', color: '#D4AF37',
      lessons: [
        { title: 'ما هو التوحيد؟', content: '<p><strong>الفقرة الأساسية:</strong> التوحيد هو إفراد الله بالعبادة، وهو الهدف من خلقنا.</p><p><strong>الدليل:</strong> ﴿وَمَا خَلَقْتُ الْجِنَّ وَالْإِنسَ إِلَّا لِيَعْبُدُونِ﴾</p><p><strong>الشرح:</strong> التوحيد أساس الدين، يشمل توحيد الربوبية والألوهية والأسماء والصفات.</p><p><strong>النموذج التطبيقي:</strong> ردد لا إله إلا الله مئة مرة.</p><p><strong>الفائدة:</strong> معرفة الهدف من الحياة تمنح السعادة.</p>', youtube: YT, book: BK },
        { title: 'أنواع التوحيد', content: '<p><strong>الفقرة الأساسية:</strong> التوحيد ثلاثة أنواع: ربوبية، ألوهية، أسماء وصفات.</p><p><strong>الدليل:</strong> سورة الإخلاص.</p><p><strong>الشرح:</strong> الربوبية: إفراد الله بالخلق والملك. الألوهية: إفراد الله بالعبادة. الأسماء والصفات: إثبات ما أثبته لنفسه.</p><p><strong>النموذج:</strong> تأمل سورة الإخلاص واستخرج أنواع التوحيد.</p><p><strong>الفائدة:</strong> التوحيد يحمي من الشرك.</p>', youtube: YT, book: BK },
        { title: 'الإيمان بأسماء الله', content: '<p><strong>الفقرة:</strong> لله الأسماء الحسنى والصفات العليا، نثبتها بلا تشبيه.</p><p><strong>الدليل:</strong> ﴿لَيْسَ كَمِثْلِهِ شَيْءٌ وَهُوَ السَّمِيعُ الْبَصِيرُ﴾</p><p><strong>الشرح:</strong> نؤمن بأن الله سميع بصير لكن صفاته لا تشبه صفات المخلوقين.</p><p><strong>النموذج:</strong> تعلم اسمًا من أسماء الله اليوم ومعناه.</p><p><strong>الفائدة:</strong> معرفة الله تزيد الإيمان.</p>', youtube: YT, book: BK },
        { title: 'محبة الله وخوفه ورجاؤه', content: '<p><strong>الفقرة:</strong> عبادة الله تقوم على الحب والخوف والرجاء.</p><p><strong>الدليل:</strong> ﴿يَدْعُونَ رَبَّهُمْ خَوْفًا وَطَمَعًا﴾</p><p><strong>الشرح:</strong> المؤمن يعبد الله محبة وخوفًا من عقابه ورجاءً لرحمته.</p><p><strong>النموذج:</strong> ادعُ الله اليوم بخوف ورجاء.</p>', youtube: YT, book: BK },
        { title: 'الإيمان بالملائكة', content: '<p><strong>الفقرة:</strong> الملائكة مخلوقات نورانية لا يعصون الله.</p><p><strong>الدليل:</strong> ﴿لَا يَسْتَكْبِرُونَ عَنْ عِبَادَتِهِ﴾</p><p><strong>النموذج:</strong> استشعر وجود الملائكة الكاتبين.</p>', youtube: YT, book: BK },
        { title: 'الإيمان بالكتب', content: '<p><strong>الفقرة:</strong> أنزل الله كتبًا على رسله، والقرآن خاتمها.</p><p><strong>الدليل:</strong> ﴿إِنَّا نَحْنُ نَزَّلْنَا الذِّكْرَ وَإِنَّا لَهُ لَحَافِظُونَ﴾</p><p><strong>النموذج:</strong> اقرأ سورة البقرة وابحث عن آيات عن الإيمان بالكتب.</p>', youtube: YT, book: BK },
        { title: 'الإيمان بالرسل', content: '<p><strong>الفقرة:</strong> الإيمان بجميع الرسل، ومحمد خاتمهم.</p><p><strong>الدليل:</strong> ﴿مَا كَانَ مُحَمَّدٌ أَبَا أَحَدٍ مِّن رِّجَالِكُمْ وَلَٰكِن رَّسُولَ اللَّهِ وَخَاتَمَ النَّبِيِّينَ﴾</p>', youtube: YT, book: BK },
        { title: 'الإيمان باليوم الآخر', content: '<p><strong>الفقرة:</strong> البعث والحساب والجنة والنار.</p><p><strong>الدليل:</strong> ﴿فَمَن يَعْمَلْ مِثْقَالَ ذَرَّةٍ خَيْرًا يَرَهُ﴾</p>', youtube: YT, book: BK },
        { title: 'الإيمان بالقدر', content: '<p><strong>الفقرة:</strong> الإيمان بالقدر خيره وشره.</p><p><strong>الدليل:</strong> ﴿إِنَّا كُلَّ شَيْءٍ خَلَقْنَاهُ بِقَدَرٍ﴾</p>', youtube: YT, book: BK },
        { title: 'الشرك بالله', content: '<p><strong>الفقرة:</strong> أعظم ذنب هو الشرك بالله.</p><p><strong>الدليل:</strong> ﴿إِنَّ اللَّهَ لَا يَغْفِرُ أَن يُشْرَكَ بِهِ﴾</p>', youtube: YT, book: BK },
        { title: 'البدعة وخطرها', content: '<p><strong>الفقرة:</strong> كل بدعة في الدين ضلالة.</p><p><strong>الدليل:</strong> حديث: «من أحدث في أمرنا هذا ما ليس منه فهو رد».</p>', youtube: YT, book: BK },
        { title: 'الولاء والبراء', content: '<p><strong>الفقرة:</strong> الولاء للمؤمنين والبراء من الكفار.</p><p><strong>الدليل:</strong> ﴿لَا تَجِدُ قَوْمًا يُؤْمِنُونَ بِاللَّهِ وَالْيَوْمِ الْآخِرِ يُوَادُّونَ مَنْ حَادَّ اللَّهَ وَرَسُولَهُ﴾</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 2, title: 'النور المبين', icon: 'fa-mosque', color: '#2E8B57',
      lessons: [
        { title: 'الماء والطهارة', content: '<p><strong>الفقرة:</strong> الطهارة شرط لصحة الصلاة.</p><p><strong>الدليل:</strong> حديث: «الطهور شطر الإيمان».</p><p><strong>النموذج:</strong> تدرب على الوضوء الصحيح.</p>', youtube: YT, book: BK },
        { title: 'فروض الوضوء', content: '<p><strong>الفقرة:</strong> النية، غسل الوجه، اليدين، مسح الرأس، غسل الرجلين.</p><p><strong>الدليل:</strong> آية المائدة: 6.</p>', youtube: YT, book: BK },
        { title: 'نواقض الوضوء', content: '<p><strong>الفقرة:</strong> الخارج من السبيلين، النوم العميق.</p><p><strong>الدليل:</strong> حديث: «لا ينصرف حتى يسمع صوتًا أو يجد ريحًا».</p>', youtube: YT, book: BK },
        { title: 'أهمية الصلاة', content: '<p><strong>الفقرة:</strong> الصلاة عمود الدين.</p><p><strong>الدليل:</strong> حديث: «العهد الذي بيننا وبينهم الصلاة».</p>', youtube: YT, book: BK },
        { title: 'شروط الصلاة وأركانها', content: '<p><strong>الفقرة:</strong> الطهارة، استقبال القبلة، النية.</p>', youtube: YT, book: BK },
        { title: 'صفة صلاة النبي ﷺ', content: '<p><strong>الفقرة:</strong> تكبيرة الإحرام، قراءة الفاتحة، الركوع، السجود...</p>', youtube: YT, book: BK },
        { title: 'صلاة الجماعة', content: '<p><strong>الفقرة:</strong> صلاة الجماعة أفضل بسبع وعشرين درجة.</p>', youtube: YT, book: BK },
        { title: 'أحكام الصيام', content: '<p><strong>الفقرة:</strong> الإمساك عن المفطرات من الفجر إلى المغرب.</p>', youtube: YT, book: BK },
        { title: 'الزكاة والصدقة', content: '<p><strong>الفقرة:</strong> الزكاة فريضة، والصدقة تطوع.</p>', youtube: YT, book: BK },
        { title: 'الصدق والأمانة', content: '<p><strong>الفقرة:</strong> الصدق واجب في البيع والشراء.</p>', youtube: YT, book: BK },
        { title: 'آداب الطعام', content: '<p><strong>الفقرة:</strong> التسمية، الأكل باليمين.</p>', youtube: YT, book: BK },
        { title: 'آداب النوم', content: '<p><strong>الفقرة:</strong> النوم على الشق الأيمن.</p>', youtube: YT, book: BK },
        { title: 'بر الوالدين', content: '<p><strong>الفقرة:</strong> طاعتهما في المعروف.</p>', youtube: YT, book: BK },
        { title: 'حفظ اللسان', content: '<p><strong>الفقرة:</strong> اجتناب الغيبة والنميمة.</p>', youtube: YT, book: BK },
        { title: 'الأمر بالمعروف', content: '<p><strong>الفقرة:</strong> من رأى منكم منكرًا فليغيره.</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 3, title: 'مع الحبيب ﷺ', icon: 'fa-heart', color: '#C0392B',
      lessons: [
        { title: 'نسبه ومولده', content: '<p>محمد بن عبد الله، ولد في مكة عام الفيل.</p>', youtube: YT, book: BK },
        { title: 'طفولته', content: '<p>كفله جده ثم عمه، حادثة شق الصدر.</p>', youtube: YT, book: BK },
        { title: 'نزول الوحي', content: '<p>في غار حراء، أول ما نزل: اقرأ.</p>', youtube: YT, book: BK },
        { title: 'الصدع بالدعوة', content: '<p>بدأ الدعوة سرًا ثم جهرًا.</p>', youtube: YT, book: BK },
        { title: 'الإسراء والمعراج', content: '<p>أسري به من المسجد الحرام إلى الأقصى.</p>', youtube: YT, book: BK },
        { title: 'الهجرة إلى المدينة', content: '<p>صحبة أبي بكر، قصة الغار.</p>', youtube: YT, book: BK },
        { title: 'غزوة بدر', content: '<p>أول معركة فاصلة.</p>', youtube: YT, book: BK },
        { title: 'غزوة أحد', content: '<p>درس في طاعة القائد.</p>', youtube: YT, book: BK },
        { title: 'غزوة الخندق', content: '<p>حفر الخندق وفكرة سلمان.</p>', youtube: YT, book: BK },
        { title: 'صلح الحديبية', content: '<p>دروس في الحكمة.</p>', youtube: YT, book: BK },
        { title: 'فتح مكة', content: '<p>العفو عند المقدرة.</p>', youtube: YT, book: BK },
        { title: 'حجة الوداع', content: '<p>إكمال الدين.</p>', youtube: YT, book: BK },
        { title: 'وفاته ﷺ', content: '<p>اللحاق بالرفيق الأعلى.</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 4, title: 'أصول الفهم السديد', icon: 'fa-brain', color: '#8E44AD',
      lessons: [
        { title: 'مصادر التلقي', content: '<p>الكتاب والسنة.</p>', youtube: YT, book: BK },
        { title: 'فهم السلف', content: '<p>الصحابة والتابعون.</p>', youtube: YT, book: BK },
        { title: 'رد المتشابه إلى المحكم', content: '<p>قاعدة أساسية.</p>', youtube: YT, book: BK },
        { title: 'الجمع بين النصوص', content: '<p>لا تناقض في الشريعة.</p>', youtube: YT, book: BK },
        { title: 'العقل والنقل', content: '<p>لا تعارض بين العقل الصريح والنقل الصحيح.</p>', youtube: YT, book: BK },
        { title: 'الإجماع حجة', content: '<p>ما أجمع عليه العلماء.</p>', youtube: YT, book: BK },
        { title: 'قطعيات الشريعة', content: '<p>لا تقبل التأويل.</p>', youtube: YT, book: BK },
        { title: 'مراعاة المآلات', content: '<p>سد الذرائع.</p>', youtube: YT, book: BK },
        { title: 'ربط الفروع بالأصول', content: '<p>وحدة الشريعة.</p>', youtube: YT, book: BK },
        { title: 'التثبت والورع', content: '<p>عدم الفتوى بغير علم.</p>', youtube: YT, book: BK },
        { title: 'العبادات توقيفية', content: '<p>لا نخترع عبادة.</p>', youtube: YT, book: BK },
        { title: 'التدرج في التعلم', content: '<p>منهج السلف.</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 5, title: 'هذه أخلاقي', icon: 'fa-smile', color: '#E67E22',
      lessons: [
        { title: 'آداب النوم', content: '<p>النوم على طهارة، الأذكار.</p>', youtube: YT, book: BK },
        { title: 'آداب الطعام', content: '<p>التسمية والحمد.</p>', youtube: YT, book: BK },
        { title: 'آداب اللباس', content: '<p>البدء باليمين.</p>', youtube: YT, book: BK },
        { title: 'آداب قضاء الحاجة', content: '<p>الاستئذان والذكر.</p>', youtube: YT, book: BK },
        { title: 'الصدق', content: '<p>الصدق يهدي إلى البر.</p>', youtube: YT, book: BK },
        { title: 'الأمانة', content: '<p>أداء الحقوق.</p>', youtube: YT, book: BK },
        { title: 'الحياء', content: '<p>شعبة من الإيمان.</p>', youtube: YT, book: BK },
        { title: 'العفة', content: '<p>غض البصر.</p>', youtube: YT, book: BK },
        { title: 'بر الوالدين', content: '<p>خفض الجناح.</p>', youtube: YT, book: BK },
        { title: 'صلة الرحم', content: '<p>زيارة الأقارب.</p>', youtube: YT, book: BK },
        { title: 'حفظ اللسان', content: '<p>اجتناب الغيبة.</p>', youtube: YT, book: BK },
        { title: 'الرفق', content: '<p>الرفق لا يأتي إلا بخير.</p>', youtube: YT, book: BK },
        { title: 'أذكار الصباح والمساء', content: '<p>آية الكرسي والمعوذات.</p>', youtube: YT, book: BK },
        { title: 'أذكار الأحوال', content: '<p>عند الأكل، النوم، الدخول.</p>', youtube: YT, book: BK },
        { title: 'فضل الذكر', content: '<p>الذكر حياة القلوب.</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 6, title: 'شبهات وردود', icon: 'fa-shield-alt', color: '#2980B9',
      lessons: [
        { title: 'من خلق الله؟', content: '<p>سؤال متناقض، الله الأول.</p>', youtube: YT, book: BK },
        { title: 'لماذا لا نرى الله؟', content: '<p>﴿لَا تُدْرِكُهُ الْأَبْصَارُ﴾.</p>', youtube: YT, book: BK },
        { title: 'القرآن من تأليف محمد؟', content: '<p>أمي لا يقرأ ولا يكتب.</p>', youtube: YT, book: BK },
        { title: 'تناقضات في القرآن؟', content: '<p>لا تناقض، النسخ رحمة.</p>', youtube: YT, book: BK },
        { title: 'تعدد زوجات النبي', content: '<p>زواج سياسي واجتماعي.</p>', youtube: YT, book: BK },
        { title: 'الإسلام انتشر بالسيف؟', content: '<p>إندونيسيا مثال.</p>', youtube: YT, book: BK },
        { title: 'المرأة في الإسلام', content: '<p>تكريم لا اضطهاد.</p>', youtube: YT, book: BK },
        { title: 'الحدود الإسلامية', content: '<p>الردع قبل التطبيق.</p>', youtube: YT, book: BK },
        { title: 'ضرب الزوجة', content: '<p>ضرب غير مبرح، آخر الحلول.</p>', youtube: YT, book: BK },
        { title: 'لماذا خلق الله الشر؟', content: '<p>ابتلاء وتمحيص.</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 7, title: 'أنا رقمي', icon: 'fa-mobile-alt', color: '#16A085',
      lessons: [
        { title: 'الوقت', content: '<p>نعمة الصحة والفراغ.</p>', youtube: YT, book: BK },
        { title: 'غض البصر الرقمي', content: '<p>النظرة سهم مسموم.</p>', youtube: YT, book: BK },
        { title: 'الموسيقى', content: '<p>البدائل الحلال.</p>', youtube: YT, book: BK },
        { title: 'التعليقات', content: '<p>حصائد الألسنة الرقمية.</p>', youtube: YT, book: BK },
        { title: 'القدوة الرقمية', content: '<p>اختيار المؤثرين.</p>', youtube: YT, book: BK },
        { title: 'الإدمان الرقمي', content: '<p>التدرج في التخفيف.</p>', youtube: YT, book: BK },
        { title: 'الألعاب الإلكترونية', content: '<p>بين الحلال والحرام.</p>', youtube: YT, book: BK },
        { title: 'التقنية نعمة', content: '<p>شكر النعمة.</p>', youtube: YT, book: BK },
        { title: 'أذكاري الرقمية', content: '<p>تطبيقات القرآن.</p>', youtube: YT, book: BK },
        { title: 'العهد مع الله', content: '<p>استخدام الشاشة في الخير.</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 8, title: 'أنا أقوى', icon: 'fa-fist-raised', color: '#E74C3C',
      lessons: [
        { title: 'لماذا خلقت؟', content: '<p>لعبادة الله.</p>', youtube: YT, book: BK },
        { title: 'التوكل', content: '<p>اليقين بالله.</p>', youtube: YT, book: BK },
        { title: 'الغضب', content: '<p>لا تغضب.</p>', youtube: YT, book: BK },
        { title: 'الحزن والخوف', content: '<p>دعاء الهم.</p>', youtube: YT, book: BK },
        { title: 'الصداقة', content: '<p>المرء على دين خليله.</p>', youtube: YT, book: BK },
        { title: 'الفشل', content: '<p>التوبة والرجوع.</p>', youtube: YT, book: BK },
        { title: 'الثقة بالنفس', content: '<p>المؤمن القوي.</p>', youtube: YT, book: BK },
        { title: 'العناية بالجسد', content: '<p>حق الجسد.</p>', youtube: YT, book: BK },
        { title: 'تزكية القلب', content: '<p>المراقبة.</p>', youtube: YT, book: BK },
        { title: 'ختام', content: '<p>أنا أقوى بإيماني.</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 9, title: 'حياة صحية', icon: 'fa-apple-alt', color: '#27AE60',
      lessons: [
        { title: 'الجسد أمانة', content: '<p>نعمة الصحة.</p>', youtube: YT, book: BK },
        { title: 'هدي النبي في الطعام', content: '<p>ثلث لطعامك.</p>', youtube: YT, book: BK },
        { title: 'آداب الطعام', content: '<p>الأكل باليمين.</p>', youtube: YT, book: BK },
        { title: 'الرياضة', content: '<p>المؤمن القوي.</p>', youtube: YT, book: BK },
        { title: 'النوم', content: '<p>النوم على الشق الأيمن.</p>', youtube: YT, book: BK },
        { title: 'النظافة', content: '<p>الطهور شطر الإيمان.</p>', youtube: YT, book: BK },
        { title: 'العادات الضارة', content: '<p>لا ضرر ولا ضرار.</p>', youtube: YT, book: BK }
      ]
    },
    {
      id: 10, title: 'أنا في الحياة', icon: 'fa-user-graduate', color: '#8E44AD',
      lessons: [
        { title: 'تحمل المسؤولية', content: '<p>كلكم راع.</p>', youtube: YT, book: BK },
        { title: 'إدارة الوقت', content: '<p>اغتنم خمسًا قبل خمس.</p>', youtube: YT, book: BK },
        { title: 'العمل الجماعي', content: '<p>يد الله مع الجماعة.</p>', youtube: YT, book: BK },
        { title: 'اتخاذ القرار', content: '<p>الاستخارة والاستشارة.</p>', youtube: YT, book: BK },
        { title: 'حل المشكلات', content: '<p>فإن مع العسر يسرًا.</p>', youtube: YT, book: BK },
        { title: 'التواصل', content: '<p>قولوا للناس حسنًا.</p>', youtube: YT, book: BK },
        { title: 'خدمة المجتمع', content: '<p>أحب الناس أنفعهم.</p>', youtube: YT, book: BK },
        { title: 'الهوية الإسلامية', content: '<p>الاعتزاز بالإسلام.</p>', youtube: YT, book: BK },
        { title: 'التخطيط للمستقبل', content: '<p>ازرع فسيلتك.</p>', youtube: YT, book: BK },
        { title: 'التوازن', content: '<p>أعط كل ذي حق حقه.</p>', youtube: YT, book: BK }
      ]
    }
  ];
}
