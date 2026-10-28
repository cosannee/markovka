/**
 * Приём заявок на новогодний адвент-календарь в Google Таблицу.
 *
 * Установка:
 * 1. Создайте Google Таблицу → Расширения → Apps Script.
 * 2. Вставьте этот код вместо содержимого Code.gs, сохраните.
 * 3. При желании укажите e-mail в NOTIFY_EMAIL — на него будет приходить письмо о каждой заявке.
 * 4. Развернуть → Новое развертывание → тип «Веб-приложение»:
 *      Выполнять как: «Я»; У кого есть доступ: «Все».
 *    Подтвердите доступы (Google предупредит о «непроверенном приложении» — это ваш скрипт,
 *    нажмите «Дополнительно» → «Перейти»).
 * 5. Скопируйте URL веб-приложения (оканчивается на /exec) и вставьте его в advent.html
 *    в строку  const ORDER_ENDPOINT = '';
 *
 * После любых правок кода: Развернуть → Управление развертываниями → ✎ → Версия «Новая» → Развернуть
 * (URL при этом не меняется).
 */

const SHEET_NAME = 'Заявки';
const NOTIFY_EMAIL = ''; // например 'you@gmail.com'; пусто — без писем

const COLUMNS = [
  ['date', 'Дата заявки'],
  ['name', 'Имя'],
  ['phone', 'Телефон'],
  ['contact', 'Telegram / мессенджер'],
  ['qty', 'Количество'],
  ['delivery', 'Способ получения'],
  ['address', 'Адрес СДЭК'],
  ['comment', 'Комментарий'],
  ['status', 'Статус'],
  ['page', 'Страница'],
];

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return ok_(); // антиспам-ловушка

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const values = Object.assign({}, p, { date: new Date(), status: 'Новая' });
    sheet.appendRow(COLUMNS.map(([key]) => clean_(values[key])));
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    const body = COLUMNS
      .filter(([key]) => key !== 'status' && p[key])
      .map(([key, title]) => `${title}: ${p[key]}`)
      .join('\n');
    MailApp.sendEmail(NOTIFY_EMAIL, `Новая заявка на адвент: ${p.name || ''}`, body);
  }
  return ok_();
}

function doGet() {
  return ContentService.createTextOutput('Приём заявок работает');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(([, title]) => title));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
  }
  return sheet;
}

// Защита от формул в ячейках (=, +, -, @ в начале строки) и ограничение длины
function clean_(v) {
  if (v instanceof Date) return v;
  const s = String(v == null ? '' : v).slice(0, 1000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
