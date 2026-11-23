/**
 * ===================================================================
 * ГОСТЕВОЙ ДОМ «У АРАМА» (г. Шахты, Ростовская обл.)
 * Google Apps Script для автоматического учёта бронирований
 * и блокировки занятых дат на сайте.
 * ===================================================================
 * 
 * ПОШАГОВАЯ ИНСТРУКЦИЯ (2 минуты):
 * 1. Откройте Google Таблицы (https://sheets.google.com) и создайте новую таблицу.
 *    Назовите её, например: «Бронирования — У Арама».
 * 2. В верхнем меню выберите: «Расширения» (Extensions) -> «Apps Script».
 * 3. В открывшемся редакторе удалите весь стандартный код и вставьте этот файл целиком.
 * 4. Нажмите синюю кнопку «Развернуть» (Deploy) вверху справа -> «Новое развертывание» (New deployment).
 * 5. Нажмите на значок шестерёнки слева от «Выберите тип» и выберите «Веб-приложение» (Web app).
 * 6. Заполните поля:
 *    - Описание: "Бронирования У Арама"
 *    - Выполнять от имени: "Меня" (Me)
 *    - У кого есть доступ: "Все" (Anyone) — ОБЯЗАТЕЛЬНО! Чтобы сайт мог отправлять заявки и проверять даты.
 * 7. Нажмите «Развернуть» (Deploy), разрешите доступ Google аккаунта и скопируйте
 *    полученный «URL веб-приложения» (заканчивается на /exec).
 * 8. Откройте сайт, внизу нажмите «Панель управления», вставьте этот URL и нажмите «Сохранить»!
 */

// Список квартир (соответствует карточкам на сайте)
var ROOMS_CONFIG = [
  { id: 'studio-comfort', name: 'Студия комфорт' },
  { id: 'room-1', name: '1-комнатная квартира' },
  { id: 'room-2-comfort', name: '2-х комнатная комфорт+' },
  { id: 'room-3-comfort', name: '3-х комнатная комфорт' },
  { id: 'studio-loft', name: 'Студия-лофт' },
  { id: 'studio-loft-premium', name: 'Студия-лофт премиум' }
];

var HEADERS = [
  'ID заявки',
  'Дата заявки',
  'Имя гостя',
  'Телефон',
  'Формат',
  'Заезд / Дата',
  'Выезд / Длительность',
  'Гостей',
  'Стоимость',
  'Статус',
  'Пожелания гостя'
];

/**
 * 1. ОБРАБОТКА POST-ЗАПРОСА (Когда гость отправляет заявку на сайте)
 * Добавляет новую строку в лист нужной квартиры со статусом «Ожидание».
 */
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data;
    
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    } else {
      data = {};
    }

    // Определяем квартиру
    var roomTitle = String(data.room || 'Студия комфорт').replace(/\s*\(.*?\)/g, '').trim();
    var sheet = getOrCreateRoomSheet(ss, roomTitle);

    var now = new Date();
    var dateOrder = Utilities.formatDate(now, 'GMT+3', 'dd.MM.yyyy HH:mm');
    var bookingId = 'UA-' + Utilities.formatDate(now, 'GMT+3', 'yyyyMMdd-HHmmss');

    var isHourly = data.tariff === 'hourly' || (data.tariffLine && String(data.tariffLine).indexOf('По часам') !== -1);
    var tariffName = isHourly ? 'По часам' : 'Посуточно';
    var checkinVal = isHourly ? (data.hourlyDate || data.checkin || '') : (data.checkin || '');
    var checkoutVal = isHourly ? (data.hourlyDuration || '3 часа') : (data.checkout || '');

    // Защита от ошибки #ERROR! в Google Таблицах:
    // Если номер телефона начинается со знака «+» (например, +7 904 444 44 91),
    // Google Sheets ошибочно парсит его как математическую формулу (=+7...),
    // что приводит к ошибке #ERROR!
    // Добавление апострофа «'» в начало сохраняет номер строго как чистый текст.
    // При этом сам апостроф в ячейке Google Таблицы не виден, отображается чистый номер +7...
    var phoneVal = String(data.phone || 'Не указан').trim();
    if (phoneVal.charAt(0) === '+') {
      phoneVal = "'" + phoneVal;
    }

    var newRow = [
      bookingId,
      dateOrder,
      data.name || 'Не указано',
      phoneVal,
      tariffName,
      checkinVal,
      checkoutVal,
      data.guests || '2',
      data.price ? (data.price + (String(data.price).indexOf('₽') === -1 ? ' ₽' : '')) : '',
      'Ожидание', // Начальный статус!
      data.comment || ''
    ];

    sheet.appendRow(newRow);

    // Добавляем выпадающий список для колонки «Статус» в новой строке
    var lastRow = sheet.getLastRow();
    var statusCell = sheet.getRange(lastRow, 10);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ожидание', 'Забронировано', 'Отклонено'], true)
      .setAllowInvalid(true)
      .build();
    statusCell.setDataValidation(rule);

    // Задаем текстовый формат для ячейки телефона (колонка 4)
    sheet.getRange(lastRow, 4).setNumberFormat('@');

    // Выравнивание и форматирование новой строки
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setVerticalAlignment('middle');
    sheet.getRange(lastRow, 1, 1, HEADERS.length).setFontFamily('Arial');
    sheet.setRowHeight(lastRow, 28);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      id: bookingId,
      room: roomTitle,
      message: 'Бронирование успешно добавлено со статусом Ожидание'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 2. ОБРАБОТКА GET-ЗАПРОСА (Сайт запрашивает занятые даты для календарей)
 * Сканирует листы всех квартир. Если статус строки === «Забронировано»,
 * возвращает эти даты сайту для закрытия.
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheetsIfMissing(ss);

    var bookedResult = {};

    ROOMS_CONFIG.forEach(function(room) {
      bookedResult[room.id] = {
        title: room.name,
        daily: [],
        hourly: []
      };

      var sheet = getSheetByNameCaseInsensitive(ss, room.name);
      if (!sheet) return;

      // Автоматически исправляем формульные ошибки в колонке телефона, если они были ранее
      fixPhoneErrors(sheet);

      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return; // Только шапка

      var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

      values.forEach(function(row) {
        var format = String(row[4] || '').trim(); // Колонка E: Формат
        var checkinRaw = row[5];                  // Колонка F: Заезд / Дата
        var checkoutRaw = row[6];                 // Колонка G: Выезд / Длительность
        var status = String(row[9] || '').trim().toLowerCase(); // Колонка J: Статус

        // Блокируем даты только когда статус подтверждён Арамом: «Забронировано»
        // Учитываем также возможные опечатки (например, «Забронироано»)
        var isBooked = (
          status === 'забронировано' ||
          status === 'забронироано' ||
          status === 'подтверждено' ||
          status.indexOf('забронирован') !== -1
        );

        if (isBooked) {
          if (format.toLowerCase().indexOf('час') !== -1) {
            // Почасовая бронь
            var hDate = formatDateISO(checkinRaw);
            if (hDate && bookedResult[room.id].hourly.indexOf(hDate) === -1) {
              bookedResult[room.id].hourly.push(hDate);
            }
          } else {
            // Посуточная бронь: получаем все даты проживания
            var dIn = formatDateISO(checkinRaw);
            var dOut = formatDateISO(checkoutRaw);
            if (dIn) {
              var stayDates = getDaysRange(dIn, dOut);
              stayDates.forEach(function(dStr) {
                if (bookedResult[room.id].daily.indexOf(dStr) === -1) {
                  bookedResult[room.id].daily.push(dStr);
                }
              });
            }
          }
        }
      });
    });

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      updatedAt: Utilities.formatDate(new Date(), 'GMT+3', 'yyyy-MM-dd HH:mm:ss'),
      booked: bookedResult
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Поиск листа без учёта регистра и лишних пробелов
 */
function getSheetByNameCaseInsensitive(ss, name) {
  var target = String(name).trim().toLowerCase();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === target) {
      return sheets[i];
    }
  }
  return null;
}

/**
 * Создать или получить лист номера
 */
function getOrCreateRoomSheet(ss, roomName) {
  var sheet = getSheetByNameCaseInsensitive(ss, roomName);
  if (!sheet) {
    sheet = ss.insertSheet(roomName);
    formatSheetHeader(sheet);
  }
  return sheet;
}

/**
 * Оформление шапки таблицы, ширины колонок и условного форматирования статусов
 */
function formatSheetHeader(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setBackground('#1055D6')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);

  // Ширина колонок
  sheet.setColumnWidth(1, 150); // ID
  sheet.setColumnWidth(2, 140); // Дата заявки
  sheet.setColumnWidth(3, 140); // Имя
  sheet.setColumnWidth(4, 140); // Телефон
  sheet.setColumnWidth(5, 110); // Формат
  sheet.setColumnWidth(6, 120); // Заезд
  sheet.setColumnWidth(7, 130); // Выезд
  sheet.setColumnWidth(8, 80);  // Гостей
  sheet.setColumnWidth(9, 110); // Стоимость
  sheet.setColumnWidth(10, 150); // Статус
  sheet.setColumnWidth(11, 260); // Пожелания

  // Колонка D (Телефон) - принудительно устанавливаем формат Plain Text (@),
  // чтобы при любом редактировании номера с плюсом не возникало #ERROR!
  try {
    sheet.getRange(2, 4, Math.max(100, sheet.getMaxRows() - 1), 1).setNumberFormat('@');
  } catch(e) {}

  // Условное форматирование для колонки J (Статус):
  // «Ожидание» -> янтарный, «Забронировано» -> зелёный, «Отклонено» -> светло-красный
  try {
    var statusRange = sheet.getRange("J2:J1000");
    var rulePending = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Ожидание")
      .setBackground("#FEF3C7")
      .setFontColor("#92400E")
      .setRanges([statusRange])
      .build();

    var ruleBooked = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Забронировано")
      .setBackground("#D1FAE5")
      .setFontColor("#065F46")
      .setRanges([statusRange])
      .build();

    var ruleRejected = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Отклонено")
      .setBackground("#FEE2E2")
      .setFontColor("#991B1B")
      .setRanges([statusRange])
      .build();

    sheet.setConditionalFormatRules([rulePending, ruleBooked, ruleRejected]);
  } catch (e) {
    // В старых версиях API правила пропускаются без падения
  }
}

/**
 * Автоматически создает листы для всех 6 комнат
 */
function setupSheetsIfMissing(ss) {
  ROOMS_CONFIG.forEach(function(room) {
    if (!getSheetByNameCaseInsensitive(ss, room.name)) {
      var sh = ss.insertSheet(room.name);
      formatSheetHeader(sh);
    }
  });

  // Удаляем стандартный пустой "Лист 1" / "Sheet1"
  var defSheet = ss.getSheetByName('Лист 1') || ss.getSheetByName('Sheet1');
  if (defSheet && ss.getSheets().length > 1 && defSheet.getLastRow() === 0) {
    try { ss.deleteSheet(defSheet); } catch(e) {}
  }
}

/**
 * Преобразует любую дату таблицы в строку YYYY-MM-DD
 */
function formatDateISO(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'GMT+3', 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  var match = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
  if (match) {
    return match[3] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[1]).slice(-2);
  }
  return null;
}

/**
 * Возвращает массив всех дней между датой заезда и выезда (включительно все ночи)
 */
function getDaysRange(inStr, outStr) {
  var list = [];
  if (!inStr) return list;
  var d1 = new Date(inStr + 'T00:00:00');
  var d2 = outStr ? new Date(outStr + 'T00:00:00') : new Date(d1.getTime() + 86400000);
  if (isNaN(d1.getTime())) return list;
  if (isNaN(d2.getTime()) || d2 <= d1) d2 = new Date(d1.getTime() + 86400000);

  var curr = new Date(d1);
  while (curr < d2) {
    list.push(Utilities.formatDate(curr, 'GMT+3', 'yyyy-MM-dd'));
    curr.setDate(curr.getDate() + 1);
  }
  return list;
}

/**
 * Автоматическое исправление ячеек с ошибкой #ERROR! в колонке телефона (Колонка D)
 * Если в ячейке находится формула вида =+7904... или значение #ERROR!,
 * функция конвертирует её в чистый текст '+7... и форматирует как Plain text.
 */
function fixPhoneErrors(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var range = sheet.getRange(2, 4, lastRow - 1, 1);
    var formulas = range.getFormulas();
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
      var formula = String(formulas[i][0] || '').trim();
      var val = String(values[i][0] || '').trim();
      if (formula && (formula.indexOf('=+') === 0 || formula.indexOf('+') === 0)) {
        var cleanPhone = formula.replace(/^=/, '');
        range.getCell(i + 1, 1).setNumberFormat('@').setValue("'" + cleanPhone);
      } else if (val === '#ERROR!' && formula) {
        var cleanPhone = formula.replace(/^=/, '');
        range.getCell(i + 1, 1).setNumberFormat('@').setValue("'" + cleanPhone);
      }
    }
  } catch (err) {
    // Безопасный перехват без остановки выполнения
  }
}
