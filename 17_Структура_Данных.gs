/**
 * СТРУКТУРА ДАННЫХ И СОПОСТАВЛЕНИЕ ПОЛЕЙ
 * Универсальная система объединения данных из всех источников
 * @fileoverview Модуль для создания и настройки сводного листа "РАБОЧИЙ АМО"
 */

/**
 * Создает структуру сводного листа "РАБОЧИЙ АМО" 
 * Объединяет данные из всех источников в единую структуру
 */
function createWorkingAmoStructure() {
  try {
    logInfo_('DATA_STRUCTURE', 'Создание структуры сводного листа РАБОЧИЙ АМО');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
    
    if (!workingSheet) {
      workingSheet = spreadsheet.insertSheet('РАБОЧИЙ АМО');
    }
    
    // Очищаем лист
    workingSheet.clear();
    
    // Применяем красивое оформление
    applyWorkingAmoStyle_(workingSheet);
    
    // Создаем заголовки всех полей
    const headers = getUnifiedHeaders_();
    setWorkingAmoHeaders_(workingSheet, headers);
    
    // Настраиваем форматирование столбцов
    formatWorkingAmoColumns_(workingSheet, headers);
    
    logInfo_('DATA_STRUCTURE', 'Структура РАБОЧИЙ АМО создана успешно');
    
    return workingSheet;
    
  } catch (error) {
    logError_('DATA_STRUCTURE', 'Ошибка создания структуры РАБОЧИЙ АМО', error);
    throw error;
  }
}

/**
 * Возвращает единые заголовки для сводного листа
 * Включает все поля из всех источников данных
 * @returns {Array} Массив заголовков
 * @private
 */
function getUnifiedHeaders_() {
  return [
    // ОСНОВНЫЕ ПОЛЯ СДЕЛКИ (из AmoCRM)
    'ID',                          // A - Уникальный номер сделки
    'Название',                    // B - Название сделки
    'Ответственный',              // C - Менеджер, ведущий сделку
    'Контакт.ФИО',               // D - Имя клиента
    'Статус',                     // E - Текущий статус сделки
    'Бюджет',                     // F - Сумма сделки
    'Дата создания',             // G - Когда создана сделка
    
    // КОНТАКТНЫЕ ДАННЫЕ
    'Контакт.Телефон',           // H - Телефон клиента
    'Email',                      // I - Email клиента
    'Номер линии MANGO OFFICE',  // J - Телефонная линия
    
    // СТАТУСЫ И РЕЗУЛЬТАТЫ
    'Теги',                       // K - Метки сделки
    'Дата закрытия',             // L - Когда закрыта сделка
    'Причина отказа',            // M - Почему отказались
    'Тип лида',                  // N - Целевой/нецелевой
    
    // ДАННЫЕ БРОНИРОВАНИЯ
    'Дата брони',                // O - На какую дату бронь
    'Время прихода',             // P - Во сколько придут
    'Кол-во гостей',             // Q - Количество гостей
    'Бар',                       // R - Какой филиал
    
    // ИСТОЧНИКИ ТРАФИКА
    'Источник',                   // S - Основной канал
    'R.Источник сделки',         // T - Детальный источник
    'R.Источник ТЕЛ сделки',     // U - Источник звонка
    'ПО',                        // V - Система привлечения
    
    // UTM МЕТКИ
    'UTM_SOURCE',                // W - utm_source
    'UTM_MEDIUM',                // X - utm_medium  
    'UTM_CAMPAIGN',              // Y - utm_campaign
    'UTM_TERM',                  // Z - utm_term
    'UTM_CONTENT',               // AA - utm_content
    'utm_referrer',              // AB - Источник перехода
    
    // АНАЛИТИЧЕСКИЕ ID
    'YM_CLIENT_ID',              // AC - ID Яндекс.Метрики
    'GA_CLIENT_ID',              // AD - ID Google Analytics
    '_ym_uid',                   // AE - Уникальный ID Метрики
    
    // ФОРМЫ И ЛЕНДИНГИ
    'FORMNAME',                  // AF - Название формы
    'FORMID',                    // AG - ID формы
    'BUTTON_TEXT',               // AH - Текст кнопки
    'REFERER',                   // AI - Страница перехода
    
    // ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ
    'DATE',                      // AJ - Дата мероприятия
    'TIME',                      // AK - Время мероприятия
    'R.Тег города',              // AL - Город
    'R.Статусы гостей',          // AM - Пришли/не пришли
    'Сарафан гости',             // AN - Сарафанное радио
    
    // КОММЕНТАРИИ
    'Комментарий МОБ',           // AO - Комментарий менеджера
    'Примечания',                // AP - Дополнительные записи
    'Комментарий',               // AQ - Общий комментарий
    
    // ДАННЫЕ ГОСТЕЙ (из RP)
    'Кол-во визитов',            // AR - Сколько раз был
    'Общая сумма',               // AS - Потратил всего
    'Первый визит',              // AT - Дата первого визита
    'Последний визит',           // AU - Дата последнего
    'Счёт, ₽'                    // AV - Сумма счета
  ];
}

/**
 * Устанавливает заголовки на листе РАБОЧИЙ АМО
 * @param {Sheet} sheet - Рабочий лист
 * @param {Array} headers - Массив заголовков
 * @private
 */
function setWorkingAmoHeaders_(sheet, headers) {
  // Устанавливаем заголовки в строку 3 (после титула и даты)
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  
  // Применяем стиль заголовков
  const headerRange = sheet.getRange(3, 1, 1, headers.length);
  headerRange
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  
  // Устанавливаем высоту строки заголовков
  sheet.setRowHeight(3, 40);
  
  // Замораживаем строки заголовков
  sheet.setFrozenRows(3);
}

/**
 * Применяет стиль к листу РАБОЧИЙ АМО
 * @param {Sheet} sheet - Рабочий лист  
 * @private
 */
function applyWorkingAmoStyle_(sheet) {
  // ЗАГОЛОВОК (строка 1)
  sheet.getRange('A1:AV1').merge();
  sheet.getRange('A1')
    .setValue('📊 РАБОЧИЙ АМО - СВОДНЫЕ ДАННЫЕ')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 50);
  
  // ВРЕМЯ ОБНОВЛЕНИЯ (строка 2)
  sheet.getRange('A2:AV2').merge();
  sheet.getRange('A2')
    .setValue(`⏰ Последнее обновление: ${getCurrentDateMoscow_().toLocaleString()}`)
    .setBackground('#f8f9fa')
    .setFontSize(11)
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(2, 25);
  
  // Фон всего листа
  sheet.getRange(1, 1, 1000, 50).setBackground('#ffffff');
}

/**
 * Настраивает форматирование столбцов
 * @param {Sheet} sheet - Рабочий лист
 * @param {Array} headers - Заголовки
 * @private
 */
function formatWorkingAmoColumns_(sheet, headers) {
  // Настраиваем ширину столбцов по типу данных
  const columnWidths = {
    // Основные поля
    'ID': 80,
    'Название': 150,
    'Ответственный': 120,
    'Контакт.ФИО': 130,
    'Статус': 140,
    'Бюджет': 100,
    'Дата создания': 110,
    
    // Контакты
    'Контакт.Телефон': 120,
    'Email': 180,
    'Номер линии MANGO OFFICE': 100,
    
    // Статусы
    'Теги': 100,
    'Дата закрытия': 110,
    'Причина отказа': 200,
    'Тип лида': 100,
    
    // Бронирование
    'Дата брони': 100,
    'Время прихода': 90,
    'Кол-во гостей': 80,
    'Бар': 80,
    
    // Источники
    'Источник': 120,
    'R.Источник сделки': 140,
    'R.Источник ТЕЛ сделки': 140,
    'ПО': 80,
    
    // UTM
    'UTM_SOURCE': 120,
    'UTM_MEDIUM': 100,
    'UTM_CAMPAIGN': 150,
    'UTM_TERM': 120,
    'UTM_CONTENT': 120,
    'utm_referrer': 150,
    
    // ID аналитики
    'YM_CLIENT_ID': 150,
    'GA_CLIENT_ID': 150,
    '_ym_uid': 150,
    
    // Формы
    'FORMNAME': 120,
    'FORMID': 80,
    'BUTTON_TEXT': 120,
    'REFERER': 200,
    
    // Комментарии (широкие столбцы)
    'Комментарий МОБ': 250,
    'Примечания': 200,
    'Комментарий': 200
  };
  
  // Устанавливаем ширину столбцов
  headers.forEach((header, index) => {
    const width = columnWidths[header] || 90; // По умолчанию 90px
    sheet.setColumnWidth(index + 1, width);
  });
  
  // Форматирование числовых столбцов
  const numberColumns = ['Бюджет', 'Кол-во гостей', 'Кол-во визитов', 'Общая сумма', 'Счёт, ₽'];
  numberColumns.forEach(columnName => {
    const columnIndex = headers.indexOf(columnName);
    if (columnIndex >= 0) {
      sheet.getRange(4, columnIndex + 1, 1000, 1)
        .setNumberFormat('#,##0.00 ₽');
    }
  });
  
  // Форматирование дат
  const dateColumns = ['Дата создания', 'Дата закрытия', 'Дата брони', 'Первый визит', 'Последний визит'];
  dateColumns.forEach(columnName => {
    const columnIndex = headers.indexOf(columnName);
    if (columnIndex >= 0) {
      sheet.getRange(4, columnIndex + 1, 1000, 1)
        .setNumberFormat('dd.mm.yyyy');
    }
  });
}

/**
 * КАРТА СОПОСТАВЛЕНИЯ ПОЛЕЙ ИЗ РАЗНЫХ ИСТОЧНИКОВ
 * Определяет, какие поля из какого источника куда попадают
 */
const FIELD_MAPPING = {
  // 1. AmoCRM - Амо Выгрузка
  'AMO_EXPORT': {
    'ID': 'ID',
    'Название': 'Название', 
    'Ответственный': 'Ответственный',
    'Статус': 'Статус',
    'Бюджет': 'Бюджет',
    'Дата создания': 'Дата создания',
    'Теги': 'Теги',
    'Дата закрытия': 'Дата закрытия',
    'YM_CLIENT_ID': 'YM_CLIENT_ID',
    'GA_CLIENT_ID': 'GA_CLIENT_ID',
    'BUTTON_TEXT': 'BUTTON_TEXT',
    'DATE': 'DATE',
    'TIME': 'TIME',
    'R.Источник сделки': 'R.Источник сделки',
    'R.Тег города': 'R.Тег города',
    'ПО': 'ПО',
    'Бар': 'Бар',
    'Дата брони': 'Дата брони',
    'Кол-во гостей': 'Кол-во гостей',
    'Время прихода': 'Время прихода',
    'Комментарий МОБ': 'Комментарий МОБ',
    'Источник': 'Источник',
    'Тип лида': 'Тип лида',
    'Причина отказа': 'Причина отказа',
    'R.Статусы гостей': 'R.Статусы гостей',
    'Сарафан гости': 'Сарафан гости',
    'UTM_MEDIUM': 'UTM_MEDIUM',
    'FORMNAME': 'FORMNAME',
    'REFERER': 'REFERER',
    'FORMID': 'FORMID',
    'Номер линии MANGO OFFICE': 'Номер линии MANGO OFFICE',
    'UTM_SOURCE': 'UTM_SOURCE',
    'UTM_TERM': 'UTM_TERM',
    'UTM_CAMPAIGN': 'UTM_CAMPAIGN',
    'UTM_CONTENT': 'UTM_CONTENT',
    'utm_referrer': 'utm_referrer',
    '_ym_uid': '_ym_uid',
    'Примечания': 'Примечания',
    'Контакт.ФИО': 'Контакт.ФИО',
    'Контакт.Телефон': 'Контакт.Телефон'
  },
  
  // 2. AmoCRM - Выгрузка Амо Полная
  'AMO_FULL_EXPORT': {
    'ID': 'ID',
    'Название': 'Название',
    'Ответственный': 'Ответственный',
    'Контакт.ФИО': 'Контакт.ФИО',
    'Статус': 'Статус',
    'Бюджет': 'Бюджет',
    'Дата создания': 'Дата создания',
    'Контакт.Телефон': 'Контакт.Телефон',
    'Причина отказа': 'Причина отказа'
    // ... остальные поля аналогично
  },
  
  // 3. Reserves RP 
  'RESERVES_RP': {
    'ID': 'ID',
    '№ заявки': 'Название', // Номер заявки → Название
    'Имя': 'Контакт.ФИО',
    'Телефон': 'Контакт.Телефон',
    'Email': 'Email',
    'Дата/время': 'Дата создания',
    'Статус': 'Статус',
    'Комментарий': 'Комментарий',
    'Счёт, ₽': 'Бюджет',
    'Гостей': 'Кол-во гостей',
    'Источник': 'Источник'
  },
  
  // 4. Guests RP
  'GUESTS_RP': {
    'Имя': 'Контакт.ФИО',
    'Телефон': 'Контакт.Телефон', 
    'Email': 'Email',
    'Кол-во визитов': 'Кол-во визитов',
    'Общая сумма': 'Общая сумма',
    'Первый визит': 'Первый визит',
    'Последний визит': 'Последний визит'
  },
  
  // 5. Заявки с Сайта
  'WEBSITE_FORMS': {
    'Name': 'Контакт.ФИО',
    'Phone': 'Контакт.Телефон',
    'Email': 'Email',
    'Date': 'Дата брони',
    'Time': 'Время прихода',
    'Quantity': 'Кол-во гостей',
    'utm_source': 'UTM_SOURCE',
    'utm_medium': 'UTM_MEDIUM',
    'utm_campaign': 'UTM_CAMPAIGN',
    'utm_term': 'UTM_TERM',
    'utm_content': 'UTM_CONTENT',
    'ym_client_id': 'YM_CLIENT_ID',
    'ga_client_id': 'GA_CLIENT_ID',
    'button_text': 'BUTTON_TEXT',
    'referer': 'REFERER',
    'formid': 'FORMID',
    'Form name': 'FORMNAME'
  },
  
  // 6. Коллтрекинг
  'CALL_TRACKING': {
    'Контакт.Номер линии MANGO OFFICE': 'Номер линии MANGO OFFICE',
    'R.Источник ТЕЛ сделки': 'R.Источник ТЕЛ сделки',
    'Название Канала': 'Источник'
  }
};

/**
 * Объединяет данные из указанного источника в РАБОЧИЙ АМО
 * @param {string} sourceSheetName - Имя листа-источника
 * @param {string} sourceType - Тип источника (ключ из FIELD_MAPPING)
 */
function mergeDataToWorkingAmo(sourceSheetName, sourceType) {
  try {
    logInfo_('DATA_MERGE', `Объединение данных из ${sourceSheetName} (тип: ${sourceType})`);
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
    const workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
    
    if (!sourceSheet) {
      logWarning_('DATA_MERGE', `Лист "${sourceSheetName}" не найден`);
      return;
    }
    
    if (!workingSheet) {
      createWorkingAmoStructure();
      return mergeDataToWorkingAmo(sourceSheetName, sourceType);
    }
    
    // Получаем данные из источника
    const sourceData = getSheetData_(sourceSheet);
    if (sourceData.length <= 1) {
      logWarning_('DATA_MERGE', `Нет данных в листе "${sourceSheetName}"`);
      return;
    }
    
    const sourceHeaders = sourceData[0];
    const sourceRows = sourceData.slice(1);
    
    // Получаем заголовки рабочего листа
    const workingHeaders = getSheetData_(workingSheet)[0];
    const fieldMapping = FIELD_MAPPING[sourceType];
    
    if (!fieldMapping) {
      logError_('DATA_MERGE', `Неизвестный тип источника: ${sourceType}`);
      return;
    }
    
    // Преобразуем и добавляем данные
    const mappedRows = sourceRows.map(sourceRow => {
      const mappedRow = new Array(workingHeaders.length).fill('');
      
      // Сопоставляем поля
      Object.entries(fieldMapping).forEach(([sourceField, targetField]) => {
        const sourceIndex = sourceHeaders.indexOf(sourceField);
        const targetIndex = workingHeaders.indexOf(targetField);
        
        if (sourceIndex >= 0 && targetIndex >= 0 && sourceRow[sourceIndex] !== undefined) {
          mappedRow[targetIndex] = sourceRow[sourceIndex];
        }
      });
      
      return mappedRow;
    });
    
    // Добавляем данные в рабочий лист
    if (mappedRows.length > 0) {
      const startRow = workingSheet.getLastRow() + 1;
      workingSheet.getRange(startRow, 1, mappedRows.length, workingHeaders.length)
        .setValues(mappedRows);
      
      logInfo_('DATA_MERGE', `Добавлено ${mappedRows.length} записей из ${sourceSheetName}`);
    }
    
  } catch (error) {
    logError_('DATA_MERGE', `Ошибка объединения данных из ${sourceSheetName}`, error);
  }
}

/**
 * Полное обновление РАБОЧИЙ АМО из всех источников
 */
function updateWorkingAmoFromAllSources() {
  try {
    logInfo_('DATA_MERGE', 'Начало полного обновления РАБОЧИЙ АМО из всех источников');
    
    // Создаем/очищаем рабочий лист
    const workingSheet = createWorkingAmoStructure();
    
    // Источники данных для объединения
    const dataSources = [
      { sheetName: 'Амо Выгрузка', type: 'AMO_EXPORT' },
      { sheetName: 'Выгрузка Амо Полная', type: 'AMO_FULL_EXPORT' },
      { sheetName: 'Reserves RP', type: 'RESERVES_RP' },
      { sheetName: 'Guests RP', type: 'GUESTS_RP' },
      { sheetName: 'Заявки с Сайта', type: 'WEBSITE_FORMS' },
      { sheetName: 'КоллТрекинг', type: 'CALL_TRACKING' }
    ];
    
    let totalMerged = 0;
    
    // Объединяем данные из каждого источника
    dataSources.forEach(source => {
      try {
        const beforeCount = workingSheet.getLastRow() - 3; // Минус заголовки
        mergeDataToWorkingAmo(source.sheetName, source.type);
        const afterCount = workingSheet.getLastRow() - 3;
        const merged = Math.max(0, afterCount - beforeCount);
        totalMerged += merged;
        
        if (merged > 0) {
          logInfo_('DATA_MERGE', `✅ ${source.sheetName}: добавлено ${merged} записей`);
        } else {
          logInfo_('DATA_MERGE', `⚠️ ${source.sheetName}: нет новых данных`);
        }
        
      } catch (sourceError) {
        logWarning_('DATA_MERGE', `Ошибка обработки ${source.sheetName}`, sourceError);
      }
    });
    
    // Обновляем время последнего обновления
    workingSheet.getRange('A2')
      .setValue(`⏰ Последнее обновление: ${getCurrentDateMoscow_().toLocaleString()} | Всего записей: ${totalMerged}`);
    
    logInfo_('DATA_MERGE', `Полное обновление завершено. Всего объединено: ${totalMerged} записей`);
    
    // Применяем автофильтры
    const dataRange = workingSheet.getRange(3, 1, workingSheet.getLastRow() - 2, getUnifiedHeaders_().length);
    workingSheet.setAutoFilter(dataRange);
    
    return totalMerged;
    
  } catch (error) {
    logError_('DATA_MERGE', 'Ошибка полного обновления РАБОЧИЙ АМО', error);
    throw error;
  }
}

/**
 * Диагностика структуры данных РАБОЧИЙ АМО
 */
function diagnoseWorkingAmoStructure() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
    
    if (!workingSheet) {
      console.log('❌ Лист "РАБОЧИЙ АМО" не существует');
      console.log('🔧 Используйте createWorkingAmoStructure() для создания');
      return;
    }
    
    const data = getSheetData_(workingSheet);
    const headers = data[0];
    const rows = data.slice(1);
    
    console.log('📊 ДИАГНОСТИКА СТРУКТУРЫ РАБОЧИЙ АМО:');
    console.log(`📋 Всего столбцов: ${headers.length}`);
    console.log(`📋 Всего записей: ${rows.length}`);
    
    // Анализ заполненности ключевых полей
    const keyFields = ['ID', 'Название', 'Статус', 'Контакт.ФИО', 'Контакт.Телефон', 'Источник', 'Причина отказа'];
    
    console.log('\n🔍 ЗАПОЛНЕННОСТЬ КЛЮЧЕВЫХ ПОЛЕЙ:');
    keyFields.forEach(field => {
      const columnIndex = headers.indexOf(field);
      if (columnIndex >= 0) {
        let filled = 0;
        rows.forEach(row => {
          if (row[columnIndex] && String(row[columnIndex]).trim() !== '') {
            filled++;
          }
        });
        const percentage = rows.length > 0 ? ((filled / rows.length) * 100).toFixed(1) : '0';
        console.log(`• ${field}: ${filled}/${rows.length} (${percentage}%)`);
      } else {
        console.log(`• ${field}: ❌ столбец не найден`);
      }
    });
    
    // Анализ источников данных
    const sourceIndex = headers.indexOf('Источник');
    if (sourceIndex >= 0) {
      const sourceStats = {};
      rows.forEach(row => {
        const source = String(row[sourceIndex] || 'Не указан').trim();
        sourceStats[source] = (sourceStats[source] || 0) + 1;
      });
      
      console.log('\n📈 РАСПРЕДЕЛЕНИЕ ПО ИСТОЧНИКАМ:');
      Object.entries(sourceStats)
        .sort(([,a], [,b]) => b - a)
        .slice(0, 10)
        .forEach(([source, count]) => {
          console.log(`• ${source}: ${count}`);
        });
    }
    
    console.log('\n✅ Диагностика завершена');
    
  } catch (error) {
    console.error('❌ Ошибка диагностики:', error);
  }
}
