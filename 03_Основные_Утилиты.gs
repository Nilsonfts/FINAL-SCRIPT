/**
 * МОДУЛЬ: ОСНОВНЫЕ УТИЛИТЫ
 * Базовые функции для работы с данными
 */

/**
 * УТИЛИТЫ ФОРМАТИРОВАНИЯ И ОБРАБОТКИ ДАННЫХ
 */

/**
 * Форматирование чисел
 */
function formatNumber(value) {
  if (!value && value !== 0) return 0;
  
  // Преобразуем в число
  let num;
  if (typeof value === 'number') {
    num = value;
  } else if (typeof value === 'string') {
    // Удаляем пробелы, запятые как разделители тысяч и заменяем запятую на точку
    const cleanValue = value.toString()
      .replace(/\s/g, '')
      .replace(/,(\d{3})/g, '$1') // убираем запятые как разделители тысяч
      .replace(',', '.'); // заменяем запятую на точку как десятичный разделитель
    
    num = parseFloat(cleanValue);
  } else {
    num = parseFloat(value);
  }
  
  return isNaN(num) ? 0 : num;
}

/**
 * Форматирование валюты
 */
function formatCurrency(value, currency = 'RUB') {
  const num = formatNumber(value);
  
  if (num === 0) return '0 ₽';
  
  // Форматируем число с пробелами как разделителями тысяч
  const formatted = num.toLocaleString('ru-RU', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });
  
  // Добавляем символ валюты
  switch (currency.toUpperCase()) {
    case 'RUB':
    case 'RUR':
      return formatted + ' ₽';
    case 'USD':
      return '$' + formatted;
    case 'EUR':
      return '€' + formatted;
    default:
      return formatted + ' ' + currency;
  }
}

/**
 * Форматирование процентов
 */
function formatPercent(value, decimals = 1) {
  const num = formatNumber(value);
  
  if (num === 0) return '0%';
  
  // Проверяем и исправляем параметр decimals
  let validDecimals = decimals;
  if (typeof decimals !== 'number' || isNaN(decimals) || decimals < 0 || decimals > 100) {
    validDecimals = 1; // Значение по умолчанию
  }
  
  // Округляем decimals до целого числа
  validDecimals = Math.floor(validDecimals);
  
  return (num * 100).toFixed(validDecimals) + '%';
}

/**
 * Безопасное форматирование процентов (с дополнительной защитой)
 */
function safeFormatPercent(value, decimals = 1) {
  try {
    return formatPercent(value, decimals);
  } catch (error) {
    console.warn('Ошибка форматирования процента:', error, 'value:', value, 'decimals:', decimals);
    const num = formatNumber(value);
    return (num * 100).toFixed(1) + '%';
  }
}

/**
 * Вычисление процента с защитой от деления на ноль
 */
function calculatePercent(part, total, decimals = 1) {
  if (!total || total === 0) return 0;
  
  const percent = part / total;
  return formatPercent(percent, decimals);
}

/**
 * Безопасное деление с возвратом 0 при делении на ноль
 */
function safeDivide(dividend, divisor, defaultValue = 0) {
  if (!divisor || divisor === 0) return defaultValue;
  return dividend / divisor;
}

/**
 * Нормализация телефонов
 */
function normalizePhone(phone) {
  if (!phone) return '';
  
  // Удаляем все кроме цифр
  const cleaned = phone.toString().replace(/\D/g, '');
  
  // Если номер начинается с 8, заменяем на 7
  if (cleaned.startsWith('8') && cleaned.length === 11) {
    return '7' + cleaned.substring(1);
  }
  
  // Если номер начинается с +7, убираем +
  if (cleaned.startsWith('7') && cleaned.length === 11) {
    return cleaned;
  }
  
  return cleaned;
}

/**
 * Проверка успешности статуса
 */
function isSuccessStatus(status) {
  if (!status) return false;
  
  const successStatuses = [
    'успешно реализовано',
    'закрыто и реализовано', 
    'продано',
    'оплачено',
    'договор подписан',
    'сделка заключена',
    'won',
    'closed won',
    'success'
  ];
  
  const lowerStatus = status.toString().toLowerCase();
  return successStatuses.some(successStatus => 
    lowerStatus.includes(successStatus)
  );
}

/**
 * Проверка статуса отказа
 */
function isRefusalStatus(status) {
  if (!status) return false;
  
  const refusalStatuses = [
    'закрыто и не реализовано',
    'отказ',
    'не дозвонились',
    'дубль',
    'брак',
    'спам',
    'lost',
    'closed lost',
    'rejected'
  ];
  
  const lowerStatus = status.toString().toLowerCase();
  return refusalStatuses.some(refusalStatus => 
    lowerStatus.includes(refusalStatus)
  );
}

/**
 * Парсинг UTM источника
 */
function parseUtmSource(row) {
  if (!row || !CONFIG.WORKING_AMO_COLUMNS) return 'unknown';
  
  const utmSource = row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || '';
  const utmMedium = row[CONFIG.WORKING_AMO_COLUMNS.UTM_MEDIUM] || '';
  const dealSource = row[CONFIG.WORKING_AMO_COLUMNS.DEAL_SOURCE] || '';
  
  // Комбинируем источники для определения канала
  const combined = (utmSource + ' ' + utmMedium + ' ' + dealSource).toLowerCase();
  
  if (combined.includes('yandex') && combined.includes('cpc')) return 'yandex_direct';
  if (combined.includes('google') && combined.includes('cpc')) return 'google_ads';
  if (combined.includes('facebook')) return 'facebook_ads';
  if (combined.includes('vk')) return 'vk_ads';
  if (combined.includes('yandex') && !combined.includes('cpc')) return 'yandex_organic';
  if (combined.includes('google') && !combined.includes('cpc')) return 'google_organic';
  if (combined.includes('direct')) return 'direct';
  if (combined.includes('referral')) return 'referral';
  if (combined.includes('site')) return 'site';
  
  return utmSource || dealSource || 'unknown';
}

/**
 * Определение типа канала из строки данных
 */
function getChannelType(row) {
  if (!row || !CONFIG.WORKING_AMO_COLUMNS) return 'Неизвестный';
  
  // Получаем данные из строки
  const utmSource = (row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || '').toString().toLowerCase();
  const utmMedium = (row[CONFIG.WORKING_AMO_COLUMNS.UTM_MEDIUM] || '').toString().toLowerCase();
  const dealSource = (row[CONFIG.WORKING_AMO_COLUMNS.DEAL_SOURCE] || '').toString().toLowerCase();
  const utmCampaign = (row[CONFIG.WORKING_AMO_COLUMNS.UTM_CAMPAIGN] || '').toString().toLowerCase();
  
  // Комбинируем все источники для анализа
  const combined = `${utmSource} ${utmMedium} ${dealSource} ${utmCampaign}`.toLowerCase();
  
  // Определяем тип канала по приоритету
  
  // Яндекс.Директ
  if (combined.includes('yandex') && (combined.includes('cpc') || combined.includes('direct'))) {
    return 'Яндекс.Директ';
  }
  
  // Google Ads
  if (combined.includes('google') && combined.includes('cpc')) {
    return 'Google Ads';
  }
  
  // Facebook/Instagram реклама
  if (combined.includes('facebook') || combined.includes('instagram') || combined.includes('fb')) {
    return 'Facebook/Instagram';
  }
  
  // VKontakte реклама
  if (combined.includes('vk') || combined.includes('vkontakte')) {
    return 'ВКонтакте';
  }
  
  // Telegram реклама
  if (combined.includes('telegram') || combined.includes('tg')) {
    return 'Telegram';
  }
  
  // Email маркетинг
  if (combined.includes('email') || combined.includes('newsletter') || utmMedium.includes('email')) {
    return 'Email маркетинг';
  }
  
  // Органический поиск
  if (combined.includes('organic') || 
      (combined.includes('yandex') && !combined.includes('cpc')) ||
      (combined.includes('google') && !combined.includes('cpc'))) {
    return 'Органический поиск';
  }
  
  // Прямые переходы
  if (combined.includes('direct') || combined.includes('(direct)') || 
      utmSource === 'direct' || dealSource.includes('direct')) {
    return 'Прямые переходы';
  }
  
  // Рефералы
  if (combined.includes('referral') || combined.includes('ref')) {
    return 'Рефералы';
  }
  
  // Сайт/Лендинг
  if (combined.includes('site') || combined.includes('landing') || combined.includes('website')) {
    return 'Сайт';
  }
  
  // Колл-трекинг
  if (combined.includes('call') || combined.includes('phone') || combined.includes('звонок')) {
    return 'Телефонные звонки';
  }
  
  // Если есть источник, но тип не определен
  if (utmSource && utmSource !== 'unknown' && utmSource.trim() !== '') {
    return `${utmSource.charAt(0).toUpperCase()}${utmSource.slice(1)}`;
  }
  
  if (dealSource && dealSource !== 'unknown' && dealSource.trim() !== '') {
    return `${dealSource.charAt(0).toUpperCase()}${dealSource.slice(1)}`;
  }
  
  return 'Неизвестный';
}

/**
 * Получение детального описания канала
 */
function getChannelDetails(row) {
  if (!row || !CONFIG.WORKING_AMO_COLUMNS) return { type: 'Неизвестный', source: '', medium: '', campaign: '' };
  
  const utmSource = row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || '';
  const utmMedium = row[CONFIG.WORKING_AMO_COLUMNS.UTM_MEDIUM] || '';
  const utmCampaign = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CAMPAIGN] || '';
  const dealSource = row[CONFIG.WORKING_AMO_COLUMNS.DEAL_SOURCE] || '';
  
  return {
    type: getChannelType(row),
    source: utmSource || dealSource,
    medium: utmMedium,
    campaign: utmCampaign,
    combined: `${utmSource}/${utmMedium}/${utmCampaign}`.replace(/\/+/g, '/').replace(/\/$/, '')
  };
}

/**
 * Получение периодов дат
 */
function getDatePeriods(dateValue) {
  if (!dateValue) return { yearMonth: 'unknown', quarter: 'unknown', year: 'unknown' };
  
  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue);
    } else {
      return { yearMonth: 'unknown', quarter: 'unknown', year: 'unknown' };
    }
    
    if (isNaN(date.getTime())) {
      return { yearMonth: 'unknown', quarter: 'unknown', year: 'unknown' };
    }
    
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const quarter = Math.ceil(month / 3);
    
    return {
      yearMonth: `${year}-${month.toString().padStart(2, '0')}`,
      quarter: `Q${quarter} ${year}`,
      year: year.toString()
    };
  } catch (error) {
    return { yearMonth: 'unknown', quarter: 'unknown', year: 'unknown' };
  }
}

/**
 * Получение данных из листа РАБОЧИЙ_АМО
 */
function getWorkingAmoData() {
  console.log('📊 Загружаем данные из РАБОЧИЙ_АМО...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKING_AMO);
    
    if (!sheet) {
      console.log('❌ Лист РАБОЧИЙ_АМО не найден');
      return [];
    }
    
    // Определяем структуру таблицы
    detectTableStructure();
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log('❌ В листе РАБОЧИЙ_АМО нет данных');
      return [];
    }
    
    // Возвращаем только строки с данными (без заголовков)
    const dataRows = data.slice(1);
    
    console.log(`✅ Загружено ${dataRows.length} строк из РАБОЧИЙ_АМО`);
    
    return dataRows;
    
  } catch (error) {
    console.error('❌ Ошибка загрузки данных РАБОЧИЙ_АМО:', error);
    return [];
  }
}

/**
 * Получение заголовков из листа РАБОЧИЙ_АМО
 */
function getWorkingAmoHeaders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKING_AMO);
    
    if (!sheet) {
      console.log('❌ Лист РАБОЧИЙ_АМО не найден');
      return [];
    }
    
    const headerData = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    return headerData[0] || [];
    
  } catch (error) {
    console.error('❌ Ошибка загрузки заголовков РАБОЧИЙ_АМО:', error);
    return [];
  }
}

/**
 * Проверка существования и валидности листа РАБОЧИЙ_АМО
 */
function validateWorkingAmoSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WORKING_AMO);
    
    if (!sheet) {
      return { valid: false, error: 'Лист не найден' };
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      return { valid: false, error: 'Нет данных в листе' };
    }
    
    if (lastColumn !== 41) {
      return { 
        valid: false, 
        error: `Неправильное количество колонок: ${lastColumn} (ожидается 41)`,
        actualColumns: lastColumn,
        expectedColumns: 41
      };
    }
    
    const headers = getWorkingAmoHeaders();
    const expectedHeaders = [
      'Сделка.ID', 'Сделка.Название', 'Сделка.Статус', 'Сделка.Причина отказа (ОБ)',
      'Сделка.Тип лида', 'Сделка.R.Статусы гостей', 'Сделка.Ответственный'
    ];
    
    for (let i = 0; i < Math.min(7, expectedHeaders.length); i++) {
      if (headers[i] !== expectedHeaders[i]) {
        return {
          valid: false,
          error: `Неправильный заголовок в колонке ${i + 1}: "${headers[i]}" (ожидается "${expectedHeaders[i]}")`,
          actualHeaders: headers.slice(0, 10),
          expectedHeaders: expectedHeaders
        };
      }
    }
    
    return { 
      valid: true, 
      rowCount: lastRow - 1, 
      columnCount: lastColumn,
      message: 'Лист корректный'
    };
    
  } catch (error) {
    return { valid: false, error: error.toString() };
  }
}

/**
 * Создание или восстановление корректной структуры листа РАБОЧИЙ_АМО
 */
function createOrRepairWorkingAmoSheet() {
  console.log('🔧 Создание/восстановление листа РАБОЧИЙ_АМО...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEETS.WORKING_AMO);
    
    if (!sheet) {
      console.log('📋 Создаем новый лист РАБОЧИЙ_АМО...');
      sheet = ss.insertSheet(CONFIG.SHEETS.WORKING_AMO);
    } else {
      console.log('🧹 Очищаем существующий лист РАБОЧИЙ_АМО...');
      sheet.clear();
    }
    
    // Создаем правильные заголовки
    createWorkingAmoHeaders_(sheet);
    
    console.log('✅ Лист РАБОЧИЙ_АМО готов');
    return sheet;
    
  } catch (error) {
    console.error('❌ Ошибка создания листа РАБОЧИЙ_АМО:', error);
    throw error;
  }
}

/**
 * Функция для быстрого исправления структуры существующего листа
 */
function fixWorkingAmoStructureNow() {
  console.log('🚑 АВАРИЙНОЕ ИСПРАВЛЕНИЕ: Исправляем структуру РАБОЧИЙ_АМО...');
  
  try {
    // Проверяем текущее состояние
    const validation = validateWorkingAmoSheet();
    
    if (validation.valid) {
      console.log('✅ Лист уже имеет правильную структуру');
      return validation;
    }
    
    console.log(`❌ Обнаружены проблемы: ${validation.error}`);
    
    // Создаем новый лист с правильной структурой
    createOrRepairWorkingAmoSheet();
    
    // Запускаем чистую сборку данных
    buildWorkingAmoFileClean();
    
    console.log('✅ ИСПРАВЛЕНИЕ ЗАВЕРШЕНО: Лист РАБОЧИЙ_АМО исправлен');
    return { fixed: true, message: 'Структура исправлена и данные загружены' };
    
  } catch (error) {
    console.error('❌ ОШИБКА ИСПРАВЛЕНИЯ:', error);
    return { fixed: false, error: error.toString() };
  }
}

/**
 * Диагностика и отчет о структуре листа
 */
function diagnoseWorkingAmoStructure() {
  console.log('🔍 ДИАГНОСТИКА: Проверяем структуру РАБОЧИЙ_АМО...');
  
  try {
    const validation = validateWorkingAmoSheet();
    const headers = getWorkingAmoHeaders();
    
    const report = {
      timestamp: new Date(),
      validation: validation,
      currentHeaders: headers,
      expectedHeaders: [
        'Сделка.ID', 'Сделка.Название', 'Сделка.Статус', 'Сделка.Причина отказа (ОБ)',
        'Сделка.Тип лида', 'Сделка.R.Статусы гостей', 'Сделка.Ответственный', 'Сделка.Теги'
      ],
      recommendations: []
    };
    
    if (!validation.valid) {
      report.recommendations.push('🔧 Запустите fixWorkingAmoStructureNow() для исправления');
      report.recommendations.push('📋 Используйте createOrRepairWorkingAmoSheet() для создания нового листа');
      report.recommendations.push('🔄 Запустите buildWorkingAmoFileClean() для загрузки данных');
    }
    
    console.log('📊 ДИАГНОСТИКА ЗАВЕРШЕНА');
    console.log('Статус:', validation.valid ? '✅ Корректно' : '❌ Требует исправления');
    console.log('Колонок:', validation.columnCount || 'неизвестно');
    console.log('Строк данных:', validation.rowCount || 'неизвестно');
    
    return report;
    
  } catch (error) {
    console.error('❌ ОШИБКА ДИАГНОСТИКИ:', error);
    return { error: error.toString() };
  }
}

/**
 * УТИЛИТЫ УПРАВЛЕНИЯ ЛИСТАМИ
 */

/**
 * Создание или обновление листа с данными
 */
function createOrUpdateSheet(sheetName, headers, data) {
  console.log(`📋 Создаем/обновляем лист: ${sheetName}`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // Создаем лист если его нет
    if (!sheet) {
      console.log(`➕ Создаем новый лист: ${sheetName}`);
      sheet = ss.insertSheet(sheetName);
    } else {
      console.log(`🧹 Очищаем существующий лист: ${sheetName}`);
      sheet.clear();
    }
    
    // Если нет данных, создаем только заголовки
    if (!data || data.length === 0) {
      if (headers && headers.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        console.log(`✅ Создан лист ${sheetName} только с заголовками (${headers.length} колонок)`);
      }
      return sheet;
    }
    
    // Подготавливаем данные для записи
    const allData = [];
    
    // Добавляем заголовки если они есть
    if (headers && headers.length > 0) {
      allData.push(headers);
    }
    
    // Добавляем данные
    if (Array.isArray(data) && data.length > 0) {
      // Проверяем формат данных
      if (Array.isArray(data[0])) {
        // Данные уже в формате массива массивов
        allData.push(...data);
      } else {
        // Данные в виде объектов, преобразуем
        const dataRows = data.map(row => {
          if (Array.isArray(row)) {
            return row;
          } else if (typeof row === 'object') {
            // Если есть заголовки, используем их как ключи
            if (headers) {
              return headers.map(header => row[header] || '');
            } else {
              return Object.values(row);
            }
          } else {
            return [row];
          }
        });
        allData.push(...dataRows);
      }
    }
    
    // Записываем данные
    if (allData.length > 0) {
      const maxCols = Math.max(...allData.map(row => row.length));
      sheet.getRange(1, 1, allData.length, maxCols).setValues(allData);
      
      // Форматируем заголовки если они есть
      if (headers && headers.length > 0) {
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#e8f4fd');
      }
      
      // Авторазмер колонок
      sheet.autoResizeColumns(1, maxCols);
      
      console.log(`✅ Лист ${sheetName} создан: ${allData.length - (headers ? 1 : 0)} строк данных, ${maxCols} колонок`);
    }
    
    return sheet;
    
  } catch (error) {
    console.error(`❌ Ошибка создания листа ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Создание простого листа с таблицей данных
 */
function createSimpleDataSheet(sheetName, data, title = '') {
  console.log(`📊 Создаем простой лист с данными: ${sheetName}`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // Создаем лист если его нет
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    let currentRow = 1;
    
    // Добавляем заголовок если есть
    if (title) {
      sheet.getRange(currentRow, 1).setValue(title);
      sheet.getRange(currentRow, 1).setFontWeight('bold').setFontSize(14);
      currentRow += 2;
    }
    
    // Добавляем данные
    if (data && data.length > 0) {
      const maxCols = Math.max(...data.map(row => Array.isArray(row) ? row.length : 1));
      sheet.getRange(currentRow, 1, data.length, maxCols).setValues(data);
      
      // Форматируем первую строку как заголовки
      if (data.length > 0) {
        const headerRange = sheet.getRange(currentRow, 1, 1, maxCols);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#f0f0f0');
      }
      
      // Авторазмер колонок
      sheet.autoResizeColumns(1, maxCols);
    }
    
    console.log(`✅ Простой лист ${sheetName} создан с ${data ? data.length : 0} строками`);
    return sheet;
    
  } catch (error) {
    console.error(`❌ Ошибка создания простого листа ${sheetName}:`, error);
    throw error;
  }
}

/**
 * Удаление листа если он существует
 */
function deleteSheetIfExists(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (sheet) {
      ss.deleteSheet(sheet);
      console.log(`🗑️ Лист ${sheetName} удален`);
      return true;
    }
    
    return false;
  } catch (error) {
    console.error(`❌ Ошибка удаления листа ${sheetName}:`, error);
    return false;
  }
}

/**
 * Получение или создание листа
 */
function getOrCreateSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log(`➕ Создаем новый лист: ${sheetName}`);
      sheet = ss.insertSheet(sheetName);
    }
    
    return sheet;
  } catch (error) {
    console.error(`❌ Ошибка получения/создания листа ${sheetName}:`, error);
    throw error;
  }
}