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
 * Проверка успешности статуса с учетом всех вариантов для бара
 */
function isSuccessStatus(status) {
  if (!status) return false;
  
  // Приводим к нижнему регистру и убираем лишние пробелы
  const normalizedStatus = status.toString().toLowerCase().trim().replace(/\s+/g, ' ');
  
  return SUCCESS_STATUSES.some(successStatus => 
    normalizedStatus.includes(successStatus)
  );
}

/**
 * Дополнительная функция для проверки оплаченных статусов
 */
function isPaidStatus(status) {
  if (!status) return false;
  
  const paidStatuses = [
    'оплачено',
    'оплачен',
    'успешно реализовано',
    'успешно в рп',
    'paid',
    'payed'
  ];
  
  const normalizedStatus = status.toString().toLowerCase().trim();
  return paidStatuses.some(paidStatus => 
    normalizedStatus.includes(paidStatus)
  );
}

/**
 * Определение способа обращения (звонок или заявка) и источника
 * ИСПРАВЛЕНО: правильный подсчет звонков по колонке M
 */
function getContactMethodAndSource(row) {
  if (!row || !CONFIG.WORKING_AMO_COLUMNS) {
    return { 
      method: 'unknown',
      source: 'unknown', 
      name: 'Неизвестный', 
      type: 'Неизвестный',
      channel: 'Неизвестный'
    };
  }
  
  // ИСПРАВЛЕНО: Проверяем колонку M (Сделка.Номер линии MANGO OFFICE)
  const mangoLine = row[CONFIG.WORKING_AMO_COLUMNS.DEAL_MANGO_LINE] || '';
  const hasCallTracking = mangoLine && mangoLine.toString().trim() !== '';
  
  // Проверяем наличие UTM меток
  const utmSource = (row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || '').toLowerCase().trim();
  const hasUtm = utmSource && utmSource !== '';
  
  // ИСПРАВЛЕНО: Если есть номер в колонке M - это ЗВОНОК
  if (hasCallTracking) {
    const cleanPhone = mangoLine.toString().replace(/\D/g, '');
    if (CALL_TRACKING_MAPPING[cleanPhone]) {
      return {
        method: 'call',
        ...CALL_TRACKING_MAPPING[cleanPhone]
      };
    }
    // Неизвестный номер колл-трекинга
    return {
      method: 'call',
      source: 'call_unknown',
      name: `Звонок с номера ${mangoLine}`,
      type: 'Звонки',
      channel: `Звонок (${mangoLine})`
    };
  }
  
  // Если есть UTM метки - это ЗАЯВКА с сайта
  if (hasUtm) {
    const cleanUtm = utmSource.replace('#booking', '');
    if (UTM_SOURCE_DETAILED_MAPPING[cleanUtm]) {
      return {
        method: 'form',
        ...UTM_SOURCE_DETAILED_MAPPING[cleanUtm]
      };
    }
    // Неизвестная UTM метка
    return {
      method: 'form',
      source: utmSource,
      name: `Заявка: ${utmSource}`,
      type: 'Заявки',
      channel: utmSource.charAt(0).toUpperCase() + utmSource.slice(1)
    };
  }
  
  // Для старых данных (до июля 2025) смотрим R.Источник сделки
  const dealSource = (row[CONFIG.WORKING_AMO_COLUMNS.DEAL_SOURCE] || '').toLowerCase().trim();
  if (dealSource && dealSource !== '') {
    // Пытаемся определить метод по содержимому
    if (dealSource.includes('звон') || dealSource.includes('call') || dealSource.includes('тел')) {
      return {
        method: 'call',
        source: 'call_old',
        name: 'Звонок (старые данные)',
        type: 'Звонки',
        channel: dealSource
      };
    }
    if (dealSource.includes('сайт') || dealSource.includes('форм') || dealSource.includes('заявк')) {
      return {
        method: 'form',
        source: 'form_old',
        name: 'Заявка с сайта (старые данные)',
        type: 'Заявки',
        channel: 'Сайт'
      };
    }
    
    // Не можем определить метод
    return {
      method: 'unknown',
      source: dealSource,
      name: dealSource.charAt(0).toUpperCase() + dealSource.slice(1),
      type: 'Старые данные',
      channel: dealSource
    };
  }
  
  return { 
    method: 'unknown',
    source: 'unknown', 
    name: 'Неизвестный', 
    type: 'Неизвестный',
    channel: 'Неизвестный'
  };
}

/**
 * Проверка даты для определения периода с правильной разметкой
 */
function hasProperUtmTracking(dateValue) {
  if (!dateValue) return false;
  
  try {
    const date = new Date(dateValue);
    const july2025 = new Date('2025-07-01'); // Правильная разметка началась с июля 2025
    
    return date >= july2025;
  } catch (error) {
    return false;
  }
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

/**
 * РАСШИРЕННАЯ АНАЛИТИКА КАНАЛОВ
 */

/**
 * Вычисление стоимости привлечения клиента (CAC)
 */
function calculateCAC(channelBudget, successfulLeads) {
  if (!successfulLeads || successfulLeads === 0) return 0;
  return channelBudget / successfulLeads;
}

/**
 * Вычисление жизненной ценности клиента (LTV)
 */
function calculateLTV(averageRevenue, repeatRate = 1.2, retentionPeriod = 24) {
  return averageRevenue * repeatRate * (retentionPeriod / 12);
}

/**
 * Вычисление ROMI с детализацией
 */
function calculateDetailedROMI(revenue, budget, additionalCosts = 0) {
  const totalInvestment = budget + additionalCosts;
  if (totalInvestment === 0) return 0;
  return ((revenue - totalInvestment) / totalInvestment) * 100;
}

/**
 * Оценка качества лидов по каналу
 */
function calculateLeadQuality(channel, data) {
  if (!data || data.length === 0) return 0;
  
  const channelData = data.filter(row => getChannelType(row) === channel);
  if (channelData.length === 0) return 0;
  
  let qualityScore = 0;
  let factors = 0;
  
  // Фактор 1: Конверсия в успешные сделки
  const successRate = channelData.filter(row => 
    isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS])).length / channelData.length;
  qualityScore += successRate * 40; // 40% веса
  factors++;
  
  // Фактор 2: Средний чек
  const revenues = channelData
    .filter(row => isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS]))
    .map(row => formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]));
  
  if (revenues.length > 0) {
    const avgRevenue = revenues.reduce((sum, rev) => sum + rev, 0) / revenues.length;
    const normalizedRevenue = Math.min(avgRevenue / 1000000, 1); // Нормализуем до 1
    qualityScore += normalizedRevenue * 30; // 30% веса
    factors++;
  }
  
  // Фактор 3: Скорость закрытия (если есть даты)
  const avgCloseTime = calculateAverageCloseTime(channelData);
  if (avgCloseTime > 0) {
    const speedScore = Math.max(0, (30 - avgCloseTime) / 30); // Чем быстрее, тем лучше
    qualityScore += speedScore * 20; // 20% веса
    factors++;
  }
  
  // Фактор 4: Отсутствие отказов
  const refusalRate = channelData.filter(row => 
    isRefusalStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS])).length / channelData.length;
  qualityScore += (1 - refusalRate) * 10; // 10% веса
  factors++;
  
  return factors > 0 ? Math.min(qualityScore / factors * 10, 10) : 0; // Нормализуем до 10
}

/**
 * Расчет среднего времени закрытия сделки
 */
function calculateAverageCloseTime(channelData) {
  const closedDeals = channelData.filter(row => 
    isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS]) || 
    isRefusalStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS])
  );
  
  if (closedDeals.length === 0) return 0;
  
  let totalDays = 0;
  let validDeals = 0;
  
  closedDeals.forEach(row => {
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const closedAt = row[CONFIG.WORKING_AMO_COLUMNS.CLOSED_AT] || row[CONFIG.WORKING_AMO_COLUMNS.UPDATED_AT];
    
    if (createdAt && closedAt) {
      const created = new Date(createdAt);
      const closed = new Date(closedAt);
      const diffDays = Math.abs((closed - created) / (1000 * 60 * 60 * 24));
      
      if (diffDays > 0 && diffDays < 365) { // Исключаем аномалии
        totalDays += diffDays;
        validDeals++;
      }
    }
  });
  
  return validDeals > 0 ? totalDays / validDeals : 0;
}

/**
 * Анализ времени активности каналов (по часам)
 */
function analyzeChannelHourlyPerformance(data) {
  const hourlyStats = {};
  
  // Инициализируем статистику по часам
  for (let hour = 0; hour < 24; hour++) {
    hourlyStats[hour] = {
      hour: hour,
      leads: 0,
      success: 0,
      revenue: 0,
      channels: {}
    };
  }
  
  data.forEach(row => {
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    if (!createdAt) return;
    
    const date = new Date(createdAt);
    const hour = date.getHours();
    const channel = getChannelType(row);
    const revenue = isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS]) ? 
      formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]) : 0;
    
    if (hourlyStats[hour]) {
      hourlyStats[hour].leads++;
      if (isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS])) {
        hourlyStats[hour].success++;
      }
      hourlyStats[hour].revenue += revenue;
      
      // По каналам
      if (!hourlyStats[hour].channels[channel]) {
        hourlyStats[hour].channels[channel] = { leads: 0, success: 0, revenue: 0 };
      }
      hourlyStats[hour].channels[channel].leads++;
      if (isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS])) {
        hourlyStats[hour].channels[channel].success++;
      }
      hourlyStats[hour].channels[channel].revenue += revenue;
    }
  });
  
  return Object.values(hourlyStats);
}

/**
 * Прогнозирование результатов канала
 */
function predictChannelPerformance(channelHistory, daysToPredict = 30) {
  if (!channelHistory || channelHistory.length < 7) return null;
  
  // Берем последние 30 дней для анализа тренда
  const recentData = channelHistory.slice(-30);
  
  // Вычисляем средние показатели
  const avgLeads = recentData.reduce((sum, day) => sum + (day.leads || 0), 0) / recentData.length;
  const avgConversion = recentData.reduce((sum, day) => sum + (day.conversion || 0), 0) / recentData.length;
  const avgRevenue = recentData.reduce((sum, day) => sum + (day.revenue || 0), 0) / recentData.length;
  
  // Анализируем тренд (простая линейная регрессия)
  const leadsTrend = calculateSimpleTrend(recentData.map(d => d.leads || 0));
  const revenueTrend = calculateSimpleTrend(recentData.map(d => d.revenue || 0));
  
  return {
    predictedLeads: Math.max(0, avgLeads + (leadsTrend * daysToPredict)),
    predictedSuccess: Math.max(0, (avgLeads + (leadsTrend * daysToPredict)) * (avgConversion / 100)),
    predictedRevenue: Math.max(0, avgRevenue + (revenueTrend * daysToPredict)),
    confidence: Math.min(recentData.length / 30, 1) * 100, // Уверенность в прогнозе
    trend: leadsTrend > 0 ? 'растущий' : leadsTrend < 0 ? 'падающий' : 'стабильный'
  };
}

/**
 * Расчет простого тренда
 */
function calculateSimpleTrend(values) {
  if (values.length < 2) return 0;
  
  let sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
  const n = values.length;
  
  for (let i = 0; i < n; i++) {
    sumX += i;
    sumY += values[i];
    sumXY += i * values[i];
    sumX2 += i * i;
  }
  
  const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
  return slope;
}

/**
 * Анализ конкурентных преимуществ каналов
 */
function analyzeChannelCompetitiveAdvantages(channels) {
  if (!channels || channels.length === 0) return [];
  
  // Находим лидеров по различным метрикам
  const leaders = {
    volume: channels.reduce((prev, curr) => prev.leads > curr.leads ? prev : curr),
    conversion: channels.reduce((prev, curr) => prev.conversion > curr.conversion ? prev : curr),
    revenue: channels.reduce((prev, curr) => prev.revenue > curr.revenue ? prev : curr),
    efficiency: channels.reduce((prev, curr) => (prev.romi || 0) > (curr.romi || 0) ? prev : curr),
    quality: channels.reduce((prev, curr) => (prev.qualityScore || 0) > (curr.qualityScore || 0) ? prev : curr)
  };
  
  return channels.map(channel => {
    const advantages = [];
    const opportunities = [];
    
    // Определяем преимущества
    if (channel.channel === leaders.volume.channel) advantages.push('Лидер по объему лидов');
    if (channel.channel === leaders.conversion.channel) advantages.push('Лучшая конверсия');
    if (channel.channel === leaders.revenue.channel) advantages.push('Максимальная выручка');
    if (channel.channel === leaders.efficiency.channel) advantages.push('Высокий ROMI');
    if (channel.channel === leaders.quality.channel) advantages.push('Качественные лиды');
    
    // Определяем возможности для роста
    if (channel.conversion < leaders.conversion.conversion * 0.8) {
      opportunities.push('Улучшить конверсию');
    }
    if ((channel.romi || 0) < (leaders.efficiency.romi || 0) * 0.7) {
      opportunities.push('Оптимизировать затраты');
    }
    if (channel.leads < leaders.volume.leads * 0.5) {
      opportunities.push('Масштабировать объемы');
    }
    
    return {
      channel: channel.channel,
      advantages: advantages,
      opportunities: opportunities,
      competitivePosition: advantages.length > 2 ? 'Сильная' : advantages.length > 0 ? 'Средняя' : 'Слабая',
      improvementPotential: opportunities.length
    };
  });
}

/**
 * Вычисление эффективности менеджера
 */
function calculateManagerEfficiency(managerStats) {
  if (!managerStats) return 0;
  
  let efficiency = 0;
  let factors = 0;
  
  // Фактор 1: Конверсия (40% веса)
  if (managerStats.conversionRate !== undefined) {
    efficiency += (managerStats.conversionRate / 100) * 40;
    factors++;
  }
  
  // Фактор 2: Объем работы (30% веса)
  if (managerStats.totalDeals !== undefined) {
    const volumeScore = Math.min(managerStats.totalDeals / 50, 1); // Нормализуем к 50 сделкам
    efficiency += volumeScore * 30;
    factors++;
  }
  
  // Фактор 3: Динамика (30% веса)
  if (managerStats.currentMonthDeals !== undefined && managerStats.previousMonthDeals !== undefined) {
    const growth = managerStats.previousMonthDeals > 0 ? 
      (managerStats.currentMonthDeals - managerStats.previousMonthDeals) / managerStats.previousMonthDeals : 0;
    const growthScore = Math.max(0, Math.min(growth + 0.5, 1)); // От -50% до +50% роста
    efficiency += growthScore * 30;
    factors++;
  }
  
  return factors > 0 ? efficiency / factors * 10 : 0; // Нормализуем до 10
}

/**
 * Форматирование даты для отчетов
 */
function formatDate(dateValue, format = 'DD.MM.YYYY') {
  if (!dateValue) return '';
  
  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue);
    } else {
      return '';
    }
    
    if (isNaN(date.getTime())) return '';
    
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    const hours = date.getHours().toString().padStart(2, '0');
    const minutes = date.getMinutes().toString().padStart(2, '0');
    
    switch (format) {
      case 'DD.MM.YYYY':
        return `${day}.${month}.${year}`;
      case 'DD.MM.YYYY HH:mm':
        return `${day}.${month}.${year} ${hours}:${minutes}`;
      case 'MM/DD/YYYY':
        return `${month}/${day}/${year}`;
      case 'YYYY-MM-DD':
        return `${year}-${month}-${day}`;
      default:
        return `${day}.${month}.${year}`;
    }
  } catch (error) {
    return '';
  }
}

/**
 * МОДУЛЬ: СПЕЦИАЛИЗИРОВАННЫЙ АНАЛИЗ ДЛЯ БАРА
 * Красивые и информативные отчеты с правильными метриками
 */

/**
 * Подсчет правильных метрик для бара
 */
function calculateBarMetrics(data) {
  const metrics = {
    totalLeads: 0,           // Общее количество лидов (бронирований)
    successfulVisits: 0,     // Пришедшие гости
    totalPrepayment: 0,      // Сумма предоплат (колонка I - Бюджет)
    totalFactAmount: 0,      // Фактическая выручка (колонка P - Счет факт)
    avgCheck: 0,             // Средний чек
    conversionRate: 0,       // Конверсия бронь -> визит
    refusals: 0,             // Отказы
    callTracking: 0,         // Звонки через колл-трекинг
    newGuests: 0,            // Новые гости
    repeatGuests: 0,         // Повторные гости
    referralGuests: 0        // Гости по сарафану
  };
  
  data.forEach(row => {
    // Считаем все лиды (каждая строка = 1 лид/бронирование)
    metrics.totalLeads++;
    
    // Статус сделки
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const guestStatus = row[CONFIG.WORKING_AMO_COLUMNS.GUEST_STATUS] || '';
    const isReferral = row[CONFIG.WORKING_AMO_COLUMNS.REFERRAL_GUESTS] || '';
    
    // Успешные визиты (гость пришел)
    if (isSuccessStatus(status)) {
      metrics.successfulVisits++;
      
      // Фактическая сумма (сколько гость потратил)
      const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
      metrics.totalFactAmount += factAmount;
    } else if (isRefusalStatus(status)) {
      metrics.refusals++;
    }
    
    // Предоплата (колонка I - Бюджет)
    const prepayment = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    metrics.totalPrepayment += prepayment;
    
    // Колл-трекинг (если есть номер MANGO)
    const mangoLine = row[CONFIG.WORKING_AMO_COLUMNS.DEAL_MANGO_LINE] || '';
    if (mangoLine && mangoLine.trim() !== '') {
      metrics.callTracking++;
    }
    
    // Тип гостя
    if (guestStatus.toLowerCase().includes('новый')) {
      metrics.newGuests++;
    } else if (guestStatus.toLowerCase().includes('повторный')) {
      metrics.repeatGuests++;
    }
    
    // Сарафанное радио
    if (isReferral && (isReferral.toLowerCase() === 'да' || isReferral === '1')) {
      metrics.referralGuests++;
    }
  });
  
  // Рассчитываем производные метрики
  metrics.avgCheck = metrics.successfulVisits > 0 ? 
    metrics.totalFactAmount / metrics.successfulVisits : 0;
  
  metrics.conversionRate = metrics.totalLeads > 0 ? 
    (metrics.successfulVisits / metrics.totalLeads) * 100 : 0;
  
  return metrics;
}

/**
 * ТЕСТОВАЯ ФУНКЦИЯ ДЛЯ ПРОВЕРКИ АНАЛИЗА КОНКРЕТНОЙ СТРОКИ
 */
function testRowAnalysis(rowNumber) {
  const data = getWorkingAmoData();
  if (!data || rowNumber > data.length) {
    console.log('❌ Строка не найдена или нет данных');
    return;
  }
  
  const row = data[rowNumber - 1];
  
  console.log(`🔍 ДЕТАЛЬНЫЙ АНАЛИЗ СТРОКИ ${rowNumber}:`);
  console.log(`ID сделки (A): "${row[0]}"`);
  console.log(`Название (B): "${row[1]}"`);
  console.log(`Статус (C): "${row[2]}"`);
  console.log(`MANGO линия (М): "${row[12] || 'пусто'}"`);      // Колонка M
  console.log(`UTM_SOURCE (AB): "${row[27] || 'пусто'}"`);      // Колонка AB
  console.log(`Счет факт (P): "${row[15] || 0}"`);             // Колонка P
  
  // Анализируем тип обращения
  const contactInfo = getContactMethodAndSource(row);
  
  console.log('--- РЕЗУЛЬТАТ АНАЛИЗА ---');
  console.log(`Способ обращения: ${contactInfo.method}`);
  console.log(`Название источника: ${contactInfo.name}`);
  console.log(`Тип: ${contactInfo.type}`);
  console.log(`Канал: ${contactInfo.channel}`);
  
  // Проверяем статус
  const status = row[2] || '';
  const isSuccess = isSuccessStatus(status);
  console.log(`Успешный статус: ${isSuccess ? 'ДА' : 'НЕТ'}`);
  
  // Определяем логику подсчета
  const mangoLine = row[12] || '';
  const utmSource = row[27] || '';
  const hasMangoLine = mangoLine && mangoLine.toString().trim() !== '';
  const hasUtm = utmSource && utmSource.toString().trim() !== '';
  
  console.log('--- ЛОГИКА ПОДСЧЕТА ---');
  if (hasMangoLine) {
    console.log(`✅ ЗАСЧИТЫВАЕТСЯ КАК ЗВОНОК (есть номер в колонке М: "${mangoLine}")`);
  } else if (hasUtm) {
    console.log(`✅ ЗАСЧИТЫВАЕТСЯ КАК ЗАЯВКА (есть UTM в колонке AB: "${utmSource}")`);
  } else {
    console.log(`❓ НЕ ЗАСЧИТЫВАЕТСЯ (нет ни MANGO, ни UTM)`);
  }
  
  return {
    rowNumber,
    isCall: hasMangoLine,
    isForm: hasUtm,
    isSuccess,
    contactInfo
  };
}

/**
 * Анализ источников трафика для бара
 */
function analyzeBarChannels(data) {
  const channels = {};
  
  data.forEach(row => {
    const contactInfo = getContactMethodAndSource(row);
    const source = contactInfo.channel;
    
    if (!channels[source]) {
      channels[source] = {
        leads: 0,
        visits: 0,
        prepayment: 0,
        revenue: 0,
        conversionRate: 0,
        method: contactInfo.method,
        type: contactInfo.type
      };
    }
    
    channels[source].leads++;
    
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    if (isSuccessStatus(status)) {
      channels[source].visits++;
      const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
      channels[source].revenue += factAmount;
    }
    
    const prepayment = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    channels[source].prepayment += prepayment;
  });
  
  // Рассчитываем конверсию для каждого канала
  Object.keys(channels).forEach(channel => {
    channels[channel].conversionRate = channels[channel].leads > 0 ?
      (channels[channel].visits / channels[channel].leads) * 100 : 0;
  });
  
  return channels;
}

/**
 * Анализ менеджеров отдела бронирования
 */
function analyzeBookingManagers(data) {
  const managers = {};
  
  data.forEach(row => {
    const manager = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || 'Не указан';
    
    if (!managers[manager]) {
      managers[manager] = {
        totalBookings: 0,      // Всего бронирований
        successfulVisits: 0,   // Успешные визиты
        prepaymentSum: 0,      // Сумма предоплат
        revenueSum: 0,         // Фактическая выручка
        avgCheck: 0,           // Средний чек
        conversionRate: 0,     // Конверсия
        newGuests: 0,          // Привлечено новых гостей
        repeatGuests: 0        // Работа с повторными
      };
    }
    
    managers[manager].totalBookings++;
    
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const guestStatus = row[CONFIG.WORKING_AMO_COLUMNS.GUEST_STATUS] || '';
    
    if (isSuccessStatus(status)) {
      managers[manager].successfulVisits++;
      const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
      managers[manager].revenueSum += factAmount;
    }
    
    const prepayment = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    managers[manager].prepaymentSum += prepayment;
    
    // Тип гостя
    if (guestStatus.toLowerCase().includes('новый')) {
      managers[manager].newGuests++;
    } else if (guestStatus.toLowerCase().includes('повторный')) {
      managers[manager].repeatGuests++;
    }
  });
  
  // Рассчитываем производные метрики
  Object.keys(managers).forEach(manager => {
    const data = managers[manager];
    data.avgCheck = data.successfulVisits > 0 ?
      data.revenueSum / data.successfulVisits : 0;
    data.conversionRate = data.totalBookings > 0 ?
      (data.successfulVisits / data.totalBookings) * 100 : 0;
  });
  
  return managers;
}

function createBarDashboard() {
  console.log('🍺 Создаем дашборд для бара...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    // Создаем лист для дашборда
    const sheet = createOrUpdateSheet('🍺 ДАШБОРД БАРА');
    
    // Основные метрики
    const metrics = calculateBarMetrics(data);
    const channels = analyzeBarChannels(data);
    const managers = analyzeBookingManagers(data);
    
    let currentRow = 1;
    
    // ЗАГОЛОВОК
    const titleRange = sheet.getRange(currentRow, 1, 1, 10);
    titleRange.merge();
    titleRange.setValue('🍺 АНАЛИТИКА БАРА - ПОЛНЫЙ ДАШБОРД');
    titleRange.setBackground('#1a237e')
              .setFontColor('#ffffff')
              .setFontSize(18)
              .setFontWeight('bold')
              .setHorizontalAlignment('center');
    
    currentRow += 3;
    
    // КЛЮЧЕВЫЕ МЕТРИКИ (KPI CARDS)
    const kpiData = [
      ['📋 ВСЕГО БРОНИРОВАНИЙ', metrics.totalLeads, '✅ УСПЕШНЫХ ВИЗИТОВ', metrics.successfulVisits],
      ['💰 СУММА ПРЕДОПЛАТ', formatCurrency(metrics.totalPrepayment), '💵 ФАКТИЧЕСКАЯ ВЫРУЧКА', formatCurrency(metrics.totalFactAmount)],
      ['📊 КОНВЕРСИЯ', metrics.conversionRate.toFixed(1) + '%', '🧾 СРЕДНИЙ ЧЕК', formatCurrency(metrics.avgCheck)],
      ['📞 ЗВОНКИ (MANGO)', metrics.callTracking, '❌ ОТКАЗЫ', metrics.refusals]
    ];
    
    // Создаем красивые KPI карточки
    kpiData.forEach((row, index) => {
      const rowNum = currentRow + index;
      
      // Первая метрика
      sheet.getRange(rowNum, 1, 1, 2).merge();
      sheet.getRange(rowNum, 1).setValue(row[0])
           .setBackground('#e3f2fd')
           .setFontWeight('bold');
      
      sheet.getRange(rowNum, 3, 1, 2).merge();
      sheet.getRange(rowNum, 3).setValue(row[1])
           .setBackground('#ffffff')
           .setFontSize(14)
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
      
      // Вторая метрика
      sheet.getRange(rowNum, 6, 1, 2).merge();
      sheet.getRange(rowNum, 6).setValue(row[2])
           .setBackground('#e8f5e9')
           .setFontWeight('bold');
      
      sheet.getRange(rowNum, 8, 1, 2).merge();
      sheet.getRange(rowNum, 8).setValue(row[3])
           .setBackground('#ffffff')
           .setFontSize(14)
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
    });
    
    currentRow += kpiData.length + 2;
    
    // АНАЛИЗ ТИПОВ ГОСТЕЙ
    const guestTypeTitle = sheet.getRange(currentRow, 1, 1, 10);
    guestTypeTitle.merge();
    guestTypeTitle.setValue('👥 АНАЛИЗ ТИПОВ ГОСТЕЙ');
    guestTypeTitle.setBackground('#3949ab')
                  .setFontColor('#ffffff')
                  .setFontWeight('bold')
                  .setFontSize(14);
    
    currentRow += 2;
    
    const guestData = [
      ['🆕 Новые гости', metrics.newGuests, (metrics.newGuests / metrics.totalLeads * 100).toFixed(1) + '%'],
      ['🔄 Повторные гости', metrics.repeatGuests, (metrics.repeatGuests / metrics.totalLeads * 100).toFixed(1) + '%'],
      ['💬 По рекомендации (сарафан)', metrics.referralGuests, (metrics.referralGuests / metrics.totalLeads * 100).toFixed(1) + '%']
    ];
    
    const guestHeaders = ['Тип гостей', 'Количество', 'Доля от общего'];
    sheet.getRange(currentRow, 1, 1, 3).setValues([guestHeaders])
         .setBackground('#7986cb')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    sheet.getRange(currentRow + 1, 1, guestData.length, 3).setValues(guestData);
    
    currentRow += guestData.length + 3;
    
    // ТОП ИСТОЧНИКОВ ТРАФИКА
    const channelTitle = sheet.getRange(currentRow, 1, 1, 10);
    channelTitle.merge();
    channelTitle.setValue('📈 ТОП ИСТОЧНИКОВ ПРИВЛЕЧЕНИЯ');
    channelTitle.setBackground('#00897b')
                .setFontColor('#ffffff')
                .setFontWeight('bold')
                .setFontSize(14);
    
    currentRow += 2;
    
    // Сортируем каналы по количеству лидов
    const sortedChannels = Object.entries(channels)
      .sort((a, b) => b[1].leads - a[1].leads)
      .slice(0, 10);
    
    const channelHeaders = ['Источник', 'Способ', 'Бронирования', 'Визиты', 'Конверсия', 'Выручка', 'Средний чек'];
    sheet.getRange(currentRow, 1, 1, channelHeaders.length).setValues([channelHeaders])
         .setBackground('#4db6ac')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    const channelData = sortedChannels.map(([source, data]) => [
      source,
      data.method === 'call' ? '📞 Звонок' : data.method === 'form' ? '📝 Заявка' : '❓ Неизвестно',
      data.leads,
      data.visits,
      data.conversionRate.toFixed(1) + '%',
      formatCurrency(data.revenue),
      formatCurrency(data.visits > 0 ? data.revenue / data.visits : 0)
    ]);
    
    if (channelData.length > 0) {
      sheet.getRange(currentRow + 1, 1, channelData.length, channelHeaders.length).setValues(channelData);
    }
    
    currentRow += Math.max(channelData.length, 1) + 3;
    
    // РЕЙТИНГ МЕНЕДЖЕРОВ
    const managerTitle = sheet.getRange(currentRow, 1, 1, 10);
    managerTitle.merge();
    managerTitle.setValue('🏆 РЕЙТИНГ МЕНЕДЖЕРОВ ОТДЕЛА БРОНИРОВАНИЯ');
    managerTitle.setBackground('#d32f2f')
                .setFontColor('#ffffff')
                .setFontWeight('bold')
                .setFontSize(14);
    
    currentRow += 2;
    
    // Сортируем менеджеров по конверсии
    const sortedManagers = Object.entries(managers)
      .sort((a, b) => b[1].conversionRate - a[1].conversionRate)
      .slice(0, 15);
    
    const managerHeaders = ['Менеджер', 'Брони', 'Визиты', 'Конверсия', 'Выручка', 'Средний чек', 'Новые', 'Повторные'];
    sheet.getRange(currentRow, 1, 1, managerHeaders.length).setValues([managerHeaders])
         .setBackground('#ef5350')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    const managerData = sortedManagers.map(([manager, data]) => [
      manager,
      data.totalBookings,
      data.successfulVisits,
      data.conversionRate.toFixed(1) + '%',
      formatCurrency(data.revenueSum),
      formatCurrency(data.avgCheck),
      data.newGuests,
      data.repeatGuests
    ]);
    
    if (managerData.length > 0) {
      sheet.getRange(currentRow + 1, 1, managerData.length, managerHeaders.length).setValues(managerData);
      
      // Подсветка лучших менеджеров
      for (let i = 0; i < Math.min(3, managerData.length); i++) {
        sheet.getRange(currentRow + 1 + i, 1, 1, managerHeaders.length)
             .setBackground(i === 0 ? '#fff9c4' : i === 1 ? '#f0f4c3' : '#e6ee9c');
      }
    }
    
    // Автоматическая подгонка ширины колонок
    sheet.autoResizeColumns(1, 10);
    
    // Добавляем границы для всех таблиц
    const dataRange = sheet.getRange(1, 1, currentRow + managerData.length, 10);
    dataRange.setBorder(true, true, true, true, true, true);
    
    console.log('✅ Дашборд бара создан успешно!');
    
  } catch (error) {
    console.error('❌ Ошибка создания дашборда:', error);
    throw error;
  }
}