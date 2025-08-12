/**
 * 🛠️ УНИВЕРСАЛЬНЫЕ ФУНКЦИИ ДЛЯ ВСЕХ ОТЧЁТОВ
 * Все общие утилиты в одном месте
 * Версия: 2.0
 */

// ===== УПРАВЛЕНИЕ ЛИСТАМИ =====

/**
 * Создает лист если его нет или возвращает существующий
 * @param {string} sheetName - Название листа
 * @returns {Sheet} Объект листа
 */
function getOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    try {
      sheet = ss.insertSheet(sheetName);
      logInfo_('SHEET_CREATE', `Создан новый лист: "${sheetName}"`);
    } catch (error) {
      logError_('SHEET_CREATE', `Ошибка создания листа "${sheetName}"`, error);
      throw error;
    }
  }
  
  return sheet;
}

/**
 * Применяет шрифт PT Sans ко всем листам
 */
function applyPtSansAllSheets_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    sheets.forEach(sheet => {
      try {
        const range = sheet.getDataRange();
        if (range.getNumRows() > 0 && range.getNumColumns() > 0) {
          range.setFontFamily('PT Sans');
        }
      } catch (error) {
        logWarning_('FORMATTING', `Не удалось применить шрифт к листу "${sheet.getName()}"`, error);
      }
    });
    
    logInfo_('FORMATTING', `Шрифт PT Sans применён к ${sheets.length} листам`);
  } catch (error) {
    logError_('FORMATTING', 'Ошибка применения шрифта PT Sans', error);
  }
}

/**
 * Универсальная функция форматирования листов аналитики
 * @param {SpreadsheetApp.Sheet} sheet - Лист для форматирования
 * @param {string} title - Заголовок отчета
 */
function applySheetFormatting_(sheet, title) {
  if (!sheet) return;
  
  try {
    // Применяем базовый шрифт
    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() > 0) {
      dataRange.setFontFamily(CONFIG.DEFAULT_FONT || 'PT Sans');
      dataRange.setFontSize(10);
      
      // Форматируем заголовки (первая строка)
      if (dataRange.getNumRows() > 0) {
        const headerRange = sheet.getRange(1, 1, 1, dataRange.getNumColumns());
        headerRange
          .setFontWeight('bold')
          .setFontSize(11)
          .setBackground(CONFIG.COLORS.HEADER_BG || '#4285f4')
          .setFontColor(CONFIG.COLORS.HEADER_TEXT || '#ffffff')
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
      
      // Автоширина колонок
      for (let col = 1; col <= dataRange.getNumColumns(); col++) {
        sheet.autoResizeColumn(col);
        
        // Ограничиваем ширину колонок
        const currentWidth = sheet.getColumnWidth(col);
        if (currentWidth > 300) {
          sheet.setColumnWidth(col, 300);
        } else if (currentWidth < 80) {
          sheet.setColumnWidth(col, 80);
        }
      }
      
      // Замораживаем первую строку
      sheet.setFrozenRows(1);
    }
    
    logDebug_('FORMATTING', `Форматирование применено к листу "${sheet.getName()}"`);
    
  } catch (error) {
    logWarning_('FORMATTING', `Ошибка форматирования листа "${sheet.getName()}"`, error);
  }
}

/**
 * Обновляет время последнего обновления в ячейке A1
 * @param {string} sheetName - Название листа
 */
function updateLastUpdateTime_(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      logWarning_('TIME_UPDATE', `Лист "${sheetName}" не найден для обновления времени`);
      return;
    }
    
    const now = getCurrentDateMoscow_();
    const timeString = `Последнее обновление: ${formatDateTime_(now)}`;
    
    sheet.getRange('A1').setValue(timeString);
    sheet.getRange('A1')
      .setFontSize(9)
      .setFontStyle('italic')
      .setFontColor('#666666');
    
    logDebug_('TIME_UPDATE', `Время обновлено в листе "${sheetName}"`);
    
  } catch (error) {
    logWarning_('TIME_UPDATE', `Ошибка обновления времени в листе "${sheetName}"`, error);
  }
}

/**
 * Универсальная функция создания диаграмм
 * @param {SpreadsheetApp.Sheet} sheet - Лист для диаграммы
 * @param {string} chartType - Тип диаграммы ('pie', 'column', 'line')
 * @param {Array} data - Данные для диаграммы
 * @param {Object} options - Параметры диаграммы
 * @returns {SpreadsheetApp.EmbeddedChart|null} Созданная диаграмма или null
 */
function createChart_(sheet, chartType, data, options = {}) {
  try {
    if (!sheet || !data || data.length < 2) {
      logWarning_('CHARTS', 'Недостаточно данных для создания диаграммы');
      return null;
    }
    
    const {
      startRow = 1,
      startCol = 8,
      title = 'Диаграмма',
      width = 500,
      height = 350,
      position = { row: 3, col: 8 },
      legend = 'right',
      hAxisTitle = '',
      vAxisTitle = ''
    } = options;
    
    // Записываем данные в лист
    sheet.getRange(startRow, startCol, data.length, data[0].length).setValues(data);
    const dataRange = sheet.getRange(startRow, startCol, data.length, data[0].length);
    
    // Создаём диаграмму в зависимости от типа
    let chartBuilder;
    
    switch (chartType.toLowerCase()) {
      case 'pie':
        chartBuilder = Charts.newPieChart()
          .addRange(dataRange)
          .setOption('legend', { position: legend });
        break;
        
      case 'column':
        chartBuilder = Charts.newColumnChart()
          .addRange(dataRange)
          .setOption('legend', { position: legend })
          .setOption('hAxis', { title: hAxisTitle })
          .setOption('vAxis', { title: vAxisTitle });
        break;
        
      case 'line':
        chartBuilder = Charts.newLineChart()
          .addRange(dataRange)
          .setOption('legend', { position: legend })
          .setOption('hAxis', { title: hAxisTitle })
          .setOption('vAxis', { title: vAxisTitle })
          .setOption('curveType', 'function');
        break;
        
      default:
        logWarning_('CHARTS', `Неизвестный тип диаграммы: ${chartType}`);
        return null;
    }
    
    // Применяем общие настройки
    const chart = chartBuilder
      .setOption('title', title)
      .setOption('titleTextStyle', { fontSize: 14, bold: true })
      .setOption('chartArea', { width: '80%', height: '80%' })
      .setPosition(position.row, position.col, 0, 0)
      .setOption('width', width)
      .setOption('height', height)
      .build();
    
    sheet.insertChart(chart);
    logDebug_('CHARTS', `Диаграмма "${title}" создана в листе "${sheet.getName()}"`);
    return chart;
    
  } catch (error) {
    logWarning_('CHARTS', `Ошибка создания диаграммы "${options.title || 'Неизвестная'}"`, error);
    return null;
  }
}

// ===== ЛОГИРОВАНИЕ =====

/**
 * Функции логирования для отладки
 */
function logInfo_(module, message, details = null) {
  try {
    console.log(`[INFO] ${module}: ${message}`, details || '');
  } catch (e) {
    // Fallback если console недоступен
  }
}

function logWarning_(module, message, details = null) {
  try {
    console.warn(`[WARNING] ${module}: ${message}`, details || '');
  } catch (e) {
    // Fallback если console недоступен
  }
}

function logError_(module, message, details = null) {
  try {
    console.error(`[ERROR] ${module}: ${message}`, details || '');
  } catch (e) {
    // Fallback если console недоступен
  }
}

// ===== ЧТЕНИЕ И ОБРАБОТКА ДАННЫХ =====

/**
 * Универсальная функция чтения листа
 */
function readSheet(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Лист "${sheetName}" не найден`);
  
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { header: [], rows: [] };
  
  // Ищем строку заголовков (обычно 2-я, но проверяем первые 5)
  let headerRow = 1;
  for (let i = 0; i < Math.min(5, values.length); i++) {
    const row = values[i].map(String);
    if (row.includes('ID') && (row.includes('Статус') || row.includes('Дата создания'))) {
      headerRow = i;
      break;
    }
  }
  
  const header = values[headerRow].map(String);
  const rows = values.slice(headerRow + 1).filter(r => r.some(x => String(x).trim() !== ''));
  
  return { header, rows };
}

/**
 * Гибкое чтение листа с автоопределением заголовков
 */
function readSheetFlexible(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Лист "${sheetName}" не найден`);
  
  const vals = sh.getDataRange().getValues();
  if (vals.length === 0) return { header: [], rows: [], idx: {} };
  
  const header = (vals[0] || []).map(String);
  const rows = vals.slice(1).filter(r => r.some(x => String(x).trim() !== ''));
  
  // Базовые индексы для частых полей
  const idx = {
    id: findColumnIndex(header, ['ID', 'id', 'Идентификатор']),
    phone: findColumnIndex(header, ['Телефон', 'Phone', 'Контакт.Телефон']),
    date: findColumnIndex(header, ['Дата создания', 'DATE', 'Created']),
    status: findColumnIndex(header, ['Статус', 'Status', 'Сделка.Статус']),
    budget: findColumnIndex(header, ['Бюджет', 'Budget', 'Сумма ₽']),
    utm: findColumnIndex(header, ['utm_source', 'UTM_SOURCE', 'Источник'])
  };
  
  return { header, rows, idx };
}

// ===== ПОИСК КОЛОНОК =====

/**
 * Универсальный поиск индексов колонок
 */
function findColumns(header, columnNames) {
  const norm = header.map(h => String(h).trim().toLowerCase());
  const indices = {};
  
  for (const [key, candidates] of Object.entries(columnNames)) {
    indices[key] = -1;
    const candArray = Array.isArray(candidates) ? candidates : [candidates];
    
    for (const name of candArray) {
      const normalized = String(name).trim().toLowerCase();
      const idx = norm.findIndex(h => h === normalized || h.includes(normalized));
      if (idx > -1) {
        indices[key] = idx;
        break;
      }
    }
  }
  
  return indices;
}

/**
 * Поиск одного индекса колонки
 */
function findColumnIndex(header, candidates) {
  const norm = header.map(h => String(h).trim().toLowerCase());
  const candArray = Array.isArray(candidates) ? candidates : [candidates];
  
  for (const name of candArray) {
    const normalized = String(name).trim().toLowerCase();
    const idx = norm.findIndex(h => h === normalized || h.includes(normalized));
    if (idx > -1) return idx;
  }
  
  return -1;
}

// ===== ПРЕОБРАЗОВАНИЕ И НОРМАЛИЗАЦИЯ ДАННЫХ =====

/**
 * Нормализация статуса сделки
 */
function normalizeStatus(raw) {
  let s = String(raw || '').trim();
  s = s.replace(/^\s*все\s*бары\s*сети\s*(?:[\/\\|:\-–—]\s*)?/i, '');
  s = s.replace(/\s*(?:-|–|—)\s*.*$/, '').trim();
  
  const rules = [
    [/оплач/i, 'Оплачено'],
    [/успеш.*в\s*рп/i, 'Успешно в РП'],
    [/успеш.*реализ/i, 'Успешно реализовано'],
    [/закрыто.*не\s*реализ/i, 'Закрыто и не реализовано'],
    [/новый/i, 'Новый лид'],
    [/первич.*контакт/i, 'Первичный контакт'],
    [/квалифик/i, 'Квалификация'],
    [/предложен/i, 'Предложение отправлено']
  ];
  
  for (const [re, norm] of rules) {
    if (re.test(s)) return norm;
  }
  
  return s || 'Новый лид';
}

/**
 * Очистка и нормализация телефона
 */
function cleanPhone(phone) {
  if (!phone) return '';
  return String(phone)
    .replace(/\D/g, '')
    .replace(/^[78]/, '')
    .slice(0, 10);
}

/**
 * Преобразование в число
 */
function toNumber(v) {
  if (v == null || v === '') return 0;
  const n = Number(String(v).replace(/\s/g, '').replace(',', '.'));
  return isNaN(n) ? 0 : n;
}

/**
 * Преобразование в дату
 */
function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Форматирование даты для отчётов
 */
function formatDate(date, format = 'dd.MM.yyyy') {
  if (!date) return '';
  if (!(date instanceof Date)) date = new Date(date);
  return Utilities.formatDate(date, CONFIG.TIMEZONE, format);
}

/**
 * Очистка строки от лишних символов
 */
function cleanString(str) {
  return String(str || '').trim().replace(/\s+/g, ' ');
}

// ===== РАБОТА С ТАБЛИЦАМИ И ОФОРМЛЕНИЕ =====

/**
 * Универсальная запись таблицы с оформлением
 */
function writeTable(sheet, startRow, startCol, header, data, options = {}) {
  const opts = {
    headerBg: '#f3f3f3',
    headerColor: '#000000',
    zebraColor1: '#ffffff',
    zebraColor2: '#f8fcff',
    borderColor: '#cccccc',
    fontFamily: CONFIG.FONT,
    ...options
  };
  
  ensureSize(sheet, startRow + data.length, startCol + header.length - 1);
  
  // Заголовок
  const headerRange = sheet.getRange(startRow, startCol, 1, header.length);
  headerRange.setValues([header])
    .setFontWeight('bold')
    .setBackground(opts.headerBg)
    .setFontColor(opts.headerColor)
    .setFontFamily(opts.fontFamily)
    .setBorder(true, true, true, true, true, true, opts.borderColor, SpreadsheetApp.BorderStyle.SOLID);
  
  // Данные
  if (data.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, startCol, data.length, header.length);
    dataRange.setValues(data)
      .setFontFamily(opts.fontFamily)
      .setBorder(true, true, true, true, true, true, opts.borderColor, SpreadsheetApp.BorderStyle.SOLID);
    
    // Зебра-оформление
    const backgrounds = [];
    for (let i = 0; i < data.length; i++) {
      const color = i % 2 === 0 ? opts.zebraColor1 : opts.zebraColor2;
      backgrounds.push(Array(header.length).fill(color));
    }
    dataRange.setBackgrounds(backgrounds);
  }
  
  return { 
    headerRange: sheet.getRange(startRow, startCol, 1, header.length),
    dataRange: data.length > 0 ? sheet.getRange(startRow + 1, startCol, data.length, header.length) : null,
    totalRows: data.length + 1,
    totalCols: header.length
  };
}

/**
 * Обеспечение размера листа
 */
function ensureSize(sheet, needRows, needCols) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  
  if (maxRows < needRows) {
    sheet.insertRowsAfter(maxRows, needRows - maxRows);
  }
  if (maxCols < needCols) {
    sheet.insertColumnsAfter(maxCols, needCols - maxCols);
  }
}

/**
 * Установка шрифта для всего листа
 */
function setFontAll(sheet, fontName = CONFIG.FONT) {
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily(fontName);
}

/**
 * Автоизменение размера колонок
 */
function autoResizeColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  for (let i = 1; i <= lastCol; i++) {
    try {
      sheet.autoResizeColumn(i);
    } catch (e) {
      // Игнорируем ошибки автоизменения
    }
  }
}

/**
 * Создание секции с заголовком
 */
function createSection(sheet, row, col, title, bgColor = '#e8f4fd') {
  ensureSize(sheet, row, col + title.length);
  sheet.getRange(row, col, 1, 3)
    .merge()
    .setValue(title)
    .setBackground(bgColor)
    .setFontWeight('bold')
    .setFontFamily(CONFIG.FONT)
    .setHorizontalAlignment('center');
  return row + 1;
}

// ===== GPT ИНТЕГРАЦИЯ =====

/**
 * Универсальная функция вызова OpenAI API
 */
function callOpenAI(prompt, options = {}) {
  const {
    model = 'gpt-4o',
    temperature = 0.3,
    maxTokens = 2000,
    systemPrompt = 'Ты — опытный бизнес-аналитик.',
    responseType = 'text',
    fallbackModel = 'gpt-4o-mini'
  } = options;
  
  validateApiKey();
  
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: model,
    temperature: temperature,
    max_tokens: maxTokens,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: prompt }
    ]
  };
  
  if (responseType === 'json') {
    payload.response_format = { type: 'json_object' };
  }
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: { 'Authorization': `Bearer ${CONFIG.API_KEY}` },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 300) {
      // Пробуем запасную модель при ошибке rate limit
      if (responseCode === 429 && fallbackModel && model !== fallbackModel) {
        Logger.log(`Основная модель ${model} недоступна, пробуем ${fallbackModel}`);
        return callOpenAI(prompt, { ...options, model: fallbackModel, fallbackModel: null });
      }
      throw new Error(`OpenAI ошибка ${responseCode}: ${responseText}`);
    }
    
    if (responseType === 'json') {
      const jsonMatch = responseText.match(/{[\s\S]*}/);
      if (jsonMatch && jsonMatch[0]) {
        try {
          return JSON.parse(jsonMatch[0]);
        } catch (e) {
          throw new Error('Не удалось обработать JSON от GPT: ' + e.toString());
        }
      }
      throw new Error('Ответ GPT не содержит валидный JSON');
    } else {
      try {
        const data = JSON.parse(responseText);
        return data?.choices?.[0]?.message?.content?.trim() || '—';
      } catch (e) {
        return responseText; // Если ответ не JSON, возвращаем как есть
      }
    }
    
  } catch (e) {
    Logger.log('Ошибка GPT: ' + e.toString());
    throw e;
  }
}

// ===== УПРАВЛЕНИЕ ТРИГГЕРАМИ =====

/**
 * Установка ежечасного триггера
 */
function setHourlyTrigger(functionName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === functionName)
    .forEach(t => ScriptApp.deleteTrigger(t));
  
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyHours(1)
    .create();
    
  Logger.log(`Установлен ежечасный триггер для ${functionName}`);
}

/**
 * Установка ежедневного триггера
 */
function setDailyTrigger(functionName, hour = 6) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === functionName)
    .forEach(t => ScriptApp.deleteTrigger(t));
  
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .inTimezone(CONFIG.TIMEZONE)
    .create();
    
  Logger.log(`Установлен ежедневный триггер для ${functionName} в ${hour}:00`);
}

/**
 * Удаление всех триггеров функции
 */
function removeTrigger(functionName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === functionName)
    .forEach(t => ScriptApp.deleteTrigger(t));
    
  Logger.log(`Удалены все триггеры для ${functionName}`);
}

// ===== АГРЕГАЦИЯ И ГРУППИРОВКА ДАННЫХ =====

/**
 * Группировка данных по ключу
 */
function groupBy(array, keyFunction) {
  const groups = new Map();
  
  for (const item of array) {
    const key = keyFunction(item);
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(item);
  }
  
  return groups;
}

/**
 * Агрегация метрик по группам
 */
function aggregateMetrics(items, groupKeyFn, metrics = {}) {
  const groups = groupBy(items, groupKeyFn);
  const result = [];
  
  for (const [key, groupItems] of groups) {
    const aggregated = { key };
    
    // Стандартные метрики
    aggregated.count = groupItems.length;
    aggregated.sum = groupItems.reduce((s, item) => s + (metrics.sum ? metrics.sum(item) : 0), 0);
    aggregated.avg = aggregated.count > 0 ? aggregated.sum / aggregated.count : 0;
    
    // Кастомные метрики
    for (const [metricName, metricFn] of Object.entries(metrics)) {
      if (metricName !== 'sum') {
        aggregated[metricName] = metricFn(groupItems);
      }
    }
    
    result.push(aggregated);
  }
  
  return result.sort((a, b) => (b.count || 0) - (a.count || 0));
}

// ===== ЕДИНОЕ МЕНЮ ДЛЯ ВСЕХ ОТЧЁТОВ =====

/**
 * Создание единого меню для всех отчётов
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Основное меню
  const menu = ui.createMenu('📊 Все Отчёты');
  
  // Основные отчёты
  menu.addItem('🔄 Обновить РАБОЧИЙ АМО', 'buildWorkingFromFive')
      .addSeparator()
      .addItem('📉 Анализ причин отказов', 'dailyReportWithGPT')
      .addItem('📊 Сравнительный анализ каналов', 'buildChannelComparison')
      .addItem('🎯 Анализ первых касаний', 'buildFirstTouchReport')
      .addItem('📈 Сводная аналитика AMO', 'buildAmoDashboard_v2')
      .addSeparator();
  
  // Детальная аналитика
  menu.addItem('📞 Колл-трекинг', 'buildCalltrackingReport')
      .addItem('🌐 Лиды по каналам (детально)', 'buildChannelDetailedReports')
      .addItem('💰 Сквозная аналитика', 'buildE2E')
      .addItem('🎨 Анализ каналов привлечения', 'buildAcquisitionAnalysis')
      .addSeparator();
  
  // Операционные отчёты
  menu.addItem('📅 Ежедневная статистика', 'buildDailyStats')
      .addItem('👥 Эффективность команды', 'buildTeamPerformanceDashboard')
      .addItem('📢 Маркетинг и каналы', 'buildMarketingChannels')
      .addItem('📊 Динамика по месяцам', 'buildMonthlyDynamics')
      .addSeparator();
  
  // Специальные отчёты
  menu.addItem('🎯 Выгрузка для Яндекс.Директ', 'buildDirectConversions')
      .addItem('⭐ ТОП-40 клиентов', 'buildTop40ClientsFromGuests')
      .addItem('🌐 Аналитика сайта', 'buildSiteAnalytics')
      .addItem('🔗 Сквозная Сайт → AMO', 'buildSiteToAmoAnalytics')
      .addSeparator();
  
  // Управление
  const triggerMenu = ui.createMenu('⏰ Триггеры')
    .addItem('Установить все ежечасные', 'setupAllHourlyTriggers')
    .addItem('Удалить все триггеры', 'removeAllTriggers')
    .addSeparator()
    .addItem('Показать активные триггеры', 'showActiveTriggers');
  
  menu.addSubMenu(triggerMenu)
      .addSeparator()
      .addItem('🔄 Обновить все отчёты', 'updateAllReports')
      .addItem('ℹ️ О системе', 'showAbout')
      .addToUi();
}

// ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ МЕНЮ =====

/**
 * Установка всех ежечасных триггеров
 */
function setupAllHourlyTriggers() {
  const functions = [
    'buildWorkingFromFive',
    'buildAmoDashboard_v2', 
    'buildCalltrackingReport',
    'buildChannelDetailedReports',
    'buildE2E',
    'buildDailyStats'
  ];
  
  functions.forEach(fn => {
    try {
      setHourlyTrigger(fn);
    } catch (e) {
      Logger.log(`Ошибка установки триггера для ${fn}: ${e.toString()}`);
    }
  });
  
  SpreadsheetApp.getUi().alert(
    '✅ Установлены ежечасные триггеры',
    `Активировано триггеров: ${functions.length}\n${functions.join('\n')}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Удаление всех триггеров
 */
function removeAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  SpreadsheetApp.getUi().alert(
    '🗑️ Триггеры удалены',
    `Удалено триггеров: ${triggers.length}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Показ активных триггеров
 */
function showActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const info = triggers.map(t => 
    `• ${t.getHandlerFunction()} (${t.getEventType()})`
  ).join('\n');
  
  SpreadsheetApp.getUi().alert(
    '⏰ Активные триггеры',
    info || 'Нет активных триггеров',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Мастер-функция обновления всех отчётов
 */
function updateAllReports() {
  const ui = SpreadsheetApp.getUi();
  const start = new Date();
  
  const reports = [
    { name: 'РАБОЧИЙ АМО', fn: 'buildWorkingFromFive' },
    { name: 'Сводная аналитика', fn: 'buildAmoDashboard_v2' },
    { name: 'Колл-трекинг', fn: 'buildCalltrackingReport' },
    { name: 'Сравнительный анализ', fn: 'buildChannelComparison' },
    { name: 'Первые касания', fn: 'buildFirstTouchReport' }
  ];
  
  let success = 0;
  let errors = [];
  
  reports.forEach(report => {
    try {
      if (typeof this[report.fn] === 'function') {
        this[report.fn]();
        success++;
      } else {
        errors.push(`${report.name}: Функция не найдена`);
      }
    } catch (e) {
      errors.push(`${report.name}: ${e.toString()}`);
    }
  });
  
  const duration = Math.round((new Date() - start) / 1000);
  
  ui.alert(
    '📊 Обновление завершено',
    `Успешно: ${success}/${reports.length}\n` +
    `Время: ${duration} сек.\n` +
    (errors.length ? `\nОшибки:\n${errors.join('\n')}` : ''),
    ui.ButtonSet.OK
  );
}

/**
 * Информация о системе
 */
function showAbout() {
  SpreadsheetApp.getUi().alert(
    'ℹ️ О системе аналитики',
    'Версия: 2.0\n' +
    'Всего отчётов: 16+\n' +
    'Основной источник: РАБОЧИЙ АМО\n' +
    'GPT модель: gpt-4o\n' +
    'Шрифт: PT Sans\n\n' +
    '© 2025 Автоматизированная система аналитики',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ===== ЛОГИРОВАНИЕ И ОТЛАДКА =====

/**
 * Логирование в специальный лист
 */
function logToSheet(ss, message, level = 'INFO') {
  try {
    const logSheet = ss.getSheetByName('LOG') || ss.insertSheet('LOG');
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd.MM.yyyy HH:mm:ss');
    logSheet.appendRow([timestamp, level, message]);
    
    // Оставляем только последние 1000 записей
    const lastRow = logSheet.getLastRow();
    if (lastRow > 1000) {
      logSheet.deleteRows(2, lastRow - 1000);
    }
  } catch (e) {
    Logger.log('Ошибка записи в лог: ' + e.toString());
  }
}
