/**
 * ГЛАВНЫЙ СКРИПТ СИСТЕМЫ АНАЛИТИКИ
 * Центральная точка управления всеми модулями и функциями
 * @fileoverview Основной контроллер системы маркетинговой аналитики
 */

/**
 * Инициализация системы при первом запуске
 * Создаёт все необходимые листы, настраивает триггеры и меню
 */
function initializeSystem() {
  try {
    logInfo_('SYSTEM_INIT', 'Начало инициализации системы');
    
    // 1. Проверяем конфигурацию
    validateConfiguration_();
    
    // 2. Создаём все необходимые листы
    createAllSheets_();
    
    // 3. Применяем единое оформление
    applySystemWideFormatting_();
    
    // 4. Настраиваем триггеры
    setupTriggers_();
    
    // 5. Создаём меню
    createCustomMenu_();
    
    // 6. Выполняем первичную синхронизацию
    logInfo_('SYSTEM_INIT', 'Выполнение первичной синхронизации данных');
    syncAllData();
    
    // 7. Создаём главную страницу с навигацией
    createMainDashboard_();
    
    logInfo_('SYSTEM_INIT', 'Система успешно инициализирована');
    
    // Показываем сообщение пользователю (только если доступен UI)
    try {
      SpreadsheetApp.getUi().alert(
        'Система инициализирована!',
        'Все модули настроены и готовы к работе.\n\n' +
        'Автоматическая синхронизация данных настроена на каждые 15 минут.\n' +
        'Аналитические отчёты обновляются ежедневно в 08:00.\n\n' +
        'Используйте меню "🔄 Аналитика" для ручного управления.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      // UI недоступен (например, вызов из триггера) - просто логируем
      logInfo_('SYSTEM_INIT', 'UI недоступен, сообщение не показано');
    }
    
  } catch (error) {
    logError_('SYSTEM_INIT', 'Критическая ошибка инициализации системы', error);
    
    try {
      SpreadsheetApp.getUi().alert(
        'Ошибка инициализации!',
        `Произошла ошибка при инициализации системы:\n\n${error.message}\n\nПроверьте настройки конфигурации и попробуйте снова.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      // UI недоступен - логируем ошибку
      logError_('SYSTEM_INIT', 'UI недоступен для показа ошибки', uiError);
    }
    
    throw error;
  }
}

/**
 * Основная функция обновления всей аналитики
 * Вызывается ежедневно по расписанию
 */
function runFullAnalyticsUpdate() {
  try {
    logInfo_('FULL_UPDATE', 'Начало полного обновления аналитики');
    const startTime = new Date();
    
    // 1. Синхронизация данных
    logInfo_('FULL_UPDATE', 'Синхронизация данных из внешних источников');
    const syncResults = syncAllData();
    
    // 2. Обновление всех аналитических модулей
    const updateResults = {
      amocrm_summary: false,
      refusal_analysis: false,
      channel_analysis: false,
      lead_analysis: false,
      utm_analysis: false,
      first_touch: false,
      daily_stats: false,
      monthly_comparison: false,
      manager_performance: false,
      client_analysis: false,
      booking_analysis: false,
      beauty_analytics: false
    };
    
    // AmoCRM сводка
    try {
      logInfo_('FULL_UPDATE', 'Обновление сводной аналитики AmoCRM');
      updateAmoCrmSummary();
      updateResults.amocrm_summary = true;
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка обновления AmoCRM сводки', error);
    }
    
    // Анализ причин отказов
    try {
      logInfo_('FULL_UPDATE', 'Анализ причин отказов');
      analyzeRefusalReasons();
      updateResults.refusal_analysis = true;
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка анализа причин отказов', error);
    }
    
    // Анализ каналов
    try {
      logInfo_('FULL_UPDATE', 'Анализ эффективности каналов');
      analyzeChannelPerformance();
      updateResults.channel_analysis = true;
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка анализа каналов', error);
    }
    
    // Анализ лидов
    try {
      logInfo_('FULL_UPDATE', 'Анализ лидов по каналам');
      analyzeLeadsByChannels();
      updateResults.lead_analysis = true;
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка анализа лидов', error);
    }
    
    // UTM анализ
    try {
      logInfo_('FULL_UPDATE', 'UTM аналитика');
      analyzeUtmPerformance();
      updateResults.utm_analysis = true;
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка UTM аналитики', error);
    }
    
    // Анализ первых касаний
    try {
      logInfo_('FULL_UPDATE', 'Анализ первых касаний');
      analyzeFirstTouchAttribution();
      updateResults.first_touch = true;
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка анализа первых касаний', error);
    }
    
    // Ежедневная статистика
    try {
      logInfo_('FULL_UPDATE', 'Ежедневная статистика');
      updateDailyStatistics();
      updateResults.daily_stats = true;
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка ежедневной статистики', error);
    }
    
    // 3. Обновление главного дашборда
    try {
      logInfo_('FULL_UPDATE', 'Обновление главного дашборда');
      updateMainDashboard_(updateResults);
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка обновления главного дашборда', error);
    }
    
    // 4. Отправка отчётов
    try {
      logInfo_('FULL_UPDATE', 'Отправка email отчётов');
      sendDailyReports_(updateResults);
    } catch (error) {
      logError_('FULL_UPDATE', 'Ошибка отправки отчётов', error);
    }
    
    const duration = (new Date() - startTime) / 1000;
    const successCount = Object.values(updateResults).filter(Boolean).length;
    const totalModules = Object.keys(updateResults).length;
    
    logInfo_('FULL_UPDATE', `Полное обновление завершено за ${duration}с. Успешно: ${successCount}/${totalModules} модулей`);
    
    return {
      success: true,
      duration: duration,
      syncResults: syncResults,
      updateResults: updateResults,
      successCount: successCount,
      totalModules: totalModules
    };
    
  } catch (error) {
    logError_('FULL_UPDATE', 'Критическая ошибка полного обновления', error);
    throw error;
  }
}

/**
 * Создаёт пользовательское меню
 */
function createCustomMenu_() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.createMenu('🔄 Аналитика')
      .addSubMenu(ui.createMenu('📊 Обновление данных')
        .addItem('🔄 Синхронизация всех данных', 'syncAllData')
        .addItem('💼 Только AmoCRM', 'syncAmoCrmDataOnly')
        .addItem('📋 Только формы сайта', 'syncWebFormsDataOnly')
        .addItem('☎️ Только колл-трекинг', 'syncCallTrackingDataOnly')
        .addSeparator()
        .addItem('🎯 Яндекс.Метрика', 'syncYandexMetricaDataOnly'))
      
      .addSubMenu(ui.createMenu('📈 Аналитические отчёты')
        .addItem('🏢 Сводка AmoCRM', 'updateAmoCrmSummary')
        .addItem('❌ Анализ отказов', 'analyzeRefusalReasons')
        .addItem('🎯 Каналы привлечения', 'analyzeChannelPerformance')
        .addItem('👥 Лиды по каналам', 'analyzeLeadsByChannels')
        .addItem('🔗 UTM аналитика', 'analyzeUtmPerformance')
        .addItem('👋 Первые касания', 'analyzeFirstTouchAttribution')
        .addSeparator()
        .addItem('📅 Ежедневная статистика', 'updateDailyStatistics')
        .addItem('📊 Обновить все отчёты', 'runFullAnalyticsUpdate'))
      
      .addSubMenu(ui.createMenu('⚙️ Настройки')
        .addItem('🎨 Применить оформление', 'applySystemWideFormatting_')
        .addItem('✨ Красивое оформление РАБОЧИЙ АМО', 'applyBeautifulStyleToWorkingAmo')
        .addItem('🔄 Пересоздать триггеры', 'setupTriggers_')
        .addItem('📧 Тест email уведомлений', 'testEmailNotifications')
        .addSeparator()
        .addItem('🗑️ Очистить кэш', 'clearAllCache')
        .addItem('📋 Показать логи', 'showSystemLogs'))
      
      .addSeparator()
      .addItem('🚀 Инициализация системы', 'initializeSystem')
      .addItem('ℹ️ О системе', 'showSystemInfo')
      .addToUi();
      
    logInfo_('MENU', 'Пользовательское меню создано');
    
  } catch (uiError) {
    // UI недоступен (например, вызов из триггера) - пропускаем создание меню
    logInfo_('MENU', 'UI недоступен для создания меню - пропуск');
  }
}

/**
 * Проверяет конфигурацию системы
 * @private
 */
function validateConfiguration_() {
  const requiredProperties = [
    'OPENAI_API_KEY'
  ];
  
  const properties = PropertiesService.getScriptProperties();
  const missingProperties = [];
  
  requiredProperties.forEach(prop => {
    if (!properties.getProperty(prop)) {
      missingProperties.push(prop);
    }
  });
  
  if (missingProperties.length > 0) {
    throw new Error(`Не настроены обязательные параметры: ${missingProperties.join(', ')}`);
  }
  
  // Проверяем доступность Google Sheets API
  try {
    SpreadsheetApp.getActiveSpreadsheet().getId();
  } catch (error) {
    throw new Error('Нет доступа к Google Sheets API');
  }
  
  logInfo_('CONFIG_CHECK', 'Конфигурация системы корректна');
}

/**
 * Создаёт все необходимые листы
 * @private
 */
function createAllSheets_() {
  // Создаём основные листы данных
  const dataSheets = [
    'Амо Выгрузка',
    'Выгрузка Амо Полная', 
    'Заявки с Сайта',
    'Reserves RP',
    'Guests RP',
    'КоллТрекинг',
    'РАБОЧИЙ АМО'
  ];
  
  // Создаём аналитические листы
  const analyticsSheets = [
    'Главная',
    'Ежедневная статистика',
    'Причина отказов', 
    'СРАВНИТЕЛЬНЫЙ АНАЛИЗ',
    'FIRST-TOUCH ANALYSIS',
    'СВОДНАЯ АНАЛИТИКА AMOcrm',
    'Лиды по каналам',
    'UTM аналитика',
    'Анализ менеджеров',
    'КоллТрекинг',
    'Сравнение месяцев',
    'Клиентская аналитика',
    'Анализ броней',
    'Beauty Analytics'
  ];
  
  // Создаём служебные листы
  const serviceSheets = [
    'LOG',
    '_DIAG'
  ];
  
  const allSheets = [...dataSheets, ...analyticsSheets, ...serviceSheets];
  
  allSheets.forEach(sheetName => {
    try {
      getOrCreateSheet_(sheetName);
      logInfo_('SHEET_CREATE', `Лист "${sheetName}" создан или уже существует`);
    } catch (error) {
      logError_('SHEET_CREATE', `Ошибка создания листа "${sheetName}"`, error);
    }
  });
}

/**
 * Применяет единое оформление ко всем листам
 * @private
 */
function applySystemWideFormatting_() {
  logInfo_('FORMATTING', 'Применение единого оформления');
  
  // Применяем PT Sans ко всем листам
  applyPtSansAllSheets_();
  
  // Устанавливаем стандартную ширину столбцов
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  sheets.forEach(sheet => {
    try {
      // Автоматическая высота строк
      sheet.setRowHeights(1, sheet.getMaxRows(), 25);
      
      // Стандартная ширина столбцов
      sheet.setColumnWidths(1, sheet.getMaxColumns(), 120);
      
      // Заморозка первой строки если есть данные
      if (sheet.getLastRow() > 0) {
        sheet.setFrozenRows(1);
      }
      
    } catch (error) {
      logWarning_('FORMATTING', `Ошибка форматирования листа "${sheet.getName()}"`, error);
    }
  });
  
  logInfo_('FORMATTING', 'Единое оформление применено');
}

/**
 * Настраивает автоматические триггеры
 * @private
 */
function setupTriggers_() {
  try {
    logInfo_('TRIGGERS', 'Настройка автоматических триггеров');
    
    // Удаляем существующие триггеры
    const existingTriggers = ScriptApp.getProjectTriggers();
    existingTriggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
    });
    
    // Создаём новые триггеры
    
    // 1. Синхронизация данных каждые 15 минут
    ScriptApp.newTrigger('syncAllData')
      .timeBased()
      .everyMinutes(15)
      .create();
    
    // 2. Полное обновление аналитики каждый день в 08:00
    ScriptApp.newTrigger('runFullAnalyticsUpdate')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
    
    // 3. Еженедельный отчёт по понедельникам в 09:00
    ScriptApp.newTrigger('sendWeeklyReport')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(9)
      .create();
    
    // 4. Анализ причин отказов каждые 2 часа (для актуальности GPT анализа)
    ScriptApp.newTrigger('analyzeRefusalReasons')
      .timeBased()
      .everyHours(2)
      .create();
    
    logInfo_('TRIGGERS', 'Автоматические триггеры настроены');
    
  } catch (error) {
    logError_('TRIGGERS', 'Ошибка настройки триггеров', error);
    throw error;
  }
}

/**
 * Очищает данные на листе
 * @param {Sheet} sheet - Лист для очистки
 * @private
 */
function clearSheetData_(sheet) {
  if (!sheet) return;
  try {
    if (sheet.getLastRow() > 0) {
      sheet.clear();
    }
  } catch (error) {
    logError_('CLEAR_SHEET', `Ошибка очистки листа "${sheet.getName()}"`, error);
  }
}

/**
 * Применяет стиль заголовка к диапазону
 * @param {Range} range - Диапазон для применения стиля
 * @private
 */
function applyHeaderStyle_(range) {
  if (!range) return;
  try {
    range.setBackground(CONFIG.COLORS.HEADER_BG || '#4285f4')
         .setFontColor(CONFIG.COLORS.HEADER_TEXT || '#ffffff')
         .setFontWeight('bold')
         .setFontSize(14)
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
  } catch (error) {
    logError_('STYLE_HEADER', 'Ошибка применения стиля заголовка', error);
  }
}

/**
 * Применяет стиль подзаголовка к диапазону
 * @param {Range} range - Диапазон для применения стиля
 * @private
 */
function applySubheaderStyle_(range) {
  if (!range) return;
  try {
    range.setBackground(CONFIG.COLORS.SUBHEADER_BG || '#f1f3f4')
         .setFontColor(CONFIG.COLORS.SUBHEADER_TEXT || '#202124')
         .setFontWeight('bold')
         .setFontSize(12)
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
  } catch (error) {
    logError_('STYLE_SUBHEADER', 'Ошибка применения стиля подзаголовка', error);
  }
}

/**
 * Применяет стиль данных к диапазону
 * @param {Range} range - Диапазон для применения стиля
 * @private
 */
function applyDataStyle_(range) {
  if (!range) return;
  try {
    range.setBackground('#ffffff')
         .setFontColor('#202124')
         .setFontSize(11)
         .setVerticalAlignment('middle')
         .setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  } catch (error) {
    logError_('STYLE_DATA', 'Ошибка применения стиля данных', error);
  }
}

/**
 * Получает лист по имени
 * @param {string} sheetName - Имя листа
 * @return {Sheet|null} Лист или null если не найден
 * @private
 */
function getSheet_(sheetName) {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  } catch (error) {
    logError_('GET_SHEET', `Ошибка получения листа "${sheetName}"`, error);
    return null;
  }
}

/**
 * Получает данные с листа в виде объекта с заголовками и строками
 * @param {string|Sheet} sheetOrName - Имя листа или объект листа
 * @return {Object|Array} Объект с header и rows или массив данных (для совместимости)
 * @private
 */
function getSheetData_(sheetOrName) {
  try {
    let sheet;
    let returnFormat = 'object'; // По умолчанию возвращаем объект
    
    // Определяем тип параметра
    if (typeof sheetOrName === 'string') {
      // Передано имя листа
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetOrName);
      returnFormat = 'object';
    } else if (sheetOrName && typeof sheetOrName === 'object' && sheetOrName.getName) {
      // Передан объект Sheet
      sheet = sheetOrName;
      returnFormat = 'array'; // Для совместимости со старыми модулями
    } else {
      logError_('GET_SHEET_DATA', 'Неверный тип параметра - ожидается имя листа или объект Sheet');
      return returnFormat === 'object' ? { header: [], rows: [] } : [];
    }
    
    if (!sheet) {
      const sheetName = typeof sheetOrName === 'string' ? sheetOrName : 'неизвестный';
      logWarning_('GET_SHEET_DATA', `Лист "${sheetName}" не найден`);
      return returnFormat === 'object' ? { header: [], rows: [] } : [];
    }
    
    if (sheet.getLastRow() <= 1) {
      logWarning_('GET_SHEET_DATA', `Лист "${sheet.getName()}" пуст или содержит только заголовки`);
      return returnFormat === 'object' ? { header: [], rows: [] } : [];
    }
    
    const values = sheet.getDataRange().getValues();
    if (!values || values.length === 0) {
      return returnFormat === 'object' ? { header: [], rows: [] } : [];
    }
    
    const header = (values[0] || []).map(String);
    const rows = values.slice(1).filter(r => r.some(x => String(x).trim() !== ''));
    
    logInfo_('GET_SHEET_DATA', `Получено ${rows.length} строк из листа "${sheet.getName()}"`);
    
    // Возвращаем в нужном формате
    if (returnFormat === 'array') {
      // Для совместимости со старыми модулями - возвращаем массив
      return [header, ...rows];
    } else {
      // Новый формат - объект с header и rows
      return { header, rows };
    }
    
  } catch (error) {
    const sheetName = typeof sheetOrName === 'string' ? sheetOrName : 'неизвестный';
    logError_('GET_SHEET_DATA', `Ошибка получения данных с листа "${sheetName}"`, error);
    return returnFormat === 'object' ? { header: [], rows: [] } : [];
  }
}

/**
 * Создаёт главную страницу дашборда
 * @private
 */
function createMainDashboard_() {
  const sheet = getOrCreateSheet_('Главная');
  clearSheetData_(sheet);
  
  // Перемещаем лист в начало
  sheet.activate();
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(1);
  
  // Заголовок системы
  sheet.getRange('A1:H1').merge();
  sheet.getRange('A1').setValue('📊 СИСТЕМА МАРКЕТИНГОВОЙ АНАЛИТИКИ');
  sheet.getRange('A1').setFontSize(24).setFontWeight('bold').setHorizontalAlignment('center');
  applyHeaderStyle_(sheet.getRange('A1'));
  
  let currentRow = 3;
  
  // Информация о системе
  const systemInfo = [
    ['📈 Модуль', '📊 Статус', '🕐 Обновлено', '📝 Описание'],
    ['AmoCRM Сводка', 'Готов', '', 'Общие показатели по сделкам и конверсии'],
    ['Анализ отказов', 'Готов', '', 'AI-анализ причин отказов клиентов'],
    ['Каналы привлечения', 'Готов', '', 'Эффективность маркетинговых каналов'],
    ['Лиды по каналам', 'Готов', '', 'Детализация лидов по источникам'],
    ['UTM аналитика', 'Готов', '', 'Анализ UTM меток и кампаний'],
    ['Первые касания', 'Готов', '', 'Атрибуция первых касаний клиентов'],
    ['Ежедневная статистика', 'Готов', '', 'Операционные показатели по дням'],
    ['Колл-трекинг', 'Готов', '', 'Аналитика звонков и их источников'],
    ['Анализ менеджеров', 'Готов', '', 'Эффективность работы менеджеров']
  ];
  
  sheet.getRange(currentRow, 1, systemInfo.length, 4).setValues(systemInfo);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
  currentRow += systemInfo.length + 2;
  
  // Быстрые действия
  sheet.getRange(currentRow, 1, 1, 4).merge();
  sheet.getRange(currentRow, 1).setValue('⚡ БЫСТРЫЕ ДЕЙСТВИЯ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1));
  currentRow++;
  
  const quickActions = [
    ['🔄 Синхронизировать данные', '→ Меню: Аналитика → Синхронизация всех данных'],
    ['📊 Обновить все отчёты', '→ Меню: Аналитика → Обновить все отчёты'],
    ['❌ Анализ причин отказов', '→ Меню: Аналитика → Анализ отказов'],
    ['🎯 Эффективность каналов', '→ Меню: Аналитика → Каналы привлечения'],
    ['📧 Настройки уведомлений', '→ Меню: Аналитика → Настройки']
  ];
  
  sheet.getRange(currentRow, 1, quickActions.length, 2).setValues(quickActions);
  sheet.getRange(currentRow, 1, quickActions.length, 1).setFontWeight('bold');
  currentRow += quickActions.length + 2;
  
  // Системная информация
  sheet.getRange(currentRow, 1, 1, 4).merge();
  sheet.getRange(currentRow, 1).setValue('ℹ️ ИНФОРМАЦИЯ О СИСТЕМЕ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1));
  currentRow++;
  
  const sysInfo = [
    ['Версия системы:', '1.0.0'],
    ['Дата инициализации:', formatDate_(getCurrentDateMoscow_(), 'full')],
    ['Автосинхронизация:', 'Каждые 15 минут'],
    ['Обновление отчётов:', 'Ежедневно в 08:00'],
    ['Анализ GPT:', 'Каждые 2 часа'],
    ['Email отчёты:', 'Понедельники в 09:00']
  ];
  
  sheet.getRange(currentRow, 1, sysInfo.length, 2).setValues(sysInfo);
  sheet.getRange(currentRow, 1, sysInfo.length, 1).setFontWeight('bold');
  
  // Форматирование
  sheet.autoResizeColumns(1, 4);
  applyDataStyle_(sheet.getRange(3, 1, currentRow - 3 + sysInfo.length, 4));
}

/**
 * Обновляет главный дашборд с информацией о статусе модулей
 * @param {Object} updateResults - Результаты обновления модулей
 * @private
 */
function updateMainDashboard_(updateResults) {
  const sheet = getSheet_('Главная');
  if (!sheet) return;
  
  const now = formatDate_(getCurrentDateMoscow_(), 'DD.MM.YYYY HH:MM');
  
  // Обновляем статусы модулей (строки 4-13)
  Object.entries(updateResults).forEach((entry, index) => {
    const [module, success] = entry;
    const row = 4 + index;
    
    if (row <= 13) { // Защита от выхода за границы
      sheet.getRange(row, 2).setValue(success ? '✅ Активен' : '❌ Ошибка');
      sheet.getRange(row, 3).setValue(success ? now : 'Ошибка');
      
      // Цветовая индикация
      const statusCell = sheet.getRange(row, 2);
      statusCell.setBackground(success ? CONFIG.COLORS.SUCCESS_LIGHT : CONFIG.COLORS.ERROR_LIGHT);
    }
  });
}

/**
 * Отправляет ежедневные отчёты по email
 * @param {Object} updateResults - Результаты обновления
 * @private
 */
function sendDailyReports_(updateResults) {
  const recipients = CONFIG.EMAIL_NOTIFICATIONS.RECIPIENTS;
  if (!recipients || recipients.length === 0) return;
  
  const successCount = Object.values(updateResults).filter(Boolean).length;
  const totalModules = Object.keys(updateResults).length;
  const date = formatDate_(getCurrentDateMoscow_(), 'DD.MM.YYYY');
  
  const subject = `📊 Ежедневный отчёт аналитики за ${date}`;
  
  let body = `Отчёт о работе системы маркетинговой аналитики\n\n`;
  body += `📅 Дата: ${date}\n`;
  body += `✅ Успешно обновлено: ${successCount} из ${totalModules} модулей\n\n`;
  
  body += `Статус модулей:\n`;
  Object.entries(updateResults).forEach(([module, success]) => {
    const status = success ? '✅' : '❌';
    const moduleName = module.replace(/_/g, ' ').toUpperCase();
    body += `${status} ${moduleName}\n`;
  });
  
  if (successCount < totalModules) {
    body += `\n⚠️ Обнаружены проблемы в работе некоторых модулей. Проверьте настройки и логи системы.\n`;
  }
  
  body += `\n📊 Ссылка на отчёты: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}\n`;
  body += `\n---\nАвтоматический отчёт системы аналитики`;
  
  try {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });
    
    logInfo_('EMAIL_REPORT', `Ежедневный отчёт отправлен на ${recipients.length} адресов`);
    
  } catch (error) {
    logError_('EMAIL_REPORT', 'Ошибка отправки ежедневного отчёта', error);
  }
}

/**
 * Показывает информацию о системе
 */
function showSystemInfo() {
  const message = `
📊 СИСТЕМА МАРКЕТИНГОВОЙ АНАЛИТИКИ

🎯 Назначение:
Единая таблица маркетингово-продажной аналитики с автосбором данных, нормализацией, AI-разбором причин отказов и интерактивными графиками

📈 Функции:
• Автоматическая синхронизация данных из AmoCRM, форм сайта, колл-трекинга
• AI-анализ причин отказов с использованием GPT-4
• Интерактивные диаграммы «за всё время» и «по месяцам»
• 12 специализированных аналитических модулей
• Автоматические email отчёты

⚙️ Технические характеристики:
• 16 модулей Google Apps Script
• Интеграция с OpenAI GPT-4o-mini
• Московский часовой пояс
• PT Sans шрифт во всей системе
• Автообновление каждые 15 минут

📧 Поддержка:
Для вопросов и доработок обращайтесь к разработчику
  `;
  
  SpreadsheetApp.getUi().alert('О системе', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Очищает весь кэш системы
 */
function clearAllCache() {
  try {
    CacheService.getScriptCache().removeAll(['headers_cache', 'utm_cache', 'channels_cache']);
    CacheService.getDocumentCache().removeAll(['daily_stats', 'monthly_stats']);
    
    logInfo_('CACHE_CLEAR', 'Весь кэш системы очищен');
    
    SpreadsheetApp.getUi().alert(
      'Кэш очищен',
      'Кэш системы успешно очищен. При следующем обновлении данные будут пересчитаны заново.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    logError_('CACHE_CLEAR', 'Ошибка очистки кэша', error);
  }
}

/**
 * Показывает системные логи
 */
function showSystemLogs() {
  // Получаем последние логи из Properties
  const properties = PropertiesService.getScriptProperties();
  const logsJson = properties.getProperty('SYSTEM_LOGS') || '[]';
  
  try {
    const logs = JSON.parse(logsJson);
    const recentLogs = logs.slice(-50); // Последние 50 записей
    
    let message = 'ПОСЛЕДНИЕ ЗАПИСИ ЛОГА:\n\n';
    
    if (recentLogs.length === 0) {
      message += 'Логи пусты или не найдены.';
    } else {
      recentLogs.forEach(log => {
        message += `[${log.timestamp}] ${log.level} ${log.module}: ${log.message}\n`;
      });
    }
    
    // Показываем в HTML диалоге для лучшего отображения
    const htmlOutput = HtmlService
      .createHtmlOutput(`<pre style="font-family: monospace; font-size: 12px; white-space: pre-wrap;">${message}</pre>`)
      .setWidth(800)
      .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Системные логи');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Ошибка',
      'Не удалось загрузить системные логи: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Тестирует отправку email уведомлений
 */
function testEmailNotifications() {
  try {
    const recipients = CONFIG.EMAIL_NOTIFICATIONS.RECIPIENTS;
    
    if (!recipients || recipients.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Настройка email',
        'Email адреса для уведомлений не настроены. Добавьте их в CONFIG.EMAIL_NOTIFICATIONS.RECIPIENTS',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const subject = '🧪 Тест уведомлений системы аналитики';
    const body = `Это тестовое сообщение от системы маркетинговой аналитики.\n\nВремя отправки: ${formatDate_(getCurrentDateMoscow_(), 'full')}\n\nЕсли вы получили это сообщение, настройка email уведомлений работает корректно.`;
    
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      body: body
    });
    
    SpreadsheetApp.getUi().alert(
      'Тест отправлен',
      `Тестовое письмо отправлено на ${recipients.length} адресов:\n${recipients.join('\n')}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Ошибка отправки',
      `Не удалось отправить тестовое письмо:\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== ФУНКЦИИ ОБРАБОТКИ ДАННЫХ =====

/**
 * Чтение таблицы с универсальными утилитами
 */
function readTable(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    logWarning_('READ_TABLE', `Лист "${sheetName}" не найден`);
    return { header: [], rows: [] };
  }
  
  const values = sh.getDataRange().getValues();
  if (!values || values.length === 0) return { header: [], rows: [] };
  
  const header = (values[0] || []).map(String);
  const rows = values.slice(1).filter(r => r.some(x => String(x).trim() !== ''));
  
  return { header, rows };
}

/**
 * Канонизация заголовков по словарю
 */
function canonHeaders(table, mapping) {
  const header = [...table.header];
  
  // Ищем синонимы и заменяем на канонические названия
  for (const [canonical, synonyms] of Object.entries(mapping)) {
    const idx = findColumnIndex(header, synonyms);
    if (idx > -1) {
      header[idx] = canonical;
    }
  }
  
  return { header, rows: table.rows };
}

/**
 * Объединение таблиц по ключу
 */
function mergeByKey(tables, keyField) {
  if (tables.length === 0) return { header: [], rows: [] };
  
  const masterTable = tables[0];
  const keyIdx = masterTable.header.indexOf(keyField);
  if (keyIdx === -1) throw new Error(`Ключевое поле "${keyField}" не найдено`);
  
  // Создаём индекс по ключу для первой таблицы
  const keyToRowIndex = new Map();
  masterTable.rows.forEach((row, idx) => {
    const key = String(row[keyIdx] || '').trim();
    if (key) keyToRowIndex.set(key, idx);
  });
  
  // Объединяем заголовки всех таблиц
  let mergedHeader = [...masterTable.header];
  const tableCols = [masterTable.header.length];
  
  for (let t = 1; t < tables.length; t++) {
    const table = tables[t];
    const startCol = mergedHeader.length;
    
    table.header.forEach(h => {
      if (!mergedHeader.includes(h)) {
        mergedHeader.push(h);
      }
    });
    
    tableCols.push(mergedHeader.length);
  }
  
  // Создаём объединённые строки
  const mergedRows = masterTable.rows.map(row => {
    const newRow = new Array(mergedHeader.length).fill('');
    
    // Копируем данные из первой таблицы
    for (let i = 0; i < masterTable.header.length; i++) {
      newRow[i] = row[i];
    }
    
    return newRow;
  });
  
  return { header: mergedHeader, rows: mergedRows };
}

/**
 * Построение агрегатов по телефонам
 */
function buildAggregates(rows, phoneIdx) {
  const aggregates = new Map();
  
  if (phoneIdx === -1) return aggregates;
  
  rows.forEach(row => {
    const phone = String(row[phoneIdx] || '').trim();
    if (phone) {
      if (!aggregates.has(phone)) {
        aggregates.set(phone, []);
      }
      aggregates.get(phone).push(row);
    }
  });
  
  return aggregates;
}

/**
 * Построение карты колл-трекинга
 */
function buildCalltrackingMap(callTable) {
  const ctMap = new Map();
  
  if (!callTable.rows || callTable.rows.length === 0) return ctMap;
  
  const phoneIdx = findColumnIndex(callTable.header, ['Телефон', 'Phone']);
  const sourceIdx = findColumnIndex(callTable.header, ['Источник', 'Source']);
  
  if (phoneIdx === -1) return ctMap;
  
  callTable.rows.forEach(row => {
    const phone = String(row[phoneIdx] || '').trim();
    const source = String(row[sourceIdx] || '').trim();
    if (phone && source) {
      ctMap.set(phone, source);
    }
  });
  
  return ctMap;
}

/**
 * Обогащение данных
 */
function buildEnrichedData(canonized, siteTable, resAgg, gueAgg, ctMap, CFG) {
  const headerOrderedRaw = [...canonized.header];
  const enrichedRows = canonized.rows.map(row => [...row]);
  
  return {
    headerOrderedRaw,
    rows: enrichedRows
  };
}

/**
 * Рендеринг в рабочий лист
 */
function renderToWorkingSheet(ss, CFG, header, rows) {
  const outputSheet = getOrCreateSheet_('РАБОЧИЙ АМО');
  
  // Очищаем лист
  outputSheet.clear();
  
  // Записываем заголовки
  if (header.length > 0) {
    outputSheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
  
  // Записываем данные
  if (rows.length > 0 && header.length > 0) {
    outputSheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  
  logInfo_('RENDER', `Записано ${rows.length} строк в лист "РАБОЧИЙ АМО"`);
}

/**
 * Обновление времени в колонке TIME
 */
function updateTimeOnlyOnWorking() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const workingSheet = ss.getSheetByName('РАБОЧИЙ АМО');
    
    if (!workingSheet) {
      logWarning_('TIME_UPDATE', 'Рабочий лист не найден');
      return;
    }
    
    const lastRow = workingSheet.getLastRow();
    if (lastRow <= 1) return;
    
    const header = workingSheet.getRange(1, 1, 1, workingSheet.getLastColumn()).getValues()[0];
    const timeIdx = findColumnIndex(header, ['TIME', 'Время', 'Время обработки']);
    
    if (timeIdx > -1) {
      const currentTime = getCurrentDateMoscow_().toLocaleString();
      const timeRange = workingSheet.getRange(2, timeIdx + 1, lastRow - 1, 1);
      const timeValues = Array(lastRow - 1).fill([currentTime]);
      timeRange.setValues(timeValues);
      
      logInfo_('TIME_UPDATE', `Обновлено время в ${lastRow - 1} строках`);
    }
  } catch (error) {
    logError_('TIME_UPDATE', 'Ошибка обновления времени', error);
  }
}

/**
 * Обновление источников колл-трекинга в рабочем листе
 */
function updateCalltrackingOnWorking() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const workingSheet = ss.getSheetByName('РАБОЧИЙ АМО');
    const callSheet = ss.getSheetByName('КоллТрекинг');
    
    if (!workingSheet || !callSheet) {
      logWarning_('CT_UPDATE', 'Рабочий лист или лист колл-трекинга не найден');
      return;
    }
    
    const lastRow = workingSheet.getLastRow();
    if (lastRow <= 1) return;
    
    // Получаем карту колл-трекинга
    const callTable = readTable(ss, 'КоллТрекинг');
    const ctMap = buildCalltrackingMap(callTable);
    
    if (ctMap.size === 0) return;
    
    const header = workingSheet.getRange(1, 1, 1, workingSheet.getLastColumn()).getValues()[0];
    const phoneIdx = findColumnIndex(header, ['Телефон', 'Phone']);
    const sourceIdx = findColumnIndex(header, ['Источник ТЕЛ', 'Call Source']);
    
    if (phoneIdx === -1 || sourceIdx === -1) return;
    
    const data = workingSheet.getRange(2, 1, lastRow - 1, workingSheet.getLastColumn()).getValues();
    let updated = 0;
    
    data.forEach((row, idx) => {
      const phone = cleanPhone(row[phoneIdx]);
      if (phone && ctMap.has(phone)) {
        row[sourceIdx] = ctMap.get(phone);
        updated++;
      }
    });
    
    if (updated > 0) {
      workingSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
      logInfo_('CT_UPDATE', `Обновлено ${updated} источников колл-трекинга`);
    }
    
  } catch (error) {
    logError_('CT_UPDATE', 'Ошибка обновления колл-трекинга', error);
  }
}

/**
 * Очистка телефона для сопоставления
 */
function cleanPhone(phone) {
  if (!phone) return '';
  return String(phone)
    .replace(/\D/g, '')
    .replace(/^[78]/, '')
    .slice(0, 10);
}

/**
 * Гуманизация заголовков для отображения
 */
function humanizeHeader(header) {
  const humanMap = {
    'Сделка.ID': 'ID',
    'Сделка.Название': 'Название',
    'Сделка.Статус': 'Статус', 
    'Сделка.Бюджет': 'Бюджет',
    'Сделка.Дата создания': 'Дата создания',
    'Контакт.Телефон': 'Телефон',
    'Контакт.ФИО': 'Контакт',
    'Сделка.utm_source': 'UTM Source',
    'Сделка.utm_medium': 'UTM Medium',
    'Сделка.utm_campaign': 'UTM Campaign',
    'R.Источник ТЕЛ сделки': 'Источник ТЕЛ',
    'Reserves.Визиты': 'Res.Визиты',
    'Reserves.Сумма': 'Res.Сумма',
    'Guests.Визиты': 'Gue.Визиты',
    'Guests.Общая сумма': 'Gue.Сумма'
  };
  
  return humanMap[header] || header.replace(/^(Сделка|Контакт)\./, '');
}

/**
 * Заглушка для обновления ежедневной статистики
 * Будет заменена на полную функцию из модуля 12_Ежедневная_Статистика.gs
 */
function updateDailyStatistics() {
  try {
    logInfo_('DAILY_STATS', 'Начало обновления ежедневной статистики');
    
    // Получаем данные из рабочего листа
    const workingData = getSheetData_('РАБОЧИЙ АМО');
    if (workingData.rows.length === 0) {
      logWarning_('DAILY_STATS', 'Нет данных для анализа');
      return;
    }
    
    // Создаём базовый отчёт ежедневной статистики
    const dailySheet = getOrCreateSheet_('Ежедневная статистика');
    clearSheetData_(dailySheet);
    
    // Заголовки
    const headers = ['Дата', 'Всего сделок', 'Новые лиды', 'Конверсии', 'Среднее время обработки'];
    dailySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    applyHeaderStyle_(dailySheet.getRange(1, 1, 1, headers.length));
    
    // Получаем статистику по дням
    const dailyStats = analyzeDailyStats_(workingData);
    if (dailyStats.length > 0) {
      dailySheet.getRange(2, 1, dailyStats.length, headers.length).setValues(dailyStats);
      applyDataStyle_(dailySheet.getRange(2, 1, dailyStats.length, headers.length));
    }
    
    // Автоширина столбцов
    dailySheet.autoResizeColumns(1, headers.length);
    
    logInfo_('DAILY_STATS', `Ежедневная статистика обновлена: ${dailyStats.length} записей`);
    
  } catch (error) {
    logError_('DAILY_STATS', 'Ошибка обновления ежедневной статистики', error);
    throw error;
  }
}

/**
 * Анализирует данные по дням
 * @param {Object} workingData - Данные из рабочего листа
 * @return {Array} Массив статистики по дням
 * @private
 */
function analyzeDailyStats_(workingData) {
  const dailyMap = new Map();
  
  // Найдём индекс колонки с датой
  const dateIdx = findColumnIndex(workingData.header, ['Дата создания', 'Сделка.Дата создания', 'Date']);
  if (dateIdx === -1) {
    logWarning_('DAILY_STATS', 'Колонка с датой не найдена');
    return [];
  }
  
  // Группируем по дням
  workingData.rows.forEach(row => {
    const dateValue = row[dateIdx];
    if (!dateValue) return;
    
    let dateStr;
    if (dateValue instanceof Date) {
      dateStr = formatDate_(dateValue, 'DD.MM.YYYY');
    } else {
      const parsedDate = new Date(dateValue);
      if (isNaN(parsedDate)) return;
      dateStr = formatDate_(parsedDate, 'DD.MM.YYYY');
    }
    
    if (!dailyMap.has(dateStr)) {
      dailyMap.set(dateStr, {
        total: 0,
        new_leads: 0,
        conversions: 0
      });
    }
    
    const stats = dailyMap.get(dateStr);
    stats.total++;
    
    // Определяем новые лиды (упрощённая логика)
    const status = String(row[findColumnIndex(workingData.header, ['Статус', 'Сделка.Статус'])] || '');
    if (status.includes('Новый') || status.includes('новый') || status.includes('NEW')) {
      stats.new_leads++;
    }
    
    // Определяем конверсии (упрощённая логика)
    if (status.includes('Успешно') || status.includes('успешно') || status.includes('Закрыт') || status.includes('SUCCESS')) {
      stats.conversions++;
    }
  });
  
  // Преобразуем в массив для таблицы
  const result = [];
  for (const [date, stats] of dailyMap.entries()) {
    result.push([
      date,
      stats.total,
      stats.new_leads,
      stats.conversions,
      '~5 мин' // Заглушка для времени обработки
    ]);
  }
  
  // Сортируем по дате (новые сверху)
  result.sort((a, b) => {
    const dateA = new Date(a[0].split('.').reverse().join('-'));
    const dateB = new Date(b[0].split('.').reverse().join('-'));
    return dateB - dateA;
  });
  
  return result;
}
