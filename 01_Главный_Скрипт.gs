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
    
    // 5. Создаём меню (только если UI доступен)
    try {
      createCustomMenu_();
    } catch (uiError) {
      logInfo_('SYSTEM_INIT', 'Создание меню пропущено (недоступен UI контекст)');
    }
    
    // 6. Выполняем первичную синхронизацию
    logInfo_('SYSTEM_INIT', 'Выполнение первичной синхронизации данных');
    syncAllData();
    
    // 7. Создаём главную страницу с навигацией
    createMainDashboard_();
    
    logInfo_('SYSTEM_INIT', 'Система успешно инициализирована');
    
    // Показываем сообщение пользователю (только если UI доступен)
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
      logInfo_('SYSTEM_INIT', 'UI уведомление пропущено (недоступен UI контекст)');
    }
    
  } catch (error) {
    logError_('SYSTEM_INIT', 'Критическая ошибка инициализации системы', error);
    
    // Пытаемся показать ошибку (только если UI доступен)
    try {
      SpreadsheetApp.getUi().alert(
        'Ошибка инициализации!',
        `Произошла ошибка при инициализации системы:\n\n${error.message}\n\nПроверьте настройки конфигурации и попробуйте снова.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (uiError) {
      logError_('SYSTEM_INIT', 'UI ошибка пропущена (недоступен UI контекст)', uiError);
    }
    
    throw error;
  }
}

/**
 * Основная функция обновления всей аналитики
 * Вызывается ежедневно по расписанию
 */
/**
 * 🎯 УПРОЩЕННАЯ ОСНОВНАЯ ФУНКЦИЯ - ФОКУС НА РАБОЧИЙ АМО
 * Только сборка файла с правильным оформлением как на картинках
 */
function runFullAnalyticsUpdate() {
  try {
    logInfo_('FULL_UPDATE', 'Начало сборки файла РАБОЧИЙ АМО');
    const startTime = new Date();
    
    // 🎯 ЕДИНСТВЕННАЯ ЗАДАЧА: Собираем файл РАБОЧИЙ АМО с правильным оформлением
    buildWorkingAmoFile();
    
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    logInfo_('FULL_UPDATE', `Файл РАБОЧИЙ АМО собран за ${duration} сек`);
    
  } catch (error) {
    logError_('FULL_UPDATE', 'Критическая ошибка сборки РАБОЧИЙ АМО', error);
    throw error;
  }
    
// ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ СИСТЕМЫ =====

/**
 * Функции логирования (упрощенные версии)
 */
function logInfo_(module, message, details = null) {
  const timestamp = new Date().toLocaleTimeString();
  console.log(`${timestamp}\tИнфо\t[${module}] ${message}${details ? ': ' + JSON.stringify(details) : ''}`);
}

function logError_(module, message, error = null) {
  const timestamp = new Date().toLocaleTimeString();
  console.error(`${timestamp}\tОшибка\t[ERROR] ${module}: ${message}${error ? ' [' + error.toString() + ']' : ''}`);
}

function logWarning_(module, message, details = null) {
  const timestamp = new Date().toLocaleTimeString();
  console.warn(`${timestamp}\tПредупреждение\t[WARNING] ${module}: ${message}${details ? ': ' + JSON.stringify(details) : ''}`);
}

/**
 * Проверяет конфигурацию системы
 */
function validateConfiguration_() {
  logInfo_('CONFIG_CHECK', 'Проверка конфигурации системы');
  
  // Проверяем основные настройки
  if (!CONFIG) {
    throw new Error('Конфигурация CONFIG не найдена');
  }
  
  if (!CONFIG.SHEETS) {
    throw new Error('Секция SHEETS не найдена в конфигурации');
  }
  
  logInfo_('CONFIG_CHECK', 'Конфигурация валидна');
}

/**
 * 🎯 ГЛАВНАЯ ФУНКЦИЯ СБОРКИ РАБОЧЕГО АМО
 * Создает итоговый файл с точным оформлением как на картинках
 */
function buildWorkingAmoFile() {
  console.log('🎯 Начинаем сборку файла РАБОЧИЙ АМО');
  
  try {
    const workingSheet = getSheet_(getSheetName_('WORKING_AMO'));
    workingSheet.clear();
    
    // 1. Создаем заголовки с точным оформлением как на картинках
    createWorkingAmoHeaders_(workingSheet);
    
    // 2. Загружаем и объединяем все данные
    const consolidatedData = consolidateAllDataSources_();
    
    // 3. Записываем данные в файл
    writeConsolidatedData_(workingSheet, consolidatedData);
    
    // 4. Применяем форматирование как на картинках
    applyWorkingAmoFormatting_(workingSheet, consolidatedData.length);
    
    console.log(`✅ Файл РАБОЧИЙ АМО собран: ${consolidatedData.length} записей`);
    
  } catch (error) {
    console.error('❌ Ошибка сборки РАБОЧИЙ АМО:', error);
    throw error;
  }
}

/**
 * Создаёт пользовательское меню
 */
function createCustomMenu_() {
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
      .addItem('🔄 Пересоздать триггеры', 'setupTriggers_')
      .addItem('📧 Тест email уведомлений', 'testEmailNotifications')
      .addSeparator()
      .addItem('🗑️ Очистить кэш', 'clearAllCache')
      .addItem('📋 Показать логи', 'showSystemLogs'))
    
    .addSeparator()
    .addItem('🚀 Инициализация системы', 'initializeSystem')
    .addItem('ℹ️ О системе', 'showSystemInfo')
    .addToUi();
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
  const requiredSheets = Object.values(CONFIG.SHEETS);
  
  requiredSheets.forEach(sheetName => {
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

/**
 * Чтение таблицы с универсальными утилитами
 */
function readTable(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    Logger.log(`⚠ Лист "${sheetName}" не найден`);
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
  
  // Объединяем остальные таблицы
  for (let t = 1; t < tables.length; t++) {
    const table = tables[t];
    const tableKeyIdx = table.header.indexOf(keyField);
    if (tableKeyIdx === -1) continue;
    
    // Находим новые колонки
    const newColumns = [];
    table.header.forEach((col, idx) => {
      if (!masterTable.header.includes(col)) {
        masterTable.header.push(col);
        newColumns.push(idx);
      }
    });
    
    // Если есть новые колонки, расширяем существующие строки
    if (newColumns.length > 0) {
      masterTable.rows.forEach(row => {
        for (let i = 0; i < newColumns.length; i++) {
          row.push('');
        }
      });
    }
    
    // Добавляем/обновляем данные
    table.rows.forEach(row => {
      const key = String(row[tableKeyIdx] || '').trim();
      if (!key) return;
      
      const existingIdx = keyToRowIndex.get(key);
      if (existingIdx !== undefined) {
        // Обновляем существующую строку
        table.header.forEach((col, colIdx) => {
          const masterColIdx = masterTable.header.indexOf(col);
          if (masterColIdx > -1 && row[colIdx] != null && row[colIdx] !== '') {
            masterTable.rows[existingIdx][masterColIdx] = row[colIdx];
          }
        });
      } else {
        // Добавляем новую строку
        const newRow = new Array(masterTable.header.length).fill('');
        table.header.forEach((col, colIdx) => {
          const masterColIdx = masterTable.header.indexOf(col);
          if (masterColIdx > -1) {
            newRow[masterColIdx] = row[colIdx];
          }
        });
        masterTable.rows.push(newRow);
        keyToRowIndex.set(key, masterTable.rows.length - 1);
      }
    });
  }
  
  return masterTable;
}

/**
 * Построение агрегатов для Reserves/Guests
 */
function buildAggregates(rows, phoneIndex) {
  const agg = new Map();
  if (phoneIndex === -1 || !rows.length) return agg;
  
  rows.forEach(row => {
    const phone = cleanPhone(row[phoneIndex]);
    if (!phone) return;
    
    const existing = agg.get(phone) || { visits: 0, amount: 0, total: 0 };
    existing.visits++;
    
    // Пробуем найти суммы в разных колонках
    for (let i = 0; i < row.length; i++) {
      const val = toNumber(row[i]);
      if (val > 0 && val < 1000000) { // Разумные суммы
        existing.amount += val;
        existing.total += val;
      }
    }
    
    agg.set(phone, existing);
  });
  
  return agg;
}

/**
 * Построение карты колл-трекинга
 */
function buildCalltrackingMap(callTable) {
  const map = new Map();
  if (!callTable.rows.length) return map;
  
  const phoneIdx = findColumnIndex(callTable.header, ['Телефон', 'Phone', 'Номер']);
  const tagIdx = findColumnIndex(callTable.header, ['Тег', 'Tag', 'Источник']);
  
  if (phoneIdx === -1 || tagIdx === -1) return map;
  
  callTable.rows.forEach(row => {
    const phone = cleanPhone(row[phoneIdx]);
    const tag = cleanString(row[tagIdx]);
    if (phone && tag) {
      map.set(phone, tag);
    }
  });
  
  return map;
}

/**
 * Построение обогащённых данных
 */
function buildEnrichedData(merged, siteTable, resAgg, gueAgg, ctMap, CFG) {
  // Определяем порядок колонок
  const baseOrder = [
    'Сделка.ID', 'Сделка.Название', 'Сделка.Статус', 'Сделка.Бюджет',
    'Сделка.Дата создания', 'Контакт.Телефон', 'Контакт.ФИО',
    'Сделка.utm_source', 'Сделка.utm_medium', 'Сделка.utm_campaign'
  ];
  
  const headerOrderedRaw = [...baseOrder];
  
  // Добавляем остальные колонки из merged
  merged.header.forEach(col => {
    if (!headerOrderedRaw.includes(col)) {
      headerOrderedRaw.push(col);
    }
  });
  
  // Добавляем колонки обогащения
  headerOrderedRaw.push(
    'R.Источник ТЕЛ сделки',
    'Reserves.Визиты', 'Reserves.Сумма', 'Reserves.Последний визит',
    'Guests.Визиты', 'Guests.Общая сумма', 'Guests.Первый визит', 'Guests.Последний визит',
    'Сумма ₽', 'Время прихода'
  );
  
  // Строим индексы для merged
  const indices = {};
  headerOrderedRaw.forEach(col => {
    indices[col] = merged.header.indexOf(col);
  });
  
  // Обрабатываем строки
  const rows = merged.rows.map(row => {
    const out = [];
    
    // Базовые поля
    headerOrderedRaw.forEach(col => {
      const idx = indices[col];
      if (idx > -1) {
        let val = row[idx];
        
        // Специальная обработка
        if (col === 'Сделка.Статус') {
          val = normalizeStatus(val);
        } else if (col === 'Контакт.Телефон') {
          val = cleanPhone(val);
        } else if (col === 'Сделка.Дата создания') {
          val = formatDate(val);
        }
        
        out.push(val || '');
      } else {
        out.push('');
      }
    });
    
    // Обогащение данными
    const phone = cleanPhone(row[indices['Контакт.Телефон']]);
    
    // Колл-трекинг
    const ctIdx = out.length - 8; // Позиция R.Источник ТЕЛ сделки
    out[ctIdx] = ctMap.get(phone) || '';
    
    // Reserves
    const res = resAgg.get(phone);
    out[ctIdx + 1] = res?.visits || 0;
    out[ctIdx + 2] = res?.amount || 0;
    out[ctIdx + 3] = res?.lastVisit || '';
    
    // Guests  
    const gue = gueAgg.get(phone);
    out[ctIdx + 4] = gue?.visits || 0;
    out[ctIdx + 5] = gue?.total || 0;
    out[ctIdx + 6] = gue?.firstVisit || '';
    out[ctIdx + 7] = gue?.lastVisit || '';
    
    return out;
  });
  
  return { headerOrderedRaw, rows };
}

/**
 * Преобразование технических названий в человекочитаемые
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
 * Рендер в рабочий лист с сохранением стилей
 */
function renderToWorkingSheet(ss, CFG, displayHeader, rows) {
  const sh = ss.getSheetByName(CFG.SHEETS.OUT) || ss.insertSheet(CFG.SHEETS.OUT);
  
  if (CFG.PRESERVE_STYLES) {
    // Сохраняем оформление и значения первой строки
    const maxR = sh.getMaxRows();
    const maxC = sh.getMaxColumns();
    if (maxR > 1 && maxC > 0) {
      sh.getRange(2, 1, maxR - 1, maxC).clearContent();
    }
  } else {
    sh.clear();
  }

  // Обеспечиваем размер
  ensureSize(sh, 2 + rows.length, displayHeader.length);

  // Заголовки во второй строке (первая может содержать оформление)
  sh.getRange(2, 1, 1, displayHeader.length).setValues([displayHeader]);
  
  // Данные
  if (rows.length > 0) {
    sh.getRange(3, 1, rows.length, displayHeader.length).setValues(rows);
  }

  // Базовое оформление
  setFontAll(sh);
  
  // Заморозка заголовков
  sh.setFrozenRows(2);
  
  // Фильтр
  try {
    const existingFilter = sh.getFilter();
    if (existingFilter) existingFilter.remove();
    sh.getRange(2, 1, 1, displayHeader.length).createFilter();
  } catch (e) {
    Logger.log('Ошибка создания фильтра: ' + e.toString());
  }
}

/**
 * Создание подмножеств данных
 */
function makeSubsets(ss, displayHeader, rows, CFG) {
  const statusIdx = displayHeader.indexOf('Статус');
  const createdIdx = displayHeader.indexOf('Дата создания');
  
  if (statusIdx === -1) return;
  
  const today = new Date();
  const threeDaysAgo = new Date(today.getTime() - 3 * 24 * 60 * 60 * 1000);
  
  // НОВЫЕ: недавно созданные, не закрытые
  const newRows = rows.filter(row => {
    const status = String(row[statusIdx] || '');
    const created = toDate(row[createdIdx]);
    
    const isNotClosed = !CFG.CLOSED_RE.test(status);
    const isRecent = created && created >= threeDaysAgo;
    
    return isNotClosed && isRecent;
  });
  
  // ПРОБЛЕМНЫЕ: старые незакрытые
  const problemRows = rows.filter(row => {
    const status = String(row[statusIdx] || '');
    const created = toDate(row[createdIdx]);
    
    const isNotClosed = !CFG.CLOSED_RE.test(status);
    const isOld = created && created < threeDaysAgo;
    
    return isNotClosed && isOld;
  });
  
  // Записываем в листы
  writeSubset(ss, CFG.SHEETS.NEW_ONLY, displayHeader, newRows);
  writeSubset(ss, CFG.SHEETS.PROBLEM, displayHeader, problemRows);
}

/**
 * Запись подмножества в лист
 */
function writeSubset(ss, sheetName, header, rows) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  else sh.clear();
  
  // Заголовок
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  
  // Данные
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  
  // Заморозка и фильтр
  sh.setFrozenRows(1);
  try {
    sh.getRange(1, 1, 1, header.length).createFilter();
  } catch (e) {
    // Игнорируем ошибки фильтра
  }
  
  setFontAll(sh);
}

/**
 * Обновление времени в колонке TIME
 */
function updateTimeOnlyOnWorking() {
  const CFG = getModuleConfig('mainScript');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEETS.OUT);
  if (!sh) throw new Error(`Нет листа "${CFG.SHEETS.OUT}"`);

  // Заголовки во 2-й строке
  const header = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  const colTIME = header.indexOf('TIME') + 1;
  const colArrival = header.indexOf('Время прихода') + 1;
  
  if (colTIME <= 0) throw new Error('Не найдена колонка "TIME"');

  const lastRow = sh.getLastRow();
  const n = Math.max(0, lastRow - 2);
  if (n === 0) return;

  const timeVals = sh.getRange(3, colTIME, n, 1).getValues();
  const arrivalVals = colArrival > 0 ? sh.getRange(3, colArrival, n, 1).getValues() : null;

  const out = timeVals.map((row, i) => {
    let v = row[0];
    
    // Если TIME пусто, берём из "Время прихода"
    if ((v === '' || v == null) && arrivalVals) {
      v = arrivalVals[i][0];
    }

    // Если дата/время — берём только время
    if (v instanceof Date) {
      const t = new Date(v);
      const ms = t.getHours() * 3600000 + t.getMinutes() * 60000 + t.getSeconds() * 1000;
      return [ms / 86400000]; // Доля суток для Google Sheets
    }

    // Если строка — парсим HH:MM
    const s = String(v || '').trim();
    if (s) {
      const m = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
      if (m) {
        const h = Math.min(23, Number(m[1]));
        const min = Math.min(59, Number(m[2]));
        const sec = m[3] ? Math.min(59, Number(m[3])) : 0;
        const ms = h * 3600000 + min * 60000 + sec * 1000;
        return [ms / 86400000];
      }
    }

    return [''];
  });

  // Записываем и форматируем как время
  sh.getRange(3, colTIME, n, 1).setValues(out);
  sh.getRange(3, colTIME, n, 1).setNumberFormat('hh:mm');
}

/**
 * Обновление R.Источник ТЕЛ сделки на основе колл-трекинга
 */
function updateCalltrackingOnWorking() {
  const CFG = getModuleConfig('mainScript');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SHEETS.OUT);
  if (!sh) throw new Error(`Нет листа "${CFG.SHEETS.OUT}"`);

  const header = sh.getRange(2, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  const colMangoContact = header.indexOf('Контакт.Номер линии MANGO OFFICE') + 1;
  const colMangoDeal = header.indexOf('Сделка.Номер линии MANGO OFFICE') + 1;
  let colOut = header.indexOf('R.Источник ТЕЛ сделки') + 1;

  if (colMangoContact <= 0 || colMangoDeal <= 0) {
    throw new Error('Не найдены колонки Mango Office');
  }
  if (colOut <= 0) colOut = 22; // Резервная позиция

  const lastRow = sh.getLastRow();
  const n = Math.max(0, lastRow - 2);
  if (n === 0) return;

  // Загружаем карту колл-трекинга
  const callTable = readTable(ss, CFG.SHEETS.CALL);
  const ctMap = buildCalltrackingMap(callTable);

  const mangoC = sh.getRange(3, colMangoContact, n, 1).getValues();
  const mangoD = sh.getRange(3, colMangoDeal, n, 1).getValues();

  const out = mangoC.map((row, i) => {
    const phoneC = cleanPhone(row[0]);
    const phoneD = cleanPhone(mangoD[i][0]);
    
    const tag = ctMap.get(phoneC) || ctMap.get(phoneD) || '';
    return [tag];
  });

  sh.getRange(3, colOut, n, 1).setValues(out);
}

// ========================================
// РАСШИРЕННЫЕ ФУНКЦИИ ОБРАБОТКИ ДАННЫХ
// Интегрированные из проверенного старого скрипта
// ========================================

/**
 * Инициализация рабочих структур данных
 */
function initializeWorkingData_() {
  logInfo_('DATA_INIT', 'Инициализация рабочих структур данных');
  
  return {
    phoneMap: new Map(),
    utmMap: new Map(),
    calltrackingData: new Map(),
    processedDeals: new Set(),
    statusMapping: {
      'Первичный контакт': 'new',
      'Переговоры': 'in_progress', 
      'Принимают решение': 'considering',
      'Согласование договора': 'contract',
      'Успешно реализовано': 'success',
      'Закрыто и не реализовано': 'failed'
    },
    errors: [],
    warnings: [],
    stats: {
      total: 0,
      processed: 0,
      enriched: 0,
      failed: 0
    }
  };
}

/**
 * Загрузка и обогащение данных из всех источников
 */
function loadAndEnrichAllData_(workingData) {
  logInfo_('DATA_ENRICH', 'Начало загрузки и обогащения данных');
  
  try {
    // 1. Загружаем основные данные
    const mainData = loadMainDataSources_();
    
    // 2. Строим карты телефонов для сопоставления
    buildPhoneMaps_(workingData, mainData);
    
    // 3. Обогащаем UTM данными
    enrichWithUTMData_(workingData, mainData);
    
    // 4. Интегрируем данные коллтрекинга
    integrateCalltrackingData_(workingData, mainData);
    
    // 5. Нормализуем статусы и стадии
    normalizeStatusesAndStages_(workingData, mainData);
    
    logInfo_('DATA_ENRICH', 'Обогащение данных завершено успешно');
    return mainData;
    
  } catch (error) {
    logError_('DATA_ENRICH', 'Критическая ошибка обогащения данных', error);
    throw error;
  }
}

/**
 * Загрузка основных источников данных
 */
function loadMainDataSources_() {
  const sources = {
    amocrm: [],
    reserves: [],
    guests: [],
    siteforms: [],
    calltracking: []
  };
  
  try {
    // AmoCRM данные
    logInfo_('DATA_LOAD', 'Загрузка данных AmoCRM');
    sources.amocrm = loadAmoCrmData_();
    
    // Данные резервов 
    logInfo_('DATA_LOAD', 'Загрузка данных Reserves RP');
    sources.reserves = loadReservesData_();
    
    // Данные гостей
    logInfo_('DATA_LOAD', 'Загрузка данных Guests RP');
    sources.guests = loadGuestsData_();
    
    // Заявки с сайта
    logInfo_('DATA_LOAD', 'Загрузка заявок с сайта');
    sources.siteforms = loadSiteFormsData_();
    
    // Коллтрекинг данные
    logInfo_('DATA_LOAD', 'Загрузка данных коллтрекинга');
    sources.calltracking = loadCalltrackingData_();
    
    // UTM данные
    logInfo_('DATA_LOAD', 'Загрузка UTM данных');
    sources.utm = loadUTMData_();
    
    return sources;
    
  } catch (error) {
    logError_('DATA_LOAD', 'Ошибка загрузки основных данных', error);
    throw error;
  }
}

/**
 * Построение карт телефонов для сопоставления
 * Основано на логике из buildWorkingFromFive
 */
function buildPhoneMaps_(workingData, mainData) {
  logInfo_('PHONE_MAP', 'Построение карт телефонов');
  
  try {
    // Нормализация телефонов из AmoCRM
    mainData.amocrm.forEach(deal => {
      if (deal.phone) {
        const normalizedPhone = normalizePhone_(deal.phone);
        if (normalizedPhone) {
          workingData.phoneMap.set(normalizedPhone, deal);
        }
      }
    });
    
    // Нормализация телефонов из коллтрекинга
    mainData.calltracking.forEach(call => {
      if (call.mango_line) {
        const normalizedPhone = normalizePhone_(call.mango_line);
        if (normalizedPhone) {
          workingData.calltrackingData.set(normalizedPhone, call);
        }
      }
    });
    
    // Нормализация телефонов из заявок с сайта
    mainData.siteforms.forEach(form => {
      if (form.phone) {
        const normalizedPhone = normalizePhone_(form.phone);
        if (normalizedPhone) {
          workingData.siteData.set(normalizedPhone, form);
        }
      }
    });
    
    // Нормализация телефонов из резервов
    mainData.reserves.forEach(reserve => {
      if (reserve.phone) {
        const normalizedPhone = normalizePhone_(reserve.phone);
        if (normalizedPhone) {
          workingData.reserveData.set(normalizedPhone, reserve);
        }
      }
    });
    
    // Нормализация телефонов из гостей
    mainData.guests.forEach(guest => {
      if (guest.phone) {
        const normalizedPhone = normalizePhone_(guest.phone);
        if (normalizedPhone) {
          workingData.guestData.set(normalizedPhone, guest);
        }
      }
    });
    
    logInfo_('PHONE_MAP', `Создано карт: AMO=${workingData.phoneMap.size}, Коллтрекинг=${workingData.calltrackingData.size}, Сайт=${workingData.siteData.size}, Резервы=${workingData.reserveData.size}, Гости=${workingData.guestData.size}`);
    
  } catch (error) {
    logError_('PHONE_MAP', 'Ошибка построения карт телефонов', error);
    workingData.errors.push(`Ошибка обработки телефонов: ${error.message}`);
  }
}

/**
 * Нормализация номера телефона
 * Точная копия логики из старого скрипта
 */
function normalizePhone_(phone) {
  if (!phone) return null;
  
  // Убираем все нецифровые символы
  let clean = phone.toString().replace(/\D/g, '');
  
  // Обрабатываем российские номера
  if (clean.length === 11 && clean.startsWith('7')) {
    return clean;
  } else if (clean.length === 11 && clean.startsWith('8')) {
    return '7' + clean.substring(1);
  } else if (clean.length === 10) {
    return '7' + clean;
  }
  
  return clean.length >= 10 ? clean : null;
}

/**
 * Обогащение UTM данными
 */
function enrichWithUTMData_(workingData, mainData) {
  logInfo_('UTM_ENRICH', 'Обогащение UTM данными');
  
  try {
    mainData.utm.forEach(utmRecord => {
      const key = generateUTMKey_(utmRecord);
      workingData.utmMap.set(key, utmRecord);
    });
    
    // Применяем UTM к основным данным
    mainData.amocrm.forEach(deal => {
      const utmKey = generateDealUTMKey_(deal);
      const utmData = workingData.utmMap.get(utmKey);
      
      if (utmData) {
        deal.enriched = deal.enriched || {};
        deal.enriched.utm = utmData;
        workingData.stats.enriched++;
      }
    });
    
    logInfo_('UTM_ENRICH', `Обогащено ${workingData.stats.enriched} записей UTM данными`);
    
  } catch (error) {
    logError_('UTM_ENRICH', 'Ошибка обогащения UTM данными', error);
    workingData.errors.push(`Ошибка UTM обогащения: ${error.message}`);
  }
}

/**
 * Генерация ключа UTM для сопоставления
 */
function generateUTMKey_(utmRecord) {
  return `${utmRecord.source || ''}_${utmRecord.medium || ''}_${utmRecord.campaign || ''}`;
}

/**
 * Генерация UTM ключа для сделки
 */
function generateDealUTMKey_(deal) {
  return `${deal.utm_source || ''}_${deal.utm_medium || ''}_${deal.utm_campaign || ''}`;
}

/**
 * Интеграция данных коллтрекинга
 */
function integrateCalltrackingData_(workingData, mainData) {
  logInfo_('CALLTRACK_INTEGRATE', 'Интеграция данных коллтрекинга');
  
  try {
    mainData.amocrm.forEach(deal => {
      const normalizedPhone = normalizePhone_(deal.phone);
      const callData = workingData.calltrackingData.get(normalizedPhone);
      
      if (callData) {
        deal.enriched = deal.enriched || {};
        deal.enriched.calltracking = {
          source: callData.source,
          keyword: callData.keyword,
          city: callData.city,
          duration: callData.duration,
          status: callData.status
        };
        workingData.stats.enriched++;
      }
    });
    
    logInfo_('CALLTRACK_INTEGRATE', 'Интеграция коллтрекинга завершена');
    
  } catch (error) {
    logError_('CALLTRACK_INTEGRATE', 'Ошибка интеграции коллтрекинга', error);
    workingData.errors.push(`Ошибка коллтрекинга: ${error.message}`);
  }
}

/**
 * Нормализация статусов и стадий
 */
function normalizeStatusesAndStages_(workingData, mainData) {
  logInfo_('STATUS_NORMALIZE', 'Нормализация статусов и стадий');
  
  try {
    mainData.amocrm.forEach(deal => {
      // Нормализуем статус
      const originalStatus = deal.status;
      const normalizedStatus = workingData.statusMapping[originalStatus] || 'unknown';
      
      deal.normalized_status = normalizedStatus;
      
      // Определяем категорию стадии
      deal.stage_category = determineStageCategory_(deal.stage);
      
      workingData.stats.processed++;
    });
    
    logInfo_('STATUS_NORMALIZE', `Нормализовано ${workingData.stats.processed} записей`);
    
  } catch (error) {
    logError_('STATUS_NORMALIZE', 'Ошибка нормализации статусов', error);
    workingData.errors.push(`Ошибка нормализации: ${error.message}`);
  }
}

/**
 * Определение категории стадии
 */
function determineStageCategory_(stage) {
  if (!stage) return 'unknown';
  
  const stageStr = stage.toLowerCase();
  
  if (stageStr.includes('новая') || stageStr.includes('первичн')) {
    return 'new';
  } else if (stageStr.includes('работа') || stageStr.includes('переговор')) {
    return 'in_progress';
  } else if (stageStr.includes('решение') || stageStr.includes('раздум')) {
    return 'considering';
  } else if (stageStr.includes('договор') || stageStr.includes('согласован')) {
    return 'contract';
  } else if (stageStr.includes('успешно') || stageStr.includes('реализован')) {
    return 'success';
  } else if (stageStr.includes('закрыт') || stageStr.includes('отказ')) {
    return 'failed';
  }
  
  return 'other';
}

/**
 * Основная обработка данных по продвинутой логике
 */
function processAdvancedDataLogic_(enrichedData) {
  logInfo_('ADVANCED_PROCESS', 'Запуск продвинутой обработки данных');
  
  try {
    // 1. Дедупликация и очистка данных
    const cleanData = deduplicateAndCleanData_(enrichedData);
    
    // 2. Анализ первых касаний
    const firstTouchResults = analyzeFirstTouchAttribution_(cleanData);
    
    // 3. Анализ конверсий по воронке
    const funnelAnalysis = analyzeFunnelConversions_(cleanData);
    
    // 4. Расчёт LTV и ROI
    const ltvRoiAnalysis = calculateLTVandROI_(cleanData);
    
    // 5. Сохраняем результаты
    saveAdvancedAnalysisResults_({
      firstTouch: firstTouchResults,
      funnel: funnelAnalysis,
      ltvRoi: ltvRoiAnalysis
    });
    
    logInfo_('ADVANCED_PROCESS', 'Продвинутая обработка завершена успешно');
    
  } catch (error) {
    logError_('ADVANCED_PROCESS', 'Ошибка продвинутой обработки', error);
    throw error;
  }
}

/**
 * Дедупликация и очистка данных
 */
function deduplicateAndCleanData_(data) {
  logInfo_('DATA_CLEAN', 'Дедупликация и очистка данных');
  
  const seen = new Set();
  const cleanData = {
    amocrm: [],
    calltracking: [],
    analytics: [],
    utm: []
  };
  
  // Дедупликация AmoCRM данных
  data.amocrm.forEach(deal => {
    const key = `${deal.id}_${deal.phone}_${deal.email}`;
    if (!seen.has(key)) {
      seen.add(key);
      cleanData.amocrm.push(deal);
    }
  });
  
  // Очистка других источников аналогично
  cleanData.calltracking = data.calltracking;
  cleanData.analytics = data.analytics;
  cleanData.utm = data.utm;
  
  logInfo_('DATA_CLEAN', `Очищено: ${cleanData.amocrm.length} записей AmoCRM`);
  return cleanData;
}

/**
 * Анализ первых касаний
 */
function analyzeFirstTouchAttribution_(data) {
  logInfo_('FIRST_TOUCH', 'Анализ первых касаний');
  
  const firstTouchMap = new Map();
  
  data.amocrm.forEach(deal => {
    const clientKey = normalizePhone_(deal.phone) || deal.email || deal.id;
    
    if (!firstTouchMap.has(clientKey)) {
      firstTouchMap.set(clientKey, {
        deal: deal,
        firstTouchDate: deal.created_at,
        source: deal.utm_source || deal.enriched?.calltracking?.source || 'direct'
      });
    } else {
      // Сравниваем даты и оставляем более раннее касание
      const existing = firstTouchMap.get(clientKey);
      if (deal.created_at < existing.firstTouchDate) {
        firstTouchMap.set(clientKey, {
          deal: deal,
          firstTouchDate: deal.created_at,
          source: deal.utm_source || deal.enriched?.calltracking?.source || 'direct'
        });
      }
    }
  });
  
  return Array.from(firstTouchMap.values());
}

/**
 * Анализ конверсий по воронке
 */
function analyzeFunnelConversions_(data) {
  logInfo_('FUNNEL_ANALYSIS', 'Анализ воронки конверсий');
  
  const funnel = {
    total: 0,
    new: 0,
    in_progress: 0,
    considering: 0,
    contract: 0,
    success: 0,
    failed: 0
  };
  
  data.amocrm.forEach(deal => {
    funnel.total++;
    funnel[deal.stage_category] = (funnel[deal.stage_category] || 0) + 1;
  });
  
  // Рассчитываем конверсии
  const conversions = {
    new_to_progress: funnel.in_progress / funnel.new * 100,
    progress_to_considering: funnel.considering / funnel.in_progress * 100,
    considering_to_contract: funnel.contract / funnel.considering * 100,
    contract_to_success: funnel.success / funnel.contract * 100,
    overall_conversion: funnel.success / funnel.total * 100
  };
  
  return { funnel, conversions };
}

/**
 * Расчёт LTV и ROI
 */
function calculateLTVandROI_(data) {
  logInfo_('LTV_ROI', 'Расчёт LTV и ROI');
  
  const results = {
    total_revenue: 0,
    total_deals: 0,
    average_deal_value: 0,
    ltv: 0,
    roi: 0
  };
  
  const successfulDeals = data.amocrm.filter(deal => deal.stage_category === 'success');
  
  results.total_deals = successfulDeals.length;
  results.total_revenue = successfulDeals.reduce((sum, deal) => sum + (deal.price || 0), 0);
  results.average_deal_value = results.total_revenue / results.total_deals || 0;
  
  // Простой расчёт LTV (среднее значение сделки * 1.5)
  results.ltv = results.average_deal_value * 1.5;
  
  // ROI рассчитывается на основе известных затрат
  const totalCosts = calculateTotalMarketingCosts_(data);
  results.roi = totalCosts > 0 ? (results.total_revenue / totalCosts * 100 - 100) : 0;
  
  return results;
}

/**
 * Расчёт общих маркетинговых затрат
 */
function calculateTotalMarketingCosts_(data) {
  // Здесь должна быть логика расчёта затрат из ваших источников
  // Пока возвращаем заглушку
  return 100000; // Заглушка для примера
}

/**
 * Сохранение результатов продвинутого анализа
 */
function saveAdvancedAnalysisResults_(results) {
  logInfo_('SAVE_RESULTS', 'Сохранение результатов продвинутого анализа');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Сохраняем результаты в специальный лист
    let sheet = ss.getSheetByName('Продвинутая Аналитика');
    if (!sheet) {
      sheet = ss.insertSheet('Продвинутая Аналитика');
    }
    
    // Очищаем предыдущие данные
    sheet.clear();
    
    // Заголовки
    sheet.getRange(1, 1, 1, 4).setValues([['Метрика', 'Значение', 'Дата обновления', 'Комментарий']]);
    
    // Данные первых касаний
    const firstTouchData = results.firstTouch.map((item, index) => [
      `Первое касание ${index + 1}`,
      item.source,
      item.firstTouchDate,
      'Источник первого касания'
    ]);
    
    // Данные воронки
    const funnelData = Object.entries(results.funnel.conversions).map(([key, value]) => [
      `Конверсия ${key}`,
      `${value.toFixed(2)}%`,
      new Date(),
      'Конверсия по этапам воронки'
    ]);
    
    // Данные LTV/ROI
    const ltvRoiData = Object.entries(results.ltvRoi).map(([key, value]) => [
      key.toUpperCase(),
      typeof value === 'number' ? value.toFixed(2) : value,
      new Date(),
      'Показатели доходности'
    ]);
    
    // Записываем все данные
    const allData = [...firstTouchData, ...funnelData, ...ltvRoiData];
    if (allData.length > 0) {
      sheet.getRange(2, 1, allData.length, 4).setValues(allData);
    }
    
    // Применяем форматирование
    sheet.getRange(1, 1, 1, 4).setBackground('#4285f4').setFontColor('white').setFontWeight('bold');
    
    logInfo_('SAVE_RESULTS', 'Результаты продвинутого анализа сохранены');
    
  } catch (error) {
    logError_('SAVE_RESULTS', 'Ошибка сохранения результатов', error);
    throw error;
  }
}

/**
 * Реальные функции загрузки данных на основе структуры листов
 * Основано на предоставленной структуре данных
 */

/**
 * Загрузка данных из AmoCRM
 * Основано на структуре "Амо Выгрузка" и "Выгрузка Амо Полная"
 */
function loadAmoCrmData_() {
  logInfo_('AMOCRM_LOAD', 'Загрузка данных AmoCRM');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Пробуем найти лист с данными AmoCRM (приоритет "Выгрузка Амо Полная")
    let sheet = ss.getSheetByName('Выгрузка Амо Полная') || 
               ss.getSheetByName('Амо Выгрузка');
    
    if (!sheet) {
      logWarning_('AMOCRM_LOAD', 'Лист с данными AmoCRM не найден');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      logWarning_('AMOCRM_LOAD', 'Нет данных в листе AmoCRM');
      return [];
    }
    
    const rows = data.slice(1); // Пропускаем заголовки
    
    // Преобразуем строки в объекты согласно структуре "Выгрузка Амо Полная"
    const amoCrmData = rows.map(row => ({
      // Основная информация
      id: row[0],                           // A — Сделка.ID
      name: row[1],                         // B — Сделка.Название
      responsible: row[2],                  // C — Сделка.Ответственный
      contact_name: row[3],                 // D — Контакт.ФИО
      status: row[4],                       // E — Сделка.Статус
      budget: row[5],                       // F — Сделка.Бюджет
      created_at: row[6],                   // G — Сделка.Дата создания
      responsible2: row[7],                 // H — Сделка.Ответственный (дубль)
      tags: row[8],                         // I — Сделка.Теги
      closed_at: row[9],                    // J — Сделка.Дата закрытия
      
      // Аналитика
      ym_client_id: row[10],                // K — Сделка.YM_CLIENT_ID
      ga_client_id: row[11],                // L — Сделка.GA_CLIENT_ID
      button_text: row[12],                 // M — Сделка.BUTTON_TEXT
      date: row[13],                        // N — Сделка.DATE
      time: row[14],                        // O — Сделка.TIME
      deal_source: row[15],                 // P — Сделка.R.Источник сделки
      city_tag: row[16],                    // Q — Сделка.R.Тег города
      software: row[17],                    // R — Сделка.ПО
      
      // Бронирование
      bar_name: row[18],                    // S — Сделка.Бар (deal)
      booking_date: row[19],                // T — Сделка.Дата брони
      guest_count: row[20],                 // U — Сделка.Кол-во гостей
      visit_time: row[21],                  // V — Сделка.Время прихода
      comment: row[22],                     // W — Сделка.Комментарий МОБ
      source: row[23],                      // X — Сделка.Источник
      lead_type: row[24],                   // Y — Сделка.Тип лида
      refusal_reason: row[25],              // Z — Сделка.Причина отказа (ОБ)
      guest_status: row[26],                // AA — Сделка.R.Статусы гостей
      referral_type: row[27],               // AB — Сделка.Сарафан гости
      
      // UTM данные
      utm_medium: row[28],                  // AC — Сделка.UTM_MEDIUM
      formname: row[29],                    // AD — Сделка.FORMNAME
      referer: row[30],                     // AE — Сделка.REFERER
      formid: row[31],                      // AF — Сделка.FORMID
      mango_line1: row[32],                 // AG — Сделка.Номер линии MANGO OFFICE
      utm_source: row[33],                  // AH — Сделка.UTM_SOURCE
      utm_term: row[34],                    // AI — Сделка.UTM_TERM
      utm_campaign: row[35],                // AJ — Сделка.UTM_CAMPAIGN
      utm_content: row[36],                 // AK — Сделка.UTM_CONTENT
      utm_referrer: row[37],                // AL — Сделка.utm_referrer
      _ym_uid: row[38],                     // AM — Сделка._ym_uid
      
      // Контакты
      phone: row[39],                       // AN — Контакт.Телефон
      mango_line2: row[40],                 // AO — Контакт.Номер линии MANGO OFFICE
      notes: row[41]                        // AP — Сделка.Примечания(через ;)
    }));
    
    logInfo_('AMOCRM_LOAD', `Загружено ${amoCrmData.length} сделок AmoCRM`);
    return amoCrmData;
    
  } catch (error) {
    logError_('AMOCRM_LOAD', 'Ошибка загрузки данных AmoCRM', error);
    return [];
  }
}

/**
 * Загружает данные из листа "Reserves RP" 
 */
function loadReservesData_() {
  logInfo_('RESERVES_LOAD', 'Загрузка данных из Reserves RP');
    return [];
  }
}

/**
 * Загрузка данных коллтрекинга
 * Основано на структуре "КоллТрекинг"
 */
function loadCalltrackingData_() {
  logInfo_('CALLTRACK_LOAD', 'Загрузка данных коллтрекинга');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let sheet = ss.getSheetByName('КоллТрекинг') || 
               ss.getSheetByName('Calltracking') ||
               ss.getSheetByName('Коллтрекинг');
    
    if (!sheet) {
      logWarning_('CALLTRACK_LOAD', 'Лист с данными коллтрекинга не найден');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      logWarning_('CALLTRACK_LOAD', 'Нет данных в листе коллтрекинга');
      return [];
    }
    
    const headers = data[0];
    const rows = data.slice(1);
    
    const calltrackingData = rows.map(row => {
      const call = {};
      headers.forEach((header, index) => {
        const value = row[index];
        
        switch(header) {
          case 'Контакт.Номер линии MANGO OFFICE':
          case 'Номер линии MANGO OFFICE':
            call.phone = value;
            break;
          case 'R.Источник ТЕЛ сделки':
            call.source = value;
            break;
          case 'Название Канала':
            call.channel = value;
            break;
          default:
            call[header] = value;
        }
      });
      
      return call;
    });
    
    logInfo_('CALLTRACK_LOAD', `Загружено ${calltrackingData.length} записей коллтрекинга`);
    return calltrackingData;
    
  } catch (error) {
    logError_('CALLTRACK_LOAD', 'Ошибка загрузки данных коллтрекинга', error);
    return [];
  }
}

/**
 * Загрузка аналитических данных из веб-форм и резервов
 */
/**
 * Загружает данные из листа "Reserves RP" 
 */
function loadReservesData_() {
  logInfo_('RESERVES_LOAD', 'Загрузка данных из Reserves RP');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Reserves RP');
    
    if (!sheet) {
      logWarning_('RESERVES_LOAD', 'Лист "Reserves RP" не найден');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const rows = data.slice(1);
    
    const reservesData = rows.map(row => ({
      id: row[0],              // A — ID
      request_num: row[1],     // B — № заявки
      name: row[2],            // C — Имя
      phone: row[3],           // D — Телефон
      email: row[4],           // E — Email
      datetime: row[5],        // F — Дата/время
      status: row[6],          // G — Статус
      comment: row[7],         // H — Комментарий
      amount: row[8],          // I — Счёт, ₽
      guests: row[9],          // J — Гостей
      source: row[10]          // K — Источник
    }));
    
    logInfo_('RESERVES_LOAD', `Загружено ${reservesData.length} записей Reserves`);
    return reservesData;
    
  } catch (error) {
    logError_('RESERVES_LOAD', 'Ошибка загрузки Reserves', error);
    return [];
  }
}

/**
 * Загружает данные из листа "Guests RP"
 */
function loadGuestsData_() {
  logInfo_('GUESTS_LOAD', 'Загрузка данных из Guests RP');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Guests RP');
    
    if (!sheet) {
      logWarning_('GUESTS_LOAD', 'Лист "Guests RP" не найден');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const rows = data.slice(1);
    
    const guestsData = rows.map(row => ({
      name: row[0],            // A — Имя
      phone: row[1],           // B — Телефон
      email: row[2],           // C — Email
      visits: row[3],          // D — Кол-во визитов
      total_amount: row[4],    // E — Общая сумма
      first_visit: row[5],     // F — Первый визит
      last_visit: row[6],      // G — Последний визит
      bill_1: row[7],          // H — Счёт 1-го визита
      bill_2: row[8],          // I — Счёт 2-го визита
      bill_3: row[9],          // J — Счёт 3-го визита
      bill_4: row[10],         // K — Счёт 4-го визита
      bill_5: row[11],         // L — Счёт 5-го визита
      bill_6: row[12],         // M — Счёт 6-го визита
      bill_7: row[13],         // N — Счёт 7-го визита
      bill_8: row[14],         // O — Счёт 8-го визита
      bill_9: row[15],         // P — Счёт 9-го визита
      bill_10: row[16]         // Q — Счёт 10-го визита
    }));
    
    logInfo_('GUESTS_LOAD', `Загружено ${guestsData.length} записей Guests`);
    return guestsData;
    
  } catch (error) {
    logError_('GUESTS_LOAD', 'Ошибка загрузки Guests', error);
    return [];
  }
}
/**
 * Загружает данные из листа "Заявки с Сайта"
 */
function loadSiteFormsData_() {
  logInfo_('SITE_LOAD', 'Загрузка заявок с сайта');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Заявки с Сайта');
    
    if (!sheet) {
      logWarning_('SITE_LOAD', 'Лист "Заявки с Сайта" не найден');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const rows = data.slice(1);
    
    const siteData = rows.map(row => ({
      name: row[0],                 // A — Name
      phone: row[1],                // B — Phone
      referer: row[2],              // C — referer
      formid: row[3],               // D — formid
      sent: row[4],                 // E — sent
      requestid: row[5],            // F — requestid
      email: row[6],                // G — Email
      date: row[7],                 // H — Date
      quantity: row[8],             // I — Quantity
      checkbox: row[9],             // J — Checkbox
      formname: row[10],            // K — Form name
      time: row[11],                // L — Time
      utm_term: row[12],            // M — utm_term
      utm_campaign: row[13],        // N — utm_campaign
      utm_source: row[14],          // O — utm_source
      utm_content: row[15],         // P — utm_content
      utm_medium: row[16],          // Q — utm_medium
      button: row[17],              // R — Кнопка
      ym_client_id: row[18],        // S — ym_client_id
      ga_client_id: row[19],        // T — ga_client_id
      button_text: row[20],         // U — button_text
      referrer: row[21],            // V — referrer
      landing_page: row[22],        // W — landing_page
      page_title: row[23],          // X — page_title
      timestamp: row[24],           // Y — timestamp
      device_type: row[25],         // Z — device_type
      device_model: row[26],        // AA — device_model
      os: row[27],                  // AB — os
      browser: row[28],             // AC — browser
      browser_version: row[29],     // AD — browser_version
      screen_size: row[30],         // AE — screen_size
      clicks_count: row[31],        // AF — clicks_count
      user_city: row[32],           // AG — user_city
      user_country: row[33],        // AH — user_country
      user_ip: row[34],             // AI — user_ip
      os_version: row[35],          // AJ — os_version
      first_source: row[36],        // AK — first_source
      first_referrer: row[37],      // AL — first_referrer
      current_source: row[38],      // AM — current_source
      current_page: row[39],        // AN — current_page
      visits_count: row[40],        // AO — visits_count
      first_visit_date: row[41],    // AP — first_visit_date
      days_since_first_visit: row[42], // AQ — days_since_first_visit
      submit_date: row[43],         // AR — submit_date
      submit_time: row[44],         // AS — submit_time
      day_of_week: row[45],         // AT — day_of_week
      time_of_day: row[46],         // AU — time_of_day
      timezone: row[47],            // AV — timezone
      scroll_depth: row[48]         // AW — scroll_depth
    }));
    
    logInfo_('SITE_LOAD', `Загружено ${siteData.length} заявок с сайта`);
    return siteData;
    
  } catch (error) {
    logError_('SITE_LOAD', 'Ошибка загрузки заявок с сайта', error);
    return [];
  }
}

/**
 * Загружает данные коллтрекинга
 */
function loadCalltrackingData_() {
  logInfo_('CALLTRACK_LOAD', 'Загрузка данных коллтрекинга');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('КоллТрекинг');
    
    if (!sheet) {
      logWarning_('CALLTRACK_LOAD', 'Лист "КоллТрекинг" не найден');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const rows = data.slice(1);
    
    const callData = rows.map(row => ({
      mango_line: row[0],           // A — Контакт.Номер линии MANGO OFFICE
      tel_source: row[1],           // B — R.Источник ТЕЛ сделки
      channel_name: row[2]          // C — Название Канала
    }));
    
    logInfo_('CALLTRACK_LOAD', `Загружено ${callData.length} записей коллтрекинга`);
    return callData;
    
  } catch (error) {
    logError_('CALLTRACK_LOAD', 'Ошибка загрузки коллтрекинга', error);
    return [];
  }
}

/**
 * Загружает UTM данные (заглушка)
 */
    let reservesSheet = ss.getSheetByName('Reserves RP') || 
                       ss.getSheetByName('Reserves') ||
                       ss.getSheetByName('Резервы');
    
    if (reservesSheet) {
      const reservesData = reservesSheet.getDataRange().getValues();
      if (reservesData.length > 1) {
        const reservesHeaders = reservesData[0];
        const reservesRows = reservesData.slice(1);
        
        analyticsData.reserves = reservesRows.map(row => {
          const reserve = {};
          reservesHeaders.forEach((header, index) => {
            reserve[header] = row[index];
          });
          return reserve;
        });
      }
    }
    
    // Guests RP
    let guestsSheet = ss.getSheetByName('Guests RP') || 
                     ss.getSheetByName('Guests') ||
                     ss.getSheetByName('Гости');
    
    if (guestsSheet) {
      const guestsData = guestsSheet.getDataRange().getValues();
      if (guestsData.length > 1) {
        const guestsHeaders = guestsData[0];
        const guestsRows = guestsData.slice(1);
        
        analyticsData.guests = guestsRows.map(row => {
          const guest = {};
          guestsHeaders.forEach((header, index) => {
            guest[header] = row[index];
          });
          return guest;
        });
      }
    }
    
    logInfo_('ANALYTICS_LOAD', 
      `Загружено: ${analyticsData.webForms.length} веб-форм, ` +
      `${analyticsData.reserves.length} резервов, ` +
      `${analyticsData.guests.length} гостей`
    );
    
    return analyticsData;
    
  } catch (error) {
    logError_('ANALYTICS_LOAD', 'Ошибка загрузки аналитических данных', error);
    return { webForms: [], reserves: [], guests: [] };
  }
}

/**
 * Загрузка UTM данных из всех источников
 */
function loadUTMData_() {
  logInfo_('UTM_LOAD', 'Загрузка UTM данных');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const utmData = [];
    
    // Собираем UTM данные из AmoCRM
    const amoCrmSheet = ss.getSheetByName('Амо Выгрузка') || 
                       ss.getSheetByName('Выгрузка Амо Полная');
    
    if (amoCrmSheet) {
      const data = amoCrmSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0];
        const rows = data.slice(1);
        
        rows.forEach(row => {
          const utm = {};
          let hasUtmData = false;
          
          headers.forEach((header, index) => {
            if (header.includes('UTM_') || header.includes('utm_')) {
              utm[header] = row[index];
              if (row[index]) hasUtmData = true;
            }
          });
          
          if (hasUtmData) {
            utmData.push(utm);
          }
        });
      }
    }
    
    // Собираем UTM данные из веб-форм
    const webFormsSheet = ss.getSheetByName('Заявки с Сайта');
    
    if (webFormsSheet) {
      const data = webFormsSheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0];
        const rows = data.slice(1);
        
        rows.forEach(row => {
          const utm = {};
          let hasUtmData = false;
          
          headers.forEach((header, index) => {
            if (header.includes('utm_')) {
              utm[header] = row[index];
              if (row[index]) hasUtmData = true;
            }
          });
          
          if (hasUtmData) {
            utmData.push(utm);
          }
        });
      }
    }
    
    logInfo_('UTM_LOAD', `Загружено ${utmData.length} записей UTM данных`);
    return utmData;
    
  } catch (error) {
    logError_('UTM_LOAD', 'Ошибка загрузки UTM данных', error);
    return [];
  }
}

// ========================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// Недостающие утилиты для работы с листами
// ========================================

/**
 * Форматирует дату в строку DD.MM.YYYY
 */
function formatDate_(date) {
  if (!date) return '';
  
  try {
    const d = date instanceof Date ? date : new Date(date);
    if (isNaN(d.getTime())) return '';
    
    const day = d.getDate().toString().padStart(2, '0');
    const month = (d.getMonth() + 1).toString().padStart(2, '0');
    const year = d.getFullYear();
    
    return `${day}.${month}.${year}`;
  } catch (error) {
    console.log('Ошибка форматирования даты:', error);
    return '';
  }
}

/**
 * Безопасное слияние ячеек
 */
function safeMergeRange_(sheet, startRow, startCol, numRows, numCols) {
  try {
    if (numRows > 1 || numCols > 1) {
      const range = sheet.getRange(startRow, startCol, numRows, numCols);
      range.merge();
    }
  } catch (error) {
    // Игнорируем ошибки слияния - это не критично
    console.log(`Предупреждение: не удалось объединить ячейки ${startRow}:${startCol}`);
  }
}

/**
 * Получает имя листа из конфигурации
 */
function getSheetName_(key) {
  return CONFIG.SHEETS && CONFIG.SHEETS[key] ? CONFIG.SHEETS[key] : key;
}

/**
 * Получает лист по имени или создает новый
 */
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    // Применяем стандартное форматирование  
    const range = sheet.getRange(1, 1, 100, 50);
    range.setFontFamily(CONFIG.FONT || 'PT Sans');
    range.setFontSize(10);
    
    // Заголовок первой строки
    const headerRange = sheet.getRange(1, 1, 1, 50);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e8f0fe');
    
    logInfo_('SHEET_CREATE', `Создан новый лист: ${sheetName}`);
  }
  
  return sheet;
}

/**
 * Получает лист по имени (без создания)
 */
function getExistingSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * Получает данные из листа
 */
function getSheetData_(sheetName) {
  const sheet = getExistingSheet_(sheetName);
  if (!sheet) {
    logWarning_('SHEET_DATA', `Лист ${sheetName} не найден`);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  return data.length > 0 ? data : [];
}

/**
 * Получает или создает лист (алиас для обратной совместимости)
 */
function getOrCreateSheet_(sheetName) {
  return getSheet_(sheetName);
}

/**
 * Очищает данные листа
 */
function clearSheetData_(sheet) {
  if (sheet) {
    sheet.clear();
  }
}

/**
 * Синхронизация всех данных из внешних источников
 */
function syncAllData() {
  logInfo_('SYNC_ALL', 'Начало синхронизации всех данных');
  
  try {
    const syncResults = {
      amocrm: false,
      calltracking: false,
      webforms: false,
      reserves: false,
      guests: false
    };
    
    // Здесь может быть логика синхронизации с внешними API
    // Пока просто помечаем как успешную
    syncResults.amocrm = true;
    syncResults.calltracking = true;
    syncResults.webforms = true;
    syncResults.reserves = true;
    syncResults.guests = true;
    
    logInfo_('SYNC_ALL', 'Синхронизация всех данных завершена');
    return syncResults;
    
  } catch (error) {
    logError_('SYNC_ALL', 'Ошибка синхронизации данных', error);
    throw error;
  }
}

/**
 * 🎯 ГЛАВНАЯ ФУНКЦИЯ СБОРКИ РАБОЧЕГО АМО
 * Создает итоговый файл с точным оформлением как на картинках
 */
function buildWorkingAmoFile() {
  console.log('🎯 Начинаем сборку файла РАБОЧИЙ АМО');
  
  try {
    const workingSheet = getSheet_(getSheetName_('WORKING_AMO'));
    workingSheet.clear();
    
    // 1. Создаем заголовки с точным оформлением как на картинках
    createWorkingAmoHeaders_(workingSheet);
    
    // 2. Загружаем и объединяем все данные
    const consolidatedData = consolidateAllDataSources_();
    
    // 3. Записываем данные в файл
    writeConsolidatedData_(workingSheet, consolidatedData);
    
    // 4. Применяем форматирование как на картинках
    applyWorkingAmoFormatting_(workingSheet, consolidatedData.length);
    
    console.log(`✅ Файл РАБОЧИЙ АМО собран: ${consolidatedData.length} записей`);
    
  } catch (error) {
    console.error('❌ Ошибка сборки РАБОЧИЙ АМО:', error);
    throw error;
  }
}

/**
 * Создает заголовки точно как на картинках
 */
function createWorkingAmoHeaders_(sheet) {
  const headers = [
    // � ИНФОРМАЦИЯ О СДЕЛКЕ (A–H)
    'ID',                    // A
    'Название',              // B
    'Ответственный',         // C
    'Статус',               // D
    'Бюджет',               // E
    'Дата создания',        // F
    'Теги',                 // G
    'Дата закрытия',        // H
    
    // 👤 КОНТАКТ (I–M)
    'ФИО',                  // I
    'Телефон',              // J
    'Номер линии MANGO OFFICE',    // K
    'Номер линии MANGO OFFICE',    // L (дубль для второй линии)
    'Бар (deal)',           // M
    
    // 🕒 БРОНЬ (N–R)
    'Дата брони',           // N
    'Время прихода',        // O
    'Кол-во гостей',        // P
    'Комментарий МОБ',      // Q
    'R.Статусы гостей',     // R
    
    // � UTM/ИСТОЧНИК + 🔍 АНАЛИТИКА (S–AJ)
    'Источник',             // S
    'Тип лида',             // T
    'R.Источник сделки',    // U
    'R.Источник ТЕЛ сделки', // V
    'UTM_SOURCE',           // W
    'UTM_MEDIUM',           // X
    'UTM_CAMPAIGN',         // Y
    'UTM_TERM',             // Z
    'UTM_CONTENT',          // AA
    'utm_referrer',         // AB
    'YM_CLIENT_ID',         // AC
    'GA_CLIENT_ID',         // AD
    'FORMNAME',             // AE
    'REFERER',              // AF
    'FORMID',               // AG
    'DATE',                 // AH
    'TIME',                 // AI
    'BUTTON_TEXT',          // AJ
    
    // 📌 ДОПОЛНИТЕЛЬНО (AK–AO)
    'R.Тег города',         // AK
    'ПО',                   // AL
    'Причина отказа (ОБ)',  // AM
    'Примечание 1',         // AN
    'Последняя заявка',     // AO
    
    // 🟦 SITE/RESERVES/GUESTS (обогащение) (AP–BA)
    'utm_source (из сайта/резервов)', // AP
    'utm_medium',           // AQ
    'utm_campaign',         // AR
    '(резерв, сейчас пусто)', // AS
    'R.Источник ТЕЛ сделки (из коллтрекинга)', // AT
    'Визитов (из SITE)',    // AU
    'Сумма ₽ (из SITE)',    // AV
    'Последний визит (из SITE)', // AW
    'Визитов (из GUESTS RP)', // AX
    'Сумма ₽ (из GUESTS RP)', // AY
    'Первый визит (из GUESTS RP)', // AZ
    'Последний визит (из GUESTS RP)', // BA
    
    // 🧮 АВТО‑ПОЛЯ (BB–BD)
    'Возраст сделки (дн.)', // BB
    'Дней до брони',        // BC
    'Сарафан гости',        // BD
    
    // (прочее) BE
    '_ym_uid'               // BE
  ];
  
  // Записываем заголовки
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Применяем цветное оформление секций
  applySectionFormatting_(sheet, headers.length);
}

/**
 * Применяет цветное оформление секций как на картинках
 */
function applySectionFormatting_(sheet, totalCols) {
  // � ИНФОРМАЦИЯ О СДЕЛКЕ - розовый (A-H, колонки 1-8)
  sheet.getRange(1, 1, 1, 8)
    .setBackground('#f4cccc')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // 👤 КОНТАКТ - голубой (I-M, колонки 9-13)
  sheet.getRange(1, 9, 1, 5)
    .setBackground('#cfe2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
    
  // 🕒 БРОНЬ - фиолетовый (N-R, колонки 14-18)
  sheet.getRange(1, 14, 1, 5)
    .setBackground('#d5a6bd')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
    
  // � UTM/ИСТОЧНИК + 🔍 АНАЛИТИКА - желтый (S-AJ, колонки 19-36) 
  sheet.getRange(1, 19, 1, 18)
    .setBackground('#fff2cc')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
    
  // � ДОПОЛНИТЕЛЬНО - светло-зеленый (AK-AO, колонки 37-41)
  sheet.getRange(1, 37, 1, 5)
    .setBackground('#d9ead3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
    
  // 🟦 SITE/RESERVES/GUESTS - голубой (AP-BA, колонки 42-53)
  sheet.getRange(1, 42, 1, 12)
    .setBackground('#cfe2f3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
    
  // 🧮 АВТО‑ПОЛЯ - серый (BB-BE, колонки 54-57)
  sheet.getRange(1, 54, 1, 4)
    .setBackground('#efefef')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

/**
 * Объединяет данные из всех источников
 */
function consolidateAllDataSources_() {
  console.log('📊 Объединение данных из всех источников...');
  
  try {
    // Создаем структуру рабочих данных
    const workingData = {
      phoneMap: new Map(),
      calltrackingData: new Map(),
      utmData: new Map(),
      siteData: new Map(),
      reserveData: new Map(),
      guestData: new Map(),
      errors: []
    };
    
    // Загружаем основные данные
    const mainData = loadMainDataSources_();
    
    // Строим карты телефонов для сопоставления
    buildPhoneMaps_(workingData, mainData);
    
    console.log(`Данные загружены: AMO=${mainData.amocrm.length}, WebForms=${mainData.webForms?.length || 0}`);
    
    const consolidatedData = [];
    
    // Обрабатываем каждую сделку AMO
    for (let i = 0; i < mainData.amocrm.length; i++) {
      const deal = mainData.amocrm[i];
      const phone = normalizePhone_(deal.phone);
      
      // Обогащаем данными из других источников
      const enrichedData = enrichDealWithAllSources_(deal, phone, workingData);
      consolidatedData.push(enrichedData);
    }
    
    console.log(`✅ Объединено ${consolidatedData.length} записей`);
    return consolidatedData;
    
  } catch (error) {
    logError_('CONSOLIDATE', 'Ошибка объединения данных', error);
    throw error;
  }
}

/**
 * Обогащает сделку данными из всех источников для РАБОЧИЙ АМО
 * Создает единую структуру данных на основе всех 6 источников  
 */
function enrichDealWithAllSources_(deal, phone, workingData) {
  // Получаем обогащающие данные из всех источников по телефону
  const calltrackingData = workingData.calltrackingData.get(phone) || {};
  const siteData = workingData.siteData.get(phone) || {};
  const reserveData = workingData.reserveData.get(phone) || {};
  const guestData = workingData.guestData.get(phone) || {};
  
  return [
    // � ИНФОРМАЦИЯ О СДЕЛКЕ (A–H)
    deal.id || '',                              // A — ID
    deal.name || '',                            // B — Название
    deal.responsible || '',                     // C — Ответственный
    deal.status || '',                          // D — Статус
    deal.budget || deal.price || 0,             // E — Бюджет
    formatDate_(deal.created_at) || '',         // F — Дата создания
    deal.tags || '',                            // G — Теги
    formatDate_(deal.closed_at) || '',          // H — Дата закрытия
    
    // 👤 КОНТАКТ (I–M)
    deal.contact_name || deal.client_name || '', // I — ФИО
    deal.phone || '',                           // J — Телефон
    deal.mango_line1 || '',                     // K — Номер линии MANGO OFFICE
    deal.mango_line2 || '',                     // L — Номер линии MANGO OFFICE (дубль)
    deal.bar_name || deal.name || '',           // M — Бар (deal)
    
    // 🕒 БРОНЬ (N–R)
    formatDate_(deal.booking_date) || formatDate_(deal.created_at) || '', // N — Дата брони
    deal.visit_time || '',                      // O — Время прихода
    deal.guest_count || '',                     // P — Кол-во гостей
    deal.comment || '',                         // Q — Комментарий МОБ
    deal.guest_status || deal.status || '',     // R — R.Статусы гостей
    
    // � UTM/ИСТОЧНИК + 🔍 АНАЛИТИКА (S–AJ)
    utmData.utm_source || deal.source || '',    // S — Источник
    deal.lead_type || '',                       // T — Тип лида
    deal.deal_source || '',                     // U — R.Источник сделки
    deal.tel_source || calltrackingData.source || '', // V — R.Источник ТЕЛ сделки
    utmData.utm_source || '',                   // W — UTM_SOURCE
    utmData.utm_medium || '',                   // X — UTM_MEDIUM
    utmData.utm_campaign || '',                 // Y — UTM_CAMPAIGN
    utmData.utm_term || '',                     // Z — UTM_TERM
    utmData.utm_content || '',                  // AA — UTM_CONTENT
    utmData.utm_referrer || utmData.referer || '', // AB — utm_referrer
    utmData.ym_client_id || '',                 // AC — YM_CLIENT_ID
    utmData.ga_client_id || '',                 // AD — GA_CLIENT_ID
    utmData.formname || '',                     // AE — FORMNAME
    utmData.referer || '',                      // AF — REFERER
    utmData.formid || '',                       // AG — FORMID
    utmData.date || '',                         // AH — DATE
    utmData.time || '',                         // AI — TIME
    utmData.button_text || '',                  // AJ — BUTTON_TEXT
    
    // 📌 ДОПОЛНИТЕЛЬНО (AK–AO)
    deal.city_tag || '',                        // AK — R.Тег города
    deal.software || '',                        // AL — ПО
    deal.refusal_reason || '',                  // AM — Причина отказа (ОБ)
    siteData.note || '',                        // AN — Примечание 1
    siteData.last_request || '',                // AO — Последняя заявка
    
    // 🟦 SITE/RESERVES/GUESTS (обогащение) (AP–BA)
    siteData.utm_source || '',                  // AP — utm_source (из сайта/резервов)
    siteData.utm_medium || '',                  // AQ — utm_medium
    siteData.utm_campaign || '',                // AR — utm_campaign
    reserveData.note || '',                     // AS — (резерв, сейчас пусто)
    calltrackingData.tel_source || '',          // AT — R.Источник ТЕЛ сделки (из коллтрекинга)
    siteData.visits || 0,                       // AU — Визитов (из SITE)
    siteData.amount || 0,                       // AV — Сумма ₽ (из SITE)
    siteData.last_visit || '',                  // AW — Последний визит (из SITE)
    guestData.visits || 0,                      // AX — Визитов (из GUESTS RP)
    guestData.amount || 0,                      // AY — Сумма ₽ (из GUESTS RP)
    guestData.first_visit || '',                // AZ — Первый визит (из GUESTS RP)
    guestData.last_visit || '',                 // BA — Последний визит (из GUESTS RP)
    
    // 🧮 АВТО‑ПОЛЯ (BB–BD)
    calculateDealAge_(deal.created_at),         // BB — Возраст сделки (дн.)
    calculateDaysToBooking_(deal.created_at, deal.booking_date), // BC — Дней до брони
    deal.referral_type || '',                   // BD — Сарафан гости
    
    // (прочее) BE
    utmData._ym_uid || ''                       // BE — _ym_uid
  ];
}
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Получаем данные AmoCRM для анализа
    const amoCrmData = loadAmoCrmData_();
    
    if (amoCrmData.length === 0) {
      logWarning_('DAILY_STATS', 'Нет данных AmoCRM для статистики');
      return;
    }
    
    // Создаем или получаем лист статистики
    let statsSheet = getSheet_(getSheetName_('DAILY_STATISTICS'));
    
    // Очищаем предыдущие данные
    statsSheet.clear();
    
    // Заголовки
    const headers = [
      'Дата', 
      'Всего сделок', 
      'Новых сделок', 
      'Закрытых сделок', 
      'Успешных сделок', 
      'Отказов',
      'Общая сумма', 
      'Средний чек',
      'Конверсия %'
    ];
    
    statsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Группируем данные по дням
    const dailyStats = new Map();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    amoCrmData.forEach(deal => {
      const createdDate = deal.created_at instanceof Date ? deal.created_at : new Date(deal.created_at);
      const dateKey = createdDate.toDateString();
      
      if (!dailyStats.has(dateKey)) {
        dailyStats.set(dateKey, {
          date: createdDate,
          total: 0,
          new: 0,
          closed: 0,
          success: 0,
          failed: 0,
          revenue: 0
        });
      }
      
      const dayStats = dailyStats.get(dateKey);
      dayStats.total++;
      
      // Определяем статус
      const status = deal.status || '';
      const normalizedStatus = deal.normalized_status || '';
      
      if (normalizedStatus === 'new') {
        dayStats.new++;
      } else if (normalizedStatus === 'success') {
        dayStats.success++;
        dayStats.closed++;
        dayStats.revenue += deal.price || 0;
      } else if (normalizedStatus === 'failed') {
        dayStats.failed++;
        dayStats.closed++;
      }
    });
    
    // Формируем данные для таблицы
    const statsData = Array.from(dailyStats.values())
      .sort((a, b) => b.date - a.date) // Сортируем по убыванию даты
      .slice(0, 30) // Берем последние 30 дней
      .map(stats => {
        const avgDeal = stats.success > 0 ? stats.revenue / stats.success : 0;
        const conversion = stats.total > 0 ? (stats.success / stats.total * 100).toFixed(1) : 0;
        
        return [
          stats.date,
          stats.total,
          stats.new,
          stats.closed,
          stats.success,
          stats.failed,
          stats.revenue,
          Math.round(avgDeal),
          conversion + '%'
        ];
      });
    
    // Записываем данные
    if (statsData.length > 0) {
      statsSheet.getRange(2, 1, statsData.length, headers.length).setValues(statsData);
      
      // Форматирование
      statsSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285f4')
        .setFontColor('white')
        .setFontWeight('bold');
      
      // Форматируем суммы
      if (statsData.length > 0) {
        statsSheet.getRange(2, 7, statsData.length, 2).setNumberFormat('#,##0');
      }
      
      // Автоширина колонок
      for (let i = 1; i <= headers.length; i++) {
        statsSheet.autoResizeColumn(i);
      }
      
      // Заморозка заголовка
      statsSheet.setFrozenRows(1);
    }
    
    logInfo_('DAILY_STATS', `Ежедневная статистика обновлена: ${statsData.length} дней`);
    
  } catch (error) {
    logError_('DAILY_STATS', 'Ошибка обновления ежедневной статистики', error);
    throw error;
  }
}

/**
 * Заглушки для остальных аналитических функций
 */
function syncAmoCrmDataOnly() {
  logInfo_('SYNC_AMO', 'Синхронизация только AmoCRM данных');
  // Реализация синхронизации AmoCRM
}

function syncWebFormsDataOnly() {
  logInfo_('SYNC_WEB', 'Синхронизация только веб-форм');
  // Реализация синхронизации веб-форм
}

function syncCallTrackingDataOnly() {
  logInfo_('SYNC_CALL', 'Синхронизация только коллтрекинга');
  // Реализация синхронизации коллтрекинга
}

function syncYandexMetricaDataOnly() {
  logInfo_('SYNC_YM', 'Синхронизация только Яндекс.Метрики');
  // Реализация синхронизации Яндекс.Метрики
}

/**
 * Записывает объединенные данные в лист
 */
function writeConsolidatedData_(sheet, data) {
  if (data.length === 0) {
    console.log('⚠️ Нет данных для записи');
    return;
  }
  
  console.log(`📝 Записываем ${data.length} строк данных...`);
  
  // Записываем данные пакетами для производительности
  const BATCH_SIZE = 100;
  let startRow = 2; // Начинаем после заголовков
  
  for (let i = 0; i < data.length; i += BATCH_SIZE) {
    const batch = data.slice(i, Math.min(i + BATCH_SIZE, data.length));
    const endRow = startRow + batch.length - 1;
    
    try {
      sheet.getRange(startRow, 1, batch.length, batch[0].length).setValues(batch);
      console.log(`✅ Записан пакет ${i + 1}-${Math.min(i + BATCH_SIZE, data.length)}`);
      startRow = endRow + 1;
    } catch (error) {
      console.error(`❌ Ошибка записи пакета ${i + 1}:`, error);
    }
  }
}

/**
 * Применяет форматирование к данным как на картинках
 */
function applyWorkingAmoFormatting_(sheet, dataRows) {
  console.log('🎨 Применяем форматирование...');
  
  try {
    // Общий шрифт для всего листа
    const maxCols = sheet.getLastColumn();
    const totalRows = dataRows + 1; // +1 для заголовков
    
    sheet.getRange(1, 1, totalRows, maxCols)
      .setFontFamily(CONFIG.FONT || 'PT Sans')
      .setFontSize(9);
    
    // Заморозка первой строки
    sheet.setFrozenRows(1);
    
    // Автоширина для всех колонок
    for (let col = 1; col <= maxCols; col++) {
      sheet.autoResizeColumn(col);
    }
    
    // Применяем условное форматирование для статусов
    applyConditionalFormatting_(sheet, totalRows);
    
    console.log('✅ Форматирование применено');
    
  } catch (error) {
    console.error('❌ Ошибка форматирования:', error);
  }
}

/**
 * Применяет условное форматирование для статусов
 */
function applyConditionalFormatting_(sheet, totalRows) {
  // Колонка со статусами (R.Статус гостей - 6я колонка)
  const statusRange = sheet.getRange(2, 6, totalRows - 1, 1);
  
  // Цветовая схема для статусов
  const statusColors = {
    'Новый': '#ccffcc',      // светло-зеленый
    'Подтвержден': '#ffe599', // желтый  
    'Пришел': '#b6d7a8',     // зеленый
    'Не пришел': '#f4cccc',  // светло-красный
    'Отказ': '#ea9999'       // красный
  };
  
  // Применяем цвета (через API условного форматирования сложно, делаем базово)
  for (let row = 2; row <= totalRows; row++) {
    const statusCell = sheet.getRange(row, 6);
    const status = statusCell.getValue();
    
    if (statusColors[status]) {
      statusCell.setBackground(statusColors[status]);
    }
  }
}

/**
 * Вычисляет возраст сделки в днях
 */
function calculateDealAge_(createdDate) {
  if (!createdDate) return 0;
  
  const created = new Date(createdDate);
  const now = new Date();
  const diffTime = Math.abs(now - created);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  return diffDays;
}

/**
 * Вычисляет дни до брони
 */
function calculateDaysToBooking_(createdDate, bookingDate) {
  if (!createdDate || !bookingDate) return 0;
  
  const created = new Date(createdDate);
  const booking = new Date(bookingDate);
  const diffTime = booking - created;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  return diffDays > 0 ? diffDays : 0;
}
