/**
 * 🎯 Главный Скрипт - Центральная точка управления
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Apps Script: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit
 * 2. Создайте новый файл: + → Script file
 * 3. Назовите файл: 01_Главный_Скрипт (без .gs)
 * 4. Удалите весь код по умолчанию
 * 5. Скопируйте и вставьте код ниже
 * 6. Сохраните файл (Ctrl+S)
 */

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
    
    // Показываем сообщение пользователю
    SpreadsheetApp.getUi().alert(
      'Система инициализирована!',
      'Все модули настроены и готовы к работе.\n\n' +
      'Автоматическая синхронизация данных настроена на каждые 15 минут.\n' +
      'Аналитические отчёты обновляются ежедневно в 08:00.\n\n' +
      'Используйте меню "🔄 Аналитика" для ручного управления.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    logError_('SYSTEM_INIT', 'Критическая ошибка инициализации системы', error);
    
    SpreadsheetApp.getUi().alert(
      'Ошибка инициализации!',
      `Произошла ошибка при инициализации системы:\n\n${error.message}\n\nПроверьте настройки конфигурации и попробуйте снова.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
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

  } catch (e) {
    logToSheet(ss, `Ошибка обновления: ${e.toString()}`, 'ERROR');
    throw e;
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
