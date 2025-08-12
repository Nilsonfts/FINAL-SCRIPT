/**
 * МОДУЛЬ СИНХРОНИЗАЦИИ ДАННЫХ ИЗ GOOGLE SHEETS
 * Работа с данными из таблиц (без прямых API интеграций)
 * @fileoverview Центральный модуль для работы с данными из Google Sheets
 */

/**
 * Основная функция синхронизации всех данных из таблиц
 * Вызывается по расписанию каждые 15 минут
 */
function syncAllData() {
  try {
    logInfo_('DATA_SYNC', 'Начало обновления данных из таблиц');
    
    const startTime = new Date();
    const results = {
      amocrm_sheets: { success: false, records: 0, error: null },
      forms: { success: false, records: 0, error: null },
      calls: { success: false, records: 0, error: null },
      processing: { success: false, records: 0, error: null }
    };
    
    // 1. Обновление данных из листов AmoCRM (которые выгружает сама AmoCRM)
    try {
      logInfo_('DATA_SYNC', 'Обработка данных AmoCRM из таблиц...');
      const amocrmResult = processAmoCrmSheetsData_();
      results.amocrm_sheets = amocrmResult;
      logInfo_('DATA_SYNC', `AmoCRM таблицы: обработано ${amocrmResult.records} записей`);
    } catch (error) {
      results.amocrm_sheets.error = error.message;
      logError_('DATA_SYNC', 'Ошибка обработки данных AmoCRM', error);
    }
    
    // 2. Обработка заявок с сайта
    try {
      logInfo_('DATA_SYNC', 'Обработка форм сайта...');
      const formsResult = processWebFormsData_();
      results.forms = formsResult;
      logInfo_('DATA_SYNC', `Формы сайта: обработано ${formsResult.records} записей`);
    } catch (error) {
      results.forms.error = error.message;
      logError_('DATA_SYNC', 'Ошибка обработки форм сайта', error);
    }
    
    // 3. Обработка звонков из колл-трекинга
    try {
      logInfo_('DATA_SYNC', 'Обработка колл-трекинга...');
      const callsResult = processCallTrackingData_();
      results.calls = callsResult;
      logInfo_('DATA_SYNC', `Колл-трекинг: обработано ${callsResult.records} записей`);
    } catch (error) {
      results.calls.error = error.message;
      logError_('DATA_SYNC', 'Ошибка обработки колл-трекинга', error);
    }
    
    // 4. Объединение и обогащение данных
    try {
      logInfo_('DATA_SYNC', 'Объединение и обогащение данных...');
      const processingResult = mergeAndEnrichData_();
      results.processing = processingResult;
      logInfo_('DATA_SYNC', `Обогащение: обработано ${processingResult.records} записей`);
    } catch (error) {
      results.processing.error = error.message;
      logError_('DATA_SYNC', 'Ошибка объединения данных', error);
    }
    
    const duration = (new Date() - startTime) / 1000;
    const totalRecords = Object.values(results).reduce((sum, result) => sum + result.records, 0);
    
    logInfo_('DATA_SYNC', `Синхронизация завершена за ${duration}с. Обработано ${totalRecords} записей`);
    
    // Отправка уведомления при ошибках
    const hasErrors = Object.values(results).some(result => result.error);
    if (hasErrors) {
      sendErrorNotification_(results);
    }
    
    return results;
    
  } catch (error) {
    logError_('DATA_SYNC', 'Критическая ошибка синхронизации', error);
    throw error;
  }
}

/**
 * Обработка данных AmoCRM из таблиц Google Sheets
 * (данные уже выгружены самой AmoCRM в таблицы)
 * @returns {Object} Результат обработки
 * @private
 */
function processAmoCrmSheetsData_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let totalRecords = 0;
  
  try {
    // Используем существующую логику из основного скрипта
    const CFG = CONFIG.mainScript;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Читаем все таблицы
    const amoTable = readTable(ss, CFG.SHEETS.NEW);
    const amoFullTable = readTable(ss, CFG.SHEETS.FULL);
    const siteTable = readTable(ss, CFG.SHEETS.SITE);
    const resTable = readTable(ss, CFG.SHEETS.RES);
    const gueTable = readTable(ss, CFG.SHEETS.GUE);
    const callTable = readTable(ss, CFG.SHEETS.CALL);
    
    // Объединяем таблицы AmoCRM
    let tables = [];
    if (amoTable.rows.length > 0) tables.push(amoTable);
    if (amoFullTable.rows.length > 0) tables.push(amoFullTable);
    if (siteTable.rows.length > 0) tables.push(siteTable);
    
    if (tables.length === 0) {
      logInfo_('DATA_SYNC', 'Нет данных для обработки');
      return { success: true, records: 0, error: null };
    }
    
    // Объединяем по ключевому полю
    const merged = mergeByKey(tables, CFG.KEY);
    
    // Канонизируем заголовки
    const canonized = canonHeaders(merged, CFG.MAP);
    
    // Строим агрегаты для обогащения
    const phoneIdx = findColumnIndex(canonized.header, ['Телефон', 'Phone']);
    const resAgg = buildAggregates(resTable.rows, phoneIdx);
    const gueAgg = buildAggregates(gueTable.rows, phoneIdx);
    const ctMap = buildCalltrackingMap(callTable);
    
    // Обогащаем данные
    const enriched = buildEnrichedData(canonized, siteTable, resAgg, gueAgg, ctMap, CFG);
    
    // Преобразуем заголовки для отображения
    const displayHeader = enriched.headerOrderedRaw.map(humanizeHeader);
    
    // Сохраняем в рабочий лист
    renderToWorkingSheet(ss, CFG, displayHeader, enriched.rows);
    
    totalRecords = enriched.rows.length;
    
    logInfo_('DATA_SYNC', `Обработано и объединено ${totalRecords} записей`);
    
    return {
      success: true,
      records: totalRecords,
      error: null
    };
    
  } catch (error) {
    logError_('DATA_SYNC', 'Ошибка обработки данных AmoCRM', error);
    return {
      success: false,
      records: totalRecords,
      error: error.message
    };
  }
}

/**
 * Обработка данных форм сайта
 * @returns {Object} Результат обработки
 * @private
 */
function processWebFormsData_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let totalRecords = 0;
  
  try {
    const formsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SITE);
    
    if (!formsSheet || formsSheet.getLastRow() <= 1) {
      return {
        success: true,
        records: 0,
        error: null
      };
    }
    
    totalRecords = formsSheet.getLastRow() - 1; // Вычитаем заголовок
    
    logInfo_('DATA_SYNC', `Найдено ${totalRecords} заявок с сайта`);
    
    return {
      success: true,
      records: totalRecords,
      error: null
    };
    
  } catch (error) {
    logError_('DATA_SYNC', 'Ошибка обработки форм сайта', error);
    return {
      success: false,
      records: totalRecords,
      error: error.message
    };
  }
}

/**
 * Обработка данных колл-трекинга
 * @returns {Object} Результат обработки
 * @private
 */
function processCallTrackingData_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let totalRecords = 0;
  
  try {
    const callsSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.CALL);
    
    if (!callsSheet || callsSheet.getLastRow() <= 1) {
      return {
        success: true,
        records: 0,
        error: null
      };
    }
    
    totalRecords = callsSheet.getLastRow() - 1; // Вычитаем заголовок
    
    logInfo_('DATA_SYNC', `Найдено ${totalRecords} записей колл-трекинга`);
    
    return {
      success: true,
      records: totalRecords,
      error: null
    };
    
  } catch (error) {
    logError_('DATA_SYNC', 'Ошибка обработки колл-трекинга', error);
    return {
      success: false,
      records: totalRecords,
      error: error.message
    };
  }
}

/**
 * Объединение и обогащение всех данных
 * @returns {Object} Результат обработки
 * @private
 */
function mergeAndEnrichData_() {
  try {
    // Основная логика объединения уже выполняется в processAmoCrmSheetsData_()
    // Здесь можно добавить дополнительные этапы обработки
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const workingSheet = ss.getSheetByName(CONFIG.SHEETS.OUT);
    
    if (!workingSheet) {
      throw new Error('Не найден рабочий лист с объединёнными данными');
    }
    
    const totalRecords = Math.max(0, workingSheet.getLastRow() - 1);
    
    // Обновляем время в колонке TIME если есть
    try {
      updateTimeOnlyOnWorking();
    } catch (timeError) {
      logWarning_('DATA_SYNC', 'Не удалось обновить время в колонке TIME', timeError);
    }
    
    // Обновляем источники телефонов из колл-трекинга
    try {
      updateCalltrackingOnWorking();
    } catch (callError) {
      logWarning_('DATA_SYNC', 'Не удалось обновить источники звонков', callError);
    }
    
    return {
      success: true,
      records: totalRecords,
      error: null
    };
    
  } catch (error) {
    logError_('DATA_SYNC', 'Ошибка объединения данных', error);
    return {
      success: false,
      records: 0,
      error: error.message
    };
  }
}

/**
 * Отправка уведомления об ошибках
 * @param {Object} results - Результаты синхронизации
 * @private
 */
function sendErrorNotification_(results) {
  const errors = Object.entries(results)
    .filter(([key, result]) => result.error)
    .map(([key, result]) => `${key}: ${result.error}`)
    .join('\n');
  
  if (errors && CONFIG.EMAIL_NOTIFICATIONS && CONFIG.EMAIL_NOTIFICATIONS.RECIPIENTS) {
    try {
      MailApp.sendEmail({
        to: CONFIG.EMAIL_NOTIFICATIONS.RECIPIENTS.join(','),
        subject: '⚠️ Ошибки обработки данных',
        body: `Обнаружены ошибки при обработке данных:\n\n${errors}\n\nВремя: ${getCurrentDateMoscow_().toLocaleString()}`
      });
    } catch (emailError) {
      logError_('DATA_SYNC', 'Ошибка отправки email уведомления', emailError);
    }
  }
}

/**
 * Только обработка данных AmoCRM из таблиц (для меню)
 */
function syncAmoCrmDataOnly() {
  return processAmoCrmSheetsData_();
}

/**
 * Только обработка форм сайта (для меню)
 */
function syncWebFormsDataOnly() {
  return processWebFormsData_();
}

/**
 * Только обработка колл-трекинга (для меню)
 */
function syncCallTrackingDataOnly() {
  return processCallTrackingData_();
}
