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
    logInfo_('DATA_SYNC', 'Начало обработки данных из листов AmoCRM');
    
    // 1. Читаем основные таблицы AmoCRM
    const amoSheet = spreadsheet.getSheetByName('Амо Выгрузка');
    const amoFullSheet = spreadsheet.getSheetByName('Выгрузка Амо Полная');
    
    if (!amoSheet && !amoFullSheet) {
      logWarning_('DATA_SYNC', 'Не найдены листы AmoCRM для обработки');
      return { success: true, records: 0, error: null };
    }
    
    // 2. Объединяем данные из двух источников AmoCRM
    const unifiedData = unifyAmoCrmData_(amoSheet, amoFullSheet);
    
    // 3. Обогащаем данными из дополнительных источников
    const enrichedData = enrichWithAdditionalData_(unifiedData);
    
    // 4. Сохраняем в рабочий лист с красивым оформлением
    const workingSheet = getOrCreateSheet_('РАБОЧИЙ АМО');
    workingSheet.clear();
    
    if (enrichedData.header.length > 0) {
      workingSheet.getRange(1, 1, 1, enrichedData.header.length).setValues([enrichedData.header]);
      
      if (enrichedData.rows.length > 0) {
        workingSheet.getRange(2, 1, enrichedData.rows.length, enrichedData.header.length).setValues(enrichedData.rows);
        
        // Применяем красивое оформление с группировкой по блокам
        applyWorkingAmoBeautifulStyle_(workingSheet, enrichedData.header, enrichedData.rows.length);
      } else {
        // Применяем стиль только к заголовкам, если нет данных
        applyHeaderStyle_(workingSheet.getRange(1, 1, 1, enrichedData.header.length));
      }
    }
    
    totalRecords = enrichedData.rows.length;
    
    logInfo_('DATA_SYNC', `Обработано и объединено ${totalRecords} записей AmoCRM`);
    
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
 * Объединяет данные из "Амо Выгрузка" и "Выгрузка Амо Полная"
 * @param {Sheet} amoSheet - Лист "Амо Выгрузка"  
 * @param {Sheet} amoFullSheet - Лист "Выгрузка Амо Полная"
 * @returns {Object} Объединённые данные
 * @private
 */
function unifyAmoCrmData_(amoSheet, amoFullSheet) {
  const unifiedRecords = new Map(); // Используем ID как ключ для избежания дублей
  const unifiedHeader = new Set();
  
  // Добавляем данные из "Амо Выгрузка"
  if (amoSheet && amoSheet.getLastRow() > 1) {
    const amoData = amoSheet.getDataRange().getValues();
    const amoHeader = amoData[0];
    const amoRows = amoData.slice(1);
    
    // Добавляем заголовки
    amoHeader.forEach(h => unifiedHeader.add(String(h)));
    
    // Находим индекс ID
    const idIdx = amoHeader.findIndex(h => String(h).includes('ID') || String(h).includes('id'));
    
    amoRows.forEach(row => {
      const dealId = idIdx >= 0 ? String(row[idIdx] || '') : '';
      if (dealId) {
        const recordObj = {};
        amoHeader.forEach((header, idx) => {
          recordObj[String(header)] = row[idx] || '';
        });
        unifiedRecords.set(dealId, recordObj);
      }
    });
    
    logInfo_('DATA_SYNC', `Добавлено ${amoRows.length} записей из "Амо Выгрузка"`);
  }
  
  // Добавляем/объединяем данные из "Выгрузка Амо Полная"
  if (amoFullSheet && amoFullSheet.getLastRow() > 1) {
    const fullData = amoFullSheet.getDataRange().getValues();
    const fullHeader = fullData[0];
    const fullRows = fullData.slice(1);
    
    // Добавляем новые заголовки
    fullHeader.forEach(h => unifiedHeader.add(String(h)));
    
    // Находим индекс ID
    const idIdx = fullHeader.findIndex(h => String(h).includes('ID') || String(h).includes('id'));
    
    let addedCount = 0;
    let updatedCount = 0;
    
    fullRows.forEach(row => {
      const dealId = idIdx >= 0 ? String(row[idIdx] || '') : '';
      if (dealId) {
        if (unifiedRecords.has(dealId)) {
          // Обновляем существующую запись
          const existingRecord = unifiedRecords.get(dealId);
          fullHeader.forEach((header, idx) => {
            const headerStr = String(header);
            const newValue = row[idx];
            if (newValue && String(newValue).trim() !== '') {
              existingRecord[headerStr] = newValue;
            }
          });
          updatedCount++;
        } else {
          // Добавляем новую запись
          const recordObj = {};
          fullHeader.forEach((header, idx) => {
            recordObj[String(header)] = row[idx] || '';
          });
          unifiedRecords.set(dealId, recordObj);
          addedCount++;
        }
      }
    });
    
    logInfo_('DATA_SYNC', `Из "Выгрузка Амо Полная": добавлено ${addedCount}, обновлено ${updatedCount} записей`);
  }
  
  // Преобразуем в формат таблицы
  const headerArray = Array.from(unifiedHeader);
  const rowsArray = Array.from(unifiedRecords.values()).map(record => {
    return headerArray.map(header => record[header] || '');
  });
  
  return {
    header: headerArray,
    rows: rowsArray
  };
}

/**
 * Обогащает данные AmoCRM данными из дополнительных источников
 * @param {Object} amocrmData - Объединённые данные AmoCRM
 * @returns {Object} Обогащённые данные
 * @private
 */
function enrichWithAdditionalData_(amocrmData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Читаем дополнительные источники
    const reservesSheet = spreadsheet.getSheetByName('Reserves RP');
    const guestsSheet = spreadsheet.getSheetByName('Guests RP');
    const formsSheet = spreadsheet.getSheetByName('Заявки с Сайта'); 
    const callsSheet = spreadsheet.getSheetByName('КоллТрекинг');
    
    // Создаём карты для быстрого поиска
    const reservesMap = buildReservesMap_(reservesSheet);
    const guestsMap = buildGuestsMap_(guestsSheet);
    const formsMap = buildFormsMap_(formsSheet);
    const callsMap = buildCallsMap_(callsSheet);
    
    // Добавляем новые столбцы в заголовок
    const enrichedHeader = [...amocrmData.header];
    const additionalColumns = [
      'Res.Визиты', 'Res.Сумма', 'Res.Статус',
      'Gue.Визиты', 'Gue.Общая сумма', 'Gue.Первый визит',
      'Источник ТЕЛ', 'Канал ТЕЛ',
      'TIME', 'Время обработки'
    ];
    
    additionalColumns.forEach(col => {
      if (!enrichedHeader.includes(col)) {
        enrichedHeader.push(col);
      }
    });
    
    // Находим индексы ключевых полей
    const phoneIdx = findColumnIndex(amocrmData.header, ['Телефон', 'Контакт.Телефон', 'Phone']);
    const statusIdx = findColumnIndex(amocrmData.header, ['Статус', 'Status', 'Сделка.Статус']);
    const timeNow = getCurrentDateMoscow_().toLocaleString();
    
    // Обогащаем каждую строку
    const enrichedRows = amocrmData.rows.map(row => {
      const enrichedRow = [...row];
      
      // Дополняем до нужной длины
      while (enrichedRow.length < enrichedHeader.length) {
        enrichedRow.push('');
      }
      
      // Получаем телефон для поиска
      const phone = phoneIdx >= 0 ? cleanPhone(row[phoneIdx]) : '';
      
      // Нормализуем статус (убираем "ВСЕ БАРЫ СЕТИ /")
      if (statusIdx >= 0 && row[statusIdx]) {
        enrichedRow[statusIdx] = normalizeStatus(row[statusIdx]);
      }
      
      if (phone) {
        // Обогащение данными Reserves
        if (reservesMap.has(phone)) {
          const resData = reservesMap.get(phone);
          enrichedRow[enrichedHeader.indexOf('Res.Визиты')] = resData.visits || '';
          enrichedRow[enrichedHeader.indexOf('Res.Сумма')] = resData.sum || '';
          enrichedRow[enrichedHeader.indexOf('Res.Статус')] = resData.status || '';
        }
        
        // Обогащение данными Guests
        if (guestsMap.has(phone)) {
          const gueData = guestsMap.get(phone);
          enrichedRow[enrichedHeader.indexOf('Gue.Визиты')] = gueData.visits || '';
          enrichedRow[enrichedHeader.indexOf('Gue.Общая сумма')] = gueData.totalSum || '';
          enrichedRow[enrichedHeader.indexOf('Gue.Первый визит')] = gueData.firstVisit || '';
        }
        
        // Обогащение колл-трекингом
        if (callsMap.has(phone)) {
          const callData = callsMap.get(phone);
          enrichedRow[enrichedHeader.indexOf('Источник ТЕЛ')] = callData.source || '';
          enrichedRow[enrichedHeader.indexOf('Канал ТЕЛ')] = callData.channel || '';
        }
      }
      
      // Добавляем время обработки
      enrichedRow[enrichedHeader.indexOf('TIME')] = timeNow;
      enrichedRow[enrichedHeader.indexOf('Время обработки')] = timeNow;
      
      return enrichedRow;
    });
    
    logInfo_('DATA_SYNC', `Данные обогащены дополнительной информацией`);
    
    return {
      header: enrichedHeader,
      rows: enrichedRows
    };
    
  } catch (error) {
    logError_('DATA_SYNC', 'Ошибка обогащения данных', error);
    return amocrmData; // Возвращаем исходные данные при ошибке
  }
}

/**
 * Создаёт карту данных из Reserves RP
 * @private
 */
function buildReservesMap_(sheet) {
  const map = new Map();
  if (!sheet || sheet.getLastRow() <= 1) return map;
  
  try {
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    const phoneIdx = header.findIndex(h => String(h).toLowerCase().includes('телефон'));
    
    if (phoneIdx >= 0) {
      data.slice(1).forEach(row => {
        const phone = cleanPhone(row[phoneIdx]);
        if (phone) {
          map.set(phone, {
            visits: row[header.findIndex(h => String(h).includes('заявки'))] || '',
            sum: row[header.findIndex(h => String(h).includes('Счёт'))] || '',
            status: row[header.findIndex(h => String(h).includes('Статус'))] || ''
          });
        }
      });
    }
  } catch (error) {
    logWarning_('DATA_SYNC', 'Ошибка построения карты Reserves', error);
  }
  
  return map;
}

/**
 * Создаёт карту данных из Guests RP  
 * @private
 */
function buildGuestsMap_(sheet) {
  const map = new Map();
  if (!sheet || sheet.getLastRow() <= 1) return map;
  
  try {
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    const phoneIdx = header.findIndex(h => String(h).toLowerCase().includes('телефон'));
    
    if (phoneIdx >= 0) {
      data.slice(1).forEach(row => {
        const phone = cleanPhone(row[phoneIdx]);
        if (phone) {
          map.set(phone, {
            visits: row[header.findIndex(h => String(h).includes('визитов'))] || '',
            totalSum: row[header.findIndex(h => String(h).includes('Общая сумма'))] || '',
            firstVisit: row[header.findIndex(h => String(h).includes('Первый визит'))] || ''
          });
        }
      });
    }
  } catch (error) {
    logWarning_('DATA_SYNC', 'Ошибка построения карты Guests', error);
  }
  
  return map;
}

/**
 * Создаёт карту данных из форм сайта
 * @private  
 */
function buildFormsMap_(sheet) {
  const map = new Map();
  if (!sheet || sheet.getLastRow() <= 1) return map;
  
  // Пока что возвращаем пустую карту, логика будет добавлена позже
  return map;
}

/**
 * Создаёт карту колл-трекинга
 * @private
 */
function buildCallsMap_(sheet) {
  const map = new Map();
  if (!sheet || sheet.getLastRow() <= 1) return map;
  
  try {
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    
    // Ищем индексы столбцов по структуре: A-Номер линии, B-Источник, C-Канал
    data.slice(1).forEach(row => {
      const phoneNumber = String(row[0] || '').trim(); // A - Номер линии MANGO OFFICE
      const source = String(row[1] || '').trim();      // B - R.Источник ТЕЛ сделки  
      const channel = String(row[2] || '').trim();     // C - Название Канала
      
      if (phoneNumber && source) {
        map.set(phoneNumber, {
          source: source,
          channel: channel
        });
      }
    });
  } catch (error) {
    logWarning_('DATA_SYNC', 'Ошибка построения карты колл-трекинга', error);
  }
  
  return map;
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
    const formsSheet = spreadsheet.getSheetByName('Заявки с Сайта');
    
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
    const callsSheet = spreadsheet.getSheetByName('КоллТрекинг');
    
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
    const workingSheet = ss.getSheetByName('РАБОЧИЙ АМО');
    
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

/**
 * Применяет красивое оформление к существующему листу "РАБОЧИЙ АМО" 
 * Можно вызвать отдельно для обновления стилей без пересоздания данных
 */
function applyBeautifulStyleToWorkingAmo() {
  try {
    logInfo_('STYLE', 'Начало применения красивого оформления к РАБОЧИЙ АМО');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
    
    if (!workingSheet) {
      throw new Error('Лист "РАБОЧИЙ АМО" не найден. Сначала выполните синхронизацию данных.');
    }
    
    const lastRow = workingSheet.getLastRow();
    const lastCol = workingSheet.getLastColumn();
    
    if (lastRow < 1 || lastCol < 1) {
      throw new Error('Лист "РАБОЧИЙ АМО" пуст. Сначала загрузите данные.');
    }
    
    // Получаем заголовки
    const header = workingSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dataRows = lastRow - 1; // Вычитаем строку заголовков
    
    // Применяем красивое оформление
    applyWorkingAmoBeautifulStyle_(workingSheet, header, dataRows);
    
    logInfo_('STYLE', `Красивое оформление применено к листу с ${lastCol} столбцами и ${dataRows} строками данных`);
    
    return {
      success: true,
      message: `✨ Красивое оформление применено к листу "РАБОЧИЙ АМО"!\n📊 Обработано: ${lastCol} столбцов, ${dataRows} строк\n🎨 Добавлено: группировка блоков, чередование цветов, обводка`,
      columns: lastCol,
      rows: dataRows
    };
    
  } catch (error) {
    logError_('STYLE', 'Ошибка применения красивого оформления', error);
    return {
      success: false,
      message: `❌ Ошибка: ${error.message}`,
      error: error.message
    };
  }
}
