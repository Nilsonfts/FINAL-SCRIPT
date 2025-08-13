/**
 * ============================================
 * ГЛАВНЫЙ СКРИПТ - ЧИСТАЯ ВЕРСИЯ
 * Фокус: создание файла РАБОЧИЙ АМО
 * ============================================
 */

/**
 * ОСНОВНАЯ ФУНКЦИЯ - объединяет данные из Амо Выгрузка + Выгрузка Амо Полная
 */
function runMergedAmoAnalysis() {
  console.log('🚀 ОБЪЕДИНЕННАЯ ВЕРСИЯ: Запуск создания файла РАБОЧИЙ АМО');
  
  const startTime = new Date();
  
  try {
    // Создаем файл РАБОЧИЙ АМО из объединенных источников
    buildMergedAmoFile();
    
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    console.log(`✅ ОБЪЕДИНЕННАЯ ВЕРСИЯ: Файл РАБОЧИЙ АМО с цветовым форматированием создан за ${duration} сек`);
    
  } catch (error) {
    console.error('❌ ОБЪЕДИНЕННАЯ ВЕРСИЯ: Ошибка создания РАБОЧИЙ АМО:', error);
    
    // Попытка аварийного восстановления
    console.log('🔧 Попытка аварийного восстановления...');
    try {
      repairWorkingAmoSheet();
      console.log('✅ Аварийное восстановление выполнено - попробуйте снова');
    } catch (repairError) {
      console.error('❌ Аварийное восстановление не удалось:', repairError);
    }
    
    throw error;
  }
}

/**
 * ОСНОВНАЯ ФУНКЦИЯ - создает файл РАБОЧИЙ АМО (ЧИСТАЯ ВЕРСИЯ)
 */
function runWorkingAmoClean() {
  console.log('🚀 ЧИСТАЯ ВЕРСИЯ: Запуск создания файла РАБОЧИЙ АМО');
  
  const startTime = new Date();
  
  try {
    // Создаем файл РАБОЧИЙ АМО
    buildWorkingAmoFileClean();
    
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    console.log(`✅ ЧИСТАЯ ВЕРСИЯ: Файл РАБОЧИЙ АМО создан за ${duration} сек`);
    
  } catch (error) {
    console.error('❌ ЧИСТАЯ ВЕРСИЯ: Ошибка создания РАБОЧИЙ АМО:', error);
    throw error;
  }
}

/**
 * ОСНОВНАЯ ФУНКЦИЯ - создает файл РАБОЧИЙ АМО
 */
function runFullAnalyticsUpdate() {
  console.log('🚀 Запуск создания файла РАБОЧИЙ АМО');
  
  const startTime = new Date();
  
  try {
    // Создаем файл РАБОЧИЙ АМО
    buildWorkingAmoFile();
    
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    console.log(`✅ Файл РАБОЧИЙ АМО создан за ${duration} сек`);
    
  } catch (error) {
    console.error('❌ Ошибка создания РАБОЧИЙ АМО:', error);
    throw error;
  }
}

/**
 * СОЗДАНИЕ ОБЪЕДИНЕННОГО ФАЙЛА РАБОЧИЙ АМО
 */
function buildMergedAmoFile() {
  console.log('📊 ОБЪЕДИНЕННАЯ ВЕРСИЯ: Создание файла РАБОЧИЙ АМО...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Создаем или очищаем лист
    let sheet = ss.getSheetByName('РАБОЧИЙ АМО');
    if (!sheet) {
      sheet = ss.insertSheet('РАБОЧИЙ АМО');
    } else {
      sheet.clear();
    }
    
    // Создаем заголовки для объединенной структуры
    createMergedAmoHeaders_(sheet);
    
    // Загружаем и объединяем данные из двух источников
    const mergedData = loadAndMergeAmoSources_();
    
    // Записываем данные в лист
    if (mergedData.length > 0) {
      const range = sheet.getRange(2, 1, mergedData.length, mergedData[0].length);
      range.setValues(mergedData);
      
      // Автоподбор ширины колонок для первых колонок
      sheet.autoResizeColumns(1, Math.min(15, mergedData[0].length));
    }
    
    console.log(`✅ ОБЪЕДИНЕННАЯ ВЕРСИЯ: Записано ${mergedData.length} строк в РАБОЧИЙ АМО`);
    
    // Записываем лог
    writeProcessingLog_('MERGED_AMO', mergedData.length);
    
  } catch (error) {
    console.error('❌ ОБЪЕДИНЕННАЯ ВЕРСИЯ: Ошибка сборки РАБОЧИЙ АМО:', error);
    throw error;
  }
}

/**
 * СОЗДАНИЕ ФАЙЛА РАБОЧИЙ АМО (ЧИСТАЯ ВЕРСИЯ)
 */
function buildWorkingAmoFileClean() {
  console.log('📊 ЧИСТАЯ ВЕРСИЯ: Создание файла РАБОЧИЙ АМО...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Создаем или очищаем лист
    let sheet = ss.getSheetByName('РАБОЧИЙ АМО');
    if (!sheet) {
      sheet = ss.insertSheet('РАБОЧИЙ АМО');
    } else {
      sheet.clear();
    }
    
    // Создаем заголовки
    createWorkingAmoHeaders_(sheet);
    
    // Загружаем и обрабатываем данные
    const consolidatedData = loadAndConsolidateAllDataClean_();
    
    // Записываем данные в лист
    if (consolidatedData.length > 0) {
      const range = sheet.getRange(2, 1, consolidatedData.length, consolidatedData[0].length);
      range.setValues(consolidatedData);
      
      // Автоподбор ширины колонок для первых колонок
      sheet.autoResizeColumns(1, Math.min(10, consolidatedData[0].length));
    }
    
    console.log(`✅ ЧИСТАЯ ВЕРСИЯ: Записано ${consolidatedData.length} строк в РАБОЧИЙ АМО`);
    
  } catch (error) {
    console.error('❌ ЧИСТАЯ ВЕРСИЯ: Ошибка сборки РАБОЧИЙ АМО:', error);
    throw error;
  }
}

/**
 * СОЗДАНИЕ ФАЙЛА РАБОЧИЙ АМО
 */
function buildWorkingAmoFile() {
  console.log('📊 Создание файла РАБОЧИЙ АМО...');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Создаем или очищаем лист
    let sheet = ss.getSheetByName('РАБОЧИЙ АМО');
    if (!sheet) {
      sheet = ss.insertSheet('РАБОЧИЙ АМО');
    } else {
      sheet.clear();
      // Очищаем все правила проверки данных
      clearDataValidation_(sheet);
    }
    
    // Создаем заголовки
    createWorkingAmoHeaders_(sheet);
    
    // Загружаем и обрабатываем данные
    const consolidatedData = loadAndConsolidateAllData_();
    
    // Записываем данные в лист
    if (consolidatedData.length > 0) {
      const range = sheet.getRange(2, 1, consolidatedData.length, consolidatedData[0].length);
      range.setValues(consolidatedData);
      
      // Автоподбор ширины колонок для первых колонок
      sheet.autoResizeColumns(1, Math.min(10, consolidatedData[0].length));
    }
    
    console.log(`✅ Записано ${consolidatedData.length} строк в РАБОЧИЙ АМО`);
    
  } catch (error) {
    console.error('❌ Ошибка сборки РАБОЧИЙ АМО:', error);
    throw error;
  }
}

/**
 * ЗАГРУЗКА И ОБЪЕДИНЕНИЕ ВСЕХ ДАННЫХ (ЧИСТАЯ ВЕРСИЯ)
 */
function loadAndConsolidateAllDataClean_() {
  console.log('🔄 ЧИСТАЯ ВЕРСИЯ: Загрузка данных из всех источников...');
  
  try {
    // Загружаем данные из всех листов
    const amoCrmData = loadAmoCrmData_();
    const reservesData = loadReservesData_();
    const guestsData = loadGuestsData_();
    const siteFormsData = loadSiteFormsData_();
    const calltrackingData = loadCalltrackingData_();
    
    console.log(`📈 ЧИСТАЯ ВЕРСИЯ: Данные загружены: AMO=${amoCrmData.length}, Reserves=${reservesData.length}, Guests=${guestsData.length}, Site=${siteFormsData.length}, Calls=${calltrackingData.length}`);
    
    // Создаем карты для быстрого поиска по телефону
    const phoneToSite = createPhoneMap_(siteFormsData);
    const phoneToReserves = createPhoneMap_(reservesData);
    const phoneToGuests = createPhoneMap_(guestsData);
    const phoneToCalltracking = createPhoneMap_(calltrackingData);
    
    // Объединяем данные на основе AmoCRM
    const consolidatedData = [];
    
    amoCrmData.forEach((deal, index) => {
      // Обрабатываем ВСЕ записи, даже без телефона
      const phone = deal.phone ? normalizePhone_(deal.phone) : null;
      
      // Получаем обогащающие данные (только если есть телефон)
      const siteData = phone ? (phoneToSite.get(phone) || {}) : {};
      const reserveData = phone ? (phoneToReserves.get(phone) || {}) : {};
      const guestData = phone ? (phoneToGuests.get(phone) || {}) : {};
      const callData = phone ? (phoneToCalltracking.get(phone) || {}) : {};
      
      // Отладочная информация для первых записей
      if (index < 3) {
        console.log(`ЧИСТАЯ ВЕРСИЯ - Запись ${index}: created_at=${deal.created_at}, тип=${typeof deal.created_at}, booking_date=${deal.booking_date}, тип=${typeof deal.booking_date}, phone=${deal.phone || 'НЕТ'}`);
      }
      
      // Создаем объединенную строку данных с чистыми функциями
      const enrichedRow = createEnrichedDataRowClean_(deal, siteData, reserveData, guestData, callData);
      consolidatedData.push(enrichedRow);
    });
    
    console.log(`🎯 ЧИСТАЯ ВЕРСИЯ: Создано ${consolidatedData.length} обогащенных записей`);
    return consolidatedData;
    
  } catch (error) {
    console.error('❌ ЧИСТАЯ ВЕРСИЯ: Ошибка объединения данных:', error);
    throw error;
  }
}

/**
 * ЗАГРУЗКА И ОБЪЕДИНЕНИЕ ВСЕХ ДАННЫХ
 */
function loadAndConsolidateAllData_() {
  console.log('🔄 Загрузка данных из всех источников...');
  
  try {
    // Загружаем данные из всех листов
    const amoCrmData = loadAmoCrmData_();
    const reservesData = loadReservesData_();
    const guestsData = loadGuestsData_();
    const siteFormsData = loadSiteFormsData_();
    const calltrackingData = loadCalltrackingData_();
    
    console.log(`📈 Данные загружены: AMO=${amoCrmData.length}, Reserves=${reservesData.length}, Guests=${guestsData.length}, Site=${siteFormsData.length}, Calls=${calltrackingData.length}`);
    
    // Создаем карты для быстрого поиска по телефону
    const phoneToSite = createPhoneMap_(siteFormsData);
    const phoneToReserves = createPhoneMap_(reservesData);
    const phoneToGuests = createPhoneMap_(guestsData);
    const phoneToCalltracking = createPhoneMap_(calltrackingData);
    
    // Объединяем данные на основе AmoCRM
    const consolidatedData = [];
    
    amoCrmData.forEach((deal, index) => {
      if (!deal.phone) return; // Пропускаем записи без телефона
      
      const phone = normalizePhone_(deal.phone);
      
      // Получаем обогащающие данные
      const siteData = phoneToSite.get(phone) || {};
      const reserveData = phoneToReserves.get(phone) || {};
      const guestData = phoneToGuests.get(phone) || {};
      const callData = phoneToCalltracking.get(phone) || {};
      
      // Отладочная информация для первых записей
      if (index < 3) {
        console.log(`Запись ${index}: created_at=${deal.created_at}, тип=${typeof deal.created_at}, booking_date=${deal.booking_date}, тип=${typeof deal.booking_date}`);
      }
      
      // Создаем объединенную строку данных
      const enrichedRow = createEnrichedDataRow_(deal, siteData, reserveData, guestData, callData);
      consolidatedData.push(enrichedRow);
    });
    
    console.log(`🎯 Создано ${consolidatedData.length} обогащенных записей`);
    return consolidatedData;
    
  } catch (error) {
    console.error('❌ Ошибка объединения данных:', error);
    throw error;
  }
}

/**
 * ЗАГРУЗКА И ОБЪЕДИНЕНИЕ ДАННЫХ ИЗ ДВУХ ИСТОЧНИКОВ AMO
 */
function loadAndMergeAmoSources_() {
  console.log('🔄 ОБЪЕДИНЕНИЕ: Загрузка данных из обеих источников AMO...');
  
  try {
    // Загружаем данные из двух источников
    const fullAmoData = loadAmoFullData_();
    const regularAmoData = loadAmoRegularData_();
    
    // Загружаем данные гостей для расчета счетов
    const guestsData = loadGuestsData_();
    const phoneToGuests = createPhoneMap_(guestsData);
    
    console.log(`📊 ОБЪЕДИНЕНИЕ: Загружено Full=${fullAmoData.length}, Regular=${regularAmoData.length}, Guests=${guestsData.length}`);
    
    // Создаем карту по ID для "Выгрузка Амо Полная" (приоритет)
    const fullAmoMap = new Map();
    fullAmoData.forEach(deal => {
      if (deal.id) {
        fullAmoMap.set(deal.id.toString(), deal);
      }
    });
    
    // Объединенные данные
    const mergedDeals = new Map();
    
    // Сначала добавляем все из "Полной" (приоритет)
    fullAmoData.forEach(deal => {
      if (deal.id) {
        const id = deal.id.toString();
        // Ищем данные гостя по телефону
        const phone = deal.phone ? normalizePhoneForSearch_(deal.phone) : null;
        const guestData = phone ? phoneToGuests.get(phone) : null;
        
        // Отладочная информация для сопоставления телефонов
        if (phone && guestData) {
          console.log(`🔗 Найдено сопоставление: AMO ${deal.phone} → Гость ${guestData.phone} (${guestData.name})`);
        }
        
        mergedDeals.set(id, createMergedDealRow_(deal, null, guestData));
      }
    });
    
    // Затем дополняем пустоты из "Амо Выгрузка"
    regularAmoData.forEach(deal => {
      if (!deal.id) return;
      
      const id = deal.id.toString();
      const phone = deal.phone ? normalizePhoneForSearch_(deal.phone) : null;
      const guestData = phone ? phoneToGuests.get(phone) : null;
      
      // Отладочная информация для сопоставления телефонов
      if (phone && guestData) {
        console.log(`🔗 Дополнительное сопоставление: AMO ${deal.phone} → Гость ${guestData.phone} (${guestData.name})`);
      }
      
      if (mergedDeals.has(id)) {
        // ID уже есть - дополняем пустые поля
        const existingDeal = fullAmoMap.get(id);
        const updatedDeal = createMergedDealRow_(existingDeal, deal, guestData);
        mergedDeals.set(id, updatedDeal);
      } else {
        // Новый ID - добавляем как есть
        mergedDeals.set(id, createMergedDealRow_(null, deal, guestData));
      }
    });
    
    // Преобразуем в массив для записи
    const mergedArray = Array.from(mergedDeals.values());
    
    // Сортируем по дате создания (новые сверху)
    mergedArray.sort((a, b) => {
      const dateA = parseRussianDate_(a[9]); // J - Дата создания (сдвинулось)
      const dateB = parseRussianDate_(b[9]);
      return dateB - dateA;
    });
    
    console.log(`🎯 ОБЪЕДИНЕНИЕ: Создано ${mergedArray.length} объединенных записей с расчетом счетов`);
    return mergedArray;
    
  } catch (error) {
    console.error('❌ ОБЪЕДИНЕНИЕ: Ошибка объединения AMO источников:', error);
    throw error;
  }
}
/**
 * ЗАГРУЗКА ДАННЫХ ИЗ "ВЫГРУЗКА АМО ПОЛНАЯ" (ПРИОРИТЕТ)
 */
function loadAmoFullData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Выгрузка Амо Полная');
  
  if (!sheet) {
    console.warn('⚠️ Лист "Выгрузка Амо Полная" не найден');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const rows = data.slice(1); // Пропускаем заголовки
  
  return rows.map(row => ({
    // Маппинг согласно структуре "Выгрузка Амо Полная"
    id: row[0],                    // A — Сделка.ID
    name: row[1],                  // B — Сделка.Название  
    responsible: row[2],           // C — Сделка.Ответственный
    contact_name: row[3],          // D — Контакт.ФИО
    status: row[4],                // E — Сделка.Статус
    budget: row[5],                // F — Сделка.Бюджет
    created_at: row[6],            // G — Сделка.Дата создания
    responsible2: row[7],          // H — Сделка.Ответственный (дубль)
    tags: row[8],                  // I — Сделка.Теги
    closed_at: row[9],             // J — Сделка.Дата закрытия
    
    // Аналитика
    ym_client_id: row[10],         // K — YM_CLIENT_ID
    ga_client_id: row[11],         // L — GA_CLIENT_ID
    button_text: row[12],          // M — BUTTON_TEXT
    date: row[13],                 // N — DATE
    time: row[14],                 // O — TIME
    deal_source: row[15],          // P — R.Источник сделки
    city_tag: row[16],             // Q — R.Тег города
    software: row[17],             // R — ПО
    
    // Бронирование
    bar_name: row[18],             // S — Бар (deal)
    booking_date: row[19],         // T — Дата брони
    guest_count: row[20],          // U — Кол-во гостей
    visit_time: row[21],           // V — Время прихода
    comment: row[22],              // W — Комментарий МОБ
    source: row[23],               // X — Источник
    lead_type: row[24],            // Y — Тип лида
    refusal_reason: row[25],       // Z — Причина отказа (ОБ)
    guest_status: row[26],         // AA — R.Статусы гостей
    referral_type: row[27],        // AB — Сарафан гости
    
    // UTM данные
    utm_medium: row[28],           // AC — UTM_MEDIUM
    formname: row[29],             // AD — FORMNAME
    referer: row[30],              // AE — REFERER
    formid: row[31],               // AF — FORMID
    mango_line1: row[32],          // AG — Номер линии MANGO OFFICE
    utm_source: row[33],           // AH — UTM_SOURCE
    utm_term: row[34],             // AI — UTM_TERM
    utm_campaign: row[35],         // AJ — UTM_CAMPAIGN
    utm_content: row[36],          // AK — UTM_CONTENT
    utm_referrer: row[37],         // AL — utm_referrer
    _ym_uid: row[38],              // AM — _ym_uid
    
    // Контакты  
    phone: row[39],                // AN — Контакт.Телефон
    mango_line2: row[40],          // AO — Контакт.Номер линии MANGO OFFICE
    notes: row[41]                 // AP — Примечания
  }));
}

/**
 * ЗАГРУЗКА ДАННЫХ ИЗ "АМО ВЫГРУЗКА" (ДОПОЛНЕНИЕ)
 */
function loadAmoRegularData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Амо Выгрузка');
  
  if (!sheet) {
    console.warn('⚠️ Лист "Амо Выгрузка" не найден');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const rows = data.slice(1); // Пропускаем заголовки
  
  return rows.map(row => ({
    // Маппинг согласно структуре "Амо Выгрузка" (A→AO)
    id: row[0],                    // A — Сделка.ID
    name: row[1],                  // B — Сделка.Название
    responsible: row[2],           // C — Сделка.Ответственный  
    status: row[3],                // D — Сделка.Статус
    budget: row[4],                // E — Сделка.Бюджет
    created_at: row[5],            // F — Сделка.Дата создания
    tags: row[6],                  // G — Сделка.Теги
    closed_at: row[7],             // H — Сделка.Дата закрытия
    
    // Аналитика
    ym_client_id: row[8],          // I — YM_CLIENT_ID
    ga_client_id: row[9],          // J — GA_CLIENT_ID
    button_text: row[10],          // K — BUTTON_TEXT
    date: row[11],                 // L — DATE
    time: row[12],                 // M — TIME
    deal_source: row[13],          // N — R.Источник сделки
    city_tag: row[14],             // O — R.Тег города
    software: row[15],             // P — ПО
    
    // Бронирование  
    bar_name: row[16],             // Q — Бар (deal)
    booking_date: row[17],         // R — Дата брони
    guest_count: row[18],          // S — Кол-во гостей
    visit_time: row[19],           // T — Время прихода
    comment: row[20],              // U — Комментарий МОБ
    source: row[21],               // V — Источник
    lead_type: row[22],            // W — Тип лида
    refusal_reason: row[23],       // X — Причина отказа (ОБ)
    guest_status: row[24],         // Y — R.Статусы гостей
    referral_type: row[25],        // Z — Сарафан гости
    
    // UTM данные
    utm_medium: row[26],           // AA — UTM_MEDIUM
    formname: row[27],             // AB — FORMNAME
    referer: row[28],              // AC — REFERER
    formid: row[29],               // AD — FORMID
    mango_line1: row[30],          // AE — Номер линии MANGO OFFICE
    utm_source: row[31],           // AF — UTM_SOURCE
    utm_term: row[32],             // AG — UTM_TERM
    utm_campaign: row[33],         // AH — UTM_CAMPAIGN
    utm_content: row[34],          // AI — UTM_CONTENT
    utm_referrer: row[35],         // AJ — utm_referrer
    _ym_uid: row[36],              // AK — _ym_uid
    notes: row[37],                // AL — Примечания(через ;)
    
    // Контакты
    contact_name: row[38],         // AM — Контакт.ФИО
    phone: row[39],                // AN — Контакт.Телефон
    mango_line2: row[40]           // AO — Контакт.Номер линии MANGO OFFICE
  }));
}

/**
 * ЗАГРУЗКА ДАННЫХ AMO CRM (СТАРАЯ ВЕРСИЯ - СОВМЕСТИМОСТЬ)
 */
function loadAmoCrmData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Приоритет "Выгрузка Амо Полная"
  let sheet = ss.getSheetByName('Выгрузка Амо Полная') || 
             ss.getSheetByName('Амо Выгрузка');
  
  if (!sheet) {
    console.warn('⚠️ Лист AmoCRM не найден');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const rows = data.slice(1); // Пропускаем заголовки
  
  return rows.map(row => ({
    // Основная информация (по структуре "Выгрузка Амо Полная")
    id: row[0],                    // A — Сделка.ID
    name: row[1],                  // B — Сделка.Название
    responsible: row[2],           // C — Сделка.Ответственный
    contact_name: row[3],          // D — Контакт.ФИО
    status: row[4],                // E — Сделка.Статус
    budget: row[5],                // F — Сделка.Бюджет
    created_at: row[6],            // G — Сделка.Дата создания
    responsible2: row[7],          // H — Сделка.Ответственный (дубль)
    tags: row[8],                  // I — Сделка.Теги
    closed_at: row[9],             // J — Сделка.Дата закрытия
    
    // Аналитика
    ym_client_id: row[10],         // K — YM_CLIENT_ID
    ga_client_id: row[11],         // L — GA_CLIENT_ID
    button_text: row[12],          // M — BUTTON_TEXT
    date: row[13],                 // N — DATE
    time: row[14],                 // O — TIME
    deal_source: row[15],          // P — R.Источник сделки
    city_tag: row[16],             // Q — R.Тег города
    software: row[17],             // R — ПО
    
    // Бронирование
    bar_name: row[18],             // S — Бар (deal)
    booking_date: row[19],         // T — Дата брони
    guest_count: row[20],          // U — Кол-во гостей
    visit_time: row[21],           // V — Время прихода
    comment: row[22],              // W — Комментарий МОБ
    source: row[23],               // X — Источник
    lead_type: row[24],            // Y — Тип лида
    refusal_reason: row[25],       // Z — Причина отказа (ОБ)
    guest_status: row[26],         // AA — R.Статусы гостей
    referral_type: row[27],        // AB — Сарафан гости
    
    // UTM данные
    utm_medium: row[28],           // AC — UTM_MEDIUM
    formname: row[29],             // AD — FORMNAME
    referer: row[30],              // AE — REFERER
    formid: row[31],               // AF — FORMID
    mango_line1: row[32],          // AG — Номер линии MANGO OFFICE
    utm_source: row[33],           // AH — UTM_SOURCE
    utm_term: row[34],             // AI — UTM_TERM
    utm_campaign: row[35],         // AJ — UTM_CAMPAIGN
    utm_content: row[36],          // AK — UTM_CONTENT
    utm_referrer: row[37],         // AL — utm_referrer
    _ym_uid: row[38],              // AM — _ym_uid
    
    // Контакты
    phone: row[39],                // AN — Контакт.Телефон
    mango_line2: row[40],          // AO — Контакт.Номер линии MANGO OFFICE
    notes: row[41]                 // AP — Примечания
  }));
}

/**
 * ЗАГРУЗКА РЕЗЕРВОВ
 */
function loadReservesData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Reserves RP');
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  return data.slice(1).map(row => ({
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
}

/**
 * ЗАГРУЗКА ГОСТЕЙ
 */
function loadGuestsData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Guests RP');
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  return data.slice(1).map(row => ({
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
}

/**
 * ЗАГРУЗКА ЗАЯВОК С САЙТА
 */
function loadSiteFormsData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Заявки с Сайта');
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  return data.slice(1).map(row => ({
    name: row[0],                 // A — Name
    phone: row[1],                // B — Phone
    referer: row[2],              // C — referer
    formid: row[3],               // D — formid
    email: row[6],                // G — Email
    date: row[7],                 // H — Date
    formname: row[10],            // K — Form name
    time: row[11],                // L — Time
    utm_term: row[12],            // M — utm_term
    utm_campaign: row[13],        // N — utm_campaign
    utm_source: row[14],          // O — utm_source
    utm_content: row[15],         // P — utm_content
    utm_medium: row[16],          // Q — utm_medium
    button_text: row[20],         // U — button_text
    referrer: row[21],            // V — referrer
    landing_page: row[22],        // W — landing_page
    device_type: row[25],         // Z — device_type
    user_city: row[32],           // AG — user_city
    visits_count: row[40]         // AO — visits_count
  }));
}

/**
 * ЗАГРУЗКА КОЛЛТРЕКИНГА
 */
function loadCalltrackingData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('КоллТрекинг');
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  return data.slice(1).map(row => ({
    mango_line: row[0],           // A — Контакт.Номер линии MANGO OFFICE
    tel_source: row[1],           // B — R.Источник ТЕЛ сделки
    channel_name: row[2],         // C — Название Канала
    phone: row[0]                 // Используем mango_line как телефон для поиска
  }));
}

/**
 * СОЗДАНИЕ ОБЪЕДИНЕННОЙ СТРОКИ ДАННЫХ
 */
function createMergedDealRow_(fullDeal, regularDeal, guestData = null) {
  // Функция для выбора значения: приоритет у fullDeal, но если пусто - берем из regularDeal
  const mergeValue = (fullValue, regularValue) => {
    if (isEmptyValue_(fullValue)) {
      return regularValue || '';
    }
    return fullValue || '';
  };
  
  // Функция для очистки названий от "ВСЕ БАРЫ СЕТИ /"
  const cleanName = (name) => {
    if (!name) return '';
    return String(name).replace(/^ВСЕ БАРЫ СЕТИ\s*\/\s*/, '');
  };
  
  const fullD = fullDeal || {};
  const regD = regularDeal || {};
  
  return [
    // Новый порядок согласно пользовательским требованиям
    mergeValue(fullD.id, regD.id),                                           // A — Сделка.ID
    mergeValue(fullD.name, regD.name),                                       // B — Сделка.Название
    normalizeStatusForValidation_(cleanName(mergeValue(fullD.status, regD.status))), // C — Сделка.Статус (нормализованный)
    mergeValue(fullD.refusal_reason, regD.refusal_reason),                   // D — Сделка.Причина отказа (ОБ)
    mergeValue(fullD.lead_type, regD.lead_type),                             // E — Сделка.Тип лида
    mergeValue(fullD.guest_status, regD.guest_status),                       // F — Сделка.R.Статусы гостей
    mergeValue(fullD.responsible, regD.responsible),                         // G — Сделка.Ответственный
    mergeValue(fullD.tags, regD.tags),                                       // H — Сделка.Теги
    mergeValue(fullD.budget, regD.budget),                                   // I — Сделка.Бюджет
    formatDateClean_(mergeValue(fullD.created_at, regD.created_at)),         // J — Сделка.Дата создания
    formatDateClean_(mergeValue(fullD.closed_at, regD.closed_at)),           // K — Сделка.Дата закрытия
    mergeValue(fullD.mango_line2, regD.mango_line2),                         // L — Контакт.Номер линии MANGO OFFICE
    mergeValue(fullD.mango_line1, regD.mango_line1),                         // M — Сделка.Номер линии MANGO OFFICE
    mergeValue(fullD.contact_name, regD.contact_name),                       // N — Контакт.ФИО
    normalizePhoneClean_(mergeValue(fullD.phone, regD.phone)),               // O — Контакт.Телефон
    calculateFactAmount_(fullDeal, regularDeal, guestData),                  // P — Счет факт (НОВАЯ ЛОГИКА)
    mergeValue(fullD.date, regD.date),                                       // Q — Сделка.DATE (сдвинуто)
    mergeValue(fullD.time, regD.time),                                       // R — Сделка.TIME (сдвинуто)
    mergeValue(fullD.city_tag, regD.city_tag),                               // S — Сделка.R.Тег города (сдвинуто)
    formatDateClean_(mergeValue(fullD.booking_date, regD.booking_date)),     // T — Сделка.Дата брони (сдвинуто)
    mergeValue(fullD.software, regD.software),                               // U — Сделка.ПО (сдвинуто)
    mergeValue(fullD.referral_type, regD.referral_type),                     // V — Сделка.Сарафан гости (сдвинуто)
    mergeValue(fullD.bar_name, regD.bar_name),                               // W — Сделка.Бар (deal) (сдвинуто)
    mergeValue(fullD.deal_source, regD.deal_source),                         // X — Сделка.R.Источник сделки (сдвинуто)
    mergeValue(fullD.button_text, regD.button_text),                         // Y — Сделка.BUTTON_TEXT (сдвинуто)
    mergeValue(fullD.ym_client_id, regD.ym_client_id),                       // Z — Сделка.YM_CLIENT_ID (сдвинуто)
    mergeValue(fullD.ga_client_id, regD.ga_client_id),                       // AA — Сделка.GA_CLIENT_ID (сдвинуто)
    mergeValue(fullD.utm_source, regD.utm_source),                           // AB — Сделка.UTM_SOURCE (сдвинуто)
    mergeValue(fullD.utm_medium, regD.utm_medium),                           // AC — Сделка.UTM_MEDIUM (сдвинуто)
    mergeValue(fullD.utm_term, regD.utm_term),                               // AD — Сделка.UTM_TERM (сдвинуто)
    mergeValue(fullD.utm_campaign, regD.utm_campaign),                       // AE — Сделка.UTM_CAMPAIGN (сдвинуто)
    mergeValue(fullD.utm_content, regD.utm_content),                         // AF — Сделка.UTM_CONTENT (сдвинуто)
    mergeValue(fullD.utm_referrer, regD.utm_referrer),                       // AG — Сделка.utm_referrer (сдвинуто)
    mergeValue(fullD.visit_time, regD.visit_time),                           // AH — Сделка.Время прихода (сдвинуто)
    mergeValue(fullD.comment, regD.comment),                                 // AI — Сделка.Комментарий МОБ (сдвинуто)
    mergeValue(fullD.source, regD.source),                                   // AJ — Сделка.Источник (сдвинуто)
    mergeValue(fullD.formname, regD.formname),                               // AK — Сделка.FORMNAME (сдвинуто)
    mergeValue(fullD.referer, regD.referer),                                 // AL — Сделка.REFERER (сдвинуто)
    mergeValue(fullD.formid, regD.formid),                                   // AM — Сделка.FORMID (сдвинуто)
    mergeValue(fullD._ym_uid, regD._ym_uid),                                 // AN — Сделка._ym_uid (сдвинуто)
    mergeValue(fullD.notes, regD.notes)                                      // AO — Сделка.Примечания(через ;) (сдвинуто)
  ];
}

/**
 * СОЗДАНИЕ ЗАГОЛОВКОВ ДЛЯ ОБЪЕДИНЕННОЙ ТАБЛИЦЫ
 */
function createMergedAmoHeaders_(sheet) {
  // Заголовки в новом порядке согласно пользовательским требованиям
  const headers = [
    'Сделка.ID',                          // A
    'Сделка.Название',                    // B  
    'Сделка.Статус',                      // C
    'Сделка.Причина отказа (ОБ)',         // D
    'Сделка.Тип лида',                    // E
    'Сделка.R.Статусы гостей',            // F
    'Сделка.Ответственный',               // G
    'Сделка.Теги',                        // H
    'Сделка.Бюджет',                      // I
    'Сделка.Дата создания',               // J
    'Сделка.Дата закрытия',               // K
    'Контакт.Номер линии MANGO OFFICE',   // L
    'Сделка.Номер линии MANGO OFFICE',    // M
    'Контакт.ФИО',                        // N
    'Контакт.Телефон',                    // O
    'Счет факт',                          // P - НОВАЯ КОЛОНКА
    'Сделка.DATE',                        // Q (сдвинуто)
    'Сделка.TIME',                        // R (сдвинуто)
    'Сделка.R.Тег города',                // S (сдвинуто)
    'Сделка.Дата брони',                  // T (сдвинуто)
    'Сделка.ПО',                          // U (сдвинуто)
    'Сделка.Сарафан гости',               // V (сдвинуто)
    'Сделка.Бар (deal)',                  // W (сдвинуто)
    'Сделка.R.Источник сделки',           // X (сдвинуто)
    'Сделка.BUTTON_TEXT',                 // Y (сдвинуто)
    'Сделка.YM_CLIENT_ID',                // Z (сдвинуто)
    'Сделка.GA_CLIENT_ID',                // AA (сдвинуто)
    'Сделка.UTM_SOURCE',                  // AB (сдвинуто)
    'Сделка.UTM_MEDIUM',                  // AC (сдвинуто)
    'Сделка.UTM_TERM',                    // AD (сдвинуто)
    'Сделка.UTM_CAMPAIGN',                // AE (сдвинуто)
    'Сделка.UTM_CONTENT',                 // AF (сдвинуто)
    'Сделка.utm_referrer',                // AG (сдвинуто)
    'Сделка.Время прихода',               // AH (сдвинуто)
    'Сделка.Комментарий МОБ',             // AI (сдвинуто)
    'Сделка.Источник',                    // AJ (сдвинуто)
    'Сделка.FORMNAME',                    // AK (сдвинуто)
    'Сделка.REFERER',                     // AL (сдвинуто)
    'Сделка.FORMID',                      // AM (сдвинуто)
    'Сделка._ym_uid',                     // AN (сдвинуто)
    'Сделка.Примечания(через ;)'          // AO (сдвинуто)
  ];
  
  // Записываем заголовки
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Форматирование заголовков
  headerRange.setFontWeight('bold')
             .setBackground('#e6e6e6')
             .setHorizontalAlignment('center')
             .setWrap(true);
  
  // Настройка форматирования для всех данных (кроме заголовков)
  if (sheet.getLastRow() > 1) {
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length);
    dataRange.setWrap(false); // Обрезаем все строки данных
  }
  
  // Цветовое кодирование блоков
  applyMergedSectionColoring_(sheet, headers.length);
  
  // Применяем условное форматирование к значениям
  applyConditionalFormatting_(sheet);
}

/**
 * ЦВЕТОВОЕ КОДИРОВАНИЕ ДЛЯ ОБЪЕДИНЕННОЙ ТАБЛИЦЫ
 */
function applyMergedSectionColoring_(sheet, totalCols) {
  // Основная информация о сделке (A-K, колонки 1-11) - светло-голубой
  sheet.getRange(1, 1, 1, 11).setBackground('#cfe2f3');
  
  // Контактная информация и счета (L-P, колонки 12-16) - светло-желтый  
  sheet.getRange(1, 12, 1, 5).setBackground('#fff2cc');
  
  // Аналитика и UTM (Q-AO, колонки 17-41) - светло-зеленый
  sheet.getRange(1, 17, 1, 25).setBackground('#d9ead3');
}

/**
 * ПРИМЕНЕНИЕ УСЛОВНОГО ФОРМАТИРОВАНИЯ К ЗНАЧЕНИЯМ
 */
function applyConditionalFormatting_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  try {
    // Форматирование для "Сделка.Статус" (колонка C)
    applyStatusFormatting_(sheet, 3, lastRow); // C = 3
    
    // Форматирование для "Сделка.Тип лида" (колонка E) 
    applyLeadTypeFormatting_(sheet, 5, lastRow); // E = 5
    
    // Форматирование для "Сделка.R.Статусы гостей" (колонка F)
    applyGuestStatusFormatting_(sheet, 6, lastRow); // F = 6
    
    // Форматирование для "Сделка.Номер линии MANGO OFFICE" (колонка M)
    applyMangoLineFormatting_(sheet, 13, lastRow); // M = 13
    
    console.log('✅ Условное форматирование применено');
  } catch (error) {
    console.warn('⚠️ Ошибка применения условного форматирования:', error);
  }
}

/**
 * ФОРМАТИРОВАНИЕ СТАТУСОВ СДЕЛКИ
 */
function applyStatusFormatting_(sheet, column, lastRow) {
  const statusColors = {
    'Автораспределение': '#EDEDED',
    'Взяли в работу': '#E5D5F7', 
    'вопрос к бару': '#CDE6F9',
    'Закрыто и не реализовано': '#FAD2D3',
    'Контроль оплаты': '#FFF2C2',
    'НДЗ': '#4E3B2F',
    'Оплачено': '#D6F3C6',
    'Перенос с открытой': '#004A99',
    'Сделка.Статус': '#EDEDED',
    'успешно в РП': '#D6F3C6',
    'успешно РЕАЛИЗОВАНО': '#D6F3C6'
  };
  
  const range = sheet.getRange(2, column, lastRow - 1, 1);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    const value = String(values[i][0]).trim();
    if (statusColors[value]) {
      const cellRange = sheet.getRange(i + 2, column);
      cellRange.setBackground(statusColors[value]);
      
      // Белый текст для тёмных фонов
      if (statusColors[value] === '#4E3B2F' || statusColors[value] === '#004A99') {
        cellRange.setFontColor('#FFFFFF');
      }
    }
  }
}

/**
 * ФОРМАТИРОВАНИЕ ТИПОВ ЛИДОВ
 */
function applyLeadTypeFormatting_(sheet, column, lastRow) {
  const leadTypeColors = {
    'Целевой': '#C6E0B4',
    'Нецелевой': '#F4CCCC'
  };
  
  const range = sheet.getRange(2, column, lastRow - 1, 1);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    const value = String(values[i][0]).trim();
    if (leadTypeColors[value]) {
      sheet.getRange(i + 2, column).setBackground(leadTypeColors[value]);
    }
  }
}

/**
 * ФОРМАТИРОВАНИЕ СТАТУСОВ ГОСТЕЙ
 */
function applyGuestStatusFormatting_(sheet, column, lastRow) {
  const guestStatusColors = {
    'Новый': '#E2F0D9',
    'Повторный': '#006847'
  };
  
  const range = sheet.getRange(2, column, lastRow - 1, 1);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    const value = String(values[i][0]).trim();
    if (guestStatusColors[value]) {
      const cellRange = sheet.getRange(i + 2, column);
      cellRange.setBackground(guestStatusColors[value]);
      
      // Белый текст для тёмного фона
      if (guestStatusColors[value] === '#006847') {
        cellRange.setFontColor('#FFFFFF');
      }
    }
  }
}

/**
 * ФОРМАТИРОВАНИЕ НОМЕРОВ ЛИНИЙ MANGO
 */
function applyMangoLineFormatting_(sheet, column, lastRow) {
  const mangoLineColors = {
    '78122428017': '#F4CCCC',
    '78122709071': '#990000',
    '78123172353': '#FFE599', 
    '78123177149': '#C6E0B4',
    '78123177310': '#9DC3E6'
  };
  
  const range = sheet.getRange(2, column, lastRow - 1, 1);
  const values = range.getValues();
  
  for (let i = 0; i < values.length; i++) {
    const value = String(values[i][0]).trim();
    if (mangoLineColors[value]) {
      const cellRange = sheet.getRange(i + 2, column);
      cellRange.setBackground(mangoLineColors[value]);
      
      // Белый текст для тёмного фона
      if (mangoLineColors[value] === '#990000') {
        cellRange.setFontColor('#FFFFFF');
      }
    }
  }
}

/**
 * ЗАПИСЬ ЛОГА ОБРАБОТКИ
 */
function writeProcessingLog_(operationType, recordsCount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('LOG');
    
    // Создаем лист LOG если его нет
    if (!logSheet) {
      logSheet = ss.insertSheet('LOG');
      logSheet.getRange(1, 1, 1, 4).setValues([['Дата/Время', 'Операция', 'Записей', 'Статус']]);
      logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e6e6e6');
    }
    
    // Добавляем запись в лог
    const timestamp = new Date().toLocaleString('ru-RU');
    logSheet.appendRow([timestamp, operationType, recordsCount, 'SUCCESS']);
    
  } catch (error) {
    console.warn('Ошибка записи лога:', error);
  }
}

/**
 * ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ОБЪЕДИНЕНИЯ
 */

function isEmptyValue_(value) {
  if (value === null || value === undefined || value === '') return true;
  
  const str = String(value).trim().toLowerCase();
  
  // Список значений, которые считаем пустыми
  const emptyValues = ['none', 'не закрыта', '—', '-', 'null', 'undefined'];
  
  return emptyValues.includes(str);
}

function parseRussianDate_(dateStr) {
  if (!dateStr) return new Date(0);
  
  try {
    // Если строка в формате ДД.ММ.ГГГГ или ДД.ММ.ГГГГ ЧЧ:ММ:СС
    const match = String(dateStr).match(/^(\d{2})\.(\d{2})\.(\d{4})/);
    if (match) {
      const [, day, month, year] = match;
      return new Date(year, month - 1, day);
    }
    
    return new Date(dateStr);
  } catch (error) {
    return new Date(0);
  }
}

function normalizePhoneClean_(phone) {
  if (!phone) return '';
  
  // Убираем все нецифровые символы  
  let clean = phone.toString().replace(/\D/g, '');
  
  // Если начинается с 8, заменяем на 7
  if (clean.startsWith('8') && clean.length === 11) {
    clean = '7' + clean.substring(1);
  }
  
  // Если начинается с 7 и длина 11 - возвращаем как есть (без +)
  if (clean.startsWith('7') && clean.length === 11) {
    return clean;
  }
  
  // Если 10 цифр, добавляем 7 в начало (без +)
  if (clean.length === 10) {
    return '7' + clean;
  }
  
  // Для других случаев возвращаем очищенный номер (без +)
  return clean.length >= 10 ? clean : phone;
}

/**
 * РАСЧЕТ СЧЕТА ФАКТ ПО НОВОЙ ЛОГИКЕ
 */
function calculateFactAmount_(fullDeal, regularDeal, guestData) {
  // Получаем статус сделки
  const fullD = fullDeal || {};
  const regD = regularDeal || {};
  const status = (fullD.status || regD.status || '').toString().replace(/^ВСЕ БАРЫ СЕТИ\s*\/\s*/, '');
  
  // Проверяем успешные статусы (гость был в баре)
  const successStatuses = ['Оплачено', 'успешно РЕАЛИЗОВАНО', 'успешно в РП'];
  const isSuccessful = successStatuses.includes(status);
  
  // Получаем ID для отладки
  const dealId = fullD.id || regD.id || 'неизвестно';
  
  if (!isSuccessful) {
    console.log(`💼 Сделка ${dealId}: статус "${status}" - не успешный, счет = 0`);
    return 0; // Если статус не успешный - возвращаем 0
  }
  
  // Если статус успешный - ищем сумму у гостя
  if (guestData) {
    const guestTotal = calculateGuestTotalBills_(guestData);
    if (guestTotal > 0) {
      console.log(`💰 Сделка ${dealId}: статус "${status}" - найден гость ${guestData.name || 'без имени'}, счет гостя = ${guestTotal}`);
      return guestTotal; // Если у гостя есть счета - берем их сумму
    }
  }
  
  // Если у гостя нет счетов (0 или пусто) - берем бюджет сделки
  const budget = fullD.budget || regD.budget || 0;
  const finalAmount = isNaN(parseFloat(budget)) ? 0 : parseFloat(budget);
  
  console.log(`📊 Сделка ${dealId}: статус "${status}" - ${guestData ? 'у гостя нет счетов' : 'гость не найден'}, берем бюджет = ${finalAmount}`);
  
  return finalAmount;
}

/**
 * РАСЧЕТ СУММЫ ВСЕХ СЧЕТОВ ГОСТЯ (с отладкой)
 */
function calculateGuestTotalBills_(guestData) {
  if (!guestData) return 0;
  
  const bills = [
    guestData.bill_1,
    guestData.bill_2,
    guestData.bill_3,
    guestData.bill_4,
    guestData.bill_5,
    guestData.bill_6,
    guestData.bill_7,
    guestData.bill_8,
    guestData.bill_9,
    guestData.bill_10
  ];
  
  let total = 0;
  let validBills = [];
  
  bills.forEach((bill, index) => {
    if (bill !== null && bill !== undefined && bill !== '') {
      const numValue = parseFloat(bill);
      if (!isNaN(numValue)) {
        total += numValue;
        validBills.push(`счёт_${index + 1}: ${numValue}`);
      }
    }
  });
  
  // Отладочная информация (можно убрать после проверки)
  if (validBills.length > 0) {
    console.log(`📊 Гость ${guestData.name || 'без имени'} (${guestData.phone}): ${validBills.join(', ')} = ${total}`);
  }
  
  return total;
}

/**
 * НОРМАЛИЗАЦИЯ ТЕЛЕФОНА БЕЗ ПЛЮСА (ДЛЯ ПОИСКА)
 */
function normalizePhoneForSearch_(phone) {
  if (!phone) return '';
  
  // Убираем все нецифровые символы  
  let clean = phone.toString().replace(/\D/g, '');
  
  // Если начинается с 8, заменяем на 7
  if (clean.startsWith('8') && clean.length === 11) {
    clean = '7' + clean.substring(1);
  }
  
  // Если 10 цифр, добавляем 7 в начало
  if (clean.length === 10) {
    clean = '7' + clean;
  }
  
  return clean.length >= 10 ? clean : '';
}

/**
 * СОЗДАНИЕ КАРТЫ ТЕЛЕФОНОВ
 */
function createPhoneMap_(dataArray) {
  const phoneMap = new Map();
  
  dataArray.forEach(item => {
    if (item.phone) {
      const normalizedPhone = normalizePhone_(item.phone);
      if (normalizedPhone) {
        phoneMap.set(normalizedPhone, item);
      }
      
      // Для гостей также добавляем поиск без плюса
      const phoneForSearch = normalizePhoneForSearch_(item.phone);
      if (phoneForSearch && phoneForSearch !== normalizedPhone) {
        phoneMap.set(phoneForSearch, item);
      }
    }
  });
  
  return phoneMap;
}

/**
 * НОРМАЛИЗАЦИЯ ТЕЛЕФОНА
 */
function normalizePhone_(phone) {
  if (!phone) return null;
  
  // Убираем все нецифровые символы
  let clean = phone.toString().replace(/\D/g, '');
  
  // Если начинается с 8, заменяем на 7
  if (clean.startsWith('8') && clean.length === 11) {
    clean = '7' + clean.substring(1);
  }
  
  // Если начинается с 7 и длина 11 - оставляем как есть
  if (clean.startsWith('7') && clean.length === 11) {
    return clean;
  }
  
  // Если 10 цифр, добавляем 7 в начало
  if (clean.length === 10) {
    return '7' + clean;
  }
  
  return clean.length >= 10 ? clean : null;
}

/**
 * СОЗДАНИЕ ОБОГАЩЕННОЙ СТРОКИ ДАННЫХ (ЧИСТАЯ ВЕРСИЯ)
 */
function createEnrichedDataRowClean_(deal, siteData, reserveData, guestData, callData) {
  // Функция для очистки названий от "ВСЕ БАРЫ СЕТИ /"
  const cleanName = (name) => {
    if (!name) return '';
    return String(name).replace(/^ВСЕ БАРЫ СЕТИ\s*\/\s*/, '');
  };
  
  return [
    // ОСНОВНЫЕ ДАННЫЕ ИЗ AMO (A-AP как в "Выгрузка Амо Полная")
    deal.id || '',                              // A
    cleanName(deal.name) || '',                 // B - название очищается от префикса
    deal.responsible || '',                     // C
    deal.contact_name || '',                    // D
    deal.status || '',                          // E
    deal.budget || 0,                           // F
    formatDateClean_(deal.created_at) || '',         // G
    deal.responsible2 || deal.responsible || '', // H
    deal.tags || '',                            // I
    formatDateClean_(deal.closed_at) || '',          // J
    deal.ym_client_id || siteData.ym_client_id || '', // K
    deal.ga_client_id || siteData.ga_client_id || '', // L
    deal.button_text || siteData.button_text || '',   // M
    deal.date || siteData.date || '',                 // N
    deal.time || siteData.time || '',                 // O
    deal.deal_source || '',                           // P
    deal.city_tag || siteData.user_city || '',        // Q
    deal.software || '',                              // R
    deal.bar_name || '',                              // S
    formatDateClean_(deal.booking_date) || '',             // T
    deal.guest_count || reserveData.guests || '',     // U
    deal.visit_time || '',                            // V
    deal.comment || reserveData.comment || '',        // W
    deal.source || siteData.utm_source || '',         // X
    deal.lead_type || '',                             // Y
    deal.refusal_reason || '',                        // Z
    deal.guest_status || reserveData.status || '',    // AA
    deal.referral_type || '',                         // AB
    siteData.utm_medium || deal.utm_medium || '',     // AC
    siteData.formname || deal.formname || '',         // AD
    siteData.referer || deal.referer || '',           // AE
    siteData.formid || deal.formid || '',             // AF
    deal.mango_line1 || '',                           // AG
    siteData.utm_source || deal.utm_source || '',     // AH
    siteData.utm_term || deal.utm_term || '',         // AI
    siteData.utm_campaign || deal.utm_campaign || '', // AJ
    siteData.utm_content || deal.utm_content || '',   // AK
    siteData.referrer || deal.utm_referrer || '',     // AL
    deal._ym_uid || '',                               // AM
    deal.phone || '',                                 // AN
    deal.mango_line2 || callData.mango_line || '',    // AO
    deal.notes || '',                                 // AP
    
    // ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ ОБОГАЩЕНИЯ
    callData.tel_source || '',                        // AQ — Источник из коллтрекинга
    callData.channel_name || '',                      // AR — Канал коллтрекинга
    reserveData.amount || 0,                          // AS — Сумма из резервов
    formatDateClean_(reserveData.datetime) || '',          // AT — Дата резерва
    guestData.visits || 0,                            // AU — Количество визитов
    guestData.total_amount || 0,                      // AV — Общая сумма гостя
    formatDateClean_(guestData.first_visit) || '',         // AW — Первый визит
    formatDateClean_(guestData.last_visit) || '',          // AX — Последний визит
    siteData.device_type || '',                       // AY — Тип устройства
    siteData.landing_page || '',                      // AZ — Посадочная страница
    
    // АВТОМАТИЧЕСКИЕ ПОЛЯ
    calculateDealAgeClean_(deal.created_at),               // BA — Возраст сделки (дн.)
    calculateDaysToBookingClean_(deal.created_at, deal.booking_date), // BB — Дней до брони
    siteData.visits_count || 0                        // BC — Визитов на сайт
  ];
}

/**
 * СОЗДАНИЕ ОБОГАЩЕННОЙ СТРОКИ ДАННЫХ
 */
function createEnrichedDataRow_(deal, siteData, reserveData, guestData, callData) {
  // Функция для очистки названий от "ВСЕ БАРЫ СЕТИ /"
  const cleanName = (name) => {
    if (!name) return '';
    return String(name).replace(/^ВСЕ БАРЫ СЕТИ\s*\/\s*/, '');
  };
  
  return [
    // ОСНОВНЫЕ ДАННЫЕ ИЗ AMO (A-AP как в "Выгрузка Амо Полная")
    deal.id || '',                              // A
    cleanName(deal.name) || '',                 // B - название очищается от префикса
    deal.responsible || '',                     // C
    deal.contact_name || '',                    // D
    deal.status || '',                          // E
    deal.budget || 0,                           // F
    formatDateClean_(deal.created_at) || '',         // G
    deal.responsible2 || deal.responsible || '', // H
    deal.tags || '',                            // I
    formatDateClean_(deal.closed_at) || '',          // J
    deal.ym_client_id || siteData.ym_client_id || '', // K
    deal.ga_client_id || siteData.ga_client_id || '', // L
    deal.button_text || siteData.button_text || '',   // M
    deal.date || siteData.date || '',                 // N
    deal.time || siteData.time || '',                 // O
    deal.deal_source || '',                           // P
    deal.city_tag || siteData.user_city || '',        // Q
    deal.software || '',                              // R
    deal.bar_name || '',                              // S
    formatDateClean_(deal.booking_date) || '',             // T
    deal.guest_count || reserveData.guests || '',     // U
    deal.visit_time || '',                            // V
    deal.comment || reserveData.comment || '',        // W
    deal.source || siteData.utm_source || '',         // X
    deal.lead_type || '',                             // Y
    deal.refusal_reason || '',                        // Z
    deal.guest_status || reserveData.status || '',    // AA
    deal.referral_type || '',                         // AB
    siteData.utm_medium || deal.utm_medium || '',     // AC
    siteData.formname || deal.formname || '',         // AD
    siteData.referer || deal.referer || '',           // AE
    siteData.formid || deal.formid || '',             // AF
    deal.mango_line1 || '',                           // AG
    siteData.utm_source || deal.utm_source || '',     // AH
    siteData.utm_term || deal.utm_term || '',         // AI
    siteData.utm_campaign || deal.utm_campaign || '', // AJ
    siteData.utm_content || deal.utm_content || '',   // AK
    siteData.referrer || deal.utm_referrer || '',     // AL
    deal._ym_uid || '',                               // AM
    deal.phone || '',                                 // AN
    deal.mango_line2 || callData.mango_line || '',    // AO
    deal.notes || '',                                 // AP
    
    // ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ ОБОГАЩЕНИЯ
    callData.tel_source || '',                        // AQ — Источник из коллтрекинга
    callData.channel_name || '',                      // AR — Канал коллтрекинга
    reserveData.amount || 0,                          // AS — Сумма из резервов
    formatDateClean_(reserveData.datetime) || '',          // AT — Дата резерва
    guestData.visits || 0,                            // AU — Количество визитов
    guestData.total_amount || 0,                      // AV — Общая сумма гостя
    formatDateClean_(guestData.first_visit) || '',         // AW — Первый визит
    formatDateClean_(guestData.last_visit) || '',          // AX — Последний визит
    siteData.device_type || '',                       // AY — Тип устройства
    siteData.landing_page || '',                      // AZ — Посадочная страница
    
    // АВТОМАТИЧЕСКИЕ ПОЛЯ
    calculateDealAgeClean_(deal.created_at),               // BA — Возраст сделки (дн.)
    calculateDaysToBookingClean_(deal.created_at, deal.booking_date), // BB — Дней до брони
    siteData.visits_count || 0                        // BC — Визитов на сайт
  ];
}

/**
 * СОЗДАНИЕ ЗАГОЛОВКОВ ТАБЛИЦЫ
 */
function createWorkingAmoHeaders_(sheet) {
  // Заголовки по реальной структуре ваших данных
  const headers = [
    // ОСНОВНЫЕ ДАННЫЕ AMO (A-AP)
    'ID', 'Название', 'Ответственный', 'Контакт.ФИО', 'Статус', 'Бюджет', 'Дата создания', 'Ответственный2', 'Теги', 'Дата закрытия',
    'YM_CLIENT_ID', 'GA_CLIENT_ID', 'BUTTON_TEXT', 'DATE', 'TIME', 'R.Источник сделки', 'R.Тег города', 'ПО',
    'Бар (deal)', 'Дата брони', 'Кол-во гостей', 'Время прихода', 'Комментарий МОБ', 'Источник', 'Тип лида', 'Причина отказа (ОБ)', 'R.Статусы гостей', 'Сарафан гости',
    'UTM_MEDIUM', 'FORMNAME', 'REFERER', 'FORMID', 'Номер линии MANGO OFFICE', 'UTM_SOURCE', 'UTM_TERM', 'UTM_CAMPAIGN', 'UTM_CONTENT', 'utm_referrer', '_ym_uid',
    'Контакт.Телефон', 'Контакт.Номер линии MANGO OFFICE', 'Примечания(через ;)',
    
    // ОБОГАЩЕНИЕ ДАННЫМИ
    'R.Источник ТЕЛ (коллтрекинг)', 'Канал (коллтрекинг)', 'Сумма ₽ (резерв)', 'Дата резерва', 'Визитов (гость)', 'Сумма ₽ (гость)', 'Первый визит (гость)', 'Последний визит (гость)',
    'Тип устройства (сайт)', 'Посадочная страница (сайт)',
    
    // АВТОПОЛЯ
    'Возраст сделки (дн.)', 'Дней до брони', 'Визитов на сайт'
  ];
  
  // Записываем заголовки
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Форматирование заголовков
  headerRange.setFontWeight('bold')
             .setBackground('#e6e6e6')
             .setHorizontalAlignment('center')
             .setWrap(true);
  
  // Цветовое кодирование секций
  applySectionColoring_(sheet, headers.length);
}

/**
 * ЦВЕТОВОЕ КОДИРОВАНИЕ СЕКЦИЙ
 */
function applySectionColoring_(sheet, totalCols) {
  // Основные данные AMO (A-AP, колонки 1-42) - светло-голубой
  sheet.getRange(1, 1, 1, 42).setBackground('#cfe2f3');
  
  // Обогащение данными (AQ-AZ, колонки 43-51) - светло-зеленый  
  sheet.getRange(1, 43, 1, 8).setBackground('#d9ead3');
  
  // Автополя (BA-BC, колонки 52-54) - светло-желтый
  sheet.getRange(1, 52, 1, 3).setBackground('#fff2cc');
}

/**
 * ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ - ЧИСТЫЕ ВЕРСИИ
 */

function formatDateClean_(date) {
  if (!date) return '';
  
  try {
    let d;
    
    // Обрабатываем разные типы входных данных
    if (date instanceof Date) {
      d = date;
    } else if (typeof date === 'string') {
      // Если строка уже в формате ДД.ММ.ГГГГ, возвращаем как есть
      const datePattern = /^\d{2}\.\d{2}\.\d{4}$/;
      if (datePattern.test(date)) {
        return date;
      }
      
      // Если строка в формате ДД.ММ.ГГГГ ЧЧ:ММ:СС, извлекаем дату
      const dateTimePattern = /^(\d{2})\.(\d{2})\.(\d{4})\s+\d{2}:\d{2}:\d{2}$/;
      const match = date.match(dateTimePattern);
      if (match) {
        return `${match[1]}.${match[2]}.${match[3]}`;
      }
      
      // Пробуем распарсить строку как дату
      d = new Date(date);
    } else if (typeof date === 'number') {
      // Если это timestamp
      d = new Date(date);
    } else {
      // Пробуем привести к строке и затем к дате
      const dateStr = String(date);
      
      // Проверяем паттерны даты/времени
      const dateTimePattern = /^(\d{2})\.(\d{2})\.(\d{4})\s+\d{2}:\d{2}:\d{2}$/;
      const match = dateStr.match(dateTimePattern);
      if (match) {
        return `${match[1]}.${match[2]}.${match[3]}`;
      }
      
      d = new Date(dateStr);
    }
    
    // Проверяем валидность даты
    if (!d || isNaN(d.getTime())) {
      return '';
    }
    
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    
    return `${day}.${month}.${year}`;
  } catch (error) {
    console.warn('Ошибка форматирования даты:', error, 'Входные данные:', date);
    return '';
  }
}

function calculateDealAgeClean_(createdDate) {
  if (!createdDate) return 0;
  
  try {
    let created;
    
    if (createdDate instanceof Date) {
      created = createdDate;
    } else if (typeof createdDate === 'string') {
      // Если строка в формате ДД.ММ.ГГГГ ЧЧ:ММ:СС
      const dateTimePattern = /^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/;
      const match = createdDate.match(dateTimePattern);
      if (match) {
        // Конвертируем в формат, понятный для Date
        const [, day, month, year, hour, minute, second] = match;
        created = new Date(year, month - 1, day, hour, minute, second);
      } else {
        created = new Date(createdDate);
      }
    } else if (typeof createdDate === 'number') {
      created = new Date(createdDate);
    } else {
      const dateStr = String(createdDate);
      const dateTimePattern = /^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/;
      const match = dateStr.match(dateTimePattern);
      if (match) {
        const [, day, month, year, hour, minute, second] = match;
        created = new Date(year, month - 1, day, hour, minute, second);
      } else {
        created = new Date(dateStr);
      }
    }
    
    if (!created || isNaN(created.getTime())) {
      return 0;
    }
    
    const now = new Date();
    const diffTime = now - created;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    return diffDays > 0 ? diffDays : 0;
  } catch (error) {
    console.warn('Ошибка расчета возраста сделки:', error);
    return 0;
  }
}

function calculateDaysToBookingClean_(createdDate, bookingDate) {
  if (!createdDate || !bookingDate) return 0;
  
  try {
    let created, booking;
    
    // Функция для обработки даты в формате ДД.ММ.ГГГГ ЧЧ:ММ:СС
    const parseDate = (dateInput) => {
      if (dateInput instanceof Date) {
        return dateInput;
      } else if (typeof dateInput === 'string') {
        const dateTimePattern = /^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/;
        const match = dateInput.match(dateTimePattern);
        if (match) {
          const [, day, month, year, hour, minute, second] = match;
          return new Date(year, month - 1, day, hour, minute, second);
        }
        return new Date(dateInput);
      } else if (typeof dateInput === 'number') {
        return new Date(dateInput);
      } else {
        const dateStr = String(dateInput);
        const dateTimePattern = /^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})$/;
        const match = dateStr.match(dateTimePattern);
        if (match) {
          const [, day, month, year, hour, minute, second] = match;
          return new Date(year, month - 1, day, hour, minute, second);
        }
        return new Date(dateStr);
      }
    };
    
    created = parseDate(createdDate);
    booking = parseDate(bookingDate);
    
    if (!created || !booking || isNaN(created.getTime()) || isNaN(booking.getTime())) {
      return 0;
    }
    
    const diffTime = booking - created;
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    
    return diffDays > 0 ? diffDays : 0;
  } catch (error) {
    console.warn('Ошибка расчета дней до брони:', error);
    return 0;
  }
}

/**
 * ДИАГНОСТИКА СЧЕТОВ ГОСТЕЙ
 */
function diagnosticGuestBills() {
  console.log('🔍 ДИАГНОСТИКА СЧЕТОВ ГОСТЕЙ');
  
  try {
    const guestsData = loadGuestsData_();
    console.log(`📊 Загружено гостей: ${guestsData.length}`);
    
    let totalWithBills = 0;
    let totalBillsSum = 0;
    
    guestsData.forEach((guest, index) => {
      if (index < 5) { // Показываем первые 5 для примера
        console.log(`\n👤 Гость ${index + 1}: ${guest.name || 'без имени'}`);
        console.log(`📞 Телефон: ${guest.phone}`);
        console.log(`💰 Счета: ${guest.bill_1}, ${guest.bill_2}, ${guest.bill_3}, ${guest.bill_4}, ${guest.bill_5}`);
        console.log(`💰 Счета 6-10: ${guest.bill_6}, ${guest.bill_7}, ${guest.bill_8}, ${guest.bill_9}, ${guest.bill_10}`);
        
        const total = calculateGuestTotalBills_(guest);
        console.log(`💎 Итого счетов: ${total}`);
      }
      
      const billTotal = calculateGuestTotalBills_(guest);
      if (billTotal > 0) {
        totalWithBills++;
        totalBillsSum += billTotal;
      }
    });
    
    console.log(`\n📈 Итого: ${totalWithBills} гостей со счетами, общая сумма: ${totalBillsSum}`);
    
  } catch (error) {
    console.error('❌ Ошибка диагностики:', error);
  }
}

/**
 * ДИАГНОСТИКА СОПОСТАВЛЕНИЯ ТЕЛЕФОНОВ
 */
function diagnosticPhoneMatching() {
  console.log('🔍 ДИАГНОСТИКА СОПОСТАВЛЕНИЯ ТЕЛЕФОНОВ');
  
  try {
    const guestsData = loadGuestsData_();
    const phoneToGuests = createPhoneMap_(guestsData);
    
    console.log(`📊 Карта телефонов создана: ${phoneToGuests.size} записей`);
    
    // Тестируем разные форматы телефонов
    const testPhones = [
      '+79996106215',
      '79996106215', 
      '89996106215',
      '9996106215',
      '+7 (999) 610-62-15'
    ];
    
    testPhones.forEach(phone => {
      const normalized = normalizePhoneForSearch_(phone);
      const found = phoneToGuests.get(normalized);
      console.log(`📞 ${phone} → ${normalized} → ${found ? found.name || 'найден' : 'НЕ НАЙДЕН'}`);
    });
    
  } catch (error) {
    console.error('❌ Ошибка диагностики сопоставления:', error);
  }
}

/**
 * УТИЛИТЫ ДЛЯ ИСПРАВЛЕНИЯ ОШИБОК ВАЛИДАЦИИ
 */

function clearDataValidation_(sheet) {
  try {
    // Очищаем правила проверки данных со всего листа
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    
    if (maxRows > 0 && maxCols > 0) {
      const range = sheet.getRange(1, 1, maxRows, maxCols);
      range.clearDataValidations();
      console.log('🧹 Правила проверки данных очищены');
    }
  } catch (error) {
    console.warn('⚠️ Не удалось очистить правила валидации:', error);
  }
}

function normalizeStatusForValidation_(status) {
  if (!status || typeof status !== 'string') return '';
  
  // Маппинг статусов для соответствия правилам валидации
  const statusMapping = {
    // Входящие варианты → Допустимые значения
    'Автораспределение': 'Автораспределение',
    'Взяли в работу': 'Взяли в работу', 
    'вопрос к бару': 'вопрос к бару',
    'Закрыто и не реализовано': 'Закрыто и не реализовано',
    'Контроль оплаты': 'Контроль оплаты',
    'НДЗ': 'НДЗ',
    'Оплачено': 'Оплачено',
    'Перенос с открытой датой': 'Перенос с открытой датой',
    'Сделка.Статус': 'Сделка.Статус',
    'успешно в РП': 'успешно в РП',
    'успешно РЕАЛИЗОВАНО': 'успешно РЕАЛИЗОВАНО',
    
    // Дополнительные варианты
    'успешно реализовано': 'успешно РЕАЛИЗОВАНО',
    'Успешно реализовано': 'успешно РЕАЛИЗОВАНО',
    'УСПЕШНО РЕАЛИЗОВАНО': 'успешно РЕАЛИЗОВАНО',
    'Успешно в РП': 'успешно в РП',
    'УСПЕШНО В РП': 'успешно в РП',
    'оплачено': 'Оплачено',
    'ОПЛАЧЕНО': 'Оплачено',
    'автораспределение': 'Автораспределение',
    'АВТОРАСПРЕДЕЛЕНИЕ': 'Автораспределение',
    'взяли в работу': 'Взяли в работу',
    'ВЗЯЛИ В РАБОТУ': 'Взяли в работу',
    'контроль оплаты': 'Контроль оплаты',
    'КОНТРОЛЬ ОПЛАТЫ': 'Контроль оплаты',
    'закрыто и не реализовано': 'Закрыто и не реализовано',
    'ЗАКРЫТО И НЕ РЕАЛИЗОВАНО': 'Закрыто и не реализовано'
  };
  
  // Точное совпадение
  if (statusMapping[status]) {
    return statusMapping[status];
  }
  
  // Поиск по частичному совпадению (без учета регистра)
  const lowerStatus = status.toLowerCase();
  for (const [key, value] of Object.entries(statusMapping)) {
    if (key.toLowerCase() === lowerStatus) {
      return value;
    }
  }
  
  // Если статус не найден, возвращаем как есть, но предупреждаем
  console.warn(`⚠️ Неизвестный статус: "${status}"`);
  return status;
}

function repairWorkingAmoSheet() {
  console.log('🔧 ВОССТАНОВЛЕНИЕ ЛИСТА РАБОЧИЙ АМО');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('РАБОЧИЙ АМО');
    
    if (!sheet) {
      console.log('❌ Лист РАБОЧИЙ АМО не найден');
      return;
    }
    
    console.log('🧹 Очистка правил валидации...');
    clearDataValidation_(sheet);
    
    console.log('🔄 Нормализация статусов...');
    const data = sheet.getDataRange().getValues();
    
    if (data.length > 1) {
      // Обрабатываем строки с данными (пропускаем заголовки)
      for (let i = 1; i < data.length; i++) {
        if (data[i][2]) { // Колонка C (статус)
          data[i][2] = normalizeStatusForValidation_(data[i][2]);
        }
      }
      
      // Записываем обратно
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      console.log(`✅ Обработано ${data.length - 1} записей`);
    }
    
    console.log('✅ Восстановление завершено');
    
  } catch (error) {
    console.error('❌ Ошибка восстановления:', error);
    throw error;
  }
}
}
