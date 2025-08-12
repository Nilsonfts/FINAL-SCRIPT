/**
 * МОДУЛЬ ИМПОРТА И СИНХРОНИЗАЦИИ ДАННЫХ
 * Загрузка данных из внешних источников (AmoCRM, формы сайта, CTI)
 * @fileoverview Центральный модуль для импорта данных из всех источников
 */

/**
 * Основная функция синхронизации всех данных
 * Вызывается по расписанию каждые 15 минут
 */
function syncAllData() {
  try {
    logInfo_('DATA_SYNC', 'Начало синхронизации всех данных');
    
    const startTime = new Date();
    const results = {
      amocrm: { success: false, records: 0, error: null },
      forms: { success: false, records: 0, error: null },
      calls: { success: false, records: 0, error: null },
      metrica: { success: false, records: 0, error: null }
    };
    
    // 1. Синхронизация данных из AmoCRM
    try {
      logInfo_('DATA_SYNC', 'Синхронизация AmoCRM...');
      const amocrmResult = syncAmoCrmData_();
      results.amocrm = amocrmResult;
      logInfo_('DATA_SYNC', `AmoCRM: обработано ${amocrmResult.records} записей`);
    } catch (error) {
      results.amocrm.error = error.message;
      logError_('DATA_SYNC', 'Ошибка синхронизации AmoCRM', error);
    }
    
    // 2. Синхронизация заявок с сайта
    try {
      logInfo_('DATA_SYNC', 'Синхронизация форм сайта...');
      const formsResult = syncWebFormsData_();
      results.forms = formsResult;
      logInfo_('DATA_SYNC', `Формы сайта: обработано ${formsResult.records} записей`);
    } catch (error) {
      results.forms.error = error.message;
      logError_('DATA_SYNC', 'Ошибка синхронизации форм сайта', error);
    }
    
    // 3. Синхронизация звонков из CTI
    try {
      logInfo_('DATA_SYNC', 'Синхронизация CTI...');
      const callsResult = syncCallTrackingData_();
      results.calls = callsResult;
      logInfo_('DATA_SYNC', `CTI: обработано ${callsResult.records} записей`);
    } catch (error) {
      results.calls.error = error.message;
      logError_('DATA_SYNC', 'Ошибка синхронизации CTI', error);
    }
    
    // 4. Синхронизация данных из Яндекс.Метрики
    try {
      logInfo_('DATA_SYNC', 'Синхронизация Яндекс.Метрики...');
      const metricaResult = syncYandexMetricaData_();
      results.metrica = metricaResult;
      logInfo_('DATA_SYNC', `Метрика: обработано ${metricaResult.records} записей`);
    } catch (error) {
      results.metrica.error = error.message;
      logError_('DATA_SYNC', 'Ошибка синхронизации Метрики', error);
    }
    
    // 5. Обновление сводных данных
    try {
      logInfo_('DATA_SYNC', 'Обновление сводных данных...');
      updateConsolidatedData_();
    } catch (error) {
      logError_('DATA_SYNC', 'Ошибка обновления сводных данных', error);
    }
    
    const duration = (new Date() - startTime) / 1000;
    const totalRecords = results.amocrm.records + results.forms.records + results.calls.records + results.metrica.records;
    
    logInfo_('DATA_SYNC', `Синхронизация завершена за ${duration}с. Всего обработано: ${totalRecords} записей`);
    
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
 * Синхронизация данных из AmoCRM
 * @returns {Object} Результат синхронизации
 * @private
 */
function syncAmoCrmData_() {
  const apiUrl = PropertiesService.getScriptProperties().getProperty('AMOCRM_API_URL');
  const apiToken = PropertiesService.getScriptProperties().getProperty('AMOCRM_API_TOKEN');
  
  if (!apiUrl || !apiToken) {
    throw new Error('Не настроены параметры подключения к AmoCRM');
  }
  
  let totalRecords = 0;
  const batchSize = 250; // AmoCRM лимит
  let page = 1;
  
  // Получаем последнюю дату синхронизации
  const lastSync = getLastSyncTime_('amocrm');
  const syncTime = getCurrentDateMoscow_();
  
  try {
    do {
      const response = fetchAmoCrmDeals_(apiUrl, apiToken, page, batchSize, lastSync);
      
      if (!response || !response._embedded || !response._embedded.leads) {
        break;
      }
      
      const leads = response._embedded.leads;
      
      // Обрабатываем полученные сделки
      const processedRecords = processAmoCrmLeads_(leads);
      totalRecords += processedRecords;
      
      // Если получили меньше записей чем ожидали - это последняя страница
      if (leads.length < batchSize) {
        break;
      }
      
      page++;
      
      // Защита от бесконечного цикла
      if (page > 100) {
        logWarning_('DATA_SYNC', 'Достигнут лимит страниц AmoCRM (100)');
        break;
      }
      
      // Небольшая задержка между запросами
      Utilities.sleep(200);
      
    } while (true);
    
    // Сохраняем время успешной синхронизации
    setLastSyncTime_('amocrm', syncTime);
    
    return {
      success: true,
      records: totalRecords,
      error: null
    };
    
  } catch (error) {
    logError_('AMOCRM_SYNC', 'Ошибка синхронизации AmoCRM', error);
    return {
      success: false,
      records: totalRecords,
      error: error.message
    };
  }
}

/**
 * Выполняет запрос к API AmoCRM для получения сделок
 * @param {string} apiUrl - URL API
 * @param {string} apiToken - Токен доступа
 * @param {number} page - Номер страницы
 * @param {number} limit - Количество записей
 * @param {Date} lastSync - Дата последней синхронизации
 * @returns {Object} Ответ API
 * @private
 */
function fetchAmoCrmDeals_(apiUrl, apiToken, page, limit, lastSync) {
  const url = `${apiUrl}/api/v4/leads`;
  const params = {
    limit: limit,
    page: page,
    'with': 'contacts,companies,custom_fields_values',
    'order[updated_at]': 'asc'
  };
  
  // Добавляем фильтр по времени обновления
  if (lastSync) {
    params['filter[updated_at][from]'] = Math.floor(lastSync.getTime() / 1000);
  }
  
  const queryString = Object.keys(params)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
    .join('&');
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${apiToken}`,
      'Content-Type': 'application/json'
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(`${url}?${queryString}`, options);
    
    if (response.getResponseCode() === 204) {
      return null; // Нет данных
    }
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`AmoCRM API error: ${response.getResponseCode()} - ${response.getContentText()}`);
    }
    
    return JSON.parse(response.getContentText());
    
  } catch (error) {
    logError_('AMOCRM_API', `Ошибка запроса к AmoCRM API: ${url}`, error);
    throw error;
  }
}

/**
 * Обрабатывает сделки из AmoCRM и сохраняет в таблицу
 * @param {Array} leads - Массив сделок
 * @returns {number} Количество обработанных записей
 * @private
 */
function processAmoCrmLeads_(leads) {
  if (!leads || leads.length === 0) return 0;
  
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.RAW_DATA);
  const existingData = getSheetData_(sheet);
  const headers = existingData[0] || [];
  
  // Убеждаемся что есть все необходимые заголовки
  ensureHeaders_(sheet, CONFIG.HEADERS.LEADS_RAW);
  
  const processedRows = [];
  
  leads.forEach(lead => {
    try {
      const processedLead = processAmoCrmLead_(lead);
      if (processedLead) {
        processedRows.push(processedLead);
      }
    } catch (error) {
      logError_('AMOCRM_PROCESS', `Ошибка обработки сделки ID ${lead.id}`, error);
    }
  });
  
  // Записываем данные пакетом
  if (processedRows.length > 0) {
    appendRowsToSheet_(sheet, processedRows);
  }
  
  return processedRows.length;
}

/**
 * Обрабатывает одну сделку из AmoCRM
 * @param {Object} lead - Сделка AmoCRM
 * @returns {Array} Строка для записи в таблицу
 * @private
 */
function processAmoCrmLead_(lead) {
  // Извлекаем основные поля
  const leadId = lead.id;
  const leadName = lead.name || '';
  const budget = lead.price || 0;
  const statusId = lead.status_id;
  const pipelineId = lead.pipeline_id;
  const responsibleUserId = lead.responsible_user_id;
  const createdAt = new Date(lead.created_at * 1000);
  const updatedAt = new Date(lead.updated_at * 1000);
  const closedAt = lead.closed_at ? new Date(lead.closed_at * 1000) : null;
  
  // Определяем статус
  const status = getLeadStatusByAmoCrmId_(statusId);
  const isWon = status === 'success';
  
  // Извлекаем UTM метки из кастомных полей
  const utmData = extractUtmFromCustomFields_(lead.custom_fields_values || []);
  
  // Извлекаем контактные данные
  const contactData = extractContactData_(lead);
  
  // Определяем канал привлечения
  const channel = determineChannel_(utmData, {
    referrer: extractCustomField_(lead.custom_fields_values, 'Реферер'),
    tel_source: extractCustomField_(lead.custom_fields_values, 'Источник звонка')
  });
  
  // Формируем строку данных
  return [
    leadId,                                    // ID сделки
    leadName,                                  // Название
    formatDate_(createdAt, 'full'),           // Дата создания
    formatDate_(updatedAt, 'full'),           // Дата обновления
    closedAt ? formatDate_(closedAt, 'full') : '', // Дата закрытия
    status,                                    // Статус
    budget,                                    // Бюджет
    channel,                                   // Канал
    utmData.utm_source || '',                 // UTM Source
    utmData.utm_medium || '',                 // UTM Medium
    utmData.utm_campaign || '',               // UTM Campaign
    utmData.utm_content || '',                // UTM Content
    utmData.utm_term || '',                   // UTM Term
    contactData.name || '',                   // Имя клиента
    contactData.phone || '',                  // Телефон
    contactData.email || '',                  // Email
    extractCustomField_(lead.custom_fields_values, 'Город') || '', // Город
    responsibleUserId,                        // Ответственный
    pipelineId,                               // Воронка
    statusId,                                 // ID статуса
    'AmoCRM',                                 // Источник данных
    getCurrentDateMoscow_().toISOString()     // Время синхронизации
  ];
}

/**
 * Синхронизация заявок с сайта (через Google Формы или веб-хуки)
 * @returns {Object} Результат синхронизации
 * @private
 */
function syncWebFormsData_() {
  // Здесь может быть логика получения данных из Google Форм
  // или из внешних источников через веб-хуки
  
  try {
    let totalRecords = 0;
    
    // Пример: синхронизация из Google Форм
    const formsData = getGoogleFormsData_();
    if (formsData && formsData.length > 0) {
      totalRecords += processWebFormsData_(formsData);
    }
    
    // Пример: синхронизация из файла выгрузки
    const fileData = getWebFormsFromFile_();
    if (fileData && fileData.length > 0) {
      totalRecords += processWebFormsData_(fileData);
    }
    
    return {
      success: true,
      records: totalRecords,
      error: null
    };
    
  } catch (error) {
    return {
      success: false,
      records: 0,
      error: error.message
    };
  }
}

/**
 * Синхронизация данных из колл-трекинга
 * @returns {Object} Результат синхронизации
 * @private
 */
function syncCallTrackingData_() {
  try {
    // Получаем данные из API колл-трекинга или файла
    const callData = getCallTrackingData_();
    
    if (!callData || callData.length === 0) {
      return {
        success: true,
        records: 0,
        error: null
      };
    }
    
    const processedRecords = processCallTrackingData_(callData);
    
    return {
      success: true,
      records: processedRecords,
      error: null
    };
    
  } catch (error) {
    return {
      success: false,
      records: 0,
      error: error.message
    };
  }
}

/**
 * Синхронизация данных из Яндекс.Метрики
 * @returns {Object} Результат синхронизации
 * @private
 */
function syncYandexMetricaData_() {
  const metricaToken = PropertiesService.getScriptProperties().getProperty('YANDEX_METRICA_TOKEN');
  const counterId = PropertiesService.getScriptProperties().getProperty('YANDEX_METRICA_COUNTER_ID');
  
  if (!metricaToken || !counterId) {
    logWarning_('METRICA_SYNC', 'Не настроены параметры Яндекс.Метрики');
    return {
      success: true,
      records: 0,
      error: null
    };
  }
  
  try {
    // Получаем данные за последние несколько дней
    const dateRange = createDateRange_('week');
    const metricaData = fetchYandexMetricaData_(metricaToken, counterId, dateRange);
    
    if (!metricaData || metricaData.length === 0) {
      return {
        success: true,
        records: 0,
        error: null
      };
    }
    
    const processedRecords = processYandexMetricaData_(metricaData);
    
    return {
      success: true,
      records: processedRecords,
      error: null
    };
    
  } catch (error) {
    return {
      success: false,
      records: 0,
      error: error.message
    };
  }
}

/**
 * Получает время последней синхронизации
 * @param {string} source - Источник данных
 * @returns {Date|null} Время последней синхронизации
 * @private
 */
function getLastSyncTime_(source) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const lastSyncStr = properties.getProperty(`LAST_SYNC_${source.toUpperCase()}`);
    
    if (lastSyncStr) {
      return new Date(lastSyncStr);
    }
    
    // Если нет сохранённого времени - берём данные за последние 7 дней
    return addDays_(getCurrentDateMoscow_(), -7);
    
  } catch (error) {
    logWarning_('SYNC_TIME', `Ошибка получения времени последней синхронизации для ${source}`, error);
    return addDays_(getCurrentDateMoscow_(), -7);
  }
}

/**
 * Сохраняет время последней синхронизации
 * @param {string} source - Источник данных
 * @param {Date} syncTime - Время синхронизации
 * @private
 */
function setLastSyncTime_(source, syncTime) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty(`LAST_SYNC_${source.toUpperCase()}`, syncTime.toISOString());
  } catch (error) {
    logWarning_('SYNC_TIME', `Ошибка сохранения времени синхронизации для ${source}`, error);
  }
}

/**
 * Извлекает UTM метки из кастомных полей AmoCRM
 * @param {Array} customFields - Кастомные поля
 * @returns {Object} UTM данные
 * @private
 */
function extractUtmFromCustomFields_(customFields) {
  const utm = {};
  
  if (!customFields) return utm;
  
  customFields.forEach(field => {
    const fieldName = field.field_name ? field.field_name.toLowerCase() : '';
    const value = field.values && field.values[0] ? field.values[0].value : '';
    
    if (fieldName.includes('utm_source') || fieldName.includes('источник')) {
      utm.utm_source = value;
    } else if (fieldName.includes('utm_medium') || fieldName.includes('канал')) {
      utm.utm_medium = value;
    } else if (fieldName.includes('utm_campaign') || fieldName.includes('кампания')) {
      utm.utm_campaign = value;
    } else if (fieldName.includes('utm_content')) {
      utm.utm_content = value;
    } else if (fieldName.includes('utm_term')) {
      utm.utm_term = value;
    }
  });
  
  return utm;
}

/**
 * Извлекает контактные данные из сделки AmoCRM
 * @param {Object} lead - Сделка
 * @returns {Object} Контактные данные
 * @private
 */
function extractContactData_(lead) {
  const contact = {
    name: '',
    phone: '',
    email: ''
  };
  
  // Извлекаем из связанных контактов
  if (lead._embedded && lead._embedded.contacts && lead._embedded.contacts.length > 0) {
    const firstContact = lead._embedded.contacts[0];
    contact.name = firstContact.name || '';
    
    // Извлекаем телефон и email из кастомных полей контакта
    if (firstContact.custom_fields_values) {
      firstContact.custom_fields_values.forEach(field => {
        if (field.field_code === 'PHONE' && field.values && field.values[0]) {
          contact.phone = field.values[0].value;
        } else if (field.field_code === 'EMAIL' && field.values && field.values[0]) {
          contact.email = field.values[0].value;
        }
      });
    }
  }
  
  return contact;
}

/**
 * Извлекает значение кастомного поля
 * @param {Array} customFields - Кастомные поля
 * @param {string} fieldName - Название поля
 * @returns {string} Значение поля
 * @private
 */
function extractCustomField_(customFields, fieldName) {
  if (!customFields || !fieldName) return '';
  
  const field = customFields.find(f => 
    f.field_name && f.field_name.toLowerCase().includes(fieldName.toLowerCase())
  );
  
  if (field && field.values && field.values[0]) {
    return field.values[0].value || '';
  }
  
  return '';
}

/**
 * Обновляет консолидированные данные после синхронизации
 * @private
 */
function updateConsolidatedData_() {
  try {
    logInfo_('DATA_CONSOLIDATION', 'Начало обновления консолидированных данных');
    
    // Обновляем кэш заголовков
    clearHeadersCache_();
    
    // Запускаем основные аналитические функции
    // Эти функции будут созданы в следующих модулях
    
    logInfo_('DATA_CONSOLIDATION', 'Консолидированные данные обновлены');
    
  } catch (error) {
    logError_('DATA_CONSOLIDATION', 'Ошибка обновления консолидированных данных', error);
    throw error;
  }
}

/**
 * Отправляет уведомление об ошибках синхронизации
 * @param {Object} results - Результаты синхронизации
 * @private
 */
function sendErrorNotification_(results) {
  try {
    const emailRecipients = CONFIG.EMAIL_NOTIFICATIONS.RECIPIENTS;
    if (!emailRecipients || emailRecipients.length === 0) return;
    
    const errors = Object.entries(results)
      .filter(([source, result]) => result.error)
      .map(([source, result]) => `${source}: ${result.error}`)
      .join('\n');
    
    if (errors) {
      const subject = 'Ошибки синхронизации данных';
      const body = `Обнаружены ошибки в процессе синхронизации данных:\n\n${errors}\n\nВремя: ${formatDate_(getCurrentDateMoscow_(), 'full')}`;
      
      MailApp.sendEmail({
        to: emailRecipients.join(','),
        subject: subject,
        body: body
      });
    }
    
  } catch (error) {
    logError_('EMAIL_NOTIFICATION', 'Ошибка отправки уведомления', error);
  }
}
