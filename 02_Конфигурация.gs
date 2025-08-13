/**
 * КОНФИГУРАЦИЯ СИСТЕМЫ AMO ANALYTICS
 * Основные настройки, константы и маппинги
 */

// 🔑 API ТОКЕНЫ И ДОСТУПЫ (ЗАПОЛНИТЕ СВОИМИ ДАННЫМИ!)
const API_TOKENS = {
  // Яндекс.Метрика - https://oauth.yandex.ru/authorize?response_type=token&client_id=c2a8f5c2abeb4b9288df8b3fb6b3c9a2
  YANDEX_METRIKA_TOKEN: 'ВСТАВЬТЕ_СЮДА_ВАШ_ТОКЕН_МЕТРИКИ',
  YANDEX_METRIKA_COUNTER_ID: 'ВСТАВЬТЕ_СЮДА_ID_СЧЕТЧИКА', // например: '12345678'
  
  // Яндекс.Директ - https://oauth.yandex.ru/authorize?response_type=token&client_id=ce780928adc04f78b79c0ac932eb5c8e  
  YANDEX_DIRECT_TOKEN: 'ВСТАВЬТЕ_СЮДА_ВАШ_ТОКЕН_ДИРЕКТА',
  
  // AMO CRM
  AMO_SUBDOMAIN: 'ВСТАВЬТЕ_ПОДДОМЕН',  // например: 'mycompany' (без .amocrm.ru)
  AMO_ACCESS_TOKEN: 'ВСТАВЬТЕ_ТОКЕН_AMO',
  AMO_CLIENT_ID: 'ВСТАВЬТЕ_CLIENT_ID_AMO',
  AMO_CLIENT_SECRET: 'ВСТАВЬТЕ_CLIENT_SECRET_AMO',
  AMO_REDIRECT_URI: 'ВСТАВЬТЕ_REDIRECT_URI_AMO',
  
  // Call-tracking системы
  MANGO_API_KEY: 'ВСТАВЬТЕ_КЛЮЧ_МАНГО',
  MANGO_API_SALT: 'ВСТАВЬТЕ_СОЛЬ_МАНГО',
  COMAGIC_TOKEN: 'ВСТАВЬТЕ_ТОКЕН_КОМАДЖИК',
  
  // Дополнительные интеграции
  GOOGLE_ANALYTICS_VIEW_ID: 'ВСТАВЬТЕ_VIEW_ID_GA',
  FACEBOOK_ACCESS_TOKEN: 'ВСТАВЬТЕ_ТОКЕН_FACEBOOK',
  VK_ACCESS_TOKEN: 'ВСТАВЬТЕ_ТОКЕН_VK'
};

// Основная конфигурация системы
const CONFIG = {
  // Названия листов в Google Sheets
  SHEETS: {
    WORKING_AMO: 'РАБОЧИЙ_АМО',
    LEADS_CHANNELS: '📊 ЛИДЫ ПО КАНАЛАМ',
    MARKETING_CHANNELS: '📈 МАРКЕТИНГ И КАНАЛЫ',
    FIRST_TOUCH: '🎯 ПЕРВЫЕ КАСАНИЯ',
    CALL_TRACKING: '📞 КОЛЛ-ТРЕКИНГ',
    REFUSAL_REASONS: '🤖 ПРИЧИНЫ ОТКАЗОВ AI',
    YANDEX_DIRECT: '🎯 ЯНДЕКС ДИРЕКТ',
    SITE_CROSS: '🌐 СКВОЗНАЯ АНАЛИТИКА САЙТ',
    AMO_SUMMARY: '📋 СВОДНАЯ АНАЛИТИКА АМО',
    YANDEX_METRIKA: '📊 ЯНДЕКС МЕТРИКА',
    MONTHLY_DASHBOARD: '📅 МЕСЯЧНЫЙ ДАШБОРД',
    CHANNEL_ANALYSIS: '🔍 АНАЛИЗ КАНАЛОВ',
    VISUALIZATION: '📊 ВИЗУАЛИЗАЦИЯ',
    DIAGNOSTICS: '🔧 ДИАГНОСТИКА',
    EXPORT_IMPORT: '💾 ЭКСПОРТ/ИМПОРТ',
    SYSTEM_UTILS: '⚙️ СИСТЕМНЫЕ УТИЛИТЫ',
    DEMO_TESTS: '🎯 ДЕМО И ТЕСТЫ'
  },

  // Маппинг колонок рабочего листа АМО - ПРАВИЛЬНАЯ ВЕРСИЯ (41 столбец)
  WORKING_AMO_COLUMNS_CORRECT: {
    ID: 0,                          // A - Сделка.ID
    NAME: 1,                        // B - Сделка.Название  
    STATUS: 2,                      // C - Сделка.Статус
    REFUSAL_REASON: 3,              // D - Сделка.Причина отказа (ОБ)
    LEAD_TYPE: 4,                   // E - Сделка.Тип лида
    GUEST_STATUS: 5,                // F - Сделка.R.Статусы гостей
    RESPONSIBLE: 6,                 // G - Сделка.Ответственный
    TAGS: 7,                        // H - Сделка.Теги
    BUDGET: 8,                      // I - Сделка.Бюджет
    CREATED_AT: 9,                  // J - Сделка.Дата создания
    CLOSED_AT: 10,                  // K - Сделка.Дата закрытия
    CONTACT_MANGO: 11,              // L - Контакт.Номер линии MANGO OFFICE
    DEAL_MANGO: 12,                 // M - Сделка.Номер линии MANGO OFFICE
    CONTACT_NAME: 13,               // N - Контакт.ФИО
    PHONE: 14,                      // O - Контакт.Телефон
    FACT_AMOUNT: 15,                // P - Счет факт
    DATE: 16,                       // Q - Сделка.DATE
    TIME: 17,                       // R - Сделка.TIME
    CITY_TAG: 18,                   // S - Сделка.R.Тег города
    BOOKING_DATE: 19,               // T - Сделка.Дата брони
    SOFTWARE: 20,                   // U - Сделка.ПО
    REFERRAL_TYPE: 21,              // V - Сделка.Сарафан гости
    BAR_NAME: 22,                   // W - Сделка.Бар (deal)
    DEAL_SOURCE: 23,                // X - Сделка.R.Источник сделки
    BUTTON_TEXT: 24,                // Y - Сделка.BUTTON_TEXT
    YM_CLIENT_ID: 25,               // Z - Сделка.YM_CLIENT_ID
    GA_CLIENT_ID: 26,               // AA - Сделка.GA_CLIENT_ID
    UTM_SOURCE: 27,                 // AB - Сделка.UTM_SOURCE
    UTM_MEDIUM: 28,                 // AC - Сделка.UTM_MEDIUM
    UTM_TERM: 29,                   // AD - Сделка.UTM_TERM
    UTM_CAMPAIGN: 30,               // AE - Сделка.UTM_CAMPAIGN
    UTM_CONTENT: 31,                // AF - Сделka.UTM_CONTENT
    UTM_REFERRER: 32,               // AG - Сделка.utm_referrer
    VISIT_TIME: 33,                 // AH - Сделка.Время прихода
    COMMENT: 34,                    // AI - Сделка.Комментарий МОБ
    SOURCE: 35,                     // AJ - Сделка.Источник
    FORMNAME: 36,                   // AK - Сделка.FORMNAME
    REFERER: 37,                    // AL - Сделка.REFERER
    FORMID: 38,                     // AM - Сделka.FORMID
    YM_UID: 39,                     // AN - Сделка._ym_uid
    NOTES: 40                       // AO - Сделка.Примечания(через ;)
  },

  // Маппинг колонок рабочего листа АМО - НЕПРАВИЛЬНАЯ ВЕРСИЯ (54+ столбцов)
  WORKING_AMO_COLUMNS_INCORRECT: {
    ID: 0,                          // A - ID
    NAME: 1,                        // B - Название
    RESPONSIBLE: 2,                 // C - Ответственный
    CONTACT_NAME: 3,                // D - Контакт.ФИО
    STATUS: 4,                      // E - Статус
    BUDGET: 5,                      // F - Бюджет
    CREATED_AT: 6,                  // G - Дата создания
    RESPONSIBLE2: 7,                // H - Ответственный2
    TAGS: 8,                        // I - Теги
    CLOSED_AT: 9,                   // J - Дата закрытия
    YM_CLIENT_ID: 10,               // K - YM_CLIENT_ID
    GA_CLIENT_ID: 11,               // L - GA_CLIENT_ID
    BUTTON_TEXT: 12,                // M - BUTTON_TEXT
    DATE: 13,                       // N - DATE
    TIME: 14,                       // O - TIME
    DEAL_SOURCE: 15,                // P - R.Источник сделки
    CITY_TAG: 16,                   // Q - R.Тег города
    SOFTWARE: 17,                   // R - ПО
    BAR_NAME: 18,                   // S - Бар (deal)
    BOOKING_DATE: 19,               // T - Дата брони
    GUEST_COUNT: 20,                // U - Кол-во гостей
    VISIT_TIME: 21,                 // V - Время прихода
    COMMENT: 22,                    // W - Комментарий МОБ
    SOURCE: 23,                     // X - Источник
    LEAD_TYPE: 24,                  // Y - Тип лида
    REFUSAL_REASON: 25,             // Z - Причина отказа (ОБ)
    GUEST_STATUS: 26,               // AA - R.Статусы гостей
    REFERRAL_TYPE: 27,              // AB - Сарафан гости
    UTM_MEDIUM: 28,                 // AC - UTM_MEDIUM
    FORMNAME: 29,                   // AD - FORMNAME
    REFERER: 30,                    // AE - REFERER
    FORMID: 31,                     // AF - FORMID
    UTM_SOURCE: 32,                 // AG - UTM_SOURCE
    UTM_TERM: 33,                   // AH - UTM_TERM
    UTM_CAMPAIGN: 34,               // AI - UTM_CAMPAIGN
    UTM_CONTENT: 35,                // AJ - UTM_CONTENT
    UTM_REFERRER: 36,               // AK - utm_referrer
    YM_UID: 37,                     // AL - _ym_uid
    PHONE: 38,                      // AM - Контакт.Телефон
    CONTACT_MANGO: 39,              // AN - Контакт.Номер линии MANGO OFFICE
    NOTES: 40,                      // AO - Примечания(через ;)
    TRACKING_SOURCE: 41,            // AP - R.Источник ТЕЛ (коллтрекинг)
    TRACKING_CHANNEL: 42            // AQ - Канал (коллтрекинг)
    // ... остальные столбцы до 54
  },

  // Динамически определяемые колонки (будет установлено автоматически)
  WORKING_AMO_COLUMNS: null

  // Цветовая схема для отчетов
  COLORS: {
    // Основные цвета
    header: '#4285f4',        // Синий для заголовков
    success: '#34a853',       // Зеленый для успеха
    warning: '#fbbc04',       // Желтый для предупреждений
    error: '#ea4335',         // Красный для ошибок
    info: '#9aa0a6',          // Серый для информации
    
    // Фоновые цвета
    light_blue: '#e8f0fe',    // Светло-синий
    light_green: '#e6f4ea',   // Светло-зеленый
    light_yellow: '#fef7e0',  // Светло-желтый
    light_red: '#fce8e6',     // Светло-красный
    light_gray: '#f8f9fa',    // Светло-серый
    
    // Цвета для графиков
    chart_palette: [
      '#4285f4', '#34a853', '#fbbc04', '#ea4335', 
      '#9aa0a6', '#ff6d00', '#7b1fa2', '#00acc1',
      '#43a047', '#fb8c00', '#8e24aa', '#00897b'
    ],
    
    // Градиенты для тепловых карт
    heatmap_low: '#e8f5e8',   // Низкие значения
    heatmap_medium: '#81c784', // Средние значения  
    heatmap_high: '#2e7d32'   // Высокие значения
  },

  // Настройки форматирования
  FORMATTING: {
    // Форматы дат
    DATE_FORMAT: 'dd.MM.yyyy',
    DATETIME_FORMAT: 'dd.MM.yyyy HH:mm',
    TIME_FORMAT: 'HH:mm',
    
    // Временная зона
    TIMEZONE: 'Europe/Moscow',  // Московское время
    
    // Форматы чисел
    NUMBER_FORMAT: '#,##0',
    CURRENCY_FORMAT: '#,##0 ₽',
    PERCENT_FORMAT: '0.00%',
    
    // Размеры ячеек
    DEFAULT_ROW_HEIGHT: 21,
    HEADER_ROW_HEIGHT: 30,
    DEFAULT_COLUMN_WIDTH: 120
  },

  // Настройки аналитики
  ANALYTICS: {
    // Периоды анализа
    DEFAULT_PERIOD_DAYS: 30,
    MAX_PERIOD_DAYS: 365,
    
    // Пороговые значения
    MIN_CONVERSION_RATE: 0.01,  // 1%
    MIN_QUALITY_SCORE: 10,       // Минимальный балл качества
    MAX_QUALITY_SCORE: 100,      // Максимальный балл качества
    
    // Настройки атрибуции
    ATTRIBUTION_MODELS: {
      FIRST_TOUCH: 'first_touch',
      LAST_TOUCH: 'last_touch', 
      LINEAR: 'linear',
      TIME_DECAY: 'time_decay'
    },
    
    // Весовые коэффициенты для скоринга
    QUALITY_WEIGHTS: {
      CONVERSION_RATE: 0.4,      // 40% - конверсия
      REVENUE_PER_LEAD: 0.3,     // 30% - доход с лида
      LEAD_VOLUME: 0.2,          // 20% - объем лидов
      COST_EFFICIENCY: 0.1       // 10% - эффективность затрат
    }
  },

  // Настройки экспорта/импорта
  EXPORT: {
    // Форматы файлов
    CSV_SEPARATOR: ',',
    CSV_ENCODING: 'UTF-8',
    
    // Ограничения
    MAX_EXPORT_ROWS: 10000,
    MAX_FILE_SIZE_MB: 50,
    
    // Папки в Google Drive
    EXPORT_FOLDER: 'AMO_Analytics_Export',
    BACKUP_FOLDER: 'AMO_Analytics_Backup'
  },

  // Настройки системы
  SYSTEM: {
    // Лимиты производительности
    MAX_EXECUTION_TIME_SEC: 300,  // 5 минут
    BATCH_SIZE: 1000,             // Размер пакета для обработки
    
    // Кэширование
    CACHE_DURATION_MIN: 30,       // 30 минут
    
    // Логирование
    LOG_LEVEL: 'INFO',            // ERROR, WARN, INFO, DEBUG
    MAX_LOG_ENTRIES: 1000,
    
    // Региональные настройки
    DEFAULT_TIMEZONE: 'Europe/Moscow',  // По умолчанию Московское время
    DEFAULT_LOCALE: 'ru_RU'             // Русская локаль
  }
};

/**
 * ДОПОЛНИТЕЛЬНЫЕ КОНСТАНТЫ
 */

// Статусы для определения успешности лидов
const SUCCESS_STATUSES = [
  'Успешно реализовано',
  'Закрыто и реализовано', 
  'Продано',
  'Оплачено',
  'Договор подписан',
  'Сделка заключена',
  'Won',
  'Closed Won',
  'Success'
];

// Маппинг типов каналов
const CHANNEL_TYPES = {
  // Платный трафик
  'yandex_direct': 'Яндекс Директ',
  'google_ads': 'Google Ads',
  'facebook_ads': 'Facebook Ads',
  'vk_ads': 'ВКонтакте Реклама',
  'target_mail': 'Target Mail',
  
  // Органический трафик
  'yandex_organic': 'Яндекс (органика)',
  'google_organic': 'Google (органика)', 
  'bing_organic': 'Bing (органика)',
  
  // Социальные сети
  'vk': 'ВКонтакте',
  'facebook': 'Facebook',
  'instagram': 'Instagram',
  'telegram': 'Telegram',
  'youtube': 'YouTube',
  
  // Другие источники
  'direct': 'Прямые заходы',
  'referral': 'Переходы с сайтов',
  'email': 'Email маркетинг',
  'sms': 'SMS рассылки',
  'phone': 'Телефонные звонки',
  'offline': 'Оффлайн источники',
  'unknown': 'Неопределенный источник'
};

// UTM маппинг для определения источников
const UTM_SOURCE_MAPPING = {
  // Поисковые системы
  'yandex': 'yandex_organic',
  'google': 'google_organic',
  'bing': 'bing_organic',
  'rambler': 'rambler_organic',
  'mail': 'mail_organic',
  
  // Социальные сети  
  'vk.com': 'vk',
  'vkontakte': 'vk',
  'facebook.com': 'facebook',
  'instagram.com': 'instagram',
  'telegram': 'telegram',
  'youtube.com': 'youtube',
  
  // Реклама
  'yandex_ads': 'yandex_direct',
  'google_ads': 'google_ads',
  'facebook_ads': 'facebook_ads',
  'vk_ads': 'vk_ads',
  
  // Прочие
  '(direct)': 'direct',
  'direct': 'direct',
  '(none)': 'direct',
  'referral': 'referral'
};

/**
 * ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ КОНФИГУРАЦИИ
 */

function getChannelTypeBySource(source) {
  if (!source) return 'unknown';
  
  const lowerSource = source.toLowerCase();
  
  // Проверяем точные совпадения
  if (UTM_SOURCE_MAPPING[lowerSource]) {
    return UTM_SOURCE_MAPPING[lowerSource];
  }
  
  // Проверяем частичные совпадения
  for (const [key, value] of Object.entries(UTM_SOURCE_MAPPING)) {
    if (lowerSource.includes(key.toLowerCase())) {
      return value;
    }
  }
  
  return 'unknown';
}

function isSuccessStatus(status) {
  if (!status) return false;
  
  const lowerStatus = status.toLowerCase();
  return SUCCESS_STATUSES.some(successStatus => 
    lowerStatus.includes(successStatus.toLowerCase())
  );
}

function getColorByValue(value, min, max, colorScheme = 'default') {
  if (typeof value !== 'number' || isNaN(value)) return CONFIG.COLORS.light_gray;
  
  const ratio = (value - min) / (max - min);
  
  switch (colorScheme) {
    case 'heatmap':
      if (ratio < 0.33) return CONFIG.COLORS.heatmap_low;
      if (ratio < 0.67) return CONFIG.COLORS.heatmap_medium;
      return CONFIG.COLORS.heatmap_high;
      
    case 'performance':
      if (ratio < 0.5) return CONFIG.COLORS.light_red;
      if (ratio < 0.8) return CONFIG.COLORS.light_yellow;
      return CONFIG.COLORS.light_green;
      
    default:
      if (ratio < 0.33) return CONFIG.COLORS.light_red;
      if (ratio < 0.67) return CONFIG.COLORS.light_yellow;
      return CONFIG.COLORS.light_green;
  }
}

function validateConfiguration() {
  const errors = [];
  
  // Проверяем основные секции
  if (!CONFIG.SHEETS) errors.push('Отсутствует конфигурация SHEETS');
  if (!CONFIG.COLORS) errors.push('Отсутствует конфигурация COLORS');
  
  // Проверяем обязательные листы
  const requiredSheets = ['WORKING_AMO', 'LEADS_CHANNELS', 'AMO_SUMMARY'];
  requiredSheets.forEach(sheet => {
    if (!CONFIG.SHEETS[sheet]) {
      errors.push(`Отсутствует настройка листа: ${sheet}`);
    }
  });
  
  // Проверяем колонки только после инициализации
  if (CONFIG.WORKING_AMO_COLUMNS) {
    const requiredColumns = ['ID', 'NAME', 'STATUS', 'CREATED_AT'];
    requiredColumns.forEach(column => {
      if (CONFIG.WORKING_AMO_COLUMNS[column] === undefined) {
        errors.push(`Отсутствует настройка колонки: ${column}`);
      }
    });
  }
  
  if (errors.length > 0) {
    console.error('Ошибки конфигурации:', errors);
    return false;
  }
  
  console.log('✅ Конфигурация валидна');
  return true;
}

function validateApiTokens() {
  const warnings = [];
  
  // Проверяем токены API
  if (!API_TOKENS.YANDEX_METRIKA_TOKEN || API_TOKENS.YANDEX_METRIKA_TOKEN.startsWith('ВСТАВЬТЕ')) {
    warnings.push('🟡 Не настроен токен Яндекс.Метрики');
  }
  
  if (!API_TOKENS.YANDEX_METRIKA_COUNTER_ID || API_TOKENS.YANDEX_METRIKA_COUNTER_ID.startsWith('ВСТАВЬТЕ')) {
    warnings.push('🟡 Не настроен ID счетчика Яндекс.Метрики');
  }
  
  if (!API_TOKENS.YANDEX_DIRECT_TOKEN || API_TOKENS.YANDEX_DIRECT_TOKEN.startsWith('ВСТАВЬТЕ')) {
    warnings.push('🟡 Не настроен токен Яндекс.Директ');
  }
  
  if (!API_TOKENS.AMO_ACCESS_TOKEN || API_TOKENS.AMO_ACCESS_TOKEN.startsWith('ВСТАВЬТЕ')) {
    warnings.push('🟡 Не настроен токен AMO CRM');
  }
  
  if (!API_TOKENS.AMO_SUBDOMAIN || API_TOKENS.AMO_SUBDOMAIN.startsWith('ВСТАВЬТЕ')) {
    warnings.push('🟡 Не настроен поддомен AMO CRM');
  }
  
  if (warnings.length > 0) {
    console.warn('⚠️ Предупреждения по токенам API:');
    warnings.forEach(warning => console.warn(warning));
    console.warn('📖 См. инструкцию в README.md для настройки токенов');
    return false;
  }
  
  console.log('✅ Все API токены настроены');
  return true;
}

function getApiTokensStatus() {
  const status = {
    yandexMetrika: !!(API_TOKENS.YANDEX_METRIKA_TOKEN && !API_TOKENS.YANDEX_METRIKA_TOKEN.startsWith('ВСТАВЬТЕ')),
    yandexDirect: !!(API_TOKENS.YANDEX_DIRECT_TOKEN && !API_TOKENS.YANDEX_DIRECT_TOKEN.startsWith('ВСТАВЬТЕ')),
    amoCrm: !!(API_TOKENS.AMO_ACCESS_TOKEN && !API_TOKENS.AMO_ACCESS_TOKEN.startsWith('ВСТАВЬТЕ')),
    callTracking: !!(API_TOKENS.MANGO_API_KEY && !API_TOKENS.MANGO_API_KEY.startsWith('ВСТАВЬТЕ'))
  };
  
  const configuredCount = Object.values(status).filter(Boolean).length;
  const totalCount = Object.keys(status).length;
  
  return {
    ...status,
    configuredCount,
    totalCount,
    percentage: Math.round(configuredCount / totalCount * 100)
  };
}

/**
 * БЕЗОПАСНОЕ ФОРМАТИРОВАНИЕ ДАТ
 */
function safeFormatDate(date, format = null, timezone = null) {
  try {
    if (!date) return '';
    
    // Используем настройки по умолчанию
    const dateFormat = format || CONFIG.FORMATTING.DATE_FORMAT;
    const timeZone = timezone || CONFIG.FORMATTING.TIMEZONE || CONFIG.SYSTEM.DEFAULT_TIMEZONE;
    
    // Преобразуем в Date объект если нужно
    let dateObj;
    if (date instanceof Date) {
      dateObj = date;
    } else if (typeof date === 'string') {
      dateObj = new Date(date);
    } else if (typeof date === 'number') {
      dateObj = new Date(date);
    } else {
      return String(date);
    }
    
    // Проверяем валидность даты
    if (isNaN(dateObj.getTime())) {
      console.warn('Неверная дата:', date);
      return String(date);
    }
    
    // Форматируем с учетом временной зоны
    return Utilities.formatDate(dateObj, timeZone, dateFormat);
    
  } catch (error) {
    console.warn('Ошибка форматирования даты:', date, error.message);
    return String(date);
  }
}

/**
 * АВТОМАТИЧЕСКОЕ ОПРЕДЕЛЕНИЕ СТРУКТУРЫ ТАБЛИЦЫ
 */
function detectTableStructure() {
  console.log('🔍 Определяем структуру таблицы...');
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.WORKING_AMO);
    
    if (!sheet) {
      console.log('❌ Лист РАБОЧИЙ_АМО не найден, используем правильную структуру по умолчанию');
      CONFIG.WORKING_AMO_COLUMNS = CONFIG.WORKING_AMO_COLUMNS_CORRECT;
      return 'default';
    }
    
    const lastColumn = sheet.getLastColumn();
    
    if (lastColumn === 0) {
      console.log('❌ Таблица пуста, используем правильную структуру по умолчанию');
      CONFIG.WORKING_AMO_COLUMNS = CONFIG.WORKING_AMO_COLUMNS_CORRECT;
      return 'empty';
    }
    
    const headers = sheet.getRange(1, 1, 1, Math.min(lastColumn, 10)).getValues()[0];
    
    console.log(`📊 Обнаружено столбцов: ${lastColumn}`);
    console.log('📋 Заголовки:', headers.slice(0, 5).join(', ') + (headers.length > 5 ? '...' : ''));
    
    // Определяем структуру по количеству столбцов и ключевым заголовкам
    if (lastColumn === 41 && headers[0] === 'Сделка.ID' && headers[1] === 'Сделка.Название') {
      console.log('✅ Обнаружена ПРАВИЛЬНАЯ структура (41 столбец)');
      CONFIG.WORKING_AMO_COLUMNS = CONFIG.WORKING_AMO_COLUMNS_CORRECT;
      return 'correct';
    } 
    else if (lastColumn >= 50 && headers[0] === 'ID' && headers[1] === 'Название') {
      console.log('⚠️ Обнаружена НЕПРАВИЛЬНАЯ структура (54+ столбцов)');
      CONFIG.WORKING_AMO_COLUMNS = CONFIG.WORKING_AMO_COLUMNS_INCORRECT;
      return 'incorrect';
    }
    else {
      console.log('❓ Неизвестная структура таблицы');
      console.log('🔧 Используем правильную структуру по умолчанию');
      CONFIG.WORKING_AMO_COLUMNS = CONFIG.WORKING_AMO_COLUMNS_CORRECT;
      return 'unknown';
    }
    
  } catch (error) {
    console.error('❌ Ошибка определения структуры:', error);
    console.log('🔧 Используем правильную структуру по умолчанию');
    CONFIG.WORKING_AMO_COLUMNS = CONFIG.WORKING_AMO_COLUMNS_CORRECT;
    return 'error';
  }
}

/**
 * ФУНКЦИЯ ПРОВЕРКИ И ИСПРАВЛЕНИЯ СТРУКТУРЫ
 */
function validateAndFixStructure() {
  console.log('🔧 Проверяем и исправляем структуру...');
  
  const detectedType = detectTableStructure();
  
  if (detectedType === 'incorrect') {
    console.log('⚠️ ВНИМАНИЕ: Обнаружена неправильная структура таблицы!');
    console.log('💡 РЕКОМЕНДАЦИЯ: Используйте правильную структуру из 41 столбца');
    console.log('📖 Список правильных столбцов:');
    
    const correctHeaders = [
      'Сделка.ID', 'Сделка.Название', 'Сделка.Статус', 'Сделка.Причина отказа (ОБ)',
      'Сделка.Тип лида', 'Сделка.R.Статусы гостей', 'Сделка.Ответственный', 'Сделка.Теги',
      'Сделка.Бюджет', 'Сделка.Дата создания', 'Сделка.Дата закрытия', 'Контакт.Номер линии MANGO OFFICE',
      'Сделка.Номер линии MANGO OFFICE', 'Контакт.ФИО', 'Контакт.Телефон'
    ];
    
    correctHeaders.forEach((header, index) => {
      console.log(`${String.fromCharCode(65 + index)} - ${header}`);
    });
    
    return false;
  }
  
  console.log('✅ Структура таблицы настроена корректно');
  return true;
}

// Автоматическая валидация при загрузке
try {
  // Сначала определяем структуру таблицы
  detectTableStructure();
  
  // Затем валидируем конфигурацию
  validateConfiguration();
  const tokenStatus = validateApiTokens();
  
  console.log('🔧 Конфигурация AMO Analytics загружена успешно');
  
  if (!tokenStatus) {
    console.log('💡 Для полной функциональности настройте API токены в начале файла 02_Конфигурация.gs');
    console.log('📖 Подробная инструкция в README.md');
  }
} catch (error) {
  console.error('❌ Ошибка при загрузке конфигурации:', error);
}
