/**
 * 📊 ЦЕНТРАЛЬНЫЙ КОНФИГУРАЦИОННЫЙ ФАЙЛ
 * Все настройки всех скриптов в одном месте
 * Версия: 2.0
 */

const CONFIG = {
  // ===== ОБЩИЕ НАСТРОЙКИ =====
  API_KEY: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  TIMEZONE: 'Europe/Moscow',
  FONT: 'PT Sans',
  
  // ===== ЦВЕТОВАЯ СХЕМА =====
  COLORS: {
    HEADER_BG: '#4285f4',
    HEADER_TEXT: '#ffffff',
    SUBHEADER_BG: '#f1f3f4',
    SUBHEADER_TEXT: '#202124',
    SUCCESS_LIGHT: '#d4edda',
    ERROR_LIGHT: '#f8d7da',
    WARNING_LIGHT: '#fff3cd',
    DATA_BG: '#ffffff',
    
    // Для совместимости с утилитами заголовков
    backgrounds: {
      header: '#4285f4'
    }
  },
  
  // ===== УВЕДОМЛЕНИЯ EMAIL =====
  EMAIL_NOTIFICATIONS: {
    RECIPIENTS: [] // Добавьте email адреса при необходимости
  },
  
  // ===== НАСТРОЙКИ ОТЛАДКИ =====
  DEBUG: {
    enabled: true,
    log_level: 'INFO' // DEBUG, INFO, WARNING, ERROR
  },
  
  // ===== НАСТРОЙКИ GPT =====
  GPT: {
    MODEL: 'gpt-4o-mini',  // Используем более быструю и дешевую модель
    FALLBACK_MODEL: 'gpt-3.5-turbo',
    MAX_TOKENS: 1500,      // Уменьшаем токены
    TEMPERATURE: 0.3
  },
  
  // ===== НАСТРОЙКИ ШРИФТОВ И ФОРМАТИРОВАНИЯ =====
  DEFAULT_FONT: 'PT Sans',
  
  // ===== ГЛОБАЛЬНЫЕ НАСТРОЙКИ ДАННЫХ =====
  RAW_DATA: 'РАБОЧИЙ АМО', // Основной источник данных для всех модулей
  
  // ===== ГЛОБАЛЬНЫЕ ЛИСТЫ ДЛЯ ВСЕХ МОДУЛЕЙ =====
  SHEETS: {
    // Основные рабочие листы
    RAW_DATA: 'РАБОЧИЙ АМО',  // Основной сводный лист со всеми данными
    WORKING_AMO: 'РАБОЧИЙ АМО', // Псевдоним для совместимости
    
    // Аналитические листы
    AMOCRM_SUMMARY: 'СВОДНАЯ АНАЛИТИКА AMOcrm',
    REFUSAL_ANALYSIS: 'Причина отказов',
    CHANNEL_ANALYSIS: 'СРАВНИТЕЛЬНЫЙ АНАЛИЗ',
    LEADS_ANALYSIS: 'Лиды по каналам',
    UTM_ANALYSIS: 'UTM аналитика',
    FIRST_TOUCH_ATTRIBUTION: 'FIRST-TOUCH ANALYSIS',
    DAILY_STATISTICS: 'Ежедневная статистика',
    MANAGER_PERFORMANCE: 'Анализ менеджеров',
    MONTHLY_COMPARISON: 'Сравнение месяцев',
    CLIENT_ANALYSIS: 'Клиентская аналитика',
    BOOKING_ANALYSIS: 'Анализ броней',
    BEAUTY_ANALYTICS: 'Beauty Analytics',
    
    // Источники данных
    AMO_EXPORT: 'Амо Выгрузка',
    AMO_FULL_EXPORT: 'Выгрузка Амо Полная',
    RESERVES_RP: 'Reserves RP',
    GUESTS_RP: 'Guests RP',
    WEBSITE_FORMS: 'Заявки с Сайта',
    CALL_TRACKING: 'КоллТрекинг'
  },
  
  // ===== ГЛОБАЛЬНЫЕ АЛИАСЫ ЗАГОЛОВКОВ =====
  HEADER_ALIASES: {
    // Основные поля сделки
    'ID сделки': ['ID', 'id', 'Сделка.ID', 'Идентификатор', 'Номер'],
    'Название': ['Название', 'Name', 'Сделка.Название', 'Название сделки'],
    'Статус': ['Статус', 'Status', 'Сделка.Статус', 'Этап сделки'],
    'Бюджет': ['Бюджет', 'Budget', 'Сделка.Бюджет', 'Сумма', 'Сумма ₽'],
    'Дата создания': ['Дата создания', 'Created', 'DATE', 'Сделка.Дата создания'],
    'Дата закрытия': ['Дата закрытия', 'Сделка.Дата закрытия'],
    'Ответственный': ['Ответственный', 'Кем создана', 'Автор', 'Создатель', 'Менеджер'],
    
    // Контактная информация
    'Контакт': ['Контакт', 'Основной контакт', 'Контакт.ФИО', 'Контакт.Имя', 'ФИО'],
    'Телефон': ['Телефон', 'Phone', 'Контакт.Телефон', 'Рабочий телефон (контакт)'],
    'Email': ['Email', 'E-mail', 'Почта', 'Контакт.Email'],
    
    // UTM метки
    'UTM Source': ['utm_source', 'UTM_SOURCE', 'Источник', 'Сделка.utm_source'],
    'UTM Medium': ['utm_medium', 'UTM_MEDIUM', 'Канал', 'Сделка.utm_medium'],
    'UTM Campaign': ['utm_campaign', 'UTM_CAMPAIGN', 'Кампания', 'Сделка.utm_campaign'],
    'UTM Term': ['utm_term', 'UTM_TERM', 'Сделка.utm_term'],
    'UTM Content': ['utm_content', 'UTM_CONTENT', 'Сделка.utm_content'],
    
    // Каналы и источники
    'Канал': ['Канал', 'Channel', 'Источник', 'Source'],
    'Причина отказа': ['Причина отказа', 'Отказ', 'Комментарий отказа'],
    'Комментарий': ['Комментарий', 'Примечания', 'Заметки']
  },
  
  // ===== ОСНОВНОЙ СКРИПТ СБОРА ДАННЫХ =====
  mainScript: {
    // ===== ОСНОВНЫЕ ЛИСТЫ ДАННЫХ =====
  SHEETS: {
    // Источники данных AmoCRM
    AMO_EXPORT: 'Амо Выгрузка',           // Первый лист экспорта AmoCRM
    AMO_FULL: 'Выгрузка Амо Полная',      // Полный лист экспорта AmoCRM
    
    // Дополнительные источники данных
    SITE_FORMS: 'Заявки с Сайта',         // Заявки с форм сайта
    RESERVES: 'Reserves RP',              // Данные по резервам
    GUESTS: 'Guests RP',                  // Данные по гостям
    CALL_TRACKING: 'КоллТрекинг',         // Данные колл-трекинга
    
    // Рабочие листы системы
    OUT: 'РАБОЧИЙ АМО',                   // Основной рабочий лист с объединёнными данными
    RAW_DATA: 'РАБОЧИЙ АМО',              // Алиас для аналитических модулей
    
    // Аналитические отчёты
    REFUSAL_ANALYSIS: 'Причина отказов',
    CHANNEL_ANALYSIS: 'СРАВНИТЕЛЬНЫЙ АНАЛИЗ',
    FIRST_TOUCH_ATTRIBUTION: 'FIRST-TOUCH ANALYSIS',
    AMOCRM_SUMMARY: 'СВОДНАЯ АНАЛИТИКА AMOcrm',
    DAILY_STATISTICS: 'Ежедневная статистика',
    LEADS_ANALYSIS: 'Лиды по каналам',
    UTM_ANALYSIS: 'UTM аналитика',
    MANAGER_PERFORMANCE: 'Анализ менеджеров',
    MONTHLY_COMPARISON: 'Сравнение месяцев',
    
    // Служебные листы
    MAIN_DASHBOARD: 'Главная',
    LOGS: 'LOG',
    DIAGNOSTICS: '_DIAG'
  },
    KEY: 'Сделка.ID',
    CLOSED_RE: /(закры|успеш|неуспеш|оплач)/i,
    PHONE_SPLIT_RE: /[;,]/,
    COLORS: { NEW:'#ccffcc', URGENT:'#ffe599', OVERDUE:'#f4cccc', LONG:'#fff2cc' },
    PRESERVE_STYLES: true,
    
    // Маппинг заголовков для сбора данных
    MAP: {
      'Сделка.ID': ['ID','id','Идентификатор','Номер'],
      'Сделка.Название': ['Название','Name','Название сделки'],
      'Сделка.Дата создания': ['Дата создания','Created','DATE'],
      'Сделка.Статус': ['Статус','Status','Этап сделки'],
      'Сделка.Бюджет': ['Бюджет','Budget','Сумма'],
      'Сделка.Дата закрытия': ['Дата закрытия'],
      'Сделка.Теги': ['Теги сделки','Теги'],
      'Кем создана': ['Автор','Создатель'],
      
      // CONTACT
      'Контакт.ФИО': ['Основной контакт','Контакт.Имя','Контакт'],
      'Контакт.Телефон': ['Телефон','Phone','Рабочий телефон (контакт)'],
      'Контакт.Номер линии MANGO OFFICE': ['Номер линии MANGO OFFICE (контакт)'],
      'Сделка.Номер линии MANGO OFFICE': ['Номер линии MANGO OFFICE'],
      
      // UTM
      'Сделка.utm_source': ['utm_source','UTM_SOURCE','Источник'],
      'Сделка.utm_medium': ['utm_medium','UTM_MEDIUM','Канал'],
      'Сделка.utm_campaign': ['utm_campaign','UTM_CAMPAIGN','Кампания'],
      'Сделка.utm_term': ['utm_term','UTM_TERM'],
      'Сделка.utm_content': ['utm_content','UTM_CONTENT'],
      'Сделка.utm_referrer': ['utm_referrer','referer','REFERRER','REFERER'],
      'Сделка.Источник': ['Источник','Source'],
      'Сделка.R.Источник сделки': ['R.Источник сделки'],
      'Сделка.Тип лида': ['Тип лида'],
      
      // ANALYTICS
      'Сделка.YM_CLIENT_ID': ['YM_CLIENT_ID'],
      'Сделка._ym_uid': ['_ym_uid'],
      'Сделка.GA_CLIENT_ID': ['GA_CLIENT_ID'],
      'Сделка.FORMID': ['FORMID'],
      'Сделка.FORMNAME': ['FORMNAME'],
      'Сделка.BUTTON_TEXT': ['BUTTON_TEXT'],
      'Сделка.DATE': ['DATE'],
      'Сделка.TIME': ['TIME'],
      'Сделка.REFERER': ['REFERER']
    }
  },

  // ===== АНАЛИЗ ОТКАЗОВ =====
  refusals: {
    SOURCE_SHEET: 'РАБОЧИЙ АМО',
    OUTPUT_SHEET: 'Причина отказов',
    GPT_MODEL: 'gpt-4o-mini',  // Используем быструю модель
    FALLBACK_MODEL: 'gpt-3.5-turbo',
    MAX_REASONS: 10,
    MIN_FREQUENCY: 2,
    SAMPLE_SIZE: 50,
    
    // ТОЧНЫЕ СТАТУСЫ ОТКАЗОВ ДЛЯ ВАШЕГО ПРОЕКТА
    REFUSAL_STATUSES: [
      'Закрыто и не реализовано'  // Ваш точный статус из колонки "Статус"
    ],
    
    // Колонка со статусом - ищем по названию "Статус"
    STATUS_COLUMN_NAME: 'Статус',
    
    // Колонка с причинами отказов - ТОЧНОЕ название из ваших листов
    REFUSAL_REASON_COLUMN_NAME: 'Причина отказа',
    
    // Альтернативные названия для поиска причин отказов
    REFUSAL_REASON_ALIASES: [
      'Причина отказа',
      'Комментарий МОБ', 
      'Примечания',
      'Комментарий'
    ]
  },

  // ===== СРАВНИТЕЛЬНЫЙ АНАЛИЗ КАНАЛОВ =====
  channels: {
    SOURCE_SHEET: 'РАБОЧИЙ АМО',
    OUTPUT_SHEET: 'СРАВНИТЕЛЬНЫЙ АНАЛИЗ',
    CHART_START_DATE: '2025-04-01',
    GPT_MODEL: 'gpt-4o',
    COLORS: {
      SUCCESS: '#28a745',
      WARNING: '#ffc107', 
      DANGER: '#dc3545',
      INFO: '#17a2b8',
      LIGHT: '#f8f9fa'
    },
    CHANNEL_COLORS: {
      'Основной сайт': '#4285f4',
      'Яндекс': '#ff3333', 
      '2ГИС': '#00b956',
      'Telegram': '#0088cc',
      'Соц сети': '#833ab4'
    }
  },

  // ===== АНАЛИЗ ПЕРВЫХ КАСАНИЙ =====
  firstTouch: {
    RA_SHEET: 'РАБОЧИЙ АМО',
    OUTPUT_SHEET: 'FIRST-TOUCH ANALYSIS',
    DETAIL_LIMIT: 50,
    MIN_CLIENTS: 3,
    GPT_MODEL: 'gpt-4o',
    CHANNEL_MAPPING: {
      'Основной сайт': ['сайт', 'site', 'основной'],
      'Яндекс': ['yandex', 'яндекс', 'direct'],
      '2ГИС': ['2gis', '2гис', 'gis'],
      'Телеграм': ['telegram', 'тг', 't.me'],
      'ВКонтакте': ['vk', 'вк', 'vkontakte']
    }
  },

  // ===== КОЛЛ-ТРЕКИНГ =====
  callTracking: {
    RA_SHEET: 'РАБОЧИЙ АМО',
    DIRECTORY_SHEET: 'КоллТрекинг',
    REPORT_SHEET: 'КоллТрекинг',
    CHANNEL_COLORS: {
      'Основной сайт': '#4285f4',
      'Яндекс': '#ff3333',
      '2ГИС': '#00b956', 
      'Telegram': '#0088cc',
      'Default': '#6c757d'
    }
  },

  // ===== СВОДНАЯ АНАЛИТИКА AMO =====
  amoDashboard: {
    RA_SHEET: 'РАБОЧИЙ АМО',
    REPORT_SHEET: 'СВОДНАЯ АНАЛИТИКА AMOcrm',
    SUCCESS_STATUSES: ['Оплачено', 'Успешно в РП', 'Успешно реализовано'],
    MONTHS_BACK: 12,
    TOP_N: 10
  },

  // ===== ЛИДЫ ПО КАНАЛАМ (ДЕТАЛЬНО) =====
  channelDetails: {
    RA_SHEET: 'РАБОЧИЙ АМО',
    CHANNELS: {
      site: 'Основной сайт + Контекст',
      yandex: 'Яндекс', 
      gis2: '2ГИС',
      telegram: 'Telegram',
      vk: 'ВКонтакте'
    }
  },

  // ===== СКВОЗНАЯ АНАЛИТИКА =====
  e2e: {
    SHEETS: {
      DEALS: 'РАБОЧИЙ АМО',
      SPEND: 'Расходы маркетинга',
      GUESTS: 'Guests RP', 
      DASH: 'СКВОЗНАЯ АНАЛИТИКА'
    },
    SUCCESS_RE: /(оплач|успеш.*в.*рп|успеш.*реализ)/i,
    COLS: {
      created: ['Дата создания','Created','DATE'],
      status: ['Статус','Status'],
      budget: ['Бюджет','Budget','Сумма ₽'],
      telTag: ['R.Источник ТЕЛ сделки'],
      utm: ['utm_source','UTM_SOURCE'],
      phone: ['Телефон','Контакт.Телефон','Phone']
    }
  },

  // ===== БЬЮТИ (АНАЛИЗ КАНАЛОВ ПРИВЛЕЧЕНИЯ) =====
  beauty: {
    RA_SHEET: 'РАБОЧИЙ АМО',
    DIRECTORY_SHEET: 'КоллТрекинг',
    BUDGETS_SHEET: 'БЮДЖЕТЫ',
    ACQ_SHEET: 'АНАЛИЗ КАНАЛОВ ПРИВЛЕЧЕНИЯ'
  },

  // ===== ЕЖЕДНЕВНАЯ СТАТИСТИКА =====
  dailyStats: {
    SOURCE: 'РАБОЧИЙ АМО',
    OUTPUT: 'ЕЖЕДНЕВНАЯ СТАТИСТИКА',
    SUCCESS: ['Оплачено', 'Успешно в РП', 'Успешно реализовано'],
    DAYS_BACK: 30
  },

  // ===== ОТДЕЛ БРОНИРОВАНИЯ =====
  teamPerformance: {
    RA_SHEET: 'РАБОЧИЙ АМО',
    REPORT_SHEET: '⚡ ЭФФЕКТИВНОСТЬ КОМАНДЫ И АНАЛИТИКА',
    TOP_N: 10,
    PERFORMANCE_COLORS: {
      HIGH: '#28a745',
      MEDIUM: '#ffc107',
      LOW: '#dc3545'
    }
  },

  // ===== МАРКЕТИНГ И КАНАЛЫ =====
  marketing: {
    SOURCE: 'РАБОЧИЙ АМО',
    BUDGETS: 'Расходы маркетинга',
    OUTPUT: 'МАРКЕТИНГ И КАНАЛЫ',
    ROI_THRESHOLD: 1.5
  },

  // ===== АНАЛИТИКА ЛИДОВ ПО КАНАЛАМ =====
  leadsByChannel: {
    SOURCE: 'РАБОЧИЙ АМО',
    OUTPUT: 'АНАЛИТИКА ЛИДОВ ПО КАНАЛАМ',
    CHANNEL_MAPPING: {
      'сайт': 'Основной сайт',
      'yandex': 'Яндекс',
      '2gis': '2ГИС',
      'telegram': 'Telegram'
    }
  },

  // ===== ДИНАМИКА ПО МЕСЯЦАМ =====
  monthly: {
    SOURCE: 'РАБОЧИЙ АМО',
    OUTPUT: 'ДИНАМИКА ПО МЕСЯЦАМ',
    MONTHS_BACK: 24
  },

  // ===== ВЫГРУЗКА ЯНДЕКС.ДИРЕКТ =====
  yandexDirect: {
    SOURCE: 'РАБОЧИЙ АМО',
    OUTPUT_DIRECT: 'ВЫГРУЗКА ЯНДЕКС ДИРЕКТ',
    OUTPUT_SKIPPED: 'ПРОПУСКИ ДИРЕКТ',
    PAID_STATUSES: ['ОПЛАЧЕНО', 'Успешно в РП', 'Успешно реализовано'],
    HEADER_ALIASES: {
      ymClientId: ['ym_client_id', 'ym client id', 'ym clientid', 'ym id', 'ymid'],
      status: ['статус'],
      budget: ['бюджет', 'сумма', 'сумма ₽', 'сумма руб', 'сумма р'],
      closeDate: ['дата закрытия', 'дата завершения'],
      reserveDate: ['дата брони', 'дата резерва']
    }
  },

  // ===== TOP-40 КЛИЕНТОВ =====
  top40: {
    SOURCE: 'Guests RP',
    OUTPUT: 'ТОП-40 КЛИЕНТОВ',
    TOP_N: 40,
    MIN_AMOUNT: 1000
  },

  // ===== СКВОЗНАЯ АНАЛИТИКА САЙТ =====
  siteAnalytics: {
    SHEETS: {
      SITE: 'Заявки с Сайта',
      AMO: 'РАБОЧИЙ АМО',
      SITE_DASH: 'Аналитика Сайта',
      CROSS_DASH: 'Сквозная Сайт → AMO'
    },
    ATTRIB_WINDOW_DAYS: 30,
    SUCCESS_RE: /(оплач|успеш.*в.*рп|успеш.*реализ)/i
  },

  // ===== UTM АНАЛИТИКА =====
  utm: {
    SHEETS: {
      SITE: 'Заявки с Сайта',
      RA: 'РАБОЧИЙ АМО',
      DASH: 'UTM АНАЛИТИКА'
    },
    ATTRIB_WINDOW_DAYS: 30,
    SUCCESS_RE: /(оплач|успеш.*в.*рп|успеш.*реализ)/i
  },

  // ===== GPT ПРОМПТЫ =====
  prompts: {
    refusalAnalysis: `Ты — ведущий аналитик продаж с опытом работы 15+ лет.
    
ЗАДАЧА: Проанализируй причины отказов клиентов и дай рекомендации.

КОНТЕКСТ: Ты получишь данные о причинах отказов клиентов от покупки/бронирования в ресторанном бизнесе.

ТРЕБОВАНИЯ К АНАЛИЗУ:
1. Выдели ТОП-5 главных причин отказов
2. Для каждой причины предложи конкретные действия
3. Оцени влияние на конверсию
4. Дай общие рекомендации по улучшению

ФОРМАТ ОТВЕТА:
- Краткий вывод (2-3 предложения)
- Топ причины отказов с процентами
- Конкретные рекомендации
- Приоритетные действия

Пиши деловым языком, конкретно и по делу.`,

    channelAnalysis: `Ты — старший маркетолог-аналитик с опытом performance-маркетинга 10+ лет.
    
ЗАДАЧА: Проанализируй эффективность каналов привлечения и дай стратегические рекомендации.

КОНТЕКСТ: Ресторанный бизнес, многоканальное привлечение клиентов (сайт, Яндекс, 2ГИС, соцсети).

АНАЛИЗИРУЙ:
1. Конверсию по каналам
2. Качество лидов
3. Стоимость привлечения
4. Рентабельность каналов
5. Тренды и динамику

ДАШБОРД ДОЛЖЕН СОДЕРЖАТЬ:
- Ключевые выводы и инсайты
- Рекомендации по бюджетам
- Приоритетные каналы для масштабирования
- Проблемные зоны и способы их решения

Пиши стратегически, с фокусом на ROI и масштабирование.`,

    firstTouchAnalysis: `Ты — директор по маркетингу с экспертизой в атрибуционном моделировании.

ЗАДАЧА: Проанализируй модель первых касаний клиентов с брендом.

ФОКУС:
- Какие каналы лучше всего привлекают новых клиентов
- Путь клиента от первого контакта до покупки
- Эффективность различных точек входа
- Рекомендации по оптимизации воронки

ОТВЕТ СТРУКТУРИРУЙ:
1. Главные инсайты о первых касаниях
2. Самые эффективные каналы входа
3. Паттерны поведения клиентов
4. Рекомендации по улучшению конверсии

Будь конкретным в рекомендациях и цифрах.`
  }
};

// ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ КОНФИГА =====

/**
 * Получить настройки конкретного модуля
 */
function getModuleConfig(moduleName) {
  const moduleConfig = CONFIG[moduleName];
  if (!moduleConfig) {
    throw new Error(`Модуль "${moduleName}" не найден в конфигурации`);
  }
  return moduleConfig;
}

/**
 * Проверить наличие API ключа
 */
function validateApiKey() {
  if (!CONFIG.API_KEY) {
    throw new Error('API ключ OpenAI не найден. Установите OPENAI_API_KEY в Свойствах проекта.');
  }
  return true;
}

/**
 * Получить цвета для канала
 */
function getChannelColor(channelName, module = 'channels') {
  const colors = CONFIG[module]?.CHANNEL_COLORS || CONFIG.channels.CHANNEL_COLORS;
  return colors[channelName] || colors['Default'] || '#6c757d';
}
