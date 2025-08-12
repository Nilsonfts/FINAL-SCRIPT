/**
 * 🔗 Утилиты UTM - Анализ UTM меток
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Apps Script: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit
 * 2. Создайте новый файл: + → Script file
 * 3. Назовите файл: 06_Утилиты_UTM (без .gs)
 * 4. Удалите весь код по умолчанию
 * 5. Скопируйте и вставьте код ниже
 * 6. Сохраните файл (Ctrl+S)
 */

/**
 * УТИЛИТЫ ДЛЯ РАБОТЫ С UTM МЕТКАМИ И КАНАЛАМИ
 * Парсинг, валидация, канонизация источников трафика
 * @fileoverview Обработка UTM меток и определение каналов привлечения
 */

/**
 * Парсит и валидирует UTM метки
 * @param {Object} utmData - Объект с UTM данными
 * @returns {Object} Обработанные UTM метки
 * 
 * @example
 * parseUtmTags_({
 *   utm_source: 'yandex',
 *   utm_medium: 'cpc',
 *   utm_campaign: 'brand_2024'
 * })
 * // { source: 'yandex', medium: 'cpc', campaign: 'brand_2024', ... }
 */
function parseUtmTags_(utmData) {
  const {
    utm_source = '',
    utm_medium = '',
    utm_campaign = '',
    utm_content = '',
    utm_term = ''
  } = utmData;
  
  return {
    source: cleanUtmValue_(utm_source),
    medium: cleanUtmValue_(utm_medium),
    campaign: cleanUtmValue_(utm_campaign),
    content: cleanUtmValue_(utm_content),
    term: cleanUtmValue_(utm_term)
  };
}

/**
 * Очищает UTM значение от мусора
 * @param {string} value - UTM значение
 * @returns {string|null} Очищенное значение
 */
function cleanUtmValue_(value) {
  if (!value) return null;
  
  const cleaned = String(value)
    .toLowerCase()
    .trim()
    // Убираем кавычки
    .replace(/['"]/g, '')
    // Убираем лишние пробелы
    .replace(/\s+/g, ' ')
    .trim();
  
  // Возвращаем null для пустых или мусорных значений
  if (!cleaned || 
      cleaned === 'null' || 
      cleaned === 'undefined' || 
      cleaned === '(not set)' ||
      cleaned === '(direct)' ||
      cleaned === 'n/a' ||
      cleaned.length < 2) {
    return null;
  }
  
  return cleaned;
}

/**
 * Определяет канал по UTM меткам с учётом fallback логики
 * @param {Object} utmData - UTM данные
 * @param {Object} fallbackData - Дополнительные данные для fallback
 * @returns {string} Канонизированное название канала
 */
function determineChannel_(utmData, fallbackData = {}) {
  const utm = parseUtmTags_(utmData);
  const {
    referrer = '',
    gclid = '',
    yclid = '',
    ysclid = '',
    tel_source = ''
  } = fallbackData;
  
  // 1. Прямое определение по UTM_SOURCE
  if (utm.source) {
    const channelBySource = findChannelByUtmSource_(utm.source);
    if (channelBySource !== 'Прочее') {
      return channelBySource;
    }
  }
  
  // 2. Определение по комбинации source + medium
  if (utm.source && utm.medium) {
    const channelByCombo = findChannelBySourceMedium_(utm.source, utm.medium);
    if (channelByCombo !== 'Прочее') {
      return channelByCombo;
    }
  }
  
  // 3. Fallback по click ID
  if (gclid) return 'Google Ads';
  if (yclid || ysclid) return 'Яндекс.Директ';
  
  // 4. Fallback по телефонному источнику
  if (tel_source) {
    const channelByTel = findChannelByTelSource_(tel_source);
    if (channelByTel !== 'Прочее') {
      return channelByTel;
    }
  }
  
  // 5. Fallback по referrer
  if (referrer) {
    const channelByReferrer = findChannelByReferrer_(referrer);
    if (channelByReferrer !== 'Прочее') {
      return channelByReferrer;
    }
  }
  
  // 6. Определение по medium
  if (utm.medium) {
    return findChannelByMedium_(utm.medium);
  }
  
  // 7. По умолчанию
  return tel_source ? 'Колл-центр' : 'Основной сайт';
}

/**
 * Находит канал по UTM source
 * @param {string} source - UTM source
 * @returns {string} Название канала
 */
function findChannelByUtmSource_(source) {
  const lowerSource = source.toLowerCase();
  
  // Проверяем все каналы из конфигурации
  for (const [channel, patterns] of Object.entries(CONFIG.CHANNEL_MAPPING)) {
    if (patterns.some(pattern => 
      lowerSource.includes(pattern.toLowerCase()) || 
      pattern.toLowerCase().includes(lowerSource)
    )) {
      return channel;
    }
  }
  
  // Дополнительные проверки для специальных случаев
  if (lowerSource.includes('google') || lowerSource.includes('adwords')) {
    return 'Google Ads';
  }
  
  if (lowerSource.includes('yandex') || lowerSource.includes('ya')) {
    return 'Яндекс.Директ';
  }
  
  if (lowerSource.includes('vk') || lowerSource.includes('вконтакте')) {
    return 'ВКонтакте';
  }
  
  if (lowerSource.includes('facebook') || lowerSource.includes('instagram') || lowerSource.includes('meta')) {
    return 'Facebook';
  }
  
  return 'Прочее';
}

/**
 * Находит канал по комбинации source + medium
 * @param {string} source - UTM source
 * @param {string} medium - UTM medium
 * @returns {string} Название канала
 */
function findChannelBySourceMedium_(source, medium) {
  const lowerSource = source.toLowerCase();
  const lowerMedium = medium.toLowerCase();
  
  // CPC трафик
  if (['cpc', 'paid', 'ppc'].includes(lowerMedium)) {
    if (lowerSource.includes('yandex') || lowerSource.includes('ya')) {
      return 'Яндекс.Директ';
    }
    if (lowerSource.includes('google')) {
      return 'Google Ads';
    }
    if (lowerSource.includes('vk') || lowerSource.includes('вконтакте')) {
      return 'ВКонтакте';
    }
    return 'Реклама';
  }
  
  // Органический трафик
  if (['organic', 'natural'].includes(lowerMedium)) {
    return 'Основной сайт';
  }
  
  // Социальные сети
  if (['social', 'smm'].includes(lowerMedium)) {
    return findChannelByUtmSource_(source);
  }
  
  // Email рассылки
  if (['email', 'newsletter'].includes(lowerMedium)) {
    return 'Email';
  }
  
  // Referral трафик
  if (['referral', 'link'].includes(lowerMedium)) {
    return 'Партнёры';
  }
  
  return 'Прочее';
}

/**
 * Находит канал по medium
 * @param {string} medium - UTM medium
 * @returns {string} Название канала
 */
function findChannelByMedium_(medium) {
  const lowerMedium = medium.toLowerCase();
  
  const mediumMapping = {
    'cpc': 'Реклама',
    'paid': 'Реклама',
    'ppc': 'Реклама',
    'organic': 'Основной сайт',
    'natural': 'Основной сайт',
    'social': 'Социальные сети',
    'smm': 'Социальные сети',
    'email': 'Email',
    'newsletter': 'Email',
    'referral': 'Партнёры',
    'affiliate': 'Партнёры',
    'direct': 'Прямые заходы',
    'none': 'Прямые заходы'
  };
  
  return mediumMapping[lowerMedium] || 'Прочее';
}

/**
 * Находит канал по телефонному источнику
 * @param {string} telSource - Источник телефонного звонка
 * @returns {string} Название канала
 */
function findChannelByTelSource_(telSource) {
  if (!telSource) return 'Прочее';
  
  const lowerTel = telSource.toLowerCase();
  
  if (lowerTel.includes('яндекс') || lowerTel.includes('yandex')) {
    return 'Яндекс.Директ';
  }
  
  if (lowerTel.includes('google') || lowerTel.includes('гугл')) {
    return 'Google Ads';
  }
  
  if (lowerTel.includes('2гис') || lowerTel.includes('2gis')) {
    return '2ГИС';
  }
  
  if (lowerTel.includes('сайт') || lowerTel.includes('site')) {
    return 'Основной сайт';
  }
  
  return 'Колл-центр';
}

/**
 * Находит канал по referrer
 * @param {string} referrer - HTTP referrer
 * @returns {string} Название канала
 */
function findChannelByReferrer_(referrer) {
  if (!referrer) return 'Прочее';
  
  const lowerRef = referrer.toLowerCase();
  
  if (lowerRef.includes('yandex.ru') || lowerRef.includes('ya.ru')) {
    return 'Яндекс';
  }
  
  if (lowerRef.includes('google.')) {
    return 'Google';
  }
  
  if (lowerRef.includes('vk.com') || lowerRef.includes('vkontakte.ru')) {
    return 'ВКонтакте';
  }
  
  if (lowerRef.includes('facebook.com') || lowerRef.includes('instagram.com')) {
    return 'Facebook';
  }
  
  if (lowerRef.includes('2gis.ru') || lowerRef.includes('2gis.com')) {
    return '2ГИС';
  }
  
  if (lowerRef.includes('telegram') || lowerRef.includes('t.me')) {
    return 'Telegram';
  }
  
  return 'Партнёры';
}

/**
 * Валидирует UTM кампанию на корректность
 * @param {string} campaign - Название кампании
 * @returns {boolean} true если кампания валидна
 */
function isValidUtmCampaign_(campaign) {
  if (!campaign) return false;
  
  const cleaned = cleanUtmValue_(campaign);
  if (!cleaned) return false;
  
  // Проверяем на мусорные значения
  const trashValues = [
    'test', 'тест', 'default', 'campaign', 'кампания',
    '123', 'temp', 'временная', 'new', 'новая'
  ];
  
  return !trashValues.includes(cleaned.toLowerCase());
}

/**
 * Извлекает данные кампании из UTM меток
 * @param {Object} utmData - UTM данные
 * @returns {Object} Данные кампании
 */
function extractCampaignData_(utmData) {
  const utm = parseUtmTags_(utmData);
  
  // Пытаемся извлечь полезную информацию из названия кампании
  let campaignType = 'Прочее';
  let target = null;
  let geo = null;
  
  if (utm.campaign) {
    const campaign = utm.campaign.toLowerCase();
    
    // Определяем тип кампании
    if (campaign.includes('brand') || campaign.includes('бренд')) {
      campaignType = 'Брендинг';
    } else if (campaign.includes('search') || campaign.includes('поиск')) {
      campaignType = 'Поиск';
    } else if (campaign.includes('display') || campaign.includes('медийная')) {
      campaignType = 'Медийная';
    } else if (campaign.includes('retarget') || campaign.includes('ретаргет')) {
      campaignType = 'Ретаргетинг';
    } else if (campaign.includes('dynamic') || campaign.includes('динамич')) {
      campaignType = 'Динамическая';
    }
    
    // Извлекаем географию
    const geoPatterns = ['msk', 'москва', 'spb', 'питер', 'екб', 'нск', 'regions', 'регионы'];
    for (const pattern of geoPatterns) {
      if (campaign.includes(pattern)) {
        geo = pattern;
        break;
      }
    }
    
    // Извлекаем целевое действие
    if (campaign.includes('order') || campaign.includes('заказ')) {
      target = 'Заказ';
    } else if (campaign.includes('call') || campaign.includes('звонок')) {
      target = 'Звонок';
    } else if (campaign.includes('form') || campaign.includes('форма')) {
      target = 'Форма';
    }
  }
  
  return {
    originalCampaign: utm.campaign,
    campaignType: campaignType,
    target: target,
    geo: geo,
    isValid: isValidUtmCampaign_(utm.campaign)
  };
}

/**
 * Создаёт сводку по UTM данным для отчётности
 * @param {Array} utmDataArray - Массив UTM данных
 * @returns {Object} Сводка по UTM
 */
function createUtmSummary_(utmDataArray) {
  const summary = {
    totalRecords: utmDataArray.length,
    channels: {},
    sources: {},
    mediums: {},
    campaigns: {},
    validUtm: 0,
    invalidUtm: 0
  };
  
  utmDataArray.forEach(item => {
    const channel = determineChannel_(item);
    const utm = parseUtmTags_(item);
    
    // Подсчёт по каналам
    summary.channels[channel] = (summary.channels[channel] || 0) + 1;
    
    // Подсчёт по источникам
    if (utm.source) {
      summary.sources[utm.source] = (summary.sources[utm.source] || 0) + 1;
    }
    
    // Подсчёт по medium
    if (utm.medium) {
      summary.mediums[utm.medium] = (summary.mediums[utm.medium] || 0) + 1;
    }
    
    // Подсчёт по кампаниям
    if (utm.campaign && isValidUtmCampaign_(utm.campaign)) {
      summary.campaigns[utm.campaign] = (summary.campaigns[utm.campaign] || 0) + 1;
      summary.validUtm++;
    } else {
      summary.invalidUtm++;
    }
  });
  
  // Сортируем по убыванию
  summary.channels = sortObjectByValue_(summary.channels);
  summary.sources = sortObjectByValue_(summary.sources);
  summary.mediums = sortObjectByValue_(summary.mediums);
  summary.campaigns = sortObjectByValue_(summary.campaigns);
  
  return summary;
}

/**
 * Сортирует объект по значениям (убывание)
 * @param {Object} obj - Объект для сортировки
 * @returns {Object} Отсортированный объект
 */
function sortObjectByValue_(obj) {
  return Object.entries(obj)
    .sort(([,a], [,b]) => b - a)
    .reduce((result, [key, value]) => {
      result[key] = value;
      return result;
    }, {});
}

/**
 * Проверяет является ли источник первым касанием
 * @param {Object} utmData - UTM данные
 * @param {string} clientId - ID клиента
 * @param {Date} currentDate - Дата текущего события
 * @returns {boolean} true если это первое касание
 */
function isFirstTouch_(utmData, clientId, currentDate) {
  if (!clientId) return false;
  
  try {
    // Проверяем кэш первых касаний
    const cache = CacheService.getScriptCache();
    const cacheKey = `first_touch_${clientId}`;
    const cachedData = cache.get(cacheKey);
    
    if (cachedData) {
      const firstTouch = JSON.parse(cachedData);
      return firstTouch.date === currentDate.getTime();
    }
    
    // Если нет в кэше - считаем первым касанием
    // В реальной системе здесь была бы проверка в БД
    const firstTouchData = {
      date: currentDate.getTime(),
      channel: determineChannel_(utmData),
      utm: parseUtmTags_(utmData)
    };
    
    // Сохраняем в кэш на 30 дней
    cache.put(cacheKey, JSON.stringify(firstTouchData), 2592000);
    
    return true;
    
  } catch (error) {
    logError_('FIRST_TOUCH', `Ошибка проверки первого касания для клиента ${clientId}`, error);
    return false;
  }
}

/**
 * Нормализует название кампании для группировки
 * @param {string} campaign - Название кампании
 * @returns {string} Нормализованное название
 */
function normalizeCampaignName_(campaign) {
  if (!campaign) return 'Не указано';
  
  let normalized = cleanUtmValue_(campaign);
  if (!normalized) return 'Не указано';
  
  // Убираем временные метки и версии
  normalized = normalized
    .replace(/(_v\d+|_\d{4}|\d{6,8})/g, '')
    .replace(/(test|тест|temp|временная)/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  
  return normalized || 'Не указано';
}

/**
 * Создаёт отчёт по эффективности UTM меток
 * @param {Array} dealsData - Данные по сделкам с UTM
 * @returns {Array} Отчёт по UTM эффективности
 */
function createUtmEfficiencyReport_(dealsData) {
  const utmStats = {};
  
  dealsData.forEach(deal => {
    const channel = determineChannel_(deal);
    const utm = parseUtmTags_(deal);
    const key = `${channel}|${utm.source || 'unknown'}|${utm.medium || 'unknown'}|${normalizeCampaignName_(utm.campaign)}`;
    
    if (!utmStats[key]) {
      utmStats[key] = {
        channel: channel,
        source: utm.source,
        medium: utm.medium,
        campaign: normalizeCampaignName_(utm.campaign),
        leads: 0,
        deals: 0,
        revenue: 0,
        conversion: 0
      };
    }
    
    utmStats[key].leads++;
    
    if (deal.status === 'success' || deal.is_paid) {
      utmStats[key].deals++;
      utmStats[key].revenue += parseFloat(deal.budget) || 0;
    }
  });
  
  // Вычисляем конверсию
  const result = Object.values(utmStats).map(stat => ({
    ...stat,
    conversion: stat.leads > 0 ? (stat.deals / stat.leads * 100).toFixed(2) + '%' : '0%'
  }));
  
  // Сортируем по количеству лидов
  return result.sort((a, b) => b.leads - a.leads);
}
