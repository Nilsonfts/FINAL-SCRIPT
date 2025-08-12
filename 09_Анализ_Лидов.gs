/**
 * АНАЛИЗ ЛИДОВ ПО КАНАЛАМ
 * Детальный анализ качества и характеристик лидов в разрезе каналов привлечения
 * @fileoverview Модуль анализа конверсии, времени жизни, источников лидов по каналам
 */

/**
 * Основная функция анализа лидов по каналам
 * Создаёт детальный отчёт по качеству лидов в каждом канале
 */
function analyzeLeadsByChannels() {
  try {
    logInfo_('LEADS_ANALYSIS', 'Начало анализа лидов по каналам');
    
    const startTime = new Date();
    
    // Получаем данные для анализа лидов
    const leadsData = getLeadsAnalysisData_();
    
    if (!leadsData || leadsData.channelLeads.length === 0) {
      logWarning_('LEADS_ANALYSIS', 'Нет данных для анализа лидов');
      createEmptyLeadsReport_();
      return;
    }
    
    // Создаём или обновляем лист
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.LEADS_ANALYSIS);
    clearSheetData_(sheet);
    applySheetFormatting_(sheet, 'Анализ лидов по каналам');
    
    // Строим структуру отчёта
    createLeadsAnalysisStructure_(sheet, leadsData);
    
    // Добавляем диаграммы
    addLeadsAnalysisCharts_(sheet, leadsData);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('LEADS_ANALYSIS', `Анализ лидов завершён за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.LEADS_ANALYSIS);
    
  } catch (error) {
    logError_('LEADS_ANALYSIS', 'Ошибка анализа лидов по каналам', error);
    throw error;
  }
}

/**
 * Получает данные для анализа лидов по каналам
 * @returns {Object} Данные анализа лидов
 * @private
 */
function getLeadsAnalysisData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return null;
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Анализ лидов по каналам
  const channelLeads = analyzeChannelLeadsQuality_(headers, rows);
  
  // Временной анализ лидов
  const timeAnalysis = analyzeLeadsTimePatterns_(headers, rows);
  
  // Источники лидов
  const sourceAnalysis = analyzeLeadsSourceBreakdown_(headers, rows);
  
  // Качественные характеристики
  const qualityMetrics = analyzeLeadsQualityMetrics_(headers, rows);
  
  // Воронка конверсии
  const conversionFunnel = analyzeLeadsConversionFunnel_(headers, rows);
  
  return {
    channelLeads: channelLeads,
    timeAnalysis: timeAnalysis,
    sourceAnalysis: sourceAnalysis,
    qualityMetrics: qualityMetrics,
    conversionFunnel: conversionFunnel,
    totalLeads: rows.length,
    analysisDate: getCurrentDateMoscow_()
  };
}

/**
 * Анализирует качество лидов по каналам
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Данные по качеству лидов по каналам
 * @private
 */
function analyzeChannelLeadsQuality_(headers, rows) {
  const channelMetrics = {};
  
  // Получаем индексы важных колонок
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const closedIndex = findHeaderIndex_(headers, 'Дата закрытия');
  const phoneIndex = findHeaderIndex_(headers, 'Телефон');
  const emailIndex = findHeaderIndex_(headers, 'Email');
  const nameIndex = findHeaderIndex_(headers, 'Имя');
  const companyIndex = findHeaderIndex_(headers, 'Компания');
  
  // Анализируем каждого лида
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const createdDate = parseDate_(row[createdIndex]);
    const closedDate = parseDate_(row[closedIndex]);
    const phone = row[phoneIndex] || '';
    const email = row[emailIndex] || '';
    const name = row[nameIndex] || '';
    const company = row[companyIndex] || '';
    
    // Инициализируем канал если нет
    if (!channelMetrics[channel]) {
      channelMetrics[channel] = {
        channel: channel,
        totalLeads: 0,
        qualifiedLeads: 0,
        hotLeads: 0,
        coldLeads: 0,
        duplicateLeads: 0,
        avgDealValue: 0,
        avgConversionTime: 0,
        totalRevenue: 0,
        completenessScore: 0,
        phoneProvided: 0,
        emailProvided: 0,
        companyProvided: 0,
        conversionTimes: [],
        budgetDistribution: {
          '0-10k': 0,
          '10k-50k': 0,
          '50k-100k': 0,
          '100k+': 0
        }
      };
    }
    
    const metrics = channelMetrics[channel];
    metrics.totalLeads++;
    
    // Анализ полноты данных
    let completenessPoints = 0;
    if (name && String(name).trim()) completenessPoints++;
    if (phone && String(phone).trim()) {
      completenessPoints++;
      metrics.phoneProvided++;
    }
    if (email && String(email).trim()) {
      completenessPoints++;
      metrics.emailProvided++;
    }
    if (company && String(company).trim()) {
      completenessPoints++;
      metrics.companyProvided++;
    }
    
    metrics.completenessScore += completenessPoints;
    
    // Анализ качества лидов
    if (budget > 0 || status === 'success') {
      metrics.qualifiedLeads++;
      
      if (budget > 50000 || (status === 'success' && budget > 0)) {
        metrics.hotLeads++;
      }
    } else if (status === 'failure') {
      metrics.coldLeads++;
    }
    
    // Время конверсии
    if (createdDate && closedDate && status === 'success') {
      const conversionDays = Math.round((closedDate - createdDate) / (1000 * 60 * 60 * 24));
      if (conversionDays >= 0) {
        metrics.conversionTimes.push(conversionDays);
      }
    }
    
    // Выручка
    if (status === 'success') {
      metrics.totalRevenue += budget;
    }
    
    // Распределение по бюджету
    if (budget > 0) {
      if (budget >= 100000) {
        metrics.budgetDistribution['100k+']++;
      } else if (budget >= 50000) {
        metrics.budgetDistribution['50k-100k']++;
      } else if (budget >= 10000) {
        metrics.budgetDistribution['10k-50k']++;
      } else {
        metrics.budgetDistribution['0-10k']++;
      }
    }
  });
  
  // Вычисляем производные метрики
  return Object.values(channelMetrics).map(metrics => {
    // Средняя полнота данных
    metrics.completenessScore = metrics.totalLeads > 0 ? 
      (metrics.completenessScore / metrics.totalLeads / 4 * 100) : 0;
    
    // Средний чек
    const successfulDeals = metrics.totalRevenue > 0 ? metrics.hotLeads || 1 : 0;
    metrics.avgDealValue = successfulDeals > 0 ? metrics.totalRevenue / successfulDeals : 0;
    
    // Среднее время конверсии
    metrics.avgConversionTime = metrics.conversionTimes.length > 0 ?
      metrics.conversionTimes.reduce((a, b) => a + b, 0) / metrics.conversionTimes.length : 0;
    
    // Процентные показатели
    metrics.qualificationRate = metrics.totalLeads > 0 ? 
      (metrics.qualifiedLeads / metrics.totalLeads * 100) : 0;
    metrics.hotLeadsRate = metrics.totalLeads > 0 ? 
      (metrics.hotLeads / metrics.totalLeads * 100) : 0;
    metrics.phoneRate = metrics.totalLeads > 0 ? 
      (metrics.phoneProvided / metrics.totalLeads * 100) : 0;
    metrics.emailRate = metrics.totalLeads > 0 ? 
      (metrics.emailProvided / metrics.totalLeads * 100) : 0;
    
    return metrics;
  }).sort((a, b) => b.qualifiedLeads - a.qualifiedLeads);
}

/**
 * Анализирует временные паттерны лидов
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Временные паттерны
 * @private
 */
function analyzeLeadsTimePatterns_(headers, rows) {
  const patterns = {
    hourly: {},
    daily: {},
    monthly: {},
    seasonal: { 'Q1': 0, 'Q2': 0, 'Q3': 0, 'Q4': 0 }
  };
  
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const createdDate = parseDate_(row[createdIndex]);
    const status = row[statusIndex] || '';
    
    if (!createdDate) return;
    
    // Анализ по часам
    const hour = createdDate.getHours();
    if (!patterns.hourly[hour]) patterns.hourly[hour] = { total: 0, qualified: 0, channels: {} };
    patterns.hourly[hour].total++;
    if (status === 'success') patterns.hourly[hour].qualified++;
    patterns.hourly[hour].channels[channel] = (patterns.hourly[hour].channels[channel] || 0) + 1;
    
    // Анализ по дням недели
    const dayNames = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
    const dayName = dayNames[createdDate.getDay()];
    if (!patterns.daily[dayName]) patterns.daily[dayName] = { total: 0, qualified: 0, channels: {} };
    patterns.daily[dayName].total++;
    if (status === 'success') patterns.daily[dayName].qualified++;
    patterns.daily[dayName].channels[channel] = (patterns.daily[dayName].channels[channel] || 0) + 1;
    
    // Анализ по месяцам
    const monthName = getMonthName_(createdDate.getMonth());
    if (!patterns.monthly[monthName]) patterns.monthly[monthName] = { total: 0, qualified: 0, channels: {} };
    patterns.monthly[monthName].total++;
    if (status === 'success') patterns.monthly[monthName].qualified++;
    patterns.monthly[monthName].channels[channel] = (patterns.monthly[monthName].channels[channel] || 0) + 1;
    
    // Сезонный анализ
    const quarter = Math.floor(createdDate.getMonth() / 3) + 1;
    patterns.seasonal[`Q${quarter}`]++;
  });
  
  return patterns;
}

/**
 * Анализирует источники лидов по каналам
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Анализ источников
 * @private
 */
function analyzeLeadsSourceBreakdown_(headers, rows) {
  const sources = {};
  
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const mediumIndex = findHeaderIndex_(headers, 'UTM Medium');
  const campaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const source = row[sourceIndex] || 'direct';
    const medium = row[mediumIndex] || 'none';
    const campaign = row[campaignIndex] || 'none';
    const status = row[statusIndex] || '';
    
    const sourceKey = `${source}/${medium}`;
    
    if (!sources[channel]) sources[channel] = {};
    if (!sources[channel][sourceKey]) {
      sources[channel][sourceKey] = {
        source: source,
        medium: medium,
        leads: 0,
        conversions: 0,
        campaigns: {}
      };
    }
    
    sources[channel][sourceKey].leads++;
    if (status === 'success') {
      sources[channel][sourceKey].conversions++;
    }
    
    // Кампании в рамках источника
    sources[channel][sourceKey].campaigns[campaign] = 
      (sources[channel][sourceKey].campaigns[campaign] || 0) + 1;
  });
  
  // Преобразуем в удобный для отчёта формат
  const formattedSources = {};
  Object.keys(sources).forEach(channel => {
    formattedSources[channel] = Object.values(sources[channel])
      .map(sourceData => ({
        ...sourceData,
        conversionRate: sourceData.leads > 0 ? 
          (sourceData.conversions / sourceData.leads * 100) : 0,
        topCampaigns: Object.entries(sourceData.campaigns)
          .sort(([,a], [,b]) => b - a)
          .slice(0, 3)
      }))
      .sort((a, b) => b.leads - a.leads);
  });
  
  return formattedSources;
}

/**
 * Анализирует качественные метрики лидов
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Качественные метрики
 * @private
 */
function analyzeLeadsQualityMetrics_(headers, rows) {
  const quality = {
    overallMetrics: {
      totalLeads: rows.length,
      withPhone: 0,
      withEmail: 0,
      withCompany: 0,
      duplicates: 0,
      complete: 0
    },
    channelQuality: {}
  };
  
  const seenContacts = new Set();
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const phoneIndex = findHeaderIndex_(headers, 'Телефон');
  const emailIndex = findHeaderIndex_(headers, 'Email');
  const companyIndex = findHeaderIndex_(headers, 'Компания');
  const nameIndex = findHeaderIndex_(headers, 'Имя');
  
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const phone = row[phoneIndex] || '';
    const email = row[emailIndex] || '';
    const company = row[companyIndex] || '';
    const name = row[nameIndex] || '';
    
    // Общие метрики
    if (phone) quality.overallMetrics.withPhone++;
    if (email) quality.overallMetrics.withEmail++;
    if (company) quality.overallMetrics.withCompany++;
    
    // Полнота профиля (все поля заполнены)
    if (name && phone && email && company) {
      quality.overallMetrics.complete++;
    }
    
    // Дубликаты
    const contactKey = phone || email;
    if (contactKey) {
      if (seenContacts.has(contactKey)) {
        quality.overallMetrics.duplicates++;
      } else {
        seenContacts.add(contactKey);
      }
    }
    
    // По каналам
    if (!quality.channelQuality[channel]) {
      quality.channelQuality[channel] = {
        total: 0,
        withPhone: 0,
        withEmail: 0,
        withCompany: 0,
        complete: 0,
        avgCompleteness: 0
      };
    }
    
    const channelQual = quality.channelQuality[channel];
    channelQual.total++;
    if (phone) channelQual.withPhone++;
    if (email) channelQual.withEmail++;
    if (company) channelQual.withCompany++;
    if (name && phone && email && company) channelQual.complete++;
  });
  
  // Вычисляем процентные показатели по каналам
  Object.keys(quality.channelQuality).forEach(channel => {
    const qual = quality.channelQuality[channel];
    qual.phoneRate = qual.total > 0 ? (qual.withPhone / qual.total * 100) : 0;
    qual.emailRate = qual.total > 0 ? (qual.withEmail / qual.total * 100) : 0;
    qual.companyRate = qual.total > 0 ? (qual.withCompany / qual.total * 100) : 0;
    qual.completenessRate = qual.total > 0 ? (qual.complete / qual.total * 100) : 0;
  });
  
  return quality;
}

/**
 * Анализирует воронку конверсии лидов
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Данные воронки конверсии
 * @private
 */
function analyzeLeadsConversionFunnel_(headers, rows) {
  const funnel = {
    stages: {
      'Новые': 0,
      'В работе': 0,
      'Квалифицированы': 0,
      'Предложение': 0,
      'Закрыты успешно': 0,
      'Закрыты неуспешно': 0
    },
    channelFunnels: {}
  };
  
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    
    // Определяем стадию
    let stage;
    if (status === 'success') {
      stage = 'Закрыты успешно';
    } else if (status === 'failure') {
      stage = 'Закрыты неуспешно';
    } else if (budget > 0) {
      stage = 'Предложение';
    } else if (status === 'qualified') {
      stage = 'Квалифицированы';
    } else if (status === 'in_progress') {
      stage = 'В работе';
    } else {
      stage = 'Новые';
    }
    
    // Общая воронка
    funnel.stages[stage]++;
    
    // Воронка по каналам
    if (!funnel.channelFunnels[channel]) {
      funnel.channelFunnels[channel] = {
        'Новые': 0,
        'В работе': 0,
        'Квалифицированы': 0,
        'Предложение': 0,
        'Закрыты успешно': 0,
        'Закрыты неуспешно': 0
      };
    }
    funnel.channelFunnels[channel][stage]++;
  });
  
  return funnel;
}

/**
 * Создаёт структуру отчёта по анализу лидов
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} leadsData - Данные анализа лидов
 * @private
 */
function createLeadsAnalysisStructure_(sheet, leadsData) {
  let currentRow = 3;
  
  // 1. Общая статистика по лидам
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('📊 ОБЩАЯ СТАТИСТИКА ЛИДОВ');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const overallStats = [
    ['Всего лидов:', leadsData.totalLeads],
    ['Каналов с лидами:', leadsData.channelLeads.length],
    ['Лидов с телефоном:', `${leadsData.qualityMetrics.overallMetrics.withPhone} (${Math.round(leadsData.qualityMetrics.overallMetrics.withPhone/leadsData.totalLeads*100)}%)`],
    ['Лидов с email:', `${leadsData.qualityMetrics.overallMetrics.withEmail} (${Math.round(leadsData.qualityMetrics.overallMetrics.withEmail/leadsData.totalLeads*100)}%)`],
    ['Полные профили:', `${leadsData.qualityMetrics.overallMetrics.complete} (${Math.round(leadsData.qualityMetrics.overallMetrics.complete/leadsData.totalLeads*100)}%)`],
    ['Дубликатов:', leadsData.qualityMetrics.overallMetrics.duplicates]
  ];
  
  sheet.getRange(currentRow, 1, overallStats.length, 2).setValues(overallStats);
  sheet.getRange(currentRow, 1, overallStats.length, 1).setFontWeight('bold');
  currentRow += overallStats.length + 2;
  
  // 2. Качество лидов по каналам
  sheet.getRange(currentRow, 1, 1, 7).merge();
  sheet.getRange(currentRow, 1).setValue('🏆 КАЧЕСТВО ЛИДОВ ПО КАНАЛАМ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 7));
  currentRow++;
  
  const qualityHeaders = [['Канал', 'Всего', 'Квалифицированы', 'Горячие', 'Полнота данных', 'Среднее время сделки', 'Средний чек']];
  sheet.getRange(currentRow, 1, 1, 7).setValues(qualityHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 7));
  currentRow++;
  
  const qualityData = leadsData.channelLeads.slice(0, 10).map(item => [
    item.channel,
    item.totalLeads,
    `${item.qualifiedLeads} (${item.qualificationRate.toFixed(1)}%)`,
    `${item.hotLeads} (${item.hotLeadsRate.toFixed(1)}%)`,
    `${item.completenessScore.toFixed(1)}%`,
    item.avgConversionTime > 0 ? `${Math.round(item.avgConversionTime)} дн.` : 'Н/Д',
    formatCurrency_(item.avgDealValue)
  ]);
  
  if (qualityData.length > 0) {
    sheet.getRange(currentRow, 1, qualityData.length, 7).setValues(qualityData);
    currentRow += qualityData.length;
  }
  currentRow += 2;
  
  // 3. Временные паттерны
  sheet.getRange(currentRow, 1, 1, 4).merge();
  sheet.getRange(currentRow, 1).setValue('⏰ ВРЕМЕННЫЕ ПАТТЕРНЫ ЛИДОВ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
  currentRow++;
  
  // Топ часы
  sheet.getRange(currentRow, 1).setValue('Самые активные часы:');
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  currentRow++;
  
  const topHours = Object.entries(leadsData.timeAnalysis.hourly)
    .sort(([,a], [,b]) => b.total - a.total)
    .slice(0, 5)
    .map(([hour, data]) => [`${hour}:00-${parseInt(hour)+1}:00`, data.total, `${Math.round(data.qualified/data.total*100)}% конверсия`]);
  
  topHours.forEach(hourData => {
    sheet.getRange(currentRow, 1, 1, 3).setValues([hourData]);
    currentRow++;
  });
  currentRow++;
  
  // Топ дни недели
  sheet.getRange(currentRow, 1).setValue('Самые активные дни недели:');
  sheet.getRange(currentRow, 1).setFontWeight('bold');
  currentRow++;
  
  const topDays = Object.entries(leadsData.timeAnalysis.daily)
    .sort(([,a], [,b]) => b.total - a.total)
    .slice(0, 5)
    .map(([day, data]) => [day, data.total, `${Math.round(data.qualified/data.total*100)}% конверсия`]);
  
  topDays.forEach(dayData => {
    sheet.getRange(currentRow, 1, 1, 3).setValues([dayData]);
    currentRow++;
  });
  currentRow += 2;
  
  // 4. Детализация по топ каналам
  const topChannels = leadsData.channelLeads.slice(0, 3);
  topChannels.forEach(channelData => {
    sheet.getRange(currentRow, 1, 1, 6).merge();
    sheet.getRange(currentRow, 1).setValue(`🔍 ДЕТАЛИЗАЦИЯ: ${channelData.channel.toUpperCase()}`);
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
    currentRow++;
    
    // Основные показатели
    const channelDetails = [
      ['Объёмы и качество:', ''],
      [`• Всего лидов`, channelData.totalLeads],
      [`• Квалифицированных`, `${channelData.qualifiedLeads} (${channelData.qualificationRate.toFixed(1)}%)`],
      [`• Горячих лидов`, `${channelData.hotLeads} (${channelData.hotLeadsRate.toFixed(1)}%)`],
      [`• Холодных лидов`, channelData.coldLeads],
      ['Контактная информация:', ''],
      [`• С телефоном`, `${channelData.phoneProvided} (${channelData.phoneRate.toFixed(1)}%)`],
      [`• С email`, `${channelData.emailProvided} (${channelData.emailRate.toFixed(1)}%)`],
      [`• С компанией`, channelData.companyProvided],
      ['Финансовые показатели:', ''],
      [`• Общая выручка`, formatCurrency_(channelData.totalRevenue)],
      [`• Средний чек`, formatCurrency_(channelData.avgDealValue)],
      [`• Среднее время сделки`, channelData.avgConversionTime > 0 ? `${Math.round(channelData.avgConversionTime)} дней` : 'Н/Д']
    ];
    
    sheet.getRange(currentRow, 1, channelDetails.length, 2).setValues(channelDetails);
    
    // Выделяем заголовки разделов
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow + 5, 1).setFontWeight('bold');
    sheet.getRange(currentRow + 9, 1).setFontWeight('bold');
    
    currentRow += channelDetails.length;
    
    // Распределение по бюджетам для канала
    if (Object.values(channelData.budgetDistribution).some(val => val > 0)) {
      sheet.getRange(currentRow, 1).setValue('Распределение по бюджетам:');
      sheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
      
      Object.entries(channelData.budgetDistribution).forEach(([range, count]) => {
        if (count > 0) {
          sheet.getRange(currentRow, 1).setValue(`• ${range} руб.`);
          sheet.getRange(currentRow, 2).setValue(`${count} лидов`);
          currentRow++;
        }
      });
    }
    
    currentRow += 2;
  });
  
  // 5. Воронка конверсии
  sheet.getRange(currentRow, 1, 1, 3).merge();
  sheet.getRange(currentRow, 1).setValue('🎯 ОБЩАЯ ВОРОНКА КОНВЕРСИИ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  const funnelHeaders = [['Стадия', 'Количество', 'Процент от общего']];
  sheet.getRange(currentRow, 1, 1, 3).setValues(funnelHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  const totalFunnelLeads = Object.values(leadsData.conversionFunnel.stages).reduce((a, b) => a + b, 0);
  const funnelData = Object.entries(leadsData.conversionFunnel.stages)
    .filter(([, count]) => count > 0)
    .map(([stage, count]) => [
      stage,
      count,
      `${Math.round(count / totalFunnelLeads * 100)}%`
    ]);
  
  if (funnelData.length > 0) {
    sheet.getRange(currentRow, 1, funnelData.length, 3).setValues(funnelData);
    currentRow += funnelData.length;
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 7);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 7);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к анализу лидов
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} leadsData - Данные анализа лидов
 * @private
 */
function addLeadsAnalysisCharts_(sheet, leadsData) {
  try {
    // 1. Круговая диаграмма качества лидов
    if (leadsData.channelLeads.length > 0) {
      const qualityChartData = [['Канал', 'Квалифицированные лиды']].concat(
        leadsData.channelLeads.slice(0, 8).map(item => [item.channel, item.qualifiedLeads])
      );
      
      createChart_(sheet, 'pie', qualityChartData, {
        startCol: 9,
        title: 'Распределение квалифицированных лидов по каналам',
        position: { row: 3, col: 9 },
        legend: 'right'
      });
    }
    
    // 2. Столбчатая диаграмма полноты данных по каналам
    if (leadsData.channelLeads.length > 0) {
      const completenessChartData = [['Канал', 'Полнота данных %']].concat(
        leadsData.channelLeads.slice(0, 10).map(item => [item.channel, item.completenessScore])
      );
      
      createChart_(sheet, 'column', completenessChartData, {
        startCol: 12,
        title: 'Полнота данных лидов по каналам',
        position: { row: 3, col: 15 },
        legend: 'none',
        hAxisTitle: 'Канал',
        vAxisTitle: 'Полнота данных, %',
        width: 600
      });
    }
    
    // 3. Временное распределение лидов по часам
    if (Object.keys(leadsData.timeAnalysis.hourly).length > 0) {
      const hourlyData = [['Час', 'Количество лидов']];
      
      for (let hour = 0; hour < 24; hour++) {
        const data = leadsData.timeAnalysis.hourly[hour];
        hourlyData.push([`${hour}:00`, data ? data.total : 0]);
      }
      
      createChart_(sheet, 'line', hourlyData, {
        startCol: 15,
        startRow: 1,
        title: 'Распределение лидов по часам дня',
        position: { row: 25, col: 9 },
        legend: 'none',
        hAxisTitle: 'Час дня',
        vAxisTitle: 'Количество лидов',
        width: 700
      });
    }
    
    // 4. Воронка конверсии
    if (Object.keys(leadsData.conversionFunnel.stages).length > 0) {
      const funnelChartData = [['Стадия', 'Количество']].concat(
        Object.entries(leadsData.conversionFunnel.stages)
          .filter(([, count]) => count > 0)
          .map(([stage, count]) => [stage, count])
      );
      
      createChart_(sheet, 'column', funnelChartData, {
        startCol: 18,
        startRow: 1,
        title: 'Воронка конверсии лидов',
        position: { row: 25, col: 18 },
        legend: 'none',
        hAxisTitle: 'Стадия воронки',
        vAxisTitle: 'Количество лидов',
        width: 600
      });
    }
    
    logDebug_('CHARTS', 'Диаграммы анализа лидов добавлены');
    
  } catch (error) {
    logWarning_('CHARTS', 'Ошибка создания диаграмм анализа лидов', error);
    // Продолжаем работу даже если диаграммы не удалось создать
  }
}

/**
 * Создаёт пустой отчёт когда нет данных
 * @private
 */
function createEmptyLeadsReport_() {
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.LEADS_ANALYSIS);
  clearSheetData_(sheet);
  applySheetFormatting_(sheet, 'Анализ лидов по каналам');
  
  sheet.getRange('A3').setValue('ℹ️ Нет данных для анализа лидов по каналам');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A5').setValue('Убедитесь, что:');
  sheet.getRange('A6').setValue('• В системе есть данные о лидах');
  sheet.getRange('A7').setValue('• Данные содержат информацию о каналах привлечения');
  sheet.getRange('A8').setValue('• Выполнена синхронизация данных из CRM');
  
  updateLastUpdateTime_(CONFIG.SHEETS.LEADS_ANALYSIS);
}
