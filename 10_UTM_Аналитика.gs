/**
 * UTM АНАЛИТИКА
 * Детальный анализ UTM параметров и их эффективности
 * @fileoverview Модуль анализа UTM меток, источников трафика и кампаний
 */

/**
 * Основная функция UTM аналитики
 * Создаёт детальный отчёт по UTM параметрам и их эффективности
 */
function analyzeUtmPerformance() {
  try {
    logInfo_('UTM_ANALYSIS', 'Начало UTM аналитики');
    
    const startTime = new Date();
    
    // Получаем данные для UTM анализа
    const utmData = getUtmAnalysisData_();
    
    if (!utmData || utmData.sources.length === 0) {
      logWarning_('UTM_ANALYSIS', 'Нет данных для UTM анализа');
      createEmptyUtmReport_();
      return;
    }
    
    // Создаём или обновляем лист
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.UTM_ANALYSIS);
    clearSheetData_(sheet);
    applySheetFormatting_(sheet, 'UTM Аналитика');
    
    // Строим структуру отчёта
    createUtmAnalysisStructure_(sheet, utmData);
    
    // Добавляем диаграммы
    addUtmAnalysisCharts_(sheet, utmData);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('UTM_ANALYSIS', `UTM аналитика завершена за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.UTM_ANALYSIS);
    
  } catch (error) {
    logError_('UTM_ANALYSIS', 'Ошибка UTM аналитики', error);
    throw error;
  }
}

/**
 * Получает данные для UTM анализа
 * @returns {Object} Данные UTM анализа
 * @private
 */
function getUtmAnalysisData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return null;
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Анализ источников (UTM Source)
  const sources = analyzeUtmSources_(headers, rows);
  
  // Анализ медиумов (UTM Medium)
  const mediums = analyzeUtmMediums_(headers, rows);
  
  // Анализ кампаний (UTM Campaign)
  const campaigns = analyzeUtmCampaigns_(headers, rows);
  
  // Анализ терминов (UTM Term)
  const terms = analyzeUtmTerms_(headers, rows);
  
  // Анализ контента (UTM Content)
  const content = analyzeUtmContent_(headers, rows);
  
  // Комбинированный анализ
  const combinations = analyzeUtmCombinations_(headers, rows);
  
  // Временной анализ UTM
  const timeAnalysis = analyzeUtmTimePatterns_(headers, rows);
  
  return {
    sources: sources,
    mediums: mediums,
    campaigns: campaigns,
    terms: terms,
    content: content,
    combinations: combinations,
    timeAnalysis: timeAnalysis,
    totalRecords: rows.length,
    analysisDate: getCurrentDateMoscow_()
  };
}

/**
 * Анализирует UTM Source параметры
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Анализ источников
 * @private
 */
function analyzeUtmSources_(headers, rows) {
  const sourceStats = {};
  
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  
  rows.forEach(row => {
    const source = row[sourceIndex] || 'direct';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const channel = row[channelIndex] || 'Неизвестно';
    const createdDate = parseDate_(row[createdIndex]);
    
    if (!sourceStats[source]) {
      sourceStats[source] = {
        source: source,
        totalLeads: 0,
        successfulDeals: 0,
        totalRevenue: 0,
        avgRevenue: 0,
        conversionRate: 0,
        channels: {},
        monthlyTrend: {}
      };
    }
    
    const stats = sourceStats[source];
    stats.totalLeads++;
    
    if (status === 'success') {
      stats.successfulDeals++;
      stats.totalRevenue += budget;
    }
    
    // Каналы для источника
    stats.channels[channel] = (stats.channels[channel] || 0) + 1;
    
    // Месячные тренды
    if (createdDate) {
      const monthKey = formatDate_(createdDate, 'YYYY-MM');
      stats.monthlyTrend[monthKey] = (stats.monthlyTrend[monthKey] || 0) + 1;
    }
  });
  
  // Вычисляем производные метрики
  return Object.values(sourceStats).map(stats => {
    stats.conversionRate = stats.totalLeads > 0 ? (stats.successfulDeals / stats.totalLeads * 100) : 0;
    stats.avgRevenue = stats.successfulDeals > 0 ? (stats.totalRevenue / stats.successfulDeals) : 0;
    
    // Топ каналы для источника
    stats.topChannels = Object.entries(stats.channels)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 3)
      .map(([channel, count]) => ({ channel, count }));
    
    return stats;
  }).sort((a, b) => b.totalLeads - a.totalLeads);
}

/**
 * Анализирует UTM Medium параметры
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Анализ медиумов
 * @private
 */
function analyzeUtmMediums_(headers, rows) {
  const mediumStats = {};
  
  const mediumIndex = findHeaderIndex_(headers, 'UTM Medium');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  
  rows.forEach(row => {
    const medium = row[mediumIndex] || 'none';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const source = row[sourceIndex] || 'direct';
    
    if (!mediumStats[medium]) {
      mediumStats[medium] = {
        medium: medium,
        totalLeads: 0,
        successfulDeals: 0,
        totalRevenue: 0,
        avgRevenue: 0,
        conversionRate: 0,
        sources: {}
      };
    }
    
    const stats = mediumStats[medium];
    stats.totalLeads++;
    
    if (status === 'success') {
      stats.successfulDeals++;
      stats.totalRevenue += budget;
    }
    
    // Источники для медиума
    stats.sources[source] = (stats.sources[source] || 0) + 1;
  });
  
  // Вычисляем метрики и сортируем
  return Object.values(mediumStats).map(stats => {
    stats.conversionRate = stats.totalLeads > 0 ? (stats.successfulDeals / stats.totalLeads * 100) : 0;
    stats.avgRevenue = stats.successfulDeals > 0 ? (stats.totalRevenue / stats.successfulDeals) : 0;
    
    // Топ источники для медиума
    stats.topSources = Object.entries(stats.sources)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 3)
      .map(([source, count]) => ({ source, count }));
    
    return stats;
  }).sort((a, b) => b.totalLeads - a.totalLeads);
}

/**
 * Анализирует UTM Campaign параметры
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Анализ кампаний
 * @private
 */
function analyzeUtmCampaigns_(headers, rows) {
  const campaignStats = {};
  
  const campaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const mediumIndex = findHeaderIndex_(headers, 'UTM Medium');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  
  rows.forEach(row => {
    const campaign = row[campaignIndex] || 'none';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const source = row[sourceIndex] || 'direct';
    const medium = row[mediumIndex] || 'none';
    const createdDate = parseDate_(row[createdIndex]);
    
    if (!campaignStats[campaign]) {
      campaignStats[campaign] = {
        campaign: campaign,
        totalLeads: 0,
        successfulDeals: 0,
        totalRevenue: 0,
        avgRevenue: 0,
        conversionRate: 0,
        sourceMediumPairs: {},
        recentActivity: null,
        isActive: false
      };
    }
    
    const stats = campaignStats[campaign];
    stats.totalLeads++;
    
    if (status === 'success') {
      stats.successfulDeals++;
      stats.totalRevenue += budget;
    }
    
    // Пары источник/медиум для кампании
    const pairKey = `${source}/${medium}`;
    stats.sourceMediumPairs[pairKey] = (stats.sourceMediumPairs[pairKey] || 0) + 1;
    
    // Последняя активность
    if (createdDate && (!stats.recentActivity || createdDate > stats.recentActivity)) {
      stats.recentActivity = createdDate;
    }
  });
  
  // Определяем активные кампании (последние 30 дней)
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  
  return Object.values(campaignStats).map(stats => {
    stats.conversionRate = stats.totalLeads > 0 ? (stats.successfulDeals / stats.totalLeads * 100) : 0;
    stats.avgRevenue = stats.successfulDeals > 0 ? (stats.totalRevenue / stats.successfulDeals) : 0;
    stats.isActive = stats.recentActivity && stats.recentActivity > thirtyDaysAgo;
    
    // Топ пары источник/медиум
    stats.topPairs = Object.entries(stats.sourceMediumPairs)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 3)
      .map(([pair, count]) => ({ pair, count }));
    
    return stats;
  }).sort((a, b) => b.totalRevenue - a.totalRevenue);
}

/**
 * Анализирует UTM Term параметры
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Анализ терминов
 * @private
 */
function analyzeUtmTerms_(headers, rows) {
  const termStats = {};
  
  const termIndex = findHeaderIndex_(headers, 'UTM Term');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const campaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  
  rows.forEach(row => {
    const term = row[termIndex] || '';
    if (!term) return; // Пропускаем пустые термины
    
    const status = row[statusIndex] || '';
    const campaign = row[campaignIndex] || 'none';
    
    if (!termStats[term]) {
      termStats[term] = {
        term: term,
        totalLeads: 0,
        conversions: 0,
        conversionRate: 0,
        campaigns: {}
      };
    }
    
    const stats = termStats[term];
    stats.totalLeads++;
    
    if (status === 'success') {
      stats.conversions++;
    }
    
    stats.campaigns[campaign] = (stats.campaigns[campaign] || 0) + 1;
  });
  
  return Object.values(termStats).map(stats => {
    stats.conversionRate = stats.totalLeads > 0 ? (stats.conversions / stats.totalLeads * 100) : 0;
    
    // Топ кампании для термина
    stats.topCampaigns = Object.entries(stats.campaigns)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 2)
      .map(([campaign, count]) => ({ campaign, count }));
    
    return stats;
  }).filter(stats => stats.totalLeads >= 5) // Только термины с минимальным объёмом
    .sort((a, b) => b.totalLeads - a.totalLeads)
    .slice(0, 20); // Топ 20 терминов
}

/**
 * Анализирует UTM Content параметры
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Анализ контента
 * @private
 */
function analyzeUtmContent_(headers, rows) {
  const contentStats = {};
  
  const contentIndex = findHeaderIndex_(headers, 'UTM Content');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const campaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  
  rows.forEach(row => {
    const content = row[contentIndex] || '';
    if (!content) return; // Пропускаем пустой контент
    
    const status = row[statusIndex] || '';
    const campaign = row[campaignIndex] || 'none';
    
    if (!contentStats[content]) {
      contentStats[content] = {
        content: content,
        totalLeads: 0,
        conversions: 0,
        conversionRate: 0,
        campaigns: {}
      };
    }
    
    const stats = contentStats[content];
    stats.totalLeads++;
    
    if (status === 'success') {
      stats.conversions++;
    }
    
    stats.campaigns[campaign] = (stats.campaigns[campaign] || 0) + 1;
  });
  
  return Object.values(contentStats).map(stats => {
    stats.conversionRate = stats.totalLeads > 0 ? (stats.conversions / stats.totalLeads * 100) : 0;
    
    // Топ кампании для контента
    stats.topCampaigns = Object.entries(stats.campaigns)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 2)
      .map(([campaign, count]) => ({ campaign, count }));
    
    return stats;
  }).filter(stats => stats.totalLeads >= 3) // Только контент с минимальным объёмом
    .sort((a, b) => b.conversions - a.conversions)
    .slice(0, 15); // Топ 15 видов контента
}

/**
 * Анализирует комбинации UTM параметров
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Анализ комбинаций
 * @private
 */
function analyzeUtmCombinations_(headers, rows) {
  const combinations = {};
  
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const mediumIndex = findHeaderIndex_(headers, 'UTM Medium');
  const campaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  
  rows.forEach(row => {
    const source = row[sourceIndex] || 'direct';
    const medium = row[mediumIndex] || 'none';
    const campaign = row[campaignIndex] || 'none';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    
    const comboKey = `${source} / ${medium} / ${campaign}`;
    
    if (!combinations[comboKey]) {
      combinations[comboKey] = {
        source: source,
        medium: medium,
        campaign: campaign,
        combination: comboKey,
        totalLeads: 0,
        conversions: 0,
        revenue: 0,
        conversionRate: 0,
        avgRevenue: 0
      };
    }
    
    const combo = combinations[comboKey];
    combo.totalLeads++;
    
    if (status === 'success') {
      combo.conversions++;
      combo.revenue += budget;
    }
  });
  
  return Object.values(combinations).map(combo => {
    combo.conversionRate = combo.totalLeads > 0 ? (combo.conversions / combo.totalLeads * 100) : 0;
    combo.avgRevenue = combo.conversions > 0 ? (combo.revenue / combo.conversions) : 0;
    return combo;
  }).filter(combo => combo.totalLeads >= 5) // Минимальный объём для значимости
    .sort((a, b) => b.revenue - a.revenue)
    .slice(0, 15); // Топ 15 комбинаций
}

/**
 * Анализирует временные паттерны UTM
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Временные паттерны
 * @private
 */
function analyzeUtmTimePatterns_(headers, rows) {
  const patterns = {
    sourcesByMonth: {},
    campaignActivity: {},
    seasonalTrends: {}
  };
  
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const campaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  
  rows.forEach(row => {
    const source = row[sourceIndex] || 'direct';
    const campaign = row[campaignIndex] || 'none';
    const createdDate = parseDate_(row[createdIndex]);
    
    if (!createdDate) return;
    
    const monthKey = formatDate_(createdDate, 'YYYY-MM');
    const monthName = `${getMonthName_(createdDate.getMonth())} ${createdDate.getFullYear()}`;
    
    // Источники по месяцам
    if (!patterns.sourcesByMonth[monthKey]) {
      patterns.sourcesByMonth[monthKey] = {
        month: monthName,
        sources: {}
      };
    }
    patterns.sourcesByMonth[monthKey].sources[source] = 
      (patterns.sourcesByMonth[monthKey].sources[source] || 0) + 1;
    
    // Активность кампаний
    if (!patterns.campaignActivity[campaign]) {
      patterns.campaignActivity[campaign] = {
        firstSeen: createdDate,
        lastSeen: createdDate,
        totalLeads: 0,
        activeMonths: new Set()
      };
    }
    
    const activity = patterns.campaignActivity[campaign];
    activity.totalLeads++;
    activity.activeMonths.add(monthKey);
    
    if (createdDate < activity.firstSeen) activity.firstSeen = createdDate;
    if (createdDate > activity.lastSeen) activity.lastSeen = createdDate;
    
    // Сезонные тренды
    const quarter = Math.floor(createdDate.getMonth() / 3) + 1;
    if (!patterns.seasonalTrends[`Q${quarter}`]) {
      patterns.seasonalTrends[`Q${quarter}`] = {
        totalLeads: 0,
        topSources: {}
      };
    }
    patterns.seasonalTrends[`Q${quarter}`].totalLeads++;
    patterns.seasonalTrends[`Q${quarter}`].topSources[source] = 
      (patterns.seasonalTrends[`Q${quarter}`].topSources[source] || 0) + 1;
  });
  
  // Преобразуем активность кампаний
  Object.keys(patterns.campaignActivity).forEach(campaign => {
    const activity = patterns.campaignActivity[campaign];
    activity.duration = Math.round((activity.lastSeen - activity.firstSeen) / (1000 * 60 * 60 * 24));
    activity.activeMonths = activity.activeMonths.size;
  });
  
  return patterns;
}

/**
 * Создаёт структуру отчёта UTM аналитики
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} utmData - Данные UTM анализа
 * @private
 */
function createUtmAnalysisStructure_(sheet, utmData) {
  let currentRow = 3;
  
  // 1. Общая статистика UTM
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('📊 ОБЩАЯ UTM СТАТИСТИКА');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const overallStats = [
    ['Всего записей с UTM:', utmData.totalRecords],
    ['Уникальных источников:', utmData.sources.length],
    ['Уникальных медиумов:', utmData.mediums.length],
    ['Активных кампаний:', utmData.campaigns.filter(c => c.isActive).length],
    ['Всего кампаний:', utmData.campaigns.length]
  ];
  
  sheet.getRange(currentRow, 1, overallStats.length, 2).setValues(overallStats);
  sheet.getRange(currentRow, 1, overallStats.length, 1).setFontWeight('bold');
  currentRow += overallStats.length + 2;
  
  // 2. Топ источники (UTM Source)
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.getRange(currentRow, 1).setValue('🌐 ТОП ИСТОЧНИКИ ТРАФИКА');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const sourceHeaders = [['Источник', 'Лиды', 'Конверсии', 'Выручка', 'Конверсия %', 'Средний чек']];
  sheet.getRange(currentRow, 1, 1, 6).setValues(sourceHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const sourceData = utmData.sources.slice(0, 10).map(item => [
    item.source,
    item.totalLeads,
    item.successfulDeals,
    formatCurrency_(item.totalRevenue),
    `${item.conversionRate.toFixed(1)}%`,
    formatCurrency_(item.avgRevenue)
  ]);
  
  if (sourceData.length > 0) {
    sheet.getRange(currentRow, 1, sourceData.length, 6).setValues(sourceData);
    currentRow += sourceData.length;
  }
  currentRow += 2;
  
  // 3. Топ медиумы (UTM Medium)
  sheet.getRange(currentRow, 1, 1, 5).merge();
  sheet.getRange(currentRow, 1).setValue('📺 ТОП МЕДИУМЫ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
  currentRow++;
  
  const mediumHeaders = [['Медиум', 'Лиды', 'Конверсии', 'Выручка', 'Конверсия %']];
  sheet.getRange(currentRow, 1, 1, 5).setValues(mediumHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
  currentRow++;
  
  const mediumData = utmData.mediums.slice(0, 8).map(item => [
    item.medium,
    item.totalLeads,
    item.successfulDeals,
    formatCurrency_(item.totalRevenue),
    `${item.conversionRate.toFixed(1)}%`
  ]);
  
  if (mediumData.length > 0) {
    sheet.getRange(currentRow, 1, mediumData.length, 5).setValues(mediumData);
    currentRow += mediumData.length;
  }
  currentRow += 2;
  
  // 4. Топ кампании
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.getRange(currentRow, 1).setValue('🎯 ТОП КАМПАНИИ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const campaignHeaders = [['Кампания', 'Лиды', 'Выручка', 'Конверсия %', 'Статус', 'Последняя активность']];
  sheet.getRange(currentRow, 1, 1, 6).setValues(campaignHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const campaignData = utmData.campaigns.slice(0, 12).map(item => [
    item.campaign,
    item.totalLeads,
    formatCurrency_(item.totalRevenue),
    `${item.conversionRate.toFixed(1)}%`,
    item.isActive ? '🟢 Активна' : '🔴 Неактивна',
    item.recentActivity ? formatDate_(item.recentActivity, 'DD.MM.YYYY') : 'Н/Д'
  ]);
  
  if (campaignData.length > 0) {
    sheet.getRange(currentRow, 1, campaignData.length, 6).setValues(campaignData);
    
    // Цветовая индикация статуса кампаний
    campaignData.forEach((row, index) => {
      const statusCell = sheet.getRange(currentRow + index, 5);
      if (row[4].includes('Активна')) {
        statusCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
      } else {
        statusCell.setBackground(CONFIG.COLORS.ERROR_LIGHT);
      }
    });
    
    currentRow += campaignData.length;
  }
  currentRow += 2;
  
  // 5. Топ комбинации UTM параметров
  if (utmData.combinations.length > 0) {
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue('🔗 ТОП UTM КОМБИНАЦИИ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
    currentRow++;
    
    const comboHeaders = [['Источник / Медиум / Кампания', 'Лиды', 'Конверсии', 'Выручка', 'ROI']];
    sheet.getRange(currentRow, 1, 1, 5).setValues(comboHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
    currentRow++;
    
    const comboData = utmData.combinations.slice(0, 8).map(item => [
      item.combination,
      item.totalLeads,
      item.conversions,
      formatCurrency_(item.revenue),
      `${item.conversionRate.toFixed(1)}%`
    ]);
    
    if (comboData.length > 0) {
      sheet.getRange(currentRow, 1, comboData.length, 5).setValues(comboData);
      currentRow += comboData.length;
    }
    currentRow += 2;
  }
  
  // 6. Анализ терминов (если есть)
  if (utmData.terms.length > 0) {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue('🔍 ТОП ПОИСКОВЫЕ ТЕРМИНЫ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const termHeaders = [['Термин', 'Лиды', 'Конверсии', 'Конверсия %']];
    sheet.getRange(currentRow, 1, 1, 4).setValues(termHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const termData = utmData.terms.slice(0, 10).map(item => [
      item.term,
      item.totalLeads,
      item.conversions,
      `${item.conversionRate.toFixed(1)}%`
    ]);
    
    if (termData.length > 0) {
      sheet.getRange(currentRow, 1, termData.length, 4).setValues(termData);
      currentRow += termData.length;
    }
    currentRow += 2;
  }
  
  // 7. Детализация топ источников
  const topSources = utmData.sources.slice(0, 3);
  topSources.forEach(sourceData => {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue(`🔍 ДЕТАЛИЗАЦИЯ: ${sourceData.source.toUpperCase()}`);
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const sourceDetails = [
      ['Показатели источника:', ''],
      [`• Всего лидов`, sourceData.totalLeads],
      [`• Успешных сделок`, sourceData.successfulDeals],
      [`• Общая выручка`, formatCurrency_(sourceData.totalRevenue)],
      [`• Конверсия`, `${sourceData.conversionRate.toFixed(1)}%`],
      [`• Средний чек`, formatCurrency_(sourceData.avgRevenue)]
    ];
    
    sheet.getRange(currentRow, 1, sourceDetails.length, 2).setValues(sourceDetails);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow += sourceDetails.length;
    
    // Топ каналы источника
    if (sourceData.topChannels.length > 0) {
      sheet.getRange(currentRow, 1).setValue('Топ каналы:');
      sheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
      
      sourceData.topChannels.forEach(channelInfo => {
        sheet.getRange(currentRow, 1).setValue(`• ${channelInfo.channel}`);
        sheet.getRange(currentRow, 2).setValue(`${channelInfo.count} лидов`);
        currentRow++;
      });
    }
    
    currentRow += 2;
  });
  
  // 8. Временной анализ активности кампаний
  if (Object.keys(utmData.timeAnalysis.campaignActivity).length > 0) {
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue('📅 ЖИЗНЕННЫЙ ЦИКЛ КАМПАНИЙ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
    currentRow++;
    
    const lifecycleHeaders = [['Кампания', 'Продолжительность', 'Активных месяцев', 'Лиды', 'Последняя активность']];
    sheet.getRange(currentRow, 1, 1, 5).setValues(lifecycleHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
    currentRow++;
    
    const lifecycleData = Object.entries(utmData.timeAnalysis.campaignActivity)
      .filter(([campaign]) => campaign !== 'none')
      .sort(([,a], [,b]) => b.totalLeads - a.totalLeads)
      .slice(0, 10)
      .map(([campaign, activity]) => [
        campaign,
        activity.duration > 0 ? `${activity.duration} дн.` : '1 день',
        activity.activeMonths,
        activity.totalLeads,
        formatDate_(activity.lastSeen, 'DD.MM.YYYY')
      ]);
    
    if (lifecycleData.length > 0) {
      sheet.getRange(currentRow, 1, lifecycleData.length, 5).setValues(lifecycleData);
      currentRow += lifecycleData.length;
    }
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 6);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 6);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к UTM аналитике
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} utmData - Данные UTM анализа
 * @private
 */
function addUtmAnalysisCharts_(sheet, utmData) {
  // 1. Круговая диаграмма источников трафика
  if (utmData.sources.length > 0) {
    const sourceChartData = [['Источник', 'Лиды']].concat(
      utmData.sources.slice(0, 8).map(item => [item.source, item.totalLeads])
    );
    
    // Создаем диаграмму через универсальную функцию
    const sourceChart = createChart_(sheet, 'pie', sourceChartData, {
      startRow: 1,
      startCol: 8,
      title: 'Распределение лидов по источникам',
      position: { row: 3, col: 8 },
      width: 500,
      height: 350,
      legend: 'right'
    });
  }
  
  // 2. Столбчатая диаграмма конверсии по медиумам
  if (utmData.mediums.length > 0) {
    const mediumChartData = [['Медиум', 'Конверсия %']].concat(
      utmData.mediums.slice(0, 8).map(item => [item.medium, item.conversionRate])
    );
    
    // Создаем диаграмму через универсальную функцию
    const mediumChart = createChart_(sheet, 'column', mediumChartData, {
      startRow: 1,
      startCol: 11,
      title: 'Конверсия по медиумам',
      position: { row: 3, col: 14 },
      width: 600,
      height: 350,
      legend: 'none',
      hAxisTitle: 'Медиум',
      vAxisTitle: 'Конверсия, %'
    });
  }
  
  // 3. Временная динамика топ источников
  if (Object.keys(utmData.timeAnalysis.sourcesByMonth).length > 0) {
    // Берём топ-5 источников
    const topSourceNames = utmData.sources.slice(0, 5).map(item => item.source);
    
    // Подготавливаем данные для линейной диаграммы
    const timeChartHeaders = ['Месяц'].concat(topSourceNames);
    const timeChartData = [timeChartHeaders];
    
    Object.entries(utmData.timeAnalysis.sourcesByMonth)
      .sort(([a], [b]) => a.localeCompare(b))
      .forEach(([monthKey, monthData]) => {
        const row = [monthData.month];
        
        topSourceNames.forEach(sourceName => {
          row.push(monthData.sources[sourceName] || 0);
        });
        
        timeChartData.push(row);
      });
    
    // Создаем диаграмму через универсальную функцию
    const timeChart = createChart_(sheet, 'line', timeChartData, {
      startRow: 1,
      startCol: 15,
      title: 'Динамика топ источников по месяцам',
      position: { row: 25, col: 8 },
      width: 700,
      height: 400,
      legend: 'top',
      hAxisTitle: 'Месяц',
      vAxisTitle: 'Количество лидов'
    });
  }
}

/**
 * Создаёт пустой отчёт когда нет данных
 * @private
 */
function createEmptyUtmReport_() {
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.UTM_ANALYSIS);
  clearSheetData_(sheet);
  applySheetFormatting_(sheet, 'UTM Аналитика');
  
  sheet.getRange('A3').setValue('ℹ️ Нет данных для UTM аналитики');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A5').setValue('Убедитесь, что:');
  sheet.getRange('A6').setValue('• В данных присутствуют UTM параметры');
  sheet.getRange('A7').setValue('• UTM метки корректно заполнены в источниках');
  sheet.getRange('A8').setValue('• Выполнена синхронизация данных');
  
  updateLastUpdateTime_(CONFIG.SHEETS.UTM_ANALYSIS);
}
