/**
 * МОДУЛЬ: АНАЛИЗ КАНАЛОВ
 * Подробный анализ всех маркетинговых каналов с воронками
 */

function runChannelAnalysis() {
  console.log('Начинаем анализ каналов...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const analysis = analyzeChannels(data);
    const sheet = createChannelAnalysisReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.CHANNEL_ANALYSIS}"`);
    console.log(`Проанализировано каналов: ${analysis.channels.length}`);
    
  } catch (error) {
    console.error('Ошибка при анализе каналов:', error);
    throw error;
  }
}

function analyzeChannels(data) {
  const channels = {};
  const crossChannelAnalysis = {};
  const conversionFunnels = {};
  const attributionAnalysis = {};
  const cohortAnalysis = {};
  
  // Группируем по клиентам для анализа пути
  const clientJourneys = {};
  
  data.forEach(row => {
    const phone = normalizePhone(row[CONFIG.WORKING_AMO_COLUMNS.PHONE] || '');
    const source = parseUtmSource(row);
    const channelType = getChannelType(source);
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const budget = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    const utmCampaign = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CAMPAIGN] || '';
    const utmContent = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CONTENT] || '';
    const formName = row[CONFIG.WORKING_AMO_COLUMNS.FORMNAME] || '';
    const cityTag = row[CONFIG.WORKING_AMO_COLUMNS.CITY_TAG] || '';
    
    const isSuccess = isSuccessStatus(status);
    const isRefusal = isRefusalStatus(status);
    
    // Основной анализ каналов
    if (!channels[channelType]) {
      channels[channelType] = {
        name: channelType,
        totalTouches: 0,
        uniqueUsers: new Set(),
        firstTouches: 0,
        lastTouches: 0,
        assistedConversions: 0,
        directConversions: 0,
        totalBudget: 0,
        totalRevenue: 0,
        avgOrderValue: 0,
        customerLtv: 0,
        sources: {},
        campaigns: {},
        managers: {},
        geoDistribution: {},
        timePatterns: {},
        qualityScore: 0,
        funnel: {
          impressions: 0,
          clicks: 0,
          visits: 0,
          leads: 0,
          qualified: 0,
          opportunities: 0,
          closedWon: 0
        }
      };
    }
    
    const channel = channels[channelType];
    channel.totalTouches++;
    channel.totalBudget += budget;
    
    if (phone && phone !== '') {
      channel.uniqueUsers.add(phone);
    }
    
    if (isSuccess) {
      channel.directConversions++;
      channel.totalRevenue += factAmount;
    }
    
    // Детализация источников
    if (!channel.sources[source]) channel.sources[source] = { count: 0, conversions: 0, revenue: 0 };
    channel.sources[source].count++;
    if (isSuccess) {
      channel.sources[source].conversions++;
      channel.sources[source].revenue += factAmount;
    }
    
    // Кампании
    if (utmCampaign) {
      if (!channel.campaigns[utmCampaign]) channel.campaigns[utmCampaign] = { count: 0, conversions: 0 };
      channel.campaigns[utmCampaign].count++;
      if (isSuccess) channel.campaigns[utmCampaign].conversions++;
    }
    
    // Менеджеры
    if (responsible) {
      if (!channel.managers[responsible]) channel.managers[responsible] = { count: 0, conversions: 0 };
      channel.managers[responsible].count++;
      if (isSuccess) channel.managers[responsible].conversions++;
    }
    
    // География
    if (cityTag) {
      if (!channel.geoDistribution[cityTag]) channel.geoDistribution[cityTag] = { count: 0, conversions: 0 };
      channel.geoDistribution[cityTag].count++;
      if (isSuccess) channel.geoDistribution[cityTag].conversions++;
    }
    
    // Воронка
    channel.funnel.visits++;
    if (formName) channel.funnel.leads++;
    if (!isRefusal && status !== '') channel.funnel.qualified++;
    if (isSuccess) channel.funnel.closedWon++;
    
    // Временные паттерны
    if (createdAt) {
      const hour = new Date(createdAt).getHours();
      if (!channel.timePatterns[hour]) channel.timePatterns[hour] = { count: 0, conversions: 0 };
      channel.timePatterns[hour].count++;
      if (isSuccess) channel.timePatterns[hour].conversions++;
    }
    
    // Анализ пути клиента
    if (phone && phone !== '') {
      if (!clientJourneys[phone]) {
        clientJourneys[phone] = {
          phone: phone,
          touchPoints: [],
          firstChannel: channelType,
          lastChannel: channelType,
          finalStatus: status,
          totalValue: 0,
          journeyLength: 0
        };
      }
      
      const journey = clientJourneys[phone];
      journey.touchPoints.push({
        channel: channelType,
        source: source,
        timestamp: createdAt,
        campaign: utmCampaign,
        manager: responsible
      });
      
      journey.lastChannel = channelType;
      journey.finalStatus = status;
      if (isSuccess) journey.totalValue += factAmount;
      journey.journeyLength = journey.touchPoints.length;
    }
  });
  
  // Постобработка каналов
  Object.values(channels).forEach(channel => {
    channel.uniqueUsers = channel.uniqueUsers.size;
    channel.conversionRate = channel.totalTouches > 0 ? (channel.directConversions / channel.totalTouches * 100) : 0;
    channel.avgOrderValue = channel.directConversions > 0 ? channel.totalRevenue / channel.directConversions : 0;
    channel.costPerLead = channel.totalTouches > 0 ? channel.totalBudget / channel.totalTouches : 0;
    channel.costPerAcquisition = channel.directConversions > 0 ? channel.totalBudget / channel.directConversions : 0;
    channel.romi = channel.totalBudget > 0 ? ((channel.totalRevenue - channel.totalBudget) / channel.totalBudget * 100) : 0;
    
    // Качественная оценка канала
    channel.qualityScore = calculateChannelQuality(channel);
    
    // Топ источники, кампании, менеджеры
    channel.topSource = getTopPerformer(channel.sources);
    channel.topCampaign = getTopPerformer(channel.campaigns);
    channel.topManager = getTopPerformer(channel.managers);
    channel.topGeo = getTopPerformer(channel.geoDistribution);
    
    // Лучшее время
    channel.peakHour = getTopPerformer(channel.timePatterns);
    
    // Конверсия воронки
    channel.funnelConversion = {
      visitToLead: channel.funnel.visits > 0 ? (channel.funnel.leads / channel.funnel.visits * 100) : 0,
      leadToQualified: channel.funnel.leads > 0 ? (channel.funnel.qualified / channel.funnel.leads * 100) : 0,
      qualifiedToWon: channel.funnel.qualified > 0 ? (channel.funnel.closedWon / channel.funnel.qualified * 100) : 0
    };
  });
  
  // Анализ пересечений каналов
  const channelIntersections = analyzeChannelIntersections(Object.values(clientJourneys));
  
  // Атрибуция
  const attributionResults = calculateAttribution(Object.values(clientJourneys));
  
  return {
    channels: Object.values(channels).sort((a, b) => b.qualityScore - a.qualityScore),
    clientJourneys: Object.values(clientJourneys).sort((a, b) => b.journeyLength - a.journeyLength),
    channelIntersections,
    attribution: attributionResults,
    summary: {
      totalChannels: Object.keys(channels).length,
      totalTouches: Object.values(channels).reduce((sum, ch) => sum + ch.totalTouches, 0),
      totalRevenue: Object.values(channels).reduce((sum, ch) => sum + ch.totalRevenue, 0),
      bestChannel: Object.values(channels).sort((a, b) => b.qualityScore - a.qualityScore)[0]?.name || 'Не определен'
    }
  };
}

function calculateChannelQuality(channel) {
  // Комплексная оценка качества канала (0-100)
  const conversionWeight = 0.3;
  const roiWeight = 0.25;
  const volumeWeight = 0.2;
  const costEfficiencyWeight = 0.15;
  const stabilityWeight = 0.1;
  
  const conversionScore = Math.min(channel.conversionRate / 50 * 100, 100); // 50% = идеальная конверсия
  const roiScore = channel.romi >= 0 ? Math.min(channel.romi / 200 * 100, 100) : 0; // 200% = отличный ROMI
  const volumeScore = Math.min(channel.totalTouches / 100 * 100, 100); // 100 касаний = хороший объем
  const costScore = channel.costPerLead > 0 ? Math.max(0, 100 - (channel.costPerLead / 1000 * 100)) : 0; // 1000р = дорого
  const stabilityScore = 70; // Упрощенно, можно рассчитать через дисперсию
  
  return (conversionScore * conversionWeight +
          roiScore * roiWeight +
          volumeScore * volumeWeight +
          costScore * costEfficiencyWeight +
          stabilityScore * stabilityWeight);
}

function getTopPerformer(items) {
  if (!items || Object.keys(items).length === 0) return 'Не определен';
  
  const sorted = Object.entries(items)
    .sort((a, b) => {
      // Сортируем по конверсиям, если есть, иначе по количеству
      const aValue = a[1].conversions !== undefined ? a[1].conversions : a[1].count || a[1];
      const bValue = b[1].conversions !== undefined ? b[1].conversions : b[1].count || b[1];
      return bValue - aValue;
    });
  
  if (sorted.length === 0) return 'Не определен';
  
  const item = sorted[0][1];
  const count = item.count || item.conversions || item;
  return `${sorted[0][0]} (${count})`;
}

function analyzeChannelIntersections(journeys) {
  const intersections = {};
  
  journeys.forEach(journey => {
    if (journey.touchPoints.length > 1) {
      const channels = [...new Set(journey.touchPoints.map(tp => tp.channel))];
      if (channels.length > 1) {
        const key = channels.sort().join(' + ');
        if (!intersections[key]) {
          intersections[key] = {
            combination: key,
            customers: 0,
            totalValue: 0,
            avgJourneyLength: 0
          };
        }
        
        intersections[key].customers++;
        intersections[key].totalValue += journey.totalValue;
        intersections[key].avgJourneyLength += journey.journeyLength;
      }
    }
  });
  
  Object.values(intersections).forEach(intersection => {
    intersection.avgJourneyLength = intersection.customers > 0 ? 
      intersection.avgJourneyLength / intersection.customers : 0;
    intersection.avgValue = intersection.customers > 0 ? 
      intersection.totalValue / intersection.customers : 0;
  });
  
  return Object.values(intersections)
    .sort((a, b) => b.customers - a.customers)
    .slice(0, 10);
}

function calculateAttribution(journeys) {
  const models = {
    firstTouch: {},
    lastTouch: {},
    linear: {},
    timeDecay: {}
  };
  
  journeys.forEach(journey => {
    if (journey.totalValue > 0 && journey.touchPoints.length > 0) {
      const channels = journey.touchPoints.map(tp => tp.channel);
      const value = journey.totalValue;
      
      // First Touch
      const firstChannel = channels[0];
      models.firstTouch[firstChannel] = (models.firstTouch[firstChannel] || 0) + value;
      
      // Last Touch
      const lastChannel = channels[channels.length - 1];
      models.lastTouch[lastChannel] = (models.lastTouch[lastChannel] || 0) + value;
      
      // Linear
      const linearValue = value / channels.length;
      channels.forEach(channel => {
        models.linear[channel] = (models.linear[channel] || 0) + linearValue;
      });
      
      // Time Decay (больше веса последним касаниям)
      channels.forEach((channel, index) => {
        const weight = Math.pow(2, index) / (Math.pow(2, channels.length) - 1);
        const decayValue = value * weight;
        models.timeDecay[channel] = (models.timeDecay[channel] || 0) + decayValue;
      });
    }
  });
  
  return {
    firstTouch: Object.entries(models.firstTouch)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([channel, value]) => ({ channel, value, percentage: 0 })),
    
    lastTouch: Object.entries(models.lastTouch)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([channel, value]) => ({ channel, value, percentage: 0 })),
    
    linear: Object.entries(models.linear)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([channel, value]) => ({ channel, value, percentage: 0 })),
    
    timeDecay: Object.entries(models.timeDecay)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([channel, value]) => ({ channel, value, percentage: 0 }))
  };
}

function createChannelAnalysisReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.CHANNEL_ANALYSIS);
  
  let currentRow = 1;
  
  // Общая сводка
  currentRow = addChannelSummary(sheet, analysis.summary, currentRow);
  currentRow += 2;
  
  // Рейтинг каналов
  currentRow = addChannelRanking(sheet, analysis.channels, currentRow);
  currentRow += 2;
  
  // Анализ пересечений
  currentRow = addChannelIntersections(sheet, analysis.channelIntersections, currentRow);
  currentRow += 2;
  
  // Модели атрибуции
  currentRow = addAttributionModels(sheet, analysis.attribution, currentRow);
  currentRow += 2;
  
  // Топ клиентские пути
  addTopJourneys(sheet, analysis.clientJourneys, currentRow);
  
  return sheet;
}

function addChannelSummary(sheet, summary, startRow) {
  const headers = ['Сводка по каналам', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1565c0')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const summaryData = [
    ['Всего каналов', summary.totalChannels],
    ['Общее количество касаний', summary.totalTouches],
    ['Общая выручка', formatCurrency(summary.totalRevenue)],
    ['Лучший канал', summary.bestChannel],
    ['Среднее касаний на канал', summary.totalChannels > 0 ? Math.round(summary.totalTouches / summary.totalChannels) : 0]
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, summaryData.length, 2);
  dataRange.setValues(summaryData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + summaryData.length + 1;
}

function addChannelRanking(sheet, channels, startRow) {
  const headers = [
    'Канал',
    'Оценка качества',
    'Касания',
    'Уникальные клиенты',
    'Конверсия %',
    'CPL',
    'CPA',
    'ROMI %',
    'Средний чек',
    'Топ источник',
    'Пиковый час'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#2e7d32')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const channelData = channels.slice(0, 15).map(channel => [
    channel.name,
    channel.qualityScore.toFixed(1),
    channel.totalTouches,
    channel.uniqueUsers,
    channel.conversionRate.toFixed(1) + '%',
    formatCurrency(channel.costPerLead),
    formatCurrency(channel.costPerAcquisition),
    channel.romi.toFixed(1) + '%',
    formatCurrency(channel.avgOrderValue),
    channel.topSource,
    channel.peakHour
  ]);
  
  if (channelData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, channelData.length, headers.length);
    dataRange.setValues(channelData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    
    // Условное форматирование для оценки качества
    const qualityRange = sheet.getRange(startRow + 1, 2, channelData.length, 1);
    const qualityRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#4caf50', SpreadsheetApp.InterpolationType.NUMBER, '80')
      .setGradientMidpointWithValue('#ff9800', SpreadsheetApp.InterpolationType.NUMBER, '50')
      .setGradientMinpointWithValue('#f44336', SpreadsheetApp.InterpolationType.NUMBER, '20')
      .setRanges([qualityRange])
      .build();
    
    sheet.setConditionalFormatRules([qualityRule]);
  }
  
  return startRow + Math.max(channelData.length, 1) + 1;
}

function addChannelIntersections(sheet, intersections, startRow) {
  const headers = [
    'Комбинация каналов',
    'Клиентов',
    'Средняя длина пути',
    'Общая ценность',
    'Средняя ценность'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#7b1fa2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const intersectionData = intersections.map(intersection => [
    intersection.combination,
    intersection.customers,
    intersection.avgJourneyLength.toFixed(1),
    formatCurrency(intersection.totalValue),
    formatCurrency(intersection.avgValue)
  ]);
  
  if (intersectionData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, intersectionData.length, headers.length);
    dataRange.setValues(intersectionData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(intersectionData.length, 1) + 1;
}

function addAttributionModels(sheet, attribution, startRow) {
  const models = [
    { name: 'First Touch Attribution', data: attribution.firstTouch },
    { name: 'Last Touch Attribution', data: attribution.lastTouch },
    { name: 'Linear Attribution', data: attribution.linear },
    { name: 'Time Decay Attribution', data: attribution.timeDecay }
  ];
  
  models.forEach(model => {
    const headers = [`${model.name}`, 'Канал', 'Ценность'];
    const headerRange = sheet.getRange(startRow, 1, 1, 3);
    headerRange.setValues([['', ...headers.slice(1)]]);
    headerRange.setBackground('#f57c00')
              .setFontColor('#ffffff')
              .setFontWeight('bold')
              .setFontFamily(CONFIG.DEFAULT_FONT);
    
    const modelData = model.data.slice(0, 5).map((item, index) => [
      index === 0 ? model.name : '',
      item.channel,
      formatCurrency(item.value)
    ]);
    
    if (modelData.length > 0) {
      const dataRange = sheet.getRange(startRow + 1, 1, modelData.length, 3);
      dataRange.setValues(modelData);
      dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    }
    
    startRow += Math.max(modelData.length, 1) + 2;
  });
  
  return startRow;
}

function addTopJourneys(sheet, journeys, startRow) {
  const headers = [
    'Телефон',
    'Длина пути',
    'Первый канал',
    'Последний канал',
    'Итоговая ценность',
    'Статус'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d32f2f')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const journeyData = journeys.slice(0, 20).map(journey => [
    journey.phone,
    journey.journeyLength,
    journey.firstChannel,
    journey.lastChannel,
    formatCurrency(journey.totalValue),
    journey.finalStatus
  ]);
  
  if (journeyData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, journeyData.length, headers.length);
    dataRange.setValues(journeyData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, Math.max(headers.length, 11));
  
  return startRow + journeyData.length + 1;
}
