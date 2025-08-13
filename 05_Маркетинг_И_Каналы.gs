/**
 * МОДУЛЬ: МАРКЕТИНГ И КАНАЛЫ
 * Детальный анализ маркетинговых каналов с метриками эффективности
 */

function runMarketingChannelsAnalysis() {
  console.log('Начинаем анализ маркетинговых каналов...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const analysis = analyzeMarketingChannels(data);
    const sheet = createMarketingChannelsReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.MARKETING_CHANNELS}"`);
    console.log(`Проанализировано каналов: ${Object.keys(analysis.channels).length}`);
    
  } catch (error) {
    console.error('Ошибка при анализе маркетинговых каналов:', error);
    throw error;
  }
}

function analyzeMarketingChannels(data) {
  const channels = {};
  const monthlyData = {};
  const campaignData = {};
  
  let totalLeads = 0;
  let totalBudget = 0;
  let totalRevenue = 0;
  
  data.forEach(row => {
    const source = parseUtmSource(row);
    const channelType = getChannelType(source);
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const budget = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const utmCampaign = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CAMPAIGN] || '';
    const utmContent = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CONTENT] || '';
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    
    const isSuccess = isSuccessStatus(status);
    const isRefusal = isRefusalStatus(status);
    
    totalLeads++;
    totalBudget += budget;
    if (isSuccess) totalRevenue += factAmount;
    
    // Анализ каналов
    if (!channels[channelType]) {
      channels[channelType] = {
        type: channelType,
        sources: {},
        totalLeads: 0,
        successLeads: 0,
        refusalLeads: 0,
        totalBudget: 0,
        totalRevenue: 0,
        avgTicket: 0,
        cpl: 0, // Cost Per Lead
        cpa: 0, // Cost Per Acquisition
        romi: 0,
        conversionRate: 0,
        refusalRate: 0
      };
    }
    
    const channel = channels[channelType];
    channel.totalLeads++;
    channel.totalBudget += budget;
    if (isSuccess) {
      channel.successLeads++;
      channel.totalRevenue += factAmount;
    }
    if (isRefusal) channel.refusalLeads++;
    
    // Источники внутри канала
    if (!channel.sources[source]) {
      channel.sources[source] = {
        name: source,
        leads: 0,
        success: 0,
        budget: 0,
        revenue: 0
      };
    }
    
    channel.sources[source].leads++;
    channel.sources[source].budget += budget;
    if (isSuccess) {
      channel.sources[source].success++;
      channel.sources[source].revenue += factAmount;
    }
    
    // Анализ по месяцам
    if (createdAt) {
      const periods = getDatePeriods(createdAt);
      if (periods.yearMonth) {
        if (!monthlyData[periods.yearMonth]) {
          monthlyData[periods.yearMonth] = {};
        }
        if (!monthlyData[periods.yearMonth][channelType]) {
          monthlyData[periods.yearMonth][channelType] = {
            leads: 0,
            success: 0,
            budget: 0,
            revenue: 0
          };
        }
        
        const monthChannel = monthlyData[periods.yearMonth][channelType];
        monthChannel.leads++;
        monthChannel.budget += budget;
        if (isSuccess) {
          monthChannel.success++;
          monthChannel.revenue += factAmount;
        }
      }
    }
    
    // Анализ кампаний
    if (utmCampaign && utmCampaign !== '') {
      const campaignKey = `${channelType} | ${utmCampaign}`;
      if (!campaignData[campaignKey]) {
        campaignData[campaignKey] = {
          channel: channelType,
          campaign: utmCampaign,
          content: utmContent,
          leads: 0,
          success: 0,
          budget: 0,
          revenue: 0,
          managers: new Set()
        };
      }
      
      const campaign = campaignData[campaignKey];
      campaign.leads++;
      campaign.budget += budget;
      if (isSuccess) {
        campaign.success++;
        campaign.revenue += factAmount;
      }
      if (responsible) campaign.managers.add(responsible);
    }
  });
  
  // Вычисляем метрики для каналов
  Object.values(channels).forEach(channel => {
    channel.conversionRate = channel.totalLeads > 0 ? 
      (channel.successLeads / channel.totalLeads * 100) : 0;
    
    channel.refusalRate = channel.totalLeads > 0 ? 
      (channel.refusalLeads / channel.totalLeads * 100) : 0;
    
    channel.avgTicket = channel.successLeads > 0 ? 
      (channel.totalRevenue / channel.successLeads) : 0;
    
    channel.cpl = channel.totalLeads > 0 ? 
      (channel.totalBudget / channel.totalLeads) : 0;
    
    channel.cpa = channel.successLeads > 0 ? 
      (channel.totalBudget / channel.successLeads) : 0;
    
    channel.romi = channel.totalBudget > 0 ? 
      ((channel.totalRevenue - channel.totalBudget) / channel.totalBudget * 100) : 0;
  });
  
  // Конвертируем кампании
  const campaigns = Object.values(campaignData).map(campaign => ({
    ...campaign,
    managers: Array.from(campaign.managers).join(', '),
    conversionRate: campaign.leads > 0 ? (campaign.success / campaign.leads * 100) : 0,
    cpl: campaign.leads > 0 ? (campaign.budget / campaign.leads) : 0,
    cpa: campaign.success > 0 ? (campaign.budget / campaign.success) : 0,
    romi: campaign.budget > 0 ? ((campaign.revenue - campaign.budget) / campaign.budget * 100) : 0
  }));
  
  return {
    channels: Object.values(channels).sort((a, b) => b.totalLeads - a.totalLeads),
    monthlyData,
    campaigns: campaigns.sort((a, b) => b.leads - a.leads),
    totals: {
      leads: totalLeads,
      budget: totalBudget,
      revenue: totalRevenue,
      romi: totalBudget > 0 ? ((totalRevenue - totalBudget) / totalBudget * 100) : 0
    }
  };
}

function createMarketingChannelsReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.MARKETING_CHANNELS);
  
  let currentRow = 1;
  
  // Общие метрики
  currentRow = addOverallMetrics(sheet, analysis.totals, currentRow);
  currentRow += 2;
  
  // Анализ каналов
  currentRow = addChannelsAnalysis(sheet, analysis.channels, currentRow);
  currentRow += 2;
  
  // Топ кампании
  currentRow = addTopCampaigns(sheet, analysis.campaigns, currentRow);
  currentRow += 2;
  
  // Помесячная динамика
  addMonthlyDynamics(sheet, analysis.monthlyData, currentRow);
  
  return sheet;
}

function addOverallMetrics(sheet, totals, startRow) {
  const headers = ['Общие метрики', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1e3d59')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const metricsData = [
    ['Общее количество лидов', totals.leads],
    ['Общий маркетинговый бюджет', formatCurrency(totals.budget)],
    ['Общая выручка', formatCurrency(totals.revenue)],
    ['Общий ROMI', totals.romi.toFixed(1) + '%'],
    ['Средний CPL (по всем каналам)', totals.leads > 0 ? formatCurrency(totals.budget / totals.leads) : '0 ₽']
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, metricsData.length, 2);
  dataRange.setValues(metricsData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + metricsData.length + 1;
}

function addChannelsAnalysis(sheet, channels, startRow) {
  const headers = [
    'Тип канала',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'Отказы %',
    'Бюджет',
    'Выручка',
    'Средний чек',
    'CPL',
    'CPA',
    'ROMI %'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#0d7377')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const channelData = channels.map(channel => [
    channel.type,
    channel.totalLeads,
    channel.successLeads,
    channel.conversionRate.toFixed(1) + '%',
    channel.refusalRate.toFixed(1) + '%',
    formatCurrency(channel.totalBudget),
    formatCurrency(channel.totalRevenue),
    formatCurrency(channel.avgTicket),
    formatCurrency(channel.cpl),
    formatCurrency(channel.cpa),
    channel.romi.toFixed(1) + '%'
  ]);
  
  if (channelData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, channelData.length, headers.length);
    dataRange.setValues(channelData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    
    // Условное форматирование для ROMI
    const romiRange = sheet.getRange(startRow + 1, 11, channelData.length, 1);
    const romiRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#34a853', SpreadsheetApp.InterpolationType.NUMBER, '100')
      .setGradientMidpointWithValue('#fbbc04', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMinpointWithValue('#ea4335', SpreadsheetApp.InterpolationType.NUMBER, '-50')
      .setRanges([romiRange])
      .build();
    
    sheet.setConditionalFormatRules([romiRule]);
  }
  
  return startRow + Math.max(channelData.length, 1) + 1;
}

function addTopCampaigns(sheet, campaigns, startRow) {
  const headers = [
    'Канал',
    'Кампания',
    'Контент',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'CPL',
    'CPA',
    'ROMI %',
    'Менеджеры'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#14213d')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  // Топ-20 кампаний
  const topCampaigns = campaigns.slice(0, 20);
  const campaignData = topCampaigns.map(campaign => [
    campaign.channel,
    campaign.campaign,
    campaign.content || 'Не указан',
    campaign.leads,
    campaign.success,
    campaign.conversionRate.toFixed(1) + '%',
    formatCurrency(campaign.cpl),
    formatCurrency(campaign.cpa),
    campaign.romi.toFixed(1) + '%',
    campaign.managers || 'Не указан'
  ]);
  
  if (campaignData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, campaignData.length, headers.length);
    dataRange.setValues(campaignData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(campaignData.length, 1) + 1;
}

function addMonthlyDynamics(sheet, monthlyData, startRow) {
  const months = Object.keys(monthlyData).sort();
  if (months.length === 0) return startRow;
  
  const channelTypes = new Set();
  months.forEach(month => {
    Object.keys(monthlyData[month]).forEach(channel => {
      channelTypes.add(channel);
    });
  });
  
  const headers = ['Месяц', ...Array.from(channelTypes), 'Итого'];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#2c5f41')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const monthlyRows = months.map(month => {
    const row = [month];
    let monthTotal = 0;
    
    Array.from(channelTypes).forEach(channel => {
      const channelData = monthlyData[month][channel];
      const leads = channelData ? channelData.leads : 0;
      row.push(leads);
      monthTotal += leads;
    });
    
    row.push(monthTotal);
    return row;
  });
  
  if (monthlyRows.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, monthlyRows.length, headers.length);
    dataRange.setValues(monthlyRows);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, headers.length);
  
  return startRow + monthlyRows.length + 1;
}
