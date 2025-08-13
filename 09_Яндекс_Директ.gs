/**
 * МОДУЛЬ: ЯНДЕКС.ДИРЕКТ АНАЛИЗ
 * Специализированный анализ кампаний Яндекс.Директ
 */

function runYandexDirectAnalysis() {
  console.log('Начинаем анализ Яндекс.Директ...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const analysis = analyzeYandexDirect(data);
    const sheet = createYandexDirectReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.YA_DIRECT}"`);
    console.log(`Проанализировано кампаний: ${Object.keys(analysis.campaigns).length}`);
    
  } catch (error) {
    console.error('Ошибка при анализе Яндекс.Директ:', error);
    throw error;
  }
}

function analyzeYandexDirect(data) {
  const campaigns = {};
  const keywords = {};
  const adGroups = {};
  const deviceAnalysis = {};
  const geoAnalysis = {};
  
  let totalDirectLeads = 0;
  let totalDirectBudget = 0;
  let totalDirectRevenue = 0;
  
  data.forEach(row => {
    const source = row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || '';
    const medium = row[CONFIG.WORKING_AMO_COLUMNS.UTM_MEDIUM] || '';
    const campaign = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CAMPAIGN] || '';
    const content = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CONTENT] || '';
    const term = row[CONFIG.WORKING_AMO_COLUMNS.UTM_TERM] || '';
    
    // Фильтруем только Яндекс.Директ трафик
    if (!isYandexDirectTraffic(source, medium)) {
      return;
    }
    
    totalDirectLeads++;
    
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const budget = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const cityTag = row[CONFIG.WORKING_AMO_COLUMNS.CITY_TAG] || '';
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    
    const isSuccess = isSuccessStatus(status);
    totalDirectBudget += budget;
    if (isSuccess) totalDirectRevenue += factAmount;
    
    // Анализ кампаний
    if (campaign && campaign !== '') {
      if (!campaigns[campaign]) {
        campaigns[campaign] = {
          name: campaign,
          leads: 0,
          success: 0,
          budget: 0,
          revenue: 0,
          cpl: 0,
          cpa: 0,
          romi: 0,
          keywords: new Set(),
          adGroups: new Set(),
          cities: new Set(),
          managers: new Set()
        };
      }
      
      const camp = campaigns[campaign];
      camp.leads++;
      camp.budget += budget;
      if (isSuccess) {
        camp.success++;
        camp.revenue += factAmount;
      }
      if (term) camp.keywords.add(term);
      if (content) camp.adGroups.add(content);
      if (cityTag) camp.cities.add(cityTag);
      if (responsible) camp.managers.add(responsible);
    }
    
    // Анализ ключевых слов
    if (term && term !== '') {
      if (!keywords[term]) {
        keywords[term] = {
          keyword: term,
          leads: 0,
          success: 0,
          budget: 0,
          revenue: 0,
          campaigns: new Set(),
          avgPosition: 0,
          competitiveness: 'Средняя'
        };
      }
      
      keywords[term].leads++;
      keywords[term].budget += budget;
      if (isSuccess) {
        keywords[term].success++;
        keywords[term].revenue += factAmount;
      }
      if (campaign) keywords[term].campaigns.add(campaign);
    }
    
    // Анализ групп объявлений
    if (content && content !== '') {
      if (!adGroups[content]) {
        adGroups[content] = {
          group: content,
          leads: 0,
          success: 0,
          budget: 0,
          revenue: 0,
          campaign: campaign,
          keywords: new Set()
        };
      }
      
      adGroups[content].leads++;
      adGroups[content].budget += budget;
      if (isSuccess) {
        adGroups[content].success++;
        adGroups[content].revenue += factAmount;
      }
      if (term) adGroups[content].keywords.add(term);
    }
    
    // Географический анализ
    if (cityTag && cityTag !== '') {
      if (!geoAnalysis[cityTag]) {
        geoAnalysis[cityTag] = {
          city: cityTag,
          leads: 0,
          success: 0,
          budget: 0,
          revenue: 0,
          topCampaigns: {}
        };
      }
      
      geoAnalysis[cityTag].leads++;
      geoAnalysis[cityTag].budget += budget;
      if (isSuccess) {
        geoAnalysis[cityTag].success++;
        geoAnalysis[cityTag].revenue += factAmount;
      }
      
      if (campaign) {
        if (!geoAnalysis[cityTag].topCampaigns[campaign]) {
          geoAnalysis[cityTag].topCampaigns[campaign] = 0;
        }
        geoAnalysis[cityTag].topCampaigns[campaign]++;
      }
    }
  });
  
  // Вычисляем метрики
  Object.values(campaigns).forEach(campaign => {
    campaign.conversionRate = campaign.leads > 0 ? (campaign.success / campaign.leads * 100) : 0;
    campaign.cpl = campaign.leads > 0 ? (campaign.budget / campaign.leads) : 0;
    campaign.cpa = campaign.success > 0 ? (campaign.budget / campaign.success) : 0;
    campaign.romi = campaign.budget > 0 ? ((campaign.revenue - campaign.budget) / campaign.budget * 100) : 0;
    campaign.keywords = Array.from(campaign.keywords);
    campaign.adGroups = Array.from(campaign.adGroups);
    campaign.cities = Array.from(campaign.cities);
    campaign.managers = Array.from(campaign.managers);
  });
  
  Object.values(keywords).forEach(keyword => {
    keyword.conversionRate = keyword.leads > 0 ? (keyword.success / keyword.leads * 100) : 0;
    keyword.cpl = keyword.leads > 0 ? (keyword.budget / keyword.leads) : 0;
    keyword.romi = keyword.budget > 0 ? ((keyword.revenue - keyword.budget) / keyword.budget * 100) : 0;
    keyword.campaigns = Array.from(keyword.campaigns);
    keyword.competitiveness = determineKeywordCompetitiveness(keyword);
  });
  
  Object.values(adGroups).forEach(group => {
    group.conversionRate = group.leads > 0 ? (group.success / group.leads * 100) : 0;
    group.cpl = group.leads > 0 ? (group.budget / group.leads) : 0;
    group.romi = group.budget > 0 ? ((group.revenue - group.budget) / group.budget * 100) : 0;
    group.keywords = Array.from(group.keywords);
  });
  
  Object.values(geoAnalysis).forEach(geo => {
    geo.conversionRate = geo.leads > 0 ? (geo.success / geo.leads * 100) : 0;
    geo.cpl = geo.leads > 0 ? (geo.budget / geo.leads) : 0;
    geo.romi = geo.budget > 0 ? ((geo.revenue - geo.budget) / geo.budget * 100) : 0;
    geo.topCampaign = getTopEntry(geo.topCampaigns);
  });
  
  return {
    totalDirectLeads,
    totalDirectBudget,
    totalDirectRevenue,
    overallROMI: totalDirectBudget > 0 ? ((totalDirectRevenue - totalDirectBudget) / totalDirectBudget * 100) : 0,
    overallConversion: totalDirectLeads > 0 ? (Object.values(campaigns).reduce((sum, c) => sum + c.success, 0) / totalDirectLeads * 100) : 0,
    campaigns: Object.values(campaigns).sort((a, b) => b.leads - a.leads),
    keywords: Object.values(keywords).sort((a, b) => b.leads - a.leads),
    adGroups: Object.values(adGroups).sort((a, b) => b.leads - a.leads),
    geoAnalysis: Object.values(geoAnalysis).sort((a, b) => b.leads - a.leads)
  };
}

function isYandexDirectTraffic(source, medium) {
  const sourceLower = source.toLowerCase();
  const mediumLower = medium.toLowerCase();
  
  return (sourceLower.includes('yandex') || sourceLower.includes('direct')) &&
         (mediumLower.includes('cpc') || mediumLower.includes('ppc') || mediumLower === '');
}

function determineKeywordCompetitiveness(keyword) {
  // Простая логика определения конкурентности
  if (keyword.cpl > 500) return 'Высокая';
  if (keyword.cpl > 200) return 'Средняя';
  if (keyword.cpl > 50) return 'Низкая';
  return 'Очень низкая';
}

function createYandexDirectReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.YA_DIRECT);
  
  let currentRow = 1;
  
  // Общие метрики
  currentRow = addDirectOverview(sheet, analysis, currentRow);
  currentRow += 2;
  
  // Анализ кампаний
  currentRow = addCampaignsAnalysis(sheet, analysis.campaigns, currentRow);
  currentRow += 2;
  
  // Топ ключевые слова
  currentRow = addKeywordsAnalysis(sheet, analysis.keywords, currentRow);
  currentRow += 2;
  
  // Группы объявлений
  currentRow = addAdGroupsAnalysis(sheet, analysis.adGroups, currentRow);
  currentRow += 2;
  
  // Географический анализ
  addGeoDirectAnalysis(sheet, analysis.geoAnalysis, currentRow);
  
  return sheet;
}

function addDirectOverview(sheet, analysis, startRow) {
  const headers = ['Метрика Яндекс.Директ', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#ffcc02')
            .setFontColor('#000000')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const overviewData = [
    ['Всего лидов из Директа', analysis.totalDirectLeads],
    ['Общий бюджет', formatCurrency(analysis.totalDirectBudget)],
    ['Общая выручка', formatCurrency(analysis.totalDirectRevenue)],
    ['Общий ROMI', analysis.overallROMI.toFixed(1) + '%'],
    ['Общая конверсия', analysis.overallConversion.toFixed(1) + '%'],
    ['Активных кампаний', analysis.campaigns.length],
    ['Уникальных ключевых слов', analysis.keywords.length],
    ['Групп объявлений', analysis.adGroups.length],
    ['Средний CPL', analysis.totalDirectLeads > 0 ? formatCurrency(analysis.totalDirectBudget / analysis.totalDirectLeads) : '0 ₽']
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, overviewData.length, 2);
  dataRange.setValues(overviewData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + overviewData.length + 1;
}

function addCampaignsAnalysis(sheet, campaigns, startRow) {
  const headers = [
    'Кампания',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'Бюджет',
    'Выручка',
    'CPL',
    'CPA',
    'ROMI %',
    'Ключевых слов',
    'Групп объявлений'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d50000')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const campaignData = campaigns.map(campaign => [
    campaign.name,
    campaign.leads,
    campaign.success,
    campaign.conversionRate.toFixed(1) + '%',
    formatCurrency(campaign.budget),
    formatCurrency(campaign.revenue),
    formatCurrency(campaign.cpl),
    formatCurrency(campaign.cpa),
    campaign.romi.toFixed(1) + '%',
    campaign.keywords.length,
    campaign.adGroups.length
  ]);
  
  if (campaignData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, campaignData.length, headers.length);
    dataRange.setValues(campaignData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    
    // Условное форматирование для ROMI
    const romiRange = sheet.getRange(startRow + 1, 9, campaignData.length, 1);
    const romiRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#4caf50', SpreadsheetApp.InterpolationType.NUMBER, '200')
      .setGradientMidpointWithValue('#ff9800', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMinpointWithValue('#f44336', SpreadsheetApp.InterpolationType.NUMBER, '-100')
      .setRanges([romiRange])
      .build();
    
    sheet.setConditionalFormatRules([romiRule]);
  }
  
  return startRow + Math.max(campaignData.length, 1) + 1;
}

function addKeywordsAnalysis(sheet, keywords, startRow) {
  const headers = [
    'Ключевое слово',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'CPL',
    'ROMI %',
    'Конкурентность',
    'Кампаний'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1976d2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  // Топ-30 ключевых слов
  const keywordData = keywords.slice(0, 30).map(keyword => [
    keyword.keyword,
    keyword.leads,
    keyword.success,
    keyword.conversionRate.toFixed(1) + '%',
    formatCurrency(keyword.cpl),
    keyword.romi.toFixed(1) + '%',
    keyword.competitiveness,
    keyword.campaigns.length
  ]);
  
  if (keywordData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, keywordData.length, headers.length);
    dataRange.setValues(keywordData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(keywordData.length, 1) + 1;
}

function addAdGroupsAnalysis(sheet, adGroups, startRow) {
  const headers = [
    'Группа объявлений',
    'Кампания',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'CPL',
    'ROMI %',
    'Ключевых слов'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#388e3c')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  // Топ-20 групп
  const groupData = adGroups.slice(0, 20).map(group => [
    group.group,
    group.campaign,
    group.leads,
    group.success,
    group.conversionRate.toFixed(1) + '%',
    formatCurrency(group.cpl),
    group.romi.toFixed(1) + '%',
    group.keywords.length
  ]);
  
  if (groupData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, groupData.length, headers.length);
    dataRange.setValues(groupData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(groupData.length, 1) + 1;
}

function addGeoDirectAnalysis(sheet, geoAnalysis, startRow) {
  const headers = [
    'Город',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'CPL',
    'ROMI %',
    'Топ кампания'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#f57c00')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const geoData = geoAnalysis.map(geo => [
    geo.city,
    geo.leads,
    geo.success,
    geo.conversionRate.toFixed(1) + '%',
    formatCurrency(geo.cpl),
    geo.romi.toFixed(1) + '%',
    geo.topCampaign
  ]);
  
  if (geoData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, geoData.length, headers.length);
    dataRange.setValues(geoData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, headers.length);
  
  return startRow + geoData.length + 1;
}
