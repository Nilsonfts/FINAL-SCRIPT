/**
 * МОДУЛЬ: СКВОЗНАЯ АНАЛИТИКА САЙТА
 * Анализ пути пользователя от первого касания до конверсии
 */

function runSiteCrossAnalysis() {
  console.log('Начинаем сквозную аналитику сайта...');
  
  try {
    const amoData = getWorkingAmoData();
    const siteData = getSiteFormsData();
    
    if (!amoData || amoData.length === 0) {
      console.log('Нет данных АМО для анализа');
      return;
    }
    
    const analysis = analyzeSiteCross(amoData, siteData);
    const sheet = createSiteCrossReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.SITE_CROSS}"`);
    console.log(`Проанализировано сессий: ${analysis.totalSessions}`);
    
  } catch (error) {
    console.error('Ошибка при сквозной аналитике сайта:', error);
    throw error;
  }
}

function getSiteFormsData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SITE_FORMS);
  if (!sheet) {
    console.log('Лист "Заявки с Сайта" не найден');
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

function analyzeSiteCross(amoData, siteData) {
  const userJourneys = {};
  const formAnalysis = {};
  const pageAnalysis = {};
  const conversionFunnel = {};
  const attributionAnalysis = {};
  
  let totalSessions = 0;
  let totalConversions = 0;
  
  // Анализ данных АМО с веб-аналитикой
  amoData.forEach(row => {
    const ymClientId = row[CONFIG.WORKING_AMO_COLUMNS.YM_CLIENT_ID] || '';
    const gaClientId = row[CONFIG.WORKING_AMO_COLUMNS.GA_CLIENT_ID] || '';
    const formName = row[CONFIG.WORKING_AMO_COLUMNS.FORMNAME] || '';
    const formId = row[CONFIG.WORKING_AMO_COLUMNS.FORMID] || '';
    const buttonText = row[CONFIG.WORKING_AMO_COLUMNS.BUTTON_TEXT] || '';
    const referer = row[CONFIG.WORKING_AMO_COLUMNS.REFERER] || '';
    const visitTime = row[CONFIG.WORKING_AMO_COLUMNS.VISIT_TIME] || '';
    const utmSource = row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || '';
    const utmMedium = row[CONFIG.WORKING_AMO_COLUMNS.UTM_MEDIUM] || '';
    const utmCampaign = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CAMPAIGN] || '';
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    
    totalSessions++;
    const isSuccess = isSuccessStatus(status);
    if (isSuccess) totalConversions++;
    
    // Анализ пользовательских путей
    const clientId = ymClientId || gaClientId || 'unknown';
    if (clientId !== 'unknown') {
      if (!userJourneys[clientId]) {
        userJourneys[clientId] = {
          clientId: clientId,
          firstSource: utmSource,
          firstMedium: utmMedium,
          firstCampaign: utmCampaign,
          firstTouch: createdAt,
          interactions: [],
          finalStatus: status,
          totalTouchPoints: 0,
          conversionTime: null
        };
      }
      
      const journey = userJourneys[clientId];
      journey.interactions.push({
        source: utmSource,
        medium: utmMedium,
        campaign: utmCampaign,
        formName: formName,
        buttonText: buttonText,
        timestamp: createdAt,
        page: extractPageFromReferer(referer)
      });
      
      journey.totalTouchPoints++;
      journey.finalStatus = status;
      
      if (isSuccess && !journey.conversionTime) {
        journey.conversionTime = createdAt;
      }
    }
    
    // Анализ форм
    if (formName && formName !== '') {
      if (!formAnalysis[formName]) {
        formAnalysis[formName] = {
          name: formName,
          submissions: 0,
          conversions: 0,
          sources: {},
          pages: {},
          avgTimeToSubmit: 0,
          topButtons: {}
        };
      }
      
      const form = formAnalysis[formName];
      form.submissions++;
      if (isSuccess) form.conversions++;
      
      if (utmSource) {
        form.sources[utmSource] = (form.sources[utmSource] || 0) + 1;
      }
      
      if (buttonText) {
        form.topButtons[buttonText] = (form.topButtons[buttonText] || 0) + 1;
      }
      
      const page = extractPageFromReferer(referer);
      if (page) {
        form.pages[page] = (form.pages[page] || 0) + 1;
      }
    }
    
    // Анализ страниц
    const page = extractPageFromReferer(referer);
    if (page && page !== '') {
      if (!pageAnalysis[page]) {
        pageAnalysis[page] = {
          page: page,
          visits: 0,
          conversions: 0,
          forms: {},
          sources: {},
          avgSessionTime: 0
        };
      }
      
      const pageData = pageAnalysis[page];
      pageData.visits++;
      if (isSuccess) pageData.conversions++;
      
      if (formName) {
        pageData.forms[formName] = (pageData.forms[formName] || 0) + 1;
      }
      
      if (utmSource) {
        pageData.sources[utmSource] = (pageData.sources[utmSource] || 0) + 1;
      }
    }
    
    // Воронка конверсии
    const funnelStage = determineFunnelStage(formName, buttonText, status);
    if (!conversionFunnel[funnelStage]) {
      conversionFunnel[funnelStage] = 0;
    }
    conversionFunnel[funnelStage]++;
    
    // Атрибуция
    const attributionModel = determineAttribution(utmSource, utmMedium);
    if (!attributionAnalysis[attributionModel]) {
      attributionAnalysis[attributionModel] = {
        model: attributionModel,
        touches: 0,
        conversions: 0,
        revenue: 0
      };
    }
    
    attributionAnalysis[attributionModel].touches++;
    if (isSuccess) {
      attributionAnalysis[attributionModel].conversions++;
      attributionAnalysis[attributionModel].revenue += formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    }
  });
  
  // Обработка пользовательских путей
  const processedJourneys = Object.values(userJourneys).map(journey => {
    const conversionTime = journey.conversionTime && journey.firstTouch ? 
      Math.round((new Date(journey.conversionTime) - new Date(journey.firstTouch)) / (1000 * 60 * 60)) : 0;
    
    return {
      ...journey,
      conversionTimeHours: conversionTime,
      isConverted: isSuccessStatus(journey.finalStatus),
      uniqueSources: [...new Set(journey.interactions.map(i => i.source))],
      touchPointsCount: journey.interactions.length
    };
  }).sort((a, b) => b.touchPointsCount - a.touchPointsCount);
  
  // Обработка форм
  const processedForms = Object.values(formAnalysis).map(form => ({
    ...form,
    conversionRate: form.submissions > 0 ? (form.conversions / form.submissions * 100) : 0,
    topSource: getTopEntry(form.sources),
    topPage: getTopEntry(form.pages),
    topButton: getTopEntry(form.topButtons)
  })).sort((a, b) => b.submissions - a.submissions);
  
  // Обработка страниц
  const processedPages = Object.values(pageAnalysis).map(page => ({
    ...page,
    conversionRate: page.visits > 0 ? (page.conversions / page.visits * 100) : 0,
    topForm: getTopEntry(page.forms),
    topSource: getTopEntry(page.sources)
  })).sort((a, b) => b.visits - a.visits);
  
  return {
    totalSessions,
    totalConversions,
    overallConversionRate: totalSessions > 0 ? (totalConversions / totalSessions * 100) : 0,
    userJourneys: processedJourneys,
    formAnalysis: processedForms,
    pageAnalysis: processedPages,
    conversionFunnel,
    attributionAnalysis: Object.values(attributionAnalysis),
    insights: generateSiteInsights(processedJourneys, processedForms, processedPages)
  };
}

function extractPageFromReferer(referer) {
  if (!referer) return 'Главная';
  
  try {
    const url = new URL(referer);
    let path = url.pathname;
    
    if (path === '/' || path === '') return 'Главная';
    
    // Убираем параметры и лишние символы
    path = path.replace(/\/$/, '').replace(/^\//, '');
    
    // Преобразуем в читаемый вид
    if (path.includes('about')) return 'О компании';
    if (path.includes('service')) return 'Услуги';
    if (path.includes('contact')) return 'Контакты';
    if (path.includes('price')) return 'Цены';
    if (path.includes('portfolio')) return 'Портфолио';
    if (path.includes('blog')) return 'Блог';
    
    return path.substring(0, 20); // Ограничиваем длину
    
  } catch (e) {
    return 'Неизвестно';
  }
}

function determineFunnelStage(formName, buttonText, status) {
  if (isSuccessStatus(status)) return '4. Конверсия';
  if (formName && formName !== '') return '3. Заявка';
  if (buttonText && buttonText !== '') return '2. Взаимодействие';
  return '1. Визит';
}

function determineAttribution(source, medium) {
  if (!source) return 'Прямой заход';
  
  const sourceLower = (source || '').toString().toLowerCase();
  const mediumLower = (medium || '').toString().toLowerCase();
  
  if (mediumLower.includes('cpc') || mediumLower.includes('ppc')) {
    return 'Платный поиск';
  }
  if (mediumLower.includes('organic')) {
    return 'Органический поиск';
  }
  if (mediumLower.includes('social')) {
    return 'Социальные сети';
  }
  if (mediumLower.includes('referral')) {
    return 'Переходы с сайтов';
  }
  if (mediumLower.includes('email')) {
    return 'Email-маркетинг';
  }
  
  return 'Другие каналы';
}

function generateSiteInsights(journeys, forms, pages) {
  const insights = [];
  
  // Анализ путей конверсии
  const convertedJourneys = journeys.filter(j => j.isConverted);
  const avgTouchPoints = convertedJourneys.length > 0 ? 
    convertedJourneys.reduce((sum, j) => sum + j.touchPointsCount, 0) / convertedJourneys.length : 0;
  
  insights.push(`Среднее количество касаний до конверсии: ${avgTouchPoints.toFixed(1)}`);
  
  // Лучшая форма
  if (forms.length > 0) {
    insights.push(`Самая эффективная форма: ${forms[0].name} (${forms[0].conversionRate.toFixed(1)}% конверсия)`);
  }
  
  // Лучшая страница
  if (pages.length > 0) {
    const bestPage = pages.sort((a, b) => b.conversionRate - a.conversionRate)[0];
    insights.push(`Самая конвертирующая страница: ${bestPage.page} (${bestPage.conversionRate.toFixed(1)}%)`);
  }
  
  // Среднее время до конверсии
  const avgConversionTime = convertedJourneys.length > 0 ?
    convertedJourneys.reduce((sum, j) => sum + j.conversionTimeHours, 0) / convertedJourneys.length : 0;
  
  if (avgConversionTime > 0) {
    insights.push(`Среднее время до конверсии: ${avgConversionTime.toFixed(1)} часов`);
  }
  
  return insights;
}

function createSiteCrossReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.SITE_CROSS);
  
  let currentRow = 1;
  
  // Общие метрики и инсайты
  currentRow = addSiteOverview(sheet, analysis, currentRow);
  currentRow += 2;
  
  // Анализ форм
  currentRow = addFormsAnalysis(sheet, analysis.formAnalysis, currentRow);
  currentRow += 2;
  
  // Анализ страниц
  currentRow = addPagesAnalysis(sheet, analysis.pageAnalysis, currentRow);
  currentRow += 2;
  
  // Воронка конверсии
  currentRow = addConversionFunnel(sheet, analysis.conversionFunnel, currentRow);
  currentRow += 2;
  
  // Атрибуция
  addAttributionAnalysis(sheet, analysis.attributionAnalysis, currentRow);
  
  return sheet;
}

function addSiteOverview(sheet, analysis, startRow) {
  const headers = ['Метрика сквозной аналитики', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1e88e5')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const overviewData = [
    ['Всего сессий', analysis.totalSessions],
    ['Конверсий', analysis.totalConversions],
    ['Общая конверсия сайта', analysis.overallConversionRate.toFixed(2) + '%'],
    ['Уникальных пользовательских путей', analysis.userJourneys.length],
    ['Форм на сайте', analysis.formAnalysis.length],
    ['Анализируемых страниц', analysis.pageAnalysis.length]
  ];
  
  // Добавляем инсайты
  analysis.insights.forEach(insight => {
    overviewData.push(['💡 ' + insight.split(':')[0], insight.split(':')[1] || insight]);
  });
  
  const dataRange = sheet.getRange(startRow + 1, 1, overviewData.length, 2);
  dataRange.setValues(overviewData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + overviewData.length + 1;
}

function addFormsAnalysis(sheet, forms, startRow) {
  const headers = [
    'Название формы',
    'Отправлений',
    'Конверсий',
    'Конверсия %',
    'Топ источник',
    'Топ страница',
    'Топ кнопка'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#43a047')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const formData = forms.map(form => [
    form.name,
    form.submissions,
    form.conversions,
    form.conversionRate.toFixed(1) + '%',
    form.topSource,
    form.topPage,
    form.topButton
  ]);
  
  if (formData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, formData.length, headers.length);
    dataRange.setValues(formData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(formData.length, 1) + 1;
}

function addPagesAnalysis(sheet, pages, startRow) {
  const headers = [
    'Страница',
    'Визитов',
    'Конверсий',
    'Конверсия %',
    'Топ форма',
    'Топ источник'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#fb8c00')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const pageData = pages.slice(0, 20).map(page => [
    page.page,
    page.visits,
    page.conversions,
    page.conversionRate.toFixed(1) + '%',
    page.topForm,
    page.topSource
  ]);
  
  if (pageData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, pageData.length, headers.length);
    dataRange.setValues(pageData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(pageData.length, 1) + 1;
}

function addConversionFunnel(sheet, funnel, startRow) {
  const headers = ['Этап воронки', 'Количество', 'Процент от общего'];
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#8e24aa')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const total = Object.values(funnel).reduce((sum, val) => sum + val, 0);
  const funnelData = Object.entries(funnel)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([stage, count]) => [
      stage,
      count,
      total > 0 ? (count / total * 100).toFixed(1) + '%' : '0%'
    ]);
  
  if (funnelData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, funnelData.length, headers.length);
    dataRange.setValues(funnelData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(funnelData.length, 1) + 1;
}

function addAttributionAnalysis(sheet, attribution, startRow) {
  const headers = [
    'Модель атрибуции',
    'Касаний',
    'Конверсий',
    'Конверсия %',
    'Выручка'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d32f2f')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const attributionData = attribution
    .sort((a, b) => b.touches - a.touches)
    .map(attr => [
      attr.model,
      attr.touches,
      attr.conversions,
      attr.touches > 0 ? (attr.conversions / attr.touches * 100).toFixed(1) + '%' : '0%',
      formatCurrency(attr.revenue)
    ]);
  
  if (attributionData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, attributionData.length, headers.length);
    dataRange.setValues(attributionData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, headers.length);
  
  return startRow + attributionData.length + 1;
}
