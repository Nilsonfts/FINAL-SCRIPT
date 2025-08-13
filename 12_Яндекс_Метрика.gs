/**
 * МОДУЛЬ: ЯНДЕКС.МЕТРИКА ИНТЕГРАЦИЯ
 * Анализ данных Яндекс.Метрики в связке с АМО
 */

function runMetrikaAnalysis() {
  console.log('Начинаем анализ Яндекс.Метрики...');
  
  try {
    const amoData = getWorkingAmoData();
    if (!amoData || amoData.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const analysis = analyzeMetrikaData(amoData);
    const sheet = createMetrikaReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.METRICA}"`);
    console.log(`Проанализировано сессий: ${analysis.totalSessions}`);
    
  } catch (error) {
    console.error('Ошибка при анализе Яндекс.Метрики:', error);
    throw error;
  }
}

function analyzeMetrikaData(data) {
  const clientAnalysis = {};
  const goalAnalysis = {};
  const trafficAnalysis = {};
  const behaviorAnalysis = {};
  const geoAnalysis = {};
  
  let totalSessions = 0;
  let totalConversions = 0;
  let uniqueVisitors = new Set();
  
  data.forEach(row => {
    const ymClientId = row[CONFIG.WORKING_AMO_COLUMNS.YM_CLIENT_ID] || '';
    const ymUid = row[CONFIG.WORKING_AMO_COLUMNS.YM_UID] || '';
    const visitTime = row[CONFIG.WORKING_AMO_COLUMNS.VISIT_TIME] || '';
    const source = parseUtmSource(row);
    const medium = row[CONFIG.WORKING_AMO_COLUMNS.UTM_MEDIUM] || '';
    const campaign = row[CONFIG.WORKING_AMO_COLUMNS.UTM_CAMPAIGN] || '';
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const referer = row[CONFIG.WORKING_AMO_COLUMNS.REFERER] || '';
    const formName = row[CONFIG.WORKING_AMO_COLUMNS.FORMNAME] || '';
    const cityTag = row[CONFIG.WORKING_AMO_COLUMNS.CITY_TAG] || '';
    
    totalSessions++;
    const isSuccess = isSuccessStatus(status);
    if (isSuccess) totalConversions++;
    
    const clientKey = ymClientId || ymUid || 'unknown';
    if (clientKey !== 'unknown') {
      uniqueVisitors.add(clientKey);
    }
    
    // Анализ клиентов Метрики
    if (clientKey !== 'unknown') {
      if (!clientAnalysis[clientKey]) {
        clientAnalysis[clientKey] = {
          clientId: clientKey,
          sessions: 0,
          conversions: 0,
          firstVisit: createdAt,
          lastVisit: createdAt,
          sources: new Set(),
          campaigns: new Set(),
          forms: new Set(),
          totalTime: 0,
          avgSessionTime: 0,
          bounceRate: 0
        };
      }
      
      const client = clientAnalysis[clientKey];
      client.sessions++;
      if (isSuccess) client.conversions++;
      
      client.sources.add(source);
      if (campaign) client.campaigns.add(campaign);
      if (formName) client.forms.add(formName);
      
      if (createdAt) {
        if (!client.firstVisit || new Date(createdAt) < new Date(client.firstVisit)) {
          client.firstVisit = createdAt;
        }
        if (!client.lastVisit || new Date(createdAt) > new Date(client.lastVisit)) {
          client.lastVisit = createdAt;
        }
      }
    }
    
    // Анализ целей (конверсий)
    if (formName && formName !== '') {
      const goalKey = `goal_${formName}`;
      if (!goalAnalysis[goalKey]) {
        goalAnalysis[goalKey] = {
          goalName: formName,
          triggers: 0,
          conversions: 0,
          conversionRate: 0,
          sources: {},
          avgTimeToGoal: 0,
          value: 0
        };
      }
      
      const goal = goalAnalysis[goalKey];
      goal.triggers++;
      if (isSuccess) {
        goal.conversions++;
        goal.value += formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
      }
      
      if (!goal.sources[source]) goal.sources[source] = 0;
      goal.sources[source]++;
    }
    
    // Анализ трафика
    const trafficKey = determineTrafficType(source, medium, referer);
    if (!trafficAnalysis[trafficKey]) {
      trafficAnalysis[trafficKey] = {
        type: trafficKey,
        sessions: 0,
        users: new Set(),
        conversions: 0,
        bounceRate: 0,
        avgSessionDuration: 0,
        pageviews: 0
      };
    }
    
    trafficAnalysis[trafficKey].sessions++;
    if (clientKey !== 'unknown') {
      trafficAnalysis[trafficKey].users.add(clientKey);
    }
    if (isSuccess) trafficAnalysis[trafficKey].conversions++;
    
    // Поведенческий анализ
    const hour = visitTime ? new Date(visitTime).getHours() : 
                  (createdAt ? new Date(createdAt).getHours() : 12);
    
    if (!behaviorAnalysis[hour]) {
      behaviorAnalysis[hour] = {
        hour: hour,
        sessions: 0,
        conversions: 0,
        avgConversionRate: 0
      };
    }
    
    behaviorAnalysis[hour].sessions++;
    if (isSuccess) behaviorAnalysis[hour].conversions++;
    
    // Географический анализ
    if (cityTag && cityTag !== '') {
      if (!geoAnalysis[cityTag]) {
        geoAnalysis[cityTag] = {
          city: cityTag,
          sessions: 0,
          users: new Set(),
          conversions: 0,
          topSources: {},
          conversionRate: 0
        };
      }
      
      const geo = geoAnalysis[cityTag];
      geo.sessions++;
      if (clientKey !== 'unknown') geo.users.add(clientKey);
      if (isSuccess) geo.conversions++;
      
      if (!geo.topSources[source]) geo.topSources[source] = 0;
      geo.topSources[source]++;
    }
  });
  
  // Постобработка данных
  Object.values(clientAnalysis).forEach(client => {
    client.sources = Array.from(client.sources);
    client.campaigns = Array.from(client.campaigns);
    client.forms = Array.from(client.forms);
    client.conversionRate = client.sessions > 0 ? (client.conversions / client.sessions * 100) : 0;
    
    if (client.firstVisit && client.lastVisit) {
      const timeDiff = new Date(client.lastVisit) - new Date(client.firstVisit);
      client.customerLifetime = Math.round(timeDiff / (1000 * 60 * 60 * 24)); // дни
    }
  });
  
  Object.values(goalAnalysis).forEach(goal => {
    goal.conversionRate = goal.triggers > 0 ? (goal.conversions / goal.triggers * 100) : 0;
    goal.topSource = getTopEntry(goal.sources);
    goal.avgValue = goal.conversions > 0 ? goal.value / goal.conversions : 0;
  });
  
  Object.values(trafficAnalysis).forEach(traffic => {
    traffic.users = traffic.users.size;
    traffic.conversionRate = traffic.sessions > 0 ? (traffic.conversions / traffic.sessions * 100) : 0;
    traffic.sessionsPerUser = traffic.users > 0 ? traffic.sessions / traffic.users : 0;
  });
  
  Object.values(behaviorAnalysis).forEach(behavior => {
    behavior.avgConversionRate = behavior.sessions > 0 ? (behavior.conversions / behavior.sessions * 100) : 0;
  });
  
  Object.values(geoAnalysis).forEach(geo => {
    geo.users = geo.users.size;
    geo.conversionRate = geo.sessions > 0 ? (geo.conversions / geo.sessions * 100) : 0;
    geo.topSource = getTopEntry(geo.topSources);
    geo.sessionsPerUser = geo.users > 0 ? geo.sessions / geo.users : 0;
  });
  
  return {
    totalSessions,
    totalConversions,
    uniqueVisitors: uniqueVisitors.size,
    overallConversionRate: totalSessions > 0 ? (totalConversions / totalSessions * 100) : 0,
    clientAnalysis: Object.values(clientAnalysis).sort((a, b) => b.sessions - a.sessions),
    goalAnalysis: Object.values(goalAnalysis).sort((a, b) => b.triggers - a.triggers),
    trafficAnalysis: Object.values(trafficAnalysis).sort((a, b) => b.sessions - a.sessions),
    behaviorAnalysis: Object.values(behaviorAnalysis).sort((a, b) => a.hour - b.hour),
    geoAnalysis: Object.values(geoAnalysis).sort((a, b) => b.sessions - a.sessions),
    insights: generateMetrikaInsights(clientAnalysis, goalAnalysis, trafficAnalysis)
  };
}

function determineTrafficType(source, medium, referer) {
  if (!source && !medium) {
    if (referer && referer !== '') {
      return 'Переходы с сайтов';
    }
    return 'Прямые заходы';
  }
  
  const sourceLower = source.toLowerCase();
  const mediumLower = (medium || '').toLowerCase();
  
  if (mediumLower.includes('organic')) return 'Органический поиск';
  if (mediumLower.includes('cpc') || mediumLower.includes('ppc')) return 'Контекстная реклама';
  if (mediumLower.includes('social')) return 'Социальные сети';
  if (mediumLower.includes('email')) return 'Email-рассылки';
  if (mediumLower.includes('referral')) return 'Переходы с сайтов';
  if (sourceLower.includes('yandex') || sourceLower.includes('google')) return 'Поисковые системы';
  
  return 'Другие источники';
}

function generateMetrikaInsights(clientAnalysis, goalAnalysis, trafficAnalysis) {
  const insights = [];
  
  // Топ источник трафика
  const topTraffic = Object.values(trafficAnalysis)[0];
  if (topTraffic) {
    insights.push(`Основной источник трафика: ${topTraffic.type} (${topTraffic.sessions} сессий)`);
  }
  
  // Лучшая цель по конверсии
  const bestGoal = Object.values(goalAnalysis).sort((a, b) => b.conversionRate - a.conversionRate)[0];
  if (bestGoal) {
    insights.push(`Самая эффективная цель: ${bestGoal.goalName} (${bestGoal.conversionRate.toFixed(1)}%)`);
  }
  
  // Средняя сессионность
  const avgSessions = Object.values(clientAnalysis).length > 0 ?
    Object.values(clientAnalysis).reduce((sum, c) => sum + c.sessions, 0) / Object.values(clientAnalysis).length : 0;
  
  if (avgSessions > 0) {
    insights.push(`Среднее количество сессий на пользователя: ${avgSessions.toFixed(1)}`);
  }
  
  return insights;
}

function createMetrikaReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.METRICA);
  
  let currentRow = 1;
  
  // Общие метрики и инсайты
  currentRow = addMetrikaOverview(sheet, analysis, currentRow);
  currentRow += 2;
  
  // Анализ источников трафика
  currentRow = addTrafficAnalysis(sheet, analysis.trafficAnalysis, currentRow);
  currentRow += 2;
  
  // Анализ целей
  currentRow = addGoalsAnalysis(sheet, analysis.goalAnalysis, currentRow);
  currentRow += 2;
  
  // Поведенческий анализ
  currentRow = addBehaviorAnalysis(sheet, analysis.behaviorAnalysis, currentRow);
  currentRow += 2;
  
  // Географический анализ
  currentRow = addGeoMetrikaAnalysis(sheet, analysis.geoAnalysis, currentRow);
  currentRow += 2;
  
  // Топ клиенты
  addTopClients(sheet, analysis.clientAnalysis, currentRow);
  
  return sheet;
}

function addMetrikaOverview(sheet, analysis, startRow) {
  const headers = ['Метрика Яндекс.Метрики', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#ff6600')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const overviewData = [
    ['Всего сессий', analysis.totalSessions],
    ['Уникальных посетителей', analysis.uniqueVisitors],
    ['Конверсий', analysis.totalConversions],
    ['Общая конверсия', analysis.overallConversionRate.toFixed(2) + '%'],
    ['Среднее сессий на пользователя', analysis.uniqueVisitors > 0 ? (analysis.totalSessions / analysis.uniqueVisitors).toFixed(1) : '0'],
    ['Источников трафика', analysis.trafficAnalysis.length],
    ['Активных целей', analysis.goalAnalysis.length],
    ['Городов', analysis.geoAnalysis.length]
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

function addTrafficAnalysis(sheet, trafficAnalysis, startRow) {
  const headers = [
    'Тип трафика',
    'Сессии',
    'Пользователи',
    'Конверсии',
    'Конверсия %',
    'Сессий на пользователя'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1976d2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const trafficData = trafficAnalysis.map(traffic => [
    traffic.type,
    traffic.sessions,
    traffic.users,
    traffic.conversions,
    traffic.conversionRate.toFixed(2) + '%',
    traffic.sessionsPerUser.toFixed(1)
  ]);
  
  if (trafficData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, trafficData.length, headers.length);
    dataRange.setValues(trafficData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(trafficData.length, 1) + 1;
}

function addGoalsAnalysis(sheet, goalAnalysis, startRow) {
  const headers = [
    'Цель (форма)',
    'Срабатываний',
    'Конверсий',
    'Конверсия %',
    'Топ источник',
    'Средняя ценность'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#388e3c')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const goalData = goalAnalysis.map(goal => [
    goal.goalName,
    goal.triggers,
    goal.conversions,
    goal.conversionRate.toFixed(1) + '%',
    goal.topSource,
    formatCurrency(goal.avgValue)
  ]);
  
  if (goalData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, goalData.length, headers.length);
    dataRange.setValues(goalData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(goalData.length, 1) + 1;
}

function addBehaviorAnalysis(sheet, behaviorAnalysis, startRow) {
  const headers = [
    'Час',
    'Сессий',
    'Конверсий',
    'Конверсия %'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#7b1fa2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const behaviorData = behaviorAnalysis.map(behavior => [
    `${behavior.hour}:00`,
    behavior.sessions,
    behavior.conversions,
    behavior.avgConversionRate.toFixed(1) + '%'
  ]);
  
  if (behaviorData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, behaviorData.length, headers.length);
    dataRange.setValues(behaviorData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(behaviorData.length, 1) + 1;
}

function addGeoMetrikaAnalysis(sheet, geoAnalysis, startRow) {
  const headers = [
    'Город',
    'Сессий',
    'Пользователи',
    'Конверсии',
    'Конверсия %',
    'Топ источник'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#f57c00')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const geoData = geoAnalysis.slice(0, 15).map(geo => [
    geo.city,
    geo.sessions,
    geo.users,
    geo.conversions,
    geo.conversionRate.toFixed(1) + '%',
    geo.topSource
  ]);
  
  if (geoData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, geoData.length, headers.length);
    dataRange.setValues(geoData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(geoData.length, 1) + 1;
}

function addTopClients(sheet, clientAnalysis, startRow) {
  const headers = [
    'Client ID',
    'Сессий',
    'Конверсий',
    'Конверсия %',
    'Источники',
    'Время жизни (дни)'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d32f2f')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  // Топ-30 клиентов
  const clientData = clientAnalysis.slice(0, 30).map(client => [
    client.clientId.substring(0, 20) + '...', // Сокращаем ID
    client.sessions,
    client.conversions,
    client.conversionRate.toFixed(1) + '%',
    client.sources.slice(0, 3).join(', '),
    client.customerLifetime || 0
  ]);
  
  if (clientData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, clientData.length, headers.length);
    dataRange.setValues(clientData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, headers.length);
  
  return startRow + clientData.length + 1;
}
