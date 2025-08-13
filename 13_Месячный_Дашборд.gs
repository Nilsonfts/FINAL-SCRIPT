/**
 * МОДУЛЬ: МЕСЯЧНЫЙ ДАШБОРД
 * Комплексный анализ за текущий месяц с KPI
 */

function runMonthlyDashboard() {
  console.log('Создаем месячный дашборд...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const dashboard = createMonthlyDashboard(data);
    const sheet = createMonthlyDashboardReport(dashboard);
    
    console.log(`Создан дашборд на листе "${CONFIG.SHEETS.MONTHLY_DASHBOARD}"`);
    console.log(`Период: ${dashboard.currentMonth} ${dashboard.currentYear}`);
    
  } catch (error) {
    console.error('Ошибка при создании месячного дашборда:', error);
    throw error;
  }
}

function createMonthlyDashboard(data) {
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth();
  const currentYear = currentDate.getFullYear();
  const previousMonth = currentMonth === 0 ? 11 : currentMonth - 1;
  const previousYear = currentMonth === 0 ? currentYear - 1 : currentYear;
  
  const monthNames = [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
  ];
  
  // Основные метрики
  const currentMonthData = {
    leads: 0,
    success: 0,
    refused: 0,
    budget: 0,
    revenue: 0,
    avgDeal: 0,
    conversion: 0,
    managers: new Set(),
    sources: {},
    days: {}
  };
  
  const previousMonthData = {
    leads: 0,
    success: 0,
    refused: 0,
    budget: 0,
    revenue: 0
  };
  
  const channelAnalysis = {};
  const managerPerformance = {};
  const dailyMetrics = {};
  const weeklyTrends = {};
  const goals = {
    leadsGoal: 100, // Целевые показатели
    revenueGoal: 5000000,
    conversionGoal: 25
  };
  
  // Анализ по дням месяца
  const daysInMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day++) {
    dailyMetrics[day] = {
      day: day,
      leads: 0,
      success: 0,
      revenue: 0,
      calls: 0
    };
  }
  
  // Обработка данных
  data.forEach(row => {
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    if (!createdAt) return;
    
    const dealDate = new Date(createdAt);
    const dealMonth = dealDate.getMonth();
    const dealYear = dealDate.getFullYear();
    const dealDay = dealDate.getDate();
    
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    const budget = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    const source = parseUtmSource(row);
    const channelType = getChannelType(source);
    
    const isSuccess = isSuccessStatus(status);
    const isRefusal = isRefusalStatus(status);
    
    // Данные текущего месяца
    if (dealMonth === currentMonth && dealYear === currentYear) {
      currentMonthData.leads++;
      currentMonthData.budget += budget;
      currentMonthData.managers.add(responsible);
      
      if (isSuccess) {
        currentMonthData.success++;
        currentMonthData.revenue += factAmount;
      }
      if (isRefusal) currentMonthData.refused++;
      
      // Источники
      if (!currentMonthData.sources[source]) currentMonthData.sources[source] = 0;
      currentMonthData.sources[source]++;
      
      // По дням
      if (dealDay && dailyMetrics[dealDay]) {
        dailyMetrics[dealDay].leads++;
        if (isSuccess) {
          dailyMetrics[dealDay].success++;
          dailyMetrics[dealDay].revenue += factAmount;
        }
      }
      
      // Анализ каналов
      if (!channelAnalysis[channelType]) {
        channelAnalysis[channelType] = {
          channel: channelType,
          leads: 0,
          success: 0,
          budget: 0,
          revenue: 0,
          conversion: 0,
          romi: 0
        };
      }
      
      channelAnalysis[channelType].leads++;
      channelAnalysis[channelType].budget += budget;
      if (isSuccess) {
        channelAnalysis[channelType].success++;
        channelAnalysis[channelType].revenue += factAmount;
      }
      
      // Производительность менеджеров
      if (responsible && responsible !== '') {
        if (!managerPerformance[responsible]) {
          managerPerformance[responsible] = {
            manager: responsible,
            leads: 0,
            success: 0,
            revenue: 0,
            avgDeal: 0,
            efficiency: 0
          };
        }
        
        managerPerformance[responsible].leads++;
        if (isSuccess) {
          managerPerformance[responsible].success++;
          managerPerformance[responsible].revenue += factAmount;
        }
      }
    }
    
    // Данные предыдущего месяца для сравнения
    if (dealMonth === previousMonth && dealYear === previousYear) {
      previousMonthData.leads++;
      previousMonthData.budget += budget;
      
      if (isSuccess) {
        previousMonthData.success++;
        previousMonthData.revenue += factAmount;
      }
      if (isRefusal) previousMonthData.refused++;
    }
  });
  
  // Вычисляем финальные метрики
  currentMonthData.conversion = currentMonthData.leads > 0 ? 
    (currentMonthData.success / currentMonthData.leads * 100) : 0;
  
  currentMonthData.avgDeal = currentMonthData.success > 0 ? 
    currentMonthData.revenue / currentMonthData.success : 0;
  
  currentMonthData.managers = currentMonthData.managers.size;
  
  // Обрабатываем каналы
  Object.values(channelAnalysis).forEach(channel => {
    channel.conversion = channel.leads > 0 ? (channel.success / channel.leads * 100) : 0;
    channel.romi = channel.budget > 0 ? ((channel.revenue - channel.budget) / channel.budget * 100) : 0;
  });
  
  // Обрабатываем менеджеров
  Object.values(managerPerformance).forEach(manager => {
    manager.avgDeal = manager.success > 0 ? manager.revenue / manager.success : 0;
    manager.conversion = manager.leads > 0 ? (manager.success / manager.leads * 100) : 0;
    manager.efficiency = calculateManagerEfficiency({
      conversionRate: manager.conversion,
      totalDeals: manager.leads,
      currentMonthDeals: manager.leads
    });
  });
  
  // Создаем недельные тренды
  const weeks = groupDaysByWeeks(dailyMetrics);
  Object.entries(weeks).forEach(([weekNum, days]) => {
    weeklyTrends[weekNum] = {
      week: `Неделя ${weekNum}`,
      leads: days.reduce((sum, day) => sum + day.leads, 0),
      success: days.reduce((sum, day) => sum + day.success, 0),
      revenue: days.reduce((sum, day) => sum + day.revenue, 0)
    };
  });
  
  // Сравнение с предыдущим месяцем
  const comparison = {
    leadsChange: calculateChange(currentMonthData.leads, previousMonthData.leads),
    revenueChange: calculateChange(currentMonthData.revenue, previousMonthData.revenue),
    conversionChange: {
      current: currentMonthData.conversion,
      previous: previousMonthData.leads > 0 ? (previousMonthData.success / previousMonthData.leads * 100) : 0
    }
  };
  
  // Прогресс к целям
  const goalProgress = {
    leadsProgress: currentMonthData.leads / goals.leadsGoal * 100,
    revenueProgress: currentMonthData.revenue / goals.revenueGoal * 100,
    conversionProgress: currentMonthData.conversion / goals.conversionGoal * 100
  };
  
  return {
    currentMonth: monthNames[currentMonth],
    currentYear: currentYear,
    currentMonthData,
    previousMonthData,
    comparison,
    channelAnalysis: Object.values(channelAnalysis).sort((a, b) => b.leads - a.leads),
    managerPerformance: Object.values(managerPerformance).sort((a, b) => b.efficiency - a.efficiency),
    dailyMetrics: Object.values(dailyMetrics),
    weeklyTrends: Object.values(weeklyTrends),
    goalProgress,
    goals,
    insights: generateDashboardInsights(currentMonthData, comparison, goalProgress)
  };
}

function groupDaysByWeeks(dailyMetrics) {
  const weeks = {};
  Object.values(dailyMetrics).forEach(day => {
    const weekNum = Math.ceil(day.day / 7);
    if (!weeks[weekNum]) weeks[weekNum] = [];
    weeks[weekNum].push(day);
  });
  return weeks;
}

function calculateChange(current, previous) {
  if (previous === 0) return { value: current, percent: 100, trend: 'up' };
  
  const change = ((current - previous) / previous) * 100;
  return {
    value: current - previous,
    percent: Math.abs(change),
    trend: change >= 0 ? 'up' : 'down'
  };
}

function generateDashboardInsights(currentData, comparison, goalProgress) {
  const insights = [];
  
  // Прогресс по лидам
  if (goalProgress.leadsProgress >= 100) {
    insights.push('🎯 Цель по количеству лидов выполнена!');
  } else if (goalProgress.leadsProgress >= 80) {
    insights.push('📈 Близки к выполнению цели по лидам');
  } else {
    insights.push('⚠️ Требуется усилить работу по привлечению лидов');
  }
  
  // Конверсия
  if (currentData.conversion >= 30) {
    insights.push('💚 Отличная конверсия лидов');
  } else if (currentData.conversion >= 20) {
    insights.push('📊 Хорошая конверсия, есть потенциал для роста');
  } else {
    insights.push('🔴 Низкая конверсия, нужно улучшать качество работы с лидами');
  }
  
  // Тенденции
  if (comparison.leadsChange.trend === 'up' && comparison.leadsChange.percent > 10) {
    insights.push(`📈 Рост количества лидов на ${comparison.leadsChange.percent.toFixed(1)}% к прошлому месяцу`);
  }
  
  if (comparison.revenueChange.trend === 'up' && comparison.revenueChange.percent > 15) {
    insights.push(`💰 Значительный рост выручки: +${comparison.revenueChange.percent.toFixed(1)}%`);
  }
  
  return insights;
}

function createMonthlyDashboardReport(dashboard) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.MONTHLY_DASHBOARD);
  
  let currentRow = 1;
  
  // Заголовок дашборда
  currentRow = addDashboardHeader(sheet, dashboard, currentRow);
  currentRow += 2;
  
  // Основные KPI текущего месяца
  currentRow = addMainKpis(sheet, dashboard, currentRow);
  currentRow += 2;
  
  // Сравнение с предыдущим месяцем
  currentRow = addMonthComparison(sheet, dashboard, currentRow);
  currentRow += 2;
  
  // Прогресс к целям
  currentRow = addGoalProgress(sheet, dashboard, currentRow);
  currentRow += 2;
  
  // Производительность каналов
  currentRow = addChannelDashboard(sheet, dashboard.channelAnalysis, currentRow);
  currentRow += 2;
  
  // Топ менеджеры
  currentRow = addManagerDashboard(sheet, dashboard.managerPerformance, currentRow);
  currentRow += 2;
  
  // Недельные тренды
  currentRow = addWeeklyTrends(sheet, dashboard.weeklyTrends, currentRow);
  currentRow += 2;
  
  // Инсайты и рекомендации
  addInsightsDashboard(sheet, dashboard.insights, currentRow);
  
  return sheet;
}

function addDashboardHeader(sheet, dashboard, startRow) {
  const title = `ДАШБОРД ${dashboard.currentMonth.toUpperCase()} ${dashboard.currentYear}`;
  const titleRange = sheet.getRange(startRow, 1, 1, 8);
  titleRange.merge();
  titleRange.setValue(title);
  titleRange.setBackground('#1a237e')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontSize(16)
            .setHorizontalAlignment('center')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + 1;
}

function addMainKpis(sheet, dashboard, startRow) {
  const headers = ['Основные KPI', 'Значение', 'Статус'];
  const headerRange = sheet.getRange(startRow, 1, 1, 3);
  headerRange.setValues([headers]);
  headerRange.setBackground('#3949ab')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const kpiData = [
    ['Всего лидов за месяц', dashboard.currentMonthData.leads, getKpiStatus(dashboard.currentMonthData.leads, dashboard.goals.leadsGoal)],
    ['Успешных сделок', dashboard.currentMonthData.success, ''],
    ['Конверсия лидов', dashboard.currentMonthData.conversion.toFixed(1) + '%', getKpiStatus(dashboard.currentMonthData.conversion, dashboard.goals.conversionGoal)],
    ['Общая выручка', formatCurrency(dashboard.currentMonthData.revenue), getKpiStatus(dashboard.currentMonthData.revenue, dashboard.goals.revenueGoal)],
    ['Средний чек', formatCurrency(dashboard.currentMonthData.avgDeal), ''],
    ['Активных менеджеров', dashboard.currentMonthData.managers, ''],
    ['Маркетинговый бюджет', formatCurrency(dashboard.currentMonthData.budget), ''],
    ['ROMI', dashboard.currentMonthData.budget > 0 ? (((dashboard.currentMonthData.revenue - dashboard.currentMonthData.budget) / dashboard.currentMonthData.budget) * 100).toFixed(1) + '%' : '0%', '']
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, kpiData.length, 3);
  dataRange.setValues(kpiData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + kpiData.length + 1;
}

function getKpiStatus(current, target) {
  const progress = (current / target) * 100;
  if (progress >= 100) return '✅ Выполнено';
  if (progress >= 80) return '🟡 Близко';
  return '🔴 Отстаем';
}

function addMonthComparison(sheet, dashboard, startRow) {
  const headers = ['Сравнение с предыдущим месяцем', 'Текущий', 'Предыдущий', 'Изменение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 4);
  headerRange.setValues([headers]);
  headerRange.setBackground('#2e7d32')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const comparisonData = [
    [
      'Лиды',
      dashboard.currentMonthData.leads,
      dashboard.previousMonthData.leads,
      formatChange(dashboard.comparison.leadsChange)
    ],
    [
      'Успешные сделки',
      dashboard.currentMonthData.success,
      dashboard.previousMonthData.success,
      ''
    ],
    [
      'Выручка',
      formatCurrency(dashboard.currentMonthData.revenue),
      formatCurrency(dashboard.previousMonthData.revenue),
      formatChange(dashboard.comparison.revenueChange)
    ],
    [
      'Конверсия',
      dashboard.comparison.conversionChange.current.toFixed(1) + '%',
      dashboard.comparison.conversionChange.previous.toFixed(1) + '%',
      formatConversionChange(dashboard.comparison.conversionChange)
    ]
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, comparisonData.length, 4);
  dataRange.setValues(comparisonData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + comparisonData.length + 1;
}

function formatChange(change) {
  const arrow = change.trend === 'up' ? '↗️' : '↘️';
  const sign = change.trend === 'up' ? '+' : '-';
  return `${arrow} ${sign}${change.percent.toFixed(1)}%`;
}

function formatConversionChange(conversionChange) {
  const diff = conversionChange.current - conversionChange.previous;
  const arrow = diff >= 0 ? '↗️' : '↘️';
  const sign = diff >= 0 ? '+' : '';
  return `${arrow} ${sign}${diff.toFixed(1)}%`;
}

function addGoalProgress(sheet, dashboard, startRow) {
  const headers = ['Прогресс к целям месяца', 'Цель', 'Текущее', 'Прогресс %', 'Осталось'];
  const headerRange = sheet.getRange(startRow, 1, 1, 5);
  headerRange.setValues([headers]);
  headerRange.setBackground('#f57c00')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const progressData = [
    [
      'Лиды',
      dashboard.goals.leadsGoal,
      dashboard.currentMonthData.leads,
      Math.min(dashboard.goalProgress.leadsProgress, 100).toFixed(1) + '%',
      Math.max(0, dashboard.goals.leadsGoal - dashboard.currentMonthData.leads)
    ],
    [
      'Выручка',
      formatCurrency(dashboard.goals.revenueGoal),
      formatCurrency(dashboard.currentMonthData.revenue),
      Math.min(dashboard.goalProgress.revenueProgress, 100).toFixed(1) + '%',
      formatCurrency(Math.max(0, dashboard.goals.revenueGoal - dashboard.currentMonthData.revenue))
    ],
    [
      'Конверсия',
      dashboard.goals.conversionGoal + '%',
      dashboard.currentMonthData.conversion.toFixed(1) + '%',
      Math.min(dashboard.goalProgress.conversionProgress, 100).toFixed(1) + '%',
      dashboard.currentMonthData.conversion < dashboard.goals.conversionGoal ? 
        `+${(dashboard.goals.conversionGoal - dashboard.currentMonthData.conversion).toFixed(1)}%` : '0%'
    ]
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, progressData.length, 5);
  dataRange.setValues(progressData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + progressData.length + 1;
}

function addChannelDashboard(sheet, channels, startRow) {
  const headers = [
    'Канал',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'Выручка',
    'ROMI %'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#7b1fa2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const channelData = channels.slice(0, 8).map(channel => [
    channel.channel,
    channel.leads,
    channel.success,
    channel.conversion.toFixed(1) + '%',
    formatCurrency(channel.revenue),
    channel.romi.toFixed(1) + '%'
  ]);
  
  if (channelData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, channelData.length, headers.length);
    dataRange.setValues(channelData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(channelData.length, 1) + 1;
}

function addManagerDashboard(sheet, managers, startRow) {
  const headers = [
    'Менеджер',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'Средний чек',
    'Эффективность'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d32f2f')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const managerData = managers.slice(0, 10).map(manager => [
    manager.manager,
    manager.leads,
    manager.success,
    manager.conversion.toFixed(1) + '%',
    formatCurrency(manager.avgDeal),
    manager.efficiency.toFixed(1)
  ]);
  
  if (managerData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, managerData.length, headers.length);
    dataRange.setValues(managerData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(managerData.length, 1) + 1;
}

function addWeeklyTrends(sheet, trends, startRow) {
  const headers = ['Неделя', 'Лиды', 'Успешные', 'Выручка'];
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#00695c')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const trendData = trends.map(trend => [
    trend.week,
    trend.leads,
    trend.success,
    formatCurrency(trend.revenue)
  ]);
  
  if (trendData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, trendData.length, headers.length);
    dataRange.setValues(trendData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(trendData.length, 1) + 1;
}

function addInsightsDashboard(sheet, insights, startRow) {
  const headers = ['Инсайты и рекомендации'];
  const headerRange = sheet.getRange(startRow, 1, 1, 1);
  headerRange.setValues([headers]);
  headerRange.setBackground('#4a148c')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const insightData = insights.map(insight => [insight]);
  
  if (insightData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, insightData.length, 1);
    dataRange.setValues(insightData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    dataRange.setWrap(true);
  }
  
  sheet.autoResizeColumns(1, 8);
  
  return startRow + insightData.length + 1;
}
