/**
 * ЕЖЕДНЕВНАЯ СТАТИСТИКА
 * Детальная аналитика по дням с трендами и сравнениями
 * @fileoverview Модуль ежедневной статистики лидов, конверсий и выручки
 */

/**
 * Основная функция создания ежедневной статистики
 * Создаёт отчёт с детальной аналитикой по дням
 */
function generateDailyStatistics() {
  try {
    logInfo_('DAILY_STATS', 'Начало генерации ежедневной статистики');
    
    const startTime = new Date();
    
    // Получаем данные для ежедневной статистики
    const dailyData = getDailyStatisticsData_();
    
    if (!dailyData || dailyData.dailyStats.length === 0) {
      logWarning_('DAILY_STATS', 'Нет данных для ежедневной статистики');
      createEmptyDailyReport_();
      return;
    }
    
    // Создаём или обновляем лист
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.DAILY_STATISTICS);
    clearSheetData_(sheet);
    applySheetFormatting_(sheet, 'Ежедневная статистика');
    
    // Строим структуру отчёта
    createDailyStatisticsStructure_(sheet, dailyData);
    
    // Добавляем диаграммы
    addDailyStatisticsCharts_(sheet, dailyData);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('DAILY_STATS', `Ежедневная статистика сгенерирована за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.DAILY_STATISTICS);
    
  } catch (error) {
    logError_('DAILY_STATS', 'Ошибка генерации ежедневной статистики', error);
    throw error;
  }
}

/**
 * Получает данные для ежедневной статистики
 * @returns {Object} Данные ежедневной статистики
 * @private
 */
function getDailyStatisticsData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return null;
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Ежедневная статистика
  const dailyStats = generateDailyStats_(headers, rows);
  
  // Тренды и сравнения
  const trendAnalysis = analyzeDailyTrends_(dailyStats);
  
  // Анализ по дням недели
  const weekdayAnalysis = analyzeWeekdayPatterns_(headers, rows);
  
  // Почасовая активность
  const hourlyAnalysis = analyzeHourlyPatterns_(headers, rows);
  
  // Аномалии и выбросы
  const anomalies = detectDailyAnomalies_(dailyStats);
  
  // Сравнение с предыдущими периодами
  const periodComparison = comparePeriods_(dailyStats);
  
  return {
    dailyStats: dailyStats,
    trendAnalysis: trendAnalysis,
    weekdayAnalysis: weekdayAnalysis,
    hourlyAnalysis: hourlyAnalysis,
    anomalies: anomalies,
    periodComparison: periodComparison,
    totalDays: dailyStats.length,
    analysisDate: getCurrentDateMoscow_()
  };
}

/**
 * Генерирует ежедневную статистику
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Ежедневная статистика
 * @private
 */
function generateDailyStats_(headers, rows) {
  const dailyStats = {};
  
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  
  rows.forEach(row => {
    const createdDate = parseDate_(row[createdIndex]);
    if (!createdDate) return;
    
    const dateKey = formatDate_(createdDate, 'YYYY-MM-DD');
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const channel = row[channelIndex] || 'Неизвестно';
    const source = row[sourceIndex] || 'direct';
    
    if (!dailyStats[dateKey]) {
      dailyStats[dateKey] = {
        date: dateKey,
        dateObj: new Date(createdDate.getFullYear(), createdDate.getMonth(), createdDate.getDate()),
        dayOfWeek: getDayName_(createdDate.getDay()),
        totalLeads: 0,
        successfulDeals: 0,
        failedDeals: 0,
        inProgressDeals: 0,
        totalRevenue: 0,
        avgDealValue: 0,
        conversionRate: 0,
        channels: {},
        sources: {},
        hourlyDistribution: {}
      };
    }
    
    const dayStats = dailyStats[dateKey];
    dayStats.totalLeads++;
    
    // Подсчёт по статусам
    if (status === 'success') {
      dayStats.successfulDeals++;
      dayStats.totalRevenue += budget;
    } else if (status === 'failure') {
      dayStats.failedDeals++;
    } else {
      dayStats.inProgressDeals++;
    }
    
    // Каналы
    dayStats.channels[channel] = (dayStats.channels[channel] || 0) + 1;
    
    // Источники
    dayStats.sources[source] = (dayStats.sources[source] || 0) + 1;
    
    // Почасовое распределение
    const hour = createdDate.getHours();
    dayStats.hourlyDistribution[hour] = (dayStats.hourlyDistribution[hour] || 0) + 1;
  });
  
  // Вычисляем производные метрики и сортируем по дате
  return Object.values(dailyStats)
    .map(dayStats => {
      dayStats.conversionRate = dayStats.totalLeads > 0 ? 
        (dayStats.successfulDeals / dayStats.totalLeads * 100) : 0;
      dayStats.avgDealValue = dayStats.successfulDeals > 0 ? 
        (dayStats.totalRevenue / dayStats.successfulDeals) : 0;
      
      // Топ канал и источник дня
      dayStats.topChannel = Object.entries(dayStats.channels)
        .sort(([,a], [,b]) => b - a)[0]?.[0] || 'Нет данных';
      dayStats.topSource = Object.entries(dayStats.sources)
        .sort(([,a], [,b]) => b - a)[0]?.[0] || 'Нет данных';
      
      // Пиковый час
      dayStats.peakHour = Object.entries(dayStats.hourlyDistribution)
        .sort(([,a], [,b]) => b - a)[0]?.[0] || 'Нет данных';
      
      return dayStats;
    })
    .sort((a, b) => a.dateObj - b.dateObj);
}

/**
 * Анализирует ежедневные тренды
 * @param {Array} dailyStats - Ежедневная статистика
 * @returns {Object} Анализ трендов
 * @private
 */
function analyzeDailyTrends_(dailyStats) {
  if (dailyStats.length < 7) {
    return { trends: {}, weeklyComparison: {}, growthRates: {} };
  }
  
  const trends = {
    leads: calculateTrend_(dailyStats.map(day => day.totalLeads)),
    conversions: calculateTrend_(dailyStats.map(day => day.successfulDeals)),
    revenue: calculateTrend_(dailyStats.map(day => day.totalRevenue)),
    conversionRate: calculateTrend_(dailyStats.map(day => day.conversionRate))
  };
  
  // Сравнение с предыдущей неделей
  const weeklyComparison = {};
  if (dailyStats.length >= 14) {
    const lastWeek = dailyStats.slice(-7);
    const prevWeek = dailyStats.slice(-14, -7);
    
    weeklyComparison.leadsGrowth = calculateGrowthRate_(
      prevWeek.reduce((sum, day) => sum + day.totalLeads, 0),
      lastWeek.reduce((sum, day) => sum + day.totalLeads, 0)
    );
    
    weeklyComparison.revenueGrowth = calculateGrowthRate_(
      prevWeek.reduce((sum, day) => sum + day.totalRevenue, 0),
      lastWeek.reduce((sum, day) => sum + day.totalRevenue, 0)
    );
    
    weeklyComparison.conversionGrowth = calculateGrowthRate_(
      prevWeek.reduce((sum, day) => sum + day.successfulDeals, 0),
      lastWeek.reduce((sum, day) => sum + day.successfulDeals, 0)
    );
  }
  
  // Скользящие средние
  const movingAverages = {
    leads7day: calculateMovingAverage_(dailyStats.map(day => day.totalLeads), 7),
    revenue7day: calculateMovingAverage_(dailyStats.map(day => day.totalRevenue), 7),
    conversion7day: calculateMovingAverage_(dailyStats.map(day => day.conversionRate), 7)
  };
  
  return {
    trends: trends,
    weeklyComparison: weeklyComparison,
    movingAverages: movingAverages
  };
}

/**
 * Анализирует паттерны по дням недели
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Анализ по дням недели
 * @private
 */
function analyzeWeekdayPatterns_(headers, rows) {
  const weekdayStats = {};
  const dayNames = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
  
  // Инициализируем статистику для каждого дня недели
  dayNames.forEach(dayName => {
    weekdayStats[dayName] = {
      dayName: dayName,
      totalLeads: 0,
      successfulDeals: 0,
      totalRevenue: 0,
      avgLeadsPerDay: 0,
      conversionRate: 0,
      daysCount: 0,
      channels: {}
    };
  });
  
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  
  // Подсчёт дней для каждого дня недели
  const uniqueDates = new Set();
  
  rows.forEach(row => {
    const createdDate = parseDate_(row[createdIndex]);
    if (!createdDate) return;
    
    const dayName = getDayName_(createdDate.getDay());
    const dateKey = formatDate_(createdDate, 'YYYY-MM-DD');
    uniqueDates.add(`${dayName}-${dateKey}`);
    
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const channel = row[channelIndex] || 'Неизвестно';
    
    const dayStats = weekdayStats[dayName];
    dayStats.totalLeads++;
    
    if (status === 'success') {
      dayStats.successfulDeals++;
      dayStats.totalRevenue += budget;
    }
    
    dayStats.channels[channel] = (dayStats.channels[channel] || 0) + 1;
  });
  
  // Подсчитываем количество уникальных дней для каждого дня недели
  Array.from(uniqueDates).forEach(dateEntry => {
    const dayName = dateEntry.split('-')[0];
    weekdayStats[dayName].daysCount++;
  });
  
  // Вычисляем средние значения
  return Object.values(weekdayStats).map(dayStats => {
    dayStats.avgLeadsPerDay = dayStats.daysCount > 0 ? 
      Math.round(dayStats.totalLeads / dayStats.daysCount) : 0;
    dayStats.conversionRate = dayStats.totalLeads > 0 ? 
      (dayStats.successfulDeals / dayStats.totalLeads * 100) : 0;
    
    // Топ канал дня недели
    dayStats.topChannel = Object.entries(dayStats.channels)
      .sort(([,a], [,b]) => b - a)[0]?.[0] || 'Нет данных';
    
    return dayStats;
  }).sort((a, b) => {
    const order = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'];
    return order.indexOf(a.dayName) - order.indexOf(b.dayName);
  });
}

/**
 * Анализирует почасовые паттерны
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Почасовая активность
 * @private
 */
function analyzeHourlyPatterns_(headers, rows) {
  const hourlyStats = {};
  
  // Инициализируем статистику для каждого часа
  for (let hour = 0; hour < 24; hour++) {
    hourlyStats[hour] = {
      hour: hour,
      displayHour: `${hour.toString().padStart(2, '0')}:00`,
      totalLeads: 0,
      successfulDeals: 0,
      conversionRate: 0,
      channels: {}
    };
  }
  
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  
  rows.forEach(row => {
    const createdDate = parseDate_(row[createdIndex]);
    if (!createdDate) return;
    
    const hour = createdDate.getHours();
    const status = row[statusIndex] || '';
    const channel = row[channelIndex] || 'Неизвестно';
    
    const hourStats = hourlyStats[hour];
    hourStats.totalLeads++;
    
    if (status === 'success') {
      hourStats.successfulDeals++;
    }
    
    hourStats.channels[channel] = (hourStats.channels[channel] || 0) + 1;
  });
  
  // Вычисляем конверсию и находим топ каналы
  return Object.values(hourlyStats).map(hourStats => {
    hourStats.conversionRate = hourStats.totalLeads > 0 ? 
      (hourStats.successfulDeals / hourStats.totalLeads * 100) : 0;
    
    hourStats.topChannel = Object.entries(hourStats.channels)
      .sort(([,a], [,b]) => b - a)[0]?.[0] || 'Нет данных';
    
    return hourStats;
  });
}

/**
 * Обнаруживает аномалии в ежедневных данных
 * @param {Array} dailyStats - Ежедневная статистика
 * @returns {Array} Обнаруженные аномалии
 * @private
 */
function detectDailyAnomalies_(dailyStats) {
  if (dailyStats.length < 7) return [];
  
  const anomalies = [];
  
  // Вычисляем средние и стандартные отклонения
  const leads = dailyStats.map(day => day.totalLeads);
  const revenue = dailyStats.map(day => day.totalRevenue);
  
  const leadsAvg = leads.reduce((a, b) => a + b, 0) / leads.length;
  const revenueAvg = revenue.reduce((a, b) => a + b, 0) / revenue.length;
  
  const leadsStd = Math.sqrt(leads.reduce((sum, val) => sum + Math.pow(val - leadsAvg, 2), 0) / leads.length);
  const revenueStd = Math.sqrt(revenue.reduce((sum, val) => sum + Math.pow(val - revenueAvg, 2), 0) / revenue.length);
  
  // Ищем аномалии (значения за пределами 2 стандартных отклонений)
  dailyStats.forEach(dayStats => {
    const leadsZScore = Math.abs((dayStats.totalLeads - leadsAvg) / leadsStd);
    const revenueZScore = Math.abs((dayStats.totalRevenue - revenueAvg) / revenueStd);
    
    if (leadsZScore > 2) {
      anomalies.push({
        date: dayStats.date,
        type: dayStats.totalLeads > leadsAvg ? 'Пиковая активность' : 'Низкая активность',
        metric: 'Лиды',
        value: dayStats.totalLeads,
        deviation: leadsZScore.toFixed(2),
        description: `${dayStats.totalLeads} лидов (среднее: ${Math.round(leadsAvg)})`
      });
    }
    
    if (revenueZScore > 2 && dayStats.totalRevenue > 0) {
      anomalies.push({
        date: dayStats.date,
        type: dayStats.totalRevenue > revenueAvg ? 'Высокая выручка' : 'Низкая выручка',
        metric: 'Выручка',
        value: dayStats.totalRevenue,
        deviation: revenueZScore.toFixed(2),
        description: `${formatCurrency_(dayStats.totalRevenue)} (среднее: ${formatCurrency_(revenueAvg)})`
      });
    }
    
    // Аномалии конверсии
    if (dayStats.totalLeads >= 5 && (dayStats.conversionRate === 0 || dayStats.conversionRate >= 80)) {
      anomalies.push({
        date: dayStats.date,
        type: dayStats.conversionRate === 0 ? 'Нулевая конверсия' : 'Очень высокая конверсия',
        metric: 'Конверсия',
        value: dayStats.conversionRate,
        deviation: 'Н/Д',
        description: `${dayStats.conversionRate.toFixed(1)}% при ${dayStats.totalLeads} лидах`
      });
    }
  });
  
  return anomalies.sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 10);
}

/**
 * Сравнивает текущий период с предыдущими
 * @param {Array} dailyStats - Ежедневная статистика
 * @returns {Object} Сравнение периодов
 * @private
 */
function comparePeriods_(dailyStats) {
  const comparison = {};
  const totalDays = dailyStats.length;
  
  if (totalDays >= 30) {
    // Последние 30 дней vs предыдущие 30 дней
    const last30 = dailyStats.slice(-30);
    const prev30 = dailyStats.slice(-60, -30);
    
    comparison.monthly = {
      period: 'Месяц к месяцу',
      current: calculatePeriodStats_(last30),
      previous: calculatePeriodStats_(prev30),
      changes: {}
    };
    
    comparison.monthly.changes = {
      leads: calculateGrowthRate_(comparison.monthly.previous.totalLeads, comparison.monthly.current.totalLeads),
      revenue: calculateGrowthRate_(comparison.monthly.previous.totalRevenue, comparison.monthly.current.totalRevenue),
      conversions: calculateGrowthRate_(comparison.monthly.previous.successfulDeals, comparison.monthly.current.successfulDeals),
      avgDealValue: calculateGrowthRate_(comparison.monthly.previous.avgDealValue, comparison.monthly.current.avgDealValue)
    };
  }
  
  if (totalDays >= 14) {
    // Последние 7 дней vs предыдущие 7 дней
    const last7 = dailyStats.slice(-7);
    const prev7 = dailyStats.slice(-14, -7);
    
    comparison.weekly = {
      period: 'Неделя к неделе',
      current: calculatePeriodStats_(last7),
      previous: calculatePeriodStats_(prev7),
      changes: {}
    };
    
    comparison.weekly.changes = {
      leads: calculateGrowthRate_(comparison.weekly.previous.totalLeads, comparison.weekly.current.totalLeads),
      revenue: calculateGrowthRate_(comparison.weekly.previous.totalRevenue, comparison.weekly.current.totalRevenue),
      conversions: calculateGrowthRate_(comparison.weekly.previous.successfulDeals, comparison.weekly.current.successfulDeals),
      avgDealValue: calculateGrowthRate_(comparison.weekly.previous.avgDealValue, comparison.weekly.current.avgDealValue)
    };
  }
  
  return comparison;
}

/**
 * Вычисляет статистику за период
 * @param {Array} periodStats - Статистика за период
 * @returns {Object} Статистика периода
 * @private
 */
function calculatePeriodStats_(periodStats) {
  const totalLeads = periodStats.reduce((sum, day) => sum + day.totalLeads, 0);
  const successfulDeals = periodStats.reduce((sum, day) => sum + day.successfulDeals, 0);
  const totalRevenue = periodStats.reduce((sum, day) => sum + day.totalRevenue, 0);
  
  return {
    totalLeads: totalLeads,
    successfulDeals: successfulDeals,
    totalRevenue: totalRevenue,
    avgDealValue: successfulDeals > 0 ? totalRevenue / successfulDeals : 0,
    conversionRate: totalLeads > 0 ? (successfulDeals / totalLeads * 100) : 0,
    avgLeadsPerDay: Math.round(totalLeads / periodStats.length),
    avgRevenuePerDay: Math.round(totalRevenue / periodStats.length)
  };
}

/**
 * Создаёт структуру отчёта ежедневной статистики
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} dailyData - Данные ежедневной статистики
 * @private
 */
function createDailyStatisticsStructure_(sheet, dailyData) {
  let currentRow = 3;
  
  // 1. Общая статистика
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('📊 ОБЩАЯ СТАТИСТИКА ПО ДНЯМ');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const totalStats = calculatePeriodStats_(dailyData.dailyStats);
  const overallStats = [
    ['Анализируемый период:', `${dailyData.totalDays} дней`],
    ['Всего лидов:', totalStats.totalLeads],
    ['Средне лидов в день:', totalStats.avgLeadsPerDay],
    ['Успешных сделок:', totalStats.successfulDeals],
    ['Общая выручка:', formatCurrency_(totalStats.totalRevenue)],
    ['Средняя выручка в день:', formatCurrency_(totalStats.avgRevenuePerDay)],
    ['Общая конверсия:', `${totalStats.conversionRate.toFixed(1)}%`]
  ];
  
  sheet.getRange(currentRow, 1, overallStats.length, 2).setValues(overallStats);
  sheet.getRange(currentRow, 1, overallStats.length, 1).setFontWeight('bold');
  currentRow += overallStats.length + 2;
  
  // 2. Сравнение периодов
  if (dailyData.periodComparison.weekly) {
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue('📈 СРАВНЕНИЕ С ПРЕДЫДУЩИМИ ПЕРИОДАМИ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
    currentRow++;
    
    const comparisonHeaders = [['Метрика', 'Текущая неделя', 'Предыдущая неделя', 'Изменение', 'Тренд']];
    sheet.getRange(currentRow, 1, 1, 5).setValues(comparisonHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
    currentRow++;
    
    const weeklyComp = dailyData.periodComparison.weekly;
    const comparisonData = [
      ['Лиды', weeklyComp.current.totalLeads, weeklyComp.previous.totalLeads, 
       `${weeklyComp.changes.leads > 0 ? '+' : ''}${weeklyComp.changes.leads.toFixed(1)}%`,
       weeklyComp.changes.leads > 0 ? '📈' : weeklyComp.changes.leads < 0 ? '📉' : '➡️'],
      ['Конверсии', weeklyComp.current.successfulDeals, weeklyComp.previous.successfulDeals,
       `${weeklyComp.changes.conversions > 0 ? '+' : ''}${weeklyComp.changes.conversions.toFixed(1)}%`,
       weeklyComp.changes.conversions > 0 ? '📈' : weeklyComp.changes.conversions < 0 ? '📉' : '➡️'],
      ['Выручка', formatCurrency_(weeklyComp.current.totalRevenue), formatCurrency_(weeklyComp.previous.totalRevenue),
       `${weeklyComp.changes.revenue > 0 ? '+' : ''}${weeklyComp.changes.revenue.toFixed(1)}%`,
       weeklyComp.changes.revenue > 0 ? '📈' : weeklyComp.changes.revenue < 0 ? '📉' : '➡️']
    ];
    
    sheet.getRange(currentRow, 1, comparisonData.length, 5).setValues(comparisonData);
    
    // Цветовое выделение трендов
    comparisonData.forEach((row, index) => {
      const changeCell = sheet.getRange(currentRow + index, 4);
      const trendCell = sheet.getRange(currentRow + index, 5);
      
      if (row[3].includes('+')) {
        changeCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
        trendCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
      } else if (parseFloat(row[3]) < 0) {
        changeCell.setBackground(CONFIG.COLORS.ERROR_LIGHT);
        trendCell.setBackground(CONFIG.COLORS.ERROR_LIGHT);
      }
    });
    
    currentRow += comparisonData.length + 2;
  }
  
  // 3. Последние дни (детально)
  sheet.getRange(currentRow, 1, 1, 8).merge();
  sheet.getRange(currentRow, 1).setValue('📅 ДЕТАЛИЗАЦИЯ ПО ДНЯМ (последние 14 дней)');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 8));
  currentRow++;
  
  const dailyHeaders = [['Дата', 'День недели', 'Лиды', 'Сделки', 'Выручка', 'Конверсия', 'Топ канал', 'Пик часа']];
  sheet.getRange(currentRow, 1, 1, 8).setValues(dailyHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 8));
  currentRow++;
  
  const recentDays = dailyData.dailyStats.slice(-14);
  const dailyDetailsData = recentDays.map(day => [
    formatDate_(day.dateObj, 'DD.MM.YYYY'),
    day.dayOfWeek,
    day.totalLeads,
    day.successfulDeals,
    formatCurrency_(day.totalRevenue),
    `${day.conversionRate.toFixed(1)}%`,
    day.topChannel,
    day.peakHour !== 'Нет данных' ? `${day.peakHour}:00` : 'Н/Д'
  ]);
  
  if (dailyDetailsData.length > 0) {
    sheet.getRange(currentRow, 1, dailyDetailsData.length, 8).setValues(dailyDetailsData);
    currentRow += dailyDetailsData.length;
  }
  currentRow += 2;
  
  // 4. Анализ по дням недели
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.getRange(currentRow, 1).setValue('📊 АКТИВНОСТЬ ПО ДНЯМ НЕДЕЛИ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const weekdayHeaders = [['День недели', 'Всего лидов', 'Среднее в день', 'Конверсия', 'Выручка', 'Топ канал']];
  sheet.getRange(currentRow, 1, 1, 6).setValues(weekdayHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const weekdayData = dailyData.weekdayAnalysis.map(day => [
    day.dayName,
    day.totalLeads,
    day.avgLeadsPerDay,
    `${day.conversionRate.toFixed(1)}%`,
    formatCurrency_(day.totalRevenue),
    day.topChannel
  ]);
  
  if (weekdayData.length > 0) {
    sheet.getRange(currentRow, 1, weekdayData.length, 6).setValues(weekdayData);
    currentRow += weekdayData.length;
  }
  currentRow += 2;
  
  // 5. Почасовая активность (топ часы)
  sheet.getRange(currentRow, 1, 1, 4).merge();
  sheet.getRange(currentRow, 1).setValue('⏰ ПОЧАСОВАЯ АКТИВНОСТЬ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
  currentRow++;
  
  const hourlyHeaders = [['Время', 'Лиды', 'Конверсия', 'Топ канал']];
  sheet.getRange(currentRow, 1, 1, 4).setValues(hourlyHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
  currentRow++;
  
  // Топ 10 самых активных часов
  const topHours = dailyData.hourlyAnalysis
    .filter(hour => hour.totalLeads > 0)
    .sort((a, b) => b.totalLeads - a.totalLeads)
    .slice(0, 10);
  
  const hourlyData = topHours.map(hour => [
    hour.displayHour,
    hour.totalLeads,
    `${hour.conversionRate.toFixed(1)}%`,
    hour.topChannel
  ]);
  
  if (hourlyData.length > 0) {
    sheet.getRange(currentRow, 1, hourlyData.length, 4).setValues(hourlyData);
    currentRow += hourlyData.length;
  }
  currentRow += 2;
  
  // 6. Аномалии и выбросы
  if (dailyData.anomalies.length > 0) {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue('🚨 ОБНАРУЖЕННЫЕ АНОМАЛИИ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const anomalyHeaders = [['Дата', 'Тип аномалии', 'Метрика', 'Описание']];
    sheet.getRange(currentRow, 1, 1, 4).setValues(anomalyHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const anomalyData = dailyData.anomalies.slice(0, 8).map(anomaly => [
      formatDate_(new Date(anomaly.date), 'DD.MM.YYYY'),
      anomaly.type,
      anomaly.metric,
      anomaly.description
    ]);
    
    sheet.getRange(currentRow, 1, anomalyData.length, 4).setValues(anomalyData);
    
    // Цветовое выделение типов аномалий
    anomalyData.forEach((row, index) => {
      const typeCell = sheet.getRange(currentRow + index, 2);
      if (row[1].includes('Пиковая') || row[1].includes('Высокая')) {
        typeCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
      } else if (row[1].includes('Низкая') || row[1].includes('Нулевая')) {
        typeCell.setBackground(CONFIG.COLORS.WARNING_LIGHT);
      }
    });
    
    currentRow += anomalyData.length;
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 8);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 8);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к ежедневной статистике
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} dailyData - Данные ежедневной статистики
 * @private
 */
function addDailyStatisticsCharts_(sheet, dailyData) {
  // 1. Динамика лидов по дням
  if (dailyData.dailyStats.length > 0) {
    const dailyChartData = [['Дата', 'Лиды', 'Сделки']].concat(
      dailyData.dailyStats.slice(-30).map(day => [
        formatDate_(day.dateObj, 'DD.MM'),
        day.totalLeads,
        day.successfulDeals
      ])
    );
    
    // Создаем диаграмму через универсальную функцию
    const dailyChart = createChart_(sheet, 'line', dailyChartData, {
      startRow: 1,
      startCol: 10,
      title: 'Динамика лидов и сделок (30 дней)',
      position: { row: 3, col: 10 },
      width: 700,
      height: 400,
      legend: 'top',
      hAxisTitle: 'Дата',
      vAxisTitle: 'Количество'
    });
    
    sheet.getRange(1, 10, dailyChartData.length, 3).setValues(dailyChartData);
  }
  
  // 2. Столбчатая диаграмма по дням недели
  if (dailyData.weekdayAnalysis.length > 0) {
    const weekdayChartData = [['День недели', 'Среднее лидов в день']].concat(
      dailyData.weekdayAnalysis.map(day => [day.dayName, day.avgLeadsPerDay])
    );
    
    // Создаем диаграмму через универсальную функцию
    const weekdayChart = createChart_(sheet, 'column', weekdayChartData, {
      startRow: 1,
      startCol: 14,
      title: 'Активность по дням недели',
      position: { row: 3, col: 18 },
      width: 500,
      height: 350,
      legend: 'none',
      hAxisTitle: 'День недели',
      vAxisTitle: 'Среднее количество лидов'
    });
  }
  
  // 3. Почасовая активность
  if (dailyData.hourlyAnalysis.length > 0) {
    const hourlyChartData = [['Час', 'Лиды']].concat(
      dailyData.hourlyAnalysis.map(hour => [hour.displayHour, hour.totalLeads])
    );
    
    // Создаем диаграмму через универсальную функцию
    const hourlyChart = createChart_(sheet, 'line', hourlyChartData, {
      startRow: 1,
      startCol: 17,
      title: 'Почасовая активность',
      position: { row: 25, col: 10 },
      width: 700,
      height: 350,
      legend: 'none',
      hAxisTitle: 'Час дня',
      vAxisTitle: 'Количество лидов'
    });
  }
  
  // 4. Диаграмма выручки по дням
  if (dailyData.dailyStats.length > 0) {
    const revenueChartData = [['Дата', 'Выручка']].concat(
      dailyData.dailyStats.slice(-30)
        .filter(day => day.totalRevenue > 0)
        .map(day => [formatDate_(day.dateObj, 'DD.MM'), day.totalRevenue])
    );
    
    if (revenueChartData.length > 1) {
      // Создаем диаграмму через универсальную функцию
      const revenueChart = createChart_(sheet, 'column', revenueChartData, {
        startRow: 1,
        startCol: 20,
        title: 'Ежедневная выручка',
        position: { row: 25, col: 18 },
        width: 600,
        height: 350,
        legend: 'none',
        hAxisTitle: 'Дата',
        vAxisTitle: 'Выручка, руб.'
      });
    }
  }
}

/**
 * Вспомогательные функции для расчётов
 */

/**
 * Вычисляет тренд временного ряда
 */
function calculateTrend_(values) {
  if (values.length < 2) return { slope: 0, direction: 'Стабильно' };
  
  const n = values.length;
  const sumX = (n * (n - 1)) / 2;
  const sumY = values.reduce((a, b) => a + b, 0);
  const sumXY = values.reduce((sum, y, x) => sum + x * y, 0);
  const sumXX = (n * (n - 1) * (2 * n - 1)) / 6;
  
  const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
  
  let direction;
  if (Math.abs(slope) < 0.1) direction = 'Стабильно';
  else if (slope > 0) direction = 'Рост';
  else direction = 'Снижение';
  
  return { slope: slope, direction: direction };
}

/**
 * Вычисляет темп роста
 */
function calculateGrowthRate_(previous, current) {
  if (previous === 0) return current > 0 ? 100 : 0;
  return ((current - previous) / previous) * 100;
}

/**
 * Вычисляет скользящее среднее
 */
function calculateMovingAverage_(values, period) {
  if (values.length < period) return values;
  
  const result = [];
  for (let i = period - 1; i < values.length; i++) {
    const sum = values.slice(i - period + 1, i + 1).reduce((a, b) => a + b, 0);
    result.push(sum / period);
  }
  return result;
}

/**
 * Получает название дня недели
 */
function getDayName_(dayNumber) {
  const days = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
  return days[dayNumber];
}

/**
 * Создаёт пустой отчёт когда нет данных
 * @private
 */
function createEmptyDailyReport_() {
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.DAILY_STATISTICS);
  clearSheetData_(sheet);
  applySheetFormatting_(sheet, 'Ежедневная статистика');
  
  sheet.getRange('A3').setValue('ℹ️ Нет данных для ежедневной статистики');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A5').setValue('Убедитесь, что:');
  sheet.getRange('A6').setValue('• В системе есть данные с корректными датами');
  sheet.getRange('A7').setValue('• Данные охватывают несколько дней');
  sheet.getRange('A8').setValue('• Выполнена синхронизация данных');
  
  updateLastUpdateTime_(CONFIG.SHEETS.DAILY_STATISTICS);
}
