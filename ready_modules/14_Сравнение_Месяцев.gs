/**
 * 📅 Сравнение Месяцев - Сравнительный анализ
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Apps Script: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit
 * 2. Создайте новый файл: + → Script file
 * 3. Назовите файл: 14_Сравнение_Месяцев (без .gs)
 * 4. Удалите весь код по умолчанию
 * 5. Скопируйте и вставьте код ниже
 * 6. Сохраните файл (Ctrl+S)
 */

/**
 * СРАВНЕНИЕ ПО МЕСЯЦАМ
 * Детальный анализ динамики показателей по месяцам с трендами и прогнозами
 * @fileoverview Модуль сравнительного анализа производительности по месяцам
 */

/**
 * Основная функция создания месячного сравнения
 * Создаёт отчёт с детальным сравнением показателей по месяцам
 */
function generateMonthlyComparison() {
  try {
    logInfo_('MONTHLY_COMPARISON', 'Начало генерации месячного сравнения');
    
    const startTime = new Date();
    
    // Получаем данные для месячного анализа
    const monthlyData = getMonthlyComparisonData_();
    
    if (!monthlyData || monthlyData.monthlyStats.length === 0) {
      logWarning_('MONTHLY_COMPARISON', 'Нет данных для месячного сравнения');
      createEmptyMonthlyReport_();
      return;
    }
    
    // Создаём или обновляем лист
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.MONTHLY_COMPARISON);
    clearSheetData_(sheet);
    applySheetFormatting_(sheet, 'Сравнение по месяцам');
    
    // Строим структуру отчёта
    createMonthlyComparisonStructure_(sheet, monthlyData);
    
    // Добавляем диаграммы
    addMonthlyComparisonCharts_(sheet, monthlyData);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('MONTHLY_COMPARISON', `Месячное сравнение сгенерировано за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.MONTHLY_COMPARISON);
    
  } catch (error) {
    logError_('MONTHLY_COMPARISON', 'Ошибка генерации месячного сравнения', error);
    throw error;
  }
}

/**
 * Получает данные для месячного сравнения
 * @returns {Object} Данные месячного сравнения
 * @private
 */
function getMonthlyComparisonData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return null;
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Основная месячная статистика
  const monthlyStats = generateMonthlyStats_(headers, rows);
  
  // Сравнительный анализ месяцев
  const comparison = compareMonthlyPerformance_(monthlyStats);
  
  // Трендовый анализ
  const trendAnalysis = analyzeMonthlyTrends_(monthlyStats);
  
  // Сезонный анализ
  const seasonalAnalysis = analyzeSeasonalPatterns_(monthlyStats);
  
  // Анализ по каналам помесячно
  const channelAnalysis = analyzeMonthlyChannels_(headers, rows);
  
  // Прогнозирование
  const forecasting = generateMonthlyForecasts_(monthlyStats);
  
  // Аномалии и выбросы
  const anomalies = detectMonthlyAnomalies_(monthlyStats);
  
  return {
    monthlyStats: monthlyStats,
    comparison: comparison,
    trendAnalysis: trendAnalysis,
    seasonalAnalysis: seasonalAnalysis,
    channelAnalysis: channelAnalysis,
    forecasting: forecasting,
    anomalies: anomalies,
    totalMonths: monthlyStats.length,
    analysisDate: getCurrentDateMoscow_()
  };
}

/**
 * Генерирует месячную статистику
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Месячная статистика
 * @private
 */
function generateMonthlyStats_(headers, rows) {
  const monthlyStats = {};
  
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const closedIndex = findHeaderIndex_(headers, 'Дата закрытия');
  
  rows.forEach(row => {
    const createdDate = parseDate_(row[createdIndex]);
    if (!createdDate) return;
    
    const monthKey = formatDate_(createdDate, 'YYYY-MM');
    const monthName = `${getMonthName_(createdDate.getMonth())} ${createdDate.getFullYear()}`;
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const channel = row[channelIndex] || 'Неизвестно';
    const source = row[sourceIndex] || 'direct';
    const closedDate = parseDate_(row[closedIndex]);
    
    if (!monthlyStats[monthKey]) {
      monthlyStats[monthKey] = {
        monthKey: monthKey,
        monthName: monthName,
        year: createdDate.getFullYear(),
        month: createdDate.getMonth() + 1,
        quarter: Math.floor(createdDate.getMonth() / 3) + 1,
        totalLeads: 0,
        successfulDeals: 0,
        failedDeals: 0,
        inProgressDeals: 0,
        totalRevenue: 0,
        avgDealValue: 0,
        conversionRate: 0,
        avgDealCycle: 0,
        channels: {},
        sources: {},
        dealCycles: [],
        workingDays: 0,
        leadsPerWorkingDay: 0,
        revenuePerWorkingDay: 0
      };
    }
    
    const monthStats = monthlyStats[monthKey];
    monthStats.totalLeads++;
    
    // Анализ статусов
    if (status === 'success') {
      monthStats.successfulDeals++;
      monthStats.totalRevenue += budget;
      
      // Время закрытия сделки
      if (closedDate && closedDate >= createdDate) {
        const dealCycle = Math.round((closedDate - createdDate) / (1000 * 60 * 60 * 24));
        monthStats.dealCycles.push(dealCycle);
      }
    } else if (status === 'failure') {
      monthStats.failedDeals++;
    } else {
      monthStats.inProgressDeals++;
    }
    
    // Каналы и источники
    monthStats.channels[channel] = (monthStats.channels[channel] || 0) + 1;
    monthStats.sources[source] = (monthStats.sources[source] || 0) + 1;
  });
  
  // Вычисляем производные метрики и рабочие дни
  return Object.values(monthlyStats)
    .map(monthStats => {
      // Основные метрики
      monthStats.conversionRate = monthStats.totalLeads > 0 ? 
        (monthStats.successfulDeals / monthStats.totalLeads * 100) : 0;
      monthStats.avgDealValue = monthStats.successfulDeals > 0 ? 
        (monthStats.totalRevenue / monthStats.successfulDeals) : 0;
      monthStats.avgDealCycle = monthStats.dealCycles.length > 0 ?
        Math.round(monthStats.dealCycles.reduce((a, b) => a + b, 0) / monthStats.dealCycles.length) : 0;
      
      // Рабочие дни в месяце (исключаем выходные)
      monthStats.workingDays = getWorkingDaysInMonth_(monthStats.year, monthStats.month - 1);
      monthStats.leadsPerWorkingDay = monthStats.workingDays > 0 ? 
        Math.round(monthStats.totalLeads / monthStats.workingDays * 10) / 10 : 0;
      monthStats.revenuePerWorkingDay = monthStats.workingDays > 0 ? 
        Math.round(monthStats.totalRevenue / monthStats.workingDays) : 0;
      
      // Топ канал и источник месяца
      monthStats.topChannel = Object.entries(monthStats.channels)
        .sort(([,a], [,b]) => b - a)[0]?.[0] || 'Нет данных';
      monthStats.topSource = Object.entries(monthStats.sources)
        .sort(([,a], [,b]) => b - a)[0]?.[0] || 'Нет данных';
      
      return monthStats;
    })
    .sort((a, b) => a.monthKey.localeCompare(b.monthKey));
}

/**
 * Сравнивает производительность по месяцам
 * @param {Array} monthlyStats - Месячная статистика
 * @returns {Object} Сравнительный анализ
 * @private
 */
function compareMonthlyPerformance_(monthlyStats) {
  if (monthlyStats.length < 2) return { monthToMonth: [], bestMonths: {}, worstMonths: {} };
  
  const comparison = {
    monthToMonth: [],
    bestMonths: {},
    worstMonths: {},
    yearOverYear: []
  };
  
  // Сравнение месяц к месяцу
  for (let i = 1; i < monthlyStats.length; i++) {
    const current = monthlyStats[i];
    const previous = monthlyStats[i - 1];
    
    comparison.monthToMonth.push({
      currentMonth: current.monthName,
      previousMonth: previous.monthName,
      leadsChange: calculateGrowthRate_(previous.totalLeads, current.totalLeads),
      revenueChange: calculateGrowthRate_(previous.totalRevenue, current.totalRevenue),
      conversionChange: previous.conversionRate > 0 ? 
        ((current.conversionRate - previous.conversionRate) / previous.conversionRate * 100) : 0,
      dealValueChange: calculateGrowthRate_(previous.avgDealValue, current.avgDealValue)
    });
  }
  
  // Лучшие и худшие месяцы
  const sortedByRevenue = [...monthlyStats].sort((a, b) => b.totalRevenue - a.totalRevenue);
  const sortedByLeads = [...monthlyStats].sort((a, b) => b.totalLeads - a.totalLeads);
  const sortedByConversion = [...monthlyStats].sort((a, b) => b.conversionRate - a.conversionRate);
  
  comparison.bestMonths = {
    byRevenue: sortedByRevenue[0],
    byLeads: sortedByLeads[0],
    byConversion: sortedByConversion[0]
  };
  
  comparison.worstMonths = {
    byRevenue: sortedByRevenue[sortedByRevenue.length - 1],
    byLeads: sortedByLeads[sortedByLeads.length - 1],
    byConversion: sortedByConversion[sortedByConversion.length - 1]
  };
  
  // Сравнение год к году (если есть данные за несколько лет)
  const yearGroups = {};
  monthlyStats.forEach(month => {
    const monthNum = month.month;
    if (!yearGroups[monthNum]) yearGroups[monthNum] = [];
    yearGroups[monthNum].push(month);
  });
  
  Object.keys(yearGroups).forEach(monthNum => {
    const monthsData = yearGroups[monthNum].sort((a, b) => a.year - b.year);
    if (monthsData.length >= 2) {
      for (let i = 1; i < monthsData.length; i++) {
        comparison.yearOverYear.push({
          month: getMonthName_(parseInt(monthNum) - 1),
          currentYear: monthsData[i].year,
          previousYear: monthsData[i - 1].year,
          leadsChange: calculateGrowthRate_(monthsData[i - 1].totalLeads, monthsData[i].totalLeads),
          revenueChange: calculateGrowthRate_(monthsData[i - 1].totalRevenue, monthsData[i].totalRevenue)
        });
      }
    }
  });
  
  return comparison;
}

/**
 * Анализирует месячные тренды
 * @param {Array} monthlyStats - Месячная статистика
 * @returns {Object} Трендовый анализ
 * @private
 */
function analyzeMonthlyTrends_(monthlyStats) {
  if (monthlyStats.length < 3) return { trends: {}, movingAverages: {}, acceleration: {} };
  
  const leads = monthlyStats.map(m => m.totalLeads);
  const revenue = monthlyStats.map(m => m.totalRevenue);
  const conversion = monthlyStats.map(m => m.conversionRate);
  const dealValue = monthlyStats.map(m => m.avgDealValue);
  
  const trends = {
    leads: calculateTrendDirection_(leads),
    revenue: calculateTrendDirection_(revenue),
    conversion: calculateTrendDirection_(conversion),
    dealValue: calculateTrendDirection_(dealValue)
  };
  
  // Скользящие средние (3 месяца)
  const movingAverages = {
    leads: calculateMovingAverage_(leads, 3),
    revenue: calculateMovingAverage_(revenue, 3),
    conversion: calculateMovingAverage_(conversion, 3)
  };
  
  // Ускорение/замедление роста
  const acceleration = {};
  if (monthlyStats.length >= 6) {
    const recentGrowth = calculateAverageGrowth_(revenue.slice(-3));
    const earlierGrowth = calculateAverageGrowth_(revenue.slice(-6, -3));
    
    acceleration.revenue = {
      recent: recentGrowth,
      earlier: earlierGrowth,
      isAccelerating: recentGrowth > earlierGrowth,
      change: recentGrowth - earlierGrowth
    };
  }
  
  return { trends, movingAverages, acceleration };
}

/**
 * Анализирует сезонные паттерны
 * @param {Array} monthlyStats - Месячная статистика
 * @returns {Object} Сезонный анализ
 * @private
 */
function analyzeSeasonalPatterns_(monthlyStats) {
  const seasonalAnalysis = {
    quarters: {
      'Q1': { months: [], avgLeads: 0, avgRevenue: 0, avgConversion: 0 },
      'Q2': { months: [], avgLeads: 0, avgRevenue: 0, avgConversion: 0 },
      'Q3': { months: [], avgLeads: 0, avgRevenue: 0, avgConversion: 0 },
      'Q4': { months: [], avgLeads: 0, avgRevenue: 0, avgConversion: 0 }
    },
    monthlyAverages: {},
    peakMonths: [],
    lowMonths: []
  };
  
  // Группировка по кварталам
  monthlyStats.forEach(month => {
    const quarterKey = `Q${month.quarter}`;
    seasonalAnalysis.quarters[quarterKey].months.push(month);
  });
  
  // Вычисление средних по кварталам
  Object.keys(seasonalAnalysis.quarters).forEach(quarter => {
    const months = seasonalAnalysis.quarters[quarter].months;
    if (months.length > 0) {
      seasonalAnalysis.quarters[quarter].avgLeads = 
        Math.round(months.reduce((sum, m) => sum + m.totalLeads, 0) / months.length);
      seasonalAnalysis.quarters[quarter].avgRevenue = 
        Math.round(months.reduce((sum, m) => sum + m.totalRevenue, 0) / months.length);
      seasonalAnalysis.quarters[quarter].avgConversion = 
        Math.round(months.reduce((sum, m) => sum + m.conversionRate, 0) / months.length * 10) / 10;
    }
  });
  
  // Средние по месяцам года (1-12)
  for (let monthNum = 1; monthNum <= 12; monthNum++) {
    const monthsData = monthlyStats.filter(m => m.month === monthNum);
    if (monthsData.length > 0) {
      seasonalAnalysis.monthlyAverages[monthNum] = {
        monthName: getMonthName_(monthNum - 1),
        avgLeads: Math.round(monthsData.reduce((sum, m) => sum + m.totalLeads, 0) / monthsData.length),
        avgRevenue: Math.round(monthsData.reduce((sum, m) => sum + m.totalRevenue, 0) / monthsData.length),
        years: monthsData.length
      };
    }
  }
  
  // Пиковые и слабые месяцы
  const sortedByPerformance = Object.values(seasonalAnalysis.monthlyAverages)
    .sort((a, b) => b.avgRevenue - a.avgRevenue);
  
  seasonalAnalysis.peakMonths = sortedByPerformance.slice(0, 3);
  seasonalAnalysis.lowMonths = sortedByPerformance.slice(-3).reverse();
  
  return seasonalAnalysis;
}

/**
 * Анализирует каналы по месяцам
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Анализ каналов по месяцам
 * @private
 */
function analyzeMonthlyChannels_(headers, rows) {
  const channelAnalysis = {};
  
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  
  rows.forEach(row => {
    const createdDate = parseDate_(row[createdIndex]);
    if (!createdDate) return;
    
    const monthKey = formatDate_(createdDate, 'YYYY-MM');
    const monthName = `${getMonthName_(createdDate.getMonth())} ${createdDate.getFullYear()}`;
    const channel = row[channelIndex] || 'Неизвестно';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    
    if (!channelAnalysis[monthKey]) {
      channelAnalysis[monthKey] = {
        monthKey: monthKey,
        monthName: monthName,
        channels: {}
      };
    }
    
    if (!channelAnalysis[monthKey].channels[channel]) {
      channelAnalysis[monthKey].channels[channel] = {
        leads: 0,
        conversions: 0,
        revenue: 0,
        conversionRate: 0
      };
    }
    
    const channelStats = channelAnalysis[monthKey].channels[channel];
    channelStats.leads++;
    
    if (status === 'success') {
      channelStats.conversions++;
      channelStats.revenue += budget;
    }
  });
  
  // Вычисляем конверсию для каждого канала по месяцам
  Object.values(channelAnalysis).forEach(monthData => {
    Object.keys(monthData.channels).forEach(channel => {
      const channelStats = monthData.channels[channel];
      channelStats.conversionRate = channelStats.leads > 0 ? 
        (channelStats.conversions / channelStats.leads * 100) : 0;
    });
  });
  
  return Object.values(channelAnalysis);
}

/**
 * Генерирует прогнозы на основе трендов
 * @param {Array} monthlyStats - Месячная статистика
 * @returns {Object} Прогнозы
 * @private
 */
function generateMonthlyForecasts_(monthlyStats) {
  if (monthlyStats.length < 6) return { forecasts: [], reliability: 'low' };
  
  const lastMonths = monthlyStats.slice(-6); // Последние 6 месяцев для прогноза
  
  // Простой линейный тренд
  const leadsTrend = calculateLinearTrend_(lastMonths.map(m => m.totalLeads));
  const revenueTrend = calculateLinearTrend_(lastMonths.map(m => m.totalRevenue));
  
  // Прогноз на следующие 3 месяца
  const forecasts = [];
  const lastMonth = monthlyStats[monthlyStats.length - 1];
  
  for (let i = 1; i <= 3; i++) {
    const forecastDate = new Date(lastMonth.year, lastMonth.month - 1 + i, 1);
    const forecastMonthName = `${getMonthName_(forecastDate.getMonth())} ${forecastDate.getFullYear()}`;
    
    const forecastLeads = Math.round(leadsTrend.slope * (lastMonths.length + i) + leadsTrend.intercept);
    const forecastRevenue = Math.round(revenueTrend.slope * (lastMonths.length + i) + revenueTrend.intercept);
    
    forecasts.push({
      monthName: forecastMonthName,
      forecastLeads: Math.max(0, forecastLeads),
      forecastRevenue: Math.max(0, forecastRevenue),
      confidence: i === 1 ? 'высокая' : i === 2 ? 'средняя' : 'низкая'
    });
  }
  
  // Оценка надёжности прогноза
  const reliability = leadsTrend.r2 > 0.7 ? 'high' : leadsTrend.r2 > 0.4 ? 'medium' : 'low';
  
  return { forecasts, reliability, trendStrength: leadsTrend.r2 };
}

/**
 * Обнаруживает месячные аномалии
 * @param {Array} monthlyStats - Месячная статистика
 * @returns {Array} Аномалии
 * @private
 */
function detectMonthlyAnomalies_(monthlyStats) {
  if (monthlyStats.length < 6) return [];
  
  const anomalies = [];
  
  // Вычисляем средние и стандартные отклонения
  const leads = monthlyStats.map(m => m.totalLeads);
  const revenue = monthlyStats.map(m => m.totalRevenue);
  const conversion = monthlyStats.map(m => m.conversionRate);
  
  const leadsStats = calculateStats_(leads);
  const revenueStats = calculateStats_(revenue);
  const conversionStats = calculateStats_(conversion);
  
  // Ищем аномалии (значения за пределами 1.5 стандартных отклонений)
  monthlyStats.forEach(month => {
    const leadsZScore = Math.abs((month.totalLeads - leadsStats.mean) / leadsStats.std);
    const revenueZScore = Math.abs((month.totalRevenue - revenueStats.mean) / revenueStats.std);
    const conversionZScore = Math.abs((month.conversionRate - conversionStats.mean) / conversionStats.std);
    
    if (leadsZScore > 1.5) {
      anomalies.push({
        month: month.monthName,
        type: month.totalLeads > leadsStats.mean ? 'Высокая активность' : 'Низкая активность',
        metric: 'Лиды',
        value: month.totalLeads,
        expected: Math.round(leadsStats.mean),
        deviation: `${leadsZScore.toFixed(1)}σ`
      });
    }
    
    if (revenueZScore > 1.5 && month.totalRevenue > 0) {
      anomalies.push({
        month: month.monthName,
        type: month.totalRevenue > revenueStats.mean ? 'Высокая выручка' : 'Низкая выручка',
        metric: 'Выручка',
        value: formatCurrency_(month.totalRevenue),
        expected: formatCurrency_(revenueStats.mean),
        deviation: `${revenueZScore.toFixed(1)}σ`
      });
    }
    
    if (conversionZScore > 1.5) {
      anomalies.push({
        month: month.monthName,
        type: month.conversionRate > conversionStats.mean ? 'Высокая конверсия' : 'Низкая конверсия',
        metric: 'Конверсия',
        value: `${month.conversionRate.toFixed(1)}%`,
        expected: `${conversionStats.mean.toFixed(1)}%`,
        deviation: `${conversionZScore.toFixed(1)}σ`
      });
    }
  });
  
  return anomalies.slice(0, 8); // Топ 8 аномалий
}

/**
 * Создаёт структуру отчёта месячного сравнения
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} monthlyData - Данные месячного сравнения
 * @private
 */
function createMonthlyComparisonStructure_(sheet, monthlyData) {
  let currentRow = 3;
  
  // 1. Общая информация о периоде
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('📅 АНАЛИЗИРУЕМЫЙ ПЕРИОД');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const firstMonth = monthlyData.monthlyStats[0];
  const lastMonth = monthlyData.monthlyStats[monthlyData.monthlyStats.length - 1];
  const totalLeads = monthlyData.monthlyStats.reduce((sum, m) => sum + m.totalLeads, 0);
  const totalRevenue = monthlyData.monthlyStats.reduce((sum, m) => sum + m.totalRevenue, 0);
  
  const periodInfo = [
    ['Период анализа:', `${firstMonth.monthName} - ${lastMonth.monthName}`],
    ['Всего месяцев:', monthlyData.totalMonths],
    ['Общее количество лидов:', totalLeads],
    ['Общая выручка:', formatCurrency_(totalRevenue)],
    ['Среднее лидов в месяц:', Math.round(totalLeads / monthlyData.totalMonths)],
    ['Средняя месячная выручка:', formatCurrency_(totalRevenue / monthlyData.totalMonths)]
  ];
  
  sheet.getRange(currentRow, 1, periodInfo.length, 2).setValues(periodInfo);
  sheet.getRange(currentRow, 1, periodInfo.length, 1).setFontWeight('bold');
  currentRow += periodInfo.length + 2;
  
  // 2. Лучшие и худшие месяцы
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.getRange(currentRow, 1).setValue('🏆 ЛУЧШИЕ И ХУДШИЕ РЕЗУЛЬТАТЫ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const bestWorstHeaders = [['Категория', 'Лучший месяц', 'Показатель', 'Худший месяц', 'Показатель', 'Разница']];
  sheet.getRange(currentRow, 1, 1, 6).setValues(bestWorstHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const bestWorstData = [
    [
      'Выручка',
      monthlyData.comparison.bestMonths.byRevenue.monthName,
      formatCurrency_(monthlyData.comparison.bestMonths.byRevenue.totalRevenue),
      monthlyData.comparison.worstMonths.byRevenue.monthName,
      formatCurrency_(monthlyData.comparison.worstMonths.byRevenue.totalRevenue),
      `в ${(monthlyData.comparison.bestMonths.byRevenue.totalRevenue / 
        Math.max(1, monthlyData.comparison.worstMonths.byRevenue.totalRevenue)).toFixed(1)} раз`
    ],
    [
      'Лиды',
      monthlyData.comparison.bestMonths.byLeads.monthName,
      monthlyData.comparison.bestMonths.byLeads.totalLeads,
      monthlyData.comparison.worstMonths.byLeads.monthName,
      monthlyData.comparison.worstMonths.byLeads.totalLeads,
      `в ${(monthlyData.comparison.bestMonths.byLeads.totalLeads / 
        Math.max(1, monthlyData.comparison.worstMonths.byLeads.totalLeads)).toFixed(1)} раз`
    ],
    [
      'Конверсия',
      monthlyData.comparison.bestMonths.byConversion.monthName,
      `${monthlyData.comparison.bestMonths.byConversion.conversionRate.toFixed(1)}%`,
      monthlyData.comparison.worstMonths.byConversion.monthName,
      `${monthlyData.comparison.worstMonths.byConversion.conversionRate.toFixed(1)}%`,
      `+${(monthlyData.comparison.bestMonths.byConversion.conversionRate - 
        monthlyData.comparison.worstMonths.byConversion.conversionRate).toFixed(1)} п.п.`
    ]
  ];
  
  sheet.getRange(currentRow, 1, bestWorstData.length, 6).setValues(bestWorstData);
  currentRow += bestWorstData.length + 2;
  
  // 3. Детальная месячная статистика
  sheet.getRange(currentRow, 1, 1, 9).merge();
  sheet.getRange(currentRow, 1).setValue('📊 ДЕТАЛЬНАЯ СТАТИСТИКА ПО МЕСЯЦАМ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 9));
  currentRow++;
  
  const monthlyHeaders = [['Месяц', 'Лиды', 'Сделки', 'Выручка', 'Конверсия', 'Средний чек', 'Время сделки', 'Лиды/день', 'Топ канал']];
  sheet.getRange(currentRow, 1, 1, 9).setValues(monthlyHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 9));
  currentRow++;
  
  const monthlyDetailsData = monthlyData.monthlyStats.map(month => [
    month.monthName,
    month.totalLeads,
    month.successfulDeals,
    formatCurrency_(month.totalRevenue),
    `${month.conversionRate.toFixed(1)}%`,
    formatCurrency_(month.avgDealValue),
    month.avgDealCycle > 0 ? `${month.avgDealCycle} дн.` : 'Н/Д',
    month.leadsPerWorkingDay,
    month.topChannel
  ]);
  
  if (monthlyDetailsData.length > 0) {
    sheet.getRange(currentRow, 1, monthlyDetailsData.length, 9).setValues(monthlyDetailsData);
    currentRow += monthlyDetailsData.length;
  }
  currentRow += 2;
  
  // 4. Трендовый анализ
  if (monthlyData.trendAnalysis.trends) {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue('📈 ТРЕНДОВЫЙ АНАЛИЗ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const trendHeaders = [['Метрика', 'Тренд', 'Направление', 'Оценка']];
    sheet.getRange(currentRow, 1, 1, 4).setValues(trendHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const trends = monthlyData.trendAnalysis.trends;
    const trendData = [
      ['Лиды', trends.leads.direction, trends.leads.direction === 'Рост' ? '📈' : trends.leads.direction === 'Снижение' ? '📉' : '➡️', 
       trends.leads.direction === 'Рост' ? 'Позитивно' : trends.leads.direction === 'Снижение' ? 'Требует внимания' : 'Стабильно'],
      ['Выручка', trends.revenue.direction, trends.revenue.direction === 'Рост' ? '📈' : trends.revenue.direction === 'Снижение' ? '📉' : '➡️',
       trends.revenue.direction === 'Рост' ? 'Отлично' : trends.revenue.direction === 'Снижение' ? 'Критично' : 'Стабильно'],
      ['Конверсия', trends.conversion.direction, trends.conversion.direction === 'Рост' ? '📈' : trends.conversion.direction === 'Снижение' ? '📉' : '➡️',
       trends.conversion.direction === 'Рост' ? 'Улучшается' : trends.conversion.direction === 'Снижение' ? 'Ухудшается' : 'Без изменений'],
      ['Средний чек', trends.dealValue.direction, trends.dealValue.direction === 'Рост' ? '📈' : trends.dealValue.direction === 'Снижение' ? '📉' : '➡️',
       trends.dealValue.direction === 'Рост' ? 'Растёт качество' : trends.dealValue.direction === 'Снижение' ? 'Снижается качество' : 'Стабильно']
    ];
    
    sheet.getRange(currentRow, 1, trendData.length, 4).setValues(trendData);
    
    // Цветовое выделение трендов
    trendData.forEach((row, index) => {
      const directionCell = sheet.getRange(currentRow + index, 3);
      const evaluationCell = sheet.getRange(currentRow + index, 4);
      
      if (row[2] === '📈') {
        directionCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
        evaluationCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
      } else if (row[2] === '📉') {
        directionCell.setBackground(CONFIG.COLORS.ERROR_LIGHT);
        evaluationCell.setBackground(CONFIG.COLORS.ERROR_LIGHT);
      }
    });
    
    currentRow += trendData.length + 2;
  }
  
  // 5. Сезонный анализ
  if (monthlyData.seasonalAnalysis && monthlyData.seasonalAnalysis.peakMonths.length > 0) {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue('🗓️ СЕЗОННЫЕ ПАТТЕРНЫ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    // Пиковые месяцы
    sheet.getRange(currentRow, 1).setValue('Лучшие месяцы года:');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow++;
    
    monthlyData.seasonalAnalysis.peakMonths.forEach((month, index) => {
      sheet.getRange(currentRow, 1).setValue(`${index + 1}. ${month.monthName}`);
      sheet.getRange(currentRow, 2).setValue(`${month.avgLeads} лидов`);
      sheet.getRange(currentRow, 3).setValue(formatCurrency_(month.avgRevenue));
      currentRow++;
    });
    currentRow++;
    
    // Слабые месяцы
    sheet.getRange(currentRow, 1).setValue('Слабые месяцы года:');
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow++;
    
    monthlyData.seasonalAnalysis.lowMonths.forEach((month, index) => {
      sheet.getRange(currentRow, 1).setValue(`${index + 1}. ${month.monthName}`);
      sheet.getRange(currentRow, 2).setValue(`${month.avgLeads} лидов`);
      sheet.getRange(currentRow, 3).setValue(formatCurrency_(month.avgRevenue));
      currentRow++;
    });
    currentRow += 2;
  }
  
  // 6. Прогнозы
  if (monthlyData.forecasting.forecasts.length > 0) {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue('🔮 ПРОГНОЗ НА БЛИЖАЙШИЕ МЕСЯЦЫ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    sheet.getRange(currentRow, 1).setValue(`Надёжность прогноза: ${
      monthlyData.forecasting.reliability === 'high' ? 'Высокая' :
      monthlyData.forecasting.reliability === 'medium' ? 'Средняя' : 'Низкая'
    }`);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow += 2;
    
    const forecastHeaders = [['Месяц', 'Прогноз лидов', 'Прогноз выручки', 'Уверенность']];
    sheet.getRange(currentRow, 1, 1, 4).setValues(forecastHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const forecastData = monthlyData.forecasting.forecasts.map(forecast => [
      forecast.monthName,
      forecast.forecastLeads,
      formatCurrency_(forecast.forecastRevenue),
      forecast.confidence
    ]);
    
    sheet.getRange(currentRow, 1, forecastData.length, 4).setValues(forecastData);
    currentRow += forecastData.length + 2;
  }
  
  // 7. Аномалии
  if (monthlyData.anomalies.length > 0) {
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue('⚠️ ОБНАРУЖЕННЫЕ АНОМАЛИИ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
    currentRow++;
    
    const anomalyHeaders = [['Месяц', 'Тип аномалии', 'Метрика', 'Значение', 'Отклонение']];
    sheet.getRange(currentRow, 1, 1, 5).setValues(anomalyHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
    currentRow++;
    
    const anomalyData = monthlyData.anomalies.map(anomaly => [
      anomaly.month,
      anomaly.type,
      anomaly.metric,
      typeof anomaly.value === 'string' ? anomaly.value : anomaly.value.toString(),
      anomaly.deviation
    ]);
    
    sheet.getRange(currentRow, 1, anomalyData.length, 5).setValues(anomalyData);
    
    // Цветовое выделение аномалий
    anomalyData.forEach((row, index) => {
      const typeCell = sheet.getRange(currentRow + index, 2);
      if (row[1].includes('Высокая') || row[1].includes('Высокий')) {
        typeCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
      } else if (row[1].includes('Низкая') || row[1].includes('Низкий')) {
        typeCell.setBackground(CONFIG.COLORS.WARNING_LIGHT);
      }
    });
    
    currentRow += anomalyData.length;
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 9);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 9);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к месячному сравнению
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} monthlyData - Данные месячного сравнения
 * @private
 */
function addMonthlyComparisonCharts_(sheet, monthlyData) {
  // 1. Основная динамика лидов и выручки
  if (monthlyData.monthlyStats.length > 0) {
    const dynamicsChartData = [['Месяц', 'Лиды', 'Выручка (тыс. руб.)']].concat(
      monthlyData.monthlyStats.map(month => [
        month.monthName.substring(0, 7), // Сокращаем название месяца
        month.totalLeads,
        Math.round(month.totalRevenue / 1000)
      ])
    );
    
    const dynamicsChart = sheet.insertChart(
      Charts.newLineChart()
        .setDataRange(sheet.getRange(1, 11, dynamicsChartData.length, 3))
        .setOption('title', 'Динамика лидов и выручки по месяцам')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'top' })
        .setOption('hAxis', { title: 'Месяц', slantedText: true })
        .setOption('vAxes', {
          0: { title: 'Лиды' },
          1: { title: 'Выручка, тыс. руб.' }
        })
        .setOption('series', {
          0: { targetAxisIndex: 0, color: '#4285f4' },
          1: { targetAxisIndex: 1, color: '#34a853' }
        })
        .setPosition(3, 11, 0, 0)
        .setOption('width', 800)
        .setOption('height', 400)
        .build()
    );
    
    sheet.getRange(1, 11, dynamicsChartData.length, 3).setValues(dynamicsChartData);
  }
  
  // 2. Конверсия по месяцам
  if (monthlyData.monthlyStats.length > 0) {
    const conversionChartData = [['Месяц', 'Конверсия %']].concat(
      monthlyData.monthlyStats.map(month => [
        month.monthName.substring(0, 7),
        month.conversionRate
      ])
    );
    
    const conversionChart = sheet.insertChart(
      Charts.newLineChart()
        .setDataRange(sheet.getRange(1, 15, conversionChartData.length, 2))
        .setOption('title', 'Динамика конверсии по месяцам')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Месяц', slantedText: true })
        .setOption('vAxis', { title: 'Конверсия, %' })
        .setOption('curveType', 'function')
        .setPosition(3, 21, 0, 0)
        .setOption('width', 600)
        .setOption('height', 350)
        .build()
    );
    
    sheet.getRange(1, 15, conversionChartData.length, 2).setValues(conversionChartData);
  }
  
  // 3. Сравнение кварталов
  if (monthlyData.seasonalAnalysis && monthlyData.seasonalAnalysis.quarters) {
    const quarterChartData = [['Квартал', 'Средние лиды', 'Средняя выручка (тыс.)']];
    
    Object.entries(monthlyData.seasonalAnalysis.quarters).forEach(([quarter, data]) => {
      if (data.months.length > 0) {
        quarterChartData.push([
          quarter,
          data.avgLeads,
          Math.round(data.avgRevenue / 1000)
        ]);
      }
    });
    
    if (quarterChartData.length > 1) {
      const quarterChart = sheet.insertChart(
        Charts.newColumnChart()
          .setDataRange(sheet.getRange(1, 18, quarterChartData.length, 3))
          .setOption('title', 'Сравнение кварталов')
          .setOption('titleTextStyle', { fontSize: 14, bold: true })
          .setOption('legend', { position: 'top' })
          .setOption('hAxis', { title: 'Квартал' })
          .setOption('vAxes', {
            0: { title: 'Лиды' },
            1: { title: 'Выручка, тыс. руб.' }
          })
          .setOption('series', {
            0: { targetAxisIndex: 0, color: '#4285f4' },
            1: { targetAxisIndex: 1, color: '#34a853' }
          })
          .setPosition(25, 11, 0, 0)
          .setOption('width', 500)
          .setOption('height', 350)
          .build()
      );
      
      sheet.getRange(1, 18, quarterChartData.length, 3).setValues(quarterChartData);
    }
  }
  
  // 4. Прогноз (если есть данные)
  if (monthlyData.forecasting.forecasts.length > 0) {
    // Объединяем исторические данные с прогнозом
    const forecastChartData = [['Месяц', 'Факт лиды', 'Прогноз лиды']];
    
    // Последние 6 месяцев фактических данных
    const recentMonths = monthlyData.monthlyStats.slice(-6);
    recentMonths.forEach(month => {
      forecastChartData.push([month.monthName.substring(0, 7), month.totalLeads, null]);
    });
    
    // Прогнозы
    monthlyData.forecasting.forecasts.forEach(forecast => {
      forecastChartData.push([
        forecast.monthName.substring(0, 7),
        null,
        forecast.forecastLeads
      ]);
    });
    
    const forecastChart = sheet.insertChart(
      Charts.newLineChart()
        .setDataRange(sheet.getRange(1, 22, forecastChartData.length, 3))
        .setOption('title', 'Прогноз лидов на ближайшие месяцы')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'top' })
        .setOption('hAxis', { title: 'Месяц', slantedText: true })
        .setOption('vAxis', { title: 'Количество лидов' })
        .setOption('series', {
          0: { color: '#4285f4', lineWidth: 2 },
          1: { color: '#ea4335', lineWidth: 2, lineDashStyle: [4, 4] }
        })
        .setPosition(25, 21, 0, 0)
        .setOption('width', 600)
        .setOption('height', 350)
        .build()
    );
    
    sheet.getRange(1, 22, forecastChartData.length, 3).setValues(forecastChartData);
  }
}

/**
 * Вспомогательные функции для расчётов
 */

/**
 * Вычисляет количество рабочих дней в месяце
 */
function getWorkingDaysInMonth_(year, month) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  let workingDays = 0;
  
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month, day);
    const dayOfWeek = date.getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Не воскресенье и не суббота
      workingDays++;
    }
  }
  
  return workingDays;
}

/**
 * Вычисляет направление тренда
 */
function calculateTrendDirection_(values) {
  if (values.length < 3) return { direction: 'Недостаточно данных', slope: 0 };
  
  const trend = calculateLinearTrend_(values);
  let direction;
  
  if (Math.abs(trend.slope) < 0.1) direction = 'Стабильно';
  else if (trend.slope > 0) direction = 'Рост';
  else direction = 'Снижение';
  
  return { direction: direction, slope: trend.slope };
}

/**
 * Вычисляет линейный тренд
 */
function calculateLinearTrend_(values) {
  const n = values.length;
  if (n < 2) return { slope: 0, intercept: 0, r2: 0 };
  
  const sumX = (n * (n - 1)) / 2;
  const sumY = values.reduce((a, b) => a + b, 0);
  const sumXY = values.reduce((sum, y, x) => sum + x * y, 0);
  const sumXX = (n * (n - 1) * (2 * n - 1)) / 6;
  const sumYY = values.reduce((sum, y) => sum + y * y, 0);
  
  const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
  const intercept = (sumY - slope * sumX) / n;
  
  // Коэффициент детерминации R²
  const yMean = sumY / n;
  const ssRes = values.reduce((sum, y, x) => {
    const predicted = slope * x + intercept;
    return sum + Math.pow(y - predicted, 2);
  }, 0);
  const ssTot = values.reduce((sum, y) => sum + Math.pow(y - yMean, 2), 0);
  const r2 = ssTot > 0 ? 1 - (ssRes / ssTot) : 0;
  
  return { slope, intercept, r2 };
}

/**
 * Вычисляет средний темп роста
 */
function calculateAverageGrowth_(values) {
  if (values.length < 2) return 0;
  
  let totalGrowth = 0;
  let validPairs = 0;
  
  for (let i = 1; i < values.length; i++) {
    if (values[i - 1] > 0) {
      totalGrowth += ((values[i] - values[i - 1]) / values[i - 1]) * 100;
      validPairs++;
    }
  }
  
  return validPairs > 0 ? totalGrowth / validPairs : 0;
}

/**
 * Вычисляет базовую статистику
 */
function calculateStats_(values) {
  const n = values.length;
  const mean = values.reduce((a, b) => a + b, 0) / n;
  const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / n;
  const std = Math.sqrt(variance);
  
  return { mean, std, variance };
}

/**
 * Создаёт пустой отчёт когда нет данных
 * @private
 */
function createEmptyMonthlyReport_() {
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.MONTHLY_COMPARISON);
  clearSheetData_(sheet);
  applySheetFormatting_(sheet, 'Сравнение по месяцам');
  
  sheet.getRange('A3').setValue('ℹ️ Нет данных для месячного сравнения');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A5').setValue('Убедитесь, что:');
  sheet.getRange('A6').setValue('• В системе есть данные за несколько месяцев');
  sheet.getRange('A7').setValue('• Даты корректно заполнены');
  sheet.getRange('A8').setValue('• Выполнена синхронизация данных');
  
  updateLastUpdateTime_(CONFIG.SHEETS.MONTHLY_COMPARISON);
}
