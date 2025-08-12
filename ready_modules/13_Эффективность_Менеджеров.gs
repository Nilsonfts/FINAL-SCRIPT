/**
 * 👨‍💼 Эффективность Менеджеров - Анализ работы
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Apps Script: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit
 * 2. Создайте новый файл: + → Script file
 * 3. Назовите файл: 13_Эффективность_Менеджеров (без .gs)
 * 4. Удалите весь код по умолчанию
 * 5. Скопируйте и вставьте код ниже
 * 6. Сохраните файл (Ctrl+S)
 */

/**
 * ЭФФЕКТИВНОСТЬ МЕНЕДЖЕРОВ
 * Анализ работы менеджеров по продажам и их KPI
 * @fileoverview Модуль анализа производительности менеджеров по различным метрикам
 */

/**
 * Основная функция анализа эффективности менеджеров
 * Создаёт детальный отчёт по работе каждого менеджера
 */
function analyzeManagerPerformance() {
  try {
    logInfo_('MANAGER_PERFORMANCE', 'Начало анализа эффективности менеджеров');
    
    const startTime = new Date();
    
    // Получаем данные для анализа менеджеров
    const managersData = getManagersAnalysisData_();
    
    if (!managersData || managersData.managers.length === 0) {
      logWarning_('MANAGER_PERFORMANCE', 'Нет данных для анализа менеджеров');
      createEmptyManagersReport_();
      return;
    }
    
    // Создаём или обновляем лист
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.MANAGER_PERFORMANCE);
    clearSheetData_(sheet);
    applySheetFormatting_(sheet, 'Эффективность менеджеров');
    
    // Строим структуру отчёта
    createManagersAnalysisStructure_(sheet, managersData);
    
    // Добавляем диаграммы
    addManagersAnalysisCharts_(sheet, managersData);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('MANAGER_PERFORMANCE', `Анализ менеджеров завершён за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.MANAGER_PERFORMANCE);
    
  } catch (error) {
    logError_('MANAGER_PERFORMANCE', 'Ошибка анализа эффективности менеджеров', error);
    throw error;
  }
}

/**
 * Получает данные для анализа менеджеров
 * @returns {Object} Данные анализа менеджеров
 * @private
 */
function getManagersAnalysisData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return null;
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Анализ менеджеров
  const managers = analyzeManagers_(headers, rows);
  
  // Командная статистика
  const teamStats = calculateTeamStats_(managers);
  
  // Анализ по временным периодам
  const timeAnalysis = analyzeManagersTimePerformance_(headers, rows);
  
  // Анализ воронки продаж по менеджерам
  const funnelAnalysis = analyzeManagersFunnel_(headers, rows);
  
  // Сравнительный анализ
  const competitiveAnalysis = compareManagers_(managers);
  
  // KPI и рейтинги
  const kpiAnalysis = calculateManagerKPIs_(managers, teamStats);
  
  return {
    managers: managers,
    teamStats: teamStats,
    timeAnalysis: timeAnalysis,
    funnelAnalysis: funnelAnalysis,
    competitiveAnalysis: competitiveAnalysis,
    kpiAnalysis: kpiAnalysis,
    totalManagers: managers.length,
    analysisDate: getCurrentDateMoscow_()
  };
}

/**
 * Анализирует производительность менеджеров
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Данные по менеджерам
 * @private
 */
function analyzeManagers_(headers, rows) {
  const managerStats = {};
  
  // Определяем индексы колонок
  const managerIndex = findHeaderIndex_(headers, 'Менеджер') || findHeaderIndex_(headers, 'Ответственный');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const closedIndex = findHeaderIndex_(headers, 'Дата закрытия');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const phoneIndex = findHeaderIndex_(headers, 'Телефон');
  const emailIndex = findHeaderIndex_(headers, 'Email');
  
  if (managerIndex === -1) {
    // Если нет колонки менеджера, создаем общего менеджера
    managerStats['Общий менеджер'] = {
      name: 'Общий менеджер',
      totalLeads: 0,
      wonDeals: 0,
      lostDeals: 0,
      inProgress: 0,
      totalRevenue: 0,
      avgDealValue: 0,
      conversionRate: 0,
      avgDealCycle: 0,
      dealsTimeSum: 0,
      channelDistribution: {},
      monthlyPerformance: {},
      dailyActivity: {},
      clientsWithContact: 0,
      dealSizes: []
    };
  }
  
  rows.forEach(row => {
    const manager = managerIndex !== -1 ? (row[managerIndex] || 'Не назначен') : 'Общий менеджер';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const createdDate = parseDate_(row[createdIndex]);
    const closedDate = parseDate_(row[closedIndex]);
    const channel = row[channelIndex] || 'Неизвестно';
    const phone = row[phoneIndex] || '';
    const email = row[emailIndex] || '';
    
    // Инициализируем менеджера если нет
    if (!managerStats[manager]) {
      managerStats[manager] = {
        name: manager,
        totalLeads: 0,
        wonDeals: 0,
        lostDeals: 0,
        inProgress: 0,
        totalRevenue: 0,
        avgDealValue: 0,
        conversionRate: 0,
        avgDealCycle: 0,
        dealsTimeSum: 0,
        channelDistribution: {},
        monthlyPerformance: {},
        dailyActivity: {},
        clientsWithContact: 0,
        dealSizes: []
      };
    }
    
    const stats = managerStats[manager];
    stats.totalLeads++;
    
    // Анализ статусов
    if (status === 'success') {
      stats.wonDeals++;
      stats.totalRevenue += budget;
      stats.dealSizes.push(budget);
      
      // Время закрытия сделки
      if (createdDate && closedDate && closedDate >= createdDate) {
        const dealCycle = Math.round((closedDate - createdDate) / (1000 * 60 * 60 * 24));
        stats.dealsTimeSum += dealCycle;
      }
    } else if (status === 'failure') {
      stats.lostDeals++;
    } else {
      stats.inProgress++;
    }
    
    // Каналы
    stats.channelDistribution[channel] = (stats.channelDistribution[channel] || 0) + 1;
    
    // Месячная активность
    if (createdDate) {
      const monthKey = formatDate_(createdDate, 'YYYY-MM');
      if (!stats.monthlyPerformance[monthKey]) {
        stats.monthlyPerformance[monthKey] = { leads: 0, deals: 0, revenue: 0 };
      }
      stats.monthlyPerformance[monthKey].leads++;
      if (status === 'success') {
        stats.monthlyPerformance[monthKey].deals++;
        stats.monthlyPerformance[monthKey].revenue += budget;
      }
      
      // Ежедневная активность
      const dayKey = formatDate_(createdDate, 'YYYY-MM-DD');
      stats.dailyActivity[dayKey] = (stats.dailyActivity[dayKey] || 0) + 1;
    }
    
    // Полнота контактных данных
    if (phone || email) {
      stats.clientsWithContact++;
    }
  });
  
  // Вычисляем производные метрики
  return Object.values(managerStats).map(stats => {
    stats.conversionRate = stats.totalLeads > 0 ? (stats.wonDeals / stats.totalLeads * 100) : 0;
    stats.avgDealValue = stats.wonDeals > 0 ? (stats.totalRevenue / stats.wonDeals) : 0;
    stats.avgDealCycle = stats.wonDeals > 0 ? Math.round(stats.dealsTimeSum / stats.wonDeals) : 0;
    stats.contactRate = stats.totalLeads > 0 ? (stats.clientsWithContact / stats.totalLeads * 100) : 0;
    
    // Топ канал менеджера
    stats.topChannel = Object.entries(stats.channelDistribution)
      .sort(([,a], [,b]) => b - a)[0]?.[0] || 'Нет данных';
    
    // Активные месяцы
    stats.activeMonths = Object.keys(stats.monthlyPerformance).length;
    
    // Активные дни
    stats.activeDays = Object.keys(stats.dailyActivity).length;
    
    // Медианная сумма сделки
    if (stats.dealSizes.length > 0) {
      const sortedSizes = stats.dealSizes.sort((a, b) => a - b);
      stats.medianDealSize = sortedSizes[Math.floor(sortedSizes.length / 2)];
      stats.maxDealSize = Math.max(...sortedSizes);
      stats.minDealSize = Math.min(...sortedSizes.filter(size => size > 0));
    } else {
      stats.medianDealSize = 0;
      stats.maxDealSize = 0;
      stats.minDealSize = 0;
    }
    
    // Стабильность работы (коэффициент вариации ежемесячной активности)
    const monthlyLeads = Object.values(stats.monthlyPerformance).map(month => month.leads);
    if (monthlyLeads.length > 1) {
      const avgMonthlyLeads = monthlyLeads.reduce((a, b) => a + b, 0) / monthlyLeads.length;
      const variance = monthlyLeads.reduce((sum, leads) => sum + Math.pow(leads - avgMonthlyLeads, 2), 0) / monthlyLeads.length;
      stats.stabilityCoeff = avgMonthlyLeads > 0 ? Math.sqrt(variance) / avgMonthlyLeads : 0;
    } else {
      stats.stabilityCoeff = 0;
    }
    
    return stats;
  }).sort((a, b) => b.totalRevenue - a.totalRevenue);
}

/**
 * Вычисляет командную статистику
 * @param {Array} managers - Данные менеджеров
 * @returns {Object} Командная статистика
 * @private
 */
function calculateTeamStats_(managers) {
  const teamStats = {
    totalManagers: managers.length,
    totalLeads: 0,
    totalRevenue: 0,
    totalDeals: 0,
    avgConversionRate: 0,
    avgDealValue: 0,
    avgDealCycle: 0,
    topPerformers: [],
    underPerformers: [],
    medianRevenue: 0
  };
  
  // Суммарные показатели
  managers.forEach(manager => {
    teamStats.totalLeads += manager.totalLeads;
    teamStats.totalRevenue += manager.totalRevenue;
    teamStats.totalDeals += manager.wonDeals;
  });
  
  // Средние показатели
  if (managers.length > 0) {
    teamStats.avgConversionRate = managers.reduce((sum, mgr) => sum + mgr.conversionRate, 0) / managers.length;
    teamStats.avgDealValue = teamStats.totalDeals > 0 ? teamStats.totalRevenue / teamStats.totalDeals : 0;
    teamStats.avgDealCycle = managers
      .filter(mgr => mgr.avgDealCycle > 0)
      .reduce((sum, mgr) => sum + mgr.avgDealCycle, 0) / Math.max(1, managers.filter(mgr => mgr.avgDealCycle > 0).length);
    
    // Медианная выручка
    const revenues = managers.map(mgr => mgr.totalRevenue).sort((a, b) => a - b);
    teamStats.medianRevenue = revenues[Math.floor(revenues.length / 2)];
    
    // Топ исполнители (выше медианы по выручке)
    teamStats.topPerformers = managers.filter(mgr => mgr.totalRevenue > teamStats.medianRevenue);
    
    // Отстающие (ниже медианы и с низкой конверсией)
    teamStats.underPerformers = managers.filter(mgr => 
      mgr.totalRevenue < teamStats.medianRevenue && 
      mgr.conversionRate < teamStats.avgConversionRate
    );
  }
  
  return teamStats;
}

/**
 * Анализирует временную производительность менеджеров
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Временной анализ
 * @private
 */
function analyzeManagersTimePerformance_(headers, rows) {
  const timeAnalysis = {
    monthlyTrends: {},
    weeklyActivity: {},
    hourlyActivity: {}
  };
  
  const managerIndex = findHeaderIndex_(headers, 'Менеджер') || findHeaderIndex_(headers, 'Ответственный');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  
  rows.forEach(row => {
    const manager = managerIndex !== -1 ? (row[managerIndex] || 'Не назначен') : 'Общий менеджер';
    const createdDate = parseDate_(row[createdIndex]);
    const status = row[statusIndex] || '';
    
    if (!createdDate) return;
    
    const monthKey = formatDate_(createdDate, 'YYYY-MM');
    const dayOfWeek = createdDate.getDay(); // 0 = воскресенье
    const hour = createdDate.getHours();
    
    // Месячные тренды
    if (!timeAnalysis.monthlyTrends[monthKey]) {
      timeAnalysis.monthlyTrends[monthKey] = {};
    }
    if (!timeAnalysis.monthlyTrends[monthKey][manager]) {
      timeAnalysis.monthlyTrends[monthKey][manager] = { leads: 0, deals: 0 };
    }
    timeAnalysis.monthlyTrends[monthKey][manager].leads++;
    if (status === 'success') {
      timeAnalysis.monthlyTrends[monthKey][manager].deals++;
    }
    
    // Активность по дням недели
    if (!timeAnalysis.weeklyActivity[manager]) {
      timeAnalysis.weeklyActivity[manager] = {};
    }
    timeAnalysis.weeklyActivity[manager][dayOfWeek] = (timeAnalysis.weeklyActivity[manager][dayOfWeek] || 0) + 1;
    
    // Почасовая активность
    if (!timeAnalysis.hourlyActivity[manager]) {
      timeAnalysis.hourlyActivity[manager] = {};
    }
    timeAnalysis.hourlyActivity[manager][hour] = (timeAnalysis.hourlyActivity[manager][hour] || 0) + 1;
  });
  
  return timeAnalysis;
}

/**
 * Анализирует воронку продаж по менеджерам
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Анализ воронки
 * @private
 */
function analyzeManagersFunnel_(headers, rows) {
  const funnelAnalysis = {};
  
  const managerIndex = findHeaderIndex_(headers, 'Менеджер') || findHeaderIndex_(headers, 'Ответственный');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  
  rows.forEach(row => {
    const manager = managerIndex !== -1 ? (row[managerIndex] || 'Не назначен') : 'Общий менеджер';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    
    if (!funnelAnalysis[manager]) {
      funnelAnalysis[manager] = {
        manager: manager,
        newLeads: 0,
        qualified: 0,
        proposal: 0,
        negotiation: 0,
        closed: 0,
        lost: 0
      };
    }
    
    const funnel = funnelAnalysis[manager];
    
    // Определяем стадию воронки
    if (status === 'success') {
      funnel.closed++;
    } else if (status === 'failure') {
      funnel.lost++;
    } else if (budget > 0) {
      funnel.proposal++;
    } else if (status === 'qualified') {
      funnel.qualified++;
    } else if (status === 'negotiation') {
      funnel.negotiation++;
    } else {
      funnel.newLeads++;
    }
  });
  
  // Вычисляем конверсию между этапами
  return Object.values(funnelAnalysis).map(funnel => {
    const totalLeads = funnel.newLeads + funnel.qualified + funnel.proposal + funnel.negotiation + funnel.closed + funnel.lost;
    
    funnel.qualificationRate = totalLeads > 0 ? (funnel.qualified / totalLeads * 100) : 0;
    funnel.proposalRate = (funnel.qualified + funnel.proposal) > 0 ? (funnel.proposal / (funnel.qualified + funnel.proposal) * 100) : 0;
    funnel.closingRate = (funnel.proposal + funnel.negotiation + funnel.closed) > 0 ? (funnel.closed / (funnel.proposal + funnel.negotiation + funnel.closed) * 100) : 0;
    funnel.overallConversion = totalLeads > 0 ? (funnel.closed / totalLeads * 100) : 0;
    
    return funnel;
  }).sort((a, b) => b.closed - a.closed);
}

/**
 * Сравнивает менеджеров между собой
 * @param {Array} managers - Данные менеджеров
 * @returns {Object} Сравнительный анализ
 * @private
 */
function compareManagers_(managers) {
  if (managers.length < 2) return { rankings: {}, gaps: {} };
  
  const rankings = {
    byRevenue: [...managers].sort((a, b) => b.totalRevenue - a.totalRevenue),
    byConversion: [...managers].sort((a, b) => b.conversionRate - a.conversionRate),
    byDealValue: [...managers].sort((a, b) => b.avgDealValue - a.avgDealValue),
    bySpeed: [...managers].filter(m => m.avgDealCycle > 0).sort((a, b) => a.avgDealCycle - b.avgDealCycle),
    byVolume: [...managers].sort((a, b) => b.totalLeads - a.totalLeads)
  };
  
  // Анализ разрывов в производительности
  const gaps = {};
  
  if (rankings.byRevenue.length >= 2) {
    const top = rankings.byRevenue[0];
    const bottom = rankings.byRevenue[rankings.byRevenue.length - 1];
    
    gaps.revenueGap = {
      topManager: top.name,
      topValue: top.totalRevenue,
      bottomManager: bottom.name,
      bottomValue: bottom.totalRevenue,
      gapMultiple: bottom.totalRevenue > 0 ? top.totalRevenue / bottom.totalRevenue : 0
    };
  }
  
  if (rankings.byConversion.length >= 2) {
    const topConv = rankings.byConversion[0];
    const bottomConv = rankings.byConversion[rankings.byConversion.length - 1];
    
    gaps.conversionGap = {
      topManager: topConv.name,
      topValue: topConv.conversionRate,
      bottomManager: bottomConv.name,
      bottomValue: bottomConv.conversionRate,
      gapDifference: topConv.conversionRate - bottomConv.conversionRate
    };
  }
  
  return { rankings, gaps };
}

/**
 * Вычисляет KPI менеджеров
 * @param {Array} managers - Данные менеджеров
 * @param {Object} teamStats - Командная статистика
 * @returns {Array} KPI анализ
 * @private
 */
function calculateManagerKPIs_(managers, teamStats) {
  return managers.map(manager => {
    // Комплексный KPI Score (от 0 до 100)
    let kpiScore = 0;
    
    // Вес компонентов KPI
    const weights = {
      revenue: 0.3,        // 30% - выручка
      conversion: 0.25,    // 25% - конверсия
      volume: 0.2,         // 20% - объём работы
      speed: 0.15,         // 15% - скорость закрытия
      stability: 0.1       // 10% - стабильность
    };
    
    // Нормализуем показатели относительно команды
    const revenueScore = teamStats.totalRevenue > 0 ? 
      Math.min(100, (manager.totalRevenue / teamStats.totalRevenue) * 100 * teamStats.totalManagers) : 0;
    
    const conversionScore = Math.min(100, manager.conversionRate);
    
    const volumeScore = teamStats.totalLeads > 0 ? 
      Math.min(100, (manager.totalLeads / teamStats.totalLeads) * 100 * teamStats.totalManagers) : 0;
    
    const speedScore = manager.avgDealCycle > 0 && teamStats.avgDealCycle > 0 ? 
      Math.max(0, 100 - (manager.avgDealCycle / teamStats.avgDealCycle - 1) * 50) : 50;
    
    const stabilityScore = Math.max(0, 100 - manager.stabilityCoeff * 100);
    
    // Итоговый KPI
    kpiScore = 
      revenueScore * weights.revenue +
      conversionScore * weights.conversion +
      volumeScore * weights.volume +
      speedScore * weights.speed +
      stabilityScore * weights.stability;
    
    // Определяем категорию эффективности
    let category;
    if (kpiScore >= 80) category = '🌟 Звезда';
    else if (kpiScore >= 60) category = '🚀 Сильный';
    else if (kpiScore >= 40) category = '📈 Средний';
    else if (kpiScore >= 20) category = '📉 Слабый';
    else category = '🔴 Критический';
    
    // Рекомендации по улучшению
    const recommendations = [];
    
    if (manager.conversionRate < teamStats.avgConversionRate * 0.8) {
      recommendations.push('Улучшить навыки конверсии лидов');
    }
    if (manager.avgDealCycle > teamStats.avgDealCycle * 1.3) {
      recommendations.push('Сократить время закрытия сделок');
    }
    if (manager.totalLeads < teamStats.totalLeads / teamStats.totalManagers * 0.7) {
      recommendations.push('Увеличить активность по работе с лидами');
    }
    if (manager.stabilityCoeff > 0.5) {
      recommendations.push('Повысить стабильность работы');
    }
    if (manager.contactRate < 80) {
      recommendations.push('Улучшить качество сбора контактных данных');
    }
    
    return {
      manager: manager.name,
      kpiScore: Math.round(kpiScore),
      category: category,
      components: {
        revenue: Math.round(revenueScore),
        conversion: Math.round(conversionScore),
        volume: Math.round(volumeScore),
        speed: Math.round(speedScore),
        stability: Math.round(stabilityScore)
      },
      recommendations: recommendations,
      strengths: identifyStrengths_(manager, teamStats),
      weaknesses: identifyWeaknesses_(manager, teamStats)
    };
  }).sort((a, b) => b.kpiScore - a.kpiScore);
}

/**
 * Определяет сильные стороны менеджера
 */
function identifyStrengths_(manager, teamStats) {
  const strengths = [];
  
  if (manager.conversionRate > teamStats.avgConversionRate * 1.2) {
    strengths.push('Высокая конверсия');
  }
  if (manager.avgDealValue > teamStats.avgDealValue * 1.2) {
    strengths.push('Крупные сделки');
  }
  if (manager.avgDealCycle > 0 && manager.avgDealCycle < teamStats.avgDealCycle * 0.8) {
    strengths.push('Быстрое закрытие');
  }
  if (manager.stabilityCoeff < 0.3) {
    strengths.push('Стабильная работа');
  }
  if (manager.contactRate > 90) {
    strengths.push('Качественная работа с данными');
  }
  
  return strengths;
}

/**
 * Определяет слабые стороны менеджера
 */
function identifyWeaknesses_(manager, teamStats) {
  const weaknesses = [];
  
  if (manager.conversionRate < teamStats.avgConversionRate * 0.8) {
    weaknesses.push('Низкая конверсия');
  }
  if (manager.avgDealValue < teamStats.avgDealValue * 0.8) {
    weaknesses.push('Мелкие сделки');
  }
  if (manager.avgDealCycle > teamStats.avgDealCycle * 1.3) {
    weaknesses.push('Долгое закрытие');
  }
  if (manager.stabilityCoeff > 0.6) {
    weaknesses.push('Нестабильная работа');
  }
  if (manager.contactRate < 70) {
    weaknesses.push('Плохое качество данных');
  }
  
  return weaknesses;
}

/**
 * Создаёт структуру отчёта по менеджерам
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} managersData - Данные анализа менеджеров
 * @private
 */
function createManagersAnalysisStructure_(sheet, managersData) {
  let currentRow = 3;
  
  // 1. Командная статистика
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('👥 КОМАНДНАЯ СТАТИСТИКА');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const teamStats = [
    ['Всего менеджеров:', managersData.teamStats.totalManagers],
    ['Общее кол-во лидов:', managersData.teamStats.totalLeads],
    ['Общая выручка:', formatCurrency_(managersData.teamStats.totalRevenue)],
    ['Средняя конверсия команды:', `${managersData.teamStats.avgConversionRate.toFixed(1)}%`],
    ['Средний чек команды:', formatCurrency_(managersData.teamStats.avgDealValue)],
    ['Среднее время сделки:', managersData.teamStats.avgDealCycle > 0 ? `${Math.round(managersData.teamStats.avgDealCycle)} дн.` : 'Н/Д'],
    ['Топ исполнителей:', managersData.teamStats.topPerformers.length],
    ['Требуют внимания:', managersData.teamStats.underPerformers.length]
  ];
  
  sheet.getRange(currentRow, 1, teamStats.length, 2).setValues(teamStats);
  sheet.getRange(currentRow, 1, teamStats.length, 1).setFontWeight('bold');
  currentRow += teamStats.length + 2;
  
  // 2. Рейтинг менеджеров по KPI
  sheet.getRange(currentRow, 1, 1, 7).merge();
  sheet.getRange(currentRow, 1).setValue('🏆 РЕЙТИНГ МЕНЕДЖЕРОВ ПО KPI');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 7));
  currentRow++;
  
  const kpiHeaders = [['Менеджер', 'KPI Score', 'Категория', 'Лиды', 'Сделки', 'Выручка', 'Конверсия']];
  sheet.getRange(currentRow, 1, 1, 7).setValues(kpiHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 7));
  currentRow++;
  
  const kpiData = managersData.kpiAnalysis.map((kpi, index) => {
    const manager = managersData.managers.find(m => m.name === kpi.manager);
    return [
      kpi.manager,
      kpi.kpiScore,
      kpi.category,
      manager.totalLeads,
      manager.wonDeals,
      formatCurrency_(manager.totalRevenue),
      `${manager.conversionRate.toFixed(1)}%`
    ];
  });
  
  if (kpiData.length > 0) {
    sheet.getRange(currentRow, 1, kpiData.length, 7).setValues(kpiData);
    
    // Цветовое выделение категорий
    kpiData.forEach((row, index) => {
      const categoryCell = sheet.getRange(currentRow + index, 3);
      const kpiCell = sheet.getRange(currentRow + index, 2);
      
      if (row[2].includes('Звезда')) {
        categoryCell.setBackground('#d4edda');
        kpiCell.setBackground('#d4edda');
      } else if (row[2].includes('Сильный')) {
        categoryCell.setBackground('#cce5ff');
        kpiCell.setBackground('#cce5ff');
      } else if (row[2].includes('Средний')) {
        categoryCell.setBackground('#fff3cd');
        kpiCell.setBackground('#fff3cd');
      } else if (row[2].includes('Слабый')) {
        categoryCell.setBackground('#ffe6cc');
        kpiCell.setBackground('#ffe6cc');
      } else if (row[2].includes('Критический')) {
        categoryCell.setBackground('#f8d7da');
        kpiCell.setBackground('#f8d7da');
      }
    });
    
    currentRow += kpiData.length;
  }
  currentRow += 2;
  
  // 3. Детальная статистика по менеджерам
  sheet.getRange(currentRow, 1, 1, 9).merge();
  sheet.getRange(currentRow, 1).setValue('📋 ДЕТАЛЬНАЯ СТАТИСТИКА МЕНЕДЖЕРОВ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 9));
  currentRow++;
  
  const detailHeaders = [['Менеджер', 'Лиды', 'Сделки', 'Выручка', 'Средний чек', 'Время сделки', 'Топ канал', 'Активных дней', 'Полнота данных']];
  sheet.getRange(currentRow, 1, 1, 9).setValues(detailHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 9));
  currentRow++;
  
  const detailData = managersData.managers.map(manager => [
    manager.name,
    manager.totalLeads,
    manager.wonDeals,
    formatCurrency_(manager.totalRevenue),
    formatCurrency_(manager.avgDealValue),
    manager.avgDealCycle > 0 ? `${manager.avgDealCycle} дн.` : 'Н/Д',
    manager.topChannel,
    manager.activeDays,
    `${manager.contactRate.toFixed(1)}%`
  ]);
  
  if (detailData.length > 0) {
    sheet.getRange(currentRow, 1, detailData.length, 9).setValues(detailData);
    currentRow += detailData.length;
  }
  currentRow += 2;
  
  // 4. Детализация топ-3 менеджеров
  const top3Managers = managersData.managers.slice(0, 3);
  top3Managers.forEach((manager, index) => {
    const kpiInfo = managersData.kpiAnalysis.find(kpi => kpi.manager === manager.name);
    
    sheet.getRange(currentRow, 1, 1, 6).merge();
    sheet.getRange(currentRow, 1).setValue(`${index + 1}. ${manager.name.toUpperCase()} - ${kpiInfo.category}`);
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
    currentRow++;
    
    const managerDetails = [
      ['Основные показатели:', ''],
      [`• Всего лидов`, manager.totalLeads],
      [`• Успешных сделок`, `${manager.wonDeals} (${manager.conversionRate.toFixed(1)}%)`],
      [`• Общая выручка`, formatCurrency_(manager.totalRevenue)],
      [`• Средний чек`, formatCurrency_(manager.avgDealValue)],
      [`• Медианная сделка`, formatCurrency_(manager.medianDealSize)],
      [`• Время закрытия сделок`, manager.avgDealCycle > 0 ? `${manager.avgDealCycle} дней` : 'Н/Д'],
      ['Активность:', ''],
      [`• Активных месяцев`, manager.activeMonths],
      [`• Активных дней`, manager.activeDays],
      [`• Топ канал`, manager.topChannel],
      [`• Полнота контактов`, `${manager.contactRate.toFixed(1)}%`]
    ];
    
    sheet.getRange(currentRow, 1, managerDetails.length, 2).setValues(managerDetails);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow + 7, 1).setFontWeight('bold');
    currentRow += managerDetails.length;
    
    // KPI компоненты
    if (kpiInfo) {
      sheet.getRange(currentRow, 1).setValue('KPI компоненты:');
      sheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
      
      const kpiComponents = [
        [`• Выручка`, `${kpiInfo.components.revenue}/100`],
        [`• Конверсия`, `${kpiInfo.components.conversion}/100`],
        [`• Объём работы`, `${kpiInfo.components.volume}/100`],
        [`• Скорость`, `${kpiInfo.components.speed}/100`],
        [`• Стабильность`, `${kpiInfo.components.stability}/100`]
      ];
      
      sheet.getRange(currentRow, 1, kpiComponents.length, 2).setValues(kpiComponents);
      currentRow += kpiComponents.length;
      
      // Рекомендации
      if (kpiInfo.recommendations.length > 0) {
        sheet.getRange(currentRow, 1).setValue('Рекомендации:');
        sheet.getRange(currentRow, 1).setFontWeight('bold');
        currentRow++;
        
        kpiInfo.recommendations.forEach(recommendation => {
          sheet.getRange(currentRow, 1).setValue(`• ${recommendation}`);
          currentRow++;
        });
      }
    }
    
    currentRow += 2;
  });
  
  // 5. Анализ разрывов в команде
  if (managersData.competitiveAnalysis.gaps.revenueGap) {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue('📊 АНАЛИЗ РАЗРЫВОВ В КОМАНДЕ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const gap = managersData.competitiveAnalysis.gaps.revenueGap;
    const convGap = managersData.competitiveAnalysis.gaps.conversionGap;
    
    const gapAnalysis = [
      ['Разрыв по выручке:', ''],
      [`• Лидер: ${gap.topManager}`, formatCurrency_(gap.topValue)],
      [`• Аутсайдер: ${gap.bottomManager}`, formatCurrency_(gap.bottomValue)],
      [`• Разница`, `в ${gap.gapMultiple.toFixed(1)} раз`],
      ['', ''],
      ['Разрыв по конверсии:', ''],
      [`• Лидер: ${convGap.topManager}`, `${convGap.topValue.toFixed(1)}%`],
      [`• Аутсайдер: ${convGap.bottomManager}`, `${convGap.bottomValue.toFixed(1)}%`],
      [`• Разница`, `${convGap.gapDifference.toFixed(1)} п.п.`]
    ];
    
    sheet.getRange(currentRow, 1, gapAnalysis.length, 2).setValues(gapAnalysis);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow + 5, 1).setFontWeight('bold');
    currentRow += gapAnalysis.length;
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 9);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 9);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к анализу менеджеров
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} managersData - Данные анализа менеджеров
 * @private
 */
function addManagersAnalysisCharts_(sheet, managersData) {
  // 1. Столбчатая диаграмма выручки по менеджерам
  if (managersData.managers.length > 0) {
    const revenueChartData = [['Менеджер', 'Выручка']].concat(
      managersData.managers.slice(0, 8).map(manager => [manager.name, manager.totalRevenue])
    );
    
    const revenueChart = sheet.insertChart(
      Charts.newColumnChart()
        .setDataRange(sheet.getRange(1, 11, revenueChartData.length, 2))
        .setOption('title', 'Выручка по менеджерам')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Менеджер', slantedText: true })
        .setOption('vAxis', { title: 'Выручка, руб.' })
        .setPosition(3, 11, 0, 0)
        .setOption('width', 600)
        .setOption('height', 350)
        .build()
    );
    
    sheet.getRange(1, 11, revenueChartData.length, 2).setValues(revenueChartData);
  }
  
  // 2. Диаграмма рассеяния: конверсия vs объём
  if (managersData.managers.length > 0) {
    const scatterData = [['Конверсия %', 'Количество лидов', 'Менеджер']].concat(
      managersData.managers.map(manager => [manager.conversionRate, manager.totalLeads, manager.name])
    );
    
    const scatterChart = sheet.insertChart(
      Charts.newScatterChart()
        .setDataRange(sheet.getRange(1, 14, scatterData.length, 3))
        .setOption('title', 'Конверсия vs Объём работы')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Конверсия, %' })
        .setOption('vAxis', { title: 'Количество лидов' })
        .setPosition(3, 18, 0, 0)
        .setOption('width', 500)
        .setOption('height', 350)
        .build()
    );
    
    sheet.getRange(1, 14, scatterData.length, 3).setValues(scatterData);
  }
  
  // 3. Круговая диаграмма KPI категорий
  if (managersData.kpiAnalysis.length > 0) {
    const categoryCount = {};
    managersData.kpiAnalysis.forEach(kpi => {
      const category = kpi.category.replace(/^[^\w]+\s/, ''); // Убираем эмодзи
      categoryCount[category] = (categoryCount[category] || 0) + 1;
    });
    
    const categoryChartData = [['Категория', 'Количество']].concat(
      Object.entries(categoryCount).map(([category, count]) => [category, count])
    );
    
    const categoryChart = sheet.insertChart(
      Charts.newPieChart()
        .setDataRange(sheet.getRange(1, 18, categoryChartData.length, 2))
        .setOption('title', 'Распределение менеджеров по категориям')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'right' })
        .setOption('chartArea', { width: '80%', height: '80%' })
        .setPosition(25, 11, 0, 0)
        .setOption('width', 500)
        .setOption('height', 350)
        .build()
    );
    
    sheet.getRange(1, 18, categoryChartData.length, 2).setValues(categoryChartData);
  }
  
  // 4. Гистограмма KPI Score
  if (managersData.kpiAnalysis.length > 0) {
    const kpiChartData = [['Менеджер', 'KPI Score']].concat(
      managersData.kpiAnalysis.map(kpi => [kpi.manager, kpi.kpiScore])
    );
    
    const kpiChart = sheet.insertChart(
      Charts.newColumnChart()
        .setDataRange(sheet.getRange(1, 21, kpiChartData.length, 2))
        .setOption('title', 'KPI Score менеджеров')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Менеджер', slantedText: true })
        .setOption('vAxis', { title: 'KPI Score', minValue: 0, maxValue: 100 })
        .setPosition(25, 18, 0, 0)
        .setOption('width', 600)
        .setOption('height', 350)
        .build()
    );
    
    sheet.getRange(1, 21, kpiChartData.length, 2).setValues(kpiChartData);
  }
}

/**
 * Создаёт пустой отчёт когда нет данных
 * @private
 */
function createEmptyManagersReport_() {
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.MANAGER_PERFORMANCE);
  clearSheetData_(sheet);
  applySheetFormatting_(sheet, 'Эффективность менеджеров');
  
  sheet.getRange('A3').setValue('ℹ️ Нет данных для анализа менеджеров');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A5').setValue('Убедитесь, что:');
  sheet.getRange('A6').setValue('• В данных указаны ответственные менеджеры');
  sheet.getRange('A7').setValue('• Есть информация о статусах сделок');
  sheet.getRange('A8').setValue('• Выполнена синхронизация данных из CRM');
  
  updateLastUpdateTime_(CONFIG.SHEETS.MANAGER_PERFORMANCE);
}
