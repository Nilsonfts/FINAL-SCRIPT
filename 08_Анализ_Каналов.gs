/**
 * АНАЛИЗ КАНАЛОВ ПРИВЛЕЧЕНИЯ
 * Детальная аналитика эффективности маркетинговых каналов
 * @fileoverview Модуль анализа ROI, конверсии и эффективности каналов привлечения
 */

/**
 * Основная функция анализа эффективности каналов
 * Создаёт отчёт с метриками по всем каналам привлечения
 */
function analyzeChannelPerformance() {
  try {
    logInfo_('CHANNEL_ANALYSIS', 'Начало анализа каналов привлечения');
    
    const startTime = new Date();
    
    // Получаем данные для анализа
    const channelData = getChannelPerformanceData_();
    
    if (!channelData || channelData.allTimeData.length === 0) {
      logWarning_('CHANNEL_ANALYSIS', 'Нет данных для анализа каналов');
      createEmptyChannelReport_();
      return;
    }
    
    // Создаём или обновляем лист
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.CHANNEL_ANALYSIS);
    clearSheetData_(sheet);
    applySheetFormatting_(sheet, 'Анализ каналов привлечения');
    
    // Строим структуру отчёта
    createChannelAnalysisStructure_(sheet, channelData);
    
    // Добавляем диаграммы
    addChannelAnalysisCharts_(sheet, channelData);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('CHANNEL_ANALYSIS', `Анализ каналов завершён за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.CHANNEL_ANALYSIS);
    
  } catch (error) {
    logError_('CHANNEL_ANALYSIS', 'Ошибка анализа каналов привлечения', error);
    throw error;
  }
}

/**
 * Получает данные для анализа каналов привлечения
 * @returns {Object} Данные по каналам
 * @private
 */
function getChannelPerformanceData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return null;
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Анализируем данные за всё время
  const allTimeData = analyzeChannelDataAllTime_(headers, rows);
  
  // Анализируем данные по месяцам
  const monthlyData = analyzeChannelDataByMonth_(headers, rows);
  
  // Тренды каналов
  const channelTrends = analyzeChannelTrends_(headers, rows);
  
  // Корреляционный анализ
  const correlationData = analyzeChannelCorrelation_(headers, rows);
  
  return {
    allTimeData: allTimeData,
    monthlyData: monthlyData,
    channelTrends: channelTrends,
    correlationData: correlationData,
    totalRecords: rows.length,
    analysisDate: getCurrentDateMoscow_()
  };
}

/**
 * Анализирует данные каналов за всё время
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Данные по каналам за всё время
 * @private
 */
function analyzeChannelDataAllTime_(headers, rows) {
  const channelStats = {};
  
  // Получаем индексы важных колонок
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const campaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  
  // Анализируем каждую строку
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const createdDate = parseDate_(row[createdIndex]);
    const utmSource = row[sourceIndex] || '';
    const utmCampaign = row[campaignIndex] || '';
    
    // Инициализируем канал если нет
    if (!channelStats[channel]) {
      channelStats[channel] = {
        channel: channel,
        totalLeads: 0,
        wonDeals: 0,
        lostDeals: 0,
        inProgress: 0,
        totalRevenue: 0,
        averageRevenue: 0,
        conversion: 0,
        avgDealCycle: 0,
        topSources: {},
        topCampaigns: {},
        monthlyDistribution: {},
        costPerLead: 0, // Будет рассчитано позже из рекламных расходов
        roi: 0
      };
    }
    
    const stats = channelStats[channel];
    stats.totalLeads++;
    
    // Подсчёт по статусам
    if (status === 'success') {
      stats.wonDeals++;
      stats.totalRevenue += budget;
    } else if (status === 'failure') {
      stats.lostDeals++;
    } else {
      stats.inProgress++;
    }
    
    // Топ источники по каналу
    if (utmSource) {
      stats.topSources[utmSource] = (stats.topSources[utmSource] || 0) + 1;
    }
    
    // Топ кампании по каналу
    if (utmCampaign) {
      stats.topCampaigns[utmCampaign] = (stats.topCampaigns[utmCampaign] || 0) + 1;
    }
    
    // Распределение по месяцам
    if (createdDate) {
      const monthKey = formatDate_(createdDate, 'YYYY-MM');
      stats.monthlyDistribution[monthKey] = (stats.monthlyDistribution[monthKey] || 0) + 1;
    }
  });
  
  // Вычисляем производные метрики
  return Object.values(channelStats).map(stats => {
    stats.conversion = stats.totalLeads > 0 ? (stats.wonDeals / stats.totalLeads * 100) : 0;
    stats.averageRevenue = stats.wonDeals > 0 ? (stats.totalRevenue / stats.wonDeals) : 0;
    
    // Сортируем топы
    stats.topSources = Object.entries(stats.topSources)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 3)
      .reduce((obj, [source, count]) => {
        obj[source] = count;
        return obj;
      }, {});
    
    stats.topCampaigns = Object.entries(stats.topCampaigns)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 3)
      .reduce((obj, [campaign, count]) => {
        obj[campaign] = count;
        return obj;
      }, {});
    
    return stats;
  }).sort((a, b) => b.totalRevenue - a.totalRevenue); // Сортируем по выручке
}

/**
 * Анализирует данные каналов по месяцам
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Array} Данные по каналам по месяцам
 * @private
 */
function analyzeChannelDataByMonth_(headers, rows) {
  const monthlyStats = {};
  
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const createdDate = parseDate_(row[createdIndex]);
    
    if (!createdDate) return;
    
    const monthKey = formatDate_(createdDate, 'YYYY-MM');
    const monthName = `${getMonthName_(createdDate.getMonth())} ${createdDate.getFullYear()}`;
    
    if (!monthlyStats[monthKey]) {
      monthlyStats[monthKey] = {
        month: monthName,
        monthKey: monthKey,
        channels: {}
      };
    }
    
    if (!monthlyStats[monthKey].channels[channel]) {
      monthlyStats[monthKey].channels[channel] = {
        leads: 0,
        won: 0,
        revenue: 0,
        conversion: 0
      };
    }
    
    const channelStats = monthlyStats[monthKey].channels[channel];
    channelStats.leads++;
    
    if (status === 'success') {
      channelStats.won++;
      channelStats.revenue += budget;
    }
  });
  
  // Вычисляем конверсию и сортируем
  return Object.entries(monthlyStats)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, monthData]) => {
      // Вычисляем конверсию для каждого канала в месяце
      Object.keys(monthData.channels).forEach(channel => {
        const stats = monthData.channels[channel];
        stats.conversion = stats.leads > 0 ? (stats.won / stats.leads * 100) : 0;
      });
      
      return monthData;
    });
}

/**
 * Анализирует тренды каналов
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Трендовые данные
 * @private
 */
function analyzeChannelTrends_(headers, rows) {
  const trends = {};
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  
  // Группируем по неделям для более детального анализа трендов
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const createdDate = parseDate_(row[createdIndex]);
    const status = row[statusIndex] || '';
    
    if (!createdDate) return;
    
    // Получаем начало недели
    const weekStart = new Date(createdDate);
    weekStart.setDate(weekStart.getDate() - weekStart.getDay() + 1); // Понедельник
    const weekKey = formatDate_(weekStart, 'YYYY-MM-DD');
    
    if (!trends[channel]) {
      trends[channel] = {};
    }
    
    if (!trends[channel][weekKey]) {
      trends[channel][weekKey] = {
        week: weekKey,
        leads: 0,
        conversions: 0
      };
    }
    
    trends[channel][weekKey].leads++;
    if (status === 'success') {
      trends[channel][weekKey].conversions++;
    }
  });
  
  // Вычисляем тренды (рост/падение)
  const trendAnalysis = {};
  Object.keys(trends).forEach(channel => {
    const weekData = Object.values(trends[channel]).sort((a, b) => a.week.localeCompare(b.week));
    
    if (weekData.length >= 2) {
      const recent = weekData.slice(-4); // Последние 4 недели
      const earlier = weekData.slice(-8, -4); // Предыдущие 4 недели
      
      const recentAvg = recent.reduce((sum, week) => sum + week.leads, 0) / recent.length;
      const earlierAvg = earlier.length > 0 ? earlier.reduce((sum, week) => sum + week.leads, 0) / earlier.length : recentAvg;
      
      const growth = earlierAvg > 0 ? ((recentAvg - earlierAvg) / earlierAvg * 100) : 0;
      
      trendAnalysis[channel] = {
        recentAverage: recentAvg,
        earlierAverage: earlierAvg,
        growthPercent: growth,
        trend: growth > 5 ? 'Рост' : growth < -5 ? 'Спад' : 'Стабильно'
      };
    }
  });
  
  return trendAnalysis;
}

/**
 * Анализирует корреляцию между каналами
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Корреляционные данные
 * @private
 */
function analyzeChannelCorrelation_(headers, rows) {
  // Анализируем какие каналы часто работают вместе
  const clientChannels = {};
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const phoneIndex = findHeaderIndex_(headers, 'Телефон');
  const emailIndex = findHeaderIndex_(headers, 'Email');
  
  rows.forEach(row => {
    const channel = row[channelIndex] || 'Неизвестно';
    const phone = row[phoneIndex] || '';
    const email = row[emailIndex] || '';
    
    const clientId = phone || email;
    if (!clientId) return;
    
    if (!clientChannels[clientId]) {
      clientChannels[clientId] = new Set();
    }
    
    clientChannels[clientId].add(channel);
  });
  
  // Считаем пересечения каналов
  const channelPairs = {};
  Object.values(clientChannels).forEach(channels => {
    const channelArray = Array.from(channels);
    
    if (channelArray.length > 1) {
      for (let i = 0; i < channelArray.length; i++) {
        for (let j = i + 1; j < channelArray.length; j++) {
          const pair = [channelArray[i], channelArray[j]].sort().join(' + ');
          channelPairs[pair] = (channelPairs[pair] || 0) + 1;
        }
      }
    }
  });
  
  // Топ корреляций
  const topCorrelations = Object.entries(channelPairs)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 5)
    .map(([pair, count]) => ({ pair, count }));
  
  return {
    multiChannelClients: Object.keys(clientChannels).filter(id => clientChannels[id].size > 1).length,
    totalClients: Object.keys(clientChannels).length,
    topCorrelations: topCorrelations
  };
}

/**
 * Создаёт структуру отчёта по анализу каналов
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} channelData - Данные анализа каналов
 * @private
 */
function createChannelAnalysisStructure_(sheet, channelData) {
  let currentRow = 3;
  
  // 1. Общая статистика
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('📊 ОБЩАЯ СТАТИСТИКА КАНАЛОВ');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const generalStats = [
    ['Всего каналов:', channelData.allTimeData.length],
    ['Общий объём лидов:', channelData.totalRecords],
    ['Мультиканальных клиентов:', channelData.correlationData.multiChannelClients],
    ['Анализируемый период:', formatDate_(channelData.analysisDate, 'DD.MM.YYYY')]
  ];
  
  sheet.getRange(currentRow, 1, generalStats.length, 2).setValues(generalStats);
  sheet.getRange(currentRow, 1, generalStats.length, 1).setFontWeight('bold');
  currentRow += generalStats.length + 2;
  
  // 2. Топ каналов по эффективности
  sheet.getRange(currentRow, 1, 1, 6).merge();
  sheet.getRange(currentRow, 1).setValue('🏆 ТОП КАНАЛОВ ПО ЭФФЕКТИВНОСТИ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const topChannelsHeaders = [['Канал', 'Лиды', 'Сделки', 'Выручка', 'Конверсия', 'Средний чек']];
  sheet.getRange(currentRow, 1, 1, 6).setValues(topChannelsHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 6));
  currentRow++;
  
  const topChannelsData = channelData.allTimeData.slice(0, 10).map(item => [
    item.channel,
    item.totalLeads,
    item.wonDeals,
    formatCurrency_(item.totalRevenue),
    `${item.conversion.toFixed(1)}%`,
    formatCurrency_(item.averageRevenue)
  ]);
  
  if (topChannelsData.length > 0) {
    sheet.getRange(currentRow, 1, topChannelsData.length, 6).setValues(topChannelsData);
    currentRow += topChannelsData.length;
  }
  currentRow += 2;
  
  // 3. Детализация по топ каналам
  const topChannels = channelData.allTimeData.slice(0, 3);
  topChannels.forEach((channelStats, index) => {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue(`🔍 ДЕТАЛИЗАЦИЯ: ${channelStats.channel.toUpperCase()}`);
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    // Основные метрики канала
    const channelDetails = [
      ['Общие показатели:', ''],
      ['• Всего лидов', channelStats.totalLeads],
      ['• Успешных сделок', channelStats.wonDeals],
      ['• В работе', channelStats.inProgress],
      ['• Отклонено', channelStats.lostDeals],
      ['Финансовые показатели:', ''],
      ['• Общая выручка', formatCurrency_(channelStats.totalRevenue)],
      ['• Средний чек', formatCurrency_(channelStats.averageRevenue)],
      ['• Конверсия', `${channelStats.conversion.toFixed(1)}%`]
    ];
    
    sheet.getRange(currentRow, 1, channelDetails.length, 2).setValues(channelDetails);
    
    // Выделяем заголовки разделов
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow + 5, 1).setFontWeight('bold');
    
    currentRow += channelDetails.length;
    
    // Топ источники и кампании
    if (Object.keys(channelStats.topSources).length > 0) {
      sheet.getRange(currentRow, 1).setValue('Топ источники:');
      sheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
      
      Object.entries(channelStats.topSources).forEach(([source, count]) => {
        sheet.getRange(currentRow, 1).setValue(`• ${source}`);
        sheet.getRange(currentRow, 2).setValue(`${count} лидов`);
        currentRow++;
      });
    }
    
    currentRow += 2;
  });
  
  // 4. Трендовый анализ
  if (Object.keys(channelData.channelTrends).length > 0) {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue('📈 ТРЕНДОВЫЙ АНАЛИЗ (4 недели)');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const trendHeaders = [['Канал', 'Текущий объём', 'Тренд', 'Изменение']];
    sheet.getRange(currentRow, 1, 1, 4).setValues(trendHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const trendData = Object.entries(channelData.channelTrends)
      .sort(([,a], [,b]) => Math.abs(b.growthPercent) - Math.abs(a.growthPercent))
      .slice(0, 8)
      .map(([channel, trend]) => [
        channel,
        Math.round(trend.recentAverage),
        trend.trend,
        trend.growthPercent > 0 ? `+${trend.growthPercent.toFixed(1)}%` : `${trend.growthPercent.toFixed(1)}%`
      ]);
    
    if (trendData.length > 0) {
      sheet.getRange(currentRow, 1, trendData.length, 4).setValues(trendData);
      
      // Цветовая индикация трендов
      trendData.forEach((row, index) => {
        const trendCell = sheet.getRange(currentRow + index, 3);
        const changeCell = sheet.getRange(currentRow + index, 4);
        
        if (row[2] === 'Рост') {
          trendCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
          changeCell.setBackground(CONFIG.COLORS.SUCCESS_LIGHT);
        } else if (row[2] === 'Спад') {
          trendCell.setBackground(CONFIG.COLORS.ERROR_LIGHT);
          changeCell.setBackground(CONFIG.COLORS.ERROR_LIGHT);
        }
      });
      
      currentRow += trendData.length;
    }
    currentRow += 2;
  }
  
  // 5. Корреляция каналов
  if (channelData.correlationData.topCorrelations.length > 0) {
    sheet.getRange(currentRow, 1, 1, 3).merge();
    sheet.getRange(currentRow, 1).setValue('🔗 ВЗАИМОДЕЙСТВИЕ КАНАЛОВ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
    currentRow++;
    
    sheet.getRange(currentRow, 1).setValue(`Клиентов с несколькими каналами: ${channelData.correlationData.multiChannelClients} из ${channelData.correlationData.totalClients}`);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow += 2;
    
    const correlationHeaders = [['Комбинация каналов', 'Клиентов', 'Доля']];
    sheet.getRange(currentRow, 1, 1, 3).setValues(correlationHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
    currentRow++;
    
    const correlationData = channelData.correlationData.topCorrelations.map(item => [
      item.pair,
      item.count,
      `${(item.count / channelData.correlationData.totalClients * 100).toFixed(1)}%`
    ]);
    
    if (correlationData.length > 0) {
      sheet.getRange(currentRow, 1, correlationData.length, 3).setValues(correlationData);
      currentRow += correlationData.length;
    }
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 6);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 6);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к анализу каналов
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} channelData - Данные анализа каналов
 * @private
 */
function addChannelAnalysisCharts_(sheet, channelData) {
  // 1. Круговая диаграмма выручки по каналам
  if (channelData.allTimeData.length > 0) {
    const revenueChartData = [['Канал', 'Выручка']].concat(
      channelData.allTimeData.slice(0, 8).map(item => [item.channel, item.totalRevenue])
    );
    
    // Создаем диаграмму через универсальную функцию
    const revenueChart = createChart_(sheet, 'pie', revenueChartData, {
      startRow: 1,
      startCol: 8,
      title: 'Распределение выручки по каналам (за всё время)',
      position: { row: 3, col: 8 },
      width: 500,
      height: 350,
      legend: 'right'
    });
  }
  
  // 2. Столбчатая диаграмма конверсии по каналам
  if (channelData.allTimeData.length > 0) {
    const conversionChartData = [['Канал', 'Конверсия %']].concat(
      channelData.allTimeData.slice(0, 10).map(item => [item.channel, item.conversion])
    );
    
    // Создаем диаграмму через универсальную функцию
    const conversionChart = createChart_(sheet, 'column', conversionChartData, {
      startRow: 1,
      startCol: 11,
      title: 'Конверсия по каналам',
      position: { row: 3, col: 14 },
      width: 600,
      height: 350,
      legend: 'none',
      hAxisTitle: 'Канал',
      vAxisTitle: 'Конверсия, %'
    });
  }
  
  // 3. Динамика топ каналов по месяцам
  if (channelData.monthlyData.length > 0) {
    // Берём топ-5 каналов для динамики
    const topChannelNames = channelData.allTimeData.slice(0, 5).map(item => item.channel);
    
    // Подготавливаем данные для линейной диаграммы
    const monthlyChartHeaders = ['Месяц'].concat(topChannelNames);
    const monthlyChartData = [monthlyChartHeaders];
    
    channelData.monthlyData.forEach(monthData => {
      const row = [monthData.month];
      
      topChannelNames.forEach(channelName => {
        const channelStats = monthData.channels[channelName];
        row.push(channelStats ? channelStats.leads : 0);
      });
      
      monthlyChartData.push(row);
    });
    
    // Создаем диаграмму через универсальную функцию
    const monthlyChart = createChart_(sheet, 'line', monthlyChartData, {
      startRow: 1,
      startCol: 15,
      title: 'Динамика топ каналов по месяцам',
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
function createEmptyChannelReport_() {
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.CHANNEL_ANALYSIS);
  clearSheetData_(sheet);
  applySheetFormatting_(sheet, 'Анализ каналов привлечения');
  
  sheet.getRange('A3').setValue('ℹ️ Нет данных для анализа каналов привлечения');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A5').setValue('Убедитесь, что:');
  sheet.getRange('A6').setValue('• В системе есть данные о лидах с указанием каналов');
  sheet.getRange('A7').setValue('• Выполнена синхронизация данных');
  sheet.getRange('A8').setValue('• Каналы корректно определяются по UTM меткам');
  
  updateLastUpdateTime_(CONFIG.SHEETS.CHANNEL_ANALYSIS);
}
