/**
 * СВОДНАЯ АНАЛИТИКА AMOCRM
 * Общие показатели по лидам, конверсии, доходам из AmoCRM
 * @fileoverview Главный дашборд с ключевыми метриками AmoCRM
 */

/**
 * Основная функция обновления сводной аналитики AmoCRM
 * Создаёт итоговую таблицу с ключевыми метриками
 */
function updateAmoCrmSummary() {
  try {
    logInfo_('AMOCRM_SUMMARY', 'Начало обновления сводной аналитики AmoCRM');
    
    const startTime = new Date();
    
    // Получаем данные за всё время
    const allTimeData = getAmoCrmAnalyticsData_();
    
    // Получаем данные по месяцам
    const monthlyData = getAmoCrmMonthlyData_();
    
    // Создаём или обновляем лист
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.AMOCRM_SUMMARY);
    
    // Очищаем лист
    clearSheetData_(sheet);
    
    // Применяем стандартные настройки
    applySheetFormatting_(sheet, 'AmoCRM: Сводная аналитика');
    
    // Создаём структуру отчёта
    createAmoCrmSummaryStructure_(sheet, allTimeData, monthlyData);
    
    // Добавляем диаграммы
    addAmoCrmSummaryCharts_(sheet, allTimeData, monthlyData);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('AMOCRM_SUMMARY', `Сводная аналитика AmoCRM обновлена за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.AMOCRM_SUMMARY);
    
  } catch (error) {
    logError_('AMOCRM_SUMMARY', 'Ошибка обновления сводной аналитики AmoCRM', error);
    throw error;
  }
}

/**
 * Получает данные для аналитики AmoCRM за весь период
 * @returns {Object} Аналитические данные
 * @private
 */
function getAmoCrmAnalyticsData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) {
    return createEmptyAmoCrmData_();
  }
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Фильтруем только данные из AmoCRM
  const amocrmRows = rows.filter(row => {
    const sourceIndex = findHeaderIndex_(headers, 'Источник данных');
    return sourceIndex >= 0 && row[sourceIndex] === 'AmoCRM';
  });
  
  if (amocrmRows.length === 0) {
    return createEmptyAmoCrmData_();
  }
  
  return analyzeAmoCrmData_(headers, amocrmRows);
}

/**
 * Анализирует данные AmoCRM и возвращает метрики
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Аналитические метрики
 * @private
 */
function analyzeAmoCrmData_(headers, rows) {
  const metrics = {
    // Основные показатели
    totalLeads: 0,
    wonDeals: 0,
    lostDeals: 0,
    inProgressDeals: 0,
    totalRevenue: 0,
    averageDealSize: 0,
    conversionRate: 0,
    
    // Разбивка по каналам
    channelStats: {},
    
    // Разбивка по статусам
    statusStats: {},
    
    // Разбивка по ответственным
    managerStats: {},
    
    // Временная динамика
    dailyStats: {},
    
    // UTM аналитика
    utmStats: {
      sources: {},
      mediums: {},
      campaigns: {}
    },
    
    // Топы
    topChannels: [],
    topManagers: [],
    topCampaigns: []
  };
  
  // Получаем индексы важных колонок
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const managerIndex = findHeaderIndex_(headers, 'Ответственный');
  const utmSourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const utmMediumIndex = findHeaderIndex_(headers, 'UTM Medium');
  const utmCampaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  
  // Анализируем каждую строку
  rows.forEach(row => {
    metrics.totalLeads++;
    
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const channel = row[channelIndex] || 'Неизвестно';
    const createdDate = parseDate_(row[createdIndex]);
    const manager = row[managerIndex] || 'Неназначен';
    const utmSource = row[utmSourceIndex] || '';
    const utmMedium = row[utmMediumIndex] || '';
    const utmCampaign = row[utmCampaignIndex] || '';
    
    // Подсчёт по статусам
    if (status === 'success') {
      metrics.wonDeals++;
      metrics.totalRevenue += budget;
    } else if (status === 'failure') {
      metrics.lostDeals++;
    } else {
      metrics.inProgressDeals++;
    }
    
    // Статистика по статусам
    metrics.statusStats[status] = (metrics.statusStats[status] || 0) + 1;
    
    // Статистика по каналам
    if (!metrics.channelStats[channel]) {
      metrics.channelStats[channel] = {
        leads: 0,
        won: 0,
        revenue: 0,
        conversion: 0
      };
    }
    metrics.channelStats[channel].leads++;
    if (status === 'success') {
      metrics.channelStats[channel].won++;
      metrics.channelStats[channel].revenue += budget;
    }
    
    // Статистика по менеджерам
    if (!metrics.managerStats[manager]) {
      metrics.managerStats[manager] = {
        leads: 0,
        won: 0,
        revenue: 0,
        conversion: 0
      };
    }
    metrics.managerStats[manager].leads++;
    if (status === 'success') {
      metrics.managerStats[manager].won++;
      metrics.managerStats[manager].revenue += budget;
    }
    
    // Временная статистика
    if (createdDate) {
      const dateKey = formatDate_(createdDate, 'YYYY-MM-DD');
      if (!metrics.dailyStats[dateKey]) {
        metrics.dailyStats[dateKey] = {
          leads: 0,
          won: 0,
          revenue: 0
        };
      }
      metrics.dailyStats[dateKey].leads++;
      if (status === 'success') {
        metrics.dailyStats[dateKey].won++;
        metrics.dailyStats[dateKey].revenue += budget;
      }
    }
    
    // UTM статистика
    if (utmSource) {
      metrics.utmStats.sources[utmSource] = (metrics.utmStats.sources[utmSource] || 0) + 1;
    }
    if (utmMedium) {
      metrics.utmStats.mediums[utmMedium] = (metrics.utmStats.mediums[utmMedium] || 0) + 1;
    }
    if (utmCampaign) {
      metrics.utmStats.campaigns[utmCampaign] = (metrics.utmStats.campaigns[utmCampaign] || 0) + 1;
    }
  });
  
  // Вычисляем производные метрики
  metrics.averageDealSize = metrics.wonDeals > 0 ? metrics.totalRevenue / metrics.wonDeals : 0;
  metrics.conversionRate = metrics.totalLeads > 0 ? (metrics.wonDeals / metrics.totalLeads * 100) : 0;
  
  // Вычисляем конверсию по каналам
  Object.keys(metrics.channelStats).forEach(channel => {
    const stats = metrics.channelStats[channel];
    stats.conversion = stats.leads > 0 ? (stats.won / stats.leads * 100) : 0;
  });
  
  // Вычисляем конверсию по менеджерам
  Object.keys(metrics.managerStats).forEach(manager => {
    const stats = metrics.managerStats[manager];
    stats.conversion = stats.leads > 0 ? (stats.won / stats.leads * 100) : 0;
  });
  
  // Создаём топы
  metrics.topChannels = Object.entries(metrics.channelStats)
    .sort(([,a], [,b]) => b.revenue - a.revenue)
    .slice(0, 10)
    .map(([channel, stats]) => ({ channel, ...stats }));
  
  metrics.topManagers = Object.entries(metrics.managerStats)
    .sort(([,a], [,b]) => b.revenue - a.revenue)
    .slice(0, 10)
    .map(([manager, stats]) => ({ manager, ...stats }));
  
  metrics.topCampaigns = Object.entries(metrics.utmStats.campaigns)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 10)
    .map(([campaign, leads]) => ({ campaign, leads }));
  
  return metrics;
}

/**
 * Получает данные AmoCRM с разбивкой по месяцам
 * @returns {Array} Данные по месяцам
 * @private
 */
function getAmoCrmMonthlyData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) return [];
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return [];
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Фильтруем только данные из AmoCRM
  const amocrmRows = rows.filter(row => {
    const sourceIndex = findHeaderIndex_(headers, 'Источник данных');
    return sourceIndex >= 0 && row[sourceIndex] === 'AmoCRM';
  });
  
  // Группируем по месяцам
  const monthlyStats = {};
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  
  amocrmRows.forEach(row => {
    const createdDate = parseDate_(row[createdIndex]);
    if (!createdDate) return;
    
    const monthKey = formatDate_(createdDate, 'YYYY-MM');
    const monthName = `${getMonthName_(createdDate.getMonth())} ${createdDate.getFullYear()}`;
    
    if (!monthlyStats[monthKey]) {
      monthlyStats[monthKey] = {
        month: monthName,
        leads: 0,
        won: 0,
        revenue: 0,
        conversion: 0
      };
    }
    
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    
    monthlyStats[monthKey].leads++;
    if (status === 'success') {
      monthlyStats[monthKey].won++;
      monthlyStats[monthKey].revenue += budget;
    }
  });
  
  // Вычисляем конверсию и сортируем по дате
  return Object.entries(monthlyStats)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, stats]) => {
      stats.conversion = stats.leads > 0 ? (stats.won / stats.leads * 100) : 0;
      return stats;
    });
}

/**
 * Создаёт структуру сводного отчёта AmoCRM
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} allTimeData - Данные за всё время
 * @param {Array} monthlyData - Данные по месяцам
 * @private
 */
function createAmoCrmSummaryStructure_(sheet, allTimeData, monthlyData) {
  let currentRow = 3; // Начинаем после заголовка
  
  // 1. Общие показатели за всё время
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('📊 ОБЩИЕ ПОКАЗАТЕЛИ ЗА ВСЁ ВРЕМЯ');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const allTimeMetrics = [
    ['Всего лидов:', allTimeData.totalLeads],
    ['Успешных сделок:', allTimeData.wonDeals],
    ['Конверсия:', `${allTimeData.conversionRate.toFixed(1)}%`],
    ['Общая выручка:', `${formatCurrency_(allTimeData.totalRevenue)}`],
    ['Средний чек:', `${formatCurrency_(allTimeData.averageDealSize)}`],
    ['В работе:', allTimeData.inProgressDeals],
    ['Проваленных:', allTimeData.lostDeals]
  ];
  
  sheet.getRange(currentRow, 1, allTimeMetrics.length, 2).setValues(allTimeMetrics);
  sheet.getRange(currentRow, 1, allTimeMetrics.length, 1).setFontWeight('bold');
  currentRow += allTimeMetrics.length + 2;
  
  // 2. Топ каналов
  sheet.getRange(currentRow, 1, 1, 5).merge();
  sheet.getRange(currentRow, 1).setValue('🎯 ТОП-5 КАНАЛОВ ПО ВЫРУЧКЕ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
  currentRow++;
  
  const channelHeaders = [['Канал', 'Лиды', 'Сделки', 'Выручка', 'Конверсия']];
  sheet.getRange(currentRow, 1, 1, 5).setValues(channelHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
  currentRow++;
  
  const topChannelsData = allTimeData.topChannels.slice(0, 5).map(item => [
    item.channel,
    item.leads,
    item.won,
    formatCurrency_(item.revenue),
    `${item.conversion.toFixed(1)}%`
  ]);
  
  if (topChannelsData.length > 0) {
    sheet.getRange(currentRow, 1, topChannelsData.length, 5).setValues(topChannelsData);
    currentRow += topChannelsData.length;
  }
  currentRow += 2;
  
  // 3. Динамика по месяцам (таблица)
  sheet.getRange(currentRow, 1, 1, 5).merge();
  sheet.getRange(currentRow, 1).setValue('📈 ДИНАМИКА ПО МЕСЯЦАМ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
  currentRow++;
  
  const monthlyHeaders = [['Месяц', 'Лиды', 'Сделки', 'Выручка', 'Конверсия']];
  sheet.getRange(currentRow, 1, 1, 5).setValues(monthlyHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 5));
  currentRow++;
  
  if (monthlyData.length > 0) {
    const monthlyTableData = monthlyData.map(item => [
      item.month,
      item.leads,
      item.won,
      formatCurrency_(item.revenue),
      `${item.conversion.toFixed(1)}%`
    ]);
    
    sheet.getRange(currentRow, 1, monthlyTableData.length, 5).setValues(monthlyTableData);
    currentRow += monthlyTableData.length;
  }
  currentRow += 2;
  
  // 4. Статистика по UTM
  sheet.getRange(currentRow, 1, 1, 3).merge();
  sheet.getRange(currentRow, 1).setValue('🔗 ТОП UTM ИСТОЧНИКИ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  if (Object.keys(allTimeData.utmStats.sources).length > 0) {
    const topSources = Object.entries(allTimeData.utmStats.sources)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 5)
      .map(([source, count]) => [source, count, `${(count / allTimeData.totalLeads * 100).toFixed(1)}%`]);
    
    const utmHeaders = [['UTM Source', 'Лиды', 'Доля']];
    sheet.getRange(currentRow, 1, 1, 3).setValues(utmHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
    currentRow++;
    
    sheet.getRange(currentRow, 1, topSources.length, 3).setValues(topSources);
    currentRow += topSources.length;
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 5);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 5);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к сводному отчёту AmoCRM
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} allTimeData - Данные за всё время
 * @param {Array} monthlyData - Данные по месяцам
 * @private
 */
function addAmoCrmSummaryCharts_(sheet, allTimeData, monthlyData) {
  // 1. Круговая диаграмма по каналам (выручка)
  if (allTimeData.topChannels.length > 0) {
    const channelChartData = [['Канал', 'Выручка']].concat(
      allTimeData.topChannels.slice(0, 8).map(item => [item.channel, item.revenue])
    );
    
    const channelChart = sheet.insertChart(
      Charts.newPieChart()
        .setDataRange(sheet.getRange(1, 7, channelChartData.length, 2))
        .setOption('title', 'Выручка по каналам (за всё время)')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'right' })
        .setOption('chartArea', { width: '80%', height: '80%' })
        .setPosition(3, 7, 0, 0)
        .setOption('width', 500)
        .setOption('height', 350)
        .build()
    );
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 7, channelChartData.length, 2).setValues(channelChartData);
  }
  
  // 2. Линейная диаграмма динамики по месяцам
  if (monthlyData.length > 0) {
    const monthlyChartData = [['Месяц', 'Лиды', 'Сделки', 'Выручка']].concat(
      monthlyData.map(item => [item.month, item.leads, item.won, item.revenue])
    );
    
    const monthlyChart = sheet.insertChart(
      Charts.newColumnChart()
        .setDataRange(sheet.getRange(1, 10, monthlyChartData.length, 4))
        .setOption('title', 'Динамика показателей по месяцам')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'top' })
        .setOption('hAxis', { title: 'Месяц', slantedText: true })
        .setOption('vAxes', {
          0: { title: 'Количество' },
          1: { title: 'Выручка, руб.' }
        })
        .setOption('series', {
          2: { targetAxisIndex: 1, type: 'line' }  // Выручка на правой оси
        })
        .setPosition(20, 7, 0, 0)
        .setOption('width', 600)
        .setOption('height', 400)
        .build()
    );
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 10, monthlyChartData.length, 4).setValues(monthlyChartData);
  }
  
  // 3. Столбчатая диаграмма конверсии по каналам
  if (allTimeData.topChannels.length > 0) {
    const conversionChartData = [['Канал', 'Конверсия %']].concat(
      allTimeData.topChannels.slice(0, 8).map(item => [item.channel, item.conversion])
    );
    
    const conversionChart = sheet.insertChart(
      Charts.newColumnChart()
        .setDataRange(sheet.getRange(1, 15, conversionChartData.length, 2))
        .setOption('title', 'Конверсия по каналам')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Канал', slantedText: true })
        .setOption('vAxis', { title: 'Конверсия, %' })
        .setPosition(45, 7, 0, 0)
        .setOption('width', 500)
        .setOption('height', 350)
        .build()
    );
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 15, conversionChartData.length, 2).setValues(conversionChartData);
  }
}

/**
 * Создаёт пустую структуру данных AmoCRM
 * @returns {Object} Пустые данные
 * @private
 */
function createEmptyAmoCrmData_() {
  return {
    totalLeads: 0,
    wonDeals: 0,
    lostDeals: 0,
    inProgressDeals: 0,
    totalRevenue: 0,
    averageDealSize: 0,
    conversionRate: 0,
    channelStats: {},
    statusStats: {},
    managerStats: {},
    dailyStats: {},
    utmStats: { sources: {}, mediums: {}, campaigns: {} },
    topChannels: [],
    topManagers: [],
    topCampaigns: []
  };
}

/**
 * Форматирует число как валюту
 * @param {number} amount - Сумма
 * @returns {string} Отформатированная сумма
 * @private
 */
function formatCurrency_(amount) {
  if (!amount || isNaN(amount)) return '0 ₽';
  
  return Math.round(amount).toLocaleString('ru-RU') + ' ₽';
}
