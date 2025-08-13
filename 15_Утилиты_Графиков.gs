/**
 * МОДУЛЬ: УТИЛИТЫ ВИЗУАЛИЗАЦИИ И ГРАФИКОВ
 * Создание диаграмм и графиков для всех модулей
 */

function createCharts() {
  console.log('Создаем графики и диаграммы...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для создания графиков');
      return;
    }
    
    // Создаем разные типы графиков
    createChannelDistributionChart(data);
    createConversionFunnelChart(data);
    createTimeAnalysisChart(data);
    createManagerPerformanceChart(data);
    createRevenueAnalysisChart(data);
    
    console.log('Графики созданы успешно');
    
  } catch (error) {
    console.error('Ошибка при создании графиков:', error);
    throw error;
  }
}

function createChannelDistributionChart(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LEADS_CHANNELS);
    if (!sheet) return;
    
    // Собираем данные для диаграммы
    const channelData = {};
    
    data.forEach(row => {
      const source = parseUtmSource(row);
      const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
      const isSuccess = isSuccessStatus(status);
      
      if (!channelData[source]) {
        channelData[source] = { total: 0, success: 0 };
      }
      
      channelData[source].total++;
      if (isSuccess) channelData[source].success++;
    });
    
    // Сортируем по количеству лидов
    const sortedChannels = Object.entries(channelData)
      .sort((a, b) => b[1].total - a[1].total)
      .slice(0, 10);
    
    if (sortedChannels.length === 0) return;
    
    // Находим свободное место для графика
    const lastRow = sheet.getLastRow();
    const chartStartRow = lastRow + 3;
    
    // Создаем данные для графика
    const chartHeaders = ['Источник', 'Всего лидов', 'Успешных'];
    sheet.getRange(chartStartRow, 1, 1, 3).setValues([chartHeaders]);
    
    const chartData = sortedChannels.map(([channel, data]) => [
      channel,
      data.total,
      data.success
    ]);
    
    sheet.getRange(chartStartRow + 1, 1, chartData.length, 3).setValues(chartData);
    
    // Создаем диаграмму
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(chartStartRow, 1, chartData.length + 1, 3))
      .setPosition(chartStartRow, 5, 0, 0)
      .setOption('title', 'Распределение лидов по каналам')
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('colors', CONFIG.COLORS.chart_palette)
      .build();
    
    sheet.insertChart(chart);
    
  } catch (error) {
    console.warn('Не удалось создать диаграмму распределения каналов:', error);
  }
}

function createConversionFunnelChart(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SITE_CROSS);
    if (!sheet) return;
    
    // Анализируем воронку конверсии
    let totalVisits = 0;
    let totalForms = 0;
    let totalLeads = 0;
    let totalSuccess = 0;
    
    data.forEach(row => {
      const formName = row[CONFIG.WORKING_AMO_COLUMNS.FORMNAME] || '';
      const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
      const isSuccess = isSuccessStatus(status);
      
      totalVisits++;
      if (formName && formName !== '') totalForms++;
      totalLeads++;
      if (isSuccess) totalSuccess++;
    });
    
    const lastRow = sheet.getLastRow();
    const chartStartRow = lastRow + 3;
    
    // Данные воронки
    const funnelData = [
      ['Этап', 'Количество', 'Конверсия %'],
      ['Визиты', totalVisits, 100],
      ['Заполнение форм', totalForms, totalVisits > 0 ? (totalForms / totalVisits * 100).toFixed(1) : 0],
      ['Лиды', totalLeads, totalVisits > 0 ? (totalLeads / totalVisits * 100).toFixed(1) : 0],
      ['Успешные', totalSuccess, totalVisits > 0 ? (totalSuccess / totalVisits * 100).toFixed(1) : 0]
    ];
    
    sheet.getRange(chartStartRow, 1, funnelData.length, 3).setValues(funnelData);
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(chartStartRow, 1, funnelData.length, 3))
      .setPosition(chartStartRow, 5, 0, 0)
      .setOption('title', 'Воронка конверсии')
      .setOption('width', 500)
      .setOption('height', 350)
      .setOption('colors', ['#4285f4', '#ea4335'])
      .build();
    
    sheet.insertChart(chart);
    
  } catch (error) {
    console.warn('Не удалось создать диаграмму воронки конверсии:', error);
  }
}

function createTimeAnalysisChart(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.FIRST_TOUCH);
    if (!sheet) return;
    
    // Анализ по времени (часам)
    const timeData = {};
    
    data.forEach(row => {
      const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
      const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
      const isSuccess = isSuccessStatus(status);
      
      if (createdAt) {
        const hour = new Date(createdAt).getHours();
        if (!timeData[hour]) {
          timeData[hour] = { total: 0, success: 0 };
        }
        timeData[hour].total++;
        if (isSuccess) timeData[hour].success++;
      }
    });
    
    const lastRow = sheet.getLastRow();
    const chartStartRow = lastRow + 3;
    
    // Сортируем по часам
    const sortedHours = Object.entries(timeData)
      .sort((a, b) => parseInt(a[0]) - parseInt(b[0]));
    
    if (sortedHours.length === 0) return;
    
    const timeChartData = [['Час', 'Лиды', 'Успешные']];
    sortedHours.forEach(([hour, data]) => {
      timeChartData.push([
        `${hour}:00`,
        data.total,
        data.success
      ]);
    });
    
    sheet.getRange(chartStartRow, 1, timeChartData.length, 3).setValues(timeChartData);
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(chartStartRow, 1, timeChartData.length, 3))
      .setPosition(chartStartRow, 5, 0, 0)
      .setOption('title', 'Активность лидов по часам')
      .setOption('width', 600)
      .setOption('height', 300)
      .setOption('curveType', 'smooth')
      .setOption('colors', ['#4285f4', '#34a853'])
      .build();
    
    sheet.insertChart(chart);
    
  } catch (error) {
    console.warn('Не удалось создать временную диаграмму:', error);
  }
}

function createManagerPerformanceChart(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.AMO_SUMMARY);
    if (!sheet) return;
    
    // Анализ производительности менеджеров
    const managerData = {};
    
    data.forEach(row => {
      const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
      const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
      const isSuccess = isSuccessStatus(status);
      
      if (responsible && responsible !== '') {
        if (!managerData[responsible]) {
          managerData[responsible] = { total: 0, success: 0 };
        }
        managerData[responsible].total++;
        if (isSuccess) managerData[responsible].success++;
      }
    });
    
    // Топ-10 менеджеров по количеству лидов
    const topManagers = Object.entries(managerData)
      .sort((a, b) => b[1].total - a[1].total)
      .slice(0, 10);
    
    if (topManagers.length === 0) return;
    
    const lastRow = sheet.getLastRow();
    const chartStartRow = lastRow + 3;
    
    const managerChartData = [['Менеджер', 'Всего лидов', 'Конверсия %']];
    topManagers.forEach(([manager, data]) => {
      const conversionRate = data.total > 0 ? (data.success / data.total * 100) : 0;
      managerChartData.push([
        manager.length > 15 ? manager.substring(0, 15) + '...' : manager,
        data.total,
        conversionRate.toFixed(1)
      ]);
    });
    
    sheet.getRange(chartStartRow, 1, managerChartData.length, 3).setValues(managerChartData);
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(sheet.getRange(chartStartRow, 1, managerChartData.length, 3))
      .setPosition(chartStartRow, 5, 0, 0)
      .setOption('title', 'Производительность менеджеров')
      .setOption('width', 700)
      .setOption('height', 400)
      .setOption('series', {
        0: { type: 'columns', color: '#4285f4' },
        1: { type: 'line', color: '#ea4335', targetAxisIndex: 1 }
      })
      .setOption('vAxes', {
        0: { title: 'Количество лидов' },
        1: { title: 'Конверсия %', minValue: 0, maxValue: 100 }
      })
      .build();
    
    sheet.insertChart(chart);
    
  } catch (error) {
    console.warn('Не удалось создать диаграмму производительности менеджеров:', error);
  }
}

function createRevenueAnalysisChart(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.MONTHLY_DASHBOARD);
    if (!sheet) return;
    
    // Анализ выручки по месяцам
    const monthlyRevenue = {};
    
    data.forEach(row => {
      const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
      const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
      const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
      const isSuccess = isSuccessStatus(status);
      
      if (createdAt && isSuccess && factAmount > 0) {
        const periods = getDatePeriods(createdAt);
        const monthKey = periods.yearMonth;
        
        if (!monthlyRevenue[monthKey]) {
          monthlyRevenue[monthKey] = 0;
        }
        monthlyRevenue[monthKey] += factAmount;
      }
    });
    
    // Сортируем по месяцам
    const sortedMonths = Object.entries(monthlyRevenue)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .slice(-12); // Последние 12 месяцев
    
    if (sortedMonths.length === 0) return;
    
    const lastRow = sheet.getLastRow();
    const chartStartRow = lastRow + 3;
    
    const revenueChartData = [['Месяц', 'Выручка']];
    sortedMonths.forEach(([month, revenue]) => {
      revenueChartData.push([month, revenue]);
    });
    
    sheet.getRange(chartStartRow, 1, revenueChartData.length, 2).setValues(revenueChartData);
    
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.AREA)
      .addRange(sheet.getRange(chartStartRow, 1, revenueChartData.length, 2))
      .setPosition(chartStartRow, 4, 0, 0)
      .setOption('title', 'Динамика выручки по месяцам')
      .setOption('width', 600)
      .setOption('height', 300)
      .setOption('colors', ['#34a853'])
      .setOption('isStacked', false)
      .build();
    
    sheet.insertChart(chart);
    
  } catch (error) {
    console.warn('Не удалось создать диаграмму выручки:', error);
  }
}

/**
 * ДОПОЛНИТЕЛЬНЫЕ УТИЛИТЫ ДЛЯ СОЗДАНИЯ ГРАФИКОВ
 */

function createPieChart(sheet, dataRange, title, position) {
  try {
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dataRange)
      .setPosition(position.row, position.col, 0, 0)
      .setOption('title', title)
      .setOption('width', 400)
      .setOption('height', 300)
      .setOption('colors', CONFIG.COLORS.chart_palette)
      .setOption('pieSliceText', 'percentage')
      .build();
    
    sheet.insertChart(chart);
    return chart;
    
  } catch (error) {
    console.warn(`Не удалось создать круговую диаграмму "${title}":`, error);
    return null;
  }
}

function createBarChart(sheet, dataRange, title, position, options = {}) {
  try {
    let chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dataRange)
      .setPosition(position.row, position.col, 0, 0)
      .setOption('title', title)
      .setOption('width', options.width || 500)
      .setOption('height', options.height || 300)
      .setOption('colors', options.colors || CONFIG.COLORS.chart_palette);
    
    if (options.horizontal) {
      chartBuilder = chartBuilder.setOption('orientation', 'horizontal');
    }
    
    const chart = chartBuilder.build();
    sheet.insertChart(chart);
    return chart;
    
  } catch (error) {
    console.warn(`Не удалось создать столбчатую диаграмму "${title}":`, error);
    return null;
  }
}

function createLineChart(sheet, dataRange, title, position, options = {}) {
  try {
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(position.row, position.col, 0, 0)
      .setOption('title', title)
      .setOption('width', options.width || 600)
      .setOption('height', options.height || 300)
      .setOption('curveType', options.smooth ? 'smooth' : 'normal')
      .setOption('colors', options.colors || CONFIG.COLORS.chart_palette)
      .setOption('pointSize', options.pointSize || 5)
      .build();
    
    sheet.insertChart(chart);
    return chart;
    
  } catch (error) {
    console.warn(`Не удалось создать линейную диаграмму "${title}":`, error);
    return null;
  }
}

function createScatterChart(sheet, dataRange, title, position, options = {}) {
  try {
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.SCATTER)
      .addRange(dataRange)
      .setPosition(position.row, position.col, 0, 0)
      .setOption('title', title)
      .setOption('width', options.width || 500)
      .setOption('height', options.height || 400)
      .setOption('colors', options.colors || CONFIG.COLORS.chart_palette)
      .setOption('pointSize', options.pointSize || 8)
      .build();
    
    sheet.insertChart(chart);
    return chart;
    
  } catch (error) {
    console.warn(`Не удалось создать точечную диаграмму "${title}":`, error);
    return null;
  }
}

/**
 * УТИЛИТА ДЛЯ СОЗДАНИЯ ДАШБОРДА С МНОЖЕСТВЕННЫМИ ГРАФИКАМИ
 */
function createVisualizationDashboard() {
  console.log('Создаем дашборд визуализации...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) return;
    
    const sheet = createOrUpdateSheet('📊 ВИЗУАЛИЗАЦИЯ');
    
    // Очищаем существующие графики
    const charts = sheet.getCharts();
    charts.forEach(chart => sheet.removeChart(chart));
    
    let currentRow = 1;
    let currentCol = 1;
    
    // 1. Распределение лидов по источникам (круговая)
    const sourceData = analyzeSourceDistribution(data);
    if (sourceData.length > 0) {
      createSourcePieChart(sheet, sourceData, { row: currentRow, col: currentCol });
      currentCol += 6;
    }
    
    // 2. Конверсия по каналам (столбчатая)
    const channelConversion = analyzeChannelConversion(data);
    if (channelConversion.length > 0) {
      createChannelConversionChart(sheet, channelConversion, { row: currentRow, col: currentCol });
      currentRow += 20;
      currentCol = 1;
    }
    
    // 3. Временные тренды (линейная)
    const timeData = analyzeTimeTrends(data);
    if (timeData.length > 0) {
      createTimeTrendsChart(sheet, timeData, { row: currentRow, col: currentCol });
      currentCol += 6;
    }
    
    // 4. Производительность менеджеров (комбинированная)
    const managerData = analyzeManagerPerformance(data);
    if (managerData.length > 0) {
      createManagerChart(sheet, managerData, { row: currentRow, col: currentCol });
    }
    
    console.log('Дашборд визуализации создан');
    
  } catch (error) {
    console.error('Ошибка при создании дашборда визуализации:', error);
  }
}

function analyzeSourceDistribution(data) {
  const sources = {};
  
  data.forEach(row => {
    const source = parseUtmSource(row);
    sources[source] = (sources[source] || 0) + 1;
  });
  
  return Object.entries(sources)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .map(([source, count]) => [source, count]);
}

function analyzeChannelConversion(data) {
  const channels = {};
  
  data.forEach(row => {
    const channelType = getChannelType(parseUtmSource(row));
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const isSuccess = isSuccessStatus(status);
    
    if (!channels[channelType]) {
      channels[channelType] = { total: 0, success: 0 };
    }
    
    channels[channelType].total++;
    if (isSuccess) channels[channelType].success++;
  });
  
  return Object.entries(channels)
    .map(([channel, data]) => [
      channel,
      data.total,
      data.total > 0 ? (data.success / data.total * 100) : 0
    ])
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
}

function analyzeTimeTrends(data) {
  const trends = {};
  
  data.forEach(row => {
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    if (!createdAt) return;
    
    const periods = getDatePeriods(createdAt);
    const monthKey = periods.yearMonth;
    
    if (!trends[monthKey]) {
      trends[monthKey] = { total: 0, success: 0 };
    }
    
    trends[monthKey].total++;
    if (isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '')) {
      trends[monthKey].success++;
    }
  });
  
  return Object.entries(trends)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .slice(-12)
    .map(([month, data]) => [
      month,
      data.total,
      data.success,
      data.total > 0 ? (data.success / data.total * 100) : 0
    ]);
}

function analyzeManagerPerformance(data) {
  const managers = {};
  
  data.forEach(row => {
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    if (!responsible) return;
    
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const isSuccess = isSuccessStatus(status);
    
    if (!managers[responsible]) {
      managers[responsible] = { total: 0, success: 0 };
    }
    
    managers[responsible].total++;
    if (isSuccess) managers[responsible].success++;
  });
  
  return Object.entries(managers)
    .filter(([_, data]) => data.total >= 5) // Минимум 5 лидов
    .map(([manager, data]) => [
      manager.length > 20 ? manager.substring(0, 20) + '...' : manager,
      data.total,
      data.total > 0 ? (data.success / data.total * 100) : 0
    ])
    .sort((a, b) => b[1] - a[1])
    .slice(0, 15);
}

function createSourcePieChart(sheet, sourceData, position) {
  const startRow = position.row;
  const headers = [['Источник', 'Лиды']];
  const data = headers.concat(sourceData);
  
  const range = sheet.getRange(startRow, 1, data.length, 2);
  range.setValues(data);
  
  createPieChart(sheet, range, 'Распределение лидов по источникам', 
                { row: startRow + data.length + 2, col: 1 });
}

function createChannelConversionChart(sheet, channelData, position) {
  const startRow = position.row;
  const headers = [['Канал', 'Лиды', 'Конверсия %']];
  const data = headers.concat(channelData);
  
  const range = sheet.getRange(startRow, 1, data.length, 3);
  range.setValues(data);
  
  createBarChart(sheet, range, 'Конверсия по каналам', 
                { row: startRow + data.length + 2, col: 1 }, 
                { width: 600, height: 400 });
}

function createTimeTrendsChart(sheet, timeData, position) {
  const startRow = position.row;
  const headers = [['Месяц', 'Лиды', 'Успешные', 'Конверсия %']];
  const data = headers.concat(timeData);
  
  const range = sheet.getRange(startRow, 1, data.length, 4);
  range.setValues(data);
  
  createLineChart(sheet, range, 'Временные тренды', 
                 { row: startRow + data.length + 2, col: 1 }, 
                 { smooth: true, width: 700, height: 350 });
}

function createManagerChart(sheet, managerData, position) {
  const startRow = position.row;
  const headers = [['Менеджер', 'Лиды', 'Конверсия %']];
  const data = headers.concat(managerData);
  
  const range = sheet.getRange(startRow, 1, data.length, 3);
  range.setValues(data);
  
  // Комбинированная диаграмма создается через обычный столбчатый график
  createBarChart(sheet, range, 'Производительность менеджеров', 
                { row: startRow + data.length + 2, col: 1 }, 
                { width: 800, height: 500 });
}
