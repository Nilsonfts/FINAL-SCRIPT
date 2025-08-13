/**
 * МОДУЛЬ: ЛИДЫ ПО КАНАЛАМ
 * Анализ распределения лидов по источникам трафика
 */

function runLeadsByChannels() {
  console.log('Начинаем анализ лидов по каналам...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const analysis = analyzeLeadsByChannels(data);
    const sheet = createLeadsByChannelsReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.LEADS_CHANNELS}"`);
    console.log(`Проанализировано лидов: ${analysis.totalLeads}`);
    
  } catch (error) {
    console.error('Ошибка при анализе лидов по каналам:', error);
    throw error;
  }
}

function analyzeLeadsByChannels(data) {
  const channelStats = {};
  let totalLeads = 0;
  let successfulLeads = 0;
  
  data.forEach(row => {
    const source = parseUtmSource(row);
    const channelType = getChannelType(source);
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const budget = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const isSuccess = isSuccessStatus(status);
    
    totalLeads++;
    if (isSuccess) successfulLeads++;
    
    // Инициализация канала
    if (!channelStats[source]) {
      channelStats[source] = {
        source: source,
        channelType: channelType,
        totalCount: 0,
        successCount: 0,
        totalBudget: 0,
        totalFactAmount: 0,
        byMonth: {},
        byStatus: {}
      };
    }
    
    const channel = channelStats[source];
    channel.totalCount++;
    if (isSuccess) channel.successCount++;
    channel.totalBudget += budget;
    channel.totalFactAmount += factAmount;
    
    // Анализ по месяцам
    if (createdAt) {
      const periods = getDatePeriods(createdAt);
      if (periods.yearMonth) {
        if (!channel.byMonth[periods.yearMonth]) {
          channel.byMonth[periods.yearMonth] = {
            count: 0,
            successCount: 0,
            budget: 0,
            factAmount: 0
          };
        }
        channel.byMonth[periods.yearMonth].count++;
        if (isSuccess) channel.byMonth[periods.yearMonth].successCount++;
        channel.byMonth[periods.yearMonth].budget += budget;
        channel.byMonth[periods.yearMonth].factAmount += factAmount;
      }
    }
    
    // Анализ по статусам
    if (!channel.byStatus[status]) {
      channel.byStatus[status] = 0;
    }
    channel.byStatus[status]++;
  });
  
  // Конвертируем в массив и сортируем
  const channels = Object.values(channelStats)
    .sort((a, b) => b.totalCount - a.totalCount);
  
  return {
    totalLeads,
    successfulLeads,
    channels,
    conversionRate: totalLeads > 0 ? (successfulLeads / totalLeads * 100) : 0
  };
}

function createLeadsByChannelsReport(analysis) {
  const headers = [
    'Источник',
    'Тип канала',
    'Всего лидов',
    'Успешных',
    'Конверсия %',
    'Общий бюджет',
    'Факт сумма',
    'ROMI %',
    'Лидов в месяц (среднее)',
    'Топ статусы'
  ];
  
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.LEADS_CHANNELS, headers);
  
  // Подготавливаем данные
  const reportData = [];
  
  analysis.channels.forEach(channel => {
    const conversionRate = channel.totalCount > 0 ? 
      (channel.successCount / channel.totalCount * 100) : 0;
    
    const romi = channel.totalBudget > 0 ? 
      ((channel.totalFactAmount - channel.totalBudget) / channel.totalBudget * 100) : 0;
    
    // Считаем среднее количество лидов в месяц
    const monthsWithData = Object.keys(channel.byMonth).length;
    const avgLeadsPerMonth = monthsWithData > 0 ? 
      Math.round(channel.totalCount / monthsWithData) : 0;
    
    // Топ-3 статуса
    const topStatuses = Object.entries(channel.byStatus)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([status, count]) => `${status}: ${count}`)
      .join('; ');
    
    reportData.push([
      channel.source,
      channel.channelType,
      channel.totalCount,
      channel.successCount,
      conversionRate.toFixed(1) + '%',
      formatCurrency(channel.totalBudget),
      formatCurrency(channel.totalFactAmount),
      romi.toFixed(1) + '%',
      avgLeadsPerMonth,
      topStatuses
    ]);
  });
  
  // Записываем данные
  if (reportData.length > 0) {
    const dataRange = sheet.getRange(2, 1, reportData.length, headers.length);
    dataRange.setValues(reportData);
    
    // Форматирование
    formatLeadsByChannelsReport(sheet, reportData.length);
  }
  
  // Добавляем сводку
  addLeadsByChannelsSummary(sheet, analysis, reportData.length + 3);
  
  return sheet;
}

function formatLeadsByChannelsReport(sheet, dataRows) {
  const headerRange = sheet.getRange(1, 1, 1, 10);
  headerRange.setBackground('#1e3d59')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  if (dataRows > 0) {
    // Числовые столбцы
    const numbersRange = sheet.getRange(2, 3, dataRows, 2); // Всего лидов, Успешных
    numbersRange.setNumberFormat('#,##0');
    
    // Проценты
    const percentRange1 = sheet.getRange(2, 5, dataRows, 1); // Конверсия
    percentRange1.setNumberFormat('0.0%');
    
    const percentRange2 = sheet.getRange(2, 8, dataRows, 1); // ROMI
    percentRange2.setNumberFormat('0.0%');
    
    // Деньги
    const moneyRange1 = sheet.getRange(2, 6, dataRows, 1); // Бюджет
    const moneyRange2 = sheet.getRange(2, 7, dataRows, 1); // Факт
    moneyRange1.setNumberFormat('#,##0 ₽');
    moneyRange2.setNumberFormat('#,##0 ₽');
    
    // Условное форматирование для конверсии
    const conversionRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#34a853', SpreadsheetApp.InterpolationType.NUMBER, '30')
      .setGradientMidpointWithValue('#fbbc04', SpreadsheetApp.InterpolationType.NUMBER, '15')
      .setGradientMinpointWithValue('#ea4335', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([percentRange1])
      .build();
    
    sheet.setConditionalFormatRules([conversionRule]);
  }
  
  // Автоматическая ширина столбцов
  sheet.autoResizeColumns(1, 10);
}

function addLeadsByChannelsSummary(sheet, analysis, startRow) {
  const summaryHeaders = ['Показатель', 'Значение'];
  const summaryRange = sheet.getRange(startRow, 1, 1, 2);
  summaryRange.setValues([summaryHeaders]);
  summaryRange.setBackground('#f8f9fa')
             .setFontWeight('bold')
             .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const summaryData = [
    ['Общее количество лидов', analysis.totalLeads],
    ['Успешных лидов', analysis.successfulLeads],
    ['Общая конверсия', formatPercent(analysis.successfulLeads, analysis.totalLeads)],
    ['Количество каналов', analysis.channels.length],
    ['Лучший канал по объему', analysis.channels[0]?.source || 'Нет данных'],
    ['Лучший канал по конверсии', getBestChannelByConversion(analysis.channels)]
  ];
  
  const summaryDataRange = sheet.getRange(startRow + 1, 1, summaryData.length, 2);
  summaryDataRange.setValues(summaryData);
  summaryDataRange.setFontFamily(CONFIG.DEFAULT_FONT);
}

function getBestChannelByConversion(channels) {
  if (!channels || channels.length === 0) return 'Нет данных';
  
  const bestChannel = channels
    .filter(ch => ch.totalCount >= 5) // Минимум 5 лидов для корректности
    .sort((a, b) => {
      const convA = a.totalCount > 0 ? a.successCount / a.totalCount : 0;
      const convB = b.totalCount > 0 ? b.successCount / b.totalCount : 0;
      return convB - convA;
    })[0];
  
  if (!bestChannel) return channels[0]?.source || 'Нет данных';
  
  const conversion = bestChannel.totalCount > 0 ? 
    (bestChannel.successCount / bestChannel.totalCount * 100).toFixed(1) : 0;
  
  return `${bestChannel.source} (${conversion}%)`;
}
