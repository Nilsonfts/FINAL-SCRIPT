/**
 * МОДУЛЬ: ЭКСПОРТ И ИМПОРТ ДАННЫХ
 * Утилиты для работы с внешними источниками данных
 */

function exportDataToCSV() {
  console.log('Экспортируем данные в CSV...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для экспорта');
      return;
    }
    
    // Создаем CSV-содержимое
    const csvContent = createCSVContent(data);
    
    // Создаем blob и скачиваем файл
    const blob = Utilities.newBlob(csvContent, 'text/csv', `amo_export_${getFormattedDate()}.csv`);
    
    // Сохраняем в Google Drive
    const file = DriveApp.createFile(blob);
    console.log(`Файл сохранен в Google Drive: ${file.getName()}`);
    
    return file.getDownloadUrl();
    
  } catch (error) {
    console.error('Ошибка при экспорте в CSV:', error);
    throw error;
  }
}

function createCSVContent(data) {
  // Заголовки CSV
  const headers = [
    'ID', 'Имя', 'Статус', 'Ответственный', 'Дата создания',
    'Источник', 'UTM Source', 'UTM Medium', 'UTM Campaign',
    'Название формы', 'Телефон', 'Email', 'Сумма план',
    'Сумма факт', 'Первое касание', 'Последнее касание'
  ];
  
  let csvContent = headers.join(',') + '\n';
  
  data.forEach(row => {
    const csvRow = [
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.ID] || ''),
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.NAME] || ''),
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || ''),
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || ''),
      escapeCSV(formatDate(row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT])),
      escapeCSV(parseUtmSource(row)),
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || ''),
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.UTM_MEDIUM] || ''),
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.UTM_CAMPAIGN] || ''),
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.FORMNAME] || ''),
      escapeCSV(normalizePhone(row[CONFIG.WORKING_AMO_COLUMNS.PHONE])),
      escapeCSV(row[CONFIG.WORKING_AMO_COLUMNS.EMAIL] || ''),
      escapeCSV(formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.PLAN_AMOUNT])),
      escapeCSV(formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT])),
      escapeCSV(formatDate(row[CONFIG.WORKING_AMO_COLUMNS.FIRST_TOUCH])),
      escapeCSV(formatDate(row[CONFIG.WORKING_AMO_COLUMNS.LAST_TOUCH]))
    ];
    
    csvContent += csvRow.join(',') + '\n';
  });
  
  return csvContent;
}

function escapeCSV(value) {
  if (value == null || value === '') return '';
  
  const stringValue = String(value);
  
  // Если содержит запятую, кавычки или перенос строки - заключаем в кавычки
  if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
    return '"' + stringValue.replace(/"/g, '""') + '"';
  }
  
  return stringValue;
}

function importDataFromGoogleSheets(spreadsheetUrl, sheetName) {
  console.log('Импортируем данные из Google Sheets...');
  
  try {
    // Извлекаем ID таблицы из URL
    const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);
    if (!spreadsheetId) {
      throw new Error('Некорректный URL Google Sheets');
    }
    
    // Открываем внешнюю таблицу
    const externalSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sourceSheet = externalSpreadsheet.getSheetByName(sheetName);
    
    if (!sourceSheet) {
      throw new Error(`Лист "${sheetName}" не найден`);
    }
    
    // Получаем данные
    const data = sourceSheet.getDataRange().getValues();
    if (data.length === 0) {
      console.log('Нет данных для импорта');
      return [];
    }
    
    console.log(`Импортировано ${data.length} строк`);
    return data;
    
  } catch (error) {
    console.error('Ошибка при импорте данных:', error);
    throw error;
  }
}

function extractSpreadsheetId(url) {
  const regex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

function exportAnalyticsReport() {
  console.log('Создаем аналитический отчет для экспорта...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) return null;
    
    const reportData = {
      summary: generateSummaryReport(data),
      channels: generateChannelReport(data),
      conversion: generateConversionReport(data),
      managers: generateManagerReport(data),
      timeline: generateTimelineReport(data)
    };
    
    // Создаем JSON отчет
    const jsonReport = JSON.stringify(reportData, null, 2);
    const jsonBlob = Utilities.newBlob(jsonReport, 'application/json', 
                                      `analytics_report_${getFormattedDate()}.json`);
    
    const jsonFile = DriveApp.createFile(jsonBlob);
    console.log(`JSON отчет сохранен: ${jsonFile.getName()}`);
    
    // Создаем Excel-подобный отчет в новой таблице
    const reportSpreadsheet = createReportSpreadsheet(reportData);
    
    return {
      jsonUrl: jsonFile.getDownloadUrl(),
      spreadsheetUrl: reportSpreadsheet.getUrl()
    };
    
  } catch (error) {
    console.error('Ошибка при создании отчета:', error);
    throw error;
  }
}

function generateSummaryReport(data) {
  const totalLeads = data.length;
  const successLeads = data.filter(row => isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '')).length;
  const totalRevenue = data.reduce((sum, row) => {
    const amount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    return sum + (isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '') ? amount : 0);
  }, 0);
  
  const uniqueSources = new Set(data.map(row => parseUtmSource(row))).size;
  const uniqueManagers = new Set(data.map(row => row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE])).size;
  
  return {
    totalLeads,
    successLeads,
    conversionRate: totalLeads > 0 ? (successLeads / totalLeads * 100).toFixed(2) : 0,
    totalRevenue,
    averageRevenue: successLeads > 0 ? (totalRevenue / successLeads).toFixed(2) : 0,
    uniqueSources,
    uniqueManagers
  };
}

function generateChannelReport(data) {
  const channels = {};
  
  data.forEach(row => {
    const source = parseUtmSource(row);
    const channelType = getChannelType(source);
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    const isSuccess = isSuccessStatus(status);
    
    if (!channels[channelType]) {
      channels[channelType] = {
        leads: 0,
        successLeads: 0,
        revenue: 0,
        sources: new Set()
      };
    }
    
    channels[channelType].leads++;
    channels[channelType].sources.add(source);
    
    if (isSuccess) {
      channels[channelType].successLeads++;
      channels[channelType].revenue += factAmount;
    }
  });
  
  // Конвертируем Set в массив для JSON
  const result = {};
  Object.keys(channels).forEach(channel => {
    const data = channels[channel];
    result[channel] = {
      leads: data.leads,
      successLeads: data.successLeads,
      conversionRate: data.leads > 0 ? (data.successLeads / data.leads * 100).toFixed(2) : 0,
      revenue: data.revenue,
      averageRevenue: data.successLeads > 0 ? (data.revenue / data.successLeads).toFixed(2) : 0,
      sourcesCount: data.sources.size,
      sources: Array.from(data.sources)
    };
  });
  
  return result;
}

function generateConversionReport(data) {
  const funnelSteps = {
    visitors: data.length,
    leads: data.length,
    qualified: 0,
    meetings: 0,
    proposals: 0,
    deals: 0
  };
  
  data.forEach(row => {
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    
    if (isQualifiedStatus(status)) funnelSteps.qualified++;
    if (isMeetingStatus(status)) funnelSteps.meetings++;
    if (isProposalStatus(status)) funnelSteps.proposals++;
    if (isSuccessStatus(status)) funnelSteps.deals++;
  });
  
  return {
    funnel: funnelSteps,
    conversionRates: {
      leadToQualified: funnelSteps.leads > 0 ? (funnelSteps.qualified / funnelSteps.leads * 100).toFixed(2) : 0,
      qualifiedToMeeting: funnelSteps.qualified > 0 ? (funnelSteps.meetings / funnelSteps.qualified * 100).toFixed(2) : 0,
      meetingToProposal: funnelSteps.meetings > 0 ? (funnelSteps.proposals / funnelSteps.meetings * 100).toFixed(2) : 0,
      proposalToDeal: funnelSteps.proposals > 0 ? (funnelSteps.deals / funnelSteps.proposals * 100).toFixed(2) : 0,
      overallConversion: funnelSteps.leads > 0 ? (funnelSteps.deals / funnelSteps.leads * 100).toFixed(2) : 0
    }
  };
}

function generateManagerReport(data) {
  const managers = {};
  
  data.forEach(row => {
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    if (!responsible) return;
    
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    const isSuccess = isSuccessStatus(status);
    
    if (!managers[responsible]) {
      managers[responsible] = {
        leads: 0,
        successLeads: 0,
        revenue: 0
      };
    }
    
    managers[responsible].leads++;
    
    if (isSuccess) {
      managers[responsible].successLeads++;
      managers[responsible].revenue += factAmount;
    }
  });
  
  // Добавляем расчетные метрики
  Object.keys(managers).forEach(manager => {
    const data = managers[manager];
    data.conversionRate = data.leads > 0 ? (data.successLeads / data.leads * 100).toFixed(2) : 0;
    data.averageRevenue = data.successLeads > 0 ? (data.revenue / data.successLeads).toFixed(2) : 0;
    data.revenuePerLead = data.leads > 0 ? (data.revenue / data.leads).toFixed(2) : 0;
  });
  
  return managers;
}

function generateTimelineReport(data) {
  const monthly = {};
  const daily = {};
  const hourly = Array(24).fill(0);
  
  data.forEach(row => {
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    if (!createdAt) return;
    
    const date = new Date(createdAt);
    const periods = getDatePeriods(createdAt);
    const hour = date.getHours();
    const dayOfWeek = date.getDay(); // 0 = Воскресенье
    
    // Месячная статистика
    if (!monthly[periods.yearMonth]) {
      monthly[periods.yearMonth] = { leads: 0, revenue: 0 };
    }
    monthly[periods.yearMonth].leads++;
    
    // Дневная статистика (по дням недели)
    if (!daily[dayOfWeek]) {
      daily[dayOfWeek] = { leads: 0, revenue: 0 };
    }
    daily[dayOfWeek].leads++;
    
    // Почасовая статистика
    hourly[hour]++;
    
    // Добавляем выручку для успешных лидов
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    if (isSuccessStatus(status)) {
      const revenue = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
      monthly[periods.yearMonth].revenue += revenue;
      daily[dayOfWeek].revenue += revenue;
    }
  });
  
  // Преобразуем дневную статистику в читаемый формат
  const dayNames = ['Воскресенье', 'Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
  const dailyNamed = {};
  Object.keys(daily).forEach(dayIndex => {
    const dayName = dayNames[parseInt(dayIndex)];
    dailyNamed[dayName] = daily[dayIndex];
  });
  
  return {
    monthly,
    daily: dailyNamed,
    hourly,
    trends: calculateTrends(monthly)
  };
}

function calculateTrends(monthlyData) {
  const months = Object.keys(monthlyData).sort();
  if (months.length < 2) return { growth: 0, direction: 'stable' };
  
  const lastMonth = monthlyData[months[months.length - 1]];
  const prevMonth = monthlyData[months[months.length - 2]];
  
  const leadsGrowth = prevMonth.leads > 0 ? 
    ((lastMonth.leads - prevMonth.leads) / prevMonth.leads * 100) : 0;
    
  const revenueGrowth = prevMonth.revenue > 0 ? 
    ((lastMonth.revenue - prevMonth.revenue) / prevMonth.revenue * 100) : 0;
  
  return {
    leadsGrowth: leadsGrowth.toFixed(2),
    revenueGrowth: revenueGrowth.toFixed(2),
    direction: leadsGrowth > 5 ? 'up' : (leadsGrowth < -5 ? 'down' : 'stable')
  };
}

function createReportSpreadsheet(reportData) {
  const reportSpreadsheet = SpreadsheetApp.create(`Аналитический отчет AMO - ${getFormattedDate()}`);
  
  // Лист 1: Общая сводка
  const summarySheet = reportSpreadsheet.getActiveSheet();
  summarySheet.setName('Сводка');
  createSummarySheet(summarySheet, reportData.summary);
  
  // Лист 2: Каналы
  const channelsSheet = reportSpreadsheet.insertSheet('Каналы');
  createChannelsSheet(channelsSheet, reportData.channels);
  
  // Лист 3: Конверсия
  const conversionSheet = reportSpreadsheet.insertSheet('Конверсия');
  createConversionSheet(conversionSheet, reportData.conversion);
  
  // Лист 4: Менеджеры
  const managersSheet = reportSpreadsheet.insertSheet('Менеджеры');
  createManagersSheet(managersSheet, reportData.managers);
  
  // Лист 5: Временные тренды
  const timelineSheet = reportSpreadsheet.insertSheet('Временные тренды');
  createTimelineSheet(timelineSheet, reportData.timeline);
  
  return reportSpreadsheet;
}

function createSummarySheet(sheet, summaryData) {
  const data = [
    ['ОБЩАЯ СВОДКА', ''],
    ['Всего лидов', summaryData.totalLeads],
    ['Успешных лидов', summaryData.successLeads],
    ['Конверсия', summaryData.conversionRate + '%'],
    ['Общая выручка', summaryData.totalRevenue],
    ['Средняя выручка с лида', summaryData.averageRevenue],
    ['Уникальных источников', summaryData.uniqueSources],
    ['Менеджеров', summaryData.uniqueManagers]
  ];
  
  sheet.getRange(1, 1, data.length, 2).setValues(data);
  
  // Форматирование
  sheet.getRange(1, 1, 1, 2).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
}

function createChannelsSheet(sheet, channelsData) {
  const headers = ['Канал', 'Лиды', 'Успешные', 'Конверсия %', 'Выручка', 'Средняя выручка', 'Источников'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const rows = Object.entries(channelsData).map(([channel, data]) => [
    channel,
    data.leads,
    data.successLeads,
    data.conversionRate,
    data.revenue,
    data.averageRevenue,
    data.sourcesCount
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Форматирование
  sheet.getRange(1, 1, 1, headers.length).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

function createConversionSheet(sheet, conversionData) {
  let currentRow = 1;
  
  // Воронка
  const funnelHeaders = ['Этап воронки', 'Количество'];
  sheet.getRange(currentRow, 1, 1, 2).setValues([funnelHeaders]);
  currentRow++;
  
  const funnelData = Object.entries(conversionData.funnel).map(([stage, count]) => [stage, count]);
  sheet.getRange(currentRow, 1, funnelData.length, 2).setValues(funnelData);
  currentRow += funnelData.length + 2;
  
  // Конверсии
  const conversionHeaders = ['Переход', 'Конверсия %'];
  sheet.getRange(currentRow, 1, 1, 2).setValues([conversionHeaders]);
  currentRow++;
  
  const conversionRates = Object.entries(conversionData.conversionRates).map(([transition, rate]) => [transition, rate]);
  sheet.getRange(currentRow, 1, conversionRates.length, 2).setValues(conversionRates);
  
  // Форматирование
  sheet.getRange(1, 1, 1, 2).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
  sheet.getRange(currentRow - conversionRates.length, 1, 1, 2).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
  sheet.autoResizeColumns(1, 2);
}

function createManagersSheet(sheet, managersData) {
  const headers = ['Менеджер', 'Лиды', 'Успешные', 'Конверсия %', 'Выручка', 'Средняя выручка', 'Выручка/лид'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const rows = Object.entries(managersData)
    .sort((a, b) => b[1].leads - a[1].leads)
    .map(([manager, data]) => [
      manager,
      data.leads,
      data.successLeads,
      data.conversionRate,
      data.revenue,
      data.averageRevenue,
      data.revenuePerLead
    ]);
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Форматирование
  sheet.getRange(1, 1, 1, headers.length).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
  sheet.autoResizeColumns(1, headers.length);
}

function createTimelineSheet(sheet, timelineData) {
  let currentRow = 1;
  
  // Месячные тренды
  sheet.getRange(currentRow, 1, 1, 3).setValues([['МЕСЯЧНЫЕ ТРЕНДЫ', '', '']]);
  sheet.getRange(currentRow, 1, 1, 3).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
  currentRow++;
  
  const monthlyHeaders = ['Месяц', 'Лиды', 'Выручка'];
  sheet.getRange(currentRow, 1, 1, 3).setValues([monthlyHeaders]);
  currentRow++;
  
  const monthlyRows = Object.entries(timelineData.monthly)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([month, data]) => [month, data.leads, data.revenue]);
  
  if (monthlyRows.length > 0) {
    sheet.getRange(currentRow, 1, monthlyRows.length, 3).setValues(monthlyRows);
    currentRow += monthlyRows.length + 2;
  }
  
  // Дневная статистика
  sheet.getRange(currentRow, 1, 1, 3).setValues([['СТАТИСТИКА ПО ДНЯМ НЕДЕЛИ', '', '']]);
  sheet.getRange(currentRow, 1, 1, 3).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
  currentRow++;
  
  const dailyHeaders = ['День недели', 'Лиды', 'Выручка'];
  sheet.getRange(currentRow, 1, 1, 3).setValues([dailyHeaders]);
  currentRow++;
  
  const dailyRows = Object.entries(timelineData.daily).map(([day, data]) => [day, data.leads, data.revenue]);
  
  if (dailyRows.length > 0) {
    sheet.getRange(currentRow, 1, dailyRows.length, 3).setValues(dailyRows);
  }
  
  sheet.autoResizeColumns(1, 3);
}

/**
 * ДОПОЛНИТЕЛЬНЫЕ УТИЛИТЫ СТАТУСОВ
 */
function isQualifiedStatus(status) {
  const qualifiedStatuses = ['Квалифицированный лид', 'Qualified', 'В работе', 'Обработан'];
  return qualifiedStatuses.some(s => status.toLowerCase().includes(s.toLowerCase()));
}

function isMeetingStatus(status) {
  const meetingStatuses = ['Встреча', 'Meeting', 'Презентация', 'Демо'];
  return meetingStatuses.some(s => status.toLowerCase().includes(s.toLowerCase()));
}

function isProposalStatus(status) {
  const proposalStatuses = ['Предложение', 'Proposal', 'КП', 'Коммерческое предложение'];
  return proposalStatuses.some(s => status.toLowerCase().includes(s.toLowerCase()));
}

function getFormattedDate() {
  return Utilities.formatDate(new Date(), CONFIG.FORMATTING.TIMEZONE, 'yyyy-MM-dd_HH-mm');
}
