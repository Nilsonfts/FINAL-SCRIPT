/**
 * МОДУЛЬ: СВОДНАЯ АНАЛИТИКА AMOCRM
 * Общий анализ эффективности работы с CRM
 */

function runAmoSummaryAnalysis() {
  console.log('Начинаем сводную аналитику AMOCRM...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const analysis = analyzeAmoSummary(data);
    const sheet = createAmoSummaryReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.AMO_SUMMARY}"`);
    console.log(`Проанализировано сделок: ${analysis.totalDeals}`);
    
  } catch (error) {
    console.error('Ошибка при сводной аналитике АМО:', error);
    throw error;
  }
}

function analyzeAmoSummary(data) {
  const statusAnalysis = {};
  const managerAnalysis = {};
  const pipelineAnalysis = {};
  const timeAnalysis = {};
  const budgetAnalysis = {};
  const tagsAnalysis = {};
  
  let totalDeals = 0;
  let totalBudget = 0;
  let totalFactAmount = 0;
  let successfulDeals = 0;
  let refusedDeals = 0;
  
  const currentDate = new Date();
  const currentMonth = currentDate.getMonth();
  const currentYear = currentDate.getFullYear();
  
  data.forEach(row => {
    totalDeals++;
    
    const dealId = row[CONFIG.WORKING_AMO_COLUMNS.ID] || '';
    const dealName = row[CONFIG.WORKING_AMO_COLUMNS.NAME] || '';
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    const budget = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.BUDGET]);
    const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const closedAt = row[CONFIG.WORKING_AMO_COLUMNS.CLOSED_AT];
    const tags = row[CONFIG.WORKING_AMO_COLUMNS.TAGS] || '';
    const leadType = row[CONFIG.WORKING_AMO_COLUMNS.LEAD_TYPE] || '';
    const source = parseUtmSource(row);
    
    const isSuccess = isSuccessStatus(status);
    const isRefusal = isRefusalStatus(status);
    
    totalBudget += budget;
    totalFactAmount += factAmount;
    
    if (isSuccess) successfulDeals++;
    if (isRefusal) refusedDeals++;
    
    // Анализ статусов
    if (!statusAnalysis[status]) {
      statusAnalysis[status] = {
        status: status,
        count: 0,
        totalBudget: 0,
        totalFact: 0,
        avgDealValue: 0,
        managers: new Set(),
        avgCloseTime: 0,
        sources: {}
      };
    }
    
    const statusData = statusAnalysis[status];
    statusData.count++;
    statusData.totalBudget += budget;
    statusData.totalFact += factAmount;
    statusData.managers.add(responsible);
    
    if (!statusData.sources[source]) statusData.sources[source] = 0;
    statusData.sources[source]++;
    
    // Время закрытия сделки
    if (createdAt && closedAt) {
      const closeTime = Math.round((new Date(closedAt) - new Date(createdAt)) / (1000 * 60 * 60 * 24));
      if (closeTime > 0) {
        statusData.avgCloseTime = (statusData.avgCloseTime * (statusData.count - 1) + closeTime) / statusData.count;
      }
    }
    
    // Анализ менеджеров
    if (responsible && responsible !== '') {
      if (!managerAnalysis[responsible]) {
        managerAnalysis[responsible] = {
          manager: responsible,
          totalDeals: 0,
          successDeals: 0,
          refusedDeals: 0,
          totalBudget: 0,
          totalFact: 0,
          avgDealSize: 0,
          conversionRate: 0,
          avgCloseTime: 0,
          topSources: {},
          topStatuses: {},
          currentMonthDeals: 0,
          efficiency: 0
        };
      }
      
      const manager = managerAnalysis[responsible];
      manager.totalDeals++;
      manager.totalBudget += budget;
      manager.totalFact += factAmount;
      
      if (isSuccess) manager.successDeals++;
      if (isRefusal) manager.refusedDeals++;
      
      if (!manager.topSources[source]) manager.topSources[source] = 0;
      manager.topSources[source]++;
      
      if (!manager.topStatuses[status]) manager.topStatuses[status] = 0;
      manager.topStatuses[status]++;
      
      // Текущий месяц
      if (createdAt) {
        const dealDate = new Date(createdAt);
        if (dealDate.getMonth() === currentMonth && dealDate.getFullYear() === currentYear) {
          manager.currentMonthDeals++;
        }
      }
    }
    
    // Временной анализ
    if (createdAt) {
      const periods = getDatePeriods(createdAt);
      const monthKey = periods.yearMonth;
      
      if (!timeAnalysis[monthKey]) {
        timeAnalysis[monthKey] = {
          period: monthKey,
          deals: 0,
          success: 0,
          refused: 0,
          budget: 0,
          fact: 0,
          avgDealValue: 0,
          conversionRate: 0
        };
      }
      
      const timeData = timeAnalysis[monthKey];
      timeData.deals++;
      timeData.budget += budget;
      timeData.fact += factAmount;
      
      if (isSuccess) timeData.success++;
      if (isRefusal) timeData.refused++;
    }
    
    // Анализ бюджетов
    const budgetRange = getBudgetRange(budget);
    if (!budgetAnalysis[budgetRange]) {
      budgetAnalysis[budgetRange] = {
        range: budgetRange,
        deals: 0,
        success: 0,
        conversionRate: 0,
        avgFact: 0
      };
    }
    
    budgetAnalysis[budgetRange].deals++;
    if (isSuccess) {
      budgetAnalysis[budgetRange].success++;
      budgetAnalysis[budgetRange].avgFact += factAmount;
    }
    
    // Анализ тегов
    if (tags && tags !== '') {
      const tagList = tags.split(/[,;]/).map(tag => tag.trim()).filter(tag => tag !== '');
      tagList.forEach(tag => {
        if (!tagsAnalysis[tag]) {
          tagsAnalysis[tag] = {
            tag: tag,
            deals: 0,
            success: 0,
            conversionRate: 0
          };
        }
        
        tagsAnalysis[tag].deals++;
        if (isSuccess) tagsAnalysis[tag].success++;
      });
    }
  });
  
  // Постобработка данных
  Object.values(statusAnalysis).forEach(status => {
    status.avgDealValue = status.count > 0 ? status.totalFact / status.count : 0;
    status.managers = Array.from(status.managers);
    status.topSource = getTopEntry(status.sources);
  });
  
  Object.values(managerAnalysis).forEach(manager => {
    manager.conversionRate = manager.totalDeals > 0 ? (manager.successDeals / manager.totalDeals * 100) : 0;
    manager.avgDealSize = manager.successDeals > 0 ? manager.totalFact / manager.successDeals : 0;
    manager.efficiency = calculateManagerEfficiency(manager);
    manager.topSource = getTopEntry(manager.topSources);
    manager.topStatus = getTopEntry(manager.topStatuses);
  });
  
  Object.values(timeAnalysis).forEach(time => {
    time.conversionRate = time.deals > 0 ? (time.success / time.deals * 100) : 0;
    time.avgDealValue = time.success > 0 ? time.fact / time.success : 0;
  });
  
  Object.values(budgetAnalysis).forEach(budget => {
    budget.conversionRate = budget.deals > 0 ? (budget.success / budget.deals * 100) : 0;
    budget.avgFact = budget.success > 0 ? budget.avgFact / budget.success : 0;
  });
  
  Object.values(tagsAnalysis).forEach(tag => {
    tag.conversionRate = tag.deals > 0 ? (tag.success / tag.deals * 100) : 0;
  });
  
  return {
    totalDeals,
    successfulDeals,
    refusedDeals,
    totalBudget,
    totalFactAmount,
    overallConversion: totalDeals > 0 ? (successfulDeals / totalDeals * 100) : 0,
    avgDealSize: successfulDeals > 0 ? totalFactAmount / successfulDeals : 0,
    statusAnalysis: Object.values(statusAnalysis).sort((a, b) => b.count - a.count),
    managerAnalysis: Object.values(managerAnalysis).sort((a, b) => b.efficiency - a.efficiency),
    timeAnalysis: Object.values(timeAnalysis).sort((a, b) => a.period.localeCompare(b.period)),
    budgetAnalysis: Object.values(budgetAnalysis),
    tagsAnalysis: Object.values(tagsAnalysis).sort((a, b) => b.deals - a.deals).slice(0, 20),
    kpis: calculateCrmKpis(totalDeals, successfulDeals, totalBudget, totalFactAmount, managerAnalysis)
  };
}

function getBudgetRange(budget) {
  if (budget === 0) return '0 ₽';
  if (budget <= 50000) return '1-50k ₽';
  if (budget <= 100000) return '50-100k ₽';
  if (budget <= 300000) return '100-300k ₽';
  if (budget <= 500000) return '300-500k ₽';
  return '500k+ ₽';
}

function calculateManagerEfficiency(manager) {
  // Комплексный показатель эффективности менеджера
  const conversionWeight = 0.4;
  const volumeWeight = 0.3;
  const currentMonthWeight = 0.3;
  
  const maxConversion = 100;
  const maxVolume = Math.max(manager.totalDeals, 50);
  const maxCurrentMonth = Math.max(manager.currentMonthDeals, 10);
  
  const conversionScore = (manager.conversionRate / maxConversion) * 100;
  const volumeScore = (manager.totalDeals / maxVolume) * 100;
  const currentMonthScore = (manager.currentMonthDeals / maxCurrentMonth) * 100;
  
  return (conversionScore * conversionWeight + 
          volumeScore * volumeWeight + 
          currentMonthScore * currentMonthWeight);
}

function calculateCrmKpis(totalDeals, successfulDeals, totalBudget, totalFactAmount, managerAnalysis) {
  const avgDealsPerManager = Object.keys(managerAnalysis).length > 0 ? 
    totalDeals / Object.keys(managerAnalysis).length : 0;
  
  const topManager = Object.values(managerAnalysis).sort((a, b) => b.efficiency - a.efficiency)[0];
  
  const romi = totalBudget > 0 ? ((totalFactAmount - totalBudget) / totalBudget * 100) : 0;
  
  return {
    avgDealsPerManager: avgDealsPerManager.toFixed(1),
    topManagerName: topManager ? topManager.manager : 'Не определен',
    topManagerEfficiency: topManager ? topManager.efficiency.toFixed(1) : '0',
    crmRomi: romi.toFixed(1) + '%',
    activeManagers: Object.keys(managerAnalysis).length
  };
}

function createAmoSummaryReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.AMO_SUMMARY);
  
  let currentRow = 1;
  
  // Общие KPI
  currentRow = addCrmKpis(sheet, analysis, currentRow);
  currentRow += 2;
  
  // Анализ статусов
  currentRow = addStatusAnalysis(sheet, analysis.statusAnalysis, currentRow);
  currentRow += 2;
  
  // Анализ менеджеров
  currentRow = addManagersSummary(sheet, analysis.managerAnalysis, currentRow);
  currentRow += 2;
  
  // Временная динамика
  currentRow = addTimeAnalysis(sheet, analysis.timeAnalysis, currentRow);
  currentRow += 2;
  
  // Анализ бюджетов
  currentRow = addBudgetAnalysis(sheet, analysis.budgetAnalysis, currentRow);
  currentRow += 2;
  
  // Топ теги
  addTagsAnalysis(sheet, analysis.tagsAnalysis, currentRow);
  
  return sheet;
}

function addCrmKpis(sheet, analysis, startRow) {
  const headers = ['KPI AMOCRM', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#ff6b35')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const kpiData = [
    ['Всего сделок', analysis.totalDeals],
    ['Успешных сделок', analysis.successfulDeals],
    ['Отказанных сделок', analysis.refusedDeals],
    ['Общая конверсия', analysis.overallConversion.toFixed(1) + '%'],
    ['Средний чек', formatCurrency(analysis.avgDealSize)],
    ['Общий бюджет', formatCurrency(analysis.totalBudget)],
    ['Общая выручка', formatCurrency(analysis.totalFactAmount)],
    ['ROI CRM', analysis.kpis.crmRomi],
    ['Активных менеджеров', analysis.kpis.activeManagers],
    ['Среднее сделок на менеджера', analysis.kpis.avgDealsPerManager],
    ['Топ менеджер', `${analysis.kpis.topManagerName} (${analysis.kpis.topManagerEfficiency})`]
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, kpiData.length, 2);
  dataRange.setValues(kpiData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + kpiData.length + 1;
}

function addStatusAnalysis(sheet, statusAnalysis, startRow) {
  const headers = [
    'Статус',
    'Количество',
    'Процент',
    'Общий бюджет',
    'Общий факт',
    'Средний чек',
    'Менеджеров',
    'Среднее время закрытия (дни)',
    'Топ источник'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1976d2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const totalDeals = statusAnalysis.reduce((sum, status) => sum + status.count, 0);
  
  const statusData = statusAnalysis.map(status => [
    status.status,
    status.count,
    totalDeals > 0 ? (status.count / totalDeals * 100).toFixed(1) + '%' : '0%',
    formatCurrency(status.totalBudget),
    formatCurrency(status.totalFact),
    formatCurrency(status.avgDealValue),
    status.managers.length,
    status.avgCloseTime.toFixed(1),
    status.topSource
  ]);
  
  if (statusData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, statusData.length, headers.length);
    dataRange.setValues(statusData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(statusData.length, 1) + 1;
}

function addManagersSummary(sheet, managers, startRow) {
  const headers = [
    'Менеджер',
    'Всего сделок',
    'Успешных',
    'Конверсия %',
    'Средний чек',
    'Выручка',
    'Текущий месяц',
    'Эффективность',
    'Топ источник'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#388e3c')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const managerData = managers.map(manager => [
    manager.manager,
    manager.totalDeals,
    manager.successDeals,
    manager.conversionRate.toFixed(1) + '%',
    formatCurrency(manager.avgDealSize),
    formatCurrency(manager.totalFact),
    manager.currentMonthDeals,
    manager.efficiency.toFixed(1),
    manager.topSource
  ]);
  
  if (managerData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, managerData.length, headers.length);
    dataRange.setValues(managerData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    
    // Условное форматирование для эффективности
    const efficiencyRange = sheet.getRange(startRow + 1, 8, managerData.length, 1);
    const efficiencyRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#4caf50', SpreadsheetApp.InterpolationType.NUMBER, '80')
      .setGradientMidpointWithValue('#ff9800', SpreadsheetApp.InterpolationType.NUMBER, '40')
      .setGradientMinpointWithValue('#f44336', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([efficiencyRange])
      .build();
    
    sheet.setConditionalFormatRules([efficiencyRule]);
  }
  
  return startRow + Math.max(managerData.length, 1) + 1;
}

function addTimeAnalysis(sheet, timeAnalysis, startRow) {
  const headers = [
    'Период',
    'Сделок',
    'Успешных',
    'Отказанных',
    'Конверсия %',
    'Средний чек'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#7b1fa2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const timeData = timeAnalysis.map(time => [
    time.period,
    time.deals,
    time.success,
    time.refused,
    time.conversionRate.toFixed(1) + '%',
    formatCurrency(time.avgDealValue)
  ]);
  
  if (timeData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, timeData.length, headers.length);
    dataRange.setValues(timeData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(timeData.length, 1) + 1;
}

function addBudgetAnalysis(sheet, budgetAnalysis, startRow) {
  const headers = [
    'Диапазон бюджета',
    'Сделок',
    'Успешных',
    'Конверсия %',
    'Средний факт'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#f57c00')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const budgetData = budgetAnalysis.map(budget => [
    budget.range,
    budget.deals,
    budget.success,
    budget.conversionRate.toFixed(1) + '%',
    formatCurrency(budget.avgFact)
  ]);
  
  if (budgetData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, budgetData.length, headers.length);
    dataRange.setValues(budgetData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(budgetData.length, 1) + 1;
}

function addTagsAnalysis(sheet, tagsAnalysis, startRow) {
  const headers = [
    'Тег',
    'Сделок',
    'Успешных',
    'Конверсия %'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d32f2f')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const tagData = tagsAnalysis.map(tag => [
    tag.tag,
    tag.deals,
    tag.success,
    tag.conversionRate.toFixed(1) + '%'
  ]);
  
  if (tagData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, tagData.length, headers.length);
    dataRange.setValues(tagData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, headers.length);
  
  return startRow + tagData.length + 1;
}
