/**
 * МОДУЛЬ: ПРИЧИНЫ ОТКАЗОВ + AI АНАЛИЗ
 * Анализ причин отказов с применением ИИ для категоризации
 */

function runRefusalsAiAnalysis() {
  console.log('Начинаем AI-анализ причин отказов...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const analysis = analyzeRefusalsWithAI(data);
    const sheet = createRefusalsAiReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.REFUSALS_AI}"`);
    console.log(`Проанализировано отказов: ${analysis.totalRefusals}`);
    
  } catch (error) {
    console.error('Ошибка при AI-анализе отказов:', error);
    throw error;
  }
}

function analyzeRefusalsWithAI(data) {
  const refusalReasons = {};
  const aiCategories = {};
  const sourceAnalysis = {};
  const timeAnalysis = {};
  const managerAnalysis = {};
  
  let totalRefusals = 0;
  let totalLeads = 0;
  
  data.forEach(row => {
    totalLeads++;
    
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const refusalReason = row[CONFIG.WORKING_AMO_COLUMNS.REFUSAL_REASON] || '';
    const source = parseUtmSource(row);
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const notes = row[CONFIG.WORKING_AMO_COLUMNS.NOTES] || '';
    const comment = row[CONFIG.WORKING_AMO_COLUMNS.COMMENT] || '';
    
    if (isRefusalStatus(status) || (refusalReason && refusalReason !== '')) {
      totalRefusals++;
      
      // Анализ причин отказов
      const mainReason = refusalReason || status;
      if (!refusalReasons[mainReason]) {
        refusalReasons[mainReason] = {
          reason: mainReason,
          count: 0,
          sources: {},
          managers: {},
          timePattern: {},
          detailedComments: []
        };
      }
      
      refusalReasons[mainReason].count++;
      
      // Источники отказов
      if (!refusalReasons[mainReason].sources[source]) {
        refusalReasons[mainReason].sources[source] = 0;
      }
      refusalReasons[mainReason].sources[source]++;
      
      // Менеджеры
      if (!refusalReasons[mainReason].managers[responsible]) {
        refusalReasons[mainReason].managers[responsible] = 0;
      }
      refusalReasons[mainReason].managers[responsible]++;
      
      // Комментарии для AI анализа
      const fullComment = `${refusalReason} ${notes} ${comment}`.trim();
      if (fullComment && fullComment !== '') {
        refusalReasons[mainReason].detailedComments.push(fullComment);
      }
      
      // AI категоризация
      const aiCategory = categorizeRefusalWithAI(mainReason, fullComment);
      if (!aiCategories[aiCategory]) {
        aiCategories[aiCategory] = {
          category: aiCategory,
          count: 0,
          examples: [],
          sources: new Set(),
          avgRefusalTime: 0
        };
      }
      
      aiCategories[aiCategory].count++;
      aiCategories[aiCategory].sources.add(source);
      if (aiCategories[aiCategory].examples.length < 5) {
        aiCategories[aiCategory].examples.push(mainReason);
      }
      
      // Анализ по источникам
      if (!sourceAnalysis[source]) {
        sourceAnalysis[source] = {
          source: source,
          totalLeads: 0,
          refusals: 0,
          topRefusalReasons: {}
        };
      }
      
      sourceAnalysis[source].refusals++;
      if (!sourceAnalysis[source].topRefusalReasons[mainReason]) {
        sourceAnalysis[source].topRefusalReasons[mainReason] = 0;
      }
      sourceAnalysis[source].topRefusalReasons[mainReason]++;
      
      // Анализ по времени
      if (createdAt) {
        const periods = getDatePeriods(createdAt);
        const hour = new Date(createdAt).getHours();
        
        if (!timeAnalysis[hour]) {
          timeAnalysis[hour] = { refusals: 0, total: 0 };
        }
        timeAnalysis[hour].refusals++;
        
        if (!refusalReasons[mainReason].timePattern[hour]) {
          refusalReasons[mainReason].timePattern[hour] = 0;
        }
        refusalReasons[mainReason].timePattern[hour]++;
      }
      
      // Анализ менеджеров
      if (responsible) {
        if (!managerAnalysis[responsible]) {
          managerAnalysis[responsible] = {
            manager: responsible,
            totalLeads: 0,
            refusals: 0,
            topRefusalTypes: {}
          };
        }
        
        managerAnalysis[responsible].refusals++;
        if (!managerAnalysis[responsible].topRefusalTypes[aiCategory]) {
          managerAnalysis[responsible].topRefusalTypes[aiCategory] = 0;
        }
        managerAnalysis[responsible].topRefusalTypes[aiCategory]++;
      }
    }
    
    // Считаем общее количество для источников и менеджеров
    if (!sourceAnalysis[source]) {
      sourceAnalysis[source] = {
        source: source,
        totalLeads: 0,
        refusals: 0,
        topRefusalReasons: {}
      };
    }
    sourceAnalysis[source].totalLeads++;
    
    if (responsible) {
      if (!managerAnalysis[responsible]) {
        managerAnalysis[responsible] = {
          manager: responsible,
          totalLeads: 0,
          refusals: 0,
          topRefusalTypes: {}
        };
      }
      managerAnalysis[responsible].totalLeads++;
    }
    
    // Подсчет времени
    if (createdAt) {
      const hour = new Date(createdAt).getHours();
      if (!timeAnalysis[hour]) {
        timeAnalysis[hour] = { refusals: 0, total: 0 };
      }
      timeAnalysis[hour].total++;
    }
  });
  
  // Обработка результатов
  const processedReasons = Object.values(refusalReasons)
    .sort((a, b) => b.count - a.count)
    .map(reason => ({
      ...reason,
      percentage: (reason.count / totalRefusals * 100),
      topSource: getTopEntry(reason.sources),
      topManager: getTopEntry(reason.managers),
      peakHour: getTopEntry(reason.timePattern)
    }));
  
  const processedCategories = Object.values(aiCategories)
    .sort((a, b) => b.count - a.count)
    .map(category => ({
      ...category,
      percentage: (category.count / totalRefusals * 100),
      sources: Array.from(category.sources),
      examples: category.examples.slice(0, 3)
    }));
  
  const processedSources = Object.values(sourceAnalysis)
    .sort((a, b) => b.refusals - a.refusals)
    .map(source => ({
      ...source,
      refusalRate: source.totalLeads > 0 ? (source.refusals / source.totalLeads * 100) : 0,
      topReason: getTopEntry(source.topRefusalReasons)
    }));
  
  const processedManagers = Object.values(managerAnalysis)
    .sort((a, b) => (b.refusals / b.totalLeads) - (a.refusals / a.totalLeads))
    .map(manager => ({
      ...manager,
      refusalRate: manager.totalLeads > 0 ? (manager.refusals / manager.totalLeads * 100) : 0,
      topType: getTopEntry(manager.topRefusalTypes)
    }));
  
  return {
    totalRefusals,
    totalLeads,
    refusalRate: (totalRefusals / totalLeads * 100),
    refusalReasons: processedReasons,
    aiCategories: processedCategories,
    sourceAnalysis: processedSources,
    managerAnalysis: processedManagers,
    timeAnalysis,
    insights: generateRefusalInsights(processedReasons, processedCategories, processedSources)
  };
}

function categorizeRefusalWithAI(reason, comment) {
  // Простая AI-подобная категоризация на базе ключевых слов
  const reasonLower = (reason + ' ' + comment).toLowerCase();
  
  if (reasonLower.includes('цен') || reasonLower.includes('дорог') || reasonLower.includes('бюджет')) {
    return 'Ценовые возражения';
  }
  
  if (reasonLower.includes('время') || reasonLower.includes('заня') || reasonLower.includes('позж')) {
    return 'Временные ограничения';
  }
  
  if (reasonLower.includes('конкурент') || reasonLower.includes('другой') || reasonLower.includes('уже выбр')) {
    return 'Выбор конкурентов';
  }
  
  if (reasonLower.includes('не отвеча') || reasonLower.includes('недоступ') || reasonLower.includes('не берет')) {
    return 'Недоступность клиента';
  }
  
  if (reasonLower.includes('не подход') || reasonLower.includes('не интерес') || reasonLower.includes('не нужн')) {
    return 'Отсутствие интереса';
  }
  
  if (reasonLower.includes('услов') || reasonLower.includes('требован') || reasonLower.includes('не устра')) {
    return 'Неподходящие условия';
  }
  
  if (reasonLower.includes('локац') || reasonLower.includes('район') || reasonLower.includes('далек')) {
    return 'Географические ограничения';
  }
  
  if (reasonLower.includes('качеств') || reasonLower.includes('опыт') || reasonLower.includes('репутац')) {
    return 'Сомнения в качестве';
  }
  
  return 'Прочие причины';
}

function generateRefusalInsights(reasons, categories, sources) {
  const insights = [];
  
  // Основная причина отказов
  if (reasons.length > 0) {
    insights.push(`Основная причина отказов: ${reasons[0].reason} (${reasons[0].percentage.toFixed(1)}%)`);
  }
  
  // Самая проблемная AI-категория
  if (categories.length > 0) {
    insights.push(`Главная категория проблем: ${categories[0].category} (${categories[0].percentage.toFixed(1)}%)`);
  }
  
  // Источник с наибольшим процентом отказов
  const worstSource = sources.find(s => s.refusalRate > 50);
  if (worstSource) {
    insights.push(`Проблемный источник: ${worstSource.source} (${worstSource.refusalRate.toFixed(1)}% отказов)`);
  }
  
  // Рекомендации
  if (categories[0]?.category === 'Ценовые возражения') {
    insights.push('🔍 Рекомендация: Пересмотреть ценовую политику или улучшить презентацию ценности');
  }
  
  if (categories[0]?.category === 'Недоступность клиента') {
    insights.push('📞 Рекомендация: Улучшить процесс обработки входящих звонков');
  }
  
  return insights;
}

function createRefusalsAiReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.REFUSALS_AI);
  
  let currentRow = 1;
  
  // Общие метрики и инсайты
  currentRow = addRefusalMetrics(sheet, analysis, currentRow);
  currentRow += 2;
  
  // AI категории
  currentRow = addAiCategories(sheet, analysis.aiCategories, currentRow);
  currentRow += 2;
  
  // Детальные причины
  currentRow = addDetailedReasons(sheet, analysis.refusalReasons, currentRow);
  currentRow += 2;
  
  // Анализ по источникам
  currentRow = addSourceRefusalAnalysis(sheet, analysis.sourceAnalysis, currentRow);
  currentRow += 2;
  
  // Анализ менеджеров
  addManagerRefusalAnalysis(sheet, analysis.managerAnalysis, currentRow);
  
  return sheet;
}

function addRefusalMetrics(sheet, analysis, startRow) {
  const headers = ['Метрика', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#d32f2f')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const metricsData = [
    ['Общий процент отказов', analysis.refusalRate.toFixed(1) + '%'],
    ['Всего отказов', analysis.totalRefusals],
    ['Всего лидов', analysis.totalLeads],
    ['Количество причин', analysis.refusalReasons.length],
    ['AI категорий', analysis.aiCategories.length]
  ];
  
  // Добавляем инсайты
  analysis.insights.forEach(insight => {
    metricsData.push(['💡 ' + insight.split(':')[0], insight.split(':')[1] || insight]);
  });
  
  const dataRange = sheet.getRange(startRow + 1, 1, metricsData.length, 2);
  dataRange.setValues(metricsData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + metricsData.length + 1;
}

function addAiCategories(sheet, categories, startRow) {
  const headers = [
    'AI Категория',
    'Количество',
    'Процент',
    'Примеры',
    'Источники'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1976d2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const categoryData = categories.map(category => [
    category.category,
    category.count,
    category.percentage.toFixed(1) + '%',
    category.examples.join('; '),
    category.sources.slice(0, 3).join(', ')
  ]);
  
  if (categoryData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, categoryData.length, headers.length);
    dataRange.setValues(categoryData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(categoryData.length, 1) + 1;
}

function addDetailedReasons(sheet, reasons, startRow) {
  const headers = [
    'Причина отказа',
    'Количество',
    'Процент',
    'Топ источник',
    'Топ менеджер',
    'Пиковый час'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#388e3c')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const reasonData = reasons.slice(0, 20).map(reason => [
    reason.reason,
    reason.count,
    reason.percentage.toFixed(1) + '%',
    reason.topSource,
    reason.topManager,
    reason.peakHour
  ]);
  
  if (reasonData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, reasonData.length, headers.length);
    dataRange.setValues(reasonData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(reasonData.length, 1) + 1;
}

function addSourceRefusalAnalysis(sheet, sources, startRow) {
  const headers = [
    'Источник',
    'Всего лидов',
    'Отказов',
    'Процент отказов',
    'Главная причина'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#f57c00')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const sourceData = sources.slice(0, 15).map(source => [
    source.source,
    source.totalLeads,
    source.refusals,
    source.refusalRate.toFixed(1) + '%',
    source.topReason
  ]);
  
  if (sourceData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, sourceData.length, headers.length);
    dataRange.setValues(sourceData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    
    // Условное форматирование для процента отказов
    const refusalRange = sheet.getRange(startRow + 1, 4, sourceData.length, 1);
    const refusalRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#d32f2f', SpreadsheetApp.InterpolationType.NUMBER, '80')
      .setGradientMidpointWithValue('#ff9800', SpreadsheetApp.InterpolationType.NUMBER, '40')
      .setGradientMinpointWithValue('#4caf50', SpreadsheetApp.InterpolationType.NUMBER, '10')
      .setRanges([refusalRange])
      .build();
    
    sheet.setConditionalFormatRules([refusalRule]);
  }
  
  return startRow + Math.max(sourceData.length, 1) + 1;
}

function addManagerRefusalAnalysis(sheet, managers, startRow) {
  const headers = [
    'Менеджер',
    'Всего лидов',
    'Отказов',
    'Процент отказов',
    'Основной тип проблем'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#7b1fa2')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const managerData = managers.map(manager => [
    manager.manager,
    manager.totalLeads,
    manager.refusals,
    manager.refusalRate.toFixed(1) + '%',
    manager.topType
  ]);
  
  if (managerData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, managerData.length, headers.length);
    dataRange.setValues(managerData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, headers.length);
}
