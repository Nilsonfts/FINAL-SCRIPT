/**
 * МОДУЛЬ: ПЕРВЫЕ КАСАНИЯ
 * Анализ первичных источников и точек входа клиентов
 */

function runFirstTouchAnalysis() {
  console.log('Начинаем анализ первых касаний...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    const analysis = analyzeFirstTouchPoints(data);
    const sheet = createFirstTouchReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.FIRST_TOUCH}"`);
    console.log(`Проанализировано точек входа: ${analysis.touchPoints.length}`);
    
  } catch (error) {
    console.error('Ошибка при анализе первых касаний:', error);
    throw error;
  }
}

function analyzeFirstTouchPoints(data) {
  const touchPoints = {};
  const timeAnalysis = {};
  const deviceAnalysis = {};
  const geographyAnalysis = {};
  
  data.forEach(row => {
    const source = parseUtmSource(row);
    const referer = row[CONFIG.WORKING_AMO_COLUMNS.REFERER] || '';
    const visitTime = row[CONFIG.WORKING_AMO_COLUMNS.VISIT_TIME] || '';
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const cityTag = row[CONFIG.WORKING_AMO_COLUMNS.CITY_TAG] || '';
    const formName = row[CONFIG.WORKING_AMO_COLUMNS.FORMNAME] || '';
    const buttonText = row[CONFIG.WORKING_AMO_COLUMNS.BUTTON_TEXT] || '';
    
    const isSuccess = isSuccessStatus(status);
    const touchPointKey = determineTouchPoint(source, referer, formName, buttonText);
    
    // Анализ точек входа
    if (!touchPoints[touchPointKey]) {
      touchPoints[touchPointKey] = {
        touchPoint: touchPointKey,
        source: source,
        totalLeads: 0,
        successLeads: 0,
        uniqueSources: new Set(),
        avgTimeToLead: 0,
        topForms: {},
        topButtons: {},
        conversionFunnel: {
          visits: 0,
          forms: 0,
          leads: 0,
          success: 0
        }
      };
    }
    
    const tp = touchPoints[touchPointKey];
    tp.totalLeads++;
    if (isSuccess) tp.successLeads++;
    tp.uniqueSources.add(source);
    
    if (formName) {
      tp.topForms[formName] = (tp.topForms[formName] || 0) + 1;
      tp.conversionFunnel.forms++;
    }
    
    if (buttonText) {
      tp.topButtons[buttonText] = (tp.topButtons[buttonText] || 0) + 1;
    }
    
    tp.conversionFunnel.leads++;
    if (isSuccess) tp.conversionFunnel.success++;
    
    // Анализ по времени (час дня)
    if (visitTime || createdAt) {
      const timeKey = extractHourFromTime(visitTime || createdAt);
      if (timeKey !== null) {
        if (!timeAnalysis[timeKey]) {
          timeAnalysis[timeKey] = { leads: 0, success: 0 };
        }
        timeAnalysis[timeKey].leads++;
        if (isSuccess) timeAnalysis[timeKey].success++;
      }
    }
    
    // География
    if (cityTag) {
      if (!geographyAnalysis[cityTag]) {
        geographyAnalysis[cityTag] = { leads: 0, success: 0, sources: new Set() };
      }
      geographyAnalysis[cityTag].leads++;
      if (isSuccess) geographyAnalysis[cityTag].success++;
      geographyAnalysis[cityTag].sources.add(source);
    }
  });
  
  // Обрабатываем результаты
  const processedTouchPoints = Object.values(touchPoints).map(tp => ({
    ...tp,
    uniqueSources: Array.from(tp.uniqueSources),
    conversionRate: tp.totalLeads > 0 ? (tp.successLeads / tp.totalLeads * 100) : 0,
    topForm: getTopEntry(tp.topForms),
    topButton: getTopEntry(tp.topButtons),
    funnelConversion: tp.conversionFunnel.forms > 0 ? 
      (tp.conversionFunnel.leads / tp.conversionFunnel.forms * 100) : 0
  })).sort((a, b) => b.totalLeads - a.totalLeads);
  
  const processedGeography = Object.entries(geographyAnalysis).map(([city, data]) => ({
    city,
    leads: data.leads,
    success: data.success,
    conversionRate: data.leads > 0 ? (data.success / data.leads * 100) : 0,
    sourcesCount: data.sources.size,
    topSources: Array.from(data.sources).slice(0, 3).join(', ')
  })).sort((a, b) => b.leads - a.leads);
  
  return {
    touchPoints: processedTouchPoints,
    timeAnalysis,
    geographyAnalysis: processedGeography,
    summary: {
      totalTouchPoints: processedTouchPoints.length,
      avgConversionRate: calculateAvgConversion(processedTouchPoints),
      bestTouchPoint: processedTouchPoints[0]?.touchPoint || 'Нет данных',
      peakHour: findPeakHour(timeAnalysis)
    }
  };
}

function determineTouchPoint(source, referer, formName, buttonText) {
  // Определяем основную точку касания
  let touchPoint = 'Неопределено';
  
  const sourceStr = (source || '').toString().toLowerCase();
  
  if (sourceStr.includes('директ') || sourceStr.includes('google ads')) {
    touchPoint = 'Контекстная реклама';
  } else if (sourceStr.includes('органика')) {
    touchPoint = 'Органический поиск';
  } else if (sourceStr.includes('социальн') || sourceStr.includes('vk') || sourceStr.includes('instagram')) {
    touchPoint = 'Социальные сети';
  } else if (referer) {
    if (referer.includes('yandex.') || referer.includes('google.')) {
      touchPoint = 'Поисковые системы';
    } else {
      touchPoint = 'Внешние ссылки';
    }
  } else if (formName) {
    touchPoint = `Форма: ${formName}`;
  } else if (buttonText) {
    touchPoint = `Кнопка: ${buttonText}`;
  }
  
  return touchPoint;
}

function extractHourFromTime(timeValue) {
  if (!timeValue) return null;
  
  let date;
  if (timeValue instanceof Date) {
    date = timeValue;
  } else if (typeof timeValue === 'string') {
    date = new Date(timeValue);
  } else {
    return null;
  }
  
  if (isNaN(date.getTime())) return null;
  
  return date.getHours();
}

function getTopEntry(entries) {
  if (!entries || Object.keys(entries).length === 0) return 'Не указано';
  
  const sorted = Object.entries(entries).sort((a, b) => b[1] - a[1]);
  return `${sorted[0][0]} (${sorted[0][1]})`;
}

function calculateAvgConversion(touchPoints) {
  if (!touchPoints || touchPoints.length === 0) return 0;
  
  const totalConversion = touchPoints.reduce((sum, tp) => sum + tp.conversionRate, 0);
  return totalConversion / touchPoints.length;
}

function findPeakHour(timeAnalysis) {
  if (!timeAnalysis || Object.keys(timeAnalysis).length === 0) return 'Не определено';
  
  const sorted = Object.entries(timeAnalysis).sort((a, b) => b[1].leads - a[1].leads);
  return `${sorted[0][0]}:00 (${sorted[0][1].leads} лидов)`;
}

function createFirstTouchReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.FIRST_TOUCH);
  
  let currentRow = 1;
  
  // Сводка
  currentRow = addFirstTouchSummary(sheet, analysis.summary, currentRow);
  currentRow += 2;
  
  // Анализ точек касания
  currentRow = addTouchPointsAnalysis(sheet, analysis.touchPoints, currentRow);
  currentRow += 2;
  
  // Анализ по времени
  currentRow = addTimeAnalysis(sheet, analysis.timeAnalysis, currentRow);
  currentRow += 2;
  
  // География
  addGeographyAnalysis(sheet, analysis.geographyAnalysis, currentRow);
  
  return sheet;
}

function addFirstTouchSummary(sheet, summary, startRow) {
  const headers = ['Показатель', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1e3d59')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const summaryData = [
    ['Всего точек касания', summary.totalTouchPoints],
    ['Средняя конверсия', summary.avgConversionRate.toFixed(1) + '%'],
    ['Лучшая точка касания', summary.bestTouchPoint],
    ['Пиковый час активности', summary.peakHour]
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, summaryData.length, 2);
  dataRange.setValues(summaryData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + summaryData.length + 1;
}

function addTouchPointsAnalysis(sheet, touchPoints, startRow) {
  const headers = [
    'Точка касания',
    'Всего лидов',
    'Успешных',
    'Конверсия %',
    'Источники',
    'Топ форма',
    'Топ кнопка',
    'Конверсия воронки %'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#0d7377')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const touchPointData = touchPoints.map(tp => [
    tp.touchPoint,
    tp.totalLeads,
    tp.successLeads,
    tp.conversionRate.toFixed(1) + '%',
    tp.uniqueSources.length,
    tp.topForm,
    tp.topButton,
    tp.funnelConversion.toFixed(1) + '%'
  ]);
  
  if (touchPointData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, touchPointData.length, headers.length);
    dataRange.setValues(touchPointData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(touchPointData.length, 1) + 1;
}

function addTimeAnalysis(sheet, timeAnalysis, startRow) {
  const hours = Object.keys(timeAnalysis).map(Number).sort((a, b) => a - b);
  
  const headers = ['Час', 'Лиды', 'Успешные', 'Конверсия %'];
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#14213d')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const timeData = hours.map(hour => {
    const data = timeAnalysis[hour];
    const conversionRate = data.leads > 0 ? (data.success / data.leads * 100) : 0;
    return [
      `${hour}:00`,
      data.leads,
      data.success,
      conversionRate.toFixed(1) + '%'
    ];
  });
  
  if (timeData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, timeData.length, headers.length);
    dataRange.setValues(timeData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(timeData.length, 1) + 1;
}

function addGeographyAnalysis(sheet, geographyAnalysis, startRow) {
  const headers = [
    'Город',
    'Лиды',
    'Успешные',
    'Конверсия %',
    'Кол-во источников',
    'Топ источники'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#2c5f41')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const geoData = geographyAnalysis.slice(0, 20).map(geo => [
    geo.city,
    geo.leads,
    geo.success,
    geo.conversionRate.toFixed(1) + '%',
    geo.sourcesCount,
    geo.topSources
  ]);
  
  if (geoData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, geoData.length, headers.length);
    dataRange.setValues(geoData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, headers.length);
}
