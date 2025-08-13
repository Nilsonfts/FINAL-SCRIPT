/**
 * МОДУЛЬ: КОЛЛ-ТРЕКИНГ И ЗВОНКИ
 * Анализ телефонных звонков и качества обработки
 */

function runCallTrackingAnalysis() {
  console.log('Начинаем анализ колл-трекинга...');
  
  try {
    const amoData = getWorkingAmoData();
    const callData = getCallTrackingData();
    
    if (!amoData || amoData.length === 0) {
      console.log('Нет данных АМО для анализа');
      return;
    }
    
    const analysis = analyzeCallTracking(amoData, callData);
    const sheet = createCallTrackingReport(analysis);
    
    console.log(`Создан отчет на листе "${CONFIG.SHEETS.CALL_TRACKING}"`);
    console.log(`Проанализировано звонков: ${analysis.totalCalls}`);
    
  } catch (error) {
    console.error('Ошибка при анализе колл-трекинга:', error);
    throw error;
  }
}

function getCallTrackingData() {
  // Сначала ищем стандартный лист, потом пользовательский
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.CALL_TRACKING);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('КоллТрекинг');
  }
  
  if (!sheet) {
    console.log('Лист КоллТрекинг не найден, анализируем только АМО данные');
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

function analyzeCallTracking(amoData, callData) {
  const phoneAnalysis = {};
  const mangoAnalysis = {};
  const timeAnalysis = {};
  const qualityAnalysis = {};
  
  let totalCalls = 0;
  let answeredCalls = 0;
  let successfulCalls = 0;
  
  // Анализ АМО данных
  amoData.forEach(row => {
    const phone = row[CONFIG.WORKING_AMO_COLUMNS.PHONE] || '';
    const contactMango = row[CONFIG.WORKING_AMO_COLUMNS.CONTACT_MANGO] || '';
    const dealMango = row[CONFIG.WORKING_AMO_COLUMNS.DEAL_MANGO] || '';
    const responsible = row[CONFIG.WORKING_AMO_COLUMNS.RESPONSIBLE] || '';
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
    const source = parseUtmSource(row);
    
    const isSuccess = isSuccessStatus(status);
    const normalizedPhone = normalizePhone(phone);
    
    if (phone && phone !== '') {
      totalCalls++;
      if (isSuccess) successfulCalls++;
      
      // Анализ телефонов
      if (!phoneAnalysis[normalizedPhone]) {
        phoneAnalysis[normalizedPhone] = {
          phone: normalizedPhone,
          calls: 0,
          success: 0,
          sources: new Set(),
          managers: new Set(),
          lastActivity: null
        };
      }
      
      const phoneData = phoneAnalysis[normalizedPhone];
      phoneData.calls++;
      if (isSuccess) phoneData.success++;
      phoneData.sources.add(source);
      phoneData.managers.add(responsible);
      
      if (createdAt) {
        if (!phoneData.lastActivity || new Date(createdAt) > new Date(phoneData.lastActivity)) {
          phoneData.lastActivity = createdAt;
        }
      }
    }
    
    // Анализ Mango линий
    const mangoNumber = contactMango || dealMango;
    if (mangoNumber && mangoNumber !== '') {
      if (!mangoAnalysis[mangoNumber]) {
        mangoAnalysis[mangoNumber] = {
          number: mangoNumber,
          calls: 0,
          success: 0,
          managers: new Set(),
          sources: new Set()
        };
      }
      
      mangoAnalysis[mangoNumber].calls++;
      if (isSuccess) mangoAnalysis[mangoNumber].success++;
      mangoAnalysis[mangoNumber].managers.add(responsible);
      mangoAnalysis[mangoNumber].sources.add(source);
    }
    
    // Анализ по времени
    if (createdAt) {
      const periods = getDatePeriods(createdAt);
      const hour = new Date(createdAt).getHours();
      
      if (!timeAnalysis[hour]) {
        timeAnalysis[hour] = {
          hour: hour,
          calls: 0,
          success: 0,
          avgResponseTime: 0
        };
      }
      
      timeAnalysis[hour].calls++;
      if (isSuccess) timeAnalysis[hour].success++;
    }
    
    // Анализ качества (по менеджерам)
    if (responsible && responsible !== '') {
      if (!qualityAnalysis[responsible]) {
        qualityAnalysis[responsible] = {
          manager: responsible,
          totalCalls: 0,
          successCalls: 0,
          uniqueClients: new Set(),
          avgCallDuration: 0,
          topSources: {}
        };
      }
      
      const manager = qualityAnalysis[responsible];
      manager.totalCalls++;
      if (isSuccess) manager.successCalls++;
      if (normalizedPhone) manager.uniqueClients.add(normalizedPhone);
      
      if (!manager.topSources[source]) manager.topSources[source] = 0;
      manager.topSources[source]++;
    }
  });
  
  // Обработка call-tracking данных (если есть)
  const trackingNumbers = {};
  if (callData && callData.length > 0) {
    callData.forEach(row => {
      const phone = row[0] || '';           // A - Телефон
      const code = row[1] || '';            // B - Код источника
      const description = row[2] || '';     // C - Описание
      
      if (phone && code) {
        trackingNumbers[phone] = {
          phone: phone,
          code: code,
          description: description,
          source: mapTrackingSource(code),
          calls: 0,
          amoMatches: 0
        };
        
        // Ищем соответствия в АМО данных
        amoData.forEach(amoRow => {
          const amoPhone = normalizePhone(amoRow[CONFIG.WORKING_AMO_COLUMNS.PHONE] || '');
          const normalizedTracking = normalizePhone(phone);
          
          if (amoPhone && normalizedTracking && amoPhone === normalizedTracking) {
            trackingNumbers[phone].amoMatches++;
          }
        });
      }
    });
  }
  
  // Финализируем анализ
  const processedPhones = Object.values(phoneAnalysis).map(phone => ({
    ...phone,
    sources: Array.from(phone.sources),
    managers: Array.from(phone.managers),
    conversionRate: phone.calls > 0 ? (phone.success / phone.calls * 100) : 0,
    lastActivityFormatted: phone.lastActivity ? formatDate(phone.lastActivity) : 'Не определено'
  })).sort((a, b) => b.calls - a.calls);
  
  const processedMango = Object.values(mangoAnalysis).map(mango => ({
    ...mango,
    sources: Array.from(mango.sources),
    managers: Array.from(mango.managers),
    conversionRate: mango.calls > 0 ? (mango.success / mango.calls * 100) : 0
  })).sort((a, b) => b.calls - a.calls);
  
  const processedQuality = Object.values(qualityAnalysis).map(manager => ({
    ...manager,
    uniqueClients: manager.uniqueClients.size,
    conversionRate: manager.totalCalls > 0 ? (manager.successCalls / manager.totalCalls * 100) : 0,
    avgCallsPerClient: manager.uniqueClients.size > 0 ? (manager.totalCalls / manager.uniqueClients.size) : 0,
    topSource: getTopSource(manager.topSources)
  })).sort((a, b) => b.conversionRate - a.conversionRate);
  
  return {
    totalCalls,
    answeredCalls: answeredCalls || totalCalls,
    successfulCalls,
    phoneAnalysis: processedPhones,
    mangoAnalysis: processedMango,
    trackingNumbers: Object.values(trackingNumbers),
    timeAnalysis,
    qualityAnalysis: processedQuality,
    metrics: {
      answerRate: totalCalls > 0 ? (answeredCalls || totalCalls) / totalCalls * 100 : 0,
      conversionRate: totalCalls > 0 ? successfulCalls / totalCalls * 100 : 0,
      avgCallsPerPhone: processedPhones.length > 0 ? totalCalls / processedPhones.length : 0
    }
  };
}

function getTopSource(sources) {
  if (!sources || Object.keys(sources).length === 0) return 'Не определен';
  
  const sorted = Object.entries(sources).sort((a, b) => b[1] - a[1]);
  return `${sorted[0][0]} (${sorted[0][1]})`;
}

function createCallTrackingReport(analysis) {
  const sheet = createOrUpdateSheet(CONFIG.SHEETS.CALL_TRACKING);
  
  let currentRow = 1;
  
  // Общие метрики
  currentRow = addCallMetrics(sheet, analysis, currentRow);
  currentRow += 2;
  
  // Анализ телефонов
  currentRow = addPhoneAnalysis(sheet, analysis.phoneAnalysis, currentRow);
  currentRow += 2;
  
  // Трекинг-номера (если есть)
  if (analysis.trackingNumbers && analysis.trackingNumbers.length > 0) {
    currentRow = addTrackingNumbersAnalysis(sheet, analysis.trackingNumbers, currentRow);
    currentRow += 2;
  }
  
  // Анализ Mango линий
  currentRow = addMangoAnalysis(sheet, analysis.mangoAnalysis, currentRow);
  currentRow += 2;
  
  // Качество работы менеджеров
  currentRow = addQualityAnalysis(sheet, analysis.qualityAnalysis, currentRow);
  currentRow += 2;
  
  // Анализ по времени
  addTimeCallAnalysis(sheet, analysis.timeAnalysis, currentRow);
  
  return sheet;
}

function addCallMetrics(sheet, analysis, startRow) {
  const headers = ['Метрика звонков', 'Значение'];
  const headerRange = sheet.getRange(startRow, 1, 1, 2);
  headerRange.setValues([headers]);
  headerRange.setBackground('#1e3d59')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const metricsData = [
    ['Всего звонков/лидов', analysis.totalCalls],
    ['Отвеченные звонки', analysis.answeredCalls],
    ['Успешные звонки', analysis.successfulCalls],
    ['Процент ответов', analysis.metrics.answerRate.toFixed(1) + '%'],
    ['Конверсия звонков', analysis.metrics.conversionRate.toFixed(1) + '%'],
    ['Уникальных телефонов', analysis.phoneAnalysis.length],
    ['Среднее звонков на номер', analysis.metrics.avgCallsPerPhone.toFixed(1)]
  ];
  
  const dataRange = sheet.getRange(startRow + 1, 1, metricsData.length, 2);
  dataRange.setValues(metricsData);
  dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  
  return startRow + metricsData.length + 1;
}

function addPhoneAnalysis(sheet, phoneAnalysis, startRow) {
  const headers = [
    'Телефон',
    'Всего звонков',
    'Успешных',
    'Конверсия %',
    'Источники',
    'Менеджеры',
    'Последняя активность'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#0d7377')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  // Топ-50 телефонов
  const topPhones = phoneAnalysis.slice(0, 50);
  const phoneData = topPhones.map(phone => [
    phone.phone,
    phone.calls,
    phone.success,
    phone.conversionRate.toFixed(1) + '%',
    phone.sources.slice(0, 2).join(', '),
    phone.managers.slice(0, 2).join(', '),
    phone.lastActivityFormatted
  ]);
  
  if (phoneData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, phoneData.length, headers.length);
    dataRange.setValues(phoneData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(phoneData.length, 1) + 1;
}

function addTrackingNumbersAnalysis(sheet, trackingNumbers, startRow) {
  const headers = [
    'Трекинг-номер',
    'Код источника', 
    'Описание',
    'Источник',
    'Совпадений в АМО'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#6d4c41')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const trackingData = trackingNumbers.map(tracking => [
    tracking.phone,
    tracking.code,
    tracking.description,
    tracking.source,
    tracking.amoMatches
  ]);
  
  if (trackingData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, trackingData.length, headers.length);
    dataRange.setValues(trackingData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    
    // Выделяем номера с совпадениями
    for (let i = 0; i < trackingData.length; i++) {
      if (trackingData[i][4] > 0) {
        sheet.getRange(startRow + 1 + i, 1, 1, headers.length)
             .setBackground('#e8f5e8');
      }
    }
  }
  
  return startRow + Math.max(trackingData.length, 1) + 1;
}

function addMangoAnalysis(sheet, mangoAnalysis, startRow) {
  if (!mangoAnalysis || mangoAnalysis.length === 0) {
    return startRow;
  }
  
  const headers = [
    'Mango номер',
    'Звонков',
    'Успешных',
    'Конверсия %',
    'Менеджеры',
    'Источники'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#14213d')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const mangoData = mangoAnalysis.map(mango => [
    mango.number,
    mango.calls,
    mango.success,
    mango.conversionRate.toFixed(1) + '%',
    mango.managers.join(', '),
    mango.sources.slice(0, 3).join(', ')
  ]);
  
  if (mangoData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, mangoData.length, headers.length);
    dataRange.setValues(mangoData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  return startRow + Math.max(mangoData.length, 1) + 1;
}

function addQualityAnalysis(sheet, qualityAnalysis, startRow) {
  const headers = [
    'Менеджер',
    'Всего звонков',
    'Успешных',
    'Конверсия %',
    'Уникальных клиентов',
    'Звонков на клиента',
    'Топ источник'
  ];
  
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#2c5f41')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const qualityData = qualityAnalysis.map(manager => [
    manager.manager,
    manager.totalCalls,
    manager.successCalls,
    manager.conversionRate.toFixed(1) + '%',
    manager.uniqueClients,
    manager.avgCallsPerClient.toFixed(1),
    manager.topSource
  ]);
  
  if (qualityData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, qualityData.length, headers.length);
    dataRange.setValues(qualityData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
    
    // Условное форматирование для конверсии
    const conversionRange = sheet.getRange(startRow + 1, 4, qualityData.length, 1);
    const conversionRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#34a853', SpreadsheetApp.InterpolationType.NUMBER, '50')
      .setGradientMidpointWithValue('#fbbc04', SpreadsheetApp.InterpolationType.NUMBER, '25')
      .setGradientMinpointWithValue('#ea4335', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([conversionRange])
      .build();
    
    sheet.setConditionalFormatRules([conversionRule]);
  }
  
  return startRow + Math.max(qualityData.length, 1) + 1;
}

function addTimeCallAnalysis(sheet, timeAnalysis, startRow) {
  const hours = Object.keys(timeAnalysis).map(Number).sort((a, b) => a - b);
  
  if (hours.length === 0) return startRow;
  
  const headers = ['Час', 'Звонков', 'Успешных', 'Конверсия %'];
  const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#8b5a2b')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontFamily(CONFIG.DEFAULT_FONT);
  
  const timeData = hours.map(hour => {
    const data = timeAnalysis[hour];
    const conversionRate = data.calls > 0 ? (data.success / data.calls * 100) : 0;
    return [
      `${hour}:00`,
      data.calls,
      data.success,
      conversionRate.toFixed(1) + '%'
    ];
  });
  
  if (timeData.length > 0) {
    const dataRange = sheet.getRange(startRow + 1, 1, timeData.length, headers.length);
    dataRange.setValues(timeData);
    dataRange.setFontFamily(CONFIG.DEFAULT_FONT);
  }
  
  sheet.autoResizeColumns(1, headers.length);
  
  return startRow + timeData.length + 1;
}

/**
 * МАППИНГ КОДОВ ТРЕКИНГ-НОМЕРОВ В ИСТОЧНИКИ
 */
function mapTrackingSource(code) {
  const mapping = {
    'osn_tel': 'Основной сайт',
    'ya_tel': 'Яндекс Карты', 
    'rclub_tel': 'Рестоклаб',
    '2gis_tel': '2ГИС',
    'smm_tel': 'Соцсети + Google'
  };
  
  return mapping[code] || code || 'Неопределенный источник';
}
