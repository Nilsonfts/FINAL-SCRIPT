/**
 * АНАЛИЗ ПЕРВЫХ КАСАНИЙ
 * Анализ атрибуции первого касания с клиентом
 * @fileoverview Модуль анализа первичных источников привлечения и их влияния на конверсию
 */

/**
 * Основная функция анализа первых касаний
 * Создаёт отчёт по атрибуции первого касания
 */
function analyzeFirstTouchAttribution() {
  try {
    logInfo_('FIRST_TOUCH', 'Начало анализа первых касаний');
    
    const startTime = new Date();
    
    // Получаем данные для анализа первых касаний
    const firstTouchData = getFirstTouchData_();
    
    if (!firstTouchData || firstTouchData.firstTouchAnalysis.length === 0) {
      logWarning_('FIRST_TOUCH', 'Нет данных для анализа первых касаний');
      createEmptyFirstTouchReport_();
      return;
    }
    
    // Создаём или обновляем лист
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.FIRST_TOUCH_ATTRIBUTION);
    clearSheetData_(sheet);
    applySheetFormatting_(sheet, 'Анализ первых касаний');
    
    // Строим структуру отчёта
    createFirstTouchStructure_(sheet, firstTouchData);
    
    // Добавляем диаграммы
    addFirstTouchCharts_(sheet, firstTouchData);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('FIRST_TOUCH', `Анализ первых касаний завершён за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.FIRST_TOUCH_ATTRIBUTION);
    
  } catch (error) {
    logError_('FIRST_TOUCH', 'Ошибка анализа первых касаний', error);
    throw error;
  }
}

/**
 * Получает данные для анализа первых касаний
 * @returns {Object} Данные анализа первых касаний
 * @private
 */
function getFirstTouchData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return null;
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Группируем данные по клиентам для выявления первых касаний
  const clientJourneys = groupClientJourneys_(headers, rows);
  
  // Анализируем первые касания
  const firstTouchAnalysis = analyzeFirstTouchPoints_(clientJourneys);
  
  // Сравнение первого и последнего касания
  const touchComparison = compareFirstLastTouch_(clientJourneys);
  
  // Анализ времени от первого касания до конверсии
  const conversionTimeAnalysis = analyzeFirstTouchConversionTime_(clientJourneys);
  
  // Анализ качества первых касаний
  const qualityAnalysis = analyzeFirstTouchQuality_(clientJourneys);
  
  // Путешествие клиента (customer journey)
  const journeyAnalysis = analyzeCustomerJourney_(clientJourneys);
  
  return {
    firstTouchAnalysis: firstTouchAnalysis,
    touchComparison: touchComparison,
    conversionTimeAnalysis: conversionTimeAnalysis,
    qualityAnalysis: qualityAnalysis,
    journeyAnalysis: journeyAnalysis,
    totalClients: Object.keys(clientJourneys).length,
    analysisDate: getCurrentDateMoscow_()
  };
}

/**
 * Группирует данные по клиентам для анализа путешествия
 * @param {Array} headers - Заголовки таблицы
 * @param {Array} rows - Строки данных
 * @returns {Object} Сгруппированные данные по клиентам
 * @private
 */
function groupClientJourneys_(headers, rows) {
  const clientJourneys = {};
  
  const phoneIndex = findHeaderIndex_(headers, 'Телефон');
  const emailIndex = findHeaderIndex_(headers, 'Email');
  const channelIndex = findHeaderIndex_(headers, 'Канал');
  const sourceIndex = findHeaderIndex_(headers, 'UTM Source');
  const mediumIndex = findHeaderIndex_(headers, 'UTM Medium');
  const campaignIndex = findHeaderIndex_(headers, 'UTM Campaign');
  const createdIndex = findHeaderIndex_(headers, 'Дата создания');
  const statusIndex = findHeaderIndex_(headers, 'Статус');
  const budgetIndex = findHeaderIndex_(headers, 'Бюджет');
  const nameIndex = findHeaderIndex_(headers, 'Имя');
  const companyIndex = findHeaderIndex_(headers, 'Компания');
  
  rows.forEach(row => {
    const phone = row[phoneIndex] || '';
    const email = row[emailIndex] || '';
    const channel = row[channelIndex] || 'Неизвестно';
    const source = row[sourceIndex] || 'direct';
    const medium = row[mediumIndex] || 'none';
    const campaign = row[campaignIndex] || 'none';
    const createdDate = parseDate_(row[createdIndex]);
    const status = row[statusIndex] || '';
    const budget = parseFloat(row[budgetIndex]) || 0;
    const name = row[nameIndex] || '';
    const company = row[companyIndex] || '';
    
    // Определяем уникального клиента (по телефону или email)
    const clientId = phone || email;
    if (!clientId || !createdDate) return;
    
    if (!clientJourneys[clientId]) {
      clientJourneys[clientId] = {
        clientId: clientId,
        phone: phone,
        email: email,
        name: name,
        company: company,
        touches: [],
        firstTouch: null,
        lastTouch: null,
        finalStatus: '',
        totalValue: 0,
        touchPoints: 0
      };
    }
    
    const journey = clientJourneys[clientId];
    
    // Добавляем касание
    const touch = {
      date: createdDate,
      channel: channel,
      source: source,
      medium: medium,
      campaign: campaign,
      status: status,
      budget: budget,
      touchPoint: `${source}/${medium}/${campaign}`
    };
    
    journey.touches.push(touch);
    
    // Обновляем общую информацию о клиенте
    if (status === 'success') {
      journey.totalValue += budget;
      journey.finalStatus = 'success';
    } else if (status === 'failure' && journey.finalStatus !== 'success') {
      journey.finalStatus = 'failure';
    } else if (!journey.finalStatus) {
      journey.finalStatus = status || 'in_progress';
    }
    
    // Обновляем имя и компанию если они есть
    if (name && !journey.name) journey.name = name;
    if (company && !journey.company) journey.company = company;
  });
  
  // Сортируем касания по дате и определяем первое/последнее
  Object.keys(clientJourneys).forEach(clientId => {
    const journey = clientJourneys[clientId];
    journey.touches.sort((a, b) => a.date - b.date);
    
    journey.firstTouch = journey.touches[0];
    journey.lastTouch = journey.touches[journey.touches.length - 1];
    journey.touchPoints = journey.touches.length;
    
    // Удаляем дубликаты касаний (одинаковые источники в один день)
    journey.uniqueTouches = journey.touches.filter((touch, index, arr) => {
      if (index === 0) return true;
      const prev = arr[index - 1];
      return !(
        touch.source === prev.source &&
        touch.medium === prev.medium &&
        touch.campaign === prev.campaign &&
        Math.abs(touch.date - prev.date) < 24 * 60 * 60 * 1000 // меньше суток
      );
    });
  });
  
  return clientJourneys;
}

/**
 * Анализирует первые точки касания
 * @param {Object} clientJourneys - Путешествия клиентов
 * @returns {Array} Анализ первых касаний
 * @private
 */
function analyzeFirstTouchPoints_(clientJourneys) {
  const firstTouchStats = {};
  
  Object.values(clientJourneys).forEach(journey => {
    if (!journey.firstTouch) return;
    
    const firstTouch = journey.firstTouch;
    const key = `${firstTouch.channel}`;
    
    if (!firstTouchStats[key]) {
      firstTouchStats[key] = {
        channel: firstTouch.channel,
        source: firstTouch.source,
        medium: firstTouch.medium,
        totalFirstTouches: 0,
        convertedClients: 0,
        totalRevenue: 0,
        avgRevenue: 0,
        conversionRate: 0,
        avgTouchPoints: 0,
        avgConversionTime: 0,
        conversionTimes: [],
        touchPointsSum: 0
      };
    }
    
    const stats = firstTouchStats[key];
    stats.totalFirstTouches++;
    stats.touchPointsSum += journey.touchPoints;
    
    if (journey.finalStatus === 'success') {
      stats.convertedClients++;
      stats.totalRevenue += journey.totalValue;
      
      // Время до конверсии
      const conversionTime = Math.round((journey.lastTouch.date - journey.firstTouch.date) / (1000 * 60 * 60 * 24));
      if (conversionTime >= 0) {
        stats.conversionTimes.push(conversionTime);
      }
    }
  });
  
  // Вычисляем производные метрики
  return Object.values(firstTouchStats).map(stats => {
    stats.conversionRate = stats.totalFirstTouches > 0 ? 
      (stats.convertedClients / stats.totalFirstTouches * 100) : 0;
    stats.avgRevenue = stats.convertedClients > 0 ? 
      (stats.totalRevenue / stats.convertedClients) : 0;
    stats.avgTouchPoints = stats.totalFirstTouches > 0 ? 
      Math.round(stats.touchPointsSum / stats.totalFirstTouches) : 0;
    stats.avgConversionTime = stats.conversionTimes.length > 0 ?
      Math.round(stats.conversionTimes.reduce((a, b) => a + b, 0) / stats.conversionTimes.length) : 0;
    
    return stats;
  }).sort((a, b) => b.totalFirstTouches - a.totalFirstTouches);
}

/**
 * Сравнивает первое и последнее касание
 * @param {Object} clientJourneys - Путешествия клиентов
 * @returns {Object} Сравнительный анализ
 * @private
 */
function compareFirstLastTouch_(clientJourneys) {
  const comparison = {
    sameChannelConversions: 0,
    differentChannelConversions: 0,
    channelSwitchPatterns: {},
    assistingChannels: {}
  };
  
  const convertedJourneys = Object.values(clientJourneys)
    .filter(journey => journey.finalStatus === 'success');
  
  convertedJourneys.forEach(journey => {
    const firstChannel = journey.firstTouch.channel;
    const lastChannel = journey.lastTouch.channel;
    
    if (firstChannel === lastChannel) {
      comparison.sameChannelConversions++;
    } else {
      comparison.differentChannelConversions++;
      
      // Паттерн переключения каналов
      const switchPattern = `${firstChannel} → ${lastChannel}`;
      comparison.channelSwitchPatterns[switchPattern] = 
        (comparison.channelSwitchPatterns[switchPattern] || 0) + 1;
      
      // Ассистирующие каналы (между первым и последним)
      journey.uniqueTouches.slice(1, -1).forEach(touch => {
        const assistingChannel = touch.channel;
        if (assistingChannel !== firstChannel && assistingChannel !== lastChannel) {
          comparison.assistingChannels[assistingChannel] = 
            (comparison.assistingChannels[assistingChannel] || 0) + 1;
        }
      });
    }
  });
  
  // Сортируем паттерны и ассистирующие каналы
  comparison.topSwitchPatterns = Object.entries(comparison.channelSwitchPatterns)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 10)
    .map(([pattern, count]) => ({ pattern, count }));
  
  comparison.topAssistingChannels = Object.entries(comparison.assistingChannels)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 5)
    .map(([channel, count]) => ({ channel, count }));
  
  return comparison;
}

/**
 * Анализирует время от первого касания до конверсии
 * @param {Object} clientJourneys - Путешествия клиентов
 * @returns {Object} Анализ времени конверсии
 * @private
 */
function analyzeFirstTouchConversionTime_(clientJourneys) {
  const timeAnalysis = {
    conversionTimes: [],
    timeDistribution: {
      'Мгновенно (0 дней)': 0,
      '1-3 дня': 0,
      '4-7 дней': 0,
      '8-14 дней': 0,
      '15-30 дней': 0,
      '31-60 дней': 0,
      'Более 60 дней': 0
    },
    channelTimeAnalysis: {}
  };
  
  Object.values(clientJourneys)
    .filter(journey => journey.finalStatus === 'success')
    .forEach(journey => {
      const conversionTime = Math.round((journey.lastTouch.date - journey.firstTouch.date) / (1000 * 60 * 60 * 24));
      if (conversionTime < 0) return; // Некорректные данные
      
      timeAnalysis.conversionTimes.push(conversionTime);
      
      // Распределение по временным интервалам
      if (conversionTime === 0) {
        timeAnalysis.timeDistribution['Мгновенно (0 дней)']++;
      } else if (conversionTime <= 3) {
        timeAnalysis.timeDistribution['1-3 дня']++;
      } else if (conversionTime <= 7) {
        timeAnalysis.timeDistribution['4-7 дней']++;
      } else if (conversionTime <= 14) {
        timeAnalysis.timeDistribution['8-14 дней']++;
      } else if (conversionTime <= 30) {
        timeAnalysis.timeDistribution['15-30 дней']++;
      } else if (conversionTime <= 60) {
        timeAnalysis.timeDistribution['31-60 дней']++;
      } else {
        timeAnalysis.timeDistribution['Более 60 дней']++;
      }
      
      // Анализ по каналам первого касания
      const firstChannel = journey.firstTouch.channel;
      if (!timeAnalysis.channelTimeAnalysis[firstChannel]) {
        timeAnalysis.channelTimeAnalysis[firstChannel] = {
          times: [],
          avgTime: 0,
          medianTime: 0,
          minTime: Infinity,
          maxTime: 0
        };
      }
      
      const channelAnalysis = timeAnalysis.channelTimeAnalysis[firstChannel];
      channelAnalysis.times.push(conversionTime);
      channelAnalysis.minTime = Math.min(channelAnalysis.minTime, conversionTime);
      channelAnalysis.maxTime = Math.max(channelAnalysis.maxTime, conversionTime);
    });
  
  // Вычисляем статистики
  if (timeAnalysis.conversionTimes.length > 0) {
    timeAnalysis.avgConversionTime = Math.round(
      timeAnalysis.conversionTimes.reduce((a, b) => a + b, 0) / timeAnalysis.conversionTimes.length
    );
    
    const sortedTimes = timeAnalysis.conversionTimes.sort((a, b) => a - b);
    timeAnalysis.medianConversionTime = sortedTimes[Math.floor(sortedTimes.length / 2)];
  }
  
  // Статистики по каналам
  Object.keys(timeAnalysis.channelTimeAnalysis).forEach(channel => {
    const analysis = timeAnalysis.channelTimeAnalysis[channel];
    if (analysis.times.length > 0) {
      analysis.avgTime = Math.round(analysis.times.reduce((a, b) => a + b, 0) / analysis.times.length);
      
      const sortedTimes = analysis.times.sort((a, b) => a - b);
      analysis.medianTime = sortedTimes[Math.floor(sortedTimes.length / 2)];
    }
  });
  
  return timeAnalysis;
}

/**
 * Анализирует качество первых касаний
 * @param {Object} clientJourneys - Путешествия клиентов
 * @returns {Object} Анализ качества
 * @private
 */
function analyzeFirstTouchQuality_(clientJourneys) {
  const qualityAnalysis = {
    touchPointsDistribution: {
      '1 касание': 0,
      '2-3 касания': 0,
      '4-5 касаний': 0,
      '6-10 касаний': 0,
      'Более 10 касаний': 0
    },
    qualityByFirstTouch: {},
    multiTouchConversions: 0,
    singleTouchConversions: 0
  };
  
  Object.values(clientJourneys).forEach(journey => {
    const touchCount = journey.uniqueTouches.length;
    
    // Распределение по количеству касаний
    if (touchCount === 1) {
      qualityAnalysis.touchPointsDistribution['1 касание']++;
      if (journey.finalStatus === 'success') {
        qualityAnalysis.singleTouchConversions++;
      }
    } else {
      if (journey.finalStatus === 'success') {
        qualityAnalysis.multiTouchConversions++;
      }
      
      if (touchCount <= 3) {
        qualityAnalysis.touchPointsDistribution['2-3 касания']++;
      } else if (touchCount <= 5) {
        qualityAnalysis.touchPointsDistribution['4-5 касаний']++;
      } else if (touchCount <= 10) {
        qualityAnalysis.touchPointsDistribution['6-10 касаний']++;
      } else {
        qualityAnalysis.touchPointsDistribution['Более 10 касаний']++;
      }
    }
    
    // Качество по первому касанию
    const firstChannel = journey.firstTouch.channel;
    if (!qualityAnalysis.qualityByFirstTouch[firstChannel]) {
      qualityAnalysis.qualityByFirstTouch[firstChannel] = {
        channel: firstChannel,
        singleTouch: 0,
        multiTouch: 0,
        singleTouchConversions: 0,
        multiTouchConversions: 0,
        avgTouchPoints: 0,
        totalTouchPoints: 0,
        totalClients: 0
      };
    }
    
    const channelQuality = qualityAnalysis.qualityByFirstTouch[firstChannel];
    channelQuality.totalClients++;
    channelQuality.totalTouchPoints += touchCount;
    
    if (touchCount === 1) {
      channelQuality.singleTouch++;
      if (journey.finalStatus === 'success') {
        channelQuality.singleTouchConversions++;
      }
    } else {
      channelQuality.multiTouch++;
      if (journey.finalStatus === 'success') {
        channelQuality.multiTouchConversions++;
      }
    }
  });
  
  // Вычисляем производные метрики
  Object.keys(qualityAnalysis.qualityByFirstTouch).forEach(channel => {
    const quality = qualityAnalysis.qualityByFirstTouch[channel];
    quality.avgTouchPoints = quality.totalClients > 0 ? 
      Math.round(quality.totalTouchPoints / quality.totalClients) : 0;
    quality.singleTouchRate = quality.totalClients > 0 ? 
      (quality.singleTouch / quality.totalClients * 100) : 0;
    quality.singleTouchConversionRate = quality.singleTouch > 0 ? 
      (quality.singleTouchConversions / quality.singleTouch * 100) : 0;
    quality.multiTouchConversionRate = quality.multiTouch > 0 ? 
      (quality.multiTouchConversions / quality.multiTouch * 100) : 0;
  });
  
  return qualityAnalysis;
}

/**
 * Анализирует путешествие клиента
 * @param {Object} clientJourneys - Путешествия клиентов
 * @returns {Object} Анализ путешествия
 * @private
 */
function analyzeCustomerJourney_(clientJourneys) {
  const journeyAnalysis = {
    commonPathways: {},
    channelSequences: {},
    dropOffPoints: {}
  };
  
  Object.values(clientJourneys).forEach(journey => {
    if (journey.uniqueTouches.length < 2) return;
    
    // Последовательности каналов
    const channelSequence = journey.uniqueTouches.map(touch => touch.channel).join(' → ');
    journeyAnalysis.channelSequences[channelSequence] = 
      (journeyAnalysis.channelSequences[channelSequence] || 0) + 1;
    
    // Общие пути (первые 3 касания)
    const pathway = journey.uniqueTouches.slice(0, 3).map(touch => touch.channel).join(' → ');
    if (!journeyAnalysis.commonPathways[pathway]) {
      journeyAnalysis.commonPathways[pathway] = {
        path: pathway,
        count: 0,
        conversions: 0,
        conversionRate: 0
      };
    }
    
    journeyAnalysis.commonPathways[pathway].count++;
    if (journey.finalStatus === 'success') {
      journeyAnalysis.commonPathways[pathway].conversions++;
    }
    
    // Точки оттока (где клиенты перестали взаимодействовать)
    if (journey.finalStatus === 'failure') {
      const lastChannel = journey.lastTouch.channel;
      journeyAnalysis.dropOffPoints[lastChannel] = 
        (journeyAnalysis.dropOffPoints[lastChannel] || 0) + 1;
    }
  });
  
  // Топ пути
  journeyAnalysis.topPathways = Object.values(journeyAnalysis.commonPathways)
    .map(pathway => {
      pathway.conversionRate = pathway.count > 0 ? 
        (pathway.conversions / pathway.count * 100) : 0;
      return pathway;
    })
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);
  
  // Топ последовательности
  journeyAnalysis.topSequences = Object.entries(journeyAnalysis.channelSequences)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 8)
    .map(([sequence, count]) => ({ sequence, count }));
  
  // Топ точки оттока
  journeyAnalysis.topDropOffPoints = Object.entries(journeyAnalysis.dropOffPoints)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 5)
    .map(([channel, count]) => ({ channel, count }));
  
  return journeyAnalysis;
}

/**
 * Создаёт структуру отчёта по первым касаниям
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} firstTouchData - Данные анализа первых касаний
 * @private
 */
function createFirstTouchStructure_(sheet, firstTouchData) {
  let currentRow = 3;
  
  // 1. Общая статистика
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('📊 СТАТИСТИКА ПЕРВЫХ КАСАНИЙ');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const overallStats = [
    ['Всего уникальных клиентов:', firstTouchData.totalClients],
    ['Конверсий из одного касания:', firstTouchData.qualityAnalysis.singleTouchConversions],
    ['Мультитач конверсий:', firstTouchData.qualityAnalysis.multiTouchConversions],
    ['Одинаковый канал (первый=последний):', firstTouchData.touchComparison.sameChannelConversions],
    ['Смена каналов:', firstTouchData.touchComparison.differentChannelConversions],
    ['Среднее время до конверсии:', `${firstTouchData.conversionTimeAnalysis.avgConversionTime || 0} дней`]
  ];
  
  sheet.getRange(currentRow, 1, overallStats.length, 2).setValues(overallStats);
  sheet.getRange(currentRow, 1, overallStats.length, 1).setFontWeight('bold');
  currentRow += overallStats.length + 2;
  
  // 2. Эффективность первых касаний по каналам
  sheet.getRange(currentRow, 1, 1, 7).merge();
  sheet.getRange(currentRow, 1).setValue('🎯 ЭФФЕКТИВНОСТЬ ПЕРВЫХ КАСАНИЙ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 7));
  currentRow++;
  
  const firstTouchHeaders = [['Канал первого касания', 'Первых касаний', 'Конверсий', 'Выручка', 'Конверсия %', 'Среднее касаний', 'Время до сделки']];
  sheet.getRange(currentRow, 1, 1, 7).setValues(firstTouchHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 7));
  currentRow++;
  
  const firstTouchData_array = firstTouchData.firstTouchAnalysis.slice(0, 10).map(item => [
    item.channel,
    item.totalFirstTouches,
    item.convertedClients,
    formatCurrency_(item.totalRevenue),
    `${item.conversionRate.toFixed(1)}%`,
    item.avgTouchPoints,
    item.avgConversionTime > 0 ? `${item.avgConversionTime} дн.` : 'Н/Д'
  ]);
  
  if (firstTouchData_array.length > 0) {
    sheet.getRange(currentRow, 1, firstTouchData_array.length, 7).setValues(firstTouchData_array);
    currentRow += firstTouchData_array.length;
  }
  currentRow += 2;
  
  // 3. Распределение времени до конверсии
  sheet.getRange(currentRow, 1, 1, 3).merge();
  sheet.getRange(currentRow, 1).setValue('⏰ ВРЕМЯ ДО КОНВЕРСИИ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  const timeDistHeaders = [['Временной интервал', 'Конверсий', 'Доля']];
  sheet.getRange(currentRow, 1, 1, 3).setValues(timeDistHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  const totalConversions = Object.values(firstTouchData.conversionTimeAnalysis.timeDistribution)
    .reduce((a, b) => a + b, 0);
  
  const timeDistData = Object.entries(firstTouchData.conversionTimeAnalysis.timeDistribution)
    .filter(([, count]) => count > 0)
    .map(([interval, count]) => [
      interval,
      count,
      totalConversions > 0 ? `${Math.round(count / totalConversions * 100)}%` : '0%'
    ]);
  
  if (timeDistData.length > 0) {
    sheet.getRange(currentRow, 1, timeDistData.length, 3).setValues(timeDistData);
    currentRow += timeDistData.length;
  }
  currentRow += 2;
  
  // 4. Популярные пути клиентов
  if (firstTouchData.journeyAnalysis.topPathways.length > 0) {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue('🗺️ ПОПУЛЯРНЫЕ ПУТИ КЛИЕНТОВ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const pathwaysHeaders = [['Путь клиента', 'Клиентов', 'Конверсий', 'Конверсия %']];
    sheet.getRange(currentRow, 1, 1, 4).setValues(pathwaysHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const pathwaysData = firstTouchData.journeyAnalysis.topPathways.map(item => [
      item.path,
      item.count,
      item.conversions,
      `${item.conversionRate.toFixed(1)}%`
    ]);
    
    if (pathwaysData.length > 0) {
      sheet.getRange(currentRow, 1, pathwaysData.length, 4).setValues(pathwaysData);
      currentRow += pathwaysData.length;
    }
    currentRow += 2;
  }
  
  // 5. Паттерны смены каналов
  if (firstTouchData.touchComparison.topSwitchPatterns.length > 0) {
    sheet.getRange(currentRow, 1, 1, 3).merge();
    sheet.getRange(currentRow, 1).setValue('🔄 СМЕНА КАНАЛОВ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
    currentRow++;
    
    const switchHeaders = [['Паттерн смены', 'Клиентов', 'Доля от смен']];
    sheet.getRange(currentRow, 1, 1, 3).setValues(switchHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
    currentRow++;
    
    const totalSwitches = firstTouchData.touchComparison.differentChannelConversions;
    const switchData = firstTouchData.touchComparison.topSwitchPatterns.map(item => [
      item.pattern,
      item.count,
      totalSwitches > 0 ? `${Math.round(item.count / totalSwitches * 100)}%` : '0%'
    ]);
    
    if (switchData.length > 0) {
      sheet.getRange(currentRow, 1, switchData.length, 3).setValues(switchData);
      currentRow += switchData.length;
    }
    currentRow += 2;
  }
  
  // 6. Распределение по количеству касаний
  sheet.getRange(currentRow, 1, 1, 3).merge();
  sheet.getRange(currentRow, 1).setValue('📈 РАСПРЕДЕЛЕНИЕ ПО КАСАНИЯМ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  const touchDistHeaders = [['Количество касаний', 'Клиентов', 'Доля']];
  sheet.getRange(currentRow, 1, 1, 3).setValues(touchDistHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  const totalTouchDistribution = Object.values(firstTouchData.qualityAnalysis.touchPointsDistribution)
    .reduce((a, b) => a + b, 0);
  
  const touchDistData = Object.entries(firstTouchData.qualityAnalysis.touchPointsDistribution)
    .filter(([, count]) => count > 0)
    .map(([range, count]) => [
      range,
      count,
      totalTouchDistribution > 0 ? `${Math.round(count / totalTouchDistribution * 100)}%` : '0%'
    ]);
  
  if (touchDistData.length > 0) {
    sheet.getRange(currentRow, 1, touchDistData.length, 3).setValues(touchDistData);
    currentRow += touchDistData.length;
  }
  currentRow += 2;
  
  // 7. Детализация по топ каналам первого касания
  const topFirstTouchChannels = firstTouchData.firstTouchAnalysis.slice(0, 3);
  topFirstTouchChannels.forEach(channelData => {
    sheet.getRange(currentRow, 1, 1, 4).merge();
    sheet.getRange(currentRow, 1).setValue(`🔍 ДЕТАЛИЗАЦИЯ: ${channelData.channel.toUpperCase()}`);
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
    currentRow++;
    
    const channelDetails = [
      ['Показатели первого касания:', ''],
      [`• Первых касаний`, channelData.totalFirstTouches],
      [`• Конверсий`, `${channelData.convertedClients} (${channelData.conversionRate.toFixed(1)}%)`],
      [`• Выручка`, formatCurrency_(channelData.totalRevenue)],
      [`• Средний чек`, formatCurrency_(channelData.avgRevenue)],
      [`• Среднее касаний до конверсии`, channelData.avgTouchPoints],
      [`• Среднее время до сделки`, channelData.avgConversionTime > 0 ? `${channelData.avgConversionTime} дней` : 'Н/Д']
    ];
    
    sheet.getRange(currentRow, 1, channelDetails.length, 2).setValues(channelDetails);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    currentRow += channelDetails.length + 1;
  });
  
  // 8. Ассистирующие каналы
  if (firstTouchData.touchComparison.topAssistingChannels.length > 0) {
    sheet.getRange(currentRow, 1, 1, 3).merge();
    sheet.getRange(currentRow, 1).setValue('🤝 АССИСТИРУЮЩИЕ КАНАЛЫ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
    currentRow++;
    
    sheet.getRange(currentRow, 1).setValue('Каналы, которые помогают в процессе конверсии:');
    currentRow++;
    
    const assistingHeaders = [['Канал', 'Ассистов', 'Роль в воронке']];
    sheet.getRange(currentRow, 1, 1, 3).setValues(assistingHeaders);
    applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
    currentRow++;
    
    const assistingData = firstTouchData.touchComparison.topAssistingChannels.map(item => [
      item.channel,
      item.count,
      'Промежуточное влияние'
    ]);
    
    if (assistingData.length > 0) {
      sheet.getRange(currentRow, 1, assistingData.length, 3).setValues(assistingData);
      currentRow += assistingData.length;
    }
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 7);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 7);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к анализу первых касаний
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} firstTouchData - Данные анализа первых касаний
 * @private
 */
function addFirstTouchCharts_(sheet, firstTouchData) {
  // 1. Круговая диаграмма первых касаний
  if (firstTouchData.firstTouchAnalysis.length > 0) {
    const firstTouchChartData = [['Канал', 'Первые касания']].concat(
      firstTouchData.firstTouchAnalysis.slice(0, 8).map(item => [item.channel, item.totalFirstTouches])
    );
    
    // Создаем диаграмму через универсальную функцию
    const firstTouchChart = createChart_(sheet, 'pie', firstTouchChartData, {
      startRow: 1,
      startCol: 9,
      title: 'Распределение первых касаний по каналам',
      position: { row: 3, col: 9 },
      width: 500,
      height: 350,
      legend: 'right'
    });
  }
  
  // 2. Столбчатая диаграмма конверсии первых касаний
  if (firstTouchData.firstTouchAnalysis.length > 0) {
    const conversionChartData = [['Канал', 'Конверсия %']].concat(
      firstTouchData.firstTouchAnalysis.slice(0, 10).map(item => [item.channel, item.conversionRate])
    );
    
    // Создаем диаграмму через универсальную функцию
    const conversionChart = createChart_(sheet, 'column', conversionChartData, {
      startRow: 1,
      startCol: 12,
      title: 'Конверсия первых касаний по каналам',
      position: { row: 3, col: 15 },
      width: 600,
      height: 350,
      legend: 'none',
      hAxisTitle: 'Канал',
      vAxisTitle: 'Конверсия, %'
    });
  }
  
  // 3. Диаграмма времени до конверсии
  if (Object.keys(firstTouchData.conversionTimeAnalysis.timeDistribution).length > 0) {
    const timeDistChartData = [['Временной интервал', 'Количество']].concat(
      Object.entries(firstTouchData.conversionTimeAnalysis.timeDistribution)
        .filter(([, count]) => count > 0)
        .map(([interval, count]) => [interval, count])
    );
    
    // Создаем диаграмму через универсальную функцию
    const timeDistChart = createChart_(sheet, 'column', timeDistChartData, {
      startRow: 1,
      startCol: 15,
      title: 'Распределение времени до конверсии',
      position: { row: 25, col: 9 },
      width: 700,
      height: 350,
      legend: 'none',
      hAxisTitle: 'Временной интервал',
      vAxisTitle: 'Количество конверсий'
    });
  }
  
  // 4. Диаграмма количества касаний
  if (Object.keys(firstTouchData.qualityAnalysis.touchPointsDistribution).length > 0) {
    const touchPointsChartData = [['Касания', 'Клиенты']].concat(
      Object.entries(firstTouchData.qualityAnalysis.touchPointsDistribution)
        .filter(([, count]) => count > 0)
        .map(([range, count]) => [range, count])
    );
    
    // Создаем диаграмму через универсальную функцию
    const touchPointsChart = createChart_(sheet, 'pie', touchPointsChartData, {
      startRow: 1,
      startCol: 18,
      title: 'Распределение клиентов по количеству касаний',
      position: { row: 25, col: 18 },
      width: 500,
      height: 350,
      legend: 'right'
    });
  }
}

/**
 * Создаёт пустой отчёт когда нет данных
 * @private
 */
function createEmptyFirstTouchReport_() {
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.FIRST_TOUCH_ATTRIBUTION);
  clearSheetData_(sheet);
  applySheetFormatting_(sheet, 'Анализ первых касаний');
  
  sheet.getRange('A3').setValue('ℹ️ Нет данных для анализа первых касаний');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A5').setValue('Убедитесь, что:');
  sheet.getRange('A6').setValue('• В системе есть данные с повторными клиентами');
  sheet.getRange('A7').setValue('• У клиентов указаны телефоны или email для идентификации');
  sheet.getRange('A8').setValue('• Данные содержат даты и источники касаний');
  
  updateLastUpdateTime_(CONFIG.SHEETS.FIRST_TOUCH_ATTRIBUTION);
}
