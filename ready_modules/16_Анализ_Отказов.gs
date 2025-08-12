/**
 * ❌ Анализ Отказов - Анализ причин отказов
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Apps Script: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit
 * 2. Создайте новый файл: + → Script file
 * 3. Назовите файл: 16_Анализ_Отказов (без .gs)
 * 4. Удалите весь код по умолчанию
 * 5. Скопируйте и вставьте код ниже
 * 6. Сохраните файл (Ctrl+S)
 */

/**
 * АНАЛИЗ ПРИЧИН ОТКАЗОВ С ИСПОЛЬЗОВАНИЕМ GPT
 * AI-анализ текстовых комментариев и выявление паттернов отказов
 * @fileoverview Модуль для интеллектуального анализа причин отказов клиентов
 */

/**
 * Основная функция анализа причин отказов
 * Анализирует отклонённые сделки и группирует причины
 */
function analyzeRefusalReasons() {
  try {
    logInfo_('REFUSAL_ANALYSIS', 'Начало анализа причин отказов');
    
    const startTime = new Date();
    
    // Получаем данные отклонённых сделок
    const refusedDeals = getRefusedDealsData_();
    
    if (!refusedDeals || refusedDeals.length === 0) {
      logWarning_('REFUSAL_ANALYSIS', 'Нет данных об отклонённых сделках');
      createEmptyRefusalReport_();
      return;
    }
    
    // Анализируем причины с помощью GPT
    const analysisResults = await analyzeRefusalReasonsWithGPT_(refusedDeals);
    
    // Создаём отчёт
    const sheet = getOrCreateSheet_(CONFIG.SHEETS.REFUSAL_ANALYSIS);
    clearSheetData_(sheet);
    applySheetFormatting_(sheet, 'Анализ причин отказов');
    
    // Строим структуру отчёта
    createRefusalAnalysisStructure_(sheet, analysisResults);
    
    // Добавляем диаграммы
    addRefusalAnalysisCharts_(sheet, analysisResults);
    
    const duration = (new Date() - startTime) / 1000;
    logInfo_('REFUSAL_ANALYSIS', `Анализ причин отказов завершён за ${duration}с`);
    
    // Обновляем время последнего обновления
    updateLastUpdateTime_(CONFIG.SHEETS.REFUSAL_ANALYSIS);
    
  } catch (error) {
    logError_('REFUSAL_ANALYSIS', 'Ошибка анализа причин отказов', error);
    throw error;
  }
}

/**
 * Получает данные отклонённых сделок для анализа
 * @returns {Array} Массив отклонённых сделок
 * @private
 */
function getRefusedDealsData_() {
  const rawSheet = getSheet_(CONFIG.SHEETS.RAW_DATA);
  if (!rawSheet) {
    throw new Error('Лист с сырыми данными не найден');
  }
  
  const rawData = getSheetData_(rawSheet);
  if (rawData.length <= 1) return [];
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  // Фильтруем только отклонённые сделки
  const refusedDeals = rows.filter(row => {
    const statusIndex = findHeaderIndex_(headers, 'Статус');
    return statusIndex >= 0 && row[statusIndex] === 'failure';
  });
  
  // Формируем структурированные данные
  return refusedDeals.map(row => {
    const dealId = row[findHeaderIndex_(headers, 'ID сделки')] || '';
    const dealName = row[findHeaderIndex_(headers, 'Название')] || '';
    const channel = row[findHeaderIndex_(headers, 'Канал')] || 'Неизвестно';
    const createdDate = parseDate_(row[findHeaderIndex_(headers, 'Дата создания')]);
    const budget = parseFloat(row[findHeaderIndex_(headers, 'Бюджет')]) || 0;
    const manager = row[findHeaderIndex_(headers, 'Ответственный')] || 'Неназначен';
    
    // Пытаемся извлечь комментарий о причине отказа
    const refusalReasonIndex = findHeaderIndex_(headers, 'Причина отказа');
    const commentIndex = findHeaderIndex_(headers, 'Комментарий');
    const notesIndex = findHeaderIndex_(headers, 'Примечания');
    
    let refusalComment = '';
    if (refusalReasonIndex >= 0 && row[refusalReasonIndex]) {
      refusalComment = row[refusalReasonIndex];
    } else if (commentIndex >= 0 && row[commentIndex]) {
      refusalComment = row[commentIndex];
    } else if (notesIndex >= 0 && row[notesIndex]) {
      refusalComment = row[notesIndex];
    }
    
    return {
      dealId: dealId,
      dealName: dealName,
      channel: channel,
      createdDate: createdDate,
      budget: budget,
      manager: manager,
      refusalComment: refusalComment,
      utmSource: row[findHeaderIndex_(headers, 'UTM Source')] || '',
      utmCampaign: row[findHeaderIndex_(headers, 'UTM Campaign')] || ''
    };
  }).filter(deal => deal.refusalComment && deal.refusalComment.trim().length > 0);
}

/**
 * Анализирует причины отказов с помощью GPT
 * @param {Array} refusedDeals - Массив отклонённых сделок
 * @returns {Object} Результаты анализа
 * @private
 */
async function analyzeRefusalReasonsWithGPT_(refusedDeals) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('Не настроен API ключ OpenAI');
  }
  
  // Группируем комментарии для анализа
  const comments = refusedDeals
    .map(deal => deal.refusalComment)
    .filter(comment => comment && comment.trim().length > 5);
  
  if (comments.length === 0) {
    return {
      totalRefusals: refusedDeals.length,
      categorizedReasons: {},
      insights: [],
      recommendations: [],
      channelAnalysis: {},
      monthlyTrends: {}
    };
  }
  
  // Разбиваем на батчи для анализа (GPT имеет ограничения по токенам)
  const batchSize = 20;
  const batches = [];
  for (let i = 0; i < comments.length; i += batchSize) {
    batches.push(comments.slice(i, i + batchSize));
  }
  
  let categorizedReasons = {};
  let insights = [];
  let recommendations = [];
  
  // Анализируем каждый батч
  for (let i = 0; i < batches.length; i++) {
    logInfo_('GPT_ANALYSIS', `Анализ батча ${i + 1} из ${batches.length}`);
    
    try {
      const batchResults = await analyzeRefusalBatch_(apiKey, batches[i], refusedDeals);
      
      // Объединяем результаты
      Object.keys(batchResults.categories).forEach(category => {
        if (!categorizedReasons[category]) {
          categorizedReasons[category] = [];
        }
        categorizedReasons[category] = categorizedReasons[category].concat(batchResults.categories[category]);
      });
      
      insights = insights.concat(batchResults.insights);
      recommendations = recommendations.concat(batchResults.recommendations);
      
      // Пауза между запросами
      Utilities.sleep(1000);
      
    } catch (error) {
      logError_('GPT_ANALYSIS', `Ошибка анализа батча ${i + 1}`, error);
      continue;
    }
  }
  
  // Дополнительный анализ по каналам и времени
  const channelAnalysis = analyzeRefusalsByChannel_(refusedDeals);
  const monthlyTrends = analyzeRefusalTrends_(refusedDeals);
  
  return {
    totalRefusals: refusedDeals.length,
    categorizedReasons: categorizedReasons,
    insights: insights,
    recommendations: recommendations,
    channelAnalysis: channelAnalysis,
    monthlyTrends: monthlyTrends
  };
}

/**
 * Анализирует батч комментариев с помощью GPT
 * @param {string} apiKey - API ключ OpenAI
 * @param {Array} commentsBatch - Батч комментариев
 * @param {Array} dealsData - Данные сделок для контекста
 * @returns {Object} Результаты анализа батча
 * @private
 */
async function analyzeRefusalBatch_(apiKey, commentsBatch, dealsData) {
  const prompt = createGPTPromptForRefusalAnalysis_(commentsBatch);
  
  const payload = {
    model: CONFIG.GPT.MODEL,
    messages: [
      {
        role: 'system',
        content: `Ты эксперт по анализу клиентских отказов в сфере услуг. 
        Твоя задача - проанализировать причины отказов клиентов и сгруппировать их по категориям.
        
        Отвечай строго в формате JSON со следующей структурой:
        {
          "categories": {
            "Категория1": ["причина1", "причина2"],
            "Категория2": ["причина3", "причина4"]
          },
          "insights": ["инсайт1", "инсайт2"],
          "recommendations": ["рекомендация1", "рекомендация2"]
        }
        
        Используй следующие основные категории отказов:
        - "Цена" - всё связанное со стоимостью
        - "Качество" - претензии к качеству услуг
        - "Сроки" - проблемы со временем выполнения
        - "Коммуникация" - проблемы в общении с клиентом
        - "Конкуренты" - ушли к конкурентам
        - "Личные причины" - изменились обстоятельства клиента
        - "Технические" - технические проблемы
        - "Недоверие" - сомнения в компании
        - "Прочее" - остальные причины`
      },
      {
        role: 'user',
        content: prompt
      }
    ],
    max_tokens: CONFIG.GPT.MAX_TOKENS,
    temperature: CONFIG.GPT.TEMPERATURE
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  };
  
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`OpenAI API error: ${response.getResponseCode()} - ${response.getContentText()}`);
  }
  
  const responseData = JSON.parse(response.getContentText());
  
  if (!responseData.choices || responseData.choices.length === 0) {
    throw new Error('Пустой ответ от OpenAI API');
  }
  
  try {
    const analysisResult = JSON.parse(responseData.choices[0].message.content);
    return analysisResult;
  } catch (parseError) {
    logError_('GPT_PARSE', 'Ошибка парсинга ответа GPT', parseError);
    // Возвращаем пустой результат в случае ошибки парсинга
    return {
      categories: { "Прочее": commentsBatch },
      insights: [],
      recommendations: []
    };
  }
}

/**
 * Создаёт промпт для анализа причин отказов
 * @param {Array} comments - Массив комментариев
 * @returns {string} Промпт для GPT
 * @private
 */
function createGPTPromptForRefusalAnalysis_(comments) {
  const commentsText = comments
    .slice(0, 20) // Ограничиваем количество для избежания превышения лимита токенов
    .map((comment, index) => `${index + 1}. ${comment}`)
    .join('\n');
  
  return `Проанализируй следующие причины отказов клиентов:

${commentsText}

Задачи:
1. Сгруппируй причины по логическим категориям
2. Выяви основные проблемы и закономерности  
3. Дай рекомендации по улучшению работы с клиентами

Отвечай только в формате JSON без дополнительных объяснений.`;
}

/**
 * Анализирует отказы по каналам
 * @param {Array} refusedDeals - Отклонённые сделки
 * @returns {Object} Анализ по каналам
 * @private
 */
function analyzeRefusalsByChannel_(refusedDeals) {
  const channelStats = {};
  
  refusedDeals.forEach(deal => {
    const channel = deal.channel || 'Неизвестно';
    
    if (!channelStats[channel]) {
      channelStats[channel] = {
        count: 0,
        totalBudget: 0,
        averageBudget: 0,
        commonReasons: {}
      };
    }
    
    channelStats[channel].count++;
    channelStats[channel].totalBudget += deal.budget;
    
    // Считаем частые слова в причинах отказов по каналам
    if (deal.refusalComment) {
      const words = deal.refusalComment.toLowerCase()
        .split(/\s+/)
        .filter(word => word.length > 3)
        .filter(word => !['что', 'как', 'это', 'для', 'или', 'еще', 'уже'].includes(word));
      
      words.forEach(word => {
        if (!channelStats[channel].commonReasons[word]) {
          channelStats[channel].commonReasons[word] = 0;
        }
        channelStats[channel].commonReasons[word]++;
      });
    }
  });
  
  // Вычисляем средний чек и сортируем частые причины
  Object.keys(channelStats).forEach(channel => {
    const stats = channelStats[channel];
    stats.averageBudget = stats.count > 0 ? stats.totalBudget / stats.count : 0;
    
    // Оставляем только топ-3 частые причины
    stats.commonReasons = Object.entries(stats.commonReasons)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 3)
      .reduce((obj, [word, count]) => {
        obj[word] = count;
        return obj;
      }, {});
  });
  
  return channelStats;
}

/**
 * Анализирует тренды отказов по месяцам
 * @param {Array} refusedDeals - Отклонённые сделки
 * @returns {Object} Тренды по месяцам
 * @private
 */
function analyzeRefusalTrends_(refusedDeals) {
  const monthlyStats = {};
  
  refusedDeals.forEach(deal => {
    if (!deal.createdDate) return;
    
    const monthKey = formatDate_(deal.createdDate, 'YYYY-MM');
    const monthName = `${getMonthName_(deal.createdDate.getMonth())} ${deal.createdDate.getFullYear()}`;
    
    if (!monthlyStats[monthKey]) {
      monthlyStats[monthKey] = {
        month: monthName,
        count: 0,
        totalBudget: 0,
        channels: {}
      };
    }
    
    monthlyStats[monthKey].count++;
    monthlyStats[monthKey].totalBudget += deal.budget;
    
    const channel = deal.channel || 'Неизвестно';
    monthlyStats[monthKey].channels[channel] = (monthlyStats[monthKey].channels[channel] || 0) + 1;
  });
  
  return Object.entries(monthlyStats)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, stats]) => stats);
}

/**
 * Создаёт структуру отчёта по анализу причин отказов
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} analysisResults - Результаты анализа
 * @private
 */
function createRefusalAnalysisStructure_(sheet, analysisResults) {
  let currentRow = 3;
  
  // 1. Общая статистика
  sheet.getRange('A3:B3').merge();
  sheet.getRange('A3').setValue('🔍 ОБЩАЯ СТАТИСТИКА ОТКАЗОВ');
  applyHeaderStyle_(sheet.getRange('A3:B3'));
  currentRow += 2;
  
  const generalStats = [
    ['Всего отказов:', analysisResults.totalRefusals],
    ['Категорий причин:', Object.keys(analysisResults.categorizedReasons).length],
    ['Каналов с отказами:', Object.keys(analysisResults.channelAnalysis).length]
  ];
  
  sheet.getRange(currentRow, 1, generalStats.length, 2).setValues(generalStats);
  sheet.getRange(currentRow, 1, generalStats.length, 1).setFontWeight('bold');
  currentRow += generalStats.length + 2;
  
  // 2. Категории причин отказов
  sheet.getRange(currentRow, 1, 1, 3).merge();
  sheet.getRange(currentRow, 1).setValue('📊 КАТЕГОРИИ ПРИЧИН ОТКАЗОВ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  const categoryHeaders = [['Категория', 'Количество', 'Процент']];
  sheet.getRange(currentRow, 1, 1, 3).setValues(categoryHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 3));
  currentRow++;
  
  const totalReasons = Object.values(analysisResults.categorizedReasons)
    .reduce((sum, reasons) => sum + reasons.length, 0);
  
  const categoryData = Object.entries(analysisResults.categorizedReasons)
    .sort(([,a], [,b]) => b.length - a.length)
    .map(([category, reasons]) => [
      category,
      reasons.length,
      totalReasons > 0 ? `${(reasons.length / totalReasons * 100).toFixed(1)}%` : '0%'
    ]);
  
  if (categoryData.length > 0) {
    sheet.getRange(currentRow, 1, categoryData.length, 3).setValues(categoryData);
    currentRow += categoryData.length;
  }
  currentRow += 2;
  
  // 3. Топ каналов по отказам
  sheet.getRange(currentRow, 1, 1, 4).merge();
  sheet.getRange(currentRow, 1).setValue('🎯 КАНАЛЫ С НАИБОЛЬШИМ КОЛИЧЕСТВОМ ОТКАЗОВ');
  applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
  currentRow++;
  
  const channelHeaders = [['Канал', 'Отказы', 'Средний чек', 'Основные причины']];
  sheet.getRange(currentRow, 1, 1, 4).setValues(channelHeaders);
  applySubheaderStyle_(sheet.getRange(currentRow, 1, 1, 4));
  currentRow++;
  
  const channelData = Object.entries(analysisResults.channelAnalysis)
    .sort(([,a], [,b]) => b.count - a.count)
    .slice(0, 10)
    .map(([channel, stats]) => [
      channel,
      stats.count,
      formatCurrency_(stats.averageBudget),
      Object.keys(stats.commonReasons).slice(0, 2).join(', ')
    ]);
  
  if (channelData.length > 0) {
    sheet.getRange(currentRow, 1, channelData.length, 4).setValues(channelData);
    currentRow += channelData.length;
  }
  currentRow += 2;
  
  // 4. Ключевые инсайты от GPT
  if (analysisResults.insights.length > 0) {
    sheet.getRange(currentRow, 1, 1, 2).merge();
    sheet.getRange(currentRow, 1).setValue('💡 КЛЮЧЕВЫЕ ИНСАЙТЫ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 2));
    currentRow++;
    
    analysisResults.insights.slice(0, 5).forEach((insight, index) => {
      sheet.getRange(currentRow, 1).setValue(`${index + 1}.`);
      sheet.getRange(currentRow, 2).setValue(insight);
      sheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
    });
    
    currentRow++;
  }
  
  // 5. Рекомендации от GPT
  if (analysisResults.recommendations.length > 0) {
    sheet.getRange(currentRow, 1, 1, 2).merge();
    sheet.getRange(currentRow, 1).setValue('🎯 РЕКОМЕНДАЦИИ');
    applyHeaderStyle_(sheet.getRange(currentRow, 1, 1, 2));
    currentRow++;
    
    analysisResults.recommendations.slice(0, 5).forEach((recommendation, index) => {
      sheet.getRange(currentRow, 1).setValue(`${index + 1}.`);
      sheet.getRange(currentRow, 2).setValue(recommendation);
      sheet.getRange(currentRow, 1).setFontWeight('bold');
      currentRow++;
    });
  }
  
  // Автоматическая ширина колонок
  sheet.autoResizeColumns(1, 4);
  
  // Применяем стили к данным
  const dataRange = sheet.getRange(3, 1, currentRow - 3, 4);
  applyDataStyle_(dataRange);
}

/**
 * Добавляет диаграммы к анализу причин отказов
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} analysisResults - Результаты анализа
 * @private
 */
function addRefusalAnalysisCharts_(sheet, analysisResults) {
  // 1. Круговая диаграмма категорий отказов
  if (Object.keys(analysisResults.categorizedReasons).length > 0) {
    const categoriesChartData = [['Категория', 'Количество']].concat(
      Object.entries(analysisResults.categorizedReasons)
        .sort(([,a], [,b]) => b.length - a.length)
        .map(([category, reasons]) => [category, reasons.length])
    );
    
    const categoriesChart = sheet.insertChart(
      Charts.newPieChart()
        .setDataRange(sheet.getRange(1, 6, categoriesChartData.length, 2))
        .setOption('title', 'Распределение причин отказов по категориям')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'right' })
        .setOption('chartArea', { width: '80%', height: '80%' })
        .setPosition(3, 6, 0, 0)
        .setOption('width', 500)
        .setOption('height', 350)
        .build()
    );
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 6, categoriesChartData.length, 2).setValues(categoriesChartData);
  }
  
  // 2. Столбчатая диаграмма отказов по каналам
  if (Object.keys(analysisResults.channelAnalysis).length > 0) {
    const channelsChartData = [['Канал', 'Количество отказов']].concat(
      Object.entries(analysisResults.channelAnalysis)
        .sort(([,a], [,b]) => b.count - a.count)
        .slice(0, 10)
        .map(([channel, stats]) => [channel, stats.count])
    );
    
    const channelsChart = sheet.insertChart(
      Charts.newColumnChart()
        .setDataRange(sheet.getRange(1, 9, channelsChartData.length, 2))
        .setOption('title', 'Количество отказов по каналам')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Канал', slantedText: true })
        .setOption('vAxis', { title: 'Количество отказов' })
        .setPosition(3, 12, 0, 0)
        .setOption('width', 500)
        .setOption('height', 350)
        .build()
    );
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 9, channelsChartData.length, 2).setValues(channelsChartData);
  }
  
  // 3. Линейная диаграмма трендов по месяцам
  if (analysisResults.monthlyTrends.length > 0) {
    const trendsChartData = [['Месяц', 'Количество отказов']].concat(
      analysisResults.monthlyTrends.map(item => [item.month, item.count])
    );
    
    const trendsChart = sheet.insertChart(
      Charts.newLineChart()
        .setDataRange(sheet.getRange(1, 12, trendsChartData.length, 2))
        .setOption('title', 'Динамика отказов по месяцам')
        .setOption('titleTextStyle', { fontSize: 14, bold: true })
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', { title: 'Месяц', slantedText: true })
        .setOption('vAxis', { title: 'Количество отказов' })
        .setPosition(25, 6, 0, 0)
        .setOption('width', 600)
        .setOption('height', 350)
        .build()
    );
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 12, trendsChartData.length, 2).setValues(trendsChartData);
  }
}

/**
 * Создаёт пустой отчёт когда нет данных
 * @private
 */
function createEmptyRefusalReport_() {
  const sheet = getOrCreateSheet_(CONFIG.SHEETS.REFUSAL_ANALYSIS);
  clearSheetData_(sheet);
  applySheetFormatting_(sheet, 'Анализ причин отказов');
  
  sheet.getRange('A3').setValue('ℹ️ Нет данных об отклонённых сделках для анализа');
  sheet.getRange('A3').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A5').setValue('Убедитесь, что:');
  sheet.getRange('A6').setValue('• В AmoCRM есть сделки со статусом "Отклонена"');
  sheet.getRange('A7').setValue('• У отклонённых сделок заполнены причины отказа в комментариях');
  sheet.getRange('A8').setValue('• Выполнена синхронизация данных из AmoCRM');
  
  updateLastUpdateTime_(CONFIG.SHEETS.REFUSAL_ANALYSIS);
}
