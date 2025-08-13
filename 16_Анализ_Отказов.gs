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
    const analysisResults = analyzeRefusalReasonsWithGPT_(refusedDeals);
    
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
 * Получает данные отклонённых сделок для анализа - ИСПРАВЛЕННАЯ ВЕРСИЯ
 * @returns {Array} Массив отклонённых сделок
 * @private
 */
function getRefusedDealsData_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
  
  if (!workingSheet) {
    throw new Error('Лист "РАБОЧИЙ АМО" не найден');
  }
  
  const rawData = getSheetData_(workingSheet);
  if (rawData.length <= 1) return [];
  
  const headers = rawData[0];
  const rows = rawData.slice(1);
  
  logInfo_('REFUSAL_ANALYSIS', `Анализируем ${rows.length} записей для поиска отказов`);
  
  // Ищем колонку D (индекс 3) - это статус
  const statusColumnIndex = 3; // Колонка D
  
  if (headers.length <= statusColumnIndex) {
    logError_('REFUSAL_ANALYSIS', `Колонка D не найдена. Всего колонок: ${headers.length}`);
    return [];
  }
  
  logInfo_('REFUSAL_ANALYSIS', `Анализируем колонку D: "${headers[statusColumnIndex]}"`);
  
  // Точно фильтруем по статусу "Закрыто и не реализовано"
  const refusedDeals = rows.filter(row => {
    const status = String(row[statusColumnIndex] || '').trim();
    return status === 'Закрыто и не реализовано';
  });
  
  logInfo_('REFUSAL_ANALYSIS', `Найдено ${refusedDeals.length} отказанных сделок со статусом "Закрыто и не реализовано"`);
  
  if (refusedDeals.length === 0) {
    // Проверяем, что вообще есть в колонке D
    const statusStats = {};
    rows.forEach(row => {
      const status = String(row[statusColumnIndex] || 'Не указан');
      statusStats[status] = (statusStats[status] || 0) + 1;
    });
    
    logWarning_('REFUSAL_ANALYSIS', 'НЕ НАЙДЕНО ОТКАЗОВ! Статистика в колонке D:');
    Object.entries(statusStats)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10)
      .forEach(([status, count]) => {
        logInfo_('REFUSAL_ANALYSIS', `"${status}": ${count} записей`);
      });
    
    return [];
  }
  
  // Определяем индексы нужных колонок
  const dealIdIndex = findColumnIndex(headers, ['ID', 'Сделка.ID', 'Идентификатор']);
  const dealNameIndex = findColumnIndex(headers, ['Название', 'Сделка.Название']);
  const channelIndex = findColumnIndex(headers, ['utm_source', 'UTM_SOURCE', 'Сделка.utm_source', 'Источник']);
  const createdDateIndex = findColumnIndex(headers, ['Дата создания', 'DATE', 'Сделка.Дата создания']);
  const budgetIndex = findColumnIndex(headers, ['Бюджет', 'Сделка.Бюджет', 'Сумма', 'Сумма ₽']);
  const managerIndex = findColumnIndex(headers, ['Ответственный', 'Кем создана', 'Менеджер']);
  
  // Ищем поля с комментариями (могут содержать причины отказа)
  const commentIndex = findColumnIndex(headers, [
    'Причина отказа', 
    'Комментарий', 
    'Примечания', 
    'Notes', 
    'Comment',
    'Отказ',
    'Сделка.Комментарий'
  ]);
  
  logInfo_('REFUSAL_ANALYSIS', `Найдены индексы колонок:
    ID: ${dealIdIndex >= 0 ? headers[dealIdIndex] : 'не найден'}
    Название: ${dealNameIndex >= 0 ? headers[dealNameIndex] : 'не найден'}
    Канал: ${channelIndex >= 0 ? headers[channelIndex] : 'не найден'}
    Дата: ${createdDateIndex >= 0 ? headers[createdDateIndex] : 'не найден'}
    Бюджет: ${budgetIndex >= 0 ? headers[budgetIndex] : 'не найден'}
    Менеджер: ${managerIndex >= 0 ? headers[managerIndex] : 'не найден'}
    Комментарий: ${commentIndex >= 0 ? headers[commentIndex] : 'не найден'}`);
  
  // Формируем структурированные данные
  return refusedDeals.map((row, index) => {
    const dealId = dealIdIndex >= 0 ? String(row[dealIdIndex] || '') : `deal_${index}`;
    const dealName = dealNameIndex >= 0 ? String(row[dealNameIndex] || '') : 'Без названия';
    const channel = channelIndex >= 0 ? String(row[channelIndex] || 'Неизвестно') : 'Неизвестно';
    const createdDate = createdDateIndex >= 0 ? parseDate_(row[createdDateIndex]) : new Date();
    const budget = budgetIndex >= 0 ? (parseFloat(row[budgetIndex]) || 0) : 0;
    const manager = managerIndex >= 0 ? String(row[managerIndex] || 'Неназначен') : 'Неназначен';
    const status = 'Закрыто и не реализовано';
    
    // Получаем комментарий (причину отказа)
    let refusalComment = '';
    if (commentIndex >= 0 && row[commentIndex]) {
      refusalComment = String(row[commentIndex]).trim();
    }
    
    // Если отдельного комментария нет, используем статус + любую доступную текстовую информацию
    if (!refusalComment) {
      // Пытаемся найти любые текстовые поля, которые могут содержать причину
      const textFields = [];
      row.forEach((cell, idx) => {
        const cellValue = String(cell || '').trim();
        if (cellValue && 
            cellValue.length > 5 && 
            cellValue !== status &&
            cellValue !== dealName &&
            cellValue !== dealId &&
            !cellValue.match(/^\d+$/) && // не числа
            !cellValue.match(/^\d{4}-\d{2}-\d{2}/) && // не даты
            !cellValue.match(/^[+]?[\d\s\-\(\)]{7,15}$/) // не телефоны
           ) {
          textFields.push(cellValue);
        }
      });
      
      refusalComment = textFields.length > 0 ? 
        textFields.slice(0, 2).join(' | ') : 
        'Причина не указана - статус: ' + status;
    }
    
    // UTM данные
    const utmSourceIndex = findColumnIndex(headers, ['utm_source', 'UTM_SOURCE', 'Сделка.utm_source']);
    const utmCampaignIndex = findColumnIndex(headers, ['utm_campaign', 'UTM_CAMPAIGN', 'Сделка.utm_campaign']);
    
    return {
      dealId: dealId,
      dealName: dealName,
      channel: channel,
      createdDate: createdDate,
      budget: budget,
      manager: manager,
      status: status,
      refusalComment: refusalComment,
      utmSource: utmSourceIndex >= 0 ? String(row[utmSourceIndex] || '') : '',
      utmCampaign: utmCampaignIndex >= 0 ? String(row[utmCampaignIndex] || '') : ''
    };
  }).filter(deal => deal.refusalComment && deal.refusalComment.trim().length > 0);
}

/**
 * Анализирует причины отказов с помощью GPT
 * @param {Array} refusedDeals - Массив отклонённых сделок
 * @returns {Object} Результаты анализа
 * @private
 */
function analyzeRefusalReasonsWithGPT_(refusedDeals) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    logWarning_('REFUSAL_ANALYSIS', 'API ключ OpenAI не настроен, используем локальный анализ');
    return analyzeRefusalReasonsLocally_(refusedDeals);
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
  const batchSize = 10; // Уменьшаем размер батча с 20 до 10
  const batches = [];
  for (let i = 0; i < comments.length; i += batchSize) {
    batches.push(comments.slice(i, i + batchSize));
  }
  
  // Ограничиваем количество батчей для избежания лимитов API
  const maxBatches = Math.min(batches.length, 10); // Максимум 10 батчей
  const processBatches = batches.slice(0, maxBatches);
  
  logInfo_('GPT_ANALYSIS', `Будет обработано ${processBatches.length} батчей из ${batches.length} (ограничение API)`);
  
  let categorizedReasons = {};
  let insights = [];
  let recommendations = [];
  let successfulBatches = 0;
  
  // Анализируем каждый батч
  for (let i = 0; i < processBatches.length; i++) {
    logInfo_('GPT_ANALYSIS', `Анализ батча ${i + 1} из ${processBatches.length}`);
    
    try {
      const batchResults = analyzeRefusalBatch_(apiKey, processBatches[i], refusedDeals);
      
      // Объединяем результаты
      Object.keys(batchResults.categories).forEach(category => {
        if (!categorizedReasons[category]) {
          categorizedReasons[category] = [];
        }
        categorizedReasons[category] = categorizedReasons[category].concat(batchResults.categories[category]);
      });
      
      insights = insights.concat(batchResults.insights);
      recommendations = recommendations.concat(batchResults.recommendations);
      successfulBatches++;
      
      // Увеличиваем паузу между запросами до 3 секунд
      if (i < processBatches.length - 1) {
        logInfo_('GPT_ANALYSIS', 'Пауза 3 секунды перед следующим запросом...');
        Utilities.sleep(3000);
      }
      
    } catch (error) {
      logError_('GPT_ANALYSIS', `Ошибка анализа батча ${i + 1}`, error);
      
      // Если получили 429 ошибку, делаем длинную паузу
      if (error.toString().includes('429') || error.toString().includes('Rate limit')) {
        logWarning_('GPT_ANALYSIS', 'Достигнут лимит API, пауза 10 секунд');
        Utilities.sleep(10000);
      }
      continue;
    }
  }
  
  // Если ни один батч не удалось обработать, используем локальный анализ
  if (successfulBatches === 0) {
    logWarning_('GPT_ANALYSIS', 'Не удалось обработать ни одного батча через GPT, используем локальный анализ');
    return analyzeRefusalReasonsLocally_(refusedDeals);
  }
  
  logInfo_('GPT_ANALYSIS', `Успешно обработано ${successfulBatches} из ${processBatches.length} батчей`);
  
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
 * Локальный анализ причин отказов без GPT
 * @param {Array} refusedDeals - Массив отклонённых сделок
 * @returns {Object} Результаты анализа
 * @private
 */
function analyzeRefusalReasonsLocally_(refusedDeals) {
  logInfo_('LOCAL_ANALYSIS', 'Выполняем локальный анализ причин отказов');
  
  const categorizedReasons = {
    'Цена': [],
    'Качество': [],
    'Сроки': [],
    'Коммуникация': [],
    'Конкуренты': [],
    'Личные причины': [],
    'Технические': [],
    'Недоверие': [],
    'Прочее': []
  };
  
  // Ключевые слова для категоризации
  const keywords = {
    'Цена': ['дорого', 'цена', 'стоимость', 'дешевле', 'бюджет', 'деньги', 'рубл'],
    'Качество': ['качество', 'плохо', 'некачественно', 'брак', 'дефект'],
    'Сроки': ['срок', 'время', 'долго', 'быстро', 'поздно', 'график'],
    'Коммуникация': ['общение', 'связь', 'менеджер', 'консультант', 'звонок', 'отвеч'],
    'Конкуренты': ['конкурент', 'другой', 'альтернатива', 'выбрал'],
    'Личные причины': ['личн', 'семь', 'переехал', 'болезнь', 'изменил'],
    'Технические': ['техническ', 'сайт', 'ошибк', 'проблем'],
    'Недоверие': ['не доверяю', 'сомнени', 'подозрительно', 'мошенник']
  };
  
  refusedDeals.forEach(deal => {
    const comment = deal.refusalComment.toLowerCase();
    let categorized = false;
    
    // Проверяем каждую категорию
    Object.entries(keywords).forEach(([category, words]) => {
      if (words.some(word => comment.includes(word))) {
        categorizedReasons[category].push(deal.refusalComment);
        categorized = true;
      }
    });
    
    // Если не удалось категоризировать, помещаем в "Прочее"
    if (!categorized) {
      categorizedReasons['Прочее'].push(deal.refusalComment);
    }
  });
  
  const insights = [
    'Локальный анализ выполнен без использования GPT',
    'Категоризация основана на ключевых словах'
  ];
  
  const recommendations = [
    'Рекомендуется настроить API ключ OpenAI для более точного анализа',
    'Проверьте лимиты API OpenAI и попробуйте позже'
  ];
  
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
function analyzeRefusalBatch_(apiKey, commentsBatch, dealsData) {
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
    let responseContent = responseData.choices[0].message.content;
    
    // Убираем markdown форматирование если присутствует
    responseContent = responseContent.replace(/```json\s*/g, '').replace(/```\s*/g, '');
    responseContent = responseContent.trim();
    
    const analysisResult = JSON.parse(responseContent);
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
    .slice(0, 10) // Ограичиваем до 10 комментариев для экономии токенов
    .map((comment, index) => `${index + 1}. ${comment}`)
    .join('\n');
  
  return `Проанализируй следующие причины отказов клиентов и верни ТОЛЬКО валидный JSON:

${commentsText}

Верни результат строго в этом JSON формате без дополнительного текста:
{
  "categories": {
    "Цена": [],
    "Качество": [],
    "Сроки": [],
    "Коммуникация": [],
    "Конкуренты": [],
    "Личные причины": [],
    "Технические": [],
    "Недоверие": [],
    "Прочее": []
  },
  "insights": [],
  "recommendations": []
}`;
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
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 6, categoriesChartData.length, 2).setValues(categoriesChartData);
    
    try {
      const categoriesChart = createChart_(sheet, 'pie', categoriesChartData, {
        title: 'Распределение причин отказов по категориям',
        position: { row: 3, col: 6 },
        width: 500,
        height: 350
      });
    } catch (chartError) {
      logWarning_('CHARTS', 'Ошибка создания диаграммы категорий', chartError);
    }
  }
  
  // 2. Столбчатая диаграмма отказов по каналам
  if (Object.keys(analysisResults.channelAnalysis).length > 0) {
    const channelsChartData = [['Канал', 'Количество отказов']].concat(
      Object.entries(analysisResults.channelAnalysis)
        .sort(([,a], [,b]) => b.count - a.count)
        .slice(0, 10)
        .map(([channel, stats]) => [channel, stats.count])
    );
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 9, channelsChartData.length, 2).setValues(channelsChartData);
    
    try {
      const channelsChart = createChart_(sheet, 'column', channelsChartData, {
        title: 'Количество отказов по каналам',
        position: { row: 3, col: 12 },
        width: 500,
        height: 350
      });
    } catch (chartError) {
      logWarning_('CHARTS', 'Ошибка создания диаграммы каналов', chartError);
    }
  }
  
  // 3. Линейная диаграмма трендов по месяцам
  if (analysisResults.monthlyTrends.length > 0) {
    const trendsChartData = [['Месяц', 'Количество отказов']].concat(
      analysisResults.monthlyTrends.map(item => [item.month, item.count])
    );
    
    // Записываем данные для диаграммы
    sheet.getRange(1, 12, trendsChartData.length, 2).setValues(trendsChartData);
    
    try {
      const trendsChart = createChart_(sheet, 'line', trendsChartData, {
        title: 'Динамика отказов по месяцам',
        position: { row: 25, col: 6 },
        width: 600,
        height: 350
      });
    } catch (chartError) {
      logWarning_('CHARTS', 'Ошибка создания диаграммы трендов', chartError);
    }
  }
}

/**
 * Создаёт пустой отчёт когда нет данных
 * @private
 */
function createEmptyRefusalReport_() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet_('Анализ отказов');
    sheet.clear();
    
    // Заголовок
    sheet.getRange('A1').setValue('📊 АНАЛИЗ ПРИЧИН ОТКАЗОВ');
    applyHeaderStyle_(sheet.getRange('A1'));
    sheet.getRange('A1').merge(sheet.getRange('A1:D1'));
    
    // Время обновления
    sheet.getRange('A2').setValue(`⏰ Последнее обновление: ${getCurrentDateMoscow_().toLocaleString()}`);
    
    // Основное сообщение
    sheet.getRange('A3').setValue('ℹ️ Нет данных об отклонённых сделках для анализа');
    applySubheaderStyle_(sheet.getRange('A3'));
    
    sheet.getRange('A5').setValue('🔍 Убедитесь, что:');
    sheet.getRange('A6').setValue('• В AmoCRM есть сделки со статусом "Закрыто и не реализовано"');
    sheet.getRange('A7').setValue('• У отклонённых сделок заполнены причины отказа в комментариях');
    sheet.getRange('A8').setValue('• Выполнена синхронизация данных из AmoCRM');
    
    // Список поддерживаемых статусов отказов
    sheet.getRange('A10').setValue('📋 Распознаваемые статусы отказов:');
    sheet.getRange('A11').setValue('• "Закрыто и не реализовано"');
    sheet.getRange('A12').setValue('• "Отклонена", "Отказ"');
    sheet.getRange('A13').setValue('• "Не реализовано", "Отклонено"');
    sheet.getRange('A14').setValue('• "Неуспешно реализовано"');
    
    logInfo_('REFUSAL_ANALYSIS', 'Создан пустой отчёт анализа отказов');
    
  } catch (error) {
    logError_('REFUSAL_ANALYSIS', 'Ошибка создания пустого отчёта', error);
  }
}

/**
 * Быстрая диагностика данных для анализа отказов
 * Запустите эту функцию для проверки данных
 */
function diagnoseRefusalData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
  
  if (!workingSheet) {
    console.log('❌ Лист "РАБОЧИЙ АМО" не найден');
    return;
  }
  
  const data = workingSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  console.log('📊 ДИАГНОСТИКА ДАННЫХ ДЛЯ АНАЛИЗА ОТКАЗОВ:');
  console.log(`📋 Всего строк: ${rows.length}`);
  console.log(`📋 Всего колонок: ${headers.length}`);
  console.log(`📋 Колонка D: "${headers[3]}" (индекс 3)`);
  
  // Статистика по колонке D (статусы)
  const statusStats = {};
  rows.forEach(row => {
    const status = String(row[3] || 'Пустая').trim();
    statusStats[status] = (statusStats[status] || 0) + 1;
  });
  
  console.log('\n📈 СТАТИСТИКА СТАТУСОВ (Колонка D):');
  Object.entries(statusStats)
    .sort(([,a], [,b]) => b - a)
    .forEach(([status, count]) => {
      const percentage = ((count / rows.length) * 100).toFixed(1);
      console.log(`• "${status}": ${count} (${percentage}%)`);
    });
  
  const refusedCount = statusStats['Закрыто и не реализовано'] || 0;
  console.log(`\n🎯 ОТКАЗАННЫЕ СДЕЛКИ: ${refusedCount}`);
  
  if (refusedCount > 0) {
    console.log('✅ Данные для анализа найдены!');
    
    // Проверяем наличие полей с комментариями
    const commentFields = headers.filter((h, idx) => {
      const normalized = String(h).toLowerCase();
      return normalized.includes('комментарий') || 
             normalized.includes('причина') ||
             normalized.includes('примечания') ||
             normalized.includes('отказ');
    });
    
    if (commentFields.length > 0) {
      console.log(`💬 Поля с комментариями: ${commentFields.join(', ')}`);
    } else {
      console.log('⚠️ Отдельные поля с комментариями не найдены, будем использовать статус и другие текстовые данные');
    }
  } else {
    console.log('❌ Отказанные сделки не найдены!');
    console.log('🔍 Проверьте, что в колонке D есть статус "Закрыто и не реализовано"');
  }
}
