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
    
    // Применяем КРАСИВОЕ оформление как в старых отчетах
    applyRefusalAnalysisBeautifulStyle_(sheet);
    
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
 * Получает данные отклонённых сделок для анализа - ОБНОВЛЕННАЯ ВЕРСИЯ
 * Работает с новой структурой сводного листа РАБОЧИЙ АМО
 * @returns {Array} Массив отклонённых сделок
 * @private
 */
function getRefusedDealsData_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
  
  if (!workingSheet) {
    logWarning_('REFUSAL_ANALYSIS', 'Лист "РАБОЧИЙ АМО" не найден, создаем структуру...');
    // Создаем структуру если не существует
    createWorkingAmoStructure();
    workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
  }
  
  const rawData = getSheetData_(workingSheet);
  if (rawData.length <= 3) { // Учитываем заголовок + 2 строки оформления
    logWarning_('REFUSAL_ANALYSIS', 'Нет данных в листе "РАБОЧИЙ АМО"');
    return [];
  }
  
  const headers = rawData[2]; // Заголовки в 3-й строке (индекс 2)
  const rows = rawData.slice(3); // Данные начинаются с 4-й строки
  
  logInfo_('REFUSAL_ANALYSIS', `Анализируем ${rows.length} записей для поиска отказов`);
  
  // Ищем колонку "Статус" в новой структуре
  const statusColumnIndex = findColumnIndex(headers, ['Статус']);
  
  if (statusColumnIndex < 0) {
    logError_('REFUSAL_ANALYSIS', 'Колонка "Статус" не найдена в новой структуре');
    return [];
  }
  
  logInfo_('REFUSAL_ANALYSIS', `Анализируем колонку "${headers[statusColumnIndex]}" (столбец ${String.fromCharCode(65 + statusColumnIndex)})`);
  
  // Точно фильтруем по статусу "Закрыто и не реализовано"
  const refusedDeals = rows.filter(row => {
    const status = String(row[statusColumnIndex] || '').trim();
    return status === 'Закрыто и не реализовано';
  });
  
  logInfo_('REFUSAL_ANALYSIS', `Найдено ${refusedDeals.length} отказанных сделок со статусом "Закрыто и не реализовано"`);
  
  if (refusedDeals.length === 0) {
    // Проверяем, что вообще есть в колонке "Статус"
    const statusStats = {};
    rows.forEach(row => {
      const status = String(row[statusColumnIndex] || 'Не указан');
      statusStats[status] = (statusStats[status] || 0) + 1;
    });
    
    logWarning_('REFUSAL_ANALYSIS', 'НЕ НАЙДЕНО ОТКАЗОВ! Статистика в колонке "Статус":');
    Object.entries(statusStats)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10)
      .forEach(([status, count]) => {
        logInfo_('REFUSAL_ANALYSIS', `"${status}": ${count} записей`);
      });
    
    return [];
  }
  
  // Определяем индексы нужных колонок из новой структуры
  const dealIdIndex = findColumnIndex(headers, ['ID']);
  const dealNameIndex = findColumnIndex(headers, ['Название']);  
  const responsibleIndex = findColumnIndex(headers, ['Ответственный']);
  const channelIndex = findColumnIndex(headers, ['Источник', 'UTM_SOURCE']);
  const createdDateIndex = findColumnIndex(headers, ['Дата создания']);
  const budgetIndex = findColumnIndex(headers, ['Бюджет']);
  const managerIndex = findColumnIndex(headers, ['Ответственный']);
  
  // ПРИОРИТЕТНО ищем точную колонку "Причина отказа" из новой структуры
  let commentIndex = findColumnIndex(headers, [
    'Причина отказа', // Основная колонка
    'Комментарий МОБ', // Альтернативный источник
    'Примечания',      // Дополнительные записи
    'Комментарий'      // Общий комментарий
  ]);
  
  if (commentIndex >= 0) {
    logInfo_('REFUSAL_ANALYSIS', `✅ Найдена колонка с причинами отказов: "${headers[commentIndex]}" (столбец ${String.fromCharCode(65 + commentIndex)})`);
  } else {
    logWarning_('REFUSAL_ANALYSIS', '⚠️ Колонка "Причина отказа" не найдена в новой структуре!');
  }
  
  logInfo_('REFUSAL_ANALYSIS', `Найдены индексы колонок в новой структуре:
    ID: ${dealIdIndex >= 0 ? `"${headers[dealIdIndex]}" (столбец ${String.fromCharCode(65 + dealIdIndex)})` : 'не найден'}
    Название: ${dealNameIndex >= 0 ? `"${headers[dealNameIndex]}" (столбец ${String.fromCharCode(65 + dealNameIndex)})` : 'не найден'}
    Канал: ${channelIndex >= 0 ? `"${headers[channelIndex]}" (столбец ${String.fromCharCode(65 + channelIndex)})` : 'не найден'}
    Дата создания: ${createdDateIndex >= 0 ? `"${headers[createdDateIndex]}" (столбец ${String.fromCharCode(65 + createdDateIndex)})` : 'не найден'}
    Бюджет: ${budgetIndex >= 0 ? `"${headers[budgetIndex]}" (столбец ${String.fromCharCode(65 + budgetIndex)})` : 'не найден'}
    Ответственный: ${managerIndex >= 0 ? `"${headers[managerIndex]}" (столбец ${String.fromCharCode(65 + managerIndex)})` : 'не найден'}
    Причины отказов: ${commentIndex >= 0 ? `"${headers[commentIndex]}" (столбец ${String.fromCharCode(65 + commentIndex)})` : 'не найден'}`);
  
  // Формируем структурированные данные из новой структуры
  return refusedDeals.map((row, index) => {
    const dealId = dealIdIndex >= 0 ? String(row[dealIdIndex] || '') : `deal_${index}`;
    const dealName = dealNameIndex >= 0 ? String(row[dealNameIndex] || '') : 'Без названия';
    const channel = channelIndex >= 0 ? String(row[channelIndex] || 'Неизвестно') : 'Неизвестно';
    const createdDate = createdDateIndex >= 0 ? parseDate_(row[createdDateIndex]) : new Date();
    const budget = budgetIndex >= 0 ? (parseFloat(row[budgetIndex]) || 0) : 0;
    const manager = managerIndex >= 0 ? String(row[managerIndex] || 'Неназначен') : 'Неназначен';
    const status = 'Закрыто и не реализовано';
    
    // Получаем причину отказа из новой структуры
    let refusalComment = '';
    if (commentIndex >= 0 && row[commentIndex]) {
      refusalComment = String(row[commentIndex]).trim();
    }
    
    // Если основная колонка пуста, ищем в альтернативных источниках
    if (!refusalComment || refusalComment === '') {
      const alternativeSources = [
        findColumnIndex(headers, ['Комментарий МОБ']),
        findColumnIndex(headers, ['Примечания']),
        findColumnIndex(headers, ['Контакт.ФИО']) // Иногда причина может быть в имени
      ];
      
      for (const altIndex of alternativeSources) {
        if (altIndex >= 0 && row[altIndex]) {
          const altValue = String(row[altIndex]).trim();
          if (altValue && altValue.length > 5 && 
              (altValue.toLowerCase().includes('отказ') || 
               altValue.toLowerCase().includes('причина') ||
               altValue.length > 20)) {
            refusalComment = altValue;
            break;
          }
        }
      }
      
      if (!refusalComment) {
        refusalComment = 'Причина отказа не указана';
      }
    }
    
    // UTM данные из новой структуры
    const utmSourceIndex = findColumnIndex(headers, ['UTM_SOURCE']);
    const utmCampaignIndex = findColumnIndex(headers, ['UTM_CAMPAIGN']);
    const utmMediumIndex = findColumnIndex(headers, ['UTM_MEDIUM']);
    const utmTermIndex = findColumnIndex(headers, ['UTM_TERM']);
    const utmContentIndex = findColumnIndex(headers, ['UTM_CONTENT']);
    
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
      utmCampaign: utmCampaignIndex >= 0 ? String(row[utmCampaignIndex] || '') : '',
      utmMedium: utmMediumIndex >= 0 ? String(row[utmMediumIndex] || '') : '',
      utmTerm: utmTermIndex >= 0 ? String(row[utmTermIndex] || '') : '',
      utmContent: utmContentIndex >= 0 ? String(row[utmContentIndex] || '') : ''
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
  const batchSize = 5; // Еще больше уменьшаем размер батча до 5
  const batches = [];
  for (let i = 0; i < comments.length; i += batchSize) {
    batches.push(comments.slice(i, i + batchSize));
  }
  
  // Ограничиваем количество батчей для избежания лимитов API (3 запроса в минуту)
  const maxBatches = Math.min(batches.length, 6); // Максимум 6 батчей за 2 минуты
  const processBatches = batches.slice(0, maxBatches);
  
  logInfo_('GPT_ANALYSIS', `Будет обработано ${processBatches.length} батчей из ${batches.length} (лимит 3 RPM)`);
  
  let categorizedReasons = {};
  let insights = [];
  let recommendations = [];
  let successfulBatches = 0;
  let requestCount = 0;
  let batchStartTime = new Date();
  
  // Анализируем каждый батч с умным rate limiting
  for (let i = 0; i < processBatches.length; i++) {
    logInfo_('GPT_ANALYSIS', `Анализ батча ${i + 1} из ${processBatches.length}`);
    
    try {
      // Умная пауза для соблюдения лимита 3 RPM
      if (requestCount >= 3) {
        const elapsedMinutes = (new Date() - batchStartTime) / 60000;
        if (elapsedMinutes < 1) {
          const waitTime = Math.ceil((1 - elapsedMinutes) * 60);
          logInfo_('GPT_ANALYSIS', `Пауза ${waitTime} секунд для соблюдения лимита 3 RPM`);
          Utilities.sleep(waitTime * 1000);
        }
        requestCount = 0;
        batchStartTime = new Date();
      }
      
      const batchResults = analyzeRefusalBatch_(apiKey, processBatches[i], refusedDeals);
      requestCount++;
      
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
      
      // Базовая пауза между запросами
      if (i < processBatches.length - 1) {
        logInfo_('GPT_ANALYSIS', 'Пауза 5 секунд между запросами...');
        Utilities.sleep(5000);
      }
      
    } catch (error) {
      logError_('GPT_ANALYSIS', `Ошибка анализа батча ${i + 1}`, error);
      requestCount++;
      
      // Если получили 429 ошибку, делаем очень длинную паузу
      if (error.toString().includes('429') || error.toString().includes('Rate limit')) {
        logWarning_('GPT_ANALYSIS', 'Достигнут лимит API, пауза 30 секунд');
        Utilities.sleep(30000);
        requestCount = 0; // Сбрасываем счетчик после длинной паузы
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
    .slice(0, 5) // Уменьшаем до 5 комментариев для экономии токенов и лучшего качества
    .map((comment, index) => `${index + 1}. ${comment}`)
    .join('\n');
  
  return `Проанализируй эти ${comments.length} причин отказов клиентов и категоризируй их:

${commentsText}

Верни результат строго в JSON формате:
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
  "insights": ["краткий инсайт 1", "краткий инсайт 2"],
  "recommendations": ["рекомендация 1", "рекомендация 2"]
}

Помести каждую причину отказа в наиболее подходящую категорию.`;
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
 * Обновленная версия для красивого оформления
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} analysisResults - Результаты анализа
 * @private
 */
function createRefusalAnalysisStructure_(sheet, analysisResults) {
  try {
    // 1. ОБЩАЯ СТАТИСТИКА (начинаем с строки 5)
    let currentRow = 5;
    
    const generalStats = [
      ['Всего отказов:', analysisResults.totalRefusals],
      ['Категорий причин:', Object.keys(analysisResults.categorizedReasons).length],
      ['Каналов с отказами:', Object.keys(analysisResults.channelAnalysis).length]
    ];
    
    generalStats.forEach(([label, value], index) => {
      sheet.getRange(currentRow + index, 1).setValue(label)
        .setFontWeight('bold')
        .setBackground('#f8f9fa');
      sheet.getRange(currentRow + index, 2).setValue(value)
        .setHorizontalAlignment('center')
        .setBackground('#ffffff');
    });
    
    // 2. КАТЕГОРИИ ПРИЧИН ОТКАЗОВ (начинаем с строки 12)
    currentRow = 12;
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
      categoryData.forEach(([category, count, percent], index) => {
        const rowNum = currentRow + index;
        
        // Чередование цветов строк
        const bgColor = index % 2 === 0 ? '#f8f9fa' : '#ffffff';
        
        sheet.getRange(rowNum, 1).setValue(category)
          .setBackground(bgColor);
        sheet.getRange(rowNum, 2).setValue(count)
          .setBackground(bgColor)
          .setHorizontalAlignment('center');
        sheet.getRange(rowNum, 3).setValue(percent)
          .setBackground(bgColor)
          .setHorizontalAlignment('center');
      });
    }
    
    // 3. КАНАЛЫ С ОТКАЗАМИ (начинаем с строки 24)
    currentRow = 24;
    const channelData = Object.entries(analysisResults.channelAnalysis)
      .sort(([,a], [,b]) => b.count - a.count)
      .slice(0, 8);
    
    if (channelData.length > 0) {
      channelData.forEach(([channel, stats], index) => {
        const rowNum = currentRow + index;
        const bgColor = index % 2 === 0 ? '#f8f9fa' : '#ffffff';
        
        sheet.getRange(rowNum, 1).setValue(channel)
          .setBackground(bgColor);
        sheet.getRange(rowNum, 2).setValue(stats.count)
          .setBackground(bgColor)
          .setHorizontalAlignment('center');
        sheet.getRange(rowNum, 3).setValue(formatCurrency_(stats.averageBudget))
          .setBackground(bgColor)
          .setHorizontalAlignment('right');
        sheet.getRange(rowNum, 4).setValue(Object.keys(stats.commonReasons).slice(0, 2).join(', '))
          .setBackground(bgColor);
      });
    }
    
    // 4. КЛЮЧЕВЫЕ ИНСАЙТЫ (начинаем с строки 35)
    if (analysisResults.insights.length > 0) {
      currentRow = 35;
      analysisResults.insights.slice(0, 8).forEach((insight, index) => {
        const rowNum = currentRow + index;
        const bgColor = index % 2 === 0 ? '#fff3e0' : '#ffffff';
        
        sheet.getRange(rowNum, 1).setValue(`${index + 1}.`)
          .setFontWeight('bold')
          .setBackground(bgColor);
        sheet.getRange(rowNum, 2, 1, 4).merge()
          .setValue(insight)
          .setBackground(bgColor)
          .setWrap(true);
      });
    }
    
    // 5. РЕКОМЕНДАЦИИ (начинаем с строки 47)
    if (analysisResults.recommendations.length > 0) {
      currentRow = 47;
      analysisResults.recommendations.slice(0, 8).forEach((recommendation, index) => {
        const rowNum = currentRow + index;
        const bgColor = index % 2 === 0 ? '#fce4ec' : '#ffffff';
        
        sheet.getRange(rowNum, 1).setValue(`${index + 1}.`)
          .setFontWeight('bold')
          .setBackground(bgColor);
        sheet.getRange(rowNum, 2, 1, 4).merge()
          .setValue(recommendation)
          .setBackground(bgColor)
          .setWrap(true);
      });
    }
    
    logInfo_('REFUSAL_STRUCTURE', 'Структура красивого отчета создана');
    
  } catch (error) {
    logError_('REFUSAL_STRUCTURE', 'Ошибка создания структуры отчета', error);
  }
}

/**
 * Добавляет красивые диаграммы к анализу причин отказов
 * Обновленная версия для нового дизайна
 * @param {Sheet} sheet - Лист таблицы
 * @param {Object} analysisResults - Результаты анализа
 * @private
 */
function addRefusalAnalysisCharts_(sheet, analysisResults) {
  try {
    // 1. Круговая диаграмма категорий отказов (правая часть, строки 4-18)
    if (Object.keys(analysisResults.categorizedReasons).length > 0) {
      const categoriesChartData = [['Категория', 'Количество']].concat(
        Object.entries(analysisResults.categorizedReasons)
          .sort(([,a], [,b]) => b.length - a.length)
          .slice(0, 8) // Топ-8 категорий для читаемости
          .map(([category, reasons]) => [category, reasons.length])
      );
      
      // Записываем данные для диаграммы в столбцы H-I
      sheet.getRange(4, 8, categoriesChartData.length, 2).setValues(categoriesChartData);
      
      try {
        const categoriesChart = createChart_(sheet, 'pie', categoriesChartData, {
          title: 'ТОП-8 причин отказов',
          position: { row: 4, col: 10 },
          width: 450,
          height: 350
        });
        logInfo_('CHARTS', 'Диаграмма категорий создана');
      } catch (chartError) {
        logWarning_('CHARTS', 'Ошибка создания диаграммы категорий', chartError);
      }
    }
    
    // 2. Столбчатая диаграмма отказов по каналам (правая часть, строки 22-36)
    if (Object.keys(analysisResults.channelAnalysis).length > 0) {
      const channelsChartData = [['Канал', 'Количество отказов']].concat(
        Object.entries(analysisResults.channelAnalysis)
          .sort(([,a], [,b]) => b.count - a.count)
          .slice(0, 8) // Топ-8 каналов
          .map(([channel, stats]) => [channel.length > 15 ? channel.substring(0, 15) + '...' : channel, stats.count])
      );
      
      // Записываем данные для диаграммы в столбцы H-I
      sheet.getRange(22, 8, channelsChartData.length, 2).setValues(channelsChartData);
      
      try {
        const channelsChart = createChart_(sheet, 'column', channelsChartData, {
          title: 'Отказы по каналам (ТОП-8)',
          position: { row: 22, col: 10 },
          width: 450,
          height: 300
        });
        logInfo_('CHARTS', 'Диаграмма каналов создана');
      } catch (chartError) {
        logWarning_('CHARTS', 'Ошибка создания диаграммы каналов', chartError);
      }
    }
    
    // 3. Тренды по месяцам (нижняя часть листа, строки 60+)
    if (analysisResults.monthlyTrends && analysisResults.monthlyTrends.length > 1) {
      const trendsChartData = [['Месяц', 'Количество отказов']].concat(
        analysisResults.monthlyTrends
          .slice(-6) // Последние 6 месяцев
          .map(item => [item.month.length > 10 ? item.month.substring(0, 10) : item.month, item.count])
      );
      
      // Записываем данные для диаграммы в столбцы A-B (строки 60+)
      sheet.getRange(60, 1, trendsChartData.length, 2).setValues(trendsChartData);
      
      try {
        const trendsChart = createChart_(sheet, 'line', trendsChartData, {
          title: 'Динамика отказов (последние 6 месяцев)',
          position: { row: 60, col: 3 },
          width: 600,
          height: 300
        });
        logInfo_('CHARTS', 'Диаграмма трендов создана');
      } catch (chartError) {
        logWarning_('CHARTS', 'Ошибка создания диаграммы трендов', chartError);
      }
    }
    
    logInfo_('CHARTS', 'Все диаграммы для анализа отказов созданы');
    
  } catch (error) {
    logError_('CHARTS', 'Ошибка создания диаграмм анализа отказов', error);
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
 * Обновленная версия для новой структуры сводного листа РАБОЧИЙ АМО
 */
function diagnoseRefusalData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
  
  if (!workingSheet) {
    console.log('❌ Лист "РАБОЧИЙ АМО" не найден');
    console.log('🔧 Используйте createWorkingAmoStructure() для создания структуры');
    return;
  }
  
  const data = workingSheet.getDataRange().getValues();
  
  // Проверяем структуру листа (заголовки должны быть в 3-й строке)
  if (data.length < 3) {
    console.log('❌ Неправильная структура листа "РАБОЧИЙ АМО"');
    console.log('🔧 Лист должен содержать: 1) титул, 2) дату обновления, 3) заголовки');
    return;
  }
  
  const headers = data[2]; // Заголовки в 3-й строке (индекс 2)
  const rows = data.slice(3); // Данные начинаются с 4-й строки
  
  console.log('📊 ДИАГНОСТИКА НОВОЙ СТРУКТУРЫ РАБОЧИЙ АМО:');
  console.log(`📋 Всего строк данных: ${rows.length}`);
  console.log(`📋 Всего столбцов: ${headers.length}`);
  
  // Проверяем ключевые поля новой структуры
  console.log('\n🔍 ПРОВЕРКА КЛЮЧЕВЫХ ПОЛЕЙ:');
  const keyFields = [
    'ID', 'Название', 'Статус', 'Причина отказа', 
    'Контакт.ФИО', 'Контакт.Телефон', 'Источник',
    'UTM_SOURCE', 'Бюджет', 'Дата создания'
  ];
  
  keyFields.forEach(field => {
    const fieldIndex = findColumnIndex(headers, [field]);
    if (fieldIndex >= 0) {
      console.log(`✅ "${field}": столбец ${String.fromCharCode(65 + fieldIndex)} ("${headers[fieldIndex]}")`);
    } else {
      console.log(`❌ "${field}": не найдено`);
    }
  });
  
  // Ищем колонку "Статус" для анализа отказов
  const statusIndex = findColumnIndex(headers, ['Статус']);
  if (statusIndex >= 0) {
    console.log(`\n📋 Колонка "Статус": "${headers[statusIndex]}" (столбец ${String.fromCharCode(65 + statusIndex)})`);
    
    // Статистика по статусам
    const statusStats = {};
    rows.forEach(row => {
      const status = String(row[statusIndex] || 'Не указан').trim();
      statusStats[status] = (statusStats[status] || 0) + 1;
    });
    
    console.log('\n📈 СТАТИСТИКА СТАТУСОВ:');
    Object.entries(statusStats)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10)
      .forEach(([status, count]) => {
        const percentage = ((count / rows.length) * 100).toFixed(1);
        console.log(`• "${status}": ${count} (${percentage}%)`);
      });
    
    const refusedCount = statusStats['Закрыто и не реализовано'] || 0;
    console.log(`\n🎯 ОТКАЗАННЫЕ СДЕЛКИ: ${refusedCount}`);
    
    // Анализ причин отказов
    const reasonIndex = findColumnIndex(headers, ['Причина отказа']);
    if (refusedCount > 0 && reasonIndex >= 0) {
      console.log(`\n💬 АНАЛИЗ ПРИЧИН ОТКАЗОВ:`);
      console.log(`📋 Колонка "Причина отказа": "${headers[reasonIndex]}" (столбец ${String.fromCharCode(65 + reasonIndex)})`);
      
      let refusedWithReasons = 0;
      let refusedWithoutReasons = 0;
      
      rows.forEach(row => {
        const status = String(row[statusIndex] || '').trim();
        if (status === 'Закрыто и не реализовано') {
          const reason = String(row[reasonIndex] || '').trim();
          if (reason && reason !== '' && reason.length > 3) {
            refusedWithReasons++;
          } else {
            refusedWithoutReasons++;
          }
        }
      });
      
      console.log(`• С причинами: ${refusedWithReasons} (${((refusedWithReasons / refusedCount) * 100).toFixed(1)}%)`);
      console.log(`• Без причин: ${refusedWithoutReasons} (${((refusedWithoutReasons / refusedCount) * 100).toFixed(1)}%)`);
      
      if (refusedWithReasons > 0) {
        console.log(`\n🚀 ГОТОВО К GPT-АНАЛИЗУ: ${refusedWithReasons} сделок с причинами отказов`);
      }
      
    } else if (reasonIndex < 0) {
      console.log(`\n❌ Колонка "Причина отказа" не найдена!`);
      
      // Ищем альтернативные источники
      const altSources = ['Комментарий МОБ', 'Примечания', 'Комментарий'];
      console.log('🔍 Поиск альтернативных источников причин отказов:');
      altSources.forEach(sourceName => {
        const altIndex = findColumnIndex(headers, [sourceName]);
        if (altIndex >= 0) {
          console.log(`✅ Найдена "${sourceName}": столбец ${String.fromCharCode(65 + altIndex)}`);
        } else {
          console.log(`❌ "${sourceName}": не найдена`);
        }
      });
    }
  } else {
    console.log('\n❌ Колонка "Статус" не найдена в новой структуре!');
  }
  
  // Проверка UTM полей
  console.log('\n🏷️ ПРОВЕРКА UTM ПОЛЕЙ:');
  const utmFields = ['UTM_SOURCE', 'UTM_MEDIUM', 'UTM_CAMPAIGN', 'UTM_TERM', 'UTM_CONTENT'];
  utmFields.forEach(utmField => {
    const utmIndex = findColumnIndex(headers, [utmField]);
    if (utmIndex >= 0) {
      // Считаем заполненность
      let filled = 0;
      rows.forEach(row => {
        if (row[utmIndex] && String(row[utmIndex]).trim() !== '') {
          filled++;
        }
      });
      const percentage = rows.length > 0 ? ((filled / rows.length) * 100).toFixed(1) : '0';
      console.log(`✅ ${utmField}: ${filled}/${rows.length} (${percentage}%)`);
    } else {
      console.log(`❌ ${utmField}: не найдено`);
    }
  });
  
  // Общая статистика заполненности
  console.log('\n📊 ОБЩАЯ СТАТИСТИКА ЗАПОЛНЕННОСТИ:');
  let totalFields = 0;
  let filledFields = 0;
  
  rows.forEach(row => {
    row.forEach(cell => {
      totalFields++;
      if (cell && String(cell).trim() !== '') {
        filledFields++;
      }
    });
  });
  
  const overallFillRate = totalFields > 0 ? ((filledFields / totalFields) * 100).toFixed(1) : '0';
  console.log(`📈 Общая заполненность данных: ${overallFillRate}%`);
  console.log(`📋 Заполнено ячеек: ${filledFields} из ${totalFields}`);
  
  console.log('\n✅ Диагностика новой структуры РАБОЧИЙ АМО завершена');
}

/**
 * Применяет красивое оформление к листу анализа отказов
 * Создает стиль как в старых красивых отчетах
 * @param {Sheet} sheet - Лист для оформления
 * @private
 */
function applyRefusalAnalysisBeautifulStyle_(sheet) {
  try {
    // Очищаем лист
    sheet.clear();
    
    // Устанавливаем фон всего листа
    const maxRows = 100;
    const maxCols = 15;
    sheet.getRange(1, 1, maxRows, maxCols)
      .setBackground('#FFFFFF')
      .setFontFamily('Arial')
      .setFontSize(10);
    
    // ЗАГОЛОВОК ОТЧЕТА (A1:N1)
    sheet.getRange('A1:N1').merge();
    sheet.getRange('A1')
      .setValue('📊 АНАЛИЗ ПРИЧИН ОТКАЗОВ')
      .setBackground('#4285f4')
      .setFontColor('#FFFFFF')
      .setFontSize(16)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 50);
    
    // ВРЕМЯ ОБНОВЛЕНИЯ (A2:N2)  
    sheet.getRange('A2:N2').merge();
    sheet.getRange('A2')
      .setValue(`⏰ Последнее обновление: ${getCurrentDateMoscow_().toLocaleString()}`)
      .setBackground('#f8f9fa')
      .setFontSize(11)
      .setFontStyle('italic')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(2, 30);
    
    // БЛОК 1: ОБЩАЯ СТАТИСТИКА (строки 4-8)
    // Заголовок блока
    sheet.getRange('A4:F4').merge();
    sheet.getRange('A4')
      .setValue('🔍 ОБЩАЯ СТАТИСТИКА ОТКАЗОВ')
      .setBackground('#e3f2fd')
      .setFontColor('#1565c0')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(4, 35);
    
    // БЛОК 2: КАТЕГОРИИ ПРИЧИН (строки 10-20)
    sheet.getRange('A10:F10').merge();
    sheet.getRange('A10')
      .setValue('📊 КАТЕГОРИИ ПРИЧИН ОТКАЗОВ')
      .setBackground('#f3e5f5')
      .setFontColor('#7b1fa2')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(10, 35);
    
    // Заголовки таблицы категорий
    const categoryHeaders = ['Категория', 'Количество', 'Процент'];
    for (let i = 0; i < categoryHeaders.length; i++) {
      sheet.getRange(11, i + 1)
        .setValue(categoryHeaders[i])
        .setBackground('#9c27b0')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    }
    sheet.setRowHeight(11, 30);
    
    // БЛОК 3: КАНАЛЫ С ОТКАЗАМИ (строки 22-32)
    sheet.getRange('A22:G22').merge();
    sheet.getRange('A22')
      .setValue('🎯 КАНАЛЫ С НАИБОЛЬШИМ КОЛИЧЕСТВОМ ОТКАЗОВ')
      .setBackground('#e8f5e8')
      .setFontColor('#388e3c')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(22, 35);
    
    // Заголовки таблицы каналов
    const channelHeaders = ['Канал', 'Отказы', 'Средний чек', 'Основные причины'];
    for (let i = 0; i < channelHeaders.length; i++) {
      sheet.getRange(23, i + 1)
        .setValue(channelHeaders[i])
        .setBackground('#4caf50')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    }
    sheet.setRowHeight(23, 30);
    
    // БЛОК 4: ИНСАЙТЫ (строки 34-44)
    sheet.getRange('A34:F34').merge();
    sheet.getRange('A34')
      .setValue('💡 КЛЮЧЕВЫЕ ИНСАЙТЫ')
      .setBackground('#fff3e0')
      .setFontColor('#f57c00')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(34, 35);
    
    // БЛОК 5: РЕКОМЕНДАЦИИ (строки 46-56)
    sheet.getRange('A46:F46').merge();
    sheet.getRange('A46')
      .setValue('🎯 РЕКОМЕНДАЦИИ')
      .setBackground('#fce4ec')
      .setFontColor('#c2185b')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(46, 35);
    
    // Устанавливаем ширину столбцов
    sheet.setColumnWidth(1, 200);  // Колонка A - широкая для названий
    sheet.setColumnWidth(2, 120);  // Колонка B 
    sheet.setColumnWidth(3, 100);  // Колонка C
    sheet.setColumnWidth(4, 150);  // Колонка D
    sheet.setColumnWidth(5, 150);  // Колонка E
    sheet.setColumnWidth(6, 80);   // Колонка F
    
    // Замораживаем первые 2 строки
    sheet.setFrozenRows(2);
    
    logInfo_('REFUSAL_STYLE', 'Красивое оформление анализа отказов применено');
    
  } catch (error) {
    logError_('REFUSAL_STYLE', 'Ошибка применения красивого оформления', error);
  }
}
