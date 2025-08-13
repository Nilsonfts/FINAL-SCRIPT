/**
 * МОДУЛЬ: ДЕМОНСТРАЦИЯ И ТЕСТИРОВАНИЕ
 * Функции для демонстрации возможностей системы и тестирования
 */

function runFullDemo() {
  console.log('🚀 Запускаем полную демонстрацию AMO Analytics...');
  
  try {
    const startTime = new Date().getTime();
    
    // Этап 1: Проверяем данные
    console.log('📊 Этап 1: Загрузка и проверка данных...');
    const data = getWorkingAmoData();
    
    if (!data || data.length === 0) {
      console.log('❌ Нет данных для демонстрации');
      return;
    }
    
    console.log(`✅ Загружено ${data.length} записей`);
    
    // Этап 2: Системная диагностика
    console.log('🔧 Этап 2: Системная диагностика...');
    const diagnostics = runSystemDiagnostics();
    console.log(`✅ Диагностика завершена со статусом: ${diagnostics.performance?.status || 'unknown'}`);
    
    // Этап 3: Создаем все аналитические отчеты
    console.log('📈 Этап 3: Создание аналитических отчетов...');
    
    const reports = [
      { name: 'Лиды по каналам', func: analyzeLeadsByChannels },
      { name: 'Маркетинг и каналы', func: analyzeMarketingChannels },
      { name: 'Первые касания', func: analyzeFirstTouchPoints },
      { name: 'Колл-трекинг', func: analyzeCallTracking },
      { name: 'Причины отказов', func: analyzeRefusalReasons },
      { name: 'Сводная аналитика', func: createAmoSummary },
      { name: 'Месячный дашборд', func: createMonthlyDashboard },
      { name: 'Анализ каналов', func: analyzeChannels }
    ];
    
    let successfulReports = 0;
    
    reports.forEach(report => {
      try {
        console.log(`  📋 Создаем: ${report.name}...`);
        report.func();
        successfulReports++;
        console.log(`  ✅ ${report.name} создан`);
      } catch (error) {
        console.warn(`  ⚠️ Ошибка в ${report.name}:`, error);
      }
    });
    
    // Этап 4: Создаем визуализации
    console.log('📊 Этап 4: Создание графиков и визуализаций...');
    try {
      createVisualizationDashboard();
      console.log('✅ Дашборд визуализации создан');
    } catch (error) {
      console.warn('⚠️ Ошибка при создании визуализации:', error);
    }
    
    // Этап 5: Экспорт данных
    console.log('💾 Этап 5: Экспорт аналитических данных...');
    try {
      const exportResult = exportAnalyticsReport();
      console.log('✅ Аналитический отчет экспортирован');
    } catch (error) {
      console.warn('⚠️ Ошибка при экспорте:', error);
    }
    
    // Этап 6: Финальная статистика
    const endTime = new Date().getTime();
    const totalTime = endTime - startTime;
    
    console.log('🎉 Демонстрация завершена!');
    console.log(`📊 Обработано записей: ${data.length}`);
    console.log(`📈 Создано отчетов: ${successfulReports}/${reports.length}`);
    console.log(`⏱️ Общее время: ${(totalTime/1000).toFixed(2)} секунд`);
    
    // Создаем сводный отчет демонстрации
    createDemoSummaryReport({
      dataCount: data.length,
      reportsCreated: successfulReports,
      totalReports: reports.length,
      executionTime: totalTime,
      diagnostics: diagnostics
    });
    
    return {
      success: true,
      dataCount: data.length,
      reportsCreated: successfulReports,
      executionTime: totalTime
    };
    
  } catch (error) {
    console.error('❌ Ошибка при демонстрации:', error);
    return { success: false, error: error.toString() };
  }
}

function createDemoSummaryReport(demoResults) {
  try {
    const sheet = createOrUpdateSheet('🎯 ДЕМО ОТЧЕТ');
    
    const reportData = [
      ['AMO ANALYTICS - ДЕМОНСТРАЦИЯ СИСТЕМЫ', ''],
      ['Дата проведения', formatDate(new Date())],
      ['', ''],
      ['РЕЗУЛЬТАТЫ ДЕМОНСТРАЦИИ', ''],
      ['Количество записей', demoResults.dataCount],
      ['Создано отчетов', `${demoResults.reportsCreated}/${demoResults.totalReports}`],
      ['Время выполнения', `${(demoResults.executionTime/1000).toFixed(2)} сек`],
      ['Производительность', demoResults.diagnostics?.performance?.status || 'N/A'],
      ['', ''],
      ['ДОСТУПНЫЕ МОДУЛИ', ''],
      ['✅ Главный скрипт', 'Основная обработка AMO данных'],
      ['✅ Конфигурация', 'Настройки и константы системы'],
      ['✅ Основные утилиты', 'Функции обработки данных'],
      ['✅ Лиды по каналам', 'Анализ эффективности каналов'],
      ['✅ Маркетинг и каналы', 'Маркетинговая аналитика'],
      ['✅ Первые касания', 'Анализ точек входа'],
      ['✅ Колл-трекинг', 'Анализ телефонных звонков'],
      ['✅ Причины отказов AI', 'ИИ-анализ причин отказов'],
      ['✅ Яндекс Директ', 'Анализ рекламных кампаний'],
      ['✅ Сквозная аналитика сайт', 'Веб-аналитика'],
      ['✅ Сводная аналитика АМО', 'Общие показатели CRM'],
      ['✅ Яндекс Метрика', 'Интеграция с Метрикой'],
      ['✅ Месячный дашборд', 'Ключевые показатели'],
      ['✅ Анализ каналов', 'Детальная атрибуция'],
      ['✅ Утилиты графиков', 'Создание визуализаций'],
      ['✅ Экспорт/импорт', 'Работа с внешними данными'],
      ['✅ Системные утилиты', 'Мониторинг и диагностика'],
      ['✅ Демо и тесты', 'Демонстрация и тестирование']
    ];
    
    sheet.getRange(1, 1, reportData.length, 2).setValues(reportData);
    
    // Форматирование
    sheet.getRange(1, 1, 1, 2).setBackground(CONFIG.COLORS.header).setFontWeight('bold').setFontSize(14);
    sheet.getRange(4, 1, 1, 2).setBackground(CONFIG.COLORS.success).setFontWeight('bold');
    sheet.getRange(10, 1, 1, 2).setBackground(CONFIG.COLORS.success).setFontWeight('bold');
    
    sheet.autoResizeColumns(1, 2);
    console.log('📋 Сводный отчет демонстрации создан');
    
  } catch (error) {
    console.error('Ошибка при создании сводного отчета:', error);
  }
}

function runPerformanceTests() {
  console.log('🔬 Запускаем тесты производительности...');
  
  try {
    const tests = [
      { name: 'Загрузка данных', test: testDataLoading },
      { name: 'Обработка UTM', test: testUtmProcessing },
      { name: 'Нормализация телефонов', test: testPhoneNormalization },
      { name: 'Расчет конверсий', test: testConversionCalculations },
      { name: 'Создание отчета', test: testReportCreation },
      { name: 'Экспорт данных', test: testDataExport }
    ];
    
    const results = [];
    
    tests.forEach(testCase => {
      console.log(`  🧪 Тестируем: ${testCase.name}...`);
      const startTime = new Date().getTime();
      
      try {
        const testResult = testCase.test();
        const endTime = new Date().getTime();
        
        results.push({
          name: testCase.name,
          success: true,
          time: endTime - startTime,
          result: testResult
        });
        
        console.log(`  ✅ ${testCase.name}: ${endTime - startTime}ms`);
        
      } catch (error) {
        const endTime = new Date().getTime();
        results.push({
          name: testCase.name,
          success: false,
          time: endTime - startTime,
          error: error.toString()
        });
        
        console.log(`  ❌ ${testCase.name}: ${error}`);
      }
    });
    
    // Создаем отчет тестирования
    createPerformanceTestReport(results);
    
    const successCount = results.filter(r => r.success).length;
    console.log(`🏁 Тестирование завершено: ${successCount}/${tests.length} тестов прошли успешно`);
    
    return results;
    
  } catch (error) {
    console.error('Ошибка при тестировании производительности:', error);
    return { error: error.toString() };
  }
}

function testDataLoading() {
  const data = getWorkingAmoData();
  return {
    recordsLoaded: data ? data.length : 0,
    firstRecord: data && data.length > 0 ? data[0] : null
  };
}

function testUtmProcessing() {
  const data = getWorkingAmoData();
  if (!data || data.length === 0) return { processed: 0 };
  
  const sampleSize = Math.min(100, data.length);
  const processed = [];
  
  for (let i = 0; i < sampleSize; i++) {
    const row = data[i];
    const utmSource = parseUtmSource(row);
    const channelType = getChannelType(utmSource);
    
    processed.push({
      originalSource: row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || '',
      parsedSource: utmSource,
      channelType: channelType
    });
  }
  
  return { processed: processed.length, samples: processed.slice(0, 5) };
}

function testPhoneNormalization() {
  const testPhones = [
    '+7 (912) 345-67-89',
    '89123456789',
    '9123456789',
    '+7-912-345-67-89',
    '7 912 345 67 89',
    'invalid phone',
    '',
    null
  ];
  
  const results = testPhones.map(phone => ({
    original: phone,
    normalized: normalizePhone(phone),
    isValid: isValidPhone(phone)
  }));
  
  return { tested: testPhones.length, results };
}

function testConversionCalculations() {
  const data = getWorkingAmoData();
  if (!data || data.length === 0) return { calculated: 0 };
  
  const sampleSize = Math.min(1000, data.length);
  let totalLeads = 0;
  let successLeads = 0;
  let totalRevenue = 0;
  
  for (let i = 0; i < sampleSize; i++) {
    const row = data[i];
    totalLeads++;
    
    const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
    if (isSuccessStatus(status)) {
      successLeads++;
      totalRevenue += formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    }
  }
  
  return {
    totalLeads,
    successLeads,
    conversionRate: totalLeads > 0 ? (successLeads / totalLeads * 100) : 0,
    totalRevenue,
    avgRevenue: successLeads > 0 ? (totalRevenue / successLeads) : 0
  };
}

function testReportCreation() {
  // Создаем тестовый лист
  const testSheetName = 'TEST_SHEET_' + Date.now();
  const sheet = createOrUpdateSheet(testSheetName);
  
  // Добавляем тестовые данные
  const testData = [
    ['Метрика', 'Значение'],
    ['Тестовый показатель 1', 100],
    ['Тестовый показатель 2', 200],
    ['Итого', 300]
  ];
  
  sheet.getRange(1, 1, testData.length, 2).setValues(testData);
  sheet.getRange(1, 1, 1, 2).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
  
  // Удаляем тестовый лист
  try {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
  } catch (error) {
    console.warn('Не удалось удалить тестовый лист:', error);
  }
  
  return { created: true, rowsAdded: testData.length };
}

function testDataExport() {
  const data = getWorkingAmoData();
  if (!data || data.length === 0) return { exported: 0 };
  
  // Тестируем создание CSV контента
  const sampleData = data.slice(0, 10);
  const csvContent = createCSVContent(sampleData);
  
  return {
    recordsExported: sampleData.length,
    csvLength: csvContent.length,
    hasHeaders: csvContent.includes('ID,Имя,Статус')
  };
}

function createPerformanceTestReport(results) {
  try {
    const sheet = createOrUpdateSheet('🧪 ТЕСТЫ ПРОИЗВОДИТЕЛЬНОСТИ');
    
    let row = 1;
    
    // Заголовок
    sheet.getRange(row, 1, 1, 4).setValues([['ТЕСТЫ ПРОИЗВОДИТЕЛЬНОСТИ', '', '', formatDate(new Date())]]);
    sheet.getRange(row, 1, 1, 4).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
    row += 2;
    
    // Заголовки таблицы
    sheet.getRange(row, 1, 1, 4).setValues([['Тест', 'Статус', 'Время (мс)', 'Результат']]);
    sheet.getRange(row, 1, 1, 4).setBackground(CONFIG.COLORS.success).setFontWeight('bold');
    row++;
    
    // Результаты тестов
    results.forEach(result => {
      const status = result.success ? '✅ PASS' : '❌ FAIL';
      const resultText = result.success ? 
        (typeof result.result === 'object' ? JSON.stringify(result.result).substring(0, 50) : result.result) :
        result.error;
      
      sheet.getRange(row, 1, 1, 4).setValues([[
        result.name,
        status,
        result.time,
        resultText
      ]]);
      
      // Подсветка статуса
      const statusColor = result.success ? CONFIG.COLORS.success : CONFIG.COLORS.error;
      sheet.getRange(row, 2).setBackground(statusColor);
      
      row++;
    });
    
    // Статистика
    row++;
    const totalTests = results.length;
    const passedTests = results.filter(r => r.success).length;
    const avgTime = results.reduce((sum, r) => sum + r.time, 0) / totalTests;
    
    sheet.getRange(row, 1, 1, 2).setValues([['СТАТИСТИКА', '']]);
    sheet.getRange(row, 1, 1, 2).setBackground(CONFIG.COLORS.warning).setFontWeight('bold');
    row++;
    
    const stats = [
      ['Всего тестов', totalTests],
      ['Прошли', passedTests],
      ['Провалились', totalTests - passedTests],
      ['Процент успеха', `${(passedTests/totalTests*100).toFixed(1)}%`],
      ['Среднее время', `${avgTime.toFixed(1)} мс`]
    ];
    
    sheet.getRange(row, 1, stats.length, 2).setValues(stats);
    
    sheet.autoResizeColumns(1, 4);
    console.log('📊 Отчет по тестам производительности создан');
    
  } catch (error) {
    console.error('Ошибка при создании отчета тестов:', error);
  }
}

function runIntegrationTests() {
  console.log('🔗 Запускаем интеграционные тесты...');
  
  try {
    const tests = [
      { name: 'Интеграция модулей', test: testModuleIntegration },
      { name: 'Целостность конфигурации', test: testConfigurationIntegrity },
      { name: 'Работа с листами', test: testSheetOperations },
      { name: 'Создание графиков', test: testChartCreation },
      { name: 'Обработка пайплайна', test: testDataPipeline }
    ];
    
    const results = [];
    
    tests.forEach(testCase => {
      console.log(`  🔗 Тестируем: ${testCase.name}...`);
      
      try {
        const testResult = testCase.test();
        results.push({
          name: testCase.name,
          success: true,
          result: testResult
        });
        console.log(`  ✅ ${testCase.name}: Успешно`);
        
      } catch (error) {
        results.push({
          name: testCase.name,
          success: false,
          error: error.toString()
        });
        console.log(`  ❌ ${testCase.name}: ${error}`);
      }
    });
    
    createIntegrationTestReport(results);
    
    const successCount = results.filter(r => r.success).length;
    console.log(`🏁 Интеграционное тестирование: ${successCount}/${tests.length} тестов прошли`);
    
    return results;
    
  } catch (error) {
    console.error('Ошибка при интеграционном тестировании:', error);
    return { error: error.toString() };
  }
}

function testModuleIntegration() {
  // Проверяем, что основные функции из разных модулей работают вместе
  const data = getWorkingAmoData();
  if (!data || data.length === 0) throw new Error('Нет данных для теста');
  
  const sampleRow = data[0];
  
  // Тестируем цепочку обработки
  const utmSource = parseUtmSource(sampleRow);
  const normalizedPhone = normalizePhone(sampleRow[CONFIG.WORKING_AMO_COLUMNS.PHONE]);
  const isSuccess = isSuccessStatus(sampleRow[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '');
  const planAmount = formatNumber(sampleRow[CONFIG.WORKING_AMO_COLUMNS.PLAN_AMOUNT]);
  
  return {
    dataLoaded: true,
    utmParsed: !!utmSource,
    phoneNormalized: !!normalizedPhone,
    statusChecked: typeof isSuccess === 'boolean',
    amountFormatted: typeof planAmount === 'number'
  };
}

function testConfigurationIntegrity() {
  // Проверяем целостность конфигурации
  const requiredKeys = ['SHEETS', 'WORKING_AMO_COLUMNS', 'COLORS'];
  const missingKeys = requiredKeys.filter(key => !CONFIG[key]);
  
  if (missingKeys.length > 0) {
    throw new Error(`Отсутствуют ключи конфигурации: ${missingKeys.join(', ')}`);
  }
  
  // Проверяем наличие основных листов
  const requiredSheets = ['LEADS_CHANNELS', 'AMO_SUMMARY', 'MARKETING_CHANNELS'];
  const missingSheets = requiredSheets.filter(sheet => !CONFIG.SHEETS[sheet]);
  
  if (missingSheets.length > 0) {
    throw new Error(`Отсутствуют настройки листов: ${missingSheets.join(', ')}`);
  }
  
  return { configValid: true, sheetsConfigured: true };
}

function testSheetOperations() {
  // Тестируем создание и операции с листами
  const testSheetName = 'INTEGRATION_TEST_' + Date.now();
  
  try {
    const sheet = createOrUpdateSheet(testSheetName);
    
    // Тестируем запись данных
    sheet.getRange(1, 1, 1, 2).setValues([['Тест', 'Значение']]);
    
    // Тестируем чтение данных
    const value = sheet.getRange(1, 1).getValue();
    
    if (value !== 'Тест') {
      throw new Error('Ошибка записи/чтения данных');
    }
    
    // Удаляем тестовый лист
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    
    return { sheetCreated: true, dataWritten: true, dataRead: true, sheetDeleted: true };
    
  } catch (error) {
    throw new Error(`Ошибка операций с листами: ${error}`);
  }
}

function testChartCreation() {
  // Тестируем создание графиков
  const testSheetName = 'CHART_TEST_' + Date.now();
  
  try {
    const sheet = createOrUpdateSheet(testSheetName);
    
    // Добавляем тестовые данные для графика
    const testData = [
      ['Категория', 'Значение'],
      ['A', 10],
      ['B', 20],
      ['C', 15]
    ];
    
    sheet.getRange(1, 1, testData.length, 2).setValues(testData);
    
    // Создаем график
    const chart = createPieChart(
      sheet,
      sheet.getRange(1, 1, testData.length, 2),
      'Тестовый график',
      { row: 6, col: 1 }
    );
    
    // Удаляем тестовый лист
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
    
    return { chartCreated: !!chart };
    
  } catch (error) {
    throw new Error(`Ошибка создания графика: ${error}`);
  }
}

function testDataPipeline() {
  // Тестируем полный пайплайн обработки данных
  const data = getWorkingAmoData();
  if (!data || data.length === 0) throw new Error('Нет данных для теста пайплайна');
  
  // Тестируем обработку небольшой выборки
  const sampleData = data.slice(0, 10);
  const processedData = [];
  
  sampleData.forEach(row => {
    const processed = {
      id: row[CONFIG.WORKING_AMO_COLUMNS.ID],
      source: parseUtmSource(row),
      channelType: getChannelType(parseUtmSource(row)),
      phone: normalizePhone(row[CONFIG.WORKING_AMO_COLUMNS.PHONE]),
      status: row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '',
      isSuccess: isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || ''),
      revenue: formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]),
      period: getDatePeriods(row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT])
    };
    
    processedData.push(processed);
  });
  
  return {
    originalRecords: sampleData.length,
    processedRecords: processedData.length,
    successfullyProcessed: processedData.filter(p => p.id && p.source).length
  };
}

function createIntegrationTestReport(results) {
  try {
    const sheet = createOrUpdateSheet('🔗 ИНТЕГРАЦИОННЫЕ ТЕСТЫ');
    
    let row = 1;
    
    // Заголовок
    sheet.getRange(row, 1, 1, 3).setValues([['ИНТЕГРАЦИОННЫЕ ТЕСТЫ', '', formatDate(new Date())]]);
    sheet.getRange(row, 1, 1, 3).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
    row += 2;
    
    // Заголовки таблицы
    sheet.getRange(row, 1, 1, 3).setValues([['Тест', 'Статус', 'Результат']]);
    sheet.getRange(row, 1, 1, 3).setBackground(CONFIG.COLORS.success).setFontWeight('bold');
    row++;
    
    // Результаты тестов
    results.forEach(result => {
      const status = result.success ? '✅ PASS' : '❌ FAIL';
      const resultText = result.success ? 
        JSON.stringify(result.result).substring(0, 100) :
        result.error.substring(0, 100);
      
      sheet.getRange(row, 1, 1, 3).setValues([[
        result.name,
        status,
        resultText
      ]]);
      
      // Подсветка статуса
      const statusColor = result.success ? CONFIG.COLORS.success : CONFIG.COLORS.error;
      sheet.getRange(row, 2).setBackground(statusColor);
      
      row++;
    });
    
    sheet.autoResizeColumns(1, 3);
    console.log('📊 Отчет интеграционных тестов создан');
    
  } catch (error) {
    console.error('Ошибка при создании отчета интеграционных тестов:', error);
  }
}

/**
 * ГЛАВНАЯ ФУНКЦИЯ ПОЛНОГО ТЕСТИРОВАНИЯ
 */
function runCompleteTestSuite() {
  console.log('🎯 Запускаем полный набор тестов AMO Analytics...');
  
  try {
    const testResults = {
      demo: runFullDemo(),
      performance: runPerformanceTests(),
      integration: runIntegrationTests(),
      system: runSystemDiagnostics()
    };
    
    console.log('🏆 ПОЛНОЕ ТЕСТИРОВАНИЕ ЗАВЕРШЕНО');
    console.log(`📊 Демо: ${testResults.demo?.success ? 'Успешно' : 'С ошибками'}`);
    console.log(`⚡ Производительность: ${testResults.performance?.filter(r => r.success).length || 0} тестов прошли`);
    console.log(`🔗 Интеграция: ${testResults.integration?.filter(r => r.success).length || 0} тестов прошли`);
    console.log(`🔧 Система: ${testResults.system?.systemHealth?.status || 'unknown'}`);
    
    return testResults;
    
  } catch (error) {
    console.error('❌ Ошибка при полном тестировании:', error);
    return { error: error.toString() };
  }
}
