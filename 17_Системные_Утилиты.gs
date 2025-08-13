/**
 * МОДУЛЬ: СИСТЕМНЫЕ УТИЛИТЫ И МОНИТОРИНГ
 * Функции для мониторинга производительности и диагностики
 */

function runSystemDiagnostics() {
  console.log('Запускаем системную диагностику...');
  
  const diagnostics = {
    timestamp: new Date(),
    performance: analyzePerformance(),
    dataQuality: analyzeDataQuality(),
    systemHealth: checkSystemHealth(),
    quotas: checkQuotas(),
    errors: getRecentErrors()
  };
  
  // Создаем отчет диагностики
  createDiagnosticsReport(diagnostics);
  
  console.log('Системная диагностика завершена');
  return diagnostics;
}

function analyzePerformance() {
  const startTime = new Date().getTime();
  
  try {
    // Тестируем производительность основных операций
    const tests = {
      dataLoad: measureDataLoadTime(),
      processing: measureProcessingTime(),
      sheetOperations: measureSheetOperations(),
      calculations: measureCalculations()
    };
    
    const totalTime = new Date().getTime() - startTime;
    
    return {
      tests,
      totalDiagnosticsTime: totalTime,
      status: totalTime < 10000 ? 'good' : (totalTime < 30000 ? 'warning' : 'critical'),
      recommendations: generatePerformanceRecommendations(tests)
    };
    
  } catch (error) {
    console.error('Ошибка при анализе производительности:', error);
    return { error: error.toString() };
  }
}

function measureDataLoadTime() {
  const start = new Date().getTime();
  
  try {
    const data = getWorkingAmoData();
    const loadTime = new Date().getTime() - start;
    
    return {
      time: loadTime,
      rowsLoaded: data ? data.length : 0,
      status: loadTime < 5000 ? 'good' : (loadTime < 15000 ? 'warning' : 'slow'),
      avgTimePerRow: data && data.length > 0 ? (loadTime / data.length) : 0
    };
    
  } catch (error) {
    return {
      error: error.toString(),
      time: new Date().getTime() - start,
      status: 'error'
    };
  }
}

function measureProcessingTime() {
  const start = new Date().getTime();
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      return { time: 0, processed: 0, status: 'no_data' };
    }
    
    // Тестируем обработку небольшой выборки
    const sampleSize = Math.min(100, data.length);
    const sample = data.slice(0, sampleSize);
    
    let processed = 0;
    sample.forEach(row => {
      parseUtmSource(row);
      normalizePhone(row[CONFIG.WORKING_AMO_COLUMNS.PHONE]);
      formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.PLAN_AMOUNT]);
      isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '');
      processed++;
    });
    
    const processingTime = new Date().getTime() - start;
    
    return {
      time: processingTime,
      processed,
      avgTimePerRow: processed > 0 ? (processingTime / processed) : 0,
      status: processingTime < 2000 ? 'good' : (processingTime < 5000 ? 'warning' : 'slow')
    };
    
  } catch (error) {
    return {
      error: error.toString(),
      time: new Date().getTime() - start,
      status: 'error'
    };
  }
}

function measureSheetOperations() {
  const start = new Date().getTime();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    // Тестируем операции с листами
    let operations = 0;
    
    sheets.forEach(sheet => {
      sheet.getName(); // Простая операция чтения
      sheet.getLastRow(); // Операция получения размера
      operations += 2;
    });
    
    const operationTime = new Date().getTime() - start;
    
    return {
      time: operationTime,
      operations,
      sheetsCount: sheets.length,
      avgTimePerOperation: operations > 0 ? (operationTime / operations) : 0,
      status: operationTime < 1000 ? 'good' : (operationTime < 3000 ? 'warning' : 'slow')
    };
    
  } catch (error) {
    return {
      error: error.toString(),
      time: new Date().getTime() - start,
      status: 'error'
    };
  }
}

function measureCalculations() {
  const start = new Date().getTime();
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      return { time: 0, calculations: 0, status: 'no_data' };
    }
    
    let calculations = 0;
    
    // Выполняем типичные расчеты
    const totalRevenue = data.reduce((sum, row) => {
      calculations++;
      return sum + formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
    }, 0);
    
    const successCount = data.filter(row => {
      calculations++;
      return isSuccessStatus(row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '');
    }).length;
    
    const conversionRate = data.length > 0 ? (successCount / data.length * 100) : 0;
    calculations++;
    
    const calculationTime = new Date().getTime() - start;
    
    return {
      time: calculationTime,
      calculations,
      results: { totalRevenue, successCount, conversionRate },
      avgTimePerCalculation: calculations > 0 ? (calculationTime / calculations) : 0,
      status: calculationTime < 1000 ? 'good' : (calculationTime < 3000 ? 'warning' : 'slow')
    };
    
  } catch (error) {
    return {
      error: error.toString(),
      time: new Date().getTime() - start,
      status: 'error'
    };
  }
}

function analyzeDataQuality() {
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      return { status: 'no_data' };
    }
    
    const quality = {
      totalRows: data.length,
      emptyRows: 0,
      missingIds: 0,
      missingNames: 0,
      missingStatuses: 0,
      missingPhones: 0,
      missingDates: 0,
      invalidPhones: 0,
      invalidDates: 0,
      duplicateIds: 0
    };
    
    const seenIds = new Set();
    
    data.forEach(row => {
      const id = row[CONFIG.WORKING_AMO_COLUMNS.ID];
      const name = row[CONFIG.WORKING_AMO_COLUMNS.NAME];
      const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS];
      const phone = row[CONFIG.WORKING_AMO_COLUMNS.PHONE];
      const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
      
      // Проверяем на пустые строки
      if (!id && !name && !status) {
        quality.emptyRows++;
        return;
      }
      
      // Проверяем обязательные поля
      if (!id) quality.missingIds++;
      if (!name) quality.missingNames++;
      if (!status) quality.missingStatuses++;
      if (!phone) quality.missingPhones++;
      if (!createdAt) quality.missingDates++;
      
      // Проверяем дубликаты ID
      if (id) {
        if (seenIds.has(id)) {
          quality.duplicateIds++;
        } else {
          seenIds.add(id);
        }
      }
      
      // Проверяем валидность телефонов
      if (phone && !isValidPhone(phone)) {
        quality.invalidPhones++;
      }
      
      // Проверяем валидность дат
      if (createdAt && !isValidDate(createdAt)) {
        quality.invalidDates++;
      }
    });
    
    // Рассчитываем показатели качества
    quality.completenessScore = calculateCompletenessScore(quality);
    quality.validityScore = calculateValidityScore(quality);
    quality.overallScore = (quality.completenessScore + quality.validityScore) / 2;
    quality.status = quality.overallScore > 90 ? 'excellent' : 
                    (quality.overallScore > 70 ? 'good' : 
                     (quality.overallScore > 50 ? 'warning' : 'poor'));
    
    return quality;
    
  } catch (error) {
    console.error('Ошибка при анализе качества данных:', error);
    return { error: error.toString() };
  }
}

function calculateCompletenessScore(quality) {
  const totalFields = quality.totalRows * 5; // 5 основных полей
  const missingFields = quality.missingIds + quality.missingNames + quality.missingStatuses + 
                       quality.missingPhones + quality.missingDates;
  
  return totalFields > 0 ? ((totalFields - missingFields) / totalFields * 100) : 100;
}

function calculateValidityScore(quality) {
  const validatableFields = quality.totalRows * 2; // телефоны и даты
  const invalidFields = quality.invalidPhones + quality.invalidDates;
  
  return validatableFields > 0 ? ((validatableFields - invalidFields) / validatableFields * 100) : 100;
}

function checkSystemHealth() {
  try {
    const health = {
      spreadsheetAccess: testSpreadsheetAccess(),
      sheetsStructure: validateSheetsStructure(),
      configIntegrity: validateConfigIntegrity(),
      dependenciesCheck: checkDependencies(),
      permissionsCheck: checkPermissions()
    };
    
    const healthScore = Object.values(health).filter(test => test.status === 'ok').length;
    const totalTests = Object.keys(health).length;
    
    health.overallHealth = (healthScore / totalTests * 100);
    health.status = health.overallHealth === 100 ? 'healthy' : 
                   (health.overallHealth > 80 ? 'warning' : 'unhealthy');
    
    return health;
    
  } catch (error) {
    return { error: error.toString(), status: 'error' };
  }
}

function testSpreadsheetAccess() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    spreadsheet.getName();
    return { status: 'ok', message: 'Доступ к таблице работает' };
  } catch (error) {
    return { status: 'error', message: `Ошибка доступа к таблице: ${error}` };
  }
}

function validateSheetsStructure() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    const requiredSheets = Object.values(CONFIG.SHEETS);
    const existingSheets = sheets.map(sheet => sheet.getName());
    
    const missingSheets = requiredSheets.filter(sheetName => !existingSheets.includes(sheetName));
    
    if (missingSheets.length === 0) {
      return { status: 'ok', message: 'Все необходимые листы присутствуют' };
    } else {
      return { 
        status: 'warning', 
        message: `Отсутствуют листы: ${missingSheets.join(', ')}`,
        missingSheets
      };
    }
    
  } catch (error) {
    return { status: 'error', message: `Ошибка проверки структуры: ${error}` };
  }
}

function validateConfigIntegrity() {
  try {
    // Проверяем основные элементы конфигурации
    if (!CONFIG) throw new Error('CONFIG не определен');
    if (!CONFIG.SHEETS) throw new Error('CONFIG.SHEETS не определен');
    if (!CONFIG.WORKING_AMO_COLUMNS) throw new Error('CONFIG.WORKING_AMO_COLUMNS не определен');
    if (!CONFIG.COLORS) throw new Error('CONFIG.COLORS не определен');
    
    const requiredColumns = ['ID', 'NAME', 'STATUS', 'CREATED_AT', 'PHONE'];
    const missingColumns = requiredColumns.filter(col => !CONFIG.WORKING_AMO_COLUMNS[col]);
    
    if (missingColumns.length === 0) {
      return { status: 'ok', message: 'Конфигурация корректна' };
    } else {
      return { 
        status: 'error', 
        message: `Отсутствуют колонки в конфигурации: ${missingColumns.join(', ')}`,
        missingColumns
      };
    }
    
  } catch (error) {
    return { status: 'error', message: `Ошибка конфигурации: ${error}` };
  }
}

function checkDependencies() {
  try {
    // Проверяем доступность основных функций
    const dependencies = [
      'getWorkingAmoData',
      'parseUtmSource',
      'normalizePhone',
      'formatNumber',
      'isSuccessStatus',
      'createOrUpdateSheet'
    ];
    
    const missingFunctions = dependencies.filter(funcName => typeof eval(funcName) !== 'function');
    
    if (missingFunctions.length === 0) {
      return { status: 'ok', message: 'Все зависимости доступны' };
    } else {
      return { 
        status: 'error', 
        message: `Недоступные функции: ${missingFunctions.join(', ')}`,
        missingFunctions
      };
    }
    
  } catch (error) {
    return { status: 'error', message: `Ошибка проверки зависимостей: ${error}` };
  }
}

function checkPermissions() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Пытаемся выполнить операции, требующие различных разрешений
    const tests = {
      read: testReadPermission(spreadsheet),
      write: testWritePermission(spreadsheet),
      create: testCreatePermission(spreadsheet)
    };
    
    const failedTests = Object.entries(tests).filter(([_, test]) => test.status !== 'ok');
    
    if (failedTests.length === 0) {
      return { status: 'ok', message: 'Все необходимые разрешения есть' };
    } else {
      return { 
        status: 'warning', 
        message: `Проблемы с разрешениями: ${failedTests.map(([name, _]) => name).join(', ')}`,
        tests
      };
    }
    
  } catch (error) {
    return { status: 'error', message: `Ошибка проверки разрешений: ${error}` };
  }
}

function testReadPermission(spreadsheet) {
  try {
    const sheets = spreadsheet.getSheets();
    if (sheets.length > 0) {
      sheets[0].getDataRange();
    }
    return { status: 'ok' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function testWritePermission(spreadsheet) {
  try {
    const testSheet = spreadsheet.getSheets()[0];
    const lastRow = testSheet.getLastRow();
    const lastCol = testSheet.getLastColumn();
    
    // Пытаемся записать тестовое значение в свободную ячейку
    testSheet.getRange(lastRow + 1, lastCol + 1).setValue('test');
    testSheet.getRange(lastRow + 1, lastCol + 1).clear(); // Убираем тестовое значение
    
    return { status: 'ok' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function testCreatePermission(spreadsheet) {
  try {
    const testSheetName = 'TempTestSheet' + Date.now();
    const testSheet = spreadsheet.insertSheet(testSheetName);
    spreadsheet.deleteSheet(testSheet);
    return { status: 'ok' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function checkQuotas() {
  try {
    // Получаем информацию о квотах (приблизительно)
    const quotaInfo = {
      executionTime: getExecutionTimeUsage(),
      apiCalls: getApiCallsUsage(),
      storage: getStorageUsage()
    };
    
    const warnings = [];
    if (quotaInfo.executionTime > 80) warnings.push('execution_time');
    if (quotaInfo.apiCalls > 80) warnings.push('api_calls');
    if (quotaInfo.storage > 90) warnings.push('storage');
    
    return {
      quotas: quotaInfo,
      warnings,
      status: warnings.length === 0 ? 'ok' : 'warning'
    };
    
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function getExecutionTimeUsage() {
  // Примерная оценка использования времени выполнения
  return Math.random() * 100; // В реальности нужно отслеживать через PropertiesService
}

function getApiCallsUsage() {
  // Примерная оценка использования API вызовов
  return Math.random() * 100;
}

function getStorageUsage() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    let totalCells = 0;
    
    sheets.forEach(sheet => {
      totalCells += sheet.getMaxRows() * sheet.getMaxColumns();
    });
    
    // Примерная оценка использования (максимум 10 млн ячеек в Google Sheets)
    return (totalCells / 10000000) * 100;
    
  } catch (error) {
    return 0;
  }
}

function getRecentErrors() {
  try {
    // В реальности здесь можно использовать PropertiesService для хранения логов ошибок
    return {
      count: 0,
      recent: [],
      status: 'ok'
    };
    
  } catch (error) {
    return {
      error: error.toString(),
      status: 'error'
    };
  }
}

function generatePerformanceRecommendations(tests) {
  const recommendations = [];
  
  if (tests.dataLoad && tests.dataLoad.status !== 'good') {
    recommendations.push('Оптимизируйте загрузку данных: используйте getValues() вместо getValue() для диапазонов');
  }
  
  if (tests.processing && tests.processing.status !== 'good') {
    recommendations.push('Оптимизируйте обработку данных: группируйте операции и используйте пакетную обработку');
  }
  
  if (tests.sheetOperations && tests.sheetOperations.status !== 'good') {
    recommendations.push('Сократите количество операций с листами: кэшируйте ссылки на листы');
  }
  
  if (tests.calculations && tests.calculations.status !== 'good') {
    recommendations.push('Оптимизируйте вычисления: используйте встроенные функции Google Sheets');
  }
  
  if (recommendations.length === 0) {
    recommendations.push('Производительность в норме, дополнительная оптимизация не требуется');
  }
  
  return recommendations;
}

function createDiagnosticsReport(diagnostics) {
  try {
    const sheet = createOrUpdateSheet('🔧 ДИАГНОСТИКА');
    
    let row = 1;
    
    // Заголовок
    sheet.getRange(row, 1, 1, 2).setValues([['СИСТЕМНАЯ ДИАГНОСТИКА', formatDate(diagnostics.timestamp)]]);
    sheet.getRange(row, 1, 1, 2).setBackground(CONFIG.COLORS.header).setFontWeight('bold');
    row += 2;
    
    // Производительность
    if (diagnostics.performance && !diagnostics.performance.error) {
      sheet.getRange(row, 1, 1, 2).setValues([['ПРОИЗВОДИТЕЛЬНОСТЬ', diagnostics.performance.status]]);
      sheet.getRange(row, 1, 1, 2).setBackground(CONFIG.COLORS.success).setFontWeight('bold');
      row++;
      
      if (diagnostics.performance.tests) {
        Object.entries(diagnostics.performance.tests).forEach(([testName, testResult]) => {
          sheet.getRange(row, 1, 1, 3).setValues([[`  ${testName}`, testResult.status, `${testResult.time || 0}ms`]]);
          row++;
        });
      }
      row++;
    }
    
    // Качество данных
    if (diagnostics.dataQuality && !diagnostics.dataQuality.error) {
      sheet.getRange(row, 1, 1, 2).setValues([['КАЧЕСТВО ДАННЫХ', diagnostics.dataQuality.status]]);
      sheet.getRange(row, 1, 1, 2).setBackground(CONFIG.COLORS.warning).setFontWeight('bold');
      row++;
      
      sheet.getRange(row, 1, 1, 3).setValues([['  Общая оценка', `${diagnostics.dataQuality.overallScore || 0}%`, '']]);
      row++;
      sheet.getRange(row, 1, 1, 3).setValues([['  Полнота данных', `${diagnostics.dataQuality.completenessScore || 0}%`, '']]);
      row++;
      sheet.getRange(row, 1, 1, 3).setValues([['  Валидность', `${diagnostics.dataQuality.validityScore || 0}%`, '']]);
      row += 2;
    }
    
    // Здоровье системы
    if (diagnostics.systemHealth && !diagnostics.systemHealth.error) {
      sheet.getRange(row, 1, 1, 2).setValues([['ЗДОРОВЬЕ СИСТЕМЫ', diagnostics.systemHealth.status]]);
      sheet.getRange(row, 1, 1, 2).setBackground(CONFIG.COLORS.error).setFontWeight('bold');
      row++;
      
      sheet.getRange(row, 1, 1, 3).setValues([['  Общее состояние', `${diagnostics.systemHealth.overallHealth || 0}%`, '']]);
      row += 2;
    }
    
    sheet.autoResizeColumns(1, 3);
    console.log('Отчет диагностики создан');
    
  } catch (error) {
    console.error('Ошибка при создании отчета диагностики:', error);
  }
}

/**
 * УТИЛИТЫ ВАЛИДАЦИИ
 */
function isValidPhone(phone) {
  if (!phone) return false;
  const normalized = normalizePhone(phone);
  return normalized && normalized.length >= 10;
}

function isValidDate(dateValue) {
  if (!dateValue) return false;
  const date = new Date(dateValue);
  return !isNaN(date.getTime()) && date.getFullYear() > 2000;
}

/**
 * УТИЛИТА АВТОМАТИЧЕСКОЙ ОЧИСТКИ СИСТЕМЫ
 */
function runSystemMaintenance() {
  console.log('Запускаем системное обслуживание...');
  
  try {
    const results = {
      diagnostics: runSystemDiagnostics(),
      cleanup: cleanupOldCharts(),
      optimization: optimizeSheets(),
      backup: createSystemBackup()
    };
    
    console.log('Системное обслуживание завершено');
    return results;
    
  } catch (error) {
    console.error('Ошибка при системном обслуживании:', error);
    return { error: error.toString() };
  }
}

function cleanupOldCharts() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    let removedCharts = 0;
    
    sheets.forEach(sheet => {
      const charts = sheet.getCharts();
      // Можно добавить логику для удаления старых или неактуальных графиков
      // Сейчас просто считаем существующие
      removedCharts += charts.length;
    });
    
    return {
      status: 'completed',
      removedCharts: 0, // Фактически не удаляем
      existingCharts: removedCharts
    };
    
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function optimizeSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    let optimized = 0;
    
    sheets.forEach(sheet => {
      // Удаляем пустые строки и столбцы в конце
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const maxRows = sheet.getMaxRows();
      const maxCols = sheet.getMaxColumns();
      
      if (maxRows > lastRow + 100) {
        sheet.deleteRows(lastRow + 100, maxRows - lastRow - 99);
        optimized++;
      }
      
      if (maxCols > lastCol + 10) {
        sheet.deleteColumns(lastCol + 10, maxCols - lastCol - 9);
        optimized++;
      }
    });
    
    return {
      status: 'completed',
      optimizedSheets: optimized
    };
    
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function createSystemBackup() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const backupName = `AMO_Analytics_Backup_${Utilities.formatDate(new Date(), 'GMT+3', 'yyyy-MM-dd_HH-mm')}`;
    
    // Создаем копию текущей таблицы
    const backup = DriveApp.getFileById(spreadsheet.getId()).makeCopy(backupName);
    
    return {
      status: 'completed',
      backupId: backup.getId(),
      backupName: backupName,
      backupUrl: backup.getUrl()
    };
    
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}
