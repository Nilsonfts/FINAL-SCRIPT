/**
 * БЫСТРЫЙ ТЕСТ СИСТЕМЫ
 * Проверяет работоспособность основных компонентов
 */

function quickSystemTest() {
  console.log('🧪 БЫСТРЫЙ ТЕСТ СИСТЕМЫ...');
  
  try {
    // 1. Проверяем конфигурацию
    console.log('1️⃣ Проверка конфигурации...');
    const configValid = validateConfiguration();
    console.log(`   Конфигурация: ${configValid ? '✅' : '❌'}`);
    
    // 2. Проверяем структуру листа
    console.log('2️⃣ Проверка структуры листа РАБОЧИЙ_АМО...');
    const structureReport = diagnoseWorkingAmoStructure();
    console.log(`   Структура: ${structureReport.validation?.valid ? '✅' : '❌'}`);
    
    // 3. Проверяем утилиты
    console.log('3️⃣ Проверка утилит...');
    const testNumber = formatNumber('1,234.56');
    const testPhone = normalizePhone('+7 (123) 456-78-90');
    const testStatus = isSuccessStatus('Оплачено');
    
    console.log(`   formatNumber: ${testNumber} (ожидается: 1234.56) ${testNumber === 1234.56 ? '✅' : '❌'}`);
    console.log(`   normalizePhone: ${testPhone} (ожидается: 71234567890) ${testPhone === '71234567890' ? '✅' : '❌'}`);
    console.log(`   isSuccessStatus: ${testStatus} (ожидается: true) ${testStatus === true ? '✅' : '❌'}`);
    
    // 4. Проверяем загрузку данных
    console.log('4️⃣ Проверка загрузки данных...');
    const data = getWorkingAmoData();
    console.log(`   Загружено строк: ${data.length} ${data.length > 0 ? '✅' : '❌'}`);
    
    if (data.length > 0) {
      console.log(`   Первая строка: ID=${data[0][0]}, Название=${data[0][1]}, Статус=${data[0][2]}`);
    }
    
    // 5. Итог
    console.log('🎯 РЕЗУЛЬТАТ ТЕСТА:');
    const allGood = configValid && structureReport.validation?.valid && data.length > 0;
    console.log(allGood ? '✅ ВСЁ РАБОТАЕТ КОРРЕКТНО!' : '❌ ЕСТЬ ПРОБЛЕМЫ, ТРЕБУЕТСЯ ИСПРАВЛЕНИЕ');
    
    return {
      success: allGood,
      config: configValid,
      structure: structureReport.validation?.valid,
      dataRows: data.length,
      utilities: testNumber === 1234.56 && testPhone === '71234567890' && testStatus === true
    };
    
  } catch (error) {
    console.error('❌ ОШИБКА ТЕСТА:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * ПОЛНАЯ ДИАГНОСТИКА СИСТЕМЫ
 */
function fullSystemDiagnostic() {
  console.log('🔍 ПОЛНАЯ ДИАГНОСТИКА СИСТЕМЫ...');
  
  try {
    // Быстрый тест
    const quickTest = quickSystemTest();
    
    // Дополнительные проверки
    console.log('🔬 ДОПОЛНИТЕЛЬНЫЕ ПРОВЕРКИ...');
    
    // Проверяем API токены
    const tokenStatus = getApiTokensStatus();
    console.log(`API токены: ${tokenStatus.configuredCount}/${tokenStatus.totalCount} (${tokenStatus.percentage}%)`);
    
    // Проверяем структуру колонок детально
    if (CONFIG.WORKING_AMO_COLUMNS) {
      const requiredColumns = ['ID', 'NAME', 'STATUS', 'RESPONSIBLE', 'PHONE', 'CREATED_AT'];
      const missingColumns = requiredColumns.filter(col => CONFIG.WORKING_AMO_COLUMNS[col] === undefined);
      
      if (missingColumns.length === 0) {
        console.log('✅ Все обязательные колонки настроены');
      } else {
        console.log('❌ Отсутствуют колонки:', missingColumns.join(', '));
      }
    }
    
    console.log('🎯 ФИНАЛЬНЫЙ СТАТУС:');
    if (quickTest.success) {
      console.log('✅ СИСТЕМА ПОЛНОСТЬЮ ГОТОВА К РАБОТЕ!');
      console.log('🚀 Можете запускать аналитику: runCompleteAnalysis()');
    } else {
      console.log('❌ СИСТЕМА ТРЕБУЕТ ИСПРАВЛЕНИЙ');
      console.log('🔧 Запустите: fixWorkingAmoStructureNow()');
    }
    
    return quickTest;
    
  } catch (error) {
    console.error('❌ ОШИБКА ДИАГНОСТИКИ:', error);
    return { success: false, error: error.toString() };
  }
}
