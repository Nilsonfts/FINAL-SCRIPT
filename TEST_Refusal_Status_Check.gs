/**
 * ТЕСТОВАЯ ФУНКЦИЯ ДЛЯ ДИАГНОСТИКИ СТАТУСОВ
 * Показывает все статусы сделок для настройки анализа отказов
 */

/**
 * Анализирует все статусы в РАБОЧИЙ АМО для диагностики
 * Показывает статистику по всем статусам сделок
 */
function analyzeAllStatusesForRefusal() {
  try {
    console.log('🔍 ДИАГНОСТИКА СТАТУСОВ для анализа отказов');
    console.log('=================================================');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
    
    if (!workingSheet) {
      console.log('❌ Лист "РАБОЧИЙ АМО" не найден');
      return;
    }
    
    const rawData = workingSheet.getDataRange().getValues();
    if (rawData.length <= 1) {
      console.log('❌ Нет данных в листе "РАБОЧИЙ АМО"');
      return;
    }
    
    const headers = rawData[0];
    const rows = rawData.slice(1);
    
    console.log(`📊 Всего записей: ${rows.length}`);
    
    // Находим колонку со статусом
    const statusIndex = findColumnIndex(headers, ['Статус', 'Status', 'Сделка.Статус']);
    
    if (statusIndex < 0) {
      console.log('❌ Колонка со статусом не найдена');
      console.log('📋 Доступные заголовки:', headers);
      return;
    }
    
    console.log(`✅ Колонка статуса найдена: ${headers[statusIndex]} (индекс ${statusIndex})`);
    
    // Собираем статистику по статусам
    const statusStats = {};
    let emptyStatuses = 0;
    
    rows.forEach(row => {
      const status = String(row[statusIndex] || '').trim();
      if (!status) {
        emptyStatuses++;
      } else {
        statusStats[status] = (statusStats[status] || 0) + 1;
      }
    });
    
    console.log('\n📈 СТАТИСТИКА СТАТУСОВ:');
    console.log('========================');
    
    // Сортируем статусы по количеству
    const sortedStatuses = Object.entries(statusStats)
      .sort(([,a], [,b]) => b - a);
    
    sortedStatuses.forEach(([status, count]) => {
      const percentage = ((count / rows.length) * 100).toFixed(1);
      console.log(`📌 "${status}": ${count} записей (${percentage}%)`);
    });
    
    if (emptyStatuses > 0) {
      const percentage = ((emptyStatuses / rows.length) * 100).toFixed(1);
      console.log(`⚠️  Пустые статусы: ${emptyStatuses} записей (${percentage}%)`);
    }
    
    // Проверяем какие из найденных статусов могут быть отказами
    console.log('\n🔍 АНАЛИЗ НА ПРЕДМЕТ ОТКАЗОВ:');
    console.log('=============================');
    
    const refusedKeywords = [
      'закрыто', 'отклон', 'отказ', 'failure', 'rejected', 
      'closed', 'реализов', 'провал', 'неуспешно'
    ];
    
    let potentialRefusals = 0;
    
    sortedStatuses.forEach(([status, count]) => {
      const statusLower = status.toLowerCase();
      const isRefusal = refusedKeywords.some(keyword => statusLower.includes(keyword));
      
      if (isRefusal) {
        console.log(`❌ ВОЗМОЖНЫЙ ОТКАЗ: "${status}" (${count} записей)`);
        potentialRefusals += count;
      } else {
        console.log(`✅ Успешный/промежуточный: "${status}" (${count} записей)`);
      }
    });
    
    console.log(`\n📊 ИТОГО потенциальных отказов: ${potentialRefusals} из ${rows.length} записей`);
    console.log(`📈 Процент отказов: ${((potentialRefusals / rows.length) * 100).toFixed(1)}%`);
    
    return {
      totalRecords: rows.length,
      statusStats: statusStats,
      potentialRefusals: potentialRefusals,
      refusalPercentage: (potentialRefusals / rows.length) * 100
    };
    
  } catch (error) {
    console.error('❌ Ошибка анализа статусов:', error);
    return { error: error.message };
  }
}

/**
 * Тестирует новую функцию getRefusedDealsData_
 */
function testRefusedDealsFunction() {
  try {
    console.log('🧪 ТЕСТ ФУНКЦИИ getRefusedDealsData_');
    console.log('====================================');
    
    const refusedDeals = getRefusedDealsData_();
    
    console.log(`📊 Найдено отказанных сделок: ${refusedDeals.length}`);
    
    if (refusedDeals.length > 0) {
      console.log('\n📋 ПРИМЕРЫ НАЙДЕННЫХ ОТКАЗОВ:');
      console.log('==============================');
      
      refusedDeals.slice(0, 5).forEach((deal, index) => {
        console.log(`\n${index + 1}. Сделка: ${deal.dealName}`);
        console.log(`   ID: ${deal.dealId}`);
        console.log(`   Статус: ${deal.status}`);
        console.log(`   Причина: ${deal.refusalComment}`);
        console.log(`   Канал: ${deal.channel}`);
        console.log(`   Менеджер: ${deal.manager}`);
      });
      
      if (refusedDeals.length > 5) {
        console.log(`\n... и еще ${refusedDeals.length - 5} отказов`);
      }
    } else {
      console.log('❌ Отказанные сделки не найдены');
      console.log('🔍 Запускаем диагностику статусов...');
      analyzeAllStatusesForRefusal();
    }
    
    return refusedDeals;
    
  } catch (error) {
    console.error('❌ Ошибка тестирования:', error);
    return [];
  }
}
