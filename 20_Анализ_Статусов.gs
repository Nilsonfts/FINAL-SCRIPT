/**
 * МОДУЛЬ: АНАЛИЗ СТАТУСОВ
 * Анализ всех статусов в системе для правильной настройки
 */

function analyzeAllStatuses() {
  console.log('📊 Анализируем все статусы в системе...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    // Собираем все уникальные статусы
    const statusMap = {};
    const successCount = {};
    const refusalCount = {};
    
    data.forEach(row => {
      const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
      if (status) {
        // Считаем количество каждого статуса
        if (!statusMap[status]) {
          statusMap[status] = 0;
        }
        statusMap[status]++;
        
        // Проверяем, является ли статус успешным
        if (isSuccessStatus(status)) {
          if (!successCount[status]) {
            successCount[status] = 0;
          }
          successCount[status]++;
        }
        
        // Проверяем, является ли статус отказом
        if (isRefusalStatus(status)) {
          if (!refusalCount[status]) {
            refusalCount[status] = 0;
          }
          refusalCount[status]++;
        }
      }
    });
    
    // Создаем отчет
    const sheet = createOrUpdateSheet('📊 АНАЛИЗ СТАТУСОВ');
    
    let currentRow = 1;
    
    // Заголовок
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue('АНАЛИЗ ВСЕХ СТАТУСОВ В СИСТЕМЕ')
         .setBackground('#1a237e')
         .setFontColor('#ffffff')
         .setFontSize(16)
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
    
    currentRow += 3;
    
    // Сводка
    const totalLeads = data.length;
    const totalSuccess = Object.values(successCount).reduce((sum, count) => sum + count, 0);
    const totalRefusals = Object.values(refusalCount).reduce((sum, count) => sum + count, 0);
    
    const summaryData = [
      ['Всего лидов:', totalLeads],
      ['Успешных (оплачено):', totalSuccess, (totalSuccess / totalLeads * 100).toFixed(1) + '%'],
      ['Отказов:', totalRefusals, (totalRefusals / totalLeads * 100).toFixed(1) + '%'],
      ['Других статусов:', totalLeads - totalSuccess - totalRefusals, ((totalLeads - totalSuccess - totalRefusals) / totalLeads * 100).toFixed(1) + '%']
    ];
    
    summaryData.forEach(row => {
      sheet.getRange(currentRow, 1).setValue(row[0]).setFontWeight('bold');
      sheet.getRange(currentRow, 2).setValue(row[1]);
      if (row[2]) {
        sheet.getRange(currentRow, 3).setValue(row[2]);
      }
      currentRow++;
    });
    
    currentRow += 2;
    
    // Все статусы
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue('ВСЕ СТАТУСЫ (отсортированы по количеству)')
         .setBackground('#3949ab')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    currentRow += 2;
    
    const headers = ['Статус', 'Количество', 'Процент', 'Тип', 'Учитывается как'];
    sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers])
         .setBackground('#7986cb')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    currentRow++;
    
    // Сортируем статусы по количеству
    const sortedStatuses = Object.entries(statusMap).sort((a, b) => b[1] - a[1]);
    
    const statusData = sortedStatuses.map(([status, count]) => {
      let type = 'Другое';
      let accounting = 'Не учитывается';
      let rowColor = '#ffffff';
      
      if (successCount[status]) {
        type = 'Успешный';
        accounting = 'Оплачено';
        rowColor = '#e8f5e9';
      } else if (refusalCount[status]) {
        type = 'Отказ';
        accounting = 'Потеря';
        rowColor = '#ffebee';
      }
      
      return {
        data: [
          status,
          count,
          (count / totalLeads * 100).toFixed(1) + '%',
          type,
          accounting
        ],
        color: rowColor
      };
    });
    
    // Записываем данные с цветовой индикацией
    statusData.forEach(({ data, color }) => {
      const range = sheet.getRange(currentRow, 1, 1, headers.length);
      range.setValues([data]);
      range.setBackground(color);
      currentRow++;
    });
    
    // Успешные статусы
    currentRow += 2;
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue('✅ УСПЕШНЫЕ СТАТУСЫ (гость оплатил)')
         .setBackground('#2e7d32')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    currentRow += 2;
    
    Object.entries(successCount)
      .sort((a, b) => b[1] - a[1])
      .forEach(([status, count]) => {
        sheet.getRange(currentRow, 1).setValue(status);
        sheet.getRange(currentRow, 2).setValue(count);
        sheet.getRange(currentRow, 3).setValue((count / totalLeads * 100).toFixed(1) + '%');
        sheet.getRange(currentRow, 1, 1, 3).setBackground('#e8f5e9');
        currentRow++;
      });
    
    // Статусы отказов
    currentRow += 2;
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue('❌ СТАТУСЫ ОТКАЗОВ')
         .setBackground('#d32f2f')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    currentRow += 2;
    
    Object.entries(refusalCount)
      .sort((a, b) => b[1] - a[1])
      .forEach(([status, count]) => {
        sheet.getRange(currentRow, 1).setValue(status);
        sheet.getRange(currentRow, 2).setValue(count);
        sheet.getRange(currentRow, 3).setValue((count / totalLeads * 100).toFixed(1) + '%');
        sheet.getRange(currentRow, 1, 1, 3).setBackground('#ffebee');
        currentRow++;
      });
    
    // Автоподгонка колонок
    sheet.autoResizeColumns(1, 5);
    
    console.log('✅ Анализ статусов завершен');
    console.log(`Найдено уникальных статусов: ${Object.keys(statusMap).length}`);
    console.log(`Успешных: ${Object.keys(successCount).length}`);
    console.log(`Отказов: ${Object.keys(refusalCount).length}`);
    
    return {
      allStatuses: statusMap,
      successStatuses: successCount,
      refusalStatuses: refusalCount
    };
    
  } catch (error) {
    console.error('❌ Ошибка анализа статусов:', error);
    throw error;
  }
}

/**
 * Быстрая проверка конкретного статуса
 */
function checkStatus(statusToCheck) {
  console.log(`Проверяем статус: "${statusToCheck}"`);
  console.log(`Успешный: ${isSuccessStatus(statusToCheck) ? 'ДА ✅' : 'НЕТ ❌'}`);
  console.log(`Отказ: ${isRefusalStatus(statusToCheck) ? 'ДА ❌' : 'НЕТ ✅'}`);
  console.log(`Оплачен: ${isPaidStatus(statusToCheck) ? 'ДА 💰' : 'НЕТ'}`);
}

// Примеры использования:
// checkStatus('Оплачено');
// checkStatus('Успешно РЕАЛИЗОВАНО');
// checkStatus('успешно в РП');
