/**
 * МОДУЛЬ: АНАЛИЗ СПОСОБОВ ОБРАЩЕНИЯ (ЗВОНКИ vs ЗАЯВКИ)
 * Разделяем анализ по способу первичного контакта
 */

function analyzeContactMethods() {
  console.log('📞 ИСПРАВЛЕННЫЙ АНАЛИЗ: Способы обращения клиентов...');
  
  try {
    const data = getWorkingAmoData();
    if (!data || data.length === 0) {
      console.log('Нет данных для анализа');
      return;
    }
    
    // Структуры для анализа
    const stats = {
      total: 0,
      calls: 0,
      forms: 0,
      unknown: 0,
      callsSuccess: 0,
      formsSuccess: 0,
      callsRevenue: 0,
      formsRevenue: 0
    };
    
    // ИСПРАВЛЕНО: Анализ по номерам из колонки M (индекс 12)
    const phoneAnalysis = {
      '78123172353': { name: 'Основной сайт', calls: 0, success: 0, revenue: 0, avgCheck: 0 },
      '78122428017': { name: 'Яндекс Карты', calls: 0, success: 0, revenue: 0, avgCheck: 0 },
      '78122709071': { name: 'Рестоклаб', calls: 0, success: 0, revenue: 0, avgCheck: 0 },
      '78123177149': { name: '2Гис', calls: 0, success: 0, revenue: 0, avgCheck: 0 },
      '78123177310': { name: 'Соц сети + Google', calls: 0, success: 0, revenue: 0, avgCheck: 0 }
    };
    
    // Детальный анализ по каналам
    const channelAnalysis = {};
    
    // Анализ по UTM источникам
    const utmAnalysis = {};
    
    // Временные периоды
    const periodAnalysis = {
      beforeJuly2025: { total: 0, calls: 0, forms: 0, unknown: 0 },
      afterJuly2025: { total: 0, calls: 0, forms: 0, unknown: 0 }
    };
    
    console.log(`📊 Обрабатываем ${data.length} строк данных...`);
    
    // ИСПРАВЛЕНО: Анализируем каждую строку по правильной логике
    data.forEach((row, index) => {
      stats.total++;
      
      // ИСПРАВЛЕНО: Читаем прямо по индексам колонок
      const mangoLine = row[12] || '';                    // Колонка M - Сделка.Номер линии MANGO OFFICE
      const status = row[2] || '';                        // Колонка C - Сделка.Статус
      const factAmount = formatNumber(row[15] || 0);      // Колонка P - Счет факт
      const utmSource = (row[27] || '').toString().trim(); // Колонка AB - UTM_SOURCE
      const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
      const isSuccess = isSuccessStatus(status);
      
      // ИСПРАВЛЕНО: Определяем тип по четкой логике
      const hasMangoLine = mangoLine && mangoLine.toString().trim() !== '';
      const hasUtm = utmSource && utmSource !== '';
      
      let contactMethod = 'unknown';
      let contactChannel = 'Неизвестный';
      
      if (hasMangoLine) {
        // Это ЗВОНОК (есть номер в колонке M)
        contactMethod = 'call';
        stats.calls++;
        
        if (isSuccess) {
          stats.callsSuccess++;
          stats.callsRevenue += factAmount;
        }
        
        // Анализ конкретного номера
        const cleanPhone = mangoLine.toString().replace(/\D/g, '');
        if (phoneAnalysis[cleanPhone]) {
          contactChannel = phoneAnalysis[cleanPhone].name;
          phoneAnalysis[cleanPhone].calls++;
          if (isSuccess) {
            phoneAnalysis[cleanPhone].success++;
            phoneAnalysis[cleanPhone].revenue += factAmount;
          }
        } else {
          contactChannel = `Неизвестный номер (${mangoLine})`;
        }
      } 
      else if (hasUtm) {
        // Это ЗАЯВКА (есть UTM, но нет MANGO)
        contactMethod = 'form';
        stats.forms++;
        contactChannel = utmSource;
        
        if (isSuccess) {
          stats.formsSuccess++;
          stats.formsRevenue += factAmount;
        }
        
        // Анализ UTM источников
        if (!utmAnalysis[utmSource]) {
          utmAnalysis[utmSource] = {
            forms: 0,
            success: 0,
            revenue: 0,
            avgCheck: 0,
            conversion: 0
          };
        }
        utmAnalysis[utmSource].forms++;
        if (isSuccess) {
          utmAnalysis[utmSource].success++;
          utmAnalysis[utmSource].revenue += factAmount;
        }
      } 
      else {
        // Неопределенный тип
        stats.unknown++;
        contactMethod = 'unknown';
        contactChannel = 'Неопределенный';
      }
      
      // Анализ по периодам
      const period = hasProperUtmTracking(createdAt) ? 'afterJuly2025' : 'beforeJuly2025';
      periodAnalysis[period].total++;
      if (contactMethod === 'call') periodAnalysis[period].calls++;
      else if (contactMethod === 'form') periodAnalysis[period].forms++;
      else periodAnalysis[period].unknown++;
      
      // Детальный анализ по каналам
      if (!channelAnalysis[contactChannel]) {
        channelAnalysis[contactChannel] = {
          name: contactChannel,
          type: contactMethod === 'call' ? 'Звонки' : contactMethod === 'form' ? 'Заявки' : 'Неизвестный',
          calls: 0,
          forms: 0,
          total: 0,
          success: 0,
          revenue: 0,
          avgCheck: 0,
          conversion: 0
        };
      }
      
      channelAnalysis[contactChannel].total++;
      if (contactMethod === 'call') channelAnalysis[contactChannel].calls++;
      else if (contactMethod === 'form') channelAnalysis[contactChannel].forms++;
      
      if (isSuccess) {
        channelAnalysis[contactChannel].success++;
        channelAnalysis[contactChannel].revenue += factAmount;
      }
    });
    
    // Рассчитываем производные метрики
    Object.values(channelAnalysis).forEach(channel => {
      channel.avgCheck = channel.success > 0 ? channel.revenue / channel.success : 0;
      channel.conversion = channel.total > 0 ? (channel.success / channel.total * 100) : 0;
    });
    
    Object.values(phoneAnalysis).forEach(phone => {
      phone.avgCheck = phone.success > 0 ? phone.revenue / phone.success : 0;
    });
    
    Object.values(utmAnalysis).forEach(utm => {
      utm.avgCheck = utm.success > 0 ? utm.revenue / utm.success : 0;
      utm.conversion = utm.forms > 0 ? (utm.success / utm.forms * 100) : 0;
    });
    
    // ИСПРАВЛЕННЫЙ ВЫВОД СТАТИСТИКИ
    console.log('');
    console.log('📊 ИСПРАВЛЕННАЯ СТАТИСТИКА ЗВОНКОВ:');
    console.log('=====================================');
    console.log(`📞 Всего звонков (колонка М заполнена): ${stats.calls}`);
    console.log(`✅ Успешных звонков: ${stats.callsSuccess}`);
    console.log(`💰 Выручка от звонков: ${formatCurrency(stats.callsRevenue)}`);
    console.log(`📈 Конверсия звонков: ${stats.calls > 0 ? (stats.callsSuccess / stats.calls * 100).toFixed(1) : 0}%`);
    console.log('');
    console.log(`📝 Всего заявок (UTM без MANGO): ${stats.forms}`);
    console.log(`✅ Успешных заявок: ${stats.formsSuccess}`);
    console.log(`💰 Выручка от заявок: ${formatCurrency(stats.formsRevenue)}`);
    console.log(`📈 Конверсия заявок: ${stats.forms > 0 ? (stats.formsSuccess / stats.forms * 100).toFixed(1) : 0}%`);
    console.log('');
    console.log(`❓ Неопределенных записей: ${stats.unknown}`);
    console.log(`📊 Всего обработано: ${stats.total}`);
    console.log('=====================================');
    
    // Детальная статистика по телефонам
    console.log('');
    console.log('📞 ДЕТАЛЬНАЯ СТАТИСТИКА ПО НОМЕРАМ:');
    Object.entries(phoneAnalysis).forEach(([phone, data]) => {
      if (data.calls > 0) {
        console.log(`${phone} (${data.name}): ${data.calls} звонков, ${data.success} успешных`);
      }
    });
    
    // Создаем красивый отчет
    const sheet = createOrUpdateSheet('📞 ЗВОНКИ vs ЗАЯВКИ');
    let currentRow = 1;
    
    // Заголовок
    const titleRange = sheet.getRange(currentRow, 1, 1, 10);
    titleRange.merge();
    titleRange.setValue('📞 АНАЛИЗ СПОСОБОВ ОБРАЩЕНИЯ: ЗВОНКИ vs ЗАЯВКИ');
    titleRange.setBackground('#1a237e')
              .setFontColor('#ffffff')
              .setFontSize(18)
              .setFontWeight('bold')
              .setHorizontalAlignment('center');
    
    currentRow += 3;
    
    // ОБЩАЯ СТАТИСТИКА
    sheet.getRange(currentRow, 1, 1, 8).merge();
    sheet.getRange(currentRow, 1).setValue('📊 ОБЩАЯ СТАТИСТИКА')
         .setBackground('#283593')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setFontSize(14);
    
    currentRow += 2;
    
    // KPI карточки
    const kpiData = [
      ['📞 ЗВОНКИ', stats.calls, '📝 ЗАЯВКИ', stats.forms],
      ['Конверсия звонков', stats.calls > 0 ? (stats.callsSuccess / stats.calls * 100).toFixed(1) + '%' : '0%', 
       'Конверсия заявок', stats.forms > 0 ? (stats.formsSuccess / stats.forms * 100).toFixed(1) + '%' : '0%'],
      ['Выручка от звонков', formatCurrency(stats.callsRevenue), 
       'Выручка от заявок', formatCurrency(stats.formsRevenue)],
      ['Средний чек (звонки)', formatCurrency(stats.callsSuccess > 0 ? stats.callsRevenue / stats.callsSuccess : 0),
       'Средний чек (заявки)', formatCurrency(stats.formsSuccess > 0 ? stats.formsRevenue / stats.formsSuccess : 0)]
    ];
    
    kpiData.forEach((row, index) => {
      const rowNum = currentRow + index;
      
      // Звонки
      sheet.getRange(rowNum, 1, 1, 2).merge();
      sheet.getRange(rowNum, 1).setValue(row[0])
           .setBackground('#bbdefb')
           .setFontWeight('bold');
      
      sheet.getRange(rowNum, 3, 1, 2).merge();
      sheet.getRange(rowNum, 3).setValue(row[1])
           .setBackground('#e3f2fd')
           .setFontSize(14)
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
      
      // Заявки
      sheet.getRange(rowNum, 6, 1, 2).merge();
      sheet.getRange(rowNum, 6).setValue(row[2])
           .setBackground('#c8e6c9')
           .setFontWeight('bold');
      
      sheet.getRange(rowNum, 8, 1, 2).merge();
      sheet.getRange(rowNum, 8).setValue(row[3])
           .setBackground('#e8f5e9')
           .setFontSize(14)
           .setFontWeight('bold')
           .setHorizontalAlignment('center');
    });
    
    currentRow += kpiData.length + 2;
    
    // СРАВНЕНИЕ ПО ПЕРИОДАМ
    sheet.getRange(currentRow, 1, 1, 10).merge();
    sheet.getRange(currentRow, 1).setValue('📅 СРАВНЕНИЕ ПО ПЕРИОДАМ')
         .setBackground('#00897b')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setFontSize(14);
    
    currentRow += 2;
    
    const periodHeaders = ['Период', 'Всего', 'Звонки', '% звонков', 'Заявки', '% заявок', 'Неизвестно'];
    sheet.getRange(currentRow, 1, 1, periodHeaders.length).setValues([periodHeaders])
         .setBackground('#4db6ac')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    currentRow++;
    
    const periodData = [
      ['До июля 2025 (старая система)', 
       periodAnalysis.beforeJuly2025.total,
       periodAnalysis.beforeJuly2025.calls,
       periodAnalysis.beforeJuly2025.total > 0 ? 
         (periodAnalysis.beforeJuly2025.calls / periodAnalysis.beforeJuly2025.total * 100).toFixed(1) + '%' : '0%',
       periodAnalysis.beforeJuly2025.forms,
       periodAnalysis.beforeJuly2025.total > 0 ? 
         (periodAnalysis.beforeJuly2025.forms / periodAnalysis.beforeJuly2025.total * 100).toFixed(1) + '%' : '0%',
       periodAnalysis.beforeJuly2025.unknown],
      ['С июля 2025 (новая UTM разметка)', 
       periodAnalysis.afterJuly2025.total,
       periodAnalysis.afterJuly2025.calls,
       periodAnalysis.afterJuly2025.total > 0 ? 
         (periodAnalysis.afterJuly2025.calls / periodAnalysis.afterJuly2025.total * 100).toFixed(1) + '%' : '0%',
       periodAnalysis.afterJuly2025.forms,
       periodAnalysis.afterJuly2025.total > 0 ? 
         (periodAnalysis.afterJuly2025.forms / periodAnalysis.afterJuly2025.total * 100).toFixed(1) + '%' : '0%',
       periodAnalysis.afterJuly2025.unknown]
    ];
    
    sheet.getRange(currentRow, 1, periodData.length, periodHeaders.length).setValues(periodData);
    currentRow += periodData.length + 2;
    
    // АНАЛИЗ ТЕЛЕФОНОВ КОЛЛ-ТРЕКИНГА
    sheet.getRange(currentRow, 1, 1, 10).merge();
    sheet.getRange(currentRow, 1).setValue('☎️ ДЕТАЛЬНЫЙ АНАЛИЗ ТЕЛЕФОНОВ КОЛЛ-ТРЕКИНГА')
         .setBackground('#1976d2')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setFontSize(14);
    
    currentRow += 2;
    
    const phoneHeaders = ['Номер', 'Источник', 'Звонков', 'Успешных', 'Конверсия', 'Выручка', 'Средний чек'];
    sheet.getRange(currentRow, 1, 1, phoneHeaders.length).setValues([phoneHeaders])
         .setBackground('#42a5f5')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    currentRow++;
    
    const phoneData = Object.entries(phoneAnalysis).map(([phone, data]) => [
      phone,
      data.name,
      data.calls,
      data.success,
      data.calls > 0 ? (data.success / data.calls * 100).toFixed(1) + '%' : '0%',
      formatCurrency(data.revenue),
      formatCurrency(data.avgCheck)
    ]);
    
    sheet.getRange(currentRow, 1, phoneData.length, phoneHeaders.length).setValues(phoneData);
    currentRow += phoneData.length + 2;
    
    // ТОП КАНАЛОВ ПО ЭФФЕКТИВНОСТИ
    sheet.getRange(currentRow, 1, 1, 10).merge();
    sheet.getRange(currentRow, 1).setValue('🏆 ТОП КАНАЛОВ ПО ЭФФЕКТИВНОСТИ')
         .setBackground('#d32f2f')
         .setFontColor('#ffffff')
         .setFontWeight('bold')
         .setFontSize(14);
    
    currentRow += 2;
    
    const channelHeaders = ['Канал', 'Тип', 'Всего', 'Звонки', 'Заявки', 'Успешных', 'Конверсия', 'Выручка', 'Ср. чек'];
    sheet.getRange(currentRow, 1, 1, channelHeaders.length).setValues([channelHeaders])
         .setBackground('#ef5350')
         .setFontColor('#ffffff')
         .setFontWeight('bold');
    
    currentRow++;
    
    const sortedChannels = Object.values(channelAnalysis)
      .sort((a, b) => b.revenue - a.revenue)
      .slice(0, 20);
    
    const channelData = sortedChannels.map(channel => [
      channel.name,
      channel.type,
      channel.total,
      channel.calls,
      channel.forms,
      channel.success,
      channel.conversion.toFixed(1) + '%',
      formatCurrency(channel.revenue),
      formatCurrency(channel.avgCheck)
    ]);
    
    sheet.getRange(currentRow, 1, channelData.length, channelHeaders.length).setValues(channelData);
    
    // Выделяем топ-3 канала
    for (let i = 0; i < Math.min(3, channelData.length); i++) {
      const rowRange = sheet.getRange(currentRow + i, 1, 1, channelHeaders.length);
      rowRange.setBackground(i === 0 ? '#fff9c4' : i === 1 ? '#f0f4c3' : '#e6ee9c');
    }
    
    // Автоподгонка колонок
    sheet.autoResizeColumns(1, 10);
    
    console.log('✅ Анализ способов обращения завершен');
    console.log(`Всего: ${stats.total} | Звонки: ${stats.calls} | Заявки: ${stats.forms}`);
    
  } catch (error) {
    console.error('❌ Ошибка анализа:', error);
    throw error;
  }
}

/**
 * ИСПРАВЛЕННАЯ функция для проверки источника конкретной строки
 */
function testSourceDetection(rowNumber) {
  const data = getWorkingAmoData();
  if (!data || rowNumber > data.length) {
    console.log('❌ Строка не найдена или нет данных');
    return;
  }
  
  const row = data[rowNumber - 1];
  
  console.log('');
  console.log(`🔍 ИСПРАВЛЕННЫЙ АНАЛИЗ СТРОКИ ${rowNumber}:`);
  console.log('=============================================');
  console.log(`ID сделки (A): "${row[0] || 'пусто'}"`);
  console.log(`Название (B): "${row[1] || 'пусто'}"`);
  console.log(`Статус (C): "${row[2] || 'пусто'}"`);
  console.log(`MANGO линия (М, индекс 12): "${row[12] || 'пусто'}"`);
  console.log(`UTM_SOURCE (AB, индекс 27): "${row[27] || 'пусто'}"`);
  console.log(`Счет факт (P, индекс 15): "${row[15] || 'пусто'}"`);
  console.log('');
  
  // ИСПРАВЛЕННАЯ логика определения
  const mangoLine = row[12] || '';
  const utmSource = row[27] || '';
  const status = row[2] || '';
  
  const hasMangoLine = mangoLine && mangoLine.toString().trim() !== '';
  const hasUtm = utmSource && utmSource.toString().trim() !== '';
  const isSuccess = isSuccessStatus(status);
  
  console.log('📊 РЕЗУЛЬТАТ ИСПРАВЛЕННОГО АНАЛИЗА:');
  console.log('-----------------------------------');
  
  if (hasMangoLine) {
    console.log(`✅ ТИП: ЗВОНОК (найден номер в колонке М)`);
    console.log(`📞 Номер: ${mangoLine}`);
    const cleanPhone = mangoLine.toString().replace(/\D/g, '');
    console.log(`🧹 Очищенный номер: ${cleanPhone}`);
    
    // Проверяем в маппинге
    const mapping = {
      '78123172353': 'Основной сайт',
      '78122428017': 'Яндекс Карты', 
      '78122709071': 'Рестоклаб',
      '78123177149': '2Гис',
      '78123177310': 'Соц сети + Google'
    };
    
    if (mapping[cleanPhone]) {
      console.log(`🎯 Источник: ${mapping[cleanPhone]}`);
    } else {
      console.log(`❓ Неизвестный номер в маппинге`);
    }
  } 
  else if (hasUtm) {
    console.log(`✅ ТИП: ЗАЯВКА (найден UTM, но нет MANGO)`);
    console.log(`🌐 UTM: ${utmSource}`);
  } 
  else {
    console.log(`❓ ТИП: НЕОПРЕДЕЛЕННЫЙ (нет ни MANGO, ни UTM)`);
  }
  
  console.log(`💰 Статус успешный: ${isSuccess ? 'ДА' : 'НЕТ'}`);
  console.log(`📈 Засчитается в статистике: ${hasMangoLine ? 'ЗВОНКИ' : hasUtm ? 'ЗАЯВКИ' : 'НЕ ЗАСЧИТАЕТСЯ'}`);
  console.log('=============================================');
  
  return {
    rowNumber,
    mangoLine,
    utmSource,
    hasMangoLine,
    hasUtm,
    isSuccess,
    type: hasMangoLine ? 'call' : hasUtm ? 'form' : 'unknown'
  };
}
