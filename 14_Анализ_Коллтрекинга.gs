/**
 * МОДУЛЬ: АНАЛИЗ СПОСОБОВ ОБРАЩЕНИЯ (ЗВОНКИ vs ЗАЯВКИ)
 * Разделяем анализ по способу первичного контакта
 */

function analyzeContactMethods() {
  console.log('📞 Анализируем способы обращения клиентов...');
  
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
    
    // Детальный анализ по каналам
    const channelAnalysis = {};
    
    // Анализ по номерам колл-трекинга
    const phoneAnalysis = {
      '78123172353': { name: 'Основной сайт', calls: 0, success: 0, revenue: 0, avgCheck: 0 },
      '78122428017': { name: 'Яндекс Карты', calls: 0, success: 0, revenue: 0, avgCheck: 0 },
      '78122709071': { name: 'Рестоклаб', calls: 0, success: 0, revenue: 0, avgCheck: 0 },
      '78123177149': { name: '2Гис', calls: 0, success: 0, revenue: 0, avgCheck: 0 },
      '78123177310': { name: 'Соц сети + Google', calls: 0, success: 0, revenue: 0, avgCheck: 0 }
    };
    
    // Анализ по UTM источникам
    const utmAnalysis = {};
    
    // Временные периоды
    const periodAnalysis = {
      beforeJuly2025: { total: 0, calls: 0, forms: 0, unknown: 0 },
      afterJuly2025: { total: 0, calls: 0, forms: 0, unknown: 0 }
    };
    
    // Анализируем каждую строку
    data.forEach(row => {
      stats.total++;
      
      const contactInfo = getContactMethodAndSource(row);
      const status = row[CONFIG.WORKING_AMO_COLUMNS.STATUS] || '';
      const factAmount = formatNumber(row[CONFIG.WORKING_AMO_COLUMNS.FACT_AMOUNT]);
      const createdAt = row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT];
      const isSuccess = isSuccessStatus(status);
      
      // Общая статистика по методам
      if (contactInfo.method === 'call') {
        stats.calls++;
        if (isSuccess) {
          stats.callsSuccess++;
          stats.callsRevenue += factAmount;
        }
      } else if (contactInfo.method === 'form') {
        stats.forms++;
        if (isSuccess) {
          stats.formsSuccess++;
          stats.formsRevenue += factAmount;
        }
      } else {
        stats.unknown++;
      }
      
      // Анализ по периодам
      const period = hasProperUtmTracking(createdAt) ? 'afterJuly2025' : 'beforeJuly2025';
      periodAnalysis[period].total++;
      if (contactInfo.method === 'call') periodAnalysis[period].calls++;
      else if (contactInfo.method === 'form') periodAnalysis[period].forms++;
      else periodAnalysis[period].unknown++;
      
      // Детальный анализ по каналам
      const channelKey = contactInfo.channel;
      if (!channelAnalysis[channelKey]) {
        channelAnalysis[channelKey] = {
          name: contactInfo.channel,
          type: contactInfo.type,
          calls: 0,
          forms: 0,
          total: 0,
          success: 0,
          revenue: 0,
          avgCheck: 0,
          conversion: 0
        };
      }
      
      channelAnalysis[channelKey].total++;
      if (contactInfo.method === 'call') channelAnalysis[channelKey].calls++;
      else if (contactInfo.method === 'form') channelAnalysis[channelKey].forms++;
      
      if (isSuccess) {
        channelAnalysis[channelKey].success++;
        channelAnalysis[channelKey].revenue += factAmount;
      }
      
      // Анализ телефонов колл-трекинга
      if (contactInfo.method === 'call') {
        const mangoLine = row[CONFIG.WORKING_AMO_COLUMNS.DEAL_MANGO_LINE] || '';
        const cleanPhone = mangoLine.toString().replace(/\D/g, '');
        if (phoneAnalysis[cleanPhone]) {
          phoneAnalysis[cleanPhone].calls++;
          if (isSuccess) {
            phoneAnalysis[cleanPhone].success++;
            phoneAnalysis[cleanPhone].revenue += factAmount;
          }
        }
      }
      
      // Анализ UTM источников
      if (contactInfo.method === 'form') {
        const utmSource = row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || 'no_utm';
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
 * Быстрая проверка источника для конкретной строки
 */
function testSourceDetection(rowNumber) {
  const data = getWorkingAmoData();
  if (rowNumber > data.length) {
    console.log('Строка не найдена');
    return;
  }
  
  const row = data[rowNumber - 1];
  const contactInfo = getContactMethodAndSource(row);
  
  console.log('=== АНАЛИЗ СТРОКИ', rowNumber, '===');
  console.log('MANGO линия:', row[CONFIG.WORKING_AMO_COLUMNS.DEAL_MANGO_LINE] || 'нет');
  console.log('UTM_SOURCE:', row[CONFIG.WORKING_AMO_COLUMNS.UTM_SOURCE] || 'нет');
  console.log('R.Источник сделки:', row[CONFIG.WORKING_AMO_COLUMNS.DEAL_SOURCE] || 'нет');
  console.log('Дата создания:', row[CONFIG.WORKING_AMO_COLUMNS.CREATED_AT]);
  console.log('---');
  console.log('Способ обращения:', contactInfo.method);
  console.log('Определенный источник:', contactInfo.name);
  console.log('Тип:', contactInfo.type);
  console.log('Канал:', contactInfo.channel);
  console.log('Код:', contactInfo.source);
}
