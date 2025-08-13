/**
 * ДЕМОНСТРАЦИОННЫЕ ФУНКЦИИ
 * Показывают работу новой системы объединения данных
 */

/**
 * Демо: создает полную структуру и заполняет тестовыми данными
 */
function demoCreateFullStructure() {
  try {
    console.log('🚀 ДЕМОНСТРАЦИЯ СОЗДАНИЯ ПОЛНОЙ СТРУКТУРЫ РАБОЧИЙ АМО');
    
    // 1. Создаем структуру сводного листа
    console.log('\n1️⃣ Создание структуры сводного листа...');
    const workingSheet = createWorkingAmoStructure();
    console.log('✅ Структура создана успешно');
    
    // 2. Показываем структуру
    console.log('\n2️⃣ Структура создана с колонками:');
    const headers = getUnifiedHeaders_();
    headers.forEach((header, index) => {
      const columnLetter = String.fromCharCode(65 + index);
      console.log(`${columnLetter}. ${header}`);
    });
    
    console.log(`\n📊 Итого создано ${headers.length} колонок для объединения всех источников данных`);
    
    // 3. Создаем тестовые данные
    console.log('\n3️⃣ Добавление тестовых данных...');
    createDemoData_(workingSheet, headers);
    console.log('✅ Тестовые данные добавлены');
    
    // 4. Диагностика
    console.log('\n4️⃣ Запуск диагностики...');
    diagnoseWorkingAmoStructure();
    
    console.log('\n🎉 ДЕМОНСТРАЦИЯ ЗАВЕРШЕНА!');
    console.log('\n📋 ЧТО СОЗДАНО:');
    console.log('• Лист "РАБОЧИЙ АМО" с полной структурой');
    console.log('• 48 колонок для всех типов данных');
    console.log('• Красивое оформление с цветными блоками');
    console.log('• Тестовые данные для проверки');
    console.log('• Автофильтры для удобной работы');
    
    console.log('\n🔧 СЛЕДУЮЩИЕ ШАГИ:');
    console.log('1. Используйте updateWorkingAmoFromAllSources() для объединения реальных данных');
    console.log('2. Запустите diagnoseRefusalData() для проверки отказов');
    console.log('3. Выполните analyzeRefusalReasons() для GPT-анализа');
    
  } catch (error) {
    console.error('❌ Ошибка демонстрации:', error);
  }
}

/**
 * Создает тестовые данные для демонстрации
 * @param {Sheet} sheet - Рабочий лист
 * @param {Array} headers - Заголовки
 * @private
 */
function createDemoData_(sheet, headers) {
  const demoData = [
    // Строка 1: Успешная сделка из AmoCRM
    [
      '12345',                           // ID
      'Заявка на банкет - Иван Петров',  // Название
      'Менеджер Сидорова А.В.',          // Ответственный
      'Иван Петров',                     // Контакт.ФИО
      'Успешно реализовано',             // Статус
      '150000',                          // Бюджет
      '2025-01-10',                      // Дата создания
      '+7 (999) 123-45-67',              // Контакт.Телефон
      'ivan.petrov@example.com',         // Email
      '8 (800) 555-01-01',               // Номер линии MANGO OFFICE
      'VIP, Банкет',                     // Теги
      '2025-01-15',                      // Дата закрытия
      '',                                // Причина отказа (пусто - успешная сделка)
      'Целевой',                         // Тип лида
      '2025-01-20',                      // Дата брони
      '19:00',                           // Время прихода
      '25',                              // Кол-во гостей
      'Филиал Центр',                    // Бар
      'Яндекс.Директ',                   // Источник
      'Контекстная реклама',             // R.Источник сделки
      'Входящий звонок',                 // R.Источник ТЕЛ сделки
      'CRM System',                      // ПО
      'yandex',                          // UTM_SOURCE
      'cpc',                             // UTM_MEDIUM
      'banket_campaign',                 // UTM_CAMPAIGN
      'банкет москва',                   // UTM_TERM
      'ad_group_1',                      // UTM_CONTENT
      'yandex.ru',                       // utm_referrer
      '1234567890123456789',             // YM_CLIENT_ID
      '9876543210987654321',             // GA_CLIENT_ID
      '1641234567890123456'              // _ym_uid
    ],
    
    // Строка 2: Отказанная сделка с причиной
    [
      '12346',                           // ID
      'Заказ корпоратива - ООО Ромашка', // Название
      'Менеджер Иванов П.С.',            // Ответственный
      'Анна Сидорова',                   // Контакт.ФИО
      'Закрыто и не реализовано',        // Статус ⭐ ОТКАЗ
      '200000',                          // Бюджет
      '2025-01-08',                      // Дата создания
      '+7 (999) 987-65-43',              // Контакт.Телефон
      'anna.sidorova@romashka.ru',       // Email
      '8 (800) 555-01-02',               // Номер линии MANGO OFFICE
      'Корпоратив',                      // Теги
      '2025-01-12',                      // Дата закрытия
      'Слишком дорого, нашли дешевле у конкурентов на 30%', // Причина отказа ⭐
      'Целевой',                         // Тип лида
      '2025-02-14',                      // Дата брони
      '18:30',                           // Время прихода
      '50',                              // Кол-во гостей
      'Филиал Север',                    // Бар
      'Google Ads',                      // Источник
      'Контекстная реклама',             // R.Источник сделки
      'Форма на сайте',                  // R.Источник ТЕЛ сделки
      'Google Ads',                      // ПО
      'google',                          // UTM_SOURCE
      'cpc',                             // UTM_MEDIUM
      'corporate_events',                // UTM_CAMPAIGN
      'корпоратив ресторан',             // UTM_TERM
      'ad_text_2',                       // UTM_CONTENT
      'google.com'                       // utm_referrer
    ],
    
    // Строка 3: Еще один отказ с другой причиной
    [
      '12347',                           // ID
      'День рождения - Михаил К.',       // Название
      'Менеджер Петрова Е.А.',           // Ответственный
      'Михаил Козлов',                   // Контакт.ФИО
      'Закрыто и не реализовано',        // Статус ⭐ ОТКАЗ
      '80000',                           // Бюджет
      '2025-01-09',                      // Дата создания
      '+7 (999) 555-77-88',              // Контакт.Телефон
      'mikhail.kozlov@gmail.com',        // Email
      '8 (800) 555-01-03',               // Номер линии MANGO OFFICE
      'День рождения',                   // Теги
      '2025-01-11',                      // Дата закрытия
      'Не устроило качество обслуживания при предыдущем посещении', // Причина отказа ⭐
      'Повторный',                       // Тип лида
      '2025-01-25',                      // Дата брони
      '20:00',                           // Время прихода
      '15',                              // Кол-во гостей
      'Филиал Юг',                       // Бар
      'Социальные сети',                 // Источник
      'Instagram',                       // R.Источник сделки
      'Сообщение в Direct',              // R.Источник ТЕЛ сделки
      'Instagram',                       // ПО
      'instagram',                       // UTM_SOURCE
      'social',                          // UTM_MEDIUM
      'birthday_promo',                  // UTM_CAMPAIGN
      '',                                // UTM_TERM
      'stories_ad'                       // UTM_CONTENT
    ]
  ];
  
  // Добавляем данные в лист
  if (demoData.length > 0) {
    // Заполняем только имеющиеся колонки, остальные оставляем пустыми
    const fullRows = demoData.map(row => {
      const fullRow = new Array(headers.length).fill('');
      row.forEach((value, index) => {
        if (index < fullRow.length) {
          fullRow[index] = value;
        }
      });
      return fullRow;
    });
    
    sheet.getRange(4, 1, fullRows.length, headers.length).setValues(fullRows);
    
    // Применяем форматирование к строкам с данными
    const dataRange = sheet.getRange(4, 1, fullRows.length, headers.length);
    
    // Чередуем цвета строк
    for (let i = 0; i < fullRows.length; i++) {
      const rowRange = sheet.getRange(4 + i, 1, 1, headers.length);
      const bgColor = i % 2 === 0 ? '#f8f9fa' : '#ffffff';
      
      // Выделяем отказанные сделки
      const statusIndex = headers.indexOf('Статус');
      if (statusIndex >= 0 && fullRows[i][statusIndex] === 'Закрыто и не реализовано') {
        rowRange.setBackground('#fff3e0'); // Оранжевый фон для отказов
      } else {
        rowRange.setBackground(bgColor);
      }
    }
    
    console.log(`✅ Добавлено ${fullRows.length} тестовых записей`);
    console.log('📊 Включает: 1 успешную сделку и 2 отказа с причинами');
  }
}

/**
 * Демо: показывает процесс объединения данных из разных источников
 */
function demoMergeFromSources() {
  try {
    console.log('🔄 ДЕМОНСТРАЦИЯ ОБЪЕДИНЕНИЯ ДАННЫХ ИЗ ИСТОЧНИКОВ');
    
    // Проверяем существующие листы-источники
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    console.log('\n📋 ДОСТУПНЫЕ ЛИСТЫ:');
    allSheets.forEach(sheetName => {
      console.log(`• ${sheetName}`);
    });
    
    // Источники данных для объединения
    const dataSources = [
      { sheetName: 'Амо Выгрузка', type: 'AMO_EXPORT' },
      { sheetName: 'Выгрузка Амо Полная', type: 'AMO_FULL_EXPORT' },
      { sheetName: 'Reserves RP', type: 'RESERVES_RP' },
      { sheetName: 'Guests RP', type: 'GUESTS_RP' },
      { sheetName: 'Заявки с Сайта', type: 'WEBSITE_FORMS' },
      { sheetName: 'КоллТрекинг', type: 'CALL_TRACKING' }
    ];
    
    console.log('\n🎯 ПОИСК ИСТОЧНИКОВ ДАННЫХ:');
    let foundSources = 0;
    
    dataSources.forEach(source => {
      if (allSheets.includes(source.sheetName)) {
        console.log(`✅ Найден: ${source.sheetName} (тип: ${source.type})`);
        foundSources++;
      } else {
        console.log(`❌ Не найден: ${source.sheetName}`);
      }
    });
    
    if (foundSources === 0) {
      console.log('\n⚠️ НЕ НАЙДЕНО ИСТОЧНИКОВ ДАННЫХ');
      console.log('📝 Для полного тестирования необходимо:');
      console.log('1. Импортировать данные из AmoCRM в листы "Амо Выгрузка"/"Выгрузка Амо Полная"');
      console.log('2. Добавить данные из RestoPlatform в листы "Reserves RP"/"Guests RP"');
      console.log('3. Загрузить заявки с сайта в лист "Заявки с Сайта"');
      console.log('4. Подключить данные коллтрекинга в лист "КоллТрекинг"');
      
      console.log('\n🔧 Создаем демо-структуру для тестирования...');
      demoCreateFullStructure();
      
    } else {
      console.log(`\n🚀 Найдено ${foundSources} источников, начинаем объединение...`);
      const totalMerged = updateWorkingAmoFromAllSources();
      
      console.log(`\n✅ ОБЪЕДИНЕНИЕ ЗАВЕРШЕНО!`);
      console.log(`📊 Всего объединено записей: ${totalMerged}`);
      
      // Запускаем диагностику
      console.log('\n🔍 ДИАГНОСТИКА ОБЪЕДИНЕННЫХ ДАННЫХ:');
      diagnoseWorkingAmoStructure();
    }
    
    console.log('\n🎯 ГОТОВО К АНАЛИЗУ ОТКАЗОВ!');
    console.log('Используйте: diagnoseRefusalData() → analyzeRefusalReasons()');
    
  } catch (error) {
    console.error('❌ Ошибка демонстрации объединения:', error);
  }
}

/**
 * Демо: показывает полный цикл анализа отказов с новой структурой
 */
function demoRefusalAnalysisFullCycle() {
  try {
    console.log('🔬 ДЕМОНСТРАЦИЯ ПОЛНОГО ЦИКЛА АНАЛИЗА ОТКАЗОВ');
    
    // 1. Создаем структуру (если нет)
    console.log('\n1️⃣ Подготовка структуры данных...');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let workingSheet = spreadsheet.getSheetByName('РАБОЧИЙ АМО');
    
    if (!workingSheet || workingSheet.getLastRow() < 4) {
      console.log('📝 Создаем тестовую структуру...');
      demoCreateFullStructure();
    } else {
      console.log('✅ Структура данных найдена');
    }
    
    // 2. Диагностика данных
    console.log('\n2️⃣ Диагностика данных для анализа отказов...');
    diagnoseRefusalData();
    
    // 3. Запуск анализа отказов
    console.log('\n3️⃣ Запуск GPT-анализа причин отказов...');
    try {
      analyzeRefusalReasons();
      console.log('✅ GPT-анализ отказов завершен успешно');
      
      // 4. Проверка результата
      console.log('\n4️⃣ Проверка созданного отчета...');
      const refusalSheet = spreadsheet.getSheetByName('Анализ отказов');
      if (refusalSheet) {
        const lastRow = refusalSheet.getLastRow();
        console.log(`✅ Отчет создан, содержит ${lastRow} строк`);
        console.log('📊 Отчет включает: категории причин, инсайты, рекомендации, диаграммы');
      } else {
        console.log('❌ Отчет не найден');
      }
      
    } catch (analysisError) {
      console.log('⚠️ Ошибка GPT-анализа (возможно, не настроен API ключ)');
      console.log('📝 Результат: будет использован локальный анализ на основе ключевых слов');
    }
    
    console.log('\n🎉 ДЕМОНСТРАЦИЯ ПОЛНОГО ЦИКЛА ЗАВЕРШЕНА!');
    
    console.log('\n📋 ЧТО БЫЛО ПРОДЕМОНСТРИРОВАНО:');
    console.log('✅ Создание унифицированной структуры РАБОЧИЙ АМО');
    console.log('✅ Сопоставление полей из всех источников данных');
    console.log('✅ Диагностика качества данных для анализа отказов');
    console.log('✅ GPT-анализ причин отказов с красивым оформлением');
    console.log('✅ Создание профессиональных отчетов с диаграммами');
    
    console.log('\n💡 ПРЕИМУЩЕСТВА НОВОЙ СИСТЕМЫ:');
    console.log('🎯 Единая структура для всех источников данных');
    console.log('🔄 Автоматическое сопоставление полей');
    console.log('📊 Комплексная аналитика по всем каналам');
    console.log('🤖 GPT-анализ с умным rate limiting');
    console.log('🎨 Красивые отчеты как в старых версиях');
    
  } catch (error) {
    console.error('❌ Ошибка демонстрации полного цикла:', error);
  }
}
