/**
 * ТЕСТОВАЯ ФУНКЦИЯ ДЛЯ КРАСИВОГО ОФОРМЛЕНИЯ
 * Демонстрирует работу новой системы стилизации
 */

/**
 * Тестовая функция для проверки группировки столбцов в "РАБОЧИЙ АМО"
 * Показывает, как будут сгруппированы столбцы по блокам
 */
function testWorkingAmoColumnGrouping() {
  try {
    // Примерная структура заголовков из РАБОЧИЙ АМО
    const testHeaders = [
      'Сделка.ID', 'Сделка.Название', 'Контакт.ID', 'Контакт.Имя', 'Контакт.Телефон', 'Контакт.Email',
      'Сделка.Статус', 'Сделка.Этап', 'Сделка.Pipeline', 
      'Сделка.Дата создания', 'Сделка.Дата закрытия', 'TIME',
      'UTM_Source', 'UTM_Medium', 'UTM_Campaign', 'Источник ТЕЛ', 'Канал ТЕЛ',
      'Res.Визиты', 'Res.Сумма', 'Res.Статус',
      'Gue.Визиты', 'Gue.Общая сумма', 'Gue.Первый визит',
      'Время обработки'
    ];
    
    // Получаем группировку столбцов  
    const groups = getWorkingAmoColumnGroups_(testHeaders);
    
    console.log('🎨 ГРУППИРОВКА СТОЛБЦОВ для листа "РАБОЧИЙ АМО":');
    console.log('================================================');
    
    groups.forEach((group, index) => {
      const columns = testHeaders.slice(group.startIndex, group.endIndex + 1);
      console.log(`\n📦 БЛОК ${index + 1}: ${group.name}`);
      console.log(`🎨 Цвет: ${group.color}`);
      console.log(`📍 Столбцы: ${String.fromCharCode(65 + group.startIndex)}-${String.fromCharCode(65 + group.endIndex)}`);
      console.log(`📋 Содержание: ${columns.join(', ')}`);
    });
    
    console.log('\n✨ ВОЗМОЖНОСТИ ОФОРМЛЕНИЯ:');
    console.log('• Каждый блок имеет свой цвет фона');
    console.log('• Заголовки блоков выделены более ярким оттенком');
    console.log('• Строки чередуются между белым и светло-серым');
    console.log('• Блоки обведены рамками для четкого разделения');
    console.log('• Автоматический подбор ширины столбцов');
    console.log('• Заморозка строки заголовков');
    
    return {
      success: true,
      groups: groups,
      totalColumns: testHeaders.length,
      message: `Тест группировки завершен! Найдено ${groups.length} блоков из ${testHeaders.length} столбцов`
    };
    
  } catch (error) {
    console.error('❌ Ошибка тестирования группировки:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Показывает предварительный просмотр цветовой схемы
 */
function previewColorScheme() {
  try {
    const blockColors = [
      '#E3F2FD', // Светло-голубой
      '#F3E5F5', // Светло-фиолетовый  
      '#E8F5E8', // Светло-зеленый
      '#FFF3E0', // Светло-оранжевый
      '#FCE4EC', // Светло-розовый
      '#F1F8E9', // Светло-салатовый
      '#E0F2F1', // Светло-бирюзовый
      '#FFF8E1'  // Светло-желтый
    ];
    
    console.log('🎨 ЦВЕТОВАЯ СХЕМА для блоков:');
    console.log('============================');
    
    blockColors.forEach((color, index) => {
      const darkerColor = darkenColor_(color, 0.3);
      const lighterColor = lightenColor_(color, 0.8);
      
      console.log(`\n🎨 Блок ${index + 1}:`);
      console.log(`   Основной: ${color}`);
      console.log(`   Заголовок: ${darkerColor} (темнее на 30%)`);
      console.log(`   Данные: ${lighterColor} (светлее на 80%)`);
    });
    
    console.log('\n✨ ДОПОЛНИТЕЛЬНЫЕ ЦВЕТА:');
    console.log(`📄 Четные строки: #FFFFFF (белый)`);
    console.log(`📄 Нечетные строки: #F8F9FA (очень светло-серый)`);
    console.log(`🔲 Рамки блоков: #CCCCCC (светло-серый)`);
    console.log(`🔳 Рамки заголовков: #666666 (средне-серый)`);
    
  } catch (error) {
    console.error('❌ Ошибка предварительного просмотра:', error);
  }
}
