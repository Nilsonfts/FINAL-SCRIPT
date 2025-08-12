/**
 * 📅 Утилиты Дат - Работа с датами
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Apps Script: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit
 * 2. Создайте новый файл: + → Script file
 * 3. Назовите файл: 04_Утилиты_Дат (без .gs)
 * 4. Удалите весь код по умолчанию
 * 5. Скопируйте и вставьте код ниже
 * 6. Сохраните файл (Ctrl+S)
 */

/**
 * УТИЛИТЫ ДЛЯ РАБОТЫ С ДАТАМИ И ВРЕМЕНЕМ
 * Работа с московским часовым поясом, периодами отчётности, календарными расчётами
 * @fileoverview Управление датами, периодами и временными интервалами
 */

/**
 * Получает текущую дату в московском часовом поясе
 * @returns {Date} Текущая дата MSK
 */
function getCurrentDateMoscow_() {
  return new Date(new Date().toLocaleString("en-US", { timeZone: "Europe/Moscow" }));
}

/**
 * Преобразует дату в московский часовой пояс
 * @param {Date|string} date - Исходная дата
 * @returns {Date} Дата в MSK
 */
function convertToMoscowTime_(date) {
  if (!date) return null;
  
  let dateObj;
  if (typeof date === 'string') {
    dateObj = new Date(date);
  } else if (date instanceof Date) {
    dateObj = new Date(date);
  } else {
    return null;
  }
  
  if (isNaN(dateObj.getTime())) return null;
  
  return new Date(dateObj.toLocaleString("en-US", { timeZone: "Europe/Moscow" }));
}

/**
 * Форматирует дату в строку в заданном формате
 * @param {Date} date - Дата для форматирования
 * @param {string} format - Формат ('YYYY-MM-DD', 'DD.MM.YYYY', 'DD/MM/YYYY', 'full')
 * @returns {string} Отформатированная дата
 */
function formatDate_(date, format = 'DD.MM.YYYY') {
  if (!date || isNaN(date.getTime())) return '';
  
  const moscowDate = convertToMoscowTime_(date);
  const day = String(moscowDate.getDate()).padStart(2, '0');
  const month = String(moscowDate.getMonth() + 1).padStart(2, '0');
  const year = moscowDate.getFullYear();
  const hours = String(moscowDate.getHours()).padStart(2, '0');
  const minutes = String(moscowDate.getMinutes()).padStart(2, '0');
  
  switch (format.toLowerCase()) {
    case 'yyyy-mm-dd':
      return `${year}-${month}-${day}`;
    case 'dd.mm.yyyy':
      return `${day}.${month}.${year}`;
    case 'dd/mm/yyyy':
      return `${day}/${month}/${year}`;
    case 'mm/dd/yyyy':
      return `${month}/${day}/${year}`;
    case 'full':
      return `${day}.${month}.${year} ${hours}:${minutes}`;
    case 'iso':
      return moscowDate.toISOString();
    case 'ru':
      const months = [
        'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
        'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
      ];
      return `${day} ${months[moscowDate.getMonth()]} ${year} г.`;
    default:
      return `${day}.${month}.${year}`;
  }
}

/**
 * Парсит строку даты в различных форматах
 * @param {string} dateString - Строка с датой
 * @returns {Date|null} Распарсенная дата или null
 */
function parseDate_(dateString) {
  if (!dateString) return null;
  
  const str = String(dateString).trim();
  if (!str) return null;
  
  // Попробуем стандартные форматы
  let date;
  
  // ISO формат (2024-01-15T10:30:00Z)
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/.test(str)) {
    date = new Date(str);
    if (!isNaN(date.getTime())) {
      return convertToMoscowTime_(date);
    }
  }
  
  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    date = new Date(str + 'T00:00:00');
    if (!isNaN(date.getTime())) {
      return convertToMoscowTime_(date);
    }
  }
  
  // DD.MM.YYYY или DD/MM/YYYY
  const ddmmMatch = str.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
  if (ddmmMatch) {
    const [, day, month, year] = ddmmMatch;
    date = new Date(year, month - 1, day);
    if (!isNaN(date.getTime())) {
      return convertToMoscowTime_(date);
    }
  }
  
  // DD.MM.YYYY HH:MM
  const ddmmTimeMatch = str.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})\s+(\d{1,2}):(\d{2})$/);
  if (ddmmTimeMatch) {
    const [, day, month, year, hour, minute] = ddmmTimeMatch;
    date = new Date(year, month - 1, day, hour, minute);
    if (!isNaN(date.getTime())) {
      return convertToMoscowTime_(date);
    }
  }
  
  // Попробуем встроенный парсер как последний вариант
  date = new Date(str);
  if (!isNaN(date.getTime())) {
    return convertToMoscowTime_(date);
  }
  
  return null;
}

/**
 * Получает начало дня для даты (00:00:00)
 * @param {Date} date - Исходная дата
 * @returns {Date} Начало дня
 */
function getStartOfDay_(date) {
  if (!date) return null;
  
  const moscowDate = convertToMoscowTime_(date);
  return new Date(moscowDate.getFullYear(), moscowDate.getMonth(), moscowDate.getDate());
}

/**
 * Получает конец дня для даты (23:59:59)
 * @param {Date} date - Исходная дата
 * @returns {Date} Конец дня
 */
function getEndOfDay_(date) {
  if (!date) return null;
  
  const moscowDate = convertToMoscowTime_(date);
  return new Date(moscowDate.getFullYear(), moscowDate.getMonth(), moscowDate.getDate(), 23, 59, 59, 999);
}

/**
 * Получает первый день месяца
 * @param {Date} date - Дата в месяце
 * @returns {Date} Первый день месяца
 */
function getStartOfMonth_(date) {
  if (!date) return null;
  
  const moscowDate = convertToMoscowTime_(date);
  return new Date(moscowDate.getFullYear(), moscowDate.getMonth(), 1);
}

/**
 * Получает последний день месяца
 * @param {Date} date - Дата в месяце
 * @returns {Date} Последний день месяца
 */
function getEndOfMonth_(date) {
  if (!date) return null;
  
  const moscowDate = convertToMoscowTime_(date);
  return new Date(moscowDate.getFullYear(), moscowDate.getMonth() + 1, 0, 23, 59, 59, 999);
}

/**
 * Получает первый день года
 * @param {Date} date - Дата в году
 * @returns {Date} Первый день года
 */
function getStartOfYear_(date) {
  if (!date) return null;
  
  const moscowDate = convertToMoscowTime_(date);
  return new Date(moscowDate.getFullYear(), 0, 1);
}

/**
 * Получает последний день года
 * @param {Date} date - Дата в году
 * @returns {Date} Последний день года
 */
function getEndOfYear_(date) {
  if (!date) return null;
  
  const moscowDate = convertToMoscowTime_(date);
  return new Date(moscowDate.getFullYear(), 11, 31, 23, 59, 59, 999);
}

/**
 * Добавляет дни к дате
 * @param {Date} date - Исходная дата
 * @param {number} days - Количество дней для добавления
 * @returns {Date} Новая дата
 */
function addDays_(date, days) {
  if (!date || isNaN(days)) return null;
  
  const result = new Date(convertToMoscowTime_(date));
  result.setDate(result.getDate() + days);
  return convertToMoscowTime_(result);
}

/**
 * Добавляет месяцы к дате
 * @param {Date} date - Исходная дата
 * @param {number} months - Количество месяцев для добавления
 * @returns {Date} Новая дата
 */
function addMonths_(date, months) {
  if (!date || isNaN(months)) return null;
  
  const result = new Date(convertToMoscowTime_(date));
  result.setMonth(result.getMonth() + months);
  return convertToMoscowTime_(result);
}

/**
 * Вычисляет разность между датами в днях
 * @param {Date} date1 - Первая дата
 * @param {Date} date2 - Вторая дата
 * @returns {number} Разность в днях (date1 - date2)
 */
function getDaysDifference_(date1, date2) {
  if (!date1 || !date2) return 0;
  
  const moscow1 = convertToMoscowTime_(date1);
  const moscow2 = convertToMoscowTime_(date2);
  
  const diffTime = moscow1.getTime() - moscow2.getTime();
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

/**
 * Получает список всех месяцев в заданном периоде
 * @param {Date} startDate - Начальная дата
 * @param {Date} endDate - Конечная дата
 * @returns {Array<Object>} Массив с информацией о месяцах
 */
function getMonthsInPeriod_(startDate, endDate) {
  if (!startDate || !endDate) return [];
  
  const months = [];
  let current = getStartOfMonth_(startDate);
  const end = getEndOfMonth_(endDate);
  
  while (current <= end) {
    const monthEnd = getEndOfMonth_(current);
    
    months.push({
      start: new Date(current),
      end: new Date(monthEnd),
      year: current.getFullYear(),
      month: current.getMonth() + 1,
      monthName: getMonthName_(current.getMonth()),
      key: `${current.getFullYear()}-${String(current.getMonth() + 1).padStart(2, '0')}`
    });
    
    current = addMonths_(current, 1);
  }
  
  return months;
}

/**
 * Получает название месяца на русском языке
 * @param {number} monthIndex - Индекс месяца (0-11)
 * @returns {string} Название месяца
 */
function getMonthName_(monthIndex) {
  const months = [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
  ];
  
  return months[monthIndex] || 'Неизвестно';
}

/**
 * Определяет, является ли дата рабочим днём
 * @param {Date} date - Проверяемая дата
 * @returns {boolean} true если рабочий день
 */
function isWorkingDay_(date) {
  if (!date) return false;
  
  const moscowDate = convertToMoscowTime_(date);
  const dayOfWeek = moscowDate.getDay();
  
  // 0 = воскресенье, 6 = суббота
  return dayOfWeek !== 0 && dayOfWeek !== 6;
}

/**
 * Получает количество рабочих дней в периоде
 * @param {Date} startDate - Начальная дата
 * @param {Date} endDate - Конечная дата
 * @returns {number} Количество рабочих дней
 */
function getWorkingDaysCount_(startDate, endDate) {
  if (!startDate || !endDate) return 0;
  
  let count = 0;
  let current = getStartOfDay_(startDate);
  const end = getEndOfDay_(endDate);
  
  while (current <= end) {
    if (isWorkingDay_(current)) {
      count++;
    }
    current = addDays_(current, 1);
  }
  
  return count;
}

/**
 * Создаёт диапазон дат для отчётов
 * @param {string} period - Тип периода ('today', 'yesterday', 'week', 'month', 'quarter', 'year', 'custom')
 * @param {Date} customStart - Начальная дата для custom периода
 * @param {Date} customEnd - Конечная дата для custom периода
 * @returns {Object} Объект с началом и концом периода
 */
function createDateRange_(period, customStart = null, customEnd = null) {
  const now = getCurrentDateMoscow_();
  let start, end;
  
  switch (period.toLowerCase()) {
    case 'today':
      start = getStartOfDay_(now);
      end = getEndOfDay_(now);
      break;
      
    case 'yesterday':
      const yesterday = addDays_(now, -1);
      start = getStartOfDay_(yesterday);
      end = getEndOfDay_(yesterday);
      break;
      
    case 'week':
      // Текущая неделя (понедельник - воскресенье)
      const dayOfWeek = now.getDay() === 0 ? 7 : now.getDay(); // Воскресенье = 7
      start = getStartOfDay_(addDays_(now, 1 - dayOfWeek));
      end = getEndOfDay_(now);
      break;
      
    case 'last_week':
      const lastWeekEnd = addDays_(now, -now.getDay());
      const lastWeekStart = addDays_(lastWeekEnd, -6);
      start = getStartOfDay_(lastWeekStart);
      end = getEndOfDay_(lastWeekEnd);
      break;
      
    case 'month':
      start = getStartOfMonth_(now);
      end = getEndOfMonth_(now);
      break;
      
    case 'last_month':
      const lastMonth = addMonths_(now, -1);
      start = getStartOfMonth_(lastMonth);
      end = getEndOfMonth_(lastMonth);
      break;
      
    case 'quarter':
      const quarterStartMonth = Math.floor(now.getMonth() / 3) * 3;
      start = new Date(now.getFullYear(), quarterStartMonth, 1);
      end = new Date(now.getFullYear(), quarterStartMonth + 3, 0, 23, 59, 59, 999);
      break;
      
    case 'year':
      start = getStartOfYear_(now);
      end = getEndOfYear_(now);
      break;
      
    case 'last_year':
      const lastYear = now.getFullYear() - 1;
      start = new Date(lastYear, 0, 1);
      end = new Date(lastYear, 11, 31, 23, 59, 59, 999);
      break;
      
    case 'custom':
      if (customStart && customEnd) {
        start = getStartOfDay_(parseDate_(customStart));
        end = getEndOfDay_(parseDate_(customEnd));
      } else {
        throw new Error('Для custom периода необходимо указать customStart и customEnd');
      }
      break;
      
    default:
      // По умолчанию - текущий месяц
      start = getStartOfMonth_(now);
      end = getEndOfMonth_(now);
  }
  
  return {
    start: convertToMoscowTime_(start),
    end: convertToMoscowTime_(end),
    period: period,
    description: getDateRangeDescription_(start, end, period)
  };
}

/**
 * Создаёт описание периода на русском языке
 * @param {Date} start - Начальная дата
 * @param {Date} end - Конечная дата
 * @param {string} period - Тип периода
 * @returns {string} Описание периода
 */
function getDateRangeDescription_(start, end, period) {
  if (!start || !end) return 'Неопределённый период';
  
  const startFormatted = formatDate_(start, 'DD.MM.YYYY');
  const endFormatted = formatDate_(end, 'DD.MM.YYYY');
  
  switch (period.toLowerCase()) {
    case 'today':
      return `Сегодня (${startFormatted})`;
    case 'yesterday':
      return `Вчера (${startFormatted})`;
    case 'week':
      return `Текущая неделя (${startFormatted} - ${endFormatted})`;
    case 'last_week':
      return `Прошлая неделя (${startFormatted} - ${endFormatted})`;
    case 'month':
      return `Текущий месяц (${getMonthName_(start.getMonth())} ${start.getFullYear()})`;
    case 'last_month':
      return `Прошлый месяц (${getMonthName_(start.getMonth())} ${start.getFullYear()})`;
    case 'quarter':
      return `Текущий квартал (${startFormatted} - ${endFormatted})`;
    case 'year':
      return `Текущий год (${start.getFullYear()})`;
    case 'last_year':
      return `Прошлый год (${start.getFullYear()})`;
    default:
      return `${startFormatted} - ${endFormatted}`;
  }
}

/**
 * Проверяет, попадает ли дата в заданный диапазон
 * @param {Date} date - Проверяемая дата
 * @param {Date} rangeStart - Начало диапазона
 * @param {Date} rangeEnd - Конец диапазона
 * @returns {boolean} true если дата в диапазоне
 */
function isDateInRange_(date, rangeStart, rangeEnd) {
  if (!date || !rangeStart || !rangeEnd) return false;
  
  const checkDate = convertToMoscowTime_(date);
  const start = convertToMoscowTime_(rangeStart);
  const end = convertToMoscowTime_(rangeEnd);
  
  return checkDate >= start && checkDate <= end;
}

/**
 * Создаёт массив дат для заданного периода
 * @param {Date} startDate - Начальная дата
 * @param {Date} endDate - Конечная дата
 * @param {string} interval - Интервал ('day', 'week', 'month')
 * @returns {Array<Date>} Массив дат
 */
function createDateArray_(startDate, endDate, interval = 'day') {
  if (!startDate || !endDate) return [];
  
  const dates = [];
  let current = new Date(convertToMoscowTime_(startDate));
  const end = convertToMoscowTime_(endDate);
  
  while (current <= end) {
    dates.push(new Date(current));
    
    switch (interval.toLowerCase()) {
      case 'day':
        current = addDays_(current, 1);
        break;
      case 'week':
        current = addDays_(current, 7);
        break;
      case 'month':
        current = addMonths_(current, 1);
        break;
      default:
        current = addDays_(current, 1);
    }
  }
  
  return dates;
}

/**
 * Получает временную метку для кэширования (обновляется каждый час)
 * @returns {string} Временная метка
 */
function getHourlyTimestamp_() {
  const now = getCurrentDateMoscow_();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}-${String(now.getHours()).padStart(2, '0')}`;
}

/**
 * Получает временную метку для кэширования (обновляется каждый день)
 * @returns {string} Временная метка
 */
function getDailyTimestamp_() {
  const now = getCurrentDateMoscow_();
  return formatDate_(now, 'YYYY-MM-DD');
}
