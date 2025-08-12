/**
 * 📝 Утилиты Заголовков - Обработка метаданных
 * 
 * ИНСТРУКЦИЯ ПО УСТАНОВКЕ:
 * 1. Откройте Google Apps Script: https://script.google.com/d/1u77EB0EuYkyO21LhwjVjLJOgHy_dyHZ7TL0rc8yAbbsuLRZG20l0vfUH/edit
 * 2. Создайте новый файл: + → Script file
 * 3. Назовите файл: 05_Утилиты_Заголовков (без .gs)
 * 4. Удалите весь код по умолчанию
 * 5. Скопируйте и вставьте код ниже
 * 6. Сохраните файл (Ctrl+S)
 */

/**
 * УТИЛИТЫ ДЛЯ РАБОТЫ С ЗАГОЛОВКАМИ И КОЛОНКАМИ
 * Нормализация, поиск, безопасное чтение/запись
 * @fileoverview Универсальные функции для работы с данными листов
 */

/**
 * Нормализует строку для сравнения заголовков
 * @param {string} str - Исходная строка
 * @returns {string} Нормализованная строка
 * 
 * @example
 * normalize_('  Сделка.ID  ') // 'сделка id'
 * normalize_('UTM_SOURCE') // 'utm source'
 */
function normalize_(str) {
  if (!str) return '';
  
  return String(str)
    .toLowerCase()
    .trim()
    // Убираем спецсимволы и заменяем пробелами
    .replace(/[\.\_\-\(\)\[\]\{\}\/\\:;,]/g, ' ')
    // Убираем множественные пробелы
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Находит индекс колонки по алиасам
 * @param {string[]} headerRow - Строка заголовков
 * @param {string[]} aliases - Возможные названия колонки
 * @returns {number} Индекс колонки или -1
 * 
 * @example
 * findHeaderIndex_(['ID', 'Статус', 'Телефон'], ['phone', 'телефон']) // 2
 */
function findHeaderIndex_(headerRow, aliases) {
  if (!headerRow || !aliases) return -1;
  
  const normalizedHeader = headerRow.map(h => normalize_(h));
  const normalizedAliases = aliases.map(a => normalize_(a));
  
  for (let i = 0; i < normalizedHeader.length; i++) {
    const header = normalizedHeader[i];
    
    // Точное совпадение
    if (normalizedAliases.includes(header)) {
      return i;
    }
    
    // Частичное совпадение
    for (const alias of normalizedAliases) {
      if (header.includes(alias) || alias.includes(header)) {
        return i;
      }
    }
  }
  
  return -1;
}

/**
 * Находит индексы всех колонок по словарю алиасов
 * @param {string[]} headerRow - Строка заголовков
 * @param {Object} aliasMap - Словарь алиасов из конфига
 * @returns {Object} Объект с индексами колонок
 * 
 * @example
 * findAllColumns_(['ID', 'Статус'], {id: ['ID', 'id'], status: ['Статус']})
 * // {id: 0, status: 1}
 */
function findAllColumns_(headerRow, aliasMap = CONFIG.HEADER_ALIASES) {
  const indices = {};
  
  for (const [key, aliases] of Object.entries(aliasMap)) {
    const index = findHeaderIndex_(headerRow, aliases);
    indices[key] = index >= 0 ? index : null;
  }
  
  return indices;
}

/**
 * Безопасное чтение данных с листа
 * @param {SpreadsheetApp.Sheet} sheet - Лист
 * @param {Object} options - Опции чтения
 * @returns {Object} Объект с заголовками и данными
 */
function safeReadSheet_(sheet, options = {}) {
  const {
    headerRow = 1,
    startDataRow = 2,
    includeEmpty = false,
    maxRows = null
  } = options;
  
  try {
    logDebug_('READ', `Чтение листа "${sheet.getName()}"`);
    
    // Проверяем, что лист существует и не пустой
    if (!sheet) {
      throw new Error('Лист не передан');
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < headerRow || lastCol === 0) {
      logWarning_('READ', `Лист "${sheet.getName()}" пустой`);
      return { headers: [], data: [], indices: {}, rowCount: 0, columnCount: 0 };
    }
    
    // Читаем заголовки
    const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    
    // Читаем данные
    let data = [];
    if (lastRow >= startDataRow) {
      const endRow = maxRows ? Math.min(lastRow, startDataRow + maxRows - 1) : lastRow;
      const dataRange = sheet.getRange(startDataRow, 1, endRow - startDataRow + 1, lastCol);
      data = dataRange.getValues();
      
      // Фильтруем пустые строки
      if (!includeEmpty) {
        data = data.filter(row => row.some(cell => cell !== '' && cell !== null && cell !== undefined));
      }
    }
    
    // Находим индексы колонок
    const indices = findAllColumns_(headers);
    
    logInfo_('READ', `Лист "${sheet.getName()}": ${data.length} строк, ${headers.length} колонок`);
    
    return {
      headers,
      data,
      indices,
      rowCount: data.length,
      columnCount: headers.length,
      lastRow: lastRow,
      lastColumn: lastCol
    };
    
  } catch (error) {
    logError_('READ', `Ошибка чтения листа "${sheet.getName()}"`, error);
    return { headers: [], data: [], indices: {}, error: error.toString(), rowCount: 0, columnCount: 0 };
  }
}

/**
 * Безопасная запись данных на лист
 * @param {SpreadsheetApp.Sheet} sheet - Лист
 * @param {Array[]} data - Данные для записи
 * @param {Object} options - Опции записи
 * @returns {boolean} Успешность операции
 */
function safeWriteSheet_(sheet, data, options = {}) {
  const {
    startRow = 1,
    startCol = 1,
    clearSheet = false,
    headers = null,
    preserveFormulas = false,
    batchSize = 1000
  } = options;
  
  try {
    logDebug_('WRITE', `Запись на лист "${sheet.getName()}": ${data.length} строк`);
    
    if (!sheet || !data || data.length === 0) {
      throw new Error('Недостаточно данных для записи');
    }
    
    // Очищаем лист если нужно
    if (clearSheet) {
      if (preserveFormulas) {
        sheet.getRange(startRow, startCol, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
      } else {
        sheet.clear();
        sheet.clearFormats();
      }
    }
    
    // Записываем заголовки
    let currentRow = startRow;
    if (headers && headers.length > 0) {
      const headerRange = sheet.getRange(currentRow, startCol, 1, headers.length);
      headerRange.setValues([headers]);
      
      // Форматируем заголовки
      headerRange
        .setFontWeight('bold')
        .setFontSize(11)
        .setBackground(CONFIG.COLORS.backgrounds.header)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
        
      currentRow++;
    }
    
    // Записываем данные порциями
    if (data.length > 0) {
      const numCols = data[0].length;
      
      // Проверяем размер листа
      ensureSheetSize_(sheet, currentRow + data.length - 1, startCol + numCols - 1);
      
      // Записываем данными порциями для избежания таймаутов
      for (let i = 0; i < data.length; i += batchSize) {
        const batch = data.slice(i, i + batchSize);
        const batchStartRow = currentRow + i;
        
        const dataRange = sheet.getRange(batchStartRow, startCol, batch.length, numCols);
        dataRange.setValues(batch);
        
        // Небольшая пауза между батчами
        if (i + batchSize < data.length) {
          Utilities.sleep(100);
        }
      }
      
      // Базовое форматирование данных
      const dataRange = sheet.getRange(currentRow, startCol, data.length, numCols);
      dataRange
        .setFontFamily(CONFIG.DEFAULT_FONT)
        .setFontSize(10)
        .setVerticalAlignment('middle');
    }
    
    logInfo_('WRITE', `Успешно записано на лист "${sheet.getName()}": ${data.length} строк`);
    return true;
    
  } catch (error) {
    logError_('WRITE', `Ошибка записи на лист "${sheet.getName()}"`, error);
    return false;
  }
}

/**
 * Автоматический поиск строки с заголовками
 * @param {SpreadsheetApp.Sheet} sheet - Лист
 * @param {number} maxRows - Максимальное количество строк для поиска
 * @returns {number} Номер строки с заголовками
 */
function findHeaderRow_(sheet, maxRows = 10) {
  if (!sheet || sheet.getLastRow() === 0) return -1;
  
  try {
    const lastRow = Math.min(sheet.getLastRow(), maxRows);
    const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
    
    // Ищем строку с максимальным количеством непустых ячеек
    let bestRow = -1;
    let maxScore = 0;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      let score = 0;
      
      // Количество непустых ячеек
      const nonEmptyCount = row.filter(cell => cell !== '' && cell !== null).length;
      score += nonEmptyCount;
      
      // Бонус за типичные заголовки
      const hasTypicalHeaders = row.some(cell => {
        const normalized = normalize_(String(cell));
        return ['id', 'статус', 'дата', 'телефон', 'сумма', 'бюджет', 'контакт', 'utm'].some(h => 
          normalized.includes(h)
        );
      });
      if (hasTypicalHeaders) score += 10;
      
      // Штраф за цифры (заголовки обычно не содержат только цифры)
      const hasOnlyNumbers = row.every(cell => {
        const str = String(cell).trim();
        return str === '' || !isNaN(Number(str));
      });
      if (hasOnlyNumbers) score -= 5;
      
      if (score > maxScore) {
        maxScore = score;
        bestRow = i + 1; // Преобразуем в 1-based индекс
      }
    }
    
    return bestRow > 0 ? bestRow : 1;
    
  } catch (error) {
    logError_('HEADERS', `Ошибка поиска строки заголовков в "${sheet.getName()}"`, error);
    return 1;
  }
}

/**
 * Проверяет и расширяет размер листа при необходимости
 * @param {SpreadsheetApp.Sheet} sheet - Лист
 * @param {number} neededRows - Необходимое количество строк
 * @param {number} neededCols - Необходимое количество колонок
 */
function ensureSheetSize_(sheet, neededRows, neededCols) {
  try {
    const currentRows = sheet.getMaxRows();
    const currentCols = sheet.getMaxColumns();
    
    if (neededRows > currentRows) {
      const rowsToAdd = neededRows - currentRows;
      sheet.insertRowsAfter(currentRows, rowsToAdd);
      logDebug_('RESIZE', `Добавлено ${rowsToAdd} строк в "${sheet.getName()}"`);
    }
    
    if (neededCols > currentCols) {
      const colsToAdd = neededCols - currentCols;
      sheet.insertColumnsAfter(currentCols, colsToAdd);
      logDebug_('RESIZE', `Добавлено ${colsToAdd} колонок в "${sheet.getName()}"`);
    }
    
  } catch (error) {
    logError_('RESIZE', `Ошибка изменения размера листа "${sheet.getName()}"`, error);
  }
}

/**
 * Создаёт маппинг заголовков для быстрого доступа
 * @param {string[]} headers - Массив заголовков
 * @returns {Object} Объект для быстрого доступа по имени
 * 
 * @example
 * const map = createHeaderMap_(['ID', 'Статус', 'Телефон'])
 * map['статус'] // 1
 * map['СТАТУС'] // 1
 */
function createHeaderMap_(headers) {
  const map = {};
  
  headers.forEach((header, index) => {
    if (header) {
      // Сохраняем оригинальное название
      map[header] = index;
      
      // Сохраняем нормализованное название
      const normalized = normalize_(header);
      if (normalized) {
        map[normalized] = index;
      }
      
      // Сохраняем в верхнем регистре
      map[header.toUpperCase()] = index;
      
      // Сохраняем в нижнем регистре
      map[header.toLowerCase()] = index;
    }
  });
  
  return map;
}

/**
 * Валидирует наличие обязательных колонок
 * @param {Object} indices - Объект с индексами колонок
 * @param {string[]} required - Список обязательных колонок
 * @throws {Error} Если обязательные колонки не найдены
 */
function validateRequiredColumns_(indices, required) {
  const missing = required.filter(col => indices[col] === null || indices[col] === undefined || indices[col] < 0);
  
  if (missing.length > 0) {
    const error = `Не найдены обязательные колонки: ${missing.join(', ')}`;
    logError_('VALIDATION', error);
    throw new Error(error);
  }
  
  logDebug_('VALIDATION', `Все обязательные колонки найдены: ${required.join(', ')}`);
}

/**
 * Получает значение ячейки с безопасным преобразованием типа
 * @param {*} value - Значение ячейки
 * @param {string} type - Тип для преобразования (string, number, date, boolean)
 * @returns {*} Преобразованное значение
 */
function safeCellValue_(value, type = 'string') {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  
  try {
    switch (type.toLowerCase()) {
      case 'string':
        return String(value).trim();
        
      case 'number':
        const num = Number(value);
        return isNaN(num) ? null : num;
        
      case 'date':
        if (value instanceof Date) return value;
        const date = new Date(value);
        return isNaN(date.getTime()) ? null : date;
        
      case 'boolean':
        if (typeof value === 'boolean') return value;
        const str = String(value).toLowerCase().trim();
        return ['true', '1', 'да', 'yes', 'y'].includes(str);
        
      case 'phone':
        return cleanPhone_(value);
        
      case 'email':
        const email = String(value).trim().toLowerCase();
        return email.includes('@') ? email : null;
        
      default:
        return value;
    }
  } catch (error) {
    logWarning_('CELL_VALUE', `Ошибка преобразования значения "${value}" в тип "${type}"`);
    return null;
  }
}

/**
 * Очищает и нормализует номер телефона
 * @param {string} phone - Номер телефона
 * @returns {string|null} Очищенный номер телефона
 */
function cleanPhone_(phone) {
  if (!phone) return null;
  
  // Удаляем все кроме цифр и знака +
  let cleaned = String(phone).replace(/[^\d+]/g, '');
  
  // Убираем ведущие нули и семёрки
  cleaned = cleaned.replace(/^[07]+/, '');
  
  // Если номер начинается с 9 и состоит из 10 цифр - это российский номер
  if (/^9\d{9}$/.test(cleaned)) {
    cleaned = '+7' + cleaned;
  }
  // Если 11 цифр и начинается с 7
  else if (/^7\d{10}$/.test(cleaned)) {
    cleaned = '+' + cleaned;
  }
  // Если есть международный код
  else if (cleaned.startsWith('+')) {
    // Оставляем как есть
  }
  // Иначе добавляем российский код
  else if (/^\d{10}$/.test(cleaned)) {
    cleaned = '+7' + cleaned;
  }
  
  return cleaned.length >= 10 ? cleaned : null;
}

/**
 * Получает данные из строки по индексам колонок с типизацией
 * @param {Array} row - Строка данных
 * @param {Object} indices - Индексы колонок
 * @param {Object} types - Типы колонок
 * @returns {Object} Объект с типизированными данными
 */
function extractRowData_(row, indices, types = {}) {
  const result = {};
  
  for (const [key, index] of Object.entries(indices)) {
    if (index !== null && index >= 0 && index < row.length) {
      const type = types[key] || 'string';
      result[key] = safeCellValue_(row[index], type);
    } else {
      result[key] = null;
    }
  }
  
  return result;
}

/**
 * Копирует лист с сохранением форматирования
 * @param {SpreadsheetApp.Sheet} sourceSheet - Исходный лист
 * @param {string} newName - Имя нового листа
 * @param {SpreadsheetApp.Spreadsheet} targetSpreadsheet - Целевая таблица
 * @returns {SpreadsheetApp.Sheet} Новый лист
 */
function copySheet_(sourceSheet, newName, targetSpreadsheet = null) {
  try {
    const ss = targetSpreadsheet || sourceSheet.getParent();
    
    // Удаляем лист если он уже существует
    const existingSheet = ss.getSheetByName(newName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // Копируем лист
    const newSheet = sourceSheet.copyTo(ss).setName(newName);
    
    logInfo_('COPY', `Лист "${sourceSheet.getName()}" скопирован как "${newName}"`);
    return newSheet;
    
  } catch (error) {
    logError_('COPY', `Ошибка копирования листа "${sourceSheet.getName()}"`, error);
    throw error;
  }
}

// Функции логирования (будут определены в отдельном модуле)
function logDebug_(module, message, details = null) {
  if (CONFIG.DEBUG.enabled && CONFIG.DEBUG.log_level === 'DEBUG') {
    console.log(`[DEBUG] ${module}: ${message}`, details || '');
  }
}

function logInfo_(module, message, details = null) {
  if (CONFIG.DEBUG.enabled) {
    console.log(`[INFO] ${module}: ${message}`, details || '');
  }
}

function logWarning_(module, message, details = null) {
  console.warn(`[WARNING] ${module}: ${message}`, details || '');
}

function logError_(module, message, details = null) {
  console.error(`[ERROR] ${module}: ${message}`, details || '');
}
