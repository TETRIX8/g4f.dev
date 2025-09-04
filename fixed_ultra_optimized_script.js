// === ИСПРАВЛЕННАЯ УЛЬТРА-ОПТИМИЗИРОВАННАЯ ВЕРСИЯ ===
// Исправления:
// 1. Убраны все setTimeout/clearTimeout (не поддерживаются в Google Apps Script)
// 2. Заменены асинхронные операции на синхронные
// 3. Упрощена система буферизации логов
// 4. Исправлены все веб-специфичные функции

// === Глобальные константы ===
const SHEET_PRODUCTS = "1. Товары";
const SHEET_PURCHASES = "3 Закупки";
const SHEET_SALES = "4 Продажи";

// ID файла на Google Диске
const CACHE_FILE_ID = "1Mk2tr9z1ZA1uxAzNZdWy9z4qUHXLCfoT";

// Настройки производительности
const DEBUG = true;
const PERFORMANCE_LOGGING = true;
const CACHE_TTL_SECONDS = 7200;
const MEMORY_CACHE_TTL = 1800;
const BATCH_SIZE = 1000;
const PRELOAD_ON_STARTUP = true;

// Глобальные переменные для ультра-быстрого доступа
let PRODUCTS_CACHE = null;
let PRODUCTS_INDEX = null;
let CACHE_TIMESTAMP = 0;
let LAST_SHEET_DATA_HASH = null;
let CALL_STATISTICS = {};
let PERFORMANCE_METRICS = {};

// === Система логирования производительности ===
function startTimer(operationName) {
  if (!PERFORMANCE_LOGGING) return null;
  
  const timerId = `${operationName}_${Date.now()}_${Math.random()}`;
  PERFORMANCE_METRICS[timerId] = {
    name: operationName,
    startTime: Date.now(),
    startMemory: getMemoryUsage()
  };
  
  return timerId;
}

function endTimer(timerId, additionalInfo = '') {
  if (!PERFORMANCE_LOGGING || !timerId || !PERFORMANCE_METRICS[timerId]) return;
  
  const metric = PERFORMANCE_METRICS[timerId];
  const duration = Date.now() - metric.startTime;
  const memoryUsed = getMemoryUsage() - metric.startMemory;
  
  // Обновляем статистику
  if (!CALL_STATISTICS[metric.name]) {
    CALL_STATISTICS[metric.name] = {
      count: 0,
      totalTime: 0,
      minTime: Infinity,
      maxTime: 0,
      avgTime: 0,
      lastCall: 0
    };
  }
  
  const stats = CALL_STATISTICS[metric.name];
  stats.count++;
  stats.totalTime += duration;
  stats.minTime = Math.min(stats.minTime, duration);
  stats.maxTime = Math.max(stats.maxTime, duration);
  stats.avgTime = stats.totalTime / stats.count;
  stats.lastCall = Date.now();
  
  const logMsg = `⏱️ ${metric.name}: ${duration}ms${memoryUsed ? ` | Память: ${memoryUsed}KB` : ''}${additionalInfo ? ` | ${additionalInfo}` : ''}`;
  log(logMsg);
  
  // Предупреждения о медленных операциях
  if (duration > 1000) {
    log(`🐌 МЕДЛЕННАЯ ОПЕРАЦИЯ: ${metric.name} заняла ${duration}ms!`);
  } else if (duration > 100) {
    log(`⚠️ Долгая операция: ${metric.name} заняла ${duration}ms`);
  }
  
  delete PERFORMANCE_METRICS[timerId];
}

function getMemoryUsage() {
  try {
    return Math.round((JSON.stringify(PRODUCTS_CACHE || {}).length + 
                      JSON.stringify(PRODUCTS_INDEX || {}).length) / 1024);
  } catch (e) {
    return 0;
  }
}

// === Главный триггер onEdit (УЛЬТРА-ОПТИМИЗИРОВАННЫЙ) ===
function onEdit(e) {
  const timerId = startTimer('onEdit_full');
  
  if (!e || !e.range) {
    endTimer(timerId, 'Нет события');
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();

  log(`🎯 onEdit: лист="${sheetName}", ячейка=${range.getA1Notation()}, строка=${row}, колонка=${col}`);

  // Ультра-быстрая предварительная проверка
  const quickCheckTimer = startTimer('onEdit_quickCheck');
  
  if (col !== 2 || row < 2 || range.getNumRows() > 1 || range.getNumColumns() > 1) {
    endTimer(quickCheckTimer, 'Пропуск - не подходящий диапазон');
    endTimer(timerId, 'Завершено - пропуск');
    return;
  }

  if (sheetName !== SHEET_PURCHASES && sheetName !== SHEET_SALES) {
    endTimer(quickCheckTimer, 'Пропуск - не подходящий лист');
    endTimer(timerId, 'Завершено - неподходящий лист');
    return;
  }
  
  endTimer(quickCheckTimer, 'Проверка пройдена');

  // Читаем значение
  const valueTimer = startTimer('getValue');
  const value = range.getValue();
  endTimer(valueTimer, `Значение: "${value}"`);

  log(`✏️ Изменение: лист="${sheetName}", ячейка=${range.getA1Notation()}, значение="${value}"`);

  // Обработка пустого значения
  if (!value) {
    const clearTimer = startTimer('clearRow');
    const colsToClean = sheetName === SHEET_PURCHASES ? [1, 3, 4] : [1, 3, 4, 5];
    clearRowBatch(sheet, row, colsToClean);
    endTimer(clearTimer, `Очищено ${colsToClean.length} колонок`);
    endTimer(timerId, 'Завершено - очистка');
    return;
  }

  // Основная обработка
  const mode = sheetName === SHEET_PURCHASES ? "purchase" : "sale";
  handleBarcodeUltraOptimized(value, sheet, row, mode);
  
  endTimer(timerId, `Режим: ${mode}`);
}

// === УЛЬТРА-ОПТИМИЗИРОВАННАЯ обработка штрих-кода ===
function handleBarcodeUltraOptimized(barcode, sheet, row, mode) {
  const timerId = startTimer('handleBarcode_full');
  
  const bc = String(barcode).trim();
  if (!bc) {
    log("⚠️ Пустой штрих-код");
    endTimer(timerId, 'Пустой штрих-код');
    return;
  }

  log(`🔍 Поиск товара: "${bc}"`);

  // Ультра-быстрый поиск с индексированием
  const searchTimer = startTimer('productSearch');
  const product = findProductUltraFast(bc);
  endTimer(searchTimer, product ? `Найден: ${product.name}` : 'Не найден');

  if (!product) {
    const errorTimer = startTimer('setErrorMessage');
    sheet.getRange(row, 3).setValue("❌ Не найден");
    endTimer(errorTimer);
    log(`❌ Товар не найден: "${bc}"`);
    endTimer(timerId, 'Товар не найден');
    return;
  }

  log(`✅ Товар найден: "${product.name}" (цена: ${product.price})`);

  // Подготовка данных для записи
  const prepareTimer = startTimer('prepareData');
  const now = new Date();
  const updates = [];
  
  if (mode === "purchase") {
    updates.push([row, 1, now], [row, 3, product.name], [row, 4, 1]);
    log(`📥 Закупка: "${product.name}", количество=1`);
  } else if (mode === "sale") {
    updates.push([row, 1, now], [row, 3, product.name], [row, 4, 1], [row, 5, product.price]);
    log(`📤 Продажа: "${product.name}", цена=${product.price}`);
  }
  endTimer(prepareTimer, `Подготовлено ${updates.length} обновлений`);

  // Выполняем все обновления одним батчем
  const updateTimer = startTimer('batchUpdate');
  batchUpdateCellsOptimized(sheet, updates);
  endTimer(updateTimer, `Обновлено ${updates.length} ячеек`);

  // Авто-переход для продаж
  if (mode === "sale") {
    const navigationTimer = startTimer('autoNavigation');
    const nextRow = row + 1;
    sheet.setActiveSelection(`B${nextRow}`);
    endTimer(navigationTimer, `Переход к B${nextRow}`);
    log(`🎯 Авто-переход к строке ${nextRow}`);
  }
  
  endTimer(timerId, `Успешно обработан ${mode}`);
}

// === УЛЬТРА-БЫСТРЫЙ поиск с индексированием ===
function findProductUltraFast(barcode) {
  const timerId = startTimer('findProduct');
  
  // Проверяем индекс в памяти (O(1) доступ)
  if (PRODUCTS_INDEX && PRODUCTS_INDEX[barcode]) {
    const product = PRODUCTS_INDEX[barcode];
    endTimer(timerId, `Найден в индексе: ${product.name}`);
    log(`⚡ Мгновенный поиск в индексе: "${product.name}"`);
    return product;
  }

  // Загружаем данные если нужно
  const loadTimer = startTimer('loadProductsIfNeeded');
  const products = getAllProductsUltraOptimized();
  endTimer(loadTimer, `Загружено ${Object.keys(products).length} товаров`);
  
  const result = products[barcode] || null;
  endTimer(timerId, result ? `Найден: ${result.name}` : 'Не найден');
  
  return result;
}

// === УЛЬТРА-ОПТИМИЗИРОВАННОЕ получение товаров ===
function getAllProductsUltraOptimized() {
  const timerId = startTimer('getAllProducts');
  
  try {
    // 1. Проверяем индекс в памяти (самый быстрый - O(1))
    const memoryCheckTimer = startTimer('memoryCheck');
    if (PRODUCTS_INDEX && (Date.now() - CACHE_TIMESTAMP) < (MEMORY_CACHE_TTL * 1000)) {
      endTimer(memoryCheckTimer, `Используем индекс: ${Object.keys(PRODUCTS_INDEX).length} товаров`);
      endTimer(timerId, 'Из индекса памяти');
      log(`⚡ Ультра-быстрый доступ к индексу: ${Object.keys(PRODUCTS_INDEX).length} товаров`);
      return PRODUCTS_INDEX;
    }
    endTimer(memoryCheckTimer, 'Индекс устарел или отсутствует');

    // 2. Проверяем CacheService
    const cacheServiceTimer = startTimer('cacheServiceCheck');
    const cacheKey = "products_ultra_v6_fixed";
    const cached = CacheService.getUserCache().get(cacheKey);
    
    if (cached) {
      const data = JSON.parse(cached);
      if (data.timestamp && (Date.now() - data.timestamp) < (CACHE_TTL_SECONDS * 1000)) {
        // Строим индекс для быстрого доступа
        const indexBuildTimer = startTimer('buildIndex');
        PRODUCTS_INDEX = data.products;
        PRODUCTS_CACHE = data.products;
        CACHE_TIMESTAMP = Date.now();
        endTimer(indexBuildTimer, `Построен индекс: ${Object.keys(PRODUCTS_INDEX).length} товаров`);
        
        endTimer(cacheServiceTimer, 'Загружено из CacheService');
        endTimer(timerId, 'Из CacheService');
        log(`✅ Быстрая загрузка из CacheService: ${Object.keys(PRODUCTS_INDEX).length} товаров`);
        return PRODUCTS_INDEX;
      }
    }
    endTimer(cacheServiceTimer, 'CacheService пуст или устарел');

    // 3. Проверяем изменения данных
    const hashCheckTimer = startTimer('hashCheck');
    const currentDataHash = getSheetDataHashOptimized();
    if (LAST_SHEET_DATA_HASH === currentDataHash && PRODUCTS_INDEX) {
      endTimer(hashCheckTimer, 'Данные не изменились');
      endTimer(timerId, 'Данные актуальны');
      log(`✅ Данные актуальны, используем существующий индекс`);
      return PRODUCTS_INDEX;
    }
    endTimer(hashCheckTimer, `Новый хеш: ${currentDataHash}`);

    // 4. Пытаемся загрузить из JSON
    const jsonTimer = startTimer('jsonLoad');
    const jsonData = readJsonCacheUltraOptimized();
    if (jsonData && jsonData.hash === currentDataHash && jsonData.products) {
      // Строим индекс
      const indexTimer = startTimer('buildIndexFromJson');
      PRODUCTS_INDEX = jsonData.products;
      PRODUCTS_CACHE = jsonData.products;
      CACHE_TIMESTAMP = Date.now();
      LAST_SHEET_DATA_HASH = currentDataHash;
      endTimer(indexTimer, `Индекс построен: ${Object.keys(PRODUCTS_INDEX).length} товаров`);
      
      // Сохраняем в CacheService
      saveToCacheService(jsonData.products, currentDataHash);
      
      endTimer(jsonTimer, 'Загружено из JSON');
      endTimer(timerId, 'Из JSON файла');
      log(`✅ Загрузка из JSON с построением индекса: ${Object.keys(PRODUCTS_INDEX).length} товаров`);
      return PRODUCTS_INDEX;
    }
    endTimer(jsonTimer, 'JSON недоступен или устарел');

    // 5. Загружаем из таблицы (критический путь)
    log(`🔄 Критическая загрузка из таблицы...`);
    const sheetTimer = startTimer('sheetLoad');
    const products = loadFromSheetUltraOptimized();
    endTimer(sheetTimer, `Загружено ${Object.keys(products).length} товаров`);
    
    // Строим индекс
    const finalIndexTimer = startTimer('buildFinalIndex');
    PRODUCTS_INDEX = products;
    PRODUCTS_CACHE = products;
    CACHE_TIMESTAMP = Date.now();
    LAST_SHEET_DATA_HASH = currentDataHash;
    endTimer(finalIndexTimer, `Финальный индекс: ${Object.keys(products).length} товаров`);

    // Сохраняем во все кеши
    saveToAllCaches(products, currentDataHash);

    endTimer(timerId, 'Из Google Sheets');
    log(`✅ Критическая загрузка завершена: ${Object.keys(products).length} товаров`);
    return products;

  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    log(`❌ Критическая ошибка загрузки: ${error.message}`);
    return {};
  }
}

// === ОПТИМИЗИРОВАННОЕ получение хеша данных ===
function getSheetDataHashOptimized() {
  const timerId = startTimer('getDataHash');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PRODUCTS);
    if (!sheet) {
      endTimer(timerId, 'Лист не найден');
      return "no_sheet";
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      endTimer(timerId, 'Пустой лист');
      return "empty_sheet";
    }
    
    // Более точный хеш на основе диапазона данных
    const dataRange = sheet.getDataRange();
    const numRows = dataRange.getNumRows();
    const numCols = dataRange.getNumColumns();
    
    // Читаем только последнюю строку для хеша (экономим время)
    const lastRowData = sheet.getRange(lastRow, 1, 1, Math.min(numCols, 5)).getValues()[0];
    const lastModified = lastRowData.join('|');
    
    const hash = `${numRows}_${numCols}_${lastModified.length}_${lastModified.slice(0, 20)}`;
    endTimer(timerId, `Хеш: ${hash}`);
    return hash;
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    return "error_" + Date.now();
  }
}

// === УЛЬТРА-ОПТИМИЗИРОВАННОЕ чтение JSON ===
function readJsonCacheUltraOptimized() {
  const timerId = startTimer('readJsonCache');
  
  try {
    const file = DriveApp.getFileById(CACHE_FILE_ID);
    const content = file.getContentAsString();
    
    if (!content || content.trim() === '') {
      endTimer(timerId, 'JSON пуст');
      return null;
    }
    
    const parseTimer = startTimer('parseJson');
    const data = JSON.parse(content);
    endTimer(parseTimer, `Размер: ${content.length} символов`);
    
    if (!data.products || !data.timestamp) {
      endTimer(timerId, 'Неверная структура');
      return null;
    }
    
    const age = Date.now() - data.timestamp;
    if (age > (CACHE_TTL_SECONDS * 1000)) {
      endTimer(timerId, `Устарел на ${Math.round(age/1000)}с`);
      return null;
    }
    
    endTimer(timerId, `Успешно загружен, возраст: ${Math.round(age/1000)}с`);
    return data;
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    return null;
  }
}

// === УЛЬТРА-ОПТИМИЗИРОВАННАЯ загрузка из таблицы ===
function loadFromSheetUltraOptimized() {
  const timerId = startTimer('loadFromSheet');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  
  if (!sheet) {
    endTimer(timerId, 'Лист не найден');
    return {};
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    endTimer(timerId, 'Нет данных');
    return {};
  }

  // КРИТИЧЕСКАЯ ОПТИМИЗАЦИЯ: читаем весь диапазон одним вызовом
  const readTimer = startTimer('readSheetData');
  const range = sheet.getRange(2, 1, lastRow - 1, 5);
  const values = range.getValues();
  endTimer(readTimer, `Прочитано ${values.length} строк`);
  
  // Обработка в памяти с предварительной аллокацией
  const processTimer = startTimer('processData');
  const products = Object.create(null); // Быстрый объект без прототипа
  let validCount = 0;
  let emptyBarcodes = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const barcode = String(row[3] || '').trim();
    
    if (!barcode) {
      emptyBarcodes++;
      continue;
    }
    
    // Создаем объект товара с минимальными данными
    products[barcode] = {
      name: row[1] || 'Без названия',
      price: row[4] || 0,
      row: i + 2,
      index: validCount // Для быстрой сортировки
    };
    
    validCount++;
  }
  
  endTimer(processTimer, `Обработано: ${validCount} товаров, пропущено: ${emptyBarcodes} пустых`);
  endTimer(timerId, `Итого товаров: ${validCount}`);
  
  log(`📊 Загрузка из таблицы: ${validCount} товаров из ${values.length} строк (пропущено ${emptyBarcodes} пустых штрих-кодов)`);
  return products;
}

// === Сохранение в CacheService ===
function saveToCacheService(products, dataHash) {
  const timerId = startTimer('saveToCacheService');
  
  try {
    const cacheData = {
      products: products,
      timestamp: Date.now(),
      hash: dataHash
    };
    
    CacheService.getUserCache().put("products_ultra_v6_fixed", JSON.stringify(cacheData), CACHE_TTL_SECONDS);
    endTimer(timerId, `Сохранено ${Object.keys(products).length} товаров`);
    log(`💾 Сохранение в CacheService: ${Object.keys(products).length} товаров`);
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    log(`⚠️ Ошибка сохранения в CacheService: ${error.message}`);
  }
}

// === Сохранение во все кеши ===
function saveToAllCaches(products, dataHash) {
  const timerId = startTimer('saveAllCaches');
  
  try {
    const timestamp = Date.now();
    const cacheData = {
      products: products,
      timestamp: timestamp,
      hash: dataHash
    };

    // CacheService (быстро)
    const cacheServiceTimer = startTimer('saveCacheService');
    CacheService.getUserCache().put("products_ultra_v6_fixed", JSON.stringify(cacheData), CACHE_TTL_SECONDS);
    endTimer(cacheServiceTimer);
    
    // JSON файл (медленнее, но персистентно)
    const jsonTimer = startTimer('saveJsonFile');
    const file = DriveApp.getFileById(CACHE_FILE_ID);
    file.setContent(JSON.stringify(cacheData, null, 2));
    endTimer(jsonTimer, `Размер файла: ${JSON.stringify(cacheData).length} символов`);
    
    endTimer(timerId, `Сохранено на всех уровнях: ${Object.keys(products).length} товаров`);
    log(`💾 Сохранение завершено: ${Object.keys(products).length} товаров`);
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    log(`⚠️ Ошибка сохранения: ${error.message}`);
  }
}

// === ОПТИМИЗИРОВАННОЕ батчевое обновление ===
function batchUpdateCellsOptimized(sheet, updates) {
  const timerId = startTimer('batchUpdate');
  
  if (!updates || updates.length === 0) {
    endTimer(timerId, 'Нет обновлений');
    return;
  }
  
  try {
    // Группируем обновления по строкам для минимизации вызовов API
    const rowUpdates = {};
    
    for (const [row, col, value] of updates) {
      if (!rowUpdates[row]) rowUpdates[row] = {};
      rowUpdates[row][col] = value;
    }
    
    // Выполняем обновления построчно
    let totalUpdated = 0;
    for (const [row, cols] of Object.entries(rowUpdates)) {
      const rowNum = parseInt(row);
      const colNums = Object.keys(cols).map(Number).sort((a, b) => a - b);
      
      // Если колонки подряд, используем диапазон
      if (colNums.length > 1 && colNums[colNums.length - 1] - colNums[0] === colNums.length - 1) {
        const startCol = colNums[0];
        const values = colNums.map(col => cols[col]);
        sheet.getRange(rowNum, startCol, 1, values.length).setValues([values]);
        totalUpdated += values.length;
      } else {
        // Иначе обновляем по одной
        for (const col of colNums) {
          sheet.getRange(rowNum, col).setValue(cols[col]);
          totalUpdated++;
        }
      }
    }
    
    endTimer(timerId, `Обновлено ${totalUpdated} ячеек в ${Object.keys(rowUpdates).length} строках`);
    log(`📝 Оптимизированное обновление: ${totalUpdated} ячеек`);
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    log(`❌ Ошибка батчевого обновления: ${error.message}`);
  }
}

// === ОПТИМИЗИРОВАННАЯ очистка строки ===
function clearRowBatch(sheet, row, cols) {
  const timerId = startTimer('clearRow');
  
  if (!cols || cols.length === 0) {
    endTimer(timerId, 'Нет колонок для очистки');
    return;
  }
  
  try {
    // Если колонки подряд, очищаем диапазоном
    const sortedCols = [...cols].sort((a, b) => a - b);
    if (sortedCols.length > 1 && sortedCols[sortedCols.length - 1] - sortedCols[0] === sortedCols.length - 1) {
      const startCol = sortedCols[0];
      sheet.getRange(row, startCol, 1, sortedCols.length).clearContent();
    } else {
      // Иначе очищаем по одной
      sortedCols.forEach(col => sheet.getRange(row, col).clearContent());
    }
    
    endTimer(timerId, `Очищено ${cols.length} колонок`);
    log(`🧹 Очищена строка ${row}, колонки: ${cols.join(', ')}`);
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    log(`❌ Ошибка очистки строки: ${error.message}`);
  }
}

// === УПРОЩЁННОЕ логирование без setTimeout ===
function log(message) {
  if (!DEBUG) return;
  
  try {
    const timestamp = new Date().toLocaleTimeString('ru-RU', {
      hour12: false,
      hour: '2-digit',
      minute: '2-digit', 
      second: '2-digit',
      fractionalSecondDigits: 3
    });
    
    const logEntry = `${timestamp} | ${message}`;
    
    // Прямое сохранение логов без буферизации (для надёжности)
    const cache = CacheService.getUserCache();
    const existingLogs = JSON.parse(cache.get("logs_ultra_v3_fixed") || "[]");
    
    existingLogs.push(logEntry);
    
    // Ограничиваем размер
    if (existingLogs.length > 200) {
      existingLogs.splice(0, existingLogs.length - 200);
    }
    
    cache.put("logs_ultra_v3_fixed", JSON.stringify(existingLogs), 3600);
    
  } catch (error) {
    console.error("Ошибка логирования:", error);
  }
}

// === Получение логов ===
function getUserLogs() {
  return JSON.parse(CacheService.getUserCache().get("logs_ultra_v3_fixed") || "[]");
}

// === Очистка логов ===
function clearLogs() {
  CacheService.getUserCache().remove("logs_ultra_v3_fixed");
  log("🧹 Логи очищены");
}

// === Статистика производительности ===
function getPerformanceStats() {
  const stats = {
    callStatistics: CALL_STATISTICS,
    cacheInfo: {
      productsInMemory: PRODUCTS_INDEX ? Object.keys(PRODUCTS_INDEX).length : 0,
      cacheAge: CACHE_TIMESTAMP ? Math.round((Date.now() - CACHE_TIMESTAMP) / 1000) : 0,
      memoryUsage: getMemoryUsage(),
      lastDataHash: LAST_SHEET_DATA_HASH
    },
    activeTimers: Object.keys(PERFORMANCE_METRICS).length
  };
  
  return stats;
}

// === Показать отчёт о производительности ===
function showPerformanceReport() {
  const stats = getPerformanceStats();
  
  let report = ["📊 ДЕТАЛЬНЫЙ ОТЧЁТ О ПРОИЗВОДИТЕЛЬНОСТИ:", ""];
  
  // Статистика кеша
  report.push("🧠 СОСТОЯНИЕ КЕША:");
  report.push(`• Товаров в памяти: ${stats.cacheInfo.productsInMemory}`);
  report.push(`• Возраст кеша: ${stats.cacheInfo.cacheAge}с`);
  report.push(`• Использование памяти: ${stats.cacheInfo.memoryUsage}KB`);
  report.push(`• Хеш данных: ${stats.cacheInfo.lastDataHash || 'Неизвестен'}`);
  report.push("");
  
  // Статистика вызовов
  report.push("⏱️ СТАТИСТИКА ОПЕРАЦИЙ:");
  const sortedStats = Object.entries(stats.callStatistics)
    .sort(([,a], [,b]) => b.totalTime - a.totalTime);
  
  for (const [operation, data] of sortedStats.slice(0, 10)) {
    report.push(`• ${operation}:`);
    report.push(`  - Вызовов: ${data.count}`);
    report.push(`  - Среднее время: ${data.avgTime.toFixed(1)}мс`);
    report.push(`  - Диапазон: ${data.minTime}-${data.maxTime}мс`);
    report.push(`  - Общее время: ${data.totalTime}мс`);
  }
  
  if (sortedStats.length > 10) {
    report.push(`... и ещё ${sortedStats.length - 10} операций`);
  }
  
  const reportText = report.join('\n');
  SpreadsheetApp.getUi().alert(reportText);
  log("📊 Отчёт о производительности показан пользователю");
}

// === Расширенный тест производительности ===
function advancedPerformanceTest() {
  log("🚀 Запуск расширенного теста производительности...");
  
  const testResults = [];
  
  // Тест 1: Загрузка кеша
  const loadStart = Date.now();
  const products = getAllProductsUltraOptimized();
  const loadTime = Date.now() - loadStart;
  testResults.push(`Загрузка кеша: ${loadTime}мс`);
  
  // Тест 2: Поиск существующих товаров
  const existingBarcodes = Object.keys(products).slice(0, 20);
  let searchTime = 0;
  let found = 0;
  
  for (const barcode of existingBarcodes) {
    const start = Date.now();
    if (findProductUltraFast(barcode)) {
      found++;
    }
    searchTime += Date.now() - start;
  }
  
  testResults.push(`Поиск ${existingBarcodes.length} товаров: ${searchTime}мс (${found} найдено)`);
  testResults.push(`Средняя скорость поиска: ${(searchTime/existingBarcodes.length).toFixed(2)}мс/товар`);
  
  // Тест 3: Поиск несуществующих товаров
  const nonExistentBarcodes = ['FAKE001', 'FAKE002', 'FAKE003', 'FAKE004', 'FAKE005'];
  let notFoundTime = 0;
  
  for (const barcode of nonExistentBarcodes) {
    const start = Date.now();
    findProductUltraFast(barcode);
    notFoundTime += Date.now() - start;
  }
  
  testResults.push(`Поиск несуществующих товаров: ${notFoundTime}мс`);
  
  // Тест 4: Нагрузочный тест
  const stressTestStart = Date.now();
  for (let i = 0; i < 100; i++) {
    const randomBarcode = existingBarcodes[i % existingBarcodes.length];
    findProductUltraFast(randomBarcode);
  }
  const stressTime = Date.now() - stressTestStart;
  testResults.push(`Нагрузочный тест (100 поисков): ${stressTime}мс`);
  
  // Общая статистика
  const totalTime = Date.now() - loadStart;
  testResults.push(`Общее время теста: ${totalTime}мс`);
  testResults.push(`Товаров в кеше: ${Object.keys(products).length}`);
  testResults.push(`Использование памяти: ${getMemoryUsage()}KB`);
  
  const report = ["🎯 РАСШИРЕННЫЙ ТЕСТ ПРОИЗВОДИТЕЛЬНОСТИ:", "", ...testResults].join('\n');
  SpreadsheetApp.getUi().alert(report);
  log(`📊 Расширенный тест завершён за ${totalTime}мс`);
}

// === Управление кешами ===
function clearAllCaches() {
  const timerId = startTimer('clearAllCaches');
  
  try {
    // Очищаем память
    PRODUCTS_CACHE = null;
    PRODUCTS_INDEX = null;
    CACHE_TIMESTAMP = 0;
    LAST_SHEET_DATA_HASH = null;
    CALL_STATISTICS = {};
    
    // Очищаем CacheService
    const cache = CacheService.getUserCache();
    cache.remove("products_ultra_v6_fixed");
    
    // Очищаем JSON файл
    const file = DriveApp.getFileById(CACHE_FILE_ID);
    file.setContent('{"products": {}, "timestamp": 0, "hash": "cleared"}');
    
    endTimer(timerId);
    SpreadsheetApp.getUi().alert("✅ Все кеши и статистика очищены");
    log("🧹 Полная очистка всех кешей выполнена");
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    SpreadsheetApp.getUi().alert(`❌ Ошибка очистки: ${error.message}`);
  }
}

function forceRefreshCache() {
  const timerId = startTimer('forceRefresh');
  
  try {
    log("🔄 Принудительное обновление всех кешей...");
    
    // Сбрасываем все кеши
    PRODUCTS_CACHE = null;
    PRODUCTS_INDEX = null;
    CACHE_TIMESTAMP = 0;
    LAST_SHEET_DATA_HASH = null;
    
    // Загружаем свежие данные
    const products = getAllProductsUltraOptimized();
    
    endTimer(timerId, `Обновлено ${Object.keys(products).length} товаров`);
    SpreadsheetApp.getUi().alert(`✅ Принудительное обновление завершено: ${Object.keys(products).length} товаров`);
    log(`✅ Принудительное обновление завершено: ${Object.keys(products).length} товаров`);
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    const errorMsg = `❌ Ошибка обновления: ${error.message}`;
    SpreadsheetApp.getUi().alert(errorMsg);
    log(errorMsg);
  }
}

// === Тест поиска товара ===
function testFind() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("Введите штрих-код для ультра-быстрого теста:");
  if (res.getSelectedButton() !== ui.Button.OK) return;
  
  const barcode = res.getResponseText();
  if (!barcode) return;

  const startTime = Date.now();
  const product = findProductUltraFast(barcode);
  const searchTime = Date.now() - startTime;
  
  if (product) {
    ui.alert(`✅ УЛЬТРА-БЫСТРЫЙ ПОИСК!\nВремя: ${searchTime}мс\n\nТовар: "${product.name}"\nЦена: ${product.price}\nСтрока: ${product.row}`);
  } else {
    ui.alert(`❌ Не найден за ${searchTime}мс\n\nВозможно, нужно обновить кеш?`);
  }
  
  log(`🔍 Тест поиска "${barcode}": ${searchTime}мс, результат: ${product ? 'найден' : 'не найден'}`);
}

// === Показать логи ===
function showLogs() {
  const html = HtmlService.createHtmlOutputFromFile('Logs')
    .setTitle('📋 Логи (Исправленная ультра-версия)')
    .setWidth(800)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Детальные логи производительности');
}

// Алиасы для совместимости
const refreshProductsCache = forceRefreshCache;
const performanceTest = advancedPerformanceTest;
const showCacheStats = showPerformanceReport;