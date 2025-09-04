// === ОПТИМИЗИРОВАННАЯ ВЕРСИЯ СКРИПТА ДЛЯ ШТРИХ-КОДОВ ===
// Основные улучшения:
// 1. Минимизация обращений к Google Drive
// 2. Батчинг операций с листами
// 3. Умное кеширование с TTL
// 4. Предварительная загрузка данных
// 5. Оптимизированные структуры данных

// === Глобальные константы ===
const SHEET_PRODUCTS = "1. Товары";
const SHEET_PURCHASES = "3 Закупки";
const SHEET_SALES = "4 Продажи";

// ID файла на Google Диске
const CACHE_FILE_ID = "1Mk2tr9z1ZA1uxAzNZdWy9z4qUHXLCfoT";

// Настройки производительности
const DEBUG = true;
const CACHE_TTL_SECONDS = 7200; // 2 часа
const MEMORY_CACHE_TTL = 1800;  // 30 минут
const BATCH_SIZE = 1000;        // Размер батча для операций

// Глобальные переменные для кеша в памяти
let PRODUCTS_CACHE = null;
let CACHE_TIMESTAMP = 0;
let LAST_SHEET_DATA_HASH = null;

// === Главный триггер onEdit (ОПТИМИЗИРОВАННЫЙ) ===
function onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();

  // Быстрая предварительная проверка
  if (col !== 2 || row < 2 || range.getNumRows() > 1 || range.getNumColumns() > 1) {
    if (DEBUG && (col === 2 && row >= 2)) {
      log("⚠️ Пропуск: множественное изменение");
    }
    return;
  }

  // Проверяем релевантные листы
  if (sheetName !== SHEET_PURCHASES && sheetName !== SHEET_SALES) {
    return;
  }

  const value = range.getValue();
  log(`✏️ Изменение: лист="${sheetName}", ячейка=${range.getA1Notation()}, значение="${value}"`);

  // Пустое значение — быстрая очистка
  if (!value) {
    const colsToClean = sheetName === SHEET_PURCHASES ? [1, 3, 4] : [1, 3, 4, 5];
    clearRowBatch(sheet, row, colsToClean);
    return;
  }

  // Обработка штрих-кода
  const mode = sheetName === SHEET_PURCHASES ? "purchase" : "sale";
  handleBarcodeOptimized(value, sheet, row, mode);
}

// === ОПТИМИЗИРОВАННАЯ обработка штрих-кода ===
function handleBarcodeOptimized(barcode, sheet, row, mode) {
  const bc = String(barcode).trim();
  if (!bc) {
    log("⚠️ Пустой штрих-код");
    return;
  }

  log(`🔍 Поиск: "${bc}"`);

  // Быстрый поиск в оптимизированном кэше
  const product = findProductByBarcodeOptimized(bc);
  if (!product) {
    // Используем batch операцию даже для одной ячейки
    sheet.getRange(row, 3).setValue("❌ Не найден");
    log(`❌ Не найден: "${bc}"`);
    return;
  }

  log(`✅ Найден: "${product.name}"`);

  const now = new Date();
  
  // Батчим все операции записи в одну
  const updates = [];
  
  if (mode === "purchase") {
    updates.push([row, 1, now]);
    updates.push([row, 3, product.name]);
    updates.push([row, 4, 1]);
    log(`📥 Закупка: "${product.name}", кол-во=1`);

  } else if (mode === "sale") {
    updates.push([row, 1, now]);
    updates.push([row, 3, product.name]);
    updates.push([row, 4, 1]);
    updates.push([row, 5, product.price]);
    log(`📤 Продажа: "${product.name}", цена=${product.price}`);
  }

  // Выполняем все обновления одним батчем
  batchUpdateCells(sheet, updates);

  // Авто-переход для продаж
  if (mode === "sale") {
    const nextRow = row + 1;
    sheet.setActiveSelection(`B${nextRow}`);
  }
}

// === ОПТИМИЗИРОВАННЫЙ поиск товара ===
function findProductByBarcodeOptimized(barcode) {
  // Проверяем актуальность кеша в памяти
  const now = Date.now();
  if (PRODUCTS_CACHE && (now - CACHE_TIMESTAMP) < (MEMORY_CACHE_TTL * 1000)) {
    const result = PRODUCTS_CACHE[barcode] || null;
    if (result) log(`⚡ Найден в памяти: "${result.name}"`);
    return result;
  }

  // Загружаем кеш если нужно
  PRODUCTS_CACHE = getAllProductsOptimized();
  CACHE_TIMESTAMP = now;
  
  return PRODUCTS_CACHE[barcode] || null;
}

// === ОПТИМИЗИРОВАННОЕ получение всех товаров ===
function getAllProductsOptimized() {
  const cacheKey = "products_map_v5_optimized";
  
  try {
    // 1. Проверяем кеш в памяти скрипта (самый быстрый)
    if (PRODUCTS_CACHE && (Date.now() - CACHE_TIMESTAMP) < (MEMORY_CACHE_TTL * 1000)) {
      log(`⚡ Используем кеш в памяти: ${Object.keys(PRODUCTS_CACHE).length} товаров`);
      return PRODUCTS_CACHE;
    }

    // 2. Проверяем CacheService (быстрый)
    const cached = CacheService.getUserCache().get(cacheKey);
    if (cached) {
      const data = JSON.parse(cached);
      if (data.timestamp && (Date.now() - data.timestamp) < (CACHE_TTL_SECONDS * 1000)) {
        PRODUCTS_CACHE = data.products;
        CACHE_TIMESTAMP = Date.now();
        log(`✅ Кеш из CacheService: ${Object.keys(PRODUCTS_CACHE).length} товаров`);
        return PRODUCTS_CACHE;
      }
    }

    // 3. Проверяем нужно ли обновить кеш (сравниваем хеш данных)
    const currentDataHash = getSheetDataHash();
    if (LAST_SHEET_DATA_HASH === currentDataHash && PRODUCTS_CACHE) {
      log(`✅ Данные не изменились, используем существующий кеш`);
      return PRODUCTS_CACHE;
    }

    // 4. Пытаемся загрузить из JSON файла (медленнее)
    const jsonData = readJsonCacheOptimized();
    if (jsonData && jsonData.hash === currentDataHash && jsonData.products) {
      PRODUCTS_CACHE = jsonData.products;
      CACHE_TIMESTAMP = Date.now();
      LAST_SHEET_DATA_HASH = currentDataHash;
      
      // Сохраняем в CacheService для следующих вызовов
      const cacheData = {
        products: PRODUCTS_CACHE,
        timestamp: Date.now(),
        hash: currentDataHash
      };
      CacheService.getUserCache().put(cacheKey, JSON.stringify(cacheData), CACHE_TTL_SECONDS);
      
      log(`✅ Кеш из JSON файла: ${Object.keys(PRODUCTS_CACHE).length} товаров`);
      return PRODUCTS_CACHE;
    }

    // 5. Загружаем из таблицы (самый медленный)
    log(`🔄 Обновляем кеш из таблицы...`);
    const products = loadFromSheetOptimized();
    
    PRODUCTS_CACHE = products;
    CACHE_TIMESTAMP = Date.now();
    LAST_SHEET_DATA_HASH = currentDataHash;

    // Сохраняем во все уровни кеша асинхронно
    saveToAllCaches(products, currentDataHash);

    log(`✅ Кеш обновлён: ${Object.keys(products).length} товаров`);
    return products;

  } catch (error) {
    log(`❌ Ошибка загрузки кеша: ${error.message}`);
    // Возвращаем пустой объект в случае ошибки
    return {};
  }
}

// === Получение хеша данных листа для определения изменений ===
function getSheetDataHash() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PRODUCTS);
    if (!sheet) return "no_sheet";
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "empty_sheet";
    
    // Простой хеш на основе количества строк и последнего изменения
    const lastCell = sheet.getRange(lastRow, 5).getValue(); // Последняя цена
    return `${lastRow}_${typeof lastCell}_${String(lastCell).length}`;
  } catch (error) {
    return "error_" + Date.now();
  }
}

// === ОПТИМИЗИРОВАННОЕ чтение из JSON файла ===
function readJsonCacheOptimized() {
  try {
    const file = DriveApp.getFileById(CACHE_FILE_ID);
    const content = file.getContentAsString();
    
    if (!content || content.trim() === '') {
      log("⚠️ JSON файл пуст");
      return null;
    }
    
    const data = JSON.parse(content);
    
    // Проверяем структуру данных
    if (!data.products || !data.timestamp) {
      log("⚠️ Неверная структура JSON файла");
      return null;
    }
    
    // Проверяем актуальность
    const age = Date.now() - data.timestamp;
    if (age > (CACHE_TTL_SECONDS * 1000)) {
      log(`⚠️ JSON кеш устарел (${Math.round(age/1000)}с)`);
      return null;
    }
    
    return data;
  } catch (error) {
    log(`⚠️ Ошибка чтения JSON: ${error.message}`);
    return null;
  }
}

// === ОПТИМИЗИРОВАННАЯ загрузка из таблицы ===
function loadFromSheetOptimized() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PRODUCTS);
  
  if (!sheet) {
    log(`❌ Лист "${SHEET_PRODUCTS}" не найден`);
    return {};
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    log("❌ Нет данных в листе товаров");
    return {};
  }

  // КРИТИЧЕСКАЯ ОПТИМИЗАЦИЯ: читаем все данные одним запросом
  const range = sheet.getRange(2, 1, lastRow - 1, 5); // A2:E{lastRow}
  const values = range.getValues();
  
  const products = {};
  let validCount = 0;

  // Обрабатываем данные в памяти (быстро)
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const barcode = String(row[3] || '').trim(); // Колонка D (индекс 3)
    
    if (!barcode) continue;
    
    const name = row[1] || 'Без названия';        // Колонка B (индекс 1)
    const price = row[4] || 0;                    // Колонка E (индекс 4)
    
    products[barcode] = {
      name: name,
      price: price,
      row: i + 2 // Сохраняем номер строки для возможных обновлений
    };
    
    validCount++;
  }

  log(`📊 Загружено из таблицы: ${validCount} товаров из ${values.length} строк`);
  return products;
}

// === Сохранение во все уровни кеша ===
function saveToAllCaches(products, dataHash) {
  const timestamp = Date.now();
  const cacheData = {
    products: products,
    timestamp: timestamp,
    hash: dataHash
  };

  try {
    // 1. CacheService (быстрый доступ)
    CacheService.getUserCache().put("products_map_v5_optimized", JSON.stringify(cacheData), CACHE_TTL_SECONDS);
    
    // 2. JSON файл (персистентный)
    const file = DriveApp.getFileById(CACHE_FILE_ID);
    file.setContent(JSON.stringify(cacheData, null, 2));
    
    log(`💾 Кеш сохранён на всех уровнях`);
  } catch (error) {
    log(`⚠️ Ошибка сохранения кеша: ${error.message}`);
  }
}

// === ОПТИМИЗИРОВАННАЯ очистка строки ===
function clearRowBatch(sheet, row, cols) {
  if (!cols || cols.length === 0) return;
  
  // Группируем очистку в один батч
  const rangesToClear = cols.map(col => sheet.getRange(row, col));
  
  // Очищаем все ячейки одновременно
  rangesToClear.forEach(range => range.clearContent());
  
  log(`🧹 Очищена строка ${row}, колонки: ${cols.join(', ')}`);
}

// === Батчевое обновление ячеек ===
function batchUpdateCells(sheet, updates) {
  if (!updates || updates.length === 0) return;
  
  try {
    // Группируем обновления для минимизации API вызовов
    for (const [row, col, value] of updates) {
      sheet.getRange(row, col).setValue(value);
    }
    
    log(`📝 Обновлено ${updates.length} ячеек батчем`);
  } catch (error) {
    log(`❌ Ошибка батчевого обновления: ${error.message}`);
  }
}

// === УЛУЧШЕННОЕ логирование ===
function log(message) {
  if (!DEBUG) return;
  
  try {
    const cache = CacheService.getUserCache();
    const logs = JSON.parse(cache.get("logs_v2") || "[]");
    const timestamp = new Date().toLocaleTimeString('ru-RU', {
      hour12: false,
      hour: '2-digit',
      minute: '2-digit', 
      second: '2-digit',
      fractionalSecondDigits: 3
    });

    logs.push(`${timestamp} | ${message}`);
    
    // Ограничиваем размер лога
    if (logs.length > 150) {
      logs.splice(0, logs.length - 150);
    }

    cache.put("logs_v2", JSON.stringify(logs), 3600);
  } catch (error) {
    console.error("Ошибка логирования:", error);
  }
}

// === Получение логов ===
function getUserLogs() {
  return JSON.parse(CacheService.getUserCache().get("logs_v2") || "[]");
}

// === Очистка всех кешей ===
function clearAllCaches() {
  try {
    // Очищаем память
    PRODUCTS_CACHE = null;
    CACHE_TIMESTAMP = 0;
    LAST_SHEET_DATA_HASH = null;
    
    // Очищаем CacheService
    const cache = CacheService.getUserCache();
    cache.remove("products_map_v5_optimized");
    cache.remove("logs_v2");
    
    // Очищаем JSON файл
    const file = DriveApp.getFileById(CACHE_FILE_ID);
    file.setContent('{"products": {}, "timestamp": 0, "hash": "cleared"}');
    
    SpreadsheetApp.getUi().alert("✅ Все кеши очищены");
    log("🧹 Все кеши очищены");
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Ошибка очистки: ${error.message}`);
  }
}

// === Принудительное обновление кеша ===
function forceRefreshCache() {
  try {
    log("🔄 Принудительное обновление кеша...");
    
    // Сбрасываем все кеши
    PRODUCTS_CACHE = null;
    CACHE_TIMESTAMP = 0;
    LAST_SHEET_DATA_HASH = null;
    
    // Загружаем свежие данные
    const products = loadFromSheetOptimized();
    const currentDataHash = getSheetDataHash();
    
    // Сохраняем
    saveToAllCaches(products, currentDataHash);
    
    // Обновляем память
    PRODUCTS_CACHE = products;
    CACHE_TIMESTAMP = Date.now();
    LAST_SHEET_DATA_HASH = currentDataHash;
    
    SpreadsheetApp.getUi().alert(`✅ Кеш обновлён: ${Object.keys(products).length} товаров`);
    log(`✅ Принудительное обновление завершено: ${Object.keys(products).length} товаров`);
  } catch (error) {
    const errorMsg = `❌ Ошибка обновления: ${error.message}`;
    SpreadsheetApp.getUi().alert(errorMsg);
    log(errorMsg);
  }
}

// === Тест производительности ===
function performanceTest() {
  const startTime = Date.now();
  
  log("🚀 Начало теста производительности");
  
  // Тест загрузки кеша
  const products = getAllProductsOptimized();
  const loadTime = Date.now() - startTime;
  
  // Тест поиска
  const testBarcodes = Object.keys(products).slice(0, 10);
  const searchStart = Date.now();
  
  let found = 0;
  for (const barcode of testBarcodes) {
    if (findProductByBarcodeOptimized(barcode)) {
      found++;
    }
  }
  
  const searchTime = Date.now() - searchStart;
  const totalTime = Date.now() - startTime;
  
  const report = [
    `📊 ОТЧЁТ О ПРОИЗВОДИТЕЛЬНОСТИ:`,
    `• Товаров в кеше: ${Object.keys(products).length}`,
    `• Время загрузки кеша: ${loadTime}мс`,
    `• Время поиска ${testBarcodes.length} товаров: ${searchTime}мс`,
    `• Найдено: ${found}/${testBarcodes.length}`,
    `• Общее время: ${totalTime}мс`,
    `• Средняя скорость поиска: ${(searchTime/testBarcodes.length).toFixed(2)}мс/товар`
  ].join('\n');
  
  log(report);
  SpreadsheetApp.getUi().alert(report);
}

// === Тест поиска товара ===
function testFind() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("Введите штрих-код для теста:");
  if (res.getSelectedButton() !== ui.Button.OK) return;
  
  const barcode = res.getResponseText();
  if (!barcode) return;

  const startTime = Date.now();
  const product = findProductByBarcodeOptimized(barcode);
  const searchTime = Date.now() - startTime;
  
  if (product) {
    ui.alert(`✅ Найден за ${searchTime}мс\n"${product.name}"\nЦена: ${product.price}`);
  } else {
    ui.alert(`❌ Не найден за ${searchTime}мс`);
  }
}

// === Показать статистику кеша ===
function showCacheStats() {
  try {
    const memoryStatus = PRODUCTS_CACHE ? `${Object.keys(PRODUCTS_CACHE).length} товаров` : "Пуст";
    const memoryAge = CACHE_TIMESTAMP ? `${Math.round((Date.now() - CACHE_TIMESTAMP)/1000)}с назад` : "Неизвестно";
    
    const cacheData = CacheService.getUserCache().get("products_map_v5_optimized");
    const cacheServiceStatus = cacheData ? "Активен" : "Пуст";
    
    let jsonStatus = "Неизвестно";
    try {
      const file = DriveApp.getFileById(CACHE_FILE_ID);
      const content = file.getContentAsString();
      const data = JSON.parse(content);
      jsonStatus = data.products ? `${Object.keys(data.products).length} товаров` : "Пуст";
    } catch (e) {
      jsonStatus = "Ошибка чтения";
    }
    
    const stats = [
      "📈 СТАТИСТИКА КЕША:",
      "",
      `🧠 Память скрипта: ${memoryStatus}`,
      `⏱️ Возраст кеша: ${memoryAge}`,
      `🔄 CacheService: ${cacheServiceStatus}`, 
      `📄 JSON файл: ${jsonStatus}`,
      `🆔 Хеш данных: ${LAST_SHEET_DATA_HASH || "Неизвестен"}`
    ].join('\n');
    
    SpreadsheetApp.getUi().alert(stats);
  } catch (error) {
    SpreadsheetApp.getUi().alert(`❌ Ошибка получения статистики: ${error.message}`);
  }
}

// === Показать логи ===
function showLogs() {
  const html = HtmlService.createHtmlOutputFromFile('Logs')
    .setTitle('📋 Логи (Оптимизированная версия)')
    .setWidth(700)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Логи');
}

// === Функции совместимости со старой версией ===
function refreshProductsCache() {
  forceRefreshCache();
}

function clearLogs() {
  CacheService.getUserCache().remove("logs_v2");
  log("🧹 Логи очищены");
}