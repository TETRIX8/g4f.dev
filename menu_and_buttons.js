// === ДОПОЛНЕНИЕ К СКРИПТУ: МЕНЮ И КНОПКИ ===
// Добавьте этот код в конец вашего основного скрипта

// === Создание пользовательского меню при открытии таблицы ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Создаём главное меню "Штрих-коды"
  const menu = ui.createMenu('🔍 Штрих-коды')
    
    // Основные функции
    .addItem('📋 Показать логи', 'showLogs')
    .addItem('📊 Отчёт о производительности', 'showPerformanceReport')
    .addItem('🧪 Тест производительности', 'advancedPerformanceTest')
    .addSeparator()
    
    // Управление кешем
    .addSubMenu(ui.createMenu('💾 Управление кешем')
      .addItem('🔄 Обновить кеш', 'forceRefreshCache')
      .addItem('📈 Статистика кеша', 'showCacheStats') 
      .addItem('🧹 Очистить все кеши', 'clearAllCaches')
      .addItem('🗑️ Очистить логи', 'clearLogs'))
    
    // Диагностика
    .addSubMenu(ui.createMenu('🔧 Диагностика')
      .addItem('🔍 Тест поиска товара', 'testFind')
      .addItem('⚡ Предзагрузка данных', 'preloadData')
      .addItem('🛡️ Проверка целостности', 'validateDataIntegrity')
      .addItem('📱 Статус системы', 'showSystemStatus'))
    
    // Настройки
    .addSubMenu(ui.createMenu('⚙️ Настройки')
      .addItem('🐛 Включить/выключить отладку', 'toggleDebugMode')
      .addItem('📊 Настройки производительности', 'showPerformanceSettings')
      .addItem('💡 Справка', 'showHelp'))
    
    .addToUi();
  
  log("📋 Пользовательское меню создано");
}

// === Функции для меню ===

// Показать статистику кеша (алиас для совместимости)
function showCacheStats() {
  showPerformanceReport();
}

// Переключение режима отладки
function toggleDebugMode() {
  // Поскольку DEBUG - константа, показываем текущий статус
  const currentStatus = DEBUG ? "включена" : "выключена";
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Режим отладки',
    `Отладка сейчас ${currentStatus}.\n\nДля изменения отредактируйте константу DEBUG в коде скрипта.`,
    ui.ButtonSet.OK
  );
  
  log(`🐛 Проверка режима отладки: ${currentStatus}`);
}

// Показать настройки производительности
function showPerformanceSettings() {
  const settings = [
    `🔧 ТЕКУЩИЕ НАСТРОЙКИ ПРОИЗВОДИТЕЛЬНОСТИ:`,
    ``,
    `• Отладка: ${DEBUG ? 'включена' : 'выключена'}`,
    `• Логирование производительности: ${PERFORMANCE_LOGGING ? 'включено' : 'выключено'}`,
    `• TTL кеша: ${CACHE_TTL_SECONDS}с (${Math.round(CACHE_TTL_SECONDS/3600)}ч)`,
    `• TTL памяти: ${MEMORY_CACHE_TTL}с (${Math.round(MEMORY_CACHE_TTL/60)}мин)`,
    `• Размер батча: ${BATCH_SIZE}`,
    `• Предзагрузка: ${PRELOAD_ON_STARTUP ? 'включена' : 'выключена'}`,
    ``,
    `📝 Для изменения настроек отредактируйте константы в коде скрипта.`
  ].join('\n');
  
  SpreadsheetApp.getUi().alert(settings);
  log("⚙️ Показаны настройки производительности");
}

// Показать справку
function showHelp() {
  const help = [
    `💡 СПРАВКА ПО СИСТЕМЕ ШТРИХ-КОДОВ`,
    ``,
    `🎯 ОСНОВНОЕ ИСПОЛЬЗОВАНИЕ:`,
    `• Введите штрих-код в колонку B (строка ≥2)`,
    `• Система автоматически найдёт товар и заполнит данные`,
    `• Для продаж автоматически переходит к следующей строке`,
    ``,
    `📊 МОНИТОРИНГ:`,
    `• "Показать логи" - детальные логи всех операций`,
    `• "Отчёт о производительности" - статистика скорости`,
    `• "Тест производительности" - проверка скорости работы`,
    ``,
    `💾 УПРАВЛЕНИЕ:`,
    `• "Обновить кеш" - принудительное обновление данных`,
    `• "Очистить кеши" - полная очистка при проблемах`,
    ``,
    `🔧 ПРИ ПРОБЛЕМАХ:`,
    `1. Проверьте логи`,
    `2. Обновите кеш`,
    `3. Проверьте статистику производительности`,
    ``,
    `📞 Версия: Ультра-оптимизированная с логированием`
  ].join('\n');
  
  SpreadsheetApp.getUi().alert(help);
  log("💡 Показана справка пользователю");
}

// Статус системы
function showSystemStatus() {
  const timerId = startTimer('systemStatus');
  
  try {
    const stats = getPerformanceStats();
    const cacheSize = stats.cacheInfo.productsInMemory;
    const cacheAge = stats.cacheInfo.cacheAge;
    const memoryUsage = stats.cacheInfo.memoryUsage;
    
    // Проверяем доступность компонентов
    let sheetStatus = "❌ Недоступен";
    let jsonFileStatus = "❌ Недоступен";
    
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_PRODUCTS);
      if (sheet && sheet.getLastRow() > 1) {
        sheetStatus = `✅ Доступен (${sheet.getLastRow() - 1} товаров)`;
      }
    } catch (e) {
      sheetStatus = `❌ Ошибка: ${e.message}`;
    }
    
    try {
      const file = DriveApp.getFileById(CACHE_FILE_ID);
      const content = file.getContentAsString();
      if (content && content.length > 10) {
        jsonFileStatus = `✅ Доступен (${Math.round(content.length/1024)}KB)`;
      }
    } catch (e) {
      jsonFileStatus = `❌ Ошибка: ${e.message}`;
    }
    
    // Определяем общий статус системы
    let systemHealth = "🟢 Отлично";
    if (cacheSize === 0) systemHealth = "🔴 Критично";
    else if (cacheAge > 14400) systemHealth = "🟡 Требует внимания";
    else if (memoryUsage > 10000) systemHealth = "🟡 Высокая нагрузка";
    
    const status = [
      `🖥️ СТАТУС СИСТЕМЫ ШТРИХ-КОДОВ`,
      ``,
      `🎯 Общее состояние: ${systemHealth}`,
      ``,
      `💾 КЕШИРОВАНИЕ:`,
      `• Товаров в памяти: ${cacheSize}`,
      `• Возраст кеша: ${cacheAge}с (${Math.round(cacheAge/60)}мин)`,
      `• Использование памяти: ${memoryUsage}KB`,
      ``,
      `📊 КОМПОНЕНТЫ:`,
      `• Лист товаров: ${sheetStatus}`,
      `• JSON файл: ${jsonFileStatus}`,
      `• Активных таймеров: ${stats.activeTimers}`,
      ``,
      `⏱️ ПРОИЗВОДИТЕЛЬНОСТЬ:`,
      `• Всего операций: ${Object.keys(stats.callStatistics).length}`,
      `• Последняя активность: ${getLastActivityTime()}`,
      ``,
      `🔧 РЕКОМЕНДАЦИИ:`,
      `${getSystemRecommendations(stats)}`
    ].join('\n');
    
    endTimer(timerId, 'Статус системы показан');
    SpreadsheetApp.getUi().alert(status);
    log("🖥️ Показан статус системы");
    
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    SpreadsheetApp.getUi().alert(`❌ Ошибка получения статуса: ${error.message}`);
  }
}

// Получить время последней активности
function getLastActivityTime() {
  if (!CALL_STATISTICS || Object.keys(CALL_STATISTICS).length === 0) {
    return "Неизвестно";
  }
  
  let lastActivity = 0;
  for (const stats of Object.values(CALL_STATISTICS)) {
    if (stats.lastCall > lastActivity) {
      lastActivity = stats.lastCall;
    }
  }
  
  if (lastActivity === 0) return "Неизвестно";
  
  const diff = Date.now() - lastActivity;
  if (diff < 60000) return `${Math.round(diff/1000)}с назад`;
  if (diff < 3600000) return `${Math.round(diff/60000)}мин назад`;
  return `${Math.round(diff/3600000)}ч назад`;
}

// Получить рекомендации по системе
function getSystemRecommendations(stats) {
  const recommendations = [];
  
  if (stats.cacheInfo.productsInMemory === 0) {
    recommendations.push("• Обновите кеш товаров");
  }
  
  if (stats.cacheInfo.cacheAge > 14400) {
    recommendations.push("• Кеш устарел, рекомендуется обновление");
  }
  
  if (stats.cacheInfo.memoryUsage > 10000) {
    recommendations.push("• Высокое использование памяти");
  }
  
  if (Object.keys(stats.callStatistics).length === 0) {
    recommendations.push("• Система не использовалась");
  }
  
  // Анализируем производительность
  const slowOperations = Object.entries(stats.callStatistics)
    .filter(([name, data]) => data.avgTime > 100)
    .map(([name]) => name);
  
  if (slowOperations.length > 0) {
    recommendations.push(`• Медленные операции: ${slowOperations.join(', ')}`);
  }
  
  if (recommendations.length === 0) {
    return "Система работает оптимально! 🎉";
  }
  
  return recommendations.join('\n');
}

// Проверка целостности данных
function validateDataIntegrity() {
  const timerId = startTimer('dataIntegrityCheck');
  log("🛡️ Запуск проверки целостности данных...");
  
  const issues = [];
  let warnings = 0;
  let errors = 0;
  
  try {
    // Проверка 1: Доступность листа товаров
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PRODUCTS);
    
    if (!sheet) {
      issues.push(`❌ Лист "${SHEET_PRODUCTS}" не найден`);
      errors++;
    } else {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        issues.push("⚠️ Лист товаров пуст");
        warnings++;
      } else {
        issues.push(`✅ Лист товаров: ${lastRow - 1} строк данных`);
      }
    }
    
    // Проверка 2: Структура данных в листе
    if (sheet && sheet.getLastRow() >= 2) {
      const headers = sheet.getRange(1, 1, 1, 5).getValues()[0];
      const expectedHeaders = ['A', 'B', 'C', 'D', 'E']; // Примерные заголовки
      
      // Проверяем несколько строк данных
      const sampleData = sheet.getRange(2, 1, Math.min(5, sheet.getLastRow() - 1), 5).getValues();
      let emptyBarcodes = 0;
      let emptyNames = 0;
      
      for (const row of sampleData) {
        if (!row[3] || String(row[3]).trim() === '') emptyBarcodes++;
        if (!row[1] || String(row[1]).trim() === '') emptyNames++;
      }
      
      if (emptyBarcodes > 0) {
        issues.push(`⚠️ Найдено ${emptyBarcodes} строк с пустыми штрих-кодами`);
        warnings++;
      }
      
      if (emptyNames > 0) {
        issues.push(`⚠️ Найдено ${emptyNames} строк с пустыми названиями`);
        warnings++;
      }
    }
    
    // Проверка 3: JSON файл кеша
    try {
      const file = DriveApp.getFileById(CACHE_FILE_ID);
      const content = file.getContentAsString();
      
      if (!content || content.trim() === '') {
        issues.push("⚠️ JSON файл кеша пуст");
        warnings++;
      } else {
        const data = JSON.parse(content);
        if (data.products && Object.keys(data.products).length > 0) {
          issues.push(`✅ JSON кеш: ${Object.keys(data.products).length} товаров`);
        } else {
          issues.push("⚠️ JSON кеш не содержит товаров");
          warnings++;
        }
      }
    } catch (e) {
      issues.push(`❌ Ошибка доступа к JSON файлу: ${e.message}`);
      errors++;
    }
    
    // Проверка 4: Состояние кеша в памяти
    if (PRODUCTS_INDEX && Object.keys(PRODUCTS_INDEX).length > 0) {
      issues.push(`✅ Кеш в памяти: ${Object.keys(PRODUCTS_INDEX).length} товаров`);
      
      // Проверяем актуальность
      const cacheAge = Date.now() - CACHE_TIMESTAMP;
      if (cacheAge > MEMORY_CACHE_TTL * 1000) {
        issues.push(`⚠️ Кеш в памяти устарел (${Math.round(cacheAge/1000)}с)`);
        warnings++;
      }
    } else {
      issues.push("⚠️ Кеш в памяти пуст");
      warnings++;
    }
    
    // Проверка 5: CacheService
    const cacheData = CacheService.getUserCache().get("products_ultra_v6");
    if (cacheData) {
      try {
        const data = JSON.parse(cacheData);
        if (data.products) {
          issues.push(`✅ CacheService: ${Object.keys(data.products).length} товаров`);
        }
      } catch (e) {
        issues.push("❌ Ошибка чтения CacheService");
        errors++;
      }
    } else {
      issues.push("⚠️ CacheService пуст");
      warnings++;
    }
    
    // Проверка 6: Производительность
    if (CALL_STATISTICS && Object.keys(CALL_STATISTICS).length > 0) {
      const slowOps = Object.entries(CALL_STATISTICS)
        .filter(([name, stats]) => stats.avgTime > 500)
        .map(([name, stats]) => `${name} (${stats.avgTime.toFixed(0)}мс)`);
      
      if (slowOps.length > 0) {
        issues.push(`⚠️ Медленные операции: ${slowOps.join(', ')}`);
        warnings++;
      } else {
        issues.push("✅ Производительность в норме");
      }
    }
    
    // Формируем отчёт
    let status = "🟢 Отлично";
    if (errors > 0) status = "🔴 Критичные ошибки";
    else if (warnings > 2) status = "🟡 Требует внимания";
    else if (warnings > 0) status = "🟡 Незначительные проблемы";
    
    const report = [
      `🛡️ ПРОВЕРКА ЦЕЛОСТНОСТИ ДАННЫХ`,
      ``,
      `📊 Общий статус: ${status}`,
      `• Ошибок: ${errors}`,
      `• Предупреждений: ${warnings}`,
      ``,
      `📋 ДЕТАЛИ:`,
      ...issues.map(issue => `${issue}`),
      ``,
      `🕒 Проверка завершена: ${new Date().toLocaleTimeString('ru-RU')}`
    ].join('\n');
    
    endTimer(timerId, `Ошибок: ${errors}, предупреждений: ${warnings}`);
    SpreadsheetApp.getUi().alert(report);
    log(`🛡️ Проверка целостности завершена: ${errors} ошибок, ${warnings} предупреждений`);
    
  } catch (error) {
    endTimer(timerId, `Критическая ошибка: ${error.message}`);
    SpreadsheetApp.getUi().alert(`❌ Критическая ошибка проверки: ${error.message}`);
    log(`❌ Критическая ошибка проверки целостности: ${error.message}`);
  }
}

// Предзагрузка данных (публичная функция)
function preloadData() {
  const timerId = startTimer('manualPreload');
  
  try {
    log("🚀 Ручная предзагрузка данных...");
    SpreadsheetApp.getUi().alert("🔄 Запуск предзагрузки данных...\nЭто может занять несколько секунд.");
    
    // Очищаем кеш для свежих данных
    PRODUCTS_CACHE = null;
    PRODUCTS_INDEX = null;
    CACHE_TIMESTAMP = 0;
    
    // Загружаем данные
    const products = getAllProductsUltraOptimized();
    const productCount = Object.keys(products).length;
    
    endTimer(timerId, `Предзагружено ${productCount} товаров`);
    
    SpreadsheetApp.getUi().alert(`✅ Предзагрузка завершена!\n\nЗагружено товаров: ${productCount}\nСистема готова к работе.`);
    log(`✅ Ручная предзагрузка завершена: ${productCount} товаров`);
    
  } catch (error) {
    endTimer(timerId, `Ошибка: ${error.message}`);
    SpreadsheetApp.getUi().alert(`❌ Ошибка предзагрузки: ${error.message}`);
    log(`❌ Ошибка ручной предзагрузки: ${error.message}`);
  }
}