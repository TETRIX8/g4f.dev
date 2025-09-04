// === ДОПОЛНИТЕЛЬНЫЕ ВАРИАНТЫ КНОПОК ===
// Альтернативные способы создания кнопок в Google Sheets

// === ВАРИАНТ 1: Кнопки через рисунки (Drawing) ===
// Эти функции помогут создать красивые кнопки прямо на листе

function createDrawingButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Создаём лист для кнопок управления, если его нет
  let controlSheet;
  try {
    controlSheet = ss.getSheetByName("🎛️ Управление");
  } catch (e) {
    controlSheet = ss.insertSheet("🎛️ Управление");
    setupControlSheet(controlSheet);
  }
  
  SpreadsheetApp.getUi().alert(
    "📝 Инструкция по созданию кнопок",
    "1. Перейдите на лист '🎛️ Управление'\n" +
    "2. Используйте меню 'Вставка' → 'Рисунок'\n" +
    "3. Создайте кнопку (прямоугольник с текстом)\n" +
    "4. Нажмите на кнопку → '⋮' → 'Назначить скрипт'\n" +
    "5. Введите название функции (например: showLogs)\n\n" +
    "Готовые функции для кнопок:\n" +
    "• showLogs\n" +
    "• showPerformanceReport\n" +
    "• forceRefreshCache\n" +
    "• advancedPerformanceTest\n" +
    "• clearAllCaches",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  log("🎛️ Инструкция по созданию кнопок показана");
}

function setupControlSheet(sheet) {
  // Настраиваем лист управления
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);
  
  // Заголовок
  sheet.getRange("A1:D1").merge();
  sheet.getRange("A1").setValue("🎛️ ПАНЕЛЬ УПРАВЛЕНИЯ СИСТЕМОЙ ШТРИХ-КОДОВ");
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");
  
  // Инструкции
  const instructions = [
    ["📋 ОСНОВНЫЕ ФУНКЦИИ:", "", "", ""],
    ["Показать логи", "→ showLogs", "", ""],
    ["Отчёт производительности", "→ showPerformanceReport", "", ""],
    ["Тест производительности", "→ advancedPerformanceTest", "", ""],
    ["", "", "", ""],
    ["💾 УПРАВЛЕНИЕ КЕШЕМ:", "", "", ""],
    ["Обновить кеш", "→ forceRefreshCache", "", ""],
    ["Очистить кеши", "→ clearAllCaches", "", ""],
    ["Статистика кеша", "→ showSystemStatus", "", ""],
    ["", "", "", ""],
    ["🔧 ДИАГНОСТИКА:", "", "", ""],
    ["Тест поиска", "→ testFind", "", ""],
    ["Проверка целостности", "→ validateDataIntegrity", "", ""],
    ["Предзагрузка данных", "→ preloadData", "", ""],
    ["", "", "", ""],
    ["💡 КАК СОЗДАТЬ КНОПКУ:", "", "", ""],
    ["1. Вставка → Рисунок", "", "", ""],
    ["2. Создайте прямоугольник", "", "", ""],
    ["3. Добавьте текст", "", "", ""],
    ["4. Сохраните и закройте", "", "", ""],
    ["5. Нажмите на кнопку → ⋮", "", "", ""],
    ["6. Назначить скрипт", "", "", ""],
    ["7. Введите название функции", "", "", ""]
  ];
  
  sheet.getRange(3, 1, instructions.length, 4).setValues(instructions);
  
  // Стилизация
  sheet.getRange("A3:A13").setFontWeight("bold");
  sheet.getRange("B3:B13").setFontStyle("italic");
  
  log("🎛️ Лист управления настроен");
}

// === ВАРИАНТ 2: Кнопки через формулы с гиперссылками ===
function createHyperlinkButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet;
  
  try {
    sheet = ss.getSheetByName("🔗 Быстрые кнопки");
  } catch (e) {
    sheet = ss.insertSheet("🔗 Быстрые кнопки");
  }
  
  // Создаём кнопки через гиперссылки
  const buttons = [
    ["📋 Показать логи", '=HYPERLINK("#gid=0", "📋 Логи")'],
    ["📊 Производительность", '=HYPERLINK("#gid=0", "📊 Отчёт")'],
    ["🔄 Обновить кеш", '=HYPERLINK("#gid=0", "🔄 Обновить")'],
    ["🧪 Тест", '=HYPERLINK("#gid=0", "🧪 Тест")'],
    ["🧹 Очистить", '=HYPERLINK("#gid=0", "🧹 Очистить")']
  ];
  
  // Заголовок
  sheet.getRange("A1").setValue("🔗 БЫСТРЫЕ КНОПКИ (через гиперссылки)");
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold");
  
  // Размещаем кнопки
  for (let i = 0; i < buttons.length; i++) {
    const row = i + 3;
    sheet.getRange(row, 1).setValue(buttons[i][0]);
    sheet.getRange(row, 2).setFormula(buttons[i][1]);
    
    // Стилизация кнопок
    sheet.getRange(row, 2).setBackground("#4285f4").setFontColor("white").setFontWeight("bold");
  }
  
  // Инструкция
  sheet.getRange("A10").setValue("❗ Примечание: Гиперссылки не могут напрямую вызывать функции.");
  sheet.getRange("A11").setValue("Используйте меню '🔍 Штрих-коды' или рисунки с назначенными скриптами.");
  
  SpreadsheetApp.getUi().alert("✅ Лист с быстрыми кнопками создан!\n\nДля полноценных кнопок используйте рисунки или меню.");
  log("🔗 Лист с гиперссылками создан");
}

// === ВАРИАНТ 3: Макрос для быстрого создания кнопки ===
function createQuickButton() {
  const ui = SpreadsheetApp.getUi();
  
  // Запрашиваем название функции
  const functionResponse = ui.prompt(
    "Создание кнопки",
    "Введите название функции для кнопки:\n\n" +
    "Доступные функции:\n" +
    "• showLogs\n" +
    "• showPerformanceReport\n" +
    "• forceRefreshCache\n" +
    "• advancedPerformanceTest\n" +
    "• clearAllCaches\n" +
    "• testFind\n" +
    "• validateDataIntegrity",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (functionResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const functionName = functionResponse.getResponseText().trim();
  if (!functionName) {
    ui.alert("❌ Название функции не может быть пустым");
    return;
  }
  
  // Запрашиваем текст кнопки
  const textResponse = ui.prompt(
    "Текст кнопки",
    "Введите текст для отображения на кнопке:",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (textResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const buttonText = textResponse.getResponseText().trim() || functionName;
  
  // Инструкция по созданию
  const instruction = [
    "📝 ИНСТРУКЦИЯ ПО СОЗДАНИЮ КНОПКИ:",
    "",
    "1. Перейдите в меню 'Вставка' → 'Рисунок'",
    "2. Нажмите 'Фигуры' → выберите прямоугольник",
    "3. Нарисуйте кнопку нужного размера",
    "4. Дважды кликните на фигуру для добавления текста",
    `5. Введите текст: "${buttonText}"`,
    "6. Настройте цвет и стиль по желанию",
    "7. Нажмите 'Сохранить и закрыть'",
    "8. Кликните на созданную кнопку",
    "9. Нажмите на три точки (⋮) в правом верхнем углу",
    "10. Выберите 'Назначить скрипт'",
    `11. Введите: ${functionName}`,
    "12. Нажмите 'OK'",
    "",
    "✅ Готово! Кнопка будет вызывать функцию при нажатии."
  ].join('\n');
  
  ui.alert("Создание кнопки", instruction, ui.ButtonSet.OK);
  log(`🔘 Инструкция создания кнопки для функции ${functionName} показана`);
}

// === Функции-обёртки для кнопок (упрощённые названия) ===
function logs() { showLogs(); }
function performance() { showPerformanceReport(); }
function refresh() { forceRefreshCache(); }
function test() { advancedPerformanceTest(); }
function clear() { clearAllCaches(); }
function find() { testFind(); }
function status() { showSystemStatus(); }
function integrity() { validateDataIntegrity(); }
function preload() { preloadData(); }

// === Специальные функции для кнопок ===
function quickRefreshAndTest() {
  SpreadsheetApp.getUi().alert("🔄 Быстрое обновление и тест...");
  forceRefreshCache();
  Utilities.sleep(1000); // Пауза между операциями
  advancedPerformanceTest();
}

function emergencyReset() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "🚨 ЭКСТРЕННЫЙ СБРОС",
    "Это действие очистит все кеши и перезапустит систему.\n\nПродолжить?",
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    clearAllCaches();
    Utilities.sleep(2000);
    forceRefreshCache();
    ui.alert("✅ Экстренный сброс выполнен успешно!");
    log("🚨 Выполнен экстренный сброс системы");
  }
}

function showQuickStats() {
  const stats = getPerformanceStats();
  const quickInfo = [
    `📊 БЫСТРАЯ СТАТИСТИКА:`,
    ``,
    `• Товаров в кеше: ${stats.cacheInfo.productsInMemory}`,
    `• Возраст кеша: ${Math.round(stats.cacheInfo.cacheAge/60)}мин`,
    `• Память: ${stats.cacheInfo.memoryUsage}KB`,
    `• Операций выполнено: ${Object.keys(stats.callStatistics).length}`,
    ``,
    `Для подробного отчёта используйте кнопку "Производительность"`
  ].join('\n');
  
  SpreadsheetApp.getUi().alert(quickInfo);
  log("📊 Показана быстрая статистика");
}