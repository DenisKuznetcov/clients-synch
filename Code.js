// Обработчик события редактирования
function onEdit(e) {
  const sheet = e.source.getActiveSheet();

  // Проверяем, что редактируемый лист — "Rus. Master table"
  if (sheet.getName() === "Rus. Master table") {
    const range = e.range;
    const editedRow = range.getRow();
    const editedCol = range.getColumn();
    const editedValue = range.getValue();
    const targetSheet = e.source.getSheetByName("Eng. Transfered from master");

    // Задаем колонки, которые активируют перевод
    const triggerColumns = new Set([3, 7, 8, 11, 12, 16, 25]);
    const targetRange = targetSheet.getRange(editedRow, editedCol);
    const sourceRange = sheet.getRange(editedRow, editedCol);

    // Если значение в ячейке не пустое
    if (editedValue) {
      // Если редактируемая колонка — в списке триггерных, выполняем перевод
      if (triggerColumns.has(editedCol)) {
        const targetValue = targetRange.getValue();

        // Если целевая ячейка пуста или значение на русском, переводим
        if (!targetValue || /[\u0400-\u04FF]/.test(editedValue)) {
          const translatedText = LanguageApp.translate(editedValue, "ru", "en");
          targetRange.setValue(translatedText);
        } else {
          targetRange.setValue(editedValue);
        }
      } else {
        targetRange.setValue(editedValue);
      }

      // Копируем правила валидации данных
      const dataValidation = sourceRange.getDataValidation();
      if (dataValidation) {
        sourceRange.copyTo(targetRange);
      }

      // Копируем форматирование
      sourceRange.copyFormatToRange(targetRange.getGridId(), editedCol, editedCol, editedRow, editedRow);

    } else {
      // Если значение пустое, очищаем целевую ячейку
      targetRange.clearContent();
    }
  }
}

/**
 * Функция для синхронизации и перевода данных между листами
 * Синхронизирует все данные между листами "Rus. Master table" и "Eng. Transfered from master" и в нужных колонках переводит их/
 * Запускается по кнопке на листе "btn" либо временному триггеру
 */
function reconcileAndTranslate() {
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rus. Master table");
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Eng. Transfered from master");

  // Задаем колонки, которые активируют перевод
  const triggerColumns = new Set([3, 7, 8, 11, 12, 16, 25]);

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  const translations = new Map();

  // Проходим по всем строкам и колонкам
  for (let row = 1; row <= lastRow; row++) {
    for (let col = 1; col <= lastCol; col++) {
      const sourceRange = sourceSheet.getRange(row, col);
      const targetRange = targetSheet.getRange(row, col);

      const sourceValue = sourceRange.getValue();
      const targetValue = targetRange.getValue();

      // Если значения в ячейках не пустые
      if (sourceValue !== "" || targetValue !== "") {
        // Если колонка — в списке триггерных, выполняем перевод
        if (triggerColumns.has(col)) {
          if (!targetValue || /[\u0400-\u04FF]/.test(sourceValue)) {
            let translatedText = translations.get(sourceValue);
            // Если перевод не был сохранен, выполняем новый
            if (translatedText === undefined || translatedText === null) {
              translatedText = LanguageApp.translate(sourceValue, "ru", "en");
              translations.set(sourceValue, translatedText);
            }
            targetRange.setValue(translatedText);
          } else {
            targetRange.setValue(sourceValue);
          }
        } else {
          targetRange.setValue(sourceValue);
        }

        // Копируем dropdown с исходной "окраской"
        const dataValidation = sourceRange.getDataValidation();
        if (dataValidation) {
          sourceRange.copyTo(targetRange);
        }

        // Копируем форматирование
        sourceRange.copyFormatToRange(targetRange.getGridId(), col, col, row, row);

      } else if (targetValue) {
        // Если целевая ячейка не пуста, очищаем её
        targetSheet.getRange(row, col).clearContent();
      }
    }
  }

  // Показываем сообщение об успешной синхронизации
  try {
    SpreadsheetApp.getUi().alert(`Sheets "${sourceSheet.getName()}" and "${targetSheet.getName()}" are successfully reconciled`);
  } catch (e) {
    console.log("Couldn't call getUi() in this context");
  }
}