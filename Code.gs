/** @OnlyCurrentDoc */

/**
 * Добавляет меню с командой вызова функции обновления значений служебной ячейки (для обновления вычислнений функций, ссылающихся на эту ячейку)
 *
 **/
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name: 'Обновить',
      functionName: 'refresh',
    },
  ];
  sheet.addMenu('TI', entries);
}

function refresh() {
  const updateDateRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('UPDATE_DATE').getCell(1, 1);
  if (updateDateRange != null) {
    updateDateRange.setValue(new Date());
  } else {
    SpreadsheetApp.getUi().ui.alert('You should specify the named range "UPDATE_DATE" for using this function.');
  }
}
