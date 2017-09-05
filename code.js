var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var menuSheet = spreadsheet.getSheetByName('Menu');
var optionsSheet = spreadsheet.getSheetByName('Opciones');
var sheet = spreadsheet.getActiveSheet();
var lastRows = [];

var LAST_COLUMN_IN_MENU = 'J';
var LAST_ROW_IN_MENU = 3;
 var START_ROW_OPTIONS = 2;
var START_ROW_MENU = 3;

var OPTIONS_COLUMNS = ['A', 'B', 'C', 'D', 'E'];
var MENU_COLUMNS = ['B', 'C', 'D', 'E', 'F', 'G', 'H'];

var MENU_ITEMS = 5;

var OPTIONS = {
    'Breakfast': 'A:A',
    'Morning Snack': 'B:B',
    'Lunch': 'C:C',
    'Afternoon Snack': 'D:D',
    'Diner': 'E:E'
};

function onOpen() {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var menuSheet = spreadsheet.getSheetByName('Menu');
    var optionsSheet = spreadsheet.getSheetByName('Opciones');
    var sheet = spreadsheet.getActiveSheet();

    var menuEntries = [
        {name: "Generate Menu", functionName: "generateMenu"}
    ];
    spreadsheet.addMenu("Food Menu", getLastRowsFromMenu);
}

function getLastRowsFromMenu() {
    optionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Opciones");


    var keyRange;
    var range;
    var lastRow;

    for (var key in OPTIONS) {
        if (OPTIONS.hasOwnProperty(key)) {
            keyRange = OPTIONS[key];
            range = optionsSheet.getRange(keyRange).getValues();
            lastRow = range.filter(String).length;
            if (lastRow !== undefined && lastRow !== null) {
                lastRows.push(lastRow);
            }
        }
    }

    menuSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
    var lastRowCell;
    var lastRowNumber = LAST_ROW_IN_MENU;

    for (var i = 0, lastRowValue; (lastRowValue = lastRows[i]); i++) {
        lastRowCell = menuSheet.getRange(LAST_COLUMN_IN_MENU + lastRowNumber);

        lastRowCell.setValue(lastRowValue);
        lastRowNumber++;
    }
    randomize();
    setDailyMealFormula();
}

function randomize() {
    var currentCell, menu_row, lastRowInColumn;
    var generator = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Generator");

    for (var i = 0; i < MENU_COLUMNS.length; i++) {
        var menu_column = MENU_COLUMNS[i]; //B
        for (var j = 0; j < OPTIONS_COLUMNS.length; j++) {
            lastRowInColumn = lastRows[j] - 1;
            menu_row = START_ROW_MENU + j;
            currentCell = generator.getRange(menu_column + menu_row);
            currentCell.setFormula("=RANDBETWEEN(K3*" + START_ROW_OPTIONS + ";" + lastRowInColumn + ")");
        }
    }
}

function setDailyMealFormula() {
    var menu_entries = START_ROW_MENU + MENU_ITEMS - 1;
    var currentCell, menu_row, lastRowInColumn, option_column, menu_column, menuRange;

    menuSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");

    for (var i = 0; i < MENU_COLUMNS.length; i++) {
        menu_column = MENU_COLUMNS[i]; //B

        for (var j = 0; j < OPTIONS_COLUMNS.length; j++) {
            option_column = OPTIONS_COLUMNS[j];
            // minus headers to offset rows
            lastRowInColumn = lastRows[j] - 1;
            menu_row = START_ROW_MENU + j;
            menuRange = menu_column + menu_row;

            currentCell = menuSheet.getRange(menuRange);
            currentCell.setFormula("=INDEX(Opciones!$" + option_column + "$" + START_ROW_OPTIONS + ":$" + option_column + "$" + lastRowInColumn + ";Generator!$" + menuRange + ")");
        }
    }

}
