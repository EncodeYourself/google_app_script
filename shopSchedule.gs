function preprocess() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const range = sheet.getRange("D2:AH11");
    const values = range.getValues();

    for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
            let formatted_value = values[row][col].toString().trim().replace("ё","е").replace("й", "и").replace(" ", " ").replace("ъ","ь");
            values[row][col] = formatted_value;
            }}
        range.setValues(values);
}

function set_address() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    let current_cell = sheet.getRange("A14").activate();
    let range = sheet.getRange("B2:B11");
    let values = range.getValues();
    values.forEach(function(value) {
        current_cell.setValue(value);
        current_cell = current_cell.offset(0,2);
        current_cell.activate();
        }) 
}

function get_random_color() {
    // Формирование случайного цвета.
    const letters = '0123456789ABCDEF';
    let color = '#';
    for (let i = 0; i < 6; i++) {
        color += letters[Math.floor(Math.random() * 16)];
        }
    return color;
}

function set_color() {
    // Использовал словарь для хранения цветов для каждого сотрудника
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getRange("D2:AH11");
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
        
        let colorDict = {};

        for (let j = 0; j < values[i].length; j++) {

        let value = values[i][j];

        if (!(value in colorDict)) {

                colorDict[value] = get_random_color();
            } 
        sheet.getRange(i + 2, j + 4).setBackground(colorDict[value]);
            }
        }
        }

function count_values() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getRange("D2:AH11");
    const values = range.getValues();
    let new_column = sheet.getRange("A15");

    for (let i = 0; i < values.length; i++) {
        let count_dict = {};

        for (let j = 0; j < values[i].length; j++) {
            let value = values[i][j];

            if (count_dict[value]) {
            count_dict[value]++;
            } else {
            count_dict[value] = 1;
    }
}

    let current_cell = new_column.activate();
    const keys = Object.keys(count_dict);
    const dvalues = Object.values(count_dict);
    for (let i = 0; i < keys.length; i++) {
        current_cell.setValue(keys[i]);
        current_cell = current_cell.offset(0, 1);
        current_cell.setValue(dvalues[i]);
        current_cell = current_cell.offset(1, -1);
        }
        new_column = new_column.offset(0, 2);
    }}


function main() {
    preprocess();
    set_address();
    count_values();
    set_color();
}