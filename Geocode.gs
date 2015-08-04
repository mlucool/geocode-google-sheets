/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 * Usage: Select a continuous column of address and select Add-Ons->Geocode Selected Cells
 * Encoding will put the lat/lng in the columns labled Lat/Lng
 * Based off of: https://vilimpoc.org/blog/2013/07/11/google-spreadsheet-geocoding-macro/
 */


/*
 * Short: Use Google's geocoding to convert addresses to GPS coordinates in a Google spreadsheet.
 * Tip: Select a continuous column of address and click "Add-Ons->Geocode Selected Cells". Encoding will puts the results in the columns labeled "Lat" and "Lng".

 See Add-Ons->Instructions for more details.
 */

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item.
 *
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 *
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [
        {
            name: "Geocode Selected Cells",
            functionName: "geocodeSelectedCells"
        },
        {
            name: "Instructions",
            functionName: "geocodeSelectedCellsHelp"
        }
    ];
    sheet.addMenu("Add-ons", entries);
}

function geocodeSelectedCellsHelp() {
    // Display a modal dialog box with custom HtmlService content.
    var htmlOutput = HtmlService
        .createHtmlOutputFromFile("GeoHelpDialog")
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(425)
        .setHeight(175);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Geocode Instructions');
}

function showPrompt(missingColumn) {
    var ui = SpreadsheetApp.getUi(); // Same variations.
    var word = (missingColumn[missingColumn.length - 1] == 's') ? "them" : "it";
    var result = ui.alert(
        'Add Columns',
        'You are missing ' + missingColumn + ' for results would you like me to create ' + word + ' for you?',
        ui.ButtonSet.YES_NO);

    // Process the user's response.
    return result == ui.Button.YES;
}

function geocodeSelectedCells() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var addresses = sheet.getActiveRange();

    // We expect only the column to be encoded selected
    if (addresses.getNumColumns() == 0) {
        Browser.msgBox("Please select a address/location column to encode");
    } else if (addresses.getNumColumns() != 1) {
        Browser.msgBox("Please select only one address/location column to encode");
        return;
    }

    // Find where to put results
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var headerValues = headerRange.getValues();
    var latColumn = -1;
    var lngColumn = -1;
    var row = null;
    for (row in headerValues) {
        for (var col in headerValues[row]) {
            if (headerValues[row][col].toLowerCase() == "lat") {
                latColumn = parseInt(col) + 1;
            } else if (headerValues[row][col].toLowerCase() == "lng") {
                lngColumn = parseInt(col) + 1;
            }
        }
    }
    if (latColumn == -1 && lngColumn == -1) {
        if (!showPrompt("the Lat/Lng columns")) {
            return;
        }
        latColumn = sheet.getLastColumn() + 1;
        lngColumn = latColumn + 1;
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.insertColumnAfter(latColumn);
        setValue(sheet, 1, lngColumn, 'Lng');
        setValue(sheet, 1, latColumn, 'Lat');
    } else if (latColumn == -1) {
        if (!showPrompt("a Lat column")) {
            return;
        }
        sheet.insertColumnAfter(lngColumn);
        latColumn = lngColumn + 1;
        setValue(sheet, 1, latColumn, 'Lat');
    } else if (lngColumn == -1) {
        if (!showPrompt("a Lng column")) {
            return;
        }
        sheet.insertColumnAfter(latColumn);
        lngColumn = latColumn + 1;
        setValue(sheet, 1, lngColumn, 'Lng');
        setValue(sheet, 1, latColumn, 'Lat');
    }

    // Let's Encode
    var geocoder = Maps.newGeocoder().setRegion('de');
    row = 1;
    // Skip header if selected
    if (addresses.getRow() == 1) {
        ++row;
    }

    var cell = null;
    for (var r = row; r <= addresses.getNumRows(); ++r) {
        cell = addresses.getCell(r, 1);
        if (!sheet.getRange(cell.getRow(), latColumn).isBlank() || !sheet.getRange(cell.getRow(), lngColumn).isBlank()) {
            var ui = SpreadsheetApp.getUi();
            var response = ui.alert('The results cells lat/lng are not empty. Do you wish to override?', ui.ButtonSet.YES_NO);
            // Process the user's response.
            if (response == ui.Button.YES) {
                break;
            } else {
                return;
            }
        }
    }

    for (row; row <= addresses.getNumRows(); ++row) {
        cell = addresses.getCell(row, 1);
        var address = cell.getValue();

        // Geocode the address and plug the lat, lng pair into the
        if (address == "") {
            continue;
        }
        var location = geocoder.geocode(address);

        // Only change cells if geocoder seems to have gotten a
        // valid response.
        if (location.status == 'OK') {
            lat = location["results"][0]["geometry"]["location"]["lat"];
            lng = location["results"][0]["geometry"]["location"]["lng"];

            setValue(sheet, cell.getRow(), latColumn, lat);
            setValue(sheet, cell.getRow(), lngColumn, lng);
        } else {
            setValue(sheet, cell.getRow(), latColumn, location.status);
            setValue(sheet, cell.getRow(), lngColumn, location.status);
        }
    }
}

function setValue(sheet, row, column, val) {
    sheet.getRange(row, column).setValue(val);
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
    onOpen(e);
}