/* eslint-disable no-undef */
/* eslint-disable @typescript-eslint/no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

/**
 * Writes the event source id to the document when ExecuteFunction runs.
 * @param event {Office.AddinCommands.Event}
 */
function readValue() {
  console.log("reached insode write table read data ");
  Excel.run(async function (context) {
    // Get the active worksheet
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Load all the used range of the worksheet
    var range = sheet.getUsedRange();
    range.load("values");

    // Synchronize the document state by executing the queued commands
    await context.sync();
    // Access the data in the range
    var data = range.values;
    // Do something with the data...
    console.log(data);
  }).catch(function (error) {
    console.log("Error: " + error);
  });

  // Excel.run(async function (context) {
  //   // Get the active worksheet
  //   var sheet = context.workbook.worksheets.getActiveWorksheet();

  //   // Load the tables collection
  //   var tables = sheet.tables.load("items/name");

  //   // Synchronize the document state by executing the queued commands
  //   await context.sync();
  //   // Loop through each table in the worksheet
  //   tables.items.forEach(async function (table) {
  //     // Load the table's properties and dataBodyRange of columns
  //     await table.load(
  //       "name, showHeaders, showTotals, style, columns/items/name, columns/items/dataBodyRange/format, columns/items/dataBodyRange/rowCount, columns/items/dataBodyRange/columnCount, columns/items/dataBodyRange/values, values"
  //     );

  //     // Synchronize the document state by executing the queued commands
  //     return context.sync().then(async function () {
  //       // Access the data and formatting of the table
  //       var tableName = table.name;
  //       var hasHeaders = table.showHeaders;
  //       var hasTotals = table.showTotals;
  //       var style = table.style;

  //       var columns = table.columns.items;
  //       var columnData = [];
  //       for (var i = 0; i < columns.length; i++) {
  //         var column = columns[i];
  //         console.log(column);
  //         var columnName = column.name;
  //         var dataBodyRange = column.dataBodyRange;
  //         var dataBodyValues = dataBodyRange?.values;
  //         var rowCount = dataBodyRange?.rowCount;
  //         var columnCount = dataBodyRange?.columnCount;
  //         var dataBodyFormat = dataBodyRange?.format;
  //         columnData.push({
  //           name: columnName,
  //           data: dataBodyValues,
  //           rowCount: rowCount,
  //           columnCount: columnCount,
  //           format: dataBodyFormat,
  //         });
  //       }

  //       // Do something with the table data and formatting...
  //       console.log("Table Name: " + tableName);
  //       console.log("Has Headers: " + hasHeaders);
  //       console.log("Has Totals: " + hasTotals);
  //       console.log("Style: " + style);
  //       console.log("Column Data: " + JSON.stringify(columnData));
  //     });
  //   });
  // }).catch(function (error) {
  //   console.log("Error: " + error);
  // });
  //
  //
  //
  //
  // Excel.run(function (context) {
  //   // Get the current worksheet
  //   var sheet = context.workbook.worksheets.getActiveWorksheet();

  //   // Get all the tables in the worksheet
  //   var tables = sheet.tables;

  //   // Load the table data
  //   tables.load("items/name, columns/items/name, rows/items");

  //   // Execute the request
  //   return context.sync().then(function () {
  //     // Loop through each table and display its data
  //     tables.items.forEach(function (table) {
  //       console.log("Table name: " + table.name);

  //       // Loop through each row in the table
  //       table.rows.items.forEach(function (row) {
  //         // Loop through each column in the row
  //         table.columns.items.forEach(function (column) {
  //           var cell = row.getCell(column.name);
  //           console.log(column.name + ": " + cell.values[0][0]);
  //         });
  //       });
  //     });
  //   });
  // }).catch(function (error) {
  //   console.log("Error: " + error);
  // });

  //   Excel.run(function (context) {
  //     // Get the current worksheet
  //     console.log("reached insode read table data ");
  //     var sheet = context.workbook.worksheets.getActiveWorksheet();

  //     // Get the first table in the worksheet
  //     var table = sheet.tables.getItemAt(0);

  //     // Load the table data and column properties
  //     table.load("name, columns/items/name, columns/items/count, columns/items/items");
  //     console.log(table.columns.items, " table.columns.items.");
  //     // Execute the request
  //     return context.sync().then(function () {
  //       // Loop through each column in the table
  //       table.columns.items.forEach(function (column) {
  //         console.log("Column name: " + column.name);
  //         console.log("Number of cells in the column: " + column.count);

  //         // Loop through each cell in the column
  //         column.items.forEach(function (cell) {
  //           // Check if the cell is part of a merged range
  //           if (cell.isMerged) {
  //             var mergedRange = cell.getBoundingRect();
  //             console.log("This cell is part of a merged range: " + mergedRange.address);
  //           }
  //         });
  //       });
  //     });
  //   }).catch(function (error) {
  //     console.log("Error: " + error);
  //   });

  //   Excel.run(async function (context) {
  //     // Get the current worksheet
  //     var sheet = context.workbook.worksheets.getActiveWorksheet();

  //     // Get the used range of the worksheet
  //     var range = sheet.getUsedRange();
  //     range.load("values");

  //     // Execute the request
  //     await context.sync();
  //     // Loop through each row in the range
  //     for (var i = 0; i < range.values.length; i++) {
  //       var row = range.values[i];
  //       // Loop through each cell in the row
  //       for (var j = 0; j < row.length; j++) {
  //         var cellValue = row[j];
  //         console.log(cellValue);
  //       }
  //     }
  //   }).catch(function (error) {
  //     console.log("Error: " + error);
  //   });
}

function writeValue() {
  console.log("reached insode write table value");
  var data = [
    ["Alice", 25, "Female"],
    ["Name", "Age", "Gender"],
    ["Bob", 30, "Male"],
    ["Bob", 30, "Male"],
    ["Charlie", 40, "Male"],
  ];

  // eslint-disable-next-line no-undef
  Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("A1:C5");
    range.values = data;
    range.format.autofitColumns();
    sheet.tables.add(range, true);
    return context.sync();
  }).catch(function (error) {
    console.log(error);
  });
}
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
Office.actions.associate("writeValue", writeValue);
Office.actions.associate("readValue", readValue);
