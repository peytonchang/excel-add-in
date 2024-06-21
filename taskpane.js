/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

  }
});


async function jsonToSheet(jsonData) {
  await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Ensure jsonData is an array and not empty
      if (!Array.isArray(jsonData) || jsonData.length === 0) {
          document.getElementById("messageArea").innerText = "Invalid JSON data.";
          return;
      }

      // Extract headers
      const headers = Object.keys(jsonData[0]);
      const headerRow = [headers];

      // Extract rows
      const rows = jsonData.map(item => headers.map(header => item[header]));

      // Combine headers and rows
      const data = headerRow.concat(rows);

      // Determine the range to write to (starting at A1)
      const range = sheet.getRangeByIndexes(0, 0, data.length, headers.length);
      range.values = data;

      await context.sync();
      document.getElementById("messageArea").innerText = "Data written to sheet successfully.";
  }).catch(function(error) {
      console.error("Error:", error);
      document.getElementById("messageArea").innerText = "Error: " + error.message;
  });
}

document.getElementById("capture-btn").addEventListener("click", function() {
    captureData();
});

document.getElementById("paste-button").addEventListener("click", function() {
    pasteData();
})

document.getElementById("btn3").addEventListener("click", function() {
  const jsonData = document.getElementById("jsonInput").value;
  try {
      const parsedData = JSON.parse(jsonData);
      jsonToSheet(parsedData);
  } catch (error) {
      document.getElementById("messageArea").innerText = "Invalid JSON format: " + error.message;
  }
});

document.getElementById("btnGenerateSQL").addEventListener("click", function() {
  generateSQLInsertStatements();
});

async function generateSQLInsertStatements() {
  await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');  // Load the sheet name
      const range = sheet.getUsedRange();
      range.load('values');
      
      const tableDefSheet = context.workbook.worksheets.getItem("Table Def");
      const tableDefRange = tableDefSheet.getUsedRange();
      tableDefRange.load('values');

      await context.sync();

      const data = range.values;
      const tableDefData = tableDefRange.values;
      
      if (data.length === 0) {
          document.getElementById("sqlOutput").value = "The sheet is empty.";
          return;
      }

      if (tableDefData.length === 0) {
          document.getElementById("sqlOutput").value = "The Table Def sheet is empty.";
          return;
      }

      const headers = data[0];
      const rows = data.slice(1);

      // Specify the table name explicitly
      const tableName = "ref_condition";  // Replace with your table name

      // Parse table definition
      const tableDef = parseTableDef(tableDefData);

      // Validate and generate SQL INSERT statements
      try {
          let sqlStatements = generateInsertSQL(tableDef, tableName, headers, rows);
          document.getElementById("sqlOutput").value = sqlStatements;
      } catch (error) {
          console.error("Validation error:", error);
          document.getElementById("sqlOutput").value = error.message;
      }
  }).catch(function(error) {
      console.error("Error:", error);
      document.getElementById("sqlOutput").value = "Error: " + error.message;
  });
}

function parseTableDef(tableDefData) {
  const tableDef = {};
  tableDefData.slice(1).forEach(row => {
      const tableName = row[0];
      if (!tableDef[tableName]) {
          tableDef[tableName] = {
              columns: [],
              types: {},
              nullable: {}
          };
      }
      tableDef[tableName].columns.push(row[2]);
      tableDef[tableName].types[row[2]] = row[3]; // Assuming the 'Type' column is the fourth column (index 3)
      tableDef[tableName].nullable[row[2]] = row[4].toLowerCase() === 'yes';
  });
  return tableDef;
}

function excelDateToJSDate(serial) {
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  const isoDateString = date.toISOString();
  return isoDateString.substring(0, 10) + ' ' + isoDateString.substring(11, 19);
}

function generateInsertSQL(tableDef, tableName, headers, rows) {
  if (!tableDef[tableName]) {
      throw new Error(`Table ${tableName} not found in Table Def sheet.`);
  }
  const columns = tableDef[tableName].columns;
  const types = tableDef[tableName].types;
  const nullable = tableDef[tableName].nullable;

  // Regex dictionary for type validation
  const regexes = {
    'bigint': /^-?(0|[1-9]\d*)$/,
    'bigint(20)': /^-?(0|[1-9]\d*)$/,
    'bit': /^[01]$/,
    'bit(1)': /^[01]$/,
    'datetime': /^\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}:\d{2}$/,
    'decimal': /^-?\d+(\.\d+)?$/,
    'int': /^-?(0|[1-9]\d*)$/,
    'int(11)': /^-?(0|[1-9]\d*)$/,
    'smallint(6)': /^-?(0|[1-9]\d*)$/,
    'varchar': /(^.*)$/,
    'varchar(10)': /^.{0,10}$/,
    'varchar(20)': /^.{0,20}$/,
    'varchar(255)': /^.{0,255}$/,
    'varchar(40)': /^.{0,40}$/,
    'varchar(4096)': /^.{0,4096}/,
    'varchar(512)': /^.{0,512}$/,
    'varchar(60)': /^.{0,60}$/,
    'varchar(80)': /^.{0,80}$/
      // Add other types and their regex patterns as needed
  };

  // Columns to ignore
  const ignoreColumns = ['id', 'version', 'created_date_time', 'created_by_user_id', 'modified_date_time', 'modified_by_user_id'];

  // Filter out ignored columns
  const filteredColumns = columns.filter(col => !ignoreColumns.includes(col));
  const filteredHeaders = headers.filter(header => !ignoreColumns.includes(header));

  let sqlStatements = `INSERT INTO ${tableName} (${filteredColumns.join(", ")})\nVALUES\n`;
  let errorMessages = "";

  const valueStatements = rows.map((row, rowIndex) => {
      const values = filteredColumns.map((col, colIndex) => {
          let value = row[headers.indexOf(col)];
          const type = types[col];

          // Convert Excel date serial to human-readable format for datetime fields
          if (type === 'datetime' && typeof value === 'number') {
              value = excelDateToJSDate(value);
          }

          // Validate value against its type
          if (value === "" && !nullable[col]) {
              errorMessages += `Column ${col} at row ${rowIndex + 2} cannot be null.\n`;
          }
          if (value !== "" && regexes[type] && !regexes[type].test(value)) {
              errorMessages += `Type mismatch at row: ${rowIndex + 2}, column: ${col}, expected type: ${type}, value: ${value}\n`;
          }

          return value === "" ? "NULL" : `'${value}'`;
      }).join(", ");
      return `(${values})`;
  });

  if (errorMessages) {
      throw new Error(errorMessages);
  }

  sqlStatements += valueStatements.join(",\n") + ";";

  return sqlStatements;
}

async function captureData() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getActiveCell();
            range.load("values");
            await context.sync();

            const cellValue = range.values[0][0];
            document.getElementById("input-box").value = cellValue;
        });
    } catch (error) {
        console.error(error);
    }
}

async function pasteData() {
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getActiveCell();
            const textBoxValue = document.getElementById("input-box").value;
            range.values = [[textBoxValue]];
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}