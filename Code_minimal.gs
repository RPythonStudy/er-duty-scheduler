/**
 * ER Shift Management System - Minimal Fast Version
 * Only reads Config and displays empty calendar
 */

const CONFIG_SHEET_NAME = 'Config';

// --- Web App Entry Point ---

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('ER 근무 신청 시스템')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- API Functions ---

/**
 * Get calendar configuration (Year, Month only)
 * ULTRA FAST - Only reads 2 cells
 */
function getCalendarConfig(reqYear, reqMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  // Create Config sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.getRange("A1:A2").setValues([["Year"], ["Month"]]);
    sheet.getRange("B1:B2").setValues([[new Date().getFullYear()], [new Date().getMonth() + 1]]);
  }
  
  // Use requested date if provided, otherwise use config
  const configYear = sheet.getRange("B1").getValue();
  const configMonth = sheet.getRange("B2").getValue();
  
  const year = reqYear ? parseInt(reqYear) : configYear;
  const month = reqMonth ? parseInt(reqMonth) : configMonth;
  
  return {
    year: year,
    month: month
  };
}
