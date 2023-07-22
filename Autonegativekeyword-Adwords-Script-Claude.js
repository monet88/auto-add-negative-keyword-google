// Constants
const SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1w8co5s_qE-1dyyHb1ftQUWRH5MP-zFeDLB9cvpvmydg/edit#gid=0";

// Helper functions
const getSheetByName = name => spreadsheet.getSheetByName(name);

const getRangeValues = range => range.getValues()[0]; 

const addNegativesToList = (listName, negatives) => {
  const list = getNegativeList(listName);
  if (list) {
    list.addNegativeKeywords(negatives);
  } else {
    Logger.log(`Negative list ${listName} not found!`);
  }
}

// Main function
function main() {
  
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  
  const sheetData = spreadsheet.getSheets()
    .map(sheet => ({
      timestamp: getSheetTimestamp(sheet),
      name: sheet.getName()
    }))
    .sort((a, b) => a.timestamp - b.timestamp);
    
  sheetData.forEach(({name}) => {
    processSheet(getSheetByName(name));
  });

  Logger.log("Finished processing sheets");

}

// Refactored processSheet function
function processSheet(sheet) {

  const settings = getSettings(sheet);
  
  const adGroups = getAdGroups(sheet);

  adGroups.forEach(({name, negatives, queries}) => {
    
    const keywords = getKeywords(sheet, queries);
    
    const newNegatives = getNewNegatives(settings, keywords);

    addNegativesToList(negatives, newNegatives);

  });

  setTimestamp(sheet);

}

// Get sheet timestamp helper
function getSheetTimestamp(sheet) {
  return sheet.getRange(1, 7).getValue() || new Date().setDate(new Date().getDate() - 1000);
}

// Additional helper functions...

// Main function invocation
main();