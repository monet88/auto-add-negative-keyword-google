var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1w8co5s_qE-1dyyHb1ftQUWRH5MP-zFeDLB9cvpvmydg/edit#gid=0";
var SS = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
function main() {
  // Get all the sheets in the spreadsheet
  var sheets = SS.getSheets();

  // Create an array of objects with timestamp and sheet name
  var timesArray = sheets.map((sheet, i) => {
    // Get the timestamp from cell (1, 7) in the sheet
    // If the cell is empty, use a timestamp calculated based on the current date minus i days
    var timestamp = sheet.getRange(1, 7).getValue() || new Date().setDate(new Date().getDate() - (1000+i));
    return {timestamp, name: sheet.getName()};
  })
  // Sort the array based on the timestamp in ascending order
  .sort((a, b) => a.timestamp - b.timestamp);

  // Iterate over each object in the timesArray
  timesArray.forEach(({name}) => {
    // Get the sheet with the corresponding name
    var sheet = SS.getSheetByName(name);
    // Process the sheet
    processSheet(sheet);
  });

  // Log a message indicating that the function has finished running
  Logger.log("Finished");
}
// Refactored function to process the sheet
function processSheet(sheet) {
  // Get the settings for the sheet
  const SETTINGS = getSettings(sheet);
  
  // Get the ad groups data from the sheet
  const adGroupsData = getAdGroupsData(sheet);
  
  // Iterate through each ad group data
  adGroupsData.forEach(data => {
    // Get the keywords for the current ad group data
    const keywords = getKeywords(sheet, data.columnNumber);
    
    // Get the negatives based on the settings and keywords
    const negs = getNegatives(SETTINGS, data, keywords);
    
    // Add the negatives to the ad group
    addNegatives(SETTINGS, data, negs);
  });
  
  // Set the value of cell (1, 7) to the current date
  sheet.getRange(1, 7).setValue(new Date());
}
function getSettings(sheet) {
  // Get the range of cells containing the settings data
  var settingsRange = sheet.getRange("A2:F2");
  
  // Get the values from the settings range
  var settingsValues = settingsRange.getValues()[0];

  // Destructure the settings values into individual variables
  var [
    CAMPAIGN_NAME,
    MIN_QUERY_CLICKS,
    MAX_QUERY_CONVERSIONS,
    DATE_RANGE,
    NEGATIVE_MATCH_TYPE,
    CAMPAIGN_LEVEL_QUERIES
  ] = settingsValues;

  // Return an object containing the settings
  return {
    CAMPAIGN_NAME,
    MIN_QUERY_CLICKS,
    MAX_QUERY_CONVERSIONS,
    DATE_RANGE,
    NEGATIVE_MATCH_TYPE,
    CAMPAIGN_LEVEL_QUERIES: CAMPAIGN_LEVEL_QUERIES === "Yes"
  };
}
function getAdGroupsData(sheet) {
  // Initialize an array to store the ad groups data
  var adGroups = [];

  // Get all the data from the sheet
  var data = sheet.getDataRange().getValues();
  
  // Iterate over the ad groups data starting from index 1
  for (var i = 1; i < data[3].length; i++) {
    // Check if the ad group has a value at the current index
    if (data[3][i]) {
      // Create an object representing the ad group and push it to the adGroups array
      adGroups.push({
        name: data[4][i], // Store the name of the ad group
        negativeList: data[5][i], // Store the negative list of the ad group
        minMatches: data[3][i], // Store the minimum matches of the ad group
        columnNumber: i + 1 // Store the column number of the ad group
      });
    }
  }
  
  // Return the adGroups array
  return adGroups;
}
function getKeywords(sheet, columnNumber) {
  // Create an empty array to store the keywords
  const keywords = [];

  // Get the last row number of the sheet
  const lastRow = sheet.getLastRow();

  // Iterate through each row starting from row 7
  for (let row = 7; row <= lastRow; row++) {
    // Get the value from the specified column and row
    const value = sheet.getRange(row, columnNumber).getValue();

    // If the value is empty, exit the loop
    if (!value) {
      break;
    }

    // Convert the value to lowercase and add it to the keywords array
    keywords.push(value.toString().toLowerCase());
  }

  // Return the array of keywords
  return keywords;
}
// Refactored function to get negative queries
// Parameters:
// - SETTINGS: object containing settings for the function
// - adGroupData: data for the ad group
// - keywords: array of keywords
// Returns:
// - an array of negative queries

function getNegatives(SETTINGS, adGroupData, keywords) {
  // Build the query using the SETTINGS and adGroupData
  var query = buildQuery(SETTINGS, adGroupData);

  // Run the query using the AdWordsApp.report method
  var report = AdWordsApp.report(query);

  // Get the rows from the report
  var rows = report.rows();

  // Initialize an empty array to store the negative queries
  var negatives = [];

  // Loop through each row in the report
  while(rows.hasNext()){
    // Get the current row
    var row = rows.next();

    // Get the query from the row
    var query = row.Query;

    // Count the number of matches between the query and the keywords
    var matches = countMatches(query, keywords);

    // Check if the number of matches is less than the minimum matches required
    if(matches < adGroupData.minMatches){
      // If it is, add the query to the negatives array
      negatives.push(query);
    }
  }

  // Return the array of negative queries
  return negatives;
}
// This function takes a query string and an array of keywords as input
// It counts the number of keywords that appear in the query string
function countMatches(query, keywords) {
  // Initialize a variable to store the count of matches
  var count = 0;
  
  // Iterate over each keyword in the array
  for (var i = 0; i < keywords.length; i++) {
    // Check if the query string includes the current keyword
    if (query.includes(keywords[i])) {
      // If the keyword is found in the query, increment the count
      count++;
    }
  }
  
  // Return the final count of matches
  return count;
}
function buildQuery(SETTINGS, adGroupData) {
  // Destructure the necessary variables from the SETTINGS object
  const { CAMPAIGN_NAME, MIN_QUERY_CLICKS, MAX_QUERY_CONVERSIONS, CAMPAIGN_LEVEL_QUERIES, DATE_RANGE } = SETTINGS;

  // Destructure the adGroupData object to get the name property
  const { name } = adGroupData;

  // Initialize the query string with the SELECT statement
  let query = `SELECT Query FROM SEARCH_QUERY_PERFORMANCE_REPORT WHERE CampaignName = '${CAMPAIGN_NAME}'`;

  // Add the condition for minimum query clicks if provided
  if(MIN_QUERY_CLICKS){
    query += ` AND Clicks > ${MIN_QUERY_CLICKS}`;
  }

  // Add the condition for maximum query conversions if provided
  if(MAX_QUERY_CONVERSIONS){
    query += ` AND Conversions < ${MAX_QUERY_CONVERSIONS}`;
  }

  // Add the condition for ad group level queries if not enabled at the campaign level
  if(!CAMPAIGN_LEVEL_QUERIES){
    query += ` AND AdGroupName = '${name}'`;
  }

  // Add the condition for the date range if not set to "ALL_TIME"
  if(DATE_RANGE != "ALL_TIME"){
    query += ` DURING ${DATE_RANGE}`;
  }

  // Return the final query string
  return query;
}
function addNegatives(SETTINGS, adGroupData, negatives) {
  // Find the ad group based on its name and the campaign name
  var adGroup = AdWordsApp.adGroups()
    .withCondition("Name = '" + adGroupData.name + "'")
    .withCondition("CampaignName = '" + SETTINGS["CAMPAIGN_NAME"] + "'")
    .getFirst();

  if (adGroup) {
    // Iterate over each negative keyword
    negatives.forEach(negative => {
      // Add the match type to the negative keyword
      var neg = addMatchType(negative, SETTINGS);

      if (adGroupData.negativeList) {
        // Add the negative keyword to the negative list
        addNegativeToList(adGroupData.negativeList, neg);
      } else {
        // Add the negative keyword directly to the ad group
        adGroup.createNegativeKeyword(neg);
      }
    });
  } else {
    // Log an error message if the ad group is not found
    Logger.log(
      "AdGroup '" +
        adGroupData.name +
        "' in Campaign '" +
        SETTINGS["CAMPAIGN_NAME"] +
        "' not found in the account. Check the AdGroup name is correct in the sheet."
    );
  }
}
function addMatchType(word, SETTINGS) {
  // Get the match type from the SETTINGS object and convert it to lowercase
  const matchType = SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase();
  
  // Based on the match type, modify the word and return the modified word
  if (matchType === "broad") {
    // If the match type is "broad", simply trim the word and return it
    return word.trim();
  } else if (matchType === "bmm") {
    // If the match type is "bmm", split the word into an array of words,
    // prepend each word with a "+", join the words with spaces, trim the result, and return it
    return word.split(" ").map(word => "+" + word).join(" ").trim();
  } else if (matchType === "phrase") {
    // If the match type is "phrase", surround the word with double quotes, trim it, and return it
    return '"' + word.trim() + '"';
  } else if (matchType === "exact") {
    // If the match type is "exact", surround the word with square brackets, trim it, and return it
    return '[' + word.trim() + ']';
  } else {
    // If the match type is not recognized, throw an error with a specific message
    throw new Error("Error: Match type not recognized. Please provide one of Broad, BMM, Exact or Phrase");
  }
}

// This function adds a negative keyword to a specified negative keyword list in AdWords.
function addNegativeToList(negativeListName, neg) {
  // Find the negative keyword list with the given name
  var negativeListIterator = AdWordsApp.negativeKeywordLists()
    .withCondition("Name = '" + negativeListName + "'")
    .get();
  
  // Check if the negative keyword list exists
  if (negativeListIterator.hasNext()) {
    var negativeList = negativeListIterator.next();
    
    // Add the negative keyword to the list
    negativeList.addNegativeKeywords([neg]);
  } else {
    // Log an error message if the negative keyword list doesn't exist
    Logger.log("The shared negative '" + negativeListName + "' list can't be found");
  }
}
  