var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1w8co5s_qE-1dyyHb1ftQUWRH5MP-zFeDLB9cvpvmydg/edit#gid=0";
var SS = SpreadsheetApp.openByUrl(SPREADSHEET_URL);

function main() {
  var sheets = SS.getSheets();
  var timesArray = sheets.map((sheet, i) => {
    var timestamp = sheet.getRange(1, 7).getValue() || new Date().setDate(new Date().getDate() - (1000+i));
    return {timestamp, name: sheet.getName()};
  }).sort((a, b) => a.timestamp - b.timestamp);

  timesArray.forEach(time => {
    var sheet = SS.getSheetByName(time.name);
    processSheet(sheet);
  });

  Logger.log("Finished");
}

function processSheet(sheet) {
  var SETTINGS = getSettings(sheet);
  var adGroupsData = getAdGroupsData(sheet);
  
  adGroupsData.forEach(data => {
    var keywords = getKeywords(sheet, data.columnNumber);
    var negs = getNegatives(SETTINGS, data, keywords);
    addNegatives(SETTINGS, data, negs);
  });

  sheet.getRange(1, 7).setValue(new Date());
}

function getSettings(sheet) {
  var settingsRange = sheet.getRange("A2:F2");
  var settingsValues = settingsRange.getValues()[0];
  
  return {
    CAMPAIGN_NAME: settingsValues[0],
    MIN_QUERY_CLICKS: settingsValues[1],
    MAX_QUERY_CONVERSIONS: settingsValues[2],
    DATE_RANGE: settingsValues[3],
    NEGATIVE_MATCH_TYPE: settingsValues[4],
    CAMPAIGN_LEVEL_QUERIES: settingsValues[5] == "Yes"
  };
}

function getAdGroupsData(sheet) {
  var adGroups = [];
  var col = 2;
  
  while(sheet.getRange(4, col).getValue()){
    adGroups.push({
      name: sheet.getRange(5, col).getValue(),
      negativeList: sheet.getRange(6, col).getValue(),
      minMatches: sheet.getRange(4, col).getValue(),
      columnNumber: col
    });
    col++;
  }
  
  return adGroups;
}

function getKeywords(sheet, columnNumber) {
  var row = 7;
  var keywords = [];
  
  while(sheet.getRange(row, columnNumber).getValue()){
    keywords.push(String(sheet.getRange(row, columnNumber).getValue()).toLowerCase());
    row++;
  }
  
  return keywords;
}

function getNegatives(SETTINGS, adGroupData, keywords) {
  var query = buildQuery(SETTINGS, adGroupData);
  var report = AdWordsApp.report(query);
  var rows = report.rows();
  var negatives = [];
  
  while(rows.hasNext()){
    var row = rows.next();
    var query = row.Query;
    var matches = keywords.filter(keyword => query.includes(keyword)).length;
    
    if(matches < adGroupData.minMatches){
      negatives.push(query);
    }
  }
  
  return negatives;
}

function buildQuery(SETTINGS, adGroupData) {
  var query =  "SELECT Query FROM SEARCH_QUERY_PERFORMANCE_REPORT WHERE CampaignName = '" + SETTINGS["CAMPAIGN_NAME"] + "'";
  
  if(SETTINGS["MIN_QUERY_CLICKS"]){
    query += " AND Clicks > " + SETTINGS["MIN_QUERY_CLICKS"];
  }

  if(SETTINGS["MAX_QUERY_CONVERSIONS"]){
    query += " AND Conversions < " + SETTINGS["MAX_QUERY_CONVERSIONS"];
  }

  if(!SETTINGS["CAMPAIGN_LEVEL_QUERIES"]){
    query += ' AND AdGroupName = "' + adGroupData.name + '"';
  }

  if(SETTINGS["DATE_RANGE"] != "ALL_TIME"){
    query += " DURING " + SETTINGS["DATE_RANGE"];
  }
  
  return query;
}

function addNegatives(SETTINGS, adGroupData, negatives) {
  var adGroupIterator = AdWordsApp.adGroups()
    .withCondition("Name = '" + adGroupData.name + "'")
    .withCondition("CampaignName = '" + SETTINGS["CAMPAIGN_NAME"] + "'")
    .get();
  
  if(adGroupIterator.hasNext()){
    var adGroup = adGroupIterator.next();
    
    negatives.forEach(negative => {
      var neg = addMatchType(negative, SETTINGS);
      
      if(adGroupData.negativeList){
        addNegativeToList(adGroupData.negativeList, neg);
      } else {
        adGroup.createNegativeKeyword(neg);
      }
    });
  } else {
    Logger.log("AdGroup '" + adGroupData.name + "' in Campaign '" + SETTINGS["CAMPAIGN_NAME"] + "' not found in the account. Check the AdGroup name is correct in the sheet.");
  }
}

function addMatchType(word, SETTINGS){
    var matchType = SETTINGS["NEGATIVE_MATCH_TYPE"].toLowerCase();
    
    if(matchType == "broad"){
      return word.trim();
    } else if(matchType == "bmm"){
      return word.split(" ").map(word => "+" + word).join(" ").trim();
    } else if(matchType == "phrase"){
      return '"' + word.trim() + '"';
    } else if(matchType == "exact"){
      return '[' + word.trim() + ']';
    } else {
      throw("Error: Match type not recognised. Please provide one of Broad, BMM, Exact or Phrase");
    }
  }
function addNegativeToList(negativeListName, neg){
  var listIter = AdWordsApp.negativeKeywordLists().withCondition("Name = '"+ negativeListName +"'").get();
  
  if(listIter.hasNext()){
    var negativeList = listIter.next();
    negativeList.addNegativeKeywords([neg]);
  } else {
    Logger.log("The shared negative '"+negativeListName+"' list can't be found");
  }
}
  