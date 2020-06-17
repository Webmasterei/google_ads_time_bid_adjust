/*
Author: Bernhard Prange
Web: https://webmasterei-prange.de
Email: info@webmasterei-prange.de
We do not provide Warranty in any way for the use of this Script.
*/

var spreadSheetUrl = "https://docs.google.com/spreadsheets/d/1E9A0bFIOAjM7V2BxP4Bv9uxk_SqpweKngWhx7u0zJgY/edit"; // Type Spreadsheet URL Here
var accountLabel = "auto_schedule";
var ss = openSpreadsheet(spreadSheetUrl);
var daysOfWeek = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]

function main() {
  var accountSelector = MccApp.accounts();
  accountSelector.withCondition("Labels CONTAINS '" + accountLabel +"'");
  accountSelector.executeInParallel('processClientAccount', null);
}

function processClientAccount() {
  var account = AdWordsApp.currentAccount();
  var sheetName = account.getName()+" - "+account.getCustomerId();
  var sheet = accountSheet(sheetName);
  if (!sheet.getRange("D1").getValue()) {
    writeSheetHeader(sheet, account)  
  }
  var rules = getRows(sheet)
  if (rules){
    iterateCampaigns(rules);
    iterateShoppingCampaigns(rules);
  }
}

function writeSheetHeader(sheet, account){
  sheet.getRange("A1").setValue("Ad Schedule");
  sheet.getRange("D1").setValue(account.getName()+" "+account.getCustomerId());
  sheet.getRange(2,1).setValue("Hour");
  for (var i = 0; i < daysOfWeek.length; i++){
    sheet.getRange(2, i+2).setValue(daysOfWeek[i])
  }
  for (var t = 0; t < 24; t++){
    sheet.getRange(t+3, 1).setValue(t)
  }
  sheet.getRange(1,1,2,8).setFontWeight("bold");
  sheet.getRange(3,1,24,1).setFontWeight("bold");
}

function openSpreadsheet(spreadSheetUrl) {
  var ss = SpreadsheetApp.openByUrl(spreadSheetUrl);
  return ss;
}

function accountSheet(name) {
  try {
    var sheet = ss.getSheetByName(name).activate();
  }
  catch(err) {
    var sheet = ss.insertSheet(name)
  }
  return sheet
}
function getRows(sheet){
  var range = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn());
  var rows = []
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  for (var i = 2; i <= numCols; i++) {
    for (var j = 2; j <= numRows; j++) {
      var row = {}
      row["dayOfWeek"] = range.getCell(1,i).getValue()
      row["startHour"] = range.getCell(j,1).getValue()
      row["endHour"] = row["startHour"]+1;
      if (row["endHour"]==24){
      row["endHour"] = 24;
      }
      row["bidModifier"] = parseFloat(range.getCell(j,i).getValue())
      if(row["bidModifier"]<0.1){
        row["bidModifier"] = 0.1
      }
      rows.push(row)
    }
  }
  return rows;
}
function iterateCampaigns(rules) {
  var campaignIterator = AdWordsApp.campaigns()
      .withCondition('LabelNames CONTAINS_ANY [' + accountLabel + ']')
      .get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    removeAdRules(campaign)
    addAdRules(campaign, rules)
  }
}
function iterateShoppingCampaigns(rules) {
  var campaignIterator = AdWordsApp.shoppingCampaigns()
      .withCondition('LabelNames CONTAINS_ANY [' + accountLabel + ']')
      .get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    removeAdRules(campaign)
    addAdRules(campaign, rules)
  }
}
function removeAdRules(campaign){
  var schedulesSelector = campaign.targeting().adSchedules().get()
  while (schedulesSelector.hasNext()){
    var schedule = schedulesSelector.next()
    schedule.remove()
  }  
}
function addAdRules(campaign, rules){
    var nextHours = getNextHours()
    var j = 0;
    for (var i = 0; i < rules.length; i++) {
      var rule = rules[i];
      if(rule["dayOfWeek"] == nextHours["dayOfWeek"] && rule["startHour"] >= nextHours["startHour"] && j<6 && rule["bidModifier"] != 1){
        var bidModifier = rule["bidModifier"];
        var schedule = [rule["dayOfWeek"], rule["startHour"], 0, rule["endHour"], 0, bidModifier]
        campaign.addAdSchedule(rule["dayOfWeek"], rule["startHour"], 0, rule["endHour"], 0, bidModifier);
        j++;
      }
    }
}
function getNextHours(){
  var hour = 1000*60*60;
  var now = new Date()
  var then = new Date(now.getTime() + 4 * hour);
  var fromTo = {}
  var timeZone = AdWordsApp.currentAccount().getTimeZone();
  fromTo["day"] = Utilities.formatDate(now, timeZone, 'uu')
  fromTo["dayOfWeek"] = daysOfWeek[fromTo["day"]-1];
  fromTo["startHour"] = Utilities.formatDate(now, timeZone, 'HH')  
  return fromTo
      
}
function round(value, decimals) {
  return Number(Math.round(value+'e'+decimals)+'e-'+decimals);
}
