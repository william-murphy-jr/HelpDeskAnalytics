const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw Data'); 
const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calculated Data');
const lastRow = sourceSheet.getLastRow();
const lastColumn = sourceSheet.getLastColumn();
const searchRange = sourceSheet.getRange(2,1, lastRow, lastColumn);
const rangeValues = searchRange.getValues(); // [2 - lastRow, 1 to lastColumn]

const headerTitles = ['Facility', 'Number of Tickets', 'Total Hours', 'Average Time to Close','',
                      'Topic Category', 'Number of Tickets', 'Total Hours', 'Average Time to Close', '',
                      'ASR', 'Number of Tickets', 'Total Hours', 'Average Time to Close'];

const __DEBUG = true; // Turn Logger On

function onOpen(e) {
  targetSheet.getRange(2, 1, 50, 50).clearContent();
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Custom');
  menu.addItem('Facility Tally', 'facilityTally');
  menu.addItem('Category Tally', 'categoryTally');
  menu.addItem('ASR Tally', 'asrTally');
  menu.addToUi();
  addHeaders(headerTitles, 5);
}

/**
* Add Headers and add a blank column as a seperator seperate column
*/
function addHeaders(headers){
//  headers = headerTitles; // Uncomment for self-contained testing ONLY
  for (let i = 0; i < headers.length; i++) {
    const col = i + 1; 
      targetSheet.getRange(1, col).setValue(headers[i]);
  }
  __DEBUG && Logger.log(headers);
  __DEBUG && Logger.log('\n\n', rangeValues);
}

function calculateMetrics(col, key) {
  const metricData = crunchNums(key);
  targetSheet.getRange(2,col,metricData.length, 4).clearContent().setValues(metricData);
  
  __DEBUG && Logger.log(`\n${key}Data: \n`, metricData);
  __DEBUG && Logger.log(`\n${key}Data.length: `, metricData.length);
} // calculateMetrics

function facilityTally() {
  calculateMetrics(1, 'facility');
} // faciltityTally

function categoryTally() {
 calculateMetrics(6, 'category'); 
} // categoryTally

function asrTally() {
  calculateMetrics(11, 'asr');
} // asrTally

function getData() {
  let rowsData = [];
  for (let i = 0; i < rangeValues.length; i++){
    const tempFacility = rangeValues[i][4];
    const tempASR = rangeValues[i][5];
    const tempCategory = rangeValues[i][12]; 
    const tempStart = rangeValues[i][15];
    const tempClose = rangeValues[i][17]; 
    let tempTimeToClose = 'ticket is not closed';
    if (tempClose) {
      const date1 = new Date(tempStart);
      const date2 = new Date(tempClose);
      const timeDiff = date2.getTime() - date1.getTime();
      tempTimeToClose = timeDiff / (1000 * 3600);
    }
    const row = {
      facility: tempFacility,
      asr: tempASR,
      category: tempCategory,
      timeToClose: tempTimeToClose
    };
    rowsData.push(row);
  }
//  __DEBUG && Logger.log('rowsData: >>>> ', rowsData);
  return rowsData;
}

function crunchNums(key) {
  let rowsData = getData();
  rowsData.sort(compare(key));
  __DEBUG && Logger.log('\n\nrowsData: \n', rowsData);
  
  let result = [];
  let comparisonEl = rowsData[0];
  let count = 0;
  let time = 0;
  let averageTime = 0;
  __DEBUG && Logger.log('\n\nrtypeof comparisonEl.time: ', typeof comparisonEl.time);
  
  for (let i = 0; i < rowsData.length; i++){
    let currentEl = rowsData[i];
    if (currentEl[key] != comparisonEl[key]) {
      if (count > 0){
        averageTime = time/count;
        result.push([comparisonEl[key], count, time.toFixed(2), averageTime.toFixed(2)]);
      }
      comparisonEl = currentEl;
      count = 0;
      time = 0;
    } else {
      count++;
      const timeToClose = parseFloat(currentEl.timeToClose);
      const currTimeToClose = Number.isNaN(timeToClose) ? 0 : timeToClose
      time += parseFloat(currTimeToClose);
    }
  }
  averageTime = time/count;
  result.push([comparisonEl[key], count, time.toFixed(2), averageTime.toFixed(2)]);
  __DEBUG && Logger.log('\n\nresult: \n', result);
  return result;
}


function compare(key) {
  return function(a, b) {
    return a[key] < b[key] ? -1 : 1;
  }
}
