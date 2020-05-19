const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1'); 
const lr = sourceSheet.getLastRow();
const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');

function onOpen(e) {
  targetSheet.getRange(2, 1, 50, 50).clearContent();
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Custom');
  menu.addItem('Facility Tally', 'facilityTally');
  menu.addItem('Category Tally', 'categoryTally');
  menu.addToUi();
}

function facilityTally() {
    var facilities = [];
  
  // clear tally column in target sheet
  targetSheet.getRange(2, 1, 15, 2).clearContent();
  
  // push content from all cells in column M into facilities array
  for (let i = 2; i <lr+1; i++) {
    var data = sourceSheet.getRange(i, 5).getValue();
    facilities.push(data);
  }
  
  // sort facilities array
  facilities.sort();
  
  // initialize variables for counting unique elements in our array
  let comparisonVal = facilities[0];
  let cnt = 0;
  let row = 2;
  
  // loop through the array
  for (let i = 0; i <= facilities.length; i++){
    // if the array element is not the same as the comparison value
    // place it in a cell in column 4 with corresponding count in column 5
    if (facilities[i] != comparisonVal) {
       targetSheet.getRange(row,1).setValue(comparisonVal);
       targetSheet.getRange(row,2).setValue(cnt);
       row++; // move to the next row
      comparisonVal = facilities[i]; // set the comparison value to the current array element
      cnt = 1; // reset the count to one
    } else { 
      cnt++; //if the array element IS the same as the comparison value, simply increment the count
    }
  }  
}


function categoryTally() {
  let topicsAndTimes = [];
  targetSheet.getRange(2, 4, lr, 4).clearContent();
  // loop through source sheet getting category, start time, and close time for each ticket
  for (let i = 2; i < lr+1; i++){
    let category = sourceSheet.getRange(i, 13).getValue(); // get category
    let facility = sourceSheet.getRange(i, 5).getValue(); // get facility
    let asr = sourceSheet.getRange(i, 6).getValue(); // get ASR 
    let startTime = sourceSheet.getRange(i, 16).getValue();
    let closeTime = sourceSheet.getRange(i, 18).getValue(); 
    // if the ticket was closed, calculate the number of hours from open to close
    if (closeTime) {
      let date1 = new Date(startTime);
      let date2 = new Date(closeTime);
      let timeDiff = date2.getTime() - date1.getTime();
      let hoursDiff = timeDiff / (1000 * 3600);
      // push ticket category and time to close into 2 dimensional array
      topicsAndTimes.push([category, hoursDiff]);
    }
  }
  // sort the array by category
  topicsAndTimes.sort();
  
  // initialize variables for counting unique elements in our array
  let comparisonVal = topicsAndTimes[0][0];
  let hrs = 0;
  let cntCat = 0;
  let cntAsr = 0;
  let cntFac = 0; 
  let row = 2;
  let numTopics = 0;
  
  // loop through the array
  for (let i = 0; i < topicsAndTimes.length; i++){
    // if the first element of each sub-array is not the same as the comparison value
    // place it in a cell in column 4 with corresponding count in column 5 and time totals in column 6
    if (topicsAndTimes[i][0] != comparisonVal) {
       targetSheet.getRange(row,4).setValue(comparisonVal);
       targetSheet.getRange(row,5).setValue(cntCat);
       targetSheet.getRange(row,6).setValue(hrs.toFixed(2));
       row++; // move to the next row
       comparisonVal = topicsAndTimes[i][0]; // set the comparison value to the first element current sub array
       hrs = topicsAndTimes[i][1]; // set hours to second element of current sub array
       cntCat = 1; // reset the count to one
       numTopics++; //tracking the total number of topics
     }  else { 
       hrs+= topicsAndTimes[i][1]; //if the array element IS the same as the comparison value, add to the total hours
       cntCat++; //and increment the count
    }
  }
  
  let totalTix = 0;
  let totalTime = 0;
  // for each row, use the number of tickets and total hours to calculate the average time to close
  for (let i = 2; i < numTopics+2; i++){
    let numTix = targetSheet.getRange(i, 5).getValue();
    let time = targetSheet.getRange(i, 6).getValue();
    let averageTime = (time / numTix)/24.0;
    targetSheet.getRange(i, 7).setValue(averageTime);
    // tracking the total number of tickets and total time for all tickets
    totalTix += numTix;
    totalTime += time;
  }
  
  let totalAverageTime = (totalTime / totalTix) / 24.0;
  // write information about totals to last row
  targetSheet.getRange(numTopics+2, 4).setValue('Total');
  targetSheet.getRange(numTopics+2, 5).setValue(totalTix);
  targetSheet.getRange(numTopics+2, 6).setValue(totalTime);
  targetSheet.getRange(numTopics+2, 7).setValue(totalAverageTime);  
}

