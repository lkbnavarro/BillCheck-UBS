'use strict';

const xlsx = require("xlsx");

function readSBR(){
    // Read Excel Files for SBR
    const workbookSBR = xlsx.readFile("C:/Users/laurence.k.navarro/Desktop/BillCheck-UBS/assets/test_input/Service Billing Report under SO2021 - November 2022_v1.2.xlsx", {cellDates:true});
    const worksheetSBR = workbookSBR.Sheets["Resource Reporting - MC"];
    let range = xlsx.utils.decode_range(worksheetSBR['!ref']);
    pushSBRtoArray(range, worksheetSBR);
    pushRLtoArray(range);
};

function readRL(){
    // Read Excel Files for Resource List
    const workbookRL = xlsx.readFile("C:/Users/laurence.k.navarro/Desktop/BillCheck-UBS/assets/test_input/UBS Resource List_1207.xlsx", {cellDates:true});
    const worksheetRL = workbookRL.Sheets["UBS Staff List"];
    pushRLtoArray(range, worksheetRL);
};



function createArray(){
  const outputHeader = ["STREAM", "Enterprise ID", "Hermes Role", "Hermes Level", "Location Category", "Daily Rate (Onshore)", "Daily Rate (Onshore)",
                    "Billable Hours (SBR)", "Billable Days (SBR)", "Gross Amount", "Net Amount", "Bill Rate (MME)", "Billable Days (MME)", "Gross Amount", "Net Amount",
                  "Billable Days (Var)", "Net Amount (Var)"];
  let streamValues = [];
  let resourceName = [];                             // from SBR
  let resourceFullName = [];                         // from Resource List
  let enterpriseID = [];                             // from Resource List
  let hermesRole = [];                               // from SBR
  let hermesLevel = [];                              // from SBR
  let locationCat = [];                              // from SBR
  let dailyrateOn = [];                              // from SBR
  let dailyrateOff = [];                              // from SBR
  let sbrBillHrs = [];                              // from SBR
  let sbrBillDays = [];                              // from SBR

};

function pushSBRtoArray(range, worksheetSBR, streamValues, resourceName, hermesRole, hermesLevel, locationCat, dailyrateOn, dailyrateOff, sbrBillHrs, sbrBillDays){
  for (let rowNum = 13; rowNum <= range.e.r; rowNum++) {
    // Get the cell value
    let streamCell = worksheetSBR["C" + rowNum];
    if (streamCell) {
      streamValues.push(streamCell.v);
    }
    let resourceNameCell = worksheetSBR["G"+rowNum];
    if (resourceNameCell){
      resourceName.push(resourceNameCell.v);
    }
    let hermesRoleCell = worksheetSBR["H" + rowNum];
    if (hermesRoleCell) {
      hermesRole.push(hermesRoleCell.v);
    }
    let hermesLevelCell = worksheetSBR["I" + rowNum];
    if (hermesLevelCell) {
      hermesLevel.push(hermesLevelCell.v);
    }
    let locationCatCell = worksheetSBR["J" + rowNum];
    if (locationCatCell) {
      locationCat.push(locationCatCell.v);
    }
  
    let dailyrateOnCell = worksheetSBR["N" + rowNum];
    if (dailyrateOnCell) {
      dailyrateOn.push(dailyrateOnCell.v);
    }
    let dailyrateOffCell = worksheetSBR["O" + rowNum];
    if (dailyrateOffCell) {
      dailyrateOff.push(dailyrateOffCell.v);
    }
    let sbrBillHrsCell = worksheetSBR["AG" + rowNum];
    if (sbrBillHrsCell) {
      sbrBillHrs.push(sbrBillHrsCell.v);
    }
    let sbrBillDaysCell = worksheetSBR["AH" + rowNum];
    if (sbrBillDaysCell) {
      sbrBillDays.push(sbrBillDaysCell.v);
    }
  }
};

function pushRLtoArray(range, worksheetRL){
  for (let rowNum = 2; rowNum <= range.e.r; rowNum++) {
    let resourceFullNameCell = worksheetRL["B" + rowNum];
    if (resourceFullNameCell) {
    resourceFullName.push(resourceFullNameCell.v);
    }
    let enterpriseIDCell = worksheetRL["C"+rowNum];
    if (enterpriseIDCell){
    enterpriseID.push(enterpriseIDCell.v);
    }
  }
};

function replaceStream(){
  const oldStream = 'PAY & MANAGE LIQUIDITY';
  const newStream = 'INNOVATION INITIATIVES';
  for (let i = 0; i < streamValues.length; i++){
    if(streamValues[i] === oldStream){
      streamValues.splice(i, 1, newStream);
    }
  }
};

function calculateGross(){
    let grossAmountOn = [];
    let grossAmountOff = [];
    let grossAmount = [];

    for (let i = 0; i < array1.length; i++){
        sumArray[i] = array1[i] + array2[i];
    }  
}


// For Gross Amount

/* let grossAmount = [];

for (let i = 0; i < sbrBillDays.length; i++){
  grossAmount.push(sbrBillDays[i] * 
}
*/

/* COMPARING ARRAYS ANG GETTING THEIR ENTERPRISE ID

const newResourceName = resourceName.map(item => item.trim());
const newResourceFullName = resourceFullName.map(item => item.trim());

let count = 0;

for (let i = 0; i < newResourceName.length; i++) {
  let flag = true;
  for (let j = 0; j < newResourceFullName.length; j++) {
      if (newResourceName[i] === newResourceFullName[j]) {
          flag = false;
          break;
      }
  }
  if (flag) {
      console.log(newResourceName[i]);
  }
}


for (let i = 0; i < newResourceFullName.length; i++) {
  let flag = true;
  for (let j = 0; j < newResourceName.length; j++) {
      if (newResourceFullName[i] === newResourceName[j]) {
          flag = false;
          break;
      }
  }
  if (flag) {
      console.log(newResourceFullName[i]);
  }
}

*/

function pushValues(){
    const newWorkbook = xlsx.utils.book_new();
    let outputData = [];
    outputData.push(outputHeader);

    for(let i=0; i<streamValues.length; i++) {
      outputData.push([streamValues[i], enterpriseID[i], hermesRole[i], hermesLevel[i], locationCat[i], dailyrateOn[i], dailyrateOff[i], sbrBillHrs[i], 
        sbrBillDays[i]])
    };
};


function writeFile(){
    const newWorksheet = xlsx.utils.aoa_to_sheet(outputData);
    newWorksheet["!cols"] = [ { wch: 40 }, {wch: 30}, {wch:23}, {wch:12}, {wch:22}, {wch:19}, {wch:19}, {wch:18}, {wch:17}, {wch:30}, {wch:30}, 
      {wch:30}, {wch:30}, {wch:30}, {wch:30}, {wch:30}, {wch:30}];
    
    // Add the new worksheet to the new workbook
    xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet 1');
    
    // Write the new workbook to a file
    xlsx.writeFile(newWorkbook, 'new-test.xlsx');
};

function testOutput(){
    console.log(sbrBillDays);
    console.log(sbrBillHrs);
    console.log(dailyrateOff); 
  };


function testExecute(){
    readSBR();
    readRL();
    createArray();
    pushSBRtoArray();
    pushRLtoArray();
    replaceStream();
    testOutput();
};

testExecute();