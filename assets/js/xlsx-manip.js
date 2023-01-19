const xlsx = require("xlsx");
const path = require('path');

let files = []

function openFileDialog(fileNumber) {
  if (window.File && window.FileReader && window.FileList && window.Blob) {
    var fileInput = document.createElement("input");
    fileInput.type = "file";
    fileInput.addEventListener("change", function () {
      var file = this.files[0];
      var filePath = file.path;
      var fileName = file.name;

      pushFile(fileNumber, filePath)
      document.getElementById(fileNumber + "-selected").innerHTML = "File Name: " + fileName;
    });
    fileInput.click();
  } else {
    alert("The File APIs are not fully supported in this browser.");
  }
}

function pushFile(fileNumber, filePath) {
  let fileFound = files.find((file) => file?.id == fileNumber)
  if (fileFound == undefined) {
      files.push({
          id: fileNumber,
          path: filePath
      })
  } else {
      let index = files.findIndex((file) => file.fileNumber == fileFound.fileNumber)
      files[index] = {
          id: fileNumber,
          path: filePath
      }
  }
}
// const outputHeader = ["STREAM", "Enterprise ID", "Hermes Role", "Hermes Level", "Location Category", "Daily Rate (Onshore)", "Daily Rate (Onshore)",
// "Billable Hours (SBR)", "Billable Days (SBR)", "Gross Amount", "Net Amount", "Bill Rate (MME)", "Billable Days (MME)", "Gross Amount", "Net Amount",
// "Billable Days (Var)", "Net Amount (Var)"];

// let streamValues = [];
// let resourceName = [];                             // from SBR
// let resourceFullName = [];                         // from Resource List
// let enterpriseID = [];                             // from Resource List
// let hermesRole = [];                               // from SBR
// let hermesLevel = [];                              // from SBR
// let locationCat = [];                              // from SBR
// let dailyrateOn = [];                              // from SBR
// let dailyrateOff = [];                              // from SBR
// let sbrBillHrs = [];                              // from SBR
// let sbrBillDays = [];                              // from SBR

function getFileWorksheetJson(fileId, sheetName) {
  let excelFile = files.find((file) => file?.id == fileId)
  const workbook = xlsx.readFile(excelFile.path, { cellDates:true });
  const worksheet = workbook.Sheets[sheetName];
 
  return xlsx.utils.sheet_to_json(worksheet)
}

function getServiceBillingJson(){
  // Read Excel Files for SBR
  // Get data that has number on column 1
  return getFileWorksheetJson("file-1", "Resource Reporting - MC")
  .filter((item) => typeof item['__EMPTY_1'] === 'number')
};

function getResourceListJson(){
  // Read Excel Files for Resource List
  return getFileWorksheetJson("file-3", "UBS Staff List")
};

function getResourceTrendJson(){
  // Read Excel Files for Resource Trend
  return getFileWorksheetJson("file-4", "Raw Data")
};

function createRows(sbrFileJson, resourceListJson, resourceTrendJson) {
  return sbrFileJson.map((data) => {
    const row = {}
    let fullName = data['__EMPTY_6']
    let onshoreRate = data['__EMPTY_13']
    let offshoreRate = data['__EMPTY_14']
    let billableDays = data['__EMPTY_25']
    let enterpriseId = getEnterpriseId(fullName, resourceListJson)
    let grossAmount = calculateGrossAmount(onshoreRate, offshoreRate, billableDays)
    let volDiscount = calculateVolDiscount(grossAmount)
    let strategicInnovFund = calculateStrategicInnovFund(grossAmount, volDiscount)
    let billRate = getBillrate(enterpriseId, resourceTrendJson)
    let hours = getHours(enterpriseId, resourceTrendJson)
    let stream = getStream(data['__EMPTY_2'])
   

    row['STREAM'] = stream
    row['Enterprise ID'] = enterpriseId
    row['Hermes Role'] = data['__EMPTY_7']
    row['Hermes Level'] = data['__EMPTY_8']
    row['Location Category'] = data['__EMPTY_9']
    row['Daily Rate (Onshore)'] = onshoreRate
    row['Daily Rate (Offshore)'] = offshoreRate
    row['Billable Hours A'] = data['__EMPTY_24']
    row['Billable Days A'] = data['__EMPTY_25']
    row['Gross Amount A'] = grossAmount
    row['Vol. Discount'] = volDiscount
    row['Strategic Innov. Fund'] = strategicInnovFund
    row['Net Amount A'] = "N/A"
    row['Bill Rate'] = billRate
    row['Billable Hours B'] = hours
    row['Billable Days B'] = "N/A"
    row['Gross Amount B'] = "N/A"
    row['Net Amount B'] = "N/A"
    row['Billable Days C'] = "N/A"
    row['Net Amount C'] = "N/A"
    
    return row
  })
}

function getEnterpriseId(fullName, resourceListJson) {
  let resourceItem = resourceListJson.find((item) => item["Full Name"] == fullName)
  return resourceItem?.['Enterprise ID'] || "N/A"
}

function getBillrate(enterpriseId, resourceTrendJson) {
  let resourceItem = resourceTrendJson.find((item) => item["Name"] == enterpriseId && item["Category"] == "Bill Rate")
  return resourceItem?.['Quantity'] || "N/A"
}

function getHours(enterpriseId, resourceTrendJson) {
  let resourceItem = resourceTrendJson.find((item) => item["Name"] == enterpriseId && item["Category"] == "Hours")
  return resourceItem?.['Quantity'] || "N/A"
}

function getStream(stream) {
  return (stream == 'PAY & MANAGE LIQUIDITY') ? 'INNOVATION INITIATIVES' : stream
}

function calculateGrossAmount(onshoreRate, offshoreRate, billableDays) {
  return (onshoreRate * billableDays) + (offshoreRate * billableDays)
}

function calculateVolDiscount(grossAmount) {
  let percentage = 5.15
  let percentageToMultiplier = (percentage / 100);
  return grossAmount * percentageToMultiplier
}

function calculateStrategicInnovFund(grossAmount, volDiscount) {
  let percentage = 1.93
  let percentageToMultiplier = (percentage / 100);
  return (grossAmount+ volDiscount) * percentageToMultiplier
}

function createExcel() {
    let serviceBillingJson = getServiceBillingJson();
    let resourceListJson = getResourceListJson();
    let resourceTrendJson = getResourceTrendJson()
    let outputJson = createRows(serviceBillingJson, resourceListJson, resourceTrendJson)
    
    var workbook = xlsx.utils.book_new();
    var worksheet = xlsx.utils.json_to_sheet(outputJson);
    xlsx.utils.book_append_sheet(workbook, worksheet, "testSheet");
    xlsx.writeFile(workbook, "textBook.xlsx");
}



// function pushSBRtoArray(range, worksheetSBR, streamValues, resourceName, hermesRole, hermesLevel, locationCat, dailyrateOn, dailyrateOff, sbrBillHrs, sbrBillDays){
//   for (let rowNum = 13; rowNum <= range.e.r; rowNum++) {
//     // Get the cell value
//     let streamCell = worksheetSBR["C" + rowNum];
//     if (streamCell) {
//       streamValues.push(streamCell.v);
//     }
//     let resourceNameCell = worksheetSBR["G"+rowNum];
//     if (resourceNameCell){
//       resourceName.push(resourceNameCell.v);
//     }
//     let hermesRoleCell = worksheetSBR["H" + rowNum];
//     if (hermesRoleCell) {
//       hermesRole.push(hermesRoleCell.v);
//     }
//     let hermesLevelCell = worksheetSBR["I" + rowNum];
//     if (hermesLevelCell) {
//       hermesLevel.push(hermesLevelCell.v);
//     }
//     let locationCatCell = worksheetSBR["J" + rowNum];
//     if (locationCatCell) {
//       locationCat.push(locationCatCell.v);
//     }
  
//     let dailyrateOnCell = worksheetSBR["N" + rowNum];
//     if (dailyrateOnCell) {
//       dailyrateOn.push(dailyrateOnCell.v);
//     }
//     let dailyrateOffCell = worksheetSBR["O" + rowNum];
//     if (dailyrateOffCell) {
//       dailyrateOff.push(dailyrateOffCell.v);
//     }
//     let sbrBillHrsCell = worksheetSBR["AG" + rowNum];
//     if (sbrBillHrsCell) {
//       sbrBillHrs.push(sbrBillHrsCell.v);
//     }
//     let sbrBillDaysCell = worksheetSBR["AH" + rowNum];
//     if (sbrBillDaysCell) {
//       sbrBillDays.push(sbrBillDaysCell.v);
//     }
//   }
// };

// function pushRLtoArray(range, worksheetRL){
//   for (let rowNum = 2; rowNum <= range.e.r; rowNum++) {
//     let resourceFullNameCell = worksheetRL["B" + rowNum];
//     if (resourceFullNameCell) {
//     resourceFullName.push(resourceFullNameCell.v);
//     }
//     let enterpriseIDCell = worksheetRL["C"+rowNum];
//     if (enterpriseIDCell){
//     enterpriseID.push(enterpriseIDCell.v);
//     }
//   }
// };

// function replaceStream(){
//   const oldStream = 'PAY & MANAGE LIQUIDITY';
//   const newStream = 'INNOVATION INITIATIVES';
//   for (let i = 0; i < streamValues.length; i++){
//     if(streamValues[i] === oldStream){
//       streamValues.splice(i, 1, newStream);
//     }
//   }
// };

// function calculateGross(){
//     let grossAmountOn = [];
//     let grossAmountOff = [];
//     let grossAmount = [];

//     for (let i = 0; i < array1.length; i++){
//         sumArray[i] = array1[i] + array2[i];
//     }  
// }


// // For Gross Amount

// /* let grossAmount = [];

// for (let i = 0; i < sbrBillDays.length; i++){
//   grossAmount.push(sbrBillDays[i] * 
// }
// */

// /* COMPARING ARRAYS ANG GETTING THEIR ENTERPRISE ID

// const newResourceName = resourceName.map(item => item.trim());
// const newResourceFullName = resourceFullName.map(item => item.trim());

// let count = 0;

// for (let i = 0; i < newResourceName.length; i++) {
//   let flag = true;
//   for (let j = 0; j < newResourceFullName.length; j++) {
//       if (newResourceName[i] === newResourceFullName[j]) {
//           flag = false;
//           break;
//       }
//   }
//   if (flag) {
//       console.log(newResourceName[i]);
//   }
// }


// for (let i = 0; i < newResourceFullName.length; i++) {
//   let flag = true;
//   for (let j = 0; j < newResourceName.length; j++) {
//       if (newResourceFullName[i] === newResourceName[j]) {
//           flag = false;
//           break;
//       }
//   }
//   if (flag) {
//       console.log(newResourceFullName[i]);
//   }
// }

// */

// function pushValues(){
//     const newWorkbook = xlsx.utils.book_new();
//     let outputData = [];
//     outputData.push(outputHeader);

//     for(let i=0; i<streamValues.length; i++) {
//       outputData.push([streamValues[i], enterpriseID[i], hermesRole[i], hermesLevel[i], locationCat[i], dailyrateOn[i], dailyrateOff[i], sbrBillHrs[i], 
//         sbrBillDays[i]])
//     };
// };


// function writeFile(){
//     const newWorksheet = xlsx.utils.aoa_to_sheet(outputData);
//     newWorksheet["!cols"] = [ { wch: 40 }, {wch: 30}, {wch:23}, {wch:12}, {wch:22}, {wch:19}, {wch:19}, {wch:18}, {wch:17}, {wch:30}, {wch:30}, 
//       {wch:30}, {wch:30}, {wch:30}, {wch:30}, {wch:30}, {wch:30}];
    
//     // Add the new worksheet to the new workbook
//     xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet 1');
    
//     // Write the new workbook to a file
//     xlsx.writeFile(newWorkbook, 'new-test.xlsx');
// };

// function testOutput(){
//     console.log(sbrBillDays);
//     console.log(sbrBillHrs);
//     console.log(dailyrateOff); 
//   };


// function testExecute(){
//     readSBR();
//     readRL();
//     createArray();
//     pushSBRtoArray();
//     pushRLtoArray();
//     replaceStream();
//     testOutput();
// };

// testExecute();