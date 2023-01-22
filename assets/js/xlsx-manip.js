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
    let billableDays = data['__EMPTY_33']
    let enterpriseId = getEnterpriseId(fullName, resourceListJson)
    let grossAmountA = calculateGrossAmountA(onshoreRate, offshoreRate, billableDays)
    let volDiscountA = calculatevolDiscountA(grossAmountA)
    let strategicInnovFundA = calculatestrategicInnovFundA(grossAmountA, volDiscountA)
    let billRate = getBillrate(enterpriseId, resourceTrendJson)
    let hours = getHours(enterpriseId, resourceTrendJson)
    let stream = getStream(data['__EMPTY_2'])
    let netAmountA = calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA)
    let grossAmountB = calculateGrossAmountB(hours, billRate)
    let volDiscountB = calculatevolDiscountB(grossAmountB)
    let strategicInnovFundB = calculatestrategicInnovFundB(grossAmountB, volDiscountB)
    let netAmountB = calculateNetAmountB(grossAmountB, volDiscountB, strategicInnovFundB)
    let location = data['__EMPTY_9']
    let billDaysMME = getBillDaysMME(location, hours, fullName)
    let billDaysVar = calculateBillDaysVar(billDaysMME, billableDays)
    let NetAmountVar = calculateNetAmountVar(netAmountB, netAmountA)
   

    row['STREAM'] = stream
    row['Enterprise ID'] = enterpriseId
    row['Hermes Role'] = data['__EMPTY_7']
    row['Hermes Level'] = data['__EMPTY_8']
    row['Location Category'] = location
    row['Daily Rate (Onshore)'] = onshoreRate
    row['Daily Rate (Offshore)'] = offshoreRate
    row['Billable Hours (SBR)'] = data['__EMPTY_32']
    row['Billable Days (SBR)'] = billableDays
    row['Gross Amount (SBR)'] = grossAmountA
    row['Vol. Discount'] = volDiscountA
    row['Strategic Innov. Fund (SBR)'] = strategicInnovFundA
    row['Net Amount (SBR)'] = netAmountA
    row['Bill Rate'] = billRate
    row['Billable Hours (MME)'] = hours
    row['Billable Days (MME)'] = billDaysMME
    row['Gross Amount (MME)'] = grossAmountB
    row['Vol. Discount (MME)'] = volDiscountB
    row['Strategic Innov. Fund (MME)'] = grossAmountB
    row['Net Amount (MME)'] = netAmountB
    row['Billable Days (Var)'] = billDaysVar
    row['Net Amount (Var)'] = NetAmountVar
    
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

function calculateGrossAmountA(onshoreRate, offshoreRate, billableDays) {
  return (onshoreRate * billableDays) + (offshoreRate * billableDays)
}

function calculateGrossAmountB(hours, billRate) {
  return hours * billRate
}

function calculatevolDiscountA(grossAmountA) {
  let percentage = 5.15
  let percentageToMultiplier = (percentage / 100);
  return grossAmountA * percentageToMultiplier
}

function calculatevolDiscountB(grossAmountB) {
  let percentage = 5.15
  let percentageToMultiplier = (percentage / 100);
  return grossAmountB * percentageToMultiplier
}

function calculatestrategicInnovFundA(grossAmountA, volDiscountA) {
  let percentage = 1.93
  let percentageToMultiplier = (percentage / 100);
  return (grossAmountA+ volDiscountA) * percentageToMultiplier
}

function calculatestrategicInnovFundB(grossAmountB, volDiscountB) {
  let percentage = 1.93
  let percentageToMultiplier = (percentage / 100);
  return (grossAmountB+ volDiscountB) * percentageToMultiplier
}

function calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA) {
  return grossAmountA+volDiscountA+strategicInnovFundA
}

function calculateNetAmountB(grossAmountB, volDiscountB, strategicInnovFundB) {
  return grossAmountB+volDiscountB+strategicInnovFundB
}

function calculateBillDaysVar(billDaysMME, billableDays) {
  return billableDays-billDaysMME
}

function calculateNetAmountVar(netAmountB, netAmountA) {
  return netAmountB-netAmountA
}


function getBillDaysMME(location, hours, fullName){
  let result = 0
  if (fullName == 'Largados, Mickie Rose'){
    result = hours / 8
  } else if (fullName == 'Garcia, Julius Armstrong'){
    result = hours / 9
  } else {
    if (location == 'Philippines Offshore'){
      result = hours / 9
    } else if (location == 'Switzerland Local'){
      result = hours / 8
    } else if (location == 'Philippines Landed in CH'){
      result = hours / 9
    } else if (location == 'Philippines Landed in SG'){
      result = hours / 8
    } else if (location == 'Singapore Local'){
      result = hours / 8
    } else if (location == 'India Local'){
      result = hours / 9
    } else if (location == 'Switzerland - Manno Local'){
      result = hours / 8
    } else if (location == 'Philippines Landed in HK'){
      result = hours / 7.5
    } else if (location == 'Hong Kong Local'){
      result = hours / 7.5
    } else if (location == 'India Landed in SG'){
      result = hours / 8
    } else {
      result
      }
    }
  return result
  }

function createExcel() {
    let serviceBillingJson = getServiceBillingJson();
    let resourceListJson = getResourceListJson();
    let resourceTrendJson = getResourceTrendJson()
    let outputJson = createRows(serviceBillingJson, resourceListJson, resourceTrendJson)
    
    var workbook = xlsx.utils.book_new();
    var worksheet = xlsx.utils.json_to_sheet(outputJson);
    worksheet["!cols"] = [ { wch: 40 }, {wch: 30}, {wch:23}, {wch:12}, {wch:22}, {wch:19}, {wch:19}, {wch:18}, {wch:17}, {wch:25}, {wch:25}, 
      {wch:25}, {wch:25}, {wch:25}, {wch:25}, {wch:25}, {wch:25}, {wch:25}, {wch:25}, {wch:25}, {wch:25}, {wch:25}];
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