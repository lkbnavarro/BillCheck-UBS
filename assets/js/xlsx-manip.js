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

function getInnovationBillingJson(){
  // Read Excel Files for Resource List
  return getFileWorksheetJson("file-2", "myTE Innov bookings")
};

function getResourceListJson(){
  // Read Excel Files for Resource List
  return getFileWorksheetJson("file-3", "UBS Staff List")
};

function getResourceTrendJson(){
  // Read Excel Files for Resource Trend
  return getFileWorksheetJson("file-4", "Raw Data")
};

function getPWAJson(){
  // Read Excel Files for Resource List
  return getFileWorksheetJson("file-5", "PWAInput")
};

function createRow(data) {
  const row = {}
  row['STREAM'] = data.stream
  row['Enterprise ID'] = data.enterpriseId
  row['Hermes Role'] = data.hermesRole
  row['Hermes Level'] = data.hermesLevel
  row['Location Category'] = data.locationCategory
  row['Daily Rate'] = data.dailyRate
  row['Billable Hours A'] = data.billableHoursA
  row['Billable Days A'] = data.billableDaysA
  row['Gross Amount A'] = data.grossAmountA
  row['Vol. Discount A'] = data.volDiscountA
  row['Strategic Innov. Fund A'] = data.strategicInnovFundA
  row['Net Amount A'] = data.netAmountA
  row['Bill Rate'] = data.billRate
  row['Billable Hours B'] = data.billableHoursB
  row['Billable Days B'] = data.billableDaysB
  row['Gross Amount B'] = data.grossAmountB
  row['Vol. Discount B'] = data.volDiscountB
  row['Strategic Innov. Fund B'] = data.strategicInnovFundB
  row['Net Amount B'] = data.netAmountB
  row['Billable Days C'] = data.billableDaysC
  row['Net Amount C'] = data.netAmountC

  return row
}

function createRowsFromSBR(sbrFileJson, resourceListJson, resourceTrendJson) {
  return sbrFileJson.map((data) => {
    //SBR
    let fullName = data['__EMPTY_6']
    let stream = getStream(data['__EMPTY_2'])
    let enterpriseId = getEnterpriseId(fullName, resourceListJson)
    let hermesRole = data['__EMPTY_7']
    let hermesLevel = data['__EMPTY_8']
    let locationCategory = data['__EMPTY_9']
    let onshoreRate = data['__EMPTY_13']
    let offshoreRate = data['__EMPTY_14']
    let dailyRate = calculateDailyRate(onshoreRate, offshoreRate)
    let billableHoursA = data['__EMPTY_32']
    let billableDaysA = data['__EMPTY_33']
    let grossAmountA = calculateGrossAmountA(dailyRate, billableDaysA)
    let volDiscountA = calculatevolDiscountA(grossAmountA)
    let strategicInnovFundA = calculatestrategicInnovFundA(grossAmountA, volDiscountA)
    let netAmountA = calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA)
    // MME
    let billRate = getBillrate(enterpriseId, resourceTrendJson)
    let billableHoursB = getHours(enterpriseId, resourceTrendJson)
    let billableDaysB = getBillDaysMME(locationCategory, billableHoursB, fullName)
    let grossAmountB = calculateGrossAmountB(billableHoursB, billRate)
    let volDiscountB = calculatevolDiscountB(grossAmountB)
    let strategicInnovFundB = calculatestrategicInnovFundB(grossAmountB, volDiscountB)
    let netAmountB = calculateNetAmountB(grossAmountB, volDiscountB, strategicInnovFundB)
    // Variance
    let billDaysVar = calculateBillDaysVar(billableDaysB, billableDaysA)
    let NetAmountVar = calculateNetAmountVar(netAmountB, netAmountA)

    return createRow({
      stream: stream,
      enterpriseId: enterpriseId,
      hermesRole: hermesRole,
      hermesLevel: hermesLevel,
      locationCategory: locationCategory,
      dailyRate: dailyRate,
      billableHoursA: billableHoursA,
      billableDaysA: billableDaysA,
      grossAmountA: grossAmountA,
      volDiscountA: volDiscountA,
      strategicInnovFundA: strategicInnovFundA,
      netAmountA: netAmountA,
      billRate: billRate,
      billableHoursB: billableHoursB,
      billableDaysB: billableDaysB,
      grossAmountB: grossAmountB,
      volDiscountB: volDiscountB,
      strategicInnovFundB: strategicInnovFundB,
      netAmountB: netAmountB,
      billableDaysC: billDaysVar,
      netAmountC: NetAmountVar
    })
  })
}

function createRowsFromPwa(pwaJson,resourceTrendJson) {
  return pwaJson.map((data) => {

    let stream = data['Segment']
    let enterpriseId = data['Enterprise ID']
    let locationCategory = data["OMP Work Location"]
    let dailyRate = data["Daily Rate"]
    let billableHoursA = data["Actual Work Hours"]
    let billableDaysA = data["Actual Work PDs"]

    let grossAmountA = calculateGrossAmountA(dailyRate, billableDaysA)
    let volDiscountA = calculatevolDiscountA(grossAmountA)
    let strategicInnovFundA = calculatestrategicInnovFundA(grossAmountA, volDiscountA)
    let netAmountA = calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA)
    // MME
    let billRate = getBillrate(enterpriseId, resourceTrendJson)
    let billableHoursB = getHours(enterpriseId, resourceTrendJson)
    let billableDaysB = getBillDaysPWA(locationCategory, billableHoursB)
    let grossAmountB = calculateGrossAmountB(billableHoursB, billRate)
    let volDiscountB = calculatevolDiscountB(grossAmountB)
    let strategicInnovFundB = calculatestrategicInnovFundB(grossAmountB, volDiscountB)
    let netAmountB = calculateNetAmountB(grossAmountB, volDiscountB, strategicInnovFundB)
    // Variance
    let billDaysVar = calculateBillDaysVar(billableDaysB, billableDaysA)
    let NetAmountVar = calculateNetAmountVar(netAmountB, netAmountA)

    return createRow({
      stream: stream,
      enterpriseId: enterpriseId,
      hermesRole: "N/A",
      hermesLevel: "N/A",
      locationCategory: locationCategory,
      dailyRate: dailyRate,
      billableHoursA: billableHoursA,
      billableDaysA: billableDaysA,
      grossAmountA: grossAmountA,
      volDiscountA: volDiscountA,
      strategicInnovFundA: strategicInnovFundA,
      netAmountA: netAmountA,
      billRate: billRate,
      billableHoursB: billableHoursB,
      billableDaysB: billableDaysB,
      grossAmountB: grossAmountB,
      volDiscountB: volDiscountB,
      strategicInnovFundB: strategicInnovFundB,
      netAmountB: netAmountB,
      billableDaysC: billDaysVar,
      netAmountC: NetAmountVar
    })
  })
}



function createRowsFromInnovationBilling(innovationBillingJson,resourceTrendJson) {
  return innovationBillingJson.map((data) => {
    let stream = data['STREAM']
    let enterpriseId = ['Enterprise ID']
    let locationCategory = data['Location']
    let dailyRate = data['GROSS daily rate']
    let billableHoursA = data["Hrs\r\n(based on myTE)_1"]
    let hrsperday = data["hrs/day"]
    let billableDaysA = calculateBillDaysAInnovationBilling(billableHoursA, hrsperday)
    let grossAmountA = calculateGrossAmountA(dailyRate, billableDaysA)
    let volDiscountA = calculatevolDiscountA(grossAmountA)
    let strategicInnovFundA = calculatestrategicInnovFundA(grossAmountA, volDiscountA)
    let netAmountA = calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA)
    // MME
    let billRate = getBillrate(enterpriseId, resourceTrendJson)
    let billableHoursB = getHours(enterpriseId, resourceTrendJson)
    let billableDaysB = getBillDaysPWA(locationCategory, billableHoursB)
    let grossAmountB = calculateGrossAmountB(billableHoursB, billRate)
    let volDiscountB = calculatevolDiscountB(grossAmountB)
    let strategicInnovFundB = calculatestrategicInnovFundB(grossAmountB, volDiscountB)
    let netAmountB = calculateNetAmountB(grossAmountB, volDiscountB, strategicInnovFundB)
    // Variance
    let billDaysVar = calculateBillDaysVar(billableDaysB, billableDaysA)
    let NetAmountVar = calculateNetAmountVar(netAmountB, netAmountA)

    return createRow({
      stream: stream,
      enterpriseId: enterpriseId,
      hermesRole: "N/A",
      hermesLevel: "N/A",
      locationCategory: locationCategory,
      dailyRate: dailyRate,
      billableHoursA: billableHoursA,
      billableDaysA: billableDaysA,
      grossAmountA: grossAmountA,
      volDiscountA: volDiscountA,
      strategicInnovFundA: strategicInnovFundA,
      netAmountA: netAmountA,
      billRate: billRate,
      billableHoursB: billableHoursB,
      billableDaysB: billableDaysB,
      grossAmountB: grossAmountB,
      volDiscountB: volDiscountB,
      strategicInnovFundB: strategicInnovFundB,
      netAmountB: netAmountB,
      billableDaysC: billDaysVar,
      netAmountC: NetAmountVar
    })
  })
}



function getEnterpriseId(fullName, resourceListJson) {
  let resourceItem = resourceListJson.find((item) => item["Full Name"] == fullName)
  return resourceItem?.['Enterprise ID'] || "N/A"
}

function getBillrate(enterpriseId, resourceTrendJson) {
  let resourceItem = resourceTrendJson.find((item) => item["Name"] == enterpriseId && item["Category"] == "Bill Rate")
  return resourceItem?.['Quantity'] || "0"
}

function getHours(enterpriseId, resourceTrendJson) {
  let resourceItem = resourceTrendJson.find((item) => item["Name"] == enterpriseId && item["Category"] == "Hours")
  return resourceItem?.['Quantity'] || "0"
}

function getStream(stream) {
  return (stream == 'PAY & MANAGE LIQUIDITY') ? 'INNOVATION INITIATIVES' : stream
}

function calculateDailyRate(onshoreRate, offshoreRate){
  return (onshoreRate + offshoreRate)
}

function calculateGrossAmountA(dailyRate, billableDaysA) {
  return (dailyRate * billableDaysA)
}

function calculateGrossAmountB(billableHoursB, billRate) {
  return (billableHoursB * billRate)
}

function calculatevolDiscountA(grossAmountA) {
  let percentage = 5.15
  let percentageToMultiplier = (percentage / 100);
  return (grossAmountA * percentageToMultiplier)
}

function calculatevolDiscountB(grossAmountB) {
  let percentage = 5.15
  let percentageToMultiplier = (percentage / 100);
  return (grossAmountB * percentageToMultiplier)
}

function calculatestrategicInnovFundA(grossAmountA, volDiscountA) {
  let percentage = 1.93
  let percentageToMultiplier = (percentage / 100);
  return ((parseInt(grossAmountA) + parseInt(volDiscountA)) * percentageToMultiplier)
}

function calculatestrategicInnovFundB(grossAmountB, volDiscountB) {
  let percentage = 1.93
  let percentageToMultiplier = (percentage / 100);
  return ((grossAmountB + volDiscountB) * percentageToMultiplier)
}

function calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA) {
  return (grossAmountA+volDiscountA+strategicInnovFundA)
}

function calculateNetAmountB(grossAmountB, volDiscountB, strategicInnovFundB) {
  return (grossAmountB+volDiscountB+strategicInnovFundB)
}

function calculateBillDaysVar(billableDaysB, billableDaysA) {
  return Math.abs(billableDaysA-billableDaysB)
}

function calculateNetAmountVar(netAmountB, netAmountA) {
  return (netAmountA-netAmountB)
}

function getBillDaysMME(locationCategory, billableHoursB, fullName){
  let result = 0
  if (fullName == 'Largados, Mickie Rose'){
    result = billableHoursB / 8
  } else if (fullName == 'Garcia, Julius Armstrong'){
    result = billableHoursB / 9
  } else {
    if (locationCategory == 'Philippines Offshore'){
      result = billableHoursB / 9
    } else if (locationCategory == 'Switzerland Local'){
      result = billableHoursB / 8
    } else if (locationCategory == 'Philippines Landed in CH'){
      result = billableHoursB / 9
    } else if (locationCategory == 'Philippines Landed in SG'){
      result = billableHoursB / 8
    } else if (locationCategory == 'Singapore Local'){
      result = billableHoursB / 8
    } else if (locationCategory == 'India Local'){
      result = billableHoursB / 9
    } else if (locationCategory == 'Switzerland - Manno Local'){
      result = billableHoursB / 8
    } else if (locationCategory == 'Philippines Landed in HK'){
      result = billableHoursB / 7.5
    } else if (locationCategory == 'Hong Kong Local'){
      result = billableHoursB / 7.5
    } else if (locationCategory == 'India Landed in SG'){
      result = billableHoursB / 8
    } else {
      result
      }
    }
  return (result)
  }

function getBillDaysPWA(locationCategory, billableHoursB){
  let result = 0
  if (locationCategory == 'Landed - CH'){
    result = billableHoursB / 9
  } else if (locationCategory == 'Local'){
    result = billableHoursB / 8 
  } else if (locationCategory == 'Offshore'){
    result = billableHoursB / 9
  } else {
    result
  }
  return (result)
}

function calculateBillDaysAInnovationBilling(billableHoursA, hrsperday){
  return (billableHoursA / hrsperday)
}

function createExcel() {
    let serviceBillingJson = getServiceBillingJson();
    let innovationBillingJson = getInnovationBillingJson();
    let resourceListJson = getResourceListJson();
    let resourceTrendJson = getResourceTrendJson();
    let pwaJson = getPWAJson();

    // console.log(serviceBillingJson)
    // console.log(innovationBillingJson)
    // console.log(resourceListJson)
    // console.log(resourceTrendJson)
    // console.log(pwaJson)
    
    let sbrRows = createRowsFromSBR(serviceBillingJson, resourceListJson, resourceTrendJson)
    let pwaRows = createRowsFromPwa(pwaJson, resourceTrendJson)
    let innovationRows = createRowsFromInnovationBilling(innovationBillingJson, resourceTrendJson)
    let mergedRows = [...sbrRows, ...pwaRows,...innovationRows]


    var workbook = xlsx.utils.book_new();
    var worksheet = xlsx.utils.json_to_sheet(mergedRows);
    worksheet["!cols"] = [ { wch: 40 }, {wch: 20}, {wch:23}, {wch:12}, {wch:22}, {wch:19}, {wch:13}, {wch:18}, {wch:17}, {wch:25}, {wch:25}, 
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