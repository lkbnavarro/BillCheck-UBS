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
  const workbook = xlsx.readFile(excelFile.path, { cellDates: true });
  const worksheet = workbook.Sheets[sheetName];

  return xlsx.utils.sheet_to_json(worksheet)
}

function getServiceBillingJson() {
  // Read Excel Files for SBR
  // Get data that has number on column 1
  return getFileWorksheetJson("file-1", "Resource Reporting - MC")
    .filter((item) => typeof item['__EMPTY_1'] === 'number')
};

function getInnovationBillingJson() {
  // Read Excel Files for Resource List
  const innovationJson = getFileWorksheetJson("file-2", "myTE Innov bookings")
  const totalRowIndex = innovationJson.findIndex((row) => row['Contract Type'].toUpperCase() == 'TOTAL')
  return innovationJson.slice(0, totalRowIndex)
};

function getResourceListJson() {
  // Read Excel Files for Resource List
  return getFileWorksheetJson("file-3", "UBS Staff List")
};

function getResourceTrendJson() {
  // Read Excel Files for Resource Trend
  return getFileWorksheetJson("file-4", "Raw Data")
};

function getPWAJson() {
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
    let strategicInnovFundA = calculateStrategicInnovFundA(grossAmountA, volDiscountA)
    let netAmountA = calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA)
    // MME
    let billRate = getBillrate(enterpriseId, resourceTrendJson)
    let billableHoursB = getHours(enterpriseId, resourceTrendJson)
    let billableDaysB = getBillDaysMME(locationCategory, billableHoursB, fullName)
    let grossAmountB = calculateGrossAmountB(billableHoursB, billRate)
    let volDiscountB = calculatevolDiscountB(grossAmountB)
    let strategicInnovFundB = calculateStrategicInnovFundB(grossAmountB, volDiscountB)
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

function createRowsFromPwa(pwaJson, resourceTrendJson) {
  return pwaJson.map((data) => {

    let stream = data['Segment']
    let enterpriseId = data['Enterprise ID']
    let locationCategory = data["OMP Work Location"]
    let dailyRate = data["Daily Rate"]
    let billableHoursA = data["Actual Work Hours"]
    let billableDaysA = data["Actual Work PDs"]

    let grossAmountA = calculateGrossAmountA(dailyRate, billableDaysA)
    let volDiscountA = calculatevolDiscountA(grossAmountA)
    let strategicInnovFundA = calculateStrategicInnovFundA(grossAmountA, volDiscountA)
    let netAmountA = calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA)
    // MME
    let billRate = getBillrate(enterpriseId, resourceTrendJson)
    let billableHoursB = getHours(enterpriseId, resourceTrendJson)
    let billableDaysB = getBillDaysPWA(locationCategory, billableHoursB)
    let grossAmountB = calculateGrossAmountB(billableHoursB, billRate)
    let volDiscountB = calculatevolDiscountB(grossAmountB)
    let strategicInnovFundB = calculateStrategicInnovFundB(grossAmountB, volDiscountB)
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



function createRowsFromInnovationBilling(innovationBillingJson, resourceTrendJson) {
  return innovationBillingJson.map((data) => {
    let stream = data['STREAM']
    let enterpriseId = data['Enterprise ID']
    let locationCategory = data['Location']
    let dailyRate = data['GROSS daily rate']
    let billableHoursA = data["Hrs\r\n(based on myTE)_1"]
    let hrsperday = data["hrs/day"]
    let billableDaysA = calculateBillDaysAInnovationBilling(billableHoursA, hrsperday)
    let grossAmountA = calculateGrossAmountA(dailyRate, billableDaysA)
    let volDiscountA = calculatevolDiscountA(grossAmountA)
    let strategicInnovFundA = calculateStrategicInnovFundA(grossAmountA, volDiscountA)
    let netAmountA = calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA)
    // MME
    let billRate = getBillrate(enterpriseId, resourceTrendJson)
    let billableHoursB = getHours(enterpriseId, resourceTrendJson)
    let billableDaysB = getBillDaysPWA(locationCategory, billableHoursB)
    let grossAmountB = calculateGrossAmountB(billableHoursB, billRate)
    let volDiscountB = calculatevolDiscountB(grossAmountB)
    let strategicInnovFundB = calculateStrategicInnovFundB(grossAmountB, volDiscountB)
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

function calculateDailyRate(onshoreRate, offshoreRate) {
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

function calculateStrategicInnovFundA(grossAmountA, volDiscountA) {
  let percentage = 1.93
  let percentageToMultiplier = (percentage / 100);
  return ((parseInt(grossAmountA) + parseInt(volDiscountA)) * percentageToMultiplier)
}

function calculateStrategicInnovFundB(grossAmountB, volDiscountB) {
  let percentage = 1.93
  let percentageToMultiplier = (percentage / 100);
  return ((grossAmountB + volDiscountB) * percentageToMultiplier)
}

function calculateNetAmountA(grossAmountA, volDiscountA, strategicInnovFundA) {
  return (grossAmountA + volDiscountA + strategicInnovFundA)
}

function calculateNetAmountB(grossAmountB, volDiscountB, strategicInnovFundB) {
  return (grossAmountB + volDiscountB + strategicInnovFundB)
}

function calculateBillDaysVar(billableDaysB, billableDaysA) {
  return Math.abs(billableDaysA - billableDaysB)
}

function calculateNetAmountVar(netAmountB, netAmountA) {
  return (netAmountA - netAmountB)
}

function getBillDaysMME(locationCategory, billableHoursB, fullName) {
  let result = 0
  if (fullName == 'Largados, Mickie Rose') {
    result = billableHoursB / 8
  } else if (fullName == 'Garcia, Julius Armstrong') {
    result = billableHoursB / 9
  } else {
    if (locationCategory == 'Philippines Offshore') {
      result = billableHoursB / 9
    } else if (locationCategory == 'Switzerland Local') {
      result = billableHoursB / 8
    } else if (locationCategory == 'Philippines Landed in CH') {
      result = billableHoursB / 9
    } else if (locationCategory == 'Philippines Landed in SG') {
      result = billableHoursB / 8
    } else if (locationCategory == 'Singapore Local') {
      result = billableHoursB / 8
    } else if (locationCategory == 'India Local') {
      result = billableHoursB / 9
    } else if (locationCategory == 'Switzerland - Manno Local') {
      result = billableHoursB / 8
    } else if (locationCategory == 'Philippines Landed in HK') {
      result = billableHoursB / 7.5
    } else if (locationCategory == 'Hong Kong Local') {
      result = billableHoursB / 7.5
    } else if (locationCategory == 'India Landed in SG') {
      result = billableHoursB / 8
    } else {
      result
    }
  }
  return (result)
}

function getBillDaysPWA(locationCategory, billableHoursB) {
  let result = 0
  if (locationCategory == 'Landed - CH') {
    result = billableHoursB / 9
  } else if (locationCategory == 'Local') {
    result = billableHoursB / 8
  } else if (locationCategory == 'Offshore') {
    result = billableHoursB / 9
  } else {
    result
  }
  return (result)
}

function calculateBillDaysAInnovationBilling(billableHoursA, hrsperday) {
  return (billableHoursA / hrsperday)
}

function distinct(iterable) {
  let set = new Set();
  return iterable.filter(item => {
    let value = JSON.stringify(item);
    if (set.has(value)) {
      return false;
    } else {
      set.add(value);
      return true;
    }
  });
}

function filterByIdStreamRole(rows) {
  return rows.filter((item, index) => rows.findIndex(x =>
    x['Enterprise ID'] === item['Enterprise ID'] && x['STREAM'] === item['STREAM']
    && x['Hermes Role'] === item['Hermes Role']) === index)
}

function roundToTwoDecimal(number) {
  return parseFloat(number.toFixed(2))
}

function formatRows(array) {
  array.forEach(item => {
    item['Net Amount C'] = roundToTwoDecimal(item['Net Amount C'])
    item['Billable Days C'] = roundToTwoDecimal(item['Billable Days C'])
  })
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
  let innovationRows = createRowsFromInnovationBilling(innovationBillingJson, resourceTrendJson)
  let pwaRows = createRowsFromPwa(pwaJson, resourceTrendJson)
  let mergedDistinctRows = distinct([...sbrRows, ...pwaRows, ...innovationRows])
  let finalArray = filterByIdStreamRole(mergedDistinctRows)
  formatRows(finalArray)

  var workbook = xlsx.utils.book_new();
  var worksheet = xlsx.utils.json_to_sheet(finalArray);
  worksheet["!cols"] = [{ wch: 40 }, { wch: 20 }, { wch: 23 }, { wch: 12 }, { wch: 22 }, { wch: 19 }, { wch: 13 }, { wch: 18 }, { wch: 17 }, { wch: 25 }, { wch: 25 },
  { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }, { wch: 25 }];

  xlsx.utils.book_append_sheet(workbook, worksheet, "testSheet");
  xlsx.writeFile(workbook, "textBook.xlsx");
}