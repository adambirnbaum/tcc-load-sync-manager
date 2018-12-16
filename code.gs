// var isHtmlServiceFinished = false;
// Logger = BetterLog.useSpreadsheet('13ZlDwJ_o2sRoKyoYRzRwc4CnmheVQMrmk2xftxQjUPY');
function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  /****************************************************************************************************************
  *
  * Naming convention
  *
  * Master spreadsheet => Source
  * Trucks Needed spreadsheet => Tn
  * Prepay => Prepay
  *
  *****************************************************************************************************************/
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("3CC Load Sync Manager")
      .addItem("Sync Master to Trucks Needed/Prepay", "menuItem1")
      .addSeparator()
      .addItem("Update \'Week of Ship Date\' List", "menuItem2")
      .addItem("Reset \'Ready to Sync\' Column to \'No\'", "menuItem3")
      .addToUi();
}

function menuItem1() {
  /****************************************************************************************************************
  *
  * This is the main function for running most of the sub-functions required to sync Master to TN and Prepay
  *
  *****************************************************************************************************************/

  // Logger.log('****************** JUST CLICKED REQUEST TO SYNC ******************');

  // Get settings
  var myActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  myActiveSpreadsheet.toast('Loading settings', 'Loading - 0%');
  var settings = getSettings();

  // Pull all source data into memory
  myActiveSpreadsheet.toast('Retrieving data from Master spreadsheet', 'Loading - 25%');
  var allDataFromSource = getAllDataFromSource(settings);

  if (allDataFromSource.length == 0) {
    throw 'No rows have been selected to sync';
  }

  // COPY START option to move this section back behind user confirmation
  myActiveSpreadsheet.toast('Analyzing data from Trucks Needed spreadsheet', 'Loading - 50%');
  var allDataFromTn = getAllDataFromTn(allDataFromSource, settings);

  myActiveSpreadsheet.toast('Analyzing data from Prepay spreadsheet', 'Loading - 75%');
  var allDataFromPrepay = getAllDataFromPrepay(allDataFromSource, settings);
  // COPY END option to move this section back behind user confirmation

  // Confirm with user they would like to go forward with sync
  myActiveSpreadsheet.toast('Loading complete.  Ready to sync', 'Loading - 100%');
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'From the ' + myActiveSpreadsheet.getName() + ' sheet, ' + allDataFromSource.objectRowsDataReadyToSyncToTn.length + ' rows will be synced to Trucks Needed,\nof which ' +
     allDataFromSource.objectRowsDataReadyToSyncToPrepay.length + ' rows will be synced to Prepay.',
     'Are you sure you want to continue with the sync?',
     ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes"
    // Acquire lock to ensure only one instance of the script is being run
    myActiveSpreadsheet.toast('Locking script to prevent simultaneous use', 'Status');
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);   // Wait for up to 10 seconds to acquire lock

    // OPTIONAL - THIS IS WHERE FUNCTIONS UP ABOVE BETWEEN COPY START AND COPY END COULD BE PASTED

    myActiveSpreadsheet.toast('Syncing selected rows from Master to Trucks Needed/Prepay', 'Status');
    var rowsCopiedForTN = syncSourceToTn(allDataFromSource, allDataFromTn, settings);
    var rowsCopiedForPrepay = syncSourceToPrepay(allDataFromSource, allDataFromPrepay, settings);
    lock.releaseLock();

    // Collect info to send confirmation email
    prepToSendConfirmationEmail(rowsCopiedForTN, rowsCopiedForPrepay, ui);


    /*
    var wait = 0;
    var timebetween = 500;                        // 0.5 seconds
    var timeout = 120000;                         // 2 minutes
    while (isHtmlServiceFinished == false) {
      Utilities.sleep(timebetween);
      wait += timebetween;
      if (wait >= timeout) {
        //Logger.log('ERROR: timed out after ' + timeout.toString() + ' seconds.');
        throw 'ERROR: timed out after ' + timeout.toString() + ' seconds.';
      }
    }

    //Logger.log("HTML service just finished");
    */

    // If html modal is submitted successfully, it will call function readyToSendEmail()
    // If html modal cancel button is selected, it will call function cancelSendEmail()

  } else {

    // User clicked "No" or X in the title bar.
  }
}

function prepToSendConfirmationEmail(rowsCopiedForTN, rowsCopiedForPrepay, ui) {
  /****************************************************************************************************************
  *
  * This function starts the HTML Service dialogue box
  *
  *****************************************************************************************************************/
  // Logger.log('**** Start prepToSendConfirmationEmail()... ');
  // isHtmlServiceFinished = true;

  var html = HtmlService.createHtmlOutputFromFile('confirmation-email-options').setHeight(350);
  ui.showModalDialog(html, 'Success! ' + rowsCopiedForTN + ' rows synced to TN and ' + rowsCopiedForPrepay + ' rows synced to Prepay');
}

function readyToSendEmail(isTruckSchedulingNeeded, traderNotes) {
  /****************************************************************************************************************
  *
  * This function gets everything ready to send email and sends confirmation email.  Then resets sync column to no and alerts user
  *
  *****************************************************************************************************************/
  //Logger.log('**** Start readyToSendEmail()... ');
  // Send confirmation email
  SpreadsheetApp.getActiveSpreadsheet().toast('Sending confirmation email', 'Status');

  // >>>> These functions are being run again because they are not passed into this function
  var settings = getSettings();
  var allDataFromSource = getAllDataFromSource(settings);

  var emailSent = sendConfirmationEmail(allDataFromSource, settings.trucksNeededUrl, settings.prepayUrl, settings.recipientEmail, settings.traderEmail, isTruckSchedulingNeeded, traderNotes);

  // Reset sync column and display success dialogue
  resetSyncToNo(allDataFromSource, settings);
  SpreadsheetApp.flush();
  if (emailSent == 1) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Sync complete and email confirmation sent.  Review Trucks Needed to ensure loads were synced properly.  If you synced Prepays, review Prepay spreadsheet to add any prepay notes.');
  } else {
    throw 'Email confirmation failed to send (but loads did sync).'
  }
}

function cancelSendEmail() {
  /****************************************************************************************************************
  *
  * This function runs if user cancels email send when in Html Service dialogue box
  *
  *****************************************************************************************************************/
  // Logger.log('**** Start cancelSendEmail()... ');
  // >>>> These functions are being run again because they are not passed into this function
  SpreadsheetApp.getActiveSpreadsheet().toast('Finishing sync without email send', 'Status');
  var settings = getSettings();
  var allDataFromSource = getAllDataFromSource(settings);

  resetSyncToNo(allDataFromSource, settings);
  SpreadsheetApp.flush();
  var ui = SpreadsheetApp.getUi();
  ui.alert('Sync complete, but no email confirmation was sent.  Review Trucks Needed to ensure loads were synced properly.  If you synced Prepays, review Prepay spreadsheet to add any prepay notes.');
}

function menuItem2() {
  /****************************************************************************************************************
  *
  * This menu command runs the ship date week of function after getting necessary settings, and displays alert after complete
  *
  *****************************************************************************************************************/

  // Confirm with user they would like to update ship date week of function
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Please confirm",
    "Are you sure you want to update the \'Week of Ship Date\' list?",
    ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    var settings = getSettings();
    updateShipDateWeekOfList(settings.trucksNeededUrl, settings.weekOfValidationStartRowNum, settings.weekOfValidationStartColNum);
    SpreadsheetApp.flush();
    ui.alert("\'Ship Date Week of\' list has been updated");
  } else {
     // User clicked "No" or X in the title bar.
  }
}

function menuItem3() {
  /****************************************************************************************************************
  *
  * This menu command runs the reset sync to no function after getting necessary settings, and displays alert after complete
  *
  *****************************************************************************************************************/

  // Confirm with user they would like to reset the sync column
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    "Please confirm",
    "Are you sure you want to reset the \'Ready to Sync\' column to \'No\'?",
     ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    var settings = getSettings();
    var allDataFromSource = getAllDataFromSource(settings);
    resetSyncToNo(allDataFromSource, settings);
    SpreadsheetApp.flush();
    ui.alert("\'Ready to Sync\' column has been successfully reset to \'No\'");
  } else {
    // User clicked "No" or X in the title bar.
  }
}

function getSettings() {
  /****************************************************************************************************************
  *
  * This function gets all settings from the settings sheet for use by other functions
  *
  *****************************************************************************************************************/
  // Logger.log('**** Start getSettings() ...');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if (sheet == null) {
      throw 'Error:  The Settings sheet could not be found';
  }
  var data = sheet.getDataRange().getValues();

  var traderEmail = data[4][4];
  var recipientEmail = data[5][4];

  var sourceHeaderRowNum = data[8][4];
  var sourceStartRowNum = data[9][4];
  var sourceStartColNum = letterToColumn(data[10][4]);
  var sourceEndColNum = letterToColumn(data[11][4]);

  var trucksNeededUrl = data[14][4];

  var tnHeaderRowNum = data[16][4];
  var tnStartRowNum = data[17][4];
  var tnStartColNum = letterToColumn(data[18][4]);
  var tnEndColNum = letterToColumn(data[19][4]);

  var prepayUrl = data[22][4];

  var prepayHeaderRowNum = data[24][4];
  var prepayStartRowNum = data[25][4];
  var prepayStartColNum = letterToColumn(data[26][4]);
  var prepayEndColNum = letterToColumn(data[27][4]);

  var weekOfValidationStartRowNum = data[48][4];
  var weekOfValidationStartColNum = data[49][4];
  var d = new Date();

  var traderRequiredColumnsObject = {
    "truckNo" : "Truck NO.",
    "traderName" : "Trader Name",
    "masterRecord" : "Master Record #",
    "ccCwTradeConfNo" : "3CC CW # / Trade Conf No",
    "readyToSync" : "Ready to Sync?",
    "status" : "Status",
    "acceptedByResponsibleParty" : "Accepted By / Responsible Party",
    "customer" : "Customer",
    "shipTo" : "Ship To",
    "expirationDate" : "Expiration Date",
    "prepaymentDateFromCustomer" : "Prepayment Date from Customer",
    "prepaymentAmountFromCustomer" : "Prepayment Amount from Customer",
    "sellTerms" : "Sell Terms",
    "customerContractNumber" : "Customer Contract Number",
    "customerOffload" : "Customer Offload #",
    "customerProduct" : "Customer Product",
    "customerContractFfa" : "Customer Contract FFA",
    "customerContractMiu" : "Customer Contract MIU",
    "customerQuantityLbs" : "Customer Quantity / Lbs",
    "sellRate" : "Sell / Rate",
    "vendor" : "Vendor",
    "ccTradeConfirmationNumberBuy" : "3CC Trade Confirmation Number BUY",
    "vendorContractNumber" : "Vendor Contract Number",
    "pickup" : "Pickup #",
    "shipFrom" : "Ship From",
    "prepayDateRequestedForVendor" : "Prepay Date Requested for Vendor",
    "prepayAmountForVendor" : "Prepay Amount for Vendor",
    "vendorProduct" : "Vendor Product",
    "vendorContractFfa" : "Vendor Contract FFA",
    "vendorContractMiu" : "Vendor Contract MIU",
    "vendorQuantityLbs" : "Vendor Quantity / Lbs",
    "buyRate" : "Buy / Rate",
    "buyTerms" : "Buy Terms",
    "broker" : "Broker",
    "brokerNumber" : "Broker Number",
    "brokerFeeInDollars" : "Broker Fee in Dollars",
    "weekOfShipDate" : "Week of Ship Date",
    "estimatedShipDate" : "Estimated Ship Date",
    "bolSentToAllParties" : "BOL sent to all parties?",
    "bolProduct" : "BOL Product",
    "billingWeights" : "Billing Weights"};

  var traderRequiredColumnNames = Object.keys(traderRequiredColumnsObject);

  return {traderEmail:traderEmail,
          recipientEmail:recipientEmail,

          sourceHeaderRowNum:sourceHeaderRowNum,
          sourceStartRowNum:sourceStartRowNum,
          sourceStartColNum:sourceStartColNum,
          sourceEndColNum:sourceEndColNum,

          trucksNeededUrl:trucksNeededUrl,

          tnHeaderRowNum:tnHeaderRowNum,
          tnStartRowNum:tnStartRowNum,
          tnStartColNum:tnStartColNum,
          tnEndColNum:tnEndColNum,

          prepayUrl:prepayUrl,

          prepayHeaderRowNum:prepayHeaderRowNum,
          prepayStartRowNum:prepayStartRowNum,
          prepayStartColNum:prepayStartColNum,
          prepayEndColNum:prepayEndColNum,

          weekOfValidationStartRowNum:weekOfValidationStartRowNum,
          weekOfValidationStartColNum:weekOfValidationStartColNum,
          d:d,

          traderRequiredColumnsObject:traderRequiredColumnsObject,
          traderRequiredColumnNames:traderRequiredColumnNames};
}

function getAllDataFromSource(settings) {
  /****************************************************************************************************************
  *
  * This function pulls all data from Source into memory.  Additionally, it creates the various arrays
  *        used by other functions
  *
  *****************************************************************************************************************/

  // Logger.log('**** Start getAllDataFromSource() ...');
  var sourceHeaderNames = [];
  var sourceStartRowNum = settings.sourceStartRowNum;
  var sourceStartColNum = settings.sourceStartColNum;
  var sourceTotalColumns = settings.sourceEndColNum - sourceStartColNum + 1;
  var sourceHeaderRowNum = settings.sourceHeaderRowNum;

  // Pull in all data from Source
  var sheetSource = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheetSource.getName() == "Settings" || sheetSource.getName() == "Master Template TO DUPLICATE") {
    throw "Error:  Please open the sheet in Master you would like to sync from";
  }

  var range = sheetSource.getRange(sourceStartRowNum, sourceStartColNum, sheetSource.getLastRow(), sourceTotalColumns);
  var objectRowsData = getRowsData(sheetSource, range, sourceHeaderRowNum);

  var headersRange = sheetSource.getRange(sourceHeaderRowNum, sourceStartColNum, 1, sourceTotalColumns);
  var headers = headersRange.getValues()[0];
  sourceHeaderNames = normalizeHeaders(headers);

  // filter to keep only rows that are Ready to Sync
  var objectRowsDataReadyToSyncToTn = objectRowsData.filter(function(x) {
    return x.readyToSync == "Yes";
  });

  if (objectRowsDataReadyToSyncToTn.length < 1) {
    throw 'Error: Must select at least one load to sync';
  }

  confirmTraderRequiredDataIsComplete(settings.traderRequiredColumnNames, settings.traderRequiredColumnsObject, objectRowsDataReadyToSyncToTn);

  for (i = 0; i < objectRowsDataReadyToSyncToTn.length; i++) {
    if ((isValidDate(objectRowsDataReadyToSyncToTn[i]["prepayDateRequestedForVendor"]) && !(isNumber(objectRowsDataReadyToSyncToTn[i]["prepayAmountForVendor"]))) || (!(isValidDate(objectRowsDataReadyToSyncToTn[i]["prepayDateRequestedForVendor"])) && (isNumber(objectRowsDataReadyToSyncToTn[i]["prepayAmountForVendor"])))) {
      throw 'Error: ' + objectRowsDataReadyToSyncToTn[i]["masterRecord"] + ' should have both vendor prepay date and amount, or else enter NA for both.';
    }
  }

  var objectRowsDataReadyToSyncToPrepay = objectRowsDataReadyToSyncToTn.filter(function(x) {
    return isValidDate(x.prepayDateRequestedForVendor) && isNumber(x.prepayAmountForVendor);
  });

  var weekofShipDatesToSyncToTn = getArrayFromObject(objectRowsDataReadyToSyncToTn, "weekOfShipDate")
  var weekofShipDatesToSyncToPrepay = getArrayFromObject(objectRowsDataReadyToSyncToPrepay, "weekOfShipDate")

  var allMasterRecordsInSource = sheetSource.getRange(sourceStartRowNum, sourceStartColNum + sourceHeaderNames.indexOf("masterRecord"), sheetSource.getLastRow(), sourceTotalColumns).getValues().map(function (row) { return row[0]; });

  return {objectRowsData:objectRowsData,
          objectRowsDataReadyToSyncToTn:objectRowsDataReadyToSyncToTn,
          objectRowsDataReadyToSyncToPrepay:objectRowsDataReadyToSyncToPrepay,
          weekofShipDatesToSyncToTn:weekofShipDatesToSyncToTn,
          weekofShipDatesToSyncToPrepay:weekofShipDatesToSyncToPrepay,
          sourceHeaderNames:sourceHeaderNames,
          allMasterRecordsInSource:allMasterRecordsInSource};
}

function getAllDataFromTn(allDataFromSource, settings) {
  /****************************************************************************************************************
  *
  * This function gets object with a key of the sheet name and value of an array of the master IDs on that sheet.
  *
  *****************************************************************************************************************/
  // Logger.log('**** Start getAllDataFromTn ...');
  var tnHeadersNames = {};
  var objectRowsTnData = {};
  var ss = SpreadsheetApp.openByUrl(settings.trucksNeededUrl);

  var uniqueWeekofShipDatesToSyncToTn = allDataFromSource.weekofShipDatesToSyncToTn.getUnique();

  var tnStartRowNum = settings.tnStartRowNum;
  var tnStartColNum = settings.tnStartColNum;
  var tnTotalColumns = settings.tnEndColNum - tnStartColNum + 1;
  var tnHeaderRowNum = settings.tnHeaderRowNum;

  if (ss == null) {        //
      throw 'Error:  Trucks Needed spreadsheet could not be opened.  Check to ensure the right sheet Url is used in Settings tab';
  }

  // get object with list of header names for each tab in Tn
  for (i = 0; i < uniqueWeekofShipDatesToSyncToTn.length; i++) {
    var sheet = ss.getSheetByName(uniqueWeekofShipDatesToSyncToTn[i]);
    if (sheet == null) {
      throw 'Error:  The sheet \"' + uniqueWeekofShipDatesToSyncToTn[i] + '\" does not exist in Trucks Needed.';
    }
    var dataRange = sheet.getRange(tnStartRowNum, tnStartColNum, sheet.getLastRow(), tnTotalColumns);
    objectRowsTnData[uniqueWeekofShipDatesToSyncToTn[i]] = getRowsData(sheet, dataRange, tnHeaderRowNum);

    var headersRange = sheet.getRange(tnHeaderRowNum, tnStartColNum, 1, tnTotalColumns);
    var headers = headersRange.getValues()[0];
    tnHeadersNames[uniqueWeekofShipDatesToSyncToTn[i]] = normalizeHeaders(headers);

    if (tnHeadersNames[uniqueWeekofShipDatesToSyncToTn[i]].indexOf("masterRecord") == -1) {
      throw "Error:  Could not find masterRecord column in the Trucks Needed tab '" + uniqueWeekofShipDatesToSyncToTn[i] + "'";
    }
  }
  return {objectRowsTnData:objectRowsTnData,
          tnHeadersNames:tnHeadersNames};
}

function getAllDataFromPrepay(allDataFromSource, settings) {
  /****************************************************************************************************************
  *
  * This function gets object with a key of the sheet name and value of an array of the master IDs on that sheet.
  *
  *****************************************************************************************************************/

  // Logger.log('**** Start getAllDataFromPrepay ...');
  var prepayHeadersNames = {};
  var objectRowsPrepayData = {};
  var ss = SpreadsheetApp.openByUrl(settings.prepayUrl);

  var uniqueWeekofShipDatesToSyncToPrepay = allDataFromSource.weekofShipDatesToSyncToPrepay.getUnique();

  var prepayStartRowNum = settings.prepayStartRowNum;
  var prepayStartColNum = settings.prepayStartColNum;
  var prepayTotalColumns = settings.prepayEndColNum - prepayStartColNum + 1;
  var prepayHeaderRowNum = settings.prepayHeaderRowNum;

  if (ss == null) {        //
      throw 'Error:  Prepay spreadsheet could not be opened.  Check to ensure the right sheet Url is used in Settings tab';
  }

  // get object with list of header names for each tab in Prepay
  for (i = 0; i < uniqueWeekofShipDatesToSyncToPrepay.length; i++) {
    var sheet = ss.getSheetByName(uniqueWeekofShipDatesToSyncToPrepay[i]);
    if (sheet == null) {
      throw 'Error:  The sheet \"' + uniqueWeekofShipDatesToSyncToPrepay[i] + '\" does not exist in Prepay.';
    }
    var headersRange = sheet.getRange(prepayHeaderRowNum, prepayStartColNum, 1, prepayTotalColumns);
    var headers = headersRange.getValues()[0];
    prepayHeadersNames[uniqueWeekofShipDatesToSyncToPrepay[i]] = normalizeHeaders(headers);

    if (prepayHeadersNames[uniqueWeekofShipDatesToSyncToPrepay[i]].indexOf("masterRecord") == -1) {
      throw "Error:  Could not find masterRecord column in the Prepay tab '" + uniqueWeekofShipDatesToSyncToPrepay[i] + "'";
    }
  }
  return {prepayHeadersNames:prepayHeadersNames};
}

function syncSourceToTn(allDataFromSource, allDataFromTn, settings) {
  // Logger.log("**** Starting syncSourceToTn() ...");
  // Initialize variables
  var firstRowInTnWeekOfTab = {};
  var masterRecordsInTn = {};
  var numberOfRowsCopiedToTn = 0;
  var sheetMaster = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ssTn = SpreadsheetApp.openByUrl(settings.trucksNeededUrl);
  var d = settings.d;

  var tnStartColNum = settings.tnStartColNum;
  var tnTotalColumns = settings.tnEndColNum - tnStartColNum + 1;
  var objectRowsDataReadyToSyncToTn = allDataFromSource.objectRowsDataReadyToSyncToTn;
  var weekofShipDatesToSyncToTn = allDataFromSource.weekofShipDatesToSyncToTn;
  var tnHeadersNames = allDataFromTn.tnHeadersNames;
  var objectRowsTnData = allDataFromTn.objectRowsTnData;

  var uniqueWeekofShipDatesToSyncToTn = weekofShipDatesToSyncToTn.getUnique();

  for (i = 0; i < uniqueWeekofShipDatesToSyncToTn.length; i++) {
    masterRecordsInTn[uniqueWeekofShipDatesToSyncToTn[i]] = getArrayFromObject(objectRowsTnData[uniqueWeekofShipDatesToSyncToTn[i]], "masterRecord");  // >>>> NOTE BY ASSUMING INDEX 0 ALL HEADERS NEED TO BE IDENTICAL FOR TN SHEETS
  }
  for (i = 0; i < objectRowsDataReadyToSyncToTn.length; i++) {
    var tnData = [];
    tnData[0] = [];
    for (j = 0; j < tnHeadersNames[weekofShipDatesToSyncToTn[0]].length; j++) {  // >>>> NOTE BY ASSUMING INDEX 0 ALL HEADERS NEED TO BE IDENTICAL FOR TN SHEETS
      if (tnHeadersNames[weekofShipDatesToSyncToTn[0]][j] == "dateLastSynced") {
        tnData[0].push(d);
      } else if (tnHeadersNames[weekofShipDatesToSyncToTn[0]][j] in objectRowsDataReadyToSyncToTn[i]) {
        tnData[0].push(objectRowsDataReadyToSyncToTn[i][tnHeadersNames[weekofShipDatesToSyncToTn[0]][j]]);
      } else {
        tnData[0].push("");
      }
    }
    var sheetTn = ssTn.getSheetByName(weekofShipDatesToSyncToTn[i]);
    if (!(weekofShipDatesToSyncToTn[i] in firstRowInTnWeekOfTab)) {
      var destinationRange = sheetTn.getRange(sheetTn.getLastRow() + 2, tnStartColNum, 1, tnTotalColumns);
      firstRowInTnWeekOfTab[weekofShipDatesToSyncToTn[i]] = true;
    } else {
      var destinationRange = sheetTn.getRange(sheetTn.getLastRow() + 1, tnStartColNum, 1, tnTotalColumns);
    }
    destinationRange.setValues(tnData);
    numberOfRowsCopiedToTn++;
    if (masterRecordsInTn[weekofShipDatesToSyncToTn[i]].indexOf(objectRowsDataReadyToSyncToTn[i]["masterRecord"]) > -1) {
      destinationRange.setBackground('#ffff00');
    }
    setDateLastSyncedInMaster(objectRowsDataReadyToSyncToTn[i]["masterRecord"], allDataFromSource, settings);
  }
  return numberOfRowsCopiedToTn;
}

function setDateLastSyncedInMaster(masterRecord, allDataFromSource, settings) {
  var sheetMaster = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sourceStartRowNum = settings.sourceStartRowNum;
  var sourceStartColNum = settings.sourceStartColNum;
  var d = settings.d;
  var sourceHeaderNames = allDataFromSource.sourceHeaderNames;
  var masterRecordRowLocation = allDataFromSource.allMasterRecordsInSource.indexOf(masterRecord);
  var dateLastSyncedColLocation = sourceHeaderNames.indexOf("dateLastSynced");
  sheetMaster.getRange(masterRecordRowLocation + sourceStartRowNum, dateLastSyncedColLocation + sourceStartColNum, 1, 1).setValue(d);
}

function syncSourceToPrepay(allDataFromSource, allDataFromPrepay, settings) {
  // Logger.log("**** Starting syncSourceToPrepay() ...");
  // Initialize variables
  var firstRowInPrepayWeekOfTab = {};
  var numberOfRowsCopiedToPrepay = 0;
  var sheetMaster = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ssPrepay = SpreadsheetApp.openByUrl(settings.prepayUrl);
  var d = settings.d;

  var prepayStartColNum = settings.prepayStartColNum;
  var prepayTotalColumns = settings.prepayEndColNum - prepayStartColNum + 1;
  var objectRowsDataReadyToSyncToPrepay = allDataFromSource.objectRowsDataReadyToSyncToPrepay;
  var weekofShipDatesToSyncToPrepay = allDataFromSource.weekofShipDatesToSyncToPrepay;
  var prepayHeadersNames = allDataFromPrepay.prepayHeadersNames;

  // var objectRowsPrepayData = allDataFromPrepay.objectRowsPrepayData;
  // var masterRecordsInTn = getArrayFromObject(objectRowsTnData, "masterRecord");

  for (i = 0; i < objectRowsDataReadyToSyncToPrepay.length; i++) {
    var prepayData = [];
    prepayData[0] = [];
    for (j = 0; j < prepayHeadersNames[weekofShipDatesToSyncToPrepay[0]].length; j++) {  // >>>> NOTE BY ASSUMING INDEX 0 ALL HEADERS NEED TO BE IDENTICAL FOR TN SHEETS
      if (prepayHeadersNames[weekofShipDatesToSyncToPrepay[0]][j] == "dateLastSynced") { // Prepay does not currently have this column, so this will be skipped
        prepayData[0].push(d);
      } else if (prepayHeadersNames[weekofShipDatesToSyncToPrepay[0]][j] in objectRowsDataReadyToSyncToPrepay[i]) {
        prepayData[0].push(objectRowsDataReadyToSyncToPrepay[i][prepayHeadersNames[weekofShipDatesToSyncToPrepay[0]][j]]);
      } else {
        prepayData[0].push("");
      }
    }
    var sheetPrepay = ssPrepay.getSheetByName(weekofShipDatesToSyncToPrepay[i]);
    if (!(weekofShipDatesToSyncToPrepay[i] in firstRowInPrepayWeekOfTab)) {
      var destinationRange = sheetPrepay.getRange(sheetPrepay.getLastRow() + 1, prepayStartColNum, 1, prepayTotalColumns);
      firstRowInPrepayWeekOfTab[weekofShipDatesToSyncToPrepay[i]] = true;
    } else {
      var destinationRange = sheetPrepay.getRange(sheetPrepay.getLastRow() + 1, prepayStartColNum, 1, prepayTotalColumns);
    }
    destinationRange.setValues(prepayData);
    destinationRange.setBackgroundColor('#ff0000');
    numberOfRowsCopiedToPrepay++;
  }
  return numberOfRowsCopiedToPrepay;
}

function sendConfirmationEmail(allDataFromSource, trucksNeededUrl, prepayUrl, recipientEmail, traderEmail, isTruckSchedulingNeeded, traderNotes) {
  /****************************************************************************************************************
  *
  * This function sends a confirmation email with info about the loads that have been synced
  *
  *****************************************************************************************************************/
  // Note quota limit of 1500 recipients / day per https://script.google.com/dashboard
  // Logger.log("**** sendConfirmationEmail() ...");

  // Initialize variables
  var isPrepay = "";
  var isPrepayCount = 0;
  var emailTable = '<table width="600" style="border:1px solid #333"><tr><th>3CC #</th><th>Prepay?</th><th>Week of</th><th>Est Ship Date</th><th>Customer</th><th>Ship To</th><th>Vendor</th><th>Ship From</th></tr>';
  var objectRowsDataReadyToSyncToTn = allDataFromSource.objectRowsDataReadyToSyncToTn;
  var masterRecordsSyncedToTn = getArrayFromObject(objectRowsDataReadyToSyncToTn, "masterRecord");

  for (i = 0; i < objectRowsDataReadyToSyncToTn.length; i++) {
    if (isValidDate(objectRowsDataReadyToSyncToTn[i]["prepayDateRequestedForVendor"]) && isNumber(objectRowsDataReadyToSyncToTn[i]["prepayAmountForVendor"]))  {
      isPrepay = "Yes";
      isPrepayCount++;
    } else {
      isPrepay = "No";
    }
    if (isValidDate(objectRowsDataReadyToSyncToTn[i]["estimatedShipDate"])) {
      var estShipDateForEmail = objectRowsDataReadyToSyncToTn[i]["estimatedShipDate"].toLocaleDateString('en-US');
    } else {
      var estShipDateForEmail = objectRowsDataReadyToSyncToTn[i]["estimatedShipDate"];
    }
    emailTable += ('<tr><td>' + objectRowsDataReadyToSyncToTn[i]["masterRecord"] +
                  '</td><td>' + isPrepay +
                  '</td><td>' + objectRowsDataReadyToSyncToTn[i]["weekOfShipDate"] +
                  '</td><td>' + estShipDateForEmail +
                  '</td><td>' + objectRowsDataReadyToSyncToTn[i]["customer"] +
                  '</td><td>' + objectRowsDataReadyToSyncToTn[i]["shipTo"] +
                  '</td><td>' + objectRowsDataReadyToSyncToTn[i]["vendor"] +
                  '</td><td>' + objectRowsDataReadyToSyncToTn[i]["shipFrom"] +
                  '</td></tr>');
  }
  emailTable += '</table><p>Feel free to reply to this email with any questions for the trader.</p>';
  if (isPrepayCount > 0) {
    var emailPrepay = '<p>There <strong>are prepay(s)</strong> to review in the <a href="' + prepayUrl + '">Prepay Sheet</a>.</p>';
  } else {
    var emailPrepay = '<p>There <strong>are not</strong> any new prepays to review.</p>';
  }
  if (isTruckSchedulingNeeded == "Yes") {
    var emailTruckScheduling = '<p>Truck(s) <strong>do</strong> need to be scheduled.</p>';
  } else {
    var emailTruckScheduling = '<p>Truck(s) <strong>do not</strong> need to be scheduled.</p>';
  }
  if (traderNotes == "") {
    var emailTraderNotes = '<p>There were no additional notes from the trader.</p>';
  } else {
    var emailTraderNotes = '<p>Additional notes from trader: ' + traderNotes + '</p>';
  }

  //Generate the subject and body of the email
  var emailSubject = masterRecordsSyncedToTn.join(", ") + ' ready for review';
  var emailBody = '<p>' + masterRecordsSyncedToTn.join(", ") + ' ready for review in the <a href="' + trucksNeededUrl + '">Trucks Needed Sheet</a>.' +
                  emailPrepay +
                  emailTruckScheduling +
                  emailTraderNotes +
                  '<p>The following loads were synced:</p>' +
                  emailTable;
  //Logger.log('emailTable = ' + emailTable);
  //Logger.log('emailBody = ' + emailBody);
  MailApp.sendEmail({
    to: recipientEmail,
    cc: traderEmail,
    replyTo: traderEmail,
    subject: emailSubject,
    htmlBody: emailBody
  });
  return 1;
}

function resetSyncToNo(allDataFromSource, settings) {
  // Logger.log("**** resetSyncToNo() ...");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sourceHeaderNames = allDataFromSource.sourceHeaderNames;
  if (sheet.getName() == "Settings") { throw "Error:  Please open the sheet in Master for which you would like to reset sync settings"; }
  var readyToSyncColLocation = sourceHeaderNames.indexOf("readyToSync");
  if (readyToSyncColLocation === -1) { throw "Unable to reset \'Ready to Sync\' column because column was not found"; }
  var readyToSyncColNumber = readyToSyncColLocation + settings.sourceStartColNum;
  var range = sheet.getRange(settings.sourceStartRowNum,readyToSyncColNumber,sheet.getLastRow()-1,1);
  range.setValue("No");
}

function updateShipDateWeekOfList(trucksNeededUrl, weekOfValidationStartRowNum, weekOfValidationStartColNum) {
  var sheetNameArray = [];
  var tnSheets = SpreadsheetApp.openByUrl(trucksNeededUrl).getSheets();
  sheetNameArray = tnSheets
    .filter(function(x) {
      return !x.isSheetHidden();     // filter out any hidden sheets
    })
    .map(function(y) {
      return [y.getName()];
    });
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  sheet.getRange(weekOfValidationStartRowNum, weekOfValidationStartColNum, 100, 1).clearContent();         // Will clear up to 100 rows
  var range = sheet.getRange(weekOfValidationStartRowNum, weekOfValidationStartColNum, sheetNameArray.length, 1);
  range.setValues(sheetNameArray);
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// ********************************              Helper Functions              *************************************************
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

Array.prototype.getUnique = function() {
   var u = {}, a = [];
   for(var i = 0, l = this.length; i < l; ++i){
      if(u.hasOwnProperty(this[i])) {
         continue;
      }
      a.push(this[i]);
      u[this[i]] = 1;
   }
   return a;
}

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function isValidDate(d) {
  if ( Object.prototype.toString.call(d) === "[object Date]" ) {
    // it is a date
    if ( isNaN( d.getTime() ) ) {  // d.valueOf() could also work
      // date is not valid
      return false;
    }
    else {
      // date is valid
      return true;
    }
  }
  else {
    // not a date
    return false;
  }
}

function isNumber(val) {
  return !isNaN(parseFloat(val)) && isFinite(val);
}

function getArrayFromObject(obj, key) {
  var a = [];
  for (var i = 0; i < obj.length; i++) {
    a.push(obj[i][key]);
  }
  return a;
}

function copyTo(source, tn, cellColor) {
  if (typeof cellColor === 'undefined') { cellColor = 'undefined'; }
  var sourceSheet = source.getSheet();
  var tnSheet = tn.getSheet();
  var sourceData = source.getValues();
  var tnRange = tnSheet.getRange(
    tn.getRow(),        // Top row of Tn
    tn.getColumn(),     // left col of Tn
    sourceData.length,           // # rows in source
    sourceData[0].length);       // # cols in source (elements in first row)
  tnRange.setValues(sourceData);
  if (cellColor != 'undefined') {
    tnRange.setBackgroundColor(cellColor);
  }
  // SpreadsheetApp.flush();
}

function getArrayWithoutBlanksAtEnd(myArray) {
  var arrayNonBlankIndexes = [];
  myArray.forEach(function(element, index) {
    if (element != "") {
      arrayNonBlankIndexes.push(index);
    }
  });
  var lastRowNum = arrayNonBlankIndexes[arrayNonBlankIndexes.length - 1] + 1;
  if (isNaN(lastRowNum)) {
    lastRowNum = 1;
  }
  var slicedArray = myArray.slice(0, lastRowNum);       // when running slice end is not included
  return slicedArray;
}

function confirmTraderRequiredDataIsComplete(traderRequiredColumnNames, traderRequiredColumnsObject, objectRowsData) {
  //Logger.log("At beginning, headerNameArray = " + headerNameArray);
  //Logger.log("At beginning, objectRowsData = " + objectRowsData);
  for (var i = 0; i < objectRowsData.length; i++) {
    for (var j = 0; j < traderRequiredColumnNames.length; j++) {
      //Logger.log("objectRowsData[" + i + "] = " + objectRowsData[i]);
      //Logger.log("headerNameArray[" + j + "] = " + headerNameArray[j]);
      if (!(traderRequiredColumnNames[j] in objectRowsData[i])) {
        throw "The syncing load " + objectRowsData[i]["masterRecord"] + " is missing required data in the " + traderRequiredColumnsObject[traderRequiredColumnNames[j]] + " column";
      }
      if (objectRowsData[i][traderRequiredColumnNames[j]] == "") {
        throw "The syncing load " + objectRowsData[i]["masterRecord"] + " is missing required data in the " + traderRequiredColumnsObject[traderRequiredColumnNames[j]] + " column";
      }
    }
  }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Sheet processing library functions from https://developers.google.com/apps-script/articles/mail_merge#section4
//
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}
