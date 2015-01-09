/*
* 
* NAME:
* 
* Fio Bank API Transaction List to Google Spreadsheet
* 
* LICENSE:
* 
* Copyright (C) 2015 Václav VESELÝ ⊂ ICTOI, s.r.o.; www.ictoi.com
* 
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
* 
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License for more details.
* 
* You should have received a copy of the GNU General Public License
* along with this program.  If not, see <http://www.gnu.org/licenses/>.
*
*/

/**
* FIO bank API token
* @type {string}
* @const
*/
var FIO_IMPORT_TOKEN = "";

/*
 * gets bank account data from fio
 * @param {string} fio bank account token
 *
 * http://www.fio.cz/docs/cz/API_Bankovnictvi.pdf
 * http://www.fio.cz/bankovni-sluzby/api-bankovnictvi
 *
 */
function getFioBankApiTransactionList(fioToken) {

  // show toast message
  SpreadsheetApp.getActiveSpreadsheet().toast("Wait until script ends.", ":()", 4);

  // check variable defined and set default if not
  fioToken = (fioToken || FIO_IMPORT_TOKEN); // FIO_IMPORT_TOKEN script wide constant of FIO token

  var dateToday = new Date();
  //fromDate = (fromDate || dateToday);
  var fromDate = new Date();
  fromDate.setDate(dateToday.getDate() - 90);
  fromDate = Utilities.formatDate(fromDate, "Europe/Prague", "yyyy-MM-dd");
  //toDate = (toDate || Utilities.formatDate(dateToday, "Europe/Prague", "yyyy-MM-dd"));
  var toDate = (toDate || Utilities.formatDate(dateToday, "Europe/Prague", "yyyy-MM-dd"));

  // get date hash
  var fetchTimestamp = Utilities.formatDate(new Date(), "Europe/Prague", "yyyy-MM-dd_hh:mm:ss");
  var tokenHash = Utilities.base64Encode(fromDate + fetchTimestamp + fioToken);

  // catch exception
  try {
    // cache handler
    var urlContent = null;
    var privateCache = CacheService.getPrivateCache();
    //publicCache.remove(addressHash);
    var cacheContent = privateCache.get(tokenHash);
    if (cacheContent != null) {
      urlContent = cacheContent;
    } else {
      // fetch source url    
      var urlResponse = UrlFetchApp.fetch(encodeURI("https://www.fio.cz/ib_api/rest/periods/" + fioToken + "/" + fromDate + "/" + toDate + "/transactions.json"));
      var urlResponseCode = urlResponse.getResponseCode();
      var urlContent = urlResponse.getContentText();

      // cache for one day
      privateCache.put(tokenHash, urlContent, 86400);
    }
  } catch (e) {
    var aEx = "Oops! Can not get source URL content.";
    Logger.log(e + " / " + aEx);
    throw aEx;
  }

  // catch exception
  try {
    // parse json
    var parsedJson = Utilities.jsonParse(urlContent);
    //var parsedXml = Xml.parse(urlContent, false);

    //get root elements
    var fioIban = parsedJson.accountStatement.info.iban;
    var fioTransactions = parsedJson.accountStatement.transactionList.transaction;

    // set variables
    var resultArray = [];
    var resultRow = [];
    var tempArray = [];

    // reverse loop results
    var i = fioTransactions.length;
    while (i--) {
      // push iban
      resultRow.push(fioIban);
      for (var actualColumn in fioTransactions[i]) {
        var actualValue = fioTransactions[i][actualColumn];
        if (actualValue == null) {
          resultRow.push("");
        } else {
          var pushValue = null;
          var switchValue = actualValue.id;
          // specific value switch
          switch (switchValue) {
          case 0: // Datum
            var dateArray = actualValue.value.split("-");
            pushValue = new Date(dateArray[0], (dateArray[1] - 1), dateArray[2].substring(0, 2), 12, 0, 0);
            break;
          case 9: // Provedl
            var nameArray = actualValue.value.split(", ");
            pushValue = nameArray[1] + " " + nameArray[0].toUpperCase();
            break;
          case 25: // Komentář
            pushValue = "";
            break;
          default:
            pushValue = actualValue.value;
          }

          // push to array
          resultRow.push(pushValue);
        }
      }

      // push to array
      resultArray.push(resultRow);
      resultRow = [];
    }

  } catch (e) {
    var aEx = "Oops! Can not parse FIO data.";
    Logger.log(e + " / " + aEx);
    throw aEx;
  }

  // catch exception
  try {

    // get data range
    var aSht = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMPORT_SHEET_NAME); // IMPORT_SHEET_NAME script wide constant of fio transaction import sheet name
    var aShtData = aSht.getDataRange();
    var aShtLRow = aShtData.getLastRow();
    var aShtLCol = aShtData.getLastColumn();
    var aShtLastPaymentId = aSht.getRange(2, 3, 1, 1).getValues(); //last payment id

    // find last filled row
    var curPaymentId = null;
    var lastPaymentIndex = null;
    for (var k = 0, len = resultArray.length; k < len; k++) {
      curPaymentId = resultArray[k][2];
      if (curPaymentId == aShtLastPaymentId[0][0]) {
        lastPaymentIndex = k;
        break;
      }
    }

    // slice array according to last filled row
    resultArray = resultArray.slice(0, lastPaymentIndex);

    // clear data range
    //var aShtClearRange = aSht.getRange(2, 1, aShtLRow, aShtLCol);
    //aShtClearRange.clearContent();
    aSht.insertRows(2, lastPaymentIndex)

    // flush
    SpreadsheetApp.flush();

    // fill data range
    var aShtFillRange = aSht.getRange(2, 1, lastPaymentIndex, resultArray[0].length);
    //var aShtFillRange = aSht.getRange(2, 1, resultArray.length, 21);
    aShtFillRange.setValues(resultArray);

    // flush
    SpreadsheetApp.flush();

  } catch (e) {
    var aEx = "Oops! Can not write data to spreadsheet.";
    Logger.log(e + " / " + aEx);
    throw aEx;
  }

  // show toast message
  SpreadsheetApp.getActiveSpreadsheet().toast("Hurray! Success.", ":()", 4);

}

/*
 * returns flatten "1D" array
 * @param {array} inputArray array to flatten
 * @return {array} flattenArray flatten array
 */
function getFlattenArray(inputArray) {

  var flattenArray = [];
  for (var i = 0, l = inputArray.length; i < l; i++) {
    var variableType = Object.prototype.toString.call(inputArray[i]).split(' ').pop().split(']').shift().toLowerCase();
    if (variableType) {
      flattenArray = flattenArray.concat(/^(array|collection|arguments|object)$/.test(variableType) ? getFlattenArray(inputArray[i]) : inputArray[i]);
    }
  }
  return flattenArray;
}
