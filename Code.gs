/***********************************************************************
Copyright 2019 Google Inc.
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at
http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
Note that these code samples being shared are not official Google
products and are not formally supported.
************************************************************************/

/**
 * @fileoverview Campaign Manager User Tool
 * Custom code to:
 * -- help bulk export user lists
 * -- bulk deactivate CM user profiles
 * -- bulk activate CM user profiles*
 */

/**
 * Add 'Campaign Manager Tools' menu to spreadsheet when user opens it.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Campaign Manager Tools')
      .addItem('List all User Profiles', 'listAllUserProfiles')
      .addItem('Deactivate User Profiles', 'deactivateUserProfilesFromSheet')
      .addItem('Activate User Profiles', 'activateUserProfilesFromSheet')
      .addToUi();
}

/**
 * Get list of accounts that user has access to, then list all user profiles
 *     in each account for which they have user admin rights.
 */
function listAllUserProfiles() {
  SpreadsheetApp.getUi().alert(
      'When you click "OK" you may be asked to grant permission for this ' +
      'code to run.\nPlease select a Google account with admin access to ' +
      'the Campaign Manager accounts for which you\'d like to list the users.');
  // Define list of fields to be output to spreadsheet.
  const fields = ['id', 'name', 'email', 'accountId', 'profileId', 'active'];

  // Get list of profiles for your email address
  Logger.log('Getting User Profiles for authenticated user...');
  let userProfiles = getUserProfiles();

  // Create map to deduplicate output across multiple userProfiles
  const accountUserProfileMap = new Map();

  // Iterate through each userProfile adding accountUserProfiles to output
  for (let i = 0; i < userProfiles.length; i++) {
    // If you have sufficient access, list all of the users on each account
    const userProfile = userProfiles[i];
    Logger.log(Utilities.formatString(
        'Getting User Profiles for all authenticated accounts of ' +
            'Administrator Profile %s...',
        userProfile.id));
    const accountUserProfiles = getAccountUserProfiles(userProfile.id);
    Logger.log(Utilities.formatString(
        '%s Account User Profiles found for %s...', accountUserProfiles.length,
        userProfile.id));

    for (let accountUserProfile of accountUserProfiles) {
      // Store userProfileId with each accountUserProfile for later usage
      accountUserProfile['profileId'] = userProfile.id;
      // Add to deduplicated map to be output
      accountUserProfileMap.set(accountUserProfile.id, accountUserProfile);
    }
  }

  Logger.log(
      accountUserProfileMap.size + ' unique Account User Profiles found...');

  // Output data to sheet in current workbook
  writeToSheet('User Profiles', accountUserProfileMap, fields);
}

/**
 * Read list of CM user profiles from Sheet and activate them
 */
function activateUserProfilesFromSheet() {
  SpreadsheetApp.getUi().alert(
      'When you click "OK" you may be asked to grant permission for ' +
      'this code to run.\nPlease select a Google account with admin access ' +
      'to the Campaign Manager accounts for which you\'d like to activate ' +
      'users.');

  // Read in data from sheet
  const sheetData = readDataFromSheet('Activate User Profiles');
  const accountUserProfileMap =
      getAccountUserProfilesMapFromSheetData(sheetData);

  // Set request details for patching AccountUserProfile
  const request = {'active': true};

  const outputMap = patchAccountUserProfiles(accountUserProfileMap, request);

  // Write log data to sheet
  writeToSheet(
      'Active User Profiles Log', outputMap,
      sheetData.fields.concat(['timestamp']));
}

/**
 * Read list of CM user profiles from Sheet and deactivate them
 */
function deactivateUserProfilesFromSheet() {
  SpreadsheetApp.getUi().alert(
      'When you click "OK" you may be asked to grant permission for ' +
      'this code to run.\nPlease select a Google account with admin access ' +
      'to the Campaign Manager accounts for which you\'d like to deactivate ' +
      'users.');

  // Read in data from sheet
  const sheetData = readDataFromSheet('Deactivate User Profiles');
  const accountUserProfileMap =
      getAccountUserProfilesMapFromSheetData(sheetData);

  // Set request details for patching AccountUserProfile
  const request = {'active': false};

  // Patch AccountUserProfiles
  outputMap = patchAccountUserProfiles(accountUserProfileMap, request);

  // Write log data to sheet
  writeToSheet(
      'Deactivate User Profiles Log', outputMap,
      sheetData.fields.concat(['timestamp']));
}

/**
 * Write data to a Google Sheet
 * @param {string} sheetName - the name of the spreadsheet sheet
 * @param {!Map} outputMap - a Map containing data tuples
 * @param {!Array} fields - array of columns to be written to sheet
 */
function writeToSheet(sheetName, outputMap, fields) {
  const dataToWrite = [];

  // Add Headers to output
  dataToWrite.push(fields);

  // Iterate through outputMap and add the relevant fields to the relevant
  // columns
  for (let row of outputMap) {
    const rowValues = [];
    for (i = 0; i < fields.length; i++) {
      rowValues.push(row[1][fields[i]]);
    }
    dataToWrite.push(rowValues);
  }

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  // Create sheet if it doesn't exist, then activate it
  try {
    Logger.log('Writing data to ' + sheetName + '...');
    spreadSheet.insertSheet(sheetName);
  } catch (err) {
    const regex = /already exists/;
    if (err.toString().match(regex)) {
      Logger.log('Sheet already exists. Clearing data...');
    } else {
      throw (err);
    }
  }
  const sheet = spreadSheet.getSheetByName(sheetName);
  sheet.activate();

  // Clear any previous results then write output
  sheet.clearContents();
  sheet.getRange(1, 1, dataToWrite.length, fields.length)
      .setValues(dataToWrite);
}

/**
 * Transform data read from sheet into a map of deduplicated user profiles
 * @param {!Object} sheetData - object with array of field names and array of
 *     data rows
 * @return {!Object} object with array of field names and array of data rows
 */
function getAccountUserProfilesMapFromSheetData(sheetData) {
  // Convert data table into array of objects with field names
  const dataArrayOfObjs =
      sheetData.data.map(x => getObjectsFromArray(sheetData.fields, x));

  const accountUserProfileMap = new Map();

  // Add rows into a Map for deduplication
  for (i = 0; i < dataArrayOfObjs.length; i++) {
    accountUserProfileMap.set(dataArrayOfObjs[i].id, dataArrayOfObjs[i]);
  }

  return accountUserProfileMap;
}

/**
 * Read data from Google Sheet
 * @param {string} sheetName - the name of the spreadsheet sheet
 * @return {!Object} object with array of field names and array of data rows
 */
function readDataFromSheet(sheetName) {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getSheetByName(sheetName);
  sheet.activate();
  const range = sheet.getDataRange();
  // Read data
  const dataArray = range.getValues();

  // Get field names from header
  let fields = dataArray.shift();

  return {fields: fields, data: dataArray};
}

/**
 * Transform row of data into object with field names
 * @param {!Array} fields - the field names
 * @param {!Array} array - the row of data
 * @return {!Object} the data object
 */
function getObjectsFromArray(fields, array) {
  const obj = {};
  for (i = 0; i < fields.length; i++) {
    obj[fields[i]] = array[i];
  }
  return obj;
}

/**
 * Patch account user profiles
 * @param {!Map} accountUserProfileMap - deduplicated map of user profiles
 * @param {!Object} request - body of the API request
 * @return {!Map} map of user profiles with updated status
 */
function patchAccountUserProfiles(accountUserProfileMap, request) {
  for (let [id, accountUserProfile] of accountUserProfileMap) {
    const profileId = accountUserProfile.profileId;

    try {
      // Use Patch request instead of update, to only update additional key
      // values part of the placement object
      const response = DoubleClickCampaigns.AccountUserProfiles.patch(
          request, profileId, id);

      Logger.log(
          'Successfully patched AccountUserProfile: ' +
          JSON.stringify(accountUserProfile));

      accountUserProfile = response;
      accountUserProfile['timestamp'] = Date.now();
      accountUserProfile['profileId'] = profileId;
      accountUserProfileMap.set(id, accountUserProfile);
    } catch (err) {
      Logger.log(
          'Error patching AccountUserProfile: ' +
          JSON.stringify(accountUserProfile) + '. Error message: ' + err);
    }
  }
  return accountUserProfileMap;
}
