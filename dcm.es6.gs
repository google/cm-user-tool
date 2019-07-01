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
 * @typedef {{ items: !Array.<!Object>, nextPageToken: string }}
 * @private
 */
let PageData;
/**
 * @typedef {{ id: number, accountId: number}}
 */
let UserProfileItem;
/**
 * @typedef {{
 *  status: string,
 *  objectIds: !Array.<string>
 * }}
 */
let AdvertiserFilter;
/**
 * @typedef {{
 *  id: number,
 *  active: boolean,
 *  email: string,
 *  name: string,
 *  accountId: number,
 *  subaccountId: number,
 *  advertiserFilter: *,
 *  userRoleId: number
 * }}
 */
let AccountUserProfileItem;
/**
 * @typedef {{
 *   id: number,
 *   name: string
 * }}
 */
let MinimalItem;
/**
 * A minimal implementation for selecting only id and name properties from
 * DCM API entities.
 *
 * @param {!Object} item - the item to transform
 * @return {!MinimalItem} the transformed object
 */
const MINIMAL_ITEM_TRANSFORM = (item) => ({id: item.id, name: item.name});
/**
 * Indicates the maximum number of pages to fetch in a single call.
 * @type {number}
 */
const MAX_PAGES_TO_FETCH = 10;
const EXCLUDED_ACCOUNTS = {
    //'2515': true,
    // '7480': true
};

/**
 * This utility function fetches all resource items from the DCM API by
 * accumulating subsequent page fetches.
 *
 * @private
 *
 * @param {function(!Object):!PageData} fetchPageFn - a function
 *   which implements a single list call to the DCM API
 * @param {!Object=} params - optional additional params to be sent with the API
 *   call
 * @param {(function(!Object):!{id: number})=} transformItemFn - an optional
 *   function to transform response entities
 * @return {!Array.<!Object>} the list of entities
 */
function fetchAll_(
    fetchPageFn, params = undefined, transformItemFn = undefined) {
  let page;
  let pageParams = Object.assign({}, params);
  let items = [];
  let threshold = MAX_PAGES_TO_FETCH;

  do {
    page = fetchPageFn(pageParams);
    items = items.concat(page.items);
    pageParams.pageToken = page.nextPageToken;
  } while (page.nextPageToken && threshold--);

  if (transformItemFn) {
    items = items.map(transformItemFn);
  }

  return items;
}

/**
 * Creates a page data Object from a DCM API response
 * @param {!Object} response - the response object from a DCM API call
 * @param {string} itemsPropName - the name of the property holding the
 *   respons items
 * @return {!PageData}
 */
function makePageData_(response, itemsPropName) {
  return {
    items: response[itemsPropName],
    nextPageToken: response.nextPageToken
  };
}

/**
 * Fetches all user profiles.
 *
 * @return {!Array.<!UserProfileItem>} the list of profiles
 */
function getUserProfiles() {
  return fetchAll_(
             (params) => makePageData_(
                 DoubleClickCampaigns.UserProfiles.list(), 'items'),
             undefined,
             (profile) =>
                 ({id: profile.profileId, accountId: profile.accountId}))
      .filter((item) => !EXCLUDED_ACCOUNTS[item.accountId]);
}

/**
 * Fetches all account users.
 *
 * @param {number} profileId - the profile ID associated with the API call
 * @param {DoubleClickCampaigns.AccountUserProfiles.ListParams=} params -
 *   optional params to be sent with the API call
 * @return {!Array.<!AccountUserProfileItem>} the list of
 *   account users
 */
function getAccountUserProfiles(profileId, params = undefined) {
  return fetchAll_(
      (params) => makePageData_(
          DoubleClickCampaigns.AccountUserProfiles.list(profileId, params),
          'accountUserProfiles'),
      params, (user) => ({
                id: user.id,
                active: user.active,
                email: user.email,
                name: user.name,
                accountId: user.accountId,
                subaccountId: user.subaccountId,
                advertiserFilter: user.advertiserFilter,
                userRoleId: user.userRoleId
              }));
}

/**
 * Fetches all accounts.
 *
 * @param {number} profileId - the profile ID associated with the API call
 * @param {DoubleClickCampaigns.Accounts.ListParams=} params - optional params
 *   to be sent with the API call
 * @return {!Array.<!MinimalItem>} the list of accounts
 */
function getAccounts(profileId, params = undefined) {
  return fetchAll_(
      (params) => makePageData_(
          DoubleClickCampaigns.Accounts.list(profileId, params), 'accounts'),
      params, MINIMAL_ITEM_TRANSFORM);
}

/**
 * Fetches all subaccounts.
 *
 * @param {number} profileId - the profile ID associated with the API call
 * @param {DoubleClickCampaigns.Subaccounts.ListParams=} params - optional
 *   params to be sent with the API call
 * @return {!Array.<!MinimalItem>} the list of subaccounts
 */
function getSubaccounts(profileId, params) {
  return fetchAll_(
      (params) => makePageData_(
          DoubleClickCampaigns.Subaccounts.list(profileId, params),
          'subaccounts'),
      params, MINIMAL_ITEM_TRANSFORM);
}

/**
 * Fetches all advertisers
 *
 * @param {number} profileId - the profile ID associated with the API call
 * @param {DoubleClickCampaigns.Advertisers.ListParams=} params - optional
 *   params to be sent with the API call
 * @return {!Array.<!MinimalItem>} the list of advertisers
 */
function getAdvertisers(profileId, params) {
  return fetchAll_(
      (params) => makePageData_(
          DoubleClickCampaigns.Advertisers.list(profileId, params),
          'advertisers'),
      params, MINIMAL_ITEM_TRANSFORM);
}

/**
 * Fetches all user roles.
 *
 * @param {number} profileId - the profile ID associated with the API call
 * @param {DoubleClickCampaigns.UserRoles.ListParams=} params - optional params
 *   to be sent with the API call
 * @return {!Array.<!DoubleClickCampaigns.UserRole>} the list of user roles.
 */
function getUserRoles(profileId, params) {
  return fetchAll_(
      (params) => makePageData_(
          DoubleClickCampaigns.UserRoles.list(profileId, params), 'userRoles'),
      params, (role) => ({
                id: role.id,
                name: role.name,
                permissions: role.permissions,
                defaultUserRole: role.defaultUserRole
              }));
}

/**
 * Fetches all user role permissions.
 *
 * @param {number} profileId - the profile ID associated with the API call
 * @param {!DoubleClickCampaigns.UserRolePermissions.ListParams=} params -
 *   optional params to be sent with the API call
 * @return {!Array.<!DoubleClickCampaigns.UserRolePermission>} the list of user
 *     roles.
 */
function getUserRolePermissions(profileId, params) {
  return fetchAll_(
      (params) => makePageData_(
          DoubleClickCampaigns.UserRolePermissions.list(profileId, params),
          'userRolePermissions'),
      params);
}

/**
 * Returns the first role in DCM which matches the role name and subaccount ID
 * (if provided).
 *
 * @param {string} roleName - the name of the role to search
 * @param {number} profileId - the profileId associated with the API call
 * @param {number=} subaccountId - an optional subaccount id within which to
 *   search
 *
 * @return {!DoubleClickCampaigns.UserRole|undefined} the first role which
 *     matches the role name and
 *   the subaccount ID or undefined if the role could not be found
 */
function findUserRoleByName(roleName, profileId, subaccountId = undefined) {
  let searchParams = {
    searchString: roleName,
    subaccountId: subaccountId,
    accountUserRoleOnly: !subaccountId
  };
  let roles = getUserRoles(profileId, searchParams);

  return roles.find((role) => role.name === roleName);
}
