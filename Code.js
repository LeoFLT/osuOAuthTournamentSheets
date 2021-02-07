const SP = PropertiesService.getScriptProperties();
const SS = SpreadsheetApp.getActiveSpreadsheet();
/**
 * @type {{
 * mode: number,
 * tournamentAcronym: string,
 * redirectUri: string,
 * registrationEndDate: string,
 * osuClientId: string, 
 * osuClientSecret: string,
 * discordClientId: string,
 * discordClientSecret: string,
 * discordBotToken: string,
 * discordGuildId: 'snowflake',
 * discordRoles: 'snowflake[]' 
 * }}
 */
const SECRET = SP.getProperties();
const CACHE = CacheService.getScriptCache();
const LOCK = LockService.getScriptLock();
// this the url that goes on the osu! account settings 'Application Callback URL'
const REDIRECT_URI = SECRET.redirectUri;
// Discord Guild ID (snowflake=string) to add players to
const GUILD = SECRET.discordGuildId;
// query mode for players (standard/mania/taiko/ctb)
const MODE = SECRET.mode;
// Array of Discord Role IDs (snowflakes=strings) to add to players
// stored as a string in the format '0123456789,1012131415' and split afterwards;
const ROLES_TO_GIVE = SECRET.discordRoles;
const TOURNEY_PREFIX = SECRET.tournamentAcronym;
// The date after which new registrations will not be allowed
const REGISTRATION_END_DATE = SECRET.registrationEndDate;
// working sheet, realistically the only thing you would change in this script
const SHEET = '_DATA';

// endDate.toUTCString().replace('GMT', 'UTC')
// this is the code that gets executed when the REDIRECT_URI is called from a browser
function doGet(e) {
  const date = new Date().getTime();
  if (date && (date > REGISTRATION_END_DATE.getTime())) {
    let page = HtmlService
      .createTemplateFromFile('Registration-Over');
    page.endDate = REGISTRATION_END_DATE.toUTCString().replace('GMT', 'UTC');
    page.forumPostURL = FORUM_POST;

    return page
      .evaluate()
      .setTitle(`${TOURNEY_PREFIX} - Registration Period Over`);
  }
  // abstract the state from the URL
  const state = e.parameter.state;
  // error parameter in the url = user denied either of the oauth provider's consent screens
  if (e.parameter.hasOwnProperty('error')) {
    return HtmlService.createTemplateFromFile('Access-denied')
      .evaluate()
      .setTitle(`${TOURNEY_PREFIX} - Authorization Failed`);
  }
  // no state = nothing to do
  if (!state) {
    return HtmlService.createTemplateFromFile('Unauthorized')
      .evaluate()
      .setTitle(`${TOURNEY_PREFIX} - Unauthorized`);
  }
  if (state === 'osu') {
    // abstract the code from the URL
    const token = e.parameter.code;
    if (!token) return HtmlService.createHtmlOutputFromFile('Unauthorized');
    const authToken = getOsuToken(token);
    if (!authToken) return HtmlService.createHtmlOutputFromFile('Unauthorized');
    const user = queryUser(authToken);
    if (!user) {
      return HtmlService.createTemplateFromFile('Error')
        .evaluate()
        .setTitle(`${TOURNEY_PREFIX} - Error`);
    }
    // get a range delimited by the dimensions where there is data (equivalent to Ctrl + A)
    const range = SS.getRange(`${SHEET}!A1`).getDataRegion();
    // range = [[header_id, header_username, header_rank, header_badge_count, header_avatar_url]]
    const userIsPresent = range.getValues().some(r => r[1] === user.id);
    if (userIsPresent) {
      let page = HtmlService.createTemplateFromFile('Already-registered')
      page.url = `https://discord.com/api/oauth2/authorize?client_id=${SECRET.discordClientId}&redirect_uri=${REDIRECT_URI}&response_type=code&scope=identify%20guilds.join&state=discord.${user.id}`;
      return page
        .evaluate()
        .setTitle(`${TOURNEY_PREFIX} - Player Already Registered`);
    };

    // appending one row to the end of the range
    const addToRange = [
      [
        new Date(),
        user.id,
        user.username,
        user.rank,
        user.pp,
        user.statistics.play_count,
        new Date(user.join_date),
        user.badgeCount,
        user.avatar_url,
        user.country_code
      ]
    ];
    // start at the row directly after the last, first column and span 1 row, addToRange[0] columns
    SS.getSheetByName(SHEET).getRange(range.getLastRow() + 1, 1, 1, addToRange[0].length).setValues(addToRange);
    let page = HtmlService
      .createTemplateFromFile('Success')
    page.url = `https://discord.com/api/oauth2/authorize?client_id=${SECRET.discordClientId}&redirect_uri=${REDIRECT_URI}&response_type=code&scope=identify%20guilds.join&state=discord.${user.id}`;

    return page
      .evaluate()
      .setTitle(`${TOURNEY_PREFIX} - Player Registration`);
  }

  if (state.includes('discord')) {
    // abstracting auth code from url
    const token = e.parameter.code;
    const authToken = getDiscordToken(token);
    if (!authToken) {
      return HtmlService.createTemplateFromFile('Unauthorized')
        .evaluate()
        .setTitle(`${TOURNEY_PREFIX} - Error`)
    };
    const regexp = /^(?:discord.)(\d+)$/ig;
    const uid = parseInt(regexp.exec(e.parameter.state)[1], 10);
    let query = discordJoinServer(authToken);

    const range = SS.getRange(`${SHEET}!A1`).getDataRegion();
    let data = range.getValues();
    let i = 1;
    let insertRow;
    for (row of data) {
      if (row[1] === uid) {
        insertRow = i;
        break;
      } i++;
    }
    // finding the row where the osu! userID is and associating the Discord Tag + Discord userID to it 
    SS.getSheetByName(SHEET).getRange(insertRow, range.getLastColumn() - 1, 1, 2).setValues([[query.discordTag, query.discordId]]);
    // 201: member succesffully joined the server
    if (query.response === 201) {
      return HtmlService.createTemplateFromFile('Discord201')
        .evaluate()
        .setTitle(`${TOURNEY_PREFIX} - Server joined successfully`);
    };
    // 204: member already joined the server, roles added
    if (query.response === 204) {
      return HtmlService.createTemplateFromFile('Discord204')
        .evaluate()
        .setTitle(`${TOURNEY_PREFIX} - Player Already Registered`);
    }
  }
  else {
    return HtmlService.createTemplateFromFile('Unauthorized')
      .evaluate()
      .setTitle(`${TOURNEY_PREFIX} - Error`)
  };
}

// URL to be used on the forum post (maybe shorten it?)
function returnForumURL(returnUriOnly) {
  const redirectUri = REDIRECT_URI;
  const result = `https://osu.ppy.sh/oauth/authorize?client_id=${SECRET.osuClientId}&redirect_uri=${redirectUri}&response_type=code&scope=identify&state=osu`;
  if (returnUriOnly) return redirectUri;
  return result;
}
// osu AUTHORIZATION CODE GRANT (add new users) function
const getOsuToken = ((authCode) => {
  const url = 'https://osu.ppy.sh/oauth/token';
  const fetchToken = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      "grant_type": "authorization_code",
      "client_id": SECRET.osuClientId,
      "client_secret": SECRET.osuClientSecret,
      "redirect_uri": REDIRECT_URI,
      "code": authCode
    }),
    muteHttpExceptions: true
  })
  // only return the access_token, we're not storing refresh_tokens
  if (fetchToken.getResponseCode() !== 200) return null;
  const result = JSON.parse(fetchToken).access_token;
  return result;
});

// osu! CLIENT CREDENTIALS GRANT (update existing users) service
const getOsuClientService = (() => {
  return OAuth2.createService('osu! Client Credentials')
    .setAuthorizationBaseUrl('https://osu.ppy.sh/oauth/authorize')
    .setTokenUrl('https://osu.ppy.sh/oauth/token')
    .setClientId(SECRET.osuClientId)
    .setClientSecret(SECRET.osuClientSecret)

    // storing final token on the script env
    .setPropertyStore(PropertiesService.getScriptProperties())
    // cache for the key:value pairs
    .setCache(CacheService.getScriptCache())
    // setting a lock to prevent a race condition
    .setLock(LockService.getScriptLock())
    // client code grant implementation for osu! requires this to be public
    .setScope('public')
    .setGrantType('client_credentials')
});

// Discord AUTHORIZATION CODE GRANT (add new users to guild) function
const getDiscordToken = ((authCode) => {
  const url = 'https://discord.com/api/oauth2/token';
  const fetchToken = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: {
      "client_id": SECRET.discordClientId,
      "client_secret": SECRET.discordClientSecret,
      "grant_type": "authorization_code",
      "code": authCode,
      "redirect_uri": REDIRECT_URI,
      "scope": "identify guilds.join"
    },
    muteHttpExceptions: true
  })
  // no token
  if (fetchToken.getResponseCode() !== 200) return null;

  // only return the access_token, we're not storing refresh_tokens
  const result = JSON.parse(fetchToken).access_token;
  return result;
});

/**
 * Query the user whose token is passed as an argument
 * @param {string} token The token.
 * @return { osuUser } Object representing a user for a valid token, null otherwise.
 */
function queryUser(token) {
  const url = `https://osu.ppy.sh/api/v2/me/${MODE}`;
  const fetchUser = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json',
      'Content-Type': 'application/json',
    }
  });

  if (fetchUser.getResponseCode() !== 200) return null;

  let result = JSON.parse(fetchUser);
  result.badgeCount = 0;

  let filterRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('_filtered_badges!A1')
    .getDataRegion(SpreadsheetApp.Dimension.ROWS)
    .getValues()
    .flat(2)
    .filter(i => i);

  let expression = filterRange.slice(1).join('|');
  // crappy way to ignore badges based on a regexp but it works
  const ignoredBadges = new RegExp(expression, 'i');

  for (badge in result.badges) {
    let currentBadge = result.badges[badge].description;
    if (!ignoredBadges.test(currentBadge.toLowerCase())) result.badgeCount++;
  }
  // shorthands for common stats
  result.rank = result.statistics.pp_rank;
  result.pp = result.statistics.pp;

  // reduce potential cache size overflow
  result.page = null;
  result.monthly_playcounts = null;
  result.replays_watched_counts = null;
  result.user_achievements = null;

  return result;
}
/**
 * @param {string} token The oauth2 user token.
 * @return {number} HTTPResponse code indicating the result of the operation.
 */
function discordJoinServer(token) {
  const baseURL = 'https://discord.com/api/v8';
  const urlUsers = `${baseURL}/users/@me`
  const rolesToGive = ROLES_TO_GIVE.split(',');
  const requestUser = UrlFetchApp.fetch(urlUsers, {
    method: 'get',
    headers: {
      'Authorization': `Bearer ${token}`,
    }
  })
  const user = JSON.parse(requestUser);

  const guildUserFetch = ((rolesToGive, method) => {
    let urlGuilds = `https://discordapp.com/api/v8/guilds/${GUILD}/members/${user.id}`;
    let params = {
      method: 'get',
      headers: {
        'Authorization': `Bot ${SECRET.discordBotToken}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    // member not in guild
    if (method === 'put') {
      const rolesArr = rolesToGive;
      params.method = 'put';
      params.payload = JSON.stringify({
        "access_token": token,
        "roles": rolesArr
      });
      const request = UrlFetchApp.fetch(urlGuilds, params);
      return { user: request.user, response: request.getResponseCode() };
    }
    // member in guild
    if (method === 'patch') {
      urlGuilds += `/roles/${rolesToGive}`;
      const role = rolesToGive;
      // yes I called it 'patch' even though it actually 'put's the resource, fight me
      params.method = 'put';
      params.payload = JSON.stringify({
        "access_token": token,
        "roles": role
      });
      const request = UrlFetchApp.fetch(urlGuilds, params);
      return { user: request.user, response: request.getResponseCode() };
    }

    const request = UrlFetchApp.fetch(urlGuilds, params);
    return { user: request.user, response: request.getResponseCode() };
  });

  const discordTag = `${user.username}#${user.discriminator}`;
  const discordId = user.id;
  let requestGuild = guildUserFetch();

  // user not in guild
  if (requestGuild.response === 404) {
    let result = guildUserFetch(rolesToGive, 'put');
    // @ts-ignore
    return { discordTag, discordId, response: result.response };
  }

  // user in guild, update instead
  if (requestGuild.response === 200) {
    let response;
    for (currentRole of rolesToGive) {
      response = guildUserFetch(currentRole, 'patch').response;
    }
    // @ts-ignore
    return { discordTag, discordId, response: response };
  }
  // @ts-ignore      
  return { discordTag, discordId, response: requestGuild.response };
}
/** 
 * @param {(number|string)} userId The user ID to query, can be the username (auto detects which of the two it is)
 * @param {(number|string)} mode The mode to query for: 1 - standard; 2 - mania; 3 - taiko; 4 - catch
 * @typedef {Object} osuUser An object representing the osu! user for the queried user id
 * @property {string} username The user's current username
 * @property {number} rank The user's rank for the mode
 * @property {number} badgeCount The amount of tournament badges for a player (excludes Contributor/mapping/etc.)
 * @property {URL} avatar_url The URL for the user's avatar
 * @returns {osuUser} Object representing the osu! user. https://osu.ppy.sh/docs/index.html?javascript#user
*/
function getUser(userId, mode) {
  const osuService = getOsuClientService();
  userId = parseInt(userId, 10);
  /*
  enum | Mode
  ------------------
  1 |	osu!standard
  2 |	osu!mania
  3 |	osu!taiko
  4 |	osu!catch
  */
  switch (parseInt(mode, 10)) {
    case 2:
      gameMode = 'mania';
      break;
    case 3:
      gameMode = 'taiko';
      break;
    case 4:
      gameMode = 'fruits';
      break;
    case 1: default: gameMode = 'osu';
  }
  // no authorization
  if (!osuService.hasAccess()) throw new Error('Missing osu! authorization.');

  const url = `https://osu.ppy.sh/api/v2/users/${userId}/${gameMode}`;
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + osuService.getAccessToken(),
      Accept: 'application/json',
      'Content-Type': 'application/json',
    }
  });

  let result = JSON.parse(response);
  result.badgeCount = 0;

  let filterRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('_filtered_badges!A1')
    .getDataRegion(SpreadsheetApp.Dimension.ROWS)
    .getValues()
    .flat(2)
    .filter(i => i);

  let expression = filterRange.slice(1).join('|');
  // crappy way to ignore badges based on a regexp but it works
  const ignoredBadges = new RegExp(expression, 'i');
  console.log(expression, ignoredBadges);
  for (badge in result.badges) {
    let currentBadge = result.badges[badge].description;
    // if the badge's description (lowercased) doesn't match our regExp
    // add a badge to the user's badgeCount property
    if (!ignoredBadges.test(currentBadge.toLowerCase())) result.badgeCount++;
  }

  // shorthands for common stats
  result.rank = result.statistics.pp_rank;
  result.pp = result.statistics.pp;
  result.mode = gameMode;

  // reduce potential cache size overflow
  result.page = null;
  result.monthly_playcounts = null;
  result.replays_watched_counts = null;
  result.user_achievements = null;

  return result;
}


/**
 * Updates the osu! user information
 */
function updateUsers() {
  const range = SS.getRangeByName(`${SHEET}!B1:J`);
  let data = range.getValues();
  for (row of data) {
    // header row, skip this iteration
    if (row[0] === 'id') continue;
    // empty row, skip this iteration
    if (!row[0]) continue;

    const user = getUser(row[0], MODE);
    if (user.username === 'RESTRICTED') {
      row[1] += ' [RESTRICTED]';
      continue;
    }
    else if (user.hasOwnProperty(username)) {
      row[1] = user.username;
      row[2] = user.rank;
      row[3] = user.pp;
      row[4] = user.statistics.play_count;
      row[5] = new Date(user.join_date);
      row[6] = user.badgeCount;
      row[7] = user.avatar_url;
      row[8] = user.country_code;
    }
  }
  const rangeToAdd = [1, 2, range.getLastRow(), range.getLastColumn() - 1];
  // pushing the same range we queried, prevents race condition-related errors
  return SS.getSheetByName(SHEET).getRange(...rangeToAdd).setValues(data);
}
/**
 * Deletes all registered users from the sheet
 */
function deleteUsers() {
  const range = SS.getRangeByName(`${SHEET}!A1`).getDataRegion();
  const rangeToDelete = [2, 1, range.getLastRow(), range.getLastColumn()];
  return SS.getSheetByName(`${SHEET}`).getRange(...rangeToDelete).clearContent();
}

function bumpSheetVersion(bumpType) {
  const rangeVersion = SS.getRange('Instructions!F49');
  const rangeDate = SS.getRange('Instructions!G49');
  const version = rangeVersion.getValue();
  let newVersion;
  let bump;
  let regExp;

  switch (bumpType) {
    case 'patch':
      newVersion = version.slice(0, -1);
      bump = parseInt(version.slice(-1));
      newVersion += ++bump;
      rangeVersion.setValue(newVersion);
      rangeDate.setValue(('|  ' + new Date().toUTCString()).replace('GMT', 'UTC'));
      break;
    case 'minor':
      regExp = /(?:\.)(\d+)(?:\.\d+)/;
      bump = parseInt(version.match(regExp)[1]);
      newVersion = version.replace(regExp, `.${++bump}.0`);
      rangeVersion.setValue(newVersion);
      rangeDate.setValue(('|  ' + new Date().toUTCString()).replace('GMT', 'UTC'));
      break;
    case 'major':
      regExp = /(\d+)(?:\.\d+\.\d+)/;
      bump = parseInt(version.match(regExp)[1]);
      newVersion = version.replace(regExp, `${++bump}.0.0`);
      rangeVersion.setValue(newVersion);
      rangeDate.setValue(('|  ' + new Date().toUTCString()).replace('GMT', 'UTC'));
      break;
  }
}

function projectVersioning() {
  const UI = SpreadsheetApp.getUi();
  UI.createMenu('Project versioning')
    .addItem('Patch', 'bumpPatch')
    .addItem('Minor', 'bumpMinor')
    .addItem('Major', 'bumpMajor')
    .addToUi();
}

function bumpPatch() { return bumpSheetVersion('patch') }
function bumpMinor() { return bumpSheetVersion('minor') }
function bumpMajor() { return bumpSheetVersion('major') }