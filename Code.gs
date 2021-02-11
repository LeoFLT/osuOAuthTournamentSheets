const SP = PropertiesService.getScriptProperties();
const SS = SpreadsheetApp.getActiveSpreadsheet();
/**
 * @type {{
 * mode: number,
 * tournamentName: string,
 * tournamentIcon: string,
 * forumPostURL: string,
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
const ROLES_TO_GIVE = SECRET.discordRoles ? SECRET.discordRoles.split(',') : '';
const TOURNAMENT_NAME = SECRET.tournamentName;
const TOURNAMENT_ACRONYM = SECRET.tournamentAcronym;
const TOURNAMENT_ICON = SECRET.tournamentIcon ? SECRET.tournamentIcon : '';
const FORUM_POST_URL = SECRET.forumPostURL ? SECRET.forumPostURL : 'https://osu.ppy.sh/home';
// The date after which new registrations will not be allowed
const REGISTRATION_END_DATE = SECRET.registrationEndDate ? new Date(SECRET.registrationEndDate) : '';
// working sheet, realistically the only thing you would change in this script
const SHEET = '_DATA';
// properties that are added to all URL objects
const URL_PROPERTIES = {
  tourName: TOURNAMENT_NAME,
  tourIcon: TOURNAMENT_ICON,
  tourAcronym: TOURNAMENT_ACRONYM,
  forumPostURL: FORUM_POST_URL
};
// Registration page's theming
const THEMING = {
  bg_color: '#36393f',
  nav_color_primary: '#7f04e1',
  nav_color_secondary: '#cc95f8',
  button_color: '#7289da',
  user_rank_color: '#eeeeee',
  title_text_color: '#ffffff',
  body_text_color: '#8e9094',
  text_hover_color: '#295c77',
  header_text_hover_color: '#dddddd',
  avatar_roundness: '50%',
  logo_roundness: '50%',
  // any font from this list https://fonts.google.com/
  font_family: 'Noto Sans'
};
/** Dictionary object for generating OAuth2 redirect URLs */
const GenerateURI = {
  /** @property {string} osu Generate an osu! OAuth2 URI */
  osu: `https://osu.ppy.sh/oauth/authorize?client_id=${SECRET.osuClientId}&redirect_uri=${SECRET.redirectUri}&response_type=code&scope=identify&state=%7B%22step%22%3A%22osu%22%7D`,
  /** @param {{id: number, username: string}} params The parameters to feed into the URL */
  discord: (params) => `https://discord.com/api/oauth2/authorize?client_id=${SECRET.discordClientId}&redirect_uri=${REDIRECT_URI}&response_type=code&scope=identify%20guilds.join&state=%7B%22step%22%3A%22discord%22%2C%22osu_id%22%3A%22${params.id}%22%2C%22osu_username%22%3A%22${params.username}%22%7D&prompt=none`
};
/**
 * Mimmicks importing local stylesheets e.g. ./css/index.css
 * @param {string} file The project file you're importing
 * @returns {HtmlService} HtmlService.HtmlOutput file, for appending as HTML code to another HtmlService.HtmlTemplate file.
 */
function include(file, ...overrides) {
  let page = HtmlService.createTemplateFromFile(file);
  for (obj of overrides) {
    Object.assign(page, obj);
  }
  return page
    .evaluate()
    .getContent();
}
TemplateService.include().append()
/**
 * Service class for managing page templates
 * @class
 */
class TemplateService {
  /**
   * Creates a HtmlService.HtmlTemplate file based on a project file
   * @param {string} file The project file to import
   */
  static createTemplateFromFile(file) {
    this.htmlFile = HtmlService.createTemplateFromFile(file);
    return this;
  }
  /** 
   * @method include Inserts a list of files inside the template as a header, similar to importing a stylesheet in HTML
   * @param {...string} fileList The list of strings that each represent a project file to be imported
   * @param {...{objList: any}} overrides Not implemented 
   * @returns {HtmlService.HtmlTemplate} Returns a HtmlTemplate object, for chaining
  */
  static include(...fileList) {
    let itemsToInclude = [];
    for (const file of fileList) {
      itemsToInclude = [...itemsToInclude, include(file, THEMING)];
    }
    this.htmlFile.stylesheets = itemsToInclude.join('\n');
    return this;
  }
  /** 
   * @method append Appends a list of objects as properties of the HtmlService.HtmlTemplate file
   * @param {...{objList: any}} objList The list of object(s) to append to the file as properties
   * @returns {HtmlService.HtmlTemplate} Returns a HtmlTemplate object, for chaining
   */
  static append(...objList) {
    for (const obj of objList) {
      Object.assign(this.htmlFile, obj);
    }
    return this;
  }
  /**
   * @method render Renders the HtmlTemplate object
   * @param {string} title The web page title
   * @returns {HtmlService.HtmlOutput} Returns a HtmlOutput object, to be served as webpage
   */
  static render(title) {
    if (title) return this.htmlFile.evaluate(title);
    return this.htmlFile.evaluate();
  }
}

// this is the code that gets executed when the REDIRECT_URI is called from a browser
function doGet(e) {
  // no state = nothing to do
  if (!e.parameter.state) {
    return TemplateService
      .createTemplateFromFile('Error')
      .include('index.css')
      .append(URL_PROPERTIES, {
        error_header: '404 Not Found',
        error_body: 'We couldn\'t find the page you were looking for',
        error_footer: null
      })
      .render(`${TOURNAMENT_ACRONYM} - Not Found`);
  }

  // abstract the state from the URL
  // error parameter in the url = user denied either of the oauth provider's consent screens
  const state = JSON.parse(e.parameter.state);

  const date = new Date().getTime();
  if (REGISTRATION_END_DATE ? (date > REGISTRATION_END_DATE.getTime()) : false) {
    return TemplateService
      .createTemplateFromFile('Registration-Over')
      .include('index.css')
      .append(URL_PROPERTIES, {
        endDate: REGISTRATION_END_DATE.toUTCString().replace('GMT', 'UTC')
      })
      .render(`${TOURNAMENT_ACRONYM} - Registration Period Over`);
  }

  if (state.step === 'osu') {
    if (e.parameter.hasOwnProperty('error')) {
      return TemplateService
        .createTemplateFromFile('Access-Denied')
        .include('index.css')
        .append(URL_PROPERTIES, { resource_denied: 'osu!' })
        .render(`${TOURNAMENT_ACRONYM} - Authorization Failed`);
    }
    // abstract the code from the URL
    const token = e.parameter.code;
    if (!token) {
      return TemplateService
        .createTemplateFromFile('Error')
        .include('index.css')
        .append(URL_PROPERTIES, {
          error_header: '400 Bad Request',
          error_body: 'Your request did not return an authentication code'
        })
        .render(`${TOURNAMENT_ACRONYM} - Bad Request`);
    }
    const authToken = getOsuToken(token);

    if (!authToken) {
      return TemplateService
        .createTemplateFromFile('Error')
        .include('index.css')
        .append(URL_PROPERTIES, {
          error_header: '400 Bad Request',
          error_body: 'Your authentication token is invalid or has expired'
        })
        .render(`${TOURNAMENT_ACRONYM} - Error`);
    }
    let user;
    try { user = queryUser(authToken); }
    catch (e) {

    }
    if (!user) {
      return TemplateService
        .createTemplateFromFile('Error')
        .include('index.css')
        .append(URL_PROPERTIES, {
          error_header: '400 Bad Request',
          error_body: 'Failed to query your osu! profile info, possibly because you attempted to do something you shouldn\'t'
        })
        .render(`${TOURNAMENT_ACRONYM} - Bad Request`);
    }

    if (user.hasOwnProperty('is_restricted')) {
      if (user.is_restricted === true) {
        return TemplateService
          .createTemplateFromFile('Error')
          .include('index.css')
          .append(URL_PROPERTIES, {
            error_header: '401 Unauthorized',
            error_body: 'Your osu! account is currently restricted. Restricted players may not interact in any multiplayer activities',
            error_footer: 'You may close the page'
          })
          .render(`${TOURNAMENT_ACRONYM} - Unauthorized`);
      }
    }

    // get a range delimited by the dimensions where there is data (equivalent to Ctrl + A)
    const range = SS.getRange(`${SHEET}!A1`).getDataRegion();
    // range = [[header_id, header_username, header_rank, header_badge_count, header_avatar_url]]
    const userIsPresent = range.getValues().some(r => r[1] === user.id);
    if (userIsPresent) {
      return TemplateService
        .createTemplateFromFile('Already-Registered')
        .include('index.css')
        .append(URL_PROPERTIES, {
          url: GenerateURI.discord(user),
          id: user.id,
          username: user.username,
          rank: user.rank
        })
        .render(`${TOURNAMENT_ACRONYM} - Player Already Registered`);
    }

    const addToRange = [
      new Date(),
      user.id,
      user.username,
      user.rank,
      user.pp,
      user.statistics.play_count,
      new Date(user.join_date),
      user.badgeCount,
      `https://a.ppy.sh/${user.id}`,
      user.country_code
    ];

    // append a row to the worksheet
    SS.getSheetByName(SHEET).appendRow(addToRange);

    return TemplateService
      .createTemplateFromFile('Registration-Success')
      .include('index.css')
      .append(URL_PROPERTIES, {
        url: GenerateURI.discord(user),
        id: user.id,
        username: user.username,
        rank: user.rank
      })
      .render(`${TOURNAMENT_ACRONYM} - Player Registered Successfully`);
  }

  if (state.step === 'discord') {
    if (e.parameter.hasOwnProperty('error')) {
      return TemplateService
        .createTemplateFromFile('Access-Denied')
        .include('index.css')
        .append(URL_PROPERTIES, { resource_denied: 'Discord' })
        .render(`${TOURNAMENT_ACRONYM} - Authorization Denied`);
    }
    // abstracting auth code from url
    const token = e.parameter.code;
    if (!token) {
      return TemplateService
        .createTemplateFromFile('Unauthorized')
        .include('index.css')
        .append(URL_PROPERTIES)
        .render(`${TOURNAMENT_ACRONYM} - Error`);
    }

    let authToken;
    try { authToken = getDiscordToken(token); }
    catch (e) {
      const error = [
        new Error(`authToken assertion failed for User ${state.osu_username} @ ${new Date().toUTCString()}, user_id: ${state.osu_id}`),
        e.stack
      ];
      console.log(...error);

      return TemplateService
        .createTemplateFromFile('Unauthorized')
        .include('index.css')
        .append(URL_PROPERTIES, { error: 'Error joining server/giving Role.' })
        .render(`${TOURNAMENT_ACRONYM} - Error`);
    }
    const uid = parseInt(state.osu_id);
    const username = state.osu_username

    let query;
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

    try { query = discordJoinServer(authToken, username); }
    catch (e) {
      let error = [new Error('discordJoinServer() threw an exception\nTimestamp: ' + new Date().toUTCString().replace('GMT', 'UTC')), e.stack];
      console.log(...error);

      return TemplateService
        .createTemplateFromFile('Error')
        .include('index.css')
        .append(URL_PROPERTIES, { error: 'Error while assigning Discord Role.' })
        .render(`${TOURNAMENT_ACRONYM} - Error`);
    }
    // 201: member succesffully joined the server
    if (query.response === 201) {
      // finding the row where the osu! userID is and associating the Discord Tag + Discord userID to it
      SS.getSheetByName(SHEET).getRange(insertRow, range.getLastColumn() - 2, 1, 3).setValues([[query.discordTag, query.discordId, false]]);
      return TemplateService
        .createTemplateFromFile('Discord20x')
        .include('index.css')
        .append(URL_PROPERTIES, {
          outcome: 'Server joined successfully',
          id: uid,
          username: username,
          discord_tag: query.discordTag
        })
        .render(`${TOURNAMENT_ACRONYM} - Server joined successfully`);
    }
    // 204: member already joined the server, roles added
    if (query.response === 204) {
      // finding the row where the osu! userID is and associating the Discord Tag + Discord userID to it
      SS.getSheetByName(SHEET).getRange(insertRow, range.getLastColumn() - 2, 1, 3).setValues([[query.discordTag, query.discordId, true]]);
      return TemplateService
        .createTemplateFromFile('Discord20x')
        .include('index.css')
        .append(URL_PROPERTIES, {
          outcome: 'Player Role assigned successfully',
          id: uid,
          username: username,
          discord_tag: query.discordTag
        })
        .render(`${TOURNAMENT_ACRONYM} - Player already in the server`);
    }
  }
  else {
    return TemplateService
      .createTemplateFromFile('Error')
      .include('index.css')
      .append(URL_PROPERTIES, {
        error_header: '500 Internal Server Error',
        error_body: 'The server encountered an internal error and could not complete your request',
        error_footer: 'You may close this window and try again at a later time'
      })
      .render(`${TOURNAMENT_ACRONYM} - Error`);
  }
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
  // no token, malformed token or invalid query
  if (fetchToken.getResponseCode() !== 200) throw new Error('Discord: 401 Unauthorized');

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

  if (fetchUser.getResponseCode() !== 200) throw new Error('osu!: 401 Unauthorized');

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
/**
 * @param {string} token The oauth2 user token.
 * @return {number} HTTPResponse code indicating the result of the operation.
 */
function discordJoinServer(token, nick) {
  try {
    const baseURL = 'https://discord.com/api/v8';
    let urlUsers = `${baseURL}/users/@me`
    const rolesToUrl = ROLES_TO_GIVE;

    const requestUser = UrlFetchApp.fetch(urlUsers, {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${token}`,
      },
    });
    const user = JSON.parse(requestUser);

    const guildUserFetch = ((roles, method, nickToGive) => {
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
        params.method = 'put';
        params.payload = JSON.stringify({
          "nick": nickToGive,
          "access_token": token,
          "roles": roles
        });
      }
      // member in guild
      if (method === 'patch') {
        urlGuilds += `/roles/${roles}`;
        // yes I called it 'patch' even though it actually 'put's the resource, fight me
        params.method = 'put';
        params.payload = JSON.stringify({
          "access_token": token
        });
      }

      const request = UrlFetchApp.fetch(urlGuilds, params);
      return { user: request.user, response: request.getResponseCode() };
    });

    const discordTag = `${user.username}#${user.discriminator}`;
    const discordId = user.id;
    let requestGuild = guildUserFetch();

    // user not in guild
    if (requestGuild.response === 404) {
      let result = guildUserFetch(rolesToUrl, 'put', nick);
      // @ts-ignore
      return { discordTag, discordId, response: result.response };
    }

    // user in guild, update instead
    if (requestGuild.response === 200) {
      let response;
      for (currentRole of rolesToUrl) {
        response = guildUserFetch(currentRole, 'patch').response;
      }
      // @ts-ignore
      return { discordTag, discordId, response: response };
    }
    // @ts-ignore      
    return { discordTag, discordId, response: requestGuild.response };
  }
  catch (e) {
    return new Error('discordJoinServer()');
  }
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
  const range = SS.getRangeByName(`${SHEET}!B2:J`);
  let data = range.getValues();
  for (row of data) {
    try {
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
        row[7] = `https://a.ppy.sh/${user.id}`;
        row[8] = user.country_code;
      }
    }
    catch (e) {
      console.log(new Error(`Querying user_id ${row[0]} failed.`, e.stack));
    }
  }
  const rangeToAdd = [1, 1, range.getLastRow(), range.getLastColumn()];
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