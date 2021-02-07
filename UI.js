function onOpen() {
  const UI = SpreadsheetApp.getUi();
  let createUI = UI.createMenu('Spreadsheet Settings')
    .addItem('Authorize Scripts', 'authorize')
    .addItem('Show Instructions', 'showInstructions')
    .addItem('Set Script Properties', 'setEnvVars')
    .addSeparator()
    .addSubMenu(
      UI.createMenu('Player Info Management')
        .addItem('Update Player Info', 'updateUsers')
        .addItem('Delete Player List', 'deleteUsersPrompt')
    )
    .addSeparator()
    .addSubMenu(
      UI.createMenu('Update Trigger Settings')
        .addItem('Add Triggers', 'setTriggers')
        .addItem('Remove Triggers', 'removeTriggers')
    )
  return createUI.addToUi();
}

function authorize() {
  const UI = SpreadsheetApp.getUi();
  return UI.alert('Already authorized');
}

function deleteUsersPrompt() {
  const UI = SpreadsheetApp.getUi();
  const alert = UI.alert('Delete player list', 'Are you sure you want to delete the entire player list?', UI.ButtonSet.YES_NO);
  if (alert === UI.Button.YES) {
    return deleteUsers();
  }
}

function showInstructions() {
  const UI = SpreadsheetApp.getUi();
  let html;
  if (REDIRECT_URI) html = HtmlService.createTemplateFromFile('Setup-UI-2');
  else html = HtmlService.createTemplateFromFile('Setup-UI-1');

  UI.showModalDialog(
    html.evaluate()
      .setWidth(1000)
      .setHeight(1000),
    `Sheet setup - ${REDIRECT_URI ? 'Part 2' : 'Part 1'}`);
}

function setEnvVars() {
  const UI = SpreadsheetApp.getUi();
  const prompt = (title, message) => UI.prompt(title, message, UI.ButtonSet.OK_CANCEL);

  const redirectUri = UI.prompt('Enter your project\'s Redirect Uri', 'Get it by deploying the Apps Script Project as a web app\n\nCancel: no change', UI.ButtonSet.OK_CANCEL);
  if (redirectUri.getSelectedButton() === UI.Button.OK) {
    const result = redirectUri.getResponseText().trim();
    return PropertiesService.getScriptProperties().setProperty('redirectUri', result);
  }

  const tournamentAcronym = prompt('Enter your Tournament\'s acronym (e.g. My osu! Tournament => MOT)', `Current acronym: ${SECRET.acronym ? SECRET.acronym : 'No acronym set'}\n\nCancel: no change`);
  if (tournamentAcronym.getSelectedButton() === UI.Button.OK) {
    const result = tournamentAcronym.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('tournamentAcronym', result);
  }

  const tournamentMode = prompt('Enter the tournament game mode', `1: Standard\n2: Mania\n3: Taiko\n4: Catch The Beat\n\nCurrent mode: ${SECRET.mode ? SECRET.mode : 'No mode set'}\n\nCancel: no change`);
  if (tournamentMode.getSelectedButton() === UI.Button.OK) {
    const result = parseInt(tournamentMode.getResponseText().trim(), 10);
    let finalResult;
    switch (result) {
      case 2: finalResult = 'mania'; break;
      case 3: finalResult = 'taiko'; break;
      case 4: finalResult = 'fruits'; break;
      case 1: default: finalResult = 'osu';
    }
    PropertiesService.getScriptProperties().setProperty('mode', finalResult);
  }

  const registrationEndDate = prompt('Enter your registration deadline', `Current end date: ${REGISTRATION_END_DATE ? REGISTRATION_END_DATE : 'No end date set (signups open forever)'}\n\nFormat: ${new Date().toUTCString()}\n\nTimezone codes are supported (https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Date#several_ways_to_create_a_date_object)\n\nLeave the field empty to have signups always open\n\nCancel: no change`,)
  if (registrationEndDate.getSelectedButton() === UI.Button.OK) {
    const result = registrationEndDate.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('registrationEndDate', result);
  }

  const osuClientIdPrompt = prompt('Enter your osu! OAuth Client ID', `Current osu! OAuth Client ID: ${SECRET.osuClientId ? SECRET.osuClientId : 'No osu! Client ID set'}\n\nCancel: no change`);
  if (osuClientIdPrompt.getSelectedButton() === UI.Button.OK) {
    const result = osuClientIdPrompt.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('osuClientId', result);
  }

  const osuClientSecretPrompt = prompt('Enter your osu! OAuth Client Secret', `Current osu! OAuth Client Secret: ${SECRET.osuClientSecret ? SECRET.osuClientSecret : 'No osu! Client Secret set'}\n\nCancel: no change`);
  if (osuClientSecretPrompt.getSelectedButton() === UI.Button.OK) {
    const result = osuClientSecretPrompt.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('osuClientSecret', result);
  }

  const discordClientIdPrompt = prompt('Enter your Discord OAuth Client ID', `Current Discord OAuth Client ID: ${SECRET.discordClientId ? SECRET.discordClientId : 'No Discord Client ID set'}\n\nCancel: no change`);
  if (discordClientIdPrompt.getSelectedButton() === UI.Button.OK) {
    const result = discordClientIdPrompt.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('discordClientId', result);
  }

  const discordClientSecretPrompt = prompt('Enter your Discord OAuth Client Secret', `Current Discord OAuth Client Secret: ${SECRET.discordClientSecret ? SECRET.discordClientSecret : 'No Discord Client Secret set'}\n\nCancel: no change`);
  if (discordClientSecretPrompt.getSelectedButton() === UI.Button.OK) {
    const result = discordClientSecretPrompt.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('discordClientSecret', result);
  }

  const discordBotToken = prompt('Enter your Discord Bot Token', `Current Discord Bot Token: ${SECRET.discordBotToken ? SECRET.discordBotToken : 'No Discord bot token set'}\n\nCancel: no change`);
  if (discordBotToken.getSelectedButton() === UI.Button.OK) {
    const result = discordBotToken.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('discordBotToken', result);
  }

  const discordGuildId = prompt('Enter your Discord Guild ID', `Current Discord Guild ID: ${SECRET.discordGuildId ? SECRET.discordGuildId : 'No Discord Guild ID set'}\n\nCancel: no change`);
  if (discordGuildId.getSelectedButton() === UI.Button.OK) {
    const result = discordGuildId.getResponseText().trim();
    PropertiesService.getScriptProperties().setProperty('result', result);
  }

  const discordRoles = prompt('Enter your Discord Player Roles', `Current Discord Player Role(s): ${SECRET.discordRoles ? SECRET.discordRoles : 'No Discord Player Roles set'}\n\n(separate multiple Roles with commas e.g. 123456789,987654321)\n\nCancel: no change`);
  if (discordRoles.getSelectedButton() === UI.Button.OK) {
    const result = discordRoles.getResponseText().trim();
    let finalResult = result.trim().replace(/\s/g, '');
    PropertiesService.getScriptProperties().setProperty('discordRoles', finalResult);
  }
}

function setTriggers() {
  const UI = SpreadsheetApp.getUi();
  const SP = PropertiesService.getScriptProperties();
  if (SP.getProperty('playerUpdateTriggerId')) return UI.alert('Update Trigger already added.');

  const alert = (title, message) => UI.alert(title, message, UI.ButtonSet.YES_NO);
  const setTriggerAlert = alert('Create Triggers?', 'This will create a trigger for updating player Usernames/Ranks/Badge Counts once a day.\n\nDo You want to continue?');

  if (setTriggerAlert === UI.Button.YES) {
    const trigger = ScriptApp.newTrigger('updateUsers')
      .timeBased()
      .atHour(1)
      .everyDays(1)
      .create()
      .getUniqueId();
    SP.setProperty('playerUpdateTriggerId', trigger);
    return UI.alert('Trigger created successfully.');
  }
}

function removeTriggers() {
  const UI = SpreadsheetApp.getUi();
  const triggerId = SP.getProperty('playerUpdateTriggerId');
  if (!triggerId) return UI.alert('No triggers to delete.');

  const alert = (title, message) => UI.alert(title, message, UI.ButtonSet.YES_NO);
  const removeTriggerAlert = alert('Remove Triggers?', 'Warning: this will remove all the triggers for updating player information.\n\nDo you want to continue?');

  if (removeTriggerAlert === UI.Button.YES) {
    const allTriggers = ScriptApp.getProjectTriggers();
    SP.deleteProperty('playerUpdateTriggerId');

    for (const trigger of allTriggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        SP.deleteProperty('playerUpdateTriggerId');
        return UI.alert('Trigger deleted successfully.');
      }
    }

    return UI.alert('No triggers to delete.');
  }
}