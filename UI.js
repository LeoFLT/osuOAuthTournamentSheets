function onOpen() {
  const UI = SpreadsheetApp.getUi();
  UI.createMenu('Spreadsheet Settings')
    .addItem('Authorize Scripts', 'authorize')
    .addItem('Show Instructions', 'forumPostURL')
    .addItem('Set Script Properties', 'setEnvVars')
    .addToUi();
}

function authorize() {
  const UI =  SpreadsheetApp.getUi();
  return UI.alert('Already authorized');
}

function forumPostURL() {
  const UI = SpreadsheetApp.getUi();
  let html = HtmlService.createTemplateFromFile('Setup-UI');
  UI.showModalDialog(
    html.evaluate()
      .setWidth(1000)
      .setHeight(1000), 'Sheet setup')
}

function setEnvVars() {
  const UI = SpreadsheetApp.getUi();
  const prompt = (title, message) => UI.prompt(title, message, UI.ButtonSet.OK_CANCEL);
  const tournamentAcronym = prompt('Enter your Tournament\'s acronym (e.g. My osu! Tournament => MOT)', 'Cancel: no change'); 
  const osuClientIdPrompt = prompt('Enter your osu! OAuth Client ID', 'Cancel: no change');
  const osuClientSecretPrompt = prompt('Enter your osu! OAuth Client Secret', 'Cancel: no change');
  const discordClientIdPrompt = prompt('Enter your Discord OAuth Client ID', 'Cancel: no change');
  const discordClientSecretPrompt = prompt('Enter your Discord OAuth Client Secret', 'Cancel: no change');
  const discordBotToken = prompt('Enter your Discord Bot Token', 'Cancel: no change');
  const discordGuildId = prompt('Enter your Discord Guild ID', 'Cancel: no change');
  const discordRoles = prompt('Enter your Discord player roles', '(separate with commas e.g. 123456789,987654321)\nCancel: no change');

  let propertiesToAdd = {};
  if (tournamentAcronym.getSelectedButton() === UI.Button.OK) {
    const result = tournamentAcronym.getResponseText().trim();
    propertiesToAdd.tournamentAcronym = result;
  }
  if (osuClientIdPrompt.getSelectedButton() === UI.Button.OK) {
    const result = osuClientIdPrompt.getResponseText().trim();
    propertiesToAdd.osuClientId = result;
  }
  if (osuClientSecretPrompt.getSelectedButton() === UI.Button.OK) {
    const result = osuClientSecretPrompt.getResponseText().trim();
    propertiesToAdd.osuClientSecret = result;
  }
  if (discordClientIdPrompt.getSelectedButton() === UI.Button.OK) {
    const result = discordClientIdPrompt.getResponseText().trim();
    propertiesToAdd.discordClientId = result;
  }
  if (discordClientSecretPrompt.getSelectedButton() === UI.Button.OK) {
    const result = discordClientSecretPrompt.getResponseText().trim();
    propertiesToAdd.discordClientSecret = result;
  }
  if (discordBotToken.getSelectedButton() === UI.Button.OK) {
    const result = discordBotToken.getResponseText().trim();
    propertiesToAdd.discordBotToken = result;
  }
  if (discordGuildId.getSelectedButton() === UI.Button.OK) {
    const result = discordGuildId.getResponseText().trim();
    propertiesToAdd.discordGuildId = result;
  }
  if (discordRoles.getSelectedButton() === UI.Button.OK) {
    const result = discordRoles.getResponseText().trim();
    let finalResult = result.trim().replace(/\s/g,'');
    propertiesToAdd.discordRoles = finalResult;
  }
  PropertiesService.getScriptProperties().setProperties(propertiesToAdd);
}
