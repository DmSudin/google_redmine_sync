class application {

  constructor() {

    this.reset();
    this.init();
  }

  reset() {

    this.currentChange = {
      'projectName': undefined,
      properties: {},
    };

    this.range = undefined;
    this.oldValue = undefined;
    this.newValue = undefined;
  }

  init() {

    this.trackedColumns = {
      'status': { 'titleTable': 'Статус', 'titleRedmine': 'Статус', 'index': 15 },
      'pm': { 'titleTable': 'ПМ отв-й', 'titleRedmine': 'Ответственный PM', 'index': 8 },
      'unitLead': {'titleTable': 'ЮнитЛид', 'titleRedmine': 'Ответственный Unit Lead', 'index': 10 }
    };

    this.tableTitlesRowIndex = 2;

    this.sheetProjects = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ответственные и проекты");
    this.sheetSettings = SpreadsheetApp.getActive().getSheetByName('Redmine_sync');
    this.projectNameColumnIndex = 2;

    this.notifyDuration = 10;
    this.cellNotify = this.sheetProjects.getRange(1, 3).getCell(1, 1);

    this.pmTitleTable = this.sheetProjects.getRange(this.tableTitlesRowIndex, this.trackedColumns.pm.index).getValue();
    this.pmTitleRedmine = 'Ответственный PM';
    this.statusTitleTable = this.sheetProjects.getRange(this.tableTitlesRowIndex, this.trackedColumns.status.index).getValue();
    this.statusTitleRedmine = 'Статус';
    this.unitLeadTitleTable = this.sheetProjects.getRange(this.tableTitlesRowIndex, this.trackedColumns.unitLead.index).getValue();
    this.unitLeadTitleRedmine = 'Ответственный Unit Lead';

    this.redmineKey = 'e2306b943c5e70ff7ba20b8bcfa95b289d78e103';
    this.trackedProjects = this.getTrackedProjects();
  }

  getTrackedProjects() {
    const result = [];
    const redmineColumnIndex = 4;
    const projectColumnIndex = 5;
    const lastSheetRow = this.sheetSettings.getLastRow();

    const range = this.sheetSettings.getRange(`D2:D${lastSheetRow}`);
    //possible bug - empty cell among range in service_info - getNextDataCell()
    const lastProjectsRowIndex = range.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();

    for (let i = 2; i<= lastProjectsRowIndex; i++) {
      const elem = new Object();
      elem.redmineAlias = this.sheetSettings.getRange(i, redmineColumnIndex).getValue();
      elem.projectName = this.sheetSettings.getRange(i, projectColumnIndex).getValue();
      result.push(elem);
    }

    return result;
  }

  isTrackedProjectChange(changedRange) {
    const projectName = this.getProjectName(changedRange);
    if (this.trackedProjects.filter(item => item.projectName === projectName).length > 0) {
      this.currentChange.projectName = projectName;
      return true;
    }
    return false;
  }

  getProjectName() {
    return this.sheetProjects.getRange(this.range.getRow(), this.projectNameColumnIndex).getValue();
  }

  getProjectRedmineAlias(projectName) {
    return this.trackedProjects.filter(item => item.projectName === projectName)[0].redmineAlias;
  }

  async handleChange() {

    if ( this.isManyCellsChanged() ) {
      return;
    }

    this.currentChange.projectName = this.getProjectName();
    this.loadProjectDataFromTable();
    this.publishToRedmine(this.currentChange);
  }

  loadProjectDataFromTable() {
      Object.keys(this.trackedColumns).forEach(key => {
        const val = this.sheetProjects.getRange(this.range.getRowIndex(), this.trackedColumns[key].index).getValue();
        this.currentChange.properties[key] = val;
      });
  }

  async publishToRedmine(change) {
    const redmineAlias = this.getProjectRedmineAlias(change.projectName);
    const url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/Shared_Info.json?key=e2306b943c5e70ff7ba20b8bcfa95b289d78e103`;
    let textContent = '';
    const props = change.properties;

    Object.keys(props).forEach(key => {
      const propName = this.trackedColumns[key].titleRedmine;
      let propValue = '';

      if ( ((propName) === this.trackedColumns.pm.titleRedmine) || ((propName) === this.trackedColumns.unitLead.titleRedmine) ) {
        // propValue = `\"${props[key]}\":https://egamings.slack.com/team/${this.getSlackLink(props[key])}`;
        propValue = this.getSlackLink(props[key]);
      } else propValue = props[key];

      textContent += `*${propName}*: ${propValue}\r\n`;
      // add slack to the name prop
    });

    const data = {
      "wiki_page":
      {
        "text": `${textContent}`,
      },
    };

    const options = {
      method: 'put',
      'contentType': 'application/json',
      'payload': JSON.stringify(data),
    };

    const response = await UrlFetchApp.fetch(url, options);

    if (response.getResponseCode() === 204) { // Rest_WikiPages API
      const projectName = this.getProjectName();
      // Browser.msgBox(`Изменения по проекту ${projectName} внесены в Redmine Wiki`);
      this.showNotify(redmineAlias, projectName);
    } else Browser.msgBox(`Что-то пошло не так при внесении изменений в Redmine Wiki`);

    this.reset();

  }

  isTrackedFieldsChanged() {
    return this.range.getColumn() === this.statusColumnIndex || this.range.getColumn() === this.pmColumnIndex;
  }

  isManyCellsChanged() {
      return ((this.range.getNumColumns() !==1) || (this.range.getNumRows() !== 1));
  }

  isColumnTracked(columnNumber) {
    let result = false;
    Object.entries(this.trackedColumns).map((elem) => elem[1].index).forEach(item => {
      if (Math.trunc(columnNumber) === Math.trunc(item)) {
        result = true;
      }
    });

    return result;
  }

  isProjectsSheetEdited() {
    return this.source.getActiveSheet().getName() === this.sheetProjects.getName();
  }

  async getProjectData(projectName) { // unused, but just for case, for the future
    const redmineAlias = this.getProjectRedmineAlias(projectName);
    const url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/Shared_Info.json?key=e2306b943c5e70ff7ba20b8bcfa95b289d78e103`;

    const responce = await UrlFetchApp.fetch(url, {
          contentType: 'application/json; charset=utf-8'
        }).getContentText();

    const result = JSON.parse(responce).wiki_page.text;
    this.getProjectRedmineData(result);
  }

  getProjectRedmineData(rawText) { // unused, but just for case, for the future
    const arr = rawText.split('\r\n');
    const regExp = `^[*][А-Яа-яA-Za-z ]+[*]: `;
    const resultArr = arr.filter(item => item.search(regExp) !== -1);

    for (let i = 0; i < resultArr.length; i++) {
      const elem = this.extractProperty(resultArr[i]);
      this.currentChange.properties[elem.key] = elem.value;
    }
  }

  extractProperty(strProperty) {
    const strSplit = '*: ';
    const pos = strProperty.indexOf(strSplit) + strSplit.length;

    const propName = strProperty.substring(1, pos - strSplit.length);
    const propValue = strProperty.includes(this.pmTitleRedmine) ? this.extractUserName(strProperty.substring(pos)) : strProperty.substring(pos);

    return {
      key: propName,
      value: propValue
    };
  }

  extractUserName(rawName) {
    const splitStr = `":https://egamings.slack.com`;
    const pos = rawName.indexOf(splitStr);
    return rawName.substring(1, pos);
  }

  getSlackLink(username) {

    let usernameToOutput = username;
    let slackID = '';
    const usernamesColumnIndex = 1;
    const slackIdColumnIndex = 2;
    const userNameAliasIndex = 3; // if userName is written by russian (UnitLeads), we output alias by english

    const lastRow = this.sheetSettings.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
      if (this.sheetSettings.getRange(i, usernamesColumnIndex).getValue() === username) {
        slackID =  this.sheetSettings.getRange(i, slackIdColumnIndex).getValue();
        const usernameAlias = this.sheetSettings.getRange(i, userNameAliasIndex).getValue();
        if ( usernameAlias.length > 0 ) usernameToOutput = usernameAlias;
        break;
      }
    }
    if (!slackID) {
      return usernameToOutput;
    }

    return `\"${usernameToOutput}\":https://egamings.slack.com/team/${slackID}`;
  }

  showNotify(redmineAlias, projectName) {
    let url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/`;
    let text = `Изменения в проекте ${projectName} занесены в Redmine Wiki. Открыть >>`; //todo hyperlink
    let textWithLink = SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(text.length - 10, text.length,  url).build();

    this.cellNotify.setBackgroundRGB(10,199, 145);
    this.cellNotify.setRichTextValue(textWithLink);

    SpreadsheetApp.flush();
    Utilities.sleep(this.notifyDuration * 1000);

    this.cellNotify.setValue('');
    this.cellNotify.setBackgroundRGB(254, 254, 254);
  }

}

app = new application();

function onEdit(event)
{

  //Redmine sync
  app.range = event.range;

  if ( ( event.source.getActiveSheet().getName() === app.sheetProjects.getName() )
    && app.isTrackedProjectChange(event.range)
    && ( app.isColumnTracked(event.range.getColumn()) )
    && (event.oldValue !== event.value)
    ) {

    app.oldValue = event.oldValue;
    app.newValue = event.value;
    app.range = event.range;

    app.handleChange();
  }

  var r = event.source.getActiveRange();
  var idCol = event.range.getColumn();
  if (idCol <= 22) {
    let userMail = Session.getActiveUser().getEmail();
    let currentMessage = r.getComment();
    if (userMail) {
      userMail = "\n" + userRmail;
    }
    let message = "Changed: " + getTime() + userMail + '\n\n' + currentMessage;
    // r.setComment(message);
    //Logger.log(r.getComment());
  }




}

function getTime() {
  var today = new Date();
  return Utilities.formatDate(today, 'GMT+03:00', 'dd.MM.yy HH:mm');
}

function consol() {
  Logger.log(Session.getActiveUser().getEmail());
}

function showNotify() {
  const sheetProjects = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ответственные и проекты");
  const cellNotify = sheetProjects.getRange(1, 3).getCell(1, 1);

  // const val = cellNotify.getValue();
  // Logger.log(`value=${val}`);
  cellNotify.setBackgroundRGB(10,199, 145);
  cellNotify.setValue('test');
  SpreadsheetApp.flush();
  Utilities.sleep(4 * 1000);

  cellNotify.setValue('');
  cellNotify.setBackgroundRGB(254, 254, 254);
}