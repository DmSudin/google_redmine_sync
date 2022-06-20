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
      'status': { 'titleTable': 'Статус', 'titleRedmine': 'Статус', 'columnIndex': 15 },
      'pm': { 'titleTable': 'ПМ отв-й', 'titleRedmine': 'Ответственный PM', 'columnIndex': 8 },
      'unitLead': {'titleTable': 'ЮнитЛид', 'titleRedmine': 'Ответственный Unit Lead', 'columnIndex': 10 }
    };
    this.defaultProjectData = {
      rowIndex: null,
      status: null,
      pm: null,
      unitLead: null,
    };

    this.requestsForUpdate = [];

    // original
    this.tableURL = 'https://docs.google.com/spreadsheets/d/1pZtZn8cAxxPDzwQNAkPs_aLneZaTx7RpQrNN9OLh3cg';

    // my copy
    // this.tableURL = 'https://docs.google.com/spreadsheets/d/1ArRJQ_20pOxefiM7yBJqHVFPuuMwtbYLHfvIPLya7MI'

    this.tableTitlesRowIndex = 2;

    this.sheetProjects = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ответственные и проекты");
    this.sheetSettings = SpreadsheetApp.getActive().getSheetByName('Redmine_sync');
    this.projectNameColumnIndex = 2;

    this.cellNotify = this.sheetProjects.getRange(1, 3).getCell(1, 1);
    this.notifyDuration = 10;

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
    // const lastProjectsRowIndex = range.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();
    const lastProjectsRowIndex = range.getDisplayValues().length + 1;

    for (let i = 2; i <= lastProjectsRowIndex; i++) {
      const redmineAlias = this.sheetSettings.getRange(i, redmineColumnIndex).getValue();
      const projectName = this.sheetSettings.getRange(i, projectColumnIndex).getValue();

      if (redmineAlias.length && projectName.length) {
        const projectItem = new Object();
        projectItem.redmineAlias = redmineAlias;
        projectItem.projectName = projectName;
        // shows if table data of project is equal to redmine wiki data
        projectItem.redmineData = {};
        Object.assign(projectItem.redmineData, this.defaultProjectData);
        projectItem.tableData = this.loadProjectDataFromTable(projectItem.projectName);
        result.push(projectItem);
      }
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
    return this.trackedProjects.filter(projectItem => projectItem.projectName === projectName)[0].redmineAlias;
  }

  async handleChange() {

    if ( this.isManyCellsChanged() ) {
      return;
    }

    this.currentChange.projectName = this.getProjectName();
    this.loadProjectDataFromTable();
    this.publishChangeToRedmine(this.currentChange);
  }

  loadProjectDataFromTable(projectName) {

    if (!projectName) {
      // called by cell change handler
      Object.keys(this.trackedColumns).forEach(key => {
        const val = this.sheetProjects.getRange(this.range.getRowIndex(), this.trackedColumns[key].columnIndex).getValue();
        this.currentChange.properties[key] = val;
      });
    } else {
      // called by sheduled task of updating projects
      const result = {};
      Object.assign(result, this.defaultProjectData);
      const projectRowIndex = this.getTableRowIndexForProject(projectName);
      const lasColumnIndex = this.sheetProjects.getLastColumn();
      const projectRowValues = this.sheetProjects.getSheetValues(projectRowIndex, 1, 1, lasColumnIndex);


      result['rowIndex'] = projectRowIndex;
      for (let key of Object.keys(this.trackedColumns)) {
        const value = projectRowValues[0][this.trackedColumns[key].columnIndex - 1];

        if (key === 'unitLead') {
          result[`${key}`] = this.getSlackName(value);

        } else {
          result[`${key}`] = value;
        }
      }
      return result;
    }
  }

  getTableRowIndexForProject(projectName) {
    const lastRowIndex = this.sheetProjects.getLastRow();
    for (let i = 3; i <= lastRowIndex; i++) {
      if (this.sheetProjects.getRange(i, this.projectNameColumnIndex).getValue() === projectName) {
        return i;
      }
    }
    return null;
  }

  async fetchProjectsDataFromRedmine() {
    for (let i = 0; i < this.trackedProjects.length; i++) {
      this.trackedProjects[i].redmineData = await this.loadProjectDataFromRedmine(this.trackedProjects[i].redmineAlias);
    }
  }

  async publishChangeToRedmine(change) {
    const redmineAlias = this.getProjectRedmineAlias(change.projectName);
    const url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/Shared_Info.json?key=e2306b943c5e70ff7ba20b8bcfa95b289d78e103`;
    let textContent = `"Таблица ответственных":${this.getLinkCellProject(change.projectName)}\r\n\r\n`;
    const props = change.properties;

    Object.keys(props).forEach(key => {
      const propName = this.trackedColumns[key].titleRedmine;
      let propValue = '';

      if ( ((propName) === this.trackedColumns.pm.titleRedmine) || ((propName) === this.trackedColumns.unitLead.titleRedmine) ) {
        propValue = this.getSlackLink(props[key]);
      } else propValue = props[key];

      textContent += `*${propName}*: ${propValue}\r\n`;
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
      this.showNotify(redmineAlias, projectName);
    } else Browser.msgBox(`Something went wrong while updating Redmine Wiki`);

    this.reset();
  }

  generateRequestForUpdate(redmineAlias, tableData) {
    const data = {};
    Object.assign(data, tableData);
    const url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/Shared_Info.json?key=e2306b943c5e70ff7ba20b8bcfa95b289d78e103`;
    const hyperLinkSuffix = `/edit#gid=0&range=${tableData.rowIndex}:${tableData.rowIndex}`;
    delete data['rowIndex'];
    let textContent = `"Таблица ответственных":${this.tableURL}${hyperLinkSuffix}\r\n\r\n`;
    for (const [key, value] of Object.entries(data)) {
      const propName = this.trackedColumns[key].titleRedmine;
      let propValue;
      switch (key) {
        case 'pm': propValue = this.getSlackLink(value);
        break;

        case 'unitLead': propValue = this.getSlackLink(value, true);
        break;

        default: propValue = value;
        break;
      }
      textContent += `*${propName}*: ${propValue}\r\n`;
    }

    const dataToSend = {
      "wiki_page":
      {
        "text": `${textContent}`,
      },
    };

    const request = {
      'url': url,
      'method': 'put',
      'contentType': 'application/json',
      'payload': JSON.stringify(dataToSend),
    };

    this.requestsForUpdate.push(request);
  }

  isManyCellsChanged() {
      return ((this.range.getNumColumns() !==1) || (this.range.getNumRows() !== 1));
  }

  isColumnTracked(columnNumber) {
    let result = false;
    Object.entries(this.trackedColumns).map((elem) => elem[1].columnIndex).forEach(item => {
      if (Math.trunc(columnNumber) === Math.trunc(item)) {
        result = true;
      }
    });

    return result;
  }

  isProjectsSheetEdited() {
    return this.source.getActiveSheet().getName() === this.sheetProjects.getName();
  }

  isEqualProjectData(tableData, redmineData) {
    let isDiffValueFound;

    const keys = Object.keys(tableData);
    if (keys.length !== Object.keys(redmineData).length) return false;

    for (let i = 0; i < keys.length; i++) {
      const propName = keys[i];
      if (tableData[propName] !== redmineData[propName]) {
        isDiffValueFound = true;
        break;
      }
    }

    if (!isDiffValueFound) return true;
    return false;
  }

  async loadProjectDataFromRedmine(redmineAlias) {
    const url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/Shared_Info.json?key=e2306b943c5e70ff7ba20b8bcfa95b289d78e103`;

    const responce = await UrlFetchApp.fetch(url, {
          contentType: 'application/json; charset=utf-8'
        }).getContentText();

    const json = JSON.parse(responce).wiki_page.text;
    return this.parseResponseFromRedmine(json);
  }

  parseResponseFromRedmine(rawText) {

    const projectData = {
      rowIndex: null,
      status: null,
      pm: null,
      unitLead: null,
    };

    // extract row number in table for project
    const tableLinkText = rawText.split('\r\n\r\n')[0];
    const regexp = `range=`;
    const projectRangeText = tableLinkText.slice(tableLinkText.search(regexp) + regexp.length);
    const rowNumber = parseInt(projectRangeText.split(`:`)[1]);
    projectData['rowIndex'] = rowNumber;


    const delimiterIndex = rawText.search('\r\n\r\n') + 1;
    const propertiesText = rawText.slice(delimiterIndex);

    const arr = propertiesText.split('\r\n');
    // regular expr for recognize strings as properties, e.g.:
    // `*PROP_NAME*: PROP_VALUE`
    const regExp = `^[*][А-Яа-яA-Za-z ]+[*]: `;
    const resultArr = arr.filter(item => item.search(regExp) !== -1);

    const entries = Object.entries(this.trackedColumns);
    for (let i = 0; i < resultArr.length; i++) {
      let elem = this.extractProperty(resultArr[i]);
      let ourEntry = entries.filter(entry => entry[1].titleRedmine === elem.key);
      const propName = ourEntry[0][0];
      projectData[propName] = elem.value;

      // required only if wiki update is fired upon cell change
      // to remove further
      // this.currentChange.properties[elem.key] = elem.value;
    }
    return projectData;
  }

  extractProperty(strProperty) {
    const strSplit = '*: ';
    const pos = strProperty.indexOf(strSplit) + strSplit.length;
    let propValue = null;

    const propName = strProperty.substring(1, pos - strSplit.length);
    if ((strProperty.includes(this.trackedColumns.pm.titleRedmine)) || (strProperty.includes(this.trackedColumns.unitLead.titleRedmine))) {
      propValue = this.extractUserName(strProperty.substring(pos));
    } else  {
      propValue = strProperty.substring(pos);
    }

    return {
      key: propName,
      value: propValue
    };
  }

  extractUserName(rawName) {
    const splitStr = `":https://egamings.slack.com`;
    const pos = rawName.indexOf(splitStr);
    if (pos === -1) return rawName;

    return rawName.substring(1, pos);
  }

  getSlackLink(username, isSlackNameGiven = false) {
    if (!username) return '';
    let usernameToOutput = username;
    let slackID = '';
    const tableUsernamesColumnIndex = 1;
    const slackIdColumnIndex = 2;
    const slackUserNamesColumnIndex = 3; // if userName is in russian (e.g. Unit Lead), we output alias in english
    const lastRow = this.sheetSettings.getLastRow();

    if (isSlackNameGiven) {
      for (let i = 1; i < lastRow; i++) {
        if (this.sheetSettings.getRange(i, slackUserNamesColumnIndex).getValue() === username) {
          slackID = this.sheetSettings.getRange(i, slackIdColumnIndex).getValue();
          break;
        }
      }
    } else {
      const userSlackName = this.getSlackName(username);
      if ( userSlackName.length > 0 ) usernameToOutput = userSlackName;
      for (let i = 1; i <= lastRow; i++) {
        if (this.sheetSettings.getRange(i, tableUsernamesColumnIndex).getValue() === username) {
          slackID =  this.sheetSettings.getRange(i, slackIdColumnIndex).getValue();
          break;
        }
      }
    }
    if (!slackID) {
      return usernameToOutput;
    }
    return `\"${usernameToOutput}\":https://egamings.slack.com/team/${slackID}`;
  }

  getSlackName(username) {
    let result = null;
    const lastSheetRow = this.sheetSettings.getLastRow();
    const usernamesColumnIndex = 1;
    const userSlackNamesIndex = 3;
    for (let i = 2; i <= lastSheetRow; i++)  {
      if (this.sheetSettings.getRange(i, usernamesColumnIndex).getValue() === username) {
        result = this.sheetSettings.getRange(i, userSlackNamesIndex).getValue();
        break;
      }
    }
    if (result && result.length) return result;
    return username;
    }

  getLinkCellProject(projectName) {
    const linkSuffix = '/edit#gid=0&range=';
    const projectRowIndex = this.getTableRowIndexForProject(projectName);

    if (!projectRowIndex || !projectName) {
      return '';
    }

    return `${this.tableURL}${linkSuffix}${projectRowIndex}:${projectRowIndex}`;
  }

  showNotify(redmineAlias, projectName) {
    let url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/`;
    let text = `Изменения в проекте ${projectName} занесены в Redmine Wiki. Открыть >>`;
    let textWithLink = SpreadsheetApp.newRichTextValue().setText(text).setLinkUrl(text.length - 10, text.length,  url).build();

    this.cellNotify.setBackgroundRGB(10,199, 145);
    this.cellNotify.setRichTextValue(textWithLink);

    SpreadsheetApp.flush();
    Utilities.sleep(this.notifyDuration * 1000);
    this.clearNotify();
  }

  clearNotify() {
    this.cellNotify.setValue('');
    this.cellNotify.setBackgroundRGB(254, 254, 254);
  }

  async synchronizeWithRedmineWiki() {
    let isChangesExists = false;
    await this.fetchProjectsDataFromRedmine();

    for (const projectItem of this.trackedProjects) {
      if (!this.isEqualProjectData(projectItem.tableData, projectItem.redmineData)) {
        isChangesExists = true;
        Logger.log(`Info about project ${projectItem.projectName} was changed. Updating...`);
        this.generateRequestForUpdate(projectItem.redmineAlias, projectItem.tableData);
      }
    }
    // Logger.log(`fetched projects data: ${JSON.stringify(this.trackedProjects)}`);

    if (isChangesExists) {
      await this.sendRequestsForUpdate();
    } else {
      Logger.log(`Everything is up-to-date.`);
    }
  }

  async sendRequestsForUpdate() {
    // max requests for fetchAll method = 100
    // see https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app#fetchallrequests

    const lastIndex = this.requestsForUpdate.length - 1;
    const responseCodes = [];
    if (lastIndex < 99) {
      const responses = await UrlFetchApp.fetchAll(this.requestsForUpdate);
      responseCodes.push(...responses.map((response) => response.getResponseCode()));

    } else {
      // need to test when projects amount will be more than 99
      let requestsPack = [];
      const chunkSize = 99;
      for (let i = 0; i < this.requestsForUpdate.length; i += chunkSize) {
        requestsPack.push(this.requestsForUpdate.slice(i, i + chunkSize));
      }
        for (const chunk of requestsPack) {
          const responses = await UrlFetchApp.fetchAll(chunk);
          responseCodes.push(...responses.map((response) => response.getResponseCode()));
        }
    }
    responseCodes.sort((a, b) => a - b);

    if ((responseCodes[0] === responseCodes[responseCodes.length - 1]) && (responseCodes[0] === 204)) {
      Logger.log('All done successfully');
    } else {
      Logger.log('some errors occured.');
    }
  }

}

app = new application();

function onOpen() {
  app.clearNotify();
}

function onEdit(event) {
  var r = event.source.getActiveRange();
  var idCol = event.range.getColumn();
  if (idCol <= 22) {
    let userMail = Session.getActiveUser().getEmail();
    let currentMessage = r.getComment();
    if (userMail) {
      userMail = "\n" + userMail;
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

  cellNotify.setBackgroundRGB(10,199, 145);
  cellNotify.setValue('test');
  SpreadsheetApp.flush();
  Utilities.sleep(4 * 1000);

  cellNotify.setValue('');
  cellNotify.setBackgroundRGB(254, 254, 254);
}

// SCR #363124
async function runRedmineSynch() {
  await app.synchronizeWithRedmineWiki();
}
