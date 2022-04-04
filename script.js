class application {

  constructor() {
    this.sheetProjects = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ответственные и проекты");
    this.sheetSettings = SpreadsheetApp.getActive().getSheetByName('service_info');

    this.range = undefined;
    this.oldValue = undefined;
    this.newValue = undefined;

    this.projectNameColumnIndex = 2;

    this.trackedColumns = {
      'status': {'titleTable': 'Статус', 'titleRedmine': 'Статус', 'index': 15 },
      'pm': {'titleTable': 'ПМ отв-й', titleRedmine: 'Ответственный PM', index: 8 },
      // 'statusColumnIndex': 15,
      // 'pmColumnIndex': 8,
    };

    this.pmTitleTable = this.sheetProjects.getRange(1, this.trackedColumns.pm.index).getValue();
    this.pmTitleRedmine = 'Ответственный PM';

    this.statusTitleTable = this.sheetProjects.getRange(1, this.trackedColumns.status.index).getValue();
    this.statusTitleRedmine = 'Статус';

    this.currentChange = {
      'projectName': undefined,
      properties: {},
    };

    this.trackedProjects = this.getTrackedProjects();
    this.redmineKey = 'e2306b943c5e70ff7ba20b8bcfa95b289d78e103';
  }

  getTrackedProjects() {
    //TODO case: exclude on-hold and closed projects
    const result = [];
    const redmineColumnIndex = 4;
    const projectColumnIndex = 5;
    const lastSheetRow = this.sheetSettings.getLastRow();

    const range = this.sheetSettings.getRange(`D2:D${lastSheetRow}`);
    //TODO empty cell among range - getNextDataCell()
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
    //TODO case: handle depends on edited sheet

    if ( this.isManyCellsChanged() ) {
      return;
    }

    this.currentChange.projectName = this.getProjectName();
    // Browser.msgBox('before getProjectData');

    // TODO fill this.currentChange by table values
    // await this.getProjectData(this.currentChange.projectName);
    this.loadProjectDataFromTable();

    Browser.msgBox(JSON.stringify(this.currentChange));

    this.publishToRedmine(this.currentChange);
  }

    loadProjectDataFromTable() {
      Browser.msgBox('loadProjectDataFromTable');
      Object.keys(this.trackedColumns).forEach(key => {
        const val = this.sheetProjects.getRange(this.range.getRowIndex(), this.trackedColumns[key].index).getValue();
        Browser.msgBox(`val: ${val}`);
        this.currentChange.properties[key] = val;
      });
      Browser.msgBox(JSON.stringify(this.currentChange));
    }


  //{"Статус":"Active Dev","Ответственный PM":"\"Andrew Boyarchuk\":https://egamings.slack.com/team/U01HWENP170"}
  // *Статус*: Active Dev\r\n*Ответственный PM*: "Andrew Boyarchuk":https://egamings.slack.com/team/U01HWENP170 - wiki
  publishToRedmine(change) {
    const redmineAlias = this.getProjectRedmineAlias(change.projectName);
    const props = change.properties;
    const url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/Shared_Info.json?key=e2306b943c5e70ff7ba20b8bcfa95b289d78e103`;
    let textContent = '';

    Object.keys(props).forEach(key => {
      textContent += `*${key}*: ${props[key]}\r\n`;
    });
    Browser.msgBox(`text content: ${textContent}`);
    //    continue: upload data to redmine
  }

  isTrackedFieldsChanged() {
    return this.range.getColumn() === this.statusColumnIndex || this.range.getColumn() === this.pmColumnIndex;
  }

  isManyCellsChanged() {
      //TODO reachable case?
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
    Browser.msgBox(this.source.getActiveSheet().getName());
    return this.source.getActiveSheet().getName() === this.sheetProjects.getName();
  }

  async getProjectData(projectName) {
    const redmineAlias = this.getProjectRedmineAlias(projectName);

    const url = `https://tracker.egamings.com/projects/${redmineAlias}/wiki/Shared_Info.json?key=e2306b943c5e70ff7ba20b8bcfa95b289d78e103`;

    const responce = await UrlFetchApp.fetch(url, {
          contentType: 'application/json; charset=utf-8'
        }).getContentText();

    const result = JSON.parse(responce).wiki_page.text;

    // Browser.msgBox(`responce = ${responce}`);
    // Browser.msgBox(`raw string = ${result}`);

    // Logger.log(json.wiki_page.text);
    // Logger.log(`project data: ${this.getProjectData(json.wiki_page.text)}`);

    this.getProjectRedmineData(result);
  }

  // *Статус*: Active Dev *Ответственный PM*: "Andrew Boyarchuk":https://egamings.slack.com/team/U01HWENP170
  // [*Статус*: Active Dev, *Ответственный PM*: "Andrew Boyarchuk":https://egamings.slack.com/team/U01HWENP170, ]

  getProjectRedmineData(rawText) {
    const result = {};
    //TODO refactoring

    // *Статус*: Active Dev *Ответственный PM*: "Andrew Boyarchuk":https://egamings.slack.com/team/U01HWENP170
    // Browser.msgBox(`rawText = ${rawText}`);


    const arr = rawText.split('\r\n');
    const regExp = `^[*][А-Яа-яA-Za-z ]+[*]: `;
    const resultArr = arr.filter(item => item.search(regExp) !== -1);

    for (let i = 0; i < resultArr.length; i++) {
      const elem = this.getProperty(resultArr[i]);
      this.currentChange.properties[elem.key] = elem.value;
    }

    Browser.msgBox(JSON.stringify(result));
    // return result;

    // Logger.log(result);
    // Logger.log(getSlackID('Andrew Boyarchuk'));
  }

  getProperty(strProperty) { //rename ?
    const result = new Object();
    const strSplit = '*: ';
    const pos = strProperty.indexOf(strSplit) + strSplit.length;

    const propName = strProperty.substring(1, pos - strSplit.length);
    // const propValue = strProperty.substring(pos);
    const propValue = strProperty.includes(this.pmTitleRedmine) ? this.extractUserName(strProperty.substring(pos)) : strProperty.substring(pos);

    // Logger.log(key);
    // Logger.log(value);
    // return extractName(value);
    // Browser.msgBox(`key: ${key}, value: ${value}`);
    return {
      key: propName,
      value: propValue
    };
  }

  // "Andrew Boyarchuk":https://egamings.slack.com/team/U01HWENP170
  extractUserName(rawName) {
    const splitStr = `":https://egamings.slack.com`;
    const pos = rawName.indexOf(splitStr);
    return rawName.substring(1, pos);
  }

  getSlackID(username) {
    //TODO case: no such username
    let result = '';
    const usernamesColumnIndex = 1;
    const slackIdColumnIndex = 2;

    const lastRow = this.sheetSettings.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
      Logger.log(this.sheetSettings.getRange(i, 1).getValue());
      if (this.sheetSettings.getRange(i, usernamesColumnIndex).getValue() === username) {
        result =  this.sheetSettings.getRange(i, slackIdColumnIndex).getValue();
        break;
      }
    }
    return result;
  }

}

app = new application();

function onOpen() {
  Browser.msgBox(`onOpen`);

  app.trackedProjects = app.getTrackedProjects();
  // Browser.msgBox(app.trackedProjects);
}

function onEdit(event) {
  app.range = event.range;

  Browser.msgBox('onEdit start');
  // Browser.msgBox(`app.isColumnTracked: ${app.isColumnTracked(event.range.getColumn())}`);

  if ( ( event.source.getActiveSheet().getName() === app.sheetProjects.getName() )
    && app.isTrackedProjectChange(event.range)
    && ( app.isColumnTracked(event.range.getColumn()) )
    && (event.oldValue !== event.value)
    ) {
      Browser.msgBox('will handle this change');
      // Browser.msgBox(`current project edited: ${app.currentChange.projectName}`);

    app.oldValue = event.oldValue;
    app.newValue = event.value;
    app.range = event.range;

    app.handleChange();
  }
}



























  // async getProjectData() {
  //   const url = `https://tracker.egamings.com/projects/${this.projectID}/wiki/Shared_Info.xml?key=${this.redmineKey}`;
  //   Browser.msgBox(url);
  //   const xml = await UrlFetchApp.fetch(url, {
  //     contentType: 'application/xml; charset=utf-8'
  //   }).getContentText();

  //   const doc = XmlService.parse(xml);
  //   Browser.msgBox(`has root: ${doc.hasRootElement}`);
  //   const root = doc.getRootElement();
  //   const text = root.getText();
  //   // const text = root.getChild('text');
  //   Browser.msgBox(text);
  // }