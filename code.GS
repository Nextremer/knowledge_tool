var FILE_NAME_PHRASES = 'phrases.xml';
var SHEET_NAME_CHECKER = 'checker_table'

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .createMenu('Tool-Dialog-Menu')
        .addItem('Open', 'openDialog')
        .addToUi();
}

function openDialog() {
  var html = HtmlService.createTemplateFromFile('sidebar').evaluate();
  SpreadsheetApp.getUi().showSidebar(html);
}

function getPhrases() {
  var groupList　 = [];
  var groupNameList = [];
  var phraseList = [];
  var phraseListToGroup = [];

  var ownFileId = SpreadsheetApp.getActiveSpreadsheet().getId()
  var ownFolderId = DriveApp.getFileById(ownFileId).getParents().next().getId();
  var ownPhrasesFileId = '';

  var files = DriveApp.getFolderById(ownFolderId).getFiles();
  while(files.hasNext()){
    var file = files.next();
    if(file.getName() === FILE_NAME_PHRASES){
      ownPhrasesFileId = file.getId();
    }
  }

  var data = DriveApp.getFileById(ownPhrasesFileId).getBlob().getDataAsString();
  data = data.replace(/（/g, '(');
  data = data.replace(/）/g, ')');

  var xml = XmlService.parse(data);
  groupList = xml.getRootElement().getChildren();

  for (var i=0; i<groupList.length; i++){
    groupNameList.push(groupList[i].getAttribute('id').getValue());
    var phrases = groupList[i].getChildren();

    for (var j=0; j< phrases.length; j++){
      phraseList.push(phrases[j].getText());
      phraseListToGroup.push(groupList[i].getAttribute('id').getValue());
    }
  }

  var res = new Object();
  res.groupList = groupNameList;
  res.phraseList = phraseList;
  res.phraseListToGroup = phraseListToGroup;
  res.xml = data;

  return res;
}

function setText(inputText){
  var activeCell = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  var newVal = (activeCell.getValue() == '' ? inputText : activeCell.getValue() + '; ' + inputText);
  activeCell.setValue(newVal);
}

function setXml(xml){
  var ownFileId = SpreadsheetApp.getActiveSpreadsheet().getId()
  var ownFolderId = DriveApp.getFileById(ownFileId).getParents().next().getId();
  var ownPhrasesFileId = '';

  var files = DriveApp.getFolderById(ownFolderId).getFiles();
  while(files.hasNext()){
    var file = files.next();
    if(file.getName() === 'test.xml'){
      ownPhrasesFileId = file.getId();
    }
  }
  DriveApp.getFileById(ownPhrasesFileId).setContent("test");
}

function getCheckers(){
  var checkerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_CHECKER);
  var lastRow = checkerSheet.getLastRow();
  var checkerName = [], checkerValue = [], value2Name = [];
  var res = new Object();

  for (var i=2; i<lastRow; i++){
    var name = checkerSheet.getRange(i, 1).getValue();
    checkerName.push(name);

    var values = checkerSheet.getRange(i, 2).getValue().replace(/"/g, '').replace(/ /g, '');
    var arrayValues = values.split(',');

    for (j=0; j<arrayValues.length; j++){
      checkerValue.push(arrayValues[j]);
      value2Name.push(name);
    }
  }

  res.checkerName = checkerName;
  res.checkerValue = checkerValue;
  res.value2Name = value2Name;

  return res;
}
