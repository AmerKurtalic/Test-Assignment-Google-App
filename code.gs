function appendSheetToDoc(sheetName, doc, body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName(sheetName);
  var dataRange = dataSheet.getDataRange();
  var dataValues = dataRange.getValues();

  if (sheetName !== "") {
    body.appendPageBreak();
  }

  if (sheetName === "Claims") {
    var dataValuesStr = dataValues.map(function(row) {
      return row.map(function(cell) {
        return cell.toString();
      });
    });

    var text = dataValuesStr.map(function(row) {
      return row.join("\t");
    }).join("\n");

    body.appendParagraph(text);
  
  } else {

    body.appendParagraph("");

    dataValues = dataValues.filter(function(row) {
      return row.some(function(cell) {
        return cell !== "";
      });
    });
    
    dataValues = dataValues.map(function(row) {
      return row.filter(function(cell) {
        return cell !== "";
      });
    });
    
    body.appendTable(dataValues.map(function(row) {
      return row.map(function(cell) {
        return cell.toString();
      });
    }));
     var dataValuesStr = dataValues.map(function(row) {
    return row.map(function(cell) {
      return cell.toString();
    });
  });
  }
}

function createDocument() {
  var ui = SpreadsheetApp.getUi();
  var fileName = ui.prompt("File Name", ui.ButtonSet.OK_CANCEL);
  if (fileName.getSelectedButton() == ui.Button.OK) {
    var reportHeader = ui.prompt("Doc Header", ui.ButtonSet.OK_CANCEL);
    if (reportHeader.getSelectedButton() == ui.Button.OK) {
      var folderList = DriveApp.getFolders();
      var folderNames = [];
      var folderIds = {};
      while (folderList.hasNext()) {
        var folder = folderList.next();
        folderNames.push(folder.getName());
        folderIds[folder.getName()] = folder.getId();
      }
      var destination = ui.prompt("Destination Folder", folderNames.join("\n"), ui.ButtonSet.OK_CANCEL);
      if (destination.getSelectedButton() == ui.Button.OK) {
        var folderName = destination.getResponseText();
        var destination_id = folderIds[folderName];
        var destination = DriveApp.getFolderById(destination_id);
        var doc = DocumentApp.create(fileName.getResponseText());
        var docID = doc.getId();
        var file = DriveApp.getFileById(docID);
        file.moveTo(destination);
        var body = doc.getBody();
        var imageBlob = DriveApp.getFileById('1N2an3WJId_kLalzOwIt0pxaYM8jB4X3S').getBlob();
        var image = body.insertImage(0, imageBlob);
        var imageParagraph = image.getParent();
            imageParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            image.setWidth(600);
            image.setHeight(500);
            body.insertParagraph(1, reportHeader.getResponseText())
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
        appendSheetToDoc('Cover Page', doc, body);
        appendSheetToDoc('Abstract', doc, body);
        appendSheetToDoc('Classification Codes', doc, body);
        appendSheetToDoc('Claims', doc, body);
        appendSheetToDoc('Application Events', doc, body);
      }
    }
  }
}

function addMenu()
{
  var menu = SpreadsheetApp.getUi().createMenu('Create as Document');
  menu.addItem('Create Doc', 'createDocument');
  menu.addToUi();
}

function onOpen(e)
{
  addMenu();
}
