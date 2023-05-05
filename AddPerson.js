/////
//Dodanie osoby, której dotyczą faktury
function addPerson() {
  //Get name and mail of new person
  const mainDoc = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = mainDoc.getActiveSheet();
  const nameSheet = activeSheet.getName();

  const ui = SpreadsheetApp.getUi();
  const respondName = ui.prompt("Imię i Nazwisko nowej osoby:");
  const name = respondName.getResponseText();
  const respondMail = ui.prompt("Mail do " + name + ": ");
  const mail = respondMail.getResponseText();

  //Create new document with title as name of person
  const newPersonSS = SpreadsheetApp.create(name);
  const fileId = newPersonSS.getId();
  const id = DriveApp.getFileById(fileId);

  //Move doc to the FSIT folder using DriveApp service
  const destinationFolderId = "ID_FOLDER";
  DriveApp.getFolderById(destinationFolderId).addFile(id);

  //Format doc with heads and mail info
  const fileToExport = SpreadsheetApp.openById(fileId);
  const fileToExportSheet = fileToExport.getActiveSheet();
  //Naming current tab
  fileToExportSheet.setName(nameSheet);
  //Creating heads
  exportHeadsNewDoc_(fileToExportSheet);
  //Writing mail
  fileToExport.insertSheet("Mail");
  const mailSheet = fileToExport.getSheetByName("Mail");
  mailSheet.getRange(1, 1, 1, 1).setValue("Mail").setFontWeight("bold");
  mailSheet.getRange(2, 1, 1, 1).setValue(mail);

  //Write down new information about new Person inside 'Osoby' tab in Main Doc
  const osobyTab = mainDoc.getSheetByName("Osoby");
  const lastRow = osobyTab.getLastRow();
  const values = [[name, fileId, mail]];
  osobyTab.getRange(lastRow + 1, 1, 1, 3).setValues(values);
  ui.alert(
    "Stworzono nowy dokument dla " +
      name +
      " oraz zaktualizowano zakładkę Osoby"
  );
}
