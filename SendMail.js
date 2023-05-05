function sendMails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName("Osoby");
  const lastRow = ws.getLastRow();
  const osobyValues = ws.getRange(2, 1, lastRow - 1, 3).getValues();

  const now = new Date();
  const setDate = now.toLocaleDateString("pl-PL", {
    year: "2-digit",
    month: "2-digit",
    day: "2-digit",
  });

  osobyValues.forEach((el, i) => {
    send_(el[1]);
    ws.getRange(i + 2, 4, 1, 1).setValue(setDate);
  });
}

function send_(id) {
  const ss = SpreadsheetApp.openById(id);
  const ws = ss.getActiveSheet();
  const headers = ws.getRange("A1:G1").getValues().flat();
  const mailSheet = ss.getSheetByName("Mail");
  const email = mailSheet.getRange("A2").getValue();
  const faktura = headers[0];
  const zlecono = headers[1];
  const rodzajUmowy = "Umowa";
  const rodzajKwoty = "Podatek";
  const opis = headers[5];
  const kwota = headers[4];
  const uwagi = headers[6];

  //Filtoranie po statusie - tylko te wiersze, które nie zostały już wysłane
  //statusFilter_(ws, headers) zwraca tablice obiektów z parametrami: data, row;
  const statusFilteredData = statusFilterMails_(ws);

  //Podział na zapłacone rachunki (alreadyPaidData) i czekające na realizację (waitingData)
  const alreadyPaidData = statusFilteredData.filter((el) => el.data[0] != "");
  const waitingData = statusFilteredData.filter((el) => el.data[0] === "");

  //Extracting values from alreadyPaidData and waitingData
  const alreadyPaidValues = alreadyPaidData.map((el) => el.data);
  const totalPaid = alreadyPaidValues
    .reduce((acc, cur) => acc + Number(cur[4]), 0)
    .toFixed(2);
  const waitingValues = waitingData.map((el) => el.data);
  const totalWaiting = waitingValues
    .reduce((acc, cur) => acc + Number(cur[4]), 0)
    .toFixed(2);

  //Checking if alreadyPaidValues or waitingValues aren't empty. If yes modify html outcome
  const isPaid = alreadyPaidValues.length != 0 ? true : false;
  const isWaiting = waitingValues.length != 0 ? true : false;

  //Checkig if both alreadyPaidValues and waitingValues are empty. If yes don't send mail at all
  const isSending =
    alreadyPaidValues.length === 0 && waitingValues.length === 0 ? false : true;

  const htmlTemplate = HtmlService.createTemplateFromFile("temp");
  //Passing variables to html template
  htmlTemplate.faktura = faktura;
  htmlTemplate.opis = opis;
  htmlTemplate.zlecono = zlecono;
  htmlTemplate.kwota = kwota;
  htmlTemplate.uwagi = uwagi;
  htmlTemplate.rodzajUmowy = rodzajUmowy;
  htmlTemplate.rodzajKwoty = rodzajKwoty;
  htmlTemplate.totalPaid = totalPaid;
  htmlTemplate.totalWaiting = totalWaiting;
  htmlTemplate.alreadyPaidValues = alreadyPaidValues;
  htmlTemplate.isPaid = isPaid;
  htmlTemplate.isWaiting = isWaiting;
  htmlTemplate.waitingValues = waitingValues;

  const htmlForEmail = htmlTemplate.evaluate().getContent();

  Logger.log(alreadyPaidValues);
  Logger.log(waitingValues);

  if (isSending) {
    GmailApp.sendEmail(
      email,
      "Wypłaty",
      "Pleas open this email with client that supports HTML",
      { htmlBody: htmlForEmail }
    );
  } else {
    Logger.log("Nothing to send");
  }

  //Updating status
  updatingStatus_(ws, alreadyPaidData, waitingData);

  //Sorting data
}

//////
//Filtrowanie po statusie
function statusFilterMails_(ws) {
  const statusCol = 8; //Column of Status

  const rows = ws.getLastRow() - 1;
  //Zczytywanie kolumny ze statusami
  const statusData = ws.getRange(2, statusCol, rows, 1).getValues().flat(); //flat() method reduces array of arrays (names) to regular array

  //rowsToExport -> rows without status or with czekające... status
  const rowsToSend = statusData
    .map((el, i) => [el, i + 2])
    .filter((el) => el[0] != "Wysłano")
    .map((el) => el[1]);
  const dataRowPair = rowsToSend.map((el) => ({
    data: ws.getRange(el, 1, 1, 10).getDisplayValues().flat(),
    row: el,
  }));

  return dataRowPair;
}

//////
//Updating status
function updatingStatus_(ws, paid, waiting) {
  //Paid update
  const rowPaid = paid.map((el) => el.row).flat();
  rowPaid.forEach((row) => {
    ws.getRange(row, 8, 1, 1).setValue("Wysłano");
  });

  //Waiting updata
  const rowWaiting = waiting.map((el) => el.row).flat();
  rowWaiting.forEach((row) => {
    ws.getRange(row, 8, 1, 1).setValue("Powiadomiono");
  });
}
