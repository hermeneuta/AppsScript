function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Wypłaty')
    .addItem('Eksportuj dane', 'eksportDanych')
    .addItem('Dodaj osobę', 'addPerson')
    .addItem('Wyślij powiadomienia', 'sendMails')
    .addToUi();
}

/*
Eksport zdefiniowanych informacji do dokumentu określonej osoby
*/
function eksportDanych() {
  //aktywacja dokumentu z którego będzie dokonany eksport
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getActiveSheet();
  const wsName = ws.getName();
  const ui = SpreadsheetApp.getUi();
  
  if (wsName != 'Osoby') {
    //czyszczenie dokumentu z wierszy nie zawierających imienia i nazwiska
    cleaningRows_(ws);

    //osoby, których dotyczą faktury
    const osobySheet = ss.getSheetByName('Osoby');
    const osobyRows = osobySheet.getLastRow() - 1;
    const osoby = osobySheet.getRange(2, 1, osobyRows, 3).getValues();

    //Obiekt zawierający informacje o osobach
    const users = osoby.map(osoba => ({ name: osoba[0], id: osoba[1], mail: osoba[2] }));

    //Sortowanie danych ze względu na osoby
    const dataSorted = sortData_(ws, osoby);

    if (dataSorted != 'Brak Danych Osobowych') {
      //Rozsyłanie danych do określonych osób
      dataSorted.forEach((person) => {
        const exportForPerson = [];
        person
          //.map(el => el.map(elem => [elem]))
          .forEach(data => {
            exportForPerson.push(data.data);
            updateStatus_(ws, data);
          })
        exportToPerson_(exportForPerson, users, wsName);
      })
      ui.alert('Eksport przebiegł pomyślnie!');
    } else {
      ui.alert('Nie wszystkie osoby mają utworzone dokumenty!');
      return
    } 
  } else {
    ui.alert('Eksport nie może być przeprowadzony z zakładki "Osoby". Zmień zakładkę.')
    return
  }
  
}

//Clean before export -> rows without names
function cleaningRows_(ws) {
  const lastCol = ws.getLastColumn();
  const lastRow = ws.getLastRow();
  const values = ws.getRange(2, 1, lastRow, lastCol).getValues();
  Logger.log(values);
  values.reverse().forEach((r,i)=>{
      if (r[1]===''){
         ws.deleteRow(values.length-i+1);
      }
  });
}

////
//Sortowanie danych ze względu na osoby
function sortData_(ws, osoby) {

  //Filtrowanie po statusie (statusFilter_(ws)) i zwracanie tablicy z obiektami o parametrach: data, row
  const dataRowPair = statusFilter_(ws);
  const names = dataRowPair.map(el => el.data[1]);
  
  //Tablica z unikalnymi imionami osób, których dotyczą rachunki
  const uniqueNames = [...new Set(names)];

  //Sprawdzanie czy wszystkie osoby mają utworzone dokumenty z rachunkami
  if (osobyTest_(uniqueNames, osoby)) {
  //Grupowanie rachunków ze względu na osoby
  const valuesAll = uniqueNames.map(uniqName => {
    const valuesForPerson = [];
    dataRowPair.forEach(el => {
    if (el.data[1] === uniqName) {
      valuesForPerson.push({data: el.data, row: el.row});
    }
  })
  return valuesForPerson;
  })
    //Obiekt z rachunkami dla danej osoby wraz z informacją o numerze wiersza danego rachunku
   return valuesAll;
   } else {
    return 'Brak Danych Osobowych';
  }
  }

/////
//Sprawdzanie czy wszystkie osoby mają utworzone dokumenty z rachunkami
function osobyTest_(osobyFromMain, osobyFromOsobyArray) {
  //osobyFromMain zawierają wszystkie unikalne imiona osób z główego dokumentu, do których chce się eksportować dane
  //osobyFromOsobyObj zawierają wszystkie informacje dotyczące danej osoby z zakładki Osoby z dokumentu głównego
  const osobyFromOsoby = osobyFromOsobyArray.map(el => el[0]);
  
  const condition = osobyFromMain.filter(el => !osobyFromOsoby.includes(el));
  
  if (condition.length != 0) {return false}
  else {return true}

}

/////
//Sortowanie względem pierwszej kolumny (rachunki niezapłacone ostatnie)
function sorting_(ws) {
  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  const range = ws.getRange(2,1,lastRow,lastCol);
  range.sort(1);
}

////
//Eksport do pliku przypisanego osobie
function exportToPerson_(rawData, users, wsName) {

  //Jakiej osoby dotyczą dane
  const userName = rawData[0][1];
  const userObj = users.filter(user => user.name === userName);
  //Zdefiniowanie pliku danej osoby
  const fileId = userObj[0].id;
  
  const fileToExport = SpreadsheetApp.openById(fileId);
  const fileToExportSheet = fileToExport.getSheetByName(wsName);
  //Logger.log(rawData);
  //Format danych - tylko potrzebne kolumny
  const data = rawData
      .map(el => [el[0], el[3], el[4], el[5], el[6], el[7], el[8]]);

  //Eksport danych do pliku danej osoby
  //0.Sortowanie danych
  sorting_(fileToExportSheet);

  //1.Czyszczenie wiersza z pustą wartością w kolumnie faktura
  const lastRowBefore = fileToExportSheet.getLastRow();
  const fakturyCol = fileToExportSheet.getRange(2,1,lastRowBefore,1).getValues().flat();
  Logger.log(fakturyCol);
  fakturyCol.forEach((el, i) => {
    if (el === '') {
      fileToExportSheet.getRange(i+2, 1, 1, 8).clearContent();
    }
  })
  //2.Nanoszenie danych
  const lastRowAfter = fileToExportSheet.getLastRow();
  fileToExportSheet.getRange(lastRowAfter+1,1,data.length,data[0].length).setValues(data);
  //Dopasowywanie rozmiaru kolumn
  fileToExportSheet.autoResizeColumns(1, fileToExportSheet.getLastColumn());
}

//////
//Filtrowanie po statusie
function statusFilter_(ws) {
  const statusCol = 11; //Column of Status 

  const rows = ws.getLastRow() - 1;
  //Zczytywanie kolumny ze statusami
  const statusData = ws.getRange(2, statusCol, rows, 1).getValues().flat(); //flat() method reduces array of arrays (names) to regular array
  Logger.log(statusData)
  //rowsToExport -> rows without status or with czekające... status
  const rowsToSend = statusData.map((el, i) => [el, i+2]).filter(el => el[0] != 'Przesłano').map(el => el[1]);
  Logger.log(rowsToSend)
  const dataRowPair = rowsToSend.map(el => ({data: ws.getRange(el, 1, 1, 10).getValues().flat(), row: el}));

  return dataRowPair;
}

function updateStatus_(ws, data) {
  const i = data.row;
  data.data[0] != '' ? ws.getRange(i, 11, 1, 1).setValue('Przesłano') : ws.getRange(i, 11, 1, 1).setValue('Czekające...');
}

function exportHeadsNewDoc_(fileToExportSheet) {
//definicja nagłówków
  const heads = ['Faktura', 'Zlecono', 'Rodzaj umowy', 'Rodzaj kwoty', 'Kwota', 'Opis', 'Uwagi', 'Status'];
  heads.forEach((head, i) => {
    fileToExportSheet.getRange(1,i+1,1,1).setValue(head).setFontWeight("bold");
  });
}



