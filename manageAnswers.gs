/**
 * The function returns the index of the last added row hiding duplicates based on fiscal code (that should be unique for each person)
 * @param sheet Google Sheet reference
 * @param fiscalCode Fiscal code 
 * @returns The index of the last row added
 */
function getAddedRow(sheet, fiscalCode){
  let rows = [];

  // Find all the rows containing the inserted fiscal code
  for(let row=2; row<=sheet.getLastRow(); row++) {
    if(getCell(sheet, row, CODICE_FISCALE).getValue() == fiscalCode) 
      rows.push(
        {
          date : new Date(getCell(sheet, row, INFO_INSERIMENTO).getValue()),
          index : row
        }
      );
  }

  // Sort in ascending order
  rows.sort((a, b)=> a.date - b.date);

  let newestRow = rows.pop();

  // Copy the newest data if present and if the newest entry doesn't have any data
  if(getCell(sheet, newestRow.index, LINK_CERTIFICATO_MEDICO).getValue().length == 0){
    for(let r=rows.length-1; r>0; r--){
      let content = getCell(sheet, rows[r].index, LINK_CERTIFICATO_MEDICO).getFormula();
      if(content.length > 0){
        getCell(sheet, newestRow.index, LINK_CERTIFICATO_MEDICO).setFormula(content);
        break;
      }
    }
  }

  if(getCell(sheet, newestRow.index, LINK_PRIVACY).getValue().length == 0){
    for(let r=rows.length-1; r>0; r--){
      let content = getCell(sheet, rows[r].index, LINK_PRIVACY).getFormula();
      if(content.length > 0){
        getCell(sheet, newestRow.index, LINK_PRIVACY).setFormula(content);
        break;
      }
    }
  }

  if(getCell(sheet, newestRow.index, LINK_RATA_1).getValue().length == 0){
    for(let r=rows.length-1; r>0; r--){
      let content = getCell(sheet, rows[r].index, LINK_RATA_1).getFormula();
      if(content.length > 0){
        getCell(sheet, newestRow.index, LINK_RATA_1).setFormula(content);
        break;
      }
    }
  }

  if(getCell(sheet, newestRow.index, LINK_RATA_2).getValue().length == 0){
    for(let r=rows.length-1; r>0; r--){
      let content = getCell(sheet, rows[r].index, LINK_RATA_2).getFormula();
      if(content.length > 0){
        getCell(sheet, newestRow.index, LINK_RATA_2).setFormula(content);
        break;
      }
    }
  }

  // Delete (hide for now) the older rows
  rows.forEach(row => sheet.hideRows(row.index));

  return newestRow.index;
}

/**
 * Function add a hyperlink to a cell. If the link is provided as parameter the content of the cell will be overwritten.  
 * @param sheet Google Sheet reference
 * @param row Row affected by the modification 
 * @param columnName Name of the column where the link must be changed
 * @param label Label placed instead of the full link
 * @param documentType Name of the document contained by the link (for logs)
 * @param link Link to substitute
 * @returns nothing
 */
function changeLinkToLabel(sheet, row, columnName, label, documentType, link = null){
  // let headerRow = sheet.getRange("1:1").getValues()[0];
  // let column = headerRow.indexOf(columnName)+1;
  // let cell = sheet.getRange(row, column);7
  let cell = getCell(sheet, row, columnName);
  let content = cell.getValue(); //cell.getValue(); 

  // The link to modify the response is provided by the form
  if(content.length == 0 && link != null){
    cell.setFormula(Utilities.formatString('=HYPERLINK("%s";"%s")', link, label));
    Logger.log("Aggiunta del link %s", documentType);
    return;
  }

  // Otherwise check for link inserted automatically 
  if (content.indexOf("drive.google") > -1) {   
    // Link found
    Logger.log("Aggiunta del link %s", documentType);
    cell.setFormula(Utilities.formatString('=HYPERLINK("%s";"%s")', content.toString(), label));
  }else{
    // Check if there is a strange content in the cell
    if(content.length != 0 && content != label) 
      Logger.log("(%s) Contenuto della cella (%d,%d) errato o personalizzato: %s", documentType, row, column, content);
    // else 
      // nop
  }
}

const UPPERCASE = 0;
const LOWECASE = 1;
const FISRT_CAPITAL = 2; 
/**
 * Function to format text fields
 * @param sheet Google Sheet reference
 * @param row Row of reference within resides the searched cell 
 * @param columnName Name of the cell's column to format
 * @param mode UPPERCASE, LOWECASE or FISRT_CAPITAL
 */
function formatField(sheet, row, columnName, mode){
  let cell = getCell(sheet, row, columnName);
  let content = cell.getValue();
  switch(mode){
    case UPPERCASE:
      cell.setValue(content.split(' ').map(word => word.trim().toUpperCase()).join(' '));
    break;
    case LOWECASE:
      cell.setValue(content.split(' ').map(word => word.trim().toLowerCase()).join(' '));
    break;
    case FISRT_CAPITAL:
      cell.setValue(content.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' '));
    break;
    default:
      Logger.log("Metodo di formattazione non trovato per la cella (%d, %d)", row, column);
  }
}

/**
 * Function to set the formula in the columns for payments based on the presence of links in the columns with the proof of payment
 * @param sheet Google Sheet reference
 * @param row Row of reference within resides the cells to modify 
 */
function setPayments(sheet, row){
  let linkReceipt = getCell(sheet, row, LINK_RATA_1);
  let paymentCell = getCell(sheet, row, PRIMA_RATA).setHorizontalAlignment("center");
  paymentCell.setFormula('=IF(ISBLANK(' + linkReceipt.getA1Notation()+ '); 0; 1)');

  linkReceipt = getCell(sheet, row, LINK_RATA_2);
  paymentCell = getCell(sheet, row, SECONDA_RATA).setHorizontalAlignment("center");
  paymentCell.setFormula('=IF(ISBLANK(' + linkReceipt.getA1Notation()+ '); 0; 1)');
}

/**
 * Find the age of a person given when he/she was born and a reference date
 * @param bornDate Date of born
 * @param date Date of reference
 * @returns Age of person at specific date
 */
function getAge(bornDate, date) {
  let eta =  new Date(bornDate);
  let eMonth = (eta.getMonth() + 1 ) % 12;
  let eDay = eta.getDate();

  let etaAtDate = new Date(date);
  let eaxMonth = (etaAtDate.getMonth() + 1) % 12;
  let eaxDay = etaAtDate.getDate();

  let age = etaAtDate.getFullYear() - 1 - eta.getFullYear();

  if(eMonth < eaxMonth || (eMonth == eaxMonth && eDay <= eaxDay)) return age + 1;
  return age;
}

/**
 * Function to add an event in the expiration day of the Certificato medico.
 * The function deletes the previous event substituted by the new one
 */
function addExpirationEvent(sheet, row){
  let completeName = getCell(sheet, row, NOME_ATLETA).getValue() + " " + getCell(sheet, row, COGNOME_ATLETA).getValue();
  let cExpirationDate = getCell(sheet, row, SCADENZA_CERTIFICATO).getValue();
  if(cExpirationDate == null || cExpirationDate.length < 1){
    Logger.log("Nessuna data di scadenza certificato medico trovata per %s", completeName);
    return;
  }
  let bornDate = getCell(sheet, row, DATA_NASCITA_ATLETA).getValue();
  let email = getCell(sheet, row, EMAIL).getValue();
  let linkToModify = getCell(sheet, row, LINK_RISPOSTA).getRichTextValue().getLinkUrl();
  let fiscalCode = getCell(sheet, row, CODICE_FISCALE).getValue().toString().trim().replace("\n", "");

  if(isNaN(new Date(cExpirationDate))){
    Logger.log("Scadenza certificato per %s non valida (%s)", completeName, cExpirationDate);
    return;
  } 

  cExpirationDate = new Date(cExpirationDate);
  let age = getAge(bornDate, cExpirationDate);

  let newEventName = Utilities.formatString("Scandenza certificato medico %s (%d anni)", completeName, age);

  let calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0];
  if(calendar == null) Logger.log("Impossibile ottenere calendario \"Amministrazione atleti\"");

  // Check 20y in case someone miss-clicked
  let fromDate = cExpirationDate.getFullYear() - 10 + "-" + cExpirationDate.getMonth() + "-" + cExpirationDate.getDate(); // Date from which we need to check past events
  let toDate = cExpirationDate.getFullYear() + 10 + "-" + cExpirationDate.getMonth() + "-" + cExpirationDate.getDate();   // Date to which we need to check past events

  // Delete the previous inserted expiration events (both previous and miss-clicked events)
  let previousEvents = calendar.getEvents(new Date(fromDate), new Date(toDate), { search: completeName }).filter((event) => {
    // Second check on the fiscal code (in case two people have the same complete name)
    return event.getDescription().search(fiscalCode) > -1;
  }); 

  // Delete all old events
  let res = 0;
  for (let i = 0; i < previousEvents.length; i++, res++) {
    previousEvents[i].deleteEvent();
  }

  if(res > 0) Logger.log("Eliminati " + res + " eventi passati per l'altleta %s (%s)", completeName, fiscalCode);
  else Logger.log("Non è stato eliminato nessun evento passato per l'atleta %s (%s)", completeName, fiscalCode);

  let options = {description : completeName + "\n" + fiscalCode + "\n" + email + "\n" + linkToModify};
  let calendarEvent = calendar.createAllDayEvent(newEventName, cExpirationDate, options); 
  Logger.log("Evento creato: " + newEventName + " il " + cExpirationDate);
}

// function updateEvents(){
//   let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   for(let row=2; row<sheet.getLastRow()+1; row++) addExpirationEvent(sheet, row);
// }

function manageAnswers() {
  /**** Retrieve form information ****/
  let form = FormApp.openByUrl('https://docs.google.com/forms/d/1nQDBpDaQiw5KCdtacIS6Vg3WyXmzmKxLHaG3L6iBWog/edit');
  let formResponses = form.getResponses()[form.getResponses().length-1].getItemResponses();
  let linkToModify = form.getResponses()[form.getResponses().length-1].getEditResponseUrl();  
  let cf = formResponses[INDEX_CODICE_FISCALE].getResponse();
  let email = formResponses[INDEX_EMAIL_CONTATTO].getResponse();
  Logger.log("Esecuzione per %s: Link da aggiungere %s per la mail %s", cf, linkToModify, email);

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let lastModifiedRow = getAddedRow(sheet, cf);
  Logger.log("Ultima riga aggiunta/modificata: %d", lastModifiedRow);

  /**** Link to text for modifying response ****/
  changeLinkToLabel(sheet,lastModifiedRow, LINK_RISPOSTA, TEXT_LINK_RISPOSTA, "modifica risposta", linkToModify);

  /**** Link to text for certificato medico ****/
  changeLinkToLabel(sheet, lastModifiedRow, LINK_CERTIFICATO_MEDICO, TEXT_LINK_CERTIFICATO, "certificato medico");
 
  /**** Link to text for certificato medico aggiornato ****/
  changeLinkToLabel(sheet, lastModifiedRow, LINK_CERTIFICATO_MEDICO_AGGIORNATO, TEXT_LINK_CERTIFICATO, "certificato medico");

  /**** Link to text for foglio privacy ****/
  changeLinkToLabel(sheet, lastModifiedRow, LINK_PRIVACY, TEXT_LINK_FOGLIO_PRIVACY, "foglio privacy");

  /**** Link to text for 1° rata ****/
  changeLinkToLabel(sheet, lastModifiedRow, LINK_RATA_1, TEXT_LINK_PRIMA_RATA, "ricevuta 1° rata");

  /**** Link to text for 1° rata ****/
  changeLinkToLabel(sheet, lastModifiedRow, LINK_RATA_2, TEXT_LINK_SECONDA_RATA, "ricevuta 2° rata");

  /**** Text formatting ****/
  formatField(sheet, lastModifiedRow, EMAIL, LOWECASE);
  formatField(sheet, lastModifiedRow, INDIRIZZO_EMAIL, LOWECASE);
  formatField(sheet, lastModifiedRow, NOME_ATLETA, FISRT_CAPITAL);
  formatField(sheet, lastModifiedRow, COGNOME_ATLETA, FISRT_CAPITAL);
  formatField(sheet, lastModifiedRow, COMUNE_NASCITA, FISRT_CAPITAL);
  formatField(sheet, lastModifiedRow, PROVINCIA_NASCITA, FISRT_CAPITAL);
  formatField(sheet, lastModifiedRow, INDIRIZZO_RESIDENZA, FISRT_CAPITAL);
  formatField(sheet, lastModifiedRow, COMUNE_RESIDENZA, FISRT_CAPITAL);
  formatField(sheet, lastModifiedRow, PROVINCIA_RESIDENZA, FISRT_CAPITAL);
  formatField(sheet, lastModifiedRow, NOME_GENITORE, FISRT_CAPITAL);
  formatField(sheet, lastModifiedRow, COGNOME_GENITORE, FISRT_CAPITAL);

  /**** Add age ****/
  let birthDateCell = getCell(sheet, lastModifiedRow, DATA_NASCITA_ATLETA);  
  let ageCell = getCell(sheet, lastModifiedRow, ETA);
  ageCell.setFormula('=DATEDIF(' + birthDateCell.getA1Notation() + '; TODAY(); "Y")');

  /**** Add expiration event in calendar ****/
  addExpirationEvent(sheet, lastModifiedRow);

  /**** Set payments flags ****/
  setPayments(sheet, lastModifiedRow);

  /**** Send email ****/
  let completeName = getCell(sheet, lastModifiedRow, NOME_ATLETA).getValue() + " " + getCell(sheet, lastModifiedRow, COGNOME_ATLETA).getValue();
  if(sendEmail(ISCRIZIONE_MODIFICA, email, linkToModify, completeName) == true) 
    Logger.log("Email inviata correttamente a %s (%s)", completeName, email);
  else 
    throw new Error("Errore invio email per %s (%s)", completeName, email)
}

/**
 * Function to run manually to fix each time different issues
 */
function fixIssues(){
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let cf=["CRNFBN79E07L407T"];

  cf.forEach(cf=>{
    let lastModifiedRow = getAddedRow(sheet, cf);
    formatField(sheet, lastModifiedRow, EMAIL, LOWECASE);
    formatField(sheet, lastModifiedRow, INDIRIZZO_EMAIL, LOWECASE);
    formatField(sheet, lastModifiedRow, NOME_ATLETA, FISRT_CAPITAL);
    formatField(sheet, lastModifiedRow, COGNOME_ATLETA, FISRT_CAPITAL);
    formatField(sheet, lastModifiedRow, COMUNE_NASCITA, FISRT_CAPITAL);
    formatField(sheet, lastModifiedRow, PROVINCIA_NASCITA, FISRT_CAPITAL);
    formatField(sheet, lastModifiedRow, INDIRIZZO_RESIDENZA, FISRT_CAPITAL);
    formatField(sheet, lastModifiedRow, COMUNE_RESIDENZA, FISRT_CAPITAL);
    formatField(sheet, lastModifiedRow, PROVINCIA_RESIDENZA, FISRT_CAPITAL);
    formatField(sheet, lastModifiedRow, NOME_GENITORE, FISRT_CAPITAL);
    formatField(sheet, lastModifiedRow, COGNOME_GENITORE, FISRT_CAPITAL);
  });

}







