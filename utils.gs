
const ISCRIZIONE_MODIFICA = 0;
const REMINDER_PAGAMENTO = 1;
const REMINDER_ISCRIZIONI = 2;
const ISCRIZIONE_COLLEGE = 3;
const REMINDER_SCADENZA_CM = 4;
/**
 * Function to send email
 * @param type Content ()
 * @param receiver The recipient
 * @param data Link to modify the response of the Google Module
 * @param name Name of the recipient
 * @returns True if sending succeeds, false otherwise
 */
function sendEmail(type, receiver, data, name){
  //
  let htmlEmail;
  let subject;
  try{
    switch(type){
      case ISCRIZIONE_MODIFICA:
        htmlEmail = HtmlService.createHtmlOutputFromFile('EmailIscrizione.html').getContent();
        htmlEmail = htmlEmail.replaceAll("LINK_TO_MODIFY_RESPONSE", data);
        htmlEmail = htmlEmail.replaceAll("NAME_OF_THE_ATHLETE", name);
        subject = 'Conferma iscrizione o aggiornamento certificato medico';
        break;
      case REMINDER_ISCRIZIONI:
        htmlEmail = HtmlService.createHtmlOutputFromFile('EmailReminderIscrizioni.html').getContent();
        htmlEmail = htmlEmail.replaceAll("ATHLETE_NAME", name);
        subject = name + " è ora di iscriversi!";
        break;
      case ISCRIZIONE_COLLEGE:
        break;
      case REMINDER_PAGAMENTO:
        htmlEmail = HtmlService.createHtmlOutputFromFile('EmailPagamenti.html').getContent();
        htmlEmail = htmlEmail.replaceAll("ATHLETE_NAME", name);
        htmlEmail = htmlEmail.replaceAll("NUMERO_RATA", data['numero_rata']);
        htmlEmail = htmlEmail.replaceAll("LINK_TO_MODIFY_RESPONSE", data['link_risposta']);
        subject = "Notifica pagamenti - " + name;
        break;
      case REMINDER_SCADENZA_CM:
        htmlEmail = HtmlService.createHtmlOutputFromFile('EmailScadenzaCM.html').getContent();
        htmlEmail = htmlEmail.replaceAll('NAME_OF_THE_ATHLETE', name);
        htmlEmail = htmlEmail.replaceAll('DAY_OF_EXPIRATION', data['date']);
        htmlEmail = htmlEmail.replaceAll('LINK_TO_MODIFY_RESPONSE', data['url']);
        subject = 'Scadenza certificato medico ' + name;
        break;
      default:
        Logger.log("Tipo di email non riconosciuto (%d)", type);
        return false;
    }
    let options = {
      htmlBody: htmlEmail,
      name: "Team JKR"
    };
    MailApp.sendEmail(receiver, subject, '', options);
    return true;
  }catch(e){
    Logger.log("Errore mandando la mail: " + e.message);
    return false;
  }
}

/**
 * Function to get the cell with specified row and column
 * @param sheet Google Sheet reference
 * @param row Row of reference within resides the searched cell 
 * @param columnName Column name of the cell
 * @returns Reference (Range) to the cell row and column name provided
 */
function getCell(sheet, row, columnName){
    let headerRow = sheet.getRange("1:1").getValues()[0];
    let column = headerRow.indexOf(columnName)+1;
    let cell = sheet.getRange(row, column);
    return cell;
  }

/**
 * Function to return the row index of the correspondent athlete (from 2 to `getLastRow()`)
 * @param sheet Google Sheet reference
 * @param fiscalCode Tax number of the athlete (unique)
 * @returns The number of the corresponding row, -1 if not found
 */
function getRow(sheet, fiscalCode){
  let lastRow = sheet.getLastRow();
  for(let i=1; i<=lastRow; i++){
    if(getCell(sheet, row, CODICE_FISCALE).getValue() == fiscalCode) return i;
  }
  return -1;
}