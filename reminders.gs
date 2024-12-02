/**
 * Function to reminder the previous athletes to subscribe also for the new year
 * @returns null
 */
function reminderIscrizioni(){
	// Controllo esecuzione: lo script deve essere a settembre
	let thisMonth = new Date().getMonth();
	if(thisMonth != SETTEMBRE){
		Logger.log("Mese d'esecuzione non corretto: lo script viene eseguito a settembre")
		return;
	} 

	// Se dicembre o gennaio esegui lo script
	let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
	
	let headerRow = sheet.getRange("1:1").getValues()[0];

	let nameIndex = headerRow.indexOf(NOME_ATLETA);
	let surnameIndex = headerRow.indexOf(COGNOME_ATLETA);
	let emailIndex = headerRow.indexOf(EMAIL);
	
	let athletes = sheet.getRange("2:" + sheet.getLastRow()).getValues();
	
	athletes.forEach(athlete => {
		let completeName = athlete[nameIndex] + " " + athlete[surnameIndex];
		let name = athlete[nameIndex];
		let email = athlete[emailIndex];

		if(sendEmail(REMINDER_ISCRIZIONI, email, null, name)) Logger.log("Email di notifica reminder iscrizione a %s inviata correttamente", completeName);
		else throw Error("Problemi nell'invio email notifica reminder iscrizione a %s", completeName);
	});
}

/**
 * Function to send two reminders for payments: one before december, the other before the end of January
 * @returns null
 */
function reminderPagamenti(){
	// Controllo esecuzione: lo script deve essere eseguito a dicembre e a gennaio 10 giorni prima della fine del mese
	let thisMonth = new Date().getMonth();
	if(thisMonth != NOVEMBRE && thisMonth != GENNAIO){
		Logger.log("Mese d'esecuzione non corretto: lo script viene eseguito a dicembre e a gennaio")
		return;
	} 

	// Se dicembre o gennaio esegui lo script
	let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	
	let headerRow = sheet.getRange("1:1").getValues()[0];

  let nameIndex = headerRow.indexOf(NOME_ATLETA);
  let surnameIndex = headerRow.indexOf(COGNOME_ATLETA);
  let firstPaymentIndex = headerRow.indexOf(PRIMA_RATA);
  let secondPaymentIndex = headerRow.indexOf(SECONDA_RATA);
  let emailIndex = headerRow.indexOf(EMAIL);

  let athletes = sheet.getRange("2:" + sheet.getLastRow()).getValues();
  
  let row = 2;
  for(athlete of athletes){
    let completeName = athlete[nameIndex] + " " + athlete[surnameIndex];
    let email = athlete[emailIndex];
    let responseLink = getCell(sheet, row, LINK_RISPOSTA).getRichTextValue().getLinkUrl();
    row++;

		// 1: paid, 0: not paid, other: future expansions
    let firstPayment = athlete[firstPaymentIndex];
    let secondPayment = athlete[secondPaymentIndex];

		let firstHalf = new Date().getMonth();                        // letiable to check if it's the first half of the year at execution
		firstHalf = firstHalf >= 8 && firstHalf <= 11 ? true : false;

    // The email is not sent if the payments are done within deadline
    if((firstHalf && firstPayment == 1)||(!firstHalf && firstPayment == 1 && secondPayment == 1)) {
			Logger.log("L'atleta %s ha effetuato tutti i pagamenti (prima rata: %d, seconda rata: %d) entro il %s periodo (%s)", 
        completeName, firstPayment, secondPayment, firstHalf ? "primo" : "secondo", email);
			continue;     // No payment to do for the specific deadline
		} 

		let numeroRata = ""; 
		if(firstPayment == 0) numeroRata += "prima";
		if(!firstHalf && secondPayment == 0) 
			if(numeroRata.length == 0) numeroRata += "seconda";
			else numeroRata += " e seconda";

    let data = {'numero_rata' : numeroRata, 'link_risposta' : responseLink};
		if(sendEmail(REMINDER_PAGAMENTO, email, data, completeName)) Logger.log("Email di notifica pagamento a %s inviata correttamente (prima rata: %d, seconda rata: %d, primo periodo: %s)", completeName, firstPayment, secondPayment, firstHalf);
		else Logger.log("Problemi nell'invio email notifica pagamento a %s", completeName);

		//Logger.log("Atleta: %s Email: %s Prima rata: %d Seconda rata: %d", completeName, email, firstPayment, secondPayment);
	}
}
	
/**
 * Function to calculate the interval of dates given a number of months in advance
 * @param monthInAdvance How many month in advance 
 * @returns Array with [first date, last date] of the thisMonth + monthInAdvance
 */
function getNextDates(monthInAdvance){
  let today = new Date();
  let nextMonth;
  let nextYear = today.getFullYear();
  
  // Set the beginning of the interval 
  nextMonth = (today.getMonth() + monthInAdvance) % 12;
  if(today.getMonth() + monthInAdvance > 11)
    nextYear += 1;
  let startDate = new Date(nextYear, nextMonth, 1);

  // Set the end of the interval (first day of the month after)
  nextMonth = (today.getMonth() + monthInAdvance + 1) % 12;
  nextYear = today.getFullYear();
  if(today.getMonth() + monthInAdvance + 1 > 11)
    nextYear += 1;

  let endDate = new Date(nextYear, nextMonth, 1);
  return [startDate, endDate];
}

/**
 *  Function to check the next certificates that will expire (base on the advance given by constants)
 */
function reminderCertificatoMedico(){
  let calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0];
  let events = [];
  
  // Expirations for under 18 
  let dates = getNextDates(ANTICIPO_AGONISTI_MINORENNI);
  let startDate = dates[0];
  let endDate = dates[1];

  Logger.log("Controllo eventi (minorennni) dal %s al %s", startDate.toDateString(), endDate.toDateString());
  let temp = 0;
	// Add all events of under 18 into the array of the people that are notified
  calendar.getEvents(startDate, endDate).forEach((event) => {
    let age = event.getTitle().match("[0-9]{1,2}");
    if(age === null) return false;
    if(age[0]>11 && age[0]<18) {events.push(event); temp++};	// From 11 to 18 they need to be notified with a lot of advance
  });
  Logger.log("Trovati %d eventi", temp);
  

  // Expirations for over 18
  dates = getNextDates(ANTICIPO_ALTRI_ATLETI);
  startDate = dates[0];
  endDate = dates[1];

  Logger.log("Controllo eventi (maggiorenni) dal %s al %s", startDate.toDateString(), endDate.toDateString());
  temp = 0;
	// Add all events of over 18 into the array of the people that are notified
  calendar.getEvents(startDate, endDate).forEach((event) => {
    let age = event.getTitle().match("[0-9]{1,2}");
    if(age === null) return false;
    if(age[0]<12 || age[0]>17 ) {events.push(event); temp++};
  });
  Logger.log("Trovati %d eventi", temp);


  events.forEach(event => {
    let description = event.getDescription();
    if(description.search("cadenza") == 0){
      Logger.log("Evento non corretto (%s)", event.getTitle());
      return;
    }
    let completeName = description.split("\n")[DESC_COMPLETE_NAME]
    let email = description.split("\n")[DESC_EMAIL]
    let urlToForm = description.split("\n")[DESC_URL]
    let date = LanguageApp.translate(Utilities.formatDate(event.getAllDayEndDate(), 'Europa/Roma', 'dd MMMM yyyy'), 'en', 'it');

    let data = {'date' : date, 'url' : urlToForm};
    if(sendEmail(REMINDER_SCADENZA_CM, email, data, completeName)) 
      Logger.log("Invio email per scadenza a " + completeName + " completato");
    else 
      throw new Error("Errore nell'invio email a " + completeName);
  })

}