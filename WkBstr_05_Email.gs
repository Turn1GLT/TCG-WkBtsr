// **********************************************
// function subGetEmailAddress()
//
// This function gets the email addresses from 
// the configuration file
//
// **********************************************

function subGetEmailAddress(Player, shtPlayers, EmailAddresses){
  
  // Players Sheets for Email addresses
  var colEmail = 3;
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var rowPlayer = 0;
  var PlyrRowStart = 3;
  
  var PlayerNames = shtPlayers.getRange(PlyrRowStart,2,NbPlayers,1).getValues();
  
  // Find Players rows
  for (var row = 0; row < NbPlayers; row++){
    if (PlayerNames[row] == Player) rowPlayer = row + PlyrRowStart;
    if (rowPlayer > 0) row = NbPlayers + 1;
  }
  
  // Get Email addresses using the players rows
  EmailAddresses[0] = shtPlayers.getRange(rowPlayer,colEmail+1).getValue();
  EmailAddresses[1] = shtPlayers.getRange(rowPlayer,colEmail).getValue();
    
  return EmailAddresses;
}

// **********************************************
// function fcnSendConfirmEmailEN()
//
// This function generates the confirmation email in English
// after a match report has been submitted
//
// **********************************************

function fcnSendCnfrmEmailEN(Player, Week, EmailAddresses, PackData, shtConfig) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  var Address;
  var Language;
  
  // Open Email Templates
  var ssEmailID = shtConfig.getRange(47,2).getValue();  
  var Headers = SpreadsheetApp.openById(ssEmailID).getSheetByName('Templates').getRange(12,2,20,1).getValues();
  
  // Get Document URLs
  var UrlValues = shtConfig.getRange(17,2,3,1).getValues();
  var urlStandings = UrlValues[0][0];
  var urlCardPool = UrlValues[1][0];
  var urlMatchReporter = UrlValues[2][0];
 
  // League Name
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeEN = shtConfig.getRange(13,2).getValue();
  var LeagueNameEN = Location + ' ' + LeagueTypeEN;
  
  // Player Email and Language Preference
  Language = EmailAddresses[0];
  Address  = EmailAddresses[1];
 
  // Set Email Subject
  EmailSubject = LeagueNameEN + " - Week " + Week + " - Weekly Booster" ;
    
  // Start of Email Message
  EmailMessage = '<html><body>';
  
  EmailMessage += 'Hi ' + Player + ',<br><br>You have succesfully added a Booster to your Card Pool for the ' + LeagueNameEN + ', Week ' + Week + '.' +
    '<br><br>Here is the list of cards added to your pool.';
  
  // Builds the Pack Table
  EmailMessage = subMatchReportTable(EmailMessage, Headers, PackData, Language, 1);
     
  EmailMessage += "<br><br>Click below to access your Card Pool."+
        "<br>"+ urlCardPool +
          "<br><br>Thank you for using TCG Booster League Manager from Turn 1 Gaming Leagues & Tournaments";
  
  // End of Email Message
  EmailMessage += '</body></html>';
  
  // Send Email to player
  MailApp.sendEmail(Address, EmailSubject, "",{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
}

// **********************************************
// function fcnSendConfirmEmailFR()
//
// This function generates the confirmation email in English
// after a match report has been submitted
//
// **********************************************

function fcnSendCnfrmEmailFR(Player, Week, EmailAddresses, PackData, shtConfig) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  var Address;
  var Language;
  
  // Open Email Templates
  var ssEmailID = shtConfig.getRange(47,2).getValue();
  var ssEmail = SpreadsheetApp.openById(ssEmailID);
  var shtEmailTemplates = ssEmail.getSheetByName('Templates');
  //var Headers = shtEmailTemplates.getRange(3,3,29,1).getValues();
  
  var Headers = SpreadsheetApp.openById(ssEmailID).getSheetByName('Templates').getRange(12,3,20,1).getValues();
  
  // Get Document URLs
  var UrlValues = shtConfig.getRange(17,2,3,1).getValues();
  var urlStandings = UrlValues[0][0];
  var urlCardPool = UrlValues[1][0];
  var urlMatchReporter = UrlValues[2][0];
 
  // League Name
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeFR = shtConfig.getRange(14,2).getValue();
  var LeagueNameFR = LeagueTypeFR + ' ' + Location;
  
  // Player Email and Language Preference
  Language = EmailAddresses[0];
  Address  = EmailAddresses[1];
 
  // Set Email Subject
  EmailSubject = LeagueNameFR + " - Semaine " + Week + " - Booster de Semaine" ;
    
  // Start of Email Message
  EmailMessage = '<html><body>';
  
  EmailMessage += 'Bonjour ' + Player + ',<br><br>Vous avez ajouté avec succès un booster à votre Pool de Cartes pour la ' + LeagueNameFR + ', Semaine ' + Week + '.' +
    '<br><br>Voici la liste des cartes ajoutées à votre pool.';
  
  // Builds the Pack Table
  EmailMessage = subMatchReportTable(EmailMessage, Headers, PackData, Language, 1);
  
  EmailMessage += "<br><br>Cliquez ci-dessous pour accéder à votre Pool de Cartes:"+
        "<br>"+ urlCardPool +
          "<br><br>Merci d'utiliser TCG Booster League Manager de Turn 1 Gaming Leagues & Tournaments";
  
  // End of Email Message
  EmailMessage += '</body></html>';
  
  // Send Email to player
  MailApp.sendEmail(Address, EmailSubject, "",{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
}

// **********************************************
// function subMatchReportTable()
//
// This function generates the HTML table that displays 
// the Match Data and Booster Pack Data
//
// **********************************************

function subMatchReportTable(EmailMessage, Headers, PackData, Language, Param){
  
  var Item = Headers[16][0];
  var CardNumber = Headers[17][0];
  var CardName = Headers[18][0];
  var CardRarity = Headers[19][0];
    
  for(var row=0; row<16; ++row){

    // Translate MatchData if necessary
    if (Language == 'English' && PackData[row][0] == 'Oui') PackData[row][0] = 'Yes';
    if (Language == 'English' && PackData[row][0] == 'Non') PackData[row][0] = 'No' ;
    if (Language == 'Français' && PackData[row][0] == 'Yes') PackData[row][0] = 'Oui';
    if (Language == 'Français' && PackData[row][0] == 'No' ) PackData[row][0] = 'Non';
    
    // Start of Pack Table
    if(row == 0 && Param == 1) {
      // English
      if(Language == 'English') EmailMessage += '<br><br><font size="4"><b>'+'Set: '+PackData[row][1]+'<br>';
      
      // French
      if(Language == 'Français') EmailMessage += '<br><br><font size="4"><b>'+'Set: '+PackData[row][1]+'<br>';

      EmailMessage += '</b></font><br><table style="border-collapse:collapse;" border = 1 cellpadding = 5><th>'+Item+'</th><th>'+CardNumber+'</th><th>'+CardName+'</th><th>'+CardRarity+'</th>';
    }
    
    // Pack Data for the first 14 Cards 
    if(row > 0 && row < 15 && Param == 1) {
      EmailMessage += '<tr><td>'+Headers[row][0]+'</td><td><center>'+PackData[row][1]+'</td><td>'+PackData[row][2]+'</td><td><center>'+PackData[row][3]+'</td></tr>';
    }
      
    // Pack Data for the first 14 Cards 
    if(row == 15 && PackData[15][2] == 'Masterpiece' && Param == 1) {
      
      // Masterpiece Card Information according to language
      if(PackData[15][2] == 'Masterpiece' && Language == 'English')  PackData[15][2] = "Card 14 is a Masterpiece";
      if(PackData[15][2] == 'Masterpiece' && Language == 'Français') PackData[15][2] = "Carte 14 est une Masterpiece";
      
      // Adds Masterpiece Info to Table if Masterpiece is present
      EmailMessage += '<tr><td>'+Headers[row][0]+'</td><td><center>'+PackData[row][1]+'</td><td>'+PackData[row][2]+'</td><td><center>'+PackData[row][3]+'</td></tr>';
    }
    
  }
  return EmailMessage +'</table>';
}