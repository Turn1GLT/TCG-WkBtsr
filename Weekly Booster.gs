// **********************************************
// function fcnWeekBstr_Master()
//
// This function adds the Booster to the Player
// Card Pool and checks the Main Weekly Booster Table
//
// **********************************************

function fcnWeekBstr_Master() {

  // Function Sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtWeekBstr = ss.getActiveSheet();
  var shtWeekBstrTable = ss.getSheetByName('Weekly Booster Table');
  var shtPlayers = ss.getSheetByName('Players');
  
  // Main File
  var ssMainId = shtWeekBstr.getRange(4, 1).getValue();
  var shtConfig = SpreadsheetApp.openById(ssMainId).getSheetByName('Config');
  
  // Open Card Pool DB and List Sheets for that Player to send in parameter
  var ssPlyrCardDBId = shtConfig.getRange(31, 2).getValue();
  var ssPlyrCardListEnId = shtConfig.getRange(32, 2).getValue();
  var ssPlyrCardListFrId = shtConfig.getRange(33, 2).getValue();
  var shtPlyrCardDB;
  var shtPlyrCardListEn;
  var shtPlyrCardListFr;
  
  // Function Values
  var WeekNum = shtWeekBstr.getRange(1, 2).getValue();
  var Player = shtWeekBstr.getRange(2, 2).getValue();
  var NbPlayer = shtPlayers.getRange(2, 1).getValue();
  var BoosterData = shtWeekBstr.getRange(6, 2, 16, 1).getValues();
  var PlayerTable = shtWeekBstrTable.getRange(5,2,NbPlayer,1).getValues();
  
  // Function Variables
  var CardMissing = 0;
  var BoosterAllowed = 0;
  var BoosterCheck;
  var RowPlayer;
  var rngBstrCheck;
  var PopulateStatus;
  var ErrorMsg
  var ConfirmationMsg;
  
  // Email Addresses Array
  var EmailAddresses = new Array(2); // 0= Language Preference, 1= email address
    
  // UI Variables
  var ui = SpreadsheetApp.getUi();
  
  // Verify that Booster is allowed for selected week
  
  if(Player != ''){
    // Find Player Row
    for(var i = 0; i < NbPlayer; i++){
      Logger.log(PlayerTable[i][0]);
      if(PlayerTable[i][0] == Player) {
        RowPlayer = i+5;
        Logger.log(RowPlayer);
        i = NbPlayer;
      }
    }
    // Open Player Card Pool DB and Lists
    shtPlyrCardDB = SpreadsheetApp.openById(ssPlyrCardDBId).getSheetByName(Player);
    shtPlyrCardListEn = SpreadsheetApp.openById(ssPlyrCardListEnId).getSheetByName(Player);
    shtPlyrCardListFr = SpreadsheetApp.openById(ssPlyrCardListFrId).getSheetByName(Player);
  }
  
  // Get Weekly Booster Check
  if(WeekNum != '' && Player != ''){
    rngBstrCheck = shtWeekBstrTable.getRange(RowPlayer, WeekNum+2);
    BoosterCheck = rngBstrCheck.getValue();
  
    // If there is no value in the Check Table, the Booster is allowed
    if(BoosterCheck == '') BoosterAllowed = 1;
  }
  
  // Verify all Booster information is present
  for (i = 0; i < 16; i++){
    if(BoosterData[i][0] == '') CardMissing = 1;
  }

  // If All information is present, execute
  if(BoosterAllowed == 1 && WeekNum != '' && Player != '' && CardMissing == 0){
    
    // Add Booster to Card DB and Regenerate Card Pool List
    PopulateStatus = fcnPopulateCardDB(Player, BoosterData, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr);
    
    if(PopulateStatus == 'Card DB Populate: Complete') {
      
      // Send Email to Confirm
      EmailAddresses = subGetEmailAddress(Player, shtPlayers, EmailAddresses);
      
      fcnSendCnfrmEmailEN(Player, WeekNum, EmailAddresses, shtConfig);
      
      // OPENS PROMPT TO NAME NEW DECK AND RENAME INSERTED TAB
      ConfirmationMsg = "Thank you, This booster was added to " + Player + "'s Card Pool."
      ui.alert(ConfirmationMsg);
    
      // Clear All Info on OK
      shtWeekBstr.getRange(1, 2, 21, 1).clearContent();
      
      // Add Check to Weekly Booster Table
      rngBstrCheck.setValue('X');
    }
  }
  
  // If Week Number is missing, display Error Message
  if(WeekNum == '' ){
    ErrorMsg = "Please Select Week Number before Submitting";
    ui.alert("ERROR",ErrorMsg, ui.ButtonSet.OK);
  }

  // If Name is missing, display Error Message
  if(Player == ''){
    ErrorMsg = "Please Select Player Name before Submitting";
    ui.alert("ERROR",ErrorMsg, ui.ButtonSet.OK);
  }

  // If Booster Information is missing, display Error Message
  if(BoosterAllowed == 1 && CardMissing == 1){
    ErrorMsg = "Please Make sure you entered all Booster Information before Submitting";
    ui.alert("ERROR",ErrorMsg, ui.ButtonSet.OK);
  }

  // If Booster was already added, display Error Message
  if(BoosterAllowed == 0 && WeekNum != '' && Player != ''){
    ErrorMsg = "A booster has already been added to " + Player + "'s Card Pool, please select another week";
    ui.alert("ERROR", ErrorMsg, ui.ButtonSet.OK)
  }
  
  // If Populate is not Complete, send Status
  if(BoosterAllowed == 1 && WeekNum != '' && Player != '' && CardMissing == 0 && PopulateStatus != 'Card DB Populate: Complete'){
    ErrorMsg = PopulateStatus;
    ui.alert("ERROR",ErrorMsg, ui.ButtonSet.OK);
  }
    
}


// **********************************************
// function fcnPopulateCardDB()
//
// This function populates the selected player
// Card Pool Database
//
// **********************************************

function fcnPopulateCardDB(Player, BoosterData, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr){
  
  var Status = 'Card DB Populate: Not Started';
  var SetName;
  var CardNum;
  var CardQty;
  var SetNameDB;
  var SetNumDB;
  var SetNumMstr;
  var SetNameMstr;
  var SetCardDB; // Array 300 rows x 2 columns. Rows = Cards 1-300, Columns 0 = Qty, 1 = Card Number
  
  var PlyrCardDBMaxCol = shtPlyrCardDB.getMaxColumns();
  
  // Get Set Name of selected booster
  SetName = BoosterData[0][0];
  Logger.log('Booster Set Name: %s',SetName);
    
  // Loop through Card DB to find the Appropriate Set
  for(var setcol = 1; setcol <= PlyrCardDBMaxCol; setcol++){
    SetNameDB = shtPlyrCardDB.getRange(6,setcol).getValue();
    Logger.log('Set Name DB: %s',SetNameDB);
    
    // If Set Name is found
    if(SetName == SetNameDB) {
      
      Status = 'Card DB Populate: Card Set Found';
      
      // Update Masterpiece card
      if(BoosterData[15][0] == 'Yes'){
        
        // Check if Masterpiece is Valid
        
        Status = 'Card DB Populate: Analyzing Masterpiece Validity';
        // Get Set Number from Card Database for Masterpiece
        SetNumDB =  shtPlyrCardDB.getRange(4, setcol).getValue();
        Logger.log('Set Num DB: %s', SetNumDB)
        // If Set Number = Even Number, assign the matching Set Number for Masterpieces
        switch(SetNumDB){
          case 2: SetNumDB = 1; break;
          case 4: SetNumDB = 3; break;
          case 6: SetNumDB = 5; break;
          case 8: SetNumDB = 7; break;
        }
        
        Logger.log('Set Num DB: %s', SetNumDB)
        
        CardNum = BoosterData[14][0];
        Logger.log('Card Num: %s',CardNum);
        
        if(CardNum != ''){
          // Find Masterpiece Card List with SetNumDB
          for(var Colsetnum = 33; Colsetnum <= PlyrCardDBMaxCol; Colsetnum++){
            // Get Set Number from Masterpiece series
            SetNumMstr = shtPlyrCardDB.getRange(4,Colsetnum).getValue();
            // When Set Number is found for Masterpiece
            if(SetNumDB == SetNumMstr){
              Status = 'Card DB Populate: Masterpiece Set Found';
              // Get Masterpiece Set Name
              SetNameMstr = shtPlyrCardDB.getRange(6,Colsetnum).getValue();
              Logger.log('Set Number Found: %s at column %s',SetNumMstr, Colsetnum);
              Logger.log('Masterpiece Set Name %s',SetNameMstr);
              Logger.log('-----------');
              
              if(SetNameMstr == '' || CardNum > 54) Status = 'Masterpiece Not Valid';
              
              if(SetNameMstr != '' && CardNum <= 54){
                Status = 'Processing Masterpiece';
                SetCardDB = shtPlyrCardDB.getRange(7, Colsetnum-2, 60, 2).getValues();
                Logger.log('Card Num DB: %s',SetCardDB[CardNum][1]);
                CardQty = SetCardDB[CardNum][0];
                if(CardQty == '') SetCardDB[CardNum][0] = 0;
                SetCardDB[CardNum][0] += 1;
                Status = 'Card DB Populate: Masterpiece Card Quantity Updated';
                Logger.log('Card Qty: %s',SetCardDB[CardNum][0]);
                // Update the Card DB for selected Set
                shtPlyrCardDB.getRange(7, Colsetnum-2, 60, 2).setValues(SetCardDB);
                Status = 'Card DB Populate: Masterpiece Card List Updated';
                Colsetnum = PlyrCardDBMaxCol + 1;
              }
            }
          }
        }
      }  
      
      // If there is no Masterpiece Set Name Error, process regular Cards
      if(Status != 'Masterpiece Not Valid'){
        // Get Set Card List where:
        // Col[0] = Qty and Col[1] = Card Number
        // Row[0] = Header and Row[1-284] = Card Number
        SetCardDB = shtPlyrCardDB.getRange(7, setcol-2, 300, 2).getValues();
        
        // Loop through each card to update the quantity for regular cards
        for (var card = 1; card <=14; card++){
          CardNum = BoosterData[card][0];
          // First 13 cards
          if(CardNum != '' && card < 14 && SetCardDB[CardNum][1] == CardNum){
            CardQty = SetCardDB[CardNum][0];
            if(CardQty == '') SetCardDB[CardNum][0] = 0;
            SetCardDB[CardNum][0] += 1;
          }
          // Last card if not Masterpiece
          if(CardNum != '' && card == 14 && SetCardDB[CardNum][1] == CardNum && BoosterData[15][0] != 'Yes'){
            CardQty = SetCardDB[CardNum][0];
            if(CardQty == '') SetCardDB[CardNum][0] = 0;
            SetCardDB[CardNum][0] += 1;
          }
          Status = 'Card DB Populate: Card Quantities Updated';
        }
        
        // Update the Card DB for selected Set
        shtPlyrCardDB.getRange(7, setcol-2, 300, 2).setValues(SetCardDB);
        Status = 'Card DB Populate: Card List Updated';
        
        Status = 'Card DB Populate: Card Database Updated';  
        
        // Update Card Pool Lists
        fcnUpdateCardList(Player, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr);
        
        Status = 'Card DB Populate: Complete';  
        
        // Exits the Loop
        setcol = PlyrCardDBMaxCol + 1;
      }
    }
  }
  return Status; 
}

// **********************************************
// function fcnUpdateCardList()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateCardList(Player, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr){
  
  // Variables
  var CardListEnMaxRows = shtPlyrCardListEn.getMaxRows();
  var CardListFrMaxRows = shtPlyrCardListFr.getMaxRows();
  var rngCardListEn = shtPlyrCardListEn.getRange(6, 1, CardListEnMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var rngCardListFr = shtPlyrCardListFr.getRange(6, 1, CardListFrMaxRows-6, 5); // 0 = Card Qty, 1 = Card Number, 2 = Card Name, 3 = Rarity, 4 = Set Name 
  var CardList; // Where Card Data will be populated
  
  var CardDBSetTotal = shtPlyrCardDB.getRange(2,1,1,48).getValues(); // Gets Sum of all set Quantity, if > 0, set is present in card pool
  var CardTotal = shtPlyrCardDB.getRange(3,7).getValue();
  var SetData;
  var SetName;
  var colSet;
  var CardNb = 0;
    
  // Clear Player Card Pool
  rngCardListEn.clearContent();
  CardList = rngCardListEn.getValues();
  rngCardListFr.clearContent();
  CardList = rngCardListFr.getValues();
    
  // Look for Set with cards present in pool
  for (var col = 0; col <= 48; col++){   
    // if Set Card Quantity > 0, Set has cards in pool, Loop through all cards in Set
    if(CardDBSetTotal[0][col] > 0){
      colSet = col + 1;
      SetName = shtPlyrCardDB.getRange(6,colSet+2).getValue();

      // Get all Cards Data from set
      SetData = shtPlyrCardDB.getRange(7, colSet, 300, 4).getValues();

      // Loop through each card in Set and get Card Data
      for (var CardID = 1; CardID <= 299; CardID++){
        if(SetData[CardID][0] > 0) {
          CardList[CardNb][0] = SetData[CardID][0]; // Quantity
          CardList[CardNb][1] = SetData[CardID][1]; // Card Number (ID)
          CardList[CardNb][2] = SetData[CardID][2]; // Card Name
          CardList[CardNb][3] = SetData[CardID][3]; // Card Rarity
          CardList[CardNb][4] = SetName;            // Set Name    
          CardNb++;
        }
      }
    }
  }
  // Updates the Player Card Pool
  rngCardListEn.setValues(CardList);
  shtPlyrCardListEn.getRange(3,1).setValue(CardTotal);
  rngCardListFr.setValues(CardList);
  shtPlyrCardListFr.getRange(3,1).setValue(CardTotal);
  
  
  // Return Value
}

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

function fcnSendCnfrmEmailEN(Player, Week, EmailAddresses, shtConfig) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  var Address;
  var Language;
  
  // Get Document URLs
  var UrlValues = shtConfig.getRange(17,2,3,1).getValues();
  var urlStandings = UrlValues[0][0];
  var urlCardPool = UrlValues[1][0];
  var urlMatchReporter = UrlValues[2][0];
 
  // League Name
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeEN = shtConfig.getRange(13,2).getValue();
  var LeagueNameEN = Location + ' ' + LeagueTypeEN;
 
  // Set Email Subject
  EmailSubject = LeagueNameEN + " - Week " + Week + " - Weekly Booster" ;
    
  // Start of Email Message
  EmailMessage = '<html><body>';
  
  EmailMessage += 'Hi ' + Player + ',<br><br>You have succesfully added a Booster to your Card Pool for the ' + LeagueNameEN + ', Week ' + Week + '.';
     
  EmailMessage += "<br><br>Click below to access your Card Pool:"+
        "<br>"+ urlCardPool +
          "<br><br>Thank you for using TCG Booster League Manager from Turn 1 Gaming Leagues & Tournaments";
  
  // End of Email Message
  EmailMessage += '</body></html>';
  
  // Sends email to both players with the Match Data
  Language = EmailAddresses[0];
  Address  = EmailAddresses[1];
  
  if(Language == 'English'){
    MailApp.sendEmail(Address, EmailSubject, "",{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
  }
}