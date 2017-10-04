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
  var UpdtListStatus;
  var ErrorMsg
  var ConfirmationMsg;
  
  // Email Addresses Array
  var EmailAddresses = new Array(2); // 0= Language Preference, 1= email address
    
  // UI Variables
  var ui = SpreadsheetApp.getUi();
  
  // Create Array of 16x4 where each row is Card 1-14 and each column is Card Info
  var PackData = new Array(16); // 0 = Set Name, 1-14 = Card Numbers, 15 = Card 14 is Masterpiece (Y-N)
  for(var cardnum = 0; cardnum < 16; cardnum++){
    PackData[cardnum] = new Array(4); // 0= Card in Pack, 1= Card Number, 2= Card Name, 3= Card Rarity
    for (var val = 0; val < 4; val++) PackData[cardnum][val] = '';
  }
  
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
    PackData = fcnPopulateCardDB(ss, Player, BoosterData, PackData, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr);
    // Function Status is in PackData[0][3]
    PopulateStatus = PackData[0][3];
    Logger.log(PopulateStatus);
    
    if(PopulateStatus == 'Card DB Populate: Complete') {
      
      // Send Email to Confirm
      EmailAddresses = subGetEmailAddress(Player, shtPlayers, EmailAddresses);
      
      if(EmailAddresses[0] == 'English')  fcnSendCnfrmEmailEN(Player, WeekNum, EmailAddresses, PackData, shtConfig);
      if(EmailAddresses[0] == 'Français') fcnSendCnfrmEmailFR(Player, WeekNum, EmailAddresses, PackData, shtConfig);
      
      // OPENS PROMPT TO NAME NEW DECK AND RENAME INSERTED TAB
      ConfirmationMsg = "Thank you, This booster was added to " + Player + "'s Card Pool."
      ui.alert("SUCCESS",ConfirmationMsg,ui.ButtonSet.OK);
      
      // Update Card Pool Lists
      UpdtListStatus = fcnUpdateCardList(Player, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr);
      Logger.log(UpdtListStatus);
      
      // Clear All Info on OK
//      shtWeekBstr.getRange(1, 2, 21, 1).clearContent();
      
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