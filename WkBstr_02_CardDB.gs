// **********************************************
// function fcnPopulateCardDB()
//
// This function populates the selected player
// Card Pool Database
//
// **********************************************

function fcnPopulateCardDB(ss, Player, BoosterData, PackData, shtPlyrCardDB, shtPlyrCardListEn, shtPlyrCardListFr){
  
  var Status = 'Card DB Populate: Not Started';
  var SetName;
  var CardNum;
  var CardQty;
  var SetNameDB;
  var SetNumDB;
  var SetNumMstr;
  var SetNameMstr;
  var MasterpieceValid;
  var SetCardDB; // Array 300 rows x 2 columns. Rows = Cards 1-300, Columns 0 = Qty, 1 = Card Number
  
  var PlyrCardDBMaxCol = shtPlyrCardDB.getMaxColumns();
  
  // Get Set Name of selected booster
  SetName = BoosterData[0][0];
  
  // Set the Card Set Name in Pack Data
  PackData[0][0] = 'Set Name';
  PackData[0][1] = SetName;
  
  // Loop through Card DB to find the Appropriate Set
  for(var setcol = 1; setcol <= PlyrCardDBMaxCol; setcol++){
    SetNameDB = shtPlyrCardDB.getRange(6,setcol).getValue();
    
    // If Set Name is found
    if(SetName == SetNameDB) {
      
      Status = 'Card DB Populate: Card Set Found';Logger.log(Status);
      
      // Process Masterpiece card if present
      if(BoosterData[15][0] == 'Yes'){
        Status = 'Card DB Populate: Analyzing Masterpiece Validity';Logger.log(Status);
        
        // Get Set Number from Card Database for Masterpiece
        SetNumDB =  shtPlyrCardDB.getRange(4, setcol).getValue();
        
        // If Set Number = Even Number, assign the matching Set Number for Masterpieces
        switch(SetNumDB){
          case 2: SetNumDB = 1; break;
          case 4: SetNumDB = 3; break;
          case 6: SetNumDB = 5; break;
          case 8: SetNumDB = 7; break;
        }
        // Masterpiece Card must be Card 14 in Pack
        CardNum = BoosterData[14][0];
        Logger.log(CardNum);
        if(CardNum != ''){
          // Find Masterpiece Card List with SetNumDB
          for(var Colsetnum = 33; Colsetnum <= PlyrCardDBMaxCol; Colsetnum++){
            // Get Set Number from Masterpiece series
            SetNumMstr = shtPlyrCardDB.getRange(4,Colsetnum).getValue();
            // When Set Number is found for Masterpiece
            Logger.log("%s - %s",SetNumDB,SetNumMstr);
            if(SetNumDB == SetNumMstr){
              Status = 'Card DB Populate: Masterpiece Set Found'; Logger.log(Status);
              // Get Masterpiece Set Name
              SetNameMstr = shtPlyrCardDB.getRange(6,Colsetnum).getValue();
              
              // If Masterpiece Set Name is null, set doesn't have a Masterpiece Series, Reject Pack
              if(SetNameMstr != '' && CardNum <= 54) {
                MasterpieceValid = 1;
                Status = 'Card DB Populate: Masterpiece Card Found';
              }
              
              // If Masterpiece Set Name is null, set doesn't have a Masterpiece Series, Reject Pack
              if(SetNameMstr == '') {
                MasterpieceValid = -1;
                Status = 'Card DB Populate: Set does not have Masterpiece Series';
              }
              
              // If Masterpiece Card Number is greater than 54, Card Number is Invalid, Reject Pack
              if(CardNum > 54) {
                MasterpieceValid = -1;
                Status = 'Card DB Populate: Masterpiece Card Number is not valid';
              }              
              
              // If Masterpiece Set is Valid, Process Masterpiece Card
              if(MasterpieceValid == 1){
                Status = 'Card DB Populate: Processing Masterpiece';
                SetCardDB = shtPlyrCardDB.getRange(7, Colsetnum-2, 60, 4).getValues();
                Logger.log('Card Num DB: %s',SetCardDB[CardNum][1]);
                CardQty = SetCardDB[CardNum][0];
                if(CardQty == '') SetCardDB[CardNum][0] = 0;
                SetCardDB[CardNum][0] += 1;
                Status = 'Card DB Populate: Masterpiece Card Quantity Updated'; Logger.log(Status);
                
                // Store Card Info to return to Main Function
                PackData[14][0] = 14;                    // Card in Pack
                PackData[14][1] = SetCardDB[CardNum][1]; // Card Number
                PackData[14][2] = SetCardDB[CardNum][2]; // Card Name
                PackData[14][3] = SetCardDB[CardNum][3]; // Card Rarity
                
                // If Masterpiece is present, specify it in Pack Data[15]
                PackData[15][2] = 'Masterpiece';

                Logger.log('Card Qty: %s',SetCardDB[CardNum][0]);
                // Update the Card DB for selected Set
                shtPlyrCardDB.getRange(7, Colsetnum-2, 60, 4).setValues(SetCardDB);
                Status = 'Card DB Populate: Masterpiece Card List Updated'; Logger.log(Status);
                Colsetnum = PlyrCardDBMaxCol + 1;
              }
            }
          }
        }
      }  
      
      // If there is no Masterpiece Error, process regular Cards
      if(MasterpieceValid != -1){
        // Get Set Card List where:
        // Col[0] = Qty and Col[1] = Card Number
        // Row[0] = Header and Row[1-284] = Card Number
        SetCardDB = shtPlyrCardDB.getRange(7, setcol-2, 300, 4).getValues();
        
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
            
            // If Masterpiece is not present, specify it in Pack Data[15]
            PackData[15][2] = 'No Masterpiece';
          }
          
          // If Pack Data is null, populate it
          if(PackData[card][0] == ''){
            // Store Card Info to return to Main Function
            PackData[card][0] = card;               // Card in Pack
            PackData[card][1] = SetCardDB[CardNum][1]; // Card Number
            PackData[card][2] = SetCardDB[CardNum][2]; // Card Name
            PackData[card][3] = SetCardDB[CardNum][3]; // Card Rarity
          }
          Status = 'Card DB Populate: Card Quantities Updated';
        }
        
        // Update the Card DB for selected Set
        shtPlyrCardDB.getRange(7, setcol-2, 300, 4).setValues(SetCardDB);
        
        Status = 'Card DB Populate: Complete';  
        
        // Exits the Loop
        setcol = PlyrCardDBMaxCol + 1;
      }
    }
  }
  // Send Status through PackData
  PackData[0][3] = Status;
//  var shtTest = ss.getSheetByName('Test');
//  shtTest.getRange(1,1,16,4).setValues(PackData);
  return PackData; 
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
  
  var Status;
    
  // Clear Player Card Pool
  rngCardListEn.clearContent();
  CardList = rngCardListEn.getValues();
  rngCardListFr.clearContent();
  CardList = rngCardListFr.getValues();
  
  Status = 'Card List Update: Player Card Lists Cleared';
    
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
    Status = 'Card List Update: Player Cards Processed';
  }
  // Updates the Player Card Pool
  rngCardListEn.setValues(CardList);
  shtPlyrCardListEn.getRange(3,1).setValue(CardTotal);
  rngCardListFr.setValues(CardList);
  shtPlyrCardListFr.getRange(3,1).setValue(CardTotal);
  
  Status = 'Card List Update: Complete';
  
  return Status;
}