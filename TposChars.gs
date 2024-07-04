/*function testid() {
  Logger.log(SpreadsheetApp.getActive().getId())
}*/

function savechar() {
  //Grabbing active sheet info and DB ID from Config Tab
  var spreadsheet = SpreadsheetApp.getActive();
  var spreadsheetid = SpreadsheetApp.getActive().getId();
  var dbsheet = SpreadsheetApp.openById(spreadsheet.getSheetByName("Config").getRange('B3').getValue());
  
  //Popup check to verify if they want to continue with save
  if(Browser.msgBox("Are you sure you want to save?\\nThis may overwrite previously saved data!",Browser.Buttons.OK_CANCEL)=="cancel")return;

  //Most of below is now grabbing data ranges from the active sheet and saving them to array variables. <variable>.push are adding info to an existing array
  var charname = spreadsheet.getSheetByName("Leveler").getRange('F5').getValue();
  var rolledstats = spreadsheet.getSheetByName("Leveler").getRange('A2:B9').getValues();
  var charinfo = spreadsheet.getSheetByName("Leveler").getRange('E4:F10').getValues();
  var charrace = spreadsheet.getSheetByName("Leveler").getRange('G1:H2').getValues();
  charrace.push(["Half Race Check",spreadsheet.getSheetByName("Leveler").getRange('F2').getValue()]);
  
  var charclass = spreadsheet.getSheetByName("Leveler").getRange('J1:K7').getValues();
  charclass.push(["Half Class Check",spreadsheet.getSheetByName("Leveler").getRange('I2').getValue()]);
  
  var charotherinfo = spreadsheet.getSheetByName("Leveler").getRange('E12:F14').getValues();
  charotherinfo.push(spreadsheet.getSheetByName("Leveler").getRange('G12:H14').getValues()[0]);
  charotherinfo.push(spreadsheet.getSheetByName("Leveler").getRange('G12:H14').getValues()[1]);
  charotherinfo.push(spreadsheet.getSheetByName("Leveler").getRange('G12:H14').getValues()[2]);
  charotherinfo.push(spreadsheet.getSheetByName("Leveler").getRange('J9:K9').getValues()[0]);
  charotherinfo.push(spreadsheet.getSheetByName("Leveler").getRange('H4:I4').getValues()[0]);
  charotherinfo.push(spreadsheet.getSheetByName("Leveler").getRange('H6:I6').getValues()[0]);
  charotherinfo.push(spreadsheet.getSheetByName("Leveler").getRange('H22:I22').getValues()[0]);
  charotherinfo.push(spreadsheet.getSheetByName("Leveler").getRange('E15:F15').getValues()[0]);
  
  var charpoints = [["Level",spreadsheet.getSheetByName("Leveler").getRange('B13').getValue(),'']];
  charpoints.push(["Charisma",spreadsheet.getSheetByName("Leveler").getRange('E18').getValue(),spreadsheet.getSheetByName("Leveler").getRange('E19').getValue()]);
  charpoints.push(["Wisdom",spreadsheet.getSheetByName("Leveler").getRange('F18').getValue(),spreadsheet.getSheetByName("Leveler").getRange('F19').getValue()]);
  charpoints.push(["Constitution",spreadsheet.getSheetByName("Leveler").getRange('G18').getValue(),spreadsheet.getSheetByName("Leveler").getRange('G19').getValue()]);
  charpoints.push(["Evasion",spreadsheet.getSheetByName("Leveler").getRange('H18').getValue(),spreadsheet.getSheetByName("Leveler").getRange('H19').getValue()]);
  charpoints.push(["Intimidation",spreadsheet.getSheetByName("Leveler").getRange('I18').getValue(),spreadsheet.getSheetByName("Leveler").getRange('I19').getValue()]);
  charpoints.push(["Mobility",spreadsheet.getSheetByName("Leveler").getRange('J18').getValue(),spreadsheet.getSheetByName("Leveler").getRange('J19').getValue()]);

  var equipadds = spreadsheet.getSheetByName("Leveler").getRange('A28:I48').getValues();
  var tomesinfo = spreadsheet.getSheetByName("Leveler").getRange('A53:C75').getValues();
  var customspellslist = "";
  var custrow=0;
 
  var customspellslist = spreadsheet.getSheetByName("LearnedSpellsNatPowers").getRange(('A4:A800')).getValues();
  var selectedspells = spreadsheet.getSheetByName("SpellBook").getRange('J3:J50').getValues();
  var backpacksize = spreadsheet.getSheetByName("Backpack").getRange('B8').getValue();
  var backpackinv1 = spreadsheet.getSheetByName("Backpack").getRange('A10:I40').getValues();
  var backpackinv2 = spreadsheet.getSheetByName("Backpack").getRange('A41:I70').getValues();
  var backpackinv3 = spreadsheet.getSheetByName("Backpack").getRange('A71:I100').getValues();
  var backpackinv4 = spreadsheet.getSheetByName("Backpack").getRange('A101:I130').getValues();


  var equipeditems = spreadsheet.getSheetByName("Character").getRange('B35:C55').getValues();
  var subclassinfo = [["Blacksmith",spreadsheet.getSheetByName("Character").getRange('C58').getValue(),spreadsheet.getSheetByName("Character").getRange('C59').getValue(),'']];
  subclassinfo.push(["Enchanter",spreadsheet.getSheetByName("Character").getRange('C61').getValue(),spreadsheet.getSheetByName("Character").getRange('C62').getValue(),'']);
  subclassinfo.push(["Alchemist",spreadsheet.getSheetByName("Character").getRange('C64').getValue(),spreadsheet.getSheetByName("Character").getRange('C65').getValue(),'']);
  subclassinfo.push(["Chef",spreadsheet.getSheetByName("Character").getRange('C67').getValue(),spreadsheet.getSheetByName("Character").getRange('C68').getValue(),'']);
  subclassinfo.push(["Spira Guard",spreadsheet.getSheetByName("Character").getRange('C70').getValue(),spreadsheet.getSheetByName("Character").getRange('C71').getValue(),spreadsheet.getSheetByName("Character").getRange('C72').getValue()]);
  subclassinfo.push(["Breeder",spreadsheet.getSheetByName("Character").getRange('E58').getValue(),spreadsheet.getSheetByName("Character").getRange('E59').getValue(),'']);
  subclassinfo.push(["Crafter",spreadsheet.getSheetByName("Character").getRange('E61').getValue(),spreadsheet.getSheetByName("Character").getRange('E62').getValue(),'']);
  subclassinfo.push(["Farmer",spreadsheet.getSheetByName("Character").getRange('E64').getValue(),spreadsheet.getSheetByName("Character").getRange('E65').getValue(),'']);
  subclassinfo.push(["Miner",spreadsheet.getSheetByName("Character").getRange('E67').getValue(),spreadsheet.getSheetByName("Character").getRange('E68').getValue(),'']);
  subclassinfo.push(["Fisherman",spreadsheet.getSheetByName("Character").getRange('E70').getValue(),spreadsheet.getSheetByName("Character").getRange('E71').getValue(),'']);

  var piggyinv = spreadsheet.getSheetByName("PiggyBank").getRange('A6:B11').getValues();
  var backstoryinfo = spreadsheet.getSheetByName("BackStory").getRange('B22').getValue();
  var achievtracker = spreadsheet.getSheetByName("Achievements").getRange(('F2:G'+(spreadsheet.getSheetByName("Achievements").getRange('A1:A').getValues().filter(String).length))).getValues();
  var temprow=0;
  if(spreadsheet.getSheetByName("Quest_Tracker").getRange('A1:A').getValues().filter(String).length==1){
    temprow = 2;
  } else {temprow = spreadsheet.getSheetByName("Quest_Tracker").getRange('A1:A').getValues().filter(String).length;}
  var questtracker1 = spreadsheet.getSheetByName("Quest_Tracker").getRange(('A2:A'+(temprow))).getValues();
  var questtracker2 = spreadsheet.getSheetByName("Quest_Tracker").getRange(('F2:I'+(temprow))).getValues();
  var selectedtitles = spreadsheet.getSheetByName("Character").getRange('C77:C135').getValues();
  temprow = 0;
  if(spreadsheet.getSheetByName("Locations_Visited").getRange('A1:A').getValues().filter(String).length==1){
    temprow = 2;
  } else {temprow = spreadsheet.getSheetByName("Locations_Visited").getRange('A1:A').getValues().filter(String).length;}
  var locationsvisited = spreadsheet.getSheetByName("Locations_Visited").getRange(('A2:A'+(temprow))).getValues();

  // This Section we are checking to verify if there is already a Character in the DB or not. selectedrow would be the row number to save data too whether that be a new line or existing line.
  var selectedrow = 1;
  var rowswithdata = (dbsheet.getSheetByName("DB").getRange('A1:A').getValues().filter(String).length);
  var savecheck = dbsheet.getSheetByName("DB").getRange('A:A').getValues();
  for(var rowsy = 0;rowswithdata>=rowsy;rowsy++){
    if(savecheck[rowsy]==serialize(charname)){
      selectedrow = rowsy+1;
      break;
    } else selectedrow = rowswithdata+1;
  }

  //Exporting all the arrays to DB sheet. Sends arrays to serialize function to change to json serial format
  dbsheet.getSheetByName('DB').getRange(('A'+selectedrow)).setValue(serialize(charname));
  dbsheet.getSheetByName('DB').getRange(('B'+selectedrow)).setValue(spreadsheetid);
  dbsheet.getSheetByName('DB').getRange(('C'+selectedrow)).setValue(serialize(rolledstats));
  dbsheet.getSheetByName('DB').getRange(('D'+selectedrow)).setValue(serialize(charinfo));
  dbsheet.getSheetByName('DB').getRange(('E'+selectedrow)).setValue(serialize(charrace));
  dbsheet.getSheetByName('DB').getRange(('F'+selectedrow)).setValue(serialize(charclass));
  dbsheet.getSheetByName('DB').getRange(('G'+selectedrow)).setValue(serialize(charotherinfo));
  dbsheet.getSheetByName('DB').getRange(('H'+selectedrow)).setValue(serialize(charpoints));
  dbsheet.getSheetByName('DB').getRange(('I'+selectedrow)).setValue(serialize(equipadds));
  dbsheet.getSheetByName('DB').getRange(('J'+selectedrow)).setValue(serialize(tomesinfo));
  dbsheet.getSheetByName('DB').getRange(('K'+selectedrow)).setValue(serialize(customspellslist));
  dbsheet.getSheetByName('DB').getRange(('L'+selectedrow)).setValue(serialize(custrow));
  dbsheet.getSheetByName('DB').getRange(('M'+selectedrow)).setValue(serialize(selectedspells));
  dbsheet.getSheetByName('DB').getRange(('N'+selectedrow)).setValue(serialize(backpacksize));
  dbsheet.getSheetByName('DB').getRange(('O'+selectedrow)).setValue(serialize(backpackinv1));
  dbsheet.getSheetByName('DB').getRange(('P'+selectedrow)).setValue(serialize(backpackinv2));
  dbsheet.getSheetByName('DB').getRange(('Q'+selectedrow)).setValue(serialize(backpackinv3));
  dbsheet.getSheetByName('DB').getRange(('R'+selectedrow)).setValue(serialize(backpackinv4));
  dbsheet.getSheetByName('DB').getRange(('S'+selectedrow)).setValue(serialize(equipeditems));
  dbsheet.getSheetByName('DB').getRange(('T'+selectedrow)).setValue(serialize(subclassinfo));
  dbsheet.getSheetByName('DB').getRange(('U'+selectedrow)).setValue(serialize(piggyinv));
  dbsheet.getSheetByName('DB').getRange(('V'+selectedrow)).setValue(serialize(backstoryinfo));
  dbsheet.getSheetByName('DB').getRange(('W'+selectedrow)).setValue(serialize(achievtracker));
  dbsheet.getSheetByName('DB').getRange(('X'+selectedrow)).setValue(serialize(questtracker1));
  dbsheet.getSheetByName('DB').getRange(('Y'+selectedrow)).setValue(serialize(questtracker2));
  dbsheet.getSheetByName('DB').getRange(('Z'+selectedrow)).setValue(serialize(selectedtitles));
  dbsheet.getSheetByName('DB').getRange(('AA'+selectedrow)).setValue(serialize(locationsvisited));

}

function loadchar(){
  var spreadsheet = SpreadsheetApp.getActive();
  var dbsheet = SpreadsheetApp.openById(spreadsheet.getSheetByName("Leveler").getRange('J12').getValue());

 if(Browser.msgBox("Are you sure you want to load?\\nThis may overwrite current character!",Browser.Buttons.OK_CANCEL)=="cancel")return;


  var charname = deserialize(spreadsheet.getSheetByName("Leveler").getRange('J13').getValue())
  var selectedrow = 1;
  var rowswithdata = (dbsheet.getSheetByName("DB").getRange('A1:A').getValues().filter(String).length);
  var savecheck = dbsheet.getSheetByName("DB").getRange('A:A').getValues();
  
  //Looking for character based on selected name
  for(var rowsy = 0;rowswithdata>=rowsy;rowsy++){
    if(savecheck[rowsy]==serialize(charname)){
      selectedrow = rowsy+1;
      break;
    };
  }
  if(selectedrow==1) {
    Browser.msgBox("No Char Found");
    return;    
  }

  //Importing and Deserializing selected character from DB back into arrays
  var rolledstats = deserialize(dbsheet.getSheetByName('DB').getRange(('C'+selectedrow)).getValue());
  var charinfo = deserialize(dbsheet.getSheetByName('DB').getRange(('D'+selectedrow)).getValue());
  var charrace = deserialize(dbsheet.getSheetByName('DB').getRange(('E'+selectedrow)).getValue());
  var charclass = deserialize(dbsheet.getSheetByName('DB').getRange(('F'+selectedrow)).getValue());
  var charotherinfo = deserialize(dbsheet.getSheetByName('DB').getRange(('G'+selectedrow)).getValue());
  var charpoints = deserialize(dbsheet.getSheetByName('DB').getRange(('H'+selectedrow)).getValue());
  var equipadds = deserialize(dbsheet.getSheetByName('DB').getRange(('I'+selectedrow)).getValue());
  var tomesinfo = deserialize(dbsheet.getSheetByName('DB').getRange(('J'+selectedrow)).getValue());
  var customspellslist = deserialize(dbsheet.getSheetByName('DB').getRange(('K'+selectedrow)).getValue());
  var custrow = deserialize(dbsheet.getSheetByName('DB').getRange(('L'+selectedrow)).getValue());
  var selectedspells = deserialize(dbsheet.getSheetByName('DB').getRange(('M'+selectedrow)).getValue());
  var backpacksize = deserialize(dbsheet.getSheetByName('DB').getRange(('N'+selectedrow)).getValue());
  var backpackinv1 = deserialize(dbsheet.getSheetByName('DB').getRange(('O'+selectedrow)).getValue());
  var backpackinv2 = deserialize(dbsheet.getSheetByName('DB').getRange(('P'+selectedrow)).getValue());
  var backpackinv3 = deserialize(dbsheet.getSheetByName('DB').getRange(('Q'+selectedrow)).getValue());
  var backpackinv4 = deserialize(dbsheet.getSheetByName('DB').getRange(('R'+selectedrow)).getValue());
  var equipeditems = deserialize(dbsheet.getSheetByName('DB').getRange(('S'+selectedrow)).getValue());
  var subclassinfo = deserialize(dbsheet.getSheetByName('DB').getRange(('T'+selectedrow)).getValue());
  var piggyinv = deserialize(dbsheet.getSheetByName('DB').getRange(('U'+selectedrow)).getValue());
  var backstoryinfo = deserialize(dbsheet.getSheetByName('DB').getRange(('V'+selectedrow)).getValue());
  var achievtracker = deserialize(dbsheet.getSheetByName('DB').getRange(('W'+selectedrow)).getValue());
  var questtracker1 = deserialize(dbsheet.getSheetByName('DB').getRange(('X'+selectedrow)).getValue());
  var questtracker2 = deserialize(dbsheet.getSheetByName('DB').getRange(('Y'+selectedrow)).getValue());
  var selectedtitles = deserialize(dbsheet.getSheetByName('DB').getRange(('Z'+selectedrow)).getValue());
  var locationsvisited = deserialize(dbsheet.getSheetByName('DB').getRange(('AA'+selectedrow)).getValue());

  //Going through each array and getting corrisponding variable to set in proper field on sheet.
  rolledstats.forEach(function(value) {
    switch(value[0]){
      case "Attack":
      spreadsheet.getSheetByName("Leveler").getRange('B2').setValue(value[1]);
      break;
      case "Defense":
      spreadsheet.getSheetByName("Leveler").getRange('B3').setValue(value[1]);
      break;
      case "Accuracy":
      spreadsheet.getSheetByName("Leveler").getRange('B4').setValue(value[1]);
      break;
      case "HP":
      spreadsheet.getSheetByName("Leveler").getRange('B5').setValue(value[1]);
      break;
      case "SP":
      spreadsheet.getSheetByName("Leveler").getRange('B6').setValue(value[1]);
      break;
      case "Hunger":
      spreadsheet.getSheetByName("Leveler").getRange('B7').setValue(value[1]);
      break;
      case "Thirst":
      spreadsheet.getSheetByName("Leveler").getRange('B8').setValue(value[1]);
      break;
      case "Energy":
      spreadsheet.getSheetByName("Leveler").getRange('B9').setValue(value[1]);
      break;
    }
  });
  charinfo.forEach(function(value) {
    switch(value[0].toString()){
      case "Starting Town:":
      spreadsheet.getSheetByName("Leveler").getRange('F4').setValue(value[1]);
      break;
      case "Character Name:":
      spreadsheet.getSheetByName("Leveler").getRange('F5').setValue(value[1]);
      break;
      case "Sex:":
      spreadsheet.getSheetByName("Leveler").getRange('F6').setValue(value[1]);
      break;
      case "Age:":
      spreadsheet.getSheetByName("Leveler").getRange('F7').setValue(value[1]);
      break;
      case "Height:":
      spreadsheet.getSheetByName("Leveler").getRange('F8').setValue(value[1]);
      break;
      case "Weight:":
      spreadsheet.getSheetByName("Leveler").getRange('F9').setValue(value[1]);
      break;
      case "Alignment:":
      spreadsheet.getSheetByName("Leveler").getRange('F10').setValue(value[1]);
      break;
    }
  });
  charrace.forEach(function(value) {
    switch(value[0].toString()){
      case "Race:":
      spreadsheet.getSheetByName("Leveler").getRange('H1').setValue(value[1]);
      break;
      case "Half Race:":
      spreadsheet.getSheetByName("Leveler").getRange('H2').setValue(value[1]);
      break;
      case "Half Race Check":
      spreadsheet.getSheetByName("Leveler").getRange('F2').setValue(value[1]);
      break;
    }
  });
  charclass.forEach(function(value) {
    switch(value[0].toString()){
      case "Class:":
      spreadsheet.getSheetByName("Leveler").getRange('K1').setValue(value[1]);
      break;
      case "Half Class:":
      spreadsheet.getSheetByName("Leveler").getRange('K2').setValue(value[1]);
      break;
      case "Enhanced Class:":
      spreadsheet.getSheetByName("Leveler").getRange('K3').setValue(value[1]);
      break;
      case "Enhanced Half Class:":
      spreadsheet.getSheetByName("Leveler").getRange('K4').setValue(value[1]);
      break;
      case "Select Element1:":
      spreadsheet.getSheetByName("Leveler").getRange('K5').setValue(value[1]);
      break;
      case "Select Element2:":
      spreadsheet.getSheetByName("Leveler").getRange('K6').setValue(value[1]);
      break;
      case "Select Element3:":
      spreadsheet.getSheetByName("Leveler").getRange('K7').setValue(value[1]);
      break;
      case "Half Class Check":
      spreadsheet.getSheetByName("Leveler").getRange('I2').setValue(value[1]);
      break;
    }
  });
  charotherinfo.forEach(function(value) {
    switch(value[0].toString()){
      case "Damage Orb (0/1):":
      spreadsheet.getSheetByName("Leveler").getRange('F12').setValue(value[1]);
      break;
      case "Damage Orb (1/1):":
      spreadsheet.getSheetByName("Leveler").getRange('F12').setValue(value[1]);
      break;
      case "Spell (Counters):":
      spreadsheet.getSheetByName("Leveler").getRange('F13').setValue(value[1]);
      break;
      case "Other Bonuses:":
      spreadsheet.getSheetByName("Leveler").getRange('F14').setValue(value[1]);
      break;
      case "Resistances:":
      spreadsheet.getSheetByName("Leveler").getRange('H12').setValue(value[1]);
      break;
      case "Vulnerabilities:":
      spreadsheet.getSheetByName("Leveler").getRange('H13').setValue(value[1]);
      break;
      case "Loyalty Score:":
      spreadsheet.getSheetByName("Leveler").getRange('H14').setValue(value[1]);
      break;
      case "Time Spell Rank:":
      spreadsheet.getSheetByName("Leveler").getRange('K9').setValue(value[1]);
      break;
      case "Spell Slots Number:":
      spreadsheet.getSheetByName("Leveler").getRange('I4').setValue(value[1]);
      break;
      case "Astrological Sign:":
      spreadsheet.getSheetByName("Leveler").getRange('I6').setValue(value[1]);
      break;
      case "Faction:":
      spreadsheet.getSheetByName("Leveler").getRange('I22').setValue(value[1]);
      break;
      case "Creature Type Override:":
      spreadsheet.getSheetByName("Leveler").getRange('F15').setValue(value[1]);
      break;
    }
  });
  charpoints.forEach(function(value) {
    switch(value[0].toString()){
      case "Level":
      spreadsheet.getSheetByName("Leveler").getRange('B13').setValue(value[1]);
      break;
      case "Charisma":
      spreadsheet.getSheetByName("Leveler").getRange('E18').setValue(value[1]);
      spreadsheet.getSheetByName("Leveler").getRange('E19').setValue(value[2]);
      break;
      case "Wisdom":
      spreadsheet.getSheetByName("Leveler").getRange('F18').setValue(value[1]);
      spreadsheet.getSheetByName("Leveler").getRange('F19').setValue(value[2]);
      break;
      case "Constitution":
      spreadsheet.getSheetByName("Leveler").getRange('G18').setValue(value[1]);
      spreadsheet.getSheetByName("Leveler").getRange('G19').setValue(value[2]);
      break;
      case "Evasion":
      spreadsheet.getSheetByName("Leveler").getRange('H18').setValue(value[1]);
      spreadsheet.getSheetByName("Leveler").getRange('H19').setValue(value[2]);
      break;
      case "Intimidation":
      spreadsheet.getSheetByName("Leveler").getRange('I18').setValue(value[1]);
      spreadsheet.getSheetByName("Leveler").getRange('I19').setValue(value[2]);
      break;
      case "Mobility":
      spreadsheet.getSheetByName("Leveler").getRange('J18').setValue(value[1]);
      spreadsheet.getSheetByName("Leveler").getRange('J19').setValue(value[2]);
      break;
    }
  });
  
  spreadsheet.getSheetByName("Leveler").getRange('A28:I48').setValues(equipadds);
  spreadsheet.getSheetByName("Leveler").getRange('A53:C75').setValues(tomesinfo);
  spreadsheet.getSheetByName("LearnedSpellsNatPowers").getRange('A4:A800').setValues(customspellslist);
  spreadsheet.getSheetByName("Backpack").getRange('B8').setValue(backpacksize);
  spreadsheet.getSheetByName("Backpack").getRange('A10:I40').setValues(backpackinv1);
  spreadsheet.getSheetByName("Backpack").getRange('A41:I70').setValues(backpackinv2);
  spreadsheet.getSheetByName("Backpack").getRange('A71:I100').setValues(backpackinv3);
  spreadsheet.getSheetByName("Backpack").getRange('A101:I130').setValues(backpackinv4);
  spreadsheet.getSheetByName("Character").getRange('B35:C55').setValues(equipeditems);

  subclassinfo.forEach(function(value) {
    switch(value[0].toString()){
      case "Blacksmith":
      spreadsheet.getSheetByName("Character").getRange('C58').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('C59').setValue(value[2]);
      break;
      case "Enchanter":
      spreadsheet.getSheetByName("Character").getRange('C61').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('C62').setValue(value[2]);
      break;
      case "Alchemist":
      spreadsheet.getSheetByName("Character").getRange('C64').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('C65').setValue(value[2]);
      break;
      case "Chef":
      spreadsheet.getSheetByName("Character").getRange('C67').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('C68').setValue(value[2]);
      break;
      case "Spira Guard":
      spreadsheet.getSheetByName("Character").getRange('C70').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('C71').setValue(value[2]);
      spreadsheet.getSheetByName("Character").getRange('C72').setValue(value[3]);
      break;
      case "Breeder":
      spreadsheet.getSheetByName("Character").getRange('E58').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('E59').setValue(value[2]);
      break;
      case "Crafter":
      spreadsheet.getSheetByName("Character").getRange('E61').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('E62').setValue(value[2]);
      break;
      case "Farmer":
      spreadsheet.getSheetByName("Character").getRange('E64').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('E65').setValue(value[2]);
      break;
      case "Miner":
      spreadsheet.getSheetByName("Character").getRange('E67').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('E68').setValue(value[2]);
      break;
      case "Fisherman":
      spreadsheet.getSheetByName("Character").getRange('E70').setValue(value[1]);
      spreadsheet.getSheetByName("Character").getRange('E71').setValue(value[2]);
      break;
    }
  });
  
  spreadsheet.getSheetByName("PiggyBank").getRange('A6:B11').setValues(piggyinv);
  spreadsheet.getSheetByName("BackStory").getRange('B22').setValue(backstoryinfo);
  tempsel = achievtracker.length+1;
  spreadsheet.getSheetByName("Achievements").getRange(('F2:G'+tempsel)).setValues(achievtracker);
  tempsel = questtracker1.length+1;
  spreadsheet.getSheetByName("Quest_Tracker").getRange(('A2:A'+tempsel)).setValues(questtracker1);
  tempsel = questtracker2.length+1;
  spreadsheet.getSheetByName("Quest_Tracker").getRange(('F2:I'+tempsel)).setValues(questtracker2);
  spreadsheet.getSheetByName("SpellBook").getRange('J3:J50').setValues(selectedspells);
  spreadsheet.getSheetByName("Character").getRange('C77:C135').setValues(selectedtitles);
  
  tempsel = locationsvisited.length+1;
  spreadsheet.getSheetByName("Locations_Visited").getRange(('A2:A'+tempsel)).setValues(locationsvisited);

}

function serialize(a) {
  return JSON.stringify(a);
}

function deserialize(a) {
  return JSON.parse(a);
}