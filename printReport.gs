function thisIsTheShit(){
  
  var important_sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Links');
  
  var important_sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Play Console - EN');
  var important_sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Play Console - SP');
  var important_sheet4 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Play Console - FR');
  var important_sheet5 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Play Console - IT');
  var important_sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Play Console - RU');
  var important_sheet7 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Play Console - CN_S');
  var important_sheet8 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Play Console - PT');
  var important_sheet9 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Google Play Console - DE');
  
  var important_sheet10 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('App Store Connect - EN');
  var important_sheet11 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('App Store Connect - SP');
  var important_sheet12 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('App Store Connect - FR');
  var important_sheet13 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('App Store Connect - IT');
  var important_sheet14 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('App Store Connect - RU');
  var important_sheet15 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('App Store Connect - CN_S');
  var important_sheet16 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('App Store Connect - PT');
  var important_sheet17 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('App Store Connect - DE');  
  
  var dataRange_important_sheet1 = important_sheet1.getDataRange();
  var values_important_sheet1 = dataRange_important_sheet1.getValues();
  for ( i = 2 ; i < values_important_sheet1.length +1 ; i++) {
    for ( j = 2 ; j <= 7 ; j++) {
      important_sheet1.getRange(i,j).setValue("");
      important_sheet1.getRange(i,j).setBackground("white");
    }
  }
  important_sheet1.getRange(2,13).setValue("");
  
  
  var dataRange_important_sheet2 = important_sheet2.getDataRange();
  var values_important_sheet2 = dataRange_important_sheet2.getValues();
  var dataRange_important_sheet3 = important_sheet3.getDataRange();
  var values_important_sheet3 = dataRange_important_sheet3.getValues();
  var dataRange_important_sheet4 = important_sheet4.getDataRange();
  var values_important_sheet4 = dataRange_important_sheet4.getValues();
  var dataRange_important_sheet5 = important_sheet5.getDataRange();
  var values_important_sheet5 = dataRange_important_sheet5.getValues();
  var dataRange_important_sheet6 = important_sheet6.getDataRange();
  var values_important_sheet6 = dataRange_important_sheet6.getValues();
  var dataRange_important_sheet7 = important_sheet7.getDataRange();
  var values_important_sheet7 = dataRange_important_sheet7.getValues();
  var dataRange_important_sheet8 = important_sheet8.getDataRange();
  var values_important_sheet8 = dataRange_important_sheet8.getValues();
  var dataRange_important_sheet9 = important_sheet9.getDataRange();
  var values_important_sheet9 = dataRange_important_sheet9.getValues();
  
  var dataRange_important_sheet10 = important_sheet10.getDataRange();
  var values_important_sheet10 = dataRange_important_sheet10.getValues();
  var dataRange_important_sheet11 = important_sheet11.getDataRange();
  var values_important_sheet11 = dataRange_important_sheet11.getValues();
  var dataRange_important_sheet12 = important_sheet12.getDataRange();
  var values_important_sheet12 = dataRange_important_sheet12.getValues();
  var dataRange_important_sheet13 = important_sheet13.getDataRange();
  var values_important_sheet13 = dataRange_important_sheet13.getValues();
  var dataRange_important_sheet14 = important_sheet14.getDataRange();
  var values_important_sheet14 = dataRange_important_sheet14.getValues();
  var dataRange_important_sheet15 = important_sheet15.getDataRange();
  var values_important_sheet15 = dataRange_important_sheet15.getValues();
  var dataRange_important_sheet16 = important_sheet16.getDataRange();
  var values_important_sheet16 = dataRange_important_sheet16.getValues();
  var dataRange_important_sheet17 = important_sheet17.getDataRange();
  var values_important_sheet17 = dataRange_important_sheet17.getValues();  
  
  
  var lastRow1ST_GAME_EN = lastRow(important_sheet2,"A");
  var lastRow2ND_GAME_EN	= lastRow(important_sheet2,"D");
  var lastRow3RD_GAME_EN = lastRow(important_sheet2,"G");
  var lastRow4TH_GAME_EN	= lastRow(important_sheet2,"J");
  var lastRow5TH_GAME_EN = lastRow(important_sheet2,"M");
  var lastRow6TH_GAME_EN	= lastRow(important_sheet2,"P");
  
  var lastRow1ST_GAME_EN_IOS_3 = lastRow(important_sheet10,"A");
  var lastRow1ST_GAME_EN_IOS_2 = lastRow(important_sheet10,"B");
  var lastRow1ST_GAME_EN_IOS_1 = lastRow(important_sheet10,"C");
  
  var lastRow2ND_GAME_EN_IOS_3 = lastRow(important_sheet10,"H");
  var lastRow2ND_GAME_EN_IOS_2 = lastRow(important_sheet10,"I");
  var lastRow2ND_GAME_EN_IOS_1 = lastRow(important_sheet10,"J");
  
  var lastRow3RD_GAME_EN_IOS_3 = lastRow(important_sheet10,"O");
  var lastRow3RD_GAME_EN_IOS_2 = lastRow(important_sheet10,"P");
  var lastRow3RD_GAME_EN_IOS_1 = lastRow(important_sheet10,"Q");
  
  var lastRow4TH_GAME_EN_IOS_3 = lastRow(important_sheet10,"V");
  var lastRow4TH_GAME_EN_IOS_2 = lastRow(important_sheet10,"W");
  var lastRow4TH_GAME_EN_IOS_1 = lastRow(important_sheet10,"X");
  
  var lastRow5TH_GAME_EN_IOS_3 = lastRow(important_sheet10,"AC");
  var lastRow5TH_GAME_EN_IOS_2 = lastRow(important_sheet10,"AD");
  var lastRow5TH_GAME_EN_IOS_1 = lastRow(important_sheet10,"AE");
  
  var lastRow6TH_GAME_EN_IOS_3 = lastRow(important_sheet10,"AJ");
  var lastRow6TH_GAME_EN_IOS_2 = lastRow(important_sheet10,"AK");
  var lastRow6TH_GAME_EN_IOS_1 = lastRow(important_sheet10,"AL");
  
  
  var lastRow1ST_GAME_SP = lastRow(important_sheet3,"A");
  var lastRow2ND_GAME_SP	= lastRow(important_sheet3,"D");
  var lastRow3RD_GAME_SP = lastRow(important_sheet3,"G");
  var lastRow4TH_GAME_SP	= lastRow(important_sheet3,"J");
  var lastRow5TH_GAME_SP = lastRow(important_sheet3,"M");
  var lastRow6TH_GAME_SP	= lastRow(important_sheet3,"P");
  
  var lastRow1ST_GAME_SP_IOS_3 = lastRow(important_sheet11,"A");
  var lastRow1ST_GAME_SP_IOS_2 = lastRow(important_sheet11,"B");
  var lastRow1ST_GAME_SP_IOS_1 = lastRow(important_sheet11,"C");
  
  var lastRow2ND_GAME_SP_IOS_3 = lastRow(important_sheet11,"H");
  var lastRow2ND_GAME_SP_IOS_2 = lastRow(important_sheet11,"I");
  var lastRow2ND_GAME_SP_IOS_1 = lastRow(important_sheet11,"J");
  
  var lastRow3RD_GAME_SP_IOS_3 = lastRow(important_sheet11,"O");
  var lastRow3RD_GAME_SP_IOS_2 = lastRow(important_sheet11,"P");
  var lastRow3RD_GAME_SP_IOS_1 = lastRow(important_sheet11,"Q");
  
  var lastRow4TH_GAME_SP_IOS_3 = lastRow(important_sheet11,"V");
  var lastRow4TH_GAME_SP_IOS_2 = lastRow(important_sheet11,"W");
  var lastRow4TH_GAME_SP_IOS_1 = lastRow(important_sheet11,"X");
  
  var lastRow5TH_GAME_SP_IOS_3 = lastRow(important_sheet11,"AC");
  var lastRow5TH_GAME_SP_IOS_2 = lastRow(important_sheet11,"AD");
  var lastRow5TH_GAME_SP_IOS_1 = lastRow(important_sheet11,"AE");
  
  var lastRow6TH_GAME_SP_IOS_3 = lastRow(important_sheet11,"AJ");
  var lastRow6TH_GAME_SP_IOS_2 = lastRow(important_sheet11,"AK");
  var lastRow6TH_GAME_SP_IOS_1 = lastRow(important_sheet11,"AL");
  
  
  var lastRow1ST_GAME_FR = lastRow(important_sheet4,"A");
  var lastRow2ND_GAME_FR	= lastRow(important_sheet4,"D");
  var lastRow3RD_GAME_FR = lastRow(important_sheet4,"G");
  var lastRow4TH_GAME_FR	= lastRow(important_sheet4,"J");
  var lastRow5TH_GAME_FR = lastRow(important_sheet4,"M");
  var lastRow6TH_GAME_FR	= lastRow(important_sheet4,"P");
  
  var lastRow1ST_GAME_FR_IOS_3 = lastRow(important_sheet12,"A");
  var lastRow1ST_GAME_FR_IOS_2 = lastRow(important_sheet12,"B");
  var lastRow1ST_GAME_FR_IOS_1 = lastRow(important_sheet12,"C");
  
  var lastRow2ND_GAME_FR_IOS_3 = lastRow(important_sheet12,"H");
  var lastRow2ND_GAME_FR_IOS_2 = lastRow(important_sheet12,"I");
  var lastRow2ND_GAME_FR_IOS_1 = lastRow(important_sheet12,"J");
  
  var lastRow3RD_GAME_FR_IOS_3 = lastRow(important_sheet12,"O");
  var lastRow3RD_GAME_FR_IOS_2 = lastRow(important_sheet12,"P");
  var lastRow3RD_GAME_FR_IOS_1 = lastRow(important_sheet12,"Q");
  
  var lastRow4TH_GAME_FR_IOS_3 = lastRow(important_sheet12,"V");
  var lastRow4TH_GAME_FR_IOS_2 = lastRow(important_sheet12,"W");
  var lastRow4TH_GAME_FR_IOS_1 = lastRow(important_sheet12,"X");
  
  var lastRow5TH_GAME_FR_IOS_3 = lastRow(important_sheet12,"AC");
  var lastRow5TH_GAME_FR_IOS_2 = lastRow(important_sheet12,"AD");
  var lastRow5TH_GAME_FR_IOS_1 = lastRow(important_sheet12,"AE");
  
  var lastRow6TH_GAME_FR_IOS_3 = lastRow(important_sheet12,"AJ");
  var lastRow6TH_GAME_FR_IOS_2 = lastRow(important_sheet12,"AK");
  var lastRow6TH_GAME_FR_IOS_1 = lastRow(important_sheet12,"AL");  
  
  
  var lastRow1ST_GAME_IT = lastRow(important_sheet5,"A");
  var lastRow2ND_GAME_IT	= lastRow(important_sheet5,"D");
  var lastRow3RD_GAME_IT = lastRow(important_sheet5,"G");
  var lastRow4TH_GAME_IT	= lastRow(important_sheet5,"J");
  var lastRow5TH_GAME_IT = lastRow(important_sheet5,"M");
  var lastRow6TH_GAME_IT	= lastRow(important_sheet5,"P");
  
  var lastRow1ST_GAME_IT_IOS_3 = lastRow(important_sheet13,"A");
  var lastRow1ST_GAME_IT_IOS_2 = lastRow(important_sheet13,"B");
  var lastRow1ST_GAME_IT_IOS_1 = lastRow(important_sheet13,"C");
  
  var lastRow2ND_GAME_IT_IOS_3 = lastRow(important_sheet13,"H");
  var lastRow2ND_GAME_IT_IOS_2 = lastRow(important_sheet13,"I");
  var lastRow2ND_GAME_IT_IOS_1 = lastRow(important_sheet13,"J");
  
  var lastRow3RD_GAME_IT_IOS_3 = lastRow(important_sheet13,"O");
  var lastRow3RD_GAME_IT_IOS_2 = lastRow(important_sheet13,"P");
  var lastRow3RD_GAME_IT_IOS_1 = lastRow(important_sheet13,"Q");
  
  var lastRow4TH_GAME_IT_IOS_3 = lastRow(important_sheet13,"V");
  var lastRow4TH_GAME_IT_IOS_2 = lastRow(important_sheet13,"W");
  var lastRow4TH_GAME_IT_IOS_1 = lastRow(important_sheet13,"X");
  
  var lastRow5TH_GAME_IT_IOS_3 = lastRow(important_sheet13,"AC");
  var lastRow5TH_GAME_IT_IOS_2 = lastRow(important_sheet13,"AD");
  var lastRow5TH_GAME_IT_IOS_1 = lastRow(important_sheet13,"AE");
  
  var lastRow6TH_GAME_IT_IOS_3 = lastRow(important_sheet13,"AJ");
  var lastRow6TH_GAME_IT_IOS_2 = lastRow(important_sheet13,"AK");
  var lastRow6TH_GAME_IT_IOS_1 = lastRow(important_sheet13,"AL");    
  
  
  var lastRow1ST_GAME_RU = lastRow(important_sheet6,"A");
  var lastRow2ND_GAME_RU	= lastRow(important_sheet6,"D");
  var lastRow3RD_GAME_RU = lastRow(important_sheet6,"G");
  var lastRow4TH_GAME_RU	= lastRow(important_sheet6,"J");
  var lastRow5TH_GAME_RU = lastRow(important_sheet6,"M");
  var lastRow6TH_GAME_RU	= lastRow(important_sheet6,"P");
  
  var lastRow1ST_GAME_RU_IOS_3 = lastRow(important_sheet14,"A");
  var lastRow1ST_GAME_RU_IOS_2 = lastRow(important_sheet14,"B");
  var lastRow1ST_GAME_RU_IOS_1 = lastRow(important_sheet14,"C");
  
  var lastRow2ND_GAME_RU_IOS_3 = lastRow(important_sheet14,"H");
  var lastRow2ND_GAME_RU_IOS_2 = lastRow(important_sheet14,"I");
  var lastRow2ND_GAME_RU_IOS_1 = lastRow(important_sheet14,"J");
  
  var lastRow3RD_GAME_RU_IOS_3 = lastRow(important_sheet14,"O");
  var lastRow3RD_GAME_RU_IOS_2 = lastRow(important_sheet14,"P");
  var lastRow3RD_GAME_RU_IOS_1 = lastRow(important_sheet14,"Q");
  
  var lastRow4TH_GAME_RU_IOS_3 = lastRow(important_sheet14,"V");
  var lastRow4TH_GAME_RU_IOS_2 = lastRow(important_sheet14,"W");
  var lastRow4TH_GAME_RU_IOS_1 = lastRow(important_sheet14,"X");
  
  var lastRow5TH_GAME_RU_IOS_3 = lastRow(important_sheet14,"AC");
  var lastRow5TH_GAME_RU_IOS_2 = lastRow(important_sheet14,"AD");
  var lastRow5TH_GAME_RU_IOS_1 = lastRow(important_sheet14,"AE");
  
  var lastRow6TH_GAME_RU_IOS_3 = lastRow(important_sheet14,"AJ");
  var lastRow6TH_GAME_RU_IOS_2 = lastRow(important_sheet14,"AK");
  var lastRow6TH_GAME_RU_IOS_1 = lastRow(important_sheet14,"AL");     
  
  
  var lastRow1ST_GAME_CN_S = lastRow(important_sheet7,"A");
  var lastRow2ND_GAME_CN_S	= lastRow(important_sheet7,"D");
  var lastRow3RD_GAME_CN_S = lastRow(important_sheet7,"G");
  var lastRow4TH_GAME_CN_S	= lastRow(important_sheet7,"J");
  var lastRow5TH_GAME_CN_S = lastRow(important_sheet7,"M");
  var lastRow6TH_GAME_CN_S	= lastRow(important_sheet7,"P");
  
  var lastRow1ST_GAME_CN_S_IOS_3 = lastRow(important_sheet15,"A");
  var lastRow1ST_GAME_CN_S_IOS_2 = lastRow(important_sheet15,"B");
  var lastRow1ST_GAME_CN_S_IOS_1 = lastRow(important_sheet15,"C");
  
  var lastRow2ND_GAME_CN_S_IOS_3 = lastRow(important_sheet15,"H");
  var lastRow2ND_GAME_CN_S_IOS_2 = lastRow(important_sheet15,"I");
  var lastRow2ND_GAME_CN_S_IOS_1 = lastRow(important_sheet15,"J");
  
  var lastRow3RD_GAME_CN_S_IOS_3 = lastRow(important_sheet15,"O");
  var lastRow3RD_GAME_CN_S_IOS_2 = lastRow(important_sheet15,"P");
  var lastRow3RD_GAME_CN_S_IOS_1 = lastRow(important_sheet15,"Q");
  
  var lastRow4TH_GAME_CN_S_IOS_3 = lastRow(important_sheet15,"V");
  var lastRow4TH_GAME_CN_S_IOS_2 = lastRow(important_sheet15,"W");
  var lastRow4TH_GAME_CN_S_IOS_1 = lastRow(important_sheet15,"X");
  
  var lastRow5TH_GAME_CN_S_IOS_3 = lastRow(important_sheet15,"AC");
  var lastRow5TH_GAME_CN_S_IOS_2 = lastRow(important_sheet15,"AD");
  var lastRow5TH_GAME_CN_S_IOS_1 = lastRow(important_sheet15,"AE");
  
  var lastRow6TH_GAME_CN_S_IOS_3 = lastRow(important_sheet15,"AJ");
  var lastRow6TH_GAME_CN_S_IOS_2 = lastRow(important_sheet15,"AK");
  var lastRow6TH_GAME_CN_S_IOS_1 = lastRow(important_sheet15,"AL");   
  
  
  var lastRow1ST_GAME_PT = lastRow(important_sheet8,"A");
  var lastRow2ND_GAME_PT	= lastRow(important_sheet8,"D");
  var lastRow3RD_GAME_PT = lastRow(important_sheet8,"G");
  var lastRow4TH_GAME_PT	= lastRow(important_sheet8,"J");
  var lastRow5TH_GAME_PT = lastRow(important_sheet8,"M");
  var lastRow6TH_GAME_PT	= lastRow(important_sheet8,"P");
  
  var lastRow1ST_GAME_PT_IOS_3 = lastRow(important_sheet16,"A");
  var lastRow1ST_GAME_PT_IOS_2 = lastRow(important_sheet16,"B");
  var lastRow1ST_GAME_PT_IOS_1 = lastRow(important_sheet16,"C");
  
  var lastRow2ND_GAME_PT_IOS_3 = lastRow(important_sheet16,"H");
  var lastRow2ND_GAME_PT_IOS_2 = lastRow(important_sheet16,"I");
  var lastRow2ND_GAME_PT_IOS_1 = lastRow(important_sheet16,"J");
  
  var lastRow3RD_GAME_PT_IOS_3 = lastRow(important_sheet16,"O");
  var lastRow3RD_GAME_PT_IOS_2 = lastRow(important_sheet16,"P");
  var lastRow3RD_GAME_PT_IOS_1 = lastRow(important_sheet16,"Q");
  
  var lastRow4TH_GAME_PT_IOS_3 = lastRow(important_sheet16,"V");
  var lastRow4TH_GAME_PT_IOS_2 = lastRow(important_sheet16,"W");
  var lastRow4TH_GAME_PT_IOS_1 = lastRow(important_sheet16,"X");
  
  var lastRow5TH_GAME_PT_IOS_3 = lastRow(important_sheet16,"AC");
  var lastRow5TH_GAME_PT_IOS_2 = lastRow(important_sheet16,"AD");
  var lastRow5TH_GAME_PT_IOS_1 = lastRow(important_sheet16,"AE");
  
  var lastRow6TH_GAME_PT_IOS_3 = lastRow(important_sheet16,"AJ");
  var lastRow6TH_GAME_PT_IOS_2 = lastRow(important_sheet16,"AK");
  var lastRow6TH_GAME_PT_IOS_1 = lastRow(important_sheet16,"AL");   
  
  
  var lastRow1ST_GAME_DE = lastRow(important_sheet9,"A");
  var lastRow2ND_GAME_DE	= lastRow(important_sheet9,"D");
  var lastRow3RD_GAME_DE = lastRow(important_sheet9,"G");
  var lastRow4TH_GAME_DE	= lastRow(important_sheet9,"J");
  var lastRow5TH_GAME_DE = lastRow(important_sheet9,"M");
  var lastRow6TH_GAME_DE	= lastRow(important_sheet9,"P");
  
  var lastRow1ST_GAME_DE_IOS_3 = lastRow(important_sheet17,"A");
  var lastRow1ST_GAME_DE_IOS_2 = lastRow(important_sheet17,"B");
  var lastRow1ST_GAME_DE_IOS_1 = lastRow(important_sheet17,"C");
  
  var lastRow2ND_GAME_DE_IOS_3 = lastRow(important_sheet17,"H");
  var lastRow2ND_GAME_DE_IOS_2 = lastRow(important_sheet17,"I");
  var lastRow2ND_GAME_DE_IOS_1 = lastRow(important_sheet17,"J");
  
  var lastRow3RD_GAME_DE_IOS_3 = lastRow(important_sheet17,"O");
  var lastRow3RD_GAME_DE_IOS_2 = lastRow(important_sheet17,"P");
  var lastRow3RD_GAME_DE_IOS_1 = lastRow(important_sheet17,"Q");
  
  var lastRow4TH_GAME_DE_IOS_3 = lastRow(important_sheet17,"V");
  var lastRow4TH_GAME_DE_IOS_2 = lastRow(important_sheet17,"W");
  var lastRow4TH_GAME_DE_IOS_1 = lastRow(important_sheet17,"X");
  
  var lastRow5TH_GAME_DE_IOS_3 = lastRow(important_sheet17,"AC");
  var lastRow5TH_GAME_DE_IOS_2 = lastRow(important_sheet17,"AD");
  var lastRow5TH_GAME_DE_IOS_1 = lastRow(important_sheet17,"AE");
  
  var lastRow6TH_GAME_DE_IOS_3 = lastRow(important_sheet17,"AJ");
  var lastRow6TH_GAME_DE_IOS_2 = lastRow(important_sheet17,"AK");
  var lastRow6TH_GAME_DE_IOS_1 = lastRow(important_sheet17,"AL");     
  
  
  var ui = SpreadsheetApp.getUi();
  
  var numOfAditions = 0 ;
  var numOfAditions2 = 0 ;
  
  var numOfDaysSinceDate = 0;
  
  var dateAndroid;
  
  //Feel free to edit WasTheLastTimeTooLongAgo as it's the amount of days that need to pass so that a Google Play Console entry appears with red background
  var WasTheLastTimeTooLongAgo = 3;
  
  //Feel free to edit WasTheLastTimeTooLongAgo2 as it's the amount of days that need to pass so that a Google Play Console entry appears with red background
  var WasTheLastTimeTooLongAgo2 = 7;
  
  //Setting up the current date
  important_sheet1.getRange(2,13).setValue(Utilities.formatDate(new Date(), "GMT+2", "yyyy-MM-dd"));
  
  
  
  for ( i = 2 ; i <= 9 ; i++) {
    for ( j = 2 ; j <= 7 ; j++) {
      numOfAditions = numOfAditions + 1;
      switch(numOfAditions) 
      {
        case 1:
          dateAndroid = important_sheet2.getRange(lastRow1ST_GAME_EN, 1).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 2: 
          dateAndroid = important_sheet2.getRange(lastRow2ND_GAME_EN, 4).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 3: 
          dateAndroid = important_sheet2.getRange(lastRow3RD_GAME_EN, 7).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 4: 
          dateAndroid = important_sheet2.getRange(lastRow4TH_GAME_EN, 10).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 5: 
          dateAndroid = important_sheet2.getRange(lastRow5TH_GAME_EN, 13).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 6: 
          dateAndroid = important_sheet2.getRange(lastRow6TH_GAME_EN, 16).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
          //SPANISH
        case 7: 
          dateAndroid = important_sheet3.getRange(lastRow1ST_GAME_SP, 1).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 8: 
          dateAndroid = important_sheet3.getRange(lastRow2ND_GAME_SP, 4).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 9: 
          dateAndroid = important_sheet3.getRange(lastRow3RD_GAME_SP, 7).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 10: 
          dateAndroid = important_sheet3.getRange(lastRow4TH_GAME_SP, 10).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 11: 
          dateAndroid = important_sheet3.getRange(lastRow5TH_GAME_SP, 13).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 12: 
          dateAndroid = important_sheet3.getRange(lastRow6TH_GAME_SP, 16).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
          //FRENCH
        case 13:
          dateAndroid = important_sheet4.getRange(lastRow1ST_GAME_FR, 1).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 14: 
          dateAndroid = important_sheet4.getRange(lastRow2ND_GAME_FR, 4).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 15: 
          dateAndroid = important_sheet4.getRange(lastRow3RD_GAME_FR, 7).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 16: 
          dateAndroid = important_sheet4.getRange(lastRow4TH_GAME_FR, 10).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 17: 
          dateAndroid = important_sheet4.getRange(lastRow5TH_GAME_FR, 13).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 18: 
          dateAndroid = important_sheet4.getRange(lastRow6TH_GAME_FR, 16).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
          //ITALIAN
        case 19: 
          dateAndroid = important_sheet5.getRange(lastRow1ST_GAME_IT, 1).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 20: 
          dateAndroid = important_sheet5.getRange(lastRow2ND_GAME_IT, 4).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 21: 
          dateAndroid = important_sheet5.getRange(lastRow3RD_GAME_IT, 7).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 22: 
          dateAndroid = important_sheet5.getRange(lastRow4TH_GAME_IT, 10).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 23: 
          dateAndroid = important_sheet5.getRange(lastRow5TH_GAME_IT, 13).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 24: 
          dateAndroid = important_sheet5.getRange(lastRow6TH_GAME_IT, 16).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
          //RUSSIAN
        case 25: 
          dateAndroid = important_sheet6.getRange(lastRow1ST_GAME_RU, 1).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 26: 
          dateAndroid = important_sheet6.getRange(lastRow2ND_GAME_RU, 4).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 27: 
          dateAndroid = important_sheet6.getRange(lastRow3RD_GAME_RU, 7).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 28: 
          dateAndroid = important_sheet6.getRange(lastRow4TH_GAME_RU, 10).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 29: 
          dateAndroid = important_sheet6.getRange(lastRow5TH_GAME_RU, 13).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 30: 
          dateAndroid = important_sheet6.getRange(lastRow6TH_GAME_RU, 16).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
          //CHINESE SIMPLIFIED
        case 31: 
          dateAndroid = important_sheet7.getRange(lastRow1ST_GAME_CN_S, 1).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 32: 
          dateAndroid = important_sheet7.getRange(lastRow2ND_GAME_CN_S, 4).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 33: 
          dateAndroid = important_sheet7.getRange(lastRow3RD_GAME_CN_S, 7).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 34: 
          dateAndroid = important_sheet7.getRange(lastRow4TH_GAME_CN_S, 10).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 35: 
          dateAndroid = important_sheet7.getRange(lastRow5TH_GAME_CN_S, 13).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 36: 
          dateAndroid = important_sheet7.getRange(lastRow6TH_GAME_CN_S, 16).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
          //PORTUGUESE
        case 37: 
          dateAndroid = important_sheet8.getRange(lastRow1ST_GAME_PT, 1).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 38: 
          dateAndroid = important_sheet8.getRange(lastRow1ST_GAME_PT, 4).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 39: 
          dateAndroid = important_sheet8.getRange(lastRow1ST_GAME_PT, 7).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 40: 
          dateAndroid = important_sheet8.getRange(lastRow1ST_GAME_PT, 10).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 41: 
          dateAndroid = important_sheet8.getRange(lastRow1ST_GAME_PT, 13).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 42: 
          dateAndroid = important_sheet8.getRange(lastRow1ST_GAME_PT, 16).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
          //GERMAN
        case 43: 
          dateAndroid = important_sheet9.getRange(lastRow1ST_GAME_DE, 1).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 44: 
          dateAndroid = important_sheet9.getRange(lastRow2ND_GAME_DE, 4).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 45:
          dateAndroid = important_sheet9.getRange(lastRow3RD_GAME_DE, 7).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 46: 
          dateAndroid = important_sheet9.getRange(lastRow4TH_GAME_DE, 10).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 47: 
          dateAndroid = important_sheet9.getRange(lastRow5TH_GAME_DE, 13).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        case 48: 
          dateAndroid = important_sheet9.getRange(lastRow6TH_GAME_DE, 16).getDisplayValue();
          if (dateAndroid != 'Last version'){
            important_sheet1.getRange(i,j).setValue(dateAndroid);
          }
          break;
        default:
          ui.alert("Smt wrong with Android bruh...");
          break;
      }
      
      numOfDaysSinceDate = important_sheet1.getRange(3,13).getValue() - important_sheet1.getRange(i,j+14).getValue();
      
      if ( numOfDaysSinceDate >= WasTheLastTimeTooLongAgo ){
        important_sheet1.getRange(i,j).setBackground("red");
      } else {
        important_sheet1.getRange(i,j).setBackground("white");
      }
      
    }
  }
  
  var date1;
  var date2;
  var date3;
  
  for ( i = 11 ; i <= 18 ; i++) {
    for ( j = 2 ; j <= 7 ; j++) {
      numOfAditions2 = numOfAditions2 + 1;
      switch(numOfAditions2) 
      {
        case 1:
          date1 = important_sheet10.getRange(lastRow1ST_GAME_EN_IOS_3, 1).getDisplayValue();
          date2 = important_sheet10.getRange(lastRow1ST_GAME_EN_IOS_2, 2).getDisplayValue();
          date3 = important_sheet10.getRange(lastRow1ST_GAME_EN_IOS_1, 3).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 2:
          date1 = important_sheet10.getRange(lastRow2ND_GAME_EN_IOS_3, 8).getDisplayValue();
          date2 = important_sheet10.getRange(lastRow2ND_GAME_EN_IOS_2, 9).getDisplayValue();
          date3 = important_sheet10.getRange(lastRow2ND_GAME_EN_IOS_1, 10).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 3:
          date1 = important_sheet10.getRange(lastRow3RD_GAME_EN_IOS_3, 15).getDisplayValue();
          date2 = important_sheet10.getRange(lastRow3RD_GAME_EN_IOS_2, 16).getDisplayValue();
          date3 = important_sheet10.getRange(lastRow3RD_GAME_EN_IOS_1, 17).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 4:
          date1 = important_sheet10.getRange(lastRow4TH_GAME_EN_IOS_3, 22).getDisplayValue();
          date2 = important_sheet10.getRange(lastRow4TH_GAME_EN_IOS_2, 23).getDisplayValue();
          date3 = important_sheet10.getRange(lastRow4TH_GAME_EN_IOS_1, 24).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 5:
          date1 = important_sheet10.getRange(lastRow5TH_GAME_EN_IOS_3, 29).getDisplayValue();
          date2 = important_sheet10.getRange(lastRow5TH_GAME_EN_IOS_2, 30).getDisplayValue();
          date3 = important_sheet10.getRange(lastRow5TH_GAME_EN_IOS_1, 31).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 6:
          date1 = important_sheet10.getRange(lastRow6TH_GAME_EN_IOS_3, 36).getDisplayValue();
          date2 = important_sheet10.getRange(lastRow6TH_GAME_EN_IOS_2, 37).getDisplayValue();
          date3 = important_sheet10.getRange(lastRow6TH_GAME_EN_IOS_1, 38).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 7:
          date1 = important_sheet11.getRange(lastRow1ST_GAME_SP_IOS_3, 1).getDisplayValue();
          date2 = important_sheet11.getRange(lastRow1ST_GAME_SP_IOS_2, 2).getDisplayValue();
          date3 = important_sheet11.getRange(lastRow1ST_GAME_SP_IOS_1, 3).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 8:
          date1 = important_sheet11.getRange(lastRow2ND_GAME_SP_IOS_3, 8).getDisplayValue();
          date2 = important_sheet11.getRange(lastRow2ND_GAME_SP_IOS_2, 9).getDisplayValue();
          date3 = important_sheet11.getRange(lastRow2ND_GAME_SP_IOS_1, 10).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 9:
          date1 = important_sheet11.getRange(lastRow3RD_GAME_SP_IOS_3, 15).getDisplayValue();
          date2 = important_sheet11.getRange(lastRow3RD_GAME_SP_IOS_2, 16).getDisplayValue();
          date3 = important_sheet11.getRange(lastRow3RD_GAME_SP_IOS_1, 17).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 10:
          date1 = important_sheet11.getRange(lastRow4TH_GAME_SP_IOS_3, 22).getDisplayValue();
          date2 = important_sheet11.getRange(lastRow4TH_GAME_SP_IOS_2, 23).getDisplayValue();
          date3 = important_sheet11.getRange(lastRow4TH_GAME_SP_IOS_1, 24).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 11:
          date1 = important_sheet11.getRange(lastRow5TH_GAME_SP_IOS_3, 29).getDisplayValue();
          date2 = important_sheet11.getRange(lastRow5TH_GAME_SP_IOS_2, 30).getDisplayValue();
          date3 = important_sheet11.getRange(lastRow5TH_GAME_SP_IOS_1, 31).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 12:
          date1 = important_sheet11.getRange(lastRow6TH_GAME_SP_IOS_3, 36).getDisplayValue();
          date2 = important_sheet11.getRange(lastRow6TH_GAME_SP_IOS_2, 37).getDisplayValue();
          date3 = important_sheet11.getRange(lastRow6TH_GAME_SP_IOS_1, 38).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 13:
          date1 = important_sheet12.getRange(lastRow1ST_GAME_FR_IOS_3, 1).getDisplayValue();
          date2 = important_sheet12.getRange(lastRow1ST_GAME_FR_IOS_2, 2).getDisplayValue();
          date3 = important_sheet12.getRange(lastRow1ST_GAME_FR_IOS_1, 3).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 14:
          date1 = important_sheet12.getRange(lastRow2ND_GAME_FR_IOS_3, 8).getDisplayValue();
          date2 = important_sheet12.getRange(lastRow2ND_GAME_FR_IOS_2, 9).getDisplayValue();
          date3 = important_sheet12.getRange(lastRow2ND_GAME_FR_IOS_1, 10).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 15:
          date1 = important_sheet12.getRange(lastRow3RD_GAME_FR_IOS_3, 15).getDisplayValue();
          date2 = important_sheet12.getRange(lastRow3RD_GAME_FR_IOS_2, 16).getDisplayValue();
          date3 = important_sheet12.getRange(lastRow3RD_GAME_FR_IOS_1, 17).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 16:
          date1 = important_sheet12.getRange(lastRow4TH_GAME_FR_IOS_3, 22).getDisplayValue();
          date2 = important_sheet12.getRange(lastRow4TH_GAME_FR_IOS_2, 23).getDisplayValue();
          date3 = important_sheet12.getRange(lastRow4TH_GAME_FR_IOS_1, 24).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 17:
          date1 = important_sheet12.getRange(lastRow5TH_GAME_FR_IOS_3, 29).getDisplayValue();
          date2 = important_sheet12.getRange(lastRow5TH_GAME_FR_IOS_2, 30).getDisplayValue();
          date3 = important_sheet12.getRange(lastRow5TH_GAME_FR_IOS_1, 31).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 18:
          date1 = important_sheet12.getRange(lastRow6TH_GAME_FR_IOS_3, 36).getDisplayValue();
          date2 = important_sheet12.getRange(lastRow6TH_GAME_FR_IOS_2, 37).getDisplayValue();
          date3 = important_sheet12.getRange(lastRow6TH_GAME_FR_IOS_1, 38).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 19:
          date1 = important_sheet13.getRange(lastRow1ST_GAME_IT_IOS_3, 1).getDisplayValue();
          date2 = important_sheet13.getRange(lastRow1ST_GAME_IT_IOS_2, 2).getDisplayValue();
          date3 = important_sheet13.getRange(lastRow1ST_GAME_IT_IOS_1, 3).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 20:
          date1 = important_sheet13.getRange(lastRow2ND_GAME_IT_IOS_3, 8).getDisplayValue();
          date2 = important_sheet13.getRange(lastRow2ND_GAME_IT_IOS_2, 9).getDisplayValue();
          date3 = important_sheet13.getRange(lastRow2ND_GAME_IT_IOS_1, 10).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 21:
          date1 = important_sheet13.getRange(lastRow3RD_GAME_IT_IOS_3, 15).getDisplayValue();
          date2 = important_sheet13.getRange(lastRow3RD_GAME_IT_IOS_2, 16).getDisplayValue();
          date3 = important_sheet13.getRange(lastRow3RD_GAME_IT_IOS_1, 17).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 22:
          date1 = important_sheet13.getRange(lastRow4TH_GAME_IT_IOS_3, 22).getDisplayValue();
          date2 = important_sheet13.getRange(lastRow4TH_GAME_IT_IOS_2, 23).getDisplayValue();
          date3 = important_sheet13.getRange(lastRow4TH_GAME_IT_IOS_1, 24).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 23:
          date1 = important_sheet13.getRange(lastRow5TH_GAME_IT_IOS_3, 29).getDisplayValue();
          date2 = important_sheet13.getRange(lastRow5TH_GAME_IT_IOS_2, 30).getDisplayValue();
          date3 = important_sheet13.getRange(lastRow5TH_GAME_IT_IOS_1, 31).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 24:
          date1 = important_sheet13.getRange(lastRow6TH_GAME_IT_IOS_3, 36).getDisplayValue();
          date2 = important_sheet13.getRange(lastRow6TH_GAME_IT_IOS_2, 37).getDisplayValue();
          date3 = important_sheet13.getRange(lastRow6TH_GAME_IT_IOS_1, 38).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 25:
          date1 = important_sheet14.getRange(lastRow1ST_GAME_RU_IOS_3, 1).getDisplayValue();
          date2 = important_sheet14.getRange(lastRow1ST_GAME_RU_IOS_2, 2).getDisplayValue();
          date3 = important_sheet14.getRange(lastRow1ST_GAME_RU_IOS_1, 3).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 26:
          date1 = important_sheet14.getRange(lastRow2ND_GAME_RU_IOS_3, 8).getDisplayValue();
          date2 = important_sheet14.getRange(lastRow2ND_GAME_RU_IOS_2, 9).getDisplayValue();
          date3 = important_sheet14.getRange(lastRow2ND_GAME_RU_IOS_1, 10).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 27:
          date1 = important_sheet14.getRange(lastRow3RD_GAME_RU_IOS_3, 15).getDisplayValue();
          date2 = important_sheet14.getRange(lastRow3RD_GAME_RU_IOS_2, 16).getDisplayValue();
          date3 = important_sheet14.getRange(lastRow3RD_GAME_RU_IOS_1, 17).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 28:
          date1 = important_sheet14.getRange(lastRow4TH_GAME_RU_IOS_3, 22).getDisplayValue();
          date2 = important_sheet14.getRange(lastRow4TH_GAME_RU_IOS_2, 23).getDisplayValue();
          date3 = important_sheet14.getRange(lastRow4TH_GAME_RU_IOS_1, 24).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 29:
          date1 = important_sheet14.getRange(lastRow5TH_GAME_RU_IOS_3, 29).getDisplayValue();
          date2 = important_sheet14.getRange(lastRow5TH_GAME_RU_IOS_2, 30).getDisplayValue();
          date3 = important_sheet14.getRange(lastRow5TH_GAME_RU_IOS_1, 31).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 30:
          date1 = important_sheet14.getRange(lastRow6TH_GAME_RU_IOS_3, 36).getDisplayValue();
          date2 = important_sheet14.getRange(lastRow6TH_GAME_RU_IOS_2, 37).getDisplayValue();
          date3 = important_sheet14.getRange(lastRow6TH_GAME_RU_IOS_1, 38).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 31:
          date1 = important_sheet15.getRange(lastRow1ST_GAME_CN_S_IOS_3, 1).getDisplayValue();
          date2 = important_sheet15.getRange(lastRow1ST_GAME_CN_S_IOS_2, 2).getDisplayValue();
          date3 = important_sheet15.getRange(lastRow1ST_GAME_CN_S_IOS_1, 3).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 32:
          date1 = important_sheet15.getRange(lastRow2ND_GAME_CN_S_IOS_3, 8).getDisplayValue();
          date2 = important_sheet15.getRange(lastRow2ND_GAME_CN_S_IOS_2, 9).getDisplayValue();
          date3 = important_sheet15.getRange(lastRow2ND_GAME_CN_S_IOS_1, 10).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 33:
          date1 = important_sheet15.getRange(lastRow3RD_GAME_CN_S_IOS_3, 15).getDisplayValue();
          date2 = important_sheet15.getRange(lastRow3RD_GAME_CN_S_IOS_2, 16).getDisplayValue();
          date3 = important_sheet15.getRange(lastRow3RD_GAME_CN_S_IOS_1, 17).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 34:
          date1 = important_sheet15.getRange(lastRow4TH_GAME_CN_S_IOS_3, 22).getDisplayValue();
          date2 = important_sheet15.getRange(lastRow4TH_GAME_CN_S_IOS_2, 23).getDisplayValue();
          date3 = important_sheet15.getRange(lastRow4TH_GAME_CN_S_IOS_1, 24).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 35:
          date1 = important_sheet15.getRange(lastRow5TH_GAME_CN_S_IOS_3, 29).getDisplayValue();
          date2 = important_sheet15.getRange(lastRow5TH_GAME_CN_S_IOS_2, 30).getDisplayValue();
          date3 = important_sheet15.getRange(lastRow5TH_GAME_CN_S_IOS_1, 31).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 36:
          date1 = important_sheet15.getRange(lastRow6TH_GAME_CN_S_IOS_3, 36).getDisplayValue();
          date2 = important_sheet15.getRange(lastRow6TH_GAME_CN_S_IOS_2, 37).getDisplayValue();
          date3 = important_sheet15.getRange(lastRow6TH_GAME_CN_S_IOS_1, 38).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 37:
          date1 = important_sheet16.getRange(lastRow1ST_GAME_PT_IOS_3, 1).getDisplayValue();
          date2 = important_sheet16.getRange(lastRow1ST_GAME_PT_IOS_2, 2).getDisplayValue();
          date3 = important_sheet16.getRange(lastRow1ST_GAME_PT_IOS_1, 3).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 38:
          date1 = important_sheet16.getRange(lastRow2ND_GAME_PT_IOS_3, 8).getDisplayValue();
          date2 = important_sheet16.getRange(lastRow2ND_GAME_PT_IOS_2, 9).getDisplayValue();
          date3 = important_sheet16.getRange(lastRow2ND_GAME_PT_IOS_1, 10).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 39:
          date1 = important_sheet16.getRange(lastRow3RD_GAME_PT_IOS_3, 15).getDisplayValue();
          date2 = important_sheet16.getRange(lastRow3RD_GAME_PT_IOS_2, 16).getDisplayValue();
          date3 = important_sheet16.getRange(lastRow3RD_GAME_PT_IOS_1, 17).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 40:
          date1 = important_sheet16.getRange(lastRow4TH_GAME_PT_IOS_3, 22).getDisplayValue();
          date2 = important_sheet16.getRange(lastRow4TH_GAME_PT_IOS_2, 23).getDisplayValue();
          date3 = important_sheet16.getRange(lastRow4TH_GAME_PT_IOS_1, 24).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 41:
          date1 = important_sheet16.getRange(lastRow5TH_GAME_PT_IOS_3, 29).getDisplayValue();
          date2 = important_sheet16.getRange(lastRow5TH_GAME_PT_IOS_2, 30).getDisplayValue();
          date3 = important_sheet16.getRange(lastRow5TH_GAME_PT_IOS_1, 31).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 42:
          date1 = important_sheet16.getRange(lastRow6TH_GAME_PT_IOS_3, 36).getDisplayValue();
          date2 = important_sheet16.getRange(lastRow6TH_GAME_PT_IOS_2, 37).getDisplayValue();
          date3 = important_sheet16.getRange(lastRow6TH_GAME_PT_IOS_1, 38).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 43:
          date1 = important_sheet17.getRange(lastRow1ST_GAME_DE_IOS_3, 1).getDisplayValue();
          date2 = important_sheet17.getRange(lastRow1ST_GAME_DE_IOS_2, 2).getDisplayValue();
          date3 = important_sheet17.getRange(lastRow1ST_GAME_DE_IOS_1, 3).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 44:
          date1 = important_sheet17.getRange(lastRow2ND_GAME_DE_IOS_3, 8).getDisplayValue();
          date2 = important_sheet17.getRange(lastRow2ND_GAME_DE_IOS_2, 9).getDisplayValue();
          date3 = important_sheet17.getRange(lastRow2ND_GAME_DE_IOS_1, 10).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 45:
          date1 = important_sheet17.getRange(lastRow3RD_GAME_DE_IOS_3, 15).getDisplayValue();
          date2 = important_sheet17.getRange(lastRow3RD_GAME_DE_IOS_2, 16).getDisplayValue();
          date3 = important_sheet17.getRange(lastRow3RD_GAME_DE_IOS_1, 17).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 46:
          date1 = important_sheet17.getRange(lastRow4TH_GAME_DE_IOS_3, 22).getDisplayValue();
          date2 = important_sheet17.getRange(lastRow4TH_GAME_DE_IOS_2, 23).getDisplayValue();
          date3 = important_sheet17.getRange(lastRow4TH_GAME_DE_IOS_1, 24).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 47:
          date1 = important_sheet17.getRange(lastRow5TH_GAME_DE_IOS_3, 29).getDisplayValue();
          date2 = important_sheet17.getRange(lastRow5TH_GAME_DE_IOS_2, 30).getDisplayValue();
          date3 = important_sheet17.getRange(lastRow5TH_GAME_DE_IOS_1, 31).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        case 48:
          date1 = important_sheet17.getRange(lastRow6TH_GAME_DE_IOS_3, 36).getDisplayValue();
          date2 = important_sheet17.getRange(lastRow6TH_GAME_DE_IOS_2, 37).getDisplayValue();
          date3 = important_sheet17.getRange(lastRow6TH_GAME_DE_IOS_1, 38).getDisplayValue();
          if (date1 != '3 ★' && date2 != '2 ★' && date3 != '1 ★'){
            important_sheet1.getRange(i,j).setValue(oldestDate(date1, date2, date3));
          }
          break;
        default:
          ui.alert("Smt wrong with iOS bruh...");
          break;
      }
      
      numOfDaysSinceDate = important_sheet1.getRange(3,13).getValue() - important_sheet1.getRange(i,j+14).getValue();
      
      if ( numOfDaysSinceDate >= WasTheLastTimeTooLongAgo2 ){
        important_sheet1.getRange(i,j).setBackground("red");
      } else {
        important_sheet1.getRange(i,j).setBackground("white");
      }
      
      
    }
  }
}
