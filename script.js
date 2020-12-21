/**
 * Manipulating Google Spreadsheet
 *
 * I had a google spreadsheet from which I needed to fetch some data 
 * and populate it in a particular format in a new sheet.
 * 
 */
function processSpreadsheet() 
{
  var ss              = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet       = ss.getSheetByName('Content'); //Sheet with data
  var processedSheet  = ss.getSheetByName('Result'); //Sheet in which result will be populated
  lastRow             = mainSheet.getLastRow();
  lastColumn          = mainSheet.getLastColumn();

  var itemsArray = mainSheet.getRange(1,1,lastRow,lastColumn).getValues();
  processedSheet.getRange(1,1, 1, 20).setValues([['ID', 'Company ID', 'Guest Name', 'Guest Phone', 'Guest Email', 'Gender', 'DOB', 'Spouse', 'Spouse DOB', 'Anniv Dt', 'Addr', 'Place', 'Country', 'Company', 'Comp Addr', 'Comp Pin', 'Comp Country', 'Lifetime Value', 'Marrital Status', 'Total Nights']]);
  var id = 3;
  var companyId = 1;
  
  for(var i=0; i<itemsArray.length; i++)
  {
    var addRow = 'yes';
    var marritalStatus = 'unmarried';
    var gender = 'male';
    
    if(itemsArray[i][0] === '' || itemsArray[i][0] === null)
    {
        addRow = 'no';
    }
    else 
    {
      addRow = 'yes';
      if(itemsArray[i][0] === 'Mrs.' || itemsArray[i][0] === 'Ms.') 
      {
        if(itemsArray[i][0] === 'Mrs.') 
        {
          marritalStatus = 'married';
        }
        
        gender = 'female';        
      }
      
      id = id + 1;
      
      var guestName = itemsArray[i][1];
      var guestPhone = itemsArray[i][2];
      var guestEmail = itemsArray[i][3];
      var guestDob = itemsArray[i][4];
      
      var spouseName = itemsArray[i][11];
      
      if(spouseName !== '' || spouseName !== null)
      {
        marritalStatus = 'married';
      }
      
      var spouseDob    = itemsArray[i][12];
      if(spouseDob === '' || spouseDob === null) 
      {
        spouseDob = '0000-00-00';
      }
      
      var annivDt      = itemsArray[i][5];
      if(annivDt === '' || annivDt === null) 
      {
        annivDt = '0000-00-00';
      }
      
      var guestAddr    = itemsArray[i][6] + itemsArray[i][7] + itemsArray[i][8];
      var guestCity    = itemsArray[i][9];
      var guestCountry = itemsArray[i][10];
      var guestComp    = itemsArray[i][13];
      var compAddr     = itemsArray[i][14] + itemsArray[i][15] + itemsArray[i][16] + itemsArray[i][17];
      var guestComp    = itemsArray[i][13];
      var compPin      = itemsArray[i][18];
      var compCountry  = itemsArray[i][19];
      var lifeTimeVal  = 0;
      var totalNights  = 0;
    }
    
    if(addRow == 'yes')
    {
      var lastCell = processedSheet.getRange(processedSheet.getLastRow()+1,1, 1, 20).setValues([[id, companyId, guestName, guestPhone, guestEmail, gender, guestDob, spouseName, spouseDob, annivDt, guestAddr, guestCity, guestCountry, guestComp, compAddr, compPin, compCountry, lifeTimeVal, marritalStatus, totalNights]]);
    }
    
  }
  
}
