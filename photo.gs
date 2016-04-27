
var templateID = "1f16dlHEG4cE8p02nRFnKBKgYOxyyjBDHM2pDI_SHhSU"; // template for google doc
var sheetID = "1rAAyT9CS0KHhNFvXb19deo--nO3jStRCUplm2ohoIJk"; //for testing...

function photo_update(e){
  var s =  e.source.getActiveSheet();
  //var s = SpreadsheetApp.openById(sheetID).getSheets()[0]; //in case source doesn't work 
  GetData(s, templateID);
}
