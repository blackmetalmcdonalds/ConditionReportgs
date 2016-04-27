
//var docTemplateid = "17iRvirlypijs1nN9Q7_alt0ytbNu-tKuMHgbBHOvKLk"; // Painting template
//var sheetid = '1ewB_98Joydbu1Ils1G0SfS1MVdKRsKsK8m1ep_baFTA'; //response sheet


var folderid = '0BwUTWvTmYT6nTlNyajdZUzRORUk';//destination folder for report docs
var archiveID = '0BwUTWvTmYT6nS1JkQXJxbzlTbm8'; //archived docs

function RegExpescape(text){
 // var rep =  text.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
  return text.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
}

function archive(filename){
  var oldfile = DriveApp.getFileById(idFromFileName(filename));
  oldfile.makeCopy('archived:' + oldfile.getName(),DriveApp.getFolderById(archiveID)); //copy to archive folder and rename with 'archived'
  oldfile.setTrashed(true);}
  
  
function tryORInsert(s,colname, headers){

	var colnum =  getColByName(colname, headers);
	if (colnum == -1){
		var lastcol = s.getLastColumn();
		s.insertColumnAfter(lastcol);
		s.getRange(1, lastcol +1).setValue(colname);
	return lastcol + 1;}
	else{
		return colnum;
	}

	}


function GetData(s, docTemplateid){
  var data = s.getDataRange().getValues();
  var headers = data.shift();
  var filecol =  tryORInsert(s,'PDF', headers);
  var donecol = tryORInsert(s,'PDF-Timestamp', headers);
  var changecol = tryORInsert(s,'Changestamp',headers);
  for (var row = 0; row < data.length; row++) {
    if (printOrReprint(data[row][0], data[row][donecol], data[row][changecol], data[row][filecol]) &&
       data[row][0] !=''){
       var nd =new Date();
       var dname = data[row][1] + '_' + betterDate(nd);
       var folder = DriveApp.getFolderById(folderid);
       var copyId  = DriveApp.getFileById(docTemplateid).makeCopy(dname,folder).getId();
       var copyDoc = DocumentApp.openById(copyId);
       var copyBody = copyDoc.getActiveSection();
       var textbody = copyBody.getText();
      for (var col = 1; col < headers.length; col++) {
        var tag = '<<' + headers[col] + '>>';
        tag = RegExpescape(tag);
        var text = data[row][col];
        //var tpos = textbody.indexOf(tag)
        var pos = copyBody.findText(tag);
        if (tag != '<<undefined>>' &&  pos != undefined){
          if (isPhoto(text)){
            var post = pos.getElement();
            var par= post.getParent();
            var imname = text.slice(text.lastIndexOf('/')+1);
            var img = DriveApp.getFilesByName(imname).next().getBlob();
            var i = par.insertInlineImage(0, img);
            post.removeFromParent();
            var cell = par.getParent();
            
            //i = resize(600,450, i);
            i = resizebybounds(cell,i);
            var stylesp = {};
            stylesp[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
            par.setAttributes(stylesp);
            
            var stylesc = {};
            stylesc[DocumentApp.Attribute.PADDING_RIGHT] = 0;
            stylesc[DocumentApp.Attribute.PADDING_LEFT] = 0;
            stylesc[DocumentApp.Attribute.PADDING_TOP] = 0;
            stylesc[DocumentApp.Attribute.PADDING_BOTTOM] = 0;
            
            cell.setAttributes(stylesc);
           }
          
          else{
            copyBody.replaceText(tag,text);
          }}
          
        }
    
    s.getRange(row+2,donecol+1).setValue(nd);
    var ra = s.getRange(row+2, getColByName('PDF', headers)+1).setValue('https://docs.google.com/document/d/' + copyId + '/edit');
    copyDoc.saveAndClose();

      }
  }
}
  

function resizebybounds(cell,image){ // gets cell dimensions in points and converts them to approx pixels
  var w = 600 //currently hardcoded IN PIXELS until i figure out why cell.getWidth won't work
  var h = cell.getParentRow().getMinimumHeight() * 1.33; //cell height in points times an approximate to point conversion
  resize(w,h,image);
}

function resize(maxW,maxH, img) {
  var h = img.getHeight();
  var w = img.getWidth();
  var ratio = h / w;
  if (h > maxH){
	h = maxH;
	w = h / ratio;
	} 
   else if(w > maxW && ratio <= 1) {
    w = maxW;
    h = w * ratio;
  }
  img.setHeight(h)
  img.setWidth(w);
  return img;
}


function isPhoto(text) {
  var exp = /(.+).(jpg|png|gif|bmp)/;
  return exp.test(text);
}

function archiveAndTrue(filename){
  var oldfile = DriveApp.getFileById(idFromFileName(filename));
  oldfile.makeCopy('archived:' + oldfile.getName(),DriveApp.getFolderById(archiveID)); //copy to archive folder and rename with 'archived'
  oldfile.setTrashed(true);
  return true;}

function printOrReprint(created, printed, changed, filename) {
  var c = (filename==''  || printed=='') ;
  if (filename==''  || printed=='') {return true;} // if this has never been printed then print
  else if(changed != undefined && changed > printed){return archiveAndTrue(filename);}//the entry has changed, move folder to archives and 
    
  else {return false;}
  }

function idFromFileName(str){
  return str.replace("https://docs.google.com/document/d/","").replace('/edit','');
}

function getColByName(name, headers){
  var colindex = headers.indexOf(name);
  return parseInt(colindex);
}

function betterDate(date) {
  return Utilities.formatDate(date, "GMT", "MM/dd/yyyyHH:mm").toString();
}