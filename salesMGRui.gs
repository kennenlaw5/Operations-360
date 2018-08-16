function salesMGRjeff() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveRange(ss.getRange("A6"));
}
function salesMGRben() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveCell().getRow();
  if(row<40){ss.setActiveRange(ss.getRange("B65"));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A40"));
}
function salesMGRrobb() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveCell().getRow();
  if(row<74){ss.setActiveRange(ss.getRange("B99"));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A74"));
}
function salesMGRanna() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveCell().getRow();
  if(row<108){ss.setActiveRange(ss.getRange("B134"));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A108"));
}
function salesMGRseth() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveCell().getRow();
  if(row<142){ss.setActiveRange(ss.getRange("B168"));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A142"));
}
function salesMGRdean() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var row = ss.getActiveCell().getRow();
  if(row<176){ss.setActiveRange(ss.getRange("B201"));ss.getActiveRange().getValue();}
  ss.setActiveRange(ss.getRange("A176"));
}
function importInfo(day){
  //Created By Kennen Lawrence
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var source1=ss.getSheetByName("Report");
  var source2=ss.getSheetByName("Report (1)");
  var sourceTemp;
  var target=ss.getSheetByName("GSM");
  var range1=source1.getRange(1,1,parseInt(source1.getLastRow())).getValues();
  var avg1=0;var avg2=0;
  var numRows1=0;var numCol1=0;var numCol2=0;var col1=0;var col2=0;
  var jeff=[0,0,0,0];
  var ben=[0,0,0,0];
  var robb=[0,0,0,0];
  var anna=[0,0,0,0];
  var seth=[0,0,0,0];
  var dean=[0,0,0,0];
  var teams=[7,41,75,109,143,177];
  for(var i=0;i<source1.getLastRow();i++){
    if(range1[i]==""){i=source1.getLastRow()}else{numRows1+=1;}
  }
  for(i=0;i<source1.getLastColumn();i++){
    if(range1[0][i]==""){i=source1.getLastColumn()}else{numCol1+=1;}
  }
  range1=source2.getRange(1,1,parseInt(source2.getLastRow())).getValues();
  var numRows2=0;
  for(i=0;i<source2.getLastRow();i++){
    if(range1[i]==""){i=source2.getLastRow()}else{numRows2+=1;}
  }
  for(i=0;i<source2.getLastColumn();i++){
    if(range1[0][i]==""){i=source2.getLastRow()}else{numCol2+=1;}
  }
  range1=source1.getRange(1,1,numRows1,numCol1).getValues();
  var range2=source2.getRange(1,1,numRows2,numCol2).getValues();
  for(i=1;i<numRows1;i++){
    for(var j=1;j<numCol1;j++){
      avg1+=parseInt(range1[i][j]);
    }
  }
  for(i=1;i<numRows2;i++){
    for(var j=1;j<numCol2;j++){
      avg2+=parseInt(range2[i][j]);
    }
  }
  avg1=avg1/(numRows1*numCol1);
  avg2=avg2/(numRows2*numCol2);
  if(avg2<avg1){sourceTemp=source2;source2=source1;source1=sourceTemp;sourceTemp=numRows1;numRows1=numRows2;numRows2=sourceTemp;sourceTemp=range1;range1=range2;range2=sourceTemp;sourceTemp=numCol1;numCol1=numCol2;numCol2=sourceTemp;}
  //Check for colNumbers
  for(i=0;i<numCol1;i++){
    if(range1[0][i]=="Warranty"){col1=parseInt(i);}
    if(range1[0][i]=="Product"){col2=parseInt(i);}
  }
  for(i=0;i<numRows1;i++){
    if(range1[i][0]=="Jeffery Englert"){if(col1!=0){jeff[0]+=range1[i][col1];}if(col2!=0){jeff[1]+=range1[i][col2];}}
    else if(range1[i][0]=="Ben Brahler"){if(col1!=0){ben[0]+=range1[i][col1];}if(col2!=0){ben[1]+=range1[i][col2];}}
    else if(range1[i][0]=="Robb Ashby"){if(col1!=0){robb[0]+=range1[i][col1];}if(col2!=0){robb[1]+=range1[i][col2];}}
    else if(range1[i][0]=="Anna Wright"){if(col1!=0){anna[0]+=range1[i][col1];}if(col2!=0){anna[1]+=range1[i][col2];}}
    else if(range1[i][0]=="Seth Carmitchel-Ewing"){if(col1!=0){seth[0]+=range1[i][col1];}if(col2!=0){seth[1]+=range1[i][col2];}}
    else if(range1[i][0]=="Alan Wentland"||range1[i][0]=="Mark Sanders"){if(col1!=0){dean[0]+=range1[i][col1];}if(col2!=0){dean[1]+=range1[i][col2];}}
  }
  Logger.log(numCol1+"\n"+numCol2);
  for(i=0;i<numCol2;i++){
    if(range2[0][i]=="Product"){col1=parseInt(i);}
    if(range2[0][i]=="Product<br> & Reserve"){col2=parseInt(i);}
  }
  for(i=0;i<numRows2;i++){
    if(range2[i][0]=="Jeffery Englert"){if(col1!=0){jeff[2]+=range2[i][col1];}if(col2!=0){jeff[3]+=range2[i][col2];}}
    else if(range2[i][0]=="Ben Brahler"){if(col1!=0){ben[2]+=range2[i][col1];}if(col2!=0){ben[3]+=range2[i][col2];}}
    else if(range2[i][0]=="Robb Ashby"){if(col1!=0){robb[2]+=range2[i][col1];}if(col2!=0){robb[3]+=range2[i][col2];}}
    else if(range2[i][0]=="Anna Wright"){if(col1!=0){anna[2]+=range2[i][col1];}if(col2!=0){anna[3]+=range2[i][col2];}}
    else if(range2[i][0]=="Seth Carmitchel-Ewing"){if(col1!=0){seth[2]+=range2[i][col1];}if(col2!=0){seth[3]+=range2[i][col2];}}
    else if(range2[i][0]=="Alan Wentland"||range2[i][0]=="Mark Sanders"){if(col1!=0){dean[2]+=range2[i][col1];}if(col2!=0){dean[3]+=range2[i][col2];}}
  }
  var order=[jeff,ben,robb,anna,seth,dean];
  var row=0;
  for(var k in teams){
    row=parseInt(teams[k])+parseInt(day);
    target.getRange(row,7,1,4).setValues([order[k]]);
  }
  Logger.log(order);
}
function performance(day){
  //Created By Kennen Lawrence
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var source=ss.getSheetByName("Performance");
  var target=ss.getSheetByName("GSM");
  var numRows=source.getLastRow();var numCol=source.getLastColumn();
  var range=source.getRange(1,1,numRows,numCol).getValues();
  var rows=0;var col=0;var name;var current;var team;
  var split;var type;var stock;var found=false;var j;var k;
  var teamRows=[7,41,75,109,143,177];
  var jeff=[0,0,0];
  var ben=[0,0,0];
  var robb=[0,0,0];
  var anna=[0,0,0];
  var seth=[0,0,0];
  var dean=[0,0,0];
  
  var tjeff=["Kiersten Peterson","Brian Neal","Jonthan Wingfield","Omar Johnson","Jeremy Sanchez","Ian Hudgens","Roger Surroz"];
  var tben=["Demitri Gavito","Patrick Quinlan","Tony Moomau","Karen Timmons","Troy","Stephen Giese"];
  var trobb=["Agymang Spencer","Jacob Ford","Kathy Powell","Chris Castro","Jeffrey Tucker","Conner Graves"];
  var tanna=["Sam Nejad","Connor Hanlon","Ace Taylor-Brown","Jenny Kim","Andrew Sapoznik","Erin Vangilder"];
  var tseth=["Jeffrey Hanson","Chuck Northrup","Christopher Leirer","Alexander Duquette","Marlowe Jones","Shaun Welch","Craig Smeton"];
  var tdean=["Tim Green","Ben Wegener","Josh Ackerman"];
  
  var teamCA=[tjeff,tben,trobb,tanna,tseth,tdean];
  var teams=[jeff,anna,robb,anna,seth,dean];
  for(var i=0;i<numRows;i++){
    if(range[i][0]==""){i=numRows;}else{rows+=1;}
  }
  numRows=rows;
  for(i=0;i<numCol;i++){
    if(range[0][i]==""){i=numCol;}
    else if(range[0][i]=="Split"){split=parseInt(i);col+=1;}
    else if(range[0][i]=="Type"){type=parseInt(i);col+=1;}
    else if(range[0][i]=="Stock#"){stock=parseInt(i);col+=1;}
    else{col+=1;}
  }
  range=source.getRange(1,1,numRows,numCol).getValues();
  for(i=1;i<numRows;i++){
    if(!isNaN(parseInt(range[i][0]))){
      if(found==false){
        //Determine "name's" team
        for(j=0;j<teamCA.length;j++){
          for(k=0;k<teamCA[j].length;k++){
            if(teamCA[j][k]==name){team=parseInt(j);found=true;k=teamCA[j].length-1;j=teamCA.length-1;Logger.log("FOUND");}
          }
        }
      }
      if(found==true){
        if(range[i][type]=="New"){
          teams[team][0]+=range[i][split];
        }else if(range[i][type]=="Pre-Owned"){
          if(range[i][stock][2]!="L"){teams[team][1]+=range[i][split];}
          else if(range[i][stock][2]=="L"){teams[team][2]+=range[i][split];}
        }else{Logger.log("Bad Type; Row: "+i+"\nType: "+range[i][type]);}
      }else{Logger.log("CA Not found; Name: "+name);}
    }
    else{
      name=range[i][0].split("<strong>");
      name=name[1];
      name=name.split("</strong>");
      name=name[0];
      Logger.log(name);
      found=false;
      //return 
    }
  }
  var row=0;
  for(var k in teamRows){
    row=parseInt(teamRows[k])+parseInt(day);
    target.getRange(row,4,1,3).setValues([teams[k]]);
  }
  Logger.log(teams);
}
function fullImport(){
  //Created By Kennen Lawrence
  var ui=SpreadsheetApp.getUi();
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("GSM"));
  var input = ui.prompt('Day to import:','Please type the day being imported:',ui.ButtonSet.OK_CANCEL);
  if (input.getSelectedButton() == ui.Button.OK) {
    var day=parseInt(input.getResponseText());
    Logger.log(day);
    performance(day);
    importInfo(day);
  } else { 
    ss.toast("Import cancelled! No data was modified!", "Import Cancelled")
    return; 
  }
  var source=ss.getSheetByName("Performance");
  var source1=ss.getSheetByName("Report");
  var source2=ss.getSheetByName("Report (1)");
  ss.deleteSheet(source1);
  ss.deleteSheet(source2);
  ss.deleteSheet(source);
  ss.toast('Import completed successfully!','Import Complete!',5);
}