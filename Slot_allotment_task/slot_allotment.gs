//Script for notification on lab allotments

function myfunction_generic(){
  
  var date = new Date();
  var now = date.toLocaleTimeString();

  var formattedDate = Utilities.formatDate(date, "GMT+2", "HH:mm:ss");
  if((date.getDay() == 0)||(date.getDay() == 6)){}
  else{
    //Get Today's date in GMT format
    var todayDate = Utilities.formatDate(new Date(), "GMT+2", "dd.MM.yy");

    //Collection of the data from the spreadsheet "Kopie von Jahresplan_2019_Belegung"  into an array
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2020_Belegung");
    var rows = sheet.getDataRange().getValues();
    //retrieving the last column in the spreadsheet
    
    var RowsCount = 0;
    RowsCount = NumberDependentOn_Day();
    //Loop through every row of the spreadsheet "Kopie von Jahresplan_2019_Belegung" to extract the next three days' data
    rows.forEach(function(row, index)
                 {
                   
                   if(index !== 0)
                   {
                     
                     //Check to see if the column 3 data contains date information
                     if(isValidDate(row[3]))
                     {
                       
                       // translate the date into GMT format
                       var row3 = isDate(row[3]); 
                       
                       //retrieving the row corresponding to the third day from today
                       if(row3 == todayDate)
                       {
                         //Extracting 3 days information into an array called "data"
                         var startRow = index +1//- RowsCount-1; // First row of data to process
                         var numRows = RowsCount; // Number of rows to process
                         var startCol = 1; // first column to process
                         var numCols = 120; //Number of columns to process
                         
                         var dataRange = sheet.getRange(startRow,startCol, numRows,numCols);
                         // Fetch values for each row in the Range.
                         var data = dataRange.getValues();
                         
                         var i = 6; // gives the column index of the requester
                         var j= i+3;  //gives the column index of the requestee
                         var k= j+1; //gives the column index of the prev-mail information
                         
                         var check_for_proceed; //15
                         for (s=0; s < 4;s++){
                          check_for_proceed= extract_funct( sheet,data, startRow,row3,i,j,k);  // function call for Kabinets
                           
                           if (check_for_proceed == 1){
                             
                             Afteraction_kabinetB(sheet,startRow,k,RowsCount,i); 
                           }
                           check_for_proceed = 0;
                           i= k+3
                           j=i+3
                           k=j+1
                         
                         }
                         
                       }
                     }
                   }
                 });

  }
}


function extract_funct(sheet,data, appendRow,date,l,m,n){
  
  var default_id ="volker.uebele@valeo.com"

  var data_mail = mailspreadsheet_func()//function call to create an array of mail ids
  var data_available = 0;
 //data is an array containing information about the lab allotment. 
  mainloop:   for (var i in data ) 
    {
      var rowData = data[i];
      var rowData_index = i;
      
      //on saturday and Sunday, necessary information need not be extracted
      var SATorSUN = rowData[2]
      if(SATorSUN == 'So.' || SATorSUN =='Sa.')
      {
        var Cancel= 'True';
        var times = 0;
      }
      
      if(times == 2)
      {
        Cancel ='False';
      }
      
      
      if(Cancel =='True' && times < 3)
      {
        times = times +1;
        appendRow = appendRow +1 //appendRow tells you the row index in which the data about the mail sent must be appended
        continue mainloop;
      }
      
      //check to see there is a request made. if there is no request continue to the next day or next section of the day i.e pm
      var checkReq= rowData[l]
      
      if( checkReq == 'Cal-1001')
      {
        appendRow=length_strike(sheet,appendRow,n);
        continue mainloop;
      }
      
      //MODIFIED

      var dateData;
       if(isValidDate(rowData[3]))
      {
        dateData = isDate(rowData[3]);
      }
      
      var length_req = checkReq.length
      if(length_req == 0)
      {
        appendRow=length_strike(sheet,appendRow,n)
        continue mainloop;
      }else
      {
        if(isValidDate(rowData[3]))
      {
        dateData = isDate(rowData[3]);
      }
        //Obtain the requester by removing unnecessary information from the column index
        var requestersplit,requester_space,requester_underscore,requester,size,TaskID,requesterSet,blank;
        var requesterInter = rowData[l].split("\n");
        var check =requesterInter[0].length;
        if( requesterInter[0].length < 12)
        {
        TaskID= requesterInter[0];

        requesterSet = requesterInter[requesterInter.length -1]
        
        requestersplit = requesterSet.split("_");
        requester_space = requestersplit[requestersplit.length -1]
        blank = requester_space.search(/ /)
        requester_underscore= requester_space.split(" ")
       
          size = requester_underscore.length -1
        
        requester = requester_underscore[size]
         }else{
           requesterSet =requesterInter[0].split(" ");
           TaskID = requesterSet[0];
           requester = requesterSet[requesterSet.length -1];
        }
        
        var slash = requester.search("/");
        if( slash != -1){
          var sheet_TR = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TR");
          var row_last = sheet_TR.getLastRow();
          var data_TR= sheet_TR.getRange(4, 2, row_last-4, 1).getValues();
          var formated_data = Array_formatting(data_TR);
          var req_data,inter_data;
    
          loopcheck:for(ds =0; ds < formated_data.length;ds++){
            req_data = formated_data[ds].substring(3,13);
            if( req_data == TaskID){
              inter_data = Array_formatting(sheet_TR.getRange(ds+4, 6).getValues());
              requester = inter_data[0];
              break loopcheck;
            }    
          }
        }
        

        //the requestee for kabinet A
        var requestee = rowData[m];
        var strike_status = strikethrough(sheet,appendRow,m);// function call to check if the requestee has been striked out
        if(strike_status == "True"){
          appendRow=length_strike(sheet,appendRow,n)
          continue mainloop;
        }
       
        var requestee_length = requestee.length;
        
        if( requestee_length ==0  )
        {
          appendRow=length_strike(sheet,appendRow,n)//function call if there is no requestee
          continue mainloop;
        }
        
        //Check if mail was sent previously 
        var prevEmail = rowData[n];
        var length_prev = prevEmail.length;
        
        //Enter this section if a mail was sent previosuly. extract the recipients of the previosuly sent mail
        if(length_prev > 0)
        {
          var prevEmail_split = prevEmail.split ("--");
          var prevRequester = prevEmail_split[1]
          var prevRequestee = prevEmail_split[2]
 
          //if the requester and requestee has changed, send a mail again with the new requester and requestee.
          if((requester != prevRequester) || (requestee != prevRequestee))
          {
            var ReqChange = 'True' // flag to indicate a change
            var appenddata = [" "];
            appending_func(sheet, appenddata, appendRow, n);
            }else{
              ReqChange = 'False'
            }
        }else{
          var Req = 'True' // send a mail as previously a mail was not sent
          date_of_request = 'null';
          }        
       
        
        if(ReqChange == 'True' || Req == 'True')
        {
         appendRow= final_step(sheet,data_mail,data,rowData,requester,requestee,appendRow,l,m,n,rowData_index,TaskID,dateData) // function to accumulate data for the mail
         ReqChange ='False'
         Req = 'False'
         data_available =1;
        }else{
          appendRow = appendRow +1
        }
      }  
    }
  return data_available;
}

// Test if value is a date and if so format
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" ){
    return false;}
  return true;
}

//Convert the date to the required format
function isDate(sDate) {
 return Utilities.formatDate(sDate, "GMT+0200", "dd.MM.yy");
}

//function to add data in the prev-mail columns
function appending_func(sheet, appenddata,appendRow,k){
 
  var appending = sheet.getRange(appendRow, k+1, 1, 1);
  appending.setValue(appenddata)
  appending.setBackground('pink')
}

//function to check if the data in the column in striked
function strikethrough(sheet, rowindex, columnindex){
  var strike_range = sheet.getRange(rowindex, columnindex+1)
  var strike_cell = strike_range.getFontLine()

  var strike
  
  if( strike_cell == "line-through"){
    strike = "True"
  }
  return strike
}

//a function to clear existing prev mail information in case the request has been removed or if requestee has been striked out
function length_strike(sheet,appendRow,n){  
  var appenddata=[" "]
  appending_func(sheet, appenddata, appendRow, n)
  appending_func(sheet, appenddata, appendRow, n+1)
  appending_func(sheet, appenddata, appendRow, n+2)
  appendRow = appendRow +1
  return appendRow;
}

//collection of the data from the spreadsheet "Kopie von Lists" that contains mail id informations
function mailspreadsheet_func(){
  var MailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lists");
  var LastRow_mail = MailSheet.getLastRow();
  var dataRange_mail = MailSheet.getRange(3, 14, LastRow_mail, 3);
  var data_mail = dataRange_mail.getValues();

  return data_mail;
}

//function to retrieve the mail ids of the requester and the requestee
function mailid(data_mail,name1,name2){

  var Request= [0,0];
  var namerequest = [name1,name2]
  
  for(var i=0; i<2;i++){

   data_mail.forEach(function(mail){
     
     if( namerequest[i] == mail[0])
     {
       Request[i] = mail[2];                     
     }
   }); 
  }
       return Request;
}

//function to obtain the recipients name
function Recipeint_name(data_mail, name1,name2){
  var Request_name= [0,0];
  var namerequest =[name1,name2]

  
  for(var i=0; i<2;i++){
      data_mail.forEach(function(mail){
        
        if( namerequest[i] == mail[0])
        {
          Request_name[i] = mail[1]
        } 
        
      }); 
  }
  return Request_name;
}

//function to check if the requester or requestee information is available in the spreadsheet
function Recipeint_Recheck(data_mail,name1,name2){
  var Request_recheck = [0,0];
  var namerequest =[name1,name2]

  
   for(var i=0; i<2;i++){

     data_mail.forEach(function(mail){
       
       if( namerequest[i] == mail[0])
       {
         
         Request_recheck[i] = mail[0]
       }  
     }); 
   }
  return Request_recheck;
}


//function to retreive default mail ids, names
function default_set(Request_recheck,request)
{
  var default_id ="volker.uebele@valeo.com" //MAIL ID TO BE CHANGED IN CASE THE ID IS NOT AVAILABLE IN THE KOPIE VON LISTS
  var returnvalue =[0,0,0]
  
  if (Request_recheck !== request){
    returnvalue[0] = 'True'
    returnvalue[1] = default_id
    returnvalue[2] = "Volker Uebele"
  }else{
    returnvalue[0] = 'False'
  }
  
  return returnvalue
}

//function to find the kabinet name
function kabinetname(l){
  var kabinetname;
  
  switch(l)
  {
    case 6:
      kabinetname = 'Kabinet A'
      break;
    case 13:
      kabinetname = 'Kabinet B'
      break;
    case 20:
      kabinetname = 'Kabinet C'
      break;
    case 27:
      kabinetname = 'Kabinet 7'
      break;
    case 34:
      kabinetname = 'Kabinet D'
      break;
    case 41:
      kabinetname = 'Kabinet D'
      break;
    case 48:
      kabinetname = 'Magnetic field'
      break; 
    case 55:
      kabinetname = 'Electric field'
      break; 
    case 62:
      kabinetname = 'MPx PP1'
      break;
    case 69:
      kabinetname = 'MPx PP2'
      break;
    case 76:
      kabinetname = 'MPx ET'
      break;
    case 83:
      kabinetname = 'Kabinet PRE1 LAB'
      break;
    case 90:
      kabinetname = 'Kabinet PRE2 G21'
      break;
    case 97:
      kabinetname = 'Kabinet PRE3 G21'
      break;
    case 97:
      kabinetname = 'Kabinet PRE4 G21'
      break;
    default:
      kabinetname = 'Default Kabinet'
      break;
        
  }
  
  return kabinetname
}

//function to accumulate data for the mail
function final_step(sheet,data_mail,data,rowData,requester,requestee,appendRow,l,m,n,rowData_index,TaskID,dateData){
  
  var default_id ="volker.uebele@valeo.com"  //MAIL ID TO BE CHANGED IN CASE THE ID IS NOT AVAILABLE IN THE KOPIE VON LISTS
  var Requester_id,Requestee_id, Requester_name, Requestee_name,Requester_recheck,Requestee_recheck,Request_id,Request_name,Request_recheck; 
  
  Request_id = mailid(data_mail,requester,requestee)//function call to obtain the mail ids
  Requester_id = Request_id[0]
  Requestee_id= Request_id[1]
  
  
 
  var appenddata = [TaskID+"--"+requester +"--" + requestee+ "--"+Requester_id+"--"+Requestee_id];
  appending_func(sheet, appenddata, appendRow, n+1)
  appenddata =[ dateData]
  appending_func(sheet, appenddata, appendRow, n+2)
  appendRow = appendRow +1;

  return appendRow
  
}

function NumberDependentOn_Day(){
 var date = new Date();
 var Today = date.getDay();
 var result =0;
  
  switch( Today){
    case 1: 
      result = 10;
      break;
    case 2: 
      result= 8;
      break;
    case 3:
      result = 6;
      break;
    case 4:
      //MODIFIED
      result = 4;
      break;
    case 5:
      result = 2;
      break;
  
  }
 return result; 

}

function Sending_mail(StartDate,Content, EndDate,l){
  if(Content != " " ){
    var Content_split = Content.split("--");
    var TaskID = Content_split[0];
    var recipientsTORer = Content_split[3];  //MAIL ID TO BE CHANGED IN CASE THE ID IS NOT AVAILABLE IN THE KOPIE VON LISTS
    var recipientsTORee = Content_split[4];
    var Request_id, Requester_id,Requestee_id,Request_name,Requester_name,Requestee_name,Request_recheck,Requester_recheck,Requestee_recheck

    var default_id ="volker.uebele@valeo.com"
    var data_mail = mailspreadsheet_func()
    
    var kabinet = kabinetname(l);
    var subject = 'Sending email about slot allotment Kabinet:: '+ kabinet;
    var regards = 'Have a pleasant day. <br /><br /> Regards,'
    var sender = 'EMC Team'
    var date ;
    var message1,message2
    var requester = Content_split[1];
    var requestee = Content_split[2];
    Request_id = mailid(data_mail,requester,requestee)//function call to obtain the mail ids
    Requester_id = Request_id[0]
    Requestee_id= Request_id[1]
    
    Request_name =Recipeint_name(data_mail,requester,requestee)//function call to obtain the recipients name
    Requester_name =Request_name[0]
    Requestee_name=Request_name[1]
    
    Request_recheck= Recipeint_Recheck(data_mail,requester,requestee)//function call to check if an entry for the requester and requestee is present in the spreadhsheet 'Lists'
    Requester_recheck=Request_recheck[0]
    Requestee_recheck=Request_recheck[1]
    
    var error_requestee, error_requester;
    var values=[0,0,0];//if the entry is not available set the default ids and names
    values= default_set(Requestee_recheck,requestee)
    if(values[0] == 'True'){Requestee_id = values[1]; Requestee_name = values[2]; error_requestee = 1;}else{ error_requestee = 0;}
    
    values=[0,0,0];
    values = default_set(Requester_recheck,requester)
    if(values[0] == 'True'){Requester_id = values[1]; Requester_name = values[2]; error_requester = 1;}else{ error_requester = 0;}
    
    var start = isDate(StartDate);
    var end = isDate(EndDate);
    
    if(start == end){
      var same_day = 1;
    }
    
    var S1,S2;
    if(Requester_name !=Requestee_name){
      if( error_requester == 0){
        if( same_day == 1){
          message1 ='Hello ' + Requester_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'on '+ ' '+ '<b>'+start+'</b>'+  ' '+'.' +'<br />' +'To make sure that your measurements are scheduled on time, kindly contact '+ '<b>"'+Requestee_name+'</b>"' +' for lab support.' +"<br />"+"<br />"+regards +"<br />"+sender;
          S1 = Requester_id;
        }else{
          message1 ='Hello ' + Requester_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'from '+ ' '+ '<b>'+start+'</b>' + ' ' +"to"+ '<b>'+  ' '+ end+'</b>'+'.' + "<br />"+'To make sure that your measurements are scheduled on time, kindly contact '+ '<b>"'+Requestee_name+'</b>"' +' for lab support.' +"<br />"+"<br />"+regards +"<br />"+sender;
          S1 = Requester_id;
        }
      } else if(error_requester == 1){
        if( same_day == 1){
          message1 = 'Hello ' + Requester_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'on '+ ' '+ '<b>'+start+'</b>'+'.'+'<br />' +'There exists some problem in the Requester, may be a mismatch of the initials in the belegung sheet and Lists, as a result a mail is sent to you. Kindly rectify the problem' +"<br />"+"<br />"+regards +"<br />"+sender;
          S1 =default_id;
        }else{
          message1 = 'Hello ' + Requester_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'from '+ ' '+ '<b>'+start+'</b>'+' ' + "to"+ '<b>'+ ' '+ end+'</b>'+'.' + "<br />"+'There exists some problem in the Requestee may be a mismatch of the initials in the belegung sheet and Lists, as a result a mail is sent to you. Kindly rectify the problem' +"<br />"+"<br />"+regards +"<br />"+sender;
          error_requester =0;
          S1 =default_id;
        }
      }
      
      if( error_requestee == 0){
        if( same_day == 1){
          message2 = 'Hello ' + Requestee_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'on '+ ' '+ '<b>'+start+'</b>' +'.'+'<br />' +'Kindly contact '+ '<b>"'+Requester_name+'</b>"' +' for lab details.' +"<br />"+"<br />"+regards +"<br />"+sender ;
          S2= Requestee_id;
        }else{
          message2 = 'Hello ' + Requestee_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'from '+ ' '+ '<b>'+start+'</b>' +' ' + "to"+ '<b>'+ ' '+ end +'</b>'+'.' + "<br />"+'Kindly contact '+ '<b>"'+Requester_name+'</b>"' +' for lab details.' +"<br />"+"<br />"+regards +"<br />"+sender ;
          S2= Requestee_id;
        }
      }else if(error_requestee == 1){
        if( same_day == 1){
          message2 = 'Hello ' + Requestee_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'on '+ ' '+ '<b>'+start+'</b>' +' ' +  '.' +'<br />'+'There exists some problem in the Requester may be a  be a mismatch of the initials in the belegung sheet and Lists, as a result a mail is sent to you. Kindly rectify the problem '+"<br />"+"<br />"+regards +"<br />"+sender ;
          S2 =default_id;
        }else{
          message2 = 'Hello ' + Requestee_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'for '+ ' '+ '<b>'+start+'</b>' +' ' + "to"+ '<b>'+  ' '+end + '</b>'+'.' +"<br />"+'There exists some problem in the Requester may be a  be a mismatch of the initials in the belegung sheet and Lists, as a result a mail is sent to you. Kindly rectify the problem '+"<br />"+"<br />"+regards +"<br />"+sender ;
          error_requestee =0;
          S2 =default_id;
        }
      }
       
  same_day =0;
  //cc: default_id 
 
  MailApp.sendEmail({to: S1,subject: subject,htmlBody: message1}); //send mail to requester
  MailApp.sendEmail({to: S2,subject: subject,htmlBody: message2}); //send mail to requestee
      
    }else
    {
      if( error_requester == 0){
        if( same_day == 1){
          message1 ='Hello ' + Requester_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'on '+ ' '+ '<b>'+start+'</b>'+  ' '+'.' +'<br />' +"<br />"+"<br />"+regards +"<br />"+sender;
          S1 = Requester_id;
        }else{
          message1 ='Hello ' + Requester_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'from '+ ' '+ '<b>'+start+'</b>' + ' ' +"to"+ '<b>'+  ' '+ end+'</b>'+'.' + "<br />"+"<br />"+"<br />"+regards +"<br />"+sender;
          S1 = Requester_id;
        }
      } else if(error_requester == 1){
        if( same_day == 1){
          message1 = 'Hello ' + Requester_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'on '+ ' '+ '<b>'+start+'</b>'+'.'+'<br />' +'There exists some problem in the Requestee may be a mismatch of the initials in the belegung sheet and Lists, as a result a mail is sent to you. Kindly rectify the problem' +"<br />"+"<br />"+regards +"<br />"+sender;
          S1 =default_id;
        }else{
          message1 = 'Hello ' + Requester_name + ','+ '<br />'+ '<br />' + 'The lab request is made as per '+ '"<b><i>'+TaskID+'</i></b>"' +' '+ 'from '+ ' '+ '<b>'+start+'</b>'+' ' + "to"+ '<b>'+ ' '+ end+'</b>'+'.' + "<br />"+'There exists some problem in the Requestee may be a mismatch of the initials in the belegung sheet and Lists, as a result a mail is sent to you. Kindly rectify the problem' +"<br />"+"<br />"+regards +"<br />"+sender;
          error_requester =0;
          S1 =default_id;
        }
      }

     MailApp.sendEmail({to:S1 ,subject: subject,htmlBody: message1}); //send mail to requester
    }
  }
}

function Compare(array2) {
  var result_index =[];
 
  
  var i=0;
  result_index[0]=0;
  var a =0;
  var j=0;
  
  do{
    
    if(array2[a] != array2[i]){ 
      j= j+1;
    result_index[j]= i;
      
      a = i;
    }
    
    i= i+1;
  }while(i != array2.length);
  
  result_index[result_index.length] = array2.length ;
  return result_index;

}

function Monday(array1,dateColumn,l,sheet,startRow){

  var array_size = array1.length;
  var result_firstLevel = Compare(array1);
  var Inter_date=[]
  
  var array_sheet=[];
  var i=0;
  var check =0;
  var a=0;
 
  do{
    a= array1.length;
    i = i+1;
    if( result_firstLevel[1] ==0){
    
     array_sheet.push(array1.splice(0, 1));
      Inter_date.push(dateColumn.splice(0, 1));
    }else{
    
    array_sheet.push(array1.splice(0, result_firstLevel[i]-check));
    Inter_date.push(dateColumn.splice(0, result_firstLevel[i]-check));  
    }
    check = check+ a-array1.length;
    
  }while (array1.length != 0) 


  i=0;
  var Content,Startdate,EndDate;
  var appendRow, appenddata,inc;
  appendRow = startRow;
  
  do{
    //MODIFIED
    if(array_sheet[i][0] != " "){
      Content = array_sheet[i][0];
      Startdate =Inter_date[i][0];
      EndDate =Inter_date[i][Inter_date[i].length-1]
       Sending_mail(Startdate,Content, EndDate,l)
      
      appenddata = Content;
      for(g=0;g < array_sheet[i].length; g++){
        appending_func(sheet, appenddata, appendRow, l+4)
        appendRow = appendRow +1;
      }
  
  i=i+1;
    }else
    {
     appendRow= appendRow + array_sheet[i].length;
      i=i+1;
    }
  }while(i != array_sheet.length);
  
}

function OtherDays(array1,array2,dateColumn,l,sheet,startRow,array5,array6,array7){
 
  
   var array3 =[];
  for(m=0;m< array1.length;m++){
    array3[m]=array1[m];
  }
  var array_size = array1.length;
  var result_firstLevel = Compare(array2);
  var Inter_date=[]
  
  var array_prevmail=[];
  var array_sheet=[];
  var i=0;
  var check =0;
  var a=0;
 
  do{
    a= array2.length;
    i = i+1;
    if( result_firstLevel[1] ==0){
      array_prevmail.push(array3.splice(0, 1))
      array_sheet.push(array2.splice(0, 1));
      Inter_date.push(dateColumn.splice(0, 1));
    }else{
      array_prevmail.push(array3.splice(0, result_firstLevel[i]- check))
      array_sheet.push(array2.splice(0, result_firstLevel[i]-check));
      Inter_date.push(dateColumn.splice(0, result_firstLevel[i]-check));  
    }
    check =check+ a-array2.length;
    
  }while (array2.length != 0) 

 
    a=0;
    i=0;
   var array1= [];
    array2.length=0;

    do{
     
     array1.push( second_level_comparison(array_prevmail[i]));
     
     i=i+1; 
      
    
    }while(i!= array_prevmail.length );
  i=0;
  var Content,Startdate,EndDate;
  var appendRow, appenddata,inc,start_index,end_index,end_check;
  inc =0;
  var hole=0;
  var btw=0;
  appendRow = startRow
  do{
    if( array1[i].length != 2){
      Content = array_sheet[i][0];
      
      Startdate =Inter_date[i][array1[i][1]];
      
      EndDate =Inter_date[i][array1[i][array1[i].length-2]]
      Sending_mail(Startdate,Content, EndDate,l)
      loopfor1:for(v=btw;v<array5.length;v++){
            if((array6[inc] == array5[v])&& array7[v] == " "){
              btw = v+1;
             break loopfor1;
            }else{
            hole= hole+1
            }
          }    
      appendRow = appendRow+hole;
      hole=0;
      appenddata = Content;
      for(g=0;g < array_sheet[i].length; g++){
        appending_func(sheet, appenddata, appendRow, l+4)
        
        if(g != array_sheet[i].length -1){
        loopfor2:for(v=btw;v<array5.length;v++){
            if((array6[g+1] == array5[v])&& array7[v] == " "){
              btw =v+1
             break loopfor2;
            }else{
            hole= hole+1
            }
          }
        }
        appendRow = appendRow +1+hole;
        hole =0;
      }
      inc = inc + array_sheet[i].length;
    }else{
    inc = inc + array_sheet[i].length;
    }
  
  i=i+1;
  }while(i != array1.length);
  

}

function second_level_comparison(array2) {
  var result_index =[];

  var i=0;
  result_index[0]=0;
  var a =0;
  var j=0;
  
    do{
      
      if(array2[i] == " "){ 
        j= j+1;
        result_index[j]= i;
      }
      
      i= i+1;
    }while(i != array2.length);

  result_index[result_index.length]= array2.length
  return result_index;
  
}

//fow kabinet 
function Afteraction_kabinetB(sheet,startRow,k,RowsCount,l,data_mail){

  var date = new Date();
  var Today = date.getDay();
  var array1=[];
  var array1_range, array3_range,array4_range,array5_range;
  var array2=[];
  var array3=[];
  var array4=[];
  var array5=[];
  var inter_1 =[];
  var inter_3 =[];
  var i =0;
  
  var dateColumn=[];
  var dateColumn_range;
  var dateColumn_inter=[];
  dateColumn_range= sheet.getRange(startRow,k+3, RowsCount,1); 
  dateColumn = dateColumn_range.getValues();
  do{
    dateColumn_inter[i] = dateColumn[i][0];
      i = i+1;
    }while(i != dateColumn.length)
  dateColumn.length =0;
  dateColumn = dateColumn_inter; 
   
  i=0
  var date = new Date();
  var now = date.toLocaleTimeString();
  
  var formattedDate = Utilities.formatDate(date, "GMT+0200", "HH:mm:ss");

  if(Today == 1 && formattedDate < "07:30:00"){ 
    dateColumn =[]
    array1_range= sheet.getRange(startRow,k+2, RowsCount,1); 
    array1 = array1_range.getValues();
    do{
      
        inter_1[i] = array1[i][0];  
       
    
      i = i+1;
    }while(i != array1.length)
      
    array1.length =0;
    array1 = inter_1;
    //am and pm seperately
    
    array3 = array1.filter(function(arra1,index){
      if(index%2 == 0){
       return arra1;
      }
    })
    
    array4 = array1.filter(function(arra1,index){
      if(index%2 != 0){
       return arra1;
      }
    })
    var flag_consecutive=0;
    loopbreak:for(m=0;m+1< array1.length;m++){
      if((array1[m] == array1[m+1])&& (dateColumn_inter[m+1]-dateColumn_inter[m] <= 1)&& (array1[m] != " ")&&(array1[m].length != 0)){
       
       flag_consecutive =1;
        break loopbreak;
      }
    }
    dateColumn = dateColumn_inter.filter(function(date_arr,index){
     if(index %2 == 0) return date_arr
    })
    
    var dateColumn11= dateColumn_inter.filter(function(date_arr,index){
     if(index %2 != 0) return date_arr 
    })

    if(flag_consecutive == 0){
     array2 = Monday_kabinetB(array3,array4,dateColumn,l,sheet,startRow,dateColumn11); 
    }else{
     
    array2 = Monday(array1,dateColumn_inter,l,sheet,startRow); 
    }
    //mail and appending function
    
  }else{
    
    array1_range = sheet.getRange(startRow,k+1, RowsCount,1); //prev mail
    var dateColumn_inter11
    var flag_blank1= 0;
    var flag_blank3 =0;

    var array11 = array1_range.getValues();
    array3_range = sheet.getRange(startRow,k+2, RowsCount,1); // new data
    var array33 =array3_range.getValues();
    var array3_inter =array3_range.getValues();
  
    
    if(formattedDate > "10:00:00"){
     array1 = array11.splice(1, array11.length);
     array3 = array33.splice(1,array33.length);
      dateColumn_inter11 = dateColumn_inter.splice(1, dateColumn_inter.length)
      
      startRow = startRow+1
    }else{
     array1 = array11;
     array3 = array33;
      dateColumn_inter11 = dateColumn_inter
      
    }
 
    var inter_array3=[]
    i=0;
     do{
        inter_array3[i] = array3_inter[i][0];  
      
      i = i+1;
    }while(i != array3_inter.length)
      
    array3_inter.length =0;
    array3_inter = inter_array3;
    i =0;
    
    
    
    do{
        inter_1[i] = array1[i][0];  
      if(inter_1[i] == " "){
        flag_blank1 = flag_blank1+1;
      }
      i = i+1;
    }while(i != array1.length)
      
    array1.length =0;
    array1 = inter_1;
    i =0;
  
    
    do{
        inter_3[i] = array3[i][0];  
        if(inter_3[i] == " "){
        flag_blank3 = flag_blank3+1;
      }
      i = i+1;
    }while(i != array3.length)
    array3.length =0;
    array3 = inter_3;
    
    inter_1 = array1.filter(function(arra1,index){
        if( index %2 == 0){
          return arra1
        }
      })
    inter_3 = array3.filter(function(arra1,index){
        if( index %2 == 0){
          return arra1
        }
      })
      
    array4 = array1.filter(function(arra1,index){
        if( index %2 != 0){
          return arra1
        }
      })
     
    array5=array3.filter(function(arra3,index){
        if( index%2 != 0){
         return arra3
        }
      })
   

 var arr=[]
   arr.push(second_level_comparison(array1));
    var flag_consecutive =0;
    for(j=1;j+1<= arr[0].length-2;j++){//condition update
      if((arr[0][j+1]- arr[0][j] <=3)&&(array3[arr[0][j+1]]==array3[arr[0][j]])&& (dateColumn_inter11[arr[0][j+1]] != " ")&&(dateColumn_inter11[arr[0][j]] != " ")&&
        (dateColumn_inter11[arr[0][j+1]].getDay()-dateColumn_inter11[arr[0][j]].getDay() <= 1) ){
        flag_consecutive =1;
        break;
      }      
    }
    var rr=[], cc=[],ll=[];
    var dateColumn_inter12=[]
    if(flag_consecutive == 1){
      
      for(j=0;j+1<= arr[0].length-2; j++){
        rr[j]=array3[arr[0][j+1]];
        cc[j]=array1[arr[0][j+1]]; 
        ll[j]=array3[arr[0][j+1]];
        dateColumn_inter12[j]= dateColumn_inter11[arr[0][j+1]];
      }
    }
        
    if( flag_blank1 == flag_blank3 ){}else{
      if((flag_consecutive ==0)){
      OtherDays_kabinetB(inter_1,inter_3,array4,array5,dateColumn_inter11,l,sheet,startRow);
      }else{
      OtherDays(cc,rr,dateColumn_inter12,l,sheet,startRow,array3,ll,array1);
      }
  }
    flag_blank1 =0;
    flag_blank3 =0;
  }


}

function compare2arrays(array3,array4){
 var value=0;
  for( i=0; i< array3.length;i++){
   if(array3[i] != array4[i])
   {
    value = 1;
   }
  }
 return value;
}
function Monday_kabinetB(array1,array2,dateColumn,l,sheet,startRow,dateColumn11){
  var array_size = array1.length;
  var result_firstLevel = Compare(array1);
  var Inter_date=[]
  var Inter_date2=[];
  
  var array_sheet1=[];
  var i=0;
  var check =0;
  var a=0;
 
  do{
    a= array1.length;
    i = i+1;
    if( result_firstLevel[1] ==0){
    
     array_sheet1.push(array1.splice(0, 1));
      Inter_date.push(dateColumn.splice(0, 1));
    }else{
    
    array_sheet1.push(array1.splice(0, result_firstLevel[i]-check));
    Inter_date.push(dateColumn.splice(0, result_firstLevel[i]-check));  
    }
    check = check+ a-array1.length;
    
  }while (array1.length != 0) 

  var array2_size= array2.length;
  var result_firstLevel2 = Compare(array2);
  
  var array_sheet2=[];
  i=0;
  check =0;
  a=0;
 
  do{
    a= array2.length;
    i = i+1;
    if( result_firstLevel2[1] ==0){
    
     array_sheet2.push(array2.splice(0, 1));
     Inter_date2.push(dateColumn11.splice(0, 1));
    }else{
    
    array_sheet2.push(array2.splice(0, result_firstLevel2[i]-check));
    Inter_date2.push(dateColumn11.splice(0, result_firstLevel2[i]-check)); 
    }
    check = check+ a-array2.length;
    
  }while (array2.length != 0) 

  i=0;
  var Content,Startdate,EndDate;
  var appendRow, appenddata,inc;
  appendRow = startRow;
  
  do{
    if( array_sheet1[i][0] != " "){
      Content = array_sheet1[i][0];
      Startdate =Inter_date[i][0];
      EndDate =Inter_date[i][Inter_date[i].length-1]
       Sending_mail(Startdate,Content, EndDate,l) 
      
      appenddata = Content;
      for(g=0;g < array_sheet1[i].length; g++){
        appending_func(sheet, appenddata, appendRow, l+4)
        appendRow = appendRow +2;
      }
  
      i=i+1;
    }else{
     appendRow = appendRow + 2*(array_sheet1[i].length); 
      i=i+1;
    }
  }while(i != array_sheet1.length);
  
  appendRow = startRow+1;
  i=0;
  do{

    if( array_sheet2[i][0] != " "){
      Content = array_sheet2[i][0];
      Startdate =Inter_date2[i][0];
      EndDate =Inter_date2[i][Inter_date2[i].length-1]
       Sending_mail(Startdate,Content, EndDate,l)
      
      appenddata = Content;
      for(g=0;g < array_sheet2[i].length; g++){
        appending_func(sheet, appenddata, appendRow, l+4)
        appendRow = appendRow +2;
      }
  
      i=i+1;
    }else{
      appendRow = appendRow + 2*(array_sheet2[i].length);
      i=i+1;
    }
  }while(i != array_sheet2.length);
 
}

function OtherDays_kabinetB(array1,array2,array4,array5,dateColumn_inter,l,sheet,startRow){
  var dateColumn = dateColumn_inter.filter(function(date_arr,index){
     if(index %2 == 0) return date_arr
    })

 kabinetB_check(array1,array2,dateColumn,l,sheet,startRow)
 
 dateColumn = dateColumn_inter.filter(function(date_arr,index){
     if(index %2 != 0) return date_arr
    })
 var inter
 if(dateColumn.length != array4.length ){
   inter = dateColumn.splice(1,dateColumn.length);
  }else{
   inter = dateColumn;
  }
  
 kabinetB_check(array4,array5,inter,l,sheet,startRow+1)

}

function kabinetB_check(array1,array2,dateColumn_inter,l,sheet,startRow){
 var array3 =[];
  array3 = array1;
  var array_size = array1.length;
  var result_firstLevel = Compare(array2);
  var dateColumn = dateColumn_inter
  var Inter_date=[]
  
  var array_prevmail=[];
  var array_sheet=[];
  var i=0;
  var check =0;
  var a=0;
 
  do{
    a= array2.length;
    i = i+1;
    if( result_firstLevel[1] ==0){
      array_prevmail.push(array3.splice(0, 1))
      array_sheet.push(array2.splice(0, 1));
      Inter_date.push(dateColumn.splice(0, 1));
    }else{
      array_prevmail.push(array3.splice(0, result_firstLevel[i]- check))
      array_sheet.push(array2.splice(0, result_firstLevel[i]-check));
      Inter_date.push(dateColumn.splice(0, result_firstLevel[i]-check));  
    }
    check =check+ a-array2.length;
    
  }while (array2.length != 0) 

 
    a=0;
    i=0;
    array1= [];
    array2.length=0;

    do{
     
     array1.push(second_level_comparison(array_prevmail[i]));
     
     i=i+1; 
      
    
    }while(i!= array_prevmail.length );
  i=0;
  var Content,Startdate,EndDate;
  var appendRow, appenddata,inc;
  inc =0;

  do{
    if( array1[i].length != 2){
      Content = array_sheet[i][0];
      Startdate =Inter_date[i][array1[i][1]];
      EndDate =Inter_date[i][array1[i][array1[i].length-2]]
       Sending_mail(Startdate,Content, EndDate,l)
      appendRow = startRow+2*inc;
      appenddata = Content;
      for(g=0;g < array_sheet[i].length; g++){
        appending_func(sheet, appenddata, appendRow, l+4)
        appendRow = appendRow +2;
      }
      inc = inc + array_sheet[i].length;
    }else{
    inc = inc + array_sheet[i].length;
    }
  
  i=i+1;
  }while(i != array1.length);
}

function Array_formatting(input_data){
 var temp=[];
  var i=0;
 do{
    temp[i] = input_data[i][0];
      i = i+1;
    }while(i != input_data.length)
   return temp;
}
