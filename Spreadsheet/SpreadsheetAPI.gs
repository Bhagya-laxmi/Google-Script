function myfunction_generic(){
  //Get Today's date in GMT format
  var todayDate = Utilities.formatDate(new Date(), "GMT+1", "dd.MM.yy");
  //Get the date of the third day from today in GMT format
  var thirdDay = new Date();
  thirdDay.setDate(thirdDay.getDate()+3);
  thirdDay = Utilities.formatDate(thirdDay, "GMT+1", "dd.MM.yy");
  
  
  //Collection of the data from the spreadsheet "Kopie von Jahresplan_2019_Belegung"  into an array
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kopie von Jahresplan_2019_Belegung");
  var rows = sheet.getDataRange().getValues();
  //retrieving the last column in the spreadsheet
  var Lastcols = sheet.getLastColumn();
 
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
         if(row3 == thirdDay)
         {
           //Extracting 3 days information into an array called "data"
           var startRow = index -3; // First row of data to process
           var numRows = 6; // Number of rows to process
           var startCol = 1; // first column to process
           var numCols = Lastcols; //Number of columns to process
           
           var dataRange = sheet.getRange(startRow,startCol, numRows,numCols);
           // Fetch values for each row in the Range.
           var data = dataRange.getValues();
           var i = 6; // gives the column index of the requester
           var j= 9;  //gives the column index of the requestee
           var k= Lastcols -13; //gives the column index of the prev-mail information
             
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k);  // function call for Kabinet A
       
           i= i+5
           j=j+5
           k=k+1
       
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for Kabinet B
      
           i= i+5
           j=j+5
           k=k+1
       
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for Kabinet C
          
           i= i+5
           j=j+5
           k=k+1
       
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for Kabinet D
       
           i=i+5
           j=j+5
           k=k+1
   
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for Kabinet 7
    
           i= i+5
           j=j+5
           k=k+1
     
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for MPx PP1
           
           i= i+5
           j=j+5
           k=k+1
    
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for MPx PP2
        
           i= i+5
           j=j+5
           k=k+1
    
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for MPx ET
    
           i= i+5
           j=j+5
           k=k+1
    
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for MPx Helmholz
    
           i= i+5
           j=j+5
           k=k+1
     
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for PRE1 LAB
   
           i= i+5
           j=j+5
           k=k+1
    
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for PRE2 LAB
    
           i= i+5
           j=j+5
           k=k+1
 
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k);//function call for PRE3 LAB
          
           i= i+5
           j=j+5
           k=k+1
           
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for PRE4 G21
           
           i= i+5
           j=j+5
           k=k+1
           
           extract_funct( sheet,data, startRow,Lastcols,row3,i,j,k); //function call for PRE5 G21
           
           
         }
       }
     }
   });
             
  Browser.msgBox('End generic function')
  
}




function extract_funct(sheet,data, appendRow,Lastcols,date,l,m,n){
  
  var default_id ="bhagyalaxmi.dinesh@valeo.com"

  var data_mail = mailspreadsheet_func()//function call to create an array of mail ids

 //data is an array containing information about the lab allotment. 
  mainloop:   for (var i in data ) 
    {
      var rowData = data[i];
      if(isValidDate(rowData[1]))
      {
        var dateData = isDate(rowData[1]);
      }
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
      var length_req = checkReq.length
      if(length_req == 0)
      {
        appendRow=length_strike(sheet,appendRow,n)
        continue mainloop;
      }else
      {
        //Obtain the requester by removing unnecessary information from the column index
        var requesterInter = rowData[l].split("n");
        var requesterSet = requesterInter[requesterInter.length -1]
        var requestersplit,requester_space,requester_underscore,requester,size;
       
        requestersplit = requesterSet.split("_");
        requester_space = requestersplit[requestersplit.length -1]
        var blank = requester_space.search(/ /)
        requester_underscore= requester_space.split(" ")
        if(blank == requester_underscore.length){
          size = requester_underscore.length -2
        }else
        {
          size = requester_underscore.length -1
        }
        requester = requester_underscore[size]

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
        var prevEmail = rowData[n -1];
        var length_prev = prevEmail.length;
        
        //Enter this section if a mail was sent previosuly. extract the recipients of the previosuly sent mail
        if(length_prev > 0)
        {
          var prevEmail_split = prevEmail.split ("--");
          var prevRequester = prevEmail_split[0]
          var prevRequestee = prevEmail_split[prevEmail_split.length-1]
          //if the requester and requestee has changed, send a mail again with the new requester and requestee.
          if((requester != prevRequester) || (requestee != prevRequestee))
          {
            var ReqChange = 'True' // flag to indicate a change
            }else{
              ReqChange = 'False'
            }
        }else{
          var Req = 'True' // send a mail as previously a mail was not sent
          }        
       
        if(ReqChange == 'True' || Req == 'True')
        {
         appendRow= final_step(sheet,data_mail,rowData,requester,requestee,appendRow,l,m,n) // function to accumulate data for the mail
         ReqChange ='False'
         
        }else{
          appendRow = appendRow +1
        }
      }  
    }
}

// Test if value is a date and if so format
function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" ){
    return false;}
  return true;
}

//Convert the date to the required format
function isDate(sDate) {
 return Utilities.formatDate(sDate, "GMT+1", "dd.MM.yy");
}

//function to add data in the prev-mail columns
function appending_func(sheet, appenddata,appendRow,k){
 
  var appending = sheet.getRange(appendRow, k, 1, 1);
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
  appendRow = appendRow +1
  return appendRow;
}

//collection of the data from the spreadsheet "Kopie von Lists" that contains mail id informations
function mailspreadsheet_func(){
  var MailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kopie von Lists");
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
  var default_id ="bhagyalaxmi.dinesh@valeo.com"
  var returnvalue =[0,0,0]
  
  if (Request_recheck !== request){
    returnvalue[0] = 'True'
    returnvalue[1] = default_id
    returnvalue[2] = "abc"
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
    case 11:
      kabinetname = 'Kabinet B'
      break;
    case 16:
      kabinetname = 'Kabinet C'
      break;
    case 21:
      kabinetname = 'Kabinet D'
      break;
    case 26:
      kabinetname = 'Kabinet 7'
      break;
    case 31:
      kabinetname = 'Kabinet MPx PP1'
      break;
    case 36:
      kabinetname = 'Kabinet MPx PP2'
      break; 
    case 41:
      kabinetname = 'Kabinet MPx ET'
      break; 
    case 46:
      kabinetname = 'Kabinet MPx Helmholz'
      break;
    case 51:
      kabinetname = 'Kabinet PRE1 LAB'
      break;
    case 56:
      kabinetname = 'Kabinet PRE2 LAB'
      break;
    case 61:
      kabinetname = 'Kabinet PRE3 LAB'
      break;
    case 66:
      kabinetname = 'Kabinet PRE4 G21'
      break;
    case 71:
      kabinetname = 'Kabinet PRE5 G21'
      break;
    default:
      kabinetname = 'Default Kabinet'
      break;
        
  }
  
  return kabinetname
}

//function to accumulate data for the mail
function final_step(sheet,data_mail,rowData,requester,requestee,appendRow,l,m,n){
  
  var default_id ="bhagyalaxmi.dinesh@valeo.com"
  var Requester_id,Requestee_id, Requester_name, Requestee_name,Requester_recheck,Requestee_recheck,Request_id,Request_name,Request_recheck; 
  var kabinet = kabinetname(l);
  var subject = 'Sending email about slot allotment Kabinet:: '+ kabinet;

  Request_id = mailid(data_mail,requester,requestee)//function call to obtain the mail ids
  Requester_id = Request_id[0]
  Requestee_id= Request_id[1]
  
  Request_name =Recipeint_name(data_mail,requester,requestee)//function call to obtain the recipients name
  Requester_name =Request_name[0]
  Requestee_name=Request_name[1]

  Request_recheck= Recipeint_Recheck(data_mail,requester,requestee)//function call to check if an entry for the requester and requestee is present in the spreadhsheet 'Lists'
  Requester_recheck=Request_recheck[0]
  Requestee_recheck=Request_recheck[1]
  
  var values=[0,0,0];//if the entry is not available set the default ids and names
  values= default_set(Requestee_recheck,requestee)
  if(values[0] == 'True'){Requestee_id = values[1]; Requestee_name = values[2]}
  
  values = default_set(Requester_recheck,requester)
  if(values[0] == 'True'){Requester_id = values[1]; Requester_name = values[2]}
  
  // curently send all the data to one mail ID
  var recipientsTO = default_id;
  
  var message =rowData[2]+"<br />"+ rowData[l] +"<br />"+"Requester is: "+requester +"--"+ Requester_name +"<br />" +"Requestee is: "+ rowData[m] +" --"+ Requestee_name+ "<br />"+"Requester mail id: " + Requester_id+"<br />"+ "Requestee mail id: "+ Requestee_id ;
  
  MailApp.sendEmail({to: recipientsTO,subject: subject,htmlBody: message}); //send mail
  var appenddata = [requester +"--" + requestee] 
  appending_func(sheet, appenddata, appendRow, n)
  appendRow = appendRow +1;
  
  return appendRow
  
}

