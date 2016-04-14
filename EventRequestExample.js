//Emails of people who may need to approve an event
var foodAppEmail = ;
var avAppEmail = ;
var techAppEmail = ;
var facAppEmail = ;
var transAppEmail = ; 
var cfoAppEmail = ;
var headmasterAssistantEmail = ;
var adminEmail = ; //CHANGE?
var msAssistantEmail = ;
var usAssistantEmail = ;
var calEmail = ;

//Maps room name -> calendar
var rooms = {};
rooms["Fake Room"] = "u1p0jq9baccu8vvbmgn447b330@group.calendar.google.com";
//I deleted these mappings



//Maps room name -> contact
var roomContacts = {};
roomContacts["Fake Room"] = "email address goes here";
//Deleted these mappings


/*
Retrieves unread emails from inbox
Calls readEmail on each of the emails
*/
function getInbox() {
  var numUnread = GmailApp.getInboxUnreadCount();
  Logger.log(numUnread);
  if (numUnread > 0) {
    var unreadThreads = GmailApp.getInboxThreads(0, numUnread);
    for (var i = 0; i < unreadThreads.length; i++) {
      if (unreadThreads[i].getFirstMessageSubject().equals(
        "Form Submission Received for form: Facilities Request Form")) {
        var messages = unreadThreads[i].getMessages();
        for (var j = 0; j < messages.length; j++) {
          if (messages[j].isUnread()) {
           readEmail(messages[j]); 
          }
        }
      }
    }
    GmailApp.markThreadsRead(unreadThreads);
  }
}

/*
Parses the email and updates variables based on information
Calls methods to create the event, create spreadsheet, and update master spreadsheet
*/
//(t)
function readEmail(message) {
  //var message = GmailApp.getInboxThreads(0, 1)[0].getMessages()[0]; //COMMENT OUT
  
  /*Switched to using getBody rather than getRawContent. Seems to fix issue where
  markup in the e-mail will switch the encoding and make it unreadable. 10/02/15
  
  Update: 10/05/15. Nope. Switched back to getRawContent(). Weird things happen using getBody() -
  the content is garbled by <br> html code. 
  
  TO DO: Make the script use getBody ONLY if it's in base 64 encoding.
  
  */
  var body = message.getRawContent(); 
  
  //This handles when people put crazy markup in the text
  if(body.indexOf("Content-Transfer-Encoding: base64") >= 0)
  {
   body = message.getBody(); 
  }
 // This should be used for base 64 only
   
  var elements = body.split("\n");
  for (var a = 0; a < elements.length; a++) {
    elements[a] = elements[a].trim();
    //Logger.log(elements[a]);
  }
  
  var eventName;
  var date;
  var startTime;
  var endTime;
  var room;
  var contact;
  var contactEmail;
  var roomEmail;
  var roomContact;
  var chap1 = "";
  var chap2 = "";
  var chap3 = "";
  
  var foodService = false;
  var avService = false;
  var techService = false;
  var facService = false;
  var transService = false;
  
  var startDate;
  var endDate;
  var otherName = "";
  
  for (var j = 0; j < elements.length; j++){  
    if (elements[j].equals("<strong>Event Name</strong>")) {
      eventName = elements[j + 5];
    } else if (elements[j].equals("<strong>Date of your event</strong>")) {
      startDate = elements[j + 5];
      endDate = elements[j + 5];
    } else if (elements[j].equals("<strong>Event Start Time</strong>")) {
      //Check to see if user entered time without minutes, e.g. 12 instead of 12:00.
      if(elements[j + 5].indexOf(":") < 0)
      {
        //Logger.log(elements[j + 5].indexOf(":"));
        elements[j + 5] = elements[j + 5] + ":00";
      }
      
      if(elements[j + 33].indexOf(":") < 0)
      {
        elements[j + 33] = elements[j + 33] + ":00";
      }
      
      if (!elements[j + 5].charAt(2).equals(":")) {
        //Logger.log(elements[j+5]);
        elements[j + 5] = "0" + elements[j + 5];
      }
      if (!elements[j + 33].charAt(2).equals(":")) {
        //Logger.log(elements[j+33]);
        elements[j + 33] = "0" + elements[j + 33];
      }
      startTime = elements[j + 5] + elements[j + 19]
      endTime = elements[j + 33] + elements[j + 47];
    } else if (elements[j].equals("<strong>Event Location</strong>")) {
      room = elements[j + 5];
      roomEmail = rooms[room];
      roomContact = roomContacts[room];
    } else if (elements[j].equals("<strong>Please specify where your event will be held</strong>")) {
      otherName = elements[j + 5];
    } else if (elements[j].equals("<strong>Your Name</strong>")) {
      contact = elements[j + 5];
    } else if (elements[j].equals("<strong>Email</strong>")) {
      contactEmail = elements[j + 5];
    } else if (elements[j].equals("<strong>Chaperone 1 Email</strong>")) {
      chap1 = elements[j + 5];
    } else if (elements[j].equals("<strong>Chaperone 2 Email</strong>")) {
      chap2 = elements[j + 5];
    } else if (elements[j].equals("<strong>Chaperone 3 Email</strong>")) {
      chap3 = elements[j + 5];
    }
    else if (elements[j].equals("<strong>Service Required</strong>")) {
      var services = elements[j + 5];
      foodService = (services.indexOf("Catering") >= 0);
      avService = (services.indexOf("Audio/Visual") >= 0);
      techService = (services.indexOf("Technology") >= 0);
      facService = (services.indexOf("Facilities Setup") >= 0);
      transService = (services.indexOf("Transportation") >= 0);
    } else if (elements[j].equals("<strong>Do you need to begin setting up prior to the time of the event?</strong>")
               && elements[j + 5].equals("Yes")) {
       
      //Check to see if user entered time without minutes, e.g. 12 instead of 12:00.
      if(elements[j + 33].indexOf(":") < 0)
      {
        elements[j + 33] = elements[j + 33] + ":00";
      }
      if (!elements[j + 33].charAt(2).equals(":")) {
        elements[j + 33] = "0" + elements[j + 33];
      }
      startTime = elements[j + 33].substring(0,5) + elements[j + 47];
      startDate = elements[j + 19];   
    } else if (elements[j].equals("<strong>Will you need time to clear the event space after the end of the event?</strong>")
               && elements[j + 5].equals("Yes")) {
       
      //Check to see if user entered time without minutes, e.g. 12 instead of 12:00.
      if(elements[j + 33].indexOf(":") < 0)
      {
        elements[j + 33] = elements[j + 33] + ":00";
      }
      
      if (!elements[j + 33].charAt(2).equals(":")) {
        elements[j + 33] = "0" + elements[j + 33];
      }
      
      //Use the first 5 characters of the time to erase any stray entries (so 09:00am becomes 09:00 since AM will already be appended.
      endTime = elements[j + 33].substring(0,5) + elements[j + 47];
      endDate = elements[j + 19];
    }
  }
  
  Logger.log(eventName);
  Logger.log(startDate);
  
  if (!startDate.equals(endDate)) {
   date = startDate + " - " + endDate; 
  } else {
   date = startDate; 
  }
  Logger.log(endDate);
 
 
  startDate = validateDate(startDate);
  endDate = validateDate(endDate);
  
  Logger.log(startTime);
  Logger.log(endTime);
  
  //create event
  var syear = startDate.substring(6, 10);
  var smonth = startDate.substring(0, 2) - 1;
  var sday = startDate.substring(3, 5);
  var sHour = startTime.substring(0, 2) - 0 + 12;
  if (startTime.charAt(5).equals('A')) {
    sHour -= 12;
  } else if (sHour == 24) {
    sHour = 12;
  }
  
  
  var eyear = endDate.substring(6, 10);
  var emonth = endDate.substring(0, 2) - 1;
  var eday = endDate.substring(3, 5);
  var sMin = startTime.substring(3, 5); 
  var eHour = endTime.substring(0, 2) - 0 + 12;
  if (endTime.charAt(5).equals('A')) {
    eHour -= 12;
  } else if (eHour == 24) {
    eHour = 12;
  }
  
  var eMin = endTime.substring(3, 5); 
  
  var startCalDate = new Date(syear, smonth, sday, sHour, sMin);
  var endCalDate = new Date(eyear, emonth, eday, eHour, eMin);
  
  
  Logger.log("Start: " + startCalDate);
  Logger.log("Year:" + eyear + "Month: " + emonth + "Day: " + eday + "Hour: " + eHour + "Min: " + eMin + "End: " + endCalDate);

  
  var event = CalendarApp.getDefaultCalendar().createEvent(
    eventName, startCalDate, endCalDate);
 
  event.setLocation(room);
  event.addGuest(contactEmail);
  event.addGuest(roomEmail);
  
   //create spreadsheet
  var spreadsheet = SpreadsheetApp.create(eventName + " " + date);
  spreadsheet.addEditors([foodAppEmail, avAppEmail, techAppEmail, facAppEmail, transAppEmail, cfoAppEmail,
                          contactEmail, headmasterAssistantEmail, adminEmail]);
  spreadsheet.addViewers([msAssistantEmail, usAssistantEmail, roomContact, calEmail]);
  
  
  //Add chaperones as viewers if appropriate
  if(chap1.length > 1)
  {
    spreadsheet.addViewers([chap1]);
  }
  
  if(chap2.length > 1)
  {
    spreadsheet.addViewers([chap2]);
  }
  
  if(chap3.length > 1)
  {
    spreadsheet.addViewers([chap3]);
  }
  
  
  
  
  var url = spreadsheet.getUrl();
  //Add URL address of speeadsheet to the event
  event.setDescription(url);
  
  if (foodService) {
    notifyCreatedToApprove(foodAppEmail, eventName, url);
  }
  if (avService) {
   notifyCreatedToApprove(avAppEmail, eventName, url); 
  }
  if (techService) {
   notifyCreatedToApprove(techAppEmail, eventName, url); 
  }
  if (facService) {
    notifyCreatedToApprove(facAppEmail, eventName, url);
  }
  if (transService) {
    notifyCreatedToApprove(transAppEmail, eventName, url);
  }
  //notifyCreatedToApprove(cfoAppEmail, eventName, url);
  notifyCreated(headmasterAssistantEmail, eventName, url);
  notifyCreated(contactEmail, eventName, url);
  
  if(chap1.length > 1)
  { 
    notifyCreated(chap1, eventName, url);
  }
  if(chap2.length > 1)
  {
    notifyCreated(chap2, eventName, url);
  }
  if(chap3.length > 1)
  {
    notifyCreated(chap3, eventName, url);
  }
  
  
  //TO DO: ADD E-MAIL NOTIFICATION FOR CHAPERONES
  
  var roomApproved = "";
  if (checkRoomAvailable(room, url, startCalDate, endCalDate)) {
    roomApproved = "Yes";
  }
  
  //CHANGED - 9/25/15: used startDate instead of date so master sheet sorting by date will work properly
  addToMaster(eventName, startDate, (startTime + " - " + endTime), room, contact, contactEmail, url, 
              foodService, avService, techService, facService, transService, roomApproved, otherName);
  
  //update individual spreadsheet
  var newSheet = SpreadsheetApp.openByUrl(url);
  var approvalsRow = ["Approved?", "NO", "APPROVALS:", "Food:"]
  if (foodService) {
    approvalsRow.push("Not Approved");
  } else {
    approvalsRow.push("X");
  }
  if (avService) {
    approvalsRow.push("AV:", "Not Approved");
  } else {
    approvalsRow.push("AV: ", "X");
  }                   
  if (techService) {
    approvalsRow.push("Technology:", "Not Approved");
  } else {
    approvalsRow.push("Technology: ", "X");
  }
  if (facService) {
    approvalsRow.push("Facilities:", "Not Approved");
  } else {
    approvalsRow.push("Facilities: ", "X");
  }
  if (transService) {
    approvalsRow.push("Transportation:", "Not Approved");
  } else {
    approvalsRow.push("Transportation: ", "X");
  }
  if (roomApproved.equals("")) {
    approvalsRow.push("Room:", "Not Approved");
  } else {
    approvalsRow.push("Room: ", roomApproved);
  }
  approvalsRow.push("CFO:", "Not Approved");
  newSheet.appendRow(["The following row lists who has approved your event thus far. It is your responsibility " +
                      "to makes sure your event gets approved. Editing the following row will NOT affect " +
                      "the approval status of your event."]);
  newSheet.appendRow(approvalsRow);
  var submitterRow = ["SUBMITTER INFO:"];
  var generalInfoRow = ["GENERAL INFO:"];
  var roomInfoRow = ["ROOM INFO:"]
  var foodRow = ["FOOD DETAILS:"];
  var avRow = ["AV DETAILS:"];
  var techRow = ["TECHNOLOGY DETAILS:"];
  var facRow = ["FACILITIES DETAILS:"];
  var transRow = ["TRANSPORTATION DETAILS:"];
  
  for (var a = 0; a < elements.length; a++) { 
    if (elements[a].equals("<strong>Your Name</strong>")) {
      submitterRow.push(elements[a + 5], elements[a + 19], elements[a + 33]); 
    } else if (elements[a].equals("<strong>Event Name</strong>")) {
      generalInfoRow.push(elements[a + 5], elements[a + 19]); 
    } else if (elements[a].equals("<strong>Event Start Time</strong>")) {
      generalInfoRow.push(elements[a + 5] + " " + elements[a + 19] + " - " +
                         elements[a + 33] + " " + elements[a + 47]); 
    } else if (elements[a].equals("<strong>Event Location</strong>")) {
      roomInfoRow.push(elements[a + 5]); 
    } else if (elements[a].equals("<strong>Do you need to begin setting up prior to the time of the event?</strong>")
              && elements[a + 5].equals("Yes")) {
      roomInfoRow.push("Room Start Date: " + elements[a + 19]);
      roomInfoRow.push("Room Start Time: " + elements[a + 33] + " " + elements[a + 47]);
    } else if (elements[a].equals("<strong>Will you need time to clear the event space after the end of the event?</strong>")
              && elements[a + 5].equals("Yes")) {
      roomInfoRow.push("Room End Date: " + elements[a + 19]);
      roomInfoRow.push("Room End Time: " + elements[a + 33] + " " + elements[a + 47]);
    } else if (elements[a].equals("<strong># Attending</strong>") && (!elements[a + 5].equals(""))) {
      generalInfoRow.push(elements[a + 5] + " attending"); 
    } else if (elements[a].equals("<strong>Description</strong>")) {
      generalInfoRow.push("Description: " + elements[a + 5]);
    } else if (elements[a].equals("<strong>Chaperone 1 Email</strong>") && (!elements[a + 5].equals(""))) {
      generalInfoRow.push("Chaperone 1: " + elements[a + 5]);
    } else if (elements[a].equals("<strong>Chaperone 2 Email</strong>") && (!elements[a + 5].equals(""))) {
      generalInfoRow.push("Chaperone 2: " + elements[a + 5]);
    } else if (elements[a].equals("<strong>Chaperone 3 Email</strong>") && (!elements[a + 5].equals(""))) {
      generalInfoRow.push("Chaperone 3: " + elements[a + 5]);
      //Old way
      //generalInfoRow.push("Chaperone 1:", elements[a + 5], elements[a + 19], elements[a + 33], url);
      //notifyChap(elements[a + 5], elements[a + 19], elements[a + 33]);
    } else if (elements[a].equals("<strong>Set-Up Time</strong>")) {
      facRow.push("Set-Up Time:", elements[a + 5], "Take-Down Time:", elements[a + 19]); 
    } else if (elements[a].equals("<strong>Number of Tables</strong>")) {
      facRow.push(elements[a + 5] + " tables"); 
    } else if (elements[a].equals("<strong>Number of Chairs</strong>")) {
      facRow.push(elements[a + 5] + " chairs"); 
    } else if (elements[a].equals("<strong>Number of Coat Racks</strong>")) {
      facRow.push(elements[a + 5] + " coat racks"); 
    } else if (elements[a].equals("<strong>Number of Trash Cans</strong>")) {
      facRow.push(elements[a + 5] + " trash cans"); 
    } else if (elements[a].equals("<strong>Hours Heat & Ventilation is Needed</strong>")) {
      facRow.push(elements[a + 5] + " hours heat/ventilation"); 
    } else if (elements[a].equals("<strong>Please specify other items needed</strong>")) {
      facRow.push("Other items: " + elements[a + 5]); 
    } else if (elements[a].equals("<strong>Facilities Request Details</strong>")) {
      facRow.push("Additional Details: " +elements[a + 5]); 
    } else if (elements[a].equals("<strong>Budget Account Number</strong>")) {
      foodRow.push("Budget Account #:", elements[a + 5], "Cost Estimate:", "(Sage fill-in)",
                   "Estimate Verified by Requester:", "");
    } else if (elements[a].equals("<strong>Service Time</strong>") && (!elements[a + 5].equals(""))) {
      foodRow.push("Service Time:", elements[a + 5]); 
    } else if (elements[a].equals("<strong>Number in Party</strong>") && (!elements[a + 5].equals(""))) {
      foodRow.push(elements[a + 5] + " people"); 
    } else if (elements[a].equals("<strong>Type of Service:</strong>") && (!elements[a + 5].equals(""))) {
      foodRow.push(elements[a + 5] + "Service"); 
    } else if (elements[a].equals("<strong>Food Service Request Details</strong>")) {
      foodRow.push("Additional Details: " + elements[a + 5]); 
    } else if (elements[a].equals("<strong>Please check all the items you will need</strong>")) {
      avRow.push(elements[a + 5]); 
    } else if (elements[a].equals("<strong>AV Request Details</strong>")) {
      avRow.push("Additional Details: " + elements[a + 5]); 
    }  else if (elements[a].equals("<strong>Please specify any non-AV tech needs.</strong>")) {
      techRow.push(elements[a + 5]); 
    } else if (elements[a].equals("Van"))
    {
      transRow.push("Type of Request: Van");
    }
    else if (elements[a].equals("Bus"))
    {
      transRow.push("Type of Request: Bus");
    }
    
    
    
    
    /*else if (elements[a].equals("<strong>Number of Vans</strong>") && (!elements[a + 5].equals(""))) {
      transRow.push(elements[a + 5] + " Van(s)"); 
    } else if (elements[a].equals("<strong>Number of Buses</strong>") && (!elements[a + 5].equals(""))) {
      transRow.push(elements[a + 5] + " Bus(es)"); 
    } else if (elements[a].equals("<strong>Number of Students</strong>") && (!elements[a + 5].equals(""))) {
      transRow.push(elements[a + 5] + " Student(s)"); 
    } else if (elements[a].equals("<strong>Number of Adults</strong>") && (!elements[a + 5].equals(""))) {
      transRow.push(elements[a + 5] + " Adult(s)"); 
    } else if (elements[a].equals("<strong>Transportation Request Details</strong>")) {
      transRow.push("Additional Details: " + elements[a + 5]); 
    }*/
  }
  
  
  newSheet.appendRow(submitterRow);
  newSheet.appendRow(generalInfoRow);
  newSheet.appendRow(roomInfoRow)
  if (foodRow.length > 1) {
    newSheet.appendRow(foodRow);
  }
  if (avRow.length > 1) {
    newSheet.appendRow(avRow);
  }
  if (techRow.length > 1) {
    newSheet.appendRow(techRow);
  }
  if (facRow.length > 1) {
    newSheet.appendRow(facRow);
  }
  if (transRow.length > 1) {
    newSheet.appendRow(transRow);
  }
  var range = newSheet.getRange("A2:Q10");
  range.setWrap(true);
  newSheet.getRange("A2:Q2").setBackground("PowderBlue");
  newSheet.getRange("A4:Q4").setBackground("PowderBlue");
  newSheet.getRange("A6:Q6").setBackground("PowderBlue");
  newSheet.getRange("A8:Q8").setBackground("PowderBlue");
  newSheet.getRange("A10:Q10").setBackground("PowderBlue");
  if (newSheet.getRange("A6").getValue().equals("FOOD DETAILS:")) {
    newSheet.getRange("E6").setBackground("yellow");
    newSheet.getRange("G6").setBackground("yellow");
  }
  newSheet.insertImage("http://www.pingry.org/uploaded/images2/logo.png", 1, newSheet.getSheets()[0].getLastRow() + 2);
}

/*
Checks if a room is free on the calendar for the time requested
*/
function checkRoomAvailable(room, url, startTime, endTime) {
  var calID = rooms[room];
  var events = CalendarApp.getCalendarById(calID).getEvents(startTime, endTime);
  if (events.length > 1) {
    var message = "This is an automated message informing you than an event " +
      "has been requested in the " + room + ". All the details of the event " +
      "can be viewed in the following spreadsheet: " + url +". It appears that " +
      "this event conflicts with something previously scheduled. Please provide " +
      "feedback to the headmaster's assistant. Thank you." +
      "\n\nPlease do not reply to this message."
    var subject = "Event Requested in the " + room;
    GmailApp.sendEmail(roomContacts[room], subject, message);
    return false;
  } else {
    var message = "This is an automated message informing you than an event " +
      "has been scheduled in the " + room + ". All the details of the event " +
      "can be viewed in the following spreadsheet: " + url +". If this is an " +
      "issue, please notify the headmaster's assistant. Thank you." +
      "\n\nPlease do not reply to this message."
    var subject = "Event Scheduled in the " + room;
    GmailApp.sendEmail(roomContacts[room], subject, message);
    return true;
  }
}

/*
Adds a row to the master calendar with the information from the event
*/
function addToMaster(eventName, date, time, room, contact, cEmail, url, food, av, tec, fac, trans, roomApp, other) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  var foodSig; var avSig; var tecSig; var facSig; var transSig;
  if (food) {
    foodSig = "";
  } else {
    foodSig = "X";
  }
  if (av) {
    avSig = "";
  } else {
    avSig = "X";
  }
  if (tec) {
    tecSig = "";
  } else {
    tecSig = "X";
  }
  if (fac) {
    facSig = "";
  } else {
    facSig = "X";
  }
  if (trans) {
    transSig = "";
  } else {
    transSig = "X";
  }
  if (room.equals("Other")) {
   room = other; 
  }
  sheet.appendRow([eventName, date, time, room, contact, cEmail, 
                   url, foodSig, avSig, tecSig, facSig, transSig, "", roomApp, "NO"]);
  var cells = sheet.getRange(sheet.getLastRow(), 1, 1, 15);
  cells.setBackground("yellow");

}

/*
Goes through each unapproved event and checks if it has been approved
Updates spreadsheet/event accordingly
*/
function checkIfApproved() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var approvals = sheet.getSheetValues(i, 8, 1, 8)[0];
    if (approvals[7].equals("NO")) {
      var nowApproved = true;
      
      var food = approvals[0];
      if (food.equals("")) {
       nowApproved = false; 
      } else if (!food.equals("X")) {
        if (food.length < 13) {
          nowApproved = false;
          sheet.getRange(i, 8).setValue(food + " - cost not verified");
        } else if (food.substring(food.length - 12, food.length - 9).equals("not")) {
          //check if it has been approved by owner
          var currUrl = sheet.getSheetValues(i, 7, 1, 1)[0][0];
          var currSheet = SpreadsheetApp.openByUrl(currUrl).getSheets()[0];
          if (currSheet.getRange("G6").getValue().equals("")) {
            nowApproved = false;
          } else {
           var cost = currSheet.getRange("E6").getValue(); 
           sheet.getRange(i, 8).setValue(food.substring(0, food.length - 12) + "verified " + cost);
          }
         }
      }
      
      for (var j = 1; j < 7; j++) {
        if (approvals[j].equals("")) {
          nowApproved = false;
        }
      }
      if (nowApproved) {
        //TO DO: Edit the individual spreadsheet to say approved and make read-only.
        
        var a1 = "O" + i.toString();
        sheet.getRange(a1).setValue("YES");
        var a1Range = "A" + i.toString() + ":" + a1;
        sheet.getRange(a1Range).setBackground("white");
        
        var eventName = sheet.getSheetValues(i, 1, 1, 1)
        var spreadsheetUrl = sheet.getSheetValues(i, 7, 1, 1)[0][0];
        
        notifyOwner(sheet.getSheetValues(i, 1, 1, 7)[0]);
        
        //TO DO Also notify chaperones that event is approved
        /* var currUrl = sheet.getSheetValues(i, 7, 1, 1)[0][0];
          var currSheet = SpreadsheetApp.openByUrl(currUrl).getSheets()[0];*/
      }
    }
    updateApprovalSpreadsheet(sheet.getSheetValues(i, 7, 1, 1)[0][0], approvals);
  }

}

/*
Notifies somebody that an event has been created which they need to approve
*/
function notifyCreatedToApprove(email, eventName, ssUrl) {
  var thisUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var message = "This is an automated email informing you that an event, " +
    eventName +  ", has been submitted and needs your approval.\n" +
    "You can view/edit the details of the event here: " + ssUrl + ".\n" +
    "To approve the event, write your initials in the appropriate box in the " +
    "following spreadsheet: " + thisUrl + ".\n\n Please do not reply to this email.";
  var subject = "Event to Approve: " + eventName;
  GmailApp.sendEmail(email, subject, message);
}

/*
Notifies somebody that an event has been created
*/
function notifyCreated(email, eventName, ssUrl) {
  var message = "This is an automated email informing you than an event, " +
    eventName + ", has been submitted.\nYou can view/edit the details of " +
    "the event here: " + ssUrl + ".\n\n Please do not reply to this email.";
  var subject = "Event Created: " + eventName;
  GmailApp.sendEmail(email, subject, message);
}

/*
Notifies event owner that their event has been approved
*/
function notifyOwner(eventInfo) {
  var event = eventInfo[0];
  var date = eventInfo[1];
  var month = date.getMonth() + 1; //Added a +1 here - for some reason was always a month behind.
  var day = date.getDate();
  var year = date.getYear();
  var time = eventInfo[2];
  var room = eventInfo[3];
  var owner = eventInfo[4];
  var email = eventInfo[5];
  var url = eventInfo[6];
  
  var message = "Dear " + owner + ",\n" +
    "This is an automated email informing you that your event, " +
    event + ", has been approved for " + month + "/" + day + "/" + year +
    " at " + time + ", " + room +
    ". You can view all the finalized details for your event at " + url +
    ".\nThank you. \n\n Please do not reply to this email.";   
  var subject = "Event Approved: " + event;
  GmailApp.sendEmail(email, subject, message);
  
}

/*
Sees which events still need to be approved and sends out reminders accordingly
*/
function reminders() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  var food = [];
  var av = [];
  var tech = [];
  var fac = [];
  var trans = [];
  var cfo = [];
  var roomNames = [];
  var roomUrls = [];
  for (var i = 2; i <= sheet.getLastRow(); i++) {
    var name = sheet.getSheetValues(i, 1, 1, 1)[0][0];
    var url = sheet.getSheetValues(i, 7, 1, 1)[0][0];
    var approvals = sheet.getSheetValues(i, 8, 1, 9)[0];
    if (approvals[7].equals("NO")) {
      sendUnapprovedEmail(sheet.getSheetValues(i, 6, 1, 1)[0][0], sheet.getSheetValues(i, 1, 1, 1)[0][0], url);
      if (approvals[0].equals("")) {
        food.push(name + ": " + url);
      }
      if (approvals[1].equals("")) {
        av.push(name + ": " + url);
      }
      if (approvals[2].equals("")) {
        tech.push(name + ": " + url);
      }
      if (approvals[3].equals("")) {
        fac.push(name + ": " + url);
      }
      if (approvals[5].equals("")) {
        cfo.push(name + ": " + url);
      }
      if (approvals[4].equals("")) {
       trans.push(name + ": " + url); 
      }
      if (approvals[6].equals("")) {
        roomUrls.push(name + ": " + url);
        roomNames.push(sheet.getSheetValues(i, 4, 1, 1)[0][0]);
      }
    }
  }
  
  if (food.length != 0) {
    sendReminderEmail(food, foodAppEmail);
  }
  if (av.length != 0) {
    sendReminderEmail(av, avAppEmail);
  }
  if (tech.length != 0) {
    sendReminderEmail(tech, techAppEmail);
  }
  if (fac.length != 0) {
    sendReminderEmail(fac, facAppEmail);
  }
  if (cfo.length != 0) {
    sendReminderEmail(cfo, cfoAppEmail);
  }
  if (trans.length != 0) {
    sendReminderEmail(trans, transAppEmail);
  }
  if (roomNames.length != 0) {
   sendRoomReminderEmail(roomUrls, roomNames); 
  }
}

/*
Sends a reminder email of events that have not been approved
*/
function sendReminderEmail(urls, email) {
 var subject = "Events Pending Approval";
 var thisUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
 var message = "This is an automated email reminder that you still have not " +
   "approved the following event(s):\n\n";
  for (var i = 0; i < urls.length; i++) {
   message = message + urls[i] + "\n\n";
  }
  message += "To approve the event, write your initials in the appropriate box in the " +
    "following spreadsheet: " + thisUrl + ".\n\n Please do not reply to this email."; 
  GmailApp.sendEmail(email, subject, message);
}

/*
Sends a reminder email for rooms that have not been approved
*/
function sendRoomReminderEmail(urls, roomNames) {
  var subject = "Room Request Pending Approval";
  var thisUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  for (var i = 0; i < urls.length; i++) {
    if (roomNames[i] != null) {
      var message = "This is an automated email reminder that you still have not " +
     "approved a room request for an event in " + roomNames[i] +
     ".\nThe information for this event can be found in the following spreadsheet: " +
     urls[i] + ". To approve the event, contact the headmaster's assistant." +
     ".\n\n Please do not reply to this email.";
    var email = roomContacts[roomNames[i]];
    GmailApp.sendEmail(email, subject, message);
    }
   
  }
}

/*
Notifies owner that their event has not yet been approved
*/
function sendUnapprovedEmail(email, event, url) {
  var message = "This is an automated message informing you that your event, " +
    event + ", has not been approved. You can view the details of the event at the " +
    "following url: " + url + ". It is your responsibility to make sure the event is " +
    "approved. If you have any questions, contact the headmaster's assistant." +
    "\n\nPlease do not respond to this email.";
  GmailApp.sendEmail(email, ("Unapproved Event: " + event), message);
}

/*
Informs the chaparones of the event
*/
function notifyChap(c1, c2, c3, url) {
  var message = "This is an automated message informing you that an event request has been " +
    "submitted and you were listed as a chaperone. You can view all the details of the event " +
    "at the following url: " + url + "./n/nPlease do not reply to this message.";
  if (!c1.equals("")) {
   GmailApp.sendEmail(c1, "Chaperoning an Event", message);
  }
  if (!c2.equals("")) {
   GmailApp.sendEmail(c2, "Chaperoning an Event", message);
  }
  if (!c3.equals("")) {
   GmailApp.sendEmail(c3, "Chaperoning an Event", message);
  }
}

/*
Updates the individual spreadsheet based on who has approved the event
*/
function updateApprovalSpreadsheet(spreadsheetUrl, approvals) {
  var sheet = SpreadsheetApp.openByUrl(spreadsheetUrl).getSheets()[0];
  
  
  
  //Check Food
  if (!approvals[0].equals("") && !approvals[0].equals("X")) {
    sheet.getRange("E2").setValue(approvals[0]);
    sheet.getRange("E2").setBackground("PaleGreen");
    sheet.getRange("D2").setBackground("PaleGreen");
  }
  //Check to see if it's awaiting cost approval
  else if(approvals[0].indexOf("not") >= 0)
  {
    sheet.getRange("E2").setValue(approvals[0]);
    sheet.getRange("E2").setBackground("Yellow");
    sheet.getRange("D2").setBackground("Yellow");
  }
  else if(approvals[0].equals("X"))
  {
   sheet.getRange("E2").setValue("N/A");
    sheet.getRange("E2").setBackground("PowderBlue");
    sheet.getRange("D2").setBackground("PowderBlue");
    
  }
  else {
   sheet.getRange("E2").setValue("Not Approved"); 
    sheet.getRange("E2").setBackground("Red");
    sheet.getRange("D2").setBackground("Red");
  }
  
  
   //Check AV
  if (!approvals[1].equals("") && !approvals[1].equals("X")) {
    sheet.getRange("G2").setValue(approvals[1]);
    sheet.getRange("G2").setBackground("PaleGreen");
    sheet.getRange("F2").setBackground("PaleGreen");
  }
  else if(approvals[1].equals("X"))
  {
   sheet.getRange("G2").setValue("N/A");
    sheet.getRange("G2").setBackground("PowderBlue");
    sheet.getRange("F2").setBackground("PowderBlue");
    
  }
  else {
   sheet.getRange("G2").setValue("Not Approved"); 
    sheet.getRange("G2").setBackground("Red");
    sheet.getRange("F2").setBackground("Red");
  }
  
   //Check Tech
  if (!approvals[2].equals("") && !approvals[2].equals("X")) {
    sheet.getRange("I2").setValue(approvals[2]);
    sheet.getRange("I2").setBackground("PaleGreen");
    sheet.getRange("H2").setBackground("PaleGreen");
  }
  else if(approvals[2].equals("X"))
  {
   sheet.getRange("I2").setValue("N/A");
    sheet.getRange("I2").setBackground("PowderBlue");
    sheet.getRange("H2").setBackground("PowderBlue");
    
  }
  else {
   sheet.getRange("I2").setValue("Not Approved"); 
    sheet.getRange("I2").setBackground("Red");
    sheet.getRange("H2").setBackground("Red");
  }
  
   //Check Facilities
  if (!approvals[3].equals("") && !approvals[3].equals("X")) {
    sheet.getRange("K2").setValue(approvals[3]);
    sheet.getRange("K2").setBackground("PaleGreen");
    sheet.getRange("J2").setBackground("PaleGreen");
  }
  else if(approvals[3].equals("X"))
  {
   sheet.getRange("K2").setValue("N/A");
    sheet.getRange("K2").setBackground("PowderBlue");
    sheet.getRange("J2").setBackground("PowderBlue");
    
  }
  else {
   sheet.getRange("K2").setValue("Not Approved"); 
    sheet.getRange("K2").setBackground("Red");
    sheet.getRange("J2").setBackground("Red");
  }
  
   //Check Transportation
  if (!approvals[4].equals("") && !approvals[4].equals("X")) {
    sheet.getRange("M2").setValue(approvals[4]);
    sheet.getRange("M2").setBackground("PaleGreen");
    sheet.getRange("L2").setBackground("PaleGreen");
  }
  else if(approvals[4].equals("X"))
  {
   sheet.getRange("M2").setValue("N/A");
    sheet.getRange("M2").setBackground("PowderBlue");
    sheet.getRange("L2").setBackground("PowderBlue");
    
  }
  else {
   sheet.getRange("M2").setValue("Not Approved"); 
    sheet.getRange("M2").setBackground("Red");
    sheet.getRange("L2").setBackground("Red");
  }
  
  //NOTE: CFO and ROOM are in different orders on individual
  //Spreadsheet vs master spreadsheet. That is why we are 
  //updating approvals[6] before 5
  
   //Check Room
  if (approvals[6].toUpperCase().equals("YES")) {
    sheet.getRange("O2").setValue("Approved");
    sheet.getRange("O2").setBackground("PaleGreen");
    sheet.getRange("N2").setBackground("PaleGreen");
  }
  else {
   sheet.getRange("O2").setValue("Not Approved " + approvals[6]); 
    sheet.getRange("O2").setBackground("Red");
    sheet.getRange("N2").setBackground("Red");
  }
  
   //Check CFO
  /*No need to check for X value, since CFO approval is
  never Not Applicable */
  if (!approvals[5].equals("")) {
    sheet.getRange("Q2").setValue(approvals[5]);
    sheet.getRange("Q2").setBackground("PaleGreen");
    sheet.getRange("P2").setBackground("PaleGreen");
  }
  else {
   sheet.getRange("Q2").setValue("Not Approved"); 
    sheet.getRange("Q2").setBackground("Red");
    sheet.getRange("P2").setBackground("Red");
  }
  
  if (!approvals[7].equals("")) {
    sheet.getRange("B2").setValue(approvals[7]);
  }
  
  
  
  
  //
  /**if (!approvals[1].equals("")) {
    sheet.getRange("G2").setValue(approvals[1]);
  }
  else {
   sheet.getRange("G2").setValue("Not Approved"); 
  }
  
  if (!approvals[2].equals("")) {
    sheet.getRange("I2").setValue(approvals[2]);
  }
  else {
   sheet.getRange("I2").setValue("Not Approved"); 
  }
  
  if (!approvals[3].equals("")) {
   sheet.getRange("K2").setValue(approvals[3]);
  }
  else {
   sheet.getRange("K2").setValue("Not Approved"); 
  }
  
  if (!approvals[4].equals("")) {
   sheet.getRange("M2").setValue(approvals[4]);
  }else {
   sheet.getRange("M2").setValue("Not Approved"); 
  }
  
  if (!approvals[5].equals("")) {
   sheet.getRange("Q2").setValue(approvals[5]);
  }
  if (!approvals[6].equals("")) {
   sheet.getRange("O2").setValue(approvals[6]);
  }
  
  }*/
 
 
}


//Script to move past, approved events into a separate past events sheet
function moveEvents() {
  
  //Column that currently contains dates. 
  var dateColumn = 2;
  
  //Column containing final approval status.
  var approvalColumn = 15;
  
  var sheet = SpreadsheetApp.getActive().getSheetByName("Sheet1");
  var pastEventSheet = SpreadsheetApp.getActive().getSheetByName("Past Events");
  
  var d = new Date();
  var currentDate = new Date(d.getMonth()+1 + "/" + d.getDate() + "/" + d.getYear());
  
  for (var row = 2; row <= sheet.getLastRow(); row++) 
  {
    
    var date = sheet.getRange(row, dateColumn).getValue();
    
    if(currentDate > date && sheet.getRange(row, approvalColumn).getValue() == "YES")
    {
      var eventInfo = sheet.getRange(row, 1, 1, sheet.getLastColumn());
      pastEventSheet.appendRow(eventInfo.getValues()[0]);
      //Logger.log(eventInfo.getValues());
      sheet.deleteRow(row);
      row = row - 1;  
    }
  }
  
  //Sort sheets by date
  sheet.sort(dateColumn);
  pastEventSheet.sort(dateColumn);
  
}

//This function will make sure a date is formatted properly. Example: 4/1/16 becomes 04/01/2016
function validateDate(dateToVerify)
{
  if(dateToVerify.charAt(2) != "/")
 {
   dateToVerify = "0" + dateToVerify;

 }
  if(dateToVerify.charAt(5) != "/")
  {
    dateToVerify = dateToVerify.substring(0,3) + "0" + dateToVerify.substring(3);
  }
  if(dateToVerify.substring(6).length != 4)
  {
   dateToVerify = dateToVerify.substring(0,6) + "20" + dateToVerify.substring(6); 
    
  }
  return dateToVerify;
}
