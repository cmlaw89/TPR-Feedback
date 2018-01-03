function onOpen() {
  //Adds the feedback menu to the Translation Proofreading Sheet (Schedule)
  
  SpreadsheetApp.getUi()
      .createMenu('Wallace')
      .addItem('Submit Feedback', 'openDatePicker')
      .addItem('View Feedback', 'viewFeedback')
      .addToUi();
}


function openDatePicker() {
  var html = HtmlService.createTemplateFromFile('Date_Picker').evaluate().setHeight(300).setWidth(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Please select the feedback date.");
}


function viewFeedback() {
  var html = HtmlService.createTemplateFromFile("View_Feedback_Index");
  html = html.evaluate()
  .setTitle("View Feedback")
  .setHeight(450)
  .setWidth(750)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  
  SpreadsheetApp.getUi().showModalDialog(html, "Submitted Feedback")
}


function tprFeedback(date, month) {
  //Opens the feedback sidebar
  var months = {0: "January", 
                1: "February",
                2: "March",
                3: "April",
                4: "May",
                5: "June",
                6: "July",
                7: "August",
                8: "September",
                9: "October",
                10: "November",
                11: "December"} 
  
  var user =  Session.getActiveUser().getEmail().split("@")[0].substr(0,1).toUpperCase() + Session.getActiveUser().getEmail().split("@")[0].substr(1)
  var all_cases = getCases(user, month);
  if (all_cases == false) {
    SpreadsheetApp.getUi().alert("You have no cases assigned for this month.", SpreadsheetApp.getUi().ButtonSet.OK);
  }  
  else {
    var cases = all_cases[date]
    if (cases == undefined) {
      SpreadsheetApp.getUi().alert("You have no cases assigned for this day.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
    else {
      var year = SpreadsheetApp.getActiveSpreadsheet().getName().split(" ")[2].slice(2, 4);
      var html = HtmlService.createTemplateFromFile('Index');
      html.cases = cases;
      html.all_cases = all_cases;
      html.user = user;
      html.month_year = "-" + pad(month+1) + "-" + year;
      html.month = months[month]
      html = html.evaluate().setTitle("Translation Proofreading Feedback");
      SpreadsheetApp.getUi().showSidebar(html);
    }
  }
}


function submitFeedback(values) {
  
  //Delete existing feedback with the same case ID
  var feedback_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TPR Feedback");
  var case_ID = values[1];
  var case_ID_column = feedback_sheet.getRange(2, 2, feedback_sheet.getLastRow() - 1, 1).getValues();
  var row_index = -1;
  var i = 0;
  while (row_index == -1 && i < case_ID_column.length) {
    if (case_ID_column[i][0] == case_ID) {
      row_index = i;
    }
    i += 1;
  }
  if (row_index != -1) {
    feedback_sheet.deleteRow(row_index + 2);
  }
  
  //Add the new feedback
  feedback_sheet.getRange(feedback_sheet.getLastRow() + 1, 1, 1, values.length).setValues([values]);
}




function getCases(user, month) {
  //Searches the case ID columns of the active month sheet.
  //Returns an array of the case IDs for the specified translation proofreader (user)

  var today = new Date();
  var today_date = today.getDate();
  var today_month = today.getMonth();
   
  var months = {0: "JAN", 
                1: "FEB",
                2: "MAR",
                3: "APR",
                4: "MAY",
                5: "JUN",
                6: "JUL",
                7: "AUG",
                8: "SEP",
                9: "OCT",
                10: "NOV",
                11: "DEC"}  
  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var month_sheet = ss.getSheetByName(months[month]);
  
  var cases = [];
  var num_rows = month_sheet.getMaxRows();
  
  //Find the start row for each week in the schedule sheet
  var week_indexes = []
  var column_A_values = month_sheet.getRange("A:A").getValues();
  var sheet_height = column_A_values.length;
  for (var i = 0; i < sheet_height; i++) {
    if (column_A_values[i][0] == "Date") {
      week_indexes.push(i);
    }
  }
  week_indexes.push(sheet_height)
  
  for (var i = 0; i < week_indexes.length - 1; i++) {
    for (var j = 0; j < 7; j++) {
      cases = cases.concat(month_sheet.getRange(week_indexes[i] + 1, (7*j+2), week_indexes[i+1] - week_indexes[i], 4).getValues());
    }
  }
  
  //Find the index of tomorrows date if the current month is selected
  var date_index = 0;
  if (today_month == month) {
    var found = false;
    while (!found && date_index < cases.length) {
      if (cases[date_index][0] == today_date+1) {
        found = true;
      }
      else {
        date_index += 1;
      }
    }
  }
  else {
    date_index = cases.length;
  }
  
  var assigned_cases = {};
  
  var day_indexes = [];
  for (i = 0; i < week_indexes.length - 1; i++) {
    for (j = 0; j < 7; j++) {
      day_indexes.push(j*(week_indexes[i + 1]-week_indexes[i]) + 7*(week_indexes[i] - week_indexes[0]));
    }
  }
  
  var day = "";
  for (var i = 0; i < date_index; i++) {
    if (day_indexes.indexOf(i) != -1) {
      day = cases[i][0];
    }
    if (cases[i][0].length == 7) {
      var t_proofreader = cases[i][2].trim()
      if (Object.keys(assigned_cases).indexOf(t_proofreader) == -1) {
        assigned_cases[t_proofreader] = {};
        assigned_cases[t_proofreader][day] = [[cases[i][0], cases[i][3]]]
      }
      else {
        if (Object.keys(assigned_cases[t_proofreader]).indexOf(day.toString()) == -1) {
          assigned_cases[t_proofreader][day] = [[cases[i][0], cases[i][3]]]
        }
        else {
          assigned_cases[t_proofreader][day].push([cases[i][0], cases[i][3]]);
        }
      }
    }
  }
  
  if (Object.keys(assigned_cases).indexOf(user) != -1) {
    return assigned_cases[user];
  }
  else {
    return false;
  }
}




function getOutstanding(cases) {
  //Get list of cases with incomplete feedback for the selected month
  
  var TPR_Feedback_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TPR Feedback");
  var complete = TPR_Feedback_sheet.getRange(2, 3, TPR_Feedback_sheet.getLastRow() - 1, 1).getValues();
  var completed_cases = [];
  for (i = 0; i < complete.length; i++) {
    completed_cases.push(complete[i][0].slice(0, 7));
  }
  
  var all_cases = [];
  for (var i = 0; i < Object.keys(cases).length; i++) {
    var day = Object.keys(cases)[i];
    all_cases = all_cases.concat(cases[day]);
  }
      
  
  var incomplete_cases = []
  for (i = 0; i < all_cases.length; i++) {
    if (completed_cases.indexOf(all_cases[i]) == -1) {
      incomplete_cases.push(all_cases[i]);
    }
  }
  
  return incomplete_cases
}




function getFeedback() {
  //Extracts the submitted feedback for the active user
  
  var user = Session.getActiveUser().getEmail().split("@")[0].substr(0,1).toUpperCase() + Session.getActiveUser().getEmail().split("@")[0].substr(1);
  var TPR_Feedback_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TPR Feedback");
  var feedback = []
  for (var i = 2; i < TPR_Feedback_sheet.getLastRow() + 1; i++) {
    var entry = TPR_Feedback_sheet.getRange(i, 2, 1, 17).getValues()[0];
    if (entry[0] == user) {
      entry[9] = entry[9].toString();
      entry[16] = entry[16].toString();
      feedback.push(entry.slice(1));
    }
  }
  return feedback.reverse();
}




function include(filename) {
  //Adds stylesheet and javascript to Index.html
  
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}



function pad(n) {
    return (n < 10) ? ("0" + n) : n;
}