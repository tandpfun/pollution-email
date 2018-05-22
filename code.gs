// add menu to Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Send Emails")
  .addItem("Send Email Batch","createEmail")
  .addToUi();
}

/**
 * take the range of data in sheet
 * use it to build an HTML email body
 */
function createEmail() {
  var thisWorkbook = SpreadsheetApp.getActiveSpreadsheet();
  var thisSheet = thisWorkbook.getSheetByName('Form Responses 1');

  // get the data range of the sheet
  var allRange = thisSheet.getDataRange();
  
  // get all the data in this range
  var allData = allRange.getValues();
  
  // get the header row
  var headers = allData.shift();
  
  // create header index map
  var headerIndexes = indexifyHeaders(headers);
  
  allData.forEach(function(row,i) {
    if (!row[headerIndexes["Status"]]) {
      var   htmlBody = 
        "Hi " + row[headerIndexes["Name"]] +",<br><br>" +
          "Thanks for responding to the Pollution Form!<br><br><br>" +
            "<strong>Here are your choices:</strong><br><br>" +
              "<strong>What city do you live in? </strong>" + row[headerIndexes["What city do you live in?"]] + "<br>" +
                "<strong>What % renewable are you at home? </strong>" + row[headerIndexes["What % renewable are you at home?"]] + "0%<br>" + 
                  "<strong>What are you most interested in? </strong>" + row[headerIndexes["What are you most interested in?"]] + "<br>" +
                    "<strong>What do you do to help the environment </strong>" + row[headerIndexes["What do you do to help the environment?"]] + "<br><br>" +
                      "Here's what you can do to be better:<br><br>";
      
       if (row[headerIndexes["What city do you live in?"]] === "san francisco" || row[headerIndexes["What city do you live in?"]] === "sf" || row[headerIndexes["What city do you live in?"]] === "San Francisco" || row[headerIndexes["What city do you live in?"]] === "SF") {
                         htmlBody = htmlBody + "The AI has detected that you live in SF based off of your answers! This is what it says about your area:<br>Because you live in SF, you get access to deciding where your energy comes from. Be sure to check out http://sfwater.org/index.aspx?page=963<br><br>"
                        } else {
                          htmlBody = htmlBody + "The AI has detected that you may not live in SF based on your answers. If you think this is an error, be sure to check this out: http://sfwater.org/index.aspx?page=963<br><br>"
                        }
      if (row[headerIndexes["What % renewable are you at home?"]] === "0" || row[headerIndexes["What % renewable are you at home?"]] === "1" || row[headerIndexes["What % renewable are you at home?"]] === "2" || row[headerIndexes["What % renewable are you at home?"]] === "3") {
        htmlBody = htmlBody + "The AI has detected that you are under 30% renewable! This is what it says about your electricity:<br>Because you put " + row[headerIndexes["What % renewable are you at home?"]] + "0% for your renuabbility percent. A cool way to become more renewable is to check out different electricity options in your area! For San Francisco, there is Clean Power SF, a company that allows people in SF ot be 100% renewable: http://sfwater.org/index.aspx?page=963<br><br>"
      } else {
        htmlBody = htmlBody + "The AI has detected that you have over 40% renewability! That probably means that you have already heard of Clean Power SF (http://sfwater.org/index.aspx?page=963)! Be sure to spread the word!<br><br>"
      }
      
      if (row[headerIndexes["How good would you describe yourself for the environment?"]] === "Great!  I do things that help the earth multiple times a day!") {
        htmlBody = htmlBody + "Great Job For Caring Greatly! You <strong>really</strong> help the environment! Teach others to be better and show what you do to the world!<br><br>"
      } else {
        if (row[headerIndexes["How good would you describe yourself for the environment?"]] === "Good! I do things that help the earth every day!") {
        htmlBody = htmlBody + "Great Job For Caring Greatly! You <strong>really</strong> help the environment! Teach others to be better and show what you do to the world! Also, maybe write to someone or a company telling them how they can do better as well.<br><br>"
      } else {
         if (row[headerIndexes["How good would you describe yourself for the environment?"]] === "Ok. I do things that help the earth every 2-4 days!") {
        htmlBody = htmlBody + "Great Job For Helping The Environment! You are around the average amount of people for the question, \"How good would you describe yourself for the environment?\". It may be hard to be different than everyone else, but our world is all we have. Maybe recycle & compost more and write to some bigger companies and tell them to be better as well.<br><br>"
      } else {
         if (row[headerIndexes["How good would you describe yourself for the environment?"]] === "Not that good. I do things that help the earth every week!") {
        htmlBody = htmlBody + "Don't feel bad, it's hard to care for the environment. Try looking up different things you can do to be better. Taking a shorter shower can even do the trick!<br><br>"
      } else {
        htmlBody = htmlBody + "Don't feel bad, it's hard to care for the environment. Learn from others and maybe websites what you can do to be better. Taking shorter showers is a great way to start!<br><br>"
      }
      }
      }
      }
      htmlBody = htmlBody + "The AI has detected that you have chosen " + row[headerIndexes["What are you most interested in?"]] + " as your answer to one of the questions! Great Choice! Here's some things to check out:<br>Al Gore's Website: https://algore.com<br>NASA Climate: https://climate.nasa.gov/evidence/<br>Water Calculator: https://watercalculator.org<br>The Ocean Cleanup: https://www.theoceancleanup.com/<br>Clean Power SF: https://sfwater.org/index.aspx?page=963<br><br>This Pollution AI is still in BETA and may not be 100% correct. If you see a problem, please reply to this email and we will get back to you shortly.<br><br><br>Thanks for being interested in helping the environment,<br><strong>The Pollution Group</strong><br><br>The Pollution Email AI © 2018 Thijs Simonian"
      var timestamp = sendEmail(row[headerIndexes["Email"]],htmlBody);
      thisSheet.getRange(i + 2, headerIndexes["Status"] + 1).setValue(timestamp);
    }
    else {
      Logger.log("No email sent for this row: " + i + 1);
    }
  });
}
function replyEmail() {
  var thisWorkbook = SpreadsheetApp.getActiveSpreadsheet();
  var thisSheet = thisWorkbook.getSheetByName('Form Responses 1');

  // get the data range of the sheet
  var allRange = thisSheet.getDataRange();
  
  // get all the data in this range
  var allData = allRange.getValues();
  
  // get the header row
  var headers = allData.shift();
  
  // create header index map
  var headerIndexes = indexifyHeaders(headers);
  
  allData.forEach(function(row,i) {
    if (!row[headerIndexes["Status"]]) {
      var   htmlBody = 
        "Hi " + row[headerIndexes["Name"]] +",<br><br>" +
          "Thanks for responding to the Pollution Form!<br><br><br>" +
            "<strong>Here is your followup to your form!</strong><br><br>" +
              row[headerIndexes["Custom Reply"]] + "<br>" +
                 "Thanks,<br>" +
                    "<strong>The Pollution Group</strong>";
      
      var timestamp = sendEmail(row[headerIndexes["Email"]],htmlBody);
      thisSheet.getRange(i + 2, headerIndexes["StatusReply"] + 1).setValue(timestamp);
    }
    else {
      Logger.log("No email sent for this row: " + i + 1);
    }
  });
}
  

/**
 * create index from column headings
 * @param {[object]} headers is an array of column headings
 * @return {{object}} object of column headings as key value pairs with index number
 */
function indexifyHeaders(headers) {
  
  var index = 0;
  return headers.reduce (
    // callback function
    function(p,c) {
    
      //skip cols with blank headers
      if (c) {
        // can't have duplicate column names
        if (p.hasOwnProperty(c)) {
          throw new Error('duplicate column name: ' + c);
        }
        p[c] = index;
      }
      index++;
      return p;
    },
    {} // initial value for reduce function to use as first argument
  );
}

/**
 * send email from GmailApp service
 * @param {string} recipient is the email address to send email to
 * @param {string} body is the html body of the email
 * @return {object} new date object to write into spreadsheet to confirm email sent
 */
function sendEmail(recipient,body) {
  
  GmailApp.sendEmail(
    recipient,
    "Pollution Group Form Comfirmation", 
    "",
    {
      htmlBody: body
    }
  );
  
  return new Date();
}
