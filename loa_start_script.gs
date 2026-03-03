// function to send an email that a LOA is starting
function loa_start(){
  // access the Google Sheet
  let wb = SpreadsheetApp.openByUrl('google-sheet-url');

  // access the sheet with the data
  let sheet = wb.getSheetByName('sheet_name');

  // creating data array
 let data = sheet.getRange('A2:I62').getDisplayValues();

 // declaring empty array dataArray
 const dataArray = []

 // for loop to iterate through nth elements of innerArray and assigning to dataArray
 for (  let i = 0; i <  9; i++ ){
  dataArray.push(data.map(innerArray => innerArray[i]))
}


 // creating object
 const dataObject = {
  'Last Name': dataArray[0],
  'First Name': dataArray[1],
  'Assignment': dataArray[2],
  'Location': dataArray[3],
  'Begin Date': dataArray[4],
  'End Date': dataArray[5],
  'Replacement': dataArray[6],
  'School Contact': dataArray[7],
  'Mailing List': dataArray[8]

 }

// creating dynamic start and end date variabnles
  let today = new Date();
  let start_date = new Date(today.getFullYear(),today.getMonth(),1); 
  let end_date = new Date(today.getFullYear(),today.getMonth() +1 ,0);

  // create template object for dynamically constructing html
  let htmlTemplate = HtmlService.createTemplateFromFile('loa_start');

 // for loop to to send emails
  for (let i = 0 ; i < dataObject['First Name'].length;i++ ){
    // returning all LOA starting within range of start_date and end_date
    if (new Date(dataObject['Begin Date'][i]) >= start_date && new Date(dataObject['Begin Date'][i]) <= end_date && dataObject['Location'][i] == 'RHS' ) {

      // define html variables
      htmlTemplate.firstName = dataObject['First Name'][i];
      htmlTemplate.lastName = dataObject['Last Name'][i];
      htmlTemplate.assignment= dataObject['Assignment'][i];
      htmlTemplate.location = dataObject['Location'][i];
      htmlTemplate.startDate = dataObject['Begin Date'][i];
      htmlTemplate.endDate = dataObject['End Date'][i];
      htmlTemplate.replacement = dataObject['Replacement'][i];

  // evaluates the template and returns htmloutput object
    let htmlForEmail = htmlTemplate.evaluate().getContent();

    // sending LOA Email
      GmailApp.sendEmail(dataObject['Mailing List'][i],
      `Leave of Absence Start Alert: ${dataObject['First Name'][i]} ${dataObject['Last Name'][i]}`,
      'This email contains html',{htmlBody: htmlForEmail})
    };
  };
}
 
