# Leave of Absence Google Apps Email Script 

A district was having an issue in which staff records were not being maintained correctly in the SIS, affecting data accuracy, state reporting, and data security, e.g. teachers that were on a leave of absence were still assigned to classes, teachers and staff still having access to the SIS despite being on leave, incorrect course transaction dates, etc.

To resolve the issue, I created a Google Apps Script that sends an email to registrars at each school informing them of the upcoming or ending leave of absence.

## Script Overview
The script automates the process of notifying relevant stakeholders when a Leave of Absence (LOA) is scheduled to begin and end. It parses employee data from a centralized Google Sheet, filters for specific criteria, and sends personalized HTML emails via Gmail.

### Core Functionality
~~~ javascript
// creating function
function loa_start(){

 // access the Google Sheet
  let wb = SpreadsheetApp.openByUrl('google-sheet-url');

  // access the sheet with the data
  let sheet = wb.getSheetByName('sheet_name');

  // creating data array
 let data = sheet.getRange('A2:I62').getDisplayValues();

 // declaring empty array dataArray
 const dataArray = []

}
~~~
Extracts data from the centralized Google Sheet by accessing the sheet via url, name, and retieving values based on the range specficied in the *sheet.getRange()* method

~~~ javascript
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
~~~
After the data has been extracted into an array, the data is then transformed into an object using via mapping using the arrow function.

~~~javascript

// creating dynamic start and end date variabnles
  let today = new Date();
  let start_date = new Date(today.getFullYear(),today.getMonth(),1); 
  let end_date = new Date(today.getFullYear(),today.getMonth() +1 ,0);

// create template object for dynamically constructing html
  let htmlTemplate = HtmlService.createTemplateFromFile('loa_start');

// for loop to to send emails
  for (let i = 0 ; i < dataObject['First Name'].length;i++ ){
    // returning all LOA starting within range of start_date and end_date
    if (new Date(dataObject['Begin Date'][i]) >= start_date && new Date(dataObject['Begin Date'][i]) <= end_date ) {

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
~~~
- The *start_date* and *end_date* variables are created for filtering of the data in the for loop.
- An HTML template is imported from both templates created for the LOA starting and ending
- A for loop is created to iterate through the object *i* times, not to exceed the length of the number of values in the object. There is also a condition to ensure that the data returned has a beginning LOA date within the specified range of *start_date* and *end_date* variables at a specific. When those conditions are met, a set of dynamic HTML variables are defined for each iteration through the for loop and used to create an email that is sent to a specified contact informing them of the employee that is starting or ending their LOA.

