function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Alat otomasi✨')
      .addItem('Buat TOR', 'createTOR')
      .addItem('Buat form presensi', 'createForm')
      .addToUi();
}

function createForm() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startRow = 14; // Starting row of the data
  var numRows = 28; // Number of forms to create
  
  var folderId = "..."; // Replace with the desired folder ID
  
  var folder = DriveApp.getFolderById(folderId);
  
  for (var i = 0; i < numRows; i++) {
    var rowIndex = startRow + i;
    var formTitle = sheet.getRange(rowIndex, 4).getValue();
    var thirdItemTitle = "Custom question sentence here that will be paired with specific rows value '" + sheet.getRange(rowIndex, 4).getValue() + "'!";
    var formUrl = sheet.getRange(rowIndex, 9).getValue();
    
    if (formUrl === "" && sheet.getRange(rowIndex, 8).getValue() !== "") {
      var form = FormApp.create(formTitle);
      form.setDescription("Specific form description here\n\nCan be used to put instruction.\n\nOr other greetings.\n\nAnd other info in general");

      form.setCollectEmail(true); // Activate the "Collect email addresses" feature
      // Activate the "Collect email addresses" feature and set it to "Responder input"
      form.setCollectEmail(true);
      form.setRequireLogin(false);

      // Question 1 (Short answer text)
      var nameItem = form.addTextItem();
      nameItem.setTitle('Name').setRequired(true);

      // Question 2 (Linear scale)
      var scaleItem = form.addScaleItem();
      scaleItem.setTitle('Custom statement here #1')
        .setBounds(1, 5)
        .setLabels(lower='Highly disagree', upper='Highly agree')
        .setRequired(true);

      // Question 3 (Linear scale)
      var scaleItem = form.addScaleItem();
      scaleItem.setTitle('Custom statement here #2')
        .setBounds(1, 5)
        .setLabels(lower='Highly disagree', upper='Highly agree')
        .setRequired(true);

      // Question 4 (Linear scale)
      var scaleItem = form.addScaleItem();
      scaleItem.setTitle('Custom statement here #3')
        .setBounds(1, 5)
        .setLabels(lower='Highly disagree', upper='Highly agree')
        .setRequired(true);

      // Question 5 (Linear scale)
      var scaleItem = form.addScaleItem();
      scaleItem.setTitle('Custom statement here #4')
        .setBounds(1, 5)
        .setLabels(lower='Highly disagree', upper='Highly agree')
        .setRequired(true);

      // Question 6 (Linear scale)
      var scaleItem = form.addScaleItem();
      scaleItem.setTitle('Custom statement here #5')
        .setBounds(1, 5)
        .setLabels(lower='Highly disagree', upper='Highly agree')
        .setRequired(true);

      // Question 7 (Long answer text)
      var longAnswerItem = form.addParagraphTextItem();
      longAnswerItem.setTitle(thirdItemTitle).setRequired(true);

      // Question 8 (Long answer text)
      var feedbackItem = form.addParagraphTextItem();
      feedbackItem.setTitle('Custom statement here #6')
        .setRequired(true)
        .setHelpText("Custom question description here");

      formUrl = form.getPublishedUrl()
      shortenUrl = form.shortenFormUrl(formUrl)
      
      sheet.getRange(rowIndex, 9).setValue(shortenUrl); // Store the shortened Google Form link in column I
    }
  }
}
