function myFunction() {
  var threads = GmailApp.search('label:unseen');
  var threadsCount = threads.length;

  var labelUnseen = GmailApp.getUserLabelByName('unseen');
  var labelSeen = GmailApp.getUserLabelByName("seen");

　var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow() + 1;

  var itemCount = 2;
  
  if (threadsCount == 0) {
    return;
  }

  var values = [];
  for (var i = 0; i < threadsCount; i++) {
　　values[i] = [];

    var counter = 0;
    var contents = threads[i].getMessages()[0].getBody().split("");

    for (var j = 0; j < contents.length; j++) {
    　if (contents[j].match(/name|age/)) {
        var parts = contents[j].split(":");
        values[i][counter++] = parts[1];
        if (parts[0] == "\nage") {break;}

      }
    }

    threads[i].addLabel(labelSeen);
    threads[i].removeLabel(labelUnseen);
  }

  // getRange(row, column, numRows, numColumns)
  var range = sheet.getRange(lastRow, 1 , threadsCount, itemCount);
  range.setValues(values);
}


function onEdit(e) {
  var column = e.range.getColumn();
  var color="white"

  if (column != 3) {return;}

  if (e.value == "o") {
    color="gray";
  }
  
  var row = e.range.getRow();
  var column = e.range.getColumn();

  e.source.getActiveSheet().getRange(row, 1, 1, column).setBackground(color);
}
