/**
 * FASHION FIZZ BD - SIMPLE AUTH SYSTEM
 */

function checkLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  // Skip headers (i=1), loop through rows
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
      return { status: 'success', user: username };
    }
  }
  return { status: 'fail', message: 'Invalid Username or Password' };
}

