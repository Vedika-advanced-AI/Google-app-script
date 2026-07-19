function sendSecurityEmail(email, subject, bodyHtml) {
  const finalHtml = `
    <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #eee; border-radius: 8px;">
      <div style="text-align: center; margin-bottom: 20px;">
        <img src="${CONFIG.LOGO_URL}" width="150" alt="Vedika AI Studio" />
      </div>
      <h3 style="color: #333;">${subject}</h3>
      <p style="color: #555; font-size: 15px;">${bodyHtml}</p>
      <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
      <p style="font-size: 12px; color: #999; text-align: center;">This is an automated security alert from AI Studio.</p>
    </div>
  `;
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: finalHtml });
}

function logUserAction(email, eventType, device, ip, details) {
  const rowIndex = findUserByEmail(email);
  if (rowIndex > 0) {
    const data = getMasterSheet().getDataRange().getValues();
    const userSheetId = data[rowIndex][8]; // User_Sheet_ID column
    const userSheet = SpreadsheetApp.openById(userSheetId).getActiveSheet();
    
    const timestamp = new Date().toISOString();
    userSheet.appendRow([timestamp, eventType, device, ip, details]);
    
    // Security email is sent from verifyLogin() for LOGIN events
    // This function handles logging only
  }
}

function logApiUsage(params) {
  const masterSheet = getMasterSheet();
  const data = masterSheet.getDataRange().getValues();
  
  // Find user by email or apiKey
  let userRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === params.email || data[i][5] === params.apiKey) {
      userRowIndex = i;
      break;
    }
  }
  
  if (userRowIndex === -1) {
    return {status: "error", message: "User not found"};
  }
  
  const rowNum = userRowIndex + 1;
  const currentRequests = Number(data[userRowIndex][9]) || 0; // Column J: Total_Requests
  const currentTokens = Number(data[userRowIndex][10]) || 0; // Column K: Total_Tokens
  const currentErrors = Number(data[userRowIndex][11]) || 0; // Column L: Error_Count
  
  const tokensUsed = Number(params.tokens) || 0;
  const statusCode = params.status_code || 200;
  
  // Update Master Sheet: Increment requests and add tokens
  masterSheet.getRange(rowNum, 10).setValue(currentRequests + 1);
  masterSheet.getRange(rowNum, 11).setValue(currentTokens + tokensUsed);
  
  // Increment Error_Count if status code is 4xx or 5xx
  if (statusCode >= 400 && statusCode < 600) {
    masterSheet.getRange(rowNum, 12).setValue(currentErrors + 1);
  }
  
  // Log to user's personal spreadsheet (model-agnostic logging)
  const userSheetId = data[userRowIndex][8];
  const userSheet = SpreadsheetApp.openById(userSheetId).getActiveSheet();
  const timestamp = new Date().toISOString();
  const model = params.model_name || params.model || "Unknown"; // Accept model_name from Hugging Face Gateway
  userSheet.appendRow([timestamp, "API_USAGE", model, statusCode, `Tokens: ${tokensUsed}`]);
  
  return {status: "success", message: "Usage logged"};
}

/**
 * Calculate user health score based on error rate
 * @param {string} email - User email
 * @returns {object} Health score percentage
 */
function calculateHealthScore(email) {
  const rowIndex = findUserByEmail(email);
  if (rowIndex === -1) {
    return {status: "error", message: "User not found"};
  }
  
  const data = getMasterSheet().getDataRange().getValues();
  const totalRequests = Number(data[rowIndex][9]) || 0; // Column J: Total_Requests
  const errorCount = Number(data[rowIndex][11]) || 0; // Column L: Error_Count
  
  if (totalRequests === 0) {
    return {status: "success", health_score: 100, message: "No requests yet"};
  }
  
  const healthScore = ((totalRequests - errorCount) / totalRequests) * 100;
  return {status: "success", health_score: Math.round(healthScore * 100) / 100};
}

/**
 * Retrieve full API usage logs for a specific user
 * @param {string} email - User email
 * @returns {array} JSON array of historical API request data
 */
function getFullLogs(email) {
  const rowIndex = findUserByEmail(email);
  if (rowIndex === -1) {
    return {status: "error", message: "User not found"};
  }
  
  const data = getMasterSheet().getDataRange().getValues();
  const userSheetId = data[rowIndex][8]; // Column I: User_Sheet_ID
  const userSheet = SpreadsheetApp.openById(userSheetId).getActiveSheet();
  const sheetData = userSheet.getDataRange().getValues();
  
  // Skip header row, filter only API_USAGE entries
  const logs = [];
  for (let i = 1; i < sheetData.length; i++) {
    if (sheetData[i][1] === "API_USAGE") {
      logs.push({
        timestamp: sheetData[i][0],
        model: sheetData[i][2],
        tokens: sheetData[i][4] ? sheetData[i][4].replace("Tokens: ", "") : 0,
        status: sheetData[i][3] || 200
      });
    }
  }
  
  return {status: "success", logs: logs, total_count: logs.length};
}
