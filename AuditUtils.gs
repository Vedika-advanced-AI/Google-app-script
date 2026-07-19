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
    
    if (eventType === "LOGIN") {
      const emailBody = `A new login was detected on your account.<br><br><b>Device:</b> ${device}<br><b>IP Address:</b> ${ip}<br><b>Time:</b> ${timestamp}<br><br>If this wasn't you, please contact support immediately.`;
      sendSecurityEmail(email, "Security Alert: New Login Detected", emailBody);
    }
  }
}
