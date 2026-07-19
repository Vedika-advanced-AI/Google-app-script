function doPost(e) {
  let response = {status: "error", message: "Bad Request"};
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;

    switch (action) {
      // 1. Frontend Actions
      case 'request_register_otp': response = requestOTP(params.email, false); break;
      case 'verify_register': response = verifyRegistration(params); break;
      case 'request_login_otp': response = requestOTP(params.email, true); break;
      case 'verify_login': response = verifyLogin(params); break;
      case 'view_api_key': response = fetchApiKey(params.email); break;
      case 'log_out': 
        logUserAction(params.email, "LOGOUT", params.device, params.ip, "User logged out manually");
        response = {status: "success", message: "Logged out"}; 
        break;

      // 2. Admin Actions (Status change)
      case 'update_status': response = changeAccountStatus(params.email, params.status); break;

      // 3. Hugging Face Sync Action (Super Fast API Check)
      case 'hf_verify_key':
        const masterSheet = getMasterSheet();
        const data = masterSheet.getDataRange().getValues();
        // Skip header row, search for matching API key in Column F (index 5)
        let userRow = null;
        for (let i = 1; i < data.length; i++) {
          if (data[i][5] === params.apiKey) {
            userRow = data[i];
            break;
          }
        }
        
        if (userRow && userRow[6] === "Active") { // Column G: Status
          response = {status: "authorized", user: userRow[1], role: userRow[3]};
        } else {
          response = {status: "unauthorized", message: "Key invalid, blocked, or deleted."};
        }
        break;

      // 4. API Usage Logging Endpoint (Hugging Face Background Call)
      case 'log_api_usage':
        response = logApiUsage(params);
        break;
    }
  } catch (err) {
    response = {status: "error", message: err.toString()};
  }
  
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}
