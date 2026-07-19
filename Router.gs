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
        const data = getMasterSheet().getDataRange().getValues();
        const userRow = data.find(row => row[5] === params.apiKey); // Column F: API_Key
        
        if (userRow && userRow[6] === "Active") { // Column G: Status
          response = {status: "authorized", user: userRow[1], role: userRow[3]};
        } else {
          response = {status: "unauthorized", message: "Key invalid, blocked, or deleted."};
        }
        break;
    }
  } catch (err) {
    response = {status: "error", message: err.toString()};
  }
  
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}
