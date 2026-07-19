function requestOTP(email, isLogin) {
  const rowIndex = findUserByEmail(email);
  if (isLogin && rowIndex === -1) return {status: "error", message: "Email not registered."};
  if (!isLogin && rowIndex > 0) return {status: "error", message: "Email already registered."};
  
  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  CacheService.getScriptCache().put("OTP_" + email, otp, CONFIG.OTP_EXPIRY);
  
  sendSecurityEmail(email, "Your AI Studio OTP", `Your verification code is: <h2 style="color:#007BFF; letter-spacing: 2px;">${otp}</h2><br>This code is valid for 5 minutes.`);
  return {status: "success", message: "OTP sent successfully."};
}

function verifyRegistration(params) {
  const cachedOtp = CacheService.getScriptCache().get("OTP_" + params.email);
  if (cachedOtp !== params.otp) return {status: "error", message: "Invalid or expired OTP."};
  
  // Student/Organization Logic
  const orgName = (params.role === "Student") ? "N/A" : (params.organization || "Not Provided");
  
  const userSheetId = createUserSheet(params.email);
  const newApiKey = "sk_live_" + Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  
  getMasterSheet().appendRow([
    Utilities.getUuid(), params.email, params.name, params.role, orgName, 
    newApiKey, "Active", "FALSE", userSheetId, 0, 0
  ]);
  
  CacheService.getScriptCache().remove("OTP_" + params.email);
  sendSecurityEmail(params.email, "Welcome to AI Studio", `Hi ${params.name}, your account is successfully created!`);
  return {status: "success", message: "Registration complete.", name: params.name, requests: 0, tokens: 0};
}

function verifyLogin(params) {
  const cachedOtp = CacheService.getScriptCache().get("OTP_" + params.email);
  if (cachedOtp !== params.otp) return {status: "error", message: "Invalid or expired OTP."};
  
  const rowIndex = findUserByEmail(params.email);
  const data = getMasterSheet().getDataRange().getValues()[rowIndex];
  
  if (data[6] !== "Active") return {status: "error", message: `Account is ${data[6]}`}; // Blocked/Deleted check
  
  logUserAction(params.email, "LOGIN", params.device, params.ip, "Successful login");
  CacheService.getScriptCache().remove("OTP_" + params.email);
  sendSecurityEmail(params.email, "Security Alert: Login Successful", 
    `Hi ${data[2]},<br><br>Your account was just accessed.<br><b>Device:</b> ${params.device}<br><b>IP Address:</b> ${params.ip}<br><b>Time:</b> ${new Date().toISOString()}<br><br>If this wasn't you, please contact support immediately.`);
  return {status: "success", message: "Login verified.", name: data[2], requests: data[9] || 0, tokens: data[10] || 0};
}

// 1-Time Show API Key Logic
function fetchApiKey(email) {
  const master = getMasterSheet();
  const rowIndex = findUserByEmail(email);
  if (rowIndex === -1) return {status: "error", message: "User not found."};
  
  const rowNum = rowIndex + 1;
  const isKeySeenValue = master.getRange(rowNum, 8).getValue(); // Column H: Key_Seen
  const fullApiKey = master.getRange(rowNum, 6).getValue(); // Column F: API_Key
  
  // Handle both boolean and string representations of "seen" status
  const isKeySeen = (isKeySeenValue === true || isKeySeenValue === "TRUE" || isKeySeenValue === "true");
  
  if (!isKeySeen) {
    master.getRange(rowNum, 8).setValue("TRUE"); // Mark as seen (use string for consistency)
    return {status: "success", api_key: fullApiKey, message: "Save this key! It won't be shown fully again."};
  } else {
    const maskedKey = fullApiKey.substring(0, 12) + "****************" + fullApiKey.substring(fullApiKey.length - 4);
    return {status: "success", api_key: maskedKey, message: "Key already viewed."};
  }
}

// Account Blocking/Unblocking Logic
function changeAccountStatus(email, newStatus) { // newStatus: 'Active', 'Blocked', 'Deleted'
  const rowIndex = findUserByEmail(email);
  if (rowIndex !== -1) {
    getMasterSheet().getRange(rowIndex + 1, 7).setValue(newStatus); // Column G: Status
    return {status: "success", message: `Account marked as ${newStatus}`};
  }
  return {status: "error", message: "User not found"};
}
