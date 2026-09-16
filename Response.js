function success(data = null, message = 'Success') {
  return {
    success: true,
    message,
    data
  };
}

function failure(message, error = 'BAD_REQUEST') {
  return {
    success: false,
    message,
    error
  };
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}