function doGet(e) {
  return jsonResponse(
    success({
      service: 'School Management System API',
      status: 'online'
    })
  );
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');

    return jsonResponse(
      success(body, 'Request received')
    );

  } catch (error) {
    return jsonResponse(
      failure(error.message, 'SERVER_ERROR')
    );
  }
}