/**
 * Router.js
 * The single Apps Script web-app endpoint, exposing an action-based API.
 *
 * WHY ACTION-BASED RATHER THAN REST
 * An Apps Script web app is one URL driven by doGet/doPost. It cannot serve
 * path segments (GET /students/:id) or PUT/DELETE verbs, so REST routes cannot
 * be expressed literally. The agreed equivalent is:
 *
 *   GET  ?action=health
 *   POST {"action":"students.create","payload":{...}}
 *
 * Request pipeline:
 *   doGet/doPost -> handleRequest_ -> parseRequest_ -> route table -> action
 *                -> standard response envelope (Response.js)
 *
 * IMPLEMENTATION STATUS (Phase 1): only CONFIG.ACTIONS.HEALTH is routed.
 * Every other action name in CONFIG.ACTIONS is a reserved identifier for a
 * later phase and currently returns NOT_FOUND.
 */

/**
 * GET entry point.
 * @param {Object} e Apps Script event object.
 * @return {ContentService.TextOutput} JSON response envelope.
 */
function doGet(e) {
  return handleRequest_(e);
}

/**
 * POST entry point.
 * @param {Object} e Apps Script event object.
 * @return {ContentService.TextOutput} JSON response envelope.
 */
function doPost(e) {
  return handleRequest_(e);
}

/* ==========================================================================
 * Route table
 * ======================================================================== */

/**
 * Build the action -> handler table.
 *
 * Built on demand rather than as a top-level constant, so the router never
 * depends on the order in which Apps Script evaluates the project files.
 *
 * @return {Object<string, Function>} Map of action name to handler.
 */
function getRoutes_() {
  const routes = {};
  routes[CONFIG.ACTIONS.HEALTH] = handleHealth_;
  return routes;
}

/**
 * @return {string[]} Actions the router currently answers.
 */
function listAvailableActions_() {
  return Object.keys(getRoutes_());
}

/* ==========================================================================
 * Request parsing
 * ======================================================================== */

/**
 * Parse an Apps Script event into a normalised {action, payload} request.
 *
 * Accepts:
 *   - GET  query parameters, where every parameter except `action` becomes part
 *          of the payload (e.g. ?action=students.get&id=STU-1)
 *   - POST a single JSON body of exactly {"action":"...","payload":{...}}
 *
 * Both may be combined; body payload values win over query parameters.
 *
 * @param {Object} e Apps Script event object.
 * @return {{action: string, payload: Object}} The normalised request.
 * @throws {Error} VALIDATION_ERROR for malformed or incomplete requests.
 */
function parseRequest_(e) {
  const params = (e && e.parameter) || {};

  const queryPayload = {};
  Object.keys(params).forEach(function (key) {
    if (key !== 'action') queryPayload[key] = params[key];
  });

  let action = toTrimmedString_(params.action);
  let bodyPayload = null;

  const rawBody = e && e.postData && e.postData.contents;
  if (isNonEmptyString_(rawBody)) {
    let body;
    try {
      body = JSON.parse(rawBody);
    } catch (err) {
      throwError_(
        'Request body is not valid JSON.',
        ERROR_CODES.VALIDATION_ERROR,
        { received: String(rawBody).slice(0, 200) }
      );
    }

    if (body === null || typeof body !== 'object' || Array.isArray(body)) {
      throwError_(
        'Request body must be a JSON object like {"action":"...","payload":{...}}.',
        ERROR_CODES.VALIDATION_ERROR,
        { receivedType: Array.isArray(body) ? 'array' : typeof body }
      );
    }

    const unexpected = Object.keys(body).filter(function (key) {
      return key !== 'action' && key !== 'payload';
    });
    if (unexpected.length) {
      throwError_(
        'Unexpected key(s) in request body: ' + unexpected.join(', ') + '. Send exactly {"action":"...","payload":{...}}.',
        ERROR_CODES.VALIDATION_ERROR,
        { unexpectedKeys: unexpected }
      );
    }

    if (isNonEmptyString_(body.action)) {
      action = toTrimmedString_(body.action);
    }

    if (body.payload !== undefined && body.payload !== null) {
      if (typeof body.payload !== 'object' || Array.isArray(body.payload)) {
        throwError_(
          '"payload" must be a JSON object.',
          ERROR_CODES.VALIDATION_ERROR,
          { receivedType: Array.isArray(body.payload) ? 'array' : typeof body.payload }
        );
      }
      bodyPayload = body.payload;
    }
  }

  if (action === '') {
    throwError_(
      'Missing required "action". Use ?action=health on GET, or {"action":"...","payload":{...}} in a POST body.',
      ERROR_CODES.VALIDATION_ERROR,
      { availableActions: listAvailableActions_() }
    );
  }

  return {
    action: action,
    payload: Object.assign({}, queryPayload, bodyPayload || {})
  };
}

/* ==========================================================================
 * Dispatch
 * ======================================================================== */

/**
 * Shared pipeline for doGet and doPost.
 *
 * Centralises error handling, so every expected (structured) and unexpected
 * error becomes a standard failure envelope. The frontend therefore never
 * receives an HTML error page or a raw stack trace.
 *
 * @param {Object} e Apps Script event object.
 * @return {ContentService.TextOutput} JSON response envelope.
 */
function handleRequest_(e) {
  try {
    const request = parseRequest_(e);
    const handler = getRoutes_()[request.action];

    if (!handler) {
      return jsonResponse(
        failure(
          'Unknown action: "' + request.action + '".',
          ERROR_CODES.NOT_FOUND,
          { action: request.action, availableActions: listAvailableActions_() }
        )
      );
    }

    const envelope = handler(request.payload, request);

    // Guard against a handler returning something that is not an envelope,
    // which would otherwise break the frontend contract silently.
    if (envelope === null || typeof envelope !== 'object' || typeof envelope.success !== 'boolean') {
      throwError_(
        'Action "' + request.action + '" returned an invalid response envelope.',
        ERROR_CODES.SERVER_ERROR,
        { action: request.action }
      );
    }

    return jsonResponse(envelope);
  } catch (error) {
    return jsonResponse(
      failure(
        error && error.message ? error.message : 'Unexpected server error.',
        (error && error.code) || ERROR_CODES.SERVER_ERROR,
        (error && error.details) || null
      )
    );
  }
}

/* ==========================================================================
 * Actions
 * ======================================================================== */

/**
 * GET ?action=health
 *
 * Verifies the full request chain end to end:
 *   React / browser -> Apps Script web app -> SpreadsheetApp.openById -> Sheets
 *
 * It genuinely opens the configured spreadsheet and lists the tabs that
 * actually exist, so a populated `sheets` array is real proof of access from
 * web-app execution context. It deliberately exposes no credentials, tokens,
 * or secrets -- only the spreadsheet ID, which the frontend needs anyway.
 *
 * `missingSheets` lists tabs Config.js expects but that do not exist yet.
 * That is a data-setup problem rather than a service failure, so `status`
 * stays 'online' as long as the spreadsheet itself was reachable.
 *
 * @return {Object} Success envelope with service diagnostics.
 */
function handleHealth_() {
  const expectedSheets = Object.keys(CONFIG.SHEETS).map(function (key) {
    return CONFIG.SHEETS[key];
  });

  // Throws SERVER_ERROR (caught by handleRequest_) when the spreadsheet cannot
  // be opened, which is exactly the failure this action exists to expose.
  const spreadsheet = getSpreadsheet_();
  const sheets = listSheetNames_();

  const missingSheets = expectedSheets.filter(function (name) {
    return sheets.indexOf(name) === -1;
  });

  return success(
    {
      service: CONFIG.SERVICE_NAME,
      status: 'online',
      spreadsheetId: CONFIG.SHEET_ID,
      spreadsheetName: spreadsheet.getName(),
      sheets: sheets,
      expectedSheets: expectedSheets,
      missingSheets: missingSheets,
      timeZone: getTimeZone_(),
      serverTime: nowIso_()
    },
    'Service healthy'
  );
}