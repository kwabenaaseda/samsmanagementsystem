/**
 * Api.js
 *
 * Entry point for the Apps Script Execution API (scripts.run).
 *
 * WHY THIS EXISTS
 * The web-app transport (/exec, doGet/doPost) cannot serve authenticated
 * cross-origin requests from the standalone React frontend: Google's web-app
 * auth layer answers unauthenticated callers with a 302/401 before doPost ever
 * runs, and that response carries no CORS headers. The Execution API
 * (POST https://script.googleapis.com/v1/scripts/{scriptId}:run) replaces it:
 * identity travels in the OAuth bearer token, and googleapis.com serves proper
 * CORS headers for authorized JavaScript origins.
 *
 * scripts.run can only invoke top-level functions with JSON-serializable
 * parameters, so this adapter:
 *   1. receives (action, payload, auth) from the frontend,
 *   2. synthesizes the same event object parseRequest_() already understands,
 *   3. delegates to the EXISTING handleRequest_() pipeline (Router.js) so the
 *      route table, __auth stripping, and response envelopes are identical,
 *   4. unwraps the ContentService TextOutput back into a plain object,
 *      which scripts.run returns to the caller as response.result.
 *
 * No route, handler, or business module is modified by this file.
 *
 * IDENTITY: the frontend passes the GIS access token as `auth` (third
 * parameter) AND sends the same token as the request's bearer token. The
 * backend's Auth.js verifies the token server-side against Google's userinfo
 * endpoint — exactly as it does for the web-app transport. Nothing here reads
 * or trusts a caller-supplied email.
 */

/**
 * Execute one SAMS API action through the existing router pipeline.
 *
 * @param {string} action Action name, e.g. "auth.me" or "students.list".
 * @param {Object=} payload Action payload (may include anything the handlers accept).
 * @param {Object=} auth Reserved auth transport block, { access_token: string }.
 * @return {Object} The standard SAMS envelope { success, message, data?... }.
 */
function apiRun(action, payload, auth) {
  const safeAction = typeof action === 'string' ? action : '';
  const safePayload = payload !== null && typeof payload === 'object' && !Array.isArray(payload)
    ? payload
    : {};
  const safeAuth = auth !== null && typeof auth === 'object' && !Array.isArray(auth)
    ? auth
    : {};

  // __auth must travel INSIDE the payload, never beside it.
  //
  // parseRequest_() (Router.js) accepts exactly two top-level body keys and
  // rejects everything else — a strictness worth keeping, since it stops
  // arbitrary fields being smuggled past the parser. handleRequest_() then
  // reads the reserved block from request.payload.__auth, installs the token in
  // the request-scoped slot, and deletes __auth before dispatch so no handler
  // ever sees transport data.
  //
  // Placing the block inside the payload therefore satisfies both rules and
  // keeps ONE parser and ONE stripping rule for the web-app and Execution API
  // transports alike.
  const accessToken =
    typeof safeAuth.access_token === 'string' ? safeAuth.access_token : '';
  const bodyPayload =
    accessToken === ''
      ? safePayload
      : Object.assign({}, safePayload, { __auth: safeAuth });

  const envelope = handleRequest_({
    parameter: { action: safeAction },
    postData: {
      contents: JSON.stringify({
        action: safeAction,
        payload: bodyPayload,
      }),
    },
  });

  // handleRequest_ always returns a JSON TextOutput envelope (Response.js),
  // so JSON.parse here cannot fail for any handled path.
  return JSON.parse(envelope.getContent());
}
