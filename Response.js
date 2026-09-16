/**
 * Response.js
 * The single response envelope returned by every API action.
 *
 * CONTRACT FOR THE FRONTEND
 * Apps Script web apps cannot reliably return arbitrary HTTP status codes, so
 * the HTTP status is always 200. The frontend must therefore branch on the
 * application-level fields, never on the HTTP status:
 *
 *   payload.success  -> boolean, true for success, false for failure
 *   payload.message  -> human-readable summary (safe to display)
 *   payload.data     -> result payload, present when success === true
 *   payload.error    -> machine-readable error code, present when success === false
 *   payload.details  -> optional machine-readable context (e.g. missing fields)
 */

/**
 * Application-level error codes.
 * These are the only values placed in payload.error.
 */
const ERROR_CODES = {
  /** Caller sent something invalid or incomplete. */
  VALIDATION_ERROR: 'VALIDATION_ERROR',
  /** Caller is not identified / not logged in. */
  UNAUTHORIZED: 'UNAUTHORIZED',
  /** Caller is identified but not permitted to perform this action. */
  FORBIDDEN: 'FORBIDDEN',
  /** A requested record, sheet, or action does not exist. */
  NOT_FOUND: 'NOT_FOUND',
  /** The request conflicts with current state (e.g. duplicate ID). */
  CONFLICT: 'CONFLICT',
  /** Unexpected backend failure. */
  SERVER_ERROR: 'SERVER_ERROR'
};

function success(data = null, message = 'Success') {
  return {
    success: true,
    message,
    data
  };
}

/**
 * Build a failure envelope.
 *
 * Signature preserved from the original implementation: failure(message, error).
 * `details` is an additive optional third argument, so existing callers
 * (e.g. Router.js) keep working unchanged.
 *
 * NOTE: the default error code was previously the literal 'BAD_REQUEST', which
 * is not part of the taxonomy. It is now VALIDATION_ERROR. No existing caller
 * relied on the default -- Router.js always passed its code explicitly.
 */
function failure(message, error = ERROR_CODES.VALIDATION_ERROR, details = null) {
  return {
    success: false,
    message,
    error,
    details
  };
}

/**
 * Serialize a response envelope as a JSON ContentService result.
 * Signature preserved from the original implementation.
 */
function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}