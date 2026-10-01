/************************************************
 * EDEN BI V4
 * Core/Auth.js
 *
 * Auth functions currently live in Code.js:
 *   loginUser, logoutUser, validateAuthToken_, getUsersSheet_
 *
 * IMPORTANT: do not call functions at the top level of any .js file.
 * Apps Script runs every file's top-level code on each request, so a bare
 * call such as validateAuthToken_() runs before login and fails with
 * "Access denied. Please login again."
 ************************************************/
