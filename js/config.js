// ---------------------------------------------------------------------------------------------
//  DAL3 Virtual Lab settings. After deploying the Google Apps Script web app (see SETUP.md),
//  paste its URL below, commit, and push. Until then the lab runs in DEMO mode (saved on this device only).
// ---------------------------------------------------------------------------------------------
export const CONFIG = {
  API_URL: "https://script.google.com/macros/s/AKfycbxidw-w5SoBPa0PIyL5zM_2SItm4csHHjER8QX9Vj6VB_1xeiKHTG7tBkNIbOQ0mx8/exec",                       // e.g. "https://script.google.com/macros/s/AKfycb.../exec"
  CLASS_CODE_HINT: "DAL3-2026",      // shown on the login screen
  // Demo videos are set in js/videos.js (Google Drive links + which lesson part shows which chapter).
  AUTOSAVE_SECONDS: 60,              // how often a student's live work is sent to the trainer board
  BOARD_REFRESH_SECONDS: 10,
};
