/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  const ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('BattleWriter');
  DocumentApp.getUi().showSidebar(ui);
}

function loadState() {
  const state = {
    lastPoll: Date.now(),
    lastWordCount: 0,
  };

  const rawState = PropertiesService.getDocumentProperties().getProperties();

  if (typeof rawState.lastPoll === 'string') {
    state.lastPoll = parseInt(rawState.lastPoll);
  }

  if (typeof rawState.lastWordCount === 'string') {
    state.lastWordCount = parseFloat(rawState.lastWordCount);
  }

  if (typeof rawState.currentSprint === 'string') {
    state.currentSprint = JSON.parse(rawState.currentSprint);
  }

  for (const key in rawState) {
    if (key.startsWith('sprint:')) {
      state[key] = JSON.parse(rawState[key]);
    }
  }

  return state;
}

function storeState(state) {
  const rawState = {};

  rawState.lastPoll = '' + state.lastPoll;
  rawState.lastWordCount = '' + state.lastWordCount;

  if (state.currentSprint) {
    rawState.currentSprint = JSON.stringify(state.currentSprint);
  }

  for (const key in state) {
    if (key.startsWith('sprint:')) {
      rawState[key] = JSON.stringify(state[key]);
    }
  }

  PropertiesService.getDocumentProperties().setProperties(rawState, true);
}

/**
 * Idempotent BattleWriter poll function.
 * Call this as often as is appropriate for memory usage / rate limiting.
 * Will update and return the current BattleWriter state for the current document.
 * 
 * @return {Object}
 */
function poll() {
  const state = loadState();

  const now = Date.now();

  if (state.currentSprint && (now - state.currentSprint.endTime) > 1000 * 60 * 5) {
    state['sprint:' + state.currentSprint.endTime] = state.currentSprint;
    delete state.currentSprint;
  }

  // it's too memory intensive to use regex to get a real word count
  // so this is an approximation.
  // english average word len is 4.7, other langs are larger, so we just use 5
  const wordCount = DocumentApp.getActiveDocument().getBody().getText().length / 5;

  if ((now - state.lastPoll) >= 1000) {
    state.lastPoll = now;

    if (wordCount !== state.lastWordCount) {
      if (state.currentSprint) {
        state.currentSprint.endTime = now;
        state.currentSprint.endWordCount = wordCount;
      } else {
        state.currentSprint = {
          startTime: now,
          startWordCount: state.lastWordCount,
          endTime: now,
          endWordCount: wordCount,
        };
      }
      state.lastWordCount = wordCount;
    }

    storeState(state);
  }

  return state;
}





