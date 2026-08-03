const { requireString, safeErrorMessage } = require('./errors.cjs');

function getAvailableLanguages(session) {
  return Array.isArray(session?.availableSpellCheckerLanguages)
    ? session.availableSpellCheckerLanguages.filter((item) => typeof item === 'string')
    : [];
}

function toSpellCheckState(settings, session) {
  const availableLanguages = getAvailableLanguages(session);
  const language = settings.spellCheckLanguage || availableLanguages[0] || 'en-US';
  return {
    enabled: Boolean(settings.spellCheckEnabled),
    language,
    availableLanguages,
    dictionaryStatus: availableLanguages.includes(language) ? 'ready' : 'missing',
  };
}

function applySpellCheckState(session, state, traceId, log) {
  if (!session) return;
  try {
    session.setSpellCheckerEnabled(Boolean(state.enabled));
    if (state.dictionaryStatus === 'ready' && state.language) {
      session.setSpellCheckerLanguages([state.language]);
    }
  } catch (error) {
    log('error', 'apply spell check state failed', traceId, {
      error: safeErrorMessage(error),
    });
  }
}

function createSpellService(options) {
  const { getEventWindow, log, settings } = options;

  async function apply(event, nextSettings, traceId) {
    const session = getEventWindow(event)?.webContents.session;
    const state = toSpellCheckState(nextSettings, session);
    applySpellCheckState(session, state, traceId, log);
    return state;
  }

  async function getState(event, _payload, traceId) {
    const result = await apply(event, await settings.readSettings(), traceId);
    log('info', 'spell getState', traceId, result);
    return result;
  }

  async function setEnabled(event, payload, traceId) {
    const result = await settings.set(event, {
      key: 'spellCheckEnabled',
      value: payload?.enabled,
    }, traceId);
    const state = await apply(event, result.settings, traceId);
    log('info', 'spell setEnabled', traceId, state);
    return state;
  }

  async function setLanguage(event, payload, traceId) {
    const language = requireString(payload?.language, 'language', 64);
    const result = await settings.set(event, {
      key: 'spellCheckLanguage',
      value: language,
    }, traceId);
    const state = await apply(event, result.settings, traceId);
    log('info', 'spell setLanguage', traceId, state);
    return state;
  }

  return {
    applyInitial: (session, currentSettings, traceId) => {
      const state = toSpellCheckState(currentSettings, session);
      applySpellCheckState(session, state, traceId, log);
    },
    getState,
    setEnabled,
    setLanguage,
  };
}

module.exports = {
  applySpellCheckState,
  createSpellService,
  toSpellCheckState,
};
