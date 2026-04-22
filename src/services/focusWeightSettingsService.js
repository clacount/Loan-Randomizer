(function initializeFocusWeightSettingsService(globalScope) {
  const SETTINGS_STORAGE_KEY = 'loan-randomizer-settings-v1';

  function clampPercent(value, fallback = 0) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) {
      return fallback;
    }
    return Math.max(0, Math.min(100, Math.round(parsed)));
  }

  function normalizeFocusWeightPair(primaryValue, secondaryValue, defaultPrimary = 70) {
    const fallbackPrimary = clampPercent(defaultPrimary, 70);
    const normalizedPrimary = clampPercent(primaryValue, fallbackPrimary);

    if (secondaryValue === undefined || secondaryValue === null || secondaryValue === '') {
      return {
        primary: normalizedPrimary,
        secondary: 100 - normalizedPrimary
      };
    }

    const normalizedSecondary = clampPercent(secondaryValue, 100 - fallbackPrimary);
    const normalizedTotal = normalizedPrimary + normalizedSecondary;

    if (normalizedTotal === 100) {
      return {
        primary: normalizedPrimary,
        secondary: normalizedSecondary
      };
    }

    return {
      primary: normalizedPrimary,
      secondary: 100 - normalizedPrimary
    };
  }

  function getDefaultFocusWeights() {
    return {
      consumerFocused: {
        consumer: 70,
        mortgage: 30
      },
      mortgageFocused: {
        consumer: 30,
        mortgage: 70
      }
    };
  }

  function readSettings() {
    try {
      const raw = globalScope.localStorage?.getItem(SETTINGS_STORAGE_KEY);
      return raw ? JSON.parse(raw) : {};
    } catch (error) {
      return {};
    }
  }

  function writeSettings(settings) {
    try {
      globalScope.localStorage?.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
    } catch (error) {
      // no-op
    }
  }

  function normalizeFocusWeights(weights = {}) {
    const defaults = getDefaultFocusWeights();
    const consumerPair = normalizeFocusWeightPair(
      weights?.consumerFocused?.consumer,
      weights?.consumerFocused?.mortgage,
      defaults.consumerFocused.consumer
    );
    const mortgagePair = normalizeFocusWeightPair(
      weights?.mortgageFocused?.mortgage,
      weights?.mortgageFocused?.consumer,
      defaults.mortgageFocused.mortgage
    );

    return {
      consumerFocused: {
        consumer: consumerPair.primary,
        mortgage: consumerPair.secondary
      },
      mortgageFocused: {
        consumer: mortgagePair.secondary,
        mortgage: mortgagePair.primary
      }
    };
  }

  function getSavedFocusWeights() {
    const settings = readSettings();
    return normalizeFocusWeights(settings.focusWeights);
  }

  function saveFocusWeights(weights) {
    const settings = readSettings();
    settings.focusWeights = normalizeFocusWeights(weights);
    writeSettings(settings);
    return settings.focusWeights;
  }

  function getConsumerFocusedWeights() {
    return getSavedFocusWeights().consumerFocused;
  }

  function getMortgageFocusedWeights() {
    return getSavedFocusWeights().mortgageFocused;
  }

  globalScope.FocusWeightSettingsService = {
    getDefaultFocusWeights,
    getSavedFocusWeights,
    saveFocusWeights,
    normalizeFocusWeightPair,
    getConsumerFocusedWeights,
    getMortgageFocusedWeights
  };
})(window);
