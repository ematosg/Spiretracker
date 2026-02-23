(() => {
  function getRulesConfig(camp) {
    const defaults = {
      difficultyDowngrades: true,
      falloutCheckOnStress: true,
      clearStressOnFallout: true
    };
    if (!camp) return defaults;
    const profile = camp.rulesProfile || 'Core';
    if (profile === 'Quickstart') {
      return {
        difficultyDowngrades: true,
        falloutCheckOnStress: false,
        clearStressOnFallout: false
      };
    }
    if (profile === 'Custom') {
      return Object.assign({}, defaults, camp.customRules || {});
    }
    return defaults;
  }

  function totalStressForFallout(pc) {
    const tracks = ['blood', 'mind', 'silver', 'shadow', 'reputation'];
    let total = 0;
    tracks.forEach(track => {
      const filled = (pc.stressFilled && pc.stressFilled[track]) ? pc.stressFilled[track].length : 0;
      total += Math.min(10, filled);
    });
    return total;
  }

  function falloutSeverityForTotalStress(total) {
    if (total >= 9) return 'Severe';
    if (total >= 5) return 'Moderate';
    return 'Minor';
  }

  function stressClearAmountForSeverity(severity) {
    if (severity === 'Severe') return 7;
    if (severity === 'Moderate') return 5;
    return 3;
  }

  const engine = {
    getRulesConfig,
    totalStressForFallout,
    falloutSeverityForTotalStress,
    stressClearAmountForSeverity
  };

  if (typeof window !== 'undefined') {
    window.SpireRulesEngine = engine;
  }
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = engine;
  }
})();
