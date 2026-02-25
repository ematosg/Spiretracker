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
    return tracks.reduce((total, track) => {
      const filled = (pc.stressFilled && pc.stressFilled[track]) ? pc.stressFilled[track].length : 0;
      let free = 0;
      if (Array.isArray(pc.resistances)) {
        pc.resistances.forEach((r) => {
          if (!r || !r.name) return;
          if (String(r.name).toLowerCase() !== track) return;
          free += Math.max(0, parseInt(r.value, 10) || 0);
        });
      }
      if (track === 'blood' && Array.isArray(pc.inventory)) {
        pc.inventory.forEach((item) => {
          if (!item || item.type !== 'armor') return;
          free += Math.max(0, parseInt(item.resistance, 10) || 0);
        });
      }
      return total + Math.min(10, Math.max(0, filled - free));
    }, 0);
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
