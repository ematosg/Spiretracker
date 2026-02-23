import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { execSync } from 'node:child_process';
import { createRequire } from 'node:module';

const APP_PATH = '/Users/ernestomatos/Downloads/files (9)/app.js';
const RULES_PATH = '/Users/ernestomatos/Downloads/files (9)/rules-engine.js';
const require = createRequire(import.meta.url);

function appSource() {
  return fs.readFileSync(APP_PATH, 'utf8');
}

function extractFunction(source, name) {
  const sig = `function ${name}(`;
  const start = source.indexOf(sig);
  assert.notEqual(start, -1, `Missing function: ${name}`);
  let parenDepth = 0;
  let seenOpen = false;
  let bodyStart = -1;
  for (let i = start + sig.length - 1; i < source.length; i++) {
    const ch = source[i];
    if (ch === '(') {
      parenDepth++;
      seenOpen = true;
    } else if (ch === ')') {
      parenDepth--;
      if (seenOpen && parenDepth === 0) {
        bodyStart = source.indexOf('{', i);
        break;
      }
    }
  }
  assert.notEqual(bodyStart, -1, `Missing body start for: ${name}`);
  let depth = 0;
  for (let i = bodyStart; i < source.length; i++) {
    const ch = source[i];
    if (ch === '{') depth++;
    if (ch === '}') {
      depth--;
      if (depth === 0) return source.slice(start, i + 1);
    }
  }
  throw new Error(`Unclosed function body for: ${name}`);
}

function loadFns(names, seed = {}) {
  const src = appSource();
  const ctx = Object.assign({}, seed);
  const code = names.map(n => extractFunction(src, n)).join('\n\n') +
    `\n\nglobalThis.__fns = { ${names.join(', ')} };`;
  vm.runInNewContext(code, ctx);
  return ctx.__fns;
}

function plain(obj) {
  return JSON.parse(JSON.stringify(obj));
}

test('app.js parses with Node syntax check', () => {
  execSync(`node --check "${APP_PATH}"`, { stdio: 'pipe' });
});

test('native blocking dialogs are not used', () => {
  const src = appSource();
  assert.equal(/\bconfirm\s*\(/.test(src), false, 'confirm() should not be used');
  assert.equal(/\bprompt\s*\(/.test(src), false, 'prompt() should not be used');
  assert.equal(/\balert\s*\(/.test(src), false, 'alert() should not be used');
});

test('rules engine module exports expected helpers', () => {
  const engine = require(RULES_PATH);
  assert.equal(typeof engine.getRulesConfig, 'function');
  assert.equal(typeof engine.totalStressForFallout, 'function');
  assert.equal(typeof engine.falloutSeverityForTotalStress, 'function');
  assert.equal(typeof engine.stressClearAmountForSeverity, 'function');
});

test('rules profile mapping behaves as expected', () => {
  const engine = require(RULES_PATH);
  const { getRulesConfig } = loadFns(['getRulesConfig'], { RulesEngine: engine });

  const core = plain(getRulesConfig({ rulesProfile: 'Core' }));
  assert.deepEqual(core, {
    difficultyDowngrades: true,
    falloutCheckOnStress: true,
    clearStressOnFallout: true
  });

  const quick = plain(getRulesConfig({ rulesProfile: 'Quickstart' }));
  assert.deepEqual(quick, {
    difficultyDowngrades: true,
    falloutCheckOnStress: false,
    clearStressOnFallout: false
  });

  const custom = plain(getRulesConfig({
    rulesProfile: 'Custom',
    customRules: { falloutCheckOnStress: false }
  }));
  assert.deepEqual(custom, {
    difficultyDowngrades: true,
    falloutCheckOnStress: false,
    clearStressOnFallout: true
  });
});

test('fallout helper thresholds are correct', () => {
  const engine = require(RULES_PATH);
  const { falloutSeverityForTotalStress, stressClearAmountForSeverity } = loadFns([
    'falloutSeverityForTotalStress',
    'stressClearAmountForSeverity'
  ], { RulesEngine: engine });

  assert.equal(falloutSeverityForTotalStress(0), 'Minor');
  assert.equal(falloutSeverityForTotalStress(4), 'Minor');
  assert.equal(falloutSeverityForTotalStress(5), 'Moderate');
  assert.equal(falloutSeverityForTotalStress(8), 'Moderate');
  assert.equal(falloutSeverityForTotalStress(9), 'Severe');

  assert.equal(stressClearAmountForSeverity('Minor'), 3);
  assert.equal(stressClearAmountForSeverity('Moderate'), 5);
  assert.equal(stressClearAmountForSeverity('Severe'), 7);
});

test('total stress for fallout clamps each track at 10', () => {
  const engine = require(RULES_PATH);
  const { totalStressForFallout } = loadFns(['totalStressForFallout'], { RulesEngine: engine });
  const pc = {
    stressFilled: {
      blood: Array.from({ length: 13 }, (_, i) => i),
      mind: Array.from({ length: 9 }, (_, i) => i),
      silver: [],
      shadow: Array.from({ length: 22 }, (_, i) => i),
      reputation: Array.from({ length: 1 }, (_, i) => i)
    }
  };
  assert.equal(totalStressForFallout(pc), 10 + 9 + 0 + 10 + 1);
});

test('maybeTriggerFallout uses injected rng and creates fallout when roll is below total', () => {
  const engine = require(RULES_PATH);
  const logs = [];
  const toasts = [];
  const camp = { rulesProfile: 'Core', customRules: {} };
  let seq = 0;

  const { maybeTriggerFallout } = loadFns(
    [
      'getRulesConfig',
      'totalStressForFallout',
      'falloutSeverityForTotalStress',
      'stressClearAmountForSeverity',
      'clearStressForFallout',
      'maybeTriggerFallout'
    ],
    {
      RulesEngine: engine,
      currentCampaign: () => camp,
      appendLog: (msg) => logs.push(msg),
      showToast: (msg) => toasts.push(msg),
      generateId: (pfx = 'id') => `${pfx}-${++seq}`
    }
  );

  const pc = {
    id: 'pc-1',
    stressFilled: {
      blood: Array.from({ length: 6 }, (_, i) => i),
      mind: Array.from({ length: 1 }, (_, i) => i),
      silver: [],
      shadow: [],
      reputation: []
    },
    fallout: []
  };

  maybeTriggerFallout(pc, 'blood', 1, { rng: () => 0.1 });
  assert.equal(pc.fallout.length, 1);
  assert.equal(pc.fallout[0].severity, 'Moderate');
  assert.equal(logs.length > 0, true);
  assert.equal(toasts.length > 0, true);
});

test('maybeTriggerFallout does nothing when roll meets or exceeds total', () => {
  const engine = require(RULES_PATH);
  const camp = { rulesProfile: 'Core', customRules: {} };

  const { maybeTriggerFallout } = loadFns(
    [
      'getRulesConfig',
      'totalStressForFallout',
      'falloutSeverityForTotalStress',
      'stressClearAmountForSeverity',
      'clearStressForFallout',
      'maybeTriggerFallout'
    ],
    {
      RulesEngine: engine,
      currentCampaign: () => camp,
      appendLog: () => {},
      showToast: () => {},
      generateId: (pfx = 'id') => `${pfx}-x`
    }
  );

  const pc = {
    id: 'pc-2',
    stressFilled: {
      blood: Array.from({ length: 4 }, (_, i) => i),
      mind: [],
      silver: [],
      shadow: [],
      reputation: []
    },
    fallout: []
  };

  maybeTriggerFallout(pc, 'blood', 1, { rng: () => 0.9 });
  assert.equal(pc.fallout.length, 0);
});
