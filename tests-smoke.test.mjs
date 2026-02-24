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

test('getRecentEntityActions returns newest-first actions for one entity', () => {
  const camp = {
    logs: [
      { action: 'Edited notes', target: 'pc-1', type: 'action', time: '2026-02-20T10:00:00.000Z' },
      { action: 'Added task', target: 'pc-2', type: 'action', time: '2026-02-20T10:05:00.000Z' },
      { action: 'Added fallout', target: 'pc-1', type: 'action', time: '2026-02-20T10:10:00.000Z' },
      { action: 'Session 2 start', target: '', type: 'session', time: '2026-02-20T10:20:00.000Z' },
      { action: 'Edited inventory', target: 'pc-1', type: 'action', time: '2026-02-20T10:30:00.000Z' }
    ]
  };
  const { getRecentEntityActions } = loadFns(['getRecentEntityActions'], {
    currentCampaign: () => camp
  });

  const actions = plain(getRecentEntityActions('pc-1', 2));
  assert.equal(actions.length, 2);
  assert.equal(actions[0].action, 'Edited inventory');
  assert.equal(actions[1].action, 'Added fallout');
});

test('deriveAutomationSuggestions produces relationship and clock suggestions from session signals', () => {
  const camp = {
    currentSession: 3,
    clocks: [],
    logs: [
      { session: 3, type: 'action', action: 'Added fallout' },
      { session: 3, type: 'action', action: 'Fallout triggered (Moderate)' },
      { session: 3, type: 'action', action: 'Edited relationship type' },
      { session: 3, type: 'action', action: 'Edited relationship direction' },
      { session: 3, type: 'action', action: 'Added relationship' },
      { session: 3, type: 'action', action: 'Added member' },
      { session: 3, type: 'action', action: 'Added task' },
      { session: 3, type: 'action', action: 'Edited task' },
      { session: 3, type: 'action', action: 'Added task' },
      { session: 2, type: 'action', action: 'Edited relationship type' }
    ]
  };
  const { deriveAutomationSuggestions } = loadFns(['deriveAutomationSuggestions']);
  const out = plain(deriveAutomationSuggestions(camp, {
    pendingFalloutCount: 4,
    highStressPcCount: 2,
    openTasksCount: 7
  }));
  const keys = out.map((s) => s.key);
  assert.equal(keys.includes('clock-fallout-aftermath'), true);
  assert.equal(keys.includes('clock-faction-backlash'), true);
  assert.equal(keys.includes('review-web'), true);
  assert.equal(keys.includes('clock-operation-pressure'), true);
});

test('normalizeScenarioPack keeps valid extensions and rejects empty packs', () => {
  const { normalizeScenarioPack } = loadFns(['normalizeScenarioPack'], {
    generateId: () => 'pack-generated'
  });
  const empty = normalizeScenarioPack({ id: 'x', name: 'Empty' });
  assert.equal(empty, null);

  const pack = plain(normalizeScenarioPack({
    id: 'city-pack',
    name: 'City Pack',
    scenePrompts: { complications: ['A ward is sealed.'] },
    npcTemplates: { fixer2: { label: 'Fixer 2', role: 'Fixer' } }
  }));
  assert.equal(pack.id, 'city-pack');
  assert.equal(pack.scenePrompts.complications.length, 1);
  assert.equal(Object.keys(pack.npcTemplates).includes('fixer2'), true);
});

test('getEffectiveScenePromptPools merges active pack prompts with defaults', () => {
  const { getActiveScenarioPacks, getEffectiveScenePromptPools } = loadFns(
    ['getActiveScenarioPacks', 'getEffectiveScenePromptPools'],
    {
      currentCampaign: () => ({
        scenarioPacks: [
          { enabled: true, scenePrompts: { complications: ['A ward is sealed.'], factionReactions: [], twists: [] } },
          { enabled: false, scenePrompts: { complications: ['Disabled entry'], factionReactions: [], twists: [] } }
        ]
      }),
      SCENE_COMPLICATIONS: ['Base complication'],
      SCENE_FACTION_REACTIONS: ['Base reaction'],
      SCENE_TWISTS: ['Base twist']
    }
  );
  const pools = plain(getEffectiveScenePromptPools());
  assert.equal(pools.complications.includes('Base complication'), true);
  assert.equal(pools.complications.includes('A ward is sealed.'), true);
  assert.equal(pools.complications.includes('Disabled entry'), false);
  assert.equal(pools.factionReactions.includes('Base reaction'), true);
  assert.equal(pools.twists.includes('Base twist'), true);
  // Ensure dependency was loaded and callable
  assert.equal(Array.isArray(getActiveScenarioPacks()), true);
});

test('exportCampaign strips gm-only entities, secrets, and non-party messages for player export', () => {
  const camp = {
    entities: {
      a: { id: 'a', name: 'Alpha', gmOnly: false, gmNotes: 'secret' },
      b: { id: 'b', name: 'Beta', gmOnly: true, gmNotes: 'hidden' }
    },
    relationships: {
      r1: { id: 'r1', source: 'a', target: 'b', secret: false, type: 'Ally' },
      r2: { id: 'r2', source: 'a', target: 'a', secret: true, type: 'Enemy' },
      r3: { id: 'r3', source: 'a', target: 'a', secret: false, type: 'Ally' }
    },
    positions: { a: { x: 1, y: 2 }, b: { x: 2, y: 3 } },
    messages: [
      { id: 'm1', target: 'party', text: 'hello' },
      { id: 'm2', target: 'gm', text: 'secret whisper' }
    ],
    logs: [
      { action: 'Visible edit', target: 'a' },
      { action: 'Hidden entity edit', target: 'b' },
      { action: 'Secret relation edit', target: 'r2' }
    ]
  };
  const { exportCampaign } = loadFns(['exportCampaign'], {
    currentCampaign: () => camp
  });
  const playerSafe = JSON.parse(exportCampaign(false));
  assert.equal(!!playerSafe.entities.a, true);
  assert.equal(!!playerSafe.entities.b, false);
  assert.equal(playerSafe.entities.a.gmNotes, undefined);
  assert.equal(!!playerSafe.relationships.r1, false);
  assert.equal(!!playerSafe.relationships.r2, false);
  assert.equal(!!playerSafe.relationships.r3, true);
  assert.equal(playerSafe.messages.length, 1);
  assert.equal(playerSafe.messages[0].target, 'party');
  assert.equal(playerSafe.logs.length, 1);
  assert.equal(playerSafe.logs[0].target, 'a');
});

test('deriveGmWhisperTargets prioritizes non-GM campaign members', () => {
  const { deriveGmWhisperTargets } = loadFns(['deriveGmWhisperTargets']);
  const camp = {
    gmUsers: ['alice', 'gm2'],
    memberUsers: ['alice', 'gm2', 'bob', 'cara']
  };
  const users = {
    alice: { createdAt: 'x' },
    gm2: { createdAt: 'x' },
    bob: { createdAt: 'x' },
    cara: { createdAt: 'x' },
    zed: { createdAt: 'x' }
  };
  const out = plain(deriveGmWhisperTargets(camp, users, 'alice'));
  assert.deepEqual(out, ['bob', 'cara']);
});
