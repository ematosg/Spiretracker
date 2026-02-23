import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { execSync } from 'node:child_process';

const APP_PATH = '/Users/ernestomatos/Downloads/files (9)/app.js';

function appSource() {
  return fs.readFileSync(APP_PATH, 'utf8');
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

test('core rule helpers exist', () => {
  const src = appSource();
  assert.equal(src.includes('function getRulesConfig'), true);
  assert.equal(src.includes('function setPCStressLevel'), true);
  assert.equal(src.includes('function maybeTriggerFallout'), true);
});
