import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('manifest.xml declares Word, Excel, and PowerPoint hosts without Word-only requirements', () => {
  const xml = fs.readFileSync(path.join(ROOT, 'manifest.xml'), 'utf8');
  assert.match(xml, /<Host Name="Document"\s*\/>/);
  assert.match(xml, /<Host Name="Workbook"\s*\/>/);
  assert.match(xml, /<Host Name="Presentation"\s*\/>/);
  assert.doesNotMatch(xml, /<Set Name="WordApi"/, 'WordApi requirement in base manifest blocks Excel and PowerPoint');
  assert.match(xml, /https:\/\/localhost:3443\/taskpane\.html/);
  assert.match(xml, /<Permissions>ReadWriteDocument<\/Permissions>/);
  assert.match(xml, /<Id>32c92c95-b390-486a-a782-02b7e30e8545<\/Id>/);
});

test('install.sh and uninstall.sh pass bash syntax check and cover Word, Excel, and PowerPoint wef paths', () => {
  const installPath = path.join(ROOT, 'install.sh');
  const uninstallPath = path.join(ROOT, 'uninstall.sh');

  execFileSync('bash', ['-n', installPath]);
  execFileSync('bash', ['-n', uninstallPath]);

  const installSrc = fs.readFileSync(installPath, 'utf8');
  const uninstallSrc = fs.readFileSync(uninstallPath, 'utf8');

  for (const container of ['com.microsoft.Word', 'com.microsoft.Excel', 'com.microsoft.Powerpoint']) {
    assert.ok(installSrc.includes(container), `install.sh missing ${container}`);
    assert.ok(uninstallSrc.includes(container), `uninstall.sh missing ${container}`);
  }
});
