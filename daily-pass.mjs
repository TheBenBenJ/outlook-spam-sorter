#!/opt/homebrew/bin/node
// Daily pass over Outlook.com, with no agent and no browser window. What this script does never depends on what the mails say: mail content is hostile. Nothing is sent and no link is opened.
// Junk folder: known senders go back to the inbox, sure spam goes to Deleted Items, the rest waits for an interactive session.
// Inbox: the original of each stuck mail is deposited into Gmail, then archived; what Gmail already has is archived; recovered senders are added to the safe senders list.
// daily-pass.mjs [trial]     trial = dry run: classify and print, without deleting, moving, depositing or recording anything. Only one browser at a time on this profile.
// Strings in French are either Outlook's French UI (selectors) or text that reaches the notification and the Dashboard: keep them as they are.
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';

const HERE = path.dirname(fileURLToPath(import.meta.url)), WORK = path.join(HERE, 'work'), EML = path.join(WORK, 'eml');
const DRY_RUN = process.argv.includes('trial');
// "trial" is the only argument: anything else, such as "essai" (the dry-run word before the rename), stops here rather than run a real pass.
if (process.argv.slice(2).some(a => a !== 'trial')) { console.error('usage: daily-pass.mjs [trial]'); process.exit(2); }
const OUTLOOK = 'https://outlook.live.com/mail/0/';
const SPAM_CONFIDENCE = 0.8, MAX_DURATION = 30 * 60 * 1000;
const { profil: profileDir, reinjecter: reinjectScript } = JSON.parse(fs.readFileSync(path.join(HERE, 'senders.json')));

const log = (...a) => console.log(new Date().toTimeString().slice(0, 8), ...a);
const sleep = ms => new Promise(r => setTimeout(r, ms));

function runScript(name, args = [], input) {
  const r = spawnSync(path.join(HERE, name), args, { input, encoding: 'utf8', maxBuffer: 1 << 26 });
  if (r.status !== 0) throw new Error(`${name} : ${(r.stderr || r.error || '').toString().trim().slice(-200)}`);
  return r.stdout;
}

function analyze(file, snapshot, args) {
  fs.writeFileSync(path.join(WORK, file), JSON.stringify(snapshot));
  return JSON.parse(runScript('analyze.py', [...args, file]));
}

// List rows -> Haiku's verdict. Without a verdict nothing moves: the row waits for the next run.
function judge(rows) {
  if (!rows.length) return () => ({});
  try {
    const verdicts = JSON.parse(runScript('classify.py', [], JSON.stringify(rows.map((l, n) => ({ n, from: l.from, objet: l.objet })))));
    return l => verdicts[rows.indexOf(l)] || {};
  } catch (e) {
    log('Haiku verdicts unavailable:', e.message);
    return () => ({});
  }
}

async function openFolder(page, folder) {
  await page.goto(OUTLOOK + folder);
  const ready = await page.waitForSelector('[role="treeitem"]', { timeout: 90000 }).then(() => true, () => false);
  // The Dashboard looks for "Session Outlook" in this message to show its "reconnect" tile.
  if (!ready || !page.url().startsWith(OUTLOOK)) throw new Error('Session Outlook expirée, lancer /outlook-spam-sorter');
  await sleep(3000);
}

// The list is virtualised: only the rows close to the viewport are in the page.
const readList = page => page.evaluate(async () => {
  const z = ossListScroller();
  if (z) { z.scrollTop = 0; await new Promise(r => setTimeout(r, 500)); }
  const seen = new Map();
  for (let k = 0, idle = 0; k < 900 && idle < 14; k++) {
    const before = seen.size;
    for (const o of ossListRows()) {
      const id = o.getAttribute('data-convid');
      if (id && !seen.has(id)) seen.set(id, { id, from: o.querySelector('[title*="@"]')?.getAttribute('title') || '', l: (o.getAttribute('aria-label') || '').slice(0, 300) });
    }
    idle = seen.size === before ? idle + 1 : 0;
    if (!z) break;
    z.scrollTop += z.clientHeight * 0.9;
    await new Promise(r => setTimeout(r, 260));
  }
  if (z) z.scrollTop = 0;
  return [...seen.values()];
});

// Click through JavaScript: an Outlook overlay sometimes intercepts simulated clicks. Never "Vider le dossier" (empty folder): only the exact label is clicked.
const clickButton = (page, label) => page.evaluate(label => {
  const b = [...document.querySelectorAll('button')].filter(b => b.getClientRects().length && !b.disabled && b.getAttribute('aria-disabled') !== 'true')
    .find(b => (b.getAttribute('aria-label') || b.innerText || '').trim() === label);
  if (b) b.click();
  return !!b;
}, label);

async function gone(page, ids, ms = 6000) {
  for (let t = 0; t < ms; t += 150) {
    const left = await page.evaluate(ids => ossListRows().some(o => ids.includes(o.getAttribute('data-convid'))), ids);
    if (!left) return true;
    await sleep(150);
  }
  return false;
}

// Open dialog: its text, or null. If it mentions permanent deletion, cancel it and stop the run.
async function guardDialog(page) {
  const text = await page.evaluate(() => {
    const d = [...document.querySelectorAll('[role="dialog"],[role="alertdialog"]')].filter(x => x.getClientRects().length).pop();
    return d ? d.innerText.slice(0, 200) : null;
  });
  if (text && /définitivement/i.test(text)) {
    await page.evaluate(() => [...[...document.querySelectorAll('[role="dialog"],[role="alertdialog"]')].pop().querySelectorAll('button')].find(b => /Annuler/.test(b.innerText))?.click());
    throw new Error('Outlook propose une suppression définitive : annulée, passage arrêté');
  }
  return text;
}

// Open each wanted row, one at a time, and apply an action to it. -> rows never found
async function eachRow(page, ids, action) {
  const pending = new Set(ids);
  for (let idle = 0; pending.size && idle < 150;) {
    const id = await page.evaluate(wanted => {
      const o = ossListRows().find(o => wanted.includes(o.getAttribute('data-convid')));
      if (o) { o.click(); return o.getAttribute('data-convid'); }
      const z = ossListScroller();
      if (z) z.scrollTop += z.clientHeight * 0.9;
      return null;
    }, [...pending]);
    if (!id) { idle++; await sleep(350); continue; }
    idle = 0;
    pending.delete(id);
    await sleep(1500);
    try {
      await action(id);
    } catch (e) {
      if (/définitive/.test(e.message)) throw e;
      log('action failed:', e.message.split('\n')[0].slice(0, 120));
      await page.keyboard.press('Escape').catch(() => {});
    }
  }
  return pending;
}

// Tick the wanted rows and click a ribbon button, in batches: faster than one at a time. -> number of rows gone
async function inBatches(page, ids, button) {
  const pending = new Set(ids);
  let done = 0;
  await page.evaluate(() => { const z = ossListScroller(); if (z) z.scrollTop = 0; });
  for (let empty = 0; pending.size && empty < 150;) {
    const r = await page.evaluate(async wanted => {
      const rows = ossListRows().filter(o => wanted.includes(o.getAttribute('data-convid'))).slice(0, 25);
      if (!rows.length) {
        const z = ossListScroller();
        if (z) z.scrollTop += z.clientHeight * 0.9;
        return { ticked: [] };
      }
      const toggle = o => o.querySelector('[role="checkbox"],input[type="checkbox"]')?.click();
      for (const o of rows) {
        if (o.getAttribute('aria-selected') !== 'true') toggle(o);
        await new Promise(r => setTimeout(r, 40));
      }
      await new Promise(r => setTimeout(r, 450));
      // nothing else may be selected, such as a mail left open
      const intruder = ossListRows().some(o => o.getAttribute('aria-selected') === 'true' && !wanted.includes(o.getAttribute('data-convid')));
      if (intruder) rows.forEach(toggle);
      return { ticked: rows.map(o => o.getAttribute('data-convid')), intruder };
    }, [...pending]);
    if (!r.ticked.length) { empty++; await sleep(350); continue; }
    if (r.intruder) { await page.keyboard.press('Escape'); await sleep(800); empty++; continue; }
    if (!await clickButton(page, button)) throw new Error(`bouton « ${button} » introuvable`);
    await sleep(500);
    await guardDialog(page);
    if (!await gone(page, r.ticked)) { empty++; continue; }
    empty = 0;
    r.ticked.forEach(id => pending.delete(id));
    done += r.ticked.length;
    await sleep(1200);
  }
  return done;
}

async function junkFolder(page) {
  await openFolder(page, 'junkemail');
  const snapshot = await readList(page);
  const { lignes: byClass } = analyze('junk.json', snapshot, ['--lines']);
  const toJudge = [...byClass.connu_gmail, ...byClass.a_classer];
  const verdictOf = judge(toJudge);
  const restore = [...byClass.legitime, ...byClass.connu_gmail.filter(l => verdictOf(l).verdict === 'legitime')];
  const newSpam = byClass.a_classer.filter(l => verdictOf(l).verdict === 'spam' && verdictOf(l).confiance >= SPAM_CONFIDENCE);
  const remove = [...byClass.spam_connu, ...newSpam];
  const waiting = toJudge.filter(l => !restore.includes(l) && !remove.includes(l));
  log(`junk: ${snapshot.length} rows, ${restore.length} to restore, ${remove.length} spam, ${waiting.length} to review`);
  waiting.forEach(l => log('  to review:', l.from, '|', l.objet.slice(0, 70), JSON.stringify(verdictOf(l))));
  const report = { examined: snapshot.length, restored: [], deleted: 0, waiting: waiting.length, newSpam };
  if (DRY_RUN) {
    restore.forEach(l => log('  to restore:', l.from, '|', l.objet.slice(0, 70)));
    remove.forEach(l => log('  spam:', l.from, '|', l.objet.slice(0, 70)));
    return report;
  }
  // Outlook sends them back to the inbox, where the inbox sweep then recovers them into Gmail. Its dialog, ticked by default, also adds the sender to the safe senders list.
  const byId = new Map(restore.map(l => [l.id, l]));
  const lost = await eachRow(page, [...byId.keys()], async id => {
    if (!await clickButton(page, 'Courrier légitime')) throw new Error(`"Courrier légitime" not found for ${byId.get(id).from}`);
    await sleep(800);
    if (await guardDialog(page)) await clickButton(page, 'OK');
    if (await gone(page, [id])) report.restored.push(byId.get(id).from);
    else { report.waiting++; log('  not restored:', byId.get(id).from); }
  });
  report.waiting += lost.size;
  await page.keyboard.press('Escape').catch(() => {});
  report.deleted = await inBatches(page, remove.map(l => l.id), 'Supprimer');
  report.waiting += remove.length - report.deleted;
  log(`junk: ${report.restored.length} restored, ${report.deleted} deleted`);
  return report;
}

async function inbox(page) {
  await openFolder(page, 'inbox');
  const snapshot = await readList(page);
  const analysis = analyze('inbox.json', snapshot, ['inbox']);
  const verdictOf = judge(analysis.recuperer);
  const toRecover = analysis.recuperer.filter(l => verdictOf(l).verdict === 'legitime');
  const waiting = analysis.recuperer.filter(l => !toRecover.includes(l));
  log(`inbox: ${snapshot.length} rows, ${toRecover.length} to recover, ${analysis.archiver.length} to archive, ${analysis.a_verifier_total + waiting.length} to review, ${analysis.recuperer_restants + analysis.archiver_restants} left for a later run`);
  [...analysis.a_verifier, ...waiting].forEach(l => log('  to review:', l.from, '|', l.objet.slice(0, 70), JSON.stringify(verdictOf(l))));
  const report = { examined: snapshot.length, deposited: 0, recovered: new Set(), archived: 0, waiting: analysis.a_verifier_total + waiting.length, remaining: analysis.recuperer_restants + analysis.archiver_restants, approve: [] };
  if (DRY_RUN) {
    toRecover.forEach(l => log('  to recover:', l.from, '|', l.objet.slice(0, 70)));
    analysis.archiver.forEach(l => log('  to archive:', l.from, '|', l.objet.slice(0, 70)));
    return report;
  }
  const byId = new Map(toRecover.map(l => [l.id, l]));
  const lost = await eachRow(page, [...byId.keys()], async id => {
    const file = path.join(EML, id.replace(/[^A-Za-z0-9]/g, '_').slice(-40) + '.eml');
    await page.locator('[role="main"] button[aria-label="More items"]').first().click({ timeout: 8000 });
    await page.getByRole('menu').getByRole('button', { name: 'Télécharger', exact: true }).click({ timeout: 5000 });   // not an attachment's download button
    const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 30000 }), page.getByRole('menuitem', { name: 'Téléchargez en EML' }).click({ timeout: 5000 })]);
    await dl.saveAs(file);   // no Escape key here: it would close the mail, and "Archiver" would then act on nothing
    // Reinjection contract: the script prints one JSON line per file; statut is depose, deja, refuse or echec.
    const lastLine = spawnSync(reinjectScript, [file], { encoding: 'utf8' }).stdout.trim().split('\n').pop();
    const r = JSON.parse(lastLine || '{"statut":"echec","motif":"no response"}');
    if (r.statut === 'refuse') { report.waiting++; log('  refused:', r.de, '|', (r.objet || '').slice(0, 50), r.motif); return; }
    if (r.statut === 'echec') { report.remaining++; log('  deposit failed:', r.de, r.motif); return; }
    report.recovered.add(byId.get(id).from);
    if (r.statut === 'depose') report.deposited++;
    if (await clickButton(page, 'Archiver') && await gone(page, [id])) report.archived++;
    else log('  not archived:', r.de);
  });
  report.remaining += lost.size;
  await page.keyboard.press('Escape').catch(() => {});
  report.archived += await inBatches(page, analysis.archiver.map(l => l.id), 'Archiver');
  report.approve = analysis.approuver.filter(x => report.recovered.has(x));
  log(`inbox: ${report.deposited} deposited, ${report.recovered.size} senders recovered, ${report.archived} archived`);
  return report;
}

async function safeSenders(page, addresses) {
  if (!addresses.length) return [];
  await page.goto(OUTLOOK + 'options/mail/junkEmail');
  const d = page.locator('[role="dialog"]').last();
  const addButton = d.getByRole('button', { name: 'Ajouter un expéditeur approuvé' });
  await addButton.waitFor({ timeout: 30000 });
  const field = d.locator('input[placeholder*="abc123"]');
  for (const a of addresses) {
    if (!await field.isVisible()) await addButton.click();
    await field.fill(a);
    await field.press('Enter');
    await sleep(600);
  }
  const text = await d.innerText();
  const added = addresses.filter(a => text.includes(a));
  await d.getByRole('button', { name: 'Enregistrer' }).click();
  await sleep(2500);
  log(`safe senders: ${added.length}/${addresses.length} added`);
  return added;
}

// The reason goes to outlook-state.json, the notification and the Dashboard: it stays in French.
function fail(reason) {
  if (!DRY_RUN) runScript('record.py', ['--failure', reason]);
  console.log(`SUMMARY: ${reason}`);
}

fs.mkdirSync(EML, { recursive: true, mode: 0o700 });
setTimeout(() => { fail('Le passage Outlook dépasse 30 minutes : arrêté'); process.exit(1); }, MAX_DURATION).unref();
let ctx, code = 0;
try {
  ctx = await chromium.launchPersistentContext(profileDir, { channel: 'chrome', headless: true, acceptDownloads: true, viewport: { width: 1500, height: 1000 } });
  // Page helpers, prefixed so they cannot clash with Outlook's own globals.
  await ctx.addInitScript(() => {
    window.ossListRows = () => [...document.querySelectorAll('[role="option"]')];
    window.ossListScroller = () => {
      for (let z = document.querySelector('[role="option"]'); z; z = z.parentElement) {
        if (/(auto|scroll)/.test(getComputedStyle(z).overflowY) && z.scrollHeight > z.clientHeight) return z;
      }
      return null;
    };
  });
  const page = ctx.pages()[0] || await ctx.newPage();
  const junk = await junkFolder(page);
  const box = await inbox(page);
  const toReview = junk.waiting + box.waiting;
  if (!DRY_RUN) {
    const approved = await safeSenders(page, box.approve.filter(x => !junk.restored.includes(x)));
    const perDomain = {};
    junk.newSpam.forEach(l => { const d = l.from.split('@').pop(); perDomain[d] = (perDomain[d] || 0) + 1; });
    const spamDomains = Object.keys(perDomain).filter(d => perDomain[d] >= 3);
    runScript('record.py', ['--examined', junk.examined + box.examined, '--restored', junk.restored.length, '--deleted', junk.deleted, '--to-review', toReview,
      '--recovered', box.deposited, '--archived', box.archived, '--remaining', box.remaining, '--spam', ...spamDomains, '--legitimate', ...junk.restored, '--approved', ...approved].map(String));
  }
  // The summary is the notification text: French on purpose.
  console.log(`SUMMARY: ${DRY_RUN ? 'essai, rien de changé. ' : ''}${junk.restored.length} remis en boîte, ${junk.deleted} spams supprimés, ${box.deposited} déposés dans Gmail, ${box.archived} archivés, ${toReview} à vérifier.`);
} catch (e) {
  code = 1;
  fail(/ProcessSingleton|SingletonLock|already in use/.test(e.message) ? 'Navigateur déjà ouvert sur le profil Outlook : passage sauté' : e.message.split('\n')[0].slice(0, 150));
} finally {
  await ctx?.close().catch(() => {});
}
process.exit(code);
