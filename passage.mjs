#!/opt/homebrew/bin/node
// Passage quotidien sur Outlook.com, sans agent ni fenêtre : ce que fait ce script ne dépend pas de ce que disent les mails, qui sont hostiles. Rien n'est envoyé, aucun lien n'est ouvert.
// Dossier indésirable : les expéditeurs connus repartent en boîte, le spam sûr va dans Éléments supprimés, le reste attend une séance interactive.
// Boîte de réception : l'original des mails coincés est déposé dans Gmail puis archivé, ce qui y est déjà part aux archives, les expéditeurs récupérés entrent sur la liste blanche.
// passage.mjs [essai]     essai : classe et affiche, sans rien supprimer, déplacer, déposer ni consigner. Une seule fenêtre à la fois sur ce profil.
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';

const ICI = path.dirname(fileURLToPath(import.meta.url)), TRAVAIL = path.join(ICI, 'travail'), EML = path.join(TRAVAIL, 'eml');
const ESSAI = process.argv.includes('essai');
const OUTLOOK = 'https://outlook.live.com/mail/0/';
const CONFIANCE_SPAM = 0.8, DUREE_MAX = 30 * 60 * 1000;
const { profil, reinjecter } = JSON.parse(fs.readFileSync(path.join(ICI, 'expediteurs.json')));

const log = (...a) => console.log(new Date().toTimeString().slice(0, 8), ...a);
const attendre = ms => new Promise(r => setTimeout(r, ms));

function script(nom, args = [], input) {
  const r = spawnSync(path.join(ICI, nom), args, { input, encoding: 'utf8', maxBuffer: 1 << 26 });
  if (r.status !== 0) throw new Error(`${nom} : ${(r.stderr || r.error || '').toString().trim().slice(-200)}`);
  return r.stdout;
}

function analyser(nom, releve, args) {
  fs.writeFileSync(path.join(TRAVAIL, nom), JSON.stringify(releve));
  return JSON.parse(script('analyser.py', [...args, nom]));
}

// Lignes de la liste -> verdict de Haiku. Sans avis, rien ne bouge : la ligne attend le passage suivant.
function juger(lignes) {
  if (!lignes.length) return () => ({});
  try {
    const avis = JSON.parse(script('classer.py', [], JSON.stringify(lignes.map((l, n) => ({ n, from: l.from, objet: l.objet })))));
    return l => avis[lignes.indexOf(l)] || {};
  } catch (e) {
    log('avis de Haiku indisponible :', e.message);
    return () => ({});
  }
}

async function ouvrir(page, dossier) {
  await page.goto(OUTLOOK + dossier);
  const pret = await page.waitForSelector('[role="treeitem"]', { timeout: 90000 }).then(() => true, () => false);
  if (!pret || !page.url().startsWith(OUTLOOK)) throw new Error('Session Outlook expirée, lancer /tri-spam-outlook');
  await attendre(3000);
}

// La liste est virtualisée : seules les lignes proches de l'écran sont dans la page.
const relever = page => page.evaluate(async () => {
  const z = triListe();
  if (z) { z.scrollTop = 0; await new Promise(r => setTimeout(r, 500)); }
  const vus = new Map();
  for (let k = 0, calme = 0; k < 900 && calme < 14; k++) {
    const avant = vus.size;
    for (const o of triLignes()) {
      const id = o.getAttribute('data-convid');
      if (id && !vus.has(id)) vus.set(id, { id, from: o.querySelector('[title*="@"]')?.getAttribute('title') || '', l: (o.getAttribute('aria-label') || '').slice(0, 300) });
    }
    calme = vus.size === avant ? calme + 1 : 0;
    if (!z) break;
    z.scrollTop += z.clientHeight * 0.9;
    await new Promise(r => setTimeout(r, 260));
  }
  if (z) z.scrollTop = 0;
  return [...vus.values()];
});

// Clic en JavaScript : un calque d'Outlook intercepte parfois les clics simulés. Jamais « Vider le dossier » : seul le nom exact est cliqué.
const cliquer = (page, nom) => page.evaluate(nom => {
  const b = [...document.querySelectorAll('button')].filter(b => b.getClientRects().length && !b.disabled && b.getAttribute('aria-disabled') !== 'true')
    .find(b => (b.getAttribute('aria-label') || b.innerText || '').trim() === nom);
  if (b) b.click();
  return !!b;
}, nom);

async function partis(page, ids, ms = 6000) {
  for (let t = 0; t < ms; t += 150) {
    const reste = await page.evaluate(ids => triLignes().some(o => ids.includes(o.getAttribute('data-convid'))), ids);
    if (!reste) return true;
    await attendre(150);
  }
  return false;
}

// Boîte de dialogue ouverte : son texte, ou null. Si elle parle de suppression définitive, on annule et on s'arrête.
async function garderDialogue(page) {
  const texte = await page.evaluate(() => {
    const d = [...document.querySelectorAll('[role="dialog"],[role="alertdialog"]')].filter(x => x.getClientRects().length).pop();
    return d ? d.innerText.slice(0, 200) : null;
  });
  if (texte && /définitivement/i.test(texte)) {
    await page.evaluate(() => [...[...document.querySelectorAll('[role="dialog"],[role="alertdialog"]')].pop().querySelectorAll('button')].find(b => /Annuler/.test(b.innerText))?.click());
    throw new Error('Outlook propose une suppression définitive : annulée, passage arrêté');
  }
  return texte;
}

// Ouvre chaque ligne voulue, une à une, et lui applique un geste. -> lignes jamais trouvées
async function chaqueLigne(page, ids, geste) {
  const restants = new Set(ids);
  for (let calme = 0; restants.size && calme < 150;) {
    const id = await page.evaluate(voulues => {
      const o = triLignes().find(o => voulues.includes(o.getAttribute('data-convid')));
      if (o) { o.click(); return o.getAttribute('data-convid'); }
      const z = triListe();
      if (z) z.scrollTop += z.clientHeight * 0.9;
      return null;
    }, [...restants]);
    if (!id) { calme++; await attendre(350); continue; }
    calme = 0;
    restants.delete(id);
    await attendre(1500);
    try {
      await geste(id);
    } catch (e) {
      if (/définitive/.test(e.message)) throw e;
      log('geste en échec :', e.message.split('\n')[0].slice(0, 120));
      await page.keyboard.press('Escape').catch(() => {});
    }
  }
  return restants;
}

// Coche les lignes voulues et clique un bouton du ruban, par lots : plus rapide qu'une à une. -> nombre de lignes parties
async function parLots(page, ids, bouton) {
  const restants = new Set(ids);
  let faits = 0;
  await page.evaluate(() => { const z = triListe(); if (z) z.scrollTop = 0; });
  for (let vides = 0; restants.size && vides < 150;) {
    const r = await page.evaluate(async voulues => {
      const lignes = triLignes().filter(o => voulues.includes(o.getAttribute('data-convid'))).slice(0, 25);
      if (!lignes.length) {
        const z = triListe();
        if (z) z.scrollTop += z.clientHeight * 0.9;
        return { coches: [] };
      }
      const basculer = o => o.querySelector('[role="checkbox"],input[type="checkbox"]')?.click();
      for (const o of lignes) {
        if (o.getAttribute('aria-selected') !== 'true') basculer(o);
        await new Promise(r => setTimeout(r, 40));
      }
      await new Promise(r => setTimeout(r, 450));
      // rien d'autre ne doit être sélectionné, un mail resté ouvert par exemple
      const intrus = triLignes().some(o => o.getAttribute('aria-selected') === 'true' && !voulues.includes(o.getAttribute('data-convid')));
      if (intrus) lignes.forEach(basculer);
      return { coches: lignes.map(o => o.getAttribute('data-convid')), intrus };
    }, [...restants]);
    if (!r.coches.length) { vides++; await attendre(350); continue; }
    if (r.intrus) { await page.keyboard.press('Escape'); await attendre(800); vides++; continue; }
    if (!await cliquer(page, bouton)) throw new Error(`bouton « ${bouton} » introuvable`);
    await attendre(500);
    await garderDialogue(page);
    if (!await partis(page, r.coches)) { vides++; continue; }
    vides = 0;
    r.coches.forEach(id => restants.delete(id));
    faits += r.coches.length;
    await attendre(1200);
  }
  return faits;
}

async function indesirable(page) {
  await ouvrir(page, 'junkemail');
  const releve = await relever(page);
  const { lignes } = analyser('indesirable.json', releve, ['--lignes']);
  const aJuger = [...lignes.connu_gmail, ...lignes.a_classer];
  const avis = juger(aJuger);
  const remettre = [...lignes.legitime, ...lignes.connu_gmail.filter(l => avis(l).verdict === 'legitime')];
  const nouveauxSpams = lignes.a_classer.filter(l => avis(l).verdict === 'spam' && avis(l).confiance >= CONFIANCE_SPAM);
  const supprimer = [...lignes.spam_connu, ...nouveauxSpams];
  const attente = aJuger.filter(l => !remettre.includes(l) && !supprimer.includes(l));
  log(`indésirable : ${releve.length} lignes, ${remettre.length} à remettre, ${supprimer.length} spams, ${attente.length} à vérifier`);
  attente.forEach(l => log('  à vérifier :', l.from, '|', l.objet.slice(0, 70), JSON.stringify(avis(l))));
  const bilan = { examines: releve.length, remis: [], supprimes: 0, attente: attente.length, nouveauxSpams };
  if (ESSAI) {
    remettre.forEach(l => log('  à remettre :', l.from, '|', l.objet.slice(0, 70)));
    supprimer.forEach(l => log('  spam :', l.from, '|', l.objet.slice(0, 70)));
    return bilan;
  }
  // Outlook les renvoie en boîte, d'où le balayage suivant les récupère dans Gmail. Sa boîte de dialogue, cochée d'office, met aussi l'expéditeur sur la liste blanche.
  const parId = new Map(remettre.map(l => [l.id, l]));
  const perdus = await chaqueLigne(page, [...parId.keys()], async id => {
    if (!await cliquer(page, 'Courrier légitime')) throw new Error(`« Courrier légitime » introuvable pour ${parId.get(id).from}`);
    await attendre(800);
    if (await garderDialogue(page)) await cliquer(page, 'OK');
    if (await partis(page, [id])) bilan.remis.push(parId.get(id).from);
    else { bilan.attente++; log('  non remis :', parId.get(id).from); }
  });
  bilan.attente += perdus.size;
  await page.keyboard.press('Escape').catch(() => {});
  bilan.supprimes = await parLots(page, supprimer.map(l => l.id), 'Supprimer');
  bilan.attente += supprimer.length - bilan.supprimes;
  log(`indésirable : ${bilan.remis.length} remis, ${bilan.supprimes} supprimés`);
  return bilan;
}

async function boite(page) {
  await ouvrir(page, 'inbox');
  const releve = await relever(page);
  const a = analyser('boite.json', releve, ['boite']);
  const avis = juger(a.recuperer);
  const aRecuperer = a.recuperer.filter(l => avis(l).verdict === 'legitime');
  const attente = a.recuperer.filter(l => !aRecuperer.includes(l));
  log(`boîte : ${releve.length} lignes, ${aRecuperer.length} à récupérer, ${a.archiver.length} à archiver, ${a.a_verifier_total + attente.length} à vérifier, ${a.recuperer_restants + a.archiver_restants} pour un prochain passage`);
  [...a.a_verifier, ...attente].forEach(l => log('  à vérifier :', l.from, '|', l.objet.slice(0, 70), JSON.stringify(avis(l))));
  const bilan = { examines: releve.length, deposes: 0, recuperes: new Set(), archives: 0, attente: a.a_verifier_total + attente.length, restants: a.recuperer_restants + a.archiver_restants, approuver: [] };
  if (ESSAI) {
    aRecuperer.forEach(l => log('  à récupérer :', l.from, '|', l.objet.slice(0, 70)));
    a.archiver.forEach(l => log('  à archiver :', l.from, '|', l.objet.slice(0, 70)));
    return bilan;
  }
  const parId = new Map(aRecuperer.map(l => [l.id, l]));
  const perdus = await chaqueLigne(page, [...parId.keys()], async id => {
    const fichier = path.join(EML, id.replace(/[^A-Za-z0-9]/g, '_').slice(-40) + '.eml');
    await page.locator('[role="main"] button[aria-label="More items"]').first().click({ timeout: 8000 });
    await page.getByRole('menu').getByRole('button', { name: 'Télécharger', exact: true }).click({ timeout: 5000 });   // pas celui d'une pièce jointe
    const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 30000 }), page.getByRole('menuitem', { name: 'Téléchargez en EML' }).click({ timeout: 5000 })]);
    await dl.saveAs(fichier);   // pas de touche Échap ici : elle fermerait le mail, et « Archiver » n'agirait plus sur rien
    const sortie = spawnSync(reinjecter, [fichier], { encoding: 'utf8' }).stdout.trim().split('\n').pop();
    const r = JSON.parse(sortie || '{"statut":"echec","motif":"pas de réponse"}');
    if (r.statut === 'refuse') { bilan.attente++; log('  refusé :', r.de, '|', (r.objet || '').slice(0, 50), r.motif); return; }
    if (r.statut === 'echec') { bilan.restants++; log('  échec du dépôt :', r.de, r.motif); return; }
    bilan.recuperes.add(parId.get(id).from);
    if (r.statut === 'depose') bilan.deposes++;
    if (await cliquer(page, 'Archiver') && await partis(page, [id])) bilan.archives++;
    else log('  non archivé :', r.de);
  });
  bilan.restants += perdus.size;
  await page.keyboard.press('Escape').catch(() => {});
  bilan.archives += await parLots(page, a.archiver.map(l => l.id), 'Archiver');
  bilan.approuver = a.approuver.filter(x => bilan.recuperes.has(x));
  log(`boîte : ${bilan.deposes} déposés, ${bilan.recuperes.size} expéditeurs récupérés, ${bilan.archives} archivés`);
  return bilan;
}

async function listeBlanche(page, adresses) {
  if (!adresses.length) return [];
  await page.goto(OUTLOOK + 'options/mail/junkEmail');
  const d = page.locator('[role="dialog"]').last();
  const ajouter = d.getByRole('button', { name: 'Ajouter un expéditeur approuvé' });
  await ajouter.waitFor({ timeout: 30000 });
  const champ = d.locator('input[placeholder*="abc123"]');
  for (const a of adresses) {
    if (!await champ.isVisible()) await ajouter.click();
    await champ.fill(a);
    await champ.press('Enter');
    await attendre(600);
  }
  const texte = await d.innerText();
  const ajoutes = adresses.filter(a => texte.includes(a));
  await d.getByRole('button', { name: 'Enregistrer' }).click();
  await attendre(2500);
  log(`liste blanche : ${ajoutes.length}/${adresses.length} ajoutées`);
  return ajoutes;
}

function echec(motif) {
  if (!ESSAI) script('consigner.py', ['--echec', motif]);
  console.log(`RÉSUMÉ : ${motif}`);
}

fs.mkdirSync(EML, { recursive: true, mode: 0o700 });
setTimeout(() => { echec('Le passage Outlook dépasse 30 minutes : arrêté'); process.exit(1); }, DUREE_MAX).unref();
let ctx, code = 0;
try {
  ctx = await chromium.launchPersistentContext(profil, { channel: 'chrome', headless: true, acceptDownloads: true, viewport: { width: 1500, height: 1000 } });
  await ctx.addInitScript(() => {
    window.triLignes = () => [...document.querySelectorAll('[role="option"]')];
    window.triListe = () => {
      for (let z = document.querySelector('[role="option"]'); z; z = z.parentElement) {
        if (/(auto|scroll)/.test(getComputedStyle(z).overflowY) && z.scrollHeight > z.clientHeight) return z;
      }
      return null;
    };
  });
  const page = ctx.pages()[0] || await ctx.newPage();
  const ind = await indesirable(page);
  const bo = await boite(page);
  const aVerifier = ind.attente + bo.attente;
  if (!ESSAI) {
    const approuves = await listeBlanche(page, bo.approuver.filter(x => !ind.remis.includes(x)));
    const familles = {};
    ind.nouveauxSpams.forEach(l => { const d = l.from.split('@').pop(); familles[d] = (familles[d] || 0) + 1; });
    const spam = Object.keys(familles).filter(d => familles[d] >= 3);
    script('consigner.py', ['--examines', ind.examines + bo.examines, '--remis', ind.remis.length, '--supprimes', ind.supprimes, '--a-verifier', aVerifier,
      '--recuperes', bo.deposes, '--archives', bo.archives, '--restants', bo.restants, '--spam', ...spam, '--legitimes', ...ind.remis, '--approuves', ...approuves].map(String));
  }
  console.log(`RÉSUMÉ : ${ESSAI ? 'essai, rien de changé. ' : ''}${ind.remis.length} remis en boîte, ${ind.supprimes} spams supprimés, ${bo.deposes} déposés dans Gmail, ${bo.archives} archivés, ${aVerifier} à vérifier.`);
} catch (e) {
  code = 1;
  echec(/ProcessSingleton|SingletonLock|already in use/.test(e.message) ? 'Navigateur déjà ouvert sur le profil Outlook : passage sauté' : e.message.split('\n')[0].slice(0, 150));
} finally {
  await ctx?.close().catch(() => {});
}
process.exit(code);
