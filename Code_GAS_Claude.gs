// ══════════════════════════════════════════════════════════════════════
// AB RENOV 35 — Google Apps Script v3 — Claude (Anthropic)
// Déployer : Web App · Exécuter en tant que : Moi · Accès : Tout le monde
// Propriétés du script : ANTHROPIC_API_KEY = sk-ant-...
// ══════════════════════════════════════════════════════════════════════

const SHEET_LIBELLES = 'libelles';
const SHEET_PRIX     = 'prix_metier';
const SHEET_DEVIS    = 'devis';
const SHEET_LIGNES   = 'lignes_devis';
const SHEET_IMPORTS  = 'imports_historique';
const CLAUDE_MODEL   = 'claude-sonnet-4-5';

// ══ ROUTING ══════════════════════════════════════════════════════════

function doGet(e) {
  let result;
  try {
    if (e.parameter.payload) {
      result = routePost(JSON.parse(decodeURIComponent(e.parameter.payload)));
    } else {
      const action = e.parameter.action || 'list';
      switch (action) {
        case 'list':       result = listLibelles(); break;
        case 'listprix':   result = listPrix();     break;
        case 'listdevis':  result = listDevis();    break;
        case 'listimports': result = listImports();   break;
        case 'getdevis':   result = getDevis(e.parameter.id); break;
        case 'listlignes': result = listLignes(e.parameter.devis_id); break;
        default:           result = { error: 'Action inconnue : ' + action };
      }
    }
  } catch (err) { result = { error: err.message }; }
  return buildResponse(result);
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData ? e.postData.contents : '{}'); }
  catch (_) { return buildResponse({ error: 'JSON invalide' }); }
  let result;
  try { result = routePost(body); }
  catch (err) { result = { error: err.message }; }
  return buildResponse(result);
}

function routePost(body) {
  switch (body.action) {
    case 'generate':          return generate(body);
    case 'save':              return saveLibelle(body);
    case 'list':              return listLibelles();
    case 'update':            return updateLibelle(body);
    case 'delete':            return deleteLibelle(body);
    case 'saveprix':          return savePrix(body);
    case 'updateprix':        return updatePrix(body);
    case 'deleteprix':        return deletePrix(body);
    case 'savedevis':         return saveDevis(body);
    case 'updatedevis':       return updateDevis(body);
    case 'deletedevis':       return deleteDevis(body);
    case 'finalizedevis':     return finalizeDevis(body);
    case 'dupliquedevis':     return dupliqueDevis(body);
    case 'structurer':        return structurerTexte(body);
    case 'analyser_fichier':  return analyserFichier(body);
    case 'optimiserdevis':    return optimiserDevis(body);
    case 'analyser_document': return analyserDocument(body);
    case 'valider_import':    return validerImport(body);
    case 'save_import':       return saveImport(body);
    case 'update_import_statut': return updateImportStatut(body);
    case 'delete_import':     return deleteImport(body);
    default: return { error: 'Action inconnue : ' + body.action };
  }
}

function buildResponse(data) {
  const json = JSON.stringify(data);
  const out = ContentService.createTextOutput(json);
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

// CORS preflight — required for POST requests with Content-Type: application/json
// Access-Control-Allow-Origin: * is added automatically by Google Apps Script
// infrastructure for deployed web apps (Execute as: Me, Access: Anyone).
function doOptions(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══ CLAUDE — APPEL PRINCIPAL ═════════════════════════════════════════

function callClaude(system, user, maxTokens) {
  const apiKey = getApiKey();
  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: maxTokens || 1024,
    system: system,
    messages: [{ role: 'user', content: user }]
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    const err = JSON.parse(res.getContentText());
    throw new Error('Claude ' + res.getResponseCode() + ': ' + (err.error?.message || res.getContentText()));
  }
  const data = JSON.parse(res.getContentText());
  return data.content?.[0]?.text || '{}';
}

// ══ CLAUDE — APPEL AVEC FICHIER (PDF ou IMAGE) ══════════════════════

function callClaudeWithFile(system, userText, b64, mimeType, maxTokens) {
  const apiKey = getApiKey();
  // Construire le contenu utilisateur avec le fichier
  const userContent = [];
  if (mimeType === 'application/pdf') {
    userContent.push({
      type: 'document',
      source: { type: 'base64', media_type: 'application/pdf', data: b64 }
    });
  } else {
    // Image (jpg, png, webp)
    const imgType = mimeType || 'image/jpeg';
    userContent.push({
      type: 'image',
      source: { type: 'base64', media_type: imgType, data: b64 }
    });
  }
  userContent.push({ type: 'text', text: userText });

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: maxTokens || 4096,
    system: system,
    messages: [{ role: 'user', content: userContent }]
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    const err = JSON.parse(res.getContentText());
    throw new Error('Claude ' + res.getResponseCode() + ': ' + (err.error?.message || res.getContentText()));
  }
  const data = JSON.parse(res.getContentText());
  return data.content?.[0]?.text || '{}';
}

// ══ GÉNÉRATION LIBELLÉ ═══════════════════════════════════════════════

function generate(body) {
  const prestation = (body.prestation || '').trim();
  const style = body.style || 'Standard';
  if (!prestation) return { error: 'Prestation manquante.' };

  const exemples  = chercherExemplesProches(prestation, 3);
  const prixMetier = chercherPrixMetier(prestation);

  const styleInstructions = {
    'Commercial': 'Libellé valorisant, bénéfice client visible, donne envie de signer. 2-3 phrases.',
    'Standard':   '1-2 phrases claires et directes.',
    'Technique':  '2-3 phrases avec terminologie DTU, normes, finitions.',
    'Court':      '1 phrase concise, max 15 mots.'
  };

  let sys = `Tu es conducteur de travaux et économiste du bâtiment pour AB RENOV 35 (Rennes).
Rédige des libellés de devis professionnels.
STYLE "${style}" : ${styleInstructions[style] || styleInstructions['Standard']}
Règles : Prix HT, marché rennais Grand Ouest 2025, MO seule sauf indication.
Réponds UNIQUEMENT en JSON valide sans markdown :
{"libelle":"...","prix_bas":null,"prix_haut":null,"unite":"m²","prix_valide":false}`;

  if (exemples.length > 0) {
    sys += '\n\nEXEMPLES AB RENOV :';
    exemples.forEach((ex, i) => { sys += `\n[${i+1}] "${ex.prestation}" → "${ex.libelle}"`; });
  }
  if (prixMetier) {
    sys += `\n\nPRIX RENNES : ${prixMetier.prix_bas}–${prixMetier.prix_haut}€/${prixMetier.unite} (prix_valide:true)`;
  }

  const raw = callClaude(sys, `Libellé style "${style}" pour : "${prestation}"`, 512);
  let data;
  try { data = JSON.parse(raw.replace(/```json|```/g, '').trim()); }
  catch (_) { return { error: 'IA non parsable', raw }; }

  let confidence = 'faible';
  if (prixMetier && exemples.length >= 2) confidence = 'eleve';
  else if (prixMetier || exemples.length >= 1) confidence = 'moyen';

  return {
    libelle:     data.libelle || '',
    prix_bas:    data.prix_bas  || (prixMetier ? prixMetier.prix_bas  : null),
    prix_haut:   data.prix_haut || (prixMetier ? prixMetier.prix_haut : null),
    unite:       data.unite     || (prixMetier ? prixMetier.unite     : 'm²'),
    prix_valide: !!prixMetier,
    confidence
  };
}

// ══ STRUCTURATION TEXTE → DEVIS ══════════════════════════════════════

function structurerTexte(body) {
  const texte = (body.texte || '').trim();
  if (!texte) return { error: 'Texte manquant.' };

  const prixDispos = getSheetData(SHEET_PRIX).slice(0, 10)
    .map(r => ({ p: r[1], bas: r[2], haut: r[3], u: r[4] }));

  const sys = `Tu es métreur et conducteur de travaux pour AB RENOV 35 (Rennes).
Analyse ce texte et génère un devis structuré complet.
Pour chaque prestation retourne un objet JSON avec :
lot (Électricité|Plomberie|Sanitaire|Fourniture|Pose|Maçonnerie|Peinture|Carrelage|Menuiserie|Isolation|Divers),
designation (libellé professionnel), quantite (nombre ou null), unite (m²|ml|m³|unité|forfait|heure|jour|point|équipement),
pu_ht (nombre ou null), confiance_prix (faible|moyen|eleve).
Si la description est vague, déduis les lots habituels.
Prix HT marché rennais 2025.
Retourne UNIQUEMENT le JSON : {"lignes":[...],"hypotheses":"..."}
PRIX REF : ${JSON.stringify(prixDispos)}`;

  const raw = callClaude(sys, `Texte : "${texte}"`, 2048);
  try {
    const clean = raw.replace(/```json|```/g, '').trim();
    const parsed = JSON.parse(clean);
    if (Array.isArray(parsed)) return { lignes: parsed };
    return { lignes: parsed.lignes || [], hypotheses: parsed.hypotheses || '' };
  } catch (_) { return { error: 'Structuration échouée.', raw }; }
}

// ══ ANALYSE FICHIER (PDF/IMAGE) → DEVIS ═════════════════════════════

function analyserFichier(body) {
  const filename = body.filename || 'document';
  if (!body.data && !body.texte) return { error: 'Données manquantes.' };

  const sys = `Tu es conducteur de travaux pour AB RENOV 35.
Extrais toutes les prestations de ce document.
Retourne UNIQUEMENT le JSON valide sans markdown :
[{"lot":"Électricité|Plomberie|Sanitaire|Fourniture|Pose|Maçonnerie|Peinture|Carrelage|Menuiserie|Isolation|Divers","designation":"...","quantite":null,"unite":"m²","pu_ht":null}]`;

  let raw;
  try {
    if (body.texte && body.texte.length > 50) {
      raw = callClaude(sys, 'Extrais les prestations du document : ' + filename + '\n\n' + body.texte.substring(0, 25000), 4096);
    } else {
      raw = callClaudeWithFile(sys, 'Extrais toutes les prestations du fichier : ' + filename, body.data, body.type || 'image/jpeg', 4096);
    }
  } catch(e) { return { error: e.message, lignes: [] }; }

  try {
    const lignes = JSON.parse(raw.replace(/```json|```/g, '').trim());
    return { lignes: Array.isArray(lignes) ? lignes : [] };
  } catch (_) { return { error: 'Analyse échouée.', raw }; }
}

// ══ OPTIMISER DEVIS ══════════════════════════════════════════════════

function optimiserDevis(body) {
  const lignes = body.lignes || [];
  if (!lignes.length) return { suggestions: [] };

  const prixRef = getSheetData(SHEET_PRIX).slice(0, 10)
    .map(r => ({ p: r[1], bas: r[2], haut: r[3], u: r[4] }));

  const sys = `Tu es économiste du bâtiment (marché rennais 2025).
Analyse ce devis et retourne un tableau JSON de suggestions d'amélioration.
Pour chaque problème : {"ligne_index":N,"message":"explication","nouveau_libelle":"optionnel","prix":null}
Vérifie : libellés vagues, quantités nulles, prix aberrants.
RETOURNE UNIQUEMENT le JSON valide sans markdown.
PRIX REF : ${JSON.stringify(prixRef)}`;

  const input = JSON.stringify(lignes.map((l, i) => ({
    i, lot: l.lot, des: (l.des || l.designation || '').substring(0, 80), qte: l.qte, pu: l.pu
  })));

  const raw = callClaude(sys, `Devis : ${input}`, 1024);
  try {
    const s = JSON.parse(raw.replace(/```json|```/g, '').trim());
    return { suggestions: Array.isArray(s) ? s.slice(0, 8) : [] };
  } catch (_) { return { suggestions: [] }; }
}

// ══ ANALYSE DOCUMENT BASE MÉTIER ═════════════════════════════════════

function analyserDocument(body) {
  const filename = body.filename || 'document';
  const docType = body.doc_type || 'divers', prixType = body.prix_type, marge = parseFloat(body.marge) || 0;

  const margeTxt = (prixType === 'pro' && marge > 0) ? 'Prix PRO : multiplier par ' + (1 + marge/100) + ' pour prix vente.' : '';
  const sys = 'Tu es expert en analyse de documents batiment pour AB RENOV 35.\n'
    + 'Extrais TOUTES les lignes de prix sans exception.\n'
    + 'Retourne UNIQUEMENT ce JSON valide sans markdown ni explication :\n'
    + '{"fournisseur":"nom","nb_lignes":0,"nb_nouveaux":0,"nb_mis_a_jour":0,"nb_doublons":0,"nb_incertains":0,"confiance":85,'
    + '"lignes":[{"designation":"libelle complet","unite":"unite","pu_ht":0,"reference":""}]}\n'
    + 'Type document : ' + docType + '. ' + margeTxt;

  let raw;
  try {
    if (body.texte && body.texte.length > 50) {
      // Texte brut extrait du PDF — méthode légère et fiable
      raw = callClaude(sys, 'Analyse ce document ' + docType + ' (' + filename + ') :\n\n' + body.texte.substring(0, 25000), 4096);
    } else if (body.data) {
      // Fichier base64 — images et petits PDF
      raw = callClaudeWithFile(sys, 'Analyse ce document ' + docType + ' : ' + filename, body.data, body.type || 'image/jpeg', 4096);
    } else {
      return { error: 'Données manquantes.' };
    }
  } catch(e) {
    Logger.log('Erreur Claude: ' + e.message);
    return { error: e.message, nb_lignes: 0, confiance: 0, lignes: [] };
  }

  try {
    const clean = raw.replace(/```json/g,'').replace(/```/g,'').trim();
    const parsed = JSON.parse(clean);
    return {
      fournisseur:   parsed.fournisseur   || 'Inconnu',
      nb_lignes:     parsed.nb_lignes     || (parsed.lignes||[]).length,
      nb_nouveaux:   parsed.nb_nouveaux   || 0,
      nb_mis_a_jour: parsed.nb_mis_a_jour || 0,
      nb_doublons:   parsed.nb_doublons   || 0,
      nb_incertains: parsed.nb_incertains || 0,
      confiance:     parsed.confiance     || 50,
      lignes:        parsed.lignes        || []
    };
  } catch(e) {
    Logger.log('Parse error: ' + e.message + ' raw: ' + raw.substring(0,500));
    return { error: 'Analyse echouee : ' + e.message, nb_lignes: 0, confiance: 0, lignes: [] };
  }
}


// ══ HISTORIQUE IMPORTS ════════════════════════════════════════════════

function listImports() {
  return { imports: getSheetData(SHEET_IMPORTS).map(r => ({
    id: r[0], date: r[1], fichier: r[2], type: r[3],
    fournisseur: r[4], lignes: r[5], conf: r[6], statut: r[7]
  })).filter(i => i.id).reverse() };
}

function saveImport(body) {
  const imp = body.imp || {};
  if (!imp.id) return { error: 'ID manquant.' };
  const sheet = getOrCreateSheet(SHEET_IMPORTS, ['id','date','fichier','type','fournisseur','lignes','conf','statut']);
  const impId = 'IMP_' + Date.now();
  sheet.appendRow([impId, imp.date||'', imp.fichier||'', imp.type||'', imp.fournisseur||'', imp.lignes||0, imp.conf||0, imp.statut||'à vérifier']);
  // Mettre à jour l'ID dans la réponse pour que le client puisse le référencer
  imp.id = impId;
  return { success: true, id: impId };
}

function updateImportStatut(body) {
  return updateSheetRow(SHEET_IMPORTS, body.id, (s, r) => {
    s.getRange(r, 8).setValue(body.statut || 'à vérifier');
  });
}

function deleteImport(body) {
  return deleteSheetRow(SHEET_IMPORTS, body.id);
}

// ══ VALIDER IMPORT ════════════════════════════════════════════════════

function validerImport(body) {
  const data = body.data || {}, lignes = data.lignes || [];
  let integres = 0;

  const sheetPrix = getOrCreateSheet(SHEET_PRIX, ['id','prestation','prix_bas','prix_haut','unite']);
  const sheetLib  = getOrCreateSheet(SHEET_LIBELLES, ['id','date','prestation','libelle','prix_bas','prix_haut','unite','prix_abrenov']);
  const existingPrix = getSheetData(SHEET_PRIX);
  const existingLib  = getSheetData(SHEET_LIBELLES);

  lignes.forEach(l => {
    if (!l.designation || !l.pu_ht) return;
    const norm = normaliser(l.designation);

    // 1. Sauvegarder dans prix_metier
    const foundPrix = existingPrix.findIndex(r => normaliser(String(r[1]||'')) === norm);
    if (foundPrix < 0) {
      sheetPrix.appendRow(['PRIX_'+Date.now()+'_'+integres, l.designation, l.pu_ht*0.9, l.pu_ht*1.1, l.unite||'m²']);
    }

    // 2. Sauvegarder dans bibliothèque libelles
    const foundLib = existingLib.findIndex(r => normaliser(String(r[2]||'')) === norm);
    if (foundLib < 0) {
      // Générer un libellé pro avec Claude
      let libelle = l.designation;
      try {
        const sys = 'Tu es conducteur de travaux pour AB RENOV 35. '
          + 'Reformule ce libellé de façon professionnelle et commerciale en 1-2 phrases maximum. '
          + 'Réponds UNIQUEMENT avec le libellé reformulé, sans guillemets ni explication.';
        libelle = callClaude(sys, l.designation, 200) || l.designation;
        libelle = libelle.trim().replace(/^"|"$/g,'');
      } catch(_) { libelle = l.designation; }

      sheetLib.appendRow([
        'LIB_'+Date.now()+'_'+integres,
        new Date().toISOString(),
        l.designation,
        libelle,
        l.pu_ht * 0.9,
        l.pu_ht * 1.1,
        l.unite || 'm²',
        l.pu_ht
      ]);
    }

    integres++;
    Utilities.sleep(300); // Éviter le rate limit Claude
  });

  return { success: true, nb_integres: integres };
}

// ══ CRUD LIBELLÉS ════════════════════════════════════════════════════

function listLibelles() {
  return { libelles: getSheetData(SHEET_LIBELLES).map(r => ({
    id: r[0], date: r[1], prestation: r[2], libelle: r[3],
    prix_bas: r[4]||null, prix_haut: r[5]||null, unite: r[6]||'m²', prix_abrenov: r[7]||null
  })).filter(l => l.id && l.libelle) };
}
function saveLibelle(body) {
  const sheet = getOrCreateSheet(SHEET_LIBELLES, ['id','date','prestation','libelle','prix_bas','prix_haut','unite','prix_abrenov']);
  const libelle = (body.libelle||'').trim();
  if (!libelle) return { error: 'Libellé manquant.' };
  if (getSheetData(SHEET_LIBELLES).some(r => normaliser(String(r[3])) === normaliser(libelle))) return { doublon: true };
  const id = 'LIB_' + Date.now();
  sheet.appendRow([id, new Date().toISOString(), body.prestation||'', libelle, body.prix_bas||'', body.prix_haut||'', body.unite||'m²', body.prix_abrenov||'']);
  return { success: true, id };
}
function updateLibelle(body) {
  return updateSheetRow(SHEET_LIBELLES, body.id, (s, r) => {
    if (body.prestation  !== undefined) s.getRange(r,3).setValue(body.prestation);
    if (body.libelle     !== undefined) s.getRange(r,4).setValue(body.libelle);
    if (body.prix_bas    !== undefined) s.getRange(r,5).setValue(body.prix_bas||'');
    if (body.prix_haut   !== undefined) s.getRange(r,6).setValue(body.prix_haut||'');
    if (body.unite       !== undefined) s.getRange(r,7).setValue(body.unite||'m²');
    if (body.prix_abrenov!== undefined) s.getRange(r,8).setValue(body.prix_abrenov||'');
  });
}
function deleteLibelle(body) { return deleteSheetRow(SHEET_LIBELLES, body.id); }

// ══ CRUD PRIX MÉTIER ═════════════════════════════════════════════════

function listPrix() {
  return { prix: getSheetData(SHEET_PRIX).map(r => ({
    id: r[0], prestation: r[1], prix_bas: r[2], prix_haut: r[3], unite: r[4]
  })).filter(p => p.prestation) };
}
function savePrix(body) {
  const sheet = getOrCreateSheet(SHEET_PRIX, ['id','prestation','prix_bas','prix_haut','unite']);
  const id = 'PRIX_' + Date.now();
  sheet.appendRow([id, body.prestation||'', body.prix_bas||'', body.prix_haut||'', body.unite||'m²']);
  return { success: true, id };
}
function updatePrix(body) {
  return updateSheetRow(SHEET_PRIX, body.id, (s, r) => {
    if (body.prestation !== undefined) s.getRange(r,2).setValue(body.prestation);
    if (body.prix_bas   !== undefined) s.getRange(r,3).setValue(body.prix_bas||'');
    if (body.prix_haut  !== undefined) s.getRange(r,4).setValue(body.prix_haut||'');
    if (body.unite      !== undefined) s.getRange(r,5).setValue(body.unite||'m²');
  });
}
function deletePrix(body) { return deleteSheetRow(SHEET_PRIX, body.id); }

// ══ CRUD DEVIS ═══════════════════════════════════════════════════════

function listDevis() {
  return { devis: getSheetData(SHEET_DEVIS).map(r => ({
    id: r[0], numero: r[1], date: r[2], client: r[3], chantier: r[4],
    statut: r[5], tva_taux: r[6], total_ht: r[7], total_tva: r[8], total_ttc: r[9], note: r[10]
  })).filter(d => d.id) };
}
function getDevis(id) {
  if (!id) return { error: 'ID manquant.' };
  const row = getSheetData(SHEET_DEVIS).find(r => String(r[0]) === String(id));
  if (!row) return { error: 'Devis introuvable.' };
  return { devis: { id: row[0], numero: row[1], date: row[2], client: row[3], chantier: row[4], statut: row[5], tva_taux: row[6], total_ht: row[7], total_tva: row[8], total_ttc: row[9], note: row[10] }, lignes: listLignes(id).lignes };
}
function saveDevis(body) {
  const sheetD = getOrCreateSheet(SHEET_DEVIS,  ['id','numero','date','client','chantier','statut','tva_taux','total_ht','total_tva','total_ttc','note']);
  const sheetL = getOrCreateSheet(SHEET_LIGNES, ['id','devis_id','ordre','lot','designation','quantite','unite','pu_ht','total_ht','tva_taux','source_libelle','libelle_id']);
  const lignes = body.lignes || [], tva = body.tva_taux || 10;
  const totalHT  = lignes.reduce((a,l) => a + (l.type==='sep'?0:(l.qte!=null&&l.pu!=null?l.qte*l.pu:0)), 0);
  const totalTVA = Math.round(totalHT*tva)/100, totalTTC = Math.round((totalHT+totalTVA)*100)/100;
  let id = body.id;
  if (id) {
    updateSheetRow(SHEET_DEVIS, id, (s,r) => {
      s.getRange(r,3).setValue(body.date||''); s.getRange(r,4).setValue(body.client||'');
      s.getRange(r,5).setValue(body.chantier||''); s.getRange(r,6).setValue(body.statut||'brouillon');
      s.getRange(r,7).setValue(tva); s.getRange(r,8).setValue(Math.round(totalHT*100)/100);
      s.getRange(r,9).setValue(totalTVA); s.getRange(r,10).setValue(totalTTC); s.getRange(r,11).setValue(body.description||'');
    });
    const al = sheetL.getDataRange().getValues();
    for (let i = al.length-1; i >= 1; i--) { if (String(al[i][1])===String(id)) sheetL.deleteRow(i+1); }
  } else {
    id = 'DEV_' + Date.now();
    sheetD.appendRow([id, body.numero||genNumeroServer(), body.date||new Date().toISOString().split('T')[0], body.client||'', body.chantier||'', body.statut||'brouillon', tva, Math.round(totalHT*100)/100, totalTVA, totalTTC, body.description||'']);
  }
  lignes.forEach((l, ordre) => {
    const lid = l.id || ('LIG_'+Date.now()+'_'+ordre);
    if (l.type==='sep') { sheetL.appendRow([lid,id,ordre,l.lot||'LOT','SEPARATEUR','','','','','','','']); }
    else { sheetL.appendRow([lid,id,ordre,l.lot||'',l.des||l.designation||'',l.qte||'',l.unit||l.unite||'m²',l.pu||'',l.qte&&l.pu?Math.round(l.qte*l.pu*100)/100:'',tva,l.src||l.source||'',l.libelle_id||'']); }
  });
  return { success: true, id, total_ht: Math.round(totalHT*100)/100, total_ttc: totalTTC };
}
function listLignes(devisId) {
  if (!devisId) return { lignes: [] };
  return { lignes: getSheetData(SHEET_LIGNES).filter(r => String(r[1])===String(devisId)).sort((a,b) => (a[2]||0)-(b[2]||0)).map(r => ({
    id: r[0], devis_id: r[1], ordre: r[2], lot: r[3],
    designation: r[4]==='SEPARATEUR'?undefined:r[4], type: r[4]==='SEPARATEUR'?'sep':undefined,
    quantite: r[5]||null, unite: r[6]||'m²', pu_ht: r[7]||null, source_libelle: r[10]||''
  })) };
}
function updateDevis(body) {
  return updateSheetRow(SHEET_DEVIS, body.id, (s,r) => {
    if (body.client   !== undefined) s.getRange(r,4).setValue(body.client);
    if (body.chantier !== undefined) s.getRange(r,5).setValue(body.chantier);
    if (body.statut   !== undefined) s.getRange(r,6).setValue(body.statut);
  });
}
function deleteDevis(body) {
  const id = body.id; if (!id) return { error: 'ID manquant.' };
  const sl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LIGNES);
  if (sl) { const d = sl.getDataRange().getValues(); for (let i = d.length-1; i >= 1; i--) { if (String(d[i][1])===String(id)) sl.deleteRow(i+1); } }
  return deleteSheetRow(SHEET_DEVIS, id);
}
function finalizeDevis(body) {
  const id = body.id; if (!id) return { error: 'ID manquant.' };
  updateSheetRow(SHEET_DEVIS, id, (s,r) => s.getRange(r,6).setValue('finalise'));
  const lignes = (body.lignes||[]).filter(l => l.type!=='sep' && l.pu && l.qte);
  let maj = 0;
  const sheet = getOrCreateSheet(SHEET_PRIX, ['id','prestation','prix_bas','prix_haut','unite']);
  const prixRows = getSheetData(SHEET_PRIX);
  lignes.forEach(l => {
    const des = l.des||l.designation||'', norm = normaliser(des).substring(0,30);
    if (!norm) return;
    const pu = parseFloat(l.pu);
    const found = prixRows.findIndex(r => normaliser(String(r[1]||'')).substring(0,30)===norm);
    if (found >= 0) { sheet.getRange(found+2,3).setValue(Math.min(parseFloat(prixRows[found][2])||pu,pu)); sheet.getRange(found+2,4).setValue(Math.max(parseFloat(prixRows[found][3])||pu,pu)); }
    else { sheet.appendRow(['PRIX_'+Date.now()+'_'+maj, des, pu*0.85, pu*1.15, l.unit||l.unite||'m²']); }
    maj++;
  });
  return { success: true, prix_maj: maj };
}
function dupliqueDevis(body) {
  const src = getDevis(body.id); if (src.error) return src;
  return saveDevis({ ...src.devis, id: null, numero: genNumeroServer(), client: '', statut: 'brouillon', lignes: src.lignes.map(l => ({ ...l, id: 'LIG_'+Date.now()+'_'+Math.random() })) });
}

// ══ RECHERCHE ════════════════════════════════════════════════════════

function chercherExemplesProches(prestation, max) {
  try {
    const rows = getSheetData(SHEET_LIBELLES); if (!rows.length) return [];
    const mots = normaliser(prestation).split(' ').filter(m => m.length > 3);
    return rows.map(r => {
      let s = 0;
      mots.forEach(m => { if (normaliser(String(r[2]||'')).includes(m)) s+=2; if (normaliser(String(r[3]||'')).includes(m)) s+=1; });
      return { score: s, prestation: r[2], libelle: r[3] };
    }).filter(x => x.score > 0 && x.libelle).sort((a,b) => b.score-a.score).slice(0,max);
  } catch (_) { return []; }
}

function chercherPrixMetier(prestation) {
  try {
    const rows = getSheetData(SHEET_PRIX); if (!rows.length) return null;
    const mots = normaliser(prestation).split(' ').filter(m => m.length > 3);
    let best = null, score = 0;
    rows.forEach(r => { let s = 0; mots.forEach(m => { if (normaliser(String(r[1]||'')).includes(m)) s++; }); if (s > score) { score = s; best = { prestation: r[1], prix_bas: r[2], prix_haut: r[3], unite: r[4] }; } });
    return score > 0 ? best : null;
  } catch (_) { return null; }
}

// ══ SHEETS UTILS ═════════════════════════════════════════════════════

function getSheetData(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) return [];
  const d = s.getDataRange().getValues(); if (d.length < 2) return [];
  const f = String(d[0][0]).toLowerCase();
  const h = isNaN(f) && !f.startsWith('lib_') && !f.startsWith('prix_') && !f.startsWith('dev_') && !f.startsWith('lig_') && !f.startsWith('imp_');
  return d.slice(h?1:0).filter(r => r.some(c => c!==''&&c!==null));
}
function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(headers); s.getRange(1,1,1,headers.length).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function updateSheetRow(name, id, fn) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) return { error: 'Feuille introuvable : ' + name };
  const d = s.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) { if (String(d[i][0])===String(id)) { fn(s, i+1); return { success: true }; } }
  return { error: 'Introuvable : ' + id };
}
function deleteSheetRow(name, id) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) return { error: 'Feuille introuvable.' };
  const d = s.getDataRange().getValues();
  for (let i = d.length-1; i >= 1; i--) { if (String(d[i][0])===String(id)) { s.deleteRow(i+1); return { success: true }; } }
  return { error: 'Introuvable : ' + id };
}
function normaliser(str) {
  return (str||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9 ]/g,' ').replace(/\s+/g,' ').trim();
}
function getApiKey() {
  const k = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!k) throw new Error('ANTHROPIC_API_KEY manquante dans les propriétés du script.');
  return k;
}
function genNumeroServer() {
  const d = new Date();
  return 'DEVIS-' + d.getFullYear() + String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0') + '-' + String(Math.floor(Math.random()*900)+100);
}

// ══ INIT (exécuter une fois manuellement) ════════════════════════════

function initPrixMetier() {
  const sheet = getOrCreateSheet(SHEET_PRIX, ['id','prestation','prix_bas','prix_haut','unite']);
  if (getSheetData(SHEET_PRIX).length > 2) { Logger.log('Déjà initialisé.'); return; }
  [['PRIX_001','Pose carrelage sol 60×60',25,45,'m²'],['PRIX_002','Pose carrelage mural',30,55,'m²'],['PRIX_003','Enduit de façade',15,28,'m²'],['PRIX_004','Ravalement de façade',18,35,'m²'],['PRIX_005','Pose parquet flottant',18,32,'m²'],['PRIX_006','Peinture intérieure murs plafonds',8,20,'m²'],['PRIX_007','Peinture façade extérieure',15,30,'m²'],['PRIX_008','Installation WC suspendu',280,450,'unité'],['PRIX_009','Remplacement robinetterie lavabo',80,160,'unité'],['PRIX_010',"Pose douche à l'italienne",800,1500,'forfait'],['PRIX_011','Mise aux normes tableau électrique',600,1200,'forfait'],['PRIX_012','Pose prise ou interrupteur',35,75,'unité'],['PRIX_013','Isolation combles soufflage',20,40,'m²'],['PRIX_014','Isolation murs intérieurs doublage',25,50,'m²'],['PRIX_015','Isolation plancher bas',15,35,'m²'],['PRIX_016','Pose fenêtre PVC double vitrage',150,300,'unité'],['PRIX_017','Pose porte intérieure',120,250,'unité'],['PRIX_018','Pose parquet massif',25,50,'m²'],['PRIX_019','Cloison placo standard',28,55,'m²'],['PRIX_020','Faux plafond BA13',18,38,'m²'],['PRIX_021','Pose meuble vasque salle de bain',150,280,'unité'],['PRIX_022','Pose baignoire',200,400,'unité'],['PRIX_023','Plomberie alimentation évacuation',600,1200,'forfait'],['PRIX_024','Pose luminaire',40,90,'unité'],['PRIX_025','Création point électrique',55,110,'point'],['PRIX_026','Pose radiateur électrique',80,180,'unité'],['PRIX_027','Pose porte fenêtre PVC',200,380,'unité'],['PRIX_028','Démolition cloison',15,35,'m²'],['PRIX_029','Pose revêtement sol souple vinyl',8,18,'m²'],['PRIX_030','Traitement humidité injection',20,45,'ml']]
    .forEach(r => sheet.appendRow(r));
  getOrCreateSheet(SHEET_LIBELLES, ['id','date','prestation','libelle','prix_bas','prix_haut','unite','prix_abrenov']);
  getOrCreateSheet(SHEET_DEVIS,    ['id','numero','date','client','chantier','statut','tva_taux','total_ht','total_tva','total_ttc','note']);
  getOrCreateSheet(SHEET_LIGNES,   ['id','devis_id','ordre','lot','designation','quantite','unite','pu_ht','total_ht','tva_taux','source_libelle','libelle_id']);
  Logger.log('Init OK — 30 prix + 4 onglets créés.');
}
