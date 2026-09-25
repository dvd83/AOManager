import React, { useState, useEffect } from "react";
import { supabase } from "./supabaseClient.js";
import * as XLSX from "xlsx";
import logoWordmark from "./assets/logo-wordmark.svg";
import logoWordmarkDark from "./assets/logo-wordmark-dark.svg";
import {
  LayoutDashboard, FilePlus2, FileText, ListChecks, Scale, Users,
  ClipboardCheck, BarChart3, FileBarChart2, History, ChevronRight,
  AlertTriangle, CheckCircle2, Sparkles, ArrowLeft, Info,
  ShieldCheck, Building2, Search, Bell, Circle, Download, Upload, Plus, Trash2, PenLine, Menu, X, Lock
} from "lucide-react";

/* =========================================================================
   AO MANAGER — prototype fonctionnel
   Cycle de vie complet : cadrage -> documents (rédigés en ligne) ->
   exigences -> critères -> fournisseurs -> réception -> évaluation
   pondérée -> comparaison -> synthèse -> export (Word/Excel réels).
   Données 100% fictives. Pas de backend : l'état vit en mémoire.
   ========================================================================= */

const C = {
  ink: "#182234", inkSoft: "#4B5872", bg: "#FAF9F7", surface: "#FFFFFF",
  border: "#E7E4DD", borderSoft: "#F0EEE8",
  accent: "#FD5312", accentDark: "#C2400D", accentSoft: "#FFEDE4",
  amber: "#A9720F", amberSoft: "#F6EDD9",
  green: "#3C7A54", greenSoft: "#E4EFE7",
  red: "#AE4A41", redSoft: "#F5E7E4", slateSoft: "#EEEDE7",
};

const STATUS_META = {
  draft:       { label: "Brouillon",       color: C.inkSoft,  bg: C.slateSoft },
  preparation: { label: "En préparation",  color: C.amber,    bg: C.amberSoft },
  validation:  { label: "En validation",   color: C.accentDark, bg: C.accentSoft },
  ongoing:     { label: "En cours",        color: C.accent,   bg: C.accentSoft },
  evaluation:  { label: "En évaluation",   color: "#8A5A2E",  bg: "#F3E7D8" },
  completed:   { label: "Terminé",         color: C.green,    bg: C.greenSoft },
};

const CRITICALITY_META = {
  Bloquante: C.red, Critique: "#B4633A", Majeure: C.amber, Normale: C.inkSoft, Souhaitable: "#7B8A9A",
};

const TENDER_TYPES = ["IT", "Logiciel", "Infrastructure", "Réseau", "Télécom", "Cybersécurité", "Prestations", "Matériel", "Maintenance", "Services", "Autre"];

const TABS = [
  { key: "info", label: "Informations", icon: Info, phase: 1 },
  { key: "need", label: "Besoin & périmètre", icon: ListChecks, phase: 1 },
  { key: "requirements", label: "Exigences", icon: ClipboardCheck, phase: 1 },
  { key: "criteria", label: "Critères", icon: Scale, phase: 1 },
  { key: "documents", label: "Documents", icon: FileText, phase: 1 },
  { key: "suppliers", label: "Fournisseurs", icon: Users, phase: 2 },
  { key: "evaluation", label: "Évaluation", icon: BarChart3, phase: 2 },
  { key: "comparison", label: "Comparaison", icon: BarChart3, phase: 2 },
  { key: "synthesis", label: "Synthèse", icon: FileBarChart2, phase: 2 },
  { key: "history", label: "Historique", icon: History, phase: 0 },
];

// Seuils des marchés publics — Accord intercantonal sur les marchés publics (AIMP 2019).
// Valeurs en CHF HT, cycle en vigueur jusqu'au 31.12.2027 (inchangées depuis 2024/2025).
// Sources : AIMP 2019 texte consolidé (DTAP/bpuk.ch), circulaire des seuils 2026-2027, rsGE L 6.05.
// À vérifier avant publication auprès des Achats / du service juridique : les seuils évoluent par cycle de 2 ans.
// Application dédiée aux achats IT : uniquement Fournitures (matériel, licences) et Services (prestations, logiciels, maintenance).
// Seuils tels que communiqués par l'utilisateur (pratique retenue en interne) : au-delà du seuil de gré à gré,
// un AO devient nécessaire ; sur invitation jusqu'à 250k ; procédure ouverte nationale jusqu'à 650k ; au-delà, international.
// À reconfirmer avec les Achats en cas de doute, ces seuils pouvant évoluer.
const SEUILS_MP_GE = {
  "Fournitures": { gre_a_gre: 100000, invitation: 250000, amp: 650000 },
  "Services":    { gre_a_gre: 150000, invitation: 250000, amp: 650000 },
};

function suggestProcedure(nature, budget) {
  const s = SEUILS_MP_GE[nature];
  if (!s || !budget) return null;
  const b = Number(budget);
  if (b < s.gre_a_gre) return { label: "Procédure de gré à gré", detail: `sous ${s.gre_a_gre.toLocaleString("fr-CH")} CHF` };
  if (b < s.invitation) return { label: "Procédure sur invitation", detail: `entre ${s.gre_a_gre.toLocaleString("fr-CH")} et ${s.invitation.toLocaleString("fr-CH")} CHF` };
  if (s.amp && b >= s.amp) return { label: "Procédure ouverte internationale (soumise aux accords internationaux)", detail: `dès ${s.amp.toLocaleString("fr-CH")} CHF` };
  return { label: "Procédure ouverte nationale (AO intérieur)", detail: `entre ${s.invitation.toLocaleString("fr-CH")} et ${s.amp.toLocaleString("fr-CH")} CHF` };
}

// Délais minimaux de remise des offres selon la procédure (jours calendaires depuis la publication/le lancement).
// Sources : AIMP 2019 art. 46 al. 2 (internationale, 40j) et al. 4 (nationale, 20j) ; pratique cantonale genevoise
// pour la procédure sur invitation (25j). Le gré à gré n'a pas de délai minimal légal.
function minSubmissionDays(procedureLabel) {
  const l = (procedureLabel || "").toLowerCase();
  if (l.includes("gré à gré")) return null;
  if (l.includes("invitation")) return { days: 25, source: "pratique cantonale genevoise pour la procédure sur invitation" };
  if (l.includes("internationale")) return { days: 40, source: "AIMP 2019, art. 46 al. 2 (marchés soumis aux accords internationaux)" };
  if (l.includes("ouverte")) return { days: 20, source: "AIMP 2019, art. 46 al. 4 (marchés non soumis aux accords internationaux)" };
  return null;
}
function daysBetween(d1, d2) {
  if (!d1 || !d2) return null;
  const ms = new Date(d2) - new Date(d1);
  return Math.round(ms / (1000 * 60 * 60 * 24));
}
// Retourne les avertissements de calendrier (délai de remise des offres, marge avant décision) pour un AO donné.
function addDaysISO(dateStr, days) {
  const d = new Date(dateStr);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}
function scheduleWarnings(tender, procedureLabel) {
  const warnings = [];
  const minDeadline = minSubmissionDays(procedureLabel);
  if (minDeadline && tender.dateLaunch && tender.dateClose) {
    const actual = daysBetween(tender.dateLaunch, tender.dateClose);
    if (actual != null && actual < minDeadline.days) {
      warnings.push({
        level: "warn",
        text: `Le délai entre lancement et clôture n'est que de ${actual} jour(s) — le minimum usuel pour cette procédure est de ${minDeadline.days} jours (${minDeadline.source}).`,
        field: "dateClose", value: addDaysISO(tender.dateLaunch, minDeadline.days),
        applyLabel: `Reporter la clôture à ${minDeadline.days} jours après le lancement`,
      });
    }
  }
  if (tender.dateClose && tender.dateDecision) {
    const gap = daysBetween(tender.dateClose, tender.dateDecision);
    if (gap != null && gap < 14) {
      warnings.push({
        level: "info",
        text: `Seulement ${gap} jour(s) entre la clôture et la décision — l'analyse des offres prend généralement 2 à 4 semaines selon la complexité (recommandation pratique, pas une obligation légale).`,
        field: "dateDecision", value: addDaysISO(tender.dateClose, 21),
        applyLabel: "Reporter la décision à 3 semaines après la clôture",
      });
    }
  }
  return warnings;
}

// Explications en langage simple, pour des personnes qui découvrent les marchés publics.
const PROCEDURE_HELP = {
  "gré à gré": "Vous choisissez vous-même un fournisseur et lui attribuez directement l'AO, sans mise en concurrence formelle.",
  "sur invitation": "Vous invitez vous-même au moins 3 fournisseurs de votre choix à faire une offre. Pas de publication publique.",
  "ouverte nationale": "L'AO est publié publiquement (sur SIMAP) : tout fournisseur suisse ou établi en Suisse intéressé peut déposer une offre.",
  "ouverte internationale": "L'AO est publié publiquement (sur SIMAP), ouvert aussi aux fournisseurs des pays parties aux accords internationaux sur les marchés publics.",
};
function procedureHelpText(label) {
  const l = (label || "").toLowerCase();
  if (l.includes("gré à gré")) return PROCEDURE_HELP["gré à gré"];
  if (l.includes("invitation")) return PROCEDURE_HELP["sur invitation"];
  if (l.includes("internationale")) return PROCEDURE_HELP["ouverte internationale"];
  if (l.includes("ouverte")) return PROCEDURE_HELP["ouverte nationale"];
  return null;
}

/* --------------------------- MODÈLES DE DOCUMENTS (rédaction guidée) -------------------------- */

const TABLE_DOC_NAMES = ["Série de prix"];
const AUTO_TABLE_DOC_NAMES = ["Cahier de réponses"];
const QUESTIONS_DOC_NAMES = ["Annexe A - Compréhension des besoins"];
// Formulaire K2 du Guide romand pour l'invitation à soumissionner et l'adjudication de marchés
// publics (procédure ouverte) — cf. www.vd.ch/etat-droit-finances/marches-publics/guide-romand.
const K2_DOC_NAME = "K2 - Dossier d'appel d'offres (procédure ouverte)";

// Placeholders obligatoires — jamais d'invention : toute information non fournie explicitement doit rester visible comme telle.
const A_COMPLETER = "[À COMPLÉTER PAR ACHATS]";
const A_VALIDER = "[À VALIDER PAR ACHATS]";

const TEXT_SCHEMAS = {
  // Structure calquée sur la logique réelle d'un service Achats : document opérationnel de consultation,
  // distinct du CDC. Aucune information (procédure, dates, critères, montants) n'est déduite ou inventée :
  // ce qui n'est pas explicitement saisi reste affiché comme [À COMPLÉTER PAR ACHATS] / [À VALIDER PAR ACHATS].
  "Procédure AO": (t) => [
    { key: "objet", label: "1. Objet de la procédure", seed: `${t.object || A_COMPLETER}\nCette procédure se réfère au Cahier des charges de l'AO ${t.reference}, qui détaille le besoin et les exigences.` },
    { key: "type_procedure", label: "3. Type de procédure", seed: `Procédure retenue : ${t.procedure || A_VALIDER}\nJustification / base réglementaire : ${A_VALIDER}` },
    { key: "perimetre", label: "4. Périmètre de la consultation", seed: `${(t.scope.included || []).length ? "Renvoi au Cahier des charges pour le détail du périmètre (inclus : " + t.scope.included.join(", ") + ")." : A_COMPLETER}\nLe soumissionnaire doit répondre à l'ensemble des exigences du cahier des charges.` },
    { key: "modalites", label: "7. Modalités de remise des offres", seed: A_VALIDER },
    { key: "recevabilite", label: "9. Critères de recevabilité (administrative / technique / financière / éliminatoire)", seed: A_COMPLETER },
    { key: "qr", label: "11. Questions / réponses", seed: `Point de contact : ${t.buyer || A_COMPLETER}\nCanal, délai et modalités de publication des réponses : ${A_VALIDER}` },
    { key: "validite", label: "12. Validité des offres", seed: A_COMPLETER },
    { key: "confidentialite", label: "13. Confidentialité et communication", seed: A_VALIDER },
    { key: "conditions_part", label: "14. Conditions particulières (lots, options, contrat-cadre, plafonds)", seed: `${A_COMPLETER} si applicable` },
  ],
  // Structure calquée sur le template officiel Achats — chapitres 1, 3, 6 (le reste : "Réponse du soumissionnaire", exigences, planification, glossaire, est généré automatiquement à l'export).
  "Cahier des charges": (t) => [
    { key: "obj_doc", label: "1.1. Objet du document", seed: t.object || "" },
    { key: "situation", label: "3.1. Situation actuelle", seed: t.need.context || "" },
    { key: "objectifs", label: "3.2. Objectifs", seed: t.need.objectives || "" },
    { key: "parties_prenantes", label: "3.3. Principales parties prenantes", seed: `Responsable métier : ${t.responsibleMetier || A_COMPLETER}\nAcheteur : ${t.buyer || A_COMPLETER}` },
    { key: "perimetre", label: "3.4. Périmètre de l'offre", seed: [
        ...(t.scope.included || []).map(x => `Inclus : ${x}`),
        ...(t.scope.excluded || []).map(x => `Exclu : ${x}`),
      ].join("\n") || A_COMPLETER },
    { key: "evolutions", label: "3.5. Évolutions probables", seed: A_COMPLETER },
    { key: "phases", label: "3.6. Phasage du projet (si module « Projet phasé » activé)", seed: "Phase 1 – Validation technique et cadrage détaillé : " + A_COMPLETER },
    { key: "licences_desc", label: "4. Licences (si module « Licences » activé)", seed: A_COMPLETER },
    { key: "hebergement_desc", label: "4. Mode d'hébergement de la solution (si module « Hébergement / SaaS » activé)", seed: A_COMPLETER },
    { key: "support_desc", label: "4. Prestations de support et accompagnement technique (si module « Support » activé)", seed: A_COMPLETER },
    { key: "cadre_legal", label: "6.1. Cadre légal", seed: A_VALIDER },
    { key: "contraintes_reg", label: "6.2. Contraintes réglementaires du présent appel d'offre", seed: t.need.constraints || A_COMPLETER },
    { key: "paiement", label: "6.4. Modalités de paiement", seed: A_VALIDER },
    { key: "duree_contrat", label: "6.5. Durée du contrat (si module « Contrat-cadre » activé)", seed: A_VALIDER },
    { key: "mono_multi", label: "6.5. Mono/multi-adjudicataire (si module « Contrat-cadre » activé)", seed: A_VALIDER },
    { key: "montant_max", label: "6.5. Montant maximum du contrat (si module « Montants plafonds » activé)", seed: t.budget ? `${Number(t.budget).toLocaleString("fr-CH")} CHF (${A_VALIDER} pour le caractère ferme/plafond)` : A_COMPLETER },
    { key: "montant_annuel_max", label: "6.5. Montant annuel maximum (si module « Montants plafonds » activé)", seed: A_COMPLETER },
    { key: "resiliation", label: "6.6. Résiliation / Pénalités de retard", seed: A_COMPLETER },
  ],
  "SLA": (t) => [
    { key: "niveaux", label: "1. Niveaux de service couverts", seed: "" },
    { key: "disponibilite", label: "2. Disponibilité garantie", seed: "" },
    { key: "delais", label: "3. Délais de prise en charge et de résolution", seed: "" },
    { key: "penalites", label: "4. Pénalités en cas de non-respect", seed: "" },
    { key: "reporting", label: "5. Reporting et suivi", seed: "" },
    { key: "contacts", label: "6. Contacts et escalade", seed: `Responsable métier : ${t.responsibleMetier || "—"}` },
  ],
  "NDA": (t) => [
    { key: "parties", label: "1. Parties", seed: `L'organisation et le fournisseur soumissionnaire pour l'AO ${t.reference}.` },
    { key: "objet", label: "2. Objet de la confidentialité", seed: `Informations échangées dans le cadre de ${t.title}.` },
    { key: "duree", label: "3. Durée de l'obligation", seed: "" },
    { key: "obligations", label: "4. Obligations des parties", seed: "" },
    { key: "sanctions", label: "5. Sanctions en cas de manquement", seed: "" },
  ],
  "Contrat": (t) => [
    { key: "objet", label: "1. Objet du contrat", seed: t.object || "" },
    { key: "duree", label: "2. Durée", seed: "" },
    { key: "prix", label: "3. Prix et modalités de paiement", seed: t.budget ? `Budget de référence : ${Number(t.budget).toLocaleString("fr-CH")} CHF` : "" },
    { key: "obligations", label: "4. Obligations des parties", seed: "" },
    { key: "resiliation", label: "5. Résiliation", seed: "" },
    { key: "pi", label: "6. Propriété intellectuelle", seed: "" },
    { key: "droit", label: "7. Droit applicable et for", seed: A_VALIDER },
  ],
  "Développement durable": (t) => [
    { key: "engagements", label: "1. Engagements attendus du fournisseur", seed: "" },
    { key: "criteres", label: "2. Critères pris en compte dans l'évaluation", seed: "" },
  ],
};

function genericSchema(t) { return [{ key: "contenu", label: "Contenu du document", seed: "" }]; }

// Structure d'une série de prix : une ligne par position (référence, intitulé, quantité), le
// soumissionnaire complète le tarif unitaire — le tarif total par ligne et le total général
// sont calculés automatiquement (jamais saisis directement).
function defaultPriceScheduleRows() {
  return [{ ref: "", label: "", quantity: 1, unitPrice: "" }];
}

/* --------------------------------- UTILITAIRES --------------------------------- */

function chf(n) { if (!n && n !== 0) return "—"; return n.toLocaleString("fr-CH") + " CHF"; }

function computeProgress(t) {
  let score = 0, total = 6;
  if (t.need.context) score++;
  if (t.scope.included.length) score++;
  if (t.documents.length && t.documents.every(d => !d.mandatory || d.status !== "À préparer")) score++;
  if (t.requirements.length) score++;
  if (t.criteria.length && t.criteria.reduce((s, c) => s + c.weight, 0) === 100) score++;
  if (t.suppliers.length && t.suppliers.some(s => Object.keys(s.evaluations || {}).length)) score++;
  return Math.round((score / total) * 100);
}

function tabCompletion(t) {
  return {
    info: !!(t.direction && t.buyer),
    need: !!t.need.context && t.scope.included.length > 0,
    documents: t.documents.length > 0 && missingMandatoryDocs(t).length === 0,
    requirements: t.requirements.length > 0,
    criteria: t.criteria.length > 0 && weightSum(t) === 100,
    suppliers: t.suppliers.length > 0,
    evaluation: t.criteria.length > 0 && t.suppliers.length > 0 && t.suppliers.every(s => Object.keys(s.evaluations || {}).length >= t.criteria.filter(c => c.name !== "Prix").length),
    comparison: t.criteria.length > 0 && t.suppliers.length > 0,
    synthesis: t.criteria.length > 0 && t.suppliers.length > 0,
    history: true,
  };
}

// Checklist de préparation avant publication (phase 1 uniquement) — les soumissionnaires
// (phase 2, après réception des offres) n'en font pas partie, ils ne sont pas encore connus.
function readinessChecks(t) {
  const checks = [
    { label: "Besoin défini", ok: !!t.need.context }, { label: "Périmètre défini", ok: t.scope.included.length > 0 },
    { label: "Exigences définies", ok: t.requirements.length > 0 }, { label: "Critères définis", ok: t.criteria.length > 0 },
    { label: "Pondération = 100%", ok: t.criteria.length > 0 && weightSum(t) === 100 },
    { label: "Documents obligatoires prêts", ok: missingMandatoryDocs(t).length === 0 },
  ];
  const pct = Math.round((checks.filter(c => c.ok).length / checks.length) * 100);
  return { checks, pct, ready: pct === 100 };
}

const PUBLISHED_STATUSES = ["ongoing", "evaluation", "completed"];

// Détermine la prochaine action prioritaire pour guider le chef de projet.
function nextStepFor(t) {
  if (!t.need.context) return { tab: "need", icon: ListChecks, text: "Décrivez le besoin de cet AO pour démarrer.", cta: "Renseigner le besoin" };
  if (t.scope.included.length === 0) return { tab: "need", icon: ListChecks, text: "Précisez le périmètre (inclus / exclus) avant de poursuivre.", cta: "Définir le périmètre" };
  if (t.requirements.length === 0) return { tab: "requirements", icon: ClipboardCheck, text: "Ajoutez au moins une exigence.", cta: "Ajouter des exigences" };
  if (t.criteria.length === 0) return { tab: "criteria", icon: Scale, text: "Définissez les critères d'évaluation et leur pondération.", cta: "Définir les critères" };
  if (weightSum(t) !== 100) return { tab: "criteria", icon: Scale, text: `La pondération totalise ${weightSum(t)}% — ajustez-la à 100%.`, cta: "Corriger la pondération" };
  if (missingMandatoryDocs(t).length > 0) return { tab: "documents", icon: FileText, text: `${missingMandatoryDocs(t).length} document(s) obligatoire(s) à finaliser avant publication.`, cta: "Compléter les documents" };
  // Phase 1 (construction) terminée : les soumissionnaires ne sont pas encore connus, ce n'est
  // pas une tâche en attente mais le passage naturel à la phase 2, une fois l'AO publié.
  if (t.suppliers.length === 0) {
    if (!PUBLISHED_STATUSES.includes(t.status)) return { tab: "info", icon: CheckCircle2, text: "AO prêt à être publié — utilisez le bouton « Publier l'AO » en haut de page.", cta: "Voir le récapitulatif", done: true };
    return { tab: "suppliers", icon: Users, text: "AO publié — enregistrez ici les soumissionnaires au fil de la réception des offres.", cta: "Enregistrer un soumissionnaire", done: true };
  }
  const notFullyEvaluated = t.suppliers.some(s => Object.keys(s.evaluations || {}).length < t.criteria.filter(c => c.name !== "Prix").length);
  if (notFullyEvaluated) return { tab: "evaluation", icon: BarChart3, text: "Certains fournisseurs n'ont pas encore été notés sur tous les critères.", cta: "Poursuivre l'évaluation" };
  return { tab: "synthesis", icon: FileBarChart2, text: "L'AO est prêt : consultez la synthèse et exportez le dossier.", cta: "Voir la synthèse", done: true };
}

function missingMandatoryDocs(t) { return t.documents.filter(d => d.mandatory && !["Validé", "Reçu", "Généré", "Archivé"].includes(d.status)); }
function weightSum(t) { return t.criteria.reduce((s, c) => s + c.weight, 0); }

function priceScores(tender) {
  const totals = tender.suppliers.map(s => ({ id: s.id, total: (s.price.initial || 0) + (s.price.annual || 0) * 3 + (s.price.maintenance || 0) * 3 + (s.price.migration || 0) }));
  const min = Math.min(...totals.map(t => t.total).filter(v => v > 0));
  const out = {};
  totals.forEach(t => { out[t.id] = t.total > 0 ? +((min / t.total) * 5).toFixed(2) : 0; });
  return out;
}

// Exigences (non éliminatoires, incluses au CDC) rattachées à un critère — même logique de
// rapprochement par catégorie que la grille Excel du jury, réutilisée dans l'évaluation guidée.
function requirementsForCriterion(tender, criterionName) {
  const pool = tender.requirements.filter(r => r.criticality !== "Bloquante" && r.includeInCDC !== false);
  const exact = pool.filter(r => r.category && r.category.trim().toLowerCase() === criterionName.trim().toLowerCase());
  if (exact.length) return exact;
  return pool.filter(r => r.category && (r.category.toLowerCase().includes(criterionName.toLowerCase()) || criterionName.toLowerCase().includes(r.category.toLowerCase())));
}

function criterionAverage(supplier, criterionId) {
  const notes = supplier.evaluations?.[criterionId];
  if (!notes || !notes.length) return null;
  return notes.reduce((s, n) => s + n.note, 0) / notes.length;
}

function computeSupplierScores(tender) {
  const pScores = priceScores(tender);
  return tender.suppliers.map(sup => {
    let total = 0;
    const breakdown = tender.criteria.map(c => {
      const avg = c.name === "Prix" ? pScores[sup.id] : criterionAverage(sup, c.id);
      const weighted = avg != null ? (avg / 5) * c.weight : 0;
      total += weighted;
      return { criterionId: c.id, criterionName: c.name, avg, weighted, weight: c.weight };
    });
    return { supplierId: sup.id, supplierName: sup.name, breakdown, total: +total.toFixed(1) };
  }).sort((a, b) => b.total - a.total);
}

// Génère le classeur d'évaluation complet, sur le modèle du jury officiel : barème, pondérations,
// critères éliminatoires, une fiche d'évaluation vierge par fournisseur (catégories, exigences, notes et
// commentaires à main levée, moyenne/pondération/note finale/seuil calculés par formule), et une synthèse
// avec classement + recommandation d'adjudication + bloc de signature du jury (générique, sans nom réel).
function excelSheetName(base) { return String(base || "Fournisseur").replace(/[\\/?*[\]:]/g, "").slice(0, 31); }

function buildEvaluationWorkbookSheets(tender) {
  const threshold = tender.evaluationThreshold || 60;
  const sheets = [];

  // 1. Barème des notes
  const baremeRows = [["Note", "Signification"], [0, "Non conforme / absent"], [1, "Très insuffisant"], [2, "Insuffisant"], [3, "Conforme"], [4, "Bon"], [5, "Excellent"]];
  sheets.push({
    name: "Barème des notes", cols: [8, 50], rows: baremeRows,
    styles: [{ row: 0, cols: [0, 1], style: XLS_STYLE.tableHeader }],
  });

  // 2. Pondérations par catégorie
  const critRows = []; const critFormulas = []; const critStyles = []; const critMerges = [];
  critRows.push(["PONDÉRATIONS PAR CATÉGORIE"]); critMerges.push({ r1: 0, c1: 0, r2: 0, c2: 2 }); critStyles.push({ row: 0, cols: [0, 1, 2], style: XLS_STYLE.title });
  critRows.push([]);
  critRows.push(["N°", "Catégorie", "Pondération (%)"]); critStyles.push({ row: 2, cols: [0, 1, 2], style: XLS_STYLE.tableHeader });
  tender.criteria.forEach((c, i) => { critRows.push([i, c.name, c.weight]); critStyles.push({ row: critRows.length - 1, cols: [0, 1, 2], style: XLS_STYLE.data }); });
  const totalRow = critRows.length + 1;
  critRows.push([null, "TOTAL", null]); critStyles.push({ row: critRows.length - 1, cols: [0, 1, 2], style: XLS_STYLE.total });
  critFormulas.push({ ref: `C${totalRow}`, f: `SUM(C4:C${totalRow - 1})` });
  critRows.push([]);
  critRows.push(["Méthode de calcul :"]); critStyles.push({ row: critRows.length - 1, cols: [0], style: XLS_STYLE.label });
  critRows.push(["1. Chaque exigence est notée de 0 à 5"]);
  critRows.push(["2. Moyenne des notes par catégorie"]);
  critRows.push(["3. Note catégorie = (Moyenne/5) × Pondération × 100"]);
  critRows.push(["4. Note finale = Somme des notes catégories"]);
  sheets.push({ name: "Pondérations", cols: [6, 34, 16], rows: critRows, formulas: critFormulas, styles: critStyles, merges: critMerges, rowHeights: [22] });

  // 3. Critères éliminatoires
  const elimRequirements = tender.requirements.filter(r => r.criticality === "Bloquante" && r.includeInCDC !== false);
  const elimRows = []; const elimStyles = []; const elimMerges = [];
  const elimCols = 2 + tender.suppliers.length * 2;
  elimRows.push(["Critères éliminatoires"]); elimMerges.push({ r1: 0, c1: 0, r2: 0, c2: elimCols - 1 }); elimStyles.push({ row: 0, cols: Array.from({ length: elimCols }, (_, i) => i), style: XLS_STYLE.title });
  elimRows.push(["En cas de non-conformité, l'organisation se réserve le droit d'éliminer le soumissionnaire (délibération)."]); elimMerges.push({ r1: 1, c1: 0, r2: 1, c2: elimCols - 1 }); elimStyles.push({ row: 1, cols: [0], style: XLS_STYLE.subtitle });
  const elimHeader = ["N°", "Catégorie / Exigence"];
  tender.suppliers.forEach(s => { elimHeader.push(`${s.name}\nConforme`); elimHeader.push("Comm."); });
  elimRows.push(elimHeader); elimStyles.push({ row: 2, cols: Array.from({ length: elimCols }, (_, i) => i), style: XLS_STYLE.tableHeader });
  if (elimRequirements.length === 0) {
    elimRows.push(["", "Aucune exigence éliminatoire définie pour cet AO.", ...tender.suppliers.flatMap(() => ["", ""])]);
  } else {
    elimRequirements.forEach(r => { elimRows.push([r.id, r.description, ...tender.suppliers.flatMap(() => ["", ""])]); elimStyles.push({ row: elimRows.length - 1, cols: Array.from({ length: elimCols }, (_, i) => i), style: XLS_STYLE.data }); });
  }
  sheets.push({ name: "Critères Éliminatoires", cols: [8, 60, ...tender.suppliers.flatMap(() => [12, 24])], rows: elimRows, styles: elimStyles, merges: elimMerges, rowHeights: [22] });

  // 4. Une fiche d'évaluation par fournisseur (structure identique pour tous, pour permettre les renvois du classement)
  const nonPriceCriteria = tender.criteria.filter(c => c.name !== "Prix");
  const priceCriterion = tender.criteria.find(c => c.name === "Prix");
  const evalRequirements = tender.requirements.filter(r => r.criticality !== "Bloquante" && r.includeInCDC !== false);
  let noteFinaleRowNum = null;
  const supplierSheetNames = {};

  tender.suppliers.forEach(supplier => {
    const rows = []; const formulas = []; const styles = []; const merges = []; const noteRanges = [];
    rows.push([`ÉVALUATION - ${supplier.name}`]); merges.push({ r1: 0, c1: 0, r2: 0, c2: 3 }); styles.push({ row: 0, cols: [0, 1, 2, 3], style: XLS_STYLE.title });
    rows.push([]);
    rows.push(["Soumissionnaire :", supplier.name]); styles.push({ row: 2, cols: [0], style: XLS_STYLE.label });
    rows.push([null, null, "Critères de sélection"]); styles.push({ row: 3, cols: [2], style: XLS_STYLE.label });
    rows.push(["N°", "Catégorie / Exigence", "Note /5", "Commentaires"]); styles.push({ row: 4, cols: [0, 1, 2, 3], style: XLS_STYLE.tableHeader });
    const categoryRefs = [];
    nonPriceCriteria.forEach((c, idx) => {
      rows.push([idx, `${c.name} (${c.weight}%)`]); styles.push({ row: rows.length - 1, cols: [0, 1, 2, 3], style: XLS_STYLE.category });
      const exactMatch = evalRequirements.filter(r => r.category && r.category.trim().toLowerCase() === c.name.trim().toLowerCase());
      const matched = exactMatch.length ? exactMatch : evalRequirements.filter(r => r.category && (r.category.toLowerCase().includes(c.name.toLowerCase()) || c.name.toLowerCase().includes(r.category.toLowerCase())));
      const startRow = rows.length + 1;
      if (matched.length === 0) {
        rows.push(["", "(Aucune exigence spécifique associée — noter la catégorie globalement)", "", ""]);
      } else {
        matched.forEach(r => rows.push([r.id, r.description, "", ""]));
      }
      for (let rr = startRow - 1; rr < rows.length; rr++) styles.push({ row: rr, cols: [0, 1, 2, 3], style: XLS_STYLE.data });
      const endRow = rows.length;
      noteRanges.push(`C${startRow}:C${endRow}`);
      rows.push([null, "Moyenne catégorie", null, null]); styles.push({ row: rows.length - 1, cols: [0, 1, 2, 3], style: XLS_STYLE.total });
      const avgRow = rows.length;
      formulas.push({ ref: `C${avgRow}`, f: `IFERROR(AVERAGE(C${startRow}:C${endRow}),0)` });
      rows.push([]);
      categoryRefs.push({ name: c.name, weight: c.weight, avgRow });
    });
    if (priceCriterion) {
      const pScore = priceScores(tender)[supplier.id] ?? 0;
      rows.push([nonPriceCriteria.length, `${priceCriterion.name} (${priceCriterion.weight}%) — calculé automatiquement (prix le plus bas = note maximale)`]); styles.push({ row: rows.length - 1, cols: [0, 1, 2, 3], style: XLS_STYLE.category });
      rows.push(["", "Score prix (voir onglet Évaluation de l'application pour le détail du TCO)", +pScore.toFixed(2), ""]); styles.push({ row: rows.length - 1, cols: [0, 1, 2, 3], style: XLS_STYLE.data });
      const priceRow = rows.length;
      rows.push([]);
      categoryRefs.push({ name: priceCriterion.name, weight: priceCriterion.weight, avgRow: priceRow });
    }
    rows.push(["N°", "Catégorie", "Moyenne /5", "Pondération", "Note pondérée"]); styles.push({ row: rows.length - 1, cols: [0, 1, 2, 3, 4], style: XLS_STYLE.tableHeader });
    const finalStart = rows.length + 1;
    categoryRefs.forEach((cat, i) => {
      rows.push([i, cat.name, null, cat.weight, null]); styles.push({ row: rows.length - 1, cols: [0, 1, 2, 3, 4], style: XLS_STYLE.data });
      const r = rows.length;
      formulas.push({ ref: `C${r}`, f: `C${cat.avgRow}` });
      formulas.push({ ref: `E${r}`, f: `(C${r}/5)*D${r}*100` });
    });
    const finalEnd = rows.length;
    rows.push([null, "NOTE FINALE /100", null, null, null]); styles.push({ row: rows.length - 1, cols: [0, 1, 2, 3, 4], style: XLS_STYLE.total });
    noteFinaleRowNum = rows.length;
    formulas.push({ ref: `E${noteFinaleRowNum}`, f: `SUM(E${finalStart}:E${finalEnd})` });
    rows.push(["Seuil minimum :", null, null, null, threshold]); styles.push({ row: rows.length - 1, cols: [0], style: XLS_STYLE.label });
    const seuilRow = rows.length;
    rows.push(["STATUT :", null, null, null, null]); styles.push({ row: rows.length - 1, cols: [0, 4], style: XLS_STYLE.total });
    const statutRow = rows.length;
    formulas.push({ ref: `E${statutRow}`, f: `IF(E${noteFinaleRowNum}>=E${seuilRow},"ADMIS","REFUSÉ")` });

    const sheetName = excelSheetName(`Éval. ${supplier.name}`);
    supplierSheetNames[supplier.id] = sheetName;
    sheets.push({ name: sheetName, cols: [8, 55, 12, 12, 14], rows, formulas, styles, merges, rowHeights: [22], dataValidations: noteRanges.length ? [{ sqref: noteRanges.join(" "), list: ["0", "1", "2", "3", "4", "5"] }] : undefined });
  });

  // 5. Synthèse et classement
  const synRows = []; const synFormulas = []; const synStyles = []; const synMerges = [];
  synRows.push([`GRILLE D'ÉVALUATION - ${tender.title || "Appel d'offres"}`]); synMerges.push({ r1: 0, c1: 0, r2: 0, c2: 5 }); synStyles.push({ row: 0, cols: [0, 1, 2, 3, 4, 5], style: XLS_STYLE.title });
  synRows.push(["Synthèse des évaluations et classement final"]); synMerges.push({ r1: 1, c1: 0, r2: 1, c2: 5 }); synStyles.push({ row: 1, cols: [0], style: XLS_STYLE.subtitle });
  synRows.push([]);
  synRows.push(["CLASSEMENT DES SOUMISSIONNAIRES"]); synMerges.push({ r1: 3, c1: 0, r2: 3, c2: 5 }); synStyles.push({ row: 3, cols: [0, 1, 2, 3, 4, 5], style: XLS_STYLE.category });
  synRows.push(["Rang", "Soumissionnaire", "Note /100", "Seuil min.", "Statut", "Observations"]); synStyles.push({ row: 4, cols: [0, 1, 2, 3, 4, 5], style: XLS_STYLE.tableHeader });
  tender.suppliers.forEach((supplier, i) => {
    synRows.push([i + 1, supplier.name, null, threshold, null, ""]);
    synStyles.push({ row: synRows.length - 1, cols: [0, 1, 2, 3, 4, 5], style: XLS_STYLE.data });
    const r = synRows.length;
    if (noteFinaleRowNum) synFormulas.push({ ref: `C${r}`, f: `'${supplierSheetNames[supplier.id]}'!E${noteFinaleRowNum}` });
    synFormulas.push({ ref: `E${r}`, f: `IF(C${r}>=D${r},"ADMIS","REFUSÉ")` });
  });
  synRows.push([]);
  synRows.push(["RECOMMANDATION D'ADJUDICATION"]); synMerges.push({ r1: synRows.length - 1, c1: 0, r2: synRows.length - 1, c2: 5 }); synStyles.push({ row: synRows.length - 1, cols: [0, 1, 2, 3, 4, 5], style: XLS_STYLE.category });
  synRows.push([]);
  synRows.push(["Soumissionnaire recommandé :", ""]);
  synRows.push(["Date de la séance d'évaluation :", ""]);
  synRows.push(["Président du jury :", "", "Signature :", ""]);
  synRows.push([]);
  synRows.push(["Membres du jury :"]);
  for (let i = 1; i <= 4; i++) synRows.push([`Membre ${i} :`, ""]);
  sheets.push({ name: "Synthèse et Classement", cols: [8, 34, 12, 12, 12, 30], rows: synRows, formulas: synFormulas, styles: synStyles, merges: synMerges, rowHeights: [22] });

  return sheets;
}

/* --------------------------- EXPORT : vrai .docx (ZIP/OOXML) et XLSX (Excel) --------------------------- */
// Aucune librairie de génération docx n'étant disponible dans ce bac à sable navigateur,
// le fichier .docx est construit ici directement : un .docx est un zip contenant du XML.
// On écrit donc un mini writer ZIP (méthode "stored", sans compression) et le XML Word minimal nécessaire.

function downloadBlob(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// Palette et bordures réutilisées pour la mise en forme des classeurs générés.
const XLS_BORDER = { top: { style: "thin", color: { rgb: "D9D9D9" } }, bottom: { style: "thin", color: { rgb: "D9D9D9" } }, left: { style: "thin", color: { rgb: "D9D9D9" } }, right: { style: "thin", color: { rgb: "D9D9D9" } } };
const XLS_STYLE = {
  title: { fill: { patternType: "solid", fgColor: { rgb: "182234" } }, font: { bold: true, sz: 14, color: { rgb: "FFFFFF" } }, alignment: { vertical: "center", horizontal: "left", wrapText: true } },
  subtitle: { font: { italic: true, sz: 10, color: { rgb: "666666" } } },
  tableHeader: { fill: { patternType: "solid", fgColor: { rgb: "FD5312" } }, font: { bold: true, sz: 11, color: { rgb: "FFFFFF" } }, alignment: { vertical: "center", horizontal: "center", wrapText: true }, border: XLS_BORDER },
  category: { fill: { patternType: "solid", fgColor: { rgb: "EEEDE7" } }, font: { bold: true, sz: 11, color: { rgb: "182234" } }, border: XLS_BORDER },
  data: { border: XLS_BORDER, alignment: { vertical: "top", wrapText: true } },
  total: { fill: { patternType: "solid", fgColor: { rgb: "FFEDE4" } }, font: { bold: true, sz: 12, color: { rgb: "182234" } }, border: XLS_BORDER },
  label: { font: { bold: true, sz: 11, color: { rgb: "182234" } } },
};
// Correspondance entre les styles ci-dessus et l'index de cellXfs dans styles.xml (cf. downloadXLSX).
const XLS_STYLE_XF_ENTRIES = [[XLS_STYLE.title, 1], [XLS_STYLE.subtitle, 2], [XLS_STYLE.tableHeader, 3], [XLS_STYLE.category, 4], [XLS_STYLE.data, 5], [XLS_STYLE.total, 6], [XLS_STYLE.label, 7]];

// Générateur XLSX maison (OOXML assemblé à la main, comme les .docx plus bas) — remplace l'usage
// de la librairie "xlsx" pour ÉCRIRE des fichiers : son édition gratuite accepte bien qu'on lui
// passe des styles de cellule (ws[addr].s = ...) mais les ignore silencieusement à l'écriture, et
// ne sait pas du tout écrire de validation de données (listes déroulantes). Résultat concret avant
// ce correctif : aucun des classeurs générés (Exigences, grille du jury, etc.) n'affichait la mise
// en forme orange/marine pourtant définie dans XLS_STYLE. Ce writer applique réellement les styles
// et supporte les listes déroulantes par feuille.
const XLS_STYLE_XF = new Map(XLS_STYLE_XF_ENTRIES);

function xlsxColLetter(c) {
  let s = "", n = c + 1;
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function downloadXLSX(filename, sheets) {
  const sheetParts = []; const sheetRelsOverrides = []; const workbookSheetEls = []; const workbookRelEls = [];
  sheets.forEach((sh, sIdx) => {
    const sheetId = sIdx + 1;
    const cellsByAddr = new Map(); // "r,c" -> { value, xf, formula }
    const getXf = (style) => (style && XLS_STYLE_XF.has(style)) ? XLS_STYLE_XF.get(style) : 0;

    (sh.rows || []).forEach((rowArr, r) => {
      rowArr.forEach((value, c) => {
        if (value === null || value === undefined || value === "") return;
        cellsByAddr.set(`${r},${c}`, { value });
      });
    });
    (sh.styles || []).forEach(({ row, cols, style }) => {
      const xf = getXf(style);
      (cols || [0]).forEach(c => {
        const key = `${row},${c}`;
        const existing = cellsByAddr.get(key) || {};
        cellsByAddr.set(key, { ...existing, xf });
      });
    });
    (sh.formulas || []).forEach(({ ref, f }) => {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      if (!m) return;
      let c = 0; for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64); c -= 1;
      const r = Number(m[2]) - 1;
      const key = `${r},${c}`;
      const existing = cellsByAddr.get(key) || {};
      cellsByAddr.set(key, { ...existing, formula: f });
    });

    let maxRow = 0, maxCol = 0;
    cellsByAddr.forEach((_, key) => { const [r, c] = key.split(",").map(Number); if (r > maxRow) maxRow = r; if (c > maxCol) maxCol = c; });
    (sh.merges || []).forEach(m => { if (m.r2 > maxRow) maxRow = m.r2; if (m.c2 > maxCol) maxCol = m.c2; });

    const rowsMap = new Map();
    cellsByAddr.forEach((cellInfo, key) => {
      const [r, c] = key.split(",").map(Number);
      if (!rowsMap.has(r)) rowsMap.set(r, []);
      rowsMap.get(r).push({ c, ...cellInfo });
    });

    let sheetDataXml = "";
    for (let r = 0; r <= maxRow; r++) {
      const cells = (rowsMap.get(r) || []).sort((a, b) => a.c - b.c);
      if (cells.length === 0 && !(sh.rowHeights && sh.rowHeights[r])) continue;
      const h = sh.rowHeights && sh.rowHeights[r] ? ` ht="${sh.rowHeights[r]}" customHeight="1"` : "";
      let rowXml = `<row r="${r + 1}"${h}>`;
      cells.forEach(cl => {
        const ref = `${xlsxColLetter(cl.c)}${r + 1}`;
        const xfAttr = cl.xf ? ` s="${cl.xf}"` : "";
        if (cl.formula) {
          rowXml += `<c r="${ref}"${xfAttr}><f>${xmlEscape(cl.formula)}</f></c>`;
        } else if (typeof cl.value === "number") {
          rowXml += `<c r="${ref}"${xfAttr}><v>${cl.value}</v></c>`;
        } else if (cl.value !== undefined) {
          rowXml += `<c r="${ref}" t="inlineStr"${xfAttr}><is><t xml:space="preserve">${xmlEscape(String(cl.value))}</t></is></c>`;
        } else {
          rowXml += `<c r="${ref}"${xfAttr}/>`;
        }
      });
      rowXml += "</row>";
      sheetDataXml += rowXml;
    }

    const colsXml = sh.cols ? `<cols>${sh.cols.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>` : "";
    const mergesXml = (sh.merges || []).length ? `<mergeCells count="${sh.merges.length}">${sh.merges.map(m => `<mergeCell ref="${xlsxColLetter(m.c1)}${m.r1 + 1}:${xlsxColLetter(m.c2)}${m.r2 + 1}"/>`).join("")}</mergeCells>` : "";
    const dvXml = (sh.dataValidations || []).length ? `<dataValidations count="${sh.dataValidations.length}">${sh.dataValidations.map(dv => `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Valeur invalide" error="Choisissez une valeur dans la liste." sqref="${dv.sqref}"><formula1>"${dv.list.join(",")}"</formula1></dataValidation>`).join("")}</dataValidations>` : "";

    sheetParts.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${colsXml}<sheetData>${sheetDataXml}</sheetData>${mergesXml}${dvXml}</worksheet>`);

    const safeName = xmlEscape(excelSheetName(sh.name));
    workbookSheetEls.push(`<sheet name="${safeName}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`);
    workbookRelEls.push(`<Relationship Id="rId${sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheetId}.xml"/>`);
    sheetRelsOverrides.push(`<Override PartName="/xl/worksheets/sheet${sheetId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
  });
  workbookRelEls.push(`<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${workbookSheetEls.join("")}</sheets>
  <calcPr calcId="0" fullCalcOnLoad="1"/>
</workbook>`;

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetRelsOverrides.join("\n  ")}
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${workbookRelEls.join("\n  ")}
</Relationships>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="8">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><i/><sz val="10"/><color rgb="FF666666"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF182234"/><name val="Calibri"/></font>
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FF182234"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF182234"/><name val="Calibri"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF182234"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFD5312"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEEEDE7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFEDE4"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="6" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const files = [
    { name: "[Content_Types].xml", data: strToBytes(contentTypesXml) },
    { name: "_rels/.rels", data: strToBytes(rootRelsXml) },
    { name: "xl/workbook.xml", data: strToBytes(workbookXml) },
    { name: "xl/_rels/workbook.xml.rels", data: strToBytes(workbookRelsXml) },
    { name: "xl/styles.xml", data: strToBytes(stylesXml) },
    ...sheetParts.map((xml, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: strToBytes(xml) })),
  ];
  downloadBlob(filename, makeZip(files), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

// Classeur "Cahier de réponses" — mêmes styles que le reste de l'app, avec une vraie liste
// déroulante Excel sur la colonne Réponse (Oui / Non / Partiellement).
// Construit une feuille "rapport" standard : bandeau titre marine, en-tête de colonnes orange,
// lignes de données bordées — le même habillage sur tous les classeurs générés par l'app.
function styledReportSheet(name, title, headers, dataRows, colWidths, dataValidations) {
  const nCols = headers.length;
  const allCols = Array.from({ length: nCols }, (_, i) => i);
  const rows = [[title], headers, ...dataRows];
  const styles = [
    { row: 0, cols: allCols, style: XLS_STYLE.title },
    { row: 1, cols: allCols, style: XLS_STYLE.tableHeader },
  ];
  dataRows.forEach((_, i) => styles.push({ row: i + 2, cols: allCols, style: XLS_STYLE.data }));
  return { name, cols: colWidths, rows, styles, merges: [{ r1: 0, c1: 0, r2: 0, c2: nCols - 1 }], rowHeights: [22], dataValidations };
}

// Un onglet par catégorie d'exigence plutôt qu'une feuille unique — c'est la convention réellement
// utilisée dans un cahier de réponses (chaque domaine/catégorie a son propre onglet), ce qui rend le
// classeur exploitable directement par le comité d'évaluation et fait le lien explicite avec la
// section « Exigences et spécifications » du cahier des charges (cf. buildCahierDesChargesBody).
function downloadResponseSheetXLSX(filename, sheetName, requirements) {
  const header = ["ID", "Exigence", "Criticité", "Obligatoire", "Réponse", "Justificatif / commentaire"];
  const byCat = {};
  requirements.forEach(r => { (byCat[r.category] = byCat[r.category] || []).push(r); });
  const categories = Object.keys(byCat);
  const sheets = (categories.length ? categories : [sheetName]).map(cat => {
    const reqs = byCat[cat] || [];
    const dataRows = reqs.map(r => [r.id, r.description, r.criticality, r.mandatory ? "Oui" : "Non", "", ""]);
    const lastRow = reqs.length + 2;
    return styledReportSheet(cat, cat.toUpperCase(), header, dataRows, [12, 55, 14, 12, 18, 45],
      reqs.length ? [{ sqref: `E3:E${lastRow}`, list: ["Oui", "Non", "Partiellement"] }] : []);
  });
  downloadXLSX(filename, sheets);
}

/* ---- ZIP writer (stored / non compressé) ---- */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ bytes[i]) & 0xFF];
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
function strToBytes(str) { return new TextEncoder().encode(str); }
function base64ToBytes(b64) {
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}
function makeZip(files) {
  const time = 0, date = ((2026 - 1980) << 9) | (1 << 5) | 1;
  const localParts = [], centralParts = [];
  let offset = 0;
  files.forEach(f => {
    const nameBytes = strToBytes(f.name);
    const crc = crc32(f.data);
    const size = f.data.length;
    const local = new Uint8Array(30 + nameBytes.length + size);
    const dv = new DataView(local.buffer);
    dv.setUint32(0, 0x04034b50, true); dv.setUint16(4, 20, true); dv.setUint16(6, 0, true);
    dv.setUint16(8, 0, true); dv.setUint16(10, time, true); dv.setUint16(12, date, true);
    dv.setUint32(14, crc, true); dv.setUint32(18, size, true); dv.setUint32(22, size, true);
    dv.setUint16(26, nameBytes.length, true); dv.setUint16(28, 0, true);
    local.set(nameBytes, 30); local.set(f.data, 30 + nameBytes.length);
    localParts.push(local);

    const central = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(central.buffer);
    cdv.setUint32(0, 0x02014b50, true); cdv.setUint16(4, 20, true); cdv.setUint16(6, 20, true);
    cdv.setUint16(8, 0, true); cdv.setUint16(10, 0, true); cdv.setUint16(12, time, true); cdv.setUint16(14, date, true);
    cdv.setUint32(16, crc, true); cdv.setUint32(20, size, true); cdv.setUint32(24, size, true);
    cdv.setUint16(28, nameBytes.length, true); cdv.setUint16(30, 0, true); cdv.setUint16(32, 0, true);
    cdv.setUint16(34, 0, true); cdv.setUint16(36, 0, true); cdv.setUint32(38, 0, true);
    cdv.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);
    offset += local.length;
  });
  const centralStart = offset;
  let centralSize = 0; centralParts.forEach(c => centralSize += c.length);
  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, 0x06054b50, true); edv.setUint16(8, files.length, true); edv.setUint16(10, files.length, true);
  edv.setUint32(12, centralSize, true); edv.setUint32(16, centralStart, true);
  const out = new Uint8Array(offset + centralSize + 22);
  let p = 0;
  localParts.forEach(l => { out.set(l, p); p += l.length; });
  centralParts.forEach(c => { out.set(c, p); p += c.length; });
  out.set(eocd, p);
  return out;
}

/* ---- Générateurs de fragments WordprocessingML (document.xml) ---- */
function xmlEscape(s) { return String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function runProps({ bold, italic, size, color, font }) {
  let p = "";
  if (bold) p += "<w:b/>";
  if (italic) p += "<w:i/>";
  if (font) p += `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`;
  if (color) p += `<w:color w:val="${color}"/>`;
  if (size) p += `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`;
  return p ? `<w:rPr>${p}</w:rPr>` : "";
}
function runsForText(text, opts = {}) {
  const lines = String(text ?? "").split("\n");
  return lines.map((line, i) => `<w:r>${runProps(opts)}<w:t xml:space="preserve">${xmlEscape(line)}</w:t>${i < lines.length - 1 ? "<w:br/>" : ""}</w:r>`).join("");
}
function para(text, { align, spacingAfter = 160, spacingBefore = 0, shade, ...runOpts } = {}) {
  const jc = align ? `<w:jc w:val="${align}"/>` : "";
  const shd = shade ? `<w:shd w:val="clear" w:fill="${shade}"/>` : "";
  return `<w:p><w:pPr><w:spacing w:before="${spacingBefore}" w:after="${spacingAfter}"/>${jc}${shd}</w:pPr>${runsForText(text, runOpts)}</w:p>`;
}
// Palette du template Word — identique à celle des classeurs Excel générés (marine + orange),
// pour une identité visuelle cohérente sur tous les documents produits par l'app.
const DOCX_NAVY = "182234", DOCX_ORANGE = "FD5312", DOCX_ORANGE_SOFT = "FFEDE4", DOCX_GRAY_SOFT = "EEEDE7", DOCX_BORDER = "D9D9D9", DOCX_INK_SOFT = "4B5872";
let __bookmarkSeq = 0;
function heading(text, level = 1, bookmarkId) {
  const sizes = { 1: 30, 2: 24, 3: 21 };
  let bm = "";
  if (bookmarkId) {
    __bookmarkSeq++;
    bm = `<w:bookmarkStart w:id="${__bookmarkSeq}" w:name="${bookmarkId}"/><w:bookmarkEnd w:id="${__bookmarkSeq}"/>`;
  }
  const border = level === 1 ? `<w:pBdr><w:bottom w:val="single" w:sz="10" w:space="6" w:color="${DOCX_ORANGE}"/></w:pBdr>` : "";
  return `<w:p><w:pPr><w:spacing w:before="320" w:after="160"/><w:keepNext/>${border}</w:pPr>${bm}<w:r><w:rPr><w:b/><w:sz w:val="${sizes[level] || 26}"/><w:szCs w:val="${sizes[level] || 26}"/><w:color w:val="${DOCX_NAVY}"/></w:rPr><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
}
// Entrée de table des matières cliquable, pointant vers le signet (bookmark) posé par heading().
function tocLink(text, bookmarkId, level = 1) {
  const indent = (level - 1) * 300;
  const size = level === 1 ? 20 : 18;
  return `<w:p><w:pPr><w:ind w:left="${indent}"/><w:spacing w:after="60"/></w:pPr><w:hyperlink w:anchor="${bookmarkId}"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/><w:sz w:val="${size}"/><w:b w:val="${level === 1 ? "1" : "0"}"/></w:rPr><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:hyperlink></w:p>`;
}
// Constructeur de chapitres qui enregistre automatiquement chaque titre pour la table des matières.
function makeChapterBuilder() {
  const toc = [];
  let body = "";
  function H(text, level, id) { toc.push({ text, level, id }); body += heading(text, level, id); }
  return { H, toc: () => toc, body: () => body, raw: (x) => { body += x; } };
}
function pageBreak() { return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`; }
function docTable(headerCells, rows) {
  const cols = headerCells.length || 1;
  const totalWidth = 9020;
  const colWidth = Math.floor(totalWidth / cols);
  const grid = Array(cols).fill(0).map(() => `<w:gridCol w:w="${colWidth}"/>`).join("");
  const borders = `<w:tblBorders><w:top w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:left w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:bottom w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:right w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:insideH w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:insideV w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/></w:tblBorders>`;
  function cell(text, opts = {}) {
    const shd = opts.shade ? `<w:shd w:val="clear" w:fill="${opts.shade}"/>` : "";
    return `<w:tc><w:tcPr><w:tcW w:w="${colWidth}" w:type="dxa"/>${shd}<w:vAlign w:val="center"/></w:tcPr><w:p>${runsForText(text, { bold: opts.bold, size: 18, color: opts.color })}</w:p></w:tc>`;
  }
  const headerTrPr = headerCells.some(h => h) ? `<w:trPr><w:tblHeader/></w:trPr>` : "";
  const headerTr = headerCells.some(h => h) ? `<w:tr>${headerTrPr}${headerCells.map(h => cell(h, { bold: true, shade: DOCX_ORANGE, color: "FFFFFF" })).join("")}</w:tr>` : "";
  const dataRows = rows.map((r, i) => `<w:tr>${r.map(v => cell(v, i % 2 === 1 ? { shade: "FBFAF8" } : {})).join("")}</w:tr>`).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="${totalWidth}" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid>${grid}</w:tblGrid>${headerTr}${dataRows}</w:tbl><w:p/>`;
}
function shadedBox(lines, fill = DOCX_GRAY_SOFT) {
  const content = lines.map(l => para(l.text, { align: "center", bold: l.bold, italic: l.italic, size: l.size || 20, spacingAfter: 40, color: l.bold ? DOCX_NAVY : undefined })).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="9020" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="6" w:color="${DOCX_NAVY}"/><w:left w:val="single" w:sz="6" w:color="${DOCX_NAVY}"/><w:bottom w:val="single" w:sz="6" w:color="${DOCX_NAVY}"/><w:right w:val="single" w:sz="6" w:color="${DOCX_NAVY}"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="9020"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9020" w:type="dxa"/><w:shd w:val="clear" w:fill="${fill}"/></w:tcPr>${content}</w:tc></w:tr></w:tbl><w:p/>`;
}
function imageParagraph(cx, cy, align = "left") {
  const jc = align ? `<w:jc w:val="${align}"/>` : "";
  return `<w:p><w:pPr>${jc}</w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="logo"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

function coverBlockXml(tender, docTypeLabel) {
  let x = "";
  x += `<w:p><w:pPr><w:spacing w:before="0" w:after="200"/><w:pBdr><w:bottom w:val="single" w:sz="14" w:space="8" w:color="${DOCX_ORANGE}"/></w:pBdr></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="${DOCX_ORANGE}"/></w:rPr><w:t xml:space="preserve">APPEL D'OFFRES</w:t></w:r></w:p>`;
  x += para(`AO – ${tender.title}`, { align: "center", bold: true, size: 34, color: DOCX_NAVY, spacingAfter: 80, spacingBefore: 400 });
  x += para("Appel d'offre", { align: "center", size: 22, color: DOCX_INK_SOFT, spacingAfter: 20 });
  x += para(docTypeLabel, { align: "center", bold: true, size: 24, color: DOCX_ORANGE, spacingAfter: 500 });
  x += shadedBox([
    { text: "Soumissionnaire", bold: true, size: 20 },
    { text: "Raison sociale – Adresse – Timbre", size: 18 },
    { text: "(Date et signature : dernière page)", italic: true, size: 16 },
  ]);
  x += para("", { spacingAfter: 2000 });
  x += para("Ce document est confidentiel. Il est fourni sous la condition qu'il ne soit ni reproduit, ni copié, ni prêté ou divulgué, directement ou indirectement, sans autorisation.", { italic: true, size: 16, color: DOCX_INK_SOFT, shade: DOCX_GRAY_SOFT });
  x += pageBreak();
  return x;
}

function metaAndHistoryXml(tender) {
  let x = "";
  x += heading("Historique du document", 2);
  x += docTable(["", ""], [["Chef de Projet", tender.responsibleMetier || "—"], ["Statut", "Travail"]]);
  x += docTable(["Version", "Date du changement", "Raison"], [["0.1", new Date().toLocaleDateString("fr-CH"), "Création"]]);
  x += para("", { spacingAfter: 120 });
  x += heading("Approbation", 2);
  x += docTable(["Nom et Service", "Date", "Signature"], [["COPIL", "", ""], ["Équipe projet", "", ""], ["Achats", "", ""]]);
  x += para("", { spacingAfter: 200 });
  return x;
}

// Table label/valeur (2 colonnes, sans en-tête) — utilisée pour les fiches d'exigence.
function labelValueTable(rows) {
  const totalWidth = 9020, c1 = Math.floor(totalWidth * 0.35), c2 = totalWidth - c1;
  const borders = `<w:tblBorders><w:top w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:left w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:bottom w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:right w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:insideH w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/><w:insideV w:val="single" w:sz="4" w:color="${DOCX_BORDER}"/></w:tblBorders>`;
  const trs = rows.map(([label, value]) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="${c1}" w:type="dxa"/><w:shd w:val="clear" w:fill="${DOCX_GRAY_SOFT}"/><w:vAlign w:val="center"/></w:tcPr><w:p>${runsForText(label, { bold: true, size: 18, color: DOCX_NAVY })}</w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="${c2}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr><w:p>${runsForText(value, { size: 18 })}</w:p></w:tc></w:tr>`).join("");
  return `<w:tbl><w:tblPr><w:tblW w:w="${totalWidth}" w:type="dxa"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="${c1}"/><w:gridCol w:w="${c2}"/></w:tblGrid>${trs}</w:tbl><w:p/>`;
}

// Fiche d'exigence au format officiel : Essentielle (O/N) / Titre / Description / Remarque / Criticité / Conformité / Réponse.
function requirementCardXml(req) {
  return labelValueTable([
    ["Référence", req.id],
    ["Essentielle (O/N)", req.mandatory ? "O" : "N"],
    ["Description", req.description],
    ["Remarque", req.remark || "—"],
    ["Criticité", req.criticality],
    ["Conformité (CO / PC / NC)", ""],
    ["Réponse", ""],
  ]);
}

function conformiteLegendXml() {
  return docTable(["Code", "Signification"], [
    ["CO", "Conforme à l'exigence — le soumissionnaire n'a aucune réserve à émettre."],
    ["PC", "Partiellement conforme — la partie non couverte est explicitée en Réponse."],
    ["NC", "Non conforme — le soumissionnaire ne peut pas répondre à l'exigence."],
  ]);
}

function packageDocx(bodyXml, filename, tender) {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;
  const coreProps = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${xmlEscape(filename)}</dc:title><dc:creator>AO Manager</dc:creator></cp:coreProperties>`;

  const files = [
    { name: "[Content_Types].xml", data: strToBytes(contentTypes) },
    { name: "_rels/.rels", data: strToBytes(rootRels) },
    { name: "docProps/core.xml", data: strToBytes(coreProps) },
    { name: "word/document.xml", data: strToBytes(documentXml) },
  ];
  const zipBytes = makeZip(files);
  downloadBlob(filename, zipBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
}

// Glossaire standard ajouté en fin de document généré.
const GLOSSARY_TERMS = [
  ["AO", "Appel d'offres"],
    ["CDC", "Cahier des charges — document définissant le besoin, le contexte et les exigences"],
  ["SLA", "Service Level Agreement — accord de niveau de service"],
  ["NDA", "Non-Disclosure Agreement — accord de confidentialité"],
  ["CO / PC / NC", "Conforme / Partiellement conforme / Non conforme — niveaux de réponse du soumissionnaire à une exigence"],
  ["Bloquante", "Exigence éliminatoire : son non-respect entraîne l'exclusion de l'offre, sans notation sur une échelle"],
  ["Gré à gré", "Attribution directe à un fournisseur choisi, sans mise en concurrence formelle"],
  ["Procédure sur invitation", "Consultation d'au moins 3 fournisseurs choisis, sans publication publique"],
  ["Procédure ouverte", "AO publié publiquement (SIMAP), accessible à tout fournisseur intéressé"],
];
function glossaryXml() {
  let x = pageBreak();
  x += heading("Glossaire & abréviations", 1);
  x += docTable(["Terme", "Définition"], GLOSSARY_TERMS);
  return x;
}

function generateDocumentFile(tender, doc) {
  if (QUESTIONS_DOC_NAMES.includes(doc.name)) {
    const questions = tender.needQuestions || [];
    let body = coverBlockXml(tender, doc.name);
    body += heading("Questions au soumissionnaire", 1);
    body += para("Le soumissionnaire est invité à répondre librement à chacune des questions ci-dessous, afin de démontrer sa compréhension du besoin et sa capacité de conseil.");
    questions.forEach((q, i) => {
      body += labelValueTable([["Identifiant", `Question n°${i + 1}`], ["Question", q.question], ["Réponse", ""]]);
    });
    packageDocx(body, `${tender.reference}_${slug(doc.name)}.docx`, tender);
    return;
  }
  if (TABLE_DOC_NAMES.includes(doc.name)) {
    const rows = doc.content?.rows || defaultPriceScheduleRows();
    const header = ["Référence", "Intitulé de la ligne", "Quantité", "Tarif unitaire (CHF)", "Tarif total (CHF)"];
    const out = []; const styles = []; const merges = []; const formulas = [];
    out.push([`SÉRIE DE PRIX — ${tender.reference} — ${tender.title || ""}`.trim()]);
    merges.push({ r1: 0, c1: 0, r2: 0, c2: header.length - 1 });
    styles.push({ row: 0, cols: Array.from({ length: header.length }, (_, i) => i), style: XLS_STYLE.title });
    out.push([]);
    out.push(header);
    styles.push({ row: out.length - 1, cols: Array.from({ length: header.length }, (_, i) => i), style: XLS_STYLE.tableHeader });
    const firstRow = out.length + 1;
    rows.forEach(r => {
      const rowNum = out.length + 1;
      out.push([r.ref || "", r.label || "", Number(r.quantity) || 0, r.unitPrice === "" || r.unitPrice == null ? "" : Number(r.unitPrice), null]);
      formulas.push({ ref: `E${rowNum}`, f: `C${rowNum}*D${rowNum}` });
      styles.push({ row: out.length - 1, cols: [0, 1, 2, 3, 4], style: XLS_STYLE.data });
    });
    const lastRow = out.length;
    out.push(["", "Total", "", "", rows.length ? null : 0]);
    styles.push({ row: out.length - 1, cols: [0, 1, 2, 3, 4], style: XLS_STYLE.total });
    if (rows.length) formulas.push({ ref: `E${out.length}`, f: `SUM(E${firstRow}:E${lastRow})` });
    out.push([]);
    out.push(["Remarque : le total est à reporter dans le document récapitulatif de l'offre. Toutes les positions doivent être complétées — aucun prix à 0 ne peut être accepté."]);
    styles.push({ row: out.length - 1, cols: [0], style: XLS_STYLE.subtitle });
    downloadXLSX(`${tender.reference}_${slug(doc.name)}.xlsx`, [{
      name: doc.name, cols: [16, 44, 10, 16, 16], rows: out, formulas, styles, merges, rowHeights: [22],
    }]);
    return;
  }
  if (AUTO_TABLE_DOC_NAMES.includes(doc.name)) {
    const requirements = tender.requirements.filter(r => r.includeInCDC !== false);
    downloadResponseSheetXLSX(`${tender.reference}_${slug(doc.name)}.xlsx`, doc.name, requirements);
    return;
  }
  const schema = (TEXT_SCHEMAS[doc.name] || genericSchema)(tender);
  const values = doc.content?.sections || {};

  if (doc.name === "Procédure AO") {
    packageDocx(buildProcedureAOBody(tender, schema, values), `${tender.reference}_${slug(doc.name)}.docx`, tender);
    return;
  }
  if (doc.name === "Cahier des charges") {
    packageDocx(buildCahierDesChargesBody(tender, schema, values), `${tender.reference}_${slug(doc.name)}.docx`, tender);
    return;
  }
  if (doc.name === K2_DOC_NAME) {
    packageDocx(buildK2Body(tender), `${tender.reference}_K2.docx`, tender);
    return;
  }

  let body = coverBlockXml(tender, doc.name);
  body += para("TABLE DES MATIÈRES", { align: "center", bold: true, size: 24, color: DOCX_NAVY, spacingAfter: 240 });
  schema.forEach(s => { body += para(s.label, { spacingAfter: 60 }); });
  body += pageBreak();
  body += metaAndHistoryXml(tender);
  schema.forEach(s => {
    const val = values[s.key] !== undefined ? values[s.key] : s.seed;
    body += heading(s.label, 1);
    body += para(val && val.trim() ? val : "(à compléter)");
  });
  body += glossaryXml();
  packageDocx(body, `${tender.reference}_${slug(doc.name)}.docx`, tender);
}

// Assemble le document "Procédure AO" au format opérationnel Achats : sections numérotées 1 à 16,
// tableaux réels pour tout ce qui est structuré (identification, documents, calendrier, checklist,
// critères, contacts), et placeholders explicites [À COMPLÉTER PAR ACHATS] / [À VALIDER PAR ACHATS]
// pour toute information non saisie — jamais d'invention.
// Assemble le Cahier des charges au format officiel (7 chapitres, structure retrouvée dans vos
// documents réels) : Introduction, Réponse du soumissionnaire (convention CO/PC/NC), Présentation générale,
// Exigences et spécifications (fiches par exigence), Planification, Exigences réglementaires et modalités
// contractuelles, Glossaire. Comme pour la Procédure AO : aucune donnée contractuelle n'est inventée.
function buildCahierDesChargesBody(tender, schema, values) {
  const v = (key) => { const s = schema.find(x => x.key === key); const val = values[key] !== undefined ? values[key] : s?.seed; return val && val.trim() ? val : A_COMPLETER; };
  const opt = tender.cdcOptions || {};
  const c = makeChapterBuilder();

  // 1. Introduction
  c.H("1. Introduction", 1, "b_cdc_1");
  c.H("1.1. Objet du document", 2, "b_cdc_11");
  c.raw(para(v("obj_doc")));
  c.raw(para("Le soumissionnaire doit procéder à la vérification de tous les documents et références qui lui sont transmis et signaler toute erreur, manque ou incompatibilité constaté."));

  // 2. Réponse du soumissionnaire (boilerplate constant)
  c.H("2. Réponse du soumissionnaire", 1, "b_cdc_2");
  c.H("2.1. Étapes à suivre par le soumissionnaire", 2, "b_cdc_21");
  c.H("2.1.1. Étape 1", 3, "b_cdc_211");
  c.raw(para("Pour chaque groupe d'exigences, le soumissionnaire décrit sous la rubrique « Réponse » la solution proposée, en démontrant explicitement la compréhension et la prise en compte de l'exigence. Les réserves éventuelles sont mentionnées clairement."));
  c.H("2.1.2. Étape 2", 3, "b_cdc_212");
  c.raw(para("Le soumissionnaire renseigne la rubrique « Conformité » en mentionnant CO, PC ou NC (voir convention ci-dessous). À défaut de précision, l'exigence est considérée comme non conforme. Pour une exigence essentielle, seule la mention CO est autorisée — toute autre mention peut entraîner l'élimination automatique de l'offre."));
  c.H("2.1.3. Étape 3", 3, "b_cdc_213");
  c.raw(para("Après avoir complété l'ensemble des rubriques, le soumissionnaire date et signe le document."));
  c.H("2.2. Convention de présentation des exigences", 2, "b_cdc_22");
  c.raw(para("Chaque exigence est présentée sous forme de fiche (référence, caractère essentiel, description, remarque, criticité, conformité, réponse). Les niveaux de conformité possibles sont :"));
  c.raw(conformiteLegendXml());

  // 3. Présentation générale
  c.H("3. Présentation générale", 1, "b_cdc_3");
  c.raw(para("Ce chapitre présente les principaux besoins et grandes caractéristiques du service à mettre en œuvre. En cas de contradiction avec les spécifications détaillées, ces dernières priment."));
  c.H("3.1. Situation actuelle", 2, "b_cdc_31");
  c.raw(para(v("situation")));
  c.H("3.2. Objectifs", 2, "b_cdc_32");
  c.raw(para(v("objectifs")));
  c.H("3.3. Principales parties prenantes", 2, "b_cdc_33");
  c.raw(para(v("parties_prenantes")));
  c.H("3.4. Périmètre de l'offre", 2, "b_cdc_34");
  c.raw(para(v("perimetre")));
  c.H("3.5. Évolutions probables", 2, "b_cdc_35");
  c.raw(para(v("evolutions")));
  if (opt.phases) {
    c.H("3.6. Phasage du projet", 2, "b_cdc_36");
    c.raw(para(v("phases")));
  }

  // 4. Exigences et spécifications (+ modules optionnels : licences, hébergement, support)
  c.H("4. Exigences et spécifications", 1, "b_cdc_4");
  c.raw(para("Chaque exigence ci-dessous doit être reprise et complétée par le soumissionnaire selon la convention définie au chapitre 2.2."));
  let sub = 0;
  if (opt.licences) { sub++; c.H(`4.${sub}. Licences`, 2, "b_cdc_4_lic"); c.raw(para(v("licences_desc"))); }
  if (opt.hebergement) { sub++; c.H(`4.${sub}. Mode d'hébergement de la solution`, 2, "b_cdc_4_heb"); c.raw(para(v("hebergement_desc"))); }
  if (opt.support) { sub++; c.H(`4.${sub}. Prestations de support et accompagnement technique`, 2, "b_cdc_4_sup"); c.raw(para(v("support_desc"))); }
  const byCat = tender.requirements.filter(r => r.includeInCDC !== false).reduce((acc, r) => { (acc[r.category] = acc[r.category] || []).push(r); return acc; }, {});
  if (Object.keys(byCat).length === 0) {
    sub++;
    c.H(`4.${sub}. Exigences détaillées`, 2, "b_cdc_4_req");
    c.raw(para(A_COMPLETER));
  } else {
    Object.entries(byCat).forEach(([cat, reqs]) => {
      sub++;
      c.H(`4.${sub}. ${cat}`, 2, `b_cdc_4_${slug(cat)}`);
      c.raw(para(`Le soumissionnaire répond à chacune des ${reqs.length} exigence(s) ci-dessous dans l'onglet « ${cat} » du cahier de réponses joint (colonnes Réponse et Justificatif / commentaire), en respectant la convention définie au chapitre 2.2.`, { italic: true, size: 18 }));
      reqs.forEach(r => { c.raw(requirementCardXml(r)); });
    });
  }

  // 5. Planification
  c.H("5. Planification", 1, "b_cdc_5");
  c.raw(docTable(["Étape", "Date"], [
    ["Lancement de l'AO", tender.dateLaunch || A_COMPLETER],
    ["Clôture de la réception des offres", tender.dateClose || A_COMPLETER],
    ["Décision d'adjudication", tender.dateDecision || A_COMPLETER],
    ["Début du contrat", A_COMPLETER],
  ]));

  // 6. Exigences réglementaires et modalités contractuelles (+ modules optionnels : contrat-cadre, montants)
  c.H("6. Exigences réglementaires et modalités contractuelles", 1, "b_cdc_6");
  c.H("6.1. Cadre légal", 2, "b_cdc_61");
  c.raw(para(v("cadre_legal")));
  c.H("6.2. Contraintes réglementaires du présent appel d'offre", 2, "b_cdc_62");
  c.raw(para(v("contraintes_reg")));
  c.H("6.3. Point de contact principal", 2, "b_cdc_63");
  c.raw(para(`Acheteur en charge : ${tender.buyer || A_COMPLETER}`));
  c.H("6.4. Modalités de paiement", 2, "b_cdc_64");
  c.raw(para(v("paiement")));

  if (opt.montants) {
    c.H("6.5. Montants", 2, "b_cdc_65m");
    c.raw(docTable(["Montant maximum", "Montant annuel maximum"], [[v("montant_max"), v("montant_annuel_max")]]));
  }

  if (opt.contratCadre) {
    c.H("6.6. Durée du contrat-cadre", 2, "b_cdc_66");
    c.raw(para(v("duree_contrat")));
    c.H("6.7. Mono/multi-adjudicataire", 2, "b_cdc_67");
    c.raw(para(v("mono_multi")));
    c.H("6.8. Absence d'exclusivité", 2, "b_cdc_68");
    c.raw(para("Au titre du présent contrat-cadre, l'adjudicataire ne bénéficie d'aucun droit d'exclusivité de réponse aux besoins de l'organisation en matière d'objet du présent marché. Par conséquent, l'organisation peut décider de lancer d'autres appels d'offres du même objet et de recourir à des tiers au contrat-cadre pour la réalisation de ces prestations. L'adjudicataire du présent contrat-cadre ne peut soulever aucune réclamation contre cela ni prétendre au versement d'une indemnité."));
    c.H("6.9. Changement d'adjudicataire", 2, "b_cdc_69");
    c.raw(para("Le contrat-cadre constitue un système fermé. Aucun nouveau prestataire ne peut entrer dans le contrat-cadre après son adjudication. Demeure réservé le changement de statut légal d'un prestataire, lié au droit des sociétés, notamment suite à une fusion, scission, transformation ou à un transfert de patrimoine du prestataire."));
    c.H("6.10. Changement de régime fiscal ou douanier", 2, "b_cdc_610");
    c.raw(para("En cas de changement de régime fiscal et/ou douanier, l'organisation se réserve le droit de résilier le contrat pour des raisons financières, sans que cela n'ouvre droit à des indemnités au profit de l'adjudicataire."));
  }

  c.H("6.11. Résiliation / Pénalités de retard", 2, "b_cdc_611");
  c.raw(para(v("resiliation")));

  c.raw(pageBreak());
  c.H("7. Glossaire & abréviations", 1, "b_cdc_7");
  c.raw(docTable(["Terme", "Définition"], GLOSSARY_TERMS));

  let out = coverBlockXml(tender, "Cahier des charges");
  out += tocBlockXml(c.toc());
  out += pageBreak();
  out += c.body();
  return out;
}

function buildProcedureAOBody(tender, schema, values) {
  const v = (key) => { const s = schema.find(x => x.key === key); const val = values[key] !== undefined ? values[key] : s?.seed; return val && val.trim() ? val : A_COMPLETER; };
  const c = makeChapterBuilder();

  c.H("1. Objet de la procédure", 1, "b_pao_1");
  c.raw(para(v("objet")));

  c.H("2. Identification du marché", 1, "b_pao_2");
  c.raw(docTable(["Champ", "Valeur"], [
    ["Référence AO", tender.reference],
    ["Intitulé", tender.title || A_COMPLETER],
    ["Service demandeur", tender.service || A_COMPLETER],
    ["Chef de projet", tender.responsibleMetier || A_COMPLETER],
    ["Acheteur", tender.buyer || A_COMPLETER],
    ["Type de marché", tender.type || A_COMPLETER],
    ["Montant estimatif", tender.budget ? chf(Number(tender.budget)) : A_COMPLETER],
    ["Lots / options", A_COMPLETER + " si applicable"],
  ]));

  c.H("3. Type de procédure", 1, "b_pao_3");
  c.raw(para(v("type_procedure")));

  c.H("4. Périmètre de la consultation", 1, "b_pao_4");
  c.raw(para(v("perimetre")));

  c.H("5. Documents du dossier d'appel d'offres", 1, "b_pao_5");
  c.raw(docTable(["Document", "Obligatoire"], tender.documents.length ? tender.documents.map(d => [d.name, d.mandatory ? "Oui" : "Non"]) : [[A_COMPLETER, ""]]));

  c.H("6. Calendrier de la consultation", 1, "b_pao_6");
  c.raw(docTable(["Étape", "Date", "Responsable"], [
    ["Publication / invitation", tender.dateLaunch || A_COMPLETER, "Achats"],
    ["Ouverture de la consultation", tender.dateLaunch || A_COMPLETER, "Achats"],
    ["Délai pour questions", A_COMPLETER, "Soumissionnaire"],
    ["Réponses aux questions", A_COMPLETER, "Organisation / Achats"],
    ["Date limite de remise des offres", tender.dateClose || A_COMPLETER, "Soumissionnaire"],
    ["Analyse de recevabilité", A_COMPLETER, "Achats"],
    ["Analyse technique", A_COMPLETER, "Métier / IT"],
    ["Analyse financière", A_COMPLETER, "Achats"],
    ["Décision d'adjudication", tender.dateDecision || A_COMPLETER, "Instance compétente"],
    ["Notification", A_COMPLETER, "Achats"],
    ["Début du marché", A_COMPLETER, "Organisation"],
  ]));

  c.H("7. Modalités de remise des offres", 1, "b_pao_7");
  c.raw(para(v("modalites")));

  c.H("8. Composition de l'offre du soumissionnaire", 1, "b_pao_8");
  const checklist = tender.documents.length ? tender.documents.map(d => `☐ ${d.name}${d.mandatory ? "" : " (si applicable)"}`).join("\n") : `☐ ${A_COMPLETER}`;
  c.raw(para(`Le soumissionnaire doit remettre :\n${checklist}`));

  c.H("9. Critères de recevabilité", 1, "b_pao_9");
  c.raw(para(v("recevabilite")));

  c.H("10. Critères d'évaluation / adjudication", 1, "b_pao_10");
  c.raw(docTable(["Critère", "Pondération", "Méthode d'évaluation"], tender.criteria.length ? tender.criteria.map(cr => [cr.name, `${cr.weight}%`, A_VALIDER]) : [["Prix", A_COMPLETER, A_COMPLETER], ["Qualité technique", A_COMPLETER, A_COMPLETER]]));

  c.H("11. Questions / réponses", 1, "b_pao_11");
  c.raw(para(v("qr")));

  c.H("12. Validité des offres", 1, "b_pao_12");
  c.raw(para(v("validite")));

  c.H("13. Confidentialité et communication", 1, "b_pao_13");
  c.raw(para(v("confidentialite")));

  c.H("14. Conditions particulières de la consultation", 1, "b_pao_14");
  c.raw(para(v("conditions_part")));

  c.H("15. Contacts", 1, "b_pao_15");
  c.raw(docTable(["Rôle", "Nom", "Email", "Téléphone"], [
    ["Acheteur", tender.buyer || A_COMPLETER, A_COMPLETER, A_COMPLETER],
    ["Chef de projet", tender.responsibleMetier || A_COMPLETER, A_COMPLETER, A_COMPLETER],
  ]));

  c.H("16. Historique / validation du document", 1, "b_pao_16");
  c.raw(docTable(["Version", "Date", "Raison"], [["0.1", new Date().toLocaleDateString("fr-CH"), "Création"]]));
  c.raw(docTable(["Rôle", "Nom", "Date", "Statut"], [["Chef de projet", tender.responsibleMetier || A_COMPLETER, "", "☐ Validé"], ["Achats", A_COMPLETER, "", "☐ Validé"]]));

  c.raw(pageBreak());
  c.H("Contrôle avant publication", 1, "b_pao_ctrl");
  c.raw(para([
    "☐ Type de procédure validé par Achats", "☐ Calendrier validé", "☐ Documents du dossier présents",
    "☐ Modalités de remise validées", "☐ Critères d'évaluation validés", "☐ Exigences éliminatoires vérifiées",
    "☐ Conditions contractuelles vérifiées", "☐ Contacts vérifiés", "☐ Cohérence CDC / Procédure vérifiée",
    "☐ Aucune information générée sans source",
  ].join("\n")));

  c.raw(pageBreak());
  c.H("Glossaire & abréviations", 1, "b_pao_gloss");
  c.raw(docTable(["Terme", "Définition"], GLOSSARY_TERMS));

  let out = coverBlockXml(tender, "Procédure AO");
  out += tocBlockXml(c.toc());
  out += pageBreak();
  out += c.body();
  return out;
}

// Assemble le formulaire K2 du Guide romand pour l'invitation à soumissionner et l'adjudication de
// marchés publics (Etat de Vaud et cantons romands, www.simap.ch) — "Dossier d'appel d'offres
// (procédure ouverte)". Structure et numérotation des chapitres calquées sur l'Annexe K2 officielle
// (11 chapitres), condensées pour un marché de services informatiques (les rubriques propres aux
// marchés de travaux de construction — codes CFC/CAN, organigramme de chantier, badge du personnel
// de chantier, entreprise générale/totale — sont omises). Comme pour les autres documents générés,
// aucune donnée contractuelle ou réglementaire propre à l'adjudicateur n'est inventée : ce qui n'est
// pas saisi dans l'AO reste affiché comme [À COMPLÉTER PAR ACHATS] / [À VALIDER PAR ACHATS].
function buildK2Body(tender) {
  const t = tender;
  const c = makeChapterBuilder();

  c.H("1. Pouvoir adjudicateur", 1, "b_k2_1");
  c.H("1.1. Nom et adresse de l'adjudicateur", 2, "b_k2_11");
  c.H("1.1.1. Pouvoir adjudicateur", 3, "b_k2_111");
  c.raw(docTable(["Champ", "Valeur"], [
    ["Direction / Service", [t.direction, t.service].filter(Boolean).join(" — ") || A_COMPLETER],
    ["Sponsor", t.sponsor || A_COMPLETER],
  ]));
  c.H("1.1.2. Organisateur de la procédure", 3, "b_k2_112");
  c.raw(docTable(["Champ", "Valeur"], [
    ["Acheteur en charge de la procédure", t.buyer || A_COMPLETER],
    ["Responsable métier", t.responsibleMetier || A_COMPLETER],
  ]));

  c.H("2. Objet du marché", 1, "b_k2_2");
  c.H("2.1. Objet et étendue du marché", 2, "b_k2_21");
  c.raw(para(t.object || A_COMPLETER));
  c.H("2.2. Titre du projet", 2, "b_k2_22");
  c.raw(para(t.title || A_COMPLETER));
  c.H("2.3. Référence/numéro du projet", 2, "b_k2_23");
  c.raw(para(t.reference));
  c.H("2.4. Vocabulaire commun des marchés publics (CPV)", 2, "b_k2_24");
  c.raw(para(A_COMPLETER));
  c.H("2.5. Pour les marchés de services — Classification centrale de produits (CPC)", 2, "b_k2_25");
  c.raw(para(A_COMPLETER));
  c.H("2.6. Description du marché", 2, "b_k2_26");
  c.H("2.6.1. Nature et importance", 3, "b_k2_261");
  c.raw(para(t.need?.context || A_COMPLETER));
  c.H("2.6.2. Étapes de réalisation", 3, "b_k2_262");
  c.raw(para(t.need?.objectives || A_COMPLETER));
  c.H("2.6.3. Lieu d'exécution du marché", 3, "b_k2_263");
  c.raw(para(A_COMPLETER));

  c.H("3. Calendrier prévisionnel de la procédure", 1, "b_k2_3");
  c.raw(docTable(["Étape", "Date"], [
    ["Publication de l'appel d'offres", t.dateLaunch || A_COMPLETER],
    ["Délai pour poser des questions", A_COMPLETER],
    ["Délai pour la remise des offres", t.dateClose || A_COMPLETER],
    ["Ouverture des offres", A_COMPLETER],
    ["Décision d'adjudication", t.dateDecision || A_COMPLETER],
    ["Délai de recours", A_COMPLETER],
    ["Début du marché", A_COMPLETER],
  ]));

  c.H("4. Bases légales", 1, "b_k2_4");
  c.raw(para("Le marché est :"));
  c.raw(para(`☐ soumis   ☐ non soumis   aux Accords internationaux (AMP 2012 et Accord bilatéral entre la Suisse et l'Union européenne) ;\n☐ soumis à l'Accord intercantonal sur les marchés publics (AIMP 2019) ;\n☐ soumis à la législation cantonale sur les marchés publics.`));

  c.H("5. Conditions de participation", 1, "b_k2_5");
  c.H("5.1. Rappel des obligations et contrôles", 2, "b_k2_51");
  c.raw(para("Pour participer à la procédure, le soumissionnaire (et ses sous-traitants éventuels) doit (doivent) respecter, conformément à l'art. 12 AIMP 2019, les dispositions relatives à la protection des travailleurs et aux conditions de travail en vigueur, les obligations d'annonce et d'autorisation relatives au travail au noir, les dispositions relatives à l'égalité de traitement salarial entre femmes et hommes, les conventions fondamentales de l'Organisation internationale du travail (OIT) en cas de prestations exécutées à l'étranger, ainsi que les dispositions légales en matière de protection de l'environnement. Le soumissionnaire (et ses sous-traitants éventuels) doit en outre être à jour dans le paiement des impôts et des cotisations sociales, et ne pas avoir conclu d'accords illicites affectant la concurrence (art. 26 al. 1 AIMP 2019)."));
  c.H("5.2. Peines conventionnelles", 2, "b_k2_52");
  c.raw(para(`L'adjudicateur inclura des peines conventionnelles dans le contrat à conclure avec le soumissionnaire retenu, pour le cas où ce dernier ou ses sous-traitants éventuels ne respecteraient pas leurs obligations : ☐ NON   ☐ OUI\nModalités (conditions, montants) : ${A_COMPLETER}`));

  c.H("6. Critères d'aptitude", 1, "b_k2_6");
  c.H("6.1. Critères d'aptitude, sous-critères et éléments d'appréciation", 2, "b_k2_61");
  c.raw(para("Les critères d'aptitude, énoncés dans le tableau ci-après, sont évalués de manière binaire (critère rempli ou non rempli). Seront admis à l'évaluation des critères d'adjudication les soumissionnaires ayant rempli tous les critères d'aptitude."));
  c.raw(docTable(["Critère d'aptitude", "Élément d'appréciation / justificatif demandé"], [
    ["Capacité économique et financière", A_COMPLETER],
    ["Capacité technique et professionnelle", A_COMPLETER],
    ["Références pour des prestations comparables", A_COMPLETER],
  ]));

  c.H("7. Critères d'adjudication", 1, "b_k2_7");
  c.H("7.1. Critères d'adjudication, sous-critères et éléments d'appréciation", 2, "b_k2_71");
  c.raw(para("Pour déterminer l'offre économiquement la plus avantageuse (art. 41 AIMP 2019), l'adjudicateur prend en compte au moins deux critères, dont le prix et la qualité de l'offre. Les critères sont les suivants :"));
  c.raw(docTable(["Critère", "Pondération", "Méthode de notation"], t.criteria.length
    ? t.criteria.map(cr => [cr.name, `${cr.weight}%`, A_VALIDER])
    : [["Prix", A_COMPLETER, A_COMPLETER], ["Qualité de l'offre", A_COMPLETER, A_COMPLETER]]));
  c.H("7.2. Échelle de notes", 2, "b_k2_72");
  c.raw(para("L'échelle de notes est de 0 à 5 (0 constituant la plus mauvaise note et 5 la meilleure note). Le critère prix est noté jusqu'au centième ; un critère ou sous-critère qualitatif est noté jusqu'à la demi-note."));
  c.H("7.3. Notation des critères d'adjudication", 2, "b_k2_73");
  c.H("7.3.1. Notation du critère qualité de l'offre", 3, "b_k2_731");
  c.raw(para(A_COMPLETER));
  c.H("7.3.2. Notation du critère prix", 3, "b_k2_732");
  c.raw(para(`La notation du prix (montant TTC) sera effectuée selon la méthode de notation suivante (cf. annexe T2 du Guide romand) : ${A_COMPLETER}`));
  c.H("7.4. Offres équivalentes", 2, "b_k2_74");
  c.raw(para("Si des offres obtiennent exactement le même nombre de points, l'adjudicateur, pour les départager, favorisera le soumissionnaire ayant acquis la meilleure note sur le critère le plus fortement pondéré, et ainsi de suite de critère en critère du plus important au moins important."));

  c.H("8. Exigences pour participer à la procédure d'adjudication", 1, "b_k2_8");
  c.H("8.1. Délai pour la remise des offres", 2, "b_k2_81");
  c.raw(para(`L'offre doit parvenir au plus tard le ${t.dateClose || A_COMPLETER}, à l'adresse ou via la plateforme indiquée par l'adjudicateur. Il appartient au soumissionnaire de tout mettre en œuvre pour respecter cette échéance ; les offres remises hors délai seront exclues de la procédure.`));
  c.H("8.2. Présentation de l'offre", 2, "b_k2_82");
  c.raw(para(`Le soumissionnaire doit déposer son offre : ☐ au format papier   ☐ au format électronique via la plateforme Simap.ch   ☐ au choix.\nNombre d'exemplaires (format papier) : ${A_COMPLETER}`));
  c.H("8.3. Conditions de recevabilité de l'offre", 2, "b_k2_83");
  c.raw(para(A_COMPLETER));
  c.H("8.4. Langue", 2, "b_k2_84");
  c.raw(para(A_VALIDER));
  c.H("8.5. Confidentialité et propriété des documents et informations", 2, "b_k2_85");
  c.raw(para("Les documents remis par l'adjudicateur et par le soumissionnaire dans le cadre de la présente procédure sont confidentiels et demeurent la propriété de leur auteur respectif ; ils ne peuvent être ni reproduits, ni communiqués à des tiers, sans autorisation préalable."));
  c.H("8.6. Durée de validité de l'offre", 2, "b_k2_86");
  c.raw(para(A_COMPLETER));
  c.H("8.7. Taxe sur la valeur ajoutée", 2, "b_k2_87");
  c.raw(para(A_VALIDER));

  c.H("9. Procédure d'adjudication", 1, "b_k2_9");
  c.H("9.1. Délai pour poser des questions", 2, "b_k2_91");
  c.raw(para(A_COMPLETER));
  c.H("9.2. Ouverture des offres", 2, "b_k2_92");
  c.raw(para(A_COMPLETER));
  c.H("9.3. Examen et évaluation des offres", 2, "b_k2_93");
  c.raw(para("L'adjudicateur examine la recevabilité des offres, puis évalue les critères d'aptitude et les critères d'adjudication définis ci-avant."));
  c.H("9.4. Décision d'adjudication", 2, "b_k2_94");
  c.raw(para(`La décision d'adjudication est notifiée à l'ensemble des soumissionnaires. Date prévisionnelle : ${t.dateDecision || A_COMPLETER}.`));
  c.H("9.5. Voies de recours", 2, "b_k2_95");
  c.raw(para("La décision d'adjudication peut faire l'objet d'un recours dans le délai et selon les modalités indiquées dans la décision elle-même, conformément à la législation cantonale applicable."));

  c.H("10. Engagement du soumissionnaire quant à la procédure", 1, "b_k2_10");
  c.raw(para("En déposant son offre, le soumissionnaire s'engage à respecter les règles de la présente procédure et atteste l'exactitude des informations fournies."));

  c.H("11. Engagement de l'adjudicateur quant à la procédure", 1, "b_k2_11e");
  c.raw(para("L'adjudicateur s'engage à conduire la présente procédure conformément aux principes de transparence, d'égalité de traitement et de non-discrimination entre les soumissionnaires."));

  c.raw(pageBreak());
  c.H("Annexes à compléter", 1, "b_k2_annexes");
  c.raw(para("Les annexes cochées ci-après font partie intégrante des documents d'appel d'offres et doivent être retournées complétées à l'adjudicateur dans le même délai que celui fixé pour la remise des offres."));
  c.raw(docTable(["Annexe", "Objet"], [
    ["P1", "Engagement sur l'honneur"],
    ["P4", "Caractéristiques du soumissionnaire"],
    ["P5", "Assurances et garanties"],
    ["P6", "Engagement à respecter l'égalité entre femmes et hommes"],
    ["P7", "Respect des conditions de travail internationales"],
    ["P8", "Engagement sur l'honneur du sous-traitant"],
    ["Q1", "Organisation qualité du soumissionnaire"],
    ["Q2", "Organigramme structurel du soumissionnaire"],
    ["Q4", "Capacité en personnel et formation des personnes-clés"],
    ["Q7", "Liste de références de services (hors construction)"],
    ["R1", "Montant de l'offre en rapport avec le cahier des charges"],
    ["R6", "Moyens et ressources pour l'exécution du marché"],
    ["R9", "Qualifications des personnes-clés désignées"],
    ["R13", "Qualités et adéquation des solutions techniques proposées"],
    ["R14", "Degré de compréhension du cahier des charges"],
  ]));
  c.raw(para("Documents remis à chaque soumissionnaire : le cahier des charges du marché, la série de prix, les conditions générales du contrat.\nAutres informations : www.simap.ch"));

  let out = coverBlockXml(t, "K2 — Dossier d'appel d'offres (procédure ouverte)");
  out += tocBlockXml(c.toc());
  out += pageBreak();
  out += c.body();
  return out;
}

function tocBlockXml(entries) {
  let x = para("TABLE DES MATIÈRES", { align: "center", bold: true, size: 24, color: DOCX_NAVY, spacingAfter: 240 });
  entries.forEach(e => { x += tocLink(e.text, e.id, e.level); });
  return x;
}

function slug(s) { return s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, ""); }

/* --------------------------- ASSISTANT IA (aide à la rédaction) --------------------------- */
// Fonction centrale utilisée par tous les boutons "Aide IA" de l'app.
// Renvoie l'objet JSON demandé dans le prompt, ou lève une erreur explicite.
// Appelle la fonction Supabase Edge "ai-assist", qui relaie vers l'API Anthropic
// côté serveur (la clé API n'est jamais exposée au navigateur).
async function callAssist(prompt, maxTokens = 1500) {
  const { data, error } = await supabase.functions.invoke("ai-assist", { body: { prompt, maxTokens } });
  if (error) {
    const serverMessage = error.context?.body ? await error.context.json?.().catch(() => null) : null;
    throw new Error(serverMessage?.error || error.message || "Connexion à l'assistant impossible. Réessayez, ou complétez manuellement.");
  }
  if (!data?.text) throw new Error("Réponse vide de l'assistant.");
  return data.text;
}

// Extrait le premier objet/tableau JSON valide et équilibré d'un texte, en ignorant les
// crochets isolés qui peuvent apparaître dans du texte libre autour (ex. placeholders du type
// "[À COMPLÉTER PAR ACHATS]"). Une simple regex gourmande s'accroche sur ces faux positifs.
function extractJSON(text) {
  for (let i = 0; i < text.length; i++) {
    const open = text[i];
    if (open !== "{" && open !== "[") continue;
    const close = open === "{" ? "}" : "]";
    let depth = 0, inStr = false, escape = false;
    for (let j = i; j < text.length; j++) {
      const c = text[j];
      if (inStr) {
        if (escape) escape = false;
        else if (c === "\\") escape = true;
        else if (c === '"') inStr = false;
        continue;
      }
      if (c === '"') { inStr = true; continue; }
      if (c === open) depth++;
      else if (c === close) {
        depth--;
        if (depth === 0) {
          try { return JSON.parse(text.slice(i, j + 1)); } catch { break; }
        }
      }
    }
  }
  throw new Error("Réponse de l'assistant illisible.");
}

async function callClaudeJSON(prompt, maxTokens = 1500) {
  const text = await callAssist(prompt, maxTokens);
  try {
    return extractJSON(text);
  } catch (e) {
    if (!/[}\]]\s*$/.test(text.trim())) {
      throw new Error("La réponse de l'assistant a été coupée (trop longue). Réessayez, ou réduisez la portée de la demande.");
    }
    throw e;
  }
}

// Variante texte libre — pour les questions ouvertes ("quels sont les critères habituels pour ce type de besoin ?").
// Utilise les connaissances générales du modèle : pas de recherche web en direct, donc pas de "scraping" réel —
// c'est un raisonnement basé sur ce que le modèle sait du marché.
async function callClaudeText(prompt, maxTokens = 800) {
  return (await callAssist(prompt, maxTokens)).trim();
}

/* --------------------------------- COMPOSANTS UI GÉNÉRIQUES --------------------------------- */

function Badge({ children, color, bg }) {
  return <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-xs font-medium" style={{ color, backgroundColor: bg }}>{children}</span>;
}
function StatusBadge({ status }) { const m = STATUS_META[status] || STATUS_META.draft; return <Badge color={m.color} bg={m.bg}>{m.label}</Badge>; }
function ProgressBar({ value, dark }) {
  return <div className="w-full h-1.5 rounded-full transition-all" style={{ backgroundColor: dark ? "rgba(255,255,255,0.14)" : C.borderSoft }}>
    <div className="h-1.5 rounded-full transition-all duration-500" style={{ width: `${value}%`, backgroundColor: value >= 80 ? C.green : value >= 40 ? C.accent : C.amber }} />
  </div>;
}
function Card({ children, className = "", style = {} }) {
  return <div className={`rounded-xl transition-shadow ${className}`} style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, boxShadow: "0 1px 2px rgba(24,34,52,0.04)", ...style }}>{children}</div>;
}
function SectionTitle({ children, sub }) {
  return <div className="mb-4">
    <h2 className="text-[15px] font-semibold" style={{ color: C.ink }}>{children}</h2>
    {sub && <p className="text-sm mt-0.5" style={{ color: C.inkSoft }}>{sub}</p>}
  </div>;
}
function GhostButton({ children, onClick, disabled, icon: Icon }) {
  return <button onClick={onClick} disabled={disabled}
    className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-full text-xs font-medium disabled:opacity-40 transition-colors hover:bg-black/[0.03]"
    style={{ border: `1px solid ${C.border}`, color: C.inkSoft, backgroundColor: C.surface }}>
    {Icon && <Icon size={13} />}{children}
  </button>;
}
function PrimaryButton({ children, onClick, disabled, icon: Icon }) {
  return <button onClick={onClick} disabled={disabled}
    className="flex items-center gap-2 px-5 py-2.5 rounded-full text-sm font-semibold text-white disabled:opacity-40 transition-transform active:scale-[0.98] hover:brightness-110"
    style={{ backgroundColor: C.accent }}>
    {Icon && <Icon size={14} />}{children}
  </button>;
}

// Widget "poser une question librement à l'assistant" — utilisable dans plusieurs onglets.
// Répond à partir des connaissances générales du modèle (pas de scraping web en direct depuis le navigateur).
// Widget "poser une question librement à l'assistant" — les suggestions retournées sont structurées
// et intégrables en un clic dans les bons champs de l'AO (exigence, critère, périmètre, contrainte).
function AIQuestionBox({ tender, updateTender, placeholder, contextNote }) {
  const [question, setQuestion] = useState("");
  const [summary, setSummary] = useState("");
  const [items, setItems] = useState([]);
  const [added, setAdded] = useState({});
  const [asking, setAsking] = useState(false);
  const [error, setError] = useState("");

  async function ask() {
    if (!question.trim()) return;
    setAsking(true); setError(""); setSummary(""); setItems([]); setAdded({});
    try {
      const context = `Titre de l'AO : ${tender.title}\nType : ${tender.type || "non précisé"}\nObjet : ${tender.object || "non précisé"}\nContexte : ${tender.need?.context || "non précisé"}\nExigences déjà définies : ${(tender.requirements || []).map(r => r.description).join(" | ") || "aucune"}\nCritères déjà définis : ${(tender.criteria || []).map(c => `${c.name} (${c.weight}%)`).join(", ") || "aucun"}`;
      const prompt = `Tu es un assistant Achats qui aide un chef de projet à préparer un appel d'offres IT. ` +
        `Réponds à la question ci-dessous en te basant sur tes connaissances générales du marché et des pratiques usuelles d'achat public/IT (pas une recherche en temps réel — précise-le dans le résumé). N'invente aucune règle juridique ou procédure interne spécifique. ` +
        `Réponds UNIQUEMENT avec un objet JSON valide, sans balises markdown, de la forme : ` +
        `{"summary": "résumé en 1-2 phrases de ta réponse", "items": [{"type": "exigence"|"critere"|"inclus"|"exclus"|"contrainte", "label": "titre court de la suggestion", "detail": "explication ou justification en 1 phrase", "criticality": "Bloquante"|"Critique"|"Majeure"|"Normale"|"Souhaitable" (uniquement si type=exigence), "weight": nombre (uniquement si type=critere)}]}. ` +
        `Propose entre 3 et 8 items concrets et actionnables, directement liés à la question.\n\n` +
        `Contexte de l'AO :\n${context}\n\nQuestion du chef de projet : ${question}`;
      const res = await callClaudeJSON(prompt, 2000);
      setSummary(res.summary || "");
      setItems(Array.isArray(res.items) ? res.items : []);
    } catch (e) {
      setError(`La question a échoué (${e.message || "erreur inconnue"}).`);
    } finally { setAsking(false); }
  }

  function integrate(item, idx) {
    if (item.type === "exigence") {
      const prefix = "GEN";
      updateTender(prev => {
        const n = prev.requirements.filter(r => r.id.startsWith(prefix)).length + 1;
        const newReq = { id: `${prefix}-${String(n).padStart(3, "0")}`, category: "Général", description: item.label + (item.detail ? ` — ${item.detail}` : ""), criticality: item.criticality || "Normale", mandatory: item.criticality === "Bloquante", verificationMethod: "" };
        return { ...prev, requirements: [...prev.requirements, newReq] };
      });
    } else if (item.type === "critere") {
      updateTender(prev => ({ ...prev, criteria: [...prev.criteria, { id: `c${Date.now()}_${idx}`, name: item.label, weight: Number(item.weight) || 0 }] }));
    } else if (item.type === "inclus" || item.type === "exclus") {
      const key = item.type === "inclus" ? "included" : "excluded";
      updateTender(prev => ({ ...prev, scope: { ...prev.scope, [key]: [...(prev.scope[key] || []), item.label] } }));
    } else if (item.type === "contrainte") {
      updateTender(prev => ({ ...prev, need: { ...prev.need, constraints: [prev.need.constraints, item.label].filter(Boolean).join("\n") } }));
    }
    setAdded(a => ({ ...a, [idx]: true }));
  }

  const TYPE_LABEL = { exigence: "Exigence", critere: "Critère", inclus: "Périmètre (inclus)", exclus: "Périmètre (exclus)", contrainte: "Contrainte" };

  return (
    <Card className="p-5">
      <div className="flex items-center gap-2 mb-1">
        <Sparkles size={15} style={{ color: C.accentDark }} />
        <div className="text-sm font-semibold" style={{ color: C.ink }}>Poser une question à l'assistant</div>
      </div>
      {contextNote && <div className="text-xs mb-3" style={{ color: C.inkSoft }}>{contextNote}</div>}
      <div className="flex flex-col sm:flex-row gap-2">
        <input value={question} onChange={e => setQuestion(e.target.value)} onKeyDown={e => e.key === "Enter" && ask()}
          placeholder={placeholder || "Ex. Quels sont les critères d'évaluation habituels pour ce type de besoin ?"}
          className="flex-1 px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} />
        <button onClick={ask} disabled={asking || !question.trim()} className="flex items-center justify-center gap-1.5 px-4 py-2 rounded-full text-sm font-medium disabled:opacity-40 shrink-0" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
          {asking ? "Recherche…" : "Demander"}
        </button>
      </div>
      {error && <div className="text-xs mt-2" style={{ color: C.red }}>{error}</div>}
      {summary && <div className="mt-3 text-sm" style={{ color: C.ink }}>{summary}</div>}
      {items.length > 0 && (
        <div className="mt-3 space-y-2">
          {items.map((item, idx) => (
            <div key={idx} className="flex items-start gap-3 p-3 rounded" style={{ border: `1px solid ${C.borderSoft}` }}>
              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2 flex-wrap">
                  <Badge color={C.accentDark} bg={C.accentSoft}>{TYPE_LABEL[item.type] || item.type}</Badge>
                  {item.type === "critere" && item.weight != null && <span className="text-xs" style={{ color: C.inkSoft }}>{item.weight}%</span>}
                  {item.type === "exigence" && item.criticality && <Badge color={CRITICALITY_META[item.criticality] || C.inkSoft} bg={C.slateSoft}>{item.criticality}</Badge>}
                </div>
                <div className="text-sm mt-1" style={{ color: C.ink }}>{item.label}</div>
                {item.detail && <div className="text-xs mt-0.5" style={{ color: C.inkSoft }}>{item.detail}</div>}
              </div>
              {added[idx] ? (
                <span className="flex items-center gap-1 text-xs shrink-0" style={{ color: C.green }}><CheckCircle2 size={13} /> Ajouté</span>
              ) : (
                <button onClick={() => integrate(item, idx)} className="flex items-center gap-1 px-3 py-1.5 rounded-full text-xs font-medium shrink-0 text-white" style={{ backgroundColor: C.accent }}>
                  <Plus size={12} /> Intégrer
                </button>
              )}
            </div>
          ))}
        </div>
      )}
      <div className="text-xs mt-3" style={{ color: C.inkSoft }}>Basé sur les connaissances générales du modèle, pas une recherche web en direct — relisez avant d'intégrer.</div>
    </Card>
  );
}

/* --------------------------------- SIDEBAR / SHELL --------------------------------- */

function useIsMobile() {
  const [isMobile, setIsMobile] = useState(() => typeof window !== "undefined" && window.innerWidth < 860);
  useEffect(() => {
    function onResize() { setIsMobile(window.innerWidth < 860); }
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);
  return isMobile;
}

function Sidebar({ view, setView, isMobile, open, onClose, onLogout, userEmail, isAdmin }) {
  const navItems = [
    { key: "dashboard", label: "Tableau de bord", icon: LayoutDashboard },
    { key: "new", label: "Nouvel appel d'offres", icon: FilePlus2 },
    { key: "suppliers-registry", label: "Fournisseurs (registre)", icon: Building2 },
    ...(isAdmin ? [{ key: "admin", label: "Administration", icon: ShieldCheck }] : []),
  ];
  const soon = [
    { label: "Modèles d'AO", icon: ListChecks },
    ...(isAdmin ? [] : [{ label: "Administration", icon: ShieldCheck }]),
  ];
  function pick(key) { setView(key); if (isMobile) onClose(); }

  const content = (
    <>
      <div className="px-5 pt-6 pb-5 flex items-start justify-between" style={{ borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
        <div>
          <img src={logoWordmarkDark} alt="AO Manager" className="h-7 w-auto mb-2" />
          <div className="text-xs" style={{ color: "#8C97AC" }}>Cockpit des appels d'offres</div>
        </div>
        {isMobile && <button onClick={onClose} className="text-white/70 p-1"><X size={18} /></button>}
      </div>
      <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
        {navItems.map(it => {
          const Icon = it.icon; const active = view === it.key;
          return <button key={it.key} onClick={() => pick(it.key)}
            className="w-full flex items-center gap-2.5 px-3 py-2.5 rounded text-sm text-left transition-all duration-150 hover:bg-white/[0.06]"
            style={{ backgroundColor: active ? "rgba(255,255,255,0.1)" : "transparent", color: active ? "#FFFFFF" : "#B7C0D1", boxShadow: active ? "inset 2px 0 0 0 " + C.accent : "none" }}>
            <Icon size={16} strokeWidth={2} />{it.label}
          </button>;
        })}
        <div className="pt-4 mt-4" style={{ borderTop: "1px solid rgba(255,255,255,0.08)" }}>
          <div className="px-3 pb-2 text-[11px] uppercase tracking-wider" style={{ color: "#66738A" }}>À venir</div>
          {soon.map((it, i) => { const Icon = it.icon; return <div key={i} className="w-full flex items-center gap-2.5 px-3 py-2 rounded text-sm" style={{ color: "#5C6880" }}><Icon size={16} strokeWidth={2} />{it.label}</div>; })}
        </div>
      </nav>
      <div className="px-5 py-4 text-xs space-y-1.5" style={{ color: "#66738A", borderTop: "1px solid rgba(255,255,255,0.08)" }}>
        {userEmail && <div className="truncate" title={userEmail}>{userEmail}</div>}
        <div>Vos AO sont sauvegardés automatiquement.</div>
        <button onClick={onLogout} className="underline hover:text-white/80 block">Se déconnecter</button>
      </div>
    </>
  );

  if (isMobile) {
    return (
      <>
        {open && <div className="fixed inset-0 z-40" style={{ backgroundColor: "rgba(0,0,0,0.4)" }} onClick={onClose} />}
        <aside className="fixed top-0 left-0 h-screen w-72 z-50 flex flex-col transition-transform duration-200"
          style={{ backgroundColor: C.ink, transform: open ? "translateX(0)" : "translateX(-100%)" }}>
          {content}
        </aside>
      </>
    );
  }
  return (
    <aside className="w-64 shrink-0 h-screen sticky top-0 hidden md:flex flex-col" style={{ backgroundColor: C.ink }}>
      {content}
    </aside>
  );
}

function initialsFor(email) {
  if (!email) return "?";
  const local = email.split("@")[0];
  const parts = local.split(/[._-]+/).filter(Boolean);
  const chars = parts.length >= 2 ? [parts[0][0], parts[1][0]] : [local[0], local[1] || ""];
  return chars.join("").toUpperCase();
}

function TopBar({ crumbs, onMenuClick, isMobile, userEmail }) {
  return (
    <div className="flex items-center justify-between px-4 sm:px-8 py-4" style={{ borderBottom: `1px solid ${C.border}`, backgroundColor: C.surface }}>
      <div className="flex items-center gap-3 min-w-0">
        {isMobile && <button onClick={onMenuClick} className="p-1 shrink-0" style={{ color: C.ink }}><Menu size={20} /></button>}
        <div className="flex items-center gap-1.5 text-sm min-w-0 overflow-hidden">
          {crumbs.map((c, i) => <span key={i} className="flex items-center gap-1.5 truncate">{i > 0 && <ChevronRight size={13} className="shrink-0" />}<span className="truncate" style={{ color: i === crumbs.length - 1 ? C.ink : C.inkSoft, fontWeight: i === crumbs.length - 1 ? 600 : 400 }}>{c}</span></span>)}
        </div>
      </div>
      <div className="flex items-center gap-4">
        <Search size={17} style={{ color: C.inkSoft }} />
        <Bell size={17} style={{ color: C.inkSoft }} />
        <div title={userEmail} className="w-7 h-7 rounded-full flex items-center justify-center text-xs font-semibold text-white shrink-0" style={{ background: `linear-gradient(135deg, ${C.accent}, ${C.accentDark})` }}>{initialsFor(userEmail)}</div>
      </div>
    </div>
  );
}

/* --------------------------------- DASHBOARD --------------------------------- */

const STATUS_ICON = { draft: PenLine, preparation: ListChecks, validation: ShieldCheck, ongoing: Users, evaluation: BarChart3, completed: CheckCircle2 };

function greetingName(email) {
  if (!email) return "";
  const local = email.split("@")[0].split(/[._-]+/)[0];
  return local ? local[0].toUpperCase() + local.slice(1) : "";
}

function Dashboard({ tenders, openTender, goNew, onDelete, userEmail, firstName, isAdmin }) {
  const [statusFilter, setStatusFilter] = useState(null);
  const greeting = firstName || greetingName(userEmail);
  const counts = {}; Object.keys(STATUS_META).forEach(k => counts[k] = tenders.filter(t => t.status === k).length);
  const visibleTenders = statusFilter ? tenders.filter(t => t.status === statusFilter) : tenders;
  const actions = [];
  tenders.forEach(t => {
    const missing = missingMandatoryDocs(t);
    if (missing.length) actions.push({ ref: t.reference, text: `${missing.length} document${missing.length > 1 ? "s" : ""} manquant${missing.length > 1 ? "s" : ""}`, id: t.id });
    if (t.status === "evaluation" && t.suppliers.some(s => Object.keys(s.evaluations).length < t.criteria.length - 1)) actions.push({ ref: t.reference, text: "Votre évaluation est attendue", id: t.id });
    if (t.status === "preparation" && t.criteria.length === 0) actions.push({ ref: t.reference, text: "Critères d'évaluation à définir", id: t.id });
  });
  const today = new Date().toLocaleDateString("fr-CH", { weekday: "long", day: "numeric", month: "long" });
  const activeCount = tenders.filter(t => !["draft", "completed"].includes(t.status)).length;

  return (
    <div className="px-4 sm:px-8 py-5 sm:py-7 max-w-6xl">
      <div className="rounded-2xl px-6 sm:px-8 py-7 mb-8 relative overflow-hidden ao-scale-in" style={{ background: `linear-gradient(135deg, ${C.ink}, #263454)` }}>
        <div className="absolute -right-10 -top-16 w-56 h-56 rounded-full opacity-20" style={{ background: `radial-gradient(circle, ${C.accent}, transparent 70%)` }} />
        <div className="absolute right-24 -bottom-20 w-40 h-40 rounded-full opacity-10" style={{ background: `radial-gradient(circle, ${C.accent}, transparent 70%)` }} />
        <div className="relative flex flex-col sm:flex-row sm:items-end justify-between gap-4">
          <div>
            <div className="text-xs uppercase tracking-wider mb-1" style={{ color: "#8C97AC" }}>{today}</div>
            <h1 className="text-2xl font-semibold text-white">{greeting ? `Bonjour, ${greeting}` : "Bonjour"}</h1>
            <p className="text-sm mt-1.5" style={{ color: "#B7C0D1" }}>{activeCount > 0 ? `${activeCount} appel${activeCount > 1 ? "s" : ""} d'offres actif${activeCount > 1 ? "s" : ""} en ce moment.` : "Aucun AO actif pour l'instant — lancez-en un nouveau."}</p>
          </div>
          <PrimaryButton onClick={goNew} icon={FilePlus2}>Nouvel appel d'offres</PrimaryButton>
        </div>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 mb-8 ao-stagger">
        {Object.entries(STATUS_META).map(([key, meta]) => {
          const Icon = STATUS_ICON[key] || Circle;
          const active = statusFilter === key;
          return (
            <button key={key} onClick={() => setStatusFilter(f => f === key ? null : key)} className="text-left">
              <Card className="p-4 hover:shadow-md hover:-translate-y-0.5 transition-all cursor-pointer" style={active ? { border: `1.5px solid ${meta.color}`, boxShadow: `0 0 0 3px ${meta.bg}` } : {}}>
                <div className="w-8 h-8 rounded-lg flex items-center justify-center mb-2.5" style={{ backgroundColor: meta.bg }}>
                  <Icon size={15} style={{ color: meta.color }} />
                </div>
                <div className="text-2xl font-semibold tabular-nums" style={{ color: C.ink }}>{counts[key]}</div>
                <div className="text-xs mt-0.5" style={{ color: meta.color }}>{meta.label}</div>
              </Card>
            </button>
          );
        })}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="col-span-2">
          <div className="flex items-center justify-between gap-3 mb-4">
            <h2 className="text-[15px] font-semibold" style={{ color: C.ink }}>{statusFilter ? `AO — ${STATUS_META[statusFilter].label}` : "Mes appels d'offres"}</h2>
            {statusFilter && <button onClick={() => setStatusFilter(null)} className="text-xs font-medium flex items-center gap-1 shrink-0" style={{ color: C.accentDark }}><X size={12} /> Effacer le filtre</button>}
          </div>
          {tenders.length === 0 ? (
            <Card className="px-6 py-14 flex flex-col items-center text-center ao-scale-in">
              <div className="w-12 h-12 rounded-full flex items-center justify-center mb-4" style={{ backgroundColor: C.accentSoft }}>
                <FilePlus2 size={20} style={{ color: C.accentDark }} />
              </div>
              <div className="text-sm font-semibold" style={{ color: C.ink }}>Aucun appel d'offres pour l'instant</div>
              <p className="text-sm mt-1.5 max-w-sm" style={{ color: C.inkSoft }}>Créez votre premier AO — l'assistant vous guide pas à pas, de la définition du besoin jusqu'à la synthèse finale.</p>
              <div className="mt-5"><PrimaryButton onClick={goNew} icon={FilePlus2}>Créer mon premier AO</PrimaryButton></div>
            </Card>
          ) : visibleTenders.length === 0 ? (
            <Card className="px-6 py-10 text-center text-sm ao-scale-in" style={{ color: C.inkSoft }}>Aucun AO avec ce statut.</Card>
          ) : (
            <Card className="ao-stagger">
              {visibleTenders.map((t, i) => {
                const progress = computeProgress(t);
                return (
                  <div key={t.id} role="button" tabIndex={0} onClick={() => openTender(t.id)} onKeyDown={e => e.key === "Enter" && openTender(t.id)}
                    className="w-full text-left px-5 py-4 flex items-center gap-4 transition-colors hover:bg-black/[0.02] cursor-pointer group"
                    style={{ borderBottom: i < visibleTenders.length - 1 ? `1px solid ${C.borderSoft}` : "none" }}>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2">
                        <span className="text-xs font-mono" style={{ color: C.inkSoft }}>{t.reference}</span>
                        <StatusBadge status={t.status} />
                      </div>
                      <div className="text-sm font-medium mt-1" style={{ color: C.ink }}>{t.title}</div>
                      <div className="mt-2 w-full max-w-xs"><ProgressBar value={progress} /></div>
                    </div>
                    <div className="text-sm font-medium tabular-nums" style={{ color: C.inkSoft }}>{progress}%</div>
                    {isAdmin && (
                      <button onClick={e => { e.stopPropagation(); onDelete(t); }} title="Supprimer cet AO (admin)" className="p-1.5 rounded hover:bg-black/5 transition-colors">
                        <Trash2 size={15} style={{ color: C.inkSoft }} />
                      </button>
                    )}
                    <ChevronRight size={16} className="transition-transform group-hover:translate-x-0.5" style={{ color: C.inkSoft }} />
                  </div>
                );
              })}
            </Card>
          )}
        </div>
        <div>
          <SectionTitle>Actions requises</SectionTitle>
          <Card>
            {actions.length === 0 && (
              <div className="px-5 py-8 flex flex-col items-center text-center">
                <CheckCircle2 size={20} style={{ color: C.green }} className="mb-2" />
                <div className="text-sm" style={{ color: C.inkSoft }}>Aucune action en attente.</div>
              </div>
            )}
            {actions.map((a, i) => (
              <button key={i} onClick={() => openTender(a.id)} className="w-full text-left px-5 py-3.5 flex items-start gap-3 transition-colors hover:bg-black/[0.02]" style={{ borderBottom: i < actions.length - 1 ? `1px solid ${C.borderSoft}` : "none" }}>
                <AlertTriangle size={15} style={{ color: C.amber, marginTop: 2 }} />
                <div><div className="text-xs font-mono" style={{ color: C.inkSoft }}>{a.ref}</div><div className="text-sm" style={{ color: C.ink }}>{a.text}</div></div>
              </button>
            ))}
          </Card>
        </div>
      </div>
    </div>
  );
}

/* --------------------------------- NOUVEL AO (wizard) --------------------------------- */

function NewTenderWizard({ onCreate, onCancel }) {
  const [step, setStep] = useState(1);
  const [form, setForm] = useState({
    reference: "AO-2026-005", title: "", object: "", direction: "", service: "",
    responsibleMetier: "", buyer: "", sponsor: "", type: "", category: "IT",
    procedure: "Procédure ouverte", confidentiality: "Interne", budget: "",
    natureMarche: "Services",
    dateLaunch: "", dateClose: "", dateDecision: "",
  });
  const [rawNeed, setRawNeed] = useState("");
  const [structured, setStructured] = useState(null);
  const [loadingAI, setLoadingAI] = useState(false);
  const [aiError, setAiError] = useState("");

  function set(field, value) { setForm(f => ({ ...f, [field]: value })); }

  const step1Missing = ["title", "object", "direction", "responsibleMetier", "buyer"].filter(f => !form[f]?.trim());
  const step1Valid = step1Missing.length === 0;

  async function structureNeed() {
    if (!rawNeed.trim()) return;
    setLoadingAI(true); setAiError(""); setStructured(null);
    try {
      const prompt = `Tu es un assistant Achats. ` +
        `À partir de la description libre d'un besoin d'achat ci-dessous, produis UNIQUEMENT un objet JSON valide, sans balises markdown ni texte autour, avec exactement ces clés : ` +
        `"contexte" (string), "problematique" (string), "objectifs" (array de strings courtes), "perimetre" (array de strings courtes), ` +
        `"exigences_potentielles" (array de strings courtes), "contraintes" (array de strings courtes), "risques" (array de strings courtes), "indicateurs_reussite" (array de strings courtes). ` +
        `Reste générique et n'invente aucune règle juridique ou procédure interne spécifique. Description du besoin : """${rawNeed}"""`;
      setStructured(await callClaudeJSON(prompt, 3000));
    } catch (e) {
      setAiError(`La structuration automatique a échoué (${e.message || "erreur inconnue"}). Vous pouvez continuer et compléter le besoin manuellement.`);
    } finally { setLoadingAI(false); }
  }

  function finish() {
    const need = structured
      ? { context: structured.contexte || "", problem: structured.problematique || "", objectives: (structured.objectifs || []).join(" · "), results: (structured.indicateurs_reussite || []).join(" · "), constraints: (structured.contraintes || []).join(" · "), dependencies: "" }
      : { context: rawNeed, problem: "", objectives: "", results: "", constraints: "", dependencies: "" };
    const scope = structured ? { included: structured.perimetre || [], excluded: [], interfaces: [], hypotheses: [] } : { included: [], excluded: [], interfaces: [], hypotheses: [] };
    const requirements = structured ? (structured.exigences_potentielles || []).map((desc, i) => ({ id: `REQ-${String(i + 1).padStart(3, "0")}`, category: "À qualifier", description: desc, criticality: "Normale", mandatory: false, verificationMethod: "À définir", includeInCDC: true })) : [];
    const tender = {
      id: `t${Date.now()}`, ...form, status: "draft", need, scope, requirements, criteria: [], suppliers: [], needQuestions: [],
      documents: [
        { id: "d1", name: "Procédure AO", category: "Procédure", mandatory: true, owner: "Achats", status: "À préparer" },
        { id: "d9", name: K2_DOC_NAME, category: "Procédure", mandatory: true, owner: "Achats", status: "À préparer" },
        { id: "d2", name: "Cahier des charges", category: "Technique", mandatory: true, owner: "Métier", status: "À préparer" },
        { id: "d3", name: "Cahier de réponses", category: "Technique", mandatory: true, owner: "Métier", status: "À préparer" },
        { id: "d7", name: "Série de prix", category: "Financier", mandatory: true, owner: "Achats", status: "À préparer" },
        { id: "d8", name: "Annexe A - Compréhension des besoins", category: "Technique", mandatory: false, owner: "Métier", status: "À préparer" },
      ],
      history: [{ date: new Date().toISOString().slice(0, 16).replace("T", " "), user: "Vous", action: "Création de l'AO" }],
    };
    onCreate(tender);
  }

  const inputStyle = { border: `1px solid ${C.border}`, backgroundColor: C.surface };

  return (
    <div className="px-4 sm:px-8 py-5 sm:py-7 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1.5 text-sm mb-5 transition-colors hover:opacity-70" style={{ color: C.inkSoft }}><ArrowLeft size={14} /> Retour au tableau de bord</button>
      <h1 className="text-xl font-semibold mb-1" style={{ color: C.ink }}>Nouvel appel d'offres</h1>
      <p className="text-sm mb-6" style={{ color: C.inkSoft }}>Trois courtes étapes, puis l'assistant vous accompagnera dans chaque onglet.</p>

      <div className="flex items-center mb-8">
        {[
          { n: 1, label: "Identification", icon: Info },
          { n: 2, label: "Type d'AO", icon: ListChecks },
          { n: 3, label: "Besoin", icon: Sparkles },
        ].map((s, idx, arr) => {
          const Icon = s.icon;
          const state = s.n < step ? "done" : s.n === step ? "active" : "todo";
          return (
            <React.Fragment key={s.n}>
              <div className="flex flex-col items-center gap-1.5">
                <div className="w-9 h-9 rounded-full flex items-center justify-center text-xs font-semibold transition-all duration-300"
                  style={{
                    backgroundColor: state === "todo" ? C.surface : C.accent,
                    color: state === "todo" ? C.inkSoft : "#fff",
                    border: `1.5px solid ${state === "todo" ? C.border : C.accent}`,
                    boxShadow: state === "active" ? `0 0 0 4px ${C.accentSoft}` : "none",
                  }}>
                  {state === "done" ? <CheckCircle2 size={16} /> : <Icon size={15} />}
                </div>
                <span className="text-[11px] font-medium whitespace-nowrap" style={{ color: state === "todo" ? C.inkSoft : C.ink }}>{s.label}</span>
              </div>
              {idx < arr.length - 1 && <div className="flex-1 h-0.5 mx-2 mb-5 rounded-full transition-colors duration-300" style={{ backgroundColor: s.n < step ? C.accent : C.borderSoft }} />}
            </React.Fragment>
          );
        })}
      </div>

      {step === 1 && (
        <Card className="p-6 ao-view-enter" key="step1">
          <SectionTitle>Identification</SectionTitle>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {[["reference", "Référence AO"], ["title", "Titre"], ["object", "Objet"], ["direction", "Direction"], ["service", "Service"], ["responsibleMetier", "Responsable métier"], ["buyer", "Acheteur"], ["sponsor", "Sponsor"], ["budget", "Budget estimatif (CHF)"], ["dateLaunch", "Date souhaitée de lancement"], ["dateClose", "Date souhaitée de clôture"], ["dateDecision", "Date souhaitée de décision"]].map(([field, label]) => (
              <label key={field} className="text-sm">
                <div className="mb-1" style={{ color: C.inkSoft }}>{label}{["title", "object", "direction", "responsibleMetier", "buyer"].includes(field) && <span style={{ color: C.accent }}> *</span>}</div>
                <input value={form[field]} onChange={e => set(field, e.target.value)} type={field.startsWith("date") ? "date" : field === "budget" ? "number" : "text"} className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle} />
              </label>
            ))}
            <label className="text-sm">
              <div className="mb-1" style={{ color: C.inkSoft }}>Nature du marché</div>
              <select value={form.natureMarche} onChange={e => set("natureMarche", e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle}>
                {Object.keys(SEUILS_MP_GE).map(n => <option key={n} value={n}>{n}</option>)}
              </select>
              <div className="text-xs mt-1" style={{ color: C.inkSoft }}>« Fournitures » = matériel/licences ; « Services » = prestations, logiciel en mode service, maintenance.</div>
            </label>
          </div>

          {form.budget && (() => {
            const sug = suggestProcedure(form.natureMarche, form.budget);
            if (!sug) return null;
            const warnings = scheduleWarnings(form, sug.label);
            return (
              <div className="flex items-start gap-3 mt-4 px-4 py-3 rounded-lg" style={{ backgroundColor: C.accentSoft, border: `1px solid ${C.accent}` }}>
                <Scale size={15} style={{ color: C.accentDark, marginTop: 2 }} />
                <div className="text-sm" style={{ color: C.ink }}>
                  <div className="text-xs font-medium mb-1" style={{ color: C.inkSoft }}>À quoi ça sert ?</div>
                  Le montant de l'AO détermine, par la loi, comment vous devez trouver un fournisseur. Avec {Number(form.budget).toLocaleString("fr-CH")} CHF pour un marché de type « {form.natureMarche} », la règle vous demande : <strong>{sug.label}</strong> ({sug.detail}).
                  {procedureHelpText(sug.label) && <div className="mt-1.5">👉 {procedureHelpText(sug.label)}</div>}
                  <div className="text-xs mt-2" style={{ color: C.inkSoft }}>Seuils selon la pratique retenue en interne — en cas de doute, demandez confirmation aux Achats avant de lancer l'AO.</div>
                  {warnings.map((w, i) => (
                    <div key={i} className="flex items-start gap-2 mt-2 pt-2" style={{ borderTop: `1px solid ${C.accent}` }}>
                      <AlertTriangle size={13} style={{ color: w.level === "warn" ? C.red : C.amber, marginTop: 2, flexShrink: 0 }} />
                      <span className="text-xs flex-1" style={{ color: C.ink }}>{w.text}</span>
                      <button onClick={() => set(w.field, w.value)} className="text-xs font-medium px-2 py-1 rounded-full shrink-0 text-white" style={{ backgroundColor: C.accent }}>{w.applyLabel}</button>
                    </div>
                  ))}
                </div>
              </div>
            );
          })()}

          <div className="flex items-center justify-end gap-3 mt-6">
            {!step1Valid && <span className="text-xs" style={{ color: C.inkSoft }}>Champs requis : Titre, Objet, Direction, Responsable métier, Acheteur</span>}
            <PrimaryButton onClick={() => setStep(2)} disabled={!step1Valid}>Continuer</PrimaryButton>
          </div>
        </Card>
      )}

      {step === 2 && (
        <Card className="p-6 ao-view-enter" key="step2">
          <SectionTitle sub="Le modèle choisi déterminera les documents et sections proposés par défaut.">Type d'appel d'offres</SectionTitle>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5 ao-stagger">
            {TENDER_TYPES.map(type => (
              <button key={type} onClick={() => set("type", type)} className="px-3 py-3 rounded text-sm text-left transition-all duration-150 hover:-translate-y-0.5" style={{ border: `1px solid ${form.type === type ? C.accent : C.border}`, backgroundColor: form.type === type ? C.accentSoft : C.surface, color: form.type === type ? C.accentDark : C.ink, boxShadow: form.type === type ? `0 0 0 3px ${C.accentSoft}` : "none" }}>{type}</button>
            ))}
          </div>
          <div className="flex justify-between mt-6">
            <button onClick={() => setStep(1)} className="px-4 py-2 rounded text-sm" style={{ color: C.inkSoft }}>Retour</button>
            <PrimaryButton onClick={() => setStep(3)} disabled={!form.type}>Continuer</PrimaryButton>
          </div>
        </Card>
      )}

      {step === 3 && (
        <Card className="p-6 ao-view-enter" key="step3">
          <SectionTitle sub="Décrivez librement le besoin, puis laissez l'assistant proposer une structuration.">Définition du besoin</SectionTitle>
          <textarea value={rawNeed} onChange={e => setRawNeed(e.target.value)} rows={4} placeholder="Ex. « Nous voulons remplacer notre outil actuel de gestion des postes et améliorer le patching. »" className="w-full px-3 py-2.5 rounded text-sm outline-none resize-none" style={inputStyle} />
          <div className="flex items-center gap-3 mt-3">
            <button onClick={structureNeed} disabled={!rawNeed.trim() || loadingAI} className="flex items-center gap-2 px-3.5 py-2 rounded text-sm font-medium disabled:opacity-40" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
              <Sparkles size={14} /> {loadingAI ? "Structuration en cours…" : "Structurer le besoin"}
            </button>
            {aiError && <span className="text-xs" style={{ color: C.red }}>{aiError}</span>}
          </div>
          {structured && (
            <div className="mt-5 pt-5 space-y-4 ao-fade-in" style={{ borderTop: `1px solid ${C.borderSoft}` }}>
              <div className="text-xs font-medium" style={{ color: C.accentDark }}>Proposition de l'assistant — à valider avant création</div>
              {[["Contexte", structured.contexte], ["Problématique", structured.problematique]].map(([label, val]) => val && (
                <div key={label}><div className="text-xs font-medium mb-1" style={{ color: C.inkSoft }}>{label}</div><div className="text-sm" style={{ color: C.ink }}>{val}</div></div>
              ))}
              {[["Objectifs", structured.objectifs], ["Périmètre potentiel", structured.perimetre], ["Exigences potentielles", structured.exigences_potentielles], ["Contraintes", structured.contraintes], ["Risques", structured.risques], ["Indicateurs de réussite", structured.indicateurs_reussite]].map(([label, arr]) => arr && arr.length > 0 && (
                <div key={label}><div className="text-xs font-medium mb-1" style={{ color: C.inkSoft }}>{label}</div><ul className="text-sm space-y-0.5" style={{ color: C.ink }}>{arr.map((v, i) => <li key={i}>· {v}</li>)}</ul></div>
              ))}
            </div>
          )}
          <div className="flex justify-between mt-6">
            <button onClick={() => setStep(2)} className="px-4 py-2 rounded text-sm" style={{ color: C.inkSoft }}>Retour</button>
            <PrimaryButton onClick={finish}>Créer l'appel d'offres</PrimaryButton>
          </div>
        </Card>
      )}
    </div>
  );
}

/* --------------------------------- ONGLET INFORMATIONS / BESOIN --------------------------------- */

function inferNature(t) {
  if (t.natureMarche) return t.natureMarche;
  if (t.type === "Matériel") return "Fournitures";
  return "Services";
}

function InfoTab({ t, updateTender }) {
  function set(field, val) { updateTender(prev => ({ ...prev, [field]: val })); }
  const inputStyle = { border: `1px solid ${C.border}` };
  const nature = inferNature(t);
  const sug = t.budget ? suggestProcedure(nature, t.budget) : null;
  const mismatch = sug && t.procedure && !sug.label.toLowerCase().includes(t.procedure.toLowerCase().replace("procédure ", "").split(" ")[0]);
  const warnings = sug ? scheduleWarnings(t, sug.label) : [];

  return (
    <div className="space-y-5">
      <Card className="p-6">
        <SectionTitle sub="Modifiable à tout moment.">Identification</SectionTitle>
        <label className="block text-sm mb-4">
          <div className="mb-1" style={{ color: C.inkSoft }}>Objet</div>
          <textarea rows={2} value={t.object || ""} onChange={e => set("object", e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none resize-y" style={inputStyle} />
        </label>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-x-6 gap-y-4">
          {[["direction", "Direction"], ["service", "Service"], ["responsibleMetier", "Responsable métier"], ["buyer", "Acheteur"], ["sponsor", "Sponsor"], ["type", "Type"], ["category", "Catégorie"], ["procedure", "Procédure"], ["confidentiality", "Confidentialité"]].map(([field, label]) => (
            <label key={field} className="text-sm">
              <div className="mb-1" style={{ color: C.inkSoft }}>{label}</div>
              <input value={t[field] || ""} onChange={e => set(field, e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle} />
            </label>
          ))}
          <label className="text-sm">
            <div className="mb-1" style={{ color: C.inkSoft }}>Budget estimatif (CHF)</div>
            <input type="number" value={t.budget || ""} onChange={e => set("budget", e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle} />
          </label>
          {[["dateLaunch", "Date de lancement"], ["dateClose", "Date de clôture"], ["dateDecision", "Date de décision"]].map(([field, label]) => (
            <label key={field} className="text-sm">
              <div className="mb-1" style={{ color: C.inkSoft }}>{label}</div>
              <input type="date" value={t[field] || ""} onChange={e => set(field, e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle} />
            </label>
          ))}
        </div>
      </Card>
      {sug && (
        <Card className="p-5" style={{ backgroundColor: mismatch ? C.amberSoft : C.accentSoft, borderColor: mismatch ? C.amber : C.accent }}>
          <div className="flex items-start gap-3">
            <Scale size={15} style={{ color: mismatch ? C.amber : C.accentDark, marginTop: 2 }} />
            <div className="text-sm" style={{ color: C.ink }}>
              <div className="text-xs font-medium mb-1" style={{ color: C.inkSoft }}>À quoi ça sert ?</div>
              Le montant de l'AO détermine, par la loi, comment vous devez trouver un fournisseur. Ici, avec {chf(Number(t.budget))} pour un marché de type « {nature} », la règle vous demande : <strong>{sug.label}</strong> ({sug.detail}).
              {procedureHelpText(sug.label) && <div className="mt-1.5">👉 {procedureHelpText(sug.label)}</div>}
              {mismatch && <div className="mt-1.5">⚠️ La procédure choisie pour cet AO (« {t.procedure} ») semble différente — vérifiez avec les Achats avant de continuer.</div>}
              <div className="text-xs mt-2" style={{ color: C.inkSoft }}>Seuils selon la pratique retenue en interne — en cas de doute, demandez confirmation aux Achats.</div>
              {warnings.map((w, i) => (
                <div key={i} className="flex items-start gap-2 mt-2 pt-2" style={{ borderTop: `1px solid ${mismatch ? C.amber : C.accent}` }}>
                  <AlertTriangle size={13} style={{ color: w.level === "warn" ? C.red : C.amber, marginTop: 2, flexShrink: 0 }} />
                  <span className="text-xs flex-1" style={{ color: C.ink }}>{w.text}</span>
                  <button onClick={() => set(w.field, w.value)} className="text-xs font-medium px-2 py-1 rounded-full shrink-0 text-white" style={{ backgroundColor: C.accent }}>{w.applyLabel}</button>
                </div>
              ))}
            </div>
          </div>
        </Card>
      )}
    </div>
  );
}

function NeedTab({ t, updateTender }) {
  const needFields = [
    ["context", "Contexte"], ["problem", "Problème à résoudre"], ["objectives", "Objectifs"],
    ["results", "Résultats attendus"], ["constraints", "Contraintes"], ["dependencies", "Dépendances"],
  ];
  const scopeFields = [
    ["included", "Inclus", "Un élément par ligne"],
    ["excluded", "Exclus", "Un élément par ligne. Ex. : maintenance, formation, migration des données…"],
    ["interfaces", "Interfaces / dépendances", "Un élément par ligne"],
    ["hypotheses", "Hypothèses", "Un élément par ligne"],
  ];
  function setNeed(key, val) { updateTender(prev => ({ ...prev, need: { ...prev.need, [key]: val } })); }
  function setScope(key, text) { updateTender(prev => ({ ...prev, scope: { ...prev.scope, [key]: text.split("\n").map(s => s.trim()).filter(Boolean) } })); }
  const inputStyle = { border: `1px solid ${C.border}` };

  const [rawNeed, setRawNeed] = useState(t.need.context || "");
  const [loadingAI, setLoadingAI] = useState(false);
  const [aiError, setAiError] = useState("");

  async function structureNeed() {
    if (!rawNeed.trim()) return;
    setLoadingAI(true); setAiError("");
    try {
      const prompt = `Tu es un assistant Achats. ` +
        `À partir de la description libre d'un besoin d'achat ci-dessous, produis UNIQUEMENT un objet JSON valide, sans balises markdown ni texte autour, avec exactement ces clés : ` +
        `"contexte" (string), "problematique" (string), "objectifs" (array de strings courtes), "perimetre" (array de strings courtes), ` +
        `"contraintes" (array de strings courtes). ` +
        `Reste générique et n'invente aucune règle juridique ou procédure interne spécifique. Description du besoin : """${rawNeed}"""`;
      const structured = await callClaudeJSON(prompt, 3000);
      updateTender(prev => ({
        ...prev,
        need: {
          ...prev.need,
          context: structured.contexte || prev.need.context,
          problem: structured.problematique || prev.need.problem,
          objectives: (structured.objectifs || []).join(" · ") || prev.need.objectives,
          constraints: (structured.contraintes || []).join(" · ") || prev.need.constraints,
        },
        scope: { ...prev.scope, included: structured.perimetre?.length ? structured.perimetre : prev.scope.included },
      }));
    } catch (e) {
      setAiError(`La structuration automatique a échoué (${e.message || "erreur inconnue"}). Vous pouvez compléter les champs manuellement.`);
    } finally { setLoadingAI(false); }
  }

  return (
    <div className="space-y-6">
      <Card className="p-6">
        <SectionTitle sub="Décrivez librement le besoin, puis laissez l'assistant préremplir le contexte, la problématique, les objectifs, les contraintes et le périmètre inclus ci-dessous.">Structurer le besoin avec l'assistant</SectionTitle>
        <textarea value={rawNeed} onChange={e => setRawNeed(e.target.value)} rows={4} placeholder="Ex. « Nous voulons remplacer notre outil actuel de gestion des postes et améliorer le patching. »" className="w-full px-3 py-2.5 rounded text-sm outline-none resize-none" style={inputStyle} />
        <div className="flex items-center gap-3 mt-3">
          <button onClick={structureNeed} disabled={!rawNeed.trim() || loadingAI} className="flex items-center gap-2 px-3.5 py-2 rounded text-sm font-medium disabled:opacity-40" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
            <Sparkles size={14} /> {loadingAI ? "Structuration en cours…" : "Structurer le besoin (IA)"}
          </button>
          {aiError && <span className="text-xs" style={{ color: C.red }}>{aiError}</span>}
        </div>
      </Card>
      <Card className="p-6">
        <SectionTitle sub="Ces informations alimentent automatiquement le Cahier des charges.">Besoin</SectionTitle>
        <div className="space-y-4">
          {needFields.map(([key, label]) => (
            <label key={key} className="block text-sm">
              <div className="mb-1 font-medium" style={{ color: C.inkSoft }}>{label}</div>
              <textarea rows={2} value={t.need[key] || ""} onChange={e => setNeed(key, e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none resize-y" style={inputStyle} />
            </label>
          ))}
        </div>
      </Card>
      <Card className="p-6">
        <SectionTitle sub="Un élément par ligne. Exemple : vous achetez les serveurs mais pas la maintenance associée — indiquez « Maintenance » en Exclus.">Périmètre</SectionTitle>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
          {scopeFields.map(([key, label, ph]) => (
            <label key={key} className="block text-sm">
              <div className="mb-1 font-medium" style={{ color: C.inkSoft }}>{label}</div>
              <textarea rows={3} value={(t.scope[key] || []).join("\n")} onChange={e => setScope(key, e.target.value)} placeholder={ph} className="w-full px-3 py-2 rounded text-sm outline-none resize-y" style={inputStyle} />
            </label>
          ))}
        </div>
      </Card>
      <AIQuestionBox tender={t} updateTender={updateTender} placeholder="Ex. Qu'est-ce qui se fait habituellement sur le marché pour ce type de besoin ?" contextNote="Pour prendre un peu de recul sur les pratiques du marché avant de figer le besoin." />
    </div>
  );
}

/* --------------------------------- ONGLET DOCUMENTS (rédaction + export) --------------------------------- */

function DocStatusSelect({ value, onChange }) {
  const statuses = ["À préparer", "En cours", "À valider", "Validé", "Généré", "Envoyé", "Reçu", "Contrôlé", "Archivé"];
  return <select value={value} onChange={e => onChange(e.target.value)} className="text-xs rounded px-2 py-1 outline-none" style={{ border: `1px solid ${C.border}` }}>
    {statuses.map(s => <option key={s} value={s}>{s}</option>)}
  </select>;
}

// Modules optionnels du Cahier des charges — n'apparaissent dans le document que si cochés,
// pour s'adapter à la nature réelle du besoin (licence, support, hébergement, contrat-cadre, projet phasé).
const CDC_MODULES = [
  { key: "licences", label: "Licences logicielles", hint: "L'AO porte sur l'achat ou le renouvellement de licences." },
  { key: "hebergement", label: "Hébergement / SaaS", hint: "La solution est hébergée par le fournisseur (cloud, SaaS)." },
  { key: "support", label: "Support et accompagnement technique", hint: "Le marché inclut des prestations de support en plus de la fourniture." },
  { key: "phases", label: "Projet phasé", hint: "Le projet se déroule en plusieurs phases (ex. cadrage puis réalisation)." },
  { key: "contratCadre", label: "Contrat-cadre pluriannuel", hint: "Engagement sur plusieurs années — ajoute durée, exclusivité, changement d'adjudicataire." },
  { key: "montants", label: "Montants plafonds", hint: "Le marché a un montant maximum (total et/ou annuel) à préciser." },
];

function CDCStructurePanel({ tender, updateTender }) {
  const options = tender.cdcOptions || {};
  function toggle(key) { updateTender(prev => ({ ...prev, cdcOptions: { ...(prev.cdcOptions || {}), [key]: !prev.cdcOptions?.[key] } })); }
  return (
    <Card className="p-5 mb-3" style={{ backgroundColor: C.accentSoft, borderColor: C.accent }}>
      <div className="text-sm font-semibold mb-1" style={{ color: C.ink }}>Structure du document</div>
      <div className="text-xs mb-3" style={{ color: C.inkSoft }}>Cochez ce qui correspond à votre besoin : les sections concernées apparaîtront dans le document généré, les autres resteront masquées.</div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
        {CDC_MODULES.map(m => (
          <label key={m.key} className="flex items-start gap-2 p-2 rounded cursor-pointer" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}` }}>
            <input type="checkbox" checked={!!options[m.key]} onChange={() => toggle(m.key)} className="mt-0.5" />
            <div>
              <div className="text-sm" style={{ color: C.ink }}>{m.label}</div>
              <div className="text-xs" style={{ color: C.inkSoft }}>{m.hint}</div>
            </div>
          </label>
        ))}
      </div>
    </Card>
  );
}

const CDC_FIELD_MODULE = {
  phases: "phases", licences_desc: "licences", hebergement_desc: "hebergement", support_desc: "support",
  duree_contrat: "contratCadre", mono_multi: "contratCadre", montant_max: "montants", montant_annuel_max: "montants",
};

function TextDocEditor({ doc, tender, updateTender, onSave, onGenerate }) {
  const rawSchema = (TEXT_SCHEMAS[doc.name] || genericSchema)(tender);
  const schema = doc.name === "Cahier des charges"
    ? rawSchema.filter(s => !CDC_FIELD_MODULE[s.key] || tender.cdcOptions?.[CDC_FIELD_MODULE[s.key]])
    : rawSchema;
  const values = doc.content?.sections || {};
  const [drafting, setDrafting] = useState(false);
  const [draftError, setDraftError] = useState("");
  function get(key, seed) { return values[key] !== undefined ? values[key] : seed; }
  function set(key, val) { onSave({ sections: { ...values, [key]: val } }); }

  async function autoDraft() {
    setDrafting(true); setDraftError("");
    try {
      const context = `Titre de l'AO : ${tender.title}\nObjet : ${tender.object}\nContexte : ${tender.need.context}\nObjectifs : ${tender.need.objectives}\nContraintes : ${tender.need.constraints}\nPérimètre inclus : ${(tender.scope.included || []).join(", ")}\nPérimètre exclus : ${(tender.scope.excluded || []).join(", ")}\nExigences déjà définies : ${tender.requirements.map(r => `${r.id} (${r.category}, ${r.criticality}) — ${r.description}`).join(" | ") || "aucune"}\nProcédure saisie : ${tender.procedure || "non renseignée"}\nDates : lancement ${tender.dateLaunch || "non renseignée"}, clôture ${tender.dateClose || "non renseignée"}, décision ${tender.dateDecision || "non renseignée"}`;
      const sectionList = schema.map(s => `- "${s.key}": ${s.label}`).join("\n");
      const prompt = `Tu es un assistant Achats qui prépare un document opérationnel de consultation — pas un texte marketing. Rédige un projet de contenu, en français administratif sobre, pour chaque section ci-dessous d'un document « ${doc.name} » d'un appel d'offres. ` +
        `RÈGLE ABSOLUE : n'invente STRICTEMENT AUCUNE information factuelle — aucun type de procédure, aucune base légale, aucune date, aucun délai, aucun canal de dépôt, aucun critère d'adjudication, aucune pondération, aucune condition de recevabilité, aucun document obligatoire, aucune règle Achats. ` +
        `Si une information nécessaire à une section n'est pas explicitement présente dans le contexte fourni ci-dessous, écris exactement "[À COMPLÉTER PAR ACHATS]" (information manquante) ou "[À VALIDER PAR ACHATS]" (information qui nécessite une décision/validation, ex. type de procédure, base légale) au lieu de la deviner ou de la généraliser. Ne remplace jamais un de ces deux placeholders déjà présents dans le contexte par une supposition. ` +
        `Réponds UNIQUEMENT avec un objet JSON valide, sans balises markdown et sans aucun texte avant ou après l'accolade ouvrante/fermante, dont les clés sont exactement celles listées ci-dessous et les valeurs sont le texte rédigé (string, quelques phrases maximum par section) :\n${sectionList}\n\nContexte de l'AO :\n${context}`;
      const drafted = await callClaudeJSON(prompt, 6000);
      onSave({ sections: { ...values, ...drafted } });
    } catch (e) {
      setDraftError(`La rédaction automatique a échoué (${e.message || "erreur inconnue"}). Vous pouvez compléter les sections manuellement, ou réessayer.`);
    } finally { setDrafting(false); }
  }

  return (
    <>
    {doc.name === "Cahier des charges" && <CDCStructurePanel tender={tender} updateTender={updateTender} />}
    <Card className="p-6 mt-3">
      <div className="flex items-center justify-between mb-1">
        <div className="text-sm font-semibold" style={{ color: C.ink }}>{doc.name} — rédaction guidée</div>
        <div className="flex items-center gap-3">
          <button onClick={autoDraft} disabled={drafting} className="flex items-center gap-1.5 px-3 py-1.5 rounded text-xs font-medium disabled:opacity-40" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
            <Sparkles size={13} /> {drafting ? "Rédaction en cours…" : "Rédiger automatiquement (IA)"}
          </button>
          <span className="text-xs" style={{ color: C.inkSoft }}>Enregistrement automatique</span>
        </div>
      </div>
      {draftError && <div className="text-xs mb-2" style={{ color: C.red }}>{draftError}</div>}
      {TEXT_SCHEMAS[doc.name] && (
        <div className="flex items-center gap-1.5 mb-4 text-xs" style={{ color: C.accentDark }}>
          <ShieldCheck size={12} /> Structure basée sur le template officiel Achats
        </div>
      )}
      {!TEXT_SCHEMAS[doc.name] && <div className="mb-4" />}
      <div className="space-y-4">
        {schema.map(s => (
          <label key={s.key} className="block text-sm">
            <div className="mb-1 font-medium" style={{ color: C.inkSoft }}>{s.label}</div>
            <textarea rows={3} value={get(s.key, s.seed)} onChange={e => set(s.key, e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none resize-y" style={{ border: `1px solid ${C.border}` }} />
          </label>
        ))}
      </div>
      <div className="flex justify-end mt-5">
        <PrimaryButton icon={Download} onClick={() => onGenerate()}>Générer le document (Word)</PrimaryButton>
      </div>
    </Card>
    </>
  );
}

function TableDocEditor({ doc, onSave, onGenerate }) {
  const rows = doc.content?.rows || defaultPriceScheduleRows();
  function updateCell(i, field, val) { onSave({ rows: rows.map((r, idx) => idx === i ? { ...r, [field]: val } : r) }); }
  function addRow() { onSave({ rows: [...rows, { ref: "", label: "", quantity: 1, unitPrice: "" }] }); }
  function removeRow(i) { onSave({ rows: rows.filter((_, idx) => idx !== i) }); }
  const lineTotal = r => (Number(r.quantity) || 0) * (Number(r.unitPrice) || 0);
  const grandTotal = rows.reduce((sum, r) => sum + lineTotal(r), 0);
  return (
    <Card className="p-6 mt-3">
      <div className="flex items-center justify-between mb-1">
        <div className="text-sm font-semibold" style={{ color: C.ink }}>{doc.name} — rédaction guidée</div>
        <GhostButton icon={Plus} onClick={addRow}>Ajouter une ligne</GhostButton>
      </div>
      <div className="text-xs mb-4" style={{ color: C.inkSoft }}>Une ligne par position (référence, intitulé, quantité) ; le tarif total par ligne et le total général sont calculés automatiquement.</div>
      <div className="overflow-x-auto -mx-1 px-1">
      <table className="w-full text-sm min-w-[640px]">
        <thead><tr style={{ borderBottom: `1px solid ${C.borderSoft}` }}>
          <th className="text-left py-2 text-xs font-medium w-36" style={{ color: C.inkSoft }}>Référence</th>
          <th className="text-left py-2 text-xs font-medium" style={{ color: C.inkSoft }}>Intitulé de la ligne</th>
          <th className="text-left py-2 text-xs font-medium w-24" style={{ color: C.inkSoft }}>Quantité</th>
          <th className="text-left py-2 text-xs font-medium w-32" style={{ color: C.inkSoft }}>Tarif unitaire (CHF)</th>
          <th className="text-right py-2 text-xs font-medium w-32" style={{ color: C.inkSoft }}>Tarif total (CHF)</th>
          <th className="w-8"></th>
        </tr></thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i} style={{ borderBottom: `1px solid ${C.borderSoft}` }}>
              <td className="py-1.5 pr-2"><input value={r.ref} onChange={e => updateCell(i, "ref", e.target.value)} className="w-full px-2 py-1.5 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></td>
              <td className="py-1.5 pr-2"><input value={r.label} onChange={e => updateCell(i, "label", e.target.value)} className="w-full px-2 py-1.5 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></td>
              <td className="py-1.5 pr-2"><input type="number" value={r.quantity} onChange={e => updateCell(i, "quantity", e.target.value)} className="w-full px-2 py-1.5 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></td>
              <td className="py-1.5 pr-2"><input type="number" value={r.unitPrice} onChange={e => updateCell(i, "unitPrice", e.target.value)} className="w-full px-2 py-1.5 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></td>
              <td className="py-1.5 pr-2 text-right tabular-nums" style={{ color: C.inkSoft }}>{chf(lineTotal(r))}</td>
              <td className="py-1.5"><button onClick={() => removeRow(i)}><Trash2 size={14} style={{ color: C.red }} /></button></td>
            </tr>
          ))}
        </tbody>
        <tfoot><tr>
          <td colSpan={4} className="py-2.5 pr-2 text-sm font-semibold text-right" style={{ color: C.ink }}>Total</td>
          <td className="py-2.5 pr-2 text-sm font-semibold text-right tabular-nums" style={{ color: C.ink }}>{chf(grandTotal)}</td>
          <td></td>
        </tr></tfoot>
      </table>
      </div>
      <div className="flex justify-end mt-5"><PrimaryButton icon={Download} onClick={() => onGenerate()}>Générer le document (Excel)</PrimaryButton></div>
    </Card>
  );
}

function QuestionsDocEditor({ doc, tender, updateTender, onGenerate }) {
  const questions = tender.needQuestions || [];
  const [text, setText] = useState("");
  const [suggesting, setSuggesting] = useState(false);
  const [suggestError, setSuggestError] = useState("");
  function addQuestion() {
    if (!text.trim()) return;
    updateTender(prev => ({ ...prev, needQuestions: [...(prev.needQuestions || []), { id: `Q${(prev.needQuestions || []).length + 1}`, question: text.trim() }] }));
    setText("");
  }
  function removeQuestion(id) { updateTender(prev => ({ ...prev, needQuestions: (prev.needQuestions || []).filter(q => q.id !== id) })); }

  async function suggestQuestions() {
    setSuggesting(true); setSuggestError("");
    try {
      const opt = tender.cdcOptions || {};
      const activeModules = Object.entries(opt).filter(([, v]) => v).map(([k]) => ({
        licences: "achat/renouvellement de licences", hebergement: "solution hébergée (cloud/SaaS)", support: "support et accompagnement technique",
        phases: "projet mené en plusieurs phases", contratCadre: "contrat-cadre pluriannuel", montants: "montants plafonnés",
      }[k] || k)).join(", ") || "aucun module spécifique activé";
      const context = `Titre : ${tender.title}\nObjet : ${tender.object}\nContexte : ${tender.need?.context || "non renseigné"}\nObjectifs : ${tender.need?.objectives || "non renseignés"}\nContraintes : ${tender.need?.constraints || "non renseignées"}\n` +
        `Périmètre inclus : ${(tender.scope?.included || []).join(", ") || "non renseigné"}\nPérimètre exclus : ${(tender.scope?.excluded || []).join(", ") || "non renseigné"}\n` +
        `Caractéristiques du marché : ${activeModules}\n` +
        `Exigences déjà définies (par catégorie) : ${Object.entries((tender.requirements || []).reduce((acc, r) => { (acc[r.category] = acc[r.category] || []).push(r.description); return acc; }, {})).map(([cat, ds]) => `${cat} (${ds.length})`).join(", ") || "aucune"}\n` +
        `Critères d'évaluation déjà définis : ${(tender.criteria || []).map(c => `${c.name} (${c.weight}%)`).join(", ") || "aucun"}\n` +
        `Questions déjà posées : ${questions.map(q => q.question).join(" | ") || "aucune"}`;
      const prompt = `Tu es un assistant Achats qui aide un chef de projet à préparer l'annexe « Compréhension des besoins » d'un appel d'offres IT. ` +
        `Cette annexe pose des questions OUVERTES au soumissionnaire pour évaluer sa compréhension du besoin et sa capacité de conseil — pas des exigences à cocher conforme/non conforme. ` +
        `Style attendu, par exemple : « Décrivez l'approche que vous recommandez pour... », « Présentez un ou plusieurs projets comparables auxquels vous avez participé », « Décrivez le modèle de gouvernance que vous recommandez pour... ». ` +
        `À partir du contexte complet de l'AO ci-dessous (besoin, périmètre, exigences déjà définies, critères, caractéristiques du marché), propose 4 à 6 questions pertinentes et spécifiques à CET AO — évite les questions génériques qui iraient pour n'importe quel projet. Ne propose pas de question déjà posée. N'invente aucun fait sur l'AO qui ne soit pas dans le contexte fourni. ` +
        `Réponds UNIQUEMENT avec un tableau JSON d'objets, sans texte autour, où chaque objet a la clé "question" (string).\n\nContexte de l'AO :\n${context}`;
      const suggestions = await callClaudeJSON(prompt, 2000);
      if (!Array.isArray(suggestions)) throw new Error("Format de réponse inattendu");
      updateTender(prev => {
        const base = prev.needQuestions || [];
        const added = suggestions.map((s, i) => ({ id: `Q${base.length + i + 1}`, question: s.question || "" })).filter(q => q.question.trim());
        return { ...prev, needQuestions: [...base, ...added] };
      });
    } catch (e) {
      setSuggestError(`La suggestion automatique a échoué (${e.message || "erreur inconnue"}).`);
    } finally { setSuggesting(false); }
  }

  return (
    <Card className="p-6 mt-3">
      <div className="flex items-start justify-between gap-3 mb-1">
        <div className="text-sm font-semibold" style={{ color: C.ink }}>{doc.name} — rédaction guidée</div>
        <button onClick={suggestQuestions} disabled={suggesting || !tender.need?.context} title={!tender.need?.context ? "Renseignez d'abord le besoin (onglet Besoin & périmètre)" : ""}
          className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium disabled:opacity-40 shrink-0" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
          <Sparkles size={13} /> {suggesting ? "Suggestion en cours…" : "Suggérer des questions (IA)"}
        </button>
      </div>
      <div className="text-xs mb-3" style={{ color: C.inkSoft }}>Posez des questions ouvertes au soumissionnaire pour évaluer sa compréhension du besoin et sa capacité de conseil (approche recommandée, retours d'expérience, gouvernance envisagée...) — à la différence des exigences, il n'y a pas de bonne ou mauvaise réponse attendue ici. L'assistant IA se base sur tout ce que vous avez déjà renseigné pour cet AO (besoin, périmètre, exigences, critères).</div>
      {suggestError && <div className="text-xs mb-3" style={{ color: C.red }}>{suggestError}</div>}
      <div className="flex flex-col sm:flex-row gap-2 mb-4">
        <textarea rows={2} value={text} onChange={e => setText(e.target.value)} placeholder="Ex. Décrivez l'approche que vous recommandez pour..." className="flex-1 px-3 py-2 rounded text-sm outline-none resize-none" style={{ border: `1px solid ${C.border}` }} />
        <PrimaryButton icon={Plus} onClick={addQuestion}>Ajouter</PrimaryButton>
      </div>
      {questions.length === 0 && <div className="text-sm mb-4" style={{ color: C.inkSoft }}>Aucune question pour le moment.</div>}
      <div className="space-y-2 mb-5">
        {questions.map((q, i) => (
          <div key={q.id} className="flex items-start gap-3 p-3 rounded" style={{ border: `1px solid ${C.borderSoft}` }}>
            <span className="text-xs font-mono mt-0.5" style={{ color: C.inkSoft }}>Q{i + 1}</span>
            <div className="flex-1 text-sm" style={{ color: C.ink }}>{q.question}</div>
            <button onClick={() => removeQuestion(q.id)}><Trash2 size={14} style={{ color: C.red }} /></button>
          </div>
        ))}
      </div>
      <div className="flex justify-end">
        <PrimaryButton icon={Download} disabled={!questions.length} onClick={() => onGenerate()}>Générer le document (Word)</PrimaryButton>
      </div>
    </Card>
  );
}

function AutoTablePreview({ doc, tender, onGenerate }) {
  const included = tender.requirements.filter(r => r.includeInCDC !== false);
  return (
    <Card className="p-6 mt-3">
      <div className="text-sm font-semibold mb-2" style={{ color: C.ink }}>{doc.name} — généré depuis la liste des exigences</div>
      {included.length === 0
        ? <div className="text-sm" style={{ color: C.inkSoft }}>Ajoutez des exigences (incluses dans le CDC) dans l'onglet « Exigences » avant de générer ce document.</div>
        : <div className="text-sm mb-1" style={{ color: C.inkSoft }}>{included.length} exigence(s) seront intégrées avec leur format de réponse attendu (Oui/Non ou texte + justificatif).</div>}
      <div className="flex justify-end mt-4"><PrimaryButton icon={Download} disabled={!included.length} onClick={() => onGenerate()}>Générer le document (Excel)</PrimaryButton></div>
    </Card>
  );
}

function K2Preview({ doc, tender, onGenerate }) {
  return (
    <Card className="p-6 mt-3">
      <div className="text-sm font-semibold mb-2" style={{ color: C.ink }}>{doc.name}</div>
      <div className="text-sm mb-1" style={{ color: C.inkSoft }}>Formulaire officiel du Guide romand pour les marchés publics (procédure ouverte, www.simap.ch), généré automatiquement à partir des informations déjà saisies dans cet AO (identification, calendrier, critères). Les champs propres à l'adjudicateur non couverts par l'app restent affichés comme [À COMPLÉTER PAR ACHATS] / [À VALIDER PAR ACHATS] — à compléter directement dans le fichier Word généré.</div>
      <div className="text-xs mt-2" style={{ color: C.inkSoft }}>Basé sur : direction/service, acheteur, responsable métier, objet, calendrier (lancement/clôture/décision), critères d'évaluation ({tender.criteria.length} défini(s)).</div>
      <div className="flex justify-end mt-4"><PrimaryButton icon={Download} onClick={() => onGenerate()}>Générer le document (Word)</PrimaryButton></div>
    </Card>
  );
}

function DocumentsTab({ t, updateTender }) {
  const [openId, setOpenId] = useState(null);

  function setStatus(docId, status) { updateTender(prev => ({ ...prev, documents: prev.documents.map(d => d.id === docId ? { ...d, status } : d) })); }
  function setContent(docId, content) { updateTender(prev => ({ ...prev, documents: prev.documents.map(d => d.id === docId ? { ...d, content } : d) })); }
  function handleGenerate(doc) { generateDocumentFile(t, doc); setStatus(doc.id, "Généré"); }
  function exportAll() { t.documents.forEach((doc, i) => setTimeout(() => handleGenerate(doc), i * 450)); }

  const openDoc = t.documents.find(d => d.id === openId);

  return (
    <div>
      <Card>
        <div className="px-6 pt-6 pb-2 flex items-start justify-between">
          <SectionTitle sub="Rédigez chaque document directement en ligne, puis générez le fichier correspondant (Word ou Excel).">Documents de l'AO</SectionTitle>
          <GhostButton icon={Download} onClick={exportAll}>Exporter tout le dossier</GhostButton>
        </div>
        <div className="overflow-x-auto">
        <table className="w-full text-sm min-w-[720px]">
          <thead><tr style={{ borderBottom: `1px solid ${C.borderSoft}` }}>
            {["Document", "Catégorie", "Obligatoire", "Responsable", "Statut", ""].map(h => <th key={h} className="text-left px-6 py-2 text-xs font-medium" style={{ color: C.inkSoft }}>{h}</th>)}
          </tr></thead>
          <tbody>
            {t.documents.map(d => (
              <tr key={d.id} style={{ borderBottom: `1px solid ${C.borderSoft}`, backgroundColor: openId === d.id ? C.accentSoft : "transparent" }}>
                <td className="px-6 py-3" style={{ color: C.ink }}>{d.name}</td>
                <td className="px-6 py-3" style={{ color: C.inkSoft }}>{d.category}</td>
                <td className="px-6 py-3" style={{ color: C.inkSoft }}>{d.mandatory ? "Oui" : "Non"}</td>
                <td className="px-6 py-3" style={{ color: C.inkSoft }}>{d.owner}</td>
                <td className="px-6 py-3"><DocStatusSelect value={d.status} onChange={s => setStatus(d.id, s)} /></td>
                <td className="px-6 py-3">
                  <div className="flex items-center gap-2">
                    <GhostButton icon={PenLine} onClick={() => setOpenId(openId === d.id ? null : d.id)}>{openId === d.id ? "Fermer" : "Rédiger"}</GhostButton>
                    <GhostButton icon={Download} onClick={() => handleGenerate(d)}>Télécharger</GhostButton>
                  </div>
                </td>
              </tr>
            ))}
            {t.documents.length === 0 && <tr><td colSpan={6} className="px-6 py-6 text-sm" style={{ color: C.inkSoft }}>Aucun document pour le moment.</td></tr>}
          </tbody>
        </table>
        </div>
      </Card>

      {openDoc && TABLE_DOC_NAMES.includes(openDoc.name) && <TableDocEditor doc={openDoc} onSave={c => setContent(openDoc.id, c)} onGenerate={() => handleGenerate(openDoc)} />}
      {openDoc && AUTO_TABLE_DOC_NAMES.includes(openDoc.name) && <AutoTablePreview doc={openDoc} tender={t} onGenerate={() => handleGenerate(openDoc)} />}
      {openDoc && QUESTIONS_DOC_NAMES.includes(openDoc.name) && <QuestionsDocEditor doc={openDoc} tender={t} updateTender={updateTender} onGenerate={() => handleGenerate(openDoc)} />}
      {openDoc && openDoc.name === K2_DOC_NAME && <K2Preview doc={openDoc} tender={t} onGenerate={() => handleGenerate(openDoc)} />}
      {openDoc && !TABLE_DOC_NAMES.includes(openDoc.name) && !AUTO_TABLE_DOC_NAMES.includes(openDoc.name) && !QUESTIONS_DOC_NAMES.includes(openDoc.name) && openDoc.name !== K2_DOC_NAME && <TextDocEditor doc={openDoc} tender={t} updateTender={updateTender} onSave={c => setContent(openDoc.id, c)} onGenerate={() => handleGenerate(openDoc)} />}
    </div>
  );
}

/* --------------------------------- ONGLET EXIGENCES --------------------------------- */

// Types d'exigences standards retrouvés dans des cahiers des charges réels — guide le choix
// au lieu de laisser une catégorie libre. "Autre…" reste disponible pour un besoin spécifique.
const STANDARD_REQUIREMENT_CATEGORIES = [
  "Sécurité", "Fonctionnel", "Intégration", "Support",
  "Documentation", "Conception et réalisation", "Exploitation",
  "Réglementaire SI", "Standards IT et processus informatiques",
];

function RequirementsTab({ t, updateTender }) {
  const byCat = {}; t.requirements.forEach(r => { (byCat[r.category] = byCat[r.category] || []).push(r); });
  const [form, setForm] = useState({ category: STANDARD_REQUIREMENT_CATEGORIES[0], customCategory: "", description: "", criticality: "Normale", mandatory: true, verificationMethod: "", includeInCDC: true });
  const [suggesting, setSuggesting] = useState(false);
  const [suggestError, setSuggestError] = useState("");
  function exportExcel() {
    const header = ["ID", "Catégorie", "Exigence", "Criticité", "Obligatoire", "Méthode de vérification", "Dans le CDC"];
    const dataRows = t.requirements.map(r => [r.id, r.category, r.description, r.criticality, r.mandatory ? "Oui" : "Non", r.verificationMethod, r.includeInCDC === false ? "Non" : "Oui"]);
    downloadXLSX(`${t.reference}_exigences.xlsx`, [styledReportSheet("Exigences", "EXIGENCES", header, dataRows, [10, 22, 55, 14, 14, 30, 14])]);
  }
  async function suggestRequirements() {
    setSuggesting(true); setSuggestError("");
    try {
      const prompt = `Tu es un assistant Achats. À partir du contexte d'appel d'offres ci-dessous, propose 5 à 8 exigences pertinentes à ajouter à la liste, en les répartissant dans les types standards suivants lorsque pertinent : ${STANDARD_REQUIREMENT_CATEGORIES.join(", ")} (ou un autre type court si aucun ne convient). ` +
        `N'invente aucun chiffre ou règle spécifique non mentionnée ; reste générique et professionnel. Réponds UNIQUEMENT avec un tableau JSON d'objets, sans texte autour, où chaque objet a les clés : "category" (string courte), "description" (string), "criticality" (une valeur parmi Bloquante, Critique, Majeure, Normale, Souhaitable), "mandatory" (booléen), "verificationMethod" (string courte).\n\n` +
        `Titre : ${t.title}\nObjet : ${t.object}\nContexte : ${t.need.context}\nObjectifs : ${t.need.objectives}\nPérimètre inclus : ${(t.scope.included || []).join(", ")}\nExigences déjà présentes : ${t.requirements.map(r => r.description).join(" | ") || "aucune"}`;
      const suggestions = await callClaudeJSON(prompt, 2500);
      if (!Array.isArray(suggestions)) throw new Error("Format de réponse inattendu");
      updateTender(prev => {
        let existingIds = prev.requirements.map(r => r.id);
        const newReqs = suggestions.map(s => {
          const prefix = (s.category || "GEN").trim().slice(0, 3).toUpperCase();
          const n = existingIds.filter(id => id.startsWith(prefix)).length + 1;
          const id = `${prefix}-${String(n).padStart(3, "0")}`;
          existingIds = [...existingIds, id];
          return { id, category: s.category || "Général", description: s.description || "", criticality: s.criticality || "Normale", mandatory: !!s.mandatory, verificationMethod: s.verificationMethod || "", includeInCDC: true };
        });
        return { ...prev, requirements: [...prev.requirements, ...newReqs] };
      });
    } catch (e) {
      setSuggestError(`La suggestion automatique a échoué (${e.message || "erreur inconnue"}).`);
    } finally { setSuggesting(false); }
  }
  function addRequirement() {
    const category = form.category === "Autre…" ? form.customCategory.trim() : form.category;
    if (!category || !form.description.trim()) return;
    const prefix = category.trim().slice(0, 3).toUpperCase();
    const n = t.requirements.filter(r => r.id.startsWith(prefix)).length + 1;
    const newReq = { id: `${prefix}-${String(n).padStart(3, "0")}`, category, description: form.description, criticality: form.criticality, mandatory: form.mandatory, verificationMethod: form.verificationMethod, includeInCDC: form.includeInCDC };
    updateTender(prev => ({ ...prev, requirements: [...prev.requirements, newReq] }));
    upsertLibraryItem("requirement", { category, description: form.description, criticality: form.criticality, mandatory: form.mandatory, verification_method: form.verificationMethod }).catch(() => {});
    setForm(f => ({ ...f, description: "", verificationMethod: "" }));
  }
  function addFromLibrary(item) {
    const category = item.category || "Général";
    const prefix = category.trim().slice(0, 3).toUpperCase();
    const n = t.requirements.filter(r => r.id.startsWith(prefix)).length + 1;
    const newReq = { id: `${prefix}-${String(n).padStart(3, "0")}`, category, description: item.description, criticality: item.criticality || "Normale", mandatory: !!item.mandatory, verificationMethod: item.verification_method || "", includeInCDC: true };
    updateTender(prev => ({ ...prev, requirements: [...prev.requirements, newReq] }));
  }
  function removeRequirement(id) { updateTender(prev => ({ ...prev, requirements: prev.requirements.filter(r => r.id !== id) })); }
  function toggleIncludeInCDC(id) { updateTender(prev => ({ ...prev, requirements: prev.requirements.map(r => r.id === id ? { ...r, includeInCDC: r.includeInCDC === false } : r) })); }
  function setCriticality(id, criticality) { updateTender(prev => ({ ...prev, requirements: prev.requirements.map(r => r.id === id ? { ...r, criticality } : r) })); }
  const inputStyle = { border: `1px solid ${C.border}` };

  return (
    <div className="space-y-4">
      <LibraryPicker kind="requirement" title="Réutiliser depuis un autre AO" onPick={addFromLibrary} />
      <Card className="p-5">
        <div className="flex items-start justify-between gap-3 mb-1">
          <SectionTitle sub="Choisissez le type parmi les catégories standards utilisées dans les cahiers des charges — chaque exigence sera reprise automatiquement dans le cahier des charges et le cahier de réponses, sauf si vous décochez « Inclure dans le CDC ».">Ajouter une exigence</SectionTitle>
          <button onClick={suggestRequirements} disabled={suggesting || !t.need.context} title={!t.need.context ? "Renseignez d'abord le besoin (onglet Besoin & périmètre)" : ""}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium disabled:opacity-40 shrink-0" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
            <Sparkles size={13} /> {suggesting ? "Suggestion en cours…" : "Suggérer des exigences (IA)"}
          </button>
        </div>
        {suggestError && <div className="text-xs mb-3" style={{ color: C.red }}>{suggestError}</div>}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <label className="text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Type d'exigence</div>
            <select value={form.category} onChange={e => setForm(f => ({ ...f, category: e.target.value }))} className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle}>
              {STANDARD_REQUIREMENT_CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
              <option value="Autre…">Autre…</option>
            </select>
          </label>
          {form.category === "Autre…" && (
            <label className="text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Préciser le type</div>
              <input value={form.customCategory} onChange={e => setForm(f => ({ ...f, customCategory: e.target.value }))} placeholder="Ex. Migration, Formation…" className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle} /></label>
          )}
          <label className="text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Criticité</div>
            <select value={form.criticality} onChange={e => setForm(f => ({ ...f, criticality: e.target.value }))} className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle}>
              {Object.keys(CRITICALITY_META).map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </label>
          <div className="text-xs sm:col-span-2 -mt-2" style={{ color: C.inkSoft }}>Bloquante = essentielle, non évaluable sur une échelle : l'offre est conforme ou éliminée · Souhaitable = facultative, un plus non obligatoire.</div>
          <label className="text-sm col-span-2"><div className="mb-1" style={{ color: C.inkSoft }}>Description de l'exigence</div>
            <textarea rows={2} value={form.description} onChange={e => setForm(f => ({ ...f, description: e.target.value }))} className="w-full px-3 py-2 rounded text-sm outline-none resize-none" style={inputStyle} /></label>
          <label className="text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Méthode de vérification</div>
            <input value={form.verificationMethod} onChange={e => setForm(f => ({ ...f, verificationMethod: e.target.value }))} placeholder="Ex. Démonstration, documentation…" className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle} /></label>
          <div className="flex items-center gap-4 mt-6">
            <label className="text-sm flex items-center gap-2"><input type="checkbox" checked={form.mandatory} onChange={e => setForm(f => ({ ...f, mandatory: e.target.checked }))} /> Obligatoire</label>
            <label className="text-sm flex items-center gap-2"><input type="checkbox" checked={form.includeInCDC} onChange={e => setForm(f => ({ ...f, includeInCDC: e.target.checked }))} /> Inclure dans le CDC</label>
          </div>
        </div>
        <div className="flex justify-end mt-4"><PrimaryButton icon={Plus} onClick={addRequirement}>Ajouter à la liste</PrimaryButton></div>
      </Card>

      <Card>
        <div className="px-6 pt-6 flex items-start justify-between">
          <SectionTitle>Liste des exigences</SectionTitle>
          <GhostButton icon={Download} disabled={!t.requirements.length} onClick={exportExcel}>Exporter (Excel)</GhostButton>
        </div>
        {Object.keys(byCat).length === 0 && <div className="px-6 pb-6 text-sm" style={{ color: C.inkSoft }}>Aucune exigence définie — ajoutez-en une ci-dessus.</div>}
        {Object.entries(byCat).map(([cat, reqs]) => (
          <div key={cat} className="px-6 pb-4">
            <div className="text-xs font-semibold uppercase tracking-wide mt-3 mb-2" style={{ color: C.inkSoft }}>{cat}</div>
            <div className="space-y-2">
              {reqs.map(r => (
                <div key={r.id} className="flex items-start gap-3 p-3 rounded" style={{ border: `1px solid ${C.borderSoft}`, opacity: r.includeInCDC === false ? 0.55 : 1 }}>
                  <span className="text-xs font-mono mt-0.5" style={{ color: C.inkSoft }}>{r.id}</span>
                  <div className="flex-1"><div className="text-sm" style={{ color: C.ink }}>{r.description}</div><div className="text-xs mt-1" style={{ color: C.inkSoft }}>Vérification : {r.verificationMethod || "—"}{r.mandatory ? " · Obligatoire" : ""}{r.includeInCDC === false ? " · Exclue du CDC" : ""}</div></div>
                  <select value={r.criticality} onChange={e => setCriticality(r.id, e.target.value)} className="text-xs font-medium rounded-full px-2.5 py-1 outline-none shrink-0" style={{ color: CRITICALITY_META[r.criticality] || C.inkSoft, backgroundColor: C.slateSoft, border: "none" }}>
                    {Object.keys(CRITICALITY_META).map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                  <button onClick={() => toggleIncludeInCDC(r.id)} title={r.includeInCDC === false ? "Inclure dans le CDC" : "Exclure du CDC"} className="p-1 rounded hover:bg-black/5">
                    {r.includeInCDC === false ? <Circle size={14} style={{ color: C.inkSoft }} /> : <CheckCircle2 size={14} style={{ color: C.green }} />}
                  </button>
                  <button onClick={() => removeRequirement(r.id)}><Trash2 size={14} style={{ color: C.red }} /></button>
                </div>
              ))}
            </div>
          </div>
        ))}
      </Card>
    </div>
  );
}

/* --------------------------------- ONGLET CRITÈRES --------------------------------- */

function CriteriaTab({ t, updateTender }) {
  const total = weightSum(t);
  const [newCriterion, setNewCriterion] = useState({ name: "", weight: "" });
  const [suggesting, setSuggesting] = useState(false);
  const [suggestError, setSuggestError] = useState("");
  function setWeight(id, w) { updateTender(prev => ({ ...prev, criteria: prev.criteria.map(c => c.id === id ? { ...c, weight: Number(w) } : c) })); }
  // Ajuste un critère via +/- et redistribue automatiquement la différence sur les autres, pour rester à 100%.
  function adjustWeight(id, delta) {
    updateTender(prev => {
      const target = prev.criteria.find(c => c.id === id);
      if (!target) return prev;
      const newTargetWeight = Math.max(0, Math.min(100, target.weight + delta));
      const actualDelta = newTargetWeight - target.weight;
      if (actualDelta === 0) return prev;
      const others = prev.criteria.filter(c => c.id !== id);
      const othersTotal = others.reduce((s, c) => s + c.weight, 0);
      let updated = prev.criteria.map(c => {
        if (c.id === id) return { ...c, weight: newTargetWeight };
        if (othersTotal === 0) return c;
        const share = Math.round((c.weight / othersTotal) * -actualDelta);
        return { ...c, weight: Math.max(0, c.weight + share) };
      });
      const diff = 100 - updated.reduce((s, c) => s + c.weight, 0);
      if (diff !== 0) {
        const idx = updated.findIndex(c => c.id !== id);
        if (idx >= 0) updated[idx] = { ...updated[idx], weight: Math.max(0, updated[idx].weight + diff) };
      }
      return { ...prev, criteria: updated };
    });
  }
  function removeCriterion(id) { updateTender(prev => ({ ...prev, criteria: prev.criteria.filter(c => c.id !== id) })); }
  function addCriterion() {
    if (!newCriterion.name.trim()) return;
    updateTender(prev => ({ ...prev, criteria: [...prev.criteria, { id: `c${Date.now()}`, name: newCriterion.name.trim(), weight: Number(newCriterion.weight) || 0 }] }));
    upsertLibraryItem("criterion", { description: newCriterion.name.trim(), weight: Number(newCriterion.weight) || null }).catch(() => {});
    setNewCriterion({ name: "", weight: "" });
  }
  function addCriterionFromLibrary(item) {
    updateTender(prev => ({ ...prev, criteria: [...prev.criteria, { id: `c${Date.now()}`, name: item.description, weight: item.weight || 0 }] }));
  }
  async function suggestCriteria() {
    setSuggesting(true); setSuggestError("");
    try {
      const prompt = `Tu es un assistant Achats. Propose une grille de 4 à 6 critères d'évaluation pour l'appel d'offres ci-dessous, avec une pondération en % dont le total fait exactement 100. Inclure systématiquement un critère « Prix ». ` +
        `Réponds UNIQUEMENT avec un tableau JSON d'objets, sans texte autour, où chaque objet a les clés "name" (string courte) et "weight" (nombre entier).\n\n` +
        `Titre : ${t.title}\nObjet : ${t.object}\nContexte : ${t.need.context}\nExigences : ${t.requirements.map(r => `${r.category} — ${r.description}`).join(" | ") || "aucune"}`;
      const suggestions = await callClaudeJSON(prompt, 1500);
      if (!Array.isArray(suggestions)) throw new Error("Format de réponse inattendu");
      updateTender(prev => ({ ...prev, criteria: suggestions.map((s, i) => ({ id: `c${Date.now()}_${i}`, name: s.name || "Critère", weight: Number(s.weight) || 0 })) }));
    } catch (e) {
      setSuggestError(`La suggestion automatique a échoué (${e.message || "erreur inconnue"}).`);
    } finally { setSuggesting(false); }
  }
  return (
    <div className="space-y-4">
    <LibraryPicker kind="criterion" title="Réutiliser depuis un autre AO" onPick={addCriterionFromLibrary} />
    <Card className="p-6">
      <div className="flex items-start justify-between gap-3 mb-1">
        <SectionTitle sub="La « pondération » = l'importance de chaque critère, en %. Le total doit faire 100%. Ajoutez un critère « Prix » pour que son coût soit noté automatiquement dans l'onglet Évaluation.">Critères d'évaluation</SectionTitle>
        <button onClick={suggestCriteria} disabled={suggesting} className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium disabled:opacity-40 shrink-0" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
          <Sparkles size={13} /> {suggesting ? "Suggestion en cours…" : "Suggérer des critères (IA)"}
        </button>
      </div>
      {suggestError && <div className="text-xs mb-3" style={{ color: C.red }}>{suggestError}</div>}
      {t.criteria.length > 0 && <div className="text-xs mb-3" style={{ color: C.inkSoft }}>Utilisez les boutons − / + pour ajuster un critère : les autres se réajustent automatiquement pour rester à 100%. « Suggérer des critères » remplace la liste actuelle.</div>}
      {t.criteria.length === 0 && <div className="text-sm mb-4" style={{ color: C.inkSoft }}>Aucun critère défini — ajoutez-en un ci-dessous.</div>}
      <div className="space-y-2.5 mb-5">
        {t.criteria.map(c => (
          <div key={c.id} className="flex items-center gap-3">
            <div className="w-40 text-sm shrink-0" style={{ color: C.ink }}>{c.name}</div>
            <div className="flex-1"><ProgressBar value={c.weight} /></div>
            <button onClick={() => adjustWeight(c.id, -5)} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0" style={{ border: `1px solid ${C.border}`, color: C.inkSoft }}>−</button>
            <input type="number" value={c.weight} onChange={e => setWeight(c.id, e.target.value)} className="w-16 px-2 py-1.5 rounded text-sm text-right outline-none shrink-0" style={{ border: `1px solid ${C.border}` }} />
            <button onClick={() => adjustWeight(c.id, 5)} className="w-7 h-7 rounded-full flex items-center justify-center shrink-0" style={{ border: `1px solid ${C.border}`, color: C.inkSoft }}>+</button>
            <span className="text-sm w-4 shrink-0" style={{ color: C.inkSoft }}>%</span>
            <button onClick={() => removeCriterion(c.id)}><Trash2 size={14} style={{ color: C.red }} /></button>
          </div>
        ))}
      </div>
      <div className="flex items-end gap-3 pt-4" style={{ borderTop: `1px solid ${C.borderSoft}` }}>
        <label className="text-sm flex-1"><div className="mb-1" style={{ color: C.inkSoft }}>Nouveau critère</div>
          <input value={newCriterion.name} onChange={e => setNewCriterion(f => ({ ...f, name: e.target.value }))} placeholder="Ex. Technique, Prix, Support…" className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
        <label className="text-sm w-28"><div className="mb-1" style={{ color: C.inkSoft }}>Pondération %</div>
          <input type="number" value={newCriterion.weight} onChange={e => setNewCriterion(f => ({ ...f, weight: e.target.value }))} className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
        <PrimaryButton icon={Plus} onClick={addCriterion}>Ajouter</PrimaryButton>
      </div>
      {t.criteria.length > 0 && (
        <div className="mt-4 flex items-center gap-2">
          {total === 100 ? <CheckCircle2 size={15} style={{ color: C.green }} /> : <AlertTriangle size={15} style={{ color: C.red }} />}
          <span className="text-sm" style={{ color: total === 100 ? C.green : C.red }}>Total : {total}% {total !== 100 && "— doit être égal à 100%"}</span>
        </div>
      )}
      <div className="flex items-center gap-3 mt-5 pt-4" style={{ borderTop: `1px solid ${C.borderSoft}` }}>
        <label className="text-sm flex items-center gap-2">
          <span style={{ color: C.inkSoft }}>Seuil minimum d'admission</span>
          <input type="number" min="0" max="100" value={t.evaluationThreshold ?? 60} onChange={e => updateTender(prev => ({ ...prev, evaluationThreshold: Number(e.target.value) }))} className="w-16 px-2 py-1.5 rounded text-sm text-right outline-none" style={{ border: `1px solid ${C.border}` }} />
          <span style={{ color: C.inkSoft }}>/ 100</span>
        </label>
        <span className="text-xs" style={{ color: C.inkSoft }}>— une offre sous ce score est proposée comme "REFUSÉ" dans la grille d'évaluation exportée.</span>
      </div>
    </Card>
    <AIQuestionBox tender={t} updateTender={updateTender} placeholder="Ex. Quels sont les critères d'évaluation habituels pour ce type de besoin ?" contextNote="Pour benchmarker vos critères par rapport aux pratiques usuelles du marché." />
    </div>
  );
}

/* --------------------------------- ONGLET FOURNISSEURS --------------------------------- */

// Alimente le registre partagé des fournisseurs au fil des AO — best-effort, ne bloque jamais
// la création du fournisseur dans l'AO si ça échoue (ex. doublon d'email déjà connu).
async function upsertSupplierRegistry(name, email) {
  if (!name?.trim()) return;
  const looksLikeEmail = /.+@.+\..+/.test(email || "");
  await supabase.from("suppliers_registry").insert({ name: name.trim(), email: looksLikeEmail ? email.trim() : null }).select();
}

// Alimente la bibliothèque partagée d'exigences/critères — best-effort, évite les doublons
// évidents (même texte déjà présent) avant d'insérer.
async function upsertLibraryItem(kind, fields) {
  const description = (fields.description || "").trim();
  if (!description) return;
  const { data: existing } = await supabase.from("library_items").select("id").eq("kind", kind).ilike("description", description).limit(1);
  if (existing && existing.length) return;
  await supabase.from("library_items").insert({ kind, description, ...fields });
}

// Widget de réutilisation : liste les éléments déjà utilisés dans d'autres AO (exigences ou
// critères), avec recherche, pour les ajouter en un clic plutôt que de tout ressaisir.
function LibraryPicker({ kind, onPick, title }) {
  const [open, setOpen] = useState(false);
  const [items, setItems] = useState(null);
  const [query, setQuery] = useState("");

  function toggle() {
    setOpen(o => !o);
    if (!open && items === null) {
      supabase.from("library_items").select("id, category, description, criticality, mandatory, verification_method, weight")
        .eq("kind", kind).order("created_at", { ascending: false }).limit(200)
        .then(({ data }) => setItems(data || []));
    }
  }

  const filtered = (items || []).filter(it => !query.trim() || it.description.toLowerCase().includes(query.toLowerCase()) || (it.category || "").toLowerCase().includes(query.toLowerCase()));

  return (
    <div className="mb-3">
      <button onClick={toggle} className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium shrink-0" style={{ border: `1px solid ${C.border}`, color: C.inkSoft, backgroundColor: C.surface }}>
        <Building2 size={13} /> {title}
      </button>
      {open && (
        <Card className="p-4 mt-2 ao-fade-in">
          <input value={query} onChange={e => setQuery(e.target.value)} placeholder="Rechercher…" className="w-full px-3 py-2 rounded text-sm outline-none mb-3" style={{ border: `1px solid ${C.border}` }} />
          {items === null && <div className="text-xs" style={{ color: C.inkSoft }}>Chargement…</div>}
          {items !== null && filtered.length === 0 && <div className="text-xs" style={{ color: C.inkSoft }}>Aucun élément { query ? "ne correspond à la recherche" : "réutilisable pour l'instant — la bibliothèque se remplit au fil des AO" }.</div>}
          <div className="space-y-1.5 max-h-64 overflow-y-auto">
            {filtered.map(it => (
              <div key={it.id} className="flex items-start gap-2 p-2 rounded" style={{ border: `1px solid ${C.borderSoft}` }}>
                <div className="flex-1 min-w-0 text-sm" style={{ color: C.ink }}>
                  {it.category && <span className="text-xs mr-1.5" style={{ color: C.inkSoft }}>[{it.category}]</span>}
                  {it.description}{it.weight != null && <span className="text-xs ml-1.5" style={{ color: C.inkSoft }}>({it.weight}%)</span>}
                </div>
                <button onClick={() => onPick(it)} className="flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium shrink-0 text-white" style={{ backgroundColor: C.accent }}>
                  <Plus size={11} /> Ajouter
                </button>
              </div>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}

function SuppliersTab({ t, updateTender }) {
  const [form, setForm] = useState({ name: "", contact: "" });
  const [showRegistry, setShowRegistry] = useState(false);
  const [registryItems, setRegistryItems] = useState(null);
  const [registryQuery, setRegistryQuery] = useState("");
  const inputStyle = { border: `1px solid ${C.border}` };
  const notYetPublished = missingMandatoryDocs(t).length > 0 || t.criteria.length === 0;
  const existingNames = new Set(t.suppliers.map(s => s.name.toLowerCase()));

  function insertSupplier(name, contact) {
    const expectedDocs = t.documents.filter(d => d.mandatory).map(d => ({ name: d.name, received: false }));
    updateTender(prev => ({ ...prev, suppliers: [...prev.suppliers, {
      id: `s${Date.now()}`, name: name.trim(), contact: (contact || "").trim(),
      documents: expectedDocs.length ? expectedDocs : [{ name: "Cahier de réponses", received: false }, { name: "Offre financière", received: false }],
      price: { initial: 0, annual: 0, maintenance: 0, migration: 0 }, evaluations: {}, files: [],
    }] }));
  }
  const [uploadingFor, setUploadingFor] = useState(null);
  async function uploadFiles(supplierId, fileList) {
    setUploadingFor(supplierId);
    const newFiles = [];
    for (const file of Array.from(fileList)) {
      const path = `${t.id}/${supplierId}/${Date.now()}_${file.name}`;
      const { error } = await supabase.storage.from("tender-files").upload(path, file);
      if (!error) newFiles.push({ id: path, name: file.name, path, size: file.size, type: file.type, uploadedAt: new Date().toISOString() });
    }
    if (newFiles.length) {
      updateTender(prev => ({ ...prev, suppliers: prev.suppliers.map(s => s.id !== supplierId ? s : { ...s, files: [...(s.files || []), ...newFiles] }) }));
    }
    setUploadingFor(null);
  }
  async function downloadFile(file) {
    const { data, error } = await supabase.storage.from("tender-files").createSignedUrl(file.path, 60);
    if (!error && data) window.open(data.signedUrl, "_blank");
  }
  async function removeFile(supplierId, file) {
    await supabase.storage.from("tender-files").remove([file.path]);
    updateTender(prev => ({ ...prev, suppliers: prev.suppliers.map(s => s.id !== supplierId ? s : { ...s, files: (s.files || []).filter(f => f.id !== file.id) }) }));
  }
  function addSupplier() {
    if (!form.name.trim()) return;
    insertSupplier(form.name, form.contact);
    upsertSupplierRegistry(form.name, form.contact).catch(() => {});
    setForm({ name: "", contact: "" });
  }
  function addFromRegistry(item) { insertSupplier(item.name, item.email); }
  function toggleRegistry() {
    setShowRegistry(o => !o);
    if (!showRegistry && registryItems === null) {
      supabase.from("suppliers_registry").select("id, name, email, domain, region").order("name")
        .then(({ data }) => setRegistryItems(data || []));
    }
  }
  const filteredRegistry = (registryItems || []).filter(it => !registryQuery.trim() || it.name.toLowerCase().includes(registryQuery.toLowerCase()));
  function removeSupplier(id) { updateTender(prev => ({ ...prev, suppliers: prev.suppliers.filter(s => s.id !== id) })); }
  function toggleDoc(supplierId, idx) {
    updateTender(prev => ({ ...prev, suppliers: prev.suppliers.map(s => s.id !== supplierId ? s : { ...s, documents: s.documents.map((d, i) => i === idx ? { ...d, received: !d.received } : d) }) }));
  }
  function setPrice(supplierId, field, value) {
    updateTender(prev => ({ ...prev, suppliers: prev.suppliers.map(s => s.id !== supplierId ? s : { ...s, price: { ...s.price, [field]: Number(value) || 0 } }) }));
  }

  return (
    <div className="space-y-4">
      {notYetPublished && (
        <div className="flex items-start gap-3 px-4 py-3 rounded-lg" style={{ backgroundColor: C.amberSoft, border: `1px solid ${C.amber}` }}>
          <AlertTriangle size={15} style={{ color: C.amber, marginTop: 2 }} />
          <div className="text-sm" style={{ color: C.ink }}>Cet onglet sert à enregistrer les <strong>soumissionnaires ayant déposé une offre</strong>, une fois l'AO publié. Les critères et/ou documents obligatoires ne sont pas encore finalisés — vous pouvez continuer, mais vérifiez que l'AO est bien prêt à être publié avant de solliciter des offres.</div>
        </div>
      )}
      <div>
        <button onClick={toggleRegistry} className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium shrink-0" style={{ border: `1px solid ${C.border}`, color: C.inkSoft, backgroundColor: C.surface }}>
          <Building2 size={13} /> Choisir depuis le registre des fournisseurs
        </button>
        {showRegistry && (
          <Card className="p-4 mt-2 ao-fade-in">
            <input value={registryQuery} onChange={e => setRegistryQuery(e.target.value)} placeholder="Rechercher un fournisseur…" className="w-full px-3 py-2 rounded text-sm outline-none mb-3" style={inputStyle} />
            {registryItems === null && <div className="text-xs" style={{ color: C.inkSoft }}>Chargement…</div>}
            {registryItems !== null && filteredRegistry.length === 0 && <div className="text-xs" style={{ color: C.inkSoft }}>Aucun fournisseur dans le registre pour l'instant.</div>}
            <div className="space-y-1.5 max-h-64 overflow-y-auto">
              {filteredRegistry.map(it => {
                const already = existingNames.has(it.name.toLowerCase());
                return (
                  <div key={it.id} className="flex items-start gap-2 p-2 rounded" style={{ border: `1px solid ${C.borderSoft}` }}>
                    <div className="flex-1 min-w-0 text-sm" style={{ color: C.ink }}>
                      {it.name}
                      <div className="text-xs mt-0.5 flex items-center gap-1.5 flex-wrap" style={{ color: C.inkSoft }}>
                        <span>{it.email || "—"}</span>
                        {it.domain && <Badge color={C.accentDark} bg={C.accentSoft}>{it.domain}</Badge>}
                        {it.region && <Badge color={C.inkSoft} bg={C.slateSoft}>{it.region}</Badge>}
                      </div>
                    </div>
                    {already ? (
                      <span className="text-xs shrink-0" style={{ color: C.inkSoft }}>Déjà ajouté</span>
                    ) : (
                      <button onClick={() => addFromRegistry(it)} className="flex items-center gap-1 px-2 py-1 rounded-full text-xs font-medium shrink-0 text-white" style={{ backgroundColor: C.accent }}>
                        <Plus size={11} /> Ajouter
                      </button>
                    )}
                  </div>
                );
              })}
            </div>
          </Card>
        )}
      </div>
      <Card className="p-5">
        <SectionTitle sub="Les documents obligatoires de l'onglet Documents sont repris automatiquement comme checklist attendue.">Ajouter un nouveau soumissionnaire</SectionTitle>
        <div className="flex items-end gap-3">
          <label className="text-sm flex-1"><div className="mb-1" style={{ color: C.inkSoft }}>Nom du fournisseur</div>
            <input value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle} /></label>
          <label className="text-sm flex-1"><div className="mb-1" style={{ color: C.inkSoft }}>Contact</div>
            <input value={form.contact} onChange={e => setForm(f => ({ ...f, contact: e.target.value }))} placeholder="email@fournisseur.example" className="w-full px-3 py-2 rounded text-sm outline-none" style={inputStyle} /></label>
          <PrimaryButton icon={Plus} onClick={addSupplier}>Ajouter</PrimaryButton>
        </div>
      </Card>

      {t.suppliers.length === 0 && <Card className="p-6 text-sm" style={{ color: C.inkSoft }}>Aucun fournisseur enregistré — ajoutez-en un ci-dessus.</Card>}
      {t.suppliers.map(s => {
        const missing = s.documents.filter(d => !d.received);
        return (
          <Card key={s.id} className="p-5">
            <div className="flex items-center justify-between">
              <div><div className="text-sm font-semibold" style={{ color: C.ink }}>{s.name}</div><div className="text-xs mt-0.5" style={{ color: C.inkSoft }}>{s.contact || "—"}</div></div>
              <div className="flex items-center gap-2">
                {missing.length === 0 ? <Badge color={C.green} bg={C.greenSoft}>Dossier complet</Badge> : <Badge color={C.amber} bg={C.amberSoft}>{missing.length} document{missing.length > 1 ? "s" : ""} manquant{missing.length > 1 ? "s" : ""}</Badge>}
                <button onClick={() => removeSupplier(s.id)}><Trash2 size={14} style={{ color: C.red }} /></button>
              </div>
            </div>
            <div className="mt-3 flex flex-wrap gap-2">
              {s.documents.map((d, i) => (
                <button key={i} onClick={() => toggleDoc(s.id, i)} className="flex items-center gap-1 text-xs px-2 py-1 rounded" style={{ backgroundColor: d.received ? C.greenSoft : C.redSoft, color: d.received ? C.green : C.red }}>
                  {d.received ? <CheckCircle2 size={12} /> : <AlertTriangle size={12} />} {d.name}
                </button>
              ))}
            </div>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mt-4 pt-4" style={{ borderTop: `1px solid ${C.borderSoft}` }}>
              {[["initial", "Prix initial"], ["annual", "Coût annuel"], ["maintenance", "Maintenance"], ["migration", "Migration"]].map(([field, label]) => (
                <label key={field} className="text-xs"><div className="mb-1" style={{ color: C.inkSoft }}>{label} (CHF)</div>
                  <input type="number" value={s.price[field] || ""} onChange={e => setPrice(s.id, field, e.target.value)} className="w-full px-2 py-1.5 rounded text-sm outline-none" style={inputStyle} /></label>
              ))}
            </div>
            <div className="mt-4 pt-4" style={{ borderTop: `1px solid ${C.borderSoft}` }}>
              <div className="flex items-center justify-between mb-2">
                <div className="text-xs font-medium" style={{ color: C.inkSoft }}>Documents reçus de ce fournisseur ({(s.files || []).length})</div>
                <label className="flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium cursor-pointer" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
                  <Upload size={12} /> {uploadingFor === s.id ? "Envoi…" : "Ajouter des fichiers"}
                  <input type="file" multiple className="hidden" disabled={uploadingFor === s.id} onChange={e => { if (e.target.files.length) uploadFiles(s.id, e.target.files); e.target.value = ""; }} />
                </label>
              </div>
              {(s.files || []).length === 0 ? (
                <div className="text-xs" style={{ color: C.inkSoft }}>Aucun fichier — glissez-y les offres, annexes ou Excel reçus du fournisseur (Excel, Word, PDF…).</div>
              ) : (
                <div className="space-y-1.5">
                  {s.files.map(f => (
                    <div key={f.id} className="flex items-center gap-2 px-2.5 py-1.5 rounded text-xs" style={{ border: `1px solid ${C.borderSoft}` }}>
                      <FileText size={13} style={{ color: C.inkSoft }} className="shrink-0" />
                      <button onClick={() => downloadFile(f)} className="flex-1 min-w-0 text-left truncate hover:underline" style={{ color: C.ink }}>{f.name}</button>
                      <span className="shrink-0" style={{ color: C.inkSoft }}>{f.size ? `${(f.size / 1024).toFixed(0)} ko` : ""}</span>
                      <button onClick={() => removeFile(s.id, f)} className="shrink-0"><Trash2 size={12} style={{ color: C.red }} /></button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </Card>
        );
      })}
    </div>
  );
}

/* --------------------------------- ÉVALUATION GUIDÉE (pas à pas) --------------------------------- */

function GuidedEvaluation({ t, updateTender, evaluatorName, onExit }) {
  const criteria = t.criteria.filter(c => c.name !== "Prix");
  const [si, setSi] = useState(0);
  const [ci, setCi] = useState(0);
  const supplier = t.suppliers[si];
  const criterion = criteria[ci];
  const totalSteps = t.suppliers.length * criteria.length;
  const stepNum = si * criteria.length + ci + 1;

  const existing = supplier.evaluations[criterion.id]?.find(n => n.evaluator === evaluatorName);
  const [note, setNote] = useState(existing?.note ?? null);
  const [comment, setComment] = useState(existing?.comment ?? "");

  useEffect(() => {
    const en = supplier.evaluations[criterion.id]?.find(n => n.evaluator === evaluatorName);
    setNote(en?.note ?? null);
    setComment(en?.comment ?? "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [si, ci]);

  function persist(nextNote, nextComment) {
    updateTender(prev => ({ ...prev, suppliers: prev.suppliers.map(s => {
      if (s.id !== supplier.id) return s;
      const ex = s.evaluations[criterion.id] || [];
      const idx = ex.findIndex(n => n.evaluator === evaluatorName);
      const updated = idx >= 0 ? ex.map((n, i) => i === idx ? { ...n, note: nextNote, comment: nextComment } : n) : [...ex, { evaluator: evaluatorName, note: nextNote, comment: nextComment }];
      return { ...s, evaluations: { ...s.evaluations, [criterion.id]: updated } };
    }) }));
  }
  function pick(v) { setNote(v); persist(v, comment); }
  function onCommentBlur() { if (note != null) persist(note, comment); }

  function goNext() {
    if (ci < criteria.length - 1) setCi(i => i + 1);
    else if (si < t.suppliers.length - 1) { setSi(i => i + 1); setCi(0); }
    else onExit();
  }
  function goPrev() {
    if (ci > 0) setCi(i => i - 1);
    else if (si > 0) { setSi(i => i - 1); setCi(criteria.length - 1); }
  }

  const reqs = requirementsForCriterion(t, criterion.name);
  const files = supplier.files || [];
  const isFirst = si === 0 && ci === 0;
  const isLast = si === t.suppliers.length - 1 && ci === criteria.length - 1;

  async function downloadFile(f) {
    const { data, error } = await supabase.storage.from("tender-files").createSignedUrl(f.path, 60);
    if (!error && data) window.open(data.signedUrl, "_blank");
  }

  return (
    <Card className="p-6 ao-view-enter" key={`${si}-${ci}`}>
      <div className="flex items-center justify-between mb-1">
        <div className="text-xs font-medium" style={{ color: C.inkSoft }}>Étape {stepNum} / {totalSteps}</div>
        <button onClick={onExit} className="text-xs font-medium" style={{ color: C.accentDark }}>Quitter le mode guidé</button>
      </div>
      <div className="w-full mb-5"><ProgressBar value={Math.round((stepNum / totalSteps) * 100)} /></div>

      <div className="text-xs uppercase tracking-wider mb-1" style={{ color: C.inkSoft }}>{supplier.name}</div>
      <div className="text-lg font-semibold mb-4" style={{ color: C.ink }}>{criterion.name} <span className="text-sm font-normal" style={{ color: C.inkSoft }}>· pondération {criterion.weight}%</span></div>

      {reqs.length > 0 && (
        <div className="mb-4 p-3 rounded" style={{ backgroundColor: C.slateSoft }}>
          <div className="text-xs font-medium mb-1.5" style={{ color: C.inkSoft }}>Exigences liées</div>
          <ul className="text-sm space-y-1" style={{ color: C.ink }}>
            {reqs.map(r => <li key={r.id}>· {r.description}</li>)}
          </ul>
        </div>
      )}

      <div className="mb-5">
        <div className="text-xs font-medium mb-1.5" style={{ color: C.inkSoft }}>Documents du fournisseur ({files.length})</div>
        {files.length === 0 ? (
          <div className="text-xs" style={{ color: C.inkSoft }}>Aucun document reçu de ce fournisseur.</div>
        ) : (
          <div className="flex flex-wrap gap-2">
            {files.map(f => (
              <button key={f.id} onClick={() => downloadFile(f)} className="flex items-center gap-1.5 px-2.5 py-1.5 rounded text-xs hover:bg-black/5" style={{ border: `1px solid ${C.border}`, color: C.ink }}>
                <FileText size={12} style={{ color: C.inkSoft }} /> {f.name}
              </button>
            ))}
          </div>
        )}
      </div>

      <div className="mb-4">
        <div className="text-xs font-medium mb-2" style={{ color: C.inkSoft }}>Votre note (0 à 5)</div>
        <div className="flex items-center gap-2">
          {[0, 1, 2, 3, 4, 5].map(v => (
            <button key={v} onClick={() => pick(v)} className="w-10 h-10 rounded-lg text-sm font-semibold transition-all"
              style={note === v ? { backgroundColor: C.accent, color: "#fff", boxShadow: `0 0 0 3px ${C.accentSoft}` } : { border: `1px solid ${C.border}`, color: C.inkSoft, backgroundColor: C.surface }}>
              {v}
            </button>
          ))}
        </div>
      </div>
      <label className="block text-sm mb-6">
        <div className="mb-1" style={{ color: C.inkSoft }}>Commentaire (optionnel)</div>
        <textarea rows={3} value={comment} onChange={e => setComment(e.target.value)} onBlur={onCommentBlur} className="w-full px-3 py-2 rounded text-sm outline-none resize-none" style={{ border: `1px solid ${C.border}` }} />
      </label>

      <div className="flex items-center justify-between pt-4" style={{ borderTop: `1px solid ${C.borderSoft}` }}>
        <button onClick={goPrev} disabled={isFirst} className="flex items-center gap-1.5 px-3.5 py-2 rounded text-sm disabled:opacity-30" style={{ color: C.inkSoft }}><ArrowLeft size={14} /> Précédent</button>
        <PrimaryButton onClick={goNext} icon={isLast ? CheckCircle2 : ChevronRight}>{isLast ? "Terminer" : "Suivant"}</PrimaryButton>
      </div>
    </Card>
  );
}

/* --------------------------------- ONGLET ÉVALUATION --------------------------------- */

function EvaluationTab({ t, updateTender, evaluatorName }) {
  const [guided, setGuided] = useState(false);
  if (t.criteria.length === 0 || t.suppliers.length === 0) return <Card className="p-6 text-sm" style={{ color: C.inkSoft }}>Définissez d'abord les critères et les fournisseurs pour démarrer l'évaluation.</Card>;
  if (guided) return <GuidedEvaluation t={t} updateTender={updateTender} evaluatorName={evaluatorName || "Vous"} onExit={() => setGuided(false)} />;

  function addNote(criterionId, supplierId, evaluator, note) {
    updateTender(prev => ({ ...prev, suppliers: prev.suppliers.map(s => {
      if (s.id !== supplierId) return s;
      const existing = s.evaluations[criterionId] || [];
      const idx = existing.findIndex(n => n.evaluator === evaluator);
      const updatedNotes = idx >= 0 ? existing.map((n, i) => i === idx ? { ...n, note } : n) : [...existing, { evaluator, note, comment: "" }];
      return { ...s, evaluations: { ...s.evaluations, [criterionId]: updatedNotes } };
    }) }));
  }

  function exportGrid() {
    downloadXLSX(`${t.reference}_grille_evaluation.xlsx`, buildEvaluationWorkbookSheets(t));
  }

  function exportFinancial() {
    const header = ["Fournisseur", "Prix initial", "Coûts récurrents (3 ans)", "Migration", "TCO 3 ans"];
    const dataRows = t.suppliers.map(s => {
      const recurring = ((s.price.annual || 0) + (s.price.maintenance || 0)) * 3;
      const tco = (s.price.initial || 0) + recurring + (s.price.migration || 0);
      return [s.name, s.price.initial || 0, recurring, s.price.migration || 0, tco];
    });
    downloadXLSX(`${t.reference}_analyse_financiere.xlsx`, [styledReportSheet("Analyse financière", "ANALYSE FINANCIÈRE", header, dataRows, [30, 16, 22, 16, 16])]);
  }

  return (
    <div className="space-y-6">
      <div className="flex items-start justify-between gap-3 px-4 py-3 rounded-lg" style={{ backgroundColor: C.slateSoft, color: C.inkSoft }}>
        <div className="text-sm">
          Notez chaque fournisseur de <strong>0</strong> (pas du tout conforme) à <strong>5</strong> (excellent) pour chaque critère. L'app calcule le score pondéré automatiquement — vous n'avez qu'à noter.
        </div>
        <button onClick={() => setGuided(true)} className="flex items-center gap-1.5 px-3.5 py-2 rounded-full text-xs font-semibold text-white shrink-0" style={{ backgroundColor: C.accent }}>
          <Sparkles size={13} /> Évaluation guidée
        </button>
      </div>
      <Card className="p-4" style={{ backgroundColor: C.accentSoft, borderColor: C.accent }}>
        <div className="flex items-start justify-between gap-3">
          <div className="text-sm" style={{ color: C.ink }}>
            <div className="font-semibold mb-0.5">Grille d'évaluation du jury (Excel)</div>
            <div className="text-xs" style={{ color: C.inkSoft }}>Une fiche par fournisseur (catégories, exigences, notes et commentaires à main levée), avec formules Excel live (moyenne par catégorie, note pondérée, score final, seuil d'admission), les critères éliminatoires détectés automatiquement, et une synthèse avec classement, recommandation et bloc de signature du jury — à distribuer directement pour notation.</div>
          </div>
          <GhostButton icon={Download} disabled={!t.criteria.length || !t.suppliers.length} onClick={exportGrid}>Exporter la grille du jury</GhostButton>
        </div>
      </Card>
      <div className="flex justify-end gap-2">
        <GhostButton icon={Download} onClick={exportFinancial}>Exporter l'analyse financière (Excel)</GhostButton>
      </div>
      {t.criteria.filter(c => c.name !== "Prix").map(c => (
        <Card key={c.id} className="p-5">
          <div className="text-sm font-semibold mb-3" style={{ color: C.ink }}>{c.name} <span className="font-normal" style={{ color: C.inkSoft }}>· pondération {c.weight}%</span></div>
          <div className="grid gap-3 overflow-x-auto" style={{ gridTemplateColumns: `repeat(${t.suppliers.length}, minmax(180px,1fr))` }}>
            {t.suppliers.map(s => {
              const notes = s.evaluations[c.id] || [];
              const avg = criterionAverage(s, c.id);
              const spread = notes.length > 1 ? Math.max(...notes.map(n => n.note)) - Math.min(...notes.map(n => n.note)) : 0;
              return (
                <div key={s.id} className="p-3 rounded" style={{ border: `1px solid ${C.borderSoft}` }}>
                  <div className="text-xs font-medium mb-2 truncate" style={{ color: C.ink }}>{s.name}</div>
                  {notes.map((n, i) => <div key={i} className="flex items-center justify-between text-xs mb-1" style={{ color: C.inkSoft }}><span>{n.evaluator}</span><span className="font-semibold" style={{ color: C.ink }}>{n.note}/5</span></div>)}
                  {notes.length === 0 && <div className="text-xs mb-1" style={{ color: C.inkSoft }}>Pas encore noté</div>}
                  <div className="flex items-center gap-1 mt-2">{[0, 1, 2, 3, 4, 5].map(v => <button key={v} onClick={() => addNote(c.id, s.id, evaluatorName || "Vous", v)} className="w-6 h-6 rounded text-xs flex items-center justify-center" style={{ border: `1px solid ${C.border}`, backgroundColor: C.surface, color: C.inkSoft }}>{v}</button>)}</div>
                  {spread >= 3 && <div className="flex items-center gap-1 mt-2 text-xs" style={{ color: C.red }}><AlertTriangle size={12} /> Écart important entre évaluateurs</div>}
                  {avg != null && <div className="text-xs mt-2" style={{ color: C.inkSoft }}>Moyenne : {avg.toFixed(1)}/5</div>}
                </div>
              );
            })}
          </div>
        </Card>
      ))}
      {t.criteria.some(c => c.name === "Prix") && (
        <Card className="p-5">
          <div className="text-sm font-semibold mb-3" style={{ color: C.ink }}>Prix <span className="font-normal" style={{ color: C.inkSoft }}>· note calculée automatiquement (prix le plus bas = note maximale)</span></div>
          <div className="grid gap-3 overflow-x-auto" style={{ gridTemplateColumns: `repeat(${t.suppliers.length}, minmax(180px,1fr))` }}>
            {t.suppliers.map(s => {
              const total = (s.price.initial || 0) + (s.price.annual || 0) * 3 + (s.price.maintenance || 0) * 3 + (s.price.migration || 0);
              return <div key={s.id} className="p-3 rounded" style={{ border: `1px solid ${C.borderSoft}` }}>
                <div className="text-xs font-medium mb-1 truncate" style={{ color: C.ink }}>{s.name}</div>
                <div className="text-xs" style={{ color: C.inkSoft }}>TCO 3 ans : {chf(total)}</div>
                <div className="text-xs" style={{ color: C.inkSoft }}>Note : {(priceScores(t)[s.id] ?? 0).toFixed(1)}/5</div>
              </div>;
            })}
          </div>
        </Card>
      )}
    </div>
  );
}

/* --------------------------------- ONGLET COMPARAISON --------------------------------- */

function ComparisonTab({ t }) {
  if (t.criteria.length === 0 || t.suppliers.length === 0) return <Card className="p-6 text-sm" style={{ color: C.inkSoft }}>Aucune donnée à comparer pour le moment.</Card>;
  const scores = computeSupplierScores(t);
  function exportComparison() {
    const header = ["Critère", "Poids", ...t.suppliers.map(s => s.name)];
    const dataRows = t.criteria.map(c => {
      const row = [c.name, c.weight];
      scores.forEach(sc => { const b = sc.breakdown.find(x => x.criterionId === c.id); row.push(b && b.avg != null ? b.weighted.toFixed(1) : "—"); });
      return row;
    });
    dataRows.push(["Score total", "", ...scores.map(sc => sc.total)]);
    const colWidths = [24, 12, ...t.suppliers.map(() => 20)];
    downloadXLSX(`${t.reference}_comparatif_fournisseurs.xlsx`, [styledReportSheet("Comparatif", "COMPARATIF FOURNISSEURS", header, dataRows, colWidths)]);
  }
  return (
    <Card>
      <div className="px-6 pt-6 flex items-start justify-between"><SectionTitle>Comparaison des fournisseurs</SectionTitle><GhostButton icon={Download} onClick={exportComparison}>Exporter (Excel)</GhostButton></div>
      <div className="overflow-x-auto">
      <table className="w-full text-sm min-w-[560px]">
        <thead><tr style={{ borderBottom: `1px solid ${C.borderSoft}` }}>
          <th className="text-left px-6 py-2 text-xs font-medium" style={{ color: C.inkSoft }}>Critère</th>
          <th className="text-left px-2 py-2 text-xs font-medium" style={{ color: C.inkSoft }}>Poids</th>
          {t.suppliers.map(s => <th key={s.id} className="text-right px-4 py-2 text-xs font-medium" style={{ color: C.inkSoft }}>{s.name}</th>)}
        </tr></thead>
        <tbody>
          {t.criteria.map(c => (
            <tr key={c.id} style={{ borderBottom: `1px solid ${C.borderSoft}` }}>
              <td className="px-6 py-2.5" style={{ color: C.ink }}>{c.name}</td>
              <td className="px-2 py-2.5" style={{ color: C.inkSoft }}>{c.weight}%</td>
              {scores.map(sc => { const b = sc.breakdown.find(x => x.criterionId === c.id); return <td key={sc.supplierId} className="text-right px-4 py-2.5 tabular-nums" style={{ color: C.ink }}>{b && b.avg != null ? b.weighted.toFixed(1) : "—"}</td>; })}
            </tr>
          ))}
          <tr><td className="px-6 py-3 font-semibold" style={{ color: C.ink }}>Score total</td><td className="px-2 py-3" />{scores.map(sc => <td key={sc.supplierId} className="text-right px-4 py-3 font-semibold tabular-nums" style={{ color: C.accentDark }}>{sc.total}</td>)}</tr>
        </tbody>
      </table>
      </div>
      <div className="px-6 py-4 flex gap-3" style={{ borderTop: `1px solid ${C.borderSoft}` }}>
        {scores.map((sc, i) => <Badge key={sc.supplierId} color={i === 0 ? C.green : C.inkSoft} bg={i === 0 ? C.greenSoft : C.slateSoft}>{i + 1}{i === 0 ? "er" : "e"} — {sc.supplierName.split("—")[1]?.trim() || sc.supplierName}</Badge>)}
      </div>
    </Card>
  );
}

/* --------------------------------- ONGLET SYNTHÈSE --------------------------------- */

function SynthesisTab({ t }) {
  if (t.criteria.length === 0 || t.suppliers.length === 0) return <Card className="p-6 text-sm" style={{ color: C.inkSoft }}>La synthèse sera disponible une fois l'évaluation renseignée.</Card>;
  const scores = computeSupplierScores(t);
  const winner = scores[0];
  const supplierMeta = t.suppliers.reduce((acc, s) => { acc[s.id] = s; return acc; }, {});

  function strengthsWeaknesses(sc) { const sorted = [...sc.breakdown].filter(b => b.avg != null).sort((a, b) => b.avg - a.avg); return { best: sorted[0], worst: sorted[sorted.length - 1] }; }

  function exportRanking() {
    const header = ["Fournisseur", "Score technique", "Score financier", "Score global", "Classement"];
    const dataRows = scores.map((sc, i) => {
      const priceBreak = sc.breakdown.find(b => b.criterionName === "Prix");
      const techTotal = sc.breakdown.filter(b => b.criterionName !== "Prix").reduce((s, b) => s + b.weighted, 0);
      return [sc.supplierName, +techTotal.toFixed(1), priceBreak ? priceBreak.weighted.toFixed(1) : "—", sc.total, i + 1];
    });
    downloadXLSX(`${t.reference}_synthese.xlsx`, [styledReportSheet("Synthèse", "SYNTHÈSE DE L'ÉVALUATION", header, dataRows, [30, 18, 18, 16, 14])]);
  }

  function exportSynthesisWord() {
    const missingBySupplier = t.suppliers.map(s => `${s.name} : ${s.documents.filter(d => !d.received).length} document(s) manquant(s)`).join("\n");
    const rankingText = scores.map((sc, i) => `${i + 1}. ${sc.supplierName} — ${sc.total}/100`).join("\n");
    const sections = [
      { label: "1. Objet de l'AO", text: t.object },
      { label: "2. Besoin", text: t.need.context },
      { label: "3. Fournisseurs consultés", text: t.suppliers.map(s => s.name).join("\n") },
      { label: "4. Offres reçues", text: `${t.suppliers.length} offre(s) reçue(s).` },
      { label: "5. Conformité documentaire", text: missingBySupplier },
      { label: "6. Évaluation technique", text: t.criteria.filter(c => c.name !== "Prix").map(c => `${c.name} (${c.weight}%)`).join("\n") },
      { label: "7. Évaluation financière", text: "Notation du prix selon la méthode du prix le plus bas (voir onglet Évaluation)." },
      { label: "8. Résultats", text: rankingText },
      { label: "9. Classement", text: rankingText },
      { label: "10. Points forts / points faibles", text: scores.map(sc => { const { best, worst } = strengthsWeaknesses(sc); return `${sc.supplierName} — Point fort : ${best ? best.criterionName : "—"} · Point à surveiller : ${worst ? worst.criterionName : "—"}`; }).join("\n") },
      { label: "11. Risques", text: "À compléter par l'équipe métier / Achats." },
      { label: "12. Recommandation", text: `${winner.supplierName} arrive en tête avec un score de ${winner.total}/100. Cette recommandation est une proposition soumise à validation.` },
      { label: "13. Décision", text: "À compléter à l'issue de la validation." },
    ];
    let body = coverBlockXml(t, "Synthèse finale");
    body += metaAndHistoryXml(t);
    sections.forEach(s => { body += heading(s.label, 1); body += para(s.text && s.text.trim() ? s.text : "(à compléter)"); });
    packageDocx(body + glossaryXml(), `${t.reference}_synthese.docx`, t);
  }

  return (
    <div className="space-y-5">
      <div className="flex justify-end gap-2">
        <GhostButton icon={Download} onClick={exportSynthesisWord}>Télécharger la synthèse (Word)</GhostButton>
        <GhostButton icon={Download} onClick={exportRanking}>Exporter le classement (Excel)</GhostButton>
      </div>
      <Card className="p-6" style={{ backgroundColor: C.accentSoft, borderColor: C.accent }}>
        <div className="text-xs font-medium mb-1" style={{ color: C.accentDark }}>Recommandation — proposition soumise à validation humaine</div>
        <div className="text-sm" style={{ color: C.ink }}>Sur la base des scores pondérés calculés, <strong>{winner.supplierName}</strong> arrive en tête avec un score de <strong>{winner.total}/100</strong>. Cette recommandation ne constitue pas une décision d'attribution et doit être validée par les instances compétentes.</div>
      </Card>
      {scores.map((sc, i) => {
        const { best, worst } = strengthsWeaknesses(sc);
        const missing = supplierMeta[sc.supplierId].documents.filter(d => !d.received);
        return (
          <Card key={sc.supplierId} className="p-5">
            <div className="flex items-center justify-between mb-3"><div className="text-sm font-semibold" style={{ color: C.ink }}>#{i + 1} — {sc.supplierName}</div><div className="text-sm font-semibold" style={{ color: C.accentDark }}>{sc.total}/100</div></div>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-sm">
              <div><div className="text-xs" style={{ color: C.inkSoft }}>Conformité documentaire</div><div style={{ color: missing.length ? C.amber : C.green }}>{missing.length ? `${missing.length} manquant(s)` : "Complète"}</div></div>
              {best && <div><div className="text-xs" style={{ color: C.inkSoft }}>Point fort</div><div style={{ color: C.ink }}>{best.criterionName} ({best.avg.toFixed(1)}/5)</div></div>}
              {worst && <div><div className="text-xs" style={{ color: C.inkSoft }}>Point à surveiller</div><div style={{ color: C.ink }}>{worst.criterionName} ({worst.avg.toFixed(1)}/5)</div></div>}
            </div>
          </Card>
        );
      })}
    </div>
  );
}

/* --------------------------------- ONGLET HISTORIQUE + CONTRÔLE QUALITÉ --------------------------------- */

function HistoryTab({ t }) {
  return (
    <Card>
      <div className="px-6 pt-6"><SectionTitle>Historique</SectionTitle></div>
      <div className="px-6 pb-6">{t.history.map((h, i) => <div key={i} className="flex gap-4 py-3" style={{ borderBottom: i < t.history.length - 1 ? `1px solid ${C.borderSoft}` : "none" }}><div className="text-xs font-mono w-36 shrink-0" style={{ color: C.inkSoft }}>{h.date}</div><div className="text-sm" style={{ color: C.ink }}><span className="font-medium">{h.user}</span> — {h.action}</div></div>)}</div>
    </Card>
  );
}

function ConfirmDialog({ open, title, message, confirmLabel = "Confirmer", danger, onConfirm, onCancel }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center px-4" style={{ backgroundColor: "rgba(24,34,52,0.45)" }} onClick={onCancel}>
      <div onClick={e => e.stopPropagation()} className="w-full max-w-sm rounded-xl p-5" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, boxShadow: "0 10px 30px rgba(0,0,0,0.15)" }}>
        <div className="text-sm font-semibold mb-2" style={{ color: C.ink }}>{title}</div>
        <div className="text-sm mb-5" style={{ color: C.inkSoft }}>{message}</div>
        <div className="flex justify-end gap-2">
          <button onClick={onCancel} className="px-3.5 py-2 rounded-full text-sm font-medium transition-colors hover:bg-black/[0.03]" style={{ border: `1px solid ${C.border}`, color: C.inkSoft }}>Annuler</button>
          <button onClick={onConfirm} className="px-3.5 py-2 rounded-full text-sm font-semibold text-white" style={{ backgroundColor: danger ? C.red : C.accent }}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

function QualityCheck({ t }) {
  const { checks, pct } = readinessChecks(t);
  return (
    <Card className="p-5 mb-6">
      <div className="flex items-center justify-between mb-3"><div className="text-sm font-semibold" style={{ color: C.ink }}>AO prêt à être publié à {pct}%</div></div>
      <ProgressBar value={pct} />
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-1.5 mt-4 ao-stagger">
        {checks.map(c => <div key={c.label} className="flex items-center gap-2 text-sm">{c.ok ? <CheckCircle2 size={14} style={{ color: C.green }} /> : <Circle size={14} style={{ color: C.amber }} />}<span style={{ color: c.ok ? C.ink : C.inkSoft }}>{c.label}</span></div>)}
      </div>
    </Card>
  );
}

/* --------------------------------- DÉTAIL AO --------------------------------- */

function GuidanceBanner({ tender, onJump }) {
  const step = nextStepFor(tender);
  const Icon = step.icon;
  return (
    <button onClick={() => onJump(step.tab)} className="w-full text-left flex items-center gap-3 px-5 py-3.5 rounded-lg mb-5 transition-colors hover:brightness-[0.98]"
      style={{ backgroundColor: step.done ? C.greenSoft : C.accentSoft, border: `1px solid ${step.done ? C.green : C.accent}` }}>
      <div key={step.done ? "done" : step.tab} className="w-8 h-8 rounded-full flex items-center justify-center shrink-0 ao-pop" style={{ backgroundColor: step.done ? C.green : C.accent }}>
        {step.done ? <CheckCircle2 size={16} color="#fff" /> : <Icon size={15} color="#fff" />}
      </div>
      <div className="flex-1">
        <div className="text-xs font-medium" style={{ color: step.done ? C.green : C.accentDark }}>{step.done ? "AO prêt" : "Prochaine étape"}</div>
        <div className="text-sm" style={{ color: C.ink }}>{step.text}</div>
      </div>
      <span className="text-xs font-medium flex items-center gap-1 shrink-0" style={{ color: step.done ? C.green : C.accentDark }}>{step.cta} <ChevronRight size={13} /></span>
    </button>
  );
}

function TenderDetail({ tender, updateTender, back, onDelete, isAdmin, evaluatorName }) {
  const [tab, setTab] = useState("info");
  const [showCheck, setShowCheck] = useState(false);
  const [confirmingPublish, setConfirmingPublish] = useState(false);
  const progress = computeProgress(tender);
  const completion = tabCompletion(tender);
  const { ready } = readinessChecks(tender);
  const published = PUBLISHED_STATUSES.includes(tender.status);

  function exportAllDocuments() { tender.documents.forEach((doc, i) => setTimeout(() => generateDocumentFile(tender, doc), i * 450)); }
  function publish() {
    updateTender(prev => ({
      ...prev, status: "ongoing",
      history: [...prev.history, { date: new Date().toISOString().slice(0, 16).replace("T", " "), user: "Vous", action: "AO publié — passage à la réception des offres" }],
    }));
    setConfirmingPublish(false);
    setTab("suppliers");
  }

  const StatusIcon = STATUS_ICON[tender.status] || Circle;
  const darkGhostBtn = "flex items-center gap-2 px-3.5 py-2 rounded-lg text-sm font-medium transition-colors text-white hover:bg-white/10";
  const darkGhostStyle = { border: "1px solid rgba(255,255,255,0.18)", backgroundColor: "rgba(255,255,255,0.06)" };

  return (
    <div className="px-4 sm:px-8 py-5 sm:py-7 max-w-6xl">
      <button onClick={back} className="flex items-center gap-1.5 text-sm mb-5 transition-colors hover:opacity-70" style={{ color: C.inkSoft }}><ArrowLeft size={14} /> Retour au tableau de bord</button>

      <div className="rounded-2xl px-6 sm:px-8 py-6 sm:py-7 mb-6 relative overflow-hidden ao-scale-in" style={{ background: `linear-gradient(135deg, ${C.ink}, #263454)` }}>
        <div className="absolute -right-10 -top-16 w-56 h-56 rounded-full opacity-20" style={{ background: `radial-gradient(circle, ${C.accent}, transparent 70%)` }} />
        <div className="relative">
          <div className="flex flex-col lg:flex-row lg:items-start justify-between gap-5">
            <div className="min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-xs font-mono" style={{ color: "#8C97AC" }}>{tender.reference}</span>
                <StatusBadge status={tender.status} />
              </div>
              <h1 className="text-2xl font-semibold text-white mt-1.5 flex items-center gap-2.5">
                <StatusIcon size={20} style={{ color: "#8C97AC" }} className="shrink-0" />
                <span className="truncate">{tender.title}</span>
              </h1>
              <div className="flex items-center gap-3 mt-4 max-w-xs">
                <div className="flex-1"><ProgressBar value={progress} dark /></div>
                <span className="text-sm font-medium tabular-nums shrink-0" style={{ color: "#B7C0D1" }}>{progress}%</span>
              </div>
            </div>
            <div className="flex items-center gap-2 flex-wrap shrink-0">
              {!published && (
                <button onClick={() => setConfirmingPublish(true)} disabled={!ready} title={ready ? "" : "Complétez la checklist de préparation (100%) avant de publier"}
                  className="flex items-center gap-2 px-3.5 py-2 rounded-lg text-sm font-semibold text-white disabled:opacity-40 transition-transform active:scale-[0.98] hover:brightness-110" style={{ backgroundColor: C.green }}>
                  <CheckCircle2 size={14} /> Publier l'AO
                </button>
              )}
              <button onClick={exportAllDocuments} className={darkGhostBtn} style={darkGhostStyle}><Download size={14} /> Exporter le dossier</button>
              <button onClick={() => setShowCheck(s => !s)} className={darkGhostBtn} style={showCheck ? { backgroundColor: C.accent, border: `1px solid ${C.accent}` } : darkGhostStyle}><ShieldCheck size={14} /> {showCheck ? "Masquer le détail" : "Détail de préparation"}</button>
              {isAdmin && (
                <button onClick={() => onDelete(tender)} title="Supprimer cet AO (admin)" className="p-2 rounded-lg transition-colors hover:bg-white/10" style={darkGhostStyle}>
                  <Trash2 size={15} color="#fff" />
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

      <GuidanceBanner tender={tender} onJump={setTab} />
      {showCheck && <QualityCheck t={tender} />}
      <ConfirmDialog open={confirmingPublish} confirmLabel="Publier l'AO"
        title="Publier cet appel d'offres ?"
        message="L'AO passe en phase de réception des offres — le statut devient « En cours » et l'onglet Fournisseurs s'ouvre pour enregistrer les soumissionnaires au fil de la réception. Les informations de cadrage restent modifiables ensuite si besoin."
        onConfirm={publish} onCancel={() => setConfirmingPublish(false)} />

      <div className="mb-6 space-y-4">
        {[
          { phase: 1, label: "Construction de l'AO" },
          { phase: 2, label: "Réception & évaluation" },
          { phase: 0, label: "Suivi" },
        ].map(group => {
          const groupTabs = TABS.filter(tb => tb.phase === group.phase);
          if (!groupTabs.length) return null;
          const locked = group.phase === 2 && !published;
          return (
            <div key={group.phase}>
              <div className="text-[11px] font-semibold uppercase tracking-wider mb-2 flex items-center gap-1.5" style={{ color: C.inkSoft }}>
                {group.label}{locked && <Lock size={11} />}
              </div>
              <div className="flex flex-wrap gap-2">
                {groupTabs.map(tb => {
                  const Icon = tb.icon; const active = tab === tb.key; const done = completion[tb.key];
                  const tabLocked = locked && !active;
                  return (
                    <button key={tb.key} onClick={() => !tabLocked && setTab(tb.key)} disabled={tabLocked}
                      title={tabLocked ? "Disponible une fois l'AO publié" : ""}
                      className="flex items-center gap-1.5 px-3.5 py-2 rounded-full text-sm font-medium transition-all duration-150 disabled:cursor-not-allowed"
                      style={active
                        ? { backgroundColor: C.accent, color: "#fff", border: `1px solid ${C.accent}`, boxShadow: `0 2px 8px ${C.accentSoft}` }
                        : tabLocked
                        ? { backgroundColor: C.borderSoft, color: C.inkSoft, border: `1px solid ${C.borderSoft}`, opacity: 0.6 }
                        : { backgroundColor: C.surface, color: C.inkSoft, border: `1px solid ${C.border}` }}
                      onMouseEnter={e => { if (!tabLocked) e.currentTarget.style.transform = "translateY(-1px)"; }}
                      onMouseLeave={e => { e.currentTarget.style.transform = ""; }}>
                      {tabLocked ? <Lock size={12} /> : <Icon size={14} />} {tb.label}
                      {done && !tabLocked && <CheckCircle2 size={12} style={{ color: active ? "#fff" : C.green }} />}
                    </button>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
      {tab === "info" && <InfoTab t={tender} updateTender={updateTender} />}
      {tab === "need" && <NeedTab t={tender} updateTender={updateTender} />}
      {tab === "documents" && <DocumentsTab t={tender} updateTender={updateTender} />}
      {tab === "requirements" && <RequirementsTab t={tender} updateTender={updateTender} />}
      {tab === "criteria" && <CriteriaTab t={tender} updateTender={updateTender} />}
      {tab === "suppliers" && <SuppliersTab t={tender} updateTender={updateTender} />}
      {tab === "evaluation" && <EvaluationTab t={tender} updateTender={updateTender} evaluatorName={evaluatorName} />}
      {tab === "comparison" && <ComparisonTab t={tender} />}
      {tab === "synthesis" && <SynthesisTab t={tender} />}
      {tab === "history" && <HistoryTab t={tender} />}
    </div>
  );
}

/* --------------------------------- APP --------------------------------- */

const STORAGE_KEY = "tenders";

const APP_URL = typeof window !== "undefined" ? window.location.origin + import.meta.env.BASE_URL : "";

function LoginScreen() {
  const [mode, setMode] = useState("login"); // login | signup | forgot
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [firstName, setFirstName] = useState("");
  const [lastName, setLastName] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");

  async function submit(e) {
    e.preventDefault();
    setBusy(true); setError(""); setInfo("");
    try {
      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
      } else if (mode === "signup") {
        const { error } = await supabase.auth.signUp({
          email, password,
          options: { data: { first_name: firstName.trim(), last_name: lastName.trim() }, emailRedirectTo: APP_URL },
        });
        if (error) throw error;
        setInfo("Compte créé. Vérifiez votre boîte mail pour confirmer votre adresse, puis connectez-vous.");
        setMode("login");
      } else {
        const { error } = await supabase.auth.resetPasswordForEmail(email, { redirectTo: APP_URL });
        if (error) throw error;
        setInfo("Si un compte existe avec cet email, un lien de réinitialisation vient d'être envoyé.");
      }
    } catch (err) {
      setError(err.message || "Une erreur est survenue.");
    } finally {
      setBusy(false);
    }
  }

  const titles = { login: "Connexion", signup: "Créer un compte", forgot: "Mot de passe oublié" };

  return (
    <div className="flex items-center justify-center min-h-screen px-4" style={{ backgroundColor: C.bg }}>
      <div className="w-full max-w-sm p-6 rounded-xl" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, boxShadow: "0 10px 30px rgba(0,0,0,0.08)" }}>
        <div className="mb-2">
          <img src={logoWordmark} alt="AO Manager" className="h-8 w-auto" />
        </div>
        <div className="text-lg font-semibold mb-4" style={{ color: C.ink }}>{titles[mode]}</div>
        <form onSubmit={submit} className="space-y-3">
          {mode === "signup" && (
            <div className="grid grid-cols-2 gap-3">
              <label className="block text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Prénom</div>
                <input required value={firstName} onChange={e => setFirstName(e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
              <label className="block text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Nom</div>
                <input required value={lastName} onChange={e => setLastName(e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
            </div>
          )}
          <label className="block text-sm">
            <div className="mb-1" style={{ color: C.inkSoft }}>Email</div>
            <input type="email" required value={email} onChange={e => setEmail(e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} />
          </label>
          {mode !== "forgot" && (
            <label className="block text-sm">
              <div className="mb-1" style={{ color: C.inkSoft }}>Mot de passe</div>
              <input type="password" required minLength={6} value={password} onChange={e => setPassword(e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} />
            </label>
          )}
          {mode === "login" && (
            <button type="button" onClick={() => { setMode("forgot"); setError(""); setInfo(""); }} className="text-xs" style={{ color: C.accentDark }}>Mot de passe oublié ?</button>
          )}
          {error && <div className="text-xs" style={{ color: C.red }}>{error}</div>}
          {info && <div className="text-xs" style={{ color: C.green }}>{info}</div>}
          <button type="submit" disabled={busy} className="w-full px-4 py-2.5 rounded-full text-sm font-semibold text-white disabled:opacity-50" style={{ backgroundColor: C.accent }}>
            {busy ? "Veuillez patienter…" : mode === "login" ? "Se connecter" : mode === "signup" ? "Créer mon compte" : "Envoyer le lien de réinitialisation"}
          </button>
        </form>
        {mode === "forgot" ? (
          <button onClick={() => { setMode("login"); setError(""); setInfo(""); }} className="w-full text-center text-xs mt-4" style={{ color: C.inkSoft }}>Retour à la connexion</button>
        ) : (
          <button onClick={() => { setMode(m => m === "login" ? "signup" : "login"); setError(""); setInfo(""); }} className="w-full text-center text-xs mt-4" style={{ color: C.inkSoft }}>
            {mode === "login" ? "Pas encore de compte ? Créez-en un" : "Déjà un compte ? Connectez-vous"}
          </button>
        )}
      </div>
    </div>
  );
}

function ResetPasswordScreen({ onDone }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function submit(e) {
    e.preventDefault();
    if (password !== confirm) { setError("Les deux mots de passe ne correspondent pas."); return; }
    setBusy(true); setError("");
    const { error } = await supabase.auth.updateUser({ password });
    setBusy(false);
    if (error) { setError(error.message || "Une erreur est survenue."); return; }
    onDone();
  }

  return (
    <div className="flex items-center justify-center min-h-screen px-4" style={{ backgroundColor: C.bg }}>
      <div className="w-full max-w-sm p-6 rounded-xl" style={{ backgroundColor: C.surface, border: `1px solid ${C.border}`, boxShadow: "0 10px 30px rgba(0,0,0,0.08)" }}>
        <div className="mb-2"><img src={logoWordmark} alt="AO Manager" className="h-8 w-auto" /></div>
        <div className="text-lg font-semibold mb-1" style={{ color: C.ink }}>Choisir un nouveau mot de passe</div>
        <p className="text-xs mb-4" style={{ color: C.inkSoft }}>Suite à votre demande de réinitialisation.</p>
        <form onSubmit={submit} className="space-y-3">
          <label className="block text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Nouveau mot de passe</div>
            <input type="password" required minLength={6} value={password} onChange={e => setPassword(e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
          <label className="block text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Confirmer le mot de passe</div>
            <input type="password" required minLength={6} value={confirm} onChange={e => setConfirm(e.target.value)} className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
          {error && <div className="text-xs" style={{ color: C.red }}>{error}</div>}
          <button type="submit" disabled={busy} className="w-full px-4 py-2.5 rounded-full text-sm font-semibold text-white disabled:opacity-50" style={{ backgroundColor: C.accent }}>
            {busy ? "Veuillez patienter…" : "Mettre à jour le mot de passe"}
          </button>
        </form>
      </div>
    </div>
  );
}

/* --------------------------------- ADMINISTRATION --------------------------------- */

function AdminPanel({ currentUserId }) {
  const [users, setUsers] = useState(null);
  const [error, setError] = useState("");
  const [busyId, setBusyId] = useState(null);
  const [pendingDeleteUser, setPendingDeleteUser] = useState(null);

  async function load() {
    setError("");
    const { data, error } = await supabase.functions.invoke("admin-users", { body: { action: "list" } });
    if (error) { setError(error.message || "Chargement impossible."); return; }
    setUsers(data.users || []);
  }
  useEffect(() => { load(); }, []);

  async function act(action, userId) {
    setBusyId(userId); setError("");
    const { error } = await supabase.functions.invoke("admin-users", { body: { action, userId } });
    setBusyId(null);
    if (error) { setError(error.message || "Action impossible."); return; }
    load();
  }

  return (
    <div className="px-4 sm:px-8 py-5 sm:py-7 max-w-4xl">
      <h1 className="text-xl font-semibold" style={{ color: C.ink }}>Administration</h1>
      <p className="text-sm mt-1 mb-6" style={{ color: C.inkSoft }}>Gestion des comptes — vous voyez et gérez ici tous les utilisateurs de l'outil.</p>

      {error && <div className="text-sm mb-4 px-4 py-2.5 rounded-lg" style={{ backgroundColor: C.redSoft, color: C.red }}>{error}</div>}

      {users === null ? (
        <div className="text-sm" style={{ color: C.inkSoft }}>Chargement…</div>
      ) : (
        <Card className="ao-stagger">
          {users.map((u, i) => (
            <div key={u.id} className="px-5 py-4 flex items-center gap-4" style={{ borderBottom: i < users.length - 1 ? `1px solid ${C.borderSoft}` : "none" }}>
              <div className="flex-1 min-w-0">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-sm font-medium truncate" style={{ color: C.ink }}>{u.email}</span>
                  {u.role === "admin" && <Badge color={C.accentDark} bg={C.accentSoft}>Admin</Badge>}
                  {u.revoked && <Badge color={C.red} bg={C.redSoft}>Accès révoqué</Badge>}
                  {u.id === currentUserId && <span className="text-xs" style={{ color: C.inkSoft }}>(vous)</span>}
                </div>
                <div className="text-xs mt-1" style={{ color: C.inkSoft }}>{u.tenderCount} appel{u.tenderCount > 1 ? "s" : ""} d'offres</div>
              </div>
              {u.id !== currentUserId && (
                <div className="flex items-center gap-2 shrink-0">
                  <GhostButton disabled={busyId === u.id} onClick={() => act(u.revoked ? "restore" : "revoke", u.id)}>
                    {u.revoked ? "Restaurer l'accès" : "Révoquer l'accès"}
                  </GhostButton>
                  <button onClick={() => setPendingDeleteUser(u)} disabled={busyId === u.id} title="Supprimer ce compte" className="p-2 rounded-lg transition-colors hover:bg-black/5 disabled:opacity-40" style={{ border: `1px solid ${C.border}` }}>
                    <Trash2 size={15} style={{ color: C.red }} />
                  </button>
                </div>
              )}
            </div>
          ))}
          {users.length === 0 && <div className="px-5 py-6 text-sm" style={{ color: C.inkSoft }}>Aucun utilisateur.</div>}
        </Card>
      )}

      <ConfirmDialog open={!!pendingDeleteUser} danger confirmLabel="Supprimer définitivement"
        title="Supprimer ce compte ?"
        message={pendingDeleteUser ? `${pendingDeleteUser.email} et tous ses appels d'offres (${pendingDeleteUser.tenderCount}) seront définitivement supprimés. Cette action est irréversible.` : ""}
        onConfirm={() => { const u = pendingDeleteUser; setPendingDeleteUser(null); act("delete", u.id); }}
        onCancel={() => setPendingDeleteUser(null)} />
    </div>
  );
}

/* --------------------------------- REGISTRE FOURNISSEURS --------------------------------- */

function SupplierRegistryPage({ isAdmin }) {
  const [suppliers, setSuppliers] = useState(null);
  const [error, setError] = useState("");
  const [form, setForm] = useState({ name: "", email: "", domain: "", region: "" });
  const [query, setQuery] = useState("");
  const [domainFilter, setDomainFilter] = useState("");
  const [regionFilter, setRegionFilter] = useState("");
  const [importing, setImporting] = useState(false);
  const [importSummary, setImportSummary] = useState("");

  async function load() {
    const { data, error } = await supabase.from("suppliers_registry").select("id, name, email, domain, region, created_at").order("name");
    if (error) { setError(error.message); return; }
    setSuppliers(data || []);
  }
  useEffect(() => { load(); }, []);

  async function addManual() {
    if (!form.name.trim()) return;
    const { error } = await supabase.from("suppliers_registry").insert({ name: form.name.trim(), email: form.email.trim() || null, domain: form.domain.trim() || null, region: form.region.trim() || null });
    if (error) { setError(error.message); return; }
    setForm({ name: "", email: "", domain: "", region: "" }); setError("");
    load();
  }

  async function remove(id) {
    const { error } = await supabase.from("suppliers_registry").delete().eq("id", id);
    if (error) { setError(error.message); return; }
    load();
  }

  async function handleFile(e) {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    setImporting(true); setError(""); setImportSummary("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      const existingEmails = new Set((suppliers || []).map(s => (s.email || "").toLowerCase()).filter(Boolean));
      const toInsert = [];
      let skipped = 0;
      for (const row of rows) {
        const keys = Object.keys(row);
        const nameKey = keys.find(k => /nom|name|fournisseur/i.test(k)) ?? keys[0];
        const emailKey = keys.find(k => /mail/i.test(k)) ?? keys[1];
        const domainKey = keys.find(k => /m[ée]tier|domaine|activit/i.test(k));
        const regionKey = keys.find(k => /r[ée]gion|zone|canton|ville/i.test(k));
        const name = String(row[nameKey] ?? "").trim();
        const email = String(row[emailKey] ?? "").trim();
        if (!name) { skipped++; continue; }
        if (email && existingEmails.has(email.toLowerCase())) { skipped++; continue; }
        if (email) existingEmails.add(email.toLowerCase());
        toInsert.push({ name, email: email || null, domain: domainKey ? String(row[domainKey] ?? "").trim() || null : null, region: regionKey ? String(row[regionKey] ?? "").trim() || null : null });
      }
      if (toInsert.length) {
        const { error } = await supabase.from("suppliers_registry").insert(toInsert);
        if (error) throw error;
      }
      setImportSummary(`${toInsert.length} fournisseur${toInsert.length > 1 ? "s" : ""} importé${toInsert.length > 1 ? "s" : ""}${skipped ? `, ${skipped} ligne${skipped > 1 ? "s" : ""} ignorée${skipped > 1 ? "s" : ""} (doublon ou nom manquant)` : ""}.`);
      load();
    } catch (err) {
      setError(`Import échoué (${err.message || "erreur inconnue"}).`);
    } finally { setImporting(false); }
  }

  const domains = [...new Set((suppliers || []).map(s => s.domain).filter(Boolean))].sort();
  const regions = [...new Set((suppliers || []).map(s => s.region).filter(Boolean))].sort();
  const filtered = (suppliers || []).filter(s =>
    (!query.trim() || s.name.toLowerCase().includes(query.toLowerCase()) || (s.email || "").toLowerCase().includes(query.toLowerCase())) &&
    (!domainFilter || s.domain === domainFilter) && (!regionFilter || s.region === regionFilter));

  return (
    <div className="px-4 sm:px-8 py-5 sm:py-7 max-w-4xl">
      <h1 className="text-xl font-semibold" style={{ color: C.ink }}>Fournisseurs (registre)</h1>
      <p className="text-sm mt-1 mb-6" style={{ color: C.inkSoft }}>
        Alimenté automatiquement à chaque fournisseur ajouté dans un appel d'offres{isAdmin ? " — vous pouvez aussi compléter la liste manuellement ou en importer une." : "."}
      </p>

      {error && <div className="text-sm mb-4 px-4 py-2.5 rounded-lg" style={{ backgroundColor: C.redSoft, color: C.red }}>{error}</div>}
      {importSummary && <div className="text-sm mb-4 px-4 py-2.5 rounded-lg" style={{ backgroundColor: C.greenSoft, color: C.green }}>{importSummary}</div>}

      {isAdmin && (
        <Card className="p-5 mb-5">
          <SectionTitle>Ajouter des fournisseurs</SectionTitle>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
            <label className="text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Nom</div>
              <input value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
            <label className="text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Email</div>
              <input value={form.email} onChange={e => setForm(f => ({ ...f, email: e.target.value }))} placeholder="contact@fournisseur.example" className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
            <label className="text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Métier / domaine</div>
              <input value={form.domain} onChange={e => setForm(f => ({ ...f, domain: e.target.value }))} placeholder="Ex. Infrastructure, Développement…" className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
            <label className="text-sm"><div className="mb-1" style={{ color: C.inkSoft }}>Région / zone</div>
              <input value={form.region} onChange={e => setForm(f => ({ ...f, region: e.target.value }))} placeholder="Ex. Genève, Suisse romande…" className="w-full px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} /></label>
          </div>
          <div className="flex flex-wrap items-center gap-3">
            <PrimaryButton icon={Plus} onClick={addManual}>Ajouter</PrimaryButton>
            <label className="inline-flex items-center gap-2 text-sm font-medium px-3.5 py-2 rounded-full cursor-pointer transition-colors hover:brightness-95" style={{ border: `1px solid ${C.accent}`, color: C.accentDark, backgroundColor: C.accentSoft }}>
              <Upload size={14} /> {importing ? "Import en cours…" : "Importer un fichier Excel"}
              <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile} disabled={importing} />
            </label>
          </div>
          <div className="text-xs mt-2" style={{ color: C.inkSoft }}>Fichier .xlsx / .xls / .csv — colonnes nom et email obligatoires ; métier/domaine et région/zone optionnels (en-têtes libres).</div>
        </Card>
      )}

      <div className="flex flex-wrap gap-3 mb-4">
        <input value={query} onChange={e => setQuery(e.target.value)} placeholder="Rechercher un fournisseur…" className="flex-1 min-w-[200px] px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }} />
        {domains.length > 0 && (
          <select value={domainFilter} onChange={e => setDomainFilter(e.target.value)} className="px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }}>
            <option value="">Tous les métiers</option>
            {domains.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
        )}
        {regions.length > 0 && (
          <select value={regionFilter} onChange={e => setRegionFilter(e.target.value)} className="px-3 py-2 rounded text-sm outline-none" style={{ border: `1px solid ${C.border}` }}>
            <option value="">Toutes les régions</option>
            {regions.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        )}
      </div>

      {suppliers === null ? (
        <div className="text-sm" style={{ color: C.inkSoft }}>Chargement…</div>
      ) : (
        <Card className="ao-stagger">
          {filtered.length === 0 && <div className="px-5 py-8 text-sm text-center" style={{ color: C.inkSoft }}>Aucun fournisseur{query || domainFilter || regionFilter ? " ne correspond à la recherche" : " enregistré pour l'instant — il se remplira au fil des AO"}.</div>}
          {filtered.map((s, i) => (
            <div key={s.id} className="px-5 py-3.5 flex items-center gap-4" style={{ borderBottom: i < filtered.length - 1 ? `1px solid ${C.borderSoft}` : "none" }}>
              <div className="flex-1 min-w-0">
                <div className="text-sm font-medium" style={{ color: C.ink }}>{s.name}</div>
                <div className="text-xs mt-0.5 flex items-center gap-2 flex-wrap" style={{ color: C.inkSoft }}>
                  <span>{s.email || "—"}</span>
                  {s.domain && <Badge color={C.accentDark} bg={C.accentSoft}>{s.domain}</Badge>}
                  {s.region && <Badge color={C.inkSoft} bg={C.slateSoft}>{s.region}</Badge>}
                </div>
              </div>
              {isAdmin && <button onClick={() => remove(s.id)} title="Supprimer" className="p-1.5 rounded hover:bg-black/5 transition-colors"><Trash2 size={14} style={{ color: C.inkSoft }} /></button>}
            </div>
          ))}
        </Card>
      )}
    </div>
  );
}

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = pas encore vérifié, null = pas connecté
  const [tenders, setTenders] = useState([]);
  const [view, setView] = useState("dashboard");
  const [selectedId, setSelectedId] = useState(null);
  const isMobile = useIsMobile();
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const [saveError, setSaveError] = useState(false);
  const [pendingDelete, setPendingDelete] = useState(null);
  const syncTimers = React.useRef({});

  // Authentification : récupère la session en cours et réagit aux connexions/déconnexions.
  // Un clic sur le lien de réinitialisation de mot de passe ouvre une session "recovery" —
  // on l'intercepte pour afficher l'écran de choix du nouveau mot de passe avant l'app.
  const [passwordRecovery, setPasswordRecovery] = useState(false);
  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session ?? null));
    const { data: sub } = supabase.auth.onAuthStateChange((event, s) => {
      if (event === "PASSWORD_RECOVERY") setPasswordRecovery(true);
      setSession(s);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  // Rôle de l'utilisateur connecté (admin voit/gère tous les AO et utilisateurs, sinon uniquement les siens).
  const [profile, setProfile] = useState(null);
  useEffect(() => {
    if (!session) { setProfile(null); return; }
    let cancelled = false;
    supabase.from("profiles").select("role, email, first_name, last_name").eq("id", session.user.id).single()
      .then(({ data }) => { if (!cancelled) setProfile(data || { role: "user" }); });
    return () => { cancelled = true; };
  }, [session]);
  const isAdmin = profile?.role === "admin";

  // Chargement des AO : la sécurité par ligne garantit qu'un utilisateur ne voit que les siens,
  // et qu'un admin voit ceux de tout le monde. On garde le propriétaire de côté (jamais dans
  // les objets AO eux-mêmes) pour ne jamais réinitialiser/supprimer par erreur les AO d'autrui.
  const [ownerByTenderId, setOwnerByTenderId] = useState({});
  useEffect(() => {
    if (!session) { setTenders([]); setOwnerByTenderId({}); setLoaded(session === null); return; }
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase.from("tenders").select("id, data, user_id").order("created_at", { ascending: true });
      if (!cancelled) {
        if (!error && data) {
          setTenders(data.map(row => ({ ...row.data, id: row.id })));
          setOwnerByTenderId(Object.fromEntries(data.map(row => [row.id, row.user_id])));
        }
        setLoaded(true);
      }
    })();
    return () => { cancelled = true; };
  }, [session]);

  // Sauvegarde d'un AO dans Supabase, avec un léger anti-rebond pour éviter un appel réseau à chaque frappe.
  function syncTenderToSupabase(tender) {
    clearTimeout(syncTimers.current[tender.id]);
    syncTimers.current[tender.id] = setTimeout(async () => {
      const { error } = await supabase.from("tenders")
        .update({ reference: tender.reference, title: tender.title, data: tender })
        .eq("id", tender.id);
      setSaveError(!!error);
    }, 500);
  }

  function openTender(id) { setSelectedId(id); setView("detail"); }
  function updateTender(updater) {
    setTenders(prev => {
      const next = prev.map(t => t.id === selectedId ? (typeof updater === "function" ? updater(t) : updater) : t);
      const changed = next.find(t => t.id === selectedId);
      if (changed) syncTenderToSupabase(changed);
      return next;
    });
  }
  async function createTender(t) {
    const id = crypto.randomUUID();
    const withId = { ...t, id };
    setTenders(prev => [...prev, withId]);
    setSelectedId(id); setView("detail");
    const { error } = await supabase.from("tenders").insert({ id, user_id: session.user.id, reference: withId.reference, title: withId.title, data: withId });
    setSaveError(!!error);
  }
  function deleteTender(t) { setPendingDelete(t); }
  async function confirmDeleteTender() {
    const t = pendingDelete;
    setTenders(prev => prev.filter(x => x.id !== t.id));
    if (selectedId === t.id) { setSelectedId(null); setView("dashboard"); }
    setPendingDelete(null);
    await supabase.from("tenders").delete().eq("id", t.id);
  }
  const selected = tenders.find(t => t.id === selectedId);

  if (session === undefined) {
    return <div className="flex items-center justify-center min-h-screen text-sm" style={{ backgroundColor: C.bg, color: C.inkSoft }}>Chargement…</div>;
  }
  if (passwordRecovery) {
    return <ResetPasswordScreen onDone={() => setPasswordRecovery(false)} />;
  }
  if (session === null) {
    return <LoginScreen />;
  }
  if (!loaded) {
    return <div className="flex items-center justify-center min-h-screen text-sm" style={{ backgroundColor: C.bg, color: C.inkSoft }}>Chargement de vos appels d'offres…</div>;
  }

  return (
    <div className="flex min-h-screen" style={{ backgroundColor: C.bg, fontFamily: "Inter, ui-sans-serif, system-ui, -apple-system, sans-serif" }}>
      <Sidebar view={view === "detail" ? "" : view} setView={v => { setView(v); setSelectedId(null); }} isMobile={isMobile} open={drawerOpen} onClose={() => setDrawerOpen(false)} onLogout={() => supabase.auth.signOut()} userEmail={session.user.email} isAdmin={isAdmin} />
      <div className="flex-1 min-w-0">
        <TopBar isMobile={isMobile} onMenuClick={() => setDrawerOpen(true)} userEmail={session.user.email} crumbs={view === "dashboard" ? ["Tableau de bord"] : view === "new" ? ["Tableau de bord", "Nouvel appel d'offres"] : view === "admin" ? ["Administration"] : view === "suppliers-registry" ? ["Fournisseurs (registre)"] : ["Tableau de bord", selected?.reference || ""]} />
        {saveError && <div className="text-xs text-center py-1.5" style={{ backgroundColor: C.redSoft, color: C.red }}>La sauvegarde automatique a échoué pour la dernière modification — vos données restent visibles ici, mais pourraient ne pas persister après fermeture.</div>}
        <div key={view === "detail" ? `detail-${selectedId}` : view} className="ao-view-enter">
          {view === "dashboard" && <Dashboard tenders={tenders} openTender={openTender} goNew={() => setView("new")} onDelete={deleteTender} userEmail={session.user.email} firstName={profile?.first_name} isAdmin={isAdmin} />}
          {view === "new" && <NewTenderWizard onCreate={createTender} onCancel={() => setView("dashboard")} />}
          {view === "detail" && selected && <TenderDetail tender={selected} updateTender={updateTender} back={() => setView("dashboard")} onDelete={deleteTender} isAdmin={isAdmin} evaluatorName={(profile?.first_name && profile?.last_name) ? `${profile.first_name} ${profile.last_name}` : (profile?.first_name || session.user.email)} />}
          {view === "admin" && isAdmin && <AdminPanel currentUserId={session.user.id} />}
          {view === "suppliers-registry" && <SupplierRegistryPage isAdmin={isAdmin} />}
        </div>
      </div>
      <ConfirmDialog open={!!pendingDelete} danger confirmLabel="Supprimer"
        title="Supprimer cet appel d'offres ?"
        message={pendingDelete ? `${pendingDelete.reference} — ${pendingDelete.title} sera définitivement supprimé. Cette action est irréversible.` : ""}
        onConfirm={confirmDeleteTender} onCancel={() => setPendingDelete(null)} />
    </div>
  );
}
