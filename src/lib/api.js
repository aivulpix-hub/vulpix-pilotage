// Persistance locale (navigateur) — simple, gratuite, sans backend.
// Deux filets de sécurité :
//  1. un instantané automatique par jour, les 7 derniers jours sont conservés
//     (protège contre une erreur de saisie ou une donnée corrompue) ;
//  2. la date du dernier export JSON, pour rappeler de sauvegarder hors du navigateur
//     (seule protection contre un vidage de cache ou un changement d'appareil).
const STORAGE_KEY = "vulpix-pilotage-data-v1";
const SNAPSHOTS_KEY = "vulpix-pilotage-snapshots-v1";
const LAST_EXPORT_KEY = "vulpix-pilotage-last-export";
const MAX_SNAPSHOTS = 7;

const EMPTY = { clients: [], sessions: [], mrrHistory: [], reminders: [], settings: {} };

export async function loadData() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return EMPTY;
  try {
    return JSON.parse(raw);
  } catch (e) {
    throw new Error("Données locales illisibles. Restaure un instantané dans Réglages.");
  }
}

export async function saveData(data) {
  const json = JSON.stringify(data);
  localStorage.setItem(STORAGE_KEY, json);
  writeDailySnapshot(json);
  return { ok: true };
}

function readSnapshots() {
  try {
    return JSON.parse(localStorage.getItem(SNAPSHOTS_KEY) || "{}");
  } catch (e) {
    return {};
  }
}

function writeDailySnapshot(json) {
  const today = new Date().toISOString().slice(0, 10);
  const snaps = readSnapshots();
  if (snaps[today]) return; // un seul instantané par jour : le premier état de la journée
  snaps[today] = json;
  const days = Object.keys(snaps).sort();
  while (days.length > MAX_SNAPSHOTS) delete snaps[days.shift()];
  try {
    localStorage.setItem(SNAPSHOTS_KEY, JSON.stringify(snaps));
  } catch (e) {
    // quota plein : on garde au moins les données principales
  }
}

export function listSnapshots() {
  const snaps = readSnapshots();
  return Object.keys(snaps)
    .sort()
    .reverse()
    .map((day) => {
      let sessions = 0, clients = 0;
      try {
        const d = JSON.parse(snaps[day]);
        sessions = (d.sessions || []).length;
        clients = (d.clients || []).length;
      } catch (e) { /* instantané illisible */ }
      return { day, sessions, clients };
    });
}

export function readSnapshot(day) {
  const snaps = readSnapshots();
  return snaps[day] ? JSON.parse(snaps[day]) : null;
}

export function getLastExport() {
  const v = localStorage.getItem(LAST_EXPORT_KEY);
  return v ? new Date(v) : null;
}

export function markExported() {
  localStorage.setItem(LAST_EXPORT_KEY, new Date().toISOString());
}

export function downloadBackup(data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `vulpix-sauvegarde-${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  markExported();
}
