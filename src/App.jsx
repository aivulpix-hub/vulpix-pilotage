import React, { useState, useEffect, useMemo, useCallback } from "react";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell,
} from "recharts";
import * as XLSX from "xlsx";
import {
  Plus, X, Trash2, Clock, TrendingUp, Users, Settings as SettingsIcon,
  Download, ChevronRight, Bell, Check, Gauge as GaugeIcon, Wallet,
} from "lucide-react";
import { loadData, saveData, downloadBackup, getLastExport, listSnapshots, readSnapshot } from "./lib/api.js";

/* ============================== TOKENS ============================== */
const C = {
  // surfaces
  bg: "#F1F2F4",        // fond d'application, gris froid très clair
  card: "#FFFFFF",      // surfaces
  sunken: "#F6F7F9",    // lignes et zones internes aux surfaces
  line: "#E3E5EA",      // filets
  // texte
  ink: "#16182B",       // navy Vulpix, texte principal
  ink2: "#3B4054",
  gray: "#6A6F80",
  faint: "#9A9FAD",
  // marque
  orange: "#FF8C00",    // réservé aux marqueurs importants (onglet actif, progression, "en cours")
  orangeDeep: "#E67300",
  orangeText: "#A94F00",
  orangeBg: "#FFF3E3",
  // zones de performance (jauges, taux horaires)
  zRed: "#D64541",
  zOrange: "#F08A24",
  zGreen: "#2F9E5E",
  zBlue: "#3A6BD6",
  // statuts
  green: "#1F7A45",
  greenBg: "#E7F4EC",
  red: "#B8322C",
  redBg: "#FBEAE9",
  blue: "#2553C9",
  blueBg: "#E8EEFB",
};

const FONT = "'Geist', -apple-system, 'Segoe UI', Roboto, sans-serif";
const MONO = FONT; // chiffres : même police, en chiffres tabulaires (réglé globalement)

/* ============================== HELPERS ============================== */
const uid = () => Math.random().toString(36).slice(2, 10);
const euro = (n) =>
  (n || 0).toLocaleString("fr-FR", { maximumFractionDigits: 0 }) + "€";
const euro1 = (n) =>
  (n || 0).toLocaleString("fr-FR", { maximumFractionDigits: 1 }) + "€";
const monthKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
const monthLabel = (key) => {
  const [y, m] = key.split("-");
  const names = ["Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Août", "Sep", "Oct", "Nov", "Déc"];
  return `${names[parseInt(m, 10) - 1]} ${y.slice(2)}`;
};
const daysBetween = (a, b) => Math.floor((b - a) / (1000 * 60 * 60 * 24));
const todayISO = () => new Date().toISOString().slice(0, 10);

/* ============================== SEED DATA ============================== */
const seedClients = () => [
  {
    id: uid(), name: "SCL / Sporting Club Luberon", sector: "Association sportive", zone: "Lauris",
    mrr: 209, status: "actif", payment: "virement", startDate: "2026-01-01",
    postsPerWeek: 3, lastContact: "", source: "", notes: "Fabien Guigues (DS), Nicolas Lespagnol (président)",
    upgrades: [], echeances: [], signal: "Florent a évoqué une possible hausse de fréquence de publication",
  },
  {
    id: uid(), name: "Jacky / JB Couverture", sector: "Couvreur", zone: "Aix-en-Provence / Corse",
    mrr: 149, status: "actif", payment: "virement", startDate: "2026-01-01",
    postsPerWeek: 3, lastContact: "", source: "", notes: "Déploiement Balagne puis Pays d'Aix",
    upgrades: [], echeances: [],
  },
  {
    id: uid(), name: "Loïc Champon (Help Services / Miroiterie / Trako)", sector: "Multi-entités BTP", zone: "Pertuis",
    mrr: 159, status: "actif", payment: "virement", startDate: "2026-01-01",
    postsPerWeek: 3, lastContact: "", source: "", notes: "",
    upgrades: [{ id: uid(), montant: 259, date: todayISO(), statut: "en attente", label: "Miroiterie Pertuisienne seule" }],
    echeances: [],
  },
  {
    id: uid(), name: "Jordan / JBM Rénovation", sector: "Couvreur/Rénovation", zone: "Lauris–Pertuis",
    mrr: 99, status: "actif", payment: "stripe", startDate: "2026-05-18",
    postsPerWeek: 2, lastContact: "", source: "", notes: "Toitu'rénov by JBM",
    upgrades: [], echeances: [{ id: uid(), label: "Fin période ferme 3 mois", date: "2026-08-18" }],
  },
  {
    id: uid(), name: "Yohann Serrurerie", sector: "Serrurier", zone: "Port-de-Bouc",
    mrr: 99, status: "actif", payment: "stripe", startDate: "2026-02-01",
    postsPerWeek: 2, lastContact: "", source: "", notes: "Label Serrurier de France, A2P",
    upgrades: [], echeances: [],
  },
  {
    id: uid(), name: "TMP / Philippe Fantini", sector: "Menuiserie PVC", zone: "Marseille",
    mrr: 59, status: "actif", payment: "stripe", startDate: "2026-01-01",
    postsPerWeek: 1, lastContact: "", source: "", notes: "Dossier Leasecom en négociation",
    upgrades: [{ id: uid(), montant: 149, date: todayISO(), statut: "en attente", label: "Upgrade proposé" }],
    echeances: [],
  },
  {
    id: uid(), name: "Charton Paysage & Piscine", sector: "Paysagiste", zone: "PACA",
    mrr: 59, status: "actif", payment: "virement", startDate: "2026-01-01",
    postsPerWeek: 2, lastContact: "", source: "", notes: "Facture manuelle le 1er du mois",
    upgrades: [{ id: uid(), montant: 99, date: todayISO(), statut: "en attente", label: "Upgrade proposé" }],
    echeances: [],
  },
  {
    id: uid(), name: "Sanchez", sector: "BTP", zone: "PACA",
    mrr: 59, status: "standby", payment: "virement", startDate: "2026-01-01",
    postsPerWeek: 0, lastContact: "", source: "", notes: "Client difficile, forte exigence pour peu d'investissement",
    upgrades: [], echeances: [],
  },
  {
    id: uid(), name: "Nexus / Antoine Baudon", sector: "Consultance métallerie-serrurerie B2B", zone: "PACA",
    mrr: 0, status: "prospect", payment: "", startDate: "", postsPerWeek: 0,
    lastContact: "", source: "Contact professionnel", notes: "GEO comme différenciateur B2B",
    upgrades: [{ id: uid(), montant: 199, date: todayISO(), statut: "en attente", label: "149€ ou 249€ selon scope site" }],
    echeances: [],
  },
  {
    id: uid(), name: "Pedro Buch / PB Toiture", sector: "Couvreur", zone: "Aix-en-Provence",
    mrr: 0, status: "prospect", payment: "", startDate: "", postsPerWeek: 0,
    lastContact: "", source: "Introduit par Jacky (beau-père)", notes: "Bloqué en contrat Linkeo 48 mois — clause Loi Hamon en piste",
    upgrades: [], echeances: [{ id: uid(), label: "Milestone légal / sortie contrat mois 13", date: "2027-06-01" }],
  },
];

const defaultSettings = {
  tjmSeuils: [25, 40, 60],
  mrrObjectifPrincipal: 1500,
  mrrObjectifPhase1: 1000,
  urssafRate: 23.1,
  coutsOutils: 34.98,
  stripeFeeRate: 2,
  tarifReferenceLivrable: 45,
  poidsCategories: { ads: 10, whatsapp: 10, gmb: 10, maintenance: 5, rapports: 5, autre: 0, posts: 60 },
  sessionsParMoisMaintenance: { ads: 4, whatsapp: 4, gmb: 2, maintenance: 2, rapports: 1, autre: 2 },
  rcProSouscrite: false,
  // Seuils fiscaux (CA annuel, en €) — à confirmer avec un expert-comptable
  caRepereLiberatoire: 29000,
  caPlafondMax: 33000,
  seuilTva: 37500,
};

const CATEGORIES = [
  { key: "posts", label: "Posts (réseaux)" },
  { key: "ads", label: "Ads (suivi/optim)" },
  { key: "whatsapp", label: "WhatsApp / photos" },
  { key: "gmb", label: "GMB / avis" },
  { key: "maintenance", label: "Maintenance site web" },
  { key: "rapports", label: "Rapports mensuels" },
  { key: "geoposts", label: "Google Posts géolocalisés" },
  { key: "livrable", label: "Site web / Livrable" },
  { key: "autre", label: "Autre" },
  { key: "rdv", label: "RDV / Visio (hors TJM)" },
];

/* ============================== STORAGE (localStorage via lib/api.js) ============================== */
function useVulpixData() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [lastSaved, setLastSaved] = useState(null);
  const [error, setError] = useState(null);
  const [saving, setSaving] = useState(false);

  const fetchData = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const loaded = await loadData();
      const isEmpty = !loaded.clients || loaded.clients.length === 0;
      const merged = {
        clients: isEmpty ? seedClients() : loaded.clients,
        sessions: loaded.sessions || [],
        mrrHistory: loaded.mrrHistory && loaded.mrrHistory.length > 0 ? loaded.mrrHistory : [{ month: monthKey(new Date()), mrr: 813 }],
        reminders: loaded.reminders || [],
        caHistory: loaded.caHistory || [],
        settings: {
          ...defaultSettings,
          ...loaded.settings,
          poidsCategories: { ...defaultSettings.poidsCategories, ...(loaded.settings?.poidsCategories || {}) },
          sessionsParMoisMaintenance: { ...defaultSettings.sessionsParMoisMaintenance, ...(loaded.settings?.sessionsParMoisMaintenance || {}) },
        },
      };
      // Migration : anciens repères par défaut (28 000 / 34 000) remplacés par les nouveaux (29 000 / 33 000)
      if (merged.settings.caRepereLiberatoire === 28000 && merged.settings.caPlafondMax === 34000) {
        merged.settings.caRepereLiberatoire = 29000;
        merged.settings.caPlafondMax = 33000;
        await saveData(merged);
      }
      setData(merged);
      if (isEmpty) await saveData(merged);
    } catch (e) {
      setError(e.message || "Erreur de chargement");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { fetchData(); }, [fetchData]);

  const save = useCallback(async (next) => {
    setData(next);
    setSaving(true);
    try {
      await saveData(next);
      setLastSaved(new Date());
      setError(null);
    } catch (e) {
      setError(e.message || "Échec de la sauvegarde");
    } finally {
      setSaving(false);
    }
  }, []);

  return { data, loading, save, lastSaved, error, saving, refetch: fetchData };
}

/* ============================== GAUGE ============================== */
function Gauge({ value, max, zones, label, sublabel, unit = "€" }) {
  const pct = Math.max(0, Math.min(100, (value / max) * 100));
  return (
    <div style={{ flex: 1, minWidth: 220 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 8 }}>
        <div style={{ fontSize: 11, fontWeight: 600, color: C.gray }}>{label}</div>
        <div style={{ fontSize: 11, color: C.gray }}>{sublabel}</div>
      </div>
      <div style={{ fontSize: 30, fontWeight: 600, color: C.ink, marginBottom: 10 }}>
        {unit === "€" ? euro(value) : euro1(value) + "/h"}
      </div>
      <div style={{ position: "relative", height: 10, borderRadius: 6, overflow: "hidden", background: "#eee", display: "flex" }}>
        {zones.map((z, i) => (
          <div key={i} style={{ width: `${z.width}%`, background: z.color }} />
        ))}
        <div
          style={{
            position: "absolute", left: `calc(${pct}% - 2px)`, top: -3, width: 4, height: 16,
            borderRadius: 2, background: C.ink, boxShadow: "0 0 0 2px #fff",
          }}
        />
      </div>
    </div>
  );
}

function zoneColorForTJM(value, seuils) {
  if (value < seuils[0]) return { color: C.zRed, label: "Rouge" };
  if (value < seuils[1]) return { color: C.zOrange, label: "Orange" };
  if (value < seuils[2]) return { color: C.zGreen, label: "Vert" };
  return { color: C.zBlue, label: "Bleu" };
}

function computeMonthStats(data, monthStr) {
  const { clients, sessions, settings } = data;
  const activeClients = clients.filter((c) => c.status === "actif");
  const mrrCourant = activeClients.reduce((s, c) => s + (Number(c.mrr) || 0), 0);
  // Un mois passé est évalué avec le MRR figé de ce mois (historique), pas avec le MRR d'aujourd'hui.
  const figé = (data.mrrHistory || []).find((h) => h.month === monthStr);
  const isCurrent = monthStr === monthKey(new Date());
  const mrrActif = !isCurrent && figé ? Number(figé.mrr) || 0 : mrrCourant;
  const mrrSource = isCurrent ? "courant" : figé ? "figé" : "courant (pas d'historique pour ce mois)";
  const monthSessions = sessions.filter((s) => s.date && s.date.startsWith(monthStr));
  const prodSessions = monthSessions.filter((s) => {
    if (s.category === "rdv") return false;
    if (s.category === "livrable" && (s.value || 0) === 0) return false;
    if (!s.clients || s.clients.length === 0) return true;
    return s.clients.some((sc) => {
      const cl = clients.find((x) => x.id === sc.id);
      return cl && cl.status === "actif";
    });
  });
  const totalMinutes = prodSessions.reduce((s, x) => s + (x.duration || 0), 0);
  const totalHours = totalMinutes / 60;
  const tjmHoraire = totalHours > 0 ? mrrActif / totalHours : 0;
  const rdvMinutes = monthSessions.filter((s) => s.category === "rdv").reduce((s, x) => s + (x.duration || 0), 0);
  const allMinutes = monthSessions.reduce((s, x) => s + (x.duration || 0), 0);
  const investMinutes = allMinutes - totalMinutes - rdvMinutes;

  const byCategory = {};
  prodSessions.forEach((s) => {
    if (!byCategory[s.category]) byCategory[s.category] = { minutes: 0, value: 0 };
    byCategory[s.category].minutes += s.duration || 0;
    byCategory[s.category].value += s.value || 0;
  });
  const categoryBreakdown = Object.entries(byCategory)
    .map(([key, v]) => ({
      key,
      label: (CATEGORIES.find((c) => c.key === key) || {}).label || key,
      hours: v.minutes / 60,
      value: v.value,
      rate: v.minutes > 0 ? v.value / (v.minutes / 60) : 0,
    }))
    .sort((a, b) => b.hours - a.hours);

  return { monthSessions, totalHours, tjmHoraire, investMinutes, rdvMinutes, allMinutes, mrrActif, mrrSource, categoryBreakdown };
}

/* ============================== RADIAL GAUGE ============================== */
function polarToCartesian(cx, cy, r, angleDeg) {
  const a = (angleDeg * Math.PI) / 180;
  return { x: cx + r * Math.cos(a), y: cy - r * Math.sin(a) };
}
function arcPath(cx, cy, r, angleStart, angleEnd) {
  const start = polarToCartesian(cx, cy, r, angleStart);
  const end = polarToCartesian(cx, cy, r, angleEnd);
  const largeArc = Math.abs(angleStart - angleEnd) > 180 ? 1 : 0;
  return `M ${start.x} ${start.y} A ${r} ${r} 0 ${largeArc} 1 ${end.x} ${end.y}`;
}

function RadialGauge({ value, max, segments, label, sublabel, displayValue }) {
  const cx = 110, cy = 112, r = 88, sw = 11, gap = 3;
  const clamped = Math.max(0, Math.min(max, value));
  const needleAngle = 180 - (clamped / max) * 180;
  const needleTip = polarToCartesian(cx, cy, r - 16, needleAngle);

  return (
    <div style={{ flex: "1 1 240px", minWidth: 0 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", gap: 8 }}>
        <div className="vx-g-label" style={{ fontSize: 13, fontWeight: 500, color: C.gray }}>{label}</div>
        <div className="vx-g-sub" style={{ fontSize: 12, color: C.faint, textAlign: "right" }}>{sublabel}</div>
      </div>
      <svg viewBox="0 0 220 122" role="img" aria-label={`${label} : ${displayValue}`}
        style={{ width: "100%", maxWidth: 280, display: "block", margin: "10px auto 0" }}>
        {segments.map((seg, k) => {
          const a0 = 180 - (seg.from / max) * 180;
          const a1 = 180 - (seg.to / max) * 180;
          const start = Math.min(180, a0 - (seg.from > 0 ? gap / 2 : 0));
          const end = Math.max(0, a1 + (seg.to < max ? gap / 2 : 0));
          if (start <= end) return null;
          return <path key={k} d={arcPath(cx, cy, r, start, end)} fill="none" stroke={seg.color} strokeWidth={sw} strokeLinecap="butt" />;
        })}
        <line x1={cx} y1={cy} x2={needleTip.x} y2={needleTip.y} stroke={C.ink} strokeWidth={2.5} strokeLinecap="round" />
        <circle cx={cx} cy={cy} r={6} fill={C.ink} />
        <circle cx={cx} cy={cy} r={2.2} fill={C.card} />
      </svg>
      <div className="vx-g-value" style={{ textAlign: "center", fontSize: 34, fontWeight: 600, letterSpacing: "-0.03em", color: C.ink, marginTop: 2 }}>{displayValue}</div>
    </div>
  );
}

/* ============================== CALCULS ============================== */
function useComputed(data) {
  return useMemo(() => {
    if (!data) return null;
    const { clients, sessions, settings } = data;
    const activeClients = clients.filter((c) => c.status === "actif");
    const mrrActif = activeClients.reduce((s, c) => s + (c.mrr || 0), 0);

    const now = new Date();
    const curMonth = monthKey(now);
    const monthSessions = sessions.filter((s) => s.date && s.date.startsWith(curMonth));
    const prodSessions = monthSessions.filter((s) => {
      if (s.category === "rdv") return false;
    if (s.category === "livrable" && (s.value || 0) === 0) return false;
      if (!s.clients || s.clients.length === 0) return true;
      return s.clients.some((sc) => {
        const cl = clients.find((x) => x.id === sc.id);
        return cl && cl.status === "actif";
      });
    });
    const totalMinutes = prodSessions.reduce((s, x) => s + (x.duration || 0), 0);
    const totalHours = totalMinutes / 60;
    const tjmHoraire = totalHours > 0 ? mrrActif / totalHours : 0;

    const prospectMinutes = monthSessions.reduce((s, x) => s + (x.duration || 0), 0) - totalMinutes;

    const mrrStripe = activeClients.filter((c) => c.payment === "stripe").reduce((s, c) => s + (c.mrr || 0), 0);
    const fraisStripe = mrrStripe * (settings.stripeFeeRate / 100);
    const netEstime = mrrActif * (1 - settings.urssafRate / 100) - settings.coutsOutils - fraisStripe;

    const upsidePotentiel = clients.reduce((s, c) => {
      const pending = (c.upgrades || []).filter((u) => u.statut === "en attente");
      // gain = nouveau tarif − tarif actuel ; un upgrade déjà atteint (ou à 0 €) ne compte pas
      return s + pending.reduce((s2, u) => s2 + Math.max(0, (Number(u.montant) || 0) - (Number(c.mrr) || 0)), 0);
    }, 0);

    const capacite = { actuel: activeClients.length, max: 13 };

    return { activeClients, mrrActif, totalHours, tjmHoraire, netEstime, upsidePotentiel, capacite, curMonth, fraisStripe, prospectMinutes };
  }, [data]);
}

function computeRecommendations(data, computed) {
  if (!data || !computed) return [];
  const recos = [];
  const now = new Date();
  data.clients.forEach((c) => {
    (c.upgrades || []).forEach((u) => {
      if (u.statut === "en attente") {
        const d = daysBetween(new Date(u.date), now);
        if (d >= 30) {
          recos.push({
            id: "up-" + u.id, priority: "orange", clientId: c.id,
            text: `Relancer ${c.name} — upgrade ${euro(u.montant)} proposé il y a ${d} jours`,
          });
        }
      }
    });
    (c.echeances || []).forEach((e) => {
      const d = daysBetween(now, new Date(e.date));
      if (d <= 30 && d >= -3) {
        recos.push({
          id: "ech-" + e.id, priority: d <= 7 ? "rouge" : "orange", clientId: c.id,
          text: `${c.name} — échéance "${e.label}" ${d < 0 ? "dépassée" : `dans ${d} jours`} (${new Date(e.date).toLocaleDateString("fr-FR")})`,
        });
      }
    });
    if (c.signal) {
      recos.push({ id: "sig-" + c.id, priority: "jaune", clientId: c.id, text: `${c.name} — signal : ${c.signal}` });
    }
  });

  if (!data.settings.rcProSouscrite && computed.mrrActif >= 900) {
    recos.push({ id: "rcpro", priority: "rouge", text: `RC Pro non souscrite — MRR à ${euro(computed.mrrActif)}, seuil 1000€ proche` });
  }
  if (computed.capacite.actuel >= 11) {
    recos.push({ id: "cap", priority: "orange", text: `Capacité : ${computed.capacite.actuel}/${computed.capacite.max} clients actifs — approche du plafond` });
  }
  data.reminders.filter((r) => !r.done).forEach((r) => {
    const cl = r.clientId && data.clients.find((c) => c.id === r.clientId);
    const text = cl ? `${cl.name.split(" / ")[0]} — ${r.text}` : r.text;
    recos.push({ id: "rem-" + r.id, priority: "jaune", clientId: r.clientId || null, text, manual: true, reminderId: r.id });
  });

  const order = { rouge: 0, orange: 1, jaune: 2 };
  return recos.sort((a, b) => order[a.priority] - order[b.priority]);
}

const PRIORITY_STYLE = {
  rouge: { accent: C.zRed, label: "Urgent" },
  orange: { accent: C.zOrange, label: "À traiter" },
  jaune: { accent: C.faint, label: "À noter" },
};

/* ============================== UI ATOMS ============================== */
function Card({ children, style }) {
  return (
    <div style={{ background: C.card, borderRadius: 14, border: `1px solid ${C.line}`, padding: 20, ...style }}>
      {children}
    </div>
  );
}
function SectionTitle({ children, right }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, gap: 12 }}>
      <h2 style={{ margin: 0, fontSize: 15, fontWeight: 600, color: C.ink, letterSpacing: "-0.01em" }}>{children}</h2>
      {right}
    </div>
  );
}
function Badge({ children, color, bg }) {
  return (
    <span style={{
      fontSize: 11, fontWeight: 500, padding: "2px 9px 2px 7px", borderRadius: 999,
      background: bg, color, display: "inline-flex", alignItems: "center", gap: 6,
    }}>
      <span style={{ width: 6, height: 6, borderRadius: "50%", background: color, display: "inline-block" }} />
      {children}
    </span>
  );
}
function Btn({ children, onClick, variant = "primary", small, style }) {
  const base = {
    cursor: "pointer", fontFamily: FONT, fontWeight: 500,
    fontSize: small ? 12 : 13, padding: small ? "6px 11px" : "9px 15px",
    borderRadius: 10, display: "inline-flex", alignItems: "center", gap: 7, lineHeight: 1.2,
    border: "1px solid transparent",
  };
  const variants = {
    primary: { background: C.ink, color: "#fff" },
    ghost: { background: C.card, color: C.ink, border: `1px solid ${C.line}` },
    danger: { background: C.card, color: C.red, border: `1px solid ${C.redBg}` },
  };
  return (
    <button className="vx-btn" onClick={onClick} style={{ ...base, ...variants[variant], ...style }}>
      {children}
    </button>
  );
}
function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 5, fontSize: 12, color: C.gray, fontWeight: 500 }}>
      {label}
      {children}
    </label>
  );
}
const inputStyle = {
  fontFamily: FONT, fontSize: 14, padding: "8px 11px", borderRadius: 10,
  border: `1px solid ${C.line}`, color: C.ink, background: C.card,
};

/* ============================== DASHBOARD TAB ============================== */
/* ============================== SEUILS FISCAUX ============================== */
// CA de l'année : mois passés = CA encaissé saisi (à défaut, MRR figé du mois) ;
// mois en cours et à venir = CA saisi, à défaut MRR actuel (projection à MRR constant).
function computeCaYear(data, year, curMonth, mrrActif) {
  const hist = data.caHistory || [];
  const mrrHist = data.mrrHistory || [];
  let passe = 0, estime = 0, restants = 0;
  for (let m = 1; m <= 12; m++) {
    const key = `${year}-${String(m).padStart(2, "0")}`;
    const saisi = hist.find((h) => h.month === key);
    if (key < curMonth) {
      const v = saisi ? Number(saisi.ca) || 0 : Number((mrrHist.find((h) => h.month === key) || {}).mrr) || 0;
      passe += v;
      estime += v;
    } else {
      restants += 1;
      estime += saisi ? Number(saisi.ca) || 0 : mrrActif;
    }
  }
  return { passe, estime, restants };
}

function SeuilsFiscauxCard({ data, computed }) {
  const st = { ...defaultSettings, ...data.settings };
  const curYear = new Date().getFullYear();
  const years = [...new Set([curYear, curYear + 1])];
  const [year, setYear] = useState(years.includes(2027) ? 2027 : curYear);
  const { passe, estime, restants } = computeCaYear(data, year, computed.curMonth, computed.mrrActif);
  const rows = [
    { label: "Repère prudent (versement libératoire)", val: st.caRepereLiberatoire },
    { label: "Plafond haut estimé (versement libératoire)", val: st.caPlafondMax },
    { label: "Franchise de TVA", val: st.seuilTva },
  ];
  const mrrMoyenAutorise = restants > 0 ? Math.max(0, (st.caRepereLiberatoire - passe) / restants) : null;
  return (
    <Card>
      <SectionTitle right={
        <select style={inputStyle} value={year} onChange={(e) => setYear(parseInt(e.target.value, 10))} aria-label="Année">
          {years.map((y) => <option key={y} value={y}>{y}</option>)}
        </select>
      }>Seuils fiscaux</SectionTitle>
      <div style={{ fontSize: 12, color: C.gray, margin: "-6px 0 14px", lineHeight: 1.5 }}>
        CA estimé {year} : <b style={{ color: C.ink }}>{euro(estime)}</b>
        {" "}(CA déjà acquis {euro(passe)}, puis MRR actuel de {euro(computed.mrrActif)} pour les {restants} mois restants ou saisis dans Réglages).
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
        {rows.map((r) => {
          const pct = r.val > 0 ? (estime / r.val) * 100 : 0;
          const over = estime > r.val;
          // vert < 70 %, jaune 70-90 %, orange 90-100 %, rouge si dépassé
          const color = over ? C.zRed : pct >= 90 ? C.zOrange : pct >= 70 ? "#F2C230" : C.zGreen;
          const textColor = over ? C.zRed : pct >= 90 ? C.orangeText : pct >= 70 ? "#8A6500" : C.green;
          const status = over ? `dépassé de ${euro(estime - r.val)}` : `marge ${euro(r.val - estime)}`;
          return (
            <div key={r.label}>
              <div style={{ display: "flex", justifyContent: "space-between", gap: 10, flexWrap: "wrap", fontSize: 13, color: C.ink }}>
                <span style={{ fontWeight: 600 }}>{r.label}</span>
                <span style={{ color: textColor }}>{euro(estime)} / {euro(r.val)} · {status}</span>
              </div>
              <div style={{ height: 10, background: C.sunken, borderRadius: 6, marginTop: 6, overflow: "hidden" }}>
                <div style={{ width: `${Math.min(100, pct)}%`, height: 10, background: color, borderRadius: 6 }} />
              </div>
              <div style={{ fontSize: 11, color: C.faint, marginTop: 4 }}>Équivaut à {euro(r.val / 12)} de MRR constant sur 12 mois</div>
            </div>
          );
        })}
      </div>
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap", fontSize: 11, color: C.gray, marginTop: 14 }}>
        {[["Vert", C.zGreen, "moins de 70 %"], ["Jaune", "#F2C230", "70 à 90 %"], ["Orange", C.zOrange, "90 à 100 %"], ["Rouge", C.zRed, "dépassé"]].map(([n, col, t]) => (
          <span key={n} style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
            <span style={{ width: 10, height: 10, borderRadius: 3, background: col }} />{n} : {t}
          </span>
        ))}
      </div>
      {mrrMoyenAutorise !== null && (
        <div style={{ fontSize: 13, color: C.ink, marginTop: 16, padding: "10px 12px", background: C.sunken, borderRadius: 10, lineHeight: 1.5 }}>
          Pour rester sous le repère prudent, le CA moyen doit rester à <b>{euro(mrrMoyenAutorise)} par mois</b> sur les {restants} mois restants de {year}.
        </div>
      )}
      <div style={{ fontSize: 11, color: C.faint, marginTop: 12, lineHeight: 1.5 }}>
        Le CA de l'année N conditionne le versement libératoire de l'année N+2 (le CA 2027 joue sur 2029). Franchise de TVA : 37 500 € pour les prestations de services. Les repères sont des estimations à valider avec un expert-comptable ; ajustables dans Réglages.
      </div>
    </Card>
  );
}

const PIE_COLORS = ["#FF8C00", "#3A6BD6", "#2F9E5E", "#F2C230", "#8E5BD6", "#D64541", "#17A2B8", "#E667A0", "#7A8B3A", "#5B6275", "#B07A4A", "#2D4A8A"];

function ConcentrationCard({ computed }) {
  const [view, setView] = useState("client");
  const actifs = (computed.activeClients || []).filter((c) => (Number(c.mrr) || 0) > 0);
  const total = actifs.reduce((s, c) => s + (Number(c.mrr) || 0), 0);
  const groups = {};
  actifs.forEach((c) => {
    const key = view === "client" ? c.name : ((c.sector || "").trim() || "Non renseigné");
    groups[key] = (groups[key] || 0) + (Number(c.mrr) || 0);
  });
  const slices = Object.entries(groups)
    .map(([name, value]) => ({ name, value, pct: total > 0 ? (value / total) * 100 : 0 }))
    .sort((a, b) => b.value - a.value);
  const top = slices[0];
  // vert < 20 %, jaune 20-25 %, orange 25-35 %, rouge au-delà
  const tone = (pct) => (pct > 35 ? [C.zRed, C.red] : pct > 25 ? [C.zOrange, C.orangeText] : pct >= 20 ? ["#F2C230", "#8A6500"] : [C.zGreen, C.green]);
  const [, topText] = top ? tone(top.pct) : [null, C.ink];
  const unit = view === "client" ? "client" : "secteur";
  return (
    <Card>
      <SectionTitle right={
        <select style={inputStyle} value={view} onChange={(e) => setView(e.target.value)} aria-label="Vue">
          <option value="client">Par client</option>
          <option value="secteur">Par secteur</option>
        </select>
      }>Dépendance du MRR</SectionTitle>
      {slices.length === 0 ? (
        <div style={{ fontSize: 12, color: C.gray, padding: "8px 0" }}>Aucun client actif avec un MRR : le camembert apparaîtra dès la première saisie.</div>
      ) : (
        <>
          <div style={{ display: "flex", gap: 20, flexWrap: "wrap", alignItems: "center" }}>
            <div style={{ width: 200, height: 200, flexShrink: 0 }}>
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={slices} dataKey="value" nameKey="name" innerRadius={52} outerRadius={92} paddingAngle={1} stroke={C.card} strokeWidth={2}>
                    {slices.map((sl, i) => <Cell key={sl.name} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                  </Pie>
                  <Tooltip formatter={(v, n) => [`${euro(v)} (${total > 0 ? Math.round((v / total) * 100) : 0} %)`, n]} contentStyle={{ fontFamily: FONT, fontSize: 12, borderRadius: 10, border: `1px solid ${C.line}`, boxShadow: "none" }} />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div style={{ flex: 1, minWidth: 220, display: "flex", flexDirection: "column", gap: 6 }}>
              {slices.map((sl, i) => (
                <div key={sl.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: C.ink }}>
                  <span style={{ width: 10, height: 10, borderRadius: 3, background: PIE_COLORS[i % PIE_COLORS.length], flexShrink: 0 }} />
                  <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{sl.name}</span>
                  <span style={{ color: C.gray }}>{euro(sl.value)}</span>
                  <span style={{ width: 44, textAlign: "right", fontWeight: 600, color: tone(sl.pct)[1] }}>{Math.round(sl.pct)} %</span>
                </div>
              ))}
            </div>
          </div>
          <div style={{ fontSize: 13, color: C.ink, marginTop: 14, padding: "10px 12px", background: C.sunken, borderRadius: 10, lineHeight: 1.5 }}>
            Le plus gros {unit} pèse <b style={{ color: topText }}>{Math.round(top.pct)} % du MRR</b> ({top.name}, {euro(top.value)} sur {euro(total)}).
            {" "}{top.pct > 35 ? "Dépendance forte : perdre ce poste ferait chuter le revenu." : top.pct > 25 ? "À surveiller : viser moins de 25 %." : top.pct >= 20 ? "Limite haute : au-delà de 25 %, la dépendance devient sensible." : "Répartition saine (moins de 20 %)."}
          </div>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", fontSize: 11, color: C.gray, marginTop: 12 }}>
            {[["Vert", C.zGreen, "moins de 20 %"], ["Jaune", "#F2C230", "20 à 25 %"], ["Orange", C.zOrange, "25 à 35 %"], ["Rouge", C.zRed, "plus de 35 %"]].map(([n, col, t]) => (
              <span key={n} style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                <span style={{ width: 10, height: 10, borderRadius: 3, background: col }} />{n} : {t}
              </span>
            ))}
          </div>
        </>
      )}
      <div style={{ fontSize: 11, color: C.faint, marginTop: 12, lineHeight: 1.5 }}>
        Calculé sur les clients au statut « actif ». Le camembert se met à jour dès que vous modifiez un MRR, un statut ou un secteur dans la fiche client.
      </div>
    </Card>
  );
}

function DashboardTab({ data, computed, recos, onToggleReminder }) {
  const tjmMax = Math.max(80, data.settings.tjmSeuils[2] * 1.5); // échelle fixe : au-delà, l'aiguille reste en butée
  const tjmZone = zoneColorForTJM(computed.tjmHoraire, data.settings.tjmSeuils);
  const [s0, s1, s2] = data.settings.tjmSeuils;

  const history = [...data.mrrHistory].sort((a, b) => (a.month > b.month ? 1 : -1));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {/* JAUGES + INDICATEURS */}
      <Card style={{ padding: "22px 24px 20px" }}>
        <div className="vx-gauges" style={{ display: "flex", gap: 36, flexWrap: "wrap" }}>
          <RadialGauge
            value={computed.mrrActif}
            max={data.settings.mrrObjectifPrincipal}
            segments={[
              { from: 0, to: data.settings.mrrObjectifPhase1, color: C.zOrange },
              { from: data.settings.mrrObjectifPhase1, to: data.settings.mrrObjectifPrincipal, color: C.zGreen },
            ]}
            label="MRR actif"
            sublabel={`${Math.round((computed.mrrActif / data.settings.mrrObjectifPrincipal) * 100)} % de ${euro(data.settings.mrrObjectifPrincipal)}`}
            displayValue={euro(computed.mrrActif)}
          />
          <RadialGauge
            value={computed.tjmHoraire}
            max={tjmMax}
            segments={[
              { from: 0, to: s0, color: C.zRed },
              { from: s0, to: s1, color: C.zOrange },
              { from: s1, to: s2, color: C.zGreen },
              { from: s2, to: tjmMax, color: C.zBlue },
            ]}
            label="Taux horaire du mois"
            sublabel={`Zone ${{ Rouge: "rouge", Orange: "orange", Vert: "verte", Bleu: "bleue" }[tjmZone.label]}`}
            displayValue={`${euro1(computed.tjmHoraire)}/h`}
          />
        </div>

        <div className="vx-stats" style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", marginTop: 22, borderTop: `1px solid ${C.line}` }}>
          {(() => {
            const prev = [...data.mrrHistory].filter((h) => h.month < computed.curMonth).sort((a, b) => (a.month < b.month ? 1 : -1))[0];
            const delta = prev ? computed.mrrActif - Number(prev.mrr) : null;
            return [
              { label: "Net estimé", value: euro(computed.netEstime), sub: "par mois, après charges" },
              { label: "Upside potentiel", value: euro(computed.upsidePotentiel), sub: "upgrades en attente" },
              { label: "Capacité", value: `${computed.capacite.actuel} / ${computed.capacite.max}`, sub: "clients actifs" },
              {
                label: "Évolution",
                value: delta === null ? "–" : `${delta >= 0 ? "+" : "−"}${euro(Math.abs(delta))}`,
                sub: prev ? `vs ${monthLabel(prev.month)}` : "pas encore d'historique",
                color: delta === null ? C.ink : delta >= 0 ? C.green : C.red,
              },
            ];
          })().map((m, k) => (
            <div key={k} className="vx-stat" style={{ padding: "16px 16px 0", borderLeft: k > 0 ? `1px solid ${C.line}` : "none" }}>
              <div style={{ fontSize: 12, color: C.gray }}>{m.label}</div>
              <div style={{ fontSize: 21, fontWeight: 600, letterSpacing: "-0.02em", color: m.color || C.ink, marginTop: 3 }}>{m.value}</div>
              <div style={{ fontSize: 12, color: C.faint, marginTop: 2 }}>{m.sub}</div>
            </div>
          ))}
        </div>
      </Card>

      <SeuilsFiscauxCard data={data} computed={computed} />

      <ConcentrationCard computed={computed} />

      {/* HISTORIQUE */}
      <Card>
        <SectionTitle>Évolution du MRR</SectionTitle>
        {history.length < 2 ? (
          <div style={{ fontSize: 12, color: C.gray, padding: "20px 0" }}>
            La courbe apparaîtra dès le deuxième mois : le MRR de chaque mois est désormais enregistré automatiquement.
          </div>
        ) : (
          <div style={{ height: 220 }}>
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={history.map((h) => ({ ...h, label: monthLabel(h.month) }))} margin={{ top: 8, right: 24, left: 0, bottom: 0 }}>
                <CartesianGrid stroke={C.line} vertical={false} />
                <XAxis dataKey="label" tick={{ fontSize: 11, fill: C.gray }} axisLine={{ stroke: C.line }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: C.gray }} axisLine={false} tickLine={false} width={50} />
                <Tooltip formatter={(v) => [euro(v), "MRR"]} contentStyle={{ fontFamily: FONT, fontSize: 12, borderRadius: 10, border: `1px solid ${C.line}`, boxShadow: "none" }} />
                <Line type="monotone" dataKey="mrr" stroke={C.orange} strokeWidth={2.5} dot={{ r: 3, fill: C.card, stroke: C.orange, strokeWidth: 2 }} activeDot={{ r: 5 }} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        )}
      </Card>

      {/* RECOMMANDATIONS */}
      <Card>
        <SectionTitle right={<span style={{ fontSize: 12, color: C.faint }}>{recos.length} en cours</span>}>À traiter</SectionTitle>
        {recos.length === 0 ? (
          <div style={{ fontSize: 12, color: C.gray, padding: "8px 0" }}>Rien à traiter pour l'instant.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column" }}>
            {recos.map((r, i) => {
              const style = PRIORITY_STYLE[r.priority];
              return (
                <div key={r.id} style={{
                  display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10,
                  padding: "11px 4px 11px 14px", borderLeft: `3px solid ${style.accent}`,
                  borderTop: i > 0 ? `1px solid ${C.line}` : "none",
                }}>
                  <div style={{ fontSize: 13.5, color: C.ink, lineHeight: 1.5 }}>
                    <span style={{ fontSize: 11, fontWeight: 600, color: style.accent, marginRight: 8 }}>{style.label}</span>
                    {r.text}
                  </div>
                  {r.manual && (
                    <button onClick={() => onToggleReminder(r.reminderId)} style={{ border: `1px solid ${C.line}`, background: "#fff", borderRadius: 8, padding: 5, cursor: "pointer", flexShrink: 0 }}>
                      <Check size={13} color={C.gray} />
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </Card>
    </div>
  );
}

/* ============================== CLIENTS TAB ============================== */
const STATUS_STYLE = {
  actif: { bg: C.greenBg, text: C.green, label: "Actif" },
  prospect: { bg: C.blueBg, text: C.blue, label: "Prospect" },
  standby: { bg: "#fef9c3", text: "#854d0e", label: "Standby" },
  perdu: { bg: C.redBg, text: C.red, label: "Perdu" },
};

function ClientDrawer({ client, onSave, onClose, onDelete }) {
  const [c, setC] = useState(client);
  const set = (k, v) => setC((prev) => ({ ...prev, [k]: v }));

  const addUpgrade = () => set("upgrades", [...(c.upgrades || []), { id: uid(), montant: 0, date: todayISO(), statut: "en attente", label: "" }]);
  const updUpgrade = (id, patch) => set("upgrades", c.upgrades.map((u) => (u.id === id ? { ...u, ...patch } : u)));
  const rmUpgrade = (id) => set("upgrades", c.upgrades.filter((u) => u.id !== id));

  const addEcheance = () => set("echeances", [...(c.echeances || []), { id: uid(), label: "", date: todayISO() }]);
  const updEcheance = (id, patch) => set("echeances", c.echeances.map((e) => (e.id === id ? { ...e, ...patch } : e)));
  const rmEcheance = (id) => set("echeances", c.echeances.filter((e) => e.id !== id));

  const addMeeting = () => set("meetings", [...(c.meetings || []), { id: uid(), date: todayISO(), subject: "" }]);
  const updMeeting = (id, patch) => set("meetings", c.meetings.map((m) => (m.id === id ? { ...m, ...patch } : m)));
  const rmMeeting = (id) => set("meetings", c.meetings.filter((m) => m.id !== id));

  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(26,26,46,0.4)", zIndex: 50,
      display: "flex", justifyContent: "flex-end",
    }} onClick={onClose}>
      <div
        style={{ width: "min(480px, 100%)", background: "#fff", height: "100%", overflowY: "auto", padding: 24 }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <div style={{ fontSize: 16, fontWeight: 600, color: C.ink }}>{client.id ? "Fiche client" : "Nouveau contact"}</div>
          <button onClick={onClose} style={{ border: "none", background: "transparent", cursor: "pointer" }}><X size={18} /></button>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <Field label="Nom / entreprise">
            <input style={inputStyle} value={c.name} onChange={(e) => set("name", e.target.value)} />
          </Field>
          <div style={{ display: "flex", gap: 10 }}>
            <Field label="Secteur"><input style={inputStyle} value={c.sector} onChange={(e) => set("sector", e.target.value)} /></Field>
            <Field label="Zone"><input style={inputStyle} value={c.zone} onChange={(e) => set("zone", e.target.value)} /></Field>
          </div>
          <div style={{ display: "flex", gap: 10 }}>
            <Field label="Statut">
              <select style={inputStyle} value={c.status} onChange={(e) => set("status", e.target.value)}>
                <option value="prospect">Prospect</option>
                <option value="actif">Actif</option>
                <option value="standby">Standby</option>
                <option value="perdu">Perdu</option>
              </select>
            </Field>
            <Field label="MRR (€)"><input type="number" style={inputStyle} value={c.mrr} onChange={(e) => set("mrr", parseFloat(e.target.value) || 0)} /></Field>
          </div>
          <div style={{ display: "flex", gap: 10 }}>
            <Field label="Paiement">
              <select style={inputStyle} value={c.payment} onChange={(e) => set("payment", e.target.value)}>
                <option value="">—</option>
                <option value="stripe">Stripe (CB)</option>
                <option value="virement">Virement</option>
              </select>
            </Field>
            <Field label="Posts / semaine">
              <input type="number" style={{ ...inputStyle, ...(c.status === "actif" && !(Number(c.postsPerWeek) > 0) ? { borderColor: C.zRed } : {}) }} value={c.postsPerWeek} onChange={(e) => set("postsPerWeek", parseFloat(e.target.value) || 0)} />
              {c.status === "actif" && !(Number(c.postsPerWeek) > 0) && <span style={{ fontSize: 11, color: C.red, fontWeight: 400 }}>Obligatoire pour valoriser ses posts</span>}
            </Field>
          </div>
          <Field label="Source (si prospect)"><input style={inputStyle} value={c.source || ""} onChange={(e) => set("source", e.target.value)} /></Field>
          <Field label="Dernier contact"><input type="date" style={inputStyle} value={c.lastContact || ""} onChange={(e) => set("lastContact", e.target.value)} /></Field>
          <Field label="Signal détecté"><input style={inputStyle} value={c.signal || ""} onChange={(e) => set("signal", e.target.value)} placeholder="ex: intérêt exprimé pour upgrade" /></Field>
          <Field label="Notes"><textarea style={{ ...inputStyle, minHeight: 60 }} value={c.notes || ""} onChange={(e) => set("notes", e.target.value)} /></Field>

          {/* UPGRADES */}
          <div style={{ marginTop: 8 }}>
            <SectionTitle right={<Btn small variant="ghost" onClick={addUpgrade}><Plus size={12} />Ajouter</Btn>}>Upgrades</SectionTitle>
            {(c.upgrades || []).map((u) => (
              <div key={u.id} style={{ display: "flex", gap: 6, marginBottom: 6, alignItems: "center" }}>
                <input style={{ ...inputStyle, flex: 1 }} placeholder="Libellé" value={u.label} onChange={(e) => updUpgrade(u.id, { label: e.target.value })} />
                <input type="number" style={{ ...inputStyle, width: 70 }} value={u.montant} onChange={(e) => updUpgrade(u.id, { montant: parseFloat(e.target.value) || 0 })} />
                <select style={{ ...inputStyle, width: 100 }} value={u.statut} onChange={(e) => updUpgrade(u.id, { statut: e.target.value })}>
                  <option value="en attente">En attente</option>
                  <option value="accepté">Accepté</option>
                  <option value="refusé">Refusé</option>
                </select>
                <button onClick={() => rmUpgrade(u.id)} style={{ border: "none", background: "transparent", cursor: "pointer" }}><Trash2 size={14} color={C.red} /></button>
              </div>
            ))}
          </div>

          {/* ECHEANCES */}
          <div style={{ marginTop: 8 }}>
            <SectionTitle right={<Btn small variant="ghost" onClick={addEcheance}><Plus size={12} />Ajouter</Btn>}>Échéances</SectionTitle>
            {(c.echeances || []).map((e) => (
              <div key={e.id} style={{ display: "flex", gap: 6, marginBottom: 6, alignItems: "center" }}>
                <input style={{ ...inputStyle, flex: 1 }} placeholder="Libellé" value={e.label} onChange={(ev) => updEcheance(e.id, { label: ev.target.value })} />
                <input type="date" style={{ ...inputStyle, width: 140 }} value={e.date} onChange={(ev) => updEcheance(e.id, { date: ev.target.value })} />
                <button onClick={() => rmEcheance(e.id)} style={{ border: "none", background: "transparent", cursor: "pointer" }}><Trash2 size={14} color={C.red} /></button>
              </div>
            ))}
          </div>

          {/* RENDEZ-VOUS */}
          <div style={{ marginTop: 8 }}>
            <SectionTitle right={<Btn small variant="ghost" onClick={addMeeting}><Plus size={12} />Ajouter</Btn>}>
              Rendez-vous / échanges (hors TJM)
            </SectionTitle>
            {(c.meetings || []).length === 0 && (
              <div style={{ fontSize: 11, color: C.gray, marginBottom: 6 }}>Aucun rdv tracé pour l'instant.</div>
            )}
            {(c.meetings || []).map((m) => (
              <div key={m.id} style={{ display: "flex", gap: 6, marginBottom: 6, alignItems: "center" }}>
                <input type="date" style={{ ...inputStyle, width: 140 }} value={m.date} onChange={(e) => updMeeting(m.id, { date: e.target.value })} />
                <input style={{ ...inputStyle, flex: 1 }} placeholder="Sujet échangé" value={m.subject} onChange={(e) => updMeeting(m.id, { subject: e.target.value })} />
                <button onClick={() => rmMeeting(m.id)} style={{ border: "none", background: "transparent", cursor: "pointer" }}><Trash2 size={14} color={C.red} /></button>
              </div>
            ))}
          </div>

          <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
            <Btn onClick={() => onSave(c)}>Enregistrer</Btn>
            {client.id && <Btn variant="danger" onClick={() => onDelete(c.id)}><Trash2 size={13} />Supprimer</Btn>}
          </div>
        </div>
      </div>
    </div>
  );
}

function ClientsTab({ data, save }) {
  const [filter, setFilter] = useState("tous");
  const [editing, setEditing] = useState(null);

  const filtered = data.clients.filter((c) => filter === "tous" || c.status === filter);

  const upsert = (c) => {
    const exists = data.clients.some((x) => x.id === c.id);
    const clients = exists ? data.clients.map((x) => (x.id === c.id ? c : x)) : [...data.clients, { ...c, id: uid() }];
    save({ ...data, clients });
    setEditing(null);
  };
  const remove = (id) => {
    save({ ...data, clients: data.clients.filter((c) => c.id !== id) });
    setEditing(null);
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    const clientsSheet = data.clients.map((c) => ({
      Nom: c.name, Statut: c.status, Secteur: c.sector, Zone: c.zone, MRR: c.mrr,
      Paiement: c.payment, "Posts/semaine": c.postsPerWeek, "Dernier contact": c.lastContact,
      Notes: c.notes, Signal: c.signal || "",
      Upgrades: (c.upgrades || []).map((u) => `${u.label} ${u.montant}€ (${u.statut})`).join(" | "),
      Echeances: (c.echeances || []).map((e) => `${e.label} — ${e.date}`).join(" | "),
      RDV: (c.meetings || []).map((m) => `${m.date} : ${m.subject}`).join(" | "),
    }));
    const sessionsSheet = data.sessions.map((s) => ({
      Date: s.date, Catégorie: s.category, "Durée (min)": s.duration, Valeur: s.value, Note: s.note || "",
      Clients: (s.clients || []).map((x) => x.name).join(", "),
    }));
    const mrrSheet = data.mrrHistory.map((h) => ({ Mois: h.month, MRR: h.mrr }));
    const remindersSheet = data.reminders.map((r) => ({ Texte: r.text, Date: r.date || "", Fait: r.done ? "Oui" : "Non" }));
    const settingsSheet = [
      { Paramètre: "Seuil TJM rouge/orange", Valeur: data.settings.tjmSeuils[0] },
      { Paramètre: "Seuil TJM orange/vert", Valeur: data.settings.tjmSeuils[1] },
      { Paramètre: "Seuil TJM vert/bleu", Valeur: data.settings.tjmSeuils[2] },
      { Paramètre: "Objectif MRR Phase 1", Valeur: data.settings.mrrObjectifPhase1 },
      { Paramètre: "Objectif MRR cible", Valeur: data.settings.mrrObjectifPrincipal },
      { Paramètre: "Taux URSSAF (%)", Valeur: data.settings.urssafRate },
      { Paramètre: "Coûts outils (€)", Valeur: data.settings.coutsOutils },
      { Paramètre: "Frais Stripe (%)", Valeur: data.settings.stripeFeeRate },
      { Paramètre: "RC Pro souscrite", Valeur: data.settings.rcProSouscrite ? "Oui" : "Non" },
    ];
    wb.SheetNames.push("Clients"); wb.Sheets["Clients"] = XLSX.utils.json_to_sheet(clientsSheet);
    wb.SheetNames.push("Temps"); wb.Sheets["Temps"] = XLSX.utils.json_to_sheet(sessionsSheet);
    wb.SheetNames.push("Historique MRR"); wb.Sheets["Historique MRR"] = XLSX.utils.json_to_sheet(mrrSheet);
    wb.SheetNames.push("Rappels"); wb.Sheets["Rappels"] = XLSX.utils.json_to_sheet(remindersSheet);
    wb.SheetNames.push("Réglages"); wb.Sheets["Réglages"] = XLSX.utils.json_to_sheet(settingsSheet);
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `vulpix-export-${todayISO()}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const exportJSON = () => downloadBackup(data);


  const importJSON = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result);
        if (parsed.clients && parsed.settings) {
          save(parsed);
          alert("Sauvegarde restaurée avec succès.");
        } else {
          alert("Fichier invalide.");
        }
      } catch {
        alert("Erreur de lecture du fichier.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          {["tous", "actif", "prospect", "standby", "perdu"].map((f) => (
            <button
              key={f} onClick={() => setFilter(f)}
              style={{
                cursor: "pointer", padding: "6px 14px", borderRadius: 20, fontSize: 11, fontWeight: 600,
                fontFamily: FONT, background: filter === f ? C.ink : "#fff", color: filter === f ? "#fff" : C.ink,
                border: `1px solid ${filter === f ? C.ink : C.line}`, textTransform: "capitalize",
              }}
            >
              {f}
            </button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Btn variant="ghost" onClick={exportJSON}><Download size={13} />Sauvegarde JSON</Btn>
          <label style={{
            border: `1px solid ${C.line}`, borderRadius: 10, padding: "9px 16px", fontSize: 12, fontWeight: 600,
            fontFamily: FONT, color: C.ink, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: 6,
          }}>
            Restaurer
            <input type="file" accept=".json" onChange={importJSON} style={{ display: "none" }} />
          </label>
          <Btn variant="ghost" onClick={exportExcel}><Download size={13} />Export Excel</Btn>
          <Btn onClick={() => setEditing({ name: "", sector: "", zone: "", status: "prospect", mrr: 0, payment: "", postsPerWeek: 0, upgrades: [], echeances: [] })}>
            <Plus size={13} />Ajouter
          </Btn>
        </div>
      </div>

      <div style={{ background: C.card, border: `1px solid ${C.line}`, borderRadius: 14, overflow: "hidden" }}>
        {filtered.length === 0 && (
          <div style={{ padding: 20, fontSize: 12, color: C.gray, textAlign: "center" }}>Aucun contact dans ce filtre.</div>
        )}
        {filtered.map((c, i) => {
          const st = STATUS_STYLE[c.status];
          const pendingUpgrade = (c.upgrades || []).find((u) => u.statut === "en attente");
          return (
            <div key={c.id} onClick={() => setEditing(c)} style={{
              display: "flex", alignItems: "center", gap: 14, cursor: "pointer",
              padding: "14px 16px", borderTop: i > 0 ? `1px solid ${C.line}` : "none",
            }}>
              <div style={{ width: 4, alignSelf: "stretch", borderRadius: 2, background: st.text, flexShrink: 0 }} />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.ink }}>{c.name}</div>
                  <span style={{ fontSize: 11, color: st.text, fontWeight: 600 }}>{st.label}</span>
                </div>
                <div style={{ fontSize: 11, color: C.gray, marginTop: 2 }}>
                  {c.sector}{c.zone && `, ${c.zone}`}
                  {pendingUpgrade && <span style={{ color: C.orangeText, fontWeight: 600 }}> — upgrade {euro(pendingUpgrade.montant)} en attente</span>}
                </div>
              </div>
              {c.mrr > 0 && <div style={{ fontSize: 15, fontWeight: 600, fontFamily: FONT, color: C.ink, flexShrink: 0 }}>{euro(c.mrr)}</div>}
              <ChevronRight size={15} color={C.gray} style={{ flexShrink: 0 }} />
            </div>
          );
        })}
      </div>

      {editing && (
        <ClientDrawer client={editing} onSave={upsert} onClose={() => setEditing(null)} onDelete={remove} />
      )}
    </div>
  );
}

/* ============================== TEMPS TAB ============================== */
function SessionForm({ data, save }) {
  const [category, setCategory] = useState("posts");
  const [duration, setDuration] = useState(65);
  const [date, setDate] = useState(todayISO());
  const [postCounts, setPostCounts] = useState({});
  const [selectedClients, setSelectedClients] = useState([]);
  const [manualValue, setManualValue] = useState(null);
  const [note, setNote] = useState("");
  const [lastResult, setLastResult] = useState(null);

  const activeClients = data.clients.filter((c) => c.status === "actif" && c.mrr > 0);
  const billableClients = data.clients.filter((c) => c.status === "actif" || c.status === "prospect");
  const w = data.settings.poidsCategories;
  const div = data.settings.sessionsParMoisMaintenance;

  const computeValue = () => {
    if (category === "rdv") return 0;
    if (category === "posts" || category === "geoposts") {
      return activeClients.reduce((sum, c) => {
        const count = postCounts[c.id] || 0;
        if (!count) return sum;
        const postsPerMonth = (Number(c.postsPerWeek) || 0) * 4.33;
        const valuePerPost = postsPerMonth > 0 ? (Number(c.mrr) * (w.posts / 100)) / postsPerMonth : 0;
        return sum + count * valuePerPost;
      }, 0);
    }
    if (["ads", "whatsapp", "gmb", "maintenance", "rapports"].includes(category)) {
      return selectedClients.reduce((sum, id) => {
        const c = activeClients.find((x) => x.id === id);
        if (!c) return sum;
        const weight = w[category] || 0;
        const sessions = div[category] || 4;
        return sum + (c.mrr * (weight / 100)) / sessions;
      }, 0);
    }
    return manualValue !== null ? manualValue : (duration / 60) * data.settings.tarifReferenceLivrable;
  };

  const value = computeValue();
  const hours = duration / 60;
  const rate = hours > 0 ? value / hours : 0;
  const zone = zoneColorForTJM(rate, data.settings.tjmSeuils);

  const totalPosts = Object.values(postCounts).reduce((a, b) => a + (b || 0), 0);

  // Rythme de publication manquant : on le corrige ici, sans quitter le formulaire
  const [rhythmDraft, setRhythmDraft] = useState({});
  const valuePerPost = (c) => {
    const ppm = (Number(c.postsPerWeek) || 0) * 4.33;
    return ppm > 0 ? (Number(c.mrr) * (w.posts / 100)) / ppm : 0;
  };
  const saveRhythm = (c) => {
    const n = parseFloat(rhythmDraft[c.id]);
    if (!n || n <= 0) return;
    save({ ...data, clients: data.clients.map((x) => (x.id === c.id ? { ...x, postsPerWeek: n } : x)) });
    setRhythmDraft({ ...rhythmDraft, [c.id]: "" });
  };
  const blockedClients = (category === "posts" || category === "geoposts")
    ? activeClients.filter((c) => (postCounts[c.id] || 0) > 0 && valuePerPost(c) === 0)
    : [];

  const logSession = () => {
    const clientsInvolved =
      category === "posts" || category === "geoposts"
        ? Object.entries(postCounts).filter(([, n]) => n > 0).map(([id, n]) => ({ id, name: activeClients.find((c) => c.id === id)?.name, count: n }))
        : selectedClients.map((id) => ({ id, name: (data.clients.find((c) => c.id === id) || {}).name }));
    const session = { id: uid(), date, category, duration, value: Math.round(value * 100) / 100, note, clients: clientsInvolved };
    save({ ...data, sessions: [...data.sessions, session] });
    setLastResult({ value, rate, zone });
    setPostCounts({}); setSelectedClients([]); setManualValue(null); setNote("");
  };

  return (
    <Card>
      <SectionTitle>Loguer une session</SectionTitle>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
        <Field label="Date"><input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        <Field label="Catégorie">
          <select style={inputStyle} value={category} onChange={(e) => setCategory(e.target.value)}>
            {CATEGORIES.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
          </select>
        </Field>
        <Field label="Durée (min)"><input type="number" style={{ ...inputStyle, width: 90 }} value={duration} onChange={(e) => setDuration(parseFloat(e.target.value) || 0)} /></Field>
      </div>

      {(category === "posts" || category === "geoposts") && (
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600, marginBottom: 6 }}>Nombre de posts par client (détail optionnel)</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(180px,1fr))", gap: 8 }}>
            {activeClients.map((c) => {
              const vpp = valuePerPost(c);
              return (
              <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ fontSize: 12, flex: 1, lineHeight: 1.3 }}>
                  {c.name.split(" / ")[0]}
                  <br />
                  {vpp > 0
                    ? <span style={{ fontSize: 11, color: C.faint }}>{euro1(vpp)} / post</span>
                    : (
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 4, marginTop: 2 }}>
                        <input type="number" min="1" placeholder="posts/sem."
                          aria-label={`Posts par semaine pour ${c.name}`}
                          style={{ ...inputStyle, width: 82, padding: "3px 6px", fontSize: 11, borderColor: C.zRed }}
                          value={rhythmDraft[c.id] || ""}
                          onChange={(e) => setRhythmDraft({ ...rhythmDraft, [c.id]: e.target.value })}
                          onKeyDown={(e) => { if (e.key === "Enter") saveRhythm(c); }} />
                        <button onClick={() => saveRhythm(c)} style={{ border: `1px solid ${C.line}`, background: C.card, borderRadius: 6, padding: "3px 7px", fontSize: 11, cursor: "pointer", fontFamily: FONT }}>OK</button>
                      </span>
                    )}
                </span>
                <input type="number" min="0" style={{ ...inputStyle, width: 50, padding: "5px 6px" }}
                  value={postCounts[c.id] || ""}
                  onChange={(e) => setPostCounts({ ...postCounts, [c.id]: parseFloat(e.target.value) || 0 })} />
              </div>
              );
            })}
          </div>
          <div style={{ fontSize: 11, color: C.gray, marginTop: 6 }}>Total : {totalPosts} posts</div>
        </div>
      )}

      {["ads", "whatsapp", "gmb", "maintenance", "rapports"].includes(category) && (
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600, marginBottom: 6 }}>Clients concernés</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {activeClients.map((c) => {
              const on = selectedClients.includes(c.id);
              return (
                <button key={c.id}
                  onClick={() => setSelectedClients(on ? selectedClients.filter((x) => x !== c.id) : [...selectedClients, c.id])}
                  style={{
                    border: `1px solid ${on ? C.orange : C.line}`, background: on ? C.orangeBg : "#fff",
                    color: on ? C.orangeText : C.ink, borderRadius: 20, padding: "5px 12px", fontSize: 11,
                    fontFamily: FONT, fontWeight: 600, cursor: "pointer",
                  }}>
                  {c.name.split(" / ")[0]}
                </button>
              );
            })}
          </div>
        </div>
      )}

      {category === "rdv" && (
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600, marginBottom: 6 }}>Avec qui ? (optionnel — laisse vide pour un rdv réseau / nouveau contact)</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
            {data.clients.filter((c) => c.status !== "perdu").map((c) => {
              const on = selectedClients.includes(c.id);
              return (
                <button key={c.id}
                  onClick={() => setSelectedClients(on ? selectedClients.filter((x) => x !== c.id) : [...selectedClients, c.id])}
                  style={{
                    border: `1px solid ${on ? C.orange : C.line}`, background: on ? C.orangeBg : "#fff",
                    color: on ? C.orangeText : C.ink, borderRadius: 20, padding: "5px 12px", fontSize: 11,
                    fontFamily: FONT, fontWeight: 600, cursor: "pointer",
                  }}>
                  {c.name.split(" / ")[0]} {c.status === "prospect" && "· prospect"}
                </button>
              );
            })}
          </div>
          <div style={{ fontSize: 10, color: C.gray }}>
            Compté dans ton temps total du mois, mais exclu du TJM. Mets le nom du contact et le sujet dans la note si la personne n'a pas de fiche.
          </div>
        </div>
      )}

      {(category === "livrable" || category === "autre") && (
        <div style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600, marginBottom: 6 }}>Client concerné (actif ou prospect)</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
            {billableClients.map((c) => {
              const on = selectedClients.includes(c.id);
              return (
                <button key={c.id}
                  onClick={() => setSelectedClients(on ? selectedClients.filter((x) => x !== c.id) : [...selectedClients, c.id])}
                  style={{
                    border: `1px solid ${on ? C.orange : C.line}`, background: on ? C.orangeBg : "#fff",
                    color: on ? C.orangeText : C.ink, borderRadius: 20, padding: "5px 12px", fontSize: 11,
                    fontFamily: FONT, fontWeight: 600, cursor: "pointer",
                  }}>
                  {c.name.split(" / ")[0]} {c.status === "prospect" && "· prospect"}
                </button>
              );
            })}
          </div>
          <Field label={`Valeur de la tâche (€) — suggérée à ${data.settings.tarifReferenceLivrable}€/h, ajustable`}>
            <input type="number" style={{ ...inputStyle, width: 140 }}
              value={manualValue !== null ? manualValue : Math.round((duration / 60) * data.settings.tarifReferenceLivrable * 100) / 100}
              onChange={(e) => setManualValue(e.target.value === "" ? null : parseFloat(e.target.value))} />
          </Field>
          <div style={{ fontSize: 10, color: C.gray, marginTop: 4 }}>
            Mets 0€ uniquement si c'est un vrai geste offert (hors production) — sinon laisse la suggestion, elle compte dans ton TJM comme du vrai travail client.
          </div>
        </div>
      )}

      <Field label="Note (optionnel)"><input style={inputStyle} value={note} onChange={(e) => setNote(e.target.value)} /></Field>

      {category === "rdv" ? (
      <div style={{ marginTop: 14, background: C.bg, borderRadius: 10, padding: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600 }}>Temps de rendez-vous</div>
          <div style={{ fontSize: 19, fontWeight: 600, letterSpacing: "-0.02em", color: C.ink }}>{(duration / 60).toFixed(1)}h</div>
        </div>
        <div style={{ fontSize: 11, color: C.gray, fontWeight: 600, textAlign: "right" }}>Hors TJM — n'affecte pas ton taux horaire</div>
      </div>
      ) : (
      <div style={{ marginTop: 14, background: C.bg, borderRadius: 10, padding: 12, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600 }}>Valeur session estimée</div>
          <div style={{ fontSize: 19, fontWeight: 600, letterSpacing: "-0.02em", color: C.ink }}>{euro1(value)}</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600 }}>Taux horaire session</div>
          <div style={{ fontSize: 19, fontWeight: 600, letterSpacing: "-0.02em", color: value > 0 ? zone.color : C.faint }}>{euro1(rate)}/h</div>
        </div>
      </div>
      )}

      <div style={{ marginTop: 12, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
        <Btn onClick={logSession} style={blockedClients.length ? { opacity: 0.4, pointerEvents: "none" } : undefined}>
          <Clock size={13} />Enregistrer la session
        </Btn>
        {blockedClients.length > 0 && (
          <span style={{ fontSize: 12, color: C.red }}>
            Renseigne le nombre de posts par semaine de {blockedClients.map((c) => c.name.split(" / ")[0]).join(", ")} pour enregistrer.
          </span>
        )}
      </div>
    </Card>
  );
}

function TempsTab({ data, save, computed }) {
  const [selectedMonth, setSelectedMonth] = useState(computed.curMonth);

  const availableMonths = useMemo(() => {
    const set = new Set([computed.curMonth]);
    data.sessions.forEach((s) => { if (s.date) set.add(s.date.slice(0, 7)); });
    return Array.from(set).sort();
  }, [data.sessions, computed.curMonth]);

  const stats = useMemo(() => computeMonthStats(data, selectedMonth), [data, selectedMonth]);
  const monthSessions = stats.monthSessions.slice().sort((a, b) => (a.date < b.date ? 1 : -1));
  const isCurrentMonth = selectedMonth === computed.curMonth;

  const idx = availableMonths.indexOf(selectedMonth);
  const goPrev = () => { if (idx > 0) setSelectedMonth(availableMonths[idx - 1]); };
  const goNext = () => { if (idx < availableMonths.length - 1) setSelectedMonth(availableMonths[idx + 1]); };

  const deleteSession = (id) => save({ ...data, sessions: data.sessions.filter((s) => s.id !== id) });

  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState({});
  const startEdit = (s) => { setEditingId(s.id); setDraft({ date: s.date, duration: s.duration, value: s.value, note: s.note || "", category: s.category }); };
  const cancelEdit = () => setEditingId(null);
  const saveEdit = (id) => {
    save({
      ...data,
      sessions: data.sessions.map((s) => (s.id === id ? {
        ...s, date: draft.date, category: draft.category,
        duration: Number(draft.duration) || 0,
        value: draft.category === "rdv" ? 0 : Number(draft.value) || 0,
        note: draft.note,
      } : s)),
    });
    setEditingId(null);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <button onClick={goPrev} disabled={idx <= 0} style={{
          border: `1px solid ${C.line}`, background: "#fff", borderRadius: 8, width: 30, height: 30,
          cursor: idx <= 0 ? "default" : "pointer", opacity: idx <= 0 ? 0.3 : 1, fontFamily: FONT, fontSize: 14,
        }}>‹</button>
        <div style={{ fontSize: 13, fontWeight: 600, color: C.ink, minWidth: 90, textAlign: "center" }}>
          {monthLabel(selectedMonth)} {isCurrentMonth && <span style={{ color: C.orange, fontWeight: 600 }}>· en cours</span>}
        </div>
        <button onClick={goNext} disabled={idx >= availableMonths.length - 1} style={{
          border: `1px solid ${C.line}`, background: "#fff", borderRadius: 8, width: 30, height: 30,
          cursor: idx >= availableMonths.length - 1 ? "default" : "pointer", opacity: idx >= availableMonths.length - 1 ? 0.3 : 1, fontFamily: FONT, fontSize: 14,
        }}>›</button>
        {!isCurrentMonth && (
          <Btn small variant="ghost" onClick={() => setSelectedMonth(computed.curMonth)}>Revenir au mois en cours</Btn>
        )}
      </div>

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        <Card style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600 }}>Temps de prod</div>
          <div style={{ fontSize: 24, fontWeight: 600, letterSpacing: "-0.02em", color: C.ink }}>{stats.totalHours.toFixed(1)}h</div>
        </Card>
        <Card style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600 }}>Taux horaire du mois</div>
          <div style={{ fontSize: 24, fontWeight: 600, letterSpacing: "-0.02em", color: C.ink }}>{euro1(stats.tjmHoraire)}/h</div>
          <div style={{ fontSize: 11, color: C.faint, marginTop: 2 }}>sur un MRR de {euro(stats.mrrActif)}{stats.mrrSource !== "courant" ? ` (${stats.mrrSource})` : ""}</div>
        </Card>
        <Card style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600 }}>Sessions loguées</div>
          <div style={{ fontSize: 24, fontWeight: 600, letterSpacing: "-0.02em", color: C.ink }}>{monthSessions.length}</div>
        </Card>
        <Card style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600 }}>Temps investissement (hors TJM)</div>
          <div style={{ fontSize: 24, fontWeight: 600, letterSpacing: "-0.02em", color: C.gray }}>{(stats.investMinutes / 60).toFixed(1)}h</div>
          <div style={{ fontSize: 10, color: C.gray, marginTop: 2 }}>Prospection + livrables offerts</div>
        </Card>
        <Card style={{ flex: 1, minWidth: 200 }}>
          <div style={{ fontSize: 11, color: C.gray, fontWeight: 600 }}>RDV / visio (hors TJM)</div>
          <div style={{ fontSize: 24, fontWeight: 600, letterSpacing: "-0.02em", color: C.gray }}>{(stats.rdvMinutes / 60).toFixed(1)}h</div>
          <div style={{ fontSize: 10, color: C.gray, marginTop: 2 }}>Temps total du mois : {(stats.allMinutes / 60).toFixed(1)}h</div>
        </Card>
      </div>

      {isCurrentMonth && <SessionForm data={data} save={save} />}

      <Card>
        <SectionTitle>Taux horaire par catégorie — {monthLabel(selectedMonth)}</SectionTitle>
        {stats.categoryBreakdown.length === 0 ? (
          <div style={{ fontSize: 12, color: C.gray }}>Aucune donnée pour ce mois.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {stats.categoryBreakdown.map((cat) => {
              const zone = zoneColorForTJM(cat.rate, data.settings.tjmSeuils);
              const barMax = Math.max(...stats.categoryBreakdown.map((c) => c.hours), 1);
              return (
                <div key={cat.key} style={{ display: "flex", alignItems: "center", gap: 12, padding: "8px 10px", borderRadius: 8, background: C.bg }}>
                  <div style={{ width: 150, fontSize: 12, fontWeight: 600, color: C.ink, flexShrink: 0 }}>{cat.label}</div>
                  <div style={{ flex: 1, height: 8, background: "#eee", borderRadius: 4, overflow: "hidden" }}>
                    <div style={{ width: `${(cat.hours / barMax) * 100}%`, height: "100%", background: zone.color, borderRadius: 4 }} />
                  </div>
                  <div style={{ width: 55, fontSize: 11, color: C.gray, textAlign: "right", flexShrink: 0 }}>{cat.hours.toFixed(1)}h</div>
                  <div style={{ width: 75, fontSize: 12, fontWeight: 600, color: zone.color, textAlign: "right", flexShrink: 0 }}>{euro1(cat.rate)}/h</div>
                </div>
              );
            })}
          </div>
        )}
      </Card>

      <Card>
        <SectionTitle>Historique — {monthLabel(selectedMonth)}</SectionTitle>
        {monthSessions.length === 0 ? (
          <div style={{ fontSize: 12, color: C.gray }}>Aucune session loguée ce mois-là.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {monthSessions.map((s) => {
              const rate = s.duration > 0 ? s.value / (s.duration / 60) : 0;
              const zone = zoneColorForTJM(rate, data.settings.tjmSeuils);
              const who = s.clients?.length > 0 ? s.clients.map((c) => c.name?.split(" / ")[0]).join(", ") : "";
              if (editingId === s.id) {
                return (
                  <div key={s.id} style={{ padding: 12, borderRadius: 10, background: C.sunken, border: `1px solid ${C.line}` }}>
                    <div style={{ fontSize: 12, color: C.gray, marginBottom: 8 }}>{who || "Sans client"}</div>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-end" }}>
                      <Field label="Date"><input type="date" style={inputStyle} value={draft.date} onChange={(e) => setDraft({ ...draft, date: e.target.value })} /></Field>
                      <Field label="Catégorie">
                        <select style={inputStyle} value={draft.category} onChange={(e) => setDraft({ ...draft, category: e.target.value })}>
                          {CATEGORIES.map((c) => <option key={c.key} value={c.key}>{c.label}</option>)}
                        </select>
                      </Field>
                      <Field label="Durée (min)"><input type="number" style={{ ...inputStyle, width: 80 }} value={draft.duration} onChange={(e) => setDraft({ ...draft, duration: e.target.value })} /></Field>
                      {draft.category !== "rdv" && (
                        <Field label="Valeur (€)"><input type="number" step="0.01" style={{ ...inputStyle, width: 90 }} value={draft.value} onChange={(e) => setDraft({ ...draft, value: e.target.value })} /></Field>
                      )}
                      <Field label="Note"><input style={{ ...inputStyle, width: 200 }} value={draft.note} onChange={(e) => setDraft({ ...draft, note: e.target.value })} /></Field>
                    </div>
                    <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                      <Btn small onClick={() => saveEdit(s.id)}><Check size={13} />Enregistrer</Btn>
                      <Btn small variant="ghost" onClick={cancelEdit}>Annuler</Btn>
                    </div>
                  </div>
                );
              }
              return (
                <div key={s.id} style={{
                  display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10,
                  padding: "8px 10px", borderRadius: 8, background: C.sunken, fontSize: 13,
                }}>
                  <div style={{ minWidth: 0 }}>
                    <strong style={{ fontWeight: 600 }}>{new Date(s.date).toLocaleDateString("fr-FR")}</strong> · {CATEGORIES.find((c) => c.key === s.category)?.label} · {s.duration} min
                    {who && <span style={{ color: C.gray }}> — {who}</span>}
                    {s.note && <span style={{ color: C.faint }}> · {s.note}</span>}
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0 }}>
                    {s.category === "rdv"
                      ? <span style={{ fontWeight: 500, color: C.gray, fontSize: 12 }}>hors TJM</span>
                      : <span style={{ fontWeight: 600, color: zone.color }}>{euro1(rate)}/h</span>}
                    <button onClick={() => startEdit(s)} aria-label="Modifier la session" style={{ border: "none", background: "transparent", cursor: "pointer", padding: 4, color: C.gray, fontSize: 12, fontFamily: FONT }}>Modifier</button>
                    <button onClick={() => deleteSession(s.id)} aria-label="Supprimer la session" style={{ border: "none", background: "transparent", cursor: "pointer", padding: 4 }}>
                      <Trash2 size={13} color={C.faint} />
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </Card>
    </div>
  );
}

/* ============================== REGLAGES TAB ============================== */
function ReglagesTab({ data, save, lastExport, onBackup }) {
  const [s, setS] = useState(data.settings);
  const [newMonth, setNewMonth] = useState(monthKey(new Date()));
  const [newMrr, setNewMrr] = useState(data.mrrHistory[data.mrrHistory.length - 1]?.mrr || 0);
  const [reminderText, setReminderText] = useState("");
  const [reminderDate, setReminderDate] = useState("");
  const [reminderClient, setReminderClient] = useState("");
  const [mrrConfirm, setMrrConfirm] = useState("");
  const [caMonth, setCaMonth] = useState(monthKey(new Date()));
  const [caAmount, setCaAmount] = useState(0);
  const [caConfirm, setCaConfirm] = useState("");

  const commit = (patch) => {
    const next = { ...s, ...patch };
    setS(next);
    save({ ...data, settings: next });
  };

  const addMrrPoint = () => {
    if (!newMonth || !newMrr) {
      setMrrConfirm("Renseigne un mois et un montant valides.");
      return;
    }
    const exists = data.mrrHistory.find((h) => h.month === newMonth);
    const history = exists
      ? data.mrrHistory.map((h) => (h.month === newMonth ? { ...h, mrr: newMrr } : h))
      : [...data.mrrHistory, { month: newMonth, mrr: newMrr }];
    save({ ...data, mrrHistory: history });
    setMrrConfirm(`✅ ${monthLabel(newMonth)} enregistré à ${euro(newMrr)}`);
    setTimeout(() => setMrrConfirm(""), 3000);
  };

  const addCaPoint = () => {
    if (!caMonth) {
      setCaConfirm("Renseigne un mois valide.");
      return;
    }
    const hist = data.caHistory || [];
    const exists = hist.find((h) => h.month === caMonth);
    const next = exists
      ? hist.map((h) => (h.month === caMonth ? { ...h, ca: caAmount } : h))
      : [...hist, { month: caMonth, ca: caAmount }];
    save({ ...data, caHistory: next });
    setCaConfirm(`✅ ${monthLabel(caMonth)} enregistré à ${euro(caAmount)}`);
    setTimeout(() => setCaConfirm(""), 3000);
  };
  const removeCaPoint = (month) => save({ ...data, caHistory: (data.caHistory || []).filter((h) => h.month !== month) });

  const addReminder = () => {
    if (!reminderText.trim()) return;
    save({ ...data, reminders: [...data.reminders, { id: uid(), text: reminderText, date: reminderDate, clientId: reminderClient || null, done: false }] });
    setReminderText(""); setReminderDate(""); setReminderClient("");
  };
  const removeReminder = (id) => save({ ...data, reminders: data.reminders.filter((r) => r.id !== id) });

  const [snapshots] = useState(() => listSnapshots());
  const restoreSnapshot = (day) => {
    const snap = readSnapshot(day);
    if (!snap) return;
    const ok = window.confirm(`Revenir à l'état du ${new Date(day).toLocaleDateString("fr-FR")} ? Les modifications faites depuis seront perdues.`);
    if (ok) save(snap);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <Card>
        <SectionTitle right={<Btn small onClick={onBackup}><Download size={13} />Sauvegarder maintenant</Btn>}>Sauvegardes</SectionTitle>
        <p style={{ margin: "0 0 14px", fontSize: 13, color: C.gray, lineHeight: 1.55, maxWidth: 620 }}>
          {lastExport
            ? `Dernière sauvegarde téléchargée le ${lastExport.toLocaleDateString("fr-FR")}. `
            : "Aucune sauvegarde téléchargée pour l'instant. "}
          Seul un fichier téléchargé te protège d'un vidage du navigateur ou d'un changement d'appareil.
        </p>
        <div style={{ fontSize: 13, fontWeight: 500, color: C.ink, marginBottom: 8 }}>Instantanés automatiques (7 derniers jours)</div>
        {snapshots.length === 0 ? (
          <div style={{ fontSize: 13, color: C.faint }}>Le premier instantané sera créé à ta prochaine modification.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column" }}>
            {snapshots.map((sn, k) => (
              <div key={sn.day} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10,
                padding: "9px 0", borderTop: k > 0 ? `1px solid ${C.line}` : "none", fontSize: 13 }}>
                <span>
                  {new Date(sn.day).toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" })}
                  <span style={{ color: C.faint }}>{`, ${sn.clients} fiches, ${sn.sessions} sessions`}</span>
                </span>
                <Btn small variant="ghost" onClick={() => restoreSnapshot(sn.day)}>Restaurer</Btn>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card>
        <SectionTitle>Seuils du taux horaire (€/h brut)</SectionTitle>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <Field label="Rouge → orange"><input type="number" style={{ ...inputStyle, width: 90 }} value={s.tjmSeuils[0]} onChange={(e) => commit({ tjmSeuils: [parseFloat(e.target.value) || 0, s.tjmSeuils[1], s.tjmSeuils[2]] })} /></Field>
          <Field label="Orange → vert"><input type="number" style={{ ...inputStyle, width: 90 }} value={s.tjmSeuils[1]} onChange={(e) => commit({ tjmSeuils: [s.tjmSeuils[0], parseFloat(e.target.value) || 0, s.tjmSeuils[2]] })} /></Field>
          <Field label="Vert → bleu"><input type="number" style={{ ...inputStyle, width: 90 }} value={s.tjmSeuils[2]} onChange={(e) => commit({ tjmSeuils: [s.tjmSeuils[0], s.tjmSeuils[1], parseFloat(e.target.value) || 0] })} /></Field>
        </div>
      </Card>

      <Card>
        <SectionTitle>Objectifs MRR</SectionTitle>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <Field label="Seuil Phase 1 (€)"><input type="number" style={{ ...inputStyle, width: 110 }} value={s.mrrObjectifPhase1} onChange={(e) => commit({ mrrObjectifPhase1: parseFloat(e.target.value) || 0 })} /></Field>
          <Field label="Cible stabilisation (€)"><input type="number" style={{ ...inputStyle, width: 110 }} value={s.mrrObjectifPrincipal} onChange={(e) => commit({ mrrObjectifPrincipal: parseFloat(e.target.value) || 0 })} /></Field>
        </div>
      </Card>

      <Card>
        <SectionTitle>Finances</SectionTitle>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="Taux URSSAF (%)"><input type="number" step="0.1" style={{ ...inputStyle, width: 100 }} value={s.urssafRate} onChange={(e) => commit({ urssafRate: parseFloat(e.target.value) || 0 })} /></Field>
          <Field label="Coûts outils (€/mois)"><input type="number" style={{ ...inputStyle, width: 100 }} value={s.coutsOutils} onChange={(e) => commit({ coutsOutils: parseFloat(e.target.value) || 0 })} /></Field>
          <Field label="Frais Stripe (%)"><input type="number" step="0.1" style={{ ...inputStyle, width: 90 }} value={s.stripeFeeRate} onChange={(e) => commit({ stripeFeeRate: parseFloat(e.target.value) || 0 })} /></Field>
          <Field label="Tarif réf. livrables (€/h)"><input type="number" style={{ ...inputStyle, width: 100 }} value={s.tarifReferenceLivrable} onChange={(e) => commit({ tarifReferenceLivrable: parseFloat(e.target.value) || 0 })} /></Field>
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 500, color: C.ink, paddingBottom: 9 }}>
            <input type="checkbox" checked={s.rcProSouscrite} onChange={(e) => commit({ rcProSouscrite: e.target.checked })} />
            RC Pro souscrite
          </label>
        </div>
      </Card>

      <Card>
        <SectionTitle>Seuils fiscaux et CA encaissé</SectionTitle>
        <p style={{ margin: "-6px 0 12px", fontSize: 13, color: C.gray, lineHeight: 1.55, maxWidth: 640 }}>
          Les seuils sont des estimations à valider avec un expert-comptable. Saisis ici le CA réellement encaissé chaque mois : sans saisie, le tableau de bord projette le MRR actuel.
        </p>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <Field label="Repère prudent (€ de CA/an)"><input type="number" style={{ ...inputStyle, width: 120, background: "#FFF9C4" }} value={s.caRepereLiberatoire ?? 29000} onChange={(e) => commit({ caRepereLiberatoire: parseFloat(e.target.value) || 0 })} /></Field>
          <Field label="Plafond haut estimé (€)"><input type="number" style={{ ...inputStyle, width: 120, background: "#FFF9C4" }} value={s.caPlafondMax ?? 33000} onChange={(e) => commit({ caPlafondMax: parseFloat(e.target.value) || 0 })} /></Field>
          <Field label="Franchise de TVA (€)"><input type="number" style={{ ...inputStyle, width: 120, background: "#FFF9C4" }} value={s.seuilTva ?? 37500} onChange={(e) => commit({ seuilTva: parseFloat(e.target.value) || 0 })} /></Field>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap", marginTop: 16 }}>
          <Field label="Mois"><input type="month" style={inputStyle} value={caMonth} onChange={(e) => setCaMonth(e.target.value)} /></Field>
          <Field label="CA encaissé (€)"><input type="number" style={{ ...inputStyle, width: 120 }} value={caAmount} onChange={(e) => setCaAmount(parseFloat(e.target.value) || 0)} /></Field>
          <Btn onClick={addCaPoint}><Plus size={13} />Enregistrer</Btn>
          {caConfirm && <span style={{ fontSize: 12, color: C.green, fontWeight: 600 }}>{caConfirm}</span>}
        </div>
        {(data.caHistory || []).length > 0 && (
          <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 4 }}>
            {[...(data.caHistory || [])].sort((a, b) => (a.month < b.month ? 1 : -1)).map((h) => (
              <div key={h.month} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 12, padding: "6px 10px", background: C.bg, borderRadius: 8 }}>
                <span style={{ color: C.gray }}>{monthLabel(h.month)}</span>
                <span style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <b style={{ color: C.ink }}>{euro(h.ca)}</b>
                  <button onClick={() => removeCaPoint(h.month)} aria-label={`Supprimer ${monthLabel(h.month)}`} style={{ border: "none", background: "transparent", cursor: "pointer", padding: 2 }}><Trash2 size={13} color={C.gray} /></button>
                </span>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card>
        <SectionTitle>Pondération du temps par catégorie (% du MRR client)</SectionTitle>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          {Object.entries(s.poidsCategories).map(([k, v]) => (
            <Field key={k} label={k}>
              <input type="number" style={{ ...inputStyle, width: 80 }} value={v}
                onChange={(e) => commit({ poidsCategories: { ...s.poidsCategories, [k]: parseFloat(e.target.value) || 0 } })} />
            </Field>
          ))}
        </div>
        <div style={{ fontSize: 11, color: C.gray, marginTop: 8 }}>Total actuel : {Object.values(s.poidsCategories).reduce((a, b) => a + b, 0)}%</div>
      </Card>

      <Card>
        <SectionTitle>Historique MRR</SectionTitle>
        <p style={{ margin: "-6px 0 12px", fontSize: 13, color: C.gray }}>Le MRR du mois en cours est enregistré automatiquement. Utilise ce formulaire seulement pour corriger ou compléter un mois passé.</p>
        <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
          <Field label="Mois"><input type="month" style={inputStyle} value={newMonth} onChange={(e) => setNewMonth(e.target.value)} /></Field>
          <Field label="MRR (€)"><input type="number" style={{ ...inputStyle, width: 110 }} value={newMrr} onChange={(e) => setNewMrr(parseFloat(e.target.value) || 0)} /></Field>
          <Btn onClick={addMrrPoint}><Plus size={13} />Ajouter</Btn>
          {mrrConfirm && <span style={{ fontSize: 12, color: C.green, fontWeight: 600 }}>{mrrConfirm}</span>}
        </div>
        {data.mrrHistory.length > 0 && (
          <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 4 }}>
            {[...data.mrrHistory].sort((a, b) => (a.month < b.month ? 1 : -1)).map((h) => (
              <div key={h.month} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, padding: "6px 10px", background: C.bg, borderRadius: 8 }}>
                <span style={{ color: C.gray }}>{monthLabel(h.month)}</span>
                <span style={{ fontWeight: 600, color: C.ink }}>{euro(h.mrr)}</span>
              </div>
            ))}
          </div>
        )}
      </Card>

      <Card>
        <SectionTitle>Rappels manuels</SectionTitle>
        <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
          <input style={{ ...inputStyle, flex: 1, minWidth: 200 }} placeholder="ex : penser à relancer Nexus" value={reminderText} onChange={(e) => setReminderText(e.target.value)} />
          <select style={inputStyle} value={reminderClient} onChange={(e) => setReminderClient(e.target.value)}>
            <option value="">Aucun client lié</option>
            {data.clients.map((c) => <option key={c.id} value={c.id}>{c.name.split(" / ")[0]}</option>)}
          </select>
          <input type="date" style={inputStyle} value={reminderDate} onChange={(e) => setReminderDate(e.target.value)} />
          <Btn onClick={addReminder}><Plus size={13} />Ajouter</Btn>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {data.reminders.map((r) => {
            const cl = r.clientId && data.clients.find((c) => c.id === r.clientId);
            return (
              <div key={r.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: C.bg, borderRadius: 8, padding: "8px 10px" }}>
                <span style={{ fontSize: 12, textDecoration: r.done ? "line-through" : "none", color: r.done ? C.gray : C.ink }}>
                  {cl && <strong>{cl.name.split(" / ")[0]} — </strong>}
                  {r.text} {r.date && <span style={{ color: C.gray }}> · {new Date(r.date).toLocaleDateString("fr-FR")}</span>}
                </span>
                <button onClick={() => removeReminder(r.id)} style={{ border: "none", background: "transparent", cursor: "pointer" }}>
                  <Trash2 size={13} color={C.gray} />
                </button>
              </div>
            );
          })}
        </div>
      </Card>

      <div style={{ fontSize: 12, color: C.faint, textAlign: "center", padding: "4px 0 8px" }}>
        Version du 10/10/2026, mise à jour automatique via GitHub
      </div>
    </div>
  );
}

/* ============================== APP ============================== */
export default function VulpixPilotage() {
  const { data, loading, save, lastSaved, error, saving } = useVulpixData();
  const [tab, setTab] = useState("dashboard");
  const [lastExport, setLastExport] = useState(() => getLastExport());
  const computed = useComputed(data);

  // Fige automatiquement le MRR du mois en cours dans l'historique.
  // Au changement de mois, la dernière valeur du mois précédent reste figée.
  useEffect(() => {
    if (!data || !computed) return;
    const cur = computed.curMonth;
    const entry = (data.mrrHistory || []).find((h) => h.month === cur);
    if (entry && Number(entry.mrr) === computed.mrrActif) return;
    const mrrHistory = entry
      ? data.mrrHistory.map((h) => (h.month === cur ? { ...h, mrr: computed.mrrActif } : h))
      : [...(data.mrrHistory || []), { month: cur, mrr: computed.mrrActif }];
    save({ ...data, mrrHistory });
  }, [data, computed, save]);

  const toggleReminder = (id) => {
    save({ ...data, reminders: data.reminders.map((r) => (r.id === id ? { ...r, done: true } : r)) });
  };

  if (loading || !data || !computed) {
    return (
      <div style={{ fontFamily: FONT, minHeight: "100vh", background: C.bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: C.gray, fontSize: 13 }}>Chargement…</div>
      </div>
    );
  }

  const recos = computeRecommendations(data, computed);

  const tabs = [
    { key: "dashboard", label: "Pilotage", icon: GaugeIcon },
    { key: "clients", label: "Clients", icon: Users },
    { key: "temps", label: "Temps", icon: Clock },
    { key: "reglages", label: "Réglages", icon: SettingsIcon },
  ];

  const daysSinceExport = lastExport ? Math.floor((Date.now() - lastExport.getTime()) / 86400000) : null;
  const backupDue = daysSinceExport === null || daysSinceExport >= 7;
  const backupNow = () => { downloadBackup(data); setLastExport(getLastExport()); };

  return (
    <div className="vx-app" style={{ fontFamily: FONT, minHeight: "100vh", background: C.bg, color: C.ink }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600;700&display=swap');
        .vx-app { font-variant-numeric: tabular-nums; -webkit-font-smoothing: antialiased; }
        .vx-app input, .vx-app select, .vx-app textarea { outline: none; transition: border-color .15s, box-shadow .15s; }
        .vx-app input:focus, .vx-app select:focus, .vx-app textarea:focus { border-color: ${C.ink2} !important; box-shadow: 0 0 0 3px rgba(22,24,43,.08); }
        .vx-app button:focus-visible { outline: 2px solid ${C.orange}; outline-offset: 2px; }
        .vx-btn { transition: filter .15s, background .15s; }
        .vx-btn:hover { filter: brightness(0.94); }
        .vx-bottomnav { display: none; }
        @media (max-width: 640px) {
          .vx-topnav { display: none !important; }
          .vx-bottomnav { display: flex; }
          .vx-main { padding-bottom: 96px !important; }
          .vx-stats { grid-template-columns: repeat(2, 1fr) !important; }
          .vx-stat:nth-child(3) { border-left: none !important; }
          .vx-gauges { flex-wrap: nowrap !important; gap: 14px !important; }
          .vx-gauges > div { flex: 1 1 0 !important; }
          .vx-g-sub { display: none; }
          .vx-g-label { font-size: 12px !important; }
          .vx-g-value { font-size: 21px !important; }
        }
        @media (prefers-reduced-motion: reduce) { .vx-app * { transition: none !important; } }
      `}</style>

      <div className="vx-main" style={{ maxWidth: 1000, margin: "0 auto", padding: "20px 20px 60px" }}>
        <header style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 16, marginBottom: 20 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
            <img src="/favicon.png" alt="" width={34} height={34} style={{ borderRadius: 9, display: "block" }} />
            <div>
              <div style={{ fontSize: 16, fontWeight: 600, letterSpacing: "-0.01em", lineHeight: 1.2 }}>Vulpix</div>
              <div style={{ fontSize: 12, color: C.faint }}>
                {saving ? "Enregistrement…" : lastSaved ? `Enregistré à ${lastSaved.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}` : "Pilotage"}
              </div>
            </div>
          </div>
          <nav className="vx-topnav" style={{ display: "flex", gap: 2, background: C.card, border: `1px solid ${C.line}`, borderRadius: 12, padding: 3 }}>
            {tabs.map((t) => {
              const Icon = t.icon;
              const active = tab === t.key;
              return (
                <button key={t.key} onClick={() => setTab(t.key)} aria-current={active ? "page" : undefined} style={{
                  border: "none", cursor: "pointer", fontFamily: FONT, fontSize: 13, fontWeight: 500,
                  padding: "7px 13px", borderRadius: 9, display: "flex", alignItems: "center", gap: 7,
                  background: active ? C.sunken : "transparent", color: active ? C.ink : C.gray,
                }}>
                  <Icon size={15} color={active ? C.orange : C.faint} />{t.label}
                </button>
              );
            })}
          </nav>
        </header>

        {error && (
          <div role="alert" style={{ background: C.redBg, borderRadius: 12, padding: "11px 14px", marginBottom: 16, fontSize: 13, color: C.red }}>
            {error}. Tes modifications restent affichées mais ne sont peut-être pas enregistrées.
          </div>
        )}

        {backupDue && (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap",
            background: C.orangeBg, borderRadius: 12, padding: "10px 12px 10px 14px", marginBottom: 16 }}>
            <div style={{ fontSize: 13, color: C.orangeText }}>
              {daysSinceExport === null
                ? "Tes données n'existent que dans ce navigateur. Fais une première sauvegarde."
                : `Dernière sauvegarde il y a ${daysSinceExport} jours.`}
            </div>
            <Btn small onClick={backupNow}><Download size={13} />Sauvegarder maintenant</Btn>
          </div>
        )}

        {tab === "dashboard" && <DashboardTab data={data} computed={computed} recos={recos} onToggleReminder={toggleReminder} />}
        {tab === "clients" && <ClientsTab data={data} save={save} />}
        {tab === "temps" && <TempsTab data={data} save={save} computed={computed} />}
        {tab === "reglages" && <ReglagesTab data={data} save={save} lastExport={lastExport} onBackup={backupNow} />}
      </div>

      <nav className="vx-bottomnav" style={{
        position: "fixed", left: 0, right: 0, bottom: 0, zIndex: 40, background: "rgba(255,255,255,0.94)",
        backdropFilter: "blur(12px)", WebkitBackdropFilter: "blur(12px)", borderTop: `1px solid ${C.line}`,
        padding: "6px 8px calc(6px + env(safe-area-inset-bottom))", justifyContent: "space-around",
      }}>
        {tabs.map((t) => {
          const Icon = t.icon;
          const active = tab === t.key;
          return (
            <button key={t.key} onClick={() => setTab(t.key)} aria-current={active ? "page" : undefined} style={{
              border: "none", background: "transparent", cursor: "pointer", fontFamily: FONT,
              display: "flex", flexDirection: "column", alignItems: "center", gap: 3, padding: "6px 10px",
              fontSize: 11, fontWeight: 500, color: active ? C.ink : C.faint, minWidth: 64,
            }}>
              <Icon size={20} color={active ? C.orange : C.faint} />
              {t.label}
            </button>
          );
        })}
      </nav>
    </div>
  );
}
