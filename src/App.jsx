import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ---------------------------------------------------------------------------
// Connexion à la base de données Supabase (stockage clé/valeur simple)
// ---------------------------------------------------------------------------
const SUPABASE_URL = "https://vxfsoesgypgtzixmyzkl.supabase.co";
const SUPABASE_KEY = "sb_publishable_cNUJK-vN4wVqxGyBNXYUlw_kS6NUgU6";

async function storageGet(key) {
  try {
    const res = await fetch(
      `${SUPABASE_URL}/rest/v1/app_data?key=eq.${encodeURIComponent(key)}&select=value`,
      { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } }
    );
    if (!res.ok) return null;
    const rows = await res.json();
    if (!rows.length) return null;
    return JSON.parse(rows[0].value);
  } catch (e) {
    console.error("Erreur de lecture Supabase", e);
    return null;
  }
}

async function storageSet(key, value) {
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/app_data`, {
      method: "POST",
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: `Bearer ${SUPABASE_KEY}`,
        "Content-Type": "application/json",
        Prefer: "resolution=merge-duplicates",
      },
      body: JSON.stringify({ key, value: JSON.stringify(value) }),
    });
  } catch (e) {
    console.error("Erreur d'écriture Supabase", e);
  }
}
import {
  LayoutGrid,
  Users,
  Trophy,
  Search,
  Phone,
  Wallet,
  Package,
  TrendingUp,
  AlertTriangle,
  X,
  ChevronRight,
  ChevronLeft,
  Facebook,
  MessageCircle,
  Music2,
  Plus,
  Settings,
  FileBarChart,
  ClipboardList,
  MessageSquare,
  LogOut,
  Copy,
  Check,
  ShieldCheck,
  UserCog,
  Palette,
  Boxes,
  Truck,
  CheckCircle2,
  KeyRound,
  Trash2,
  ListChecks,
  Database,
  ArrowUpFromLine,
  Tag,
} from "lucide-react";

// ---------------------------------------------------------------------------
// Tokens
// ---------------------------------------------------------------------------
const ink = {
  canvas: "#070707",
  canvasDeep: "#030303",
  panel: "#121212",
  panelSoft: "#181818",
  card: "#151515",
  cardLight: "#1E1E1E",

  orange: "#FF6A00",
  orangeDeep: "#D95500",
  orangeSoft: "rgba(255,106,0,0.16)",

  green: "#38B54A",
  greenSoft: "rgba(56,181,74,0.16)",

  red: "#EF4444",
  redSoft: "rgba(239,68,68,0.16)",

  yellow: "#FFB238",
  yellowDeep: "#E0982A",
  blue: "#4FA8D8",
  violet: "#6B5CA5",

  text: "#FFFFFF",
  textSoft: "#C9C9C9",
  textMuted: "#7C7C7C",

  line: "rgba(255,255,255,0.08)",
  lineStrong: "rgba(255,255,255,0.14)",

  // alias — noms utilisés ailleurs dans le fichier, pointent vers la nouvelle palette
  petrol: "#FF6A00",
  petrolDeep: "#D95500",
  ochre: "#FFB238",
  ochreDeep: "#E0982A",
  rouge: "#EF4444",
  rougeDeep: "#B83A1E",
  ink900: "#FFFFFF",
  ink600: "#C9C9C9",
  ink300: "#7C7C7C",
  bleu: "#4FA8D8",
  bleuDeep: "#3684AD",
};

const LOGO_DATA_URI = "data:image/webp;base64,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";

function LogoMark({ size = 44 }) {
  return (
    <img
      src={LOGO_DATA_URI}
      alt="Le Sérigraphe"
      className="shrink-0 object-contain"
      style={{
        height: size,
        width: size,
        filter: "drop-shadow(0 10px 24px rgba(255,106,0,0.28))",
      }}
    />
  );
}

function PremiumCard({ children, className = "", style = {}, onClick }) {
  return (
    <div onClick={onClick} className={`app-card rounded-3xl ${className}`} style={style}>
      {children}
    </div>
  );
}

function StatusPill({ label, color }) {
  return (
    <span
      className="inline-flex items-center gap-1.5 rounded-full px-3 py-1 text-[11px] font-semibold whitespace-nowrap"
      style={{ background: `${color}22`, color }}
    >
      <span className="h-1.5 w-1.5 rounded-full" style={{ background: color }} />
      {label}
    </span>
  );
}

const STATUTS = [
  "Dossier à suivre",
  "En attente de validation",
  "Commande confirmée",
  "Conception",
  "En production",
  "Prêt / à livrer",
  "Livré",
];

const STATUT_COLOR = {
  "Dossier à suivre": ink.ink300,
  "En attente de validation": ink.ochre,
  "Commande confirmée": ink.petrol,
  Conception: ink.bleu,
  "En production": ink.petrol,
  "Prêt / à livrer": ink.petrolDeep,
  Livré: "#5FA85B",
};

const URGENCE_COLOR = {
  normal: ink.ink300,
  urgent: ink.ochre,
  tres_urgent: ink.rouge,
};

const SOURCE_ICON = { Facebook: Facebook, WhatsApp: MessageCircle, TikTok: Music2 };

const REGLAGES_DEFAUT = { seuilInactiviteJours: 14, seuilCommandeInactiveJours: 3, frequenceRelance: "hebdomadaire" };

const ADMIN_AUTH_DEFAUT = { nom: "Félix", motDePasse: "serigraphe2026" };

const CATEGORIES_DEFAUT = ["T-shirt", "Sachet", "Sac", "Tableau", "Casquette", "Carte de visite"];

const ROLE_OPTIONS = [
  { value: "commercial", label: "Commercial(e)" },
  { value: "graphiste", label: "Pôle graphique" },
  { value: "production", label: "Production" },
  { value: "livraison", label: "Livraison" },
];

const URGENCE_OPTIONS = [
  { value: "normal", label: "Normal" },
  { value: "urgent", label: "Urgent" },
  { value: "tres_urgent", label: "Très urgent" },
];

const POLE_STATUT = {
  graphiste: "Conception",
  production: "En production",
  livraison: "Prêt / à livrer",
};
const STATUT_SUIVANT = {
  Conception: "En production",
  "En production": "Prêt / à livrer",
  "Prêt / à livrer": "Livré",
};
const POLE_LABEL = {
  graphiste: "Pôle graphique",
  production: "Production",
  livraison: "Livraison",
};

const TEMPLATES_DEFAUT = [
  {
    id: "devis",
    titre: "Relance devis sans réponse",
    texte:
      "Bonjour {nom}, c'est Le Sérigraphe 👋 Vous aviez demandé un devis, êtes-vous toujours intéressé(e) ? On reste dispo pour finaliser votre commande.",
  },
  {
    id: "paiement",
    titre: "Relance solde à payer",
    texte:
      "Bonjour {nom}, petit rappel amical : il reste un solde à régler sur votre commande. Dites-nous quand ça vous arrange pour finaliser 🙏",
  },
  {
    id: "fidelisation",
    titre: "Relance fidélisation",
    texte:
      "Bonjour {nom}, ça fait un moment ! Besoin d'un réassort ou d'un nouveau projet ? On a de nouvelles offres qui pourraient vous plaire 😊",
  },
];

// ---------------------------------------------------------------------------
// Données de démonstration
// ---------------------------------------------------------------------------
const AUJOURD_HUI = new Date("2026-07-31");

const CLIENTS_INIT = [
  {
    id: "LSG-2026-0001",
    nom: "Chantal Adjovi",
    type: "client",
    telephone: "+229 97 12 34 56",
    source: "Facebook",
    dateEntree: "2026-05-03",
    statut: "Livré",
    commandes: [
      { id: 1, date: "2026-05-10", description: "T-shirts événement (50 pièces)", montant: 85000, montantPaye: 85000, couts: [] },
      { id: 2, date: "2026-06-20", description: "Casquettes brodées (30 pièces)", montant: 120000, montantPaye: 60000, couts: [] },
      { id: 3, date: "2026-07-22", description: "Réassort T-shirts", montant: 90000, montantPaye: 90000, couts: [
        { id: 1, description: "Achat T-shirts", montant: 35000 },
        { id: 2, description: "Déplacement fournisseur", montant: 5000 },
      ] },
    ],
  },
  {
    id: "LSG-2026-0002",
    nom: "Ets. Béton Rouge",
    type: "client",
    telephone: "+229 96 44 21 09",
    source: "WhatsApp",
    dateEntree: "2026-06-11",
    statut: "En production",
    commandes: [
      { id: 1, date: "2026-06-15", description: "Uniformes chantier", montant: 210000, montantPaye: 100000, couts: [] },
    ],
  },
  {
    id: "LSG-2026-0003",
    nom: "Yves Kpossou",
    type: "prestataire",
    telephone: "+229 95 30 18 77",
    source: "WhatsApp",
    dateEntree: "2026-04-02",
    statut: "Livré",
    commandes: [
      { id: 1, date: "2026-04-18", description: "Sous-traitance broderie", montant: 45000, montantPaye: 45000, couts: [] },
      { id: 2, date: "2026-06-02", description: "Sous-traitance DTF", montant: 38000, montantPaye: 38000, couts: [] },
    ],
  },
  {
    id: "LSG-2026-0004",
    nom: "Église Renaissance",
    type: "client",
    telephone: "+229 61 05 88 40",
    source: "TikTok",
    dateEntree: "2026-07-01",
    statut: "En attente de validation",
    commandes: [],
  },
  {
    id: "LSG-2026-0005",
    nom: "Sandrine Houngbédji",
    type: "client",
    telephone: "+229 97 88 12 03",
    source: "Facebook",
    dateEntree: "2026-03-14",
    statut: "Livré",
    commandes: [
      { id: 1, date: "2026-03-20", description: "Cartes de visite transparentes", montant: 35000, montantPaye: 35000, couts: [] },
      { id: 2, date: "2026-04-25", description: "Flyers + kakémono", montant: 60000, montantPaye: 60000, couts: [] },
      { id: 3, date: "2026-05-30", description: "T-shirts personnalisés", montant: 150000, montantPaye: 150000, couts: [] },
      { id: 4, date: "2026-06-28", description: "Packaging boutique", montant: 95000, montantPaye: 95000, couts: [] },
    ],
  },
  {
    id: "LSG-2026-0006",
    nom: "Marcel Dossou",
    type: "client",
    telephone: "+229 94 20 71 62",
    source: "TikTok",
    dateEntree: "2026-07-10",
    statut: "En attente de validation",
    commandes: [],
  },
  {
    id: "LSG-2026-0007",
    nom: "Collège Saint-Michel",
    type: "client",
    telephone: "+229 90 15 44 28",
    source: "WhatsApp",
    dateEntree: "2026-06-02",
    statut: "Dossier à suivre",
    commandes: [],
  },
  {
    id: "LSG-2026-0008",
    nom: "Ines Agossou",
    type: "client",
    telephone: "+229 66 77 09 15",
    source: "Facebook",
    dateEntree: "2026-06-25",
    statut: "Commande confirmée",
    commandes: [{ id: 1, date: "2026-07-01", description: "Sacs personnalisés", montant: 70000, montantPaye: 0, couts: [
      { id: 1, description: "Achat sacs vierges", montant: 30000 },
    ] }],
  },
];

// ---------------------------------------------------------------------------
// Aides
// ---------------------------------------------------------------------------
const fmt = (n) => n.toLocaleString("fr-FR") + " F";

function statutPaiement(cmd) {
  if (cmd.montantPaye <= 0) return "Non payé";
  if (cmd.montantPaye < cmd.montant) return "Acompte versé";
  return "Soldé";
}
const PAIEMENT_COLOR = { "Non payé": ink.rouge, "Acompte versé": ink.ochre, Soldé: "#5FA85B" };

function totalCouts(cmd) {
  return (cmd.couts || []).reduce((s, c) => s + c.montant, 0);
}
function margeCommande(cmd) {
  return cmd.montant - totalCouts(cmd);
}
function totalClient(c) {
  return c.commandes.reduce((s, cmd) => s + cmd.montant, 0);
}
function joursDepuis(dateStr) {
  return Math.floor((AUJOURD_HUI - new Date(dateStr)) / 86400000);
}
function derniereActivite(c) {
  if (c.commandes.length === 0) return c.dateEntree;
  return c.commandes.reduce((max, cmd) => (cmd.date > max ? cmd.date : max), c.commandes[0].date);
}
function maintenant() {
  const d = new Date();
  return { date: d.toISOString().slice(0, 10), heure: d.toTimeString().slice(0, 5) };
}
function dansLeMois(dateStr, ref = AUJOURD_HUI) {
  const d = new Date(dateStr);
  return d.getFullYear() === ref.getFullYear() && d.getMonth() === ref.getMonth();
}
function dansLeJour(dateStr, ref = AUJOURD_HUI) {
  return dateStr === ref.toISOString().slice(0, 10);
}

// ---------------------------------------------------------------------------
// Le tampon — badge numéro de commande fixe (élément signature)
// ---------------------------------------------------------------------------
function Tampon({ id, small, color = ink.petrol }) {
  return (
    <span
      className={`inline-flex items-center justify-center rounded-full border-2 border-dashed shrink-0 ${
        small ? "px-2 py-0.5 text-[10px]" : "px-3 py-1 text-xs"
      }`}
      style={{
        borderColor: color,
        color,
        fontFamily: "'Inter', monospace",
        letterSpacing: "0.02em",
        transform: "rotate(-2deg)",
        background: `${color}1F`,
      }}
    >
      {id}
    </span>
  );
}

function TypeBadge({ type }) {
  const isClient = type === "client";
  return (
    <span
      className="inline-flex items-center rounded px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wide"
      style={{
        background: isClient ? "rgba(31,75,76,0.1)" : "rgba(193,68,60,0.1)",
        color: isClient ? ink.petrol : ink.rouge,
      }}
    >
      {isClient ? "Client" : "Prestataire"}
    </span>
  );
}

function StatutBadge({ statut }) {
  const c = STATUT_COLOR[statut] || ink.ink300;
  return (
    <span
      className="inline-flex items-center gap-1.5 rounded px-2 py-1 text-xs font-medium"
      style={{ background: `${c}1A`, color: c }}
    >
      <span className="h-1.5 w-1.5 rounded-full" style={{ background: c }} />
      {statut}
    </span>
  );
}

const MESSAGES_PAR_ROLE = {
  commercial: [
    "Bonjour {nom}, cette journée sera une belle journée !",
    "{nom}, un sourire au téléphone s'entend aussi bien qu'en vrai — bonne journée !",
    "Astuce du jour, {nom} : réponds vite aux nouveaux contacts, la rapidité fait la différence.",
    "{nom}, chaque client mérite d'être écouté avant d'être vendu. Bonne journée à toi !",
    "Aujourd'hui, c'est un bon jour pour relancer tes clients fidèles, {nom} !",
    "{nom}, un client satisfait en amène deux autres. Fais de ton mieux aujourd'hui !",
    "Courage {nom}, chaque commande compte pour l'atelier. On avance ensemble !",
    "{nom}, la patience avec un client hésitant paie toujours à la fin. Bonne journée !",
  ],
  graphiste: [
    "Bonjour {nom}, que ton trait soit précis aujourd'hui !",
    "{nom}, un bon visuel vaut mille mots — donne le meilleur de ta créativité !",
    "Astuce du jour, {nom} : relis deux fois avant d'exporter, ça évite les allers-retours.",
    "{nom}, chaque conception que tu livres fait gagner du temps à toute l'équipe.",
    "Aujourd'hui, prends le temps de peaufiner les détails qui font la différence, {nom}.",
    "{nom}, ton travail donne vie aux idées des clients. Continue comme ça !",
    "Courage {nom}, une belle conception aujourd'hui, c'est une production plus fluide demain.",
  ],
  production: [
    "Bonjour {nom}, une production bien faite évite bien des retours !",
    "{nom}, chaque pièce qui sort de tes mains représente l'atelier. Fais-la bien.",
    "Astuce du jour, {nom} : vérifie la qualité avant de valider, ça rassure toute la chaîne.",
    "{nom}, ton rythme d'aujourd'hui fait avancer toute l'équipe.",
    "Bon courage {nom}, la précision d'aujourd'hui évite les reprises de demain.",
    "{nom}, une commande bien produite, c'est un client content. Merci pour ton sérieux.",
  ],
  livraison: [
    "Bonjour {nom}, chaque livraison à l'heure construit la confiance du client.",
    "{nom}, sois prudent sur la route, la commande compte sur toi pour arriver en bon état.",
    "Astuce du jour, {nom} : confirme toujours la réception avec le client avant de repartir.",
    "{nom}, un sourire à la livraison laisse une bonne impression durable.",
    "Bon courage {nom}, la dernière étape est aussi la plus visible pour le client.",
    "{nom}, merci de faire le lien final entre l'atelier et le client.",
  ],
};

function MessageDuJour({ nom, pole = "commercial" }) {
  const liste = MESSAGES_PAR_ROLE[pole] || MESSAGES_PAR_ROLE.commercial;
  const [message] = useState(() => {
    const tpl = liste[Math.floor(Math.random() * liste.length)];
    return tpl.replace("{nom}", nom);
  });
  return (
    <div
      className="rounded-3xl p-4 flex items-center gap-3"
      style={{ background: `${ink.ochre}1A`, border: `1px solid ${ink.ochre}` }}
    >
      <span className="text-xl shrink-0">☀️</span>
      <p className="text-sm font-medium" style={{ color: ink.ink900 }}>{message}</p>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Tableau de bord
// ---------------------------------------------------------------------------
// Système 1 — clients inactifs (relance générale : saluer, rappeler notre présence)
function clientsARelancer(clients, seuil) {
  const clientsOnly = clients.filter((c) => c.type === "client");
  const statutsExclus = ["Conception", "En production", "Prêt / à livrer", "Livré", "Dossier à suivre", "En attente de validation"];
  return clientsOnly.filter(
    (c) =>
      joursDepuis(derniereActivite(c)) > seuil &&
      c.commandes.length > 0 &&
      !statutsExclus.includes(c.statut)
  );
}

// Système 2 — commandes inactives : statut "Dossier à suivre" ou "En attente de validation" depuis trop longtemps
function commandesARelancer(clients, seuil) {
  const clientsOnly = clients.filter((c) => c.type === "client");
  const statutsConcernes = ["Dossier à suivre", "En attente de validation"];
  return clientsOnly.filter(
    (c) => statutsConcernes.includes(c.statut) && joursDepuis(derniereActivite(c)) > seuil
  );
}

function ARelancerListe({ clients, reglages, onSelect }) {
  const seuil = reglages.seuilInactiviteJours;
  const seuilCommande = reglages.seuilCommandeInactiveJours;
  const listeClients = clientsARelancer(clients, seuil);
  const listeCommandes = commandesARelancer(clients, seuilCommande);
  const map = new Map();
  listeClients.forEach((c) => map.set(c.id, c));
  listeCommandes.forEach((c) => map.set(c.id, c));
  const liste = [...map.values()];

  return (
    <div className="rounded-3xl overflow-hidden" style={{ border: `1px solid ${ink.line}` }}>
      {liste.length === 0 ? (
        <p className="text-xs p-4" style={{ color: ink.ink600 }}>Rien à relancer pour le moment.</p>
      ) : (
        liste.map((c, i) => {
          const commandeInactive = ["Dossier à suivre", "En attente de validation"].includes(c.statut);
          const joursInactif = joursDepuis(derniereActivite(c));
          return (
            <div
              key={c.id}
              onClick={() => onSelect(c)}
              className="flex items-center justify-between gap-3 px-4 py-3 cursor-pointer"
              style={{ background: i % 2 ? ink.panel : "transparent", borderTop: i ? `1px solid ${ink.line}` : "none", borderLeft: `3px solid ${ink.rouge}` }}
            >
              <div className="min-w-0">
                <div className="text-sm font-semibold mb-1 truncate" style={{ color: ink.ink900 }}>{c.nom}</div>
                <Tampon id={c.id} small color={ink.rouge} />
              </div>
              <div className="text-xs text-right shrink-0" style={{ color: ink.rouge }}>
                {commandeInactive ? `commande inactive — ${joursInactif} j` : `client inactif — ${joursInactif} j`}
              </div>
            </div>
          );
        })
      )}
    </div>
  );
}

function Dashboard({ clients, reglages, currentUser, missions, setView }) {
  const clientsOnly = clients.filter((c) => c.type === "client");
  const isAdminDash = currentUser.roles.includes("admin");
  const caPeriode = clientsOnly.reduce(
    (s, c) =>
      s +
      c.commandes
        .filter((cmd) => (isAdminDash ? dansLeMois(cmd.date) : dansLeJour(cmd.date)))
        .reduce((s2, cmd) => s2 + Number(cmd.montant || 0), 0),
    0
  );
  const nouveauxCeMois = clients.filter((c) => dansLeMois(c.dateEntree)).length;
  const seuil = reglages.seuilInactiviteJours;
  const totalARelancer = clientsARelancer(clients, seuil).length + commandesARelancer(clients, reglages.seuilCommandeInactiveJours).length;
  const statutsExclusInactivite = ["Conception", "En production", "Prêt / à livrer", "Livré"];
  const inactifs = clientsOnly.filter(
    (c) => joursDepuis(derniereActivite(c)) > seuil && c.commandes.length > 0 && !statutsExclusInactivite.includes(c.statut)
  );
  const parSource = ["Facebook", "WhatsApp", "TikTok"].map((s) => ({
    source: s,
    n: clients.filter((c) => c.source === s).length,
  }));
  const maxSource = Math.max(...parSource.map((s) => s.n), 1);
  const topClient = [...clientsOnly].sort((a, b) => totalClient(b) - totalClient(a))[0];

  const commandesEnAttente = clients.filter((c) =>
    ["Conception", "En production"].includes(c.statut)
  );
  const livraisonEnAttente = clients.filter((c) => c.statut === "Prêt / à livrer");
  const missionsAFaire = missions.filter((m) => m.statut === "a_faire" && (isAdminDash || m.assigneA === currentUser.nom));
  const missionsUrgentes = missionsAFaire.filter((m) => m.urgence === "urgent").length;
  const missionsTresUrgentes = missionsAFaire.filter((m) => m.urgence === "tres_urgent").length;

  const stat = (label, value, Icon, color) => (
    <div className="rounded-3xl p-5 flex-1 min-w-[180px]" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
      <div className="flex items-center justify-between mb-3">
        <span className="text-xs font-medium uppercase tracking-wide" style={{ color: ink.ink600 }}>
          {label}
        </span>
        <Icon size={16} style={{ color }} />
      </div>
      <div className="text-2xl font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: ink.ink900 }}>
        {value}
      </div>
    </div>
  );

  return (
    <div className="space-y-6 pb-4">
      {!currentUser.roles.includes("admin") && <MessageDuJour nom={currentUser.nom} pole="commercial" />}

      <div>
        <h1 className="text-2xl sm:text-3xl font-extrabold mb-1" style={{ color: ink.ink900 }}>
          {isAdminDash ? `Bonjour ${currentUser.nom} 👋` : `Salut ${currentUser.nom} 👋`}
        </h1>
        <p className="text-sm" style={{ color: ink.ink600 }}>
          {isAdminDash ? "Voici l'activité de l'atelier aujourd'hui." : "Prêt·e à avancer sur tes commandes ?"}
        </p>
      </div>

      <div
        className="rounded-3xl p-5 overflow-hidden relative"
        style={{ background: "linear-gradient(135deg, #151515 0%, #090909 100%)", border: `1px solid ${ink.line}` }}
      >
        <div className="absolute -right-10 -top-10 h-36 w-36 rounded-full blur-2xl" style={{ background: "rgba(255,106,0,0.25)" }} />
        <div className="relative z-10">
          <div className="text-sm mb-2" style={{ color: ink.ink600 }}>{isAdminDash ? "Chiffre d'affaires du mois" : "Chiffre d'affaires du jour"}</div>
          <div className="text-4xl font-extrabold mb-4" style={{ color: ink.ink900 }}>{fmt(caPeriode)}</div>
          <div className="flex flex-wrap items-center gap-2">
            <StatusPill label={`${nouveauxCeMois} nouveaux contacts`} color={ink.green} />
            <StatusPill label={`${commandesEnAttente.length} en cours`} color={ink.orange} />
          </div>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-2.5">
        <button
          onClick={() => setView("missions")}
          className="rounded-2xl p-3 text-left"
          style={{ background: ink.petrol, color: "#fff" }}
        >
          <div className="flex items-center justify-between mb-1.5">
            <ListChecks size={17} />
            {(missionsUrgentes > 0 || missionsTresUrgentes > 0) && (
              <span
                className="text-[9px] font-bold h-4 min-w-4 px-1 rounded-full flex items-center justify-center"
                style={{ background: missionsTresUrgentes > 0 ? URGENCE_COLOR.tres_urgent : URGENCE_COLOR.urgent }}
              >
                {missionsTresUrgentes + missionsUrgentes}
              </span>
            )}
          </div>
          <div className="text-xs font-semibold leading-tight">{isAdminDash ? "Attribution" : "Missions"}</div>
          <div className="text-[10px] opacity-80">{missionsAFaire.length} à faire</div>
        </button>

        <button
          onClick={() => setView("commande_attente")}
          className="rounded-2xl p-3 text-left"
          style={{ background: ink.ochre, color: "#fff" }}
        >
          <div className="flex items-center justify-between mb-1.5">
            <Package size={17} />
            <span className="text-sm font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800 }}>
              {commandesEnAttente.length}
            </span>
          </div>
          <div className="text-xs font-semibold leading-tight">Commande</div>
          <div className="text-[10px] opacity-80">en attente</div>
        </button>

        <button
          onClick={() => setView("livraison_attente")}
          className="rounded-2xl p-3 text-left"
          style={{ background: "#3E7C5A", color: "#fff" }}
        >
          <div className="flex items-center justify-between mb-1.5">
            <Truck size={17} />
            <span className="text-sm font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800 }}>
              {livraisonEnAttente.length}
            </span>
          </div>
          <div className="text-xs font-semibold leading-tight">Livraison</div>
          <div className="text-[10px] opacity-80">prêtes à livrer</div>
        </button>

        <button
          onClick={() => setView("a_relancer")}
          className="rounded-2xl p-3 text-left"
          style={{ background: ink.rouge, color: "#fff" }}
        >
          <div className="flex items-center justify-between mb-1.5">
            <AlertTriangle size={17} />
            <span className="text-sm font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800 }}>
              {totalARelancer}
            </span>
          </div>
          <div className="text-xs font-semibold leading-tight">À relancer</div>
          <div className="text-[10px] opacity-80">clients concernés</div>
        </button>

        {isAdminDash && (
          <button
            onClick={() => setView("bilan")}
            className="rounded-2xl p-3 text-left"
            style={{ background: "#6B5CA5", color: "#fff" }}
          >
            <div className="flex items-center justify-between mb-1.5">
              <FileBarChart size={17} />
            </div>
            <div className="text-xs font-semibold leading-tight">Bilan</div>
            <div className="text-[10px] opacity-80">chiffres, exports</div>
          </button>
        )}
      </div>

      <div className="flex flex-wrap gap-4">
        {stat(isAdminDash ? "Chiffre d'affaires (mois)" : "Chiffre d'affaires (jour)", fmt(caPeriode), Wallet, ink.petrol)}
        {stat("Nouveaux contacts (ce mois)", nouveauxCeMois, TrendingUp, ink.ochre)}
        {stat(`Clients inactifs (+${seuil}j)`, inactifs.length, AlertTriangle, ink.rouge)}
        {stat("Client le plus fidèle", topClient?.nom.split(" ")[0] || "—", Trophy, ink.petrol)}
      </div>

      <div className="grid md:grid-cols-2 gap-4">
        <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <h3 className="text-sm font-semibold mb-4" style={{ color: ink.ink900 }}>
            Contacts par source
          </h3>
          <div className="space-y-3">
            {parSource.map((s) => {
              const Icon = SOURCE_ICON[s.source];
              return (
                <div key={s.source} className="flex items-center gap-3">
                  <Icon size={14} style={{ color: ink.ink600 }} className="shrink-0" />
                  <span className="text-xs w-16 shrink-0" style={{ color: ink.ink600 }}>
                    {s.source}
                  </span>
                  <div className="flex-1 h-2.5 rounded-full overflow-hidden" style={{ background: ink.canvasDeep }}>
                    <div
                      className="h-full rounded-full"
                      style={{ width: `${(s.n / maxSource) * 100}%`, background: ink.petrol }}
                    />
                  </div>
                  <span className="text-xs font-semibold w-4 text-right" style={{ color: ink.ink900 }}>
                    {s.n}
                  </span>
                </div>
              );
            })}
          </div>
        </div>

        <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <h3 className="text-sm font-semibold mb-4" style={{ color: ink.ink900 }}>
            À relancer en priorité
          </h3>
          {inactifs.length === 0 ? (
            <p className="text-xs" style={{ color: ink.ink600 }}>
              Aucun client inactif pour le moment.
            </p>
          ) : (
            <div className="space-y-2">
              {inactifs.map((c) => (
                <div key={c.id} className="flex items-center justify-between text-xs">
                  <div className="flex items-center gap-2">
                    <Tampon id={c.id} small color={ink.rouge} />
                    <span style={{ color: ink.ink900 }}>{c.nom}</span>
                  </div>
                  <span style={{ color: ink.rouge }}>{joursDepuis(derniereActivite(c))} j sans activité</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Liste clients
// ---------------------------------------------------------------------------
function detailsCommande(c) {
  if (c.commandes && c.commandes.length > 0) {
    return c.commandes.map((cmd) => cmd.description).join(", ");
  }
  return c.besoin || "—";
}

function categoriesClient(c) {
  const cats = [...new Set((c.commandes || []).map((cmd) => cmd.categorie).filter(Boolean))];
  return cats.length ? cats.join(", ") : "—";
}

function ClientsList({ clients, onSelect, filter, setFilter, query, setQuery, statutForce }) {
  const [queryBesoin, setQueryBesoin] = useState("");

  const lignes = clients.flatMap((c) => (c.commandes || []).map((cmd) => ({ client: c, cmd })));

  const filtered = lignes.filter(({ client: c, cmd }) => {
    if (statutForce && !statutForce.includes(c.statut)) return false;
    if (filter !== "tous" && c.type !== filter) return false;
    if (query && !c.nom.toLowerCase().includes(query.toLowerCase()) && !c.id.toLowerCase().includes(query.toLowerCase()))
      return false;
    if (queryBesoin && !cmd.description.toLowerCase().includes(queryBesoin.toLowerCase())) return false;
    return true;
  });

  return (
    <div className="space-y-4">
      {statutForce && (
        <div className="rounded-2xl px-3 py-2 text-xs font-medium" style={{ background: `${ink.bleu}1A`, color: ink.bleu, border: `1px solid ${ink.bleu}` }}>
          Filtré sur : {statutForce.join(" / ")}
        </div>
      )}
      <div className="flex flex-wrap items-center gap-3">
        <div
          className="flex items-center gap-2 rounded-2xl px-3 py-2 flex-1 min-w-[180px]"
          style={{ background: ink.panel, border: `1px solid ${ink.line}` }}
        >
          <Search size={14} style={{ color: ink.ink600 }} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Nom ou numéro..."
            className="bg-transparent outline-none text-sm flex-1"
            style={{ color: ink.ink900 }}
          />
        </div>
        <div
          className="flex items-center gap-2 rounded-2xl px-3 py-2 flex-1 min-w-[180px]"
          style={{ background: ink.panel, border: `1px solid ${ink.bleu}` }}
        >
          <ListChecks size={14} style={{ color: ink.bleu }} />
          <input
            value={queryBesoin}
            onChange={(e) => setQueryBesoin(e.target.value)}
            placeholder="Article commandé (ex. t-shirt)..."
            className="bg-transparent outline-none text-sm flex-1"
            style={{ color: ink.ink900 }}
          />
        </div>
        {["tous", "client", "prestataire"].map((f) => (
          <button
            key={f}
            onClick={() => setFilter(f)}
            className="text-xs font-medium px-3 py-2 rounded-2xl capitalize transition"
            style={{
              background: filter === f ? ink.petrol : ink.panel,
              color: filter === f ? ink.panel : ink.ink600,
              border: `1px solid ${filter === f ? ink.petrol : ink.line}`,
            }}
          >
            {f}
          </button>
        ))}
      </div>

      <div className="space-y-3">
        {filtered.map(({ client: c, cmd }) => (
          <PremiumCard
            key={cmd.id}
            className="p-4 cursor-pointer"
            onClick={() => onSelect(c)}
            style={{ borderLeft: `3px solid ${STATUT_COLOR[c.statut] || ink.ink300}` }}
          >
            <div className="flex items-start justify-between gap-2 mb-1.5">
              <div className="min-w-0">
                <div className="text-sm font-bold truncate" style={{ color: ink.ink900 }}>{cmd.description}</div>
                <div className="text-[11px] flex items-center gap-1.5 flex-wrap mt-0.5" style={{ color: ink.ink600 }}>
                  <Tampon id={c.id} small /> {c.nom}
                </div>
              </div>
              <ChevronRight size={16} className="shrink-0" style={{ color: ink.ink300 }} />
            </div>
            <div className="flex flex-wrap gap-1.5 mb-2">
              <StatutBadge statut={c.statut} />
              {cmd.categorie && (
                <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded" style={{ background: `${ink.bleu}1A`, color: ink.bleu }}>
                  {cmd.categorie}
                </span>
              )}
              {cmd.nonConfirmee && (
                <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded" style={{ background: `${ink.ochre}1A`, color: ink.ochre }}>
                  Non confirmée
                </span>
              )}
            </div>
            <div className="flex items-center justify-between text-xs">
              <span style={{ color: ink.ink600 }}>{cmd.date}</span>
              {!cmd.nonConfirmee && (
                <span className="font-bold" style={{ fontFamily: "'Inter', monospace", color: ink.ink900 }}>{fmt(cmd.montant)}</span>
              )}
            </div>
          </PremiumCard>
        ))}
        {filtered.length === 0 && (
          <PremiumCard className="p-8 text-center">
            <p className="text-sm" style={{ color: ink.ink600 }}>Aucune commande trouvée.</p>
          </PremiumCard>
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Parcours (kanban)
// ---------------------------------------------------------------------------
function Parcours({ clients, onSelect }) {
  const tries = [...clients].sort((a, b) => STATUTS.indexOf(a.statut) - STATUTS.indexOf(b.statut));

  return (
    <div className="rounded-3xl overflow-hidden" style={{ border: `1px solid ${ink.line}` }}>
      {tries.map((c, i) => (
        <div
          key={c.id}
          onClick={() => onSelect(c)}
          className="flex items-center gap-3 px-4 py-3 cursor-pointer"
          style={{
            background: i % 2 ? ink.panel : "transparent",
            borderTop: i ? `1px solid ${ink.line}` : "none",
            borderLeft: `3px solid ${STATUT_COLOR[c.statut]}`,
          }}
        >
          <div className="min-w-0 flex-1">
            <div className="text-sm font-semibold mb-1 truncate" style={{ color: ink.ink900 }}>{c.nom}</div>
            <Tampon id={c.id} small />
          </div>
          <StatutBadge statut={c.statut} />
          <ChevronRight size={15} style={{ color: ink.ink300 }} className="shrink-0" />
        </div>
      ))}
      {tries.length === 0 && (
        <p className="text-xs p-4" style={{ color: ink.ink600 }}>Aucun client pour le moment.</p>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Fidélité
// ---------------------------------------------------------------------------
function Fidelite({ clients }) {
  const classement = [...clients.filter((c) => c.type === "client")]
    .sort((a, b) => totalClient(b) - totalClient(a))
    .slice(0, 8);

  return (
    <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
      <h3 className="text-sm font-semibold mb-1" style={{ color: ink.ink900 }}>
        Classement de fidélité
      </h3>
      <p className="text-xs mb-4" style={{ color: ink.ink600 }}>
        Basé sur le montant total dépensé. Les prestataires ne sont pas comptés.
      </p>
      <div className="space-y-2">
        {classement.map((c, i) => (
          <div
            key={c.id}
            className="flex items-center gap-3 rounded-2xl px-3 py-2.5"
            style={{ background: i === 0 ? `${ink.ochre}17` : "transparent" }}
          >
            <span
              className="text-sm font-bold w-6 text-center"
              style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: i === 0 ? ink.ochreDeep : ink.ink300 }}
            >
              {i + 1}
            </span>
            <Tampon id={c.id} small color={i === 0 ? ink.ochreDeep : ink.petrol} />
            <span className="text-sm font-medium flex-1" style={{ color: ink.ink900 }}>
              {c.nom}
            </span>
            <span className="text-xs" style={{ color: ink.ink600 }}>
              {c.commandes.length} commande{c.commandes.length > 1 ? "s" : ""}
            </span>
            <span
              className="text-sm font-semibold w-24 text-right"
              style={{ fontFamily: "'Inter', monospace", color: ink.ink900 }}
            >
              {fmt(totalClient(c))}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Réglages (admin)
// ---------------------------------------------------------------------------
function Reglages({ reglages, onSave, adminAuth, onSaveAdminAuth, categories, onSaveCategories }) {
  const [seuil, setSeuil] = useState(reglages.seuilInactiviteJours);
  const [seuilCommande, setSeuilCommande] = useState(reglages.seuilCommandeInactiveJours);
  const [frequence, setFrequence] = useState(reglages.frequenceRelance);
  const [saved, setSaved] = useState(false);
  const [nouveauMdp, setNouveauMdp] = useState("");
  const [mdpSaved, setMdpSaved] = useState(false);
  const [nouvelleCategorie, setNouvelleCategorie] = useState("");

  function handleSave() {
    onSave({
      seuilInactiviteJours: Number(seuil) || 14,
      seuilCommandeInactiveJours: Number(seuilCommande) || 3,
      frequenceRelance: frequence,
    });
    setSaved(true);
    setTimeout(() => setSaved(false), 1800);
  }

  function handleSaveMdp() {
    if (!nouveauMdp.trim()) return;
    onSaveAdminAuth({ ...adminAuth, motDePasse: nouveauMdp.trim() });
    setNouveauMdp("");
    setMdpSaved(true);
    setTimeout(() => setMdpSaved(false), 1800);
  }

  function ajouterCategorie() {
    const val = nouvelleCategorie.trim();
    if (!val || categories.includes(val)) return;
    onSaveCategories([...categories, val]);
    setNouvelleCategorie("");
  }

  function supprimerCategorie(cat) {
    onSaveCategories(categories.filter((c) => c !== cat));
  }

  return (
    <div className="space-y-4 max-w-md">
      <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-1" style={{ color: ink.ink900 }}>Système 1 — Clients inactifs</h3>
        <p className="text-xs mb-3" style={{ color: ink.ink600 }}>
          Au bout de combien de jours sans activité un client (déjà en cours de commande) passe "inactif" — utile pour le
          saluer, rappeler qu'on est présent.
        </p>
        <div className="flex items-center gap-2">
          <input
            type="number"
            min="1"
            value={seuil}
            onChange={(e) => setSeuil(e.target.value)}
            className="w-24 rounded-2xl px-3 py-2 text-sm"
            style={inputStyle}
          />
          <span className="text-sm" style={{ color: ink.ink600 }}>jours (ex. 14 = deux semaines)</span>
        </div>
      </div>

      <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-1" style={{ color: ink.ink900 }}>Système 2 — Commandes inactives</h3>
        <p className="text-xs mb-3" style={{ color: ink.ink600 }}>
          Au bout de combien de jours un dossier resté en statut "Dossier à suivre" ou "En attente de validation" doit
          être relancé.
        </p>
        <div className="flex items-center gap-2">
          <input
            type="number"
            min="1"
            value={seuilCommande}
            onChange={(e) => setSeuilCommande(e.target.value)}
            className="w-24 rounded-2xl px-3 py-2 text-sm"
            style={inputStyle}
          />
          <span className="text-sm" style={{ color: ink.ink600 }}>jours (ex. 3 = trois jours)</span>
        </div>
      </div>

      <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-1" style={{ color: ink.ink900 }}>Cadence des relances</h3>
        <p className="text-xs mb-3" style={{ color: ink.ink600 }}>
          La cadence que le commercial doit suivre pour relancer un client sans réponse. Dans ce prototype, elle s'affiche
          comme consigne au commercial — l'envoi automatique réel demanderait une intégration WhatsApp Business API.
        </p>
        <select value={frequence} onChange={(e) => setFrequence(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm" style={inputStyle}>
          <option value="quotidienne">Tous les jours</option>
          <option value="deux_jours">Tous les 2 jours</option>
          <option value="hebdomadaire">Toutes les semaines</option>
        </select>
      </div>

      <button onClick={handleSave} className="rounded-2xl px-4 py-2.5 text-sm font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
        {saved ? "Enregistré ✓" : "Enregistrer les réglages"}
      </button>

      <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-1 flex items-center gap-1.5" style={{ color: ink.ink900 }}>
          <Tag size={14} /> Catégories d'articles
        </h3>
        <p className="text-xs mb-3" style={{ color: ink.ink600 }}>
          Utilisées pour catégoriser chaque commande (ex. T-shirt, Sachet, Sac, Tableau) et pour filtrer la base de données clients par article.
        </p>
        <div className="flex flex-wrap gap-2 mb-3">
          {categories.map((cat) => (
            <span
              key={cat}
              className="text-xs font-medium px-2.5 py-1.5 rounded-2xl flex items-center gap-1.5"
              style={{ background: ink.canvasDeep, color: ink.ink900 }}
            >
              {cat}
              <button onClick={() => supprimerCategorie(cat)}>
                <X size={11} style={{ color: ink.rouge }} />
              </button>
            </span>
          ))}
          {categories.length === 0 && <span className="text-xs" style={{ color: ink.ink600 }}>Aucune catégorie pour l'instant.</span>}
        </div>
        <div className="flex gap-2">
          <input
            value={nouvelleCategorie}
            onChange={(e) => setNouvelleCategorie(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && ajouterCategorie()}
            placeholder="Ex. Mug"
            className="flex-1 rounded-2xl px-3 py-2 text-sm"
            style={inputStyle}
          />
          <button onClick={ajouterCategorie} className="rounded-2xl px-4 py-2 text-sm font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
            Ajouter
          </button>
        </div>
      </div>

      <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-1 flex items-center gap-1.5" style={{ color: ink.ink900 }}>
          <KeyRound size={14} /> Sécurité — ton mot de passe administrateur
        </h3>
        <p className="text-xs mb-3" style={{ color: ink.ink600 }}>Identifiant : {adminAuth.nom}</p>
        <input
          type="password"
          value={nouveauMdp}
          onChange={(e) => setNouveauMdp(e.target.value)}
          placeholder="Nouveau mot de passe"
          className="w-full rounded-2xl px-3 py-2 text-sm mb-2"
          style={inputStyle}
        />
        <button
          onClick={handleSaveMdp}
          disabled={!nouveauMdp.trim()}
          className="rounded-2xl px-4 py-2.5 text-sm font-semibold"
          style={{ background: nouveauMdp.trim() ? ink.petrol : ink.ink300, color: "#fff" }}
        >
          {mdpSaved ? "Mot de passe changé ✓" : "Changer le mot de passe"}
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Bilan (jour / mois — admin)
// ---------------------------------------------------------------------------
const PERIODE_OPTIONS = [
  { value: "jour", label: "Aujourd'hui" },
  { value: "deux_jours", label: "2 jours" },
  { value: "semaine", label: "1 semaine" },
  { value: "mois", label: "Ce mois" },
  { value: "deux_mois", label: "2 mois" },
  { value: "an", label: "1 an" },
  { value: "personnalise", label: "Personnalisé" },
];

function rangeForPeriode(periode, custom) {
  const end = new Date(AUJOURD_HUI);
  end.setHours(23, 59, 59, 999);
  let start = new Date(AUJOURD_HUI);
  if (periode === "jour") {
    start.setHours(0, 0, 0, 0);
  } else if (periode === "deux_jours") {
    start.setDate(start.getDate() - 1);
    start.setHours(0, 0, 0, 0);
  } else if (periode === "semaine") {
    start.setDate(start.getDate() - 6);
    start.setHours(0, 0, 0, 0);
  } else if (periode === "mois") {
    start = new Date(AUJOURD_HUI.getFullYear(), AUJOURD_HUI.getMonth(), 1);
  } else if (periode === "deux_mois") {
    start.setMonth(start.getMonth() - 2);
    start.setHours(0, 0, 0, 0);
  } else if (periode === "an") {
    start.setFullYear(start.getFullYear() - 1);
    start.setHours(0, 0, 0, 0);
  } else if (periode === "personnalise" && custom.debut && custom.fin) {
    return { start: new Date(custom.debut + "T00:00:00"), end: new Date(custom.fin + "T23:59:59") };
  }
  return { start, end };
}

function dansIntervalle(dateStr, start, end) {
  const d = new Date(dateStr + "T12:00:00");
  return d >= start && d <= end;
}

function Bilan({ clients }) {
  const [periode, setPeriode] = useState("mois");
  const [debutPerso, setDebutPerso] = useState("");
  const [finPerso, setFinPerso] = useState("");
  const { start, end } = rangeForPeriode(periode, { debut: debutPerso, fin: finPerso });
  const dansPeriode = (dateStr) => dansIntervalle(dateStr, start, end);

  const clientsOnly = clients.filter((c) => c.type === "client");
  const nouveaux = clients.filter((c) => dansPeriode(c.dateEntree));
  const commandesPeriode = clientsOnly.flatMap((c) => c.commandes.filter((cmd) => dansPeriode(cmd.date)));
  const ca = commandesPeriode.reduce((s, cmd) => s + cmd.montant, 0);
  const margeBrute = commandesPeriode.reduce((s, cmd) => s + margeCommande(cmd), 0);
  const panierMoyen = commandesPeriode.length ? Math.round(ca / commandesPeriode.length) : 0;
  const impaye = clientsOnly.flatMap((c) => c.commandes).reduce((s, cmd) => s + (cmd.montant - cmd.montantPaye), 0);

  const parDescription = {};
  commandesPeriode.forEach((cmd) => {
    const key = cmd.description;
    if (!parDescription[key]) parDescription[key] = { n: 0, total: 0 };
    parDescription[key].n += 1;
    parDescription[key].total += cmd.montant;
  });
  const topProduits = Object.entries(parDescription).sort((a, b) => b[1].n - a[1].n).slice(0, 5);

  const label = `${start.toLocaleDateString("fr-FR")} → ${end.toLocaleDateString("fr-FR")}`;

  function exporterBilanExcel() {
    const wb = XLSX.utils.book_new();
    const resume = [
      ["Bilan Le Sérigraphe", label],
      [],
      ["Nouveaux contacts", nouveaux.length],
      ["Chiffre d'affaires", ca],
      ["Marge brute", margeBrute],
      ["Commandes", commandesPeriode.length],
      ["Panier moyen", panierMoyen],
      ["Reste à encaisser (total)", impaye],
      [],
      ["Commandes qui reviennent le plus", "", "Occurrences", "Total (F)"],
      ...topProduits.map(([desc, v]) => [desc, "", v.n, v.total]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(resume);
    XLSX.utils.book_append_sheet(wb, ws, "Bilan");
    XLSX.writeFile(wb, `bilan-${periode}-${AUJOURD_HUI.toISOString().slice(0, 10)}.xlsx`);
  }

  function exporterContactsExcel() {
    const wb = XLSX.utils.book_new();
    const lignes = clients.map((c) => ({
      "N° client": c.id,
      Nom: c.nom,
      Type: c.type,
      Statut: c.statut,
      Téléphone: c.telephone,
      Source: c.source,
      "Date d'entrée": c.dateEntree,
      Besoin: c.besoin || "",
      "Total facturé (F)": totalClient(c),
      "Nombre de commandes": c.commandes.length,
    }));
    const ws = XLSX.utils.json_to_sheet(lignes);
    XLSX.utils.book_append_sheet(wb, ws, "Contacts");
    XLSX.writeFile(wb, `contacts-le-serigraphe-${AUJOURD_HUI.toISOString().slice(0, 10)}.xlsx`);
  }

  const card = (labelC, value) => (
    <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
      <div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>{labelC}</div>
      <div className="text-xl font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: ink.ink900 }}>{value}</div>
    </div>
  );

  return (
    <div className="space-y-5">
      <div className="flex flex-wrap items-center gap-2">
        {PERIODE_OPTIONS.map((p) => (
          <button
            key={p.value}
            onClick={() => setPeriode(p.value)}
            className="text-xs font-semibold px-3 py-1.5 rounded-2xl"
            style={{
              background: periode === p.value ? ink.petrol : ink.panel,
              color: periode === p.value ? "#fff" : ink.ink600,
              border: `1px solid ${periode === p.value ? ink.petrol : ink.line}`,
            }}
          >
            {p.label}
          </button>
        ))}
      </div>

      {periode === "personnalise" && (
        <div className="flex flex-wrap items-center gap-2">
          <input type="date" value={debutPerso} onChange={(e) => setDebutPerso(e.target.value)} className="rounded-2xl px-3 py-2 text-sm" style={inputStyle} />
          <span className="text-xs" style={{ color: ink.ink600 }}>→</span>
          <input type="date" value={finPerso} onChange={(e) => setFinPerso(e.target.value)} className="rounded-2xl px-3 py-2 text-sm" style={inputStyle} />
        </div>
      )}

      <p className="text-xs" style={{ color: ink.ink600 }}>
        Métriques pour <strong style={{ color: ink.ink900 }}>{label}</strong> — utile pour ta comptabilité ou un point avec un investisseur.
      </p>

      <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
        {card("Nouveaux contacts", nouveaux.length)}
        {card("Chiffre d'affaires", fmt(ca))}
        {card("Marge brute", fmt(margeBrute))}
        {card("Commandes", commandesPeriode.length)}
        {card("Panier moyen", fmt(panierMoyen))}
        {card("Reste à encaisser (total)", fmt(impaye))}
      </div>

      <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-3" style={{ color: ink.ink900 }}>Commandes qui reviennent le plus</h3>
        {topProduits.length === 0 ? (
          <p className="text-xs" style={{ color: ink.ink600 }}>Pas encore assez de commandes sur cette période pour un classement.</p>
        ) : (
          <div className="space-y-2">
            {topProduits.map(([desc, v]) => (
              <div key={desc} className="flex items-center justify-between text-sm">
                <span style={{ color: ink.ink900 }}>{desc}</span>
                <span className="text-xs" style={{ color: ink.ink600 }}>
                  {v.n}× · {fmt(v.total)}
                </span>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="rounded-3xl p-5" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-1" style={{ color: ink.ink900 }}>Exports</h3>
        <p className="text-xs mb-3" style={{ color: ink.ink600 }}>Fichiers Excel (.xlsx), ouvrables directement dans Excel ou Google Sheets.</p>
        <div className="flex flex-wrap gap-2">
          <button onClick={exporterBilanExcel} className="flex items-center gap-1.5 rounded-2xl px-3 py-2 text-xs font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
            <FileBarChart size={13} /> Exporter le bilan (Excel)
          </button>
          <button onClick={exporterContactsExcel} className="flex items-center gap-1.5 rounded-2xl px-3 py-2 text-xs font-semibold" style={{ background: ink.ochre, color: "#fff" }}>
            <Users size={13} /> Exporter les contacts (Excel)
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Utilisateurs (admin) — créer des comptes et leur affecter un rôle
// ---------------------------------------------------------------------------
function Utilisateurs({ utilisateurs, onSave }) {
  const [nom, setNom] = useState("");
  const [motDePasse, setMotDePasse] = useState("");
  const [rolesChoisis, setRolesChoisis] = useState([]);

  function toggleRole(value) {
    setRolesChoisis((r) => (r.includes(value) ? r.filter((x) => x !== value) : [...r, value]));
  }

  function ajouter() {
    if (!nom.trim() || !motDePasse.trim() || rolesChoisis.length === 0) return;
    const nouveau = { id: Date.now(), nom: nom.trim(), motDePasse: motDePasse.trim(), roles: rolesChoisis };
    onSave([...utilisateurs, nouveau]);
    setNom("");
    setMotDePasse("");
    setRolesChoisis([]);
  }

  function supprimer(id) {
    onSave(utilisateurs.filter((u) => u.id !== id));
  }

  return (
    <div className="space-y-4">
      <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-3" style={{ color: ink.ink900 }}>Ajouter un utilisateur</h3>
        <div className="space-y-2.5">
          <input
            value={nom}
            onChange={(e) => setNom(e.target.value)}
            placeholder="Prénom (ex. Sandrine)"
            className="w-full rounded-2xl px-3 py-2 text-sm"
            style={inputStyle}
          />
          <input
            value={motDePasse}
            onChange={(e) => setMotDePasse(e.target.value)}
            placeholder="Mot de passe"
            className="w-full rounded-2xl px-3 py-2 text-sm"
            style={inputStyle}
          />
          <div>
            <label className="text-[11px] font-medium block mb-1.5" style={{ color: ink.ink600 }}>
              Accès (un ou plusieurs)
            </label>
            <div className="flex flex-wrap gap-2">
              {ROLE_OPTIONS.map((r) => {
                const actif = rolesChoisis.includes(r.value);
                return (
                  <button
                    key={r.value}
                    onClick={() => toggleRole(r.value)}
                    className="text-xs font-medium px-3 py-1.5 rounded-2xl"
                    style={{
                      background: actif ? ink.petrol : ink.canvasDeep,
                      color: actif ? "#fff" : ink.ink600,
                      border: `1px solid ${actif ? ink.petrol : ink.line}`,
                    }}
                  >
                    {r.label}
                  </button>
                );
              })}
            </div>
          </div>
          <button
            onClick={ajouter}
            disabled={!nom.trim() || !motDePasse.trim() || rolesChoisis.length === 0}
            className="w-full rounded-2xl py-2.5 text-sm font-semibold"
            style={{
              background: nom.trim() && motDePasse.trim() && rolesChoisis.length > 0 ? ink.petrol : ink.ink300,
              color: "#fff",
            }}
          >
            Ajouter
          </button>
        </div>
      </div>

      <div className="rounded-3xl overflow-hidden" style={{ border: `1px solid ${ink.line}` }}>
        {utilisateurs.length === 0 ? (
          <p className="text-xs p-4" style={{ color: ink.ink600 }}>Aucun utilisateur créé pour le moment.</p>
        ) : (
          utilisateurs.map((u, i) => (
            <div
              key={u.id}
              className="flex items-center justify-between gap-3 px-4 py-3"
              style={{ background: i % 2 ? ink.panel : "transparent", borderTop: i ? `1px solid ${ink.line}` : "none" }}
            >
              <div>
                <div className="text-sm font-semibold" style={{ color: ink.ink900 }}>{u.nom}</div>
                <div className="text-[11px]" style={{ color: ink.ink600 }}>
                  {(u.roles || []).map((r) => ROLE_OPTIONS.find((o) => o.value === r)?.label).join(" · ")}
                </div>
              </div>
              <button onClick={() => supprimer(u.id)} className="p-2 rounded-2xl" style={{ background: ink.canvasDeep }}>
                <Trash2 size={14} style={{ color: ink.rouge }} />
              </button>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Mission spécifique — accessible à tout le monde
// ---------------------------------------------------------------------------
function Missions({ missions, personnes, currentUser, isAdmin, onAdd, onToggle, onDelete }) {
  const [destinataire, setDestinataire] = useState(personnes[0] || "");
  const [texte, setTexte] = useState("");
  const [urgence, setUrgence] = useState("normal");
  const [deadline, setDeadline] = useState("");
  const personnesAffichees = isAdmin ? personnes : [currentUser.nom];

  function assigner() {
    if (!destinataire || !texte.trim()) return;
    onAdd({
      id: Date.now(),
      assigneA: destinataire,
      texte: texte.trim(),
      statut: "a_faire",
      urgence,
      deadline: deadline || null,
      creePar: currentUser.nom,
    });
    setTexte("");
    setUrgence("normal");
    setDeadline("");
  }

  return (
    <div className="space-y-4">
      {isAdmin && (
        <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <h3 className="text-sm font-semibold mb-3" style={{ color: ink.ink900 }}>Assigner une mission</h3>
          <div className="space-y-2.5">
            <select value={destinataire} onChange={(e) => setDestinataire(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm" style={inputStyle}>
              {personnes.map((p) => (
                <option key={p} value={p}>{p}</option>
              ))}
            </select>
            <textarea
              value={texte}
              onChange={(e) => setTexte(e.target.value)}
              rows={2}
              placeholder="Ex. Préparer 20 t-shirts pour l'église Renaissance avant vendredi"
              className="w-full rounded-2xl px-3 py-2 text-sm"
              style={inputStyle}
            />
            <div>
              <label className="text-[11px] font-medium block mb-1.5" style={{ color: ink.ink600 }}>Deadline (date et heure)</label>
              <input
                type="datetime-local"
                value={deadline}
                onChange={(e) => setDeadline(e.target.value)}
                className="w-full rounded-2xl px-3 py-2 text-sm"
                style={inputStyle}
              />
            </div>
            <div>
              <label className="text-[11px] font-medium block mb-1.5" style={{ color: ink.ink600 }}>Niveau d'urgence</label>
              <div className="flex flex-wrap gap-2">
                {URGENCE_OPTIONS.map((u) => {
                  const actif = urgence === u.value;
                  return (
                    <button
                      key={u.value}
                      onClick={() => setUrgence(u.value)}
                      className="text-xs font-semibold px-3 py-1.5 rounded-2xl flex items-center gap-1.5"
                      style={{
                        background: actif ? URGENCE_COLOR[u.value] : ink.canvasDeep,
                        color: actif ? "#fff" : ink.ink600,
                        border: `1px solid ${actif ? URGENCE_COLOR[u.value] : ink.line}`,
                      }}
                    >
                      <span className="h-1.5 w-1.5 rounded-full" style={{ background: actif ? "#fff" : URGENCE_COLOR[u.value] }} />
                      {u.label}
                    </button>
                  );
                })}
              </div>
            </div>
            <button
              onClick={assigner}
              disabled={!texte.trim()}
              className="w-full rounded-2xl py-2.5 text-sm font-semibold"
              style={{ background: texte.trim() ? ink.petrol : ink.ink300, color: "#fff" }}
            >
              Assigner
            </button>
          </div>
        </div>
      )}

      <div className="space-y-2.5">
        {personnesAffichees.map((p) => {
          const mesMissions = missions.filter((m) => m.assigneA === p);
          const cestMoi = p === currentUser.nom;
          return (
            <div
              key={p}
              className="rounded-3xl overflow-hidden"
              style={{ border: `1px solid ${cestMoi ? ink.petrol : ink.line}` }}
            >
              <div className="flex items-center justify-between px-4 py-2.5" style={{ background: ink.panel }}>
                <span className="text-sm font-semibold" style={{ color: ink.ink900 }}>
                  {p} {cestMoi && <span style={{ color: ink.petrol }}>(toi)</span>}
                </span>
                <span
                  className="text-[10px] font-bold rounded-full h-4 w-4 flex items-center justify-center"
                  style={{ background: ink.canvasDeep, color: ink.ink600 }}
                >
                  {mesMissions.filter((m) => m.statut === "a_faire").length}
                </span>
              </div>
              {mesMissions.length === 0 ? (
                <p className="text-xs px-4 py-2.5" style={{ color: ink.ink600, background: ink.canvas }}>
                  Rien pour l'instant.
                </p>
              ) : (
                <div className="p-2.5 space-y-2" style={{ background: ink.canvas }}>
                  {mesMissions.map((m) => {
                    const coul = URGENCE_COLOR[m.urgence || "normal"];
                    const label = URGENCE_OPTIONS.find((u) => u.value === (m.urgence || "normal"))?.label;
                    return (
                      <div
                        key={m.id}
                        className="flex items-start gap-2.5 rounded-2xl p-3"
                        style={{ background: ink.panel, border: `1px solid ${ink.line}`, borderLeft: `3px solid ${coul}`, opacity: m.statut === "faite" ? 0.55 : 1 }}
                      >
                        <button
                          onClick={() => (isAdmin || cestMoi) && onToggle(m.id)}
                          className="mt-0.5 shrink-0"
                          style={{ color: m.statut === "faite" ? "#5FA85B" : ink.ink300 }}
                        >
                          <CheckCircle2 size={18} />
                        </button>
                        <div className="flex-1">
                          <span
                            className="text-sm block"
                            style={{ color: ink.ink900, textDecoration: m.statut === "faite" ? "line-through" : "none" }}
                          >
                            {m.texte}
                          </span>
                          {m.urgence && m.urgence !== "normal" && (
                            <span
                              className="text-[10px] font-semibold inline-flex items-center gap-1 mt-1 px-1.5 py-0.5 rounded"
                              style={{ background: `${coul}1A`, color: coul }}
                            >
                              <span className="h-1.5 w-1.5 rounded-full" style={{ background: coul }} />
                              {label}
                            </span>
                          )}
                          {m.deadline && (
                            <div
                              className="text-[11px] mt-1"
                              style={{ color: new Date(m.deadline) < new Date() && m.statut === "a_faire" ? ink.rouge : ink.ink600 }}
                            >
                              Deadline : {new Date(m.deadline).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}
                              {new Date(m.deadline) < new Date() && m.statut === "a_faire" ? " — dépassée" : ""}
                            </div>
                          )}
                        </div>
                        {isAdmin && (
                          <button onClick={() => onDelete(m.id)} className="shrink-0">
                            <Trash2 size={14} style={{ color: ink.rouge }} />
                          </button>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Base de données clients (admin) — filtres article / période / CA
// ---------------------------------------------------------------------------
function BaseClients({ clients, categories, onSelect, onImportContacts, importMessage, onClearImportMessage, isAdmin, onNouveauContact }) {
  const [query, setQuery] = useState("");
  const [filtreArticle, setFiltreArticle] = useState("tous");
  const [periode, setPeriode] = useState("an");
  const [debutPerso, setDebutPerso] = useState("");
  const [finPerso, setFinPerso] = useState("");
  const [caMin, setCaMin] = useState("");
  const [caMax, setCaMax] = useState("");

  const { start, end } = rangeForPeriode(periode, { debut: debutPerso, fin: finPerso });

  const filtres = clients.filter((c) => {
    if (query && !c.nom.toLowerCase().includes(query.toLowerCase()) && !c.id.toLowerCase().includes(query.toLowerCase()))
      return false;
    if (filtreArticle !== "tous" && !(c.commandes || []).some((cmd) => cmd.categorie === filtreArticle)) return false;
    const aUneCommandeDansPeriode = (c.commandes || []).some((cmd) => dansIntervalle(cmd.date, start, end));
    const dateEntreeDansPeriode = dansIntervalle(c.dateEntree, start, end);
    if (!aUneCommandeDansPeriode && !dateEntreeDansPeriode) return false;
    const ca = totalClient(c);
    if (caMin && ca < Number(caMin)) return false;
    if (caMax && ca > Number(caMax)) return false;
    return true;
  });

  function exporterExcel() {
    const wb = XLSX.utils.book_new();
    const lignes = filtres.map((c) => ({
      "N° client": c.id,
      Nom: c.nom,
      Type: c.type,
      Statut: c.statut,
      Téléphone: c.telephone,
      Source: c.source,
      "Date d'entrée": c.dateEntree,
      Besoin: c.besoin || "",
      "Articles commandés": categoriesClient(c),
      "Total facturé (F)": totalClient(c),
      "Nombre de commandes": c.commandes.length,
    }));
    const ws = XLSX.utils.json_to_sheet(lignes);
    XLSX.utils.book_append_sheet(wb, ws, "Contacts");
    XLSX.writeFile(wb, `contacts-le-serigraphe-${AUJOURD_HUI.toISOString().slice(0, 10)}.xlsx`);
  }

  return (
    <div className="space-y-4">
      <button
        onClick={onNouveauContact}
        className="w-full flex items-center justify-center gap-1.5 rounded-2xl py-2.5 text-sm font-semibold"
        style={{ background: ink.rouge, color: "#fff" }}
      >
        <Plus size={15} /> Nouveau contact
      </button>

      {isAdmin && (
        <div className="flex gap-2">
          {onImportContacts && (
            <label
              className="flex-1 flex items-center justify-center gap-1.5 rounded-2xl py-2.5 text-sm font-semibold cursor-pointer"
              style={{ background: ink.panel, border: `1px solid ${ink.line}`, color: ink.ink900 }}
            >
              <ArrowUpFromLine size={15} /> Importer
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) onImportContacts(file);
                  e.target.value = "";
                }}
              />
            </label>
          )}
          <button
            onClick={exporterExcel}
            className="flex-1 flex items-center justify-center gap-1.5 rounded-2xl py-2.5 text-sm font-semibold"
            style={{ background: ink.ochre, color: "#fff" }}
          >
            <FileBarChart size={15} /> Exporter ({filtres.length})
          </button>
        </div>
      )}

      {importMessage && (
        <div
          className="rounded-2xl px-3 py-2.5 text-xs font-medium flex items-center justify-between gap-2"
          style={{
            background: importMessage.type === "succes" ? "rgba(95,168,91,0.15)" : `${ink.rouge}1A`,
            color: importMessage.type === "succes" ? "#5FA85B" : ink.rouge,
          }}
        >
          <span>{importMessage.texte}</span>
          <button onClick={onClearImportMessage}>
            <X size={13} />
          </button>
        </div>
      )}

      <div
        className="flex items-center gap-2 rounded-2xl px-3 py-2"
        style={{ background: ink.panel, border: `1px solid ${ink.line}` }}
      >
        <Search size={14} style={{ color: ink.ink600 }} />
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="Chercher un nom ou un numéro..."
          className="bg-transparent outline-none text-sm flex-1"
          style={{ color: ink.ink900 }}
        />
      </div>

      <div className="rounded-3xl p-4 space-y-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <div>
          <label className="text-[11px] font-medium block mb-1.5" style={{ color: ink.ink600 }}>Filtrer par article</label>
          <select value={filtreArticle} onChange={(e) => setFiltreArticle(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm" style={inputStyle}>
            <option value="tous">Tous les articles</option>
            {categories.map((cat) => (
              <option key={cat} value={cat}>{cat}</option>
            ))}
          </select>
        </div>

        <div>
          <label className="text-[11px] font-medium block mb-1.5" style={{ color: ink.ink600 }}>Filtrer par période</label>
          <div className="flex flex-wrap gap-2 mb-2">
            {PERIODE_OPTIONS.map((p) => (
              <button
                key={p.value}
                onClick={() => setPeriode(p.value)}
                className="text-xs font-semibold px-3 py-1.5 rounded-2xl"
                style={{
                  background: periode === p.value ? ink.petrol : ink.canvasDeep,
                  color: periode === p.value ? "#fff" : ink.ink600,
                }}
              >
                {p.label}
              </button>
            ))}
          </div>
          {periode === "personnalise" && (
            <div className="flex flex-wrap items-center gap-2">
              <input type="date" value={debutPerso} onChange={(e) => setDebutPerso(e.target.value)} className="rounded-2xl px-3 py-2 text-sm" style={inputStyle} />
              <span className="text-xs" style={{ color: ink.ink600 }}>→</span>
              <input type="date" value={finPerso} onChange={(e) => setFinPerso(e.target.value)} className="rounded-2xl px-3 py-2 text-sm" style={inputStyle} />
            </div>
          )}
        </div>

        <div>
          <label className="text-[11px] font-medium block mb-1.5" style={{ color: ink.ink600 }}>Filtrer par chiffre d'affaires (F)</label>
          <div className="flex items-center gap-2">
            <input type="number" value={caMin} onChange={(e) => setCaMin(e.target.value)} placeholder="Min" className="w-full rounded-2xl px-3 py-2 text-sm" style={inputStyle} />
            <span className="text-xs" style={{ color: ink.ink600 }}>—</span>
            <input type="number" value={caMax} onChange={(e) => setCaMax(e.target.value)} placeholder="Max" className="w-full rounded-2xl px-3 py-2 text-sm" style={inputStyle} />
          </div>
        </div>
      </div>

      <div className="rounded-3xl overflow-hidden" style={{ border: `1px solid ${ink.line}` }}>
        <table className="w-full text-sm table-fixed">
          <colgroup>
            <col style={{ width: "16%" }} />
            <col style={{ width: "20%" }} />
            <col className="hidden sm:table-column" style={{ width: "16%" }} />
            <col className="hidden sm:table-column" style={{ width: "14%" }} />
            <col style={{ width: "20%" }} />
            <col style={{ width: "14%" }} />
          </colgroup>
          <thead>
            <tr style={{ background: ink.canvasDeep }}>
              <th className="text-left px-3 py-3 text-xs font-semibold uppercase tracking-wide" style={{ color: ink.ink600 }}>Identifiant</th>
              <th className="text-left px-3 py-3 text-xs font-semibold uppercase tracking-wide" style={{ color: ink.ink600 }}>Nom</th>
              <th className="hidden sm:table-cell text-left px-3 py-3 text-xs font-semibold uppercase tracking-wide" style={{ color: ink.ink600 }}>Numéro</th>
              <th className="hidden sm:table-cell text-left px-3 py-3 text-xs font-semibold uppercase tracking-wide" style={{ color: ink.ink600 }}>Date d'arrivée</th>
              <th className="text-left px-3 py-3 text-xs font-semibold uppercase tracking-wide" style={{ color: ink.ink600 }}>Articles commandés</th>
              <th className="text-left px-3 py-3 text-xs font-semibold uppercase tracking-wide" style={{ color: ink.ink600 }}>Chiffre d'affaires</th>
            </tr>
          </thead>
          <tbody>
            {filtres.map((c, i) => (
              <tr
                key={c.id}
                onClick={() => onSelect(c)}
                className="cursor-pointer transition align-top"
                style={{ background: i % 2 ? ink.panel : "transparent", borderTop: `1px solid ${ink.line}` }}
              >
                <td className="px-3 py-3">
                  <Tampon id={c.id} small />
                </td>
                <td className="px-3 py-3 font-medium text-xs whitespace-normal break-words" style={{ color: ink.ink900 }}>{c.nom}</td>
                <td className="hidden sm:table-cell px-3 py-3 text-xs" style={{ color: ink.ink600 }}>{c.telephone}</td>
                <td className="hidden sm:table-cell px-3 py-3 text-xs" style={{ color: ink.ink600 }}>{c.dateEntree}</td>
                <td className="px-3 py-3 text-xs whitespace-normal break-words" style={{ color: ink.ink600 }}>{categoriesClient(c)}</td>
                <td className="px-3 py-3 font-medium text-xs" style={{ fontFamily: "'Inter', monospace", color: ink.ink900 }}>{fmt(totalClient(c))}</td>
              </tr>
            ))}
            {filtres.length === 0 && (
              <tr>
                <td colSpan={6} className="px-4 py-8 text-center text-xs" style={{ color: ink.ink600 }}>Aucun résultat pour ces filtres.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Journal({ journal }) {
  const [filtreAuteur, setFiltreAuteur] = useState("tous");
  const auteurs = [...new Set(journal.map((j) => j.auteur))];
  const filtres = filtreAuteur === "tous" ? journal : journal.filter((j) => j.auteur === filtreAuteur);
  const tries = [...filtres].sort((a, b) => (b.date + b.heure).localeCompare(a.date + a.heure));

  return (
    <div className="space-y-4">
      <div className="flex items-center gap-2">
        <span className="text-xs" style={{ color: ink.ink600 }}>Filtrer par :</span>
        <select value={filtreAuteur} onChange={(e) => setFiltreAuteur(e.target.value)} className="rounded-2xl px-3 py-1.5 text-xs" style={inputStyle}>
          <option value="tous">Tout le monde</option>
          {auteurs.map((a) => <option key={a} value={a}>{a}</option>)}
        </select>
      </div>
      <div className="rounded-3xl overflow-hidden" style={{ border: `1px solid ${ink.line}` }}>
        {tries.length === 0 ? (
          <p className="text-xs p-4" style={{ color: ink.ink600 }}>Aucune action enregistrée pour le moment.</p>
        ) : (
          tries.map((j, i) => (
            <div
              key={j.id}
              className="flex items-start gap-3 px-4 py-3 text-sm"
              style={{ background: i % 2 ? ink.panel : "transparent", borderTop: i ? `1px solid ${ink.line}` : "none" }}
            >
              <span className="text-[11px] shrink-0 w-28" style={{ fontFamily: "'Inter', monospace", color: ink.ink600 }}>
                {j.date} {j.heure}
              </span>
              <span className="text-[11px] font-semibold shrink-0 w-24 truncate" style={{ color: ink.petrol }}>
                {j.auteur}
              </span>
              <span style={{ color: ink.ink900 }}>{j.action}</span>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Messages de relance (commercial, modifiables)
// ---------------------------------------------------------------------------
function Messages({ templates, onSave }) {
  const [edits, setEdits] = useState({});
  const [savedId, setSavedId] = useState(null);

  function handleChange(id, val) {
    setEdits((e) => ({ ...e, [id]: val }));
  }
  function handleSave(tpl) {
    const texte = edits[tpl.id] !== undefined ? edits[tpl.id] : tpl.texte;
    onSave(templates.map((t) => (t.id === tpl.id ? { ...t, texte } : t)));
    setSavedId(tpl.id);
    setTimeout(() => setSavedId(null), 1500);
  }

  return (
    <div className="space-y-4">
      <p className="text-xs" style={{ color: ink.ink600 }}>
        Modifie ces modèles à ta façon. Utilise <code>{"{nom}"}</code> pour insérer automatiquement le nom du client.
      </p>
      {templates.map((tpl) => (
        <div key={tpl.id} className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>{tpl.titre}</h3>
          <textarea
            rows={3}
            value={edits[tpl.id] !== undefined ? edits[tpl.id] : tpl.texte}
            onChange={(e) => handleChange(tpl.id, e.target.value)}
            className="w-full rounded-2xl px-3 py-2 text-sm"
            style={inputStyle}
          />
          <button
            onClick={() => handleSave(tpl)}
            className="mt-2 rounded-2xl px-3 py-1.5 text-xs font-semibold"
            style={{ background: ink.petrol, color: "#fff" }}
          >
            {savedId === tpl.id ? "Enregistré ✓" : "Enregistrer ce modèle"}
          </button>
        </div>
      ))}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Mon bilan du jour (commercial)
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Vue : Bilan d'activité du pôle graphique
// ---------------------------------------------------------------------------
function BilanGraphiste({ clients, journal, currentUser }) {
  const aujourdhui = new Date().toISOString().slice(0, 10);
  const mesActions = journal.filter((j) => j.auteur === currentUser.nom && j.date === aujourdhui);
  const missionsTraitees = mesActions.filter((a) => a.action.includes("comme faite"));
  const commandesTraitees = mesActions.filter((a) => a.action.startsWith('A terminé "Conception"'));
  const clientsRestants = clients.filter((c) => c.statut === "Conception");
  const [observation, setObservation] = useState("");
  const [copie, setCopie] = useState(false);

  const detailRestantes = clientsRestants
    .map((c) => `- ${c.nom} : ${(c.commandes.map((cmd) => cmd.description).join(", ") || "besoin exprimé — " + (c.besoin || "non précisé"))}`)
    .join("\n");

  const rapport =
    `Bilan d'activité — ${currentUser.nom} — ${aujourdhui}\n\n` +
    `Missions spécifiques traitées (${missionsTraitees.length}) :\n` +
    (missionsTraitees.length ? missionsTraitees.map((a) => `- ${a.heure} : ${a.action}`).join("\n") : "- Aucune") +
    `\n\nCommandes terminées aujourd'hui (${commandesTraitees.length}) :\n` +
    (commandesTraitees.length ? commandesTraitees.map((a) => `- ${a.heure} : ${a.action}`).join("\n") : "- Aucune") +
    `\n\nCommandes restant à traiter (${clientsRestants.length}) :\n` +
    (clientsRestants.length ? detailRestantes : "- Aucune") +
    `\n\nObservation / demande particulière :\n${observation.trim() || "(aucune)"}`;

  function copier() {
    navigator.clipboard?.writeText(rapport);
    setCopie(true);
    setTimeout(() => setCopie(false), 1500);
  }

  return (
    <div className="space-y-4 pb-4">
      <div className="grid grid-cols-2 gap-3">
        <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <div className="text-[10px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>Missions traitées</div>
          <div className="text-2xl font-extrabold" style={{ color: "#5FA85B" }}>{missionsTraitees.length}</div>
        </div>
        <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <div className="text-[10px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>Commandes traitées</div>
          <div className="text-2xl font-extrabold" style={{ color: "#5FA85B" }}>{commandesTraitees.length}</div>
        </div>
      </div>

      {commandesTraitees.length > 0 && (
        <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>Détail des commandes traitées</h3>
          <div className="space-y-1.5">
            {commandesTraitees.map((a) => (
              <div key={a.id} className="text-xs" style={{ color: ink.ink600 }}>
                <span style={{ color: ink.ink300 }}>{a.heure}</span> — {a.action}
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>
          Commandes restant à traiter ({clientsRestants.length})
        </h3>
        {clientsRestants.length === 0 ? (
          <p className="text-xs" style={{ color: ink.ink600 }}>Aucune.</p>
        ) : (
          <div className="space-y-1.5">
            {clientsRestants.map((c) => (
              <div key={c.id} className="text-xs" style={{ color: ink.ink600 }}>
                <span className="font-semibold" style={{ color: ink.ink900 }}>{c.nom}</span> — {c.commandes.map((cmd) => cmd.description).join(", ") || "besoin : " + (c.besoin || "non précisé")}
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>Observation / demande particulière</h3>
        <textarea
          value={observation}
          onChange={(e) => setObservation(e.target.value)}
          rows={3}
          placeholder="Ex. besoin de plus de temps sur telle commande, matériel manquant..."
          className="w-full rounded-2xl px-3 py-2 text-sm"
          style={inputStyle}
        />
      </div>

      <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-3" style={{ color: ink.ink900 }}>Rapport à envoyer</h3>
        <pre className="text-xs whitespace-pre-wrap rounded-2xl p-3" style={{ background: "#0B0B0B", color: ink.ink900 }}>{rapport}</pre>
        <button
          onClick={copier}
          className="mt-3 flex items-center gap-1.5 rounded-2xl px-3 py-2 text-xs font-semibold"
          style={{ background: ink.petrol, color: "#fff" }}
        >
          {copie ? <Check size={13} /> : <Copy size={13} />} {copie ? "Copié !" : "Copier le rapport"}
        </button>
      </div>
    </div>
  );
}

function MonBilan({ journal, currentUser }) {
  const aujourdhui = new Date().toISOString().slice(0, 10);
  const mesActions = journal.filter((j) => j.auteur === currentUser.nom && j.date === aujourdhui);
  const [copie, setCopie] = useState(false);
  const [observation, setObservation] = useState("");

  const compte = {
    contacts: mesActions.filter((a) => a.action.startsWith("A ajouté le contact")).length,
    statuts: mesActions.filter((a) => a.action.startsWith("A changé le statut")).length,
    commandes: mesActions.filter((a) => a.action.startsWith("A ajouté une commande")).length,
    relances: mesActions.filter((a) => a.action.startsWith("A envoyé une relance")).length,
  };

  const rapport = `Rapport journalier — ${currentUser.nom} — ${aujourdhui}\n\n` +
    `• Nouveaux contacts ajoutés : ${compte.contacts}\n` +
    `• Statuts mis à jour : ${compte.statuts}\n` +
    `• Commandes enregistrées : ${compte.commandes}\n` +
    `• Relances envoyées : ${compte.relances}\n\n` +
    `Détail :\n` + mesActions.map((a) => `- ${a.heure} : ${a.action}`).join("\n") +
    `\n\nObservation / demande particulière :\n${observation.trim() || "(aucune)"}`;

  function copier() {
    navigator.clipboard?.writeText(rapport);
    setCopie(true);
    setTimeout(() => setCopie(false), 1500);
  }

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        {[
          ["Contacts ajoutés", compte.contacts],
          ["Statuts changés", compte.statuts],
          ["Commandes", compte.commandes],
          ["Relances", compte.relances],
        ].map(([label, v]) => (
          <div key={label} className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
            <div className="text-[11px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>{label}</div>
            <div className="text-xl font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: ink.ink900 }}>{v}</div>
          </div>
        ))}
      </div>

      <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>Observation / demande particulière</h3>
        <textarea
          value={observation}
          onChange={(e) => setObservation(e.target.value)}
          rows={3}
          placeholder="Ex. client difficile à convaincre, besoin de visuels supplémentaires pour argumenter..."
          className="w-full rounded-2xl px-3 py-2 text-sm"
          style={inputStyle}
        />
      </div>

      <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>Rapport à envoyer</h3>
        <pre className="text-xs whitespace-pre-wrap rounded-2xl p-3" style={{ background: ink.canvasDeep, color: ink.ink900 }}>
          {rapport}
        </pre>
        <button onClick={copier} className="mt-3 flex items-center gap-1.5 rounded-2xl px-3 py-2 text-xs font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
          {copie ? <Check size={13} /> : <Copy size={13} />} {copie ? "Copié !" : "Copier le rapport"}
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Écran de connexion — identification admin / commerciale
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Vue : File d'attente d'un pôle (graphiste / production / livraison)
// ---------------------------------------------------------------------------
function FileAttente({ clients, pole, onValider, onImportVisuel, currentUser }) {
  const statutCible = POLE_STATUT[pole];
  const items = clients.filter((c) => c.statut === statutCible);

  function handleFichier(clientId, e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => onImportVisuel(clientId, reader.result);
    reader.readAsDataURL(file);
  }

  return (
    <div className="space-y-3">
      {(pole === "production" || pole === "livraison") && currentUser && (
        <MessageDuJour nom={currentUser.nom} pole={pole} />
      )}
      {items.length === 0 ? (
        <div className="rounded-3xl p-6 text-center" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <CheckCircle2 size={28} style={{ color: ink.ink300 }} className="mx-auto mb-2" />
          <p className="text-sm" style={{ color: ink.ink600 }}>Rien en attente pour le moment.</p>
        </div>
      ) : (
        items.map((c) => {
          const deadlinePassee = c.deadlineEtape && new Date(c.deadlineEtape) < new Date();
          return (
            <div key={c.id} className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}`, borderLeft: `3px solid ${STATUT_COLOR[statutCible]}` }}>
              <div className="flex items-start justify-between mb-2">
                <div>
                  <div className="text-sm font-semibold mb-1" style={{ color: ink.ink900 }}>{c.nom}</div>
                  <Tampon id={c.id} small color={STATUT_COLOR[statutCible]} />
                </div>
              </div>
              {c.commandes.length > 0 && (
                <div className="mt-2 space-y-1">
                  {c.commandes.map((cmd) => (
                    <div key={cmd.id} className="text-xs flex items-center justify-between" style={{ color: ink.ink600 }}>
                      <span>{cmd.description}</span>
                      {pole === "livraison" && (
                        <span style={{ fontFamily: "'Inter', monospace" }}>{fmt(cmd.montant)}</span>
                      )}
                    </div>
                  ))}
                </div>
              )}
              {c.deadlineEtape && (
                <div className="text-xs mt-2 font-medium" style={{ color: deadlinePassee ? ink.rouge : ink.ink600 }}>
                  Deadline : {new Date(c.deadlineEtape).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}
                  {deadlinePassee ? " — dépassée" : ""}
                </div>
              )}

              {pole === "graphiste" && (
                <div className="mt-3 rounded-2xl p-3" style={{ background: ink.panelSoft, border: `1px dashed ${ink.line}` }}>
                  {c.visuelConception ? (
                    <div className="flex items-center gap-3">
                      <img src={c.visuelConception} alt="Visuel importé" className="h-14 w-14 rounded-xl object-cover shrink-0" />
                      <div className="min-w-0">
                        <div className="text-xs font-semibold" style={{ color: ink.ink900 }}>Visuel importé</div>
                        <label className="text-[11px] underline cursor-pointer" style={{ color: ink.bleu }}>
                          Remplacer
                          <input type="file" accept="image/jpeg,image/jpg" className="hidden" onChange={(e) => handleFichier(c.id, e)} />
                        </label>
                      </div>
                    </div>
                  ) : (
                    <label className="flex items-center gap-2 text-xs font-medium cursor-pointer" style={{ color: ink.bleu }}>
                      <Plus size={14} />
                      Importer le visuel final (JPEG)
                      <input type="file" accept="image/jpeg,image/jpg" className="hidden" onChange={(e) => handleFichier(c.id, e)} />
                    </label>
                  )}
                </div>
              )}

              {pole === "production" && c.visuelConception && (
                <div className="mt-3 rounded-2xl overflow-hidden" style={{ border: `1px solid ${ink.line}` }}>
                  <div className="px-3 py-1.5 text-[10px] font-bold uppercase tracking-wide" style={{ background: ink.panelSoft, color: ink.bleu }}>
                    Maquette du graphiste
                  </div>
                  <img src={c.visuelConception} alt="Maquette finale" className="w-full max-h-64 object-contain" style={{ background: "#0B0B0B" }} />
                </div>
              )}

              <button
                onClick={() => onValider(c)}
                className="w-full mt-3 flex items-center justify-center gap-1.5 rounded-2xl py-2.5 text-sm font-semibold"
                style={{ background: ink.petrol, color: "#fff" }}
              >
                <CheckCircle2 size={15} /> FAIT — passer à l'étape suivante
              </button>
            </div>
          );
        })
      )}
    </div>
  );
}


function BottomNav({ view, setView, isAdmin, hasRole, onAjouterCommande }) {
  let homeId = "dashboard";
  if (!isAdmin && !hasRole("commercial")) {
    if (hasRole("graphiste")) homeId = "dashboard_graphiste";
    else if (hasRole("production")) homeId = "dashboard_production";
    else if (hasRole("livraison")) homeId = "file_livraison";
    else homeId = "missions";
  }
  const commandeEnabled = isAdmin || hasRole("commercial");
  const commandeId = commandeEnabled
    ? "clients"
    : hasRole("graphiste")
    ? "file_graphiste"
    : hasRole("production")
    ? "file_production"
    : homeId;

  const items = [
    { id: homeId, label: "Accueil", Icon: LayoutGrid, action: () => setView(homeId) },
    {
      id: commandeId,
      label: "Commande",
      Icon: Users,
      action: () => setView(commandeId),
    },
    isAdmin
      ? { id: "base_clients", label: "Client", Icon: Database, center: true, action: () => setView("base_clients") }
      : {
          id: "create",
          label: commandeEnabled ? "Enregistrer" : "Mission",
          Icon: Plus,
          center: true,
          action: commandeEnabled ? onAjouterCommande : () => setView("missions"),
        },
    { id: "missions", label: "Missions", Icon: ListChecks, action: () => setView("missions") },
    {
      id: isAdmin ? "bilan" : hasRole("commercial") ? "monbilan" : "missions",
      label: isAdmin || hasRole("commercial") ? "Bilan" : "Plus",
      Icon: isAdmin || hasRole("commercial") ? FileBarChart : ClipboardList,
      action: () => setView(isAdmin ? "bilan" : hasRole("commercial") ? "monbilan" : "missions"),
    },
  ];

  return (
    <div className="fixed bottom-0 left-0 right-0 z-40 px-2 pb-4 pt-2.5 sm:hidden" style={{ background: "rgba(10,10,10,0.94)", backdropFilter: "blur(20px)", borderTop: `1px solid ${ink.line}` }}>
      <div className="grid grid-cols-5 items-end max-w-md mx-auto">
        {items.map(({ id, label, Icon, center, action }) => {
          const active = view === id;
          if (center) {
            return (
              <button key={id} onClick={action} className="flex flex-col items-center gap-0.5 -mt-7">
                <div className="rounded-full flex items-center justify-center shadow-xl orange-gradient" style={{ height: 52, width: 52 }}>
                  <Icon size={22} color="#fff" />
                </div>
                <span className="text-[9px] font-medium truncate max-w-full" style={{ color: ink.orange }}>{label}</span>
              </button>
            );
          }
          return (
            <button key={id + label} onClick={action} className="flex flex-col items-center gap-0.5 px-0.5 min-w-0">
              <Icon size={18} style={{ color: active ? ink.orange : ink.ink300 }} />
              <span className="text-[9px] font-medium truncate max-w-full" style={{ color: active ? ink.orange : ink.ink300 }}>{label}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

      * {
        font-family: 'Inter', sans-serif;
        -webkit-tap-highlight-color: transparent;
        box-sizing: border-box;
      }

      html, body { background: #070707; overflow-x: hidden; max-width: 100%; }
      #root, .app-shell { overflow-x: hidden; max-width: 100vw; }
      img, svg { max-width: 100%; }

      button { transition: transform 0.12s ease, opacity 0.12s ease, background-color 0.2s ease; }
      button:active { transform: scale(0.96); opacity: 0.9; }

      input, textarea, select { outline: none; transition: border-color 0.15s ease; }
      input::placeholder, textarea::placeholder { color: #7C7C7C; }

      .app-card {
        background: linear-gradient(145deg, #171717, #101010);
        border: 1px solid rgba(255,255,255,0.08);
        box-shadow: 0 12px 35px rgba(0,0,0,0.35);
      }

      .orange-gradient { background: linear-gradient(135deg, #FF7A00, #FF4D00); }

      ::-webkit-scrollbar { width: 10px; height: 10px; }
      ::-webkit-scrollbar-thumb { background: #3A3A3A; border-radius: 6px; }
      ::-webkit-scrollbar-track { background: transparent; }
    `}</style>
  );
}

// ---------------------------------------------------------------------------
// Vue : Tableau de bord du pôle graphique
// ---------------------------------------------------------------------------
function GraphisteDashboard({ clients, journal, currentUser, setView }) {
  const enAttente = clients.filter((c) => c.statut === "Conception");
  const aujourdhui = new Date().toISOString().slice(0, 10);
  const traiteesAujourdhui = journal.filter(
    (j) => j.auteur === currentUser.nom && j.date === aujourdhui && j.action.startsWith('A terminé "Conception"')
  ).length;

  return (
    <div className="space-y-5 pb-4">
      <MessageDuJour nom={currentUser.nom} pole="graphiste" />

      <div>
        <h1 className="text-2xl sm:text-3xl font-extrabold mb-1" style={{ color: ink.ink900 }}>
          Salut {currentUser.nom} 👋
        </h1>
        <p className="text-sm" style={{ color: ink.ink600 }}>Voici tes conceptions du jour.</p>
      </div>

      <button
        onClick={() => setView("file_graphiste")}
        className="w-full rounded-3xl p-5 text-left relative overflow-hidden"
        style={{ background: "linear-gradient(135deg, #151515 0%, #090909 100%)", border: `1px solid ${ink.line}` }}
      >
        <div className="absolute -right-10 -top-10 h-32 w-32 rounded-full blur-2xl" style={{ background: "rgba(79,168,216,0.3)" }} />
        <div className="relative z-10 flex items-center justify-between">
          <div>
            <div className="text-sm mb-1" style={{ color: ink.ink600 }}>Commande</div>
            <div className="text-3xl font-extrabold" style={{ color: ink.ink900 }}>{enAttente.length} en attente</div>
          </div>
          <ChevronRight size={22} style={{ color: ink.bleu }} />
        </div>
      </button>

      <div className="grid grid-cols-2 gap-3">
        <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <div className="text-[10px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>Traitées aujourd'hui</div>
          <div className="text-2xl font-extrabold" style={{ color: "#5FA85B" }}>{traiteesAujourdhui}</div>
        </div>
        <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <div className="text-[10px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>Restantes</div>
          <div className="text-2xl font-extrabold" style={{ color: ink.rouge }}>{enAttente.length}</div>
        </div>
      </div>

      <button
        onClick={() => setView("monbilan")}
        className="w-full flex items-center justify-between rounded-2xl p-4"
        style={{ background: ink.panel, border: `1px solid ${ink.line}` }}
      >
        <div className="flex items-center gap-2.5">
          <ClipboardList size={18} style={{ color: ink.bleu }} />
          <span className="text-sm font-semibold" style={{ color: ink.ink900 }}>Mon bilan d'activité</span>
        </div>
        <ChevronRight size={16} style={{ color: ink.ink300 }} />
      </button>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Mini-composant : signaler un besoin particulier (persisté via le journal)
// ---------------------------------------------------------------------------
function BesoinsParticuliers({ journal, currentUser, onAjouter }) {
  const [texte, setTexte] = useState("");
  const aujourdhui = new Date().toISOString().slice(0, 10);
  const mesBesoins = journal.filter(
    (j) => j.auteur === currentUser.nom && j.date === aujourdhui && j.action.startsWith("A signalé un besoin particulier")
  );

  function envoyer() {
    if (!texte.trim()) return;
    onAjouter(texte);
    setTexte("");
  }

  return (
    <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
      <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>Besoins particuliers / observations</h3>
      <div className="flex gap-2 mb-2">
        <input
          value={texte}
          onChange={(e) => setTexte(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && envoyer()}
          placeholder="Ex. rupture de stock de tissu, machine à réviser..."
          className="flex-1 rounded-2xl px-3 py-2 text-sm"
          style={inputStyle}
        />
        <button onClick={envoyer} className="shrink-0 rounded-2xl px-3 py-2 text-xs font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
          Signaler
        </button>
      </div>
      {mesBesoins.length > 0 && (
        <div className="space-y-1.5 mt-2">
          {mesBesoins.map((b) => (
            <div key={b.id} className="text-xs" style={{ color: ink.ink600 }}>
              <span style={{ color: ink.ink300 }}>{b.heure}</span> — {b.action.replace('A signalé un besoin particulier : ', '')}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Tableau de bord de la production
// ---------------------------------------------------------------------------
function ProductionDashboard({ clients, journal, currentUser, setView, onAjouterObservation }) {
  const enAttente = clients.filter((c) => c.statut === "En production");
  const aujourdhui = new Date().toISOString().slice(0, 10);
  const traiteesAujourdhui = journal.filter(
    (j) => j.auteur === currentUser.nom && j.date === aujourdhui && j.action.startsWith('A terminé "En production"')
  ).length;

  return (
    <div className="space-y-5 pb-4">
      <MessageDuJour nom={currentUser.nom} pole="production" />

      <div>
        <h1 className="text-2xl sm:text-3xl font-extrabold mb-1" style={{ color: ink.ink900 }}>
          Salut {currentUser.nom} 👋
        </h1>
        <p className="text-sm" style={{ color: ink.ink600 }}>Voici ta production du jour.</p>
      </div>

      <button
        onClick={() => setView("file_production")}
        className="w-full rounded-3xl p-5 text-left relative overflow-hidden"
        style={{ background: "linear-gradient(135deg, #151515 0%, #090909 100%)", border: `1px solid ${ink.line}` }}
      >
        <div className="absolute -right-10 -top-10 h-32 w-32 rounded-full blur-2xl" style={{ background: "rgba(255,106,0,0.28)" }} />
        <div className="relative z-10 flex items-center justify-between">
          <div>
            <div className="text-sm mb-1" style={{ color: ink.ink600 }}>Commande</div>
            <div className="text-3xl font-extrabold" style={{ color: ink.ink900 }}>{enAttente.length} en attente</div>
          </div>
          <ChevronRight size={22} style={{ color: ink.petrol }} />
        </div>
      </button>

      <div className="grid grid-cols-2 gap-3">
        <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <div className="text-[10px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>Traitées aujourd'hui</div>
          <div className="text-2xl font-extrabold" style={{ color: "#5FA85B" }}>{traiteesAujourdhui}</div>
        </div>
        <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <div className="text-[10px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>Restantes</div>
          <div className="text-2xl font-extrabold" style={{ color: ink.rouge }}>{enAttente.length}</div>
        </div>
      </div>

      <BesoinsParticuliers journal={journal} currentUser={currentUser} onAjouter={onAjouterObservation} />

      <button
        onClick={() => setView("bilan_production")}
        className="w-full flex items-center justify-between rounded-2xl p-4"
        style={{ background: ink.panel, border: `1px solid ${ink.line}` }}
      >
        <div className="flex items-center gap-2.5">
          <ClipboardList size={18} style={{ color: ink.petrol }} />
          <span className="text-sm font-semibold" style={{ color: ink.ink900 }}>Rapport d'activité</span>
        </div>
        <ChevronRight size={16} style={{ color: ink.ink300 }} />
      </button>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Vue : Rapport d'activité de la production
// ---------------------------------------------------------------------------
function BilanProduction({ clients, journal, currentUser, onAjouterObservation }) {
  const aujourdhui = new Date().toISOString().slice(0, 10);
  const mesActions = journal.filter((j) => j.auteur === currentUser.nom && j.date === aujourdhui);
  const commandesTraitees = mesActions.filter((a) => a.action.startsWith('A terminé "En production"'));
  const besoins = mesActions.filter((a) => a.action.startsWith("A signalé un besoin particulier"));
  const clientsRestants = clients.filter((c) => c.statut === "En production");
  const [copie, setCopie] = useState(false);

  const detailRestantes = clientsRestants
    .map((c) => `- ${c.nom} : ${c.commandes.map((cmd) => cmd.description).join(", ") || "non précisé"}`)
    .join("\n");

  const rapport =
    `Rapport d'activité — Production — ${currentUser.nom} — ${aujourdhui}\n\n` +
    `Commandes traitées aujourd'hui (${commandesTraitees.length}) :\n` +
    (commandesTraitees.length ? commandesTraitees.map((a) => `- ${a.heure} : ${a.action}`).join("\n") : "- Aucune") +
    `\n\nCommandes restant à traiter (${clientsRestants.length}) :\n` +
    (clientsRestants.length ? detailRestantes : "- Aucune") +
    `\n\nBesoins particuliers signalés (${besoins.length}) :\n` +
    (besoins.length ? besoins.map((a) => `- ${a.heure} : ${a.action.replace('A signalé un besoin particulier : ', '')}`).join("\n") : "- Aucun");

  function copier() {
    navigator.clipboard?.writeText(rapport);
    setCopie(true);
    setTimeout(() => setCopie(false), 1500);
  }

  return (
    <div className="space-y-4 pb-4">
      <div className="grid grid-cols-2 gap-3">
        <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <div className="text-[10px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>Commandes traitées</div>
          <div className="text-2xl font-extrabold" style={{ color: "#5FA85B" }}>{commandesTraitees.length}</div>
        </div>
        <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <div className="text-[10px] uppercase tracking-wide mb-1" style={{ color: ink.ink600 }}>Commandes restantes</div>
          <div className="text-2xl font-extrabold" style={{ color: ink.rouge }}>{clientsRestants.length}</div>
        </div>
      </div>

      {commandesTraitees.length > 0 && (
        <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
          <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>Détail des commandes traitées</h3>
          <div className="space-y-1.5">
            {commandesTraitees.map((a) => (
              <div key={a.id} className="text-xs" style={{ color: ink.ink600 }}>
                <span style={{ color: ink.ink300 }}>{a.heure}</span> — {a.action}
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-2" style={{ color: ink.ink900 }}>
          Commandes restant à traiter ({clientsRestants.length})
        </h3>
        {clientsRestants.length === 0 ? (
          <p className="text-xs" style={{ color: ink.ink600 }}>Aucune.</p>
        ) : (
          <div className="space-y-1.5">
            {clientsRestants.map((c) => (
              <div key={c.id} className="text-xs" style={{ color: ink.ink600 }}>
                <span className="font-semibold" style={{ color: ink.ink900 }}>{c.nom}</span> — {c.commandes.map((cmd) => cmd.description).join(", ") || "non précisé"}
              </div>
            ))}
          </div>
        )}
      </div>

      <BesoinsParticuliers journal={journal} currentUser={currentUser} onAjouter={onAjouterObservation} />

      <div className="rounded-3xl p-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
        <h3 className="text-sm font-semibold mb-3" style={{ color: ink.ink900 }}>Rapport à envoyer</h3>
        <pre className="text-xs whitespace-pre-wrap rounded-2xl p-3" style={{ background: ink.canvasDeep, color: ink.ink900 }}>{rapport}</pre>
        <button onClick={copier} className="mt-3 flex items-center gap-1.5 rounded-2xl px-3 py-2 text-xs font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
          {copie ? <Check size={13} /> : <Copy size={13} />} {copie ? "Copié !" : "Copier le rapport"}
        </button>
      </div>
    </div>
  );
}

function LoginScreen({ onLogin, adminAuth, utilisateurs }) {
  const [etape, setEtape] = useState("choix");
  const [nom, setNom] = useState("");
  const [motDePasse, setMotDePasse] = useState("");
  const [erreur, setErreur] = useState("");

  function tenterAdmin() {
    if (nom.trim().toLowerCase() === adminAuth.nom.trim().toLowerCase() && motDePasse === adminAuth.motDePasse) {
      onLogin({ roles: ["admin"], nom: adminAuth.nom });
    } else {
      setErreur("Nom ou mot de passe incorrect.");
    }
  }

  function tenterUtilisateur() {
    const trouve = utilisateurs.find(
      (u) => u.nom.trim().toLowerCase() === nom.trim().toLowerCase() && u.motDePasse === motDePasse
    );
    if (trouve) {
      onLogin({ roles: trouve.roles && trouve.roles.length ? trouve.roles : ["commercial"], nom: trouve.nom });
    } else {
      setErreur("Nom ou mot de passe incorrect.");
    }
  }

  function retour() {
    setEtape("choix");
    setNom("");
    setMotDePasse("");
    setErreur("");
  }

  return (
    <div className="min-h-screen w-full flex items-center justify-center p-5" style={{ background: ink.canvas }}>
      <GlobalStyle />
      <div className="w-full max-w-sm rounded-[32px] p-6 app-card">
        <div className="flex flex-col items-center text-center mb-1">
          <LogoMark size={96} />
          <div
            className="text-xl font-bold"
            style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: ink.ink900 }}
          >
            LE SÉRIGRAPHE
          </div>
        </div>
        <p className="text-xs mb-6 text-center" style={{ color: ink.ink600 }}>Identifie-toi pour accéder au suivi clients.</p>

        {etape === "choix" && (
          <div className="space-y-2.5">
            <button
              onClick={() => { setEtape("admin"); setErreur(""); }}
              className="w-full flex items-center gap-3 rounded-2xl px-4 py-3 text-left"
              style={{ background: ink.petrol, color: "#fff" }}
            >
              <ShieldCheck size={18} />
              <div>
                <div className="text-sm font-semibold">Administrateur</div>
                <div className="text-[11px] opacity-80">Accès complet</div>
              </div>
            </button>
            <button
              onClick={() => { setEtape("utilisateur"); setErreur(""); }}
              className="w-full flex items-center gap-3 rounded-2xl px-4 py-3 text-left"
              style={{ background: ink.canvasDeep, color: ink.ink900 }}
            >
              <UserCog size={18} />
              <div>
                <div className="text-sm font-semibold">Utilisateur</div>
                <div className="text-[11px]" style={{ color: ink.ink600 }}>Commercial, pôle graphique, production, livraison</div>
              </div>
            </button>
          </div>
        )}

        {(etape === "admin" || etape === "utilisateur") && (
          <div className="space-y-3">
            <div>
              <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Nom</label>
              <input
                value={nom}
                onChange={(e) => setNom(e.target.value)}
                placeholder={etape === "admin" ? "Félix" : "Ex. Gloria"}
                className="w-full rounded-2xl px-3 py-2 text-sm mt-1"
                style={inputStyle}
                autoFocus
              />
            </div>
            <div>
              <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Mot de passe</label>
              <input
                type="password"
                value={motDePasse}
                onChange={(e) => setMotDePasse(e.target.value)}
                className="w-full rounded-2xl px-3 py-2 text-sm mt-1"
                style={inputStyle}
                onKeyDown={(e) => e.key === "Enter" && (etape === "admin" ? tenterAdmin() : tenterUtilisateur())}
              />
            </div>
            {erreur && <p className="text-xs" style={{ color: ink.rouge }}>{erreur}</p>}
            <button
              onClick={etape === "admin" ? tenterAdmin : tenterUtilisateur}
              disabled={!nom.trim() || !motDePasse}
              className="w-full rounded-2xl py-2.5 text-sm font-semibold"
              style={{ background: nom.trim() && motDePasse ? ink.petrol : ink.ink300, color: "#fff" }}
            >
              Entrer
            </button>
            <button onClick={retour} className="text-xs" style={{ color: ink.ink600 }}>
              ← Retour
            </button>
          </div>
        )}
      </div>
    </div>
  );
}


function SelecteurCategorie({ value, onChange, categories, onAddCategorie }) {
  const [ouvert, setOuvert] = useState(false);
  const [nouvelle, setNouvelle] = useState("");

  function valider() {
    const val = nouvelle.trim();
    if (!val) return;
    onAddCategorie(val);
    onChange(val);
    setNouvelle("");
    setOuvert(false);
  }

  return (
    <div className="mb-2">
      <div className="flex gap-2">
        <select value={value} onChange={(e) => onChange(e.target.value)} className="flex-1 rounded-2xl px-3 py-2 text-sm" style={inputStyle}>
          <option value="">Catégorie — non précisée</option>
          {categories.map((cat) => (
            <option key={cat} value={cat}>{cat}</option>
          ))}
        </select>
        <button
          onClick={() => setOuvert((o) => !o)}
          className="shrink-0 rounded-2xl px-3 py-2 text-xs font-semibold flex items-center gap-1"
          style={{ background: ouvert ? ink.panel : ink.bleu, color: ouvert ? ink.ink600 : "#fff" }}
        >
          <Plus size={13} /> Article
        </button>
      </div>
      {ouvert && (
        <div className="flex gap-2 mt-2">
          <input
            value={nouvelle}
            onChange={(e) => setNouvelle(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && valider()}
            placeholder="Ex. Mug"
            className="flex-1 rounded-2xl px-3 py-2 text-sm"
            style={inputStyle}
            autoFocus
          />
          <button onClick={valider} className="shrink-0 rounded-2xl px-3 py-2 text-xs font-semibold" style={{ background: ink.orange, color: "#fff" }}>
            Ajouter
          </button>
        </div>
      )}
    </div>
  );
}

function CommandeCard({ cmd, clientId, onAddCout, onSolder, onEdit, onDelete, categories, onAddCategorie }) {
  const [ouvert, setOuvert] = useState(false);
  const [descCout, setDescCout] = useState("");
  const [montantCout, setMontantCout] = useState("");
  const [editOuvert, setEditOuvert] = useState(false);
  const [editDesc, setEditDesc] = useState(cmd.description);
  const [editMontant, setEditMontant] = useState(String(cmd.montant || ""));
  const [editMontantPaye, setEditMontantPaye] = useState(String(cmd.montantPaye || ""));
  const [editCategorie, setEditCategorie] = useState(cmd.categorie || "");
  const [editNonConfirmee, setEditNonConfirmee] = useState(!!cmd.nonConfirmee);
  const [confirmSupp, setConfirmSupp] = useState(false);
  const statutP = statutPaiement(cmd);
  const reste = cmd.montant - cmd.montantPaye;
  const couts = cmd.couts || [];
  const marge = margeCommande(cmd);

  function ajouter() {
    if (!descCout.trim() || !montantCout) return;
    onAddCout(clientId, cmd.id, { id: Date.now(), description: descCout.trim(), montant: Number(montantCout) });
    setDescCout("");
    setMontantCout("");
    setOuvert(false);
  }

  function sauverEdit() {
    if (!editDesc.trim()) return;
    if (!editNonConfirmee && !editMontant) return;
    onEdit(clientId, cmd.id, {
      description: editDesc.trim(),
      montant: Number(editMontant || 0),
      montantPaye: Number(editMontantPaye || 0),
      categorie: editCategorie || null,
      nonConfirmee: editNonConfirmee,
    });
    setEditOuvert(false);
  }

  function supprimer() {
    onDelete(clientId, cmd.id);
  }


  if (editOuvert) {
    return (
      <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.bleu}` }}>
        <input value={editDesc} onChange={(e) => setEditDesc(e.target.value)} placeholder="Description" className="w-full rounded-2xl px-3 py-2 text-sm mb-2" style={inputStyle} />
        <SelecteurCategorie value={editCategorie} onChange={setEditCategorie} categories={categories} onAddCategorie={onAddCategorie} />
        <div className="grid grid-cols-2 gap-2 mb-2">
          <input value={editMontant} onChange={(e) => setEditMontant(e.target.value)} type="number" placeholder="Montant (F)" disabled={editNonConfirmee} className="w-full rounded-2xl px-3 py-2 text-sm" style={{ ...inputStyle, opacity: editNonConfirmee ? 0.5 : 1 }} />
          <input value={editMontantPaye} onChange={(e) => setEditMontantPaye(e.target.value)} type="number" placeholder="Déjà payé (F)" disabled={editNonConfirmee} className="w-full rounded-2xl px-3 py-2 text-sm" style={{ ...inputStyle, opacity: editNonConfirmee ? 0.5 : 1 }} />
        </div>
        <label className="flex items-center gap-2 mb-2 text-xs font-medium cursor-pointer" style={{ color: ink.ochre }}>
          <input type="checkbox" checked={editNonConfirmee} onChange={(e) => setEditNonConfirmee(e.target.checked)} />
          Commande non confirmée (pas encore de montant)
        </label>
        <div className="flex gap-2">
          <button onClick={sauverEdit} className="flex-1 rounded-2xl py-2 text-xs font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
            Enregistrer
          </button>
          <button onClick={() => setEditOuvert(false)} className="rounded-2xl py-2 px-3 text-xs font-semibold" style={{ background: ink.canvasDeep, color: ink.ink600 }}>
            Annuler
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
      <div className="flex items-start justify-between mb-1.5">
        <span className="text-sm font-medium" style={{ color: ink.ink900 }}>{cmd.description}</span>
        <span
          className="text-[10px] font-semibold px-1.5 py-0.5 rounded shrink-0 ml-2"
          style={{ background: cmd.nonConfirmee ? `${ink.ochre}1A` : `${PAIEMENT_COLOR[statutP]}1A`, color: cmd.nonConfirmee ? ink.ochre : PAIEMENT_COLOR[statutP] }}
        >
          {cmd.nonConfirmee ? "Non confirmée" : statutP}
        </span>
      </div>
      <div className="text-[11px] mb-2 flex items-center gap-1.5" style={{ color: ink.ink600 }}>
        {cmd.date}
        {cmd.categorie && (
          <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded" style={{ background: `${ink.bleu}1A`, color: ink.bleu }}>
            {cmd.categorie}
          </span>
        )}
      </div>
      {!cmd.nonConfirmee && (
        <div className="flex items-center justify-between text-xs mb-2" style={{ fontFamily: "'Inter', monospace" }}>
          <span style={{ color: ink.ink900 }}>{fmt(cmd.montant)}</span>
          {reste > 0 && <span style={{ color: ink.rouge }}>reste {fmt(reste)}</span>}
        </div>
      )}

      {couts.length > 0 && (
        <div className="rounded-2xl p-2 mb-2 space-y-1" style={{ background: ink.canvasDeep }}>
          {couts.map((c) => (
            <div key={c.id} className="flex items-center justify-between text-[11px]" style={{ color: ink.ink600 }}>
              <span>{c.description}</span>
              <span style={{ fontFamily: "'Inter', monospace" }}>{fmt(c.montant)}</span>
            </div>
          ))}
        </div>
      )}

      <div className="flex items-center justify-between gap-2 mb-2">
        <div className="text-xs font-semibold" style={{ color: marge >= 0 ? "#5FA85B" : ink.rouge }}>
          Marge brute : {fmt(marge)}
        </div>
      </div>

      <div className="flex items-center gap-1.5 flex-wrap">
        {reste > 0 && (
          <button
            onClick={() => onSolder(clientId, cmd.id)}
            className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl"
            style={{ background: "#5FA85B", color: "#fff" }}
          >
            ✓ Marquer soldé
          </button>
        )}
        <button
          onClick={() => setOuvert((o) => !o)}
          className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl"
          style={{ background: ouvert ? ink.canvasDeep : ink.petrol, color: ouvert ? ink.ink600 : "#fff" }}
        >
          {ouvert ? "Annuler" : "+ Ajouter un coût"}
        </button>
        <button
          onClick={() => setEditOuvert(true)}
          className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl"
          style={{ background: ink.canvasDeep, color: ink.ink600 }}
        >
          Modifier
        </button>
        {!confirmSupp ? (
          <button
            onClick={() => setConfirmSupp(true)}
            className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl flex items-center gap-1"
            style={{ background: "transparent", color: ink.rouge, border: `1px solid ${ink.rouge}` }}
          >
            <Trash2 size={11} /> Supprimer
          </button>
        ) : (
          <>
            <button
              onClick={supprimer}
              className="text-[11px] font-bold px-2.5 py-1.5 rounded-2xl"
              style={{ background: ink.rouge, color: "#fff" }}
            >
              Confirmer la suppression
            </button>
            <button
              onClick={() => setConfirmSupp(false)}
              className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl"
              style={{ background: ink.canvasDeep, color: ink.ink600 }}
            >
              Annuler
            </button>
          </>
        )}
      </div>

      {ouvert && (
        <div className="mt-2 space-y-1.5">
          <input
            value={descCout}
            onChange={(e) => setDescCout(e.target.value)}
            placeholder="Ex. Achat t-shirts"
            className="w-full rounded-2xl px-3 py-1.5 text-xs"
            style={inputStyle}
          />
          <div className="flex gap-1.5">
            <input
              value={montantCout}
              onChange={(e) => setMontantCout(e.target.value)}
              type="number"
              placeholder="Montant (F)"
              className="flex-1 rounded-2xl px-3 py-1.5 text-xs"
              style={inputStyle}
            />
            <button onClick={ajouter} className="rounded-2xl px-3 py-1.5 text-xs font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
              Ajouter
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function FicheClient({ client, onClose, onChangeStatut, onAddCommande, onAddCout, onSolderCommande, onEditCommande, onDeleteCommande, templates, onRelance, categories, onAddCategorie, onSetDeadlineEtape, onEditClient, onDeleteClient, modeGestion = false }) {
  const [ajoutOuvert, setAjoutOuvert] = useState(false);
  const [desc, setDesc] = useState("");
  const [montant, setMontant] = useState("");
  const [montantPaye, setMontantPaye] = useState("");
  const [categorieChoisie, setCategorieChoisie] = useState("");
  const [nonConfirmee, setNonConfirmee] = useState(false);
  const [relanceOuverte, setRelanceOuverte] = useState(false);
  const [copieId, setCopieId] = useState(null);
  const [editClientOuvert, setEditClientOuvert] = useState(false);
  const [editNom, setEditNom] = useState(client.nom);
  const [editTelephone, setEditTelephone] = useState(client.telephone || "");
  const [editSource, setEditSource] = useState(client.source || "Facebook");
  const [editBesoin, setEditBesoin] = useState(client.besoin || "");
  const [editType, setEditType] = useState(client.type);
  const [confirmSuppClient, setConfirmSuppClient] = useState(false);

  function sauverEditClient() {
    if (!editNom.trim()) return;
    onEditClient(client.id, {
      nom: editNom.trim(),
      telephone: editTelephone.trim(),
      source: editSource,
      besoin: editBesoin.trim(),
      type: editType,
    });
    setEditClientOuvert(false);
  }

  function supprimerClient() {
    onDeleteClient(client.id);
    onClose();
  }

  function handleAddCommande() {
    if (!desc.trim()) return;
    if (!nonConfirmee && !montant) return;
    onAddCommande(client.id, {
      id: Date.now(),
      date: AUJOURD_HUI.toISOString().slice(0, 10),
      description: desc.trim(),
      montant: Number(montant || 0),
      montantPaye: Number(montantPaye || 0),
      categorie: categorieChoisie || null,
      nonConfirmee,
      couts: [],
    });
    setDesc("");
    setMontant("");
    setMontantPaye("");
    setCategorieChoisie("");
    setNonConfirmee(false);
    setAjoutOuvert(false);
  }

  function copierMessage(tpl) {
    const texte = tpl.texte.replace("{nom}", client.nom);
    navigator.clipboard?.writeText(texte);
    setCopieId(tpl.id);
    onRelance(client, tpl.titre);
    setTimeout(() => setCopieId(null), 1500);
  }

  return (
    <div className="w-full max-w-md mx-auto sm:max-w-2xl">
      <button
        onClick={onClose}
        className="flex items-center gap-1.5 text-sm font-semibold mb-4"
        style={{ color: ink.ink600 }}
      >
        <ChevronLeft size={18} /> Retour
      </button>
      <div className="pb-10">
        <div className="flex items-center gap-2 flex-wrap mb-4">
          <Tampon id={client.id} small />
          <span className="text-base font-bold" style={{ color: ink.ink900 }}>{client.nom}</span>
          {client.telephone && (
            <span className="text-xs" style={{ color: ink.ink600 }}>· {client.telephone}</span>
          )}
          <TypeBadge type={client.type} />
        </div>

        <div className="grid grid-cols-2 gap-2 mb-4">
          <div>
            <label className="text-[10px] font-medium uppercase tracking-wide block mb-1" style={{ color: ink.ink600 }}>
              Statut
            </label>
            <div className="relative">
              <select
                value={client.statut}
                onChange={(e) => onChangeStatut(client.id, e.target.value)}
                className="w-full text-xs font-semibold rounded-2xl pl-2.5 pr-6 py-2 outline-none appearance-none"
                style={{
                  background: ink.panel,
                  border: `2px solid ${STATUT_COLOR[client.statut]}`,
                  color: STATUT_COLOR[client.statut],
                }}
              >
                {STATUTS.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
              <ChevronRight
                size={12}
                className="absolute right-2 top-1/2 -translate-y-1/2 pointer-events-none rotate-90"
                style={{ color: STATUT_COLOR[client.statut] }}
              />
            </div>
          </div>
          <div>
            <label className="text-[10px] font-medium uppercase tracking-wide block mb-1" style={{ color: ink.ink600 }}>
              Deadline étape
            </label>
            <input
              type="datetime-local"
              value={client.deadlineEtape ? client.deadlineEtape.slice(0, 16) : ""}
              onChange={(e) => onSetDeadlineEtape(client.id, e.target.value || null)}
              className="w-full text-xs rounded-2xl px-2.5 py-2"
              style={inputStyle}
            />
          </div>
        </div>

        {modeGestion && (
          <>
            <div className="flex gap-2 mb-4">
              <button
                onClick={() => setEditClientOuvert((o) => !o)}
                className="flex-1 flex items-center justify-center gap-1.5 rounded-2xl py-2.5 text-xs font-semibold"
                style={{ background: editClientOuvert ? ink.panel : ink.bleu, color: editClientOuvert ? ink.ink600 : "#fff" }}
              >
                {editClientOuvert ? "Annuler" : "Modifier les infos"}
              </button>
              {!confirmSuppClient ? (
                <button
                  onClick={() => setConfirmSuppClient(true)}
                  className="flex-1 flex items-center justify-center gap-1.5 rounded-2xl py-2.5 text-xs font-semibold"
                  style={{ background: "transparent", color: ink.rouge, border: `1px solid ${ink.rouge}` }}
                >
                  <Trash2 size={13} /> Supprimer le client
                </button>
              ) : (
                <div className="flex-1 flex gap-2">
                  <button
                    onClick={supprimerClient}
                    className="flex-1 rounded-2xl py-2.5 text-xs font-bold"
                    style={{ background: ink.rouge, color: "#fff" }}
                  >
                    Confirmer
                  </button>
                  <button
                    onClick={() => setConfirmSuppClient(false)}
                    className="rounded-2xl px-3 py-2.5 text-xs font-semibold"
                    style={{ background: ink.panel, color: ink.ink600 }}
                  >
                    Annuler
                  </button>
                </div>
              )}
            </div>

            {confirmSuppClient && (
              <div className="rounded-2xl p-3 mb-4 text-xs" style={{ background: `${ink.rouge}1A`, color: ink.rouge }}>
                Cette action supprimera définitivement {client.nom} et toutes ses commandes. Clique sur "Confirmer" pour valider.
              </div>
            )}

            {editClientOuvert && (
              <div className="rounded-2xl p-4 mb-4 space-y-2.5" style={{ background: ink.panel, border: `1px solid ${ink.bleu}` }}>
                <div>
                  <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Nom</label>
                  <input value={editNom} onChange={(e) => setEditNom(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle} />
                </div>
                <div>
                  <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Téléphone / WhatsApp</label>
                  <input value={editTelephone} onChange={(e) => setEditTelephone(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle} />
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Source</label>
                    <select value={editSource} onChange={(e) => setEditSource(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle}>
                      <option>Facebook</option>
                      <option>WhatsApp</option>
                      <option>TikTok</option>
                      <option>Autre</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Type</label>
                    <select value={editType} onChange={(e) => setEditType(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle}>
                      <option value="client">Client</option>
                      <option value="prestataire">Prestataire</option>
                    </select>
                  </div>
                </div>
                <div>
                  <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Besoin exprimé</label>
                  <textarea value={editBesoin} onChange={(e) => setEditBesoin(e.target.value)} rows={2} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle} />
                </div>
                <button onClick={sauverEditClient} className="w-full rounded-2xl py-2.5 text-sm font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
                  Enregistrer
                </button>
              </div>
            )}
          </>
        )}

        {ajoutOuvert ? (
          <div className="rounded-2xl p-3 mb-4" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
            <input value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Description (ex. T-shirts x50)" className="w-full rounded-2xl px-3 py-2 text-sm mb-2" style={inputStyle} />
            <SelecteurCategorie value={categorieChoisie} onChange={setCategorieChoisie} categories={categories} onAddCategorie={onAddCategorie} />
            <div className="grid grid-cols-2 gap-2 mb-2">
              <input value={montant} onChange={(e) => setMontant(e.target.value)} type="number" placeholder="Montant (F)" disabled={nonConfirmee} className="w-full rounded-2xl px-3 py-2 text-sm" style={{ ...inputStyle, opacity: nonConfirmee ? 0.5 : 1 }} />
              <input value={montantPaye} onChange={(e) => setMontantPaye(e.target.value)} type="number" placeholder="Déjà payé (F)" disabled={nonConfirmee} className="w-full rounded-2xl px-3 py-2 text-sm" style={{ ...inputStyle, opacity: nonConfirmee ? 0.5 : 1 }} />
            </div>
            <label className="flex items-center gap-2 mb-2 text-xs font-medium cursor-pointer" style={{ color: ink.ochre }}>
              <input type="checkbox" checked={nonConfirmee} onChange={(e) => setNonConfirmee(e.target.checked)} />
              Commande non confirmée (pas encore de montant)
            </label>
            <div className="flex gap-2">
              <button onClick={handleAddCommande} className="flex-1 rounded-2xl py-2 text-xs font-semibold" style={{ background: ink.petrol, color: "#fff" }}>
                Enregistrer
              </button>
              <button onClick={() => setAjoutOuvert(false)} className="rounded-2xl py-2 px-3 text-xs font-semibold" style={{ background: ink.canvasDeep, color: ink.ink600 }}>
                Annuler
              </button>
            </div>
          </div>
        ) : (
          <button
            onClick={() => setAjoutOuvert(true)}
            className="w-full mb-4 rounded-2xl py-2.5 text-sm font-semibold"
            style={{ background: ink.petrol, color: "#fff" }}
          >
            + Enregistrer une commande
          </button>
        )}

        {client.besoin && (
          <div className="rounded-2xl p-3 mb-6" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
            <div className="flex items-center gap-1.5 text-[11px] mb-1" style={{ color: ink.ink600 }}>
              <ListChecks size={11} /> Besoin exprimé
            </div>
            <div className="text-sm font-medium" style={{ color: ink.ink900 }}>
              {client.besoin}
            </div>
          </div>
        )}

        <div className="rounded-2xl p-4 mb-6" style={{ background: ink.petrol }}>
          <div className="text-[11px] uppercase tracking-wide" style={{ color: "rgba(255,255,255,0.85)" }}>
            Total facturé
          </div>
          <div className="text-2xl font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: "#fff" }}>
            {fmt(totalClient(client))}
          </div>
        </div>

        <h3 className="text-xs font-semibold uppercase tracking-wide mb-3" style={{ color: ink.ink600 }}>
          Commandes & paiements
        </h3>
        {client.commandes.length === 0 ? (
          <p className="text-xs" style={{ color: ink.ink600 }}>
            Aucune commande enregistrée pour l'instant.
          </p>
        ) : (
          <div className="space-y-3">
            {client.commandes.map((cmd) => (
              <CommandeCard
                key={cmd.id}
                cmd={cmd}
                clientId={client.id}
                onAddCout={onAddCout}
                onSolder={onSolderCommande}
                onEdit={onEditCommande}
                onDelete={onDeleteCommande}
                categories={categories}
                onAddCategorie={onAddCategorie}
              />
            ))}
          </div>
        )}

        {relanceOuverte ? (
          <div className="space-y-2 mt-2">
            {templates.map((tpl) => (
              <div key={tpl.id} className="rounded-2xl p-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
                <div className="text-xs font-semibold mb-1" style={{ color: ink.ink900 }}>{tpl.titre}</div>
                <p className="text-[11px] mb-2" style={{ color: ink.ink600 }}>{tpl.texte.replace("{nom}", client.nom)}</p>
                <button
                  onClick={() => copierMessage(tpl)}
                  className="flex items-center gap-1.5 rounded-2xl px-3 py-1.5 text-xs font-semibold"
                  style={{ background: copieId === tpl.id ? "#5FA85B" : ink.ochre, color: "#fff" }}
                >
                  {copieId === tpl.id ? <Check size={12} /> : <Copy size={12} />}
                  {copieId === tpl.id ? "Copié — colle-le dans WhatsApp" : "Copier ce message"}
                </button>
              </div>
            ))}
            <button onClick={() => setRelanceOuverte(false)} className="text-xs" style={{ color: ink.ink600 }}>
              ← Fermer
            </button>
          </div>
        ) : (
          <button
            onClick={() => setRelanceOuverte(true)}
            className="w-full mt-2 rounded-2xl py-2.5 text-sm font-semibold"
            style={{ background: ink.ochre, color: "#fff" }}
          >
            Relancer par WhatsApp
          </button>
        )}
      </div>
    </div>
  );
}

function prochainNumero(clients) {
  const nums = clients.map((c) => parseInt(c.id.split("-")[2], 10)).filter((n) => !isNaN(n));
  const next = (nums.length ? Math.max(...nums) : 0) + 1;
  return `LSG-2026-${String(next).padStart(4, "0")}`;
}

const inputStyle = {
  background: "#111111",
  border: `1px solid ${ink.line}`,
  color: ink.ink900,
  borderRadius: "16px",
};

// ---------------------------------------------------------------------------
// Modale : Enregistrer une commande (recherche client → ajoute une commande)
// ---------------------------------------------------------------------------
function AjouterCommandeModal({ clients, categories, onAddCommande, onAddCategorie, onClose }) {
  const [recherche, setRecherche] = useState("");
  const [clientChoisi, setClientChoisi] = useState(null);
  const [desc, setDesc] = useState("");
  const [montant, setMontant] = useState("");
  const [montantPaye, setMontantPaye] = useState("");
  const [categorieChoisie, setCategorieChoisie] = useState("");
  const [nonConfirmee, setNonConfirmee] = useState(false);

  const resultats = recherche.trim()
    ? clients.filter(
        (c) => c.nom.toLowerCase().includes(recherche.toLowerCase()) || c.id.toLowerCase().includes(recherche.toLowerCase())
      ).slice(0, 20)
    : clients.slice(0, 20);

  function enregistrer() {
    if (!clientChoisi || !desc.trim()) return;
    if (!nonConfirmee && !montant) return;
    onAddCommande(clientChoisi.id, {
      id: Date.now(),
      date: AUJOURD_HUI.toISOString().slice(0, 10),
      description: desc.trim(),
      montant: Number(montant || 0),
      montantPaye: Number(montantPaye || 0),
      categorie: categorieChoisie || null,
      nonConfirmee,
      couts: [],
    });
    onClose();
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.7)" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} className="w-full max-w-sm rounded-[32px] p-5 app-card max-h-[85vh] overflow-y-auto">
        <div className="flex items-center justify-between mb-4">
          <h2 className="text-xl font-extrabold" style={{ color: ink.ink900 }}>Enregistrer une commande</h2>
          <button onClick={onClose} className="p-3 rounded-2xl" style={{ background: ink.panelSoft }}>
            <X size={18} style={{ color: ink.ink600 }} />
          </button>
        </div>

        {!clientChoisi ? (
          <>
            <div className="flex items-center gap-2 rounded-2xl px-3 py-2 mb-3" style={{ background: ink.panel, border: `1px solid ${ink.line}` }}>
              <Search size={14} style={{ color: ink.ink600 }} />
              <input
                value={recherche}
                onChange={(e) => setRecherche(e.target.value)}
                placeholder="Chercher un client par nom ou numéro..."
                className="bg-transparent outline-none text-sm flex-1"
                style={{ color: ink.ink900 }}
                autoFocus
              />
            </div>
            <div className="space-y-2 max-h-72 overflow-y-auto">
              {resultats.map((c) => (
                <button
                  key={c.id}
                  onClick={() => setClientChoisi(c)}
                  className="w-full flex items-center justify-between gap-2 rounded-2xl px-3 py-2.5 text-left"
                  style={{ background: ink.panelSoft, border: `1px solid ${ink.line}` }}
                >
                  <div className="min-w-0">
                    <div className="text-sm font-semibold truncate" style={{ color: ink.ink900 }}>{c.nom}</div>
                    <div className="text-[11px]" style={{ color: ink.ink600 }}>{c.telephone || "—"}</div>
                  </div>
                  <Tampon id={c.id} small />
                </button>
              ))}
              {resultats.length === 0 && (
                <p className="text-xs text-center py-4" style={{ color: ink.ink600 }}>Aucun client trouvé.</p>
              )}
            </div>
          </>
        ) : (
          <>
            <div className="flex items-center gap-2 rounded-2xl px-3 py-2 mb-4" style={{ background: ink.panelSoft, border: `1px solid ${ink.line}` }}>
              <Tampon id={clientChoisi.id} small />
              <span className="text-sm font-semibold" style={{ color: ink.ink900 }}>{clientChoisi.nom}</span>
              <button onClick={() => setClientChoisi(null)} className="ml-auto text-[11px] underline" style={{ color: ink.bleu }}>
                Changer
              </button>
            </div>
            <input value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Description (ex. T-shirts x50)" className="w-full rounded-2xl px-3 py-2 text-sm mb-2" style={inputStyle} />
            <SelecteurCategorie value={categorieChoisie} onChange={setCategorieChoisie} categories={categories} onAddCategorie={onAddCategorie} />
            <div className="grid grid-cols-2 gap-2 mb-2">
              <input value={montant} onChange={(e) => setMontant(e.target.value)} type="number" placeholder="Montant (F)" disabled={nonConfirmee} className="w-full rounded-2xl px-3 py-2 text-sm" style={{ ...inputStyle, opacity: nonConfirmee ? 0.5 : 1 }} />
              <input value={montantPaye} onChange={(e) => setMontantPaye(e.target.value)} type="number" placeholder="Déjà payé (F)" disabled={nonConfirmee} className="w-full rounded-2xl px-3 py-2 text-sm" style={{ ...inputStyle, opacity: nonConfirmee ? 0.5 : 1 }} />
            </div>
            <label className="flex items-center gap-2 mb-4 text-xs font-medium cursor-pointer" style={{ color: ink.ochre }}>
              <input type="checkbox" checked={nonConfirmee} onChange={(e) => setNonConfirmee(e.target.checked)} />
              Commande non confirmée (pas encore de montant)
            </label>
            <button
              onClick={enregistrer}
              disabled={!desc.trim() || (!nonConfirmee && !montant)}
              className="w-full rounded-2xl py-3 text-sm font-bold"
              style={{ background: desc.trim() && (nonConfirmee || montant) ? ink.petrol : ink.ink300, color: "#fff" }}
            >
              Enregistrer la commande
            </button>
          </>
        )}
      </div>
    </div>
  );
}

function NouveauContactModal({ clients, onAdd, onClose }) {
  const [nom, setNom] = useState("");
  const [telephone, setTelephone] = useState("");
  const [source, setSource] = useState("Facebook");
  const [type, setType] = useState("client");
  const [besoin, setBesoin] = useState("");

  const numero = prochainNumero(clients);
  const canSave = nom.trim().length > 0;

  function handleSave() {
    if (!canSave) return;
    onAdd({
      id: numero,
      nom: nom.trim(),
      type,
      telephone,
      source,
      besoin,
      dateEntree: AUJOURD_HUI.toISOString().slice(0, 10),
      statut: "Dossier à suivre",
      commandes: [],
    });
    onClose();
  }

  return (
    <div className="fixed inset-0 z-30 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.7)" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} className="w-full max-w-sm rounded-3xl p-5" style={{ background: ink.canvas, border: `1px solid ${ink.line}` }}>
        <div className="flex items-start justify-between mb-4">
          <div>
            <h2 className="text-lg font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: ink.ink900 }}>
              Nouveau contact
            </h2>
            <Tampon id={numero} small />
          </div>
          <button onClick={onClose} className="p-1.5 rounded-2xl" style={{ background: ink.panel }}>
            <X size={16} style={{ color: ink.ink600 }} />
          </button>
        </div>

        <div className="space-y-3">
          <div>
            <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Nom *</label>
            <input value={nom} onChange={(e) => setNom(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle} placeholder="Ex. Chantal Adjovi" />
          </div>
          <div>
            <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Téléphone / WhatsApp</label>
            <input value={telephone} onChange={(e) => setTelephone(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle} placeholder="+229 ..." />
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Source</label>
              <select value={source} onChange={(e) => setSource(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle}>
                <option>Facebook</option>
                <option>WhatsApp</option>
                <option>TikTok</option>
                <option>Autre</option>
              </select>
            </div>
            <div>
              <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Type</label>
              <select value={type} onChange={(e) => setType(e.target.value)} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle}>
                <option value="client">Client</option>
                <option value="prestataire">Prestataire</option>
              </select>
            </div>
          </div>
          <div>
            <label className="text-[11px] font-medium" style={{ color: ink.ink600 }}>Besoin exprimé</label>
            <textarea value={besoin} onChange={(e) => setBesoin(e.target.value)} rows={2} className="w-full rounded-2xl px-3 py-2 text-sm mt-1" style={inputStyle} placeholder="Ex. 50 t-shirts pour événement" />
          </div>
        </div>

        <button
          onClick={handleSave}
          disabled={!canSave}
          className="w-full mt-5 rounded-2xl py-2.5 text-sm font-semibold transition"
          style={{ background: canSave ? ink.petrol : ink.ink300, color: "#fff" }}
        >
          Ajouter le contact
        </button>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------
export default function CRMPrototype() {
  const [currentUser, setCurrentUser] = useState(null);
  const [view, setView] = useState("dashboard");
  const [clients, setClients] = useState([]);
  const [reglages, setReglages] = useState(REGLAGES_DEFAUT);
  const [templates, setTemplates] = useState(TEMPLATES_DEFAUT);
  const [journal, setJournal] = useState([]);
  const [utilisateurs, setUtilisateurs] = useState([]);
  const [adminAuth, setAdminAuth] = useState(ADMIN_AUTH_DEFAUT);
  const [missions, setMissions] = useState([]);
  const [categories, setCategories] = useState(CATEGORIES_DEFAUT);
  const [loading, setLoading] = useState(true);
  const [selected, setSelected] = useState(null);
  const [filter, setFilter] = useState("tous");
  const [query, setQuery] = useState("");
  const [modalOuvert, setModalOuvert] = useState(false);
  const [importMessage, setImportMessage] = useState(null);
  const [selectedModeGestion, setSelectedModeGestion] = useState(false);
  const [ajouterCommandeOuvert, setAjouterCommandeOuvert] = useState(false);

  function selectClient(c, gestion = false) {
    setSelected(c);
    setSelectedModeGestion(gestion);
  }
  const [menuOuvert, setMenuOuvert] = useState(false);
  const [alerteAcquittee, setAlerteAcquittee] = useState(false);
  const [missionAlerteAcquittee, setMissionAlerteAcquittee] = useState(false);
  const audioCtxRef = useRef(null);

  function initAudio() {
    try {
      if (!audioCtxRef.current) {
        audioCtxRef.current = new (window.AudioContext || window.webkitAudioContext)();
      } else if (audioCtxRef.current.state === "suspended") {
        audioCtxRef.current.resume();
      }
    } catch (e) {}
  }

  function jouerBip() {
    const ctx = audioCtxRef.current;
    if (!ctx) return;
    try {
      [0, 0.4].forEach((delai) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = "sine";
        osc.frequency.value = 880;
        const t = ctx.currentTime + delai;
        gain.gain.setValueAtTime(0.0001, t);
        gain.gain.exponentialRampToValueAtTime(0.25, t + 0.02);
        gain.gain.exponentialRampToValueAtTime(0.0001, t + 0.3);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(t);
        osc.stop(t + 0.32);
      });
    } catch (e) {}
  }

  const clientsOnly = clients.filter((c) => c.type === "client");
  const inactifsUrgents = [...clientsARelancer(clients, reglages.seuilInactiviteJours), ...commandesARelancer(clients, reglages.seuilCommandeInactiveJours)];

  const [maintenantTick, setMaintenantTick] = useState(() => Date.now());
  useEffect(() => {
    const id = setInterval(() => setMaintenantTick(Date.now()), 60000);
    return () => clearInterval(id);
  }, []);

  const missionsSonnantes = missions.filter((m) => {
    if (!m.deadline || m.statut !== "a_faire") return false;
    if (!currentUser || m.assigneA !== currentUser.nom) return false;
    const delta = new Date(m.deadline).getTime() - maintenantTick;
    return delta <= 30 * 60000; // deadline dans moins de 30 min, ou déjà dépassée
  });

  useEffect(() => {
    if (missionsSonnantes.length > 0 && !missionAlerteAcquittee) {
      jouerBip();
      const id = setInterval(jouerBip, 10000);
      return () => clearInterval(id);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser, missionsSonnantes.length, missionAlerteAcquittee, maintenantTick]);

  useEffect(() => {
    if (currentUser?.roles?.includes("commercial") && inactifsUrgents.length > 0 && !alerteAcquittee) {
      jouerBip();
      const id = setInterval(jouerBip, 10000);
      return () => clearInterval(id);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser, inactifsUrgents.length, alerteAcquittee]);

  useEffect(() => {
    (async () => {
      try {
        const [rc, rr, rt, rj, ru, ra, rm, rcat] = await Promise.all([
          storageGet("clients"),
          storageGet("reglages"),
          storageGet("templates"),
          storageGet("journal"),
          storageGet("utilisateurs"),
          storageGet("adminAuth"),
          storageGet("missions"),
          storageGet("categories"),
        ]);
        setClients(rc || CLIENTS_INIT);
        setReglages(rr ? { ...REGLAGES_DEFAUT, ...rr } : REGLAGES_DEFAUT);
        setTemplates(rt || TEMPLATES_DEFAUT);
        setJournal(rj || []);
        setUtilisateurs(ru || []);
        setAdminAuth(ra || ADMIN_AUTH_DEFAUT);
        setMissions(rm || []);
        setCategories(rcat || CATEGORIES_DEFAUT);
      } catch (e) {
        setClients(CLIENTS_INIT);
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  async function persist(next) {
    setClients(next);
    await storageSet("clients", next);
  }

  async function persistReglages(next) {
    setReglages(next);
    await storageSet("reglages", next);
  }

  async function persistTemplates(next) {
    setTemplates(next);
    await storageSet("templates", next);
  }

  async function persistUtilisateurs(next) {
    setUtilisateurs(next);
    await storageSet("utilisateurs", next);
  }

  async function persistAdminAuth(next) {
    setAdminAuth(next);
    await storageSet("adminAuth", next);
  }

  async function persistMissions(next) {
    setMissions(next);
    await storageSet("missions", next);
  }

  async function persistCategories(next) {
    setCategories(next);
    await storageSet("categories", next);
  }

  function handleAddCategorie(nom) {
    if (categories.includes(nom)) return;
    persistCategories([...categories, nom]);
    logAction(`A ajouté l'article "${nom}" aux catégories`);
  }

  async function logAction(action) {
    const { date, heure } = maintenant();
    const entry = { id: Date.now() + Math.random(), date, heure, auteur: currentUser.nom, role: currentUser.roles.join(", "), action };
    const next = [...journal, entry];
    setJournal(next);
    await storageSet("journal", next);
  }

  function handleAdd(nouveauClient) {
    persist([nouveauClient, ...clients]);
    logAction(`A ajouté le contact ${nouveauClient.nom} (${nouveauClient.id})`);
  }

  function handleImportContacts(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!rows.length) {
          setImportMessage({ type: "erreur", texte: "Le fichier ne contient aucune ligne exploitable." });
          return;
        }
        const trouverChamp = (row, ...noms) => {
          for (const cle of Object.keys(row)) {
            if (noms.some((n) => cle.trim().toLowerCase() === n)) return String(row[cle]).trim();
          }
          return "";
        };
        let clientsCourant = clients;
        const nouveaux = [];
        rows.forEach((row) => {
          const nom = trouverChamp(row, "nom", "name", "client");
          if (!nom) return;
          const id = prochainNumero([...clientsCourant, ...nouveaux]);
          const nouveau = {
            id,
            nom,
            type: (trouverChamp(row, "type").toLowerCase() === "prestataire" ? "prestataire" : "client"),
            telephone: trouverChamp(row, "téléphone", "telephone", "numéro", "numero", "phone"),
            source: trouverChamp(row, "source") || "Autre",
            besoin: trouverChamp(row, "besoin"),
            dateEntree: AUJOURD_HUI.toISOString().slice(0, 10),
            statut: "Dossier à suivre",
            commandes: [],
          };
          nouveaux.push(nouveau);
        });
        if (!nouveaux.length) {
          setImportMessage({ type: "erreur", texte: "Aucun contact valide trouvé — vérifie qu'il y a bien une colonne 'Nom'." });
          return;
        }
        persist([...nouveaux, ...clients]);
        logAction(`A importé ${nouveaux.length} contact${nouveaux.length > 1 ? "s" : ""} depuis un fichier`);
        setImportMessage({ type: "succes", texte: `${nouveaux.length} contact${nouveaux.length > 1 ? "s" : ""} importé${nouveaux.length > 1 ? "s" : ""} avec succès.` });
      } catch (err) {
        console.error(err);
        setImportMessage({ type: "erreur", texte: "Impossible de lire ce fichier. Vérifie qu'il s'agit bien d'un Excel (.xlsx) ou d'un CSV." });
      }
    };
    reader.readAsBinaryString(file);
  }

  function handleChangeStatut(id, statut) {
    const client = clients.find((c) => c.id === id);
    const next = clients.map((c) => (c.id === id ? { ...c, statut } : c));
    persist(next);
    setSelected((s) => (s && s.id === id ? { ...s, statut } : s));
    logAction(`A changé le statut de ${client?.nom} → ${statut}`);
  }

  function handleAddCommande(id, commande) {
    const client = clients.find((c) => c.id === id);
    const next = clients.map((c) => (c.id === id ? { ...c, commandes: [...c.commandes, commande] } : c));
    persist(next);
    const updated = next.find((c) => c.id === id);
    setSelected(updated);
    logAction(`A ajouté une commande pour ${client?.nom} (${fmt(commande.montant)})`);
  }

  function handleRelance(client, titreModele) {
    logAction(`A envoyé une relance à ${client.nom} (${titreModele})`);
  }

  function handleAddCout(clientId, commandeId, cout) {
    const client = clients.find((c) => c.id === clientId);
    const cmd = client?.commandes.find((cm) => cm.id === commandeId);
    const next = clients.map((c) =>
      c.id === clientId
        ? {
            ...c,
            commandes: c.commandes.map((cm) =>
              cm.id === commandeId ? { ...cm, couts: [...(cm.couts || []), cout] } : cm
            ),
          }
        : c
    );
    persist(next);
    const updated = next.find((c) => c.id === clientId);
    setSelected(updated);
    logAction(`A ajouté un coût "${cout.description}" (${fmt(cout.montant)}) sur la commande "${cmd?.description}" de ${client?.nom}`);
  }

  function handleSolderCommande(clientId, commandeId) {
    const client = clients.find((c) => c.id === clientId);
    const cmd = client?.commandes.find((cm) => cm.id === commandeId);
    const next = clients.map((c) =>
      c.id === clientId
        ? { ...c, commandes: c.commandes.map((cm) => (cm.id === commandeId ? { ...cm, montantPaye: cm.montant } : cm)) }
        : c
    );
    persist(next);
    setSelected(next.find((c) => c.id === clientId));
    logAction(`A marqué la commande "${cmd?.description}" de ${client?.nom} comme soldée`);
  }

  function handleEditCommande(clientId, commandeId, updates) {
    const client = clients.find((c) => c.id === clientId);
    const next = clients.map((c) =>
      c.id === clientId
        ? { ...c, commandes: c.commandes.map((cm) => (cm.id === commandeId ? { ...cm, ...updates } : cm)) }
        : c
    );
    persist(next);
    setSelected(next.find((c) => c.id === clientId));
    logAction(`A modifié la commande "${updates.description}" de ${client?.nom}`);
  }

  function handleDeleteCommande(clientId, commandeId) {
    const client = clients.find((c) => c.id === clientId);
    const cmd = client?.commandes.find((cm) => cm.id === commandeId);
    const next = clients.map((c) =>
      c.id === clientId ? { ...c, commandes: c.commandes.filter((cm) => cm.id !== commandeId) } : c
    );
    persist(next);
    setSelected(next.find((c) => c.id === clientId));
    logAction(`A supprimé la commande "${cmd?.description}" de ${client?.nom}`);
  }

  function handleEditClient(clientId, updates) {
    const client = clients.find((c) => c.id === clientId);
    const next = clients.map((c) => (c.id === clientId ? { ...c, ...updates } : c));
    persist(next);
    setSelected(next.find((c) => c.id === clientId));
    logAction(`A modifié les infos de ${client?.nom} (${updates.nom})`);
  }

  function handleDeleteClient(clientId) {
    const client = clients.find((c) => c.id === clientId);
    const next = clients.filter((c) => c.id !== clientId);
    persist(next);
    logAction(`A supprimé le client ${client?.nom} (${clientId})`);
  }

  function handleValiderEtape(pole, client) {
    const statutActuel = POLE_STATUT[pole];
    const suivant = STATUT_SUIVANT[statutActuel];
    const next = clients.map((c) => (c.id === client.id ? { ...c, statut: suivant } : c));
    persist(next);
    const details = client.commandes.map((cmd) => cmd.description).join(", ") || "aucune commande enregistrée";
    const verbe = pole === "livraison" ? "A livré la commande de" : `A terminé "${statutActuel}" pour`;
    logAction(`${verbe} ${client.nom} (${details}) → ${suivant}`);
  }

  function handleSetDeadlineEtape(clientId, valeur) {
    const client = clients.find((c) => c.id === clientId);
    const next = clients.map((c) => (c.id === clientId ? { ...c, deadlineEtape: valeur } : c));
    persist(next);
    setSelected((s) => (s && s.id === clientId ? { ...s, deadlineEtape: valeur } : s));
    logAction(`A défini la deadline d'étape pour ${client?.nom}${valeur ? " : " + new Date(valeur).toLocaleString("fr-FR") : " (retirée)"}`);
  }

  function handleImportVisuel(clientId, dataUrl) {
    const client = clients.find((c) => c.id === clientId);
    const next = clients.map((c) => (c.id === clientId ? { ...c, visuelConception: dataUrl } : c));
    persist(next);
    logAction(`A importé le visuel final pour ${client?.nom}`);
  }

  function handleAjouterObservation(texte) {
    if (!texte.trim()) return;
    logAction(`A signalé un besoin particulier : "${texte.trim()}"`);
  }

  function handleAddMission(mission) {
    persistMissions([...missions, mission]);
    logAction(`A assigné une mission à ${mission.assigneA} : "${mission.texte}"`);
  }

  function handleToggleMission(id) {
    const m = missions.find((x) => x.id === id);
    const next = missions.map((x) => (x.id === id ? { ...x, statut: x.statut === "faite" ? "a_faire" : "faite" } : x));
    persistMissions(next);
    if (m) logAction(`A marqué la mission "${m.texte}" comme ${m.statut === "faite" ? "à faire" : "faite"}`);
  }

  function handleDeleteMission(id) {
    persistMissions(missions.filter((x) => x.id !== id));
  }

  function roleLabel(r) {
    if (r === "admin") return "Administrateur";
    if (r === "commercial") return "Commercial(e)";
    return POLE_LABEL[r] || r;
  }

  const isAdmin = !!currentUser?.roles?.includes("admin");
  const hasRole = (r) => !!currentUser?.roles?.includes(r);
  const personnesMissions = [adminAuth.nom, ...utilisateurs.map((u) => u.nom)].filter(
    (n, i, arr) => arr.indexOf(n) === i
  );

  const navAdmin = [
    { id: "dashboard", label: "Tableau de bord", Icon: LayoutGrid },
    { id: "clients", label: "Commande", Icon: Users },
    { id: "base_clients", label: "Base de données clients", Icon: Database },
    { id: "fidelite", label: "Fidélité", Icon: Trophy },
    { id: "bilan", label: "Bilan", Icon: FileBarChart },
    { id: "missions", label: "Attribution de mission", Icon: ListChecks },
    { id: "reglages", label: "Réglages", Icon: Settings },
    { id: "journal", label: "Journal d'activité", Icon: ClipboardList },
    { id: "utilisateurs", label: "Utilisateurs", Icon: UserCog },
  ];

  let nav = [];
  if (isAdmin) {
    nav = navAdmin;
  } else {
    if (hasRole("commercial")) {
      nav.push(
        { id: "dashboard", label: "Tableau de bord", Icon: LayoutGrid },
        { id: "clients", label: "Commande", Icon: Users },
        { id: "base_clients", label: "Client", Icon: Database },
        { id: "parcours", label: "Parcours", Icon: Package },
        { id: "fidelite", label: "Fidélité", Icon: Trophy },
        { id: "messages", label: "Messages de relance", Icon: MessageSquare },
        { id: "monbilan", label: "Mon bilan du jour", Icon: ClipboardList }
      );
    }
    if (hasRole("graphiste")) {
      nav.push({ id: "dashboard_graphiste", label: "Tableau de bord", Icon: LayoutGrid });
      nav.push({ id: "file_graphiste", label: "Commande", Icon: Palette });
      nav.push({ id: "monbilan", label: "Mon bilan d'activité", Icon: ClipboardList });
    }
    if (hasRole("production")) {
      nav.push({ id: "dashboard_production", label: "Tableau de bord", Icon: LayoutGrid });
      nav.push({ id: "file_production", label: "Commande", Icon: Boxes });
      nav.push({ id: "bilan_production", label: "Rapport d'activité", Icon: ClipboardList });
    }
    if (hasRole("livraison")) nav.push({ id: "file_livraison", label: "File — Livraison", Icon: Truck });
    nav.push({ id: "missions", label: "Mission spécifique", Icon: ListChecks });
  }

  function defaultViewFor(roles) {
    if (roles.includes("admin") || roles.includes("commercial")) return "dashboard";
    if (roles.includes("graphiste")) return "dashboard_graphiste";
    if (roles.includes("production")) return "dashboard_production";
    if (roles.includes("livraison")) return "file_livraison";
    return "missions";
  }

  if (!currentUser) {
    return (
      <LoginScreen
        onLogin={(u) => { initAudio(); setCurrentUser(u); setView(defaultViewFor(u.roles)); setAlerteAcquittee(false); setMissionAlerteAcquittee(false); }}
        adminAuth={adminAuth}
        utilisateurs={utilisateurs}
      />
    );
  }

  return (
    <div className="app-shell min-h-screen w-full flex" style={{ background: ink.canvas }}>
      <GlobalStyle />

      {hasRole("commercial") && inactifsUrgents.length > 0 && !alerteAcquittee && (
        <div
          className="fixed top-0 left-0 right-0 z-40 flex flex-wrap items-center justify-between gap-2 px-4 py-3"
          style={{ background: ink.rouge }}
        >
          <div className="flex items-center gap-2 text-white text-xs font-semibold">
            <AlertTriangle size={16} className="animate-pulse shrink-0" />
            {inactifsUrgents.length} client{inactifsUrgents.length > 1 ? "s" : ""} à relancer (clients inactifs + commandes en attente)
          </div>
          <div className="flex items-center gap-2 ml-auto">
            <button
              onClick={() => { setView("dashboard"); setModalOuvert(false); setSelected(null); }}
              className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl"
              style={{ background: "rgba(255,255,255,0.2)", color: "#fff" }}
            >
              Voir la liste
            </button>
            <button
              onClick={() => setAlerteAcquittee(true)}
              className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl"
              style={{ background: "#fff", color: ink.rouge }}
            >
              Valider — j'ai vu
            </button>
          </div>
        </div>
      )}

      {missionsSonnantes.length > 0 && !missionAlerteAcquittee && (
        <div
          className="fixed left-0 right-0 z-40 flex flex-wrap items-center justify-between gap-2 px-4 py-3"
          style={{
            background: ink.ochre,
            top: hasRole("commercial") && inactifsUrgents.length > 0 && !alerteAcquittee ? 52 : 0,
          }}
        >
          <div className="flex items-center gap-2 text-white text-xs font-semibold">
            <ListChecks size={16} className="animate-pulse shrink-0" />
            {missionsSonnantes.length} mission{missionsSonnantes.length > 1 ? "s" : ""} à terminer — deadline proche ou dépassée
          </div>
          <div className="flex items-center gap-2 ml-auto">
            <button
              onClick={() => { setView("missions"); setModalOuvert(false); setSelected(null); }}
              className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl"
              style={{ background: "rgba(255,255,255,0.2)", color: "#fff" }}
            >
              Voir mes missions
            </button>
            <button
              onClick={() => setMissionAlerteAcquittee(true)}
              className="text-[11px] font-semibold px-2.5 py-1.5 rounded-2xl"
              style={{ background: "#fff", color: ink.ochreDeep }}
            >
              Valider — j'ai vu
            </button>
          </div>
        </div>
      )}

      <div
        className="w-full flex"
        style={{
          paddingTop:
            (hasRole("commercial") && inactifsUrgents.length > 0 && !alerteAcquittee ? 52 : 0) +
            (missionsSonnantes.length > 0 && !missionAlerteAcquittee ? 52 : 0),
        }}
      >
      {/* Sidebar */}
      <aside className="w-56 shrink-0 p-5 hidden sm:flex flex-col" style={{ borderRight: `1px solid ${ink.line}` }}>
        <div className="mb-8 flex items-center gap-3">
          <LogoMark size={44} />
          <div>
            <div
              className="text-base font-bold leading-tight"
              style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: ink.ink900 }}
            >
              LE SÉRIGRAPHE
            </div>
            <div className="text-[11px] tracking-wide" style={{ color: ink.ink600 }}>
              {currentUser.roles.map(roleLabel).join(" · ")} — {currentUser.nom}
            </div>
          </div>
        </div>
        <nav className="space-y-1">
          {nav.map(({ id, label, Icon }) => (
            <button
              key={id}
              onClick={() => setView(id)}
              className="w-full flex items-center gap-2.5 px-3 py-2 rounded-2xl text-sm font-medium transition"
              style={{
                background: view === id ? ink.petrol : "transparent",
                color: view === id ? "#fff" : ink.ink600,
              }}
            >
              <Icon size={15} />
              {label}
            </button>
          ))}
        </nav>
        <div className="mt-auto pt-6 space-y-2">
          <button
            onClick={() => setCurrentUser(null)}
            className="w-full flex items-center justify-center gap-1.5 rounded-2xl py-2 text-xs font-medium"
            style={{ background: "transparent", color: ink.ink600, border: `1px solid ${ink.line}` }}
          >
            <LogOut size={13} /> Changer de compte
          </button>
        </div>
      </aside>

      {/* Main */}
      <main className="flex-1 w-full min-w-0 p-6 pb-28 sm:pb-6 max-w-5xl overflow-x-hidden">
        <div className="mb-5 flex items-center justify-between sm:hidden">
          <div className="flex items-center gap-2">
            <LogoMark size={32} />
            <div className="text-base font-bold" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: ink.ink900 }}>
              LE SÉRIGRAPHE
            </div>
          </div>
          <button
            onClick={() => setMenuOuvert((m) => !m)}
            className="h-9 w-9 rounded-2xl flex items-center justify-center"
            style={{ background: ink.panel, border: `1px solid ${ink.line}` }}
          >
            <Settings size={16} style={{ color: ink.ink600 }} />
          </button>
        </div>

        {menuOuvert && (
          <PremiumCard className="p-2 mb-5 sm:hidden">
            {nav.filter((n) => n.id !== "base_clients").map(({ id, label, Icon }) => (
              <button
                key={id}
                onClick={() => { setView(id); setMenuOuvert(false); }}
                className="w-full flex items-center gap-3 px-3 py-3 text-sm font-medium rounded-2xl"
                style={{ background: view === id ? ink.orangeSoft : "transparent", color: view === id ? ink.orange : ink.ink900 }}
              >
                <Icon size={16} />
                {label}
              </button>
            ))}
            <div style={{ borderTop: `1px solid ${ink.line}` }} className="mt-1 pt-1">
              <button
                onClick={() => { setCurrentUser(null); setMenuOuvert(false); }}
                className="w-full flex items-center gap-3 px-3 py-3 text-sm font-medium rounded-2xl"
                style={{ color: ink.ink600 }}
              >
                <LogOut size={16} /> Changer de compte
              </button>
            </div>
          </PremiumCard>
        )}

        {selected ? null : (
          <h1 className="text-2xl font-bold mb-6" style={{ fontFamily: "'Inter', sans-serif", fontWeight: 800, color: ink.ink900 }}>
            {view === "commande_attente" ? "Commande en attente" : view === "livraison_attente" ? "Livraison en attente" : view === "a_relancer" ? "À relancer" : nav.find((n) => n.id === view)?.label}
          </h1>
        )}

        {selected ? (
          <FicheClient
            client={selected}
            onClose={() => setSelected(null)}
            onChangeStatut={handleChangeStatut}
            onAddCommande={handleAddCommande}
            onAddCout={handleAddCout}
            onSolderCommande={handleSolderCommande}
            onEditCommande={handleEditCommande}
            onDeleteCommande={handleDeleteCommande}
            templates={templates}
            onRelance={handleRelance}
            categories={categories}
            onAddCategorie={handleAddCategorie}
            onSetDeadlineEtape={handleSetDeadlineEtape}
            onEditClient={handleEditClient}
            onDeleteClient={handleDeleteClient}
            modeGestion={selectedModeGestion}
          />
        ) : loading ? (
          <p className="text-sm" style={{ color: ink.ink600 }}>Chargement…</p>
        ) : (
          <>
            {view === "dashboard" && (isAdmin || hasRole("commercial")) && (
              <Dashboard clients={clients} reglages={reglages} currentUser={currentUser} missions={missions} setView={setView} />
            )}
            {view === "clients" && (isAdmin || hasRole("commercial")) && (
              <ClientsList
                clients={clients}
                onSelect={selectClient}
                filter={filter}
                setFilter={setFilter}
                query={query}
                setQuery={setQuery}
              />
            )}
            {view === "parcours" && (isAdmin || hasRole("commercial")) && <Parcours clients={clients} onSelect={selectClient} />}
            {view === "fidelite" && (isAdmin || hasRole("commercial")) && <Fidelite clients={clients} />}
            {view === "base_clients" && (isAdmin || hasRole("commercial")) && (
              <BaseClients
                clients={clients}
                categories={categories}
                onSelect={(c) => selectClient(c, true)}
                onImportContacts={isAdmin ? handleImportContacts : undefined}
                importMessage={importMessage}
                onClearImportMessage={() => setImportMessage(null)}
                isAdmin={isAdmin}
                onNouveauContact={() => setModalOuvert(true)}
              />
            )}
            {view === "commande_attente" && (isAdmin || hasRole("commercial")) && (
              <ClientsList
                clients={clients}
                onSelect={selectClient}
                filter={filter}
                setFilter={setFilter}
                query={query}
                setQuery={setQuery}
                statutForce={["Conception", "En production"]}
              />
            )}
            {view === "livraison_attente" && (isAdmin || hasRole("commercial")) && (
              <ClientsList
                clients={clients}
                onSelect={selectClient}
                filter={filter}
                setFilter={setFilter}
                query={query}
                setQuery={setQuery}
                statutForce={["Prêt / à livrer"]}
              />
            )}
            {view === "a_relancer" && (isAdmin || hasRole("commercial")) && (
              <ARelancerListe clients={clients} reglages={reglages} onSelect={selectClient} />
            )}
            {view === "bilan" && isAdmin && <Bilan clients={clients} />}
            {view === "reglages" && isAdmin && (
              <Reglages
                reglages={reglages}
                onSave={persistReglages}
                adminAuth={adminAuth}
                onSaveAdminAuth={persistAdminAuth}
                categories={categories}
                onSaveCategories={persistCategories}
              />
            )}
            {view === "journal" && isAdmin && <Journal journal={journal} />}
            {view === "utilisateurs" && isAdmin && (
              <Utilisateurs utilisateurs={utilisateurs} onSave={persistUtilisateurs} />
            )}
            {view === "messages" && hasRole("commercial") && <Messages templates={templates} onSave={persistTemplates} />}
            {view === "monbilan" && hasRole("commercial") && <MonBilan journal={journal} currentUser={currentUser} />}
            {view === "monbilan" && !hasRole("commercial") && hasRole("graphiste") && (
              <BilanGraphiste clients={clients} journal={journal} currentUser={currentUser} />
            )}
            {view === "dashboard_graphiste" && hasRole("graphiste") && (
              <GraphisteDashboard clients={clients} journal={journal} currentUser={currentUser} setView={setView} />
            )}
            {view === "file_graphiste" && hasRole("graphiste") && (
              <FileAttente
                clients={clients}
                pole="graphiste"
                onValider={(c) => handleValiderEtape("graphiste", c)}
                onImportVisuel={handleImportVisuel}
              />
            )}
            {view === "dashboard_production" && hasRole("production") && (
              <ProductionDashboard clients={clients} journal={journal} currentUser={currentUser} setView={setView} onAjouterObservation={handleAjouterObservation} />
            )}
            {view === "file_production" && hasRole("production") && (
              <FileAttente clients={clients} pole="production" onValider={(c) => handleValiderEtape("production", c)} currentUser={currentUser} />
            )}
            {view === "bilan_production" && hasRole("production") && (
              <BilanProduction clients={clients} journal={journal} currentUser={currentUser} onAjouterObservation={handleAjouterObservation} />
            )}
            {view === "file_livraison" && hasRole("livraison") && (
              <FileAttente clients={clients} pole="livraison" onValider={(c) => handleValiderEtape("livraison", c)} currentUser={currentUser} />
            )}
            {view === "missions" && (
              <Missions
                missions={missions}
                personnes={personnesMissions}
                currentUser={currentUser}
                isAdmin={isAdmin}
                onAdd={handleAddMission}
                onToggle={handleToggleMission}
                onDelete={handleDeleteMission}
              />
            )}
          </>
        )}
      </main>

      {modalOuvert && (
        <NouveauContactModal clients={clients} onAdd={handleAdd} onClose={() => setModalOuvert(false)} />
      )}
      {ajouterCommandeOuvert && (
        <AjouterCommandeModal
          clients={clients}
          categories={categories}
          onAddCommande={handleAddCommande}
          onAddCategorie={handleAddCategorie}
          onClose={() => setAjouterCommandeOuvert(false)}
        />
      )}
      <BottomNav
        view={view}
        setView={setView}
        isAdmin={isAdmin}
        hasRole={hasRole}
        onAjouterCommande={() => setAjouterCommandeOuvert(true)}
      />
      </div>
    </div>
  );
}
