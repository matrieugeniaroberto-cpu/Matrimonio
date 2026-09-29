import React, { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { supabase, DATA_TABLE, DATA_ROW_ID } from "./lib/supabase";
import {
  LayoutGrid, Users, CheckSquare, PieChart, Briefcase, Table2, Clock,
  Camera, Plus, X, Trash2, Check, ChevronRight, Sun, Moon, AlertCircle,
  CalendarClock, Palette, Utensils, Phone, Mail, Tag, StickyNote,
  Link2, Mic, ImageIcon, PenTool, MapPin, Bell, Settings as SettingsIcon,
  Mailbox, ExternalLink, Sparkles, Upload, Printer, GripVertical,
  Building2, Star, Scale, Contact, Pencil
} from "lucide-react";

/* ============================== helpers ============================== */
const uid = () => Math.random().toString(36).slice(2, 10);
function reorderById(arr, draggedId, targetId) {
  if (!draggedId || draggedId === targetId) return arr;
  const fromIdx = arr.findIndex((x) => x.id === draggedId);
  const toIdx = arr.findIndex((x) => x.id === targetId);
  if (fromIdx < 0 || toIdx < 0) return arr;
  const copy = [...arr];
  const [moved] = copy.splice(fromIdx, 1);
  copy.splice(toIdx, 0, moved);
  return copy;
}
function shade(hex, pct) {
  const f = parseInt(hex.slice(1), 16), t = pct < 0 ? 0 : 255, p = Math.abs(pct);
  const R = f >> 16, G = (f >> 8) & 0xff, B = f & 0xff;
  return "#" + (0x1000000 + (Math.round((t - R) * p) + R) * 0x10000 +
    (Math.round((t - G) * p) + G) * 0x100 + (Math.round((t - B) * p) + B)).toString(16).slice(1);
}
function fmtMoney(v, currency) {
  const n = Number(v || 0);
  return currency === "UF" ? `${n.toLocaleString("es-CL", { minimumFractionDigits: 2 })} UF` : `$${n.toLocaleString("es-CL")}`;
}
function daysUntil(d) { if (!d) return null; return Math.ceil((new Date(d) - new Date()) / 86400000); }
function mapsEmbed(address) { return `https://www.google.com/maps?q=${encodeURIComponent(address)}&output=embed`; }
function mapsLink(address) { return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(address)}`; }
function truthy(v) { return /^(s[ií]|true|yes|x|1)$/i.test(String(v || "").trim()); }
function pick(row, keys) {
  for (const k of keys) for (const rk of Object.keys(row)) if (rk.trim().toLowerCase() === k.toLowerCase()) return row[rk];
  return "";
}
function matchEvents(text, events) {
  if (!text) return [];
  const names = String(text).split(/[,;]/).map((s) => s.trim().toLowerCase()).filter(Boolean);
  return events.filter((e) => names.some((n) => e.name.toLowerCase().includes(n))).map((e) => e.id);
}
function matchEvent(text, events) {
  if (!text) return null;
  const n = String(text).trim().toLowerCase();
  const found = events.find((e) => e.name.toLowerCase().includes(n));
  return found ? found.id : null;
}
async function readExcelRows(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}
function compressImage(file, maxWidth = 1400, quality = 0.82) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = Math.max(1, Math.round(img.width * scale));
        canvas.height = Math.max(1, Math.round(img.height * scale));
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}
function ImportExcel({ T, label, hint, onRows }) {
  const ref = useRef(null);
  async function handle(e) {
    const f = e.target.files?.[0];
    if (!f) return;
    try {
      const rows = await readExcelRows(f);
      onRows(rows);
    } catch (err) { alert("No se pudo leer el archivo. Usa un .xlsx, .xls o .csv exportado desde Excel o Google Sheets."); }
    e.target.value = "";
  }
  return (
    <span className="inline-flex flex-col items-end">
      <button onClick={() => ref.current?.click()} title={hint} className="flex items-center gap-1.5 rounded-full font-medium px-3 py-1.5 text-xs" style={{ background: "transparent", color: T.primaryDark, border: `1px solid ${T.soft}` }}>
        <Upload size={13} /> {label || "Importar Excel"}
      </button>
      <input ref={ref} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handle} />
    </span>
  );
}
function printHtml(title, bodyHtml) {
  const win = window.open("", "_blank");
  if (!win) { alert("Habilita las ventanas emergentes para poder imprimir."); return; }
  win.document.write(`<html><head><title>${title}</title><style>
    body{font-family:system-ui,sans-serif;color:#3B2A20;padding:24px;}
    h1{font-family:Georgia,serif;font-size:22px;margin-bottom:16px;}
    table{width:100%;border-collapse:collapse;font-size:13px;}
    th,td{border:1px solid #ddd;padding:6px 10px;text-align:left;}
    th{background:#f2ece2;}
  </style></head><body>${bodyHtml}</body></html>`);
  win.document.close();
  win.focus();
  win.print();
}

/* ============================== theme ============================== */
// Light mode derives from the active colour scheme. Dark mode is a fixed
// near-black / soft-rose palette, independent of the scheme.
function buildTheme(mode, scheme) {
  if (mode === "dark") {
    return {
      mode, bg: "#121013", card: "#1B181C", cardAlt: "#242025", line: "#332C31",
      text: "#EDE6E8", muted: "#9C8F94",
      primary: "#D98FA8", primaryDark: "#C97591", secondary: "#C7A6AE",
      secondaryDark: "#B98D97", soft: "#3A2E33", alert: "#E2708A",
    };
  }
  const [c1, c2, c3, c4] = scheme.colors;
  return {
    mode, bg: "#F6EEE2", card: "#FFFCF7", cardAlt: "#FBF3E7", line: "#E8DCC7",
    text: "#3B2A20", muted: "#8A7A64",
    primary: c1, primaryDark: shade(c1, -0.25),
    secondary: c2, secondaryDark: shade(c2, -0.25),
    soft: c3, alert: c4 || "#7C3140",
  };
}
const DEFAULT_SCHEMES = [
  { id: "s1", name: "Autumn Yuyito", colors: ["#B5652E", "#767A4A", "#CDA36C", "#7C3140"] },
  { id: "s2", name: "Terracota y oliva", colors: ["#C86F4A", "#7A7B4A", "#F2E3D1", "#4A3B32"] },
  { id: "s3", name: "Blush y sage", colors: ["#C9A0A0", "#8A9A7E", "#F3EAE1", "#3F3A34"] },
  { id: "s4", name: "Vino y dorado", colors: ["#6E2A34", "#C9A24B", "#F1E7D6", "#2E2420"] },
  { id: "s5", name: "Eucalipto y crema", colors: ["#7C9182", "#C9BBA0", "#F5F1E7", "#3B4A3F"] },
  { id: "s6", name: "Lavanda y gris", colors: ["#9A8FB0", "#B9AEB0", "#F0EAF0", "#443E4A"] },
  { id: "s7", name: "Clásico blanco y negro", colors: ["#B79A5B", "#2B2B2B", "#F7F5F0", "#8A8A8A"] },
  { id: "s8", name: "Desert rose", colors: ["#C9776B", "#B5652E", "#EFDCC8", "#5A3A2E"] },
];
const RSVP_STYLES = (T) => ({
  attending: { bg: T.mode === "dark" ? "#2E3A2A" : "#E7EBD9", fg: T.secondaryDark, label: "Confirma" },
  pending: { bg: T.mode === "dark" ? "#3A3020" : "#F3E4CB", fg: T.primaryDark, label: "Sin responder" },
  declined: { bg: T.mode === "dark" ? "#3A2226" : "#EFD9DC", fg: T.alert, label: "No asiste" },
});

const EVENT_TYPES = ["Ceremonia", "Recepción", "After", "Otro"];
const initialEvents = [
  { id: "e1", name: "Ceremonia Civil", type: "Ceremonia", date: "2026-11-14", time: "12:00", address: "" },
  { id: "e2", name: "Ceremonia Religiosa", type: "Ceremonia", date: "2026-12-05", time: "17:00", address: "" },
];
const DEFAULT_GROUPS = ["Familia Rober", "Familia Yuyito", "Amigos L1 (cercanos)", "Amigos Universidad"];
const DEFAULT_BUDGET_CATS = ["Lugar", "Catering", "Fotografía", "Flores", "Música", "Vestuario", "Papelería", "Otro"];
const CATERING_SECTIONS = ["Entrada", "Plato principal", "Postre", "Cóctail"];
const MODULE_TAGS = ["Vestuario", "Flores", "Decoración", "Catering", "Fotografía", "Papelería", "Música", "Otro"];
const TASK_TEMPLATE = [
  { title: "Reservar el lugar", tag: "Logística" },
  { title: "Contratar catering", tag: "Proveedores" },
  { title: "Definir el menú", tag: "Proveedores" },
  { title: "Enviar save the date", tag: "Papelería" },
  { title: "Enviar invitaciones", tag: "Papelería" },
  { title: "Probar vestido / traje", tag: "Vestuario" },
  { title: "Confirmar proveedores", tag: "Proveedores" },
  { title: "Armar cronograma del día", tag: "Logística" },
  { title: "Enviar lista de canciones al DJ", tag: "Logística" },
  { title: "Confirmar transporte", tag: "Logística" },
  { title: "Hacer el plano de mesas", tag: "Logística" },
  { title: "Coordinar ensayo", tag: "Logística" },
];
const DEFAULT_USERS = [
  { id: "u1", name: "Eugenia (Yuyito)", email: "eugenia.neira.h@gmail.com", role: "Novia" },
  { id: "u2", name: "Roberto", email: "robertoavendanoceballos@gmail.com", role: "Novio" },
  { id: "u3", name: "Cuenta de respaldo", email: "matrieugeniaroberto@gmail.com", role: "Respaldo" },
];

/* ============================== App ============================== */
export default function WeddingPlanner() {
  const [mode, setMode] = useState("light");
  const [themeName] = useState("Clásico");
  const [schemes, setSchemes] = useState(DEFAULT_SCHEMES);
  const [activeScheme, setActiveScheme] = useState("s1");
  const T = buildTheme(mode, schemes.find((s) => s.id === activeScheme) || schemes[0]);

  const [tab, setTab] = useState("overview");
  const [couple, setCouple] = useState({ a: "Yuyito", b: "Roberto", photo: null });
  const [engagementDate, setEngagementDate] = useState("");
  const [events, setEvents] = useState(initialEvents);
  const [etapas, setEtapas] = useState([]);
  const [activeEvent, setActiveEvent] = useState("all");

  const [groups, setGroups] = useState(DEFAULT_GROUPS);
  const [guests, setGuests] = useState([]);
  const [todoTags, setTodoTags] = useState(["Papelería", "Vestuario", "Logística", "Proveedores"]);
  const [todos, setTodos] = useState([]);
  const [budgetCats, setBudgetCats] = useState(DEFAULT_BUDGET_CATS);
  const [budget, setBudget] = useState([]);
  const [suppliers, setSuppliers] = useState([]);
  const [catering, setCatering] = useState({});
  const [seating, setSeating] = useState({});
  const [schedules, setSchedules] = useState({});
  const [scrapbook, setScrapbook] = useState([]);
  const [inspiration, setInspiration] = useState([]);
  const [venues, setVenues] = useState([]);
  const [invitations, setInvitations] = useState([]);
  const [users, setUsers] = useState(DEFAULT_USERS);
  const [notifPrefs, setNotifPrefs] = useState({ rsvp: true, tasks: true, suppliers: true });
  const [dismissed, setDismissed] = useState([]);
  const [reminders, setReminders] = useState([]);

  const [loaded, setLoaded] = useState(false);
  const [saveStatus, setSaveStatus] = useState("idle"); // idle | saving | saved | error
  const [saveError, setSaveError] = useState("");
  const saveTimer = useRef(null);
  const lastSavedAtRef = useRef(null); // guarda el updated_at que nosotros mismos escribimos, para ignorar el eco de Realtime

  function applyData(d) {
    if (d.couple) setCouple(d.couple);
    if (d.engagementDate !== undefined) setEngagementDate(d.engagementDate);
    if (d.events) setEvents(d.events);
    if (d.etapas) setEtapas(d.etapas);
    if (d.groups) setGroups(d.groups);
    if (d.guests) setGuests(d.guests);
    if (d.todoTags) setTodoTags(d.todoTags);
    if (d.todos) setTodos(d.todos);
    if (d.budgetCats) setBudgetCats(d.budgetCats);
    if (d.budget) setBudget(d.budget);
    if (d.suppliers) setSuppliers(d.suppliers);
    if (d.catering) setCatering(d.catering);
    if (d.seating) setSeating(d.seating);
    if (d.schedules) setSchedules(d.schedules);
    if (d.scrapbook) setScrapbook(d.scrapbook);
    if (d.inspiration) setInspiration(d.inspiration);
    if (d.venues) setVenues(d.venues);
    if (d.invitations) setInvitations(d.invitations);
    if (d.users) setUsers(d.users);
    if (d.notifPrefs) setNotifPrefs(d.notifPrefs);
    if (d.dismissed) setDismissed(d.dismissed);
    if (d.reminders) setReminders(d.reminders);
    if (d.schemes) setSchemes(d.schemes);
    if (d.activeScheme) setActiveScheme(d.activeScheme);
    if (d.mode) setMode(d.mode);
  }
  function snapshot() {
    return { couple, engagementDate, events, etapas, groups, guests, todoTags, todos, budgetCats, budget, suppliers, catering, seating, schedules, scrapbook, inspiration, venues, invitations, users, notifPrefs, dismissed, reminders, schemes, activeScheme, mode };
  }

  // Carga inicial desde Supabase.
  useEffect(() => {
    (async () => {
      try {
        const { data, error } = await supabase.from(DATA_TABLE).select("data, updated_at").eq("id", DATA_ROW_ID).maybeSingle();
        if (error) throw error;
        if (data?.data) applyData(data.data);
        if (data?.updated_at) lastSavedAtRef.current = data.updated_at;
      } catch (e) {
        console.error("No se pudo cargar desde Supabase:", e.message);
        setSaveStatus("error");
        setSaveError(e.message || String(e));
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  // Escucha cambios en tiempo real (por ejemplo, si Roberto edita desde otro dispositivo).
  useEffect(() => {
    const channel = supabase
      .channel("wedding_data_changes")
      .on("postgres_changes", { event: "*", schema: "public", table: DATA_TABLE, filter: `id=eq.${DATA_ROW_ID}` }, (payload) => {
        const row = payload.new;
        if (!row || row.updated_at === lastSavedAtRef.current) return; // es nuestro propio guardado, ignorar
        applyData(row.data);
        lastSavedAtRef.current = row.updated_at;
      })
      .subscribe();
    return () => { supabase.removeChannel(channel); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Guardado automático (debounced) cada vez que cambian los datos.
  useEffect(() => {
    if (!loaded) return;
    setSaveStatus("saving");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      const nowIso = new Date().toISOString();
      try {
        const { error } = await supabase.from(DATA_TABLE).upsert({ id: DATA_ROW_ID, data: snapshot(), updated_at: nowIso });
        if (error) throw error;
        lastSavedAtRef.current = nowIso;
        setSaveStatus("saved");
        setSaveError("");
      } catch (e) {
        console.error("No se pudo guardar en Supabase:", e.message);
        setSaveStatus("error");
        setSaveError(e.message || String(e));
      }
    }, 700);
    return () => clearTimeout(saveTimer.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loaded, couple, engagementDate, events, etapas, groups, guests, todoTags, todos, budgetCats, budget, suppliers, catering, seating, schedules, scrapbook, inspiration, venues, invitations, users, notifPrefs, dismissed, reminders, schemes, activeScheme, mode]);

  async function saveNow() {
    setSaveStatus("saving");
    const nowIso = new Date().toISOString();
    try {
      const { error } = await supabase.from(DATA_TABLE).upsert({ id: DATA_ROW_ID, data: snapshot(), updated_at: nowIso });
      if (error) throw error;
      lastSavedAtRef.current = nowIso;
      setSaveStatus("saved");
      setSaveError("");
    } catch (e) { setSaveStatus("error"); setSaveError(e.message || String(e)); }
  }
  async function clearSavedData() {
    if (!window.confirm("¿Borrar todos los datos guardados? Esto no se puede deshacer.")) return;
    try { await supabase.from(DATA_TABLE).delete().eq("id", DATA_ROW_ID); } catch (e) { /* nada guardado aún */ }
    window.location.reload();
  }

  const fileRef = useRef(null);
  function handlePhoto(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    compressImage(file).then((dataUrl) => setCouple((c) => ({ ...c, photo: dataUrl, photoPos: undefined }))).catch(() => alert("No se pudo procesar esa imagen."));
  }

  function exportData() {
    const blob = new Blob([JSON.stringify(snapshot(), null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "boda-datos.json"; a.click();
    URL.revokeObjectURL(url);
  }
  function importData(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try { applyData(JSON.parse(reader.result)); }
      catch (e) { alert("El archivo no es un respaldo válido."); }
    };
    reader.readAsText(file);
  }

  const NAV = [
    { id: "overview", label: "Resumen", icon: LayoutGrid, scoped: false },
    { id: "guests", label: "Invitados", icon: Users, scoped: true },
    { id: "ceremonies", label: "Eventos", icon: MapPin, scoped: false },
    { id: "venues", label: "Lugares", icon: Building2, scoped: false },
    { id: "todos", label: "Tareas", icon: CheckSquare, scoped: false },
    { id: "budget", label: "Presupuesto", icon: PieChart, scoped: true },
    { id: "suppliers", label: "Proveedores", icon: Briefcase, scoped: false },
    { id: "catering", label: "Catering", icon: Utensils, scoped: true },
    { id: "invitations", label: "Invitaciones", icon: Mailbox, scoped: true },
    { id: "seating", label: "Mesas", icon: Table2, scoped: true },
    { id: "schedule", label: "Cronograma", icon: Clock, scoped: true },
    { id: "scrapbook", label: "Scrapbook", icon: StickyNote, scoped: false },
    { id: "inspiration", label: "Inspiración", icon: Sparkles, scoped: false },
    { id: "schemes", label: "Esquemas de color", icon: Palette, scoped: false },
    { id: "notifications", label: "Notificaciones", icon: Bell, scoped: false },
    { id: "settings", label: "Configuración", icon: SettingsIcon, scoped: false },
  ];

  const openTasks = todos.filter((t) => !t.done).length;
  const overdueTasks = todos.filter((t) => !t.done && t.due && new Date(t.due) < new Date()).length;
  const pendingRsvp = guests.filter((g) => g.rsvp === "pending").length;
  const unpaidSuppliers = suppliers.filter((s) => s.paymentStatus !== "Pagado").length;

  return (
    <div className="flex w-full min-h-screen" style={{ background: T.bg, color: T.text, fontFamily: "system-ui, sans-serif" }}>
      <aside className="w-60 shrink-0 p-5 flex flex-col gap-1 overflow-y-auto" style={{ borderRight: `1px solid ${T.line}` }}>
        <div className="flex items-center gap-3 mb-2">
          <div className="w-11 h-11 rounded-full overflow-hidden flex items-center justify-center shrink-0" style={{ background: T.soft, border: `2px solid ${T.primary}` }}>
            {couple.photo ? <img src={couple.photo} alt="" className="w-full h-full object-cover" style={{ objectPosition: couple.photoPos || "50% 50%" }} /> : <Camera size={18} color={T.primaryDark} />}
          </div>
          <div style={{ fontFamily: "Georgia, serif" }} className="text-lg leading-tight">
            {couple.a} <span style={{ color: T.primary }}>&</span> {couple.b}
          </div>
        </div>
        <button onClick={() => setMode(mode === "light" ? "dark" : "light")} className="flex items-center gap-2 text-xs mb-2 px-2 py-1.5 rounded-lg w-fit" style={{ color: T.muted, border: `1px solid ${T.line}` }}>
          {mode === "light" ? <Moon size={13} /> : <Sun size={13} />} {mode === "light" ? "Modo oscuro" : "Modo día"}
        </button>
        <div className="text-xs mb-5 px-1" style={{ color: saveStatus === "error" ? T.alert : T.muted }}>
          {saveStatus === "saving" && "Guardando…"}
          {saveStatus === "saved" && "Guardado ✓"}
          {saveStatus === "error" && `Error al guardar: ${saveError || "desconocido"}`}
          {saveStatus === "idle" && "\u00A0"}
        </div>
        {NAV.map((n) => {
          const Icon = n.icon; const active = tab === n.id;
          return (
            <button key={n.id} onClick={() => setTab(n.id)} className="flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm text-left"
              style={{ background: active ? T.primary : "transparent", color: active ? "#fff" : T.text }}>
              <Icon size={16} />{n.label}
            </button>
          );
        })}
      </aside>

      <main className="flex-1 p-8 max-w-6xl">
        {(NAV.find((n) => n.id === tab)?.scoped || tab === "overview") && (
          <EventSwitcher T={T} events={events} activeEvent={activeEvent} setActiveEvent={setActiveEvent} />
        )}

        {tab === "overview" && (
          <Overview T={T} couple={couple} setCouple={setCouple} events={events} guests={guests} todos={todos} budget={budget}
            suppliers={suppliers} schedules={schedules} fileRef={fileRef} handlePhoto={handlePhoto} setTab={setTab}
            activeEvent={activeEvent} openTasks={openTasks} overdueTasks={overdueTasks} pendingRsvp={pendingRsvp}
            unpaidSuppliers={unpaidSuppliers} scheme={schemes.find((s) => s.id === activeScheme)} engagementDate={engagementDate} />
        )}
        {tab === "guests" && <Guests T={T} guests={guests} setGuests={setGuests} events={events} activeEvent={activeEvent} groups={groups} setGroups={setGroups} />}
        {tab === "ceremonies" && <Ceremonies T={T} events={events} setEvents={setEvents} etapas={etapas} setEtapas={setEtapas} guests={guests} budget={budget} todos={todos} setTodos={setTodos} todoTags={todoTags} schedules={schedules} setSchedules={setSchedules} />}
        {tab === "venues" && <Venues T={T} venues={venues} setVenues={setVenues} events={events} guests={guests} couple={couple} setSuppliers={setSuppliers} />}
        {tab === "todos" && <Todos T={T} todos={todos} setTodos={setTodos} events={events} tags={todoTags} setTags={setTodoTags} couple={couple} />}
        {tab === "budget" && <Budget T={T} budget={budget} setBudget={setBudget} activeEvent={activeEvent} events={events} cats={budgetCats} setCats={setBudgetCats} suppliers={suppliers} />}
        {tab === "suppliers" && <Suppliers T={T} suppliers={suppliers} setSuppliers={setSuppliers} cats={budgetCats} setCats={setBudgetCats} />}
        {tab === "catering" && <Catering T={T} catering={catering} setCatering={setCatering} activeEvent={activeEvent} events={events} />}
        {tab === "invitations" && <Invitations T={T} invitations={invitations} setInvitations={setInvitations} events={events} activeEvent={activeEvent} guests={guests} setGuests={setGuests} groups={groups} scheme={schemes.find((s) => s.id === activeScheme)} />}
        {tab === "seating" && <Seating T={T} seating={seating} setSeating={setSeating} activeEvent={activeEvent} events={events} guests={guests} />}
        {tab === "schedule" && <Schedule T={T} schedules={schedules} setSchedules={setSchedules} activeEvent={activeEvent} events={events} suppliers={suppliers} />}
        {tab === "scrapbook" && <Scrapbook T={T} items={scrapbook} setItems={setScrapbook} />}
        {tab === "inspiration" && <Inspiration T={T} boards={inspiration} setBoards={setInspiration} events={events} />}
        {tab === "schemes" && <Schemes T={T} schemes={schemes} setSchemes={setSchemes} activeScheme={activeScheme} setActiveScheme={setActiveScheme} mode={mode} />}
        {tab === "notifications" && (
          <Notifications T={T} guests={guests} todos={todos} suppliers={suppliers} events={events} reminders={reminders}
            setReminders={setReminders} dismissed={dismissed} setDismissed={setDismissed} notifPrefs={notifPrefs} />
        )}
        {tab === "settings" && (
          <Settings T={T} mode={mode} setMode={setMode} themeName={themeName} engagementDate={engagementDate}
            setEngagementDate={setEngagementDate} events={events} setEvents={setEvents} users={users} setUsers={setUsers}
            notifPrefs={notifPrefs} setNotifPrefs={setNotifPrefs} couple={couple} setCouple={setCouple}
            exportData={exportData} importData={importData} saveNow={saveNow} saveStatus={saveStatus} saveError={saveError} clearSavedData={clearSavedData} />
        )}
      </main>
    </div>
  );
}

/* ============================== shared primitives ============================== */
function SectionTitle({ T, children, action }) {
  return <div className="flex items-center justify-between mb-5"><h1 style={{ fontFamily: "Georgia, serif" }} className="text-3xl">{children}</h1>{action}</div>;
}
function Card({ T, children, style, onClick, draggable, onDragStart, onDragOver, onDrop }) {
  return <div onClick={onClick} draggable={draggable} onDragStart={onDragStart} onDragOver={onDragOver} onDrop={onDrop} className="rounded-2xl p-5" style={{ background: T.card, border: `1px solid ${T.line}`, cursor: onClick ? "pointer" : "default", ...style }}>{children}</div>;
}
function Button({ T, children, onClick, variant = "solid", small }) {
  const solid = variant === "solid";
  return (
    <button onClick={onClick} className={`flex items-center gap-1.5 rounded-full font-medium ${small ? "px-3 py-1.5 text-xs" : "px-4 py-2 text-sm"}`}
      style={solid ? { background: T.primary, color: "#fff" } : { background: "transparent", color: T.primaryDark, border: `1px solid ${T.soft}` }}>
      {children}
    </button>
  );
}
function Field({ T, label, children }) {
  return <label className="flex flex-col gap-1 text-sm"><span style={{ color: T.muted }}>{label}</span>{children}</label>;
}
function inputStyle(T) { return { background: T.mode === "dark" ? T.cardAlt : "#fff", border: `1px solid ${T.line}`, borderRadius: 10, padding: "8px 10px", fontSize: 14, color: T.text }; }
function Toggle({ T, checked, onChange }) {
  return (
    <button onClick={() => onChange(!checked)} className="w-11 h-6 rounded-full relative transition-colors" style={{ background: checked ? T.primary : T.line }}>
      <span className="absolute top-0.5 w-5 h-5 rounded-full bg-white transition-all" style={{ left: checked ? 22 : 2 }} />
    </button>
  );
}
function Bar({ T, pct, color }) {
  return <div className="w-full h-2 rounded-full overflow-hidden" style={{ background: T.line }}><div className="h-full rounded-full" style={{ width: `${Math.min(pct, 100)}%`, background: color || T.primary }} /></div>;
}
function EmptyState({ T, text }) {
  return <div className="text-sm py-10 text-center rounded-2xl" style={{ color: T.muted, border: `1px dashed ${T.line}` }}>{text}</div>;
}
function Modal({ T, title, onClose, children }) {
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-auto p-8" style={{ background: "rgba(0,0,0,0.5)" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} className="w-full max-w-2xl rounded-2xl p-6" style={{ background: T.card, marginTop: 20 }}>
        <div className="flex items-center justify-between mb-5">
          <h2 style={{ fontFamily: "Georgia, serif" }} className="text-xl">{title}</h2>
          <button onClick={onClose}><X size={18} color={T.muted} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}
function EventSwitcher({ T, events, activeEvent, setActiveEvent }) {
  return (
    <div className="flex gap-2 mb-6 flex-wrap">
      <button onClick={() => setActiveEvent("all")} className="px-4 py-2 rounded-full text-sm font-medium"
        style={activeEvent === "all" ? { background: T.secondary, color: "#fff" } : { background: T.cardAlt, color: T.muted, border: `1px solid ${T.line}` }}>Todos</button>
      {events.map((ev) => (
        <button key={ev.id} onClick={() => setActiveEvent(ev.id)} className="px-4 py-2 rounded-full text-sm font-medium"
          style={activeEvent === ev.id ? { background: T.secondary, color: "#fff" } : { background: T.cardAlt, color: T.muted, border: `1px solid ${T.line}` }}>
          {ev.name}{ev.date && <span className="ml-2 opacity-75">· {new Date(ev.date).toLocaleDateString("es-CL", { day: "numeric", month: "short" })}</span>}
        </button>
      ))}
    </div>
  );
}
// Foto con encuadre ajustable: cuando editing=true se arrastra para elegir qué parte se ve.
// La posición se guarda como "x% y%" (object-position) y solo se confirma al soltar.
function PositionedPhoto({ src, pos, onChange, editing, className = "", style }) {
  const [live, setLive] = useState(null);
  const drag = useRef(null);
  const [x, y] = String(live || pos || "50% 50%").split(" ").map((v) => parseFloat(v));
  function down(e) {
    if (!editing) return;
    e.preventDefault(); e.stopPropagation();
    e.currentTarget.setPointerCapture(e.pointerId);
    const r = e.currentTarget.getBoundingClientRect();
    drag.current = { sx: e.clientX, sy: e.clientY, x, y, w: r.width, h: r.height };
  }
  function move(e) {
    const d = drag.current;
    if (!d) return;
    const clamp = (n) => Math.max(0, Math.min(100, n));
    setLive(`${clamp(d.x - ((e.clientX - d.sx) / d.w) * 150).toFixed(1)}% ${clamp(d.y - ((e.clientY - d.sy) / d.h) * 150).toFixed(1)}%`);
  }
  function up() { if (drag.current && live) onChange?.(live); drag.current = null; setLive(null); }
  return (
    <img src={src} alt="" draggable={false} onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerCancel={up}
      className={`object-cover ${className}`} style={{ objectPosition: `${x}% ${y}%`, cursor: editing ? "grab" : undefined, touchAction: editing ? "none" : undefined, userSelect: "none", ...style }} />
  );
}
const photoSrc = (p) => (typeof p === "string" ? p : p?.src);
const photoPos = (p) => (typeof p === "string" ? undefined : p?.pos);

function pickEvent(activeEvent, events) { return activeEvent === "all" ? events[0]?.id : activeEvent; }

/* ============================== Overview ============================== */
function Overview({ T, couple, setCouple, events, guests, todos, budget, suppliers, schedules, fileRef, handlePhoto, setTab, activeEvent, openTasks, overdueTasks, pendingRsvp, unpaidSuppliers, scheme, engagementDate }) {
  const evFilter = activeEvent === "all" ? null : activeEvent;
  const scopedGuests = evFilter ? guests.filter((g) => g.eventIds.includes(evFilter)) : guests;
  const scopedBudget = evFilter ? budget.filter((b) => b.eventId === evFilter || b.eventId === "shared") : budget;
  const attending = scopedGuests.filter((g) => g.rsvp === "attending").length;
  const rsvpPct = scopedGuests.length ? Math.round((attending / scopedGuests.length) * 100) : 0;
  const spent = scopedBudget.reduce((s, b) => s + (b.mode === "qty" ? Number(b.qty || 0) * Number(b.unitCost || 0) : Number(b.amount || 0)), 0);
  const paidPct = scopedBudget.length ? Math.round((scopedBudget.filter((b) => b.paid).length / scopedBudget.length) * 100) : 0;
  const tasksPct = todos.length ? Math.round(((todos.length - openTasks) / todos.length) * 100) : 0;
  const bookedSuppliers = suppliers.filter((s) => s.paymentStatus === "Pagado").length;
  const supPct = suppliers.length ? Math.round((bookedSuppliers / suppliers.length) * 100) : 0;
  const nextEvent = events.filter((e) => e.date).sort((a, b) => a.date.localeCompare(b.date))[0];
  const dEngaged = daysUntil(engagementDate);
  const [framing, setFraming] = useState(false);

  const upcoming = Object.entries(schedules).flatMap(([evId, items]) => (items || []).map((it) => ({ ...it, evId })))
    .filter((it) => !evFilter || it.evId === evFilter).sort((a, b) => (a.start || "").localeCompare(b.start || "")).slice(0, 4);

  return (
    <div>
      <div className="relative rounded-3xl overflow-hidden mb-8 flex items-end" style={{ height: 260, background: `linear-gradient(135deg, ${T.soft}, ${T.cardAlt})` }}>
        {couple.photo && <PositionedPhoto src={couple.photo} pos={couple.photoPos} editing={framing} onChange={(p) => setCouple((c) => ({ ...c, photoPos: p }))} className="absolute inset-0 w-full h-full" style={{ zIndex: framing ? 20 : 0 }} />}
        <div className="absolute inset-0 pointer-events-none" style={{ background: couple.photo && !framing ? "linear-gradient(0deg, rgba(10,8,9,0.82), rgba(10,8,9,0.05))" : "none" }} />
        {framing && <div className="absolute top-4 left-1/2 -translate-x-1/2 z-30 text-xs px-3 py-1.5 rounded-full pointer-events-none" style={{ background: "rgba(0,0,0,0.6)", color: "#fff" }}>Arrastra la foto para encuadrarla</div>}
        <div className="relative z-30 p-6 flex items-end justify-between w-full pointer-events-none">
          <div>
            <div style={{ fontFamily: "Georgia, serif", color: couple.photo ? "#fff" : T.text }} className="text-4xl">{couple.a} & {couple.b}</div>
            <div className="text-sm mt-1" style={{ color: couple.photo ? "#F1E6D6" : T.muted }}>
              {engagementDate ? `Comprometidos hace ${Math.max(dEngaged * -1, 0)} días` : "Dos celebraciones, una sola historia"}
            </div>
          </div>
          <div className="flex gap-2 pointer-events-auto">
            {couple.photo && (
              <button onClick={() => setFraming((v) => !v)} className="flex items-center gap-2 px-4 py-2 rounded-full text-sm font-medium" style={{ background: framing ? T.primary : "rgba(255,255,255,0.92)", color: framing ? "#fff" : "#3B2A20" }}>
                {framing ? <><Check size={15} /> Listo</> : <><Pencil size={15} /> Ajustar encuadre</>}
              </button>
            )}
            <button onClick={() => fileRef.current?.click()} className="flex items-center gap-2 px-4 py-2 rounded-full text-sm font-medium" style={{ background: "rgba(255,255,255,0.92)", color: "#3B2A20" }}>
              <Camera size={15} /> {couple.photo ? "Cambiar foto" : "Subir foto"}
            </button>
          </div>
          <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={handlePhoto} />
        </div>
      </div>

      <div className="grid gap-4 mb-6" style={{ gridTemplateColumns: `repeat(${events.length}, 1fr)` }}>
        {events.map((ev) => {
          const d = daysUntil(ev.date);
          return (
            <Card T={T} key={ev.id} onClick={() => setTab("ceremonies")}>
              <div className="text-sm font-medium mb-2" style={{ color: T.primaryDark }}>{ev.name}</div>
              <div className="flex items-baseline gap-2">
                <span style={{ fontFamily: "Georgia, serif" }} className="text-4xl">{d !== null ? Math.max(d, 0) : "—"}</span>
                <span className="text-sm" style={{ color: T.muted }}>días</span>
              </div>
              <div className="text-xs mt-2" style={{ color: T.muted }}>{ev.type}{ev.time ? ` · ${ev.time}` : ""}</div>
            </Card>
          );
        })}
      </div>

      <div className="grid grid-cols-4 gap-4 mb-6">
        <StatCard T={T} label="Invitados" value={scopedGuests.length} sub={`${attending} confirmados`} onClick={() => setTab("guests")} />
        <StatCard T={T} label="Tareas pendientes" value={openTasks} sub={`${todos.length} en total`} onClick={() => setTab("todos")} />
        <StatCard T={T} label="Gastado" value={fmtMoney(spent, "CLP")} sub={`${suppliers.length} proveedores`} onClick={() => setTab("budget")} />
        <StatCard T={T} label="Proveedores" value={suppliers.length} sub={`${bookedSuppliers} pagados`} onClick={() => setTab("suppliers")} />
      </div>

      <div className="grid grid-cols-3 gap-4 mb-6">
        <Card T={T} style={{ gridColumn: "span 2" }}>
          <div className="text-sm font-medium mb-3">Progreso</div>
          <div className="flex flex-col gap-3">
            <ProgressRow T={T} label="RSVP confirmados" pct={rsvpPct} color={T.secondary} />
            <ProgressRow T={T} label="Ítems de presupuesto pagados" pct={paidPct} color={T.primary} />
            <ProgressRow T={T} label="Tareas completadas" pct={tasksPct} color={T.soft} />
            <ProgressRow T={T} label="Proveedores pagados" pct={supPct} color={T.alert} />
          </div>
        </Card>
        <Card T={T}>
          <div className="text-sm font-medium mb-3">RSVP tracker</div>
          <div className="flex items-center justify-center mb-3">
            <div className="w-24 h-24 rounded-full flex items-center justify-center" style={{ background: `conic-gradient(${T.secondary} ${rsvpPct * 3.6}deg, ${T.line} 0deg)` }}>
              <div className="w-16 h-16 rounded-full flex items-center justify-center text-sm font-medium" style={{ background: T.card }}>{rsvpPct}%</div>
            </div>
          </div>
          <div className="text-xs flex justify-between" style={{ color: T.muted }}><span>Confirman</span><span>{attending}</span></div>
          <div className="text-xs flex justify-between" style={{ color: T.muted }}><span>Esperando</span><span>{scopedGuests.filter((g) => g.rsvp === "pending").length}</span></div>
        </Card>
      </div>

      <div className="grid grid-cols-3 gap-4 mb-6">
        <Card T={T} onClick={() => setTab("notifications")} style={{ borderColor: (overdueTasks + pendingRsvp + unpaidSuppliers) ? T.alert : T.line }}>
          <div className="flex items-center gap-2 text-sm font-medium mb-3"><AlertCircle size={15} color={T.alert} /> Necesita atención</div>
          <div className="flex flex-col gap-2 text-xs" style={{ color: T.muted }}>
            {pendingRsvp > 0 && <div>{pendingRsvp} invitados sin responder</div>}
            {overdueTasks > 0 && <div>{overdueTasks} tareas atrasadas</div>}
            {unpaidSuppliers > 0 && <div>{unpaidSuppliers} proveedores sin pagar</div>}
            {pendingRsvp + overdueTasks + unpaidSuppliers === 0 && <div>Todo al día ✓</div>}
          </div>
        </Card>
        <Card T={T} onClick={() => setTab("schedule")}>
          <div className="flex items-center gap-2 text-sm font-medium mb-3"><CalendarClock size={15} color={T.primary} /> Lo que viene</div>
          <div className="flex flex-col gap-2 text-xs">
            {upcoming.length === 0 && <span style={{ color: T.muted }}>Nada programado aún</span>}
            {upcoming.map((it) => <div key={it.id} className="flex justify-between"><span>{it.title}</span><span style={{ color: T.muted }}>{it.start}</span></div>)}
          </div>
        </Card>
        <Card T={T} onClick={() => setTab("schemes")}>
          <div className="flex items-center gap-2 text-sm font-medium mb-3"><Palette size={15} color={T.secondary} /> Tu paleta</div>
          <div className="text-xs mb-2" style={{ color: T.muted }}>{scheme?.name}</div>
          <div className="flex gap-1.5">{(scheme?.colors || []).map((c, i) => <div key={i} className="w-8 h-8 rounded-lg" style={{ background: c }} />)}</div>
        </Card>
      </div>

      <div className="flex gap-3">
        <button onClick={() => setTab("schedule")} className="flex-1 flex items-center justify-between px-5 py-3 rounded-xl text-sm font-medium" style={{ background: T.cardAlt, color: T.primaryDark }}>Ir a Cronograma <ChevronRight size={15} /></button>
        <button onClick={() => setTab("seating")} className="flex-1 flex items-center justify-between px-5 py-3 rounded-xl text-sm font-medium" style={{ background: T.cardAlt, color: T.primaryDark }}>Ir a Mesas <ChevronRight size={15} /></button>
        <button onClick={() => setTab("ceremonies")} className="flex-1 flex items-center justify-between px-5 py-3 rounded-xl text-sm font-medium" style={{ background: T.cardAlt, color: T.primaryDark }}>Ir a Ceremonias <ChevronRight size={15} /></button>
      </div>
    </div>
  );
}
function StatCard({ T, label, value, sub, onClick }) {
  return <Card T={T} onClick={onClick}><div className="text-xs" style={{ color: T.muted }}>{label}</div><div style={{ fontFamily: "Georgia, serif" }} className="text-2xl mt-1">{value}</div><div className="text-xs mt-1" style={{ color: T.secondary }}>{sub}</div></Card>;
}
function ProgressRow({ T, label, pct, color }) {
  return <div><div className="flex justify-between text-xs mb-1"><span style={{ color: T.muted }}>{label}</span><span>{pct}%</span></div><Bar T={T} pct={pct} color={color} /></div>;
}

/* ============================== Eventos + Etapas ============================== */
// Un "evento" es lo que antes llamábamos ceremonia (ej. Matrimonio Civil):
// es el nivel al que se asignan invitados, tareas, presupuesto, catering,
// invitaciones, mesas y cronograma. Una "etapa" es una parte del día dentro
// de ese evento (Ceremonia, Recepción, After, etc.) — puramente descriptiva.
function Ceremonies({ T, events, setEvents, etapas, setEtapas, guests, budget, todos, setTodos, todoTags, schedules, setSchedules }) {
  const [form, setForm] = useState(null);
  const [etapaForm, setEtapaForm] = useState({}); // { [eventId]: {name,type,time} }

  function startNew() { setForm({ id: uid(), name: "", date: "", time: "", address: "" }); }
  function save() {
    if (!form.name.trim()) return;
    setEvents((ev) => (ev.some((x) => x.id === form.id) ? ev.map((x) => (x.id === form.id ? form : x)) : [...ev, form]));
    setForm(null);
  }
  function remove(id) { setEvents((ev) => ev.filter((x) => x.id !== id)); setEtapas((et) => et.filter((x) => x.eventId !== id)); }
  function importEvents(rows) {
    const created = rows.map((row) => ({ id: uid(), name: pick(row, ["Nombre", "Name"]), date: pick(row, ["Fecha", "Date"]), time: pick(row, ["Hora", "Time"]), address: pick(row, ["Direccion", "Dirección", "Address"]) })).filter((e) => e.name);
    if (created.length) setEvents((ev) => [...ev, ...created]);
  }

  function addEtapa(eventId) {
    const draft = etapaForm[eventId] || {};
    if (!draft.name?.trim()) return;
    setEtapas((et) => [...et, { id: uid(), eventId, name: draft.name, type: draft.type || EVENT_TYPES[0], time: draft.time || "" }]);
    setEtapaForm((f) => ({ ...f, [eventId]: { name: "", type: EVENT_TYPES[0], time: "" } }));
  }
  function removeEtapa(id) { setEtapas((et) => et.filter((x) => x.id !== id)); }

  function addTemplateTasks(eventId) {
    const existingTitles = new Set(todos.filter((t) => t.scope === eventId).map((t) => t.title));
    const toAdd = TASK_TEMPLATE.filter((t) => !existingTitles.has(t.title))
      .map((t) => ({ id: uid(), title: t.title, tag: todoTags.includes(t.tag) ? t.tag : todoTags[0], checklist: [], assignedTo: "Ambos", due: "", scope: eventId, done: false }));
    if (toAdd.length) setTodos((ts) => [...ts, ...toAdd]);
  }

  function generateSchedule(eventId, evEtapas) {
    const existing = new Set((schedules[eventId] || []).map((s) => `${s.title}__${s.start}`));
    const toAdd = evEtapas.filter((et) => !existing.has(`${et.name}__${et.time}`))
      .map((et) => ({ id: uid(), title: et.name, start: et.time, end: "", runsUntilEnd: false, where: "", withSupplier: "", remindMe: false, notes: `Generado desde la etapa "${et.type}"` }));
    if (toAdd.length) setSchedules((s) => ({ ...s, [eventId]: [...(s[eventId] || []), ...toAdd] }));
  }

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="flex gap-2">
          <ImportExcel T={T} label="Importar eventos" hint="Columnas: Nombre, Fecha, Hora, Direccion" onRows={importEvents} />
          <Button T={T} onClick={startNew}><Plus size={15} /> Agregar evento</Button>
        </div>}>
        Eventos
      </SectionTitle>
      <p className="text-xs mb-4" style={{ color: T.muted }}>Cada evento (ej. Matrimonio Civil, Matrimonio Religioso) es donde se asignan invitados, presupuesto, catering, invitaciones, mesas y cronograma. Dentro de cada evento puedes agregar etapas (Ceremonia, Recepción, After, etc.).</p>
      <div className="flex flex-col gap-4">
        {events.map((ev) => {
          const g = guests.filter((x) => x.eventIds.includes(ev.id));
          const attending = g.filter((x) => x.rsvp === "attending").length;
          const b = budget.filter((x) => x.eventId === ev.id);
          const spent = b.reduce((s, x) => s + (x.mode === "qty" ? Number(x.qty || 0) * Number(x.unitCost || 0) : Number(x.amount || 0)), 0);
          const paidPct = b.length ? Math.round((b.filter((x) => x.paid).length / b.length) * 100) : 0;
          const d = daysUntil(ev.date);
          const evEtapas = etapas.filter((x) => x.eventId === ev.id).slice().sort((a, b2) => (a.time || "").localeCompare(b2.time || ""));
          const draft = etapaForm[ev.id] || { name: "", type: EVENT_TYPES[0], time: "" };

          return (
            <Card T={T} key={ev.id}>
              <div className="flex items-start justify-between mb-3">
                <div>
                  <div style={{ fontFamily: "Georgia, serif" }} className="text-2xl">{ev.name}</div>
                  <div className="text-xs mt-1" style={{ color: T.muted }}>
                    {ev.date ? new Date(ev.date).toLocaleDateString("es-CL", { day: "numeric", month: "long", year: "numeric" }) : "Sin fecha"}
                    {ev.time ? ` · ${ev.time} hrs` : ""}{d !== null ? ` · faltan ${Math.max(d, 0)} días` : ""}
                  </div>
                </div>
                <div className="flex gap-2">
                  <Button T={T} small variant="outline" onClick={() => addTemplateTasks(ev.id)}>+ Tareas típicas</Button>
                  <Button T={T} small variant="outline" onClick={() => setForm(ev)}>Editar</Button>
                  {events.length > 1 && <button onClick={() => remove(ev.id)}><Trash2 size={14} color={T.muted} /></button>}
                </div>
              </div>

              <div className="grid grid-cols-3 gap-4 mb-4">
                <div><div className="text-xs mb-1" style={{ color: T.muted }}>Invitados ({attending}/{g.length})</div><Bar T={T} pct={g.length ? (attending / g.length) * 100 : 0} color={T.secondary} /></div>
                <div><div className="text-xs mb-1" style={{ color: T.muted }}>Presupuesto pagado</div><Bar T={T} pct={paidPct} color={T.primary} /></div>
                <div><div className="text-xs mb-1" style={{ color: T.muted }}>Gastado</div><div className="text-sm font-medium">{fmtMoney(spent, "CLP")}</div></div>
              </div>

              {ev.address && (
                <div className="mb-4">
                  <div className="rounded-xl overflow-hidden mb-2" style={{ border: `1px solid ${T.line}` }}>
                    <iframe title={ev.name} src={mapsEmbed(ev.address)} width="100%" height="200" style={{ border: 0, display: "block" }} loading="lazy" />
                  </div>
                  <a href={mapsLink(ev.address)} target="_blank" rel="noreferrer" className="text-xs flex items-center gap-1" style={{ color: T.primaryDark }}><MapPin size={12} /> {ev.address} <ExternalLink size={11} /></a>
                </div>
              )}

              <div style={{ borderTop: `1px solid ${T.line}`, paddingTop: 14 }}>
                <div className="flex items-center justify-between mb-2">
                  <div className="text-xs font-medium" style={{ color: T.muted }}>Etapas de este evento</div>
                  {evEtapas.length > 0 && <Button T={T} small variant="outline" onClick={() => generateSchedule(ev.id, evEtapas)}>Generar en cronograma</Button>}
                </div>
                <div className="flex flex-wrap gap-2 mb-3">
                  {evEtapas.length === 0 && <span className="text-xs" style={{ color: T.muted }}>Sin etapas aún</span>}
                  {evEtapas.map((et) => (
                    <span key={et.id} className="text-xs px-2.5 py-1.5 rounded-full flex items-center gap-1.5" style={{ background: T.cardAlt }}>
                      {et.name}{et.time ? ` · ${et.time}` : ""} <span style={{ color: T.muted }}>· {et.type}</span>
                      <button onClick={() => removeEtapa(et.id)}><X size={11} color={T.muted} /></button>
                    </span>
                  ))}
                </div>
                <div className="flex gap-2 flex-wrap items-end">
                  <input placeholder="Nombre de la etapa" style={{ ...inputStyle(T), width: 180 }} value={draft.name} onChange={(e) => setEtapaForm((f) => ({ ...f, [ev.id]: { ...draft, name: e.target.value } }))} />
                  <select style={inputStyle(T)} value={draft.type} onChange={(e) => setEtapaForm((f) => ({ ...f, [ev.id]: { ...draft, type: e.target.value } }))}>{EVENT_TYPES.map((t) => <option key={t}>{t}</option>)}</select>
                  <input type="time" style={inputStyle(T)} value={draft.time} onChange={(e) => setEtapaForm((f) => ({ ...f, [ev.id]: { ...draft, time: e.target.value } }))} />
                  <Button T={T} small variant="outline" onClick={() => addEtapa(ev.id)}><Plus size={13} /> Agregar etapa</Button>
                </div>
                <p className="text-xs mt-2" style={{ color: T.muted }}>"Generar en cronograma" crea un ítem base por cada etapa (título y hora) sin pisar lo que ya hayas cargado a mano en Cronograma.</p>
              </div>
            </Card>
          );
        })}
      </div>

      {form && (
        <Modal T={T} title={events.some((x) => x.id === form.id) ? "Editar evento" : "Nuevo evento"} onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2"><Field T={T} label="Nombre"><input style={{ ...inputStyle(T), width: "100%" }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} autoFocus placeholder="Ej. Matrimonio Civil" /></Field></div>
            <Field T={T} label="Fecha"><input type="date" style={inputStyle(T)} value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} /></Field>
            <Field T={T} label="Horario"><input type="time" style={inputStyle(T)} value={form.time} onChange={(e) => setForm({ ...form, time: e.target.value })} /></Field>
            <div className="col-span-2"><Field T={T} label="Dirección / lugar"><input style={{ ...inputStyle(T), width: "100%" }} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} placeholder="Ej. Parque Bicentenario, Vitacura" /></Field></div>
          </div>
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== Guests ============================== */
// Lee un archivo .vcf (contactos exportados desde iPhone, Mac, Android o Google) y devuelve [{ name, email, phone }].
function parseVCard(text) {
  const unfolded = text.replace(/\r\n/g, "\n").replace(/\n[ \t]/g, "");
  const cards = unfolded.split(/BEGIN:VCARD/i).slice(1);
  const clean = (v) => v.replace(/\\,/g, ",").replace(/\;/g, ";").replace(/\\n/gi, " ").trim();
  return cards.map((card) => {
    let fn = "", n = "", email = "", phone = "";
    for (const line of card.split("\n")) {
      const idx = line.indexOf(":");
      if (idx < 0) continue;
      const key = line.slice(0, idx).replace(/^item\d+\./i, "").split(";")[0].toUpperCase();
      const val = line.slice(idx + 1);
      if (key === "FN" && !fn) fn = clean(val);
      else if (key === "N" && !n) { const [last, first] = val.split(";"); n = clean(`${first || ""} ${last || ""}`); }
      else if (key === "EMAIL" && !email) email = clean(val);
      else if (key === "TEL" && !phone) phone = clean(val);
    }
    return { name: fn || n, email, phone };
  }).filter((c) => c.name);
}

function Guests({ T, guests, setGuests, events, activeEvent, groups, setGroups }) {
  const [form, setForm] = useState(null);
  const [newGroup, setNewGroup] = useState("");
  const [manageGroups, setManageGroups] = useState(false);
  const [contacts, setContacts] = useState(null); // { list, selected: Set, group, query }
  const dragRef = useRef(null);
  const vcfRef = useRef(null);
  const visible = activeEvent === "all" ? guests : guests.filter((g) => g.eventIds.includes(activeEvent));
  const S = RSVP_STYLES(T);
  const canPickContacts = typeof navigator !== "undefined" && "contacts" in navigator && "select" in (navigator.contacts || {});

  function startNew() { setForm({ id: uid(), name: "", eventIds: activeEvent === "all" ? [] : [activeEvent], rsvp: "pending", ageGroup: "Adulto", group: groups[0] || "", inviteSent: false, restrictions: "", notes: "", phone: "", email: "" }); }
  function save() { if (!form.name.trim()) return; setGuests((g) => (g.some((x) => x.id === form.id) ? g.map((x) => (x.id === form.id ? form : x)) : [...g, form])); setForm(null); }
  function importGuests(rows) {
    const mapRsvp = (v) => { const s = String(v || "").toLowerCase(); if (/asiste|confirm|s[ií]/.test(s)) return "attending"; if (/no asiste|declin/.test(s)) return "declined"; return "pending"; };
    const created = rows.map((row) => {
      // También acepta el CSV que exporta Google Contactos (First Name / Last Name / E-mail 1 - Value / Phone 1 - Value).
      const name = pick(row, ["Nombre", "Name"]) || [pick(row, ["First Name", "Given Name"]), pick(row, ["Last Name", "Family Name"])].filter(Boolean).join(" ");
      if (!name) return null;
      const evIds = matchEvents(pick(row, ["Evento", "Eventos", "Event"]), events);
      return {
        id: uid(), name, group: pick(row, ["Grupo", "Group"]) || groups[0] || "",
        ageGroup: pick(row, ["GrupoEdad", "Edad", "AgeGroup"]) || "Adulto",
        rsvp: mapRsvp(pick(row, ["RSVP", "Rsvp"])),
        inviteSent: truthy(pick(row, ["InvitacionEnviada", "Invitación enviada", "InviteSent"])),
        restrictions: pick(row, ["Restricciones", "Restrictions"]),
        notes: pick(row, ["Notas", "Notes"]),
        phone: pick(row, ["Telefono", "Teléfono", "Phone", "Phone 1 - Value"]),
        email: pick(row, ["Email", "Correo", "E-mail", "E-mail 1 - Value"]),
        eventIds: evIds.length ? evIds : (activeEvent !== "all" ? [activeEvent] : []),
      };
    }).filter(Boolean);
    if (created.length) setGuests((gs) => [...gs, ...created]);
  }
  function openContactReview(list) {
    const known = new Set(guests.map((g) => g.name.trim().toLowerCase()));
    const fresh = list.filter((c) => !known.has(c.name.trim().toLowerCase()));
    if (!fresh.length) { alert(list.length ? "Todos esos contactos ya están en la lista de invitados." : "No se encontraron contactos."); return; }
    setContacts({ list: fresh, selected: new Set(fresh.length <= 30 ? fresh.map((_, i) => i) : []), group: groups[0] || "", query: "" });
  }
  async function pickFromPhone() {
    try {
      const res = await navigator.contacts.select(["name", "email", "tel"], { multiple: true });
      openContactReview(res.map((c) => ({ name: c.name?.[0] || "", email: c.email?.[0] || "", phone: c.tel?.[0] || "" })).filter((c) => c.name));
    } catch (e) { /* el usuario canceló */ }
  }
  async function handleVcf(e) {
    const f = e.target.files?.[0];
    e.target.value = "";
    if (!f) return;
    openContactReview(parseVCard(await f.text()));
  }
  function addSelectedContacts() {
    const chosen = contacts.list.filter((_, i) => contacts.selected.has(i));
    const created = chosen.map((c) => ({ id: uid(), name: c.name, email: c.email, phone: c.phone, group: contacts.group, eventIds: activeEvent === "all" ? [] : [activeEvent], rsvp: "pending", ageGroup: "Adulto", inviteSent: false, restrictions: "", notes: "" }));
    setGuests((gs) => [...gs, ...created]);
    setContacts(null);
  }
  function addGroup() { const g = newGroup.trim(); if (g && !groups.includes(g)) setGroups((gs) => [...gs, g]); setNewGroup(""); }
  function renameGroup(oldName, name) {
    const n = name.trim();
    if (!n || n === oldName || groups.includes(n)) return;
    setGroups((gs) => gs.map((g) => (g === oldName ? n : g)));
    setGuests((gs) => gs.map((g) => (g.group === oldName ? { ...g, group: n } : g)));
  }
  function deleteGroup(name) {
    const count = guests.filter((g) => g.group === name).length;
    if (!window.confirm(count ? `¿Borrar el grupo "${name}"? Sus ${count} invitado(s) quedarán "Sin grupo" (no se borran).` : `¿Borrar el grupo "${name}"?`)) return;
    setGroups((gs) => gs.filter((g) => g !== name));
    setGuests((gs) => gs.map((g) => (g.group === name ? { ...g, group: "" } : g)));
  }
  function printGuests() {
    const rows = visible.map((g) => `<tr><td>${g.name}</td><td>${g.group || "Sin grupo"}</td><td>${g.ageGroup}</td><td>${S[g.rsvp].label}</td><td>${g.inviteSent ? "Sí" : "No"}</td><td>${g.phone || ""}</td><td>${g.email || ""}</td><td>${g.restrictions || ""}</td></tr>`).join("");
    printHtml("Lista de invitados", `<h1>Lista de invitados</h1><table><tr><th>Nombre</th><th>Grupo</th><th>Edad</th><th>RSVP</th><th>Invitación</th><th>Teléfono</th><th>Email</th><th>Restricciones</th></tr>${rows}</table>`);
  }

  const shownContacts = contacts ? contacts.list.map((c, i) => ({ ...c, i })).filter((c) => !contacts.query || `${c.name} ${c.email} ${c.phone}`.toLowerCase().includes(contacts.query.toLowerCase())) : [];

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="flex gap-2 flex-wrap justify-end">
          <Button T={T} small variant="outline" onClick={printGuests}><Printer size={13} /> Imprimir / PDF</Button>
          <Button T={T} small variant="outline" onClick={() => setManageGroups(true)}><Users size={13} /> Grupos</Button>
          {canPickContacts && <Button T={T} small variant="outline" onClick={pickFromPhone}><Contact size={13} /> Desde contactos del cel</Button>}
          <Button T={T} small variant="outline" onClick={() => vcfRef.current?.click()}><Contact size={13} /> Importar contactos (.vcf)</Button>
          <input ref={vcfRef} type="file" accept=".vcf,text/vcard,text/x-vcard" className="hidden" onChange={handleVcf} />
          <ImportExcel T={T} label="Importar Excel" hint="Columnas: Nombre, Grupo, GrupoEdad, RSVP, InvitacionEnviada, Restricciones, Notas, Telefono, Email, Evento. También acepta el CSV de Google Contactos." onRows={importGuests} />
          <Button T={T} onClick={startNew}><Plus size={15} /> Agregar invitado</Button>
        </div>}>
        Invitados
      </SectionTitle>
      <p className="text-xs mb-4" style={{ color: T.muted }}>Cada invitado es único y puede asignarse a uno o varios eventos — usa el filtro arriba para verlos todos o por separado. Para traer contactos: en el iPhone abre Contactos → un contacto → Compartir contacto → Guardar en Archivos, y súbelo con "Importar contactos (.vcf)". En el Mac: app Contactos → selecciona varios → Archivo → Exportar → vCard.</p>
      <div className="flex flex-col gap-2">
        {visible.length === 0 && <EmptyState T={T} text="No hay invitados para este filtro todavía." />}
        {visible.map((g) => {
          const st = S[g.rsvp];
          return (
            <Card T={T} key={g.id} style={{ padding: "12px 16px" }} onClick={() => setForm({ email: "", ...g })} draggable onDragStart={(e) => { e.stopPropagation(); dragRef.current = g.id; }} onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.stopPropagation(); setGuests((gs) => reorderById(gs, dragRef.current, g.id)); }}>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <GripVertical size={13} color={T.muted} style={{ cursor: "grab" }} />
                  <div>
                    <div className="text-sm font-medium">{g.name}</div>
                    <div className="text-xs mt-0.5 flex items-center gap-1.5 flex-wrap" style={{ color: T.muted }}>
                      {g.group || "Sin grupo"} · {g.ageGroup} · {g.inviteSent ? "Invitación enviada" : "Sin enviar"}
                      {g.email && <span className="flex items-center gap-1"><Mail size={11} />{g.email}</span>}
                      {g.phone && <span className="flex items-center gap-1"><Phone size={11} />{g.phone}</span>}
                      {g.eventIds.map((id) => <span key={id} className="px-1.5 py-0.5 rounded" style={{ background: T.cardAlt }}>{events.find((e) => e.id === id)?.name}</span>)}
                    </div>
                  </div>
                </div>
                <div className="flex items-center gap-3">
                  <span className="text-xs px-2.5 py-1 rounded-full" style={{ background: st.bg, color: st.fg }}>{st.label}</span>
                  <button onClick={(e) => { e.stopPropagation(); setGuests((gs) => gs.filter((x) => x.id !== g.id)); }}><Trash2 size={14} color={T.muted} /></button>
                </div>
              </div>
            </Card>
          );
        })}
      </div>

      {form && (
        <Modal T={T} title={form.name ? "Editar invitado" : "Nuevo invitado"} onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <Field T={T} label="Nombre"><input style={inputStyle(T)} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} autoFocus /></Field>
            <Field T={T} label="Grupo de edad"><select style={inputStyle(T)} value={form.ageGroup} onChange={(e) => setForm({ ...form, ageGroup: e.target.value })}><option>Adulto</option><option>Niño</option></select></Field>
            <Field T={T} label="Grupo asignado"><select style={inputStyle(T)} value={form.group} onChange={(e) => setForm({ ...form, group: e.target.value })}><option value="">Sin grupo</option>{groups.map((gr) => <option key={gr}>{gr}</option>)}</select></Field>
            <Field T={T} label="RSVP"><select style={inputStyle(T)} value={form.rsvp} onChange={(e) => setForm({ ...form, rsvp: e.target.value })}><option value="pending">Sin responder</option><option value="attending">Confirma</option><option value="declined">No asiste</option></select></Field>
            <Field T={T} label="Teléfono"><input type="tel" style={inputStyle(T)} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
            <Field T={T} label="Email"><input type="email" style={inputStyle(T)} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
            <div className="col-span-2 flex items-center gap-3"><Toggle T={T} checked={form.inviteSent} onChange={(v) => setForm({ ...form, inviteSent: v })} /><span className="text-sm">Invitación enviada</span></div>
            <div className="col-span-2">
              <Field T={T} label="Asignado a evento(s) — puede ser uno o varios">
                <div className="flex gap-2 mt-1">
                  {events.map((ev) => (
                    <button key={ev.id} onClick={() => setForm((f) => ({ ...f, eventIds: f.eventIds.includes(ev.id) ? f.eventIds.filter((id) => id !== ev.id) : [...f.eventIds, ev.id] }))}
                      className="px-3 py-2 rounded-lg text-xs" style={form.eventIds.includes(ev.id) ? { background: T.secondary, color: "#fff" } : { background: T.cardAlt, border: `1px solid ${T.line}`, color: T.muted }}>
                      {ev.name}
                    </button>
                  ))}
                </div>
              </Field>
            </div>
            <div className="col-span-2"><Field T={T} label="Restricciones alimentarias / accesibilidad"><input style={inputStyle(T)} value={form.restrictions} onChange={(e) => setForm({ ...form, restrictions: e.target.value })} /></Field></div>
            <div className="col-span-2"><Field T={T} label="Notas"><textarea style={{ ...inputStyle(T), width: "100%", minHeight: 60 }} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></Field></div>
            <div className="col-span-2 flex gap-2 items-center pt-2" style={{ borderTop: `1px solid ${T.line}` }}>
              <input placeholder="Nuevo grupo (ej. Colegio)" style={{ ...inputStyle(T), flex: 1 }} value={newGroup} onChange={(e) => setNewGroup(e.target.value)} />
              <Button T={T} small variant="outline" onClick={addGroup}><Plus size={13} /> Agregar grupo</Button>
            </div>
          </div>
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}

      {manageGroups && (
        <Modal T={T} title="Grupos de invitados" onClose={() => setManageGroups(false)}>
          <p className="text-xs mb-3" style={{ color: T.muted }}>Cambia el nombre directo en el campo. Al borrar un grupo, sus invitados no se borran: quedan "Sin grupo".</p>
          <div className="flex flex-col gap-2">
            {groups.length === 0 && <EmptyState T={T} text="No hay grupos." />}
            {groups.map((gr) => (
              <div key={gr} className="flex items-center gap-2">
                <input defaultValue={gr} style={{ ...inputStyle(T), flex: 1 }} onBlur={(e) => { const v = e.target.value.trim(); if (!v || (v !== gr && groups.includes(v))) e.target.value = gr; else renameGroup(gr, v); }} onKeyDown={(e) => e.key === "Enter" && e.currentTarget.blur()} />
                <span className="text-xs w-20 text-right" style={{ color: T.muted }}>{guests.filter((g) => g.group === gr).length} invitados</span>
                <button onClick={() => deleteGroup(gr)} title="Borrar grupo"><Trash2 size={15} color={T.alert} /></button>
              </div>
            ))}
            <div className="flex gap-2 items-center pt-3 mt-1" style={{ borderTop: `1px solid ${T.line}` }}>
              <input placeholder="Nuevo grupo (ej. Colegio)" style={{ ...inputStyle(T), flex: 1 }} value={newGroup} onChange={(e) => setNewGroup(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addGroup()} />
              <Button T={T} small variant="outline" onClick={addGroup}><Plus size={13} /> Agregar</Button>
            </div>
          </div>
        </Modal>
      )}

      {contacts && (
        <Modal T={T} title={`Importar contactos (${contacts.list.length})`} onClose={() => setContacts(null)}>
          <div className="flex gap-2 mb-3 flex-wrap items-center">
            <input placeholder="Buscar…" style={{ ...inputStyle(T), flex: 1, minWidth: 140 }} value={contacts.query} onChange={(e) => setContacts({ ...contacts, query: e.target.value })} />
            <select style={inputStyle(T)} value={contacts.group} onChange={(e) => setContacts({ ...contacts, group: e.target.value })}><option value="">Sin grupo</option>{groups.map((gr) => <option key={gr}>{gr}</option>)}</select>
          </div>
          <div className="flex gap-3 text-xs mb-2" style={{ color: T.primaryDark }}>
            <button onClick={() => setContacts({ ...contacts, selected: new Set([...contacts.selected, ...shownContacts.map((c) => c.i)]) })}>Marcar visibles</button>
            <button onClick={() => setContacts({ ...contacts, selected: new Set() })}>Desmarcar todos</button>
          </div>
          <div className="flex flex-col gap-1 overflow-auto" style={{ maxHeight: 360 }}>
            {shownContacts.map((c) => (
              <label key={c.i} className="flex items-center gap-3 px-2 py-1.5 rounded-lg text-sm cursor-pointer" style={{ background: contacts.selected.has(c.i) ? T.cardAlt : "transparent" }}>
                <input type="checkbox" checked={contacts.selected.has(c.i)} onChange={() => { const s = new Set(contacts.selected); s.has(c.i) ? s.delete(c.i) : s.add(c.i); setContacts({ ...contacts, selected: s }); }} />
                <span className="flex-1">{c.name}</span>
                <span className="text-xs" style={{ color: T.muted }}>{c.email || c.phone}</span>
              </label>
            ))}
          </div>
          <p className="text-xs mt-2" style={{ color: T.muted }}>Los contactos que ya están como invitados (mismo nombre) no aparecen.</p>
          <div className="flex gap-2 mt-4 justify-end">
            <Button T={T} variant="outline" onClick={() => setContacts(null)}>Cancelar</Button>
            <Button T={T} onClick={addSelectedContacts}><Check size={14} /> Agregar {contacts.selected.size} invitado(s)</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== Lugares / centros de eventos ============================== */
const VENUE_TYPES = ["Centro de eventos", "Hacienda / Fundo", "Viña", "Salón", "Hotel", "Casa / Jardín", "Restaurante", "Playa", "Otro"];
const VENUE_STATUS = ["Por visitar", "Visitado", "Favorito", "Descartado", "Reservado"];
const VENUE_VIBES = ["Rústico", "Campestre", "Elegante", "Boho", "Moderno", "Clásico", "Industrial", "Minimalista", "Romántico", "Playero", "Jardín / Bosque"];
const VENUE_SERVICES = ["Banquetera / Catering", "DJ / Música", "Barra / Bebidas", "Decoración", "Flores", "Mobiliario", "Iluminación", "Coordinador(a) del día", "Fotografía", "Torta", "Estacionamiento", "Alojamiento", "Ceremonia en el lugar"];
const VENUE_ASPECTS = ["Belleza del lugar", "Vibe / decoración", "Ubicación y acceso", "Capacidad y espacios", "Relación precio-calidad", "Comida / banquetera", "Atención y flexibilidad", "Plan lluvia / clima", "Estacionamiento y logística", "Fotos (luz y rincones)"];
const MAX_VENUE_PHOTOS = 8;

function venueAvg(v, who) {
  const vals = VENUE_ASPECTS.map((a) => Number(v.ratings?.[a]?.[who])).filter((n) => n > 0);
  return vals.length ? vals.reduce((s, n) => s + n, 0) / vals.length : null;
}
function venueCombined(v) {
  const vals = [venueAvg(v, "a"), venueAvg(v, "b")].filter((n) => n !== null);
  return vals.length ? vals.reduce((s, n) => s + n, 0) / vals.length : null;
}
function venueTotal(v, headcount) {
  const base = Number(v.priceBase || 0), pp = Number(v.pricePerPerson || 0);
  return base || pp ? base + pp * headcount : null;
}
const fmtScore = (n) => (n === null || n === undefined ? "—" : n.toFixed(1));

function ScoreSelect({ T, value, onChange }) {
  return (
    <select style={{ ...inputStyle(T), padding: "4px 6px", width: 64 }} value={value || ""} onChange={(e) => onChange(e.target.value ? Number(e.target.value) : "")}>
      <option value="">—</option>
      {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((n) => <option key={n} value={n}>{n}</option>)}
    </select>
  );
}
function Chip({ T, active, children, onClick }) {
  return (
    <button onClick={onClick} className="px-3 py-1.5 rounded-full text-xs"
      style={active ? { background: T.secondary, color: "#fff" } : { background: T.cardAlt, color: T.muted, border: `1px solid ${T.line}` }}>{children}</button>
  );
}

function Venues({ T, venues, setVenues, events, guests, couple, setSuppliers }) {
  const [form, setForm] = useState(null);
  const [statusFilter, setStatusFilter] = useState("Todos");
  const [sortBy, setSortBy] = useState("score");
  const [compareIds, setCompareIds] = useState([]);
  const [comparing, setComparing] = useState(false);
  const defaultHeadcount = guests.filter((g) => g.rsvp !== "declined").length || 100;
  const [headcount, setHeadcount] = useState(defaultHeadcount);
  const photoRef = useRef(null);
  const nameA = couple.a || "Novia", nameB = couple.b || "Novio";

  function startNew() {
    setForm({ id: uid(), name: "", type: VENUE_TYPES[0], status: "Por visitar", vibes: [], eventId: "", address: "", contact: "", phone: "", email: "", website: "", visitDate: "",
      capacityMin: "", capacityMax: "", space: "Ambos", rainPlan: false, parking: "", curfew: "", currency: "CLP", priceBase: "", pricePerPerson: "", deposit: "", priceNotes: "",
      services: [], exclusiveSuppliers: false, photos: [], ratings: {}, overall: { a: "", b: "" }, pros: "", cons: "", notes: "" });
  }
  function save() { if (!form.name.trim()) return; setVenues((vs) => (vs.some((x) => x.id === form.id) ? vs.map((x) => (x.id === form.id ? form : x)) : [...vs, form])); setForm(null); }
  function remove(id) { if (!window.confirm("¿Borrar este lugar?")) return; setVenues((vs) => vs.filter((x) => x.id !== id)); setCompareIds((c) => c.filter((x) => x !== id)); setForm(null); }
  function toggleIn(key, val) { setForm((f) => ({ ...f, [key]: f[key].includes(val) ? f[key].filter((x) => x !== val) : [...f[key], val] })); }
  function setRating(aspect, who, val) { setForm((f) => ({ ...f, ratings: { ...f.ratings, [aspect]: { ...(f.ratings[aspect] || {}), [who]: val } } })); }
  async function addPhotos(e) {
    const files = Array.from(e.target.files || []);
    e.target.value = "";
    const room = MAX_VENUE_PHOTOS - form.photos.length;
    if (files.length > room) alert(`Máximo ${MAX_VENUE_PHOTOS} fotos por lugar (para que la app siga rápida). Se agregarán ${Math.max(room, 0)}.`);
    try {
      const urls = await Promise.all(files.slice(0, Math.max(room, 0)).map((f) => compressImage(f, 1000, 0.72)));
      setForm((f) => ({ ...f, photos: [...f.photos, ...urls] }));
    } catch (err) { alert("No se pudo procesar alguna imagen."); }
  }
  function toSupplier(v) {
    const total = venueTotal(v, headcount);
    setSuppliers((s) => [...s, { id: uid(), company: v.name, category: "Lugar", cost: total ?? "", phone: v.phone, email: v.email, address: v.address, paymentStatus: "No pagado", depositPaid: false, notes: `Agregado desde Lugares${total ? ` (total estimado para ${headcount} invitados)` : ""}.` }]);
    alert(`"${v.name}" se agregó a Proveedores (categoría Lugar).`);
  }
  function toggleCompare(id) { setCompareIds((c) => (c.includes(id) ? c.filter((x) => x !== id) : [...c, id])); }

  const statusColor = (st) => ({ Favorito: T.primary, Reservado: T.secondaryDark, Descartado: T.muted, Visitado: T.secondary }[st] || T.primaryDark);
  const sorted = venues
    .filter((v) => statusFilter === "Todos" || v.status === statusFilter)
    .slice()
    .sort((x, y) => {
      if (sortBy === "score") return (venueCombined(y) ?? -1) - (venueCombined(x) ?? -1);
      if (sortBy === "price") return (venueTotal(x, headcount) ?? Infinity) - (venueTotal(y, headcount) ?? Infinity);
      if (sortBy === "capacity") return Number(y.capacityMax || 0) - Number(x.capacityMax || 0);
      return x.name.localeCompare(y.name);
    });
  const compared = compareIds.map((id) => venues.find((v) => v.id === id)).filter(Boolean);

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="flex gap-2">
          {compared.length >= 2 && <Button T={T} variant="outline" onClick={() => setComparing(true)}><Scale size={14} /> Comparar ({compared.length})</Button>}
          <Button T={T} onClick={startNew}><Plus size={15} /> Agregar lugar</Button>
        </div>}>
        Lugares
      </SectionTitle>
      <p className="text-xs mb-4" style={{ color: T.muted }}>Evalúen cada centro de eventos por separado ({nameA} y {nameB}, de 1 a 10), guarden fotos, precios y qué incluye. Marca "Comparar" en 2 o más para verlos lado a lado.</p>

      <div className="flex gap-2 mb-3 flex-wrap items-center">
        {["Todos", ...VENUE_STATUS].map((st) => <Chip key={st} T={T} active={statusFilter === st} onClick={() => setStatusFilter(st)}>{st}</Chip>)}
      </div>
      <div className="flex gap-4 mb-5 flex-wrap items-center text-sm">
        <label className="flex items-center gap-2"><span style={{ color: T.muted }}>Ordenar por</span>
          <select style={inputStyle(T)} value={sortBy} onChange={(e) => setSortBy(e.target.value)}>
            <option value="score">Mejor nota</option><option value="price">Menor precio</option><option value="capacity">Mayor capacidad</option><option value="name">Nombre</option>
          </select>
        </label>
        <label className="flex items-center gap-2"><span style={{ color: T.muted }}>Invitados estimados (para calcular totales)</span>
          <input type="number" min="0" style={{ ...inputStyle(T), width: 90 }} value={headcount} onChange={(e) => setHeadcount(Number(e.target.value) || 0)} />
        </label>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
        {sorted.length === 0 && <div className="col-span-full"><EmptyState T={T} text="Sin lugares todavía. Agrega el primero que estén evaluando." /></div>}
        {sorted.map((v) => {
          const total = venueTotal(v, headcount);
          const tooSmall = v.capacityMax && Number(v.capacityMax) < headcount;
          return (
            <Card T={T} key={v.id} style={{ padding: 0, overflow: "hidden", opacity: v.status === "Descartado" ? 0.6 : 1 }} onClick={() => setForm({ ...v })}>
              <div className="h-40 flex items-center justify-center relative" style={{ background: T.cardAlt }}>
                {v.photos?.[0] ? <PositionedPhoto src={photoSrc(v.photos[0])} pos={photoPos(v.photos[0])} className="w-full h-full" /> : <Building2 size={32} color={T.muted} />}
                <span className="absolute top-2 left-2 text-xs px-2 py-1 rounded-full" style={{ background: statusColor(v.status), color: "#fff" }}>{v.status}</span>
                {v.photos?.length > 1 && <span className="absolute bottom-2 right-2 text-xs px-2 py-0.5 rounded-full" style={{ background: "rgba(0,0,0,0.55)", color: "#fff" }}>{v.photos.length} fotos</span>}
                <label onClick={(e) => e.stopPropagation()} className="absolute top-2 right-2 text-xs px-2 py-1 rounded-full flex items-center gap-1 cursor-pointer" style={{ background: compareIds.includes(v.id) ? T.primary : "rgba(255,255,255,0.9)", color: compareIds.includes(v.id) ? "#fff" : "#3B2A20" }}>
                  <input type="checkbox" className="hidden" checked={compareIds.includes(v.id)} onChange={() => toggleCompare(v.id)} /><Scale size={11} /> Comparar
                </label>
              </div>
              <div className="p-4">
                <div className="text-base font-medium">{v.name}</div>
                <div className="text-xs mt-0.5" style={{ color: T.muted }}>{v.type}{v.address && ` · ${v.address}`}</div>
                <div className="flex gap-2 mt-3 text-xs flex-wrap">
                  <span className="px-2 py-1 rounded-lg flex items-center gap-1" style={{ background: T.cardAlt }}><Star size={11} color={T.primary} /> {nameA} {fmtScore(venueAvg(v, "a"))}</span>
                  <span className="px-2 py-1 rounded-lg flex items-center gap-1" style={{ background: T.cardAlt }}><Star size={11} color={T.secondary} /> {nameB} {fmtScore(venueAvg(v, "b"))}</span>
                  <span className="px-2 py-1 rounded-lg font-medium" style={{ background: T.soft, color: T.text }}>Pareja {fmtScore(venueCombined(v))}</span>
                </div>
                <div className="flex justify-between mt-3 text-xs" style={{ color: T.muted }}>
                  <span style={{ color: tooSmall ? T.alert : T.muted }}><Users size={11} className="inline mr-1" />{v.capacityMin || v.capacityMax ? `${v.capacityMin || "?"}–${v.capacityMax || "?"} pers.` : "Capacidad —"}{tooSmall && " (no caben)"}</span>
                  <span className="font-medium" style={{ color: T.text }}>{total !== null ? fmtMoney(total, v.currency) : "Precio —"}</span>
                </div>
                {(v.vibes?.length > 0 || v.services?.length > 0) && (
                  <div className="flex gap-1 mt-3 flex-wrap">
                    {v.vibes.map((x) => <span key={x} className="text-xs px-1.5 py-0.5 rounded" style={{ background: T.soft, color: T.text }}>{x}</span>)}
                    {v.services?.length > 0 && <span className="text-xs px-1.5 py-0.5 rounded" style={{ background: T.cardAlt, color: T.muted }}>Incluye {v.services.length} servicio(s)</span>}
                  </div>
                )}
              </div>
            </Card>
          );
        })}
      </div>

      {form && (
        <Modal T={T} title={venues.some((x) => x.id === form.id) ? form.name || "Editar lugar" : "Nuevo lugar"} onClose={() => setForm(null)}>
          <div className="flex flex-col gap-6">
            <section className="grid grid-cols-2 gap-4">
              <div className="col-span-2"><Field T={T} label="Nombre del lugar"><input style={{ ...inputStyle(T), width: "100%" }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} autoFocus /></Field></div>
              <Field T={T} label="Tipo de lugar"><select style={inputStyle(T)} value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>{VENUE_TYPES.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field T={T} label="Estado"><select style={inputStyle(T)} value={form.status} onChange={(e) => setForm({ ...form, status: e.target.value })}>{VENUE_STATUS.map((x) => <option key={x}>{x}</option>)}</select></Field>
              <Field T={T} label="Para qué evento"><select style={inputStyle(T)} value={form.eventId} onChange={(e) => setForm({ ...form, eventId: e.target.value })}><option value="">Cualquiera / por definir</option>{events.map((ev) => <option key={ev.id} value={ev.id}>{ev.name}</option>)}</select></Field>
              <Field T={T} label="Fecha de visita"><input type="date" style={inputStyle(T)} value={form.visitDate} onChange={(e) => setForm({ ...form, visitDate: e.target.value })} /></Field>
            </section>

            <section>
              <div className="flex justify-between items-center mb-2">
                <h3 className="text-sm font-medium">Fotos ({form.photos.length}/{MAX_VENUE_PHOTOS})</h3>
                {form.photos.length < MAX_VENUE_PHOTOS && <Button T={T} small variant="outline" onClick={() => photoRef.current?.click()}><Camera size={13} /> Agregar fotos</Button>}
                <input ref={photoRef} type="file" accept="image/*" multiple className="hidden" onChange={addPhotos} />
              </div>
              <div className="flex gap-2 flex-wrap">
                {form.photos.length === 0 && <span className="text-xs" style={{ color: T.muted }}>Sin fotos. La primera será la portada.</span>}
                {form.photos[0] && (
                  <div className="w-full mb-1">
                    <div className="h-40 rounded-lg overflow-hidden relative" style={{ border: `1px solid ${T.line}` }}>
                      <PositionedPhoto editing src={photoSrc(form.photos[0])} pos={photoPos(form.photos[0])} className="w-full h-full"
                        onChange={(p) => setForm((f) => ({ ...f, photos: [{ src: photoSrc(f.photos[0]), pos: p }, ...f.photos.slice(1)] }))} />
                      <span className="absolute top-2 left-2 text-xs px-2 py-1 rounded-full pointer-events-none" style={{ background: "rgba(0,0,0,0.55)", color: "#fff" }}>Portada · arrastra para encuadrar</span>
                    </div>
                  </div>
                )}
                {form.photos.map((p, i) => (
                  <div key={i} className="relative w-28 h-24 rounded-lg overflow-hidden" style={{ border: `1px solid ${T.line}` }}>
                    <a href={photoSrc(p)} target="_blank" rel="noreferrer"><PositionedPhoto src={photoSrc(p)} pos={photoPos(p)} className="w-full h-full" /></a>
                    <button onClick={() => setForm((f) => ({ ...f, photos: f.photos.filter((_, j) => j !== i) }))} className="absolute top-1 right-1 w-5 h-5 rounded-full flex items-center justify-center" style={{ background: "rgba(0,0,0,0.55)" }}><X size={11} color="#fff" /></button>
                    {i > 0 && <button onClick={() => setForm((f) => ({ ...f, photos: [f.photos[i], ...f.photos.filter((_, j) => j !== i)] }))} className="absolute bottom-1 left-1 text-[10px] px-1.5 rounded" style={{ background: "rgba(0,0,0,0.55)", color: "#fff" }}>Portada</button>}
                  </div>
                ))}
              </div>
            </section>

            <section>
              <h3 className="text-sm font-medium mb-2">Evaluación (1 a 10)</h3>
              <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${T.line}` }}>
                <div className="grid text-xs font-medium px-3 py-2" style={{ gridTemplateColumns: "1fr 72px 72px", background: T.cardAlt, color: T.muted }}><span>Aspecto</span><span>{nameA}</span><span>{nameB}</span></div>
                {VENUE_ASPECTS.map((a) => (
                  <div key={a} className="grid items-center px-3 py-1.5 text-sm" style={{ gridTemplateColumns: "1fr 72px 72px", borderTop: `1px solid ${T.line}` }}>
                    <span>{a}</span>
                    <ScoreSelect T={T} value={form.ratings[a]?.a} onChange={(n) => setRating(a, "a", n)} />
                    <ScoreSelect T={T} value={form.ratings[a]?.b} onChange={(n) => setRating(a, "b", n)} />
                  </div>
                ))}
                <div className="grid items-center px-3 py-2 text-sm font-medium" style={{ gridTemplateColumns: "1fr 72px 72px", borderTop: `1px solid ${T.line}`, background: T.cardAlt }}>
                  <span>Promedio (pareja: {fmtScore(venueCombined(form))})</span><span>{fmtScore(venueAvg(form, "a"))}</span><span>{fmtScore(venueAvg(form, "b"))}</span>
                </div>
                <div className="grid items-center px-3 py-1.5 text-sm" style={{ gridTemplateColumns: "1fr 72px 72px", borderTop: `1px solid ${T.line}` }}>
                  <span>Nota "de guata" (sensación general)</span>
                  <ScoreSelect T={T} value={form.overall?.a} onChange={(n) => setForm((f) => ({ ...f, overall: { ...f.overall, a: n } }))} />
                  <ScoreSelect T={T} value={form.overall?.b} onChange={(n) => setForm((f) => ({ ...f, overall: { ...f.overall, b: n } }))} />
                </div>
              </div>
            </section>

            <section className="grid grid-cols-2 gap-4">
              <h3 className="col-span-2 text-sm font-medium">Precio</h3>
              <Field T={T} label="Moneda"><select style={inputStyle(T)} value={form.currency} onChange={(e) => setForm({ ...form, currency: e.target.value })}><option>CLP</option><option>UF</option></select></Field>
              <Field T={T} label="Arriendo / precio fijo"><input type="number" style={inputStyle(T)} value={form.priceBase} onChange={(e) => setForm({ ...form, priceBase: e.target.value })} /></Field>
              <Field T={T} label="Precio por persona"><input type="number" style={inputStyle(T)} value={form.pricePerPerson} onChange={(e) => setForm({ ...form, pricePerPerson: e.target.value })} /></Field>
              <Field T={T} label="Reserva / pie"><input type="number" style={inputStyle(T)} value={form.deposit} onChange={(e) => setForm({ ...form, deposit: e.target.value })} /></Field>
              <div className="col-span-2 text-sm px-3 py-2 rounded-lg" style={{ background: T.cardAlt }}>Total estimado para {headcount} invitados: <b>{venueTotal(form, headcount) !== null ? fmtMoney(venueTotal(form, headcount), form.currency) : "—"}</b></div>
              <div className="col-span-2"><Field T={T} label="Qué incluye el precio / condiciones"><textarea style={{ ...inputStyle(T), width: "100%", minHeight: 50 }} value={form.priceNotes} onChange={(e) => setForm({ ...form, priceNotes: e.target.value })} /></Field></div>
            </section>

            <section>
              <h3 className="text-sm font-medium mb-2">Servicios incluidos</h3>
              <div className="flex gap-2 flex-wrap">{VENUE_SERVICES.map((x) => <Chip key={x} T={T} active={form.services.includes(x)} onClick={() => toggleIn("services", x)}>{form.services.includes(x) && "✓ "}{x}</Chip>)}</div>
              <div className="flex items-center gap-3 mt-3"><Toggle T={T} checked={form.exclusiveSuppliers} onChange={(v) => setForm({ ...form, exclusiveSuppliers: v })} /><span className="text-sm">Obliga a usar sus proveedores (no se puede traer externos)</span></div>
            </section>

            <section>
              <h3 className="text-sm font-medium mb-2">Vibe de la decoración</h3>
              <div className="flex gap-2 flex-wrap">{VENUE_VIBES.map((x) => <Chip key={x} T={T} active={form.vibes.includes(x)} onClick={() => toggleIn("vibes", x)}>{x}</Chip>)}</div>
            </section>

            <section className="grid grid-cols-2 gap-4">
              <h3 className="col-span-2 text-sm font-medium">Espacio y logística</h3>
              <Field T={T} label="Capacidad mínima (personas)"><input type="number" style={inputStyle(T)} value={form.capacityMin} onChange={(e) => setForm({ ...form, capacityMin: e.target.value })} /></Field>
              <Field T={T} label="Capacidad máxima (personas)"><input type="number" style={inputStyle(T)} value={form.capacityMax} onChange={(e) => setForm({ ...form, capacityMax: e.target.value })} /></Field>
              <Field T={T} label="Espacios"><select style={inputStyle(T)} value={form.space} onChange={(e) => setForm({ ...form, space: e.target.value })}><option>Interior</option><option>Exterior</option><option>Ambos</option></select></Field>
              <Field T={T} label="Hora de término / toque de queda"><input style={inputStyle(T)} placeholder="Ej. 04:00" value={form.curfew} onChange={(e) => setForm({ ...form, curfew: e.target.value })} /></Field>
              <Field T={T} label="Estacionamiento (cupos)"><input style={inputStyle(T)} value={form.parking} onChange={(e) => setForm({ ...form, parking: e.target.value })} /></Field>
              <div className="flex items-center gap-3 mt-5"><Toggle T={T} checked={form.rainPlan} onChange={(v) => setForm({ ...form, rainPlan: v })} /><span className="text-sm">Tiene plan lluvia</span></div>
            </section>

            <section className="grid grid-cols-2 gap-4">
              <h3 className="col-span-2 text-sm font-medium">Contacto</h3>
              <div className="col-span-2"><Field T={T} label="Dirección / comuna"><input style={{ ...inputStyle(T), width: "100%" }} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></Field></div>
              {form.address && <a href={mapsLink(form.address)} target="_blank" rel="noreferrer" className="col-span-2 text-xs flex items-center gap-1" style={{ color: T.primaryDark }}><MapPin size={12} /> Ver en Google Maps</a>}
              <Field T={T} label="Persona de contacto"><input style={inputStyle(T)} value={form.contact} onChange={(e) => setForm({ ...form, contact: e.target.value })} /></Field>
              <Field T={T} label="Teléfono"><input type="tel" style={inputStyle(T)} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
              <Field T={T} label="Email"><input type="email" style={inputStyle(T)} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
              <Field T={T} label="Web / Instagram"><input style={inputStyle(T)} value={form.website} onChange={(e) => setForm({ ...form, website: e.target.value })} /></Field>
            </section>

            <section className="grid grid-cols-2 gap-4">
              <Field T={T} label="👍 Lo mejor"><textarea style={{ ...inputStyle(T), minHeight: 60 }} value={form.pros} onChange={(e) => setForm({ ...form, pros: e.target.value })} /></Field>
              <Field T={T} label="👎 Lo que no nos convence"><textarea style={{ ...inputStyle(T), minHeight: 60 }} value={form.cons} onChange={(e) => setForm({ ...form, cons: e.target.value })} /></Field>
              <div className="col-span-2"><Field T={T} label="Notas"><textarea style={{ ...inputStyle(T), width: "100%", minHeight: 50 }} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></Field></div>
            </section>
          </div>
          <div className="flex gap-2 mt-6 justify-end flex-wrap">
            {venues.some((x) => x.id === form.id) && <Button T={T} variant="outline" onClick={() => remove(form.id)}><Trash2 size={13} /> Eliminar</Button>}
            {venues.some((x) => x.id === form.id) && <Button T={T} variant="outline" onClick={() => toSupplier(form)}><Briefcase size={13} /> Pasar a Proveedores</Button>}
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}

      {comparing && compared.length >= 2 && <VenueCompare T={T} venues={compared} headcount={headcount} nameA={nameA} nameB={nameB} onClose={() => setComparing(false)} onRemove={toggleCompare} />}
    </div>
  );
}

function VenueCompare({ T, venues, headcount, nameA, nameB, onClose, onRemove }) {
  // Devuelve el id del mejor lugar según una métrica (para destacarlo), ignorando vacíos y empates totales.
  function best(fn, lowest = false) {
    const vals = venues.map((v) => ({ id: v.id, n: fn(v) })).filter((x) => x.n !== null && x.n !== undefined && x.n !== "" && !Number.isNaN(Number(x.n)));
    if (vals.length < 2) return null;
    const target = lowest ? Math.min(...vals.map((x) => Number(x.n))) : Math.max(...vals.map((x) => Number(x.n)));
    const winners = vals.filter((x) => Number(x.n) === target);
    return winners.length === vals.length ? null : new Set(winners.map((x) => x.id));
  }
  const yesNo = (b) => (b ? "✓" : "—");
  const rows = [
    { group: "General" },
    { label: "Tipo", get: (v) => v.type },
    { label: "Estado", get: (v) => v.status },
    { label: "Vibe", get: (v) => v.vibes?.join(", ") || "—" },
    { group: "Notas" },
    { label: `Promedio ${nameA}`, get: (v) => fmtScore(venueAvg(v, "a")), best: best((v) => venueAvg(v, "a")) },
    { label: `Promedio ${nameB}`, get: (v) => fmtScore(venueAvg(v, "b")), best: best((v) => venueAvg(v, "b")) },
    { label: "Promedio pareja", get: (v) => fmtScore(venueCombined(v)), best: best(venueCombined), strong: true },
    { label: "Nota de guata", get: (v) => `${v.overall?.a || "—"} / ${v.overall?.b || "—"}` },
    ...VENUE_ASPECTS.map((a) => ({ label: a, get: (v) => `${v.ratings?.[a]?.a || "—"} / ${v.ratings?.[a]?.b || "—"}`, best: best((v) => { const s = [v.ratings?.[a]?.a, v.ratings?.[a]?.b].filter((n) => n > 0); return s.length ? s.reduce((x, y) => x + y, 0) / s.length : null; }), small: true })),
    { group: "Precio" },
    { label: "Arriendo / fijo", get: (v) => (v.priceBase ? fmtMoney(v.priceBase, v.currency) : "—") },
    { label: "Por persona", get: (v) => (v.pricePerPerson ? fmtMoney(v.pricePerPerson, v.currency) : "—") },
    { label: `Total estimado (${headcount} inv.)`, get: (v) => { const t = venueTotal(v, headcount); return t !== null ? fmtMoney(t, v.currency) : "—"; }, best: venues.every((v) => v.currency === venues[0].currency) ? best((v) => venueTotal(v, headcount), true) : null, strong: true },
    { label: "Reserva / pie", get: (v) => (v.deposit ? fmtMoney(v.deposit, v.currency) : "—") },
    { label: "Qué incluye", get: (v) => v.priceNotes || "—", small: true },
    { group: "Espacio y logística" },
    { label: "Capacidad", get: (v) => `${v.capacityMin || "?"}–${v.capacityMax || "?"}`, best: best((v) => v.capacityMax) },
    { label: "Espacios", get: (v) => v.space },
    { label: "Plan lluvia", get: (v) => yesNo(v.rainPlan) },
    { label: "Estacionamiento", get: (v) => v.parking || "—" },
    { label: "Hora de término", get: (v) => v.curfew || "—" },
    { label: "Proveedores obligatorios", get: (v) => (v.exclusiveSuppliers ? "Sí" : "No") },
    { group: "Servicios incluidos" },
    { label: "Cantidad", get: (v) => v.services?.length || 0, best: best((v) => v.services?.length || 0) },
    ...VENUE_SERVICES.map((s) => ({ label: s, get: (v) => yesNo(v.services?.includes(s)), small: true })),
    { group: "Comentarios" },
    { label: "Lo mejor", get: (v) => v.pros || "—", small: true },
    { label: "Lo que no convence", get: (v) => v.cons || "—", small: true },
  ];
  const cols = `180px repeat(${venues.length}, minmax(160px, 1fr))`;
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-auto p-4 md:p-8" style={{ background: "rgba(0,0,0,0.5)" }} onClick={onClose}>
      <div onClick={(e) => e.stopPropagation()} className="w-full max-w-6xl rounded-2xl p-6" style={{ background: T.card }}>
        <div className="flex items-center justify-between mb-5">
          <h2 style={{ fontFamily: "Georgia, serif" }} className="text-xl">Comparar lugares</h2>
          <button onClick={onClose}><X size={18} color={T.muted} /></button>
        </div>
        <p className="text-xs mb-3" style={{ color: T.muted }}>Las notas se muestran como {nameA} / {nameB}. Lo destacado es el mejor en esa fila.</p>
        <div className="overflow-x-auto">
          <div className="grid gap-x-3 min-w-fit" style={{ gridTemplateColumns: cols }}>
            <div />
            {venues.map((v) => (
              <div key={v.id} className="pb-3">
                <div className="h-28 rounded-lg overflow-hidden flex items-center justify-center mb-2" style={{ background: T.cardAlt }}>
                  {v.photos?.[0] ? <PositionedPhoto src={photoSrc(v.photos[0])} pos={photoPos(v.photos[0])} className="w-full h-full" /> : <Building2 size={24} color={T.muted} />}
                </div>
                <div className="flex items-start justify-between gap-1">
                  <span className="text-sm font-medium">{v.name}</span>
                  <button onClick={() => onRemove(v.id)} title="Quitar de la comparación"><X size={13} color={T.muted} /></button>
                </div>
              </div>
            ))}
            {rows.map((r, i) => r.group ? (
              <div key={i} className="text-xs font-medium uppercase tracking-wide pt-4 pb-1" style={{ gridColumn: "1 / -1", color: T.primaryDark, borderBottom: `1px solid ${T.line}` }}>{r.group}</div>
            ) : (
              <React.Fragment key={i}>
                <div className={`py-1.5 ${r.small ? "text-xs" : "text-sm"}`} style={{ color: T.muted, borderBottom: `1px solid ${T.line}` }}>{r.label}</div>
                {venues.map((v) => {
                  const win = r.best?.has(v.id);
                  return <div key={v.id} className={`py-1.5 px-2 ${r.small ? "text-xs" : "text-sm"} ${r.strong || win ? "font-medium" : ""}`} style={{ borderBottom: `1px solid ${T.line}`, background: win ? T.soft : "transparent", borderRadius: win ? 6 : 0 }}>{r.get(v)}</div>;
                })}
              </React.Fragment>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== Todos ============================== */
function Todos({ T, todos, setTodos, events, tags, setTags, couple }) {
  const [form, setForm] = useState(null);
  const [step, setStep] = useState("");
  const dragRef = useRef(null);
  const stepDragRef = useRef(null);
  function startNew() { setForm({ id: uid(), title: "", tag: tags[0] || "", checklist: [], assignedTo: "Ambos", due: "", scope: "shared", done: false }); }
  function save() { if (!form.title.trim()) return; setTodos((t) => (t.some((x) => x.id === form.id) ? t.map((x) => (x.id === form.id ? form : x)) : [...t, form])); setForm(null); }
  function importTodos(rows) {
    const created = rows.map((row) => {
      const title = pick(row, ["Titulo", "Título", "Title"]);
      if (!title) return null;
      const scopeText = pick(row, ["Evento", "Scope", "Para"]);
      const eid = matchEvent(scopeText, events);
      return { id: uid(), title, tag: pick(row, ["Etiqueta", "Tag"]) || tags[0] || "", checklist: [], assignedTo: pick(row, ["AsignadoA", "Asignado a", "Assigned"]) || "Ambos", due: pick(row, ["FechaLimite", "Fecha límite", "Due"]), scope: eid || "shared", done: truthy(pick(row, ["Hecho", "Done"])) };
    }).filter(Boolean);
    if (created.length) setTodos((t) => [...t, ...created]);
  }

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="flex gap-2">
          <ImportExcel T={T} label="Importar Excel" hint="Columnas: Titulo, Etiqueta, AsignadoA, FechaLimite, Evento" onRows={importTodos} />
          <Button T={T} onClick={startNew}><Plus size={15} /> Agregar tarea</Button>
        </div>}>
        Tareas
      </SectionTitle>
      <div className="flex flex-col gap-2">
        {todos.length === 0 && <EmptyState T={T} text="Sin tareas todavía." />}
        {todos.map((t) => (
          <Card T={T} key={t.id} style={{ padding: "10px 16px" }} draggable onDragStart={() => (dragRef.current = t.id)} onDragOver={(e) => e.preventDefault()} onDrop={() => setTodos((ts) => reorderById(ts, dragRef.current, t.id))}>
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <GripVertical size={14} color={T.muted} style={{ cursor: "grab" }} />
                <button className="flex items-center gap-3" onClick={() => setTodos((ts) => ts.map((x) => (x.id === t.id ? { ...x, done: !x.done } : x)))}>
                  <span style={{ width: 18, height: 18, border: `1.5px solid ${T.soft}`, borderRadius: 6, background: t.done ? T.secondary : "transparent" }} className="flex items-center justify-center">{t.done && <Check size={12} color="#fff" />}</span>
                  <span className="text-sm text-left" style={{ textDecoration: t.done ? "line-through" : "none", color: t.done ? T.muted : T.text }} onClick={(e) => { e.stopPropagation(); setForm(t); }}>{t.title}</span>
                </button>
              </div>
              <div className="flex items-center gap-3 text-xs" style={{ color: T.muted }}>
                {t.tag && <span className="px-2 py-1 rounded-full flex items-center gap-1" style={{ background: T.cardAlt }}><Tag size={10} />{t.tag}</span>}
                {t.checklist?.length > 0 && <span>{t.checklist.filter((c) => c.done).length}/{t.checklist.length}</span>}
                <span>{t.assignedTo}</span>
                {t.due && <span>{new Date(t.due).toLocaleDateString("es-CL")}</span>}
                <button onClick={() => setTodos((ts) => ts.filter((x) => x.id !== t.id))}><Trash2 size={14} /></button>
              </div>
            </div>
          </Card>
        ))}
      </div>

      {form && (
        <Modal T={T} title={form.title ? "Editar tarea" : "Nueva tarea"} onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2"><Field T={T} label="Título"><input style={{ ...inputStyle(T), width: "100%" }} value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} autoFocus /></Field></div>
            <Field T={T} label="Etiqueta"><select style={inputStyle(T)} value={form.tag} onChange={(e) => setForm({ ...form, tag: e.target.value })}>{tags.map((tg) => <option key={tg}>{tg}</option>)}</select></Field>
            <Field T={T} label="Asignado a"><select style={inputStyle(T)} value={form.assignedTo} onChange={(e) => setForm({ ...form, assignedTo: e.target.value })}><option>{couple.a}</option><option>{couple.b}</option><option>Ambos</option></select></Field>
            <Field T={T} label="Fecha límite"><input type="date" style={inputStyle(T)} value={form.due} onChange={(e) => setForm({ ...form, due: e.target.value })} /></Field>
            <Field T={T} label="Para"><select style={inputStyle(T)} value={form.scope} onChange={(e) => setForm({ ...form, scope: e.target.value })}><option value="shared">Todos los eventos</option>{events.map((ev) => <option key={ev.id} value={ev.id}>{ev.name}</option>)}</select></Field>
            <div className="col-span-2">
              <Field T={T} label="Checklist">
                <div className="flex flex-col gap-1.5 mt-1">
                  {form.checklist.map((c) => (
                    <div key={c.id} className="flex items-center gap-2 text-sm" draggable onDragStart={() => (stepDragRef.current = c.id)} onDragOver={(e) => e.preventDefault()} onDrop={() => setForm({ ...form, checklist: reorderById(form.checklist, stepDragRef.current, c.id) })}>
                      <GripVertical size={12} color={T.muted} style={{ cursor: "grab" }} />
                      <input type="checkbox" checked={c.done} onChange={() => setForm({ ...form, checklist: form.checklist.map((x) => (x.id === c.id ? { ...x, done: !x.done } : x)) })} />
                      <span style={{ textDecoration: c.done ? "line-through" : "none" }}>{c.text}</span>
                      <button onClick={() => setForm({ ...form, checklist: form.checklist.filter((x) => x.id !== c.id) })}><X size={12} color={T.muted} /></button>
                    </div>
                  ))}
                  <div className="flex gap-2 mt-1">
                    <input placeholder="Nuevo paso" style={{ ...inputStyle(T), flex: 1 }} value={step} onChange={(e) => setStep(e.target.value)}
                      onKeyDown={(e) => { if (e.key === "Enter" && step.trim()) { setForm({ ...form, checklist: [...form.checklist, { id: uid(), text: step, done: false }] }); setStep(""); } }} />
                    <Button T={T} small variant="outline" onClick={() => { if (step.trim()) { setForm({ ...form, checklist: [...form.checklist, { id: uid(), text: step, done: false }] }); setStep(""); } }}>Agregar</Button>
                  </div>
                </div>
              </Field>
            </div>
          </div>
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== Budget ============================== */
function Budget({ T, budget, setBudget, activeEvent, events, cats, setCats, suppliers }) {
  const [form, setForm] = useState(null);
  const dragRef = useRef(null);
  const visible = activeEvent === "all" ? budget : budget.filter((b) => b.eventId === activeEvent || b.eventId === "shared");
  const total = (b) => (b.mode === "qty" ? Number(b.qty || 0) * Number(b.unitCost || 0) : Number(b.amount || 0));
  const totalCLP = visible.filter((b) => b.currency !== "UF").reduce((s, b) => s + total(b), 0);
  const totalUF = visible.filter((b) => b.currency === "UF").reduce((s, b) => s + total(b), 0);
  function startNew() { setForm({ id: uid(), title: "", category: cats[0], mode: "unique", amount: "", qty: "", unitCost: "", paid: false, currency: "CLP", eventId: pickEvent(activeEvent, events), supplierId: "" }); }
  function save() { if (!form.title.trim()) return; setBudget((b) => (b.some((x) => x.id === form.id) ? b.map((x) => (x.id === form.id ? form : x)) : [...b, form])); setForm(null); }
  function applySupplier(supplierId) {
    const s = suppliers.find((x) => x.id === supplierId);
    if (!s) { setForm({ ...form, supplierId: "" }); return; }
    setForm({ ...form, supplierId, title: form.title || s.company, category: s.category || form.category, mode: "unique", amount: s.cost || form.amount });
  }
  function importBudget(rows) {
    const created = rows.map((row) => {
      const title = pick(row, ["Titulo", "Título", "Title"]);
      if (!title) return null;
      const isQty = /cantidad|qty/i.test(pick(row, ["Tipo", "Mode"]));
      return {
        id: uid(), title, category: pick(row, ["Categoria", "Categoría", "Category"]) || cats[0],
        mode: isQty ? "qty" : "unique", amount: pick(row, ["Monto", "Amount"]), qty: pick(row, ["Cantidad", "Qty"]), unitCost: pick(row, ["PrecioUnitario", "Precio unitario", "UnitCost"]),
        paid: truthy(pick(row, ["Pagado", "Paid"])), currency: /uf/i.test(pick(row, ["Moneda", "Currency"])) ? "UF" : "CLP",
        eventId: matchEvent(pick(row, ["Evento", "Event"]), events) || "shared", supplierId: "",
      };
    }).filter(Boolean);
    if (created.length) setBudget((b) => [...b, ...created]);
  }

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="flex gap-2">
          <ImportExcel T={T} label="Importar Excel" hint="Columnas: Titulo, Categoria, Tipo, Monto, Cantidad, PrecioUnitario, Pagado, Moneda, Evento" onRows={importBudget} />
          <Button T={T} onClick={startNew}><Plus size={15} /> Agregar ítem</Button>
        </div>}>
        Presupuesto
      </SectionTitle>
      <div className="flex gap-4 mb-5">
        <Card T={T} style={{ flex: 1 }}><div className="text-xs" style={{ color: T.muted }}>Total CLP</div><div style={{ fontFamily: "Georgia, serif" }} className="text-2xl">{fmtMoney(totalCLP, "CLP")}</div></Card>
        <Card T={T} style={{ flex: 1 }}><div className="text-xs" style={{ color: T.muted }}>Total UF</div><div style={{ fontFamily: "Georgia, serif" }} className="text-2xl">{fmtMoney(totalUF, "UF")}</div></Card>
      </div>
      <div className="flex flex-col gap-2">
        {visible.length === 0 && <EmptyState T={T} text="Sin ítems de presupuesto todavía." />}
        {visible.map((b) => (
          <Card T={T} key={b.id} style={{ padding: "10px 16px" }} onClick={() => setForm(b)} draggable onDragStart={(e) => { e.stopPropagation(); dragRef.current = b.id; }} onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.stopPropagation(); setBudget((bs) => reorderById(bs, dragRef.current, b.id)); }}>
            <div className="flex items-center justify-between text-sm">
              <div className="flex items-center gap-2"><GripVertical size={13} color={T.muted} style={{ cursor: "grab" }} /><span className="font-medium">{b.title}</span><span className="ml-2 text-xs" style={{ color: T.muted }}>{b.category} · {events.find((e) => e.id === b.eventId)?.name || "Compartido"}{b.supplierId && ` · ${suppliers.find((s) => s.id === b.supplierId)?.company || "proveedor"}`}</span></div>
              <div className="flex items-center gap-4">
                <span className="text-xs px-2 py-1 rounded-full" style={{ background: b.paid ? T.cardAlt : "transparent", color: b.paid ? T.secondaryDark : T.alert, border: b.paid ? "none" : `1px solid ${T.alert}` }}>{b.paid ? "Pagado" : "No pagado"}</span>
                <span>{fmtMoney(total(b), b.currency)}</span>
                <button onClick={(e) => { e.stopPropagation(); setBudget((bs) => bs.filter((x) => x.id !== b.id)); }}><Trash2 size={14} color={T.muted} /></button>
              </div>
            </div>
          </Card>
        ))}
      </div>

      {form && (
        <Modal T={T} title={form.title ? "Editar ítem" : "Nuevo ítem"} onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2"><Field T={T} label="Título"><input style={{ ...inputStyle(T), width: "100%" }} value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} autoFocus /></Field></div>
            <Field T={T} label="Proveedor (opcional)">
              <select style={inputStyle(T)} value={form.supplierId || ""} onChange={(e) => applySupplier(e.target.value)}>
                <option value="">Sin proveedor</option>{suppliers.map((s) => <option key={s.id} value={s.id}>{s.company}</option>)}
              </select>
            </Field>
            <Field T={T} label="Categoría"><select style={inputStyle(T)} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>{cats.map((c) => <option key={c}>{c}</option>)}</select></Field>
            <Field T={T} label="Moneda"><select style={inputStyle(T)} value={form.currency} onChange={(e) => setForm({ ...form, currency: e.target.value })}><option>CLP</option><option>UF</option></select></Field>
            <Field T={T} label="Tipo de gasto"><select style={inputStyle(T)} value={form.mode} onChange={(e) => setForm({ ...form, mode: e.target.value })}><option value="unique">Único</option><option value="qty">Por cantidad</option></select></Field>
            <Field T={T} label="Evento"><select style={inputStyle(T)} value={form.eventId} onChange={(e) => setForm({ ...form, eventId: e.target.value })}>{events.map((ev) => <option key={ev.id} value={ev.id}>{ev.name}</option>)}<option value="shared">Compartido</option></select></Field>
            {form.mode === "unique" ? (
              <Field T={T} label="Monto"><input type="number" style={inputStyle(T)} value={form.amount} onChange={(e) => setForm({ ...form, amount: e.target.value })} /></Field>
            ) : (
              <>
                <Field T={T} label="Cantidad"><input type="number" style={inputStyle(T)} value={form.qty} onChange={(e) => setForm({ ...form, qty: e.target.value })} /></Field>
                <Field T={T} label="Precio unitario"><input type="number" style={inputStyle(T)} value={form.unitCost} onChange={(e) => setForm({ ...form, unitCost: e.target.value })} /></Field>
              </>
            )}
            <div className="flex items-center gap-3 mt-5"><Toggle T={T} checked={form.paid} onChange={(v) => setForm({ ...form, paid: v })} /><span className="text-sm">Pagado</span></div>
          </div>
          {form.supplierId && <p className="text-xs mt-2" style={{ color: T.muted }}>El monto se precargó desde el proveedor — ajústalo si el costo cambió.</p>}
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== Suppliers ============================== */
function Suppliers({ T, suppliers, setSuppliers, cats, setCats }) {
  const [form, setForm] = useState(null);
  const [newCat, setNewCat] = useState("");
  const dragRef = useRef(null);
  function startNew() { setForm({ id: uid(), company: "", category: cats[0], cost: "", phone: "", email: "", address: "", paymentStatus: "No pagado", depositPaid: false, notes: "" }); }
  function save() { if (!form.company.trim()) return; setSuppliers((s) => (s.some((x) => x.id === form.id) ? s.map((x) => (x.id === form.id ? form : x)) : [...s, form])); setForm(null); }
  function importSuppliers(rows) {
    const created = rows.map((row) => {
      const company = pick(row, ["Compania", "Compañía", "Company"]);
      if (!company) return null;
      return {
        id: uid(), company, category: pick(row, ["Categoria", "Categoría", "Category"]) || cats[0], cost: pick(row, ["Costo", "Cost"]),
        phone: pick(row, ["Telefono", "Teléfono", "Phone"]), email: pick(row, ["Email"]), address: pick(row, ["Direccion", "Dirección", "Address"]),
        paymentStatus: pick(row, ["EstadoPago", "Estado de pago", "PaymentStatus"]) || "No pagado",
        depositPaid: truthy(pick(row, ["ReservaPagada", "Reserva pagada", "DepositPaid"])), notes: pick(row, ["Notas", "Notes"]),
      };
    }).filter(Boolean);
    if (created.length) setSuppliers((s) => [...s, ...created]);
  }

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="flex gap-2">
          <ImportExcel T={T} label="Importar Excel" hint="Columnas: Compania, Categoria, Costo, Telefono, Email, Direccion, EstadoPago, ReservaPagada, Notas" onRows={importSuppliers} />
          <Button T={T} onClick={startNew}><Plus size={15} /> Agregar proveedor</Button>
        </div>}>
        Proveedores
      </SectionTitle>
      <div className="flex flex-col gap-2">
        {suppliers.length === 0 && <EmptyState T={T} text="Sin proveedores todavía." />}
        {suppliers.map((s) => (
          <Card T={T} key={s.id} style={{ padding: "10px 16px" }} onClick={() => setForm(s)} draggable onDragStart={(e) => { e.stopPropagation(); dragRef.current = s.id; }} onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.stopPropagation(); setSuppliers((ss) => reorderById(ss, dragRef.current, s.id)); }}>
            <div className="flex items-center justify-between text-sm">
              <div className="flex items-center gap-2"><GripVertical size={13} color={T.muted} style={{ cursor: "grab" }} /><span className="font-medium">{s.company}</span><span className="ml-2 text-xs" style={{ color: T.muted }}>{s.category}</span></div>
              <div className="flex items-center gap-4 text-xs" style={{ color: T.muted }}>
                {s.phone && <span className="flex items-center gap-1"><Phone size={11} />{s.phone}</span>}
                <span className="px-2 py-1 rounded-full" style={{ background: s.paymentStatus === "Pagado" ? T.cardAlt : "transparent", color: s.paymentStatus === "Pagado" ? T.secondaryDark : T.alert, border: s.paymentStatus === "Pagado" ? "none" : `1px solid ${T.alert}` }}>{s.paymentStatus}</span>
                <span>{fmtMoney(s.cost, "CLP")}</span>
                <button onClick={(e) => { e.stopPropagation(); setSuppliers((ss) => ss.filter((x) => x.id !== s.id)); }}><Trash2 size={14} /></button>
              </div>
            </div>
          </Card>
        ))}
      </div>

      {form && (
        <Modal T={T} title={form.company ? "Editar proveedor" : "Nuevo proveedor"} onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <Field T={T} label="Compañía"><input style={inputStyle(T)} value={form.company} onChange={(e) => setForm({ ...form, company: e.target.value })} autoFocus /></Field>
            <Field T={T} label="Categoría de presupuesto"><select style={inputStyle(T)} value={form.category} onChange={(e) => setForm({ ...form, category: e.target.value })}>{cats.map((c) => <option key={c}>{c}</option>)}</select></Field>
            <Field T={T} label="Costo total"><input type="number" style={inputStyle(T)} value={form.cost} onChange={(e) => setForm({ ...form, cost: e.target.value })} /></Field>
            <Field T={T} label="Estado de pago"><select style={inputStyle(T)} value={form.paymentStatus} onChange={(e) => setForm({ ...form, paymentStatus: e.target.value })}><option>No pagado</option><option>Parcial</option><option>Pagado</option></select></Field>
            <Field T={T} label="Teléfono"><input style={inputStyle(T)} value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })} /></Field>
            <Field T={T} label="Email"><input style={inputStyle(T)} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field>
            <div className="col-span-2"><Field T={T} label="Dirección"><input style={{ ...inputStyle(T), width: "100%" }} value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })} /></Field></div>
            <div className="flex items-center gap-3 mt-5"><Toggle T={T} checked={form.depositPaid} onChange={(v) => setForm({ ...form, depositPaid: v })} /><span className="text-sm">Reserva pagada</span></div>
            <div className="col-span-2"><Field T={T} label="Notas"><textarea style={{ ...inputStyle(T), width: "100%", minHeight: 60 }} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></Field></div>
            <div className="col-span-2 flex gap-2 items-center pt-2" style={{ borderTop: `1px solid ${T.line}` }}>
              <input placeholder="Nueva categoría de presupuesto" style={{ ...inputStyle(T), flex: 1 }} value={newCat} onChange={(e) => setNewCat(e.target.value)} />
              <Button T={T} small variant="outline" onClick={() => { if (newCat.trim()) { setCats((c) => [...c, newCat]); setNewCat(""); } }}><Plus size={13} /> Agregar categoría</Button>
            </div>
          </div>
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== Catering ============================== */
function Catering({ T, catering, setCatering, activeEvent, events }) {
  const evId = pickEvent(activeEvent, events);
  const [inputs, setInputs] = useState({});
  const dragRef = useRef(null);
  const data = catering[evId] || {};
  function addItem(section) {
    const val = (inputs[section] || "").trim();
    if (!val) return;
    setCatering((c) => ({ ...c, [evId]: { ...c[evId], [section]: [...(c[evId]?.[section] || []), { id: uid(), name: val }] } }));
    setInputs((i) => ({ ...i, [section]: "" }));
  }
  function removeItem(section, id) { setCatering((c) => ({ ...c, [evId]: { ...c[evId], [section]: c[evId][section].filter((x) => x.id !== id) } })); }
  function importCatering(rows) {
    const grouped = {};
    rows.forEach((row) => {
      const secRaw = pick(row, ["Seccion", "Sección", "Section"]);
      const name = pick(row, ["Nombre", "Name"]);
      const sec = CATERING_SECTIONS.find((s) => s.toLowerCase() === String(secRaw).trim().toLowerCase());
      if (!sec || !name) return;
      grouped[sec] = grouped[sec] || [];
      grouped[sec].push({ id: uid(), name });
    });
    if (Object.keys(grouped).length) {
      setCatering((c) => {
        const cur = c[evId] || {};
        const merged = { ...cur };
        Object.keys(grouped).forEach((sec) => { merged[sec] = [...(cur[sec] || []), ...grouped[sec]]; });
        return { ...c, [evId]: merged };
      });
    }
  }

  return (
    <div>
      <SectionTitle T={T} action={<ImportExcel T={T} label="Importar Excel" hint="Columnas: Seccion (Entrada/Plato principal/Postre/Cóctail), Nombre" onRows={importCatering} />}>
        Catering{activeEvent === "all" && <span className="text-sm ml-2" style={{ color: T.muted }}>· editando {events.find((e) => e.id === evId)?.name}</span>}
      </SectionTitle>
      <div className="grid grid-cols-2 gap-4">
        {CATERING_SECTIONS.map((sec) => (
          <Card T={T} key={sec}>
            <div className="text-sm font-medium mb-3 flex items-center gap-2"><Utensils size={14} color={T.primary} />{sec}</div>
            <div className="flex flex-col gap-1.5 mb-3">
              {(data[sec] || []).map((it) => <div key={it.id} className="flex items-center justify-between text-sm px-3 py-2 rounded-lg" style={{ background: T.cardAlt }} draggable onDragStart={() => (dragRef.current = it.id)} onDragOver={(e) => e.preventDefault()} onDrop={() => setCatering((c) => ({ ...c, [evId]: { ...c[evId], [sec]: reorderById(c[evId][sec], dragRef.current, it.id) } }))}><span className="flex items-center gap-2"><GripVertical size={12} color={T.muted} style={{ cursor: "grab" }} />{it.name}</span><button onClick={() => removeItem(sec, it.id)}><X size={12} color={T.muted} /></button></div>)}
              {(!data[sec] || data[sec].length === 0) && <span className="text-xs" style={{ color: T.muted }}>Sin opciones aún</span>}
            </div>
            <div className="flex gap-2">
              <input placeholder={`Agregar a ${sec.toLowerCase()}`} style={{ ...inputStyle(T), flex: 1 }} value={inputs[sec] || ""} onChange={(e) => setInputs({ ...inputs, [sec]: e.target.value })} onKeyDown={(e) => e.key === "Enter" && addItem(sec)} />
              <Button T={T} small variant="outline" onClick={() => addItem(sec)}><Plus size={13} /></Button>
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ============================== Invitaciones ============================== */
function Invitations({ T, invitations, setInvitations, events, activeEvent, guests, setGuests, groups, scheme }) {
  const [form, setForm] = useState(null);
  function startNew() { setForm({ id: uid(), tipo: "Save the date", eventId: pickEvent(activeEvent, events), diseno: scheme?.name || "", fechaEnvio: "", groupFilter: "Todos", sent: false }); }
  function targetGuests(inv) {
    return guests.filter((g) => g.eventIds.includes(inv.eventId) && (inv.groupFilter === "Todos" || g.group === inv.groupFilter));
  }
  function save() { setInvitations((iv) => (iv.some((x) => x.id === form.id) ? iv.map((x) => (x.id === form.id ? form : x)) : [...iv, form])); setForm(null); }
  function markSent(inv) {
    const ids = new Set(targetGuests(inv).map((g) => g.id));
    setGuests((gs) => gs.map((g) => (ids.has(g.id) ? { ...g, inviteSent: true } : g)));
    setInvitations((iv) => iv.map((x) => (x.id === inv.id ? { ...x, sent: true } : x)));
  }

  return (
    <div>
      <SectionTitle T={T} action={<Button T={T} onClick={startNew}><Plus size={15} /> Nuevo envío</Button>}>Invitaciones</SectionTitle>
      <div className="flex flex-col gap-3">
        {invitations.length === 0 && <EmptyState T={T} text="Sin campañas de invitación todavía." />}
        {invitations.map((inv) => {
          const list = targetGuests(inv);
          return (
            <Card T={T} key={inv.id}>
              <div className="flex items-center justify-between mb-2">
                <div>
                  <div className="text-sm font-medium">{inv.tipo} · {events.find((e) => e.id === inv.eventId)?.name}</div>
                  <div className="text-xs mt-1" style={{ color: T.muted }}>Diseño: {inv.diseno || "—"} · Envío: {inv.fechaEnvio || "sin definir"} · Destinatarios: {inv.groupFilter}</div>
                </div>
                <div className="flex gap-2 items-center">
                  <span className="text-xs" style={{ color: T.muted }}>{list.length} invitados</span>
                  <Button T={T} small variant="outline" onClick={() => setForm(inv)}>Editar</Button>
                  {!inv.sent && <Button T={T} small onClick={() => markSent(inv)}>Marcar enviada</Button>}
                  {inv.sent && <span className="text-xs px-2 py-1 rounded-full" style={{ background: T.cardAlt, color: T.secondaryDark }}>Enviada</span>}
                </div>
              </div>
            </Card>
          );
        })}
      </div>

      {form && (
        <Modal T={T} title="Campaña de invitación" onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <Field T={T} label="Tipo"><select style={inputStyle(T)} value={form.tipo} onChange={(e) => setForm({ ...form, tipo: e.target.value })}><option>Save the date</option><option>Invitación formal</option></select></Field>
            <Field T={T} label="Evento"><select style={inputStyle(T)} value={form.eventId} onChange={(e) => setForm({ ...form, eventId: e.target.value })}>{events.map((ev) => <option key={ev.id} value={ev.id}>{ev.name}</option>)}</select></Field>
            <Field T={T} label="Diseño / paleta"><input style={inputStyle(T)} value={form.diseno} onChange={(e) => setForm({ ...form, diseno: e.target.value })} /></Field>
            <Field T={T} label="Fecha de envío"><input type="date" style={inputStyle(T)} value={form.fechaEnvio} onChange={(e) => setForm({ ...form, fechaEnvio: e.target.value })} /></Field>
            <Field T={T} label="Listado de invitados"><select style={inputStyle(T)} value={form.groupFilter} onChange={(e) => setForm({ ...form, groupFilter: e.target.value })}><option>Todos</option>{groups.map((g) => <option key={g}>{g}</option>)}</select></Field>
          </div>
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== Seating ============================== */
function Seating({ T, seating, setSeating, activeEvent, events, guests }) {
  const evId = pickEvent(activeEvent, events);
  const tables = seating[evId] || [];
  const eventGuests = guests.filter((g) => g.eventIds.includes(evId));
  const assignedIds = new Set(tables.flatMap((t) => t.guestIds));
  const unassigned = eventGuests.filter((g) => !assignedIds.has(g.id));
  const dragId = useRef(null);
  const canvasRef = useRef(null);

  function addTable(shape) {
    setSeating((s) => ({ ...s, [evId]: [...(s[evId] || []), { id: uid(), name: `Mesa ${((s[evId] || []).length) + 1}`, shape, seats: shape === "redonda" ? 8 : 10, guestIds: [], x: 40 + (((s[evId] || []).length) % 4) * 160, y: 40 + Math.floor(((s[evId] || []).length) / 4) * 160 }] }));
  }
  function updateTable(id, patch) { setSeating((s) => ({ ...s, [evId]: s[evId].map((t) => (t.id === id ? { ...t, ...patch } : t)) })); }
  function removeTable(id) { setSeating((s) => ({ ...s, [evId]: s[evId].filter((t) => t.id !== id) })); }
  function assign(tableId, guestId) { if (!guestId) return; updateTable(tableId, { guestIds: [...tables.find((t) => t.id === tableId).guestIds, guestId] }); }
  function unassign(tableId, guestId) { updateTable(tableId, { guestIds: tables.find((t) => t.id === tableId).guestIds.filter((id) => id !== guestId) }); }
  function onDrop(e) {
    e.preventDefault();
    const id = dragId.current;
    if (!id) return;
    const rect = canvasRef.current.getBoundingClientRect();
    updateTable(id, { x: Math.max(0, e.clientX - rect.left - 60), y: Math.max(0, e.clientY - rect.top - 60) });
  }
  function printSeating() {
    const rows = tables.map((t) => `<tr><td>${t.name}</td><td>${t.shape}</td><td>${t.guestIds.length}/${t.seats}</td><td>${t.guestIds.map((id) => guests.find((g) => g.id === id)?.name || "").join(", ")}</td></tr>`).join("");
    printHtml("Plano de mesas", `<h1>Plano de mesas · ${events.find((e) => e.id === evId)?.name || ""}</h1><table><tr><th>Mesa</th><th>Forma</th><th>Ocupación</th><th>Invitados</th></tr>${rows}</table>`);
  }

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="flex gap-2">
          <Button T={T} small variant="outline" onClick={printSeating}><Printer size={13} /> Imprimir / PDF</Button>
          <Button T={T} small variant="outline" onClick={() => addTable("redonda")}><Plus size={13} /> Mesa redonda</Button>
          <Button T={T} small onClick={() => addTable("rectangular")}><Plus size={13} /> Mesa rectangular</Button>
        </div>}>
        Plano de mesas {activeEvent === "all" && <span className="text-sm ml-2" style={{ color: T.muted }}>· editando {events.find((e) => e.id === evId)?.name}</span>}
      </SectionTitle>

      <Card T={T} style={{ marginBottom: 16 }}>
        <div className="text-sm font-medium mb-1">Sin asignar: {unassigned.length}</div>
        <div className="flex flex-wrap gap-2">
          {unassigned.length === 0 && <span className="text-xs" style={{ color: T.muted }}>{eventGuests.length === 0 ? "Aún no hay invitados asignados a esta ceremonia — ve a Invitados." : "Todos los invitados de esta ceremonia ya tienen mesa."}</span>}
          {unassigned.map((g) => <span key={g.id} className="text-xs px-2.5 py-1.5 rounded-full" style={{ background: T.cardAlt, color: T.text }}>{g.name}</span>)}
        </div>
      </Card>

      <div ref={canvasRef} onDragOver={(e) => e.preventDefault()} onDrop={onDrop} className="relative rounded-2xl mb-4" style={{ height: 420, background: T.cardAlt, border: `1px dashed ${T.line}` }}>
        {tables.length === 0 && <div className="absolute inset-0 flex items-center justify-center text-sm" style={{ color: T.muted }}>Agrega una mesa y arrástrala al plano</div>}
        {tables.map((t) => (
          <div key={t.id} draggable onDragStart={() => (dragId.current = t.id)}
            className="absolute flex flex-col items-center justify-center text-xs p-2 cursor-move"
            style={{ left: t.x, top: t.y, width: 120, height: t.shape === "redonda" ? 120 : 90, borderRadius: t.shape === "redonda" ? "50%" : 14, background: T.card, border: `2px solid ${T.primary}` }}>
            <input value={t.name} onChange={(e) => updateTable(t.id, { name: e.target.value })} className="bg-transparent text-center font-medium outline-none w-full" style={{ color: T.text }} />
            <span style={{ color: T.muted }}>{t.guestIds.length}/{t.seats} puestos</span>
            <button onClick={() => removeTable(t.id)} className="absolute -top-2 -right-2 w-5 h-5 rounded-full flex items-center justify-center" style={{ background: T.alert }}><X size={11} color="#fff" /></button>
          </div>
        ))}
      </div>

      <div className="grid grid-cols-2 gap-4">
        {tables.map((t) => (
          <Card T={T} key={t.id}>
            <div className="flex items-center justify-between mb-2 text-sm font-medium">{t.name}
              <input type="number" value={t.seats} onChange={(e) => updateTable(t.id, { seats: Number(e.target.value) })} style={{ ...inputStyle(T), width: 60 }} />
            </div>
            <div className="flex flex-col gap-1.5 mb-2">
              {t.guestIds.map((gid) => { const g = guests.find((x) => x.id === gid); return (
                <div key={gid} className="flex items-center justify-between text-xs px-2.5 py-1.5 rounded-lg" style={{ background: T.cardAlt }}>{g?.name || "—"}<button onClick={() => unassign(t.id, gid)}><X size={12} color={T.muted} /></button></div>
              ); })}
              {t.guestIds.length === 0 && <span className="text-xs" style={{ color: T.muted }}>Sin invitados en esta mesa</span>}
            </div>
            <select style={{ ...inputStyle(T), width: "100%" }} value="" onChange={(e) => assign(t.id, e.target.value)}>
              <option value="">{unassigned.length ? "+ Asignar invitado" : "No hay invitados sin asignar"}</option>
              {unassigned.map((g) => <option key={g.id} value={g.id}>{g.name}</option>)}
            </select>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ============================== Schedule ============================== */
function Schedule({ T, schedules, setSchedules, activeEvent, events, suppliers }) {
  const evId = pickEvent(activeEvent, events);
  const items = (schedules[evId] || []).slice().sort((a, b) => (a.start || "").localeCompare(b.start || ""));
  const [form, setForm] = useState(null);
  function startNew() { setForm({ id: uid(), title: "", start: "", end: "", runsUntilEnd: false, where: "", withSupplier: "", remindMe: false, notes: "" }); }
  function save() {
    if (!form.title.trim()) return;
    setSchedules((s) => {
      const list = s[evId] || [];
      const exists = list.some((x) => x.id === form.id);
      return { ...s, [evId]: exists ? list.map((x) => (x.id === form.id ? form : x)) : [...list, form] };
    });
    setForm(null);
  }
  function remove(id) { setSchedules((s) => ({ ...s, [evId]: s[evId].filter((x) => x.id !== id) })); }
  function importSchedule(rows) {
    const created = rows.map((row) => {
      const title = pick(row, ["Titulo", "Título", "Title"]);
      if (!title) return null;
      const end = pick(row, ["HoraTermino", "Hora término", "End"]);
      return { id: uid(), title, start: pick(row, ["Hora", "HoraInicio", "Start"]), end, runsUntilEnd: !!end, where: pick(row, ["Ubicacion", "Ubicación", "Location"]), withSupplier: pick(row, ["Proveedor", "Supplier"]), remindMe: truthy(pick(row, ["Recordar", "RemindMe"])), notes: pick(row, ["Notas", "Notes"]) };
    }).filter(Boolean);
    if (created.length) setSchedules((s) => ({ ...s, [evId]: [...(s[evId] || []), ...created] }));
  }
  function printSchedule() {
    const rows = items.map((it) => `<tr><td>${it.start}${it.runsUntilEnd && it.end ? " – " + it.end : ""}</td><td>${it.title}</td><td>${it.where || ""}</td><td>${it.withSupplier || ""}</td></tr>`).join("");
    printHtml("Cronograma", `<h1>Cronograma · ${events.find((e) => e.id === evId)?.name || ""}</h1><table><tr><th>Hora</th><th>Ítem</th><th>Ubicación</th><th>Proveedor</th></tr>${rows}</table>`);
  }

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="flex gap-2">
          <Button T={T} small variant="outline" onClick={printSchedule}><Printer size={13} /> Imprimir / PDF</Button>
          <ImportExcel T={T} label="Importar Excel" hint="Columnas: Titulo, Hora, HoraTermino, Ubicacion, Proveedor, Notas" onRows={importSchedule} />
          <Button T={T} onClick={startNew}><Plus size={15} /> Agregar ítem</Button>
        </div>}>
        Cronograma {activeEvent === "all" && <span className="text-sm ml-2" style={{ color: T.muted }}>· editando {events.find((e) => e.id === evId)?.name}</span>}
      </SectionTitle>
      <div className="flex flex-col">
        {items.length === 0 && <EmptyState T={T} text="Sin ítems en el cronograma todavía." />}
        {items.map((it) => (
          <div key={it.id} className="flex items-center gap-4 py-3 cursor-pointer" style={{ borderLeft: `2px solid ${T.soft}`, paddingLeft: 16, marginLeft: 4 }} onClick={() => setForm(it)}>
            <span className="text-sm font-medium w-24" style={{ color: T.primaryDark }}>{it.start}{it.runsUntilEnd && it.end ? ` – ${it.end}` : ""}</span>
            <div className="flex-1">
              <div className="text-sm">{it.title}</div>
              {(it.where || it.withSupplier) && (
                <div className="text-xs flex items-center gap-1" style={{ color: T.muted }}>
                  {it.where && <a href={mapsLink(it.where)} target="_blank" rel="noreferrer" onClick={(e) => e.stopPropagation()} className="flex items-center gap-1" style={{ color: T.primaryDark }}><MapPin size={11} />{it.where}</a>}
                  {it.where && it.withSupplier ? " · " : ""}{it.withSupplier}
                </div>
              )}
            </div>
            {it.remindMe && <span className="text-xs" style={{ color: T.secondary }}>Recordatorio</span>}
            <button onClick={(e) => { e.stopPropagation(); remove(it.id); }}><Trash2 size={14} color={T.muted} /></button>
          </div>
        ))}
      </div>

      {form && (
        <Modal T={T} title={form.title ? "Editar ítem" : "Nuevo ítem"} onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2"><Field T={T} label="Título"><input style={{ ...inputStyle(T), width: "100%" }} value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} autoFocus /></Field></div>
            <Field T={T} label="Hora de inicio"><input type="time" style={inputStyle(T)} value={form.start} onChange={(e) => setForm({ ...form, start: e.target.value })} /></Field>
            <div className="flex items-center gap-3 mt-5"><Toggle T={T} checked={form.runsUntilEnd} onChange={(v) => setForm({ ...form, runsUntilEnd: v })} /><span className="text-sm">Tiene hora de término</span></div>
            {form.runsUntilEnd && <Field T={T} label="Hora de término"><input type="time" style={inputStyle(T)} value={form.end} onChange={(e) => setForm({ ...form, end: e.target.value })} /></Field>}
            <Field T={T} label="Ubicación"><input style={inputStyle(T)} value={form.where} onChange={(e) => setForm({ ...form, where: e.target.value })} placeholder="Dirección o lugar" /></Field>
            <Field T={T} label="Con (proveedor)"><select style={inputStyle(T)} value={form.withSupplier} onChange={(e) => setForm({ ...form, withSupplier: e.target.value })}><option value="">Sin proveedor</option>{suppliers.map((s) => <option key={s.id} value={s.company}>{s.company}</option>)}</select></Field>
            <div className="flex items-center gap-3 mt-5"><Toggle T={T} checked={form.remindMe} onChange={(v) => setForm({ ...form, remindMe: v })} /><span className="text-sm">Recuérdamelo</span></div>
            <div className="col-span-2"><Field T={T} label="Notas"><textarea style={{ ...inputStyle(T), width: "100%", minHeight: 60 }} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} /></Field></div>
          </div>
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== Scrapbook ============================== */
const SCRAP_TYPES = [
  { id: "note", label: "Nota de texto", icon: StickyNote }, { id: "link", label: "Link", icon: Link2 },
  { id: "checklist", label: "Checklist", icon: CheckSquare }, { id: "sketch", label: "Sketch", icon: PenTool },
  { id: "voice", label: "Nota de voz", icon: Mic }, { id: "photo", label: "Foto", icon: ImageIcon },
];
function Scrapbook({ T, items, setItems }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const fileRef = useRef(null);
  function addItem(type) { setMenuOpen(false); if (type === "photo") { fileRef.current?.click(); return; } setItems((its) => [...its, { id: uid(), type, text: "", checklist: [] }]); }
  function handlePhoto(e) { const f = e.target.files?.[0]; if (!f) return; compressImage(f).then((dataUrl) => setItems((its) => [...its, { id: uid(), type: "photo", text: dataUrl }])).catch(() => alert("No se pudo procesar esa imagen.")); }
  function update(id, patch) { setItems((its) => its.map((it) => (it.id === id ? { ...it, ...patch } : it))); }
  function remove(id) { setItems((its) => its.filter((it) => it.id !== id)); }

  return (
    <div>
      <SectionTitle T={T} action={
        <div className="relative">
          <Button T={T} onClick={() => setMenuOpen((v) => !v)}><Plus size={15} /> Agregar</Button>
          {menuOpen && (
            <div className="absolute right-0 mt-2 rounded-xl overflow-hidden z-10" style={{ background: T.card, border: `1px solid ${T.line}` }}>
              {SCRAP_TYPES.map((s) => { const Icon = s.icon; return <button key={s.id} onClick={() => addItem(s.id)} className="flex items-center gap-2 px-4 py-2.5 text-sm w-full text-left" style={{ color: T.text }}><Icon size={14} color={T.primary} />{s.label}</button>; })}
            </div>
          )}
          <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={handlePhoto} />
        </div>}>
        Scrapbook
      </SectionTitle>
      <div className="flex flex-wrap gap-4">
        {items.length === 0 && <EmptyState T={T} text="El scrapbook está vacío — agrega notas, fotos o ideas." />}
        {items.map((it) => (
          <div key={it.id} className="relative w-48 p-3 rounded-lg" style={{ background: it.type === "note" ? "#F6E79C" : T.card, border: it.type === "note" ? "none" : `1px solid ${T.line}`, boxShadow: "0 3px 8px rgba(0,0,0,0.12)", transform: "rotate(-1deg)" }}>
            <button onClick={() => remove(it.id)} className="absolute top-1 right-1"><X size={12} color={it.type === "note" ? "#6b5c1f" : T.muted} /></button>
            {it.type === "note" && <textarea value={it.text} onChange={(e) => update(it.id, { text: e.target.value })} placeholder="Escribe algo…" className="bg-transparent w-full text-sm outline-none" style={{ color: "#4a3d10", minHeight: 80 }} />}
            {it.type === "link" && <input value={it.text} onChange={(e) => update(it.id, { text: e.target.value })} placeholder="https://…" className="bg-transparent w-full text-sm outline-none" style={{ color: T.primaryDark }} />}
            {it.type === "photo" && <img src={it.text} alt="" className="w-full h-32 object-cover rounded" />}
            {it.type === "sketch" && <div className="h-24 flex items-center justify-center text-xs rounded" style={{ background: T.cardAlt, color: T.muted }}><PenTool size={16} className="mr-1" /> Sketch</div>}
            {it.type === "voice" && <div className="h-16 flex items-center justify-center text-xs rounded gap-2" style={{ background: T.cardAlt, color: T.muted }}><Mic size={16} /> Nota de voz</div>}
            {it.type === "checklist" && <ChecklistScrap T={T} item={it} onChange={(patch) => update(it.id, patch)} />}
          </div>
        ))}
      </div>
    </div>
  );
}
function ChecklistScrap({ T, item, onChange }) {
  const [val, setVal] = useState("");
  return (
    <div className="flex flex-col gap-1 text-sm" style={{ color: T.text }}>
      {item.checklist.map((c) => (
        <label key={c.id} className="flex items-center gap-1.5">
          <input type="checkbox" checked={c.done} onChange={() => onChange({ checklist: item.checklist.map((x) => (x.id === c.id ? { ...x, done: !x.done } : x)) })} />
          <span style={{ textDecoration: c.done ? "line-through" : "none" }}>{c.text}</span>
        </label>
      ))}
      <input placeholder="+ agregar" value={val} onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => { if (e.key === "Enter" && val.trim()) { onChange({ checklist: [...item.checklist, { id: uid(), text: val, done: false }] }); setVal(""); } }}
        className="bg-transparent text-xs outline-none mt-1" style={{ color: T.muted }} />
    </div>
  );
}

/* ============================== Inspiración ============================== */
const PIN_TYPES = [
  { id: "photo", label: "Foto", icon: ImageIcon }, { id: "link", label: "Link", icon: Link2 },
  { id: "note", label: "Nota", icon: StickyNote }, { id: "color", label: "Color", icon: Palette },
];
function Inspiration({ T, boards, setBoards, events }) {
  const [groupBy, setGroupBy] = useState("Todos");
  const [subFilter, setSubFilter] = useState(null);
  const [form, setForm] = useState(null);
  const [openBoard, setOpenBoard] = useState(null);

  function startNew() { setForm({ id: uid(), name: "", groupType: "Macro evento", eventId: events[0]?.id || "", moduleTag: MODULE_TAGS[0], items: [] }); }
  function save() { if (!form.name.trim()) return; setBoards((b) => (b.some((x) => x.id === form.id) ? b.map((x) => (x.id === form.id ? form : x)) : [...b, form])); setForm(null); }
  function removeBoard(id) { setBoards((b) => b.filter((x) => x.id !== id)); if (openBoard === id) setOpenBoard(null); }

  const filtered = boards.filter((b) => {
    if (groupBy === "Todos") return true;
    if (groupBy === "Macro evento") return b.groupType === "Macro evento";
    if (groupBy === "Evento") return b.groupType === "Evento" && (!subFilter || b.eventId === subFilter);
    if (groupBy === "Módulo") return b.groupType === "Módulo" && (!subFilter || b.moduleTag === subFilter);
    return true;
  });

  return (
    <div>
      <SectionTitle T={T} action={<Button T={T} onClick={startNew}><Plus size={15} /> Nuevo tablero</Button>}>Inspiración</SectionTitle>
      <p className="text-xs mb-4" style={{ color: T.muted }}>Agrupa tableros por evento, por tipo de módulo, o déjalos como inspiración general de la boda (macro evento).</p>

      <div className="flex gap-2 mb-3 flex-wrap">
        {["Todos", "Macro evento", "Evento", "Módulo"].map((g) => (
          <button key={g} onClick={() => { setGroupBy(g); setSubFilter(null); }} className="px-3.5 py-1.5 rounded-full text-xs font-medium"
            style={groupBy === g ? { background: T.secondary, color: "#fff" } : { background: T.cardAlt, color: T.muted, border: `1px solid ${T.line}` }}>{g}</button>
        ))}
      </div>
      {groupBy === "Evento" && (
        <div className="flex gap-2 mb-5 flex-wrap">
          {events.map((ev) => (
            <button key={ev.id} onClick={() => setSubFilter(subFilter === ev.id ? null : ev.id)} className="px-3 py-1 rounded-full text-xs"
              style={subFilter === ev.id ? { background: T.primary, color: "#fff" } : { background: T.cardAlt, color: T.muted, border: `1px solid ${T.line}` }}>{ev.name}</button>
          ))}
        </div>
      )}
      {groupBy === "Módulo" && (
        <div className="flex gap-2 mb-5 flex-wrap">
          {MODULE_TAGS.map((m) => (
            <button key={m} onClick={() => setSubFilter(subFilter === m ? null : m)} className="px-3 py-1 rounded-full text-xs"
              style={subFilter === m ? { background: T.primary, color: "#fff" } : { background: T.cardAlt, color: T.muted, border: `1px solid ${T.line}` }}>{m}</button>
          ))}
        </div>
      )}

      <div className="grid grid-cols-3 gap-4">
        {filtered.length === 0 && <EmptyState T={T} text="Sin tableros para este filtro todavía." />}
        {filtered.map((b) => (
          <Card T={T} key={b.id} onClick={() => setOpenBoard(openBoard === b.id ? null : b.id)}>
            <div className="flex items-start justify-between mb-2">
              <div>
                <div className="text-sm font-medium">{b.name}</div>
                <div className="text-xs mt-1" style={{ color: T.muted }}>
                  {b.groupType === "Evento" ? events.find((e) => e.id === b.eventId)?.name : b.groupType === "Módulo" ? b.moduleTag : "General"}
                </div>
              </div>
              <button onClick={(e) => { e.stopPropagation(); setForm(b); }}><Tag size={13} color={T.muted} /></button>
            </div>
            <div className="flex gap-1.5">
              {b.items.slice(0, 4).map((it) => {
                const PinIcon = (PIN_TYPES.find((p) => p.id === it.type) || {}).icon;
                return (
                  <div key={it.id} className="w-9 h-9 rounded-lg overflow-hidden flex items-center justify-center" style={{ background: it.type === "color" ? it.content : T.cardAlt }}>
                    {it.type === "photo" && <img src={it.content} alt="" className="w-full h-full object-cover" />}
                    {it.type !== "photo" && it.type !== "color" && PinIcon && <PinIcon size={13} color={T.muted} />}
                  </div>
                );
              })}
              {b.items.length === 0 && <span className="text-xs" style={{ color: T.muted }}>Vacío</span>}
            </div>
          </Card>
        ))}
      </div>

      {openBoard && boards.find((b) => b.id === openBoard) && (
        <BoardDetail T={T} board={boards.find((b) => b.id === openBoard)} setBoards={setBoards} onClose={() => setOpenBoard(null)} onDelete={() => removeBoard(openBoard)} />
      )}

      {form && (
        <Modal T={T} title={boards.some((x) => x.id === form.id) ? "Editar tablero" : "Nuevo tablero"} onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <div className="col-span-2"><Field T={T} label="Nombre"><input style={{ ...inputStyle(T), width: "100%" }} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} autoFocus placeholder="Ej. Vestido, Flores, Torta…" /></Field></div>
            <Field T={T} label="Agrupar por"><select style={inputStyle(T)} value={form.groupType} onChange={(e) => setForm({ ...form, groupType: e.target.value })}><option>Macro evento</option><option>Evento</option><option>Módulo</option></select></Field>
            {form.groupType === "Evento" && <Field T={T} label="Evento"><select style={inputStyle(T)} value={form.eventId} onChange={(e) => setForm({ ...form, eventId: e.target.value })}>{events.map((ev) => <option key={ev.id} value={ev.id}>{ev.name}</option>)}</select></Field>}
            {form.groupType === "Módulo" && <Field T={T} label="Módulo"><select style={inputStyle(T)} value={form.moduleTag} onChange={(e) => setForm({ ...form, moduleTag: e.target.value })}>{MODULE_TAGS.map((m) => <option key={m}>{m}</option>)}</select></Field>}
          </div>
          <div className="flex gap-2 mt-6 justify-end">
            {boards.some((x) => x.id === form.id) && <Button T={T} variant="outline" onClick={() => { removeBoard(form.id); setForm(null); }}><Trash2 size={13} /> Eliminar</Button>}
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}
function BoardDetail({ T, board, setBoards, onClose, onDelete }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const fileRef = useRef(null);
  function patch(items) { setBoards((bs) => bs.map((b) => (b.id === board.id ? { ...b, items } : b))); }
  function addPin(type) {
    setMenuOpen(false);
    if (type === "photo") { fileRef.current?.click(); return; }
    patch([...board.items, { id: uid(), type, content: type === "color" ? "#B5652E" : "" }]);
  }
  function handlePhoto(e) { const f = e.target.files?.[0]; if (!f) return; compressImage(f).then((dataUrl) => patch([...board.items, { id: uid(), type: "photo", content: dataUrl }])).catch(() => alert("No se pudo procesar esa imagen.")); }
  function update(id, content) { patch(board.items.map((it) => (it.id === id ? { ...it, content } : it))); }
  function remove(id) { patch(board.items.filter((it) => it.id !== id)); }

  return (
    <Modal T={T} title={board.name} onClose={onClose}>
      <div className="flex justify-between items-center mb-4">
        <span className="text-xs" style={{ color: T.muted }}>{board.items.length} elementos</span>
        <div className="relative flex gap-2">
          <button onClick={onDelete} className="text-xs flex items-center gap-1" style={{ color: T.alert }}><Trash2 size={13} /> Eliminar tablero</button>
          <Button T={T} small onClick={() => setMenuOpen((v) => !v)}><Plus size={13} /> Agregar</Button>
          {menuOpen && (
            <div className="absolute right-0 top-8 rounded-xl overflow-hidden z-10" style={{ background: T.card, border: `1px solid ${T.line}` }}>
              {PIN_TYPES.map((p) => { const Icon = p.icon; return <button key={p.id} onClick={() => addPin(p.id)} className="flex items-center gap-2 px-4 py-2.5 text-sm w-full text-left" style={{ color: T.text }}><Icon size={14} color={T.primary} />{p.label}</button>; })}
            </div>
          )}
          <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={handlePhoto} />
        </div>
      </div>
      <div className="flex flex-wrap gap-3">
        {board.items.length === 0 && <EmptyState T={T} text="Este tablero está vacío — agrega fotos, links, notas o colores." />}
        {board.items.map((it) => (
          <div key={it.id} className="relative w-36 rounded-lg overflow-hidden" style={{ border: `1px solid ${T.line}` }}>
            <button onClick={() => remove(it.id)} className="absolute top-1 right-1 z-10 w-5 h-5 rounded-full flex items-center justify-center" style={{ background: "rgba(0,0,0,0.5)" }}><X size={11} color="#fff" /></button>
            {it.type === "photo" && <img src={it.content} alt="" className="w-full h-32 object-cover" />}
            {it.type === "color" && <input type="color" value={it.content} onChange={(e) => update(it.id, e.target.value)} className="w-full h-32 border-none" />}
            {it.type === "link" && <input value={it.content} onChange={(e) => update(it.id, e.target.value)} placeholder="https://…" className="w-full h-32 text-xs p-2 outline-none" style={{ background: T.cardAlt, color: T.primaryDark }} />}
            {it.type === "note" && <textarea value={it.content} onChange={(e) => update(it.id, e.target.value)} placeholder="Idea…" className="w-full h-32 text-xs p-2 outline-none" style={{ background: "#F6E79C", color: "#4a3d10" }} />}
          </div>
        ))}
      </div>
    </Modal>
  );
}

/* ============================== Esquemas de color ============================== */
function Schemes({ T, schemes, setSchemes, activeScheme, setActiveScheme, mode }) {
  const [form, setForm] = useState(null);
  function startNew() { setForm({ id: uid(), name: "", colors: ["#B5652E", "#767A4A", "#CDA36C", "#7C3140"] }); }
  function save() { if (!form.name.trim()) return; setSchemes((s) => [...s, form]); setForm(null); }
  return (
    <div>
      <SectionTitle T={T} action={<Button T={T} onClick={startNew}><Plus size={15} /> Nueva paleta</Button>}>Esquemas de color</SectionTitle>
      {mode === "dark" && <p className="text-xs mb-4" style={{ color: T.muted }}>El modo oscuro usa un acento rosado fijo — las paletas de abajo se aplican al modo día.</p>}
      <div className="grid grid-cols-3 gap-4">
        {schemes.map((s) => (
          <Card T={T} key={s.id} style={{ border: activeScheme === s.id ? `2px solid ${T.primary}` : `1px solid ${T.line}` }}>
            <div className="text-sm font-medium mb-3">{s.name}</div>
            <div className="flex gap-1.5 mb-3">{s.colors.map((c, i) => <div key={i} className="w-9 h-9 rounded-lg" style={{ background: c }} />)}</div>
            <div className="flex gap-2">
              <Button T={T} small onClick={() => setActiveScheme(s.id)}>{activeScheme === s.id ? "En uso" : "Usar esta paleta"}</Button>
              {schemes.length > 1 && <button onClick={() => setSchemes((all) => all.filter((x) => x.id !== s.id))}><Trash2 size={14} color={T.muted} /></button>}
            </div>
          </Card>
        ))}
      </div>
      {form && (
        <Modal T={T} title="Nueva paleta" onClose={() => setForm(null)}>
          <Field T={T} label="Nombre"><input style={inputStyle(T)} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} autoFocus /></Field>
          <div className="flex gap-3 mt-4">{form.colors.map((c, i) => <input key={i} type="color" value={c} onChange={(e) => { const n = [...form.colors]; n[i] = e.target.value; setForm({ ...form, colors: n }); }} className="w-12 h-12 rounded-lg border-none" />)}</div>
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={save}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}

/* ============================== Notificaciones ============================== */
function Notifications({ T, guests, todos, suppliers, events, reminders, setReminders, dismissed, setDismissed, notifPrefs }) {
  const [text, setText] = useState("");
  const [date, setDate] = useState("");
  const alerts = [];
  if (notifPrefs.rsvp) guests.filter((g) => g.rsvp === "pending").forEach((g) => alerts.push({ id: `rsvp-${g.id}`, text: `${g.name} no ha respondido el RSVP`, kind: "RSVP" }));
  if (notifPrefs.tasks) todos.filter((t) => !t.done && t.due && new Date(t.due) < new Date()).forEach((t) => alerts.push({ id: `task-${t.id}`, text: `Tarea atrasada: ${t.title}`, kind: "Tarea" }));
  if (notifPrefs.suppliers) suppliers.filter((s) => s.paymentStatus !== "Pagado").forEach((s) => alerts.push({ id: `sup-${s.id}`, text: `${s.company} tiene pago pendiente`, kind: "Proveedor" }));
  const visible = alerts.filter((a) => !dismissed.includes(a.id));

  function addReminder() { if (!text.trim()) return; setReminders((r) => [...r, { id: uid(), text, date }]); setText(""); setDate(""); }

  return (
    <div>
      <SectionTitle T={T}>Notificaciones</SectionTitle>
      <div className="flex flex-col gap-2 mb-6">
        {visible.length === 0 && <EmptyState T={T} text="Sin alertas pendientes — todo al día." />}
        {visible.map((a) => (
          <Card T={T} key={a.id} style={{ padding: "10px 16px" }}>
            <div className="flex items-center justify-between text-sm">
              <span><span className="px-2 py-0.5 rounded-full text-xs mr-2" style={{ background: T.cardAlt, color: T.alert }}>{a.kind}</span>{a.text}</span>
              <button onClick={() => setDismissed((d) => [...d, a.id])}><X size={14} color={T.muted} /></button>
            </div>
          </Card>
        ))}
      </div>

      <div className="text-sm font-medium mb-3">Recordatorios propios</div>
      <Card T={T} style={{ marginBottom: 16 }}>
        <div className="flex gap-3 items-end flex-wrap">
          <Field T={T} label="Recordatorio"><input style={{ ...inputStyle(T), width: 260 }} value={text} onChange={(e) => setText(e.target.value)} /></Field>
          <Field T={T} label="Fecha"><input type="date" style={inputStyle(T)} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
          <Button T={T} small onClick={addReminder}><Plus size={13} /> Agregar</Button>
        </div>
      </Card>
      <div className="flex flex-col gap-2">
        {reminders.map((r) => (
          <Card T={T} key={r.id} style={{ padding: "10px 16px" }}>
            <div className="flex items-center justify-between text-sm"><span>{r.text}</span><div className="flex items-center gap-3 text-xs" style={{ color: T.muted }}>{r.date}<button onClick={() => setReminders((rs) => rs.filter((x) => x.id !== r.id))}><Trash2 size={14} /></button></div></div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* ============================== Configuración ============================== */
function Settings({ T, mode, setMode, themeName, engagementDate, setEngagementDate, events, setEvents, users, setUsers, notifPrefs, setNotifPrefs, couple, setCouple, exportData, importData, saveNow, saveStatus, saveError, clearSavedData }) {
  const [form, setForm] = useState(null);
  const importRef = useRef(null);
  function startNewUser() { setForm({ id: uid(), name: "", email: "", role: "" }); }
  function saveUser() { if (!form.name.trim()) return; setUsers((u) => (u.some((x) => x.id === form.id) ? u.map((x) => (x.id === form.id ? form : x)) : [...u, form])); setForm(null); }

  return (
    <div>
      <SectionTitle T={T}>Configuración</SectionTitle>

      <div className="text-sm font-medium mb-3">Los novios</div>
      <Card T={T} style={{ marginBottom: 24 }}>
        <div className="grid grid-cols-2 gap-4">
          <Field T={T} label="Nombre 1"><input style={inputStyle(T)} value={couple.a} onChange={(e) => setCouple({ ...couple, a: e.target.value })} /></Field>
          <Field T={T} label="Nombre 2"><input style={inputStyle(T)} value={couple.b} onChange={(e) => setCouple({ ...couple, b: e.target.value })} /></Field>
          <Field T={T} label="Fecha de compromiso"><input type="date" style={inputStyle(T)} value={engagementDate} onChange={(e) => setEngagementDate(e.target.value)} /></Field>
        </div>
      </Card>

      <div className="text-sm font-medium mb-3">Apariencia</div>
      <Card T={T} style={{ marginBottom: 24 }}>
        <div className="grid grid-cols-2 gap-4">
          <Field T={T} label="Tema"><select style={inputStyle(T)} value={themeName} disabled><option>Clásico</option></select></Field>
          <Field T={T} label="Modo">
            <div className="flex gap-2">
              <button onClick={() => setMode("light")} className="px-3 py-2 rounded-lg text-xs flex items-center gap-1.5" style={mode === "light" ? { background: T.primary, color: "#fff" } : { background: T.cardAlt, color: T.muted, border: `1px solid ${T.line}` }}><Sun size={13} /> Día</button>
              <button onClick={() => setMode("dark")} className="px-3 py-2 rounded-lg text-xs flex items-center gap-1.5" style={mode === "dark" ? { background: T.primary, color: "#fff" } : { background: T.cardAlt, color: T.muted, border: `1px solid ${T.line}` }}><Moon size={13} /> Noche</button>
            </div>
          </Field>
        </div>
      </Card>

      <div className="text-sm font-medium mb-3">Eventos</div>
      <Card T={T} style={{ marginBottom: 24 }}>
        <p className="text-xs mb-3" style={{ color: T.muted }}>Agrega eventos y sus etapas (ceremonia, recepción, after, etc.) desde el módulo Eventos.</p>
        <div className="flex flex-wrap gap-2">{events.map((ev) => <span key={ev.id} className="text-xs px-2.5 py-1.5 rounded-full" style={{ background: T.cardAlt }}>{ev.name}</span>)}</div>
      </Card>

      <div className="text-sm font-medium mb-3">Guardado automático</div>
      <Card T={T} style={{ marginBottom: 24 }}>
        <p className="text-xs mb-3" style={{ color: T.muted }}>
          La app guarda los datos solos, en segundo plano, cada vez que cambias algo — no necesitas hacer nada. Se guardan en tu base de datos de Supabase, compartida entre quien abra este mismo sitio (tú y Roberto), así que trata la URL como algo privado de la pareja.
        </p>
        <div className="flex items-center gap-3 mb-3">
          <span className="text-xs px-2.5 py-1 rounded-full" style={{ background: T.cardAlt, color: saveStatus === "error" ? T.alert : T.secondaryDark }}>
            {saveStatus === "saving" ? "Guardando…" : saveStatus === "error" ? "No se pudo guardar" : "Guardado ✓"}
          </span>
          <Button T={T} small variant="outline" onClick={saveNow}>Guardar ahora</Button>
        </div>
        {saveStatus === "error" && <p className="text-xs mb-3" style={{ color: T.alert }}>Detalle del error: {saveError || "desconocido"}</p>}
        <div className="flex gap-2 flex-wrap">
          <Button T={T} small onClick={exportData}>Descargar respaldo (JSON)</Button>
          <Button T={T} small variant="outline" onClick={() => importRef.current?.click()}>Restaurar desde archivo</Button>
          <input ref={importRef} type="file" accept="application/json" className="hidden" onChange={(e) => e.target.files?.[0] && importData(e.target.files[0])} />
        </div>
        <p className="text-xs mt-3" style={{ color: T.muted }}>El respaldo en JSON sigue sirviendo para llevarte una copia offline o mover los datos a otra sesión.</p>
        <div style={{ borderTop: `1px solid ${T.line}`, marginTop: 14, paddingTop: 14 }}>
          <button onClick={clearSavedData} className="text-xs" style={{ color: T.alert }}>Borrar todos los datos guardados</button>
        </div>
      </Card>

      <div className="text-sm font-medium mb-3">Notificaciones</div>
      <Card T={T} style={{ marginBottom: 24 }}>
        <div className="flex flex-col gap-3">
          <div className="flex items-center justify-between"><span className="text-sm">RSVP pendientes</span><Toggle T={T} checked={notifPrefs.rsvp} onChange={(v) => setNotifPrefs({ ...notifPrefs, rsvp: v })} /></div>
          <div className="flex items-center justify-between"><span className="text-sm">Tareas atrasadas</span><Toggle T={T} checked={notifPrefs.tasks} onChange={(v) => setNotifPrefs({ ...notifPrefs, tasks: v })} /></div>
          <div className="flex items-center justify-between"><span className="text-sm">Proveedores sin pagar</span><Toggle T={T} checked={notifPrefs.suppliers} onChange={(v) => setNotifPrefs({ ...notifPrefs, suppliers: v })} /></div>
        </div>
      </Card>

      <div className="flex items-center justify-between mb-3"><div className="text-sm font-medium">Usuarios</div><Button T={T} small onClick={startNewUser}><Plus size={13} /> Agregar usuario</Button></div>
      <p className="text-xs mb-3" style={{ color: T.muted }}>Directorio de acceso — esta app corre localmente en tu navegador, así que esto es un registro informativo y no un inicio de sesión real.</p>
      <div className="flex flex-col gap-2">
        {users.map((u) => (
          <Card T={T} key={u.id} style={{ padding: "10px 16px" }}>
            <div className="flex items-center justify-between text-sm">
              <div><span className="font-medium">{u.name}</span><span className="ml-2 text-xs" style={{ color: T.muted }}>{u.role}</span></div>
              <div className="flex items-center gap-3 text-xs" style={{ color: T.muted }}><Mail size={12} />{u.email}<button onClick={() => setForm(u)}>Editar</button><button onClick={() => setUsers((us) => us.filter((x) => x.id !== u.id))}><Trash2 size={13} /></button></div>
            </div>
          </Card>
        ))}
      </div>

      {form && (
        <Modal T={T} title="Usuario" onClose={() => setForm(null)}>
          <div className="grid grid-cols-2 gap-4">
            <Field T={T} label="Nombre"><input style={inputStyle(T)} value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })} autoFocus /></Field>
            <Field T={T} label="Rol"><input style={inputStyle(T)} value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })} /></Field>
            <div className="col-span-2"><Field T={T} label="Email"><input style={{ ...inputStyle(T), width: "100%" }} value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })} /></Field></div>
          </div>
          <div className="flex gap-2 mt-6 justify-end">
            <Button T={T} variant="outline" onClick={() => setForm(null)}>Cancelar</Button>
            <Button T={T} onClick={saveUser}><Check size={14} /> Guardar</Button>
          </div>
        </Modal>
      )}
    </div>
  );
}
