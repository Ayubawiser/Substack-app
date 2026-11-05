import { useEffect, useMemo, useRef, useState } from "react";
import { ensureFirebaseUser, loadWorkspace, saveWorkspace } from "./firebase";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Users, FolderKanban, FileCheck2, Bell, Search,
  Plus, Menu, X, ChevronDown, MoreHorizontal, CalendarDays,
  AlertTriangle, CheckCircle2, Clock3, FileText, ArrowUpRight,
  ClipboardList, Settings, Download, Trash2, Pencil, ShieldCheck, Upload, FileSpreadsheet
} from "lucide-react";

type DocStatus = "Submitted" | "Missing" | "Expiring soon" | "Expired";
type DocumentRecord = { id: string; name: string; status: DocStatus; expiry: string };
type Subcontractor = {
  id: string; company: string; contact: string; email: string; phone: string;
  documents: DocumentRecord[]; lastContact: string; projects: string[];
};
type Project = { id: string; name: string; location: string; status: "Active" | "Planning" | "Completed" };

const today = new Date();
const datePlus = (days: number) => {
  const d = new Date(today); d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
};
const makeDocs = (insurance: DocStatus, licence: DocStatus, safety: DocStatus, expiryDays = 100): DocumentRecord[] => [
  { id: crypto.randomUUID(), name: "Insurance certificate", status: insurance, expiry: datePlus(expiryDays) },
  { id: crypto.randomUUID(), name: "Business licence", status: licence, expiry: datePlus(180) },
  { id: crypto.randomUUID(), name: "Safety document", status: safety, expiry: datePlus(240) }
];

const initialProjects: Project[] = [
  { id: "p1", name: "Riverside Office Renovation", location: "Manchester, UK", status: "Active" },
  { id: "p2", name: "West End Retail Fit-out", location: "London, UK", status: "Active" },
  { id: "p3", name: "Northgate Residential", location: "Leeds, UK", status: "Planning" }
];
const initialSubs: Subcontractor[] = [
  { id: "s1", company: "Apex Electrical Ltd", contact: "Daniel Brooks", email: "daniel@apex.example", phone: "+44 7700 900101", documents: makeDocs("Submitted", "Submitted", "Submitted", 120), lastContact: datePlus(-3), projects: ["p1", "p2"] },
  { id: "s2", company: "Northstar Plumbing", contact: "Olivia Reed", email: "olivia@northstar.example", phone: "+44 7700 900102", documents: makeDocs("Expiring soon", "Submitted", "Submitted", 12), lastContact: datePlus(-1), projects: ["p1"] },
  { id: "s3", company: "BuildRight Interiors", contact: "James Carter", email: "james@buildright.example", phone: "+44 7700 900103", documents: makeDocs("Missing", "Submitted", "Missing"), lastContact: datePlus(-8), projects: ["p2", "p3"] },
  { id: "s4", company: "Summit Roofing Co.", contact: "Amelia Hayes", email: "amelia@summit.example", phone: "+44 7700 900104", documents: makeDocs("Submitted", "Expired", "Submitted", -4), lastContact: datePlus(-5), projects: ["p3"] },
  { id: "s5", company: "Clearline Flooring", contact: "Noah Wilson", email: "noah@clearline.example", phone: "+44 7700 900105", documents: makeDocs("Submitted", "Submitted", "Submitted"), lastContact: datePlus(-2), projects: ["p1"] }
];

const navItems = [
  { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { id: "subcontractors", label: "Subcontractors", icon: Users },
  { id: "projects", label: "Projects", icon: FolderKanban },
  { id: "documents", label: "Documents", icon: FileCheck2 },
  { id: "followups", label: "Follow-ups", icon: Bell }
];

function docStatus(doc: DocumentRecord): DocStatus {
  if (doc.status === "Missing") return "Missing";
  if (doc.expiry) {
    const days = Math.ceil((new Date(doc.expiry + "T23:59:59").getTime() - new Date().getTime()) / 86400000);
    if (days < 0) return "Expired";
    if (days <= 30) return "Expiring soon";
  }
  return doc.status === "Expired" ? "Expired" : "Submitted";
}
function subStatus(sub: Subcontractor) {
  const statuses = sub.documents.map(docStatus);
  if (statuses.some(s => s === "Missing" || s === "Expired")) return "Action required";
  if (statuses.some(s => s === "Expiring soon")) return "Review needed";
  return "Cleared";
}
function Badge({ children }: { children: string }) {
  const cls = children === "Cleared" || children === "Submitted" || children === "Active" ? "green"
    : children === "Expiring soon" || children === "Review needed" || children === "Planning" ? "amber"
    : children === "Expired" || children === "Missing" || children === "Action required" ? "red" : "gray";
  return <span className={`badge ${cls}`}>{children}</span>;
}

type ImportRow = Record<string, unknown>;
const clean = (value: unknown) => String(value ?? "").trim();
const keyOf = (value: unknown) => clean(value).toLowerCase();
const validStatus = (value: unknown): DocStatus => {
  const s = clean(value).toLowerCase();
  if (s === "missing") return "Missing";
  if (s === "expired") return "Expired";
  if (s === "expiring soon") return "Expiring soon";
  return "Submitted";
};
const downloadBlob = (blob: Blob, filename: string) => {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
};
const rowsToCsv = (rows: Record<string, unknown>[]) => {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  const quote = (v: unknown) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  return [headers.map(quote).join(","), ...rows.map(row => headers.map(h => quote(row[h])).join(","))].join("\r\n");
};

function App() {
  const [page, setPage] = useState("dashboard");
  const [mobileOpen, setMobileOpen] = useState(false);
  const [subs, setSubs] = useState<Subcontractor[]>(() => {
    try { const saved = localStorage.getItem("subtrack-subs"); return saved ? JSON.parse(saved) : initialSubs; } catch { return initialSubs; }
  });
  const [projects, setProjects] = useState<Project[]>(() => {
    try { const saved = localStorage.getItem("subtrack-projects"); return saved ? JSON.parse(saved) : initialProjects; } catch { return initialProjects; }
  });
  const [search, setSearch] = useState("");
  const [modal, setModal] = useState<"sub" | "project" | "edit-sub" | null>(null);
  const [editing, setEditing] = useState<Subcontractor | null>(null);
  const [selectedSub, setSelectedSub] = useState<string | null>(null);
  const [subForm, setSubForm] = useState({ company: "", contact: "", email: "", phone: "" });
  const [projectForm, setProjectForm] = useState({ name: "", location: "" });
  const [toast, setToast] = useState("");
  const [filter, setFilter] = useState("All");
  const [dataModal, setDataModal] = useState<"import" | "export" | null>(null);
  const [importMode, setImportMode] = useState<"merge" | "replace">("merge");
  const [importBusy, setImportBusy] = useState(false);
  const [cloudUserId, setCloudUserId] = useState<string | null>(null);
  const [cloudReady, setCloudReady] = useState(false);
  const [cloudMessage, setCloudMessage] = useState("Connecting to cloud…");
  const importInput = useRef<HTMLInputElement>(null);

  const persistSubs = (next: Subcontractor[]) => { setSubs(next); localStorage.setItem("subtrack-subs", JSON.stringify(next)); };
  const persistProjects = (next: Project[]) => { setProjects(next); localStorage.setItem("subtrack-projects", JSON.stringify(next)); };

  useEffect(() => {
    let active = true;
    (async () => {
      try {
        const user = await ensureFirebaseUser();
        if (!active) return;
        setCloudUserId(user.uid);
        const remote = await loadWorkspace(user.uid);
        if (!active) return;
        if (remote) {
          if (Array.isArray(remote.subs)) {
            setSubs(remote.subs as Subcontractor[]);
            localStorage.setItem("subtrack-subs", JSON.stringify(remote.subs));
          }
          if (Array.isArray(remote.projects)) {
            setProjects(remote.projects as Project[]);
            localStorage.setItem("subtrack-projects", JSON.stringify(remote.projects));
          }
          setCloudMessage("Cloud connected · data loaded");
        } else {
          const localSubs = (() => { try { const v = localStorage.getItem("subtrack-subs"); return v ? JSON.parse(v) as Subcontractor[] : initialSubs; } catch { return initialSubs; } })();
          const localProjects = (() => { try { const v = localStorage.getItem("subtrack-projects"); return v ? JSON.parse(v) as Project[] : initialProjects; } catch { return initialProjects; } })();
          setSubs(localSubs); setProjects(localProjects);
          await saveWorkspace(user.uid, localSubs, localProjects);
          setCloudMessage("Cloud connected · workspace created");
        }
        setCloudReady(true);
      } catch (error) {
        console.error("Firebase connection failed", error);
        if (active) setCloudMessage("Cloud connection failed · check Firebase setup");
      }
    })();
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!cloudReady || !cloudUserId) return;
    const timer = window.setTimeout(() => {
      saveWorkspace(cloudUserId, subs, projects)
        .then(() => setCloudMessage("All changes saved to cloud"))
        .catch(error => { console.error("Cloud save failed", error); setCloudMessage("Cloud save failed · retry after checking connection"); });
    }, 500);
    return () => window.clearTimeout(timer);
  }, [subs, projects, cloudReady, cloudUserId]);
  const cleared = subs.filter(s => subStatus(s) === "Cleared").length;
  const missing = subs.reduce((n, s) => n + s.documents.filter(d => docStatus(d) === "Missing" || docStatus(d) === "Expired").length, 0);
  const expiring = subs.reduce((n, s) => n + s.documents.filter(d => docStatus(d) === "Expiring soon").length, 0);
  const allDocs = subs.flatMap(s => s.documents.map(d => ({ ...d, company: s.company, subId: s.id })));
  const filteredSubs = useMemo(() => subs.filter(s => {
    const match = `${s.company} ${s.contact} ${s.email}`.toLowerCase().includes(search.toLowerCase());
    return match && (filter === "All" || subStatus(s) === filter);
  }), [subs, search, filter]);

  const notify = (message: string) => { setToast(message); window.setTimeout(() => setToast(""), 2600); };
  const openAddSub = () => { setEditing(null); setSubForm({ company: "", contact: "", email: "", phone: "" }); setModal("sub"); };
  const openEditSub = (s: Subcontractor) => { setEditing(s); setSubForm({ company: s.company, contact: s.contact, email: s.email, phone: s.phone }); setModal("edit-sub"); };
  const saveSub = (e: React.FormEvent) => {
    e.preventDefault();
    if (!subForm.company.trim()) return;
    if (editing) {
      persistSubs(subs.map(s => s.id === editing.id ? { ...s, ...subForm } : s));
      notify("Subcontractor updated");
    } else {
      persistSubs([{ id: crypto.randomUUID(), ...subForm, documents: makeDocs("Missing", "Missing", "Missing"), lastContact: "Never", projects: [] }, ...subs]);
      notify("Subcontractor added");
    }
    setModal(null);
  };
  const saveProject = (e: React.FormEvent) => {
    e.preventDefault();
    if (!projectForm.name.trim()) return;
    persistProjects([{ id: crypto.randomUUID(), ...projectForm, status: "Planning" }, ...projects]);
    setProjectForm({ name: "", location: "" }); setModal(null); notify("Project created");
  };
  const updateDoc = (subId: string, docId: string, patch: Partial<DocumentRecord>) => {
    persistSubs(subs.map(s => s.id === subId ? { ...s, documents: s.documents.map(d => d.id === docId ? { ...d, ...patch } : d) } : s));
    notify("Document updated");
  };
  const deleteSub = (id: string) => {
    if (window.confirm("Delete this subcontractor and its document records?")) {
      persistSubs(subs.filter(s => s.id !== id)); notify("Subcontractor deleted");
    }
  };

  const exportRows = () => {
    const subRows = subs.map(s => ({
      Company: s.company, Contact: s.contact, Email: s.email, Phone: s.phone,
      "Last Contact": s.lastContact,
      Projects: s.projects.map(id => projects.find(p => p.id === id)?.name ?? "").filter(Boolean).join("; ")
    }));
    const projectRows = projects.map(p => ({ Name: p.name, Location: p.location, Status: p.status }));
    const documentRows = subs.flatMap(s => s.documents.map(d => ({
      Company: s.company, Document: d.name, Status: docStatus(d), Expiry: d.expiry
    })));
    const followupRows = subs.map(s => ({ Company: s.company, "Last Contact": s.lastContact, "Open Issues": s.documents.filter(d => ["Missing", "Expired", "Expiring soon"].includes(docStatus(d))).map(d => d.name).join("; ") }));
    return { subRows, projectRows, documentRows, followupRows };
  };
  const exportExcel = () => {
    const rows = exportRows();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.subRows), "Subcontractors");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.projectRows), "Projects");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.documentRows), "Documents");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.followupRows), "Follow-ups");
    XLSX.writeFile(wb, "subtrack-report.xlsx");
    setDataModal(null); notify("Excel workbook exported");
  };
  const exportCsvFiles = () => {
    const rows = exportRows();
    [
      ["subtrack-subcontractors.csv", rows.subRows],
      ["subtrack-projects.csv", rows.projectRows],
      ["subtrack-documents.csv", rows.documentRows],
      ["subtrack-follow-ups.csv", rows.followupRows]
    ].forEach(([name, data]) => downloadBlob(new Blob([rowsToCsv(data as Record<string, unknown>[])], { type: "text/csv;charset=utf-8" }), name as string));
    setDataModal(null); notify("CSV reports downloaded");
  };
  const importFile = async (file?: File) => {
    if (!file) return;
    setImportBusy(true);
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
      const sheetRows = (name: string) => {
        const sheet = workbook.Sheets[name];
        return sheet ? XLSX.utils.sheet_to_json<ImportRow>(sheet, { defval: "" }) : [];
      };
      const normalize = (rows: ImportRow[]) => rows.map(row => {
        const out: Record<string, unknown> = {};
        Object.entries(row).forEach(([k, v]) => out[k.toLowerCase().replace(/[^a-z0-9]/g, "")] = v);
        return out;
      });
      let subRows = normalize(sheetRows("Subcontractors"));
      let projectRows = normalize(sheetRows("Projects"));
      let docRows = normalize(sheetRows("Documents"));
      if (!subRows.length && !projectRows.length && !docRows.length) {
        const first = workbook.Sheets[workbook.SheetNames[0]];
        const generic = normalize(first ? XLSX.utils.sheet_to_json<ImportRow>(first, { defval: "" }) : []);
        if (generic.some(r => r.company && (r.document || r.expiry || r.status))) docRows = generic;
        else if (generic.some(r => r.company && (r.email || r.contact || r.phone))) subRows = generic;
        else if (generic.some(r => r.name && (r.location || r.status))) projectRows = generic;
        else throw new Error("No recognised columns. Use the Subcontractors, Projects or Documents template.");
      }

      let nextProjects: Project[] = importMode === "replace" ? [] : [...projects];
      const projectIdByName = new Map<string, string>();
      nextProjects.forEach(p => projectIdByName.set(keyOf(p.name), p.id));
      projectRows.forEach(r => {
        const name = clean(r.name || r.project);
        if (!name) return;
        const existingId = projectIdByName.get(keyOf(name));
        if (existingId) {
          nextProjects = nextProjects.map(p => p.id === existingId ? { ...p, location: clean(r.location) || p.location, status: (["Active","Planning","Completed"].includes(clean(r.status)) ? clean(r.status) : p.status) as Project["status"] } : p);
        } else {
          const id = crypto.randomUUID();
          nextProjects.push({ id, name, location: clean(r.location), status: (["Active","Planning","Completed"].includes(clean(r.status)) ? clean(r.status) : "Planning") as Project["status"] });
          projectIdByName.set(keyOf(name), id);
        }
      });

      let nextSubs: Subcontractor[] = importMode === "replace" ? [] : [...subs];
      const subIdByCompany = new Map<string, string>();
      nextSubs.forEach(s => subIdByCompany.set(keyOf(s.company), s.id));
      subRows.forEach(r => {
        const company = clean(r.company || r.companyname);
        if (!company) return;
        const nameList = clean(r.projects).split(/[;,]/).map(x => x.trim()).filter(Boolean);
        const linkedProjects = nameList.map(n => projectIdByName.get(keyOf(n))).filter((x): x is string => Boolean(x));
        const existingId = subIdByCompany.get(keyOf(company));
        if (existingId) {
          nextSubs = nextSubs.map(s => s.id === existingId ? {
            ...s, contact: clean(r.contact) || s.contact, email: clean(r.email) || s.email,
            phone: clean(r.phone) || s.phone, lastContact: clean(r.lastcontact) || s.lastContact,
            projects: linkedProjects.length ? linkedProjects : s.projects
          } : s);
        } else {
          const id = crypto.randomUUID();
          nextSubs.push({ id, company, contact: clean(r.contact), email: clean(r.email), phone: clean(r.phone),
            lastContact: clean(r.lastcontact) || "Never", projects: linkedProjects, documents: [] });
          subIdByCompany.set(keyOf(company), id);
        }
      });

      docRows.forEach(r => {
        const company = clean(r.company || r.companyname);
        const docName = clean(r.document || r.documentname);
        if (!company || !docName) return;
        let subId = subIdByCompany.get(keyOf(company));
        if (!subId) {
          subId = crypto.randomUUID();
          nextSubs.push({ id: subId, company, contact: "", email: "", phone: "", lastContact: "Never", projects: [], documents: [] });
          subIdByCompany.set(keyOf(company), subId);
        }
        const doc: DocumentRecord = { id: crypto.randomUUID(), name: docName, status: validStatus(r.status), expiry: clean(r.expiry) };
        nextSubs = nextSubs.map(s => s.id === subId ? { ...s, documents: [...s.documents.filter(d => keyOf(d.name) !== keyOf(docName)), doc] } : s);
      });
      if (!subRows.length && !docRows.length && !projectRows.length) throw new Error("The selected file contains no importable rows.");
      persistProjects(nextProjects);
      persistSubs(nextSubs);
      setDataModal(null);
      notify(`Import complete: ${nextSubs.length} subcontractors, ${nextProjects.length} projects`);
    } catch (error) {
      window.alert(error instanceof Error ? error.message : "Could not read this file. Please check its format.");
    } finally {
      setImportBusy(false);
      if (importInput.current) importInput.current.value = "";
    }
  };

  const title = navItems.find(n => n.id === page)?.label ?? "SubTrack";

  const Sidebar = () => <aside className={`sidebar ${mobileOpen ? "open" : ""}`}>
    <div className="brand"><div className="brand-mark"><ShieldCheck size={21}/></div><span>subtrack</span><button className="mobile-close" onClick={() => setMobileOpen(false)}><X size={19}/></button></div>
    <div className="workspace"><div className="workspace-avatar">AC</div><div><strong>Acme Contractors</strong><small>Workspace</small></div><ChevronDown size={15}/></div>
    <div className="nav-label">WORKSPACE</div>
    <nav>{navItems.map(item => <button key={item.id} className={`nav-item ${page === item.id ? "active" : ""}`} onClick={() => { setPage(item.id); setMobileOpen(false); }}><item.icon size={18}/><span>{item.label}</span>{item.id === "followups" && (missing + expiring > 0) && <i>{missing + expiring}</i>}</button>)}</nav>
    <div className="sidebar-bottom"><button className="nav-item"><Settings size={18}/><span>Settings</span></button><div className="profile"><div className="profile-avatar">JD</div><div><strong>Jordan Davis</strong><small>Project administrator</small></div><MoreHorizontal size={18}/></div></div>
  </aside>;

  return <div className="app-shell">
    {mobileOpen && <div className="scrim" onClick={() => setMobileOpen(false)} />}
    <Sidebar/>
    <main className="main">
      <header className="topbar"><button className="icon-button mobile-menu" onClick={() => setMobileOpen(true)}><Menu size={21}/></button><div className="breadcrumbs">Workspace <span>/</span> <strong>{title}</strong></div><div className="top-actions"><div className="searchbox"><Search size={16}/><input placeholder="Search anything..." value={search} onChange={e => setSearch(e.target.value)}/><kbd>⌘ K</kbd></div><button className="icon-button notification" onClick={() => setPage("followups")}><Bell size={18}/><i/></button><div className="top-avatar">JD</div></div></header>
      <div className="content">
        {page === "dashboard" && <>
          <div className="page-heading"><div><div className="eyebrow">OVERVIEW</div><h1>Good morning, Jordan <span>✦</span></h1><p>Here's what's happening across your subcontractors today.</p></div><button className="primary" onClick={openAddSub}><Plus size={17}/> Add subcontractor</button></div>
          <div className="stats-grid">
            <div className="stat-card"><div className="stat-top"><span>Total subcontractors</span><span className="stat-icon blue"><Users size={18}/></span></div><strong>{subs.length}</strong><small>Across all projects</small></div>
            <div className="stat-card"><div className="stat-top"><span>Fully cleared</span><span className="stat-icon green"><CheckCircle2 size={18}/></span></div><strong>{cleared}</strong><small>{subs.length ? Math.round(cleared / subs.length * 100) : 0}% of subcontractors</small></div>
            <div className="stat-card"><div className="stat-top"><span>Missing / expired</span><span className="stat-icon red"><AlertTriangle size={18}/></span></div><strong>{missing}</strong><small>Documents need attention</small></div>
            <div className="stat-card"><div className="stat-top"><span>Expiring soon</span><span className="stat-icon amber"><Clock3 size={18}/></span></div><strong>{expiring}</strong><small>Within the next 30 days</small></div>
          </div>
          <div className="dashboard-grid">
            <section className="panel wide"><div className="panel-heading"><div><h2>Subcontractor status</h2><p>Clearance overview across your workspace</p></div><button className="text-button" onClick={() => setPage("subcontractors")}>View all <ArrowUpRight size={15}/></button></div>
              <div className="status-bar">{["Cleared","Review needed","Action required"].map((label, i) => <div key={label} style={{ width: `${subs.length ? (subs.filter(s => subStatus(s) === label).length / subs.length * 100) : 0}%` }} className={`bar-segment seg-${i}`}/>)}</div>
              <div className="legend">{[["Cleared",cleared,"green-dot"],["Review needed",subs.filter(s => subStatus(s) === "Review needed").length,"amber-dot"],["Action required",subs.filter(s => subStatus(s) === "Action required").length,"red-dot"]].map(([label,count,dot]) => <div key={String(label)}><i className={String(dot)}/><span>{label}</span><strong>{count}</strong></div>)}</div>
            </section>
            <section className="panel"><div className="panel-heading"><div><h2>Needs attention</h2><p>Items to follow up</p></div><span className="attention-count">{missing + expiring}</span></div>
              <div className="attention-list">{subs.filter(s => subStatus(s) !== "Cleared").slice(0,4).map(s => <button className="attention-row" key={s.id} onClick={() => { setSelectedSub(s.id); setPage("documents"); }}><div className="attention-icon"><FileText size={17}/></div><div><strong>{s.company}</strong><small>{s.documents.filter(d => ["Missing","Expired","Expiring soon"].includes(docStatus(d))).length} document issue(s)</small></div><ArrowUpRight size={15}/></button>)}{subs.every(s => subStatus(s) === "Cleared") && <div className="empty">All subcontractors are cleared.</div>}</div>
            </section>
          </div>
          <section className="panel table-panel"><div className="panel-heading"><div><h2>Recent subcontractors</h2><p>Latest updates from your directory</p></div><button className="text-button" onClick={() => setPage("subcontractors")}>View directory <ArrowUpRight size={15}/></button></div><SubTable items={subs.slice(0,5)} onEdit={openEditSub} onDelete={deleteSub} onOpen={id => { setSelectedSub(id); setPage("documents"); }}/></section>
        </>}
        {page === "subcontractors" && <><div className="page-heading"><div><div className="eyebrow">DIRECTORY</div><h1>Subcontractors</h1><p>Manage your subcontractor relationships and clearance status.</p></div><button className="primary" onClick={openAddSub}><Plus size={17}/> Add subcontractor</button></div><div className="panel table-panel"><div className="toolbar"><div className="filter-tabs">{["All","Cleared","Review needed","Action required"].map(f => <button className={filter === f ? "selected" : ""} onClick={() => setFilter(f)} key={f}>{f}</button>)}</div><span className="muted">{filteredSubs.length} results</span></div><SubTable items={filteredSubs} onEdit={openEditSub} onDelete={deleteSub} onOpen={id => { setSelectedSub(id); setPage("documents"); }}/></div></>}
        {page === "projects" && <><div className="page-heading"><div><div className="eyebrow">WORKSPACE</div><h1>Projects</h1><p>Organise subcontractors by the jobs they're working on.</p></div><button className="primary" onClick={() => setModal("project")}><Plus size={17}/> New project</button></div><div className="project-grid">{projects.map(p => <div className="project-card" key={p.id}><div className="project-card-top"><div className="project-icon"><FolderKanban size={20}/></div><Badge>{p.status}</Badge></div><h3>{p.name}</h3><p><span><CalendarDays size={14}/>{p.location || "Location not set"}</span></p><div className="project-footer"><span>{subs.filter(s => s.projects.includes(p.id)).length} subcontractors</span><button className="text-button" onClick={() => { setSearch(""); setPage("subcontractors"); }}>View directory <ArrowUpRight size={14}/></button></div></div>)}{projects.length === 0 && <div className="empty">No projects yet. Create your first project.</div>}</div></>}
        {page === "documents" && <><div className="page-heading"><div><div className="eyebrow">COMPLIANCE</div><h1>Document tracking</h1><p>See which documents are submitted, missing or approaching expiry.</p></div></div><div className="panel table-panel"><div className="panel-heading"><div><h2>All documents</h2><p>{allDocs.length} records across {subs.length} subcontractors</p></div><div className="data-actions"><button className="secondary" onClick={() => setDataModal("import")}><Upload size={15}/> Import data</button><button className="primary" onClick={() => setDataModal("export")}><Download size={15}/> Export reports</button></div></div><div className="document-list">{subs.flatMap(s => s.documents.map(d => ({ ...d, company: s.company, subId: s.id }))).map(d => <div className="document-row" key={d.id}><div className="doc-file"><div className="file-icon"><FileText size={18}/></div><div><strong>{d.name}</strong><small>{d.company}</small></div></div><Badge>{docStatus(d)}</Badge><div className="doc-expiry"><small>Expiry date</small><strong>{d.expiry ? new Date(d.expiry + "T12:00:00").toLocaleDateString() : "Not set"}</strong></div><select aria-label="Update document status" value={d.status} onChange={e => updateDoc(d.subId, d.id, { status: e.target.value as DocStatus })}><option>Submitted</option><option>Missing</option><option>Expiring soon</option><option>Expired</option></select><input aria-label="Expiry date" type="date" value={d.expiry} onChange={e => updateDoc(d.subId, d.id, { expiry: e.target.value })}/></div>)}</div></div></>}
        {page === "followups" && <><div className="page-heading"><div><div className="eyebrow">ACTION CENTRE</div><h1>Follow-ups</h1><p>Keep document requests moving and avoid last-minute surprises.</p></div></div><div className="panel table-panel"><div className="panel-heading"><div><h2>Follow-up queue</h2><p>Subcontractors with missing, expired or expiring documents</p></div><span className="attention-count">{subs.filter(s => subStatus(s) !== "Cleared").length}</span></div>{subs.filter(s => subStatus(s) !== "Cleared").map(s => <div className="follow-row" key={s.id}><div className="company-avatar">{s.company.slice(0,1)}</div><div className="follow-main"><strong>{s.company}</strong><small>{s.documents.filter(d => ["Missing","Expired","Expiring soon"].includes(docStatus(d))).map(d => d.name).join(", ")}</small></div><div className="follow-last"><small>Last contacted</small><strong>{s.lastContact === "Never" ? "Never" : new Date(s.lastContact + "T12:00:00").toLocaleDateString()}</strong></div><Badge>{subStatus(s)}</Badge><button className="secondary small" onClick={() => { persistSubs(subs.map(x => x.id === s.id ? { ...x, lastContact: datePlus(0) } : x)); notify("Follow-up date recorded"); }}>Log contact</button></div>)}{subs.every(s => subStatus(s) === "Cleared") && <div className="empty">Nothing to follow up. Great work!</div>}</div></>}
      </div>
    </main>
    {dataModal && <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget && !importBusy) setDataModal(null); }}>
      <div className="modal data-modal">
        <div className="modal-head"><div><h2>{dataModal === "import" ? "Import existing data" : "Export reports"}</h2><p>{dataModal === "import" ? "Bring your spreadsheet records into SubTrack." : "Download your workspace data in a spreadsheet format."}</p></div><button type="button" className="icon-button" onClick={() => setDataModal(null)}><X size={19}/></button></div>
        {dataModal === "export" ? <>
          <button className="data-option" onClick={exportExcel}><FileSpreadsheet size={21}/><span><strong>Excel workbook (.xlsx)</strong><small>One workbook with Subcontractors, Projects, Documents and Follow-ups tabs. You can also open it in Google Sheets.</small></span><ArrowUpRight size={16}/></button>
          <button className="data-option" onClick={exportCsvFiles}><FileText size={21}/><span><strong>CSV reports</strong><small>Downloads separate CSV files for each data section. CSV files can be imported into Google Sheets.</small></span><ArrowUpRight size={16}/></button>
        </> : <>
          <div className="import-mode"><strong>How should we handle matching records?</strong><label><input type="radio" checked={importMode === "merge"} onChange={() => setImportMode("merge")}/> Add new records and update matching records</label><label><input type="radio" checked={importMode === "replace"} onChange={() => setImportMode("replace")}/> Replace current records with the spreadsheet data</label></div>
          <input ref={importInput} type="file" accept=".xlsx,.xls,.csv" className="file-input" onChange={e => importFile(e.target.files?.[0])}/>
          <button className="primary import-pick" disabled={importBusy} onClick={() => importInput.current?.click()}><Upload size={16}/>{importBusy ? "Importing..." : "Choose Excel or CSV file"}</button>
          <p className="import-help">For Google Sheets: open your sheet, choose File → Download → Microsoft Excel (.xlsx) or Comma-separated values (.csv), then select that file here. For best results, use the four tabs from an exported SubTrack workbook.</p>
          <p className="import-help">Import supports company, contact, email, phone, project, document, status and expiry columns. Keep a backup before choosing Replace.</p>
        </>}
        <div className="modal-actions"><button type="button" className="secondary" onClick={() => setDataModal(null)}>Close</button></div>
      </div>
    </div>}
    {modal && <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) setModal(null); }}><form className="modal" onSubmit={modal === "project" ? saveProject : saveSub}><div className="modal-head"><div><h2>{modal === "project" ? "Create project" : editing ? "Edit subcontractor" : "Add subcontractor"}</h2><p>{modal === "project" ? "Add a project to your workspace." : "Enter the company's contact details."}</p></div><button type="button" className="icon-button" onClick={() => setModal(null)}><X size={19}/></button></div>{modal === "project" ? <><label>Project name<input required value={projectForm.name} onChange={e => setProjectForm({ ...projectForm, name: e.target.value })} placeholder="e.g. Central Office Refurbishment"/></label><label>Location<input value={projectForm.location} onChange={e => setProjectForm({ ...projectForm, location: e.target.value })} placeholder="City, country"/></label></> : <><label>Company name<input required value={subForm.company} onChange={e => setSubForm({ ...subForm, company: e.target.value })} placeholder="Company Ltd"/></label><label>Contact person<input value={subForm.contact} onChange={e => setSubForm({ ...subForm, contact: e.target.value })} placeholder="Full name"/></label><label>Email address<input type="email" value={subForm.email} onChange={e => setSubForm({ ...subForm, email: e.target.value })} placeholder="name@company.com"/></label><label>Phone number<input value={subForm.phone} onChange={e => setSubForm({ ...subForm, phone: e.target.value })} placeholder="+1 555 000 0000"/></label></>}<div className="modal-actions"><button type="button" className="secondary" onClick={() => setModal(null)}>Cancel</button><button className="primary" type="submit">{modal === "project" ? "Create project" : editing ? "Save changes" : "Add subcontractor"}</button></div></form></div>}
    {toast && <div className="toast"><CheckCircle2 size={17}/>{toast}</div>}
  </div>;
}

function SubTable({ items, onEdit, onDelete, onOpen }: { items: Subcontractor[]; onEdit: (s: Subcontractor) => void; onDelete: (id: string) => void; onOpen: (id: string) => void }) {
  if (!items.length) return <div className="empty">No subcontractors found. Try changing your search or add a subcontractor.</div>;
  return <div className="table-scroll"><table><thead><tr><th>COMPANY</th><th>CONTACT</th><th>DOCUMENTS</th><th>STATUS</th><th>LAST CONTACT</th><th></th></tr></thead><tbody>{items.map(s => <tr key={s.id}><td><button className="company-link" onClick={() => onOpen(s.id)}><span className="company-avatar">{s.company.slice(0,1)}</span><strong>{s.company}</strong></button></td><td><strong className="contact-name">{s.contact || "—"}</strong><small className="cell-sub">{s.email || "No email"}</small></td><td><span className="doc-count">{s.documents.filter(d => docStatus(d) === "Submitted").length} / {s.documents.length} submitted</span></td><td><Badge>{subStatus(s)}</Badge></td><td className="last-contact">{s.lastContact === "Never" ? "Never" : new Date(s.lastContact + "T12:00:00").toLocaleDateString()}</td><td><div className="row-actions"><button title="Edit" onClick={() => onEdit(s)}><Pencil size={15}/></button><button title="Delete" onClick={() => onDelete(s.id)}><Trash2 size={15}/></button><button title="View documents" onClick={() => onOpen(s.id)}><ArrowUpRight size={15}/></button></div></td></tr>)}</tbody></table></div>;
}

export default App;