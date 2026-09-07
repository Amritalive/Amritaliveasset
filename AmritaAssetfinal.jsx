import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Search, Plus, Upload, Download, MoreVertical, Pencil, Trash2, Eye, X,
  ChevronUp, ChevronDown, ChevronLeft, ChevronRight, Server, UserCheck,
  Wrench, Building2, AlertTriangle, Cast, Filter, Inbox,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/* Design tokens                                                       */
/* ------------------------------------------------------------------ */
const C = {
  bg: "#0A0F1C",
  panel: "#0F1729",
  panelAlt: "#141F35",
  header: "#080D18",
  border: "#1D2A42",
  borderLight: "#28374F",
  text: "#E6EBF5",
  muted: "#8996AD",
  faint: "#586479",
  cyan: "#2FD9C7",
  cyanDim: "rgba(47,217,199,0.13)",
  red: "#FF6161",
  redDim: "rgba(255,97,97,0.13)",
  amber: "#F5A83C",
  amberDim: "rgba(245,168,60,0.13)",
  green: "#3FDD8B",
  greenDim: "rgba(63,221,139,0.13)",
  blue: "#5B93FF",
  blueDim: "rgba(91,147,255,0.13)",
  violet: "#B18CFF",
  violetDim: "rgba(177,140,255,0.13)",
  gray: "#8996AD",
  grayDim: "rgba(137,150,173,0.13)",
};

const ACCESSORY_OPTIONS = ["Monitor 1", "Monitor 2", "CPU", "Keyboard", "Mouse", "Headset", "Speaker", "Printer"];
const DEPARTMENTS = ["OTT", "Web"];
const STATUSES = ["Assigned", "Available", "In Maintenance", "Decommissioned"];
const STATUS_COLOR = {
  Assigned: C.blue,
  Available: C.green,
  "In Maintenance": C.amber,
  Decommissioned: C.gray,
};
const DEPT_COLOR = { OTT: C.cyan, Web: C.violet };
const PAGE_SIZE = 8;
const STORAGE_KEY = "amrita-ott-asset-tracker-v1";
const IPV4_RE = /^((25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)\.){3}(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)$/;
const CURRENT_YEAR = new Date().getFullYear();

/* ------------------------------------------------------------------ */
/* Sample data                                                         */
/* ------------------------------------------------------------------ */
const uid = () => "AST-" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 6).toUpperCase();

function seedData() {
  return [
    {
      id: uid(),
      userName: "Aditya Menon",
      department: "OTT",
      computerName: "EDIT-STUDIO-01",
      ipAddress: "192.168.1.45",
      hardware: { cpu: "Apple M2 Max", ram: "64GB RAM", storage: "2TB SSD", os: "macOS Sonoma" },
      yearOfPurchase: 2023,
      accessories: [
        { name: "Monitor 1", assetNumber: "ALT-AST-1001", serialNumber: "SM24J812378972" },
        { name: "Monitor 2", assetNumber: "ALT-AST-1002", serialNumber: "SM24J812378973" },
        { name: "Keyboard", assetNumber: "ALT-AST-1003", serialNumber: "KB2245981" },
        { name: "Mouse", assetNumber: "ALT-AST-1004", serialNumber: "MS2245982" },
        { name: "Headset", assetNumber: "ALT-AST-1005", serialNumber: "HS9981276" },
      ],
      status: "Assigned",
      remarks: "Primary editing rig for the Original Series team. Color-calibrated monthly.",
    },
    {
      id: uid(),
      userName: "Priya Raghavan",
      department: "OTT",
      computerName: "ENCODE-RACK-03",
      ipAddress: "192.168.1.12",
      hardware: { cpu: "Intel Xeon W-2295", ram: "128GB RAM", storage: "8TB RAID", os: "Ubuntu 22.04 LTS" },
      yearOfPurchase: 2022,
      accessories: [
        { name: "Keyboard", assetNumber: "ATV-OTT-0231", serialNumber: "CN25645235" },
        { name: "Mouse", assetNumber: "ATV-OTT-0232", serialNumber: "CN25645236" },
      ],
      status: "Assigned",
      remarks: "Dedicated transcode node, cluster B. Handles 4K HDR encode queue.",
    },
    {
      id: uid(),
      userName: "Karthik Subramanian",
      department: "Web",
      computerName: "WEBDEV-STN-07",
      ipAddress: "192.168.2.31",
      hardware: { cpu: "Apple M3 Pro", ram: "36GB RAM", storage: "1TB SSD", os: "macOS Sequoia" },
      yearOfPurchase: 2024,
      accessories: [
        { name: "Monitor 1", assetNumber: "ALT-AST-2091", serialNumber: "SM26K120044" },
        { name: "Keyboard", assetNumber: "ALT-AST-2092", serialNumber: "KB2260114" },
        { name: "Mouse", assetNumber: "ALT-AST-2093", serialNumber: "MS2260115" },
        { name: "Speaker", assetNumber: "ALT-AST-2094", serialNumber: "SP7701245" },
      ],
      status: "Assigned",
      remarks: "Frontend dev station for the playback web app, React / Next.js.",
    },
    {
      id: uid(),
      userName: "Divya Nair",
      department: "OTT",
      computerName: "QA-BENCH-02",
      ipAddress: "192.168.1.88",
      hardware: { cpu: "Intel i7-13700", ram: "32GB RAM", storage: "1TB SSD", os: "Windows 11 Pro" },
      yearOfPurchase: 2021,
      accessories: [
        { name: "Monitor 1", assetNumber: "ALT-AST-3312", serialNumber: "SM21G550091" },
        { name: "Monitor 2", assetNumber: "ALT-AST-3313", serialNumber: "SM21G550092" },
        { name: "Headset", assetNumber: "ALT-AST-3314", serialNumber: "HS4471902" },
        { name: "Printer", assetNumber: "ALT-AST-3315", serialNumber: "PR8820013" },
      ],
      status: "In Maintenance",
      remarks: "GPU fan replacement pending — ticket #4482. ETA 3 business days.",
    },
    {
      id: uid(),
      userName: "Rohan Iyer",
      department: "Web",
      computerName: "MGFX-RENDER-05",
      ipAddress: "192.168.2.14",
      hardware: { cpu: "AMD Threadripper 3970X", ram: "256GB RAM", storage: "4TB NVMe", os: "Windows 11 Pro" },
      yearOfPurchase: 2020,
      accessories: [
        { name: "Monitor 1", assetNumber: "ALT-AST-4410", serialNumber: "SM20D330871" },
        { name: "Monitor 2", assetNumber: "ALT-AST-4411", serialNumber: "SM20D330872" },
        { name: "Speaker", assetNumber: "ALT-AST-4412", serialNumber: "SP1120087" },
      ],
      status: "Available",
      remarks: "Spare render node. Reassign after Q3 onboarding batch.",
    },
  ];
}

/* ------------------------------------------------------------------ */
/* Helpers                                                              */
/* ------------------------------------------------------------------ */
function hardwareLine(hw = {}) {
  return [hw.cpu, hw.ram, hw.storage, hw.os].filter(Boolean).join(" / ");
}

function normKey(k) {
  return String(k || "").trim().toLowerCase();
}

function parseKVList(str) {
  const map = {};
  if (!str) return map;
  String(str)
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean)
    .forEach((seg) => {
      const idx = seg.indexOf(":");
      if (idx === -1) return;
      const name = seg.slice(0, idx).trim();
      const value = seg.slice(idx + 1).trim();
      if (name) map[name.toLowerCase()] = value;
    });
  return map;
}

function rowToAsset(row) {
  const map = {};
  Object.entries(row).forEach(([k, v]) => (map[normKey(k)] = v));
  const get = (...aliases) => {
    for (const a of aliases) {
      if (map[a] !== undefined && map[a] !== "") return String(map[a]).trim();
    }
    return "";
  };

  const hwRaw = get("hardware properties", "hardware");
  const hwParts = hwRaw.split("/").map((s) => s.trim());
  const hardware = {
    cpu: hwParts[0] || "",
    ram: hwParts[1] || "",
    storage: hwParts[2] || "",
    os: hwParts[3] || "",
  };

  const deptRaw = get("department", "dept");
  const department = DEPARTMENTS.find((d) => d.toLowerCase() === deptRaw.toLowerCase()) || "OTT";

  const statusRaw = get("status");
  const status = STATUSES.find((s) => s.toLowerCase() === statusRaw.toLowerCase()) || "Available";

  const accNames = get("using machine accessories", "accessories")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  const assetMap = parseKVList(get("asset numbers", "asset number"));
  const serialMap = parseKVList(get("serial numbers", "serial number"));
  const accessories = accNames.map((name) => ({
    name,
    assetNumber: assetMap[name.toLowerCase()] || "",
    serialNumber: serialMap[name.toLowerCase()] || "",
  }));

  const yearRaw = get("year of purchase", "year", "purchase year");
  const yearOfPurchase = parseInt(yearRaw, 10) || CURRENT_YEAR;

  return {
    id: uid(),
    userName: get("user name", "username", "employee name", "name"),
    department,
    computerName: get("computer name", "computername", "pc name"),
    ipAddress: get("machine ip address", "ip address", "ip"),
    hardware,
    yearOfPurchase,
    accessories,
    status,
    remarks: get("updates / remarks", "updates/remarks", "remarks", "notes", "updates"),
  };
}

function assetToExportRow(a) {
  return {
    "User Name": a.userName,
    Department: a.department,
    "Computer Name": a.computerName,
    "Machine IP Address": a.ipAddress,
    "Hardware Properties": hardwareLine(a.hardware),
    "Year of Purchase": a.yearOfPurchase,
    "Using Machine Accessories": a.accessories.map((x) => x.name).join(", "),
    "Asset Numbers": a.accessories.map((x) => `${x.name}: ${x.assetNumber}`).join(", "),
    "Serial Numbers": a.accessories.map((x) => `${x.name}: ${x.serialNumber}`).join(", "),
    Status: a.status,
    "Updates / Remarks": a.remarks,
  };
}

/* ------------------------------------------------------------------ */
/* Small presentational bits                                           */
/* ------------------------------------------------------------------ */
function Badge({ color, dim, children, mono }) {
  return (
    <span
      className={"inline-flex items-center gap-1.5 px-2 py-0.5 text-xs rounded-sm border-l-2 whitespace-nowrap" + (mono ? " font-mono" : "")}
      style={{ background: dim, color, borderColor: color }}
    >
      {children}
    </span>
  );
}

function StatusBadge({ status }) {
  const color = STATUS_COLOR[status] || C.gray;
  return (
    <span className="inline-flex items-center gap-1.5 px-2 py-1 text-xs font-medium rounded-sm" style={{ background: `${color}20`, color }}>
      <span className="w-1.5 h-1.5 rounded-full" style={{ background: color }} />
      {status}
    </span>
  );
}

function IconBtn({ onClick, title, children, danger }) {
  return (
    <button
      onClick={onClick}
      title={title}
      className="p-2 rounded-md transition-colors"
      style={{ color: danger ? C.red : C.muted }}
      onMouseEnter={(e) => (e.currentTarget.style.background = danger ? C.redDim : C.panelAlt)}
      onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
    >
      {children}
    </button>
  );
}

function MetricCard({ icon, label, value, accent, sub }) {
  return (
    <div
      className="relative overflow-hidden rounded-md p-4 flex items-start justify-between"
      style={{ background: C.panel, border: `1px solid ${C.border}` }}
    >
      <div className="absolute left-0 top-0 bottom-0 w-1" style={{ background: accent }} />
      <div className="pl-2">
        <p className="text-xs uppercase tracking-wide" style={{ color: C.faint }}>{label}</p>
        <p className="text-2xl font-semibold mt-1 font-display" style={{ color: C.text }}>{value}</p>
        {sub && <p className="text-xs mt-1" style={{ color: C.muted }}>{sub}</p>}
      </div>
      <div className="p-2 rounded-md" style={{ background: `${accent}18`, color: accent }}>
        {icon}
      </div>
    </div>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  const color = toast.type === "error" ? C.red : toast.type === "info" ? C.blue : C.green;
  return (
    <div
      className="fixed bottom-5 right-5 z-50 px-4 py-3 rounded-md flex items-center gap-2 text-sm shadow-lg animate-toast"
      style={{ background: C.panel, border: `1px solid ${color}`, color: C.text }}
    >
      <span className="w-2 h-2 rounded-full" style={{ background: color }} />
      {toast.message}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Add / Edit modal                                                     */
/* ------------------------------------------------------------------ */
function AssetFormModal({ initial, onClose, onSave }) {
  const isEdit = !!initial;
  const [form, setForm] = useState(() => ({
    userName: initial?.userName || "",
    department: initial?.department || "OTT",
    computerName: initial?.computerName || "",
    ipAddress: initial?.ipAddress || "",
    cpu: initial?.hardware?.cpu || "",
    ram: initial?.hardware?.ram || "",
    storage: initial?.hardware?.storage || "",
    os: initial?.hardware?.os || "",
    yearOfPurchase: initial?.yearOfPurchase || CURRENT_YEAR,
    status: initial?.status || "Available",
    remarks: initial?.remarks || "",
  }));
  const [accState, setAccState] = useState(() => {
    const existing = {};
    (initial?.accessories || []).forEach((a) => (existing[a.name] = a));
    const state = {};
    ACCESSORY_OPTIONS.forEach((name) => {
      state[name] = {
        checked: !!existing[name],
        assetNumber: existing[name]?.assetNumber || "",
        serialNumber: existing[name]?.serialNumber || "",
      };
    });
    return state;
  });
  const [errors, setErrors] = useState({});

  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const setAcc = (name, patch) => setAccState((s) => ({ ...s, [name]: { ...s[name], ...patch } }));

  const validate = () => {
    const e = {};
    if (!form.userName.trim()) e.userName = "User name is required.";
    if (!form.computerName.trim()) e.computerName = "Computer name is required.";
    if (!IPV4_RE.test(form.ipAddress.trim())) e.ipAddress = "Enter a valid IPv4 address, e.g. 192.168.1.45.";
    const yr = Number(form.yearOfPurchase);
    if (!yr || yr < 1990 || yr > CURRENT_YEAR + 1) e.yearOfPurchase = `Enter a year between 1990 and ${CURRENT_YEAR + 1}.`;
    Object.entries(accState).forEach(([name, v]) => {
      if (v.checked && (!v.assetNumber.trim() || !v.serialNumber.trim())) {
        e[`acc_${name}`] = "Asset number and serial number are required for a selected accessory.";
      }
    });
    setErrors(e);
    return Object.keys(e).length === 0;
  };

  const submit = () => {
    if (!validate()) return;
    const accessories = ACCESSORY_OPTIONS.filter((n) => accState[n].checked).map((n) => ({
      name: n,
      assetNumber: accState[n].assetNumber.trim(),
      serialNumber: accState[n].serialNumber.trim(),
    }));
    onSave({
      id: initial?.id || uid(),
      userName: form.userName.trim(),
      department: form.department,
      computerName: form.computerName.trim(),
      ipAddress: form.ipAddress.trim(),
      hardware: { cpu: form.cpu.trim(), ram: form.ram.trim(), storage: form.storage.trim(), os: form.os.trim() },
      yearOfPurchase: Number(form.yearOfPurchase),
      accessories,
      status: form.status,
      remarks: form.remarks.trim(),
    });
  };

  const inputStyle = { background: C.panelAlt, border: `1px solid ${C.border}`, color: C.text };
  const label = "text-xs font-medium mb-1.5 block";

  return (
    <div className="fixed inset-0 z-50 flex items-start md:items-center justify-center p-4 overflow-y-auto" style={{ background: "rgba(4,7,14,0.72)" }} onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="w-full rounded-md my-6 animate-modal" style={{ maxWidth: "780px", background: C.panel, border: `1px solid ${C.border}` }}>
        <div className="flex items-center justify-between px-6 py-4" style={{ borderBottom: `1px solid ${C.border}` }}>
          <h2 className="text-lg font-semibold font-display" style={{ color: C.text }}>{isEdit ? "Edit asset" : "Add new asset"}</h2>
          <IconBtn onClick={onClose} title="Close"><X size={18} /></IconBtn>
        </div>

        <div className="px-6 py-5 space-y-6 max-h-[70vh] overflow-y-auto">
          {/* Identity */}
          <section>
            <p className="text-xs uppercase tracking-wide mb-3" style={{ color: C.faint }}>Assignment</p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className={label} style={{ color: C.muted }}>User name</label>
                <input className="w-full rounded-md px-3 py-2 text-sm outline-none" style={inputStyle} value={form.userName} onChange={(e) => set("userName", e.target.value)} placeholder="e.g. Aditya Menon" />
                {errors.userName && <p className="text-xs mt-1" style={{ color: C.red }}>{errors.userName}</p>}
              </div>
              <div>
                <label className={label} style={{ color: C.muted }}>Department</label>
                <select className="w-full rounded-md px-3 py-2 text-sm outline-none" style={inputStyle} value={form.department} onChange={(e) => set("department", e.target.value)}>
                  {DEPARTMENTS.map((d) => <option key={d} value={d}>{d}</option>)}
                </select>
              </div>
              <div>
                <label className={label} style={{ color: C.muted }}>Computer name</label>
                <input className="w-full rounded-md px-3 py-2 text-sm font-mono outline-none" style={inputStyle} value={form.computerName} onChange={(e) => set("computerName", e.target.value.toUpperCase())} placeholder="EDIT-STUDIO-01" />
                {errors.computerName && <p className="text-xs mt-1" style={{ color: C.red }}>{errors.computerName}</p>}
              </div>
              <div>
                <label className={label} style={{ color: C.muted }}>Machine IP address</label>
                <input className="w-full rounded-md px-3 py-2 text-sm font-mono outline-none" style={inputStyle} value={form.ipAddress} onChange={(e) => set("ipAddress", e.target.value)} placeholder="192.168.1.45" />
                {errors.ipAddress && <p className="text-xs mt-1" style={{ color: C.red }}>{errors.ipAddress}</p>}
              </div>
              <div>
                <label className={label} style={{ color: C.muted }}>Status</label>
                <select className="w-full rounded-md px-3 py-2 text-sm outline-none" style={inputStyle} value={form.status} onChange={(e) => set("status", e.target.value)}>
                  {STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
              <div>
                <label className={label} style={{ color: C.muted }}>Year of purchase</label>
                <input type="number" className="w-full rounded-md px-3 py-2 text-sm outline-none" style={inputStyle} value={form.yearOfPurchase} onChange={(e) => set("yearOfPurchase", e.target.value)} />
                {errors.yearOfPurchase && <p className="text-xs mt-1" style={{ color: C.red }}>{errors.yearOfPurchase}</p>}
              </div>
            </div>
          </section>

          {/* Hardware */}
          <section>
            <p className="text-xs uppercase tracking-wide mb-3" style={{ color: C.faint }}>Hardware properties</p>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              {[["cpu", "CPU"], ["ram", "RAM"], ["storage", "Storage"], ["os", "Operating system"]].map(([k, l]) => (
                <div key={k}>
                  <label className={label} style={{ color: C.muted }}>{l}</label>
                  <input className="w-full rounded-md px-3 py-2 text-sm outline-none" style={inputStyle} value={form[k]} onChange={(e) => set(k, e.target.value)} placeholder={k === "cpu" ? "Apple M2 Max" : k === "ram" ? "64GB RAM" : k === "storage" ? "2TB SSD" : "macOS"} />
                </div>
              ))}
            </div>
          </section>

          {/* Accessories */}
          <section>
            <p className="text-xs uppercase tracking-wide mb-3" style={{ color: C.faint }}>Using machine accessories</p>
            <div className="space-y-2">
              {ACCESSORY_OPTIONS.map((name) => (
                <div key={name} className="rounded-md p-3" style={{ background: C.panelAlt, border: `1px solid ${accState[name].checked ? C.cyan + "55" : C.border}` }}>
                  <label className="flex items-center gap-2 cursor-pointer select-none">
                    <input type="checkbox" className="w-4 h-4" checked={accState[name].checked} onChange={(e) => setAcc(name, { checked: e.target.checked })} />
                    <span className="text-sm font-medium" style={{ color: C.text }}>{name}</span>
                  </label>
                  {accState[name].checked && (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mt-2 pl-6">
                      <input className="rounded-md px-3 py-1.5 text-xs font-mono outline-none" style={inputStyle} placeholder="Asset number, e.g. ALT-AST-1001" value={accState[name].assetNumber} onChange={(e) => setAcc(name, { assetNumber: e.target.value })} />
                      <input className="rounded-md px-3 py-1.5 text-xs font-mono outline-none" style={inputStyle} placeholder="Serial number" value={accState[name].serialNumber} onChange={(e) => setAcc(name, { serialNumber: e.target.value })} />
                    </div>
                  )}
                  {errors[`acc_${name}`] && <p className="text-xs mt-1.5 pl-6" style={{ color: C.red }}>{errors[`acc_${name}`]}</p>}
                </div>
              ))}
            </div>
          </section>

          {/* Remarks */}
          <section>
            <label className={label} style={{ color: C.muted }}>Updates / remarks</label>
            <textarea rows={3} className="w-full rounded-md px-3 py-2 text-sm outline-none resize-none" style={inputStyle} value={form.remarks} onChange={(e) => set("remarks", e.target.value)} placeholder="Maintenance notes, assignment history, hardware condition…" />
          </section>
        </div>

        <div className="flex items-center justify-end gap-3 px-6 py-4" style={{ borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} className="px-4 py-2 text-sm rounded-md" style={{ color: C.muted }}>Cancel</button>
          <button onClick={submit} className="px-4 py-2 text-sm rounded-md font-medium" style={{ background: C.cyan, color: "#04211D" }}>
            {isEdit ? "Save changes" : "Add asset"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* View details modal                                                   */
/* ------------------------------------------------------------------ */
function ViewModal({ asset, onClose }) {
  if (!asset) return null;
  const row = "flex justify-between gap-4 py-2 text-sm";
  return (
    <div className="fixed inset-0 z-50 flex items-start md:items-center justify-center p-4 overflow-y-auto" style={{ background: "rgba(4,7,14,0.72)" }} onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="w-full rounded-md my-6 animate-modal" style={{ maxWidth: "680px", background: C.panel, border: `1px solid ${C.border}` }}>
        <div className="flex items-center justify-between px-6 py-4" style={{ borderBottom: `1px solid ${C.border}` }}>
          <div>
            <h2 className="text-lg font-semibold font-display" style={{ color: C.text }}>{asset.userName}</h2>
            <p className="text-xs font-mono mt-0.5" style={{ color: C.muted }}>{asset.computerName}</p>
          </div>
          <IconBtn onClick={onClose} title="Close"><X size={18} /></IconBtn>
        </div>
        <div className="px-6 py-4 max-h-[70vh] overflow-y-auto">
          <div className="grid grid-cols-2 gap-x-6" style={{ borderBottom: `1px solid ${C.border}` }}>
            <div className={row} style={{ color: C.muted }}><span>Department</span><Badge color={DEPT_COLOR[asset.department]} dim={`${DEPT_COLOR[asset.department]}20`}>{asset.department}</Badge></div>
            <div className={row} style={{ color: C.muted }}><span>Status</span><StatusBadge status={asset.status} /></div>
            <div className={row} style={{ color: C.muted }}><span>IP address</span><span className="font-mono" style={{ color: C.text }}>{asset.ipAddress}</span></div>
            <div className={row} style={{ color: C.muted }}><span>Year of purchase</span><span style={{ color: C.text }}>{asset.yearOfPurchase}</span></div>
          </div>

          <p className="text-xs uppercase tracking-wide mt-4 mb-2" style={{ color: C.faint }}>Hardware</p>
          <p className="text-sm font-mono" style={{ color: C.text }}>{hardwareLine(asset.hardware) || "—"}</p>

          <p className="text-xs uppercase tracking-wide mt-5 mb-2" style={{ color: C.faint }}>Accessories &amp; tags</p>
          {asset.accessories.length === 0 ? (
            <p className="text-sm" style={{ color: C.muted }}>No accessories assigned.</p>
          ) : (
            <div className="rounded-md overflow-hidden" style={{ border: `1px solid ${C.border}` }}>
              {asset.accessories.map((a, i) => (
                <div key={a.name} className="grid grid-cols-3 gap-2 px-3 py-2 text-xs" style={{ background: i % 2 ? C.panelAlt : "transparent" }}>
                  <span style={{ color: C.text }}>{a.name}</span>
                  <span className="font-mono" style={{ color: C.cyan }}>{a.assetNumber || "—"}</span>
                  <span className="font-mono" style={{ color: C.muted }}>{a.serialNumber || "—"}</span>
                </div>
              ))}
            </div>
          )}

          <p className="text-xs uppercase tracking-wide mt-5 mb-2" style={{ color: C.faint }}>Updates / remarks</p>
          <p className="text-sm" style={{ color: C.text }}>{asset.remarks || "No remarks on file."}</p>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Delete confirm modal                                                 */
/* ------------------------------------------------------------------ */
function DeleteModal({ asset, onClose, onConfirm }) {
  if (!asset) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(4,7,14,0.72)" }} onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="w-full rounded-md animate-modal" style={{ maxWidth: "420px", background: C.panel, border: `1px solid ${C.border}` }}>
        <div className="p-6">
          <div className="w-10 h-10 rounded-md flex items-center justify-center mb-4" style={{ background: C.redDim, color: C.red }}>
            <AlertTriangle size={20} />
          </div>
          <h3 className="text-base font-semibold" style={{ color: C.text }}>Delete this asset record?</h3>
          <p className="text-sm mt-2" style={{ color: C.muted }}>
            <span className="font-mono" style={{ color: C.text }}>{asset.computerName}</span> assigned to {asset.userName} will be permanently removed. This can't be undone.
          </p>
        </div>
        <div className="flex items-center justify-end gap-3 px-6 py-4" style={{ borderTop: `1px solid ${C.border}` }}>
          <button onClick={onClose} className="px-4 py-2 text-sm rounded-md" style={{ color: C.muted }}>Cancel</button>
          <button onClick={onConfirm} className="px-4 py-2 text-sm rounded-md font-medium" style={{ background: C.red, color: "#2A0808" }}>Delete asset</button>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Row action menu                                                      */
/* ------------------------------------------------------------------ */
function RowMenu({ onView, onEdit, onDelete, open, setOpen }) {
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open, setOpen]);

  return (
    <div className="relative" ref={ref}>
      <IconBtn onClick={() => setOpen(!open)} title="Actions"><MoreVertical size={16} /></IconBtn>
      {open && (
        <div className="absolute right-0 top-9 z-30 w-40 rounded-md py-1 animate-modal" style={{ background: C.panelAlt, border: `1px solid ${C.border}`, boxShadow: "0 8px 24px rgba(0,0,0,0.4)" }}>
          <button onClick={() => { onView(); setOpen(false); }} className="w-full flex items-center gap-2 px-3 py-2 text-xs text-left hover:brightness-125" style={{ color: C.text }}>
            <Eye size={14} /> View details
          </button>
          <button onClick={() => { onEdit(); setOpen(false); }} className="w-full flex items-center gap-2 px-3 py-2 text-xs text-left hover:brightness-125" style={{ color: C.text }}>
            <Pencil size={14} /> Edit
          </button>
          <button onClick={() => { onDelete(); setOpen(false); }} className="w-full flex items-center gap-2 px-3 py-2 text-xs text-left" style={{ color: C.red }}>
            <Trash2 size={14} /> Delete
          </button>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Main app                                                             */
/* ------------------------------------------------------------------ */
export default function AmritaAssetTracker() {
  const [assets, setAssets] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [search, setSearch] = useState("");
  const [filterDept, setFilterDept] = useState("All");
  const [filterStatus, setFilterStatus] = useState("All");
  const [sort, setSort] = useState({ key: "userName", dir: "asc" });
  const [page, setPage] = useState(1);
  const [formModal, setFormModal] = useState(null); // { mode, data }
  const [viewAsset, setViewAsset] = useState(null);
  const [deleteAsset, setDeleteAsset] = useState(null);
  const [openMenuId, setOpenMenuId] = useState(null);
  const [toast, setToast] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef(null);

  // load / persist
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      setAssets(raw ? JSON.parse(raw) : seedData());
    } catch {
      setAssets(seedData());
    }
    setLoaded(true);
  }, []);
  useEffect(() => {
    if (!loaded) return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(assets));
  }, [assets, loaded]);

  const showToast = useCallback((message, type = "success") => {
    setToast({ message, type });
    const t = setTimeout(() => setToast(null), 3200);
    return () => clearTimeout(t);
  }, []);

  // derived
  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return assets.filter((a) => {
      if (filterDept !== "All" && a.department !== filterDept) return false;
      if (filterStatus !== "All" && a.status !== filterStatus) return false;
      if (!q) return true;
      const hay = [
        a.userName,
        a.ipAddress,
        a.computerName,
        ...a.accessories.map((x) => x.assetNumber),
        ...a.accessories.map((x) => x.serialNumber),
      ].join(" ").toLowerCase();
      return hay.includes(q);
    });
  }, [assets, search, filterDept, filterStatus]);

  const sorted = useMemo(() => {
    const arr = [...filtered];
    arr.sort((a, b) => {
      let va = a[sort.key], vb = b[sort.key];
      if (typeof va === "string") va = va.toLowerCase();
      if (typeof vb === "string") vb = vb.toLowerCase();
      if (va < vb) return sort.dir === "asc" ? -1 : 1;
      if (va > vb) return sort.dir === "asc" ? 1 : -1;
      return 0;
    });
    return arr;
  }, [filtered, sort]);

  const totalPages = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
  const pageClamped = Math.min(page, totalPages);
  const pageRows = sorted.slice((pageClamped - 1) * PAGE_SIZE, pageClamped * PAGE_SIZE);

  useEffect(() => { setPage(1); }, [search, filterDept, filterStatus]);

  const metrics = useMemo(() => ({
    total: assets.length,
    active: assets.filter((a) => a.status === "Assigned").length,
    maintenance: assets.filter((a) => a.status === "In Maintenance").length,
    departments: new Set(assets.map((a) => a.department)).size,
  }), [assets]);

  const handleSort = (key) => setSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "asc" }));

  const saveAsset = (asset) => {
    setAssets((prev) => {
      const exists = prev.some((a) => a.id === asset.id);
      return exists ? prev.map((a) => (a.id === asset.id ? asset : a)) : [asset, ...prev];
    });
    showToast(formModal?.mode === "edit" ? "Asset updated." : "Asset added.");
    setFormModal(null);
  };

  const doDelete = () => {
    setAssets((prev) => prev.filter((a) => a.id !== deleteAsset.id));
    showToast("Asset deleted.", "info");
    setDeleteAsset(null);
  };

  const exportExcel = () => {
    const rows = sorted.map(assetToExportRow);
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 20 }, { wch: 10 }, { wch: 18 }, { wch: 15 }, { wch: 36 }, { wch: 10 }, { wch: 28 }, { wch: 36 }, { wch: 36 }, { wch: 14 }, { wch: 40 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Assets");
    XLSX.writeFile(wb, "Amrita_OTT_Asset_Inventory.xlsx");
    showToast(`Exported ${rows.length} record${rows.length === 1 ? "" : "s"}.`);
  };

  const importFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const imported = json.map(rowToAsset).filter((a) => a.userName || a.computerName);
        if (imported.length === 0) {
          showToast("No usable rows found in that file.", "error");
          return;
        }
        setAssets((prev) => [...imported, ...prev]);
        showToast(`Imported ${imported.length} record${imported.length === 1 ? "" : "s"}.`);
      } catch {
        showToast("Couldn't read that file. Check the format and try again.", "error");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const clearFilters = () => { setSearch(""); setFilterDept("All"); setFilterStatus("All"); };

  const SortHeader = ({ label: l, k, className = "" }) => (
    <th
      className={"px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide cursor-pointer select-none whitespace-nowrap " + className}
      style={{ color: sort.key === k ? C.cyan : C.faint }}
      onClick={() => handleSort(k)}
    >
      <span className="inline-flex items-center gap-1">
        {l}
        {sort.key === k && (sort.dir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />)}
      </span>
    </th>
  );

  return (
    <div className="min-h-screen w-full font-sans" style={{ background: C.bg, color: C.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
        .font-sans { font-family: 'Inter', system-ui, sans-serif; }
        .font-display { font-family: 'Space Grotesk', system-ui, sans-serif; }
        .font-mono { font-family: 'JetBrains Mono', ui-monospace, monospace; }
        * { box-sizing: border-box; }
        input:focus, select:focus, textarea:focus, button:focus-visible { outline: 2px solid ${C.cyan}; outline-offset: 1px; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 4px; }
        @keyframes pulseDot { 0%,100% { opacity:1; } 50% { opacity:0.35; } }
        .pulse-dot { animation: pulseDot 2s ease-in-out infinite; }
        @keyframes modalIn { from { opacity:0; transform: translateY(6px); } to { opacity:1; transform: translateY(0); } }
        .animate-modal { animation: modalIn 0.16s ease-out; }
        @keyframes toastIn { from { opacity:0; transform: translateY(8px);} to {opacity:1; transform:translateY(0);} }
        .animate-toast { animation: toastIn 0.2s ease-out; }
        @media (prefers-reduced-motion: reduce) { .pulse-dot, .animate-modal, .animate-toast { animation: none; } }
      `}</style>

      {/* Header */}
      <header className="sticky top-0 z-40" style={{ background: C.header, borderBottom: `1px solid ${C.border}` }}>
        <div className="max-w-7xl mx-auto px-4 md:px-8 py-3 flex flex-col lg:flex-row lg:items-center gap-3 lg:gap-6">
          <div className="flex items-center gap-3 shrink-0">
            <div className="w-9 h-9 rounded-md flex items-center justify-center" style={{ background: C.cyanDim, color: C.cyan }}>
              <Cast size={18} />
            </div>
            <div>
              <h1 className="text-base font-semibold font-display leading-tight" style={{ color: C.text }}>Amrita Live OTT <span style={{ color: C.faint, fontWeight: 500 }}>· Asset Tracker</span></h1>
              <div className="flex items-center gap-1.5 mt-0.5">
                <span className="w-1.5 h-1.5 rounded-full pulse-dot" style={{ background: C.green }} />
                <span className="text-xs" style={{ color: C.muted }}>All systems operational</span>
              </div>
            </div>
          </div>

          <div className="flex-1 flex items-center gap-2 min-w-0" onDragOver={(e) => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={(e) => { e.preventDefault(); setDragOver(false); importFile(e.dataTransfer.files?.[0]); }}>
            <div className="relative flex-1 min-w-0" style={{ maxWidth: "480px" }}>
              <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: C.faint }} />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="Search by user, IP, asset or serial number…"
                className="w-full rounded-md pl-9 pr-3 py-2 text-sm outline-none"
                style={{ background: C.panel, border: `1px solid ${dragOver ? C.cyan : C.border}`, color: C.text }}
              />
            </div>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => { importFile(e.target.files?.[0]); e.target.value = ""; }} />
            <button onClick={() => fileInputRef.current?.click()} className="flex items-center gap-1.5 px-3 py-2 text-sm rounded-md" style={{ background: C.panel, border: `1px solid ${C.border}`, color: C.text }}>
              <Upload size={15} /> <span className="hidden md:inline">Import</span>
            </button>
            <button onClick={exportExcel} className="flex items-center gap-1.5 px-3 py-2 text-sm rounded-md" style={{ background: C.panel, border: `1px solid ${C.border}`, color: C.text }}>
              <Download size={15} /> <span className="hidden md:inline">Export</span>
            </button>
            <button onClick={() => setFormModal({ mode: "add" })} className="flex items-center gap-1.5 px-3 py-2 text-sm rounded-md font-medium" style={{ background: C.cyan, color: "#04211D" }}>
              <Plus size={15} /> <span>Add asset</span>
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 md:px-8 py-6">
        {/* Metrics */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
          <MetricCard icon={<Server size={18} />} label="Total assets tracked" value={metrics.total} accent={C.cyan} />
          <MetricCard icon={<UserCheck size={18} />} label="Active / assigned devices" value={metrics.active} accent={C.blue} />
          <MetricCard icon={<Wrench size={18} />} label="Devices in maintenance" value={metrics.maintenance} accent={C.amber} />
          <MetricCard icon={<Building2 size={18} />} label="Total departments" value={metrics.departments} accent={C.violet} />
        </div>

        {/* Filter bar */}
        <div className="flex flex-wrap items-center gap-3 mb-4">
          <div className="flex items-center gap-1.5 text-xs" style={{ color: C.faint }}>
            <Filter size={13} /> Filters
          </div>
          <select value={filterDept} onChange={(e) => setFilterDept(e.target.value)} className="rounded-md px-3 py-1.5 text-sm outline-none" style={{ background: C.panel, border: `1px solid ${C.border}`, color: C.text }}>
            <option value="All">All departments</option>
            {DEPARTMENTS.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
          <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} className="rounded-md px-3 py-1.5 text-sm outline-none" style={{ background: C.panel, border: `1px solid ${C.border}`, color: C.text }}>
            <option value="All">All statuses</option>
            {STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          {(search || filterDept !== "All" || filterStatus !== "All") && (
            <button onClick={clearFilters} className="text-xs px-2 py-1.5 rounded-md" style={{ color: C.cyan }}>Clear filters</button>
          )}
          <span className="text-xs ml-auto" style={{ color: C.faint }}>{sorted.length} of {assets.length} record{assets.length === 1 ? "" : "s"}</span>
        </div>

        {/* Table */}
        <div className="rounded-md overflow-hidden" style={{ background: C.panel, border: `1px solid ${C.border}` }}>
          <div className="overflow-x-auto">
            <table className="w-full text-sm border-collapse">
              <thead>
                <tr style={{ background: C.header, borderBottom: `1px solid ${C.border}` }}>
                  <SortHeader label="User" k="userName" className="sticky left-0 z-20" />
                  <SortHeader label="Department" k="department" />
                  <SortHeader label="Computer" k="computerName" />
                  <SortHeader label="IP address" k="ipAddress" />
                  <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide" style={{ color: C.faint }}>Hardware</th>
                  <SortHeader label="Year" k="yearOfPurchase" />
                  <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide" style={{ color: C.faint }}>Accessories</th>
                  <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide" style={{ color: C.faint }}>Tags</th>
                  <SortHeader label="Status" k="status" />
                  <th className="px-4 py-3 text-left text-xs font-semibold uppercase tracking-wide" style={{ color: C.faint }}>Remarks</th>
                  <th className="px-4 py-3 text-right text-xs font-semibold uppercase tracking-wide" style={{ color: C.faint }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {pageRows.map((a, idx) => (
                  <tr key={a.id} style={{ background: idx % 2 ? C.panelAlt : "transparent", borderBottom: `1px solid ${C.border}` }}>
                    <td className="px-4 py-3 whitespace-nowrap sticky left-0" style={{ background: idx % 2 ? C.panelAlt : C.panel, color: C.text, fontWeight: 500 }}>{a.userName}</td>
                    <td className="px-4 py-3"><Badge color={DEPT_COLOR[a.department]} dim={`${DEPT_COLOR[a.department]}20`}>{a.department}</Badge></td>
                    <td className="px-4 py-3 font-mono whitespace-nowrap" style={{ color: C.text }}>{a.computerName}</td>
                    <td className="px-4 py-3 font-mono whitespace-nowrap" style={{ color: C.muted }}>{a.ipAddress}</td>
                    <td className="px-4 py-3" style={{ minWidth: "220px" }}>
                      <p className="text-xs" style={{ color: C.text }}>{a.hardware.cpu || "—"}</p>
                      <p className="text-xs mt-0.5" style={{ color: C.faint }}>{[a.hardware.ram, a.hardware.storage, a.hardware.os].filter(Boolean).join(" · ")}</p>
                    </td>
                    <td className="px-4 py-3 whitespace-nowrap" style={{ color: C.muted }}>{a.yearOfPurchase}</td>
                    <td className="px-4 py-3" style={{ minWidth: "180px" }}>
                      <div className="flex flex-wrap gap-1">
                        {a.accessories.slice(0, 3).map((x) => (
                          <span key={x.name} className="text-xs px-1.5 py-0.5 rounded-sm" style={{ background: C.grayDim, color: C.muted }}>{x.name}</span>
                        ))}
                        {a.accessories.length > 3 && (
                          <span className="text-xs px-1.5 py-0.5 rounded-sm" style={{ background: C.grayDim, color: C.faint }}>+{a.accessories.length - 3}</span>
                        )}
                        {a.accessories.length === 0 && <span className="text-xs" style={{ color: C.faint }}>—</span>}
                      </div>
                    </td>
                    <td className="px-4 py-3 whitespace-nowrap">
                      {a.accessories.length > 0 ? (
                        <button onClick={() => setViewAsset(a)} className="text-xs font-mono" style={{ color: C.cyan }}>{a.accessories.length} tagged</button>
                      ) : <span className="text-xs" style={{ color: C.faint }}>—</span>}
                    </td>
                    <td className="px-4 py-3"><StatusBadge status={a.status} /></td>
                    <td className="px-4 py-3 text-xs" style={{ color: C.muted, maxWidth: "220px" }}>
                      <span className="block truncate" title={a.remarks}>{a.remarks || "—"}</span>
                    </td>
                    <td className="px-4 py-3 text-right">
                      <RowMenu
                        open={openMenuId === a.id}
                        setOpen={(v) => setOpenMenuId(v ? a.id : null)}
                        onView={() => setViewAsset(a)}
                        onEdit={() => setFormModal({ mode: "edit", data: a })}
                        onDelete={() => setDeleteAsset(a)}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {sorted.length === 0 && (
            <div className="flex flex-col items-center justify-center py-16 px-4 text-center">
              <div className="w-12 h-12 rounded-md flex items-center justify-center mb-3" style={{ background: C.panelAlt, color: C.faint }}>
                <Inbox size={22} />
              </div>
              <p className="text-sm font-medium" style={{ color: C.text }}>No assets match your filters</p>
              <p className="text-xs mt-1" style={{ color: C.muted }}>Try a different search term, or clear the department and status filters.</p>
              <button onClick={clearFilters} className="mt-4 text-xs px-3 py-1.5 rounded-md" style={{ background: C.cyanDim, color: C.cyan }}>Clear filters</button>
            </div>
          )}

          {/* Pagination */}
          {sorted.length > 0 && (
            <div className="flex items-center justify-between px-4 py-3 flex-wrap gap-2" style={{ borderTop: `1px solid ${C.border}` }}>
              <span className="text-xs" style={{ color: C.faint }}>
                Showing {(pageClamped - 1) * PAGE_SIZE + 1}–{Math.min(pageClamped * PAGE_SIZE, sorted.length)} of {sorted.length}
              </span>
              <div className="flex items-center gap-1">
                <IconBtn onClick={() => setPage((p) => Math.max(1, p - 1))} title="Previous page"><ChevronLeft size={16} /></IconBtn>
                {Array.from({ length: totalPages }, (_, i) => i + 1)
                  .filter((n) => n === 1 || n === totalPages || Math.abs(n - pageClamped) <= 1)
                  .reduce((acc, n, i, arr) => { if (i > 0 && n - arr[i - 1] > 1) acc.push("…"); acc.push(n); return acc; }, [])
                  .map((n, i) => n === "…" ? (
                    <span key={`e${i}`} className="px-2 text-xs" style={{ color: C.faint }}>…</span>
                  ) : (
                    <button key={n} onClick={() => setPage(n)} className="w-8 h-8 rounded-md text-xs font-medium" style={{ background: n === pageClamped ? C.cyanDim : "transparent", color: n === pageClamped ? C.cyan : C.muted }}>
                      {n}
                    </button>
                  ))}
                <IconBtn onClick={() => setPage((p) => Math.min(totalPages, p + 1))} title="Next page"><ChevronRight size={16} /></IconBtn>
              </div>
            </div>
          )}
        </div>
      </main>

      {formModal && <AssetFormModal initial={formModal.data} onClose={() => setFormModal(null)} onSave={saveAsset} />}
      {viewAsset && <ViewModal asset={viewAsset} onClose={() => setViewAsset(null)} />}
      {deleteAsset && <DeleteModal asset={deleteAsset} onClose={() => setDeleteAsset(null)} onConfirm={doDelete} />}
      <Toast toast={toast} />
    </div>
  );
}
