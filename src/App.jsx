import React, { useState, useEffect, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Clock, LogIn, LogOut, Users, CheckCircle2, XCircle, AlertTriangle,
  Calendar, Search, Filter, Download, FileText, Settings as SettingsIcon,
  LayoutDashboard, ClipboardList, ShieldCheck, LogOut as LogOutIcon,
  Edit3, PlusCircle, ChevronDown, X, Building2, History, Ban, Check,
  Wallet, DollarSign, TrendingUp, CalendarDays, Eye, Printer, AlertOctagon,
  Trash2,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/* Constants & storage helpers                                        */
/* ------------------------------------------------------------------ */

const KEYS = {
  employees: "ats:employees",
  attendance: "ats:attendance",
  audit: "ats:audit",
  settings: "ats:settings",
  holidays: "ats:holidays",
};

const DAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

const DEFAULT_SETTINGS = {
  shiftStart: "08:00",
  shiftEnd: "17:00",
  standardHours: 8,
  gracePeriodMinutes: 10,
  otThresholdHours: 8,
  workingDays: [1, 2, 3, 4, 5],
  // Payroll settings
  otMultiplier: 1.25,       // overtime is paid at this multiple of the hourly rate
  pagibigRate: 0.02,        // 2%
  pagibigCap: 10000,        // Pag-IBIG is computed on gross capped at this amount
  sssRate: 0.05,            // 5%
  sssCap: 35000,            // SSS is computed on gross capped at this amount
  philhealthRate: 0.025,    // 2.5%
};

function seedEmployees() {
  return [
    { id: "u-admin", employeeId: "ADM-001", name: "Dana Cruz", username: "admin", password: "admin123", role: "admin", email: "dana@company.com", department: "Operations", position: "HR Administrator", dateHired: "2021-03-01", status: "active", expectedClockIn: null, gracePeriodOverride: null, payRateType: "daily", payRate: 900, rateHistory: [] },
    { id: "u-1", employeeId: "EMP-1001", name: "Marco Villanueva", username: "marco", password: "employee123", role: "employee", email: "marco@company.com", department: "Engineering", position: "Software Engineer", dateHired: "2022-06-15", status: "active", expectedClockIn: null, gracePeriodOverride: null, payRateType: "daily", payRate: 700, rateHistory: [] },
    { id: "u-2", employeeId: "EMP-1002", name: "Liza Fernandez", username: "liza", password: "employee123", role: "employee", email: "liza@company.com", department: "Sales", position: "Account Executive", dateHired: "2023-01-10", status: "active", expectedClockIn: "08:30", gracePeriodOverride: null, payRateType: "hourly", payRate: 87.5, rateHistory: [] },
    { id: "u-3", employeeId: "EMP-1003", name: "Ravi Santos", username: "ravi", password: "employee123", role: "employee", email: "ravi@company.com", department: "Engineering", position: "QA Analyst", dateHired: "2020-11-20", status: "inactive", expectedClockIn: null, gracePeriodOverride: null, payRateType: "daily", payRate: 658, rateHistory: [] },
    { id: "u-4", employeeId: "EMP-1004", name: "Peachy Domingo", username: "peachy", password: "employee123", role: "employee", email: "peachy@company.com", department: "Sales", position: "Sales Associate", dateHired: "2024-02-05", status: "active", expectedClockIn: null, gracePeriodOverride: null, payRateType: "hourly", payRate: 80, rateHistory: [] },
  ];
}

function seedHolidays() {
  return [];
}

function offsetDateStr(daysAgo) {
  const d = new Date();
  d.setDate(d.getDate() - daysAgo);
  return d.toLocaleDateString("en-CA");
}

function seedAttendance() {
  const mk = (employeeId, daysAgo, inHM, outHM, note) => {
    const date = offsetDateStr(daysAgo);
    const actualClockIn = `${date}T${inHM}:00`;
    const actualClockOut = outHM ? `${date}T${outHM}:00` : null;
    return { employeeId, date, actualClockIn, actualClockOut, note };
  };
  const raw = [
    mk("EMP-1001", 4, "07:58", "17:10"),
    mk("EMP-1001", 3, "08:22", "18:45"),
    mk("EMP-1001", 2, "08:05", "17:00"),
    mk("EMP-1002", 4, "08:45", "17:30"),
    mk("EMP-1002", 3, "08:20", "19:20"),
    mk("EMP-1004", 3, "09:15", "17:05"),
    mk("EMP-1004", 2, "07:55", "17:00"),
  ];
  return raw.map((r, i) => buildRecordFromTimes(r, DEFAULT_SETTINGS, `seed-${i}`, seedEmployees()));
}

function seedAudit() {
  return [
    { id: "a-1", timestamp: new Date().toISOString(), actor: "System", action: "SYSTEM_INIT", details: "Attendance system initialized with default settings and demo data." },
  ];
}

/* ------------------------------------------------------------------ */
/* Time / calculation helpers                                         */
/* ------------------------------------------------------------------ */

function getEmployeeSchedule(employee, settings) {
  return {
    scheduledClockIn: employee?.expectedClockIn || settings.shiftStart,
    grace: employee?.gracePeriodOverride ?? settings.gracePeriodMinutes,
  };
}

function buildRecordFromTimes(r, settings, id, employees) {
  const employee = employees.find((e) => e.employeeId === r.employeeId);
  const { scheduledClockIn, grace } = getEmployeeSchedule(employee, settings);
  const scheduled = new Date(`${r.date}T${scheduledClockIn}:00`);
  const actualIn = new Date(r.actualClockIn);
  const lateRaw = Math.round((actualIn - scheduled) / 60000);
  const lateMinutes = Math.max(0, lateRaw - grace);
  const status = lateMinutes > 0 ? "LATE" : "ON_TIME";

  let totalHours = null, regularHours = null, otHours = 0, otStatus = "NONE";
  if (r.actualClockOut) {
    totalHours = Math.round(((new Date(r.actualClockOut) - actualIn) / 3600000) * 100) / 100;
    regularHours = Math.min(totalHours, settings.standardHours);
    otHours = Math.max(0, Math.round((totalHours - settings.otThresholdHours) * 100) / 100);
    otStatus = otHours > 0 ? "PENDING" : "NONE";
  }

  return {
    id,
    employeeId: r.employeeId,
    date: r.date,
    scheduledClockIn,
    actualClockIn: r.actualClockIn,
    actualClockOut: r.actualClockOut || null,
    attendanceStatus: status,
    lateMinutes,
    totalHours,
    regularHours,
    otHours,
    otStatus,
    otApprovedBy: null,
    otApprovedAt: null,
    otNote: "",
    notes: r.note || "",
    editedBy: null,
    editedAt: null,
  };
}

function fmtTime(iso) {
  if (!iso) return "—";
  return new Date(iso).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}
function fmtDate(dateStr) {
  if (!dateStr) return "—";
  return new Date(`${dateStr}T00:00:00`).toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}
function fmtHM(hm) {
  if (!hm) return "—";
  const [h, m] = hm.split(":").map(Number);
  const d = new Date();
  d.setHours(h, m, 0, 0);
  return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}
function fmtDuration(mins) {
  if (!mins || mins <= 0) return "0 min";
  const h = Math.floor(mins / 60), m = mins % 60;
  if (h === 0) return `${m} min`;
  if (m === 0) return `${h}h`;
  return `${h}h ${m}m`;
}
function fmtHours(h) {
  if (h === null || h === undefined) return "—";
  return `${h.toFixed(2)}h`;
}
function todayStr() {
  return new Date().toLocaleDateString("en-CA");
}
function isWorkingDay(dateStr, settings) {
  const dow = new Date(`${dateStr}T00:00:00`).getDay();
  return settings.workingDays.includes(dow);
}

/* ------------------------------------------------------------------ */
/* Payroll calculation helpers                                        */
/* ------------------------------------------------------------------ */

function getHourlyRate(employee) {
  if (!employee) return 0;
  const rate = Number(employee.payRate) || 0;
  return employee.payRateType === "hourly" ? rate : rate / 8;
}
function getDailyRate(employee) {
  if (!employee) return 0;
  const rate = Number(employee.payRate) || 0;
  return employee.payRateType === "hourly" ? rate * 8 : rate;
}
function fmtMoney(n) {
  if (n === null || n === undefined || isNaN(n)) return "₱0.00";
  return `₱${n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

/* Period helpers: monthly / semi-monthly / custom -> { from, to } (YYYY-MM-DD, inclusive) */
function lastDayOfMonth(year, month1to12) {
  return new Date(year, month1to12, 0).getDate();
}
function computePeriodRange(periodType, ym, half, customFrom, customTo) {
  const [year, month] = ym.split("-").map(Number); // month is 1-12
  if (periodType === "monthly") {
    const last = lastDayOfMonth(year, month);
    return { from: `${ym}-01`, to: `${ym}-${String(last).padStart(2, "0")}` };
  }
  if (periodType === "semi-monthly") {
    const last = lastDayOfMonth(year, month);
    return half === "1"
      ? { from: `${ym}-01`, to: `${ym}-15` }
      : { from: `${ym}-16`, to: `${ym}-${String(last).padStart(2, "0")}` };
  }
  return { from: customFrom || todayStr(), to: customTo || todayStr() };
}

/* Compute a single employee's payroll figures for a set of attendance records
   already filtered to that employee + the chosen period. Records with no
   clock-out are excluded from hours/pay and flip needsReview on. */
function computeEmployeePayroll(employee, records, holidaySet, settings) {
  const hourlyRate = getHourlyRate(employee);
  const dailyRate = getDailyRate(employee);

  let hoursWorked = 0;
  let approvedOTHours = 0;
  let pendingOTHours = 0;
  let holidayDays = 0;
  let daysWorked = 0;
  let missingCount = 0;
  const missingDates = [];

  records.forEach((r) => {
    if (!r.actualClockOut) {
      missingCount += 1;
      missingDates.push(r.date);
      return; // excluded entirely from hours/pay
    }
    const hours = r.totalHours || 0;
    hoursWorked += hours;
    daysWorked += 1;
    if (r.otStatus === "APPROVED") approvedOTHours += r.otHours || 0;
    if (r.otStatus === "PENDING") pendingOTHours += r.otHours || 0;
    if (holidaySet.has(r.date)) holidayDays += 1;
  });

  hoursWorked = Math.round(hoursWorked * 100) / 100;
  approvedOTHours = Math.round(approvedOTHours * 100) / 100;

  // All worked hours (incl. OT hours) are paid straight-time here; the OT
  // premium below tops approved OT hours up to the OT multiplier.
  const regularPay = hoursWorked * hourlyRate;
  const otPremium = approvedOTHours * hourlyRate * (settings.otMultiplier - 1);
  const holidayPay = holidayDays * dailyRate;
  const grossPay = regularPay + otPremium + holidayPay;

  const pagibig = Math.min(grossPay, settings.pagibigCap) * settings.pagibigRate;
  const sss = Math.round(Math.min(grossPay, settings.sssCap) * settings.sssRate * 100) / 100;
  const philhealth = grossPay * settings.philhealthRate;
  const netPay = grossPay - pagibig - sss;

  return {
    hourlyRate, dailyRate, hoursWorked, approvedOTHours, pendingOTHours,
    holidayDays, daysWorked, missingCount, missingDates,
    regularPay, otPremium, holidayPay, grossPay,
    pagibig, sss, philhealth, netPay,
    needsReview: missingCount > 0,
  };
}

/* Live/derived display status for a record (handles NO_CLOCK_OUT for past days) */
function displayStatus(rec) {
  if (!rec.actualClockOut) {
    const recDate = rec.date;
    if (recDate < todayStr()) return "NO_CLOCK_OUT";
    return rec.attendanceStatus === "LATE" ? "LATE" : "CLOCKED_IN";
  }
  return rec.attendanceStatus;
}

const STATUS_STYLE = {
  ON_TIME: "bg-emerald-100 text-emerald-700 border-emerald-200",
  CLOCKED_IN: "bg-emerald-100 text-emerald-700 border-emerald-200",
  LATE: "bg-red-100 text-red-700 border-red-200",
  NO_CLOCK_OUT: "bg-amber-100 text-amber-800 border-amber-200",
  ABSENT: "bg-red-100 text-red-700 border-red-200",
  CLOCKED_OUT: "bg-gray-100 text-gray-600 border-gray-200",
};
const OT_STYLE = {
  NONE: "bg-gray-100 text-gray-500 border-gray-200",
  PENDING: "bg-orange-100 text-orange-700 border-orange-200",
  APPROVED: "bg-blue-100 text-blue-700 border-blue-200",
  REJECTED: "bg-red-100 text-red-700 border-red-200",
};

function StatusPill({ status, map = STATUS_STYLE, children }) {
  return (
    <span className={`inline-flex items-center px-2.5 py-1 rounded-full text-xs font-semibold border ${map[status] || "bg-gray-100 text-gray-600 border-gray-200"}`}>
      {children || status.replace(/_/g, " ")}
    </span>
  );
}

/* ------------------------------------------------------------------ */
/* Small UI primitives                                                 */
/* ------------------------------------------------------------------ */

function Card({ children, className = "" }) {
  return <div className={`bg-white rounded-xl border border-gray-200 shadow-sm ${className}`}>{children}</div>;
}

function Btn({ children, onClick, variant = "primary", size = "md", disabled, type = "button", className = "" }) {
  const base = "inline-flex items-center justify-center gap-1.5 font-semibold rounded-lg transition disabled:opacity-40 disabled:cursor-not-allowed";
  const sizes = { sm: "px-2.5 py-1.5 text-xs", md: "px-4 py-2 text-sm", lg: "px-6 py-3 text-base" };
  const variants = {
    primary: "bg-slate-900 text-white hover:bg-slate-800",
    accent: "bg-amber-500 text-white hover:bg-amber-600",
    success: "bg-emerald-600 text-white hover:bg-emerald-700",
    danger: "bg-red-600 text-white hover:bg-red-700",
    ghost: "bg-transparent text-slate-600 hover:bg-slate-100 border border-slate-200",
    subtle: "bg-slate-100 text-slate-700 hover:bg-slate-200",
  };
  return (
    <button type={type} disabled={disabled} onClick={onClick} className={`${base} ${sizes[size]} ${variants[variant]} ${className}`}>
      {children}
    </button>
  );
}

function Field({ label, children }) {
  return (
    <label className="block mb-3">
      <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">{label}</span>
      {children}
    </label>
  );
}
const inputCls = "w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-400 focus:border-amber-400";

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4" onClick={onClose}>
      <div className={`bg-white rounded-2xl shadow-2xl w-full ${wide ? "max-w-3xl" : "max-w-lg"} max-h-[90vh] overflow-y-auto`} onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100 sticky top-0 bg-white rounded-t-2xl">
          <h3 className="font-bold text-slate-900 text-lg">{title}</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-700"><X size={20} /></button>
        </div>
        <div className="p-6">{children}</div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Root App                                                            */
/* ------------------------------------------------------------------ */

export default function App() {
  const [loading, setLoading] = useState(true);
  const [employees, setEmployees] = useState([]);
  const [attendance, setAttendance] = useState([]);
  const [auditLog, setAuditLog] = useState([]);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [holidays, setHolidays] = useState([]);
  const [currentUser, setCurrentUser] = useState(null);
  const [toast, setToast] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        const results = await Promise.all(
          Object.values(KEYS).map((k) =>
            window.storage.get(k, true).catch(() => null)
          )
        );
        const [empRes, attRes, audRes, setRes, holRes] = results;

        let emp = empRes ? JSON.parse(empRes.value) : seedEmployees();
        let set = setRes ? { ...DEFAULT_SETTINGS, ...JSON.parse(setRes.value) } : DEFAULT_SETTINGS;
        let att = attRes ? JSON.parse(attRes.value) : seedAttendance();
        let aud = audRes ? JSON.parse(audRes.value) : seedAudit();
        let hol = holRes ? JSON.parse(holRes.value) : seedHolidays();

        if (!empRes) await window.storage.set(KEYS.employees, JSON.stringify(emp), true);
        if (!setRes) await window.storage.set(KEYS.settings, JSON.stringify(set), true);
        if (!attRes) await window.storage.set(KEYS.attendance, JSON.stringify(att), true);
        if (!audRes) await window.storage.set(KEYS.audit, JSON.stringify(aud), true);
        if (!holRes) await window.storage.set(KEYS.holidays, JSON.stringify(hol), true);

        setEmployees(emp);
        setSettings(set);
        setAttendance(att);
        setAuditLog(aud);
        setHolidays(hol);
      } catch (e) {
        console.error("Storage load failed", e);
        setEmployees(seedEmployees());
        setSettings(DEFAULT_SETTINGS);
        setAttendance(seedAttendance());
        setAuditLog(seedAudit());
        setHolidays(seedHolidays());
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const notify = useCallback((msg, tone = "success") => {
    setToast({ msg, tone });
    setTimeout(() => setToast(null), 2600);
  }, []);

  const persist = useCallback(async (key, data) => {
    try {
      await window.storage.set(key, JSON.stringify(data), true);
    } catch (e) {
      console.error("Save failed", e);
      notify("Couldn't save — changes may not persist.", "error");
    }
  }, [notify]);

  const addAudit = useCallback((actor, action, details) => {
    setAuditLog((prev) => {
      const next = [{ id: `a-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`, timestamp: new Date().toISOString(), actor, action, details }, ...prev];
      persist(KEYS.audit, next);
      return next;
    });
  }, [persist]);

  const saveEmployees = useCallback((updater, auditInfo) => {
    setEmployees((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      persist(KEYS.employees, next);
      return next;
    });
    if (auditInfo) addAudit(auditInfo.actor, auditInfo.action, auditInfo.details);
  }, [persist, addAudit]);

  const saveAttendance = useCallback((updater, auditInfo) => {
    setAttendance((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      persist(KEYS.attendance, next);
      return next;
    });
    if (auditInfo) addAudit(auditInfo.actor, auditInfo.action, auditInfo.details);
  }, [persist, addAudit]);

  const saveSettings = useCallback((next, auditInfo) => {
    setSettings(next);
    persist(KEYS.settings, next);
    if (auditInfo) addAudit(auditInfo.actor, auditInfo.action, auditInfo.details);
  }, [persist, addAudit]);

  const saveHolidays = useCallback((updater, auditInfo) => {
    setHolidays((prev) => {
      const next = typeof updater === "function" ? updater(prev) : updater;
      persist(KEYS.holidays, next);
      return next;
    });
    if (auditInfo) addAudit(auditInfo.actor, auditInfo.action, auditInfo.details);
  }, [persist, addAudit]);

  if (loading) {
    return (
      <div className="min-h-[600px] flex items-center justify-center bg-slate-50 font-sans">
        <div className="flex flex-col items-center gap-3 text-slate-400">
          <Clock className="animate-pulse" size={32} />
          <span className="text-sm font-medium">Loading attendance system…</span>
        </div>
      </div>
    );
  }

  return (
    <div className="ats-theme min-h-[700px] bg-slate-50 font-sans text-slate-900" style={{ fontFamily: "Inter, ui-sans-serif, system-ui, sans-serif" }}>
      <style>{`
        .ats-theme { --brand: #7FCDF4; --brand-dark: #4BAEDC; --accent: #FFC400; --bg: #FFFFFF; --bg-soft: #F7FBFD; --text: #222222; --text-secondary: #66727A; --border: #DDECF3; }

        /* Backgrounds */
        .ats-theme .bg-slate-50 { background-color: #F7FBFD !important; }
        .ats-theme .bg-slate-100 { background-color: #EAF5FC !important; }
        .ats-theme .bg-slate-200 { background-color: #DDECF3 !important; }
        .ats-theme .bg-slate-800 { background-color: #4BAEDC !important; }
        .ats-theme .bg-slate-900 { background-color: #2E86B8 !important; }
        .ats-theme .bg-slate-950 { background-color: #23698F !important; }
        .ats-theme .bg-amber-100 { background-color: #FFF3CC !important; }
        .ats-theme .bg-amber-500 { background-color: #FFC400 !important; color: #222222 !important; }
        .ats-theme .bg-amber-600 { background-color: #E6B000 !important; }
        .ats-theme .hover\\:bg-slate-50:hover { background-color: #F7FBFD !important; }
        .ats-theme .hover\\:bg-slate-100:hover { background-color: #EAF5FC !important; }
        .ats-theme .hover\\:bg-slate-200:hover { background-color: #DDECF3 !important; }
        .ats-theme .hover\\:bg-slate-800:hover { background-color: #3C9BCB !important; }
        .ats-theme .hover\\:bg-amber-600:hover { background-color: #E6B000 !important; }
        .ats-theme .hover\\:bg-amber-100:hover { background-color: #FFF3CC !important; }

        /* Text */
        .ats-theme .text-slate-900 { color: #222222 !important; }
        .ats-theme .text-slate-800 { color: #222222 !important; }
        .ats-theme .text-slate-700 { color: #33404A !important; }
        .ats-theme .text-slate-600 { color: #66727A !important; }
        .ats-theme .text-slate-500 { color: #66727A !important; }
        .ats-theme .text-slate-400 { color: #8B96A0 !important; }
        .ats-theme .text-slate-300 { color: #AEB9C2 !important; }
        .ats-theme .text-amber-400 { color: #FFC400 !important; }
        .ats-theme .text-amber-500 { color: #B38700 !important; }
        .ats-theme .text-amber-600 { color: #B38700 !important; }
        .ats-theme .hover\\:text-slate-700:hover { color: #33404A !important; }
        .ats-theme .hover\\:text-slate-800:hover { color: #222222 !important; }

        /* Borders */
        .ats-theme .border-slate-100 { border-color: #EAF2F7 !important; }
        .ats-theme .border-slate-200 { border-color: #DDECF3 !important; }
        .ats-theme .border-slate-300 { border-color: #C9DCE6 !important; }
        .ats-theme .border-slate-800 { border-color: #3C9BCB !important; }

        /* Login hero gradient panel */
        .ats-theme .from-slate-800 { --tw-gradient-from: #4BAEDC var(--tw-gradient-from-position); --tw-gradient-to: rgb(75 174 220 / 0) var(--tw-gradient-to-position); --tw-gradient-stops: var(--tw-gradient-from), var(--tw-gradient-to); }
        .ats-theme .to-slate-950 { --tw-gradient-to: #23698F var(--tw-gradient-to-position); }

        /* Inputs / focus states */
        .ats-theme .focus\\:ring-amber-400:focus { --tw-ring-color: #4BAEDC !important; }
        .ats-theme .focus\\:border-amber-400:focus { border-color: #4BAEDC !important; }
      `}</style>
      {toast && (
        <div className={`fixed top-4 right-4 z-[100] px-4 py-3 rounded-lg shadow-lg text-sm font-semibold text-white ${toast.tone === "error" ? "bg-red-600" : "bg-slate-900"}`}>
          {toast.msg}
        </div>
      )}
      {!currentUser ? (
        <LoginScreen employees={employees} onLogin={setCurrentUser} />
      ) : currentUser.role === "admin" ? (
        <AdminApp
          currentUser={currentUser}
          onLogout={() => setCurrentUser(null)}
          employees={employees}
          attendance={attendance}
          auditLog={auditLog}
          settings={settings}
          holidays={holidays}
          saveEmployees={saveEmployees}
          saveAttendance={saveAttendance}
          saveSettings={saveSettings}
          saveHolidays={saveHolidays}
          notify={notify}
        />
      ) : (
        <EmployeeApp
          currentUser={currentUser}
          onLogout={() => setCurrentUser(null)}
          employees={employees}
          attendance={attendance}
          settings={settings}
          saveAttendance={saveAttendance}
          notify={notify}
        />
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Login                                                               */
/* ------------------------------------------------------------------ */

function LoginScreen({ employees, onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [now, setNow] = useState(new Date());

  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  const handleSubmit = (e) => {
    e.preventDefault();
    const user = employees.find((e2) => e2.username.toLowerCase() === username.trim().toLowerCase());
    if (!user || user.password !== password) {
      setError("Incorrect username or password.");
      return;
    }
    if (user.status !== "active") {
      setError("This account has been deactivated. Contact your Admin.");
      return;
    }
    setError("");
    onLogin(user);
  };

  const angle = ((now.getHours() % 12) + now.getMinutes() / 60) * 30;
  const minAngle = now.getMinutes() * 6;

  return (
    <div className="min-h-[700px] flex items-center justify-center bg-slate-900 p-6">
      <div className="w-full max-w-4xl grid md:grid-cols-2 rounded-2xl overflow-hidden shadow-2xl">
        <div className="hidden md:flex flex-col justify-between bg-gradient-to-br from-slate-800 to-slate-950 p-10 text-white">
          <div>
            <div className="flex items-center gap-2 text-amber-400 font-bold text-sm tracking-widest uppercase mb-8">
              <Clock size={18} /> TimeKeep
            </div>
            <h1 className="text-3xl font-bold leading-tight mb-3">Attendance &amp; Time Tracking, done right.</h1>
            <p className="text-slate-400 text-sm leading-relaxed">Clock in, clock out, and let the system handle regular hours, overtime, and lateness — accurately, every time.</p>
          </div>
          <div className="flex items-center gap-6">
            <svg width="88" height="88" viewBox="0 0 100 100">
              <circle cx="50" cy="50" r="46" fill="none" stroke="#334155" strokeWidth="3" />
              {[...Array(12)].map((_, i) => (
                <line key={i} x1="50" y1="7" x2="50" y2="13" stroke="#64748b" strokeWidth="2"
                  transform={`rotate(${i * 30} 50 50)`} />
              ))}
              <line x1="50" y1="50" x2={50 + 22 * Math.sin((angle * Math.PI) / 180)} y2={50 - 22 * Math.cos((angle * Math.PI) / 180)} stroke="#f59e0b" strokeWidth="3" strokeLinecap="round" />
              <line x1="50" y1="50" x2={50 + 32 * Math.sin((minAngle * Math.PI) / 180)} y2={50 - 32 * Math.cos((minAngle * Math.PI) / 180)} stroke="#e2e8f0" strokeWidth="2" strokeLinecap="round" />
              <circle cx="50" cy="50" r="2.5" fill="#f59e0b" />
            </svg>
            <div className="font-mono">
              <div className="text-2xl font-bold tabular-nums">{now.toLocaleTimeString()}</div>
              <div className="text-slate-400 text-xs">{now.toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" })}</div>
            </div>
          </div>
        </div>
        <div className="bg-white p-10 flex flex-col justify-center">
          <h2 className="text-xl font-bold text-slate-900 mb-1">Sign in</h2>
          <p className="text-sm text-slate-500 mb-6">Use your employee or admin account.</p>
          <form onSubmit={handleSubmit}>
            <Field label="Username">
              <input className={inputCls} value={username} onChange={(e) => setUsername(e.target.value)} autoFocus />
            </Field>
            <Field label="Password">
              <input type="password" className={inputCls} value={password} onChange={(e) => setPassword(e.target.value)} />
            </Field>
            {error && <div className="text-sm text-red-600 font-medium mb-3 flex items-center gap-1.5"><AlertTriangle size={14} />{error}</div>}
            <Btn type="submit" variant="accent" size="lg" className="w-full mt-2">Sign in <LogIn size={16} /></Btn>
          </form>
          <div className="mt-6 pt-5 border-t border-slate-100 text-xs text-slate-400 space-y-1">
            <div className="font-semibold text-slate-500">Demo accounts</div>
            <div>Admin — <span className="font-mono">admin / admin123</span></div>
            <div>Employee — <span className="font-mono">marco / employee123</span></div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Employee App                                                        */
/* ------------------------------------------------------------------ */

function EmployeeApp({ currentUser, onLogout, employees, attendance, settings, saveAttendance, notify }) {
  const [now, setNow] = useState(new Date());
  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  const today = todayStr();
  const myRecords = attendance.filter((r) => r.employeeId === currentUser.employeeId).sort((a, b) => b.date.localeCompare(a.date));
  const openRecord = myRecords.find((r) => r.date === today && !r.actualClockOut);
  const todayRecord = myRecords.find((r) => r.date === today);

  const { scheduledClockIn, grace } = getEmployeeSchedule(currentUser, settings);

  const handleClockIn = () => {
    const alreadyOpen = attendance.find((r) => r.employeeId === currentUser.employeeId && !r.actualClockOut);
    if (alreadyOpen) {
      notify("You're already clocked in. Clock out first.", "error");
      return;
    }
    const nowIso = new Date().toISOString();
    const scheduled = new Date(`${today}T${scheduledClockIn}:00`);
    const lateRaw = Math.round((new Date(nowIso) - scheduled) / 60000);
    const lateMinutes = Math.max(0, lateRaw - grace);
    const rec = {
      id: `att-${Date.now()}`,
      employeeId: currentUser.employeeId,
      date: today,
      scheduledClockIn,
      actualClockIn: nowIso,
      actualClockOut: null,
      attendanceStatus: lateMinutes > 0 ? "LATE" : "ON_TIME",
      lateMinutes,
      totalHours: null,
      regularHours: null,
      otHours: 0,
      otStatus: "NONE",
      otApprovedBy: null,
      otApprovedAt: null,
      otNote: "",
      notes: "",
      editedBy: null,
      editedAt: null,
    };
    saveAttendance((prev) => [rec, ...prev], { actor: currentUser.name, action: "CLOCK_IN", details: `Clocked in at ${fmtTime(nowIso)}` });
    notify(lateMinutes > 0 ? `Clocked in — ${fmtDuration(lateMinutes)} late.` : "Clocked in on time!");
  };

  const handleClockOut = () => {
    const open = attendance.find((r) => r.employeeId === currentUser.employeeId && !r.actualClockOut);
    if (!open) {
      notify("You need to clock in first.", "error");
      return;
    }
    const nowIso = new Date().toISOString();
    const totalHours = Math.round(((new Date(nowIso) - new Date(open.actualClockIn)) / 3600000) * 100) / 100;
    const regularHours = Math.min(totalHours, settings.standardHours);
    const otHours = Math.max(0, Math.round((totalHours - settings.otThresholdHours) * 100) / 100);
    saveAttendance(
      (prev) => prev.map((r) => r.id === open.id ? { ...r, actualClockOut: nowIso, totalHours, regularHours, otHours, otStatus: otHours > 0 ? "PENDING" : "NONE" } : r),
      { actor: currentUser.name, action: "CLOCK_OUT", details: `Clocked out at ${fmtTime(nowIso)} — ${totalHours}h worked` }
    );
    notify(otHours > 0 ? `Clocked out. ${otHours}h of potential OT recorded.` : "Clocked out — have a good one!");
  };

  const hoursSoFar = openRecord ? Math.max(0, (now - new Date(openRecord.actualClockIn)) / 3600000) : (todayRecord?.totalHours || 0);
  const ringPct = Math.min(1, hoursSoFar / settings.standardHours);
  const circumference = 2 * Math.PI * 54;

  return (
    <div className="max-w-lg mx-auto p-6">
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center gap-2 text-slate-900 font-bold">
          <Clock className="text-amber-500" size={20} /> TimeKeep
        </div>
        <button onClick={onLogout} className="text-sm text-slate-500 hover:text-slate-800 flex items-center gap-1"><LogOutIcon size={15} /> Sign out</button>
      </div>

      <Card className="p-6 mb-5">
        <div className="text-sm text-slate-400 font-medium">{now.toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" })}</div>
        <h1 className="text-xl font-bold text-slate-900 mb-4">Hi, {currentUser.name.split(" ")[0]}</h1>

        <div className="flex flex-col items-center py-4">
          <div className="relative w-40 h-40 mb-4">
            <svg width="160" height="160" viewBox="0 0 120 120" className="-rotate-90">
              <circle cx="60" cy="60" r="54" fill="none" stroke="#f1f5f9" strokeWidth="10" />
              <circle cx="60" cy="60" r="54" fill="none" stroke={openRecord ? "#10b981" : "#94a3b8"} strokeWidth="10" strokeLinecap="round"
                strokeDasharray={circumference} strokeDashoffset={circumference * (1 - ringPct)} style={{ transition: "stroke-dashoffset 1s linear" }} />
            </svg>
            <div className="absolute inset-0 flex flex-col items-center justify-center">
              <span className="text-2xl font-bold font-mono tabular-nums text-slate-900">{hoursSoFar.toFixed(2)}h</span>
              <span className="text-xs text-slate-400">of {settings.standardHours}h shift</span>
            </div>
          </div>

          <StatusPill status={openRecord ? "CLOCKED_IN" : todayRecord ? "CLOCKED_OUT" : "CLOCKED_OUT"}>
            {openRecord ? "Clocked In" : todayRecord ? "Clocked Out" : "Not Clocked In"}
          </StatusPill>

          <div className="grid grid-cols-2 gap-3 w-full mt-6">
            <Btn onClick={handleClockIn} disabled={!!openRecord} variant="success" size="lg"><LogIn size={16} /> Clock In</Btn>
            <Btn onClick={handleClockOut} disabled={!openRecord} variant="danger" size="lg"><LogOut size={16} /> Clock Out</Btn>
          </div>
        </div>
      </Card>

      <Card className="p-5 mb-5">
        <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-3">Today's schedule</div>
        <div className="flex justify-between text-sm mb-2"><span className="text-slate-500">Expected clock-in</span><span className="font-semibold font-mono">{fmtHM(scheduledClockIn)}</span></div>
        <div className="flex justify-between text-sm"><span className="text-slate-500">Standard hours</span><span className="font-semibold">{settings.standardHours}h</span></div>
      </Card>

      {todayRecord && (
        <Card className="p-5 mb-5">
          <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-3">Today's attendance</div>
          <div className="space-y-2 text-sm">
            <div className="flex justify-between"><span className="text-slate-500">Clock-in</span><span className="font-mono font-semibold">{fmtTime(todayRecord.actualClockIn)}</span></div>
            <div className="flex justify-between"><span className="text-slate-500">Clock-out</span><span className="font-mono font-semibold">{fmtTime(todayRecord.actualClockOut)}</span></div>
            <div className="flex justify-between items-center"><span className="text-slate-500">Status</span><StatusPill status={displayStatus(todayRecord)} /></div>
            {todayRecord.lateMinutes > 0 && (
              <div className="flex justify-between"><span className="text-slate-500">Late by</span><span className="font-semibold text-red-600">{fmtDuration(todayRecord.lateMinutes)}</span></div>
            )}
            <div className="flex justify-between"><span className="text-slate-500">Total hours</span><span className="font-semibold">{fmtHours(todayRecord.totalHours)}</span></div>
            {todayRecord.otHours > 0 && (
              <div className="flex justify-between items-center"><span className="text-slate-500">Overtime</span>
                <span className="flex items-center gap-2 font-semibold">{fmtHours(todayRecord.otHours)} <StatusPill status={todayRecord.otStatus} map={OT_STYLE} /></span>
              </div>
            )}
          </div>
        </Card>
      )}

      <Card className="p-5">
        <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-3 flex items-center gap-1.5"><History size={13}/> Recent history</div>
        <div className="space-y-2">
          {myRecords.slice(0, 7).map((r) => (
            <div key={r.id} className="flex items-center justify-between text-sm py-1.5 border-b border-slate-50 last:border-0">
              <span className="text-slate-600">{fmtDate(r.date)}</span>
              <span className="font-mono text-xs text-slate-500">{fmtTime(r.actualClockIn)}–{fmtTime(r.actualClockOut)}</span>
              <StatusPill status={displayStatus(r)} />
            </div>
          ))}
          {myRecords.length === 0 && <div className="text-sm text-slate-400 py-3 text-center">No attendance records yet.</div>}
        </div>
      </Card>

      <p className="text-center text-xs text-slate-400 mt-6">Records are locked once submitted. Contact your Admin to correct an entry.</p>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Admin App shell                                                     */
/* ------------------------------------------------------------------ */

function AdminApp(props) {
  const { currentUser, onLogout } = props;
  const [page, setPage] = useState("dashboard");

  const nav = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "attendance", label: "Attendance Records", icon: ClipboardList },
    { id: "employees", label: "Employees", icon: Users },
    { id: "payroll", label: "Payroll", icon: Wallet },
    { id: "reports", label: "Reports", icon: FileText },
    { id: "settings", label: "Settings", icon: SettingsIcon },
    { id: "audit", label: "Audit Log", icon: ShieldCheck },
  ];

  return (
    <div className="flex min-h-[700px]">
      <aside className="w-60 bg-slate-900 text-white flex flex-col shrink-0">
        <div className="px-5 py-5 flex items-center gap-2 font-bold text-lg border-b border-slate-800">
          <Clock className="text-amber-400" size={20} /> TimeKeep
        </div>
        <nav className="flex-1 py-4 space-y-1 px-3">
          {nav.map((n) => {
            const Icon = n.icon;
            const active = page === n.id;
            return (
              <button key={n.id} onClick={() => setPage(n.id)}
                className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-lg text-sm font-medium transition ${active ? "bg-amber-500 text-white" : "text-slate-400 hover:bg-slate-800 hover:text-white"}`}>
                <Icon size={16} /> {n.label}
              </button>
            );
          })}
        </nav>
        <div className="p-4 border-t border-slate-800">
          <div className="text-xs text-slate-400 mb-2">Signed in as</div>
          <div className="text-sm font-semibold mb-3">{currentUser.name}</div>
          <button onClick={onLogout} className="w-full flex items-center gap-2 text-xs text-slate-400 hover:text-white"><LogOutIcon size={14}/> Sign out</button>
        </div>
      </aside>
      <main className="flex-1 bg-slate-50 overflow-y-auto p-6">
        {page === "dashboard" && <AdminDashboard {...props} />}
        {page === "attendance" && <AttendancePage {...props} />}
        {page === "employees" && <EmployeesPage {...props} />}
        {page === "payroll" && <PayrollPage {...props} />}
        {page === "reports" && <ReportsPage {...props} />}
        {page === "settings" && <SettingsPage {...props} />}
        {page === "audit" && <AuditPage {...props} />}
      </main>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Dashboard                                                           */
/* ------------------------------------------------------------------ */

function StatCard({ icon: Icon, label, value, tone = "slate" }) {
  const tones = {
    slate: "bg-slate-100 text-slate-600",
    emerald: "bg-emerald-100 text-emerald-600",
    amber: "bg-amber-100 text-amber-600",
    red: "bg-red-100 text-red-600",
    blue: "bg-blue-100 text-blue-600",
  };
  return (
    <Card className="p-4 flex items-center gap-3">
      <div className={`w-10 h-10 rounded-lg flex items-center justify-center shrink-0 ${tones[tone]}`}><Icon size={18} /></div>
      <div>
        <div className="text-xl font-bold text-slate-900 leading-tight">{value}</div>
        <div className="text-xs text-slate-500">{label}</div>
      </div>
    </Card>
  );
}

function AdminDashboard({ employees, attendance, settings }) {
  const today = todayStr();
  const activeEmployees = employees.filter((e) => e.role === "employee" && e.status === "active");
  const todayRecords = attendance.filter((r) => r.date === today);
  const clockedIn = todayRecords.filter((r) => !r.actualClockOut).length;
  const clockedOut = todayRecords.filter((r) => r.actualClockOut).length;
  const pendingOT = attendance.filter((r) => r.otStatus === "PENDING");
  const totalOTHoursMonth = attendance.filter((r) => r.date.slice(0, 7) === today.slice(0, 7) && r.otStatus === "APPROVED").reduce((s, r) => s + (r.otHours || 0), 0);

  const lateToday = todayRecords.filter((r) => r.attendanceStatus === "LATE").map((r) => ({ ...r, employee: employees.find((e) => e.employeeId === r.employeeId) }));

  const clockedInIds = new Set(todayRecords.filter((r) => !r.actualClockOut).map((r) => r.employeeId));
  const recordedIds = new Set(todayRecords.map((r) => r.employeeId));
  const absentToday = isWorkingDay(today, settings) && new Date().getHours() >= 12
    ? activeEmployees.filter((e) => !recordedIds.has(e.employeeId))
    : [];

  const recent = [...attendance].sort((a, b) => (b.actualClockIn || "").localeCompare(a.actualClockIn || "")).slice(0, 6);

  return (
    <div>
      <h1 className="text-2xl font-bold text-slate-900 mb-1">Dashboard</h1>
      <p className="text-slate-500 text-sm mb-6">{new Date().toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric", year: "numeric" })}</p>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <StatCard icon={Users} label="Total employees" value={activeEmployees.length} tone="slate" />
        <StatCard icon={LogIn} label="Clocked in now" value={clockedIn} tone="emerald" />
        <StatCard icon={LogOut} label="Clocked out today" value={clockedOut} tone="slate" />
        <StatCard icon={Calendar} label="Today's attendance" value={`${recordedIds.size}/${activeEmployees.length}`} tone="blue" />
        <StatCard icon={AlertTriangle} label="Pending OT approvals" value={pendingOT.length} tone="amber" />
        <StatCard icon={Clock} label="Approved OT this month" value={`${totalOTHoursMonth.toFixed(1)}h`} tone="blue" />
        <StatCard icon={XCircle} label="Late today" value={lateToday.length} tone="red" />
        <StatCard icon={Ban} label="Absent today" value={absentToday.length} tone="red" />
      </div>

      <div className="grid md:grid-cols-2 gap-5">
        <Card className="p-5">
          <div className="font-bold text-slate-900 mb-3 flex items-center gap-1.5"><AlertTriangle size={16} className="text-red-500"/> Late employees today</div>
          {lateToday.length === 0 && <div className="text-sm text-slate-400 py-4 text-center">No late arrivals today.</div>}
          <div className="space-y-2">
            {lateToday.map((r) => (
              <div key={r.id} className="flex items-center justify-between text-sm py-2 border-b border-slate-50 last:border-0">
                <div>
                  <div className="font-semibold text-slate-800">{r.employee?.name}</div>
                  <div className="text-xs text-slate-400">{r.employeeId} · {r.employee?.department}</div>
                </div>
                <div className="text-right">
                  <div className="text-xs text-slate-500 font-mono">{fmtHM(r.scheduledClockIn)} → {fmtTime(r.actualClockIn)}</div>
                  <div className="text-xs font-semibold text-red-600">{fmtDuration(r.lateMinutes)} late</div>
                </div>
              </div>
            ))}
          </div>
        </Card>

        <Card className="p-5">
          <div className="font-bold text-slate-900 mb-3 flex items-center gap-1.5"><History size={16} className="text-slate-400"/> Recent activity</div>
          <div className="space-y-2">
            {recent.map((r) => {
              const emp = employees.find((e) => e.employeeId === r.employeeId);
              return (
                <div key={r.id} className="flex items-center justify-between text-sm py-2 border-b border-slate-50 last:border-0">
                  <div>
                    <div className="font-semibold text-slate-800">{emp?.name || r.employeeId}</div>
                    <div className="text-xs text-slate-400">{fmtDate(r.date)}</div>
                  </div>
                  <StatusPill status={displayStatus(r)} />
                </div>
              );
            })}
          </div>
        </Card>
      </div>

      {absentToday.length > 0 && (
        <Card className="p-5 mt-5">
          <div className="font-bold text-slate-900 mb-3 flex items-center gap-1.5"><Ban size={16} className="text-red-500"/> Absent today (no clock-in recorded)</div>
          <div className="flex flex-wrap gap-2">
            {absentToday.map((e) => (
              <span key={e.employeeId} className="px-3 py-1.5 rounded-lg bg-red-50 text-red-700 text-xs font-semibold border border-red-100">{e.name} · {e.department}</span>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Attendance Records page                                             */
/* ------------------------------------------------------------------ */

const DATE_PRESETS = ["All", "Today", "This Week", "This Month", "Custom"];
const STATUS_FILTERS = ["All", "ON_TIME", "LATE", "NO_CLOCK_OUT", "CLOCKED_IN"];
const OT_FILTERS = ["All", "NONE", "PENDING", "APPROVED", "REJECTED"];

function inDateRange(dateStr, preset, customFrom, customTo) {
  const d = new Date(`${dateStr}T00:00:00`);
  const now = new Date();
  if (preset === "Today") return dateStr === todayStr();
  if (preset === "This Week") {
    const start = new Date(now); start.setDate(now.getDate() - now.getDay());
    start.setHours(0,0,0,0);
    return d >= start;
  }
  if (preset === "This Month") return dateStr.slice(0, 7) === todayStr().slice(0, 7);
  if (preset === "Custom") {
    if (customFrom && d < new Date(customFrom)) return false;
    if (customTo && d > new Date(customTo)) return false;
    return true;
  }
  return true;
}

function AttendancePage({ employees, attendance, settings, saveAttendance, currentUser, notify }) {
  const [search, setSearch] = useState("");
  const [deptFilter, setDeptFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [otFilter, setOtFilter] = useState("All");
  const [datePreset, setDatePreset] = useState("All");
  const [customFrom, setCustomFrom] = useState("");
  const [customTo, setCustomTo] = useState("");
  const [sortDesc, setSortDesc] = useState(true);
  const [editRecord, setEditRecord] = useState(null);
  const [otRecord, setOtRecord] = useState(null);

  const departments = useMemo(() => ["All", ...new Set(employees.map((e) => e.department))], [employees]);

  const rows = useMemo(() => {
    return attendance
      .map((r) => ({ ...r, employee: employees.find((e) => e.employeeId === r.employeeId) }))
      .filter((r) => {
        if (search && !(r.employee?.name.toLowerCase().includes(search.toLowerCase()) || r.employeeId.toLowerCase().includes(search.toLowerCase()))) return false;
        if (deptFilter !== "All" && r.employee?.department !== deptFilter) return false;
        if (statusFilter !== "All" && displayStatus(r) !== statusFilter) return false;
        if (otFilter !== "All" && r.otStatus !== otFilter) return false;
        if (!inDateRange(r.date, datePreset, customFrom, customTo)) return false;
        return true;
      })
      .sort((a, b) => sortDesc ? b.date.localeCompare(a.date) : a.date.localeCompare(b.date));
  }, [attendance, employees, search, deptFilter, statusFilter, otFilter, datePreset, customFrom, customTo, sortDesc]);

  const handleApproveOT = (record, note) => {
    saveAttendance(
      (prev) => prev.map((r) => r.id === record.id ? { ...r, otStatus: "APPROVED", otApprovedBy: currentUser.name, otApprovedAt: new Date().toISOString(), otNote: note } : r),
      { actor: currentUser.name, action: "OT_APPROVED", details: `Approved ${record.otHours}h OT for ${record.employeeId} on ${record.date}${note ? ` — "${note}"` : ""}` }
    );
    notify("Overtime approved.");
    setOtRecord(null);
  };
  const handleRejectOT = (record, note) => {
    saveAttendance(
      (prev) => prev.map((r) => r.id === record.id ? { ...r, otStatus: "REJECTED", otApprovedBy: currentUser.name, otApprovedAt: new Date().toISOString(), otNote: note } : r),
      { actor: currentUser.name, action: "OT_REJECTED", details: `Rejected ${record.otHours}h OT for ${record.employeeId} on ${record.date}${note ? ` — "${note}"` : ""}` }
    );
    notify("Overtime rejected.");
    setOtRecord(null);
  };
  const handleSaveEdit = (record, changes, reason) => {
    saveAttendance(
      (prev) => prev.map((r) => {
        if (r.id !== record.id) return r;
        const merged = { ...r, ...changes, editedBy: currentUser.name, editedAt: new Date().toISOString() };
        if (merged.actualClockIn && merged.actualClockOut) {
          const totalHours = Math.round(((new Date(merged.actualClockOut) - new Date(merged.actualClockIn)) / 3600000) * 100) / 100;
          merged.totalHours = totalHours;
          merged.regularHours = Math.min(totalHours, settings.standardHours);
          merged.otHours = Math.max(0, Math.round((totalHours - settings.otThresholdHours) * 100) / 100);
          if (merged.otHours === 0) merged.otStatus = "NONE";
          else if (r.otStatus === "NONE") merged.otStatus = "PENDING";
        }
        return merged;
      }),
      { actor: currentUser.name, action: "RECORD_EDITED", details: `Edited attendance for ${record.employeeId} on ${record.date}${reason ? ` — reason: "${reason}"` : ""}` }
    );
    notify("Record updated.");
    setEditRecord(null);
  };

  return (
    <div>
      <h1 className="text-2xl font-bold text-slate-900 mb-4">Attendance Records</h1>

      <Card className="p-4 mb-4">
        <div className="flex flex-wrap gap-3 items-end">
          <div className="flex-1 min-w-[180px]">
            <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">Search</span>
            <div className="relative">
              <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input className={inputCls + " pl-8"} placeholder="Employee name or ID" value={search} onChange={(e) => setSearch(e.target.value)} />
            </div>
          </div>
          <div>
            <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">Department</span>
            <select className={inputCls} value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)}>
              {departments.map((d) => <option key={d}>{d}</option>)}
            </select>
          </div>
          <div>
            <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">Status</span>
            <select className={inputCls} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
              {STATUS_FILTERS.map((s) => <option key={s} value={s}>{s.replace(/_/g, " ")}</option>)}
            </select>
          </div>
          <div>
            <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">OT status</span>
            <select className={inputCls} value={otFilter} onChange={(e) => setOtFilter(e.target.value)}>
              {OT_FILTERS.map((s) => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div>
            <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">Date range</span>
            <select className={inputCls} value={datePreset} onChange={(e) => setDatePreset(e.target.value)}>
              {DATE_PRESETS.map((s) => <option key={s}>{s}</option>)}
            </select>
          </div>
          {datePreset === "Custom" && (
            <>
              <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">From</span><input type="date" className={inputCls} value={customFrom} onChange={(e) => setCustomFrom(e.target.value)} /></div>
              <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">To</span><input type="date" className={inputCls} value={customTo} onChange={(e) => setCustomTo(e.target.value)} /></div>
            </>
          )}
          <Btn variant="ghost" onClick={() => setSortDesc((s) => !s)}><ChevronDown size={14} className={sortDesc ? "" : "rotate-180"} /> Date</Btn>
        </div>
      </Card>

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
              <th className="px-4 py-3">Employee</th>
              <th className="px-4 py-3">Date</th>
              <th className="px-4 py-3">Scheduled</th>
              <th className="px-4 py-3">In / Out</th>
              <th className="px-4 py-3">Status</th>
              <th className="px-4 py-3">Late</th>
              <th className="px-4 py-3">Hours</th>
              <th className="px-4 py-3">OT</th>
              <th className="px-4 py-3">Actions</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id} className="border-b border-slate-50 hover:bg-slate-50/60">
                <td className="px-4 py-3">
                  <div className="font-semibold text-slate-800">{r.employee?.name || r.employeeId}</div>
                  <div className="text-xs text-slate-400">{r.employeeId} · {r.employee?.department}</div>
                </td>
                <td className="px-4 py-3 whitespace-nowrap">{fmtDate(r.date)}</td>
                <td className="px-4 py-3 font-mono text-xs">{fmtHM(r.scheduledClockIn)}</td>
                <td className="px-4 py-3 font-mono text-xs whitespace-nowrap">{fmtTime(r.actualClockIn)} – {fmtTime(r.actualClockOut)}</td>
                <td className="px-4 py-3"><StatusPill status={displayStatus(r)} /></td>
                <td className="px-4 py-3">{r.lateMinutes > 0 ? <span className="text-red-600 font-semibold text-xs">{fmtDuration(r.lateMinutes)}</span> : <span className="text-slate-300 text-xs">—</span>}</td>
                <td className="px-4 py-3 text-xs">
                  <div>{fmtHours(r.totalHours)}</div>
                  {r.regularHours != null && <div className="text-slate-400">{r.regularHours}h reg</div>}
                </td>
                <td className="px-4 py-3">
                  {r.otHours > 0 ? <div className="flex flex-col gap-1"><span className="text-xs font-semibold">{r.otHours}h</span><StatusPill status={r.otStatus} map={OT_STYLE} /></div> : <span className="text-slate-300 text-xs">—</span>}
                </td>
                <td className="px-4 py-3">
                  <div className="flex gap-1.5">
                    <button title="Edit record" onClick={() => setEditRecord(r)} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><Edit3 size={14} /></button>
                    {r.otStatus === "PENDING" && (
                      <button title="Review OT" onClick={() => setOtRecord(r)} className="p-1.5 rounded hover:bg-amber-100 text-amber-600"><AlertTriangle size={14} /></button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
            {rows.length === 0 && (
              <tr><td colSpan={9} className="text-center text-slate-400 py-10">No records match these filters.</td></tr>
            )}
          </tbody>
        </table>
      </Card>

      {editRecord && <RecordEditModal record={editRecord} onClose={() => setEditRecord(null)} onSave={handleSaveEdit} />}
      {otRecord && <OTReviewModal record={otRecord} onClose={() => setOtRecord(null)} onApprove={handleApproveOT} onReject={handleRejectOT} />}
    </div>
  );
}

function RecordEditModal({ record, onClose, onSave }) {
  const [inTime, setInTime] = useState(record.actualClockIn ? record.actualClockIn.slice(11, 16) : "");
  const [outTime, setOutTime] = useState(record.actualClockOut ? record.actualClockOut.slice(11, 16) : "");
  const [notes, setNotes] = useState(record.notes || "");
  const [reason, setReason] = useState("");

  const submit = () => {
    const changes = {
      actualClockIn: inTime ? `${record.date}T${inTime}:00` : null,
      actualClockOut: outTime ? `${record.date}T${outTime}:00` : null,
      notes,
    };
    onSave(record, changes, reason);
  };

  return (
    <Modal title={`Edit record — ${record.employee?.name || record.employeeId}`} onClose={onClose}>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Clock-in time"><input type="time" className={inputCls} value={inTime} onChange={(e) => setInTime(e.target.value)} /></Field>
        <Field label="Clock-out time"><input type="time" className={inputCls} value={outTime} onChange={(e) => setOutTime(e.target.value)} /></Field>
      </div>
      <Field label="Notes"><textarea className={inputCls} rows={2} value={notes} onChange={(e) => setNotes(e.target.value)} /></Field>
      <Field label="Reason for correction (logged to audit trail)"><input className={inputCls} value={reason} onChange={(e) => setReason(e.target.value)} placeholder="e.g. Forgot to clock out, corrected per employee request" /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="primary" onClick={submit}>Save changes</Btn>
      </div>
    </Modal>
  );
}

function OTReviewModal({ record, onClose, onApprove, onReject }) {
  const [note, setNote] = useState("");
  return (
    <Modal title="Review overtime" onClose={onClose}>
      <div className="mb-4 bg-slate-50 rounded-lg p-4 text-sm space-y-1.5">
        <div className="flex justify-between"><span className="text-slate-500">Employee</span><span className="font-semibold">{record.employee?.name}</span></div>
        <div className="flex justify-between"><span className="text-slate-500">Date</span><span className="font-semibold">{fmtDate(record.date)}</span></div>
        <div className="flex justify-between"><span className="text-slate-500">Worked</span><span className="font-semibold">{fmtTime(record.actualClockIn)} – {fmtTime(record.actualClockOut)} ({record.totalHours}h)</span></div>
        <div className="flex justify-between"><span className="text-slate-500">Overtime</span><span className="font-bold text-amber-600">{record.otHours}h</span></div>
      </div>
      <Field label="Note (optional)"><textarea className={inputCls} rows={2} value={note} onChange={(e) => setNote(e.target.value)} placeholder="Reason for approval/rejection" /></Field>
      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="danger" onClick={() => onReject(record, note)}><XCircle size={15}/> Reject</Btn>
        <Btn variant="success" onClick={() => onApprove(record, note)}><CheckCircle2 size={15}/> Approve OT</Btn>
      </div>
    </Modal>
  );
}

/* ------------------------------------------------------------------ */
/* Employees page                                                      */
/* ------------------------------------------------------------------ */

function EmployeesPage({ employees, attendance, saveEmployees, saveAttendance, currentUser, notify }) {
  const [search, setSearch] = useState("");
  const [modalEmp, setModalEmp] = useState(null);
  const [creating, setCreating] = useState(false);
  const [historyEmp, setHistoryEmp] = useState(null);
  const [rateHistoryEmp, setRateHistoryEmp] = useState(null);
  const [deleteEmp, setDeleteEmp] = useState(null);

  const rows = employees.filter((e) => (e.name.toLowerCase().includes(search.toLowerCase()) || e.employeeId.toLowerCase().includes(search.toLowerCase())));

  const handleSave = (data, isNew) => {
    if (isNew) {
      const rec = { id: `u-${Date.now()}`, status: "active", rateHistory: [], ...data };
      saveEmployees((prev) => [...prev, rec], { actor: currentUser.name, action: "EMPLOYEE_ADDED", details: `Added employee ${rec.name} (${rec.employeeId})` });
      notify("Employee added.");
    } else {
      saveEmployees((prev) => prev.map((e) => {
        if (e.id !== data.id) return e;
        let rateHistory = e.rateHistory || [];
        const rateChanged = data.payRate !== undefined && (Number(e.payRate) !== Number(data.payRate) || e.payRateType !== data.payRateType);
        if (rateChanged) {
          rateHistory = [
            { rate: e.payRate ?? null, payRateType: e.payRateType ?? null, effectiveUntil: todayStr(), changedBy: currentUser.name, changedAt: new Date().toISOString() },
            ...rateHistory,
          ];
        }
        return { ...e, ...data, rateHistory };
      }), { actor: currentUser.name, action: "EMPLOYEE_EDITED", details: `Edited employee ${data.name} (${data.employeeId})` });
      notify("Employee updated.");
    }
    setModalEmp(null);
    setCreating(false);
  };

  const toggleStatus = (emp) => {
    if (emp.id === currentUser.id) {
      notify("You can't deactivate your own account.", "error");
      return;
    }
    const newStatus = emp.status === "active" ? "inactive" : "active";
    saveEmployees((prev) => prev.map((e) => e.id === emp.id ? { ...e, status: newStatus } : e), { actor: currentUser.name, action: "EMPLOYEE_STATUS_CHANGED", details: `${emp.name} set to ${newStatus}` });
    notify(`${emp.name} ${newStatus === "active" ? "activated" : "deactivated"}.`);
  };

  const handleDelete = (emp, alsoDeleteRecords) => {
    if (emp.id === currentUser.id) {
      notify("You can't delete your own account.", "error");
      return;
    }
    if (emp.role === "admin" && employees.filter((e) => e.role === "admin").length <= 1) {
      notify("You can't delete the last remaining admin.", "error");
      return;
    }
    saveEmployees(
      (prev) => prev.filter((e) => e.id !== emp.id),
      { actor: currentUser.name, action: "EMPLOYEE_DELETED", details: `Deleted ${emp.role} ${emp.name} (${emp.employeeId})${alsoDeleteRecords ? " and their attendance records" : ""}` }
    );
    if (alsoDeleteRecords) {
      saveAttendance((prev) => prev.filter((r) => r.employeeId !== emp.employeeId));
    }
    setDeleteEmp(null);
    notify(`${emp.name} deleted.`);
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <h1 className="text-2xl font-bold text-slate-900">Employees &amp; Admins</h1>
        <Btn variant="accent" onClick={() => setCreating(true)}><PlusCircle size={16}/> Add account</Btn>
      </div>

      <Card className="p-3 mb-4">
        <div className="relative max-w-xs">
          <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input className={inputCls + " pl-8"} placeholder="Search employees" value={search} onChange={(e) => setSearch(e.target.value)} />
        </div>
      </Card>

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
              <th className="px-4 py-3">Employee</th><th className="px-4 py-3">Role</th><th className="px-4 py-3">Department</th><th className="px-4 py-3">Position</th>
              <th className="px-4 py-3">Date hired</th><th className="px-4 py-3">Expected clock-in</th><th className="px-4 py-3">Pay rate</th><th className="px-4 py-3">Status</th><th className="px-4 py-3">Actions</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((e) => (
              <tr key={e.id} className="border-b border-slate-50 hover:bg-slate-50/60">
                <td className="px-4 py-3"><div className="font-semibold text-slate-800">{e.name}</div><div className="text-xs text-slate-400">{e.employeeId} · {e.email}</div></td>
                <td className="px-4 py-3">
                  <span className={`px-2 py-0.5 rounded text-xs font-semibold ${e.role === "admin" ? "bg-slate-800 text-white" : "bg-slate-100 text-slate-600"}`}>{e.role === "admin" ? "Admin" : "Employee"}</span>
                </td>
                <td className="px-4 py-3">{e.department}</td>
                <td className="px-4 py-3">{e.position}</td>
                <td className="px-4 py-3">{fmtDate(e.dateHired)}</td>
                <td className="px-4 py-3 font-mono text-xs">{e.role === "admin" ? "—" : (e.expectedClockIn ? fmtHM(e.expectedClockIn) : "Default")}</td>
                <td className="px-4 py-3">
                  {e.role === "admin" ? <span className="text-slate-300 text-xs">—</span> : (
                    <div className="font-mono text-xs">
                      <div className="font-semibold text-slate-700">{fmtMoney(e.payRate || 0)} <span className="text-slate-400 font-sans">/ {e.payRateType === "hourly" ? "hr" : "day"}</span></div>
                    </div>
                  )}
                </td>
                <td className="px-4 py-3"><StatusPill status={e.status === "active" ? "ON_TIME" : "CLOCKED_OUT"}>{e.status === "active" ? "Active" : "Inactive"}</StatusPill></td>
                <td className="px-4 py-3">
                  <div className="flex gap-1.5">
                    <button title="Edit" onClick={() => setModalEmp(e)} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><Edit3 size={14} /></button>
                    {e.role === "employee" && <button title="Attendance history" onClick={() => setHistoryEmp(e)} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><History size={14} /></button>}
                    {e.role === "employee" && <button title="Pay rate history" onClick={() => setRateHistoryEmp(e)} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><TrendingUp size={14} /></button>}
                    <button title={e.id === currentUser.id ? "You can't deactivate yourself" : (e.status === "active" ? "Deactivate" : "Activate")} disabled={e.id === currentUser.id} onClick={() => toggleStatus(e)} className={`p-1.5 rounded hover:bg-slate-100 disabled:opacity-30 disabled:cursor-not-allowed ${e.status === "active" ? "text-red-500" : "text-emerald-500"}`}>
                      {e.status === "active" ? <Ban size={14} /> : <Check size={14} />}
                    </button>
                    <button title={e.id === currentUser.id ? "You can't delete yourself" : "Delete permanently"} disabled={e.id === currentUser.id} onClick={() => setDeleteEmp(e)} className="p-1.5 rounded hover:bg-red-50 text-red-600 disabled:opacity-30 disabled:cursor-not-allowed">
                      <Trash2 size={14} />
                    </button>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>

      {(modalEmp || creating) && (
        <EmployeeModal employee={modalEmp} employees={employees} onClose={() => { setModalEmp(null); setCreating(false); }} onSave={handleSave} />
      )}
      {historyEmp && <EmployeeHistoryModal employee={historyEmp} attendance={attendance} onClose={() => setHistoryEmp(null)} />}
      {rateHistoryEmp && <RateHistoryModal employee={rateHistoryEmp} onClose={() => setRateHistoryEmp(null)} />}
      {deleteEmp && (
        <DeleteEmployeeModal
          employee={deleteEmp}
          recordCount={attendance.filter((r) => r.employeeId === deleteEmp.employeeId).length}
          onClose={() => setDeleteEmp(null)}
          onConfirm={handleDelete}
        />
      )}
    </div>
  );
}

function DeleteEmployeeModal({ employee, recordCount, onClose, onConfirm }) {
  const [alsoDeleteRecords, setAlsoDeleteRecords] = useState(true);
  const [confirmText, setConfirmText] = useState("");
  const canDelete = confirmText.trim().toUpperCase() === "DELETE";

  return (
    <Modal title={`Delete ${employee.name}?`} onClose={onClose}>
      <div className="flex items-start gap-3 bg-red-50 border border-red-200 text-red-800 rounded-lg px-4 py-3 text-sm mb-4">
        <AlertOctagon size={18} className="mt-0.5 shrink-0" />
        <div>
          This permanently removes <strong>{employee.name}</strong> ({employee.employeeId}) and their login. This cannot be undone — if you just want to block their access, use Deactivate instead.
        </div>
      </div>

      {recordCount > 0 && (
        <label className="flex items-start gap-2 text-sm mb-4 bg-slate-50 rounded-lg p-3">
          <input type="checkbox" className="mt-0.5" checked={alsoDeleteRecords} onChange={(e) => setAlsoDeleteRecords(e.target.checked)} />
          <span>Also delete their <strong>{recordCount}</strong> attendance record{recordCount === 1 ? "" : "s"}. Leaving this unchecked keeps the records (they'll no longer be linked to a listed employee).</span>
        </label>
      )}

      <Field label={<>Type <span className="font-mono font-bold">DELETE</span> to confirm</>}>
        <input className={inputCls} value={confirmText} onChange={(e) => setConfirmText(e.target.value)} placeholder="DELETE" autoFocus />
      </Field>

      <div className="flex justify-end gap-2 mt-2">
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="danger" disabled={!canDelete} onClick={() => onConfirm(employee, alsoDeleteRecords)}><Trash2 size={15} /> Delete permanently</Btn>
      </div>
    </Modal>
  );
}

function RateHistoryModal({ employee, onClose }) {
  const history = employee.rateHistory || [];
  return (
    <Modal title={`Pay rate history — ${employee.name}`} onClose={onClose}>
      <div className="mb-4 bg-slate-50 rounded-lg p-4 text-sm flex items-center justify-between">
        <div>
          <div className="text-xs text-slate-400 uppercase tracking-wide">Current rate</div>
          <div className="font-bold text-slate-900">{fmtMoney(employee.payRate || 0)} <span className="text-slate-400 font-normal">/ {employee.payRateType === "hourly" ? "hour" : "day"}</span></div>
        </div>
        <DollarSign className="text-amber-500" size={22} />
      </div>
      {history.length === 0 ? (
        <div className="text-sm text-slate-400 py-4 text-center">No previous rate changes on record.</div>
      ) : (
        <div className="space-y-2 max-h-72 overflow-y-auto">
          {history.map((h, i) => (
            <div key={i} className="flex items-center justify-between text-sm py-2 border-b border-slate-50 last:border-0">
              <div>
                <div className="font-semibold text-slate-700">{fmtMoney(h.rate || 0)} <span className="text-slate-400 font-normal text-xs">/ {h.payRateType === "hourly" ? "hour" : "day"}</span></div>
                <div className="text-xs text-slate-400">Effective until {fmtDate(h.effectiveUntil)}</div>
              </div>
              <div className="text-right text-xs text-slate-400">
                <div>Changed by {h.changedBy}</div>
                <div>{new Date(h.changedAt).toLocaleString()}</div>
              </div>
            </div>
          ))}
        </div>
      )}
    </Modal>
  );
}

function EmployeeModal({ employee, employees, onClose, onSave }) {
  const isNew = !employee;
  const [form, setForm] = useState(employee || {
    role: "employee",
    employeeId: `EMP-${1000 + employees.filter((e) => e.role === "employee").length + 1}`,
    name: "", username: "", password: "employee123", email: "", department: "", position: "",
    dateHired: todayStr(), expectedClockIn: "", gracePeriodOverride: "",
    payRateType: "daily", payRate: "", rateHistory: [],
  });
  const set = (k, v) => setForm((f) => {
    const next = { ...f, [k]: v };
    if (k === "role" && isNew) {
      next.employeeId = v === "admin"
        ? `ADM-${100 + employees.filter((e) => e.role === "admin").length + 1}`
        : `EMP-${1000 + employees.filter((e) => e.role === "employee").length + 1}`;
    }
    return next;
  });

  const submit = () => {
    if (!form.name || !form.username) return;
    const payload = {
      ...form,
      expectedClockIn: form.expectedClockIn || null,
      gracePeriodOverride: form.gracePeriodOverride ? Number(form.gracePeriodOverride) : null,
      payRateType: form.role === "employee" ? (form.payRateType || "daily") : null,
      payRate: form.role === "employee" ? Number(form.payRate) || 0 : null,
    };
    onSave(payload, isNew);
  };

  return (
    <Modal title={isNew ? "Add employee" : "Edit employee"} onClose={onClose} wide>
      <Field label="Account type">
        <div className="flex gap-2">
          {["employee", "admin"].map((r) => (
            <button key={r} type="button" onClick={() => set("role", r)}
              className={`flex-1 px-3 py-2 rounded-lg text-sm font-semibold border ${form.role === r ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-500 border-slate-200"}`}>
              {r === "admin" ? "Admin" : "Employee"}
            </button>
          ))}
        </div>
        {form.role === "admin" && <p className="text-xs text-slate-400 mt-1.5">Admins have full system access: employee management, attendance edits, OT approval, reports, and settings.</p>}
      </Field>
      <div className="grid grid-cols-2 gap-x-4">
        <Field label="Full name"><input className={inputCls} value={form.name} onChange={(e) => set("name", e.target.value)} /></Field>
        <Field label="Employee ID"><input className={inputCls} value={form.employeeId} onChange={(e) => set("employeeId", e.target.value)} /></Field>
        <Field label="Username"><input className={inputCls} value={form.username} onChange={(e) => set("username", e.target.value)} /></Field>
        <Field label="Password"><input className={inputCls} value={form.password} onChange={(e) => set("password", e.target.value)} /></Field>
        <Field label="Email"><input className={inputCls} value={form.email} onChange={(e) => set("email", e.target.value)} /></Field>
        <Field label="Department"><input className={inputCls} value={form.department} onChange={(e) => set("department", e.target.value)} /></Field>
        <Field label="Position"><input className={inputCls} value={form.position} onChange={(e) => set("position", e.target.value)} /></Field>
        <Field label="Date hired"><input type="date" className={inputCls} value={form.dateHired} onChange={(e) => set("dateHired", e.target.value)} /></Field>
        {form.role === "employee" && (
          <>
            <Field label="Expected clock-in (blank = company default)"><input type="time" className={inputCls} value={form.expectedClockIn || ""} onChange={(e) => set("expectedClockIn", e.target.value)} /></Field>
            <Field label="Grace period override, minutes (blank = default)"><input type="number" className={inputCls} value={form.gracePeriodOverride ?? ""} onChange={(e) => set("gracePeriodOverride", e.target.value)} /></Field>
          </>
        )}
      </div>
      {form.role === "employee" && (
        <div className="border-t border-slate-100 pt-4 mt-1">
          <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-3 flex items-center gap-1.5"><DollarSign size={13}/> Pay rate</div>
          <div className="grid grid-cols-2 gap-x-4">
            <Field label="Rate type">
              <div className="flex gap-2">
                {["daily", "hourly"].map((t) => (
                  <button key={t} type="button" onClick={() => set("payRateType", t)}
                    className={`flex-1 px-3 py-2 rounded-lg text-sm font-semibold border capitalize ${form.payRateType === t ? "bg-slate-900 text-white border-slate-900" : "bg-white text-slate-500 border-slate-200"}`}>
                    {t}
                  </button>
                ))}
              </div>
            </Field>
            <Field label={form.payRateType === "hourly" ? "Rate per hour" : "Rate per day"}>
              <input type="number" step="0.01" className={inputCls} value={form.payRate ?? ""} onChange={(e) => set("payRate", e.target.value)} placeholder="0.00" />
            </Field>
          </div>
          {!isNew && (form.rateHistory || []).length > 0 && (
            <p className="text-xs text-slate-400 -mt-2 mb-2">Changing the rate above will automatically move the current rate into this employee's rate history.</p>
          )}
        </div>
      )}
      <div className="flex justify-end gap-2 mt-3">
        <Btn variant="ghost" onClick={onClose}>Cancel</Btn>
        <Btn variant="primary" onClick={submit}>{isNew ? "Add employee" : "Save changes"}</Btn>
      </div>
    </Modal>
  );
}

function EmployeeHistoryModal({ employee, attendance, onClose }) {
  const rows = attendance.filter((r) => r.employeeId === employee.employeeId).sort((a, b) => b.date.localeCompare(a.date));
  return (
    <Modal title={`Attendance history — ${employee.name}`} onClose={onClose} wide>
      <div className="max-h-96 overflow-y-auto">
        <table className="w-full text-sm">
          <thead><tr className="text-left text-xs uppercase text-slate-400 border-b border-slate-100"><th className="py-2">Date</th><th className="py-2">In/Out</th><th className="py-2">Status</th><th className="py-2">Hours</th><th className="py-2">OT</th></tr></thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.id} className="border-b border-slate-50">
                <td className="py-2">{fmtDate(r.date)}</td>
                <td className="py-2 font-mono text-xs">{fmtTime(r.actualClockIn)}–{fmtTime(r.actualClockOut)}</td>
                <td className="py-2"><StatusPill status={displayStatus(r)} /></td>
                <td className="py-2">{fmtHours(r.totalHours)}</td>
                <td className="py-2">{r.otHours > 0 ? <StatusPill status={r.otStatus} map={OT_STYLE}>{r.otHours}h {r.otStatus}</StatusPill> : "—"}</td>
              </tr>
            ))}
            {rows.length === 0 && <tr><td colSpan={5} className="text-center text-slate-400 py-6">No records.</td></tr>}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

/* ------------------------------------------------------------------ */
/* Reports page                                                        */
/* ------------------------------------------------------------------ */

function ReportsPage({ employees, attendance }) {
  const [empFilter, setEmpFilter] = useState("All");
  const [deptFilter, setDeptFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [otFilter, setOtFilter] = useState("All");
  const [from, setFrom] = useState(offsetDateStr(30));
  const [to, setTo] = useState(todayStr());
  const [showPrint, setShowPrint] = useState(false);

  const departments = useMemo(() => ["All", ...new Set(employees.map((e) => e.department))], [employees]);

  const filtered = useMemo(() => {
    return attendance
      .map((r) => ({ ...r, employee: employees.find((e) => e.employeeId === r.employeeId) }))
      .filter((r) => {
        if (empFilter !== "All" && r.employeeId !== empFilter) return false;
        if (deptFilter !== "All" && r.employee?.department !== deptFilter) return false;
        if (statusFilter !== "All" && displayStatus(r) !== statusFilter) return false;
        if (otFilter !== "All" && r.otStatus !== otFilter) return false;
        if (from && r.date < from) return false;
        if (to && r.date > to) return false;
        return true;
      })
      .sort((a, b) => a.date.localeCompare(b.date));
  }, [attendance, employees, empFilter, deptFilter, statusFilter, otFilter, from, to]);

  const summary = useMemo(() => {
    const totalRegular = filtered.reduce((s, r) => s + (r.regularHours || 0), 0);
    const totalOT = filtered.reduce((s, r) => s + (r.otHours || 0), 0);
    const approvedOT = filtered.filter((r) => r.otStatus === "APPROVED").reduce((s, r) => s + r.otHours, 0);
    const pendingOT = filtered.filter((r) => r.otStatus === "PENDING").reduce((s, r) => s + r.otHours, 0);
    const workingDays = new Set(filtered.map((r) => r.date)).size;
    const lateCount = filtered.filter((r) => r.attendanceStatus === "LATE").length;
    const lateMinutesTotal = filtered.reduce((s, r) => s + (r.lateMinutes || 0), 0);
    return { totalRegular, totalOT, approvedOT, pendingOT, workingDays, lateCount, lateMinutesTotal, records: filtered.length };
  }, [filtered]);

  const exportExcel = () => {
    const data = filtered.map((r) => ({
      "Employee Name": r.employee?.name || r.employeeId,
      "Employee ID": r.employeeId,
      "Department": r.employee?.department || "",
      "Date": r.date,
      "Scheduled Clock-In": r.scheduledClockIn,
      "Actual Clock-In": fmtTime(r.actualClockIn),
      "Status": displayStatus(r).replace(/_/g, " "),
      "Late By (min)": r.lateMinutes || 0,
      "Clock-Out": fmtTime(r.actualClockOut),
      "Total Hours": r.totalHours ?? "",
      "Regular Hours": r.regularHours ?? "",
      "OT Hours": r.otHours || 0,
      "OT Status": r.otStatus,
      "OT Approved By": r.otApprovedBy || "",
      "Notes": r.notes || "",
    }));
    const summaryRows = [
      {}, { "Employee Name": "SUMMARY" },
      { "Employee Name": "Total working days", "Employee ID": summary.workingDays },
      { "Employee Name": "Total regular hours", "Employee ID": summary.totalRegular.toFixed(2) },
      { "Employee Name": "Total OT hours", "Employee ID": summary.totalOT.toFixed(2) },
      { "Employee Name": "Approved OT hours", "Employee ID": summary.approvedOT.toFixed(2) },
      { "Employee Name": "Pending OT hours", "Employee ID": summary.pendingOT.toFixed(2) },
      { "Employee Name": "Late occurrences", "Employee ID": summary.lateCount },
    ];
    const ws = XLSX.utils.json_to_sheet([...data, ...summaryRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance Report");
    XLSX.writeFile(wb, `attendance-report-${from}_to_${to}.xlsx`);
  };

  return (
    <div>
      <h1 className="text-2xl font-bold text-slate-900 mb-4">Reports</h1>
      <Card className="p-4 mb-5">
        <div className="grid md:grid-cols-3 lg:grid-cols-6 gap-3 items-end">
          <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase">Employee</span>
            <select className={inputCls} value={empFilter} onChange={(e) => setEmpFilter(e.target.value)}>
              <option value="All">All</option>
              {employees.filter((e) => e.role === "employee").map((e) => <option key={e.employeeId} value={e.employeeId}>{e.name}</option>)}
            </select>
          </div>
          <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase">Department</span>
            <select className={inputCls} value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)}>{departments.map((d) => <option key={d}>{d}</option>)}</select>
          </div>
          <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase">Status</span>
            <select className={inputCls} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>{STATUS_FILTERS.map((s) => <option key={s} value={s}>{s.replace(/_/g," ")}</option>)}</select>
          </div>
          <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase">OT status</span>
            <select className={inputCls} value={otFilter} onChange={(e) => setOtFilter(e.target.value)}>{OT_FILTERS.map((s) => <option key={s}>{s}</option>)}</select>
          </div>
          <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase">From</span><input type="date" className={inputCls} value={from} onChange={(e) => setFrom(e.target.value)} /></div>
          <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase">To</span><input type="date" className={inputCls} value={to} onChange={(e) => setTo(e.target.value)} /></div>
        </div>
        <div className="flex gap-2 mt-4">
          <Btn variant="success" onClick={exportExcel}><Download size={15}/> Export Excel (.xlsx)</Btn>
          <Btn variant="primary" onClick={() => setShowPrint(true)}><FileText size={15}/> Print / Save as PDF</Btn>
        </div>
      </Card>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-5">
        <StatCard icon={Calendar} label="Working days covered" value={summary.workingDays} />
        <StatCard icon={Clock} label="Total regular hours" value={summary.totalRegular.toFixed(1)} tone="slate" />
        <StatCard icon={AlertTriangle} label="Total OT hours" value={summary.totalOT.toFixed(1)} tone="amber" />
        <StatCard icon={CheckCircle2} label="Approved OT hours" value={summary.approvedOT.toFixed(1)} tone="blue" />
      </div>

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead><tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
            <th className="px-4 py-3">Employee</th><th className="px-4 py-3">Date</th><th className="px-4 py-3">In/Out</th><th className="px-4 py-3">Status</th><th className="px-4 py-3">Hours</th><th className="px-4 py-3">OT</th>
          </tr></thead>
          <tbody>
            {filtered.map((r) => (
              <tr key={r.id} className="border-b border-slate-50">
                <td className="px-4 py-3 font-semibold">{r.employee?.name}</td>
                <td className="px-4 py-3">{fmtDate(r.date)}</td>
                <td className="px-4 py-3 font-mono text-xs">{fmtTime(r.actualClockIn)}–{fmtTime(r.actualClockOut)}</td>
                <td className="px-4 py-3"><StatusPill status={displayStatus(r)} /></td>
                <td className="px-4 py-3">{fmtHours(r.totalHours)}</td>
                <td className="px-4 py-3">{r.otHours > 0 ? `${r.otHours}h (${r.otStatus})` : "—"}</td>
              </tr>
            ))}
            {filtered.length === 0 && <tr><td colSpan={6} className="text-center text-slate-400 py-8">No records in this range.</td></tr>}
          </tbody>
        </table>
      </Card>

      {showPrint && <PrintReportModal rows={filtered} summary={summary} from={from} to={to} onClose={() => setShowPrint(false)} />}
    </div>
  );
}

function PrintReportModal({ rows, summary, from, to, onClose }) {
  useEffect(() => {
    const style = document.createElement("style");
    style.id = "print-style-injected";
    style.innerHTML = `@media print { body * { visibility: hidden !important; } .print-area, .print-area * { visibility: visible !important; } .print-area { position: fixed; top:0; left:0; width:100%; padding: 24px; } }`;
    document.head.appendChild(style);
    return () => { document.getElementById("print-style-injected")?.remove(); };
  }, []);

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-6 py-4 border-b sticky top-0 bg-white">
          <h3 className="font-bold text-lg">Payroll Attendance Report</h3>
          <div className="flex gap-2">
            <Btn variant="primary" onClick={() => window.print()}>Print / Save as PDF</Btn>
            <button onClick={onClose} className="text-slate-400 hover:text-slate-700 ml-2"><X size={20} /></button>
          </div>
        </div>
        <div className="print-area p-6">
          <h2 className="text-xl font-bold">Attendance Report</h2>
          <p className="text-sm text-slate-500 mb-4">{fmtDate(from)} — {fmtDate(to)}</p>
          <div className="grid grid-cols-4 gap-3 mb-5 text-sm">
            <div><div className="text-slate-500 text-xs">Working days</div><div className="font-bold">{summary.workingDays}</div></div>
            <div><div className="text-slate-500 text-xs">Regular hours</div><div className="font-bold">{summary.totalRegular.toFixed(2)}</div></div>
            <div><div className="text-slate-500 text-xs">OT hours (total)</div><div className="font-bold">{summary.totalOT.toFixed(2)}</div></div>
            <div><div className="text-slate-500 text-xs">OT approved</div><div className="font-bold">{summary.approvedOT.toFixed(2)}</div></div>
          </div>
          <table className="w-full text-xs border-collapse">
            <thead><tr className="border-b-2 border-slate-800 text-left"><th className="py-1.5 pr-2">Employee</th><th className="py-1.5 pr-2">Date</th><th className="py-1.5 pr-2">In</th><th className="py-1.5 pr-2">Out</th><th className="py-1.5 pr-2">Status</th><th className="py-1.5 pr-2">Reg</th><th className="py-1.5 pr-2">OT</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.id} className="border-b border-slate-200">
                  <td className="py-1 pr-2">{r.employee?.name}</td>
                  <td className="py-1 pr-2">{r.date}</td>
                  <td className="py-1 pr-2">{fmtTime(r.actualClockIn)}</td>
                  <td className="py-1 pr-2">{fmtTime(r.actualClockOut)}</td>
                  <td className="py-1 pr-2">{displayStatus(r).replace(/_/g," ")}</td>
                  <td className="py-1 pr-2">{r.regularHours ?? ""}</td>
                  <td className="py-1 pr-2">{r.otHours ? `${r.otHours} (${r.otStatus})` : ""}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Payroll page                                                        */
/* ------------------------------------------------------------------ */

function monthLabel(ym) {
  const [y, m] = ym.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" });
}

function PayrollPage({ employees, attendance, settings, holidays, saveHolidays, currentUser, notify }) {
  const [periodType, setPeriodType] = useState("monthly");
  const [ym, setYm] = useState(todayStr().slice(0, 7));
  const [half, setHalf] = useState("1");
  const [customFrom, setCustomFrom] = useState(offsetDateStr(14));
  const [customTo, setCustomTo] = useState(todayStr());
  const [payslipEmp, setPayslipEmp] = useState(null);
  const [showHolidays, setShowHolidays] = useState(false);
  const [showPrint, setShowPrint] = useState(false);

  const { from, to } = computePeriodRange(periodType, ym, half, customFrom, customTo);
  const holidaySet = useMemo(() => new Set(holidays.map((h) => h.date)), [holidays]);

  const activeEmployees = employees.filter((e) => e.role === "employee");

  const payrollRows = useMemo(() => {
    return activeEmployees.map((emp) => {
      const records = attendance.filter((r) => r.employeeId === emp.employeeId && r.date >= from && r.date <= to);
      const figures = computeEmployeePayroll(emp, records, holidaySet, settings);
      return { employee: emp, ...figures };
    }).sort((a, b) => a.employee.name.localeCompare(b.employee.name));
  }, [activeEmployees, attendance, from, to, holidaySet, settings]);

  const totals = useMemo(() => payrollRows.reduce((acc, r) => ({
    hours: acc.hours + r.hoursWorked,
    ot: acc.ot + r.approvedOTHours,
    gross: acc.gross + r.grossPay,
    net: acc.net + r.netPay,
    needsReview: acc.needsReview + (r.needsReview ? 1 : 0),
  }), { hours: 0, ot: 0, gross: 0, net: 0, needsReview: 0 }), [payrollRows]);

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();

    const summaryData = payrollRows.map((r) => ({
      "Employee Name": r.employee.name,
      "Employee ID": r.employee.employeeId,
      "Department": r.employee.department,
      "Pay Type": r.employee.payRateType,
      "Rate": r.employee.payRate,
      "Hours Worked": r.hoursWorked,
      "Approved OT Hours": r.approvedOTHours,
      "Holiday Days": r.holidayDays,
      "Regular Pay": Number(r.regularPay.toFixed(2)),
      "OT Premium": Number(r.otPremium.toFixed(2)),
      "Holiday Pay": Number(r.holidayPay.toFixed(2)),
      "Gross Pay": Number(r.grossPay.toFixed(2)),
      "Pag-IBIG": Number(r.pagibig.toFixed(2)),
      "SSS": Number(r.sss.toFixed(2)),
      "PhilHealth (info only)": Number(r.philhealth.toFixed(2)),
      "Net Pay": Number(r.netPay.toFixed(2)),
      "Needs Review": r.needsReview ? `Yes (${r.missingCount} missing clock-out${r.missingCount > 1 ? "s" : ""})` : "",
    }));
    summaryData.push({}, {
      "Employee Name": "TOTAL",
      "Hours Worked": Number(totals.hours.toFixed(2)),
      "Approved OT Hours": Number(totals.ot.toFixed(2)),
      "Gross Pay": Number(totals.gross.toFixed(2)),
      "Net Pay": Number(totals.net.toFixed(2)),
    });
    const summaryWs = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");

    payrollRows.forEach((r) => {
      const e = r.employee;
      const rows = [
        ["PAYSLIP"],
        ["Name of Employee:", e.name],
        ["Employee ID:", e.employeeId],
        ["Pay Period:", `${fmtDate(from)} — ${fmtDate(to)}`],
        [`${e.payRateType === "hourly" ? "Hourly" : "Daily"} Rate:`, e.payRate],
        [],
        ["Hours Rendered", r.hoursWorked, "Amount", Number(r.regularPay.toFixed(2))],
        ["Approved OT Hours", r.approvedOTHours, "OT Premium", Number(r.otPremium.toFixed(2))],
        ["Holiday Pay (days)", r.holidayDays, "Amount", Number(r.holidayPay.toFixed(2))],
        [],
        ["TOTAL GROSS (w/ OT)", "", "", Number(r.grossPay.toFixed(2))],
        [],
        ["LESS: DEDUCTIONS (Law mandated Benefits)"],
        ["PAG-IBIG CONTRIBUTION", "", "", Number(r.pagibig.toFixed(2))],
        ["SSS CONTRIBUTION", "", "", Number(r.sss.toFixed(2))],
        ["PhilHealth (info only, not deducted)", "", "", Number(r.philhealth.toFixed(2))],
        [],
        ["TOTAL NET PAY", "", "", Number(r.netPay.toFixed(2))],
      ];
      if (r.needsReview) {
        rows.push([], [`NEEDS REVIEW: ${r.missingCount} missing clock-out(s) on ${r.missingDates.join(", ")} — excluded from hours/pay.`]);
      }
      const ws = XLSX.utils.aoa_to_sheet(rows);
      const sheetName = e.name.slice(0, 28).replace(/[\\/?*[\]:]/g, "");
      XLSX.utils.book_append_sheet(wb, ws, sheetName || e.employeeId);
    });

    XLSX.writeFile(wb, `payroll-${from}_to_${to}.xlsx`);
    notify("Payroll workbook exported.");
  };

  return (
    <div>
      <h1 className="text-2xl font-bold text-slate-900 mb-1">Payroll</h1>
      <p className="text-slate-500 text-sm mb-4">{fmtDate(from)} — {fmtDate(to)}</p>

      <Card className="p-4 mb-4">
        <div className="flex flex-wrap gap-3 items-end">
          <div>
            <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">Period</span>
            <select className={inputCls} value={periodType} onChange={(e) => setPeriodType(e.target.value)}>
              <option value="monthly">Monthly</option>
              <option value="semi-monthly">Semi-monthly</option>
              <option value="custom">Custom range</option>
            </select>
          </div>
          {periodType !== "custom" ? (
            <>
              <div>
                <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">Month</span>
                <input type="month" className={inputCls} value={ym} onChange={(e) => setYm(e.target.value)} />
              </div>
              {periodType === "semi-monthly" && (
                <div>
                  <span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">Half</span>
                  <select className={inputCls} value={half} onChange={(e) => setHalf(e.target.value)}>
                    <option value="1">1st — 15th</option>
                    <option value="2">16th — end of month</option>
                  </select>
                </div>
              )}
            </>
          ) : (
            <>
              <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">From</span><input type="date" className={inputCls} value={customFrom} onChange={(e) => setCustomFrom(e.target.value)} /></div>
              <div><span className="block text-xs font-semibold text-slate-500 mb-1 uppercase tracking-wide">To</span><input type="date" className={inputCls} value={customTo} onChange={(e) => setCustomTo(e.target.value)} /></div>
            </>
          )}
          <Btn variant="ghost" onClick={() => setShowHolidays(true)}><CalendarDays size={15}/> Manage holidays</Btn>
          <div className="flex-1" />
          <Btn variant="success" onClick={exportExcel}><Download size={15}/> Export Excel (.xlsx)</Btn>
          <Btn variant="primary" onClick={() => setShowPrint(true)}><Printer size={15}/> Print / Save as PDF</Btn>
        </div>
      </Card>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-5">
        <StatCard icon={Users} label="Employees in run" value={payrollRows.length} />
        <StatCard icon={Clock} label="Total hours" value={totals.hours.toFixed(1)} tone="slate" />
        <StatCard icon={Wallet} label="Total gross pay" value={fmtMoney(totals.gross)} tone="blue" />
        <StatCard icon={AlertOctagon} label="Needs review" value={totals.needsReview} tone={totals.needsReview > 0 ? "red" : "slate"} />
      </div>

      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400">
              <th className="px-4 py-3">Employee</th>
              <th className="px-4 py-3">Hours</th>
              <th className="px-4 py-3">Approved OT</th>
              <th className="px-4 py-3">Holiday days</th>
              <th className="px-4 py-3">Gross</th>
              <th className="px-4 py-3">Deductions</th>
              <th className="px-4 py-3">Net pay</th>
              <th className="px-4 py-3">Actions</th>
            </tr>
          </thead>
          <tbody>
            {payrollRows.map((r) => (
              <tr key={r.employee.id} className="border-b border-slate-50 hover:bg-slate-50/60">
                <td className="px-4 py-3">
                  <div className="flex items-center gap-2">
                    <div className="font-semibold text-slate-800">{r.employee.name}</div>
                    {r.needsReview && (
                      <span title={`${r.missingCount} missing clock-out(s) — excluded from hours/pay`} className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-100 text-amber-800 border border-amber-200">
                        <AlertOctagon size={11}/> NEEDS REVIEW
                      </span>
                    )}
                  </div>
                  <div className="text-xs text-slate-400">{r.employee.employeeId} · {r.employee.department}</div>
                </td>
                <td className="px-4 py-3">{r.hoursWorked.toFixed(2)}h</td>
                <td className="px-4 py-3">{r.approvedOTHours > 0 ? `${r.approvedOTHours.toFixed(2)}h` : <span className="text-slate-300">—</span>}</td>
                <td className="px-4 py-3">{r.holidayDays > 0 ? r.holidayDays : <span className="text-slate-300">—</span>}</td>
                <td className="px-4 py-3 font-semibold">{fmtMoney(r.grossPay)}</td>
                <td className="px-4 py-3 text-xs text-slate-500">{fmtMoney(r.pagibig + r.sss)}</td>
                <td className="px-4 py-3 font-bold text-slate-900">{fmtMoney(r.netPay)}</td>
                <td className="px-4 py-3">
                  <button title="View payslip" onClick={() => setPayslipEmp(r)} className="p-1.5 rounded hover:bg-slate-100 text-slate-500"><Eye size={14} /></button>
                </td>
              </tr>
            ))}
            {payrollRows.length === 0 && <tr><td colSpan={8} className="text-center text-slate-400 py-10">No employees to run payroll for.</td></tr>}
          </tbody>
        </table>
      </Card>

      {payslipEmp && <PayslipModal row={payslipEmp} from={from} to={to} onClose={() => setPayslipEmp(null)} />}
      {showHolidays && <HolidaysModal holidays={holidays} saveHolidays={saveHolidays} currentUser={currentUser} notify={notify} onClose={() => setShowHolidays(false)} />}
      {showPrint && <PayrollPrintModal rows={payrollRows} totals={totals} from={from} to={to} onClose={() => setShowPrint(false)} />}
    </div>
  );
}

function PayslipLines({ row, from, to }) {
  const e = row.employee;
  return (
    <div className="text-sm">
      <div className="text-center mb-4">
        <div className="font-bold text-lg tracking-wide">PAYSLIP</div>
        <div className="text-xs text-slate-400">{fmtDate(from)} — {fmtDate(to)}</div>
      </div>
      <div className="space-y-1 mb-4">
        <div className="flex justify-between"><span className="text-slate-500">Name of Employee:</span><span className="font-semibold">{e.name}</span></div>
        <div className="flex justify-between"><span className="text-slate-500">Employee ID:</span><span className="font-semibold">{e.employeeId}</span></div>
        <div className="flex justify-between"><span className="text-slate-500">{e.payRateType === "hourly" ? "Hourly" : "Daily"} Rate:</span><span className="font-semibold">{fmtMoney(e.payRate || 0)}</span></div>
      </div>
      <table className="w-full mb-3">
        <tbody>
          <tr className="border-b border-slate-100"><td className="py-1.5 text-slate-500">Hours Rendered ({row.hoursWorked.toFixed(2)}h)</td><td className="py-1.5 text-right font-semibold">{fmtMoney(row.regularPay)}</td></tr>
          <tr className="border-b border-slate-100"><td className="py-1.5 text-slate-500">Overtime Premium ({row.approvedOTHours.toFixed(2)}h approved @ {settingsMultiplierLabel(row)})</td><td className="py-1.5 text-right font-semibold">{fmtMoney(row.otPremium)}</td></tr>
          <tr className="border-b border-slate-100"><td className="py-1.5 text-slate-500">Holiday Pay ({row.holidayDays} day{row.holidayDays === 1 ? "" : "s"})</td><td className="py-1.5 text-right font-semibold">{fmtMoney(row.holidayPay)}</td></tr>
          <tr className="border-t-2 border-slate-800"><td className="py-2 font-bold">TOTAL GROSS (w/ OT)</td><td className="py-2 text-right font-bold">{fmtMoney(row.grossPay)}</td></tr>
        </tbody>
      </table>
      <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-2">Less: Deductions (law mandated benefits)</div>
      <table className="w-full mb-3">
        <tbody>
          <tr className="border-b border-slate-50"><td className="py-1.5 text-slate-500">Pag-IBIG Contribution</td><td className="py-1.5 text-right">{fmtMoney(row.pagibig)}</td></tr>
          <tr className="border-b border-slate-50"><td className="py-1.5 text-slate-500">SSS Contribution</td><td className="py-1.5 text-right">{fmtMoney(row.sss)}</td></tr>
          <tr className="border-b border-slate-50"><td className="py-1.5 text-slate-500">PhilHealth <span className="text-slate-300">(info only, not deducted)</span></td><td className="py-1.5 text-right text-slate-400">{fmtMoney(row.philhealth)}</td></tr>
        </tbody>
      </table>
      <div className="flex justify-between items-center bg-slate-900 text-white rounded-lg px-4 py-3 font-bold">
        <span>TOTAL NET PAY</span><span>{fmtMoney(row.netPay)}</span>
      </div>
      {row.needsReview && (
        <div className="mt-4 flex items-start gap-2 bg-amber-50 border border-amber-200 text-amber-800 rounded-lg px-3 py-2 text-xs">
          <AlertOctagon size={14} className="mt-0.5 shrink-0"/>
          <span>{row.missingCount} day{row.missingCount > 1 ? "s" : ""} in this period ({row.missingDates.map((d) => fmtDate(d)).join(", ")}) had no clock-out and were excluded from hours and pay. Fix the record in Attendance Records to include them.</span>
        </div>
      )}
    </div>
  );
}

function settingsMultiplierLabel(row) {
  // derived purely for display; the multiplier itself lives in Settings
  if (!row.approvedOTHours) return "OT rate";
  const premiumPerHour = row.hourlyRate ? row.otPremium / row.approvedOTHours / row.hourlyRate : 0;
  return `${(1 + premiumPerHour).toFixed(2)}x`;
}

function PayslipModal({ row, from, to, onClose }) {
  return (
    <Modal title={`Payslip — ${row.employee.name}`} onClose={onClose}>
      <PayslipLines row={row} from={from} to={to} />
    </Modal>
  );
}

function PayrollPrintModal({ rows, totals, from, to, onClose }) {
  useEffect(() => {
    const style = document.createElement("style");
    style.id = "print-style-injected-payroll";
    style.innerHTML = `@media print { body * { visibility: hidden !important; } .print-area, .print-area * { visibility: visible !important; } .print-area { position: fixed; top:0; left:0; width:100%; padding: 24px; } .payslip-page { page-break-after: always; } }`;
    document.head.appendChild(style);
    return () => { document.getElementById("print-style-injected-payroll")?.remove(); };
  }, []);

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-3xl max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-6 py-4 border-b sticky top-0 bg-white">
          <h3 className="font-bold text-lg">Payroll — Payslips</h3>
          <div className="flex gap-2">
            <Btn variant="primary" onClick={() => window.print()}>Print / Save as PDF</Btn>
            <button onClick={onClose} className="text-slate-400 hover:text-slate-700 ml-2"><X size={20} /></button>
          </div>
        </div>
        <div className="print-area p-6">
          <h2 className="text-xl font-bold mb-1">Payroll Summary</h2>
          <p className="text-sm text-slate-500 mb-4">{fmtDate(from)} — {fmtDate(to)}</p>
          <table className="w-full text-xs border-collapse mb-8">
            <thead><tr className="border-b-2 border-slate-800 text-left"><th className="py-1.5 pr-2">Employee</th><th className="py-1.5 pr-2">Hours</th><th className="py-1.5 pr-2">OT</th><th className="py-1.5 pr-2">Gross</th><th className="py-1.5 pr-2">Deductions</th><th className="py-1.5 pr-2">Net</th></tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.employee.id} className="border-b border-slate-200">
                  <td className="py-1 pr-2">{r.employee.name}{r.needsReview ? " *" : ""}</td>
                  <td className="py-1 pr-2">{r.hoursWorked.toFixed(2)}</td>
                  <td className="py-1 pr-2">{r.approvedOTHours.toFixed(2)}</td>
                  <td className="py-1 pr-2">{fmtMoney(r.grossPay)}</td>
                  <td className="py-1 pr-2">{fmtMoney(r.pagibig + r.sss)}</td>
                  <td className="py-1 pr-2 font-semibold">{fmtMoney(r.netPay)}</td>
                </tr>
              ))}
              <tr className="border-t-2 border-slate-800 font-bold"><td className="py-1.5 pr-2">TOTAL</td><td className="py-1.5 pr-2">{totals.hours.toFixed(2)}</td><td className="py-1.5 pr-2">{totals.ot.toFixed(2)}</td><td className="py-1.5 pr-2">{fmtMoney(totals.gross)}</td><td className="py-1.5 pr-2"></td><td className="py-1.5 pr-2">{fmtMoney(totals.net)}</td></tr>
            </tbody>
          </table>
          {totals.needsReview > 0 && <p className="text-xs text-slate-500 mb-8">* Needs review — one or more missing clock-outs excluded from hours/pay in this period.</p>}

          {rows.map((r) => (
            <div key={r.employee.id} className="payslip-page mb-10 border-t pt-6">
              <PayslipLines row={r} from={from} to={to} />
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function HolidaysModal({ holidays, saveHolidays, currentUser, notify, onClose }) {
  const [date, setDate] = useState(todayStr());
  const [name, setName] = useState("");

  const addHoliday = () => {
    if (!date) return;
    if (holidays.some((h) => h.date === date)) { notify("That date is already on the holiday list.", "error"); return; }
    const rec = { id: `hol-${Date.now()}`, date, name: name || "Holiday" };
    saveHolidays((prev) => [...prev, rec].sort((a, b) => a.date.localeCompare(b.date)), { actor: currentUser.name, action: "HOLIDAY_ADDED", details: `Added holiday ${rec.date} (${rec.name})` });
    setName("");
    notify("Holiday added.");
  };

  const removeHoliday = (h) => {
    saveHolidays((prev) => prev.filter((x) => x.id !== h.id), { actor: currentUser.name, action: "HOLIDAY_REMOVED", details: `Removed holiday ${h.date} (${h.name})` });
    notify("Holiday removed.");
  };

  return (
    <Modal title="Manage holidays" onClose={onClose}>
      <p className="text-xs text-slate-400 mb-4">Employees who clock in and out on one of these dates automatically receive one extra day's pay on top of hours worked.</p>
      <div className="flex gap-2 items-end mb-4">
        <Field label="Date"><input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        <Field label="Name (optional)"><input className={inputCls} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Independence Day" /></Field>
        <Btn variant="accent" onClick={addHoliday} className="mb-3"><PlusCircle size={15}/> Add</Btn>
      </div>
      <div className="space-y-1.5 max-h-64 overflow-y-auto">
        {holidays.length === 0 && <div className="text-sm text-slate-400 py-4 text-center">No holidays configured yet.</div>}
        {holidays.map((h) => (
          <div key={h.id} className="flex items-center justify-between text-sm py-2 px-3 rounded-lg bg-slate-50">
            <div><span className="font-mono font-semibold">{fmtDate(h.date)}</span> <span className="text-slate-500 ml-2">{h.name}</span></div>
            <button onClick={() => removeHoliday(h)} className="text-red-500 hover:text-red-700"><X size={15} /></button>
          </div>
        ))}
      </div>
    </Modal>
  );
}

/* ------------------------------------------------------------------ */
/* Settings page                                                       */
/* ------------------------------------------------------------------ */

function SettingsPage({ settings, saveSettings, currentUser, notify }) {
  const [form, setForm] = useState(settings);
  const set = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const toggleDay = (d) => setForm((f) => ({ ...f, workingDays: f.workingDays.includes(d) ? f.workingDays.filter((x) => x !== d) : [...f.workingDays, d].sort() }));

  const submit = () => {
    saveSettings(form, { actor: currentUser.name, action: "SETTINGS_UPDATED", details: `Shift ${form.shiftStart}-${form.shiftEnd}, grace ${form.gracePeriodMinutes}min, OT threshold ${form.otThresholdHours}h` });
    notify("Settings saved.");
  };

  return (
    <div className="max-w-xl">
      <h1 className="text-2xl font-bold text-slate-900 mb-4">Shift &amp; Attendance Settings</h1>
      <Card className="p-6">
        <div className="grid grid-cols-2 gap-x-4">
          <Field label="Shift start"><input type="time" className={inputCls} value={form.shiftStart} onChange={(e) => set("shiftStart", e.target.value)} /></Field>
          <Field label="Shift end"><input type="time" className={inputCls} value={form.shiftEnd} onChange={(e) => set("shiftEnd", e.target.value)} /></Field>
          <Field label="Standard working hours"><input type="number" step="0.5" className={inputCls} value={form.standardHours} onChange={(e) => set("standardHours", Number(e.target.value))} /></Field>
          <Field label="OT threshold (hours)"><input type="number" step="0.5" className={inputCls} value={form.otThresholdHours} onChange={(e) => set("otThresholdHours", Number(e.target.value))} /></Field>
          <Field label="Late grace period (minutes)"><input type="number" className={inputCls} value={form.gracePeriodMinutes} onChange={(e) => set("gracePeriodMinutes", Number(e.target.value))} /></Field>
        </div>
        <Field label="Working days">
          <div className="flex gap-2">
            {DAY_NAMES.map((d, i) => (
              <button key={i} onClick={() => toggleDay(i)} className={`w-10 h-10 rounded-lg text-xs font-bold ${form.workingDays.includes(i) ? "bg-amber-500 text-white" : "bg-slate-100 text-slate-400"}`}>{d}</button>
            ))}
          </div>
        </Field>
        <p className="text-xs text-slate-400 mb-4">These are company-wide defaults. Individual employees can have a custom expected clock-in and grace period set on their profile, which overrides these values.</p>
        <Btn variant="accent" onClick={submit}>Save settings</Btn>
      </Card>

      <h1 className="text-2xl font-bold text-slate-900 mb-4 mt-8">Payroll Settings</h1>
      <Card className="p-6">
        <div className="grid grid-cols-2 gap-x-4">
          <Field label="Overtime multiplier (e.g. 1.25 = 125%)">
            <input type="number" step="0.01" className={inputCls} value={form.otMultiplier} onChange={(e) => set("otMultiplier", Number(e.target.value))} />
          </Field>
        </div>
        <p className="text-xs text-slate-400 mb-4">Approved overtime hours are paid at this multiple of the employee's hourly rate. Straight-time pay for the hours already worked is calculated separately, so this only affects the OT premium.</p>

        <div className="text-xs font-bold uppercase tracking-wide text-slate-400 mb-3 mt-2">Mandated contributions</div>
        <div className="grid grid-cols-3 gap-x-4">
          <Field label="Pag-IBIG rate">
            <div className="relative"><input type="number" step="0.001" className={inputCls} value={form.pagibigRate} onChange={(e) => set("pagibigRate", Number(e.target.value))} /></div>
          </Field>
          <Field label="Pag-IBIG cap (₱ of gross)">
            <input type="number" step="1" className={inputCls} value={form.pagibigCap} onChange={(e) => set("pagibigCap", Number(e.target.value))} />
          </Field>
          <div />
          <Field label="SSS rate">
            <input type="number" step="0.001" className={inputCls} value={form.sssRate} onChange={(e) => set("sssRate", Number(e.target.value))} />
          </Field>
          <Field label="SSS cap (₱ of gross)">
            <input type="number" step="1" className={inputCls} value={form.sssCap} onChange={(e) => set("sssCap", Number(e.target.value))} />
          </Field>
          <div />
          <Field label="PhilHealth rate">
            <input type="number" step="0.001" className={inputCls} value={form.philhealthRate} onChange={(e) => set("philhealthRate", Number(e.target.value))} />
          </Field>
        </div>
        <p className="text-xs text-slate-400 mb-4">Rates are expressed as decimals (2% = 0.02). Pag-IBIG and SSS are computed on gross pay capped at the amount above, then subtracted to get net pay. PhilHealth is shown on the payslip for reference but — matching the company's payroll sheet — is not deducted from net pay.</p>
        <Btn variant="accent" onClick={submit}>Save payroll settings</Btn>
      </Card>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Audit log page                                                      */
/* ------------------------------------------------------------------ */

function AuditPage({ auditLog }) {
  return (
    <div>
      <h1 className="text-2xl font-bold text-slate-900 mb-4">Audit Log</h1>
      <Card className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead><tr className="border-b border-slate-100 text-left text-xs uppercase tracking-wide text-slate-400"><th className="px-4 py-3">When</th><th className="px-4 py-3">Actor</th><th className="px-4 py-3">Action</th><th className="px-4 py-3">Details</th></tr></thead>
          <tbody>
            {auditLog.map((a) => (
              <tr key={a.id} className="border-b border-slate-50">
                <td className="px-4 py-3 whitespace-nowrap text-xs text-slate-500">{new Date(a.timestamp).toLocaleString()}</td>
                <td className="px-4 py-3 font-semibold">{a.actor}</td>
                <td className="px-4 py-3"><span className="px-2 py-0.5 rounded bg-slate-100 text-slate-600 text-xs font-mono">{a.action}</span></td>
                <td className="px-4 py-3 text-slate-600">{a.details}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
      <p className="text-xs text-slate-400 mt-3">This log is append-only and cannot be edited by any account, including Admin, from within the app.</p>
    </div>
  );
}
