import React, { useState, useEffect, useCallback } from 'react';
import ReactDOM from 'react-dom';
import ExcelJS from 'exceljs';
import { api, SimpleRow, PivotResult } from './api';

// ─── Report catalogue ──────────────────────────────────────────────────────────

type ReportKind = 'simple' | 'pivot' | 'daterange' | 'futureage' | 'futureweight' | 'custom' | 'animaledit';

interface ReportDef {
  id: string;
  icon: string;
  title: string;
  desc: string;
  kind: ReportKind;
  group: string;
  load?: () => Promise<SimpleRow[] | PivotResult>;
  filterCols?: string[];
  totalCols?: string[];
  defaultHidden?: string[];
}

const REPORTS: ReportDef[] = [
  // Custom (top)
  { id: 'animal_edit', icon: '✏️', title: 'View / Edit Animal', desc: 'Search by last 4 digits, view or edit all animal details', kind: 'animaledit', group: 'Custom' },
  { id: 'custom',            icon: '🔧', title: 'Custom Filter',         desc: 'Combine any filters for a tailored report',               kind: 'custom',  group: 'Custom' },
  // Herd
  { id: 'all_animals',       icon: '🐄', title: 'All Animals',          desc: 'Complete list incl. dead & sold, last weight & pregnancy', kind: 'simple',  group: 'Herd',     load: () => api.allAnimals() },
  { id: 'heifers_1624',      icon: '🐮', title: 'Heifers 16–24 m',      desc: 'Female 16–24 months, alive',                              kind: 'simple',  group: 'Herd',     load: () => api.bySex('F', 16, 24) },
  { id: 'heifers_1216',      icon: '🐄', title: 'Heifers 12–16 m',      desc: 'Female 12–16 months, alive',                              kind: 'simple',  group: 'Herd',     load: () => api.bySex('F', 12, 16) },
  { id: 'young_females',     icon: '🐣', title: 'Young Females 6–12 m', desc: 'Female 6–12 months, alive',                               kind: 'simple',  group: 'Herd',     load: () => api.bySex('F', 6, 12) },
  { id: 'female_calves',     icon: '🐣', title: 'Female Calves 0–6 m',  desc: 'Female calves under 6 months, alive',                     kind: 'simple',  group: 'Herd',     load: () => api.bySex('F', 0, 6) },
  { id: 'bulls',             icon: '🐂', title: 'Bulls',                 desc: 'Male animals with bull tag, alive',                        kind: 'simple',  group: 'Herd',     load: () => api.bulls() },
  { id: 'males_16plus',      icon: '🐃', title: 'Males > 16 m',         desc: 'Males over 16 months (excl. bulls), alive',               kind: 'simple',  group: 'Herd',     load: () => api.males16plus() },
  { id: 'males_1216',        icon: '🐃', title: 'Males 12–16 m',        desc: 'Males 12–16 months, alive',                               kind: 'simple',  group: 'Herd',     load: () => api.males1216() },
  { id: 'young_males',       icon: '🐂', title: 'Young Males 6–12 m',   desc: 'Males 6–12 months, alive',                                kind: 'simple',  group: 'Herd',     load: () => api.youngMales() },
  { id: 'male_calves',       icon: '🐣', title: 'Male Calves 0–12 m',   desc: 'Male calves under 12 months, alive',                      kind: 'simple',  group: 'Herd',     load: () => api.maleCalves() },
  // Pregnancy
  { id: 'current_pregnant',  icon: '🤰', title: 'Currently Pregnant',   desc: 'Alive females with latest pregnancy check',               kind: 'simple',  group: 'Pregnancy', load: () => api.currentPregnant() },
  { id: 'pregnancy_history', icon: '🔬', title: 'Pregnancy History',    desc: 'All pregnancy checks per female, last 3 years',           kind: 'pivot',   group: 'Pregnancy', load: () => api.pregnancyHistory() },
  { id: 'lost_pregnancy',    icon: '💔', title: 'Lost Pregnancy',        desc: 'Animals that lost their pregnancy',                        kind: 'simple',  group: 'Pregnancy', load: () => api.lostPregnancy() },
  { id: 'pending_births',    icon: '👶', title: 'Pending Father',        desc: 'Births not yet assigned a father',                         kind: 'simple',  group: 'Pregnancy', load: () => api.pendingBirths() },
  // Medical
  { id: 'medical_simple',    icon: '📋', title: 'Medical Summary',      desc: 'Procedure summary with animal count, last 5 years',        kind: 'simple',  group: 'Medical',  load: () => api.medicalSimple() },
  { id: 'medical_group',     icon: '💉', title: 'Group Treatments',     desc: 'Group procedures per animal, last 5 years',               kind: 'pivot',   group: 'Medical',  load: () => api.medicalGroup() },
  { id: 'medical_individual',icon: '🩺', title: 'Individual Treatments',desc: 'Individual treatments per animal, last 2 years',           kind: 'pivot',   group: 'Medical',  load: () => api.medicalIndividual() },
  // Weights & AI
  { id: 'weights',           icon: '⚖️', title: 'Weighing History',     desc: 'All weighings per animal, last 3 years',                  kind: 'pivot',   group: 'Weights',  load: () => api.weights() },
  { id: 'ai_history',        icon: '🧬', title: 'AI History',           desc: 'Artificial insemination records, last 3 years',            kind: 'pivot',   group: 'AI',       load: () => api.aiHistory() },
  // Parameterised
  { id: 'births',            icon: '🐣', title: 'Births – Chronological', desc: 'All births newest first — ear tag, sex, parents, weight',  kind: 'simple',  group: 'Events', load: () => api.births() },
  { id: 'close_births',      icon: '⚠️', title: 'Close Births < 9 m',    desc: 'Cows with two births less than 9 months apart',            kind: 'simple',  group: 'Events', load: () => api.closeBirths() },
  { id: 'ai_prepared',       icon: '💉', title: 'AI – Prepared',          desc: 'Cows prepared for insemination – disappear after 3 weeks or when inseminated', kind: 'simple', group: 'Events', load: () => api.aiPrepared() },
  { id: 'ai_inseminations',  icon: '🧬', title: 'AI – Inseminations',     desc: 'All artificial inseminations – cow, bull name & code, date', kind: 'simple', group: 'Events', load: () => api.aiInseminations() },
  { id: 'calving',           icon: '🐄', title: 'Calvings in Period',   desc: 'Births within a chosen date range',                        kind: 'daterange',group: 'Events' },
  { id: 'not_calving',       icon: '⏳', title: 'Expected but No Birth',desc: 'Expected calvings that didn\'t happen',                    kind: 'daterange',group: 'Events' },
  { id: 'dead',              icon: '💀', title: 'Deaths in Period',      desc: 'Deaths within a chosen date range',                        kind: 'daterange',group: 'Events' },
  { id: 'future_age',        icon: '📅', title: 'Age on Future Date',    desc: 'Animals in an age window on a future date',               kind: 'futureage',group: 'Events' },
  { id: 'future_weight',     icon: '📈', title: 'Weight on Future Date', desc: 'Animals in a weight window on a future date',             kind: 'futureweight',group: 'Events' },
  // Bales
  { id: 'bale_production',   icon: '🌾', title: 'Bales Production',      desc: 'All bale production incl. area, date, weight, count',      kind: 'simple', group: 'Bales',  load: () => api.baleProduction(), filterCols: ['bale_type', 'area', 'year'], totalCols: ['count', 'total_t'], defaultHidden: ['created_time'] },
  { id: 'bale_farm_entry',   icon: '🚜', title: 'Bales Farm Entry',       desc: 'Bales moved from field to farm (not purchased)',           kind: 'simple', group: 'Bales',  load: () => api.baleFarmEntry(),   filterCols: ['bale_type', 'year'],             totalCols: ['count', 'total_t'], defaultHidden: ['created_time'] },
  { id: 'bale_purchase',     icon: '🛒', title: 'Bales Purchase',          desc: 'Purchased bales entered into farm stock',                 kind: 'simple', group: 'Bales',  load: () => api.balePurchase(),    filterCols: ['bale_type', 'year'],             totalCols: ['count', 'total_t'], defaultHidden: ['created_time'] },
  { id: 'bale_consumption',  icon: '📦', title: 'Bales Consumption',      desc: 'All bale consumption — field or farm, weight & tonnage',   kind: 'simple', group: 'Bales',  load: () => api.baleConsumption(), filterCols: ['bale_type', 'location', 'year'], totalCols: ['count', 'total_t'], defaultHidden: ['created_time'] },
  // Grains
  { id: 'cereal_production', icon: '🌽', title: 'Grains Production',  desc: 'Own cereal harvest only',                               kind: 'simple', group: 'Grains', load: () => api.cerealProduction(),  filterCols: ['cereal_type', 'year'],           totalCols: ['harvest_t'],     defaultHidden: ['created_time'] },
  { id: 'cereal_purchase',   icon: '🛒', title: 'Grains Purchase',    desc: 'Purchased cereals (Shrot de soia, Shrot de floare, …)', kind: 'simple', group: 'Grains', load: () => api.cerealPurchase(),    filterCols: ['cereal_type', 'supplier', 'year'], totalCols: ['purchase_t'], defaultHidden: ['created_time'] },
  { id: 'cereal_consumption',icon: '🍽️', title: 'Grains Consumption', desc: 'All cereal consumption entries',                        kind: 'simple', group: 'Grains', load: () => api.cerealConsumption(), filterCols: ['cereal_type', 'year'],           totalCols: ['consumption_t'], defaultHidden: ['created_time'] },
  { id: 'cereal_sale',       icon: '💰', title: 'Grains Sale',        desc: 'All cereal sales — type, quantity, client, date',       kind: 'simple', group: 'Grains', load: () => api.cerealSale(),        filterCols: ['cereal_type', 'client', 'year'], totalCols: ['sale_t'],        defaultHidden: ['created_time'] },
];

const ANIMAL_GROUPS = ['Custom', 'Herd', 'Pregnancy', 'Medical', 'Weights', 'AI', 'Events'];
const FEED_GROUPS   = ['Bales', 'Grains'];

// ─── Helpers ───────────────────────────────────────────────────────────────────

function formatDate(val: string): string {
  const m = val.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? `${m[3]}.${m[2]}.${m[1]}` : val;
}

function formatColHeader(col: string): string {
  if (/^date:\d{4}-\d{2}-\d{2}$/.test(col))           return col.slice(5);
  if (/^date:\d{4}-\d{2}-\d{2}:type$/.test(col))      return `Type (${col.split(':')[1]})`;
  if (/^date:\d{4}-\d{2}-\d{2}:med$/.test(col))       return `Med (${col.split(':')[1]})`;
  if (/^treat:\d+:type$/.test(col))                    return `Treat ${col.split(':')[1]} Type`;
  if (/^treat:\d+:date$/.test(col))                    return `Treat ${col.split(':')[1]} Date`;
  if (/^ai:\d+:type$/.test(col))                       return `AI ${col.split(':')[1]} Type`;
  if (/^ai:\d+:bull$/.test(col))                       return `AI ${col.split(':')[1]} Bull`;
  if (/^ai:\d+:date$/.test(col))                       return `AI ${col.split(':')[1]} Date`;
  return col.replace(/_/g, ' ').replace(/^\w/, c => c.toUpperCase());
}

function normBool(s: string): string | null {
  const u = s.trim().toUpperCase();
  if (u === 'YES' || u === 'DA' || u === 'TRUE') return 'DA';
  if (u === 'NO'  || u === 'NU' || u === 'FALSE') return 'NU';
  return null;
}

const str = (v: unknown): string => (v == null ? '' : String(v));

function normColorIn(s: string): string {
  const u = s.toUpperCase();
  if (['BLACK', 'B', 'N', 'NEGRU'].includes(u)) return 'BLACK';
  if (['RED', 'R', 'ROSU'].includes(u)) return 'RED';
  return s;
}

function normColor(s: string): string | null {
  const u = s.trim().toUpperCase();
  if (u === 'BLACK' || u === 'B' || u === 'N' || u === 'NEGRU') return 'NEGRU';
  if (u === 'RED'   || u === 'R' || u === 'ROSU')               return 'ROSU';
  return null;
}

function formatCell(val: unknown, col: string): string {
  if (val === null || val === undefined) return '';
  if (val === true)  return 'DA';
  if (val === false) return 'NU';
  const s = String(val);
  const b = normBool(s);
  if (b !== null) return b;
  if (col === 'color') { const c = normColor(s); if (c !== null) return c; }
  if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}/.test(val) && !col.startsWith('date:')) return formatDate(val);
  return s;
}

function colMinWidth(col: string): number {
  if (col === 'ear_tag') return 140;
  if (col.includes('date') || /^\d{4}-\d{2}-\d{2}/.test(col)) return 96;
  if (col.includes('weight') || col.includes('kg')) return 72;
  if (col === 'status') return 72;
  if (col === 'company') return 160;
  return 110;
}

async function exportXLSX(columns: string[], rows: SimpleRow[], title: string) {
  const headers = columns.map(formatColHeader);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Farm Reports';
  const ws = wb.addWorksheet('Report');

  // Column widths (calculated before adding rows)
  ws.columns = columns.map((col, i) => {
    const maxDataLen = rows.reduce((max, r) => {
      const v = r[col];
      return Math.max(max, v === null || v === undefined ? 0 : String(v).length);
    }, 0);
    return {
      width: Math.min(Math.max(headers[i].length, maxDataLen) + 3, 42),
    };
  });

  // Header row — bold, blue background, white text
  const headerRow = ws.addRow(headers);
  headerRow.eachCell(cell => {
    cell.font  = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1565C0' } };
    cell.alignment = { vertical: 'middle', wrapText: false };
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFAAAAAA' } } };
  });
  headerRow.height = 22;

  // Freeze top row
  ws.views = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2', activeCell: 'A1', showGridLines: true }];

  // AutoFilter on the entire header row
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };

  // Data rows — zebra striping
  rows.forEach((row, ri) => {
    const dataRow = ws.addRow(
      columns.map(c => {
        const v = row[c];
        return v === null || v === undefined ? '' : v;
      })
    );
    if (ri % 2 === 1) {
      dataRow.eachCell({ includeEmpty: true }, cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F4FF' } };
      });
    }
    dataRow.height = 18;
  });

  // Download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${title.replace(/[^\w\s-]/g, '').replace(/\s+/g, '_')}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

// ─── Custom filter state ───────────────────────────────────────────────────────

interface CustomFilter {
  sex: 'both' | 'M' | 'F';
  viu: boolean; vandut: boolean; mort: boolean;
  minAge: string; maxAge: string;
  minKg: string; maxKg: string;
  companies: string[] | null;  // null = all, [] = none, [...] = specific
  groups:    string[] | null;
}

const defaultCustom: CustomFilter = {
  sex: 'both', viu: true, vandut: false, mort: false,
  minAge: '', maxAge: '', minKg: '', maxKg: '',
  companies: null, groups: null,
};

function buildCustomParams(f: CustomFilter): Record<string, string> {
  const p: Record<string, string> = {};
  if (f.sex !== 'both') p.sex = f.sex;
  const st = [f.viu && 'VIU', f.vandut && 'VANDUT', f.mort && 'MORT'].filter(Boolean).join(',');
  if (st) p.status = st;
  if (f.minAge) p.min_age_months = f.minAge;
  if (f.maxAge) p.max_age_months = f.maxAge;
  if (f.companies !== null && f.companies.length > 0) p.company = f.companies.join(',');
  if (f.groups    !== null && f.groups.length    > 0) p.group_name = f.groups.join(',');
  if (f.minKg) p.min_weight_kg = f.minKg;
  if (f.maxKg) p.max_weight_kg = f.maxKg;
  return p;
}

// ─── MultiSelect dropdown ─────────────────────────────────────────────────────

function MultiSelect({
  options, value, onChange, placeholder,
}: {
  options: string[];
  value: string[] | null;
  onChange: (v: string[] | null) => void;
  placeholder: string;
}) {
  const [open, setOpen] = React.useState(false);
  const [pos, setPos]   = React.useState({ top: 0, left: 0, width: 0 });
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  const panelRef   = React.useRef<HTMLDivElement>(null);

  // Close on outside click
  React.useEffect(() => {
    const handler = (e: MouseEvent) => {
      const t = e.target as Node;
      if (
        triggerRef.current && !triggerRef.current.contains(t) &&
        panelRef.current   && !panelRef.current.contains(t)
      ) setOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, []);

  const handleToggle = () => {
    if (!open && triggerRef.current) {
      const r = triggerRef.current.getBoundingClientRect();
      setPos({ top: r.bottom + 4, left: r.left, width: r.width });
    }
    setOpen(o => !o);
  };

  const isChecked  = (item: string) => value === null || value.includes(item);
  const allChecked = value === null || value.length === options.length;

  const toggle = (item: string) => {
    if (value === null) {
      onChange(options.filter(o => o !== item));
    } else if (value.includes(item)) {
      onChange(value.filter(o => o !== item));
    } else {
      const next = [...value, item];
      onChange(next.length === options.length ? null : next);
    }
  };

  const checkedCount = value === null ? options.length : value.length;
  const label = options.length === 0
    ? placeholder
    : allChecked   ? `All (${options.length})`
    : checkedCount === 0 ? 'None selected'
    : `${checkedCount} / ${options.length} selected`;

  const panel = open && options.length > 0 && ReactDOM.createPortal(
    <div ref={panelRef} style={{
      ...ms.panel,
      position: 'fixed',
      top:  pos.top,
      left: pos.left,
      width: Math.max(pos.width, 260),
    }}>
      <div style={ms.topRow}>
        <button style={ms.quickBtn} onClick={() => onChange(null)}>All</button>
        <button style={ms.quickBtn} onClick={() => onChange([])}>None</button>
      </div>
      <div style={ms.divider} />
      <div style={ms.list}>
        {options.map(opt => (
          <label key={opt} style={ms.item}>
            <input
              type="checkbox"
              checked={isChecked(opt)}
              onChange={() => toggle(opt)}
              style={{ marginRight: 8, accentColor: '#1565C0', cursor: 'pointer' }}
            />
            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {opt}
            </span>
          </label>
        ))}
      </div>
    </div>,
    document.body
  );

  return (
    <div style={{ position: 'relative', width: 220 }}>
      <button ref={triggerRef} style={ms.trigger} onClick={handleToggle}>
        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1 }}>
          {label}
        </span>
        <span style={{ flexShrink: 0, marginLeft: 6, color: '#9CA3AF' }}>{open ? '▲' : '▼'}</span>
      </button>
      {panel}
    </div>
  );
}

const ms = {
  trigger: {
    display: 'flex', alignItems: 'center', width: '100%', height: 32,
    border: '1px solid #D1D5DB', borderRadius: 6, background: '#FFF',
    padding: '0 10px', fontSize: 13, color: '#111827', cursor: 'pointer',
    fontFamily: 'inherit', textAlign: 'left' as const,
  } as React.CSSProperties,
  panel: {
    position: 'absolute' as const, top: 36, left: 0, zIndex: 100,
    background: '#FFF', border: '1px solid #D1D5DB', borderRadius: 8,
    boxShadow: '0 4px 16px rgba(0,0,0,0.13)', width: 260, overflow: 'hidden',
  } as React.CSSProperties,
  topRow:   { display: 'flex', gap: 6, padding: '8px 10px' } as React.CSSProperties,
  quickBtn: {
    flex: 1, height: 26, border: '1px solid #D1D5DB', borderRadius: 5,
    background: '#F9FAFB', fontSize: 12, fontWeight: '600' as const,
    cursor: 'pointer', color: '#374151',
  } as React.CSSProperties,
  divider:  { height: 1, background: '#E5E7EB', margin: '0' } as React.CSSProperties,
  list:     { maxHeight: 220, overflowY: 'auto' as const, padding: '4px 0' },
  item:     {
    display: 'flex', alignItems: 'center', padding: '6px 12px',
    cursor: 'pointer', fontSize: 13, color: '#111827',
    transition: 'background 0.1s',
  } as React.CSSProperties,
};

// ─── Column config (persisted in localStorage) ────────────────────────────────

interface ColConfig {
  order:  string[];  // full column order (all columns including hidden)
  hidden: string[];  // column keys to hide
}

const LS_KEY = 'farm-col-configs';

function loadColConfigs(): Record<string, ColConfig> {
  try { return JSON.parse(localStorage.getItem(LS_KEY) ?? '{}'); }
  catch { return {}; }
}

function applyColConfig(cols: string[], cfg: ColConfig | undefined, defaultHidden: string[] = []): string[] {
  if (!cfg) return cols.filter(c => !defaultHidden.includes(c));
  const ordered = cfg.order.filter(c => cols.includes(c));
  const newCols  = cols.filter(c => !cfg.order.includes(c));
  return [...ordered, ...newCols].filter(c => !cfg.hidden.includes(c));
}

// ─── Password Gate ────────────────────────────────────────────────────────────

const PW = '9824';
const PW_KEY = 'farm-auth';

function PasswordGate({ children }: { children: React.ReactNode }) {
  const [auth, setAuth] = useState(() => localStorage.getItem(PW_KEY) === PW);
  const [input, setInput] = useState('');
  const [shake, setShake] = useState(false);

  const submit = () => {
    if (input === PW) {
      localStorage.setItem(PW_KEY, PW);
      setAuth(true);
    } else {
      setShake(true);
      setInput('');
      setTimeout(() => setShake(false), 600);
    }
  };

  if (auth) return <>{children}</>;

  return (
    <div style={{ minHeight: '100vh', background: '#1A2B4A', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ background: '#FFF', borderRadius: 16, padding: '48px 40px', width: 340, boxShadow: '0 24px 64px rgba(0,0,0,0.4)', textAlign: 'center' }}>
        <div style={{ fontSize: 48, marginBottom: 16 }}>🐄</div>
        <div style={{ fontSize: 22, fontWeight: 800, color: '#1A2B4A', marginBottom: 6 }}>Farm Reports</div>
        <div style={{ fontSize: 14, color: '#6B7280', marginBottom: 32 }}>Bitte Passwort eingeben</div>
        <input
          type="password"
          value={input}
          onChange={e => setInput(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && submit()}
          placeholder="Passwort"
          autoFocus
          style={{
            width: '100%', boxSizing: 'border-box' as const,
            padding: '12px 16px', fontSize: 18, letterSpacing: 6,
            border: `2px solid ${shake ? '#EF4444' : '#E5E7EB'}`,
            borderRadius: 8, outline: 'none', textAlign: 'center',
            animation: shake ? 'shake 0.5s ease' : 'none',
            marginBottom: 16,
          }}
        />
        <button
          onClick={submit}
          style={{ width: '100%', padding: '12px', background: '#1A2B4A', color: '#FFF', border: 'none', borderRadius: 8, fontSize: 15, fontWeight: 700, cursor: 'pointer' }}
        >
          Weiter →
        </button>
      </div>
      <style>{`@keyframes shake { 0%,100%{transform:translateX(0)} 20%{transform:translateX(-8px)} 40%{transform:translateX(8px)} 60%{transform:translateX(-6px)} 80%{transform:translateX(6px)} }`}</style>
    </div>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────

function AppInner() {
  const [mode, setMode]       = useState<'animals' | 'feed'>('animals');
  const [activeId, setActiveId] = useState<string | null>(null);
  const [columns, setColumns] = useState<string[]>([]);
  const [rows, setRows]       = useState<SimpleRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError]     = useState<string | null>(null);
  const [paramsReady, setParamsReady] = useState(false);

  // Params state
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo,   setDateTo]   = useState('');
  const [fDate,    setFDate]    = useState('');
  const [fMin,     setFMin]     = useState('');
  const [fMax,     setFMax]     = useState('');
  const [custom,   setCustom]   = useState<CustomFilter>(defaultCustom);
  const [colConfigs, setColConfigs] = useState<Record<string, ColConfig>>(loadColConfigs);
  const [configOpen, setConfigOpen] = useState(false);

  const activeDef = REPORTS.find(r => r.id === activeId);
  const displayColumns = applyColConfig(columns, activeId ? colConfigs[activeId] : undefined, activeDef?.defaultHidden ?? []);

  const saveColConfig = (reportId: string, cfg: ColConfig) => {
    const next = { ...colConfigs, [reportId]: cfg };
    setColConfigs(next);
    localStorage.setItem(LS_KEY, JSON.stringify(next));
  };

  const needsParams = activeDef &&
    ['daterange','futureage','futureweight','custom'].includes(activeDef.kind);

  const runReport = useCallback(async () => {
    if (!activeDef) return;
    if (activeDef.kind === 'animaledit') return;
    setLoading(true);
    setError(null);
    try {
      let c: string[] = [];
      let r: SimpleRow[] = [];

      if (activeDef.kind === 'pivot') {
        const res = await (activeDef.load!() as Promise<PivotResult>);
        c = res.columns;
        r = res.rows;
      } else if (activeDef.kind === 'simple') {
        r = await (activeDef.load!() as Promise<SimpleRow[]>);
        c = r.length > 0 ? Object.keys(r[0]) : [];
      } else if (activeDef.kind === 'daterange') {
        if (activeDef.id === 'dead')       r = await api.dead(dateFrom, dateTo);
        else if (activeDef.id === 'calving')    r = await api.calving(dateFrom, dateTo);
        else if (activeDef.id === 'not_calving')r = await api.notCalving(dateFrom, dateTo);
        c = r.length > 0 ? Object.keys(r[0]) : [];
      } else if (activeDef.kind === 'futureage') {
        r = await api.futureAge(fDate, Number(fMin) || 0, Number(fMax) || 999);
        c = r.length > 0 ? Object.keys(r[0]) : [];
      } else if (activeDef.kind === 'futureweight') {
        r = await api.futureWeight(fDate, Number(fMin) || 0, Number(fMax) || 9999);
        c = r.length > 0 ? Object.keys(r[0]) : [];
      } else if (activeDef.kind === 'custom') {
        r = await api.custom(buildCustomParams(custom));
        c = r.length > 0 ? Object.keys(r[0]) : [];
      }

      setColumns(c);
      setRows(r);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Unknown error');
    } finally {
      setLoading(false);
    }
  }, [activeDef, dateFrom, dateTo, fDate, fMin, fMax, custom]);

  const selectReport = (id: string) => {
    const def = REPORTS.find(r => r.id === id)!;
    setActiveId(id);
    setColumns([]);
    setRows([]);
    setError(null);
    setParamsReady(false);
    if (!['daterange','futureage','futureweight','custom'].includes(def.kind)) {
      setParamsReady(true);
    }
  };

  useEffect(() => {
    if (paramsReady && activeId) runReport();
  }, [paramsReady, activeId, runReport]);

  return (
    <div style={layout.root}>
      <Sidebar activeId={activeId} onSelect={selectReport} mode={mode} onModeChange={m => { setMode(m); setActiveId(null); setColumns([]); setRows([]); setError(null); setParamsReady(false); }} />
      <div style={layout.main}>
        {!activeDef ? (
          <Welcome />
        ) : activeDef.kind === 'animaledit' ? (
          <AnimalEditPanel />
        ) : (
          <>
            <Header
              title={activeDef.title}
              icon={activeDef.icon}
              rowCount={displayColumns.length > 0 ? rows.length : 0}
              hasData={rows.length > 0}
              onExport={() => void exportXLSX(displayColumns, rows, activeDef.title)}
              onRefresh={paramsReady ? runReport : undefined}
              onConfigCols={columns.length > 0 ? () => setConfigOpen(true) : undefined}
            />
            {needsParams && (
              <ParamsBar
                kind={activeDef.kind}
                dateFrom={dateFrom} setDateFrom={setDateFrom}
                dateTo={dateTo}     setDateTo={setDateTo}
                fDate={fDate}       setFDate={setFDate}
                fMin={fMin}         setFMin={setFMin}
                fMax={fMax}         setFMax={setFMax}
                custom={custom}     setCustom={setCustom}
                onApply={() => {
                  if (!paramsReady) setParamsReady(true);
                  else runReport();
                }}
                label={activeDef.kind === 'futureage' ? 'months' : activeDef.kind === 'futureweight' ? 'kg' : ''}
              />
            )}
            <div style={layout.tableWrap}>
              {loading && <Loader />}
              {!loading && error && <Err msg={error} onRetry={runReport} />}
              {!loading && !error && !paramsReady && (
                <Empty msg="Set the parameters above and click Apply" />
              )}
              {!loading && !error && paramsReady && rows.length === 0 && columns.length === 0 && !loading && (
                <Empty msg="No data found" />
              )}
              {!loading && !error && columns.length > 0 && (
                <DataTable
                  key={activeId}
                  columns={displayColumns}
                  rows={rows}
                  filterCols={activeDef?.filterCols}
                  totalCols={activeDef?.totalCols}
                />
              )}
            </div>
          </>
        )}
      </div>
      {configOpen && activeDef && columns.length > 0 && (
        <ColConfigOverlay
          reportTitle={activeDef.title}
          allColumns={columns}
          config={activeId ? colConfigs[activeId] : undefined}
          onSave={cfg => activeId && saveColConfig(activeId, cfg)}
          onClose={() => setConfigOpen(false)}
        />
      )}
    </div>
  );
}

// ─── Sidebar ──────────────────────────────────────────────────────────────────

function Sidebar({ activeId, onSelect, mode, onModeChange }: {
  activeId: string | null;
  onSelect: (id: string) => void;
  mode: 'animals' | 'feed';
  onModeChange: (m: 'animals' | 'feed') => void;
}) {
  const groups = mode === 'animals' ? ANIMAL_GROUPS : FEED_GROUPS;
  return (
    <div style={sb.wrap}>
      {/* Mode switcher */}
      <div style={sb.tabs}>
        <button
          style={{ ...sb.tab, ...(mode === 'animals' ? sb.tabActive : {}) }}
          onClick={() => onModeChange('animals')}
          title="Animal Reports"
        >
          🐄
        </button>
        <button
          style={{ ...sb.tab, ...(mode === 'feed' ? sb.tabActive : {}) }}
          onClick={() => onModeChange('feed')}
          title="Feed Reports"
        >
          🌾
        </button>
        <span style={sb.brandText}>{mode === 'animals' ? 'Animals' : 'Feed'}</span>
      </div>
      <div style={sb.scroll}>
        {groups.map(g => {
          const items = REPORTS.filter(r => r.group === g);
          return (
            <div key={g}>
              <div style={sb.groupLabel}>{g.toUpperCase()}</div>
              {items.map(r => (
                <button
                  key={r.id}
                  style={{ ...sb.item, ...(activeId === r.id ? sb.itemActive : {}) }}
                  onClick={() => onSelect(r.id)}
                >
                  <span style={sb.itemIcon}>{r.icon}</span>
                  <span style={sb.itemTitle}>{r.title}</span>
                </button>
              ))}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── Header ───────────────────────────────────────────────────────────────────

function Header({
  title, icon, rowCount, hasData, onExport, onRefresh, onConfigCols,
}: {
  title: string; icon: string; rowCount: number; hasData: boolean;
  onExport: () => void; onRefresh?: () => void; onConfigCols?: () => void;
}) {
  return (
    <div style={hdr.bar}>
      <span style={hdr.icon}>{icon}</span>
      <span style={hdr.title}>{title}</span>
      <div style={hdr.actions}>
        {hasData && (
          <span style={hdr.count}>{rowCount.toLocaleString()} rows</span>
        )}
        {onRefresh && (
          <button style={hdr.btnRefresh} onClick={onRefresh} title="Refresh">
            ↺ Refresh
          </button>
        )}
        {onConfigCols && (
          <button style={hdr.btnCols} onClick={onConfigCols} title="Configure columns">
            ⚙ Columns
          </button>
        )}
        {hasData && (
          <button style={hdr.btnExport} onClick={onExport}>
            ↓ Excel Export
          </button>
        )}
      </div>
    </div>
  );
}

// ─── Parameter bars ───────────────────────────────────────────────────────────

interface ParamsBarProps {
  kind: ReportKind;
  dateFrom: string; setDateFrom: (v: string) => void;
  dateTo: string;   setDateTo:   (v: string) => void;
  fDate: string;    setFDate:    (v: string) => void;
  fMin: string;     setFMin:     (v: string) => void;
  fMax: string;     setFMax:     (v: string) => void;
  custom: CustomFilter; setCustom: React.Dispatch<React.SetStateAction<CustomFilter>>;
  onApply: () => void;
  label: string;
}

function ParamsBar(p: ParamsBarProps) {
  return (
    <div style={pb.bar}>
      {p.kind === 'daterange' && (
        <div style={pb.row}>
          <Field label="From" width={140}>
            <input style={pb.input} type="date" onClick={e => (e.currentTarget as HTMLInputElement).showPicker?.()} value={p.dateFrom} onChange={e => p.setDateFrom(e.target.value)} />
          </Field>
          <Field label="To" width={140}>
            <input style={pb.input} type="date" onClick={e => (e.currentTarget as HTMLInputElement).showPicker?.()} value={p.dateTo} onChange={e => p.setDateTo(e.target.value)} />
          </Field>
          <button style={pb.apply} onClick={p.onApply}>Apply</button>
        </div>
      )}
      {(p.kind === 'futureage' || p.kind === 'futureweight') && (
        <div style={pb.row}>
          <Field label="Future date" width={160}>
            <input style={pb.input} type="date" onClick={e => (e.currentTarget as HTMLInputElement).showPicker?.()} value={p.fDate} onChange={e => p.setFDate(e.target.value)} />
          </Field>
          <Field label={`Min ${p.label}`} width={100}>
            <input style={pb.input} type="number" min={0} value={p.fMin} onChange={e => p.setFMin(e.target.value)} placeholder="0" />
          </Field>
          <Field label={`Max ${p.label}`} width={100}>
            <input style={pb.input} type="number" min={0} value={p.fMax} onChange={e => p.setFMax(e.target.value)} placeholder={p.kind === 'futureage' ? '999' : '9999'} />
          </Field>
          <button style={pb.apply} onClick={p.onApply}>Apply</button>
        </div>
      )}
      {p.kind === 'custom' && (
        <CustomFilterBar custom={p.custom} setCustom={p.setCustom} onApply={p.onApply} />
      )}
    </div>
  );
}

function Field({ label, width, children }: { label: string; width: number; children: React.ReactNode }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 4, width }}>
      <span style={pb.label}>{label}</span>
      {children}
    </div>
  );
}

function CustomFilterBar({ custom, setCustom, onApply }: {
  custom: CustomFilter;
  setCustom: React.Dispatch<React.SetStateAction<CustomFilter>>;
  onApply: () => void;
}) {
  const up = (patch: Partial<CustomFilter>) => setCustom(p => ({ ...p, ...patch }));

  // Groups are canonical from the app definition — all shown even if empty in DB
  const ALL_APP_GROUPS = [
    'Parcela 1 - Ferma', 'Parcela 2 - Ferma', 'Parcela 3 - Ferma',
    'Parcela 4 - Ferma', 'Parcela 5 - Ferma', 'Parcela 6 - Ferma',
    'Grupa 1 - Negestante', 'Grupa 2 - Gestatie mica', 'Grupa 3 - Gestatie mare',
    'Grupa 4 - Juninci',
    'Vitei Ingrasare', 'Vitele', 'Intarcati', 'Reforme', 'Insamintate', 'Generala',
  ];

  const [companyOpts, setCompanyOpts] = React.useState<string[]>([]);
  const groupOpts = ALL_APP_GROUPS;

  React.useEffect(() => {
    api.metaCompanies().then(setCompanyOpts).catch(() => {});
  }, []);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      <div style={pb.row}>
        <Field label="Sex" width={188}>
          <div style={{ display: 'flex', gap: 4 }}>
            {(['both','F','M'] as const).map(s => (
              <button key={s} style={custom.sex === s ? pb.toggleOn : pb.toggle}
                onClick={() => up({ sex: s })}>
                {s === 'both' ? 'Both' : s === 'F' ? 'Female' : 'Male'}
              </button>
            ))}
          </div>
        </Field>
        <Field label="Status" width={220}>
          <div style={{ display: 'flex', gap: 4 }}>
            {([['viu','VIU'],['vandut','VANDUT'],['mort','MORT']] as const).map(([k,l]) => (
              <button key={k} style={custom[k] ? pb.toggleOn : pb.toggle}
                onClick={() => up({ [k]: !custom[k] } as Partial<CustomFilter>)}>
                {l}
              </button>
            ))}
          </div>
        </Field>
        <Field label="Age min (m)" width={96}>
          <input style={pb.input} type="number" min={0} value={custom.minAge}
            onChange={e => up({ minAge: e.target.value })} placeholder="0" />
        </Field>
        <Field label="Age max (m)" width={96}>
          <input style={pb.input} type="number" min={0} value={custom.maxAge}
            onChange={e => up({ maxAge: e.target.value })} placeholder="999" />
        </Field>
        <Field label="Weight min (kg)" width={110}>
          <input style={pb.input} type="number" min={0} value={custom.minKg}
            onChange={e => up({ minKg: e.target.value })} placeholder="0" />
        </Field>
        <Field label="Weight max (kg)" width={110}>
          <input style={pb.input} type="number" min={0} value={custom.maxKg}
            onChange={e => up({ maxKg: e.target.value })} placeholder="9999" />
        </Field>
        <Field label="Company" width={220}>
          <MultiSelect
            options={companyOpts}
            value={custom.companies}
            onChange={v => up({ companies: v })}
            placeholder="All companies"
          />
        </Field>
        <Field label="Group" width={220}>
          <MultiSelect
            options={groupOpts}
            value={custom.groups}
            onChange={v => up({ groups: v })}
            placeholder="All groups"
          />
        </Field>
        <button style={{ ...pb.apply, alignSelf: 'flex-end' }} onClick={onApply}>Apply</button>
      </div>
    </div>
  );
}

// ─── Animal View / Edit Panel ─────────────────────────────────────────────────

const EDIT_GROUPS = [
  'Parcela 1 - Ferma', 'Parcela 2 - Ferma', 'Parcela 3 - Ferma',
  'Parcela 4 - Ferma', 'Parcela 5 - Ferma', 'Parcela 6 - Ferma',
  'Grupa 1 - Negestante', 'Grupa 2 - Gestatie mica', 'Grupa 3 - Gestatie mare',
  'Grupa 4 - Juninci',
  'Vitei Ingrasare', 'Vitele', 'Intarcati', 'Reforme', 'Insamintate', 'Generala',
];

function AeLabel({ text }: { text: string }) {
  return <div style={ae.label}>{text}</div>;
}

function AeChips({ options, value, onChange, labels }: {
  options: string[]; value: string;
  onChange?: (v: string) => void; labels?: Record<string, string>;
}) {
  return (
    <div style={{ display: 'flex', gap: 4, flexWrap: 'wrap' as const }}>
      {options.map(o => (
        <button key={o}
          style={value === o ? ae.chipOn : ae.chip}
          onClick={onChange ? () => onChange(o) : undefined}
          disabled={!onChange}
        >{labels?.[o] ?? o}</button>
      ))}
    </div>
  );
}

function AeField({ label, children, style }: { label: string; children: React.ReactNode; style?: React.CSSProperties }) {
  return <div style={style}><div style={ae.label}>{label}</div>{children}</div>;
}

function Val({ children, date }: { children?: React.ReactNode; date?: boolean }) {
  const s = children == null || children === '' ? '—' : children;
  return <div style={ae.val}>{date && typeof s === 'string' && s !== '—' ? formatDate(s) : s}</div>;
}

function AeFlag({ label, value, onChange }: { label: string; value: boolean; onChange?: (v: boolean) => void }) {
  if (onChange) return (
    <button style={value ? ae.toggleOn : ae.toggle} onClick={() => onChange(!value)}>
      {label}: {value ? 'ON' : 'OFF'}
    </button>
  );
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 5, fontSize: 13 }}>
      <span style={{ fontWeight: 700, color: value ? C.green : '#CBD5E1' }}>{value ? '✓' : '✗'}</span>
      <span style={{ color: value ? '#111827' : '#9CA3AF', fontWeight: value ? 600 : 400 }}>{label}</span>
    </span>
  );
}

function Txt({ v, date: isDate }: { v: string; date?: boolean }) {
  return <div style={ae.val}>{v ? (isDate ? formatDate(v) : v) : '—'}</div>;
}
function Inp({ v, set, type = 'text' }: { v: string; set: (x: string) => void; type?: string }) {
  return <input style={ae.inp} type={type} value={v} onChange={e => set(e.target.value)} />;
}

function AnimalEditPanel() {
  const [last4,       setLast4]       = useState('');
  const [searching,   setSearching]   = useState(false);
  const [candidates,  setCandidates]  = useState<SimpleRow[] | null>(null);
  const [selTag,      setSelTag]      = useState<string | null>(null);
  const [animal,      setAnimal]      = useState<Record<string, unknown> | null>(null);
  const [snapshot,    setSnapshot]    = useState<Record<string, unknown> | null>(null);
  const [loadingAnim, setLoadingAnim] = useState(false);
  const [tab,         setTab]         = useState<'view' | 'edit'>('view');
  const [saving,      setSaving]      = useState(false);
  const [saveMsg,     setSaveMsg]     = useState<{ ok: boolean; text: string } | null>(null);
  const [clientOpts,  setClientOpts]  = useState<string[]>([]);

  React.useEffect(() => { api.metaClients().then(setClientOpts).catch(() => {}); }, []);

  // Edit fields
  const [fTag,    setFTag]    = useState('');
  const [fComp,   setFComp]   = useState('');
  const [fBirth,  setFBirth]  = useState('');
  const [fSex,    setFSex]    = useState('');
  const [fColor,  setFColor]  = useState('');
  const [fPurp,   setFPurp]   = useState('');
  const [fStat,   setFStat]   = useState('');
  const [fBidaa,  setFBidaa]  = useState(false);
  const [fPed,    setFPed]    = useState(false);
  const [fBreed,  setFBreed]  = useState(false);
  const [fGest,   setFGest]   = useState(false);
  const [fGestWk, setFGestWk] = useState('');
  const [fBull,   setFBull]   = useState(false);
  const [fDDat,   setFDDat]   = useState('');
  const [fDReas,  setFDReas]  = useState('');
  const [fSDat,   setFSDat]   = useState('');
  const [fCli,    setFCli]    = useState('');
  const [fGrp,    setFGrp]    = useState('');
  const [fMom,    setFMom]    = useState('');
  const [fDad,    setFDad]    = useState('');
  const [fCmt,    setFCmt]    = useState('');

  const fill = (a: Record<string, unknown>) => {
    setFTag(str(a.ear_tag));
    setFComp(str(a.company));
    setFBirth(str(a.birth_date).slice(0, 10));
    setFSex(str(a.sex));
    setFColor(normColorIn(str(a.color)));
    setFPurp(str(a.purpose));
    setFStat(str(a.status));
    setFBidaa(!!a.bidaa);
    const p = a.pedigree;
    setFPed(p === true || ['yes', 'true', 'YES', 'TRUE', 'True'].includes(String(p)));
    setFBreed(!!a.breeding);
    setFGest(!!a.gestation);
    setFGestWk(a.gestation_weeks != null ? str(a.gestation_weeks) : '');
    setFBull(!!a.is_bull);
    setFDDat(str(a.death_date).slice(0, 10));
    setFDReas(str(a.death_reason));
    setFSDat(str(a.sales_date).slice(0, 10));
    setFCli(str(a.client_name));
    setFGrp(str(a.group_name));
    setFMom(str(a.mother_ear_tag));
    setFDad(str(a.father_ear_tag));
    setFCmt(str(a.comment));
  };

  const loadAnimal = (tag: string) => {
    setSelTag(tag);
    setLoadingAnim(true);
    setAnimal(null);
    setSaveMsg(null);
    api.animalByTag(tag)
      .then(a => { setAnimal(a); setSnapshot(a); fill(a); setTab('view'); })
      .catch(() => {})
      .finally(() => setLoadingAnim(false));
  };

  const doSearch = async () => {
    if (!/^\d{4}$/.test(last4)) return;
    setSearching(true);
    setCandidates(null);
    setAnimal(null);
    setSelTag(null);
    setSaveMsg(null);
    try {
      const res = await api.animalSearch(last4);
      if (res.length === 1) {
        loadAnimal(str(res[0].ear_tag));
      } else {
        setCandidates(res);
      }
    } catch { setCandidates([]); }
    finally { setSearching(false); }
  };

  const validateTag = (tag: string): string | null => {
    const m = tag.match(/^([A-Z]{2})(\d+)$/);
    if (!m) return 'must start with 2 letters followed by digits only';
    const [, country, digits] = m;
    if (country === 'DE' && digits.length !== 10) return `DE tags need exactly 10 digits (got ${digits.length})`;
    if (country === 'RO' && digits.length !== 12) return `RO tags need exactly 12 digits (got ${digits.length})`;
    if (country !== 'DE' && country !== 'RO' && (digits.length < 10 || digits.length > 12))
      return 'must have 10–12 digits after the country code';
    return null;
  };

  const validate = (): string | null => {
    if (!fComp)  return 'Company: required';
    if (!fSex)   return 'Sex: required';
    if (!fBirth) return 'Birth Date: required';
    if (!fStat)  return 'Status: required';
    const tagErr = fTag ? validateTag(fTag) : null;
    if (tagErr) return `Animal Ear Tag: ${tagErr}`;
    const momErr = fMom ? validateTag(fMom) : null;
    if (momErr) return `Mother Ear Tag: ${momErr}`;
    const dadErr = fDad ? validateTag(fDad) : null;
    if (dadErr) return `Father Ear Tag: ${dadErr}`;
    if (fStat === 'MORT'   && !fDDat) return 'Death Date: required when status is MORT';
    if (fStat === 'VANDUT' && !fSDat) return 'Sale Date: required when status is VANDUT';
    return null;
  };

  const doSave = async () => {
    if (!selTag) return;
    const validErr = validate();
    if (validErr) { setSaveMsg({ ok: false, text: validErr }); return; }
    setSaving(true);
    setSaveMsg(null);
    try {
      const body: Record<string, unknown> = {
        company: fComp, birth_date: fBirth || null, sex: fSex,
        color: fColor, purpose: fPurp || null, status: fStat,
        bidaa: fBidaa, pedigree: fPed, breeding: fBreed,
        gestation: fGest,
        gestation_weeks: fGest && fGestWk ? Number(fGestWk) : null,
        is_bull: fBull,
        death_date:   fStat === 'MORT'   ? fDDat || null  : null,
        death_reason: fStat === 'MORT'   ? fDReas || null : null,
        sales_date:   fStat === 'VANDUT' ? fSDat || null  : null,
        client_name:  fStat === 'VANDUT' ? fCli || null   : null,
        group_name: fGrp || null,
        mother_ear_tag: fMom || null,
        father_ear_tag: fDad || null,
        comment: fCmt || null,
      };
      if (fTag && fTag !== selTag) body.new_ear_tag = fTag;
      await api.animalUpdate(selTag, body);
      const newTag = (fTag && fTag !== selTag) ? fTag : selTag;
      loadAnimal(newTag);
      setSaveMsg({ ok: true, text: '✓ Saved' });
    } catch (e: unknown) {
      setSaveMsg({ ok: false, text: e instanceof Error ? e.message : 'Save failed' });
    } finally { setSaving(false); }
  };

  const isEdit = tab === 'edit';

  return (
    <div style={ae.wrap}>

      {/* ── No animal: centered search ── */}
      {!animal && !loadingAnim && (
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column' as const, alignItems: 'center', justifyContent: 'center', gap: 20 }}>
          <div style={{ fontSize: 22, fontWeight: 700, color: '#1A2B4A' }}>Search for an animal</div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <input
              style={ae.searchInp}
              value={last4}
              onChange={e => { if (/^\d{0,4}$/.test(e.target.value)) setLast4(e.target.value); }}
              onKeyDown={e => e.key === 'Enter' && doSearch()}
              maxLength={4}
              autoFocus
            />
            <button style={ae.searchBtn} onClick={doSearch} disabled={last4.length !== 4 || searching}>
              {searching ? '…' : 'Search'}
            </button>
          </div>
          {candidates !== null && (
            candidates.length === 0
              ? <div style={{ color: '#9CA3AF', fontSize: 13 }}>No animal found.</div>
              : <div style={{ display: 'flex', flexDirection: 'column' as const, gap: 8, alignItems: 'center' }}>
                  <span style={{ fontSize: 12, color: '#6B7280' }}>Multiple results — select one:</span>
                  {candidates.map(c => (
                    <button key={str(c.ear_tag)} style={{ ...ae.candBtn, width: 340, textAlign: 'center' as const }}
                      onClick={() => { setCandidates(null); loadAnimal(str(c.ear_tag)); }}>
                      {str(c.ear_tag)} · {str(c.status)} · {str(c.company)}
                    </button>
                  ))}
                </div>
          )}
        </div>
      )}

      {loadingAnim && (
        <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#9CA3AF', fontSize: 14 }}>
          Loading…
        </div>
      )}

      {animal && (
        <>
          {/* ── Top bar: ← Search | [View][Edit] | 💾 Save  ↺ Discard | saveMsg ── */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
            <button style={ae.backBtn} onClick={() => { setAnimal(null); setSelTag(null); setCandidates(null); setSaveMsg(null); setLast4(''); }}>
              ← Search
            </button>
            <div style={{ width: 1, background: '#E5E7EB', alignSelf: 'stretch', margin: '0 4px' }} />
            <button style={!isEdit ? ae.tabOn : ae.tabOff} onClick={() => setTab('view')}>View</button>
            <button style={isEdit  ? ae.tabOn : ae.tabOff} onClick={() => setTab('edit')}>Edit</button>
            {isEdit && <>
              <button style={ae.saveBtn} onClick={doSave} disabled={saving}>
                {saving ? 'Saving…' : '💾 Save'}
              </button>
              <button style={ae.discardBtn} onClick={() => { if (snapshot) fill(snapshot); }}>
                ↺ Discard
              </button>
            </>}
            {saveMsg && (
              <span style={{ marginLeft: 6, fontSize: 13, fontWeight: 700, color: saveMsg.ok ? C.green : '#C62828' }}>
                {saveMsg.text}
              </span>
            )}
            <span style={{ marginLeft: 'auto', fontSize: 11, color: '#9CA3AF', fontFamily: 'monospace' }}>
              {selTag}
            </span>
          </div>

          {/* ── Single centered column ── */}
          <div style={{ flex: 1, overflowY: 'auto' as const, display: 'flex', flexDirection: 'column' as const, alignItems: 'center', padding: '12px 0 20px' }}>
            <div style={{ width: 480, display: 'flex', flexDirection: 'column' as const, gap: 0 }}>

              <div style={ae.field}><AeLabel text="Ear Tag" />
                {isEdit ? <Inp v={fTag} set={setFTag} /> : <Txt v={fTag} />}
              </div>
              <div style={ae.field}><AeLabel text="Company" />
                <AeChips options={['Apollo','Ares','AFM','Atlas']} value={fComp} onChange={isEdit ? setFComp : undefined} />
              </div>
              <div style={ae.field}><AeLabel text="Birth Date" />
                {isEdit
                  ? <input style={ae.inp} type="date" onClick={e => (e.currentTarget as HTMLInputElement).showPicker?.()} value={fBirth} onChange={e => setFBirth(e.target.value)} />
                  : <Txt v={fBirth} date />}
              </div>
              <div style={ae.field}><AeLabel text="Sex" />
                <AeChips options={['M','F']} value={fSex} onChange={isEdit ? setFSex : undefined} />
              </div>
              <div style={ae.field}><AeLabel text="Color" />
                <AeChips options={['BLACK','RED']} value={fColor} onChange={isEdit ? setFColor : undefined}
                  labels={{ BLACK: 'Black', RED: 'Red' }} />
              </div>
              <div style={ae.field}><AeLabel text="Purpose" />
                <AeChips options={['reproductie','ingrasare']} value={fPurp} onChange={isEdit ? setFPurp : undefined}
                  labels={{ reproductie: 'Repro', ingrasare: 'Fatten' }} />
              </div>
              <div style={ae.field}><AeLabel text="Status" />
                <AeChips options={['VIU','MORT','VANDUT']} value={fStat} onChange={isEdit ? setFStat : undefined} />
              </div>
              {fStat === 'MORT' && (<>
                <div style={ae.field}><AeLabel text="Death Date" />
                  {isEdit
                    ? <input style={ae.inp} type="date" onClick={e => (e.currentTarget as HTMLInputElement).showPicker?.()} value={fDDat} onChange={e => setFDDat(e.target.value)} />
                    : <Txt v={fDDat} date />}
                </div>
                <div style={ae.field}><AeLabel text="Death Reason" />
                  {isEdit ? <Inp v={fDReas} set={setFDReas} /> : <Txt v={fDReas} />}
                </div>
              </>)}
              {fStat === 'VANDUT' && (<>
                <div style={ae.field}><AeLabel text="Sale Date" />
                  {isEdit
                    ? <input style={ae.inp} type="date" onClick={e => (e.currentTarget as HTMLInputElement).showPicker?.()} value={fSDat} onChange={e => setFSDat(e.target.value)} />
                    : <Txt v={fSDat} date />}
                </div>
                <div style={ae.field}><AeLabel text="Client" />
                  {isEdit
                    ? <select style={ae.sel} value={fCli} onChange={e => setFCli(e.target.value)}>
                        <option value="">— Select client —</option>
                        {clientOpts.map(c => <option key={c} value={c}>{c}</option>)}
                        {fCli && !clientOpts.includes(fCli) && <option value={fCli}>{fCli}</option>}
                      </select>
                    : <Txt v={fCli} />}
                </div>
              </>)}

              <div style={{ borderTop: '1px solid #E5E7EB', margin: '10px 0 10px' }} />

              <div style={ae.field}><AeLabel text="Group" />
                {isEdit
                  ? <select style={ae.sel} value={fGrp} onChange={e => setFGrp(e.target.value)}>
                      <option value="">— None —</option>
                      {EDIT_GROUPS.map(g => <option key={g} value={g}>{g}</option>)}
                    </select>
                  : <Txt v={fGrp} />}
              </div>
              <div style={ae.field}><AeLabel text="Mother Ear Tag" />
                {isEdit ? <Inp v={fMom} set={setFMom} /> : <Txt v={fMom} />}
              </div>
              <div style={ae.field}><AeLabel text="Father Ear Tag" />
                {isEdit ? <Inp v={fDad} set={setFDad} /> : <Txt v={fDad} />}
              </div>

              <div style={{ borderTop: '1px solid #E5E7EB', margin: '10px 0 10px' }} />

              <div style={ae.field}>
                <div style={{ fontSize: 10, fontWeight: 700, color: '#9CA3AF', textTransform: 'uppercase' as const, letterSpacing: 0.8, marginBottom: 7 }}>Flags</div>
                <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' as const }}>
                  <AeFlag label="Bidaa"    value={fBidaa} onChange={isEdit ? setFBidaa : undefined} />
                  <AeFlag label="Pedigree" value={fPed}   onChange={isEdit ? setFPed   : undefined} />
                  {fSex !== 'M' && <AeFlag label="Breeding" value={fBreed} onChange={isEdit ? setFBreed : undefined} />}
                  {fSex !== 'M' && <AeFlag label="Pregnant" value={fGest}  onChange={isEdit ? setFGest  : undefined} />}
                  {fSex !== 'F' && <AeFlag label="Bull"     value={fBull}  onChange={isEdit ? setFBull  : undefined} />}
                </div>
              </div>
              {fGest && (
                <div style={ae.field}><AeLabel text="Gestation weeks" />
                  {isEdit ? <Inp v={fGestWk} set={setFGestWk} type="number" /> : <Txt v={fGestWk} />}
                </div>
              )}

              <div style={{ borderTop: '1px solid #E5E7EB', margin: '10px 0 10px' }} />

              <div style={ae.field}><AeLabel text="Comment" />
                {isEdit
                  ? <textarea
                      rows={6}
                      style={{ ...ae.inp, height: 'auto', resize: 'none' as const, fontFamily: 'inherit', paddingTop: 6 }}
                      value={fCmt}
                      onChange={e => setFCmt(e.target.value)}
                    />
                  : <div style={{ ...ae.val, whiteSpace: 'pre-wrap' as const, wordBreak: 'break-word' as const }}>
                      {fCmt || '—'}
                    </div>
                }
              </div>

            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ─── Column filter dropdown (Excel-style checkboxes) ─────────────────────────

// selected: null = alle angezeigt (kein Filter), Set<string> = nur diese Werte anzeigen
function ColFilterDropdown({ col: _col, options, selected, onChange }: {
  col: string;
  options: string[];
  selected: Set<string> | null;
  onChange: (sel: Set<string> | null) => void;
}) {
  const [open, setOpen] = React.useState(false);
  const [pos,  setPos]  = React.useState({ top: 0, left: 0 });
  const triggerRef = React.useRef<HTMLButtonElement>(null);
  const panelRef   = React.useRef<HTMLDivElement>(null);
  const allSelected = selected === null;

  React.useEffect(() => {
    if (!open) return;
    const handler = (e: MouseEvent) => {
      const t = e.target as Node;
      if (
        triggerRef.current && !triggerRef.current.contains(t) &&
        panelRef.current   && !panelRef.current.contains(t)
      ) setOpen(false);
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [open]);

  const handleOpen = () => {
    if (!open && triggerRef.current) {
      const r = triggerRef.current.getBoundingClientRect();
      setPos({ top: r.bottom + 2, left: r.left });
    }
    setOpen(o => !o);
  };

  const toggle = (val: string) => {
    if (allSelected) {
      // War "alle" → deselektiere diesen einen Wert
      onChange(new Set(options.filter(o => o !== val)));
    } else {
      const next = new Set(selected!);
      next.has(val) ? next.delete(val) : next.add(val);
      // Wenn wieder alle drin → zurück zu null (kein Filter)
      onChange(next.size === options.length ? null : next);
    }
  };

  const isChecked   = (val: string) => allSelected || (selected?.has(val) ?? false);
  const activeCount = allSelected ? options.length : (selected?.size ?? 0);
  const hasFilter   = !allSelected;

  const panel = open && ReactDOM.createPortal(
    <div ref={panelRef} style={{
      position: 'fixed', top: pos.top, left: pos.left, zIndex: 9999,
      background: '#FFF', border: '1px solid #D1D5DB', borderRadius: 6,
      boxShadow: '0 4px 16px rgba(0,0,0,0.15)', minWidth: 180, maxWidth: 260,
    }}>
      <div style={{ padding: '6px 8px', borderBottom: '1px solid #E5E7EB', display: 'flex', gap: 6 }}>
        <button onClick={() => onChange(null)}           style={flt.quickBtn}>Alle</button>
        <button onClick={() => onChange(new Set())}      style={flt.quickBtn}>Keine</button>
      </div>
      <div style={{ maxHeight: 220, overflowY: 'auto', padding: '4px 0' }}>
        {options.map(val => (
          <label key={val} style={flt.item}>
            <input
              type="checkbox"
              checked={isChecked(val)}
              onChange={() => toggle(val)}
              style={{ marginRight: 7, accentColor: '#1565C0', cursor: 'pointer' }}
            />
            <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{val}</span>
          </label>
        ))}
      </div>
    </div>,
    document.body
  );

  return (
    <div style={{ position: 'relative', marginTop: 3 }}>
      <button
        ref={triggerRef}
        onClick={handleOpen}
        style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          width: '100%', height: 22, padding: '0 5px',
          fontSize: 11, border: 'none', borderRadius: 3,
          background: hasFilter ? '#FFF9C4' : '#FFFFFF',
          color: '#111827', cursor: 'pointer', fontFamily: 'inherit',
          fontWeight: hasFilter ? 700 : 400,
        }}
      >
        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', flex: 1, textAlign: 'left' }}>
          {allSelected ? 'Alle' : `${activeCount} / ${options.length}`}
        </span>
        <span style={{ marginLeft: 4, flexShrink: 0 }}>▾</span>
      </button>
      {panel}
    </div>
  );
}

const flt = {
  quickBtn: { flex: 1, height: 24, border: '1px solid #D1D5DB', borderRadius: 4, background: '#F9FAFB', fontSize: 11, fontWeight: '600' as const, cursor: 'pointer', color: '#374151' } as React.CSSProperties,
  item: { display: 'flex', alignItems: 'center', padding: '5px 10px', cursor: 'pointer', fontSize: 12, color: '#111827' } as React.CSSProperties,
};

// ─── Data table ───────────────────────────────────────────────────────────────

function DataTable({ columns, rows, filterCols = [], totalCols = [] }: {
  columns: string[];
  rows: SimpleRow[];
  filterCols?: string[];
  totalCols?: string[];
}) {
  // null = kein Filter (alle), Set<string> = nur diese Werte anzeigen
  const [filters, setFilters] = React.useState<Record<string, Set<string> | null>>({});

  const filterOptions = React.useMemo(() => {
    const opts: Record<string, string[]> = {};
    for (const col of filterCols) {
      const vals = [...new Set(rows.map(r => String(r[col] ?? '')).filter(Boolean))].sort();
      opts[col] = vals;
    }
    return opts;
  }, [rows, filterCols]);

  const filteredRows = React.useMemo(() =>
    rows.filter(row =>
      filterCols.every(col => {
        const sel = filters[col];
        if (sel === null || sel === undefined) return true;  // kein Filter
        if (sel.size === 0) return false;                    // nichts ausgewählt
        return sel.has(String(row[col] ?? ''));
      })
    ),
  [rows, filters, filterCols]);

  const totals = React.useMemo(() => {
    const t: Record<string, number> = {};
    for (const col of totalCols) {
      t[col] = filteredRows.reduce((s, r) => s + (Number(r[col]) || 0), 0);
    }
    return t;
  }, [filteredRows, totalCols]);

  const hasFilters = filterCols.length > 0;
  const hasTotals  = totalCols.length > 0 && filteredRows.length > 0;

  const setColFilter = (col: string, sel: Set<string> | null) =>
    setFilters(f => ({ ...f, [col]: sel }));

  return (
    <div style={tbl.outer}>
      <table style={tbl.table}>
        <thead>
          <tr>
            {columns.map((col, i) => (
              <th key={col} style={{
                ...tbl.th,
                minWidth: colMinWidth(col),
                ...(i === 0 ? tbl.stickyCol : {}),
                ...(i === 0 ? tbl.stickyTh : {}),
                verticalAlign: 'top',
              }}>
                <div style={{ marginBottom: hasFilters ? 1 : 0 }}>{formatColHeader(col)}</div>
                {filterCols.includes(col) && (
                  <ColFilterDropdown
                    col={col}
                    options={filterOptions[col] ?? []}
                    selected={col in filters ? filters[col] : null}
                    onChange={sel => setColFilter(col, sel)}
                  />
                )}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {filteredRows.map((row, ri) => (
            <tr key={ri} style={ri % 2 === 0 ? tbl.rowEven : tbl.rowOdd}>
              {columns.map((col, ci) => (
                <td key={col} style={{
                  ...tbl.td,
                  minWidth: colMinWidth(col),
                  ...(ci === 0 ? tbl.stickyCol : {}),
                  ...(ci === 0 ? (ri % 2 === 0 ? tbl.stickyEven : tbl.stickyOdd) : {}),
                }}>
                  {formatCell(row[col], col)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
        {hasTotals && (
          <tfoot>
            <tr>
              {columns.map((col, ci) => {
                const isTotal = totalCols.includes(col);
                return (
                  <td key={col} style={{
                    ...tbl.td,
                    minWidth: colMinWidth(col),
                    ...(ci === 0 ? tbl.stickyCol : {}),
                    background: '#E8F0FE',
                    fontWeight: 700,
                    borderTop: '2px solid #1565C0',
                    color: isTotal ? '#1565C0' : '#374151',
                  }}>
                    {ci === 0
                      ? `Total (${filteredRows.length})`
                      : isTotal
                        ? col === 'count'
                          ? String(Math.round(totals[col]))
                          : totals[col].toFixed(2) + ' t'
                        : ''}
                  </td>
                );
              })}
            </tr>
          </tfoot>
        )}
      </table>
    </div>
  );
}

// ─── Misc components ──────────────────────────────────────────────────────────

function Welcome() {
  return (
    <div style={misc.center}>
      <div style={misc.welcomeIcon}>🐄</div>
      <div style={misc.welcomeTitle}>Farm Reports</div>
      <div style={misc.welcomeDesc}>Select a report from the sidebar to get started.</div>
    </div>
  );
}

function Loader() {
  return (
    <div style={misc.center}>
      <div style={misc.spinner} />
      <div style={{ marginTop: 16, color: '#6B7280' }}>Loading…</div>
    </div>
  );
}

function Err({ msg, onRetry }: { msg: string; onRetry: () => void }) {
  return (
    <div style={misc.center}>
      <div style={{ color: '#C62828', marginBottom: 12, maxWidth: 480, textAlign: 'center' }}>
        Error: {msg}
      </div>
      <button style={misc.retryBtn} onClick={onRetry}>Retry</button>
    </div>
  );
}

function Empty({ msg }: { msg: string }) {
  return (
    <div style={misc.center}>
      <div style={{ color: '#9CA3AF', fontSize: 15 }}>{msg}</div>
    </div>
  );
}

// ─── Styles (inline, no external CSS needed) ──────────────────────────────────

const layout = {
  root:      { display: 'flex', height: '100vh', overflow: 'hidden', fontFamily: "'Inter', system-ui, sans-serif" } as React.CSSProperties,
  main:      { flex: 1, display: 'flex', flexDirection: 'column' as const, overflow: 'hidden', background: '#F8FAFC' },
  tableWrap: { flex: 1, overflow: 'auto', position: 'relative' as const },
};

const C = { navy: '#1A2B4A', blue: '#1565C0', green: '#2E7D32', blueLt: '#EFF6FF', blueAlt: '#F0F4FF' };

const sb = {
  wrap:      { width: 248, background: C.navy, display: 'flex', flexDirection: 'column' as const, overflow: 'hidden', flexShrink: 0 } as React.CSSProperties,
  tabs:      { display: 'flex', alignItems: 'center', gap: 6, padding: '12px 14px', borderBottom: '1px solid rgba(255,255,255,0.08)', flexShrink: 0 } as React.CSSProperties,
  tab:       { width: 38, height: 38, borderRadius: 8, border: '2px solid rgba(255,255,255,0.15)', background: 'rgba(255,255,255,0.06)', fontSize: 20, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, transition: 'all 0.15s' } as React.CSSProperties,
  tabActive: { background: 'rgba(255,255,255,0.22)', borderColor: 'rgba(255,255,255,0.5)' } as React.CSSProperties,
  brandText: { fontSize: 14, fontWeight: '700', color: '#FFFFFF', letterSpacing: 0.2, marginLeft: 4 } as React.CSSProperties,
  scroll:    { flex: 1, overflowY: 'auto' as const, padding: '8px 0 24px' },
  groupLabel:{ fontSize: 10, fontWeight: '700', color: 'rgba(255,255,255,0.35)', letterSpacing: 1.2, padding: '14px 16px 4px', textTransform: 'uppercase' as const },
  item:      { display: 'flex', alignItems: 'center', gap: 9, width: '100%', padding: '8px 16px', background: 'none', border: 'none', cursor: 'pointer', textAlign: 'left' as const, borderRadius: 0, color: 'rgba(255,255,255,0.75)', fontSize: 13, fontWeight: '500', transition: 'background 0.1s' },
  itemActive:{ background: 'rgba(255,255,255,0.12)', color: '#FFFFFF' },
  itemIcon:  { fontSize: 15, flexShrink: 0 },
  itemTitle: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' as const },
};

const hdr = {
  bar:      { height: 52, background: C.blue, display: 'flex', alignItems: 'center', gap: 10, paddingInline: 20, flexShrink: 0, boxShadow: '0 2px 6px rgba(0,0,0,0.18)' } as React.CSSProperties,
  icon:     { fontSize: 20 } as React.CSSProperties,
  title:    { flex: 1, fontSize: 16, fontWeight: '700', color: '#FFFFFF', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' as const } as React.CSSProperties,
  actions:  { display: 'flex', alignItems: 'center', gap: 10 } as React.CSSProperties,
  count:    { fontSize: 13, color: 'rgba(255,255,255,0.7)' } as React.CSSProperties,
  btnRefresh:{ background: 'rgba(255,255,255,0.15)', border: '1px solid rgba(255,255,255,0.3)', borderRadius: 6, color: '#FFFFFF', fontSize: 13, fontWeight: '600', padding: '5px 14px', cursor: 'pointer' } as React.CSSProperties,
  btnCols:   { background: 'rgba(255,255,255,0.15)', border: '1px solid rgba(255,255,255,0.3)', borderRadius: 6, color: '#FFFFFF', fontSize: 13, fontWeight: '600', padding: '5px 14px', cursor: 'pointer' } as React.CSSProperties,
  btnExport: { background: C.green, border: 'none', borderRadius: 6, color: '#FFFFFF', fontSize: 13, fontWeight: '600', padding: '6px 16px', cursor: 'pointer' } as React.CSSProperties,
};

const pb = {
  bar:    { background: '#FFFFFF', borderBottom: '1px solid #E5E7EB', padding: '10px 20px', flexShrink: 0, overflowX: 'auto' as const } as React.CSSProperties,
  row:    { display: 'flex', alignItems: 'flex-end', gap: 12, flexWrap: 'nowrap' as const, minWidth: 'max-content' },
  label:  { fontSize: 10, fontWeight: '700', color: '#6B7280', textTransform: 'uppercase' as const, letterSpacing: 0.8 } as React.CSSProperties,
  input:  { height: 32, border: '1px solid #D1D5DB', borderRadius: 6, padding: '0 10px', fontSize: 13, color: '#111827', width: '100%', outline: 'none', fontFamily: 'inherit' } as React.CSSProperties,
  apply:  { height: 32, background: C.blue, border: 'none', borderRadius: 6, color: '#FFF', fontSize: 13, fontWeight: '700', padding: '0 20px', cursor: 'pointer', flexShrink: 0 } as React.CSSProperties,
  toggle: { height: 28, border: '1px solid #D1D5DB', borderRadius: 5, background: '#FFF', color: '#374151', fontSize: 12, fontWeight: '600', padding: '0 10px', cursor: 'pointer' } as React.CSSProperties,
  toggleOn:{ height: 28, border: '1px solid ' + C.blue, borderRadius: 5, background: C.blue, color: '#FFF', fontSize: 12, fontWeight: '600', padding: '0 10px', cursor: 'pointer' } as React.CSSProperties,
};

const ae = {
  wrap:       { flex: 1, display: 'flex', flexDirection: 'column' as const, padding: '14px 20px', gap: 10, overflow: 'hidden', background: '#F8FAFC' } as React.CSSProperties,
  searchRow:  { display: 'flex', alignItems: 'center', gap: 10, flexShrink: 0 } as React.CSSProperties,
  searchLabel:{ fontSize: 13, fontWeight: '600', color: '#374151' } as React.CSSProperties,
  searchInp:  { height: 36, width: 96, border: '2px solid ' + C.blue, borderRadius: 6, padding: '0 8px', fontSize: 20, fontWeight: '700', textAlign: 'center' as const, letterSpacing: 7, color: '#111827', outline: 'none', fontFamily: 'inherit' } as React.CSSProperties,
  searchBtn:  { height: 36, background: C.blue, border: 'none', borderRadius: 6, color: '#FFF', fontSize: 13, fontWeight: '700', padding: '0 18px', cursor: 'pointer', flexShrink: 0 } as React.CSSProperties,
  candBtn:    { border: '1px solid #D1D5DB', borderRadius: 6, padding: '4px 12px', fontSize: 12, cursor: 'pointer', background: '#FFF', color: '#111827', fontFamily: 'inherit', fontWeight: '600' as const } as React.CSSProperties,
  tabRow:     { display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 } as React.CSSProperties,
  tabOn:      { height: 30, border: '1px solid ' + C.blue, borderRadius: 6, padding: '0 18px', fontSize: 13, cursor: 'pointer', background: C.blue, color: '#FFF', fontFamily: 'inherit', fontWeight: '700' as const } as React.CSSProperties,
  tabOff:     { height: 30, border: '1px solid #D1D5DB', borderRadius: 6, padding: '0 18px', fontSize: 13, cursor: 'pointer', background: '#FFF', color: '#374151', fontFamily: 'inherit', fontWeight: '600' as const } as React.CSSProperties,
  saveBtn:    { height: 30, background: C.green, border: 'none', borderRadius: 6, color: '#FFF', fontSize: 13, fontWeight: '700' as const, padding: '0 20px', cursor: 'pointer' } as React.CSSProperties,
  discardBtn: { height: 30, background: '#FFF', border: '1px solid #D1D5DB', borderRadius: 6, color: '#374151', fontSize: 13, fontWeight: '600' as const, padding: '0 16px', cursor: 'pointer' } as React.CSSProperties,
  backBtn:    { height: 30, background: '#FFF', border: '1px solid #D1D5DB', borderRadius: 6, color: '#374151', fontSize: 13, fontWeight: '600' as const, padding: '0 14px', cursor: 'pointer' } as React.CSSProperties,
  grid:       { flex: 1, display: 'grid', gridTemplateColumns: '1fr 1fr', overflow: 'hidden', minHeight: 0 } as React.CSSProperties,
  col:        { display: 'flex', flexDirection: 'column' as const, overflow: 'hidden', paddingRight: 20 },
  field:      { marginBottom: 7 } as React.CSSProperties,
  label:      { fontSize: 10, fontWeight: '700' as const, color: '#6B7280', textTransform: 'uppercase' as const, letterSpacing: 0.8, marginBottom: 3 } as React.CSSProperties,
  val:        { fontSize: 13, color: '#111827', padding: '3px 0' } as React.CSSProperties,
  inp:        { height: 28, border: '1px solid #D1D5DB', borderRadius: 5, padding: '0 8px', fontSize: 13, color: '#111827', outline: 'none', width: '100%', boxSizing: 'border-box' as const, fontFamily: 'inherit', background: '#FFF' } as React.CSSProperties,
  sel:        { height: 28, border: '1px solid #D1D5DB', borderRadius: 5, padding: '0 6px', fontSize: 13, color: '#111827', outline: 'none', width: '100%', boxSizing: 'border-box' as const, fontFamily: 'inherit', background: '#FFF', cursor: 'pointer' } as React.CSSProperties,
  chip:       { height: 26, border: '1px solid #D1D5DB', borderRadius: 5, padding: '0 10px', fontSize: 12, fontWeight: '600' as const, cursor: 'pointer', background: '#FFF', color: '#374151', fontFamily: 'inherit' } as React.CSSProperties,
  chipOn:     { height: 26, border: '1px solid ' + C.blue, borderRadius: 5, padding: '0 10px', fontSize: 12, fontWeight: '600' as const, cursor: 'pointer', background: C.blue, color: '#FFF', fontFamily: 'inherit' } as React.CSSProperties,
  toggle:     { height: 24, border: '1px solid #D1D5DB', borderRadius: 4, padding: '0 8px', fontSize: 11, fontWeight: '600' as const, cursor: 'pointer', background: '#F9FAFB', color: '#374151', fontFamily: 'inherit' } as React.CSSProperties,
  toggleOn:   { height: 24, border: '1px solid ' + C.green, borderRadius: 4, padding: '0 8px', fontSize: 11, fontWeight: '600' as const, cursor: 'pointer', background: C.green, color: '#FFF', fontFamily: 'inherit' } as React.CSSProperties,
};

const tbl = {
  outer:      { overflow: 'auto', height: '100%' } as React.CSSProperties,
  table:      { borderCollapse: 'collapse' as const, width: '100%', tableLayout: 'auto' as const, fontSize: 13 },
  th:         { position: 'sticky' as const, top: 0, background: C.blue, color: '#FFF', fontWeight: '600', fontSize: 12, padding: '8px 10px', whiteSpace: 'nowrap' as const, textAlign: 'left' as const, borderRight: '1px solid rgba(255,255,255,0.15)', zIndex: 2 },
  td:         { padding: '6px 10px', borderBottom: '1px solid #E5E7EB', borderRight: '1px solid #F3F4F6', whiteSpace: 'nowrap' as const, color: '#111827' },
  rowEven:    { background: '#FFFFFF' } as React.CSSProperties,
  rowOdd:     { background: '#F5F8FF' } as React.CSSProperties,
  stickyCol:  { position: 'sticky' as const, left: 0, zIndex: 1 },
  stickyTh:   { zIndex: 3, borderRight: '2px solid rgba(255,255,255,0.3)' },
  stickyEven: { background: '#FFFFFF' } as React.CSSProperties,
  stickyOdd:  { background: '#F5F8FF' } as React.CSSProperties,
};

const misc = {
  center:      { display: 'flex', flexDirection: 'column' as const, alignItems: 'center', justifyContent: 'center', height: '100%', padding: 40 },
  welcomeIcon: { fontSize: 64, marginBottom: 20 } as React.CSSProperties,
  welcomeTitle:{ fontSize: 24, fontWeight: '700', color: '#1A2B4A', marginBottom: 10 } as React.CSSProperties,
  welcomeDesc: { fontSize: 15, color: '#6B7280', textAlign: 'center' as const, maxWidth: 360 } as React.CSSProperties,
  spinner:     { width: 36, height: 36, border: '3px solid #E5E7EB', borderTopColor: C.blue, borderRadius: '50%', animation: 'spin 0.8s linear infinite' } as React.CSSProperties,
  retryBtn:    { background: C.blue, border: 'none', borderRadius: 6, color: '#FFF', fontSize: 13, fontWeight: '700', padding: '8px 20px', cursor: 'pointer' } as React.CSSProperties,
};

// ─── Column Configurator Overlay ──────────────────────────────────────────────

function ColConfigOverlay({
  reportTitle, allColumns, config, onSave, onClose,
}: {
  reportTitle: string;
  allColumns:  string[];
  config:      ColConfig | undefined;
  onSave:      (cfg: ColConfig) => void;
  onClose:     () => void;
}) {
  const initOrder = React.useMemo(() => {
    if (config?.order?.length) {
      const known = config.order.filter(c => allColumns.includes(c));
      const fresh = allColumns.filter(c => !config.order.includes(c));
      return [...known, ...fresh];
    }
    return [...allColumns];
  }, []);  // eslint-disable-line react-hooks/exhaustive-deps

  const [items,   setItems]   = React.useState<string[]>(initOrder);
  const [hidden,  setHidden]  = React.useState<Set<string>>(new Set(config?.hidden ?? []));
  const [dragIdx, setDragIdx] = React.useState<number | null>(null);

  const toggleHidden = (col: string) =>
    setHidden(h => { const n = new Set(h); n.has(col) ? n.delete(col) : n.add(col); return n; });

  const handleDragStart = (idx: number) => setDragIdx(idx);

  const handleDragOver = (e: React.DragEvent, idx: number) => {
    e.preventDefault();
    if (dragIdx === null || dragIdx === idx) return;
    const next = [...items];
    const [moved] = next.splice(dragIdx, 1);
    next.splice(idx, 0, moved);
    setItems(next);
    setDragIdx(idx);
  };

  const handleDragEnd = () => setDragIdx(null);

  const handleSave = () => {
    onSave({ order: items, hidden: Array.from(hidden) });
    onClose();
  };

  const visibleCount = items.length - hidden.size;

  return ReactDOM.createPortal(
    <div style={ov.backdrop} onClick={onClose}>
      <div style={ov.panel} onClick={e => e.stopPropagation()}>

        {/* Header */}
        <div style={ov.header}>
          <span style={{ fontWeight: 700, fontSize: 15 }}>⚙ Configure Columns</span>
          <span style={{ fontSize: 13, color: 'rgba(255,255,255,0.65)', flex: 1, marginLeft: 12, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' as const }}>
            {reportTitle} &nbsp;·&nbsp; {visibleCount} of {items.length} visible
          </span>
          <button style={ov.closeBtn} onClick={onClose}>✕</button>
        </div>

        {/* Hint */}
        <div style={ov.hint}>
          Drag ⠿ to reorder &nbsp;·&nbsp; uncheck to hide a column
        </div>

        {/* Column list */}
        <div style={ov.list}>
          {items.map((col, idx) => (
            <div
              key={col}
              draggable
              onDragStart={() => handleDragStart(idx)}
              onDragOver={e => handleDragOver(e, idx)}
              onDragEnd={handleDragEnd}
              style={{
                ...ov.item,
                opacity: dragIdx === idx ? 0.4 : 1,
                background: hidden.has(col) ? '#F9FAFB' : '#FFFFFF',
              }}
            >
              <span style={ov.handle}>⠿</span>
              <input
                type="checkbox"
                checked={!hidden.has(col)}
                onChange={() => toggleHidden(col)}
                style={{ width: 16, height: 16, marginRight: 12, accentColor: C.blue, cursor: 'pointer', flexShrink: 0 }}
              />
              <span style={{ flex: 1, fontSize: 14, color: hidden.has(col) ? '#9CA3AF' : '#111827' }}>
                {formatColHeader(col)}
              </span>
              <span style={{ fontSize: 11, color: '#D1D5DB', fontFamily: 'monospace', marginLeft: 12 }}>
                {col}
              </span>
            </div>
          ))}
        </div>

        {/* Footer */}
        <div style={ov.footer}>
          <button style={ov.btnReset} onClick={() => { setItems([...allColumns]); setHidden(new Set()); }}>
            ↺ Reset to default
          </button>
          <div style={{ display: 'flex', gap: 10 }}>
            <button style={ov.btnCancel} onClick={onClose}>Cancel</button>
            <button style={ov.btnSave}   onClick={handleSave}>Apply</button>
          </div>
        </div>
      </div>
    </div>,
    document.body,
  );
}

const ov = {
  backdrop: { position: 'fixed' as const, inset: 0, background: 'rgba(0,0,0,0.55)', zIndex: 2000, display: 'flex', alignItems: 'center', justifyContent: 'center' },
  panel:    { background: '#FFF', borderRadius: 12, width: 660, maxHeight: '88vh', display: 'flex', flexDirection: 'column' as const, overflow: 'hidden', boxShadow: '0 24px 64px rgba(0,0,0,0.35)' },
  header:   { background: C.blue, padding: '15px 20px', display: 'flex', alignItems: 'center', gap: 6, flexShrink: 0 } as React.CSSProperties,
  hint:     { padding: '7px 20px', background: '#F8FAFC', borderBottom: '1px solid #E5E7EB', fontSize: 12, color: '#6B7280', flexShrink: 0 } as React.CSSProperties,
  list:     { flex: 1, overflowY: 'auto' as const },
  item:     { display: 'flex', alignItems: 'center', padding: '9px 20px', borderBottom: '1px solid #F3F4F6', cursor: 'grab', userSelect: 'none' as const, transition: 'background 0.08s' },
  handle:   { fontSize: 18, color: '#CBD5E1', marginRight: 12, cursor: 'grab', flexShrink: 0, letterSpacing: -2 } as React.CSSProperties,
  footer:   { padding: '14px 20px', borderTop: '1px solid #E5E7EB', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0 } as React.CSSProperties,
  closeBtn: { background: 'none', border: 'none', color: '#FFF', fontSize: 18, cursor: 'pointer', padding: '0 4px', lineHeight: 1, flexShrink: 0 } as React.CSSProperties,
  btnReset: { background: 'none', border: '1px solid #D1D5DB', borderRadius: 6, padding: '7px 14px', fontSize: 13, color: '#374151', cursor: 'pointer' } as React.CSSProperties,
  btnCancel:{ background: 'none', border: '1px solid #D1D5DB', borderRadius: 6, padding: '7px 16px', fontSize: 13, color: '#374151', cursor: 'pointer' } as React.CSSProperties,
  btnSave:  { background: C.blue, border: 'none', borderRadius: 6, padding: '7px 22px', fontSize: 13, fontWeight: '700' as const, color: '#FFF', cursor: 'pointer' } as React.CSSProperties,
};

export default function App() {
  return <PasswordGate><AppInner /></PasswordGate>;
}
