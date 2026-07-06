const BASE = import.meta.env.VITE_API_BASE ?? 'http://localhost:8000';

async function getList(url: string): Promise<Record<string, unknown>[]> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);
  return res.json();
}

async function getPivot(url: string): Promise<{ columns: string[]; rows: Record<string, unknown>[] }> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${res.statusText}`);
  return res.json();
}

export type SimpleRow = Record<string, unknown>;
export type PivotResult = { columns: string[]; rows: SimpleRow[] };

// Backend paths (from routers/reports.py):
//   /animals               → all animals
//   /weights               → weights pivot
//   /pregnancy             → pregnancy history pivot
//   /pregnancy/current     → currently pregnant
//   /pregnancy/lost        → lost pregnancy
//   /medical/group         → group treatments pivot
//   /medical/individual    → individual treatments pivot
//   /medical/simple        → medical summary
//   /births/pending        → pending father
//   /animals/by-age-sex    → by sex+age filter (params: sex, min_months, max_months)
//   /animals/dead          → dead in period (params: date_from, date_to)
//   /animals/calving       → calvings in period
//   /animals/not-calving   → expected but no birth
//   /animals/future-age    → future age (params: future_date, min_months, max_months)
//   /animals/future-weight → future weight (params: future_date, min_kg, max_kg)
//   /animals/custom        → custom filter (params: sex, status, min_age_months, max_age_months,
//                                           company, group_name, is_bull, min_weight_kg, max_weight_kg)
//   /ai                    → AI history pivot

export const api = {
  allAnimals:        () => getList(`${BASE}/api/reports/animals`),
  currentPregnant:   () => getList(`${BASE}/api/reports/pregnancy/current`),
  medicalSimple:     () => getList(`${BASE}/api/reports/medical/simple`),
  lostPregnancy:     () => getList(`${BASE}/api/reports/pregnancy/lost`),
  pendingBirths:     () => getList(`${BASE}/api/reports/births/pending`),

  bySex: (sex: string, min: number, max: number) =>
    getList(`${BASE}/api/reports/animals/by-age-sex?sex=${sex}&min_months=${min}&max_months=${max}`),

  dead: (from: string, to: string) =>
    getList(`${BASE}/api/reports/animals/dead?date_from=${from}&date_to=${to}`),

  calving: (from: string, to: string) =>
    getList(`${BASE}/api/reports/animals/calving?date_from=${from}&date_to=${to}`),

  notCalving: (from: string, to: string) =>
    getList(`${BASE}/api/reports/animals/not-calving?date_from=${from}&date_to=${to}`),

  futureAge: (date: string, min: number, max: number) =>
    getList(`${BASE}/api/reports/animals/future-age?future_date=${date}&min_months=${min}&max_months=${max}`),

  futureWeight: (date: string, min: number, max: number) =>
    getList(`${BASE}/api/reports/animals/future-weight?future_date=${date}&min_kg=${min}&max_kg=${max}`),

  custom: (params: Record<string, string>) =>
    getList(`${BASE}/api/reports/animals/custom?${new URLSearchParams(params)}`),

  // Preset shortcuts via custom endpoint
  bulls:       () => getList(`${BASE}/api/reports/animals/custom?sex=M&is_bull=true`),
  males16plus: () => getList(`${BASE}/api/reports/animals/custom?sex=M&min_age_months=16&status=VIU`),
  males1216:   () => getList(`${BASE}/api/reports/animals/custom?sex=M&min_age_months=12&max_age_months=16`),
  youngMales:  () => getList(`${BASE}/api/reports/animals/custom?sex=M&min_age_months=6&max_age_months=12`),
  maleCalves:  () => getList(`${BASE}/api/reports/animals/custom?sex=M&max_age_months=12`),

  // Meta (available filter values)
  metaCompanies: () => fetch(`${BASE}/api/reports/meta/companies`).then(r => r.json()) as Promise<string[]>,
  metaGroups:    () => fetch(`${BASE}/api/reports/meta/groups`).then(r => r.json()) as Promise<string[]>,

  // Pivot reports
  weights:           () => getPivot(`${BASE}/api/reports/weights`),
  pregnancyHistory:  () => getPivot(`${BASE}/api/reports/pregnancy`),
  medicalGroup:      () => getPivot(`${BASE}/api/reports/medical/group`),
  medicalIndividual: () => getPivot(`${BASE}/api/reports/medical/individual`),
  aiHistory:         () => getPivot(`${BASE}/api/reports/ai`),
};
