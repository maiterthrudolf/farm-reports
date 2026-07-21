const BASE = import.meta.env.VITE_API_BASE ?? 'http://localhost:8080';

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
  allMothers:        () => getList(`${BASE}/api/reports/mothers`),
  births:            () => getList(`${BASE}/api/reports/births`),
  closeBirths:       () => getList(`${BASE}/api/reports/births/close`),
  aiPrepared:        () => getList(`${BASE}/api/reports/ai/prepared`),
  aiInseminations:   () => getList(`${BASE}/api/reports/ai/inseminations`),
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
  metaClients:   () => fetch(`${BASE}/api/reports/meta/clients`).then(r => r.json()) as Promise<string[]>,
  cronStatus:    () => fetch(`${BASE}/api/reports/cron-status`).then(r => r.json()) as Promise<{ job_name: string; run_at: string; status: string; animals_updated: number | null; detail: string | null }[]>,

  // Pivot reports
  weights:           () => getPivot(`${BASE}/api/reports/weights`),
  pregnancyHistory:  () => getPivot(`${BASE}/api/reports/pregnancy`),
  medicalGroup:      () => getPivot(`${BASE}/api/reports/medical/group`),
  medicalIndividual: () => getPivot(`${BASE}/api/reports/medical/individual`),
  aiHistory:         () => getPivot(`${BASE}/api/reports/ai`),

  // Feed – Bales
  baleProduction: () => getList(`${BASE}/api/reports/bales/production`).then(rows =>
    rows.map(r => ({
      bale_type:       r.bale_type,
      area:            r.area,
      year:            r.year,
      production_date: r.production_date,
      count:           r.count,
      bale_weight_kg:  r.weight_kg != null ? Math.round(Number(r.weight_kg)) : null,
      total_t:         r.total_kg  != null ? (Number(r.total_kg) / 1000).toFixed(2) : null,
      created_time:    r.created_time,
    }))
  ),
  baleFarmEntry: () => getList(`${BASE}/api/reports/bales/farm-entry`).then(rows =>
    rows.map(r => ({
      bale_type:      r.bale_type,
      year:           r.year,
      entry_date:     r.entry_date,
      count:          r.count,
      bale_weight_kg: r.weight_kg != null ? Math.round(Number(r.weight_kg)) : null,
      total_t:        r.total_kg  != null ? (Number(r.total_kg) / 1000).toFixed(2) : null,
      created_time:   r.created_time,
    }))
  ),
  balePurchase: () => getList(`${BASE}/api/reports/bales/purchase`).then(rows =>
    rows.map(r => ({
      bale_type:      r.bale_type,
      year:           r.year,
      purchase_date:  r.purchase_date,
      count:          r.count,
      bale_weight_kg: r.weight_kg != null ? Math.round(Number(r.weight_kg)) : null,
      total_t:        r.total_kg  != null ? (Number(r.total_kg) / 1000).toFixed(2) : null,
      created_time:   r.created_time,
    }))
  ),
  baleConsumption: () => getList(`${BASE}/api/reports/bales/consumption`).then(rows =>
    rows.map(r => ({
      bale_type:      r.bale_type,
      location:       r.location,
      year:           r.year,
      entry_date:     r.entry_date,
      count:          r.count,
      bale_weight_kg: r.weight_kg != null ? Math.round(Number(r.weight_kg)) : null,
      total_t:        r.total_kg  != null ? (Number(r.total_kg) / 1000).toFixed(2) : null,
      created_time:   r.created_time,
    }))
  ),

  // Feed – Cereals
  cerealProduction: () => getList(`${BASE}/api/reports/cereals/production`).then(rows =>
    rows.map(r => ({
      cereal_type:     r.cereal_type,
      area:            r.area,
      year:            r.year,
      production_date: r.production_date,
      harvest_t:       r.harvest_kg != null ? (Number(r.harvest_kg) / 1000).toFixed(2) : null,
      created_time:    r.created_time,
    }))
  ),
  cerealPurchase: () => getList(`${BASE}/api/reports/cereals/purchase`).then(rows =>
    rows.map(r => ({
      cereal_type:   r.cereal_type,
      supplier:      r.supplier,
      year:          r.year,
      purchase_date: r.purchase_date,
      purchase_t:    r.purchase_kg != null ? (Number(r.purchase_kg) / 1000).toFixed(2) : null,
      created_time:  r.created_time,
    }))
  ),
  cerealConsumption: () => getList(`${BASE}/api/reports/cereals/consumption`).then(rows =>
    rows.map(r => ({
      cereal_type:      r.cereal_type,
      year:             r.year,
      consumption_date: r.consumption_date,
      consumption_t:    r.consumption_kg != null ? (Number(r.consumption_kg) / 1000).toFixed(2) : null,
      created_time:     r.created_time,
    }))
  ),
  cerealSale: () => getList(`${BASE}/api/reports/cereals/sale`).then(rows =>
    rows.map(r => ({
      cereal_type:  r.cereal_type,
      client:       r.client,
      year:         r.year,
      sale_date:    r.sale_date,
      sale_t:       r.sale_kg != null ? (Number(r.sale_kg) / 1000).toFixed(2) : null,
      created_time: r.created_time,
    }))
  ),

  animalSearch: (last4: string) =>
    getList(`${BASE}/api/cattle/search?last4=${last4}`),

  animalByTag: async (earTag: string): Promise<Record<string, unknown>> => {
    const r = await fetch(`${BASE}/api/cattle/by-tag/${encodeURIComponent(earTag)}`);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return r.json();
  },

  animalUpdate: async (earTag: string, body: Record<string, unknown>): Promise<void> => {
    const r = await fetch(`${BASE}/api/edit/${encodeURIComponent(earTag)}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!r.ok) {
      const d = await r.json().catch(() => null);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const detail = (d as any)?.detail;
      let msg: string;
      if (Array.isArray(detail) && detail.length > 0) {
        const first = detail[0];
        const loc: unknown[] = Array.isArray(first?.loc) ? first.loc : [];
        const field = loc.length > 0 ? String(loc[loc.length - 1]) : '';
        const errMsg: string = first?.msg ?? 'Validation error';
        msg = field ? `${field}: ${errMsg}` : errMsg;
      } else {
        msg = typeof detail === 'string' ? detail : `HTTP ${r.status}`;
      }
      throw new Error(msg);
    }
  },
};
