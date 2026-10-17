// Solver's other two engines, for problems that are not linear. GRG
// Nonlinear, as a local method: the constraints folded into an augmented
// Lagrangian (the method of multipliers), each inner problem minimised by
// BFGS on forward-difference gradients with a backtracking line search, the
// variables kept within their bounds by projection; from where the cells
// are, to the nearest local optimum, as Excel's GRG goes. Evolutionary: a
// differential evolution (DE/rand/1/bin) inside the variables' bounds, which
// it needs on every one, as Excel's does; a feasible point beats an
// infeasible one, two infeasible ones by how far each is out; integer
// variables rounded; seeded, so the same problem gives the same answer.
// Pure: the problem's values come from `evaluate`, which the command does
// by setting the sheet's cells and reading its formulas.
export interface Constraint { readonly rel: "<=" | ">=" | "="; }
export interface Nlp {
  readonly n: number;
  readonly maximise: boolean;
  readonly rels: readonly ("<=" | ">=" | "=")[];
  // The objective, and each constraint's lhs - rhs, at x.
  evaluate(x: readonly number[]): Promise<{ readonly obj: number; readonly diff: readonly number[] }>;
  readonly lower: readonly number[];      // -Infinity when none
  readonly upper: readonly number[];      // Infinity when none
  readonly integer?: readonly boolean[];
  readonly start: readonly number[];
}
export type NlpResult = { readonly status: "optimal" | "converged"; readonly x: number[]; readonly obj: number; readonly evaluations: number } | { readonly status: "infeasible"; readonly x: number[]; readonly evaluations: number };

const clamp = (v: number, lo: number, hi: number) => Math.min(hi, Math.max(lo, v));
// How far each constraint is out (0 when it holds).
const outBy = (rels: readonly string[], diff: readonly number[]) => diff.map((d, i) => (rels[i] === "<=" ? Math.max(0, d) : rels[i] === ">=" ? Math.max(0, -d) : Math.abs(d)));
const total = (xs: readonly number[]) => xs.reduce((s, v) => s + v, 0);

export async function grg(p: Nlp, tol = 1e-7): Promise<NlpResult> {
  let evals = 0;
  const sign = p.maximise ? -1 : 1;
  const proj = (x: readonly number[]) => x.map((v, i) => clamp(v, p.lower[i], p.upper[i]));
  const f = async (x: readonly number[]) => { evals++; const e = await p.evaluate(x); return { obj: sign * e.obj, diff: e.diff }; };
  // The augmented Lagrangian at x for multipliers lam and penalty rho.
  const lag = (o: { obj: number; diff: readonly number[] }, lam: readonly number[], rho: number) => o.obj + o.diff.reduce((s, d, i) => {
    if (p.rels[i] === "=") return s + lam[i] * d + rho / 2 * d * d;
    const g = p.rels[i] === "<=" ? d : -d, t = Math.max(0, lam[i] + rho * g);
    return s + (t * t - lam[i] * lam[i]) / (2 * rho);
  }, 0);
  let x = proj(p.start), lam = p.rels.map(() => 0), rho = 10;
  let cur = await f(x), was: number[] | null = null;
  for (let outer = 0; outer < 40; outer++) {
    // BFGS on L(x) for these multipliers.
    const L = async (y: readonly number[]) => { const o = await f(y); return { v: lag(o, lam, rho), o }; };
    const grad = async (y: readonly number[], fy: number) => {
      const g: number[] = [];
      for (let j = 0; j < p.n; j++) {
        const h = 1e-6 * Math.max(1, Math.abs(y[j]));
        const up = y[j] + h <= p.upper[j];
        const z = [...y]; z[j] = up ? y[j] + h : y[j] - h;
        const fz = (await L(z)).v;
        g.push(up ? (fz - fy) / h : (fy - fz) / h);
      }
      return g;
    };
    let H = Array.from({ length: p.n }, (_, i) => Array.from({ length: p.n }, (_, j) => (i === j ? 1 : 0)));
    let here = await L(x), g = await grad(x, here.v);
    for (let it = 0; it < 500; it++) {
      // A component pressed against its bound stays put.
      const free = g.map((gj, j) => !((x[j] <= p.lower[j] && gj > 0) || (x[j] >= p.upper[j] && gj < 0)));
      const d = H.map((row) => -row.reduce((s, v, j) => s + v * (free[j] ? g[j] : 0), 0)).map((v, j) => (free[j] ? v : 0));
      let slope = d.reduce((s, v, j) => s + v * g[j], 0);
      if (slope >= 0) { H = H.map((row, i) => row.map((_, j) => (i === j ? 1 : 0))); continue; }
      let t = 1, next = here, y = x;
      for (let k = 0; k < 40; k++) {
        y = proj(x.map((v, j) => v + t * d[j]));
        next = await L(y);
        if (next.v <= here.v + 1e-4 * t * slope) break;
        t /= 2;
      }
      const s = y.map((v, j) => v - x[j]);
      if (Math.max(...s.map(Math.abs)) < tol * (1 + Math.max(...x.map(Math.abs)))) { x = y; here = next; break; }
      const g2 = await grad(y, next.v), q = g2.map((v, j) => v - g[j]);
      const sq = s.reduce((a, v, j) => a + v * q[j], 0);
      if (sq > 1e-12) {
        const Hq = H.map((row) => row.reduce((a, v, j) => a + v * q[j], 0)), qHq = q.reduce((a, v, j) => a + v * Hq[j], 0);
        H = H.map((row, i) => row.map((v, j) => v + ((sq + qHq) * s[i] * s[j]) / (sq * sq) - (Hq[i] * s[j] + s[i] * Hq[j]) / sq));
      }
      const gn = Math.sqrt(g2.reduce((a, v, j) => a + (((y[j] <= p.lower[j] && v > 0) || (y[j] >= p.upper[j] && v < 0)) ? 0 : v * v), 0));
      const done = gn < 1e-7 * (1 + Math.abs(next.v)) || Math.abs(next.v - here.v) < 1e-14 * (1 + Math.abs(here.v));
      x = y; here = next; g = g2;
      if (done) break;
    }
    cur = here.o;
    const out = outBy(p.rels, cur.diff);
    // Feasible, and where the last round left it: the multipliers settled.
    const still = was !== null && x.every((v, j) => Math.abs(v - was![j]) < 1e-6 * (1 + Math.abs(v)));
    if (total(out) < 1e-6 && (p.rels.length === 0 || still)) return { status: "optimal", x, obj: sign * cur.obj, evaluations: evals };
    was = [...x];
    lam = lam.map((l, i) => (p.rels[i] === "=" ? l + rho * cur.diff[i] : Math.max(0, l + rho * (p.rels[i] === "<=" ? cur.diff[i] : -cur.diff[i]))));
    rho = Math.min(rho * 4, 1e8);
  }
  return total(outBy(p.rels, cur.diff)) < 1e-4 ? { status: "converged", x, obj: sign * cur.obj, evaluations: evals } : { status: "infeasible", x, evaluations: evals };
}

// A seeded generator (mulberry32): the same run every time.
function random(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

export async function evolutionary(p: Nlp, generations = 150, seed = 1): Promise<NlpResult> {
  let evals = 0;
  const rnd = random(seed), np = Math.max(15, 8 * p.n), F = 0.7, CR = 0.9;
  const fix = (x: number[]) => x.map((v, i) => { const c = clamp(v, p.lower[i], p.upper[i]); return p.integer?.[i] ? Math.round(c) : c; });
  const score = async (x: number[]) => { evals++; const e = await p.evaluate(x); return { x, obj: (p.maximise ? -1 : 1) * e.obj, out: total(outBy(p.rels, e.diff)) }; };
  // Feasible first; then the better objective, or the smaller violation.
  const better = (a: { obj: number; out: number }, b: { obj: number; out: number }) => (a.out < 1e-9 && b.out < 1e-9 ? a.obj <= b.obj : a.out < b.out);
  const pop = [await score(fix([...p.start]))];
  while (pop.length < np) pop.push(await score(fix(p.lower.map((lo, i) => lo + rnd() * (p.upper[i] - lo)))));
  for (let gen = 0; gen < generations; gen++) {
    for (let i = 0; i < np; i++) {
      let a = i, b = i, c = i;
      while (a === i) a = Math.floor(rnd() * np);
      while (b === i || b === a) b = Math.floor(rnd() * np);
      while (c === i || c === a || c === b) c = Math.floor(rnd() * np);
      const jr = Math.floor(rnd() * p.n);
      const trial = fix(pop[i].x.map((v, j) => (j === jr || rnd() < CR ? pop[a].x[j] + F * (pop[b].x[j] - pop[c].x[j]) : v)));
      const t = await score(trial);
      if (better(t, pop[i])) pop[i] = t;
    }
  }
  const best = pop.reduce((m, q) => (better(q, m) ? q : m));
  return best.out < 1e-6 ? { status: "converged", x: best.x, obj: (p.maximise ? -1 : 1) * best.obj, evaluations: evals } : { status: "infeasible", x: best.x, evaluations: evals };
}
