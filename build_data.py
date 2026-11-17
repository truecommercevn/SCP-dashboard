#!/usr/bin/env python3
"""SCP report Excel -> DATA object per team lead (same shape the dashboard page already uses).

Inputs
  oneonones.xlsx  sheet "Manager Levels"  (OneOnOnes_*.xlsx): reporting chain of every employee
  scp.xlsx        sheet "Data"            (User Initiated 1-on-1s.xlsx): one row per weekly SCP

Rules (documented so numbers can be audited):
  Team          = everyone whose Manager Level 1..6 contains the team lead ("Last, First (id)").
  Headcount     = team members that are Active and appear in the SCP data (Vietnam only).
  Submission    = a row with Employee Status Shared/Finalized.   Rating = "Employee Rating" 1-5.
  Month bucket  = Employee Shared Date.   Recent = last 90 days to as-of; prior = the 90 days before.
  Completion    = submitted / (submitted + Not Started + Started) for rows launched in the recent window.
  Supervisor    = last non-empty Manager Level of the employee.
  Themes        = keyword matching on the free text (approximate; keyword lists are in themes.py, built from the real comments).
"""
import os, re
from collections import Counter, defaultdict
import pandas as pd

from themes import CHAL, SUCC, NONE_C, NONE_S, comp   # keyword lists live in themes.py (edit there)
THEMES = comp(CHAL)
THEMES_S = comp(SUCC)
ID = re.compile(r'^(.*?),\s*(.*?)\s*\((\w+)\)\s*$')


def disp(s):                     # "Nguyen, Amy (001212)" -> "Amy Nguyen"
    m = ID.match(str(s))
    return f'{m.group(2)} {m.group(1)}' if m else str(s)


def mid(s):
    m = ID.match(str(s)); return m.group(3) if m else None


def fdate(d): return f'{d:%b} {d.day}, {d.year}'


def themes_of(texts, T=None):
    c = Counter()
    for t in texts:
        for k, rx in (T or THEMES).items():
            if rx.search(t): c[k] += 1
    return c


def stats(g, asof):
    """g = SCP rows of a group; returns the block used in dept / mgr tables."""
    sub = g[g.okf]; rec = sub[(sub.sd > asof - pd.Timedelta(days=90))]; pri = sub[(sub.sd <= asof - pd.Timedelta(days=90)) & (sub.sd > asof - pd.Timedelta(days=180))]
    r = lambda x: round(float(x.r.mean()), 2) if x.r.notna().any() else None
    ar, ap = r(rec), r(pri)
    due = g[(g.ld > asof - pd.Timedelta(days=90)) & (g.ld <= asof)]
    return {'submissions': int(len(sub)), 'completion': round(float(due.okf.mean()), 3) if len(due) else None,
            'avg_all': r(sub), 'avg_recent': ar, 'avg_prior': ap,
            'delta': round(ar - ap, 2) if ar is not None and ap is not None else None,
            'low_recent': int((rec.r <= 2).sum()), 'high_recent': int((rec.r >= 4).sum())}


def titles(emp_path):
    """email (lower) -> job Title from the shared employee list."""
    if not emp_path or not os.path.exists(emp_path): return {}
    df = pd.read_excel(emp_path, sheet_name='Main sheet', dtype=str).dropna(subset=['Email', 'Title'])
    return {str(e).strip().lower(): str(t).strip() for e, t in zip(df['Email'], df['Title'])}


def hire_dates(emp_path):
    """email (lower) -> Original Hire Date, from the shared employee list (sheet 'Main sheet')."""
    if not emp_path or not os.path.exists(emp_path): return {}
    df = pd.read_excel(emp_path, sheet_name='Main sheet', dtype=str)
    df['hd'] = pd.to_datetime(df['Original Hire Date'], format='mixed', errors='coerce')
    df = df.dropna(subset=['Email', 'hd'])
    return {str(e).strip().lower(): h for e, h in zip(df['Email'], df['hd'])}


def week_stats(sub_launch, hire, asof, weeks=12):
    """Weekly completion over the last `weeks` COMPLETE calendar weeks (Mon-Sun) before the week of the data date
    (the current week is still open, so it is not counted).
    due = weeks in which the person was already employed; ok = weeks in which a check-in was submitted (by the date it was shared, the same date used for the monthly counts).
    A week with no submitted check-in counts as missed (the export has no row for a skipped week)."""
    cur = asof.to_period('W'); got = set(sub_launch.dropna().dt.to_period('W')); due = ok = 0
    for k in range(1, weeks + 1):
        w = cur - k
        if pd.notna(hire) and hire >= w.end_time: continue
        due += 1; ok += int(w in got)
    return ok, due


def warn(msg):
    print(f'::warning::{msg}')          # shows as a yellow annotation on the GitHub Actions run


def _tokens(x): return sorted(re.findall(r"[a-z0-9']+", str(x).lower()))


def find_manager(view_name, mgr_counts):
    """mgr_counts: {'Nguyen, Anh (001810)': number of reports}. Returns the employee id of the manager, or None."""
    want = _tokens(view_name); exact, partial = [], []
    for m, n in mgr_counts.items():
        mm = ID.match(m)
        if not mm: continue
        have = _tokens(mm.group(1) + ' ' + mm.group(2))
        if have == want: exact.append((n, mm.group(3), m))
        elif set(want) <= set(have): partial.append((n, mm.group(3), m))
    pick = exact or (partial if len(partial) == 1 else [])
    if len(pick) > 1: warn(f'More than one manager is called "{view_name}" ({", ".join(p[2] for p in pick)}); using the one with most reports.')
    return max(pick)[1] if pick else None


def build_teams(ml_path, scp_path, team_names, emp_path=None):
    ml = pd.read_excel(ml_path, sheet_name='Manager Levels', dtype=str)
    lv = [c for c in ml.columns if c.startswith('Manager Level')]
    ml['EmployeeId'] = ml['EmployeeId'].astype(str)
    mgr_counts = Counter(str(x) for _, r in ml.iterrows() for x in r[lv] if isinstance(x, str))
    paths = scp_path if isinstance(scp_path, (list, tuple)) else [scp_path]       # one file per export: read all, oldest first
    frames = [pd.read_excel(q, sheet_name='Data', dtype=str) for q in paths]
    u = pd.concat(frames, ignore_index=True)
    # the same check-in can appear in several exports: keep the newest export's version
    u = u.drop_duplicates(subset=['Employee ID', 'Launch Date'], keep='last').reset_index(drop=True)
    u['id'] = u['Employee ID'].astype(str)
    for c, n in (('Launch Date', 'ld'), ('Employee Shared Date', 'sd')): u[n] = pd.to_datetime(u[c], format='mixed', errors='coerce')
    u['r'] = pd.to_numeric(u['Employee Rating'], errors='coerce')
    u['okf'] = u['Employee Status'].isin(['Shared', 'Finalized'])
    # some rows (e.g. check-ins not opened yet) come with blank employee fields: fill them from the person's other rows first
    prof = [c for c in u.columns if c.startswith('Employee:') or c in ('Employee Username', 'Employee FirstName', 'Employee LastName', 'Employee Active Status')]
    u[prof] = u.groupby('id')[prof].transform(lambda c: c.ffill().bfill())
    u = u[(u['Employee: Country'] == 'Vietnam')].copy()
    u['name'] = u['Employee: Preferred First Name'].fillna(u['Employee FirstName']) + ' ' + u['Employee LastName']
    asof = u.sd.max().normalize()      # latest date any check-in was shared = the data date (comes from the User Initiated file)
    asof_s = f'{asof.month}/{asof.day}/{asof.year}'
    # Probation rule: hired less than 2 months before the data date -> listed by name ("New hire") and counted in Total Headcount only;
    # left out of every other number (SCP data, completion, ratings, no-SCP list).
    # Hire date = Original Hire Date (employee list, matched by e-mail); if not listed, the employee's first SCP launch date.
    hires = hire_dates(emp_path); first = u.groupby('id')['ld'].min(); mail = u.groupby('id')['Employee Username'].first().str.lower()
    hire_of = {i: hires.get(mail.get(i), first.get(i)) for i in first.index}
    tt = titles(emp_path); job_of = lambda i: tt.get(mail.get(i), '')      # job title; blank if not in the employee list
    cutoff = asof - pd.DateOffset(months=2)
    new_ids = {i for i, h in hire_of.items() if pd.notna(h) and h > cutoff}
    out = {}
    for tn in team_names:
        # find the team lead in the reporting chain: "Anh Nguyen" <-> "Nguyen, Anh (001810)" (word order and a middle initial are tolerated)
        want_id = find_manager(tn, mgr_counts)
        if want_id is None:
            warn(f'Team "{tn}" was NOT created: no manager with that name (or nobody reporting to them) in the OneOnOnes Manager Levels sheet. Check the spelling in Dashboard Access.xlsx or drop a newer OneOnOnes file.'); continue
        want = f'({want_id})'
        mem = ml[ml[lv].apply(lambda r: any(isinstance(x, str) and x.endswith(want) for x in r), axis=1)]
        lead_id = want_id
        chain = {r['EmployeeId']: [x for x in r[lv] if isinstance(x, str)] for _, r in mem.iterrows()}
        g = u[u.id.isin(chain)].copy()
        # active, Vietnam, has a record
        info = g.sort_values('ld').groupby('id').tail(1).set_index('id')
        active = info[info['Employee Active Status'] == 'Active']
        fh = lambda h: f'{h.month}/{h.day}/{h.year}' if pd.notna(h) else ''
        new_hires = []                      # probation: named, counted in headcount only, nothing else
        for i in sorted((i for i in active.index if i in new_ids), key=lambda i: active.loc[i, 'name']):
            ids_c = [mid(c) for c in chain[i]]; ids_c = ids_c[ids_c.index(lead_id):] if lead_id in ids_c else ids_c
            new_hires.append({'name': active.loc[i, 'name'], 'job': job_of(i), 'dept': active.loc[i, 'Employee: Department'] or '',
                              'sup': disp(chain[i][-1]) if chain[i] else tn, 'mgrs': ids_c, 'hire_date': fh(hire_of.get(i)), 'new': True})
        have = {i for i in active.index if i not in new_ids}
        g = g[g.id.isin(have)]
        g['dept'] = g.id.map(active['Employee: Department'])
        g['sup'] = g.id.map(lambda i: disp(chain[i][-1]) if chain[i] else tn)
        scp_ids = set(g[g['okf']].id)
        rec = g[(g.okf) & (g.sd > asof - pd.Timedelta(days=90))]; pri = g[(g.okf) & (g.sd <= asof - pd.Timedelta(days=90)) & (g.sd > asof - pd.Timedelta(days=180))]
        due = g[(g.ld > asof - pd.Timedelta(days=90)) & (g.ld <= asof)]
        ar, ap = round(float(rec.r.mean()), 2), round(float(pri.r.mean()), 2) if pri.r.notna().any() else None
        kpis = {'total_headcount': len(have) + len(new_hires), 'scp_headcount': len(scp_ids),
                'completion_rate': round(float(due['okf'].mean()), 3) if len(due) else None,
                'avg_recent': ar, 'avg_prior': ap, 'delta': round(ar - ap, 2) if ap is not None else None,
                'dist_recent': {str(k): int(v) for k, v in rec.r.dropna().astype(int).value_counts().sort_index().items()},
                'n_recent_rated': int(rec.r.notna().sum())}
        FCOL = next(c for c in g.columns if c.startswith('F - '))   # manager's written response
        W = 'E - How are you feeling about your workload and job duties this week? - Elaborate'
        cut = lambda v: ('' if pd.isna(v) else str(v).strip()[:3000])
        rec_ids = set(rec.index)
        checkins = [{'name': r['name'], 'sup': r.sup, 'r': int(r.r), 'date': fdate(r.sd), 'ym': r.sd.strftime('%b %Y'), 'mk': r.sd.strftime('%Y-%m'), 'rec': int(i in rec_ids),
                     'w': cut(r[W]), 'c': cut(r['E - Top 3 Challenges (can be new from prior week, or ongoing from prior weeks).']),
                     's': cut(r['E - Top 3 Successes from the prior week.']), 'p': cut(r['E - Top 3 Priorities for the upcoming week.']),
                     'o': cut(r['E - Any other topics you would like to include in your weekly SCP. ']),
                     'f': cut(r[FCOL]), 'fn': ('' if pd.isna(r['Facilitator FirstName']) else f"{r['Facilitator FirstName']} {r['Facilitator LastName']}")}
                    for i, r in g[g.okf & g.r.notna() & g.sd.notna()].sort_values('sd', ascending=False).iterrows()]
        s = g[g['okf'] & g.sd.notna()]
        trend = []
        for p, x in s.groupby(s.sd.dt.to_period('M')):
            trend.append({'label': p.strftime('%b %Y'), 'submissions': int(len(x)), 'avg_rating': round(float(x.r.mean()), 2) if x.r.notna().any() else None,
                          'low': int((x.r <= 2).sum()), 'high': int((x.r >= 4).sum())})
        def blocks(key):
            res = {}
            for k, x in g.groupby(key):
                ids = set(x.id); b = stats(x, asof)
                b.update({'headcount': len(ids), 'scp_headcount': len(ids & scp_ids)})
                res[k] = b
            return res
        dept, mgr = blocks('dept'), blocks('sup')
        for nh in new_hires:                 # new hires add to the headcount of their group only
            for grp, key in ((dept, nh['dept']), (mgr, nh['sup'])):
                if key in grp: grp[key]['headcount'] += 1
        dept_order = sorted(dept, key=lambda d: -dept[d]['headcount']); mgr_order = sorted(mgr, key=lambda d: -mgr[d]['headcount'])
        # no SCP: team members who are Active in the data but never submitted, plus members absent from the data (no record)
        no_scp = []
        for i, r in info.iterrows():
            if i in chain and r['Employee Active Status'] == 'Active' and i not in scp_ids and i not in new_ids:
                no_scp.append({'name': r['name'], 'num': i, 'job': job_of(i), 'dept': r['Employee: Department'] or '',
                               'supervisor': disp(chain[i][-1]) if chain[i] else tn, 'country': 'Vietnam', 'hire_date': fh(hire_of.get(i)), 'no_row': False})
        # employees table (last 12 weeks)
        emps = []; wk = {}
        for i, x in g.groupby('id'):
            if i not in scp_ids: continue
            ok, dn = week_stats(x.loc[x.okf, 'sd'], hire_of.get(i), asof); wk[i] = (ok, dn)
            last = x[x['okf']].sd.max()
            ids_chain = [mid(c) for c in chain[i]]; ids_chain = ids_chain[ids_chain.index(lead_id):] if lead_id in ids_chain else ids_chain
            emps.append({'name': x.name.iloc[0], 'job': job_of(i), 'sup': x.sup.iloc[0], 'mgrs': ids_chain,
                         'wk': sorted({int((d.normalize() - pd.Timestamp('1970-01-05')).days // 7) for d in x.loc[x.okf, 'sd'].dropna()}),
                         'hd': (lambda h: h.strftime('%Y-%m-%d') if pd.notna(h) else '')(hire_of.get(i)), 'ok': ok, 'due': dn, 'rate': round(100 * ok / dn) if dn else 0, 'last': fdate(last) if pd.notna(last) else '—'})
        emps += [dict(n, ok=None, due=None, rate=None, last='—') for n in new_hires]
        emps.sort(key=lambda e: e['name'])
        for grp, key in ((dept, 'dept'), (mgr, 'sup')):
            for k_, b_ in grp.items():
                ids_ = set(g.loc[g[key] == k_, 'id']); o_ = sum(wk[i][0] for i in ids_ if i in wk); d_ = sum(wk[i][1] for i in ids_ if i in wk)
                b_['completion'] = round(o_ / d_, 3) if d_ else None
        cnt = Counter(m for e in emps for m in e['mgrs'] if m != lead_id)
        names = {mid(c): disp(c) for ch in chain.values() for c in ch}
        mgr_filter = [{'id': m, 'name': names.get(m, m), 'n': n} for m, n in cnt.most_common()]
        # free text
        ct, st = 'E - Top 3 Challenges (can be new from prior week, or ongoing from prior weeks).', 'E - Top 3 Successes from the prior week.'
        sub = g[g['okf']].sort_values('sd', ascending=False)
        ch = sub[sub[ct].notna()]; su = sub[sub[st].notna()]
        cc, sc = themes_of(ch[ct]), themes_of(su[st], THEMES_S)
        top = lambda c: [(k, v) for k, v in c.most_common(3) if v]
        quotes = lambda d, col, low: [{'name': r['name'], 'date': fdate(r.sd), 'rating': int(r.r) if pd.notna(r.r) else None, 'text': str(r[col]).strip()}
                                      for _, r in (d[d.r <= 3] if low else d).head(12).iterrows()]
        trend_word = 'up' if (kpis['delta'] or 0) > 0 else 'down' if (kpis['delta'] or 0) < 0 else 'flat'
        pct = lambda n, d: f'{100 * n / d:.1f}%' if d else '0%'
        nr = kpis['n_recent_rated']; hi = sum(v for k, v in kpis['dist_recent'].items() if int(k) >= 4); lo = sum(v for k, v in kpis['dist_recent'].items() if int(k) <= 2)
        narrative = [f"Team sentiment is trending {trend_word}: {ar} avg rating over the last 90 days vs {ap} in the prior 90 days ({kpis['delta']:+}).",
                     f"Of {nr} rated check-ins in the last 90 days, {pct(hi, nr)} were 4–5 (Feeling Good/Great) and {pct(lo, nr)} were 1–2 (Struggling/Need Help; {lo} check-ins)."]
        ranked = [(m, b['avg_recent']) for m, b in mgr.items() if b['scp_headcount'] >= 3 and b['avg_recent'] is not None]
        if len(ranked) >= 2:
            lo_m, hi_m = min(ranked, key=lambda t: t[1]), max(ranked, key=lambda t: t[1])
            narrative.append(f'Lowest-rated supervisor group (3+ people, last 90d): {lo_m[0]} at {lo_m[1]:.2f}; highest: {hi_m[0]} at {hi_m[1]:.2f}.')
        fmt = lambda t, n: '; '.join(f'{k} ({v}, {round(100 * v / n)}% of {n} comments)' for k, v in top(t))
        theme_narr = [f'Most-mentioned challenge themes: {fmt(cc, len(ch))}.', f'Most-mentioned success themes: {fmt(sc, len(su))}.',
                      'Theme counts come from keyword matching on the free-text answers; treat them as indicative.']
        out[tn] = {'team': tn, 'asof': asof_s, 'kpis': kpis, 'checkins': checkins, 'trend': trend, 'dept': dept, 'dept_order': dept_order, 'mgr': mgr, 'mgr_order': mgr_order,
                   'no_scp': no_scp, 'new_hires': new_hires,
                   'months': sorted({c['mk'] for c in checkins}), 'min_wk': int((u.ld.min().normalize() - pd.Timestamp('1970-01-05')).days // 7), 'cur_wk': int((asof.normalize() - pd.Timestamp('1970-01-05')).days // 7),
                   'theme_rx': {'c': CHAL, 's': SUCC, 'cn': NONE_C, 'sn': NONE_S}, 'employees': emps, 'mgr_filter': mgr_filter,
                   'chal_counts': dict(cc), 'succ_counts': dict(sc), 'n_challenges': len(ch), 'n_successes': len(su), 'narrative': narrative, 'theme_narr': theme_narr}
    return out
