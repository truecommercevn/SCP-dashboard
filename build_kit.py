"""Reusable publisher: turns "one payload per person/team" into a protected site (one index.html + API).

Typical use in a dashboard's build.py:

    import build_kit as k
    views, master = k.read_access('inputs/access.xlsx')            # who may see which view (Excel)
    payloads = {'master': make_payload(None)}                       # everything, for master users
    for v in views: payloads[k.slug(v['name'])] = make_payload(v)   # one slice per view
    k.publish(index_html='index.html', out='site', api='api', payloads=payloads, views=views, master=master,
              auth=json.load(open('inputs/auth.json')), api_src='api_src')

`make_payload(view)` returns a JSON-able dict, e.g. {"DATA": ..., "META": {...}}. Optional: payloads['master'] may contain 'VIEWS': [{'slug':..., 'name':...}] - master users then get a view switcher (?view=slug).
Each top-level key becomes a
`const KEY = ...;` in the page after sign-in (see gate.html). Only the person's own payload is ever sent to the browser.
"""
import json, os, re, shutil
import openpyxl


def slug(s):
    return re.sub(r'[^a-z0-9]+', '-', s.lower()).strip('-')


def read_access(path):
    """Sheet 'Views': col A = view name (team / person / client), B = label, C = emails (; separated).
       Sheet 'Master access': col A = emails that see every view."""
    wb = openpyxl.load_workbook(path, data_only=True)
    split = lambda v: [x.strip().lower() for x in re.split(r'[;,\n]', str(v or '')) if '@' in x]
    views = [dict(name=str(r[0]).strip(), label=str(r[1] or '').strip(), emails=split(r[2]))
             for r in wb['Views'].iter_rows(min_row=2, values_only=True) if r[0] and str(r[0]).strip()]
    master = [e for r in wb['Master access'].iter_rows(min_row=2, values_only=True) for e in split(r[0])]
    if not views: raise SystemExit('access file: sheet "Views" is empty')
    return views, master


def publish(index_html, out, api, payloads, views, master, auth, api_src,
            frame_ancestors='https://*.sharepoint.com'):
    for d in (out, api):
        if os.path.isdir(d): shutil.rmtree(d)
    os.makedirs(out); os.makedirs(os.path.join(api, 'data'))
    shutil.copy(index_html, os.path.join(out, 'index.html'))              # the page, no data inside
    for key, payload in payloads.items():                                  # data, only reachable through the API
        json.dump(payload, open(os.path.join(api, 'data', key + '.json'), 'w', encoding='utf-8'), ensure_ascii=False)
    access = {}
    for v in views:
        if slug(v['name']) not in payloads: print('warning: no data for view', v['name']); continue
        for e in v['emails']: access.setdefault(e, []).append(slug(v['name']))
    for e in master: access.setdefault(e, []).append('master')
    json.dump(access, open(os.path.join(api, 'access.json'), 'w'), indent=1)
    if not auth.get('clientId') or not auth.get('tenantId'): raise SystemExit('missing clientId/tenantId')
    json.dump(auth, open(os.path.join(api, 'config.json'), 'w'))           # API verifies tokens with these
    json.dump(auth, open(os.path.join(out, 'auth-config.json'), 'w'))      # page signs in with these (public ids)
    shutil.copytree(api_src, api, dirs_exist_ok=True)                      # host.json, package.json, src/functions/data.js
    json.dump({'platform': {'apiRuntime': 'node:20'},
               'globalHeaders': {'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff',
                                 'Content-Security-Policy': f"frame-ancestors 'self' {frame_ancestors}"}},
              open(os.path.join(out, 'staticwebapp.config.json'), 'w'), indent=2)
    print(f"published {len(payloads)} views, {len(access)} sign-in emails; no email yet for:",
          [v['name'] for v in views if not v['emails']])
