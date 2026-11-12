#!/usr/bin/env python3
"""Download the newest input files from ONE SharePoint folder (Microsoft Graph, app-only) into ./inputs.

Env: TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_URL (https://tcapc.sharepoint.com/sites/performance),
     FOLDER (e.g. "SCP Dashboard/Data"), LIBRARY (default Documents), FORCE (true = ignore "unchanged"),
     FILES = JSON {"role": "regex"}, e.g. {"data": "ScpExport_(\\d{4}-\\d{2}-\\d{2})\\.xlsx", "access": "Dashboard Access\\.xlsx"}
Rule: if the regex has a group (a date), the highest group wins; otherwise the most recently modified match wins.
Role names ending in "[]" collect SEVERAL files: regex group 1 = the key (e.g. team), group 2 (optional) = date; for each key the
newest date wins. They are saved as inputs/<role>/<key>.<ext>, e.g. FILES {"teams[]": "SCP_Dashboard_Vietnam_(\\w+)_Team_(\\d{4}-\\d{2}-\\d{2})\\.html"}.
Writes inputs/<role>.xlsx, inputs/names.json (role -> real file name), inputs/auth.json (public ids), state/fingerprint.txt
and prints  changed=true|false  (add it to $GITHUB_OUTPUT).
"""
import os, re, sys, json, urllib.parse
import requests
G = 'https://graph.microsoft.com/v1.0'
env = lambda k, d=None: os.environ.get(k, d) or (sys.exit(f'missing env {k}') if d is None else d)


def get(h, url, **kw):
    r = requests.get(url, headers=h, timeout=120, **kw); r.raise_for_status(); return r


def main():
    r = requests.post(f"https://login.microsoftonline.com/{env('TENANT_ID')}/oauth2/v2.0/token", data={
        'client_id': env('CLIENT_ID'), 'client_secret': env('CLIENT_SECRET'),
        'scope': 'https://graph.microsoft.com/.default', 'grant_type': 'client_credentials'}, timeout=60)
    r.raise_for_status(); h = {'Authorization': 'Bearer ' + r.json()['access_token']}
    u = urllib.parse.urlparse(env('SITE_URL'))
    site = get(h, f"{G}/sites/{u.netloc}:{u.path}").json()['id']
    drive = next((d['id'] for d in get(h, f"{G}/sites/{site}/drives").json()['value']
                  if d['name'].lower() == env('LIBRARY', 'Documents').lower()), None) or sys.exit('library not found')
    url, items = f"{G}/drives/{drive}/root:/{urllib.parse.quote(env('FOLDER').strip('/'))}:/children?$top=200", []
    while url:
        j = get(h, url).json(); items += j['value']; url = j.get('@odata.nextLink')
    files = [x for x in items if 'file' in x and not x['name'].startswith('~$')]
    root_files = []
    if os.environ.get('SHARED_FILES'):                      # shared files at the library root (e.g. the employee list)
        url = f"{G}/drives/{drive}/root/children?$top=200"
        while url:
            j = get(h, url).json(); root_files += [x for x in j['value'] if 'file' in x and not x['name'].startswith('~$')]; url = j.get('@odata.nextLink')     # sub-folders (e.g. Code) are ignored
    picked = {}
    jobs = [(r, rx, files) for r, rx in json.loads(env('FILES')).items()] + [(r, rx, root_files) for r, rx in json.loads(os.environ.get('SHARED_FILES') or '{}').items()]
    for role, rx, pool in jobs:
        hits = [(re.fullmatch(rx, x['name'], re.I), x) for x in pool]
        hits = [(m, x) for m, x in hits if m]
        if not hits: sys.exit(f'no file for "{role}" matching {rx}')
        if role.endswith('[]'):
            best = {}
            for m, x in hits:
                key, day = m.group(1), (m.group(2) if m.lastindex and m.lastindex >= 2 else x.get('lastModifiedDateTime', ''))
                if key not in best or day > best[key][0]: best[key] = (day, x)
            for key, (day, x) in best.items(): picked[f'{role[:-2]}/{key}'] = x
        else:
            picked[role] = max(hits, key=lambda t: (t[0].group(1) if t[0].groups() else t[1].get('lastModifiedDateTime', '')))[1]
    fp = '|'.join(f"{r}:{x['name']}:{x.get('eTag')}" for r, x in sorted(picked.items()))
    old = open('state/fingerprint.txt').read().strip() if os.path.exists('state/fingerprint.txt') else ''
    changed = fp != old or os.environ.get('FORCE') == 'true'
    print('files:', {r: x['name'] for r, x in picked.items()}, '| changed' if changed else '| unchanged', file=sys.stderr)
    os.makedirs('inputs', exist_ok=True); os.makedirs('state', exist_ok=True)
    if changed:
        for role, x in picked.items():
            ext = os.path.splitext(x['name'])[1] or '.xlsx'
            os.makedirs(os.path.dirname(f'inputs/{role}{ext}') or '.', exist_ok=True)
            with open(f'inputs/{role}{ext}', 'wb') as f:
                for c in get(h, f"{G}/drives/{drive}/items/{x['id']}/content", stream=True).iter_content(1 << 20): f.write(c)
        json.dump({r: x['name'] for r, x in picked.items()}, open('inputs/names.json', 'w'))
        json.dump({'clientId': env('CLIENT_ID'), 'tenantId': env('TENANT_ID')}, open('inputs/auth.json', 'w'))   # public ids
        open('state/fingerprint.txt', 'w').write(fp)
    print('changed=' + ('true' if changed else 'false'))


main()
