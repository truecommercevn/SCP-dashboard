#!/usr/bin/env python3
"""SCP Dashboard: one protected page + one data slice per team lead.

Inputs (downloaded by fetch_kit.py into ./inputs):
  inputs/access.xlsx            who may see which team (Dashboard Access.xlsx)
  inputs/oneonones.xlsx         newest OneOnOnes_*.xlsx   (sheet "Manager Levels")
  inputs/scp.xlsx               "User Initiated 1-on-1s.xlsx" (sheet "Data")
  inputs/employees.xlsx         shared employee list at the library root (hire dates; optional, falls back to first SCP date)
"""
import argparse, glob, json, os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0, HERE)
import build_kit as k, build_data as bd


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default='site'); ap.add_argument('--api', default='api')
    ap.add_argument('--frame-ancestors', default='https://*.sharepoint.com')
    a = ap.parse_args()
    views, master = k.read_access('inputs/access.xlsx')
    names = [v['name'] for v in views]
    teams_by_name = bd.build_teams('inputs/oneonones.xlsx', 'inputs/scp.xlsx', names, 'inputs/employees.xlsx')
    teams = {k.slug(n): d for n, d in teams_by_name.items()}
    names = {k.slug(n): n for n in teams_by_name}
    if not teams: raise SystemExit('no team built; check names in Dashboard Access.xlsx')
    payloads = {slug: {'DATA': d} for slug, d in teams.items()}
    order = sorted(teams, key=lambda s: names[s])
    payloads['master'] = {'DATA': teams[order[0]], 'VIEWS': [{'slug': s, 'name': names[s]} for s in order]}
    asof = {d.get('asof') for d in teams.values()}
    k.publish(os.path.join(HERE, 'index.html'), a.out, a.api, payloads, views, master,
              json.load(open('inputs/auth.json')), os.path.join(HERE, 'api_src'), a.frame_ancestors)
    print('teams:', sorted(teams), '| as of:', sorted(asof))


main()
