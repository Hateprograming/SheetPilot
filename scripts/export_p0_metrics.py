"""Read-only, local administrator export. Does not export user text or workbook data."""
import argparse
import csv
import json
import sqlite3
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--database', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    if args.database.resolve() == args.output.resolve():
        parser.error('Output must not replace the database.')
    fields = ['operation_id','status','result_version','server_ms','waiting_ms','outcome','baseline_minutes','actual_total_minutes','net_saved_minutes']
    with sqlite3.connect(args.database.resolve().as_uri()+'?mode=ro', uri=True) as db, args.output.open('x', encoding='utf-8-sig', newline='') as stream:
        db.row_factory = sqlite3.Row
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        for op in db.execute('SELECT operation_id,status,result_version FROM p0_operations ORDER BY created_at'):
            stages = db.execute('SELECT stage,SUM(duration_ms) ms FROM p0_measurements WHERE operation_id=? GROUP BY stage', (op['operation_id'],)).fetchall()
            fb = db.execute('SELECT payload_json FROM p0_feedback WHERE operation_id=?', (op['operation_id'],)).fetchone()
            data = json.loads(fb['payload_json']) if fb else {}
            baseline, actual = data.get('baseline_minutes'), data.get('actual_total_minutes')
            writer.writerow({**dict(op), 'server_ms':sum(r['ms'] for r in stages if r['stage']!='waiting'),
                             'waiting_ms':sum(r['ms'] for r in stages if r['stage']=='waiting'),
                             'outcome':data.get('outcome'), 'baseline_minutes':baseline, 'actual_total_minutes':actual,
                             'net_saved_minutes':None if baseline is None or actual is None else baseline-actual})


if __name__ == '__main__':
    main()
