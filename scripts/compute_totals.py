import re, json

with open('index.html', encoding='utf-8') as f:
    html = f.read()

i1 = html.index('const DATA = [') + len('const DATA = ')
i2 = html.index('];', i1) + 1
data = json.loads(html[i1:i2])

# total dashboard (all rows)
may = [r for r in data if r.get('month') in ('Май','2026-05')]
jun = [r for r in data if r.get('month') in ('Июнь','2026-06')]

def agg(rows):
    qty_ship    = sum(float(r.get('shippedQty',0) or 0)         for r in rows)
    qty_plan    = sum(float(r.get('planQty',0) or 0)            for r in rows)
    qty_rel     = sum(float(r.get('releaseQty',0) or 0)         for r in rows)
    rev_fact    = sum(float(r.get('actualRevenue',0) or 0)      for r in rows)
    rev_plan    = sum(float(r.get('plannedRevenue',0) or 0)     for r in rows)
    gp_fact     = sum(float(r.get('actualGp',0) or 0)           for r in rows)
    gp_plan     = sum(float(r.get('plannedGp',0) or 0)          for r in rows)
    return qty_ship,qty_plan,qty_rel,rev_fact,rev_plan,gp_fact,gp_plan

for label, rows in [('Май', may), ('Июнь', jun)]:
    qs,qp,qr,rf,rp,gf,gp = agg(rows)
    print(f'\n=== {label} (rows={len(rows)}) ===')
    print(f'  отгружено шт:      {qs:>15,.0f}')
    print(f'  план шт:           {qp:>15,.0f}')
    print(f'  выпущено шт:       {qr:>15,.0f}')
    print(f'  факт.выручка RUB:    {rf:>15,.0f}')
    print(f'  плановая выручка:  {rp:>15,.0f}')
    print(f'  факт.валовка RUB:    {gf:>15,.0f}')
    print(f'  плановая валовка:  {gp:>15,.0f}')

# Top by actualRevenue
may_sorted = sorted(may, key=lambda r: float(r.get('actualRevenue',0) or 0), reverse=True)
print('\n=== TOP 10 Май по факт.выручке ===')
for r in may_sorted[:10]:
    print(f"  {r.get('sku',''):<14} ship={float(r.get('shippedQty',0) or 0):>7,.0f}  rev={float(r.get('actualRevenue',0) or 0):>12,.0f}  gp={float(r.get('actualGp',0) or 0):>11,.0f}")

# Save SKU list for cube query
top_skus = [r.get('sku','') for r in may_sorted[:30] if r.get('sku')]
open('data/may_top_skus.txt','w',encoding='utf-8').write('\n'.join(top_skus))
print(f'\nsaved {len(top_skus)} top skus to data/may_top_skus.txt')
