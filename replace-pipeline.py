#!/usr/bin/env python3
"""
Replaces the "What else is coming" pipeline section on products.html.

Why: the strategy of 12 September narrows the portfolio to two products -
nBooks and nLearn - with nDocs reframed as a module, nSpark as a capability
woven through rather than a thing to buy, and nSales, nWork and nTel dropped.

A "Planned" card for a capability reads as a product that does not exist yet.
Five of them read as a company that could not decide. The replacement says
what the shared platform is, which is a genuine asset and shows deliberate
architecture rather than three unrelated applications.

Also fixes a stale label: nLearn was marked "In development" and is live.

Run from the repository root:  python3 replace-pipeline.py
"""

import io

PATH = 'docs/products.html'

LABEL = ('font-size:10px;font-weight:700;text-transform:uppercase;'
         'letter-spacing:0.12em;color:var(--oxford-blue);margin-bottom:12px;')
BODY = 'font-size:14px;color:var(--slate);margin-bottom:0;line-height:1.6;'

NEW = f'''<section class="section">
  <div class="container">
    <div class="section-label">Underneath</div>
    <h2 class="section-title" style="margin-bottom:20px;">The same foundations, both products.</h2>
    <p style="font-size:16px;color:var(--slate);margin-bottom:40px;max-width:720px;line-height:1.7;">
      nBooks and nLearn are not two unrelated applications that happen to share a name. They sit on
      the same shared services &mdash; one sign-in, one audit trail, one way of storing documents.
      Those foundations are open source, and stay that way.
    </p>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:28px;" class="two-col">
      <div class="pipeline-card">
        <div style="{LABEL}">Identity</div>
        <div class="product-label">nID</div>
        <p style="{BODY}">One account across everything you use, with multi-factor authentication
        and single sign-on. Your people sign in once, and access is removed in one place when
        somebody leaves.</p>
      </div>
      <div class="pipeline-card">
        <div style="{LABEL}">Documents</div>
        <div class="product-label">nDocs</div>
        <p style="{BODY}">Storage, versioning and retention for the paperwork behind your records.
        Making Tax Digital requires records to be preserved &mdash; this is how nBooks does it,
        rather than a separate thing to buy.</p>
      </div>
      <div class="pipeline-card">
        <div style="{LABEL}">Intelligence</div>
        <div class="product-label">nSpark</div>
        <p style="{BODY}">The receipt scanning in nBooks, and the document understanding behind it.
        A capability woven through the products rather than a product of its own.</p>
      </div>
      <div class="pipeline-card">
        <div style="{LABEL}">Trail</div>
        <div class="product-label">nAudit &amp; nEvents</div>
        <p style="{BODY}">Every change recorded, and every notification sent from one place. Dull,
        and the thing you want when somebody asks who changed a figure and when.</p>
      </div>
    </div>
    <p style="margin-top:32px;font-size:14px;color:var(--slate);">
      Have a specific requirement? <a href="services.html" style="color:var(--oxford-blue);">We also build bespoke</a>.
    </p>
  </div>
</section>'''

src = io.open(PATH, encoding='utf-8').read()

start = src.find('<section class="section">\n  <div class="container">\n    <div class="section-label">What&rsquo;s next</div>')
if start == -1:
    raise SystemExit('Could not find the pipeline section. Has products.html changed?')

end = src.find('</section>', src.find('We also build bespoke', start)) + len('</section>')
if end < start:
    raise SystemExit('Could not find the end of the pipeline section.')

removed = src[start:end]
for name in ('nSales', 'nWork', 'nTel'):
    if name in removed:
        print(f'  removing {name}')

src = src[:start] + NEW + src[end:]
io.open(PATH, 'w', encoding='utf-8').write(src)

print(f'\n  replaced {len(removed)} characters with {len(NEW)}')
print('  products.html written')

