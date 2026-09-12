#!/usr/bin/env python3
"""
Brings the homepage and products page into line with the open core strategy.

The problem: the homepage offers "a perpetual licence and the source code" as
the company's whole proposition. That is NobleProg's nLearn arrangement
generalised to everything, and under open core it is not true of nBooks - a
subscriber gets an open core, complete data export, and the option to buy
perpetual rights, not ownership by default.

Adopting the Strategy §7.2 names this exact ambiguity as a diligence finding:
"If it means IP ownership in any form, an acquirer discovers that the company
does not own its own product."

What the new wording keeps, because all of it is true:
  - UK-built, UK-hosted, no US parent
  - an open-source core
  - complete data export
  - somebody who answers the phone

What it drops: the promise of ownership as the default.

What it does not add: sector modules, the accountants' edition and
fiscalisation are the strategy's differentiators and none of them exists yet.
A homepage claiming them in 2026 is writing the 2028 site early. They belong
on a roadmap, and on a practices page that does not exist yet.

Run from the repository root:  python3 homepage-open-core.py
"""

import io

CHANGES = {
    'docs/index.html': [
        # --- the hero ---------------------------------------------------
        (
            '<h1>Software you own.<br>Built by someone<br>who answers the phone.</h1>',
            '<h1>Software for the rules<br>you have to follow.<br>'
            'Built by someone who answers the phone.</h1>',
        ),
        (
            'Business software for organisations tired of renting their own data back. UK-built, UK-hosted,\n'
            '      and yours &mdash; <strong>a perpetual licence and the source code</strong>, so it keeps working\n'
            '      whatever happens to us.',
            'Accounting and training software for UK organisations, built around the obligations rather than '
            'adapted to them. UK-built, UK-hosted, with <strong>an open-source core and your data yours to '
            'take at any time</strong>.',
        ),
        # --- metadata ---------------------------------------------------
        (
            '<title>no~bull consulting | Business software you own | UK-built, UK-hosted</title>',
            '<title>no~bull consulting | UK accounting and training software | UK-built, UK-hosted</title>',
        ),
        (
            'content="no~bull consulting — business software you own"',
            'content="no~bull consulting — UK accounting and training software"',
        ),
        (
            'content="Open-source business software you own outright. UK-hosted, source code included, no lock-in."',
            'content="UK accounting and training software with an open-source core. UK-hosted, complete data export, no lock-in."',
        ),
        (
            'content="MTD VAT &amp; ITSA ready. Open source, UK-hosted, no lock-in. From £9/month."',
            'content="MTD VAT &amp; ITSA ready. Open-source core, UK-hosted, no lock-in. From £9/month."',
        ),
        (
            '"description": "UK-built, UK-hosted business software you own outright. nLearn for training coordination, with more tools in development.",',
            '"description": "UK-built, UK-hosted business software with an open-source core. nBooks for accounting and Making Tax Digital, nLearn for training coordination.",',
        ),
    ],
    'docs/products.html': [
        (
            'nSuite: nBooks (UK accounting, HMRC MTD, open source, UK-hosted) and more open-source business tools in development.',
            'nSuite: nBooks (UK accounting, HMRC MTD, UK-hosted) and nLearn (training coordination), built on shared open-source foundations.',
        ),
        (
            'Built on open source, hosted in the UK, and handed over with the source code. You own what we',
            'Built on open source, hosted in the UK, with an open-source core and complete data export. You keep what we',
        ),
    ],
}

for path, pairs in CHANGES.items():
    src = io.open(path, encoding='utf-8').read()
    original = src
    print(path)
    for old, new in pairs:
        n = src.count(old)
        if n == 0:
            print(f'    NOT FOUND: "{old[:60]}..."')
        else:
            src = src.replace(old, new)
            print(f'    {n}x changed: "{old[:60]}..."')
    if src != original:
        io.open(path, 'w', encoding='utf-8').write(src)
        print('    written')
    print()

print('Remaining ownership claims across the site:')
import glob
for p in sorted(glob.glob('docs/*.html')):
    s = io.open(p, encoding='utf-8').read().lower()
    n = s.count('you own') + s.count('own outright') + s.count('yours to keep')
    if n:
        print(f'    {p}: {n}')
print()
print('nlearn.html keeps its ownership language deliberately - NobleProg does')
print('hold a perpetual licence and the source, so on that page it is accurate.')
