#!/usr/bin/env python3
"""
Changes the site's description of its own licensing from open source to
open core.

The strategy of 12 September adopts open core: the platform and the core of
each product stay open, while fiscalisation, sector modules and the
accountants' edition are commercial. "Builds open-source business software"
describes something the company has decided not to be.

Deliberately narrow. Two distinct claims appear on this site and only one of
them is now wrong:

  "builds open-source business software"   - our own licensing. CHANGED.
  "built on open-source foundations"       - our dependencies, meaning
                                             PostgreSQL, Node, Vue, Linux.
                                             Still entirely true. UNTOUCHED.

A blanket find-and-replace would break the second while fixing the first,
which is why this names each string rather than matching a pattern.

Run from the repository root:  python3 open-core-wording.py
"""

import glob
import io

# (old, new) - exact strings, so nothing is changed by accident.
CHANGES = [
    (
        'no~bull consulting builds open-source business software for UK organisations',
        'no~bull consulting builds business software with an open-source core for UK organisations',
    ),
    (
        'About no~bull consulting - open-source software, built in the UK',
        'About no~bull consulting - business software with an open-source core, built in the UK',
    ),
    (
        'About no~bull consulting | Open-source software, built in the UK',
        'About no~bull consulting | Business software with an open-source core, built in the UK',
    ),
    (
        'that builds practical, open-source business applications',
        'that builds practical business software with an open-source core',
    ),
    (
        'Practical, open-source business applications for UK professionals and micro-businesses.',
        'Practical business software with an open-source core, for UK professionals and micro-businesses.',
    ),
    (
        'open-source software for UK organisations &mdash; from sole traders to established SMEs',
        'business software with an open-source core for UK organisations &mdash; from sole traders to established SMEs',
    ),
    # "We also build bespoke open-source tools" stays as it is: bespoke work
    # genuinely is delivered open, and that has not changed.
]

total = 0
for path in sorted(glob.glob('docs/*.html')):
    src = io.open(path, encoding='utf-8').read()
    original = src
    hits = []
    for old, new in CHANGES:
        n = src.count(old)
        if n:
            src = src.replace(old, new)
            hits.append(f'{n}x "{old[:52]}..."')
    if src != original:
        io.open(path, 'w', encoding='utf-8').write(src)
        total += len(hits)
        print(f'{path}')
        for h in hits:
            print(f'    {h}')

print(f'\n{total} replacements')

# What should remain, and should not have been touched.
print('\nStill present, and correct - these describe dependencies, not licensing:')
for path in sorted(glob.glob('docs/*.html')):
    src = io.open(path, encoding='utf-8').read()
    n = src.lower().count('open-source foundations') + src.lower().count('build on open source')
    if n:
        print(f'    {path}: {n}')
