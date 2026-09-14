"""
build_picker.py — emit the HTML for an interactive slide picker so the user can
tick the slides they want instead of reading an inventory and typing numbers.

    python3 build_picker.py corporate.pptx business_line.pptx > picker.html
    python3 build_picker.py corp.pptx --label "Corp deck" bl.pptx --label "Data & AI" \
        --preselect "1:2-12" --preselect "2:4-13"

The output is a self-contained HTML fragment to pass as `widget_code` to the
visualizer's show_widget tool. Do not wrap it in DOCTYPE/html/body, do not add a
title inside it — the surrounding explanation belongs in the chat response.

What it does for you:

- Numbers every slide exactly as PowerPoint's slide pane does, hidden slides
  included, which is the numbering merge_decks.py expects.
- Flags hidden slides with a badge, and flags likely housekeeping slides
  ("this deck contains reusable content", "use the latest template") as
  Internal, so the user can see what they are ticking.
- Starts with every box empty. The selection is the user's to make, not yours to
  propose — the badges tell them what to avoid, and that is enough. Use
  --preselect only when the user asks to start from a selection they already
  made.
- On submit, calls sendPrompt() with a compact instruction naming each deck and
  its ranges, e.g. "Merge these slides: Corp deck 2-12; Data & AI 4-13." —
  which comes back as the user's next message, ready to feed merge_decks.py.

Slide descriptions come from the slide's own text and are necessarily terse. Read
them before showing the picker and rewrite any that would not mean anything to
the user: `--describe N:text` overrides one row.

Requires python-pptx.
"""
import argparse
import html
import re
import sys

from pptx import Presentation

HOUSEKEEPING = re.compile(
    r'reusable|sharable|shareable|latest (elca )?template|only use slides|'
    r'can be found here|corporate slides|this (slide )?deck contains|'
    r'use the comments feature|do not modify|for internal use',
    re.I,
)


def slide_text(slide, limit=90):
    """Title-ish line plus the next distinct line, which is usually enough to
    recognise a slide by."""
    lines = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for line in shape.text_frame.text.splitlines():
            line = ' '.join(line.split())
            if line and line not in lines:
                lines.append(line)
                break
        if len(lines) >= 2:
            break
    if not lines:
        return '(no text)'
    text = ' — '.join(lines)
    return text[:limit].rstrip() + ('…' if len(text) > limit else '')


def parse_ranges(spec):
    out = set()
    for part in spec.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a, b = part.split('-', 1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    return out


def inventory(path):
    prs = Presentation(path)
    rows = []
    for i, slide in enumerate(prs.slides, 1):
        text = slide_text(slide)
        rows.append({
            'n': i,
            'text': text,
            'hidden': slide._element.get('show') == '0',
            'internal': bool(HOUSEKEEPING.search(text)),
        })
    return rows


BADGE = ('<span style="font-size:11px;background:var(--bg-{bg});'
         'color:var(--text-{fg});padding:2px 7px;border-radius:var(--radius)">{label}</span>')
ROW = ('<label style="display:flex;gap:10px;align-items:center;padding:7px 12px;'
       '{border}font-size:13px"><input type="checkbox" data-d="{key}" data-n="{n}"{checked}>'
       '<span style="color:var(--text-muted);min-width:20px;text-align:right">{n}</span>'
       '<span style="flex:1">{text}</span>{badge}</label>')

SCRIPT = """<script>
var boxes=[].slice.call(document.querySelectorAll('input[data-n]'));
var L=__LABELS__;
function sel(d){return boxes.filter(function(b){return b.dataset.d===d&&b.checked}).map(function(b){return +b.dataset.n}).sort(function(a,b){return a-b})}
function rng(a){var o=[],i=0;while(i<a.length){var s=i;while(i+1<a.length&&a[i+1]===a[i]+1)i++;o.push(s===i?''+a[s]:a[s]+'-'+a[i]);i++}return o.join(',')}
function total(){return Object.keys(L).reduce(function(t,k){return t+sel(k).length},0)}
function upd(){var n=total();document.getElementById('cnt').textContent=n+(n===1?' slide selected':' slides selected');if(n)document.getElementById('err').hidden=true}
boxes.forEach(function(b){b.addEventListener('change',upd)});
function set(d,v){boxes.forEach(function(b){if(b.dataset.d===d)b.checked=v});upd()}
Object.keys(L).forEach(function(k){
  document.getElementById('all-'+k).onclick=function(){set(k,true)};
  document.getElementById('none-'+k).onclick=function(){set(k,false)};
});
document.getElementById('go').onclick=function(){
  if(!total()){document.getElementById('err').hidden=false;return}
  var p=Object.keys(L).map(function(k){var s=sel(k);return s.length?L[k]+' '+rng(s):null}).filter(Boolean);
  var t='Merge these slides: '+p.join('; ')+'.';
  if(document.getElementById('notes').checked)t+=' Keep speaker notes.';
  sendPrompt(t);
};
upd();
</script>"""


def build(decks, describe):
    keys = {}
    parts = []
    counts = []
    for idx, (path, label, rows, preselect) in enumerate(decks, 1):
        key = 'd%d' % idx
        keys[key] = label
        parts.append(
            '<div style="display:flex;justify-content:space-between;align-items:center;'
            'margin:0 0 8px"><span style="font-size:16px;font-weight:500">%s</span>'
            '<span><button type="button" id="all-%s" style="font-size:12px;padding:4px 10px">All</button> '
            '<button type="button" id="none-%s" style="font-size:12px;padding:4px 10px">None</button>'
            '</span></div>' % (html.escape(label), key, key))
        parts.append('<div style="background:var(--surface-2);border:0.5px solid var(--border);'
                     'border-radius:12px;overflow:hidden;margin:0 0 1.5rem">')
        for j, row in enumerate(rows):
            if preselect is None:
                checked = False
            else:
                checked = row['n'] in preselect
            if checked:
                counts.append(1)
            badge = ''
            if row['hidden']:
                badge = BADGE.format(bg='warning', fg='warning', label='Hidden')
            elif row['internal']:
                badge = BADGE.format(bg='neutral', fg='secondary', label='Internal')
            text = describe.get((idx, row['n']), row['text'])
            parts.append(ROW.format(
                key=key, n=row['n'], text=html.escape(text), badge=badge,
                checked=' checked' if checked else '',
                border='' if j == len(rows) - 1 else 'border-bottom:0.5px solid var(--border);'))
        parts.append('</div>')

    n = len(counts)
    parts.append(
        '<div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap">'
        '<button type="button" id="go">Merge selected slides ↗</button>'
        '<span id="cnt" style="font-size:13px;color:var(--text-secondary)">%d slides selected</span>'
        '<label style="display:flex;gap:8px;align-items:center;font-size:13px;'
        'color:var(--text-secondary)"><input type="checkbox" id="notes">Keep speaker notes</label></div>'
        '<p id="err" style="font-size:13px;color:var(--text-danger);margin:8px 0 0" hidden>'
        'Tick at least one slide first.</p>' % n)

    summary = ('Interactive slide picker listing every slide in %s, each with a checkbox, '
               'so a selection can be sent back to chat for merging.'
               % ' and '.join(d[1] for d in decks))
    head = '<h2 class="sr-only">%s</h2>\n\n' % html.escape(summary)
    labels = '{' + ','.join('"%s":"%s"' % (k, v.replace('"', '\\"')) for k, v in keys.items()) + '}'
    return head + '\n'.join(parts) + '\n\n' + SCRIPT.replace('__LABELS__', labels)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('decks', nargs='+', help='paths to .pptx files, in output order')
    ap.add_argument('--label', action='append', default=[],
                    help='display name for the matching deck (defaults to filename)')
    ap.add_argument('--preselect', action='append', default=[],
                    help="DECKINDEX:RANGES, e.g. '1:2-12' — 1-based deck index. Only when "
                         "the user asks to start from a selection they already made; the "
                         "default is nothing ticked")
    ap.add_argument('--describe', action='append', default=[],
                    help="DECKINDEX:SLIDE:TEXT — replace one row's description")
    ap.add_argument('-o', '--out', help='write here instead of stdout')
    args = ap.parse_args()

    pre = {}
    for spec in args.preselect:
        i, ranges = spec.split(':', 1)
        pre[int(i)] = parse_ranges(ranges)

    describe = {}
    for spec in args.describe:
        i, n, text = spec.split(':', 2)
        describe[(int(i), int(n))] = text

    decks = []
    for i, path in enumerate(args.decks, 1):
        label = args.label[i - 1] if len(args.label) >= i else path.split('/')[-1].rsplit('.', 1)[0]
        decks.append((path, label, inventory(path), pre.get(i)))

    out = build(decks, describe)
    if args.out:
        with open(args.out, 'w') as fh:
            fh.write(out)
        print('wrote %s (%d slides across %d deck(s))'
              % (args.out, sum(len(d[2]) for d in decks), len(decks)), file=sys.stderr)
    else:
        print(out)


if __name__ == '__main__':
    main()
