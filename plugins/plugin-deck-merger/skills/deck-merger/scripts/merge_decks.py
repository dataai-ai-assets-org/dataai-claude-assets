"""
merge_decks.py — build one .pptx from slides picked out of several .pptx files,
keeping every slide byte-identical to its source.

    python3 merge_decks.py -o out.pptx \
        --deck corp.pptx  --slides 2-7,9-11 \
        --deck bl.pptx    --slides 4-8,10

Slide numbers are 1-based positions as PowerPoint shows them (hidden slides
included — run inspect_deck.py to see the numbering). `--slides all` takes
every slide. The order you write the ranges is the order the slides appear in
the output, so `--slides 9-11,2` really does put slide 2 last.

WHY THIS EXISTS: copying slides with python-pptx or pptxgenjs means re-creating
their content, which loses vector art, SmartArt, embedded objects, exact
positions and theme inheritance. This script instead copies the slide XML parts
verbatim and re-registers them in the package, so what you get out is what the
author designed.

HOW IT WORKS: the first deck is the base package. For every other deck, its
slide master, layouts, theme and media are imported alongside the base deck's
own master (never merged into it), so imported slides keep inheriting exactly
the fonts, colors and background art they were built against. Parts are renamed
on the way in to avoid collisions, relationship IDs are rewritten, and
everything left unreferenced afterwards is pruned.

Requires only the standard library. Validate the result afterwards (see
SKILL.md) — this script does not open PowerPoint.
"""
import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile

RELNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

CONTENT_TYPES = {
    'ppt/slides/': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    'ppt/slideLayouts/': 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    'ppt/slideMasters/': 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
    'ppt/theme/': 'application/vnd.openxmlformats-officedocument.theme+xml',
    'ppt/tags/': 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml',
    'ppt/charts/': 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    'ppt/notesSlides/': 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    'ppt/notesMasters/': 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
}
DIAGRAM_CT = {
    'data': 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
    'layout': 'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
    'quickStyle': 'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml',
    'colors': 'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
    'drawing': 'application/vnd.ms-office.drawingml.diagramDrawing+xml',
}
CHART_CT = {
    'chart': CONTENT_TYPES['ppt/charts/'],
    'colors': 'application/vnd.ms-office.chartcolorstyle+xml',
    'style': 'application/vnd.ms-office.chartstyle+xml',
    'chartEx': 'application/vnd.ms-office.chartex+xml',
}
DEFAULT_CT = {
    'bin': 'application/vnd.openxmlformats-officedocument.oleObject',
    'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpeg', 'gif': 'image/gif',
    'svg': 'image/svg+xml', 'emf': 'image/x-emf', 'wmf': 'image/x-wmf',
    'tiff': 'image/tiff', 'bmp': 'image/bmp', 'wdp': 'image/vnd.ms-photo',
    'mp4': 'video/mp4', 'wav': 'audio/wav', 'm4a': 'audio/mp4',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
}

# Notes and comments travel with a slide only if the target deck can host them;
# they carry author metadata and reviewer chatter that has no place in a client
# deck, so they are dropped from imported slides by default.
DROP_REL_TYPES = ('notesSlide', 'comments', 'commentAuthors', 'modernComment')


def parse_slides(spec, total):
    """'2-7,9-11' -> [2,3,4,5,6,7,9,10,11]. 'all' -> every slide."""
    if spec.strip().lower() == 'all':
        return list(range(1, total + 1))
    out = []
    for chunk in spec.split(','):
        chunk = chunk.strip()
        if not chunk:
            continue
        if '-' in chunk:
            a, b = chunk.split('-', 1)
            a, b = int(a), int(b)
            out.extend(range(a, b + 1) if a <= b else range(a, b - 1, -1))
        else:
            out.append(int(chunk))
    bad = [n for n in out if not 1 <= n <= total]
    if bad:
        raise SystemExit(f'slide number(s) out of range 1-{total}: {bad}')
    return out


def unzip(path, dest):
    with zipfile.ZipFile(path) as z:
        z.extractall(dest)
    return dest


def read(root, rel):
    with open(os.path.join(root, rel), encoding='utf-8') as f:
        return f.read()


def write(root, rel, text):
    p = os.path.join(root, rel)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, 'w', encoding='utf-8') as f:
        f.write(text)


def rels_path(part):
    d, f = os.path.split(part)
    return f'{d}/_rels/{f}.rels'


def slide_order(root):
    """Slide part names in presentation order."""
    pres = read(root, 'ppt/presentation.xml')
    rels = read(root, 'ppt/_rels/presentation.xml.rels')
    by_id = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]*slides/slide\d+\.xml)"', rels))
    out = []
    for tag in re.findall(r'<p:sldId [^>]*/>', pres):
        rid = re.search(r'r:id="(rId\d+)"', tag).group(1)
        t = by_id[rid]
        out.append(os.path.normpath(os.path.join('ppt', t)).replace('\\', '/'))
    return out


class Merger:
    def __init__(self, base_root):
        self.root = base_root
        self.ct = read(base_root, '[Content_Types].xml')
        self.copied = {}      # (deck_root, src_part) -> dest_part
        self.masters = {}     # deck_root -> dest master part

    # ---------- content types -------------------------------------------
    def ensure_default(self, ext):
        ext = ext.lstrip('.').lower()
        if not ext or re.search(rf'Extension="{ext}"', self.ct, re.I):
            return
        ct = DEFAULT_CT.get(ext, 'application/octet-stream')
        self.ct = self.ct.replace('<Default', f'<Default Extension="{ext}" ContentType="{ct}"/><Default', 1)

    def ensure_override(self, dest):
        part = '/' + dest
        if f'PartName="{part}"' in self.ct:
            return
        folder = os.path.dirname(dest) + '/'
        name = os.path.basename(dest)
        if folder in ('ppt/media/', 'ppt/embeddings/'):
            return
        if folder == 'ppt/diagrams/':
            ct = DIAGRAM_CT.get(re.match(r'([A-Za-z]+)', name).group(1))
        elif folder == 'ppt/charts/':
            ct = CHART_CT.get(re.match(r'([A-Za-z]+)', name).group(1))
        else:
            ct = CONTENT_TYPES.get(folder)
        if ct:
            self.ct = self.ct.replace('</Types>', f'<Override PartName="{part}" ContentType="{ct}"/></Types>')

    # ---------- copying --------------------------------------------------
    def alloc(self, src_part):
        """Free path in the base package mirroring the source part's folder."""
        folder, fn = os.path.split(src_part)
        folder += '/'
        m = re.fullmatch(r'(.*?)(\d+)(\.[A-Za-z0-9]+)', fn)
        stem, ext = (m.group(1), m.group(3)) if m else os.path.splitext(fn)
        d = os.path.join(self.root, folder)
        os.makedirs(d, exist_ok=True)
        n = 0
        for f in os.listdir(d):
            mm = re.fullmatch(re.escape(stem) + r'(\d+)' + re.escape(ext), f)
            if mm:
                n = max(n, int(mm.group(1)))
        return f'{folder}{stem}{n + 1}{ext}'

    def copy_part(self, deck_root, src_part, drop_rels=()):
        key = (deck_root, src_part)
        if key in self.copied:
            return self.copied[key]
        dest = self.alloc(src_part)
        self.copied[key] = dest
        os.makedirs(os.path.dirname(os.path.join(self.root, dest)), exist_ok=True)
        shutil.copy(os.path.join(deck_root, src_part), os.path.join(self.root, dest))
        self.ensure_default(os.path.splitext(dest)[1])
        self.ensure_override(dest)

        src_rels = os.path.join(deck_root, rels_path(src_part))
        if not os.path.exists(src_rels):
            return dest
        with open(src_rels, encoding='utf-8') as f:
            xml = f.read()
        out = xml
        src_dir = os.path.dirname(src_part)
        for tag in re.findall(r'<Relationship\b[^>]*/>', xml):
            if 'TargetMode="External"' in tag:
                continue
            typ = re.search(r'Type="[^"]*/([A-Za-z]+)"', tag).group(1)
            tgt = re.search(r'Target="([^"]+)"', tag).group(1)
            if typ in drop_rels:
                out = out.replace(tag, '')
                continue
            child = os.path.normpath(os.path.join(src_dir, tgt)).replace('\\', '/')
            if not os.path.exists(os.path.join(deck_root, child)):
                out = out.replace(tag, '')       # dangling in the source deck
                continue
            new_child = self.copy_part(deck_root, child, drop_rels)
            new_tgt = os.path.relpath(new_child, os.path.dirname(dest)).replace('\\', '/')
            out = out.replace(tag, tag.replace(f'Target="{tgt}"', f'Target="{new_tgt}"'))
        write(self.root, rels_path(dest), out)
        return dest

    def import_master(self, deck_root):
        if deck_root not in self.masters:
            masters = sorted(os.listdir(os.path.join(deck_root, 'ppt/slideMasters')))
            first = next(m for m in masters if m.endswith('.xml'))
            self.masters[deck_root] = self.copy_part(deck_root, 'ppt/slideMasters/' + first)
        return self.masters[deck_root]

    def import_slide(self, deck_root, src_part):
        self.import_master(deck_root)   # so the slide's layout lands under it
        return self.copy_part(deck_root, src_part, drop_rels=DROP_REL_TYPES)

    # ---------- presentation wiring --------------------------------------
    def add_presentation_rel(self, target, typ):
        p = 'ppt/_rels/presentation.xml.rels'
        xml = read(self.root, p)
        used = [int(n) for n in re.findall(r'Id="rId(\d+)"', xml)]
        rid = f'rId{max(used) + 1}'
        tag = f'<Relationship Id="{rid}" Type="{RELNS}/{typ}" Target="{target}"/>'
        write(self.root, p, xml.replace('</Relationships>', tag + '</Relationships>'))
        return rid

    def rebuild(self, slide_parts, extra_masters):
        """slide_parts: dest part names, in final order."""
        pres = read(self.root, 'ppt/presentation.xml')
        rels = read(self.root, 'ppt/_rels/presentation.xml.rels')
        rid_by_target = {os.path.normpath(os.path.join('ppt', t)).replace('\\', '/'): rid
                         for rid, t in re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels)}
        sid = max([int(n) for n in re.findall(r'<p:sldId id="(\d+)"', pres)] or [255])
        entries = []
        for part in slide_parts:
            rid = rid_by_target.get(part)
            if rid is None:
                rid = self.add_presentation_rel(
                    os.path.relpath(part, 'ppt').replace('\\', '/'), 'slide')
                rid_by_target[part] = rid
            sid += 1
            entries.append(f'<p:sldId id="{sid}" r:id="{rid}"/>')
        # Slides left out of the selection must lose their presentation
        # relationship too, otherwise they stay "referenced" and prune() keeps
        # them and all their media in the file.
        keep_rids = {rid_by_target[p] for p in slide_parts}
        rels = read(self.root, 'ppt/_rels/presentation.xml.rels')
        for tag in re.findall(r'<Relationship\b[^>]*/>', rels):
            if '/slide"' not in tag.replace('Type="' + RELNS, 'Type="'):
                continue
            rid = re.search(r'Id="(rId\d+)"', tag).group(1)
            if rid not in keep_rids:
                rels = rels.replace(tag, '')
        write(self.root, 'ppt/_rels/presentation.xml.rels', rels)

        pres = read(self.root, 'ppt/presentation.xml')
        pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>',
                      '<p:sldIdLst>' + ''.join(entries) + '</p:sldIdLst>', pres, flags=re.S)

        mlst = re.search(r'<p:sldMasterIdLst>.*?</p:sldMasterIdLst>', pres, re.S).group(0)
        mid = max(int(n) for n in re.findall(r'<p:sldMasterId id="(\d+)"', mlst))
        add = ''
        for m in extra_masters:
            rid = self.add_presentation_rel(os.path.relpath(m, 'ppt').replace('\\', '/'), 'slideMaster')
            mid += 1
            add += f'<p:sldMasterId id="{mid}" r:id="{rid}"/>'
        pres = pres.replace(mlst, mlst.replace('</p:sldMasterIdLst>', add + '</p:sldMasterIdLst>'))
        write(self.root, 'ppt/presentation.xml', pres)

    def unique_layout_ids(self):
        """PowerPoint rejects duplicate sldLayoutId values across masters."""
        d = os.path.join(self.root, 'ppt/slideMasters')
        used = set()
        for fn in sorted(os.listdir(d)):
            if not fn.endswith('.xml'):
                continue
            rel = 'ppt/slideMasters/' + fn
            xml = read(self.root, rel)

            def fix(m):
                i = int(m.group(1))
                if i in used:
                    n = max(used) + 1
                    used.add(n)
                    return m.group(0).replace(f'id="{i}"', f'id="{n}"')
                used.add(i)
                return m.group(0)

            new = re.sub(r'<p:sldLayoutId id="(\d+)"[^>]*/>', fix, xml)
            if new != xml:
                write(self.root, rel, new)

    # ---------- pruning ---------------------------------------------------
    def prune(self):
        """Drop every part no longer reachable from the package root rels."""
        reachable, queue = set(), ['_rels/.rels']
        while queue:
            rp = queue.pop()
            if rp in reachable or not os.path.exists(os.path.join(self.root, rp)):
                continue
            reachable.add(rp)
            src_dir = os.path.dirname(os.path.dirname(rp)) if rp.endswith('.rels') else ''
            if not rp.endswith('.rels'):
                continue
            with open(os.path.join(self.root, rp), encoding='utf-8') as f:
                xml = f.read()
            for tag in re.findall(r'<Relationship\b[^>]*/>', xml):
                if 'TargetMode="External"' in tag:
                    continue
                tgt = re.search(r'Target="([^"]+)"', tag).group(1)
                part = os.path.normpath(os.path.join(src_dir, tgt.lstrip('/'))).replace('\\', '/')
                if part in reachable:
                    continue
                reachable.add(part)
                queue.append(rels_path(part))

        removed = []
        for dirpath, _dirs, files in os.walk(self.root):
            for fn in files:
                full = os.path.join(dirpath, fn)
                rel = os.path.relpath(full, self.root).replace('\\', '/')
                if rel == '[Content_Types].xml' or rel in reachable:
                    continue
                if rel.endswith('.rels'):
                    owner = re.sub(r'_rels/([^/]+)\.rels$', r'\1', rel)
                    if owner in reachable:
                        continue
                os.remove(full)
                removed.append(rel)
                self.ct = re.sub(rf'<Override PartName="/{re.escape(rel)}"[^>]*/>', '', self.ct)
        return removed

    def save(self, out_path):
        write(self.root, '[Content_Types].xml', self.ct)
        out_path = os.path.abspath(out_path)
        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for dirpath, _dirs, files in os.walk(self.root):
                for fn in files:
                    full = os.path.join(dirpath, fn)
                    z.write(full, os.path.relpath(full, self.root).replace('\\', '/'))
        return out_path


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('-o', '--out', required=True)
    ap.add_argument('--deck', action='append', required=True, help='path to a .pptx (repeatable)')
    ap.add_argument('--slides', action='append', required=True,
                    help="selection for the matching --deck, e.g. '2-7,9-11' or 'all'")
    ap.add_argument('--keep-notes', action='store_true',
                    help='keep speaker notes on imported slides (dropped by default)')
    args = ap.parse_args()

    if len(args.deck) != len(args.slides):
        raise SystemExit('each --deck needs exactly one --slides')

    global DROP_REL_TYPES
    if args.keep_notes:
        DROP_REL_TYPES = tuple(t for t in DROP_REL_TYPES if t != 'notesSlide')

    tmp = tempfile.mkdtemp(prefix='deckmerge-')
    roots = [unzip(d, os.path.join(tmp, f'deck{i}')) for i, d in enumerate(args.deck)]

    merger = Merger(roots[0])
    final = []

    base_slides = slide_order(roots[0])
    for n in parse_slides(args.slides[0], len(base_slides)):
        final.append(base_slides[n - 1])

    extra_masters = []
    for root, spec, path in zip(roots[1:], args.slides[1:], args.deck[1:]):
        slides = slide_order(root)
        before = set(merger.masters)
        for n in parse_slides(spec, len(slides)):
            final.append(merger.import_slide(root, slides[n - 1]))
        if root not in before:
            extra_masters.append(merger.masters[root])
        print(f'imported {len(parse_slides(spec, len(slides)))} slide(s) from {path}')

    merger.rebuild(final, extra_masters)
    merger.unique_layout_ids()
    removed = merger.prune()
    out = merger.save(args.out)

    print(f'pruned {len(removed)} unreferenced part(s)')
    print(f'wrote {out} — {len(final)} slides')


if __name__ == '__main__':
    sys.exit(main())
