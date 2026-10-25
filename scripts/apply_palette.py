#!/usr/bin/env python3
"""Apply bounded palette edits without round-tripping unrelated PPTX objects."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
import zipfile
from xml.parsers import expat

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
SLIDE = re.compile(r'ppt/slides/slide[1-9][0-9]*\.xml\Z')
HEX = re.compile(r'#?([0-9a-fA-F]{6})\Z')


def color(value):
    match = HEX.fullmatch(str(value))
    if not match:
        raise ValueError('Colors must be six-digit RGB hex values.')
    return match[1].upper()


def nodes(raw):
    """Parse namespaces and locate byte spans, retaining original XML bytes."""
    parser = expat.ParserCreate(namespace_separator='}')
    stack, result = [], []

    def start(name, attrs):
        offset = parser.CurrentByteIndex
        # XML attributes may contain a quoted greater-than sign.
        match = re.match(br'<(?:"[^"]*"|\x27[^\x27]*\x27|[^\x27">])*>', raw[offset:])
        if not match:
            raise ValueError('Cannot locate XML opening tag.')
        tag = match[0]
        node = {'name': name, 'attrs': attrs, 'start': offset,
                'openEnd': offset + len(tag), 'parent': stack[-1] if stack else None,
                'children': [], 'selfClosing': tag.rstrip().endswith(b'/>')}
        if stack:
            stack[-1]['children'].append(node)
        stack.append(node)
        result.append(node)

    def end(_name):
        node = stack.pop()
        if node['selfClosing']:
            node['closeStart'] = node['end'] = node['openEnd']
        else:
            node['closeStart'] = parser.CurrentByteIndex
            node['end'] = raw.index(b'>', node['closeStart']) + 1

    def reject(*_args):
        raise ValueError('DTD and entity declarations are not supported.')

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.StartDoctypeDeclHandler = reject
    parser.EntityDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    parser.Parse(raw, True)
    return result


def ancestor(node, name):
    node = node['parent']
    while node:
        if node['name'] == name:
            return node
        node = node['parent']
    return None


def validate_plan(plan):
    if not isinstance(plan, dict):
        raise ValueError('Palette plan must be a JSON object.')
    if plan.get('schemaVersion') != '1.0':
        raise ValueError('Expected palette plan schemaVersion 1.0.')
    rules = plan.get('colorRules', [])
    effects = plan.get('iconEffects', [])
    protected = plan.get('protectedPictures', [])
    if not isinstance(rules, list) or not isinstance(effects, list) or not isinstance(protected, list):
        raise ValueError('Plan collections must be arrays.')
    if not rules and not effects:
        raise ValueError('Plan must contain at least one color rule or icon effect.')
    seen = set()
    for rule in rules:
        if not isinstance(rule, dict):
            raise ValueError('Each color rule must be an object.')
        rule['from'], rule['to'] = color(rule['from']), color(rule['to'])
        if rule.get('scope') not in ('text', 'shape'):
            raise ValueError('Color rule scope must be text or shape.')
        if not isinstance(rule.get('parts'), list) or not rule['parts']:
            raise ValueError('Each color rule needs an explicit nonempty parts list.')
        for part in rule['parts']:
            if not isinstance(part, str) or not SLIDE.fullmatch(part):
                raise ValueError('Only exact slide part paths are supported.')
            key = (part, rule['scope'], rule['from'])
            if key in seen:
                raise ValueError('Duplicate or conflicting color rule.')
            seen.add(key)
    seen = set()
    for effect in effects + protected:
        if not isinstance(effect, dict):
            raise ValueError('Each picture selector must be an object.')
        if not isinstance(effect.get('part'), str) or not SLIDE.fullmatch(effect['part']):
            raise ValueError('Picture selectors need an exact slide part.')
        if not re.fullmatch(r'[1-9][0-9]*', str(effect.get('shapeId', ''))):
            raise ValueError('Picture selectors need a numeric shapeId.')
        key = (effect['part'], str(effect['shapeId']))
        if key in seen:
            raise ValueError('Duplicate or protected picture selector.')
        seen.add(key)
    for effect in effects:
        if effect.get('assetClass') != 'generic-icon':
            raise ValueError('Only explicitly classified generic-icon pictures may be recolored.')
        effect['dark'], effect['light'] = color(effect['dark']), color(effect['light'])
    return plan


def edit_part(raw, part, plan):
    parsed = nodes(raw)
    replacements, changes = [], []
    rules = [r for r in plan.get('colorRules', []) if part in r['parts']]
    matched = set()
    for node in parsed:
        if node['name'] != A + '}srgbClr':
            continue
        # Never touch picture effects or embedded artwork through broad color rules.
        if ancestor(node, A + '}blip') or ancestor(node, P + '}pic'):
            continue
        is_text = any(ancestor(node, A + '}' + kind) for kind in ('rPr', 'defRPr', 'endParaRPr'))
        scope = 'text' if is_text else 'shape'
        if not is_text and not any(ancestor(node, A + '}' + kind) or ancestor(node, P + '}' + kind)
                                   for kind in ('spPr', 'tcPr')):
            continue
        value = node['attrs'].get('val', '').upper()
        for rule in rules:
            if rule['scope'] != scope or rule['from'] != value:
                continue
            tag = raw[node['start']:node['openEnd']]
            pattern = br'(\bval\s*=\s*[\x22\x27])' + value.encode() + br'([\x22\x27])'
            new, count = re.subn(pattern, lambda m: m[1] + rule['to'].encode() + m[2], tag, flags=re.I)
            if count != 1:
                raise ValueError('Cannot locate exact RGB attribute.')
            replacements.append((node['start'], node['openEnd'], new))
            changes.append({'part': part, 'scope': scope, 'from': value, 'to': rule['to']})
            matched.add(id(rule))
            break  # Simultaneous mappings: A->B and B->C do not cascade.
    for rule in rules:
        if id(rule) not in matched:
            raise ValueError(f'Color rule matched nothing in {part}: {rule["scope"]} {rule["from"]}')
    pictures = []
    for node in parsed:
        if node['name'] == P + '}pic':
            ids = [n for n in parsed if n['name'] == P + '}cNvPr' and ancestor(n, P + '}pic') is node]
            if len(ids) == 1:
                pictures.append((ids[0]['attrs'].get('id'), node))
    for selector in [*plan.get('iconEffects', []), *plan.get('protectedPictures', [])]:
        if selector['part'] != part:
            continue
        found = [n for ident, n in pictures if ident == str(selector['shapeId'])]
        if len(found) != 1:
            raise ValueError(f'Picture selector does not resolve uniquely: {part} {selector["shapeId"]}')
        if selector in plan.get('protectedPictures', []):
            continue
        blips = [n for n in parsed if n['name'] == A + '}blip' and ancestor(n, P + '}pic') is found[0]]
        if (len(blips) != 1 or R + '}embed' not in blips[0]['attrs']
                or R + '}link' in blips[0]['attrs']):
            raise ValueError('Icon must have one embedded image, not a linked picture.')
        node = blips[0]
        if node['children']:
            raise ValueError('Icon already has effects or extensions; review it rather than stacking effects.')
        effect = (f'<a:duotone xmlns:a="{A}"><a:srgbClr val="{selector["dark"]}"/>'
                  f'<a:srgbClr val="{selector["light"]}"/></a:duotone>').encode()
        tag = raw[node['start']:node['openEnd']]
        if node['selfClosing']:
            qname = re.match(br'<([^\s/>]+)', tag)[1]
            new = re.sub(br'/\s*>$', b'>', tag) + effect + b'</' + qname + b'>'
        else:
            new = tag + effect + raw[node['closeStart']:node['end']]
        replacements.append((node['start'], node['end'], new))
        changes.append({'part': part, 'scope': 'picture-effect', 'shapeId': str(selector['shapeId']),
                        'dark': selector['dark'], 'light': selector['light']})
    for start, end, new in sorted(replacements, reverse=True):
        raw = raw[:start] + new + raw[end:]
    nodes(raw)  # Validate edited XML before writing anything.
    return raw, changes


def apply_palette(source, output, plan, report=None, dry_run=False):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output:
        raise ValueError('Output must differ from the source.')
    if source.suffix.lower() != '.pptx' or output.suffix.lower() != '.pptx':
        raise ValueError('Source and output must be .pptx files.')
    if output.exists():
        raise ValueError('Output already exists; choose a new filename.')
    if report and (Path(report).resolve() in (source, output) or Path(report).exists()):
        raise ValueError('Report must be a new path distinct from source and output.')
    plan = validate_plan(plan)
    original = source.read_bytes()
    changes, payload = [], []
    targeted = {part for rule in plan.get('colorRules', []) for part in rule['parts']}
    targeted |= {r['part'] for r in plan.get('iconEffects', []) + plan.get('protectedPictures', [])}
    with zipfile.ZipFile(source) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Duplicate ZIP entries are unsupported.')
        if 'ppt/presentation.xml' not in names or '[Content_Types].xml' not in names:
            raise ValueError('Input is not a PowerPoint Open XML package.')
        if any(n.startswith('_xmlsignatures/') for n in names):
            raise ValueError('Digitally signed packages require a separate signing workflow.')
        if not targeted.issubset(names):
            raise ValueError('Plan references a missing slide part.')
        for info in archive.infolist():
            raw = archive.read(info)
            if info.filename in targeted:
                raw, delta = edit_part(raw, info.filename, plan)
                changes.extend(delta)
            payload.append((info, raw))
        comment = archive.comment
    result = {'schemaVersion': '1.0', 'sourceSha256': hashlib.sha256(original).hexdigest(),
              'dryRun': dry_run, 'changes': changes, 'changedParts': sorted({c['part'] for c in changes}),
              'preserved': 'All bytes outside selected RGB attributes and picture-effect elements; all media unchanged.',
              'limitations': 'Explicit slide RGB only; scheme colors and inherited styles unchanged. Visual review required.'}
    if source.read_bytes() != original:
        raise ValueError('Source changed during processing.')
    if not dry_run:
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open('xb') as handle:
            with zipfile.ZipFile(handle, 'w') as archive:
                archive.comment = comment
                for info, raw in payload:
                    archive.writestr(info, raw)
        result['finalSha256'] = hashlib.sha256(output.read_bytes()).hexdigest()
    if report:
        with Path(report).open('x', encoding='utf-8') as handle:
            json.dump(result, handle, ensure_ascii=False, indent=2)
            handle.write('\n')
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', required=True)
    parser.add_argument('--output', required=True)
    parser.add_argument('--plan', required=True)
    parser.add_argument('--report')
    parser.add_argument('--dry-run', action='store_true')
    args = parser.parse_args()
    try:
        plan = json.loads(Path(args.plan).read_text(encoding='utf-8-sig'))
        result = apply_palette(args.source, args.output, plan, args.report, args.dry_run)
        print(json.dumps(result, ensure_ascii=True))
    except (OSError, ValueError, KeyError, TypeError, zipfile.BadZipFile, expat.ExpatError) as error:
        parser.exit(2, f'ERROR: {error}\n')


if __name__ == '__main__':
    main()
