"""Synthetic package tests; no customer documents or runtime packages required."""
import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import sys
import unittest
import zipfile

sys.dont_write_bytecode = True

SPEC = importlib.util.spec_from_file_location('apply_palette', Path(__file__).parents[1] / 'scripts' / 'apply_palette.py')
M = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(M)
PART = 'ppt/slides/slide1.xml'
XML = f'''<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="{M.P}" xmlns:x="{M.A}" xmlns:r="{M.R}">
<p:cSld><p:spTree><p:sp><p:spPr><x:solidFill><x:srgbClr val='EC671B'/></x:solidFill></p:spPr>
<p:txBody><x:p><x:r><x:rPr><x:solidFill><x:srgbClr val="EC671B"/></x:solidFill></x:rPr>
<x:t>中文 123.45万条 &amp; 49项</x:t></x:r></x:p></p:txBody></p:sp>
<p:pic><p:nvPicPr><p:cNvPr id="7" name="generic icon"/></p:nvPicPr>
<p:blipFill><x:blip r:embed="rId1"/></p:blipFill><p:spPr><x:xfrm><x:off x="123" y="456"/></x:xfrm></p:spPr></p:pic>
<p:pic><p:nvPicPr><p:cNvPr id="8" name="evidence"/></p:nvPicPr><p:blipFill><x:blip r:embed="rId2"/></p:blipFill></p:pic>
</p:spTree></p:cSld></p:sld>'''.encode()


def rule(scope='text', before='EC671B', after='FCB824'):
    return {'from': before, 'to': after, 'scope': scope, 'parts': [PART]}


def plan(*rules):
    return {'schemaVersion': '1.0', 'colorRules': list(rules)}


class PaletteTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source, self.output = self.root / '源.pptx', self.root / 'final.pptx'
        with zipfile.ZipFile(self.source, 'w') as z:
            z.comment = b'archive metadata'
            for name, value in {PART: XML, '[Content_Types].xml': b'<Types/>',
                                'ppt/presentation.xml': b'<presentation/>',
                                'ppt/media/icon.png': b'unchanged binary\x00\xff',
                                'ppt/slides/_rels/slide1.xml.rels': b'<Relationships/>',
                                'ppt/notesSlides/notesSlide1.xml': b'<notes>private source</notes>',
                                'ppt/slideMasters/slideMaster1.xml': b'<master/>'}.items():
                z.writestr(name, value)
        self.original = self.source.read_bytes()

    def apply(self, config, **kwargs):
        return M.apply_palette(self.source, self.output, copy.deepcopy(config), **kwargs)

    def test_text_mapping_preserves_all_other_bytes(self):
        result = self.apply(plan(rule()), report=self.root / 'report.json')
        with zipfile.ZipFile(self.source) as a, zipfile.ZipFile(self.output) as b:
            self.assertEqual(a.namelist(), b.namelist())
            self.assertEqual(a.comment, b.comment)
            for name in a.namelist():
                expected = XML.replace(b'val="EC671B"', b'val="FCB824"') if name == PART else a.read(name)
                self.assertEqual(b.read(name), expected)
        self.assertEqual(self.source.read_bytes(), self.original)
        self.assertEqual(result['finalSha256'], M.hashlib.sha256(self.output.read_bytes()).hexdigest())
        self.assertEqual(json.loads((self.root / 'report.json').read_text())['changes'], result['changes'])

    def test_shape_mapping_keeps_emphasized_text(self):
        self.apply(plan(rule('shape', after='146CB5')))
        with zipfile.ZipFile(self.output) as z:
            self.assertEqual(z.read(PART), XML.replace(b"val='EC671B'", b"val='146CB5'"))

    def test_rgb_mappings_do_not_cascade(self):
        raw = XML.replace(b'val="EC671B"', b'val="FCB824"')
        config = M.validate_plan(plan(rule('shape', after='FCB824'), rule('text', 'FCB824', '103D6B')))
        edited, changes = M.edit_part(raw, PART, config)
        self.assertIn(b"val='FCB824'", edited)
        self.assertIn(b'val="103D6B"', edited)
        self.assertEqual(len(changes), 2)

    def test_icon_effect_is_bounded_and_media_unchanged(self):
        config = plan()
        config['iconEffects'] = [{'part': PART, 'shapeId': 7, 'assetClass': 'generic-icon', 'dark': '#0064AA', 'light': '#FFFFFF'}]
        config['protectedPictures'] = [{'part': PART, 'shapeId': 8}]
        self.apply(config)
        with zipfile.ZipFile(self.output) as z:
            edited = z.read(PART)
            self.assertIn(b'<x:blip r:embed="rId1"><a:duotone', edited)
            self.assertIn(b'<x:blip r:embed="rId2"/>', edited)
            self.assertEqual(z.read('ppt/media/icon.png'), b'unchanged binary\x00\xff')
            self.assertEqual(len(M.nodes(edited)) - len(M.nodes(XML)), 3)

    def test_dry_run_writes_only_report(self):
        result = self.apply(plan(rule()), dry_run=True, report=self.root / 'dry.json')
        self.assertFalse(self.output.exists())
        self.assertTrue(result['dryRun'])
        self.assertEqual(len(result['changes']), 1)

    def test_bad_plan_or_selector_never_writes_output(self):
        cases = [plan(rule(after='bad')), plan(rule(before='ABCDEF')), plan(rule(), rule()),
                 plan({'from': 'EC671B', 'to': 'FCB824', 'scope': 'all', 'parts': [PART]}),
                 plan({'from': 'EC671B', 'to': 'FCB824', 'scope': 'text', 'parts': ['ppt/theme/theme1.xml']}),
                 {'schemaVersion': '1.0', 'iconEffects': [{'part': PART, 'shapeId': 99, 'assetClass': 'generic-icon', 'dark': '0064AA', 'light': 'FFFFFF'}]},
                 {'schemaVersion': '1.0', 'iconEffects': [{'part': PART, 'shapeId': 7, 'assetClass': 'logo', 'dark': '0064AA', 'light': 'FFFFFF'}]}]
        for config in cases:
            with self.subTest(config=config), self.assertRaises(ValueError):
                self.apply(config)
            self.assertFalse(self.output.exists())

    def test_refuses_source_output_or_report_overwrite(self):
        self.output.write_bytes(b'existing')
        with self.assertRaises(ValueError):
            self.apply(plan(rule()))
        self.assertEqual(self.output.read_bytes(), b'existing')
        with self.assertRaises(ValueError):
            M.apply_palette(self.source, self.source, plan(rule()))
        self.output.unlink()
        with self.assertRaises(ValueError):
            self.apply(plan(rule()), report=self.source)
        self.assertEqual(self.source.read_bytes(), self.original)

    def test_existing_effects_linked_images_and_protected_overlap_rejected(self):
        effect = {'part': PART, 'shapeId': 7, 'assetClass': 'generic-icon', 'dark': '0064AA', 'light': 'FFFFFF'}
        config = {'schemaVersion': '1.0', 'iconEffects': [effect]}
        for raw in (XML.replace(b'<x:blip r:embed="rId1"/>', b'<x:blip r:embed="rId1"><x:grayscl/></x:blip>'),
                    XML.replace(b'r:embed="rId1"', b'r:link="rId1"')):
            with self.assertRaises(ValueError):
                M.edit_part(raw, PART, M.validate_plan(copy.deepcopy(config)))
        config['protectedPictures'] = [{'part': PART, 'shapeId': 7}]
        with self.assertRaises(ValueError):
            M.validate_plan(config)

    def test_signed_and_duplicate_packages_are_rejected(self):
        for extra in ('_xmlsignatures/sig1.xml', PART):
            with zipfile.ZipFile(self.source, 'a') as z:
                import warnings
                with warnings.catch_warnings():
                    warnings.simplefilter('ignore', UserWarning)
                    z.writestr(extra, b'<signature/>')
            with self.assertRaises(ValueError):
                self.apply(plan(rule()))
            self.assertFalse(self.output.exists())
            self.source.write_bytes(self.original)

    def test_invalid_json_shapes_are_rejected(self):
        for config in ([], {'schemaVersion': '1.0', 'colorRules': [None]},
                       {'schemaVersion': '1.0', 'iconEffects': ['not a selector']},
                       {'schemaVersion': '1.0', 'colorRules': 'not an array'}):
            with self.subTest(config=config), self.assertRaises(ValueError):
                self.apply(config)
            self.assertFalse(self.output.exists())

    def test_xml_namespaces_entities_and_scheme_colors(self):
        raw = XML.replace(b'<x:srgbClr val="EC671B"/>', b'<x:schemeClr val="accent2"/>')
        with self.assertRaisesRegex(ValueError, 'matched nothing'):
            M.edit_part(raw, PART, M.validate_plan(plan(rule())))
        with self.assertRaisesRegex(ValueError, 'DTD'):
            M.nodes(b'<!DOCTYPE x [<!ENTITY x "oops">]><x>&x;</x>')


if __name__ == '__main__':
    unittest.main()
