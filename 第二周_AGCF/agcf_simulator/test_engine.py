import copy
import re
import unittest
from engine import Message, explore, load_json, sample, simulate


class SimulationTests(unittest.TestCase):
    def setUp(self):
        self.config = load_json('demo_profiles.json')

    def run_case(self, raw=None, **kwargs):
        return simulate({'invite': raw if raw is not None else sample(), 'profiles': self.config, **kwargs})

    def test_normal_preserves_call_identity_and_callee(self):
        result = self.run_case()
        self.assertTrue(result['ok'], result.get('error'))
        before, after = Message(sample()), result['output']
        self.assertTrue(after.startswith(before.start + '\r\n'))
        for key in ['From', 'To', 'Call-ID', 'CSeq', 'Contact']:
            self.assertIn(f'{key}: {before.get(key)}\r\n', after)
        self.assertIn('P-Asserted-Identity: <sip:+862132010323@sh2.ctcims.cn>', after)
        self.assertNotIn('P-Profile-Key:', after)
        self.assertIn('orig-ioi=PBX38.sh2.ctcims.cn', after)
        self.assertIn('Max-Forwards: 69', after)
        self.assertEqual(len(re.findall(r'^Via:', after, re.M)), 2)
        self.assertIn('scscf-demo.example.invalid', result['next_hop'])
        self.assertEqual(len(result['steps']), 7)

    def test_real_pui_provenance_and_synthetic_binding_are_separate(self):
        result = self.run_case()
        identity_step = result['steps'][2]
        self.assertEqual(identity_step['data']['pui_evidence'][0]['did'], 'BRA')
        self.assertIn('XJSL_AGCFB04', identity_step['data']['pui_evidence'][0]['source_file'])
        self.assertEqual(identity_step['origin'], 'assumption')
        self.assertIn('人工', identity_step['detail'])

    def test_missing_registration_does_not_fallback_to_icscf(self):
        self.config['users'][0]['registered'] = False
        result = self.run_case()
        self.assertFalse(result['ok'])
        self.assertIsNone(result['output'])
        self.assertIsNone(result['next_hop'])
        self.assertIn('不回退', result['error'])

    def test_unknown_and_duplicate_identity_fail_closed(self):
        self.assertFalse(self.run_case(sample('+862199999999'))['ok'])
        self.config['users'].append(copy.deepcopy(self.config['users'][0]))
        self.assertFalse(self.run_case()['ok'])

    def test_existing_pcv_is_preserved(self):
        value = 'icid-value=existing-123;orig-ioi=old.example.invalid'
        raw = sample().replace('Content-Length:', f'P-Charging-Vector: {value}\r\nContent-Length:')
        result = self.run_case(raw)
        self.assertTrue(result['ok'])
        self.assertIn(f'P-Charging-Vector: {value}\r\n', result['output'])
        self.assertEqual(result['output'].count('P-Charging-Vector:'), 1)

    def test_existing_pai_replaced(self):
        raw = sample().replace('Content-Length:', 'p-asserted-identity: <sip:old@invalid>\r\nContent-Length:')
        result = self.run_case(raw)
        self.assertTrue(result['ok'])
        self.assertNotIn('old@invalid', result['output'])

    def test_unique_icid_per_new_simulation(self):
        a, b = self.run_case(), self.run_case()
        self.assertNotEqual(re.search(r'icid-value=([^;]+)', a['output'])[1], re.search(r'icid-value=([^;]+)', b['output'])[1])

    def test_ioi_mapping_can_be_disabled(self):
        self.config['ioi_to_orig_ioi'] = False
        self.assertNotIn('orig-ioi=', self.run_case()['output'])

    def test_body_is_preserved_with_utf8_byte_length(self):
        body = 'v=0\r\ns=演示\r\n'
        raw = sample().replace('Content-Length: 0', f'Content-Type: application/sdp\r\nContent-Length: {len(body.encode("utf-8"))}') + body
        result = self.run_case(raw)
        self.assertTrue(result['ok'], result.get('error'))
        self.assertEqual(result['output'].split('\r\n\r\n', 1)[1], body)

    def test_bad_lengths_and_duplicate_headers_are_rejected(self):
        for raw in [sample()+'extra', sample().replace('Content-Length: 0','Content-Length: 0\r\nl: 0'), sample().replace('Max-Forwards: 70','Max-Forwards: 70\r\nMax-Forwards: 60')]:
            with self.subTest(raw=raw):
                self.assertFalse(self.run_case(raw)['ok'])

    def test_hop_limit_and_default(self):
        self.assertFalse(self.run_case(sample().replace('Max-Forwards: 70', 'Max-Forwards: 0'))['ok'])
        self.assertFalse(self.run_case(sample().replace('Max-Forwards: 70', 'Max-Forwards: 00'))['ok'])
        self.assertIn('Max-Forwards: 0', self.run_case(sample().replace('Max-Forwards: 70', 'Max-Forwards: 1'))['output'])
        self.assertIn('Max-Forwards: 70', self.run_case(sample().replace('Max-Forwards: 70\r\n', ''))['output'])

    def test_response_bye_and_reinvite_not_supported(self):
        for raw in ['SIP/2.0 200 OK\r\n\r\n', sample().replace('INVITE', 'BYE'), sample().replace('To: <tel:13800138000>', 'To: <tel:13800138000>;tag=established')]:
            self.assertFalse(self.run_case(raw)['ok'])

    def test_compact_header_names_and_lf_accepted(self):
        raw = sample().replace('From:', 'f:').replace('To:', 't:').replace('Call-ID:', 'i:').replace('Via:', 'v:').replace('Content-Length:', 'l:').replace('\r\n','\n')
        self.assertTrue(self.run_case(raw)['ok'])

    def test_user_config_does_not_mutate_and_no_header_injection(self):
        backup = copy.deepcopy(self.config)
        self.run_case()
        self.assertEqual(self.config, backup)
        self.config['users'][0]['pai'] = 'sip:x@invalid\r\nInjected: true'
        self.assertFalse(self.run_case()['ok'])

    def test_bad_config_types(self):
        self.config['users'][0]['registered'] = 'false'
        self.assertFalse(self.run_case()['ok'])
        self.assertFalse(simulate({'invite': sample(), 'profiles': []})['ok'])

    def test_unknown_trunk_and_unsupported_ppk(self):
        self.assertFalse(self.run_case(trunk='999')['ok'])
        raw = sample().replace('Content-Length:', 'P-Profile-Key: <sip:old@invalid>\r\nContent-Length:')
        self.assertFalse(self.run_case(raw)['ok'])

    def test_route_set_is_not_silently_overwritten(self):
        raw = sample().replace('Content-Length:', 'Route: <sip:previous.example.invalid;lr>\r\nContent-Length:')
        self.assertFalse(self.run_case(raw)['ok'])

    def test_failure_export_preserves_input_and_configuration(self):
        self.config['users'][0]['registered'] = False
        result = self.run_case()
        self.assertEqual(result['input'], sample())
        self.assertEqual(result['profiles'], self.config)
        self.assertEqual(len(result['provenance']), 4)

    def test_real_evidence_longest_prefix_and_no_signal_output(self):
        result = explore('13800138000')
        self.assertTrue(result['ok'])
        self.assertEqual(result['rule']['pfx'], "K'13")
        self.assertEqual(result['rule']['rsc'], '10')
        self.assertEqual(result['routes'][0]['r'], '10')
        self.assertEqual(result['length_status'], 'within')
        self.assertNotIn('output', result)

    def test_length_boundaries_and_missing_prefix(self):
        self.assertEqual(explore('13')['length_status'], 'short')
        self.assertEqual(explore('138001380001')['length_status'], 'long')
        self.assertFalse(explore('abc')['ok'])
        self.assertFalse(explore('')['ok'])
        self.assertEqual(explore('385350123456')['rule']['pfx'], "K'385350")


if __name__ == '__main__':
    unittest.main(verbosity=2)
