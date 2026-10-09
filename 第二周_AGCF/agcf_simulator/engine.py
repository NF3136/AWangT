"""有界的离线 INVITE 规则演示，不是 SIP 协议栈或真实网元。"""
import copy
import difflib
import json
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

BASE = Path(__file__).resolve().parent
ALIASES = {'f': 'from', 't': 'to', 'i': 'call-id', 'v': 'via', 'l': 'content-length', 'c': 'content-type', 'm': 'contact'}


def load_json(name):
    return json.loads((BASE / 'data' / name).read_text(encoding='utf-8'))


def canonical(name):
    return ALIASES.get(name.lower(), name.lower())


class SimulationError(ValueError):
    pass


class Message:
    def __init__(self, raw):
        if not isinstance(raw, str) or not raw.strip():
            raise SimulationError('请输入 SIP INVITE 报文。')
        if len(raw.encode('utf-8')) > 128000 or '\x00' in raw:
            raise SimulationError('报文过大或包含非法空字符。')
        parts = re.split(r'\r?\n\r?\n', raw, maxsplit=1)
        head, self.body = parts[0], parts[1] if len(parts) == 2 else ''
        lines = head.replace('\r\n', '\n').split('\n')
        self.start = lines[0]
        match = re.fullmatch(r'INVITE\s+((?:sip:|sips:|tel:)[^\s]+)\s+SIP/2\.0', self.start)
        if not match:
            raise SimulationError('本版只处理初始 INVITE 请求；响应、ACK、BYE 等不在本版范围。')
        self.uri = match[1]
        self.headers = []
        for line in lines[1:]:
            if line.startswith((' ', '\t')) and self.headers:
                self.headers[-1][1] += ' ' + line.strip()
                continue
            if ':' not in line:
                raise SimulationError('头域格式错误：每行需为“名称: 值”。')
            name, value = line.split(':', 1)
            if not re.fullmatch(r'[A-Za-z0-9!#$%&\x27*+.^_`|~-]+', name):
                raise SimulationError('头域名称格式错误。')
            self.headers.append([name, value.strip()])
        for key in ['from', 'to', 'call-id', 'cseq', 'via']:
            values = self.all(key)
            if not values or not all(values):
                raise SimulationError(f'缺少必要头域：{key}')
            if key != 'via' and len(values) != 1:
                raise SimulationError(f'必要单值头域重复：{key}')
        for key in ['content-length', 'max-forwards']:
            if len(self.all(key)) > 1:
                raise SimulationError(f'头域重复：{key}')
        if not re.fullmatch(r'\d+\s+INVITE', self.get('cseq')):
            raise SimulationError('CSeq 必须与 INVITE 方法一致。')
        if re.search(r';\s*tag\s*=', self.get('to'), re.I):
            raise SimulationError('本版不处理对话内 re-INVITE（To 带 tag）。')
        if self.all('route'):
            raise SimulationError('本版尚未实现已有 Route 路由集的处理，请使用不含 Route 的初始演示请求。')
        if self.all('p-profile-key'):
            raise SimulationError('本版只处理 PAI 和 PCV；已有 PPK 的请求留待后续版本支持。')
        length = self.get('content-length')
        if length is not None and (not length.isdigit() or int(length) != len(self.body.encode('utf-8'))):
            raise SimulationError('Content-Length 与消息体的 UTF-8 字节数不一致。')
        hops = self.get('max-forwards')
        if hops is not None and (not hops.isdigit() or not 0 <= int(hops) <= 255):
            raise SimulationError('Max-Forwards 必须是 0～255 的整数。')
        if hops is not None and int(hops) == 0:
            raise SimulationError('Max-Forwards 为 0，停止模拟转发。')

    def all(self, key):
        return [v for k, v in self.headers if canonical(k) == key.lower()]

    def get(self, key):
        values = self.all(key)
        return values[0] if values else None

    def set(self, key, value):
        self.headers = [[k, v] for k, v in self.headers if canonical(k) != key.lower()]
        if value is not None:
            self.headers.append([key, value])

    def render(self):
        return self.start + '\r\n' + '\r\n'.join(f'{k}: {v}' for k, v in self.headers) + '\r\n\r\n' + self.body


def safe_text(value, label):
    if not isinstance(value, str) or not value or len(value) > 500 or any(ord(c) < 32 or ord(c) == 127 for c in value):
        raise SimulationError(f'{label} 为空或含非法字符。')
    return value


def sip_uri(value, label):
    safe_text(value, label)
    if not re.fullmatch(r'sip:[^\s<>"\\]+', value):
        raise SimulationError(f'{label} 必须是无尖括号的 sip: URI。')
    return value


def caller_from(value):
    # 演示范围：From 中的数字型 sip/tel 身份；不把 Request-URI 当成主叫。
    match = re.search(r'(?:<\s*|^)(?:sip:|tel:)(\+?\d+)(?=@|;|>|$)', value, re.I)
    if not match:
        raise SimulationError('无法提取数字型 From 主叫；本版不推断匿名或复杂身份。')
    return match[1]


def simulate(payload, snapshot=None):
    snapshot = snapshot or load_json('real_config.json')
    steps, changes, warnings = [], [], []
    result = {'ok': False, 'output': None, 'next_hop': None, 'steps': steps, 'changes': changes,
              'warnings': warnings, 'mode': 'registered_demo', 'generated_at': datetime.now(timezone.utc).isoformat(),
              'input': payload.get('invite', ''), 'profiles': copy.deepcopy(payload.get('profiles', load_json('demo_profiles.json'))),
              'trunk': str(payload.get('trunk', '38')), 'provenance': snapshot.get('provenance', []),
              'source_file': snapshot['source_file'], 'assumptions': [
                  '按入中继 38 人工绑定场景；未实现接收地址四元组识别。',
                  'PUI 可引用真实样本值；主叫与中继的绑定、注册状态和 S-CSCF 为测试配置。',
                  '默认 PUI 来自 B04 BRA 用户，与 JSL TG38 为跨来源人工绑定，不是同一次现场呼叫。',
                  '本版只处理 PAI 与 PCV，未实现 PPK。',
                  '本场景不执行库中的静态号码分析或 DNC 变换。',
                  '采用简化代理转发规则，未实现 AGCF 的 B2BUA、事务、重传、媒体和完整呼叫状态。',
              ]}

    def step(title, detail, origin, data=None):
        steps.append({'title': title, 'detail': detail, 'origin': origin, 'status': 'done', 'data': data or {}})

    def update(key, value, reason, origin):
        before = msg.all(key)
        msg.set(key, value)
        after = msg.all(key)
        if before != after:
            changes.append({'field': key, 'before': before, 'after': after, 'reason': reason, 'origin': origin})

    try:
        msg = Message(payload.get('invite', ''))
        caller = caller_from(msg.get('from'))
        step('解析输入信令', f'INVITE 请求；主叫 {caller}；请求目标 {msg.uri}', 'input', {'caller': caller, 'request_uri': msg.uri})
        trunk_id = str(payload.get('trunk', '38'))
        trunks = [r for r in snapshot['tables']['t_agcf_siptg'] if r['tg'] == trunk_id]
        if trunk_id != '38' or len(trunks) != 1:
            raise SimulationError('本版只打包并支持入中继 38 的演示场景。')
        trunk = trunks[0]
        step('读取中继配置', f'中继 {trunk_id}；IOI = {trunk["ioi"]}', 'real', trunk)
        config = copy.deepcopy(payload.get('profiles', load_json('demo_profiles.json')))
        if not isinstance(config, dict) or not isinstance(config.get('users'), list):
            raise SimulationError('测试配置必须包含 users 数组。')
        users = [r for r in config['users'] if isinstance(r, dict) and r.get('caller') == caller]
        if len(users) != 1:
            raise SimulationError('未找到唯一的主叫测试身份。请补充用户配置；不能从中继 IOI 猜测 PAI。')
        user = users[0]
        evidence = [r for r in snapshot.get('identity_samples', []) if r['pui'] == user.get('pai')]
        step('查找用户身份', '引用真实 PUI 样本；与入中继的绑定为人工构造。' if evidence else '按主叫精确命中一条自定义测试身份。',
             'assumption' if evidence else 'mock', {'binding': user, 'pui_evidence': evidence})
        if type(user.get('registered')) is not bool:
            raise SimulationError('registered 必须为 true 或 false。')
        if user['registered'] is not True or not user.get('scscf'):
            raise SimulationError('测试记录中没有有效的注册 S-CSCF。停止处理，不回退到静态 I-CSCF 路由。')
        next_hop = sip_uri(user['scscf'], '注册 S-CSCF')
        if ';lr' not in next_hop:
            raise SimulationError('本版下一跳只支持带 ;lr 的宽松路由 URI。')
        step('确认注册路由', f'本次使用测试注册记录：{next_hop}', 'mock', {'next_hop': next_hop})
        pai = sip_uri(user.get('pai'), 'PAI')
        update('P-Asserted-Identity', f'<{pai}>', '演示策略：采用人工配置映射，替换输入中已有 PAI。', 'assumption')
        step('补充身份头域', '已按人工配置映射设置 PAI。本版不生成 PPK。', 'assumption', {'pai': pai, 'pui_evidence': evidence})
        pcvs = msg.all('p-charging-vector')
        if pcvs:
            if len(pcvs) != 1 or not re.search(r'(?:^|;)\s*icid-value\s*=\s*(?:"[^"\r\n]+"|[^;\s]+)', pcvs[0], re.I):
                raise SimulationError('已有 PCV 缺少有效 icid-value 或存在重复；本版停止处理。')
            detail = '保留已有 PCV（本版策略），避免无条件覆盖计费关联。'
            warnings.append('已有 PCV 按本版策略原样保留；未验证可信域与运营商策略。')
        else:
            host = safe_text(config.get('icid_host'), 'ICID 生成地址')
            if not re.fullmatch(r'[A-Za-z0-9.:-]+', host):
                raise SimulationError('ICID 生成地址格式不受支持。')
            pcv = f'icid-value=demo-{uuid.uuid4().hex};icid-generated-at={host}'
            if type(config.get('ioi_to_orig_ioi')) is not bool:
                raise SimulationError('ioi_to_orig_ioi 必须为 true 或 false。')
            if config['ioi_to_orig_ioi']:
                pcv += f';orig-ioi={trunk["ioi"]}'
                warnings.append('IOI → orig-ioi 是可关闭的演示映射；尚无该现场配置的报文验证。')
            update('P-Charging-Vector', pcv, '生成演示 ICID；不声称复现华为的生成算法。', 'assumption')
            detail = '生成新的演示 ICID；生成地址为测试值。'
        step('处理计费关联', detail, 'assumption', {'pcv': msg.get('p-charging-vector')})
        local = safe_text(config.get('local_address'), '模拟 AGCF 地址')
        if not re.fullmatch(r'[A-Za-z0-9.-]+:\d{1,5}', local) or not 1 <= int(local.rsplit(':', 1)[1]) <= 65535:
            raise SimulationError('模拟 AGCF 地址需为 host:port。')
        old_via = msg.all('via')
        via = f'SIP/2.0/UDP {local};branch=z9hG4bK-demo-{uuid.uuid4().hex}'
        msg.headers.insert(0, ['Via', via])
        changes.append({'field': 'Via', 'before': old_via, 'after': msg.all('via'), 'reason': '简化代理转发：增加本地 Via，保留原 Via。', 'origin': 'assumption'})
        hops = msg.get('max-forwards')
        update('Max-Forwards', str(int(hops) - 1) if hops is not None else '70', '转发跳数减一；缺省请求设置为 70。', 'protocol')
        update('Route', f'<{next_hop}>', '使用模拟注册下一跳；保持 Request-URI 和 To。', 'mock')
        update('Content-Length', str(len(msg.body.encode('utf-8'))), '按保留的消息体重新计算 UTF-8 字节长度。', 'protocol')
        step('生成输出与下一跳', '已生成离线结果；Request-URI、To、From、Call-ID、CSeq 和消息体保留。', 'protocol', {'next_hop': next_hop})
        result.update(ok=True, output=msg.render(), next_hop=next_hop, caller=caller, request_uri=msg.uri,
                      input=payload['invite'], profiles=config, trunk=trunk_id,
                      diff='\n'.join(difflib.unified_diff(payload['invite'].splitlines(), msg.render().splitlines(), fromfile='输入 INVITE', tofile='输出 INVITE', lineterm='')))
    except (SimulationError, TypeError, KeyError, AttributeError) as exc:
        error = str(exc) if isinstance(exc, SimulationError) else '配置结构不完整，请检查测试配置。'
        result['error'] = error
        steps.append({'title': '停止处理', 'detail': error, 'origin': 'validation', 'status': 'error', 'data': {}})
    return result


def explore(number, snapshot=None):
    """静态配置证据查找，不能作为运行时路由决策或报文输出。"""
    snapshot = snapshot or load_json('real_config.json')
    if not isinstance(number, str) or not re.fullmatch(r'\d{1,32}', number):
        return {'ok': False, 'error': '请输入 1～32 位数字，用于查询被叫字冠。'}
    rows = snapshot['tables']['t_agcf_cnacld']
    matched = [r for r in rows if number.startswith(r['pfx'][2:] if r['pfx'].startswith("K'") else r['pfx'])]
    matched.sort(key=lambda r: len(r['pfx']), reverse=True)
    if not matched:
        return {'ok': False, 'error': '本快照中未匹配到字冠；这不等于完整设备必定拒绝呼叫。'}
    row = matched[0]
    minimum, maximum = int(row['minl']), int(row['maxl'])
    length_status = 'within' if minimum <= len(number) <= maximum else ('short' if len(number) < minimum else 'long')
    routes = [r for r in snapshot['tables']['t_agcf_rtana'] if r['rssc'] == '30' and r['rsc'] == row['rsc']]
    warnings = ['仅为配置关联结果，未执行实际注册、业务触发、主叫分析和链路状态选择。']
    if length_status == 'short':
        warnings.append('号码短于最小号长；按已提供说明应释放呼叫，不能按候选路由转发。')
    if length_status == 'long':
        warnings.append('号码超出最大号长；说明提到仅按最大号长分析，本页不擅自截断并输出信令。')
    if row.get('clraf') == 'YES':
        warnings.append('该字冠启用了主叫分析；DCP 与通配编码未确认，不能据此确定最终路由或主叫号码。')
    return {'ok': True, 'number': number, 'rule': row, 'length_status': length_status,
            'routes': routes, 'warnings': warnings, 'source_file': snapshot['source_file']}


def sample(caller='+862132010323'):
    return '\r\n'.join([
        'INVITE tel:13800138000 SIP/2.0',
        'Via: SIP/2.0/UDP 192.0.2.20:5060;branch=z9hG4bK-demo-input',
        'Max-Forwards: 70', f'From: <sip:{caller}@sh2.ctcims.cn>;tag=demo-caller',
        'To: <tel:13800138000>', 'Call-ID: demo-call-001@pbx.example.invalid',
        'CSeq: 1 INVITE', f'Contact: <sip:{caller}@192.0.2.20:5060>', 'Content-Length: 0', '', ''])
