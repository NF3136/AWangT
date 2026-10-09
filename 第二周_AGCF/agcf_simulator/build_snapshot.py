"""从用户导出的 CSV 提取同源证据；不连接数据库。"""
import csv
import hashlib
import json
from pathlib import Path

BASE = Path(__file__).resolve().parent
SOURCE = 'JSL_VOBB_AGCFOMU/ALLME_20260803065001.txt'


def build():
    tables, provenance = {}, []
    for name in ['同源链路配置.csv', '号码分析配置.csv', '主叫与号码处理配置.csv']:
        path = BASE.parent / name
        rows = list(csv.DictReader(path.open(encoding='utf-8-sig', newline='')))
        counts = {}
        for row in rows:
            if row['source_file'] != SOURCE:
                raise ValueError(f'{name}: 来源不一致')
            if row['record_kind'] == 'config':
                counts[row['table_name']] = counts.get(row['table_name'], 0) + 1
                tables.setdefault(row['table_name'], []).append(json.loads(row['config']))
        for row in rows:
            if row['record_kind'] == 'summary':
                if json.loads(row['config'])['row_count'] != counts.get(row['table_name'], 0):
                    raise ValueError(f'{name}: 导出不完整')
        provenance.append({'file': name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'counts': counts})
    # 仅打包本次演示相关数据，保留原始字段、ADD/MOD 和来源；不打包用户认证数据。
    filters = {
        't_agcf_siptg': lambda r: r['tg'] in ['38', '10', '11', '12', '13'],
        't_agcf_sipippair': lambda r: r['tg'] in ['38', '10', '11', '12', '13'],
        't_agcf_callsrc': lambda r: r['csc'] == '30',
        't_agcf_rtana': lambda r: r['rssc'] == '30',
        't_agcf_rt': lambda r: r['r'] == '10',
        't_agcf_srt': lambda r: r['src'] in ['5', '7', '38'],
        't_agcf_ldnset': lambda r: r['lp'] == '30',
        't_agcf_cnacld': lambda r: r['lp'] == '30',
        't_agcf_cnaclr': lambda r: r['csc'] == '30',
        't_agcf_dnc': lambda r: r['dcx'] in ['3', '4', '6'],
    }
    selected = {table: [r for r in tables.get(table, []) if predicate(r)] for table, predicate in filters.items()}
    identity_path = BASE.parent / '样例数据' / 'MSN 号码（主叫分机）new.csv'
    identities = list(csv.DictReader(identity_path.open(encoding='utf-8-sig', newline='')))
    provenance.append({'file': str(identity_path.relative_to(BASE.parent)),
                       'sha256': hashlib.sha256(identity_path.read_bytes()).hexdigest(),
                       'counts': {'ASBR 身份样本': len(identities)}})
    result = {'source_file': SOURCE, 'provenance': provenance, 'tables': selected, 'identity_samples': identities,
              'limitations': ['此快照为导入表，不是完整原始 MML。', 'DNC 缺少 DCP 字段；不执行号码变换。',
                              'CNACLR 中 E 串的通配编码尚待原始配置核对。', '没有实际用户注册状态；静态路由不等于注册的 S-CSCF。',
                              'PUI 样本来自 B04 的 BRA 用户，与本页 TG38 人工绑定；不是同一次现场呼叫。',
                              '第一版只处理 PAI 和 PCV；PPK 待后续补充。']}
    target = BASE / 'data' / 'real_config.json'
    target.parent.mkdir(exist_ok=True)
    target.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'快照已生成：{target}，{sum(map(len, selected.values()))} 条关联记录')


if __name__ == '__main__':
    build()
