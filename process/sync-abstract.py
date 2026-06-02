#!/usr/bin/env python3
"""从 md 摘要同步 ABSTRACTCN/EN、KWDCN/EN 到 config.json。"""
import re, json, sys

def extract_body(text: str) -> str:
    """去除标题行和关键词行，保留段落结构（\n\n 分隔）。"""
    text = re.sub(r'^# .+\n', '', text)
    text = re.sub(r'\*\*Keywords?\*\*.*', '', text, flags=re.M | re.I)
    text = re.sub(r'\*\*关键词\*\*.*', '', text, flags=re.M)
    # 保留空行作为段落分隔
    paras = []
    for para in text.split('\n\n'):
        para = ' '.join(l.strip() for l in para.split('\n') if l.strip())
        if para:
            paras.append(para)
    return '\a'.join(paras)

def extract_kwd(text: str) -> str:
    """提取关键词。"""
    for pattern in [r'\*\*关键词\*\*[：:]\s*(.+)', r'\*\*Keywords?\*\*[：:]\s*(.+)']:
        m = re.search(pattern, text, re.M | re.I)
        if m:
            return m.group(1)
    return ''

config_path = sys.argv[1]
cn_path = sys.argv[2]
en_path = sys.argv[3]

try:
    with open(config_path) as f:
        config = json.load(f)
except FileNotFoundError:
    config = {
        'TITLE_EN': '', 'TITLE_CN': '',
        'SID': '', 'STNAME': '', 'INSTRUCTOR': '',
        'MAJOR': '', 'SUBJECT': '', 'CNDATE': '',
        'ABSTRACTCN': '', 'ABSTRACTEN': '',
        'KWDCN': '', 'KWDEN': ''
    }

cn = open(cn_path).read()
en = open(en_path).read()

config['ABSTRACTCN'] = extract_body(cn)
config['KWDCN'] = extract_kwd(cn)
config['ABSTRACTEN'] = extract_body(en)
config['KWDEN'] = extract_kwd(en)

with open(config_path, 'w') as f:
    json.dump(config, f, ensure_ascii=False, indent=4)
    f.write('\n')

print(f'  SYNC  {cn_path} + {en_path} -> {config_path}')
