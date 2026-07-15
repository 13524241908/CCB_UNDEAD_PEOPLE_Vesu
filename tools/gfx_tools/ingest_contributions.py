#!/usr/bin/env python3
"""
读取贴图美术师填好的 Excel（内嵌 png），校验合格后自动归置贴图。

流程：
  1. 解压 xlsx，提取每个工作表内嵌的图片及其锚定行
  2. 对每个"带图片"的行做合格校验（见下）
  3. 仅当全部带图行都合格时，才执行归置：
       - 新增模式（默认）：png 放到 pngs_<尺寸>/<类别子目录>/<完整贴图ID>.png，
         并生成同名 .json（CDDA 散图格式）
       - 替换模式（--replace）：全仓库搜同名 png，找到则原位覆盖官方贴图，
         不生成 json（官方已有定义）；被覆盖的原图备份内嵌到检测表末列。
         全仓库找不到同名时，回退为新增模式（生成 json）。
  4. 在脚本同目录输出"检测/贡献表"xlsx，逐行记录模式/通过/原因/归置路径/备份

合格判定（全部满足才通过）：
  - 完成状态 == "已完成"
  - 贡献者 非空
  - png 像素尺寸在 tile_info.json 定义的尺寸集合内
  （内嵌图片本身无文件名，落地时按该行「完整贴图ID」自动命名，无需名称比对）

只要有一行不合格，整体不归置（只出检测表）。可用 --dry-run 强制只检测。

用法：
  python3 tools/gfx_tools/ingest_contributions.py <填好的.xlsx>
  python3 tools/gfx_tools/ingest_contributions.py <填好的.xlsx> --dry-run
  python3 tools/gfx_tools/ingest_contributions.py <填好的.xlsx> --replace
"""
import argparse
import json
import os
import re
import shutil
import struct
import sys
import tempfile
import zipfile
from pathlib import Path


# ---- 工作表类别 -> 旧式子目录候选（按优先级，仅在已存在时复用） ----
# 当前仓库的 pngs_* 目录是平铺结构，resolve_subdir 会直接使用根目录。
CATEGORY_SUBDIRS = {
    '主贴图-物品': ['items', 'item'],
    '主贴图-怪物': ['monsters', 'monster'],
    '尸体': ['monsters', 'monster'],
    '主贴图-地形': ['terrain'],
    '主贴图-家具': ['furniture'],
    '主贴图-载具部件': ['vehicle'],
    '主贴图-陷阱场效': ['traps', 'fields', 'trap', 'field'],
    'overmap': ['overmap'],
    '地图事件': ['overmap/extra'],
    '穿戴-中性': ['overlay/worn'],
    '穿戴-男': ['overlay/worn'],
    '穿戴-女': ['overlay/worn'],
    '手持': ['overlay/wielded'],
    '手持尸体': ['overlay/wielded'],
    '变异生化叠加': ['overlay/mutations'],  # 兼容旧工作簿
    '变异生化-中性': ['overlay/mutations'],
    '变异生化-男': ['overlay/mutations'],
    '变异生化-女': ['overlay/mutations'],
    '效果叠加': ['overlay/effects'],
    '移动模式叠加': ['overlay/effects'],
    '硬编码': ['other'],
}

# 工作表类别 -> 同尺寸 pngs_* 根目录关键词（按优先级）。仓库当前按
# items/monsters/terrain 等拆分，而不是旧脚本假设的 Chibi/normal 两类。
CATEGORY_ROOT_HINTS = {
    '主贴图-物品': ('items', 'stuff'),
    '主贴图-怪物': ('monsters', 'stuff'),
    '尸体': ('monsters', 'stuff'),
    '主贴图-地形': ('terrain', 'stuff'),
    '主贴图-家具': ('furniture', 'stuff'),
    '主贴图-载具部件': ('vehicle_parts', 'stuff'),
    '主贴图-陷阱场效': ('fields', 'stuff'),
    'overmap': ('terrain', 'stuff'),
    '地图事件': ('terrain', 'stuff'),
    '穿戴-中性': ('items', 'character', 'stuff'),
    '穿戴-男': ('items', 'character', 'stuff'),
    '穿戴-女': ('items', 'character', 'stuff'),
    '手持': ('items', 'stuff'),
    '手持尸体': ('items', 'stuff'),
    '变异生化叠加': ('items', 'character', 'stuff'),
    '变异生化-中性': ('items', 'character', 'stuff'),
    '变异生化-男': ('items', 'character', 'stuff'),
    '变异生化-女': ('items', 'character', 'stuff'),
    '效果叠加': ('stuff', 'offset', 'fields'),
    '移动模式叠加': ('stuff', 'character'),
    '硬编码': ('stuff', 'terrain', 'items'),
}

# 列名（与 tileset_workbook.py 输出一致）
COL = {'类型': 0, 'ID': 1, '名称': 2, '前缀': 3, '完整贴图ID': 4,
       '当前贴图': 5, '完成状态': 6, '提交': 7, '贡献者': 8, '备注': 9}


def png_size(path):
    """读取 PNG 宽高（不依赖 PIL）。返回 (w,h) 或 None。"""
    try:
        with open(path, 'rb') as f:
            head = f.read(24)
        if head[:8] != b'\x89PNG\r\n\x1a\n' or head[12:16] != b'IHDR':
            return None
        w, h = struct.unpack('>II', head[16:24])
        return (w, h)
    except Exception:
        return None


def load_defined_sizes(repo):
    """从 tile_info.json 读出所有已定义的 (w,h) 尺寸集合。"""
    sizes = set()
    info_path = Path(repo) / 'tile_info.json'
    try:
        info = json.load(open(info_path, encoding='utf-8'))
    except Exception:
        return sizes
    base_w = info[0].get('width', 32) if info else 32
    base_h = info[0].get('height', 32) if info else 32
    for sheet in (info[1:] if isinstance(info, list) else []):
        for key, spec in sheet.items():
            if not key.endswith('.png') or not isinstance(spec, dict):
                continue
            w = spec.get('sprite_width', base_w)
            h = spec.get('sprite_height', base_h)
            sizes.add((w, h))
    sizes.add((base_w, base_h))
    return sizes


def find_size_dirs(repo, w, h):
    """返回该尺寸对应的所有 pngs_<root>_<w>x<h> 目录（Path 列表）。"""
    suffix = f'_{w}x{h}'
    out = []
    for d in sorted(Path(repo).glob(f'pngs_*{suffix}')):
        if d.is_dir():
            out.append(d)
    return out


def pick_size_dir(repo, w, h, category):
    """按工作表类别从同尺寸目录中选择语义相符的 pngs_* 根目录。"""
    dirs = find_size_dirs(repo, w, h)
    if not dirs:
        return None
    for hint in CATEGORY_ROOT_HINTS.get(category, ('stuff',)):
        marker = f'_{hint}_'
        matches = [directory for directory in dirs if marker in directory.name]
        if matches:
            return matches[0]
    return dirs[0]


def resolve_subdir(parent, category):
    """复用已有分类子目录；平铺式 pngs_* 目录则继续放在根目录。"""
    candidates = CATEGORY_SUBDIRS.get(category, ['other'])
    for c in candidates:
        if (parent / c).is_dir():
            return parent / c
    if any(parent.glob('*.png')) or any(parent.glob('*.json')):
        return parent
    return parent / candidates[0]


def parse_rels(text):
    """解析 .rels，返回 {rId: Target}。属性顺序不固定，逐个 Relationship 分别抓。"""
    out = {}
    for m in re.finditer(r'<Relationship\b[^>]*?/?>', text):
        tag = m.group(0)
        rid = re.search(r'Id="([^"]+)"', tag)
        tgt = re.search(r'Target="([^"]+)"', tag)
        if rid and tgt:
            out[rid.group(1)] = tgt.group(1)
    return out


# Excel 列字母 <-> 0基序号
def col_to_idx(col_letters):
    n = 0
    for ch in col_letters:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1


def _unescape_xml(s):
    return (s.replace('&lt;', '<').replace('&gt;', '>')
             .replace('&quot;', '"').replace('&apos;', "'")
             .replace('&amp;', '&'))


def load_shared_strings(tmp):
    """读 sharedStrings.xml，返回字符串列表（按索引）。"""
    p = Path(tmp, 'xl', 'sharedStrings.xml')
    if not p.exists():
        return []
    text = p.read_text(encoding='utf-8')
    out = []
    # 每个 <si>...</si> 是一条；其文本是内部所有 <t> 拼接
    for si in re.findall(r'<si>(.*?)</si>', text, re.S):
        parts = re.findall(r'<t[^>]*>(.*?)</t>', si, re.S)
        out.append(_unescape_xml(''.join(parts)))
    return out


def read_sheet_cells(sheet_xml_path, shared):
    """直接解析 sheetN.xml，返回 {(row1基, col0基): value(str)}。

    不依赖 openpyxl —— 规避其在某些 WPS 导出的 styles.xml（空 <fill/>）
    + Python 3.14 下解析样式表崩溃的兼容问题。
    """
    text = Path(sheet_xml_path).read_text(encoding='utf-8')
    cells = {}
    # 同时匹配自闭合 <c .../> 和成对 <c ...>...</c>
    # （空单元格常写成自闭合，若只配对 <c>...</c> 会跨格错位吞内容）
    for cm in re.finditer(r'<c\b([^>]*?)(?:/>|>(.*?)</c>)', text, re.S):
        attrs = cm.group(1)
        body = cm.group(2) or ''
        ref = re.search(r'r="([A-Z]+)(\d+)"', attrs)
        if not ref:
            continue
        col0 = col_to_idx(ref.group(1))
        row1 = int(ref.group(2))
        ctype = re.search(r't="([^"]+)"', attrs)
        ctype = ctype.group(1) if ctype else 'n'
        vm = re.search(r'<v>(.*?)</v>', body, re.S)
        if ctype == 's':            # 共享字符串：v 是索引
            if vm:
                idx = int(vm.group(1))
                val = shared[idx] if 0 <= idx < len(shared) else ''
            else:
                val = ''
        elif ctype in ('inlineStr', 'str'):  # 内联/公式字符串
            tm = re.search(r'<t[^>]*>(.*?)</t>', body, re.S)
            if tm:
                val = _unescape_xml(tm.group(1))
            else:
                val = _unescape_xml(vm.group(1)) if vm else ''
        else:                       # 数字/其他：取 v 原文
            val = _unescape_xml(vm.group(1)) if vm else ''
        cells[(row1, col0)] = val
    return cells


def extract_data(xlsx_path):
    """解压 xlsx，手工解析（不经 openpyxl 读取）。

    返回 (sheets, tmp)，sheets[sheet_name] = {
        'cells': {(row,col): value},
        'images': {row(1基): image_abs_path},
    }
    """
    tmp = tempfile.mkdtemp(prefix='ingest_')
    with zipfile.ZipFile(xlsx_path) as z:
        z.extractall(tmp)

    shared = load_shared_strings(tmp)

    # workbook.xml: 按出现顺序拿 sheet 名 + r:id
    wb_xml = Path(tmp, 'xl', 'workbook.xml').read_text(encoding='utf-8')
    rid_to_name = {}
    for m in re.finditer(r'<sheet\b[^>]*?/?>', wb_xml):
        tag = m.group(0)
        nm = re.search(r'name="([^"]+)"', tag)
        rid = re.search(r'r:id="(rId\d+)"', tag)
        if nm and rid:
            rid_to_name[rid.group(1)] = _unescape_xml(nm.group(1))
    rid_to_target = parse_rels(
        Path(tmp, 'xl', '_rels', 'workbook.xml.rels').read_text(encoding='utf-8'))
    file_to_name = {}
    for rid, name in rid_to_name.items():
        base = os.path.basename(rid_to_target.get(rid, ''))
        if base:
            file_to_name[base] = name

    sheets = {name: {'cells': {}, 'images': {}} for name in rid_to_name.values()}

    ws_dir = Path(tmp, 'xl', 'worksheets')
    for sheet_file in ws_dir.glob('sheet*.xml'):
        sname = file_to_name.get(sheet_file.name)
        if not sname:
            continue
        sheets[sname]['cells'] = read_sheet_cells(sheet_file, shared)

        # 该 sheet 的 drawing -> 图片锚点
        rels = ws_dir / '_rels' / (sheet_file.name + '.rels')
        if not rels.exists():
            continue
        dm = re.search(r'drawing(\d+)\.xml', rels.read_text(encoding='utf-8'))
        if not dm:
            continue
        draw_file = Path(tmp, 'xl', 'drawings', f'drawing{dm.group(1)}.xml')
        draw_rels = Path(tmp, 'xl', 'drawings', '_rels',
                         f'drawing{dm.group(1)}.xml.rels')
        if not draw_file.exists() or not draw_rels.exists():
            continue
        embed_to_media = {rid: os.path.basename(tgt) for rid, tgt
                          in parse_rels(draw_rels.read_text(encoding='utf-8')).items()}
        dtext = draw_file.read_text(encoding='utf-8')
        # 锚点命名空间前缀可选：WPS 用 <xdr:from><xdr:row>，openpyxl 用 <from><row>
        for am in re.finditer(
                r'<(?:\w+:)?from>.*?<(?:\w+:)?row>(\d+)</(?:\w+:)?row>.*?'
                r'embed="(rId\d+)"', dtext, re.S):
            row0 = int(am.group(1))
            media = embed_to_media.get(am.group(2))
            if not media:
                continue
            media_path = Path(tmp, 'xl', 'drawings', 'media', media)
            if not media_path.exists():
                media_path = Path(tmp, 'xl', 'media', media)
            if media_path.exists():
                sheets[sname]['images'][row0 + 1] = str(media_path)

    return sheets, tmp


def validate_row(cells, row, defined_sizes, img_path):
    """返回 (ok, reason, info)。cells 为 {(row,col): value}。"""
    def cell(name):
        v = cells.get((row, COL[name]))
        return v.strip() if isinstance(v, str) else v

    full_id = cell('完整贴图ID') or ''
    status = cell('完成状态') or ''
    contributor = cell('贡献者')
    reasons = []

    if status != '已完成':
        reasons.append(f'完成状态为「{status or "空"}」，需为「已完成」')
    if not contributor:
        reasons.append('缺贡献者')

    size = png_size(img_path)
    if size is None:
        reasons.append('图片无法读取或非 PNG')
    elif size not in defined_sizes:
        reasons.append(f'尺寸 {size[0]}x{size[1]} 不在已定义尺寸内')

    info = {'full_id': full_id, 'status': status,
            'contributor': contributor or '', 'size': size}
    return (len(reasons) == 0, '；'.join(reasons), info)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('xlsx', help='贴图美术师填好的 Excel')
    ap.add_argument('--repo', default='.', help='贴图源仓库根目录（含 tile_info.json）')
    ap.add_argument('--report', default='归置检测表.xlsx', help='输出检测表路径')
    ap.add_argument('--dry-run', action='store_true', help='只检测不落地')
    ap.add_argument('--force', action='store_true', help='允许覆盖归置目标位置已存在的同名 png')
    ap.add_argument('--replace', action='store_true',
                    help='替换模式：全仓库搜同名 png，找到则原位替换官方贴图'
                         '（不生成 json，被覆盖文件备份进检测表）')
    args = ap.parse_args()

    repo = Path(args.repo).resolve()
    if not (repo / 'tile_info.json').exists():
        sys.exit(f'错误：{repo} 下无 tile_info.json，--repo 需指向贴图源仓库根目录')

    defined_sizes = load_defined_sizes(repo)
    if not defined_sizes:
        sys.exit('错误：未能从 tile_info.json 读到任何已定义尺寸')

    # 全仓库 png 名 -> 路径索引（用于 --replace 找官方同名贴图）
    existing_index = build_png_index(repo) if args.replace else {}

    sheets, tmp = extract_data(args.xlsx)

    # 收集所有带图行 + 校验
    records = []  # dict per image row
    for sname, sdata in sheets.items():
        rowmap = sdata['images']
        if not rowmap:
            continue
        cells = sdata['cells']
        for row, img_path in sorted(rowmap.items()):
            ok, reason, info = validate_row(cells, row, defined_sizes, img_path)
            png_name = info['full_id'] + '.png' if info['full_id'] else ''
            rec = {
                'sheet': sname, 'row': row, 'img': img_path,
                'full_id': info['full_id'], 'png_name': png_name,
                'status': info['status'], 'contributor': info['contributor'],
                'size': info['size'], 'ok': ok, 'reason': reason,
                'png_dest': None, 'json_dest': None, 'placed': False,
                'mode': '', 'backup_img': None,
            }
            # --replace：全仓库找同名官方贴图，找到则原位替换、不生成 json
            existing = existing_index.get(png_name) if args.replace else None
            if existing:
                rec['png_dest'] = str(Path(existing).relative_to(repo))
                rec['json_dest'] = ''      # 替换官方图不生成 json
                rec['mode'] = '替换'
            else:
                # 全新图：按尺寸+类别算归置位置，生成 json
                rec['png_dest'], rec['json_dest'] = plan_dest(rec, repo)
                rec['mode'] = '新增'
            records.append(rec)

    if not records:
        print('未在表格中找到任何内嵌图片，无事可做。')
        shutil.rmtree(tmp, ignore_errors=True)
        write_report(args.report, records)
        return

    all_ok = all(r['ok'] for r in records)
    n_ok = sum(1 for r in records if r['ok'])
    print(f'带图行 {len(records)} 个，合格 {n_ok}，不合格 {len(records) - n_ok}')

    do_place = all_ok and not args.dry_run
    if not all_ok:
        print('存在不合格行，整体不归置（仅输出检测表）。')
    elif args.dry_run:
        print('--dry-run：仅检测，不落地。')

    if do_place:
        for r in records:
            place_one(r, repo, args.force)

    write_report(args.report, records)
    shutil.rmtree(tmp, ignore_errors=True)
    # 清理备份临时目录（已内嵌进检测表）
    for r in records:
        bak = r.get('backup_img')
        if bak:
            shutil.rmtree(os.path.dirname(bak), ignore_errors=True)

    print(f'\n检测表已写入：{args.report}')
    if do_place:
        n_new = sum(1 for r in records if r.get('placed') and r.get('mode') == '新增')
        n_rep = sum(1 for r in records if r.get('placed') and r.get('mode') == '替换')
        print(f'已归置 {n_new + n_rep} 张（新增 {n_new}，替换 {n_rep}）。')


def build_png_index(repo):
    """扫描全仓库 pngs_* 下所有 png，返回 {文件名: 绝对路径}。

    同名出现多次时保留第一个（按目录排序稳定）。用于 --replace 找官方同名贴图。
    """
    index = {}
    for d in sorted(repo.glob('pngs_*')):
        if not d.is_dir():
            continue
        for p in d.rglob('*.png'):
            index.setdefault(p.name, str(p))
    return index


def plan_dest(rec, repo):
    """根据尺寸+类别计算 png/json 的目标路径（相对 repo）。

    返回 (png_rel, json_rel)；无法确定（缺尺寸/无对应目录/缺ID）时返回 (None, None)。
    仅计算位置，不写任何文件。
    """
    size = rec.get('size')
    if not size or not rec.get('full_id'):
        return (None, None)
    parent = pick_size_dir(repo, size[0], size[1], rec['sheet'])
    if parent is None:
        return (None, None)
    dest_dir = resolve_subdir(parent, rec['sheet'])
    png_rel = (dest_dir / rec['png_name']).relative_to(repo)
    json_rel = (dest_dir / (rec['full_id'] + '.json')).relative_to(repo)
    return (str(png_rel), str(json_rel))


def place_one(rec, repo, force):
    """归置单张 png（+ 必要时 json）。结果写回 rec['placed'] / rec['reason']。

    - 替换模式（rec['mode']=='替换'）：原位覆盖官方 png，先把原文件备份到
      rec['backup_img']（供检测表内嵌），不生成 json。
    - 新增模式：复制到目标目录，并生成 CDDA 散图 json。
    """
    png_rel = rec.get('png_dest')
    if not png_rel:
        size = rec.get('size')
        rec['ok'] = False
        rec['reason'] = f'无 {size[0]}x{size[1]} 对应的 pngs 目录' if size else '无法确定归置位置'
        return
    dest_png = repo / png_rel
    is_replace = (rec.get('mode') == '替换')

    if is_replace:
        # 覆盖官方贴图：先备份原文件（留副本给检测表内嵌），再覆盖
        if dest_png.exists():
            bak = Path(tempfile.mkdtemp(prefix='bak_')) / dest_png.name
            shutil.copy2(dest_png, bak)
            rec['backup_img'] = str(bak)
        shutil.copy2(rec['img'], dest_png)
        rec['placed'] = True
        return

    # 新增模式
    dest_png.parent.mkdir(parents=True, exist_ok=True)
    if dest_png.exists() and not force:
        rec['ok'] = False
        rec['reason'] = f'目标已存在同名 png：{png_rel}（用 --force 覆盖或 --replace 替换）'
        return
    shutil.copy2(rec['img'], dest_png)
    # 全新图生成 CDDA 散图 json
    if rec.get('json_dest'):
        dest_json = repo / rec['json_dest']
        tile_def = [{'id': [rec['full_id']], 'fg': [rec['full_id']]}]
        dest_json.write_text(
            json.dumps(tile_def, ensure_ascii=False, indent=2) + '\n',
            encoding='utf-8')
    rec['placed'] = True


def write_report(path, records):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.drawing.image import Image as XLImage

    wb = Workbook()
    ws = wb.active
    ws.title = '归置检测'
    cols = ['工作表', '行', 'ID/完整贴图ID', 'png文件名', 'json文件名', '尺寸',
            '完成状态', '贡献者', '模式', '结果', '原因',
            '归置png路径', '归置json路径', '是否已落地', '备份(被覆盖原图)']
    ws.append(cols)
    hf = Font(bold=True, color='FFFFFF')
    hfill = PatternFill('solid', fgColor='4472C4')
    okfill = PatternFill('solid', fgColor='C6EFCE')
    badfill = PatternFill('solid', fgColor='FFC7CE')
    center = Alignment(horizontal='center')
    for c in ws[1]:
        c.font, c.fill, c.alignment = hf, hfill, center
    ws.freeze_panes = 'A2'

    result_col = cols.index('结果') + 1
    backup_col = cols.index('备份(被覆盖原图)') + 1
    backup_letter = chr(ord('A') + backup_col - 1)
    for r in records:
        size_str = f'{r["size"][0]}x{r["size"][1]}' if r.get('size') else '—'
        # 替换模式不生成 json，json 名/路径列留空避免误导
        json_name = ('' if r.get('mode') == '替换'
                     else (r['full_id'] + '.json') if r.get('full_id') else '')
        ws.append([r['sheet'], r['row'], r['full_id'], r['png_name'], json_name,
                   size_str, r['status'], r['contributor'], r.get('mode') or '',
                   '通过' if r['ok'] else '不通过', r['reason'],
                   r.get('png_dest') or '', r.get('json_dest') or '',
                   '是' if r.get('placed') else '否', ''])
        rownum = ws.max_row
        ws.cell(rownum, result_col).fill = okfill if r['ok'] else badfill
        # 被覆盖的原图内嵌到「备份」列
        bak = r.get('backup_img')
        if bak and os.path.exists(bak):
            try:
                img = XLImage(bak)
                ws.add_image(img, f'{backup_letter}{rownum}')
                ws.row_dimensions[rownum].height = 50
            except Exception:
                ws.cell(rownum, backup_col).value = f'(备份: {os.path.basename(bak)})'

    widths = {'A': 16, 'B': 6, 'C': 26, 'D': 26, 'E': 26, 'F': 8, 'G': 10,
              'H': 12, 'I': 8, 'J': 8, 'K': 34, 'L': 38, 'M': 38, 'N': 10, 'O': 18}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    wb.save(path)


if __name__ == '__main__':
    main()
