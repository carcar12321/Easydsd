#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""easydsd v0.1 - DART 감사보고서 변환 도구"""

import os, re, sys, io, zipfile, threading, webbrowser, socket, time

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

try:
    from flask import Flask, request, send_file, jsonify, render_template_string
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'flask', 'openpyxl', '-q'])
    from flask import Flask, request, send_file, jsonify, render_template_string
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter

def find_free_port(start=5000, end=5099):
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    return start

PORT = find_free_port()
EDIT_COLOR = 'FFF2CC'
C = {'navy':'1F4E79','blue':'2E75B6','lblue':'DEEAF1',
     'yellow':'FFF2CC','white':'FFFFFF','lgray':'F2F2F2','orange':'C55A11'}

FIN_TABLE_MAP = [
    (['재 무 상 태 표'],         '🏦재무상태표'),
    (['포 괄 손 익 계 산 서'],    '💹포괄손익계산서'),
    (['자 본 변 동 표'],         '📈자본변동표'),
    (['현 금 흐 름 표'],         '💰현금흐름표'),
]
EXT_DESC = {
    'TOT_ASSETS':'총자산(백만원)','TOT_DEBTS':'총부채(백만원)',
    'TOT_SALES':'매출액(백만원)', 'TOT_EMPL':'총직원수',
    'GMSH_DATE':'주총일자(YYYYMMDD)','SUPV_OPIN':'감사의견코드',
    'AUDIT_CIK':'감사인CIK','CRP_RGS_NO':'법인등록번호',
}

def fill(c): return PatternFill('solid', fgColor=c)
def fnt(color='000000',bold=False,size=9,italic=False):
    return Font(color=color,bold=bold,size=size,italic=italic)
def aln(h='left',v='center',wrap=False):
    return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

# ── &amp;cr; 클린업 ────────────────────────────────────────────────────────────
def clean_cr(s, as_newline=False):
    """
    &amp;cr; → 줄바꿈 또는 공백으로 변환
    as_newline=True  : Excel 셀 내 줄바꿈 (\n)
    as_newline=False : 공백 (시트명, 제목 등에 사용)
    """
    repl = '\n' if as_newline else ' '
    s = s.replace('&amp;cr;', repl).replace('&cr;', repl)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def clean_title(s):
    """시트명/섹션 제목 정리 - &amp;cr; 제거, 공백 정리"""
    s = clean_cr(s, as_newline=False)
    # &amp; 외 나머지 엔티티도 처리
    s = s.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def is_blank_title(s):
    """의미 없는 제목(&cr; 만으로 구성) 여부"""
    cleaned = re.sub(r'[&;a-z]+', '', s).strip()
    return len(cleaned) == 0

# ── XML 파싱 ───────────────────────────────────────────────────────────────────
def parse_cell(m):
    attrs = m.group(1)
    val   = re.sub(r'<[^>]+>', '', m.group(0))
    # 엔티티 변환 순서 중요: &amp; 먼저
    val   = val.replace('&amp;cr;', '\n')   # DART 줄바꿈 → Excel 줄바꿈
    val   = val.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"')
    val   = val.replace('&cr;', '\n')        # 혹시 남아있는 &cr; 처리
    val   = val.strip()
    cs    = int(x.group(1)) if (x := re.search(r'COLSPAN="(\d+)"', attrs)) else 1
    tag   = re.match(r'<([A-Z]+)', m.group(0)).group(1)
    return dict(value=val, colspan=cs, tag=tag)

def is_num_or_decimal(val):
    v = val.strip().replace(',','').replace('(','').replace(')','').replace('%','').replace('-','').replace(' ','').split('\n')[0]
    if not v: return False
    try: float(v); return True
    except: return False

def parse_xml(xml):
    exts = re.findall(r'<EXTRACTION[^>]*ACODE="([^"]+)"[^>]*>([^<]+)</EXTRACTION>', xml)
    tables = []
    for ti, tm in enumerate(re.finditer(r'<TABLE([^>]*)>(.*?)</TABLE>', xml, re.DOTALL)):
        ctx = xml[max(0, tm.start()-600): tm.start()]
        # 앞 컨텍스트 OR 테이블 자체 내용에서 재무제표명 감지
        table_body = tm.group(0)
        fin_label = next((lbl for kws, lbl in FIN_TABLE_MAP
                          if any(kw in ctx or kw in table_body for kw in kws)), '')
        # ctx_title: &amp;cr; 클린업 + 의미없는 제목 제외
        raw_titles = re.findall(r'<(?:TITLE|P)[^>]*>([^<]{3,80})</(?:TITLE|P)>', ctx)
        ctx_titles = [clean_title(t) for t in raw_titles if not is_blank_title(t) and len(clean_title(t)) > 1]
        rows = []
        for tr in re.finditer(r'<TR[^>]*>(.*?)</TR>', tm.group(2), re.DOTALL):
            cells = [parse_cell(cm)
                     for cm in re.finditer(r'<(?:TD|TH|TU|TE)([^>]*)>.*?</(?:TD|TH|TU|TE)>', tr.group(1), re.DOTALL)]
            if cells: rows.append(cells)
        tables.append(dict(idx=ti, fin_label=fin_label,
                           ctx_title=(ctx_titles[-1] if ctx_titles else ''),
                           rows=rows, start=tm.start()))
    return exts, tables

# ── DSD → Excel ────────────────────────────────────────────────────────────────
def make_sheet_name(fin_label, ctx_title, note_n):
    """깔끔한 시트명 생성"""
    if fin_label:
        return fin_label[:31], fin_label
    # 시트명에 사용할 짧은 제목: 특수문자 제거
    short = re.sub(r'[\\/*?\[\]:]', '', ctx_title)[:10].strip()
    if not short:
        short = '주석'
    sname = f'📝{note_n:02d}_{short}'[:31]
    return sname, ctx_title or f'주석 {note_n}'

def dsd_to_excel_bytes(dsd_bytes):
    with zipfile.ZipFile(io.BytesIO(dsd_bytes)) as zf:
        files = {n: zf.read(n) for n in zf.namelist()}
    xml      = files.get('contents.xml', b'').decode('utf-8', errors='replace')
    meta_xml = files.get('meta.xml', b'').decode('utf-8', errors='replace')
    exts, tables = parse_xml(xml)
    wb = openpyxl.Workbook()

    # ① 사용안내
    ws0 = wb.active; ws0.title = '📋사용안내'; ws0.sheet_view.showGridLines = False
    guide = [
        ('DART 감사보고서 DSD - Excel 변환 도구 (easydsd v0.1)', True, C['white'], C['navy'], 13),
        ('', False, '', '', 8),
        ('【 작업 순서 】', True, C['navy'], C['lblue'], 11),
        ('  1. 이 Excel 파일의 노란색 셀을 당해년도 숫자/텍스트로 수정하세요', False, '000000', C['white'], 10),
        ('  2. 저장 후 도구로 돌아와 "Excel → DSD" 탭에서 변환하세요', False, '000000', C['white'], 10),
        ('', False, '', '', 8),
        ('【 색상 범례 】', True, C['navy'], C['lblue'], 11),
        ('  노란색 = 수정 가능 (금액, 주주명, 지분율, 텍스트 모두)', False, '000000', C['yellow'], 10),
        ('  파란색 = 헤더 (수정 불필요)', False, C['white'], C['navy'], 10),
        ('', False, '', '', 8),
        ('【 주의사항 】', True, C['navy'], C['lblue'], 11),
        ('  _원본XML 시트는 절대 수정/삭제하지 마세요 (변환에 필수)', False, C['orange'], C['white'], 10),
        ('  숫자 입력 시 콤마 포함/미포함 모두 가능', False, '000000', C['white'], 10),
        ('  음수는 -1234567 또는 (1,234,567) 형식 모두 가능', False, '000000', C['white'], 10),
    ]
    for ri, (txt, bold, fg, bg, sz) in enumerate(guide, 1):
        c = ws0.cell(ri, 1, txt); c.font = fnt(fg or '000000', bold=bold, size=sz)
        if bg: c.fill = fill(bg)
        c.alignment = aln('left', wrap=True); ws0.row_dimensions[ri].height = 21
    ws0.column_dimensions['A'].width = 65

    # ② 요약수치
    ws_e = wb.create_sheet('📊요약수치'); ws_e.sheet_view.showGridLines = False
    for ci, (h, w) in enumerate([('ACODE', 15), ('값 (수정가능)', 22), ('설명', 28)], 1):
        c = ws_e.cell(1, ci, h); c.fill = fill(C['navy']); c.font = fnt(C['white'], bold=True); c.alignment = aln('center')
        ws_e.column_dimensions[get_column_letter(ci)].width = w
    for ri, (code, val) in enumerate(exts, 2):
        ws_e.cell(ri, 1, code).font = fnt(bold=True, size=9)
        vc = ws_e.cell(ri, 2, val); vc.fill = fill(C['yellow']); vc.alignment = aln('right')
        ws_e.cell(ri, 3, EXT_DESC.get(code, '')).font = fnt(size=9, italic=True)
        ws_e.row_dimensions[ri].height = 18

    # ③ 그룹핑 전략
    # ─────────────────────────────────────────────────────────────────────────
    # 그룹 0: 재무상태표 앞까지 모든 TABLE → "📝00_서문" 1개
    # 그룹 1: 🏦재무상태표 (제목+데이터) + 바로 뒤 TABLE 1개
    # 그룹 2: 💹포괄손익계산서         + 바로 뒤 TABLE 1개
    # 그룹 3: 📈자본변동표             + 바로 뒤 TABLE 1개
    # 그룹 4: 💰현금흐름표             + 바로 뒤 TABLE 1개
    # 그룹 5+: 나머지 주석 TABLE들을 10개씩 묶어 시트 1개
    # ─────────────────────────────────────────────────────────────────────────

    FIN_ORDER = ['🏦재무상태표', '💹포괄손익계산서', '📈자본변동표', '💰현금흐름표']

    # ── 1단계: TABLE 전체를 논리 그룹으로 분류 ──────────────────────────────
    groups = []   # [(sheet_name, [table_obj, ...]), ...]

    i = 0
    # 그룹 0: 첫 fin 테이블 전까지
    pre_fin = []
    while i < len(tables) and not tables[i]['fin_label']:
        pre_fin.append(tables[i]); i += 1
    if pre_fin:
        groups.append(('📝00_서문', pre_fin, False))

    # 그룹 1~4: 각 재무제표 + 뒤 1개
    for fin_label in FIN_ORDER:
        if i >= len(tables): break
        # fin_label에 해당하는 TABLE들 수집
        fin_tbls = []
        while i < len(tables) and tables[i]['fin_label'] == fin_label:
            fin_tbls.append(tables[i]); i += 1
        if not fin_tbls: continue
        # 바로 뒤 TABLE 1개 추가 (다음 fin 그룹이 시작되기 전)
        if i < len(tables) and not tables[i]['fin_label']:
            fin_tbls.append(tables[i]); i += 1
        groups.append((fin_label[:31], fin_tbls, False))

    # 그룹 5+: 나머지 10개씩
    remaining = tables[i:]
    chunk_n = 1
    for start in range(0, len(remaining), 10):
        chunk = remaining[start:start+10]
        groups.append((f'📝{chunk_n:02d}_주석', chunk, True))  # show_titles
        chunk_n += 1

    # ── 2단계: 그룹 → 시트 생성 헬퍼 ────────────────────────────────────────
    def write_tables_to_sheet(ws, tbl_list, show_titles=False):
        """여러 TABLE을 하나의 시트에 순서대로 기록
        show_titles=True: 각 TABLE 앞에 ctx_title 구분선 삽입 (주석 시트용)
        """
        er = 1
        max_cols_all = 1
        table_start_rows = {}  # {tbl['idx']: er_at_start}
        for tbl in tbl_list:
            if not tbl['rows']: continue
            max_cols_all = max(max_cols_all,
                min(max((sum(c['colspan'] for c in row) for row in tbl['rows']), default=1), 26))
        for tbl in tbl_list:
            # 주석 구분선: ctx_title이 있으면 회색 배경 구분 행 삽입
            if show_titles and tbl.get('ctx_title'):
                div_cell = ws.cell(er, 1, tbl['ctx_title'])
                div_cell.fill = PatternFill('solid', fgColor='D9D9D9')
                div_cell.font = Font(bold=True, size=9, color='333333')
                div_cell.alignment = Alignment(horizontal='left', vertical='center')
                if max_cols_all > 1:
                    try: ws.merge_cells(start_row=er, start_column=1,
                                        end_row=er, end_column=max_cols_all)
                    except: pass
                ws.row_dimensions[er].height = 16
                er += 1
            table_start_rows[tbl['idx']] = er  # 이 TABLE의 Excel 시작 행
            for row in tbl['rows']:
                col = 1
                for cell in row:
                    if col > 26: break
                    wc = ws.cell(er, col, cell['value']); v, tag = cell['value'], cell['tag']
                    if tag in ('TH', 'TE'):
                        wc.fill = fill(C['navy']); wc.font = fnt(C['white'], bold=True, size=9)
                        wc.alignment = aln('center', wrap=True)
                    else:
                        wc.fill = fill(C['yellow']); wc.font = fnt(size=9)
                        wc.alignment = aln('right' if is_num_or_decimal(v) else 'left', wrap=True)
                    if '\n' in str(v):
                        ws.row_dimensions[er].height = max(18, 18*(str(v).count('\n')+1))
                    span = min(cell['colspan'], 26-col+1)
                    if span > 1:
                        try: ws.merge_cells(start_row=er, start_column=col,
                                            end_row=er, end_column=col+span-1)
                        except: pass
                    col += cell['colspan']
                if not ws.row_dimensions[er].height or ws.row_dimensions[er].height < 18:
                    ws.row_dimensions[er].height = 18
                er += 1
        ws.column_dimensions['A'].width = 28
        for ci in range(2, max_cols_all+1):
            ws.column_dimensions[get_column_letter(ci)].width = 18
        return table_start_rows  # {t_idx: excel_row_start}

    # ── 3단계: 그룹 → 시트 생성 ──────────────────────────────────────────────
    sheet_map = []
    for group_item in groups:
        sname, tbl_list = group_item[0], group_item[1]
        show_titles = group_item[2] if len(group_item) > 2 else False
        ws = wb.create_sheet(sname); ws.sheet_view.showGridLines = False
        t_start_rows = write_tables_to_sheet(ws, tbl_list, show_titles)
        for tbl in tbl_list:
            excel_row = t_start_rows.get(tbl['idx'], -1) if t_start_rows else -1
            sheet_map.append((sname, tbl['idx'], excel_row))

    # ④ _원본XML
    ws_raw = wb.create_sheet('_원본XML'); ws_raw.sheet_view.showGridLines = False
    ws_raw.cell(1, 1, '이 시트는 DSD 복원에 필수입니다. 절대 수정/삭제 금지!').font = fnt(C['orange'], bold=True, size=9)
    ws_raw.cell(2, 1, 'meta_xml'); ws_raw.cell(2, 2, meta_xml or '')
    ws_raw.cell(4, 1, 'sheet_name'); ws_raw.cell(4, 2, 'table_idx'); ws_raw.cell(4, 3, 'fin_label'); ws_raw.cell(4, 4, 'ctx_title'); ws_raw.cell(4, 5, 'excel_start_row')
    for ri, sm_item in enumerate(sheet_map, 5):
        sname, t_idx = sm_item[0], sm_item[1]
        excel_row = sm_item[2] if len(sm_item) > 2 else -1
        t = tables[t_idx]
        ws_raw.cell(ri, 1, sname); ws_raw.cell(ri, 2, t_idx)
        ws_raw.cell(ri, 3, t['fin_label']); ws_raw.cell(ri, 4, t['ctx_title'])
        ws_raw.cell(ri, 5, excel_row)
    for ci, w in [(1,35),(2,12),(3,22),(4,40),(5,14)]: ws_raw.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── Excel → DSD ────────────────────────────────────────────────────────────────
def is_note_ref(val):
    """주석 참조번호 여부: '5,32,33' 처럼 1~2자리 숫자가 콤마로 연결된 형태"""
    parts = val.strip().split(',')
    return (len(parts) >= 2 and
            all(p.strip().isdigit() and 1 <= len(p.strip()) <= 2 for p in parts))

def normalize_num(val):
    v = str(val).strip()
    if not v or v in ('-', ''): return v
    # 줄바꿈 포함 값: 줄바꿈을 &amp;cr; 로 복원
    if '\n' in v:
        lines = v.split('\n')
        lines = [normalize_num(l) for l in lines]
        return '&amp;cr;'.join(lines)
    # 주석 참조번호 (5,32,33 / 6,18,32,33 등): 원본 그대로 유지
    if is_note_ref(v):
        return v
    negative = v.startswith('-') or (v.startswith('(') and v.endswith(')'))
    cleaned  = v.replace(',', '').replace('(', '').replace(')', '').replace('-', '').replace(' ', '')
    if cleaned.isdigit() and len(cleaned) >= 3:
        fmt = f"{int(cleaned):,}"; v = f"({fmt})" if negative else fmt
    v = v.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
    return v

def is_edit(cell):
    f = cell.fill
    if f and f.fill_type == 'solid':
        fg = f.fgColor
        if fg and fg.type == 'rgb': return fg.rgb.upper().endswith(EDIT_COLOR.upper())
    return False

def excel_to_dsd_bytes(orig_dsd_bytes, xlsx_bytes):
    with zipfile.ZipFile(io.BytesIO(orig_dsd_bytes)) as zf:
        orig_files = {n: zf.read(n) for n in zf.namelist()}
    contents_xml = orig_files['contents.xml'].decode('utf-8', errors='replace')
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)

    # ── 매핑: {sheet_name: [(t_idx, excel_start_row), ...]} ──────────────────
    # excel_start_row: 0-based (헤더 제외, 즉 min_row=2 기준)
    mapping = {}
    if '_원본XML' in wb.sheetnames:
        ws_raw = wb['_원본XML']
        for row in ws_raw.iter_rows(min_row=5, values_only=True):
            if not row or row[0] is None or row[1] is None: continue
            sname     = str(row[0]).strip()
            t_idx     = int(row[1])
            excel_row = int(row[4]) if len(row) > 4 and row[4] is not None else -1
            if sname:
                mapping.setdefault(sname, []).append((t_idx, excel_row))

    # ── 수정값 수집 ────────────────────────────────────────────────────────────
    exts      = {}
    t_changes = {}
    for sname in wb.sheetnames:
        if sname in ('📋사용안내', '_원본XML', '_meta'): continue
        ws = wb[sname]
        if sname == '📊요약수치':
            for row in ws.iter_rows(min_row=2):
                if len(row) < 2: continue
                cc, vc = row[0], row[1]
                if cc.value and vc.value is not None and is_edit(vc):
                    exts[str(cc.value).strip()] = str(vc.value).strip()
        else:
            changes = []
            for ri, row in enumerate(ws.iter_rows(min_row=2)):
                for ci, cell in enumerate(row):
                    if is_edit(cell) and cell.value is not None:
                        changes.append((ri, ci, str(cell.value)))
            if changes:
                t_changes[sname] = changes

    # ── EXTRACTION 패치 ────────────────────────────────────────────────────────
    for ext_code, val in exts.items():
        contents_xml = re.sub(
            rf'(<EXTRACTION[^>]*ACODE="{re.escape(ext_code)}"[^>]*>)[^<]+(</EXTRACTION>)',
            rf'\g<1>{val}\g<2>', contents_xml)

    # ── TABLE 패치 ─────────────────────────────────────────────────────────────
    table_positions = [(m.start(), m.end())
                       for m in re.finditer(r'<TABLE[^>]*>.*?</TABLE>',
                                            contents_xml, re.DOTALL)]
    patches = []

    for sname, changes in t_changes.items():
        t_info_list = mapping.get(sname)  # [(t_idx, excel_start_row), ...]
        if not t_info_list: continue

        all_changes = {(r, c): v for r, c, v in changes}

        for k, (t_idx, excel_start_row) in enumerate(t_info_list):
            if t_idx >= len(table_positions): continue

            if excel_start_row >= 0:
                # excel_start_row 기반 정확한 분배
                # 다음 TABLE 시작 행까지가 이 TABLE 범위
                next_start = (t_info_list[k+1][1]
                              if k+1 < len(t_info_list) and t_info_list[k+1][1] >= 0
                              else 99999)
                local_map = {
                    (r - excel_start_row, c): v
                    for (r, c), v in all_changes.items()
                    if excel_start_row <= r < next_start
                }
            else:
                # 구버전 호환: TR 카운트 기반 누적 오프셋
                row_offset = 0
                for j in range(k):
                    ti2 = t_info_list[j][0]
                    if ti2 < len(table_positions):
                        snip = contents_xml[table_positions[ti2][0]:table_positions[ti2][1]]
                        row_offset += len(re.findall(r'<TR[^>]*>', snip))
                snip = contents_xml[table_positions[t_idx][0]:table_positions[t_idx][1]]
                tr_cnt = len(re.findall(r'<TR[^>]*>', snip))
                local_map = {
                    (r - row_offset, c): v
                    for (r, c), v in all_changes.items()
                    if row_offset <= r < row_offset + tr_cnt
                }

            if not local_map: continue

            t_start, t_end = table_positions[t_idx]
            t_text = contents_xml[t_start:t_end]
            rebuilt = []; last = 0; td_row = 0

            for tr_m in re.finditer(r'(<TR[^>]*>)(.*?)(</TR>)', t_text, re.DOTALL):
                rebuilt.append(t_text[last:tr_m.start()])
                tr_body = tr_m.group(2); new_body = []; td_last = 0; td_col = 0
                for td_m in re.finditer(
                        r'(<(?:TD|TH|TU|TE)[^>]*>)(.*?)(</(?:TD|TH|TU|TE)>)',
                        tr_body, re.DOTALL):
                    new_body.append(tr_body[td_last:td_m.start()])
                    key = (td_row, td_col)
                    if key in local_map:
                        new_body.append(
                            td_m.group(1) + normalize_num(local_map[key]) + td_m.group(3))
                    else:
                        new_body.append(td_m.group(0))
                    td_last = td_m.end(); td_col += 1
                new_body.append(tr_body[td_last:])
                rebuilt.append(tr_m.group(1) + ''.join(new_body) + tr_m.group(3))
                last = tr_m.end(); td_row += 1
            rebuilt.append(t_text[last:])
            patches.append((t_start, t_end, ''.join(rebuilt)))

    # 역방향 적용
    result = contents_xml
    for t_start, t_end, new_text in sorted(patches, key=lambda x: -x[0]):
        result = result[:t_start] + new_text + result[t_end:]

    # DSD 저장 (JPG 제거)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, data in orig_files.items():
            if os.path.splitext(name)[1].lower() in ('.jpg','.jpeg','.png','.gif','.bmp'):
                continue
            zf.writestr(name, result.encode('utf-8') if name == 'contents.xml' else data)
    return buf.getvalue()


# ── Flask 앱 ───────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

# ── 하트비트 감시자 ─────────────────────────────────────────────────────────────
# 브라우저가 살아있는 동안 JS가 2.5초마다 /api/heartbeat 를 호출함.
# 8초 이상 핑이 없으면 브라우저 창이 닫힌 것으로 간주 → os._exit(0) 으로 즉시 종료.
_last_ping = time.time()

def _watchdog():
    global _last_ping
    time.sleep(12)          # 앱 시작 직후 12초는 여유 줌 (브라우저 로딩 시간)
    while True:
        time.sleep(2)
        if time.time() - _last_ping > 8:
            os._exit(0)     # 좀비 방지: 프로세스 강제 종료

threading.Thread(target=_watchdog, daemon=True).start()

# ── HTML ───────────────────────────────────────────────────────────────────────
HTML = r'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>easydsd - DART 감사보고서 변환 도구</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Malgun Gothic','맑은 고딕',sans-serif;background:#f0f4f8;color:#1a1a2e;min-height:100vh}
.header{background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);color:white;padding:18px 28px;
  box-shadow:0 4px 20px rgba(31,78,121,.3);display:flex;align-items:center;justify-content:space-between;gap:12px}
.header-left h1{font-size:18px;font-weight:700;letter-spacing:-.5px}
.header-left p{font-size:11px;opacity:.75;margin-top:3px}
.header-right{display:flex;align-items:center;gap:10px;flex-shrink:0}
.header-badge{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);
  border-radius:20px;padding:4px 12px;font-size:11px;font-weight:600;white-space:nowrap}
.kill-btn{background:#c0392b;color:white;border:none;border-radius:8px;padding:7px 14px;
  font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap;transition:background .15s;
  box-shadow:0 2px 8px rgba(192,57,43,.5)}
.kill-btn:hover{background:#e74c3c}
.container{max-width:820px;margin:24px auto;padding:0 18px 60px}
.tabs{display:flex;gap:3px}
.tab{padding:10px 20px;border-radius:10px 10px 0 0;background:#cdd8e4;color:#4a6078;
  font-size:13px;font-weight:600;cursor:pointer;border:none;border-bottom:3px solid transparent;transition:all .2s}
.tab.active{background:white;color:#1F4E79;border-bottom:3px solid #1F4E79}
.tab:hover:not(.active){background:#bcccd8}
.tab.dev-tab{background:#2a2a2a;color:#999}
.tab.dev-tab.active{background:white;color:#333;border-bottom:3px solid #666}
.tab.dev-tab:hover:not(.active){background:#3a3a3a;color:#bbb}
.card{background:white;border-radius:0 12px 12px 12px;box-shadow:0 4px 24px rgba(0,0,0,.08);padding:28px}
.tab-content{display:none}.tab-content.active{display:block}
.step{display:flex;gap:12px;align-items:flex-start;padding:14px;margin-bottom:12px;
  background:#f7f9fc;border-radius:10px;border-left:4px solid #2E75B6}
.step-num{min-width:26px;height:26px;border-radius:50%;background:#1F4E79;color:white;
  display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px;flex-shrink:0}
.step-title{font-weight:700;font-size:13px;color:#1F4E79;margin-bottom:4px}
.step-desc{font-size:12px;color:#556;line-height:1.6}
.drop-zone{border:2px dashed #a0b8d0;border-radius:10px;padding:24px;text-align:center;
  cursor:pointer;transition:all .2s;background:#f7fbff;margin-top:8px}
.drop-zone:hover,.drop-zone.drag-over{border-color:#1F4E79;background:#e8f0f8}
.drop-zone .icon{font-size:28px;margin-bottom:5px}
.drop-zone .label{font-size:13px;color:#4a6078}
.drop-zone .sub{font-size:11px;color:#89a;margin-top:3px}
.file-badge{margin-top:7px;font-size:12px;color:#1F4E79;font-weight:600;display:none;
  background:#e8f0f8;padding:6px 12px;border-radius:6px}
.btn{width:100%;padding:12px;border:none;border-radius:8px;font-size:14px;font-weight:700;
  cursor:pointer;transition:all .2s;margin-top:14px}
.btn-blue{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:white}
.btn-blue:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(31,78,121,.35)}
.btn-blue:disabled{background:#a0b8c8;cursor:not-allowed;transform:none;box-shadow:none}
.btn-green{background:linear-gradient(135deg,#1a6b3a,#22a55a);color:white}
.btn-green:hover{transform:translateY(-1px);box-shadow:0 6px 20px rgba(26,107,58,.35)}
.btn-green:disabled{background:#8fc0a0;cursor:not-allowed;transform:none;box-shadow:none}
.prog-wrap{margin-top:14px;display:none}
.prog-bar{height:7px;background:#e0e8f0;border-radius:4px;overflow:hidden}
.prog-fill{height:100%;width:0%;background:linear-gradient(90deg,#1F4E79,#2E75B6);
  border-radius:4px;transition:width .35s ease}
.prog-text{font-size:11px;color:#4a6078;margin-top:5px;text-align:center}
.result{margin-top:14px;padding:14px 16px;border-radius:10px;display:none;align-items:center;gap:12px}
.result.ok{background:#e8f5ec;border:1px solid #6dbf8a}
.result.err{background:#fdecea;border:1px solid #e88}
.r-icon{font-size:22px}
.r-body{flex:1}
.r-title{font-weight:700;font-size:13px}
.r-sub{font-size:11px;margin-top:3px;color:#556}
.dl-btn{padding:8px 14px;color:white;border:none;border-radius:6px;font-size:12px;font-weight:600;
  cursor:pointer;white-space:nowrap;text-decoration:none;display:inline-block;transition:background .15s}
.dl-btn.green{background:#1a6b3a}.dl-btn.green:hover{background:#145530}
.dl-btn.blue{background:#1F4E79}.dl-btn.blue:hover{background:#163a5e}
.legend{display:flex;gap:12px;flex-wrap:wrap;margin-top:8px}
.leg-item{display:flex;align-items:center;gap:5px;font-size:11px;color:#556}
.leg-dot{width:13px;height:13px;border-radius:3px;flex-shrink:0}
/* 개발자 탭 */
.dev-profile{display:flex;align-items:center;gap:18px;padding:20px;
  background:linear-gradient(135deg,#1a1a2e,#16213e);border-radius:12px;margin-bottom:18px}
.dev-avatar{width:64px;height:64px;border-radius:50%;background:linear-gradient(135deg,#1F4E79,#2E75B6);
  display:flex;align-items:center;justify-content:center;font-size:28px;flex-shrink:0;
  border:3px solid rgba(255,255,255,.2)}
.dev-info h2{color:white;font-size:16px;font-weight:700;margin-bottom:3px}
.dev-sub{color:rgba(255,255,255,.6);font-size:11px;margin-bottom:7px}
.dev-badges{display:flex;gap:6px;flex-wrap:wrap}
.badge{border-radius:20px;padding:3px 10px;font-size:10px;font-weight:600}
.badge-gray{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.8)}
.badge-gold{background:linear-gradient(135deg,#b8860b,#daa520);color:white}
.badge-tech{background:rgba(46,117,182,.5);border:1px solid rgba(46,117,182,.8);color:white}
.info-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px}
.info-box{background:#f7f9fc;border-radius:10px;padding:12px 14px;border-left:3px solid #2E75B6}
.info-box .lbl{font-size:10px;color:#89a;text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px}
.info-box .val{font-size:14px;font-weight:700;color:#1F4E79}
.info-box .val a{color:#1F4E79;text-decoration:none}.info-box .val a:hover{text-decoration:underline}
.credit-box{background:#fffbf0;border:1px solid #e8d060;border-radius:12px;
  padding:18px;text-align:center;margin-bottom:16px}
.credit-title{font-size:13px;font-weight:700;color:#7a5500;margin-bottom:10px}
.credit-body{font-size:12px;color:#444;line-height:1.9}
.credit-name{font-size:16px;font-weight:800;color:#1a1a2e;margin:5px 0 2px}
.credit-sub{font-size:11px;color:#888}
.claude-chip{display:inline-block;background:linear-gradient(135deg,#7c4dff,#2196f3);
  color:white;border-radius:20px;padding:3px 12px;font-size:11px;font-weight:700;
  margin:0 4px;vertical-align:middle}
.feat-section h3{font-size:12px;font-weight:700;color:#1F4E79;margin-bottom:8px}
.feat-item{display:flex;align-items:flex-start;gap:8px;padding:8px 0;
  border-bottom:1px solid #f0f4f8;font-size:11px;color:#446;line-height:1.5}
.feat-item:last-child{border-bottom:none}
.feat-ico{font-size:14px;flex-shrink:0;margin-top:1px}
.feat-wip{color:#999;font-size:10px;margin-left:4px}
/* 종료 확인 모달 */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);
  z-index:9999;align-items:center;justify-content:center}
.modal-overlay.show{display:flex}
.modal{background:white;border-radius:14px;padding:28px 32px;max-width:360px;width:90%;text-align:center;
  box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal h3{font-size:16px;font-weight:700;color:#1a1a2e;margin-bottom:8px}
.modal p{font-size:13px;color:#556;margin-bottom:20px;line-height:1.6}
.modal-btns{display:flex;gap:10px;justify-content:center}
.modal-btns button{padding:10px 24px;border:none;border-radius:8px;font-size:13px;
  font-weight:700;cursor:pointer;transition:all .15s}
.modal-cancel{background:#e8eef4;color:#4a6078}
.modal-cancel:hover{background:#d0dce8}
.modal-confirm{background:#c0392b;color:white;box-shadow:0 3px 10px rgba(192,57,43,.4)}
.modal-confirm:hover{background:#e74c3c}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>&#128202; DART 감사보고서 변환 도구</h1>
    <p>DSD &#8596; Excel 양방향 변환 &nbsp;&#xB7;&nbsp; 재무/비재무정보 전체 수정 가능</p>
  </div>
  <div class="header-right">
    <div class="header-badge">easydsd v0.1</div>
    <button class="kill-btn" onclick="showKillModal()">&#x23FC; 프로그램 종료</button>
  </div>
</div>

<!-- 종료 확인 모달 -->
<div class="modal-overlay" id="killModal">
  <div class="modal">
    <h3>&#x26A0;&#xFE0F; 프로그램을 종료할까요?</h3>
    <p>서버가 완전히 종료됩니다.<br>브라우저 창도 함께 닫힙니다.</p>
    <div class="modal-btns">
      <button class="modal-cancel" onclick="hideKillModal()">취소</button>
      <button class="modal-confirm" onclick="doKill()">종료</button>
    </div>
  </div>
</div>

<div class="container">
  <div class="tabs">
    <button class="tab active" onclick="switchTab(0)">&#9312; DSD &#8594; Excel</button>
    <button class="tab" onclick="switchTab(1)">&#9313; Excel &#8594; DSD</button>
    <button class="tab dev-tab" onclick="switchTab(2)">개발자 정보</button>
  </div>
  <div class="card">

    <!-- 탭①: DSD→Excel -->
    <div class="tab-content active" id="tab0">
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <div class="step-title">전년도 DSD 파일을 업로드하세요</div>
          <div class="step-desc">DART에서 제출한 감사보고서 .dsd 파일을 드래그하거나 클릭해 선택하세요.<br>변환된 Excel의 <b>노란색 셀</b>을 당해년도 숫자로 수정하시면 됩니다.</div>
          <div class="drop-zone" id="dz1" onclick="document.getElementById('f1').click()"
               ondragover="dragOver(event,'dz1')" ondragleave="dragLeave('dz1')" ondrop="drop(event,'f1','dz1')">
            <div class="icon">&#128194;</div>
            <div class="label">클릭하거나 파일을 여기에 끌어다 놓으세요</div>
            <div class="sub">.dsd 파일만 가능합니다</div>
          </div>
          <input type="file" id="f1" accept=".dsd" style="display:none" onchange="setFile('f1','fb1','dz1')">
          <div class="file-badge" id="fb1"></div>
        </div>
      </div>
      <button class="btn btn-blue" id="btn1" onclick="run1()" disabled>
        &#128229;&nbsp; Excel 파일로 변환하기
      </button>
      <div class="prog-wrap" id="pw1">
        <div class="prog-bar"><div class="prog-fill" id="pf1"></div></div>
        <div class="prog-text" id="pt1">변환 중...</div>
      </div>
      <div class="result ok" id="ok1">
        <div class="r-icon">&#9989;</div>
        <div class="r-body">
          <div class="r-title" id="ok1t"></div>
          <div class="r-sub"  id="ok1s"></div>
          <div class="legend" style="margin-top:8px">
            <div class="leg-item"><div class="leg-dot" style="background:#FFF2CC;border:1px solid #ccc"></div>노란색 = 수정 가능</div>
            <div class="leg-item"><div class="leg-dot" style="background:#1F4E79"></div>파란색 = 헤더(수정불필요)</div>
            <div class="leg-item"><div class="leg-dot" style="background:#D9D9D9;border:1px solid #bbb"></div>회색줄 = 주석 구분선</div>
          </div>
        </div>
        <a class="dl-btn green" id="dl1" href="#">&#11015; 다운로드</a>
      </div>
      <div class="result err" id="er1">
        <div class="r-icon">&#10060;</div>
        <div class="r-body"><div class="r-title">변환 실패</div><div class="r-sub" id="er1m"></div></div>
      </div>
    </div>

    <!-- 탭②: Excel→DSD -->
    <div class="tab-content" id="tab1">
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <div class="step-title">원본 DSD 파일 업로드</div>
          <div class="step-desc">&#9312; 탭에서 변환할 때 사용했던 원본 .dsd 파일을 올려주세요.</div>
          <div class="drop-zone" id="dz2" onclick="document.getElementById('f2').click()"
               ondragover="dragOver(event,'dz2')" ondragleave="dragLeave('dz2')" ondrop="drop(event,'f2','dz2')">
            <div class="icon">&#128194;</div>
            <div class="label">원본 DSD 파일</div>
            <div class="sub">.dsd 파일</div>
          </div>
          <input type="file" id="f2" accept=".dsd" style="display:none" onchange="setFile('f2','fb2','dz2')">
          <div class="file-badge" id="fb2"></div>
        </div>
      </div>
      <div class="step">
        <div class="step-num">2</div>
        <div class="step-body">
          <div class="step-title">수정한 Excel 파일 업로드</div>
          <div class="step-desc">&#9312; 탭에서 다운로드해 노란색 셀을 수정한 .xlsx 파일을 올려주세요.</div>
          <div class="drop-zone" id="dz3" onclick="document.getElementById('f3').click()"
               ondragover="dragOver(event,'dz3')" ondragleave="dragLeave('dz3')" ondrop="drop(event,'f3','dz3')">
            <div class="icon">&#128202;</div>
            <div class="label">수정된 Excel 파일</div>
            <div class="sub">.xlsx 파일</div>
          </div>
          <input type="file" id="f3" accept=".xlsx" style="display:none" onchange="setFile('f3','fb3','dz3')">
          <div class="file-badge" id="fb3"></div>
        </div>
      </div>
      <button class="btn btn-green" id="btn2" onclick="run2()" disabled>
        &#128228;&nbsp; DSD 파일로 변환하기
      </button>
      <div class="prog-wrap" id="pw2">
        <div class="prog-bar"><div class="prog-fill" id="pf2"></div></div>
        <div class="prog-text" id="pt2">변환 중...</div>
      </div>
      <div class="result ok" id="ok2">
        <div class="r-icon">&#9989;</div>
        <div class="r-body"><div class="r-title" id="ok2t"></div><div class="r-sub" id="ok2s"></div></div>
        <a class="dl-btn blue" id="dl2" href="#">&#11015; DSD 다운로드</a>
      </div>
      <div class="result err" id="er2">
        <div class="r-icon">&#10060;</div>
        <div class="r-body"><div class="r-title">변환 실패</div><div class="r-sub" id="er2m"></div></div>
      </div>
    </div>

    <!-- 탭③: 개발자 정보 -->
    <div class="tab-content" id="tab2">
      <div class="dev-profile">
        <div class="dev-avatar">&#127970;</div>
        <div class="dev-info">
          <h2>Easydsd 0.1v</h2>
          <div class="dev-sub">DART 감사보고서 DSD 파일 변환 도구(양방향)</div>
          <div class="dev-badges">
            <span class="badge badge-gray">v0.1</span>
            <span class="badge badge-gold">&#129302; AI-Powered</span>
            <span class="badge badge-tech">Python + Flask</span>
          </div>
        </div>
      </div>
      <div class="info-grid">
        <div class="info-box">
          <div class="lbl">개발자 연락처</div>
          <div class="val"><a href="mailto:eeffco11@naver.com">eeffco11@naver.com</a></div>
        </div>
        <div class="info-box">
          <div class="lbl">버전</div>
          <div class="val">Easydsd 0.1v</div>
        </div>
        <div class="info-box">
          <div class="lbl">지원 파일</div>
          <div class="val">.dsd / .xlsx</div>
        </div>
        <div class="info-box">
          <div class="lbl">대상</div>
          <div class="val">감사보고서</div>
        </div>
      </div>
      <div class="credit-box">
        <div class="credit-title">&#128591; 제작 크레딧</div>
        <div class="credit-body">
          이 프로그램은 전적으로<br>
          <span class="claude-chip">Claude (Anthropic)</span> 가 설계하고 개발했습니다.
          <div class="credit-name">클로드 짱짱맨</div>
          <div class="credit-sub">전 과정을 클로드로 다함</div>
        </div>
      </div>
      <div class="feat-section">
        <h3>&#10024; 주요 기능</h3>
        <div class="feat-item"><div class="feat-ico">&#127974;</div><div>재무상태표 · 포괄손익계산서 · 자본변동표 · 현금흐름표 전체 편집</div></div>
        <div class="feat-item"><div class="feat-ico">&#128221;</div><div>주석 전체 편집 — 주주명, 지분율, 이자율, 텍스트 포함</div></div>
        <div class="feat-item"><div class="feat-ico">&#128260;</div><div>DSD → Excel → DSD 완전한 양방향 변환, XML 유효성 자동 검증</div></div>
        <div class="feat-item"><div class="feat-ico">&#128737;&#65039;</div><div>비재무 테이블(목차, 감사의견 등) 원본 보존 <span class="feat-wip">— 자꾸 꼬여서 나중에 수정할 예정</span></div></div>
        <div class="feat-item"><div class="feat-ico">&#128433;&#65039;</div><div>드래그 앤 드롭 UI — Python 없이 .exe 한 번 클릭으로 실행</div></div>
        <div class="feat-item"><div class="feat-ico">&#128163;</div><div>하트비트 감시 — 브라우저 닫으면 서버 자동 종료 (좀비 프로세스 방지)</div></div>
      </div>
    </div>

  </div>
</div>

<script>
// ── 하트비트: 2.5초마다 서버에 핑 ──────────────────────────────────────────────
setInterval(function(){
  fetch('/api/heartbeat', {method:'POST'}).catch(function(){});
}, 2500);

// ── 파일 / 드래그 ─────────────────────────────────────────────────────────────
const F = {f1:null, f2:null, f3:null};
function switchTab(n){
  document.querySelectorAll('.tab').forEach((t,i)=>t.classList.toggle('active',i===n));
  document.querySelectorAll('.tab-content').forEach((t,i)=>t.classList.toggle('active',i===n));
}
function setFile(id,bid,dzId){
  const f=document.getElementById(id).files[0]; if(!f) return;
  F[id]=f;
  const b=document.getElementById(bid);
  b.textContent='✓  '+f.name+'  ('+(f.size/1024).toFixed(0)+' KB)';
  b.style.display='block';
  document.getElementById(dzId).style.borderColor='#1F4E79';
  chk();
}
function dragOver(e,id){e.preventDefault();document.getElementById(id).classList.add('drag-over')}
function dragLeave(id){document.getElementById(id).classList.remove('drag-over')}
function drop(e,fid,did){
  e.preventDefault(); dragLeave(did);
  const dt=e.dataTransfer; if(!dt.files.length) return;
  const inp=document.getElementById(fid);
  const tr=new DataTransfer(); tr.items.add(dt.files[0]); inp.files=tr.files;
  setFile(fid, fid.replace('f','fb'), did);
}
function chk(){
  document.getElementById('btn1').disabled=!F.f1;
  document.getElementById('btn2').disabled=!(F.f2&&F.f3);
}
function hide(n){['ok','er'].forEach(p=>document.getElementById(p+n).style.display='none')}
function showOk(n,title,sub,blob,fname){
  const b=document.getElementById('ok'+n); b.style.display='flex';
  document.getElementById('ok'+n+'t').textContent=title;
  document.getElementById('ok'+n+'s').textContent=sub;
  const dl=document.getElementById('dl'+n);
  dl.href=URL.createObjectURL(blob); dl.download=fname;
}
function showErr(n,msg){
  const b=document.getElementById('er'+n); b.style.display='flex';
  document.getElementById('er'+n+'m').textContent=msg;
}

// ── 프로그레스 바 ─────────────────────────────────────────────────────────────
let progIv=null;
function startProg(n,msg){
  hide(n);
  const pw=document.getElementById('pw'+n); pw.style.display='block';
  document.getElementById('pt'+n).textContent=msg;
  document.getElementById('pf'+n).style.width='0%';
  let w=0; progIv=setInterval(()=>{w=Math.min(w+4,88);document.getElementById('pf'+n).style.width=w+'%';},200);
}
function endProg(n){
  clearInterval(progIv);
  document.getElementById('pf'+n).style.width='100%';
  setTimeout(()=>document.getElementById('pw'+n).style.display='none',500);
}
const S1=['DSD 파일 분석 중...','테이블 구조 파싱 중...','Excel 시트 생성 중...'];
const S2=['매핑 구성 중...','XML 패치 적용 중...','DSD 파일 생성 중...'];
function animS(n,steps){let i=0;return setInterval(()=>{if(i<steps.length)document.getElementById('pt'+n).textContent=steps[i++];},1000)}

// ── 변환 요청 ─────────────────────────────────────────────────────────────────
async function run1(){
  if(!F.f1) return;
  document.getElementById('btn1').disabled=true;
  startProg(1,S1[0]); const iv=animS(1,S1);
  try{
    const fd=new FormData(); fd.append('dsd',F.f1);
    const r=await fetch('/api/dsd2excel',{method:'POST',body:fd});
    clearInterval(iv); endProg(1);
    if(!r.ok){const e=await r.json();throw new Error(e.error||'변환 실패');}
    const blob=await r.blob();
    const info=JSON.parse(r.headers.get('X-Info')||'{}');
    const fname=F.f1.name.replace(/\.dsd$/i,'')+'.xlsx';
    showOk(1,'변환 완료! Excel 파일을 다운로드하세요',
      '시트 '+info.sheets+'개 · 수정가능 셀 '+info.cells+'개 · 핵심재무표 '+info.fin+'개', blob, fname);
  }catch(e){clearInterval(iv);endProg(1);showErr(1,e.message);}
  document.getElementById('btn1').disabled=false;
}
async function run2(){
  if(!F.f2||!F.f3) return;
  document.getElementById('btn2').disabled=true;
  startProg(2,S2[0]); const iv=animS(2,S2);
  try{
    const fd=new FormData(); fd.append('orig_dsd',F.f2); fd.append('xlsx',F.f3);
    const r=await fetch('/api/excel2dsd',{method:'POST',body:fd});
    clearInterval(iv); endProg(2);
    if(!r.ok){const e=await r.json();throw new Error(e.error||'변환 실패');}
    const blob=await r.blob();
    const info=JSON.parse(r.headers.get('X-Info')||'{}');
    const fname=F.f2.name.replace(/\.dsd$/i,'')+'_수정.dsd';
    showOk(2,'DSD 파일 생성 완료!',
      info.tables+'개 테이블 · '+info.cells+'개 셀 수정 · XML 검증 '+(info.xml_ok?'✓ 정상':'✗ 오류'), blob, fname);
  }catch(e){clearInterval(iv);endProg(2);showErr(2,e.message);}
  document.getElementById('btn2').disabled=false;
}

// ── 종료 모달 ─────────────────────────────────────────────────────────────────
function showKillModal(){ document.getElementById('killModal').classList.add('show') }
function hideKillModal(){ document.getElementById('killModal').classList.remove('show') }
async function doKill(){
  hideKillModal();
  try{ await fetch('/api/shutdown',{method:'POST'}); }catch(e){}
  document.body.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif;color:#556;font-size:15px;">서버가 종료되었습니다. 이 탭을 닫으세요.</div>';
}
</script>
</body>
</html>'''

# ── API 라우트 ─────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/api/heartbeat', methods=['POST'])
def api_heartbeat():
    global _last_ping
    _last_ping = time.time()
    return jsonify(ok=True)

@app.route('/api/shutdown', methods=['POST'])
def api_shutdown():
    """[프로그램 종료] 버튼에서 호출 → 즉시 프로세스 종료"""
    threading.Thread(target=lambda: (time.sleep(0.3), os._exit(0)), daemon=True).start()
    return jsonify(ok=True)

@app.route('/api/dsd2excel', methods=['POST'])
def api_dsd2excel():
    try:
        xlsx = dsd_to_excel_bytes(request.files['dsd'].read())
        wb = openpyxl.load_workbook(io.BytesIO(xlsx), data_only=True)
        cells = sum(1 for ws in wb.worksheets for row in ws.iter_rows()
                    for cell in row if cell.fill and cell.fill.fill_type == 'solid'
                    and cell.fill.fgColor and cell.fill.fgColor.type == 'rgb'
                    and cell.fill.fgColor.rgb.upper().endswith(EDIT_COLOR.upper()))
        fin = [ws.title for ws in wb.worksheets
               if any(ws.title.startswith(e) for e in ('🏦','💹','📈','💰'))]
        import json
        resp = send_file(io.BytesIO(xlsx),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name='converted.xlsx')
        resp.headers['X-Info'] = json.dumps({'sheets':len(wb.sheetnames),'cells':cells,'fin':len(fin)})
        return resp
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.route('/api/excel2dsd', methods=['POST'])
def api_excel2dsd():
    try:
        orig = request.files['orig_dsd'].read()
        xlsx = request.files['xlsx'].read()
        dsd  = excel_to_dsd_bytes(orig, xlsx)
        import xml.etree.ElementTree as ET, json
        with zipfile.ZipFile(io.BytesIO(dsd)) as z:
            xml_text = z.read('contents.xml').decode('utf-8')
        xml_ok = True
        try: ET.fromstring(xml_text)
        except: xml_ok = False
        wb = openpyxl.load_workbook(io.BytesIO(xlsx), data_only=True)
        tc = tb = 0
        for sname in wb.sheetnames:
            if sname in ('📋사용안내','_원본XML','📊요약수치'): continue
            ws = wb[sname]
            cnt = sum(1 for row in ws.iter_rows() for cell in row
                      if cell.fill and cell.fill.fill_type == 'solid'
                      and cell.fill.fgColor and cell.fill.fgColor.type == 'rgb'
                      and cell.fill.fgColor.rgb.upper().endswith(EDIT_COLOR.upper()))
            if cnt: tb += 1; tc += cnt
        resp = send_file(io.BytesIO(dsd), mimetype='application/octet-stream',
                         as_attachment=True, download_name='output.dsd')
        resp.headers['X-Info'] = json.dumps({'tables':tb,'cells':tc,'xml_ok':xml_ok})
        return resp
    except Exception as e:
        return jsonify(error=str(e)), 500

# ── 실행 ───────────────────────────────────────────────────────────────────────
def open_browser():
    time.sleep(1.5)
    webbrowser.open(f'http://127.0.0.1:{PORT}')

if __name__ == '__main__':
    print('='*50)
    print('  easydsd v0.1 - DART 감사보고서 변환 도구')
    print(f'  http://127.0.0.1:{PORT}')
    print('  종료: 브라우저 닫기 or 종료 버튼')
    print('='*50)
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='127.0.0.1', port=PORT, debug=False)
