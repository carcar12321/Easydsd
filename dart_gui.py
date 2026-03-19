#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""easydsd v0.7 - DART 감사보고서 변환 도구 + Gemini AI"""

import os, re, sys, io, zipfile, threading, webbrowser, socket, time, json

# Windows 콘솔 UTF-8 강제 (cp949 오류 방지)
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# IS_FROZEN: Flask import보다 먼저 선언 (순서 중요)
IS_FROZEN = getattr(sys, 'frozen', False) or hasattr(sys, '_MEIPASS')
BASE_DIR  = os.path.dirname(sys.executable if IS_FROZEN else os.path.abspath(__file__))

# ── Flask / openpyxl ─────────────────────────────────────────────────────────
try:
    from flask import Flask, request, send_file, jsonify, render_template_string
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    if IS_FROZEN:
        print("[ERROR] 필수 라이브러리 누락. EXE를 다시 빌드하세요.")
        sys.exit(1)
    import subprocess
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'flask', 'openpyxl', '-q'])
    from flask import Flask, request, send_file, jsonify, render_template_string
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter

# ── google-generativeai (Raw import: PyInstaller 의존성 인식용) ──────────────
import google.generativeai as genai  # noqa: E402

# ── 기본 상수 ─────────────────────────────────────────────────────────────────
def find_free_port(start=5000, end=5099):
    for p in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', p)) != 0: return p
    return start

PORT       = find_free_port()
EDIT_COLOR = 'FFF2CC'
C = {'navy':'1F4E79','blue':'2E75B6','lblue':'DEEAF1',
     'yellow':'FFF2CC','white':'FFFFFF','lgray':'F2F2F2','orange':'C55A11'}
FIN_TABLE_MAP = [
    (['재 무 상 태 표'],       '🏦재무상태표'),
    (['포 괄 손 익 계 산 서'], '💹포괄손익계산서'),
    (['자 본 변 동 표'],       '📈자본변동표'),
    (['현 금 흐 름 표'],       '💰현금흐름표'),
]
EXT_DESC = {
    'TOT_ASSETS':'총자산(백만원)','TOT_DEBTS':'총부채(백만원)',
    'TOT_SALES':'매출액(백만원)','TOT_EMPL':'총직원수',
    'GMSH_DATE':'주총일자(YYYYMMDD)','SUPV_OPIN':'감사의견코드',
    'AUDIT_CIK':'감사인CIK','CRP_RGS_NO':'법인등록번호',
}
FIN_PREFIXES = ('🏦','💹','📈','💰')

def fill(c): return PatternFill('solid', fgColor=c)
def fnt(color='000000',bold=False,size=9,italic=False):
    return Font(color=color,bold=bold,size=size,italic=italic)
def aln(h='left',v='center',wrap=False):
    return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

# ── XML 파싱 ──────────────────────────────────────────────────────────────────
def clean_cr(s, as_newline=False):
    repl='\n' if as_newline else ' '
    s=s.replace('&amp;cr;',repl).replace('&cr;',repl)
    return re.sub(r'\s+',' ',s).strip()

def clean_title(s):
    s=clean_cr(s,False)
    s=s.replace('&amp;','&').replace('&lt;','<').replace('&gt;','>').replace('&quot;','"')
    return re.sub(r'\s+',' ',s).strip()

def is_blank_title(s):
    return len(re.sub(r'[&;a-z]+','',s).strip())==0

def parse_cell(m):
    attrs=m.group(1)
    val=re.sub(r'<[^>]+>','',m.group(0))
    val=(val.replace('&amp;cr;','\n').replace('&amp;','&')
           .replace('&lt;','<').replace('&gt;','>').replace('&quot;','"')
           .replace('&cr;','\n').strip())
    cs=int(x.group(1)) if (x:=re.search(r'COLSPAN="(\d+)"',attrs)) else 1
    tag=re.match(r'<([A-Z]+)',m.group(0)).group(1)
    return dict(value=val,colspan=cs,tag=tag)

def is_num_or_decimal(val):
    v=(val.strip().replace(',','').replace('(','').replace(')','')
         .replace('%','').replace('-','').replace(' ','').split('\n')[0])
    if not v: return False
    try: float(v); return True
    except: return False

def parse_xml(xml):
    exts=re.findall(r'<EXTRACTION[^>]*ACODE="([^"]+)"[^>]*>([^<]+)</EXTRACTION>',xml)
    tables=[]
    for ti,tm in enumerate(re.finditer(r'<TABLE([^>]*)>(.*?)</TABLE>',xml,re.DOTALL)):
        ctx=xml[max(0,tm.start()-600):tm.start()]
        tbody=tm.group(0)
        fin_label=next((lbl for kws,lbl in FIN_TABLE_MAP
                        if any(kw in ctx or kw in tbody for kw in kws)),'')
        raw=re.findall(r'<(?:TITLE|P)[^>]*>([^<]{3,80})</(?:TITLE|P)>',ctx)
        ctx_titles=[clean_title(t) for t in raw if not is_blank_title(t) and len(clean_title(t))>1]
        rows=[]
        for tr in re.finditer(r'<TR[^>]*>(.*?)</TR>',tm.group(2),re.DOTALL):
            cells=[parse_cell(cm) for cm in re.finditer(
                r'<(?:TD|TH|TU|TE)([^>]*)>.*?</(?:TD|TH|TU|TE)>',tr.group(1),re.DOTALL)]
            if cells: rows.append(cells)
        tables.append(dict(idx=ti,fin_label=fin_label,
                           ctx_title=(ctx_titles[-1] if ctx_titles else ''),
                           rows=rows,start=tm.start()))
    return exts,tables

# ── cell 헬퍼 ─────────────────────────────────────────────────────────────────
def is_edit(cell):
    f=cell.fill
    if f and f.fill_type=='solid':
        fg=f.fgColor
        if fg and fg.type=='rgb': return fg.rgb.upper().endswith(EDIT_COLOR.upper())
    return False

def is_navy(cell):
    f=cell.fill
    if f and f.fill_type=='solid':
        fg=f.fgColor
        if fg and fg.type=='rgb': return fg.rgb.upper().endswith(C['navy'].upper())
    return False

def cell_num(v):
    """셀 값을 숫자(float)로 변환, 실패 시 None"""
    if v is None: return None
    s=str(v).strip().replace(',','').replace('(','').replace(')','').replace('-','').replace(' ','')
    if not s: return None
    try:
        raw=str(v).strip().replace(',','')
        neg=(str(v).strip().startswith('(') and str(v).strip().endswith(')')) or str(v).strip().startswith('-')
        n=float(s)
        return -n if neg else n
    except: return None

# ── 기능1: 롤오버 (당기→전기 이월) ──────────────────────────────────────────
def apply_rollover(wb):
    """
    재무제표 시트의 당기 값을 전기로 이월, 당기는 빈칸으로.
    각 데이터 행에서 yellow 숫자 셀을 찾아 앞절반=당기, 뒷절반=전기로 처리.
    """
    for sname in wb.sheetnames:
        if not any(sname.startswith(p) for p in FIN_PREFIXES): continue
        ws=wb[sname]
        for rowi in range(1, ws.max_row+1):
            num_cells=[]
            for ci in range(1, ws.max_column+1):
                cell=ws.cell(rowi,ci)
                if is_edit(cell) and cell.value is not None:
                    v=str(cell.value).strip().replace(',','').replace('(','').replace(')','').replace('-','')
                    if v and v.replace('.','').isdigit() and len(v)>=3:
                        num_cells.append((ci,cell))
            if len(num_cells)<2: continue
            half=len(num_cells)//2
            for (_cc,c_cell),(_pc,p_cell) in zip(num_cells[:half], num_cells[half:]):
                p_cell.value=c_cell.value
                c_cell.value=None


# ── 기능2/4: Python 수학적 검증 ──────────────────────────────────────────────
def python_verify(xlsx_bytes:bytes) -> dict:
    """
    순수 파이썬 산술/논리 검증.
    반환: {errors:[], warnings:[], info:[]}
    """
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes),data_only=True)
    errors=[]; warnings=[]; info=[]

    # ─ 1. 대차평균 검증 (재무상태표) ──────────────────────────────────────────
    bs_sheet=next((wb[s] for s in wb.sheetnames if s.startswith('🏦')),None)
    if bs_sheet:
        asset_total=liab_total=equity_total=None
        for row in bs_sheet.iter_rows(values_only=True):
            vals=[str(v or '') for v in row]
            row_txt=' '.join(vals)
            nums=[cell_num(v) for v in row if cell_num(v) is not None]
            if not nums: continue
            n=nums[0]
            if any(k in row_txt for k in ['자산총계','자산 총계','총자산']):
                asset_total=n
            elif any(k in row_txt for k in ['부채총계','부채 총계','총부채']):
                liab_total=n
            elif any(k in row_txt for k in ['자본총계','자본 총계','총자본']):
                equity_total=n
        if all(v is not None for v in [asset_total,liab_total,equity_total]):
            diff=abs(asset_total-(liab_total+equity_total))
            tol=max(abs(asset_total)*1e-6,1)
            if diff>tol:
                errors.append(
                    f"[대차불일치] 자산총계({asset_total:,.0f}) != 부채({liab_total:,.0f})+자본({equity_total:,.0f}), 차이={diff:,.0f}")
            else:
                info.append(f"[대차일치] 자산={asset_total:,.0f} = 부채+자본")
        else:
            warnings.append("[대차검증] 자산총계/부채총계/자본총계 중 일부를 찾지 못했습니다.")

    # ─ 2. 단위(자릿수) 스케일 이상 탐지 ──────────────────────────────────────
    fin_nums=[]
    note_nums=[]
    for sname in wb.sheetnames:
        ws=wb[sname]
        for row in ws.iter_rows(max_row=150,values_only=True):
            for v in row:
                n=cell_num(v)
                if n is not None and abs(n)>0:
                    if any(sname.startswith(p) for p in FIN_PREFIXES): fin_nums.append(abs(n))
                    elif sname.startswith('📝'): note_nums.append(abs(n))
    if fin_nums and note_nums:
        fin_med=sorted(fin_nums)[len(fin_nums)//2]
        note_med=sorted(note_nums)[len(note_nums)//2]
        if fin_med>0 and note_med>0:
            ratio=max(fin_med,note_med)/min(fin_med,note_med)
            if ratio>5000:
                warnings.append(
                    f"[단위불일치 의심] 재무제표 중간값({fin_med:,.0f}) vs 주석 중간값({note_med:,.0f}), "
                    f"배율={ratio:.0f}배 - 원/천원/백만원 단위 혼재 가능성")

    # ─ 3. 주석 번호 꼬임 방지 ─────────────────────────────────────────────────
    # 재무제표 본문에서 주석 번호 추출
    fin_note_refs=set()
    for sname in wb.sheetnames:
        if not any(sname.startswith(p) for p in FIN_PREFIXES): continue
        ws=wb[sname]
        for row in ws.iter_rows(max_row=200,values_only=True):
            for v in row:
                if v is None: continue
                # 주석 참조 패턴: (주N), 주N, N,M,K 형태
                for m in re.finditer(r'(?:주\s*|주석\s*)?(\d{1,2})(?:\s*,\s*(\d{1,2}))*', str(v)):
                    nums=[int(x) for x in re.findall(r'\d{1,2}',m.group(0))]
                    fin_note_refs.update(n for n in nums if 1<=n<=99)

    # 실제 존재하는 주석 시트에서 번호 추출
    existing_notes=set()
    for sname in wb.sheetnames:
        m=re.search(r'(\d{1,2})',sname)
        if m and sname.startswith('📝'):
            existing_notes.add(int(m.group(1)))

    if fin_note_refs and existing_notes:
        # 재무제표가 참조하지만 시트 없는 주석
        missing={n for n in fin_note_refs if n not in existing_notes and n>5}
        if missing:
            warnings.append(f"[주석 매핑 불완전] 재무제표 참조 주석번호 {sorted(missing)} 에 해당하는 시트 없음")
        else:
            info.append(f"[주석 매핑] 참조된 주석 {len(fin_note_refs)}개 모두 시트 존재")

    return {'errors':errors,'warnings':warnings,'info':info}

# ── Gemini: AI 시트명 분류 ────────────────────────────────────────────────────
def gemini_classify_tables(api_key:str,tables:list,model_name:str='gemini-3-flash-preview')->dict:
    if not api_key: return {}
    try:
        genai.configure(api_key=api_key)
        model=genai.GenerativeModel(model_name)
        summaries=[]
        for tbl in tables[:60]:
            vals=[c['value'] for row in tbl['rows'][:3] for c in row if c['value'].strip()][:6]
            summaries.append(f"TABLE[{tbl['idx']}] ctx={tbl['ctx_title']!r} fin={tbl['fin_label']!r} sample={vals}")
        prompt=(
            "한국 DART 감사보고서 TABLE 목록입니다. 각 TABLE의 엑셀 시트명을 제안해주세요.\n"
            "규칙: 재무상태표->'🏦재무상태표', 포괄손익->'💹포괄손익계산서', 자본변동->'📈자본변동표', 현금흐름->'💰현금흐름표',\n"
            "주석->'📝주석_[주제3~5자]', 서문/목차/감사의견->'📄서문', 31자이내\n\n"
            f"TABLE:\n{chr(10).join(summaries)}\n\n"
            'JSON만 응답: {"mapping":[{"idx":0,"name":"예시"}]}'
        )
        resp=model.generate_content(prompt)
        m=re.search(r'\{.*\}',resp.text.strip(),re.DOTALL)
        if not m: return {}
        data=json.loads(m.group(0))
        return {item['idx']:item['name'] for item in data.get('mapping',[])}
    except Exception as e:
        print(f"[Gemini classify] {e}"); return {}

# ── Gemini: 강화된 AI 교차 검증 (Python 오류 포함) ───────────────────────────
def gemini_verify_enhanced(api_key:str,fin_data:dict,note_data:dict,
                            py_result:dict,model_name:str='gemini-3-flash-preview')->str:
    if not api_key:
        return "Gemini API Key가 없습니다."
    try:
        genai.configure(api_key=api_key)
        model=genai.GenerativeModel(model_name)
        fin_text =json.dumps(fin_data, ensure_ascii=False)[:6000]
        note_text=json.dumps(note_data,ensure_ascii=False)[:6000]
        py_errors_text='\n'.join(
            [f"[오류] {e}" for e in py_result.get('errors',[])] +
            [f"[경고] {w}" for w in py_result.get('warnings',[])] +
            [f"[정보] {i}" for i in py_result.get('info',[])]
        ) or "파이썬 자동검사 이상 없음"
        prompt=(
            "당신은 한국 공인회계사(CPA) 수준의 재무제표 검증 전문가입니다.\n\n"
            f"[파이썬 수학적 1차 검사 결과]\n{py_errors_text}\n\n"
            f"[재무제표 본문]\n{fin_text}\n\n"
            f"[주석 데이터]\n{note_text}\n\n"
            "파이썬이 찾은 수학적 오류 내용과 재무제표 맥락을 합쳐서 최종 검증 리포트를 작성해 주세요.\n\n"
            "응답 형식:\n"
            "## ✅ 파이썬 수학 검사 결과\n(자동검사 결과 요약)\n\n"
            "## ✅ 일치 항목\n(AI가 확인한 일치 항목)\n\n"
            "## ❌ 불일치 항목\n(불일치 + 차이 금액)\n\n"
            "## ⚠️ 확인 필요 항목\n(데이터 부족 등)\n\n"
            "## 📋 종합 의견\n(전체 요약)"
        )
        resp=model.generate_content(prompt)
        return resp.text.strip()
    except Exception as e:
        return f"Gemini API 오류: {e}"

# ── 기능3: DSD 비교 분석 ──────────────────────────────────────────────────────
def parse_dsd_tables(dsd_bytes:bytes)->dict:
    """DSD에서 TABLE별 데이터 추출: {table_idx: [[cell_vals,...], ...]}"""
    with zipfile.ZipFile(io.BytesIO(dsd_bytes)) as zf:
        xml=zf.read('contents.xml').decode('utf-8',errors='replace')
    exts=dict(re.findall(r'<EXTRACTION[^>]*ACODE="([^"]+)"[^>]*>([^<]+)</EXTRACTION>',xml))
    tables={}
    for ti,tm in enumerate(re.finditer(r'<TABLE[^>]*>(.*?)</TABLE>',xml,re.DOTALL)):
        rows=[]
        for tr in re.finditer(r'<TR[^>]*>(.*?)</TR>',tm.group(1),re.DOTALL):
            cells=[]
            for td in re.finditer(r'<(?:TD|TH|TU|TE)[^>]*>(.*?)</(?:TD|TH|TU|TE)>',tr.group(1),re.DOTALL):
                v=re.sub(r'<[^>]+>','',td.group(1))
                v=(v.replace('&amp;cr;',' ').replace('&amp;','&').replace('&lt;','<')
                   .replace('&gt;','>').replace('&cr;',' ').strip())
                cells.append(v)
            if cells: rows.append(cells)
        if rows: tables[ti]=rows
    return {'tables':tables,'exts':exts}

def compare_dsd_bytes(dsd_a:bytes,dsd_b:bytes,api_key:str,model_name:str='gemini-3-flash-preview')->bytes:
    """두 DSD 파일 비교 → 차이 강조 Excel 반환"""
    data_a=parse_dsd_tables(dsd_a)
    data_b=parse_dsd_tables(dsd_b)

    # EXTRACTION 차이
    ext_diffs=[]
    for k in set(list(data_a['exts'].keys())+list(data_b['exts'].keys())):
        va=data_a['exts'].get(k,'(없음)'); vb=data_b['exts'].get(k,'(없음)')
        if va!=vb: ext_diffs.append((k,va,vb))

    # TABLE별 차이
    all_table_ids=sorted(set(list(data_a['tables'].keys())+list(data_b['tables'].keys())))
    table_diffs=[]  # (table_idx, row_i, col_j, val_a, val_b)
    for ti in all_table_ids:
        rows_a=data_a['tables'].get(ti,[])
        rows_b=data_b['tables'].get(ti,[])
        max_r=max(len(rows_a),len(rows_b))
        for ri in range(max_r):
            ra=rows_a[ri] if ri<len(rows_a) else []
            rb=rows_b[ri] if ri<len(rows_b) else []
            max_c=max(len(ra),len(rb))
            for ci in range(max_c):
                ca=ra[ci] if ci<len(ra) else ''
                cb=rb[ci] if ci<len(rb) else ''
                if ca!=cb: table_diffs.append((ti,ri,ci,ca,cb))

    # Gemini 요약
    ai_summary="(API Key 없음 - 요약 생략)"
    if api_key:
        try:
            genai.configure(api_key=api_key)
            model=genai.GenerativeModel(model_name)
            diff_preview=f"EXTRACTION 변경: {len(ext_diffs)}건\nTABLE 셀 변경: {len(table_diffs)}건\n"
            for k,va,vb in ext_diffs[:5]:
                diff_preview+=f"  {k}: {va} -> {vb}\n"
            num_diffs=[(ti,ri,ci,va,vb) for ti,ri,ci,va,vb in table_diffs
                       if cell_num(va) is not None or cell_num(vb) is not None]
            for ti,ri,ci,va,vb in num_diffs[:10]:
                diff_preview+=f"  TABLE[{ti}] R{ri}C{ci}: {va} -> {vb}\n"
            prompt=(
                "한국 DART 감사보고서 두 버전 간의 변경 사항입니다.\n"
                f"{diff_preview}\n"
                "주요 재무적 변동 사항을 3~5줄로 알기 쉽게 요약해 주세요. "
                "숫자 변경의 재무적 의미, 증감 방향, 주요 계정을 중심으로 설명해 주세요."
            )
            resp=model.generate_content(prompt)
            ai_summary=resp.text.strip()
        except Exception as e:
            ai_summary=f"Gemini 오류: {e}"

    # Excel 생성
    wb=openpyxl.Workbook()

    # 시트1: AI 요약
    ws0=wb.active; ws0.title='🤖AI변동요약'; ws0.sheet_view.showGridLines=False
    title_c=ws0.cell(1,1,'🔍 DSD 비교 분석 - AI 변동 요약 리포트')
    title_c.fill=PatternFill('solid',fgColor='1B4F72')
    title_c.font=Font(color='FFFFFF',bold=True,size=13)
    title_c.alignment=Alignment(horizontal='left',vertical='center')
    ws0.merge_cells('A1:D1'); ws0.row_dimensions[1].height=30
    sub_c=ws0.cell(2,1,f'비교 일시: {time.strftime("%Y-%m-%d %H:%M")}  |  EXTRACTION 변경: {len(ext_diffs)}건  |  TABLE 셀 변경: {len(table_diffs)}건')
    sub_c.font=Font(color='2E75B6',size=9,italic=True); ws0.row_dimensions[2].height=15

    # AI 요약 텍스트
    for ri,line in enumerate(ai_summary.split('\n'),4):
        c=ws0.cell(ri,1,line)
        c.font=Font(size=10); c.alignment=Alignment(horizontal='left',vertical='center',wrap_text=True)
        ws0.row_dimensions[ri].height=18
    ws0.column_dimensions['A'].width=100

    # 시트2: EXTRACTION 차이
    if ext_diffs:
        ws_ext=wb.create_sheet('📊EXTRACTION변경')
        for ci,(h,w) in enumerate([('항목코드',20),('변경전',35),('변경후',35)],1):
            c=ws_ext.cell(1,ci,h)
            c.fill=PatternFill('solid',fgColor='1F4E79')
            c.font=Font(color='FFFFFF',bold=True,size=9)
            c.alignment=Alignment(horizontal='center',vertical='center')
            ws_ext.column_dimensions[get_column_letter(ci)].width=w
        for ri,(k,va,vb) in enumerate(ext_diffs,2):
            ws_ext.cell(ri,1,k).font=Font(bold=True,size=9)
            ca=ws_ext.cell(ri,2,va); cb=ws_ext.cell(ri,3,vb)
            ca.fill=PatternFill('solid',fgColor='FCE4EC')
            cb.fill=PatternFill('solid',fgColor='E8F5E9')
            ca.font=Font(size=9); cb.font=Font(size=9)

    # 시트3: TABLE 셀 차이 (빨간 하이라이트)
    ws_diff=wb.create_sheet('🔴셀변경목록')
    for ci,(h,w) in enumerate([('TABLE','8'),('행','6'),('열','6'),('변경전','40'),('변경후','40')],1):
        c=ws_diff.cell(1,ci,h)
        c.fill=PatternFill('solid',fgColor='1F4E79')
        c.font=Font(color='FFFFFF',bold=True,size=9)
        c.alignment=Alignment(horizontal='center',vertical='center')
        ws_diff.column_dimensions[get_column_letter(ci)].width=int(w)*2
    RED=PatternFill('solid',fgColor='FF0000')
    for ri,(ti,rowi,ci,va,vb) in enumerate(table_diffs[:500],2):
        ws_diff.cell(ri,1,ti).font=Font(size=9)
        ws_diff.cell(ri,2,rowi).font=Font(size=9)
        ws_diff.cell(ri,3,ci).font=Font(size=9)
        ca=ws_diff.cell(ri,4,va); cb=ws_diff.cell(ri,5,vb)
        ca.fill=RED; cb.fill=PatternFill('solid',fgColor='FF0000') if not vb else PatternFill('solid',fgColor='FFEB3B')
        ca.font=Font(size=9,color='FFFFFF'); cb.font=Font(size=9,color='000000')
        ws_diff.row_dimensions[ri].height=15
    if len(table_diffs)>500:
        ws_diff.cell(502,1,f'(총 {len(table_diffs)}건 중 500건만 표시)').font=Font(italic=True,size=8,color='888888')

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── 재무/주석 데이터 추출 ──────────────────────────────────────────────────────
def extract_fin_and_notes(xlsx_bytes:bytes)->tuple:
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes),data_only=True)
    fin_data={}; note_data={}
    for sname in wb.sheetnames:
        if sname in ('📋사용안내','_원본XML','📊요약수치'): continue
        ws=wb[sname]
        rows_data=[]
        for row in ws.iter_rows(max_row=200,values_only=True):
            cleaned=[str(v).strip() if v is not None else '' for v in row]
            if any(cleaned): rows_data.append(cleaned)
        if any(sname.startswith(p) for p in FIN_PREFIXES):
            fin_data[sname]=rows_data[:80]
        elif sname.startswith('📝'):
            note_data[sname]=rows_data[:50]
    return fin_data,note_data

# ── DSD -> Excel 변환 (롤오버 옵션 포함) ─────────────────────────────────────
def dsd_to_excel_bytes(dsd_bytes:bytes,ai_mapping:dict=None,do_rollover:bool=False)->bytes:
    with zipfile.ZipFile(io.BytesIO(dsd_bytes)) as zf:
        files={n:zf.read(n) for n in zf.namelist()}
    xml      =files.get('contents.xml',b'').decode('utf-8',errors='replace')
    meta_xml =files.get('meta.xml',b'').decode('utf-8',errors='replace')
    exts,tables=parse_xml(xml)
    wb=openpyxl.Workbook()

    # 사용안내
    ws0=wb.active; ws0.title='📋사용안내'; ws0.sheet_view.showGridLines=False
    guide=[
        ('DART 감사보고서 DSD - Excel 변환 도구 (easydsd v0.7)',True,C['white'],C['navy'],13),
        ('',False,'','',8),
        ('【 작업 순서 】',True,C['navy'],C['lblue'],11),
        ('  1. 노란색 셀을 당해년도 숫자/텍스트로 수정하세요',False,'000000',C['white'],10),
        ('  2. 저장 후 "Excel -> DSD" 탭에서 변환하세요',False,'000000',C['white'],10),
        ('',False,'','',8),
        ('【 색상 범례 】',True,C['navy'],C['lblue'],11),
        ('  노란색 = 수정 가능 (금액, 주주명, 지분율, 텍스트 모두)',False,'000000',C['yellow'],10),
        ('  파란색 = 헤더 (수정 불필요)',False,C['white'],C['navy'],10),
        ('',False,'','',8),
        ('【 주의사항 】',True,C['navy'],C['lblue'],11),
        ('  _원본XML 시트는 절대 수정/삭제 금지!',False,C['orange'],C['white'],10),
    ]
    for ri,(txt,bold,fg,bg,sz) in enumerate(guide,1):
        cc=ws0.cell(ri,1,txt); cc.font=fnt(fg or '000000',bold=bold,size=sz)
        if bg: cc.fill=fill(bg)
        cc.alignment=aln('left',wrap=True); ws0.row_dimensions[ri].height=21
    ws0.column_dimensions['A'].width=65

    # 요약수치
    ws_e=wb.create_sheet('📊요약수치'); ws_e.sheet_view.showGridLines=False
    for ci,(h,w) in enumerate([('ACODE',15),('값 (수정가능)',22),('설명',28)],1):
        cc=ws_e.cell(1,ci,h); cc.fill=fill(C['navy']); cc.font=fnt(C['white'],bold=True)
        cc.alignment=aln('center'); ws_e.column_dimensions[get_column_letter(ci)].width=w
    for ri,(code,val) in enumerate(exts,2):
        ws_e.cell(ri,1,code).font=fnt(bold=True,size=9)
        vc=ws_e.cell(ri,2,val); vc.fill=fill(C['yellow']); vc.alignment=aln('right')
        ws_e.cell(ri,3,EXT_DESC.get(code,'')).font=fnt(size=9,italic=True)
        ws_e.row_dimensions[ri].height=18

    # 그룹핑
    FIN_ORDER=['🏦재무상태표','💹포괄손익계산서','📈자본변동표','💰현금흐름표']
    groups=[]; i=0
    pre_fin=[]
    while i<len(tables) and not tables[i]['fin_label']:
        pre_fin.append(tables[i]); i+=1
    if pre_fin: groups.append(('📝00_서문',pre_fin,False))
    for fin_label in FIN_ORDER:
        if i>=len(tables): break
        fin_tbls=[]
        while i<len(tables) and tables[i]['fin_label']==fin_label:
            fin_tbls.append(tables[i]); i+=1
        if not fin_tbls: continue
        if i<len(tables) and not tables[i]['fin_label']:
            fin_tbls.append(tables[i]); i+=1
        groups.append((fin_label[:31],fin_tbls,False))
    remaining=tables[i:]
    for chunk_n,start in enumerate(range(0,len(remaining),10),1):
        chunk=remaining[start:start+10]
        if ai_mapping:
            ai_names=[ai_mapping.get(t['idx']) for t in chunk if ai_mapping.get(t['idx'])]
            sname=ai_names[0][:31] if ai_names else f'📝{chunk_n:02d}_주석'
        else:
            sname=f'📝{chunk_n:02d}_주석'
        groups.append((sname,chunk,True))

    def write_tables_to_sheet(ws,tbl_list,show_titles=False):
        er=1; max_cols_all=1; table_start_rows={}
        for tbl in tbl_list:
            if not tbl['rows']: continue
            max_cols_all=max(max_cols_all,
                min(max((sum(c['colspan'] for c in row) for row in tbl['rows']),default=1),26))
        for tbl in tbl_list:
            if show_titles and tbl.get('ctx_title'):
                div=ws.cell(er,1,tbl['ctx_title'])
                div.fill=PatternFill('solid',fgColor='D9D9D9')
                div.font=Font(bold=True,size=9,color='333333')
                div.alignment=Alignment(horizontal='left',vertical='center')
                if max_cols_all>1:
                    try: ws.merge_cells(start_row=er,start_column=1,end_row=er,end_column=max_cols_all)
                    except: pass
                ws.row_dimensions[er].height=16; er+=1
            table_start_rows[tbl['idx']]=er
            for row in tbl['rows']:
                col=1
                for cell in row:
                    if col>26: break
                    wc=ws.cell(er,col,cell['value']); v,tag=cell['value'],cell['tag']
                    if tag in ('TH','TE'):
                        wc.fill=fill(C['navy']); wc.font=fnt(C['white'],bold=True,size=9)
                        wc.alignment=aln('center',wrap=True)
                    else:
                        wc.fill=fill(C['yellow']); wc.font=fnt(size=9)
                        wc.alignment=aln('right' if is_num_or_decimal(v) else 'left',wrap=True)
                    if '\n' in str(v):
                        ws.row_dimensions[er].height=max(18,18*(str(v).count('\n')+1))
                    span=min(cell['colspan'],26-col+1)
                    if span>1:
                        try: ws.merge_cells(start_row=er,start_column=col,end_row=er,end_column=col+span-1)
                        except: pass
                    col+=cell['colspan']
                if not ws.row_dimensions[er].height or ws.row_dimensions[er].height<18:
                    ws.row_dimensions[er].height=18
                er+=1
        ws.column_dimensions['A'].width=28
        for ci in range(2,max_cols_all+1): ws.column_dimensions[get_column_letter(ci)].width=18
        return table_start_rows

    sheet_map=[]; used=set()
    for gitem in groups:
        sraw=gitem[0]; tbl_list=gitem[1]; show_t=gitem[2] if len(gitem)>2 else False
        sname=sraw[:31]
        if sname in used: sname=(sraw[:28]+f'_{len(used)}')[:31]
        used.add(sname)
        ws=wb.create_sheet(sname); ws.sheet_view.showGridLines=False
        tsr=write_tables_to_sheet(ws,tbl_list,show_t)
        for tbl in tbl_list:
            sheet_map.append((sname,tbl['idx'],tsr.get(tbl['idx'],-1)))

    # _원본XML
    ws_r=wb.create_sheet('_원본XML'); ws_r.sheet_view.showGridLines=False
    ws_r.cell(1,1,'이 시트는 DSD 복원에 필수입니다. 절대 수정/삭제 금지!').font=fnt(C['orange'],bold=True,size=9)
    ws_r.cell(2,1,'meta_xml'); ws_r.cell(2,2,meta_xml or '')
    for hi,(h,w) in enumerate([('sheet_name',35),('table_idx',12),('fin_label',22),('ctx_title',40),('excel_start_row',14)],1):
        ws_r.cell(4,hi,h); ws_r.column_dimensions[get_column_letter(hi)].width=w
    for ri,(sname,t_idx,excel_row) in enumerate(sheet_map,5):
        t=tables[t_idx]; ws_r.cell(ri,1,sname); ws_r.cell(ri,2,t_idx)
        ws_r.cell(ri,3,t['fin_label']); ws_r.cell(ri,4,t['ctx_title']); ws_r.cell(ri,5,excel_row)

    # 롤오버 적용
    if do_rollover:
        apply_rollover(wb)

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── Excel -> DSD 변환 ──────────────────────────────────────────────────────────
def is_note_ref(val):
    parts=val.strip().split(',')
    return (len(parts)>=2 and all(p.strip().isdigit() and 1<=len(p.strip())<=2 for p in parts))

def normalize_num(val):
    v=str(val).strip()
    if not v or v in ('-',''): return v
    if '\n' in v: return '&amp;cr;'.join(normalize_num(l) for l in v.split('\n'))
    if is_note_ref(v): return v
    neg=v.startswith('-') or (v.startswith('(') and v.endswith(')'))
    cl=v.replace(',','').replace('(','').replace(')','').replace('-','').replace(' ','')
    if cl.isdigit() and len(cl)>=3:
        fmt=f"{int(cl):,}"; v=f"({fmt})" if neg else fmt
    return v.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def excel_to_dsd_bytes(orig_dsd_bytes:bytes,xlsx_bytes:bytes)->bytes:
    with zipfile.ZipFile(io.BytesIO(orig_dsd_bytes)) as zf:
        orig_files={n:zf.read(n) for n in zf.namelist()}
    contents_xml=orig_files['contents.xml'].decode('utf-8',errors='replace')
    wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes),data_only=True)

    mapping={}
    if '_원본XML' in wb.sheetnames:
        ws_r=wb['_원본XML']
        for row in ws_r.iter_rows(min_row=5,values_only=True):
            if not row or row[0] is None or row[1] is None: continue
            sname=str(row[0]).strip(); t_idx=int(row[1])
            excel_row=int(row[4]) if len(row)>4 and row[4] is not None else -1
            if sname: mapping.setdefault(sname,[]).append((t_idx,excel_row))

    exts={}; t_changes={}
    for sname in wb.sheetnames:
        if sname in ('📋사용안내','_원본XML','_meta'): continue
        ws=wb[sname]
        if sname=='📊요약수치':
            for row in ws.iter_rows(min_row=2):
                if len(row)<2: continue
                cc,vc=row[0],row[1]
                if cc.value and vc.value is not None and is_edit(vc):
                    exts[str(cc.value).strip()]=str(vc.value).strip()
        else:
            changes=[]
            for ri,row in enumerate(ws.iter_rows(min_row=2)):
                for ci,cell in enumerate(row):
                    if is_edit(cell) and cell.value is not None:
                        changes.append((ri,ci,str(cell.value)))
            if changes: t_changes[sname]=changes

    for ext_code,val in exts.items():
        contents_xml=re.sub(
            rf'(<EXTRACTION[^>]*ACODE="{re.escape(ext_code)}"[^>]*>)[^<]+(</EXTRACTION>)',
            rf'\g<1>{val}\g<2>',contents_xml)

    table_positions=[(m.start(),m.end())
                     for m in re.finditer(r'<TABLE[^>]*>.*?</TABLE>',contents_xml,re.DOTALL)]
    patches=[]
    for sname,changes in t_changes.items():
        t_info=mapping.get(sname)
        if not t_info: continue
        all_ch={(r,c):v for r,c,v in changes}
        for k,(t_idx,esr) in enumerate(t_info):
            if t_idx>=len(table_positions): continue
            if esr>=0:
                # er은 1-based, ri는 0-based(min_row=2 기준) → s = esr - 2
                s  =esr-2
                nxt_esr=t_info[k+1][1] if k+1<len(t_info) and t_info[k+1][1]>=0 else None
                nxt=(nxt_esr-2) if nxt_esr else 99999
                local_map={(r-s,c):v for (r,c),v in all_ch.items() if s<=r<nxt}
            else:
                off=0
                for j in range(k):
                    ti2=t_info[j][0]
                    if ti2<len(table_positions):
                        sn=contents_xml[table_positions[ti2][0]:table_positions[ti2][1]]
                        off+=len(re.findall(r'<TR[^>]*>',sn))
                sn=contents_xml[table_positions[t_idx][0]:table_positions[t_idx][1]]
                tc=len(re.findall(r'<TR[^>]*>',sn))
                local_map={(r-off,c):v for (r,c),v in all_ch.items() if off<=r<off+tc}
            if not local_map: continue
            ts,te=table_positions[t_idx]; tt=contents_xml[ts:te]
            rebuilt=[]; last=0; td_row=0
            for tr_m in re.finditer(r'(<TR[^>]*>)(.*?)(</TR>)',tt,re.DOTALL):
                rebuilt.append(tt[last:tr_m.start()])
                trb=tr_m.group(2); nb=[]; td_last=0; td_col=0
                for td_m in re.finditer(r'(<(?:TD|TH|TU|TE)[^>]*>)(.*?)(</(?:TD|TH|TU|TE)>)',trb,re.DOTALL):
                    nb.append(trb[td_last:td_m.start()])
                    key=(td_row,td_col)
                    if key in local_map:
                        nb.append(td_m.group(1)+normalize_num(local_map[key])+td_m.group(3))
                    else:
                        nb.append(td_m.group(0))
                    td_last=td_m.end(); td_col+=1
                nb.append(trb[td_last:])
                rebuilt.append(tr_m.group(1)+''.join(nb)+tr_m.group(3))
                last=tr_m.end(); td_row+=1
            rebuilt.append(tt[last:])
            patches.append((ts,te,''.join(rebuilt)))

    result=contents_xml
    for ts,te,nt in sorted(patches,key=lambda x:-x[0]):
        result=result[:ts]+nt+result[te:]

    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as zf:
        for name,data in orig_files.items():
            if os.path.splitext(name)[1].lower() in ('.jpg','.jpeg','.png','.gif','.bmp'): continue
            zf.writestr(name,result.encode('utf-8') if name=='contents.xml' else data)
    return buf.getvalue()

# ── Flask 앱 ──────────────────────────────────────────────────────────────────
app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=100*1024*1024

_last_ping=time.time()
def _watchdog():
    time.sleep(30 if IS_FROZEN else 12)
    while True:
        time.sleep(2)
        if time.time()-_last_ping>8: os._exit(0)
threading.Thread(target=_watchdog,daemon=True).start()


# ── HTML (r-string: 내부 작은따옴표 이스케이프 불필요) ────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>easydsd v0.7</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Malgun Gothic','맑은 고딕',sans-serif;background:#f0f4f8;color:#1a1a2e;min-height:100vh}
.header{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:white;padding:14px 24px;box-shadow:0 4px 20px rgba(31,78,121,.3)}
.hd-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.hd-top h1{font-size:17px;font-weight:700}
.hd-top p{font-size:11px;opacity:.75;margin-top:2px}
.hd-right{display:flex;align-items:center;gap:8px;flex-shrink:0}
.hd-badge{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:20px;padding:3px 12px;font-size:11px;font-weight:600}
.kill-btn{background:#c0392b;color:white;border:none;border-radius:7px;padding:6px 12px;font-size:11px;font-weight:700;cursor:pointer;transition:background .15s}
.kill-btn:hover{background:#e74c3c}
.api-bar{display:flex;align-items:center;gap:8px;background:rgba(0,0,0,.18);border-radius:8px;padding:8px 12px;flex-wrap:wrap}
.api-bar label{font-size:11px;font-weight:700;white-space:nowrap;opacity:.9}
.api-input{flex:1;min-width:180px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);border-radius:6px;color:white;padding:5px 10px;font-size:12px;font-family:monospace;outline:none}
.api-input::placeholder{opacity:.5}
.api-input:focus{background:rgba(255,255,255,.2);border-color:rgba(255,255,255,.5)}
.api-sel{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);border-radius:6px;color:white;padding:4px 7px;font-size:11px;cursor:pointer;outline:none;flex-shrink:0}
.api-note{font-size:10px;white-space:nowrap;background:rgba(255,193,7,.3);border:1px solid rgba(255,193,7,.5);border-radius:10px;padding:2px 7px;color:#fff8e1;font-weight:600}
.api-st{font-size:10px;white-space:nowrap;opacity:.8}
.api-clear-btn{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:white;border-radius:6px;padding:3px 8px;font-size:10px;cursor:pointer;white-space:nowrap}
.api-sub{margin-top:5px;font-size:10px;opacity:.75}
.api-sub a{color:#a5d6a7;font-weight:600;text-decoration:none}
.api-sub a:hover{text-decoration:underline}
.container{max-width:860px;margin:20px auto;padding:0 16px 60px}
.tabs{display:flex;gap:2px;flex-wrap:wrap}
.tab{padding:8px 14px;border-radius:10px 10px 0 0;background:#cdd8e4;color:#4a6078;font-size:11px;font-weight:600;cursor:pointer;border:none;border-bottom:3px solid transparent;transition:all .2s;white-space:nowrap}
.tab.active{background:white;color:#1F4E79;border-bottom:3px solid #1F4E79}
.tab:hover:not(.active){background:#bcccd8}
.tab.ai-tab{background:#2d1b4e;color:#b39ddb}
.tab.ai-tab.active{background:white;color:#6200ea;border-bottom:3px solid #6200ea}
.tab.ai-tab:hover:not(.active){background:#3d2b5e;color:#ce93d8}
.tab.diff-tab{background:#1a3a1a;color:#81c784}
.tab.diff-tab.active{background:white;color:#1b5e20;border-bottom:3px solid #2e7d32}
.tab.diff-tab:hover:not(.active){background:#2a4a2a;color:#a5d6a7}
.tab.dev-tab{background:#2a2a2a;color:#999}
.tab.dev-tab.active{background:white;color:#333;border-bottom:3px solid #666}
.tab.dev-tab:hover:not(.active){background:#3a3a3a;color:#bbb}
.card{background:white;border-radius:0 12px 12px 12px;box-shadow:0 4px 24px rgba(0,0,0,.08);padding:24px}
.tab-content{display:none}.tab-content.active{display:block}
.step{display:flex;gap:10px;align-items:flex-start;padding:12px;margin-bottom:10px;background:#f7f9fc;border-radius:10px;border-left:4px solid #2E75B6}
.step-num{min-width:24px;height:24px;border-radius:50%;background:#1F4E79;color:white;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:11px;flex-shrink:0}
.step-title{font-weight:700;font-size:13px;color:#1F4E79;margin-bottom:3px}
.step-desc{font-size:11px;color:#556;line-height:1.6}
.chk-row{display:flex;align-items:center;gap:7px;margin-top:9px;padding:8px 12px;border-radius:8px}
.chk-row.purple{background:#f3e5f5;border:1px solid #ce93d8}
.chk-row.teal{background:#e0f2f1;border:1px solid #80cbc4}
.chk-row input[type=checkbox]{width:14px;height:14px;cursor:pointer}
.chk-row.purple input{accent-color:#6200ea}
.chk-row.teal input{accent-color:#00695c}
.chk-row label{font-size:12px;font-weight:600;cursor:pointer}
.chk-row.purple label{color:#4a148c}
.chk-row.teal label{color:#004d40}
.chk-note{font-size:10px;margin-left:4px;opacity:.75}
.drop-zone{border:2px dashed #a0b8d0;border-radius:10px;padding:20px;text-align:center;cursor:pointer;transition:all .2s;background:#f7fbff;margin-top:8px}
.drop-zone:hover,.drop-zone.drag-over{border-color:#1F4E79;background:#e8f0f8}
.drop-zone .dz-icon{font-size:24px;margin-bottom:3px}
.drop-zone .dz-lbl{font-size:12px;color:#4a6078}
.drop-zone .dz-sub{font-size:10px;color:#89a;margin-top:2px}
.file-badge{margin-top:5px;font-size:11px;color:#1F4E79;font-weight:600;display:none;background:#e8f0f8;padding:4px 9px;border-radius:6px}
.btn{width:100%;padding:10px;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;transition:all .2s;margin-top:11px}
.btn-blue{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:white}
.btn-blue:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 5px 18px rgba(31,78,121,.35)}
.btn-blue:disabled{background:#a0b8c8;cursor:not-allowed}
.btn-green{background:linear-gradient(135deg,#1a6b3a,#22a55a);color:white}
.btn-green:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 5px 18px rgba(26,107,58,.35)}
.btn-green:disabled{background:#8fc0a0;cursor:not-allowed}
.btn-ai{background:linear-gradient(135deg,#4a148c,#7b1fa2);color:white}
.btn-ai:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 5px 18px rgba(74,20,140,.4)}
.btn-ai:disabled{background:#b39ddb;cursor:not-allowed}
.btn-diff{background:linear-gradient(135deg,#1b5e20,#388e3c);color:white}
.btn-diff:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 5px 18px rgba(27,94,32,.4)}
.btn-diff:disabled{background:#81c784;cursor:not-allowed}
.prog-wrap{margin-top:11px;display:none}
.prog-bar{height:5px;background:#e0e8f0;border-radius:4px;overflow:hidden}
.prog-fill{height:100%;width:0%;border-radius:4px;transition:width .35s ease}
.pf-blue{background:linear-gradient(90deg,#1F4E79,#2E75B6)}
.pf-ai{background:linear-gradient(90deg,#4a148c,#7b1fa2)}
.pf-diff{background:linear-gradient(90deg,#1b5e20,#388e3c)}
.prog-text{font-size:11px;color:#4a6078;margin-top:3px;text-align:center}
.result{margin-top:11px;padding:11px 14px;border-radius:10px;display:none;align-items:center;gap:9px}
.result.ok{background:#e8f5ec;border:1px solid #6dbf8a}
.result.err{background:#fdecea;border:1px solid #e88}
.result.ai-ok{background:#f3e5f5;border:1px solid #ce93d8}
.result.diff-ok{background:#e8f5e9;border:1px solid #66bb6a}
.r-icon{font-size:19px}
.r-body{flex:1}
.r-title{font-weight:700;font-size:13px}
.r-sub{font-size:11px;margin-top:2px;color:#556}
.dl-btn{padding:6px 12px;color:white;border:none;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap;text-decoration:none;display:inline-block;transition:background .15s}
.dl-btn.green{background:#1a6b3a}.dl-btn.green:hover{background:#145530}
.dl-btn.blue{background:#1F4E79}.dl-btn.blue:hover{background:#163a5e}
.dl-btn.purple{background:#4a148c}.dl-btn.purple:hover{background:#6a1b9a}
.dl-btn.forest{background:#1b5e20}.dl-btn.forest:hover{background:#145218}
.legend{display:flex;gap:9px;flex-wrap:wrap;margin-top:6px}
.leg-item{display:flex;align-items:center;gap:5px;font-size:11px;color:#556}
.leg-dot{width:12px;height:12px;border-radius:3px;flex-shrink:0}
.ai-notice{background:#f3e5f5;border:1px solid #ce93d8;border-radius:8px;padding:9px 13px;margin-bottom:11px;font-size:11px;color:#4a148c;line-height:1.6}
.ai-hdr{display:flex;align-items:center;gap:9px;padding:12px 14px;border-radius:10px;margin-bottom:12px;color:white}
.ai-hdr.purple{background:linear-gradient(135deg,#4a148c,#7b1fa2)}
.ai-hdr.green{background:linear-gradient(135deg,#1b5e20,#2e7d32)}
.ai-hdr .ico{font-size:24px}
.ai-hdr h3{font-size:14px;font-weight:700}
.ai-hdr p{font-size:10px;opacity:.8;margin-top:2px}
.pycheck-box{margin-top:10px;padding:11px 14px;border-radius:9px;background:#fff8e1;border:1px solid #f9a825;display:none}
.pycheck-box h4{font-size:11px;font-weight:700;color:#f57f17;margin-bottom:6px}
.pycheck-item{font-size:11px;padding:3px 0;border-bottom:1px solid #fff3cd;line-height:1.5}
.pycheck-item:last-child{border-bottom:none}
.pycheck-item.err{color:#c62828}
.pycheck-item.warn{color:#e65100}
.pycheck-item.info{color:#1b5e20}
.vr-box{margin-top:11px;display:none;background:#fafafa;border:1px solid #e0e0e0;border-radius:9px;padding:13px;max-height:340px;overflow-y:auto}
.vr-box h4{font-size:11px;font-weight:700;color:#4a148c;margin-bottom:6px}
.vr-box pre{font-size:11px;color:#333;white-space:pre-wrap;line-height:1.7;font-family:'Malgun Gothic',sans-serif}
.diff-notice{background:#e8f5e9;border:1px solid #80cbc4;border-radius:8px;padding:9px 13px;margin-bottom:11px;font-size:11px;color:#1b5e20;line-height:1.6}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.dev-profile{display:flex;align-items:center;gap:14px;padding:14px;background:linear-gradient(135deg,#1a1a2e,#16213e);border-radius:12px;margin-bottom:12px}
.dev-av{width:54px;height:54px;border-radius:50%;background:linear-gradient(135deg,#1F4E79,#2E75B6);display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;border:3px solid rgba(255,255,255,.2)}
.dev-info h2{color:white;font-size:14px;font-weight:700;margin-bottom:2px}
.dev-sub{color:rgba(255,255,255,.6);font-size:10px;margin-bottom:5px}
.dev-badges{display:flex;gap:5px;flex-wrap:wrap}
.badge{border-radius:20px;padding:2px 8px;font-size:10px;font-weight:600}
.bg0{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.8)}
.bg-gold{background:linear-gradient(135deg,#b8860b,#daa520);color:white}
.bg-tech{background:rgba(46,117,182,.5);border:1px solid rgba(46,117,182,.8);color:white}
.bg-ai{background:linear-gradient(135deg,#4a148c,#7b1fa2);color:white}
.ig{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:11px}
.ib{background:#f7f9fc;border-radius:9px;padding:10px 12px;border-left:3px solid #2E75B6}
.ib .lbl{font-size:10px;color:#89a;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px}
.ib .val{font-size:12px;font-weight:700;color:#1F4E79}
.ib .val a{color:#1F4E79;text-decoration:none}.ib .val a:hover{text-decoration:underline}
.cbox{background:#fffbf0;border:1px solid #e8d060;border-radius:11px;padding:13px;text-align:center;margin-bottom:11px}
.ct{font-size:12px;font-weight:700;color:#7a5500;margin-bottom:6px}
.cb{font-size:12px;color:#444;line-height:1.9}
.cn{font-size:14px;font-weight:800;color:#1a1a2e;margin:4px 0 1px}
.cs{font-size:10px;color:#888}
.cc{display:inline-block;background:linear-gradient(135deg,#7c4dff,#2196f3);color:white;border-radius:20px;padding:3px 10px;font-size:11px;font-weight:700;margin:0 3px;vertical-align:middle}
.fs h3{font-size:12px;font-weight:700;color:#1F4E79;margin-bottom:6px}
.fi{display:flex;align-items:flex-start;gap:7px;padding:5px 0;border-bottom:1px solid #f0f4f8;font-size:11px;color:#446;line-height:1.5}
.fi:last-child{border-bottom:none}
.fic{font-size:12px;flex-shrink:0;margin-top:1px}
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:9999;align-items:center;justify-content:center}
.modal-overlay.show{display:flex}
.modal{background:white;border-radius:13px;padding:22px 26px;max-width:300px;width:90%;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal h3{font-size:14px;font-weight:700;margin-bottom:5px}
.modal p{font-size:12px;color:#556;margin-bottom:14px;line-height:1.6}
.mbtns{display:flex;gap:7px;justify-content:center}
.mbtns button{padding:8px 18px;border:none;border-radius:7px;font-size:12px;font-weight:700;cursor:pointer}
.mc{background:#e8eef4;color:#4a6078}.mc:hover{background:#d0dce8}
.mx{background:#c0392b;color:white}.mx:hover{background:#e74c3c}
</style>
</head>
<body>
<div class="header">
  <div class="hd-top">
    <div>
      <h1>&#128202; DART 감사보고서 변환 도구</h1>
      <p>DSD &harr; Excel &nbsp;&#xB7;&nbsp; 롤오버 &nbsp;&#xB7;&nbsp; AI 검증 &nbsp;&#xB7;&nbsp; DSD 비교 &nbsp;&#xB7;&nbsp; easydsd v0.7</p>
    </div>
    <div class="hd-right">
      <div class="hd-badge">v0.7</div>
      <button class="kill-btn" onclick="showKill()">&#x23FC; 종료</button>
    </div>
  </div>
  <div class="api-bar">
    <label>&#128273; Gemini API Key</label>
    <input class="api-input" id="apiKey" type="password"
      placeholder="AIza... (없어도 DSD 변환은 완벽 작동 - 선택사항)"
      oninput="saveKey(this.value)" />
    <select id="modelSel" class="api-sel" onchange="saveModel(this.value)">
      <option value="gemini-3-flash-preview">3 Flash &#x2605; 최신</option>
      <option value="gemini-2.5-flash">2.5 Flash</option>
      <option value="gemini-2.0-flash">2.0 Flash (6월종료)</option>
    </select>
    <span class="api-note">&#128204; 선택사항</span>
    <span class="api-st" id="apiSt">&#x26AA; 미입력</span>
    <button class="api-clear-btn" onclick="clearKey()">&#x2715; 삭제</button>
  </div>
  <div class="api-sub">
    &#128204; API Key 없이도 DSD&harr;Excel 변환 및 롤오버는 완벽 작동합니다. &nbsp;|&nbsp;
    <a href="https://aistudio.google.com/app/apikey" target="_blank">&#128073; 1분 만에 무료 API 키 발급받는 방법 &#x2197;</a>
  </div>
</div>

<div class="modal-overlay" id="killModal">
  <div class="modal">
    <h3>&#x26A0;&#xFE0F; 종료할까요?</h3>
    <p>서버 프로세스가 완전히 종료됩니다.<br>이 탭도 닫아주세요.</p>
    <div class="mbtns">
      <button class="mc" onclick="hideKill()">취소</button>
      <button class="mx" onclick="doKill()">종료</button>
    </div>
  </div>
</div>

<div class="container">
  <div class="tabs">
    <button class="tab active" onclick="sw(0)">&#9312; DSD &#8594; Excel</button>
    <button class="tab" onclick="sw(1)">&#9313; Excel &#8594; DSD</button>
    <button class="tab ai-tab" onclick="sw(2)">&#129302; AI 재무제표 검증</button>
    <button class="tab diff-tab" onclick="sw(3)">&#128269; DSD 비교 분석</button>
    <button class="tab dev-tab" onclick="sw(4)">개발자 정보</button>
  </div>
  <div class="card">

    <!-- ① DSD -> Excel -->
    <div class="tab-content active" id="tab0">
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <div class="step-title">전년도 DSD 파일을 업로드하세요</div>
          <div class="step-desc">DART에서 제출한 .dsd 파일을 드래그하거나 클릭해 선택하세요.<br>변환된 Excel의 <b>노란색 셀</b>을 당해년도 숫자로 수정하시면 됩니다.</div>
          <div class="drop-zone" id="dz1" onclick="document.getElementById('f1').click()"
               ondragover="dov(event,'dz1')" ondragleave="dlv('dz1')" ondrop="ddrop(event,'f1','dz1')">
            <div class="dz-icon">&#128194;</div>
            <div class="dz-lbl">클릭하거나 파일을 여기에 끌어다 놓으세요</div>
            <div class="dz-sub">.dsd 파일</div>
          </div>
          <input type="file" id="f1" accept=".dsd" style="display:none" onchange="sf('f1','fb1','dz1')">
          <div class="file-badge" id="fb1"></div>
          <div class="chk-row teal">
            <input type="checkbox" id="chkRollover">
            <label for="chkRollover">&#128260; 작년 당기 실적을 올해 전기 칸으로 자동 이월하기 (Rollover)</label>
            <span class="chk-note">(재무4표 당기열 &#8594; 전기열, 당기열 비우기)</span>
          </div>
          <div class="chk-row purple">
            <input type="checkbox" id="chkAI">
            <label for="chkAI">&#129302; AI를 이용해 재무제표 및 주석 스마트 분류하기</label>
            <span class="chk-note">(Gemini API Key 필요)</span>
          </div>
        </div>
      </div>
      <button class="btn btn-blue" id="btn1" onclick="run1()" disabled>&#128229; Excel 파일로 변환하기</button>
      <div class="prog-wrap" id="pw1"><div class="prog-bar"><div class="prog-fill pf-blue" id="pf1"></div></div><div class="prog-text" id="pt1">변환 중...</div></div>
      <div class="result ok" id="ok1">
        <div class="r-icon">&#9989;</div>
        <div class="r-body">
          <div class="r-title" id="ok1t"></div><div class="r-sub" id="ok1s"></div>
          <div class="legend" style="margin-top:6px">
            <div class="leg-item"><div class="leg-dot" style="background:#FFF2CC;border:1px solid #ccc"></div>노란색=수정가능</div>
            <div class="leg-item"><div class="leg-dot" style="background:#1F4E79"></div>파란색=헤더</div>
            <div class="leg-item"><div class="leg-dot" style="background:#D9D9D9;border:1px solid #bbb"></div>회색=구분선</div>
          </div>
        </div>
        <a class="dl-btn green" id="dl1" href="#">&#11015; 다운로드</a>
      </div>
      <div class="result err" id="er1"><div class="r-icon">&#10060;</div><div class="r-body"><div class="r-title">변환 실패</div><div class="r-sub" id="er1m"></div></div></div>
    </div>

    <!-- ② Excel -> DSD -->
    <div class="tab-content" id="tab1">
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <div class="step-title">원본 DSD 파일 업로드</div>
          <div class="step-desc">&#9312; 탭에서 사용했던 원본 .dsd 파일을 올려주세요.</div>
          <div class="drop-zone" id="dz2" onclick="document.getElementById('f2').click()"
               ondragover="dov(event,'dz2')" ondragleave="dlv('dz2')" ondrop="ddrop(event,'f2','dz2')">
            <div class="dz-icon">&#128194;</div><div class="dz-lbl">원본 DSD 파일</div><div class="dz-sub">.dsd</div>
          </div>
          <input type="file" id="f2" accept=".dsd" style="display:none" onchange="sf('f2','fb2','dz2')">
          <div class="file-badge" id="fb2"></div>
        </div>
      </div>
      <div class="step">
        <div class="step-num">2</div>
        <div class="step-body">
          <div class="step-title">수정한 Excel 파일 업로드</div>
          <div class="step-desc">노란색 셀을 수정한 .xlsx 파일을 올려주세요.</div>
          <div class="drop-zone" id="dz3" onclick="document.getElementById('f3').click()"
               ondragover="dov(event,'dz3')" ondragleave="dlv('dz3')" ondrop="ddrop(event,'f3','dz3')">
            <div class="dz-icon">&#128202;</div><div class="dz-lbl">수정된 Excel 파일</div><div class="dz-sub">.xlsx</div>
          </div>
          <input type="file" id="f3" accept=".xlsx" style="display:none" onchange="sf('f3','fb3','dz3')">
          <div class="file-badge" id="fb3"></div>
        </div>
      </div>
      <button class="btn btn-green" id="btn2" onclick="run2()" disabled>&#128228; DSD 파일로 변환하기</button>
      <div class="prog-wrap" id="pw2"><div class="prog-bar"><div class="prog-fill pf-blue" id="pf2"></div></div><div class="prog-text" id="pt2">변환 중...</div></div>
      <div class="result ok" id="ok2">
        <div class="r-icon">&#9989;</div>
        <div class="r-body"><div class="r-title" id="ok2t"></div><div class="r-sub" id="ok2s"></div></div>
        <a class="dl-btn blue" id="dl2" href="#">&#11015; DSD 다운로드</a>
      </div>
      <div class="result err" id="er2"><div class="r-icon">&#10060;</div><div class="r-body"><div class="r-title">변환 실패</div><div class="r-sub" id="er2m"></div></div></div>
    </div>

    <!-- ③ AI 검증 (강화) -->
    <div class="tab-content" id="tab2">
      <div class="ai-notice">
        &#x2139;&#xFE0F; <b>이 기능은 선택사항입니다.</b> API Key 없이도 DSD 변환 기능은 완벽히 작동합니다.<br>
        &#128270; <b>Python 수학 검사</b>(대차평균, 단위 이상, 주석번호 매핑)는 API Key 없이도 실행됩니다.
      </div>
      <div class="ai-hdr purple">
        <div class="ico">&#129302;</div>
        <div><h3>AI 재무제표 교차 검증 (강화판)</h3><p>Python 수학 검사 + Gemini AI 교차 검증을 단계적으로 수행합니다</p></div>
      </div>
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <div class="step-title">수정한 Excel 파일 업로드</div>
          <div class="step-desc">easydsd로 변환 후 수정한 .xlsx 파일을 올려주세요.<br>
            <b style="color:#4a148c">&#x26A0; Gemini AI 검증은 상단 API Key가 필요합니다.</b><br>
            API Key 없이도 파이썬 수학 검사(대차평균, 자릿수, 주석 매핑)는 실행됩니다.</div>
          <div class="drop-zone" id="dz4" onclick="document.getElementById('f4').click()"
               ondragover="dov(event,'dz4')" ondragleave="dlv('dz4')" ondrop="ddrop(event,'f4','dz4')">
            <div class="dz-icon">&#128202;</div><div class="dz-lbl">수정된 Excel 파일 (.xlsx)</div><div class="dz-sub">easydsd로 변환한 Excel</div>
          </div>
          <input type="file" id="f4" accept=".xlsx" style="display:none" onchange="sf('f4','fb4','dz4')">
          <div class="file-badge" id="fb4"></div>
        </div>
      </div>
      <button class="btn btn-ai" id="btn3" onclick="run3()" disabled>&#129302; AI 교차 검증 실행하기</button>
      <div class="prog-wrap" id="pw3"><div class="prog-bar"><div class="prog-fill pf-ai" id="pf3"></div></div><div class="prog-text" id="pt3">분석 중...</div></div>
      <!-- Python 검사 결과 -->
      <div class="pycheck-box" id="pycheckBox">
        <h4>&#128270; Python 수학 검사 결과</h4>
        <div id="pycheckItems"></div>
      </div>
      <div class="result ai-ok" id="ok3">
        <div class="r-icon">&#129302;</div>
        <div class="r-body"><div class="r-title" id="ok3t"></div><div class="r-sub" id="ok3s"></div></div>
        <a class="dl-btn purple" id="dl3" href="#">&#11015; 검증결과 다운로드</a>
      </div>
      <div class="result err" id="er3"><div class="r-icon">&#10060;</div><div class="r-body"><div class="r-title">검증 실패</div><div class="r-sub" id="er3m"></div></div></div>
      <div class="vr-box" id="vrBox"><h4>&#129302; AI 검증 결과 미리보기</h4><pre id="vrText"></pre></div>
    </div>

    <!-- ④ DSD 비교 분석 -->
    <div class="tab-content" id="tab3">
      <div class="diff-notice">
        &#128269; <b>두 DSD 파일을 비교</b>하여 변경된 셀을 자동으로 찾아냅니다.<br>
        Gemini API Key가 있으면 변동 사항을 AI가 3~5줄로 요약해줍니다. (없어도 Diff 리포트 생성 가능)
      </div>
      <div class="ai-hdr green">
        <div class="ico">&#128269;</div>
        <div><h3>DSD 비교 분석 (Diff)</h3><p>수정 전/후 DSD 파일의 재무적 변동을 자동 감지합니다</p></div>
      </div>
      <div class="two-col">
        <div class="step">
          <div class="step-num">1</div>
          <div class="step-body">
            <div class="step-title">수정 전 DSD</div>
            <div class="step-desc">원본(구버전) DSD 파일</div>
            <div class="drop-zone" id="dz5" onclick="document.getElementById('f5').click()"
                 ondragover="dov(event,'dz5')" ondragleave="dlv('dz5')" ondrop="ddrop(event,'f5','dz5')">
              <div class="dz-icon">&#128194;</div><div class="dz-lbl">수정 전 DSD</div><div class="dz-sub">.dsd</div>
            </div>
            <input type="file" id="f5" accept=".dsd" style="display:none" onchange="sf('f5','fb5','dz5')">
            <div class="file-badge" id="fb5"></div>
          </div>
        </div>
        <div class="step">
          <div class="step-num">2</div>
          <div class="step-body">
            <div class="step-title">수정 후 DSD</div>
            <div class="step-desc">수정된(신버전) DSD 파일</div>
            <div class="drop-zone" id="dz6" onclick="document.getElementById('f6').click()"
                 ondragover="dov(event,'dz6')" ondragleave="dlv('dz6')" ondrop="ddrop(event,'f6','dz6')">
              <div class="dz-icon">&#128194;</div><div class="dz-lbl">수정 후 DSD</div><div class="dz-sub">.dsd</div>
            </div>
            <input type="file" id="f6" accept=".dsd" style="display:none" onchange="sf('f6','fb6','dz6')">
            <div class="file-badge" id="fb6"></div>
          </div>
        </div>
      </div>
      <button class="btn btn-diff" id="btn4" onclick="run4()" disabled>&#128269; DSD 비교 분석 실행</button>
      <div class="prog-wrap" id="pw4"><div class="prog-bar"><div class="prog-fill pf-diff" id="pf4"></div></div><div class="prog-text" id="pt4">비교 분석 중...</div></div>
      <div class="result diff-ok" id="ok4">
        <div class="r-icon">&#128269;</div>
        <div class="r-body"><div class="r-title" id="ok4t"></div><div class="r-sub" id="ok4s"></div></div>
        <a class="dl-btn forest" id="dl4" href="#">&#11015; Diff 리포트 다운로드</a>
      </div>
      <div class="result err" id="er4"><div class="r-icon">&#10060;</div><div class="r-body"><div class="r-title">비교 실패</div><div class="r-sub" id="er4m"></div></div></div>
    </div>

    <!-- ⑤ 개발자 정보 -->
    <div class="tab-content" id="tab4">
      <div class="dev-profile">
        <div class="dev-av">&#127970;</div>
        <div class="dev-info">
          <h2>Easydsd 0.7v</h2>
          <div class="dev-sub">DART 감사보고서 DSD 변환 + AI 검증 + DSD 비교 + 롤오버</div>
          <div class="dev-badges">
            <span class="badge bg0">v0.7</span>
            <span class="badge bg-gold">&#129302; AI-Powered</span>
            <span class="badge bg-tech">Python+Flask</span>
            <span class="badge bg-ai">Gemini 3 Flash</span>
          </div>
        </div>
      </div>
      <div class="ig">
        <div class="ib"><div class="lbl">개발자</div><div class="val"><a href="mailto:eeffco11@naver.com">eeffco11@naver.com</a></div></div>
        <div class="ib"><div class="lbl">버전</div><div class="val">Easydsd 0.7v</div></div>
        <div class="ib"><div class="lbl">지원 파일</div><div class="val">.dsd / .xlsx</div></div>
        <div class="ib"><div class="lbl">AI 엔진</div><div class="val">Gemini 3 Flash</div></div>
      </div>
      <div class="cbox">
        <div class="ct">&#128591; 제작 크레딧</div>
        <div class="cb">이 프로그램은 전적으로<br><span class="cc">Claude (Anthropic)</span>가 설계하고 개발했습니다.
          <div class="cn">클로드 짱짱맨</div><div class="cs">전 과정을 클로드로 다함</div>
        </div>
      </div>
      <div class="fs">
        <h3>&#10024; v0.7 주요 기능</h3>
        <div class="fi"><div class="fic">&#128260;</div><div><b>롤오버(Rollover)</b> - DSD->Excel 시 당기 실적을 전기 칸으로 자동 이월, 당기 칸 비우기</div></div>
        <div class="fi"><div class="fic">&#128270;</div><div><b>Python 수학 검증</b> - 대차평균(자산=부채+자본), 단위 자릿수 이상, 주석번호 매핑 자동 검사</div></div>
        <div class="fi"><div class="fic">&#129302;</div><div><b>AI 강화 검증</b> - Python 오류를 Gemini에 전달하여 맥락 있는 최종 검증 리포트 생성</div></div>
        <div class="fi"><div class="fic">&#128269;</div><div><b>DSD 비교 분석</b> - 두 DSD 파일 셀 단위 Diff + AI 변동 요약 + 빨간 하이라이트 Excel</div></div>
        <div class="fi"><div class="fic">&#128260;</div><div>DSD &harr; Excel 완전한 양방향 변환, XML 유효성 자동 검증</div></div>
        <div class="fi"><div class="fic">&#128163;</div><div>하트비트 감시 - 브라우저 닫으면 서버 자동 종료, 종료 시 API Key 자동 삭제</div></div>
      </div>
    </div>

  </div>
</div>

<script>
// ── API Key / 모델 관리 ───────────────────────────────────────────────────────
function loadKey(){
  var k=localStorage.getItem('easydsd_gemini_key')||'';
  document.getElementById('apiKey').value=k; updSt(k);
}
function saveKey(v){
  if(v) localStorage.setItem('easydsd_gemini_key',v);
  else  localStorage.removeItem('easydsd_gemini_key');
  updSt(v);
}
function clearKey(){
  localStorage.removeItem('easydsd_gemini_key');
  document.getElementById('apiKey').value=''; updSt('');
}
function getKey(){ return localStorage.getItem('easydsd_gemini_key')||''; }
function updSt(v){
  var e=document.getElementById('apiSt');
  if(v&&v.length>10){e.textContent='🟢 입력됨';e.style.color='#a5d6a7';}
  else{e.textContent='⚪ 미입력';e.style.color='rgba(255,255,255,.6)';}
}
function saveModel(v){ localStorage.setItem('easydsd_model',v); }
function getModel(){ return localStorage.getItem('easydsd_model')||'gemini-3-flash-preview'; }
function loadModel(){
  var m=getModel(); var sel=document.getElementById('modelSel');
  if(sel) sel.value=m;
}
loadKey(); loadModel();

// ── 하트비트 ──────────────────────────────────────────────────────────────────
setInterval(function(){ fetch('/api/heartbeat',{method:'POST'}).catch(function(){}); },2500);

// ── 탭 / 파일 / 드래그 ────────────────────────────────────────────────────────
var F={f1:null,f2:null,f3:null,f4:null,f5:null,f6:null};
function sw(n){
  document.querySelectorAll('.tab').forEach(function(t,i){t.classList.toggle('active',i===n);});
  document.querySelectorAll('.tab-content').forEach(function(t,i){t.classList.toggle('active',i===n);});
}
function sf(id,bid,dzId){
  var f=document.getElementById(id).files[0]; if(!f) return;
  F[id]=f;
  var b=document.getElementById(bid);
  b.textContent='✓  '+f.name+'  ('+(f.size/1024).toFixed(0)+' KB)';
  b.style.display='block';
  document.getElementById(dzId).style.borderColor='#1F4E79';
  chk();
}
function dov(e,id){e.preventDefault();document.getElementById(id).classList.add('drag-over');}
function dlv(id){document.getElementById(id).classList.remove('drag-over');}
function ddrop(e,fid,did){
  e.preventDefault(); dlv(did);
  var dt=e.dataTransfer; if(!dt.files.length) return;
  var inp=document.getElementById(fid);
  var tr=new DataTransfer(); tr.items.add(dt.files[0]); inp.files=tr.files;
  sf(fid,fid.replace('f','fb'),did);
}
function chk(){
  document.getElementById('btn1').disabled=!F.f1;
  document.getElementById('btn2').disabled=!(F.f2&&F.f3);
  document.getElementById('btn3').disabled=!F.f4;
  document.getElementById('btn4').disabled=!(F.f5&&F.f6);
}
function hide(n){
  ['ok','er'].forEach(function(p){document.getElementById(p+n).style.display='none';});
}

// ── 프로그레스 ────────────────────────────────────────────────────────────────
var piv=null;
function sp(n,msg,isAI){
  hide(n);
  if(n===3){document.getElementById('vrBox').style.display='none';document.getElementById('pycheckBox').style.display='none';}
  var pw=document.getElementById('pw'+n); pw.style.display='block';
  document.getElementById('pt'+n).textContent=msg;
  document.getElementById('pf'+n).style.width='0%';
  var w=0; piv=setInterval(function(){w=Math.min(w+(isAI?1:4),88);document.getElementById('pf'+n).style.width=w+'%';},isAI?400:200);
}
function ep(n){
  clearInterval(piv);
  document.getElementById('pf'+n).style.width='100%';
  setTimeout(function(){document.getElementById('pw'+n).style.display='none';},500);
}
function sok(n,t,s,blob,fname){
  var b=document.getElementById('ok'+n); b.style.display='flex';
  document.getElementById('ok'+n+'t').textContent=t;
  document.getElementById('ok'+n+'s').textContent=s;
  if(blob){var dl=document.getElementById('dl'+n);dl.href=URL.createObjectURL(blob);dl.download=fname;}
}
function ser(n,msg){
  var b=document.getElementById('er'+n); b.style.display='flex';
  document.getElementById('er'+n+'m').textContent=msg;
}
function showPyCheck(pyResult){
  var box=document.getElementById('pycheckBox');
  var items=document.getElementById('pycheckItems');
  items.innerHTML='';
  var all=(pyResult.errors||[]).map(function(e){return {t:'err',v:e};})
    .concat((pyResult.warnings||[]).map(function(w){return {t:'warn',v:w};}))
    .concat((pyResult.info||[]).map(function(i){return {t:'info',v:i};}));
  if(all.length===0){items.innerHTML='<div class="pycheck-item info">이상 없음</div>'; box.style.display='block'; return;}
  all.forEach(function(item){
    var d=document.createElement('div');
    d.className='pycheck-item '+item.t;
    d.textContent=(item.t==='err'?'❌ ':item.t==='warn'?'⚠️ ':'✅ ')+item.v;
    items.appendChild(d);
  });
  box.style.display='block';
}
var S1=['DSD 파일 분석 중...','테이블 파싱 중...','Excel 시트 생성 중...'];
var S1A=['DSD 분석 중...','Gemini AI 분류 중... (15~30초 소요)','AI 시트명 적용 중...'];
var S2=['매핑 구성 중...','XML 패치 적용 중...','DSD 생성 중...'];
var S3=['Excel 데이터 추출 중...','Python 수학 검사 중...','Gemini AI 교차 검증 중... (30~60초 소요)','결과 시트 생성 중...'];
var S4=['DSD XML 파싱 중...','셀 단위 비교 중...','Gemini AI 요약 중... (10~20초 소요)','리포트 생성 중...'];
function anim(n,steps,isAI){
  var i=0;
  return setInterval(function(){
    if(i<steps.length) document.getElementById('pt'+n).textContent=steps[i++];
  },isAI?5000:1200);
}

// ── run1: DSD -> Excel ────────────────────────────────────────────────────────
async function run1(){
  if(!F.f1) return;
  document.getElementById('btn1').disabled=true;
  var useAI=document.getElementById('chkAI').checked;
  var doRoll=document.getElementById('chkRollover').checked;
  var key=getKey();
  if(useAI&&!key){ser(1,'AI 분류를 사용하려면 Gemini API Key를 입력해주세요.');document.getElementById('btn1').disabled=false;return;}
  sp(1,useAI?S1A[0]:S1[0],useAI); var iv=anim(1,useAI?S1A:S1,useAI);
  try{
    var fd=new FormData();
    fd.append('dsd',F.f1);
    fd.append('ai_classify',useAI?'1':'0');
    fd.append('rollover',doRoll?'1':'0');
    fd.append('api_key',key);
    fd.append('model',getModel());
    var r=await fetch('/api/dsd2excel',{method:'POST',body:fd});
    clearInterval(iv); ep(1);
    if(!r.ok){var e=await r.json();throw new Error(e.error||'변환 실패');}
    var blob=await r.blob();
    var info=JSON.parse(r.headers.get('X-Info')||'{}');
    var fname=F.f1.name.replace(/\.dsd$/i,'')+'.xlsx';
    var sub='시트 '+info.sheets+'개 · 수정가능 셀 '+info.cells+'개 · 재무4표 '+info.fin+'개';
    if(doRoll) sub+=' · 🔄롤오버 적용';
    if(useAI)  sub+=' · 🤖AI분류 적용';
    sok(1,'변환 완료! Excel 파일을 다운로드하세요',sub,blob,fname);
  }catch(e){clearInterval(iv);ep(1);ser(1,e.message);}
  document.getElementById('btn1').disabled=false;
}

// ── run2: Excel -> DSD ────────────────────────────────────────────────────────
async function run2(){
  if(!F.f2||!F.f3) return;
  document.getElementById('btn2').disabled=true;
  sp(2,S2[0]); var iv=anim(2,S2);
  try{
    var fd=new FormData(); fd.append('orig_dsd',F.f2); fd.append('xlsx',F.f3);
    var r=await fetch('/api/excel2dsd',{method:'POST',body:fd});
    clearInterval(iv); ep(2);
    if(!r.ok){var e=await r.json();throw new Error(e.error||'변환 실패');}
    var blob=await r.blob();
    var info=JSON.parse(r.headers.get('X-Info')||'{}');
    var fname=F.f2.name.replace(/\.dsd$/i,'')+'_수정.dsd';
    sok(2,'DSD 파일 생성 완료!',info.tables+'개 테이블 · '+info.cells+'개 셀 수정 · XML '+(info.xml_ok?'✓ 정상':'✗ 오류'),blob,fname);
  }catch(e){clearInterval(iv);ep(2);ser(2,e.message);}
  document.getElementById('btn2').disabled=false;
}

// ── run3: AI 검증 ─────────────────────────────────────────────────────────────
async function run3(){
  if(!F.f4) return;
  var key=getKey();
  document.getElementById('btn3').disabled=true;
  sp(3,S3[0],true); var iv=anim(3,S3,true);
  try{
    var fd=new FormData(); fd.append('xlsx',F.f4); fd.append('api_key',key); fd.append('model',getModel());
    var r=await fetch('/api/verify_excel',{method:'POST',body:fd});
    clearInterval(iv); ep(3);
    if(!r.ok){var e=await r.json();throw new Error(e.error||'검증 실패');}
    var blob=await r.blob();
    var info=JSON.parse(r.headers.get('X-Info')||'{}');
    var fname=F.f4.name.replace(/\.xlsx$/i,'')+'_AI검증.xlsx';
    if(info.py_result) showPyCheck(info.py_result);
    var sub='재무시트 '+info.fin_sheets+'개 · 주석시트 '+info.note_sheets+'개 분석';
    var pyE=(info.py_result&&info.py_result.errors)?info.py_result.errors.length:0;
    var pyW=(info.py_result&&info.py_result.warnings)?info.py_result.warnings.length:0;
    if(pyE>0) sub+=' · ❌오류 '+pyE+'건';
    if(pyW>0) sub+=' · ⚠️경고 '+pyW+'건';
    if(!key) sub+=' · (Gemini 검증 생략)';
    sok(3,'AI 검증 완료! 결과를 다운로드하세요',sub,blob,fname);
    if(info.preview){document.getElementById('vrText').textContent=info.preview;document.getElementById('vrBox').style.display='block';}
  }catch(e){clearInterval(iv);ep(3);ser(3,e.message);}
  document.getElementById('btn3').disabled=false;
}

// ── run4: DSD 비교 분석 ───────────────────────────────────────────────────────
async function run4(){
  if(!F.f5||!F.f6) return;
  document.getElementById('btn4').disabled=true;
  sp(4,S4[0],true); var iv=anim(4,S4,true);
  try{
    var fd=new FormData();
    fd.append('dsd_a',F.f5); fd.append('dsd_b',F.f6);
    fd.append('api_key',getKey()); fd.append('model',getModel());
    var r=await fetch('/api/compare_dsd',{method:'POST',body:fd});
    clearInterval(iv); ep(4);
    if(!r.ok){var e=await r.json();throw new Error(e.error||'비교 실패');}
    var blob=await r.blob();
    var info=JSON.parse(r.headers.get('X-Info')||'{}');
    var fname='DSD_비교분석_'+new Date().toISOString().slice(0,10)+'.xlsx';
    sok(4,'비교 분석 완료!','EXTRACTION 변경 '+info.ext_diffs+'건 · TABLE 셀 변경 '+info.cell_diffs+'건'+(getKey()?'':'  (AI 요약 생략)'),blob,fname);
  }catch(e){clearInterval(iv);ep(4);ser(4,e.message);}
  document.getElementById('btn4').disabled=false;
}

// ── 종료 ─────────────────────────────────────────────────────────────────────
function showKill(){ document.getElementById('killModal').classList.add('show'); }
function hideKill(){ document.getElementById('killModal').classList.remove('show'); }
async function doKill(){
  hideKill();
  localStorage.removeItem('easydsd_gemini_key');
  localStorage.removeItem('easydsd_model');
  try{ await fetch('/api/shutdown',{method:'POST'}); }catch(e){}
  document.body.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif;color:#556;font-size:15px;">서버가 종료되었습니다. 이 탭을 닫으세요.</div>';
}
</script>
</body>
</html>"""


# ── Flask 라우트 ──────────────────────────────────────────────────────────────
@app.route('/')
def index(): return render_template_string(HTML)

@app.route('/api/heartbeat', methods=['POST'])
def api_heartbeat():
    global _last_ping; _last_ping=time.time(); return jsonify(ok=True)

@app.route('/api/shutdown', methods=['POST'])
def api_shutdown():
    threading.Thread(target=lambda:(time.sleep(0.3),os._exit(0)),daemon=True).start()
    return jsonify(ok=True)

@app.route('/api/dsd2excel', methods=['POST'])
def api_dsd2excel():
    try:
        dsd_bytes   = request.files['dsd'].read()
        ai_classify = request.form.get('ai_classify','0')=='1'
        do_rollover = request.form.get('rollover','0')=='1'
        api_key     = request.form.get('api_key','').strip()
        model_name  = request.form.get('model','gemini-3-flash-preview').strip()
        ai_mapping  = {}
        if ai_classify:
            if not api_key:
                return jsonify(error='AI 분류를 사용하려면 Gemini API Key를 입력해주세요.'), 400
            xml_raw = zipfile.ZipFile(io.BytesIO(dsd_bytes)).read('contents.xml').decode('utf-8',errors='replace')
            _, tables = parse_xml(xml_raw)
            ai_mapping = gemini_classify_tables(api_key, tables, model_name)
        xlsx = dsd_to_excel_bytes(dsd_bytes, ai_mapping or None, do_rollover=do_rollover)
        wb   = openpyxl.load_workbook(io.BytesIO(xlsx), data_only=True)
        cells = sum(1 for ws in wb.worksheets for row in ws.iter_rows()
                    for cell in row if cell.fill and cell.fill.fill_type=='solid'
                    and cell.fill.fgColor and cell.fill.fgColor.type=='rgb'
                    and cell.fill.fgColor.rgb.upper().endswith(EDIT_COLOR.upper()))
        fin = [ws.title for ws in wb.worksheets if any(ws.title.startswith(e) for e in FIN_PREFIXES)]
        resp = send_file(io.BytesIO(xlsx),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name='converted.xlsx')
        resp.headers['X-Info'] = json.dumps(
            {'sheets':len(wb.sheetnames),'cells':cells,'fin':len(fin),'ai':bool(ai_mapping),'rollover':do_rollover})
        return resp
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.route('/api/excel2dsd', methods=['POST'])
def api_excel2dsd():
    try:
        orig = request.files['orig_dsd'].read()
        xlsx = request.files['xlsx'].read()
        dsd  = excel_to_dsd_bytes(orig, xlsx)
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(io.BytesIO(dsd)) as z:
            xt = z.read('contents.xml').decode('utf-8')
        xml_ok = True
        try: ET.fromstring(xt)
        except: xml_ok = False
        wb = openpyxl.load_workbook(io.BytesIO(xlsx), data_only=True)
        tc = tb = 0
        for sname in wb.sheetnames:
            if sname in ('📋사용안내','_원본XML','📊요약수치'): continue
            ws = wb[sname]
            cnt = sum(1 for row in ws.iter_rows() for cell in row
                      if cell.fill and cell.fill.fill_type=='solid'
                      and cell.fill.fgColor and cell.fill.fgColor.type=='rgb'
                      and cell.fill.fgColor.rgb.upper().endswith(EDIT_COLOR.upper()))
            if cnt: tb+=1; tc+=cnt
        resp = send_file(io.BytesIO(dsd), mimetype='application/octet-stream',
                         as_attachment=True, download_name='output.dsd')
        resp.headers['X-Info'] = json.dumps({'tables':tb,'cells':tc,'xml_ok':xml_ok})
        return resp
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.route('/api/verify_excel', methods=['POST'])
def api_verify_excel():
    try:
        xlsx_bytes = request.files['xlsx'].read()
        api_key    = request.form.get('api_key','').strip()
        model_name = request.form.get('model','gemini-3-flash-preview').strip()

        # ① Python 수학 검사 (API Key 없어도 실행)
        py_result = python_verify(xlsx_bytes)

        # ② 재무/주석 데이터 추출
        fin_data, note_data = extract_fin_and_notes(xlsx_bytes)
        if not fin_data:
            return jsonify(error='재무제표 시트(🏦💹📈💰)를 찾을 수 없습니다.'), 400

        # ③ Gemini AI 검증 (API Key 있을 때만)
        if api_key:
            verify_result = gemini_verify_enhanced(api_key, fin_data, note_data, py_result, model_name)
        else:
            err_lines  = [f"[오류] {e}" for e in py_result.get('errors', [])]
            warn_lines = [f"[경고] {w}" for w in py_result.get('warnings', [])]
            info_lines = [f"[정보] {i}" for i in py_result.get('info', [])]
            verify_result = (
                "## ✅ 파이썬 수학 검사 결과\n" +
                ('\n'.join(err_lines+warn_lines+info_lines) or '이상 없음') +
                "\n\n## 📋 종합 의견\nGemini API Key가 없어 AI 교차 검증은 생략되었습니다.\nPython 자동 검사 결과만 포함됩니다."
            )

        # ④ 결과 Excel 저장
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        if '🤖AI검증결과' in wb.sheetnames: del wb['🤖AI검증결과']
        ws_v = wb.create_sheet('🤖AI검증결과', 0)
        ws_v.sheet_view.showGridLines = False
        tc = ws_v.cell(1, 1, '🤖 Gemini AI + Python 재무제표 검증 결과 (easydsd v0.7)')
        tc.fill = PatternFill('solid', fgColor='4A148C')
        tc.font = Font(color='FFFFFF', bold=True, size=12)
        tc.alignment = Alignment(horizontal='left', vertical='center')
        ws_v.merge_cells('A1:F1'); ws_v.row_dimensions[1].height = 28
        sc = ws_v.cell(2, 1,
            f'생성: {time.strftime("%Y-%m-%d %H:%M")}  |  재무: {len(fin_data)}개  |  주석: {len(note_data)}개  |'
            f'  Python오류: {len(py_result.get("errors",[]))}건  경고: {len(py_result.get("warnings",[]))}건')
        sc.font = Font(color='7B1FA2', size=9, italic=True); ws_v.row_dimensions[2].height = 16
        COLOR_MAP = {
            '## ✅': ('E8F5E9','1B5E20'),
            '## ❌': ('FFEBEE','B71C1C'),
            '## ⚠': ('FFF8E1','E65100'),
            '## 📋': ('E3F2FD','0D47A1'),
        }
        for ri, line in enumerate(verify_result.split('\n'), 4):
            cell = ws_v.cell(ri, 1, line)
            matched = next(((fg,fc) for k,(fg,fc) in COLOR_MAP.items() if line.startswith(k)), None)
            if matched:
                cell.fill = PatternFill('solid', fgColor=matched[0])
                cell.font = Font(bold=True, size=10, color=matched[1])
            else:
                cell.font = Font(size=9)
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            ws_v.row_dimensions[ri].height = 16
        ws_v.column_dimensions['A'].width = 90
        buf = io.BytesIO(); wb.save(buf)
        preview = verify_result[:600] + ('...' if len(verify_result)>600 else '')
        resp = send_file(io.BytesIO(buf.getvalue()),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name='verified.xlsx')
        resp.headers['X-Info'] = json.dumps({
            'fin_sheets' : len(fin_data),
            'note_sheets': len(note_data),
            'preview'    : preview,
            'py_result'  : py_result,
        })
        return resp
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.route('/api/compare_dsd', methods=['POST'])
def api_compare_dsd():
    try:
        dsd_a    = request.files['dsd_a'].read()
        dsd_b    = request.files['dsd_b'].read()
        api_key  = request.form.get('api_key','').strip()
        model_name = request.form.get('model','gemini-3-flash-preview').strip()
        result   = compare_dsd_bytes(dsd_a, dsd_b, api_key, model_name)
        # 통계 계산
        da = parse_dsd_tables(dsd_a); db = parse_dsd_tables(dsd_b)
        ext_diffs = sum(1 for k in set(list(da['exts'].keys())+list(db['exts'].keys()))
                        if da['exts'].get(k)!=db['exts'].get(k))
        all_ids = sorted(set(list(da['tables'].keys())+list(db['tables'].keys())))
        cell_diffs = 0
        for ti in all_ids:
            ra = da['tables'].get(ti,[]); rb = db['tables'].get(ti,[])
            for ri in range(max(len(ra),len(rb))):
                a = ra[ri] if ri<len(ra) else []; b = rb[ri] if ri<len(rb) else []
                for ci in range(max(len(a),len(b))):
                    if (a[ci] if ci<len(a) else '') != (b[ci] if ci<len(b) else ''):
                        cell_diffs += 1
        resp = send_file(io.BytesIO(result),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name='diff_report.xlsx')
        resp.headers['X-Info'] = json.dumps({'ext_diffs':ext_diffs,'cell_diffs':cell_diffs})
        return resp
    except Exception as e:
        return jsonify(error=str(e)), 500

# ── 실행 ─────────────────────────────────────────────────────────────────────
def open_browser():
    time.sleep(1.5)
    webbrowser.open(f'http://127.0.0.1:{PORT}')

if __name__ == '__main__':
    print('='*54)
    print('  easydsd v0.7 - DART 감사보고서 변환 + AI')
    print(f'  http://127.0.0.1:{PORT}')
    print('  종료: 브라우저 종료 버튼 or Ctrl+C')
    print('='*54)
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='127.0.0.1', port=PORT, debug=False)
