#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""easydsd v0.4 - DART 감사보고서 변환 도구 + Gemini AI"""

import os, re, sys, io, zipfile, threading, webbrowser, socket, time, json

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── 라이브러리 자동 설치 ───────────────────────────────────────────────────────
try:
    from flask import Flask, request, send_file, jsonify, render_template_string
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable,'-m','pip','install','flask','openpyxl','-q'])
    from flask import Flask, request, send_file, jsonify, render_template_string
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter

try:
    import google.generativeai as genai
    GENAI_AVAILABLE = True
except ImportError:
    print("  google-generativeai 설치 중... (최초 1회)")
    try:
        import subprocess
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install",
             "google-generativeai", "grpcio", "--quiet"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        import google.generativeai as genai
        GENAI_AVAILABLE = True
        print("  google-generativeai 설치 완료!")
    except Exception as _e:
        print(f"  google-generativeai 설치 실패 ({_e}) — AI 기능 비활성")
        GENAI_AVAILABLE = False
        genai = None

# ── 기본 상수 ──────────────────────────────────────────────────────────────────
def find_free_port(start=5000, end=5099):
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(('127.0.0.1', port)) != 0:
                return port
    return start

PORT       = find_free_port()
EDIT_COLOR = 'FFF2CC'
C = {'navy':'1F4E79','blue':'2E75B6','lblue':'DEEAF1',
     'yellow':'FFF2CC','white':'FFFFFF','lgray':'F2F2F2','orange':'C55A11'}
FIN_TABLE_MAP = [
    (['재 무 상 태 표'],      '🏦재무상태표'),
    (['포 괄 손 익 계 산 서'],'💹포괄손익계산서'),
    (['자 본 변 동 표'],      '📈자본변동표'),
    (['현 금 흐 름 표'],      '💰현금흐름표'),
]
EXT_DESC = {
    'TOT_ASSETS':'총자산(백만원)','TOT_DEBTS':'총부채(백만원)',
    'TOT_SALES':'매출액(백만원)','TOT_EMPL':'총직원수',
    'GMSH_DATE':'주총일자(YYYYMMDD)','SUPV_OPIN':'감사의견코드',
    'AUDIT_CIK':'감사인CIK','CRP_RGS_NO':'법인등록번호',
}

# ── 스타일 헬퍼 ────────────────────────────────────────────────────────────────
def fill(c): return PatternFill('solid', fgColor=c)
def fnt(color='000000',bold=False,size=9,italic=False):
    return Font(color=color,bold=bold,size=size,italic=italic)
def aln(h='left',v='center',wrap=False):
    return Alignment(horizontal=h,vertical=v,wrap_text=wrap)

# ── XML 파싱 유틸 ──────────────────────────────────────────────────────────────
def clean_cr(s, as_newline=False):
    repl = '\n' if as_newline else ' '
    s = s.replace('&amp;cr;', repl).replace('&cr;', repl)
    return re.sub(r'\s+', ' ', s).strip()

def clean_title(s):
    s = clean_cr(s, False)
    s = s.replace('&amp;','&').replace('&lt;','<').replace('&gt;','>').replace('&quot;','"')
    return re.sub(r'\s+', ' ', s).strip()

def is_blank_title(s):
    return len(re.sub(r'[&;a-z]+','',s).strip()) == 0

def parse_cell(m):
    attrs = m.group(1)
    val   = re.sub(r'<[^>]+>','',m.group(0))
    val   = (val.replace('&amp;cr;','\n').replace('&amp;','&')
               .replace('&lt;','<').replace('&gt;','>').replace('&quot;','"')
               .replace('&cr;','\n').strip())
    cs  = int(x.group(1)) if (x:=re.search(r'COLSPAN="(\d+)"',attrs)) else 1
    tag = re.match(r'<([A-Z]+)',m.group(0)).group(1)
    return dict(value=val, colspan=cs, tag=tag)

def is_num_or_decimal(val):
    v = (val.strip().replace(',','').replace('(','').replace(')','')
             .replace('%','').replace('-','').replace(' ','').split('\n')[0])
    if not v: return False
    try: float(v); return True
    except: return False

def parse_xml(xml):
    exts = re.findall(r'<EXTRACTION[^>]*ACODE="([^"]+)"[^>]*>([^<]+)</EXTRACTION>',xml)
    tables = []
    for ti,tm in enumerate(re.finditer(r'<TABLE([^>]*)>(.*?)</TABLE>',xml,re.DOTALL)):
        ctx   = xml[max(0,tm.start()-600):tm.start()]
        tbody = tm.group(0)
        fin_label = next((lbl for kws,lbl in FIN_TABLE_MAP
                          if any(kw in ctx or kw in tbody for kw in kws)),'')
        raw = re.findall(r'<(?:TITLE|P)[^>]*>([^<]{3,80})</(?:TITLE|P)>',ctx)
        ctx_titles = [clean_title(t) for t in raw if not is_blank_title(t) and len(clean_title(t))>1]
        rows=[]
        for tr in re.finditer(r'<TR[^>]*>(.*?)</TR>',tm.group(2),re.DOTALL):
            cells=[parse_cell(cm) for cm in re.finditer(
                r'<(?:TD|TH|TU|TE)([^>]*)>.*?</(?:TD|TH|TU|TE)>',tr.group(1),re.DOTALL)]
            if cells: rows.append(cells)
        tables.append(dict(idx=ti,fin_label=fin_label,
                           ctx_title=(ctx_titles[-1] if ctx_titles else ''),
                           rows=rows,start=tm.start()))
    return exts, tables

# ── Gemini: AI 시트명 분류 ─────────────────────────────────────────────────────
def gemini_classify_tables(api_key:str, tables:list) -> dict:
    """TABLE 목록 → Gemini → {table_idx: 추천시트명}"""
    if not GENAI_AVAILABLE or not api_key: return {}
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        summaries=[]
        for tbl in tables[:60]:
            vals=[c['value'] for row in tbl['rows'][:3] for c in row if c['value'].strip()][:6]
            summaries.append(f"TABLE[{tbl['idx']}] ctx={tbl['ctx_title']!r} fin={tbl['fin_label']!r} 샘플={vals}")
        prompt=f"""한국 DART 감사보고서 TABLE 목록입니다. 각 TABLE의 엑셀 시트명을 제안해주세요.
규칙: 재무상태표→"🏦재무상태표", 포괄손익→"💹포괄손익계산서", 자본변동→"📈자본변동표", 현금흐름→"💰현금흐름표",
주석→"📝주석_[주제3~5자]", 서문/목차/감사의견→"📄서문", 시트명31자이내 특수문자(/\\*?[]:)불가

TABLE목록:
{chr(10).join(summaries)}

JSON만 응답: {{"매핑":[{{"idx":0,"시트명":"예시"}}]}}"""
        resp=model.generate_content(prompt)
        m=re.search(r'\{.*\}',resp.text.strip(),re.DOTALL)
        if not m: return {}
        data=json.loads(m.group(0))
        return {item['idx']:item['시트명'] for item in data.get('매핑',[])}
    except Exception as e:
        print(f"[Gemini classify] {e}"); return {}

# ── Gemini: AI 교차 검증 ──────────────────────────────────────────────────────
def gemini_verify_excel(api_key:str, fin_data:dict, note_data:dict) -> str:
    """재무제표 + 주석 → Gemini → 교차검증 결과 텍스트"""
    if not GENAI_AVAILABLE or not api_key:
        return "❌ Gemini API 키가 없거나 라이브러리가 설치되지 않았습니다."
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        fin_text  = json.dumps(fin_data,  ensure_ascii=False, indent=2)[:8000]
        note_text = json.dumps(note_data, ensure_ascii=False, indent=2)[:8000]
        prompt=f"""당신은 한국 공인회계사(CPA) 수준의 재무제표 검증 전문가입니다.
아래 재무제표 본문과 주석 데이터를 교차 검증해주세요.

[재무제표 본문]
{fin_text}

[주석 데이터]
{note_text}

검증 목표:
1. 재무상태표 주요 계정 금액 ↔ 해당 주석 세부 합계 일치 여부
2. 포괄손익계산서 항목 ↔ 주석 세부 내역 일치 여부
3. 불일치·확인불가 항목 명시

응답 형식:
## ✅ 일치 항목
(일치 항목 나열)

## ❌ 불일치 항목
(불일치 항목 + 차이 금액)

## ⚠️ 확인 필요 항목
(데이터 부족 등)

## 📋 종합 의견
(전체 요약)"""
        resp=model.generate_content(prompt)
        return resp.text.strip()
    except Exception as e:
        return f"❌ Gemini API 오류: {e}"

# ── Excel에서 재무/주석 데이터 추출 ───────────────────────────────────────────
def extract_fin_and_notes(xlsx_bytes:bytes) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    fin_data={}; note_data={}
    FIN_P=('🏦','💹','📈','💰')
    for sname in wb.sheetnames:
        if sname in ('📋사용안내','_원본XML','📊요약수치'): continue
        ws=wb[sname]
        rows_data=[]
        for row in ws.iter_rows(max_row=200,values_only=True):
            cleaned=[str(v).strip() if v is not None else '' for v in row]
            if any(cleaned): rows_data.append(cleaned)
        if any(sname.startswith(p) for p in FIN_P):
            fin_data[sname]=rows_data[:80]
        elif sname.startswith('📝'):
            note_data[sname]=rows_data[:50]
    return fin_data, note_data

# ── DSD → Excel 변환 ───────────────────────────────────────────────────────────
def dsd_to_excel_bytes(dsd_bytes:bytes, ai_mapping:dict=None) -> bytes:
    with zipfile.ZipFile(io.BytesIO(dsd_bytes)) as zf:
        files={n:zf.read(n) for n in zf.namelist()}
    xml      = files.get('contents.xml',b'').decode('utf-8',errors='replace')
    meta_xml = files.get('meta.xml',b'').decode('utf-8',errors='replace')
    exts, tables = parse_xml(xml)
    wb = openpyxl.Workbook()

    # 사용안내
    ws0=wb.active; ws0.title='📋사용안내'; ws0.sheet_view.showGridLines=False
    guide=[
        ('DART 감사보고서 DSD - Excel 변환 도구 (easydsd v0.4)',True,C['white'],C['navy'],13),
        ('',False,'','',8),
        ('【 작업 순서 】',True,C['navy'],C['lblue'],11),
        ('  1. 노란색 셀을 당해년도 숫자/텍스트로 수정하세요',False,'000000',C['white'],10),
        ('  2. 저장 후 "Excel → DSD" 탭에서 변환하세요',False,'000000',C['white'],10),
        ('',False,'','',8),
        ('【 색상 범례 】',True,C['navy'],C['lblue'],11),
        ('  노란색 = 수정 가능 (금액, 주주명, 지분율, 텍스트 모두)',False,'000000',C['yellow'],10),
        ('  파란색 = 헤더 (수정 불필요)',False,C['white'],C['navy'],10),
        ('',False,'','',8),
        ('【 주의사항 】',True,C['navy'],C['lblue'],11),
        ('  _원본XML 시트는 절대 수정/삭제 금지!',False,C['orange'],C['white'],10),
    ]
    for ri,(txt,bold,fg,bg,sz) in enumerate(guide,1):
        c=ws0.cell(ri,1,txt); c.font=fnt(fg or '000000',bold=bold,size=sz)
        if bg: c.fill=fill(bg)
        c.alignment=aln('left',wrap=True); ws0.row_dimensions[ri].height=21
    ws0.column_dimensions['A'].width=65

    # 요약수치
    ws_e=wb.create_sheet('📊요약수치'); ws_e.sheet_view.showGridLines=False
    for ci,(h,w) in enumerate([('ACODE',15),('값 (수정가능)',22),('설명',28)],1):
        c=ws_e.cell(1,ci,h); c.fill=fill(C['navy']); c.font=fnt(C['white'],bold=True)
        c.alignment=aln('center'); ws_e.column_dimensions[get_column_letter(ci)].width=w
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

    # 시트 생성 헬퍼
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

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── Excel → DSD 변환 ───────────────────────────────────────────────────────────
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

def is_edit(cell):
    f=cell.fill
    if f and f.fill_type=='solid':
        fg=f.fgColor
        if fg and fg.type=='rgb': return fg.rgb.upper().endswith(EDIT_COLOR.upper())
    return False

def excel_to_dsd_bytes(orig_dsd_bytes:bytes, xlsx_bytes:bytes) -> bytes:
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
        for k,(t_idx,excel_start_row) in enumerate(t_info):
            if t_idx>=len(table_positions): continue
            if excel_start_row>=0:
                nxt=(t_info[k+1][1] if k+1<len(t_info) and t_info[k+1][1]>=0 else 99999)
                local_map={(r-excel_start_row,c):v for (r,c),v in all_ch.items()
                           if excel_start_row<=r<nxt}
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

# ── Flask 앱 ───────────────────────────────────────────────────────────────────
app=Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=100*1024*1024

_last_ping=time.time()
def _watchdog():
    time.sleep(12)
    while True:
        time.sleep(2)
        if time.time()-_last_ping>8: os._exit(0)
threading.Thread(target=_watchdog,daemon=True).start()

HTML=r'''<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>easydsd v0.4</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Malgun Gothic','맑은 고딕',sans-serif;background:#f0f4f8;color:#1a1a2e;min-height:100vh}
.header{background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);color:white;padding:14px 24px;box-shadow:0 4px 20px rgba(31,78,121,.3)}
.hd-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.hd-top h1{font-size:17px;font-weight:700}
.hd-top p{font-size:11px;opacity:.75;margin-top:2px}
.hd-right{display:flex;align-items:center;gap:8px;flex-shrink:0}
.hd-badge{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:20px;padding:3px 12px;font-size:11px;font-weight:600}
.kill-btn{background:#c0392b;color:white;border:none;border-radius:7px;padding:6px 12px;font-size:11px;font-weight:700;cursor:pointer}
.kill-btn:hover{background:#e74c3c}
.api-bar{display:flex;align-items:center;gap:8px;background:rgba(0,0,0,.18);border-radius:8px;padding:8px 12px}
.api-bar label{font-size:11px;font-weight:700;white-space:nowrap;opacity:.9}
.api-input{flex:1;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);border-radius:6px;color:white;padding:5px 10px;font-size:12px;font-family:monospace;min-width:0;outline:none}
.api-input::placeholder{opacity:.5}
.api-input:focus{background:rgba(255,255,255,.2);border-color:rgba(255,255,255,.5)}
.api-st{font-size:10px;white-space:nowrap;opacity:.8}
.api-note{font-size:10px;white-space:nowrap;background:rgba(255,193,7,.3);border:1px solid rgba(255,193,7,.5);border-radius:10px;padding:2px 7px;color:#fff8e1;font-weight:600}
.container{max-width:840px;margin:22px auto;padding:0 16px 60px}
.tabs{display:flex;gap:3px}
.tab{padding:9px 18px;border-radius:10px 10px 0 0;background:#cdd8e4;color:#4a6078;font-size:12px;font-weight:600;cursor:pointer;border:none;border-bottom:3px solid transparent;transition:all .2s}
.tab.active{background:white;color:#1F4E79;border-bottom:3px solid #1F4E79}
.tab:hover:not(.active){background:#bcccd8}
.tab.ai-tab{background:#2d1b4e;color:#b39ddb}
.tab.ai-tab.active{background:white;color:#6200ea;border-bottom:3px solid #6200ea}
.tab.ai-tab:hover:not(.active){background:#3d2b5e;color:#ce93d8}
.tab.dev-tab{background:#2a2a2a;color:#999}
.tab.dev-tab.active{background:white;color:#333;border-bottom:3px solid #666}
.tab.dev-tab:hover:not(.active){background:#3a3a3a;color:#bbb}
.card{background:white;border-radius:0 12px 12px 12px;box-shadow:0 4px 24px rgba(0,0,0,.08);padding:26px}
.tab-content{display:none}.tab-content.active{display:block}
.step{display:flex;gap:11px;align-items:flex-start;padding:13px;margin-bottom:11px;background:#f7f9fc;border-radius:10px;border-left:4px solid #2E75B6}
.step-num{min-width:25px;height:25px;border-radius:50%;background:#1F4E79;color:white;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:12px;flex-shrink:0}
.step-title{font-weight:700;font-size:13px;color:#1F4E79;margin-bottom:4px}
.step-desc{font-size:11px;color:#556;line-height:1.6}
.ai-check-row{display:flex;align-items:center;gap:8px;margin-top:10px;padding:9px 13px;background:#f3e5f5;border-radius:8px;border:1px solid #ce93d8}
.ai-check-row input[type=checkbox]{width:15px;height:15px;cursor:pointer;accent-color:#6200ea}
.ai-check-row label{font-size:12px;color:#4a148c;font-weight:600;cursor:pointer}
.ai-note{font-size:10px;color:#7b1fa2;margin-left:4px;opacity:.8}
.drop-zone{border:2px dashed #a0b8d0;border-radius:10px;padding:22px;text-align:center;cursor:pointer;transition:all .2s;background:#f7fbff;margin-top:8px}
.drop-zone:hover,.drop-zone.drag-over{border-color:#1F4E79;background:#e8f0f8}
.drop-zone .icon{font-size:26px;margin-bottom:4px}
.drop-zone .label{font-size:12px;color:#4a6078}
.drop-zone .sub{font-size:10px;color:#89a;margin-top:2px}
.file-badge{margin-top:6px;font-size:11px;color:#1F4E79;font-weight:600;display:none;background:#e8f0f8;padding:5px 10px;border-radius:6px}
.btn{width:100%;padding:11px;border:none;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;transition:all .2s;margin-top:12px}
.btn-blue{background:linear-gradient(135deg,#1F4E79,#2E75B6);color:white}
.btn-blue:hover{transform:translateY(-1px);box-shadow:0 5px 18px rgba(31,78,121,.35)}
.btn-blue:disabled{background:#a0b8c8;cursor:not-allowed;transform:none;box-shadow:none}
.btn-green{background:linear-gradient(135deg,#1a6b3a,#22a55a);color:white}
.btn-green:hover{transform:translateY(-1px);box-shadow:0 5px 18px rgba(26,107,58,.35)}
.btn-green:disabled{background:#8fc0a0;cursor:not-allowed;transform:none;box-shadow:none}
.btn-ai{background:linear-gradient(135deg,#4a148c,#7b1fa2);color:white}
.btn-ai:hover{transform:translateY(-1px);box-shadow:0 5px 18px rgba(74,20,140,.4)}
.btn-ai:disabled{background:#b39ddb;cursor:not-allowed;transform:none;box-shadow:none}
.prog-wrap{margin-top:12px;display:none}
.prog-bar{height:6px;background:#e0e8f0;border-radius:4px;overflow:hidden}
.prog-fill{height:100%;width:0%;background:linear-gradient(90deg,#1F4E79,#2E75B6);border-radius:4px;transition:width .35s ease}
.prog-fill.ai-fill{background:linear-gradient(90deg,#4a148c,#7b1fa2)}
.prog-text{font-size:11px;color:#4a6078;margin-top:4px;text-align:center}
.result{margin-top:12px;padding:12px 15px;border-radius:10px;display:none;align-items:center;gap:10px}
.result.ok{background:#e8f5ec;border:1px solid #6dbf8a}
.result.err{background:#fdecea;border:1px solid #e88}
.result.ai-ok{background:#f3e5f5;border:1px solid #ce93d8}
.r-icon{font-size:20px}
.r-body{flex:1}
.r-title{font-weight:700;font-size:13px}
.r-sub{font-size:11px;margin-top:2px;color:#556}
.dl-btn{padding:7px 13px;color:white;border:none;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap;text-decoration:none;display:inline-block;transition:background .15s}
.dl-btn.green{background:#1a6b3a}.dl-btn.green:hover{background:#145530}
.dl-btn.blue{background:#1F4E79}.dl-btn.blue:hover{background:#163a5e}
.dl-btn.purple{background:#4a148c}.dl-btn.purple:hover{background:#6a1b9a}
.legend{display:flex;gap:10px;flex-wrap:wrap;margin-top:7px}
.leg-item{display:flex;align-items:center;gap:5px;font-size:11px;color:#556}
.leg-dot{width:12px;height:12px;border-radius:3px;flex-shrink:0}
.ai-hdr{display:flex;align-items:center;gap:10px;padding:13px 15px;background:linear-gradient(135deg,#4a148c,#7b1fa2);border-radius:10px;margin-bottom:14px;color:white}
.ai-hdr .ai-ico{font-size:26px}
.ai-hdr h3{font-size:14px;font-weight:700}
.ai-hdr p{font-size:10px;opacity:.8;margin-top:2px}
.vr-box{margin-top:13px;display:none;background:#fafafa;border:1px solid #e0e0e0;border-radius:10px;padding:14px;max-height:380px;overflow-y:auto}
.vr-box h4{font-size:11px;font-weight:700;color:#4a148c;margin-bottom:7px}
.vr-box pre{font-size:11px;color:#333;white-space:pre-wrap;line-height:1.7;font-family:'Malgun Gothic',sans-serif}
.dev-profile{display:flex;align-items:center;gap:16px;padding:16px;background:linear-gradient(135deg,#1a1a2e,#16213e);border-radius:12px;margin-bottom:14px}
.dev-avatar{width:58px;height:58px;border-radius:50%;background:linear-gradient(135deg,#1F4E79,#2E75B6);display:flex;align-items:center;justify-content:center;font-size:24px;flex-shrink:0;border:3px solid rgba(255,255,255,.2)}
.dev-info h2{color:white;font-size:15px;font-weight:700;margin-bottom:2px}
.dev-sub{color:rgba(255,255,255,.6);font-size:10px;margin-bottom:6px}
.dev-badges{display:flex;gap:5px;flex-wrap:wrap}
.badge{border-radius:20px;padding:2px 9px;font-size:10px;font-weight:600}
.bg{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.8)}
.bg-gold{background:linear-gradient(135deg,#b8860b,#daa520);color:white}
.bg-tech{background:rgba(46,117,182,.5);border:1px solid rgba(46,117,182,.8);color:white}
.bg-ai{background:linear-gradient(135deg,#4a148c,#7b1fa2);color:white}
.ig{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:13px}
.ib{background:#f7f9fc;border-radius:10px;padding:11px 13px;border-left:3px solid #2E75B6}
.ib .lbl{font-size:10px;color:#89a;text-transform:uppercase;letter-spacing:.5px;margin-bottom:2px}
.ib .val{font-size:13px;font-weight:700;color:#1F4E79}
.ib .val a{color:#1F4E79;text-decoration:none}.ib .val a:hover{text-decoration:underline}
.cbox{background:#fffbf0;border:1px solid #e8d060;border-radius:12px;padding:15px;text-align:center;margin-bottom:13px}
.ct{font-size:12px;font-weight:700;color:#7a5500;margin-bottom:7px}
.cb{font-size:12px;color:#444;line-height:1.9}
.cn{font-size:15px;font-weight:800;color:#1a1a2e;margin:4px 0 1px}
.cs{font-size:10px;color:#888}
.cc{display:inline-block;background:linear-gradient(135deg,#7c4dff,#2196f3);color:white;border-radius:20px;padding:3px 11px;font-size:11px;font-weight:700;margin:0 3px;vertical-align:middle}
.fs h3{font-size:12px;font-weight:700;color:#1F4E79;margin-bottom:7px}
.fi{display:flex;align-items:flex-start;gap:7px;padding:6px 0;border-bottom:1px solid #f0f4f8;font-size:11px;color:#446;line-height:1.5}
.fi:last-child{border-bottom:none}
.fic{font-size:13px;flex-shrink:0;margin-top:1px}
.fw{color:#999;font-size:10px;margin-left:3px}
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:9999;align-items:center;justify-content:center}
.modal-overlay.show{display:flex}
.modal{background:white;border-radius:14px;padding:24px 28px;max-width:320px;width:90%;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal h3{font-size:14px;font-weight:700;margin-bottom:6px}
.modal p{font-size:12px;color:#556;margin-bottom:16px;line-height:1.6}
.mbtns{display:flex;gap:8px;justify-content:center}
.mbtns button{padding:8px 20px;border:none;border-radius:7px;font-size:12px;font-weight:700;cursor:pointer}
.mc{background:#e8eef4;color:#4a6078}.mc:hover{background:#d0dce8}
.mx{background:#c0392b;color:white}.mx:hover{background:#e74c3c}
</style>
</head>
<body>
<div class="header">
  <div class="hd-top">
    <div>
      <h1>&#128202; DART 감사보고서 변환 도구</h1>
      <p>DSD &#8596; Excel 양방향 변환 &nbsp;&#xB7;&nbsp; Gemini AI 검증 &nbsp;&#xB7;&nbsp; easydsd v0.4</p>
    </div>
    <div class="hd-right">
      <div class="hd-badge">v0.4</div>
      <button class="kill-btn" onclick="showKill()">&#x23FC; 종료</button>
    </div>
  </div>
  <div class="api-bar">
    <label>&#129302; Gemini API Key</label>
    <input class="api-input" id="apiKey" type="password"
      placeholder="AIza... (AI 기능 사용 시 입력 — 없어도 DSD 변환은 완벽 작동)"
      oninput="saveKey(this.value)" />
    <span class="api-note">&#x1F4CC; 선택사항</span>
    <span class="api-st" id="apiSt">&#x26AA; 미입력</span>
  </div>
  <div style="margin-top:5px;font-size:10px;opacity:.75">
    &#x1F449; API Key 없이도 DSD&#8596;Excel 변환은 정상 작동합니다. &nbsp;|&nbsp;
    <a href="https://aistudio.google.com/app/apikey" target="_blank"
       style="color:#a5d6a7;font-weight:600;">1분 만에 무료 API 키 발급받는 방법 &#x2197;</a>
  </div>
</div>

<div class="modal-overlay" id="killModal">
  <div class="modal">
    <h3>&#x26A0;&#xFE0F; 종료할까요?</h3>
    <p>서버가 완전히 종료됩니다.</p>
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
    <button class="tab dev-tab" onclick="sw(3)">개발자 정보</button>
  </div>
  <div class="card">

    <!-- 탭① DSD→Excel -->
    <div class="tab-content active" id="tab0">
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <div class="step-title">전년도 DSD 파일을 업로드하세요</div>
          <div class="step-desc">DART에서 제출한 .dsd 파일을 드래그하거나 클릭해 선택하세요.<br>변환된 Excel의 <b>노란색 셀</b>을 당해년도 숫자로 수정하시면 됩니다.</div>
          <div class="drop-zone" id="dz1" onclick="document.getElementById('f1').click()"
               ondragover="dov(event,'dz1')" ondragleave="dlv('dz1')" ondrop="ddrop(event,'f1','dz1')">
            <div class="icon">&#128194;</div>
            <div class="label">클릭하거나 파일을 여기에 끌어다 놓으세요</div>
            <div class="sub">.dsd 파일</div>
          </div>
          <input type="file" id="f1" accept=".dsd" style="display:none" onchange="sf('f1','fb1','dz1')">
          <div class="file-badge" id="fb1"></div>
          <div class="ai-check-row">
            <input type="checkbox" id="aiClassify">
            <label for="aiClassify">&#129302; AI를 이용해 재무제표 및 주석 스마트 분류하기</label>
            <span class="ai-note">(Gemini API Key 필요)</span>
          </div>
        </div>
      </div>
      <button class="btn btn-blue" id="btn1" onclick="run1()" disabled>&#128229;&nbsp; Excel 파일로 변환하기</button>
      <div class="prog-wrap" id="pw1"><div class="prog-bar"><div class="prog-fill" id="pf1"></div></div><div class="prog-text" id="pt1">변환 중...</div></div>
      <div class="result ok" id="ok1">
        <div class="r-icon">&#9989;</div>
        <div class="r-body">
          <div class="r-title" id="ok1t"></div><div class="r-sub" id="ok1s"></div>
          <div class="legend" style="margin-top:7px">
            <div class="leg-item"><div class="leg-dot" style="background:#FFF2CC;border:1px solid #ccc"></div>노란색=수정가능</div>
            <div class="leg-item"><div class="leg-dot" style="background:#1F4E79"></div>파란색=헤더</div>
            <div class="leg-item"><div class="leg-dot" style="background:#D9D9D9;border:1px solid #bbb"></div>회색=구분선</div>
          </div>
        </div>
        <a class="dl-btn green" id="dl1" href="#">&#11015; 다운로드</a>
      </div>
      <div class="result err" id="er1"><div class="r-icon">&#10060;</div><div class="r-body"><div class="r-title">변환 실패</div><div class="r-sub" id="er1m"></div></div></div>
    </div>

    <!-- 탭② Excel→DSD -->
    <div class="tab-content" id="tab1">
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <div class="step-title">원본 DSD 파일 업로드</div>
          <div class="step-desc">&#9312; 탭에서 사용했던 원본 .dsd 파일을 올려주세요.</div>
          <div class="drop-zone" id="dz2" onclick="document.getElementById('f2').click()"
               ondragover="dov(event,'dz2')" ondragleave="dlv('dz2')" ondrop="ddrop(event,'f2','dz2')">
            <div class="icon">&#128194;</div><div class="label">원본 DSD 파일</div><div class="sub">.dsd</div>
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
            <div class="icon">&#128202;</div><div class="label">수정된 Excel 파일</div><div class="sub">.xlsx</div>
          </div>
          <input type="file" id="f3" accept=".xlsx" style="display:none" onchange="sf('f3','fb3','dz3')">
          <div class="file-badge" id="fb3"></div>
        </div>
      </div>
      <button class="btn btn-green" id="btn2" onclick="run2()" disabled>&#128228;&nbsp; DSD 파일로 변환하기</button>
      <div class="prog-wrap" id="pw2"><div class="prog-bar"><div class="prog-fill" id="pf2"></div></div><div class="prog-text" id="pt2">변환 중...</div></div>
      <div class="result ok" id="ok2">
        <div class="r-icon">&#9989;</div>
        <div class="r-body"><div class="r-title" id="ok2t"></div><div class="r-sub" id="ok2s"></div></div>
        <a class="dl-btn blue" id="dl2" href="#">&#11015; DSD 다운로드</a>
      </div>
      <div class="result err" id="er2"><div class="r-icon">&#10060;</div><div class="r-body"><div class="r-title">변환 실패</div><div class="r-sub" id="er2m"></div></div></div>
    </div>

    <!-- 탭③ AI 검증 -->
    <div class="tab-content" id="tab2">
      <div style="background:#f3e5f5;border:1px solid #ce93d8;border-radius:8px;padding:10px 14px;margin-bottom:12px;font-size:11px;color:#4a148c;line-height:1.6">
        &#x2139;&#xFE0F; <b>이 기능은 선택사항입니다.</b> Gemini API Key를 입력한 사용자만 이용할 수 있습니다.<br>
        API Key가 없어도 &#9312; DSD&#8594;Excel, &#9313; Excel&#8594;DSD 변환은 완벽히 작동합니다.
      </div>
      <div class="ai-hdr">
        <div class="ai-ico">&#129302;</div>
        <div><h3>AI 재무제표 교차 검증</h3><p>Gemini AI가 재무제표 본문↔주석 금액 일치 여부를 자동으로 검증합니다</p></div>
      </div>
      <div class="step">
        <div class="step-num">1</div>
        <div class="step-body">
          <div class="step-title">수정한 Excel 파일 업로드</div>
          <div class="step-desc">easydsd로 변환 후 수정한 .xlsx 파일을 올려주세요.<br>
            재무상태표·포괄손익 등과 주석 시트의 금액을 AI가 교차 검증합니다.<br>
            <b style="color:#4a148c">⚠️ 상단에 Gemini API Key가 입력되어 있어야 합니다.</b></div>
          <div class="drop-zone" id="dz4" onclick="document.getElementById('f4').click()"
               ondragover="dov(event,'dz4')" ondragleave="dlv('dz4')" ondrop="ddrop(event,'f4','dz4')">
            <div class="icon">&#128202;</div><div class="label">수정된 Excel 파일 (.xlsx)</div><div class="sub">easydsd로 변환한 Excel</div>
          </div>
          <input type="file" id="f4" accept=".xlsx" style="display:none" onchange="sf('f4','fb4','dz4')">
          <div class="file-badge" id="fb4"></div>
        </div>
      </div>
      <button class="btn btn-ai" id="btn3" onclick="run3()" disabled>&#129302;&nbsp; AI 교차 검증 실행하기</button>
      <div class="prog-wrap" id="pw3"><div class="prog-bar"><div class="prog-fill ai-fill" id="pf3"></div></div><div class="prog-text" id="pt3">AI 분석 중...</div></div>
      <div class="result ai-ok" id="ok3">
        <div class="r-icon">&#129302;</div>
        <div class="r-body"><div class="r-title" id="ok3t"></div><div class="r-sub" id="ok3s"></div></div>
        <a class="dl-btn purple" id="dl3" href="#">&#11015; 검증결과 다운로드</a>
      </div>
      <div class="result err" id="er3"><div class="r-icon">&#10060;</div><div class="r-body"><div class="r-title">검증 실패</div><div class="r-sub" id="er3m"></div></div></div>
      <div class="vr-box" id="vrBox"><h4>&#129302; AI 검증 결과 미리보기</h4><pre id="vrText"></pre></div>
    </div>

    <!-- 탭④ 개발자 -->
    <div class="tab-content" id="tab3">
      <div class="dev-profile">
        <div class="dev-avatar">&#127970;</div>
        <div class="dev-info">
          <h2>Easydsd 0.4v</h2>
          <div class="dev-sub">DART 감사보고서 DSD 파일 변환 도구(양방향) + Gemini AI</div>
          <div class="dev-badges">
            <span class="badge bg">v0.4</span>
            <span class="badge bg-gold">&#129302; AI-Powered</span>
            <span class="badge bg-tech">Python+Flask</span>
            <span class="badge bg-ai">Gemini 1.5</span>
          </div>
        </div>
      </div>
      <div class="ig">
        <div class="ib"><div class="lbl">개발자 연락처</div><div class="val"><a href="mailto:eeffco11@naver.com">eeffco11@naver.com</a></div></div>
        <div class="ib"><div class="lbl">버전</div><div class="val">Easydsd 0.4v</div></div>
        <div class="ib"><div class="lbl">지원 파일</div><div class="val">.dsd / .xlsx</div></div>
        <div class="ib"><div class="lbl">AI 엔진</div><div class="val">Gemini 1.5 Flash</div></div>
      </div>
      <div class="cbox">
        <div class="ct">&#128591; 제작 크레딧</div>
        <div class="cb">이 프로그램은 전적으로<br><span class="cc">Claude (Anthropic)</span> 가 설계하고 개발했습니다.
          <div class="cn">클로드 짱짱맨</div><div class="cs">전 과정을 클로드로 다함</div>
        </div>
      </div>
      <div class="fs">
        <h3>&#10024; 주요 기능</h3>
        <div class="fi"><div class="fic">&#127974;</div><div>재무상태표·포괄손익·자본변동표·현금흐름표 전체 편집</div></div>
        <div class="fi"><div class="fic">&#128221;</div><div>주석 전체 편집 — 주주명, 지분율, 이자율, 텍스트 포함</div></div>
        <div class="fi"><div class="fic">&#129302;</div><div>Gemini AI 스마트 분류 — DSD→Excel 변환 시 시트명 자동 정렬</div></div>
        <div class="fi"><div class="fic">&#128269;</div><div>AI 교차 검증 — 재무제표 본문↔주석 금액 일치 여부 자동 확인 후 Excel 시트로 저장</div></div>
        <div class="fi"><div class="fic">&#128260;</div><div>DSD→Excel→DSD 완전한 양방향 변환, XML 유효성 자동 검증</div></div>
        <div class="fi"><div class="fic">&#128163;</div><div>하트비트 감시 — 브라우저 닫으면 서버 자동 종료 (좀비 방지)</div></div>
      </div>
    </div>

  </div>
</div>

<script>
// API Key
function loadKey(){const k=localStorage.getItem('gemini_api_key')||'';document.getElementById('apiKey').value=k;updSt(k);return k;}
function saveKey(v){localStorage.setItem('gemini_api_key',v);updSt(v);}
function updSt(v){const e=document.getElementById('apiSt');if(v&&v.length>10){e.textContent='🟢 입력됨';e.style.color='#a5d6a7';}else{e.textContent='⚪ 미입력';e.style.color='rgba(255,255,255,.6)';}}
function getKey(){return localStorage.getItem('gemini_api_key')||'';}
loadKey();

// 하트비트
setInterval(function(){fetch('/api/heartbeat',{method:'POST'}).catch(function(){});},2500);

// 파일/드래그
const F={f1:null,f2:null,f3:null,f4:null};
function sw(n){document.querySelectorAll('.tab').forEach((t,i)=>t.classList.toggle('active',i===n));document.querySelectorAll('.tab-content').forEach((t,i)=>t.classList.toggle('active',i===n));}
function sf(id,bid,dzId){const f=document.getElementById(id).files[0];if(!f)return;F[id]=f;const b=document.getElementById(bid);b.textContent='✓  '+f.name+'  ('+(f.size/1024).toFixed(0)+' KB)';b.style.display='block';document.getElementById(dzId).style.borderColor='#1F4E79';chk();}
function dov(e,id){e.preventDefault();document.getElementById(id).classList.add('drag-over')}
function dlv(id){document.getElementById(id).classList.remove('drag-over')}
function ddrop(e,fid,did){e.preventDefault();dlv(did);const dt=e.dataTransfer;if(!dt.files.length)return;const inp=document.getElementById(fid);const tr=new DataTransfer();tr.items.add(dt.files[0]);inp.files=tr.files;sf(fid,fid.replace('f','fb'),did);}
function chk(){document.getElementById('btn1').disabled=!F.f1;document.getElementById('btn2').disabled=!(F.f2&&F.f3);document.getElementById('btn3').disabled=!F.f4;}
function hide(n){['ok','er'].forEach(p=>document.getElementById(p+n).style.display='none');}

let piv=null;
function sp(n,msg,isAI){hide(n);if(n===3)document.getElementById('vrBox').style.display='none';const pw=document.getElementById('pw'+n);pw.style.display='block';document.getElementById('pt'+n).textContent=msg;document.getElementById('pf'+n).style.width='0%';let w=0;piv=setInterval(()=>{w=Math.min(w+(isAI?1:4),88);document.getElementById('pf'+n).style.width=w+'%';},isAI?400:200);}
function ep(n){clearInterval(piv);document.getElementById('pf'+n).style.width='100%';setTimeout(()=>document.getElementById('pw'+n).style.display='none',500);}
function sok(n,t,s,blob,fname){const b=document.getElementById('ok'+n);b.style.display='flex';document.getElementById('ok'+n+'t').textContent=t;document.getElementById('ok'+n+'s').textContent=s;if(blob){const dl=document.getElementById('dl'+n);dl.href=URL.createObjectURL(blob);dl.download=fname;}}
function ser(n,msg){const b=document.getElementById('er'+n);b.style.display='flex';document.getElementById('er'+n+'m').textContent=msg;}
const S1=['DSD 파일 분석 중...','테이블 파싱 중...','Excel 시트 생성 중...'];
const S1A=['DSD 분석 중...','Gemini AI 분류 중... (15~30초 소요)','AI 시트명 적용 중...'];
const S2=['매핑 구성 중...','XML 패치 적용 중...','DSD 생성 중...'];
const S3=['Excel 데이터 추출 중...','Gemini AI 교차 검증 중... (30~60초 소요)','검증 결과 시트 생성 중...'];
function anim(n,steps,isAI){let i=0;return setInterval(()=>{if(i<steps.length)document.getElementById('pt'+n).textContent=steps[i++];},isAI?5000:1000);}

async function run1(){
  if(!F.f1)return;
  document.getElementById('btn1').disabled=true;
  const useAI=document.getElementById('aiClassify').checked;
  const key=getKey();
  if(useAI&&!key){ser(1,'AI 분류를 사용하려면 Gemini API Key를 입력하세요.');document.getElementById('btn1').disabled=false;return;}
  sp(1,useAI?S1A[0]:S1[0],useAI);const iv=anim(1,useAI?S1A:S1,useAI);
  try{
    const fd=new FormData();fd.append('dsd',F.f1);fd.append('ai_classify',useAI?'1':'0');fd.append('api_key',key);
    const r=await fetch('/api/dsd2excel',{method:'POST',body:fd});
    clearInterval(iv);ep(1);
    if(!r.ok){const e=await r.json();throw new Error(e.error||'변환 실패');}
    const blob=await r.blob();const info=JSON.parse(r.headers.get('X-Info')||'{}');
    const fname=F.f1.name.replace(/\.dsd$/i,'')+'.xlsx';
    sok(1,'변환 완료! Excel 파일을 다운로드하세요','시트 '+info.sheets+'개 · 수정가능 셀 '+info.cells+'개 · 핵심재무표 '+info.fin+'개'+(useAI?' · 🤖AI분류 적용':''),blob,fname);
  }catch(e){clearInterval(iv);ep(1);ser(1,e.message);}
  document.getElementById('btn1').disabled=false;
}

async function run2(){
  if(!F.f2||!F.f3)return;
  document.getElementById('btn2').disabled=true;
  sp(2,S2[0]);const iv=anim(2,S2);
  try{
    const fd=new FormData();fd.append('orig_dsd',F.f2);fd.append('xlsx',F.f3);
    const r=await fetch('/api/excel2dsd',{method:'POST',body:fd});
    clearInterval(iv);ep(2);
    if(!r.ok){const e=await r.json();throw new Error(e.error||'변환 실패');}
    const blob=await r.blob();const info=JSON.parse(r.headers.get('X-Info')||'{}');
    const fname=F.f2.name.replace(/\.dsd$/i,'')+'_수정.dsd';
    sok(2,'DSD 파일 생성 완료!',info.tables+'개 테이블 · '+info.cells+'개 셀 수정 · XML 검증 '+(info.xml_ok?'✓ 정상':'✗ 오류'),blob,fname);
  }catch(e){clearInterval(iv);ep(2);ser(2,e.message);}
  document.getElementById('btn2').disabled=false;
}

async function run3(){
  if(!F.f4)return;
  const key=getKey();
  if(!key){ser(3,'Gemini API Key를 상단에 입력해주세요.');return;}
  document.getElementById('btn3').disabled=true;
  sp(3,S3[0],true);const iv=anim(3,S3,true);
  try{
    const fd=new FormData();fd.append('xlsx',F.f4);fd.append('api_key',key);
    const r=await fetch('/api/verify_excel',{method:'POST',body:fd});
    clearInterval(iv);ep(3);
    if(!r.ok){const e=await r.json();throw new Error(e.error||'검증 실패');}
    const blob=await r.blob();const info=JSON.parse(r.headers.get('X-Info')||'{}');
    const fname=F.f4.name.replace(/\.xlsx$/i,'')+'_AI검증.xlsx';
    sok(3,'AI 검증 완료! 결과를 다운로드하세요','재무시트 '+info.fin_sheets+'개 · 주석시트 '+info.note_sheets+'개 분석',blob,fname);
    if(info.preview){document.getElementById('vrText').textContent=info.preview;document.getElementById('vrBox').style.display='block';}
  }catch(e){clearInterval(iv);ep(3);ser(3,e.message);}
  document.getElementById('btn3').disabled=false;
}

function showKill(){document.getElementById('killModal').classList.add('show')}
function hideKill(){document.getElementById('killModal').classList.remove('show')}
async function doKill(){hideKill();try{await fetch('/api/shutdown',{method:'POST'});}catch(e){}document.body.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:sans-serif;color:#556;font-size:15px;">서버가 종료되었습니다. 이 탭을 닫으세요.</div>';}
</script>
</body>
</html>'''

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
        api_key     = request.form.get('api_key','').strip()
        ai_mapping  = {}
        if ai_classify:
            if not api_key:
                return jsonify(error='AI 분류를 사용하려면 Gemini API Key를 입력해주세요.'), 400
            if not GENAI_AVAILABLE:
                return jsonify(error='google-generativeai 라이브러리가 설치되지 않았습니다.\n실행.bat을 닫고 cmd에서 "pip install google-generativeai" 실행 후 다시 시도해주세요.'), 500
            xml=zipfile.ZipFile(io.BytesIO(dsd_bytes)).read('contents.xml').decode('utf-8',errors='replace')
            _,tables=parse_xml(xml)
            ai_mapping=gemini_classify_tables(api_key,tables)
        xlsx=dsd_to_excel_bytes(dsd_bytes,ai_mapping or None)
        wb=openpyxl.load_workbook(io.BytesIO(xlsx),data_only=True)
        cells=sum(1 for ws in wb.worksheets for row in ws.iter_rows()
                  for cell in row if cell.fill and cell.fill.fill_type=='solid'
                  and cell.fill.fgColor and cell.fill.fgColor.type=='rgb'
                  and cell.fill.fgColor.rgb.upper().endswith(EDIT_COLOR.upper()))
        fin=[ws.title for ws in wb.worksheets if any(ws.title.startswith(e) for e in('🏦','💹','📈','💰'))]
        resp=send_file(io.BytesIO(xlsx),
                       mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       as_attachment=True,download_name='converted.xlsx')
        resp.headers['X-Info']=json.dumps({'sheets':len(wb.sheetnames),'cells':cells,'fin':len(fin),'ai':bool(ai_mapping)})
        return resp
    except Exception as e: return jsonify(error=str(e)),500

@app.route('/api/excel2dsd', methods=['POST'])
def api_excel2dsd():
    try:
        orig=request.files['orig_dsd'].read(); xlsx=request.files['xlsx'].read()
        dsd=excel_to_dsd_bytes(orig,xlsx)
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(io.BytesIO(dsd)) as z: xt=z.read('contents.xml').decode('utf-8')
        xml_ok=True
        try: ET.fromstring(xt)
        except: xml_ok=False
        wb=openpyxl.load_workbook(io.BytesIO(xlsx),data_only=True)
        tc=tb=0
        for sname in wb.sheetnames:
            if sname in('📋사용안내','_원본XML','📊요약수치'): continue
            ws=wb[sname]
            cnt=sum(1 for row in ws.iter_rows() for cell in row
                    if cell.fill and cell.fill.fill_type=='solid'
                    and cell.fill.fgColor and cell.fill.fgColor.type=='rgb'
                    and cell.fill.fgColor.rgb.upper().endswith(EDIT_COLOR.upper()))
            if cnt: tb+=1; tc+=cnt
        resp=send_file(io.BytesIO(dsd),mimetype='application/octet-stream',
                       as_attachment=True,download_name='output.dsd')
        resp.headers['X-Info']=json.dumps({'tables':tb,'cells':tc,'xml_ok':xml_ok})
        return resp
    except Exception as e: return jsonify(error=str(e)),500

@app.route('/api/verify_excel', methods=['POST'])
def api_verify_excel():
    try:
        xlsx_bytes=request.files['xlsx'].read()
        api_key=request.form.get('api_key','').strip()
        if not api_key: return jsonify(error='Gemini API Key가 필요합니다.'),400
        if not GENAI_AVAILABLE: return jsonify(error='google-generativeai 라이브러리 미설치'),500
        fin_data,note_data=extract_fin_and_notes(xlsx_bytes)
        if not fin_data: return jsonify(error='재무제표 시트(🏦💹📈💰)를 찾을 수 없습니다.'),400
        verify_result=gemini_verify_excel(api_key,fin_data,note_data)
        wb=openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
        if '🤖AI검증결과' in wb.sheetnames: del wb['🤖AI검증결과']
        ws_v=wb.create_sheet('🤖AI검증결과',0)
        ws_v.sheet_view.showGridLines=False
        tc=ws_v.cell(1,1,'🤖 Gemini AI 재무제표 교차 검증 결과')
        tc.fill=PatternFill('solid',fgColor='4A148C')
        tc.font=Font(color='FFFFFF',bold=True,size=12)
        tc.alignment=Alignment(horizontal='left',vertical='center')
        ws_v.merge_cells('A1:F1'); ws_v.row_dimensions[1].height=28
        sc=ws_v.cell(2,1,f'생성: {time.strftime("%Y-%m-%d %H:%M")}  |  재무시트: {len(fin_data)}개  |  주석시트: {len(note_data)}개')
        sc.font=Font(color='7B1FA2',size=9,italic=True); ws_v.row_dimensions[2].height=16
        COLOR_MAP={'## ✅':('E8F5E9','1B5E20'),'## ❌':('FFEBEE','B71C1C'),'## ⚠️':('FFF8E1','E65100'),'## 📋':('E3F2FD','0D47A1')}
        for ri,line in enumerate(verify_result.split('\n'),4):
            cell=ws_v.cell(ri,1,line)
            matched=next(((fg,fc) for k,(fg,fc) in COLOR_MAP.items() if line.startswith(k)),None)
            if matched:
                cell.fill=PatternFill('solid',fgColor=matched[0])
                cell.font=Font(bold=True,size=10,color=matched[1])
            else:
                cell.font=Font(size=9)
            cell.alignment=Alignment(horizontal='left',vertical='center',wrap_text=True)
            ws_v.row_dimensions[ri].height=16
        ws_v.column_dimensions['A'].width=90
        buf=io.BytesIO(); wb.save(buf)
        preview=verify_result[:600]+('...' if len(verify_result)>600 else '')
        resp=send_file(io.BytesIO(buf.getvalue()),
                       mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       as_attachment=True,download_name='verified.xlsx')
        resp.headers['X-Info']=json.dumps({'fin_sheets':len(fin_data),'note_sheets':len(note_data),'preview':preview})
        return resp
    except Exception as e: return jsonify(error=str(e)),500

def open_browser():
    time.sleep(1.5)
    webbrowser.open(f'http://127.0.0.1:{PORT}')

if __name__=='__main__':
    print('='*52)
    print('  easydsd v0.4 - DART 감사보고서 변환 + AI')
    print(f'  http://127.0.0.1:{PORT}')
    print('  종료: 브라우저 종료 버튼 or Ctrl+C')
    print('='*52)
    if not GENAI_AVAILABLE:
        print('  ⚠️  google-generativeai 미설치 — AI 기능 비활성')
    threading.Thread(target=open_browser,daemon=True).start()
    app.run(host='127.0.0.1',port=PORT,debug=False)
