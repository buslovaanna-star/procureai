import streamlit as st
import pandas as pd
import openpyxl
import math
import io
from datetime import date, timedelta
from collections import defaultdict

# ─────────────────────────────────────────────
# Конфігурація сторінки
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ProcureAI — Аналіз закупівель",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

UA_MO = {'Січень':1,'Лютий':2,'Березень':3,'Квітень':4,'Травень':5,'Червень':6,
         'Липень':7,'Серпень':8,'Вересень':9,'Жовтень':10,'Листопад':11,'Грудень':12}
BUILT_IN_K = [1.1,1.15,1.18,1.05,0.95,1.34,1.34,0.9,0.92,1.2,1.5,1.56]

# ─────────────────────────────────────────────
# Хелпери
# ─────────────────────────────────────────────
def safe_float(v):
    if v is None: return None
    try:
        f = float(str(v).replace(',','.').strip())
        return f if f > 0 else None
    except: return None

def mo_num(label):
    return UA_MO.get(str(label).split()[0], 0)

def mo_year(label):
    for p in str(label).split():
        if p.isdigit() and len(p)==4: return int(p)
    return 0

# ─────────────────────────────────────────────
# Завантаження і парсинг Excel
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_excel(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheets = {}
    for name in wb.sheetnames:
        rows = list(wb[name].iter_rows(values_only=True))
        if rows: sheets[name] = rows
    return sheets

def detect_sheets(sheets):
    """Автоматично визначити які вкладки що містять"""
    result = {}
    for name, rows in sheets.items():
        nl = name.lower()
        if not rows: continue
        cols = [str(c).lower() for c in (rows[1] if len(rows)>1 else rows[0]) if c]
        has_months = any(m.lower() in ' '.join(cols) for m in UA_MO.keys())

        if 'продаж' in nl or 'sales' in nl: result['sales'] = name
        elif 'наявн' in nl or 'avail' in nl: result['avail'] = name
        elif ('залиш' in nl or 'stock' in nl) and not has_months: result['stock'] = name
        elif 'ціни' in nl or 'price' in nl or 'артикул' in ' '.join(cols[:3]): result['prices'] = name

    # Fallback: якщо не знайшли — берем по вмісту
    for name, rows in sheets.items():
        if name in result.values(): continue
        if len(rows) < 2: continue
        cols = [str(c).lower() for c in rows[1] if c]
        has_months = any(m.lower() in ' '.join(cols) for m in UA_MO.keys())
        has_pairs  = any('рентаб' in c or 'margin' in c for c in cols)
        if has_months and has_pairs and 'sales' not in result: result['sales'] = name
        elif has_months and not has_pairs and 'avail' not in result: result['avail'] = name
    return result

def parse_months(rows):
    """Знайти місяці у рядку заголовків (рядок 2, індекс 1)"""
    month_row = rows[1] if len(rows)>1 else rows[0]
    months = []
    for ci, v in enumerate(month_row):
        if ci >= 3 and v is not None and mo_num(str(v)) > 0:
            months.append((ci, str(v)))
    return months

# ─────────────────────────────────────────────
# Основний розрахунок
# ─────────────────────────────────────────────
def run_analysis(sheets, sheet_map, params):
    today    = date.today()
    CUR_MO   = today.month
    CUR_YR   = today.year
    CUR_DAY  = today.day
    cur_scale = 30.0 / CUR_DAY

    MG_MIN   = params['mg_min']
    LEAD     = params['lead']
    SAFETY   = params['safety']
    SAFETY60 = params['safety60']
    A_MULT   = params['a_mult']
    MG_A     = params['mg_a']
    MIN_MO   = params['min_months']
    MIN_QTY  = params['min_qty']
    AA       = params['avail_alpha']
    LAM      = params['lambda_val']
    LOW_AV   = params['low_avail']
    DISC_THR = params['disc_thr']

    # ── Продажі ──
    sales_rows = sheets[sheet_map['sales']]
    months = parse_months(sales_rows)
    n = len(months)
    if n == 0:
        return None, "Не знайдено місяців у вкладці Продажі"

    mo_complete = []
    mo_is_cur   = []
    for _, lbl in months:
        mn, yr = mo_num(lbl), mo_year(lbl)
        done = yr < CUR_YR or (yr == CUR_YR and mn < CUR_MO)
        cur  = mn == CUR_MO and yr == CUR_YR
        mo_complete.append(done)
        mo_is_cur.append(cur)

    exp_w = [math.exp(LAM * (i - (n-1))) for i in range(n)]

    # ── Наявність ──
    avail_map = {}
    if 'avail' in sheet_map:
        for r in sheets[sheet_map['avail']][2:]:
            if not r[0] or not str(r[0]).strip(): continue
            sku = str(r[0]).strip(); days = []
            for v in r[1:n+1]:
                dv = float(v) if v is not None else 0.0
                days.append(dv*30 if 0 < dv <= 1 else dv)
            avail_map[sku] = days

    # ── Залишки ──
    stock_map = {}
    if 'stock' in sheet_map:
        for r in sheets[sheet_map['stock']][2:]:
            if not r[0] or not str(r[0]).strip() or str(r[0])=='1*': continue
            sku = str(r[0]).strip()
            nm  = str(r[1]) if len(r)>1 and r[1] else ''
            st  = max(float(r[3]),0) if len(r)>3 and r[3] is not None else 0
            tr  = max(float(r[4]),0) if len(r)>4 and r[4] is not None else 0
            stock_map[sku] = (st, tr, nm)

    # ── Ціни ──
    price_map = {}
    if 'prices' in sheet_map:
        for r in sheets[sheet_map['prices']][3:]:
            if not r[0] or not str(r[0]).strip(): continue
            sku = str(r[0]).strip()
            price = safe_float(r[2]) if len(r)>2 else None
            avail_str = str(r[4]).lower() if len(r)>4 and r[4] else ''
            disc_raw  = r[5] if len(r)>5 else None
            sold30    = safe_float(r[8]) if len(r)>8 else None
            in_stock  = 'наявн' in avail_str or avail_str in ('1','true')
            disc = safe_float(disc_raw) or 0
            price_map[sku] = (price, in_stock, disc, sold30)

    # ── Pass 1: season K ──
    mo_cs = [0.0]*n; mo_cc = [0]*n
    sku_md = {}

    data_rows = [r for r in sales_rows[3:]
                 if r[0] and str(r[0]).strip() and not str(r[0]).startswith('⚠')]

    for r in data_rows:
        sku  = str(r[0]).strip()
        name = str(r[1]) if r[1] else ''
        av   = avail_map.get(sku, [30]*n)
        md   = []
        for i,(ci,label) in enumerate(months):
            qty  = float(r[ci])   if r[ci]   is not None else 0.0
            rent = float(r[ci+1]) if len(r)>ci+1 and r[ci+1] is not None else None
            if rent is not None and 0 < rent < 1: rent *= 100
            ad   = av[i] if i < len(av) else 30
            inc  = (rent >= MG_MIN) if rent is not None else (qty > 0)
            if inc and qty > 0 and mo_complete[i]:
                mo_cs[i] += qty; mo_cc[i] += 1
            md.append({'qty':qty,'rent':rent,'avail':ad,'include':inc,
                       'is_cur':mo_is_cur[i],'complete':mo_complete[i]})
        sku_md[sku] = (name, md)

    ca_ = [mo_cs[i]/mo_cc[i] if mo_cc[i]>0 else 0 for i in range(n)]
    cv  = [ca_[i] for i in range(n) if mo_complete[i] and ca_[i]>0]
    global_avg = sum(cv)/len(cv) if cv else 1.0
    cur_ks = [ca_[i]/global_avg for i,(_, lbl) in enumerate(months)
              if mo_complete[i] and mo_num(lbl)==CUR_MO and ca_[i]>0]
    season_K = sum(cur_ks)/len(cur_ks) if cur_ks else BUILT_IN_K[CUR_MO-1]

    # ── Pass 2: per-SKU ──
    results = []; excl_mg = []; sporadic = []

    for r in data_rows:
        sku = str(r[0]).strip()
        name, md = sku_md[sku]
        st_info = stock_map.get(sku)
        stock   = st_info[0] if st_info else 0
        transit = st_info[1] if st_info else 0
        nm      = st_info[2] if st_info else name
        if not nm: nm = name
        eff = stock + transit

        # avg/day зважений
        ws = 0.0; wd = 0.0; rd = n*30
        for i, m in enumerate(md):
            if not m['include']: continue
            wt  = exp_w[i]
            qty = m['qty'] * (cur_scale if m['is_cur'] else 1)
            ad  = min(m['avail'] * (cur_scale if m['is_cur'] else 1), 30)
            if ad > 0: ws += qty*wt; wd += ad*wt

        avg_day  = ws/wd if wd>0 else 0
        cd_clean = sum(min(m['avail']*(cur_scale if m['is_cur'] else 1),30)
                       for m in md if m['include'])
        avail_pct = round(cd_clean/rd*100) if rd>0 else 0
        avail_K   = math.pow(avail_pct/100, AA) if avail_pct>0 else 0

        rents = [m['rent'] for m in md if m['include'] and m['rent'] is not None]
        avg_margin = round(sum(rents)/len(rents),1) if rents else None

        # Фільтр маржі
        if avg_margin is not None and avg_margin < MG_MIN:
            cmo = sum(1 for m in md if m['include'] and m['qty']>0)
            ts  = sum(m['qty'] for m in md if m['include'])
            excl_mg.append({'sku':sku,'name':nm,'avg_margin':avg_margin,
                             'clean_months':cmo,'total_sold':round(ts,1)})
            continue

        cmo = sum(1 for m in md if m['include'] and m['qty']>0)
        ts  = sum(m['qty'] for m in md if m['include'])
        is_sp = not (cmo >= MIN_MO and ts >= MIN_QTY)
        reason = ''
        if is_sp:
            parts = []
            if cmo < MIN_MO:  parts.append(f"{cmo}<{MIN_MO}міс")
            if ts  < MIN_QTY: parts.append(f"{round(ts):.0f}<{MIN_QTY}шт")
            reason = ', '.join(parts)

        if avg_margin is None:   abc, am = '?', 1.0
        elif avg_margin >= MG_A: abc, am = 'A', A_MULT
        else:                    abc, am = 'B', 1.0

        # Тренд
        cm_list = [m for m in md if m['include'] and m['avail']>0]
        if len(cm_list) >= 6:
            f3 = sum(m['qty'] for m in cm_list[:3])/3
            l3 = sum(m['qty'] for m in cm_list[-3:])/3
            tr_r = round(l3/f3,2) if f3>0 else None
            trend = ('↑ зростає' if tr_r and tr_r>1.3 else
                     '↓ спадає'  if tr_r and tr_r<0.7 else '→ стабільно')
        else:
            l3 = sum(m['qty'] for m in cm_list[-3:])/3 if cm_list else 0
            trend = '↑ новий' if l3>0 else '—'

        # Ціни
        pi = price_map.get(sku)
        p_avail = pi[1] if pi else True
        p_disc  = pi[2] if pi else 0
        p_price = pi[0] if pi else None
        use_60  = not is_sp and p_avail and p_disc >= DISC_THR

        # Маржинальний дохід
        mi_day = None
        sell_price = None
        if p_price and avg_margin and 0 < avg_margin < 100:
            sell_price = p_price / (1 - avg_margin/100)
            mi_day = round(avg_day * (sell_price - p_price), 4)

        dl = round(stock/avg_day) if avg_day>0 else 999
        st = ('Критично' if dl<5 else 'Низько' if dl<15
              else 'Надлишок' if dl>90 else 'Норма')

        rec   = max(0,round(avg_day*avail_K*am*(LEAD+SAFETY  )*season_K-eff)) if not is_sp and p_avail else 0
        rec60 = max(0,round(avg_day*avail_K*am*(LEAD+SAFETY60)*season_K-eff)) if not is_sp and p_avail else 0
        zero_date = (today+timedelta(days=int(dl))).strftime('%d.%m.%Y') if 0<=dl<999 else None

        row = dict(
            sku=sku, name=nm, abc=abc, avg_margin=avg_margin,
            avail_pct=avail_pct, avail_K=round(avail_K,3),
            trend=trend, avg_day=round(avg_day,4),
            stock=stock, transit=transit, eff_stock=eff,
            days_left=dl, zero_date=zero_date, status=st,
            is_sporadic=is_sp, sporadic_reason=reason,
            season_K=round(season_K,3), rec=rec, rec_60=rec60,
            use_60=use_60, price_disc=round(p_disc,1),
            buy_price=p_price, sell_price=round(sell_price,2) if sell_price else None,
            mi_day=mi_day, low_avail=avail_pct<LOW_AV,
        )
        if is_sp: sporadic.append(row)
        else:     results.append(row)

    # ABC по MI (Pareto 70/90)
    mi_vals = sorted([r['mi_day'] for r in results if r['mi_day']], reverse=True)
    total_mi = sum(mi_vals)
    cum=0; ta=tb=None
    for v in mi_vals:
        cum+=v
        if ta is None and cum>=total_mi*0.70: ta=v
        if tb is None and cum>=total_mi*0.90: tb=v
    for r in results:
        mi=r['mi_day']
        if mi is None: r['abc_mi']='?'
        elif ta and mi>=ta: r['abc_mi']='A'
        elif tb and mi>=tb: r['abc_mi']='B'
        else: r['abc_mi']='C'

    meta = dict(season_K=season_K, global_avg=global_avg, n_months=n,
                cur_day=CUR_DAY, cur_scale=cur_scale,
                total_mi=total_mi, ta=ta, tb=tb,
                months=[lbl for _,lbl in months])
    return dict(regular=results, sporadic=sporadic, excl_mg=excl_mg, meta=meta), None

# ─────────────────────────────────────────────
# Генерація Excel-звіту
# ─────────────────────────────────────────────
def gen_excel(data, params):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    results  = data['regular']
    sporadic = data['sporadic']
    excl_mg  = data['excl_mg']
    meta     = data['meta']
    today    = date.today()

    th = Side(style="thin", color="BFBFBF")
    def tb(): return Border(left=th,right=th,top=th,bottom=th)
    def fl(c): return PatternFill("solid",start_color=c,fgColor=c)
    def hf(color="FFFFFF",bold=True,sz=10): return Font(name="Arial",bold=bold,color=color,size=sz)
    def cf(bold=False,sz=10,color="000000"): return Font(name="Arial",bold=bold,size=sz,color=color)
    def ca(): return Alignment(horizontal="center",vertical="center",wrap_text=True)
    def la(): return Alignment(horizontal="left",vertical="center")

    C=dict(main="1F4E79",red="C62828",amber="E65100",green="1B5E20",blue="0D47A1",
           gray="546E7A",orange="BF360C",teal="006064",gold="F57F17",
           red_l="FFEBEE",amber_l="FFF3E0",green_l="E8F5E9",blue_l="E3F2FD",
           gray_l="ECEFF1",orange_l="FBE9E7",gold_l="FFFDE7",low_l="FCE4EC",
           row1="EBF3FB",row2="F2F9EC",white="FFFFFF",yellow="FFF9C4")
    ST={'Критично':(C['red'],C['red_l']),'Низько':(C['amber'],C['amber_l']),
        'Норма':(C['green'],C['green_l']),'Надлишок':(C['blue'],C['blue_l'])}
    AB={'A':C['green'],'B':C['blue'],'?':C['gray']}
    AB_MI={'A':C['gold'],'B':C['blue'],'C':C['gray'],'?':C['gray']}
    TR={'↑ зростає':C['green'],'↑ новий':C['teal'],'↓ спадає':C['red'],
        '→ стабільно':C['gray'],'—':C['gray']}

    wb = Workbook()
    param_str=(f"Маржа≥{params['mg_min']}% | Lead {params['lead']}д | "
               f"Safety {params['safety']}д | α={params['avail_alpha']} | "
               f"λ={params['lambda_val']} | K={meta['season_K']:.3f}")

    def make_ws(ws, title, warn, cols, warn_c="7F3F00", warn_bg="FFF9C4"):
        lc=get_column_letter(len(cols))
        ws.merge_cells(f"A1:{lc}1"); ws["A1"]=title
        ws["A1"].font=hf(sz=9); ws["A1"].fill=fl(C['main'])
        ws["A1"].alignment=la(); ws.row_dimensions[1].height=20
        ws.merge_cells(f"A2:{lc}2"); ws["A2"]=warn
        ws["A2"].font=Font(name="Arial",size=9,italic=True,color=warn_c)
        ws["A2"].fill=fl(warn_bg); ws["A2"].alignment=la(); ws.row_dimensions[2].height=18
        for col,(h,w) in enumerate(cols,1):
            c=ws.cell(row=3,column=col,value=h)
            c.font=hf(sz=9); c.fill=fl(C['main']); c.alignment=ca(); c.border=tb()
            ws.column_dimensions[get_column_letter(col)].width=w
        ws.row_dimensions[3].height=40

    def fill_row(ws, row, r, bg, cols_def):
        dl=r.get('days_left',999); low=r.get('low_avail',False)
        tr=r.get('trend','—'); abc_mi=r.get('abc_mi','?')
        fc2,_=ST.get(r.get('status','Норма'),(C['gray'],C['gray_l']))
        for col,(key,fmt) in enumerate(cols_def,1):
            val=r.get(key)
            c=ws.cell(row=row,column=col,value=val); c.fill=fl(bg); c.border=tb()
            if fmt=='sku':
                c.value=("⚑ " if low else "")+str(val or '')
                c.font=Font(name="Arial",bold=True,size=10,color="880E4F" if low else "000000"); c.alignment=la()
            elif fmt=='name': c.font=cf(sz=9); c.alignment=la()
            elif fmt=='abc':  c.font=Font(name="Arial",bold=True,size=11,color=AB.get(val,C['gray'])); c.alignment=ca()
            elif fmt=='abc_mi': c.font=Font(name="Arial",bold=True,size=11,color=AB_MI.get(val,C['gray'])); c.alignment=ca()
            elif fmt=='pct':
                if val: c.number_format='0.0"%"'
                mg_c=(C['green'] if (val or 0)>=params['mg_a'] else C['blue'])
                c.font=Font(name="Arial",bold=True,size=10,color=mg_c); c.alignment=ca()
            elif fmt=='avail_pct':
                if val: c.number_format='0"%"'
                bc=C['red'] if (val or 0)<params['low_avail'] else C['amber'] if (val or 0)<50 else C['green']
                c.font=Font(name="Arial",bold=low,size=10,color=bc); c.alignment=ca()
            elif fmt=='mi':
                if val: c.number_format='#,##0.00'
                ta_=meta.get('ta') or 0; tb__=meta.get('tb') or 0
                col_c=(C['gold'] if (val or 0)>=ta_ else C['blue'] if (val or 0)>=tb__ else C['gray'])
                c.font=Font(name="Arial",bold=True,size=10,color=col_c); c.alignment=ca()
            elif fmt=='days':
                c.font=Font(name="Arial",bold=True,size=11,color=fc2); c.alignment=ca()
            elif fmt=='date':
                col_c=C['red'] if dl<5 else C['amber'] if dl<15 else "000000"
                if val: c.font=Font(name="Arial",bold=True,size=10,color=col_c)
                else: c.value="∞"; c.font=cf(sz=10,color="AAAAAA")
                c.alignment=ca()
            elif fmt=='trend':
                c.font=Font(name="Arial",bold=True,size=10,color=TR.get(val,C['gray'])); c.alignment=ca()
            elif fmt=='rec':
                if val and val>0: c.font=Font(name="Arial",bold=True,size=12,color=C['main'])
                else: c.value="—"; c.font=cf(sz=10,color="BBBBBB")
                c.alignment=ca()
            elif fmt=='rec60':
                if val and val>0:
                    c.font=Font(name="Arial",bold=True,size=12,color=C['amber']); c.fill=fl(C['amber_l'])
                else: c.value="—"; c.font=cf(sz=10,color="BBBBBB")
                c.alignment=ca()
            elif fmt=='avg':
                c.number_format='0.0000'; c.font=Font(name="Arial",bold=True,size=10,color=C['blue']); c.alignment=ca()
            else:
                c.font=cf(sz=10); c.alignment=ca()

    # ── Sheet 1: Замовлення ──
    ws1=wb.active; ws1.title="Замовлення"
    COLS_O=[("SKU",13),("Назва",42),("ABC\nмарж%",7),("ABC\nMI",7),
            ("Маржа %",10),("MI/день",11),("Залишок\n+транзит",10),
            ("Дата нуля",11),("Дні до нуля",9),("Тренд",11),
            ("Замовити\n14+30д",12),("Замовити\n14+60д\nзнижка",13)]
    make_ws(ws1,f"ЗАМОВЛЕННЯ | {today.strftime('%d.%m.%Y')} | {param_str}",
            f"ABC_MI (золото/синій/сірий) = клас за MI/день. Дата нуля = коли закінчиться залишок. ⚑ = наявність <{params['low_avail']}%.",
            COLS_O)
    order=[r for r in results if r['rec']>0 or r['rec_60']>0]
    order.sort(key=lambda x:x['days_left'])
    DEFS_O=[('sku','sku'),('name','name'),('abc','abc'),('abc_mi','abc_mi'),
            ('avg_margin','pct'),('mi_day','mi'),('eff_stock','num'),
            ('zero_date','date'),('days_left','days'),('trend','trend'),
            ('rec','rec'),('rec_60','rec60')]
    for ri,r in enumerate(order):
        dl=r['days_left']; low=r.get('low_avail',False)
        _,bg=ST.get(r['status'],(C['gray'],C['gray_l']))
        if low: bg=C['low_l']
        elif r['status']=='Норма': bg=C['row1'] if ri%2==0 else C['row2']
        fill_row(ws1,ri+4,r,bg,DEFS_O)
    ws1.freeze_panes="A4"; ws1.auto_filter.ref=f"A3:{get_column_letter(len(COLS_O))}{len(order)+3}"

    # ── Sheet 2: Аналіз SKU ──
    ws2=wb.create_sheet("Аналіз SKU")
    COLS_A=[("SKU",13),("Назва",42),("ABC\nмарж%",7),("ABC\nMI",7),
            ("Маржа %",10),("MI/день",11),("Наявн %",8),("Тренд",10),
            ("avg/день",10),("Залишок",8),("Дата нуля",11),("Дні до нуля",9),
            ("Замовити\n14+30д",12),("Замовити\n14+60д",12)]
    make_ws(ws2,f"Аналіз SKU — {len(results)} регулярних | {today.strftime('%d.%m.%Y')} | {param_str}",
            f"ABC_MI: A=топ 70% MI (золото), B=70-90% (синій), C=нижні 10% (сірий). ⚑ = наявність <{params['low_avail']}%.",
            COLS_A)
    DEFS_A=[('sku','sku'),('name','name'),('abc','abc'),('abc_mi','abc_mi'),
            ('avg_margin','pct'),('mi_day','mi'),('avail_pct','avail_pct'),
            ('trend','trend'),('avg_day','avg'),('stock','num'),
            ('zero_date','date'),('days_left','days'),('rec','rec'),('rec_60','rec60')]
    for ri,r in enumerate(sorted(results,key=lambda x:x['days_left'])):
        dl=r['days_left']; low=r.get('low_avail',False)
        _,bg=ST.get(r['status'],(C['gray'],C['gray_l']))
        if low: bg=C['low_l']
        elif r['status']=='Норма': bg=C['row1'] if ri%2==0 else C['row2']
        fill_row(ws2,ri+4,r,bg,DEFS_A)
    ws2.freeze_panes="A4"; ws2.auto_filter.ref=f"A3:{get_column_letter(len(COLS_A))}{len(results)+3}"

    # ── Sheet 3: ABC по MI ──
    ws3=wb.create_sheet("ABC по MI")
    COLS_M=[("SKU",13),("Назва",42),("ABC\nмарж%",7),("ABC\nMI",7),
            ("Маржа %",10),("MI/день",11),("avg/день",10),
            ("Дата нуля",11),("Дні до нуля",9),("Замовити",10)]
    mi_sorted=[r for r in results if r.get('mi_day')]
    mi_sorted.sort(key=lambda x:-(x['mi_day'] or 0))
    make_ws(ws3,f"ABC по маржинальному доходу | {today.strftime('%d.%m.%Y')}",
            f"Відсортовано за MI/день ↓. A={sum(1 for r in results if r.get('abc_mi')=='A')} "
            f"B={sum(1 for r in results if r.get('abc_mi')=='B')} "
            f"C={sum(1 for r in results if r.get('abc_mi')=='C')} SKU.",
            COLS_M)
    DEFS_M=[('sku','sku'),('name','name'),('abc','abc'),('abc_mi','abc_mi'),
            ('avg_margin','pct'),('mi_day','mi'),('avg_day','avg'),
            ('zero_date','date'),('days_left','days'),('rec','rec')]
    for ri,r in enumerate(mi_sorted):
        abc_mi=r.get('abc_mi','?')
        bg=(C['gold_l'] if abc_mi=='A' else C['blue_l'] if abc_mi=='B' else C['gray_l'])
        if ri%2!=0: bg=(C['row1'] if abc_mi=='A' else C['row2'] if abc_mi=='B' else C['white'])
        fill_row(ws3,ri+4,r,bg,DEFS_M)
    ws3.freeze_panes="A4"

    # ── Sheet 4: Разовий попит ──
    ws4=wb.create_sheet("Разовий попит")
    COLS_SP=[("SKU",13),("Назва",42),("Маржа %",10),
             ("Чист. міс.",12),("Продано (шт)",13),("Причина",28)]
    make_ws(ws4,f"Разовий попит — {len(sporadic)} SKU",
            f"Виключені з замовлень: < {params['min_months']} міс. АБО < {params['min_qty']} шт.",
            COLS_SP,"BF360C","FBE9E7")
    sp_sorted=sorted(sporadic,key=lambda x:-(x.get('avg_margin') or 0))[:200]
    for ri,r in enumerate(sp_sorted):
        row=ri+4; bg=C['orange_l'] if ri%2==0 else C['white']
        for col,val in enumerate([r['sku'],r['name'],r.get('avg_margin'),
                                   r.get('clean_months',0),r.get('total_clean_sold',0),
                                   r.get('sporadic_reason','')],1):
            c=ws4.cell(row=row,column=col,value=val); c.fill=fl(bg); c.border=tb()
            if col==1: c.font=cf(bold=True,sz=10); c.alignment=la()
            elif col==2: c.font=cf(sz=9); c.alignment=la()
            elif col==3:
                if val: c.number_format='0.0"%"'
                c.font=cf(sz=10,color=C['blue']); c.alignment=ca()
            elif col==6: c.font=Font(name="Arial",size=10,italic=True,color=C['orange']); c.alignment=la()
            else: c.font=cf(sz=10); c.alignment=ca()

    # ── Sheet 5: Ручне рішення ──
    ws5=wb.create_sheet("Ручне рішення")
    low_rows=sorted([r for r in results if r.get('low_avail')],key=lambda x:x['avail_pct'])
    COLS_L=[("SKU",13),("Назва",42),("ABC\nмарж%",7),("ABC\nMI",7),
            ("Маржа %",10),("Наявн %",9),("Тренд",11),("Дата нуля",11)]
    make_ws(ws5,f"Ручне рішення — наявність <{params['low_avail']}% | {len(low_rows)} SKU",
            "Прогноз ненадійний. Вирішіть вручну: замовити під попит або пропустити.",
            COLS_L,"880E4F","FCE4EC")
    DEFS_L=[('sku','sku'),('name','name'),('abc','abc'),('abc_mi','abc_mi'),
            ('avg_margin','pct'),('avail_pct','avail_pct'),('trend','trend'),('zero_date','date')]
    for ri,r in enumerate(low_rows):
        bg="FCE4EC" if ri%2==0 else C['white']
        fill_row(ws5,ri+4,r,bg,DEFS_L)

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────
st.title("📦 ProcureAI — Аналіз закупівель")
st.caption("Завантажте Excel-файл → налаштуйте параметри → отримайте рекомендації")

# ── Sidebar: параметри ──
with st.sidebar:
    st.header("⚙️ Параметри")

    with st.expander("💰 Маржа та ABC", expanded=True):
        mg_min  = st.slider("Мінімальна маржа (%)", 5, 50, 35)
        mg_a    = st.slider("A-клас: маржа ≥ (%)", mg_min, 70, max(mg_min+10, 50))
        a_mult  = st.slider("Бонус запасу A-класу (×)", 1.0, 2.0, 1.3, 0.05)

    with st.expander("📅 Замовлення", expanded=True):
        lead    = st.slider("Lead time (дні)", 1, 30, 14)
        safety  = st.slider("Страховий запас (дні)", 7, 60, 30)
        safety60= st.slider("Safety при знижці (дні)", 30, 90, 60)
        disc_thr= st.slider("Поріг знижки постачальника (%)", 5, 30, 10)

    with st.expander("📊 Попит", expanded=True):
        min_months= st.slider("Мін. місяців з продажами", 2, 9, 6)
        min_qty   = st.slider("Мін. продано штук (за весь період)", 1, 30, 12)
        low_avail = st.slider("Поріг низької наявності (%)", 10, 40, 20)

    with st.expander("🔢 Коефіцієнти", expanded=False):
        avail_alpha = st.slider("Коефіцієнт наявності α", 0.3, 1.5, 0.7, 0.05,
                                help="avail_K = (avail%)^α. Менше α = слабший вплив")
        lambda_val  = st.slider("Часові ваги λ", 0.05, 0.5, 0.25, 0.05,
                                help="Більше λ = сильніший акцент на останні місяці")

params = dict(mg_min=mg_min, mg_a=mg_a, a_mult=a_mult,
              lead=lead, safety=safety, safety60=safety60, disc_thr=disc_thr,
              min_months=min_months, min_qty=min_qty, low_avail=low_avail,
              avail_alpha=avail_alpha, lambda_val=lambda_val)

# ── Завантаження файлу ──
uploaded = st.file_uploader(
    "📂 Завантажте Excel-файл (вкладки: Продажі, Наявність, Залишки, Ціни постачальника)",
    type=["xlsx","xls"], help="Файл обробляється локально у браузері — дані не зберігаються")

if uploaded:
    with st.spinner("Читаємо файл..."):
        sheets = load_excel(uploaded.read())

    sheet_map = detect_sheets(sheets)

    # Показати які вкладки знайдено
    with st.expander("🔍 Знайдені вкладки", expanded=False):
        roles = {'sales':'Продажі','avail':'Наявність','stock':'Залишки','prices':'Ціни'}
        for role,label in roles.items():
            found = sheet_map.get(role,'—')
            icon = "✅" if role in sheet_map else "⚠️"
            st.write(f"{icon} **{label}**: `{found}`")
        if 'sales' not in sheet_map:
            st.error("Не знайдено вкладку з продажами. Перевірте назву.")
            st.stop()

    # ── Запуск аналізу ──
    with st.spinner(f"Аналізуємо... Це може зайняти до 30 секунд для великого файлу"):
        data, err = run_analysis(sheets, sheet_map, params)

    if err:
        st.error(f"Помилка: {err}")
        st.stop()

    meta     = data['meta']
    regular  = data['regular']
    sporadic = data['sporadic']
    excl_mg  = data['excl_mg']
    total    = len(regular)+len(sporadic)+len(excl_mg)

    # ── KPI метрики ──
    st.subheader("📊 Зведення")
    c1,c2,c3,c4,c5,c6 = st.columns(6)
    c1.metric("Всього SKU", total)
    c2.metric("Постійний попит", len(regular))
    c3.metric("Разовий попит", len(sporadic))
    c4.metric("Критично (<5дн)", sum(1 for r in regular if r['status']=='Критично'),
              delta=None, delta_color="inverse")
    c5.metric("До замовлення", sum(1 for r in regular if r['rec']>0))
    c6.metric("Season K", f"{meta['season_K']:.3f}")

    # Дата нуля
    st.subheader("📅 Коли закінчуються залишки")
    buckets=[("🔴 Вже 0",0,1),("🔴 До 7 дн",1,8),("🟡 8–14 дн",8,15),
             ("🟡 15–30 дн",15,31),("🟢 31–60 дн",31,61),("🔵 >60 дн",61,999)]
    cols_b = st.columns(len(buckets))
    for i,(lbl,lo,hi) in enumerate(buckets):
        cnt=sum(1 for r in regular if lo<=r['days_left']<hi)
        cols_b[i].metric(lbl, cnt)

    # ── Таблиці ──
    tab1,tab2,tab3,tab4,tab5 = st.tabs([
        "🛒 Замовлення","📋 Аналіз SKU","⭐ ABC по MI",
        "⚠️ Ручне рішення","📦 Разовий попит"])

    with tab1:
        order=[r for r in regular if r['rec']>0 or r['rec_60']>0]
        order.sort(key=lambda x:x['days_left'])
        if order:
            df=pd.DataFrame([{
                'SKU':       ('⚑ ' if r['low_avail'] else '')+r['sku'],
                'Назва':     r['name'][:50],
                'ABC':       r['abc'],
                'ABC_MI':    r.get('abc_mi','?'),
                'Маржа %':   r['avg_margin'],
                'MI/день':   r.get('mi_day'),
                'Залишок':   r['eff_stock'],
                'Дата нуля': r.get('zero_date','∞'),
                'Дні до 0':  r['days_left'] if r['days_left']<999 else '∞',
                'Тренд':     r['trend'],
                'Замовити 14+30':  r['rec'],
                'Замовити 14+60':  r['rec_60'],
            } for r in order])
            st.dataframe(df, use_container_width=True, height=500)
            st.caption(f"Всього: {len(order)} SKU | {sum(r['rec'] for r in order)} шт (стандарт) | {sum(r['rec_60'] for r in order)} шт (при знижці)")
        else:
            st.success("Всі залишки в нормі — замовлення не потрібні")

    with tab2:
        df2=pd.DataFrame([{
            'SKU':       ('⚑ ' if r['low_avail'] else '')+r['sku'],
            'Назва':     r['name'][:50],
            'ABC':       r['abc'],
            'ABC_MI':    r.get('abc_mi','?'),
            'Маржа %':   r['avg_margin'],
            'MI/день':   r.get('mi_day'),
            'Наявн %':   r['avail_pct'],
            'Тренд':     r['trend'],
            'avg/день':  r['avg_day'],
            'Залишок':   r['stock'],
            'Дата нуля': r.get('zero_date','∞'),
            'Дні до 0':  r['days_left'] if r['days_left']<999 else '∞',
            'Статус':    r['status'],
            'Замовити':  r['rec'],
        } for r in sorted(regular,key=lambda x:x['days_left'])])
        st.dataframe(df2, use_container_width=True, height=500)
        col_f1,col_f2=st.columns(2)
        with col_f1:
            st.write(f"**Постійний попит:** {len(regular)} SKU")
            st.write(f"**Season K (березень):** {meta['season_K']:.3f}")
        with col_f2:
            st.write(f"**Місяців даних:** {meta['n_months']}")
            st.write(f"**Поточний місяць:** {meta['cur_day']}/30 дн (×{meta['cur_scale']:.2f})")

    with tab3:
        mi_sorted=[r for r in regular if r.get('mi_day')]
        mi_sorted.sort(key=lambda x:-(x['mi_day'] or 0))
        df3=pd.DataFrame([{
            'SKU':       r['sku'],
            'Назва':     r['name'][:50],
            'ABC_MI':    r.get('abc_mi','?'),
            'MI/день':   r.get('mi_day'),
            'Маржа %':   r['avg_margin'],
            'avg/день':  r['avg_day'],
            'Дата нуля': r.get('zero_date','∞'),
            'Замовити':  r['rec'],
        } for r in mi_sorted])
        st.dataframe(df3, use_container_width=True, height=500)
        a_cnt=sum(1 for r in regular if r.get('abc_mi')=='A')
        b_cnt=sum(1 for r in regular if r.get('abc_mi')=='B')
        c_cnt=sum(1 for r in regular if r.get('abc_mi')=='C')
        st.caption(f"A (золото, топ 70% MI): {a_cnt} | B (70-90%): {b_cnt} | C (нижні 10%): {c_cnt}")

    with tab4:
        low_rows=sorted([r for r in regular if r.get('low_avail')],key=lambda x:x['avail_pct'])
        if low_rows:
            df4=pd.DataFrame([{
                'SKU':       r['sku'],
                'Назва':     r['name'][:50],
                'ABC':       r['abc'],
                'Маржа %':   r['avg_margin'],
                'Наявн %':   r['avail_pct'],
                'Тренд':     r['trend'],
                'Залишок':   r['stock'],
                'Дата нуля': r.get('zero_date','∞'),
            } for r in low_rows])
            st.warning(f"⚑ {len(low_rows)} SKU мали наявність <{params['low_avail']}% — прогноз ненадійний")
            st.dataframe(df4, use_container_width=True, height=400)
        else:
            st.success("Немає SKU з підозріло низькою наявністю")

    with tab5:
        if sporadic:
            df5=pd.DataFrame([{
                'SKU':       r['sku'],
                'Назва':     r['name'][:50],
                'Маржа %':   r.get('avg_margin'),
                'Причина':   r.get('sporadic_reason',''),
            } for r in sorted(sporadic,key=lambda x:-(x.get('avg_margin') or 0))[:300]])
            st.info(f"🔵 {len(sporadic)} SKU з нерегулярним попитом — виключені з замовлень")
            st.dataframe(df5, use_container_width=True, height=400)
        else:
            st.success("Всі SKU мають постійний попит")

    # ── Завантаження Excel ──
    st.divider()
    st.subheader("⬇️ Вивантажити звіт")
    with st.spinner("Генеруємо Excel..."):
        excel_buf = gen_excel(data, params)
    fname = f"ProcureAI_{date.today().strftime('%d%m%Y')}.xlsx"
    st.download_button(
        label="📥 Завантажити Excel-звіт (5 вкладок)",
        data=excel_buf,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)

else:
    # Інструкція
    st.info("👆 Завантажте Excel-файл щоб почати аналіз")
    with st.expander("📖 Структура Excel-файлу"):
        st.markdown("""
**Вкладка 1 — Продажі:**
- Колонки: SKU | Назва | Категорія | Місяць_кількість | Місяць_рентаб% | ...
- Пари колонок для кожного місяця (кількість + рентабельність %)

**Вкладка 2 — Наявність:**
- SKU | Місяць1 | Місяць2 | ... — кількість днів товару на складі (0–31)

**Вкладка 3 — Залишки:**
- SKU | Назва | Категорія | Залишок (шт) | В дорозі (шт)

**Вкладка 4 — Ціни постачальника:**
- Артикул | Название | Цена | Старая цена | Наличие | Скидка | Рейтинг | Відгуки | Продано за 30 дн
        """)
