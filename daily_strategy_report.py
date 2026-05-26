"""
台指策略日報自動生成 v2.6
================================================================
每天早上 8:30 在 Jupyter 執行：

    %run daily_strategy_report.py

輸出：./reports/台股策略日報_YYYY.MM.DD.pptx

安裝套件：
    pip install yfinance python-pptx pandas numpy matplotlib requests
================================================================
v2.6 變更：
  ① 資料來源：Yahoo Finance（移除 TWSE 校正）
  ② 因子表移除四列（加權指數月線/季線、台積電月線、費半雙均）
  ③ 因子表字體全面放大（行高 0.041→0.053，字體 11/12→13/14）
================================================================
"""

import os, sys, warnings
warnings.filterwarnings("ignore")
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import requests
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.dates as mdates
from matplotlib.lines import Line2D

try:
    import yfinance as yf
    YF_OK = True
except ImportError:
    YF_OK = False
    print("⚠️  yfinance 未安裝：pip install yfinance")

import matplotlib.font_manager as fm
# 強制載入雲端下載的 NotoSans 字型
font_path = Path("NotoSans.otf")
if font_path.exists():
    prop = fm.FontProperties(fname=str(font_path))
    fm.fontManager.addfont(str(font_path))
    matplotlib.rcParams['font.sans-serif'] = [prop.get_name()] + matplotlib.rcParams['font.sans-serif']
else:
    for _f in ['Microsoft JhengHei', 'PingFang TC', 'Heiti TC', 'Noto Sans CJK TC']:
        try:
            matplotlib.rcParams['font.sans-serif'] = [_f] + matplotlib.rcParams['font.sans-serif']
            break
        except Exception:
            pass
matplotlib.rcParams['axes.unicode_minus'] = False

# ═══════════════════════════════════════════════════════════════
# Config
# ═══════════════════════════════════════════════════════════════
SCRIPT_DIR    = Path(os.getcwd())
TEMPLATE_PATH = SCRIPT_DIR / "daily_template.pptx"
OUTPUT_DIR    = SCRIPT_DIR / "reports"

FINMIND_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VyX2lkIjoia3VvODYwMSIsImVtYWlsIjoic29sZGllcjg2MTAwQGdtYWlsLmNvbSIsInRva2VuX3ZlcnNpb24iOjB9._5JgdrkR3h3ogK7zaxW1t7R4UxB0rbR-_aZUm3z0HLQ"

LE, SE          = 3.5, 5.5
LX, SX          = 2.0, 2.0
EC_L, EC_S      = 5, 2
COST            = 0.0005
DYN_WINDOW      = 60
DYN_MIN_TRIG    = 3
DYN_LO, DYN_HI  = 0.5, 1.5
FWD_DAYS        = 20

BASE_L = {'fL1':1.0,'fL2':1.0,'fL3':2.0,'fL4':2.0,'fL5':2.0,'fL6':1.0,'fL7':0.5,
          'fL8':1.0,'fL9':1.0,'fL10':1.0,'fL11':0.5,'fL12':1.0}
BASE_S = {'fS1':1.5,'fS2':1.0,'fS3':2.0,'fS4':2.0,'fS5':2.0,'fS6':1.0,'fS7':0.5,
          'fS8':1.0,'fS9':1.0,'fS10':1.0,'fS11':0.5,'fS12':1.0,'fS13':1.0}


# ═══════════════════════════════════════════════════════════════
# 1. 資料抓取（Yahoo Finance + CSV fallback）
# ═══════════════════════════════════════════════════════════════
def fetch_market_data() -> pd.DataFrame:
    """Yahoo Finance；失敗時 fallback 本地 CSV"""
    if YF_OK:
        end   = datetime.now()
        start = end - timedelta(days=540)
        specs = {
            "TWII"    : "^TWII",
            "SOX"     : "^SOX",
            "TSMC_TW" : "2330.TW",
            "TSM_US"  : "TSM",
            "ELEC"    : "0053.TW",
            "FIN"     : "0055.TW",
            "USDTWD"  : "USDTWD=X",   # 正確方向：1 USD ≈ 31 TWD
        }
        frames = {}
        for name, ticker in specs.items():
            try:
                raw = yf.Ticker(ticker).history(start=start, end=end)
                if raw.empty: continue
                if raw.index.tz is not None:
                    raw.index = raw.index.tz_localize(None)
                raw.index = pd.to_datetime(raw.index).normalize()
                frames[name] = raw[["Close"]].rename(columns={"Close": name})
                if name == "TSMC_TW":
                    frames["TSMC_Vol"] = raw[["Volume"]].rename(columns={"Volume":"TSMC_Vol"})
                if name == "TWII":
                    frames["TWII_Open"] = raw[["Open"]].rename(columns={"Open":"TWII_Open"})
            except Exception as e:
                print(f"  ⚠️  {ticker}: {e}")

        if frames and "TWII" in frames:
            df = pd.concat(frames.values(), axis=1).sort_index().ffill()

            # USDTWD 合理性驗證（應介於 15~50 TWD/USD）
            if "USDTWD" in df.columns:
                rate = df["USDTWD"].iloc[-1]
                if not (15 < rate < 50):
                    print(f"  ⚠️  USDTWD 異常值 {rate:.4f}，自動取倒數修正")
                    df["USDTWD"] = 1.0 / df["USDTWD"]
                print(f"  ✓  匯率 USD/TWD = {df['USDTWD'].iloc[-1]:.2f}")

            # ADR 折溢價：1 TSM ADR = 5 股 2330.TW
            if all(c in df.columns for c in ["TSM_US","USDTWD","TSMC_TW"]):
                df["ADR_Premium"] = (df["TSM_US"] * df["USDTWD"] / 5.0
                                     / df["TSMC_TW"] - 1.0) * 100.0
            else:
                df["ADR_Premium"] = np.nan

            df = df.dropna(subset=["TWII"]).reset_index().rename(columns={"index":"date"})
            if "date" not in df.columns:
                df = df.rename(columns={df.columns[0]:"date"})
            return df

    csv_path = SCRIPT_DIR / "台指量化模型基礎資料表_包含電子金融.csv"
    if csv_path.exists():
        print(f"  ⚠️  Yahoo 失敗，改用 CSV: {csv_path.name}")
        return _load_from_csv(csv_path)
    raise RuntimeError("無法取得市場資料")


def _load_from_csv(path) -> pd.DataFrame:
    raw = pd.read_csv(path)
    col_map = {
        '日期':'date','台灣加權指數開盤價':'TWII_Open','台灣加權指數收盤價':'TWII',
        '費半收盤價':'SOX','台積電收盤':'TSMC_TW','台積電成交量':'TSMC_Vol',
        '台積電ADR折溢價(%)':'ADR_Premium',
        '電子指數(0053)收盤價':'ELEC','金融指數(0055)收盤價':'FIN',
    }
    raw = raw.rename(columns=col_map)
    raw['date'] = pd.to_datetime(raw['date'])
    for c in raw.columns:
        if c != 'date' and raw[c].dtype == object:
            raw[c] = pd.to_numeric(
                raw[c].astype(str).str.replace(',','').str.replace('%',''), errors='coerce')
    raw = raw.dropna(subset=['TWII','TWII_Open']).ffill().reset_index(drop=True)
    fi_col = '外資合計買賣超金額(百萬)'
    if fi_col in raw.columns:
        raw['FI_Net_csv'] = pd.to_numeric(raw[fi_col], errors='coerce') / 100
    return raw


def fetch_foreign_investor(start_date: str) -> pd.Series:
    try:
        url    = "https://api.finmindtrade.com/api/v4/data"
        params = {"dataset":"TaiwanStockTotalInstitutionalInvestors",
                  "start_date":start_date,"token":FINMIND_TOKEN}
        res    = requests.get(url, params=params,
                              headers={"User-Agent":"Mozilla/5.0"}, timeout=15)
        data   = res.json()
        if data.get("msg") != "success": return pd.Series(dtype=float)
        dff    = pd.DataFrame(data["data"])
        mask   = dff["name"].str.contains("外資|Foreign_Investor", case=False, na=False) & \
                 ~dff["name"].str.contains("自營商|Dealer", case=False, na=False)
        f = dff[mask].copy()
        if f.empty: return pd.Series(dtype=float)
        f["Date"]   = pd.to_datetime(f["date"]).dt.normalize()
        f["FI_Net"] = (f["buy"].astype(float)-f["sell"].astype(float))/1e8
        return f.groupby("Date")["FI_Net"].sum()
    except Exception:
        return pd.Series(dtype=float)


# ═══════════════════════════════════════════════════════════════
# 2. 因子計算
# ═══════════════════════════════════════════════════════════════
def build_factors(df: pd.DataFrame, fi: pd.Series) -> pd.DataFrame:
    d = df.copy()
    d["MA5"]  = d["TWII"].rolling(5).mean()
    d["MA10"] = d["TWII"].rolling(10).mean()
    d["MA20"] = d["TWII"].rolling(20).mean()
    d["MA60"] = d["TWII"].rolling(60).mean()
    d["斜率"]  = (d["MA60"].diff(5) / d["MA60"].shift(5)) * 100
    d["乖離"]  = (d["TWII"] - d["MA60"]) / d["MA60"] * 100
    d["STD20"] = d["TWII"].rolling(20).std()
    d["BB上"]  = d["MA20"] + 2*d["STD20"]
    d["BB下"]  = d["MA20"] - 2*d["STD20"]
    δ = d["TWII"].diff()
    d["RSI"] = 100 - (100 / (1 +
        δ.clip(lower=0).ewm(com=13,adjust=False).mean() /
        (-δ.clip(upper=0)).ewm(com=13,adjust=False).mean().replace(0,np.nan)))
    d["MOM5"]    = (d["TWII"] / d["TWII"].shift(5) - 1) * 100
    d["SOX_MA20"]= d["SOX"].rolling(20).mean()
    d["SOX_MA60"]= d["SOX"].rolling(60).mean()
    d["TS_MA20"] = d["TSMC_TW"].rolling(20).mean()
    d["TS_VolMA"]= d["TSMC_Vol"].rolling(10).mean()
    d["EF"]      = d["ELEC"] / d["FIN"]
    d["EF_MA20"] = d["EF"].rolling(20).mean()
    d["EF_MA60"] = d["EF"].rolling(60).mean()
    if not fi.empty:
        d = d.set_index("date").join(fi.rename("FI_Net"), how="left").reset_index()
        d["FI_Net"]  = d["FI_Net"].ffill().fillna(0)
        d["FI_MA"]   = d["FI_Net"].rolling(120,min_periods=30).mean()
        d["FI_STD"]  = d["FI_Net"].rolling(120,min_periods=30).std().replace(0,np.nan)
        d["FI_Z"]    = (d["FI_Net"] - d["FI_MA"]) / d["FI_STD"]
        d["FI_5MA"]  = d["FI_Net"].rolling(5).mean()
    else:
        d["FI_Net"] = 0.0; d["FI_Z"] = 0.0; d["FI_5MA"] = 0.0
    d["ADR_MA"]  = d["ADR_Premium"].rolling(120,min_periods=30).mean()
    d["ADR_STD"] = d["ADR_Premium"].rolling(120,min_periods=30).std().replace(0,np.nan)
    d["ADR_Z"]   = (d["ADR_Premium"] - d["ADR_MA"]) / d["ADR_STD"]
    d["fL1"]  = ((d["TWII"]>d["MA60"]) & (d["斜率"]>0.1)).astype(float)
    d["fL2"]  = (d["EF"]>d["EF_MA20"]).astype(float)
    d["fL3"]  = (d["FI_Z"]>1.2).astype(float)
    d["fL4"]  = ((d["SOX"]>d["SOX_MA20"]) & (d["SOX"]>d["SOX_MA60"])).astype(float)
    d["fL5"]  = (d["ADR_Z"]>0.8).astype(float)
    d["fL6"]  = (d["TSMC_TW"]>d["TS_MA20"]).astype(float)
    d["fL7"]  = (d["TSMC_Vol"]>1.5*d["TS_VolMA"]).astype(float)
    d["fL8"]  = (d["乖離"]<-8).astype(float)
    d["fL9"]  = (d["RSI"]<40).astype(float)
    d["fL10"] = (d["FI_5MA"]>0).astype(float)
    d["fL11"] = (d["TWII"]<d["BB下"]).astype(float)
    d["fL12"] = (d["MOM5"]>2).astype(float)
    d["fS1"]  = ((d["TWII"]<d["MA60"]) & (d["斜率"]<-0.1)).astype(float)
    d["fS2"]  = (d["EF"]<d["EF_MA20"]).astype(float)
    d["fS3"]  = (d["FI_Z"]<-1.2).astype(float)
    d["fS4"]  = ((d["SOX"]<d["SOX_MA20"]) & (d["SOX"]<d["SOX_MA60"])).astype(float)
    d["fS5"]  = (d["ADR_Z"]<-1.0).astype(float)
    d["fS6"]  = (d["TSMC_TW"]<d["TS_MA20"]).astype(float)
    d["fS7"]  = (d["TSMC_Vol"]>1.5*d["TS_VolMA"]).astype(float)
    d["fS8"]  = (d["乖離"]>8).astype(float)
    d["fS9"]  = (d["RSI"]>55).astype(float)
    d["fS10"] = (d["FI_5MA"]<0).astype(float)
    d["fS11"] = (d["TWII"]>d["BB上"]).astype(float)
    d["fS12"] = (d["TWII"]<d["MA60"]).astype(float)
    d["fS13"] = (d["MOM5"]<-2).astype(float)
    d["fwd20"] = d["TWII"].shift(-FWD_DAYS) / d["TWII"] - 1
    return d


def calc_dynamic(factor_arr, fwd_arr, is_long):
    N = len(factor_arr); hits = np.full(N, 0.5)
    for i in range(DYN_WINDOW, N):
        fw = factor_arr[i-DYN_WINDOW:i]; rw = fwd_arr[i-DYN_WINDOW:i]
        trig = fw == 1.0
        if trig.sum() >= DYN_MIN_TRIG:
            rs = rw[trig]; rs = rs[~np.isnan(rs)]
            if len(rs) >= DYN_MIN_TRIG:
                hits[i] = (rs>0).mean() if is_long else (rs<0).mean()
    return hits


def compute_scores(d):
    fwd   = d["fwd20"].values
    dyn_L = {k: calc_dynamic(d[k].values, fwd, True)  for k in BASE_L}
    dyn_S = {k: calc_dynamic(d[k].values, fwd, False) for k in BASE_S}
    N = len(d); ml = np.zeros(N); ms = np.zeros(N)
    for k,w in BASE_L.items():
        ml += d[k].values * w * np.clip(dyn_L[k]*2, DYN_LO, DYN_HI)
    for k,w in BASE_S.items():
        ms += d[k].values * w * np.clip(dyn_S[k]*2, DYN_LO, DYN_HI)
    gL = ((d["fL4"]==1)&((d["fL5"]==1)|(d["fL3"]==1))&(d["EF"]>d["EF_MA60"])).values
    gS = ((d["fS4"]==1)&((d["fS5"]==1)|(d["fS3"]==1))&(d["EF"]<d["EF_MA20"])).values
    return ml, ms, gL, gS


def backtest(d, ml, ms, gL, gS):
    N = len(d)
    close = d["TWII"].values; open_ = d["TWII_Open"].values
    ma10  = d["MA10"].values; ma60  = d["MA60"].values
    intra = np.where(open_>0, close/open_-1, 0)
    onight = np.zeros(N)
    onight[1:] = np.where(close[:-1]>0, open_[1:]/close[:-1]-1, 0)
    daily = np.zeros(N)
    daily[1:]  = np.where(close[:-1]>0, close[1:]/close[:-1]-1, 0)
    pos = np.zeros(N); cur=0.0; ec=0; quick=0
    for i in range(N):
        if cur == 0:
            if ml[i]>=LE and gL[i]:    cur=1.;  ec=0; quick=0
            elif ms[i]>=SE and gS[i]:  cur=-1.; ec=0; quick=0
        elif cur == 1:
            esig = ml[i]<LX or not gL[i] or close[i]<ma60[i]
            ec   = ec+1 if esig else 0
            if ec>=EC_L: cur=0.; ec=0
        else:
            if close[i]>ma10[i]: quick+=1
            else: quick=0
            if quick>=1:
                cur=0.; ec=0; quick=0; pos[i]=cur; continue
            esig = ms[i]<SX or not gS[i] or close[i]>ma60[i]
            ec   = ec+1 if esig else 0
            if ec>=EC_S: cur=0.; ec=0
        pos[i] = cur
    exp  = np.roll(pos,1); exp[0]=0
    expp = np.roll(exp,1); expp[0]=0
    ret  = np.zeros(N)
    me=(exp!=0)&(expp==0);           ret[me]=exp[me]*intra[me]
    mh=(exp!=0)&(expp==exp);         ret[mh]=exp[mh]*daily[mh]
    mx=(exp==0)&(expp!=0);           ret[mx]=expp[mx]*onight[mx]
    mr=(exp!=0)&(expp!=0)&(exp!=expp)
    ret[mr]=expp[mr]*onight[mr]+exp[mr]*intra[mr]
    ret -= np.abs(np.diff(exp, prepend=0))*COST
    trades=[]; it=False; tr=[]; cd=0; ei=0
    for i in range(N):
        e=exp[i]
        if not it and e!=0:
            it=True; tr=[ret[i]]; cd=int(e); ei=i
            trades.append({"entry_idx":i,"entry_date":d["date"].iloc[i],
                           "dir_code":cd,"entry_price":close[i]})
        elif it and e!=0:
            tr.append(ret[i])
        elif it and e==0:
            tr.append(ret[i])
            p=np.prod(1+np.array(tr))-1
            trades[-1].update({"exit_idx":i,"exit_date":d["date"].iloc[i],
                               "exit_price":close[i],"pct_return":p*100,
                               "n_days":i-ei,"is_win":p>0})
            it=False; tr=[]
    return {"daily_ret":ret,"exp":exp,"pos":pos,"cum":np.cumprod(1+ret),"trades":trades}


# ═══════════════════════════════════════════════════════════════
# 3. 狀態判斷
# ═══════════════════════════════════════════════════════════════
def determine_state(d, bt, ml, ms, gL, gS):
    i=len(d)-1; pt=int(bt["pos"][i]); py=int(bt["pos"][i-1]) if i>=1 else 0
    rec={( 1, 1):"多單續抱",( 1, 0):"多單建倉",( 1,-1):"空單轉多",
         ( 0, 1):"多單出場",( 0, 0):"空手觀望",( 0,-1):"空單出場",
         (-1, 1):"多單轉空",(-1, 0):"空單建倉",(-1,-1):"空單續抱",
         }.get((pt,py),"空手觀望")
    if   rec in ("多單續抱","多單建倉","空單轉多"): bias="偏多"
    elif rec in ("空單續抱","空單建倉","多單轉空"): bias="偏空"
    else:                                            bias="震盪"
    return {"recommendation":rec,"bias":bias,
            "ml":float(ml[i]),"ms":float(ms[i]),
            "gL":bool(gL[i]),"gS":bool(gS[i]),
            "pos_today":pt,"pos_yesterday":py,
            "last_close":float(d["TWII"].iloc[i]),
            "last_date":d["date"].iloc[i].date()}


def compute_1yr_stats(d, bt):
    cutoff  = d["date"].iloc[-1]-pd.Timedelta(days=365)
    mask    = (d["date"]>=cutoff).values
    if mask.sum()<30: return {"error":"資料不足一年"}
    sub_ret = bt["daily_ret"][mask]; sub_exp=bt["exp"][mask]
    sc=np.cumprod(1+sub_ret); tot=(sc[-1]-1)*100; yrs=mask.sum()/252
    trades_1y=[t for t in bt["trades"] if "exit_date" in t and t["entry_date"]>=cutoff]
    n_t=len(trades_1y); n_l=sum(1 for t in trades_1y if t["dir_code"]==1)
    n_s=sum(1 for t in trades_1y if t["dir_code"]==-1)
    w_l=sum(1 for t in trades_1y if t["dir_code"]==1  and t["is_win"])
    w_s=sum(1 for t in trades_1y if t["dir_code"]==-1 and t["is_win"])
    w_t=sum(1 for t in trades_1y if t["is_win"])
    return {"period":(d["date"][mask].iloc[0].date(),d["date"].iloc[-1].date()),
            "n_trades":n_t,"n_long":n_l,"n_short":n_s,
            "wr_all":(w_t/n_t*100) if n_t else 0,
            "wr_long":(w_l/n_l*100) if n_l else 0,
            "wr_short":(w_s/n_s*100) if n_s else 0,
            "strat_total":tot,
            "strat_ann":((1+tot/100)**(1/yrs)-1)*100,
            "strat_mdd":((sc/np.maximum.accumulate(sc)-1)*100).min(),
            "strat_vol":sub_ret.std()*np.sqrt(252)*100,
            "strat_sharpe":sub_ret.mean()*252/(sub_ret.std()*np.sqrt(252)+1e-9),
            "in_mkt":(sub_exp!=0).mean()*100,
            "in_long":(sub_exp==1).mean()*100,
            "in_short":(sub_exp==-1).mean()*100}


# ═══════════════════════════════════════════════════════════════
# 4. 圖檔生成
# ═══════════════════════════════════════════════════════════════
def make_stats_image(stats, out_path):
    rows = [
        ('回測期間',
         f"{stats['period'][0].strftime('%Y/%m/%d')}~{stats['period'][1].strftime('%Y/%m/%d')}",
         '#1E3A5F'),
        ('交易筆數',    f"{stats['n_trades']} 筆",                '#1E3A5F'),
        ('多單 / 空單', f"{stats['n_long']} / {stats['n_short']}", '#1E3A5F'),
        ('SEP','',''),
        ('整體勝率', f"{stats['wr_all']:.1f}%",
         '#DC2626' if stats['wr_all']>=60 else '#1E3A5F'),
        ('多頭勝率', f"{stats['wr_long']:.1f}%",
         '#DC2626' if stats['wr_long']>=60 else '#1E3A5F'),
        ('空頭勝率', f"{stats['wr_short']:.1f}%",
         '#DC2626' if stats['wr_short']>=60 else '#1E3A5F'),
        ('SEP','',''),
        ('整體累積績效', f"{stats['strat_total']:+.2f}%",
         '#DC2626' if stats['strat_total']>0 else '#16A34A'),
        ('SEP','',''),
        ('策略年化報酬', f"{stats['strat_ann']:+.2f}%",
         '#DC2626' if stats['strat_ann']>0 else '#16A34A'),
        ('年化波動度',   f"{stats['strat_vol']:.2f}%",  '#1E3A5F'),
        ('最大回撤',     f"{stats['strat_mdd']:.2f}%",  '#16A34A'),
        ('夏普比率',     f"{stats['strat_sharpe']:.2f}", '#DC2626'),
        ('SEP','',''),
        ('在市場時間', f"{stats['in_mkt']:.1f}%",   '#1E3A5F'),
        ('多單時間',   f"{stats['in_long']:.1f}%",  '#DC2626'),
        ('空單時間',   f"{stats['in_short']:.1f}%", '#16A34A'),
    ]
    fig, ax = plt.subplots(figsize=(5.3, 7.05), facecolor='white')
    ax.set_facecolor('white'); ax.axis('off')
    ax.set_xlim(0,1); ax.set_ylim(0,1)
    y=0.985; row_h=0.052; sep_h=0.013; alt=0
    for label, val, color in rows:
        if label=='SEP': y-=sep_h; continue
        bg='#F1F5F9' if alt%2==0 else '#FFFFFF'
        ax.add_patch(plt.Rectangle((0.02,y-row_h),0.96,row_h,
                                   facecolor=bg,edgecolor='#E2E8F0',linewidth=0.6))
        ax.text(0.05,y-row_h/2,label,fontsize=14,color='#374151',va='center')
        ax.text(0.95,y-row_h/2,val,  fontsize=14,color=color,
                va='center',ha='right',fontweight='bold')
        y-=row_h; alt+=1
    plt.savefig(out_path,dpi=130,facecolor='white',bbox_inches='tight',pad_inches=0.03)
    plt.close()


def make_factor_image(d, out_path):
    """
    重要因子最新狀況
    ★ v2.6：移除月線MA20/季線MA60（TWII）、費半雙均、台積電月線
             共 4 列 → 剩 14 資料列 + 4 分組標題 = 18 列
             行高 0.041→0.053，字體 11/12→13/14
    PPT 插入：5.30" × 7.24"  →  figsize=(5.3, 7.24)
    """
    i=len(d)-1; p=max(0,i-1)
    def g(c,idx=i):
        try:    return float(d[c].iloc[idx])
        except: return 0.0

    twii=g('TWII'); tp=g('TWII',p)
    twii_chg=(twii/tp-1)*100 if tp>0 else 0
    bias=g('乖離'); rsi=g('RSI'); rp=g('RSI',p); mom5=g('MOM5')
    ef=g('EF'); efp=g('EF',p); ef20=g('EF_MA20'); ef60=g('EF_MA60')
    sox=g('SOX'); soxp=g('SOX',p)
    sox_chg=(sox/soxp-1)*100 if soxp>0 else 0
    sox20=g('SOX_MA20'); sox60=g('SOX_MA60')
    tsmc=g('TSMC_TW'); tsmcp=g('TSMC_TW',p)
    tsmc_chg=(tsmc/tsmcp-1)*100 if tsmcp>0 else 0
    tvol=g('TSMC_Vol'); tvma=g('TS_VolMA')
    vr=tvol/tvma if tvma>0 else 0
    adr=g('ADR_Premium'); adrp=g('ADR_Premium',p); adrz=g('ADR_Z')
    fi=g('FI_Net'); fi5=g('FI_5MA'); fiz=g('FI_Z')

    def sta(up,dn,cond): return (up,'g') if cond else (dn,'r')
    sox_s,sox_c=(('雙均偏多','g') if sox>sox20 and sox>sox60 else
                 ('雙均偏空','r') if sox<sox20 and sox<sox60 else ('震盪','n'))

    # ── 因子列表（已移除 4 列）──────────────────────────────────
    rows = [
        # ── 加權指數（移除月線MA20、季線MA60）──
        ('GROUP','加權指數 (TAIEX)',None,None),
        ('收盤指數',   f"{twii:,.0f}",   f"{twii_chg:+.2f}%",'g' if twii_chg>0 else 'r'),
        # 月線 MA20 → 移除
        # 季線 MA60 → 移除
        ('季線乖離率', f"{bias:+.2f}%",
            '超賣' if bias<-8 else ('超買' if bias>8 else '常態'),
            'g' if bias<-8 else ('r' if bias>8 else 'n')),
        ('RSI(14)',    f"{rsi:.1f}",     f"{rsi-rp:+.1f}",
            'g' if rsi<40 else ('r' if rsi>55 else 'n')),
        ('5日動量',    f"{mom5:+.2f}%",
            '強多' if mom5>2 else ('強空' if mom5<-2 else '中性'),
            'g' if mom5>2 else ('r' if mom5<-2 else 'n')),

        # ── 風格輪動（不變）──
        ('GROUP','風格輪動 (電金比)',None,None),
        ('電金比',   f"{ef:.4f}",   f"{ef-efp:+.4f}", 'g' if ef>efp else 'r'),
        ('vs 月線',  f"{ef20:.4f}", *sta('站上','跌破',ef>ef20)),
        ('vs 季線',  f"{ef60:.4f}", *sta('站上','跌破',ef>ef60)),

        # ── 半導體（移除費半雙均、台積電月線）──
        ('GROUP','半導體 (海外+本土)',None,None),
        ('費城半導體', f"{sox:,.0f}",  f"{sox_chg:+.2f}%", 'g' if sox_chg>0 else 'r'),
        # 費半雙均 → 移除
        ('台積電現貨', f"{tsmc:,.0f}", f"{tsmc_chg:+.2f}%",'g' if tsmc_chg>0 else 'r'),
        # 台積電月線 → 移除
        ('台積量比',   f"{vr:.2f}x",
            '爆量' if vr>1.5 else '正常','r' if vr>1.5 else 'n'),

        # ── 海外資金面（ADR 折溢價原始數值移除，保留 ADR Z-score）──
        ('GROUP','海外資金面',None,None),
        ('ADR Z-score', f"{adrz:+.2f}s",
            '搶單' if adrz>0.8 else ('撤退' if adrz<-1.0 else '中性'),
            'g' if adrz>0.8 else ('r' if adrz<-1.0 else 'n')),
        ('外資今日(億)', f"{fi:+.1f}",  '買超' if fi>0 else '賣超', 'g' if fi>0 else 'r'),
        ('外資 5MA(億)', f"{fi5:+.1f}", '連買' if fi5>0 else '連賣','g' if fi5>0 else 'r'),
        ('外資 Z-score', f"{fiz:+.2f}s",
            '強買超' if fiz>1.2 else ('強賣超' if fiz<-1.2 else '中性'),
            'g' if fiz>1.2 else ('r' if fiz<-1.2 else 'n')),
    ]
    # 共 17 列（4 GROUP + 13 資料列）

    cmap = {'g':'#DC2626','r':'#16A34A','n':'#64748B'}   # 台灣慣例：漲=紅 跌=綠

    # figsize=(5.3, 7.24) 對應 PPT 插入 5.30"×7.24"
    fig, ax = plt.subplots(figsize=(5.3, 7.24), facecolor='white')
    ax.set_facecolor('white'); ax.axis('off')
    ax.set_xlim(0,1); ax.set_ylim(0,1)

    # 欄位標頭（字體放大）
    ax.text(0.04, 0.988, '因子',   fontsize=13, color='#64748B',
            fontweight='bold', va='top')
    ax.text(0.57, 0.988, '當前值', fontsize=13, color='#64748B',
            fontweight='bold', va='top', ha='center')
    ax.text(0.87, 0.988, '狀態',   fontsize=13, color='#64748B',
            fontweight='bold', va='top', ha='center')

    # 18 列 × 行高 0.053 = 0.954，起點 0.960，底部留 0.006 空白
    y = 0.960
    rh = 0.053   # ← 行高放大（原 0.041）

    for row in rows:
        if row[0] == 'GROUP':
            ax.add_patch(plt.Rectangle((0.02,y-rh),0.96,rh,
                                       facecolor='#1E3A5F',edgecolor='none'))
            ax.text(0.05,y-rh/2,row[1],
                    fontsize=13,color='white',fontweight='bold',va='center')
            y -= rh
        else:
            lbl, val, sts, ch = row
            ax.add_patch(plt.Rectangle((0.02,y-rh),0.96,rh,
                                       facecolor='#FAFAFA',edgecolor='#E2E8F0',linewidth=0.5))
            ax.text(0.04,y-rh/2, lbl, fontsize=13, color='#1E293B', va='center')
            ax.text(0.57,y-rh/2, val, fontsize=14, color='#1E3A5F',
                    ha='center',va='center',fontweight='bold')
            ax.text(0.87,y-rh/2, sts, fontsize=13, color=cmap.get(ch,'#64748B'),
                    ha='center',va='center',fontweight='bold')
            y -= rh

    plt.savefig(out_path, dpi=130, facecolor='white',
                bbox_inches='tight', pad_inches=0.03)
    plt.close()


def make_chart(d, bt, out_path):
    cutoff  = d["date"].iloc[-1]-pd.Timedelta(days=365)
    sub     = d[d["date"]>=cutoff].copy().reset_index(drop=True)
    exp_sub = bt["exp"][(d["date"]>=cutoff).values]
    fig, ax = plt.subplots(figsize=(14,5.0),facecolor='white')
    ax.set_facecolor('#F9FAFB')
    in_long=False; in_short=False; ls_dt=sub["date"].iloc[0]
    for i in range(len(sub)):
        e=exp_sub[i]; dt=sub["date"].iloc[i]
        if e==1 and not in_long:
            ls_dt=dt; in_long=True; in_short=False
        elif e==-1 and not in_short:
            ls_dt=dt; in_short=True; in_long=False
        elif e==0 and (in_long or in_short):
            ax.axvspan(ls_dt,dt,color='#DCFCE7' if in_long else '#FEE2E2',
                       alpha=0.55,zorder=1,lw=0)
            in_long=False; in_short=False
    if in_long or in_short:
        ax.axvspan(ls_dt,sub["date"].iloc[-1],
                   color='#FEE2E2' if in_long else '#DCFCE7',alpha=0.55,zorder=1,lw=0)
    ax.plot(sub["date"],sub["MA60"],color='#F43F5E',linewidth=1.4,
            linestyle='--',alpha=0.75,zorder=2)
    ax.plot(sub["date"],sub["MA10"],color='#F59E0B',linewidth=1.1,
            linestyle=(0,(3,2)),alpha=0.75,zorder=2)
    ax.plot(sub["date"],sub["TWII"],color='#1E40AF',linewidth=2.4,
            zorder=3,solid_capstyle='round')
    twii_rng=sub["TWII"].max()-sub["TWII"].min(); off=twii_rng*0.016
    trades_1y=[t for t in bt["trades"] if t["entry_date"]>=cutoff]
    for t in trades_1y:
        is_open="exit_date" not in t
        if t["dir_code"]==1:
            ax.scatter(t["entry_date"],t["entry_price"]-off,
                       marker='^',s=170,color='#DC2626',edgecolor='white',
                       linewidth=1.8,zorder=6)
            if is_open:
                ax.annotate('持倉中',xy=(t["entry_date"],t["entry_price"]-off),
                            xytext=(10,18),textcoords='offset points',
                            fontsize=9,color='#B91C1C',fontweight='bold',
                            bbox=dict(boxstyle='round,pad=0.2',facecolor='#FEE2E2',
                                      edgecolor='#DC2626',linewidth=1))
            else:
                ax.scatter(t["exit_date"],t["exit_price"],
                           marker='o',s=75,color='#4ADE80',edgecolor='#16A34A',
                           linewidth=1.5,zorder=6)
        else:
            ax.scatter(t["entry_date"],t["entry_price"]+off,
                       marker='v',s=170,color='#16A34A',edgecolor='white',
                       linewidth=1.8,zorder=6)
            if is_open:
                ax.annotate('持倉中',xy=(t["entry_date"],t["entry_price"]+off),
                            xytext=(10,-22),textcoords='offset points',
                            fontsize=9,color='#15803D',fontweight='bold',
                            bbox=dict(boxstyle='round,pad=0.2',facecolor='#DCFCE7',
                                      edgecolor='#16A34A',linewidth=1))
            else:
                ax.scatter(t["exit_date"],t["exit_price"],
                           marker='o',s=75,color='#4ADE80',edgecolor='#16A34A',
                           linewidth=1.5,zorder=6)
    legend_elems=[
        Line2D([0],[0],color='#1E40AF',lw=2.4,label='加權指數'),
        Line2D([0],[0],color='#F43F5E',lw=1.4,ls='--',label='季線(60MA)'),
        Line2D([0],[0],color='#F59E0B',lw=1.1,ls=(0,(3,2)),label='MA10'),
        Line2D([0],[0],marker='^',color='w',markerfacecolor='#DC2626',
               markersize=13,label='多單進場'),
        Line2D([0],[0],marker='o',color='w',markerfacecolor='#FCA5A5',
               markeredgecolor='#DC2626',markersize=10,label='多單出場'),
        Line2D([0],[0],marker='v',color='w',markerfacecolor='#16A34A',
               markersize=13,label='空單進場'),
        Line2D([0],[0],marker='o',color='w',markerfacecolor='#4ADE80',
               markeredgecolor='#16A34A',markersize=10,label='空單出場'),
    ]
    ax.legend(handles=legend_elems,loc='upper left',fontsize=12,ncol=7,
              frameon=True,framealpha=0.92,edgecolor='#D1D5DB',
              bbox_to_anchor=(0.0,1.0),handlelength=1.5,columnspacing=0.7)
    ymin=sub["TWII"].min()*0.98; ymax=sub["TWII"].max()*1.02
    ax.set_ylim(ymin,ymax)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x,_: f"{x:,.0f}"))
    ax.yaxis.set_major_locator(mticker.MaxNLocator(nbins=6,prune='both'))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%Y'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    plt.setp(ax.xaxis.get_majorticklabels(),rotation=0,ha='center',fontsize=12)
    ax.tick_params(axis='y',labelsize=12,colors='#374151',length=3)
    ax.tick_params(axis='x',labelsize=12,colors='#374151',length=3)
    ax.set_ylabel('加權指數',fontsize=13,color='#374151',labelpad=8)
    ax.yaxis.grid(True,alpha=0.4,linestyle='--',linewidth=0.6,color='#CBD5E1')
    ax.set_axisbelow(True)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#E5E7EB'); ax.spines['bottom'].set_color('#E5E7EB')
    plt.tight_layout(pad=0.8)
    plt.savefig(out_path,dpi=150,facecolor='white',bbox_inches='tight')
    plt.close()


# ═══════════════════════════════════════════════════════════════
# 5. PPT 套版生成
# ═══════════════════════════════════════════════════════════════
def generate_pptx(run_date, state, stats_img, factor_img, chart_img, output_path):
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
    except ImportError:
        raise ImportError("請先安裝 python-pptx：pip install python-pptx")
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"模板不存在：{TEMPLATE_PATH}")
    prs=Presentation(str(TEMPLATE_PATH)); slide=prs.slides[0]
    for shape in slide.shapes:
        if shape.name=='群組 28':
            for sub in shape.shapes:
                if sub.has_text_frame and (
                        str(run_date.year) in sub.text_frame.text or
                        '/' in sub.text_frame.text):
                    paras=list(sub.text_frame.paragraphs)
                    for p_idx,p in enumerate(paras):
                        if not p.runs: continue
                        if   p_idx==0: p.runs[0].text=str(run_date.year)
                        elif p_idx==1: p.runs[0].text=f"{run_date.month:02d}/{run_date.day:02d}"
        elif '文字方塊 4' in shape.name:
            for p in shape.text_frame.paragraphs:
                if len(p.runs)>=7:
                    rec=state["recommendation"]; bias=state["bias"]
                    p.runs[2].text=rec+"　　　　 　　"
                    if   rec in ("多單續抱","多單建倉","空單轉多"):
                        p.runs[2].font.color.rgb=RGBColor.from_string("DC2626")
                    elif rec in ("空單續抱","空單建倉","多單轉空"):
                        p.runs[2].font.color.rgb=RGBColor.from_string("16A34A")
                    else:
                        p.runs[2].font.color.rgb=RGBColor.from_string("1E293B")
                    p.runs[6].text=bias
                    if   bias=="偏多": p.runs[6].font.color.rgb=RGBColor.from_string("DC2626")
                    elif bias=="偏空": p.runs[6].font.color.rgb=RGBColor.from_string("16A34A")
                    else:              p.runs[6].font.color.rgb=RGBColor.from_string("1E293B")
    _EMU=914400
    for _s in slide.shapes:
        if _s.name=='文字方塊 29':
            _s.top=int(11.64*_EMU); _s.height=int(4.61*_EMU); break
    # 三張圖位置（對齊 2026-05-24 參考版本）
    slide.shapes.add_picture(stats_img,  Inches(0.00),Inches(4.40),
                             width=Inches(5.30),height=Inches(7.05))
    slide.shapes.add_picture(factor_img, Inches(5.70),Inches(4.40),
                             width=Inches(5.30),height=Inches(7.24))
    slide.shapes.add_picture(chart_img,  Inches(-0.08),Inches(12.25),
                             width=Inches(11.08),height=Inches(3.95))
    tb=slide.shapes.add_textbox(Inches(0.3),Inches(3.50),Inches(10.5),Inches(0.30))
    tf=tb.text_frame
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
    r=tf.paragraphs[0].add_run()
    r.text=(f"加權收盤 {state['last_close']:,.0f}  |  "
            f"多頭得分 ml={state['ml']:.2f} (門檻 {LE})  |  "
            f"空頭得分 ms={state['ms']:.2f} (門檻 {SE})  |  "
            f"大環境門票：多 {'OK' if state['gL'] else 'NG'}  "
            f"空 {'OK' if state['gS'] else 'NG'}")
    r.font.size=Pt(11); r.font.name="Microsoft JhengHei"
    r.font.color.rgb=RGBColor.from_string("475569")
    prs.save(str(output_path))


# ═══════════════════════════════════════════════════════════════
# 6. 主程式
# ═══════════════════════════════════════════════════════════════
def generate_daily_report():
    run_date=datetime.today().date()
    print("="*60)
    print(f"  台指策略日報  v15 多因子量化模型  v2.6")
    print(f"  執行日期：{run_date}  ({datetime.now().strftime('%H:%M:%S')})")
    print("="*60)

    df=fetch_market_data()
    print(f"✓ 資料：{len(df)} 個交易日  "
          f"({df['date'].iloc[0].date()} ~ {df['date'].iloc[-1].date()})")
    start_str=(df["date"].iloc[0]-timedelta(days=10)).strftime("%Y-%m-%d")
    fi=fetch_foreign_investor(start_str)
    if fi.empty and 'FI_Net_csv' in df.columns:
        fi=df.set_index('date')['FI_Net_csv'].dropna()
        print(f"✓ 外資：CSV fallback（{len(fi)} 筆）")
    elif not fi.empty:
        print(f"✓ 外資：FinMind（{len(fi)} 筆）")

    d=build_factors(df,fi)
    ml,ms,gL,gS=compute_scores(d)
    bt=backtest(d,ml,ms,gL,gS)
    state=determine_state(d,bt,ml,ms,gL,gS)
    stats=compute_1yr_stats(d,bt)

    i_last=len(d)-1
    print(f"\n  策略建議：{state['recommendation']}　｜　模型分析：{state['bias']}")
    print(f"  加權收盤：{state['last_close']:,.0f}（資料日期：{state['last_date']}）")
    print(f"  月線 MA20：{d['MA20'].iloc[i_last]:,.2f}　"
          f"季線 MA60：{d['MA60'].iloc[i_last]:,.2f}")
    print(f"  ADR 折溢價：{d['ADR_Premium'].iloc[i_last]:+.2f}%")
    print(f"  ml={state['ml']:.2f}（門檻{LE}）　ms={state['ms']:.2f}（門檻{SE}）")
    print(f"  門票　多：{'OK' if state['gL'] else 'NG'}　空：{'OK' if state['gS'] else 'NG'}")
    if "error" not in stats:
        print(f"\n  近一年　{stats['n_trades']} 筆  勝率 {stats['wr_all']:.1f}%  "
              f"累積 {stats['strat_total']:+.2f}%  "
              f"MDD {stats['strat_mdd']:.2f}%  Sharpe {stats['strat_sharpe']:.2f}")

    open_trades=[t for t in bt["trades"] if "exit_date" not in t]
    if open_trades:
        ot=open_trades[-1]
        dir_str="多單" if ot["dir_code"]==1 else "空單"
        print(f"\n  ★ 進行中部位：{dir_str}  進場 {ot['entry_date'].date()}"
              f"  進場價 {ot['entry_price']:,.0f}")

    OUTPUT_DIR.mkdir(exist_ok=True)
    stats_img  =str(OUTPUT_DIR/"_tmp_stats.png")
    factor_img =str(OUTPUT_DIR/"_tmp_factors.png")
    chart_img  =str(OUTPUT_DIR/"_tmp_chart.png")
    print("\n  生成圖檔...",end='',flush=True)
    make_stats_image(stats,stats_img)
    make_factor_image(d,factor_img)
    make_chart(d,bt,chart_img)
    print(" 完成")

    fname=f"台股策略日報_{run_date.strftime('%Y.%m.%d')}.pptx"
    output_path=OUTPUT_DIR/fname
    print(f"  套版輸出...",end='',flush=True)
    generate_pptx(run_date,state,stats_img,factor_img,chart_img,output_path)
    print(" 完成")

    for _p in (stats_img,factor_img,chart_img):
        try: Path(_p).unlink()
        except: pass

    print(f"\n{'='*60}")
    print(f"  ✅  日報已產出：{output_path}")
    print(f"{'='*60}")
    return output_path


if __name__=="__main__" or "ipykernel" in sys.modules:
    generate_daily_report()
