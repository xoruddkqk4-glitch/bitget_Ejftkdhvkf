import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
from datetime import datetime

# ======================================
# ✅ 데이터 소스: 새 공통모듈(LS_Common_new) + FDR
# - FDR을 직접 호출하지 않고, LS_Common_new.GetOhlcv(source="fdr")로 조회
# - 같은 폴더에 LS_Common_new.py 가 있어야 합니다.
#   (FinanceDataReader 미설치 시: pip install finance-datareader)
# ======================================
import LS_Common_new as LS

# ======================================
# ✅ 백테스팅 기간 설정
# - 예: "20180101" 또는 "2018-01-01"
# - None이면: 데이터에서 가능한 가장 가까운 날짜로 자동 보정
# ======================================
BACKTEST_START_STR = 20260510   # 예: "20180101" / "2018-01-01" / None
BACKTEST_END_STR   = None   # 예: "20251231" / "2025-12-31" / None

def _parse_date_str(s):
    """ 'YYYY-MM-DD' / 'YYYYMMDD' / datetime-like 를 Timestamp로 변환 """
    if s is None:
        return None
    if isinstance(s, (pd.Timestamp, datetime)):
        return pd.Timestamp(s).normalize()
    s = str(s).strip()
    if len(s) == 8 and s.isdigit():
        s = f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    return pd.to_datetime(s).normalize()

def _nearest_date(dates, target_ts: pd.Timestamp):
    """ dates(list[Timestamp]) 중 target에 가장 가까운 날짜 1개 반환 """
    return min(dates, key=lambda d: abs(pd.Timestamp(d) - target_ts))

# 한글 폰트 설정
plt.rc('font', family='Malgun Gothic')
plt.rcParams['axes.unicode_minus'] = False  # 음수 부호 깨짐 방지

# =========================
# 1) 기본 설정
# =========================
stock_name = 'SOXL'
current_date = datetime.now().strftime('%Y%m%d_%H%M%S')  # 시간 포함으로 중복 방지

# =========================
# ✅ 데이터 로딩: 새 공통모듈(LS_Common_new) + FDR
#    - FDR을 직접 호출하지 않고 LS_Common_new.GetOhlcv(source="fdr")로 조회
#    - 분할만 보정된 Close 사용(배당 미보정) → 키움/LS와 동일 기준 (adj_ok="0")
#    - 키움 엑셀과 동일한 한글 컬럼(날짜/시가/고가/저가/종가/거래량)으로 변환
#    - 전략 로직은 날짜/시가/종가만 사용하므로 결과는 기존과 동일
# =========================
FDR_LOOKBACK_DAYS = 8000   # 과거 조회 범위(달력일 기준). 충분히 크게 두면 상장 이후 전체를 가져옴
ADJ_OK = "0"               # "0"=분할만 보정(키움/LS 동일), "1"=배당까지 수정주가
DROP_TODAY_BAR = True      # 미장 진행중 당일봉(오늘) 제외

def load_data_via_common(ticker):
    """LS_Common_new.GetOhlcv(source='fdr')로 받아 한글 컬럼 DataFrame으로 변환."""
    df_std = LS.GetOhlcv("US", ticker, limit=FDR_LOOKBACK_DAYS, adj_ok=ADJ_OK, source="fdr")
    if df_std is None or len(df_std) == 0:
        raise RuntimeError(
            f"공통모듈에서 {ticker} 데이터를 받지 못했습니다. "
            f"(FinanceDataReader 설치/네트워크/티커명을 확인하세요)"
        )
    out = pd.DataFrame({
        '날짜':  pd.to_datetime(df_std.index),
        '시가':  pd.to_numeric(df_std['open'],   errors='coerce'),
        '고가':  pd.to_numeric(df_std['high'],   errors='coerce'),
        '저가':  pd.to_numeric(df_std['low'],    errors='coerce'),
        '종가':  pd.to_numeric(df_std['close'],  errors='coerce'),
        '거래량': pd.to_numeric(df_std['volume'], errors='coerce'),
    })
    out = out.dropna(subset=['종가'])
    out = out[out['종가'] > 0]
    out = out.drop_duplicates(subset=['날짜']).sort_values('날짜').reset_index(drop=True)
    if DROP_TODAY_BAR and len(out):
        today = pd.Timestamp(datetime.now().date())
        if out['날짜'].iloc[-1].normalize() == today:
            out = out.iloc[:-1].reset_index(drop=True)
    return out

print(f"[데이터] 공통모듈+FDR 실시간 로딩: {stock_name}")
data = load_data_via_common(stock_name)
data['날짜'] = pd.to_datetime(data['날짜'], errors='coerce')
data = data.dropna(subset=['날짜']).sort_values('날짜').reset_index(drop=True)
print(f"[데이터] {len(data)}봉 로딩 완료 ({data['날짜'].iloc[0].date()} ~ {data['날짜'].iloc[-1].date()})")

# ======================================
# ✅ 기간 확정 + 데이터 필터링 (가장 가까운 날짜로 자동 보정)
# ======================================
all_dates = sorted(pd.to_datetime(data['날짜']).dt.normalize().unique())
if not all_dates:
    raise ValueError("데이터에 유효한 날짜가 없습니다.")

_start_req = _parse_date_str(BACKTEST_START_STR)
_end_req   = _parse_date_str(BACKTEST_END_STR)

if _start_req is None:
    backtest_start = all_dates[0]
else:
    backtest_start = _nearest_date(all_dates, _start_req)

if _end_req is None:
    backtest_end = all_dates[-1]
else:
    backtest_end = _nearest_date(all_dates, _end_req)

# 시작/끝 뒤집힘 방지
if backtest_start > backtest_end:
    backtest_start, backtest_end = backtest_end, backtest_start

# 실제 필터링
data = data[(data['날짜'].dt.normalize() >= backtest_start) & (data['날짜'].dt.normalize() <= backtest_end)].reset_index(drop=True)
if data.empty:
    raise ValueError("설정한 기간에 해당하는 데이터가 없습니다. 날짜 범위를 확인하세요.")

print(f"[백테스팅 기간 확정] {pd.Timestamp(backtest_start).date()} ~ {pd.Timestamp(backtest_end).date()}")

# =========================
# 2) 매수/매도 원칙(그대로 유지)
# =========================
buy_threshold = -0.015            # 매수: 시가 대비 -1.5% 하락
sell_threshold = 0.025            # 매도: 매수가 대비 2.5% 상승
commission_rate = 0.0006          # 매도 시 수수료(값은 기존 그대로 유지)
holding_days_limit = 20           # 기본 최대 보유일수
lookback_period = 40              # 과거 고점 계산 기간 (30일)
decline_threshold = -0.4         # 고점 대비 하락률 -35% (기존 코드 그대로)
increase_factor_if_decline = 1.5  # 하락 조건 충족 시 수익률 조정
holding_days_limit_if_decline = 50  # 하락 조건 충족 시 최대 보유일수

# ✅ (중요) 기존 /7 하드코딩을 변수로 분리
PORTFOLIO_DIVISOR = 7

# =========================
# 3) 데이터 컬럼 초기화
# =========================
data['잔고'] = 100.0
data['평가잔고'] = 100.0  # ✅ MTM 평가잔고
data['최종 수익률'] = None
data['하락 조건 충족'] = ""

# ✅ FIX: int64로 생성되면 float 할당 시 에러 발생 → float로 생성
data['단순보유자산변동'] = 100.0

# 로그/가독성용(기존 로직 영향 없음)
data['차수'] = None
data['매수일'] = None
data['매수가'] = None
data['매도일'] = None
data['매도가'] = None
data['수익률'] = None

data['매수대상'] = None
data['매수사유'] = None
data['매도사유'] = None

data['매수시잔고'] = None
data['매도후잔고'] = None
data['잔고변화'] = None

# =========================
# 4) 포지션/잔고 변수
# =========================
max_count = 1
active_positions = {1: None}  # {차수: (buy_index, buy_date, buy_price, increase_factor, holding_limit)}
current_balance = 100.0

# (참고) 기존 코드에 있던 MDD 변수(실현기준) - 로직 영향 없음
peak_balance = current_balance
max_drawdown = 0.0

# 날짜별 활성 유닛 기록(기존과 동일)
position_history = []

# =========================
# 5) 매수/매도 조건 함수 (원칙 그대로)
# =========================
def should_buy(open_price, close_price):
    return (close_price - open_price) / open_price <= buy_threshold

def should_sell(buy_price, close_price):
    return (close_price - buy_price) / buy_price >= sell_threshold

# =========================
# 6) 사유 문자열 생성 (원칙 변경 없음)
# =========================
def fmt(x, nd=6):
    try:
        if x is None or pd.isna(x):
            return "—"
        return f"{float(x):.{nd}f}"
    except Exception:
        return "—"

def build_buy_reason(i, open_price, close_price, past_high, decline_rate, increase_factor, holding_limit, decline_ok):
    r = []
    r.append(f"매수조건: (종가-시가)/시가 <= {buy_threshold} → {fmt((close_price-open_price)/open_price, 6)}")
    if past_high is not None and not pd.isna(past_high):
        r.append(f"30일고점={fmt(past_high, 4)}, 하락률={(decline_rate*100):.2f}% (기준 {decline_threshold*100:.2f}%)")
    else:
        r.append("30일고점: 데이터 부족(초기구간)")
    r.append(f"하락조건충족={'O' if decline_ok else 'X'} → increase_factor={increase_factor}, holding_limit={holding_limit}")
    r.append(f"수익률반영: profit_rate*increase_factor 후 /{PORTFOLIO_DIVISOR} 적용")
    return " | ".join(r)

def build_sell_reason(buy_price, close_price, days_held, holding_limit):
    hit_profit = should_sell(buy_price, close_price)
    hit_days = days_held >= holding_limit
    parts = []
    parts.append(f"매도가(수수료전)={fmt(close_price,4)}")
    parts.append(f"목표수익률: (종가-매수가)/매수가 >= {sell_threshold} → {fmt((close_price-buy_price)/buy_price, 6)} ({'충족' if hit_profit else '미충족'})")
    parts.append(f"보유일: {days_held}일 / limit {holding_limit}일 ({'초과' if hit_days else '미초과'})")
    parts.append("매도트리거: " + ("목표수익률" if hit_profit else "") + (" + " if (hit_profit and hit_days) else "") + ("보유일초과" if hit_days else ""))
    return " | ".join(parts)

# =========================
# 7) 평가잔고(MTM) 계산 (원칙 변경 없음)
# =========================
def compute_mtm_balance(date_idx, current_balance, active_positions):
    """
    열린 포지션들을 '오늘 종가로 가정 청산'했을 때의 평가잔고.
    - 매도수수료 반영
    - 최종수익률 계산식은 실현과 동일: profit_rate -> /PORTFOLIO_DIVISOR
    - 여러 포지션은 곱셈 누적(=하루에 여러 건 청산 시 compounding과 동일)
    """
    close_today = data.loc[date_idx, '종가']
    if pd.isna(close_today):
        return current_balance

    factor = 1.0
    for _, buy_info in active_positions.items():
        if buy_info is None:
            continue
        buy_index, buy_date, buy_price, increase_factor, holding_limit = buy_info

        # 가정 청산가(수수료 반영)
        net_sell = close_today * (1.0 - commission_rate)

        # 실현과 동일한 수익률 계산식
        profit_rate = ((net_sell - buy_price) / buy_price) * 100.0 * increase_factor
        final_profit_rate = profit_rate / PORTFOLIO_DIVISOR

        factor *= (1.0 + final_profit_rate / 100.0)

    return current_balance * factor

# =========================
# 8) 백테스팅 루프 (매수/매도 원칙 그대로)
# =========================
initial_close = None  # ✅ 안전 초기화

for i in range(len(data)):
    open_price = data.loc[i, '시가']
    close_price = data.loc[i, '종가']
    date = data.loc[i, '날짜']

    # -------------------------
    # (1) 매도 로직: 열린 포지션 모두 체크 (원칙 동일)
    # -------------------------
    for count, buy_info in list(active_positions.items()):
        if buy_info is None:
            continue

        buy_index, buy_date, buy_price, buy_increase_factor, buy_holding_limit = buy_info
        days_held = (pd.to_datetime(date) - pd.to_datetime(buy_date)).days

        if should_sell(buy_price, close_price) or days_held >= buy_holding_limit:
            # 매도 기록은 "매수행(buy_index)"에 저장 (기존 방식 유지)
            net_sell_price = close_price * (1.0 - commission_rate)

            data.loc[buy_index, '매도일'] = date
            data.loc[buy_index, '매도가'] = net_sell_price

            # 매도사유 기록
            data.loc[buy_index, '매도사유'] = build_sell_reason(buy_price, close_price, days_held, buy_holding_limit)

            # 수익률 및 잔고 계산 (기존 공식 유지)
            profit_rate = ((net_sell_price - buy_price) / buy_price) * 100.0 * buy_increase_factor
            data.loc[buy_index, '수익률'] = profit_rate

            final_profit_rate = profit_rate / PORTFOLIO_DIVISOR
            data.loc[buy_index, '최종 수익률'] = final_profit_rate

            # 잔고 업데이트 (기존 방식 유지)
            profit_amount = current_balance * (final_profit_rate / 100.0)
            current_balance += profit_amount

            # (참고) 기존 코드 MDD 변수 업데이트(실현기준) - 로직 동일
            peak_balance = max(peak_balance, current_balance)
            drawdown = (current_balance - peak_balance) / peak_balance if peak_balance > 0 else 0.0
            max_drawdown = min(max_drawdown, drawdown)

            # 매도 후 잔고/잔고변화 기록(매수행 기준)
            data.loc[buy_index, '매도후잔고'] = current_balance
            if pd.notna(data.loc[buy_index, '매수시잔고']):
                data.loc[buy_index, '잔고변화'] = current_balance - float(data.loc[buy_index, '매수시잔고'])

            # 포지션 종료
            active_positions[count] = None

    # -------------------------
    # (2) 매수 로직 (원칙 동일)
    # -------------------------
    if should_buy(open_price, close_price):
        # 30일 고점/하락률 체크(원칙 동일)
        past_high = None
        decline_rate = None

        if i >= lookback_period:
            past_high = data.loc[i - lookback_period:i - 1, '종가'].max()
            decline_rate = (close_price - past_high) / past_high if past_high and past_high != 0 else np.nan

            decline_ok = (decline_rate <= decline_threshold) if not pd.isna(decline_rate) else False
            increase_factor = increase_factor_if_decline if decline_ok else 1
            holding_limit = holding_days_limit_if_decline if decline_ok else holding_days_limit
            data.loc[i, '하락 조건 충족'] = "O" if decline_ok else ""
        else:
            decline_ok = False
            increase_factor = 1
            holding_limit = holding_days_limit
            data.loc[i, '하락 조건 충족'] = ""

        # 사용 가능한 차수 찾기(원칙 동일)
        allocated = False
        for count in range(1, max_count + 1):
            if active_positions.get(count) is None:
                data.loc[i, '차수'] = count
                data.loc[i, '매수일'] = date
                data.loc[i, '매수가'] = close_price

                # 매수 시점 잔고 기록
                data.loc[i, '매수시잔고'] = current_balance

                # 매수사유/대상 기록 (원칙 변경 없음)
                data.loc[i, '매수대상'] = stock_name
                data.loc[i, '매수사유'] = build_buy_reason(i, open_price, close_price, past_high, decline_rate, increase_factor, holding_limit, decline_ok)

                active_positions[count] = (i, date, close_price, increase_factor, holding_limit)
                allocated = True
                break

        # 없다면 차수 추가(원칙 동일)
        if not allocated:
            max_count += 1
            data.loc[i, '차수'] = max_count
            data.loc[i, '매수일'] = date
            data.loc[i, '매수가'] = close_price

            data.loc[i, '매수시잔고'] = current_balance

            data.loc[i, '매수대상'] = stock_name
            data.loc[i, '매수사유'] = build_buy_reason(i, open_price, close_price, past_high, decline_rate, increase_factor, holding_limit, decline_ok)

            active_positions[max_count] = (i, date, close_price, increase_factor, holding_limit)

    # -------------------------
    # (3) 단순보유자산변동(기존 동일) - ✅ FIX: dtype/NaN 안전 처리
    # -------------------------
    if i == 0:
        initial_close = close_price

    if pd.notna(close_price) and pd.notna(initial_close) and float(initial_close) != 0.0:
        data.loc[i, '단순보유자산변동'] = 100.0 * (float(close_price) / float(initial_close))
    else:
        data.loc[i, '단순보유자산변동'] = np.nan

    # -------------------------
    # (4) 당일 잔고 기록(기존 동일)
    # -------------------------
    data.loc[i, '잔고'] = float(current_balance)

    # -------------------------
    # (5) 평가잔고 기록(추가)
    # -------------------------
    data.loc[i, '평가잔고'] = float(compute_mtm_balance(i, current_balance, active_positions))

    # -------------------------
    # (6) 날짜별 활성 차수 기록(기존 동일)
    # -------------------------
    weighted_units = 0.0
    active_count = 0
    for buy_info in active_positions.values():
        if buy_info is not None:
            _, _, _, inc_factor, _ = buy_info
            weighted_units += inc_factor
            active_count += 1
    position_history.append((date, weighted_units, active_count))

# =========================
# 9) 연산용 잔고(기존 안전장치 유지)
# =========================
data['잔고_연산용'] = pd.to_numeric(data['잔고'], errors='coerce')
data['잔고_연산용'] = data['잔고_연산용'].replace(0, pd.NA).ffill()
data['잔고_연산용'] = data['잔고_연산용'].fillna(100.0)

data['평가잔고_연산용'] = pd.to_numeric(data['평가잔고'], errors='coerce')
data['평가잔고_연산용'] = data['평가잔고_연산용'].replace(0, pd.NA).ffill()
data['평가잔고_연산용'] = data['평가잔고_연산용'].fillna(100.0)

# =========================
# 10) MDD 계산 (V3와 동일: 시계열 cummax 기반)
# =========================
# (A) 실현잔고 전체 MDD
real_peak = data['잔고_연산용'].cummax()
real_dd = (data['잔고_연산용'] - real_peak) / real_peak
overall_mdd = real_dd.min() * 100.0
overall_mdd_date = data.loc[real_dd.idxmin(), '날짜'] if not real_dd.empty else None

# (B) 평가잔고 전체 평가MDD
eval_peak = data['평가잔고_연산용'].cummax()
eval_dd = (data['평가잔고_연산용'] - eval_peak) / eval_peak
overall_mdd_eval = eval_dd.min() * 100.0
overall_mdd_eval_date = data.loc[eval_dd.idxmin(), '날짜'] if not eval_dd.empty else None

# =========================
# 11) 연도별 Summary (수익률 + 최대유닛 + 연도별 MDD/평가MDD + 전체MDD정보)
# =========================
data['년도'] = data['날짜'].dt.year
summary_rows = []

for year, yearly_data in data.groupby('년도'):
    if yearly_data.empty:
        continue

    start_balance = yearly_data.iloc[0]['잔고_연산용']
    end_balance = yearly_data.iloc[-1]['잔고_연산용']
    profit_rate = ((end_balance - start_balance) / start_balance) * 100 if start_balance and start_balance != 0 else float('nan')

    # 연도별 최대 유닛/시점 (기존 로직 유지)
    max_units = 0.0
    max_units_date = None
    max_units_count = 0
    for dt, weighted_units, active_count in position_history:
        if dt.year == year and weighted_units > max_units:
            max_units = weighted_units
            max_units_date = dt
            max_units_count = active_count

    # 연도별 MDD(실현)
    y_peak = yearly_data['잔고_연산용'].cummax()
    y_dd = (yearly_data['잔고_연산용'] - y_peak) / y_peak
    y_mdd = y_dd.min() * 100.0
    y_mdd_date = yearly_data.loc[y_dd.idxmin(), '날짜'] if not y_dd.empty else None

    # 연도별 평가MDD
    y_peak_eval = yearly_data['평가잔고_연산용'].cummax()
    y_dd_eval = (yearly_data['평가잔고_연산용'] - y_peak_eval) / y_peak_eval
    y_mdd_eval = y_dd_eval.min() * 100.0
    y_mdd_eval_date = yearly_data.loc[y_dd_eval.idxmin(), '날짜'] if not y_dd_eval.empty else None

    summary_rows.append({
        '년도': year,
        '시작 잔고': start_balance,
        '마지막 잔고': end_balance,
        '수익률': profit_rate,
        '최대유닛개수 시점': max_units_date,
        '최대유닛개수 및 차수': f"{max_units} (차수: {max_units_count})",

        '년도별 MDD 갱신 시점': y_mdd_date,
        '년도별 MDD': y_mdd,

        '년도별 평가MDD 갱신 시점': y_mdd_eval_date,
        '년도별 평가MDD': y_mdd_eval,

        '전체 MDD 갱신 시점': overall_mdd_date,
        '전체 MDD': overall_mdd,

        '전체 평가MDD 갱신 시점': overall_mdd_eval_date,
        '전체 평가MDD': overall_mdd_eval
    })

summary = pd.DataFrame(summary_rows)

# =========================
# 12) CAGR / 단순보유 CAGR
# =========================
start_date = data['날짜'].min()
end_date = data['날짜'].max()
years = (end_date - start_date).days / 365.25 if start_date != end_date else 0.0

initial_balance_bt = data['잔고_연산용'].iloc[0]
final_balance_bt = data['잔고_연산용'].iloc[-1]
if years <= 0 or initial_balance_bt == 0 or pd.isna(initial_balance_bt):
    cagr_bt = float('nan')
else:
    cagr_bt = ((final_balance_bt / initial_balance_bt) ** (1 / years) - 1) * 100

initial_hold = data['단순보유자산변동'].iloc[0]
final_hold = data['단순보유자산변동'].iloc[-1]
if years <= 0 or initial_hold == 0 or pd.isna(initial_hold):
    cagr_hold = float('nan')
else:
    cagr_hold = ((final_hold / initial_hold) ** (1 / years) - 1) * 100

# 단순보유 MDD
peak_hold = data['단순보유자산변동'].cummax()
dd_hold = (data['단순보유자산변동'] - peak_hold) / peak_hold
mdd_hold = dd_hold.min() * 100

# 연복리수익률_MDD 시트
mdd_cagr = pd.DataFrame({
    '지표': ['MDD', 'CAGR'],
    '백테스팅 결과 값': [overall_mdd, cagr_bt],       # ✅ 실현 시계열 MDD 기준
    '단순보유결과값': [mdd_hold, cagr_hold],
    '단위': ['%', '%']
})

# =========================
# 13) 매수매도로그 + 보유기간 평가MDD(트레이드별)
# =========================
eval_series = data[['날짜', '평가잔고_연산용']].dropna().copy()
eval_series = eval_series.sort_values('날짜').set_index('날짜')

def calc_window_mdd(series: pd.Series):
    if series is None or series.empty:
        return np.nan, None
    peak = series.cummax()
    dd = (series - peak) / peak
    mdd = dd.min() * 100.0
    mdd_date = dd.idxmin() if not dd.empty else None
    return mdd, mdd_date

trade_log = data[data['매수일'].notna()].copy()

# 보유일 계산
trade_log['보유일'] = (pd.to_datetime(trade_log['매도일']) - pd.to_datetime(trade_log['매수일'])).dt.days

# 보유기간 평가MDD 계산(매수일~매도일, 매도 없으면 마지막 날짜까지)
trade_log['보유기간 평가MDD(%)'] = np.nan
trade_log['보유기간 평가MDD 갱신일'] = None

for idx in trade_log.index:
    buy_dt = trade_log.loc[idx, '매수일']
    sell_dt = trade_log.loc[idx, '매도일']
    if pd.isna(buy_dt):
        continue
    if pd.isna(sell_dt):
        sell_dt = data['날짜'].iloc[-1]

    try:
        window = eval_series.loc[pd.to_datetime(buy_dt):pd.to_datetime(sell_dt), '평가잔고_연산용']
        mdd_w, mdd_w_date = calc_window_mdd(window)
        trade_log.loc[idx, '보유기간 평가MDD(%)'] = mdd_w
        trade_log.loc[idx, '보유기간 평가MDD 갱신일'] = mdd_w_date
    except Exception:
        pass

# 로그 컬럼 구성(동일 방식)
trade_log = trade_log[[
    '차수',
    '매수일', '매수가',
    '매수대상', '매수사유',
    '매도일', '매도가',
    '매도사유',
    '하락 조건 충족',
    '수익률', '최종 수익률',
    '매수시잔고', '매도후잔고', '잔고변화',
    '보유일',
    '보유기간 평가MDD(%)', '보유기간 평가MDD 갱신일',
    '잔고', '평가잔고'
]].copy()

# =========================
# 14) 그래프 저장
# =========================
plt.figure(figsize=(10, 6))
plt.plot(data['날짜'], data['잔고'], label='백테스팅 잔고(실현)')
plt.plot(data['날짜'], data['평가잔고'], label='평가잔고(MTM)', linestyle='--')
plt.plot(data['날짜'], data['단순보유자산변동'], label='단순보유자산변동')
plt.title(f'{stock_name} 잔고 변화 ({current_date})')
plt.xlabel('날짜')
plt.ylabel('자산')
plt.legend()
plt.grid(True)

plot_path = f'./{stock_name}_balance_plot_{current_date}.png'
plt.savefig(plot_path)
plt.close()

# =========================
# 15) 엑셀 저장
# =========================
output_path = f'./{stock_name}_떨사오팔_{current_date}.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    data.to_excel(writer, sheet_name='Data', index=False)
    summary.to_excel(writer, sheet_name='Summary', index=False)
    mdd_cagr.to_excel(writer, sheet_name='연복리수익률_MDD', index=False)
    trade_log.to_excel(writer, sheet_name='매수매도로그', index=False)

print(f"백테스팅 결과가 '{output_path}' 파일에 저장되었습니다.")
print(f"잔고 그래프가 '{plot_path}' 파일에 저장되었습니다.")























