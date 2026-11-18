# -*- coding: utf-8 -*-
# =========================================================
#  비트겟 떨사오팔 봇 — 수강생 배포용
# =========================================================
#  ※ 이 파일은 배포 버전입니다. (선생님 운영 서버본 기준)
#  ※ 시세·지표는 LS증권에서 읽고, 주문은 비트겟 USDT 무기한선물로 나갑니다.
#     그래서 LS 앱키와 비트겟 API 키가 ★둘 다★ 필요합니다.
#
#  [ 서버에 올릴 파일 ]
#   /var/autobot/bitget_Ejftkdhvkf/ 에 아래를 함께 올리세요.
#     - 비트겟떨사오팔_배포용.py   ← 이 파일
#     - LS공통모듈 폴더 안의 모든 .py
#         · LS_Common_new.py
#         · LS_API_Helper_US_new.py
#         · LS_API_Helper_KR_new.py
#         · telegram_alert.py
#         · bitget_keys.py
#
#  [ 사용 전 반드시 확인 ]
#   1) bitget_keys.py 를 열어 본인 비트겟 API 키를 넣으세요.
#      비워 두면 봇이 시작 즉시 멈춥니다(주문 안 나감). 발급 방법은 그 파일에 있습니다.
#         chmod 600 /var/autobot/bitget_Ejftkdhvkf/bitget_keys.py
#   2) telegram_alert.py 를 열어 본인 TOKEN / CHAT_ID 를 넣으세요.
#   3) myLS_StockInfo.yaml 에 본인의 LS 앱키·앱시크릿을 넣고
#      ★홈 디렉터리(/home/ec2-user)★ 에 두세요.
#      공통모듈이 ./myLS_StockInfo.yaml 상대경로로 읽기 때문에
#      반드시 홈에서 실행해야 합니다.
#         cd ~ && python3 -u /var/autobot/bitget_Ejftkdhvkf/비트겟떨사오팔_배포용.py
#      크론탭도 홈에서 돌아야 조건이 같습니다.
#   4) LS 계좌 모드가 SetChangeMode("REAL") 로 되어 있습니다.
#      본인 yaml 의 키 접두사(REAL_APP_KEY 등)와 맞추거나 바꾸세요.
#   5) 투자금은 배포본 기본값으로 낮춰 두었습니다.
#      → INITIAL_INVEST_USDT = 10 USDT (SOXL 최소주문 5의 2배)
#      ※ 비트겟 최소 주문금액은 5 USDT 입니다. 그 아래로 낮추면 주문이 거부됩니다.
#   6) 매매기록 엑셀 경로: /var/autobot/bitget_Ejftkdhvkf/비트겟떨사오팔_매매기록.xlsx
#   7) 거래 종목: SOXL
#
#  [ 이 봇의 기록 방식 ]
#   - 주문 접수 후 실제 체결가/체결수량을 조회해 엑셀에 기록합니다.
#   - 지정가가 제한시간 안에 체결되지 않으면 주문을 취소하고 알립니다.
#   - 매수·매도가 되지 않은 경우에도 그 사유(조건 미충족 내역)를 텔레그램으로 보냅니다.
#
#  [ 면책 ]
#   레버리지 선물 거래는 원금 전액 손실이 가능합니다.
#   투자 판단과 그 결과에 대한 책임은 사용자 본인에게 있습니다.
#   반드시 소액으로 충분히 검증한 뒤 사용하세요.
# =========================================================
import time
from datetime import datetime, timezone, timedelta
import pandas as pd
import pprint
import os
import json

import telegram_alert

# ✅ 데이터 조회/주문 모두 LS 사용
import LS_Common_new as LsCommon
import LS_API_Helper_US_new as LsiUS

# =========================
# ✅ Bitget 선물 주문 인프라 (신호=LS / 주문=Bitget / 엑셀=실체결값)
# =========================
import ccxt
from zoneinfo import ZoneInfo

# --- 시간 기준: 한국시간(KST) 및 미국동부시간(US_EASTERN, 썸머타임 자동 반영) ---
KST = ZoneInfo("Asia/Seoul")
US_EASTERN = ZoneInfo("America/New_York")
DAILY_CHECK_ET_HOUR = 15    # 미국 정규장 마감(16:00 ET) 10분 전 시각(시)
DAILY_CHECK_ET_MINUTE = 51  # 미국 정규장 마감(16:00 ET) 10분 전 시각(분)

# [구글 시트 연동 로거 추가]
from sheets_logger import SheetsLogger
sheet_logger = SheetsLogger(
    # webhook_url / auth_token 은 환경변수 GSHEET_WEBHOOK_URL / GSHEET_AUTH_TOKEN 에서 읽습니다 (소스에 하드코딩 금지)
    strategy_name="떨사오팔",
    exchange="Bitget"
)

# --- 비트겟 API (본인 키로 교체) ---

#   키는 서버의 /var/autobot/bitget_Ejftkdhvkf/bitget_keys.py 에서 읽는다.
#   Grid_Bitget.py / 비트겟V3.py 가 쓰는 것과 동일한 키이며, 같은 계정을 쓴다.
#   ★ 저장소·배포본에는 키를 두지 않는다. bitget_keys.py 는 서버 전용(600).
#     이 파일이 없으면 아래 플레이스홀더가 남아 검증 함수가
#     "설정되지 않았습니다" 로 즉시 걸러낸다.
try:
    from bitget_keys import BITGET_API_KEY, BITGET_SECRET, BITGET_PASSPHRASE
except ImportError:
    BITGET_API_KEY = "여기에_API_KEY"
    BITGET_SECRET = "여기에_SECRET"
    BITGET_PASSPHRASE = "여기에_PASSPHRASE"
BITGET_API_KEY = BITGET_API_KEY.strip()
BITGET_SECRET = BITGET_SECRET.strip()
BITGET_PASSPHRASE = BITGET_PASSPHRASE.strip()

# --- 심볼 (SOXL 선물) ---
BITGET_SYMBOL_CANDIDATES = ["SOXL/USDT:USDT", "SOXLUSDT"]

# --- 선물/지정가 파라미터 ---
MARGIN_MODE = "cross"
POSITION_SIDE = "long"
USE_LIMIT_ORDER = True
BUY_LIMIT_PRICE_MULTIPLIER = 1.008
SELL_LIMIT_PRICE_MULTIPLIER = 0.992
ORDER_WAIT_SEC = 2
ORDER_CHECK_COUNT = 5

# --- 실체결 회수 파라미터 ---
FILL_FETCH_RETRIES = 5
FILL_FETCH_DELAY = 0.5

def validate_api_strings():
    for name, value in {"API_KEY": BITGET_API_KEY, "SECRET": BITGET_SECRET, "PASSPHRASE": BITGET_PASSPHRASE}.items():
        if not value or "여기에" in value:
            raise RuntimeError(f"비트겟 {name} 가 설정되지 않았습니다. 본인 키로 교체하세요.")

def create_exchange():
    validate_api_strings()
    exchange = ccxt.bitget({
        "apiKey": BITGET_API_KEY,
        "secret": BITGET_SECRET,
        "password": BITGET_PASSPHRASE,
        "enableRateLimit": True,
        "options": {"defaultType": "swap"},
    })
    exchange.load_markets()
    return exchange

def resolve_symbol(exchange):
    for s in BITGET_SYMBOL_CANDIDATES:
        if s in exchange.markets:
            return s
    raise RuntimeError(f"Bitget에서 SOXL 선물 심볼을 찾지 못했습니다. 후보: {BITGET_SYMBOL_CANDIDATES}")

def get_current_futures_price(exchange, symbol) -> float:
    ticker = exchange.fetch_ticker(symbol)
    last = ticker.get("last")
    if last is None:
        raise RuntimeError("비트겟 현재가 조회 실패")
    return float(last)

def get_available_usdt(exchange) -> float:
    """비트겟 선물(swap) 계좌의 가용 USDT 잔고 조회"""
    try:
        bal = exchange.fetch_balance(params={"type": "swap"})
        free_usdt = None
        if isinstance(bal, dict):
            if "USDT" in bal and isinstance(bal["USDT"], dict):
                free_usdt = bal["USDT"].get("free")
            if free_usdt is None and "free" in bal and isinstance(bal["free"], dict):
                free_usdt = bal["free"].get("USDT")
        if free_usdt is not None:
            return float(free_usdt)
        return 0.0
    except Exception as e:
        print(f"[가용잔고 조회 경고] {e}")
        return -1.0  # 잔고 조회 실패 시 -1 반환 (주문 단계 예외처리로 위임)

def usdt_to_amount(exchange, symbol, usdt_amount, price):
    market = exchange.market(symbol)
    raw_amount = float(usdt_amount) / float(price)
    amount = float(exchange.amount_to_precision(symbol, raw_amount))
    min_amt = None
    try:
        min_amt = market.get("limits", {}).get("amount", {}).get("min")
    except Exception:
        pass
    if min_amt is not None and amount < float(min_amt):
        raise RuntimeError(f"주문수량이 최소수량보다 작습니다. 계산수량={amount}, 최소수량={min_amt}")
    if amount <= 0:
        raise RuntimeError("주문수량이 0 이하입니다.")
    return amount

def split_stage1_sell_quantity(exchange, symbol, total_qty, current_price, ratio=0.5):
    """
    1단계 목표가 통과 시 보유 수량을 ratio(기본 50%)만큼 분할 익절하기 위해
    1차 매도 수량(part1_qty)과 2단계 잔여 수량(rem_qty)을 거래소 정밀도에 맞게 계산합니다.
    최소 주문 수량(min_amt) 또는 최소 주문 금액(5 USDT) 미만이어서 분할이 불가능한 경우
    (0.0, float(total_qty), False)를 반환합니다.
    """
    try:
        market = exchange.market(symbol)
        min_amt = market.get("limits", {}).get("amount", {}).get("min")
        min_cost = market.get("limits", {}).get("cost", {}).get("min") or 5.0
    except Exception:
        min_amt = None
        min_cost = 5.0

    total_qty_f = float(total_qty)
    raw_part1 = total_qty_f * float(ratio)
    part1_qty = float(exchange.amount_to_precision(symbol, raw_part1))
    rem_qty = float(exchange.amount_to_precision(symbol, max(0.0, total_qty_f - part1_qty)))

    if part1_qty <= 0 or rem_qty <= 0:
        return 0.0, total_qty_f, False
    if min_amt is not None and (part1_qty < float(min_amt) or rem_qty < float(min_amt)):
        return 0.0, total_qty_f, False
    if float(current_price) > 0:
        if (part1_qty * float(current_price) < float(min_cost)) or (rem_qty * float(current_price) < float(min_cost)):
            return 0.0, total_qty_f, False

    return part1_qty, rem_qty, True

def price_to_precision(exchange, symbol, price):
    return float(exchange.price_to_precision(symbol, price))

def get_limit_buy_price(exchange, symbol):
    current_price = get_current_futures_price(exchange, symbol)
    return price_to_precision(exchange, symbol, current_price * BUY_LIMIT_PRICE_MULTIPLIER)

def get_limit_sell_price(exchange, symbol):
    current_price = get_current_futures_price(exchange, symbol)
    return price_to_precision(exchange, symbol, current_price * SELL_LIMIT_PRICE_MULTIPLIER)

def fetch_order_status_safe(exchange, order_id, symbol):
    try:
        return exchange.fetch_order(order_id, symbol)
    except Exception as e:
        print(f"[주문조회 실패] {e}")
        return None

def wait_order_filled(exchange, order_id, symbol, max_check_count=ORDER_CHECK_COUNT, wait_sec=ORDER_WAIT_SEC):
    last_info = None
    for i in range(max_check_count):
        order_info = fetch_order_status_safe(exchange, order_id, symbol)
        last_info = order_info
        if order_info:
            status = str(order_info.get("status", "")).lower()
            filled = float(order_info.get("filled") or 0)
            amount = float(order_info.get("amount") or 0)
            print(f"[주문체결확인 {i+1}/{max_check_count}] status={status}, filled={filled}, amount={amount}")
            if status in ("closed", "filled"):
                return True, order_info
            if amount > 0 and filled >= amount:
                return True, order_info
        time.sleep(wait_sec)
    return False, last_info

def cancel_order_safe(exchange, order_id, symbol):
    try:
        result = exchange.cancel_order(order_id, symbol)
        print("[미체결 주문 취소]", result)
        return result
    except Exception as e:
        print(f"[주문취소 실패] {e}")
        return None

def place_open_long(exchange, symbol, amount):
    params = {"reduceOnly": False, "marginMode": MARGIN_MODE, "posSide": POSITION_SIDE}
    if USE_LIMIT_ORDER:
        price = get_limit_buy_price(exchange, symbol)
        order = exchange.create_limit_buy_order(symbol, amount, price, params=params)
        return order, price
    order = exchange.create_market_buy_order(symbol, amount, params=params)
    return order, None

def place_close_long(exchange, symbol, amount):
    params = {"reduceOnly": True, "marginMode": MARGIN_MODE, "posSide": POSITION_SIDE}
    if USE_LIMIT_ORDER:
        price = get_limit_sell_price(exchange, symbol)
        order = exchange.create_limit_sell_order(symbol, amount, price, params=params)
        return order, price
    order = exchange.create_market_sell_order(symbol, amount, params=params)
    return order, None

def 체결가_조회(exchange, symbol, order_result, fallback_price, fallback_qty):
    """비트겟 주문의 실제 평균체결가/체결수량/체결금액 회수 (없으면 fetch_order 재시도 → fallback)."""
    if order_result:
        try:
            avg = order_result.get("average")
            filled = order_result.get("filled", 0) or 0
            cost = order_result.get("cost", 0) or 0
            if avg and float(avg) > 0 and float(filled) > 0:
                avg_f = float(avg); filled_f = float(filled)
                cost_f = float(cost) if cost else avg_f * filled_f
                return avg_f, filled_f, cost_f
        except Exception as e:
            print(f"[체결가 즉시 추출 실패] {e}")
    order_id = None
    try:
        if order_result:
            order_id = order_result.get("id")
    except Exception:
        pass
    if order_id:
        for attempt in range(1, FILL_FETCH_RETRIES + 1):
            try:
                time.sleep(FILL_FETCH_DELAY)
                fetched = exchange.fetch_order(order_id, symbol)
                if fetched:
                    avg = fetched.get("average")
                    filled = fetched.get("filled", 0) or 0
                    cost = fetched.get("cost", 0) or 0
                    if avg and float(avg) > 0 and float(filled) > 0:
                        avg_f = float(avg); filled_f = float(filled)
                        cost_f = float(cost) if cost else avg_f * filled_f
                        return avg_f, filled_f, cost_f
            except Exception as e:
                if attempt == FILL_FETCH_RETRIES:
                    print(f"[fetch_order 최종 실패 {attempt}/{FILL_FETCH_RETRIES}] {e}")
    print(f"⚠️ 체결가 조회 실패 → fallback 사용 (가격={fallback_price}, 수량={fallback_qty})")
    return float(fallback_price), float(fallback_qty), float(fallback_price) * float(fallback_qty)

def get_trend_ma_status(exchange, symbol):
    """
    비트겟 무기한선물 TREND_TIMEFRAME봉 기준
    직전 '완성봉(index -2)'의 종가가 TREND_MA_PERIOD MA 대비 노이즈 방지 버퍼(TREND_MA_BUFFER_PCT) 아래로 꺾였는지 판정.
    반환: (is_broken, prev_close, ma_val)
    """
    try:
        min_candles = TREND_MA_PERIOD + 2
        ohlcv = exchange.fetch_ohlcv(symbol, timeframe=TREND_TIMEFRAME, limit=TREND_OHLCV_LIMIT)
        if not ohlcv or len(ohlcv) < min_candles:
            print(f"[{TREND_TIMEFRAME} {TREND_MA_PERIOD}MA] 데이터 부족 ({len(ohlcv) if ohlcv else 0}개 수신, 최소 {min_candles}개 필요)")
            return False, 0.0, 0.0

        df_candle = pd.DataFrame(ohlcv, columns=["timestamp", "open", "high", "low", "close", "volume"])
        df_candle["close"] = pd.to_numeric(df_candle["close"], errors="coerce")

        # index -1 은 현재 실시간 진행 중인 미완성 캔들이므로 직전 완성봉(index -2) 기준 판정
        # 직전 완성봉 포함 최근 TREND_MA_PERIOD개 완성봉 구간
        completed_n = df_candle["close"].iloc[-(TREND_MA_PERIOD + 1):-1]
        ma_val = float(completed_n.mean())
        prev_close = float(df_candle["close"].iloc[-2])

        # 낮 시간대(미국 정규장 외) 횡보 노이즈로 인한 오작동 매도 방지를 위해 소폭 하단 버퍼 적용
        buf_pct = float(TREND_MA_BUFFER_PCT)
        buf_ratio = (buf_pct / 100.0) if buf_pct >= 0.05 else max(0.0, buf_pct)
        ma_break_line = ma_val * (1.0 - buf_ratio)

        is_broken = prev_close < ma_break_line
        print(
            f"[{TREND_TIMEFRAME} {TREND_MA_PERIOD}MA 분석] 직전완성봉 종가={prev_close:.4f}, "
            f"{TREND_MA_PERIOD}MA={ma_val:.4f}, 이탈기준선(-{buf_ratio*100:.2f}%)={ma_break_line:.4f}, 꺾임(이탈)={is_broken}"
        )
        return is_broken, prev_close, ma_val
    except Exception as e:
        print(f"[{TREND_TIMEFRAME} {TREND_MA_PERIOD}MA 조회 에러] {e}")
        return False, 0.0, 0.0
# ===== Bitget 인프라 끝 =====

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import RichTextProperties, Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.styles import Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows

# ✅ 모드 설정
LsCommon.SetChangeMode("REAL")    # LS 주문/데이터용

# =========================
# ✅ 모든 텔레그램 메시지 앞에 '비트겟_떨사오팔' 표식 (기존 LS 떨사오팔봇과 구분)
# =========================
_original_send = telegram_alert.SendMessage
def _prefixed_send(msg):
    return _original_send(f"🤖 [비트겟_떨사오팔]\n{msg}")
telegram_alert.SendMessage = _prefixed_send

# 통합 기록 파일 이름
FILENAME = "/var/autobot/bitget_Ejftkdhvkf/비트겟떨사오팔_매매기록.xlsx"
# PC에서 테스트할 때는 아래를 사용하세요
# FILENAME = "./비트겟떨사오팔_매매기록.xlsx"

# 차수별 2단계 상태 파일 이름
ROUND_STATE_FILE = os.path.join(os.path.dirname(FILENAME), "bitget_round_state.json") if os.path.dirname(FILENAME) else "./bitget_round_state.json"

# =========================
# ✅ 차수별 2단계 상태 관리 (JSON 기반 + 엑셀 동기화)
# =========================
def load_round_state():
    if not os.path.exists(ROUND_STATE_FILE):
        return {}
    try:
        with open(ROUND_STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[상태파일 로드 실패] {e}")
        return {}

def save_round_state(state):
    try:
        d = os.path.dirname(ROUND_STATE_FILE)
        if d and not os.path.exists(d):
            os.makedirs(d, exist_ok=True)
        with open(ROUND_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"[상태파일 저장 실패] {e}")

def _safe_nonneg_float(val, default=0.0):
    try:
        if val is None or pd.isna(val):
            return float(default)
        v = float(val)
        return v if v > 0 else float(default)
    except Exception:
        return float(default)

def get_offset_pool(stock_code) -> float:
    state = load_round_state()
    return _safe_nonneg_float(state.get(stock_code, {}).get("offset_pool_usdt", 0.0), 0.0)

def add_offset_pool(stock_code, amount_usdt) -> float:
    add_val = max(0.0, float(amount_usdt or 0.0))
    state = load_round_state()
    stock_dict = state.setdefault(stock_code, {})
    cur_pool = _safe_nonneg_float(stock_dict.get("offset_pool_usdt", 0.0), 0.0)
    new_pool = round(cur_pool + add_val, 6)
    stock_dict["offset_pool_usdt"] = new_pool
    stock_dict.setdefault("offset_total_used_usdt", 0.0)
    save_round_state(state)
    return new_pool

def consume_offset_pool(stock_code, needed_usdt):
    need_val = max(0.0, float(needed_usdt or 0.0))
    state = load_round_state()
    stock_dict = state.setdefault(stock_code, {})
    cur_pool = _safe_nonneg_float(stock_dict.get("offset_pool_usdt", 0.0), 0.0)
    used = min(cur_pool, need_val)
    new_pool = round(max(0.0, cur_pool - used), 6)
    stock_dict["offset_pool_usdt"] = new_pool
    prev_total_used = _safe_nonneg_float(stock_dict.get("offset_total_used_usdt", 0.0), 0.0)
    stock_dict["offset_total_used_usdt"] = round(prev_total_used + used, 6)
    save_round_state(state)
    return round(used, 6), new_pool

def sync_round_state_from_excel(df, stock_code):
    """
    상태 파일이 비어있거나 특정 차수가 누락된 경우,
    엑셀 매매기록의 미매도 차수들을 읽어 자동으로 상태를 복원/동기화(Self-healing).
    이미 stage=2 로 승격된 차수, 구제 적립금(offset_pool_usdt), 부분 상계 및 구제 모드 이력을 유실 없이 보존합니다.
    """
    state = load_round_state()
    stock_dict = state.setdefault(stock_code, {})
    stock_dict["offset_pool_usdt"] = _safe_nonneg_float(stock_dict.get("offset_pool_usdt", 0.0), 0.0)
    stock_dict["offset_total_used_usdt"] = _safe_nonneg_float(stock_dict.get("offset_total_used_usdt", 0.0), 0.0)
    rounds = stock_dict.setdefault("rounds", {})

    if df is not None and not df.empty and "종목" in df.columns:
        stock_data = df[df["종목"] == stock_code]

        # 1. 미매도 상태인 행(매도가가 없는 행)만 추출하여 매수가 기준 내림차순 정렬 및 1~N차 빈칸 없이 앞당김 정렬
        active_rows = stock_data[stock_data["매도가"].isna()].copy()
        excel_updated = False
        prev_rounds = dict(rounds)
        new_rounds = {}
        used_prev_keys = set()

        if not active_rows.empty:
            active_rows["_buy_price_num"] = pd.to_numeric(active_rows["매수가"], errors="coerce").fillna(0.0)
            active_rows["_orig_idx"] = active_rows.index
            active_rows = active_rows.sort_values(by=["_buy_price_num", "_orig_idx"], ascending=[False, True])

            for seq_idx, (idx, row) in enumerate(active_rows.iterrows(), start=1):
                old_r_num = int(row["차수"]) if pd.notna(row["차수"]) else seq_idx
                new_r_num = seq_idx
                if old_r_num != new_r_num:
                    df.loc[idx, "차수"] = new_r_num
                    excel_updated = True

                r_key = str(new_r_num)
                buy_price = float(row["매수가"])
                qty = float(row["수량"])
                target_price = buy_price * (1.0 + SELL_PROFIT_PCT)

                raw_date = row.get("날짜", None)
                if isinstance(raw_date, pd.Timestamp):
                    buy_date_str = raw_date.strftime("%Y-%m-%d")
                elif raw_date is not None and not pd.isna(raw_date):
                    s_dt = str(raw_date).strip()
                    buy_date_str = "" if s_dt.lower() in ("nat", "nan", "none", "") else s_dt[:10]
                else:
                    buy_date_str = ""

                excel_offset_qty = _safe_nonneg_float(row.get("구제상계수량", 0.0), 0.0)
                excel_offset_proceeds = _safe_nonneg_float(row.get("구제상계금액", 0.0), 0.0)

                # 기존 차수 상태를 (매수가, 수량) 기준으로 매칭하여 보존
                matched_prev = None
                old_key = str(old_r_num)
                if old_key in prev_rounds and old_key not in used_prev_keys:
                    p_info = prev_rounds[old_key]
                    p_init_qty = float(p_info.get("initial_quantity", p_info.get("quantity", 0.0)) or 0.0)
                    p_rem_qty = float(p_info.get("quantity", 0.0) or 0.0)
                    if abs(float(p_info.get("buy_price", 0.0)) - buy_price) < 0.05 and (abs(p_init_qty - qty) < 0.02 or abs(p_rem_qty - qty) < 0.02):
                        matched_prev = p_info
                        used_prev_keys.add(old_key)
                if matched_prev is None:
                    for pk, p_info in prev_rounds.items():
                        if pk in used_prev_keys:
                            continue
                        p_init_qty = float(p_info.get("initial_quantity", p_info.get("quantity", 0.0)) or 0.0)
                        p_rem_qty = float(p_info.get("quantity", 0.0) or 0.0)
                        if abs(float(p_info.get("buy_price", 0.0)) - buy_price) < 0.05 and (abs(p_init_qty - qty) < 0.02 or abs(p_rem_qty - qty) < 0.02):
                            matched_prev = p_info
                            used_prev_keys.add(pk)
                            break

                if matched_prev is not None:
                    prev_stage = int(matched_prev.get("stage", 1))
                    init_qty = float(matched_prev.get("initial_quantity", qty) or qty)
                    prev_part1_qty = float(matched_prev.get("part1_sold_qty", 0.0) or 0.0)
                    json_offset_qty = float(matched_prev.get("offset_sold_qty", 0.0) or 0.0)
                    prev_offset_qty = max(json_offset_qty, excel_offset_qty)
                    prev_offset_proceeds = max(float(matched_prev.get("offset_sold_proceeds", 0.0) or 0.0), excel_offset_proceeds)
                    calc_rem = round(max(0.0, init_qty - prev_part1_qty - prev_offset_qty), 8)
                    if prev_part1_qty > 0 or prev_offset_qty > 0:
                        if matched_prev.get("quantity") is not None and excel_offset_qty <= json_offset_qty + 1e-8:
                            effective_qty = float(matched_prev.get("quantity", calc_rem))
                        else:
                            effective_qty = calc_rem
                    else:
                        effective_qty = qty

                    prev_rescue_stage2 = bool(matched_prev.get("rescue_stage2", False))
                    prev_target = float(matched_prev.get("target_price", target_price) or target_price)
                    prev_pres = float(matched_prev.get("preservation_line", prev_target) or prev_target)

                    new_rounds[r_key] = {
                        "stage": prev_stage,
                        "buy_date": buy_date_str or str(matched_prev.get("buy_date", "") or ""),
                        "buy_price": buy_price,
                        "initial_quantity": init_qty,
                        "quantity": effective_qty,
                        "part1_sold_qty": prev_part1_qty,
                        "part1_sold_price": matched_prev.get("part1_sold_price", None),
                        "part1_sold_pct": matched_prev.get("part1_sold_pct", None),
                        "part1_sold_at": matched_prev.get("part1_sold_at", None),
                        "part1_pool_saved_ratio": float(matched_prev.get("part1_pool_saved_ratio", 0.0) or 0.0),
                        "target_price": prev_target if prev_rescue_stage2 else target_price,
                        "preservation_line": prev_pres,
                        "stage2_entered_at": matched_prev.get("stage2_entered_at", None),
                        "rescue_mode": bool(matched_prev.get("rescue_mode", False)),
                        "deep_loss_seen": bool(matched_prev.get("deep_loss_seen", False)),
                        "rescue_stage2": prev_rescue_stage2,
                        "offset_sold_qty": prev_offset_qty,
                        "offset_sold_proceeds": prev_offset_proceeds,
                        "offset_used_usdt": float(matched_prev.get("offset_used_usdt", 0.0) or 0.0),
                        "offset_count": int(matched_prev.get("offset_count", 0) or 0),
                    }
                else:
                    rem_qty = round(max(0.0, qty - excel_offset_qty), 8)
                    init_rescue = False
                    if buy_date_str and not ENABLE_HARD_TIMECUT:
                        try:
                            b_d = datetime.strptime(buy_date_str, "%Y-%m-%d").date()
                            held_d = (datetime.now(KST).date() - b_d).days
                            if held_d > MAX_HOLD_DAYS_NORMAL:
                                init_rescue = True
                        except Exception:
                            pass
                    new_rounds[r_key] = {
                        "stage": 1,
                        "buy_date": buy_date_str,
                        "buy_price": buy_price,
                        "initial_quantity": qty,
                        "quantity": rem_qty,
                        "part1_sold_qty": 0.0,
                        "part1_sold_price": None,
                        "part1_sold_pct": None,
                        "part1_sold_at": None,
                        "part1_pool_saved_ratio": 0.0,
                        "target_price": target_price,
                        "preservation_line": target_price,
                        "stage2_entered_at": None,
                        "rescue_mode": init_rescue,
                        "deep_loss_seen": False,
                        "rescue_stage2": False,
                        "offset_sold_qty": excel_offset_qty,
                        "offset_sold_proceeds": excel_offset_proceeds,
                        "offset_used_usdt": 0.0,
                        "offset_count": 1 if excel_offset_qty > 0 else 0,
                    }

        stock_dict["rounds"] = new_rounds

        if excel_updated:
            try:
                매매기록저장(df)
            except Exception as e:
                print(f"[차수 정렬 엑셀 자동 보정 저장 실패] {e}")

    save_round_state(state)
    return state

def update_round_stage(
    stock_code,
    round_number,
    stage,
    preservation_line=None,
    target_price=None,
    buy_price=None,
    quantity=None,
    initial_quantity=None,
    part1_sold_qty=None,
    part1_sold_price=None,
    part1_sold_pct=None,
    part1_sold_at=None,
    buy_date=None,
    part1_pool_saved_ratio=None,
    rescue_mode=None,
    deep_loss_seen=None,
    rescue_stage2=None,
    offset_sold_qty=None,
    offset_sold_proceeds=None,
    offset_used_usdt=None,
    offset_count=None,
):
    state = load_round_state()
    stock_dict = state.setdefault(stock_code, {})
    stock_dict.setdefault("offset_pool_usdt", 0.0)
    stock_dict.setdefault("offset_total_used_usdt", 0.0)
    rounds = stock_dict.setdefault("rounds", {})
    r_key = str(int(round_number))
    r_info = rounds.setdefault(r_key, {})

    r_info["stage"] = int(stage)
    if preservation_line is not None:
        r_info["preservation_line"] = float(preservation_line)
    if target_price is not None:
        r_info["target_price"] = float(target_price)
    if buy_price is not None:
        r_info["buy_price"] = float(buy_price)
    if quantity is not None:
        r_info["quantity"] = float(quantity)
    if initial_quantity is not None:
        r_info["initial_quantity"] = float(initial_quantity)
    if part1_sold_qty is not None:
        r_info["part1_sold_qty"] = float(part1_sold_qty)
    if part1_sold_price is not None:
        r_info["part1_sold_price"] = float(part1_sold_price)
    if part1_sold_pct is not None:
        r_info["part1_sold_pct"] = float(part1_sold_pct)
    if part1_sold_at is not None:
        r_info["part1_sold_at"] = str(part1_sold_at)
    if buy_date is not None:
        r_info["buy_date"] = str(buy_date)
    if part1_pool_saved_ratio is not None:
        r_info["part1_pool_saved_ratio"] = float(part1_pool_saved_ratio)
    if rescue_mode is not None:
        r_info["rescue_mode"] = bool(rescue_mode)
    if deep_loss_seen is not None:
        r_info["deep_loss_seen"] = bool(deep_loss_seen)
    if rescue_stage2 is not None:
        r_info["rescue_stage2"] = bool(rescue_stage2)
    if offset_sold_qty is not None:
        r_info["offset_sold_qty"] = float(offset_sold_qty)
    if offset_sold_proceeds is not None:
        r_info["offset_sold_proceeds"] = float(offset_sold_proceeds)
    if offset_used_usdt is not None:
        r_info["offset_used_usdt"] = float(offset_used_usdt)
    if offset_count is not None:
        r_info["offset_count"] = int(offset_count)
    if stage == 2 and not r_info.get("stage2_entered_at"):
        r_info["stage2_entered_at"] = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")

    save_round_state(state)
    return r_info

def remove_round_from_state(stock_code, round_number):
    state = load_round_state()
    r_key = str(int(round_number))
    if stock_code in state and "rounds" in state[stock_code]:
        if r_key in state[stock_code]["rounds"]:
            del state[stock_code]["rounds"][r_key]
            save_round_state(state)

def has_active_stage2_rounds(stock_code):
    """엑셀 및 상태 파일을 동기화한 뒤, 현재 2단계(stage == 2)에 진입한 미매도 차수가 1개 이상 존재하는지 확인합니다."""
    df = 매매기록불러오기()
    state = sync_round_state_from_excel(df, stock_code)
    rounds_dict = state.get(stock_code, {}).get("rounds", {})
    return any(int(r.get("stage", 1)) == 2 for r in rounds_dict.values())


# =========================
# ✅ 최초 투자금 (USDT) — 종목별 첫 매수 기준금액. 여기서만 수정하세요.
# =========================
INITIAL_INVEST_USDT = 150

# =========================
# ✅ 시장 무시 실행 토글
#   - True  : 미국장 개장 여부를 무시하고 무조건 실행 (장 마감/휴장에도 매수·매도 로직 실행)
#             → 비트겟은 24시간 거래되므로, 미국장과 무관하게 돌리고 싶을 때 사용
#   - False : 기존 동작. 미국장(LS 기준)이 열렸을 때만 실행, 닫혔으면 알림만 보내고 스킵
# =========================
IGNORE_MARKET_OPEN = True

# =========================
# ✅ 전략 설정
# =========================
BUY_DROP_PCT = 0.015
SELL_PROFIT_PCT = 0.025

MAX_HOLD_DAYS_NORMAL = 40
MAX_HOLD_DAYS_HIGH_DROP = 40  # 일반/대폭락 매수 모두 40일(41일차 구제 모드 전환) 기준으로 통일

# =========================
# ✅ 보유일별 목표 수익률 하향(B 전략) & 구제 적립금 상계(D 전략) 설정
# =========================
ENABLE_HARD_TIMECUT = False             # False: 기계적 강제 손절 폐지 -> 41일차(held_days > 40)부터 구제 모드 전환
STEPDOWN_BREAKEVEN_START_DAYS = 26       # 26~40일: 본전 탈출 구간 시작 경과일 (0~25일은 기존 +2.5% 적용)
STEPDOWN_BREAKEVEN_PCT = 0.002           # 26~40일 본전 탈출 목표(+0.2%) 및 구제 2단계 본전 래칫 잠금선
ENABLE_OFFSET_RESCUE = True             # D 전략: 하위 차수 익절 수익 적립 및 구제 상계 청산 활성화
RESCUE_REBOUND_TRIGGER_PCT = -0.10      # 41일~ 깊은 손실 차수의 반등 추세추종(stage=2) 진입 및 초기 보존선 (-10.0%)
OFFSET_PROFIT_SAVE_RATIO = 1.0          # 구제 대기 차수 보유 시 하위 익절 순수익 중 적립금 적립 비율 (1.0 = 100%)
MIN_OFFSET_ORDER_USDT = 10.0            # 부분 상계 주문 1회 최소 실행 금액 (비트겟 최소 5 USDT 대비 2배 안전마진)
MIN_REMAINING_DUST_USDT = 6.0           # 부분 상계 후 최소 잔여 평가금액 (자투리 미청산 방지)

# =========================
# ✅ 선물 계좌 보호용 매수 브레이크 설정
# =========================
MAX_ACTIVE_ROUNDS = 15                  # 최대 동시 보유 허용 차수 (초과 시 신규 매수 보류)
HEAVY_HOLD_ROUND_THRESHOLD = 10         # 다차수 과부하 기준 (10개 이상 보유 시 매수 하락 기준 강화)
HEAVY_HOLD_BUY_DROP_PCT = 0.030         # 다차수 과부하 구간 시가 대비 하락 매수 기준 (-3.0%)

# 고점 계산 기간 (2026-08-06: LS떨사오팔봇과 통일, 백테스팅 최적값)
LOOKBACK_DAYS = 40

HIGH_DROP_THRESHOLD = 0.40   # 고점대비 하락 판정 (2026-08-06: 0.35 → 0.40)
HIGH_DROP_INVEST_MULTIPLIER = 1.5   # 하락시 매수 증액 배수 (2026-08-06: 2.0 → 1.5)
HIGH_DROP_ADJ_WEIGHT = 1.5          # 하락시 복리 조정 가중치 (2026-08-06: 2.0 → 1.5)
ADJ_DIVISOR = 7.0

SELL_FEE_RATE = 0.0006  # 0.06%

# =========================
# ✅ 2단계 추세추종 및 1차 분할익절 설정
# =========================
TREND_TIMEFRAME = "1h"               # 추세 판정 캔들 시간프레임 (SOXL 변동성 대응: "4h" → "1h" 적용)
TREND_MA_PERIOD = 5                  # 이동평균선 기간 (예: 5 → 5MA, 10 → 10MA)
TREND_OHLCV_LIMIT = 10               # 추세 판정용 OHLCV 조회 캔들 수
TREND_MA_BUFFER_PCT = 0.004          # ★ 낮 시간대(장외 횡보) 1h 5MA 노이즈 방지용 버퍼 (0.004 = 0.4%: 완성봉 종가 < 5MA × 0.996 일 때 이탈 판정)
PRESERVATION_LINE_MULTIPLIER = 0.998  # 보존선 방어 매도 계수 (현재가 ≤ 보존선 × 이 값이면 매도)
STAGE1_PARTIAL_SELL_PCT = 50         # ★ 1차 목표가 통과 시 즉시 분할익절할 비율(%) — 여기서 자유롭게 조정하세요!
                                     #   (예: 50 → 50% 익절 후 50% 2단계, 30 → 30% 익절 후 70% 2단계, 0 → 분할 없이 100% 2단계, 100 → 100% 즉시 전량 익절)
                                     #   ※ 50(백분율) 또는 0.50(비율) 모두 자동 인식합니다.
STAGE2_BRIEFING_INTERVAL_HOURS = 4   # 2단계 진입 차수 보유 시 정기 브리핑 발송 간격 (시간 단위)

def get_stage1_partial_sell_config():
    """
    STAGE1_PARTIAL_SELL_PCT 설정값을 읽어 (비율 0.0~1.0, 1차익절_%문자열, 잔여_%문자열, 1차익절_%수치)를 반환합니다.
    예: STAGE1_PARTIAL_SELL_PCT = 50 (또는 0.5) -> (0.5, "50%", "50%", 50.0)
    """
    raw = float(STAGE1_PARTIAL_SELL_PCT)
    ratio = (raw / 100.0) if raw > 1.0 else raw
    ratio = max(0.0, min(1.0, ratio))
    pct_val = round(ratio * 100.0, 1)
    rem_val = round(100.0 - pct_val, 1)
    pct_str = f"{int(pct_val)}%" if pct_val.is_integer() else f"{pct_val:.1f}%"
    rem_str = f"{int(rem_val)}%" if rem_val.is_integer() else f"{rem_val:.1f}%"
    return ratio, pct_str, rem_str, pct_val

def get_round_target_profile(held_days, max_hold_days=MAX_HOLD_DAYS_NORMAL):
    """
    보유 경과일(held_days)에 따른 목표 수익률 및 매도 단계 프로필을 반환합니다.
    반환: (target_pct, allow_stage2, is_rescue_mode, phase_label)
      - 0 ~ 25일 : 정상(+2.5%), allow_stage2=True, is_rescue_mode=False
      - 26 ~ 40일: 본전구간(+0.2%), allow_stage2=False, is_rescue_mode=False
      - 41일 ~   : 구제모드, allow_stage2=False, is_rescue_mode=True
    """
    d = max(0, int(held_days))
    m_days = max(int(STEPDOWN_BREAKEVEN_START_DAYS), int(max_hold_days))
    if d < int(STEPDOWN_BREAKEVEN_START_DAYS):
        return float(SELL_PROFIT_PCT), True, False, f"정상(+{SELL_PROFIT_PCT*100:.1f}%)"
    if d <= m_days:
        return float(STEPDOWN_BREAKEVEN_PCT), False, False, f"본전구간(+{STEPDOWN_BREAKEVEN_PCT*100:.1f}%)"
    return float(STEPDOWN_BREAKEVEN_PCT), False, True, "구제모드"

# =========================
# ✅ 데이터 재시도 설정
# =========================
DATA_RETRY_COUNT = 5
DATA_RETRY_SLEEP_SEC = 4

# =========================
# ✅ 전략 전체 재실행 설정
# =========================
STRATEGY_RETRY_INTERVAL_SEC = 60
STRATEGY_RETRY_MAX_WAIT_SEC = 1800


# =========================
# ✅ 데이터 조회 재시도 래퍼
#    "LS 데이터"만 재시도 처리
# =========================
def _값이_비정상(v):
    try:
        if v is None:
            return True
        if isinstance(v, str):
            s = v.strip().lower()
            if s in ("", "none", "nan", "nat"):
                return True
        if pd.isna(v):
            return True
        fv = float(v)
        if fv <= 0:
            return True
        return False
    except Exception:
        return True


def _최근_close_유효한지(ohlcv):
    try:
        if ohlcv is None or len(ohlcv) == 0:
            return False
        if "close" not in ohlcv.columns:
            return False

        recent_close = ohlcv["close"].tail(30)
        recent_close = pd.to_numeric(recent_close, errors="coerce").dropna()
        recent_close = recent_close[recent_close > 0]

        return len(recent_close) > 0
    except Exception:
        return False


def _장체크용_ohlcv_유효성체크(ohlcv):
    try:
        if ohlcv is None:
            return False, "OHLCV가 None"
        if not isinstance(ohlcv, pd.DataFrame):
            return False, f"OHLCV 타입 이상: {type(ohlcv)}"
        if len(ohlcv) == 0:
            return False, "OHLCV가 비어 있음"

        if "open" not in ohlcv.columns:
            return False, "OHLCV에 open 컬럼이 없음"
        if "close" not in ohlcv.columns:
            return False, "OHLCV에 close 컬럼이 없음"

        last_open = pd.to_numeric(pd.Series([ohlcv["open"].iloc[-1]]), errors="coerce").iloc[0]
        last_close = pd.to_numeric(pd.Series([ohlcv["close"].iloc[-1]]), errors="coerce").iloc[0]

        if _값이_비정상(last_open):
            return False, f"시가 비정상값: {last_open}"
        if _값이_비정상(last_close):
            return False, f"종가 비정상값: {last_close}"

        return True, "OK"
    except Exception as e:
        return False, str(e)


# ✅ 원본 LS 데이터 함수 백업
_원본_GetCurrentPrice = LsiUS.GetCurrentPrice
_원본_GetOhlcv = LsiUS.GetOhlcv


def _get_yf_current_price(stock_code):
    """야후 차트 API를 직접 호출하여 최신 가격(regularMarketPrice 또는 직전 종가)을 반환합니다."""
    try:
        import requests
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{stock_code}"
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        res = requests.get(url, headers=headers, params={"range": "1d", "interval": "1m"}, timeout=10)
        if res.status_code == 200:
            result = res.json().get("chart", {}).get("result", [])
            if result:
                meta = result[0].get("meta", {})
                price = meta.get("regularMarketPrice") or meta.get("chartPreviousClose") or meta.get("previousClose")
                if price and not _값이_비정상(price):
                    return float(price)
                quotes = result[0].get("indicators", {}).get("quote", [{}])[0].get("close", [])
                valid_closes = [c for c in quotes if c is not None and not _값이_비정상(c)]
                if valid_closes:
                    return float(valid_closes[-1])
    except Exception as e:
        print(f"[야후 현재가 폴백 조회 오류] {e}")
    return None


def _GetCurrentPrice_with_retry(stock_code, *args, **kwargs):
    last_err = None

    for attempt in range(1, DATA_RETRY_COUNT + 1):
        try:
            current_price = _원본_GetCurrentPrice(stock_code, *args, **kwargs)

            if not _값이_비정상(current_price):
                return float(current_price)

            last_err = f"현재가 비정상값: {current_price}"

        except Exception as e:
            last_err = str(e)

        print(f"[LS 현재가 재시도 {attempt}/{DATA_RETRY_COUNT}] {stock_code} - {last_err}")
        time.sleep(DATA_RETRY_SLEEP_SEC)

    # ✅ LS 시세 조회 실패 시 (주말/휴장/LS API 일시 장애 대응) 야후 직접 HTTP 폴백
    yf_price = _get_yf_current_price(stock_code)
    if yf_price is not None:
        print(f"[시세소스] {stock_code} 현재가 ← 야후 직접HTTP 폴백 (LS 응답 없음 대응): {yf_price}")
        return float(yf_price)

    msg = (
        f"⚠️ LS 데이터 조회 실패\n"
        f"🔹 종목: {stock_code}\n"
        f"🔹 항목: 현재가\n"
        f"🔹 재시도: {DATA_RETRY_COUNT}회\n"
        f"🔹 마지막 오류/값: {last_err}\n"
        f"💡 LS 및 야후 가격 데이터가 불안정하여 이번 실행을 중단합니다."
    )
    telegram_alert.SendMessage(msg)
    raise RuntimeError(f"{stock_code} LS 현재가 조회 실패: {last_err}")


def _GetOhlcv_with_retry(*args, **kwargs):
    market = args[0] if len(args) > 0 else kwargs.get("market", "UNKNOWN")
    stock_code = args[1] if len(args) > 1 else kwargs.get("stock_code", "UNKNOWN")
    limit = args[2] if len(args) > 2 else kwargs.get("limit", None)
    last_err = None

    for attempt in range(1, DATA_RETRY_COUNT + 1):
        try:
            # 전략용 조회는 yfinance 우선(실패 시 LS 폴백).
            # LS 는 장중 당일 시가가 확정 전이라 최대 0.41% 어긋난다(2026-08 실측).
            # 장체크·프로브용 소량 조회(limit<=10)는 LS 기존 경로를 그대로 쓴다.
            _kw = dict(kwargs)
            if not (limit is not None and int(limit) <= 10):
                _kw.setdefault("source", "yf_auto")
            ohlcv = _원본_GetOhlcv(*args, **_kw)

            if (
                str(market).upper() == "US"
                and str(stock_code).upper() == "AAPL"
                and limit is not None
                and int(limit) <= 10
            ):
                is_valid, reason = _장체크용_ohlcv_유효성체크(ohlcv)
                if is_valid:
                    return ohlcv
                last_err = reason

            else:
                if ohlcv is None:
                    last_err = "OHLCV가 None"
                elif not isinstance(ohlcv, pd.DataFrame):
                    last_err = f"OHLCV 타입 이상: {type(ohlcv)}"
                elif len(ohlcv) == 0:
                    last_err = "OHLCV가 비어 있음"
                elif "open" not in ohlcv.columns:
                    last_err = "OHLCV에 open 컬럼이 없음"
                elif "close" not in ohlcv.columns:
                    last_err = "OHLCV에 close 컬럼이 없음"
                else:
                    last_open = ohlcv["open"].iloc[-1]

                    if _값이_비정상(last_open):
                        last_err = f"시가 비정상값: {last_open}"
                    elif not _최근_close_유효한지(ohlcv):
                        last_err = "최근 30개 close 데이터가 유효하지 않음"
                    else:
                        return ohlcv

        except Exception as e:
            last_err = str(e)

        print(f"[LS OHLCV 재시도 {attempt}/{DATA_RETRY_COUNT}] {market} {stock_code} - {last_err}")
        time.sleep(DATA_RETRY_SLEEP_SEC)

    msg = (
        f"⚠️ LS 데이터 조회 실패\n"
        f"🔹 종목: {stock_code}\n"
        f"🔹 항목: OHLCV(시가/open, 종가/close)\n"
        f"🔹 재시도: {DATA_RETRY_COUNT}회\n"
        f"🔹 마지막 오류/값: {last_err}\n"
        f"💡 LS 시가/과거 고점 계산이 불가능하여 이번 실행을 중단합니다."
    )
    telegram_alert.SendMessage(msg)
    raise RuntimeError(f"{stock_code} LS OHLCV 조회 실패: {last_err}")


# ✅ LS 데이터 함수만 재시도 버전으로 교체
LsiUS.GetCurrentPrice = _GetCurrentPrice_with_retry
LsCommon.GetOhlcv = _GetOhlcv_with_retry


# =========================
# ✅ 투자금 시트 관리
# =========================
INVEST_SHEET_NAME = "투자금"

def 투자금시트_초기화_및_불러오기(stock_code, default_invest_money):
    default_invest_money = int(default_invest_money)

    if not os.path.exists(FILENAME):
        wb = Workbook()

        ws_trade = wb.active
        ws_trade.title = "매매기록"
        ws_trade.append(["차수", "날짜", "종목", "매수가", "수량", "고점대비하락조건", "매도가", "매도일", "매도수량"])

        ws = wb.create_sheet(INVEST_SHEET_NAME)
        ws.append(["종목", "투자금", "업데이트일", "마지막수익률(%)", "조정률(%)", "다음투자금"])
        ws.append([stock_code, int(default_invest_money), datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"), "", "", int(default_invest_money)])

        wb.save(FILENAME)
        return int(default_invest_money)

    wb = load_workbook(FILENAME)

    if INVEST_SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(INVEST_SHEET_NAME)
        ws.append(["종목", "투자금", "업데이트일", "마지막수익률(%)", "조정률(%)", "다음투자금"])
        ws.append([stock_code, int(default_invest_money), datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"), "", "", int(default_invest_money)])
        wb.save(FILENAME)
        return int(default_invest_money)

    ws = wb[INVEST_SHEET_NAME]

    stock_row = None
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v == stock_code:
            stock_row = r
            break

    if stock_row is None:
        ws.append([stock_code, int(default_invest_money), datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"), "", "", int(default_invest_money)])
        wb.save(FILENAME)
        return int(default_invest_money)

    invest_val = ws.cell(row=stock_row, column=2).value
    try:
        invest_val = int(float(invest_val))
    except Exception:
        invest_val = int(default_invest_money)
        ws.cell(row=stock_row, column=2, value=int(invest_val))
        ws.cell(row=stock_row, column=3, value=datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"))
        ws.cell(row=stock_row, column=6, value=int(invest_val))
        wb.save(FILENAME)

    return int(invest_val)


def 투자금시트_업데이트(stock_code, new_invest_money, last_profit_pct, adj_pct):
    new_invest_money = int(new_invest_money)

    wb = load_workbook(FILENAME)
    if INVEST_SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(INVEST_SHEET_NAME)
        ws.append(["종목", "투자금", "업데이트일", "마지막수익률(%)", "조정률(%)", "다음투자금"])
    ws = wb[INVEST_SHEET_NAME]

    stock_row = None
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v == stock_code:
            stock_row = r
            break

    if stock_row is None:
        ws.append([stock_code, int(new_invest_money), datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"), float(last_profit_pct), float(adj_pct), int(new_invest_money)])
    else:
        ws.cell(row=stock_row, column=2, value=int(new_invest_money))
        ws.cell(row=stock_row, column=3, value=datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"))
        ws.cell(row=stock_row, column=4, value=float(last_profit_pct))
        ws.cell(row=stock_row, column=5, value=float(adj_pct))
        ws.cell(row=stock_row, column=6, value=int(new_invest_money))

    wb.save(FILENAME)


# =========================
# ✅ 매매 기록 로드/저장
# =========================
def 매매기록불러오기():
    try:
        wb = load_workbook(FILENAME, read_only=True)
        if "매매기록" in wb.sheetnames:
            df = pd.read_excel(FILENAME, engine="openpyxl", sheet_name="매매기록")
        else:
            df = pd.read_excel(FILENAME, engine="openpyxl", sheet_name=0)

        if "매도일" in df.columns:
            df["매도일"] = df["매도일"].astype(str)
        if "날짜" in df.columns:
            df["날짜"] = pd.to_datetime(df["날짜"], errors="coerce").dt.strftime("%Y-%m-%d")

        if "고점대비하락조건" in df.columns:
            def _to_bool(v):
                if isinstance(v, bool):
                    return v
                if v is None:
                    return False
                if isinstance(v, (int, float)) and not pd.isna(v):
                    return bool(int(v))
                s = str(v).strip().lower()
                if s in ("true", "t", "1", "y", "yes"):
                    return True
                if s in ("false", "f", "0", "n", "no", "", "none", "nan"):
                    return False
                return False

            df["고점대비하락조건"] = df["고점대비하락조건"].apply(_to_bool)

    except FileNotFoundError:
        print("매매 기록 파일이 없습니다.")
        df = pd.DataFrame(columns=["차수", "날짜", "종목", "매수가", "수량", "고점대비하락조건", "매도가", "매도일", "매도수량"])
    except Exception:
        df = pd.DataFrame(columns=["차수", "날짜", "종목", "매수가", "수량", "고점대비하락조건", "매도가", "매도일", "매도수량"])
    return df


def 매매기록저장(df):
    if "매도일" in df.columns:
        df["매도일"] = df["매도일"].astype(str)

    if os.path.exists(FILENAME):
        wb = load_workbook(FILENAME)
        trade_sheet_name = "매매기록"
        if trade_sheet_name not in wb.sheetnames:
            wb.create_sheet(trade_sheet_name, 0)
            wb.save(FILENAME)

        with pd.ExcelWriter(FILENAME, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name=trade_sheet_name)
    else:
        wb = Workbook()
        ws_trade = wb.active
        ws_trade.title = "매매기록"
        ws_trade.append(["차수", "날짜", "종목", "매수가", "수량", "고점대비하락조건", "매도가", "매도일", "매도수량"])
        wb.save(FILENAME)

        with pd.ExcelWriter(FILENAME, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, index=False, sheet_name="매매기록")

    try:
        손익차트시트_갱신(df)
    except Exception as e:
        print(f"손익차트 시트 갱신 실패: {e}")


# =========================
# ✅ 손익차트 시트 (날짜별 일일/누적 실현손익)
# =========================
PNL_SHEET_NAME = "손익차트"


def _실현손익_시리즈(df):
    """매도 완료 건만 골라 매도일 기준 실현손익(USDT)을 계산한다."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["매도일", "실현손익"])

    closed = df.copy()
    if "매도가" not in closed.columns:
        return pd.DataFrame(columns=["매도일", "실현손익"])

    closed = closed[pd.notna(closed["매도가"])].copy()
    if closed.empty:
        return pd.DataFrame(columns=["매도일", "실현손익"])

    buy = pd.to_numeric(closed.get("매수가"), errors="coerce")
    sell_qty = pd.to_numeric(closed.get("매도수량"), errors="coerce")
    if sell_qty.isna().all() and "수량" in closed.columns:
        sell_qty = pd.to_numeric(closed["수량"], errors="coerce")

    if "순매도가(수수료반영)" in closed.columns:
        net_sell = pd.to_numeric(closed["순매도가(수수료반영)"], errors="coerce")
    else:
        net_sell = pd.Series(float("nan"), index=closed.index)

    missing_net = net_sell.isna()
    if missing_net.any():
        raw_sell = pd.to_numeric(closed["매도가"], errors="coerce")
        net_sell = net_sell.where(~missing_net, raw_sell * (1.0 - SELL_FEE_RATE))

    closed["실현손익"] = (net_sell - buy) * sell_qty
    closed["매도일"] = pd.to_datetime(closed["매도일"], errors="coerce")
    closed = closed[closed["매도일"].notna() & closed["실현손익"].notna()]
    return closed[["매도일", "실현손익"]]


def 손익차트시트_갱신(df=None):
    """매매기록 기반 날짜별 일일/누적 실현손익 시트·차트를 갱신한다.
    일일 실현손익이 0인 날짜는 표·그래프에서 제외한다.
    """
    if df is None:
        df = 매매기록불러오기()

    pnl_rows = _실현손익_시리즈(df)
    open_cnt = 0
    if df is not None and not df.empty and "매도가" in df.columns:
        open_cnt = int(pd.isna(df["매도가"]).sum())

    col_daily = "일일 실현손익(USDT)"
    col_cum = "누적 실현손익(USDT)"

    if pnl_rows.empty:
        daily = pd.DataFrame(columns=["기간", col_daily, col_cum])
        total_pnl = 0.0
    else:
        daily_sum = pnl_rows.groupby(pnl_rows["매도일"].dt.normalize())["실현손익"].sum().sort_index()
        # 일일 손익 0인 날은 기간·그래프에서 제외
        daily_sum = daily_sum[daily_sum.abs() > 1e-12]
        daily = pd.DataFrame({
            "기간": daily_sum.index.strftime("%Y-%m-%d"),
            col_daily: daily_sum.values,
            col_cum: daily_sum.cumsum().values,
        })
        total_pnl = float(daily_sum.sum()) if len(daily_sum) else 0.0

    if os.path.exists(FILENAME):
        wb = load_workbook(FILENAME)
    else:
        wb = Workbook()
        wb.active.title = "매매기록"

    if PNL_SHEET_NAME in wb.sheetnames:
        del wb[PNL_SHEET_NAME]
    ws = wb.create_sheet(PNL_SHEET_NAME)

    ws["A1"] = "실현 손익 추이"
    ws["A1"].font = Font(bold=True, size=14, color="595959")
    ws["A2"] = "지금까지 누적 실현손익(USDT)"
    ws["B2"] = round(total_pnl, 4)
    ws["B2"].font = Font(bold=True, size=14, color="C00000")
    ws["B2"].number_format = "0.00"
    ws["A3"] = "청산 완료 건수"
    ws["B3"] = int(len(pnl_rows))
    ws["A4"] = "미청산(보유) 건수"
    ws["B4"] = open_cnt
    ws["A5"] = "기준시각(KST)"
    ws["B5"] = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")

    start_row = 7
    for r_idx, row in enumerate(dataframe_to_rows(daily, index=False, header=True), start=start_row):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == start_row:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
            elif c_idx >= 2 and isinstance(value, (int, float)):
                cell.number_format = "0.00"

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 22

    data_rows = len(daily)
    if data_rows > 0:
        cats = Reference(ws, min_col=1, min_row=start_row + 1, max_row=start_row + data_rows)
        data_daily = Reference(ws, min_col=2, min_row=start_row, max_row=start_row + data_rows)
        data_cum = Reference(ws, min_col=3, min_row=start_row, max_row=start_row + data_rows)

        bar = BarChart()
        bar.type = "col"
        bar.grouping = "clustered"
        bar.title = "실현 손익 추이 (단위: USDT)"
        bar.style = 10
        bar.width = 16
        bar.height = 12

        # X축 (세로 텍스트 회전으로 날짜 겹침 방지)
        bar.x_axis.delete = False
        bar.x_axis.axPos = "b"
        bar.x_axis.tickLblPos = "nextTo"
        bar.x_axis.majorTickMark = "out"
        bar.x_axis.txPr = RichText(
            bodyPr=RichTextProperties(
                anchor="ctr",
                anchorCtr="1",
                rot="-5400000",  # -90도 세로 회전
                spcFirstLastPara="1",
                vertOverflow="ellipsis",
                wrap="square",
            ),
            p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties()), endParaRPr=CharacterProperties())],
        )

        # Y축 (손익 금액 눈금 정보 명확히 표시)
        bar.y_axis.delete = False
        bar.y_axis.axPos = "l"
        bar.y_axis.tickLblPos = "nextTo"
        bar.y_axis.majorTickMark = "out"
        bar.y_axis.number_format = "#,##0"
        bar.y_axis.majorGridlines = ChartLines()

        bar.add_data(data_daily, titles_from_data=True)
        bar.set_categories(cats)
        bar.shape = 4
        bar.gapWidth = 80
        bar.legend.position = "r"

        if bar.series:
            bar.series[0].graphicalProperties.solidFill = "5B9BD5"

        line = LineChart()
        line.add_data(data_cum, titles_from_data=True)
        line.set_categories(cats)
        if line.series:
            s = line.series[0]
            s.graphicalProperties.line.solidFill = "C00000"
            s.graphicalProperties.line.width = 25000  # EMU
            s.marker.symbol = "circle"
            s.marker.size = 5

        bar += line
        ws.add_chart(bar, "E2")

    wb.save(FILENAME)
    return total_pnl, open_cnt


# =========================
# ✅ 전략 공통 함수
# =========================
def 과거_고점찾기(stock_code, ohlcv=None):
    if ohlcv is None:
        ohlcv = LsCommon.GetOhlcv("US", stock_code)
    if len(ohlcv) >= LOOKBACK_DAYS:
        recent_ohlcv = ohlcv[-LOOKBACK_DAYS:]
    else:
        recent_ohlcv = ohlcv
    highest_price = recent_ohlcv["close"].max()
    return highest_price


def 가능다음차수(stock_data):
    """
    빈 차수 없이 1~N차 앞당김 + 맨 뒤 차수 추가 방식:
    - 기존 미매도 보유 차수들이 1~N차로 빈칸 없이 채워지므로,
      신규 매수 차수는 항상 미매도 보유 개수 + 1 (보유 0개일 때 1차, 3개 보유 중일 때 4차)
    """
    if stock_data is None or stock_data.empty:
        return 1
    active_rows = stock_data[stock_data["매도가"].isna()]
    active_cnt = len(active_rows)
    if active_cnt < MAX_ACTIVE_ROUNDS:
        return active_cnt + 1
    return None


# =========================
def _has_rescue_waiting(rounds_map, exclude_rounds=None):
    """구제 대기 상태(rescue_mode=True, stage=1, 잔여수량>0)인 차수가 존재하는지 확인합니다."""
    excl = set()
    for er in (exclude_rounds or ()):
        try:
            excl.add(int(er))
        except Exception:
            excl.add(er)
    for rk, rv in (rounds_map or {}).items():
        try:
            rn = int(rk)
        except Exception:
            rn = rk
        if rn in excl:
            continue
        if bool(rv.get("rescue_mode", False)) and int(rv.get("stage", 1)) == 1:
            if float(rv.get("quantity", 0.0) or 0.0) > 0:
                return True
    return False


def execute_offset_rescue(exchange, symbol, stock_code, df, current_price, current_invest_money):
    """
    D 전략(트랙 B): 구제 적립금(offset_pool_usdt)을 사용하여
    구제 대기(rescue_mode=True, stage=1) 차수를 오래된 순(동률 시 고가 순)으로
    최소 주문금액(MIN_OFFSET_ORDER_USDT) 및 잔여 더스트(MIN_REMAINING_DUST_USDT)를 지켜가며
    부분 또는 전량 상계 청산합니다.
    """
    if not ENABLE_OFFSET_RESCUE or df is None or df.empty:
        return []

    try:
        market = exchange.market(symbol)
        min_amt = market.get("limits", {}).get("amount", {}).get("min")
        min_amt_f = float(min_amt) if min_amt is not None else 0.0
        min_cost = market.get("limits", {}).get("cost", {}).get("min")
        min_cost_f = max(5.0, float(min_cost)) if min_cost is not None else 5.0
    except Exception:
        min_amt_f = 0.0
        min_cost_f = 5.0

    min_order_usdt = max(min_cost_f, float(MIN_OFFSET_ORDER_USDT))
    min_dust_usdt = max(min_cost_f, float(MIN_REMAINING_DUST_USDT))
    offset_events = []

    for _ in range(20):
        pool_usdt = get_offset_pool(stock_code)
        if pool_usdt <= 0.01:
            break

        state = sync_round_state_from_excel(df, stock_code)
        rounds_dict = state.get(stock_code, {}).get("rounds", {})
        stock_data = df[df["종목"] == stock_code]
        active_rows = stock_data[stock_data["매도가"].isna()]
        if active_rows.empty:
            break

        candidates = []
        for idx, row in active_rows.iterrows():
            r_num = int(row["차수"])
            r_key = str(r_num)
            r_info = rounds_dict.get(r_key, {})
            if not bool(r_info.get("rescue_mode", False)):
                continue
            if int(r_info.get("stage", 1)) != 1:
                continue
            rem_qty = float(r_info.get("quantity", row["수량"]) or 0.0)
            if rem_qty <= 0:
                continue
            try:
                b_date = (
                    row["날짜"].date()
                    if isinstance(row["날짜"], pd.Timestamp)
                    else datetime.strptime(str(row["날짜"])[:10], "%Y-%m-%d").date()
                )
            except Exception:
                b_date = datetime.now(KST).date()
            b_price = float(row["매수가"])
            candidates.append((b_date, -b_price, idx, r_num, b_price, rem_qty, r_info))

        if not candidates:
            break

        # 집중 소각 원칙: 매수일 가장 오래된 순 -> 매수가 가장 높은 최상단 차수 우선
        candidates.sort(key=lambda x: (x[0], x[1]))
        b_date, _, idx, r_num, buy_price_f, rem_qty, r_info = candidates[0]

        loss_per_share = buy_price_f - (float(current_price) * (1.0 - SELL_FEE_RATE))
        if loss_per_share <= 0:
            sell_qty = float(exchange.amount_to_precision(symbol, rem_qty))
            if sell_qty <= 0 or (min_amt_f > 0 and sell_qty < min_amt_f) or (sell_qty * float(current_price) < min_cost_f):
                print(f"[구제 상계 대기] {stock_code} {r_num}차 최소 주문수량/금액 미달 (수량={sell_qty})")
                break
            is_full_close = True
        else:
            raw_offset_qty = pool_usdt / loss_per_share
            if raw_offset_qty >= rem_qty - 1e-8:
                sell_qty = float(exchange.amount_to_precision(symbol, rem_qty))
                if sell_qty <= 0 or (min_amt_f > 0 and sell_qty < min_amt_f) or (sell_qty * float(current_price) < min_cost_f):
                    print(f"[구제 전량 상계 불가] {stock_code} {r_num}차 최소주문 조건 미달 (수량={sell_qty}, 금액={sell_qty * float(current_price):.2f})")
                    break
                is_full_close = True
            else:
                part_qty = float(exchange.amount_to_precision(symbol, raw_offset_qty))
                if part_qty * loss_per_share > pool_usdt + 1e-6 and part_qty > 0:
                    fit_ratio = (pool_usdt / (part_qty * loss_per_share)) * 0.995
                    part_qty = float(exchange.amount_to_precision(symbol, part_qty * fit_ratio))

                after_rem = round(max(0.0, rem_qty - part_qty), 8)
                if after_rem > 0 and (
                    after_rem * float(current_price) < min_dust_usdt
                    or (min_amt_f > 0 and after_rem < min_amt_f)
                ):
                    safe_target_rem = max(
                        min_dust_usdt / float(current_price) if float(current_price) > 0 else 0.0,
                        min_amt_f,
                    )
                    max_safe_part = max(0.0, rem_qty - safe_target_rem)
                    part_qty = float(exchange.amount_to_precision(symbol, max_safe_part)) if max_safe_part > 0 else 0.0
                    if part_qty > 0 and (rem_qty - part_qty) * float(current_price) < min_dust_usdt - 1e-6:
                        part_qty = float(exchange.amount_to_precision(symbol, max_safe_part * 0.98)) if max_safe_part * 0.98 > 0 else 0.0

                if (
                    part_qty <= 0
                    or (min_amt_f > 0 and part_qty < min_amt_f)
                    or (part_qty * float(current_price) < min_order_usdt)
                ):
                    print(
                        f"[구제 상계 대기] {stock_code} {r_num}차 적립금=${pool_usdt:.2f} -> "
                        f"가능상계={part_qty}주(약 ${part_qty * float(current_price):.2f} < 최소 ${min_order_usdt:.2f})"
                    )
                    break

                sell_qty = part_qty
                is_full_close = False

        order, order_price = place_close_long(exchange, symbol, sell_qty)
        print(f"[구제 상계 매도주문 ({r_num}차, {'전량' if is_full_close else '부분'})]", order)

        fallback_sell = float(order_price) if order_price is not None else float(current_price)
        order_id = order.get("id") if isinstance(order, dict) else None
        if order_id:
            filled_ok, _ = wait_order_filled(exchange, order_id, symbol)
            if not filled_ok and USE_LIMIT_ORDER:
                cancel_order_safe(exchange, order_id, symbol)
                telegram_alert.SendMessage(f"⚠️ 떨사오팔(Bitget) {stock_code} {r_num}차 구제 상계 지정가 미체결 → 취소")
                break

        real_price, real_qty, _ = 체결가_조회(exchange, symbol, order, fallback_sell, sell_qty)
        real_price_f = float(real_price)
        real_qty_f = float(real_qty)

        actual_loss = max(0.0, (buy_price_f - real_price_f * (1.0 - SELL_FEE_RATE)) * real_qty_f)
        used_pool, remain_pool = consume_offset_pool(stock_code, actual_loss)

        raw_left = max(0.0, rem_qty - real_qty_f)
        if raw_left <= 1e-8:
            new_rem_qty = 0.0
        else:
            new_rem_qty = float(exchange.amount_to_precision(symbol, raw_left))
        if new_rem_qty <= 0 or (min_amt_f > 0 and new_rem_qty < min_amt_f):
            is_full_close = True
            new_rem_qty = 0.0

        prev_off_qty = float(r_info.get("offset_sold_qty", 0.0) or 0.0)
        prev_off_proceeds = float(r_info.get("offset_sold_proceeds", 0.0) or 0.0)
        prev_off_used = float(r_info.get("offset_used_usdt", 0.0) or 0.0)
        prev_off_cnt = int(r_info.get("offset_count", 0) or 0)

        new_off_qty = round(prev_off_qty + real_qty_f, 8)
        new_off_proceeds = round(prev_off_proceeds + (real_price_f * real_qty_f), 8)
        new_off_used = round(prev_off_used + used_pool, 6)
        new_off_cnt = prev_off_cnt + 1

        init_qty = float(r_info.get("initial_quantity", df.at[idx, "수량"]) or df.at[idx, "수량"])
        part1_qty_done = float(r_info.get("part1_sold_qty", 0.0) or 0.0)
        part1_price_done = r_info.get("part1_sold_price", None)
        part1_proceeds = (
            float(part1_price_done) * part1_qty_done
            if (part1_qty_done > 0 and part1_price_done is not None)
            else 0.0
        )

        fee_usdt = real_price_f * real_qty_f * SELL_FEE_RATE
        this_net = real_price_f * (1.0 - SELL_FEE_RATE)
        this_pnl = (this_net - buy_price_f) * real_qty_f
        this_pct = ((this_net / buy_price_f) - 1.0) * 100.0 if buy_price_f > 0 else 0.0

        if not is_full_close:
            df.at[idx, "구제상계수량"] = new_off_qty
            df.at[idx, "구제상계금액"] = new_off_proceeds

            update_round_stage(
                stock_code,
                r_num,
                stage=1,
                buy_price=buy_price_f,
                quantity=new_rem_qty,
                initial_quantity=init_qty,
                rescue_mode=True,
                deep_loss_seen=bool(r_info.get("deep_loss_seen", False)),
                rescue_stage2=False,
                offset_sold_qty=new_off_qty,
                offset_sold_proceeds=new_off_proceeds,
                offset_used_usdt=new_off_used,
                offset_count=new_off_cnt,
            )

            msg = (
                f"🛟 [구제 부분 상계 청산] {stock_code} {r_num}차 소각 진행\n"
                f"🔹 금회 상계: {real_qty_f}주 @ {real_price_f:.4f} (손실 {this_pct:+.2f}%)\n"
                f"🏦 적립금 보전: ${used_pool:.2f} USDT 차감 (잔여 적립금: ${remain_pool:.2f} USDT)\n"
                f"📊 차수 잔여 수량: {new_rem_qty}주 / 최초 {init_qty}주 (누적 상계 {new_off_qty}주)\n"
                f"📅 매수일: {b_date} | 매수가: {buy_price_f:.2f}"
            )
            telegram_alert.SendMessage(msg)

            sheet_logger.log_trade(
                action="SELL",
                symbol=symbol,
                level=f"{r_num}차(구제부분)",
                price=real_price_f,
                amount=real_qty_f,
                fee=fee_usdt,
                pnl=this_pnl,
                pnl_pct=this_pct,
                close_all=False,
                note=f"구제 적립금 부분 상계 (보전 ${used_pool:.2f}, 잔여 {new_rem_qty}주)",
            )
            offset_events.append((r_num, real_qty_f, real_price_f, False))
            break
        else:
            total_sold_qty = round(part1_qty_done + new_off_qty, 8)
            total_proceeds = part1_proceeds + new_off_proceeds
            avg_sell_price = (total_proceeds / total_sold_qty) if total_sold_qty > 0 else real_price_f
            net_sell = avg_sell_price * (1.0 - SELL_FEE_RATE)
            profit_pct = ((net_sell / buy_price_f) - 1.0) * 100.0 if buy_price_f > 0 else 0.0

            df.at[idx, "구제상계수량"] = new_off_qty
            df.at[idx, "구제상계금액"] = new_off_proceeds
            df.at[idx, "매도가"] = avg_sell_price
            df.at[idx, "매도일"] = datetime.now(KST).strftime("%Y-%m-%d")
            df.at[idx, "매도수량"] = total_sold_qty
            df.at[idx, "수수료율(매도)"] = 0.06
            df.at[idx, "순매도가(수수료반영)"] = net_sell
            df.at[idx, "수익률(%)"] = profit_pct
            df.at[idx, "조정률(수익률/7)(%)"] = 0.0
            df.at[idx, "다음투자금"] = current_invest_money

            remove_round_from_state(stock_code, r_num)
            sync_round_state_from_excel(df, stock_code)

            msg = (
                f"🛟 [구제 전량 상계 완결] {stock_code} {r_num}차 소각 완료!\n"
                f"🔹 최종 상계: {real_qty_f}주 @ {real_price_f:.4f} (합산 평균매도가 {avg_sell_price:.4f})\n"
                f"🏦 적립금 보전: 금회 ${used_pool:.2f} / 누적 ${new_off_used:.2f} USDT (잔여 적립금: ${remain_pool:.2f} USDT)\n"
                f"📊 총 청산 수량: {total_sold_qty}주 ({new_off_cnt}회 상계 완결)\n"
                f"📅 매수일: {b_date} | 매수가: {buy_price_f:.2f} | 합산수익률: {profit_pct:+.2f}%\n"
                f"💡 적립금으로 손실이 상쇄되어 다음 투자금({current_invest_money} USDT)은 감액되지 않았습니다."
            )
            telegram_alert.SendMessage(msg)

            sheet_logger.log_trade(
                action="SELL",
                symbol=symbol,
                level=f"{r_num}차(구제완결)",
                price=real_price_f,
                amount=real_qty_f,
                fee=fee_usdt,
                pnl=this_pnl,
                pnl_pct=this_pct,
                close_all=True,
                note=f"구제 적립금 전량 상계 완결 (보전 ${used_pool:.2f}, 합산수익률 {profit_pct:+.2f}%)",
            )
            offset_events.append((r_num, real_qty_f, real_price_f, True))

    if offset_events:
        매매기록저장(df)

    return offset_events


# =========================
# ✅ 매도 체크 (보유일별 목표하향 + 2단계 추세추종 + D전략 구제상계)
#    데이터=LS / 주문=Bitget 선물
# =========================
def 매도체크(exchange, symbol, stock_code, default_invest_money, ohlcv=None, is_daily_buy_turn=False):
    df = 매매기록불러오기()

    # 상태 파일 동기화 (기존 엑셀 미매도건 자동 복원 및 차수 앞당김)
    state = sync_round_state_from_excel(df, stock_code)
    stock_data = df[df["종목"] == stock_code]

    # ✅ 매도 감시 현재가는 Bitget 실시간 선물 시세 사용
    # (LS증권은 한국 장외 시간대 전산 리셋 시 전일 종가로 초기화되는 왜곡이 발생하여
    #  24시간 크론탭 기반 매도 감시에서 잘못된 청산을 유발할 수 있음.
    #  Bitget 선물은 24/7 연속 실시간 시세를 제공하므로 장외 시간대에도 안전함.)
    current_price = get_current_futures_price(exchange, symbol)
    today = datetime.now(KST).date()
    sell_orders = []
    partial_sell_events = []

    current_invest_money = 투자금시트_초기화_및_불러오기(stock_code, default_invest_money)
    highest_price_30d = 과거_고점찾기(stock_code, ohlcv)

    # 대시보드 STATE_SYNC 및 2단계(추세추종) 판정을 위해 1h 5MA 지표 조회 (보유 0개여도 조회)
    rounds_dict = state.get(stock_code, {}).get("rounds", {})
    is_ma5_broken, prev_close_4h, ma5_4h = get_trend_ma_status(exchange, symbol)

    # 0) 미매도 차수별 deep_loss_seen / rescue_mode 선행 동기화
    for _, pre_row in stock_data.iterrows():
        if pd.notna(pre_row.get("매도가", None)):
            continue
        try:
            pre_buy_date = (
                pre_row["날짜"].date()
                if isinstance(pre_row["날짜"], pd.Timestamp)
                else datetime.strptime(str(pre_row["날짜"])[:10], "%Y-%m-%d").date()
            )
        except Exception:
            continue
        pre_buy_price = float(pre_row["매수가"])
        pre_r_num = int(pre_row["차수"])
        pre_r_key = str(pre_r_num)
        pre_r_info = rounds_dict.get(pre_r_key, {})
        pre_stage = int(pre_r_info.get("stage", 1))
        pre_held_days = (today - pre_buy_date).days
        pre_max_hold = MAX_HOLD_DAYS_HIGH_DROP if bool(pre_row["고점대비하락조건"]) else MAX_HOLD_DAYS_NORMAL

        cur_deep = bool(pre_r_info.get("deep_loss_seen", False))
        cur_rescue = bool(pre_r_info.get("rescue_mode", False))
        new_deep = cur_deep or (pre_buy_price > 0 and current_price <= pre_buy_price * (1.0 + RESCUE_REBOUND_TRIGGER_PCT))
        new_rescue = cur_rescue or ((not ENABLE_HARD_TIMECUT) and pre_stage == 1 and pre_held_days > pre_max_hold)

        if new_deep != cur_deep or new_rescue != cur_rescue:
            rounds_dict[pre_r_key] = update_round_stage(
                stock_code,
                pre_r_num,
                stage=pre_stage,
                deep_loss_seen=new_deep,
                rescue_mode=new_rescue,
            )

    for index, row in stock_data.iterrows():
        if pd.notna(row.get("매도가", None)):
            continue

        try:
            if isinstance(row["날짜"], pd.Timestamp):
                buy_date = row["날짜"].date()
            else:
                buy_date = datetime.strptime(str(row["날짜"])[:10], "%Y-%m-%d").date()
        except ValueError:
            print(f"날짜 형식 오류: {row['날짜']}")
            continue

        buy_price = float(row["매수가"])
        excel_quantity = float(row["수량"])
        round_number = int(row["차수"])
        target_price = buy_price * (1.0 + SELL_PROFIT_PCT)
        max_hold_days = MAX_HOLD_DAYS_HIGH_DROP if row["고점대비하락조건"] else MAX_HOLD_DAYS_NORMAL
        held_days = (today - buy_date).days

        r_key = str(round_number)
        r_info = rounds_dict.get(r_key, {})
        current_stage = int(r_info.get("stage", 1))
        preservation_line = float(r_info.get("preservation_line", target_price))
        initial_quantity = float(r_info.get("initial_quantity", excel_quantity) or excel_quantity)
        part1_sold_qty = float(r_info.get("part1_sold_qty", 0.0) or 0.0)
        offset_sold_qty = float(r_info.get("offset_sold_qty", 0.0) or 0.0)
        if part1_sold_qty > 0 or offset_sold_qty > 0:
            calc_rem = max(0.0, initial_quantity - part1_sold_qty - offset_sold_qty)
            quantity = float(r_info.get("quantity", calc_rem) or calc_rem)
        else:
            quantity = excel_quantity
        if quantity <= 0:
            continue

        deep_loss_seen = bool(r_info.get("deep_loss_seen", False))
        rescue_stage2 = bool(r_info.get("rescue_stage2", False))

        # 1) 강제 타임컷 체크 (ENABLE_HARD_TIMECUT=True 일 때만 동작; False 면 구제 모드로 전환)
        if ENABLE_HARD_TIMECUT and is_daily_buy_turn and held_days >= max_hold_days:
            print(f"{stock_code} {round_number}차 떨사오팔 강제 매도 - {max_hold_days}일 경과")
            sell_orders.append((index, quantity, current_price * 0.98, f"보유기간 경과({max_hold_days}일)", buy_price, buy_date, round_number, "HARD_TIMECUT"))

            msg = (
                f"⚠️ 떨사오팔 일반 전략 {stock_code} {round_number}차 강제 매도 알림\n"
                f"🔹 사유: 보유기간 경과({max_hold_days}일)\n"
                f"📅 매수일: {buy_date} | 오늘: {today}\n"
                f"📊 매도 대상 수량: {quantity}주\n"
            )
            telegram_alert.SendMessage(msg)
            continue

        # 2) 1단계: 보유일별 동적 목표가 / 구제 반등 추세추종 감시 (일일 매수 점검 전 하루 1번만 확인)
        if current_stage == 1:
            if is_daily_buy_turn:
                target_pct, allow_stage2, is_rescue_phase, phase_label = get_round_target_profile(held_days, max_hold_days)

                # [구간 3] 41일차 이상 구제 모드 (트랙 A: 방식 ② 반등 추세추종)
                if is_rescue_phase and not ENABLE_HARD_TIMECUT:
                    be_price = buy_price * (1.0 + STEPDOWN_BREAKEVEN_PCT)
                    if deep_loss_seen:
                        rebound_price = buy_price * (1.0 + RESCUE_REBOUND_TRIGGER_PCT)
                        if current_price >= rebound_price:
                            if is_ma5_broken:
                                reason_str = (
                                    f"구제 반등선({RESCUE_REBOUND_TRIGGER_PCT*100:+.1f}%) 도달 + "
                                    f"{TREND_TIMEFRAME} {TREND_MA_PERIOD}MA 하회 전량 청산(완성봉 {prev_close_4h:.2f} < {TREND_MA_PERIOD}MA {ma5_4h:.2f})"
                                )
                                print(f"{stock_code} {round_number}차 {reason_str} -> 100% 반등 청산!")
                                sell_orders.append((index, quantity, current_price * 0.98, reason_str, buy_price, buy_date, round_number, "RESCUE_REBOUND"))
                            else:
                                init_pres = be_price if current_price >= be_price else rebound_price
                                updated_info = update_round_stage(
                                    stock_code,
                                    round_number,
                                    stage=2,
                                    preservation_line=init_pres,
                                    target_price=init_pres,
                                    buy_price=buy_price,
                                    quantity=quantity,
                                    initial_quantity=initial_quantity,
                                    rescue_mode=True,
                                    deep_loss_seen=True,
                                    rescue_stage2=True,
                                )
                                rounds_dict[r_key] = updated_info
                                cur_loss_pct = ((current_price / buy_price) - 1.0) * 100.0 if buy_price > 0 else 0.0
                                pres_pct = ((init_pres / buy_price) - 1.0) * 100.0 if buy_price > 0 else 0.0
                                print(f"{stock_code} {round_number}차 구제 반등선({rebound_price:.2f}) 회복 -> 100% 구제 2단계 승격(보존선 {init_pres:.2f})")
                                msg = (
                                    f"🛟 [구제 반등 추세추종 승격] {stock_code} {round_number}차 2단계 진입\n"
                                    f"🔹 반등 조건: {RESCUE_REBOUND_TRIGGER_PCT*100:+.1f}% 선({rebound_price:.2f}) 회복 & {TREND_TIMEFRAME} {TREND_MA_PERIOD}MA 상승 유지\n"
                                    f"📉 매수가: {buy_price:.2f} | 💲 현재가: {current_price:.2f} ({cur_loss_pct:+.2f}%)\n"
                                    f"🛡️ 초기 최저보존선: {init_pres:.2f} ({pres_pct:+.1f}%)\n"
                                    f"📊 잔여 물량(100%): {quantity}주 → 5분 간격 추세추종 감시 시작\n"
                                    f"💡 본전(+{STEPDOWN_BREAKEVEN_PCT*100:.1f}%, {be_price:.2f}) 돌파 시 보존선이 본전선으로 자동 상향 잠금됩니다."
                                )
                                telegram_alert.SendMessage(msg)
                        else:
                            print(f"{stock_code} {round_number}차 구제 대기 중(보유 {held_days}일): 현재가 {current_price:.2f} < 반등선 {rebound_price:.2f} (트랙 B 적립금 상계 대기)")
                    else:
                        if current_price >= be_price:
                            reason_str = f"구제 본전탈출선(+{STEPDOWN_BREAKEVEN_PCT*100:.1f}%) 도달(얕은 손실 보유 {held_days}일)"
                            print(f"{stock_code} {round_number}차 {reason_str} -> 100% 자력 탈출!")
                            sell_orders.append((index, quantity, current_price * 0.98, reason_str, buy_price, buy_date, round_number, "NORMAL"))
                        else:
                            print(f"{stock_code} {round_number}차 구제 얕은손실 대기(보유 {held_days}일): 현재가 {current_price:.2f} < 본전선 {be_price:.2f}")
                    continue

                # [구간 1·2] 0 ~ 40일 동적 목표가 평가
                dynamic_target_price = buy_price * (1.0 + target_pct)
                if current_price >= dynamic_target_price:
                    # [구간 2] 26 ~ 40일 본전 탈출 구간: 2단계/분할 없이 100% 즉시 탈출
                    if not allow_stage2:
                        reason_str = f"본전 탈출 구간 목표 도달(보유 {held_days}일, +{target_pct*100:.1f}%)"
                        print(f"{stock_code} {round_number}차 {reason_str} -> 100% 전량 즉시 탈출!")
                        sell_orders.append((index, quantity, current_price * 0.98, reason_str, buy_price, buy_date, round_number, "NORMAL"))
                        continue

                    # [구간 1] 0 ~ 25일 정상 수익 추구 구간 (+2.5%)
                    part1_ratio, part1_pct_str, rem_pct_str, part1_pct_val = get_stage1_partial_sell_config()

                    # [방식 1] 1단계 통과 시점에 이미 1h 5MA가 꺾여 있거나(is_ma5_broken == True),
                    # STAGE1_PARTIAL_SELL_PCT가 100%로 설정된 경우 즉시 100% 전량 익절
                    if is_ma5_broken or part1_ratio >= 1.0:
                        reason_str = (
                            f"1차 목표가 도달 + {TREND_TIMEFRAME} {TREND_MA_PERIOD}MA 하회 전량 익절(완성봉 {prev_close_4h:.2f} < {TREND_MA_PERIOD}MA {ma5_4h:.2f})"
                            if is_ma5_broken
                            else f"1차 목표가 도달 ({part1_pct_str} 전량 익절 설정)"
                        )
                        print(f"{stock_code} {round_number}차 {reason_str} -> 100% 전량 즉시 익절!")
                        sell_orders.append((
                            index,
                            quantity,
                            current_price * 0.98,
                            reason_str,
                            buy_price,
                            buy_date,
                            round_number,
                            "NORMAL",
                        ))
                        continue

                    # 1h 5MA 상승 추세가 살아있는 경우: 설정 비율(STAGE1_PARTIAL_SELL_PCT)만큼 즉시 분할익절 + 나머지는 최저보존선(1차 목표가) 기준으로 2단계 승격
                    part1_qty, rem_qty, can_split = (
                        split_stage1_sell_quantity(exchange, symbol, quantity, current_price, ratio=part1_ratio)
                        if part1_ratio > 0.0
                        else (0.0, float(quantity), False)
                    )
                    if can_split:
                        p_order, p_order_price = place_close_long(exchange, symbol, part1_qty)
                        print(f"[1단계 {part1_pct_str} 분할매도 주문 ({round_number}차)]", p_order)

                        fallback_part1 = float(p_order_price) if p_order_price is not None else float(current_price)
                        p_order_id = p_order.get("id") if isinstance(p_order, dict) else None
                        if p_order_id:
                            p_filled_ok, _ = wait_order_filled(exchange, p_order_id, symbol)
                            if not p_filled_ok and USE_LIMIT_ORDER:
                                cancel_order_safe(exchange, p_order_id, symbol)
                                telegram_alert.SendMessage(
                                    f"⚠️ 떨사오팔(Bitget) {stock_code} {round_number}차 1단계 {part1_pct_str} 분할매도 지정가 미체결 → 취소"
                                )
                                continue

                        real_part1_price, real_part1_qty, _ = 체결가_조회(
                            exchange, symbol, p_order, fallback_part1, part1_qty
                        )
                        rem_qty = float(exchange.amount_to_precision(symbol, max(0.0, float(quantity) - float(real_part1_qty))))
                        now_kst_str = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")

                        part1_net_sell = float(real_part1_price) * (1.0 - SELL_FEE_RATE)
                        part1_profit_pct = ((part1_net_sell / buy_price) - 1.0) * 100.0 if buy_price > 0 else 0.0
                        part1_pnl = (part1_net_sell - buy_price) * float(real_part1_qty)
                        part1_fee = float(real_part1_price) * float(real_part1_qty) * SELL_FEE_RATE

                        part1_saved_ratio = 0.0
                        part1_saved_usdt = 0.0
                        part1_pool_note = ""
                        if ENABLE_OFFSET_RESCUE and part1_pnl > 0 and _has_rescue_waiting(rounds_dict, exclude_rounds={round_number}):
                            part1_saved_ratio = max(0.0, min(1.0, float(OFFSET_PROFIT_SAVE_RATIO)))
                            part1_saved_usdt = part1_pnl * part1_saved_ratio
                            if part1_saved_usdt > 0:
                                new_pool = add_offset_pool(stock_code, part1_saved_usdt)
                                part1_pool_note = f"\n🏦 구제 적립금 반영: +${part1_saved_usdt:.2f} USDT (누적 ${new_pool:.2f} USDT)"

                        updated_info = update_round_stage(
                            stock_code,
                            round_number,
                            stage=2,
                            preservation_line=target_price,
                            target_price=target_price,
                            buy_price=buy_price,
                            quantity=rem_qty,
                            initial_quantity=initial_quantity,
                            part1_sold_qty=real_part1_qty,
                            part1_sold_price=real_part1_price,
                            part1_sold_pct=part1_pct_val,
                            part1_sold_at=now_kst_str,
                            part1_pool_saved_ratio=part1_saved_ratio,
                        )
                        rounds_dict[r_key] = updated_info
                        partial_sell_events.append((round_number, real_part1_qty, real_part1_price, rem_qty))

                        print(
                            f"{stock_code} {round_number}차 1차 목표가 돌파 -> {part1_pct_str}({real_part1_qty}주 @ {real_part1_price:.4f}) 분할익절 완료 & 잔여 {rem_pct_str}({rem_qty}주) 2단계(추세추종) 승격!"
                        )
                        msg = (
                            f"🎯 [1차 목표가 돌파 & {part1_pct_str} 분할익절] {stock_code} {round_number}차 2단계 승격\n"
                            f"🔹 1차 익절({part1_pct_str}): {real_part1_qty}주 체결 (체결가 {real_part1_price:.4f} / 수익률 {part1_profit_pct:+.2f}%)\n"
                            f"🔹 잔여 물량({rem_pct_str}): {rem_qty}주 → [2단계 추세 추종 모드 진입]\n"
                            f"📉 매수가: {buy_price:.2f} | 📈 1차 목표가: {target_price:.2f} (+{SELL_PROFIT_PCT*100:.1f}%)\n"
                            f"💲 현재가: {current_price:.2f} | 📊 {TREND_TIMEFRAME} {TREND_MA_PERIOD}MA: {ma5_4h:.2f}\n"
                            f"🛡️ 잔여 {rem_pct_str} 최저보존선: {target_price:.2f} (1차 목표 매도가)\n"
                            f"💡 잔여 {rem_pct_str}는 {TREND_TIMEFRAME}봉 {TREND_MA_PERIOD}MA 꺾임 시 초과 익절 또는 보존선 하향 터치({PRESERVATION_LINE_MULTIPLIER}) 시 즉시 방어 매도합니다."
                            f"{part1_pool_note}"
                        )
                        telegram_alert.SendMessage(msg)

                        sheet_logger.log_trade(
                            action="SELL",
                            symbol=symbol,
                            level=f"{round_number}차(1/2)",
                            price=real_part1_price,
                            amount=real_part1_qty,
                            fee=part1_fee,
                            pnl=part1_pnl,
                            pnl_pct=part1_profit_pct,
                            close_all=False,
                            note=f"1단계 목표가 돌파 {part1_pct_str} 분할익절 (잔여 {rem_qty}주 2단계 승격, 보존선 {target_price:.2f})",
                        )
                    else:
                        updated_info = update_round_stage(
                            stock_code,
                            round_number,
                            stage=2,
                            preservation_line=target_price,
                            target_price=target_price,
                            buy_price=buy_price,
                            quantity=quantity,
                            initial_quantity=initial_quantity,
                        )
                        rounds_dict[r_key] = updated_info
                        skip_note = "(분할익절 0% 설정)" if part1_ratio <= 0.0 else "(최소주문금액 제약으로 분할 생략)"
                        print(f"{stock_code} {round_number}차 1차 목표가 돌파 {skip_note} -> 전량 2단계(추세추종) 승격!")
                        msg = (
                            f"🎯 [1차 목표가 돌파] {stock_code} {round_number}차 2단계 승격\n"
                            f"🔹 진입 단계: [2단계 추세 추종 모드 {skip_note}]\n"
                            f"📈 1차 목표가: {target_price:.2f} (수익률 +{SELL_PROFIT_PCT*100:.1f}%)\n"
                            f"💲 현재가: {current_price:.2f}\n"
                            f"🛡️ 최소 이익 보존선: {target_price:.2f}\n"
                            f"💡 {TREND_TIMEFRAME}봉 {TREND_MA_PERIOD}MA 꺾임 시 초과 익절 또는 보존선 하향 터치({PRESERVATION_LINE_MULTIPLIER}) 시 즉시 방어 매도합니다."
                        )
                        telegram_alert.SendMessage(msg)
                else:
                    print(f"{stock_code} {round_number}차 1단계 일일 점검({phase_label}, 보유 {held_days}일): 현재가 {current_price:.2f} < 목표가 {dynamic_target_price:.2f}")
            else:
                print(f"{stock_code} {round_number}차 1단계 대기 중 (일일 점검 시점에만 목표/승격 확인)")

        # 3) 2단계: 추세추종 (기존 수익 추세추종 + 구제 반등 추세추종 + 본전 래칫 상향)
        elif current_stage == 2:
            if rescue_stage2:
                be_line = buy_price * (1.0 + STEPDOWN_BREAKEVEN_PCT)
                if current_price >= be_line and preservation_line < be_line - 1e-6:
                    preservation_line = be_line
                    rounds_dict[r_key] = update_round_stage(
                        stock_code,
                        round_number,
                        stage=2,
                        preservation_line=preservation_line,
                        target_price=preservation_line,
                        rescue_stage2=True,
                    )
                    print(f"{stock_code} {round_number}차 구제 추세추종 본전선({be_line:.2f}) 돌파 -> 보존선 상향 잠금!")
                    telegram_alert.SendMessage(
                        f"🛡️ [구제 추세추종 본전 잠금] {stock_code} {round_number}차\n"
                        f"🔹 현재가 {current_price:.2f}가 본전선(+{STEPDOWN_BREAKEVEN_PCT*100:.1f}%, {be_line:.2f})을 돌파했습니다.\n"
                        f"🔒 최저보존선을 {be_line:.2f}로 자동 상향하여 손실 재진입을 차단합니다."
                    )

            order_tag = "RESCUE_STAGE2" if rescue_stage2 else "NORMAL"
            pres_label = "구제 보존선 방어" if rescue_stage2 else "최소 이익 보존선 방어"
            ma_label = "구제 추세 꺾임 청산" if rescue_stage2 else "추세 꺾임 익절"

            # 시나리오 ②: 보존선 방어 매도 (현재가 <= 보존선 * 0.998)
            if current_price <= preservation_line * PRESERVATION_LINE_MULTIPLIER:
                print(f"{stock_code} {round_number}차 {pres_label} 조건 충족: 현재가 {current_price:.2f} <= {preservation_line * PRESERVATION_LINE_MULTIPLIER:.2f}")
                sell_orders.append((index, quantity, current_price * 0.98, f"{pres_label}(보존선 {preservation_line:.2f} × {PRESERVATION_LINE_MULTIPLIER} 대비 하향 터치)", buy_price, buy_date, round_number, order_tag))
            # 시나리오 ①: 추세 꺾임 익절 (TREND_TIMEFRAME TREND_MA_PERIOD MA 버퍼 기준 이탈)
            elif is_ma5_broken:
                _buf_p = float(TREND_MA_BUFFER_PCT)
                _buf_r = (_buf_p / 100.0) if _buf_p >= 0.05 else max(0.0, _buf_p)
                _ma_cut = ma5_4h * (1.0 - _buf_r)
                print(
                    f"{stock_code} {round_number}차 {ma_label} 조건 충족: "
                    f"{TREND_TIMEFRAME}봉 종가 {prev_close_4h:.2f} < {TREND_MA_PERIOD}MA 버퍼기준 {_ma_cut:.2f} (원본MA {ma5_4h:.2f})"
                )
                sell_orders.append((
                    index,
                    quantity,
                    current_price * 0.98,
                    f"{ma_label}({TREND_TIMEFRAME}봉 완성봉 {prev_close_4h:.2f} < {TREND_MA_PERIOD}MA 버퍼기준 {_ma_cut:.2f})",
                    buy_price,
                    buy_date,
                    round_number,
                    order_tag,
                ))
            else:
                print(f"{stock_code} {round_number}차 2단계 홀딩 중 (잔여 {quantity}주 | 현재가 {current_price:.2f} > 보존선 {preservation_line:.2f}, {TREND_TIMEFRAME} {TREND_MA_PERIOD}MA 상승 추세 유지)")

    closed_round_nums = {r_n for (_, _, _, _, _, _, r_n, _) in sell_orders}
    has_other_rescue_waiting = _has_rescue_waiting(rounds_dict, exclude_rounds=closed_round_nums)

    for index, quantity, sell_price, sell_reason, buy_price, buy_date, round_number, order_tag in sell_orders:
        r_key = str(round_number)
        r_info_before = rounds_dict.get(r_key, {})
        part1_qty_done = float(r_info_before.get("part1_sold_qty", 0.0) or 0.0)
        part1_price_done = r_info_before.get("part1_sold_price", None)
        part1_pct_done = r_info_before.get("part1_sold_pct", None)
        part1_saved_ratio = float(r_info_before.get("part1_pool_saved_ratio", 0.0) or 0.0)
        offset_qty_done = float(r_info_before.get("offset_sold_qty", 0.0) or 0.0)
        offset_proceeds_done = float(r_info_before.get("offset_sold_proceeds", 0.0) or 0.0)

        # ✅ 주문은 Bitget 선물 (청산), 엑셀엔 실체결값 저장
        order, order_price = place_close_long(exchange, symbol, quantity)
        print("[매도주문]", order)

        fallback_sell = float(order_price) if order_price is not None else float(current_price)
        order_id = order.get("id") if isinstance(order, dict) else None
        if order_id:
            filled_ok, _ = wait_order_filled(exchange, order_id, symbol)
            if not filled_ok and USE_LIMIT_ORDER:
                cancel_order_safe(exchange, order_id, symbol)
                telegram_alert.SendMessage(f"⚠️ 떨사오팔(Bitget) {stock_code} 매도 지정가 미체결 → 취소 ({df.at[index,'차수']}차)")
                continue

        # ★ 비트겟 실제 체결가/수량
        real_sell_price, real_sell_qty, _ = 체결가_조회(exchange, symbol, order, fallback_sell, quantity)
        quantity = real_sell_qty

        try:
            buy_price_f = float(buy_price)
        except Exception:
            buy_price_f = float(pd.to_numeric(pd.Series([buy_price]), errors="coerce").iloc[0])

        # 1차 분할익절 / 이전 부분 구제상계 / 금회 최종 매도를 합산하여 가중평균 매도가 산출
        has_part1 = (part1_qty_done > 0 and part1_price_done is not None and float(part1_price_done) > 0)
        has_offset = (offset_qty_done > 0 and offset_proceeds_done > 0)
        p1_proceeds = (float(part1_price_done) * float(part1_qty_done)) if has_part1 else 0.0
        off_proceeds = float(offset_proceeds_done) if has_offset else 0.0
        this_proceeds = float(real_sell_price) * float(real_sell_qty)

        total_sold_qty = (float(part1_qty_done) if has_part1 else 0.0) + (float(offset_qty_done) if has_offset else 0.0) + float(real_sell_qty)
        avg_sell_price = ((p1_proceeds + off_proceeds + this_proceeds) / total_sold_qty) if total_sold_qty > 0 else float(real_sell_price)

        df.at[index, "매도가"] = avg_sell_price
        df.at[index, "매도일"] = datetime.now(KST).strftime("%Y-%m-%d")
        df.at[index, "매도수량"] = total_sold_qty

        sell_price_f = float(avg_sell_price)
        net_sell = sell_price_f * (1.0 - SELL_FEE_RATE)

        if buy_price_f is None or pd.isna(buy_price_f) or buy_price_f <= 0:
            profit_pct = 0.0
        else:
            profit_pct = ((net_sell / buy_price_f) - 1.0) * 100.0

        # 구제 적립금 적립 / 손실 보전 및 복리 조정률 왜곡 방지
        this_net_sell = float(real_sell_price) * (1.0 - SELL_FEE_RATE)
        this_fill_pnl = (this_net_sell - buy_price_f) * float(real_sell_qty) if buy_price_f > 0 else 0.0
        part1_pnl_done = ((float(part1_price_done) * (1.0 - SELL_FEE_RATE) - buy_price_f) * float(part1_qty_done)) if (has_part1 and buy_price_f > 0) else 0.0

        this_saved_usdt = 0.0
        this_covered_usdt = 0.0
        pool_msg_suffix = ""
        if ENABLE_OFFSET_RESCUE:
            if this_fill_pnl > 0 and has_other_rescue_waiting:
                save_ratio = max(0.0, min(1.0, float(OFFSET_PROFIT_SAVE_RATIO)))
                this_saved_usdt = this_fill_pnl * save_ratio
                if this_saved_usdt > 0:
                    new_pool = add_offset_pool(stock_code, this_saved_usdt)
                    pool_msg_suffix = f"\n🏦 구제 적립금 적립: +${this_saved_usdt:.2f} USDT (누적 ${new_pool:.2f} USDT)"
            elif this_fill_pnl < 0 and order_tag in ("RESCUE_REBOUND", "RESCUE_STAGE2"):
                this_covered_usdt, remain_pool = consume_offset_pool(stock_code, abs(this_fill_pnl))
                if this_covered_usdt > 0:
                    pool_msg_suffix = f"\n🏦 구제 적립금 보전: ${this_covered_usdt:.2f} USDT 차감 (잔여 ${remain_pool:.2f} USDT)"

        part1_unoffset_pnl = part1_pnl_done * (1.0 - part1_saved_ratio)
        this_unoffset_pnl = this_fill_pnl - this_saved_usdt + this_covered_usdt
        total_cost_basis = buy_price_f * total_sold_qty if (buy_price_f > 0 and total_sold_qty > 0) else 0.0
        effective_profit_pct = ((part1_unoffset_pnl + this_unoffset_pnl) / total_cost_basis) * 100.0 if total_cost_basis > 0 else 0.0

        weight = HIGH_DROP_ADJ_WEIGHT if bool(df.at[index, "고점대비하락조건"]) else 1.0
        adj_pct = (effective_profit_pct / ADJ_DIVISOR) * weight

        next_invest_money = int(round(current_invest_money * (1.0 + (adj_pct / 100.0))))

        df.at[index, "수수료율(매도)"] = 0.06
        df.at[index, "순매도가(수수료반영)"] = net_sell
        df.at[index, "수익률(%)"] = profit_pct
        df.at[index, "조정률(수익률/7)(%)"] = adj_pct
        df.at[index, "다음투자금"] = next_invest_money

        투자금시트_업데이트(stock_code, next_invest_money, effective_profit_pct, adj_pct)
        current_invest_money = next_invest_money

        # 상태 파일에서 해당 차수 제거
        remove_round_from_state(stock_code, round_number)

        drop_rate = 0.0
        if highest_price_30d and highest_price_30d > 0:
            drop_rate = (highest_price_30d - current_price) / highest_price_30d * 100.0

        if has_part1:
            _, default_p1_str, default_rem_str, default_p1_val = get_stage1_partial_sell_config()
            p1_val = float(part1_pct_done) if part1_pct_done is not None else default_p1_val
            r2_val = round(100.0 - p1_val, 1)
            p1_label = f"{int(p1_val)}%" if p1_val.is_integer() else f"{p1_val:.1f}%"
            r2_label = f"{int(r2_val)}%" if r2_val.is_integer() else f"{r2_val:.1f}%"

            part2_net_sell = float(real_sell_price) * (1.0 - SELL_FEE_RATE)
            part2_profit_pct = ((part2_net_sell / buy_price_f) - 1.0) * 100.0 if buy_price_f > 0 else 0.0
            msg = (
                f"📌 떨사오팔 일반전략 {stock_code} {df.at[index, '차수']}차 2단계 잔여({r2_label}) 최종 매도 완료\n"
                f"🔹 매도 사유: {sell_reason}\n"
                f"📉 매수가(기록): {buy_price_f:.2f}\n"
                f"📈 1차({p1_label}) 기익절: {part1_qty_done}주 @ {float(part1_price_done):.4f}\n"
                f"📈 2차({r2_label}) 금회체결: {quantity}주 @ {float(real_sell_price):.4f} ({part2_profit_pct:+.2f}%)\n"
                f"📊 합산 평균매도가: {sell_price_f:.4f} (총 {total_sold_qty}주 완판)\n"
                f"📅 매수일: {buy_date} | 📅 매도완결일: {datetime.now(KST).strftime('%Y-%m-%d')}\n"
                f"💸 매도수수료: 0.06% (매도만)\n"
                f"✅ 최종 합산 수익률(수수료반영): {profit_pct:.2f}%\n"
                f"✅ 조정률(유효수익률/{int(ADJ_DIVISOR)}×가중치 {weight}): {adj_pct:.2f}%\n"
                f"💰 다음 투자금(종목별, 정수): {next_invest_money}\n"
                f"📊 고점 대비 낙폭률(현재가 기준): {drop_rate:.2f}%"
                f"{pool_msg_suffix}"
            )
        else:
            offset_hist_line = (
                f"\n🛟 기 상계청산분: {offset_qty_done}주 (합산 {total_sold_qty}주 완판)"
                if has_offset
                else ""
            )
            msg = (
                f"📌 떨사오팔 일반전략 {stock_code} {df.at[index, '차수']}차 매도 발생\n"
                f"🔹 매도 사유: {sell_reason}\n"
                f"📉 매수가(기록): {buy_price_f:.2f}\n"
                f"📈 매도가(Bitget 체결): {sell_price_f:.4f}\n"
                f"📊 매도 수량: {quantity}주{offset_hist_line}\n"
                f"📅 매수일: {buy_date}\n"
                f"📅 매도일: {datetime.now(KST).strftime('%Y-%m-%d')}\n"
                f"💸 매도수수료: 0.06% (매도만)\n"
                f"✅ 수익률(수수료반영): {profit_pct:.2f}%\n"
                f"✅ 조정률(유효수익률/{int(ADJ_DIVISOR)}×가중치 {weight}): {adj_pct:.2f}%\n"
                f"💰 다음 투자금(종목별, 정수): {next_invest_money}\n"
                f"📊 고점 대비 낙폭률(현재가 기준): {drop_rate:.2f}%"
                f"{pool_msg_suffix}"
            )
        telegram_alert.SendMessage(msg)

        # [구글 시트 체결 기록 전송 추가]
        part2_net = float(real_sell_price) * (1.0 - SELL_FEE_RATE)
        real_pnl = (part2_net - buy_price_f) * float(quantity)  # 금회 체결분 실현 손익 (USDT)
        fee_usdt = float(real_sell_price) * float(quantity) * SELL_FEE_RATE  # 금회 체결분 매도 수수료 (USDT)
        part2_pct = ((part2_net / buy_price_f) - 1.0) * 100.0 if buy_price_f > 0 else 0.0
        sheet_logger.log_trade(
            action="SELL",
            symbol=symbol,
            level=f"{df.at[index, '차수']}차(2/2)" if has_part1 else f"{df.at[index, '차수']}차",
            price=real_sell_price,
            amount=quantity,
            fee=fee_usdt,
            pnl=real_pnl,
            pnl_pct=part2_pct if has_part1 else profit_pct,
            close_all=True,
            note=f"{sell_reason} (합산수익률 {profit_pct:.2f}%, 조정률 {adj_pct:+.2f}%)" if has_part1 else f"{sell_reason} (조정률 {adj_pct:+.2f}%)"
        )

    if sell_orders:
        # 매도(강제매도/익절) 발생으로 빈 차수가 생기면 남은 미매도 차수들을 매수가 내림차순으로 1차부터 순서대로 앞당김
        state = sync_round_state_from_excel(df, stock_code)
        rounds_dict = state.get(stock_code, {}).get("rounds", {})

    # D 전략(트랙 B): 누적 구제 적립금이 있으면 최우선 구제 대기 차수 자동 부분/전량 상계 소각
    offset_events = execute_offset_rescue(exchange, symbol, stock_code, df, current_price, current_invest_money)
    if offset_events:
        state = sync_round_state_from_excel(df, stock_code)
        rounds_dict = state.get(stock_code, {}).get("rounds", {})
        stock_data = df[df["종목"] == stock_code]

    if not sell_orders and not partial_sell_events and not offset_events:
        remaining = 0 if stock_data.empty else len(stock_data[pd.isna(stock_data["매도가"])])
        now_et = datetime.now(US_EASTERN)

        # 미매도 보유 차수별 매도 기준가 목록 구성
        rounds_detail_lines = []
        has_any_rescue = False
        if not stock_data.empty:
            uncompleted_df = stock_data[pd.isna(stock_data["매도가"])].copy()
            if not uncompleted_df.empty:
                try:
                    uncompleted_df["_sort_round"] = pd.to_numeric(uncompleted_df["차수"], errors="coerce")
                    uncompleted_df = uncompleted_df.sort_values(by="_sort_round")
                except Exception:
                    pass

                for _, r_row in uncompleted_df.iterrows():
                    try:
                        r_num = int(r_row["차수"])
                    except Exception:
                        r_num = r_row["차수"]

                    try:
                        b_price = float(r_row["매수가"])
                    except Exception:
                        b_price = 0.0

                    t_price = b_price * (1.0 + SELL_PROFIT_PCT)
                    r_key = str(r_num)
                    r_state_info = rounds_dict.get(r_key, {})
                    r_stage = int(r_state_info.get("stage", 1))
                    r_pres_line = float(r_state_info.get("preservation_line", t_price))
                    r_part1_qty = float(r_state_info.get("part1_sold_qty", 0.0) or 0.0)
                    r_part1_pct = r_state_info.get("part1_sold_pct", None)
                    r_rescue_mode = bool(r_state_info.get("rescue_mode", False))
                    r_deep_loss = bool(r_state_info.get("deep_loss_seen", False))
                    r_rescue_stage2 = bool(r_state_info.get("rescue_stage2", False))
                    r_offset_sold_qty = float(r_state_info.get("offset_sold_qty", 0.0) or 0.0)
                    r_rem_qty = float(r_state_info.get("quantity", r_row.get("수량", 0.0)) or 0.0)

                    r_max_hold_days = MAX_HOLD_DAYS_HIGH_DROP if bool(r_row.get("고점대비하락조건", False)) else MAX_HOLD_DAYS_NORMAL
                    r_held_days = 0
                    force_sell_str = ""
                    r_is_rescue_phase = r_rescue_mode
                    r_allow_stage2 = True
                    r_target_pct = SELL_PROFIT_PCT
                    try:
                        if isinstance(r_row["날짜"], pd.Timestamp):
                            r_buy_date = r_row["날짜"].date()
                        else:
                            r_buy_date = datetime.strptime(str(r_row["날짜"])[:10], "%Y-%m-%d").date()
                        r_held_days = (today - r_buy_date).days
                        r_target_pct, r_allow_stage2, r_is_rescue_phase, _ = get_round_target_profile(r_held_days, r_max_hold_days)
                        if ENABLE_HARD_TIMECUT:
                            force_sell_date = r_buy_date + timedelta(days=r_max_hold_days)
                            force_sell_str = f", 강제매도일 {force_sell_date.strftime('%Y-%m-%d')}"
                        elif r_is_rescue_phase or r_rescue_mode:
                            has_any_rescue = True
                            off_str = f", 상계 {r_offset_sold_qty:g}주" if r_offset_sold_qty > 0 else ""
                            force_sell_str = f", 구제진행중(보유 {r_held_days}일, 잔여 {r_rem_qty:g}주{off_str})"
                        else:
                            rescue_date = r_buy_date + timedelta(days=r_max_hold_days + 1)
                            force_sell_str = f", 구제전환일 {rescue_date.strftime('%Y-%m-%d')}"
                    except Exception:
                        pass

                    if r_stage == 2:
                        if r_rescue_stage2:
                            stage2_tag = "🛟구제2단계 추세"
                        elif r_part1_qty > 0:
                            _, _, def_rem_str, def_p1_val = get_stage1_partial_sell_config()
                            p1_v = float(r_part1_pct) if r_part1_pct is not None else def_p1_val
                            rem_v = round(100.0 - p1_v, 1)
                            rem_lbl = f"{int(rem_v)}%" if rem_v.is_integer() else f"{rem_v:.1f}%"
                            stage2_tag = f"2단계 잔여{rem_lbl}"
                        else:
                            stage2_tag = "2단계 진입"
                        _bp = float(TREND_MA_BUFFER_PCT)
                        _br = (_bp / 100.0) if _bp >= 0.05 else max(0.0, _bp)
                        _ma_disp = ma5_4h * (1.0 - _br) if ma5_4h > 0 else 0.0
                        ma_str = f"{_ma_disp:.2f}({stage2_tag})" if _ma_disp > 0 else f"{r_pres_line:.2f}({stage2_tag})"
                        rounds_detail_lines.append(
                            f"  • {r_num}차: {ma_str} (매수가 {b_price:.2f}, 보존선 {r_pres_line:.2f}{force_sell_str})"
                        )
                    else:
                        if (r_is_rescue_phase or r_rescue_mode) and not ENABLE_HARD_TIMECUT:
                            if r_deep_loss:
                                reb_price = b_price * (1.0 + RESCUE_REBOUND_TRIGGER_PCT)
                                rounds_detail_lines.append(
                                    f"  • {r_num}차: {reb_price:.2f}(🛟구제 반등선 {RESCUE_REBOUND_TRIGGER_PCT*100:+.0f}%) (매수가 {b_price:.2f}{force_sell_str})"
                                )
                            else:
                                be_price = b_price * (1.0 + STEPDOWN_BREAKEVEN_PCT)
                                rounds_detail_lines.append(
                                    f"  • {r_num}차: {be_price:.2f}(🛟구제 본전선 +{STEPDOWN_BREAKEVEN_PCT*100:.1f}%) (매수가 {b_price:.2f}{force_sell_str})"
                                )
                        elif not r_allow_stage2:
                            dyn_t = b_price * (1.0 + r_target_pct)
                            rounds_detail_lines.append(
                                f"  • {r_num}차: {dyn_t:.2f}(본전구간 +{r_target_pct*100:.1f}%) (매수가 {b_price:.2f}{force_sell_str})"
                            )
                        else:
                            rounds_detail_lines.append(
                                f"  • {r_num}차: {t_price:.2f} (매수가 {b_price:.2f}{force_sell_str})"
                            )

        pool_bal_disp = get_offset_pool(stock_code)
        pool_footer = f"\n🏦 구제 적립금 잔액: ${pool_bal_disp:.2f} USDT" if (pool_bal_disp > 0 or has_any_rescue) else ""
        rounds_detail_str = ("\n🎯 차수별 매도 기준가:\n" + "\n".join(rounds_detail_lines) + pool_footer) if rounds_detail_lines else pool_footer

        # 1) 매일 1회 일일 매수 점검 턴 (평일 15:51 ET [썸머타임 자동 반영] 또는 force_buy):
        #    사용자 요청에 따라 매도 점검 결과(매도 없음)와 매수 결과를 순차적으로 항상 전송
        if is_daily_buy_turn:
            # 매수 메시지와 일관성을 위해 가능하면 LS 현재가 조회 시도, 실패 시 선물 현재가 사용
            disp_price = current_price
            try:
                ls_curr = LsiUS.GetCurrentPrice(stock_code)
                if ls_curr is not None and not _값이_비정상(ls_curr):
                    disp_price = float(ls_curr)
            except Exception:
                pass

            msg = (
                f"⏸️ {stock_code} 매도 없음\n"
                f"🔹 이유: 목표가 미도달 & 구제모드 대기\n"
                f"📈 현재가(LS): {disp_price:.2f}\n"
                f"🧾 미매도 보유 차수: {remaining}개"
                f"{rounds_detail_str}"
            )
            telegram_alert.SendMessage(msg)
            print(f"[매도 점검] {stock_code} 매도 없음 알림 발송 (일일 매수 점검 턴 / 잔여 {remaining}개)")
        # 2) 그 외 일반 시간대의 4시간 간격 51분 (정기 브리핑: 미국 장 마감 10분 전 기준 4시간 간격, 썸머타임 자동 반영):
        #    알림 피로도 최적화를 위해 2단계 추세추종 차수가 1개 이상 존재할 때만 발송
        elif now_et.minute == DAILY_CHECK_ET_MINUTE and (now_et.hour - DAILY_CHECK_ET_HOUR) % STAGE2_BRIEFING_INTERVAL_HOURS == 0:
            latest_state = load_round_state()
            latest_rounds = latest_state.get(stock_code, {}).get("rounds", {})
            stage2_cnt = sum(1 for r in latest_rounds.values() if r.get("stage") == 2)
            if stage2_cnt > 0:
                msg = (
                    f"⏸️ {stock_code} 매도 현황 브리핑 (정기)\n"
                    f"🔹 상태: 목표가 미도달 & 2단계 추세 유지 중\n"
                    f"📈 현재가(LS): {current_price:.2f}\n"
                    f"🧾 미매도 보유 차수: {remaining}개 (2단계 추세추종: {stage2_cnt}개)"
                    f"{rounds_detail_str}"
                )
                telegram_alert.SendMessage(msg)
            else:
                print(f"[매도 없음] {stock_code} (2단계 추세추종 0개로 {STAGE2_BRIEFING_INTERVAL_HOURS}시간 정기 브리핑 생략 / 잔여 {remaining}개)")
        else:
            print(f"[매도 없음] {stock_code} (목표가 미도달 or 2단계 홀딩 중 / 잔여 {remaining}개)")

    매매기록저장(df)
    try:
        latest_state = load_round_state()
        latest_stock_state = latest_state.get(stock_code, {})
        latest_rounds = latest_stock_state.get("rounds", {})
        has_stage2_now = any(r.get("stage") == 2 for r in latest_rounds.values())
        rescue_cnt = sum(1 for r in latest_rounds.values() if r.get("rescue_mode"))
        sheet_logger.log_bot_state({
            "symbol": symbol,
            "stock_code": stock_code,
            "updated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"),
            "offset_pool_usdt": round(_safe_nonneg_float(latest_stock_state.get("offset_pool_usdt", 0.0), 0.0), 4),
            "rescue_rounds_cnt": int(rescue_cnt),
            "trend_riding": {
                "active": bool(prev_close_4h > 0 and not is_ma5_broken),
                "stage2_active": bool(has_stage2_now),
                "prev_close_4h": round(float(prev_close_4h), 4),
                "ma5_4h": round(float(ma5_4h), 4),
                "timeframe": TREND_TIMEFRAME,
                "ma_period": TREND_MA_PERIOD,
                "ma_buffer_pct": TREND_MA_BUFFER_PCT,
                "stage1_partial_sell_pct": STAGE1_PARTIAL_SELL_PCT,
            },
            "rounds": latest_rounds,
        }, strategy_name="떨사오팔")
    except Exception as e:
        print(f"⚠️ [STATE_SYNC] 떨사오팔 상태 동기화 오류(무시): {e}")
    return bool(sell_orders or partial_sell_events or offset_events)


# =========================
# ✅ 매수 기록
# =========================
def 매수기록(stock_code, round_number, buy_price, quantity, high_drop_condition):
    df = 매매기록불러오기()
    buy_date = datetime.now(KST).strftime("%Y-%m-%d")

    new_data = {
        "차수": round_number,
        "날짜": buy_date,
        "종목": stock_code,
        "매수가": buy_price,
        "수량": quantity,
        "고점대비하락조건": high_drop_condition,
        "매도가": None,
        "매도일": None,
        "매도수량": None
    }

    df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
    매매기록저장(df)

    # 차수별 상태 파일 등록 (새로운 매수는 1단계로 진입) 후 매수가 내림차순 1~N차 정렬 동기화
    target_price = float(buy_price) * (1.0 + SELL_PROFIT_PCT)
    update_round_stage(
        stock_code,
        round_number,
        stage=1,
        preservation_line=target_price,
        target_price=target_price,
        buy_price=float(buy_price),
        quantity=float(quantity),
        initial_quantity=float(quantity),
        part1_sold_qty=0.0,
        buy_date=buy_date,
        part1_pool_saved_ratio=0.0,
        rescue_mode=False,
        deep_loss_seen=False,
        rescue_stage2=False,
        offset_sold_qty=0.0,
        offset_sold_proceeds=0.0,
        offset_used_usdt=0.0,
        offset_count=0,
    )
    sync_round_state_from_excel(df, stock_code)
    print(f"{stock_code} {round_number}차 매수 기록 및 상태(1단계) 등록 완료 - 매수가: {buy_price}, 수량: {quantity}, 고점대비하락조건: {high_drop_condition}")


# =========================
# ✅ 매수 체크
#    데이터=LS / 주문=Bitget 선물
# =========================
def 매수체크(exchange, symbol, stock_code, default_invest_money, ohlcv=None):
    df = 매매기록불러오기()
    sync_round_state_from_excel(df, stock_code)
    stock_data = df[df["종목"] == stock_code]

    # ✅ 당일(KST 기준) 이미 매수한 이력이 있는지 사전 점검 (5분 크론탭 급락 시 연속 매수 원천 차단)
    today_str = datetime.now(KST).strftime("%Y-%m-%d")
    if not stock_data.empty and "날짜" in stock_data.columns:
        bought_today = stock_data[stock_data["날짜"].astype(str).str.startswith(today_str)]
        if not bought_today.empty:
            rounds_bought = bought_today["차수"].tolist()
            print(f"⏸️ [당일 매수 완료] {stock_code}는 오늘({today_str}) 이미 {rounds_bought}차 매수가 완료되었습니다. (추가 매수 차단)")
            return False

    active_cnt = len(stock_data[stock_data["매도가"].isna()]) if not stock_data.empty else 0
    next_round = 가능다음차수(stock_data) if not stock_data.empty else 1
    if next_round is None:
        msg = (
            f"⏸️ [최대 보유 차수 제한] {stock_code} 매수 보류\n"
            f"🔹 사유: 현재 미매도 보유 차수({active_cnt}개)가 상한({MAX_ACTIVE_ROUNDS}개)에 도달했습니다.\n"
            f"💡 기존 차수 익절 또는 구제 상계 완결로 보유 차수가 줄어들면 자동 매수를 재개합니다."
        )
        telegram_alert.SendMessage(msg)
        print(f"⏸️ [매수 보류 - 최대 차수 상한] 보유 {active_cnt}개 >= {MAX_ACTIVE_ROUNDS}개")
        return False

    is_heavy_hold = active_cnt >= HEAVY_HOLD_ROUND_THRESHOLD
    effective_buy_drop_pct = HEAVY_HOLD_BUY_DROP_PCT if is_heavy_hold else BUY_DROP_PCT
    heavy_hold_note = f" [과부하 방어: 보유 {active_cnt}개 → -{effective_buy_drop_pct*100:.1f}% 강화]" if is_heavy_hold else ""

    invest_money = 투자금시트_초기화_및_불러오기(stock_code, default_invest_money)

    # ✅ 데이터는 LS
    current_price = LsiUS.GetCurrentPrice(stock_code)
    if ohlcv is None:
        ohlcv = LsCommon.GetOhlcv("US", stock_code)
    open_price = ohlcv["open"].iloc[-1]

    highest_price = 과거_고점찾기(stock_code, ohlcv)
    high_drop_condition = (highest_price - current_price) / highest_price >= HIGH_DROP_THRESHOLD if highest_price else False

    if current_price <= open_price * (1.0 - effective_buy_drop_pct):
        adjusted_investment = invest_money * HIGH_DROP_INVEST_MULTIPLIER if high_drop_condition else invest_money

        # ✅ 보유금(가용 USDT) 사전 점검 (방안 A: Skip & Alert)
        available_usdt = get_available_usdt(exchange)
        if 0 <= available_usdt < adjusted_investment:
            msg = (
                f"🚨 [보유금 부족] 떨사오팔 {stock_code} {next_round}차 매수 보류\n"
                f"🔹 사유: 비트겟 선물 가용 USDT가 필요 투자금보다 부족합니다.\n"
                f"💵 필요 투자금: 약 {adjusted_investment:.1f} USDT\n"
                f"💰 가용 잔고: {available_usdt:.2f} USDT\n"
                f"💡 조치: 비트겟 선물 계좌에 USDT를 보충하시면 다음 스케줄에 정상 매수됩니다."
            )
            telegram_alert.SendMessage(msg)
            print(f"[매수 보류 - 보유금 부족] 가용={available_usdt:.2f} USDT < 필요={adjusted_investment:.1f} USDT")
            return False

        # ✅ 주문은 Bitget 선물 (USDT ÷ 선물가), 엑셀엔 실체결값 저장(매수가=실체결가)
        futures_price = get_current_futures_price(exchange, symbol)
        amount = usdt_to_amount(exchange, symbol, adjusted_investment, futures_price)

        # ✅ 주문 실행 및 잔고부족 예외 안전 처리 (사후 방어)
        try:
            order, order_price = place_open_long(exchange, symbol, amount)
        except ccxt.InsufficientFunds as e:
            msg = (
                f"🚨 [보유금 부족] 떨사오팔 {stock_code} {next_round}차 매수 실패\n"
                f"🔹 사유: 비트겟 거래소에서 증거금 부족(InsufficientFunds)으로 주문이 거절되었습니다.\n"
                f"💵 주문 금액: 약 {adjusted_investment:.1f} USDT (수량: {amount})\n"
                f"💡 조치: 비트겟 선물 계좌의 가용 증거금을 확인 후 충전해 주세요."
            )
            telegram_alert.SendMessage(msg)
            print(f"[매수 취소 - 거래소 증거금 부족] {e}")
            return False
        except ccxt.ExchangeError as e:
            err_str = str(e).lower()
            if any(k in err_str for k in ["margin", "balance", "fund", "insufficient", "40754", "40762"]):
                msg = (
                    f"🚨 [보유금 부족] 떨사오팔 {stock_code} {next_round}차 매수 실패\n"
                    f"🔹 사유: 비트겟 거래소 증거금 부족 오류 ({e})\n"
                    f"💵 주문 금액: 약 {adjusted_investment:.1f} USDT\n"
                    f"💡 조치: 증거금을 충전하시면 다음 스케줄에 정상 실행됩니다."
                )
                telegram_alert.SendMessage(msg)
                print(f"[매수 취소 - 비트겟 증거금 에러] {e}")
                return False
            raise

        print("[매수주문]", order)

        fallback_buy = float(order_price) if order_price is not None else float(futures_price)
        order_id = order.get("id") if isinstance(order, dict) else None
        if order_id:
            filled_ok, _ = wait_order_filled(exchange, order_id, symbol)
            if not filled_ok and USE_LIMIT_ORDER:
                cancel_order_safe(exchange, order_id, symbol)
                telegram_alert.SendMessage(f"⚠️ 떨사오팔(Bitget) {stock_code} {next_round}차 매수 지정가 미체결 → 취소")
                return False

        # ★ 비트겟 실제 체결가/수량
        real_buy_price, real_buy_qty, _ = 체결가_조회(exchange, symbol, order, fallback_buy, amount)
        buy_amt = real_buy_qty

        매수기록(stock_code, next_round, real_buy_price, buy_amt, high_drop_condition)

        drop_rate = 0.0
        if highest_price and highest_price > 0:
            drop_rate = (highest_price - current_price) / highest_price * 100.0

        invest_suffix = f" → 증액({HIGH_DROP_INVEST_MULTIPLIER:.1f}배) 적용" if high_drop_condition else ""

        msg = (
            f"🟢 떨사오팔 일반전략 {stock_code} {next_round}차 매수 실행\n"
            f"🔹 매수 사유: 현재가 ≤ 시가×{(1.0-effective_buy_drop_pct):.3f}"
            f"{' + 고점 대비 하락조건 충족' if high_drop_condition else ''}{heavy_hold_note}\n"
            f"📈 시가(LS): {open_price:.2f} | 💲 현재가(LS): {current_price:.2f} | 📉 매수체결가(Bitget): {real_buy_price:.4f}\n"
            f"💵 투자금(종목별, 정수): {invest_money}{invest_suffix}\n"
            f"📊 주문 수량: {buy_amt}주\n"
            f"📅 매수일: {datetime.now(KST).strftime('%Y-%m-%d')}\n"
            f"📊 고점 대비 낙폭률(LS): {drop_rate:.2f}%"
        )
        telegram_alert.SendMessage(msg)

        # [구글 시트 체결 기록 전송 추가]
        sheet_logger.log_trade(
            action="BUY",
            symbol=symbol,
            level=f"{next_round}차",
            price=real_buy_price,
            amount=buy_amt,
            note=f"떨사 분할 매수 실행 (낙폭률 {drop_rate:.2f}%)"
        )
        try:
            latest_state = load_round_state()
            latest_stock_state = latest_state.get(stock_code, {})
            latest_rounds = latest_stock_state.get("rounds", {})
            has_stage2_now = any(r.get("stage") == 2 for r in latest_rounds.values())
            rescue_cnt = sum(1 for r in latest_rounds.values() if r.get("rescue_mode"))
            is_ma5_broken, prev_close_4h, ma5_4h = get_trend_ma_status(exchange, symbol)
            sheet_logger.log_bot_state({
                "symbol": symbol,
                "stock_code": stock_code,
                "updated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"),
                "offset_pool_usdt": round(_safe_nonneg_float(latest_stock_state.get("offset_pool_usdt", 0.0), 0.0), 4),
                "rescue_rounds_cnt": int(rescue_cnt),
                "trend_riding": {
                    "active": bool(prev_close_4h > 0 and not is_ma5_broken),
                    "stage2_active": bool(has_stage2_now),
                    "prev_close_4h": round(float(prev_close_4h), 4),
                    "ma5_4h": round(float(ma5_4h), 4),
                    "timeframe": TREND_TIMEFRAME,
                    "ma_period": TREND_MA_PERIOD,
                    "ma_buffer_pct": TREND_MA_BUFFER_PCT,
                    "stage1_partial_sell_pct": STAGE1_PARTIAL_SELL_PCT,
                },
                "rounds": latest_rounds,
            }, strategy_name="떨사오팔")
        except Exception as e:
            print(f"⚠️ [STATE_SYNC] 떨사오팔 매수 후 상태 동기화 오류(무시): {e}")

        return True

    else:
        target_buy_price = open_price * (1.0 - effective_buy_drop_pct)
        price_items = [
            ("📈 시가(LS)", open_price),
            ("💲 현재가(LS)", current_price),
            ("🎯 기준가", target_buy_price)
        ]
        # 가격 높은 순(내림차순) 정렬
        price_items.sort(key=lambda x: x[1], reverse=True)
        price_lines = "\n".join([f"{label}: {val:.2f}" for label, val in price_items])

        msg = (
            f"⏸️ {stock_code} 매수 보류\n"
            f"🔹 조건: 현재가 ≤ 시가×{(1.0-effective_buy_drop_pct):.3f}{heavy_hold_note}\n"
            f"{price_lines}\n"
            f"메모: 떨사오팔 일반 전략 조건 미충족으로 매수하지 않습니다."
        )
        telegram_alert.SendMessage(msg)
        return False


# =========================
# ✅ 매수/매도 프로세스
# =========================
def 매수매도프로세스(exchange, symbol, stock_code, default_invest_money):
    # 1. 일일 매수/매도2단계 진입 체크 시점 판별 : 미국 동부시간(ET) 평일 15:51 (미국 장 마감 10분 전, 썸머타임 자동 반영)
    #    - 썸머타임 적용(EDT, UTC-4): 15:51 EDT = 19:51 UTC = 04:51 KST
    #    - 썸머타임 해제(EST, UTC-5): 15:51 EST = 20:51 UTC = 05:51 KST
    now_utc = datetime.now(timezone.utc)
    now_et = now_utc.astimezone(US_EASTERN)
    et_tz_name = now_et.tzname() or "ET"
    target_utc_hour = (DAILY_CHECK_ET_HOUR - int(now_et.utcoffset().total_seconds() // 3600)) % 24
    target_kst_hour = (target_utc_hour + 9) % 24
    schedule_desc = (
        f"평일 {DAILY_CHECK_ET_HOUR:02d}:{DAILY_CHECK_ET_MINUTE:02d} {et_tz_name} "
        f"({target_utc_hour:02d}:{DAILY_CHECK_ET_MINUTE:02d} UTC / {target_kst_hour:02d}:{DAILY_CHECK_ET_MINUTE:02d} KST)"
    )

    is_weekday = now_et.weekday() < 5                 # 미국 정규장 기준 월~금 (0~4)
    is_buy_minute = (now_et.hour == DAILY_CHECK_ET_HOUR and now_et.minute == DAILY_CHECK_ET_MINUTE)
    force_buy = os.environ.get("FORCE_BUY_CHECK", "0") == "1"
    is_daily_buy_turn = (is_weekday and is_buy_minute) or force_buy

    # 2. 일일 점검 턴이 아닌 일반 5분 크론탭 턴에서는 2단계(stage == 2) 진입 차수가 있을 때만 매도 감시 수행
    if not is_daily_buy_turn:
        if not has_active_stage2_rounds(stock_code):
            print(
                f"[매도/매수 스킵] 현재 2단계 진입 차수가 없으며 일일 점검 시간({schedule_desc}, 현재 {now_et.strftime('%H:%M')} {et_tz_name} / {now_utc.strftime('%H:%M')} UTC)이 아니므로 다음 날 점검 시점까지 대기합니다."
            )
            return

    # ✅ OHLCV 데이터는 LS
    ohlcv = LsCommon.GetOhlcv("US", stock_code)

    # 3. [매도 감시] :
    #    - 일일 매수 점검 턴(하루 1번): 매수 확인 전 1단계 -> 2단계 진입(승격) 여부 및 타임컷/2단계 매도 조건 확인
    #    - 일반 5분 턴(2단계 진입 차수 보유 시): 2단계 차수의 매도 조건(추세 꺾임 / 보존선 방어)만 5분 간격 상시 감시
    매도_발생 = 매도체크(exchange, symbol, stock_code, default_invest_money, ohlcv, is_daily_buy_turn=is_daily_buy_turn)
    if 매도_발생:
        print("매도 완료. 계좌 정리 대기를 위해 10초 대기 중...")
        time.sleep(10)
    elif is_daily_buy_turn:
        # 매도 메시지와 매수 메시지 순차 전송 시 텔레그램 수신 순서 보장을 위해 2초 대기
        time.sleep(2)

    # 4. [일일 매수 체크] : 미국 장 마감 10분 전(평일 15:51 ET) 턴에만 실행!
    if is_daily_buy_turn:
        print(f"[일일 매수 실행] {schedule_desc} (현재: {now_et.strftime('%H:%M:%S')} {et_tz_name} / {now_utc.strftime('%H:%M:%S')} UTC) 매수 조건을 점검합니다.")
        매수체크(exchange, symbol, stock_code, default_invest_money, ohlcv)
    else:
        print(f"[매수 스킵] 현재 시각({now_et.strftime('%H:%M')} {et_tz_name} / {now_utc.strftime('%H:%M')} UTC)은 매수 시간({schedule_desc})이 아니므로 매수를 건너뜁니다.")


def 전략_1회실행():
    # ✅ 장 열림 여부도 LS 기준 (단, IGNORE_MARKET_OPEN=True 면 장 상태 무시하고 무조건 실행)
    if IGNORE_MARKET_OPEN or LsiUS.IsMarketOpen():
        exchange = create_exchange()
        symbol = resolve_symbol(exchange)
        매수매도프로세스(exchange, symbol, "SOXL", INITIAL_INVEST_USDT)
    else:
        msg = (
            "⏹️ 시장 닫힘 알림\n"
            "📅 떨사오팔일반 전략 오늘은 거래가 실행되지 않았습니다.\n"
            "🔹 이유: 미국 주식 시장 휴장/폐장 시간(LS 기준)\n"
            "💡 자동매매는 시장 개장 시 다시 시작됩니다."
        )
        telegram_alert.SendMessage(msg)
        print("시장 닫힘: 떨사오팔 일반전략 매수와 매도 조건이 실행되지 않았습니다.")


def 전략_재실행_루프():
    start_ts = time.time()
    attempt = 0
    last_err = ""

    while (time.time() - start_ts) < STRATEGY_RETRY_MAX_WAIT_SEC:
        attempt += 1
        try:
            print(f"[전략 실행 시도 {attempt}] LS주문 + LS데이터 떨사오팔 일반전략")
            전략_1회실행()
            return
        except RuntimeError as e:
            last_err = str(e)

            remain_sec = int(STRATEGY_RETRY_MAX_WAIT_SEC - (time.time() - start_ts))
            if remain_sec <= 0:
                break

            print(f"[전략 재실행 대기] {last_err}")
            print(f"{STRATEGY_RETRY_INTERVAL_SEC}초 후 재실행 / 남은 최대 대기 시간: {remain_sec}초")
            time.sleep(STRATEGY_RETRY_INTERVAL_SEC)

    print(f"[전략 종료] 최대 대기 시간 초과 - 마지막 오류: {last_err}")


if __name__ == "__main__":
    전략_재실행_루프()
