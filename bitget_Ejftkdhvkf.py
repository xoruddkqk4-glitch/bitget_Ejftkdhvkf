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
from datetime import datetime, timezone
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

# --- 시간 기준: 한국시간(KST) 고정 (서버 타임존과 무관하게 통일) ---
KST = ZoneInfo("Asia/Seoul")

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
    직전 '완성봉(index -2)'의 종가가 TREND_MA_PERIOD MA 아래로 꺾였는지 판정.
    반환: (is_broken, prev_close, ma_val)
    """
    try:
        min_candles = TREND_MA_PERIOD + 2
        ohlcv = exchange.fetch_ohlcv(symbol, timeframe=TREND_TIMEFRAME, limit=TREND_OHLCV_LIMIT)
        if not ohlcv or len(ohlcv) < min_candles:
            print(f"[{TREND_TIMEFRAME} {TREND_MA_PERIOD}MA] 데이터 부족 ({len(ohlcv) if ohlcv else 0}개 수신, 최소 {min_candles}개 필요)")
            return False, 0.0, 0.0

        df_4h = pd.DataFrame(ohlcv, columns=["timestamp", "open", "high", "low", "close", "volume"])
        df_4h["close"] = pd.to_numeric(df_4h["close"], errors="coerce")

        # index -1 은 현재 실시간 진행 중인 미완성 캔들이므로 직전 완성봉(index -2) 기준 판정
        # 직전 완성봉 포함 최근 TREND_MA_PERIOD개 완성봉 구간
        completed_n = df_4h["close"].iloc[-(TREND_MA_PERIOD + 1):-1]
        ma_val = float(completed_n.mean())
        prev_close = float(df_4h["close"].iloc[-2])

        is_broken = prev_close < ma_val
        print(f"[{TREND_TIMEFRAME} {TREND_MA_PERIOD}MA 분석] 직전완성봉 종가={prev_close:.4f}, {TREND_MA_PERIOD}MA={ma_val:.4f}, 꺾임(이탈)={is_broken}")
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

def sync_round_state_from_excel(df, stock_code):
    """
    상태 파일이 비어있거나 특정 차수가 누락된 경우,
    엑셀 매매기록의 미매도 차수들을 읽어 자동으로 stage=1 로 복원/동기화(Self-healing).
    """
    state = load_round_state()
    stock_dict = state.setdefault(stock_code, {})
    rounds = stock_dict.setdefault("rounds", {})

    if df is not None and not df.empty and "종목" in df.columns:
        stock_data = df[df["종목"] == stock_code]
        for _, row in stock_data.iterrows():
            if pd.notna(row.get("매도가", None)):
                r_key = str(int(row["차수"]))
                rounds.pop(r_key, None)
                continue

            r_key = str(int(row["차수"]))
            buy_price = float(row["매수가"])
            qty = float(row["수량"])
            target_price = buy_price * (1.0 + SELL_PROFIT_PCT)

            if r_key not in rounds:
                rounds[r_key] = {
                    "stage": 1,
                    "buy_price": buy_price,
                    "quantity": qty,
                    "target_price": target_price,
                    "preservation_line": target_price,
                    "stage2_entered_at": None,
                }
    save_round_state(state)
    return state

def update_round_stage(stock_code, round_number, stage, preservation_line=None, target_price=None, buy_price=None, quantity=None):
    state = load_round_state()
    rounds = state.setdefault(stock_code, {}).setdefault("rounds", {})
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

MAX_HOLD_DAYS_NORMAL = 20
MAX_HOLD_DAYS_HIGH_DROP = 50

# 고점 계산 기간 (2026-08-06: LS떨사오팔봇과 통일, 백테스팅 최적값)
LOOKBACK_DAYS = 40

HIGH_DROP_THRESHOLD = 0.40   # 고점대비 하락 판정 (2026-08-06: 0.35 → 0.40)
HIGH_DROP_INVEST_MULTIPLIER = 1.5   # 하락시 매수 증액 배수 (2026-08-06: 2.0 → 1.5)
HIGH_DROP_ADJ_WEIGHT = 1.5          # 하락시 복리 조정 가중치 (2026-08-06: 2.0 → 1.5)
ADJ_DIVISOR = 7.0

SELL_FEE_RATE = 0.0006  # 0.06%

# =========================
# ✅ 2단계 추세추종 설정 (상승 추세 판정 기준)
# =========================
TREND_TIMEFRAME = "4h"               # 추세 판정 캔들 시간프레임 (예: "1h", "4h", "1d")
TREND_MA_PERIOD = 5                  # 이동평균선 기간 (예: 5 → 5MA, 10 → 10MA)
TREND_OHLCV_LIMIT = 10               # 추세 판정용 OHLCV 조회 캔들 수
PRESERVATION_LINE_MULTIPLIER = 0.998  # 보존선 방어 매도 계수 (현재가 ≤ 보존선 × 이 값이면 매도)

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
    for round_num in range(1, 40):
        if round_num not in stock_data["차수"].values or (
            round_num in stock_data["차수"].values and pd.notna(stock_data.loc[stock_data["차수"] == round_num, "매도가"]).all()
        ):
            return round_num
    return None


# =========================
# ✅ 매도 체크
#    데이터=LS / 주문=LS
# =========================
# =========================
# ✅ 매도 체크 (2단계 추세 추종 & 보존선 방어)
#    데이터=LS / 주문=Bitget 선물
# =========================
def 매도체크(exchange, symbol, stock_code, default_invest_money, ohlcv=None):
    df = 매매기록불러오기()
    stock_data = df[df["종목"] == stock_code]

    # 상태 파일 동기화 (기존 엑셀 미매도건 자동 복원)
    state = sync_round_state_from_excel(df, stock_code)

    # ✅ 현재가 데이터는 LS
    current_price = LsiUS.GetCurrentPrice(stock_code)
    today = datetime.now(KST).date()
    sell_orders = []

    current_invest_money = 투자금시트_초기화_및_불러오기(stock_code, default_invest_money)
    highest_price_30d = 과거_고점찾기(stock_code, ohlcv)

    # 미매도 차수 중 2단계(추세추종) 진입 차수가 있으면 4시간봉 5MA 지표 사전 조회
    rounds_dict = state.get(stock_code, {}).get("rounds", {})
    has_stage2 = any(r.get("stage") == 2 for r in rounds_dict.values())
    is_ma5_broken = False
    prev_close_4h = 0.0
    ma5_4h = 0.0
    if has_stage2:
        is_ma5_broken, prev_close_4h, ma5_4h = get_trend_ma_status(exchange, symbol)

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
        quantity = float(row["수량"])
        round_number = int(row["차수"])
        target_price = buy_price * (1.0 + SELL_PROFIT_PCT)
        max_hold_days = MAX_HOLD_DAYS_HIGH_DROP if row["고점대비하락조건"] else MAX_HOLD_DAYS_NORMAL

        r_key = str(round_number)
        r_info = rounds_dict.get(r_key, {})
        current_stage = r_info.get("stage", 1)
        preservation_line = float(r_info.get("preservation_line", target_price))

        # 1) 타임컷 우선 체크
        if (today - buy_date).days >= max_hold_days:
            print(f"{stock_code} {round_number}차 떨사오팔 강제 매도 - {max_hold_days}일 경과")
            sell_orders.append((index, quantity, current_price * 0.98, f"보유기간 경과({max_hold_days}일)", buy_price, buy_date, round_number))

            msg = (
                f"⚠️ 떨사오팔 일반 전략 {stock_code} {round_number}차 강제 매도 알림\n"
                f"🔹 사유: 보유기간 경과({max_hold_days}일)\n"
                f"📅 매수일: {buy_date} | 오늘: {today}\n"
                f"📊 보유 수량: {quantity}주\n"
            )
            telegram_alert.SendMessage(msg)
            continue

        # 2) 1단계: 1차 목표가 감시
        if current_stage == 1:
            if current_price >= target_price:
                update_round_stage(stock_code, round_number, stage=2, preservation_line=target_price, target_price=target_price, buy_price=buy_price, quantity=quantity)
                print(f"{stock_code} {round_number}차 1차 목표가 돌파 -> 2단계(추세추종) 승격!")
                msg = (
                    f"🎯 [1차 목표가 돌파] {stock_code} {round_number}차 2단계 승격\n"
                    f"🔹 진입 단계: [2단계 추세 추종 모드]\n"
                    f"📈 1차 목표가: {target_price:.2f} (수익률 +{SELL_PROFIT_PCT*100:.1f}%)\n"
                    f"💲 현재가: {current_price:.2f}\n"
                    f"🛡️ 최소 이익 보존선: {target_price:.2f}\n"
                    f"💡 {TREND_TIMEFRAME}봉 {TREND_MA_PERIOD}MA 꺾임 시 초과 익절 또는 보존선 하향 터치({PRESERVATION_LINE_MULTIPLIER}) 시 즉시 방어 매도합니다."
                )
                telegram_alert.SendMessage(msg)
            else:
                print(f"{stock_code} {round_number}차 1단계 감시 중: 현재가 {current_price:.2f} < 목표가 {target_price:.2f}")

        # 3) 2단계: 추세추종 (시나리오 ② 보존선 방어 or 시나리오 ① 4h 5MA 이탈)
        elif current_stage == 2:
            # 시나리오 ②: 최소 이익 보존선 방어 매도 (현재가 <= 보존선 * 0.998)
            if current_price <= preservation_line * PRESERVATION_LINE_MULTIPLIER:
                print(f"{stock_code} {round_number}차 보존선 방어 매도 조건 충족: 현재가 {current_price:.2f} <= {preservation_line * PRESERVATION_LINE_MULTIPLIER:.2f}")
                sell_orders.append((index, quantity, current_price * 0.98, f"최소 이익 보존선 방어(보존선 {preservation_line:.2f} × {PRESERVATION_LINE_MULTIPLIER} 대비 하향 터치)", buy_price, buy_date, round_number))
            # 시나리오 ①: 추세 꺾임 익절 (4시간봉 5MA 이탈)
            elif is_ma5_broken:
                print(f"{stock_code} {round_number}차 추세 꺾임 익절 조건 충족: 4시간봉 종가 {prev_close_4h:.2f} < 5MA {ma5_4h:.2f}")
                sell_orders.append((index, quantity, current_price * 0.98, f"추세 꺾임 익절(4시간봉 완성봉 {prev_close_4h:.2f} < 5MA {ma5_4h:.2f})", buy_price, buy_date, round_number))
            else:
                print(f"{stock_code} {round_number}차 2단계 홀딩 중 (현재가 {current_price:.2f} > 보존선 {preservation_line:.2f}, {TREND_TIMEFRAME} {TREND_MA_PERIOD}MA 상승 추세 유지)")

    for index, quantity, sell_price, sell_reason, buy_price, buy_date, round_number in sell_orders:
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

        df.at[index, "매도가"] = real_sell_price
        df.at[index, "매도일"] = datetime.now(KST).strftime("%Y-%m-%d")
        df.at[index, "매도수량"] = real_sell_qty

        try:
            buy_price_f = float(buy_price)
        except Exception:
            buy_price_f = float(pd.to_numeric(pd.Series([buy_price]), errors="coerce").iloc[0])

        sell_price_f = float(real_sell_price)
        net_sell = sell_price_f * (1.0 - SELL_FEE_RATE)

        if buy_price_f is None or pd.isna(buy_price_f) or buy_price_f <= 0:
            profit_pct = 0.0
        else:
            profit_pct = ((net_sell / buy_price_f) - 1.0) * 100.0

        weight = HIGH_DROP_ADJ_WEIGHT if bool(df.at[index, "고점대비하락조건"]) else 1.0
        adj_pct = (profit_pct / ADJ_DIVISOR) * weight

        next_invest_money = int(round(current_invest_money * (1.0 + (adj_pct / 100.0))))

        df.at[index, "수수료율(매도)"] = 0.06
        df.at[index, "순매도가(수수료반영)"] = net_sell
        df.at[index, "수익률(%)"] = profit_pct
        df.at[index, "조정률(수익률/7)(%)"] = adj_pct
        df.at[index, "다음투자금"] = next_invest_money

        투자금시트_업데이트(stock_code, next_invest_money, profit_pct, adj_pct)
        current_invest_money = next_invest_money

        # 상태 파일에서 해당 차수 제거
        remove_round_from_state(stock_code, round_number)

        drop_rate = 0.0
        if highest_price_30d and highest_price_30d > 0:
            drop_rate = (highest_price_30d - current_price) / highest_price_30d * 100.0

        msg = (
            f"📌 떨사오팔 일반전략 {stock_code} {df.at[index, '차수']}차 매도 발생\n"
            f"🔹 매도 사유: {sell_reason}\n"
            f"📉 매수가(기록): {buy_price_f:.2f}\n"
            f"📈 매도가(Bitget 체결): {sell_price_f:.4f}\n"
            f"📊 매도 수량: {quantity}주\n"
            f"📅 매수일: {buy_date}\n"
            f"📅 매도일: {datetime.now(KST).strftime('%Y-%m-%d')}\n"
            f"💸 매도수수료: 0.06% (매도만)\n"
            f"✅ 수익률(수수료반영): {profit_pct:.2f}%\n"
            f"✅ 조정률(수익률/{int(ADJ_DIVISOR)}×가중치 {weight}): {adj_pct:.2f}%\n"
            f"💰 다음 투자금(종목별, 정수): {next_invest_money}\n"
            f"📊 고점 대비 낙폭률(현재가 기준): {drop_rate:.2f}%"
        )
        telegram_alert.SendMessage(msg)

    if not sell_orders:
        remaining = 0 if stock_data.empty else len(stock_data[pd.isna(stock_data["매도가"])])
        # 5분 크론탭에서 매 5분마다 텔레그램을 보내면 알림 폭탄이 되므로 매시 51분(정기 브리핑)에만 발송
        # 단, 2단계 추세추종 차수가 0개인 경우 1시간 단위 정기 메시지 발송 생략
        now_utc = datetime.now(timezone.utc)
        if now_utc.minute == 51:
            latest_state = load_round_state()
            latest_rounds = latest_state.get(stock_code, {}).get("rounds", {})
            stage2_cnt = sum(1 for r in latest_rounds.values() if r.get("stage") == 2)
            if stage2_cnt > 0:
                msg = (
                    f"⏸️ {stock_code} 매도 현황 브리핑 (정기)\n"
                    f"🔹 상태: 목표가 미도달 & 2단계 추세 유지 중\n"
                    f"📈 현재가(LS): {current_price:.2f}\n"
                    f"🧾 미매도 보유 차수: {remaining}개 (2단계 추세추종: {stage2_cnt}개)"
                )
                telegram_alert.SendMessage(msg)
            else:
                print(f"⏸️ {stock_code} 매도 없음 (2단계 추세추종 0개로 1시간 정기 브리핑 생략 / 잔여 {remaining}개)")
        else:
            print(f"⏸️ {stock_code} 매도 없음 (목표가 미도달 or 2단계 홀딩 중 / 잔여 {remaining}개)")

    매매기록저장(df)
    return sell_orders


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

    # 차수별 상태 파일 등록 (새로운 매수는 1단계로 진입)
    target_price = float(buy_price) * (1.0 + SELL_PROFIT_PCT)
    update_round_stage(
        stock_code,
        round_number,
        stage=1,
        preservation_line=target_price,
        target_price=target_price,
        buy_price=float(buy_price),
        quantity=float(quantity)
    )
    print(f"{stock_code} {round_number}차 매수 기록 및 상태(1단계) 등록 완료 - 매수가: {buy_price}, 수량: {quantity}, 고점대비하락조건: {high_drop_condition}")


# =========================
# ✅ 매수 체크
#    데이터=LS / 주문=Bitget 선물
# =========================
def 매수체크(exchange, symbol, stock_code, default_invest_money, ohlcv=None):
    df = 매매기록불러오기()
    stock_data = df[df["종목"] == stock_code]

    # ✅ 당일(KST 기준) 이미 매수한 이력이 있는지 사전 점검 (5분 크론탭 급락 시 연속 매수 원천 차단)
    today_str = datetime.now(KST).strftime("%Y-%m-%d")
    if not stock_data.empty and "날짜" in stock_data.columns:
        bought_today = stock_data[stock_data["날짜"].astype(str).str.startswith(today_str)]
        if not bought_today.empty:
            rounds_bought = bought_today["차수"].tolist()
            print(f"⏸️ [당일 매수 완료] {stock_code}는 오늘({today_str}) 이미 {rounds_bought}차 매수가 완료되었습니다. (추가 매수 차단)")
            return False

    next_round = 가능다음차수(stock_data) if not stock_data.empty else 1

    invest_money = 투자금시트_초기화_및_불러오기(stock_code, default_invest_money)

    # ✅ 데이터는 LS
    current_price = LsiUS.GetCurrentPrice(stock_code)
    if ohlcv is None:
        ohlcv = LsCommon.GetOhlcv("US", stock_code)
    open_price = ohlcv["open"].iloc[-1]

    highest_price = 과거_고점찾기(stock_code, ohlcv)
    high_drop_condition = (highest_price - current_price) / highest_price >= HIGH_DROP_THRESHOLD if highest_price else False

    if current_price <= open_price * (1.0 - BUY_DROP_PCT):
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
            f"🔹 매수 사유: 현재가 ≤ 시가×{(1.0-BUY_DROP_PCT):.3f}"
            f"{' + 고점 대비 하락조건 충족' if high_drop_condition else ''}\n"
            f"📈 시가(LS): {open_price:.2f} | 💲 현재가(LS): {current_price:.2f} | 📉 매수체결가(Bitget): {real_buy_price:.4f}\n"
            f"💵 투자금(종목별, 정수): {invest_money}{invest_suffix}\n"
            f"📊 주문 수량: {buy_amt}주\n"
            f"📅 매수일: {datetime.now(KST).strftime('%Y-%m-%d')}\n"
            f"📊 고점 대비 낙폭률(LS): {drop_rate:.2f}%"
        )
        telegram_alert.SendMessage(msg)
        return True

    else:
        msg = (
            f"⏸️ {stock_code} 매수 보류\n"
            f"🔹 조건: 현재가 ≤ 시가×{(1.0-BUY_DROP_PCT):.3f}\n"
            f"📈 시가(LS): {open_price:.2f}\n"
            f"💲 현재가(LS): {current_price:.2f}\n"
            f"🎯 기준가: {(open_price*(1.0-BUY_DROP_PCT)):.2f}\n"
            f"메모: 떨사오팔 일반 전략 조건 미충족으로 매수하지 않습니다."
        )
        telegram_alert.SendMessage(msg)
        return False


# =========================
# ✅ 매수/매도 프로세스
# =========================
def 매수매도프로세스(exchange, symbol, stock_code, default_invest_money):
    # ✅ OHLCV 데이터는 LS
    ohlcv = LsCommon.GetOhlcv("US", stock_code)

    # 1. [상시 매도 감시] : 5분마다 24시간 항상 실행
    매도_발생 = 매도체크(exchange, symbol, stock_code, default_invest_money, ohlcv)
    if 매도_발생:
        print("매도 완료. 계좌 정리 대기를 위해 10초 대기 중...")
        time.sleep(10)

    # 2. [일일 매수 체크] : 평일 19:51 (UTC, 미국 장 마감 10분 전) 턴에만 실행!
    now_utc = datetime.now(timezone.utc)
    is_weekday = now_utc.weekday() < 5                 # 월~금 (0~4)
    is_buy_minute = (now_utc.hour == 19 and now_utc.minute == 51)  # 19시 51분
    force_buy = os.environ.get("FORCE_BUY_CHECK", "0") == "1"

    if (is_weekday and is_buy_minute) or force_buy:
        print(f"[일일 매수 실행] 평일 19:51 UTC (현재: {now_utc.strftime('%H:%M:%S')} UTC) 매수 조건을 점검합니다.")
        매수체크(exchange, symbol, stock_code, default_invest_money, ohlcv)
    else:
        print(f"[매수 스킵] 현재 시각({now_utc.strftime('%H:%M')} UTC)은 매수 시간(평일 19:51 UTC)이 아니므로 매수를 건너뜁니다.")


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
