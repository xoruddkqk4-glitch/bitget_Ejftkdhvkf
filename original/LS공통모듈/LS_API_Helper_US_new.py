# -*- coding: utf-8 -*-
"""
=============================================================================
LS_API_Helper_US.py  —  LS증권 해외주식(미국) 헬퍼 (자체 구현본)

LS증권 OpenAPI 공식 명세를 참고하여 직접 작성했습니다.
KIS 미국 헬퍼와 동일한 함수 인터페이스를 제공합니다.

[현재 구현 — 받은 시세/계좌 문서 기반]
  · GetCurrentPrice : 현재가         (g3101 해외주식 현재가)
  · GetHoga         : 10호가          (g3106 현재가호가)
  · GetExrt         : 환율            (g3104 종목정보의 exrate)
  · GetStockName    : 종목명          (g3104)
  · PriceAdjust     : 소수점 자릿수 보정 (floatpoint)
  · GetBalance      : 외화 잔고/평가   (COSOQ00201 종합잔고평가)
  · GetMyStockList  : 보유종목        (COSOQ00201)
  · GetOrderList    : 주문체결/미체결  (COSAQ00102)
  · GetOhlcv        : 일봉 (LS_Common 위임 — 미국은 외부소스 권장)
  · IsMarketOpen    : 미국 장중 여부
  · MakeBuy/SellLimitOrder, MakeBuy/SellMarketOrder,
    CancelOrder, ModifyOrder, CancelModifyOrder, CancelAllOrders, SellAllStock
  → 호출 시 NotImplementedError 안내

거래소코드: 81=뉴욕/아멕스, 82=나스닥. keysymbol = exchcd + symbol (예: 82TSLA)
=============================================================================
"""

import json
import time
import requests
import pandas as pd
from datetime import datetime

import LS_Common_new as Common

_MARKET_URL = "/overseas-stock/market-data"
_ACCNO_URL = "/overseas-stock/accno"

# 거래소 코드 후보 (자동 탐색: 나스닥 우선, 다음 뉴욕/아멕스)
_EXCH_LIST = ["82", "81"]


def _post(url_path, tr_cd, in_block_name, in_block, out_keys):
    """해외 API 공통 POST. 응답에서 지정한 블록들을 dict로 모아 반환."""
    url = f"{Common.GetUrlBase(Common.GetNowDist())}{url_path}"
    headers = Common.GetCommonHeader(tr_cd=tr_cd)
    body = {in_block_name: in_block}
    try:
        res = requests.post(url, headers=headers, data=json.dumps(body), timeout=10)
        if res.status_code != 200:
            print(f"[{tr_cd}] HTTP {res.status_code}: {res.text[:150]}")
            return None
        j = res.json()
        result = {"rsp_cd": j.get("rsp_cd", ""), "rsp_msg": j.get("rsp_msg", "")}
        for k in out_keys:
            result[k] = j.get(k)
        return result
    except Exception as e:
        print(f"[{tr_cd}] 오류: {e}")
        return None


def _keysymbol(symbol, exchcd):
    """KEY종목코드 = 거래소코드 + 심볼 (예: 82TSLA)"""
    return f"{exchcd}{symbol}"


# ---------------------------------------------------------------------------
# 현재가 (g3101) — 거래소 자동 탐색
# ---------------------------------------------------------------------------
def _query_price_raw(symbol):
    """g3101로 현재가 블록을 조회. 거래소(나스닥→뉴욕) 자동 탐색."""
    for exchcd in _EXCH_LIST:
        in_block = {
            "delaygb": "R",
            "keysymbol": _keysymbol(symbol, exchcd),
            "exchcd": exchcd,
            "symbol": symbol,
        }
        out = _post(_MARKET_URL, "g3101", "g3101InBlock", in_block, ["g3101OutBlock"])
        if out and out.get("g3101OutBlock"):
            blk = out["g3101OutBlock"]
            try:
                if float(blk.get("price", 0) or 0) > 0:
                    blk["_exchcd"] = exchcd
                    return blk
            except Exception:
                continue
    return None


def GetCurrentPrice(stock_code):
    """현재가를 반환합니다. (g3101의 price)"""
    blk = _query_price_raw(stock_code)
    if not blk:
        return None
    try:
        return float(blk.get("price", 0) or 0)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 종목정보 (g3104) — 환율/소수점/종목명
# ---------------------------------------------------------------------------
def _query_info_raw(symbol):
    """g3104 종목정보 블록 조회 (환율 exrate, 소수점 floatpoint 포함)."""
    for exchcd in _EXCH_LIST:
        in_block = {
            "delaygb": "R",
            "keysymbol": _keysymbol(symbol, exchcd),
            "exchcd": exchcd,
            "symbol": symbol,
        }
        out = _post(_MARKET_URL, "g3104", "g3104InBlock", in_block, ["g3104OutBlock"])
        if out and out.get("g3104OutBlock"):
            blk = out["g3104OutBlock"]
            if blk.get("symbol"):
                blk["_exchcd"] = exchcd
                return blk
    return None


def GetStockName(stock_code):
    """종목명(한글)을 반환합니다. (g3104의 korname)"""
    blk = _query_info_raw(stock_code)
    return blk.get("korname", "") if blk else None


def GetExrt(stock_code="TSLA"):
    """
    원/달러 환율을 반환합니다. (g3104 종목정보의 exrate)
    대표 종목으로 조회해 환율만 추출합니다. 실패 시 0.
    """
    blk = _query_info_raw(stock_code)
    if not blk:
        return 0
    try:
        return float(blk.get("exrate", 0) or 0)
    except Exception:
        return 0


# ---------------------------------------------------------------------------
# 소수점 자릿수 보정 (floatpoint)
# ---------------------------------------------------------------------------
def PriceAdjust(price, stock_code=None, floatpoint=None):
    """
    미국 주문가를 소수점 자릿수에 맞게 보정합니다.
    floatpoint 직접 지정 시 그 자릿수로, 미지정+종목코드 있으면 g3104로 조회,
    둘 다 없으면 기본 2자리.
    """
    if floatpoint is None:
        if stock_code is not None:
            blk = _query_info_raw(stock_code)
            try:
                floatpoint = int(blk.get("floatpoint", 2)) if blk else 2
            except Exception:
                floatpoint = 2
        else:
            floatpoint = 2
    # LS floatpoint는 보통 4지만, 미국 주식 주문은 통상 소수 2자리로 맞춤.
    # 과도한 자릿수를 막기 위해 4 이상이면 2로 제한.
    fp = int(floatpoint)
    if fp > 2:
        fp = 2
    return round(float(price), fp)


# ---------------------------------------------------------------------------
# 호가 (g3106)
# ---------------------------------------------------------------------------
def GetHoga(stock_code):
    """
    10호가를 조회합니다. (g3106 현재가호가)
    반환: {'ask': [...], 'bid': [...], 'ask_rem': [...], 'bid_rem': [...], 'last': 현재가}
    """
    for exchcd in _EXCH_LIST:
        in_block = {
            "delaygb": "R",
            "keysymbol": _keysymbol(stock_code, exchcd),
            "exchcd": exchcd,
            "symbol": stock_code,
        }
        out = _post(_MARKET_URL, "g3106", "g3106InBlock", in_block, ["g3106OutBlock"])
        if not out or not out.get("g3106OutBlock"):
            continue
        blk = out["g3106OutBlock"]
        if not blk.get("price"):
            continue
        ask, bid, ask_rem, bid_rem = [], [], [], []
        for i in range(1, 11):
            a = blk.get(f"offerho{i}")
            b = blk.get(f"bidho{i}")
            ar = blk.get(f"offerrem{i}")
            br = blk.get(f"bidrem{i}")
            if a not in (None, "", "0", 0):
                ask.append(float(a))
            if b not in (None, "", "0", 0):
                bid.append(float(b))
            if ar not in (None, ""):
                ask_rem.append(float(ar))
            if br not in (None, ""):
                bid_rem.append(float(br))
        return {
            "ask": ask, "bid": bid,
            "ask_rem": ask_rem, "bid_rem": bid_rem,
            "last": float(blk.get("price", 0) or 0),
        }
    return None


# ---------------------------------------------------------------------------
# 잔고 / 보유종목 (COSOQ00201 종합잔고평가)
# ---------------------------------------------------------------------------
def _get_balance_raw():
    """해외 종합잔고평가(COSOQ00201) 원본 응답."""
    in_block = {
        "RecCnt": 1,
        "BaseDt": Common.GetNowDateStr("KR"),
        "CrcyCode": "USD",
        "AstkBalTpCode": "00",
    }
    return _post(_ACCNO_URL, "COSOQ00201", "COSOQ00201InBlock1", in_block,
                 ["COSOQ00201OutBlock2", "COSOQ00201OutBlock3", "COSOQ00201OutBlock4"])


def GetBalance():
    """
    해외 계좌 잔고 요약을 반환합니다. (COSOQ00201)
    반환 dict:
      TotalMoneyKRW : 예탁자산환산평가금액(원화)
      StockMoneyKRW : 주식환산평가금액(원화)
      RevenueKRW    : 환산평가손익(원화)
      DollarMoney   : 외화예수금(USD)
      OrderableUSD  : 외화주문가능금액(USD)
      Exrt          : 기준환율
    """
    fail = {"TotalMoneyKRW": 0, "StockMoneyKRW": 0, "RevenueKRW": 0,
            "DollarMoney": 0, "OrderableUSD": 0, "Exrt": 0}
    out = _get_balance_raw()
    if not out:
        return fail
    def _n(v):
        try:
            return float(v)
        except Exception:
            return 0.0
    blk2 = out.get("COSOQ00201OutBlock2") or {}
    blk3_list = out.get("COSOQ00201OutBlock3") or []
    blk3 = {}
    for row in blk3_list:
        if row.get("CrcyCode") == "USD":
            blk3 = row
            break
    if not blk3 and blk3_list:
        blk3 = blk3_list[0]
    return {
        "TotalMoneyKRW": _n(blk2.get("DpsastConvEvalAmt", 0)),
        "StockMoneyKRW": _n(blk2.get("StkConvEvalAmt", 0)),
        "RevenueKRW": _n(blk2.get("ConvEvalPnlAmt", 0)),
        "DollarMoney": _n(blk3.get("FcurrDps", 0)),
        "OrderableUSD": _n(blk3.get("FcurrOrdAbleAmt", 0)),
        "Exrt": _n(blk3.get("BaseXchrat", 0)),
    }


def GetMyStockList():
    """
    보유 해외종목 리스트를 반환합니다. (COSOQ00201OutBlock4)
    반환: [{StockCode, StockName, StockAmt, StockAvgPrice, StockNowPrice,
            StockEvalMoney(외화평가), StockRevenue(외화손익),
            StockRevenueRate, StockSellAbleAmt}, ...]
    """
    out = _post(_ACCNO_URL, "COSOQ00201", "COSOQ00201InBlock1",
                {"RecCnt": 1, "BaseDt": Common.GetNowDateStr("KR"),
                 "CrcyCode": "USD", "AstkBalTpCode": "00"},
                ["COSOQ00201OutBlock4"])
    if not out:
        return []
    rows = out.get("COSOQ00201OutBlock4") or []
    if isinstance(rows, dict):
        rows = [rows]
    def _n(v):
        try:
            return float(v)
        except Exception:
            return 0.0
    result = []
    for r in rows:
        qty = _n(r.get("AstkBalQty", 0))
        if qty <= 0:
            continue
        result.append({
            "StockCode": r.get("ShtnIsuNo", ""),
            "StockName": r.get("JpnMktHanglIsuNm", ""),
            "StockAmt": qty,
            "StockAvgPrice": _n(r.get("FcstckUprc", 0)),
            "StockNowPrice": _n(r.get("OvrsScrtsCurpri", 0)),
            "StockEvalMoney": _n(r.get("FcurrEvalAmt", 0)),
            "StockRevenue": _n(r.get("FcurrEvalPnlAmt", 0)),
            "StockRevenueRate": _n(r.get("PnlRat", 0)),
            "StockSellAbleAmt": _n(r.get("AstkSellAbleQty", 0)),
        })
    return result


# ---------------------------------------------------------------------------
# 매수가능 (예수금 기반) — COSOQ02701 + 현재가
# ---------------------------------------------------------------------------
def CheckPossibleBuyInfo(stock_code, price):
    """
    해외 매수가능 정보를 조회합니다.
    예수금 조회(COSOQ02701)의 외화주문가능금액과 주문단가로 수량을 계산합니다.
    반환 dict: Money(외화주문가능금액), Amount(가능수량), Exrt(기준환율)
    """
    fail = {"Money": 0, "Amount": 0, "Exrt": 0}
    out = _post(_ACCNO_URL, "COSOQ02701", "COSOQ02701InBlock1",
                {"RecCnt": 1, "CrcyCode": "USD"}, ["COSOQ02701OutBlock3"])
    if not out:
        return fail
    rows = out.get("COSOQ02701OutBlock3") or []
    if isinstance(rows, dict):
        rows = [rows]
    usd = {}
    for r in rows:
        if r.get("CrcyCode") == "USD":
            usd = r
            break
    def _n(v):
        try:
            return float(v)
        except Exception:
            return 0.0
    # 외화주문가능 또는 가환전주문가능 중 큰 값
    money = max(_n(usd.get("FcurrOrdAbleAmt", 0)), _n(usd.get("PrexchOrdAbleAmt", 0)))
    exrt = _n(usd.get("BaseXchrat", 0))
    amount = int(money // float(price)) if price and float(price) > 0 else 0
    return {"Money": money, "Amount": amount, "Exrt": exrt}


def AdjustPossibleAmt(stock_code, want_amt, price):
    """원하는 매수수량이 가능수량을 넘지 않게 보정합니다."""
    info = CheckPossibleBuyInfo(stock_code, price)
    possible = info["Amount"]
    want_amt = int(want_amt)
    return min(want_amt, possible) if possible > 0 else 0


# ---------------------------------------------------------------------------
# 주문체결/미체결 내역 (COSAQ00102)
# ---------------------------------------------------------------------------
def GetOrderList(stock_code="", side="ALL", status="ALL"):
    """
    해외 주문체결/미체결 내역. (COSAQ00102)
      side   : "ALL"=전체, "BUY"=매수, "SELL"=매도
      status : "ALL"=전체, "OPEN"=미체결, "CLOSE"=체결
    (봇 호환 시그니처. 기존 exec_yn 방식 호출도 status 자리로 들어오면 변환)

    반환 리스트의 각 항목 키(봇이 사용하는 키 포함):
      OrderNum, OrderAmt, OrderType("LIMIT"/"MARKET"),
      OrderMarket(81/82), OrderStockCode, OrderSide("BUY"/"SELL"),
      OrderPrice, ExecAmt, RemainAmt, OrgOrderNum, Status
    """
    # status 코드 변환 (COSAQ00102 ExecYn)
    exec_yn = "0"                       # 0=전체
    if status == "OPEN":
        exec_yn = "2"                   # 2=미체결
    elif status == "CLOSE":
        exec_yn = "1"                   # 1=체결
    elif status in ("0", "1", "2"):     # 구버전 exec_yn 직접 전달 호환
        exec_yn = status
    # side 코드 변환 (BnsTpCode)
    bns_tp = "0"                        # 0=전체
    if side == "BUY":
        bns_tp = "2"
    elif side == "SELL":
        bns_tp = "1"

    def _n(v):
        try:
            return float(v)
        except Exception:
            return 0.0

    result = []
    # 거래소(나스닥/뉴욕) 모두 조회해 합칩니다.
    for mkt in ("82", "81"):
        in_block = {
            "RecCnt": 1,
            "QryTpCode": "1",
            "BkseqTpCode": "1",
            "OrdMktCode": mkt,
            "BnsTpCode": bns_tp,
            "IsuNo": stock_code,
            "SrtOrdNo": 999999999,
            "OrdDt": Common.GetNowDateStr("US"),
            "ExecYn": exec_yn,
            "CrcyCode": "USD",
            "ThdayBnsAppYn": "1",
            "LoanBalHldYn": "0",
        }
        out = _post(_ACCNO_URL, "COSAQ00102", "COSAQ00102InBlock1", in_block,
                    ["COSAQ00102OutBlock3"])
        if not out:
            continue
        rows = out.get("COSAQ00102OutBlock3") or []
        if isinstance(rows, dict):
            rows = [rows]
        for r in rows:
            ord_ptn = r.get("OrdPtnNm", "")        # 매수/매도(한글)
            side_str = "BUY" if "매수" in ord_ptn else ("SELL" if "매도" in ord_ptn else "")
            ptn_code = str(r.get("OrdprcPtnCode", "00"))
            ord_type = "MARKET" if ptn_code == "03" else "LIMIT"
            result.append({
                "OrderNum": int(r.get("OrdNo", 0) or 0),
                "OrderAmt": int(_n(r.get("OrdQty", 0))),
                "OrderType": ord_type,
                "OrderMarket": r.get("OrdMktCode", mkt) or mkt,
                "OrderStockCode": r.get("ShtnIsuNo", "") or stock_code,
                "OrderSide": side_str,
                "OrderPrice": _n(r.get("OvrsOrdPrc", 0)),
                "ExecAmt": int(_n(r.get("ExecQty", 0))),
                "RemainAmt": int(_n(r.get("UnercQty", 0))),
                "OrgOrderNum": int(r.get("OrgOrdNo", 0) or 0),
                "Status": r.get("OrdTrxPtnNm", ""),
                # 기존 키 호환
                "StockCode": r.get("ShtnIsuNo", "") or stock_code,
            })
    return result


def GetOrderExecInfo(stock_code, order_num, market=None):
    """
    특정 주문번호(order_num)의 '실제 체결' 정보를 COSAQ00102(체결내역)에서 조회.

    market : 주문 시장코드("81"=NYSE, "82"=NASDAQ)를 알면 넘겨준다.
             주면 그 시장만 조회하므로 API 호출이 절반으로 줄어
             레이트리밋(IGW00201 '호출 거래건수 초과') 압박이 크게 낮아진다.
             None 이면 기존처럼 82→81 순으로 모두 조회.

    반환:
      {"price": 체결단가(float), "qty": 체결수량(int),
       "amt": 체결금액(float), "raw": 원본row(dict)}
      아직 체결이 확인되지 않으면(미체결/조회불가) None.

    ★ 주의: LS OutBlock 스키마에 따라 '체결단가' 필드명이 계정/버전별로
      다를 수 있어, 아래 여러 후보 키를 순서대로 시도하고, 그래도 없으면
      '체결금액 ÷ 체결수량'으로 역산한다. 첫 실제 체결 시 원본 row 를
      [체결가확인] 로그로 남기므로, 후보에 없는 필드명이면 그 로그를 보고
      PRICE_KEYS 맨 앞에 실제 키를 추가하면 된다.
    """
    try:
        order_num = int(order_num)
    except Exception:
        return None
    if order_num <= 0:
        return None

    def _f(v):
        try:
            return float(v)
        except Exception:
            return 0.0

    # 체결단가 후보 (OvrsOrdPrc=주문가와 대응되는 '해외체결가격'을 최우선)
    PRICE_KEYS = ["OvrsExecPrc", "ExecPrc", "OvrsExecAvrPrc",
                  "ExecAvrPrc", "AvrPrc", "MdiaAvrPrc", "ExecPric"]
    QTY_KEYS = ["ExecQty", "OvrsExecQty"]
    AMT_KEYS = ["OvrsExecAmt", "ExecAmt"]      # 역산용 체결금액 후보

    # 시장코드가 주어지면 그 시장만, 없으면 82→81 모두 조회
    markets = (str(market),) if market in ("81", "82") else ("82", "81")
    for mkt in markets:
        in_block = {
            "RecCnt": 1, "QryTpCode": "1", "BkseqTpCode": "1",
            "OrdMktCode": mkt, "BnsTpCode": "0",
            "IsuNo": stock_code, "SrtOrdNo": 999999999,
            "OrdDt": Common.GetNowDateStr("US"),
            "ExecYn": "1",                     # 1 = 체결분만
            "CrcyCode": "USD",
            "ThdayBnsAppYn": "1", "LoanBalHldYn": "0",
        }
        out = _post(_ACCNO_URL, "COSAQ00102", "COSAQ00102InBlock1", in_block,
                    ["COSAQ00102OutBlock3"])
        if not out:
            continue
        rows = out.get("COSAQ00102OutBlock3") or []
        if isinstance(rows, dict):
            rows = [rows]
        for r in rows:
            try:
                if int(r.get("OrdNo", 0) or 0) != order_num:
                    continue
            except Exception:
                continue

            # 체결수량
            qty = 0
            for k in QTY_KEYS:
                if r.get(k) not in (None, ""):
                    qty = int(_f(r.get(k)))
                    break
            if qty <= 0:
                return None                    # 아직 체결 안 됨

            # 체결단가
            price = 0.0
            for k in PRICE_KEYS:
                if r.get(k) not in (None, ""):
                    price = _f(r.get(k))
                    if price > 0:
                        break
            # 단가 못 찾으면 체결금액 ÷ 체결수량으로 역산
            if price <= 0:
                amt = 0.0
                for k in AMT_KEYS:
                    if r.get(k) not in (None, ""):
                        amt = _f(r.get(k))
                        if amt > 0:
                            break
                if amt > 0 and qty > 0:
                    price = amt / qty

            print(f"[체결가확인] {stock_code} 주문:{order_num} → "
                  f"단가:{price} 수량:{qty} | raw={r}")
            if price > 0:
                return {"price": round(price, 4), "qty": qty,
                        "amt": round(price * qty, 2), "raw": r}
            return None
    return None


# ---------------------------------------------------------------------------
# 일봉 OHLCV — LS API(g3204) 우선 + FDR 폴백
# ---------------------------------------------------------------------------
# 종목별 거래소 코드 캐시 (한 번 성공한 거래소는 다음에 바로 시도)
# 이 캐시 덕분에 같은 종목 재조회 시 거래소 자동 탐색 비용이 사라짐
_EXCHCD_CACHE = {}


def _get_ohlcv_from_ls(stock_code, limit=500, adj_ok="1"):
    """
    LS증권 해외주식 차트 API(g3204)로 일봉 OHLCV 조회.

    거래소 코드 (exchcd) 자동 탐색 (첫 성공 시 캐시):
      82 = 나스닥, 81 = 뉴욕(NYSE), 83 = 아멕스(AMEX)

    Rate Limit: 1초당 1건 (LS API 정책 준수)

    문서: https://openapi.ls-sec.co.kr  /overseas-stock/chart  g3204
    반환: DataFrame (인덱스=날짜, 컬럼=open/high/low/close/volume/value)
          실패 시 None
    """
    url = f"{Common.GetUrlBase(Common.GetNowDist())}/overseas-stock/chart"
    sujung = "Y" if adj_ok == "1" else "N"

    # 캐시된 거래소가 있으면 그것 먼저 시도, 없으면 전체 후보
    if stock_code in _EXCHCD_CACHE:
        cached = _EXCHCD_CACHE[stock_code]
        exchange_codes = [cached] + [e for e in ["82", "81", "83"] if e != cached]
    else:
        exchange_codes = ["82", "81", "83"]   # 나스닥 우선

    for exchcd in exchange_codes:
        keysymbol = f"{exchcd}{stock_code}"

        # 충분히 과거 시점부터 — limit 일치만큼 받기 위해 넉넉히 거슬러 올라감
        # type="NONE" 이면 YYYYMMDD 형식으로 반환 (g3204 가 요구하는 형식)
        sdate = Common.GetFromNowDateStr("US", "NONE", -int(limit * 2 + 365))
        edate = Common.GetNowDateStr("US", "NONE")

        body = {
            "g3204InBlock": {
                "sujung": sujung,
                "delaygb": "R",
                "keysymbol": keysymbol,
                "exchcd": exchcd,
                "symbol": stock_code,
                "gubun": "2",   # 2: 일봉
                "qrycnt": min(int(limit), 500),  # 비압축 최대 500
                "comp_yn": "N",
                "sdate": sdate,
                "edate": edate,
                "cts_date": "",
                "cts_info": "",
            }
        }
        headers = Common.GetCommonHeader(tr_cd="g3204", tr_cont="N", tr_cont_key="")
        headers["content-type"] = "application/json; charset=UTF-8"

        # LS API Rate Limit (1초당 1건) 준수
        # 호출 직전에 충분히 대기해서 직전 호출과 1초 이상 간격 보장
        time.sleep(1.1)

        try:
            res = requests.post(url, headers=headers, data=json.dumps(body), timeout=10)
        except Exception as e:
            print(f"[g3204 {exchcd}] 요청 오류: {e}")
            continue

        if res.status_code != 200:
            # HTTP 500 + 호출제한 메시지면 더 기다린 후 재시도
            if res.status_code == 500 and "거래건수" in res.text:
                print(f"[g3204 {exchcd}] Rate Limit 감지 — 2초 대기 후 재시도")
                time.sleep(2.0)
                try:
                    res = requests.post(url, headers=headers, data=json.dumps(body), timeout=10)
                except Exception as e:
                    print(f"[g3204 {exchcd}] 재시도 오류: {e}")
                    continue
                if res.status_code != 200:
                    print(f"[g3204 {exchcd}] 재시도 후에도 HTTP {res.status_code}")
                    continue
            else:
                print(f"[g3204 {exchcd}] HTTP {res.status_code}: {res.text[:150]}")
                continue

        try:
            j = res.json()
        except Exception as e:
            print(f"[g3204 {exchcd}] JSON 파싱 오류: {e}")
            continue

        if j.get("rsp_cd") != "00000":
            # 거래소가 다를 수 있음 — 다음 거래소 코드로 재시도
            continue

        rows = j.get("g3204OutBlock1", [])
        if not rows:
            continue

        # 정상 응답 — DataFrame 변환
        records = []
        for r in rows:
            d = r.get("date", "")
            try:
                op = float(r.get("open", 0) or 0)
                if d and op > 0:
                    records.append({
                        "date":   d,
                        "open":   op,
                        "high":   float(r.get("high", 0) or 0),
                        "low":    float(r.get("low", 0) or 0),
                        "close":  float(r.get("close", 0) or 0),
                        "volume": float(r.get("volume", 0) or 0),
                        "value":  float(r.get("amount", 0) or 0),
                    })
            except (ValueError, TypeError):
                continue

        if not records:
            continue

        df = pd.DataFrame(records)
        df["date"] = pd.to_datetime(df["date"], format="%Y%m%d")
        df = df.set_index("date").sort_index()

        # 성공한 거래소 코드 캐시 (다음 조회는 바로 그 거래소로)
        _EXCHCD_CACHE[stock_code] = exchcd

        return Common._standardize(df)

    # 모든 거래소 코드 실패
    return None


def _get_ohlcv_from_fdr_us(stock_code, limit=500, adj_ok="1"):
    """
    미국 종목 전용 FDR 폴백 (LS_Common 의 한국용 _get_from_fdr 우회).

    LS_Common._get_from_fdr 는 시간대를 KR 로 가정하므로 미국 종목 조회에 부적합.
    이 함수는 미국 시간대 기준으로 FDR 를 정확히 호출합니다.
    """
    try:
        import FinanceDataReader as fdr
    except ImportError:
        print("[FDR 폴백] FinanceDataReader 가 설치되지 않음")
        return None

    try:
        # 미국 시간대 기준으로 충분히 과거부터
        start = Common.GetFromNowDateStr("US", "BAR", -int(limit * 2 + 365))
        end = Common.GetNowDateStr("US", "BAR")
        # BAR 형식 (YYYY-MM-DD) 로 FDR 호출
        raw = fdr.DataReader(stock_code, start, end)
        if raw is None or len(raw) == 0:
            return None
        close_col = "Adj Close" if (adj_ok == "1" and "Adj Close" in raw.columns) else "Close"
        df = pd.DataFrame({
            "open":   raw["Open"],
            "high":   raw["High"],
            "low":    raw["Low"],
            "close":  raw[close_col],
            "volume": raw["Volume"],
        })
        return Common._standardize(df)
    except Exception as e:
        print(f"[FDR 폴백 미국] {stock_code} 조회 실패: {e}")
        return None


def GetOhlcv(area, stock_code, limit=500, adj_ok="1", source="auto"):
    """
    미국 일봉 OHLCV 조회 (LS API 우선 + FDR 폴백).

    source="auto" : LS API(g3204) 우선, 실패 시 FDR 폴백  ★기본 (권장)
    source="ls"   : LS API(g3204)만 사용 (실패해도 폴백 없음)
    source="fdr"  : FinanceDataReader (yfinance 백본) 만 사용

    Rate Limit: LS API 는 1초당 1건 — 내부에서 자동 준수
    반환: DataFrame (인덱스=날짜, 컬럼=open/high/low/close/volume/value/change)
          최근 limit 개 행. 실패 시 빈 DataFrame.
    """
    df = None

    if source == "ls":
        df = _get_ohlcv_from_ls(stock_code, limit, adj_ok)
    elif source == "fdr":
        df = _get_ohlcv_from_fdr_us(stock_code, limit, adj_ok)
    else:  # "auto" 기본 — LS 우선 → FDR 폴백
        df = _get_ohlcv_from_ls(stock_code, limit, adj_ok)
        if df is None or len(df) == 0:
            print(f"[GetOhlcv] LS API 실패 → FDR 폴백 ({stock_code})")
            df = _get_ohlcv_from_fdr_us(stock_code, limit, adj_ok)

    if df is None or len(df) == 0:
        return pd.DataFrame()
    return df.iloc[-int(limit):]


def GetOhlcvNew(area, stock_code, limit=500, adj_ok="1", source="auto"):
    """GetOhlcv 별칭 (호환용)."""
    return GetOhlcv(area, stock_code, limit, adj_ok, source)


# ---------------------------------------------------------------------------
# 미국 장중 여부 — 미국 동부시간 09:30~16:00 (썸머타임 자동)
# ---------------------------------------------------------------------------
def IsMarketOpen(check_holiday=True, probe_code="SPY"):
    """미국 정규장 시간(동부 09:30~16:00, 평일)이면 True. (서버시간 기준)"""
    now = Common.GetNow("US")
    if now.weekday() >= 5:
        return False
    t = now.hour * 60 + now.minute
    if not (9 * 60 + 30 <= t <= 16 * 60):
        return False
    if check_holiday:
        opened = Common.IsTradingDayByOhlcv("US", probe_code)
        if opened is False:
            return False
    return True


# ---------------------------------------------------------------------------
# 주문 — 미국시장주문 (COSAT00301) / 정정 (COSAT00311)
# ---------------------------------------------------------------------------
_ORDER_URL = "/overseas-stock/order"


def _find_order_mkt(stock_code):
    """주문시장코드(81뉴욕/82나스닥)를 현재가 조회로 자동 판별. 실패 시 82(나스닥)."""
    blk = _query_price_raw(stock_code)
    if blk and blk.get("_exchcd"):
        return blk["_exchcd"]
    return "82"


def _do_order(stock_code, amount, price, ord_ptn, ord_prc_ptn, org_no=0):
    """
    미국시장주문(COSAT00301) 실행 공통.
      ord_ptn     : "01"=매도, "02"=매수, "08"=취소
      ord_prc_ptn : "00"=지정가, "03"=시장가
    반환: {OrderNum, IsuNm} (실패 시 OrderNum=0)
    """
    mkt = _find_order_mkt(stock_code)
    in_block = {
        "RecCnt": 1,
        "OrdPtnCode": ord_ptn,
        "OrgOrdNo": int(org_no),
        "OrdMktCode": mkt,
        "IsuNo": stock_code,
        "OrdQty": int(abs(amount)),
        "OvrsOrdPrc": float(price) if ord_prc_ptn == "00" else 0,
        "OrdprcPtnCode": ord_prc_ptn,
        "BrkTpCode": "",
    }
    # 실패 응답에는 거부 사유(rsp_cd/rsp_msg)를 함께 담아 돌려준다.
    # → 봇이 텔레그램으로 "왜 주문이 안 나갔는지"를 그대로 알릴 수 있다.
    #   (기존 키 OrderNum/IsuNm 은 그대로라 기존 봇 로직에는 영향 없음)
    fail = {"OrderNum": 0, "IsuNm": "", "rsp_cd": "", "rsp_msg": "주문 응답 없음"}
    out = _post(_ORDER_URL, "COSAT00301", "COSAT00301InBlock1", in_block,
                ["COSAT00301OutBlock2"])
    if not out:
        return fail
    blk2 = out.get("COSAT00301OutBlock2") or {}
    ord_no = blk2.get("OrdNo", 0)
    if not ord_no:
        print(f"[US.order] 주문 실패: {out.get('rsp_msg')}")
        fail["rsp_cd"] = out.get("rsp_cd", "")
        fail["rsp_msg"] = out.get("rsp_msg", "") or "사유 미상"
        return fail
    return {"OrderNum": ord_no, "IsuNm": blk2.get("IsuNm", "")}


def MakeBuyLimitOrder(stock_code, amount, price):
    """미국 지정가 매수."""
    price = PriceAdjust(price, stock_code=stock_code)
    return _do_order(stock_code, amount, price, ord_ptn="02", ord_prc_ptn="00")


def MakeSellLimitOrder(stock_code, amount, price):
    """미국 지정가 매도."""
    price = PriceAdjust(price, stock_code=stock_code)
    return _do_order(stock_code, amount, price, ord_ptn="01", ord_prc_ptn="00")


def MakeBuyMarketOrder(stock_code, amount, slippage=0.01):
    """미국 시장가 매수 (OrdprcPtnCode=03)."""
    return _do_order(stock_code, amount, 0, ord_ptn="02", ord_prc_ptn="03")


def MakeSellMarketOrder(stock_code, amount, slippage=0.01):
    """미국 시장가 매도 (OrdprcPtnCode=03)."""
    return _do_order(stock_code, amount, 0, ord_ptn="01", ord_prc_ptn="03")


def CancelOrder(stockcode, order_num):
    """
    미국 주문 취소. (COSAT00301, OrdPtnCode=08 + 원주문번호)
    봇 호환 시그니처: CancelOrder(종목코드, 주문번호)
    미체결 목록에서 해당 주문을 찾아 수량/시장/종목을 채워 취소합니다.
    반환: 성공 시 주문번호(int) / 못 찾으면 안내 문자열 / 실패 시 False
    """
    open_orders = GetOrderList(stockcode, "ALL", "OPEN")
    target = None
    for o in open_orders:
        if o["OrderNum"] == int(order_num):
            target = o
            break
    if target is None:
        return "주문이 존재하지 않습니다."

    mkt = target.get("OrderMarket") or "82"
    isu = target.get("OrderStockCode") or stockcode
    qty = target.get("RemainAmt") or target.get("OrderAmt") or 0
    in_block = {
        "RecCnt": 1,
        "OrdPtnCode": "08",                 # 취소
        "OrgOrdNo": int(order_num),
        "OrdMktCode": mkt,
        "IsuNo": isu,
        "OrdQty": int(qty),
        "OvrsOrdPrc": 0,
        "OrdprcPtnCode": "00",
        "BrkTpCode": "",
    }
    out = _post(_ORDER_URL, "COSAT00301", "COSAT00301InBlock1", in_block,
                ["COSAT00301OutBlock2"])
    if not out:
        return False
    blk2 = out.get("COSAT00301OutBlock2") or {}
    new_no = blk2.get("OrdNo", 0) or 0
    return new_no if new_no else False


def ModifyOrder(org_order_no, stock_code, price):
    """
    미국 정정주문. (COSAT00311, OrdPtnCode=07)
    수량은 0 입력(명세), 가격만 정정. 반환: 새 주문번호(실패 시 0)
    """
    mkt = _find_order_mkt(stock_code)
    price = PriceAdjust(price, stock_code=stock_code)
    in_block = {
        "RecCnt": 1,
        "OrdPtnCode": "07",
        "OrgOrdNo": int(org_order_no),
        "OrdMktCode": mkt,
        "IsuNo": stock_code,
        "OrdQty": 0,                  # 명세: 정정은 0 입력
        "OvrsOrdPrc": float(price),
        "OrdprcPtnCode": "00",
        "BrkTpCode": "",
    }
    out = _post(_ORDER_URL, "COSAT00311", "COSAT00311InBlock1", in_block,
                ["COSAT00311OutBlock2"])
    if not out:
        return 0
    blk2 = out.get("COSAT00311OutBlock2") or {}
    return blk2.get("OrdNo", 0) or 0


def CancelModifyOrder(org_order_no, stock_code, amount=0, price=0, mode="CANCEL"):
    """KIS 호환용. mode='CANCEL'이면 취소, 'MODIFY'면 정정."""
    if mode.upper() == "MODIFY":
        return ModifyOrder(org_order_no, stock_code, price)
    return CancelOrder(stock_code, org_order_no)


def CancelAllOrders(stock_code=""):
    """미체결 해외주문을 모두 취소합니다. 취소 건수를 반환."""
    orders = GetOrderList(stock_code, "ALL", "OPEN")
    cnt = 0
    for o in orders:
        remain = o["RemainAmt"] or o["OrderAmt"]
        if remain > 0:
            r = CancelOrder(o["OrderStockCode"] or stock_code, o["OrderNum"])
            if r and not isinstance(r, str):
                cnt += 1
    return cnt


def SellAllStock(stock_code):
    """특정 해외종목을 시장가로 전량 매도합니다."""
    for s in GetMyStockList():
        if s["StockCode"] == stock_code:
            qty = s["StockSellAbleAmt"] or s["StockAmt"]
            if qty > 0:
                return MakeSellMarketOrder(stock_code, int(qty))
    print(f"[US.SellAllStock] 보유 수량 없음: {stock_code}")
    return {"OrderNum": 0, "IsuNm": ""}


def GetMarketOrderPrice(stock_code):
    """시장가 참고용 현재가."""
    return GetCurrentPrice(stock_code)
