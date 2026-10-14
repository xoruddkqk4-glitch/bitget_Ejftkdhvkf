# 비트겟 떨사오팔 봇 (Bitget SOXL Futures Trading Bot)

미국 증시 레버리지 ETF(SOXL)의 주가 시세 및 기술 지표(LS증권 API)를 모니터링하여, 기준치 이상 하락 시 비트겟(Bitget) USDT 무기한선물 롱 포지션으로 분할 매수하고 목표 수익률 도달 시 익절하는 자동매매 봇입니다.

---

## 1. 주요 전략 및 동작 방식

- **시세 및 기술 지표 수집**: LS증권 API (`LS_Common_new.py`, `LS_API_Helper_US_new.py`)
- **선물 주문 실행**: CCXT 기반 Bitget 선물 인프라 (`SOXL/USDT:USDT`, Cross Margin)
- **떨사 (분할 매수)**: 당일 시가 대비 일정 비율(`BUY_DROP_PCT = 1.5%`) 하락 시 매수 실행
- **낙폭 과대 증액 매수**: 최근 40일 고점 대비 40% 이상 하락 시 투자금 1.5배 증액 매수
- **오팔 (익절 매도)**: 각 차수별 매수가 대비 목표 수익률(`SELL_PROFIT_PCT = 2.5%`) 도달 시 자동 익절
- **매매 기록 관리**: 엑셀 시트(`비트겟떨사오팔_매매기록.xlsx`)에 실체결가, 수량, 투자금 동적 관리 및 텔레그램 실시간 알림

---

## 2. 안전 장치 (Safety & Exception Handling)

- **보유금(USDT) 부족 방어 체계 (방안 A: Skip & Alert)**:
  - **1단계 (사전 점검)**: 주문 전 비트겟 선물 계좌의 가용 USDT 잔고(`get_available_usdt`)를 조회하여 회차별 필요 투자금보다 부족할 경우 주문을 넣지 않고 안전하게 매수를 보류(Skip)합니다.
  - **2단계 (사후 방어)**: 거래소 주문 전송 시 잔고 부족으로 발생하는 `ccxt.InsufficientFunds` 및 증거금 관련 에러를 포획하여, 프로세스가 비정상 종료(Crash)되지 않도록 보호하고 텔레그램으로 상세 알림을 발송합니다.
- **미체결 지정가 취소**: 지정가 주문 후 일정 시간 체결되지 않을 시 주문 자동 취소

---

## 3. 업데이트 이력

## [2026-10-15 08:15] 업데이트 이력 (Commit ID: 1d517c8)
- **수정 내용**:
  - 보유금 부족 상황 처리 방안(방안 A: Skip & Alert) 구현
  - `get_available_usdt(exchange)` 헬퍼 함수 추가를 통한 가용 USDT 잔고 사전 점검
  - `매수체크()` 내 사전 잔고 부족 시 안전 매수 보류 및 텔레그램 알림 로직 추가
  - `place_open_long()` 호출 부 `ccxt.InsufficientFunds` 및 비트겟 증거금 에러 핸들링 추가
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 통과
  - CCXT 선물 잔고 응답 파싱 단위 테스트 통과
