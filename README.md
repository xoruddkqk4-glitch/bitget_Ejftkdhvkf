# 비트겟 떨사오팔 봇 (Bitget SOXL Futures Trading Bot)

미국 증시 레버리지 ETF(SOXL)의 주가 시세 및 기술 지표(LS증권 API)를 모니터링하여, 기준치 이상 하락 시 비트겟(Bitget) USDT 무기한선물 롱 포지션으로 분할 매수하고 목표 수익률 도달 시 익절하는 자동매매 봇입니다.

---

## 1. 주요 전략 및 동작 방식

- **시세 및 기술 지표 수집**: LS증권 API (`LS_Common_new.py`, `LS_API_Helper_US_new.py`)
- **선물 주문 실행**: CCXT 기반 Bitget 선물 인프라 (`SOXL/USDT:USDT`, Cross Margin)
- **떨사 (분할 매수)**: 당일 시가 대비 일정 비율(`BUY_DROP_PCT = 1.5%`) 하락 시 매수 실행
- **낙폭 과대 증액 매수**: 최근 40일 고점 대비 40% 이상 하락 시 투자금 1.5배 증액 매수
- **오팔 (2단계 추세 추종 매도)**: 
  - **1단계 (감시)**: 각 차수별 매수가 대비 1차 목표 수익률(`SELL_PROFIT_PCT = 2.5%`) 도달 시 즉시 매도하지 않고 **2단계(추세추종)로 승격**, 최소 이익 보존선(`target_price`) 설정
  - **2단계 시나리오 ① (추세 꺾임 익절 - 초과 수익)**: 비트겟 선물 `TREND_TIMEFRAME`봉 직전 완성봉 종가가 `TREND_MA_PERIOD`MA 하향 이탈 시 꼭대기 추세 꺾임 익절 매도
  - **2단계 시나리오 ② (보존선 방어 - 최소 이익 안전망)**: 1차 목표가 밑으로 반락 시(`현재가 <= 보존선 × PRESERVATION_LINE_MULTIPLIER`) 즉시 방어 청산하여 +2.5% 수준 확정 이익 보존
  - **타임컷**: 목표가 미도달 시 40일(일반)/50일(낙폭과대) 경과 시 강제 청산
- **단일 5분 크론탭 통합 (`1-59/5 * * * *`)**: 매 5분마다 상시 매도를 감시하되, 미국 장 마감 10분 전인 **평일 19시 51분(UTC) 턴에만 일일 매수(떨사)를 1회 실행**하도록 통합 (매도 정기 브리핑은 2단계 추세추종 차수가 1개 이상 존재할 때만 매시 51분에 발송)
- **매매 기록 관리**: 엑셀 시트(`비트겟떨사오팔_매매기록.xlsx`) 및 상태 파일(`bitget_round_state.json`)에 실체결가, 수량, 차수별 단계(stage) 동적 관리 및 텔레그램 실시간 알림

### 2단계 추세추종 설정 변수

| 변수명 | 기본값 | 설명 |
|---|---|---|
| `TREND_TIMEFRAME` | `"4h"` | 추세 판정 캔들 시간프레임 (`"1h"`, `"4h"`, `"1d"` 등) |
| `TREND_MA_PERIOD` | `5` | 이동평균선 기간 (5 → 5MA, 10 → 10MA 등) |
| `TREND_OHLCV_LIMIT` | `10` | 추세 판정용 OHLCV 조회 캔들 수 |
| `PRESERVATION_LINE_MULTIPLIER` | `0.998` | 보존선 방어 매도 계수 |

---

## 2. 안전 장치 (Safety & Exception Handling)

- **당일 중복 매수 방지 체계 (Same-day Buy Lock)**:
  - 5분 크론탭 실행 환경에서 급락 장 발생 시 5분 간격으로 다음 차수가 연속 매수되는 사고를 원천 방지하기 위해, KST 당일 이미 매수된 이력이 있을 경우 추가 매수를 즉시 차단(Skip)합니다.
- **자가 복원 상태 관리자 (Self-Healing State Manager)**:
  - `bitget_round_state.json` 파일이 손상되거나 비어 있어도 기존 엑셀 `매매기록`의 미매도 차수들을 자동으로 역추적하여 `stage=1` 상태로 복원 후 실시간 감시를 재개합니다.
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

## [2026-10-20 09:48] 업데이트 이력 (Commit ID: 36d6abf)
- **수정 내용**:
  - **차수별 2단계 추세 추종 매도 전략 도입**:
    - 1차 목표가(+2.5%) 도달 시 2단계(추세추종)로 승격 및 이익 보존선 설정
    - 비트겟 선물 4시간봉 5MA 지표 판정 함수(`get_4h_ma5_status`) 추가 (직전 완성봉 기준)
    - 시나리오 ①(4h 5MA 이탈 꼭대기 익절) 및 시나리오 ②(보존선 * 0.998 반락 방어 매도) 구현
    - 차수별 독립 상태 관리 영속화 (`bitget_round_state.json`) 및 엑셀 기반 자가 복원(`sync_round_state_from_excel`) 연동
  - **단일 5분 크론탭(`1-59/5 * * * *`) 통합 & 당일 중복 매수 방어**:
    - `매수체크()`에 KST 당일 매수 이력 존재 시 중복 매수 차단 로직 추가
    - `매수매도프로세스()`에 평일 19:51(UTC, 미국 장 마감 10분 전) 턴에만 일일 매수 자동 실행 분기 구현
    - 매도 미발생 브리핑을 매시 51분 정기 턴에만 발송하여 텔레그램 알림 피로도 최적화
  - `__main__` 가드 추가로 모듈 임포트 안정성 확보
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 정적 구문 검사 통과
  - 차수별 1단계 등록 -> 2단계 승격 -> 매도 시 상태 제거 단위 테스트 통과 (ALL PASS)

## [2026-10-20 11:29] 업데이트 이력 (Commit ID: a6b28b6)
- **수정 내용**:
  - **매도 정기 브리핑 텔레그램 알림 조건 최적화**:
    - 매시 51분 정기 턴 도달 시 2단계(추세추종) 종목이 0개(`stage2_cnt == 0`)인 경우 1시간 단위 텔레그램 정기 브리핑 발송 생략
    - 2단계 추세추종 종목이 1개 이상 존재할 때만 1시간 단위 정기 브리핑 메시지(`⏸️ SOXL 매도 현황 브리핑 (정기)`) 발송
    - 불필요한 주기적 알림 피로도를 줄이고 콘솔 로그(`print`)로만 정상 기록 유지
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 정상 통과
  - EC2 실서버 파일 배포 및 크론탭 환경 검증 완료

## [2026-10-20 16:08] 업데이트 이력 (Commit ID: f3f6b6b)
- **수정 내용**:
  - **2단계 추세추종 상승 추세 기준 하드코딩 → 변수 추출 리팩터링**:
    - 4시간봉(`"4h"`) → `TREND_TIMEFRAME` 변수로 추출
    - 5MA 이동평균 기간(`5`) → `TREND_MA_PERIOD` 변수로 추출
    - OHLCV 조회 캔들 수(`10`) → `TREND_OHLCV_LIMIT` 변수로 추출
    - 최소 필요 캔들 수(`7`) → `TREND_MA_PERIOD + 2` 자동 계산으로 변경
    - 보존선 방어 계수(`0.998`) → `PRESERVATION_LINE_MULTIPLIER` 변수로 추출
    - 함수명 `get_4h_ma5_status` → `get_trend_ma_status`로 범용화
    - 텔레그램 메시지 및 콘솔 로그 내 하드코딩 문자열 모두 변수 참조로 변경
  - 기본값은 기존과 동일하여 동작 변경 없음
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 정상 통과

## [2026-10-21 07:48] 업데이트 이력 (Commit ID: d7999c1)
- **수정 내용**:
  - **2단계 승격 알림 5분 주기 무한 반복 발송 버그 수정**:
    - `sync_round_state_from_excel()`에서 엑셀 매매기록 전체 행 순회 시 과거 매도 완료된 동일 차수 행에 의해 상태(`rounds`)가 삭제(`pop`)된 후 `stage: 1`로 강제 리셋되던 결함 해결
    - 엑셀에서 현재 미매도 상태인 행(`매도가.isna()`)만 필터링하여 활성 차수 집합을 구성하도록 변경
    - 이미 2단계로 승격된 차수의 상태(`stage`, `preservation_line`, `stage2_entered_at`)는 동기화 시 덮어쓰지 않고 영구 보존
    - 전량 매도 완료되어 활성 차수 목록에 없는 과거 차수만 `rounds` 상태에서 안전하게 정리되도록 정류화
    - `매도체크()` 내 2단계 승격 즉시 메모리 상의 `rounds_dict`도 동기화하여 중복 감시 방지
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 통과

## [2026-10-21 17:14] 업데이트 이력 (Commit ID: a18b096)
- **발생 문제**:
  - SOXL 6차 보존선 방어 매도 오작동 발생: 한국 시간 16:25(KST) 크론탭 실행 시 LS증권 해외주식 전산이 익일 정산을 위해 현재가를 전일 종가(97.55)로 리셋하면서, 봇이 급락으로 오인하여 보존선 방어 청산을 실행함 (실제 Bitget 선물 시세는 104.72로 정상)
  - 결과적으로 메시지 내 '매도 사유(보존선 102.62 × 0.998 하향 터치)'와 '매도가(104.72)'가 불일치하는 모순 발생
- **수정 내용**:
  - **매도 감시 현재가 소스를 LS증권 → Bitget 실시간 선물 시세로 변경**:
    - `매도체크()` 내 `LsiUS.GetCurrentPrice(stock_code)` → `get_current_futures_price(exchange, symbol)` 교체
    - Bitget 선물(`SOXL/USDT:USDT`)은 24/7 연속 실시간 시세를 제공하므로 한국 장외 시간대 LS증권 전산 리셋 왜곡에 의한 오작동 원천 차단
    - 매수 체크(`매수체크()`)는 미국 정규장 시가/낙폭 조건 판정 정확성을 위해 기존 LS증권 하이브리드 방식 그대로 유지
- **아키텍처 변경 요약**:
  - `[매수]`: LS증권 시가(open) + LS증권 현재가 → Bitget 주문 집행 (유지)
  - `[매도]`: ~~LS증권 현재가~~ → **Bitget 실시간 선물 시세** → Bitget 청산 (변경)
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 통과

## [2026-10-22 08:42] 업데이트 이력 (Commit ID: 657d302)
- **발생 문제 / 요청 사항**:
  - 매일 1회 일일 매수 점검 시점(미국 장 마감 10분 전, 평일 19:51 UTC)에 기존에는 매도 상태 메시지와 매수 점검 메시지가 순차적으로 발송되었으나, 최근 5분 크론탭 알림 최적화 로직(2단계 추세추종 차수가 0개이면 51분 정기 브리핑 생략)으로 인해 보유 차수가 모두 1단계인 경우 매도 메시지가 생략되고 매수 메시지만 단독 발송되던 현상 발생
- **수정 내용**:
  - **일일 매수 턴 매도 점검 메시지 발송 보장**:
    - `매도체크()` 함수에 `is_daily_buy_turn=False` 매개변수 추가
    - 평일 19:51 UTC(또는 `FORCE_BUY_CHECK=1`) 일일 매수 턴에는 2단계 추세추종 차수 유무와 무관하게, 매도 미발생 시 `⏸️ SOXL 매도 없음` 안내 메시지 텔레그램 발송 보장
    - 매수 메시지와의 시세 표시 일관성을 위해 가능 시 LS 현재가를 우선 조회하여 표시(`disp_price`)
    - 일반 5분 크론탭 실행 시에는 기존의 알림 피로도 최적화(매도 발생 시에만 즉시 발송, 매시 51분 정기 브리핑은 2단계 차수 보유 시 발송) 로직 유지
  - **매수/매도 메시지 순차 전송 딜레이 적용**:
    - `매수매도프로세스()`에서 일일 매수 턴 매도 점검 후 텔레그램 수신 순서 보장을 위해 2초 대기(`time.sleep(2)`) 후 `매수체크()` 수행
  - **콘솔 로그 이모지 치환**:
    - 윈도우 환경(cp949) 및 터미널 출력 시 `UnicodeEncodeError` 방지를 위해 print 로그의 이모지 안전 텍스트화
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 정상 통과
  - 일일 매수 점검 턴 모의 테스트: `⏸️ SOXL 매도 없음` 메시지 발송 및 매수 점검 연계 동작 정상 검증 완료

## [2026-10-22 21:40] 업데이트 이력 (Commit ID: f4faeb1)
- **요청 사항**:
  - 떨사오팔 봇의 매수 및 매도 체결 내역을 통합 대시보드 및 구글 스프레드시트(Google Sheets)에 실시간으로 자동 기록할 수 있도록 웹훅 로거(`SheetsLogger`) 연동
- **수정 내용**:
  - **구글 스프레드시트 로거 연동 (`sheets_logger.py`)**:
    - GAS(Google Apps Script) 웹훅 엔드포인트와 통신하는 비동기/안전 전송 모듈(`sheets_logger.py`) 추가
    - `bitget_Ejftkdhvkf.py` 상단에 `SheetsLogger` 인스턴스 초기화 (`webhook_url`, `auth_token="trading_secret_2026"`, `strategy_name="떨사오팔"`, `exchange="Bitget"`)
  - **매도 체결 로깅 추가**:
    - `매도체크()` 내 매도 주문 성공 및 텔레그램 메시지 발송 직후 `sheet_logger.log_trade(action="SELL", ...)` 호출
    - 체결 차수, 매도가, 수량, 수수료(USDT), 실현 손익(`pnl`), 수익률(`pnl_pct`), 매도 사유(조정률 포함) 전송
  - **매수 체결 로깅 추가**:
    - `매수체크()` 내 분할 매수 주문 성공 및 텔레그램 메시지 발송 직후 `sheet_logger.log_trade(action="BUY", ...)` 호출
    - 체결 차수, 매수가, 수량, 낙폭률 메모 전송
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 통과
  - `sheets_logger` 모듈 임포트 및 인스턴스 생성 검증 완료
  - EC2 실서버 파일 배포 및 원격 구문 검증 완료

## [2026-10-22 23:44] 업데이트 이력 (Commit ID: 3e6795d)
- **요청 사항**:
  - 구글 스프레드시트 체결 기록 전송 시 기록 시각(타임스탬프)을 한국 표준시(KST, UTC+9) 기준으로 정확하게 기록하도록 개선
- **수정 내용**:
  - **구글 스프레드시트 웹훅 로거 타임스탬프 KST 적용 (`sheets_logger.py`)**:
    - `datetime`, `timezone`, `timedelta`를 활용하여 KST(UTC+9) 타임존 객체 정의 (`KST = timezone(timedelta(hours=9))`)
    - `log_trade()` 메서드에서 외부 타임스탬프 미지정 시 기본 생성되는 `now_str`을 `datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")`로 수정하여 서버 환경(UTC 등)과 관계없이 한국 시간 기준으로 구글 시트에 일관되게 기록되도록 보장
- **검증 결과**:
  - `python -m py_compile sheets_logger.py` 및 `bitget_Ejftkdhvkf.py` 구문 검사 정상 통과
  - EC2 실서버 파일 배포 및 원격 구문 검증 완료

## [2026-10-23 08:42] 업데이트 이력 (Commit ID: f3d7bc3)
- **요청 사항**:
  - 24시간 봇 구동 중 보유금(USDT) 부족으로 매수가 보류되었거나 수동으로 추가 매수를 진행하고자 할 때, 미국 장 마감 10분 전(평일 19:51 UTC)이라는 특정 매수 스케줄과 당일 1회 매수 제한으로 인해 매수 진행이 불가능했던 문제 해결 요청
  - 매수 기준 확인 후 매수만 수동으로 1회 진행하고, 기존 24시간 봇이 매도를 계속 감시하도록 연동하는 코드 추가 요청
- **수정 내용**:
  - **독립형 수동 매수 전용 스크립트 신규 구현 (`manual_buy.py`)**:
    - 기존 24시간 매도 감시 봇(`bitget_Ejftkdhvkf.py`) 프로세스를 중단하거나 재기동할 필요 없이 언제든 1회성 매수를 안전하게 수행
    - **동일 매수 알고리즘 적용**: 당일 시가 대비 -1.5% 하락 여부(`BUY_DROP_PCT = 0.015`), 40일 고점 대비 -40% 하락 시 1.5배 증액(`HIGH_DROP_THRESHOLD = 0.40`), `비트겟떨사오팔_매매기록.xlsx` 내 투자금 시트 기준 반영
    - **비트겟 선물 주문 및 실체결 회수**: SOXL 선물 롱 포지션 오픈, 체결 대기 및 실체결가/실수량 조회
    - **엑셀 및 상태 JSON 자동 동기화**: `비트겟떨사오팔_매매기록.xlsx`에 신규 차수 기록 및 `bitget_round_state.json`에 1단계(목표가 +2.5%) 등록
    - **24시간 매도 감시 자동 연동**: 24시간 봇이 5분마다 엑셀을 읽어 미매도 차수를 자동 복원(`sync_round_state_from_excel`)하므로, 수동 매수 건도 별도 조치 없이 자동으로 24시간 매도 감시(1차 목표가 및 2단계 추세추종) 대상에 편입
    - **알림 및 로깅**: 텔레그램 `[수동 매수 실행]` 알림 발송 및 구글 스프레드시트(`SheetsLogger`) 웹훅 전송
  - **다양한 CLI 실행 모드 지원**:
    - `python manual_buy.py`: 매수 기준 충족 시에만 매수 실행 (기본 모드)
    - `python manual_buy.py --check-only`: 주문 없이 현재 시세, 매수 기준가, 가용 USDT 잔고, 다음 예정 차수만 확인
    - `python manual_buy.py --force`: 매수 조건 미달이어도 지금 즉시 1차수 강제 수동 매수
- **검증 결과**:
  - `python -m py_compile manual_buy.py` 구문 검사 정상 통과
  - `python manual_buy.py --help` CLI 인자 파싱 및 모듈 로딩 정상 검증 완료

## [2026-10-23 21:26] 업데이트 이력 (Commit ID: 9f57680)
- **요청 사항**:
  - 구글 시트 웹훅 전송 및 로깅 시 데이터 처리 속도 향상과 함께, 봇 내부 상태 동기화 및 원장 유지관리 기능 추가 요청
- **수정 내용**:
  - **HTTP 커넥션 풀링을 통한 데이터 처리 속도 대폭 향상 (`sheets_logger.py`)**:
    - `requests.Session()` 인스턴스를 초기화하여 매 요청마다 발생하던 TCP 3-way handshake 및 SSL/TLS 재협상 오버헤드를 제거
    - `log_trade()` 및 `log_snapshot()`에서 세션 커넥션을 재사용(`http_client = self._session or requests`)하여 데이터 전송 응답 속도 및 네트워크 효율 극대화
  - **봇 내부 상태 및 원장 관리 원격 제어 메서드 확장 (`sheets_logger.py`)**:
    - `log_bot_state()`: 쿨링다운, 공백 예약, 잔고부족 차수 등 봇의 실시간 내부 상태를 대시보드 및 구글 시트로 비동기 전송 (타임아웃 2초 및 예외 격리로 메인 매매 루프 완벽 보장)
    - `trigger_monthly_archive()`: 원장 데이터 누적으로 인한 시트 속도 저하를 방지하기 위해 월별 탭 자동 아카이브를 원격 트리거하는 기능 추가
    - `trigger_yearly_sharding()`: 연간 원장 파일을 별도 분리(샤딩)하여 원장 성능을 최적화하는 원격 실행 기능 추가
    - `diagnose_capacity()`: 원장 용량 및 건강도 진단 원격 조회 기능 구현
- **검증 결과**:
  - `python -m py_compile sheets_logger.py` 구문 검사 정상 통과
  - EC2 실서버 파일 배포 및 원격 구문 검증 완료

## [2026-10-24 11:05] 업데이트 이력 (Commit ID: 57c6721)
- **요청 사항**:
  - 구글 스프레드시트 연동 보안 강화 및 소스 코드 내 웹훅 URL과 인증 토큰 하드코딩 제거 요청
  - 서버 공통 환경설정 파일(`/var/autobot/gsheet.env`) 기반 자동 설정 주입 지원 요청
- **수정 내용**:
  - **구글 시트 로거 보안 강화 및 환경설정 파일 자동 로드 (`sheets_logger.py`)**:
    - `DEFAULT_ENV_FILE = "/var/autobot/gsheet.env"` (또는 환경변수 `GSHEET_ENV_FILE`)에서 `GSHEET_*` 설정을 안전하게 파싱하여 `os.environ`에 주입하는 `_load_env_file()` 구현
    - 기본 인증 토큰 하드코딩(`trading_secret_2026`)을 완전히 제거하고 `GSHEET_AUTH_TOKEN` 환경변수로만 주입받도록 보안 강화
    - 배포 ID만 환경변수에 입력된 경우에도 표준 웹앱 URL(`https://script.google.com/macros/s/{id}/exec`)로 자동 정규화하는 보정 로직 추가
    - 토큰 누락 시 명확한 경고 출력 및 예외 격리 유지
  - **떨사오팔 메인 봇 하드코딩 제거 (`bitget_Ejftkdhvkf.py`)**:
    - `SheetsLogger` 인스턴스 생성 시 소스에 하드코딩되어 있던 `webhook_url` 및 `auth_token` 매개변수 제거
    - 환경변수 및 공통 설정 파일(`gsheet.env`)을 통해 자동으로 자격증명을 주입받도록 설정 일원화
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py sheets_logger.py` 로컬 구문 컴파일 통과
  - EC2 실서버 파일 전송 (`scp`) 및 원격 구문 검증 완료
  - EC2 상에서 `sheets_logger`의 `gsheet.env` 자동 로드 및 Webhook 엔드포인트 연동 테스트 정상 완료

## [2026-10-24 13:38] 업데이트 이력 (Commit ID: bede1a8)
- **수정 내용**:
  - **`sheets_logger.py` 백그라운드 스레드 큐 및 로컬 Outbox 장애 복구 아키텍처 구축 ([`sheets_logger.py`](file:///c:/Users/user/Desktop/trading_buddy/06비트겟자동매매-4기/week_02-bitget_Ejftkdhvkf/떨사오팔봇/sheets_logger.py))**:
    - **1) 논블로킹 백그라운드 스레드 큐 (`queue.Queue`, `threading.Thread`) 도입**:
      - `async_mode=True` 기본 활성화: `log_trade`, `log_trades_batch`, `log_snapshot`, `log_bot_state` 호출 시 큐에 즉시 삽입하고 반환하여 매매 루프의 지연시간을 0ms 수준으로 최소화
      - 첫 전송 요청 시 백그라운드 워커 스레드 자동 가동 (Lazy Start)
    - **2) 로컬 Outbox 장애 복구 시스템 (`sheets_outbox_*.jsonl`) 구현**:
      - Google 서버 지연, 네트워크 단절 등으로 전송 실패 시 체결 데이터를 전략별 로컬 Outbox 파일에 안전하게 보관
      - 백그라운드 워커에서 60초 주기 재전송 시도 및 네트워크 다운 시 자동 백오프(`NETWORK_DOWN_BACKOFF_SEC=30`) 적용
      - 중복 전송 방지: 서버 측 `trade_id` 기반 중복 검증 연동 (서버의 `duplicate` 응답 수신 시에도 성공 처리)
      - 최대 재시도 초과 시 `.dead.jsonl` 격리 보관
    - **3) 봇 상태 최신 1건 압축(Coalescing) 전송**:
      - `log_bot_state`는 최신 상태 1건만 유지하도록 전송 버퍼 압축 처리
    - **4) 프로세스 정상 종료 시 안전 플러시 (`atexit`)**:
      - SIGTERM 등 종료 시 대기 중인 큐의 체결 데이터를 최대 5초간 안전하게 전송
    - **5) Webhook URL 포맷 자동 보정 및 KST 타임존 정합성 유지**:
      - Apps Script 배포 ID만 기재된 경우 정규 URL 자동 보정 및 스냅샷 일자 생성 시 KST 타임존 적용 유지
- **검증 결과**:
  - `python -m py_compile sheets_logger.py bitget_Ejftkdhvkf.py manual_buy.py` 정적 구문 컴파일 검증 완료 (오류 0건 통과)
  - EC2 실서버(`ec2:/var/autobot/bitget_Ejftkdhvkf/`) 파일 배포 및 원격 구문 검증 완료

## [2026-10-25 21:48] 업데이트 이력 (Commit ID: e280676)
- **요청 사항**:
  - 매도 미발생 안내 메시지(`⏸️ SOXL 매도 없음`)에 미매도 보유 차수별 매도 기준가 정보를 함께 표시하도록 개선
  - 매수 보류 메시지(`⏸️ SOXL 매수 보류`) 내 가격 정보(시가, 현재가, 기준가)를 실제 가격 수치 기준 내림차순(높은 가격 순)으로 정렬하여 표시하도록 개선
- **수정 내용**:
  - **매도 미발생 메시지 내 미매도 차수별 매도 기준가 표시 추가 (`bitget_Ejftkdhvkf.py`)**:
    - `매도체크()` 내 매도 미발생 시 미체결 행들을 차수 순으로 순회하여 차수별 매도 기준가(목표가, `매수가 × 1.025`), 매수가 및 2단계 추세추종 여부를 산출
    - 일일 매수 점검 턴(`is_daily_buy_turn`) 알림 및 정기 브리핑 메시지에 `🎯 차수별 매도 기준가` 목록 추가 출력
  - **매수 보류 메시지 가격 항목 내림차순 정렬 (`bitget_Ejftkdhvkf.py`)**:
    - `매수체크()` 조건 미충족 시 시가(LS), 현재가(LS), 기준가를 튜플 리스트로 묶고 가격 크기 기준 내림차순(`reverse=True`)으로 자동 정렬하여 메시지에 동적 반영
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 정상 통과
  - 매수 가격 내림차순 정렬 및 미매도 차수별 기준가 텍스트 빌드 단위 로직 검증 완료

## [2026-10-28 19:04] 업데이트 이력 (Commit ID: 8454102)
- **요청 사항**:
  - 매도 미발생 텔레그램 메시지(`⏸️ SOXL 매도 없음`)의 차수별 매도 기준가 목록에 보유하고 있는 각 차수별 강제 매도일(`강제매도일 YYYY-MM-DD`) 정보 포함 요청
  - 일반 조건 매수 시 최대 보유 기간(`MAX_HOLD_DAYS_NORMAL`)을 기존 20일에서 40일로 변경
- **수정 내용**:
  - **일반 매수 최대 보유 기간 변경 (`bitget_Ejftkdhvkf.py`)**:
    - `MAX_HOLD_DAYS_NORMAL` 상수를 `20` → `40`으로 변경 (일반 매수 40일, 고점 대비 하락 조건 매수 50일)
  - **매도 미발생 메시지 내 차수별 강제 매도일 표시 추가 (`bitget_Ejftkdhvkf.py`)**:
    - `datetime` 모듈에서 `timedelta` 추가 임포트
    - `매도체크()` 내 미매도 보유 차수별 목록(`rounds_detail_lines`) 구성 시, 각 차수의 매수일(`날짜`)과 `고점대비하락조건` 여부에 따른 최대 보유일수(일반 40일 / 고점대비하락 50일)를 합산하여 `강제매도일 YYYY-MM-DD`를 함께 출력하도록 개선
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 구문 검사 정상 통과

## [2026-10-29 23:06] 업데이트 이력 (Commit ID: 337fe42)
- **요청 사항**:
  - 대시보드 및 구글 시트와 떨사오팔 봇의 실시간 상태(2단계 추세추종 지표 및 보유 차수 상태)를 동기화하도록 `STATE_SYNC` 연동 추가
- **수정 내용**:
  - **4시간봉 5MA 추세 지표 상시 조회 (`bitget_Ejftkdhvkf.py`)**:
    - `매도체크()`에서 2단계 진입 차수 보유 여부(`has_stage2`)와 무관하게 `get_trend_ma_status(exchange, symbol)`를 매 주기 조회하여 대시보드에 실시간 추세 데이터 제공
  - **구글 시트 봇 상태 동기화 연동 (`bitget_Ejftkdhvkf.py`)**:
    - `매도체크()` 종료 시점에 `sheet_logger.log_bot_state()`를 호출하여 심볼, 종목코드, `trend_riding` 상태(`active`, `stage2_active`, `prev_close_4h`, `ma5_4h`, `timeframe`, `ma_period`) 및 최신 보유 차수(`rounds`) 정보를 비동기 전송
    - 예외 격리(`try-except`)를 적용하여 상태 동기화 오류가 매매 루프에 영향을 주지 않도록 방어
- **검증 결과**:
  - `python -m py_compile bitget_Ejftkdhvkf.py` 정적 구문 검사 통과




