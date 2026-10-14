# -*- coding: utf-8 -*-
# =========================================================
#  telegram_alert.py  —  텔레그램 알림 모듈 (수강생 배포용)
# =========================================================
#  ★ 아래 TOKEN 과 CHAT_ID 두 줄만 본인 값으로 채우면 됩니다.
#    비워 두면 봇은 정상 동작하지만 텔레그램 알림은 나가지 않습니다.
#
#  ─────────────────────────────────────────────────────────
#  [ TOKEN 받는 법 ]
#    1) 텔레그램에서 @BotFather 를 검색해 대화를 시작합니다.
#    2) /newbot 입력 → 봇 이름과 아이디(끝이 bot 이어야 함)를 정합니다.
#    3) BotFather 가 주는 문자열이 TOKEN 입니다.
#       형식 예시)  "123456789:AAF0abcdEFGHijklMNOPqrstUVWXyz123456"
#                   └ 숫자 ┘ └────────── 영문+숫자 ──────────┘
#
#  [ CHAT_ID 받는 법 ]
#    1) 방금 만든 봇과의 대화방을 열고 아무 메시지나 한 번 보냅니다.
#    2) 브라우저 주소창에 아래를 입력합니다. <TOKEN> 자리에 본인 토큰을 넣으세요.
#         https://api.telegram.org/bot<TOKEN>/getUpdates
#    3) 화면에 나온 결과에서  "chat":{"id":123456789,...  부분의
#       숫자가 CHAT_ID 입니다.
#       형식 예시)  123456789        (따옴표 없이 숫자만)
#       그룹방이면 앞에 -100 이 붙은 음수일 수 있습니다.
#       형식 예시)  -1001234567890
#
#  [ 설치 ]
#    pip3 install python-telegram-bot
#
#  ※ 이 파일에는 절대 남의 토큰을 넣지 마세요. 토큰이 유출되면
#    누구나 그 봇으로 메시지를 보낼 수 있습니다.
# =========================================================

import telegram
import asyncio

# ─────────────────────────────────────────────────────────
# ★★★ 여기 두 줄을 본인 값으로 채우세요 ★★★
# ─────────────────────────────────────────────────────────
TOKEN = "8394717466:AAGq3SGPDQrtAyfI3juqG7QrqLEDDD57Lbo"        # 예시) "123456789:AAF0abcdEFGHijklMNOPqrstUVWXyz123456"
CHAT_ID = "6847524018"      # 예시) 123456789
# ─────────────────────────────────────────────────────────


# TOKEN/CHAT_ID 가 비어 있어도 import 단계에서 죽지 않도록 처리한다.
# (봇 로직은 그대로 돌고, 알림만 건너뛴다)
bot = None
if TOKEN and str(CHAT_ID).strip():
    try:
        bot = telegram.Bot(token=TOKEN)
    except Exception as ex:
        print(f"[telegram_alert] 봇 생성 실패: {ex}")
else:
    print("[telegram_alert] TOKEN / CHAT_ID 가 비어 있습니다. "
          "telegram_alert.py 를 열어 본인 값을 넣어주세요. (알림 전송 안 됨)")


async def send_telegram_alert(msg):
    await bot.send_message(chat_id=CHAT_ID, text=msg)


def SendMessage(msg):
    # 설정이 비어 있으면 콘솔에만 남기고 조용히 넘어간다.
    if bot is None:
        print(f"[telegram_alert 미설정] {msg}")
        return
    try:
        loop = asyncio.get_event_loop()   # 기존 이벤트 루프 가져오기
        if loop.is_running():             # 이미 이벤트 루프가 실행 중인 경우
            loop.create_task(send_telegram_alert(msg))   # 비동기 작업을 추가
        else:
            loop.run_until_complete(send_telegram_alert(msg))  # 새 루프 실행
    except Exception as ex:
        print(ex)
