import time, win32con, win32api, win32gui, who_will_present_science as ps

schedule = [
    ["자료구조B", "통합과학", "통합사회", "수학", "자료구조A", "체육", "국어", "종례"],
    ["컴퓨터시스템일반", "미술", "미술", "영어(공)", "수학", "진로(CA)", "진로(CA)", "종례"],
    ["통합사회", "컴퓨터시스템일반", "통합과학", "미술", "체육", "국어", "종례"],
    ["진로", "자료구조A", "자료구조B", "자료구조B", "통합과학", "영어(공)", "HR", "종례"],
    ["자료구조A", "자료구조A", "수학", "통합사회", "컴퓨터시스템일반", "국어", "영어(전)", "종례"]
]

schedule_link = {
    "국어": "(링크)",
    "수학": "(링크)",
    "통합사회": "(링크)",
    "통합과학": "(링크)",
    "영어(공)": "(링크)",
    "영어(전)": "(링크)",
    "컴퓨터시스템일반": "(링크)",
    "자료구조A": "(링크)",
    "자료구조B": "(링크)",
    "미술": "(링크)",
    "체육": "(링크)",
    "진로": "(링크)",
    "조회": "(링크)",
    "종례": "(링크)",
    "진로(CA)": "(링크)",
    "동아리(CA)": "(링크)",
    "HR": "(링크)"
}

# 채팅방 이름
kakaoRoomName = ["고1 쌤없 반톡", "1-5반 (선생님)"]
trg = False
dow = datetime.datetime.now().weekday()

# 시간표 받아오기
def getTodaySchedule(inp_period, type="default"):
    try:
        if type == "default":
            return schedule[dow][inp_period]
        elif type == "link":
            return schedule_link.get(schedule[dow][inp_period])
    except (IndexError, TypeError):
        print("👍 오늘의 모든 교시를 마쳤습니다. 프로그램을 종료합니다 👍", end="")
        exit()

# 카카오톡 제어
def kakaoSendText(roomName, text):
    hwnd = win32gui.FindWindowEx(win32gui.FindWindow(None, roomName), None, "RichEdit50W", None)
    win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.1)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

# 과학 발표자
trees = [ps.result.get("Tree 1"), ps.result.get("Tree 2"), ps.result.get("Tree 3")]
presenter = list()

try:
    for ori in ps.result.get("Origin"):
        presenter.append(ps.getInfoByNum(ori).get("name"))

    for tree in trees:
        for branch in tree.items():
            for leaf in branch[1]:
                presenter.append(ps.getInfoByNum(leaf).get("name"))
finally:
    result = ""
    for pst in enumerate(list(set(presenter))):
        result += pst[1] if pst[0]+1 == len(list(set(presenter))) else pst[1] + ", "
        
# 작동
while True:
    now = datetime.datetime.now()
    start_hour, term_hour = 8, 1
    end_hour = 14 if dow == 2 else 15
    ok_hour = [i if i < 12 else i+1 for i in range(start_hour, end_hour, term_hour)]
    period = now.hour-7 if now.hour < 13 else now.hour-8
    send_min = 51 if now.hour < 13 else 41

    if dow in range(0,5): # 월화수목금
        now_schedule = getTodaySchedule(period-1)
        now_schedule_link = getTodaySchedule(period-1, "link")
        sciencePresenterMessage = f"☆ 통합과학: [{result}]는 발표를 준비해주세요 ☆"

        if now_schedule == getTodaySchedule(period-2):
            message = f'📢 [Bot] 이번교시는 연강입니다.\n' \
                      f'혹시 튕기거나 나갔다면 아래의 링크를 통해 다시 접속해주세요.\n' \
                      f'{now_schedule_link}'
        else:
            message = f'📢 [Bot] 현재 시간 {now.hour}시 {now.minute}분을 지나가고 있습니다.\n' \
                      f'{period}교시는 "{now_schedule}" 시간입니다.\n' \
                      f'{now_schedule_link}'

        for room in kakaoRoomName:
            if now.hour in ok_hour and now.minute == send_min:
                trg = False
                kakaoSendText(room, message)
                print("{}시 {}분 {}초, '{}'방에\n{:=^185}\n전송했습니다\n".format(now.hour, now.minute, now.second, room, f'\n{message}\n'))

                if now_schedule == "통합과학":
                    kakaoSendText(room, sciencePresenterMessage)
                    print(f"+ 통합과학 발표 대상자도 전송했습니다({result})\n")
            else:
                if not trg:
                    print(f'{send_min}분이 되면 "{getTodaySchedule(period)}" 시간 공지를 전송합니다')
                    trg = True

        if now_schedule == "종례" and not trg:
            exit()
    else:
        print("오늘은 주말입니다. 프로그램을 실행할 수 없습니다", end="")
        exit()
