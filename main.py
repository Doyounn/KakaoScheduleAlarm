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
cnt = False

def getTime(inp):
    return {
        "요일": time.strftime("%a", time.localtime(time.time())),
        "시": int(time.strftime("%H", time.localtime(time.time()))),
        "분": int(time.strftime("%M", time.localtime(time.time()))),
        "초": int(time.strftime("%S", time.localtime(time.time())))
    }.get(inp, "DEFAULT")

def getNumDay(inp):
    return {
        "Mon": 0,
        "Tue": 1,
        "Wed": 2,
        "Thu": 3,
        "Fri": 4,
    }.get(inp, "DEFAULT")

def getTodaySchedule(inp_day, inp_col):
    try:
        return schedule[getNumDay(inp_day)][inp_col]
    except (IndexError, TypeError):
        print("👍 오늘의 모든 교시를 마쳤습니다. 프로그램을 종료합니다 👍", end="")
        exit()

def getTodayScheduleLink(subject):
    return schedule_link.get(subject)

def kakaoSendText(roomName, text):
    hwnd = win32gui.FindWindowEx(win32gui.FindWindow(None, roomName), None, "RichEdit50W", None)
    win32api.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
    SendReturn(hwnd)

def SendReturn(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(0.1)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

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
    presenter = list(set(presenter))
    result = ""

    for pst in enumerate(presenter):
        result += pst[1] if pst[0]+1 == len(presenter) else pst[1] + ", "

while True:
    day, hour, minute, second = getTime("요일"), getTime("시"), getTime("분"), getTime("초")
    col = hour - 7 if hour < 13 else hour - 8 # 교시
    ok_day = ["Mon", "Tue", "Wed", "Thu", "Fri"]
    start_hour, term_hour = 8, 1
    end_hour = 14 if day == "Wed" else 15 # 수요일은 6교시까지, 다른 날은 7교시까지
    ok_hour = [i if i < 12 else i+1 for i in range(start_hour, end_hour, term_hour)] # 1시간마다 전송 (12시 제외)
    send_min = 52 if hour < 13 else 42 # 점심 전까지는 52분에, 이후에는 42분에 안내
    now_schedule = getTodaySchedule(day, col-1)
    now_schedule_link = getTodayScheduleLink(now_schedule)
    sciencePresenterMessage = f"☆ 통합과학: [{result}]는 발표를 준비해주세요 ☆"

    if now_schedule == getTodaySchedule(day, col-2):
        message = f'📢 [Bot] 이번교시는 연강입니다.\n' \
                  f'혹시 튕기거나 나갔다면 아래의 링크를 통해 다시 접속해주세요.\n' \
                  f'{now_schedule_link}'
    else:
        message = f'📢 [Bot] 현재 시간 {hour}시 {minute}분을 지나가고 있습니다.\n' \
                  f'{col}교시는 "{now_schedule}" 시간입니다.\n' \
                  f'{now_schedule_link}'

    if day in ok_day:
        for room in kakaoRoomName:
            if hour in ok_hour and minute == send_min:
                cnt = False
                kakaoSendText(room, message)
                print("{0}시 {1}분 {2}초, '{3}'방에\n{4:=^185}\n전송했습니다\n".format(hour, minute, second, room, f'\n{message}\n'))

                if now_schedule == "통합과학":
                    kakaoSendText(room, sciencePresenterMessage)
                    print(f"+ 통합과학 발표 대상자도 전송했습니다({result})\n")
            else:
                if not cnt:
                    print(f'{send_min}분이 되면 "{getTodaySchedule(day, col)}" 시간 공지를 전송합니다')
                    cnt = True
    else:
        print("오늘은 주말입니다. 프로그램을 실행할 수 없습니다", end="")
        exit()

    if now_schedule == "종례" and not cnt:
        exit()

    time.sleep(60)
