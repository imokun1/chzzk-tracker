import os
import requests
import gspread
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

load_dotenv()

CLIENT_ID = os.getenv("CHZZK_CLIENT_ID")
CLIENT_SECRET = os.getenv("CHZZK_CLIENT_SECRET")
SHEET_NAME = "치지직_버튜버_활동기록"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def to_int(value):
    try:
        return int(value)
    except:
        return 0


def parse_datetime(value):
    if not value:
        return None

    value = str(value).strip()

    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
    except:
        pass

    try:
        return datetime.strptime(value[:19], "%Y-%m-%d %H:%M:%S")
    except:
        return None


def calculate_viewer_stats(snapshots, live_id):
    viewers_list = []

    for snapshot in snapshots:
        if str(snapshot.get("live_id")) == str(live_id):
            viewers = to_int(snapshot.get("current_viewers"))
            viewers_list.append(viewers)

    if not viewers_list:
        return 0, 0

    peak_viewers = max(viewers_list)
    average_viewers = round(sum(viewers_list) / len(viewers_list))

    return peak_viewers, average_viewers


def archive_old_snapshots(snapshots_ws, archive_ws, days=30):
    snapshots = snapshots_ws.get_all_values()

    if len(snapshots) <= 1:
        return 0

    rows = snapshots[1:]
    cutoff = datetime.now() - timedelta(days=days)

    rows_to_archive = []
    row_numbers_to_delete = []

    for i, row in enumerate(rows, start=2):
        if len(row) < 2:
            continue

        snapshot_time_raw = row[1]
        snapshot_dt = parse_datetime(snapshot_time_raw)

        if snapshot_dt and snapshot_dt < cutoff:
            rows_to_archive.append(row)
            row_numbers_to_delete.append(i)

    if not rows_to_archive:
        return 0

    archive_ws.append_rows(rows_to_archive)

    for row_number in reversed(row_numbers_to_delete):
        snapshots_ws.delete_rows(row_number)

    return len(rows_to_archive)


def get_channel_info(channel_id):
    url = f"https://api.chzzk.naver.com/service/v1/channels/{channel_id}"

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    response = requests.get(url, headers=headers, timeout=15)
    data = response.json()

    content = data.get("content", {})

    return {
        "channel_id": content.get("channelId"),
        "channel_name": content.get("channelName"),
        "follower_count": to_int(content.get("followerCount")),
    }


def get_header_map(worksheet):
    headers = worksheet.row_values(1)
    return {header: index + 1 for index, header in enumerate(headers)}


def update_creator_status(
    creators_ws,
    row_number,
    header_map,
    is_live,
    live_title,
    current_viewers,
    current_followers,
    last_live_time,
    total_live_count,
    avg_live_duration,
    checked_at
):
    updates = {
        "current_live_status": "TRUE" if is_live else "FALSE",
        "current_live_title": live_title,
        "current_viewers": current_viewers,
        "current_followers": current_followers,
        "last_live_time": last_live_time,
        "total_live_count": total_live_count,
        "avg_live_duration": avg_live_duration,
        "last_checked_at": checked_at,
    }

    for column_name, value in updates.items():
        col = header_map.get(column_name)
        if col:
            creators_ws.update_cell(row_number, col, value)


def calculate_creator_summary(sessions, channel_id):
    channel_sessions = []

    for session in sessions:
        if str(session.get("channel_id")) == str(channel_id):
            channel_sessions.append(session)

    total_live_count = len(channel_sessions)

    durations = []
    last_live_time = ""

    for session in channel_sessions:
        start_time = str(session.get("start_time")).strip()
        duration = to_int(session.get("duration_minutes"))

        if start_time and (not last_live_time or start_time > last_live_time):
            last_live_time = start_time

        if duration > 0:
            durations.append(duration)

    avg_live_duration = ""

    if durations:
        avg_live_duration = round(sum(durations) / len(durations))

    return total_live_count, avg_live_duration, last_live_time


creds = Credentials.from_service_account_file("google_credentials.json", scopes=SCOPES)
client = gspread.authorize(creds)

sheet = client.open(SHEET_NAME)

creators_ws = sheet.worksheet("creators")
snapshots_ws = sheet.worksheet("live_snapshots")
archive_ws = sheet.worksheet("live_snapshots_archive")
sessions_ws = sheet.worksheet("live_sessions")
followers_ws = sheet.worksheet("follower_snapshots")
logs_ws = sheet.worksheet("collection_logs")

run_time = now_str()

try:
    creators = creators_ws.get_all_records()
    creator_header_map = get_header_map(creators_ws)

    target_channel_ids = []
    creator_rows_by_channel_id = {}

    for index, creator in enumerate(creators, start=2):
        if str(creator.get("is_active")).upper() == "TRUE":
            channel_id = str(creator.get("channel_id")).strip()

            if channel_id:
                target_channel_ids.append(channel_id)
                creator_rows_by_channel_id[channel_id] = index

    print("추적 대상 채널 수:", len(target_channel_ids))

    # 1. 팔로워 수 기록
    follower_collected_count = 0
    latest_followers_by_channel_id = {}

    for channel_id in target_channel_ids:
        try:
            channel_info = get_channel_info(channel_id)

            if channel_info["channel_id"]:
                followers_ws.append_row([
                    run_time,
                    channel_info["channel_id"],
                    channel_info["channel_name"],
                    channel_info["follower_count"]
                ])

                latest_followers_by_channel_id[channel_id] = channel_info["follower_count"]

                follower_collected_count += 1
                print("팔로워 기록:", channel_info["channel_name"], channel_info["follower_count"])

        except Exception as follower_error:
            print("팔로워 기록 실패:", channel_id, follower_error)

    # 2. 현재 라이브 목록 조회
    url = "https://openapi.chzzk.naver.com/open/v1/lives?size=20"

    headers = {
        "Client-Id": CLIENT_ID,
        "Client-Secret": CLIENT_SECRET,
    }

    response = requests.get(url, headers=headers, timeout=15)
    print("치지직 API 상태 코드:", response.status_code)

    data = response.json()
    lives = data.get("content", {}).get("data", [])

    current_live_ids = set()
    live_by_channel_id = {}

    sessions = sessions_ws.get_all_records()

    collected_count = 0
    created_session_count = 0
    updated_session_count = 0
    ended_session_count = 0

    # 3. 현재 방송 중인 라이브 기록
    for live in lives:
        channel_id = live.get("channelId")

        if channel_id not in target_channel_ids:
            continue

        live_by_channel_id[channel_id] = live

        live_id = str(live.get("liveId"))
        current_live_ids.add(live_id)

        channel_name = live.get("channelName")
        live_title = live.get("liveTitle")
        category = live.get("liveCategoryValue")
        viewers = to_int(live.get("concurrentUserCount"))
        tags = ",".join(live.get("tags", [])) if live.get("tags") else ""
        open_date = live.get("openDate")
        thumbnail_url = live.get("thumbnailUrl")
        snapshot_time = now_str()

        snapshots_ws.append_row([
            "",
            snapshot_time,
            channel_id,
            channel_name,
            live_id,
            live_title,
            category,
            viewers,
            tags,
            open_date,
            thumbnail_url,
            "TRUE"
        ])

        collected_count += 1

        snapshots = snapshots_ws.get_all_records()
        peak_viewers, average_viewers = calculate_viewer_stats(snapshots, live_id)

        existing_session_row_number = None

        for index, session in enumerate(sessions, start=2):
            if str(session.get("live_id")) == live_id:
                existing_session_row_number = index
                break

        if existing_session_row_number:
            sessions_ws.update_cell(existing_session_row_number, 9, peak_viewers)
            sessions_ws.update_cell(existing_session_row_number, 10, average_viewers)
            sessions_ws.update_cell(existing_session_row_number, 11, category)
            sessions_ws.update_cell(existing_session_row_number, 12, tags)

            updated_session_count += 1
            print("세션 업데이트:", channel_name, "/", live_title)
            print("최고 시청자:", peak_viewers, "평균 시청자:", average_viewers)

        else:
            sessions_ws.append_row([
                "",
                live_id,
                channel_id,
                channel_name,
                live_title,
                open_date,
                "",
                "",
                peak_viewers,
                average_viewers,
                category,
                tags,
                run_time
            ])

            created_session_count += 1
            print("새 세션 생성:", channel_name, "/", live_title)
            print("최고 시청자:", peak_viewers, "평균 시청자:", average_viewers)

    # 4. 방송 종료 감지
    sessions = sessions_ws.get_all_records()
    snapshots = snapshots_ws.get_all_records()

    for index, session in enumerate(sessions, start=2):
        session_live_id = str(session.get("live_id"))
        session_channel_id = str(session.get("channel_id"))
        end_time = str(session.get("end_time")).strip()

        if session_channel_id not in target_channel_ids:
            continue

        if end_time:
            continue

        peak_viewers, average_viewers = calculate_viewer_stats(snapshots, session_live_id)

        if peak_viewers > 0:
            sessions_ws.update_cell(index, 9, peak_viewers)
            sessions_ws.update_cell(index, 10, average_viewers)

        if session_live_id and session_live_id not in current_live_ids:
            ended_at = run_time
            start_time_raw = session.get("start_time")
            start_dt = parse_datetime(start_time_raw)
            end_dt = parse_datetime(ended_at)

            duration_minutes = ""

            if start_dt and end_dt:
                duration_minutes = round((end_dt - start_dt).total_seconds() / 60)

            sessions_ws.update_cell(index, 7, ended_at)
            sessions_ws.update_cell(index, 8, duration_minutes)

            ended_session_count += 1
            print("방송 종료 처리:", session.get("channel_name"), "/", session.get("live_title"))

    # 5. creators 현재 상태 업데이트
    sessions = sessions_ws.get_all_records()

    for channel_id in target_channel_ids:
        row_number = creator_rows_by_channel_id.get(channel_id)

        if not row_number:
            continue

        live = live_by_channel_id.get(channel_id)
        is_live = live is not None

        live_title = ""
        current_viewers = 0

        if is_live:
            live_title = live.get("liveTitle")
            current_viewers = to_int(live.get("concurrentUserCount"))

        current_followers = latest_followers_by_channel_id.get(channel_id, "")
        total_live_count, avg_live_duration, last_live_time = calculate_creator_summary(sessions, channel_id)

        update_creator_status(
            creators_ws=creators_ws,
            row_number=row_number,
            header_map=creator_header_map,
            is_live=is_live,
            live_title=live_title,
            current_viewers=current_viewers,
            current_followers=current_followers,
            last_live_time=last_live_time,
            total_live_count=total_live_count,
            avg_live_duration=avg_live_duration,
            checked_at=run_time
        )

        print("creators 상태 업데이트:", channel_id)

    # 6. 오래된 live_snapshots 아카이브
    archived_count = archive_old_snapshots(snapshots_ws, archive_ws, days=30)

    logs_ws.append_row([
        "",
        run_time,
        "SUCCESS",
        collected_count,
        "",
        ""
    ])

    print("전체 완료!")
    print("팔로워 기록 수:", follower_collected_count)
    print("스냅샷 기록 수:", collected_count)
    print("새 세션 수:", created_session_count)
    print("업데이트 세션 수:", updated_session_count)
    print("종료 처리 수:", ended_session_count)
    print("아카이브 이동 수:", archived_count)

except Exception as e:
    logs_ws.append_row([
        "",
        run_time,
        "ERROR",
        0,
        str(e),
        ""
    ])

    print("에러 발생:")
    print(e)