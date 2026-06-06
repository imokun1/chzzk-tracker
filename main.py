"""
CHZZK Tracker - 버튜버 활동 데이터 수집 스크립트
GitHub Actions 또는 로컬에서 주기적으로 실행되어
치지직 공식 API로 채널 정보 / 라이브 정보를 수집하고
Google Sheets에 누적 저장합니다.

업데이트:
- 모든 시계열 시트에 agency, generation 컬럼 포함
- live_snapshots, live_sessions, collection_logs, live_snapshots_archive
  시트의 'id' 컬럼(A열) 제거. A열이 비어있어서 발생하는 append 위치
  버그 해결.
"""

import os
import json
import time
import traceback
import requests
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone, timedelta

# 로컬 .env 사용 시에만 동작 (Actions에선 패스)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


# ============================================================
# 설정
# ============================================================
CLIENT_ID = os.getenv("CHZZK_CLIENT_ID")
CLIENT_SECRET = os.getenv("CHZZK_CLIENT_SECRET")
SHEET_NAME = "치지직_버튜버_활동기록"
GOOGLE_CREDS_FILE = "google_credentials.json"
GOOGLE_CREDS_ENV = "GOOGLE_CREDENTIALS_JSON"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

KST = timezone(timedelta(hours=9))

CHZZK_API_BASE = "https://openapi.chzzk.naver.com/open/v1"
LIVES_ENDPOINT = f"{CHZZK_API_BASE}/lives"
CHANNELS_ENDPOINT = f"{CHZZK_API_BASE}/channels"

MAX_RETRIES = 3
RETRY_BACKOFF = 2


# ============================================================
# 유틸 함수
# ============================================================
def now_str():
    return datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")


def to_int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def is_truthy(value):
    if value is True:
        return True
    if value is False or value is None:
        return False
    return str(value).strip().lower() == "true"


def parse_datetime(value):
    if not value:
        return None
    value = str(value).strip()
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
    except (ValueError, TypeError):
        pass
    try:
        return datetime.strptime(value[:19], "%Y-%m-%d %H:%M:%S")
    except (ValueError, TypeError):
        return None


def get_header_map(worksheet):
    headers = worksheet.row_values(1)
    return {header: index + 1 for index, header in enumerate(headers)}


def col_letter(col_number):
    result = ""
    while col_number > 0:
        col_number, remainder = divmod(col_number - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ============================================================
# Google Sheets 인증
# ============================================================
def get_sheets_client():
    creds_env = os.getenv(GOOGLE_CREDS_ENV)
    if creds_env:
        info = json.loads(creds_env)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)


# ============================================================
# 치지직 API 호출
# ============================================================
def chzzk_request(url, params=None):
    headers = {
        "Client-Id": CLIENT_ID,
        "Client-Secret": CLIENT_SECRET,
        "Content-Type": "application/json",
    }

    last_error = None
    for attempt in range(MAX_RETRIES):
        try:
            response = requests.get(url, params=params, headers=headers, timeout=15)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            last_error = e
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_BACKOFF * (attempt + 1))
            else:
                raise last_error


def get_channels_info(channel_ids):
    if not channel_ids:
        return {}

    result = {}
    chunk_size = 20
    for i in range(0, len(channel_ids), chunk_size):
        chunk = channel_ids[i:i + chunk_size]
        params = {"channelIds": ",".join(chunk)}
        data = chzzk_request(CHANNELS_ENDPOINT, params=params)
        items = data.get("content", {}).get("data", [])
        for item in items:
            cid = item.get("channelId")
            if cid:
                result[cid] = {
                    "channel_id": cid,
                    "channel_name": item.get("channelName", ""),
                    "follower_count": to_int(item.get("followerCount")),
                }
    return result


def get_all_lives(max_pages=10):
    all_lives = []
    next_token = None

    for _ in range(max_pages):
        params = {"size": 20}
        if next_token:
            params["next"] = next_token

        data = chzzk_request(LIVES_ENDPOINT, params=params)
        content = data.get("content", {})
        lives = content.get("data", [])
        all_lives.extend(lives)

        next_token = content.get("page", {}).get("next")
        if not next_token:
            break

    return all_lives


# ============================================================
# 데이터 처리
# ============================================================
def calculate_viewer_stats(snapshots_cache, live_id):
    viewers_list = [
        to_int(s.get("current_viewers"))
        for s in snapshots_cache
        if str(s.get("live_id")) == str(live_id)
    ]
    if not viewers_list:
        return 0, 0
    return max(viewers_list), round(sum(viewers_list) / len(viewers_list))


def calculate_last_live_time(sessions_cache, channel_id):
    last_live_time = ""

    for s in sessions_cache:
        if str(s.get("channel_id")) != str(channel_id):
            continue
        start = str(s.get("start_time", "")).strip()
        if start and (not last_live_time or start > last_live_time):
            last_live_time = start

    return last_live_time


# ============================================================
# 배치 업데이트 헬퍼
# ============================================================
def make_update(worksheet_title, row, col, value):
    return {
        "range": f"'{worksheet_title}'!{col_letter(col)}{row}",
        "values": [[value]],
    }


def flush_updates(spreadsheet, updates):
    if not updates:
        return
    spreadsheet.values_batch_update({
        "valueInputOption": "USER_ENTERED",
        "data": updates,
    })


# ============================================================
# 아카이브
# ============================================================
def archive_old_snapshots(snapshots_ws, archive_ws, days=30):
    all_values = snapshots_ws.get_all_values()
    if len(all_values) <= 1:
        return 0

    cutoff = datetime.now(KST).replace(tzinfo=None) - timedelta(days=days)
    rows_to_archive = []
    row_numbers_to_delete = []

    # ★ id 컬럼 제거로 snapshot_time이 이제 A열(인덱스 0)
    for i, row in enumerate(all_values[1:], start=2):
        if len(row) < 1:
            continue
        snapshot_dt = parse_datetime(row[0])  # ★ 변경: row[1] → row[0]
        if snapshot_dt and snapshot_dt < cutoff:
            rows_to_archive.append(row)
            row_numbers_to_delete.append(i)

    if not rows_to_archive:
        return 0

    archive_ws.append_rows(rows_to_archive, value_input_option="USER_ENTERED")
    for row_number in reversed(row_numbers_to_delete):
        snapshots_ws.delete_rows(row_number)

    return len(rows_to_archive)


# ============================================================
# 메인 로직
# ============================================================
def main():
    if not CLIENT_ID or not CLIENT_SECRET:
        raise RuntimeError("CHZZK_CLIENT_ID / CHZZK_CLIENT_SECRET 환경변수가 비어있습니다.")

    client = get_sheets_client()
    sheet = client.open(SHEET_NAME)

    creators_ws = sheet.worksheet("creators")
    snapshots_ws = sheet.worksheet("live_snapshots")
    archive_ws = sheet.worksheet("live_snapshots_archive")
    sessions_ws = sheet.worksheet("live_sessions")
    followers_ws = sheet.worksheet("follower_snapshots")
    logs_ws = sheet.worksheet("collection_logs")

    run_time = now_str()

    # --- 1단계: 시트 데이터 일괄 로드 ---
    creators = creators_ws.get_all_records()
    creator_header_map = get_header_map(creators_ws)
    sessions_cache = sessions_ws.get_all_records()
    snapshots_cache = snapshots_ws.get_all_records()

    target_channel_ids = []
    creator_rows_by_channel_id = {}
    creator_meta_by_channel_id = {}

    for index, creator in enumerate(creators, start=2):
        if is_truthy(creator.get("is_active")):
            channel_id = str(creator.get("channel_id", "")).strip()
            if channel_id:
                target_channel_ids.append(channel_id)
                creator_rows_by_channel_id[channel_id] = index
                creator_meta_by_channel_id[channel_id] = {
                    "agency": str(creator.get("agency", "")).strip(),
                    "generation": str(creator.get("generation", "")).strip(),
                }

    print(f"추적 대상 채널 수: {len(target_channel_ids)}")
    if not target_channel_ids:
        print("추적 대상이 없습니다. 종료.")
        # ★ id 컬럼 제거: 5개 컬럼으로 축소
        logs_ws.append_row([run_time, "SUCCESS", 0, "no targets", ""])
        return

    # --- 2단계: 공식 API로 채널 정보 일괄 조회 ---
    follower_collected_count = 0
    latest_followers_by_channel_id = {}
    follower_rows_to_append = []

    try:
        channels_info = get_channels_info(target_channel_ids)
        for cid, info in channels_info.items():
            meta = creator_meta_by_channel_id.get(cid, {})
            follower_rows_to_append.append([
                run_time,
                info["channel_id"],
                info["channel_name"],
                info["follower_count"],
                meta.get("agency", ""),
                meta.get("generation", ""),
            ])
            latest_followers_by_channel_id[cid] = info["follower_count"]
            follower_collected_count += 1
            print(f"팔로워 기록: {info['channel_name']} {info['follower_count']}")

        if follower_rows_to_append:
            followers_ws.append_rows(
                follower_rows_to_append,
                value_input_option="USER_ENTERED",
            )
    except Exception as e:
        print(f"채널 정보 조회 실패: {e}")
        traceback.print_exc()

    # --- 3단계: 라이브 목록 조회 ---
    try:
        lives = get_all_lives(max_pages=50)
        print(f"전체 라이브 수신: {len(lives)}")
    except Exception as e:
        print(f"라이브 목록 조회 실패: {e}")
        traceback.print_exc()
        lives = []

    # --- 4단계: 추적 대상 라이브만 필터 ---
    target_set = set(target_channel_ids)
    target_lives = [l for l in lives if l.get("channelId") in target_set]
    print(f"추적 대상 중 라이브 중: {len(target_lives)}")

    current_live_ids = set()
    live_by_channel_id = {}
    snapshot_rows_to_append = []
    session_rows_to_append = []
    session_updates = []

    collected_count = 0
    created_session_count = 0
    updated_session_count = 0
    ended_session_count = 0

    # --- 5단계: 라이브 데이터 처리 ---
    for live in target_lives:
        channel_id = live.get("channelId")
        live_id = str(live.get("liveId"))
        current_live_ids.add(live_id)
        live_by_channel_id[channel_id] = live

        channel_name = live.get("channelName", "")
        live_title = live.get("liveTitle", "")
        category = live.get("liveCategoryValue", "")
        viewers = to_int(live.get("concurrentUserCount"))
        tags = ",".join(live.get("tags", [])) if live.get("tags") else ""
        open_date = live.get("openDate", "")
        thumbnail_url = live.get("thumbnailUrl", "")
        snapshot_time = now_str()

        meta = creator_meta_by_channel_id.get(channel_id, {})
        agency = meta.get("agency", "")
        generation = meta.get("generation", "")

        # ★ id 컬럼 제거: 13개 컬럼 (snapshot_time이 첫 컬럼)
        snapshot_row = [
            snapshot_time, channel_id, channel_name, live_id,
            live_title, category, viewers, tags, open_date,
            thumbnail_url, "TRUE",
            agency, generation,
        ]
        snapshot_rows_to_append.append(snapshot_row)
        collected_count += 1

        snapshots_cache.append({
            "live_id": live_id,
            "current_viewers": viewers,
        })

        peak, avg = calculate_viewer_stats(snapshots_cache, live_id)

        existing_row = None
        for idx, s in enumerate(sessions_cache, start=2):
            if str(s.get("live_id")) == live_id:
                existing_row = idx
                break

        # ★ id 컬럼 제거로 모든 컬럼 번호가 1씩 감소
        # 기존: 9(peak), 10(avg), 11(category), 12(tags)
        # 변경: 8(peak), 9(avg), 10(category), 11(tags)
        if existing_row:
            session_updates.append(make_update(sessions_ws.title, existing_row, 8, peak))
            session_updates.append(make_update(sessions_ws.title, existing_row, 9, avg))
            session_updates.append(make_update(sessions_ws.title, existing_row, 10, category))
            session_updates.append(make_update(sessions_ws.title, existing_row, 11, tags))
            updated_session_count += 1
        else:
            # ★ id 컬럼 제거: 14개 컬럼 (live_id가 첫 컬럼)
            new_row = [
                live_id, channel_id, channel_name, live_title,
                open_date, "", "", peak, avg, category, tags, run_time,
                agency, generation,
            ]
            session_rows_to_append.append(new_row)
            sessions_cache.append({
                "live_id": live_id,
                "channel_id": channel_id,
                "start_time": open_date,
                "end_time": "",
                "duration_minutes": "",
            })
            created_session_count += 1

    # --- 6단계: 종료된 세션 처리 ---
    for idx, s in enumerate(sessions_cache, start=2):
        if idx > len(sessions_cache) + 1 - len(session_rows_to_append):
            continue
        session_channel_id = str(s.get("channel_id", ""))
        session_live_id = str(s.get("live_id", ""))
        end_time = str(s.get("end_time", "")).strip()

        if session_channel_id not in target_set:
            continue
        if end_time:
            continue
        if not session_live_id:
            continue
        if session_live_id in current_live_ids:
            continue

        peak, avg = calculate_viewer_stats(snapshots_cache, session_live_id)
        start_dt = parse_datetime(s.get("start_time"))
        end_dt = parse_datetime(run_time)
        duration_minutes = ""
        if start_dt and end_dt:
            duration_minutes = round((end_dt - start_dt).total_seconds() / 60)

        # ★ id 컬럼 제거로 컬럼 번호가 1씩 감소
        # 기존: 7(end_time), 8(duration), 9(peak), 10(avg)
        # 변경: 6(end_time), 7(duration), 8(peak), 9(avg)
        session_updates.append(make_update(sessions_ws.title, idx, 6, run_time))
        session_updates.append(make_update(sessions_ws.title, idx, 7, duration_minutes))
        if peak > 0:
            session_updates.append(make_update(sessions_ws.title, idx, 8, peak))
            session_updates.append(make_update(sessions_ws.title, idx, 9, avg))
        ended_session_count += 1

    # --- 7단계: append할 행들 한 번에 추가 ---
    if snapshot_rows_to_append:
        snapshots_ws.append_rows(snapshot_rows_to_append, value_input_option="USER_ENTERED")
    if session_rows_to_append:
        sessions_ws.append_rows(session_rows_to_append, value_input_option="USER_ENTERED")

    # --- 8단계: 세션 업데이트 일괄 적용 ---
    flush_updates(sheet, session_updates)

    # --- 9단계: creators 시트 상태 갱신 ---
    sessions_cache_refreshed = sessions_ws.get_all_records()
    creator_updates = []

    for channel_id in target_channel_ids:
        row_number = creator_rows_by_channel_id.get(channel_id)
        if not row_number:
            continue

        live = live_by_channel_id.get(channel_id)
        is_live = live is not None
        live_title = live.get("liveTitle", "") if is_live else ""
        current_viewers = to_int(live.get("concurrentUserCount")) if is_live else 0
        current_followers = latest_followers_by_channel_id.get(channel_id, "")

        last_live_time = calculate_last_live_time(sessions_cache_refreshed, channel_id)

        update_map = {
            "current_live_status": "TRUE" if is_live else "FALSE",
            "current_live_title": live_title,
            "current_viewers": current_viewers,
            "current_followers": current_followers,
            "last_live_time": last_live_time,
            "last_checked_at": run_time,
        }

        for col_name, val in update_map.items():
            col = creator_header_map.get(col_name)
            if col:
                creator_updates.append(make_update(creators_ws.title, row_number, col, val))

    flush_updates(sheet, creator_updates)

    # --- 10단계: 아카이브 ---
    archived_count = archive_old_snapshots(snapshots_ws, archive_ws, days=366)

    # --- 11단계: 로그 기록 ---
    # ★ id 컬럼 제거: 5개 컬럼으로 축소
    logs_ws.append_row([
        run_time, "SUCCESS", collected_count, "", "",
    ], value_input_option="USER_ENTERED")

    print("=" * 40)
    print("전체 완료!")
    print(f"팔로워 기록 수: {follower_collected_count}")
    print(f"스냅샷 기록 수: {collected_count}")
    print(f"새 세션 수: {created_session_count}")
    print(f"업데이트 세션 수: {updated_session_count}")
    print(f"종료 처리 수: {ended_session_count}")
    print(f"아카이브 이동 수: {archived_count}")


# ============================================================
# 실행
# ============================================================
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("=" * 40)
        print("에러 발생:")
        print(e)
        traceback.print_exc()
        try:
            client = get_sheets_client()
            sheet = client.open(SHEET_NAME)
            logs_ws = sheet.worksheet("collection_logs")
            # ★ id 컬럼 제거: 5개 컬럼으로 축소
            logs_ws.append_row(
                [now_str(), "ERROR", 0, str(e)[:500], ""],
                value_input_option="USER_ENTERED",
            )
        except Exception as log_err:
            print(f"로그 기록도 실패: {log_err}")
        raise
