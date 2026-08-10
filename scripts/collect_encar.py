"""
엔카 광고등록 대수 자동 수집 스크립트
- 국산(CarType.Y) / 수입(CarType.A) 각각에 대해
- 제조사 -> 차종그룹(ModelGroup) -> 모델(Model) 3단계 트리를 순회하며 대수를 수집
- 결과를 구글 시트에 (수집일자, 국산수입구분, 제조사, 차종그룹, 모델, 대수) 형태로 한 줄씩 기록

환경변수로 아래 두 값을 전달받습니다 (GitHub Actions Secrets에서 주입):
  GOOGLE_SERVICE_ACCOUNT_KEY : 구글 서비스계정 JSON 키 전체 내용 (문자열)
  SHEET_ID                   : 데이터를 쓸 구글시트 ID
"""

import json
import os
import time
import datetime
import urllib.parse
import requests
import gspread
from google.oauth2.service_account import Credentials

BASE_URL = "https://api.encar.com/search/car/list/general"

# 요청 간 최소 간격(초). 너무 빠르게 두드리면 차단될 수 있어 여유를 둡니다.
REQUEST_DELAY = 0.3
MAX_RETRIES = 3

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "application/json",
    "Origin": "https://www.encar.com",
    "Referer": "https://www.encar.com/",
}

# CarType 코드: Y=국산, A=수입 으로 추정 (최초 1회 실행 후 실제 수치로 검증 필요)
CAR_TYPES = {
    "국산": "Y",
    "수입": "A",
}


def build_query(car_type, manufacturer=None, model_group=None):
    """엔카 API 쿼리 문자열(q 파라미터)을 조립합니다."""
    if manufacturer is None:
        # 제조사 목록만 필요한 단계
        inner = f"CarType.{car_type}."
    elif model_group is None:
        # 특정 제조사의 차종그룹 목록이 필요한 단계
        inner = f"(C.CarType.{car_type}._.Manufacturer.{manufacturer}.)"
    else:
        # 특정 제조사+차종그룹의 모델 목록이 필요한 단계
        inner = (
            f"(C.CarType.{car_type}._."
            f"(C.Manufacturer.{manufacturer}._.ModelGroup.{model_group}.))"
        )
    return f"(And.Hidden.N._.{inner})"


def fetch(query):
    """API를 호출하고 JSON을 반환합니다. 실패 시 재시도합니다."""
    params = {"count": "true", "q": query}
    url = BASE_URL + "?" + urllib.parse.urlencode(params, safe="().,_")

    last_err = None
    for attempt in range(MAX_RETRIES):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            time.sleep(REQUEST_DELAY)
            return resp.json()
        except Exception as e:  # noqa: BLE001
            last_err = e
            time.sleep(1 + attempt)
    raise RuntimeError(f"요청 실패: {query[:80]}... / {last_err}")


def get_facet_list(data, aspect_name):
    """iNav.Nodes 안에서 특정 Aspect(Name)의 Facets 목록을 찾아 반환합니다."""
    nodes = data.get("iNav", {}).get("Nodes", [])
    return _find_aspect(nodes, aspect_name)


def _find_aspect(nodes, aspect_name):
    for node in nodes:
        if node.get("Name") == aspect_name:
            return node.get("Facets", [])
        # 선택된 facet 안에 중첩된 다음 단계 Aspect가 있을 수 있음
        for facet in node.get("Facets", []):
            refinements = facet.get("Refinements", {})
            if refinements:
                result = _find_aspect(refinements.get("Nodes", []), aspect_name)
                if result:
                    return result
    return []


def collect_all():
    """국산/수입 전체를 순회하며 (구분, 제조사, 차종그룹, 모델, 대수) 리스트를 만듭니다."""
    rows = []
    today = datetime.date.today().isoformat()

    for label, car_type in CAR_TYPES.items():
        print(f"=== {label} (CarType.{car_type}) 수집 시작 ===")

        # 1단계: 제조사 목록
        top_query = build_query(car_type)
        top_data = fetch(top_query)
        manufacturers = get_facet_list(top_data, "Manufacturer")

        for m in manufacturers:
            m_name = m["Value"]
            m_count = m["Count"]
            if m_count == 0:
                continue

            # 2단계: 해당 제조사의 차종그룹 목록
            mg_query = build_query(car_type, manufacturer=m_name)
            mg_data = fetch(mg_query)
            model_groups = get_facet_list(mg_data, "ModelGroup")

            if not model_groups:
                # 차종그룹이 없으면 제조사 레벨 대수만 기록
                rows.append([today, label, m_name, "", "", m_count])
                continue

            for mg in model_groups:
                mg_name = mg["Value"]
                mg_count = mg["Count"]
                if mg_count == 0:
                    continue

                # 3단계: 해당 제조사+차종그룹의 모델 목록
                model_query = build_query(car_type, manufacturer=m_name, model_group=mg_name)
                model_data = fetch(model_query)
                models = get_facet_list(model_data, "Model")

                if not models:
                    rows.append([today, label, m_name, mg_name, "", mg_count])
                    continue

                for md in models:
                    md_name = md["Value"]
                    md_count = md["Count"]
                    rows.append([today, label, m_name, mg_name, md_name, md_count])

            print(f"  {m_name} 완료 (누적 {len(rows)}행)")

    return rows


def write_to_sheet(rows):
    """수집 결과를 구글 시트에 기록합니다. 매일 새 워크시트(탭)를 날짜로 만듭니다."""
    key_json = os.environ["GOOGLE_SERVICE_ACCOUNT_KEY"]
    sheet_id = os.environ["SHEET_ID"]

    creds_info = json.loads(key_json)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    client = gspread.authorize(creds)

    sh = client.open_by_key(sheet_id)

    today = datetime.date.today().isoformat()
    try:
        ws = sh.worksheet(today)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=today, rows=str(len(rows) + 10), cols="6")

    header = ["수집일자", "국산/수입", "제조사", "차종그룹", "모델", "등록대수"]
    ws.update("A1", [header] + rows, value_input_option="RAW")

    print(f"구글시트에 {len(rows)}행 기록 완료 (탭: {today})")


if __name__ == "__main__":
    all_rows = collect_all()
    print(f"총 {len(all_rows)}행 수집 완료")
    write_to_sheet(all_rows)
