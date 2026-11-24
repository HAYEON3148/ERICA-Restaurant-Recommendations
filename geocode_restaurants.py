"""ERICA 식당 주소를 Google Maps Platform Geocoding API v4로 좌표 변환합니다.

사용법:
1) 이 파일과 erica.restaurant.data.csv를 같은 폴더에 둡니다.
2) python geocode_restaurants.py 실행
3) 터미널에 Google Maps Platform API 키를 입력합니다(입력값은 화면에 표시되지 않음).

원본 CSV는 수정하지 않으며, 결과는 별도 Excel 파일로 저장합니다.
"""

import getpass
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "erica.restaurant.data.csv"
OUTPUT_FILE = BASE_DIR / "ERICA_restaurants_geocoded_google.xlsx"

# 한양대 ERICA 주변의 결과인지 확인하기 위한 넉넉한 검증 범위
MIN_LAT, MAX_LAT = 37.25, 37.35
MIN_LNG, MAX_LNG = 126.78, 126.89


def normalize_address(value):
    """층/호수 정보를 제거하고 주소 표기를 정리합니다."""
    address = re.sub(r"\s+", " ", str(value)).strip()
    address = re.sub(r"\s+\d+\s*층.*$", "", address)
    address = re.sub(r"\s+\d+\s*호.*$", "", address)
    address = address.replace("상록구한양", "상록구 한양")
    address = address.replace("한양대학로15", "한양대학로 15")
    address = re.sub(r"^경기\s+", "경기도 ", address)
    return address.strip()


def make_query(value):
    address = normalize_address(value)
    if "안산" not in address:
        address = "경기도 안산시 상록구 " + address
    elif "경기도" not in address and address.startswith("경기 "):
        address = "경기도 " + address[3:]
    if "대한민국" not in address:
        address += " 대한민국"
    return address


def geocode_google_v4(address, api_key):
    """Google 공식 Geocoding API v4의 GeocodeAddress 엔드포인트를 호출합니다."""
    query = make_query(address)
    encoded_address = urllib.parse.quote(query, safe="")
    params = urllib.parse.urlencode({
        "regionCode": "KR",
        "languageCode": "ko",
        "locationBias.rectangle.low.latitude": "37.27",
        "locationBias.rectangle.low.longitude": "126.81",
        "locationBias.rectangle.high.latitude": "37.33",
        "locationBias.rectangle.high.longitude": "126.87",
    })
    url = f"https://geocode.googleapis.com/v4/geocode/address/{encoded_address}?{params}"

    request = urllib.request.Request(
        url,
        headers={
            "X-Goog-Api-Key": api_key,
            "X-Goog-FieldMask": "results.placeId,results.location,results.formattedAddress",
        },
        method="GET",
    )

    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        try:
            message = json.loads(body).get("error", {}).get("message", body)
        except json.JSONDecodeError:
            message = body
        return {"status": f"HTTP_{exc.code}", "message": str(message)[:500]}
    except Exception as exc:
        return {"status": "REQUEST_ERROR", "message": str(exc)[:500]}

    results = payload.get("results") or []
    if not results:
        return {"status": "NOT_FOUND", "message": "Google Geocoding 결과가 없습니다."}

    # 결과 중 안산 검증 범위 안에 있고 반환 주소에도 안산이 표시되는 좌표만 채택
    for item in results:
        location = item.get("location") or {}
        lat = location.get("latitude")
        lng = location.get("longitude")
        formatted = item.get("formattedAddress", "")

        if not isinstance(lat, (int, float)) or not isinstance(lng, (int, float)):
            continue

        in_ansan_bounds = (
            MIN_LAT <= lat <= MAX_LAT and MIN_LNG <= lng <= MAX_LNG
        )
        if in_ansan_bounds and "안산" in formatted:
            return {
                "status": "OK",
                "latitude": lat,
                "longitude": lng,
                "google_address": formatted,
                "place_id": item.get("placeId", ""),
                "message": "",
            }

    first = results[0]
    return {
        "status": "REVIEW_REQUIRED",
        "google_address": first.get("formattedAddress", ""),
        "place_id": first.get("placeId", ""),
        "message": "Google 결과가 안산 검증 범위와 일치하지 않아 좌표를 입력하지 않았습니다.",
    }


def main():
    if not INPUT_FILE.is_file():
        raise FileNotFoundError(
            f"원본 CSV를 찾을 수 없습니다: {INPUT_FILE.name}\n"
            "app.py와 같은 프로젝트 폴더에 원본 CSV를 그대로 두어 주세요."
        )

    try:
        data = pd.read_csv(INPUT_FILE, encoding="utf-8-sig")
    except UnicodeDecodeError:
        data = pd.read_csv(INPUT_FILE, encoding="cp949")

    required_columns = {"상호", "주소"}
    missing = required_columns - set(data.columns)
    if missing:
        raise ValueError(f"CSV에 필요한 열이 없습니다: {', '.join(sorted(missing))}")

    # 원본 주소별 결과를 기록하되, 층/호수 제거 후 같은 주소는 API를 한 번만 호출
    address_values = data["주소"].dropna().astype(str).unique().tolist()
    query_by_address = {address: make_query(address) for address in address_values}
    unique_queries = list(dict.fromkeys(query_by_address.values()))

    print(f"원본 파일: {INPUT_FILE.name}")
    print(f"식당 수: {len(data)}개")
    print(f"원본 주소 수: {len(address_values)}개")
    print(f"실제 Google API 요청 수(중복 정리 후): {len(unique_queries)}개")
    print("사용 API: Google Maps Platform - Geocoding API v4 / GeocodeAddress")

    api_key = getpass.getpass("Google Maps Platform API 키 입력(입력값 숨김): ").strip()
    if not api_key:
        print("API 키가 입력되지 않아 종료합니다.")
        return

    result_by_query = {}
    for number, query in enumerate(unique_queries, start=1):
        print(f"[{number}/{len(unique_queries)}] {query}", flush=True)
        result = geocode_google_v4(query, api_key)
        result_by_query[query] = result
        print(
            f"  {result['status']} | "
            f"{result.get('google_address', result.get('message', ''))}",
            flush=True,
        )
        if number < len(unique_queries):
            time.sleep(0.1)

    def result_for(address):
        if pd.isna(address):
            return {}
        return result_by_query.get(query_by_address.get(str(address), ""), {})

    output = data.copy()
    output["위도"] = output["주소"].map(lambda address: result_for(address).get("latitude"))
    output["경도"] = output["주소"].map(lambda address: result_for(address).get("longitude"))
    output["좌표 확인 상태"] = output["주소"].map(
        lambda address: result_for(address).get("status", "주소 없음")
    )
    output["Google 반환 주소"] = output["주소"].map(
        lambda address: result_for(address).get("google_address", "")
    )
    output["Google Place ID"] = output["주소"].map(
        lambda address: result_for(address).get("place_id", "")
    )
    output["오류/검토 메모"] = output["주소"].map(
        lambda address: result_for(address).get("message", "")
    )

    review_rows = []
    for original_address in address_values:
        result = result_by_query[query_by_address[original_address]]
        review_rows.append({
            "원본 주소": original_address,
            "Google 요청 주소": query_by_address[original_address],
            "위도": result.get("latitude"),
            "경도": result.get("longitude"),
            "좌표 확인 상태": result.get("status", ""),
            "Google 반환 주소": result.get("google_address", ""),
            "Google Place ID": result.get("place_id", ""),
            "오류/검토 메모": result.get("message", ""),
        })
    review = pd.DataFrame(review_rows)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        output.to_excel(writer, index=False, sheet_name="식당데이터")
        review.to_excel(writer, index=False, sheet_name="주소별좌표검토")
        for sheet_name in ("식당데이터", "주소별좌표검토"):
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cells in worksheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in cells)
                worksheet.column_dimensions[cells[0].column_letter].width = min(
                    max(max_length + 2, 12), 48
                )

    successful = sum(r.get("status") == "OK" for r in result_by_query.values())
    review_needed = sum(r.get("status") == "REVIEW_REQUIRED" for r in result_by_query.values())
    not_found = sum(r.get("status") == "NOT_FOUND" for r in result_by_query.values())
    errors = sum(
        r.get("status", "").startswith(("HTTP_", "REQUEST_ERROR"))
        for r in result_by_query.values()
    )

    print("\n========== 완료 ==========")
    print(f"좌표 확인 성공: {successful}/{len(unique_queries)}개 요청 주소")
    print(f"검토 필요: {review_needed}개 / 검색 결과 없음: {not_found}개 / 오류: {errors}개")
    print(f"엑셀 저장 위치: {OUTPUT_FILE}")
    print("원본 CSV는 변경하지 않았으며, 검증되지 않은 좌표는 임의로 입력하지 않았습니다.")


if __name__ == "__main__":
    main()
