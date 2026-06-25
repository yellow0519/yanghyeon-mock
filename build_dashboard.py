#!/usr/bin/env python3
"""
양현고등학교 모의고사 대시보드 빌드 스크립트

사용법:
  python build_dashboard.py                    # data/ 폴더의 엑셀을 읽어 docs/index.html 생성
  python build_dashboard.py --push             # 생성 후 GitHub Pages에 자동 배포
  python build_dashboard.py --push -m "5월 추가" # 커밋 메시지 지정

폴더 구조:
  data/
    2026.03/          ← 시험 회차별 폴더 (yyyy.mm 형식)
      1학년.xlsx
      2학년.xlsx
      3학년.xlsx
    2026.05/
      1학년.xlsx
      ...
  template.html       ← 대시보드 HTML 템플릿
  docs/
    index.html         ← 생성된 대시보드 (GitHub Pages 배포용)
"""

import pandas as pd
import numpy as np
import json
import os
import sys
import argparse
import subprocess
from pathlib import Path

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
TEMPLATE_PATH = BASE_DIR / "template.html"
OUTPUT_DIR = BASE_DIR / "docs"
OUTPUT_PATH = OUTPUT_DIR / "index.html"


# ──────────────────────────────────────────────
# 엑셀 파싱 함수들
# ──────────────────────────────────────────────

def safe_float(v):
    try:
        if pd.isna(v): return None
        return float(v)
    except: return None

def safe_int(v):
    try:
        if pd.isna(v): return None
        return int(float(v))
    except: return None

def safe_str(v):
    if pd.isna(v): return ""
    return str(v).strip()


def parse_grade1(filepath):
    """1학년 파싱: '성적 정보' 또는 '예상점수' 시트"""
    xl = pd.ExcelFile(filepath)
    sheet = None
    for s in xl.sheet_names:
        if '성적' in s and '정보' in s:
            sheet = s
            break
    if not sheet:
        for s in xl.sheet_names:
            if '예상점수' in s:
                sheet = s
                break
    if not sheet:
        sheet = xl.sheet_names[0]
        print(f"  [1학년] 적절한 시트를 찾지 못해 '{sheet}' 사용")

    df = pd.read_excel(filepath, sheet_name=sheet, header=None)
    students = []

    if '예상점수' in sheet:
        # 예상점수 시트: r[0]=반, r[1]=번호, r[2]=성명
        # r[3]=국어과목명, r[4]=국어점수, r[5]=국어등급
        # r[6]=수학과목명, r[7]=수학점수, r[8]=수학등급
        # r[9]=영어점수,   r[10]=영어등급
        # r[13]=탐구1과목, r[14]=탐구1점수, r[15]=탐구1등급
        # r[16]=탐구2과목, r[17]=탐구2점수, r[18]=탐구2등급
        # r[19]=한국사점수, r[20]=한국사등급, r[21]=총점, r[22]=석차
        for i in range(3, len(df)):
            r = df.iloc[i]
            try:
                cls_val = safe_int(r[0])
                if cls_val is None: continue
                s = {
                    "cls": cls_val,
                    "num": safe_int(r[1]) or 0,
                    "name": safe_str(r[2]),
                    "국어": {"score": safe_float(r[4]), "grade": safe_int(r[5])},
                    "수학": {"score": safe_float(r[7]), "grade": safe_int(r[8])},
                    "영어": {"score": safe_float(r[9]), "grade": safe_int(r[10])},
                    "사탐": {"score": safe_float(r[14]), "grade": safe_int(r[15])},
                    "과탐": {"score": safe_float(r[17]), "grade": safe_int(r[18])},
                    "한국사": {"score": safe_float(r[19]), "grade": safe_int(r[20])},
                    "total": safe_float(r[21]),
                    "rank": safe_int(r[22])
                }
                if s["name"]:
                    students.append(s)
            except Exception:
                pass
    else:
        # 성적 정보 시트: r[1]=반, r[2]=번호, r[3]=성명
        for i in range(2, len(df)):
            r = df.iloc[i]
            try:
                cls_val = safe_int(r[1])
                if cls_val is None: continue
                s = {
                    "cls": cls_val,
                    "num": safe_int(r[2]) or 0,
                    "name": safe_str(r[3]),
                    "국어": {"score": safe_float(r[4]), "grade": safe_int(r[7])},
                    "수학": {"score": safe_float(r[8]), "grade": safe_int(r[11])},
                    "영어": {"score": safe_float(r[12]), "grade": safe_int(r[15])},
                    "사탐": {"score": safe_float(r[16]), "grade": safe_int(r[19])},
                    "과탐": {"score": safe_float(r[20]), "grade": safe_int(r[23])},
                    "한국사": {"score": safe_float(r[24]), "grade": safe_int(r[27])},
                    "total": safe_float(r[28]),
                    "rank": safe_int(r[29])
                }
                if s["name"]:
                    students.append(s)
            except Exception:
                pass
    return students


def parse_grade2(filepath):
    """2학년 파싱: '모의고사' 시트"""
    xl = pd.ExcelFile(filepath)
    sheet = None
    for s in xl.sheet_names:
        if '모의고사' in s or '2학년' in s:
            sheet = s
            break
    if not sheet:
        sheet = xl.sheet_names[0]
        print(f"  [2학년] 적절한 시트를 찾지 못해 '{sheet}' 사용")

    df = pd.read_excel(filepath, sheet_name=sheet, header=None)
    students = []
    # 데이터는 3행부터 (0: 제목, 1: 헤더1, 2: 헤더2)
    for i in range(3, len(df)):
        r = df.iloc[i]
        try:
            cls_val = safe_int(r[0])
            if cls_val is None: continue
            s = {
                "cls": cls_val,
                "num": safe_int(r[1]) or 0,
                "name": safe_str(r[2]),
                "국어": {"score": safe_float(r[3]), "grade": safe_int(r[4])},
                "수학": {"score": safe_float(r[5]), "grade": safe_int(r[6])},
                "영어": {"score": safe_float(r[7]), "grade": safe_int(r[8])},
                "한국사": {"score": safe_float(r[10]), "grade": safe_int(r[11])},
                "사탐": {"score": safe_float(r[12]), "grade": safe_int(r[13])},
                "과탐": {"score": safe_float(r[14]), "grade": safe_int(r[15])},
                "total": safe_float(r[16]),
                "rank": safe_int(r[17])
            }
            if s["name"]:
                students.append(s)
        except:
            pass
    return students


def parse_grade3(filepath):
    """3학년 파싱: '예상점수' 시트"""
    xl = pd.ExcelFile(filepath)
    sheet = None
    for s in xl.sheet_names:
        if '예상점수' in s:
            sheet = s
            break
    if not sheet:
        sheet = xl.sheet_names[0]
        print(f"  [3학년] '예상점수' 시트를 찾지 못해 '{sheet}' 사용")

    df = pd.read_excel(filepath, sheet_name=sheet, header=None)
    students = []
    for i in range(3, len(df)):
        r = df.iloc[i]
        try:
            cls_val = safe_int(r[0])
            if cls_val is None: continue
            s = {
                "cls": cls_val,
                "num": safe_int(r[1]) or 0,
                "name": safe_str(r[2]),
                "국어": {"score": safe_float(r[4]), "grade": safe_int(r[5]), "sub": safe_str(r[3])},
                "수학": {"score": safe_float(r[7]), "grade": safe_int(r[8]), "sub": safe_str(r[6])},
                "영어": {"score": safe_float(r[9]), "grade": safe_int(r[10])},
                "탐구1": {"score": safe_float(r[14]), "grade": safe_int(r[15]), "sub": safe_str(r[13])},
                "탐구2": {"score": safe_float(r[17]), "grade": safe_int(r[18]), "sub": safe_str(r[16])},
                "한국사": {"score": safe_float(r[19]), "grade": safe_int(r[20])},
                "total": safe_float(r[21]),
                "rank": safe_int(r[22])
            }
            if s["name"]:
                students.append(s)
        except:
            pass
    return students


PARSERS = {
    "1": parse_grade1,
    "2": parse_grade2,
    "3": parse_grade3,
}

GRADE_ALIASES = {
    "1학년": "1", "2학년": "2", "3학년": "3",
    "1": "1", "2": "2", "3": "3",
    "grade1": "1", "grade2": "2", "grade3": "3",
}

MONTH_LABELS = {
    "03": "3월", "04": "4월", "05": "5월", "06": "6월",
    "07": "7월", "09": "9월", "10": "10월", "11": "11월",
}


# ──────────────────────────────────────────────
# 메인 빌드 로직
# ──────────────────────────────────────────────

def scan_data_folder():
    """data/ 폴더를 스캔하여 시험별 파일 목록 반환"""
    exams = {}
    if not DATA_DIR.exists():
        print(f"오류: {DATA_DIR} 폴더가 없습니다.")
        print(f"  '{DATA_DIR}' 폴더를 만들고 시험별 하위 폴더(예: 2026.03/)에 엑셀 파일을 넣어주세요.")
        sys.exit(1)

    for exam_dir in sorted(DATA_DIR.iterdir()):
        if not exam_dir.is_dir(): continue
        exam_id = exam_dir.name  # e.g. "2026.03"
        month = exam_id.split(".")[-1] if "." in exam_id else ""
        label = MONTH_LABELS.get(month, exam_id)

        files = {}
        for f in exam_dir.iterdir():
            if not f.suffix.lower() in ('.xlsx', '.xls'): continue
            fname = f.stem
            for alias, grade in GRADE_ALIASES.items():
                if alias in fname:
                    files[grade] = f
                    break

        if files:
            exams[exam_id] = {"label": label, "files": files}

    return exams


def build_json(exams_info):
    """엑셀 파일들을 파싱하여 JSON 데이터 구조 생성"""
    result = {"exams": {}, "order": []}
    total_students = 0

    for exam_id in sorted(exams_info.keys()):
        info = exams_info[exam_id]
        print(f"\n📋 {exam_id} ({info['label']}) 처리 중...")
        exam_data = {}

        for grade, filepath in sorted(info["files"].items()):
            parser = PARSERS.get(grade)
            if not parser:
                print(f"  ⚠️  {grade}학년 파서가 없습니다. 건너뜁니다.")
                continue

            print(f"  📄 {grade}학년: {filepath.name}")
            students = parser(filepath)
            print(f"     → {len(students)}명 파싱 완료")
            exam_data[grade] = students
            total_students += len(students)

        result["exams"][exam_id] = {"label": info["label"], "data": exam_data}
        result["order"].append(exam_id)

    return result, total_students


def build_html(json_data, total_students):
    """템플릿에 JSON 데이터를 삽입하여 HTML 생성"""
    if not TEMPLATE_PATH.exists():
        print(f"오류: {TEMPLATE_PATH} 파일이 없습니다.")
        sys.exit(1)

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    json_str = json.dumps(json_data, ensure_ascii=False, separators=(",", ":"))

    # 학생 수 뱃지 업데이트
    num_exams = len(json_data["order"])
    num_grades = set()
    for eid in json_data["order"]:
        for g in json_data["exams"][eid]["data"]:
            num_grades.add(g)
    badge_text = f"{total_students}명 · {len(num_grades)}개 학년 · {num_exams}회 시험"
    template = template.replace("758명 · 3개 학년 · 실시간 분석", badge_text)

    html = template.replace("{{EXAM_DATA}}", json_str)

    OUTPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    print(f"\n✅ 대시보드 생성 완료: {OUTPUT_PATH}")
    print(f"   파일 크기: {len(html):,} bytes")
    print(f"   총 학생 수: {total_students}명")
    print(f"   시험 회차: {', '.join(json_data['order'])}")


def git_push(message="대시보드 업데이트"):
    """GitHub Pages에 배포"""
    try:
        os.chdir(BASE_DIR)
        subprocess.run(["git", "add", "docs/index.html"], check=True)
        subprocess.run(["git", "commit", "-m", message], check=True)
        subprocess.run(["git", "push"], check=True)
        print(f"\n🚀 GitHub Pages 배포 완료!")
        print(f"   잠시 후 사이트에 반영됩니다.")
    except subprocess.CalledProcessError as e:
        print(f"\n⚠️  Git 명령 실패: {e}")
        print("   수동으로 git add/commit/push 해주세요.")
    except FileNotFoundError:
        print("\n⚠️  git이 설치되어 있지 않습니다.")


# ──────────────────────────────────────────────
# 엔트리 포인트
# ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="양현고 모의고사 대시보드 빌드")
    parser.add_argument("--push", action="store_true", help="빌드 후 GitHub Pages 배포")
    parser.add_argument("-m", "--message", default="대시보드 업데이트", help="커밋 메시지")
    args = parser.parse_args()

    print("=" * 50)
    print("  양현고등학교 모의고사 대시보드 빌드")
    print("=" * 50)

    # 1. 데이터 폴더 스캔
    exams_info = scan_data_folder()
    if not exams_info:
        print("\n오류: data/ 폴더에 시험 데이터가 없습니다.")
        print("  data/2026.03/ 같은 폴더를 만들고 엑셀 파일을 넣어주세요.")
        sys.exit(1)

    print(f"\n📁 발견된 시험: {', '.join(exams_info.keys())}")

    # 2. 엑셀 파싱
    json_data, total = build_json(exams_info)

    # 3. HTML 생성
    build_html(json_data, total)

    # 4. 배포 (선택)
    if args.push:
        git_push(args.message)
    else:
        print(f"\n💡 GitHub에 배포하려면:")
        print(f"   python build_dashboard.py --push")


if __name__ == "__main__":
    main()
