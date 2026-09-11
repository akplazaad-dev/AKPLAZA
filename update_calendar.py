# -*- coding: utf-8 -*-
"""
================================================================================
 AK PLAZA 카드 사은혜택 자동 입력 프로그램
================================================================================

[이 프로그램이 하는 일]
  1) 이번 달 "제휴 판촉 캘린더" PDF 파일에서 카드 사은혜택 내용을 읽어옵니다.
  2) 엑셀 양식 파일의 노란색 칸(주차별 "지원 사항" 칸)에 그 내용을 채워 넣습니다.
  3) 원본은 그대로 두고, 내용이 채워진 "새 엑셀 파일"을 만들어 줍니다.
  4) 사람이 눈으로 확인하기 좋게 "검토용 텍스트 파일"도 함께 만들어 줍니다.

[사용 방법 - 아주 간단하게]
  - 이 프로그램 파일과 같은 폴더에
      · 엑셀 양식 파일 1개 (예: 사전고지_양식.xlsx)
      · 이번 달 PDF 파일 1개
    를 넣어 두고 프로그램을 실행하면 됩니다.
  - 자세한 설치/실행 방법은 함께 들어 있는 "사용설명서.md" 파일을 보세요.

  (고급) 파일 위치를 직접 지정하고 싶으면:
      python update_calendar.py  "양식.xlsx"  "이번달.pdf"

--------------------------------------------------------------------------------
 코딩을 모르셔도 됩니다. 아래 내용은 "설정" 부분만 살짝 바꾸면 되고,
 나머지는 건드리지 않아도 잘 동작합니다.
--------------------------------------------------------------------------------
"""

import os
import re
import sys
import glob

# 윈도우 한글 콘솔에서 특수문자(★, 화살표 등)를 출력해도 오류로 멈추지 않도록,
# 화면 출력 인코딩을 UTF-8로 맞추고, 표현 불가 문자는 대체하도록 설정합니다.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ------------------------------------------------------------------------------
# [설정] 여기 값만 필요할 때 바꾸면 됩니다. (보통은 그대로 두세요)
# ------------------------------------------------------------------------------

# 엑셀에서 내용을 채울 시트(탭) 이름
SHEET_NAME = "사전고지(본부)"

# 노란색 칸(주차별 "지원 사항" 칸)이 있는 셀 위치입니다.
# 왼쪽부터: 1주차, 2주차, 3주차, 4주차, 5주차
YELLOW_CELLS = ["H9", "H14", "H18", "H22", "H26"]

# 각 주차의 "기간" 정보가 들어 있는 셀 위치입니다.
# (엑셀에서 자동으로 읽어오지만, 못 읽을 경우를 대비한 위치입니다.)
PERIOD_CELLS = ["E6", "E11", "E15", "E19", "E23"]


# ==============================================================================
# 아래부터는 프로그램 본체입니다. (수정하지 않아도 됩니다)
# ==============================================================================

def 필수라이브러리_확인():
    """PDF/엑셀을 다루는 데 필요한 부품(라이브러리)이 설치돼 있는지 확인합니다."""
    부족 = []
    try:
        import pdfplumber  # noqa: F401
    except Exception:
        부족.append("pdfplumber")
    try:
        import openpyxl  # noqa: F401
    except Exception:
        부족.append("openpyxl")
    if 부족:
        print("[오류] 다음 프로그램 부품이 설치되어 있지 않습니다:", ", ".join(부족))
        print("       먼저 '최초설치.bat' 파일을 한 번 실행해 주세요.")
        print("       (또는 명령창에서:  pip install " + " ".join(부족) + " )")
        sys.exit(1)


def 공백정리(s):
    """여러 칸의 공백/줄바꿈을 한 칸 공백으로 정리합니다."""
    return re.sub(r"\s+", " ", s).strip()


def PDF에서_글자읽기(pdf_경로):
    """PDF 파일에서 모든 글자를 읽어 옵니다."""
    import pdfplumber
    모든글자 = []
    with pdfplumber.open(pdf_경로) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            모든글자.append(t)
    return "\n".join(모든글자)


def 연월_찾기(글자):
    """PDF 제목에서 '26년 10월' 같은 연/월을 찾습니다. 못 찾으면 (None, None)."""
    m = re.search(r"(\d{2})\s*년\s*(\d{1,2})\s*월", 글자)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None, None


# ------------------------------------------------------------------------------
# 월간(한 달 내내 유효한) 공통 사은혜택 추출 규칙
#   - 이 판촉 캘린더는 매달 같은 양식으로 나오기 때문에,
#     아래처럼 항목별로 "찾는 규칙"을 정해두면 매달 잘 동작합니다.
#   - 규칙은 (이름, 찾는패턴, 만드는함수) 형태입니다.
#   - 어떤 달에 특정 항목이 없으면 그 항목은 그냥 건너뜁니다(오류 없음).
# ------------------------------------------------------------------------------

def _금액정리(s):
    """'20/40/60/100/200만' 처럼 붙어 있는 금액을 보기 좋게 그대로 둡니다."""
    return s.strip()


def 월간혜택_추출(글자, 월):
    """
    PDF 글자에서 '한 달 내내 유효한' 공통 사은혜택을 뽑아
    보기 좋은 문장 목록으로 만들어 돌려줍니다.
    각 항목은 '한 덩어리 문자열'(여러 줄 포함)입니다.
    """
    # 줄바꿈을 공백으로 바꿔 한 줄로 이어 붙입니다(줄 걸침 문제 방지).
    한줄 = 공백정리(글자.replace("\n", " "))
    항목들 = []          # 최종 결과(문장 덩어리 목록)
    추출로그 = []        # 검토용: 무엇을 찾았는지 기록

    def 추가(제목, 본문줄들, 원본=""):
        덩어리 = "\n".join(본문줄들)
        항목들.append(덩어리)
        추출로그.append((제목, 원본))

    # 1) 신한Plus 발급①  (예: 전관 합산 20/40/60/100/200만 [15% 상품권1매])
    m = re.search(r"\[신한Plus 발급①\][^\[]*?(전관\s*합산\s*[\d/]+만)\s*\[([^\]]+)\]\s*\(([^)]*)\)", 한줄)
    if m:
        금액, 혜택, 기간 = _금액정리(m.group(1)), m.group(2).strip(), m.group(3).strip()
        금액 = 금액.replace("전관합산", "전관 합산 ").replace("전관 합산  ", "전관 합산 ")
        추가("신한Plus 발급①",
             ["ㆍ[신한Plus 발급①] {} ({})".format(금액, 혜택),
              "※ {}".format(기간)],
             m.group(0))

    # 2) 신한Plus 발급②  (첫결제 조건 충족 시 모바일상품권/캐시백)
    m = re.search(r"\[신한Plus 발급②\]\s*(전관합산[^\[]*?캐시백)", 한줄)
    if m:
        본문 = (m.group(1)
                .replace("전관합산", "전관 합산 ")
                .replace("첫결제조건충족시", "첫결제 조건 충족 시 ")
                .replace("/ ", "/"))
        본문 = 공백정리(본문)
        # 신규/재발급 상세 금액
        상세 = re.search(r"(\(신규발급\).*?\(익월말[^)]*\))", 한줄)
        상세줄 = 공백정리(상세.group(1)).replace("/ ", "/") if 상세 else ""
        줄들 = ["ㆍ[신한Plus 발급②] {}".format(본문)]
        if 상세줄:
            줄들.append("※ {}".format(상세줄))
        줄들.append("※ 월간, 수/분/평/원")
        추가("신한Plus 발급②", 줄들, m.group(0))

    # 3) 신한Plus 정기  (APP쿠폰 전관 단일 10만 [10%])
    m = re.search(r"\[신한Plus 정기\]\s*(APP쿠폰전관단일[\d/]+만)\s*\[([^\]]+)\]\s*\(([^)]*)\)", 한줄)
    if m:
        본문 = m.group(1).replace("APP쿠폰전관단일", "APP쿠폰 전관 단일 ")
        추가("신한Plus 정기",
             ["ㆍ[신한Plus 정기] {} ({})".format(본문, m.group(2).strip()),
              "※ {}".format(m.group(3).strip())],
             m.group(0))

    # 4) PAYCO(포인트) ①  (전관 합산, 평일/주말 상향)
    m1 = re.search(r"①\s*(전관합산[\d/]+만)\s*\[([^\]]+)\]\s*\(([^)]*)\)", 한줄)
    if m1:
        금액 = m1.group(2).strip()  # 예: 평일7%, 주말10% 상향
        기간 = m1.group(3).strip()
        본문 = m1.group(1).replace("전관합산", "전관 합산 ")
        # 대표 혜택률(예: 7~10%)을 제목에 넣고, 평일/주말은 부가설명으로
        줄들 = ["ㆍ[PAYCO(포인트)] {} (7~10%)".format(본문)]
        평주 = re.search(r"평일\s*\d+%\s*,?\s*주말\s*\d+%", 금액)
        if 평주:
            평주문 = 공백정리(평주.group(0))
            if "," not in 평주문:                      # '평일7% 주말10%' → '평일7%, 주말10%'
                평주문 = 평주문.replace(" 주말", ", 주말")
            줄들.append("※ {}".format(평주문))
        줄들.append("※ {}".format(기간))
        추가("PAYCO 포인트 ①(전관)", 줄들, m1.group(0))

    # 5) PAYCO(포인트) ②  (해외명품/골드바 단일)
    m2 = re.search(r"②\s*(해외명품/골드바단일[\d~,]+만)\s*\[([^\]]+)\]\s*\(([^)]*)\)", 한줄)
    if m2:
        본문 = m2.group(1).replace("해외명품/골드바단일", "해외명품/골드바 단일 ")
        추가("PAYCO 포인트 ②(해외명품/골드바)",
             ["ㆍ[PAYCO(포인트)] {} ({})".format(본문, m2.group(2).strip()),
              "※ {}".format(m2.group(3).strip())],
             m2.group(0))

    # 6) PAYCO(포인트) ③  (무신사/나이키 단일)
    m3 = re.search(r"③\s*(무신사/나이키단일[\d/]+만)\s*\[([^\]]+)\]", 한줄)
    if m3:
        본문 = m3.group(1).replace("무신사/나이키단일", "무신사/나이키 단일 ")
        # 이 줄은 PDF에서 옆 칸(패션그룹/KEY MD) 글자와 겹쳐 깨져 나오므로,
        # 기간은 다른 PAYCO 항목과 같은 기본값('월간, 수/분/평/원')을 사용합니다.
        기간 = "월간, 수/분/평/원"
        추가("PAYCO 포인트 ③(무신사/나이키)",
             ["ㆍ[PAYCO(포인트)] {} ({})".format(본문, m3.group(2).strip()),
              "※ {}".format(기간)],
             m3.group(0))

    # 7) 가전  ([네이버페이/NH] 가전 단일 ...)
    #    날짜(예: 10/1~31)는 '★0.5억(10/1~31, ...)'처럼 괄호 안에 들어 있어,
    #    혜택률 다음에 나오는 '날짜가 든 첫 괄호'를 찾습니다.
    m = re.search(r"\[네이버페이/NH\]\s*(가전단일[\d/,]+만)\s*\[([^\]]+)\].*?\(([^)]*\d{1,2}/\d{1,2}~[^)]*)\)", 한줄)
    if m:
        본문 = m.group(1).replace("가전단일", "가전 단일 ")
        줄들 = ["ㆍ[네이버페이/NH] {} ({})".format(본문, m.group(2).strip()),
               "※ 신한P 네이버페이 등록 시 증정률 +1%",
               "※ {}".format(공백정리(m.group(3)))]
        추가("가전", 줄들, m.group(0))

    # 8) 가구  ([네이버페이/NH] 가구 단일 ...)
    m = re.search(r"\[네이버페이/NH\]\s*(가구단일[\d/,]+만)\s*\[([^\]]+)\].*?\(([^)]*\d{1,2}/\d{1,2}~[^)]*)\)", 한줄)
    if m:
        본문 = m.group(1).replace("가구단일", "가구 단일 ")
        줄들 = ["ㆍ[네이버페이/NH] {} ({})".format(본문, m.group(2).strip()),
               "※ 신한P 네이버페이 등록 시 증정률 +1%",
               "※ 일부 참여브랜드 9%",
               "※ {}".format(공백정리(m.group(3)))]
        추가("가구", 줄들, m.group(0))

    return 항목들, 추출로그


# ------------------------------------------------------------------------------
# 특정 날짜 행사(A*CLASS, 멤페, 공판 등) 후보 줄 찾기 (검토용)
#   - 이 부분은 PDF 디자인상 글자가 날짜 칸에 겹쳐 들어가 깨져 나올 수 있어,
#     자동으로 칸에 넣지 않고 "검토용 파일"에만 정리해 드립니다.
# ------------------------------------------------------------------------------

def 날짜행사_후보(글자):
    """날짜 구간(예: 10/16~18)이나 혜택률(%)이 들어 있는 줄을 골라 돌려줍니다."""
    후보 = []
    날짜패턴 = re.compile(r"\d{1,2}/\d{1,2}\s*~\s*\d{1,2}(?:/\d{1,2})?")
    for line in 글자.split("\n"):
        line = line.rstrip()
        if not line.strip():
            continue
        날짜들 = 날짜패턴.findall(line)
        if 날짜들:
            후보.append((line.strip(), 날짜들))
    return 후보


# ------------------------------------------------------------------------------
# 엑셀 쓰기
# ------------------------------------------------------------------------------

def 주차기간_읽기(ws):
    """엑셀 양식에서 각 주차의 '기간' 글자를 읽어옵니다(검토용/표시용)."""
    기간목록 = []
    for cell in PERIOD_CELLS:
        try:
            v = ws[cell].value
        except Exception:
            v = None
        기간목록.append(공백정리(str(v)) if v else "")
    return 기간목록


def 엑셀에_쓰기(양식_경로, 결과_경로, 셀내용):
    """
    양식 엑셀을 열어 노란색 칸에 내용을 채우고, 새 파일로 저장합니다.
    - 셀내용: {"H9": "...", "H14": "...", ...} 형태
    - 서식(줄바꿈 자동, 노란색 등)은 원래 양식의 것을 그대로 유지합니다.
    """
    import openpyxl
    from openpyxl.styles import Alignment

    wb = openpyxl.load_workbook(양식_경로)
    if SHEET_NAME in wb.sheetnames:
        ws = wb[SHEET_NAME]
    else:
        ws = wb.active
        print("[안내] '{}' 시트를 못 찾아 첫 번째 시트에 씁니다.".format(SHEET_NAME))

    for 셀, 내용 in 셀내용.items():
        칸 = ws[셀]
        칸.value = 내용
        # 줄바꿈이 보이도록 '자동 줄바꿈'을 켜고, 위-왼쪽 정렬로 맞춥니다.
        기존 = 칸.alignment
        칸.alignment = Alignment(
            wrap_text=True,
            vertical=(기존.vertical or "top"),
            horizontal=(기존.horizontal or "left"),
        )

    wb.save(결과_경로)


# ------------------------------------------------------------------------------
# 파일 자동 찾기
# ------------------------------------------------------------------------------

def 파일_자동찾기(폴더):
    """폴더 안에서 엑셀 양식 1개와 PDF 1개를 자동으로 찾습니다."""
    pdf목록 = sorted(glob.glob(os.path.join(폴더, "*.pdf")),
                    key=os.path.getmtime, reverse=True)
    xlsx목록 = [f for f in glob.glob(os.path.join(폴더, "*.xlsx"))
               if "결과" not in os.path.basename(f)
               and not os.path.basename(f).startswith("~$")]
    xlsx목록 = sorted(xlsx목록, key=os.path.getmtime, reverse=True)
    pdf = pdf목록[0] if pdf목록 else None
    xlsx = xlsx목록[0] if xlsx목록 else None
    return xlsx, pdf


# ------------------------------------------------------------------------------
# 메인 (프로그램 시작점)
# ------------------------------------------------------------------------------

def main():
    print("=" * 60)
    print(" AK PLAZA 카드 사은혜택 자동 입력 프로그램")
    print("=" * 60)

    필수라이브러리_확인()

    # 이 프로그램 파일이 있는 폴더
    기준폴더 = os.path.dirname(os.path.abspath(__file__))

    # 1) 파일 위치 정하기 (직접 지정 > 자동 찾기)
    양식_경로 = None
    pdf_경로 = None
    인자 = sys.argv[1:]
    for a in 인자:
        low = a.lower()
        if low.endswith(".xlsx"):
            양식_경로 = a
        elif low.endswith(".pdf"):
            pdf_경로 = a
    if not 양식_경로 or not pdf_경로:
        자동_xlsx, 자동_pdf = 파일_자동찾기(기준폴더)
        양식_경로 = 양식_경로 or 자동_xlsx
        pdf_경로 = pdf_경로 or 자동_pdf

    if not 양식_경로 or not os.path.exists(양식_경로):
        print("\n[오류] 엑셀 양식 파일(.xlsx)을 찾지 못했습니다.")
        print("       이 프로그램과 같은 폴더에 엑셀 양식 파일을 넣어 주세요.")
        sys.exit(1)
    if not pdf_경로 or not os.path.exists(pdf_경로):
        print("\n[오류] PDF 파일(.pdf)을 찾지 못했습니다.")
        print("       이 프로그램과 같은 폴더에 이번 달 PDF 파일을 넣어 주세요.")
        sys.exit(1)

    print("\n[사용할 파일]")
    print("  · 엑셀 양식 :", os.path.basename(양식_경로))
    print("  · PDF 파일  :", os.path.basename(pdf_경로))

    # 2) PDF 읽기
    글자 = PDF에서_글자읽기(pdf_경로)
    연, 월 = 연월_찾기(글자)
    월표시 = "{}월".format(월) if 월 else "이번달"
    연표시 = "20{}년".format(연) if 연 else ""
    print("\n[문서 인식]", 연표시, 월표시)

    # 3) 월간 공통 혜택 추출
    항목들, 추출로그 = 월간혜택_추출(글자, 월)
    if not 항목들:
        print("\n[주의] PDF에서 월간 공통 혜택을 하나도 찾지 못했습니다.")
        print("       PDF 양식이 평소와 다를 수 있습니다. 검토용 파일을 확인해 주세요.")
    else:
        print("\n[찾은 월간 공통 혜택] 총 {}건".format(len(추출로그)))
        for 제목, _ in 추출로그:
            print("   - ", 제목)

    # 4) 주차 칸에 넣을 내용 만들기
    #    - 월간 공통 혜택은 '모든 주차'에 유효하므로 5칸 모두에 넣습니다.
    셀본문 = "\n\n".join(항목들) if 항목들 else ""
    셀내용 = {cell: 셀본문 for cell in YELLOW_CELLS}

    # 5) 엑셀 저장 (새 파일)
    양식이름 = os.path.splitext(os.path.basename(양식_경로))[0]
    꼬리 = "{}{}".format(연표시.replace("년", "년"), 월표시) if 연 else 월표시
    결과이름 = "{}_{}_결과.xlsx".format(양식이름, 꼬리.replace(" ", ""))
    결과_경로 = os.path.join(기준폴더, 결과이름)
    엑셀에_쓰기(양식_경로, 결과_경로, 셀내용)
    print("\n[완료] 새 엑셀 파일을 만들었습니다:")
    print("   ->", 결과이름)

    # 6) 검토용 텍스트 파일 만들기
    검토이름 = "{}_{}_검토용.txt".format(양식이름, 꼬리.replace(" ", ""))
    검토_경로 = os.path.join(기준폴더, 검토이름)
    주차기간 = None
    try:
        import openpyxl
        wb = openpyxl.load_workbook(양식_경로, data_only=True)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
        주차기간 = 주차기간_읽기(ws)
    except Exception:
        주차기간 = ["" for _ in YELLOW_CELLS]

    날짜후보 = 날짜행사_후보(글자)
    with open(검토_경로, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\n")
        f.write(" {} {} 카드 사은혜택 - 검토용 정리\n".format(연표시, 월표시))
        f.write("=" * 70 + "\n\n")
        f.write("이 파일은 '사람이 눈으로 확인'하기 위한 참고용입니다.\n")
        f.write("아래 [월간 공통 혜택]은 엑셀 결과 파일의 모든 주차 칸에 이미 들어갔습니다.\n\n")

        f.write("-" * 70 + "\n")
        f.write("[1] 엑셀 노란색 칸에 들어간 내용 (월간 공통 혜택)\n")
        f.write("-" * 70 + "\n")
        for i, cell in enumerate(YELLOW_CELLS):
            기간표시 = ("  ({})".format(주차기간[i]) if 주차기간 and 주차기간[i] else "")
            f.write("\n■ {}주차{}  [엑셀 {}]\n".format(i + 1, 기간표시, cell))
            f.write((셀본문 if 셀본문 else "(내용 없음)") + "\n")

        f.write("\n\n")
        f.write("-" * 70 + "\n")
        f.write("[2] 특정 날짜 행사 후보 (사람이 확인 후 필요하면 직접 추가하세요)\n")
        f.write("-" * 70 + "\n")
        f.write("※ 아래는 PDF에서 날짜 구간(예: 10/16~18)이 있는 줄을 그대로 뽑은 것입니다.\n")
        f.write("※ PDF 디자인상 글자가 섞여 나올 수 있으니, 원본 PDF와 비교해 확인하세요.\n\n")
        if 날짜후보:
            for 줄, 날짜들 in 날짜후보:
                f.write("· (날짜: {})\n   {}\n\n".format(", ".join(날짜들), 줄))
        else:
            f.write("(특정 날짜 행사 후보를 찾지 못했습니다.)\n")

    print("[완료] 검토용 파일도 만들었습니다:")
    print("   ->", 검토이름)

    print("\n" + "=" * 60)
    print(" 끝났습니다! 결과 엑셀 파일을 열어 확인해 주세요.")
    print(" (특정 날짜 행사는 검토용 파일을 보고 필요하면 직접 추가하세요.)")
    print("=" * 60)


if __name__ == "__main__":
    # 오류가 나더라도 무슨 일인지 화면에 보여 줍니다.
    # (창을 열어두고 멈추는 일은 '실행하기.bat'이 담당합니다.)
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        print("\n" + "!" * 60)
        print("[예상치 못한 오류가 발생했습니다]")
        print("  ", repr(e))
        print("  아래 내용을 캡처해 담당자에게 보여 주시면 도움이 됩니다.")
        print("!" * 60)
        import traceback
        traceback.print_exc()
        sys.exit(1)
