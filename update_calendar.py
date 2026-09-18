# -*- coding: utf-8 -*-
"""
================================================================================
 AK PLAZA 카드 사은혜택 자동 입력 프로그램  (PPT / PDF 지원)
================================================================================

[이 프로그램이 하는 일]
  1) 이번 달 "제휴 판촉 캘린더" 파일(PPT 권장, PDF도 가능)에서
     카드 사은혜택 내용을 읽어옵니다.
  2) 엑셀 양식의 노란색 칸(주차별 "지원 사항" 칸)에 그 내용을 채워 넣습니다.
     - 월간(한 달 내내) 공통 혜택 → 5개 주차 칸 모두에
     - 특정 날짜 행사 → 날짜가 속한 그 주차 칸에만
  3) 원본은 그대로 두고, 내용이 채워진 "새 엑셀 파일"을 만들어 줍니다.
  4) 사람이 눈으로 확인하기 좋게 "검토용 텍스트 파일"도 함께 만들어 줍니다.

[왜 PPT를 권장하나요?]
  PDF는 글자를 화면 위치 순서로 뽑아, 이미지에 겹친 행사 글자가 뒤섞일 수
  있습니다. PPT는 글상자마다 글자가 온전히 남아 있어, 특정 날짜 행사까지
  정확하게 읽어 각 주차에 배치할 수 있습니다.

[사용 방법 - 아주 간단하게]
  - 이 프로그램 파일과 같은 폴더에
      · 엑셀 양식 파일 1개 (예: 사전고지_양식.xlsx)
      · 이번 달 PPT 파일 1개 (없으면 PDF 파일)
    를 넣어 두고 "실행하기.bat" 을 더블클릭하면 됩니다.
  - 자세한 설치/실행 방법은 함께 들어 있는 "사용설명서.md" 파일을 보세요.

  (고급) 파일을 직접 지정하려면:
      python update_calendar.py  "양식.xlsx"  "이번달.pptx"
--------------------------------------------------------------------------------
"""

import os
import re
import sys
import glob

# 윈도우 한글 콘솔에서 특수문자를 출력해도 오류로 멈추지 않도록 설정합니다.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ------------------------------------------------------------------------------
# [설정] 여기 값만 필요할 때 바꾸면 됩니다. (보통은 그대로 두세요)
# ------------------------------------------------------------------------------

SHEET_NAME = "사전고지(본부)"                      # 내용을 채울 시트(탭) 이름
YELLOW_CELLS = ["H9", "H14", "H18", "H22", "H26"]  # 1~5주차 노란색 칸 위치
PERIOD_CELLS = ["E6", "E11", "E15", "E19", "E23"]  # 1~5주차 '기간' 칸 위치

# 특정 날짜 행사에서 카드사로 인정할 낱말들
CARD_KW = ("신한", "네이버", "KB", "NH", "PAYCO", "페이", "일반", "BC", "카카오")


# ==============================================================================
# 공통 도우미 함수
# ==============================================================================

def sp(s):
    """여러 칸의 공백을 한 칸으로 정리합니다."""
    return re.sub(r"[ \t]+", " ", str(s)).strip()


def md(m, d):
    """월/일을 비교하기 쉬운 숫자로: 10월 5일 → 1005"""
    return m * 100 + d


def 날짜구간_파싱(txt):
    """'10/2~5', '10/16~22', '10/30~11/1', '10/1~31' → (시작, 끝) 숫자. 없으면 None."""
    m = re.search(r"(\d{1,2})/(\d{1,2})\s*~\s*(?:(\d{1,2})/)?(\d{1,2})", txt)
    if not m:
        return None
    m1, d1 = int(m.group(1)), int(m.group(2))
    m2 = int(m.group(3)) if m.group(3) else m1
    d2 = int(m.group(4))
    return (md(m1, d1), md(m2, d2))


def 기간정리(paren):
    """'월간, 수/분/평/원, 점-페이각 50%' → '월간, 수/분/평/원' 처럼 앞부분만 남깁니다."""
    inner = sp(paren)
    m = re.match(r"(월간|\d{1,2}/\d{1,2}[~\d/]*)\s*,\s*([수분평원광/]+)", inner)
    return "{}, {}".format(m.group(1), m.group(2)) if m else inner


def 카드사_정리(s):
    """'신한 Plus 발급 ①' → '신한Plus 발급①' 처럼 보기 좋게 다듬습니다."""
    return (sp(s).replace("신한 Plus", "신한Plus")
                 .replace("발급 ①", "발급①").replace("발급 ②", "발급②"))


def 점포정리(s):
    """'4개점' 같은 표현을 실제 점포 표기 '수/분/평/원'으로 바꿉니다. ('분' 등 단일점포는 유지)"""
    return re.sub(r"\d+\s*개점", "수/분/평/원", str(s))


def 예_아니오(질문, 기본=True):
    """사용자에게 예/아니오를 물어봅니다. 그냥 Enter를 누르면 '기본'값을 씁니다.
    (자동 실행 등으로 입력을 받을 수 없으면 기본값을 사용합니다.)"""
    표시 = "Y/n" if 기본 else "y/N"
    try:
        답 = input("{} ({}): ".format(질문, 표시)).strip().lower()
    except Exception:
        return 기본
    if 답 == "":
        return 기본
    if 답 in ("y", "yes", "네", "예", "o", "ㅛ"):
        return True
    if 답 in ("n", "no", "아니오", "아니요", "x", "ㅜ"):
        return False
    return 기본


def 기본율_노트(t1):
    """가전/가구 글상자에서 '기본 9%'와 '등록 신한P 10%'를 찾아 ※ 문구를 만듭니다."""
    기본 = re.search(r"기본\s*(\d+)\s*%", t1)
    등록 = re.search(r"등록\s*신한\s*P?\s*(\d+)\s*%", t1)
    if 기본 and 등록:
        return "※ 기본 {}%, 신한P 네이버페이 등록 시 {}%".format(기본.group(1), 등록.group(1))
    if 기본:
        return "※ 기본 {}%, 신한P 네이버페이 등록 시 증정률 +1%".format(기본.group(1))
    return "※ 신한P 네이버페이 등록 시 증정률 +1%"


# ==============================================================================
# 주차 범위 읽기 (엑셀의 '기간' 칸에서)
# ==============================================================================

def 주차범위_읽기(ws, 월):
    """
    엑셀 E칸('10월 1주차 10/1~10/7')에서 주차별 (시작,끝) 날짜를 읽어옵니다.
    못 읽으면 그 달 기준의 기본값으로 대체합니다.
    """
    구간 = []
    for cell in PERIOD_CELLS:
        try:
            v = str(ws[cell].value or "")
        except Exception:
            v = ""
        m = re.search(r"(\d{1,2})/(\d{1,2})\s*~\s*(?:(\d{1,2})/)?(\d{1,2})", v)
        if m:
            m1, d1 = int(m.group(1)), int(m.group(2))
            m2 = int(m.group(3)) if m.group(3) else m1
            d2 = int(m.group(4))
            구간.append((md(m1, d1), md(m2, d2)))
        else:
            구간.append(None)
    # 하나라도 못 읽었으면, 못 읽은 칸만 넉넉한 기본값으로 채웁니다.
    for i, g in enumerate(구간):
        if g is None:
            구간[i] = (md(월 or 1, 1), md((월 or 1) + 1, 5))
    return 구간


def 주차찾기(rng, 주차구간):
    """날짜 구간이 겹치는 주차 번호(0~4) 목록. 날짜가 없으면 모든 주차."""
    if rng is None:
        return list(range(len(주차구간)))
    s, e = rng
    return [i for i, g in enumerate(주차구간) if g and s <= g[1] and e >= g[0]]


# ==============================================================================
# PPT(.pptx) 읽기
# ==============================================================================

def PPT_도형텍스트(pptx_경로):
    """PPT의 모든 글상자 텍스트를 목록으로 돌려줍니다(그룹 안까지)."""
    from pptx import Presentation

    def 안쪽텍스트(shape):
        if shape.shape_type == 6:  # GROUP
            return "\n".join(안쪽텍스트(s) for s in shape.shapes).strip()
        return shape.text_frame.text if shape.has_text_frame else ""

    prs = Presentation(pptx_경로)
    결과 = []
    for slide in prs.slides:
        for s in slide.shapes:
            if getattr(s, "has_table", False):
                continue
            t = 안쪽텍스트(s)
            if t and t.strip():
                결과.append(t)
    return 결과


def PPT_월간혜택(도형들):
    """
    PPT에서 '한 달 내내 유효한' 공통 혜택을 뽑습니다.
    돌려주는 값: [(제목, 문장덩어리), ...]  (제목으로 발급/정기 포함여부를 걸러낼 수 있음)
    """
    항목 = []      # (정렬순서, 제목, 덩어리)

    def 넣기(순서, 제목, 줄들):
        항목.append((순서, 제목, "\n".join(줄들)))

    for t in 도형들:
        L0 = sp(t.split("\n")[0])
        t1 = sp(t.replace("\n", " "))
        c = L0.replace(" ", "")   # 공백 제거본(판별용)

        if "발급①" in c and c.startswith("[신한"):
            m = re.match(r"\[([^\]]+)\]\s*(.+?)\s*\[([^\]]+)\]\s*\(([^)]+)\)", L0)
            if m:
                넣기(1, "신한Plus 발급①",
                    ["ㆍ[{}] {} ({})".format(카드사_정리(m.group(1)), sp(m.group(2)), sp(m.group(3))),
                     "※ {}".format(기간정리(m.group(4)))])

        elif "발급②" in c and c.startswith("[신한"):
            m = re.match(r"\[([^\]]+)\]\s*(.+?)\s*\(([^)]+)\)", L0)
            if m:
                본문 = sp(m.group(2)).replace(" / ", "/")
                줄들 = ["ㆍ[{}] {}".format(카드사_정리(m.group(1)), 본문)]
                상세 = re.search(r"(\(신규 ?발급\).*?증정\s*\))", t1)
                if 상세:
                    줄들.append("※ {}".format(sp(상세.group(1)).replace(" / ", "/")))
                줄들.append("※ {}".format(기간정리(m.group(3))))
                넣기(2, "신한Plus 발급②", 줄들)

        elif "정기]" in c and c.startswith("[신한"):
            m = re.match(r"\[([^\]]+)\]\s*(.+?)\s*\[([^\]]+)\]\s*\(([^)]+)\)", L0)
            if m:
                넣기(3, "신한Plus 정기",
                    ["ㆍ[{}] {} ({})".format(카드사_정리(m.group(1)), sp(m.group(2)), sp(m.group(3))),
                     "※ {}".format(기간정리(m.group(4)))])

        elif c.startswith("[PAYCO"):
            순번 = 4
            for mm in re.finditer(r"[①②③]\s*(.+?)\s*\[([^\]]+)\]\s*\(([^)]+)\)", t1):
                본문, 혜택, 기간 = sp(mm.group(1)), sp(mm.group(2)), mm.group(3)
                if "평일" in 혜택:
                    줄들 = ["ㆍ[PAYCO(포인트)] {} (7~10%)".format(본문)]
                    pr = re.search(r"평일\s*\d+%\s*,?\s*주말\s*\d+%", 혜택)
                    if pr:
                        줄들.append("※ {}".format(sp(pr.group(0))))
                    줄들.append("※ {}".format(기간정리(기간)))
                else:
                    줄들 = ["ㆍ[PAYCO(포인트)] {} ({})".format(본문, 혜택),
                           "※ {}".format(기간정리(기간))]
                넣기(순번, "PAYCO 포인트", 줄들)
                순번 += 0.1

        elif "가전단일" in c:
            m = re.match(r"\[([^\]]+)\]\s*(가전.+?)\s*\[([^\]]+)\].*?\(([^)]*\d/\d[^)]*)\)", L0)
            if m:
                줄들 = ["ㆍ[{}] {} ({})".format(sp(m.group(1)), sp(m.group(2)), sp(m.group(3))),
                       기본율_노트(t1),
                       "※ {}".format(기간정리(m.group(4)))]
                넣기(7, "가전", [x for x in 줄들 if x])

        elif "가구단일" in c:
            m = re.match(r"\[([^\]]+)\]\s*(가구.+?)\s*\[([^\]]+)\].*?\(([^)]*\d/\d[^)]*)\)", L0)
            if m:
                줄들 = ["ㆍ[{}] {} ({})".format(sp(m.group(1)), sp(m.group(2)), sp(m.group(3))),
                       기본율_노트(t1),
                       "※ 일부 참여브랜드 9%",
                       "※ {}".format(기간정리(m.group(4)))]
                넣기(8, "가구", [x for x in 줄들 if x])

    항목.sort(key=lambda x: x[0])
    return [(제목, c) for _, 제목, c in 항목]


def PPT_날짜행사(도형들, 주차구간):
    """
    PPT에서 '특정 날짜 행사'를 뽑아, 각 행사가 속한 주차에 배치합니다.
    돌려주는 값: 주차별 목록 [[1주차행사들], [2주차...], ...]
    """
    주차별 = [[] for _ in 주차구간]
    로그 = []

    for t in 도형들:
        L0 = sp(t.split("\n")[0])
        if ("☆" not in t) and (not re.match(r"[②③]\s*\[", L0)):
            continue

        t1 = sp(t.replace("\n", " "))
        t1 = re.sub(r"★\s*[^ ]*?억", "", t1)                       # ★0.6억 등 예산 제거
        t1 = re.sub(r"★\s*신한[^,]*?/\s*네이버[^ ]*억", "", t1)    # A*CLASS식 예산 제거

        nm = re.search(r"☆\s*(.+?)\s*★", t)
        행사명 = sp(nm.group(1)) if nm else ("멤버스 페스티벌" if "멤페" in t else "행사")

        parens = list(re.finditer(r"\((\d{1,2}/\d{1,2}[^)]*)\)", t1))
        cur = 0
        for pm in parens:
            seg = t1[cur:pm.start()]
            cur = pm.end()
            날짜점포 = 점포정리(sp(pm.group(1)))   # '4개점' → '수/분/평/원'

            대상들 = re.findall(
                r"([가-힣A-Za-z*]+(?:\s[가-힣A-Za-z*]+)*\s*(?:합산|단일)\s*[\d/,~]+\s*만?)", seg)
            if not 대상들:
                continue
            본문 = sp(대상들[-1])

            혜택들 = re.findall(r"\[([^\]]*\d+\s*%[^\]]*)\]", seg)
            카드들 = [b for b in re.findall(r"\[([^\]]+)\]", seg)
                     if any(k in b for k in CARD_KW)]
            카드사 = sp(카드들[-1]) if 카드들 else ""

            혜택률 = sp(혜택들[-1]) if 혜택들 else ""
            if not 혜택률 and 카드사 and "%" in 카드사:
                혜택률 = 카드사
            # 혜택률 문자열에 카드/조건 글자가 섞이면 숫자%만 혜택률로, 나머지는 카드사로
            if 혜택률 and any(k in 혜택률 for k in CARD_KW):
                rm = re.search(r"\d+\s*%", 혜택률)
                카드사 = sp(re.sub(r"\d+\s*%", "", 혜택률)).rstrip("시 ").strip()
                혜택률 = sp(rm.group(0)) if rm else ""

            머리 = "ㆍ[{}] ".format(행사명)
            if 카드사:
                머리 += "[{}] ".format(카드사)
            머리 += 본문 + ((" ({})".format(혜택률)) if 혜택률 else "")

            줄들 = [머리]
            노트 = re.search(r"신한P\s*\d+%\s*,\s*네이버페이\s*\d+%", seg)
            if 노트:
                줄들.append("※ {}".format(sp(노트.group(0))))
            줄들.append("※ {}".format(날짜점포))

            rng = 날짜구간_파싱(날짜점포)
            덩어리 = "\n".join(줄들)
            로그.append((행사명, 날짜점포))
            for w in 주차찾기(rng, 주차구간):
                주차별[w].append((rng[0] if rng else 0, 덩어리))

    # 각 주차 안에서 날짜 순으로 정렬
    for w in range(len(주차별)):
        주차별[w].sort(key=lambda x: x[0])
        주차별[w] = [c for _, c in 주차별[w]]
    return 주차별, 로그


# ==============================================================================
# PDF(.pdf) 읽기 (PPT가 없을 때의 대비책 - 월간 공통 혜택만 추출)
# ==============================================================================

def PDF_글자(pdf_경로):
    import pdfplumber
    글자 = []
    with pdfplumber.open(pdf_경로) as pdf:
        for page in pdf.pages:
            글자.append(page.extract_text() or "")
    return "\n".join(글자)


def PDF_월간혜택(글자):
    """
    PDF에서 월간 공통 혜택만 추출합니다(줄 겹침 때문에 날짜 행사는 검토용으로만).
    돌려주는 값: [(제목, 문장덩어리), ...]
    """
    한줄 = sp(글자.replace("\n", " "))
    항목 = []

    def 넣기(제목, 줄들):
        항목.append((제목, "\n".join(줄들)))

    m = re.search(r"\[신한Plus 발급①\][^\[]*?(전관\s*합산\s*[\d/]+만)\s*\[([^\]]+)\]\s*\(([^)]*)\)", 한줄)
    if m:
        넣기("신한Plus 발급①", ["ㆍ[신한Plus 발급①] {} ({})".format(sp(m.group(1)), sp(m.group(2))),
                              "※ {}".format(sp(m.group(3)))])
    m = re.search(r"\[신한Plus 발급②\]\s*(전관합산[^\[]*?캐시백)", 한줄)
    if m:
        본문 = sp(m.group(1)).replace("첫결제조건충족시", "첫결제 조건 충족 시 ").replace("/ ", "/")
        상세 = re.search(r"(\(신규발급\).*?\(익월말[^)]*\))", 한줄)
        줄들 = ["ㆍ[신한Plus 발급②] {}".format(sp(본문))]
        if 상세:
            줄들.append("※ {}".format(sp(상세.group(1)).replace("/ ", "/")))
        줄들.append("※ 월간, 수/분/평/원")
        넣기("신한Plus 발급②", 줄들)
    m = re.search(r"\[신한Plus 정기\]\s*(APP쿠폰전관단일[\d/]+만)\s*\[([^\]]+)\]\s*\(([^)]*)\)", 한줄)
    if m:
        넣기("신한Plus 정기", ["ㆍ[신한Plus 정기] {} ({})".format(sp(m.group(1)), sp(m.group(2))),
                            "※ {}".format(sp(m.group(3)))])
    m = re.search(r"①\s*(전관합산[\d/]+만)\s*\[([^\]]+)\]\s*\(([^)]*)\)", 한줄)
    if m:
        줄들 = ["ㆍ[PAYCO(포인트)] {} (7~10%)".format(sp(m.group(1)))]
        pr = re.search(r"평일\s*\d+%\s*,?\s*주말\s*\d+%", m.group(2))
        if pr:
            줄들.append("※ {}".format(sp(pr.group(0))))
        줄들.append("※ {}".format(sp(m.group(3))))
        넣기("PAYCO 포인트 ①", 줄들)
    m = re.search(r"②\s*(해외명품/골드바단일[\d~,]+만)\s*\[([^\]]+)\]\s*\(([^)]*)\)", 한줄)
    if m:
        넣기("PAYCO 포인트 ②", ["ㆍ[PAYCO(포인트)] {} ({})".format(sp(m.group(1)), sp(m.group(2))),
                             "※ {}".format(sp(m.group(3)))])
    m = re.search(r"③\s*(무신사/나이키단일[\d/]+만)\s*\[([^\]]+)\]", 한줄)
    if m:
        넣기("PAYCO 포인트 ③", ["ㆍ[PAYCO(포인트)] {} ({})".format(sp(m.group(1)), sp(m.group(2))),
                             "※ 월간, 수/분/평/원"])
    m = re.search(r"\[네이버페이/NH\]\s*(가전단일[\d/,]+만)\s*\[([^\]]+)\].*?\(([^)]*\d{1,2}/\d{1,2}~[^)]*)\)", 한줄)
    if m:
        넣기("가전", ["ㆍ[네이버페이/NH] {} ({})".format(sp(m.group(1)), sp(m.group(2))),
                    기본율_노트(한줄), "※ {}".format(sp(m.group(3)))])
    m = re.search(r"\[네이버페이/NH\]\s*(가구단일[\d/,]+만)\s*\[([^\]]+)\].*?\(([^)]*\d{1,2}/\d{1,2}~[^)]*)\)", 한줄)
    if m:
        넣기("가구", ["ㆍ[네이버페이/NH] {} ({})".format(sp(m.group(1)), sp(m.group(2))),
                    기본율_노트(한줄), "※ 일부 참여브랜드 9%", "※ {}".format(sp(m.group(3)))])
    return 항목


def PDF_날짜행사후보(글자):
    """PDF에서 날짜 구간이 있는 줄을 그대로 뽑습니다(검토용)."""
    후보 = []
    pat = re.compile(r"\d{1,2}/\d{1,2}\s*~\s*\d{1,2}(?:/\d{1,2})?")
    for line in 글자.split("\n"):
        line = line.strip()
        if line and pat.findall(line):
            후보.append((line, pat.findall(line)))
    return 후보


# ==============================================================================
# 연/월 찾기, 엑셀 쓰기, 파일 찾기
# ==============================================================================

def 연월_찾기(글자):
    m = re.search(r"(\d{2})\s*년\s*(\d{1,2})\s*월", 글자)
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)


def 엑셀에_쓰기(양식_경로, 결과_경로, 셀내용):
    import openpyxl
    from openpyxl.styles import Alignment
    wb = openpyxl.load_workbook(양식_경로)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    for 셀, 내용 in 셀내용.items():
        칸 = ws[셀]
        칸.value = 내용
        기존 = 칸.alignment
        칸.alignment = Alignment(wrap_text=True,
                                 vertical=(기존.vertical or "top"),
                                 horizontal=(기존.horizontal or "left"))
    wb.save(결과_경로)


def 파일_자동찾기(폴더):
    def 최신(패턴):
        L = [f for f in glob.glob(os.path.join(폴더, 패턴))
             if not os.path.basename(f).startswith("~$")]
        return sorted(L, key=os.path.getmtime, reverse=True)
    pptx = 최신("*.pptx")
    pdf = 최신("*.pdf")
    xlsx = [f for f in 최신("*.xlsx") if "결과" not in os.path.basename(f)]
    return (xlsx[0] if xlsx else None,
            pptx[0] if pptx else None,
            pdf[0] if pdf else None)


# ==============================================================================
# 메인
# ==============================================================================

def 입력종류_판단(경로):
    """파일 확장자로 'ppt' / 'pdf' / None 을 돌려줍니다."""
    low = (경로 or "").lower()
    if low.endswith(".pptx"):
        return "ppt"
    if low.endswith(".pdf"):
        return "pdf"
    return None


def 라이브러리_확인(입력종류):
    """필요한 부품이 설치돼 있는지 확인하고, 없는 것의 이름 목록을 돌려줍니다."""
    필요 = ["openpyxl"] + (["pptx"] if 입력종류 == "ppt" else ["pdfplumber"])
    없음 = []
    for lib in 필요:
        try:
            __import__(lib)
        except Exception:
            없음.append("python-pptx" if lib == "pptx" else lib)
    return 없음


def 혜택_처리(양식_경로, 입력_경로, 입력종류, 발급포함, 정기포함, 기준폴더=None, 로그=print):
    """
    핵심 처리: 입력 파일에서 혜택을 뽑아 엑셀 노란색 칸을 채우고,
    결과 엑셀 + 검토용 txt 를 만들어 저장합니다.
    돌려주는 값(dict): 결과경로, 검토경로, 연, 월, 월간제목들, 주차행사건수
    (CLI와 GUI가 함께 사용합니다. 로그()로 진행 상황을 알려줍니다.)
    """
    import openpyxl
    if not 기준폴더:
        기준폴더 = os.path.dirname(os.path.abspath(입력_경로))

    로그("[사용할 파일]")
    로그("  · 엑셀 양식 : " + os.path.basename(양식_경로))
    로그("  · 입력 파일 : {} ({})".format(os.path.basename(입력_경로), 입력종류.upper()))

    wb0 = openpyxl.load_workbook(양식_경로, data_only=True)
    ws0 = wb0[SHEET_NAME] if SHEET_NAME in wb0.sheetnames else wb0.active

    # 입력 파일에서 혜택 추출
    if 입력종류 == "ppt":
        도형들 = PPT_도형텍스트(입력_경로)
        전체글자 = "\n".join(도형들)
        연, 월 = 연월_찾기(전체글자)
        주차구간 = 주차범위_읽기(ws0, 월)
        월간항목 = PPT_월간혜택(도형들)
        주차행사, 행사로그 = PPT_날짜행사(도형들, 주차구간)
        날짜후보 = None
    else:
        전체글자 = PDF_글자(입력_경로)
        연, 월 = 연월_찾기(전체글자)
        주차구간 = 주차범위_읽기(ws0, 월)
        월간항목 = PDF_월간혜택(전체글자)
        주차행사 = [[] for _ in YELLOW_CELLS]
        행사로그 = []
        날짜후보 = PDF_날짜행사후보(전체글자)

    # 발급/정기 포함 여부 반영
    def _포함(제목):
        if not 발급포함 and "발급" in 제목:
            return False
        if not 정기포함 and "정기" in 제목:
            return False
        return True
    월간항목 = [(t, c) for (t, c) in 월간항목 if _포함(t)]

    월표시 = "{}월".format(월) if 월 else "이번달"
    연표시 = "20{}년".format(연) if 연 else ""
    로그("\n[문서 인식] {} {}".format(연표시, 월표시))
    로그("[월간 공통 혜택] 총 {}건".format(len(월간항목)))
    for 제목, _ in 월간항목:
        로그("   - " + 제목)
    if not 발급포함:
        로그("   (신한Plus 발급 혜택은 제외)")
    if not 정기포함:
        로그("   (신한Plus 정기 혜택은 제외)")
    if 입력종류 == "ppt":
        로그("[특정 날짜 행사] 총 {}건 (주차별 배치)".format(len(행사로그)))

    # 주차 칸 내용 만들기
    셀내용 = {}
    for i, cell in enumerate(YELLOW_CELLS):
        조각 = [c for _, c in 월간항목]
        if i < len(주차행사) and 주차행사[i]:
            조각.append("─ [이 주 특별 행사] ─")
            조각.extend(주차행사[i])
        셀내용[cell] = "\n\n".join(조각)

    # 엑셀 저장 (새 파일)
    양식이름 = os.path.splitext(os.path.basename(양식_경로))[0]
    꼬리 = "{}{}".format(연표시, 월표시).replace(" ", "") if 연 else 월표시
    결과_경로 = os.path.join(기준폴더, "{}_{}_결과.xlsx".format(양식이름, 꼬리))
    엑셀에_쓰기(양식_경로, 결과_경로, 셀내용)
    로그("\n[완료] 새 엑셀 파일: " + os.path.basename(결과_경로))

    # 검토용 텍스트 파일
    검토_경로 = os.path.join(기준폴더, "{}_{}_검토용.txt".format(양식이름, 꼬리))
    주차기간표시 = []
    for cell in PERIOD_CELLS:
        try:
            주차기간표시.append(sp(str(ws0[cell].value or "")))
        except Exception:
            주차기간표시.append("")
    with open(검토_경로, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\n")
        f.write(" {} {} 카드 사은혜택 - 검토용 정리\n".format(연표시, 월표시))
        f.write("=" * 70 + "\n\n")
        f.write("아래 내용이 엑셀 결과 파일의 각 주차 노란색 칸에 들어갔습니다.\n")
        f.write("원본과 비교해 확인해 주세요.\n\n")
        for i, cell in enumerate(YELLOW_CELLS):
            기간 = ("  ({})".format(주차기간표시[i]) if i < len(주차기간표시) and 주차기간표시[i] else "")
            f.write("\n" + "─" * 70 + "\n")
            f.write("■ {}주차{}  [엑셀 {}]\n".format(i + 1, 기간, cell))
            f.write("─" * 70 + "\n")
            f.write((셀내용[cell] if 셀내용[cell] else "(내용 없음)") + "\n")
        if 입력종류 == "pdf" and 날짜후보:
            f.write("\n\n" + "=" * 70 + "\n")
            f.write("[참고] PDF에서 발견한 '날짜가 있는 줄'(특정 날짜 행사 후보)\n")
            f.write("       PPT 파일로 넣으면 이 행사들도 자동으로 주차에 배치됩니다.\n")
            f.write("=" * 70 + "\n")
            for 줄, 날짜들 in 날짜후보:
                f.write("· (날짜: {})\n   {}\n\n".format(", ".join(날짜들), 줄))
    로그("[완료] 검토용 파일: " + os.path.basename(검토_경로))

    return {
        "결과경로": 결과_경로,
        "검토경로": 검토_경로,
        "연": 연, "월": 월,
        "월간제목들": [t for t, _ in 월간항목],
        "주차행사건수": [len(x) for x in 주차행사],
    }


def main():
    print("=" * 60)
    print(" AK PLAZA 카드 사은혜택 자동 입력 프로그램")
    print("=" * 60)

    기준폴더 = os.path.dirname(os.path.abspath(__file__))

    # 1) 파일 정하기 (직접 지정 > 자동 찾기)
    양식_경로 = 입력_경로 = None
    입력종류 = None
    for a in sys.argv[1:]:
        low = a.lower()
        if low.endswith(".xlsx"):
            양식_경로 = a
        elif low.endswith(".pptx"):
            입력_경로, 입력종류 = a, "ppt"
        elif low.endswith(".pdf"):
            입력_경로, 입력종류 = a, "pdf"

    자동_xlsx, 자동_pptx, 자동_pdf = 파일_자동찾기(기준폴더)
    if not 양식_경로:
        양식_경로 = 자동_xlsx
    if not 입력_경로:
        if 자동_pptx:
            입력_경로, 입력종류 = 자동_pptx, "ppt"   # PPT 우선
        elif 자동_pdf:
            입력_경로, 입력종류 = 자동_pdf, "pdf"

    if not 양식_경로 or not os.path.exists(양식_경로):
        print("\n[오류] 엑셀 양식 파일(.xlsx)을 찾지 못했습니다.")
        print("       이 프로그램과 같은 폴더에 엑셀 양식 파일을 넣어 주세요.")
        sys.exit(1)
    if not 입력_경로 or not os.path.exists(입력_경로):
        print("\n[오류] PPT(.pptx) 또는 PDF(.pdf) 파일을 찾지 못했습니다.")
        print("       이 프로그램과 같은 폴더에 이번 달 PPT(권장) 또는 PDF를 넣어 주세요.")
        sys.exit(1)

    # 2) 라이브러리 확인
    없음 = 라이브러리_확인(입력종류)
    if 없음:
        print("\n[오류] 다음 부품이 설치되어 있지 않습니다:", ", ".join(없음))
        print("       먼저 '최초설치.bat' 을 한 번 실행해 주세요.")
        sys.exit(1)

    # 2-2) 포함 여부 선택 (명령행 옵션 우선, 없으면 물어봅니다)
    args_low = [a.lower() for a in sys.argv[1:]]
    발급포함 = False if ("--no-발급" in args_low or "--no-issue" in args_low) else \
             (True if ("--발급" in args_low or "--issue" in args_low) else None)
    정기포함 = False if ("--no-정기" in args_low or "--no-regular" in args_low) else \
             (True if ("--정기" in args_low or "--regular" in args_low) else None)
    print("\n[선택] 아래 항목을 엑셀에 포함할까요?  (그냥 Enter=포함, n 입력=제외)")
    if 발급포함 is None:
        발급포함 = 예_아니오("  · 신한Plus 발급①/발급② 혜택 포함?", True)
    if 정기포함 is None:
        정기포함 = 예_아니오("  · 신한Plus 정기 혜택 포함?", True)

    # 3) 처리 실행
    혜택_처리(양식_경로, 입력_경로, 입력종류, 발급포함, 정기포함, 기준폴더, 로그=print)

    print("\n" + "=" * 60)
    print(" 끝났습니다! 결과 엑셀 파일을 열어 확인해 주세요.")
    print("=" * 60)


if __name__ == "__main__":
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
