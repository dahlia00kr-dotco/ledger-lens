"""샘플(더미) 전표 데이터 생성기.

가상의 제조회사 1년치(2025-01~12) 전표를 더존 iCUBE 전표출력 형식으로 만듭니다.
모든 회사명·거래처명·금액은 가상입니다.

12월(검토월)에는 도구가 잡아내야 할 이슈를 일부러 심어 둡니다.
  - 커피머신 렌탈료 누락             → 고정거래 누락
  - 경비보안 월정료 2번 입력          → 동일 전표 중복 / 건수 증가
  - 세무자문 수수료를 복리후생비로 처리→ 기존 거래처의 새 계정
  - 사무용품을 기업업무추진비로 처리   → 기존 거래처의 새 계정
  - 비품 감가상각비 누락              → 결산전표 누락
  - 시설자금대출 이자 미계상           → 결산전표 누락
  - 전기요금 급증, 시설관리 용역비 인상→ 금액 변동
  - 신규 컨설팅 거래처                → 신규 거래처
또 이력 중 7월에는 통신비가 빠져 있어, 월별 표에 누락 셀로 보입니다.

사용법:  python tools/make_sample.py   (openpyxl 필요)
"""
import calendar
import os
import random

from openpyxl import Workbook

random.seed(20251231)
OUT = os.path.join(os.path.dirname(__file__), "..", "sample")

A = {  # 계정명: 계정코드
    "보통예금": "10300", "외상매출금": "10800", "미수수익": "11600", "선급금": "13100", "선급비용": "13300",
    "부가세대급금": "13500", "원재료": "15300", "건물감가상각누계액": "20300", "기계장치감가상각누계액": "20700",
    "비품감가상각누계액": "21300", "사용권자산감가상각누계액": "21900", "특허권": "23200", "외상매입금": "25100",
    "미지급금": "25300", "예수금": "25400", "부가세예수금": "25500", "미지급비용": "26200", "퇴직급여충당부채": "29500",
    "제품매출": "40400", "급여(제)": "50300", "퇴직급여(제)": "50800", "전력비(제)": "51600", "감가상각비(제)": "51800",
    "외주가공비(제)": "53300", "직원급여": "80200", "연구원급여": "80210", "퇴직급여": "80600", "복리후생비": "81100",
    "식대": "81120", "국민연금(판)": "81130", "건강보험(판)": "81140", "고용보험(판)": "81150", "여비교통비": "81200",
    "기업업무추진비": "81300", "통신비": "81400", "감가상각비": "81800", "지급임차료": "81900", "소모품비": "83000",
    "지급수수료": "83100", "외주용역비": "83300", "무형고정자산상각": "84000", "사용권자산상각비": "84100",
    "이자수익": "90100", "이자비용": "93100",
}
V = {  # 키: (거래처코드, 거래처명)
    "emp": ("00001", "임직원"), "nps": ("00010", "국민연금공단"), "nhis": ("00011", "건강보험공단"),
    "ei": ("00012", "고용보험"), "tax": ("00013", "관할세무서"),
    "security": ("01001", "(주)한빛시큐리티"), "water": ("01002", "(주)브루웰렌탈"), "labor": ("01003", "세무법인 한결"),
    "power": ("01004", "대한전력공사"), "outsrc": ("01005", "(주)성원가공"), "clean": ("01006", "(주)그린빌딩관리"),
    "telecom": ("01007", "(주)스마트텔레콤"), "sw": ("01008", "(주)클라우드웍스"), "dorm": ("01009", "(주)든든물류"),
    "cust1": ("02001", "(주)한결유통"), "cust2": ("02002", "서울리테일(주)"), "cust3": ("02003", "Northwind Trading Co."),
    "sup1": ("03001", "(주)대한포장"), "sup2": ("03002", "대성소재(주)"),
    "consult": ("04001", "(주)인사이트컨설팅"),
    "loan1": ("70001", "가나은행 운영자금대출_6억"), "loan2": ("70002", "가나은행 시설자금대출_2억"),
    "dep": ("70010", "가나은행 정기예금_3억"),
    "card1": ("98001", "가나카드(1111)_경영지원"), "card2": ("98002", "가나카드(2222)_연구소"), "card3": ("98003", "가나카드(3333)_생산"),
}
# 법인카드 가맹점: (코드, 이름, 계정, 금액범위, 월평균 건수)
MERCHANTS = [
    ("05001", "한식뷔페 본점", "식대", (8000, 9000), 10),
    ("05002", "카페온더힐", "식대", (4500, 28000), 6),
    ("05003", "(주)든든도시락", "식대", (60000, 180000), 3),
    ("05004", "한국고속철도", "여비교통비", (18000, 52000), 6),
    ("05005", "하이웨이주유소", "여비교통비", (50000, 90000), 4),
    ("05006", "시티택시", "여비교통비", (6000, 24000), 4),
    ("05007", "오피스몰(주)", "소모품비", (20000, 250000), 3),
    ("05008", "생활용품마트", "소모품비", (5000, 60000), 3),
    ("05009", "한우명가", "기업업무추진비", (150000, 450000), 2),
    ("05010", "역앞편의점", "복리후생비", (3000, 30000), 3),
]
NEW_MERCHANTS_DEC = [("05011", "컨벤션센터 전시회 등록", "지급수수료", (330000, 330000)), ("05012", "항공권 예약센터", "여비교통비", (128000, 128000))]


class Book:
    def __init__(self):
        self.rows = []
        self.no_by_date = {}

    def voucher(self, date, lines):
        no = self.no_by_date.get(date, 0) + 1
        self.no_by_date[date] = no
        d = date.replace("-", "/")
        for i, (acct, vkey, memo, dr, cr) in enumerate(lines, 1):
            vcode, vname = V[vkey] if isinstance(vkey, str) and vkey in V else (vkey or (None, None))
            self.rows.append([d, str(no), str(i), d, str(no), "대체", A[acct], acct, vcode, vname, memo,
                              dr or None, cr or None, None])


def r100(x):
    return int(round(x / 100.0)) * 100


def vat_purchase(acct, vkey, memo, supply):
    vat = r100(supply * 0.1)
    return [(acct, vkey, memo, supply, 0), ("부가세대급금", vkey, memo, vat, 0), ("미지급금", vkey, memo, 0, supply + vat)]


def build():
    b = Book()
    for m in range(1, 13):
        ym = f"2025-{m:02d}"
        last = calendar.monthrange(2025, m)[1]
        day = lambda d: f"{ym}-{min(d, last):02d}"
        dec = m == 12
        yy = f"25년 {m}월"

        # ---- 월초 고정 비용 ----
        sec = vat_purchase("지급수수료", "security", f"{yy} 경비/보안 월정료", 250000)
        b.voucher(day(1), sec)
        if dec:  # 중복 입력
            b.voucher(day(3), sec)
        if not dec:  # 12월 누락
            b.voucher(day(1), vat_purchase("지급수수료", "water", f"{yy} 커피머신 렌탈료 2대", 120000))
        if dec:  # 계정 오류
            b.voucher(day(1), vat_purchase("복리후생비", "labor", f"{m}월분 세무자문 수수료", 350000))
        else:
            b.voucher(day(1), vat_purchase("지급수수료", "labor", f"{m}월분 세무자문 수수료", 350000))
        b.voucher(day(1), [("지급임차료", "dorm", f"{yy} 물류창고 임차료", 1800000, 0),
                           ("미지급금", "dorm", f"{yy} 물류창고 임차료", 0, 1800000)])
        if m != 7:  # 7월 통신비 누락(이력 속 누락 예시)
            tel = r100(random.uniform(410000, 440000))
            b.voucher(day(5), vat_purchase("통신비", "telecom", f"{yy} 인터넷·전화 요금", tel))
        if dec:
            b.voucher(day(9), vat_purchase("지급수수료", "consult", "ERP 도입 컨설팅 착수금", 4000000))

        # ---- 매출 ----
        for ck, (lo, hi), cnt in [("cust1", (8e6, 15e6), 2), ("cust2", (5e6, 9e6), 1), ("cust3", (12e6, 25e6), 1)]:
            for _ in range(cnt if random.random() > 0.15 else cnt + 1):
                sup = r100(random.uniform(lo, hi))
                d = day(random.randint(5, 26))
                memo = f"{yy} 제품 매출 [모델 A-100]"
                if ck == "cust3":
                    b.voucher(d, [("외상매출금", ck, memo + " 수출", sup, 0), ("제품매출", ck, memo + " 수출", 0, sup)])
                else:
                    vat = r100(sup * 0.1)
                    b.voucher(d, [("외상매출금", ck, memo, sup + vat, 0), ("제품매출", ck, memo, 0, sup), ("부가세예수금", ck, memo, 0, vat)])

        # ---- 원재료 매입 ----
        for sk, (lo, hi) in [("sup1", (3e6, 7e6)), ("sup2", (2e6, 4e6))]:
            for _ in range(random.randint(1, 3)):
                sup = r100(random.uniform(lo, hi))
                vat = r100(sup * 0.1)
                memo = f"{yy} 원재료 매입" + (" [포장재]" if sk == "sup1" else " [원자재]")
                b.voucher(day(random.randint(3, 27)), [("원재료", sk, memo, sup, 0), ("부가세대급금", sk, memo, vat, 0), ("외상매입금", sk, memo, 0, sup + vat)])

        # ---- 법인카드 ----
        cards = ["card1", "card2", "card3"]
        tx = []
        for code, name, acct, (lo, hi), avg in MERCHANTS:
            for _ in range(max(0, int(random.gauss(avg, avg ** 0.5)))):
                tx.append((code, name, acct, r100(random.uniform(lo, hi)) if lo != hi else lo, random.choice(cards), random.randint(1, last)))
        if dec:
            tx.append(("05007", "오피스몰(주)", "기업업무추진비", 184000, "card1", 12))  # 계정 오류
            for code, name, acct, (lo, _) in NEW_MERCHANTS_DEC:
                tx.append((code, name, acct, lo, "card1", random.randint(1, last)))
        by_card_week = {}
        for t in tx:
            by_card_week.setdefault((t[4], (t[5] - 1) // 7), []).append(t)
        for (card, wk), ts in sorted(by_card_week.items(), key=lambda x: (x[0][1], x[0][0])):
            lines = []
            for code, name, acct, amt, _, d in sorted(ts, key=lambda t: t[5]):
                memo = f"{name} {m}/{d} 법인카드 사용"
                lines += [(acct, (code, name), memo, amt, 0), ("미지급비용", card, memo, 0, amt)]
            b.voucher(day(min(last, wk * 7 + 7)), lines)

        # ---- 급여 ----
        n_sales, n_rnd, n_prod = 12 + (m >= 7), 4 + (m >= 10), 18 + (1 if m in (3, 8) else 0)
        g_s, g_r, g_p = n_sales * 3_600_000, n_rnd * 4_100_000, n_prod * 2_900_000
        gross = g_s + g_r + g_p
        np_, hi_, ei_, it_ = r100(gross * 0.045), r100(gross * 0.03545), r100(gross * 0.009), r100(gross * 0.06)
        net = gross - np_ - hi_ - ei_ - it_
        pm = f"{yy} 급여"
        b.voucher(day(25), [
            ("직원급여", "emp", f"{pm}_판관 {n_sales}명", g_s, 0), ("연구원급여", "emp", f"{pm}_연구소 {n_rnd}명", g_r, 0),
            ("급여(제)", "emp", f"{pm}_생산 {n_prod}명", g_p, 0),
            ("예수금", "nps", f"{pm}_국민연금 예수금", 0, np_), ("예수금", "nhis", f"{pm}_건강보험 예수금", 0, hi_),
            ("예수금", "ei", f"{pm}_고용보험 예수금", 0, ei_), ("예수금", "tax", f"{pm}_소득세 예수금", 0, it_),
            ("미지급비용", "emp", f"{pm}_임직원 급여 {n_sales + n_rnd + n_prod}명", 0, net)])
        b.voucher(day(25), [
            ("국민연금(판)", "nps", f"{yy} 국민연금 회사부담금", np_, 0), ("미지급비용", "nps", f"{yy} 국민연금 회사부담금", 0, np_),
            ("건강보험(판)", "nhis", f"{yy} 건강보험 회사부담금", hi_, 0), ("미지급비용", "nhis", f"{yy} 건강보험 회사부담금", 0, hi_),
            ("고용보험(판)", "ei", f"{yy} 고용보험 회사부담금", r100(gross * 0.0115), 0), ("미지급비용", "ei", f"{yy} 고용보험 회사부담금", 0, r100(gross * 0.0115))])
        b.voucher(day(10), vat_purchase("외주가공비(제)", "outsrc", f"{yy} 외주가공비", r100(random.uniform(22e6, 27e6))))

        # ---- 공과금·용역 ----
        base_power = [1.35, 1.30, 1.20, 1.10, 1.15, 1.45, 1.80, 1.85, 1.40, 1.15, 1.25, 1.40][m - 1] * 1e6
        if dec:
            base_power *= 2.6  # 급증
        b.voucher(day(15), vat_purchase("전력비(제)", "power", f"{yy} 전기요금", r100(base_power * random.uniform(0.97, 1.03))))
        b.voucher(day(28), vat_purchase("외주용역비", "clean", f"{m}월 시설관리 용역비", 2_640_000 if dec else 2_200_000))

        # ---- 월말 결산전표 ----
        c = f"{m}월"
        dep = [("감가상각비", "건물감가상각누계액", "건물", 1_250_000), ("감가상각비", "비품감가상각누계액", "비품", 430_000),
               ("감가상각비(제)", "기계장치감가상각누계액", "기계장치", 2_100_000)]
        lines = []
        for exp, acc, nm, amt in dep:
            if dec and nm == "비품":
                continue  # 누락
            lines += [(exp, None, f"{c} 감가상각비 계상({nm})", amt, 0), (acc, None, f"{c} 감가상각비 계상({nm})", 0, amt)]
        lines += [("무형고정자산상각", None, f"{m}월 특허권 상각", 320_000, 0), ("특허권", None, f"{m}월 특허권 상각", 0, 320_000),
                  ("사용권자산상각비", None, f"{yy} 리스차량 사용권자산 상각", 650_000, 0),
                  ("사용권자산감가상각누계액", None, f"{yy} 리스차량 사용권자산 상각", 0, 650_000)]
        b.voucher(day(31), lines)
        rs, rr, rp = r100(g_s * 0.083), r100(g_r * 0.083), r100(g_p * 0.083)
        b.voucher(day(31), [("퇴직급여", None, f"{yy} 퇴직급여충당부채 전입(판관·연구)", rs + rr, 0),
                            ("퇴직급여(제)", None, f"{yy} 퇴직급여충당부채 전입(생산)", rp, 0),
                            ("퇴직급여충당부채", None, f"{yy} 퇴직급여충당부채 전입", 0, rs + rr + rp)])
        b.voucher(day(31), [("지급수수료", "sw", f"{m}월 선급비용 상각 [디자인 SW 연간 라이선스 25.01~25.12]", 275_000, 0),
                            ("선급비용", "sw", f"{m}월 선급비용 상각 [디자인 SW 연간 라이선스 25.01~25.12]", 0, 275_000)])
        i1 = r100(600_000_000 * 0.0485 * last / 365)
        lines = [("이자비용", "loan1", f"운영자금대출 {m}월 이자 미지급 계상", i1, 0), ("미지급비용", "loan1", f"운영자금대출 {m}월 이자 미지급 계상", 0, i1)]
        if not dec:  # 12월 미계상
            i2 = r100(200_000_000 * 0.039 * last / 365)
            lines += [("이자비용", "loan2", f"시설자금대출 {m}월 이자 미지급 계상", i2, 0), ("미지급비용", "loan2", f"시설자금대출 {m}월 이자 미지급 계상", 0, i2)]
        b.voucher(day(31), lines)
        ii = r100(300_000_000 * 0.034 * last / 365)
        b.voucher(day(31), [("미수수익", "dep", f"정기예금 {m}월 미수이자 계상", ii, 0), ("이자수익", "dep", f"정기예금 {m}월 미수이자 계상", 0, ii)])
    return b.rows


HEAD = ["결의일", "결의No", "순번", "기표일자", "기표번호", "구분", "코드", "계정과목명", "코드", "거래처명", "적요", "차변", "대변", "출력일자"]


def save(rows, name):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1_1"
    ws.append(HEAD)
    for r in rows:
        ws.append(r)
    for col, w in zip("ABCDEFGHIJKLMN", [11, 7, 5, 11, 7, 5, 7, 18, 7, 26, 52, 14, 14, 10]):
        ws.column_dimensions[col].width = w
    wb.save(os.path.join(OUT, name))


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    rows = build()
    rows.sort(key=lambda r: (r[3], int(r[4]), int(r[2])))
    save([r for r in rows if r[3] < "2025/12"], "journal_2025_01-11.xlsx")
    save([r for r in rows if r[3] >= "2025/12"], "journal_2025_12.xlsx")
    dr = sum(r[11] or 0 for r in rows)
    cr = sum(r[12] or 0 for r in rows)
    print(f"{len(rows)} lines, debit {dr:,} / credit {cr:,}")
