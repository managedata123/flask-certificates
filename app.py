from pptx import Presentation
import os
import uuid
from datetime import datetime

# 절대 경로 설정
template_path = "/Users/mildang/flask_certificates/static/certi_template.pptx"
output_dir = "/Users/mildang/flask_certificates/static"
os.makedirs(output_dir, exist_ok=True)

# 프레젠테이션 열기
prs = Presentation(template_path)

# 엑셀 대신 하드코딩된 테스트 데이터
selected_rows = [
    {
        "fomat_name": "홍길동",
        "upper_subject": "수학",
        "paid_amount": "300,000원",
        "period": "2025.03.01 ~ 2025.06.01",
    },
    {
        "fomat_name": "이순신",
        "upper_subject": "과학",
        "paid_amount": "280,000원",
        "period": "2025.04.01 ~ 2025.07.01",
    }
]

# 템플릿 슬라이드는 첫 장으로 남기고, 복사하지 않고 새로운 슬라이드 레이아웃을 사용
for row in selected_rows:
    name = row.get("fomat_name", "")
    subject = row.get("upper_subject", "")
    amount = row.get("paid_amount", "")
    period = row.get("period", "")
    date_str = datetime.today().strftime("%Y년 %m월 %d일")

    # 새 슬라이드 추가
    layout = prs.slide_layouts[0]  # 템플릿과 동일한 레이아웃 (보통 Title and Content)
    slide = prs.slides.add_slide(layout)

    # 슬라이드 내 텍스트박스 순서대로 텍스트 삽입
    try:
        slide.shapes[0].text = name        # 성명
        slide.shapes[1].text = subject     # 과목
        slide.shapes[2].text = amount      # 금액
        slide.shapes[3].text = period      # 기간
        slide.shapes[4].text = date_str    # 날인일
    except IndexError as e:
        print(f"❌ 텍스트박스 개수가 부족합니다: {e}")

# 템플릿 슬라이드(0번 인덱스) 삭제
prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

# 저장
filename = f"certificate_test_{uuid.uuid4().hex}.pptx"
output_path = os.path.join(output_dir, filename)
prs.save(output_path)
print(f"✅ 수강증 저장 완료: {output_path}")
