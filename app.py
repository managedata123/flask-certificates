from pptx import Presentation
import os
import uuid
from datetime import datetime

# 상대 경로 설정 (static 폴더 기준)
template_path = os.path.join("static", "certi_template.pptx")
output_dir = "static"
os.makedirs(output_dir, exist_ok=True)

# 프레젠테이션 템플릿 열기
prs = Presentation(template_path)

# 테스트용 하드코딩 데이터
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

# 수강증 슬라이드 생성
for row in selected_rows:
    name = row.get("fomat_name", "")
    subject = row.get("upper_subject", "")
    amount = row.get("paid_amount", "")
    period = row.get("period", "")
    date_str = datetime.today().strftime("%Y년 %m월 %d일")

    # 새 슬라이드 추가
    layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(layout)

    try:
        slide.shapes[0].text = name
        slide.shapes[1].text = subject
        slide.shapes[2].text = amount
        slide.shapes[3].text = period
        slide.shapes[4].text = date_str
    except IndexError as e:
        print(f"❌ 텍스트박스 부족: {e}")

# 템플릿 첫 슬라이드는 제거
prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

# 결과 저장
filename = f"certificate_test_{uuid.uuid4().hex}.pptx"
output_path = os.path.join(output_dir, filename)
prs.save(output_path)

print(f"✅ 수강증 저장 완료: {output_path}")
