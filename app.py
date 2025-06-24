from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from datetime import datetime
import tempfile
import os

app = Flask(__name__)

# 날짜를 "YYYY년 MM월 DD일" 형식으로 바꾸는 함수
def format_korean_date(date_obj):
    return date_obj.strftime("%Y년 %m월 %d일")

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    data = request.json.get("selectedRows", [])
    if not data:
        return jsonify({"error": "No data received"}), 400

    # 임시 폴더 만들기
    temp_dir = tempfile.mkdtemp()
    # 템플릿 파일 경로 (상대경로)
    pptx_template_path = "certi_template.pptx"
    result_pptx_path = os.path.join(temp_dir, "result.pptx")

    # 템플릿 열기
    prs = Presentation(pptx_template_path)

    # 슬라이드 여러 개면 첫 슬라이드만 남기고 제거
    while len(prs.slides) > 1:
        r_id = prs.slides._sldIdLst[-1].rId
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[-1])
        prs.part.drop_rel(r_id)

    first_slide = prs.slides[0]

    for idx, row in enumerate(data):
        # 첫 슬라이드는 그대로 사용, 나머지는 복제해서 새 슬라이드 추가
        slide = first_slide if idx == 0 else prs.slides.add_slide(first_slide.slide_layout)

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text

            # 텍스트 상자 내 텍스트 치환
            text = text.replace("성명~", f"성명: {row.get('format_name', '')}")
            text = text.replace("과목~", f"과목: {row.get('upper_subject', '')}")
            text = text.replace("금액~", f"금액: {row.get('paid_amount', '')}")
            text = text.replace("기간~", f"기간: {row.get('period', '')}")
            text = text.replace("날인일~", f"날인일: {format_korean_date(datetime.now())}")

            shape.text_frame.text = text

    # 결과 PPTX 저장
    prs.save(result_pptx_path)

    # 결과 파일을 다운로드용 응답으로 전송
    return send_file(result_pptx_path, as_attachment=True, download_name="수강증.pptx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)

