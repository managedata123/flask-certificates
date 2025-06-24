from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from datetime import datetime
import tempfile
import os

app = Flask(__name__)

def format_korean_date(date_obj):
    return date_obj.strftime("%Y년 %m월 %d일")

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    data = request.json.get("selectedRows", [])
    if not data:
        return jsonify({"error": "No data received"}), 400

    temp_dir = tempfile.mkdtemp()
    pptx_template_path = "/Users/mildang/Downloads/certi_template.pptx"  # 네가 준 템플릿 경로로 바꿈
    result_pptx_path = os.path.join(temp_dir, "result.pptx")

    prs = Presentation(pptx_template_path)

    # 템플릿 슬라이드 1개만 남기고 나머지 삭제
    while len(prs.slides) > 1:
        r_id = prs.slides._sldIdLst[-1].rId
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[-1])
        prs.part.drop_rel(r_id)

    first_slide = prs.slides[0]

    for idx, row in enumerate(data):
        slide = first_slide if idx == 0 else prs.slides.add_slide(first_slide.slide_layout)

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text
            text = text.replace("[학생이름]", str(row.get("format_name", "")))
            text = text.replace("[수강반]", str(row.get("upper_subject", "")))
            text = text.replace("[결제액]", str(row.get("paid_amount", "")))
            text = text.replace("[수강기간]", str(row.get("period", "")))
            text = text.replace("[날짜]", format_korean_date(datetime.now()))
            shape.text_frame.text = text

    prs.save(result_pptx_path)

    # 아직 PDF 변환은 안 함. PPTX 파일 그대로 보내기
    return send_file(result_pptx_path, as_attachment=True, download_name="수강증.pptx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)

