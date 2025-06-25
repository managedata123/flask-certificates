from flask import Flask, request, jsonify
from pptx import Presentation
import os
import uuid
from datetime import datetime

app = Flask(__name__)

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    try:
        data = request.get_json()
        selected_rows = data.get('selectedRows', [])
        if not selected_rows:
            return jsonify({"error": "선택된 행이 없습니다."}), 400

        # ✅ 절대 경로로 템플릿 파일 지정
        template_path = "/Users/mildang/flask_certificates/static/certi_template.pptx"
        if not os.path.exists(template_path):
            return jsonify({"error": "템플릿 파일이 존재하지 않습니다."}), 500

        prs = Presentation(template_path)

        for row in selected_rows:
            name = row.get("fomat_name", "")
            subject = row.get("upper_subject", "")
            amount = row.get("paid_amount", "")
            period = row.get("period", "")
            date_str = datetime.today().strftime("%Y년 %m월 %d일")

            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)

            try:
                slide.shapes[0].text = name
                slide.shapes[1].text = subject
                slide.shapes[2].text = amount
                slide.shapes[3].text = period
                slide.shapes[4].text = date_str
            except IndexError:
                return jsonify({"error": "슬라이드 텍스트 박스 개수가 부족합니다."}), 500

        # 템플릿 슬라이드 제거
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

        # ✅ 저장할 파일 경로
        output_dir = "/Users/mildang/flask_certificates/static"
        os.makedirs(output_dir, exist_ok=True)
        filename = f"certificate_{uuid.uuid4().hex}.pptx"
        save_path = os.path.join(output_dir, filename)

        prs.save(save_path)
        print(f"✅ 저장된 경로: {save_path}")

        return jsonify({"message": "수강증 저장 완료", "file_path": save_path})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)
