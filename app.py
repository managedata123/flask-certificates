from flask import Flask, request, jsonify, url_for
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

        template_path = "/Users/mildang/flask_certificates/static/certi_template.pptx"
        if not os.path.exists(template_path):
            return jsonify({"error": "템플릿 파일이 존재하지 않습니다."}), 500

        prs = Presentation(template_path)

        # 템플릿의 첫 슬라이드는 기본 슬라이드로 남겨두고, 이후 새 슬라이드를 추가하며 데이터 삽입
        for row in selected_rows:
            name = row.get("fomat_name", "")
            subject = row.get("upper_subject", "")
            amount = row.get("paid_amount", "")
            period = row.get("period", "")
            date_str = datetime.today().strftime("%Y년 %m월 %d일")

            slide_layout = prs.slide_layouts[0]  # 템플릿과 동일한 레이아웃
            slide = prs.slides.add_slide(slide_layout)

            # 인덱스로 텍스트박스 접근해 텍스트 변경 (VBA 방식과 동일)
            try:
                slide.shapes[0].text = name
                slide.shapes[1].text = subject
                slide.shapes[2].text = amount
                slide.shapes[3].text = period
                slide.shapes[4].text = date_str
            except IndexError:
                return jsonify({"error": "슬라이드 텍스트 박스 개수가 부족합니다."}), 500

        # 첫번째 기본 템플릿 슬라이드는 제거
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

        unique_filename = f"certificate_{uuid.uuid4().hex}.pptx"
        output_dir = "/Users/mildang/flask_certificates/static"
        save_path = os.path.join(output_dir, unique_filename)
        prs.save(save_path)

        file_url = url_for('static', filename=unique_filename, _external=True)
        return jsonify({"file_url": file_url})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)
