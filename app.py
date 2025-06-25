from flask import Flask, request, jsonify, url_for
from pptx import Presentation
import os
import uuid
from datetime import datetime

app = Flask(__name__)

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    try:
        data = request.get_json()  # 클라이언트가 보낸 JSON 데이터 받기
        selected_rows = data.get('selectedRows', [])  # 'selectedRows' 키로 데이터 받기

        if not selected_rows:
            return jsonify({"error": "선택된 행이 없습니다."}), 400

        template_path = os.path.join("static", "certi_template.pptx")
        output_dir = "static"
        os.makedirs(output_dir, exist_ok=True)

        if not os.path.exists(template_path):
            return jsonify({"error": f"템플릿 파일이 존재하지 않습니다: {template_path}"}), 500

        prs = Presentation(template_path)

        for row in selected_rows:
            name = row.get("fomat_name", "")
            subject = row.get("upper_subject", "")
            amount = row.get("paid_amount", "")
            period = row.get("period", "")
            date_str = datetime.today().strftime("%Y년 %m월 %d일")

            slide = prs.slides.add_slide(prs.slide_layouts[0])
            try:
                slide.shapes[0].text = name
                slide.shapes[1].text = subject
                slide.shapes[2].text = amount
                slide.shapes[3].text = period
                slide.shapes[4].text = date_str
            except IndexError as e:
                return jsonify({"error": f"텍스트박스 부족: {e}"}), 500

        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

        filename = f"certificate_{uuid.uuid4().hex}.pptx"
        save_path = os.path.join(output_dir, filename)
        prs.save(save_path)

        file_url = url_for('static', filename=filename, _external=True)
        return jsonify({"message": "수강증 저장 성공", "file_url": file_url})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000, debug=True)
