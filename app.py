from flask import Flask, jsonify, send_from_directory
from pptx import Presentation
import os
import uuid
from datetime import datetime

app = Flask(__name__)

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    try:
        # 📌 템플릿 경로: 프로젝트 내 static 폴더 기준 상대경로
        template_path = os.path.join("static", "certi_template.pptx")
        output_dir = "static"
        os.makedirs(output_dir, exist_ok=True)

        if not os.path.exists(template_path):
            return jsonify({"error": f"템플릿 파일이 존재하지 않습니다: {template_path}"}), 500

        prs = Presentation(template_path)

        selected_rows = [
            {
                "fomat_name": "이 시 은",
                "upper_subject": "수학 집중관리 프리패스",
                "paid_amount": "300,000원",
                "period": "24. 09. 23 ~ 25. 10. 20",
            },
            {
                "fomat_name": "이 시 은",
                "upper_subject": "영어 집중관리 프리패스",
                "paid_amount": "280,000원",
                "period": "25. 09. 23 ~ 26. 10. 20",
            }
        ]

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

        # 템플릿 슬라이드 제거
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

        filename = f"certificate_{uuid.uuid4().hex}.pptx"
        save_path = os.path.join(output_dir, filename)
        prs.save(save_path)

        print(f"파일 저장 위치: {os.path.abspath(save_path)}")  # 파일 저장 경로 출력

        return jsonify({"message": "수강증 저장 성공", "file": filename})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# 추가: 다운로드 API
@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    output_dir = "static"
    try:
        file_path = os.path.join(output_dir, filename)
        if not os.path.exists(file_path):
            return jsonify({"error": "파일이 존재하지 않습니다."}), 404
        return send_from_directory(output_dir, filename, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(host="0.0.0.0", port=10000, debug=True)
