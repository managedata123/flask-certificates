from flask import Flask, request, jsonify, url_for
from pptx import Presentation
import os
import uuid
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def home():
    return "Flask 수강증 생성 서버입니다."

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    try:
        print("✅ API 함수 진입")
        data = request.get_json()
        print(f"🔥 받은 전체 데이터: {data}")

        selected_rows = data.get('selectedRows', [])
        print(f"🔹 선택된 행 개수: {len(selected_rows)}")

        if not selected_rows:
            raise ValueError("선택된 행이 없습니다.")

        template_path = os.path.join('static', 'certi_template.pptx')
        print(f"📄 템플릿 경로: {template_path}")

        if not os.path.exists(template_path):
            raise FileNotFoundError("템플릿 파일이 존재하지 않습니다.")

        prs = Presentation(template_path)

        for row in selected_rows:
            name = row.get('fomat_name', '')
            subject = row.get('upper_subject', '')
            amount = row.get('paid_amount', '')
            period = row.get('period', '')
            date_str = datetime.today().strftime("%Y년 %m월 %d일")

            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)

            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                text = shape.text_frame.text.strip()
                if '성명~' in text:
                    shape.text_frame.text = f"성명: {name}"
                elif '과목~' in text:
                    shape.text_frame.text = f"과목: {subject}"
                elif '금액~' in text:
                    shape.text_frame.text = f"금액: {amount}"
                elif '기간~' in text:
                    shape.text_frame.text = f"기간: {period}"
                elif '날인일~' in text:
                    shape.text_frame.text = f"날인일: {date_str}"

        unique_filename = f"certificate_{uuid.uuid4().hex}.pptx"
        save_path = os.path.join('static', unique_filename)

        prs.save(save_path)
        print(f"✅ 파일 저장 완료: {save_path}")

        file_url = url_for('static', filename=unique_filename, _external=True)
        print(f"✅ 응답 파일 URL: {file_url}")

        return jsonify({"file_url": file_url})

    except Exception as e:
        print(f"❌ 에러 발생: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("✅ Flask 앱 시작됨")
    app.run(host='0.0.0.0', port=5001, debug=True)
