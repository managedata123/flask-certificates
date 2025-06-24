from flask import Flask, request, jsonify, url_for
from pptx import Presentation
import os
from datetime import datetime
import uuid

app = Flask(__name__)

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    try:
        print("API 함수 진입")
        data = request.get_json()
        print(f"받은 데이터: {data}")

        selected_rows = data.get('selectedRows', [])
        print(f"선택된 행 개수: {len(selected_rows)}")

        template_path = os.path.join('static', 'certi_template.pptx')
        print(f"템플릿 경로: {template_path}")
        if not os.path.exists(template_path):
            print("템플릿 파일 없음!")
            raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_path}")

        prs = Presentation(template_path)
        print("템플릿 파일 로드 완료")

        # 슬라이드 추가 및 내용 채우기 (생략)

        unique_filename = f"certificate_{uuid.uuid4().hex}.pptx"
        save_path = os.path.join('static', unique_filename)
        print("파일 저장 시작")
        prs.save(save_path)
        print(f"파일 저장 완료: {save_path}")

        file_url = url_for('static', filename=unique_filename, _external=True)
        print(f"응답 파일 URL: {file_url}")

        return jsonify({"file_url": file_url})

    except Exception as e:
        print(f"에러 발생: {e}")
        return jsonify({"error": str(e)}), 500
