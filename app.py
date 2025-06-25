from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/')
def home():
    return "Flask 테스트 서버입니다."

@app.route('/generate_certificates', methods=['POST'])
def generate_certificates():
    data = request.get_json()
    selected_rows = data.get('selectedRows', [])
    print("🔹 받은 selectedRows 데이터:")
    for i, row in enumerate(selected_rows, 1):
        print(f"  행 {i}: {row}")
    return jsonify({"message": f"선택된 행 {len(selected_rows)}개를 받았습니다."})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
