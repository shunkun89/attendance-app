from flask import Flask, request, render_template_string
import pandas as pd
import os

app = Flask(__name__)

df = pd.read_excel('members.xlsx', engine='openpyxl')

form_html = '''
<!doctype html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>出欠確認フォーム</title>
  <style>
    body { font-family: sans-serif; padding: 20px; max-width: 500px; margin: auto; }
    input, select { width: 100%; padding: 10px; margin: 8px 0; }
    button, input[type="submit"] { padding: 10px; width: 100%; background: #4CAF50; color: white; border: none; border-radius: 5px; }
    .error { color: red; margin-top: 10px; }
  </style>
</head>
<body>
  <h2>出欠確認フォーム</h2>
  <form method="post">
    コードを入力してください：
    <input type="text" name="code" value="{{ code or '' }}" required>
    <input type="submit" value="検索">
  </form>

  {% if error %}
    <div class="error">{{ error }}</div>
  {% endif %}

  {% if name %}
  <hr>
  <p>氏名：{{ name }}</p>
  <p>クラス：{{ class_name }}</p>
  <form method="post">
    <input type="hidden" name="code" value="{{ code }}">
    出欠：
    <select name="attendance" required>
      <option value="">選択してください</option>
      <option value="出席">出席</option>
      <option value="欠席">欠席</option>
    </select><br>
    交通手段：
    <select name="transport" required>
      <option value="">選択してください</option>
      <option value="電車">電車</option>
      <option value="バス">バス</option>
      <option value="自家用車">自家用車</option>
      <option value="徒歩">徒歩</option>
    </select><br>
    懇親会：
    <select name="party" required>
      <option value="">選択してください</option>
      <option value="参加">参加</option>
      <option value="不参加">不参加</option>
    </select><br>
    <input type="submit" name="submit" value="送信">
  </form>
  {% endif %}

  {% if submitted %}
  <hr>
  <h3>送信されました！</h3>
  <p>氏名：{{ name }}</p>
  <p>出欠：{{ attendance }}</p>
  <p>交通手段：{{ transport }}</p>
  <p>懇親会：{{ party }}</p>
  {% endif %}
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def index():
    name = class_name = attendance = transport = party = None
    submitted = False
    error = None
    code = request.form.get('code')

    if code and not request.form.get('submit'):
        if not code.isdigit():
            error = '数字のコードを入力してください。'
        else:
            person = df[df['コード'] == int(code)]
            if person.empty:
                error = '該当する情報が見つかりません。'
            else:
                name = person.iloc[0]['氏名']
                class_name = person.iloc[0]['クラス']

    elif request.form.get('submit'):  # フォーム送信処理
        code = request.form.get('code')
        person = df[df['コード'] == int(code)]
        if not person.empty:
            name = person.iloc[0]['氏名']
            class_name = person.iloc[0]['クラス']
            attendance = request.form.get('attendance')
            transport = request.form.get('transport')
            party = request.form.get('party')
            submitted = True

            # 結果をExcelファイルに保存（追記）
            response_df = pd.DataFrame([{
                'コード': code,
                '氏名': name,
                'クラス': class_name,
                '出欠': attendance,
                '交通手段': transport,
                '懇親会': party
            }])

            file_path = 'responses.xlsx'
            if os.path.exists(file_path):
                existing = pd.read_excel(file_path, engine='openpyxl')
                new_df = pd.concat([existing, response_df], ignore_index=True)
            else:
                new_df = response_df

            new_df.to_excel(file_path, index=False, engine='openpyxl')

    return render_template_string(form_html,
                                  name=name,
                                  class_name=class_name,
                                  code=code,
                                  attendance=attendance,
                                  transport=transport,
                                  party=party,
                                  submitted=submitted,
                                  error=error)

if __name__ == '__main__':
    app.run(debug=True)
