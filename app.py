import streamlit as st
import pandas as pd
import psycopg2
from datetime import datetime
import io  # メモリ内でファイルを扱うため
import os
from dotenv import load_dotenv

# Excelファイルからデータを読み込み
file_path = 'プルダウンマスター.xlsx'
data = pd.read_excel(file_path)

# 欠損値を空文字列に置換し、文字列型に変換
data['都道府県'] = data['都道府県'].fillna('').astype(str)
data['産業分類主業名'] = data['産業分類主業名'].fillna('').astype(str)
data['産業分類従業名'] = data['産業分類従業名'].fillna('').astype(str)

# 都道府県、産業分類主業名、産業分類従業名のリストを取得
prefectures = sorted(data['都道府県'].unique().tolist())
main_industries = sorted(data['産業分類主業名'].unique().tolist())
sub_industries = sorted(data['産業分類従業名'].unique().tolist())


# .envファイルをロード
load_dotenv()

# データベース接続を初期化時に開く
conn = psycopg2.connect(
    host=os.getenv('DB_HOST'),
    database=os.getenv('DB_NAME'),
    user=os.getenv('DB_USER'),
    password=os.getenv('DB_PASSWORD'),
    port=os.getenv('DB_PORT')
)

cursor = conn.cursor()

# タイトルを設定
st.title("企業リスト作成アプリ")

# セッションステートの初期化
if 'current_conditions' not in st.session_state:
    st.session_state['current_conditions'] = []
if 'current_params' not in st.session_state:
    st.session_state['current_params'] = []
if 'total_count' not in st.session_state:
    st.session_state['total_count'] = 0
if 'df_display' not in st.session_state:
    st.session_state['df_display'] = pd.DataFrame()

# ボタンを横並びで配置
col1, col2, col3 = st.columns(3)
with col1:
    get_selection = st.button('選択内容を取得')
with col2:
    export_excel = st.button('Excel出力')
with col3:
    reset_selection = st.button('選択リセット')

# 選択リセットボタンの処理
if reset_selection:
    for key in ['selected_prefs', 'selected_main_inds', 'selected_sub_inds', 'current_conditions', 'current_params', 'total_count', 'df_display']:
        if key in st.session_state:
            del st.session_state[key]
    # ページを再実行して選択をリセット
    st.experimental_rerun()

# 選択肢を横並びに配置
col_pref, col_main_ind, col_sub_ind = st.columns(3)

with col_pref:
    selected_prefs = st.multiselect('都道府県', prefectures, key='selected_prefs')

with col_main_ind:
    selected_main_inds = st.multiselect('業種', main_industries, key='selected_main_inds')

with col_sub_ind:
    selected_sub_inds = st.multiselect('副業種', sub_industries, key='selected_sub_inds')

# データベースのカラム名（'id'を最後に配置）
columns = ['会社名', '都道府県', '住所', '電話番号', '産業分類主業名', '産業分類従業名',
           '資本金', '従業員数', '設立年', '売上高', '代表者', '主要仕入先', '主要販売先', '株主', 'id']

# 「選択内容を取得」ボタンが押されたときの処理
if get_selection:
    query = "SELECT * FROM companies"
    conditions = []
    params = []

    # 各選択肢について条件を作成
    if selected_prefs:
        placeholders = ','.join(['%s'] * len(selected_prefs))
        conditions.append(f"都道府県 IN ({placeholders})")
        params.extend(selected_prefs)
    if selected_main_inds:
        placeholders = ','.join(['%s'] * len(selected_main_inds))
        conditions.append(f"産業分類主業名 IN ({placeholders})")
        params.extend(selected_main_inds)
    if selected_sub_inds:
        placeholders = ','.join(['%s'] * len(selected_sub_inds))
        conditions.append(f"産業分類従業名 IN ({placeholders})")
        params.extend(selected_sub_inds)

    if conditions:
        where_clause = " WHERE " + " AND ".join(conditions)
        query += where_clause
        count_query = "SELECT COUNT(*) FROM companies" + where_clause
    else:
        count_query = "SELECT COUNT(*) FROM companies"

    # 条件とパラメータをセッションステートに保存
    st.session_state['current_conditions'] = conditions
    st.session_state['current_params'] = params

    # 総データ数を取得
    cursor.execute(count_query, params)
    total_count = cursor.fetchone()[0]
    st.session_state['total_count'] = total_count

    # データを取得（表示用に制限）
    limit = 30
    data_query = query + f" LIMIT {limit}"
    cursor.execute(data_query, params)
    rows = cursor.fetchall()

    # データをDataFrameに変換してセッションステートに保存
    df_display = pd.DataFrame(rows, columns=columns)
    st.session_state['df_display'] = df_display

# 検索結果の表示
if st.session_state['total_count'] > 0:
    st.write(f'総データ数: {st.session_state["total_count"]}')
    st.dataframe(st.session_state['df_display'])
elif get_selection:
    st.write('条件に該当するデータがありません。')

# 「Excel出力」ボタンが押されたときの処理
if export_excel:
    if st.session_state['current_conditions']:
        conditions = st.session_state['current_conditions']
        params = st.session_state['current_params']
        query = "SELECT * FROM companies"
        if conditions:
            where_clause = " WHERE " + " AND ".join(conditions)
            query += where_clause

        cursor.execute(query, params)
        rows = cursor.fetchall()

        if not rows:
            st.warning("条件に該当するデータがありません。")
        elif len(rows) > 20000:
            st.warning(f"データが多すぎます（{len(rows)}行）。20,000行以下のデータのみ出力できます。条件を絞り込んでください。")
        else:
            df_export = pd.DataFrame(rows, columns=columns)
            # メモリ内にExcelファイルを作成
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_export.to_excel(writer, index=False)
            output.seek(0)
            data = output.getvalue()

            # ダウンロードボタンを表示
            st.download_button(
                label='Excelファイルをダウンロード',
                data=data,
                file_name=f'企業リスト_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    else:
        st.warning('先に「選択内容を取得」をクリックしてください。')

# データベース接続を閉じる（Streamlitでは接続を閉じない方が良いのでコメントアウト）
# conn.close()
