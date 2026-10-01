import streamlit as st
from google import genai
from google.genai import types
from pypdf import PdfReader
from docx import Document
from pptx import Presentation
import io

# --- ページ設定 ---
st.set_page_config(page_title="マイリサーチAI", layout="wide")
st.title("🧠 マイ・NotebookLM ＆ リサーチAI")
st.write("PDFの要約、クイズ作成、Google検索、そしてWord・PowerPointの自動生成ができます。")

# --- セッション情報の保存（画面が更新されてもデータを消さない仕組み） ---
if "ai_result" not in st.session_state:
    st.session_state.ai_result = None

# --- サイドバー（APIキー設定） ---
with st.sidebar:
    st.header("⚙️ 設定")
    api_key = st.text_input("Gemini APIキーを入力", type="password")
    st.markdown("[Google AI Studio](https://aistudio.google.com/) で無料で取得できます。")

# --- ファイル操作・ファイル生成関数 ---
def extract_text_from_pdf(file):
    reader = PdfReader(file)
    text = ""
    for page in reader.pages:
        if page.extract_text():
            text += page.extract_text()
    return text

def create_word_file(text):
    doc = Document()
    doc.add_heading('AI 解析レポート', 0)
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_ppt_file(text):
    prs = Presentation()
    # 表紙
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "AI 解析スライド"
    title_slide.placeholders[1].text = "自動生成された資料"
    # 内容を段落ごとにスライド化（長すぎる場合は分割）
    paragraphs = text.split('\n\n')
    for para in paragraphs:
        if para.strip():
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = "ポイント"
            slide.placeholders[1].text = para.strip()
    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()

# --- メイン画面 ---
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📥 入力エリア")
    uploaded_file = st.file_uploader("PDFファイルをアップロード (任意)", type=["pdf"])
    
    mode = st.radio("モードを選択", ["📄 要約する", "❓ クイズを作る", "🔍 Google検索でリサーチ"])
    user_prompt = st.text_input("追加の質問や検索キーワード (任意)")
    
    if st.button("🚀 AIに実行させる", type="primary"):
        if not api_key:
            st.error("左のサイドバーにGemini APIキーを入力してください！")
        else:
            with st.spinner("AIが処理中・検索中です..."):
                try:
                    client = genai.Client(api_key=api_key)
                    
                    # PDFの読み込み
                    file_text = extract_text_from_pdf(uploaded_file) if uploaded_file else ""
                    
                    tools = []
                    if mode == "📄 要約する":
                        prompt = f"以下の内容を分かりやすく要約してください。\n\n{file_text}"
                    elif mode == "❓ クイズを作る":
                        prompt = f"以下の内容からクイズを3問作り、解答もつけてください。\n\n{file_text}"
                    elif mode == "🔍 Google検索でリサーチ":
                        prompt = f"ユーザーの質問: {user_prompt}\n\n参考資料: {file_text}"
                        tools = [{"google_search": {}}]
                        
                    config = types.GenerateContentConfig(tools=tools, temperature=0.7)
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt,
                        config=config
                    )
                    
                    # 結果を保存
                    st.session_state.ai_result = response.text
                
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")

with col2:
    st.subheader("✨ 出力エリア")
    if st.session_state.ai_result:
        st.write(st.session_state.ai_result)
        
        st.markdown("---")
        st.write("⬇️ ファイルとしてダウンロード")
        
        # Wordダウンロードボタン
        word_data = create_word_file(st.session_state.ai_result)
        st.download_button(
            label="📄 Wordファイル (.docx) をダウンロード",
            data=word_data,
            file_name="ai_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        # PowerPointダウンロードボタン
        ppt_data = create_ppt_file(st.session_state.ai_result)
        st.download_button(
            label="📊 PowerPoint (.pptx) をダウンロード",
            data=ppt_data,
            file_name="ai_slides.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
