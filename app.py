import os
import json
import streamlit as st
from dotenv import load_dotenv

from azure.ai.contentunderstanding import ContentUnderstandingClient
from azure.core.credentials import AzureKeyCredential
from azure.identity import DefaultAzureCredential

load_dotenv()

st.set_page_config(page_title="Content Understanding - PPTX Demo", layout="wide")
st.title("Azure AI Content Understanding：PowerPoint（PPTX）読み込みデモ")

endpoint = os.getenv("CONTENTUNDERSTANDING_ENDPOINT", "").strip()
api_key = os.getenv("CONTENTUNDERSTANDING_KEY", "").strip()
analyzer_id = os.getenv("ANALYZER_ID", "prebuilt-documentSearch").strip()

with st.sidebar:
    st.header("設定")
    st.text_input("Endpoint", value=endpoint, key="endpoint")
    st.text_input("Analyzer ID", value=analyzer_id, key="analyzer_id")
    auth_mode = st.radio("認証", ["API Key", "DefaultAzureCredential"], index=0 if api_key else 1)
    st.caption("prebuilt-documentSearch は Markdown/要約/構造抽出が出るのでデモ向き。")

if not st.session_state.endpoint:
    st.error("CONTENTUNDERSTANDING_ENDPOINT を設定してください。")
    st.stop()

def make_client():
    if auth_mode == "API Key":
        if not api_key:
            st.error("CONTENTUNDERSTANDING_KEY が未設定です。")
            st.stop()
        return ContentUnderstandingClient(
            endpoint=st.session_state.endpoint,
            credential=AzureKeyCredential(api_key),
        )
    else:
        return ContentUnderstandingClient(
            endpoint=st.session_state.endpoint,
            credential=DefaultAzureCredential(),
        )

uploaded = st.file_uploader("PPTX をアップロード", type=["pptx"])

col1, col2 = st.columns([1, 1])
with col1:
    st.subheader("入力")
    if uploaded:
        st.write(f"- file: `{uploaded.name}`  size: {uploaded.size:,} bytes")
        st.download_button("アップロードしたPPTXをダウンロード", uploaded.getvalue(), file_name=uploaded.name)

with col2:
    st.subheader("実行")
    run = st.button("解析（Analyze）", type="primary", disabled=not uploaded)

if run:
    pptx_bytes = uploaded.getvalue()
    client = make_client()

    # PPTX の MIME type（Office Open XML Presentation）
    pptx_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    with st.spinner("解析中..."):
        # SDK: begin_analyze_binary が bytes を直接渡せる（LRO poller）:contentReference[oaicite:6]{index=6}
        poller = client.begin_analyze_binary(
            analyzer_id=st.session_state.analyzer_id,
            binary_input=pptx_bytes,
            content_type=pptx_mime,
        )
        result = poller.result()

    st.success("完了")

    # 返却は dict 互換（MutableMapping）:contentReference[oaicite:7]{index=7}
    st.subheader("Raw JSON（抜粋）")
    st.json(result)

    # 便利表示（よく使う場所だけ）
    st.subheader("抽出結果（見やすい表示）")
    contents = result.get("result", {}).get("contents", []) if "result" in result else result.get("contents", [])
    if not contents:
        st.warning("contents が空でした。PPTX が未対応扱いの可能性があるので、PPTX→PDF 変換投入を検討してください。")
        st.stop()

    for i, c in enumerate(contents, start=1):
        st.markdown(f"### Content #{i}: `{c.get('path', 'input')}`")
        md = c.get("markdown")
        if md:
            st.markdown("#### Markdown")
            st.markdown(md)
        fields = c.get("fields")
        if fields:
            st.markdown("#### Fields")
            st.code(json.dumps(fields, ensure_ascii=False, indent=2), language="json")
