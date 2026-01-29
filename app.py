import os
import io
import json
import base64
import tempfile
import subprocess
import time
from pathlib import Path

import fitz  # PyMuPDF
import requests
import streamlit as st
from dotenv import load_dotenv
from azure.identity import DefaultAzureCredential

# ---------------------------
# App setup
# ---------------------------
load_dotenv()

st.set_page_config(page_title="CU PPTX Slide Split Demo", layout="wide")
st.title("Content Understanding：PPTX → スライド単位解析 → 一覧UI")

# Defaults from env
ENV_CU_ENDPOINT = os.getenv("CU_ENDPOINT", "").rstrip("/")
ENV_CU_KEY = os.getenv("CU_KEY", "")
ENV_ANALYZER_ID = os.getenv("ANALYZER_ID", "prebuilt-documentSearch")
ENV_API_VERSION = os.getenv("API_VERSION", "2025-11-01")
ENV_TRANSLATOR_ENDPOINT = os.getenv("TRANSLATOR_ENDPOINT", "").rstrip("/")
ENV_TRANSLATOR_KEY = os.getenv("TRANSLATOR_KEY", "")
ENV_TRANSLATOR_REGION = os.getenv("TRANSLATOR_REGION", "")
ENV_LLM_ENDPOINT = os.getenv("LLM_ENDPOINT", "").rstrip("/")
ENV_LLM_DEPLOYMENT = os.getenv("LLM_DEPLOYMENT", "")
ENV_LLM_API_VERSION = os.getenv("LLM_API_VERSION", "2024-06-01")
ENV_AUTH_MODE = os.getenv("AUTH_MODE", "entra").lower()
ENV_LOCALE = os.getenv("LOCALE", "ja-JP")
ENV_TRANSLATE_TO_JA = os.getenv("TRANSLATE_TO_JA", "false").lower() in ("1", "true", "yes")
ENV_LLM_REFINE = os.getenv("LLM_REFINE", "false").lower() in ("1", "true", "yes")
ENV_APP_PASSWORD = os.getenv("APP_PASSWORD", "")

# Simple password gate (demo use)
if ENV_APP_PASSWORD:
    if "auth_ok" not in st.session_state:
        st.session_state.auth_ok = False
    if not st.session_state.auth_ok:
        st.info("パスワードが必要です。")
        pw = st.text_input("Password", type="password")
        if st.button("Login"):
            if pw == ENV_APP_PASSWORD:
                st.session_state.auth_ok = True
            else:
                st.error("パスワードが違います。")
        st.stop()

# ---------------------------
# Sidebar (settings)
# ---------------------------
with st.sidebar:
    st.header("設定")

    endpoint = st.text_input("Endpoint", value=ENV_CU_ENDPOINT, help="例: https://<resource>.cognitiveservices.azure.com")
    analyzer_id = st.text_input("Analyzer ID", value=ENV_ANALYZER_ID, help="例: prebuilt-documentSearch")
    api_version = st.text_input("API Version", value=ENV_API_VERSION, help="例: 2025-11-01")
    locale_options = ["ja-JP", "en-US"]
    locale = st.selectbox(
        "Locale",
        locale_options,
        index=locale_options.index(ENV_LOCALE) if ENV_LOCALE in locale_options else 0,
        help="解析言語の指定（対応はアナライザーによって異なります）",
    )
    auth_index = 0 if ENV_AUTH_MODE in ("entra", "entra_id", "aad") else 1
    auth_mode = st.radio(
        "認証方式",
        ["DefaultAzureCredential (Entra ID)", "API Key"],
        index=auth_index,
        help="Key が無効化されている場合は DefaultAzureCredential を選択"
    )

    api_key = ""
    if auth_mode == "API Key":
        api_key = st.text_input("API Key", value=ENV_CU_KEY, type="password")
        st.caption("Key 認証: Ocp-Apim-Subscription-Key を使用")
    else:
        st.caption("Entra ID 認証: Authorization: Bearer <token> を使用")
        st.caption("ローカル実行なら `az login` 済みだと通りやすいです。")

    st.divider()
    st.subheader("翻訳 (オプション)")
    translate_to_ja = st.checkbox("結果を日本語に翻訳", value=ENV_TRANSLATE_TO_JA)
    translator_endpoint = st.text_input(
        "Translator Endpoint",
        value=ENV_TRANSLATOR_ENDPOINT or "https://api.cognitive.microsofttranslator.com",
        help="Key 認証: https://api.cognitive.microsofttranslator.com / Entra ID: https://<resource>.cognitiveservices.azure.com",
    )
    if auth_mode == "API Key":
        translator_key = st.text_input("Translator Key", value=ENV_TRANSLATOR_KEY, type="password")
        translator_region = st.text_input(
            "Translator Region",
            value=ENV_TRANSLATOR_REGION,
            help="マルチリージョンのキーの場合は不要。リージョンキーなら必須",
        )
    else:
        translator_key = ""
        translator_region = ""
        st.caption("Translator も Entra ID 認証で呼び出します。")
        if "cognitive.microsofttranslator.com" in translator_endpoint:
            st.warning("Entra ID 認証では Translator のリソースエンドポイントを指定してください。")

    st.divider()
    st.subheader("LLM 再整理 (オプション)")
    refine_with_llm = st.checkbox("LLMで解析結果を再整理", value=ENV_LLM_REFINE)
    llm_endpoint = st.text_input(
        "LLM Endpoint",
        value=ENV_LLM_ENDPOINT,
        help="例: https://<resource>.openai.azure.com  または https://<resource>.services.ai.azure.com",
    )
    llm_deployment = st.text_input("LLM Deployment", value=ENV_LLM_DEPLOYMENT, help="例: gpt-4o-mini")
    llm_api_version = st.text_input("LLM API Version", value=ENV_LLM_API_VERSION)
    llm_system_prompt = st.text_area(
        "System Prompt",
        value="You are a careful analyst. Use the provided slide image and extracted JSON to produce a concise, accurate Japanese summary. If something is unclear, say so.",
        height=120,
    )

    st.divider()
    st.caption("PPTX はいったん PDF に変換し、PDF のページ(range)でスライド単位解析します。")
    st.caption("LibreOffice (soffice) が PATH に必要です。Docker 推奨。")

# Basic validation
if not endpoint:
    st.error("Endpoint を設定してください（.env の CU_ENDPOINT でも可）。")
    st.stop()

if auth_mode == "API Key" and not api_key:
    st.error("API Key を設定してください（.env の CU_KEY でも可）。")
    st.stop()

# ---------------------------
# Helpers
# ---------------------------
_CRED = DefaultAzureCredential()
_SCOPE = "https://cognitiveservices.azure.com/.default"


def convert_pptx_to_pdf(pptx_bytes: bytes, out_dir: Path) -> Path:
    """LibreOffice headless で PPTX→PDF 変換"""
    pptx_path = out_dir / "input.pptx"
    pptx_path.write_bytes(pptx_bytes)

    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(pptx_path),
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice変換に失敗:\nSTDERR:\n{proc.stderr}\nSTDOUT:\n{proc.stdout}")

    pdf_path = out_dir / "input.pdf"
    if not pdf_path.exists():
        pdfs = list(out_dir.glob("*.pdf"))
        if not pdfs:
            raise RuntimeError("PDFが生成されませんでした。")
        pdf_path = pdfs[0]
    return pdf_path


def pdf_page_thumbnail_png(pdf_path: Path, page_index: int, zoom: float = 1.0) -> bytes:
    """PDFの特定ページをPNGにして返す（UIサムネ用）"""
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_index)
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    doc.close()
    return pix.tobytes("png")


def _auth_headers() -> dict:
    """認証ヘッダーを返す（Key or Bearer）"""
    if auth_mode == "API Key":
        return {"Ocp-Apim-Subscription-Key": api_key}
    token = _CRED.get_token(_SCOPE).token
    return {"Authorization": f"Bearer {token}"}


def translate_texts(texts: list[str], to_lang: str = "ja") -> list[str]:
    """Translator Text API で配列翻訳"""
    if not translator_endpoint:
        return texts

    base = translator_endpoint.rstrip("/")
    if base.endswith("/translate"):
        url = base
    else:
        if auth_mode == "API Key":
            url = f"{base}/translate"
        else:
            if "/translator/text" in base:
                url = f"{base}/translate"
            else:
                url = f"{base}/translator/text/v3.0/translate"
    params = {"api-version": "3.0", "to": to_lang}
    if auth_mode == "API Key":
        if not translator_key:
            return texts
        headers = {
            "Ocp-Apim-Subscription-Key": translator_key,
            "Content-Type": "application/json",
        }
        if translator_region:
            headers["Ocp-Apim-Subscription-Region"] = translator_region
    else:
        headers = _auth_headers()
        headers["Content-Type"] = "application/json"

    body = [{"text": t} for t in texts]
    r = requests.post(url, params=params, headers=headers, json=body, timeout=180)
    r.raise_for_status()
    data = r.json()
    return [item["translations"][0]["text"] for item in data]


def translate_fields(obj):
    """dict/list 内の文字列だけ翻訳"""
    if isinstance(obj, str):
        return translate_texts([obj])[0]
    if isinstance(obj, list):
        return [translate_fields(v) for v in obj]
    if isinstance(obj, dict):
        return {k: translate_fields(v) for k, v in obj.items()}
    return obj


def llm_refine(content_json: dict, image_png: bytes) -> str:
    """Azure OpenAI (Foundry) で JSON+画像を使って再整理"""
    if not llm_endpoint or not llm_deployment:
        return ""

    base = llm_endpoint.rstrip("/")
    url = f"{base}/openai/deployments/{llm_deployment}/chat/completions"
    params = {"api-version": llm_api_version}

    if auth_mode == "API Key":
        if not api_key:
            return ""
        headers = {"api-key": api_key, "Content-Type": "application/json"}
    else:
        headers = _auth_headers()
        headers["Content-Type"] = "application/json"

    image_b64 = base64.b64encode(image_png).decode("utf-8")
    messages = [
        {"role": "system", "content": llm_system_prompt},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": "以下は Content Understanding の解析結果(JSON)です。"},
                {"type": "text", "text": json.dumps(content_json, ensure_ascii=False)},
                {"type": "text", "text": "以下は同じスライドの画像です。"},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}},
            ],
        },
    ]

    body = {"messages": messages, "temperature": 0.2}
    r = requests.post(url, params=params, headers=headers, json=body, timeout=180)
    r.raise_for_status()
    data = r.json()
    return data["choices"][0]["message"]["content"]


def cu_analyze_binary(binary_bytes: bytes, content_type: str, page_1based: int) -> dict:
    """
    Content Understanding AnalyzeBinary を range=page で1ページずつ解析（LRO）
    - Key 認証 or Entra ID 認証のどちらでも動くようにする
    """
    url = f"{endpoint}/contentunderstanding/analyzers/{analyzer_id}:analyzeBinary"
    params = {"api-version": api_version, "range": str(page_1based)}
    if locale:
        params["locale"] = locale

    headers = _auth_headers()
    headers["Content-Type"] = content_type

    # analyzeBinary のボディはバイナリ
    r = requests.post(url, params=params, headers=headers, data=binary_bytes, timeout=180)
    r.raise_for_status()

    op_loc = r.headers.get("Operation-Location")
    if not op_loc:
        raise RuntimeError("Operation-Location が見つかりません。")

    # Poll
    progress = st.progress(0)
    status_text = st.empty()
    ticks = 0
    poll_headers = _auth_headers()
    while True:
        rr = requests.get(op_loc, headers=poll_headers, timeout=180)
        rr.raise_for_status()
        data = rr.json()
        status = data.get("status")
        percent = (
            data.get("progress")
            or data.get("percentCompleted")
            or data.get("percentage")
        )
        if isinstance(percent, (int, float)):
            progress.progress(min(max(int(percent), 0), 100))
        else:
            ticks = (ticks + 1) % 100
            progress.progress(ticks)
        status_text.write(f"解析ステータス: {status}")
        if status in ("Succeeded", "Failed", "Canceled"):
            if status != "Succeeded":
                progress.empty()
                status_text.empty()
                raise RuntimeError(f"解析失敗: {json.dumps(data, ensure_ascii=False)}")
            progress.empty()
            status_text.empty()
            return data
        time.sleep(0.4)


# ---------------------------
# UI - upload
# ---------------------------
uploaded = st.file_uploader("PPTXをアップロード", type=["pptx"])

if not uploaded:
    st.info("PPTXをアップロードしてください。")
    st.stop()

st.write(f"file: `{uploaded.name}`  size: {uploaded.size:,} bytes")

# Layout columns
colL, colR = st.columns([0.35, 0.65], gap="large")

# ---------------------------
# Convert PPTX->PDF and prepare thumbnails
# ---------------------------
try:
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)

        with st.spinner("PPTX→PDF 変換中..."):
            pdf_path = convert_pptx_to_pdf(uploaded.getvalue(), tmp)

        pdf_bytes = pdf_path.read_bytes()
        doc = fitz.open(pdf_path)
        page_count = doc.page_count
        doc.close()

        st.success(f"変換OK：{page_count} スライド（ページ）検出")

        # Initialize selection
        if "selected_page" not in st.session_state:
            st.session_state.selected_page = 1

        # Left: slide list with thumbnails
        with colL:
            st.subheader("Slides")
            for p in range(1, page_count + 1):
                thumb = pdf_page_thumbnail_png(pdf_path, p - 1, zoom=1.0)

                c1, c2 = st.columns([0.45, 0.55], gap="small")
                with c1:
                    st.image(thumb, width="stretch")
                with c2:
                    label = f"Slide {p}"
                    if st.button(label, key=f"btn_{p}"):
                        st.session_state.selected_page = p

        # Right: selected slide analysis result
        with colR:
            st.subheader(f"Slide {st.session_state.selected_page} Result")

            # Cache per slide per uploaded file (basic cache key)
            file_sig = f"{uploaded.name}:{uploaded.size}"
            cache_key = f"cu_result:{file_sig}:page:{st.session_state.selected_page}:analyzer:{analyzer_id}:ver:{api_version}:auth:{auth_mode}"

            if cache_key not in st.session_state:
                with st.spinner("Content Understanding 解析中..."):
                    st.session_state[cache_key] = cu_analyze_binary(
                        binary_bytes=pdf_bytes,
                        content_type="application/pdf",
                        page_1based=st.session_state.selected_page,
                    )

            result = st.session_state[cache_key]
            slide_png = pdf_page_thumbnail_png(pdf_path, st.session_state.selected_page - 1, zoom=2.0)

            # Friendly view
            contents = result.get("result", {}).get("contents") or result.get("contents") or []
            if contents:
                c = contents[0]
                md = c.get("markdown")
                fields = c.get("fields")

                if translate_to_ja:
                    if auth_mode == "API Key" and not translator_key:
                        st.warning("翻訳が有効ですが Translator Key が未設定です。")
                    else:
                        if md:
                            md = translate_texts([md])[0]
                        if fields:
                            fields = translate_fields(fields)

                if md:
                    st.markdown("### Markdown")
                    if "<table" in md.lower():
                        st.markdown(md, unsafe_allow_html=True)
                    else:
                        st.markdown(md)

                if fields:
                    st.markdown("### Fields (JSON)")
                    st.code(json.dumps(fields, ensure_ascii=False, indent=2), language="json")
            else:
                st.warning("contents が見つかりませんでした（Raw JSON を確認してください）。")

            if refine_with_llm:
                llm_cache_key = f"llm:{cache_key}:dep:{llm_deployment}:ver:{llm_api_version}"
                if llm_cache_key not in st.session_state:
                    with st.spinner("LLM で再整理中..."):
                        st.session_state[llm_cache_key] = llm_refine(result, slide_png)
                llm_text = st.session_state[llm_cache_key]
                if llm_text:
                    st.markdown("### LLM 再整理")
                    st.markdown(llm_text)
                else:
                    st.warning("LLM 連携の設定が不足しています。")

            with st.expander("Raw JSON", expanded=False):
                st.json(result)

except FileNotFoundError:
    st.error(
        "LibreOffice (soffice) が見つかりません。\n\n"
        "対処:\n"
        "- ローカルに LibreOffice をインストールして `soffice` が PATH に入るようにする\n"
        "- もしくは Docker で LibreOffice 同梱イメージを使う（推奨）"
    )
except requests.HTTPError as e:
    st.error(f"HTTPエラー: {e}\n\nレスポンス: {getattr(e.response, 'text', '')}")
except Exception as e:
    st.error(f"エラー: {e}")
