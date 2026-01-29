import os
import io
import json
import base64
import tempfile
import subprocess
import time
import hashlib
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
status_area = st.container()

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

# --- LLM Semantic (Image-first) settings ---
ENV_SEMANTIC_ENABLE = os.getenv("SEMANTIC_ENABLE", "true").lower() in ("1", "true", "yes")
ENV_SEMANTIC_CONF_THRESHOLD = float(os.getenv("SEMANTIC_CONF_THRESHOLD", "0.75"))
ENV_SEMANTIC_UNKNOWN_THRESHOLD = int(os.getenv("SEMANTIC_UNKNOWN_THRESHOLD", "3"))
ENV_SEMANTIC_ZOOM = float(os.getenv("SEMANTIC_ZOOM", "2.5"))

NORMALIZED_SCHEMA = {
    "name": "normalized_slide",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "slide_number": {"type": "integer"},
            "title": {"type": "string"},
            "section": {"type": ["string", "null"]},
            "bullets": {"type": "array", "items": {"type": "string"}},
            "tables": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "caption": {"type": ["string", "null"]},
                        "headers": {"type": "array", "items": {"type": ["string", "null"]}},
                        "rows": {
                            "type": "array",
                            "items": {"type": "array", "items": {"type": ["string", "null"]}}
                        }
                    },
                    "required": ["caption", "headers", "rows"]
                }
            },
            "figures": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "caption": {"type": ["string", "null"]},
                        "labels": {"type": "array", "items": {"type": "string"}}
                    },
                    "required": ["caption", "labels"]
                }
            },
            "callouts": {"type": "array", "items": {"type": "string"}},
            "confidence": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "title": {"type": "number"},
                    "bullets": {"type": "number"},
                    "tables": {"type": "number"},
                    "figures": {"type": "number"}
                },
                "required": ["title", "bullets", "tables", "figures"]
            },
            "notes": {"type": ["string", "null"]}
        },
        "required": [
            "slide_number",
            "title",
            "section",
            "bullets",
            "tables",
            "figures",
            "callouts",
            "confidence",
            "notes"
        ]
    }
}

# --- Semantic architecture schema (image-first, CU fallback) ---
SEMANTIC_SCHEMA = {
    "name": "semantic_architecture",
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "doc_id": {"type": "string"},
            "page": {"type": "integer"},
            "title": {"type": ["string", "null"]},
            "overview": {"type": "string"},
            "layers": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "name": {"type": "string"},
                        "components": {"type": "array", "items": {"type": "string"}}
                    },
                    "required": ["name", "components"]
                }
            },
            "components": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "name": {"type": "string"},
                        "type": {"type": ["string", "null"]},
                        "role": {"type": ["string", "null"]},
                        "notes": {"type": ["string", "null"]}
                    },
                    "required": ["name", "type", "role", "notes"]
                }
            },
            "flows": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "source": {"type": "string"},
                        "target": {"type": "string"},
                        "direction": {"type": ["string", "null"], "enum": ["inbound", "outbound", "internal", None]},
                        "purpose": {"type": ["string", "null"]},
                        "data": {"type": ["string", "null"]}
                    },
                    "required": ["source", "target", "direction", "purpose", "data"]
                }
            },
            "boundaries": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "name": {"type": "string"},
                        "kind": {"type": ["string", "null"]},
                        "inside": {"type": "array", "items": {"type": "string"}},
                        "outside": {"type": "array", "items": {"type": "string"}},
                        "notes": {"type": ["string", "null"]}
                    },
                    "required": ["name", "kind", "inside", "outside", "notes"]
                }
            },
            "decisions": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "id": {"type": ["string", "null"]},
                        "topic": {"type": "string"},
                        "status": {"type": "string", "enum": ["decide", "consider", "unknown"]},
                        "options": {"type": "array", "items": {"type": "string"}},
                        "rationale": {"type": ["string", "null"]},
                        "open_questions": {"type": "array", "items": {"type": "string"}}
                    },
                    "required": ["id", "topic", "status", "options", "rationale", "open_questions"]
                }
            },
            "unknowns": {"type": "array", "items": {"type": "string"}},
            "assumptions": {"type": "array", "items": {"type": "string"}},
            "entities_from_cu": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "services": {"type": "array", "items": {"type": "string"}},
                    "products": {"type": "array", "items": {"type": "string"}},
                    "acronyms": {"type": "array", "items": {"type": "string"}},
                    "numbers": {"type": "array", "items": {"type": "string"}},
                    "other_terms": {"type": "array", "items": {"type": "string"}}
                },
                "required": ["services", "products", "acronyms", "numbers", "other_terms"]
            },
            "quality": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "confidence": {"type": "number"},
                    "unreadable_text_flags": {"type": "boolean"},
                    "needs_cu_fallback": {"type": "boolean"},
                    "reasons": {"type": "array", "items": {"type": "string"}}
                },
                "required": ["confidence", "unreadable_text_flags", "needs_cu_fallback", "reasons"]
            },
            "evidence": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "image_sha256": {"type": "string"},
                    "cu_present": {"type": "boolean"},
                    "cu_excerpt": {"type": "object", "additionalProperties": True}
                },
                "required": ["image_sha256", "cu_present", "cu_excerpt"]
            }
        },
        "required": [
            "doc_id",
            "page",
            "title",
            "overview",
            "layers",
            "components",
            "flows",
            "boundaries",
            "decisions",
            "unknowns",
            "assumptions",
            "entities_from_cu",
            "quality",
            "evidence"
        ]
    }
}

# Simple password gate (demo use)
if ENV_APP_PASSWORD:
    if "auth_ok" not in st.session_state:
        st.session_state.auth_ok = False
    if not st.session_state.auth_ok:
        status_area.info("パスワードが必要です。")
        with st.form("login_form"):
            pw = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login")
        if submitted:
            if pw == ENV_APP_PASSWORD:
                st.session_state.auth_ok = True
                st.rerun()
            else:
                status_area.error("パスワードが違います。")
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
            status_area.warning("Entra ID 認証では Translator のリソースエンドポイントを指定してください。")

    st.divider()
    st.subheader("LLM 再整理 (オプション)")
    refine_with_llm = st.checkbox("LLMで解析結果を再整理", value=ENV_LLM_REFINE)
    llm_image_only = st.checkbox("PNG画像のみでの抽出も実行", value=False)
    llm_send_image = st.checkbox("LLMに画像を送る", value=True)
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
    st.subheader("意味論ナレッジ (画像主・必要時CU補正)")
    semantic_enable = st.checkbox("意味論JSONを生成できるようにする", value=ENV_SEMANTIC_ENABLE)
    semantic_conf_threshold = st.slider("Confidence しきい値", min_value=0.5, max_value=0.95, value=float(ENV_SEMANTIC_CONF_THRESHOLD), step=0.01)
    semantic_unknown_threshold = st.slider("Unknowns しきい値", min_value=1, max_value=10, value=int(ENV_SEMANTIC_UNKNOWN_THRESHOLD), step=1)
    semantic_zoom = st.slider("画像ズーム (意味論抽出用)", min_value=1.5, max_value=4.0, value=float(ENV_SEMANTIC_ZOOM), step=0.1)

    st.divider()
    st.caption("PPTX はいったん PDF に変換し、PDF のページ(range)でスライド単位解析します。")
    st.caption("LibreOffice (soffice) が PATH に必要です。Docker 推奨。")

# Basic validation
if not endpoint:
    status_area.error("Endpoint を設定してください（.env の CU_ENDPOINT でも可）。")
    st.stop()

if auth_mode == "API Key" and not api_key:
    status_area.error("API Key を設定してください（.env の CU_KEY でも可）。")
    st.stop()

# ---------------------------
# Helpers
# ---------------------------
_CRED = DefaultAzureCredential()
_SCOPE = "https://cognitiveservices.azure.com/.default"


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


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


def llm_refine(
    slide_number: int,
    cu_result: dict | None,
    image_png: bytes,
    include_json: bool = True,
    include_image: bool = True,
) -> dict:
    """Azure OpenAI (Foundry) で JSON+画像 or 画像のみで再整理 (JSON返却)"""
    if not llm_endpoint or not llm_deployment:
        return {}

    base = llm_endpoint.rstrip("/")
    url = f"{base}/openai/deployments/{llm_deployment}/chat/completions"
    params = {"api-version": llm_api_version}

    if auth_mode == "API Key":
        if not api_key:
            return {}
        headers = {"api-key": api_key, "Content-Type": "application/json"}
    else:
        headers = _auth_headers()
        headers["Content-Type"] = "application/json"

    # CUの“使う部分”だけ渡す（トークン節約）
    cu_payload = {}
    if include_json and cu_result is not None:
        contents = cu_result.get("result", {}).get("contents") or cu_result.get("contents") or []
        if contents:
            c0 = contents[0]
            cu_payload = {
                "markdown": c0.get("markdown"),
                "fields": c0.get("fields"),
            }
        else:
            cu_payload = {"raw": cu_result}

    system = (
        "You are a careful slide reader. "
        "You will be given (1) a slide image and (2) extracted content from a document analyzer (may contain errors). "
        "If there is any conflict, prefer the slide image. "
        "If text is too small to read in the image, use the analyzer output as a fallback and reflect lower confidence. "
        "Return ONLY a JSON object that matches the provided JSON schema. "
        "Do not include markdown fences."
    )

    user_text = {
        "slide_number": slide_number,
        "task": "Normalize and correct the content into the target schema. Extract title, bullets, tables, figures, and callouts.",
        "analyzer_output": cu_payload,
    }

    image_b64 = base64.b64encode(image_png).decode("utf-8")
    messages = [
        {"role": "system", "content": f"{llm_system_prompt}\n\n{system}"},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": json.dumps(user_text, ensure_ascii=False)},
            ],
        },
    ]

    if include_image:
        messages[1]["content"].append(
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}}
        )

    body = {
        "messages": messages,
        "temperature": 0.2,
        "response_format": {"type": "json_schema", "json_schema": NORMALIZED_SCHEMA},
    }
    r = requests.post(url, params=params, headers=headers, json=body, timeout=180)
    r.raise_for_status()
    data = r.json()
    content = data["choices"][0]["message"]["content"]
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        return {"raw": content}


def semantic_llm_call(
    doc_id: str,
    page: int,
    image_png: bytes,
    cu_result: dict | None,
    use_cu: bool,
) -> dict:
    """
    Azure OpenAI で意味論ナレッジを生成。
    - use_cu=False: 画像のみ
    - use_cu=True : 画像 + CU excerpt（小さい文字の補強用）
    """
    if not llm_endpoint or not llm_deployment:
        return {}

    base = llm_endpoint.rstrip("/")
    url = f"{base}/openai/deployments/{llm_deployment}/chat/completions"
    params = {"api-version": llm_api_version}

    if auth_mode == "API Key":
        if not api_key:
            return {}
        headers = {"api-key": api_key, "Content-Type": "application/json"}
    else:
        headers = _auth_headers()
        headers["Content-Type"] = "application/json"

    image_b64 = base64.b64encode(image_png).decode("utf-8")

    # CU excerpt（トークン節約）
    cu_excerpt = {}
    if use_cu and cu_result is not None:
        contents = cu_result.get("result", {}).get("contents") or cu_result.get("contents") or []
        if contents:
            c0 = contents[0]
            cu_excerpt = {"markdown": c0.get("markdown"), "fields": c0.get("fields")}
        else:
            cu_excerpt = {"raw": cu_result}

    payload = {
        "doc_id": doc_id,
        "page": page,
        "mode": "image+cu" if use_cu else "image_only",
        "cu_excerpt": cu_excerpt if use_cu else None,
        "instructions": {
            "focus": [
                "components",
                "layers",
                "flows (arrows)",
                "trust/security/network boundaries",
                "numbered decision points (e.g., ①②③...)"
            ],
            "rules": [
                "Prefer image for structure and meaning.",
                "If text is too small to read, do NOT guess; add to unknowns and set unreadable_text_flags=true.",
                "Use CU excerpt only to improve accuracy of small text (names, acronyms, numbers).",
                "If there is conflict between image and CU, record in quality.reasons and lower confidence."
            ]
        }
    }

    system = (
        "You are an expert Azure architecture analyst. "
        "Return ONLY valid JSON matching the provided JSON schema. "
        "No markdown fences."
    )
    if not use_cu:
        system += " You only have a slide image."

    messages = [
        {"role": "system", "content": system},
        {"role": "user", "content": [{"type": "text", "text": json.dumps(payload, ensure_ascii=False)}]},
        {"role": "user", "content": [{"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}}]},
    ]

    body = {
        "messages": messages,
        "temperature": 0.2,
        "response_format": {"type": "json_schema", "json_schema": SEMANTIC_SCHEMA},
    }

    r = requests.post(url, params=params, headers=headers, json=body, timeout=180)
    r.raise_for_status()
    data = r.json()
    content = data["choices"][0]["message"]["content"]
    try:
        out = json.loads(content)
    except json.JSONDecodeError:
        out = {"raw": content}

    # evidence は常に付ける
    if isinstance(out, dict):
        out.setdefault("evidence", {})
        out["evidence"]["image_sha256"] = sha256_bytes(image_png)
        out["evidence"]["cu_present"] = bool(use_cu and cu_result is not None)
        out["evidence"]["cu_excerpt"] = cu_excerpt if (use_cu and cu_result is not None) else {}
        out.setdefault("entities_from_cu", {"services": [], "products": [], "acronyms": [], "numbers": [], "other_terms": []})
    return out


def should_fallback_to_cu(semantic: dict, conf_threshold: float, unknown_threshold: int) -> bool:
    if not isinstance(semantic, dict):
        return True
    q = semantic.get("quality", {}) or {}
    conf = float(q.get("confidence", 0.0))
    unreadable = bool(q.get("unreadable_text_flags", False))
    needs = bool(q.get("needs_cu_fallback", False))
    unknowns = semantic.get("unknowns", []) or []
    reasons = q.get("reasons", []) or []

    if conf < conf_threshold:
        return True
    if unreadable:
        return True
    if needs:
        return True
    if len(unknowns) >= unknown_threshold:
        return True
    # Reason hints
    reason_text = " ".join([str(r) for r in reasons]).lower()
    if any(k in reason_text for k in ["small", "unread", "table", "tiny", "cannot read", "illegible"]):
        return True
    return False


def build_semantic(
    doc_id: str,
    page: int,
    image_png: bytes,
    cu_result: dict | None,
    conf_threshold: float,
    unknown_threshold: int,
) -> dict:
    # Pass1: image-only
    s1 = semantic_llm_call(doc_id, page, image_png, cu_result=None, use_cu=False)
    if not isinstance(s1, dict):
        return {"raw": s1}

    # Ensure required fields exist (best-effort)
    s1.setdefault("doc_id", doc_id)
    s1.setdefault("page", page)
    s1.setdefault("layers", [])
    s1.setdefault("components", [])
    s1.setdefault("flows", [])
    s1.setdefault("boundaries", [])
    s1.setdefault("decisions", [])
    s1.setdefault("unknowns", [])
    s1.setdefault("assumptions", [])
    s1.setdefault("entities_from_cu", {"services": [], "products": [], "acronyms": [], "numbers": [], "other_terms": []})
    s1.setdefault("quality", {"confidence": 0.7, "unreadable_text_flags": False, "needs_cu_fallback": False, "reasons": []})
    s1.setdefault("evidence", {"image_sha256": sha256_bytes(image_png), "cu_present": False, "cu_excerpt": {}})

    # Pass2: CU fallback if needed and available
    if cu_result is not None and should_fallback_to_cu(s1, conf_threshold, unknown_threshold):
        s2 = semantic_llm_call(doc_id, page, image_png, cu_result=cu_result, use_cu=True)
        if isinstance(s2, dict):
            s2.setdefault("doc_id", doc_id)
            s2.setdefault("page", page)
            s2.setdefault("evidence", {})
            s2["evidence"].setdefault("image_sha256", sha256_bytes(image_png))
            s2["evidence"]["cu_present"] = True
            if "cu_excerpt" not in s2["evidence"]:
                contents = cu_result.get("result", {}).get("contents") or cu_result.get("contents") or []
                if contents:
                    c0 = contents[0]
                    s2["evidence"]["cu_excerpt"] = {
                        "markdown": c0.get("markdown"),
                        "fields": c0.get("fields"),
                    }
                else:
                    s2["evidence"]["cu_excerpt"] = {"raw": cu_result}
            return s2

    return s1


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
        percent = data.get("progress") or data.get("percentCompleted") or data.get("percentage")
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
    status_area.info("PPTXをアップロードしてください。")
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

        status_area.success(f"変換OK：{page_count} スライド（ページ）検出")

        # Initialize selection
        if "selected_page" not in st.session_state:
            st.session_state.selected_page = 1
        if "analyze_page" not in st.session_state:
            st.session_state.analyze_page = None
        if "analyzing" not in st.session_state:
            st.session_state.analyzing = False

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

            st.caption("スライドを選択後、解析ボタンを押してください。")
            if st.button("選択スライドを解析", disabled=st.session_state.analyzing):
                st.session_state.analyze_page = st.session_state.selected_page
                st.session_state.analyzing = True

            if st.session_state.analyze_page is None:
                status_area.info("解析するスライドを選んでボタンを押してください。")
                st.stop()

            # Cache per slide per uploaded file (basic cache key)
            file_sig = f"{uploaded.name}:{uploaded.size}"
            cache_key = (
                f"cu_result:{file_sig}:page:{st.session_state.analyze_page}"
                f":analyzer:{analyzer_id}:ver:{api_version}:auth:{auth_mode}"
            )

            if cache_key not in st.session_state:
                with st.spinner("Content Understanding 解析中..."):
                    st.session_state[cache_key] = cu_analyze_binary(
                        binary_bytes=pdf_bytes,
                        content_type="application/pdf",
                        page_1based=st.session_state.analyze_page,
                    )
                st.session_state.analyzing = False
            else:
                st.session_state.analyzing = False

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
                        status_area.warning("翻訳が有効ですが Translator Key が未設定です。")
                    else:
                        if md:
                            md = translate_texts([md])[0]
                        if fields:
                            fields = translate_fields(fields)

                if md:
                    st.markdown("### Markdown")
                    if "<table" in (md or "").lower():
                        st.markdown(md, unsafe_allow_html=True)
                    else:
                        st.markdown(md)

                if fields:
                    with st.expander("Fields (JSON)", expanded=False):
                        st.code(json.dumps(fields, ensure_ascii=False, indent=2), language="json")
            else:
                status_area.warning("contents が見つかりませんでした（Raw JSON を確認してください）。")

            with st.expander("Raw JSON", expanded=False):
                st.json(result)

            # --- LLM refine (existing) ---
            if refine_with_llm:
                if llm_image_only and not llm_send_image:
                    status_area.warning("画像のみ抽出は無効です。画像送信をオンにしてください。")
                if llm_image_only:
                    col_a, col_b = st.columns(2, gap="large")
                    with col_a:
                        llm_cache_key = f"llm:{cache_key}:dep:{llm_deployment}:ver:{llm_api_version}:json"
                        if llm_cache_key not in st.session_state:
                            with st.spinner("LLM で再整理中..."):
                                st.session_state[llm_cache_key] = llm_refine(
                                    st.session_state.selected_page,
                                    result,
                                    slide_png,
                                    include_json=True,
                                    include_image=llm_send_image,
                                )
                        llm_json = st.session_state[llm_cache_key]
                        if llm_json:
                            st.markdown("### LLM 再整理 (JSON + 画像)")
                            st.json(llm_json)
                        else:
                            status_area.warning("LLM 連携の設定が不足しています。")

                    with col_b:
                        llm_cache_key_img = f"llm:{cache_key}:dep:{llm_deployment}:ver:{llm_api_version}:img"
                        if llm_cache_key_img not in st.session_state:
                            with st.spinner("LLM で再整理中 (画像のみ)..."):
                                st.session_state[llm_cache_key_img] = llm_refine(
                                    st.session_state.selected_page,
                                    None,
                                    slide_png,
                                    include_json=False,
                                    include_image=llm_send_image,
                                )
                        llm_json_img = st.session_state[llm_cache_key_img]
                        if llm_json_img:
                            st.markdown("### LLM 再整理 (画像のみ)")
                            st.json(llm_json_img)
                else:
                    llm_cache_key = f"llm:{cache_key}:dep:{llm_deployment}:ver:{llm_api_version}:json"
                    if llm_cache_key not in st.session_state:
                        with st.spinner("LLM で再整理中..."):
                            st.session_state[llm_cache_key] = llm_refine(
                                st.session_state.selected_page,
                                result,
                                slide_png,
                                include_json=True,
                                include_image=llm_send_image,
                            )
                    llm_json = st.session_state[llm_cache_key]
                    if llm_json:
                        st.markdown("### LLM 再整理")
                        st.json(llm_json)
                    else:
                        status_area.warning("LLM 連携の設定が不足しています。")

            # --- Semantic knowledge (new) ---
            if semantic_enable and ENV_SEMANTIC_ENABLE:
                st.markdown("---")
                st.markdown("## 意味論ナレッジ（画像主・必要時CU補正）")
                st.caption("通常は画像のみで意味論JSONを生成し、信頼度が低い場合のみCUを使って補正します。")

                semantic_png = pdf_page_thumbnail_png(
                    pdf_path,
                    st.session_state.selected_page - 1,
                    zoom=float(semantic_zoom),
                )

                semantic_button = st.button("意味論JSONを生成", type="primary")
                if semantic_button:
                    sem_key = (
                        f"semantic:{file_sig}:page:{st.session_state.selected_page}"
                        f":dep:{llm_deployment}:ver:{llm_api_version}"
                        f":thr:{semantic_conf_threshold}:unk:{semantic_unknown_threshold}:z:{semantic_zoom}"
                    )
                    if sem_key not in st.session_state:
                        if not llm_endpoint or not llm_deployment:
                            status_area.warning("意味論生成には LLM Endpoint / Deployment の設定が必要です。")
                        else:
                            with st.spinner("意味論抽出中...（画像のみ→必要時CU補正）"):
                                st.session_state[sem_key] = build_semantic(
                                    doc_id=file_sig,
                                    page=st.session_state.selected_page,
                                    image_png=semantic_png,
                                    cu_result=result,
                                    conf_threshold=float(semantic_conf_threshold),
                                    unknown_threshold=int(semantic_unknown_threshold),
                                )

                    semantic_json = st.session_state.get(sem_key, {})
                    if semantic_json:
                        st.markdown("### 意味論JSON")
                        st.json(semantic_json)

                        # Show whether CU fallback was used (best-effort)
                        used_cu = False
                        if isinstance(semantic_json, dict):
                            ev = semantic_json.get("evidence", {}) or {}
                            used_cu = bool(ev.get("cu_present", False))
                        if used_cu:
                            st.success("CU 補正を実施しました（小さい文字・固有名詞の精度向上）。")
                        else:
                            st.info("画像のみで生成しました（CU補正なし）。")

                        with st.expander("意味論JSON (Raw)", expanded=False):
                            st.code(json.dumps(semantic_json, ensure_ascii=False, indent=2), language="json")
                    else:
                        status_area.warning("意味論JSONの生成に失敗しました（LLM設定やレスポンスを確認してください）。")

except FileNotFoundError:
    status_area.error(
        "LibreOffice (soffice) が見つかりません。\n\n"
        "対処:\n"
        "- ローカルに LibreOffice をインストールして `soffice` が PATH に入るようにする\n"
        "- もしくは Docker で LibreOffice 同梱イメージを使う（推奨）"
    )
except requests.HTTPError as e:
    status_area.error(f"HTTPエラー: {e}\n\nレスポンス: {getattr(e.response, 'text', '')}")
except Exception as e:
    status_area.error(f"エラー: {e}")
