"""
app.py -- Streamlit front-end for the SlideMorph pipeline.

Wraps generate_ai_deck() from fill_template.py in a small web UI:
  - topic / slide-count / template inputs
  - Generate button -> downloadable .pptx
  - simple in-memory rate limit: MAX_PER_HOUR generations per client

API key is read via os.getenv("GEMINI_API_KEY"). Locally we load .env into the
env with python-dotenv; on Render the platform sets the env var directly.

Run with:  streamlit run app.py
"""

import os
import time
import threading
import traceback
import uuid
from collections import defaultdict, deque

import streamlit as st
from dotenv import load_dotenv

from fill_template import generate_ai_deck, output_filename

load_dotenv()  # local dev: pull GEMINI_API_KEY from .env into os.environ

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(HERE, "templates")
OUTPUT_DIR = os.path.join(HERE, "generated")
os.makedirs(OUTPUT_DIR, exist_ok=True)

MAX_PER_HOUR = 3
WINDOW_SECONDS = 3600


def list_templates():
    """Every subfolder of templates/ that has a spec.json + template file."""
    names = []
    if os.path.isdir(TEMPLATES_DIR):
        for name in sorted(os.listdir(TEMPLATES_DIR)):
            spec = os.path.join(TEMPLATES_DIR, name, "spec.json")
            if os.path.isfile(spec):
                names.append(name)
    return names or ["neon"]


@st.cache_resource
def _rate_state():
    """Process-wide rate-limit store, shared across sessions and reruns."""
    return {"lock": threading.Lock(), "log": defaultdict(deque)}


def client_id():
    """Best-effort caller identity for rate limiting.

    Streamlit has no equivalent of Gradio's request.client.host, so we read the
    forwarded-for header set by Render's proxy and fall back to a per-session
    UUID when running locally (where no proxy header exists).
    """
    try:
        headers = st.context.headers or {}
        fwd = headers.get("X-Forwarded-For") or headers.get("x-forwarded-for")
        if fwd:
            return fwd.split(",")[0].strip()
    except Exception:  # noqa: BLE001 - header access varies by Streamlit version
        pass
    if "_client_id" not in st.session_state:
        st.session_state["_client_id"] = str(uuid.uuid4())
    return st.session_state["_client_id"]


def rate_check(cid):
    """Return (allowed, wait_seconds). Prunes entries older than WINDOW_SECONDS."""
    state = _rate_state()
    now = time.time()
    with state["lock"]:
        q = state["log"][cid]
        while q and now - q[0] > WINDOW_SECONDS:
            q.popleft()
        if len(q) >= MAX_PER_HOUR:
            return False, int(WINDOW_SECONDS - (now - q[0]))
        q.append(now)
        return True, 0


def generate(topic, count, template):
    """Run the pipeline. Returns (out_path, download_name, error_message); the
    first two are None on failure. `download_name` is the clean, user-facing
    '{topic}_{template}.pptx'; the on-disk name is prefixed with a timestamp so
    concurrent generations never overwrite each other."""
    download_name = output_filename(topic, template)
    disk_name = f"{int(time.time())}_{download_name}"
    out_path = os.path.join(OUTPUT_DIR, disk_name)
    try:
        generate_ai_deck(template, topic, int(count), out_path)
    except Exception as exc:  # noqa: BLE001 - surface any pipeline error to the UI
        traceback.print_exc()
        return None, None, f"{type(exc).__name__}: {exc}"
    return out_path, download_name, None


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.set_page_config(page_title="SlideMorph", page_icon="📊")

st.title("SlideMorph — AI Presentations That Actually Animate")
st.write("Enter a topic, pick a length and template, and get a "
         "morph-animated PowerPoint deck.")

templates = list_templates()

topic = st.text_input("Topic", placeholder="e.g. Machine Learning")
count = st.slider("Slide count", min_value=6, max_value=30, value=13, step=1)
template = st.selectbox("Template", options=templates)

if st.button("Generate", type="primary"):
    if not topic.strip():
        st.error("Please enter a topic.")
    else:
        allowed, wait = rate_check(client_id())
        if not allowed:
            mins = max(1, wait // 60)
            st.error(f"Rate limit hit. You have used {MAX_PER_HOUR} generations "
                     f"this hour. Try again in ~{mins} min.")
        else:
            with st.spinner(f"Generating {count} slides on {topic.strip()!r}… "
                            "this takes a minute."):
                out_path, download_name, error = generate(topic.strip(), count, template)
            if error:
                st.session_state.pop("result", None)
                st.error(f"Generation failed: {error}")
            else:
                # Read now so the download survives later reruns even if the
                # file is cleaned up.
                with open(out_path, "rb") as f:
                    st.session_state["result"] = {
                        "data": f.read(),
                        "name": download_name,
                        "count": count,
                        "topic": topic.strip(),
                    }

# Persist the download across the rerun that clicking the button triggers.
result = st.session_state.get("result")
if result:
    st.success(f"Done — {result['count']} slides on {result['topic']!r}.")
    st.download_button(
        label="Download .pptx",
        data=result["data"],
        file_name=result["name"],
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
