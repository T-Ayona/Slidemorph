"""
app.py -- Gradio front-end for the SlideMorph pipeline.

Wraps generate_ai_deck() from fill_template.py in a small web UI:
  - topic / slide-count / template inputs
  - Generate button -> downloadable .pptx
  - simple in-memory rate limit: MAX_PER_HOUR generations per client IP

API key is read via os.getenv("GEMINI_API_KEY"). Locally we load .env into the
env with python-dotenv; on Render the platform sets the env var directly.
"""

import os
import time
import threading
import traceback
from collections import defaultdict, deque

import gradio as gr
from dotenv import load_dotenv

from fill_template import generate_ai_deck

load_dotenv()  # local dev: pull GEMINI_API_KEY from .env into os.environ

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(HERE, "templates")
OUTPUT_DIR = os.path.join(HERE, "generated")
os.makedirs(OUTPUT_DIR, exist_ok=True)

MAX_PER_HOUR = 3
WINDOW_SECONDS = 3600
_rate_lock = threading.Lock()
_rate_log = defaultdict(deque)   # ip -> deque of unix timestamps


def list_templates():
    """Every subfolder of templates/ that has a spec.json + template file."""
    names = []
    if os.path.isdir(TEMPLATES_DIR):
        for name in sorted(os.listdir(TEMPLATES_DIR)):
            spec = os.path.join(TEMPLATES_DIR, name, "spec.json")
            if os.path.isfile(spec):
                names.append(name)
    return names or ["neon"]


def _rate_check(ip):
    """Return (allowed, wait_seconds). Prunes entries older than WINDOW_SECONDS."""
    now = time.time()
    with _rate_lock:
        q = _rate_log[ip]
        while q and now - q[0] > WINDOW_SECONDS:
            q.popleft()
        if len(q) >= MAX_PER_HOUR:
            return False, int(WINDOW_SECONDS - (now - q[0]))
        q.append(now)
        return True, 0


def generate(topic, count, template, request: gr.Request):
    """Handler for the Generate button. Returns (file, status_markdown)."""
    ip = request.client.host if request and request.client else "unknown"

    topic = (topic or "").strip()
    if not topic:
        return None, "**Error:** please enter a topic."

    ok, wait = _rate_check(ip)
    if not ok:
        mins = max(1, wait // 60)
        return None, (f"**Rate limit hit.** You have used {MAX_PER_HOUR} "
                      f"generations this hour. Try again in ~{mins} min.")

    safe_topic = "".join(c if c.isalnum() else "_" for c in topic)[:40] or "deck"
    out_name = f"{template}_{safe_topic}_{int(time.time())}.pptx"
    out_path = os.path.join(OUTPUT_DIR, out_name)

    try:
        generate_ai_deck(template, topic, int(count), out_path)
    except Exception as exc:  # noqa: BLE001 - surface any pipeline error to the UI
        traceback.print_exc()
        return None, f"**Generation failed:** `{type(exc).__name__}: {exc}`"

    return out_path, f"**Done** -- {int(count)} slides on {topic!r}. Click the file to download."


with gr.Blocks(title="SlideMorph") as demo:
    gr.Markdown("# SlideMorph -- AI Presentations That Actually Animate")
    gr.Markdown("Enter a topic, pick a length and template, and get a "
                "morph-animated PowerPoint deck.")

    with gr.Row():
        with gr.Column(scale=2):
            topic_in = gr.Textbox(label="Topic",
                                   placeholder="e.g. Machine Learning")
            count_in = gr.Slider(minimum=6, maximum=30, step=1, value=13,
                                  label="Slide count")
            template_in = gr.Dropdown(choices=list_templates(),
                                       value=list_templates()[0],
                                       label="Template")
            go_btn = gr.Button("Generate", variant="primary")

        with gr.Column(scale=1):
            file_out = gr.File(label="Generated .pptx")
            status_out = gr.Markdown()

    go_btn.click(fn=generate,
                  inputs=[topic_in, count_in, template_in],
                  outputs=[file_out, status_out])


if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0",
                server_port=int(os.environ.get("PORT", 7860)))
