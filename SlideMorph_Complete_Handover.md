# SlideMorph — Complete Project Handover Document

**Purpose of this document:** If you start a new chat or come back later, paste this whole document in and continue. It contains everything: the idea, the current status, all technical decisions, all prompts, and the full roadmap.

---

## PART 1 — WHO I AM AND WHAT I'M BUILDING

I'm a BS Computer Science student at NUML University in Pakistan. I'm building a micro-SaaS solo, with no funding, no connections, on a 1-month Claude Pro subscription. I have a Gemini API key.

**Realistic goals:**
- Ship a working prototype visible internationally
- Get accepted to Pakistan's NIC Islamabad incubator (grant + AWS credits)
- Build an open-source library credential that could land me a job or contract
- Earn a few hundred to few thousand dollars via templates + service
- Not: beat Gamma, raise millions, get acquired by Microsoft

**The product:** SlideMorph — an AI tool that generates native, editable PowerPoint decks that actually animate (using Morph transitions). Uses hand-designed animated templates. The AI never designs anything — it only fills placeholders with short text and picks which template variants to use.

**Core principle:** Fill, don't generate. The design, shapes, animations, and transitions live inside the template file. Code only swaps text, picks variants, and deletes unused slides.

**Two products, one codebase:**
- **Product A:** Open-source Python library `pptx-morph-safe` — the technical credential
- **Product B:** SaaS/demo on Hugging Face Spaces — the pitch for funders and users

---

## PART 2 — TECHNICAL DECISIONS ALREADY MADE

**Stack (all free):**
- Language: Python 3.11+
- PPTX engine: `python-pptx` + `lxml` (for XML manipulation python-pptx can't do)
- Text-fit measurement: `fontTools` + `Pillow`
- LLM: Gemini (currently using `gemini-2.0-flash`, free tier)
- Web UI: Gradio
- Hosting: Hugging Face Spaces (free, always-live public URL)
- Payments (later): Lemon Squeezy (merchant of record — solves Pakistan/Stripe problem)
- Template packs (later): Gumroad
- Version control: Git + GitHub

**Project location:** `C:\slidemorph` (moved out of OneDrive because OneDrive kept dehydrating files and breaking builds)

**Files in C:\slidemorph:**
- `fill_template.py` — the main script (working)
- `Presentation_templete_neon.pptx` — the template with animations
- `.env` — contains `GEMINI_API_KEY=...`
- `.gitignore` — excludes venv/, .env, __pycache__, *.pptx (except template)
- `venv/` — Python virtual environment
- Various `output_*.pptx` test files

**Python packages installed:**
```
python-pptx, lxml, fonttools, pillow, google-generativeai, gradio, requests, python-dotenv
```

---

## PART 3 — THE TEMPLATE STRUCTURE

Current template `Presentation_templete_neon.pptx` has 20 slides:
- Slide 1: Start (title + subtitle + presenter info)
- Slides 2-4: Wheel variants W1, W2, W3 (3 tags around wheel + content box)
- Slides 5-13: Bar variants B1-B9 (title + bullet content)
- Slides 14-19: Hexagon variants H1-H6 (title + bullet content)
- Slide 20: End slide

**All slides have Morph transition** (`byObject`, 2 seconds).

**Manually applied entrance animations** on AI_content shapes (fade, after previous, "by paragraph" so each bullet fades separately).

**Placeholder naming convention:**
- `AI_title`, `AI_subtitle` — filled with generated text
- `AI_content` — filled with bullet points
- `AI_tag_1`, `AI_tag_2`, `AI_tag_3` — wheel labels
- `USER_info` — presenter info, DO NOT TOUCH
- `end_textbox` — end slide, DO NOT TOUCH

**Morph safety:** All AI shapes get renamed to `!!AI_*` prefix so PowerPoint's Morph matches them by name across slides. Each slide's shape gets a unique name (like `!!AI_content_5`) to prevent unwanted cross-slide morphing of text boxes.

**Autofit:** Disabled (`noAutofit`) on all AI content boxes — otherwise Morph animates box size changes.

---

## PART 4 — SLIDE SELECTION RULES (COMPLETE)

Users can request 6 to 30 slides. Reject anything outside that range.

Every deck always has: 1 Start + 3 Wheel + End.

**For 6-20 slides, exact bar and hexagon counts:**
```
6:  1B, 0H
7:  2B, 0H
8:  3B, 0H
9:  3B, 1H
10: 3B, 2H
11: 3B, 3H
12: 4B, 3H
13: 4B, 4H
14: 5B, 4H
15: 5B, 5H
16: 6B, 5H
17: 6B, 6H
18: 7B, 6H
19: 8B, 6H
20: 9B, 6H (full original template)
```

Order: Start → W1 W2 W3 → all Bars → all Hexagons → End

**For 21-30 slides:**
Use full 20-slide template first, then insert extra slides between last hexagon (H6) and End slide. Extra slides alternate in blocks of 3-4:
- Block 1: 3-4 bars (start over from B1)
- Block 2: 3-4 hexagons (start over from H1)
- Block 3: bars again (continue where left off)
- Etc.

Wheel never loops. Only Bars and Hexagons loop.

---

## PART 5 — CURRENT STATUS (COMPLETED)

- ✅ Text filling with formatting preserved (no font/color loss)
- ✅ Morph transitions preserved after filling
- ✅ Manual fade animations on bullets preserved
- ✅ Slide selection logic for 6-30 slides
- ✅ Looping with alternating bars/hexagon blocks after 20 slides
- ✅ Gemini AI generating real content from topic prompt
- ✅ Adaptive bullet rules (fewer bullets if they wrap to multiple lines)
- ✅ Heading character limits enforced (single line only)
- ✅ Hanging indent on wrapped bullets

**End-to-end pipeline works from command line.** Topic in → real .pptx out.

---

## PART 6 — REMAINING WORK (4 STEPS)

Use Claude Code (desktop app → Code tab → point to `C:\slidemorph`) for all four steps.

### STEP 1 — JSON SPEC SYSTEM (2-3 days)

**Why:** Currently rules are hardcoded. Extract them to JSON so new templates can be added without code changes. Also makes the library reusable.

**Paste this prompt into Claude Code:**

```
The script currently has all rules hardcoded for the neon template. I need to refactor 
it so all template-specific rules live in a JSON spec file. This way, adding a new 
template later means just adding a JSON + a pptx — no code changes.

DO THIS:

1. Create a folder structure:
   C:\slidemorph\templates\neon\
     template.pptx  (copy the current template file here)
     spec.json      (create this new file — see structure below)

2. spec.json structure:
{
  "name": "Neon",
  "description": "Purple neon theme with wheel, bars, and hexagon designs",
  "file": "template.pptx",
  "total_slides": 20,
  "sections": {
    "start": {"slide": 1},
    "wheel": {"slides": [2, 3, 4], "count": 3},
    "bars": {"slides": [5, 6, 7, 8, 9, 10, 11, 12, 13], "count": 9},
    "hexagon": {"slides": [14, 15, 16, 17, 18, 19], "count": 6},
    "end": {"slide": 20}
  },
  "placeholders": {
    "start": {
      "AI_title": {"max_chars": 20, "single_line": true},
      "AI_subtitle": {"max_chars": 30, "single_line": true}
    },
    "wheel": {
      "AI_tag_1": {"max_chars": 14, "single_line": true},
      "AI_tag_2": {"max_chars": 14, "single_line": true},
      "AI_tag_3": {"max_chars": 14, "single_line": true},
      "AI_content": {
        "max_bullets_no_wrap": 6,
        "max_bullets_some_wrap": 4,
        "max_bullets_much_wrap": 3,
        "words_per_bullet": [4, 6]
      }
    },
    "bars": {
      "AI_title": {"max_chars": 25, "single_line": true},
      "AI_content": {
        "max_bullets_no_wrap": 5,
        "max_bullets_some_wrap": 4,
        "max_bullets_much_wrap": 3,
        "words_per_bullet": [5, 7]
      }
    },
    "hexagon": {
      "AI_title": {"max_chars": 25, "single_line": true},
      "AI_content": {
        "max_bullets_no_wrap": 5,
        "max_bullets_some_wrap": 4,
        "max_bullets_much_wrap": 3,
        "words_per_bullet": [5, 7]
      }
    }
  },
  "slide_rules": {
    "6": {"bars": 1, "hex": 0},
    "7": {"bars": 2, "hex": 0},
    "8": {"bars": 3, "hex": 0},
    "9": {"bars": 3, "hex": 1},
    "10": {"bars": 3, "hex": 2},
    "11": {"bars": 3, "hex": 3},
    "12": {"bars": 4, "hex": 3},
    "13": {"bars": 4, "hex": 4},
    "14": {"bars": 5, "hex": 4},
    "15": {"bars": 5, "hex": 5},
    "16": {"bars": 6, "hex": 5},
    "17": {"bars": 6, "hex": 6},
    "18": {"bars": 7, "hex": 6},
    "19": {"bars": 8, "hex": 6},
    "20": {"bars": 9, "hex": 6}
  },
  "loop_rules": {
    "beyond_20": "alternate bars-block (3-4) with hex-block (3-4)",
    "block_size_range": [3, 4]
  }
}

3. Refactor fill_template.py:
   - Add command-line arguments: python fill_template.py --template neon --topic "..." --count 13
   - Load spec.json at startup
   - All rule lookups now go through the spec (not hardcoded)
   - Engine code should not know anything about "neon" specifically
   - Split into clear functions:
     * load_spec(template_name)
     * build_slide_sequence(count, spec)
     * generate_content(topic, spec, sequence)
     * fill_placeholders(prs, content, spec)
     * save_output(prs, output_path)

4. Test that everything still works with the neon template:
   python fill_template.py --template neon --topic "Machine Learning" --count 13

Save output as output_v5.pptx and confirm it looks like previous outputs.
```

**Verify after done:** Open output in PowerPoint. Should look identical to previous outputs. Structure changed, output didn't.

---

### STEP 2 — GRADIO WEB UI + HUGGING FACE SPACES (3-4 days)

**Why:** A public link anyone can click to use your tool. Essential for funders, uni, Reddit posts.

**Before starting:** 
1. Create free account at huggingface.co
2. On Hugging Face, click "New Space" → SDK: Gradio → name it (e.g., `slidemorph`) → Public

**Paste this prompt into Claude Code:**

```
Build a Gradio web UI on top of the existing fill_template.py logic. Do NOT rewrite 
the engine — just wrap it.

BUILD THIS:

1. Create app.py in C:\slidemorph with a Gradio interface:
   - Title: "SlideMorph — AI Presentations that Actually Animate"
   - Brief description of the tool
   - Dropdown to select a template (only "Neon" for now, but read available templates 
     from templates/ folder dynamically)
   - Text box for the topic
   - Slider for slide count (6 to 30, default 13)
   - "Generate Presentation" button
   - File download component for the .pptx output
   - Show a template preview image (static image for now)
   - Loading spinner during generation
   - Error messages if generation fails

2. When user clicks Generate:
   - Call functions from fill_template.py
   - Save output to temp file
   - Return file for download
   - Show status to user

3. No login. No accounts. No payments. Just a public tool.

4. Create requirements.txt with:
   python-pptx, lxml, fonttools, pillow, google-generativeai, gradio, requests, 
   python-dotenv

5. Create README.md for Hugging Face Space with:
   ---
   title: SlideMorph
   emoji: 🎬
   colorFrom: purple
   colorTo: blue
   sdk: gradio
   sdk_version: 4.x
   app_file: app.py
   pinned: false
   ---
   
   Then a description of the tool.

6. Rate limit: 3 generations per IP per hour (in-memory). Show friendly message if 
   exceeded.

7. Gemini API key: load with os.getenv("GEMINI_API_KEY"). Works with both local .env 
   and Hugging Face Space secrets.

Test locally with: python app.py
Should open http://localhost:7860 with the interface.

Once local testing works, tell me step-by-step how to deploy to Hugging Face Spaces:
- How to upload files (Git or web interface)
- How to add GEMINI_API_KEY as a Space secret
- How to verify it's live
```

**After it works locally:**
Ask Claude Code: "Walk me through deploying this to Hugging Face Spaces step by step."

Follow the steps. You'll get a URL like `yourname-slidemorph.hf.space` that anyone can visit.

---

### STEP 3 — PUBLISH THE LIBRARY (2-3 days)

**Why:** This is your open-source credential. It gets you noticed by engineers at Plus AI, Gamma, SlideSpeak. It's your job/contract magnet. Independent of whether the SaaS makes money.

**Before starting:**
1. Create GitHub account (if you don't have one)
2. Create a new PUBLIC repo called `pptx-morph-safe` (leave it empty for now)
3. Create a PyPI account at pypi.org

**Paste this prompt into Claude Code:**

```
Create a standalone open-source Python library called pptx-morph-safe based on the 
reusable engine functions in fill_template.py. It should be publishable to PyPI and 
GitHub as a separate project — separate from the SaaS.

CREATE THIS FOLDER STRUCTURE:
C:\pptx-morph-safe\
├── pptx_morph_safe\
│   ├── __init__.py
│   ├── morph.py          (make_morph_safe function)
│   ├── autofit.py        (disable_autofit function)
│   ├── text_fill.py      (set_text_preserve_format function)
│   ├── slides.py         (delete_slides, duplicate_slide functions)
│   ├── media.py          (compress_media function)
│   └── validation.py     (fit_check function using fontTools)
├── tests\
│   ├── test_morph.py
│   ├── test_autofit.py
│   └── test_text_fill.py
├── examples\
│   ├── basic_fill.py
│   └── with_morph.py
├── README.md
├── LICENSE               (MIT license — full text)
├── pyproject.toml
├── .gitignore
└── setup.py

FUNCTION SPECS:

1. make_morph_safe(prs, prefix="!!"):
   - Takes a Presentation object
   - Renames shapes matching patterns (default: names starting with "AI_") to have 
     "!!" prefix
   - Assigns unique names per slide where needed
   - Returns modified presentation

2. disable_autofit(shape):
   - Takes a shape
   - Sets text frame to noAutofit via XML manipulation
   - Prevents morph from animating box size changes

3. set_text_preserve_format(shape, text_or_bullets):
   - Takes a shape and either string (single-line) or list of strings (bullets)
   - Replaces text while preserving all run-level formatting
   - Handles paragraph duplication with hanging indent preservation
   - The critical function everyone rediscovers because .text setter wipes formatting

4. delete_slides(prs, keep_indices):
   - Removes slides not in keep_indices
   - Handles relationship cleanup properly

5. duplicate_slide(prs, source_index):
   - Deep copies a slide including relationships
   - Assigns new unique rId
   - Returns new slide's index

6. compress_media(prs, max_px=1920, quality=85):
   - Iterates embedded images
   - Downscales to max_px on longest side
   - Recompresses as JPEG

7. fit_check(text, shape, font_path):
   - Uses fontTools + Pillow to measure rendered pixel width
   - Compares against shape's box width (EMU → pixels)
   - Returns (fits: bool, wrap_count: int)

WRITE README.md WITH:
- Title: pptx-morph-safe
- One-liner: "Safe manipulation of PowerPoint templates with Morph transitions"
- Problem: python-pptx doesn't support animations; naive text replacement breaks Morph
- Solution: this library handles XML-level details
- Installation: pip install pptx-morph-safe
- Quick 10-line example
- Full API reference
- Why !! naming matters (shape ID matching in Morph)
- Why autofit destroys Morph
- MIT license badge

TESTS:
Write at least one test per function using pytest. Include a small sample .pptx in 
tests/fixtures/.

WHEN DONE, TELL ME:
1. How to test locally (python -m pytest)
2. How to push to GitHub (I created empty repo at github.com/[username]/pptx-morph-safe)
3. How to publish to PyPI (as version 0.1.0)
```

**After it's built, ask Claude Code:**
1. "Walk me through pushing this to GitHub step by step."
2. "Walk me through publishing to PyPI step by step."

Once live: anyone can `pip install pptx-morph-safe`.

---

### STEP 4 — ADD MORE TEMPLATES (1 day each, only Type A)

**Type A templates:** Same wheel/bars/hexagon structure, different backgrounds. Easy — just new JSON + new .pptx.

**Type B templates:** Completely different designs. Skip until after launch — too much work.

**Paste this prompt into Claude Code for each new Type A template:**

```
Add a new template called "[name]" to templates/ folder. Same structure as neon 
(wheel + bars + hexagon), different visuals.

DO THIS:

1. Create templates/[name]/ folder
2. Copy templates/neon/spec.json to templates/[name]/spec.json
3. Update spec.json to reference [name]/template.pptx
4. Change "name" and "description" fields
5. Verify template has same placeholder names as neon
6. If placeholder box sizes differ, adjust character limits and bullet rules

After I upload the .pptx to templates/[name]/template.pptx, test with:
python fill_template.py --template [name] --topic "test topic" --count 13
```

---

## PART 7 — TIMELINE

4 weeks total on Claude Pro subscription:
- **Week 1:** Step 1 (JSON spec system)
- **Week 2:** Step 2 (Gradio + Hugging Face deploy)
- **Week 3:** Step 3 (library on PyPI + GitHub)
- **Week 4:** Buffer + Step 4 (add 1-2 Type A templates) + write blog post

---

## PART 8 — AFTER SUBSCRIPTION (no Claude Code needed)

These don't need paid Claude:
1. **Write the blog post** — "How to fill a PowerPoint template with AI without breaking its Morph animations." Document the `!!` naming, shape-ID matching, autofit problem, rotation wrap. No good public reference exists — this ranks and gets you found.
2. **Post to Show HN + r/Python + r/programming** — frame as file-format engineering, not product pitch.
3. **Post to r/PowerPoint, r/presentations, r/GradSchool, r/college, r/Professors** — post the morph GIF, not the landing page.
4. **List a template pack on Gumroad** — fastest path to first dollar.
5. **Fiverr/Upwork** — "animated morph PowerPoint" service, $25-50/deck.
6. **Apply to NIC Islamabad** — closest incubator, funds AI/deep-tech students. Grant + AWS credits. Watch their LinkedIn for cohort deadlines.
7. **Ask university's ORIC office for a nomination and intros.**
8. **Weekly build-in-public posts on X/LinkedIn** — 20 min/week, short screen recordings of morph.

---

## PART 9 — KNOWN PROBLEMS AND FIXES

**OneDrive issue:** Project must live at `C:\slidemorph` (NOT inside Documents/Desktop which OneDrive syncs). OneDrive dehydrates files and breaks builds. If you see `PackageNotFoundError` or `PermissionError`, this is why.

**Gemini quota:** Free tier hits daily limits fast. Currently using `gemini-2.0-flash`. If quota exhausted, switch to `gemini-1.5-flash` or wait for reset (midnight Pacific = 12 PM PKT next day).

**Gemini training on free tier data:** Free tier may use prompts to train Google's models. Fine for prototype. Before real paying customers, move to billed project or note it on the site.

**Enabling billing kills free tier:** In Google Cloud, enabling billing on a project silently removes free tier. Keep prototyping on a separate project with billing OFF.

**Morph and shape names:** Morph matches shapes by ID first, then by name. Names starting with `!!` are matched by name reliably. Without the prefix, matching is unreliable across slides.

**Autofit + Morph = broken animations:** Autofit resizes boxes at open time, which Morph then animates as a stretch. Always disable autofit on AI content boxes.

**Box vs text morphing:** If you want boxes to morph but text to fade separately, give text unique names per slide. If you want boxes to appear/disappear cleanly, give the whole shape unique names per slide (current setup does this).

---

## PART 10 — NEXT ACTION

**Right now: Step 1.**

1. Open Claude desktop app → Code tab
2. Point it to `C:\slidemorph`
3. Copy the Step 1 prompt from Part 6 above
4. Paste into Claude Code
5. Let it run
6. When done, test with:
   `python fill_template.py --template neon --topic "Machine Learning" --count 13`
7. Open the output in PowerPoint — should look same as before
8. Report back for Step 2

---

## PART 11 — IF USING A NEW CHAT

If this is a new chat and I've never seen this project before, here's what you need to know to help me effectively:

1. Read all 11 parts of this document first
2. I'm technical enough to run terminal commands and use Claude Code, but this is my first real project — explain unfamiliar concepts briefly
3. Give me exact prompts to paste into Claude Code, not general advice
4. My tone preference: brutally honest, problem-solver, concrete next actions, no hype
5. Answer format: short answers by default, longer only when the topic genuinely needs it
6. Never assume I know DevOps, cloud deployment, or PyPI publishing — walk me through those
7. Trust that the technical work described in Part 5 is actually done and working
8. Trust the decisions in Part 2 — don't second-guess the stack

**Where I am when I paste this in:** [FILL IN THE STEP YOU'RE ON — e.g., "Step 1 done, moving to Step 2" or "Halfway through Step 3"]

**My immediate question:** [WRITE YOUR QUESTION HERE]

---

*End of document. Save this file and paste it into any future chat to continue without losing context.*
