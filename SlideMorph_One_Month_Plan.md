# SlideMorph — One-Month Technical Execution Plan

*Use your Claude Pro subscription for the technical work in this document. Do the non-technical work (marketing, funding, Gumroad) after the subscription ends — you don't need Claude Code for that.*

---

## Current status

- ✅ Text filling with formatting preserved
- ✅ Morph transitions working
- ✅ Fade animations on bullets
- ✅ Slide selection logic (6-30 slides)
- ✅ Looping with alternating blocks
- ✅ Gemini AI generating real content
- ✅ Adaptive bullet rules

**You are here.** One working template, command-line only, no web UI, no library published.

---

## The four technical steps (in order)

### Step 1 — JSON spec system (2-3 days)

**Why:** Right now, everything is hardcoded for your neon template. If you want to add a new template, you'd rewrite the whole script. This step extracts all the rules into a JSON file, so adding a new template later means just adding a new JSON + a new .pptx. It also makes the library reusable — other people can use your engine for their own templates.

**What Claude Code will do:**
- Create a folder called `templates/` inside `C:\slidemorph`
- Create `templates/neon/` containing:
  - `template.pptx` (a copy of your current template)
  - `spec.json` (all the rules — slide counts, placeholder names, character limits, bullet rules, looping logic)
  - `preview.mp4` (a pre-recorded preview video you'll add later)
- Rewrite `fill_template.py` to read the spec file instead of having rules baked in
- Add a `--template` command-line option so you can pick which template to use

**Prompt to paste into Claude Code:**

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
   - Add a command-line argument: python fill_template.py --template neon --topic "..." --count 13
   - Load the spec.json file at startup
   - All rule lookups now go through the spec (not hardcoded)
   - The engine code should not know anything about "neon" specifically — it should work with any spec that follows this structure
   - Split the code into clear sections/functions:
     * load_spec(template_name)
     * build_slide_sequence(count, spec)
     * generate_content(topic, spec, sequence)
     * fill_placeholders(prs, content, spec)
     * save_output(prs, output_path)

4. Test that everything still works exactly as before with the neon template.

Run: python fill_template.py --template neon --topic "Machine Learning" --count 13
Save output as output_v5.pptx and verify it looks identical to previous outputs.

Show me the new file structure and the refactored code.
```

**When done:** Open the output in PowerPoint. Should look identical to before. The important thing is the code structure changed, not the output.

---

### Step 2 — Gradio web UI on Hugging Face Spaces (3-4 days)

**Why:** A public link you can share with anyone. Funders click a link, use your tool in their browser, download a real .pptx. This is what makes the product feel real.

**What Claude Code will do:**
- Create `app.py` — a Gradio interface
- Wrap your existing engine (no rewrite, just wrap it)
- Add file uploads for the template gallery (later)
- Deploy to Hugging Face Spaces (free, always live)

**Before starting:** Create a free account at huggingface.co. Then create a new Space (Gradio type). You'll get a URL like `yourname-slidemorph.hf.space`.

**Prompt to paste into Claude Code:**

```
Build a Gradio web UI on top of the existing fill_template.py logic. Do NOT rewrite 
the engine — just wrap it.

BUILD THIS:

1. Create app.py in C:\slidemorph with a Gradio interface that has:
   - A title: "SlideMorph — AI Presentations that Actually Animate"
   - A subtitle/description explaining the tool briefly
   - A dropdown to select a template (only "Neon" for now, but code should read available 
     templates from the templates/ folder dynamically)
   - A text box for the topic
   - A slider for slide count (6 to 30, default 13)
   - A "Generate Presentation" button
   - A file download component for the .pptx output
   - A section showing the template preview (for now just show a static image; later 
     we'll add a video)
   - Loading spinner while generating
   - Error messages if generation fails

2. When the user clicks Generate:
   - Call the engine functions from fill_template.py
   - Save the output to a temp file
   - Return the file for download
   - Show generation status to the user

3. Do not require login. No accounts. No payments. Just a public tool.

4. Add a requirements.txt file with all dependencies:
   python-pptx, lxml, fonttools, pillow, google-generativeai, gradio, requests, 
   python-dotenv

5. Add a README.md for the Hugging Face Space with:
   - Title
   - Emoji: 🎬
   - SDK: gradio
   - Python version: 3.11
   - App file: app.py

6. Rate limiting: since it's on free tier, add a simple in-memory rate limit — 
   3 generations per IP per hour. If exceeded, show a friendly message.

7. Make sure the Gemini API key is loaded from Hugging Face Space secrets (not from 
   .env when deployed). Locally, keep using .env. Use os.getenv("GEMINI_API_KEY") 
   which will work for both.

Test locally first with: python app.py
It should open a browser at http://localhost:7860 with the interface.

Once local testing works, tell me how to deploy to Hugging Face Spaces.
```

**After it builds:**
1. Test locally (`python app.py`) — make sure the interface works and generates files
2. Then ask Claude Code: "Now walk me through deploying this to Hugging Face Spaces step by step, including how to upload files and add the Gemini API key as a Space secret."

Follow those steps. You'll end up with a public URL you can share with anyone.

---

### Step 3 — Publish the library `pptx-morph-safe` (2-3 days)

**Why:** This is your credential. When you apply to NIC or when Plus AI's engineer searches for a python-pptx animation workaround, this is what they find. It's separate from the SaaS.

**What Claude Code will do:**
- Extract the reusable engine functions from `fill_template.py` into a proper Python package
- Add tests, README, license, examples
- Set up for `pip install pptx-morph-safe`

**Before starting:** Create a public GitHub repo called `pptx-morph-safe`. Get a GitHub account if you don't have one (free).

**Prompt to paste into Claude Code:**

```
Create a standalone open-source Python library called pptx-morph-safe based on the 
reusable engine functions in fill_template.py. It should be publishable to PyPI and 
GitHub as a separate project.

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
│   ├── basic_fill.py     (example: fill a template)
│   └── with_morph.py     (example: preserve morph while filling)
├── README.md
├── LICENSE               (MIT license)
├── pyproject.toml
├── .gitignore
└── setup.py

WHAT EACH FUNCTION SHOULD DO:

1. make_morph_safe(prs, prefix="!!"):
   - Takes a Presentation object
   - Renames all shapes matching certain patterns (default: names starting with "AI_") 
     to have the "!!" prefix so morph matches them by name
   - Also assigns unique names per slide where needed
   - Returns the modified presentation

2. disable_autofit(shape):
   - Takes a shape
   - Sets its text frame to noAutofit via XML manipulation
   - Prevents morph from animating box size changes

3. set_text_preserve_format(shape, text_or_bullets):
   - Takes a shape and either a string (single-line) or a list of strings (bullets)
   - Replaces text while preserving all run-level formatting (font, color, size, bold)
   - For bullets: handles paragraph duplication with hanging indent preservation
   - This is the function everyone rediscovers because python-pptx's .text setter 
     wipes formatting

4. delete_slides(prs, keep_indices):
   - Removes slides not in keep_indices
   - Handles relationship cleanup properly

5. duplicate_slide(prs, source_index):
   - Deep copies a slide including relationships
   - Assigns a new unique rId
   - Returns the new slide's index

6. compress_media(prs, max_px=1920, quality=85):
   - Iterates embedded images
   - Downscales to max_px on longest side
   - Recompresses as JPEG at given quality

7. fit_check(text, shape, font_path):
   - Uses fontTools + Pillow to measure rendered pixel width of text
   - Compares against shape's box width in EMU (converted to pixels)
   - Returns tuple: (fits: bool, wrap_count: int)

WRITE THE README.md WITH:
- Title: pptx-morph-safe
- One-line description: "Safe manipulation of PowerPoint templates with Morph transitions"
- The problem: python-pptx doesn't support animations, and naive text replacement 
  breaks Morph transitions
- The solution: this library handles the XML-level details
- Installation: pip install pptx-morph-safe
- Quick example (10 lines) showing filling a template
- Full API reference
- Why the !! naming convention matters (explain shape ID matching in Morph)
- Why autofit destroys Morph
- Link to your blog post (I'll write this later)
- MIT license

TESTS:
Write at least one test per function using pytest. Use a small sample template file 
in tests/fixtures/.

Once everything is created, tell me:
1. How to test the library locally (python -m pytest)
2. How to publish it to PyPI (I want to release version 0.1.0)
3. How to push it to a GitHub repo (I'll create the repo manually first)
```

**After it builds:**
1. Run the tests locally
2. Ask Claude Code to walk you through pushing to GitHub
3. Ask Claude Code to walk you through publishing to PyPI

Once it's on PyPI, anyone in the world can `pip install pptx-morph-safe`. That's your credential live.

---

### Step 4 — Add Type A templates (1 day each, only if time allows)

**Why:** More templates = product feels bigger. But only Type A (same design logic, different skin). Type B (completely different designs) needs its own JSON spec and takes 3+ days — not worth it before launch.

**What Claude Code will do:**
- Copy your neon spec.json for each new skin
- Adjust colors/background references in the spec
- Test each new template

**Prompt (repeat per new template):**

```
Add a new template called "[name]" to the templates/ folder. It uses the same design 
logic as neon (wheel, bars, hexagon) but with different visuals.

DO THIS:

1. Create templates/[name]/ folder
2. Copy templates/neon/spec.json to templates/[name]/spec.json
3. Update the spec.json to reference [name]/template.pptx
4. Change the "name" and "description" fields in the spec
5. Verify the template has the same placeholder names as neon (AI_title, AI_content, 
   AI_tag_1/2/3)
6. If placeholder names or box sizes differ, adjust character limits and bullet rules 
   in the spec accordingly

After I upload the .pptx file to templates/[name]/template.pptx, test with:
python fill_template.py --template [name] --topic "test topic" --count 13

Save output and confirm it works.
```

---

## Timeline (4 weeks)

- **Week 1:** Step 1 (JSON spec system)
- **Week 2:** Step 2 (Gradio + Hugging Face deploy)
- **Week 3:** Step 3 (library on PyPI + GitHub)
- **Week 4:** Buffer + Step 4 (Type A templates) + write blog post

If you finish early, use extra time to polish the Gradio UI or add more Type A templates.

---

## After the subscription ends (no Claude Code needed)

- Write the blog post (you can use free ChatGPT or write it yourself — you know the technical details)
- Post the blog + library to Show HN and r/Python
- Post the live demo GIF to r/PowerPoint, r/presentations, r/GradSchool
- List a template pack on Gumroad
- Apply to NIC Islamabad
- Ask your university for a nomination
- Build audience on X/LinkedIn with weekly build-in-public posts

---

## Which step to start NOW

Step 1. Copy the prompt above and paste it into Claude Code.

When it's done, come back and tell me. I'll walk you through Step 2.
