# SEO Meta Translation & Constraint Tool (Core PHP)

---

## 📌 Overview

This tool processes a CSV file containing English meta titles and/or meta descriptions, translates them into a target language, and ensures they strictly comply with SEO character limits.

If the translated content fails validation, it is automatically rewritten using AI while **preserving the exact meaning of the original English text**.

---

## 🎯 Objectives

* Translate meta content using DeepL
* Enforce strict SEO constraints:

  * Meta Title: **40–46 characters**
  * Meta Description: **150–155 characters**
* Automatically regenerate invalid content
* Preserve **exact meaning and intent**
* Process bulk data via CSV
* Use a cost-efficient multi-AI fallback system

---

## ⚙️ Tech Stack

* Core PHP (no frameworks)
* cURL (API communication)
* CSV handling (fgetcsv / fputcsv)

### APIs Used

* DeepL → Translation
* Google AI Studio → Primary rewrite
* Groq → Fallback rewrite

---

## 📂 Project Structure

```
/meta-tool
│── index.php              # Upload UI
│── upload.php             # File upload handler
│── process.php            # Main processing pipeline
│── output.csv             # Generated output
│
└── /lib
    │── deepl.php          # DeepL integration
    │── gemini.php         # Gemini API logic
    │── groq.php           # Groq API logic
    │── validator.php      # Length validation
    │── helper.php         # Common utilities
    │── logger.php         # Optional logging
```

---

## 📥 Input CSV Format

The CSV may contain:

* `meta_title` (optional)
* `meta_description` (optional)

Example:

```
url,meta_title,meta_description
/page-1,Best flight deals today,Find cheap flights and save big on your next trip.
/page-2,,Get compensation for delayed flights quickly.
```

---

## 📤 Output CSV Format

The output file appends:

```
new_title,new_description
```

---

## 🔄 Processing Flow

1. Read CSV row
2. If `meta_title` exists:

   * Translate using DeepL
   * Validate length
   * If invalid → rewrite via Gemini
   * If still invalid → rewrite via Groq
   * If still invalid → hard trim
3. Repeat same steps for `meta_description`
4. Write results to output CSV

---

## 📏 Validation Rules

### Meta Title

* Minimum: 40 characters
* Maximum: 46 characters

### Meta Description

* Minimum: 150 characters
* Maximum: 155 characters

Use multibyte-safe functions:

```
mb_strlen()
mb_substr()
```

---

## 🔴 CRITICAL RULE: Meaning Preservation (MANDATORY)

All regenerated content MUST preserve the **exact meaning of the original English text**.

### ✅ Required

* Same meaning and intent
* Same SEO purpose (informational / transactional)
* Preserve key terms and entities
* Maintain original context

### ❌ Forbidden

* Adding new information
* Removing important concepts
* Changing intent
* Introducing marketing fluff
* Replacing core keywords

---

## 🤖 AI Rewrite Prompt Standard

Every rewrite request MUST follow:

```
Rewrite in {target_language}.

STRICT RULES:
- Keep EXACT same meaning as original English text
- Do NOT add or remove information
- Do NOT change intent
- Preserve key terms
- MUST be {character_limit}
- SEO optimized
- Natural and fluent
- Return ONLY final sentence

Text:
{original_text}
```

---

## 🔁 Retry Strategy

For each field:

* 1× DeepL translation
* 2× Gemini rewrite attempts
* 2× Groq rewrite attempts

If all fail:
→ Apply hard trim as last fallback

---

## 🧠 Meaning Validation (Recommended)

### Option 1: Keyword Check (Lightweight)

* Extract important keywords from original
* Ensure presence in rewritten output

---

### Option 2: AI Semantic Validation (Advanced)

```
Compare:

Original (English):
{original}

Rewritten ({language}):
{rewritten}

Does the rewritten text preserve EXACT meaning?

Answer ONLY: YES or NO
```

If NO → retry

---

## ⚠️ Priority Order

1. Meaning preservation (highest priority)
2. Character limits
3. SEO optimization

👉 Meaning must NEVER be sacrificed.

---

## 🧩 Core Logic Rules

* Always rewrite using **original English text**
* Never rewrite from translated text
* Skip empty fields
* Retry if limits fail
* Retry if meaning changes

---

## 🧪 Edge Cases

* Very short input → requires expansion
* Very long input → requires compression
* German/French → longer words (frequent overflow)
* AI ignoring limits → retry required
* Mixed-language input

---

## 🔐 Environment Variables

Store API keys securely:

```
DEEPL_API_KEY=
GEMINI_API_KEY=
GROQ_API_KEY=
```

---

## 🚀 Optimization Tips

* Skip rewrite if length is near valid range
* Reduce API calls by validating early
* Batch process large CSV files
* Add queue system for scaling

---

## 🧾 Logging (Recommended)

Log each step:

```
original | deepl | gemini | groq | final | attempts | status
```

---

## 🌍 Multi-language Support

Make target language dynamic:

```
FR → French
DE → German
ES → Spanish
IT → Italian
```

---

## 🧩 Output Behavior Rules

| Scenario          | Action          |
| ----------------- | --------------- |
| Valid translation | Keep            |
| Too short         | Expand via AI   |
| Too long          | Compress via AI |
| Meaning changed   | Retry           |
| Still invalid     | Trim            |

---

## 🚨 Failure Handling

If all attempts fail:

* Return closest valid output
* Log failure for manual review

---

## 🔮 Future Enhancements

* Web UI (drag & drop upload)
* Multi-language batch mode
* WordPress / WooCommerce integration
* Google Sheets integration
* Background queue processing
* Semantic similarity scoring

---

## ✅ Summary

This tool is a:

* Translation engine (DeepL)
* AI rewriting system (Gemini + Groq)
* SEO validation engine
* Meaning-preserving processor
* Bulk CSV automation tool

It ensures:

✔ Accurate translation
✔ Strict SEO compliance
✔ Meaning integrity
✔ Scalable processing

---

## 🛠️ Instructions for Claude

Generate:

1. Full Core PHP project
2. Modular file structure
3. Clean reusable functions
4. Proper error handling
5. API integrations via cURL
6. CSV upload → process → download flow
7. Optional logging system

### Do NOT:

* Use frameworks
* Add unnecessary dependencies
* Overcomplicate logic

Keep code clean, minimal, and production-ready.

---