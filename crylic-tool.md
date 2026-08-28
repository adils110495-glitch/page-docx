Create a new standalone tool in a **new browser tab** for cleaning corrupted English text containing Cyrillic/lookalike Unicode characters.

## Tool Name

**Cyrillic Word Cleaner**

Suggested route:

`/cyrillic-cleaner/`

---

## 1. Use `crilic-wordss.csv` as the Main Reference

There is a CSV file named:

**`crilic-wordss.csv`**

This CSV contains the actual Cyrillic/lookalike words that need to be replaced with their correct English/Latin versions.

**The CSV must be used dynamically as the primary replacement dictionary.**

Do NOT manually hardcode only the examples from the prompt.

Read/import the mappings from `crilic-wordss.csv` and use all available entries in that file.

If the CSV contains columns such as:

* Cyrillic/incorrect word
* Correct English word
* Replacement
* Original
* Correct Latin form

detect the appropriate source and target columns automatically or map them according to the actual CSV structure.

---

## 2. Input Textarea

Create a large textarea where the user can paste:

* A single word
* Sentence
* Paragraph
* Full article
* Large amount of text

Placeholder:

**`Paste your text here...`**

---

## 3. Output Textarea

Create a second large textarea showing the cleaned result.

Placeholder:

**`Cleaned English text will appear here...`**

The output must preserve the original content.

Only corrupted Cyrillic/lookalike characters or words should be corrected.

---

## 4. CSV-Based Word Replacement

First, check the pasted text against **all mappings available in `crilic-wordss.csv`**.

For every matching incorrect/Cyrillic word:

**Incorrect word → Correct English word**

For example, if the CSV contains mappings such as:

`fоr → for`

`Prоtectiоns → Protections`

`саn → can`

`аnԁ → and`

then the tool must automatically apply those mappings.

Do not limit the system to these examples. **Every valid mapping inside `crilic-wordss.csv` must be supported.**

---

## 5. IMPORTANT — Handle Words NOT Present in CSV

The tool must NOT stop at CSV matching.

After applying the CSV mappings, perform a second **Unicode character-level scan**.

If an English word contains a Cyrillic/lookalike character that is not explicitly present as a complete word in the CSV, detect the suspicious character and convert it to the correct Latin character when the mapping is unambiguous.

For example:

`informаtion`

where `а` is Cyrillic rather than Latin:

`informаtion → information`

Similarly:

`delаyed → delayed`

`cоmpensation → compensation`

`аirline → airline`

even if that exact corrupted word does not exist in the CSV.

### Core requirement:

**CSV word mapping + generic Unicode character mapping**

Both systems must work together.

---

## 6. Character-Level Unicode Detection

Detect Cyrillic characters commonly used as Latin lookalikes.

At minimum support mappings such as:

* `а` → `a`
* `А` → `A`
* `с` → `c`
* `С` → `C`
* `е` → `e`
* `Е` → `E`
* `о` → `o`
* `О` → `O`
* `р` → `p`
* `Р` → `P`
* `х` → `x`
* `Х` → `X`
* `у` → `y`
* `У` → `Y`

Also support special characters found in `crilic-wordss.csv`, such as:

* `ԁ` → `d`

Do not assume the CSV is the only source of possible corrupted words.

---

## 7. Do NOT Rewrite the Content

This is extremely important.

The tool is **NOT an AI paraphraser, grammar checker, or content rewriter**.

If the user enters:

`Yоur flight was ԁelаyeԁ due to cancellatiоn.`

the result should be:

`Your flight was delayed due to cancellation.`

Nothing else should change.

Do NOT:

* Rewrite sentences
* Change wording
* Improve grammar
* Change sentence structure
* Summarize
* Paraphrase
* Change punctuation unnecessarily
* Change formatting

Only fix the corrupted characters/words.

---

## 8. Preserve Formatting

Preserve:

* Paragraphs
* Line breaks
* Spaces
* Punctuation
* Numbers
* Capitalization
* URLs
* Email addresses
* HTML content, if pasted
* Markdown formatting

Only modify the affected Unicode characters/words.

---

## 9. Case Preservation

Maintain the correct capitalization.

Examples:

`Yоur` → `Your`

`Fоr` → `For`

`Cоmpensatiоn` → `Compensation`

`Аre` → `Are`

Do not lowercase or uppercase the entire text.

---

## 10. Processing Order

Use this exact processing strategy:

### Step 1

Load `crilic-wordss.csv`.

### Step 2

Build a replacement dictionary from every valid CSV mapping.

### Step 3

Apply exact word-level replacements from the CSV.

### Step 4

Scan the resulting text character-by-character for remaining suspicious Cyrillic/lookalike Unicode characters.

### Step 5

Replace unambiguous lookalike characters with their Latin equivalents.

### Step 6

Scan the final output again for remaining suspicious Unicode characters.

### Step 7

Show a warning if suspicious characters still remain and cannot safely be converted.

---

## 11. Statistics

Show useful statistics after cleaning:

**Characters Replaced:** `25`

**Words Corrected:** `12`

**Total Replacements:** `37`

**Suspicious Characters Remaining:** `0`

If nothing was found:

**No Cyrillic/lookalike characters detected.**

---

## 12. Change Log

Add a **Show Changes** option.

When enabled, display something like:

| Original | Corrected | Type        | Count |
| -------- | --------- | ----------- | ----: |
| `fоr`    | `for`     | CSV mapping |    12 |
| `саn`    | `can`     | CSV mapping |     8 |
| `а`      | `a`       | Unicode     |     5 |
| `ԁ`      | `d`       | Unicode     |     2 |

This helps the user verify what was changed.

---

## 13. Buttons

Add:

### Clean Text

Processes the input.

### Copy Result

Copies only the cleaned output.

Show:

**Copied!**

after successful copying.

### Clear

Clears:

* Input
* Output
* Statistics
* Change log
* Warnings

---

## 14. Auto Clean

Add an optional toggle:

**Auto Clean**

When enabled:

* Automatically clean pasted text.
* Update the output without requiring the user to click Clean Text.

When disabled:

* Process only when the user clicks **Clean Text**.

For very large text, use a performant/debounced implementation.

---

## 15. Remaining Suspicious Characters

After processing, check whether any Cyrillic characters remain.

If none remain:

**✓ Text is clean**

If some remain:

**⚠ X suspicious Unicode characters remain**

Show the actual characters/codepoints where useful so the user can inspect them.

Do not blindly convert characters that are not confidently identifiable as Latin lookalikes.

---

## 16. CSV Must Be Maintainable

Do not permanently hardcode the CSV mappings into frontend JavaScript if avoidable.

The system should use `crilic-wordss.csv` as the source dictionary so that if the CSV is updated with additional mappings, the cleaner can use those new mappings without requiring a complete rewrite of the tool.

If the project architecture requires the CSV to be imported into a database or JSON dictionary for performance, implement that cleanly and document the relationship:

**`crilic-wordss.csv` → replacement dictionary → cleaner engine**

---

## 17. Performance

The tool should efficiently handle:

* Small text
* Multiple paragraphs
* Full articles
* Large pasted content

Avoid unnecessary API calls.

The actual character replacement should happen locally where possible.

Do not send the entire article to an AI API simply to perform character replacement.

---

## 18. Responsive Design

The complete tool must be mobile responsive.

### Desktop

Show:

**Input Textarea | Output Textarea**

side-by-side.

### Mobile

Stack them:

**Input**

↓

**Output**

↓

**Buttons**

No horizontal overflow.

All controls must remain usable on small screens.

---

## 19. New Browser Tab

Add this tool to the existing system and make it open in a **new browser tab** from the application navigation/menu.

Do not break existing navigation or modules.

---

## 20. Critical Requirement

The cleaner must have **two layers of detection**:

### Layer 1 — `crilic-wordss.csv`

Use every incorrect → correct mapping available in the CSV.

### Layer 2 — Generic Unicode Detection

Detect and fix Cyrillic/lookalike characters even when the **complete corrupted word is NOT present in the CSV**.

Therefore:

**CSV is the authoritative word-level dictionary, but it is NOT the limit of the cleaner.**

The final goal is:

**Any English text containing Cyrillic/lookalike characters → clean standard English/Latin text**

while preserving the original wording, formatting, capitalization, punctuation, and meaning exactly.
