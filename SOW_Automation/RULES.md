# SOW Automation Tool - Processing Rules

## Overview

The SOW Automation Tool processes an Excel file (clause definitions) and a Word document (SOW template with highlighted regions) to generate a final SOW document based on user answers.

---

## Step 1: Upload

- User uploads an **Excel file** and a **Word file**.
- Excel columns: Row Number (B), Clause Number (C), SOW Text (D), SOW Summary (E), EP/EM Response (F), Tips (G), Variables (I).
- Word file contains **highlighted text** (yellow highlight or colored shading) marking editable regions.

## Step 2: Review (Create Page)

The system matches highlighted Word regions with Excel rows by comparing SOW Text (column D) against highlighted text. Matched rows are displayed in a table for user input.

### User Answer Conditions

Each matched clause is presented with an answer input. The type of input depends on the clause:

| Condition | Input Type | Behavior |
|-----------|-----------|----------|
| **Hierarchical tips** (`**1 **1a **1b`) AND has placeholders | Root tip control + cascading child inputs | See "Hierarchical Tips" section below |
| Has placeholders AND has options | Options dropdown + placeholder inputs | Selecting an option shows placeholder fields |
| Has placeholders AND Tips contains "yes"/"no" | Yes/No/N/A dropdown + placeholder inputs | Selecting "Yes" or a valid option shows placeholder fields |
| Has placeholders (no options, not yes/no) | Placeholder inputs shown directly | UserAnswer is auto-set to "Yes" |
| Has `*{}*` AND Tips contains "yes"/"no" (no `[...]` placeholders) | Yes/No/N/A dropdown + append textbox | Selecting "Yes" shows append field |
| Has `*{}*` AND has options (no `[...]` placeholders) | Options dropdown + append textbox | Selecting an option shows append field |
| Has `*{}*` only (no `[...]` placeholders, not yes/no, no options) | Append textbox shown directly | UserAnswer is auto-set to "Yes" |
| Tips contains "yes"/"no" (no placeholders) | Yes/No/N/A dropdown | Simple selection |
| Has options (no placeholders) | Options dropdown | Select from parsed options |
| Default | Free-text textarea | User types custom answer |

### Hierarchical Tips

When the Tips column contains markers like `**1`, `**1a`, `**1b`, the system parses them as a hierarchical tip structure.

**Format:** `**<number><optional letter> <content>`
- `**1 yes/no` = Root tip (level 1), determines the main answer type (yes/no in this case)
- `**1a Option 1 - Fixed, Option 2 - Flex` = Child of `**1`, maps to 1st placeholder, shows options dropdown
- `**1b` = Another child of `**1`, maps to 2nd placeholder, shows free-text input

**Rules:**
1. The root tip (`**1`) determines the top-level input: Yes/No dropdown, Options dropdown, or free-text.
2. Child tips (`**1a`, `**1b`, etc.) are shown below the root input when the user selects "Yes" (or a valid option).
3. Each child tip is mapped to a placeholder (excluding `[Optional...]` and `[NOTE TO DRAFT...]`) in order.
4. Each child tip's content determines its input type:
   - Contains "Option N" patterns: Options dropdown
   - Contains "yes"/"no": Yes/No/N/A dropdown
   - Otherwise: Free-text textarea with "Replace <field>" hint
5. When the root answer is "No" or "N/A", all child inputs are hidden and disabled.
6. The tip content is displayed as a label above each child input for context.

**Example:**
- Tips: `**1 yes/no **1a Option 1 - Fixed, Option 2 - Flex **1b`
- SOW Text contains: `[pricing model] and [delivery timeline]`
- UI: Root shows Yes/No dropdown. On "Yes": child `**1a` shows options dropdown for `[pricing model]`, child `**1b` shows free-text for `[delivery timeline]`.

### Options Parsing

Options are parsed from the SOW Summary column. Patterns recognized:
- `Option 1 - text... Option 2 - text...`
- `Option 1: Label: description... Option 2: Label: description...`
- `Alternative 1: text... Alternative 2: text...`

Short labels are extracted from `Label: rest of text` patterns when available.

### Placeholder Detection

Placeholders are bracket-delimited text in the SOW Text column: `[describe something]`.

**Rules:**
1. `[text]` - Simple placeholder. User provides replacement text.
2. `*[text]*` - **Escaped placeholder**. Asterisks before `[` AND after `]` mean this is NOT a placeholder and is ignored.
3. `[NOTE TO DRAFT ...]` - Always excluded from placeholders.
4. `[Optional...]` - **Auto-removed**. Any bracket text starting with "Optional" (e.g., `[Optional]`, `[Optional:]`, `[Optional:dcvsdggg]`) is not shown as a placeholder in the UI and is deleted from the final document along with any leading space.
5. `[outer text [inner1] and [inner2]]` - **Nested placeholder**. Contains inner bracket groups.

### Nested Placeholder Handling

When a nested placeholder like `[this is [enumm] and [enum]]` is detected:

- The UI presents a **Full Replacement** / **Custom Replacement** toggle.
- **Full Replacement**: User replaces the entire outer text `[this is [enumm] and [enum]]` with a single value.
- **Custom Replacement**: User replaces each inner placeholder (`[enumm]`, `[enum]`) individually, keeping the surrounding structure. **All custom replacement fields are required** and cannot be left empty. In the generated document, the outer brackets are stripped (e.g., `[power of value1 and value2]` becomes `power of value1 and value2`).

### Append Placeholder `*{}*`

The `*{}*` pattern in SOW text is an **append placeholder** that allows the user to optionally add or append new text at that position.

**Rules:**
- `*{}*` is highlighted in the SOW Text column with a distinct style.
- When the user answers "Yes" (or equivalent), a textbox appears with the hint: "Add or append text (leave blank to remove *{}*)".
- If the user enters text, `*{}*` is replaced with that text in the final document (e.g., `"darn you *{}*"` + `"rhenerya"` → `"darn you rhenerya"`).
- If the user leaves the textbox blank, `*{}*` and any leading space are removed (e.g., `"darn you *{}*"` → `"darn you"`).
- `*{}*` can coexist with `[...]` placeholders in the same row. The append textbox appears alongside the placeholder inputs.
- A final cleanup pass removes any remaining `*{}*` patterns from the entire document.

### Placeholder Hint Text

Each placeholder text box displays a hint: `Replace <field name>` where `<field name>` is the placeholder text without brackets. For example, placeholder `[xxxx]` shows "Replace xxxx" and `[cc]` shows "Replace cc".

### Placeholder Highlight in SOW Text Column

- Placeholders in the SOW Text column are wrapped in highlight-able spans.
- When a user clicks/focuses on a replacement text box, the corresponding placeholder text in the SOW Text column is highlighted in **yellow**.
- For nested placeholders: focusing the full replacement box highlights the entire outer text; focusing an inner replacement box highlights just that inner placeholder.

### Variables

Variables allow a user's answer for one clause to be reused across multiple SOW text entries.

**How it works:**

1. **Definition**: A row in the Excel file has variable name(s) in column I. The user's answer(s) for that row define the variable values.
2. **Reference**: Any SOW text (in any row) can reference a variable using the pattern `**variableName**` (e.g., `**xxx**`).
3. **UI behavior**: When the user types or selects an answer, all `**variableName**` references across all SOW text cells update **live** in the UI to show the resolved value.
4. **Document generation**: After all answer processing, a final pass replaces every `**variableName**` occurrence in the entire Word document with the corresponding value. The replacement is case-insensitive.
5. **No separate UI section**: Variables do not have their own section in the UI. The user simply fills in the answer for the row that defines the variable; the system handles propagation.

**Variable Column Format (Column I):**

Variables in column I map **positionally** to the placeholders found in the SOW Text of the same row.

| Variable Column | SOW Text Placeholders | Mapping |
|---|---|---|
| `PLACE,STATE` | `[xxxx]` and `[cccc]` | PLACE -> `[xxxx]`, STATE -> `[cccc]` |
| `TEXT,[TXT1,TXT2]` | `[there has been [yy] in [cc]]` | TEXT -> whole `[there has been [yy] in [cc]]`, TXT1 -> `[yy]`, TXT2 -> `[cc]` |
| `CITY` | `[enter city]` | CITY -> `[enter city]` |

**Simple variables** (comma-separated names): Each name maps to the corresponding placeholder in order. When the user fills the placeholder replacement text, that value is stored as the variable's value.

**Nested variables**: Use bracket syntax `[VAR1,VAR2]` to map inner placeholders of a nested placeholder.
- The name before the brackets maps to the whole (full replacement) placeholder.
- The names inside brackets map to the inner (custom replacement) placeholders in order.

**Example:**
- Row 3: SOW Text = `"there can be [xxxx] in [cccc]"`, Variables = `PLACE,STATE`
- Row 7: SOW Text = `"The location is **PLACE** in **STATE**"`
- User fills `[xxxx]` with "mountains" and `[cccc]` with "Colorado"
- In the UI, row 7 live-updates to "The location is mountains in Colorado"
- In the final document, all `**PLACE**` become "mountains" and `**STATE**` become "Colorado"

---

## Step 3: Generate (Document Generation)

The final Word document is generated by processing each highlighted region in the original Word template:

### Answer Processing Rules

| User Answer | Action |
|-------------|--------|
| **"Yes"** with placeholder values filled | Replace placeholder keys with user values in the highlighted runs, then remove highlight formatting |
| **"Yes"** without placeholder values | Remove highlight formatting only (keep original text) |
| **"No"** | Remove the entire paragraph. If the clause is a Schedule (e.g., "Sch A"), remove the entire schedule section (Heading1 to next Heading1) |
| **"N/A"** | Replace the highlighted text with "N/A" |
| **Custom text** (non-empty) | Replace the highlighted text with the user's custom text |
| **Empty** (no answer) | Remove highlight formatting only (keep original text unchanged) |

### Formatting Preservation

- When replacing text, the original run formatting (font, bold, italic, etc.) is preserved from the first run in the region.
- Highlight/shading formatting is always removed from replaced text.
- Bullets, numbering, and paragraph-level formatting are preserved.
- Checkboxes and other non-text elements within paragraphs are preserved.

### Final Cleanup

After all user answers are applied:
1. **Variable replacement**: All `**variableName**` references in the entire document are replaced with the user's answer for the corresponding variable-defining row. Case-insensitive matching.
2. **Append placeholder cleanup**: Any remaining `*{}*` patterns (not already handled during per-row processing) are removed along with any leading space.
3. **Escape marker removal**: All `*[text]*` patterns are stripped to `[text]` (asterisks removed, brackets preserved).
3. **All remaining highlight formatting** is removed from the entire document (any highlighted text not matched to a clause).
4. **`[NOTE TO DRAFT ...]`** and **`[Optional...]`** text is removed from all runs. This is done per-run to preserve surrounding formatting, bullets, and checkboxes.

### Section Marker Deletion (`*****<name>*****`)

Section markers allow entire multi-page sections of the Word document to be conditionally included or removed. Unlike schedule deletion (which uses Heading1 styles), section markers use explicit delimiter text in the Word document.

**Excel format:**
- Column D (SOW Text): `*****<section name>*****` (exactly 5 asterisks on each side)
- Column G (Tips): Describes the choice (e.g., "Please select yes/no")
- Column J (Parent Clauses): Optional, comma-separated parent clause numbers

**Word document format:**
- The Word document contains **two** marker instances with the text `*****<section name>*****`
- The markers define the **start** and **end** of a section
- Everything between the two markers (headings, paragraphs, tables, page breaks) belongs to that section
- The text between the markers may be partially highlighted, fully highlighted, or not highlighted at all
- **Multiple markers can share the same paragraph** (e.g., `*****A_Schedule***** *****B_Schedule*****` on one line)

**Rules:**
1. Section marker rows are always included in the UI regardless of highlighted text matching (they span multiple pages and may not be highlighted).
2. In the UI, `*****` is stripped and only the clean section name is shown in bold with a Yes/No dropdown.
3. Section marker rows are skipped during per-region highlight matching to avoid false matches.
4. When the user answers **"No"**: everything between and including both `*****<name>*****` marker paragraphs is removed from the Word document. If a marker shares a paragraph with other markers, only the marker text is stripped (not the whole paragraph).
5. When the user answers **"Yes"**: only the marker text `*****<name>*****` is stripped from the paragraphs; all content between them is kept. If stripping leaves an empty paragraph, that paragraph is removed.
6. Section markers can be used as parent clauses (via Column J) to cascade the "No" answer to child rows.

**Example 1: Simple (each marker on its own line)**

Excel row:
| Row# | Clause | SOW Text | Tips |
|------|--------|----------|------|
| 15 | 6a | `*****Onboarding Services*****` | Please select Yes or No |

Word document:
```
... previous content ...
*****Onboarding Services*****
Heading: Onboarding Services Overview
Paragraph: The following onboarding services will be provided...
... (multiple pages of content) ...
*****Onboarding Services*****
... subsequent content ...
```

- User selects **Yes** → The two `*****Onboarding Services*****` lines are removed, all content between them is kept.
- User selects **No** → Everything from the first `*****Onboarding Services*****` through the second `*****Onboarding Services*****` (inclusive) is deleted.

**Example 2: Shared paragraph (multiple markers on the same line)**

Excel rows:
| Row# | Clause | SOW Text | Tips |
|------|--------|----------|------|
| 10 | Sch_A | `*****A_Schedule*****` | Please select Yes or No |
| 11 | Sch_B | `*****B_Schedule*****` | Please select Yes or No |

Word document:
```
*****A_Schedule***** *****B_Schedule*****
Content of Schedule A...
*****A_Schedule*****
Content of Schedule B...
*****B_Schedule*****
```

Here the first paragraph contains both `*****A_Schedule*****` and `*****B_Schedule*****`.

- User selects **Yes** for A, **No** for B:
  - A: Marker text `*****A_Schedule*****` is stripped from both paragraphs. Content between them is kept.
  - B: The first paragraph is shared, so only `*****B_Schedule*****` text is stripped from it (not the whole paragraph). Content between the B markers is deleted. The end `*****B_Schedule*****` paragraph is removed.
  - Result: `Content of Schedule A...` remains.

- User selects **No** for A, **Yes** for B:
  - A: The first paragraph is shared, so `*****A_Schedule*****` text is stripped. Everything between A markers (including "Content of Schedule A...") is deleted. The end `*****A_Schedule*****` paragraph is removed.
  - B: Marker text `*****B_Schedule*****` is stripped from the surviving paragraphs. Content between them is kept.
  - Result: `Content of Schedule B...` remains.

If another row has `ParentClauses = "Sch_A"` in Column J, selecting "No" for Sch_A will also hide that child row in the UI and auto-set its answer to "No".

---

## Parent-Child Clause Dependencies (Column J)

Column J in the Excel file defines parent clause dependencies. When all parents of a clause are answered "No", the child clause is automatically hidden in the UI and its answer is set to "No".

**Format:** Comma-separated clause numbers referencing the Clause Number (Column C) of parent rows.

**Rules:**
1. If a child has a single parent (e.g., `6a`), it is hidden when that parent is "No".
2. If a child has multiple parents (e.g., `6a,6b`), it is hidden only when **ALL** parents are "No".
3. Hidden child rows have their answer auto-set to "No", which triggers the standard "No" processing (paragraph removal) during document generation.
4. The dependency check runs on every answer change in the UI.

**Example:**

| Row# | Clause | SOW Text | Parent Clauses (J) |
|------|--------|----------|---------------------|
| 10 | 6a | `*****Onboarding*****` | |
| 11 | 6b | `[some highlighted text]` | |
| 12 | 6c | `[another clause]` | `6a` |
| 13 | 6d | `[depends on both]` | `6a,6b` |

- User selects "No" for 6a → Row 12 (6c) is hidden and auto-set to "No". Row 13 (6d) remains visible because 6b is not "No".
- User then selects "No" for 6b → Row 13 (6d) is now also hidden and auto-set to "No" (both parents are "No").

---

## Table Row Markers (`&&&&&table<name>&&&&&`)

Table row markers allow individual rows within a Word table to be conditionally included, modified, or removed. The Excel clause column uses a `table<name>` prefix to identify these rows, and the Word document uses `&&&&&table<name>&&&&&` delimiters to mark which table the row belongs to.

**Excel format:**
- Column C (Clause): `table<name>` (e.g., `tableA`, `tableSLA`). The `table` prefix is case-insensitive.
- Column D (SOW Text): The text content of a specific cell within the Word table (may contain `[placeholder]` patterns for user input).
- Column J (Parent Clauses): A parent clause number. If the parent answers "No", the entire Word table row (`<w:tr>`) containing the matched text is deleted.

**Word document format:**
- The Word document contains marker text `&&&&&table<name>&&&&&` (exactly 5 ampersands on each side) placed before and after the table.
- The markers identify which table the `table<name>` Excel rows refer to.
- The marker text is always stripped from the final document regardless of user answers.

**Rules:**
1. The `&&&&&table<name>&&&&&` marker text is stripped from the Word document early in processing (before answer matching), so it does not interfere with highlighted region detection.
2. Table rows are matched by comparing their SOW Text (Column D) against highlighted text within Word table cells, using the same matching logic as regular rows.
3. If the SOW Text contains `[placeholder]` patterns, the user can fill them in the UI as usual (textboxes, nested placeholders, etc.).
4. If the matched row has no placeholders and no parent dependency, the highlighted text is kept as-is (highlight formatting removed).
5. When the parent clause answers **"No"**: the entire Word table row (`<w:tr>`) containing the matched text is removed -- even if the row has multiple columns.
6. When the parent clause answers **"Yes"** (or no parent): the row is kept, placeholders are replaced, and highlight formatting is removed.

**Example:**

Excel rows:
| Row# | Clause | SOW Text | Parent (J) |
|------|--------|----------|------------|
| 10 | Sch_B | `*****B_Schedule*****` | |
| 11 | tableB | `Allocation of [Pool Percentage]` | `Sch_B` |
| 12 | tableB | `Critical Service Level [definition]` | `Sch_B` |

Word document:
```
&&&&&tableB&&&&&
| Term | Definition |
| Allocation of [Pool Percentage] | A quantity stated as a number... |
| Critical Service Level [definition] | The minimum acceptable level... |
| Agreement | As defined in the preamble... |
&&&&&tableB&&&&&
```

- User selects **Yes** for Sch_B:
  - Row 11: User fills `[Pool Percentage]` → replaced in the table cell.
  - Row 12: User fills `[definition]` → replaced in the table cell.
  - `&&&&&tableB&&&&&` markers are stripped. Table content is kept.

- User selects **No** for Sch_B:
  - Row 11 and 12 auto-set to "No" (parent dependency).
  - The entire `<w:tr>` for "Allocation of [Pool Percentage]" is deleted.
  - The entire `<w:tr>` for "Critical Service Level [definition]" is deleted.
  - The "Agreement" row is NOT affected (no Excel row targets it).
  - `&&&&&tableB&&&&&` markers are stripped.
  - If Sch_B is a section marker, the B_Schedule section is also removed.
