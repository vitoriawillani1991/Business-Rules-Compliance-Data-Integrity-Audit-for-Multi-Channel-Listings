# Title & Bullet Validation Engine  
**Business Rules Compliance & Data Integrity Audit for Multi-Channel Listings**

## Overview

This project validates already-built product titles and bullet points against:

1. **Business Rules Templates** (structure compliance)  
2. **NetSuite Saved Search data** (data integrity compliance)  

It ensures that marketplace listings (Amazon, eBay, Walmart, Webstore, Google) strictly follow predefined templates **and** that all inserted placeholder values match authoritative system data.

This is not a content generator.  
It is a **quality control and risk prevention engine** for catalog publishing pipelines.

---

# Business Impact

## 1. Reduces Manual QA Work
Replaces manual text verification with automated SKU-level validation.

## 2. Prevents Incorrect Listings Going Live
Detects:
- Wrong position (Front vs Rear)
- Incorrect side (Left/Right, Driver/Passenger)
- Incorrect ABS flags
- Wrong numeric specs (diameter, length, piston size)
- Invalid kit member data
- Template structure violations

This reduces:
- Returns
- Customer complaints
- Marketplace compliance risks
- Brand inconsistency

## 3. Improves Publishing Speed
The script outputs a structured Excel report identifying:
- Exactly which field failed
- Whether the issue is structure or data
- The expected values
- The data source used for validation

This allows fast correction with minimal investigation.

## 4. Enables Scalable Governance
The system is template-driven:
- Any new rule added to the Business Rules workbook automatically becomes enforceable.
- No need to rewrite validation logic per template.

---

# What the Script Does

For each SKU in the Titles/Bullets input file:

---

## Step 1 — Select the Correct Template

Matches:

Titles["SKU Type"] == Templates["SKU Type Updated"]


If no match is found:
- `rule_status = NO_RULE_MATCH`

---

## Step 2 — Validate Each Text Field

Fields validated (if present in both files):

- Amazon Title  
- Amazon Bullet Point 1–5  
- eBay Title  
- eBay Subtitle  
- eBay Description  
- Walmart Title  
- Webstore Title  
- Google Title  

---

# Validation Layers

## A) STRUCTURE VALIDATION (Template Compliance)

The template is converted into a tolerant regex pattern:

- Case-insensitive
- Whitespace tolerant
- Punctuation tolerant
- Placeholders `[LIKE THIS]` converted into capture groups
- Special handling for `OUTER DIAMETER MM` to reduce regex drift

If the built text does not match the template structure:

rule_status = ERROR
rule_error = STRUCTURE_FAIL


Even when structure fails, placeholder validation still runs in fallback mode.

---

## B) PLACEHOLDER DATA VALIDATION

If structure matches → CAPTURE mode  
If structure fails → FALLBACK mode

### CAPTURE Mode
- Extracts captured placeholder values
- Validates each against expected candidates

### FALLBACK Mode
- Does not rely on regex captures
- Confirms expected values exist somewhere in the text

---

# Data Sources Used for Validation

## 1. Saved Search (NetSuite)
Used for:
- Inventory Item
- Assembly/Bill of Materials

## 2. Parts_in_Package
For Kit/Package SKUs:
- Expands members
- Validates against member-level data
- Any-match wins logic

## 3. Position List
Used for:
- `[Position]`
- `[Left/Right]`
- `[Driver/Passenger]`
- `[Position for <Category>]`

---

# Special Placeholder Behavior

## Ignored Placeholders
Completely ignored during validation:

[2PIECE DESIGN M1]
[2PIECE DESIGN M2]
[2PIECE DESIGN M3]
[2PIECE DESIGN]
[SKU M1]
[SKU M2]
[SKU M3]
[YMM]
[MMY]


---

## Boolean Placeholders (Defined in "Boolean" Sheet)

Behavior:

If Saved Search contains a boolean value:
- True  → must match Boolean.Yes
- False → must match Boolean.No (may be blank)

If Saved Search value is:
- Missing
- Blank
- Not a boolean

Then:
- Keeping the literal placeholder is allowed

---

## [MIRROR COLOR]

- Black / Chrome / Gray / Satin → must match value
- "Paint to Match" → placeholder must be deleted
- Missing → literal placeholder allowed

---

## [MIRROR FEATURES]

- "Extendable, Heated" or "Heated" → expected "Heated"
- "Extendable" → expected "Non-Heated"
- Missing → literal placeholder allowed

---

## Position Logic

### [Left/Right] and [Driver/Passenger]

Derived from Position List:

| Positions Found | Expected Text |
|-----------------|--------------|
| Generic (Front/Rear only) | Left or Right |
| Left only | Left |
| Right only | Right |
| Left + Right | Left or Right |

Driver/Passenger follows the same mapping logic.

---

### [Position] Enhancements

If both Front and Rear are valid:
- Accepts:
  - "Front and Rear"
  - "Front or Rear"

If side placeholders exist in any field:
- `[Position]` must exclude side labels consistently.

---

# Numeric Validation Logic

Supports:

- Unit tolerance (mm vs inches)
- Rounding tolerance
- Flexible numeric matching
- Dual-measure interpretation

Example:
12.36" (314 mm)
Validated against:
12.36 in / 314 mm

---

# Output

The script generates an Excel report with:

One row per (SKU, FIELD) with issues.

Columns:

- SKU  
- Type  
- Rule Key  
- Field  
- rule_status (OK / ERROR / NO_RULE_MATCH)  
- placeholder_status (OK / ERROR / SKIPPED)  
- placeholder_check_mode (CAPTURE / FALLBACK / SKIPPED)  
- rule_error  
- placeholder_error  

Only problematic rows are written to the output file.

---

# File Structure

The script expects:

- Business Rules workbook (Templates + Boolean sheets)
- Titles/Bullets input file
- Saved Search export
- Parts_in_Package export
- Position list CSV

All file paths are configured at the top of the script.

---

# Technical Design Highlights

- Template-driven validation (no hardcoded rule logic)
- Flexible regex engine
- Intelligent fallback validation
- Kit expansion support
- Side-aware position logic
- Boolean rule abstraction layer
- Dual-measure numeric matching
- Detailed error traceability

