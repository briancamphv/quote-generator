# TeamLogic IT of Sherman-Celina — Quote Generator

A Python program that builds branded Word (`.docx`) quotes in two modes:

1. **Managed Services / MRR** — reads the *Managed Services Estimator v4.4*
   workbook and pulls the proposal-safe (rounded) figures from the **Results**
   sheet, then writes a managed-services proposal with services, pricing, and
   onboarding.
2. **Time & Materials** — takes an hourly rate (plus after-hours multiplier,
   minimum, and billing increment) and writes a T&M proposal with a rate table
   and auto-computed billing examples.

Both proposals use TeamLogic brand styling (logo, PMS 348 green, Montserrat),
your contact block, and the MSA + SOW "Next Steps" language.

---

## Setup

```bash
pip install python-docx openpyxl --break-system-packages
```

Keep `logo.png` (the white-background TeamLogic IT lockup) in the same folder as
`quote_generator.py`. Optional: install LibreOffice (`soffice`) for PDF export.

---

## Run

```bash
python3 quote_generator.py
```

The program asks you a few questions in the terminal:

**MRR mode** asks for the path to the estimator `.xlsm`, lets you confirm the
prospect name, and lets you choose which price(s) to show (both / onsite / remote).

**T&M mode** asks for the client name, business-hours rate, after-hours uplift %,
minimum hours, billing increment, and the business-hours window. Everything has a
sensible default (150 / 50% / 1 hr / 30 min) — press Enter to accept.

Output `.docx` (and optional `.pdf`) lands in the folder you specify.

---

## Important: estimator must be saved by Excel

The estimator's **Results** sheet is built on formulas. Excel caches the computed
values when you save. This program reads those cached values directly and does
**not** recalculate — because a headless recalc (LibreOffice) can zero-out the
estimator's Excel-specific formulas.

So: **open the estimator in Excel, enter your inputs, save, then point this
program at that saved file.** If the MRR cells read as empty/zero, the program
warns you that the file likely wasn't last saved by Excel.

---

## What it reads from the estimator

By label (not fixed cell coordinates, so minor layout shifts won't break it):

- Prospect name (`Prospect:`)
- Environment inputs (Servers, Computers, Employees, Mailboxes, Incidents, TB, …)
- Rounded MRR — *Managed Services with Onsite & Remote Support*
- Rounded MRR — *Managed Services with Remote Support*
- Rounded *Onboarding* fee

It intentionally uses only the **rounded, proposal-safe** numbers — never the
internal cost/margin figures the Results sheet flags as "DO NOT transpose."

---

## Customizing the service list

The managed-services feature list is the `MS_SERVICES` constant near the top of
the builder section. Edit names/descriptions there to match your current stack
(RMM, MDR vendor, password manager, email security, etc.).
