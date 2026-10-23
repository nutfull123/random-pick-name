# Random Pick Name Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** หน้าเว็บไฟล์เดียวที่นำเข้า Excel/CSV ให้ผู้ใช้เลือกคอลัมน์ชื่อ-นามสกุล แล้วสุ่มชื่อ (ตัดออก/สุ่มซ้ำได้) เผยแพร่เป็น Artifact

**Architecture:** `random-pick-name/index.html` ไฟล์เดียว: CSS + markup + `<script>` ที่แบ่งเป็นฟังก์ชันล้วน (parse/extract/random/pick) กับส่วน UI. Self-check รันเมื่อ URL มี `?test` และพิมพ์ผลลง console + `document.title`

**Tech Stack:** HTML/CSS/vanilla JS, SheetJS `https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js`

**Spec:** `random-pick-name/docs/superpowers/specs/2026-10-23-random-pick-name-design.md`

## Global Constraints

- Single file, no build step; only external script is SheetJS 0.18.5 from cdnjs
- รองรับ `.xlsx`, `.xls`, `.csv` (UTF-8 / UTF-8 BOM); TIS-620 CSV out of scope
- ข้อความ UI ภาษาไทยตามสเปก (copy ใน spec ใช้ตามตัวอักษร)
- ชื่อจากไฟล์แสดงด้วย `textContent` เท่านั้น ห้าม `innerHTML`
- สุ่มด้วย `crypto.getRandomValues` + rejection sampling
- Colors as CSS tokens on `:root`, dark mode via `prefers-color-scheme` (guarded by `:root:not([data-theme="light"])`) and `:root[data-theme="dark"]`; 16px gutter, no horizontal scroll at 375px
- `prefers-reduced-motion` → no spin animation
- Not a git repo: no commit steps (offer `git init` only if user asks)

## Review Focus

1. Space ขณะ focus อยู่บนปุ่ม/สวิตช์ → ต้องไม่สุ่มสองครั้ง (คีย์ลัดเฉพาะเมื่อ focus ไม่อยู่บน control)
2. กดสุ่มระหว่างแอนิเมชันยังไม่จบ → ไม่เริ่มรอบใหม่ ไม่ตัดชื่อเกิน
3. ชีตที่มีแค่แถวหัวตาราง → "คอลัมน์นี้ไม่มีรายชื่อ" ไม่ใช่หน้าสุ่มว่าง
4. CSV UTF-8 BOM → หัวคอลัมน์แรกไม่มี `﻿` ติดมา และภาษาไทยไม่เพี้ยน
5. ชื่อยาวมากบนจอ 375px → ตัดบรรทัดได้ ไม่ล้นแนวนอน

---

### Task 1: Core logic + self-check

**Files:**
- Create: `random-pick-name/index.html` (skeleton + `<script>` core functions + test runner)

**Interfaces:**
- Produces:
  - `async parseFile(file: File): Promise<{ [sheetName: string]: string[][] }>` — `.csv`: `TextDecoder('utf-8')` (strips BOM) → `XLSX.read(text, {type:'string'})`; else `XLSX.read(await file.arrayBuffer(), {type:'array'})`. Each sheet → `sheet_to_json(ws, {header:1, raw:false, defval:''})`. Unsupported extension or SheetJS throw → `throw new Error('อ่านไฟล์นี้ไม่ได้ กรุณาใช้ .xlsx, .xls หรือ .csv')`
  - `extractNames(rows: string[][], col: number): string[]` — skip row 0, `String(v).replace(/\s+/g,' ').trim()`, drop empty; no dedupe
  - `randomIndex(n: number): number` — uniform in `[0, n)`
  - `pick(pool: string[], removeAfter: boolean): { name: string, pool: string[] }` — returns new pool (no mutation); throws on empty pool

- [ ] **Step 1: Write self-check `runTests()`** (runs when `location.search.includes('test')`)

```js
eq(extractNames([['เลขที่','ชื่อ-นามสกุล'],['1','  นาย  สมชาย   ใจดี '],['2',''],['3','น.ส. สมหญิง รักเรียน']], 1),
   ['นาย สมชาย ใจดี','น.ส. สมหญิง รักเรียน'])
eq(extractNames([['ชื่อ']], 0), [])
eq(extractNames([['h'],['ก'],['ก']], 0), ['ก','ก'])               // duplicates kept
for (i<1000) assert(0 <= randomIndex(3) < 3)
// remove mode: 3 picks from 3 → all distinct, pool empty
// repeat mode: pool length unchanged
// parse xlsx: XLSX.write(aoa workbook with 2 sheets, {type:'array',bookType:'xlsx'}) → new File → parseFile → both sheet names, Thai cell intact
// parse csv: new File(['﻿ชื่อ-นามสกุล,ห้อง\nนาย ก ข,1'], 'a.csv') → rows[0][0] === 'ชื่อ-นามสกุล'
```

Result: `document.title = 'TESTS PASS'` or `'TESTS FAIL: <first failure>'`, details to `console`.

- [ ] **Step 2: Open `index.html?test` in browser pane** — Expected: `TESTS FAIL` (functions not defined)
- [ ] **Step 3: Implement the four functions above**

```js
// randomIndex: rejection sampling to avoid modulo bias
const limit = Math.floor(0x100000000 / n) * n
do { crypto.getRandomValues(buf) } while (buf[0] >= limit)
return buf[0] % n
```

- [ ] **Step 4: Reload `index.html?test`** — Expected: title `TESTS PASS`, no console errors

### Task 2: Import + column selection UI

**Files:**
- Modify: `random-pick-name/index.html`

**Interfaces:**
- Consumes: `parseFile`, `extractNames`
- Produces: `startPicking(names: string[])` called by "ใช้รายชื่อนี้"; `showStep(id: 'import'|'columns'|'picker')`

- [ ] **Step 1: Markup + wiring**
  - `#import`: drop zone (`dragover`/`drop`) + `<input type=file accept=".xlsx,.xls,.csv">` with visible label; error `<p role="alert">`
  - `#columns`: sheet `<select>` (hidden when 1 sheet), preview table of first 5 rows; each header cell is a `<button aria-pressed>`; status "พบ N รายชื่อ" / "คอลัมน์นี้ไม่มีรายชื่อ"; button "ใช้รายชื่อนี้" disabled until N > 0; "เปลี่ยนไฟล์" back to import
  - empty sheet → "ชีตนี้ไม่มีข้อมูล"
- [ ] **Step 2: Verify in browser** — simulate drop via JS (`DataTransfer` + File built with `XLSX.write`, 2 sheets, header-only sheet included). Expected: sheet select shows 2 sheets; clicking header shows correct count; header-only sheet shows "คอลัมน์นี้ไม่มีรายชื่อ" and disabled button (Review Focus 3); CSV BOM header shows clean (Review Focus 4)

### Task 3: Picker screen

**Files:**
- Modify: `random-pick-name/index.html`

**Interfaces:**
- Consumes: `pick`, `startPicking(names)`
- State: `{ allNames, pool, history, removeAfter = true, spinning = false }`

- [ ] **Step 1: Implement**
  - Big name display (`aria-live="polite"`), button "สุ่ม", switch "ตัดชื่อที่สุ่มได้ออก" (checked default), "รีเซ็ต", "เปลี่ยนไฟล์"; side panel "เหลืออีก N คน" + "ประวัติที่สุ่มได้" (newest first, `<ol>`)
  - Spin: decide winner via `pick` first; then cycle random names from `allNames`, interval growing ~40ms→200ms over ~1.5s, land on winner. Skip animation under `matchMedia('(prefers-reduced-motion: reduce)')`
  - `spinning` guard: "สุ่ม" disabled and Space ignored while spinning (Review Focus 2)
  - Space shortcut only when `document.activeElement` is not a button/input/select (Review Focus 1); `preventDefault` to stop page scroll
  - Any toggle change: `pool = [...allNames]` (history kept; removal counts only picks made while the switch is on)
  - Reset: `pool = [...allNames]`, `history = []`
  - Pool empty → "สุ่ม" disabled, show "สุ่มครบทุกคนแล้ว — กดรีเซ็ตเพื่อเริ่มใหม่"
- [ ] **Step 2: Verify in browser** — with 3 names, remove on: click สุ่ม 3× → 3 distinct in history, then disabled + message; reset → 3 left; repeat mode: count stays 3; focus switch + Space → no pick; double-click สุ่ม fast → only one history entry; `?test` still PASS

### Task 4: Styling, responsive, publish

**Files:**
- Modify: `random-pick-name/index.html`

- [ ] **Step 1: Style** — load `artifact-design` + `ui-ux-pro-max` guidance first. `<title>สุ่มชื่อ</title>`; tokens on `:root` with dark overrides per Global Constraints; Thai font from Google Fonts (e.g. Noto Sans Thai / IBM Plex Sans Thai); name display `clamp()` size with `overflow-wrap:anywhere` (Review Focus 5); two-column picker on desktop, stacked on mobile; visible focus rings; touch targets ≥ 44px
- [ ] **Step 2: Verify** — browser pane at 375×812 and desktop, light + dark: no horizontal scroll, long name (60 chars) wraps; `?test` PASS; no console errors
- [ ] **Step 3: Publish** — `Artifact` publish `random-pick-name/index.html`, `icon: "shuffle"`, description one sentence; share link with user
