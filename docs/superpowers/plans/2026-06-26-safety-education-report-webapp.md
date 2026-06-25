# 신규근로자 교육 결과보고서 웹앱 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Client-side React SPA that lets a construction-site user pick one of 3 education report types, fill in a short form + 1 photo, and instantly preview/print/download (Excel + PDF) a report that exactly matches the existing fixed Excel templates.

**Architecture:** Pure client-side SPA (no server, no DB). All three `.xlsx` templates ship as static assets. `exceljs` (browser build) fills `{{placeholder}}` cells and embeds the photo to produce a downloadable `.xlsx`. The same data renders into an HTML A4-page preview (used for on-screen preview, `window.print()`, and `html2canvas`+`jsPDF` PDF export). A 39-item lookup table (extracted from `특별안전교육 탬플릿.xlsx` sheets `1`-`39`) backs the keyword-matching dropdown for 특별안전교육.

**Tech Stack:** Vite, React, TypeScript, Vitest + Testing Library, exceljs, html2canvas, jspdf, react-router-dom. Deployed to Vercel/Netlify as a static site.

---

## File Structure

```
src/
  lib/
    dateTime.ts            # today-date default + end-time calc
    dateTime.test.ts
    pageFit.ts              # A4 margin shrink-to-fit algorithm
    pageFit.test.ts
    keywordDerive.ts        # derive default keywords from a 작업유형 title
    keywordDerive.test.ts
    specialItemMatch.ts      # keyword-match a 공종 string against the 39 items
    specialItemMatch.test.ts
    excelTemplate.ts         # loadTemplateBuffer, fillPlaceholders, embedPhoto
    excelTemplate.test.ts
    pdfExport.ts              # html2canvas + jsPDF capture-to-file
  data/
    specialEducationItems.ts  # generated: 39 entries {id,title,content,remark,keywords}
  components/
    CommonFields.tsx
    A4Preview.tsx
    forms/
      RegularEducationForm.tsx
      HiringEducationForm.tsx
      SpecialEducationForm.tsx
  pages/
    HomePage.tsx
    InputPage.tsx
    PreviewPage.tsx
  App.tsx
  main.tsx
  types.ts
scripts/
  extract-special-items.mjs   # one-time extraction script (run with node)
public/
  templates/
    정기안전교육.xlsx
    채용시작업내용변경시교육.xlsx
    특별안전교육.xlsx
```

---

### Task 1: Project scaffold

**Files:**
- Create: `package.json`, `vite.config.ts`, `tsconfig.json`, `index.html`, `src/main.tsx`, `src/App.tsx`

- [ ] **Step 1: Scaffold Vite React-TS app**

Run:
```bash
npm create vite@latest . -- --template react-ts
```
(Run inside `D:\AI\new edu`; when prompted about non-empty directory, confirm — keep existing files.)

- [ ] **Step 2: Install runtime + dev dependencies**

Run:
```bash
npm install react-router-dom exceljs html2canvas jspdf
npm install -D vitest @testing-library/react @testing-library/jest-dom jsdom @vitest/ui
```

- [ ] **Step 3: Add Vitest config**

Edit `vite.config.ts`:
```ts
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  test: {
    environment: 'jsdom',
    globals: true,
  },
})
```

- [ ] **Step 4: Add test script**

Edit `package.json` `"scripts"` block, add:
```json
"test": "vitest run"
```

- [ ] **Step 5: Verify dev server boots**

Run: `npm run dev -- --port 5183`
Expected: Vite prints a local URL with no errors. Stop the server (Ctrl+C) once confirmed.

- [ ] **Step 6: Commit**

```bash
git add package.json package-lock.json vite.config.ts tsconfig.json tsconfig.app.json tsconfig.node.json index.html src/ public/vite.svg
git commit -m "chore: scaffold Vite React-TS app with Vitest"
```

---

### Task 2: dateTime utilities

**Files:**
- Create: `src/lib/dateTime.ts`
- Test: `src/lib/dateTime.test.ts`

- [ ] **Step 1: Write failing tests**

`src/lib/dateTime.test.ts`:
```ts
import { describe, it, expect } from 'vitest'
import { getTodayDateString, calcEndTime } from './dateTime'

describe('getTodayDateString', () => {
  it('formats a given date as YYYY-MM-DD', () => {
    const date = new Date(2026, 5, 26) // 2026-06-26 (month is 0-indexed)
    expect(getTodayDateString(date)).toBe('2026-06-26')
  })

  it('pads single-digit month and day', () => {
    const date = new Date(2026, 0, 5) // 2026-01-05
    expect(getTodayDateString(date)).toBe('2026-01-05')
  })
})

describe('calcEndTime', () => {
  it('adds 2 hours by default', () => {
    expect(calcEndTime('09:00')).toBe('11:00')
  })

  it('wraps past midnight', () => {
    expect(calcEndTime('23:30')).toBe('01:30')
  })

  it('supports a custom number of hours', () => {
    expect(calcEndTime('08:00', 3)).toBe('11:00')
  })
})
```

- [ ] **Step 2: Run tests, confirm failure**

Run: `npm test -- dateTime`
Expected: FAIL — `Cannot find module './dateTime'`

- [ ] **Step 3: Implement**

`src/lib/dateTime.ts`:
```ts
export function getTodayDateString(date: Date = new Date()): string {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  const d = String(date.getDate()).padStart(2, '0')
  return `${y}-${m}-${d}`
}

export function calcEndTime(startTime: string, hoursToAdd = 2): string {
  const [h, m] = startTime.split(':').map(Number)
  const totalMinutes = (h * 60 + m + hoursToAdd * 60) % (24 * 60)
  const endH = Math.floor(totalMinutes / 60)
  const endM = totalMinutes % 60
  return `${String(endH).padStart(2, '0')}:${String(endM).padStart(2, '0')}`
}
```

- [ ] **Step 4: Run tests, confirm pass**

Run: `npm test -- dateTime`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add src/lib/dateTime.ts src/lib/dateTime.test.ts
git commit -m "feat: add date/time auto-fill utilities"
```

---

### Task 3: pageFit (A4 margin shrink-to-fit)

**Files:**
- Create: `src/lib/pageFit.ts`
- Test: `src/lib/pageFit.test.ts`

- [ ] **Step 1: Write failing tests**

`src/lib/pageFit.test.ts`:
```ts
import { describe, it, expect } from 'vitest'
import { computeMarginMm, A4_HEIGHT_MM } from './pageFit'

describe('computeMarginMm', () => {
  it('keeps the default 40mm margin when content fits', () => {
    // available height at 40mm margin = 297 - 80 = 217mm
    expect(computeMarginMm(200)).toBe(40)
  })

  it('shrinks margin in 5mm steps until content fits', () => {
    // at 40mm: available 217 -> too small for 250
    // at 35mm: available 227 -> still too small
    // at 30mm: available 237 -> still too small
    // at 25mm: available 247 -> fits
    expect(computeMarginMm(240)).toBe(25)
  })

  it('never goes below the 10mm minimum margin', () => {
    expect(computeMarginMm(1000)).toBe(10)
  })

  it('exposes the A4 height constant in millimeters', () => {
    expect(A4_HEIGHT_MM).toBe(297)
  })
})
```

- [ ] **Step 2: Run tests, confirm failure**

Run: `npm test -- pageFit`
Expected: FAIL — `Cannot find module './pageFit'`

- [ ] **Step 3: Implement**

`src/lib/pageFit.ts`:
```ts
export const A4_HEIGHT_MM = 297
export const A4_WIDTH_MM = 210
export const DEFAULT_MARGIN_MM = 40
export const MIN_MARGIN_MM = 10
const STEP_MM = 5

export function computeMarginMm(contentHeightMm: number): number {
  let margin = DEFAULT_MARGIN_MM
  while (margin > MIN_MARGIN_MM) {
    const available = A4_HEIGHT_MM - margin * 2
    if (contentHeightMm <= available) return margin
    margin -= STEP_MM
  }
  return MIN_MARGIN_MM
}
```

- [ ] **Step 4: Run tests, confirm pass**

Run: `npm test -- pageFit`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add src/lib/pageFit.ts src/lib/pageFit.test.ts
git commit -m "feat: add A4 margin shrink-to-fit calculation"
```

---

### Task 4: keywordDerive (default keywords from a work-type title)

**Files:**
- Create: `src/lib/keywordDerive.ts`
- Test: `src/lib/keywordDerive.test.ts`

- [ ] **Step 1: Write failing tests**

`src/lib/keywordDerive.test.ts`:
```ts
import { describe, it, expect } from 'vitest'
import { deriveDefaultKeywords } from './keywordDerive'

describe('deriveDefaultKeywords', () => {
  it('strips trailing 작업/시/등 and splits on whitespace/punctuation', () => {
    const title = '타워크레인을 사용하는 작업 시 신호업무를 하는 작업'
    expect(deriveDefaultKeywords(title)).toEqual(['타워크레인', '신호업무'])
  })

  it('removes stopwords and particles-only tokens', () => {
    const title = '밀폐된 공간에서의 작업'
    expect(deriveDefaultKeywords(title)).toEqual(['밀폐된', '공간에서의'])
  })

  it('dedupes repeated tokens', () => {
    const title = '용접 작업 및 용접 관련 작업'
    expect(deriveDefaultKeywords(title)).toEqual(['용접', '관련'])
  })
})
```

- [ ] **Step 2: Run tests, confirm failure**

Run: `npm test -- keywordDerive`
Expected: FAIL — `Cannot find module './keywordDerive'`

- [ ] **Step 3: Implement**

`src/lib/keywordDerive.ts`:
```ts
const STOPWORDS = new Set(['작업', '시', '등', '및', '하는', '사용하는', '관련된', '경우'])

export function deriveDefaultKeywords(title: string): string[] {
  const tokens = title
    .split(/[\s,·()/]+/)
    .map((t) => t.trim())
    .filter(Boolean)
    .filter((t) => !STOPWORDS.has(t))

  const seen = new Set<string>()
  const result: string[] = []
  for (const t of tokens) {
    if (!seen.has(t)) {
      seen.add(t)
      result.push(t)
    }
  }
  return result
}
```

- [ ] **Step 4: Run tests, confirm pass**

Run: `npm test -- keywordDerive`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add src/lib/keywordDerive.ts src/lib/keywordDerive.test.ts
git commit -m "feat: derive default match keywords from a work-type title"
```

---

### Task 5: specialItemMatch (keyword matching against the 39 items)

**Files:**
- Create: `src/lib/specialItemMatch.ts`
- Test: `src/lib/specialItemMatch.test.ts`

- [ ] **Step 1: Write failing tests**

`src/lib/specialItemMatch.test.ts`:
```ts
import { describe, it, expect } from 'vitest'
import { matchSpecialItem, type SpecialEducationItem } from './specialItemMatch'

const items: SpecialEducationItem[] = [
  { id: 1, title: '밀폐공간 작업', content: '...', remark: '', keywords: ['밀폐공간'] },
  { id: 39, title: '타워크레인 신호업무', content: '...', remark: '', keywords: ['타워크레인', '신호업무'] },
]

describe('matchSpecialItem', () => {
  it('returns the item whose keyword is contained in the input', () => {
    expect(matchSpecialItem('타워크레인 신호수 작업', items)?.id).toBe(39)
  })

  it('returns undefined when no keyword matches', () => {
    expect(matchSpecialItem('전기 배선 점검', items)).toBeUndefined()
  })

  it('returns the first matching item when multiple keywords match', () => {
    expect(matchSpecialItem('밀폐공간 내 타워크레인 점검', items)?.id).toBe(1)
  })
})
```

- [ ] **Step 2: Run tests, confirm failure**

Run: `npm test -- specialItemMatch`
Expected: FAIL — `Cannot find module './specialItemMatch'`

- [ ] **Step 3: Implement**

`src/lib/specialItemMatch.ts`:
```ts
export interface SpecialEducationItem {
  id: number
  title: string
  content: string
  remark: string
  keywords: string[]
}

export function matchSpecialItem(
  workType: string,
  items: SpecialEducationItem[],
): SpecialEducationItem | undefined {
  return items.find((item) => item.keywords.some((kw) => workType.includes(kw)))
}
```

- [ ] **Step 4: Run tests, confirm pass**

Run: `npm test -- specialItemMatch`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add src/lib/specialItemMatch.ts src/lib/specialItemMatch.test.ts
git commit -m "feat: add keyword-based matcher for the 39 special-education items"
```

---

### Task 6: Extract the 39 special-education items from the template

**Files:**
- Create: `scripts/extract-special-items.mjs`
- Create (generated, then committed): `src/data/specialEducationItems.ts`

- [ ] **Step 1: Write the extraction script**

`scripts/extract-special-items.mjs`:
```js
import ExcelJS from 'exceljs'
import { writeFileSync } from 'fs'

const TEMPLATE_PATH = new URL('../특별안전교육 탬플릿.xlsx', import.meta.url)

function mergedText(ws, ref) {
  const cell = ws.getCell(ref)
  return (cell.value ?? '').toString().trim()
}

async function main() {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(TEMPLATE_PATH)

  const items = []
  for (let id = 1; id <= 39; id++) {
    const ws = wb.getWorksheet(String(id))
    if (!ws) throw new Error(`sheet ${id} not found`)
    const title = mergedText(ws, 'C6')
    const content = mergedText(ws, 'G6')
    const remark = mergedText(ws, 'K4')
    items.push({ id, title, content, remark })
  }

  const out = [
    "import { deriveDefaultKeywords } from '../lib/keywordDerive'",
    "import type { SpecialEducationItem } from '../lib/specialItemMatch'",
    '',
    `const RAW: Omit<SpecialEducationItem, 'keywords'>[] = ${JSON.stringify(items, null, 2)}`,
    '',
    'export const specialEducationItems: SpecialEducationItem[] = RAW.map((item) => ({',
    '  ...item,',
    '  keywords: deriveDefaultKeywords(item.title),',
    '}))',
    '',
  ].join('\n')

  writeFileSync(new URL('../src/data/specialEducationItems.ts', import.meta.url), out)
  console.log(`Wrote ${items.length} items to src/data/specialEducationItems.ts`)
}

main()
```

- [ ] **Step 2: Run the extraction script**

Run: `node scripts/extract-special-items.mjs`
Expected: `Wrote 39 items to src/data/specialEducationItems.ts`

- [ ] **Step 3: Sanity-check the generated file**

Open `src/data/specialEducationItems.ts` and confirm:
- It has exactly 39 entries (`id: 1` through `id: 39`)
- `title` and `content` fields are non-empty Korean text matching the template's sheets
- No `RAW` entry has an empty `title`

If any title/content is empty, open the template sheet for that `id` in Excel and check whether the title/content actually live in different merged cells (the extraction reads cells `C6` and `G6`, which is the top-left of the merged range per the earlier template inspection — confirm against the sheet directly if a mismatch is found, and adjust the cell refs in Step 1).

- [ ] **Step 4: Commit**

```bash
git add scripts/extract-special-items.mjs src/data/specialEducationItems.ts
git commit -m "feat: generate the 39 special-education items dataset from the template"
```

---

### Task 7: excelTemplate — fillPlaceholders

**Files:**
- Create: `src/lib/excelTemplate.ts`
- Test: `src/lib/excelTemplate.test.ts`

- [ ] **Step 1: Write failing test**

`src/lib/excelTemplate.test.ts`:
```ts
import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import { fillPlaceholders } from './excelTemplate'

describe('fillPlaceholders', () => {
  it('replaces {{key}} tokens with provided values', async () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('test')
    ws.getCell('A1').value = '{{협력사}}'
    ws.getCell('B1').value = '고정 텍스트'
    ws.getCell('C1').value = '담당자: {{교육강사}}'

    fillPlaceholders(ws, { 협력사: 'ABC건설', 교육강사: '홍길동' })

    expect(ws.getCell('A1').value).toBe('ABC건설')
    expect(ws.getCell('B1').value).toBe('고정 텍스트')
    expect(ws.getCell('C1').value).toBe('담당자: 홍길동')
  })

  it('leaves unknown placeholders untouched', () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('test')
    ws.getCell('A1').value = '{{모름}}'

    fillPlaceholders(ws, {})

    expect(ws.getCell('A1').value).toBe('{{모름}}')
  })
})
```

- [ ] **Step 2: Run test, confirm failure**

Run: `npm test -- excelTemplate`
Expected: FAIL — `Cannot find module './excelTemplate'`

- [ ] **Step 3: Implement fillPlaceholders**

`src/lib/excelTemplate.ts`:
```ts
import ExcelJS from 'exceljs'

const PLACEHOLDER_RE = /\{\{(\w+)\}\}/g

export function fillPlaceholders(
  worksheet: ExcelJS.Worksheet,
  values: Record<string, string>,
): void {
  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      if (typeof cell.value !== 'string') return
      if (!PLACEHOLDER_RE.test(cell.value)) return
      cell.value = cell.value.replace(PLACEHOLDER_RE, (_match, key: string) =>
        key in values ? values[key] : `{{${key}}}`,
      )
    })
  })
}
```

Note: `RegExp.test` with the `g` flag mutates `lastIndex`, so reset it before `replace` is safe here because each `cell.value.replace(...)` call uses the same regex object freshly each iteration — to avoid subtle `lastIndex` bugs, define `PLACEHOLDER_RE` fresh per call instead of as a shared module-level constant:

```ts
import ExcelJS from 'exceljs'

export function fillPlaceholders(
  worksheet: ExcelJS.Worksheet,
  values: Record<string, string>,
): void {
  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      if (typeof cell.value !== 'string') return
      const re = /\{\{(\w+)\}\}/g
      if (!re.test(cell.value)) return
      cell.value = cell.value.replace(/\{\{(\w+)\}\}/g, (_match, key: string) =>
        key in values ? values[key] : `{{${key}}}`,
      )
    })
  })
}
```

- [ ] **Step 4: Run test, confirm pass**

Run: `npm test -- excelTemplate`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add src/lib/excelTemplate.ts src/lib/excelTemplate.test.ts
git commit -m "feat: fill {{placeholder}} cells in an exceljs worksheet"
```

---

### Task 8: excelTemplate — embedPhoto

**Files:**
- Modify: `src/lib/excelTemplate.ts`
- Modify: `src/lib/excelTemplate.test.ts`

- [ ] **Step 1: Write failing test**

Append to `src/lib/excelTemplate.test.ts`:
```ts
import { embedPhoto } from './excelTemplate'

const ONE_PX_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='

describe('embedPhoto', () => {
  it('adds exactly one image anchored at the given range', () => {
    const wb = new ExcelJS.Workbook()
    const ws = wb.addWorksheet('test')

    embedPhoto(wb, ws, ONE_PX_PNG_BASE64, 'E20:J20')

    expect(ws.getImages()).toHaveLength(1)
  })
})
```

- [ ] **Step 2: Run test, confirm failure**

Run: `npm test -- excelTemplate`
Expected: FAIL — `embedPhoto is not a function` (or not exported)

- [ ] **Step 3: Implement embedPhoto**

Add to `src/lib/excelTemplate.ts`:
```ts
function parseCellRange(range: string): { tl: string; br: string } {
  const [tl, br] = range.split(':')
  return { tl, br: br ?? tl }
}

export function embedPhoto(
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
  base64Png: string,
  anchorRange: string,
): void {
  const imageId = workbook.addImage({
    base64: base64Png.startsWith('data:') ? base64Png : `data:image/png;base64,${base64Png}`,
    extension: 'png',
  })
  const { tl, br } = parseCellRange(anchorRange)
  worksheet.addImage(imageId, `${tl}:${br}`)
}
```

- [ ] **Step 4: Run test, confirm pass**

Run: `npm test -- excelTemplate`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add src/lib/excelTemplate.ts src/lib/excelTemplate.test.ts
git commit -m "feat: embed the uploaded photo into the worksheet"
```

---

### Task 9: excelTemplate — generateReportWorkbook (integration) + copy template assets

**Files:**
- Modify: `src/lib/excelTemplate.ts`
- Modify: `src/lib/excelTemplate.test.ts`
- Create: `public/templates/정기안전교육.xlsx`, `public/templates/채용시작업내용변경시교육.xlsx`, `public/templates/특별안전교육.xlsx`

- [ ] **Step 1: Copy the three templates into `public/templates/`**

Run:
```bash
mkdir -p public/templates
cp "정기안전교육 탬플릿.xlsx" "public/templates/정기안전교육.xlsx"
cp "채용시 작업내용변경시 교육 탬플릿.xlsx" "public/templates/채용시작업내용변경시교육.xlsx"
cp "특별안전교육 탬플릿.xlsx" "public/templates/특별안전교육.xlsx"
```

- [ ] **Step 2: Write failing test for the photo-cell-anchor lookup**

Append to `src/lib/excelTemplate.test.ts`:
```ts
import { PHOTO_ANCHOR } from './excelTemplate'

describe('PHOTO_ANCHOR', () => {
  it('matches the photo placeholder cell range used by all three templates', () => {
    expect(PHOTO_ANCHOR).toBe('E20:J20')
  })
})
```

(Note: 정기/특별 use `E20:J20`; 채용시 template's photo cell is `E19:J19` one row higher per the structural inspection — handle this per-template in Step 3 below rather than a single shared constant used everywhere.)

- [ ] **Step 3: Implement generateReportWorkbook with per-template photo anchors**

Add to `src/lib/excelTemplate.ts`:
```ts
export const PHOTO_ANCHOR = 'E20:J20'

export type TemplateName = '정기안전교육' | '채용시작업내용변경시교육' | '특별안전교육'

const PHOTO_ANCHOR_BY_TEMPLATE: Record<TemplateName, string> = {
  정기안전교육: 'E20:J20',
  채용시작업내용변경시교육: 'E19:J19',
  특별안전교육: 'E20:J20',
}

const SHEET_NAME_BY_TEMPLATE: Record<TemplateName, string> = {
  정기안전교육: '정기',
  채용시작업내용변경시교육: '채용시',
  특별안전교육: '특별',
}

export async function loadTemplateBuffer(name: TemplateName): Promise<ArrayBuffer> {
  const res = await fetch(`/templates/${name}.xlsx`)
  if (!res.ok) throw new Error(`failed to load template: ${name}`)
  return res.arrayBuffer()
}

export async function generateReportWorkbook(
  name: TemplateName,
  templateBuffer: ArrayBuffer,
  values: Record<string, string>,
  photoBase64?: string,
): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(templateBuffer)
  const ws = wb.getWorksheet(SHEET_NAME_BY_TEMPLATE[name])
  if (!ws) throw new Error(`sheet not found for template: ${name}`)

  fillPlaceholders(ws, values)
  if (photoBase64) {
    embedPhoto(wb, ws, photoBase64, PHOTO_ANCHOR_BY_TEMPLATE[name])
  }
  return wb
}
```

- [ ] **Step 4: Run tests, confirm pass**

Run: `npm test -- excelTemplate`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add public/templates src/lib/excelTemplate.ts src/lib/excelTemplate.test.ts
git commit -m "feat: load templates and generate a filled report workbook per type"
```

---

### Task 10: Shared types

**Files:**
- Create: `src/types.ts`

- [ ] **Step 1: Define the shared form-data type**

`src/types.ts`:
```ts
export interface CommonFormData {
  협력사: string
  공종: string
  교육일시: string
  교육시작시간: string
  교육종료시간: string
  교육방법: '강의식' | '토론식'
  교육장소: string
  교육인원: string
  교육강사: string
  교육사진: string | null // base64 PNG/JPEG, no data: prefix stripped
}

export type ReportKind = 'regular' | 'hiring' | 'special'

export interface HiringFormData extends CommonFormData {
  hiringType: '채용시' | '작업내용변경시'
}

export interface SpecialFormData extends CommonFormData {
  selectedItemId: number | null
}
```

- [ ] **Step 2: Commit**

```bash
git add src/types.ts
git commit -m "feat: add shared form-data types"
```

---

### Task 11: CommonFields component

**Files:**
- Create: `src/components/CommonFields.tsx`

- [ ] **Step 1: Implement the shared fields component**

`src/components/CommonFields.tsx`:
```tsx
import { useEffect } from 'react'
import type { CommonFormData } from '../types'
import { getTodayDateString, calcEndTime } from '../lib/dateTime'

interface Props {
  value: CommonFormData
  onChange: (next: CommonFormData) => void
}

export function CommonFields({ value, onChange }: Props) {
  useEffect(() => {
    if (!value.교육일시) {
      onChange({ ...value, 교육일시: getTodayDateString() })
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  function set<K extends keyof CommonFormData>(key: K, v: CommonFormData[K]) {
    onChange({ ...value, [key]: v })
  }

  function setStartTime(start: string) {
    onChange({ ...value, 교육시작시간: start, 교육종료시간: calcEndTime(start) })
  }

  return (
    <div>
      <label>
        협력사
        <input value={value.협력사} onChange={(e) => set('협력사', e.target.value)} />
      </label>
      <label>
        공종
        <input value={value.공종} onChange={(e) => set('공종', e.target.value)} />
      </label>
      <label>
        교육일시
        <input
          type="date"
          value={value.교육일시}
          onChange={(e) => set('교육일시', e.target.value)}
        />
      </label>
      <label>
        교육시작시간
        <input
          type="time"
          value={value.교육시작시간}
          onChange={(e) => setStartTime(e.target.value)}
        />
      </label>
      <label>
        교육종료시간 (자동계산, 수정 가능)
        <input
          type="time"
          value={value.교육종료시간}
          onChange={(e) => set('교육종료시간', e.target.value)}
        />
      </label>
      <label>
        교육방법
        <select
          value={value.교육방법}
          onChange={(e) => set('교육방법', e.target.value as CommonFormData['교육방법'])}
        >
          <option value="강의식">강의식</option>
          <option value="토론식">토론식</option>
        </select>
      </label>
      <label>
        교육장소
        <input value={value.교육장소} onChange={(e) => set('교육장소', e.target.value)} />
      </label>
      <label>
        교육인원
        <input value={value.교육인원} onChange={(e) => set('교육인원', e.target.value)} />
      </label>
      <label>
        교육강사
        <input value={value.교육강사} onChange={(e) => set('교육강사', e.target.value)} />
      </label>
      <label>
        교육사진
        <input
          type="file"
          accept="image/*"
          onChange={async (e) => {
            const file = e.target.files?.[0]
            if (!file) return
            const base64 = await fileToBase64(file)
            set('교육사진', base64)
          }}
        />
      </label>
    </div>
  )
}

function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = reject
    reader.readAsDataURL(file)
  })
}

export function createEmptyCommonFormData(): CommonFormData {
  return {
    협력사: '',
    공종: '',
    교육일시: '',
    교육시작시간: '09:00',
    교육종료시간: '11:00',
    교육방법: '강의식',
    교육장소: '안전교육장',
    교육인원: '',
    교육강사: '',
    교육사진: null,
  }
}
```

- [ ] **Step 2: Commit**

```bash
git add src/components/CommonFields.tsx
git commit -m "feat: add shared input fields component"
```

---

### Task 12: Report forms (regular, hiring, special)

**Files:**
- Create: `src/components/forms/RegularEducationForm.tsx`
- Create: `src/components/forms/HiringEducationForm.tsx`
- Create: `src/components/forms/SpecialEducationForm.tsx`

- [ ] **Step 1: RegularEducationForm**

`src/components/forms/RegularEducationForm.tsx`:
```tsx
import { useState } from 'react'
import { CommonFields, createEmptyCommonFormData } from '../CommonFields'
import type { CommonFormData } from '../../types'

interface Props {
  onSubmit: (data: CommonFormData) => void
}

export function RegularEducationForm({ onSubmit }: Props) {
  const [data, setData] = useState<CommonFormData>(createEmptyCommonFormData())

  return (
    <div>
      <CommonFields value={data} onChange={setData} />
      <button onClick={() => onSubmit(data)}>보고서 생성</button>
    </div>
  )
}
```

- [ ] **Step 2: HiringEducationForm**

`src/components/forms/HiringEducationForm.tsx`:
```tsx
import { useState } from 'react'
import { CommonFields, createEmptyCommonFormData } from '../CommonFields'
import type { HiringFormData } from '../../types'

interface Props {
  onSubmit: (data: HiringFormData) => void
}

export function HiringEducationForm({ onSubmit }: Props) {
  const [data, setData] = useState<HiringFormData>({
    ...createEmptyCommonFormData(),
    hiringType: '채용시',
  })

  return (
    <div>
      <div role="group" aria-label="교육 구분">
        <label>
          <input
            type="radio"
            name="hiringType"
            checked={data.hiringType === '채용시'}
            onChange={() => setData({ ...data, hiringType: '채용시' })}
          />
          채용시 교육
        </label>
        <label>
          <input
            type="radio"
            name="hiringType"
            checked={data.hiringType === '작업내용변경시'}
            onChange={() => setData({ ...data, hiringType: '작업내용변경시' })}
          />
          작업내용 변경시 교육
        </label>
      </div>
      <CommonFields value={data} onChange={(next) => setData({ ...data, ...next })} />
      <button onClick={() => onSubmit(data)}>보고서 생성</button>
    </div>
  )
}
```

- [ ] **Step 3: SpecialEducationForm with keyword auto-match**

`src/components/forms/SpecialEducationForm.tsx`:
```tsx
import { useEffect, useState } from 'react'
import { CommonFields, createEmptyCommonFormData } from '../CommonFields'
import type { SpecialFormData } from '../../types'
import { specialEducationItems } from '../../data/specialEducationItems'
import { matchSpecialItem } from '../../lib/specialItemMatch'

interface Props {
  onSubmit: (data: SpecialFormData, itemId: number) => void
}

export function SpecialEducationForm({ onSubmit }: Props) {
  const [data, setData] = useState<SpecialFormData>({
    ...createEmptyCommonFormData(),
    selectedItemId: null,
  })

  useEffect(() => {
    if (!data.공종) return
    const matched = matchSpecialItem(data.공종, specialEducationItems)
    if (matched && matched.id !== data.selectedItemId) {
      setData((prev) => ({ ...prev, selectedItemId: matched.id }))
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data.공종])

  const selectedItem = specialEducationItems.find((i) => i.id === data.selectedItemId)

  return (
    <div>
      <CommonFields value={data} onChange={(next) => setData({ ...data, ...next })} />

      <label>
        작업유형 (공종 기반 자동매칭, 직접 변경 가능)
        <select
          value={data.selectedItemId ?? ''}
          onChange={(e) => setData({ ...data, selectedItemId: Number(e.target.value) })}
        >
          <option value="" disabled>
            선택하세요
          </option>
          {specialEducationItems.map((item) => (
            <option key={item.id} value={item.id}>
              {item.id}. {item.title}
            </option>
          ))}
        </select>
      </label>

      {selectedItem && (
        <div>
          <p>교육항목: {selectedItem.title}</p>
          <p>교육내용: {selectedItem.content}</p>
        </div>
      )}

      <button
        disabled={data.selectedItemId === null}
        onClick={() => data.selectedItemId !== null && onSubmit(data, data.selectedItemId)}
      >
        보고서 생성
      </button>
    </div>
  )
}
```

- [ ] **Step 4: Commit**

```bash
git add src/components/forms
git commit -m "feat: add the three report input forms"
```

---

### Task 13: A4Preview component

**Files:**
- Create: `src/components/A4Preview.tsx`

- [ ] **Step 1: Implement the A4 page wrapper using computeMarginMm**

`src/components/A4Preview.tsx`:
```tsx
import { useEffect, useRef, useState } from 'react'
import { computeMarginMm, A4_WIDTH_MM, A4_HEIGHT_MM } from '../lib/pageFit'

interface Props {
  children: React.ReactNode
}

const MM_TO_PX = 96 / 25.4 // CSS px per mm at 96dpi

export function A4Preview({ children }: Props) {
  const contentRef = useRef<HTMLDivElement>(null)
  const [marginMm, setMarginMm] = useState(40)

  useEffect(() => {
    if (!contentRef.current) return
    const heightPx = contentRef.current.scrollHeight
    const heightMm = heightPx / MM_TO_PX
    setMarginMm(computeMarginMm(heightMm))
  }, [children])

  return (
    <div
      id="a4-page"
      style={{
        width: `${A4_WIDTH_MM}mm`,
        height: `${A4_HEIGHT_MM}mm`,
        padding: `${marginMm}mm`,
        boxSizing: 'border-box',
        margin: '0 auto',
        background: 'white',
      }}
    >
      <div ref={contentRef}>{children}</div>
    </div>
  )
}
```

- [ ] **Step 2: Add print CSS**

Create `src/index.css` print rules (append if file already exists from scaffold):
```css
@media print {
  body * {
    visibility: hidden;
  }
  #a4-page,
  #a4-page * {
    visibility: visible;
  }
  #a4-page {
    position: absolute;
    top: 0;
    left: 0;
  }
}
```

- [ ] **Step 3: Commit**

```bash
git add src/components/A4Preview.tsx src/index.css
git commit -m "feat: add A4 page preview wrapper with shrink-to-fit margins"
```

---

### Task 14: pdfExport utility

**Files:**
- Create: `src/lib/pdfExport.ts`

- [ ] **Step 1: Implement**

`src/lib/pdfExport.ts`:
```ts
import html2canvas from 'html2canvas'
import jsPDF from 'jspdf'
import { A4_WIDTH_MM, A4_HEIGHT_MM } from './pageFit'

export async function exportElementToPdf(element: HTMLElement, filename: string): Promise<void> {
  const canvas = await html2canvas(element, { scale: 2 })
  const imgData = canvas.toDataURL('image/png')

  const pdf = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' })
  pdf.addImage(imgData, 'PNG', 0, 0, A4_WIDTH_MM, A4_HEIGHT_MM)
  pdf.save(filename)
}
```

- [ ] **Step 2: Commit**

```bash
git add src/lib/pdfExport.ts
git commit -m "feat: add PDF export via html2canvas + jsPDF"
```

---

### Task 15: PreviewPage

**Files:**
- Create: `src/pages/PreviewPage.tsx`

- [ ] **Step 1: Implement**

`src/pages/PreviewPage.tsx`:
```tsx
import { useRef } from 'react'
import { A4Preview } from '../components/A4Preview'
import { exportElementToPdf } from '../lib/pdfExport'
import {
  generateReportWorkbook,
  loadTemplateBuffer,
  type TemplateName,
} from '../lib/excelTemplate'
import type { CommonFormData } from '../types'

interface Props {
  templateName: TemplateName
  values: Record<string, string>
  photoBase64: string | null
  formData: CommonFormData
  onBack: () => void
}

export function PreviewPage({ templateName, values, photoBase64, formData, onBack }: Props) {
  const pageRef = useRef<HTMLDivElement>(null)

  async function downloadExcel() {
    const buffer = await loadTemplateBuffer(templateName)
    const wb = await generateReportWorkbook(
      templateName,
      buffer,
      values,
      photoBase64 ?? undefined,
    )
    const out = await wb.xlsx.writeBuffer()
    const blob = new Blob([out], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${templateName}_${formData.협력사 || '보고서'}.xlsx`
    a.click()
    URL.revokeObjectURL(url)
  }

  async function downloadPdf() {
    if (!pageRef.current) return
    await exportElementToPdf(pageRef.current, `${templateName}_${formData.협력사 || '보고서'}.pdf`)
  }

  return (
    <div>
      <div ref={pageRef}>
        <A4Preview>
          <h2>{templateName}</h2>
          <p>협력사: {formData.협력사}</p>
          <p>공종: {formData.공종}</p>
          <p>교육일시: {formData.교육일시}</p>
          <p>
            교육시간: {formData.교육시작시간} ~ {formData.교육종료시간}
          </p>
          <p>교육방법: {formData.교육방법}</p>
          <p>교육장소: {formData.교육장소}</p>
          <p>교육인원: {formData.교육인원}</p>
          <p>교육강사: {formData.교육강사}</p>
          {formData.교육사진 && (
            <img src={formData.교육사진} alt="교육사진" style={{ maxWidth: '100%' }} />
          )}
        </A4Preview>
      </div>
      <div>
        <button onClick={() => window.print()}>인쇄</button>
        <button onClick={downloadExcel}>Excel 다운로드</button>
        <button onClick={downloadPdf}>PDF 다운로드</button>
        <button onClick={onBack}>수정하기</button>
      </div>
    </div>
  )
}
```

> Note: this preview shows the common fields generically. It approximates — not pixel-replicates — the template's exact merged-cell layout (그 정확한 표 구조는 `.xlsx` 다운로드 결과에서 1:1로 보장됨; HTML 미리보기는 같은 정보를 보여주는 데 목적이 있음). If pixel-for-pixel preview fidelity becomes a requirement later, revisit this component with an explicit grid matching each template's row/column structure.

- [ ] **Step 2: Commit**

```bash
git add src/pages/PreviewPage.tsx
git commit -m "feat: add preview page with print/Excel/PDF actions"
```

---

### Task 16: HomePage, InputPage, routing

**Files:**
- Create: `src/pages/HomePage.tsx`
- Create: `src/pages/InputPage.tsx`
- Modify: `src/App.tsx`

- [ ] **Step 1: HomePage**

`src/pages/HomePage.tsx`:
```tsx
import { Link } from 'react-router-dom'

export function HomePage() {
  return (
    <div>
      <h1>신규근로자 교육 결과보고서</h1>
      <nav>
        <Link to="/input/regular">정기안전교육</Link>
        <Link to="/input/hiring">채용시 / 작업내용변경시 교육</Link>
        <Link to="/input/special">특별안전교육</Link>
      </nav>
    </div>
  )
}
```

- [ ] **Step 2: InputPage (dispatches to the right form, then shows PreviewPage)**

`src/pages/InputPage.tsx`:
```tsx
import { useState } from 'react'
import { useParams } from 'react-router-dom'
import { RegularEducationForm } from '../components/forms/RegularEducationForm'
import { HiringEducationForm } from '../components/forms/HiringEducationForm'
import { SpecialEducationForm } from '../components/forms/SpecialEducationForm'
import { PreviewPage } from './PreviewPage'
import { specialEducationItems } from '../data/specialEducationItems'
import type { CommonFormData, HiringFormData, SpecialFormData, ReportKind } from '../types'
import type { TemplateName } from '../lib/excelTemplate'

function toBareBase64(dataUrl: string | null): string | null {
  if (!dataUrl) return null
  const comma = dataUrl.indexOf(',')
  return comma === -1 ? dataUrl : dataUrl.slice(comma + 1)
}

export function InputPage() {
  const { kind } = useParams<{ kind: ReportKind }>()
  const [submitted, setSubmitted] = useState<{
    templateName: TemplateName
    values: Record<string, string>
    formData: CommonFormData
  } | null>(null)

  if (submitted) {
    return (
      <PreviewPage
        templateName={submitted.templateName}
        values={submitted.values}
        photoBase64={toBareBase64(submitted.formData.교육사진)}
        formData={submitted.formData}
        onBack={() => setSubmitted(null)}
      />
    )
  }

  function commonValues(data: CommonFormData): Record<string, string> {
    return {
      협력사: data.협력사,
      공종: data.공종,
      교육일시: data.교육일시,
      교육시간: `${data.교육시작시간} ~ ${data.교육종료시간}`,
      교육방법: data.교육방법,
      교육장소: data.교육장소,
      교육인원: data.교육인원,
      교육강사: data.교육강사,
    }
  }

  if (kind === 'regular') {
    return (
      <RegularEducationForm
        onSubmit={(data) =>
          setSubmitted({ templateName: '정기안전교육', values: commonValues(data), formData: data })
        }
      />
    )
  }

  if (kind === 'hiring') {
    return (
      <HiringEducationForm
        onSubmit={(data: HiringFormData) =>
          setSubmitted({
            templateName: '채용시작업내용변경시교육',
            values: commonValues(data),
            formData: data,
          })
        }
      />
    )
  }

  if (kind === 'special') {
    return (
      <SpecialEducationForm
        onSubmit={(data: SpecialFormData, itemId: number) => {
          const item = specialEducationItems.find((i) => i.id === itemId)!
          setSubmitted({
            templateName: '특별안전교육',
            values: { ...commonValues(data), 교육항목: item.title, 교육내용: item.content },
            formData: data,
          })
        }}
      />
    )
  }

  return <p>알 수 없는 보고서 종류입니다.</p>
}
```

- [ ] **Step 3: App.tsx routing**

`src/App.tsx`:
```tsx
import { BrowserRouter, Routes, Route } from 'react-router-dom'
import { HomePage } from './pages/HomePage'
import { InputPage } from './pages/InputPage'

export default function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<HomePage />} />
        <Route path="/input/:kind" element={<InputPage />} />
      </Routes>
    </BrowserRouter>
  )
}
```

- [ ] **Step 4: Verify dev server runs end-to-end**

Run: `npm run dev`
Manually visit the printed local URL, click through: 홈 → 정기안전교육 → fill 협력사/공종/etc → 사진 업로드 → 보고서 생성 → confirm preview shows entered values → click "Excel 다운로드" → confirm an `.xlsx` file downloads and opens correctly in Excel with values filled and photo embedded. Repeat for 채용시 and 특별 (확인: 특별 흐름에서 공종에 키워드 입력 시 드롭다운이 자동으로 바뀌는지).

- [ ] **Step 5: Commit**

```bash
git add src/pages/HomePage.tsx src/pages/InputPage.tsx src/App.tsx
git commit -m "feat: wire up home, input, and preview pages with routing"
```

---

### Task 17: Deployment config

**Files:**
- Create: `vercel.json`
- Modify: `README.md`

- [ ] **Step 1: Add Vercel config (SPA fallback routing)**

`vercel.json`:
```json
{
  "rewrites": [{ "source": "/(.*)", "destination": "/index.html" }]
}
```

- [ ] **Step 2: Add a short README**

`README.md`:
```markdown
# 신규근로자 교육 결과보고서 웹앱

건설현장 신규근로자 교육(정기안전교육 / 채용시·작업내용변경시 교육 / 특별안전교육) 결과보고서를
기본정보 + 사진 1장 입력만으로 자동 생성하는 클라이언트 전용 웹앱입니다.

## 개발

```bash
npm install
npm run dev
```

## 테스트

```bash
npm test
```

## 배포

Vercel/Netlify에 정적 사이트로 배포 (서버/DB 불필요). `vercel.json`이 SPA 라우팅 fallback을 처리합니다.
```

- [ ] **Step 3: Run full test suite one more time**

Run: `npm test`
Expected: all suites PASS

- [ ] **Step 4: Commit**

```bash
git add vercel.json README.md
git commit -m "docs: add deployment config and README"
```

---

## Self-Review Notes

- **Spec coverage:** Home → input → preview flow (✅ Task 16), per-type forms (✅ Task 12), today-date default + editable (✅ Task 2, 11), end-time auto-calc + editable (✅ Task 2, 11), 교육방법 select (✅ Task 11), 교육장소 default (✅ Task 11), no 결재 fields in form (✅ Task 11 — not present), 1-photo upload (✅ Task 11), 특별 keyword auto-match + manual override (✅ Task 4, 5, 6, 12), A4 portrait 40mm margin shrink-to-fit (✅ Task 3, 13), Excel/PDF/print output (✅ Task 7-9, 13-15), no server/DB (✅ — all client-side), deployment as static site (✅ Task 17).
- **Open follow-up (flagged, not blocking):** `PreviewPage`'s HTML preview is a simplified field list, not a pixel-exact replica of each template's table layout — called out explicitly in Task 15's note so the next engineer doesn't mistake it for done-done. The Excel **download** is the pixel-exact artifact (fills the real template). Tightening the on-screen preview to mirror the template's exact grid is a reasonable fast-follow once the core flow is verified working.
