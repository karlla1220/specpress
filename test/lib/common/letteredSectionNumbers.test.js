// Clause numbers with letters, as 3GPP specifications use them for inserted
// clauses: 5.5a (TS 38.331), 6.2A.1 (TS 38.101-1), 5.3.5.13b1 (TS 38.331),
// 9.A before 9.1 (TS 38.213).
const { test, describe } = require('node:test')
const assert = require('assert')
const { Md2Html } = require('../../../lib/md2html/md2html')
const { concatenateFiles, extractSectionNumber, parsePlaceholder } = require('../../../lib/common/specProcessor')

const ROOT = 'c:/spec'

function headings(html) {
  return [...html.matchAll(/<h(\d)[^>]*>(.*?)<\/h\1>/g)].map(m => ({ level: Number(m[1]), text: m[2] }))
}

function render(files) {
  const h = new Md2Html({ specRootPath: ROOT })
  const content = concatenateFiles(Object.keys(files).sort(), f => files[f], ROOT)
  return h.renderBody(content, false)
}

describe('path parts with letters', () => {
  test('letter after the digits belongs to the number', () => {
    assert.strictEqual(extractSectionNumber('c:/spec/05 RRC/05a Logged.md', ROOT).sectionNumber, '5.5a')
    assert.strictEqual(extractSectionNumber('c:/spec/06 Tx/02A CA/01 General.md', ROOT).sectionNumber, '6.2A.1')
    assert.strictEqual(extractSectionNumber('c:/spec/05 RRC/13b1 X.md', ROOT).sectionNumber, '5.13b1')
    assert.strictEqual(extractSectionNumber('c:/spec/07A Security/01 General.md', ROOT).sectionNumber, '7A.1')
  })

  test('00 with a letter gives a letter-only part', () => {
    assert.strictEqual(extractSectionNumber('c:/spec/09 UE/00A PUCCH cell switching.md', ROOT).sectionNumber, '9.A')
  })

  test('unchanged for plain numbers and words', () => {
    assert.strictEqual(extractSectionNumber('c:/spec/05_Architecture/01_Overview.md', ROOT).sectionNumber, '5.1')
    assert.strictEqual(extractSectionNumber('c:/spec/05 RRC/00 RRC.md', ROOT).sectionNumber, '5')
    assert.strictEqual(extractSectionNumber('c:/spec/05abc/01 X.md', ROOT).sectionNumber, '5.1')
  })

  test('derived heading name drops the lettered number', () => {
    assert.strictEqual(extractSectionNumber('c:/spec/05 RRC/05a Logged.md', ROOT).sectionHeading, 'Logged')
  })
})

describe('placeholders under lettered paths', () => {
  test('x.x and its sub-clauses take the lettered number', () => {
    const hs = headings(render({ 'c:/spec/05 RRC/05a Logged.md': '## x.x Logged\n\n### x.x.1 Sub\n\n#### x.x.1.1 Deeper' }))
    assert.deepStrictEqual(hs.map(h => h.text), ['5 RRC', '5.5a Logged', '5.5a.1 Sub', '5.5a.1.1 Deeper'])
  })

  test('9.A sorts before 9.1', () => {
    const hs = headings(render({
      'c:/spec/09 UE/00 UE.md': '# x UE procedures',
      'c:/spec/09 UE/00A PUCCH cell switching.md': '## x.x PUCCH cell switching',
      'c:/spec/09 UE/01 General.md': '## x.x General',
    }))
    assert.deepStrictEqual(hs.map(h => h.text), ['9 UE procedures', '9.A PUCCH cell switching', '9.1 General'])
  })

  test('x.xa keeps working in a plain file and in a lettered file', () => {
    assert.strictEqual(headings(render({ 'c:/spec/05 RRC/05 Logged.md': '## x.xa Logged a' }))[1].text, '5.5a Logged a')
    assert.strictEqual(headings(render({ 'c:/spec/05 RRC/05a Logged.md': '## x.xa Logged a' }))[1].text, '5.5a Logged a')
  })

  test('x.xb in a file named 05a is an error', () => {
    assert.ok(headings(render({ 'c:/spec/05 RRC/05a Logged.md': '## x.xb Other' }))[1].text.includes('E.R.R.O.R'))
  })

  test('caption below x.x.1a takes the lettered number', () => {
    const h = new Md2Html({ specRootPath: ROOT })
    const html = h.renderBody('## x.x Gen\n\n### x.x.1a Added\n\nTable x.x.x-1: Values\n\n| a |\n|---|\n| 1 |', false, null, 'c:/spec/05 RRC/03 Gen.md')
    assert.ok(html.includes('Table 5.3.1a-1: Values'), html)
  })

  test('folder headings carry the letter', () => {
    const hs = headings(render({ 'c:/spec/06 Tx/02A CA/01 General.md': '### x.x.x General' }))
    assert.deepStrictEqual(hs.map(h => h.text), ['6 Tx', '6.2A CA', '6.2A.1 General'])
  })
})

describe('written section numbers', () => {
  test('kept when they are the derived number or below it', () => {
    const hs = headings(render({
      'c:/spec/05 RRC/05a Logged.md': '## 5.5a Logged\n\n### 5.5a.1 Sub\n\n#### 5.5a.1.1 Deeper',
      'c:/spec/09 UE/00 UE.md': '# 9 UE procedures\n\n## 9.A PUCCH cell switching',
      'c:/spec/05 RRC/03 Conn.md': '## 5.3 Conn\n\n#### 5.3.5.13b1 Odd',
    }))
    assert.ok(hs.every(h => !h.text.includes('E.R.R.O.R')), JSON.stringify(hs))
    assert.ok(hs.some(h => h.text === '9.A PUCCH cell switching'))
  })

  test('caption placeholder below a written number', () => {
    const h = new Md2Html({ specRootPath: ROOT })
    const html = h.renderBody('## 5.5a Logged\n\n### 5.5a.1 Sub\n\nTable x.x.x-1: Values\n\n| a |\n|---|\n| 1 |', false, null, 'c:/spec/05 RRC/05a Logged.md')
    assert.ok(html.includes('Table 5.5a.1-1: Values'), html)
  })

  test('E.R.R.O.R when the number disagrees with the path', () => {
    const hs = headings(render({ 'c:/spec/08 Msg/02 Defs.md': '## 8.2 Defs\n\n#### 8.6.27.1 Message definition' }))
    assert.ok(hs[2].text.includes('E.R.R.O.R'), JSON.stringify(hs))
  })

  test('E.R.R.O.R when the heading level does not match the number', () => {
    const hs = headings(render({ 'c:/spec/05 RRC/01 Gen.md': '### 5.1 Gen' }))
    assert.ok(hs[1].text.includes('E.R.R.O.R'), JSON.stringify(hs))
  })

  test('E.R.R.O.R for a second heading at the derived level', () => {
    const hs = headings(render({ 'c:/spec/05 RRC/01 Gen.md': '## 5.1 Gen\n\n## 5.1 Again' }))
    assert.ok(hs[2].text.includes('E.R.R.O.R'), JSON.stringify(hs))
  })
})

describe('headings that only look like placeholders', () => {
  test('xApp and x2 are words', () => {
    assert.strictEqual(parsePlaceholder('xApp procedures'), null)
    assert.strictEqual(parsePlaceholder('x2 interface'), null)
    const hs = headings(render({ 'c:/spec/05 RRC/01 Gen.md': '## x.x Gen\n\n### xApp procedures' }))
    assert.strictEqual(hs[2].text, 'xApp procedures')
  })
})
