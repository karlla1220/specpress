const { test, describe } = require('node:test')
const assert = require('assert')
const { jsonToDocxTable, jsonFileToDocxTable } = require('../../../../lib/md2docx/handlers/jsonTableHandler')
const path = require('path')

/** Recursively collect all paragraph style IDs from a docx element tree. */
function collectStyles(el) {
  const styles = []
  const json = JSON.stringify(el)
  const re = /w:pStyle.*?w:val.*?"value":"([^"]+)"/g
  let m
  while ((m = re.exec(json)) !== null) styles.push(m[1])
  return styles
}

/** Recursively collect all text content from w:t elements. */
function collectTexts(el) {
  const texts = []
  const json = JSON.stringify(el)
  const re = /"w:t".*?\[.*?"preserve"\},"([^"]+)"/g
  let m
  while ((m = re.exec(json)) !== null) texts.push(m[1])
  return texts
}

describe('jsonToDocxTable', () => {

  test('returns null for empty data', async () => {
    assert.strictEqual(await jsonToDocxTable(null), null)
    assert.strictEqual(await jsonToDocxTable({}), null)
    assert.strictEqual(await jsonToDocxTable({ rows: [] }), null)
  })

  test('generates a Table with correct row count', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['r1'], ['r2'], ['r3']]
    })
    assert.ok(table, 'should return a table')
    // header row + 3 data rows = 4 TableRows
    const rowCount = table.root.filter(r => r.rootKey === 'w:tr').length
    assert.strictEqual(rowCount, 4)
  })

  test('header cells get TAH style', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'Header1' }],
      rows: [['val']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAH'))
  })

  test('left-aligned column gets TAL style', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A', align: 'left' }],
      rows: [['val']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAL'))
  })

  test('center-aligned column gets TAC style', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A', align: 'center' }],
      rows: [['val']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAC'))
  })

  test('right-aligned column gets TAR style', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A', align: 'right' }],
      rows: [['val']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAR'))
  })

  test('NOTE cell gets TAN style', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['NOTE: important']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAN'))
  })

  test('NOTE 2: numbered note gets TAN style', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['NOTE 2: second note']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAN'))
  })

  test('multi-line cell with NOTE gets TAN for note line only', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['Regular text\nNOTE: a note']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAL'), 'regular line should be TAL')
    assert.ok(styles.includes('TAN'), 'note line should be TAN')
  })

  test('NOTE cell has colon-tab replacement', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['NOTE: text here']]
    })
    const json = JSON.stringify(table)
    assert.ok(json.includes('NOTE:\\ttext'), 'should replace ": " with ":\\t"')
  })

  test('colspan with < marker', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }, { name: 'B' }],
      rows: [['spans two', '<']]
    })
    const json = JSON.stringify(table)
    assert.ok(json.includes('"gridSpan"') || json.includes('w:gridSpan'), 'should have colspan')
  })

  test('rowspan with ^ marker', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['spans down'], ['^']]
    })
    const json = JSON.stringify(table)
    assert.ok(json.includes('vMerge') || json.includes('w:vMerge'), 'should have rowspan merge')
  })

  test('skips header row when no column names', async () => {
    const table = await jsonToDocxTable({
      rows: [['a', 'b']]
    })
    const rowCount = table.root.filter(r => r.rootKey === 'w:tr').length
    assert.strictEqual(rowCount, 1, 'only data row, no header')
  })

})
describe('jsonFileToDocxTable', () => {

  test('reads and converts a JSON file', async () => {
    const samplePath = path.join(__dirname, '../../md2html/handlers/samples/jsonTable1.json')
    const result = await jsonFileToDocxTable(samplePath)
    assert.ok(result, 'should return a table or array')
  })

})
describe('italic/bold in JsonTable cells', () => {

  test('NOTE with italic content gets TAN style', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['NOTE: *italic* content']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAN'))
  })

  test('cell with italic text but no NOTE gets normal style', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['*italic* regular text']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAL'))
    assert.ok(!styles.includes('TAN'))
  })

  test('multi-line cell with italic in NOTE line', async () => {
    const table = await jsonToDocxTable({
      columns: [{ name: 'A' }],
      rows: [['Regular *italic*\nNOTE: *italic* note']]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAL'), 'regular line should be TAL')
    assert.ok(styles.includes('TAN'), 'note line should be TAN')
  })

})
describe('key-based rows format', () => {

  test('key-based row objects produce a table', async () => {
    const table = await jsonToDocxTable({
      columns: [{ key: 'a', name: 'Col A' }, { key: 'b', name: 'Col B' }],
      rows: [{ a: 'r1a', b: 'r1b' }]
    })
    assert.ok(table, 'should return a table')
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAH'), 'should have header style')
    assert.ok(styles.includes('TAL'), 'should have body style')
  })

  test('key-based format with alignment', async () => {
    const table = await jsonToDocxTable({
      columns: [{ key: 'x', name: 'X', align: 'center' }],
      rows: [{ x: 'val' }]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAC'), 'should have center alignment style')
  })

  test('key-based format with NOTE gets TAN', async () => {
    const table = await jsonToDocxTable({
      columns: [{ key: 'a', name: 'A' }],
      rows: [{ a: 'NOTE: important' }]
    })
    const styles = collectStyles(table)
    assert.ok(styles.includes('TAN'), 'NOTE should get TAN style')
  })

  test('key-based format with missing key produces empty cell', async () => {
    const table = await jsonToDocxTable({
      columns: [{ key: 'a', name: 'A' }, { key: 'b', name: 'B' }],
      rows: [{ a: 'only a' }]
    })
    assert.ok(table, 'should return a table')
  })


})
describe('jsonToDocxTable: list lines in a cell', () => {

  /** The table's cell paragraphs as [style, text] from the real document.xml. */
  async function cellParagraphs(table) {
    const { Document, Packer } = require('docx')
    const JSZip = require('jszip')
    const buf = await Packer.toBuffer(new Document({ sections: [{ children: [table] }] }))
    const xml = await (await JSZip.loadAsync(buf)).file('word/document.xml').async('string')
    const cells = xml.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []
    return cells.flatMap(tc => (tc.match(/<w:p>[\s\S]*?<\/w:p>|<w:p [\s\S]*?<\/w:p>/g) || []).map(p => [
      (p.match(/w:pStyle w:val="([^"]+)"/) || [])[1],
      (p.match(/<w:t[^>]*>([^<]*)</g) || []).map(t => t.replace(/<w:t[^>]*>|</g, '')).join('')]))
  }

  // 38.331 CellGroupConfig field description: B1/B2 paragraphs with a typed "-<tab>" label in a TAL cell.
  const cell = [
    'The field is mandatory present in the *RRCReconfiguration* message:',
    '',
    '- \\- in the *masterCellGroup:*',
    '',
    '  - \\- at change of AS security key,',
    '',
    '- \\- in the *secondaryCellGroup* at:',
    '',
    'Otherwise, it is optional.'
  ].join('\n')

  test('list lines become B1/B2 paragraphs and blank separator lines no paragraph', async () => {
    const table = await jsonToDocxTable({ columns: [{ name: '' }], rows: [[cell]] })
    const paras = await cellParagraphs(table)
    assert.deepStrictEqual(paras.map(p => p[0]).slice(-5), ['TAL', 'B1', 'B2', 'B1', 'TAL'], JSON.stringify(paras))
  })

  test('the bullet is followed by a tab, as in the body', async () => {
    const table = await jsonToDocxTable({ columns: [{ name: '' }], rows: [['- \\- item']] })
    const paras = await cellParagraphs(table)
    assert.deepStrictEqual(paras[paras.length - 1], ['B1', '-\titem'])
  })

  test('<u> in a cell is an underlined run', async () => {
    const table = await jsonToDocxTable({ columns: [{ name: '' }], rows: [['<u>QPSK</u> plain']] })
    assert.ok(JSON.stringify(table).includes('w:u'), 'expected an underline property')
  })
})

describe('Cell blocks: nested JsonTable link and fenced block', () => {
  const fs = require('fs')
  const os = require('os')
  const { Document, Packer } = require('docx')
  const JSZip = require('jszip')

  /** The table as Word writes it: document.xml's body, properties left out. */
  async function bodyXml(table) {
    const buf = await Packer.toBuffer(new Document({ sections: [{ children: [table] }] }))
    const xml = await (await JSZip.loadAsync(buf)).file('word/document.xml').async('string')
    return xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr'))
      .replace(/<w:(tblPr|tcPr|tblGrid|rPr|pPr)>.*?<\/w:\1>/g, '')
  }
  const nested = xml => (xml.match(/<w:tbl>/g) || []).length - 1

  test('a link line in a cell is a Word table in that cell, its path resolved from baseDir', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'jsontable-'))
    fs.mkdirSync(path.join(dir, 'assets', 'tables'), { recursive: true })
    fs.writeFileSync(path.join(dir, 'assets', 'tables', 'inner.json'), JSON.stringify({ rows: [['inner text']] }))
    const xml = await bodyXml(await jsonToDocxTable({ rows: [['Before\n\n[JsonTable](assets/tables/inner.json)']] }, { baseDir: dir }))
    assert.strictEqual(nested(xml), 1)
    assert.ok(xml.includes('>inner text<'))
    assert.ok(!xml.includes('[JsonTable]'))
    assert.ok(xml.endsWith('</w:tbl><w:p></w:p></w:tc></w:tr></w:tbl>'), 'Word ends the cell with a paragraph after the nested table')
  })

  test('readFile reads the linked table (a git or webview resolver)', async () => {
    const read = p => { assert.strictEqual(p, path.resolve('/spec', 't/n.json')); return JSON.stringify({ rows: [['from resolver']] }) }
    const xml = await bodyXml(await jsonToDocxTable({ rows: [['[JsonTable](t/n.json)']] }, { baseDir: '/spec', readFile: read }))
    assert.ok(xml.includes('>from resolver<'))
  })

  test('a nested table may nest again', async () => {
    const tables = { 'a.json': { rows: [['[JsonTable](b.json)']] }, 'b.json': { rows: [['deep']] } }
    const read = p => JSON.stringify(tables[path.basename(p)])
    const xml = await bodyXml(await jsonToDocxTable({ rows: [['[JsonTable](a.json)']] }, { baseDir: '/spec', readFile: read }))
    assert.strictEqual(nested(xml), 2)
    assert.ok(xml.includes('>deep<'))
  })

  test('a ```jsonTable fence in a cell is a Word table', async () => {
    const xml = await bodyXml(await jsonToDocxTable({ rows: [['```jsonTable\n{"rows": [["R", "F"]]}\n```']] }))
    assert.strictEqual(nested(xml), 1)
    assert.ok(xml.includes('>R<') && xml.includes('>F<'))
    assert.ok(!xml.includes('```'))
  })

  test('a missing linked table is an error paragraph, not a failed export', async () => {
    const xml = await bodyXml(await jsonToDocxTable({ rows: [['[JsonTable](missing.json)']] }, { baseDir: os.tmpdir() }))
    assert.ok(xml.includes('[JsonTable error:'))
  })
})
