const { readFileSync } = require('fs')
const { resolve } = require('path')
const { Table, TableRow, TableCell, Paragraph, WidthType, BorderStyle } = require('docx')
const { buildSpanMap, normalizeJsonTable } = require('../../common/buildSpanMap')
const { parseBullet } = require('../../common/bullets')

const TABLE_LINK = /^\[JsonTable\]\(([^)\s]+\.json)\)$/
const JSON_FENCE = /^ *```jsonTable\s*$/

const cellBorders = {
  top: { style: BorderStyle.SINGLE, size: 1 },
  bottom: { style: BorderStyle.SINGLE, size: 1 },
  left: { style: BorderStyle.SINGLE, size: 1 },
  right: { style: BorderStyle.SINGLE, size: 1 }
}


function alignStyle(col, isHeader) {
  if (isHeader) return 'TAH'
  const align = col && col.align
  if (align === 'center') return 'TAC'
  if (align === 'right') return 'TAR'
  return 'TAL'
}

/**
 * Converts parsed JSON table data to a docx Table.
 *
 * @param {object} data - JSON object with `columns` and `rows` arrays.
 * @param {object} [files] - Where a cell's nested tables are read from: `baseDir` resolves a
 *   `[JsonTable](…)` link line, `readFile(path)` reads it (default: the filesystem).
 * @returns {Promise<Table>} docx Table element.
 */
async function jsonToDocxTable(data, files = {}) {
  data = normalizeJsonTable(data)
  if (!data || !data.rows || data.rows.length === 0) return null
  const { buildInlineRuns } = require('../md2docx')

  const columns = data.columns || []
  const rowCount = data.rows.length
  const colCount = Math.max(columns.length, ...data.rows.map(r => r.length))
  const spanMap = buildSpanMap(data.rows, rowCount, colCount)
  const tableRows = []

  const MarkdownIt = require('markdown-it')
  const cellMd = new MarkdownIt({ html: true })

  async function cellRuns(text, isNote) {
    const tokens = cellMd.parseInline(text, {})
    const children = tokens.length > 0 ? tokens[0].children : []
    return buildInlineRuns(children || [], isNote ? { replaceFirstColonSpace: true } : {})
  }

  // A list line in a cell ("- \\- text", nested two spaces per level) is a B1..B9 paragraph with
  // "bullet<tab>text", as list items are in the body (exportBulletItem) and as Word writes them in cells.
  const LIST_LINE = /^( *)[-*+] (.*)$/
  async function cellListParagraph(line) {
    const m = LIST_LINE.exec(line)
    if (!m) return null
    const level = Math.min(Math.floor(m[1].length / 2) + 1, 9)
    const tokens = cellMd.parseInline(m[2], {})
    const children = (tokens.length > 0 ? tokens[0].children : []) || []
    const first = children[0]
    if (first && first.type === 'text') {
      const { bullet, rest } = parseBullet(first.content)
      first.content = bullet + '\t' + rest
    } else {
      children.unshift({ type: 'text', content: '-\t' })
    }
    return new Paragraph({ children: await buildInlineRuns(children), style: `B${level}` })
  }

  const hasHeaders = columns.some(col => col.name)
  if (hasHeaders) {
    const cells = []
    for (let c = 0; c < colCount; c++) {
      const runs = await cellRuns(String(columns[c]?.name ?? ''))
      cells.push(new TableCell({
        children: [new Paragraph({ children: runs, style: 'TAH' })],
        borders: cellBorders
      }))
    }
    tableRows.push(new TableRow({ children: cells }))
  }

  const noteRe = /^(\*\*)?NOTE(\*\*)?( \d+)?:/

  // A cell's nested table: a line that is only a JsonTable link, or a ```jsonTable fence, as in the body.
  async function nestedTables(json) {
    const items = [].concat(JSON.parse(json))
    const tables = []
    for (const item of items) {
      const t = await jsonToDocxTable(item, files)
      if (t) tables.push(t)
    }
    return tables
  }
  async function linkedTables(href) {
    const filePath = resolve(files.baseDir || '.', href)
    const read = files.readFile || (p => readFileSync(p, 'utf8'))
    return nestedTables(read(filePath))
  }

  for (let r = 0; r < rowCount; r++) {
    const cells = []
    for (let c = 0; c < colCount; c++) {
      const span = spanMap[r][c]
      if (span.skip) continue
      const text = span.value === null || span.value === undefined ? '' : String(span.value)
      const baseStyle = alignStyle(columns[c], false)
      const lines = text.split('\n')
      const paragraphs = []
      for (let k = 0; k < lines.length; k++) {
        const line = lines[k]
        const link = TABLE_LINK.exec(line.trim())
        if (link) {
          try {
            paragraphs.push(...await linkedTables(link[1]))
          } catch (e) {
            paragraphs.push(new Paragraph({ text: `[JsonTable error: ${e.message}]`, style: baseStyle }))
          }
          continue
        }
        if (JSON_FENCE.test(line)) {
          let end = k + 1
          while (end < lines.length && lines[end].trim() !== '```') end++
          try {
            paragraphs.push(...await nestedTables(lines.slice(k + 1, end).join('\n')))
          } catch (e) {
            paragraphs.push(new Paragraph({ text: `[JsonTable error: ${e.message}]`, style: baseStyle }))
          }
          k = end
          continue
        }
        // A blank line next to a list line only separates Markdown blocks; Word has no empty paragraph there.
        if (line.trim() === '' && (LIST_LINE.test(lines[k - 1] || '') || LIST_LINE.test(lines[k + 1] || ''))) continue
        const item = await cellListParagraph(line)
        if (item) { paragraphs.push(item); continue }
        const isNote = noteRe.test(line.trim())
        const runs = await cellRuns(line, isNote)
        paragraphs.push(new Paragraph({ children: runs, style: isNote ? 'TAN' : baseStyle }))
      }
      // Word ends every cell with a paragraph, also after a nested table.
      if (paragraphs.length === 0 || paragraphs[paragraphs.length - 1] instanceof Table) {
        paragraphs.push(new Paragraph({ style: baseStyle }))
      }
      const cellOpts = { children: paragraphs, borders: cellBorders }
      if (span.colspan > 1) cellOpts.columnSpan = span.colspan
      if (span.rowspan > 1) cellOpts.rowSpan = span.rowspan
      cells.push(new TableCell(cellOpts))
    }
    tableRows.push(new TableRow({ children: cells }))
  }

  return new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } })
}

/**
 * Reads a JSON file and converts it to a docx Table.
 *
 * @param {string} filePath - Absolute path to the JSON file.
 * @param {object} [files] - As for jsonToDocxTable: `baseDir` is the linking Markdown file's directory.
 * @returns {Promise<Table|null>} docx Table element, or null if data is empty.
 */
async function jsonFileToDocxTable(filePath, files = {}) {
  const content = readFileSync(filePath, 'utf8')
  const data = JSON.parse(content)
  if (Array.isArray(data)) {
    const tables = []
    for (const item of data) {
      const t = await jsonToDocxTable(item, files)
      if (t) tables.push(t)
    }
    return tables
  }
  return jsonToDocxTable(data, files)
}

module.exports = { jsonToDocxTable, jsonFileToDocxTable }
