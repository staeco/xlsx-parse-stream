/* eslint-disable no-loops/no-loops */
const Excel = require('exceljs')
const duplex = require('duplexify')
const { Readable, PassThrough, finished } = require('stream')

const matchSelector = (selector, worksheet) =>
  selector.includes('*') || selector.includes(worksheet.name)

const handleError = (err, isEnded) => {
  if (!err) return
  if (isEnded && err.message === 'FILE_ENDED') return
  if (err.message && err.message.indexOf('invalid signature') !== -1) {
    err = new Error('Legacy XLS files are not supported, use an XLSX file instead!')
  }
  throw err
}

module.exports = ({ mapHeaders, mapValues, selector = '*' } = {}) => {
  if (selector && !Array.isArray(selector)) selector = [ selector ]
  let isEnded = false
  const input = new PassThrough()
  const reader = new Excel.stream.xlsx.WorkbookReader(input)
  const createReader = async function* () {
    try {
      for await (const worksheet of reader) {
        if (!matchSelector(selector, worksheet)) continue

        let headers
        for await (const row of worksheet) {
          if (row.values.length === 0) continue // blank
          if (!headers) {
            headers = mapHeaders ? row.values.map(mapHeaders) : row.values
            out.emit('header', headers)
            continue
          }
          const item = row.values.reduce((acc, v, idx) => {
            acc[headers[idx]] = mapValues ? mapValues(v) : v
            return acc
          }, {})
          yield item
        }
      }
    } catch (err) {
      handleError(err, isEnded)
    }
  }

  const out = Readable.from(createReader())
  const final = duplex.obj(input, out)
  finished(input, () => isEnded = true)
  finished(out, (err) => {
    isEnded = true
    if (err) out.emit('error', err)
  })
  return final
}

module.exports.getSelectors = () => {
  let isEnded = false
  const input = new PassThrough()
  const reader = new Excel.stream.xlsx.WorkbookReader(input)
  const createReader = async function* () {
    yield '*'
    try {
      for await (const worksheet of reader) {
        yield worksheet.name
      }
    } catch (err) {
      handleError(err, isEnded)
    }
  }

  const out = Readable.from(createReader())
  const final = duplex.obj(input, out)
  finished(out, () => isEnded = true)
  finished(input, () => isEnded = true)
  return final
}
