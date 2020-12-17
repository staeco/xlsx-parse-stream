# xlsx-parse-stream [![Build Status](https://travis-ci.org/staeco/xlsx-parse-stream.svg?branch=master)](https://travis-ci.org/staeco/xlsx-parse-stream)

> Parse excel (XLSX) files as a through stream to JSON using exceljs

## Install

```shell
$ npm install xlsx-parse-stream
```
## Usage

```js
const excel = require('xlsx-parse-stream')
const request = require('superagent')
const through = require('through2')

// from a URL
const req = request.get('http://localhost:8000/file.xlsx')
  .pipe(excel())
  .pipe(through2.obj((row, _, cb) => {
    // row = the parsed object!
    cb()
  }))


// from the FS
fs.createReadStream(__dirname + '/file.xlsx')
  .pipe(excel())
  .pipe(through2.obj((row, _, cb) => {
    // row = the parsed object!
    cb()
  }))
```


### Options

##### selector

String or array of strings specifying the sheet names you want to parse. You can also specify `"*"` to pull from all sheets (this is the default).

When pulling from multiple sheets, the first row of each sheet will be treated as the header.

```js
// loading a specific sheet
fs.createReadStream(__dirname + '/file.xlsx')
  .pipe(excel({ selector: 'Sheet1' }))
  .pipe(through2.obj((row, _, cb) => {
    // row = the parsed object!
    cb()
  }))

  // loading multiple specific sheets
  fs.createReadStream(__dirname + '/file.xlsx')
    .pipe(excel({ selector: [ 'Sheet1', 'Sheet3' ] }))
    .pipe(through2.obj((row, _, cb) => {
      // row = the parsed object!
      cb()
    }))
```

## [License](LICENSE) (MIT)
