# w-md2docx
A tool for Markdown to Docx.

![language](https://img.shields.io/badge/language-JavaScript-orange.svg) 
[![npm version](http://img.shields.io/npm/v/w-md2docx.svg?style=flat)](https://npmjs.org/package/w-md2docx) 
[![license](https://img.shields.io/npm/l/w-md2docx.svg?style=flat)](https://npmjs.org/package/w-md2docx) 
[![npm download](https://img.shields.io/npm/dt/w-md2docx.svg)](https://npmjs.org/package/w-md2docx) 
[![npm download](https://img.shields.io/npm/dm/w-md2docx.svg)](https://npmjs.org/package/w-md2docx) 
[![jsdelivr download](https://img.shields.io/jsdelivr/npm/hm/w-md2docx.svg)](https://www.jsdelivr.com/package/npm/w-md2docx)

## Documentation
To view documentation or get support, visit [docs](https://yuda-lyu.github.io/w-md2docx/global.html).

## Core
> `w-md2docx` is based on `w-md2html` (Markdown to Html) and `w-html2docx` (Html to Docx via `win32com` of `Microsoft Word`), so the docx conversion only runs in `Windows` with `Microsoft Word` installed.

> The converter `htmlToDocx.exe` is located and, when absent (e.g. npm blocked the `postinstall` script of `w-html2docx`), downloaded automatically by `w-html2docx` on the first docx conversion. It is resolved from the current working directory, so run your program from the project root that contains `node_modules/w-html2docx`, with network access for the first conversion.

It provides four entries:
- `cvMdToDocx`: convert a Markdown file to a Docx file.
- `cvMdTo`: convert Markdown content (with attached assets) to Html/Docx content in base64.
- `rmApiServer`: a hapi service (`GET /api/health`, `GET /api/selftest`, `POST /api/convert`) that wraps `cvMdTo`, so machines without Word can convert through it.
- `rmApiClient`: a client for `rmApiServer`, reading a local Markdown file with its images and writing back the converted files.

## Installation

### Using npm(ES6 module):
```alias
npm i w-md2docx
```

#### Example:
> **Link:** [[dev source code](https://github.com/yuda-lyu/w-md2docx/blob/master/g.mjs)]
```alias
import w from 'wsemi'
import WMd2docx from './src/WMd2docx.mjs'
//import WMd2docx from 'w-md2docx/src/WMd2docx.mjs'
//import WMd2docx from 'w-md2docx'


async function test() {

    let fpIn = `./test/report.md`
    let fpOut = `./test/report.docx`
    let opt = {
        fpInTemp: './src/templates/temp_tpc.docx', //docx模板, 未給則用w-html2docx內建模板
        optMd2html: {
            imgWidthMax: '500px',
        },
        optHtml2docx: {
            imgRatioWidthMax: 0.5,
        },
    }

    let r = await WMd2docx.cvMdToDocx(fpIn, fpOut, opt)
    console.log(r)
    // => { fpOutDocx: '...\\test\\report.docx', sizeDocx: 40960, sizeHtml: 61785, ms: 6532 }

    w.fsDeleteFile(fpOut)

}
test()
    .catch((err) => {
        console.log('catch', err)
    })
```

#### API service and client:
```alias
import WMd2docx from 'w-md2docx/src/WMd2docx.mjs'

//server (Windows with Microsoft Word)
let { server } = await WMd2docx.rmApiServer({ port: 22000, token: '' })
console.log(`listening on ${server.info.uri}`)

//client (any machine)
let r = await WMd2docx.rmApiClient.cvMdTo({
    host: '127.0.0.1',
    port: 22000,
    fpInMd: './test/report.md',
    fpOutDocx: './test/report.docx',
})
console.log(r)
// => { ms: 6532, nAssets: 1, template: 'default', docx: { fp: '...', size: 40960 } }
```
