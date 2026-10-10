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

It provides five entries:
- `cvMdToDocx`: convert a Markdown file to a Docx file.
- `cvMdTo`: convert Markdown content (with attached assets) to Html/Docx content in base64.
- `addDocxToc`: add a table of contents, a table of figures and a table of tables to a docx, and renumber its pages.
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

#### Tables of contents:
> Give `toc: true` (or a settings object) to `cvMdToDocx`, `cvMdTo` or `rmApiClient.cvMdTo`, or call `addDocxToc` on a docx just converted. Word computes the entries and page numbers, so it also requires `Windows` with `Microsoft Word`, and the conversions and the TOC updates share one queue.
```alias
import WMd2docx from 'w-md2docx/src/WMd2docx.mjs'

let r = await WMd2docx.cvMdToDocx('./test/report.md', './test/report.docx', {
    fpInTemp: './src/templates/temp_tpc.docx',
    toc: true, //or settings, e.g. { labels: { fig: 'Figure', tab: 'Table' }, titles: { toc: 'Contents', fig: 'List of Figures', tab: 'List of Tables' } }
})
console.log(r.toc)
// => { skip: '', cover: true, levels: [1, 3], front: [], lists: ['toc', 'tab'], toc: 18, fig: 0, tab: 3, pagesFirst: '1', pages: 21, warns: [], ms: 3120 }
```

- Headings are paragraphs with an outline level. The tables are inserted before the first heading of the body, and only the headings of the body are listed (`maxLevels`, default 3).
- Captions are paragraphs starting with a label, a number and a space, such as `圖1 名稱` and `表1 名稱` (`labels`, default `圖`/`表`). Their numbers must run from 1 in document order and become `SEQ` fields. A paragraph like `圖1：名稱` is not a caption and is reported in `warns`.
- Pages: the content before the first heading is the cover, without a page number. The front headings (`frontHeadings`, default `摘要`/`ABSTRACT`/`Abstract`) and the tables are numbered I, II, III..., and the body from 1. A document without a cover gets no cover page. Set `pageNumbers: false` to keep the sections and page numbers as they are.
- A Markdown file whose title is the only `#` heading lists the title as the first entry and starts the body from it. Write the title as a non-heading paragraph (e.g. `<div pretitle>`) to keep it on the cover.
- The entries use the font of the first heading. `replaceTocStyles: false` keeps the `toc N` and `table of figures` styles of the template.
- A document without headings is not changed and `toc.skip` tells why. Otherwise any mismatch rejects: `addDocxToc` leaves the docx unchanged, and `cvMdToDocx` deletes the docx it has just converted.
- The service ignores `toc.timeoutMs` of a request, and `rmApiClient.cvMdTo` rejects without writing the docx when the service does not report the TOC result (a service without TOC support).
- Word is driven by `cscript` with `updateDocxToc.vbs`. Microsoft has deprecated VBScript, which becomes a feature on demand before its removal from Windows. When an update times out, the `WINWORD` process started through COM may remain and has to be ended manually.
- The TOC options `keepLines` and `keepWithObject` (default `true`) apply the same rules as `keepCaption` below to the captions, and `keepLines` also to the listed headings.

#### Keep captions with figures and tables:
> Give `keepCaption: true` (or `{ labels }`) to `cvMdToDocx`, `cvMdTo` or `rmApiClient.cvMdTo` to keep each caption on the same page as its figure or table, with or without a TOC. It only changes paragraph properties after the conversion, so Word is not called again.
```alias
let r = await WMd2docx.cvMdToDocx('./test/report.md', './test/report.docx', {
    keepCaption: true, //or { labels: { fig: 'Figure', tab: 'Table' } }
})
console.log(r.keepCaption)
// => { figs: 3, tabs: 2, keepNext: 5, keepLines: 5, changed: true, warns: [], ms: 210 }
```

- Captions are recognized as in the TOC, and each caption is kept from splitting across pages (`keepLines`).
- A caption followed by a table or a figure is kept with it, otherwise a figure right before a caption is kept with the caption (`keepNext`). Empty paragraphs between them are crossed and kept too, but a page break, a section break or a page break before stops it. When figures follow each other directly (figure, caption, figure, caption), the first caption is kept with the next figure, so they are chained.
- Table captions below their tables are not handled. The paragraphs in or holding text boxes are neither captions nor figures, which is reported in `warns`.
- The rules of the TOC do not apply: captions numbered out of order, several sections, no headings or an existing TOC are processed as usual. With `toc`, captions are kept even when the TOC is skipped or its `keepWithObject` is `false`.
- `changed: false` means nothing was needed and the docx is not rewritten. Otherwise the docx is rewritten (not by Word) through a temporary file, and a failure deletes the docx just converted and rejects.
- `labels` of `keepCaption` and of `toc` are independent. The service passes `keepCaption` on, and `rmApiClient.cvMdTo` rejects without writing the docx when the service does not report the result.
