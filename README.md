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

> The converter `htmlToDocx.exe` is not part of the npm package of `w-html2docx`: on the first docx conversion `w-html2docx` downloads it, verifies its SHA-256 and keeps it in its cache (`bin-cache/` in the package folder, or the folder given by `W_BIN_CACHE`), independent of the current working directory. A machine without network access needs `W_BIN_MIRROR` / `W_BIN_OFFLINE`, or `WHtml2docx.ensureBins()` called beforehand, see the section 外部執行檔 of the README of `w-html2docx`; a failure there rejects with an error code such as `[E_OFFLINE_MISSING]`. Only the default template of `w-html2docx` is still looked up from the current working directory: run your program from the project root that contains `node_modules/w-html2docx`, or give `fpInTemp`.

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
        fpInTemp: './src/templates/temp.docx', //docx模板, 未給則用w-html2docx內建模板
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
> Give `toc: true` (or a settings object) to `cvMdToDocx`, `cvMdTo` or `rmApiClient.cvMdTo`, or call `addDocxToc` on a docx just converted. Word computes the entries and page numbers, so it also requires `Windows` with `Microsoft Word` and `powershell.exe`, and the conversions and the TOC updates share one queue.
```alias
import WMd2docx from 'w-md2docx/src/WMd2docx.mjs'

let r = await WMd2docx.cvMdToDocx('./test/report.md', './test/report.docx', {
    fpInTemp: './src/templates/temp.docx',
    toc: true, //or settings, e.g. { labels: { fig: 'Figure', tab: 'Table' }, titles: { toc: 'Contents', fig: 'List of Figures', tab: 'List of Tables' } }
})
console.log(r.toc)
// => { skip: '', cover: true, levels: [2, 4], front: [], lists: ['toc'], toc: 26, fig: 0, tab: 0, pagesFirst: '1', pages: 10, warns: [], ms: 2589 }
```

- Headings are paragraphs with an outline level. The tables are inserted before the first heading of the body, and only the headings of the body are listed (`maxLevels`, default 3).
- Captions are paragraphs starting with a label, a number and a space, such as `圖1 名稱` and `表1 名稱` (`labels`, default `圖`/`表`). Their numbers must run from 1 in document order and become `SEQ` fields. A paragraph like `圖1：名稱` is not a caption and is reported in `warns`.
- Pages: the content before the first heading is the cover, without a page number. The front headings (`frontHeadings`, default `摘要`/`ABSTRACT`/`Abstract`) and the tables are numbered I, II, III..., and the body from 1. A document without a cover gets no cover page. Set `pageNumbers: false` to keep the sections and page numbers as they are.
- A Markdown file whose title is the only `#` heading lists the title as the first entry and starts the body from it. Write the title as a non-heading paragraph (e.g. `<div pretitle>`) to keep it on the cover.
- The entries use the font of the first heading. `replaceTocStyles: false` keeps the `toc N` and `table of figures` styles of the template.
- A document without headings is not changed and `toc.skip` tells why. Otherwise any mismatch rejects: `addDocxToc` leaves the docx unchanged, and `cvMdToDocx` deletes the docx it has just converted.
- The service ignores `toc.timeoutMs` of a request, and `rmApiClient.cvMdTo` rejects without writing the docx when the service does not report the TOC result (a service without TOC support).
- Word is driven through COM by Windows PowerShell 5.1 (`powershell.exe`, built into Windows), so the TOC needs `Windows`, `Microsoft Word` and `powershell.exe`. The script is passed on the command line (`-EncodedCommand`) and is never written to a file, because a stand-alone script file may be quarantined by antivirus software as a false positive. The paths of the docx are passed through the environment variables `W_MD2DOCX_SRC` and `W_MD2DOCX_DST`, which are set only while Word is being driven. A machine that blocks PowerShell scripts through Constrained Language Mode (AppLocker, WDAC) can not drive Word this way.
- When an update times out, `powershell.exe` is ended with its child processes, but the `WINWORD` process started through COM is not one of them and may remain, so it has to be ended manually.
- The TOC options `keepLines` and `keepWithObject` (default `true`) apply the same rules as `keepCaption` below to the captions, and `keepLines` also to the listed headings.
- The names given by `titleStyle` (default `報告目錄標題`, the style of the three TOC titles) and `headStyles` (default `章標題`/`節標題`/`小節標題`) appear in the style list of the produced docx, so use neutral names when the docx is delivered to a client. When the template has no style of that name, the package creates one (centered and bold for `titleStyle`, bold for `headStyles`). `temp.docx` (see Templates below) has `報告目錄標題`, so the three TOC titles use its own style, and the three heading styles are created.
- Do not name a style `目錄標題`: Word in Traditional Chinese uses it as the local name of its built-in style `TOC Heading`, so a style of your own with that name is renamed `目錄標題1` (the style of a template is then not found, and one the package creates is renamed too).

#### Keep captions with figures and tables:
> Give `keepCaption: true` (or `{ labels }`) to `cvMdToDocx`, `cvMdTo` or `rmApiClient.cvMdTo` to keep each caption on the same page as its figure or table, with or without a TOC. It only changes paragraph properties after the conversion, so Word is not called again.
```alias
let r = await WMd2docx.cvMdToDocx('./test/report.md', './test/report.docx', {
    keepCaption: true, //or { labels: { fig: 'Figure', tab: 'Table' } }
})
console.log(r.keepCaption)
// => { figs: 0, tabs: 0, keepNext: 0, keepLines: 0, changed: false, warns: [], ms: 4 } (test/report.md has no captions, so the docx is not rewritten)
```

- Captions are recognized as in the TOC, and each caption is kept from splitting across pages (`keepLines`).
- A caption is kept with the table or figure right after it, or with the figure right before it (`keepNext`). A figure is a paragraph holding a picture or an object without text, so a horizontal line (`---`) or text with an inline picture is not a figure. Empty paragraphs between them are crossed and kept too, but a page break or a section break between them (in an empty paragraph, or as a page break before the later one) stops it.
- A caption with a figure or table on both sides (e.g. figures following each other: figure, caption, figure, caption) follows the captions of its kind that have one on one side only: it goes to the side most of them use, and on a tie to the figure before it for a figure caption and to the table or figure after it for a table caption, as figure captions are usually below their figures and table captions above their tables. A figure or table already taken by another caption is left to it, and a caption whose both sides are taken is not kept with either. Two captions with a figure between them and nothing on their other sides are both kept with it.
- A table above a caption is not kept with it (that would take every row), so a caption below its table is only kept from splitting, and is not kept with a table after it either. The paragraphs in or holding text boxes are neither captions nor figures, which is reported in `warns`. A figure made of several paragraphs is kept only by the paragraph next to its caption.
- The rules of the TOC do not apply: captions numbered out of order, several sections, no headings or an existing TOC are processed as usual. With `toc`, captions are kept even when the TOC is skipped or its `keepWithObject` is `false`.
- `changed: false` means nothing was needed and the docx is not rewritten. Otherwise the docx is rewritten (not by Word) through a temporary file, and a failure deletes the docx just converted and rejects.
- `labels` of `keepCaption` and of `toc` are independent. The service passes `keepCaption` on, and `rmApiClient.cvMdTo` rejects without writing the docx when the service does not report the result.

#### Pictures and size limits:
> `cvMdToDocx` (so `cvMdTo` and the service too) does not embed the pictures as base64 into the html it hands to Word. That intermediate html refers to each local picture by its absolute path (`imgConvertToBase64: false` of `w-md2html`, which `optMd2html` can not override), Word reads the files, and `w-html2docx` breaks the links and embeds the pictures into the docx, so the docx does not depend on the original files.

- A report with many pictures is no longer limited by the string length of Node (about 5.4e8 characters) or by the default heap. Measured with the default heap (4288 MB): 720 PNG pictures of 1600×900 that do not compress (0.5 MB each, 393 MB in total) became a docx of 392,757,546 bytes holding all 720 pictures byte for byte, in 417 seconds with `keepCaption: true`; the intermediate html was 169 KB, and the peak of Node was about 1.3 GB (resident set) and that of Word 534 MB. Word needs about 0.6 seconds per picture.
- PNG and JPG are embedded as they are, SVG as vector (svg) with a bitmap preview (png).
- A picture file that does not exist rejects in the md to html stage and lists the absolute paths of all missing pictures (`cvMdTo` and the service list them relative to the assets). A picture that Word can not resolve (for example an unreachable URL) rejects in the html to docx stage with one `unresolved linked picture: <source>` line per picture, and no docx is written. A file that is not a picture is not detected.
- `fpOutHtml` of `cvMdToDocx` is the intermediate file for Word: it refers to the pictures by absolute local paths, so it is not self-contained. For a self-contained html use `w-md2html`, or `cvMdTo` with `out: 'html'` or `'both'` (always base64, whatever `optMd2html.imgConvertToBase64` says; `w-md2html` rejects with a message when the pictures exceed about 96 MiB of PNG).
- The hard limits left reject with a message that tells the cause, before the file is read:
  - `keepCaption` and `toc` read the whole docx into memory: at most 2 GiB - 1 bytes (`fs.readFileSync`).
  - `cvMdTo` and `POST /api/convert` return the docx and the html as base64 in one string (JSON for the service): the files together at most 401,866,734 bytes (the string length limit of V8 less 1 MiB for the rest of the response). Write a bigger docx with `cvMdToDocx`.
  - The service still limits a request (the assets and the template it carries) to `maxMb`.

#### Tall figures and captions:
> With a template that has a line grid (Word: Layout > Page Setup > Document Grid > Specify line grid only, as `temp.docx`), Word rounds the height of each paragraph up to whole grid lines, so a figure as tall as the page pushes its caption to the next page. `cvMdToDocx` (and so `cvMdTo` and the service) therefore passes `imgHeightReserveLines: 4` to `w-html2docx` unless `optHtml2docx` gives it. The height of a figure is limited to leave 4 grid lines below it: an empty paragraph between figure blocks (such as `<div data-for="br"></div>` written in the Markdown), a caption of up to 2 lines of 12pt text and 1 line of margin.
```alias
let r = await WMd2docx.cvMdToDocx('./test/report.md', './test/report.docx', {
    fpInTemp: './src/templates/temp.docx',
    optHtml2docx: {
        imgHeightReserveLines: 6, //default 4, null to disable
    },
    keepCaption: true,
})
```

- The limit lets a figure and its caption fit in one page, and `keepCaption` (or `toc`) keeps them together. With `temp.docx` (38 grid lines of 18pt a page), the tallest figure becomes (38 - 4) × 18 - 0.5 = 611.5pt instead of about 654pt.
- Give a larger number for larger or longer captions: a line of 14pt `標楷體` takes 2 grid lines. Paragraph spacing before or after, and a header or footer taller than the margin, take more room, hence the line of margin.
- `imgRatioHeightMax` of `optHtml2docx` still applies on top of it (0.95 gives about 581pt), so drop it if it was set only to make room for captions.
- Give `null` to keep the default maximum height of `w-html2docx` (0.937 of the content height). The setting has no effect with a template without a line grid, such as the built-in template of `w-html2docx`, see `imgHeightReserveLines` of `w-html2docx` for the exact conditions.

#### Templates:
> `src/templates/temp.docx` is the template that ships with the package: the default template of `rmApiServer` (`templateDef`) and the one the examples use. It is an A4 template with a line grid (38 lines of 18pt a page), and its style names carry no name of an owner or a project, so a docx made from it can be delivered to any client. The styles of a template are listed in the style list of every docx made from it, so when you make your own template, name its styles neutrally too.

| Style name | Style id | Use |
|---|---|---|
| `封面主標題` `封面次標題` `封面委託單位` `封面執行單位` `封面日期` `封面作者` | `Rpt01` … `Rpt06` | cover |
| `摘要主標題` `摘要作者` `摘要標題` | `Rpt07` `Rpt08` `Rpt09` | abstract |
| `摘要與本文` | `Rpt10` | abstract and body text |
| `報告目錄標題` | `Rpt11` | the three TOC titles (default `titleStyle`) |
| `圖本身` `圖標題` | `Rpt12-` `Rpt13-` | figure and its caption |
| `圖表註解(單個)` `圖表註解(多個)` | `Rpt14-1-` `Rpt14-2-` | note below a figure or table (one / several) |
| `表標題` `表頭` `表內容` `表內容純數字` | `Rpt15-` … `Rpt18-` | table caption, header, cells, numeric cells |
| `參考文獻` `數學公式` | `Rpt19` `Rpt20` | references, formulas |
| `項目1` … `項目7` | `Rpt211` `Rpt222` `Rpt233` `Rpt244` `Rpt255` `Rpt266` `Rpt277` | numbered items, level 1 to 7 |

- The ids end with the number the style had in the template that the package shipped before (`01` … `27`), whose style names were the same names with a three-letter prefix of a former owner, the number and the word `報告` in front (for example `<prefix>01報告封面(主標題)`), and the character-linked styles with the suffix ` 字元`. `temp.docx` keeps the page layout, fonts and paragraph formats of that template. Every style name of `temp.docx` is kept by Word when it converts (checked by the tests).
- The properties of the document (author, last modified by, the name of the attached template and its link) are empty, and `temp.docx` has no attached template.

## Breaking changes

Changes after `1.0.18`, together with the dependencies `w-md2html` `^1.1.12` and `w-html2docx` `^1.0.40`:

1. **The intermediate html of `cvMdToDocx` is no longer self-contained.** It always refers to the pictures by absolute local paths instead of embedding base64, and `optMd2html.imgConvertToBase64` no longer has any effect on it. Before, the pictures of the whole report were joined into one html string, which failed with `Invalid string length` at about 400 MB of pictures, and ran out of memory (the process aborted, nothing to catch) at about 80 to 100 MB with the default heap. A program that uses `fpOutHtml` as a self-contained html must call `w-md2html` instead.
2. **The pictures embedded in the docx differ.** Before, every picture was decoded and encoded again as PNG at 96 dpi; now the picture files are embedded as they are: PNG and JPG keep their format and bytes (JPG was a PNG), SVG is embedded as vector (svg) with a bitmap preview (it was a 1067 px PNG). The displayed size is the same, except for a picture that is smaller than the page width and does not carry a resolution of 96 dpi, which Word now displays at 96 dpi (0.75 pt per pixel, as a browser does): a JPG of 72 dpi (100 px: 99.75 pt before, 75.00 pt now) and a PNG without a resolution (`pHYs`) chunk (40 px: 39.75 pt before, 30.00 pt now). A picture wider than the page is scaled to the page width as before, and a PNG with a resolution keeps its size. The Word stage takes about 0.6 seconds per picture instead of about 0.3.
3. **A missing picture always rejects.** Before, a missing picture of the Markdown syntax was only logged and left as a broken picture in the docx, and one of a raw `<img>` line rejected with `fp[…] does not exist`. Now both reject with `N image file(s) do not exist:` followed by the absolute paths; `cvMdTo` and the service turn it into `the referenced file (pics/a.png) does not exist, please make sure it is included in assets` (or `the referenced files (…) do not exist, please make sure they are included in assets`). `allowMissingAssets` of `rmApiClient.cvMdTo` is removed, because the service can not convert a report with a missing picture; the client keeps checking before it sends the assets, and the message no longer ends with `Set allowMissingAssets:true to skip this check`.
4. **A picture that Word can not resolve rejects.** Before, the docx kept a broken 24×24 picture link and the conversion succeeded. Now `cvMdToDocx` rejects with the `unresolved linked picture: <source>` lines of `w-html2docx` and writes no docx.
5. **`cvMdTo` returns a self-contained html.** With `out: 'both'` the html is no longer the intermediate html of the docx (which would point into the work folder that is deleted afterwards): it is produced by `w-md2html` with the pictures embedded as base64 and is identical to `out: 'html'`. With `out: 'html'` or `'both'`, `optMd2html.imgConvertToBase64: false` is ignored. The html is produced before the docx, so too large or missing pictures reject without starting Word.
6. **The diagnosis of the converter changed.** `checkDocxReady()` of `src/utils.mjs` returns a Promise and `getFpExe()` is removed. `exeFound` and `exePath` (also of `GET /api/health`) tell whether `htmlToDocx.exe` is ready in the cache of `w-html2docx` (`ensureBins({ offline: true, mirror: '' })`: the cache only, no download, a hit is verified), no longer whether it lies in `src/` of the current working directory. A failure to obtain the converter rejects with `[E_…]` and the pointer to the README of `w-html2docx` instead of the old hint about downloading from the working directory.
7. **New errors at the hard limits**, see above: a docx over 2 GiB - 1 bytes for `keepCaption` and `toc`, and a result over 401,866,734 bytes for `cvMdTo` and the service (before, `ERR_FS_FILE_TOO_LARGE` or `ERR_STRING_TOO_LONG` without the cause).
8. **The default `titleStyle` of `toc` is now `報告目錄標題`** (a neutral name, like the default `headStyles`). Before, the default was the TOC-title style of the template that the package shipped before, whose name began with the prefix of a former owner, so every docx made with `toc` and no `titleStyle` listed that name in its style list. A template that has no style named `報告目錄標題` (that includes your copy of the earlier template, which has only the old style) now gets a created centered bold 16pt style for the three TOC titles, where the old style was centered, regular, 14pt. To keep the old look, give `titleStyle` the full name of the old style. A template that has `報告目錄標題`, or a `titleStyle` that you give, is not affected. There is no fallback to the old name.
9. **The earlier template file is removed, and the package ships `temp.docx` instead.** `temp.docx` is the default template of `rmApiServer` (`templateDef`, so the service now lists `temp.docx` only) and the one every example uses. It has the same layout (page, grid, fonts and paragraph formats) with neutral style names and ids (see Templates above), and its document properties and attached-template link are empty. A client that asks the service for the earlier template by `templateName`, or relies on its style names, has to change; whoever wants the old look keeps a copy of that template of their own and gives its path as `fpInTemp` (or `templateBase64`) and the name of its title style as `titleStyle`.
