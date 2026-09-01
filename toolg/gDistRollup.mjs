import rollupFiles from 'w-package-tools/src/rollupFiles.mjs'


let fdSrc = './src'
let fdTar = './dist'


async function rp() {

    await rollupFiles({ //rollupFiles預設會clean folder
        fns: 'WMd2docx.mjs',
        fdSrc,
        fdTar,
        hookNameDist: () => 'w-md2docx',
        // nameDistType: 'kebabCase', //直接由hookNameDist給予
        globals: {
            'path': 'path',
            'fs': 'fs',
            'os': 'os',
            'url': 'url',
            'process': 'process',
            'child_process': 'child_process',
            'sharp': 'sharp', //w-image-proc
            'highlight.js': 'highlight.js', //w-md2html
            'marked': 'marked', //w-md2html
            'marked-katex-extension': 'marked-katex-extension', //w-md2html
            'marked-footnote': 'marked-footnote', //w-md2html
            'marked-highlight': 'marked-highlight', //w-md2html
            'dompurify': 'dompurify', //w-md2html
            'jsdom': 'jsdom', //w-md2html(動態載入)
            '@hapi/hapi': '@hapi/hapi', //ApiServer
        },
        external: [
            'path',
            'fs',
            'os',
            'url',
            'process',
            'child_process',
            'sharp', //w-image-proc
            'highlight.js', //w-md2html
            'marked', //w-md2html
            'marked-katex-extension', //w-md2html
            'marked-footnote', //w-md2html
            'marked-highlight', //w-md2html
            'dompurify', //w-md2html
            'jsdom', //w-md2html(動態載入)
            '@hapi/hapi', //ApiServer
        ],
    })

}
rp()
    .catch((err) => {
        console.log(err)
    })
