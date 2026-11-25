import { BIN_NAME, getBin } from './binEnsure.mjs'


//binEnsureChild: 供 test/unit-binEnsure.test.mjs 起之子行程, 經 test/tools/binEnsure.mjs 之 getBin(即 before 實際使用之綁定)取得 htmlToDocx.exe
//  快取夾與鏡像由環境變數 T_CACHE、T_MIRROR 指定; 每行輸出一個 JSON: 開始等候他方之鎖時輸出 { ev: 'wait' }, 結束時輸出 { result: { fp, steps, err } };
//  不使用 W_BIN_* 環境變數(由父行程清除)
async function main() {
    //out: 寫出一行 JSON, resolve 於寫入完成(管線輸出於 Windows 為非同步, 結束前須等寫完, 否則 process.exit 會截斷)
    let out = (o) => {
        return new Promise((resolve) => {
            process.stdout.write(`${JSON.stringify(o)}\n`, resolve)
        })
    }
    let steps = []
    let fp = null
    let err = null
    try {
        let r = await getBin({
            cacheDir: process.env.T_CACHE,
            mirror: process.env.T_MIRROR,
            offline: false,
            onProgress: (v) => {
                steps.push(v.step === 'mirror' ? `mirror:${v.result}` : v.step === 'cache' ? `cache:${v.hit}` : v.step)
                if (v.step === 'wait') {
                    out({ ev: 'wait' })
                }
            },
        })
        fp = r[BIN_NAME]
    }
    catch (e) {
        err = { code: e.code, message: String(e.message) }
    }
    await out({ result: { fp, steps, err } })
}


main().then(() => {
    process.exit(0)
})
