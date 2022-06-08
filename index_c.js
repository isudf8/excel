const path = require('./path.json')
const readline = require('readline');
const fs = require('fs');
const xl = require('xlsx');
let excelDir = "./excel"           // excel配置表目录
let outDir = "./out"            // 导出目录
let outJsonFolder = "excel"     // 导出的json目录

// 前后端解析索引的行
let hang = 3 // 前端3 后端4
// 表格列表 example: ["iconopen","mainTask"]
let excelFiles = []
// 表格与表单的数据类型 example:{iconopen:{id:string}}
let excelAndSheets = {}

// // 创建文件夹
// fs.mkdir(outDir, { recursive: true }, () => {
// //     fs.mkdir(outDir + "/" + outJsonFolder, { recursive: true }, () => { })
// })

function exeExcelConfig() {
    let str = ""
    if (hang == 4) {
        // 引入json文件
        excelFiles.forEach(element => {
            str += `const ${element} = require('../../../app/excel/${element}.json')\n`
        });
        str += "\n"
    }

    // 添加类型定义
    excelFiles.forEach(excelName => {
        str += `// type for ${excelName}.excel\n`
        for (const sheetName in excelAndSheets[excelName]) {
            str += `export type ${toUpperCaseByStr(excelName)}${toUpperCaseByStr(sheetName)} = {`

            //遍历索引存入类型定义
            let keyObjList = excelAndSheets[excelName][sheetName]
            for (const key in keyObjList) {
                str += `${key.split(",")[0]}:${keyObjList[key].value},`
            }
            // 去掉最后一个逗号 ","
            if (str.charAt(str.length - 1) == ",") {
                str = str.substring(0, str.length - 1)
            }

            str += "}"
            str += "\n"
        }

    });
    str += "\n"
    // 添加解析json表代码
    str +=
        `export default class ConfProxy<T> {
    static getKey(args: string[]): string {
        let out = ""
        args.forEach(element => {
            out += element
            out += "_"
        });
        return out
    }
    static setKey(target: any, args: string[]): string {
        let out = ""
        args.forEach(element => {
            out += target[element]
            out += "_"
        });
        return out
    }
    pool: { [key: string]: T } = {}
    constructor(conf: T[], ...args: string[]) {
        conf.forEach(element => {
            this.pool[ConfProxy.setKey(element, args)] = element
        });
    }
    getItem(...args: string[]): T | null {
        return this.pool[ConfProxy.getKey(args)]
    }
}
export class ConfListProxy<T> {
    pool: { [key: string]: T[] } = {}
    constructor(conf: T[], ...args: string[]) {
        conf.forEach(element => {
            let key = ConfProxy.setKey(element, args)
            if (this.pool[key] == null) {
                this.pool[key] = []
            }
            this.pool[key].push(element)
        });
    }
    getItemList(...args: string[]): T[] | null {
        return this.pool[ConfProxy.getKey(args)]
    }
}`

    if (hang == 3) {
        str += "\n"
        str += `
function load_jsons(jsonFiles: string[]) {
    return new Promise((resolve, reject) => {
        cc.resources.load(jsonFiles, (err, data: cc.JsonAsset[]) => {
            if (err) {
                reject(err);
            } else {
                let jsonArray: { [key: string]: cc.JsonAsset } = {}
                data.forEach(element => {
                    jsonArray[element.name] = element.json
                });
                resolve(jsonArray);
            }
        });
    });
}`
    }
    str += "\n"
    if (hang == 4) {
        // 添加数据类型
        for (const excelName in excelAndSheets) {
            str += `// array for excel: ${excelName}\n`
            for (const sheetName in excelAndSheets[excelName]) {
                if (hang == 3) {
                    str += `let ${excelName}`
                } else if (hang == 4) {
                    str += `export let ${excelName}`
                }
                str += `${toUpperCaseByStr(sheetName)} = <${toUpperCaseByStr(excelName)}${toUpperCaseByStr(sheetName)}[]>${excelName}.${sheetName}\n`
            }
        }
    }

    if (hang == 3) {
        str += `export class ExcelCfg {\n    //定义类型【复制】\n`
        for (const excelName in excelAndSheets) {
            for (const sheetName in excelAndSheets[excelName]) {
                str += `    static ${excelName}${toUpperCaseByStr(sheetName)}Json: ${toUpperCaseByStr(excelName)}${toUpperCaseByStr(sheetName)}[]\n`
            }
        }
        str += '    // 读取数据字段【复制】\n'
        // // 根据表中配置的*等信息，生成读取数据字段
        // for (const excelName in excelAndSheets) {
        //     for (const sheetName in excelAndSheets[excelName]) {
        //         for (const key in excelAndSheets[excelName][sheetName]) {
        //             let _keyArray = (excelAndSheets[excelName][sheetName][key].keys).split(",")
        //             // 字段定义的名称
        //             let title = ''
        //             _keyArray.forEach(element => {
        //                 title += toUpperCaseByStr(element)
        //             });
        //             if (excelAndSheets[excelName][sheetName][key].isKey == 1) {

        //                 str += `    static ${excelName}${toUpperCaseByStr(sheetName)}${title}: ConfProxy<${toUpperCaseByStr(excelName)}${toUpperCaseByStr(sheetName)}>\n`
        //             } else if (excelAndSheets[excelName][sheetName][key].isKey == 2) {
        //                 str += `    static ${excelName}${toUpperCaseByStr(sheetName)}${title}List: ConfListProxy<${toUpperCaseByStr(excelName)}${toUpperCaseByStr(sheetName)}>\n`
        //             }
        //         }
        //     }
        // }
        str += `    static async initJson(callback: Function) {\n`
        let jsonUrls = []
        for (const excelName in excelAndSheets) {
            let ex = ""
            if (jsonUrls.length % 3 == 0 && jsonUrls.length != 0) {
                ex = `\n        `
            }
            jsonUrls.push(`${ex}'./${outJsonFolder}/${excelName}'`)

        }
        str += `        let jsonUrls: string[] = [${jsonUrls}]\n`
        str += `        let jsonObjs = await load_jsons(jsonUrls)\n`
        // 添加初始化json为可读字段的逻辑
        for (const excelName in excelAndSheets) {
            str += `        let ${excelName} = jsonObjs['${excelName}']\n`
            for (const sheetName in excelAndSheets[excelName]) {
                str += `        ExcelCfg.${excelName}${toUpperCaseByStr(sheetName)}Json = ${excelName}['${sheetName}']\n`
                for (const key in excelAndSheets[excelName][sheetName]) {
                    let _keyArray = (excelAndSheets[excelName][sheetName][key].keys).split(",")
                    // 字段定义的名称
                    let title1 = ''
                    let title2 = ''
                    _keyArray.forEach((element, index) => {
                        title1 += toUpperCaseByStr(element)
                        title2 += `"${element}"`
                        if (index != _keyArray.length - 1) {
                            title2 += ", "
                        }
                    });

                    // 取用方式由前端自己写
                    // if (excelAndSheets[excelName][sheetName][key].isKey == 1) {
                    //     str += `        GameCfg.${excelName}${toUpperCaseByStr(sheetName)}${title1} = new ConfProxy(GameCfg.${excelName}${toUpperCaseByStr(sheetName)}Json,${title2})\n`
                    // } else if (excelAndSheets[excelName][sheetName][key].isKey == 2) {
                    //     str += `        GameCfg.${excelName}${toUpperCaseByStr(sheetName)}${title1}List = new ConfListProxy(GameCfg.${excelName}${toUpperCaseByStr(sheetName)}Json,${title2})\n`
                    // }
                }
            }
        }

        str += "        // 这里配置需要读取表的字段\n\n"
        str += `        callback()\n`
        str += `    }`
        str += `\n}`

        // for
        // console.log("===表:", excelName, excelAndSheets[excelName])
    }
    // if (hang == 3) {
    //     // 添加有类型定义的配置表
    //     str += "\n"
    //     str += `export class GameCfg {\n`
    //     for (const excelName in excelAndSheets) {
    //         // str += `    ${excelName} = {\n`
    //         // // 遍历表单
    //         // for (const sheetName in excelAndSheets[excelName]) {
    //         //     // 遍历索引，找出所有key
    //         //     let obj = excelAndSheets[excelName][sheetName]

    //         //     Object.keys(obj).forEach((key, index) => {
    //         //         // 索引组
    //         //         let _keys = obj[key].keys
    //         //         let desc1 = ""
    //         //         let desc2 = ""
    //         //         let aaaa = _keys.split(",")
    //         //         aaaa.forEach(element => {
    //         //             desc1 += toUpperCaseByStr(element)
    //         //             desc2 += `, "${element}"`
    //         //         })
    //         //         if (obj[key].isKey == 1) {
    //         //             str += `        ${sheetName}${desc1}: new ConfProxy(${excelName}${toUpperCaseByStr(sheetName)}${desc2})`
    //         //             // str += `        ${sheetName}_${_keys.replace(",", "_")}: new ConfProxy(${excelName}_${sheetName}${desc2})`
    //         //         } else if (obj[key].isKey == 2) {
    //         //             str += `        ${sheetName}${desc1}List: new ConfListProxy(${excelName}${toUpperCaseByStr(sheetName)}${desc2})`
    //         //             // str += `        ${sheetName}_${_keys.replace(",", "_")}List: new ConfListProxy(${excelName}_${sheetName}${desc2})`
    //         //         }
    //         //         // str += `\n`
    //         //         if (obj[key].isKey > 0) {
    //         //             str += index == Object.keys(obj).length - 1 ? `\n` : `,\n`
    //         //         }
    //         //     })
    //         // }

    //         // str += "    }"
    //         // str += "\n"
    //     }
    //     // str = str.substring(0, str.length - 2)
    //     str += `}`
    // }

    // 写入文件
    let outPath = hang == 4 ? `${outDir}cfg/excelConfig.ts` : `${outDir}/script/excelConfig.ts`
    fs.writeFile(outPath, str, function (err) {
        if (err) { console.log("excelConfig写入失败:", err) }
    })
}

// 字符串开头变为大写
function toUpperCaseByStr(txt) {
    return txt.slice(0, 1).toUpperCase() + txt.slice(1);
}

// 分析配置表
function doExcel(excel) {
    let fileDir = `${excelDir}/${excel}`
    let workbook = xl.readFile(fileDir)
    // 获取 Excel 中所有表名
    const sheetNames = workbook.SheetNames; // 返回 ['sheet1', 'sheet2']
    let _sheetNames = []
    sheetNames.forEach(element => {
        if (element.substr(0, 1) != "!") {
            _sheetNames.push(element)
        }
    })
    if (_sheetNames.length == 0) { return }
    console.log("...start excel " + excel)
    let excelName = excel.split(".")[0]
    // 加到【excelAndSheets】中
    excelAndSheets[excelName] = {}
    let obj = {}
    _sheetNames.forEach(sheetName => {
        let sheet = workbook.Sheets[sheetName]
        excelAndSheets[excelName][sheetName] = {}
        doSheet(obj, excelName, sheetName, sheet)
    })
    // console.log("aaaa:", `${outDir}${outJsonFolder}/${excelName}.json`)

    let _outDir = hang == 3 ? outDir + "resources/" : outDir
    fs.writeFile(`${_outDir}${outJsonFolder}/${excelName}.json`, JSON.stringify(obj), function (err) {
        if (err) { console.log(err) }
    })
}
//分析表单  obj数据，表名，表单名，表单数据内容
function doSheet(obj, excelName, sheetName, sheet) {
    if (obj[sheetName] == null) {
        obj[sheetName] = []
    }

    // 索引列表  example:{A:"",B:"id",C:"hp",D:"attack"}
    let keyList = {}
    // 类型列表 example:{A:"number",B:"string",C:"boolean",D:"{}[]"}
    let typeList = {}

    let _tempObj = {}
    for (const key in sheet) {
        // 单元格定位id example: A25
        const sheetKey = sheet[key];
        // A25->25
        let _hang = key.replace(/[^0-9]+/ig, "");
        // A25->A
        let _lie = key.replace(/[^A-Z]+/ig, "");
        if (_lie == "A") {
            if (Object.keys(_tempObj).length > 0) {
                obj[sheetName].push(_tempObj)
                _tempObj = {}
            }
        }
        if (sheet["A" + _hang] == null) {
            // 第一列没有配置参数,就跳过
            continue
        }

        // 储存类型
        if (_hang == 2) {
            typeList[_lie] = sheetKey.v
        }
        // 储存索引
        if (_hang == hang && sheetKey.v && typeList[_lie]) {
            // 正则抽取真正的值(去掉所有的*)
            let realVal = (sheetKey.v).replace(/[\*]/g, '')
            // 判断下作为key的字段
            let isKey = 0   // 0:不是key，1:单key，2:列表key
            if (sheetKey.v.slice(0, 2) == "**") {
                isKey = 2
            } else if (sheetKey.v.slice(0, 1) == "*") {
                isKey = 1
            }
            keyList[_lie] = realVal.split(",")[0]
            // id:string
            excelAndSheets[excelName][sheetName][keyList[_lie]] = {
                isKey: isKey,
                value: typeList[_lie],
                keys: realVal
            }
        }
        // 储存数据
        if (_hang >= 5 && keyList[_lie] != null && typeList[_lie] != null) {
            // 根据类型转化数据内容,在做储存处理
            let _val
            if (typeList[_lie] == 'number') {
                if (typeof sheetKey.v != 'number') {
                    _val = Number(sheetKey.v)
                    if (isNaN(_val)) {
                        console.log('\x1b[33m%s\x1b[0m', "警告,出现NaN数据类型,表名:" + excelName + " 表单:" + sheetName + " 索引:" + keyList[_lie] + " 表格ID:" + key)
                    }
                } else {
                    _val = sheetKey.v
                }
            } else if (typeList[_lie] == 'boolean') {
                _val = sheetKey.v == "true" ? true : false
            } else if (typeList[_lie][0] == "{" || typeList[_lie].slice(-1) == "]") {
                try {
                    _val = JSON.parse(sheetKey.v)
                } catch (error) {
                    console.log('\x1b[33m%s\x1b[0m', "警告:json格式错误,表名:" + excelName + " 表单:" + sheetName + " 索引:" + keyList[_lie] + " 表格ID:" + key)
                }
            } else if (typeList[_lie] == 'string') {
                if (typeof sheetKey.v == 'number') {
                    _val = (sheetKey.v).toString()
                } else {
                    _val = sheetKey.v
                }
            } else {
                _val = sheetKey.v
            }
            _tempObj[keyList[_lie]] = _val
        }
    }
    // 储存最后一个
    obj[sheetName].push(_tempObj)
}

// //--------------------start-------------------  控制台人机交互
function readSyncByRl(tips) {
    tips = tips || '> ';

    return new Promise((resolve) => {
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        rl.question(tips, (answer) => {
            rl.close();
            resolve(answer.trim());
        });
    });
}

function start(res){
    // console.log(res);
    if (res != "c" && res != "s" && res != "C" && res != "S") {
        console.log("非法输入")
        return
    }
    if (res == "c" || res == "C") {
        hang = 3
    } else {
        hang = 4
    }

    if (hang == 3) {
        outDir = path.clientPath
    } else {
        outDir = path.serverPath
    }

    // console.log('\x1b[37m') // 控制台文字切为白色
    // 解析表格
    fs.readdirSync(excelDir).forEach(excelName => {
        if ((excelName.substr(-4) == ".xls" || excelName.substr(-5) == ".xlsx") && excelName.substr(0, 1) != "~") {
            excelFiles.push(excelName.split(".")[0])
            doExcel(excelName)
        }
    })
    // 生成配置表读取文件(ts)
    exeExcelConfig()

    console.log('\x1b[32m%s\x1b[0m', "...completed!")
    // readline.question('enter key exit ')
    readSyncByRl('enter key exit:').then((res) => {

    })
}

start("c")
// //---------------------end--------------------  控制台人机交互