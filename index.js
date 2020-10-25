const XLSX = require('xlsx')
const fs = require('fs')
const {
  isNil,
  isEmpty,
  intersection,
  without,
  concat,
  not,
  repeat,
} = require('ramda')

// parameter
const SIZE = '尺碼'
// const RANGE = 'A2:M10'
const SHEETNAME = '資料'
const RANGE = 'A2:M992'
// const RANGE = 'A922:M923'
const OUTPUTFILE = 'out.csv'

// main flow
var workbook = XLSX.readFile('source.xlsx')
var ws = workbook['Sheets'][SHEETNAME]
ws['!ref'] = RANGE

var oriData = XLSX.utils.sheet_to_csv(ws, { header: 1 })

// 步驟一: 資料整理(先拉出出所有資料的丈量單位&資料)

var modifyData = oriData
  .split('\n')
  .map((d) => {
    if (/,{12,}/.test(d) || isEmpty(d)) return false // check
    var cols = d.split(/,/)
    var [
      model,
      generalModel,
      size,
      productType,
      descript1,
      descript2,
      ...units
    ] = cols

    var pUnits = units
      .map((d) => {
        // d :: '袖長:64'
        if (checkUnit(d)) {
          return d.split(/:/)[0]
        } else {
          return
        }
      })
      .filter((d) => not(isNil(d)))
    var vUnits = units
      .map((d) => {
        // d :: '袖長:64'
        if (checkUnit(d)) {
          return d.split(/:/)
        } else {
          return
        }
      })
      .filter((d) => not(isNil(d)))
    return [
      model,
      generalModel,
      size,
      productType,
      pUnits,
      vUnits,
      descript1,
      descript2,
    ]
  })
  .filter((d) => d)

// 步驟二: 資料整理(先統計出各型號的丈量單位)

var unitifyData = getAllModelUniqueUnits(modifyData)

// 步驟三: 逐筆資料匯整分類(Grouping By 通用型號)

var finalData = modifyData.reduce((acc, cur) => {
  var [
    model,
    generalModel,
    size,
    productType,
    pUnits,
    vUnits,
    descript1,
    descript2,
  ] = cur
  if (generalModel === undefined) {
    return acc
  }
  var aggUnits = unitifyData[generalModel]
  var sizeData = repeat('--', aggUnits.length)
  sizeData[0] = size
  vUnits.forEach(([unitKey, unitVal]) => {
    if (aggUnits.indexOf(unitKey) > -1) {
      sizeData[aggUnits.indexOf(unitKey)] = unitVal
    } else {
      throw new Error('unit is not within Model Unit')
    }
  })
  if (acc[generalModel] === undefined) {
    // init acc
    return {
      ...acc,
      [generalModel]: [sizeData],
    }
  } else {
    // insert acc[generalModel]
    return {
      ...acc,
      [generalModel]: concat(acc[generalModel], [sizeData]),
    }
  }
}, {})

console.log('modifyData', modifyData)
console.log('unitifyData', unitifyData)
console.log('final', finalData)

// 步驟四: 寫入 csv 檔(組 html code)

var str = ''

Object.keys(finalData).map((guid) => {
  // console.log('guid', guid)
  str +=
    `${guid},` +
    `<table width="100%" border="0"><tbody><tr><td>版型建議</td><td> </td><td> </td><td> </td><td> </td><td> </td><td>單位cm</td></tr></tbody></table><table width="100%" border="1" style="font-size: 9pt; text-align:center;"><tbody><tr>`
  unitifyData[guid].map((d1) => {
    str += `<td>${d1}</td>`
  })
  str += `</tr>`
  finalData[guid].map((d2) => {
    str += `<tr>`
    d2.map((d3) => {
      str += `<td>${d3}</td>`
    })
    str += `</tr>`
  })
  str += `</tbody></table>`
  str += '\r\n'
})

console.log('str', str)
fs.writeFile(OUTPUTFILE, str, function (err, data) {
  if (err) {
    return console.log(err)
  }
  console.log(data)
})

// helper lib
function getAllModelUniqueUnits(dataset) {
  return dataset.reduce((acc, cur) => {
    var [model, generalModel, size, productType, units, ...rest] = cur

    if (acc[generalModel]) {
      // exists
      var existUnits = intersection(acc[generalModel], units)
      return {
        ...acc,
        [generalModel]: concat(acc[generalModel], without(existUnits, units)),
      }
    } else {
      // undefined
      return { ...acc, [generalModel]: concat([SIZE], units) }
    }
  }, {})
}

// check spec if valid (檢查要排除的條件: 現況 '1000' or 空值)
function checkUnit(d) {
  return d === '1000' || isNil(d) || isEmpty(d) ? false : true
}
