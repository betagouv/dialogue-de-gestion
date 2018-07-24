Promise = require('bluebird')
const rp = require('request-promise')
const url = require('url')

const utils = require('./utils')
const sid = process.env.SHEET_SID
var qs = { key: process.env.API_KEY }

const namedRanges = ['BCIntitulé', 'BCMontantTTC', 'BCNumero', 'BCDateEJ', 'BCDatePaiement', 'BCRaP', 'BCConvention', 'BCTypeConvention']



/*
rp({
  uri: `https://sheets.googleapis.com/v4/spreadsheets/${sid}`,
  qs: qs,
  json: true
})
.then(data => {
  console.log(JSON.stringify(data, null, 2))
})
//*/


/*
var searchBatch = new url.URLSearchParams()
searchBatch.append('key', qs.key)
searchBatch.append('valueRenderOption', 'UNFORMATTED_VALUE')
searchBatch.append('dateTimeRenderOption', 'SERIAL_NUMBER')
namedRanges.forEach(range => searchBatch.append('ranges', range))
console.error(searchBatch.toString())

const conf = {
  json: true,
  uri: `https://sheets.googleapis.com/v4/spreadsheets/${sid}/values:batchGet?${searchBatch}`,

};
rp(conf)
//.then(structure)
.then(data => {
  console.log(JSON.stringify(data, null, 2))
})
//*/

function structure(orderData) {
  orderData.valueRanges.forEach(range => {
    range.values = range.values.map(rowData => rowData[0] || 0 )
  })

  var orders = orderData.valueRanges[0].values.map((v, lineNumber) => {
    var order = namedRanges.reduce((obj, name, fieldNumber) => {
      obj[name] = orderData.valueRanges[fieldNumber].values[lineNumber]
      return obj
    }, {})

    order.BCDateEJ = utils.ExcelDateToJSDate(order.BCDateEJ);
    return order
  })

  return orders
}


const named = require('./bcnamed.json')
var orders = structure(named);

orders = orders.filter((order) => order.BCTypeConvention != 'Délégation de gestion')


/*
// Current year
const startOfYear = new Date(2018, 0, 1)
orders = orders.filter((order) => startOfYear <= order.BCDateEJ)
*/


orders = orders.filter((order) => order.BCNumero)
orders = orders.slice(1)
orders.sort((a, b) => b.BCMontantTTC - a.BCMontantTTC)

console.log(JSON.stringify(orders, null, 2))
console.log(JSON.stringify(orders.length, null, 2))
console.log(JSON.stringify(orders.reduce((s,o) => { console.log(s); return s + o.BCMontantTTC }, 0), null, 2))

//*/
/*
const full = require('./fulldata.json')
console.log(full.namedRanges.map((r) => r.name).sort())
//*/
