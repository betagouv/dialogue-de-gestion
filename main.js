Promise = require('bluebird')
const rp = require('request-promise')
const url = require('url')

const utils = require('./utils')
const sid = process.env.SHEET_SID
var qs = { key: process.env.API_KEY }

const schemas = [{
  name: 'orders',
  prefix: 'BC',
  fields: [
    { name: 'Intitulé' },
    { name: 'MontantTTC' },
    { name: 'Numero' },
    { name: 'DateEJ', process: utils.ExcelDateToJSDate },
    { name: 'RaP' },
    { name: 'DatePaiement', process: utils.ExcelDateToJSDate },
    { name: 'RaR' },
    { name: 'DateRbm', process: utils.ExcelDateToJSDate },
    { name: 'Convention' },
    { name: 'TypeConvention' },
  ]
}, {
  name: 'payments',
  prefix: 'Paiement',
  fields: [
    { name: 'EJ' },
    { name: 'Date', process: utils.ExcelDateToJSDate },
    { name: 'Montant' },
  ]
}]


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


//*
var searchBatch = new url.URLSearchParams()
searchBatch.append('key', qs.key)
searchBatch.append('valueRenderOption', 'UNFORMATTED_VALUE')
searchBatch.append('dateTimeRenderOption', 'SERIAL_NUMBER')
schemas.forEach(schema => schema.fields.forEach(field => searchBatch.append('ranges', schema.prefix + field.name)))
console.error(searchBatch.toString())

const conf = {
  json: true,
  uri: `https://sheets.googleapis.com/v4/spreadsheets/${sid}/values:batchGet?${searchBatch}`,
};
rp(conf)
.then(structure)
.then(data => {
  console.log(JSON.stringify(data, null, 2))
  console.log([data.orders.length, data.payments.length])
})
//*/

function structure(data) {
  // From table to vector
  data.valueRanges.forEach(range => {
    range.values = range.values.map(rowData => rowData[0] || 0 )
  })

  return schemas.reduce((accum, schema) => {
    var objects = data.valueRanges[accum.rangeIndex].values.map((v, lineNumber) => {

      return schema.fields.reduce((object, field, fieldIndex) => {
        const value = data.valueRanges[accum.rangeIndex + fieldIndex].values[lineNumber]
        object[field.name] = field.process ? field.process(value) : value

        return object
      }, {})
    })

    accum.rangeIndex = accum.rangeIndex + schema.fields.length
    accum.models[schema.name] = objects

    return accum
  }, { rangeIndex: 0, models: {} }).models
}

/*
const named = require('./bcnamed.json')

var objects = structure(named);

console.log(objects)
var orders = objects.orders;
orders = orders.filter((order) => order.TypeConvention != 'Délégation de gestion')


if (0) {
  // Current year
  const startOfYear = new Date(2018, 0, 1)
  orders = orders.filter((order) => startOfYear <= order.BCDateEJ)
}

orders = orders.filter((order) => order.Numero)
orders = orders.slice(1)
orders.sort((a, b) => b.MontantTTC - a.MontantTTC)

console.log(JSON.stringify(orders, null, 2))
console.log(JSON.stringify(orders.length, null, 2))
console.log(JSON.stringify(orders.reduce((s,o) => { console.log(s); return s + o.MontantTTC }, 0), null, 2))

//*/
/*
const full = require('./fulldata.json')
console.log(full.namedRanges.map((r) => r.name).sort())
//*/
