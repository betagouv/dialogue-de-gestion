Promise = require('bluebird')
const fs = Promise.promisifyAll(require('fs'))
const rp = require('request-promise')
const url = require('url')

const utils = require('./utils')
const sid = process.env.SHEET_SID
var qs = { key: process.env.API_KEY }

const schemas = [{
  name: 'orders',
  prefix: 'BC',
  key: 'Numero',
  constr: () => { return { payments: [] } },
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
      }, Object.assign({}, schema.constr ? schema.constr() : {}))
    })

    accum.rangeIndex = accum.rangeIndex + schema.fields.length
    accum.models[schema.name] = objects

    return accum
  }, { rangeIndex: 0, models: {} }).models
}

function createRelationships(data) {
  const keyedOrders = data.orders.reduce((obj, order) => {
    if (! order.Numero) {
      return obj
    }

    obj[order.Numero] = order

    return obj
  }, {})

  data.payments.forEach((payment) => {
    if (! keyedOrders[payment.EJ]) {
      return
    }

    var order = keyedOrders[payment.EJ]
    if (! order.Numero) {
      return
    }

    if (order.Numero !== payment.EJ) {
      console.log([order, payment])
      throw e
    }

    order.payments.push(payment)
  })

  return data
}

function computeSums(data) {
  data.orders.forEach((order) => {
    order.RaPComputed = order.MontantTTC - order.payments.reduce((sum, payment) => sum + payment.Montant, 0)
  })

  return data
}

function restitute(data) {
  console.log(JSON.stringify(data, null, 2))
  console.warn(JSON.stringify(data.orders, null, 2))
  console.warn(JSON.stringify(data.orders.filter(o => o.RaPComputed != o.RaP), null, 2))
  console.warn([data.orders.length, data.payments.length])
}

var searchBatch = new url.URLSearchParams()
searchBatch.append('key', qs.key)
searchBatch.append('valueRenderOption', 'UNFORMATTED_VALUE')
searchBatch.append('dateTimeRenderOption', 'SERIAL_NUMBER')
schemas.forEach(schema => schema.fields.forEach(field => searchBatch.append('ranges', schema.prefix + field.name)))
const conf = {
  json: true,
  uri: `https://sheets.googleapis.com/v4/spreadsheets/${sid}/values:batchGet?${searchBatch}`,
};

(new Promise((resolve) => { resolve(require('./data.json')) }))
.catch(() => {
  console.error(searchBatch.toString())
  return rp(conf)
  .then(data => fs.writeFileAsync('data.json', JSON.stringify(data, null, 2), 'utf-8').then(() => data))
})
.then(structure)
.then(createRelationships)
.then(computeSums)
.then(restitute)
