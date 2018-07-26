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
  constr: () => { return { payments: [], reimbursements: [] } },
  fields: [
    { name: 'Intitulé' },
    { name: 'MontantTTC' },
    { name: 'Numero' },
    { name: 'DateEJ', process: utils.ExcelDateToJSDate },
    { name: 'RaP' },
    { name: 'DatePaiement', process: utils.ExcelDateToJSDate },
    { name: 'RaR' },
    { name: 'DateRbm', process: utils.ExcelDateToJSDate },
    { name: 'FondsPropres' },
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
}, {
  name: 'reimbursements',
  prefix: 'RbmAlloc',
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
      var obj = schema.fields.reduce((object, field, fieldIndex) => {
        const value = data.valueRanges[accum.rangeIndex + fieldIndex].values[lineNumber]
        object[field.name] = field.process ? field.process(value) : value

        return object
      }, Object.assign({}, schema.constr ? schema.constr() : {}))

      obj.index = lineNumber

      return obj;
    })

    accum.rangeIndex = accum.rangeIndex + schema.fields.length
    accum.models[schema.name] = objects.slice(1) // Drop header line

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

  const props = ['payments', 'reimbursements']
  props.forEach(field => {
    data[field].forEach((obj) => {
      if (! keyedOrders[obj.EJ]) {
        return
      }

      var order = keyedOrders[obj.EJ]
      if (! order.Numero) {
        return
      }

      if (order.Numero !== obj.EJ) {
        console.log([order, obj])
        throw e
      }

      order[field].push(obj)
    })//*/
  })

  return data
}

function computeSums(data) {
  data.orders.forEach((order) => {
    order.RaPComputed = order.MontantTTC - order.payments.reduce((sum, payment) => sum + payment.Montant, 0)
    order.RaRComputed = order.MontantTTC - order.FondsPropres - order.reimbursements.reduce((sum, payment) => sum + payment.Montant, 0)
  })

  return data
}

const startOfYear = new Date(2018, 0, 1)
fluxPropres = o => ['Fonds propres', 'Refacturation', 'Fonds de concours'].indexOf(o.TypeConvention) >= 0
fluxCourants = o => startOfYear <= Math.max(o.DateEJ, o.DatePaiement, o.DateRbm)

var headers = [
'Activités/Projets',
'Type de mouvement',
'Objet de la dépense',
'Prestataires/Marchés',
'Références (n° bdc/n°Conv)',
'RàP 2016 (CP 2017 sur AE<=2016)',
'AE janv-Avril',
'AE mai-aout',
'AE sept-dec',
'CP janv-Avril',
'CP mai-aout',
'CP sept-dec',
'RàP 2017 sur CP 2018',
'Statut',
'Commentaires',
]

function restitute(data) {

  var categories = [, {
    name: 'DoneMoins1',
    predicate: (o) => o.DateEJ <= startOfYear && (! o.RaP) && (! o.RaR) && (startOfYear <= Math.max(o.DateEJ, o.DatePaiement, o.DateRbm))
  }, {
    name: 'RAPMoins1',
    predicate: (o) => o.DateEJ <= startOfYear && o.RaP
  }, {
    name: 'RARMoins1',
    predicate: (o) => o.DateEJ <= startOfYear && o.RaR
  }, {
    name: 'Done',
    predicate: (o) => startOfYear < o.DateEJ && o.Numero && (! o.RaP) && (! o.RaR)
  }, {
    name: 'RAP',
    predicate: (o) => startOfYear < o.DateEJ && o.Numero && o.RaP
  }, {
    name: 'RAR',
    predicate: (o) => startOfYear < o.DateEJ && o.Numero && o.RaR
  }, {
    name: 'RAE',
    predicate: (o) => o
  }]

  data.orders.sort(function(a, b) { return a.DateEJ - b.DateEJ })
  const dialogue = categories.reduce((result, category) => {
    const selection = result.orders.filter(category.predicate)

    result.orders = result.orders.filter(o => ! category.predicate(o))
    result.output[category.name] = selection.map(o => {
      var operations = []

      var res = headers.map(o => '')
      res[0] = o.Intitulé
      res[1] = 'Dépense UO DINSIC (yc sur facture interne interministérielle)'
      // rétablissements de crédits uo dinsic (cf convention)

      operations.push(res)

      return operations
    })

    return result
  }, { orders: data.orders.filter(fluxPropres).filter(fluxCourants), output: {} })

  console.log(JSON.stringify({
    full: data,
    //RAPValidation: data.orders.filter(o => o.TypeConvention != 'Délégation de gestion' && o.RaPComputed != o.RaP),
    //RARValidation: data.orders.filter(o => ['Refacturation', 'Fonds de concours'].indexOf(o.TypeConvention) >= 0 && o.RaRComputed != o.RaR),
    dialogue: dialogue,
    counts: [data.orders.length, data.payments.length],
  }, null, 2))
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
