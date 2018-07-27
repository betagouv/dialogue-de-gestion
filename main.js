Promise = require('bluebird')
const fs = Promise.promisifyAll(require('fs'))
const rp = require('request-promise')
const url = require('url')

const utils = require('./utils')
const sid = process.env.SHEET_SID
var qs = { key: process.env.API_KEY }

const year = 2018
const startOfYear = new Date(year, 0, 1)
const endOfYear = new Date(year + 1, 0, 1)
const dates = [startOfYear, new Date(year, 4, 1), new Date(year, 8, 1), endOfYear]
periodIdx = d => dates.reduce((a, p) => a + (p < d ? 1 : 0), -1)

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
  ],
  computedFields: [{
    name: 'RAPEJ',
    formula: order => {
      return order.MontantTTC - order.payments.reduce((accum, payment) => accum + (payment.Date < startOfYear ? payment.Montant : 0), 0)
    }
  }, {
    name: 'RAREJ',
    formula: order => {
      return order.RaR + order.reimbursements.reduce((accum, reimbursement) => accum + (startOfYear < reimbursement.Date ? reimbursement.Montant : 0), 0)
    }
  }]
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
      return obj
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
    })
  })
  return data
}

function computeFields(data) {
  schemas.forEach((schema) => {
    if (! schema.computedFields) {
      return
    }

      var objects = data[schema.name]
      objects.forEach(obj => {
        schema.computedFields.forEach(field => {

          obj[field.name] = field.formula(obj)
        })
      })
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

const fluxPropres = (o) => ['Fonds propres', 'Refacturation', 'Fonds de concours', 'Transfert de crédit'].indexOf(o.TypeConvention) >= 0
const fluxCourants = (o) => startOfYear <= Math.max(o.DateEJ, o.DatePaiement, o.DateRbm)

var headers = [
'Activités/Projets',
'Type de mouvement',
'Objet de la dépense',
'Prestataires/Marchés',
'Références (n° bdc/n°Conv)',
'RàP 2016 (CP 2017 sur AE<=2016)',
'AE janv-avril',
'AE mai-aout',
'AE sept-dec',
'CP janv-avril',
'CP mai-aout',
'CP sept-dec',
'RàP 2017 sur CP 2018',
'Statut',
'Commentaires',
'Catégorie'
]

idxH = n => headers.indexOf(n)
hasCP = row => ['CP janv-avril', 'CP mai-aout', 'CP sept-dec', 'RàP 2017 sur CP 2018'].map(n => row[idxH(n)]).reduce((a, v) => a || v, false)

var categories = [{
  name: 'DoneMoins1',
  predicate: (o) => o.DateEJ <= startOfYear && (! o.RaP) && (! o.RaR) && (startOfYear <= Math.max(o.DateEJ, o.DatePaiement, o.DateRbm)),
  status: 'Terminé',
}, {
  name: 'RAPMoins1',
  predicate: (o) => o.DateEJ <= startOfYear && o.RaP,
  status: 'SF Partiel',
}, {
  name: 'RARMoins1',
  predicate: (o) => o.DateEJ <= startOfYear && o.RaR,
  status: 'EL Partiel',
}, {
  name: 'Done',
  predicate: (o) => startOfYear < o.DateEJ && o.Numero && (! o.RaP) && (! o.RaR),
  status: 'Terminé',
}, {
  name: 'RAP',
  predicate: (o) => startOfYear < o.DateEJ && o.Numero && o.RaP,
  status: 'SF Partiel',
}, {
  name: 'RAR',
  predicate: (o) => startOfYear < o.DateEJ && o.Numero && o.RaR,
  status: 'EL Partiel',
}, {
  name: 'RAE',
  predicate: (o) => o,
  status: 'Prévisionnel',
}]

function restitute(data) {
  data.orders.sort(function(a, b) { return (a.DateEJ - b.DateEJ) })

  const dialogue = categories.reduce((result, category) => {
    const selection = result.orders.filter(category.predicate)

    result.orders = result.orders.filter(o => ! category.predicate(o))
    result.output[category.name] = selection.map(o => {
      var operations = []

      var gen = (type, cat) => [
        o.Intitulé,
        type,
        o.Intitulé + ' - ' + o.Numero, // Objet de la dépense,
        o.TypeConvention, // Prestataires/Marchés,
        o.Numero, // Références (n° bdc/n°Conv),
        0, // RàP 2016 (CP 2017 sur AE<=2016)',
        0,0,0, // AE,
        0,0,0, // CP
        0, // RAP
        category.status, // Statut,
        0, // Commentaires,
        cat
      ]

      var order = gen('Dépense UO DINSIC (yc sur facture interne interministérielle)', 'Dépense AE')
      if (o.DateEJ < endOfYear) {
        order[idxH('AE janv-avril') + periodIdx(o.DateEJ)] = o.RAPEJ
      } else {
        order[idxH('Commentaires') + periodIdx(o.DateEJ)] = o.RAPEJ
      }
      operations.push(order)

      var pastPayments = gen('Dépense UO DINSIC (yc sur facture interne interministérielle)', 'Dépense CP passée')
      o.payments.forEach(payment => {
        if (payment.Date < startOfYear) {
          return
        }

        pastPayments[idxH('CP janv-avril') + periodIdx(payment.Date)] += payment.Montant
      })
      if (hasCP(pastPayments)) {
        operations.push(pastPayments)
      }

      var futurePayments = gen('Dépense UO DINSIC (yc sur facture interne interministérielle)', 'Dépense CP prévue')
      futurePayments[idxH('CP janv-avril') + periodIdx(o.DatePaiement)] += o.RaP
      if (hasCP(futurePayments)) {
        operations.push(futurePayments)
      }

      var pastReimbursements = gen('rétablissements de crédits uo dinsic (cf convention)', 'Rétablissement passé')
      o.reimbursements.forEach(reimbursement => {
        if (reimbursement.Date < startOfYear) {
          return
        }
        pastReimbursements[idxH('AE janv-avril') + periodIdx(reimbursement.Date)] += -reimbursement.Montant
        pastReimbursements[idxH('CP janv-avril') + periodIdx(reimbursement.Date)] += -reimbursement.Montant
      })
      if (hasCP(pastReimbursements)) {
        operations.push(pastReimbursements)
      }

      var futureReimbursements = gen('rétablissements de crédits uo dinsic (cf convention)', 'Rétablissement prévu')
      if (o.DateRbm < endOfYear) {
        futureReimbursements[idxH('AE janv-avril') + periodIdx(o.DateRbm)] += -o.RaR
      } else {
        futureReimbursements[idxH('Commentaires')] += -o.RaR
      }
      futureReimbursements[idxH('CP janv-avril') + periodIdx(o.DateRbm)] += -o.RaR
      if (hasCP(futureReimbursements)) {
        operations.push(futureReimbursements)
      }

      return operations
    })

    return result
  }, { orders: data.orders.filter(fluxPropres).filter(fluxCourants), output: {} })

  return {
    full: data,
    //RAPValidation: data.orders.filter(o => o.TypeConvention != 'Délégation de gestion' && o.RaPComputed != o.RaP),
    //RARValidation: data.orders.filter(o => ['Refacturation', 'Fonds de concours'].indexOf(o.TypeConvention) >= 0 && o.RaRComputed != o.RaR),
    dialogue: dialogue,
    counts: [data.orders.length, data.payments.length],
  }
}

var searchBatch = new url.URLSearchParams()
searchBatch.append('key', qs.key)
searchBatch.append('valueRenderOption', 'UNFORMATTED_VALUE')
searchBatch.append('dateTimeRenderOption', 'SERIAL_NUMBER')
schemas.forEach(schema => schema.fields.forEach(field => searchBatch.append('ranges', schema.prefix + field.name)))
const conf = {
  json: true,
  uri: `https://sheets.googleapis.com/v4/spreadsheets/${sid}/values:batchGet?${searchBatch}`,
}

var express = require('express')
var app = express()

app.get('/', function (req, res) {
  fs.readFileAsync('data.json')
  .then(content => JSON.parse(content))
  .catch((error) => {
    console.error(searchBatch.toString())
    return rp(conf)
    .then(data => {
      return fs.writeFileAsync('data.json', JSON.stringify(data, null, 2), 'utf-8')
      .then(() => data)
    })
  })
  .then(structure)
  .then(createRelationships)
  .then(computeFields)
  .then(computeSums)
  .then(restitute)
  .then(data => {
    res.header({ 'Access-Control-Allow-Origin': '*' })
    res.json(data)
  })
})

app.listen(3000, () => console.log('App listening on port 3000!'))
