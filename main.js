var XLSX = require('xlsx')
var fs = require('fs')
var _ = require('lodash')

// Parse the file
var buf = fs.readFileSync("18YearLotto.xlsx")
var wb = XLSX.read(buf, {type:'buffer'})

// Access the sheet
let Sheet = wb.Sheets.Sheet1

// Delete the excess data
delete(Sheet['!ref'])
delete(Sheet['!margins'])
delete(Sheet['A1'])
delete(Sheet['B1'])
delete(Sheet['C1'])
delete(Sheet['D1'])
delete(Sheet['E1'])
delete(Sheet['F1'])
delete(Sheet['G1'])
delete(Sheet['H1'])

// Stores the list of number frequencies
let frequencies = []

// Itterate over the remaining data
_.forEach(wb.Sheets.Sheet1, function(value, key) {
  // If this isnt an A_ row
  if (!key.includes('A')) {
    // Access the cell value
    let v = value.v
    // If this number has not yet appeared in the list of frequent numbers
    if (_.find(frequencies, { 'number': v }) == undefined) {
      // Initialise the count
      frequencies.push({
        number: v,
        frequency: 0
      })
    }
    // Get the key of the object that contains this number
    let objKey = _.findKey(frequencies, { 'number': v })
    // Increment the count
    frequencies[objKey].frequency++
  }
})

// Sort by frequency
frequencies = _.orderBy(frequencies, ['frequency'], ['desc'])

// List the frequencies
_.forEach(frequencies, function(value, key) {
  if (key !== 0) console.log(value.number+' appears '+value.frequency+' times')
})
