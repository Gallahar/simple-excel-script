const XLSX = require('xlsx')
const prompt = require('prompt-sync')

const inputFilename = prompt()('please write name of .xlsx workbook: ').trim()

const outputFilename = prompt()('choose name of output .xlsx file: ').trim()

const workbook = XLSX.readFile(inputFilename, { cellFormula: true })

const worksheet = workbook.Sheets['Лист1']

const w = XLSX.utils.sheet_to_json(worksheet)

for (let row of w) {
	if (row['год2'] || row['год3']) {
		if (
			row['период 1'] === row['период 2'] &&
			row['период 2'] === row['период 3'] &&
			row['период 1'] === row['период 3']
		) {
			row['год1'] = row['год3']
			row['год3'] = ''
			row['год2'] = ''
			continue
		}
		if (row['период 1'] === row['период 2']) {
			const temp = row['год2']
			row['год1'] = temp
			row['год2'] = ''
			if (row['период 3']) {
				const temp2 = row['год3']
				row['год2'] = temp2
				row['год3'] = ''
			}
		}
		if (row['период 2'] === row['период 3']) {
			const temp = row['год3']
			row['год2'] = temp
			row['год3'] = ''
		}

		if (row['период 1'] === row['период 3']) {
			const temp = row['год3']
			row['год1'] = temp
			row['год3'] = ''
		}
	}
}

const newWorkSheet = XLSX.utils.json_to_sheet(w)

const newWorkBook = XLSX.utils.book_new()

XLSX.utils.book_append_sheet(newWorkBook, newWorkSheet, 'Лист1')

XLSX.writeFile(newWorkBook, outputFilename, { compression: true })
