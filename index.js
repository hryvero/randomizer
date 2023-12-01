const XLSX = require('xlsx')
const https = require('https')
const TelegramBot = require('node-telegram-bot-api')
const cron = require('node-cron')
const fs = require('fs')
const dotenv = require('dotenv')
dotenv.config()


const token = process.env.TOKEN || ''
const chatId = process.env.CHAT_ID || 0
const bot = new TelegramBot(token, {polling: true, none_stop: true})
const cronTime = '30 17 * * 5' // Every Friday at 17:30
const fieldName = 'Status' // Specify the field name you want to update
const newValue = true // Specify the new value for the field

cron.schedule(cronTime, () => {

	const parsedData = parseXLSX(filePath)
	const filteredData = getFalseStatus(parsedData)
	const finalResult = randomizer(filteredData, parsedData)

	let resultString = ''
	for (const key in finalResult.randomData) {
		if (finalResult.randomData.hasOwnProperty(key)) {
			resultString += key + ': ' + finalResult.randomData[key] + ', \n '
		}
	}
	resultString = resultString.slice(0, -2)

	if (!resultString) {
		bot.sendMessage(chatId, 'Opps...It seems that list is over. Please send new file')
	}
	bot.sendMessage(chatId, resultString)
	bot.sendMessage(chatId, `${finalResult.randomData.Speaker}, are you ready to perform next Friday?`,{
		reply_markup: {
			keyboard: [['Yes'], ['No']]
		},
	})

	bot.on('message', async msg => {
		const username = msg.from.username
		console.log(username)

		// Check if the username matches the allowed username
		if (username === finalResult.randomData.Username) {

			const userId = msg.from.id
			const firstName = msg.from.first_name

			// Extracted user information
			const responseText = `User ID: ${userId}\nUsername: @${username}\nFirst Name: ${firstName}`

			// Send the extracted information back to the user
			updateFieldValue(filePath, finalResult.targetRow, fieldName, newValue)
			await bot.sendMessage(chatId, responseText)
		} else {
			// Inform the user that only the specified username is allowed
			await bot.sendMessage(chatId, `Sorry, only @${finalResult.randomData.Username} is allowed to answer.`)
		}
	})



})

const filePath = 'list.xlsx'

bot.on('message', (msg) => {

	// Check if the message has a document
	if (msg.document) {
		const fileId = msg.document.file_id

		bot.getFile(fileId).then((fileInfo) => {
			const fileUrl = `https://api.telegram.org/file/bot${token}/${fileInfo.file_path}`

			// Download the file
			const downloadPath = './list.xlsx'
			const fileStream = fs.createWriteStream(downloadPath)

			https.get(fileUrl, (response) => {
				response.pipe(fileStream)

				fileStream.on('finish', () => {
					fileStream.close()
					bot.sendMessage(chatId, 'File is successfully downloaded!')
				})
			})


		})
	}
})

function parseXLSX(filePath) {
	try {
		const workbook = XLSX.readFile(filePath)

		const sheetName = workbook.SheetNames[0]
		const sheet = workbook.Sheets[sheetName]

		const data = XLSX.utils.sheet_to_json(sheet, {header: 'Topic'})

		return data
	} catch (error) {
    bot.sendMessage(chatId, `Error parsing XLSX file: ${error.message}`)
		return null
	}
}


function getFalseStatus (parsedData){
	return parsedData.filter((item) => !item.Status)
}

function randomizer(filteredData, parsedData){

	let targetRow
	if (!filteredData || filteredData.length < 2) {
		const index = parsedData.findIndex((element) => {
			return filteredData.some((compareElement) => {
				// Compare elements based on Topic and Speaker
				return element.Topic === compareElement.Topic && element.Speaker === compareElement.Speaker
			})
		})
		targetRow = index + 1
		return {randomData: parsedData[index], targetRow}

	} else {
		const result = Math.floor(Math.random() * filteredData.length)
		targetRow = result + 1
		return {randomData: filteredData[result], targetRow}
	}
}

function updateFieldValue(filePath, targetRow, fieldName, newValue) {
	try {
		// Read the Excel file
		const workbook = XLSX.readFile(filePath)

		// Assume the first sheet is the one you want to modify
		const sheetName = workbook.SheetNames[0]
		const sheet = workbook.Sheets[sheetName]

		// Check if the targetRow is within the range of the sheet
		const range = XLSX.utils.decode_range(sheet['!ref'])
		if (targetRow < range.s.r || targetRow > range.e.r) {
			console.error('Target row is out ranged.')
			return
		}

		// Find the column index for the specified fieldName
		const columnIndex = Object.keys(sheet)
			.filter(cell => XLSX.utils.decode_cell(cell).r === range.s.r && sheet[cell].v === fieldName)
			.map(cell => XLSX.utils.decode_cell(cell).c)[0]

		if (columnIndex === undefined) {
			console.error(`Field "${fieldName}" is not found.`)
			return
		}

		// Update the value of the specified field in the target row
		const cellToUpdate = XLSX.utils.encode_cell({r: targetRow, c: columnIndex})

		// Check if the cell exists before updating its value
		if (sheet[cellToUpdate]) {
			sheet[cellToUpdate].v = newValue
		} else {
			console.error(`Cell ${cellToUpdate} does not exist.`)
			return
		}

		// Update the worksheet with the new content
		workbook.Sheets[sheetName] = sheet

		// Write the modified workbook back to the same file
		XLSX.writeFile(workbook, filePath)

		console.log(`Field "${fieldName}" in row  ${targetRow} successfully updated.`)
	} catch (error) {
		console.error('Error in updating file:', error.message)
	}
}


