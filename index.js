const XLSX = require('xlsx')
const https = require('https')
const TelegramBot = require('node-telegram-bot-api')
const cron = require('node-cron');
const fs = require('fs');
const dotenv = require('dotenv')
dotenv.config()


const token = process.env.TOKEN || ""
const chatId = process.env.CHAT_ID || 0
const bot = new TelegramBot(token, {polling: true, none_stop: true})


cron.schedule('* * * * *', () => {

  const parsedData = parseXLSX(filePath)
  const filteredData = getFalseStatus(parsedData)
  const finalResult = randomizer(filteredData, parsedData)

  let resultString = ''
  for (const key in finalResult) {
    if (finalResult.hasOwnProperty(key)) {
      resultString += key + ': ' + finalResult[key] + ', \n '
    }
  }
  resultString = resultString.slice(0, -2)

  if (!resultString) {
    bot.sendMessage(chatId, 'Something went wrong')
  }
  bot.sendMessage(chatId, resultString)
});
bot.on('message', (msg) => {

  // Check if the message has a document
  if (msg.document) {
    const fileId = msg.document.file_id;

    bot.getFile(fileId).then((fileInfo) => {
      const fileUrl = `https://api.telegram.org/file/bot${token}/${fileInfo.file_path}`;

      // Download the file
      const downloadPath = `./list.xlsx`;
      const fileStream = fs.createWriteStream(downloadPath);

      https.get(fileUrl, (response) => {
        response.pipe(fileStream);

        fileStream.on('finish', () => {
          fileStream.close();
          console.log(`File downloaded: ${downloadPath}`);
          bot.sendMessage(chatId, 'File downloaded successfully!');
        });
      });


    });
  }
});


const filePath = 'list.xlsx'

function parseXLSX(filePath) {
  try {
    // Read the XLSX file
    const workbook = XLSX.readFile(filePath);

    // Assume the first sheet is the one you're interested in
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Parse the sheet into a JSON object
    const data = XLSX.utils.sheet_to_json(sheet, {header: 'Topic'});

    return data;
  } catch (error) {
    console.error('Error parsing XLSX file:', error.message);
    return null;
  }
}


const fieldName = 'Status' // Specify the field name you want to update
const newValue = true // Specify the new value for the field

const getFalseStatus = (parsedData) => {
  return parsedData.filter((item) => !item.Status)
}

const randomizer = (filteredData, parsedData) => {
  let targetRow
  if (!filteredData || filteredData.length < 2) {
    const index = parsedData.findIndex((element) => {
      return filteredData.some((compareElement) => {
        // Compare elements based on Topic and Speaker
        return element.Topic === compareElement.Topic && element.Speaker === compareElement.Speaker;
      });
    });
    targetRow = index + 1
    updateFieldValue(filePath, targetRow, fieldName, newValue)
    return parsedData[index]

  } else {
    const result = Math.floor(Math.random() * filteredData.length);
    targetRow = result + 1;
    updateFieldValue(filePath, targetRow, fieldName, newValue)
    return filteredData[result]
  }
}

function updateFieldValue(filePath, targetRow, fieldName, newValue) {
  try {
    // Read the Excel file
    const workbook = XLSX.readFile(filePath);

    // Assume the first sheet is the one you want to modify
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Check if the targetRow is within the range of the sheet
    const range = XLSX.utils.decode_range(sheet['!ref']);
    if (targetRow < range.s.r || targetRow > range.e.r) {
      console.error('Target row is out of range.');
      return;
    }

    // Find the column index for the specified fieldName
    const columnIndex = Object.keys(sheet)
      .filter(cell => XLSX.utils.decode_cell(cell).r === range.s.r && sheet[cell].v === fieldName)
      .map(cell => XLSX.utils.decode_cell(cell).c)[0];

    if (columnIndex === undefined) {
      console.error(`Field "${fieldName}" not found.`);
      return;
    }

    // Update the value of the specified field in the target row
    const cellToUpdate = XLSX.utils.encode_cell({r: targetRow, c: columnIndex});

    // Check if the cell exists before updating its value
    if (sheet[cellToUpdate]) {
      sheet[cellToUpdate].v = newValue;
    } else {
      console.error(`Cell ${cellToUpdate} does not exist.`);
      return;
    }

    // Update the worksheet with the new content
    workbook.Sheets[sheetName] = sheet;

    // Write the modified workbook back to the same file
    XLSX.writeFile(workbook, filePath);

    console.log(`Field "${fieldName}" in row ${targetRow} updated successfully.`);
  } catch (error) {
    console.error('Error updating field value in Excel file:', error.message);
  }
}


