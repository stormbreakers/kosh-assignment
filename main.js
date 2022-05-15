const fs = require('fs');
const { read, utils, writeFile } = require('xlsx');
const filePath = './asset/read.docx';

function readFileFromPath(filePath) {
  return new Promise((resolve, reject) => {
    const body = [];
    const readableStream = fs.createReadStream(filePath);

    readableStream.on('error', function (error) {
      reject(err);
    });

    readableStream.on('data', (chunk) => {
      body.push(chunk);
    });

    readableStream.on('end', (chunk) => {
      const buffer = Buffer.concat(body);
      resolve(buffer.toString('base64'));
    });
  });
}

function readData(data) {
  const wb = read(data, { type: 'base64' });
  return utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
}

function _formatData(data) {
  const descriptionIndex = data[0].findIndex((val) => val === 'Description');
  return data.map((val, index) => {
    if (index === 0) {
      val.push('Flag', 'Description');
    } else {
      const includesAeps = val?.[descriptionIndex]?.includes('AEPS');
      if (includesAeps) {
        val.push('AEPS')
      }
      const includesFeechg = val?.[descriptionIndex]?.includes('FEE CHG');
      if (includesFeechg) {
        val.push('FEE CHG')
      }
      if (!includesAeps && !includesFeechg) {
        val.push('')
      }
      val.push(...val[descriptionIndex].split('/'))
    }
    return val;
  });
}

function _writeDateToExcel(data) {
  const outputData = utils.aoa_to_sheet(data);
  const newWb = utils.book_new();
  utils.book_append_sheet(newWb, outputData, "Sheet1");

  /* generate an XLSX file */
  writeFile(newWb, './asset/output.xlsx', { type: 'file' });
}

readFileFromPath(filePath)
  .then(readData)
  .then(_formatData)
  .then(_writeDateToExcel)
  .catch((err) => console.log(err))