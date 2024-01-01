const readXlsxFile = require('read-excel-file/node')

const email = (req, res) => {
    console.log("working");
    readXlsxFile('/home/tanya/C++/excelmail/input.xlsx').then((rows) => {
      res.json(rows);
    });
  };
  //Hi Tanyaaaa!!



module.exports = { email };