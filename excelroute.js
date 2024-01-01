const express = require("express");
//const { Email } = require("read-excel-file/node");
const { email } = require("../excelmail/excelcontroller");


//const { Email } = require("read-excel-file");
const router = express.Router();

router.get("/email", email);

//app.get('/email', (req, res) => {
//    res.send('This is the /email route!');
//  });
  

module.exports = router;


