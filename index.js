const express = require('express');
const app = express();
const dotenv = require("dotenv").config();
const port = process.env.PORT || 8000;
const readXlsxFile = require('read-excel-file/node')
const excelcontroller = require("../excelmail/excelcontroller");
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const fs = require('fs');

app.use(express.json())
app.use("/", require("../excelmail/excelroute"));


let transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: "tanyavashistha11@gmail.com" ,
      pass: "nhynqpdnhdupjcot" ,
      
    }
  });

  let mailOptions = {
    from: "tanyavashistha11@gmail.com",
    to: "tanyav1103@gmail.com",
    subject: 'excelmailer',
    text: 'Hi from your nodemailer project'
  };

  transporter.sendMail(mailOptions, function(err, data) {
    if (err) {
      console.log("Error " + err);
    } else {
      console.log("Email sent successfully");
    }
  });

const fileContents = fs.readFileSync('./input.xlsx');
const wb = xlsx.read(fileContents);
console.log(wb.SheetNames);





app.listen ( port, ()=> {
    console.log(`server is running on port ${port}`);
});

