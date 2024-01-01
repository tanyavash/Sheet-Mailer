const xlsx = require('xlsx');
const fs = require('fs');
const nodemailer = require('nodemailer');

const fileContents = fs.readFileSync('./input.xlsx');
const wb = xlsx.read(fileContents);
//console.log(wb.SheetNames);
const ws = wb.Sheets["Sheet1"];
//console.log(ws);

const data = xlsx.utils.sheet_to_json(ws);
//console.log(data);

fs.writeFileSync('./jsondata.json', JSON.stringify(data, null, 2));


//const fs = require('fs');

// Replace 'path/to/your/file.json' with the actual path to your JSON file
const filePath = '/home/tanya/C++/Project1/jsondata.json';

// Read the JSON file
fs.readFile(filePath, 'utf8', (err, data) => {
    // console.log(data);
    
  if (err) {
    console.error('Error reading JSON file:', err);
    return;
  }


  try {
    let transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: "tanyavashistha11@gmail.com" ,
        pass: "nhynqpdnhdupjcot" ,
        
      }
    });

    // Parse the JSON data
    const jsonData = JSON.parse(data);
    jsonData.forEach((element) => {
        let mailOptions = {
            from: "tanyavashistha11@gmail.com",
            to: element["EMAIL"],
            subject: 'excelmailer',
            text: `Hi ${element["NAME"]} from your Tanya your id is ${element["ID"]}`
          };
          transporter.sendMail(mailOptions, function(err, data) {
            if (err) {
              console.log("Error " + err);
            } else {
              console.log("Email sent successfully");
            }
         });
          
    });
    // Extract email values for all objects
    // const emailValues = jsonData.map(obj => obj.EMAIL);

    // Filter out undefined values (in case "EMAIL" is missing in some objects)
    // const validEmailValues = emailValues.filter(email => email !== undefined);

    // Check if any valid email values exist
//if (validEmailValues.length > 0) {
//
   
//    
//      
       
//
//
//
//  // Use the email values as needed
//  console.log('Email Values:', validEmailValues);
//} else {
//  console.log('No valid "EMAIL" values found in the JSON data.');
// }
  }

catch (parseError) {
console.error('Error parsing JSON:', parseError);
}});
//
//
//