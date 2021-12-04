var nodemailer = require('nodemailer');
var XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');  // name of excel sheet

var i=1;
var count=0;

for(i;i<=1000;i++){

  var first_sheet_name = workbook.SheetNames[0];
  var address_of_cell = `A${i}`; // specify the column in excel in which there are emails stored
  
  var worksheet = workbook.Sheets[first_sheet_name];
  var desired_cell = worksheet[address_of_cell];
  var desired_value = (desired_cell ? desired_cell.v : undefined)
  //email is stored in this desired_value variable

  if(desired_value!==undefined){

      c++; //count variable

      //mail module
      // fill all 5 fields
      let transporter = nodemailer.createTransport({
        service:'gmail',
        auth: {
          type: 'OAuth2',
          user: '',  
          pass: '',  
          clientId: '',
          clientSecret: '',
          refreshToken: ''
        },
      });
    
    // fill 1 field
    var mailOptions = {
      from: ` ` ,
      to:desired_value,

      //subject here
      subject: 'Sending Email using Node.js',

      //text body here
      text: 
`start..


end..
`,

      // add file name and address
      attachments: [{
        filename: '',
        path: '',
        contentType: 'application/pdf'
      },{
        filename: '',
        path: '',
        contentType: 'application/pdf'
      }],
    };
    
    transporter.sendMail(mailOptions, function(error, info){
      if (error) {
        console.log(error);
      } else {
        console.log('Email sent: ' + info.response);
      }
    });
  } 
}

console.log(c);
