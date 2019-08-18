// import default values
const config = require("./config.json")

// require node modules
const nodemailer = require("nodemailer");

const Excel = require('exceljs/modern.nodejs');

// workbook for reading from excel sheet
const workbook = new Excel.Workbook();

const filename = "student_list.xlsx"

// define transporter for sending email. GMAIL API
const transporter = nodemailer.createTransport({

    service: "gmail",

    auth: {
        type: "OAuth2",
        user: config.email,
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        refreshToken: config.refreshToken,
        accessToken: config.accessToken
    }

});

// read actual file
workbook.xlsx.readFile(filename)

    .then(() => {
        // use workbook
        const worksheet = workbook.getWorksheet('Sheet1')

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {

            if (row.values.length < 10) return // if the work book does not contain recommendation break the loop

            const name = row.values[1].split(',').join(' ') // join the names because they come separated by commas

            const email = row.values[2]

            const ipscoresOne = row.values[3]

            const ipscoresTwo = row.values[4]

            const ipscoresThree = row.values[5]

            const ipscoresFour = row.values[6]

            const attendanceScore = row.values[7]

            const recommendation = row.values[10] == "No" ? 'not qualified' : 'qualified'

            const reason = row.values[11]

            //console.log([name,email,recommendation,reason, ipscoresOne, ipscoresTwo, ipscoresThree, ipscoresFour, attendanceScore])

            // call send email function
            sendEmail(email, name, recommendation, ipscoresOne, ipscoresTwo, ipscoresThree, ipscoresFour, attendanceScore)

        });
    });


// function to send email
const sendEmail = (email, name, recommendation, ipscoresOne, ipscoresTwo, ipscoresThree, ipscoresFour, attendanceScore) => {
    const mailOptions = {
        from: `Moringa School <${config.email}>`, // sender email
        to: email,
        subject: `Status for ${name}`,
        html:
            `<p>Hello ${name}. <br><br>
        Below are your marks:
        <ul><li><strong>IP1/31:</strong>${ipscoresOne}</li> <li><strong>IP2/22:</strong>${ipscoresTwo}</li> 
        <li><strong>IP3/22:</strong>${ipscoresThree}</li> <li><strong>IP4/28:</strong>${ipscoresFour}</li> 
        <li><strong>Attendance:</strong> ${attendanceScore * 100}%</li></ul> <br> 
        You have ${recommendation}. <br><br> Thank you for choosing Moringa school. <br><br> Kind regards, <br> Moringa School</p>`
    }
    // sending actual email
    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.log(error)
        } else {
            console.log(info)
        }
    })
}