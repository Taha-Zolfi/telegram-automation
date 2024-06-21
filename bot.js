const TelegramBot = require('node-telegram-bot-api');
const XLSX = require('xlsx');
const sqlite3 = require('sqlite3').verbose();
const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const { TelegramClient } = require('telegram');
const { StringSession } = require('telegram/sessions');
const { Api } = require('telegram/tl');
const fs = require('fs');
const path = require('path');

// Telegram Bot Token
const token = '7045174404:AAHDO2G0TnJ-6wbxMdcI-nFoTyhowSeZcK8';
const bot = new TelegramBot(token, { polling: true });

// Excel File Path
const filePath = 'public/data.xlsx';

// Reading the Excel file
const workbook = XLSX.readFile(filePath);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Connecting to the SQLite database
const db = new sqlite3.Database('./public/users.db');

// Create users table if it doesn't exist
db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS users (chat_id TEXT UNIQUE)`);
});

// Function to add data to Excel
function addToExcel(data) {
    const newRow = [
        XLSX.utils.decode_range(worksheet['!ref']).e.r + 1, // Row number
        data.firstName,
        data.lastName,
        data.grade,
        data.gender,
        data.province,
        data.region,
        data.mobile
    ];
    const newRowIndex = XLSX.utils.decode_range(worksheet['!ref']).e.r + 1;

    newRow.forEach((value, index) => {
        const cellRef = XLSX.utils.encode_cell({ c: index, r: newRowIndex });
        worksheet[cellRef] = { t: 's', v: value };
    });

    worksheet['!ref'] = XLSX.utils.encode_range(worksheet['!ref'].split(':')[0], XLSX.utils.encode_cell({ c: newRow.length - 1, r: newRowIndex }));

    XLSX.writeFile(workbook, filePath);
}

// Function to check if a user is registered
function isUserRegistered(chatId) {
    return new Promise((resolve, reject) => {
        db.get("SELECT chat_id FROM users WHERE chat_id = ?", [chatId], (err, row) => {
            if (err) {
                reject(err);
            } else {
                resolve(!!row);
            }
        });
    });
}

// Function to add a user to the database
function addUserToDatabase(chatId) {
    return new Promise((resolve, reject) => {
        db.run("INSERT INTO users (chat_id) VALUES (?)", [chatId], function(err) {
            if (err) {
                reject(err);
            } else {
                resolve();
            }
        });
    });
}

// Variable to store user information
let userStates = {};

// Telegram Bot Handlers
bot.onText(/\/start/, async (msg) => {
    const chatId = msg.chat.id;
    try {
        const registered = await isUserRegistered(chatId);
        if (registered) {
            bot.sendMessage(chatId, 'You have already registered.');
            return;
        }
        userStates[chatId] = { step: 'firstName' };
        bot.sendMessage(chatId, `Welcome! Please enter your first name.`);
    } catch (err) {
        console.error(err);
        bot.sendMessage(chatId, 'An error occurred. Please try again later.');
    }
});

bot.on('message', async (msg) => {
    const chatId = msg.chat.id;

    if (!userStates[chatId]) return;

    const state = userStates[chatId];

    switch (state.step) {
        case 'firstName':
            state.firstName = msg.text;
            state.step = 'lastName';
            bot.sendMessage(chatId, 'Please enter your last name.');
            break;
        case 'lastName':
            state.lastName = msg.text;
            state.step = 'grade';
            bot.sendMessage(chatId, 'Please select your grade.', {
                reply_markup: {
                    keyboard: [['1st', '2nd', '3rd', '4th'], ['5th', '6th', '7th', '8th'], ['9th', '10th', '11th', '12th']],
                    one_time_keyboard: true
                }
            });
            break;
        case 'grade':
            state.grade = msg.text;
            state.step = 'gender';
            bot.sendMessage(chatId, 'Please select your gender.', {
                reply_markup: {
                    keyboard: [['Male', 'Female']],
                    one_time_keyboard: true
                }
            });
            break;
        case 'gender':
            state.gender = msg.text;
            state.step = 'province';
            bot.sendMessage(chatId, 'Please enter your province.');
            break;
        case 'province':
            state.province = msg.text;
            state.step = 'region';
            bot.sendMessage(chatId, 'Please enter your region.');
            break;
        case 'region':
            state.region = msg.text;
            state.step = 'mobile';
            bot.sendMessage(chatId, 'Please enter your mobile number.');
            break;
        case 'mobile':
            state.mobile = msg.text;

            try {
                await addUserToDatabase(chatId);
                addToExcel(state);
                bot.sendMessage(chatId, `Congratulations! You have been registered. Your row number is: ${XLSX.utils.decode_range(worksheet['!ref']).e.r}`);
                delete userStates[chatId];
            } catch (err) {
                console.error(err);
                bot.sendMessage(chatId, 'An error occurred. Please try again later.');
            }
            break;
    }
});

// Express Server for Sending Messages
const app = express();
const port = 3000;

// Your Telegram API credentials
const API_ID = 23787541; 
const API_HASH = "fc1f17f7d2e81b0ad904228f002c01c9";
const DATA_FILE_PATH = "public/data.xlsx";

let client;
let phoneNumber; 
let authCodeHash; 
const stringSession = new StringSession("");

// Set up multer for handling file uploads
const upload = multer({ dest: "public/" });

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static("public"));

// Load phone numbers from Excel file
function loadPhoneNumbers(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet);
  const phoneNumbers = data.map(entry => String(entry.mobile)); 
  return phoneNumbers;
}

// Serve the phone number form
app.get("/", (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html lang="fa">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Enter Phone Number</title>
      <link rel="stylesheet" href="/styles.css">
    </head>
    <body>
      <div class="container">
        <h2>Enter Phone Number</h2>
        <form action="/sendCode" method="post">
          <label for="phoneNumber">Phone Number:</label>
          <input type="text" id="phoneNumber" name="phoneNumber" required><br><br>
          <input type="submit" value="Send Code">
        </form>
      </div>
    </body>
    </html>
  `);
});

// Handle phone number submission and send the code
app.post("/sendCode", async (req, res) => {
  phoneNumber = req.body.phoneNumber;

  client = new TelegramClient(stringSession, API_ID, API_HASH, {
    connectionRetries: 5,
  });

  try {
    await client.connect();
    const result = await client.invoke(
      new Api.auth.SendCode({
        phoneNumber: phoneNumber,
        apiId: API_ID,
        apiHash: API_HASH,
        settings: new Api.CodeSettings({}),
      })
    );

    authCodeHash = result.phoneCodeHash;

    res.send(`
      <!DOCTYPE html>
      <html lang="fa">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Enter Authentication Code</title>
        <link rel="stylesheet" href="/styles.css">
      </head>
      <body>
        <div class="container">
          <h2>Enter Authentication Code</h2>
          <form action="/authenticate" method="post">
            <label for="authCode">Authentication Code:</label>
            <input type="text" id="authCode" name="authCode" required><br><br>
            <input type="submit" value="Submit">
          </form>
        </div>
      </body>
      </html>
    `);
  } catch (err) {
    console.error("Failed to send authentication code:", err);
    res.status(500).send("Failed to send authentication code.");
  }
});

// Handle authentication code submission and then prompt for the message and file/image path
app.post("/authenticate", async (req, res) => {
  const authCode = req.body.authCode;

  try {
    await client.invoke(
      new Api.auth.SignIn({
        phoneNumber: phoneNumber,
        phoneCodeHash: authCodeHash,
        phoneCode: authCode,
      })
    );

    res.send(`
      <!DOCTYPE html>
      <html lang="fa">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Send Message</title>
        <link rel="stylesheet" href="/styles.css">
      </head>
      <body>
        <div class="container">
          <h2>Send Message</h2>
          <form action="/sendMessage" method="post" enctype="multipart/form-data">
            <label for="message">Message:</label>
            <input type="text" id="message" name="message" required><br><br>
            <label for="imagePath">Image:</label>
            <input type="file" id="imagePath" name="imagePath" accept="image/*"><br><br>
            <label for="filePath">File:</label>
            <input type="file" id="filePath" name="filePath"><br><br>
            <input type="submit" value="Send">
          </form>
        </div>
      </body>
      </html>
    `);
  } catch (err) {
    console.error("Failed to authenticate:", err);
    res.status(500).send("Failed to authenticate.");
  }
});

// Handle message submission and send the message
app.post("/sendMessage", upload.fields([{ name: 'imagePath', maxCount: 1 }, { name: 'filePath', maxCount: 1 }]), async (req, res) => {
  const MESSAGE = req.body.message;
  const IMAGE_PATH = req.files['imagePath'] ? req.files['imagePath'][0].path : null;
  const FILE_PATH = req.files['filePath'] ? req.files['filePath'][0].path : null;

  const TARGET_PHONE_NUMBERS = loadPhoneNumbers(DATA_FILE_PATH);

  let logs = "";

  try {
    for (let i = 0; i < TARGET_PHONE_NUMBERS.length; i++) {
      const phone = TARGET_PHONE_NUMBERS[i];

      const contact = new Api.InputPhoneContact({
        clientId: BigInt(i),
        phone: phone,
        firstName: `User${i}`,
        lastName: "",
      });

      const result = await client.invoke(
        new Api.contacts.ImportContacts({ contacts: [contact] })
      );

      if (result.imported.length > 0) {
        const user = result.imported[0];

        try {
          if (IMAGE_PATH) {
            const file = await client.uploadFile({
              file: IMAGE_PATH,
              workers: 1,
            });

            await client.invoke(
              new Api.messages.SendMedia({
                peer: user.userId,
                media: new Api.InputMediaUploadedPhoto({
                  file: file,
                  caption: MESSAGE,
                }),
                message: MESSAGE,
                randomId: BigInt(Math.floor(Math.random() * 0xFFFFFFFFFFFFFFFF))
              })
            );

            const log = `Image with caption sent to ${user.userId}`;
            logs += `${log}\n`;
            console.log(log);
          } else if (FILE_PATH) {
            await client.sendFile(user.userId, {
              file: FILE_PATH,
              caption: MESSAGE,
              forceDocument: true
            });
            const log = `File with caption sent to ${user.userId}`;
            logs += `${log}\n`;
            console.log(log);
          }
        } catch (err) {
          const log = `Failed to send file to ${user.userId}: ${err}`;
          logs += `${log}\n`;
          console.error(log);
        }
      } else {
        const log = `Failed to import contact ${phone}`;
        logs += `${log}\n`;
        console.error(log);
      }

      await new Promise((resolve) => setTimeout(resolve, 3000));
    }

    res.send(`
      <!DOCTYPE html>
      <html lang="fa">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Send Results</title>
        <link rel="stylesheet" href="/styles.css">
      </head>
      <body>
        <div class="container">
          <h2>Send Results</h2>
          <pre>${logs}</pre>
        </div>
      </body>
      </html>
    `);
  } catch (err) {
    console.error("Failed to send messages:", err);
    res.status(500).send("Failed to send messages.");
  } finally {
    if (IMAGE_PATH) {
      fs.unlinkSync(IMAGE_PATH);
    }
    if (FILE_PATH) {
      fs.unlinkSync(FILE_PATH);
    }
  }
});

// Serve static files from the public directory
app.use("/public", express.static(path.join(__dirname, "public")));

// Start the server
app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
