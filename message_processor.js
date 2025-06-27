const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { ImapFlow } = require('imapflow');
const nodemailer = require('nodemailer');
const express = require('express');

const USER_SETTINGS_PATH = path.join(__dirname, 'user_settings.json');

function loadSettings() {
  try {
    const data = fs.readFileSync(USER_SETTINGS_PATH, 'utf8');
    return JSON.parse(data);
  } catch (err) {
    console.error('Unable to load user settings:', err);
    process.exit(1);
  }
}

function saveSettings(settings) {
  try {
    fs.writeFileSync(USER_SETTINGS_PATH, JSON.stringify(settings, null, 4));
  } catch (err) {
    console.error('Unable to save user settings:', err);
  }
}

async function promptSettings(settings) {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  function ask(question, defaultValue) {
    return new Promise(resolve => {
      rl.question(`${question} (${defaultValue}): `, answer => {
        resolve(answer || defaultValue);
      });
    });
  }

  settings.SendFrom = await ask('Enter the email address you want to send from', settings.SendFrom);
  settings.SendTo = await ask('Enter the email address you want to send to', settings.SendTo);
  settings.Smtp = await ask('Enter the SMTP server to use', settings.Smtp);
  settings.Encryption = await ask('Select the encryption type: Unencrypted or Encrypted', settings.Encryption);
  settings.Interval = parseInt(await ask('Enter the interval (minutes) between scans', settings.Interval), 10);
  rl.close();
  saveSettings(settings);
  return settings;
}

async function startProcessing(settings) {
  const imap = new ImapFlow({
    host: settings.Smtp,
    port: settings.Encryption === 'Encrypted' ? 993 : 143,
    secure: settings.Encryption === 'Encrypted',
    auth: {
      user: settings.SendFrom,
      pass: settings.RI
    }
  });

  const transporter = nodemailer.createTransport({
    host: settings.Smtp,
    secure: settings.Encryption === 'Encrypted',
    auth: {
      user: settings.SendFrom,
      pass: settings.RI
    }
  });

  await imap.connect();
  await imap.mailboxOpen('INBOX');

  async function scanMailbox() {
    const messages = await imap.search({ seen: false });
    for await (let msg of imap.fetch(messages, { source: true })) {
      const mailOptions = {
        from: settings.SendFrom,
        to: settings.SendTo,
        subject: msg.envelope.subject,
        text: `HEADER\n\n${msg.source.toString()}\n\nFOOTER`
      };
      await transporter.sendMail(mailOptions);
      await imap.messageFlagsAdd(msg.uid, ['\\Seen']);
    }
  }

  // initial scan then repeat at interval without closing the connection
  await scanMailbox();
  setInterval(() => {
    scanMailbox().catch(err => console.error('Failed to process emails:', err));
  }, settings.Interval * 60000);
}

async function main() {
  let settings = loadSettings();
  settings = await promptSettings(settings);

  const app = express();
  app.get('/status', (req, res) => res.send('running'));

  await startProcessing(settings);
  app.listen(3000, () => {
    console.log('Message processor server running on port 3000');
  });
}

main();
