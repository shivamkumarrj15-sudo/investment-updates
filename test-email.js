// Quick email test - run: node test-email.js
require('dotenv').config();
const nodemailer = require('nodemailer');

const sender   = process.env.SENDER_EMAIL;
const password = process.env.SENDER_APP_PASSWORD;
const receiver = process.env.RECEIVER_EMAIL;

console.log('🔍 Email Config Check:');
console.log(`   SENDER_EMAIL:        ${sender      ? '✅ Set (' + sender + ')'      : '❌ MISSING'}`);
console.log(`   SENDER_APP_PASSWORD: ${password    ? '✅ Set (****)'                 : '❌ MISSING'}`);
console.log(`   RECEIVER_EMAIL:      ${receiver    ? '✅ Set (' + receiver + ')'     : '❌ MISSING'}`);

if (!sender || !password || !receiver) {
  console.error('\n❌ Missing .env variables. Create a .env file with the above keys.');
  process.exit(1);
}

async function testEmail() {
  console.log('\n📧 Testing SMTP connection (smtp.gmail.com:465)...');

  const transporter = nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 465,
    secure: true,
    auth: { user: sender, pass: password },
    tls: { rejectUnauthorized: false },
    connectionTimeout: 15000,
  });

  try {
    await transporter.verify();
    console.log('✅ SMTP connection OK!');
  } catch (err) {
    console.error('❌ SMTP verify FAILED:', err.message);
    console.error('\n   Possible fixes:');
    console.error('   1. Make sure SENDER_APP_PASSWORD is a Gmail App Password (not your regular Gmail password)');
    console.error('   2. Go to: myaccount.google.com > Security > 2-Step Verification > App Passwords');
    console.error('   3. Generate a new 16-char App Password and update .env / GitHub Secrets');
    process.exit(1);
  }

  console.log('\n📨 Sending test email...');
  try {
    const info = await transporter.sendMail({
      from: `"📈 Test Bot" <${sender}>`,
      to: receiver,
      subject: `✅ Test Email - Investment Tracker Bot Working! (${new Date().toLocaleString('en-IN', {timeZone:'Asia/Kolkata'})})`,
      html: `
        <div style="font-family:sans-serif;padding:20px;max-width:600px;">
          <h2 style="color:#0f172a;">✅ Email Test Successful!</h2>
          <p>Yeh test email Investment Tracker Bot ne bheja hai.</p>
          <p><strong>Time (IST):</strong> ${new Date().toLocaleString('en-IN', {timeZone:'Asia/Kolkata'})}</p>
          <p><strong>From:</strong> ${sender}</p>
          <p><strong>To:</strong> ${receiver}</p>
          <hr/>
          <p style="color:#64748b;font-size:12px;">Agar yeh email aaya, toh SMTP configuration sahi hai. ✅</p>
        </div>
      `
    });
    console.log(`✅ Test email SENT! Message-ID: ${info.messageId}`);
    console.log(`\n🎉 Email system is working correctly!`);
    console.log(`   Check your inbox: ${receiver}`);
  } catch (err) {
    console.error('❌ sendMail FAILED:', err.message);
    process.exit(1);
  }
}

testEmail();
