const express = require("express");
const nodemailer = require("nodemailer");
const xlsx = require("xlsx");
const dotenv = require("dotenv");
dotenv.config();

const app = express();
const PORT = 5000;

// Load student data from Excel
const workbook = xlsx.readFile("studentData.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const students = xlsx.utils.sheet_to_json(worksheet);

// Nodemailer configuration
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

// Send confirmation emails
async function sendEmails() {
  for (const student of students) {
    const mailOptions = {
      from: `"ABACUS CLUB" <${process.env.EMAIL_USER}>`,
      to: student.Email,
      subject: "📚 Day 3: Let’s Proceed Further - Open Source Workshop",
      html: `
        <div style="font-family: 'Segoe UI', sans-serif; color: #333; padding: 20px; background-color: #ffffff;">
          <p style="font-size: 18px; color: #1a73e8;"><strong>Hi ${student.Name},</strong></p>

          <p style="font-size: 16px;">
            👋 Welcome to <strong>Day 3</strong> of the <strong>Open Source Contribution Workshop</strong> organized by the ABACUS Club!
          </p>

          <p style="font-size: 16px; color: #444;">
            Here's a quick recap of what we've covered so far:
          </p>

          <ul style="font-size: 15px; margin-left: 20px; color: #444;">
            <li><strong>Day 1:</strong> Introduction to Git & GitHub – how to initialize a repo, commit changes, and push code.</li>
            <li><strong>Day 2:</strong> Explored forking, cloning, and how to create issues on GitHub.</li>
          </ul>

          <p style="font-size: 16px; color: #444;">
            ✅ Today, we will <strong>proceed further</strong> on this exciting journey into open source collaboration.
          </p>

          <p style="font-size: 15px; margin-left: 20px;">
            <strong>🧑‍🏫 Mentor:</strong> Shivam Kumar (CSE Cyber Security, 2022–26)<br/>
            <strong>🕒 Time:</strong> 2:00 PM – 5:00 PM<br/>
            <strong>📍 Venue:</strong> Room No. 60, Academic Block
          </p>

          <p style="font-size: 16px; color: #d93025;">
            💻 Please carry your <strong>laptop</strong> — it’s essential for hands-on participation.
          </p>

          <p style="font-size: 16px;">
            Let’s keep learning, building, and contributing — together! 🚀
          </p>

          <p style="font-size: 15px;">
            Need assistance? Contact us: <strong style="color: #1a73e8;">+91 88252 91561</strong>
          </p>

          <p style="margin-top: 30px;">
            Warm regards,<br/>
            <strong>Team ABACUS</strong><br/>
            Government Engineering College, West Champaran
          </p>

          <p style="font-size: 12px; color: #888;">
            (This is an automated email. Please do not reply directly.)
          </p>
        </div>
      `,
    };

    try {
      const info = await transporter.sendMail(mailOptions);
      console.log(`✅ Email sent to ${student.Email}: ${info.messageId}`);
    } catch (error) {
      console.error(`❌ Failed to send to ${student.Email}: ${error.message}`);
    }
  }
}

// Start server and send emails
app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  sendEmails();
});
