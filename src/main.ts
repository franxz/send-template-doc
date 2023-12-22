import fs from "fs";
import path from "path";
const dotenv = require("dotenv");
const nodemailer = require("nodemailer");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const modifiedDocPath = "../temp/modified_document.docx";

type EmailAttachment = {
  filename: string;
  path: string;
  encoding: "base64";
};

(function main() {
  dotenv.config();

  if (!process.env.TEMPLATE_FILENAME || !process.env.INVOICE_NUMBER)
    throw new Error("Missing TEMPLATE_FILENAME or INVOICE_NUMBER env vars");

  replacePlaceholdersInDocx(process.env.TEMPLATE_FILENAME, {
    INVOICE_NUMBER: process.env.INVOICE_NUMBER,
    INVOICE_DATE: new Date().toLocaleDateString("es"),
  });

  sendEmail();
})();

function replacePlaceholdersInDocx(
  templatePath: string,
  replacements: Record<string, string>
) {
  const content = fs.readFileSync(
    path.resolve(__dirname, templatePath),
    "binary"
  );

  // Unzip the content of the file
  const zip = new PizZip(content);

  // This will parse the template, and will throw an error if the template is
  // invalid, for example, if the template is "{user" (no closing tag)
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
  doc.render(replacements);

  // Get the zip document and generate it as a nodebuffer
  const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
  });

  // buf is a nodejs Buffer, you can either write it to a
  // file or res.send it with express for example.
  fs.writeFileSync(path.resolve(__dirname, modifiedDocPath), buf);
}

function sendEmail() {
  const user = process.env.USER_EMAIL;

  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user,
      pass: process.env.USER_PASS,
    },
  });

  // Add the modified doc file as an attachment
  const attachments: EmailAttachment[] = [
    {
      filename: `Invoice_FM_${new Date()
        .toLocaleDateString("en")
        .replace(/\//g, "-")}.docx`,
      path: path.resolve(__dirname, modifiedDocPath),
      encoding: "base64",
    },
  ];

  const mailOptions = {
    from: user,
    to: process.env.TARGET_EMAIL,
    subject: process.env.SUBJECT,
    text: process.env.TEXT,
    attachments,
  };

  // Send the email
  transporter.sendMail(mailOptions, (error: Error, info: any) => {
    fs.rmSync(path.resolve(__dirname, modifiedDocPath));
    if (error) {
      console.error(error);
    } else {
      console.log("Email sent: " + info.response);
    }
  });
}
