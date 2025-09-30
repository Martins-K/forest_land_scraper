import dotenv from "dotenv";
dotenv.config();

import nodemailer from "nodemailer";
import { LandForestData } from "./types";

export async function sendEmail(
  newItems: LandForestData[],
  fileName: string,
  cutoffHours: number = 25
): Promise<void> {
  try {
    // Create transporter (configure with your email service)
    const transporter = nodemailer.createTransport({
      service: process.env.EMAIL_SERVICE || "gmail",
      auth: {
        user: process.env.EMAIL_USER_LAND,
        pass: process.env.EMAIL_PASS_LAND,
      },
    });

    // Create HTML table for new items
    const itemsTable =
      newItems.length > 0
        ? `
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 10pt;">
        <thead>
          <tr style="background-color: #f2f2f2;">
            <th style="text-align: left;">Link</th>
            <th style="text-align: left;">Price</th>
            <th style="text-align: left;">District</th>
            <th style="text-align: left;">Area</th>
            <th style="text-align: left;">Cadastre</th>
            <th style="text-align: left;">Date</th>
          </tr>
        </thead>
        <tbody>
          ${newItems
            .map(
              (item) => `
            <tr>
              <td><a href="${
                item.link
              }" style="color: #0066cc; text-decoration: none;">View Listing</a></td>
              <td>${item.price}</td>
              <td>${item.districtText || "N/A"}</td>
              <td>${item.areaText || "N/A"}</td>
              <td>${item.cadastreText || "N/A"}</td>
              <td>${item.date}</td>
            </tr>
          `
            )
            .join("")}
        </tbody>
      </table>
    `
        : '<p style="color: #666; font-style: italic;">No new items found in this run.</p>';

    const emailBody = `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Land Scraper Report</title>
</head>
<body>
    <div style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto;">
        <h2 style="color: #333; border-bottom: 2px solid #0066cc; padding-bottom: 10px;">
            🏞️ Land Scraper Report
        </h2>
        
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-bottom: 20px;">
            <p style="margin: 0; color: #555;">
                <strong>Scraping completed on:</strong> ${new Date().toLocaleString("en-US", {
                  timeZone: "Europe/Riga",
                  year: "numeric",
                  month: "2-digit",
                  day: "2-digit",
                  hour: "2-digit",
                  minute: "2-digit",
                  second: "2-digit",
                  hour12: false,
                })}<br>
                <strong>New land listings found in the past ${cutoffHours} hours:</strong> ${
      newItems.length
    }
            </p>
        </div>

        <h3 style="color: #0066cc;">New Land Listings</h3>
        ${itemsTable}

        <div style="margin-top: 30px; padding: 15px; background-color: #f0f8ff; border-radius: 5px;">
            <p style="margin: 0; color: #555;">
                <strong>Note:</strong> The complete dataset including previous listings (scraped on 9/26/2025, 12:25:16 PM) is attached as an Excel file.
                Some previously found listings may no longer be available (removed in the portal).
                This email contains only the new items discovered in the latest run.
                ${
                  cutoffHours === 73
                    ? "<br><strong>Monday run:</strong> Includes listings from the weekend (past 72 hours)."
                    : ""
                }
            </p>
        </div>

        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd;">
            <p style="color: #666;">
                At your service,<br>
                <strong>Land Scraping Service</strong>
            </p>
        </div>
    </div>
</body>
</html>
    `;

    const mailOptions = {
      from: `Land Scraping Service <${process.env.EMAIL_USER_LAND}>`,
      to: process.env.RECIPIENT_EMAIL_LAND,
      cc: process.env.CC_EMAIL_LAND || "",
      bcc: process.env.BCC_EMAIL_LAND || "",
      subject: `Land Scraper Report - ${new Date().toLocaleDateString()} (${
        newItems.length
      } new items)`,
      html: emailBody,
      attachments: [
        {
          filename: "lands-scraped.xlsx",
          path: fileName,
        },
      ],
    };

    const info = await transporter.sendMail(mailOptions);
    console.log("Email sent successfully:", info.messageId);
  } catch (error) {
    console.error("Error sending email:", error);
    throw error; // Re-throw to handle in the main script if needed
  }
}
