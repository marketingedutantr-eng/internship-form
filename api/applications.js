import { google } from "googleapis";
import { randomUUID } from "node:crypto";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const d = req.body || {};
    const requiredFields = [
      "name", "phone", "email", "college", "department",
      "year", "domain", "startDate", "motivation"
    ];

    for (const field of requiredFields) {
      if (typeof d[field] !== "string" || !d[field].trim()) {
        return res.status(400).json({ error: `Please provide ${field}.` });
      }
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) {
      return res.status(400).json({ error: "Please enter a valid email." });
    }
    if (d.phone.replace(/\D/g, "").length < 7) {
      return res.status(400).json({ error: "Please enter a valid phone number." });
    }
    if (d.motivation.trim().length < 15) {
      return res.status(400).json({ error: "Please provide at least 15 characters about your motivation." });
    }

    const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"]
    });
    const sheets = google.sheets({ version: "v4", auth });
    const spreadsheetId = process.env.GOOGLE_SHEET_ID;
    if (!spreadsheetId) throw new Error("Google Sheet ID is missing.");

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
    const tabExists = spreadsheet.data.sheets?.some(
      sheet => sheet.properties.title === "Applications"
    );

    if (!tabExists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ addSheet: { properties: { title: "Applications" } } }]
        }
      });
    }

    const headers = [
      "Application ID", "Submitted At", "Name", "WhatsApp Number",
      "Email", "College", "College Mail ID", "Department", "Year",
      "Interested Domain", "Preferred Start Date", "Why This Internship"
    ];

    const existing = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: "Applications!A1:L1"
    });

    if (!existing.data.values?.length) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: "Applications!A1:L1",
        valueInputOption: "RAW",
        requestBody: { values: [headers] }
      });
    }

    const applicationId =
      "ET-" + new Date().getFullYear() + "-" +
      randomUUID().slice(0, 8).toUpperCase();

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Applications!A:L",
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          applicationId,
          new Date().toISOString(),
          d.name.trim(),
          d.phone.trim(),
          d.email.trim(),
          d.college.trim(),
          (d.collegeEmail || "").trim(),
          d.department.trim(),
          d.year.trim(),
          d.domain.trim(),
          d.startDate.trim(),
          d.motivation.trim()
        ]]
      }
    });

    return res.status(200).json({ success: true, applicationId });
  } catch (error) {
    console.error("Application error:", error);
    return res.status(500).json({
      error: "Unable to save your application. Please try again."
    });
  }
}
