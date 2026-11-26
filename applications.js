
import { google } from "googleapis";
import { randomUUID } from "node:crypto";
import { Readable } from "node:stream";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const d = req.body || {};

    // Required fields. Motivation and resume are optional.
    const requiredFields = [
      "name", "phone", "email", "college",
      "department", "year", "domain", "startDate"
    ];

    for (const field of requiredFields) {
      if (typeof d[field] !== "string" || !d[field].trim()) {
        return res.status(400).json({
          error: `Please provide ${field}.`
        });
      }
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) {
      return res.status(400).json({
        error: "Please enter a valid email."
      });
    }

    if (d.phone.replace(/\D/g, "").length < 7) {
      return res.status(400).json({
        error: "Please enter a valid phone number."
      });
    }

    // Preferred Start Month must be in YYYY-MM format.
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(d.startDate)) {
      return res.status(400).json({
        error: "Please select a valid preferred start month."
      });
    }

    const motivation =
      typeof d.motivation === "string"
        ? d.motivation.trim()
        : "";

    if (motivation.length > 3000) {
      return res.status(400).json({
        error: "Motivation answer is too long."
      });
    }

    // Connect to Google services.
    const credentials = JSON.parse(
      process.env.GOOGLE_SERVICE_ACCOUNT_JSON
    );

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
      ]
    });

    const sheets = google.sheets({ version: "v4", auth });
    const drive = google.drive({ version: "v3", auth });

    const spreadsheetId = process.env.GOOGLE_SHEET_ID;

    if (!spreadsheetId) {
      throw new Error("Google Sheet ID is missing.");
    }

    // Upload resume to Google Drive if one was provided.
    let resumeLink = "";

    if (d.resume) {
      const resume = d.resume;

      if (
        typeof resume.fileName !== "string" ||
        typeof resume.mimeType !== "string" ||
        typeof resume.dataBase64 !== "string"
      ) {
        return res.status(400).json({
          error: "Invalid resume file."
        });
      }

      const allowedFiles = {
        ".pdf": "application/pdf",
        ".doc": "application/msword",
        ".docx":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      };

      const extension = resume.fileName
        .toLowerCase()
        .match(/\.(pdf|doc|docx)$/)?.[0];

      if (!extension || allowedFiles[extension] !== resume.mimeType) {
        return res.status(400).json({
          error: "Only PDF, DOC and DOCX resumes are accepted."
        });
      }

      const fileBuffer = Buffer.from(resume.dataBase64, "base64");

      // 3 MB limit to keep the request within serverless body limits.
      if (
        fileBuffer.length === 0 ||
        fileBuffer.length > 3 * 1024 * 1024 ||
        fileBuffer.toString("base64") !== resume.dataBase64
      ) {
        return res.status(400).json({
          error: "Resume must be a valid file smaller than 3 MB."
        });
      }

      const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;

      if (!folderId) {
        throw new Error("Google Drive folder ID is missing.");
      }

      const safeName = resume.fileName
        .replace(/[^a-zA-Z0-9._-]/g, "_")
        .slice(0, 150);

      const uploadedFile = await drive.files.create({
        requestBody: {
          name: `${Date.now()}_${safeName}`,
          mimeType: resume.mimeType,
          parents: [folderId]
        },
        media: {
          mimeType: resume.mimeType,
          body: Readable.from(fileBuffer)
        },
        fields: "id, webViewLink",
        supportsAllDrives: true
      });

      resumeLink =
        uploadedFile.data.webViewLink ||
        `https://drive.google.com/file/d/${uploadedFile.data.id}/view`;
    }

    // Check that the Applications sheet exists.
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId
    });

    const tabExists = spreadsheet.data.sheets?.some(
      sheet => sheet.properties.title === "Applications"
    );

    if (!tabExists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: { title: "Applications" }
              }
            }
          ]
        }
      });
    }

    // Updated headers.
    const headers = [
      "Application ID",
      "Submitted At",
      "Name",
      "WhatsApp Number",
      "Email",
      "College",
      "College Mail ID",
      "Department",
      "Year",
      "Interested Domain",
      "Preferred Start Month",
      "Why This Internship",
      "Resume Link"
    ];

    // Update header row without deleting existing applications.
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: "Applications!A1:M1",
      valueInputOption: "RAW",
      requestBody: {
        values: [headers]
      }
    });

    const applicationId =
      "ET-" +
      new Date().getFullYear() +
      "-" +
      randomUUID().slice(0, 8).toUpperCase();

    // Save application in Google Sheets.
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: "Applications!A:M",
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
          motivation,
          resumeLink
        ]]
      }
    });

    return res.status(200).json({
      success: true,
      applicationId
    });

  } catch (error) {
    console.error("Application error:", error);

    return res.status(500).json({
      error: "Unable to save your application. Please try again."
    });
  }
}
