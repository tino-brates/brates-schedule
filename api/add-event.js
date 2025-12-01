// api/add-event.js
import { google } from "googleapis";

function getQuarterSheetInfo(dateStr) {
  if (!dateStr) throw new Error("date manquante");
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) throw new Error("date invalide");

  const year = d.getFullYear();
  const month = d.getMonth();
  const q = Math.floor(month / 3) + 1;

  const sheetTitle = `Q${q} ${year}`;

  const monthNames = [
    "January","February","March",
    "April","May","June",
    "July","August","September",
    "October","November","December"
  ];
  const monthName = monthNames[month];

  return { sheetTitle, monthName, year };
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Méthode non autorisée" });
    return;
  }

  try {
    const {
      date,
      time,
      event,
      location,
      provider,
      camera,
      delivery,
      internet,
      client,
      crewTC,
      crewDG,
      crewJB
    } = req.body || {};

    if (!date || !event) {
      res.status(400).json({ error: "date et event obligatoires" });
      return;
    }

    const {
      GOOGLE_SERVICE_ACCOUNT_EMAIL,
      GOOGLE_PRIVATE_KEY,
      GOOGLE_SHEET_ID
    } = process.env;

    if (!GOOGLE_SERVICE_ACCOUNT_EMAIL || !GOOGLE_PRIVATE_KEY || !GOOGLE_SHEET_ID) {
      res.status(500).json({ error: "Config env Google manquante" });
      return;
    }

    const jwtClient = new google.auth.JWT(
      GOOGLE_SERVICE_ACCOUNT_EMAIL,
      null,
      GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
      ["https://www.googleapis.com/auth/spreadsheets"]
    );

    await jwtClient.authorize();

    const sheets = google.sheets({ version: "v4", auth: jwtClient });

    const { sheetTitle, monthName } = getQuarterSheetInfo(date);

    const range = `'${sheetTitle}'!A:Z`;
    const getResp = await sheets.spreadsheets.values.get({
      spreadsheetId: GOOGLE_SHEET_ID,
      range
    });

    const rows = getResp.data.values || [];

    let monthRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i] && rows[i][0] && rows[i][0].toLowerCase() === monthName.toLowerCase()) {
        monthRowIndex = i;
        break;
      }
    }

    if (monthRowIndex === -1) {
      res.status(500).json({ error: `Ligne du mois ${monthName} introuvable dans ${sheetTitle}` });
      return;
    }

    let insertRowIndex = monthRowIndex + 1;
    while (
      insertRowIndex < rows.length &&
      rows[insertRowIndex] &&
      rows[insertRowIndex][0] &&
      rows[insertRowIndex][0].trim() !== "" &&
      !isNaN(Date.parse(rows[insertRowIndex][0]))
    ) {
      insertRowIndex++;
    }

    const excelDate = (() => {
      const d = new Date(date + "T00:00:00");
      const epoch = new Date(Date.UTC(1899, 11, 30));
      return (d - epoch) / (1000 * 60 * 60 * 24);
    })();

    const crewTCVal = crewTC ? "x" : "";
    const crewDGVal = crewDG ? "x" : "";
    const crewJBVal = crewJB ? "x" : "";

    const rowValues = [
      excelDate,
      "",
      time || "",
      crewTCVal,
      crewDGVal,
      crewJBVal,
      event || "",
      location || "",
      provider || "",
      camera || "",
      delivery || "",
      internet || "",
      client || "",
      ""
    ];

    const targetRange = `'${sheetTitle}'!A${insertRowIndex + 1}:N${insertRowIndex + 1}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SHEET_ID,
      range: targetRange,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [rowValues]
      }
    });

    res.status(200).json({ ok: true });
  } catch (err) {
    console.error("Erreur add-event:", err);
    res.status(500).json({ error: err.message || "Erreur interne" });
  }
}
