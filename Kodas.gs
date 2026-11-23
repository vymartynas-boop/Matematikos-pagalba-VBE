
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Matematikos pagalba VBE egzaminams")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function askMathAI(message, history) {
  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  if (!apiKey) {
    throw new Error("Nerastas Gemini API raktas. Patikrink Script Properties.");
  }

  if (!message || !message.trim()) {
    throw new Error("Įrašyk matematikos uždavinį.");
  }

  const systemPrompt = `
Tu esi „Matematikos pagalba VBE egzaminams“.
Esi kantrus, profesionalus Lietuvos matematikos mokytojas.

Tavo tikslas – padėti mokiniams pasiruošti matematikos VBE.

TAISYKLĖS:
1. Visada atsakyk lietuviškai.
2. Paaiškink, ką reikia rasti.
3. Pateik reikalingas formules.
4. Sprendimą išskaidyk į sunumeruotus žingsnius.
5. Paaiškink, kodėl atliekamas kiekvienas veiksmas.
6. Pateik aiškų galutinį atsakymą.
7. Jei įmanoma, patikrink rezultatą.
8. Naudok mokyklinio matematikos kurso terminus.
9. Jei mokinys nesupranta, paaiškink paprasčiau.
10. Niekada neišgalvok matematinių taisyklių ar rezultatų.
11. Jei sąlyga neaiški, paprašyk patikslinti.

Elkis kaip matematikos mokytojas, o ne skaičiuotuvas.
`;

  const contents = [];

  // Ankstesni pokalbio pranešimai
  (history || []).slice(-8).forEach(function(item) {
    if (
      item.content &&
      (item.role === "user" || item.role === "assistant")
    ) {
      contents.push({
        role: item.role === "assistant" ? "model" : "user",
        parts: [{ text: String(item.content).slice(0, 5000) }]
      });
    }
  });

  // Dabartinis klausimas
  contents.push({
    role: "user",
    parts: [{ text: message.slice(0, 8000) }]
  });

  const model = "gemini-3.5-flash-lite";

  const url =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model + ":generateContent";

const vbeContext = findRelevantVbeText_(message);

const finalSystemPrompt = systemPrompt +
  (vbeContext
    ? "\n\nOFICIALI VBE MEDŽIAGA:\n" + vbeContext +
      "\n\nAtsakydamas remkis šia medžiaga, kai ji susijusi su klausimu."
    : "\n\nŠiam klausimui tinkamos VBE medžiagos nerasta. Neišgalvok oficialių užduočių ar vertinimo kriterijų.");

const payload = {
  systemInstruction: {
    parts: [{ text: finalSystemPrompt }]
  },
  contents: contents,
  generationConfig: {
    maxOutputTokens: 2500
  }
};

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-goog-api-key": apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const responseText = response.getContentText();

  let result;

  try {
    result = JSON.parse(responseText);
  } catch (e) {
    throw new Error("Nepavyko perskaityti Gemini atsakymo.");
  }

  if (status !== 200) {
    const errorMessage = result.error
      ? result.error.message
      : responseText;

    throw new Error(
      "Gemini API klaida (" + status + "): " + errorMessage
    );
  }

  if (
    !result.candidates ||
    !result.candidates.length ||
    !result.candidates[0].content
  ) {
    throw new Error("Gemini nepateikė atsakymo. Bandyk dar kartą.");
  }

  const parts = result.candidates[0].content.parts || [];

  const answer = parts
    .filter(function(part) {
      return typeof part.text === "string";
    })
    .map(function(part) {
      return part.text;
    })
    .join("\n");

  if (!answer) {
    throw new Error("Gautas tuščias Gemini atsakymas.");
  }

  return answer;
}


const VBE_FOLDER_ID = "1J5M1UyC6Sj2ytuvufa7g-zPvZ5aTPJ7K";

function scanVbeDrive() {
  const root = DriveApp.getFolderById(VBE_FOLDER_ID);

  const spreadsheet = SpreadsheetApp.create(
    "VBE matematikos žinių bazė – dokumentų indeksas"
  );

  const sheet = spreadsheet.getSheets()[0];
  sheet.setName("Dokumentų indeksas");

  const headers = [
    "Metai / aplankas",
    "Dokumento pavadinimas",
    "Dokumento ID",
    "Tipas",
    "Nuoroda"
  ];

  const rows = [headers];

  collectVbeFiles_(root, root.getName(), rows);

  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, headers.length)
      .setValues(rows);

    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#DCE6FA");

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  PropertiesService.getScriptProperties()
    .setProperty("VBE_INDEX_SHEET_ID", spreadsheet.getId());

  Logger.log("Rasta dokumentų: " + (rows.length - 1));
  Logger.log("Žinių bazės indeksas: " + spreadsheet.getUrl());

  return spreadsheet.getUrl();
}

function collectVbeFiles_(folder, folderPath, rows) {
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();

    if (file.getMimeType() === "application/pdf") {
      rows.push([
        folderPath,
        file.getName(),
        file.getId(),
        file.getMimeType(),
        file.getUrl()
      ]);
    }
  }

  const folders = folder.getFolders();

  while (folders.hasNext()) {
    const subfolder = folders.next();

    collectVbeFiles_(
      subfolder,
      folderPath + " / " + subfolder.getName(),
      rows
    );
  }
}

function testPdfExtraction() {
  const root = DriveApp.getFolderById(VBE_FOLDER_ID);
  const pdf = findFirstPdf_(root);

  if (!pdf) {
    throw new Error("PDF dokumentų nerasta.");
  }

  Logger.log("Bandomas dokumentas: " + pdf.getName());

  let tempDocId = null;

  try {
    const converted = Drive.Files.create(
      {
        name: "TEMP - " + pdf.getName(),
        mimeType: "application/vnd.google-apps.document"
      },
      pdf.getBlob(),
      { fields: "id, name" }
    );

    tempDocId = converted.id;

    const text = DocumentApp.openById(tempDocId)
      .getBody()
      .getText();

    Logger.log("Ištraukta simbolių: " + text.length);
    Logger.log("TEKSTO PRADŽIA:\n" + text.substring(0, 1500));

  } finally {
    if (tempDocId) {
      DriveApp.getFileById(tempDocId).setTrashed(true);
    }
  }
}

function findFirstPdf_(folder) {
  const files = folder.getFilesByType(MimeType.PDF);

  if (files.hasNext()) {
    return files.next();
  }

  const folders = folder.getFolders();

  while (folders.hasNext()) {
    const found = findFirstPdf_(folders.next());
    if (found) return found;
  }

  return null;
}
function buildVbeKnowledgeBase() {
  const root = DriveApp.getFolderById(VBE_FOLDER_ID);

  const spreadsheetId = PropertiesService
    .getScriptProperties()
    .getProperty("VBE_INDEX_SHEET_ID");

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  let sheet = spreadsheet.getSheetByName("Žinių bazė");

  if (!sheet) {
    sheet = spreadsheet.insertSheet("Žinių bazė");
  }

  sheet.clearContents();

  sheet.appendRow([
    "Metai / aplankas",
    "Dokumentas",
    "Turinio dalis",
    "Tekstas"
  ]);

  processVbeFolder_(root, root.getName(), sheet);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 4);
PropertiesService.getScriptProperties()
  .setProperty("VBE_KNOWLEDGE_SHEET_ID", spreadsheet.getId());
  Logger.log("Žinių bazė sukurta. Iš viso dalių: " + (sheet.getLastRow() - 1));
}

function processVbeFolder_(folder, folderPath, sheet) {
  const files = folder.getFilesByType(MimeType.PDF);

  while (files.hasNext()) {
    const pdf = files.next();
    let tempDocId = null;

    try {
      Logger.log("Apdorojamas: " + pdf.getName());

      const converted = Drive.Files.create(
        {
          name: "TEMP - " + pdf.getName(),
          mimeType: "application/vnd.google-apps.document"
        },
        pdf.getBlob(),
        { fields: "id" }
      );

      tempDocId = converted.id;

      const content = DocumentApp.openById(tempDocId)
        .getBody()
        .getText();

      const chunkSize = 7000;

      for (let i = 0; i < content.length; i += chunkSize) {
        sheet.appendRow([
          folderPath,
          pdf.getName(),
          Math.floor(i / chunkSize) + 1,
          content.substring(i, i + chunkSize)
        ]);
      }

    } catch (error) {
      Logger.log("KLAIDA su " + pdf.getName() + ": " + error.message);

    } finally {
      if (tempDocId) {
        DriveApp.getFileById(tempDocId).setTrashed(true);
      }
    }
  }

  const folders = folder.getFolders();

  while (folders.hasNext()) {
    const subfolder = folders.next();

    processVbeFolder_(
      subfolder,
      folderPath + " / " + subfolder.getName(),
      sheet
    );
  }
}

function findKnowledgeBase() {
  const id = PropertiesService.getScriptProperties()
    .getProperty("VBE_KNOWLEDGE_SHEET_ID");

  Logger.log("Žinių bazės ID: " + id);

  if (id) {
    Logger.log("Žinių bazės nuoroda: " +
      SpreadsheetApp.openById(id).getUrl());
  }
}
function findRelevantVbeText_(question) {
  const id = PropertiesService.getScriptProperties()
    .getProperty("VBE_KNOWLEDGE_SHEET_ID");

  if (!id) return "";

  const sheet = SpreadsheetApp.openById(id)
    .getSheetByName("Žinių bazė");

  if (!sheet) return "";

  const rows = sheet.getDataRange().getValues();
  const words = question.toLowerCase()
    .replace(/[^\p{L}\p{N}]/gu, " ")
    .split(/\s+/)
    .filter(w => w.length > 2);

  const results = [];

  for (let i = 1; i < rows.length; i++) {
    const folder = String(rows[i][0]);
    const filename = String(rows[i][1]);
    const chunk = String(rows[i][3]);

    if (!chunk) continue;

    const searchable = (folder + " " + filename + " " + chunk)
      .toLowerCase();

    let score = 0;

    words.forEach(word => {
      if (searchable.includes(word)) score++;
    });

    if (score > 0) {
      results.push({
        score: score,
        text: "Dokumentas: " + filename +
          "\nAplankas: " + folder +
          "\nTurinys:\n" + chunk
      });
    }
  }

  results.sort((a, b) => b.score - a.score);

  return results.slice(0, 3)
    .map(item => item.text)
    .join("\n\n---\n\n");
}

function askMathAIWithImage(message, history, mimeType, base64Data) {
  try {
    const apiKey = PropertiesService
      .getScriptProperties()
      .getProperty("GEMINI_API_KEY");

    if (!apiKey) {
      throw new Error("Nerastas Gemini API raktas.");
    }

    // Leidžiami nuotraukų formatai
    const allowedTypes = [
      "image/jpeg",
      "image/png",
      "image/webp"
    ];

    if (!allowedTypes.includes(mimeType)) {
      throw new Error("Galima įkelti tik JPG, PNG arba WEBP nuotraukas.");
    }

    if (!base64Data || base64Data.length > 6 * 1024 * 1024) {
      throw new Error("Nuotrauka per didelė. Pabandyk įkelti mažesnę.");
    }

    const systemPrompt = `
Tu esi matematikos mokytojas, padedantis mokiniams pasiruošti
matematikos valstybiniam brandos egzaminui (VBE).

Tavo užduotis:
- Atidžiai išanalizuoti mokinio įkeltą matematikos uždavinio nuotrauką.
- Perskaityti joje pateiktą sąlygą, formules, skaičius ir simbolius.
- Jei sąlyga neįskaitoma, paprašyti mokinio įkelti aiškesnę nuotrauką.
- Spręsti uždavinį nuosekliai, žingsnis po žingsnio.
- Paaiškinti, kodėl atliekamas kiekvienas veiksmas.
- Naudoti mokyklinę matematikos terminologiją.
- Pritaikyti paaiškinimo sudėtingumą mokinio pasirinktam A arba B kursui.
- Neskubėti pateikti vien tik galutinio atsakymo.
- Jei mokinys pateikia savo sprendimą, padėti rasti klaidą.

Atsakyk lietuviškai, aiškiai ir suprantamai.
`;

    // Įtraukiame ankstesnį pokalbį
    const contents = (history || []).map(item => ({
      role: item.role === "assistant" ? "model" : "user",
      parts: [{ text: String(item.content || "") }]
    }));

    // Pridedame dabartinį klausimą ir nuotrauką
    contents.push({
      role: "user",
      parts: [
        {
          text: message || "Išanalizuok nuotraukoje pateiktą matematikos uždavinį."
        },
        {
          inline_data: {
            mime_type: mimeType,
            data: base64Data
          }
        }
      ]
    });

    const model = "gemini-3.5-flash-lite";

    const url =
      "https://generativelanguage.googleapis.com/v1beta/models/" +
      model +
      ":generateContent?key=" +
      apiKey;

    const payload = {
      systemInstruction: {
        parts: [{ text: systemPrompt }]
      },
      contents: contents,
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 2500
      }
    };

    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());

    if (result.error) {
      throw new Error(result.error.message);
    }

    if (
      !result.candidates ||
      !result.candidates[0] ||
      !result.candidates[0].content
    ) {
      throw new Error("Nepavyko gauti atsakymo iš Gemini.");
    }

    return result.candidates[0].content.parts
      .map(part => part.text || "")
      .join("");

  } catch (error) {
    throw new Error("Nepavyko išanalizuoti nuotraukos: " + error.message);
  }
}
