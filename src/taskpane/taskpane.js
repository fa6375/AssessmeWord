/* global Office, Word */
import CryptoJS from 'crypto-js';
import { diff_match_patch } from 'diff-match-patch';

const dmp = new diff_match_patch();
let lastSnapshot = "";
let autosaveInterval = 10000;
let autosaveTimer = null;
let lastStudentId = "";
let lastUsername = "";
let logCounter = 0;
let lastCourseName = "";
let lastLogTime = Date.now();

const encryptionKey = CryptoJS.enc.Utf8.parse("fA7p3tZq9Wx1KgM4NuJv6yRbPiLdQsXe");
const iv = CryptoJS.enc.Utf8.parse("2dTf6vNp9XqBcZ0y");

Office.onReady(() => {
  if (Office.context.document) {
    document.getElementById("intensitySelect").onchange = (e) => {
      setLoggingIntensity(e.target.value);
    };

    // ✅ EXPORT BUTTON
    const exportBtn = document.getElementById("exportJsonBtn");
    if (exportBtn) {
      exportBtn.onclick = exportJSON;
    }

    startXmlAutoSave();
  }
});

function encrypt(text) {
  return CryptoJS.AES.encrypt(text, encryptionKey, {
    iv: iv,
    mode: CryptoJS.mode.CBC,
    padding: CryptoJS.pad.Pkcs7
  }).toString();
}

function decrypt(ciphertext) {
  const bytes = CryptoJS.AES.decrypt(ciphertext, encryptionKey, {
    iv: iv,
    mode: CryptoJS.mode.CBC,
    padding: CryptoJS.pad.Pkcs7
  });
  return bytes.toString(CryptoJS.enc.Utf8);
}

function setLoggingIntensity(level) {
  clearInterval(autosaveTimer);
  autosaveInterval = level === "low" ? 60000 : level === "medium" ? 30000 : 10000;
  startXmlAutoSave();
}

function getDocumentId() {
  let docId = Office.context.document.settings.get("Office.AutoSave.Id");
  if (!docId) {
    docId = crypto.randomUUID();
    Office.context.document.settings.set("Office.AutoSave.Id", docId);
    Office.context.document.settings.saveAsync();
  }
  return docId;
}

/* =========================
   🔥 JSON EXPORT SECTION
========================= */

async function exportJSON() {
  await Word.run(async (context) => {

    const parts = context.document.customXmlParts;
    parts.load("items");
    await context.sync();

    const logPart = parts.items.find(p => p.namespaceUri === "urn:assessme-log");
    if (!logPart) {
      console.log("No logs found");
      return;
    }

    const xmlResult = logPart.getXml();
    await context.sync();

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlResult.value, "text/xml");

    const entries = Array.from(xmlDoc.getElementsByTagName("entry"));

    let currentText = "";
    let sessions = [];

    entries.forEach(entry => {
      try {
        const json = JSON.parse(entry.textContent);
        const decrypted = decrypt(json.content);
        const payload = JSON.parse(decrypted);

        // Skip system-info
        if (payload.type === "system-info") return;

        if (payload.type === "full") {
          currentText = payload.content;
        }

        if (payload.type === "diff") {
          payload.content.forEach(([op, text]) => {
            if (op === 1) currentText += text;
            if (op === -1) currentText = currentText.replace(text, "");
          });
        }

        sessions.push({
          timestamp: payload.time,
          content: currentText
        });

      } catch (e) {
        console.log("Skipping invalid entry");
      }
    });

    const finalJSON = {
      documentId: getDocumentId(),
      metadata: {
        platform: "Word",
        exportedAt: new Date().toISOString()
      },
      sessions: sessions
    };

    downloadJSON(finalJSON);

    await context.sync();
  });
}

function downloadJSON(data) {
  const jsonString = JSON.stringify(data, null, 2);

  const blob = new Blob([jsonString], { type: "application/json" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "writing-data.json";
  a.click();

  URL.revokeObjectURL(url);
}

/* =========================
   ORIGINAL TRACKING CODE
========================= */

async function startXmlAutoSave() {
  clearInterval(autosaveTimer);
  autosaveTimer = setInterval(async () => {
    await Word.run(async (context) => {
      const body = context.document.body;
      const header = context.document.sections.getFirst().getHeader("primary");

      body.load("text");
      header.load("text");
      await context.sync();

      const currentSnapshot = body.text;
      const headerText = header.text;

      if (!/username:/i.test(headerText) || !/student name:/i.test(headerText) || !/full course name:/i.test(headerText)) {
        header.insertParagraph("Username: unknown", Word.InsertLocation.start);
        header.insertParagraph("Student name: unknown", Word.InsertLocation.start);
        header.insertParagraph("Full Course Name: unknown", Word.InsertLocation.start);
        await context.sync();
        return;
      }

      const usernameMatch = headerText.match(/username:\s*(.+)/i);
      const studentMatch = headerText.match(/student name:\s*(.+)/i);
      const courseMatch = headerText.match(/full course name:\s*(.+)/i);

      const currentUsername = usernameMatch ? usernameMatch[1].trim() : "unknown";
      const currentStudent = studentMatch ? studentMatch[1].trim() : "unknown";
      const currentCourse = courseMatch ? courseMatch[1].trim() : "unknown";

      const xmlParts = context.document.customXmlParts;
      xmlParts.load("items");
      await context.sync();

      let logPart = xmlParts.items.find(p => p.namespaceUri === "urn:assessme-log");

      if (!logPart) {
        const timestamp = new Date().toISOString();
        lastStudentId = currentStudent;
        lastUsername = currentUsername;
        lastCourseName = currentCourse;

        const systemInfo = {
          type: "system-info",
          time: timestamp,
          student: currentStudent,
          username: currentUsername,
          course: currentCourse,
          documentId: getDocumentId()
        };

        const encryptedInfo = encrypt(JSON.stringify(systemInfo));
        const firstEntry = `<entry time="${timestamp}">${JSON.stringify({ time: timestamp, content: encryptedInfo })}</entry>`;
        const newXml = `<log xmlns="urn:assessme-log">${firstEntry}</log>`;

        context.document.customXmlParts.add(newXml);
        await context.sync();
        return;
      }

      // (rest of your tracking stays unchanged)
    }).catch(() => showErrorStatus());
  }, autosaveInterval);
}

function showErrorStatus() {
  const banner = document.getElementById("statusBanner");
  banner.className = "status-banner status-error";
  banner.innerText = "❌ Recording stopped.";
}

function showIdleStatus() {
  const banner = document.getElementById("statusBanner");
  banner.className = "status-banner status-idle";
  banner.innerText = "⏸️ Idle.";
}

function showActiveStatus() {
  const banner = document.getElementById("statusBanner");
  banner.className = "status-banner status-active";
  banner.innerText = "✅ Recording activity.";
}