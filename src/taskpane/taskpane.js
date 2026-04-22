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
        // First time setup if no log exists
        const timestamp = new Date().toISOString();
        lastStudentId = currentStudent;
        lastUsername = currentUsername;
        lastCourseName = currentCourse;        

        const systemInfo = {
          type: "system-info",
          time: timestamp,
          platform: Office.context.diagnostics.platform,
          host: Office.context.diagnostics.host,
          version: Office.context.diagnostics.version,
          language: Office.context.displayLanguage,
          student: currentStudent,
          username: currentUsername,
          course: currentCourse,
          documentId: getDocumentId()
        };

        const encryptedInfo = encrypt(JSON.stringify(systemInfo));
        const firstEntry = `<entry time="${timestamp}">${JSON.stringify({ time: timestamp, content: encryptedInfo, documentId: systemInfo.documentId })}</entry>`;
        const newXml = `<log xmlns="urn:assessme-log">${firstEntry}</log>`;

        context.document.customXmlParts.add(newXml);
        await context.sync();
        return;
      }

      // Otherwise: Log already exists → load XML
      let xmlResult = logPart.getXml();
      await context.sync();

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlResult.value, "text/xml");

      // --- 1. Update system-info if Username or Student Name changed ---
      const firstEntry = xmlDoc.getElementsByTagName("entry")[0];
      if (firstEntry) {
        const json = JSON.parse(firstEntry.textContent);
        const decrypted = decrypt(json.content);
        const systemInfo = JSON.parse(decrypted);

        let changed = false;

        if (systemInfo.username !== currentUsername) {
          systemInfo.username = currentUsername;
          changed = true;
        }
        if (systemInfo.student !== currentStudent) {
          systemInfo.student = currentStudent;
          changed = true;
        }
        if (systemInfo.course !== currentCourse) {
          systemInfo.course = currentCourse;
          changed = true;
      }
      

        if (changed) {
          const newEncrypted = encrypt(JSON.stringify(systemInfo));
          json.content = newEncrypted;
          firstEntry.textContent = JSON.stringify(json);

          lastStudentId = currentStudent;
          lastUsername = currentUsername;
          lastCourseName = currentCourse;
        }
      }

      // --- 2. Add new diff if body text meaningfully changed ---
      const diffs = dmp.diff_main(lastSnapshot, currentSnapshot);
      dmp.diff_cleanupSemantic(diffs);

      const meaningfulDiffs = diffs.filter(([op, data]) => op !== 0 && data.trim().length > 0);
      const totalChangedLength = meaningfulDiffs.reduce((sum, [_, data]) => sum + data.length, 0);

      if (totalChangedLength > 5) {
        const timestamp = new Date().toISOString();
        const documentId = getDocumentId();

        let payload;
        if (logCounter % 10 === 0) {
          payload = {
            time: timestamp,
            type: "full",
            content: currentSnapshot,
            documentId
          };
        } else {
          payload = {
            time: timestamp,
            type: "diff",
            content: meaningfulDiffs,
            documentId
          };
        }

        lastSnapshot = currentSnapshot;
        logCounter++;

        const encryptedContent = encrypt(JSON.stringify(payload));
        const logJson = JSON.stringify({
          time: timestamp,
          content: encryptedContent,
          documentId
        });

        const newEntry = xmlDoc.createElement("entry");
        newEntry.setAttribute("time", timestamp);
        newEntry.textContent = logJson;
        xmlDoc.documentElement.appendChild(newEntry);
        lastLogTime = Date.now();
        showActiveStatus();
      }

      // --- 3. Save everything back once ---
      const serializer = new XMLSerializer();
      const finalUpdatedXml = serializer.serializeToString(xmlDoc);

      await logPart.delete();
      context.document.customXmlParts.add(finalUpdatedXml);
      await context.sync();
    }).catch(error => {
      showErrorStatus();
    });
    const now = Date.now();
    if (now - lastLogTime > 60000) {  // 60 seconds
      showIdleStatus();
    }
  }, autosaveInterval);
}

function showErrorStatus() {
  const banner = document.getElementById("statusBanner");
  banner.className = "status-banner status-error";
  banner.innerText = "❌ Recording stopped. Critical error occurred.";
}

function showIdleStatus() {
  const banner = document.getElementById("statusBanner");
  banner.className = "status-banner status-idle";
  banner.innerText = "⏸️ Recording paused due to inactivity. Continue when you are ready.";
}

function showActiveStatus() {
  const banner = document.getElementById("statusBanner");
  banner.className = "status-banner status-active";
  banner.innerText = "✅ AssessMe is recording your activity.";
}
