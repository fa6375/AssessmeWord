Office.onReady(() => {
    document.getElementById("extractLogs").onclick = exportJSON;
});

const SECRET_KEY = "your-secret-key";

// 📦 Read XML
async function getTrackingData(context) {
    const parts = context.document.customXmlParts;
    parts.load("items");
    await context.sync();

    let logs = [];

    for (const part of parts.items) {
        const xml = part.getXml();
        await context.sync();

        try {
            const parsed = JSON.parse(xml.value);
            logs.push(parsed);
        } catch (e) {}
    }

    return logs;
}

// 🔐 Decrypt
function decryptLogs(logs) {
    return logs.map(log => {
        try {
            const bytes = CryptoJS.AES.decrypt(log, SECRET_KEY);
            return JSON.parse(bytes.toString(CryptoJS.enc.Utf8));
        } catch {
            return null;
        }
    }).filter(l => l !== null);
}

// 🔁 Convert diffs → snapshots
function buildSnapshots(logs) {
    let snapshots = [];
    let currentText = "";

    logs.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));

    logs.forEach(log => {
        if (log.type === "snapshot") {
            currentText = log.content;
        }

        if (log.type === "diff") {
            log.content.forEach(([op, text]) => {
                if (op === 1) currentText += text;
                if (op === -1) {
                    currentText = currentText.replace(text, "");
                }
            });
        }

        snapshots.push({
            timestamp: log.timestamp,
            type: "snapshot",
            content: currentText
        });
    });

    return snapshots;
}

// 📦 Build final JSON
function buildFinalJSON(snapshots) {
    return {
        documentId: crypto.randomUUID(),
        metadata: {
            platform: "Word",
            createdAt: new Date().toISOString()
        },
        sessions: snapshots
    };
}

// 📄 Export JSON into Word (or console/file)
async function exportJSON() {
    await Word.run(async (context) => {

        const raw = await getTrackingData(context);
        const decrypted = decryptLogs(raw);
        const snapshots = buildSnapshots(decrypted);
        const finalJSON = buildFinalJSON(snapshots);

        const jsonString = JSON.stringify(finalJSON, null, 2);

        const newDoc = context.application.createDocument();
        newDoc.body.insertText(jsonString, "Start");

        await context.sync();
    });
}