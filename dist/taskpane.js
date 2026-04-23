Office.onReady(() => {
    document.getElementById("exportJsonBtn").onclick = exportJSON;
});

const SECRET_KEY = "your-secret-key";

// 📦 Read XML logs
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

// 🔐 Decrypt logs
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

// 💾 DOWNLOAD JSON FILE
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

// 🚀 MAIN EXPORT FUNCTION
async function exportJSON() {
    await Word.run(async (context) => {

        const raw = await getTrackingData(context);
        const decrypted = decryptLogs(raw);
        const snapshots = buildSnapshots(decrypted);
        const finalJSON = buildFinalJSON(snapshots);

        downloadJSON(finalJSON);

        await context.sync();
    });
}