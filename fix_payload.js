const fs = require('fs');

let html = fs.readFileSync('new_jso_project/Index.html', 'utf8');

// I might have replaced it incorrectly earlier if the block didn't exactly match.
// Let's check if the old data object is still there.

const oldDataObjRegex = /const data = {\s*leader: leader,\s*associate: associate,\s*duration: document\.getElementById\('stopwatch'\)\.textContent,\s*checklistResults: checklistResults,\s*notes: combinedNotes,\s*function: func,\s*totalCount: totalCount,\s*totalSeconds: seconds\s*};/;

const newDataObj = `
        const fbPace = document.querySelector('input[name="fb-pace"]:checked')?.value || "N/A";
        const fbSkills = document.querySelector('input[name="fb-skills"]:checked')?.value || "N/A";
        const fbTraining = document.querySelector('input[name="fb-training"]:checked')?.value || "N/A";
        const fbEnvironment = document.querySelector('input[name="fb-environment"]:checked')?.value || "N/A";
        const fbNotes = document.getElementById('fb-notes').value || "None";
        const signatureBase64 = signatureEmpty ? null : canvas.toDataURL("image/png");

        const data = {
            leader: leader,
            associate: associate,
            duration: document.getElementById('stopwatch').textContent,
            checklistResults: checklistResults,
            function: func,
            totalCount: totalCount,
            totalSeconds: seconds,
            activityLogs: formattedLogs,
            fbPace: fbPace,
            fbSkills: fbSkills,
            fbTraining: fbTraining,
            fbEnvironment: fbEnvironment,
            fbNotes: fbNotes,
            signature: signatureBase64
        };`;

if (html.match(oldDataObjRegex)) {
    html = html.replace(oldDataObjRegex, newDataObj);
    fs.writeFileSync('new_jso_project/Index.html', html);
    console.log("Fixed payload.");
} else {
    console.log("Could not find the data object block.");
}
