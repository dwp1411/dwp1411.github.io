const fs = require('fs');

// The original git restore blew away the Index.html changes from earlier steps too, not just Code.gs!
// We lost the voice to text and the form transformation because of `git restore new_jso_project/Index.html`
// when trying to fix Code.gs earlier.
// I need to rebuild all 3 steps for Index.html.

let html = fs.readFileSync('new_jso_project/Index.html', 'utf8');

// 1. ADD VOICE
const voiceScript = `
      // Speech Recognition
      let recognition;
      if ('webkitSpeechRecognition' in window) {
          recognition = new webkitSpeechRecognition();
          recognition.continuous = false;
          recognition.interimResults = false;
      }

      function startVoiceRecognition(targetId, btn) {
          if (!recognition) {
              alert("Voice to text is not supported in this browser. Please use Chrome.");
              return;
          }

          const targetInput = document.getElementById(targetId);
          const originalText = btn.innerHTML;

          recognition.onstart = function() {
              btn.innerHTML = '🔴 Listening...';
              btn.style.color = 'red';
          };

          recognition.onresult = function(event) {
              const transcript = event.results[0][0].transcript;
              if (targetInput.value) {
                 targetInput.value += ' ' + transcript;
              } else {
                 targetInput.value = transcript;
              }
          };

          recognition.onerror = function(event) {
              console.error("Speech error", event.error);
              alert("Microphone error. Please try again or type manually.");
          };

          recognition.onend = function() {
              btn.innerHTML = originalText;
              btn.style.color = '';
          };

          recognition.start();
      }
`;
if (!html.includes("webkitSpeechRecognition")) {
    html = html.replace('// 1. Submit Basic Form', voiceScript + '\n\n      // 1. Submit Basic Form');
}

// Ensure dynamic checklists still exist. (Wait, the git checkout blew them away too).
// Let's re-run the checklist injection first.
