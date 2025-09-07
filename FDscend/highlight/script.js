document.addEventListener("DOMContentLoaded", (event) => {
  const codeInput = document.getElementById("codeInput");
  const styleSelect = document.getElementById("styleSelect");
  const languageSelect = document.getElementById("languageSelect");
  const highlightButton = document.getElementById("highlightButton");
  const copyButton = document.getElementById("copyButton");
  const highlightedCode = document.getElementById("highlightedCode");
  const highlightStyleLink = document.getElementById("highlight-style");

  var highlighted = false;

  function updateBackgroundColor() {
    const hljsElement = document.querySelector(".hljs");
    if (hljsElement) {
      const computedStyle = window.getComputedStyle(hljsElement);
      const backgroundColor = computedStyle.backgroundColor;
      document.querySelectorAll(".hljs *").forEach((element) => {
        element.style.backgroundColor = backgroundColor;
      });
      // console.log('Background color:', backgroundColor);
    }
  }

  highlightButton.addEventListener("click", () => {
    const code = codeInput.value;
    const language = languageSelect.value;
    if (language === "auto") {
      highlightedCode.innerHTML = hljs.highlightAuto(code).value;
    } else {
      highlightedCode.innerHTML = hljs.highlight(code, { language }).value;
    }
    updateBackgroundColor();

    highlighted = true;
  });

  styleSelect.addEventListener("change", () => {
    const style = styleSelect.value;
    highlightStyleLink.href = `highlight/styles/${style}.min.css`;
    setTimeout(updateBackgroundColor, 100);
    // console.log('updateBackgroundColor');
  });

  copyButton.addEventListener("click", () => {
    const range = document.createRange();
    range.selectNodeContents(highlightedCode);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    document.execCommand("copy");

    if (highlighted) window.chrome.webview.postMessage("codeCopiedToClipboard");
    else window.chrome.webview.postMessage("notHighLighted");
  });
});
