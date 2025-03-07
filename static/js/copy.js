document.addEventListener("DOMContentLoaded", function () {
    const copyButton = document.getElementById("copy-button");
    const errorText = document.getElementById("error-text");

    if (copyButton && errorText) {
        copyButton.addEventListener("click", function () {
            navigator.clipboard.writeText(errorText.innerText).then(() => {
                copyButton.innerText = "✅ Copied!";
                setTimeout(() => {
                    copyButton.innerText = "📋 Copy";
                }, 2000);
            });
        });
    }
});