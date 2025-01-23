let apiKey = ""; // Global variable for the API key
if (!window.addinInitialized) {
    window.addinInitialized = true; // Prevent duplicate initialization
    console.log("Initializing add-in...");

    Office.onReady((info) => {
        if (info.host === Office.HostType.Outlook) {
            // Initialize Outlook-specific code here
        } else if (info.host === Office.HostType.Word) {
            // Initialize Word-specific code here
            console.log("Office is ready");
            initializeEventListeners();

            // Load settings
            const savedApiKey = localStorage.getItem("openaiApiKey");
            if (savedApiKey) {
                apiKey = savedApiKey;
                console.log("API Key loaded from local storage.");
            }
        }
        // Add other host checks as needed
    });
} else {
    console.log("Add-in already initialized. Skipping initialization.");
}

function initializeEventListeners() {
    console.log("Initializing event listeners...");

    const sendButton = document.getElementById("send");
    sendButton.addEventListener("click", sendPrompt);

    const insertButton = document.getElementById("insert");
    insertButton.addEventListener("click", insertText);

    const replaceButton = document.getElementById("replace");
    replaceButton.addEventListener("click", replaceText);

    const copyButton = document.getElementById("copy");
    copyButton.addEventListener("click", copyToClipboard);

    const useSelectionToggle = document.getElementById("useSelection");
    const useWholeDocumentToggle = document.getElementById("useWholeDocument");

    // Ensure mutual exclusivity of context toggles
    useSelectionToggle.addEventListener("change", () => {
        if (useSelectionToggle.checked) {
            useWholeDocumentToggle.checked = false;
        }
    });

    useWholeDocumentToggle.addEventListener("change", () => {
        if (useWholeDocumentToggle.checked) {
            useSelectionToggle.checked = false;
        }
    });

    const settingsButton = document.getElementById("settingsButton");
    settingsButton.addEventListener("click", openSettings);

    const saveApiKeyButton = document.getElementById("saveApiKey");
    saveApiKeyButton.addEventListener("click", saveSettings);

    console.log("Event listeners initialized.");
}

function openSettings() {
    console.log("Opening settings...");
    const settingsModal = new bootstrap.Modal(document.getElementById("settingsModal"));
    settingsModal.show();
}

function saveSettings() {
    console.log("Saving settings...");
    const newApiKey = document.getElementById("apiKeyInput").value.trim();
    const selectedModel = document.getElementById("defaultModel").value;
    const selectedCharacter = document.getElementById("character").value;

    if (newApiKey) {
        apiKey = newApiKey; // Update the global variable
        localStorage.setItem("openaiApiKey", apiKey); // Save to local storage
        console.log("API Key saved successfully!");
    }

    localStorage.setItem("defaultModel", selectedModel);
    console.log("Default model updated to:", selectedModel);

    localStorage.setItem("characterStyle", selectedCharacter);
    console.log("Character style updated to:", selectedCharacter);

    showToast("Settings saved successfully!");
}

// Add this function to your JavaScript code
function addToastStyles() {
    const style = document.createElement('style');
    style.innerHTML = `
        .toast {
            opacity: 0;
            transition: opacity 2s ease-out; /* Adjust the duration as needed */
        }

        .toast.fade-in {
            opacity: 1;
        }

        .toast.fade-out {
            opacity: 0;
        }
    `;
    document.head.appendChild(style);
}

addToastStyles();

function showToast(message) {
    const toastContainer = document.createElement("div");
    toastContainer.className = "toast-container position-fixed top-0 end-0 p-3";
    document.body.appendChild(toastContainer);

    const toast = document.createElement("div");
    toast.className = "toast align-items-center text-white bg-primary border-0";
    toast.innerHTML = `
        <div class="d-flex">
            <div class="toast-body">${message}</div>
            <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
        </div>
    `;
    toastContainer.appendChild(toast);

    // Add fade-in class immediately to trigger the fade-in effect
    requestAnimationFrame(() => {
        toast.classList.add('fade-in');
    });

    const bsToast = new bootstrap.Toast(toast);
    bsToast.show();

    toast.addEventListener("hidden.bs.toast", () => {
        toastContainer.remove();
    });

    // Automatically fade out the toast after 5 seconds
    setTimeout(() => {
        toast.classList.remove('fade-in');
        toast.classList.add('fade-out');
    }, 5000);
}

async function sendPrompt() {
    console.log("sendPrompt function triggered");
    const useSelection = document.getElementById("useSelection").checked;
    const useWholeDocument = document.getElementById("useWholeDocument").checked;
    const character = localStorage.getItem("characterStyle") || "none";
    const selectedModel = localStorage.getItem("defaultModel") || "gpt-4";
    let promptInput = document.getElementById("prompt").value;
    let finalPrompt = "";

    try {
        if (useSelection) {
            console.log("Using selected text as context...");
            await Word.run(async (context) => {
                const range = context.document.getSelection();
                range.load("text");
                await context.sync();

                if (range.text && range.text.trim().length > 0) {
                    finalPrompt = `Context: "${range.text.trim()}". `;
                    console.log("Selected text:", range.text.trim());
                } else {
                    throw new Error("No text selected.");
                }
            });
        } else if (useWholeDocument) {
            console.log("Using the whole document as context...");
            await Word.run(async (context) => {
                const body = context.document.body;
                body.load("text");
                await context.sync();

                if (body.text && body.text.trim().length > 0) {
                    finalPrompt = `Context: "${body.text.trim()}". `;
                    console.log("Whole document text:", body.text.trim());
                } else {
                    throw new Error("The document is empty.");
                }
            });
        }

        if (character !== "none") {
            finalPrompt += `Write in the style of ${character}. `;
        }

        if (promptInput.trim()) {
            finalPrompt += `Instruction: "${promptInput.trim()}".`;
        }

        const responseDiv = document.getElementById("response");
        responseDiv.innerText = "Loading...";

        const res = await fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: selectedModel,
                messages: [{ role: "user", content: finalPrompt }],
                max_tokens: 1000
            })
        });

        if (!res.ok) {
            throw new Error(`HTTP error! status: ${res.status}`);
        }

        const data = await res.json();
        responseDiv.innerText = data.choices[0].message.content.trim();
    } catch (err) {
        console.error("Error in sendPrompt:", err.message);
        const responseDiv = document.getElementById("response");
        responseDiv.innerText = `Error: ${err.message}`;
    }
}

async function insertText() {
    console.log("insertText function triggered");
    const response = document.getElementById("response").innerText;

    await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("font");
        await context.sync();

        console.log("Selected text formatting:", range.font);

        const newRange = range.insertText(`\n${response}`, Word.InsertLocation.after);

        newRange.font.name = range.font.name;
        newRange.font.size = range.font.size;
        newRange.font.bold = range.font.bold;
        newRange.font.italic = range.font.italic;
        newRange.font.color = range.font.color;
        newRange.font.highlightColor = range.font.highlightColor;

        await context.sync();
        console.log("Inserted text below the selection with matching formatting:", response);
    }).catch((error) => {
        console.error("Error in insertText:", error);
    });
}

async function replaceText() {
    console.log("replaceText function triggered");
    const response = document.getElementById("response").innerText;

    await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(response, Word.InsertLocation.replace);
        await context.sync();
        console.log("Replacement successful with:", response);
    }).catch((error) => {
        console.error("Error in replaceText:", error);
    });
}

function copyToClipboard() {
    console.log("copyToClipboard function triggered");
    const response = document.getElementById("response").innerText;
    navigator.clipboard.writeText(response).then(() => {
        showToast("Response copied to clipboard!");
    }).catch((err) => {
        console.error("Failed to copy text:", err);
        showToast("Failed to copy text to clipboard.");
    });
}

document.addEventListener("DOMContentLoaded", function () {
    // Initialize Bootstrap tooltips
    var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'));
    var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
        return new bootstrap.Tooltip(tooltipTriggerEl);
    });
});
