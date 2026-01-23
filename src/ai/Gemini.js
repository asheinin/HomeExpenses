/**
 * Handles interactions with the Gemini API.
 */

/**
 * Calls the Gemini API with a given prompt.
 * @param {string} prompt The prompt to send to Gemini.
 * @returns {string|null} The response from Gemini, or null if failed.
 */
function callGemini(prompt) {
    const apiKey = getGeminiApiKey();
    if (!apiKey) {
        console.error("Gemini API Key not found in Script Properties.");
        return null;
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

    const payload = {
        contents: [
            {
                parts: [
                    { text: prompt }
                ]
            }
        ],
        generationConfig: {
            temperature: 0.7,
            topK: 40,
            topP: 0.95,
            maxOutputTokens: 1024,
        }
    };

    const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        const response = UrlFetchApp.fetch(url, options);
        const json = JSON.parse(response.getContentText());

        if (response.getResponseCode() !== 200) {
            console.error("Gemini API Error: " + response.getContentText());
            return null;
        }

        if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts[0]) {
            return json.candidates[0].content.parts[0].text;
        } else {
            console.error("Unexpected Gemini API response format: " + JSON.stringify(json));
            return null;
        }
    } catch (e) {
        console.error("Error calling Gemini API: " + e.toString());
        return null;
    }
}

/**
 * Retrieves the Gemini API Key from Script Properties.
 * @returns {string|null}
 */
function getGeminiApiKey() {
    return PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
}

/**
 * Helper function to set the Gemini API Key.
 * Can be run manually by the user or via a UI prompt.
 */
function setGeminiApiKey() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Set Gemini API Key", "Enter your Gemini API key from Google AI Studio:", ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.OK) {
        const key = response.getResponseText().trim();
        if (key) {
            PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", key);
            ui.alert("API Key saved successfully.");
        } else {
            ui.alert("Invalid API Key.");
        }
    }
}
