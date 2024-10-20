let clipboardData = [];
let allParagraphsData = []; // Store all paragraphs and numbering

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Load all paragraphs and their numbering on add-in load
    loadAllParagraphsData();

    document.getElementById("logStyleContentButton").onclick = getListInfoFromSelection;
    document.getElementById("clearContentButton").onclick = clearCopiedContent;
  }
});

// Helper function to clean and normalize text
function normalizeText(text) {
  return text
    .trim()
    .replace(/\s+/g, " ") // Replace multiple spaces with single space
    .replace(/[^\x20-\x7E]/g, ""); // Clean up non-ASCII characters
}

// Function to load all paragraphs' numbering and text
async function loadAllParagraphsData() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      allParagraphsData = []; // Reset the array
      let parentNumbering = [];
      let lastNumbering = ""; // Store the last used numbering

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,isListItem");
        await context.sync();

        let text = normalizeText(paragraph.text);

        if (text.length <= 1) {
          continue;
        }

        if (paragraph.isListItem) {
          paragraph.listItem.load("level,listString");
          await context.sync();

          const level = paragraph.listItem.level;
          const listString = paragraph.listItem.listString || "";

          // Adjust the parentNumbering array based on the current level
          if (level <= parentNumbering.length) {
            parentNumbering = parentNumbering.slice(0, level);
          }

          parentNumbering[level] = listString;

          let fullNumbering = "";
          for (let j = 0; j <= level; j++) {
            if (parentNumbering[j]) {
              fullNumbering += `${parentNumbering[j]}.`;
            }
          }

          fullNumbering = fullNumbering.replace(/\.$/, ""); // Remove trailing dot
          lastNumbering = fullNumbering; // Update the last used numbering

          allParagraphsData.push({
            key: fullNumbering,
            value: text,
            originalText: paragraph.text.trim(), // Store original text for matching
            isListItem: true,
          });
        } else {
          // If it's not a list item, use the last known numbering + .text
          const key = lastNumbering ? `${lastNumbering}.text` : `text_${i + 1}`;
          allParagraphsData.push({
            key: key,
            value: text,
            originalText: paragraph.text.trim(), // Store original text for matching
            isListItem: false,
          });
        }
      }

      console.log("All paragraphs data loaded:", allParagraphsData);
    });
  } catch (error) {
    console.error("An error occurred while loading all paragraphs data:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
  }
}

// Function to handle copying data based on user selection
async function getListInfoFromSelection() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      // Don't reset clipboardData anymore - let it accumulate
      let newSelections = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const selectedParagraph = paragraphs.items[i];
        selectedParagraph.load("text");
        await context.sync();

        const selectedText = selectedParagraph.text.trim();
        const normalizedSelectedText = normalizeText(selectedText);

        // Find the matching paragraph using both normalized and original text
        const matchingParagraphData = allParagraphsData.find(
          (para) => para.value === normalizedSelectedText || para.originalText === selectedText
        );

        if (matchingParagraphData) {
          // Check if this item is already in clipboardData
          const isDuplicate = clipboardData.some(
            (item) => item.key === matchingParagraphData.key && item.value === matchingParagraphData.value
          );

          if (!isDuplicate) {
            newSelections.push({
              key: matchingParagraphData.key,
              value: matchingParagraphData.value,
            });
          }
        } else {
          console.log("No match found for:", selectedText);
        }
      }

      if (newSelections.length > 0) {
        // Add new selections to existing clipboardData
        clipboardData = [...clipboardData, ...newSelections];

        // Sort clipboardData based on key
        clipboardData.sort((a, b) => {
          // Extract numbers from keys for proper numerical sorting
          const aMatch = a.key.match(/\d+/g);
          const bMatch = b.key.match(/\d+/g);

          if (aMatch && bMatch) {
            return parseInt(aMatch[0]) - parseInt(bMatch[0]);
          }
          return a.key.localeCompare(b.key);
        });

        updateCopiedContentDisplay();
        const clipboardString = formatClipboardData();
        await copyToClipboard(clipboardString);

        console.log("Updated clipboard data:", clipboardString);
      } else {
        console.log("No new paragraphs to add.");
      }
    });
  } catch (error) {
    console.error("An error occurred while copying data:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
  }
}

function formatClipboardData() {
  return `{\n${clipboardData.map((pair) => `"${pair.key}": "${pair.value}"`).join(",\n")}\n}`;
}

function updateCopiedContentDisplay() {
  const copiedContentElement = document.getElementById("copiedContent");
  copiedContentElement.innerHTML = "";

  clipboardData.forEach((pair) => {
    const keySpan = `<span class="key">${pair.key}</span>`;
    const valueSpan = `<span class="value">${pair.value}</span>`;
    const formattedPair = `<div class="pair">${keySpan}: ${valueSpan}</div>`;
    copiedContentElement.innerHTML += formattedPair;
  });

  copiedContentElement.scrollTop = copiedContentElement.scrollHeight;
}

async function copyToClipboard(text) {
  try {
    // Try to use the modern clipboard API first
    await navigator.clipboard.writeText(text);
    showCopyMessage(true);
  } catch (err) {
    // Fall back to the older execCommand method
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.position = "fixed";
    textArea.style.left = "-999999px";
    textArea.style.top = "-999999px";
    document.body.appendChild(textArea);

    try {
      textArea.focus();
      textArea.select();
      const successful = document.execCommand("copy");
      showCopyMessage(successful);
    } catch (err) {
      console.error("Unable to copy to clipboard", err);
      showCopyMessage(false);
    } finally {
      document.body.removeChild(textArea);
    }
  }
}

function showCopyMessage(successful) {
  const copyMessage = document.getElementById("copyMessage");
  copyMessage.style.display = "block";
  copyMessage.textContent = successful ? "Content added and copied to clipboard!" : "Failed to copy content";
  copyMessage.style.color = successful ? "green" : "red";

  setTimeout(() => {
    copyMessage.style.display = "none";
  }, 3000);
}

function clearCopiedContent() {
  clipboardData = [];
  document.getElementById("copiedContent").innerHTML = "";
}
