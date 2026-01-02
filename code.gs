/**
 * ======================================================
 * ⚙️ SETTINGS & CONFIGURATION
 * ======================================================
 */
const CONFIG = {
  // List of Sheet Tabs where the script will run.
  TARGET_SHEETS: ["MANAM", "8CUTS", "OOMA", "DTF", "MO'COOKIES", "EXPRESS", "FOUNDATION", "CENTRAL KITCHEN"],
  
  // Name of the Master Sheet.
  OVERALL_SHEET: "OVERALL",
  
  // The internal email address to ignore (we don't want to track our own replies).
  EMAIL_RECEIVER: "scm.nonfood@momentgroup.ph"
};

// ======================================================
// 1. MANUAL ENTRY AUTOMATION
// Triggers instantly when you type in a cell.
// ======================================================
function automationsTrigger(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  // --- SAFETY CHECKS ---
  // Stop if editing the header, wrong sheet, or deleting data.
  if (range.getRow() === 1) return;
  if (!CONFIG.TARGET_SHEETS.includes(sheet.getName()) && sheet.getName() !== CONFIG.OVERALL_SHEET) return;
  if (range.getValue() === "") return;

  const cols = getColumnMap(sheet);
  const editedCol = range.getColumn();
  const row = range.getRow();
  const val = String(range.getValue());

  // ------------------------------------------------------
  // SCENARIO A: You typed a "PR NUMBER"
  // ------------------------------------------------------
  if (cols["PR NUMBER"] && editedCol === cols["PR NUMBER"]) {
    
    // Rule 1: Format Check (Must be PR-SHOP-NUMBER)
    const prFormatRegex = /^PR-[A-Z0-9']+-[\d]+$/i; 
    if (!prFormatRegex.test(val)) return; 

    // Rule 2: Duplicate Check
    if (isDuplicate(sheet, cols["PR NUMBER"], val, row)) {
      e.source.toast("⚠️ Notice: PR Duplicate found. Auto-fill skipped.", "Info", -1);
      return; 
    }

    // Rule 3: Gmail Search & Auto-Fill
    if (cols["DESCRIPTION"]) {
      // Get current values to see if they are empty
      const currentDesc = sheet.getRange(row, cols["DESCRIPTION"]).getValue();
      const currentStatus = cols["STATUS"] ? sheet.getRange(row, cols["STATUS"]).getValue() : "";
      const currentDate = cols["PR DATE RECEIVED"] ? sheet.getRange(row, cols["PR DATE RECEIVED"]).getValue() : "";

      // Protection: Only run if the target cells are empty (prevents overwriting)
      if (currentDesc === "" && currentStatus === "" && currentDate === "") {
        try {
          // Manual search looks for ANY email (read or unread) to help you fill data
          const threads = GmailApp.search(`"${val}"`, 0, 1);
          
          if (threads.length > 0) {
            const thread = threads[0];
            const msg = thread.getMessages()[0];
            
            // Clean the Subject Line (Remove "Re:", "Fwd:")
            const cleanSubject = msg.getSubject().replace(/^(Re:|Fwd:)\s*/i, "").trim();
            
            // Write Description & Set Text Wrap to CLIP
            sheet.getRange(row, cols["DESCRIPTION"])
                 .setValue(cleanSubject)
                 .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

            // Set Status
            if (cols["STATUS"]) sheet.getRange(row, cols["STATUS"]).setValue("CANVASSING");

            // Set Date (Condition: 1 Message Only + External Sender)
            if (cols["PR DATE RECEIVED"]) {
              const isSingleConvo = thread.getMessageCount() === 1;
              const isExternalSender = !msg.getFrom().includes(CONFIG.EMAIL_RECEIVER);

              if (isSingleConvo && isExternalSender) {
                sheet.getRange(row, cols["PR DATE RECEIVED"]).setValue(msg.getDate());
              }
            }
            e.source.toast("✅ Auto-filled from Email", "Success");
          }
        } catch (err) {
          // Silent fail for manual entry is fine
        }
      } 
    }

    // Rule 4: Create Blue Clickable Link
    if (range.getFormula() === "") {
      const link = `https://mail.google.com/mail/u/0/#search/${encodeURIComponent(val)}`;
      range.setFormula(`=HYPERLINK("${link}", "${val}")`);
    }
  }

  // ------------------------------------------------------
  // SCENARIO B: You typed a "SUPPLIER"
  // ------------------------------------------------------
  else if (cols["SUPPLIER"] && editedCol === cols["SUPPLIER"]) {
    const prVal = sheet.getRange(row, cols["PR NUMBER"]).getValue();
    const branchVal = sheet.getRange(row, cols["BRANCH"]).getValue();
    
    // Check if PR + Branch + Supplier combo already exists
    if (prVal && branchVal && checkComboDuplicate(sheet, cols, prVal, branchVal, val, row)) {
      e.source.toast("⛔ REJECTED: PR + Branch + Supplier combo exists!", "Error", -1);
      range.setValue(""); 
    }
  }

  // ------------------------------------------------------
  // SCENARIO C: You typed a "PO NUMBER"
  // ------------------------------------------------------
  else if (cols["PO NUMBER"] && editedCol === cols["PO NUMBER"]) {
    if (isDuplicate(sheet, cols["PO NUMBER"], val, row)) {
      e.source.toast("⛔ REJECTED: PO Number already exists!", "Error", -1);
      range.setValue(""); 
    }
  }
}

// ======================================================
// 2. AUTOMATIC BACKGROUND SCANNER
// Runs in background every few minutes.
// ======================================================
function scanEmailsForPRs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Map code in Subject to Sheet Name
  const shopMap = { "MANAM": "MANAM", "8CUTS": "8CUTS", "OOMA": "OOMA", "DTF": "DTF", "MO": "MO'COOKIES", "CK": "CENTRAL KITCHEN" };

  // SEARCH RULES:
  // 1. is:unread -> Only look for new emails.
  // 2. subject:(PR) -> Must have "PR" in title.
  // 3. -from:internal -> Ignore internal emails.
  // 4. after:2025/12/31 -> DATE GUARD. Only emails from 2026 onwards.
  const query = `is:unread subject:(PR) -from:${CONFIG.EMAIL_RECEIVER} after:2025/12/31`;
  
  try {
    // Get up to 50 unread threads (Batch Size)
    const threads = GmailApp.search(query, 0, 50);
    
    // Create virtual "Buckets" to organize emails by Shop
    const groupedData = {};

    // --- Step 1: Filter and Sort Emails ---
    threads.forEach(thread => {
      // Strict Rule: Only process new requests (exactly 1 message in conversation)
      if (thread.getMessageCount() !== 1) return;

      const msg = thread.getMessages()[0]; 
      
      // Double check unread status
      if (!msg.isUnread()) return;

      const subject = msg.getSubject();
      
      // Pattern Match: Finds "PR MANAM 123" or "PR-DTF-123"
      const pattern = /PR[\s-]*\b(MANAM|8CUTS|OOMA|DTF|MO|CK)\b[\s-]*(\d+)/i;
      const match = subject.match(pattern);

      if (match) {
        const code = match[1].toUpperCase();
        const num = match[2];
        const sheetName = shopMap[code];
        
        if (sheetName) {
          if (!groupedData[sheetName]) groupedData[sheetName] = [];
          
          // Add email to the correct shop bucket
          groupedData[sheetName].push({
            prStr: `PR-${code}-${num}`,
            date: msg.getDate(),
            desc: subject.replace(/^(Re:|Fwd:)\s*/i, "").trim(),
            messageObj: msg // Save this so we can mark it read later
          });
        }
      }
    });

    // --- Step 2: Process Buckets (One Shop at a time) ---
    Object.keys(groupedData).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      const cols = getColumnMap(sheet);
      if (!cols["PR NUMBER"]) return;

      // Batch Read: Get all existing PRs in one go (Fast)
      const lastRow = sheet.getLastRow();
      let existingPrs = [];
      if (lastRow > 1) {
        existingPrs = sheet.getRange(2, cols["PR NUMBER"], lastRow - 1, 1).getValues().flat().map(String);
      }

      // Loop through new emails for this shop
      groupedData[sheetName].forEach(item => {
        // Check if PR is already in the sheet (Avoid Duplicates)
        const isDup = existingPrs.some(p => p.includes(item.prStr));

        if (!isDup) {
          const nextRow = sheet.getLastRow() + 1;
          const link = `https://mail.google.com/mail/u/0/#search/${encodeURIComponent(item.prStr)}`;

          // 1. Write PR Number with Link
          sheet.getRange(nextRow, cols["PR NUMBER"]).setFormula(`=HYPERLINK("${link}", "${item.prStr}")`);
          
          // 2. Fill other details
          if (cols["PR DATE RECEIVED"]) sheet.getRange(nextRow, cols["PR DATE RECEIVED"]).setValue(item.date);
          
          // 3. Write Description and SET CLIP WRAPPING
          if (cols["DESCRIPTION"]) {
            sheet.getRange(nextRow, cols["DESCRIPTION"])
                 .setValue(item.desc)
                 .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
          }

          if (cols["STATUS"]) sheet.getRange(nextRow, cols["STATUS"]).setValue("CANVASSING");
          if (cols["REMARKS"]) sheet.getRange(nextRow, cols["REMARKS"]).setValue("Auto-generated via Email");
          if (cols["PR CATEGORY"]) sheet.getRange(nextRow, cols["PR CATEGORY"]).setValue("SIMPLE");

          console.log(`✅ Added ${item.prStr}`);
          
          // Mark as READ (Success)
          item.messageObj.markRead();
        } else {
          // If duplicate, mark read anyway so we don't get stuck processing it forever
          item.messageObj.markRead();
        }
      });
    });

  } catch (err) {
    console.log("❌ Scanner Error: " + err);
  }
}

// ======================================================
// 3. HELPER TOOLS (Utilities)
// ======================================================

function getColumnMap(sheet) {
  if (sheet.getLastColumn() < 1) return {};
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => map[String(h).trim().toUpperCase()] = i + 1);
  return map;
}

function isDuplicate(sheet, colIndex, value, currentRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const data = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues().flat();
  return data.some((v, i) => String(v) === String(value) && (i + 2) !== currentRow);
}

function checkComboDuplicate(sheet, cols, pr, branch, supplier, currentRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  
  const prs = sheet.getRange(2, cols["PR NUMBER"], lastRow - 1, 1).getValues().flat();
  const branches = sheet.getRange(2, cols["BRANCH"], lastRow - 1, 1).getValues().flat();
  const suppliers = sheet.getRange(2, cols["SUPPLIER"], lastRow - 1, 1).getValues().flat();

  for (let i = 0; i < prs.length; i++) {
    if ((i + 2) === currentRow) continue; 
    if (String(prs[i]) === String(pr) && 
        String(branches[i]) === String(branch) && 
        String(suppliers[i]) === String(supplier)) {
      return true;
    }
  }
  return false;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚡ Admin Tools')
    .addItem('Repair STATUS Dropdowns', 'fixDataValidation')
    .addToUi();
}

/**
 * REPAIR FUNCTION
 * Finds the STATUS column and applies strict Dropdown (Chips) with specific values.
 */
function fixDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // The specific list of allowed values
  const statusValues = [
    "CANCELLED", 
    "COMPLETED", 
    "INCOMPLETE SPECS", 
    "R&M/FINANCE APPROVAL", 
    "CANVASSING", 
    "ON HOLD"
  ];

  // Build the Rule: Dropdown (Chips) + Reject Invalid Input
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusValues, true) // true = Show dropdown (Chips)
    .setAllowInvalid(false)                 // Reject invalid input
    .build();

  CONFIG.TARGET_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    // Dynamically find which column is "STATUS"
    const cols = getColumnMap(sheet);
    const statusCol = cols["STATUS"];

    if (statusCol) {
      // Apply to Row 2 down to the last possible row
      sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
      console.log(`✅ Fixed STATUS for sheet: ${name}`);
    } else {
      console.log(`⚠️ STATUS column not found in: ${name}`);
    }
  });
  
  ss.toast("✅ Status Dropdowns Repaired (Values Enforced)");
}