// Google Apps Script for BlackHomeschoolers Check-in System (Container-bound)
// Updated to support: Member, Non-Member, Facilitator, Visitor, Visiting Facilitator

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Server-side functions that can be called from HTML
function testConnection() {
  return { success: true, message: "Connection successful" };
}

function handleCheckInOut(data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName("CheckInLog");

    if (!sheet) {
      sheet = spreadsheet.insertSheet("CheckInLog");
      sheet
        .getRange(1, 1, 1, 7)
        .setValues([
          [
            "ID",
            "Name",
            "Type",
            "Action",
            "Location",
            "Timestamp",
            "Input Method",
          ],
        ]);

      // Format headers
      const headerRange = sheet.getRange(1, 1, 1, 7);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4285F4");
      headerRange.setFontColor("white");
    }

    const timestamp = new Date();

    // Determine input method based on person type and ID format
    let inputMethod = "ID Card";
    if (data.type === "Visitor" || data.type === "Visiting Facilitator") {
      // Check if ID looks auto-generated (contains underscore and timestamp)
      if (data.id && data.id.includes("_") && /\d{13}/.test(data.id)) {
        inputMethod = "Name Entry (Auto-ID)";
      } else if (!data.id || data.id === data.name) {
        inputMethod = "Name Entry";
      }
    }

    sheet.appendRow([
      data.id || "",
      data.name || "",
      data.type || "",
      data.actionType || "",
      data.location || "",
      timestamp,
      inputMethod,
    ]);

    // Auto-resize columns for better readability
    sheet.autoResizeColumns(1, 7);

    return { success: true, message: "Check-in/out logged successfully" };
  } catch (error) {
    console.error("Error in handleCheckInOut:", error);
    return { success: false, error: error.toString() };
  }
}

function addMember(data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName("Members");

    if (!sheet) {
      sheet = spreadsheet.insertSheet("Members");
      sheet
        .getRange(1, 1, 1, 4)
        .setValues([["ID", "Name", "Type", "Date Added"]]);

      // Format headers
      const headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#34A853");
      headerRange.setFontColor("white");
    }

    // Validate person type for Members sheet
    const validMemberTypes = ["Member", "Facilitator"];
    if (!validMemberTypes.includes(data.type)) {
      return {
        success: false,
        error: `Invalid type for Members sheet: ${data.type}. Use addNonMember for other types.`,
      };
    }

    // Check if ID already exists across all sheets
    const existingMember = findPersonById(data.id);
    if (existingMember.found) {
      return {
        success: false,
        error: `ID ${data.id} already exists as ${existingMember.type} in ${existingMember.sheet}`,
      };
    }

    // Validate required fields
    if (!data.id || !data.name) {
      return {
        success: false,
        error: "ID and Name are required for Members and Facilitators",
      };
    }

    const timestamp = new Date();
    sheet.appendRow([data.id, data.name, data.type, timestamp]);

    sheet.autoResizeColumns(1, 4);

    return { success: true, message: `${data.type} added successfully` };
  } catch (error) {
    console.error("Error in addMember:", error);
    return { success: false, error: error.toString() };
  }
}

function addNonMember(data) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName("NonMembers");

    if (!sheet) {
      sheet = spreadsheet.insertSheet("NonMembers");
      sheet
        .getRange(1, 1, 1, 5)
        .setValues([["ID", "Name", "Type", "Has ID Card", "Date Added"]]);

      // Format headers
      const headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#FF9800");
      headerRange.setFontColor("white");
    }

    // Validate person type for NonMembers sheet
    const validNonMemberTypes = [
      "Non-Member",
      "Visitor",
      "Visiting Facilitator",
    ];
    if (!validNonMemberTypes.includes(data.type)) {
      return {
        success: false,
        error: `Invalid type for NonMembers sheet: ${data.type}. Use addMember for Members/Facilitators.`,
      };
    }

    // Handle ID requirements based on type
    let finalId = data.id;
    let hasIdCard = true;

    if (data.type === "Visitor" || data.type === "Visiting Facilitator") {
      hasIdCard = false;

      if (!finalId) {
        // Generate unique ID for visitors without cards
        finalId =
          data.name.replace(/\s+/g, "").toLowerCase() + "_" + Date.now();
      }
    } else if (data.type === "Non-Member") {
      // Non-Members must have ID cards
      if (!finalId) {
        return {
          success: false,
          error: "ID card number is required for Non-Members",
        };
      }
      hasIdCard = true;
    }

    // Validate required fields
    if (!finalId || !data.name) {
      return {
        success: false,
        error: "Name is required, and ID is required for Non-Members",
      };
    }

    // Check if ID already exists across all sheets
    const existingPerson = findPersonById(finalId);
    if (existingPerson.found) {
      return {
        success: false,
        error: `ID ${finalId} already exists as ${existingPerson.type} in ${existingPerson.sheet}`,
      };
    }

    const timestamp = new Date();
    sheet.appendRow([
      finalId,
      data.name,
      data.type,
      hasIdCard ? "Yes" : "No",
      timestamp,
    ]);

    sheet.autoResizeColumns(1, 5);

    return {
      success: true,
      message: `${data.type} ${data.name} added successfully`,
    };
  } catch (error) {
    console.error("Error in addNonMember:", error);
    return { success: false, error: error.toString() };
  }
}

// Helper function to find a person by ID across all sheets
function findPersonById(id) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Check Members sheet
    const membersSheet = spreadsheet.getSheetByName("Members");
    if (membersSheet) {
      const membersData = membersSheet.getDataRange().getValues();
      for (let i = 1; i < membersData.length; i++) {
        if (
          membersData[i][0] &&
          membersData[i][0].toString() === id.toString()
        ) {
          return {
            found: true,
            sheet: "Members",
            type: membersData[i][2] || "Member",
            name: membersData[i][1] || "",
          };
        }
      }
    }

    // Check NonMembers sheet
    const nonMembersSheet = spreadsheet.getSheetByName("NonMembers");
    if (nonMembersSheet) {
      const nonMembersData = nonMembersSheet.getDataRange().getValues();
      for (let i = 1; i < nonMembersData.length; i++) {
        if (
          nonMembersData[i][0] &&
          nonMembersData[i][0].toString() === id.toString()
        ) {
          return {
            found: true,
            sheet: "NonMembers",
            type: nonMembersData[i][2] || "Non-Member",
            name: nonMembersData[i][1] || "",
          };
        }
      }
    }

    return { found: false };
  } catch (error) {
    console.error("Error in findPersonById:", error);
    return { found: false };
  }
}

function loadAllMembers() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Load Members (Members and Facilitators with ID cards)
    let membersSheet = spreadsheet.getSheetByName("Members");
    const members = {};

    if (membersSheet) {
      const membersData = membersSheet.getDataRange().getValues();
      for (let i = 1; i < membersData.length; i++) {
        if (membersData[i][0]) {
          const id = membersData[i][0].toString();
          members[id] = {
            name: membersData[i][1] || "",
            type: membersData[i][2] || "Member",
          };
        }
      }
    } else {
      membersSheet = spreadsheet.insertSheet("Members");
      membersSheet
        .getRange(1, 1, 1, 4)
        .setValues([["ID", "Name", "Type", "Date Added"]]);

      const headerRange = membersSheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#34A853");
      headerRange.setFontColor("white");
    }

    // Load Non-Members (Non-Members, Visitors, Visiting Facilitators)
    let nonMembersSheet = spreadsheet.getSheetByName("NonMembers");
    const nonMembers = {};

    if (nonMembersSheet) {
      const nonMembersData = nonMembersSheet.getDataRange().getValues();
      for (let i = 1; i < nonMembersData.length; i++) {
        if (nonMembersData[i][0]) {
          const id = nonMembersData[i][0].toString();
          nonMembers[id] = {
            name: nonMembersData[i][1] || "",
            type: nonMembersData[i][2] || "Non-Member",
          };
        }
      }
    } else {
      nonMembersSheet = spreadsheet.insertSheet("NonMembers");
      nonMembersSheet
        .getRange(1, 1, 1, 5)
        .setValues([["ID", "Name", "Type", "Has ID Card", "Date Added"]]);

      const headerRange = nonMembersSheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#FF9800");
      headerRange.setFontColor("white");
    }

    const totalMembers = Object.keys(members).length;
    const totalNonMembers = Object.keys(nonMembers).length;

    return {
      success: true,
      members: members,
      nonMembers: nonMembers,
      summary: {
        totalMembers: totalMembers,
        totalNonMembers: totalNonMembers,
        totalPeople: totalMembers + totalNonMembers,
      },
    };
  } catch (error) {
    console.error("Error in loadAllMembers:", error);
    return { success: false, error: error.toString() };
  }
}

function setupSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Create CheckInLog sheet with enhanced tracking
    let checkInSheet = spreadsheet.getSheetByName("CheckInLog");
    if (!checkInSheet) {
      checkInSheet = spreadsheet.insertSheet("CheckInLog");
      checkInSheet
        .getRange(1, 1, 1, 7)
        .setValues([
          [
            "ID",
            "Name",
            "Type",
            "Action",
            "Location",
            "Timestamp",
            "Input Method",
          ],
        ]);

      const headerRange = checkInSheet.getRange(1, 1, 1, 7);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4285F4");
      headerRange.setFontColor("white");
    }

    // Create Members sheet (Members and Facilitators with ID cards)
    let membersSheet = spreadsheet.getSheetByName("Members");
    if (!membersSheet) {
      membersSheet = spreadsheet.insertSheet("Members");
      membersSheet
        .getRange(1, 1, 1, 4)
        .setValues([["ID", "Name", "Type", "Date Added"]]);

      const headerRange = membersSheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#34A853");
      headerRange.setFontColor("white");

      // Add sample data comment
      membersSheet.getRange(2, 1).setNote("Sample: 12345");
      membersSheet.getRange(2, 2).setNote("Sample: John Doe");
      membersSheet.getRange(2, 3).setNote("Sample: Member or Facilitator");
    }

    // Create NonMembers sheet (Non-Members, Visitors, Visiting Facilitators)
    let nonMembersSheet = spreadsheet.getSheetByName("NonMembers");
    if (!nonMembersSheet) {
      nonMembersSheet = spreadsheet.insertSheet("NonMembers");
      nonMembersSheet
        .getRange(1, 1, 1, 5)
        .setValues([["ID", "Name", "Type", "Has ID Card", "Date Added"]]);

      const headerRange = nonMembersSheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#FF9800");
      headerRange.setFontColor("white");

      // Add sample data comments
      nonMembersSheet
        .getRange(2, 1)
        .setNote("Sample: 67890 or auto-generated for visitors");
      nonMembersSheet.getRange(2, 2).setNote("Sample: Jane Smith");
      nonMembersSheet
        .getRange(2, 3)
        .setNote("Sample: Non-Member, Visitor, or Visiting Facilitator");
      nonMembersSheet
        .getRange(2, 4)
        .setNote("Yes for Non-Members, No for Visitors");
    }

    // Auto-resize all columns in all sheets
    [checkInSheet, membersSheet, nonMembersSheet].forEach((sheet) => {
      if (sheet) {
        const maxCols = sheet.getLastColumn();
        if (maxCols > 0) {
          sheet.autoResizeColumns(1, maxCols);
        }
      }
    });

    return "Setup completed successfully - All sheets created with proper formatting";
  } catch (error) {
    console.error("Error in setupSheets:", error);
    return "Setup failed: " + error.toString();
  }
}

// Additional utility function to get summary statistics
function getSystemStats() {
  try {
    const result = loadAllMembers();
    if (!result.success) {
      return result;
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const checkInSheet = spreadsheet.getSheetByName("CheckInLog");

    let totalCheckIns = 0;
    let todayCheckIns = 0;

    if (checkInSheet) {
      const logData = checkInSheet.getDataRange().getValues();
      const today = new Date();
      today.setHours(0, 0, 0, 0);

      for (let i = 1; i < logData.length; i++) {
        if (logData[i][5]) {
          // Timestamp column
          totalCheckIns++;
          const logDate = new Date(logData[i][5]);
          logDate.setHours(0, 0, 0, 0);
          if (logDate.getTime() === today.getTime()) {
            todayCheckIns++;
          }
        }
      }
    }

    return {
      success: true,
      stats: {
        ...result.summary,
        totalCheckIns: totalCheckIns,
        todayCheckIns: todayCheckIns,
      },
    };
  } catch (error) {
    console.error("Error in getSystemStats:", error);
    return { success: false, error: error.toString() };
  }
}
