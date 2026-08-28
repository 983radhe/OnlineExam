/*******************************************************
 * VIKAS ONLINE EXAM - COMPLETE SERVER CODE
 * Google Apps Script - Code.gs
 *
 * IMPORTANT:
 * 1. Run setupExamSystem() once.
 * 2. Run createAdmin() once if you use the admin system.
 * 3. Run setupDriveFolders() once. It creates the PDF and
 *    camera folders automatically and stores their IDs.
 * 4. Keep your existing sheet data. The code only creates
 *    missing sheets/headers.
 *******************************************************/

const EXAM_ID = "TEST-EXAM";
const PDF_FOLDER_PROPERTY = "PDF_FOLDER_ID";
const CAMERA_FOLDER_PROPERTY = "CAMERA_FOLDER_ID";

/* ======================================================
   1. SHEET SETUP
====================================================== */

function setupExamSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = {
    Students: [
      "Student ID","Name","Mobile","Roll No","Email",
      "Photo URL","Status","Created At"
    ],

    Exams: [
      "Exam ID","Exam Name","Description","Total Questions",
      "Duration","Marks Per Question","Negative Marks",
      "Status","Created At"
    ],

    Questions: [
      "Question ID","Exam ID","Question","Option A","Option B",
      "Option C","Option D","Correct Answer","Marks"
    ],

    Attempts: [
      "Attempt ID","Student ID","Exam ID","Start Time","Submit Time",
      "Status","Score","Percentage","Attempted","Correct","Wrong","Skipped"
    ],

    Answers: [
      "Attempt ID","Question ID","Question No","Student Answer",
      "Correct Answer","Is Correct","Marks","Answered At"
    ],

    LoginLogs: [
      "Log ID","Student ID","Roll No","Login Time",
      "Device","Browser","IP","Status"
    ],

    VerificationCodes: [
      "Code ID","Exam ID","Code","Created At",
      "Expires At","Status","Used By"
    ],

    SecurityLogs: [
      "Log ID","Attempt ID","Student ID","Roll No",
      "Event","Violation Count","Time","Action"
    ],

    CameraLogs: [
      "Capture ID","Attempt ID","Student ID","Roll No",
      "Capture Time","Image URL"
    ],

    Results: [
      "Result ID","Attempt ID","Student ID","Roll No",
      "Student Name","Exam Name","Score","Percentage",
      "Result Status","PDF URL","Created At"
    ],

    PDFReports: [
      "Report ID","Attempt ID","Student ID","Roll No",
      "Student Name","PDF URL","Created At"
    ],

    VerificationRequests: [
      "Request ID","Student ID","Student Name","Roll No",
      "Mobile","Exam ID","Requested At","Status","Verification Code"
    ]
  };

  Object.keys(sheets).forEach(function(name) {
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    const headers = sheets[name];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  });

  SpreadsheetApp.flush();

  return {
    success: true,
    message: "Online Exam System setup completed successfully."
  };
}

/* ======================================================
   2. ADMIN
====================================================== */

function createAdmin() {
  const props = PropertiesService.getScriptProperties();

  props.setProperty("ADMIN_USERNAME", "admin");
  props.setProperty("ADMIN_PASSWORD", "ChangeMe123!");

  return {
    success: true,
    message: "Admin created successfully."
  };
}

/* ======================================================
   3. DRIVE FOLDERS
====================================================== */

function setupDriveFolders() {
  const props = PropertiesService.getScriptProperties();

  let pdfFolder;
  let cameraFolder;

  const pdfId = props.getProperty(PDF_FOLDER_PROPERTY);
  const cameraId = props.getProperty(CAMERA_FOLDER_PROPERTY);

  if (pdfId) {
    try {
      pdfFolder = DriveApp.getFolderById(pdfId);
    } catch (e) {
      pdfFolder = null;
    }
  }

  if (!pdfFolder) {
    pdfFolder = DriveApp.createFolder("Vikas Online Exam - PDF Reports");
    props.setProperty(PDF_FOLDER_PROPERTY, pdfFolder.getId());
  }

  if (cameraId) {
    try {
      cameraFolder = DriveApp.getFolderById(cameraId);
    } catch (e) {
      cameraFolder = null;
    }
  }

  if (!cameraFolder) {
    cameraFolder = DriveApp.createFolder("Vikas Online Exam - Camera Captures");
    props.setProperty(CAMERA_FOLDER_PROPERTY, cameraFolder.getId());
  }

  return {
    success: true,
    pdfFolderId: pdfFolder.getId(),
    cameraFolderId: cameraFolder.getId(),
    message: "Drive folders are ready."
  };
}

/* ======================================================
   4. WEB APP
====================================================== */

function doGet(e) {
  const page =
    e && e.parameter && e.parameter.page
      ? e.parameter.page
      : "index";

  if (page === "admin") {
    return HtmlService
      .createHtmlOutputFromFile("admin")
      .setTitle("Vikas Online Exam - Admin");
  }

  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("Vikas Online Exam");
}

/* ======================================================
   5. STUDENT LOGIN
====================================================== */

function checkStudent(rollNo, mobile) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Students");

  if (!sheet) {
    return {
      success: false,
      message: "Students sheet not found."
    };
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const studentId = String(data[i][0] || "").trim();
    const name = String(data[i][1] || "").trim();
    const studentMobile = String(data[i][2] || "").trim();
    const studentRoll = String(data[i][3] || "").trim();
    const email = String(data[i][4] || "").trim();
    const photoUrl = String(data[i][5] || "").trim();
    const status = String(data[i][6] || "").trim();

    if (
      studentRoll === String(rollNo || "").trim() &&
      studentMobile === String(mobile || "").trim()
    ) {
      if (status.toLowerCase() !== "active") {
        return {
          success: false,
          message: "This student is not currently active."
        };
      }

      const logSheet = getOrCreateSheet("LoginLogs", [
        "Log ID","Student ID","Roll No","Login Time",
        "Device","Browser","IP","Status"
      ]);

      logSheet.appendRow([
        makeId("LOGIN"),
        studentId,
        studentRoll,
        new Date(),
        "Web Browser",
        "Browser",
        "",
        "Success"
      ]);

      return {
        success: true,
        student: {
          id: studentId,
          name: name,
          rollNo: studentRoll,
          mobile: studentMobile,
          email: email,
          photoUrl: photoUrl
        }
      };
    }
  }

  return {
    success: false,
    message: "The Roll Number or Mobile Number is incorrect."
  };
}

/* ======================================================
   6. VERIFICATION REQUEST
====================================================== */

function createVerificationRequest(student) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet("VerificationRequests", [
    "Request ID","Student ID","Student Name","Roll No",
    "Mobile","Exam ID","Requested At","Status","Verification Code"
  ]);

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const data = sheet.getDataRange().getValues();
    const studentId = String(student.id || "").trim();
    const rollNo = String(student.rollNo || "").trim();

    for (let i = 1; i < data.length; i++) {
      const existingStudent = String(data[i][1] || "").trim();
      const existingRoll = String(data[i][3] || "").trim();
      const status = String(data[i][7] || "").trim().toLowerCase();

      if (
        (existingStudent === studentId || existingRoll === rollNo) &&
        status === "pending"
      ) {
        return {
          success: false,
          requestId: String(data[i][0] || ""),
          status: "Pending",
          message: "A verification request is already pending. Please wait for admin approval."
        };
      }
    }

    const requestId = makeId("REQ");

    sheet.appendRow([
      requestId,
      student.id || "",
      student.name || "",
      student.rollNo || "",
      student.mobile || "",
      EXAM_ID,
      new Date(),
      "Pending",
      ""
    ]);

    SpreadsheetApp.flush();

    return {
      success: true,
      requestId: requestId,
      status: "Pending",
      message: "Verification request sent successfully. Please wait for admin approval."
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      message: "Unable to create verification request. Please try again."
    };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ======================================================
   7. ADMIN VERIFICATION REQUESTS
====================================================== */

function getAllVerificationRequests() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("VerificationRequests");

  if (!sheet) {
    return { success: true, requests: [] };
  }

  const data = sheet.getDataRange().getValues();
  const requests = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    requests.push({
      requestId: String(data[i][0] || ""),
      studentId: String(data[i][1] || ""),
      studentName: String(data[i][2] || ""),
      rollNo: String(data[i][3] || ""),
      mobile: String(data[i][4] || ""),
      examId: String(data[i][5] || ""),
      requestedAt: formatDateValue(data[i][6]),
      status: String(data[i][7] || ""),
      verificationCode: String(data[i][8] || "")
    });
  }

  return {
    success: true,
    requests: requests
  };
}

function getPendingVerificationRequests() {
  const result = getAllVerificationRequests();

  if (!result.success) return result;

  result.requests = result.requests.filter(function(r) {
    return r.status.toLowerCase() === "pending";
  });

  return result;
}

function approveVerificationRequest(requestId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("VerificationRequests");
    const codeSheet = getOrCreateSheet("VerificationCodes", [
      "Code ID","Exam ID","Code","Created At",
      "Expires At","Status","Used By"
    ]);

    if (!requestSheet) {
      return {
        success: false,
        message: "VerificationRequests sheet not found."
      };
    }

    const data = requestSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== String(requestId).trim()) continue;

      const status = String(data[i][7] || "").trim();

      if (status.toLowerCase() !== "pending") {
        return {
          success: false,
          message: "This request has already been processed."
        };
      }

      const examId = String(data[i][5] || EXAM_ID).trim();
      const code = generateSixDigitCode();
      const now = new Date();
      const expiresAt = new Date(now.getTime() + 10 * 60 * 1000);

      codeSheet.appendRow([
        makeId("CODE"),
        examId,
        code,
        now,
        expiresAt,
        "Active",
        ""
      ]);

      requestSheet.getRange(i + 1, 8).setValue("Approved");
      requestSheet.getRange(i + 1, 9).setValue(code);

      SpreadsheetApp.flush();

      return {
        success: true,
        code: code,
        expiresAt: formatDateValue(expiresAt)
      };
    }

    return {
      success: false,
      message: "Verification request not found."
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      message: "Unable to approve the verification request."
    };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function rejectVerificationRequest(requestId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("VerificationRequests");

  if (!sheet) {
    return {
      success: false,
      message: "VerificationRequests sheet not found."
    };
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== String(requestId).trim()) continue;

    const status = String(data[i][7] || "").trim();

    if (status.toLowerCase() !== "pending") {
      return {
        success: false,
        message: "This request has already been processed."
      };
    }

    sheet.getRange(i + 1, 8).setValue("Rejected");

    return {
      success: true,
      message: "Verification request rejected."
    };
  }

  return {
    success: false,
    message: "Verification request not found."
  };
}

function regenerateVerificationCode(requestId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("VerificationRequests");
    const codeSheet = ss.getSheetByName("VerificationCodes");

    if (!requestSheet || !codeSheet) {
      return {
        success: false,
        message: "Required sheet not found."
      };
    }

    const requestData = requestSheet.getDataRange().getValues();

    for (let i = 1; i < requestData.length; i++) {
      if (String(requestData[i][0]).trim() !== String(requestId).trim()) continue;

      const status = String(requestData[i][7] || "").trim();

      if (status !== "Approved") {
        return {
          success: false,
          message: "Only approved requests can generate a new code."
        };
      }

      const oldCode = String(requestData[i][8] || "").trim();
      const codeData = codeSheet.getDataRange().getValues();

      for (let j = 1; j < codeData.length; j++) {
        if (String(codeData[j][2] || "").trim() === oldCode) {
          codeSheet.getRange(j + 1, 6).setValue("Replaced");
        }
      }

      const newCode = generateSixDigitCode();
      const now = new Date();
      const expiresAt = new Date(now.getTime() + 10 * 60 * 1000);

      codeSheet.appendRow([
        makeId("CODE"),
        String(requestData[i][5] || EXAM_ID),
        newCode,
        now,
        expiresAt,
        "Active",
        ""
      ]);

      requestSheet.getRange(i + 1, 9).setValue(newCode);

      return {
        success: true,
        code: newCode,
        expiresAt: formatDateValue(expiresAt)
      };
    }

    return {
      success: false,
      message: "Verification request not found."
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      message: "Unable to regenerate the verification code."
    };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function deleteVerificationRequest(requestId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("VerificationRequests");
    const codeSheet = ss.getSheetByName("VerificationCodes");

    if (!requestSheet) {
      return {
        success: false,
        message: "VerificationRequests sheet not found."
      };
    }

    const data = requestSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() !== String(requestId).trim()) continue;

      const oldCode = String(data[i][8] || "").trim();

      if (codeSheet && oldCode) {
        revokeCode(codeSheet, oldCode);
      }

      requestSheet.deleteRow(i + 1);

      return {
        success: true,
        message: "Verification request deleted successfully."
      };
    }

    return {
      success: false,
      message: "Verification request not found."
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      message: "Unable to delete the verification request."
    };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function deleteAllVerificationRequests() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("VerificationRequests");
    const codeSheet = ss.getSheetByName("VerificationCodes");

    if (!requestSheet) {
      return {
        success: true,
        message: "There are no verification requests to delete."
      };
    }

    const lastRow = requestSheet.getLastRow();

    if (lastRow <= 1) {
      return {
        success: true,
        message: "There are no verification requests to delete."
      };
    }

    const data = requestSheet.getRange(2, 1, lastRow - 1, 9).getValues();

    if (codeSheet) {
      const codeData = codeSheet.getDataRange().getValues();

      for (let i = 1; i < codeData.length; i++) {
        const code = String(codeData[i][2] || "").trim();
        const status = String(codeData[i][5] || "").trim();

        if (code && status === "Active") {
          for (let j = 0; j < data.length; j++) {
            if (String(data[j][8] || "").trim() === code) {
              codeSheet.getRange(i + 1, 6).setValue("Revoked");
              break;
            }
          }
        }
      }
    }

    requestSheet.deleteRows(2, lastRow - 1);

    return {
      success: true,
      message: "All verification requests deleted successfully."
    };

  } catch (error) {
    console.error(error);

    return {
      success: false,
      message: "Unable to delete all verification requests."
    };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ======================================================
   8. VERIFY CODE
====================================================== */

function verifyStudentCode(code, studentId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("VerificationCodes");

  if (!sheet) {
    return {
      success: false,
      message: "VerificationCodes sheet not found."
    };
  }

  const data = sheet.getDataRange().getValues();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    const savedCode = String(data[i][2] || "").trim();
    const expiresAt = data[i][4];
    const status = String(data[i][5] || "").trim();

    if (
      savedCode === String(code || "").trim() &&
      status === "Active"
    ) {
      if (!(expiresAt instanceof Date) || expiresAt <= now) {
        sheet.getRange(i + 1, 6).setValue("Expired");

        return {
          success: false,
          message: "This verification code has expired."
        };
      }

      sheet.getRange(i + 1, 6).setValue("Used");
      sheet.getRange(i + 1, 7).setValue(studentId);

      return {
        success: true,
        message: "Verification successful."
      };
    }
  }

  return {
    success: false,
    message: "Invalid verification code."
  };
}

/* ======================================================
   9. EXAM DATA
====================================================== */

function getExamDetails(examId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Exams");

  if (!sheet) {
    return {
      success: false,
      message: "Exams sheet not found."
    };
  }

  const data = sheet.getDataRange().getValues();
  const requestedId = String(examId || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim() !== requestedId) continue;

    const status = String(data[i][7] || "").trim();

    if (status.toLowerCase() !== "active") {
      return {
        success: false,
        message: "This exam is not currently active."
      };
    }

    return {
      success: true,
      exam: {
        examId: requestedId,
        examName: String(data[i][1] || ""),
        description: String(data[i][2] || ""),
        totalQuestions: Number(data[i][3] || 0),
        duration: Number(data[i][4] || 0),
        marksPerQuestion: Number(data[i][5] || 0),
        negativeMarks: Number(data[i][6] || 0),
        status: status,
        createdAt: formatDateValue(data[i][8])
      }
    };
  }

  return {
    success: false,
    message: "Exam not found."
  };
}

function getExamQuestions(examId) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Questions");

  if (!sheet) {
    return {
      success: false,
      message: "Questions sheet not found."
    };
  }

  const data = sheet.getDataRange().getValues();
  const requestedId = String(examId || "").trim();
  const questions = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || "").trim() !== requestedId) continue;

    questions.push({
      questionId: String(data[i][0] || ""),
      question: String(data[i][2] || ""),
      optionA: String(data[i][3] || ""),
      optionB: String(data[i][4] || ""),
      optionC: String(data[i][5] || ""),
      optionD: String(data[i][6] || ""),
      marks: Number(data[i][8] || 0)
    });
  }

  if (!questions.length) {
    return {
      success: false,
      message: "No questions found for this exam."
    };
  }

  shuffleArray(questions);

  return {
    success: true,
    questions: questions,
    totalQuestions: questions.length
  };
}

/* ======================================================
   10. START ATTEMPT
====================================================== */

function startExamAttempt(studentId, examId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attempts = getOrCreateSheet("Attempts", [
    "Attempt ID","Student ID","Exam ID","Start Time","Submit Time",
    "Status","Score","Percentage","Attempted","Correct","Wrong","Skipped"
  ]);

  const attemptId = makeId("ATTEMPT");
  const startTime = new Date();

  attempts.appendRow([
    attemptId,
    studentId || "",
    examId || EXAM_ID,
    startTime,
    "",
    "In Progress",
    "",
    "",
    "",
    "",
    "",
    ""
  ]);

  return {
    success: true,
    attemptId: attemptId,
    startTime: startTime.getTime()
  };
}

/* ======================================================
   11. SAVE SECURITY EVENT
====================================================== */

function logSecurityEvent(eventData) {
  try {
    const student = eventData && eventData.student
      ? eventData.student
      : {};

    const sheet = getOrCreateSheet("SecurityLogs", [
      "Log ID","Attempt ID","Student ID","Roll No",
      "Event","Violation Count","Time","Action"
    ]);

    sheet.appendRow([
      makeId("SEC"),
      eventData.attemptId || "",
      student.id || "",
      student.rollNo || "",
      eventData.event || "Security Event",
      Number(eventData.violationCount || 1),
      new Date(),
      eventData.action || "Logged"
    ]);

    return { success: true };

  } catch (error) {
    console.error("Security logging failed:", error);
    return { success: false };
  }
}

/* ======================================================
   12. CAMERA CAPTURE
====================================================== */

function saveCameraCapture(data) {
  try {
    if (!data || !data.imageData) {
      return {
        success: false,
        message: "Camera image data is missing."
      };
    }

    const props = PropertiesService.getScriptProperties();
    let folder;

    const folderId = props.getProperty(CAMERA_FOLDER_PROPERTY);

    if (folderId) {
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (e) {
        folder = null;
      }
    }

    if (!folder) {
      folder = DriveApp.createFolder("Vikas Online Exam - Camera Captures");
      props.setProperty(CAMERA_FOLDER_PROPERTY, folder.getId());
    }

    const base64 = String(data.imageData).replace(/^data:image\/\w+;base64,/, "");
    const bytes = Utilities.base64Decode(base64);

    const name =
      "CAMERA_" +
      String(data.rollNo || "STUDENT") +
      "_" +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") +
      ".jpg";

    const blob = Utilities.newBlob(bytes, "image/jpeg", name);
    const file = folder.createFile(blob);

    file.setDescription(
      "Vikas Online Exam camera capture | Attempt: " +
      String(data.attemptId || "")
    );

    const url = file.getUrl();

    const sheet = getOrCreateSheet("CameraLogs", [
      "Capture ID","Attempt ID","Student ID","Roll No",
      "Capture Time","Image URL"
    ]);

    sheet.appendRow([
      makeId("CAP"),
      data.attemptId || "",
      data.studentId || "",
      data.rollNo || "",
      new Date(),
      url
    ]);

    return {
      success: true,
      url: url
    };

  } catch (error) {
    console.error("Camera save error:", error);

    return {
      success: false,
      message: "Unable to save camera image."
    };
  }
}

/* ======================================================
   13. FINAL SUBMISSION
   This is the important fixed submission function.
====================================================== */

function submitExam(data) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    if (!data) {
      return {
        success: false,
        message: "Submission data is missing."
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const attempts = ss.getSheetByName("Attempts");
    const answersSheet = ss.getSheetByName("Answers");
    const resultsSheet = ss.getSheetByName("Results");
    const pdfReports = ss.getSheetByName("PDFReports");
    const questionsSheet = ss.getSheetByName("Questions");
    const examsSheet = ss.getSheetByName("Exams");

    if (
      !attempts ||
      !answersSheet ||
      !resultsSheet ||
      !pdfReports ||
      !questionsSheet ||
      !examsSheet
    ) {
      return {
        success: false,
        message: "Required exam sheets are missing. Run setupExamSystem() first."
      };
    }

    const attemptId = String(data.attemptId || "").trim();

    if (!attemptId) {
      return {
        success: false,
        message: "Attempt ID is missing."
      };
    }

    const attemptData = attempts.getDataRange().getValues();
    let attemptRow = -1;
    let attemptStatus = "";

    for (let i = 1; i < attemptData.length; i++) {
      if (String(attemptData[i][0] || "").trim() === attemptId) {
        attemptRow = i + 1;
        attemptStatus = String(attemptData[i][5] || "").trim();
        break;
      }
    }

    if (attemptRow === -1) {
      return {
        success: false,
        message: "Exam attempt was not found."
      };
    }

    if (attemptStatus.toLowerCase() === "submitted") {
      return {
        success: true,
        alreadySubmitted: true,
        message: "This exam has already been submitted successfully."
      };
    }

    const studentId = String(data.studentId || "").trim();
    const rollNo = String(data.rollNo || "").trim();
    const studentName = String(data.studentName || "").trim();
    const examId = String(data.examId || EXAM_ID).trim();
    const reason = String(data.reason || "Manual Submit").trim();
    const answers = data.answers || {};

    const examInfo = getExamInfoForPdf(examsSheet, examId);

    if (!examInfo) {
      return {
        success: false,
        message: "Exam details were not found."
      };
    }

    const questionRows = questionsSheet.getDataRange().getValues();

    const answerRows = [];
    const pdfQuestions = [];

    let attempted = 0;
    let correct = 0;
    let wrong = 0;
    let skipped = 0;
    let score = 0;

    for (let i = 1; i < questionRows.length; i++) {
      const qId = String(questionRows[i][0] || "").trim();
      const qExamId = String(questionRows[i][1] || "").trim();

      if (qExamId !== examId) continue;

      const questionNo = pdfQuestions.length + 1;
      const questionText = String(questionRows[i][2] || "");
      const optionA = String(questionRows[i][3] || "");
      const optionB = String(questionRows[i][4] || "");
      const optionC = String(questionRows[i][5] || "");
      const optionD = String(questionRows[i][6] || "");
      const correctAnswer = String(questionRows[i][7] || "").trim();
      const marks = Number(questionRows[i][8] || examInfo.marksPerQuestion || 0);

      const studentAnswer = String(answers[qId] || "").trim();

      let isCorrect = false;
      let awardedMarks = 0;

      if (studentAnswer) {
        attempted++;

        if (studentAnswer === correctAnswer) {
          correct++;
          isCorrect = true;
          awardedMarks = marks;
          score += marks;
        } else {
          wrong++;
          awardedMarks = -Math.abs(examInfo.negativeMarks || 0);
          score += awardedMarks;
        }
      } else {
        skipped++;
      }

      answerRows.push([
        attemptId,
        qId,
        questionNo,
        studentAnswer,
        correctAnswer,
        isCorrect ? "Yes" : "No",
        awardedMarks,
        new Date()
      ]);

      pdfQuestions.push({
        no: questionNo,
        question: questionText,
        optionA: optionA,
        optionB: optionB,
        optionC: optionC,
        optionD: optionD,
        studentAnswer: studentAnswer || "Not Answered",
        correctAnswer: correctAnswer,
        isCorrect: isCorrect,
        marks: awardedMarks
      });
    }

    const totalMarks = calculateTotalMarks(pdfQuestions);
    const percentage = totalMarks > 0
      ? Math.max(0, (score / totalMarks) * 100)
      : 0;

    if (answerRows.length) {
      answersSheet
        .getRange(
          answersSheet.getLastRow() + 1,
          1,
          answerRows.length,
          answerRows[0].length
        )
        .setValues(answerRows);
    }

    const submitTime = new Date();

    attempts.getRange(attemptRow, 5).setValue(submitTime);
    attempts.getRange(attemptRow, 6).setValue("Submitted");
    attempts.getRange(attemptRow, 7).setValue(score);
    attempts.getRange(attemptRow, 8).setValue(percentage);
    attempts.getRange(attemptRow, 9).setValue(attempted);
    attempts.getRange(attemptRow, 10).setValue(correct);
    attempts.getRange(attemptRow, 11).setValue(wrong);
    attempts.getRange(attemptRow, 12).setValue(skipped);

    /* ---------- CREATE PDF ---------- */

    const pdfResult = createExamPdf({
      attemptId: attemptId,
      studentId: studentId,
      studentName: studentName,
      rollNo: rollNo,
      examId: examId,
      examName: examInfo.examName,
      startTime: attemptData[attemptRow - 1][3],
      submitTime: submitTime,
      reason: reason,
      score: score,
      percentage: percentage,
      attempted: attempted,
      correct: correct,
      wrong: wrong,
      skipped: skipped,
      totalMarks: totalMarks,
      questions: pdfQuestions
    });

    if (!pdfResult.success) {
      return {
        success: false,
        message:
          "Exam was submitted, but the PDF could not be created: " +
          pdfResult.message
      };
    }

    const resultId = makeId("RESULT");

    resultsSheet.appendRow([
      resultId,
      attemptId,
      studentId,
      rollNo,
      studentName,
      examInfo.examName,
      score,
      percentage,
      "Submitted",
      pdfResult.url,
      submitTime
    ]);

    pdfReports.appendRow([
      makeId("PDF"),
      attemptId,
      studentId,
      rollNo,
      studentName,
      pdfResult.url,
      submitTime
    ]);

    SpreadsheetApp.flush();

    return {
      success: true,
      alreadySubmitted: false,
      attemptId: attemptId,
      resultId: resultId,
      status: "Submitted",
      submittedAt: formatDateValue(submitTime),
      pdfUrl: pdfResult.url,
      message: "Exam submitted successfully."
    };

  } catch (error) {
    console.error("SUBMIT EXAM ERROR:", error);

    return {
      success: false,
      message:
        "Unable to complete exam submission. Please contact the administrator."
    };

  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ======================================================
   14. PDF CREATION
====================================================== */

function createExamPdf(report) {
  try {
    const props = PropertiesService.getScriptProperties();
    let folder;

    const folderId = props.getProperty(PDF_FOLDER_PROPERTY);

    if (folderId) {
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (e) {
        folder = null;
      }
    }

    if (!folder) {
      folder = DriveApp.createFolder("Vikas Online Exam - PDF Reports");
      props.setProperty(PDF_FOLDER_PROPERTY, folder.getId());
    }

    const html = buildPdfHtml(report);

    const blob = Utilities
      .newBlob(html, "text/html", "exam_report.html");

    const tempFile = folder.createFile(blob);

    const pdfBlob = tempFile
      .getAs(MimeType.PDF)
      .setName(
        "Vikas_Online_Exam_" +
        safeFileName(report.rollNo || report.studentName || report.attemptId) +
        "_" +
        safeFileName(report.attemptId) +
        ".pdf"
      );

    const pdfFile = folder.createFile(pdfBlob);

    try {
      tempFile.setTrashed(true);
    } catch (e) {}

    return {
      success: true,
      url: pdfFile.getUrl(),
      fileId: pdfFile.getId(),
      name: pdfFile.getName()
    };

  } catch (error) {
    console.error("PDF creation error:", error);

    return {
      success: false,
      message: String(error.message || error)
    };
  }
}

function buildPdfHtml(report) {
  const q = report.questions || [];

  let questionHtml = "";

  q.forEach(function(item) {
    questionHtml += `
      <div class="question">
        <div class="qtitle">
          Q${item.no}. ${escapeHtmlServer(item.question)}
        </div>

        <div>A. ${escapeHtmlServer(item.optionA)}</div>
        <div>B. ${escapeHtmlServer(item.optionB)}</div>
        <div>C. ${escapeHtmlServer(item.optionC)}</div>
        <div>D. ${escapeHtmlServer(item.optionD)}</div>

        <div class="answer">
          Student Answer:
          <b>${escapeHtmlServer(item.studentAnswer)}</b>
        </div>

        <div>
          Correct Answer:
          <b>${escapeHtmlServer(item.correctAnswer)}</b>
        </div>

        <div>
          Status:
          <b>${item.isCorrect ? "Correct" : (item.studentAnswer === "Not Answered" ? "Skipped" : "Wrong")}</b>
        </div>

        <div>
          Marks:
          <b>${Number(item.marks || 0)}</b>
        </div>
      </div>
    `;
  });

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body{
  font-family:Arial,sans-serif;
  padding:28px;
  color:#111827;
  font-size:12px;
}
h1{
  text-align:center;
  margin-bottom:5px;
  color:#1e3a8a;
}
h2{
  margin-top:25px;
  color:#1e3a8a;
}
.info{
  border:1px solid #d1d5db;
  padding:12px;
  margin:15px 0;
}
.info table{
  width:100%;
  border-collapse:collapse;
}
.info td{
  padding:5px;
  border-bottom:1px solid #e5e7eb;
}
.summary{
  width:100%;
  border-collapse:collapse;
  margin:15px 0;
}
.summary td{
  border:1px solid #d1d5db;
  padding:8px;
}
.question{
  border:1px solid #d1d5db;
  padding:10px;
  margin:10px 0;
  page-break-inside:avoid;
}
.qtitle{
  font-weight:bold;
  margin-bottom:8px;
}
.answer{
  margin-top:7px;
}
.footer{
  margin-top:25px;
  text-align:center;
  font-size:10px;
  color:#6b7280;
}
</style>
</head>
<body>

<h1>Vikas Online Exam</h1>
<div style="text-align:center;">Examination Submission Report</div>

<div class="info">
<table>
<tr><td><b>Student Name</b></td><td>${escapeHtmlServer(report.studentName)}</td></tr>
<tr><td><b>Roll Number</b></td><td>${escapeHtmlServer(report.rollNo)}</td></tr>
<tr><td><b>Student ID</b></td><td>${escapeHtmlServer(report.studentId)}</td></tr>
<tr><td><b>Exam Name</b></td><td>${escapeHtmlServer(report.examName)}</td></tr>
<tr><td><b>Exam ID</b></td><td>${escapeHtmlServer(report.examId)}</td></tr>
<tr><td><b>Attempt ID</b></td><td>${escapeHtmlServer(report.attemptId)}</td></tr>
<tr><td><b>Submission Type</b></td><td>${escapeHtmlServer(report.reason)}</td></tr>
<tr><td><b>Start Time</b></td><td>${escapeHtmlServer(formatDateValue(report.startTime))}</td></tr>
<tr><td><b>Submit Time</b></td><td>${escapeHtmlServer(formatDateValue(report.submitTime))}</td></tr>
</table>
</div>

<table class="summary">
<tr>
<td><b>Score</b><br>${Number(report.score || 0)}</td>
<td><b>Total Marks</b><br>${Number(report.totalMarks || 0)}</td>
<td><b>Percentage</b><br>${Number(report.percentage || 0).toFixed(2)}%</td>
</tr>
<tr>
<td><b>Attempted</b><br>${Number(report.attempted || 0)}</td>
<td><b>Correct</b><br>${Number(report.correct || 0)}</td>
<td><b>Wrong</b><br>${Number(report.wrong || 0)}</td>
</tr>
<tr>
<td><b>Skipped</b><br>${Number(report.skipped || 0)}</td>
<td colspan="2"><b>Status</b><br>Submitted</td>
</tr>
</table>

<h2>Question-wise Details</h2>

${questionHtml}

<div class="footer">
Generated automatically by Vikas Online Exam
</div>

</body>
</html>
`;
}

/* ======================================================
   15. HELPERS
====================================================== */

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0 && headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function makeId(prefix) {
  return prefix + "-" +
    Utilities.getUuid()
      .replace(/-/g, "")
      .substring(0, 12)
      .toUpperCase();
}

function generateSixDigitCode() {
  return String(
    Math.floor(100000 + Math.random() * 900000)
  );
}

function formatDateValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm:ss"
    );
  }

  return String(value || "");
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));

    const temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }

  return array;
}

function revokeCode(sheet, code) {
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2] || "").trim() === String(code).trim()) {
      sheet.getRange(i + 1, 6).setValue("Revoked");
    }
  }
}

function getExamInfoForPdf(sheet, examId) {
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === String(examId).trim()) {
      return {
        examName: String(data[i][1] || ""),
        marksPerQuestion: Number(data[i][5] || 0),
        negativeMarks: Number(data[i][6] || 0)
      };
    }
  }

  return null;
}

function calculateTotalMarks(questions) {
  let total = 0;

  questions.forEach(function(q) {
    total += Math.max(
      Number(q.marks || 0),
      0
    );
  });

  return total;
}

function safeFileName(value) {
  return String(value || "report")
    .replace(/[\\\/:*?"<>|#%{}[\]]/g, "_")
    .replace(/\s+/g, "_")
    .substring(0, 80);
}

function escapeHtmlServer(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/* Optional manual test */
function testDriveFolders() {
  Logger.log(setupDriveFolders());
}
