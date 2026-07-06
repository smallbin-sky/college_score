/**
 * 大學多課程評分系統 Google Apps Script 後端 (GAS)
 * 
 * 【部署說明】
 * 1. 在您的 Google 雲端硬碟建立一個新的「Google 試算表 (Google Sheets)」。
 * 2. 點選上方選單的「擴充功能」->「Apps Script」。
 * 3. 將本程式碼 (Code.gs) 的所有內容完整複製，覆蓋貼上到 GAS 編輯器中。
 * 4. 點選編輯器右上角的「部署」->「新增部署」。
 * 5. 部署類型選擇「網頁應用程式 (Web App)」。
 * 6. 設定以下部署內容：
 *    - 說明：大學評分系統雲端 API
 *    - 執行身分：我 (您的 Google 帳號)
 *    - 誰有權限存取：任何人 (Anyone)
 * 7. 點選「部署」，並在彈出的權限視窗中，點選您的帳號並授予授權（點選「進階」->「前往未命名的專案 (不安全)」以核准權限）。
 * 8. 複製部署成功後產生的「網頁應用程式 URL」，並貼回網頁的【雲端同步】面板中即可開始同步！
 */

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "大學多課程評分系統 GAS API 連線成功！系統運作正常。"
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // 設定 CORS 與處理預檢請求
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "test") {
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "連線成功！已成功與您的 Google 試算表進行串接。"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "upload") {
      // 1. 將完整的 JSON 資料庫備份至隱藏的 __CONFIG__ 工作表中，供還原使用
      var configSheet = spreadsheet.getSheetByName("__CONFIG__");
      if (!configSheet) {
        configSheet = spreadsheet.insertSheet("__CONFIG__");
        configSheet.hideSheet(); // 隱藏此工作表，避免使用者誤刪
      }
      configSheet.clear();
      configSheet.getRange(1, 1).setValue(JSON.stringify(requestData.payload));
      
      // 2. 解析資料並動態為每門課程建立/更新美化的工作表
      updateReports(spreadsheet, requestData.payload.courses);
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "資料已成功備份至雲端，並自動更新試算表報表！"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === "download") {
      var configSheet = spreadsheet.getSheetByName("__CONFIG__");
      if (!configSheet) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "雲端尚無備份資料！請先從網頁上傳備份。"
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      var jsonStr = configSheet.getRange(1, 1).getValue();
      if (!jsonStr) {
        return ContentService.createTextOutput(JSON.stringify({
          status: "error",
          message: "雲端備份資料為空！"
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        payload: JSON.parse(jsonStr)
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "未知的 action 指令"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "後端執行出錯: " + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 更新所有課程在 Sheets 中的視覺化報表
 */
function updateReports(spreadsheet, courses) {
  var activeCourseNames = [];
  for (var courseId in courses) {
    activeCourseNames.push(courses[courseId].name);
  }
  
  // 清理已經在網頁端被刪除的課程對應工作表
  var sheets = spreadsheet.getSheets();
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    if (name === "__CONFIG__") return;
    
    var isGradeSheet = name.indexOf("_成績總表") !== -1;
    var isAttSheet = name.indexOf("_出缺席表") !== -1;
    if (isGradeSheet || isAttSheet) {
      var courseName = name.replace("_成績總表", "").replace("_出缺席表", "");
      if (activeCourseNames.indexOf(courseName) === -1) {
        try {
          spreadsheet.deleteSheet(sheet);
        } catch(e) {
          // 防止最後一個工作表被刪除而出錯
        }
      }
    }
  });
  
  // 逐一更新各課程工作表
  for (var courseId in courses) {
    var course = courses[courseId];
    
    // A. 更新或建立成績總表
    var gradeSheetName = course.name + "_成績總表";
    var gradeSheet = spreadsheet.getSheetByName(gradeSheetName);
    if (!gradeSheet) {
      gradeSheet = spreadsheet.insertSheet(gradeSheetName);
    }
    gradeSheet.clear();
    writeGradesToSheet(gradeSheet, course);
    
    // B. 更新或建立出缺席表
    var attSheetName = course.name + "_出缺席表";
    var attSheet = spreadsheet.getSheetByName(attSheetName);
    if (!attSheet) {
      attSheet = spreadsheet.insertSheet(attSheetName);
    }
    attSheet.clear();
    writeAttendanceToSheet(attSheet, course);
  }
}

/**
 * 寫入與美化成績總表
 */
function writeGradesToSheet(sheet, course) {
  var weights = course.weights || { att: 20, pres: 40, part: 40 };
  var projects = course.projects || [];
  var students = course.students || [];
  var peerEvalMode = course.peerEvalMode || "none";
  
  // 1. 設置標題列與樣式
  sheet.getRange(1, 1).setValue(course.name + " 學期加權成績大表")
    .setFontSize(16)
    .setFontWeight("bold")
    .setFontColor("#1e3a8a");
  
  var maxCol = 16 + projects.length;
  sheet.getRange(1, 1, 1, maxCol).merge();
  
  // 2. 設置公式與權重描述
  var formulaText = "★ 計分公式：學期總加權成績 = (出席得分 × " + weights.att + "%) + (報告最終實得 × " + weights.pres + "%) + (發言得分 × " + weights.part + "%)";
  if (peerEvalMode === "average") {
    formulaText += " | [小組互評機制：已啟用「平均貢獻度調幅」，報告最終實得 = 報告平均基分 × 個人貢獻係數 (限制於 0.80 ~ 1.20 之間)]";
  } else {
    formulaText += " | [小組互評機制：未啟用 (互評僅作參考，報告分數直接繼承小組基分)]";
  }
  sheet.getRange(2, 1).setValue(formulaText)
    .setFontSize(9)
    .setFontColor("#4b5563")
    .setFontStyle("italic");
  sheet.getRange(2, 1, 1, maxCol).merge();
  
  // 3. 定義欄位 Header
  var headers = ["No.", "班級", "學號", "姓名", "分組", "歷史出席率", "出席評估分 (" + weights.att + "%)"];
  projects.forEach(function(p) {
    headers.push("報告原始:" + p);
  });
  headers.push("報告項目平均基分", "團隊互評得分", "個人貢獻係數", "報告最終實得", "報告加權分 (" + weights.pres + "%)", "發言次數", "發言評估分 (" + weights.part + "%)", "學期加權總成績");
  
  var headerRange = sheet.getRange(4, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground("#1e3a8a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 28);
  
  if (students.length === 0) {
    sheet.getRange(5, 1).setValue("當前課程尚無學生成冊資料。").setFontColor("#9ca3af").setFontStyle("italic");
    return;
  }
  
  // 4. 計算各小組互評平均
  var groupPeerAverages = {};
  var uniqueGroups = [];
  students.forEach(function(s) {
    if (s.group && s.group !== "未分組" && uniqueGroups.indexOf(s.group) === -1) {
      uniqueGroups.push(s.group);
    }
  });
  
  uniqueGroups.forEach(function(g) {
    var members = students.filter(function(s) { return s.group === g; });
    var sum = 0;
    members.forEach(function(m) {
      if (!m.peerEval) m.peerEval = { contribution: 80, communication: 80, punctuality: 80 };
      sum += (m.peerEval.contribution + m.peerEval.communication + m.peerEval.punctuality) / 3;
    });
    groupPeerAverages[g] = members.length > 0 ? (sum / members.length) : 80;
  });
  
  // 5. 整理與寫入學生成績列
  var rowsData = [];
  var sortedStudents = students.slice().sort(function(a, b) {
    return (a.no || 0) - (b.no || 0);
  });
  
  sortedStudents.forEach(function(s, idx) {
    // 出席率
    var totalSlots = 0, attendedSlots = 0;
    if (s.attendance) {
      for (var date in s.attendance) {
        totalSlots++;
        var status = s.attendance[date];
        if (status === "出席") attendedSlots++;
        else if (status === "遲到") attendedSlots += 0.7;
      }
    }
    var attRate = totalSlots > 0 ? (attendedSlots / totalSlots) * 100 : 100;
    
    // 報告平均
    var pSum = 0, pCount = 0;
    var pScores = [];
    projects.forEach(function(p) {
      var sc = s.presentationScores ? s.presentationScores[p] : null;
      if (sc !== undefined && sc !== null && sc !== "") {
        pSum += sc;
        pCount++;
        pScores.push(sc);
      } else {
        pScores.push("");
      }
    });
    var projectBaseAvg = pCount > 0 ? (pSum / pCount) : 0;
    
    // 互評與係數
    if (!s.peerEval) s.peerEval = { contribution: 80, communication: 80, punctuality: 80 };
    var personalPeerScore = (s.peerEval.contribution + s.peerEval.communication + s.peerEval.punctuality) / 3;
    
    var finalProjectScore = projectBaseAvg;
    var appliedCoef = 1.0;
    if (peerEvalMode === "average") {
      if (s.group && s.group !== "未分組" && groupPeerAverages[s.group] > 0) {
        var rawCoef = personalPeerScore / groupPeerAverages[s.group];
        appliedCoef = Math.min(1.2, Math.max(0.8, rawCoef));
        finalProjectScore = Math.min(100, projectBaseAvg * appliedCoef);
      }
    }
    var projectWeighted = finalProjectScore * (weights.pres / 100);
    
    // 發言
    var bonusCount = s.bonusCount || 0;
    var partScore = Math.min(100, bonusCount * 5);
    
    // 總成績
    var finalGrade = (attRate * (weights.att / 100)) + (finalProjectScore * (weights.pres / 100)) + (partScore * (weights.part / 100));
    
    var row = [
      s.no || (idx + 1),
      s.class || "",
      s.id || "",
      s.name || "",
      s.group || "未分組",
      Math.round(attRate) + "%",
      Math.round(attRate)
    ];
    
    pScores.forEach(function(score) {
      row.push(score === "" ? "-" : score);
    });
    
    row.push(
      Math.round(projectBaseAvg),
      Math.round(personalPeerScore),
      peerEvalMode === "average" && s.group !== "未分組" ? Number(appliedCoef.toFixed(2)) : "-",
      Math.round(finalProjectScore),
      Number(projectWeighted.toFixed(1)),
      bonusCount,
      partScore,
      Number(finalGrade.toFixed(1))
    );
    
    rowsData.push(row);
  });
  
  var dataRange = sheet.getRange(5, 1, rowsData.length, headers.length);
  dataRange.setValues(rowsData)
    .setFontSize(9)
    .setVerticalAlignment("middle");
  
  // 排版細節：前 5 欄對齊靠左，其餘對齊居中
  sheet.getRange(5, 1, rowsData.length, 5).setHorizontalAlignment("left");
  sheet.getRange(5, 6, rowsData.length, headers.length - 5).setHorizontalAlignment("center");
  // 總加權成績靠右並加粗
  sheet.getRange(5, headers.length, rowsData.length, 1).setHorizontalAlignment("right").setFontWeight("bold");
  
  // 劃上格線
  dataRange.setBorder(true, true, true, true, true, true, "#e5e7eb", SpreadsheetApp.BorderStyle.SOLID);
  
  // 條件格式化：總成績不及格 (<60) 紅色高亮
  var finalGradesRange = sheet.getRange(5, headers.length, rowsData.length, 1);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(60)
    .setBackground("#fee2e2")
    .setFontColor("#b91c1c")
    .setRanges([finalGradesRange])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
  
  // 自動適應欄寬，並加上邊距
  for (var col = 1; col <= headers.length; col++) {
    sheet.autoResizeColumn(col);
    var currentWidth = sheet.getColumnWidth(col);
    sheet.setColumnWidth(col, currentWidth + 12);
  }
}

/**
 * 寫入與美化出缺席明細表
 */
function writeAttendanceToSheet(sheet, course) {
  var students = course.students || [];
  
  // 1. 設置大標題
  sheet.getRange(1, 1).setValue(course.name + " 學生出缺席點名大表")
    .setFontSize(16)
    .setFontWeight("bold")
    .setFontColor("#1e3a8a");
  
  // 2. 蒐集所有不重複且排序過的點名日期
  var datesSet = [];
  students.forEach(function(s) {
    if (s.attendance) {
      for (var d in s.attendance) {
        if (datesSet.indexOf(d) === -1) {
          datesSet.push(d);
        }
      }
    }
  });
  datesSet.sort(function(a, b) {
    return new Date(a) - new Date(b);
  });
  
  var maxCol = 5 + datesSet.length;
  sheet.getRange(1, 1, 1, maxCol).merge();
  
  // 3. 說明列
  sheet.getRange(2, 1).setValue("★ 點名狀態著色說明：出席 (綠色)、遲到 (黃色)、缺席 (紅色) | 點名歷史會即時且完整呈現在下方表格中。")
    .setFontSize(9)
    .setFontColor("#4b5563")
    .setFontStyle("italic");
  sheet.getRange(2, 1, 1, maxCol).merge();
  
  // 4. 定義 Header
  var headers = ["No.", "班級", "學號", "姓名", "分組"];
  datesSet.forEach(function(date) {
    headers.push(date);
  });
  
  var headerRange = sheet.getRange(4, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground("#1e3a8a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 28);
  
  if (students.length === 0) {
    sheet.getRange(5, 1).setValue("當前課程尚無學生成冊資料。").setFontColor("#9ca3af").setFontStyle("italic");
    return;
  }
  
  // 5. 整理與寫入點名資料
  var sortedStudents = students.slice().sort(function(a, b) {
    return (a.no || 0) - (b.no || 0);
  });
  
  var rowsData = [];
  sortedStudents.forEach(function(s, idx) {
    var row = [
      s.no || (idx + 1),
      s.class || "",
      s.id || "",
      s.name || "",
      s.group || "未分組"
    ];
    
    datesSet.forEach(function(date) {
      var status = s.attendance ? s.attendance[date] : null;
      row.push(status || "未點名");
    });
    
    rowsData.push(row);
  });
  
  var dataRange = sheet.getRange(5, 1, rowsData.length, headers.length);
  dataRange.setValues(rowsData)
    .setFontSize(9)
    .setVerticalAlignment("middle");
  
  sheet.getRange(5, 1, rowsData.length, 5).setHorizontalAlignment("left");
  if (datesSet.length > 0) {
    sheet.getRange(5, 6, rowsData.length, datesSet.length).setHorizontalAlignment("center");
  }
  
  // 劃上邊框
  dataRange.setBorder(true, true, true, true, true, true, "#e5e7eb", SpreadsheetApp.BorderStyle.SOLID);
  
  // 6. 為點名狀態設定背景顏色條件格式化
  if (datesSet.length > 0) {
    var statusRange = sheet.getRange(5, 6, rowsData.length, datesSet.length);
    var rules = sheet.getConditionalFormatRules();
    
    // 出席: 綠色
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("出席")
      .setBackground("#dcfce7")
      .setFontColor("#15803d")
      .setRanges([statusRange])
      .build());
      
    // 遲到: 黃色
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("遲到")
      .setBackground("#fef3c7")
      .setFontColor("#b45309")
      .setRanges([statusRange])
      .build());
      
    // 缺席: 紅色
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("缺席")
      .setBackground("#fee2e2")
      .setFontColor("#b91c1c")
      .setRanges([statusRange])
      .build());
      
    sheet.setConditionalFormatRules(rules);
  }
  
  // 自動調整列寬
  for (var col = 1; col <= headers.length; col++) {
    sheet.autoResizeColumn(col);
    var currentWidth = sheet.getColumnWidth(col);
    sheet.setColumnWidth(col, currentWidth + 12);
  }
}
