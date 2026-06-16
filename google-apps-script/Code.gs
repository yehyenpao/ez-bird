/**
 * 羽毛球比賽系統 - Google Apps Script 後端 (終極穩定版)
 */

const CONFIG = {
  SHEET_REGISTRATION: "報名紀錄",
  SHEET_MEMBER: "隊員紀錄",
  SHEET_ROUND_ROBIN: "預賽紀錄表",
  SHEET_CHASING: "追分賽紀錄表",
  SHEET_ELIMINATION: "單淘汰追分賽",
  SHEET_POINTS: "積點統計表",
  SHEET_SPECIAL: "特殊紀錄",
  SHEET_PLAYER_DB: "球員資料庫",
  TEAMS: ["藍鳥隊", "黑鳥隊", "青鳥隊", "粉鳥隊"],
  AREAS: ["猛禽區", "小鳥區", "鳥蛋區"],
  TIMEZONE: "GMT+8"
};

/**
 * 統一入口：所有請求 (GET & POST) 都轉向這裡
 */
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const action = e.parameter.action;
    const yearMonth = e.parameter.yearMonth || "";
    let data = null;
    
    if (e.parameter.data) {
      data = JSON.parse(decodeURIComponent(e.parameter.data));
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents).data;
    }
    
    if (!action) {
      return createResponse({ status: "error", message: "缺少 action 指令" });
    }

    let result = { status: "error", message: "未定義的指令" };
    
    switch (action) {
      case "ping":
        result = { status: "success", message: "pong" };
        break;
      case "getRegistrations":
        result = { status: "success", data: helperGetData(CONFIG.SHEET_REGISTRATION, yearMonth) };
        break;
      case "getSchedule":
        result = { status: "success", data: helperGetData(CONFIG.SHEET_ROUND_ROBIN, yearMonth) };
        break;
      case "getLiveScores":
        const rrLive = helperGetData(CONFIG.SHEET_ROUND_ROBIN, yearMonth);
        const chLive = helperGetData(CONFIG.SHEET_CHASING, yearMonth);
        result = { status: "success", data: rrLive.concat(chLive) };
        break;
      case "getChasingSchedule":
        result = { status: "success", data: helperGetData(CONFIG.SHEET_CHASING, yearMonth) };
        break;
      case "getRankings":
        result = { status: "success", data: helperGetTeamRankings(yearMonth) };
        break;
      case "addRegistrations":
        result = logicAddRegistrations(data);
        break;
      case "autoGroup":
        result = logicAutoGroup(yearMonth, typeof data === "string" ? data : (data && data.mode) || "default");
        break;
      case "generateSchedule":
        result = logicGenerateSchedule(yearMonth);
        break;
      case "updateScore":
        result = logicUpdateScore(data);
        break;
      case "updateChasingScore":
        result = logicUpdateChasingScore(data);
        break;
      case "updatePlayerOrder":
        result = logicUpdatePlayerOrder(data);
        break;
      case "generateChasingSchedule":
        result = logicGenerateChasingSchedule(yearMonth, data);
        break;
      case "generateFinals":
        result = logicGenerateFinals(yearMonth, data);
        break;
      case "clearData":
        result = logicClearData(yearMonth, e.parameter.sheet);
        break;
      case "calculatePoints":
        result = logicCalculatePoints(yearMonth, data);
        break;
      case "getPointsRecords":
        result = { status: "success", data: helperGetPointsRecords(yearMonth) };
        break;
      case "getSpecialRecords":
        result = { status: "success", data: helperGetSpecialRecords() };
        break;
      case "saveSpecialRecords":
        result = logicSaveSpecialRecords(yearMonth, data);
        break;
      case "getPlayersInfo":
        result = { status: "success", data: logicGetPlayersInfo() };
        break;
      case "uploadPhoto":
        result = logicUploadPhoto(data);
        break;
      case "generateLotteryKnockout":
        result = logicGenerateLotteryKnockout(yearMonth, data);
        break;
      case "getLatestDate":
        result = { status: "success", data: logicGetLatestDate() };
        break;
      case "debugSheet":
        result = { status: "success", data: helperDebugSheet(e.parameter.sheet || CONFIG.SHEET_CHASING, e.parameter.query || "") };
        break;
      case "debugCalc":
        result = { status: "success", data: helperDebugCalc(e.parameter.yearMonth || "2026-06-14") };
        break;
      case "updateZones":
        result = logicUpdateRegistrationZones(yearMonth || "2026-06-14");
        break;
      case "updateRanks":
        result = logicUpdateRegistrationRanks(yearMonth || "2026-06-14");
        break;
    }
    
    return createResponse(result);
  } catch (err) {
    return createResponse({ status: "error", message: err.toString() });
  }
}

function createResponse(content) {
  return ContentService.createTextOutput(JSON.stringify(content))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 取得指定月份的資料
 */
function helperGetData(sheetName, yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const ymIdx = headers.indexOf("年月");
  const timeIdx = headers.indexOf("比賽時間");
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    // 處理「年月」欄位 - 統一轉為 "yyyy-MM" 字串再比對
    let rYM = ymIdx > -1 ? data[i][ymIdx] : "";
    if (rYM instanceof Date) {
      rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    } else {
      // 非 Date 時，取前 10 碼 ("2026-03-01") 防止格式不一致
      rYM = String(rYM).trim().substring(0, 10);
    }
    
    const targetDate = String(yearMonth).trim();
    if (!targetDate) {
      // 不篩選日期
    } else {
      let isMatch = false;
      if (rYM === targetDate) {
        isMatch = true;
      } else if (rYM.length === 7 && targetDate.startsWith(rYM)) {
        // 舊資料支援匹配整個月
        isMatch = true;
      } else if (targetDate.length === 7 && rYM.startsWith(targetDate)) {
        isMatch = true;
      }
      
      if (!isMatch) continue;
    }

    const obj = {};
    headers.forEach((h, idx) => {
      let val = data[i][idx];
      
      // 處理「比賽時間」欄位：轉為 HH:mm 格式
      if (idx === timeIdx && val instanceof Date) {
        val = Utilities.formatDate(val, CONFIG.TIMEZONE, "HH:mm");
      }
      // 處理「年月」欄位：統一輸出為 yyyy-MM-dd 字串
      if (idx === ymIdx && val instanceof Date) {
        val = Utilities.formatDate(val, CONFIG.TIMEZONE, "yyyy-MM-dd");
      }
      
      obj[h] = val;
    });
    result.push(obj);
  }
  return result;
}

/**
 * 診斷用：直接回傳試算表所有原始資料（不篩選年月）
 */
function helperDebugSheet(sheetName, query) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (sheetName === "list") {
    const sheets = ss.getSheets();
    return {
      sheets: sheets.map(s => s.getName())
    };
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: "找不到工作表: " + sheetName };
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { rowCount: data.length, headers: data[0] || [], rows: [] };
  
  const headers = data[0];
  
  // 找出所有包含指定 query 的行，回傳其整行資料，方便確認欄位對齊狀況
  const targetQuery = query || "2026-06-14";
  const matchingRows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowStr = row.map(v => v instanceof Date ? Utilities.formatDate(v, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss") : String(v)).join(" | ");
    if (rowStr.includes(targetQuery) || rowStr.includes("2026-05-31")) {
      matchingRows.push({
        rowNum: i + 1,
        content: row.map(v => v instanceof Date ? Utilities.formatDate(v, CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm") : String(v))
      });
    }
  }
  
  return {
    sheetName: sheetName,
    totalRows: data.length - 1,
    headers: headers,
    matchingRowsCount: matchingRows.length,
    matchingRows: matchingRows.slice(0, 100)
  };
}

function helperDebugCalc(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];
  
  log.push("Starting debugCalc for " + yearMonth);
  
  // 1. prevBalances
  const prevBalances = {};
  let pSheet = ss.getSheetByName(CONFIG.SHEET_POINTS);
  if (pSheet) {
    const pData = pSheet.getDataRange().getValues();
    log.push("pSheet total rows: " + pData.length);
    
    const targetDate = new Date(yearMonth);
    let latestPrevDate = null;
    let latestPrevDateStr = "";
    
    for (let i = 1; i < pData.length; i++) {
      let rowDate = pData[i][0]; // 年月
      if (!(rowDate instanceof Date)) rowDate = new Date(rowDate);
      if (rowDate < targetDate) {
        if (!latestPrevDate || rowDate > latestPrevDate) {
          latestPrevDate = rowDate;
          latestPrevDateStr = Utilities.formatDate(rowDate, CONFIG.TIMEZONE, "yyyy-MM-dd");
        }
      }
    }
    log.push("latestPrevDateStr: " + latestPrevDateStr);
    if (latestPrevDateStr) {
      for (let i = 1; i < pData.length; i++) {
        let rowDate = pData[i][0];
        const rowDateStr = rowDate instanceof Date ? 
          Utilities.formatDate(rowDate, CONFIG.TIMEZONE, "yyyy-MM-dd") : String(rowDate).substring(0, 10);
        if (rowDateStr === latestPrevDateStr) {
          prevBalances[pData[i][2]] = parseInt(pData[i][12]) || 0; // 姓名, 累積積點
        }
      }
    }
  }
  log.push("prevBalances['子安']: " + prevBalances["子安"]);
  
  // 2. currReg
  const currReg = helperGetData(CONFIG.SHEET_REGISTRATION, yearMonth);
  log.push("currReg count: " + currReg.length);
  const testPlayer = "子安";
  const pReg = currReg.find(p => p["姓名"] === testPlayer);
  log.push("pReg for " + testPlayer + ": " + JSON.stringify(pReg));
  
  // 3. Initialize playersMap
  const playersMap = {};
  Object.keys(prevBalances).forEach(name => {
    playersMap[name] = { 
      name: name, team: "", area: "", 
      rrRank: "-", elimRank: "-", 
      currPts: prevBalances[name], 
      guessPts: 0, refPts: 0, rrPts: 0, elimPts: 0, totalPts: 0 
    };
  });
  
  currReg.forEach(p => {
    const name = p["姓名"];
    if (!playersMap[name]) {
      playersMap[name] = { 
        name: name, team: p["隊名"] || "", area: p["區"] || "", 
        rrRank: p["循環名次"] || "-", elimRank: p["淘汰名次"] || "-", 
        currPts: prevBalances[name] || 0, 
        guessPts: 0, refPts: 0, rrPts: 0, elimPts: 0, totalPts: 0 
      };
    } else {
      playersMap[name].team = p["隊名"] || playersMap[name].team;
      playersMap[name].area = p["區"] || playersMap[name].area;
      if (p["循環名次"]) playersMap[name].rrRank = p["循環名次"];
      if (p["淘汰名次"]) playersMap[name].elimRank = p["淘汰名次"];
    }
  });
  
  log.push("playersMap['" + testPlayer + "'] after currReg: " + JSON.stringify(playersMap[testPlayer]));
  
  // 4. rrMatches
  const rrMatches = helperGetData(CONFIG.SHEET_ROUND_ROBIN, yearMonth);
  log.push("rrMatches count: " + rrMatches.length);
  
  rrMatches.forEach((m, idx) => {
    const isA = (m["A隊員1"] === testPlayer || m["A隊員2"] === testPlayer);
    const isB = (m["B隊員1"] === testPlayer || m["B隊員2"] === testPlayer);
    if (isA || isB) {
      log.push("Match " + idx + " involves " + testPlayer + ": " + JSON.stringify(m));
      const sA = parseInt(m["A隊比分"]) || 0;
      const sB = parseInt(m["B隊比分"]) || 0;
      const area = String(m["區"]);
      const isTeamMatch = area.includes("團體");
      log.push("  isTeamMatch: " + isTeamMatch + ", area: " + area);
      
      let ptsCfg = { win: 0, lose: 0 };
      if (isTeamMatch) {
        const pArea = String(playersMap[testPlayer].area);
        log.push("  pArea: " + pArea);
        if (pArea.includes("猛禽")) {
          ptsCfg = { win: 200, lose: 100 };
        } else if (pArea.includes("鳥蛋")) {
          ptsCfg = { win: 120, lose: 60 };
        } else {
          ptsCfg = { win: 160, lose: 80 };
        }
      }
      log.push("  ptsCfg: " + JSON.stringify(ptsCfg));
      
      const officialTeams = ["藍鳥隊", "黑鳥隊", "青鳥隊", "粉鳥隊"];
      const isOfficialOrTeamMatch = area.includes("男雙") || 
        area.includes("女雙") || 
        officialTeams.includes(playersMap[testPlayer].team) ||
        isTeamMatch;
        
      log.push("  isOfficialOrTeamMatch: " + isOfficialOrTeamMatch);
      if (isOfficialOrTeamMatch) {
        const win = isA ? (sA > sB) : (sB > sA);
        const pPts = win ? ptsCfg.win : ptsCfg.lose;
        log.push("  win: " + win + ", adding points: " + pPts);
        playersMap[testPlayer].rrPts += pPts;
      }
    }
  });
  
  log.push("playersMap['" + testPlayer + "'] rrPts final: " + playersMap[testPlayer].rrPts);
  
  // 5. elimination pts trace
  const rank = String(playersMap[testPlayer].elimRank || "").trim();
  log.push("testPlayer elimRank: " + rank);
  let elimPts = 0;
  if (rank === "1" || rank === "冠軍") {
    elimPts = 400;
  } else if (rank === "2" || rank === "亞軍") {
    elimPts = 350;
  } else if (rank === "3" || rank === "季軍") {
    elimPts = 300;
  } else if (rank === "4" || rank === "殿軍") {
    elimPts = 250;
  }
  log.push("testPlayer elimPts calculated: " + elimPts);
  
  return log;
}

/**
 * 寫入報名資料
 */
function logicAddRegistrations(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_REGISTRATION);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_REGISTRATION);
    sheet.appendRow(["年月", "姓名", "身份", "隊名", "區", "循環名次", "淘汰名次", "循環積分", "淘汰積分"]);
  }

  // 解析前端傳入的結構
  let dataList = Array.isArray(data) ? data : (data.items || []);
  const overwrite = (data && data.overwrite === true);
  const append   = (data && data.append   === true);

  // ── 預檢：若無明確指示，先確認是否已有同月資料 ──
  // 使用前 7 碼 (yyyy-MM) 比對，避免跨日期同月資料被漏判
  if (!overwrite && !append && dataList.length > 0) {
    const targetYM = String(dataList[0].yearMonth).substring(0, 7);
    const existing = sheet.getDataRange().getValues();
    let existCount = 0;
    for (let i = 1; i < existing.length; i++) {
      let rYM = existing[i][0];
      if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM");
      if (String(rYM).substring(0, 7) === targetYM) existCount++;
    }
    if (existCount > 0) {
      return {
        status: "warning",
        code: "ALREADY_EXISTS",
        count: existCount,
        message: `資料庫中「${targetYM}」已有 ${existCount} 筆報名資料，請確認是否覆蓋或追加。`
      };
    }
  }

  // ── 追加模式：檢查追加後是否超過 24 人上限 ──
  if (append && dataList.length > 0) {
    const targetYM = String(dataList[0].yearMonth).substring(0, 7);
    const existing = sheet.getDataRange().getValues();
    let existCount = 0;
    for (let i = 1; i < existing.length; i++) {
      let rYM = existing[i][0];
      if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM");
      if (String(rYM).substring(0, 7) === targetYM) existCount++;
    }
    const totalAfter = existCount + dataList.length;
    if (totalAfter > 24) {
      return {
        status: "error",
        message: `追加失敗：目前已有 ${existCount} 人，再追加 ${dataList.length} 人將超過 24 人上限（共 ${totalAfter} 人）。`
      };
    }
  }

  // ── 覆蓋模式：先刪除同月舊資料 ──
  if (overwrite && dataList.length > 0) {
    const targetYM = String(dataList[0].yearMonth).substring(0, 7);
    const rows = sheet.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      let rYM = rows[i][0];
      if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM");
      if (String(rYM).substring(0, 7) === targetYM) sheet.deleteRow(i + 1);
    }
  }

  // ── 寫入資料 ──
  dataList.forEach(d => {
    sheet.appendRow([
      d.yearMonth,
      d.name,
      d.role || "球員",
      d.team || "",
      d.area || "",
      "", "", "", ""
    ]);
  });

  // 同步球員資料庫
  const names = dataList.map(d => d.name);
  if (names.length > 0) syncPlayerDatabase(names);

  return { status: "success", message: `成功匯入 ${dataList.length} 筆報名資料！` };
}

/**
 * 自動產生循環賽程
 */
function logicGenerateSchedule(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const regItems = helperGetData(CONFIG.SHEET_REGISTRATION, yearMonth);
  if (regItems.length < 2) return { status: "error", message: "分組人數不足 (目前 " + regItems.length + " 人)" };
  
  let sheet = ss.getSheetByName(CONFIG.SHEET_ROUND_ROBIN);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_ROUND_ROBIN);
  } else {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      let rYM = data[i][1];
      if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (String(rYM || "").trim().substring(0, 10) === String(yearMonth).trim().substring(0, 10)) sheet.deleteRow(i + 1);
    }
  }
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["序號", "年月", "比賽時間", "輪次", "區", "場地", "A隊名", "A隊員1", "A隊員2", "A隊比分", "B隊比分", "B隊名", "B隊員1", "B隊員2", "裁判", "比賽狀態"]);
  }

  // 判斷當前模式：查看註冊表中的區域特徵與隊伍數量
  const areasFound = [...new Set(regItems.map(p => p.區).filter(a => a !== ""))];
  const teamsFound = [...new Set(regItems.map(p => p.隊名).filter(t => t !== ""))];
  // 鳥樂賽的區域名必含「場」(如「粉紅爆(C場)」)；標準模式的「猛禽區/小鳥區/孵蛋 A」等均不含「場」
  // 故僅以「場」判斷，避免「小鳥區」被誤判為鳥樂賽
  const isLottery = teamsFound.length > 4 || areasFound.some(a => a.includes("場"));
  const hasCourtInfo = areasFound.some(a => a.includes("場"));

  if (isLottery && !hasCourtInfo) {
    return { status: "error", message: "尚未完成分區！請先返回「報名管理」執行【啟動智慧分組】（或手動填入 A區/B區/C區），才能產生鳥樂賽的賽程喔！" };
  }

  if (isLottery) {
    // 【鳥樂賽】12隊預賽產生邏輯
    // 動態讀取包含「場」的區域（鳥樂賽區域名均含「場」）
    const courtAreas = areasFound.filter(a => a.includes("場")).sort();
    if (courtAreas.length === 0) {
      return { status: "error", message: "找不到任何包含「場」的分區，無法產生賽程！" };
    }

    const matchups = [[0, 2], [1, 3], [0, 1], [2, 3], [0, 3], [2, 1]];
    const roundTimes = ["13:40", "14:05", "14:30", "14:45", "15:00", "15:15"];
    let sequenceNum = 1;

    // 建立一個以「區 + 隊名」為 key 的人員查詢 Map (排除空白/大小寫干擾)
    const playersByAreaTeam = {};
    regItems.forEach(p => {
      const aKey = String(p.區 || "").trim();
      const tKey = String(p.隊名 || "").trim();
      if (!aKey || !tKey) return;
      const key = `${aKey}||${tKey}`;
      if (!playersByAreaTeam[key]) playersByAreaTeam[key] = [];
      playersByAreaTeam[key].push(String(p.姓名 || "").trim());
    });

    Logger.log("playersByAreaTeam keys: " + JSON.stringify(Object.keys(playersByAreaTeam)));

    courtAreas.forEach((area) => {
      const court = area.includes("C") ? "C" : area.includes("B") ? "B" : "A";
      const teamsInArea = [...new Set(
        regItems
          .filter(p => String(p.區 || "").trim() === String(area).trim())
          .map(p => String(p.隊名 || "").trim())
          .filter(t => t)
      )];
      
      Logger.log(`Area: ${area}, Teams: ${JSON.stringify(teamsInArea)}`);

      matchups.forEach((pair, idx) => {
        const teamA = teamsInArea[pair[0]];
        const teamB = teamsInArea[pair[1]];
        if (!teamA || !teamB) return;

        const keyA = `${String(area).trim()}||${teamA}`;
        const keyB = `${String(area).trim()}||${teamB}`;
        const pA = playersByAreaTeam[keyA] || [];
        const pB = playersByAreaTeam[keyB] || [];

        Logger.log(`Match ${sequenceNum}: ${teamA}(${pA.join(",")}) vs ${teamB}(${pB.join(",")})`);

        sheet.appendRow([
          sequenceNum++, yearMonth, roundTimes[idx], (idx + 1).toString(), area, court, 
          teamA, pA[0] || "待定", pA[1] || "待定", 0, 
          0, teamB, pB[0] || "待定", pB[1] || "待定", "", "待賽"
        ]);
      });
    });
    return { status: "success", message: "【鳥樂賽】12隊預賽賽程已成功產生！" };
  }

  // 【常規賽】產生邏輯
  const matchups = [[0, 2], [1, 3], [0, 1], [2, 3], [0, 3], [1, 2]];
  const roundTimes = ["13:40", "14:05", "14:30", "14:45", "15:00", "15:15"];
  
  // 動態對應場地 (根據關鍵字比對)
  const getCourt = (area) => {
    if (area.includes("猛") || area.includes("孵蛋 A") || area.includes("狐狸 A")) return "C";
    if (area.includes("小鳥") || area.includes("狐狸 B") || area.includes("醬板鴨 A")) return "B";
    if (area.includes("蛋") || area.includes("孵蛋 B") || area.includes("醬板鴨 B")) return "A";
    return "A";
  };
  
  let sequenceNum = 1;
  matchups.forEach((pair, idx) => {
    areasFound.forEach((area) => {
      const teamA = CONFIG.TEAMS[pair[0]];
      const teamB = CONFIG.TEAMS[pair[1]];
      const court = getCourt(area);
      const playersA = regItems.filter(p => p.隊名 === teamA && p.區 === area);
      const playersB = regItems.filter(p => p.隊名 === teamB && p.區 === area);
      
      sheet.appendRow([
        sequenceNum++, yearMonth, roundTimes[idx], (idx + 1).toString(), area, court, 
        teamA, playersA[0]?.姓名 || "待定", playersA[1]?.姓名 || "待定", 0, 
        0, teamB, playersB[0]?.姓名 || "待定", playersB[1]?.姓名 || "待定", "", "待賽"
      ]);
    });
  });
  
  return { status: "success", message: "常規賽賽程已成功產生，並已動態適配分組名稱與場地！" };
}

/**
 * 更新裁判比分 (即時同步)
 */
function logicUpdateScore(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ROUND_ROBIN);
  const rows = sheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    let rYM = rows[i][1];
    if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    else rYM = String(rYM).trim().substring(0, 10);
    
    let isYMMatch = (rYM === String(d.yearMonth));
    if (!isYMMatch && rYM.length === 7) isYMMatch = d.yearMonth.startsWith(rYM);

    // 輪次 (Index 3), 場地 (Index 5)
    if (isYMMatch && rows[i][3] == d.round && rows[i][5] == d.court) {
      if (d.startTime) sheet.getRange(i + 1, 3).setValue(d.startTime); // 比賽時間移動到第 3 欄 (Index 2)
      sheet.getRange(i + 1, 10).setValue(d.scoreA);
      sheet.getRange(i + 1, 11).setValue(d.scoreB);
      if (d.referee !== undefined) sheet.getRange(i + 1, 15).setValue(d.referee);
      
      const newStatus = d.status || "已完賽";
      sheet.getRange(i + 1, 16).setValue(newStatus);
      
      return { status: "success", message: "比分已成功更新為 " + newStatus };
    }
  }
  return { status: "error", message: "找不到該場比賽" };
}

function helperCalculateRankings(yearMonth) {
  const items = helperGetData(CONFIG.SHEET_ROUND_ROBIN, yearMonth);
  const points = {};
  items.forEach(m => {
    const sA = parseInt(m.A得分 || 0);
    const sB = parseInt(m.B得分 || 0);
    if (sA === 0 && sB === 0) return;
    
    const pA = sA > sB ? 3 : (sA < sB ? 1 : 1);
    const pB = sA > sB ? 1 : (sA < sB ? 3 : 1);
    
    (m.A球員 || "").split("/").forEach(n => points[n] = (points[n] || 0) + pA);
    (m.B球員 || "").split("/").forEach(n => points[n] = (points[n] || 0) + pB);
  });
  return points;
}

function logicClearData(yearMonth, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: "error", message: "找不到工作表" };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    let rYM = data[i][0];
    if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, Session.getScriptTimeZone(), "yyyy-MM");
    if (String(rYM).trim() === String(yearMonth).trim()) sheet.deleteRow(i + 1);
  }
  return { status: "success", message: "已清除 " + yearMonth + " 的相關資料" };
}

/**
 * 取得團隊排名計算 (積分 > 正負商)
 */
/**
 * 取得團隊排名計算 (積分 > 對戰勝場 > 正負商)
 */
function helperGetTeamRankings(yearMonth) {
  const items = helperGetData(CONFIG.SHEET_ROUND_ROBIN, yearMonth);
  if (items.length === 0) return [];

  // 1. 動態偵測場地/分組 (判斷是否為鳥樂賽)
  const areas = [...new Set(items.map(m => m["區"]))];
  const isLottery = areas.some(a => a.includes("場"));
  
  const stats = {};
  const clean = (str) => String(str || "").replace(/\s+/g, "").trim();
  
  // 找出所有參賽隊伍
  const allTeamsFound = [...new Set(items.flatMap(m => [clean(m.A隊名), clean(m.B隊名)]))].filter(t => t !== "");

  allTeamsFound.forEach(t => {
    stats[t] = { 
        name: t, 
        area: items.find(m => clean(m.A隊名) === t || clean(m.B隊名) === t)["區"],
        points: 0, scored: 0, conceded: 0, wins: 0, matches: 0 
    };
  });

  // 2. 累加數據
  items.forEach(m => {
    const tA = clean(m.A隊名); const tB = clean(m.B隊名);
    const sA = Number(m.A隊比分) || 0; const sB = Number(m.B隊比分) || 0;
    if (sA === 0 && sB === 0) return;

    if (stats[tA] && stats[tB]) {
      stats[tA].matches++; stats[tB].matches++;
      stats[tA].scored += sA; stats[tA].conceded += sB;
      stats[tB].scored += sB; stats[tB].conceded += sA;
      if (sA > sB) { stats[tA].points += 3; stats[tA].wins++; }
      else if (sB > sA) { stats[tB].points += 3; stats[tB].wins++; }
      else { stats[tA].points += 1; stats[tB].points += 1; }
    }
  });

  // 3. 排序與分組排名
  const results = Object.values(stats);
  
  // 如果是鳥樂賽，我們回傳一個按區域分組的結果，或者排序好的 12 隊列表
  return results.sort((a, b) => {
    if (b.points !== a.points) return b.points - a.points;
    if (b.wins !== a.wins) return b.wins - a.wins;
    const diffA = a.scored - a.conceded;
    const diffB = b.scored - b.conceded;
    return diffB - diffA;
  });
}

function helperBuildRoundRobinRankLookup(yearMonth) {
  const matches = helperGetData(CONFIG.SHEET_ROUND_ROBIN, yearMonth);
  if (!matches || matches.length === 0) return {};

  const clean = (value) => String(value || "").replace(/\s+/g, "").trim();
  const stats = {};
  const allTeams = [...new Set(matches.flatMap(m => [clean(m["A隊名"]), clean(m["B隊名"])]).filter(Boolean))];
  const isLottery = allTeams.length > 4;

  matches.forEach(match => {
    const teamA = clean(match["A隊名"]);
    const teamB = clean(match["B隊名"]);
    const area = String(match["區"] || "").trim();

    [teamA, teamB].forEach(team => {
      if (!team) return;
      if (!stats[team]) {
        stats[team] = {
          teamName: team,
          area: area,
          matchWins: 0,
          matchLosses: 0,
          matchPoints: 0,
          totalPoints: 0,
          totalScored: 0,
          totalConceded: 0,
          diff: 0,
          quotient: 0,
          rank: 0
        };
      }
    });
  });

  matches.forEach(match => {
    const teamA = clean(match["A隊名"]);
    const teamB = clean(match["B隊名"]);
    if (!teamA || !teamB || !stats[teamA] || !stats[teamB]) return;

    const scoreA = parseInt(match["A隊比分"] || 0, 10);
    const scoreB = parseInt(match["B隊比分"] || 0, 10);

    stats[teamA].totalScored += scoreA;
    stats[teamA].totalConceded += scoreB;
    stats[teamB].totalScored += scoreB;
    stats[teamB].totalConceded += scoreA;

    if (scoreA > scoreB) {
      stats[teamA].matchWins++;
      stats[teamB].matchLosses++;
      stats[teamA].totalPoints += isLottery ? 100 : 3;
      stats[teamB].totalPoints += isLottery ? 50 : 1;
    } else if (scoreB > scoreA) {
      stats[teamB].matchWins++;
      stats[teamA].matchLosses++;
      stats[teamB].totalPoints += isLottery ? 100 : 3;
      stats[teamA].totalPoints += isLottery ? 50 : 1;
    }
  });

  const comboGroups = {};
  matches.forEach(match => {
    const teamA = clean(match["A隊名"]);
    const teamB = clean(match["B隊名"]);
    if (!teamA || !teamB || !stats[teamA] || !stats[teamB]) return;

    const round = clean(match["輪次"]);
    const pairKey = [teamA, teamB].sort().join("::");
    const groupKey = round + "||" + pairKey;

    if (!comboGroups[groupKey]) {
      comboGroups[groupKey] = {
        teamA: teamA,
        teamB: teamB,
        wins: {}
      };
      comboGroups[groupKey].wins[teamA] = 0;
      comboGroups[groupKey].wins[teamB] = 0;
    }

    const scoreA = parseInt(match["A隊比分"] || 0, 10);
    const scoreB = parseInt(match["B隊比分"] || 0, 10);
    if (scoreA > scoreB) comboGroups[groupKey].wins[teamA]++;
    if (scoreB > scoreA) comboGroups[groupKey].wins[teamB]++;
  });

  Object.keys(comboGroups).forEach(key => {
    const group = comboGroups[key];
    const winsA = group.wins[group.teamA] || 0;
    const winsB = group.wins[group.teamB] || 0;

    if (winsA === 0 && winsB === 0) return;

    if (winsA > winsB) {
      stats[group.teamA].matchPoints += 2;
      stats[group.teamB].matchPoints += 1;
    } else if (winsB > winsA) {
      stats[group.teamB].matchPoints += 2;
      stats[group.teamA].matchPoints += 1;
    }
  });

  Object.keys(stats).forEach(team => {
    const item = stats[team];
    item.diff = item.totalScored - item.totalConceded;
    item.quotient = item.totalConceded === 0
      ? (item.totalScored > 0 ? 999 : 0)
      : (item.totalScored / item.totalConceded);
  });

  const sortFn = function(a, b) {
    if (b.matchPoints !== a.matchPoints) return b.matchPoints - a.matchPoints;
    if (b.matchWins !== a.matchWins) return b.matchWins - a.matchWins;
    if (b.diff !== a.diff) return b.diff - a.diff;
    return b.quotient - a.quotient;
  };

  const rankLookup = {};

  if (isLottery) {
    const byArea = {};
    Object.keys(stats).forEach(team => {
      const areaKey = stats[team].area || "未分區";
      if (!byArea[areaKey]) byArea[areaKey] = [];
      byArea[areaKey].push(stats[team]);
    });

    Object.keys(byArea).forEach(area => {
      byArea[area].sort(sortFn).forEach((team, index) => {
        team.rank = index + 1;
        rankLookup[team.teamName] = String(team.rank);
      });
    });
  } else {
    Object.values(stats).sort(sortFn).forEach((team, index) => {
      team.rank = index + 1;
      rankLookup[team.teamName] = String(team.rank);
    });
  }

  return rankLookup;
}

/**
 * 更新球員棒次 (1~6 棒)
 */
function logicUpdatePlayerOrder(dataList) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_REGISTRATION);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  
  // 動態尋找欄位索引
  let nameIdx = headers.indexOf("姓名");
  let orderIdx = headers.indexOf("棒次");
  
  // 如果沒有棒次欄位，則新增一個
  if (orderIdx === -1) {
    orderIdx = headers.length;
    sheet.getRange(1, orderIdx + 1).setValue("棒次");
  }
  
  dataList.forEach(d => {
    for (let i = 1; i < rows.length; i++) {
        // 比對姓名，確保去除空白
        if (String(rows[i][nameIdx]).trim() === String(d.name).trim()) {
            sheet.getRange(i + 1, orderIdx + 1).setValue(d.order);
        }
    }
  });
  return { status: "success", message: "球員棒次已成功儲存至試算表 (第 " + (orderIdx+1) + " 欄)" };
}

/**
 * 確保追分賽記錄表欄位正確。判斷基準：是否包含「年月」欄位（而非對齊結構）
 * 這樣即便表格第一行是標題文字也能正確處理。
 */
function ensureChasingHeaders(sheet) {
  const STANDARD_HEADERS = ["序號", "年月", "比賽時間", "輪次", "區", "場地", "A隊名", "A隊員1", "A隊員2", "A隊員3", "A隊比分", "B隊比分", "B隊名", "B隊員1", "B隊員2", "B隊員3", "裁判", "比賽狀態"];
  
  let rawHeaders = sheet.getDataRange().getValues()[0] || [];
  let headers = rawHeaders.map(h => String(h || "").trim());

  if (!headers.includes("年月") || !headers.includes("序號")) {
    sheet.clearContents(); 
    sheet.getRange(1, 1, 1, STANDARD_HEADERS.length).setValues([STANDARD_HEADERS]);
    return STANDARD_HEADERS;
  }
  
  if (!headers.includes("A隊員3")) {
    const idxA2 = headers.indexOf("A隊員2");
    if (idxA2 !== -1) {
      sheet.insertColumnAfter(idxA2 + 1);
      sheet.getRange(1, idxA2 + 2).setValue("A隊員3");
      headers = sheet.getDataRange().getValues()[0].map(h => String(h || "").trim());
    }
  }

  if (!headers.includes("B隊員3")) {
    const idxB2 = headers.indexOf("B隊員2");
    if (idxB2 !== -1) {
      sheet.insertColumnAfter(idxB2 + 1);
      sheet.getRange(1, idxB2 + 2).setValue("B隊員3");
      headers = sheet.getDataRange().getValues()[0].map(h => String(h || "").trim());
    }
  }
  
  return headers;
}

/**
 * 產生追分賽程 (準決賽)，接收前端自訂出戰順序表
 */
function logicGenerateChasingSchedule(yearMonth, customizedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!customizedData || customizedData.length === 0) return { status: "error", message: "沒有接收到自訂賽程資料" };

  let sheet = ss.getSheetByName(CONFIG.SHEET_CHASING);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_CHASING);
  const headers = ensureChasingHeaders(sheet);
  
  // 清除舊資料
  const oldData = sheet.getDataRange().getValues();
  for (let i = oldData.length - 1; i >= 1; i--) {
     let rYM = oldData[i][0];
     if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
     if (String(rYM).substring(0, 10) === String(yearMonth).substring(0, 10) && String(oldData[i][headers.indexOf("區")]).includes("準決賽")) {
         sheet.deleteRow(i + 1);
     }
  }

  const headIdx = {};
  headers.forEach((h, i) => headIdx[String(h).trim()] = i);
  if (headIdx["年月"] === undefined) return { status: "error", message: "標題異常：找不到「年月」" };

  let seqCount = 1;
  customizedData.forEach(p => {
    const rowObj = headers.map(() => "");
    rowObj[headIdx["序號"]] = seqCount++;
    rowObj[headIdx["年月"]] = yearMonth;
    rowObj[headIdx["比賽時間"]] = "15:30";
    rowObj[headIdx["輪次"]] = p.targetScore; 
    rowObj[headIdx["區"]] = p.area;
    rowObj[headIdx["場地"]] = p.court;
    rowObj[headIdx["A隊名"]] = p.teamA;
    rowObj[headIdx["A隊員1"]] = p.A1;
    rowObj[headIdx["A隊員2"]] = p.A2;
    rowObj[headIdx["A隊員3"]] = p.A3 || "";
    rowObj[headIdx["A隊比分"]] = 0;
    rowObj[headIdx["B隊名"]] = p.teamB;
    rowObj[headIdx["B隊員1"]] = p.B1;
    rowObj[headIdx["B隊員2"]] = p.B2;
    rowObj[headIdx["B隊員3"]] = p.B3 || "";
    rowObj[headIdx["B隊比分"]] = 0;
    rowObj[headIdx["比賽狀態"]] = "待賽";
    sheet.appendRow(rowObj);
  });
  return { status: "success", message: "追分賽程產生成功！" };
}

/**
 * 更新追分賽比分，支援自動承接
 */
function logicUpdateChasingScore(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_CHASING);
  if (!sheet) return { status: "error", message: "找不到表單" };
  
  const rows = sheet.getDataRange().getValues();
  const headIdx = {};
  rows[0].forEach((h, i) => headIdx[h] = i);
  const isRelayRound = /(\u63A5\u529B|\u8FFD\u5206)/.test(String(d.round || ""));
  
  for (let i = 1; i < rows.length; i++) {
    let rYM = rows[i][headIdx["年月"]];
    if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    
    if (String(rYM).substring(0, 10) === String(d.yearMonth).substring(0, 10) && 
        rows[i][headIdx["輪次"]] == d.round && 
        rows[i][headIdx["區"]] == d.area && 
        rows[i][headIdx["場地"]] == d.court) {
      
      sheet.getRange(i + 1, headIdx["A隊比分"] + 1).setValue(d.scoreA);
      sheet.getRange(i + 1, headIdx["B隊比分"] + 1).setValue(d.scoreB);
      if (d.referee) sheet.getRange(i + 1, headIdx["裁判"] + 1).setValue(d.referee);
      
      const newStatus = d.status || "已完賽";
      sheet.getRange(i + 1, headIdx["比賽狀態"] + 1).setValue(newStatus);
      
      if (newStatus === "已完賽") {
        const isAWin = parseInt(d.scoreA) > parseInt(d.scoreB);
        const winTeam = isAWin ? rows[i][headIdx["A隊名"]] : rows[i][headIdx["B隊名"]];
        const winPlayers = isAWin ? [rows[i][headIdx["A隊員1"]], rows[i][headIdx["A隊員2"]], rows[i][headIdx["A隊員3"]]] 
                                  : [rows[i][headIdx["B隊員1"]], rows[i][headIdx["B隊員2"]], rows[i][headIdx["B隊員3"]]];
        const loseTeam = isAWin ? rows[i][headIdx["B隊名"]] : rows[i][headIdx["A隊名"]];
        const losePlayers = isAWin ? [rows[i][headIdx["B隊員1"]], rows[i][headIdx["B隊員2"]], rows[i][headIdx["B隊員3"]]] 
                                   : [rows[i][headIdx["A隊員1"]], rows[i][headIdx["A隊員2"]], rows[i][headIdx["A隊員3"]]];

        const fullArea = String(d.area);
        let matchKey = "";
        const court = fullArea.includes("B區") ? "B" : (fullArea.includes("C區") ? "C" : "A");
        const num = fullArea.match(/\d+/);
        if (fullArea.includes("突圍")) matchKey = court + "突圍" + (num ? num[0] : "");
        else if (fullArea.includes("晉級")) matchKey = court + "晉級" + (num ? num[0] : "");

        if (matchKey) {
          const placeholders = { win: matchKey + "勝者", lose: matchKey + "敗者" };
          for (let k = 1; k < rows.length; k++) {
            let kYM = rows[k][headIdx["年月"]];
            if (kYM instanceof Date) kYM = Utilities.formatDate(kYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
            if (String(kYM).substring(0, 10) !== String(d.yearMonth).substring(0, 10)) continue;

            const check = (pre) => {
              const n = String(rows[k][headIdx[pre + "隊名"]]).trim();
              if (n === placeholders.win) {
                sheet.getRange(k + 1, headIdx[pre + "隊名"] + 1).setValue(winTeam);
                sheet.getRange(k + 1, headIdx[pre + "隊員1"] + 1).setValue(winPlayers[0]||"");
                sheet.getRange(k + 1, headIdx[pre + "隊員2"] + 1).setValue(winPlayers[1]||"");
                sheet.getRange(k + 1, headIdx[pre + "隊員3"] + 1).setValue(winPlayers[2]||"");
              } else if (n === placeholders.lose) {
                sheet.getRange(k + 1, headIdx[pre + "隊名"] + 1).setValue(loseTeam);
                sheet.getRange(k + 1, headIdx[pre + "隊員1"] + 1).setValue(losePlayers[0]||"");
                sheet.getRange(k + 1, headIdx[pre + "隊員2"] + 1).setValue(losePlayers[1]||"");
                sheet.getRange(k + 1, headIdx[pre + "隊員3"] + 1).setValue(losePlayers[2]||"");
              }
            };
            check("A"); check("B");
          }
        } else if (isRelayRound) {
          // 接力承接：找同區同場地下一個待賽
          for (let j = i + 1; j < rows.length; j++) {
            if (String(rows[j][headIdx["區"]]) === String(d.area) && 
                String(rows[j][headIdx["場地"]]) === String(d.court) &&
                String(rows[j][headIdx["比賽狀態"]]).includes("待賽")) {
              sheet.getRange(j + 1, headIdx["A隊比分"] + 1).setValue(d.scoreA);
              sheet.getRange(j + 1, headIdx["B隊比分"] + 1).setValue(d.scoreB);
              break;
            }
          }
        }
      }
      return { status: "success", message: "比分已更新！" };
    }
  }
  return { status: "error", message: "找不到比賽" };
}

/**
 * 產生決賽赛程
 */
function logicGenerateFinals(yearMonth, customizedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_CHASING);
  const headers = ensureChasingHeaders(sheet);
  const dataList = Array.isArray(customizedData) ? customizedData : ((customizedData && customizedData.items) || []);
  const overwrite = !!(customizedData && customizedData.overwrite);
  
  const oldData = sheet.getDataRange().getValues();
  let sameDayCount = 0;
  for (let i = 1; i < oldData.length; i++) {
     let rYM = oldData[i][0];
     if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
     if (String(rYM).substring(0, 10) === String(yearMonth).substring(0, 10)) {
         sameDayCount++;
     }
  }
  if (sameDayCount > 0 && !overwrite) {
    return {
      status: "warning",
      code: "ALREADY_EXISTS",
      count: sameDayCount,
      message: `追分賽紀錄表中已存在 ${sameDayCount} 筆當日資料`
    };
  }

  for (let i = oldData.length - 1; i >= 1; i--) {
     let rYM = oldData[i][0];
     if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
     if (String(rYM).substring(0, 10) === String(yearMonth).substring(0, 10) && (String(oldData[i][headers.indexOf("區")]).includes("冠軍賽") || String(oldData[i][headers.indexOf("區")]).includes("季軍賽"))) {
          sheet.deleteRow(i + 1);
     }
  }

  const headIdx = {};
  headers.forEach((h, i) => headIdx[h] = i);
  let seq = 1;
  const existing = sheet.getDataRange().getValues();
  for(let i=1; i<existing.length; i++) {
    const s = parseInt(existing[i][headIdx["序號"]]) || 0;
    if (s >= seq) seq = s + 1;
  }

  dataList.forEach(p => {
    const r = headers.map(() => "");
    r[headIdx["序號"]] = seq++;
    r[headIdx["年月"]] = yearMonth;
    r[headIdx["比賽時間"]] = "15:30";
    r[headIdx["輪次"]] = p.targetScore;
    r[headIdx["區"]] = p.area;
    r[headIdx["場地"]] = p.court;
    r[headIdx["A隊名"]] = p.teamA;
    r[headIdx["A隊員1"]] = p.A1; r[headIdx["A隊員2"]] = p.A2; r[headIdx["A隊員3"]] = p.A3 || "";
    r[headIdx["A隊比分"]] = 0;
    r[headIdx["B隊名"]] = p.teamB;
    r[headIdx["B隊員1"]] = p.B1; r[headIdx["B隊員2"]] = p.B2; r[headIdx["B隊員3"]] = p.B3 || "";
    r[headIdx["B隊比分"]] = 0;
    r[headIdx["比賽狀態"]] = "待賽";
    sheet.appendRow(r);
  });
  return { status: "success", message: "決賽賽程已產生！" };
}

/**
 * 智慧自動分組 (4隊制 + 12隊制)
 */
function logicAutoGroup(yearMonth, mode = "default") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_REGISTRATION);
  const data = sheet.getDataRange().getValues();
  const TEAMS = CONFIG.TEAMS; // ["藍鳥隊", "黑鳥隊", "青鳥隊", "粉鳥隊"]
  
  const players = [];
  for (let i = 1; i < data.length; i++) {
    let rYM = data[i][0];
    if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    else rYM = String(rYM).trim().substring(0, 10);

    const ymMatch = (rYM === String(yearMonth)) ||
                    (rYM.length === 7 && yearMonth.startsWith(rYM)) ||
                    (String(yearMonth).substring(0, 7) === String(rYM).substring(0, 7));
    if (ymMatch) {
      const aRaw    = String(data[i][4] || "").trim();  // 「區」欄位
      const roleRaw = String(data[i][2] || "").trim();  // 「身份」欄位（可能含猛禽/小鳥/鳥蛋）
      // 優先從「區」讀取技能，若欄位空白或無關鍵字，改從「身份」讀取，否則預設小鳥
      const sk = (
        ["猛禽", "小鳥", "鳥蛋"].find(s => aRaw.includes(s)) ||
        ["猛禽", "小鳥", "鳥蛋"].find(s => roleRaw.includes(s)) ||
        "小鳥"
      );
      players.push({ rowIdx: i + 1, name: data[i][1], team: String(data[i][3] || "").trim(), area: aRaw, skill: sk });
    }
  }

  if (mode === "lottery") {
    if (players.length === 0) return { status: "error", message: "無資料" };
    const g = {};
    players.forEach(p => { if (!g[p.team]) g[p.team] = []; g[p.team].push(p); });
    const keys = Object.keys(g).filter(n => n !== "").sort(() => Math.random() - 0.5);
    keys.forEach((k, idx) => {
      const target = (idx < 4) ? "粉紅爆(C場)" : (idx < 8) ? "雙打小可愛(B場)" : "尼克半熟(A場)";
      g[k].forEach(p => sheet.getRange(p.rowIdx, 5).setValue(target));
    });
    return { status: "success", message: "鳥樂賽分組完成！" };
  }

  // 風樂賣以外的模式：嚴格限制剛好 24 人 (只支援四個標準隊伍)
  if (players.length !== 24) {
    return { status: "error", message: `常規賽分組需要剛好 24 人（目前 ${players.length} 人）。\n請確認報名人數後再執行分組。` };
  }

  let mConfig = [];
  if (mode === "egg") {
    mConfig = [{ area: "孵蛋 A", skills: { "猛禽": 1, "鳥蛋": 1 } }, { area: "孵蛋 B", skills: { "猛禽": 1, "鳥蛋": 1 } }, { area: "小鳥區", skills: { "小鳥": 2 } }];
  } else if (mode === "fox") {
    mConfig = [{ area: "狐狸 A", skills: { "猛禽": 1, "小鳥": 1 } }, { area: "狐狸 B", skills: { "猛禽": 1, "小鳥": 1 } }, { area: "鳥蛋區", skills: { "鳥蛋": 2 } }];
  } else if (mode === "duck") {
    mConfig = [{ area: "猛禽區", skills: { "猛禽": 2 } }, { area: "醬板鴨 A", skills: { "小鳥": 1, "鳥蛋": 1 } }, { area: "醬板鴨 B", skills: { "小鳥": 1, "鳥蛋": 1 } }];
  } else {
    mConfig = [{ area: "猛禽區", skills: { "猛禽": 2 } }, { area: "小鳥區", skills: { "小鳥": 2 } }, { area: "鳥蛋區", skills: { "鳥蛋": 2 } }];
  }

  const tSkill = {}; TEAMS.forEach(t => tSkill[t] = { "猛禽": 0, "小鳥": 0, "鳥蛋": 0 });
  const totalReq = { "猛禽": 0, "小鳥": 0, "鳥蛋": 0 };
  mConfig.forEach(c => Object.keys(c.skills).forEach(s => totalReq[s] += c.skills[s]));

  players.forEach(p => { if (TEAMS.includes(p.team)) tSkill[p.team][p.skill]++; });
  for (const t of TEAMS) {
    for (const s of ["猛禽", "小鳥", "鳥蛋"]) {
      if (tSkill[t][s] > totalReq[s]) return { status: "error", message: `${t} ${s}過多 (${tSkill[t][s]}/${totalReq[s]})` };
    }
  }

  const slots = mConfig.map(() => TEAMS.map(() => []));
  // 有指定隊名但區域不符當前模式 → 進「隊伍專屬水桶」，保留隊名
  const teamPool = {};
  TEAMS.forEach(t => { teamPool[t] = { "猛禽": [], "小鳥": [], "鳥蛋": [] }; });
  // 完全未指定隊名 → 進「通用水桶」，隨機分配
  const generalPool = { "猛禽": [], "小鳥": [], "鳥蛋": [] };

  players.forEach(p => {
    const tIdx = TEAMS.indexOf(p.team);
    const aIdx = mConfig.findIndex(c => p.area.includes(c.area));

    if (tIdx !== -1 && aIdx !== -1) {
      // 隊名與區域都合法 → 鎖定就地
      if (slots[aIdx][tIdx].length < 2 && mConfig[aIdx].skills[p.skill]) {
        slots[aIdx][tIdx].push(p);
      } else {
        // 鎖定失敗（超額或實力不符） → 回落至隊伍專屬水桶
        teamPool[p.team][p.skill].push(p);
      }
    } else if (tIdx !== -1) {
      // 有隊名但區域不符當前模式 → 隊伍專屬水桶
      teamPool[p.team][p.skill].push(p);
    } else {
      // 完全未指定 → 通用水桶
      generalPool[p.skill].push(p);
    }
  });

  // 打亂通用水桶
  ["猛禽", "小鳥", "鳥蛋"].forEach(sk => generalPool[sk].sort(() => Math.random() - 0.5));

  const upd = [];
  for (let a = 0; a < mConfig.length; a++) {
    const req = mConfig[a].skills;
    for (let t = 0; t < TEAMS.length; t++) {
      const teamName = TEAMS[t];
      Object.keys(req).forEach(sk => {
        const needed = req[sk] - slots[a][t].filter(p => p.skill === sk).length;
        for (let k = 0; k < needed; k++) {
          let p = null;
          // 優先從「該隊伍的專屬水桶」抽取（保留匯入時指定的隊名）
          if (teamPool[teamName][sk].length > 0) {
            p = teamPool[teamName][sk].pop();
          } else {
            // 再從通用水桶抽取
            p = generalPool[sk].length > 0 ? generalPool[sk].pop() : null;
          }
          if (p) {
            p.team = teamName;
            p.area = mConfig[a].area;
            upd.push(p);
          }
        }
      });
    }
  }

  upd.forEach(p => {
    sheet.getRange(p.rowIdx, 4).setValue(p.team);
    sheet.getRange(p.rowIdx, 5).setValue(p.area);
  });

  // 計算剩餘未被分配的人數
  const leftGeneral = Object.values(generalPool).flat().length;
  const leftTeam = TEAMS.reduce((sum, t) => sum + Object.values(teamPool[t]).flat().length, 0);
  const left = leftGeneral + leftTeam;
  return {
    status: left > 0 ? "warning" : "success",
    message: left > 0
      ? `分組完成，但有 ${left} 人因實力不符或隊伍已滿無法自動填入，請手動調整。`
      : "智慧分組完成！已保留匯入時指定的隊名。"
  };
}



/**
 * 統計月度積點
 * 包含：循環賽積分、淘汰賽積分、手動積分、上月結餘。
 */
function logicCalculatePoints(yearMonth, manualData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 取得「距離本日最近且早於本日」的歷史紀錄 (不再限制於上個月)
  let pSheet = ss.getSheetByName(CONFIG.SHEET_POINTS);
  const prevBalances = {};
  
  if (pSheet) {
    const pData = pSheet.getDataRange().getValues();
    if (pData.length > 1) {
      const headers = pData[0];
      const ymIdx = headers.indexOf("年月");
      const nameIdx = headers.indexOf("姓名");
      const totalIdx = headers.indexOf("累積積點");
      
      // 找出所有早於 current date 的唯一日期，並由晚到早排序
      const targetDate = new Date(yearMonth);
      let latestPrevDate = null;
      let latestPrevDateStr = "";
      
      for (let i = 1; i < pData.length; i++) {
        let rowDate = pData[i][ymIdx];
        if (!(rowDate instanceof Date)) rowDate = new Date(rowDate);
        
        if (rowDate < targetDate) {
          if (!latestPrevDate || rowDate > latestPrevDate) {
            latestPrevDate = rowDate;
            latestPrevDateStr = Utilities.formatDate(rowDate, CONFIG.TIMEZONE, "yyyy-MM-dd");
          }
        }
      }
      
      // 如果找到了最近的歷史日期，抓取該日期的所有最後結餘
      if (latestPrevDateStr) {
        for (let i = 1; i < pData.length; i++) {
          let rowDate = pData[i][ymIdx];
          const rowDateStr = rowDate instanceof Date ? 
            Utilities.formatDate(rowDate, CONFIG.TIMEZONE, "yyyy-MM-dd") : String(rowDate).substring(0, 10);
            
          if (rowDateStr === latestPrevDateStr) {
            prevBalances[pData[i][nameIdx]] = parseInt(pData[i][totalIdx]) || 0;
          }
        }
      }
    }
  }

  // 1.1 過濾非官方球員 (僅藍、黑、青、粉鳥隊計算積分)
  const officialTeams = ["藍鳥隊", "黑鳥隊", "青鳥隊", "粉鳥隊"];
  
  // 3. 初始化全體球員字典 mapping
  const playersMap = {}; 
  
  // 先把有舊餘額的球員塞入 (即使這個月沒報名)
  Object.keys(prevBalances).forEach(name => {
    playersMap[name] = { 
      name: name, team: "", area: "", 
      rrRank: "-", elimRank: "-", 
      currPts: prevBalances[name], 
      guessPts: 0, refPts: 0, rrPts: 0, elimPts: 0, totalPts: 0 
    };
  });
  
  // 讀取當月報名清單
  const currReg = helperGetData(CONFIG.SHEET_REGISTRATION, yearMonth);
  currReg.forEach(p => {
    const name = p["姓名"];
    if (!playersMap[name]) {
      playersMap[name] = { 
        name: name, team: p["隊名"] || "", area: p["區"] || "", 
        rrRank: p["循環名次"] || "-", elimRank: p["淘汰名次"] || "-", 
        currPts: prevBalances[name] || 0, 
        guessPts: 0, refPts: 0, rrPts: 0, elimPts: 0, totalPts: 0 
      };
    } else {
      playersMap[name].team = p["隊名"] || playersMap[name].team;
      playersMap[name].area = p["區"] || playersMap[name].area;
      if (p["循環名次"]) playersMap[name].rrRank = p["循環名次"];
      if (p["淘汰名次"]) playersMap[name].elimRank = p["淘汰名次"];
    }
  });

  // 套用手動給分 (猜隊, 裁判, 包含鳥巢隊設定)
  if (manualData) {
    Object.keys(manualData).forEach(name => {
      if (!playersMap[name]) {
         playersMap[name] = { 
           name: name, 
           currPts: prevBalances[name] || 0, 
           guessPts: 0, 
           refPts: 0, 
           rrPts: 0, 
           elimPts: 0, 
           totalPts: 0, 
           team: manualData[name].team || "", 
           area: manualData[name].area || "", 
           rrRank: "-", 
           elimRank: "-" 
         };
      } else {
         if (manualData[name].team) playersMap[name].team = manualData[name].team;
         if (manualData[name].area) playersMap[name].area = manualData[name].area;
      }
      playersMap[name].guessPts += parseInt(manualData[name].guess) || 0;
      playersMap[name].refPts += parseInt(manualData[name].ref) || 0;
    });
  }

  // 讀取當月賽程與追分賽紀錄 (用於過濾與自動註冊)
  const rrMatches = helperGetData(CONFIG.SHEET_ROUND_ROBIN, yearMonth);
  const chasingMatches = helperGetData(CONFIG.SHEET_CHASING, yearMonth);

  // 防呆優化：若外部匯入賽程且無報名名單，自動掃描並建立球員基本資料
  const autoRegisterPlayer = (pName, teamName, areaName) => {
    if (pName && pName !== "待定" && !playersMap[pName]) {
      playersMap[pName] = {
        name: pName,
        team: teamName || "",
        area: areaName || "",
        rrRank: "-",
        elimRank: "-",
        currPts: prevBalances[pName] || 0,
        guessPts: 0,
        refPts: 0,
        rrPts: 0,
        elimPts: 0,
        totalPts: 0
      };
    }
  };

  rrMatches.forEach(m => {
    autoRegisterPlayer(cleanPlayerName(m["A隊員1"]), m["A隊名"], m["區"]);
    autoRegisterPlayer(cleanPlayerName(m["A隊員2"]), m["A隊名"], m["區"]);
    autoRegisterPlayer(cleanPlayerName(m["B隊員1"]), m["B隊名"], m["區"]);
    autoRegisterPlayer(cleanPlayerName(m["B隊員2"]), m["B隊名"], m["區"]);
  });

  chasingMatches.forEach(m => {
    autoRegisterPlayer(cleanPlayerName(m["A隊員1"]), m["A隊名"], m["區"]);
    autoRegisterPlayer(cleanPlayerName(m["A隊員2"]), m["A隊名"], m["區"]);
    autoRegisterPlayer(cleanPlayerName(m["A隊員3"]), m["A隊名"], m["區"]);
    autoRegisterPlayer(cleanPlayerName(m["B隊員1"]), m["B隊名"], m["區"]);
    autoRegisterPlayer(cleanPlayerName(m["B隊員2"]), m["B隊名"], m["區"]);
    autoRegisterPlayer(cleanPlayerName(m["B隊員3"]), m["B隊名"], m["區"]);
  });

  // 4. 計算當月循環賽每場得失分
  // 動態定義積分表 (包含常規與混合區)
  const areaPoints = {
    "猛禽": { win: 100, lose: 50 }, "猛禽區": { win: 100, lose: 50 },
    "小鳥": { win: 80, lose: 40 }, "小鳥區": { win: 80, lose: 40 },
    "鳥蛋": { win: 60, lose: 30 }, "鳥蛋區": { win: 60, lose: 30 },
    "孵蛋": { win: 80, lose: 40 }, "狐狸": { win: 90, lose: 45 }, "醬板鴨": { win: 70, lose: 35 }
  };
  
  rrMatches.forEach(m => {
    const sA = parseInt(m["A隊比分"]) || 0;
    const sB = parseInt(m["B隊比分"]) || 0;
    if (sA === 0 && sB === 0) return; 
    
    const area = String(m["區"]);
    let ptsCfg = { win: 0, lose: 0 };
    const isTeamMatch = area.includes("團體");
    
    // 判斷是否為鳥樂賽 (場地關鍵字)
    const isLottery = area.includes("場");
    if (isLottery) {
        ptsCfg = { win: 100, lose: 50 };
    } else if (area.includes("男雙")) {
        ptsCfg = { win: 100, lose: 50 };
    } else if (area.includes("女雙")) {
        ptsCfg = { win: 50, lose: 25 };
    } else if (isTeamMatch) {
        // 團體賽：點數於下方依球員分區動態決定
    } else {
        // 常規區模糊匹配
        const key = Object.keys(areaPoints).find(k => area.includes(k));
        ptsCfg = areaPoints[key] || { win: 0, lose: 0 };
    }
    
    const teamAPlayers = [m["A隊員1"], m["A隊員2"]].map(cleanPlayerName).filter(p => p && p !== "待定");
    const teamBPlayers = [m["B隊員1"], m["B隊員2"]].map(cleanPlayerName).filter(p => p && p !== "待定");

    teamAPlayers.forEach(p => {
      if (playersMap[p]) {
        let currentPtsCfg = ptsCfg;
        if (isTeamMatch) {
          const pArea = String(playersMap[p].area);
          if (pArea.includes("猛禽")) {
            currentPtsCfg = { win: 200, lose: 100 };
          } else if (pArea.includes("鳥蛋")) {
            currentPtsCfg = { win: 120, lose: 60 };
          } else {
            currentPtsCfg = { win: 160, lose: 80 }; // 預設為小鳥
          }
        }
        
        const isOfficialOrTeamMatch = isLottery || 
          area.includes("男雙") || 
          area.includes("女雙") || 
          officialTeams.includes(playersMap[p].team) ||
          isTeamMatch;
          
        if (isOfficialOrTeamMatch) {
          const pPts = (sA > sB) ? currentPtsCfg.win : currentPtsCfg.lose;
          playersMap[p].rrPts += pPts;
        }
      }
    });

    teamBPlayers.forEach(p => {
      if (playersMap[p]) {
        let currentPtsCfg = ptsCfg;
        if (isTeamMatch) {
          const pArea = String(playersMap[p].area);
          if (pArea.includes("猛禽")) {
            currentPtsCfg = { win: 200, lose: 100 };
          } else if (pArea.includes("鳥蛋")) {
            currentPtsCfg = { win: 120, lose: 60 };
          } else {
            currentPtsCfg = { win: 160, lose: 80 }; // 預設為小鳥
          }
        }
        
        const isOfficialOrTeamMatch = isLottery || 
          area.includes("男雙") || 
          area.includes("女雙") || 
          officialTeams.includes(playersMap[p].team) ||
          isTeamMatch;
          
        if (isOfficialOrTeamMatch) {
          const pPts = (sB > sA) ? currentPtsCfg.win : currentPtsCfg.lose;
          playersMap[p].rrPts += pPts;
        }
      }
    });
  });

  // 4.1 依目前預賽結果排序規則回填隊伍循環名次至每位球員
  const rrRankLookup = helperBuildRoundRobinRankLookup(yearMonth);
  const cleanTeamName = (value) => String(value || "").replace(/\s+/g, "").trim();
  Object.keys(playersMap).forEach(name => {
    const teamKey = cleanTeamName(playersMap[name].team);
    if (teamKey && rrRankLookup[teamKey]) {
      playersMap[name].rrRank = rrRankLookup[teamKey];
    }
  });

  // 5. 計算當月淘汰賽名次給分
  const finals = chasingMatches.filter(m => 
    (String(m["區"]).includes("冠軍賽") || String(m["區"]).includes("季軍賽") || String(m["區"]).includes("56名賽")) && 
    String(m["比賽狀態"]).includes("已完賽")
  );
  
  const finalRanks = {}; 
  finals.forEach(m => {
    const area = m["區"];
    if (!finalRanks[area]) finalRanks[area] = { sA: -1, sB: -1, teamA: m["A隊名"], teamB: m["B隊名"] };
    const sA = parseInt(m["A隊比分"]) || 0;
    const sB = parseInt(m["B隊比分"]) || 0;
    if (sA > finalRanks[area].sA || sB > finalRanks[area].sB) {
      finalRanks[area].sA = sA; finalRanks[area].sB = sB;
    }
  });

  const elimPointsMapping = {}; 
  const elimRankMapping = {}; 

  Object.keys(finalRanks).forEach(area => {
    const r = finalRanks[area];
    const isChamp = area.includes("冠軍賽");
    const is56 = area.includes("56名賽");

    // 排除男雙與女雙的標準淘汰給分邏輯 (男雙、女雙採每場勝負點數累計，在下方處理)
    if (area.includes("男雙") || area.includes("女雙")) return;

    if (r.sA > r.sB) {
      if (isChamp) { 
        elimPointsMapping[r.teamA] = 300; elimRankMapping[r.teamA] = "冠軍"; 
        elimPointsMapping[r.teamB] = 250; elimRankMapping[r.teamB] = "亞軍"; 
      } else if (is56) {
        elimPointsMapping[r.teamA] = 100; elimRankMapping[r.teamA] = "第五名";
        elimPointsMapping[r.teamB] = 100; elimRankMapping[r.teamB] = "第六名";
      } else { 
        elimPointsMapping[r.teamA] = 200; elimRankMapping[r.teamA] = "季軍"; 
        elimPointsMapping[r.teamB] = 150; elimRankMapping[r.teamB] = "殿軍"; 
      }
    } else {
      if (isChamp) { 
        elimPointsMapping[r.teamB] = 300; elimRankMapping[r.teamB] = "冠軍"; 
        elimPointsMapping[r.teamA] = 250; elimRankMapping[r.teamA] = "亞軍"; 
      } else if (is56) {
        elimPointsMapping[r.teamB] = 100; elimRankMapping[r.teamB] = "第五名";
        elimPointsMapping[r.teamA] = 100; elimRankMapping[r.teamA] = "第六名";
      } else { 
        elimPointsMapping[r.teamB] = 200; elimRankMapping[r.teamB] = "季軍"; 
        elimPointsMapping[r.teamA] = 150; elimRankMapping[r.teamA] = "殿軍"; 
      }
    }
  });

  // 分配積分至所有隊員 (僅常規與特別賽)
  chasingMatches.forEach(m => {
    const tA = m["A隊名"]; const tB = m["B隊名"];
    const isLotteryMatch = String(m["區"]).includes("猛禽") || String(m["區"]).includes("小鳥");
    const area = String(m["區"]);
    
    // 排除男雙與女雙，由下方特別計分處理
    if (area.includes("男雙") || area.includes("女雙")) return;

    [tA, tB].forEach(team => {
      if (elimPointsMapping[team]) {
        const prefix = (team === tA) ? "A隊員" : "B隊員";
        [1, 2, 3].forEach(num => {
          const pName = cleanPlayerName(m[prefix + num]);
          if (pName && pName !== "待定" && playersMap[pName]) {
            if (isLotteryMatch || officialTeams.includes(playersMap[pName].team)) {
                playersMap[pName].elimPts = elimPointsMapping[team];
                playersMap[pName].elimRank = elimRankMapping[team];
            }
          }
        });
      }
    });
  });

  // 5.1 針對「男雙」與「女雙」的追分賽/決賽進行每場勝負分累加
  chasingMatches.forEach(m => {
    const area = String(m["區"]);
    if (area.includes("男雙") || area.includes("女雙")) {
      const sA = parseInt(m["A隊比分"]) || 0;
      const sB = parseInt(m["B隊比分"]) || 0;
      if (sA === 0 && sB === 0) return;
      
      let ptsCfg = { win: 0, lose: 0 };
      if (area.includes("男雙")) {
        ptsCfg = { win: 100, lose: 50 };
      } else if (area.includes("女雙")) {
        ptsCfg = { win: 50, lose: 25 };
      }
      
      const aPts = (sA > sB) ? ptsCfg.win : ptsCfg.lose;
      const bPts = (sB > sA) ? ptsCfg.win : ptsCfg.lose;
      
      const teamAPlayers = [m["A隊員1"], m["A隊員2"], m["A隊員3"]].map(cleanPlayerName).filter(p => p && p !== "待定");
      const teamBPlayers = [m["B隊員1"], m["B隊員2"], m["B隊員3"]].map(cleanPlayerName).filter(p => p && p !== "待定");
      
      teamAPlayers.forEach(p => {
        if (playersMap[p]) {
          playersMap[p].elimPts += aPts;
          if (area.includes("冠軍賽")) {
            playersMap[p].elimRank = (sA > sB) ? "冠軍" : "亞軍";
          } else if (area.includes("季軍賽")) {
            playersMap[p].elimRank = (sA > sB) ? "季軍" : "殿軍";
          } else if (area.includes("56名賽")) {
            playersMap[p].elimRank = (sA > sB) ? "第五名" : "第六名";
          }
        }
      });
      
      teamBPlayers.forEach(p => {
        if (playersMap[p]) {
          playersMap[p].elimPts += bPts;
          if (area.includes("冠軍賽")) {
            playersMap[p].elimRank = (sB > sA) ? "冠軍" : "亞軍";
          } else if (area.includes("季軍賽")) {
            playersMap[p].elimRank = (sB > sA) ? "季軍" : "殿軍";
          } else if (area.includes("56名賽")) {
            playersMap[p].elimRank = (sB > sA) ? "第五名" : "第六名";
          }
        }
      });
    }
  });

  // 5.2 針對團體賽，直接從報名表的「淘汰名次」給予淘汰分
  Object.keys(playersMap).forEach(name => {
    const p = playersMap[name];
    const isTeamPlayer = p.area && (p.area.includes("團體") || 
      ["猛禽總部隊", "大哥隊", "雪精靈隊", "燒鳥隊", "寒冬防守", "炎夏爆擊", "春風快攻", "秋風控場"].includes(p.team));
    
    if (isTeamPlayer) {
      const rank = String(p.elimRank || "").trim();
      if (rank === "1" || rank === "冠軍") {
        p.elimPts = 400;
        p.elimRank = "冠軍";
      } else if (rank === "2" || rank === "亞軍") {
        p.elimPts = 350;
        p.elimRank = "亞軍";
      } else if (rank === "3" || rank === "季軍") {
        p.elimPts = 300;
        p.elimRank = "季軍";
      } else if (rank === "4" || rank === "殿軍") {
        p.elimPts = 250;
        p.elimRank = "殿軍";
      }
    }
  });

  // 6. 加總並排序結果陣列
  const finalArray = [];
  Object.keys(playersMap).forEach(name => {
    const p = playersMap[name];
    p.totalPts = p.currPts + p.guessPts + p.refPts + p.rrPts + p.elimPts;
    finalArray.push(p);
  });

  finalArray.sort((a, b) => b.totalPts - a.totalPts);

  // 7. 將結果寫入積點統計表
  pSheet = ss.getSheetByName(CONFIG.SHEET_POINTS);
  if (!pSheet) {
    pSheet = ss.insertSheet(CONFIG.SHEET_POINTS);
  }
  
  // 清除本月的防呆處理
  const pData = pSheet.getDataRange().getValues();
  for (let i = pData.length - 1; i >= 1; i--) {
    let rYM = pData[i][0];
    if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    if (String(rYM).substring(0, 10) === String(yearMonth).substring(0, 10)) pSheet.deleteRow(i + 1);
  }
  
  if (pSheet.getLastRow() === 0) {
    pSheet.appendRow([
      "年月", "排名", "姓名", "隊名", "區", 
      "循環名次", "淘汰名次", 
      "目前積點", "猜隊", "裁判", "循環分", "淘汰分", "累積積點"
    ]);
  }
  
  // 核心效能優化：改為 2D 陣列批次寫入 (setValues) 而不是一個一個 appendRow
  if (finalArray.length > 0) {
    const rowsToAppend = finalArray.map((p, idx) => [
      yearMonth, 
      idx + 1,        
      p.name, 
      p.team, 
      p.area, 
      p.rrRank,
      p.elimRank,
      p.currPts,
      p.guessPts,
      p.refPts,
      p.rrPts,
      p.elimPts,
      p.totalPts
    ]);
    
    const maxCols = pSheet.getMaxColumns();
    const neededCols = rowsToAppend[0].length;
    if (maxCols < neededCols) {
      pSheet.insertColumnsAfter(maxCols, neededCols - maxCols);
    }
    
    pSheet.getRange(pSheet.getLastRow() + 1, 1, rowsToAppend.length, neededCols).setValues(rowsToAppend);
  }
  
  return { 
    status: "success", 
    message: "月結積點計算完成，共統計 " + finalArray.length + " 位球員！", 
    data: finalArray 
  };
}

/**
 * 取得前台顯示專用的積點紀錄 (純讀取，不計算，效能優化)
 */
function helperGetPointsRecords(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_POINTS);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const result = [];
  const headers = data[0];
  
  // 找到對應月分且依照排名排序 (Excel 裡已經排好了)
  for (let i = 1; i < data.length; i++) {
    let rYM = data[i][0];
    if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    
    if (String(rYM).substring(0, 10) === String(yearMonth).substring(0, 10)) {
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = data[i][idx];
      });
      // 轉換為與 calculatePoints 回傳格式一致的物件
      result.push({
        name: obj["姓名"],
        team: obj["隊名"],
        area: obj["區"],
        rrRank: obj["循環名次"],
        elimRank: obj["淘汰名次"],
        currPts: obj["目前積點"],
        guessPts: obj["猜隊"],
        refPts: obj["裁判"],
        rrPts: obj["循環分"],
        elimPts: obj["淘汰分"],
        totalPts: obj["累積積點"]
      });
    }
  }
  return result;
}

/**
 * 取得所有特殊紀錄
 */
function helperGetSpecialRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_SPECIAL);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const result = [];
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    const ob = {};
    headers.forEach((h, idx) => {
      let val = data[i][idx];
      if (idx === 0 && val instanceof Date) {
        val = Utilities.formatDate(val, CONFIG.TIMEZONE, "yyyy-MM-dd");
      }
      ob[h] = val;
    });
    result.push(ob);
  }
  return result;
}

/**
 * 儲存公佈欄紀錄 (原特殊紀錄)
 */
function logicSaveSpecialRecords(yearMonth, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_SPECIAL);
  
  // 若工作表不存在，則重新建立新格式
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_SPECIAL);
    sheet.appendRow(["年月", "公佈內容"]);
  } else {
    // 檢查標題是否為舊格式，如果是則重置為新格式
    const headers = sheet.getDataRange().getValues()[0];
    if (headers && headers.includes("類型")) {
      sheet.clear();
      sheet.appendRow(["年月", "公佈內容"]);
    }
  }

  // 刪除該月舊資料
  const oldData = sheet.getDataRange().getValues();
  for (let i = oldData.length - 1; i >= 1; i--) {
    let rYM = oldData[i][0];
    if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    if (String(rYM).substring(0, 10) === String(yearMonth).substring(0, 10)) {
      sheet.deleteRow(i + 1);
    }
  }

  // 寫入新公佈內容
  if (data && data.content) {
    sheet.appendRow([yearMonth, data.content]);
  }

  return { status: "success", message: "公佈欄內容已成功發布！" };
}

/**
 * 同步球員到獨立的球員資料庫中
 */
function syncPlayerDatabase(names) {
  if (!names || names.length === 0) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_PLAYER_DB);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_PLAYER_DB);
    sheet.appendRow(["姓名", "照片網址"]);
  }

  const data = sheet.getDataRange().getValues();
  const existingNames = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const n = String(data[i][0]).trim();
    if (n) existingNames.add(n);
  }

  const newNames = [...new Set(names)].filter(n => !existingNames.has(String(n).trim()));
  newNames.forEach(n => {
    sheet.appendRow([n.trim(), ""]);
  });
}

/**
 * 獲取系統中最近的一個歷史賽事日期 (<= 今日)
 */
function logicGetLatestDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = [CONFIG.SHEET_CHASING, CONFIG.SHEET_ROUND_ROBIN, CONFIG.SHEET_REGISTRATION];
  const today = new Date();
  today.setHours(23, 59, 59, 999); // 包含今日
  
  let allDates = [];
  
  sheetsToCheck.forEach(sName => {
    const s = ss.getSheetByName(sName);
    if (!s) return;
    const data = s.getDataRange().getValues();
    if (data.length <= 1) return;
    
    // 年月欄位索引
    const ymIdx = data[0].indexOf("年月");
    if (ymIdx === -1) return;
    
    for (let i = 1; i < data.length; i++) {
        let val = data[i][ymIdx];
        if (!val) continue;
        let d = val instanceof Date ? val : new Date(val);
        if (!isNaN(d.getTime()) && d <= today) {
            allDates.push(d);
        }
    }
  });
  
  if (allDates.length === 0) return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd");
  
  // 排序並取最晚日期
  allDates.sort((a, b) => b - a);
  return Utilities.formatDate(allDates[0], CONFIG.TIMEZONE, "yyyy-MM-dd");
}

/**
 * 獲取球員名單與照片對應圖
 */
function logicGetPlayersInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PLAYER_DB);
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const playerMap = {};
  for (let i = 1; i < data.length; i++) {
    const n = String(data[i][0]).trim();
    const photoUrl = String(data[i][1]).trim();
    if (n) playerMap[n] = photoUrl;
  }
  return playerMap;
}

/**
 * 上傳球員照片 (Base64) 至 Google Drive
 */
function logicUploadPhoto(data) {
  try {
    const { name, base64Data, mimeType } = data;
    if (!name || !base64Data) {
      return { status: "error", message: "缺少必要參數 (姓名或圖片資料)" };
    }

    // 1. 尋找或建立統一照片資料夾
    const folderName = "羽球系統照片";
    let uploadFolder;
    try {
      const folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        uploadFolder = folders.next();
      } else {
        uploadFolder = DriveApp.createFolder(folderName);
        try {
          uploadFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e) {
          Logger.log("資料夾權限設定失敗: " + e.toString());
        }
      }
    } catch (err) {
      return { status: "error", message: "取得或建立資料夾失敗：" + err.toString() };
    }

    // 2. 將 base64 解碼並轉成 Blob
    const byteString = Utilities.base64Decode(base64Data.split(',')[1] || base64Data);
    let blob;
    try {
      blob = Utilities.newBlob(byteString, mimeType || "image/jpeg", `${name}_大頭貼.jpg`);
    } catch (err) {
      return { status: "error", message: "圖片編碼解析失敗：" + err.toString() };
    }

    // 3. 建立檔案並組成分享網址
    let file;
    try {
      file = uploadFolder.createFile(blob);
    } catch (err) {
      return { status: "error", message: "照片檔案寫入失敗 (可能沒有儲存空間)：" + err.toString() };
    }
    
    let sharingMsg = "";
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      sharingMsg = " (備註: 自動開放權限失敗，可能是您的 Google 帳號安全性限制，" + e.toString() + ")";
    }
    
    const fileId = file.getId();
    const photoUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;

    // 4. 更新球員資料庫
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_PLAYER_DB);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_PLAYER_DB);
      sheet.appendRow(["姓名", "照片網址"]);
    }

    const rows = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(name).trim()) {
        sheet.getRange(i + 1, 2).setValue(photoUrl);
        found = true;
        break;
      }
    }

    // 若沒發現這名球員，則新增一行
    if (!found) {
      sheet.appendRow([name.trim(), photoUrl]);
    }

    return { status: "success", message: "照片上傳成功！" + sharingMsg, photoUrl: photoUrl };
  } catch (err) {
    Logger.log("logicUploadPhoto 未預期例外: " + err.toString());
    return { status: "error", message: "上傳過程發生未預期錯誤：" + err.toString() };
  }
}

/**
 * 產生特別賽 (12隊鳥樂) 淘汰賽程表
 * 接收排序與種子名單後，按照規定的突圍 -> 晉級 -> 決戰順序寫入 SHEET_CHASING
 */
function logicGenerateLotteryKnockout(yearMonth, customizedData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!customizedData || !customizedData.ranksByArea) return { status: "error", message: "缺乏預賽名次或種子資料" };

  let sheet = ss.getSheetByName(CONFIG.SHEET_CHASING);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_CHASING);
  
  const headers = ensureChasingHeaders(sheet);
  
  // 清除舊的該月追分賽資料 (避免舊資料殘留)
  const oldData = sheet.getDataRange().getValues();
  for (let i = oldData.length - 1; i >= 1; i--) {
     let rYM = oldData[i][0];
     if (rYM instanceof Date) rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
     if (String(rYM).substring(0, 10) === String(yearMonth).substring(0, 10)) {
         sheet.deleteRow(i + 1);
     }
  }

  const headIdx = {};
  headers.forEach((h, i) => headIdx[String(h).trim()] = i);
  
  // 建立球員名單 mapping
  const rrSheet = ss.getSheetByName(CONFIG.SHEET_ROUND_ROBIN);
  const rrData = rrSheet.getDataRange().getValues();
  const rrHeadIdx = {};
  if (rrData.length > 0) {
    rrData[0].forEach((h, i) => rrHeadIdx[String(h).trim()] = i);
  }
  const teamPlayers = {};
  for(let i = 1; i < rrData.length; i++) {
     const tA = String(rrData[i][rrHeadIdx["A隊名"]] || "").trim();
     if(tA && !teamPlayers[tA]) teamPlayers[tA] = [rrData[i][rrHeadIdx["A隊員1"]], rrData[i][rrHeadIdx["A隊員2"]]];
     const tB = String(rrData[i][rrHeadIdx["B隊名"]] || "").trim();
     if(tB && !teamPlayers[tB]) teamPlayers[tB] = [rrData[i][rrHeadIdx["B隊員1"]], rrData[i][rrHeadIdx["B隊員2"]]];
  }
  function getP(tName) {
     if(!tName || tName.includes("勝") || tName.includes("敗")) return ["", ""];
     return teamPlayers[tName] || ["", ""];
  }

  // 解析前端資料
  const seedsC = customizedData.seedsC || [];
  const seedsB = customizedData.seedsB || [];
  const ranksByArea = customizedData.ranksByArea;

  // 定義種子隊伍 (sC1, sC2 為高分區種子；sB1, sB2 為低分區種子)
  const sC1 = seedsC[0] || "C種子1缺";
  const sC2 = seedsC[1] || "C種子2缺";
  const sB1 = seedsB[0] || "B種子1缺";
  const sB2 = seedsB[1] || "B種子2缺";

  const top1s = []; const top2s = []; const top3s = []; const top4s = [];
  
  // Extract groups
  for (const court in ranksByArea) {
    ranksByArea[court].forEach(s => {
      s.prelimCourt = court;
      if (s.rank === 1) top1s.push(s);
      if (s.rank === 2) top2s.push(s);
      if (s.rank === 3) top3s.push(s);
      if (s.rank === 4) top4s.push(s);
    });
  }

  // C區第一階段邏輯：1隊未中籤第1名 + 3隊第2名
  let unseededC = top1s.find(t => !seedsC.includes(t.teamName));
  if (!unseededC) unseededC = { teamName: "C未中籤第1缺" };
  let oppC = top2s.find(t => t.prelimCourt !== unseededC.prelimCourt);
  if (!oppC) oppC = top2s[0] || { teamName: "C區第2缺" }; 
  const remainingC2 = top2s.filter(t => t.teamName !== oppC.teamName);
  if (remainingC2.length < 2) remainingC2.push({ teamName: "C區第2缺" }, { teamName: "C區第2缺" });

  // B區第一階段邏輯：1隊未中籤第3名 + 3隊第4名
  let unseededB = top3s.find(t => !seedsB.includes(t.teamName));
  if (!unseededB) unseededB = { teamName: "B未中籤第3缺" };
  let oppB = top4s.find(t => t.prelimCourt !== unseededB.prelimCourt);
  if (!oppB) oppB = top4s[0] || { teamName: "B區第4缺" };
  const remainingB4 = top4s.filter(t => t.teamName !== oppB.teamName);
  if (remainingB4.length < 2) remainingB4.push({ teamName: "B區第4缺" }, { teamName: "B區第4缺" });

  let seqCount = 1;
  function pushMatch(round, area, court, teamA, teamB) {
    const row = [];
    headers.forEach(() => row.push(""));
    row[headIdx["序號"]] = seqCount++;
    row[headIdx["年月"]] = yearMonth;

    // 計算時間：從 15:30 開始，每過一 Round 加 15 分鐘 (Round 1=15:30, Round 2=15:45...)
    const roundMatch = round.match(/\d+/);
    const roundNum = roundMatch ? parseInt(roundMatch[0]) : 1;
    const offsetMins = (roundNum - 1) * 15;
    const baseTime = new Date(2000, 0, 1, 15, 30);
    baseTime.setMinutes(baseTime.getMinutes() + offsetMins);
    row[headIdx["比賽時間"]] = Utilities.formatDate(baseTime, CONFIG.TIMEZONE, "HH:mm");
    row[headIdx["輪次"]] = round; 
    row[headIdx["區"]] = area;
    row[headIdx["場地"]] = court;
    row[headIdx["A隊名"]] = teamA;
    const pA = getP(teamA); row[headIdx["A隊員1"]] = pA[0]; row[headIdx["A隊員2"]] = pA[1];
    row[headIdx["B隊名"]] = teamB;
    const pB = getP(teamB); row[headIdx["B隊員1"]] = pB[0]; row[headIdx["B隊員2"]] = pB[1];
    row[headIdx["A隊比分"]] = 0; row[headIdx["B隊比分"]] = 0;
    row[headIdx["比賽狀態"]] = "待賽";
    sheet.appendRow(row);
  }

  // ============================================================
  // 新版 4 輪 12 場：三場地全並行，猛禽=C區命名，小鳥=B區命名
  // 取消敗部賽，突圍戰落敗隊直接淘汰
  // ============================================================

  // === Round 1: 15:30，三場同步開打 ===
  // A場: 小鳥(B區) 突圍1 — 兩支第4名對決
  pushMatch("Round 1", "B區(小鳥) - 突圍1", "A場地", remainingB4[0].teamName, remainingB4[1].teamName);
  // B場: 小鳥(B區) 突圍2 — 非種子第3名 vs 來自不同區的第4名
  pushMatch("Round 1", "B區(小鳥) - 突圍2", "B場地", unseededB.teamName, oppB.teamName);
  // C場: 猛禽(C區) 突圍3 — 兩支第2名對決
  pushMatch("Round 1", "C區(猛禽) - 突圍3", "C場地", remainingC2[0].teamName, remainingC2[1].teamName);

  // === Round 2: 15:45 ===
  // A場: 小鳥(B區) 晉級1 — 種子sB1 迎戰 突圍1勝者
  pushMatch("Round 2", "B區(小鳥) - 晉級1", "A場地", sB1, "B突圍1勝者");
  // B場: 小鳥(B區) 晉級2 — 種子sB2 迎戰 突圍2勝者
  pushMatch("Round 2", "B區(小鳥) - 晉級2", "B場地", sB2, "B突圍2勝者");
  // C場: 猛禽(C區) 突圍4 — 非種子第1名 vs 剩餘第2名
  pushMatch("Round 2", "C區(猛禽) - 突圍4", "C場地", unseededC.teamName, oppC.teamName);

  // === Round 3: 16:00 ===
  // A場: 猛禽(C區) 晉級3 — 種子sC1 迎戰 突圍3勝者
  pushMatch("Round 3", "C區(猛禽) - 晉級3", "A場地", sC1, "C突圍3勝者");
  // B場: 猛禽(C區) 晉級4 — 種子sC2 迎戰 突圍4勝者
  pushMatch("Round 3", "C區(猛禽) - 晉級4", "B場地", sC2, "C突圍4勝者");
  // C場: 小鳥(B區) 季殿軍戰 — 晉級1&2 敗者爭第三
  pushMatch("Round 3", "B區(小鳥) - 季殿軍", "C場地", "B晉級1敗者", "B晉級2敗者");

  // === Round 4: 16:15，三場決戰同步開打 ===
  // A場: 小鳥(B區) 冠亞軍戰 — 晉級1&2 勝者決冠
  pushMatch("Round 4", "B區(小鳥) - 冠亞軍", "A場地", "B晉級1勝者", "B晉級2勝者");
  // B場: 猛禽(C區) 季殿軍戰 — 晉級3&4 敗者爭第三
  pushMatch("Round 4", "C區(猛禽) - 季殿軍", "B場地", "C晉級3敗者", "C晉級4敗者");
  // C場: 猛禽(C區) 冠亞軍戰 — 晉級3&4 勝者決冠
  pushMatch("Round 4", "C區(猛禽) - 冠亞軍", "C場地", "C晉級3勝者", "C晉級4勝者");



  return { status: "success" };
}

function cleanPlayerName(name) {
  if (!name) return "";
  return String(name).replace(/\([CA]\)/ig, "").trim();
}

function logicUpdateRegistrationZones(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_REGISTRATION);
  if (!sheet) return { status: "error", message: "找不到報名紀錄工作表" };
  
  const data = sheet.getDataRange().getValues();
  
  // 建立球員與其分區的映射 (根據圖片)
  const zoneMapping = {
    // 猛禽總部隊
    "阿伯": "猛禽", "Kobe": "猛禽", "建瑋": "猛禽", "Lily": "猛禽",
    "彥維": "小鳥", "哈達威": "小鳥", "寶拉": "小鳥", "庭妤": "小鳥",
    "豚豚": "鳥蛋", "龍龍": "鳥蛋",
    
    // 雪精靈隊
    "子安": "猛禽", "阿俊": "猛禽", "英忠": "猛禽", "鈞家": "猛禽",
    "Afu": "小鳥", "閃亮亮": "小鳥", "大毛": "小鳥", "薛薛": "小鳥",
    "杭杭": "鳥蛋", "佳靜": "鳥蛋",
    
    // 大哥隊
    "磊哥": "猛禽", "Jason": "猛禽", "宗霈": "猛禽", "Ken": "猛禽",
    "世峰": "小鳥", "Sandy": "小鳥", "志軒": "小鳥", "東哥": "小鳥",
    "鈺茹": "鳥蛋", "盈婷": "鳥蛋",
    
    // 燒鳥隊
    "半熟": "猛禽", "粒米": "猛禽", "靜靜": "猛禽", "文豪": "猛禽", "QQ": "猛禽",
    "嘉銘": "小鳥", "欣陵": "小鳥", "David": "小鳥", "政瑜": "小鳥",
    "葉問": "鳥蛋", "Peggy": "鳥蛋"
  };
  
  const ymIdx = data[0].indexOf("年月");
  const nameIdx = data[0].indexOf("姓名");
  const areaIdx = data[0].indexOf("區");
  
  if (ymIdx === -1 || nameIdx === -1 || areaIdx === -1) {
    return { status: "error", message: "欄位結構不完整" };
  }
  
  let updatedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    let rYM = data[i][ymIdx];
    if (rYM instanceof Date) {
      rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    } else {
      rYM = String(rYM).trim().substring(0, 10);
    }
    
    // 比對年月
    const targetYM = String(yearMonth).trim().substring(0, 10);
    if (rYM === targetYM) {
      const name = String(data[i][nameIdx]).trim();
      const cleanName = name.replace(/\([CA]\)/ig, "").trim();
      
      if (zoneMapping[cleanName]) {
        sheet.getRange(i + 1, areaIdx + 1).setValue(zoneMapping[cleanName]);
        updatedCount++;
      }
    }
  }
  
  return { status: "success", message: "成功更新了 " + updatedCount + " 位球員的分區！" };
}

function logicUpdateRegistrationRanks(yearMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_REGISTRATION);
  if (!sheet) return { status: "error", message: "找不到報名紀錄工作表" };
  
  const data = sheet.getDataRange().getValues();
  
  const rankMapping = {
    "雪精靈隊": 1,
    "燒鳥隊": 2,
    "猛禽總部隊": 3,
    "大哥隊": 4
  };
  
  const ymIdx = data[0].indexOf("年月");
  const teamIdx = data[0].indexOf("隊名");
  const elimIdx = data[0].indexOf("淘汰名次");
  
  if (ymIdx === -1 || teamIdx === -1 || elimIdx === -1) {
    return { status: "error", message: "欄位結構不完整" };
  }
  
  let updatedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    let rYM = data[i][ymIdx];
    if (rYM instanceof Date) {
      rYM = Utilities.formatDate(rYM, CONFIG.TIMEZONE, "yyyy-MM-dd");
    } else {
      rYM = String(rYM).trim().substring(0, 10);
    }
    
    const targetYM = String(yearMonth).trim().substring(0, 10);
    if (rYM === targetYM) {
      const team = String(data[i][teamIdx]).trim();
      if (rankMapping[team] !== undefined) {
        sheet.getRange(i + 1, elimIdx + 1).setValue(rankMapping[team]);
        updatedCount++;
      }
    }
  }
  
  return { status: "success", message: "成功更新了 " + updatedCount + " 位球員的淘汰名次！" };
}
