/**
 * ระบบจัดสอนแทน โรงเรียนบ้านสันดาบ
 * ปรับปรุงระบบใหม่ ครอบคลุมทุกฟังก์ชัน และอัปเดตตารางสอนปีการศึกษา 2569 เทอม 2 ครบถ้วน 100%
 */

function doGet(e) {
  // รองรับ GET API request จาก GitHub Pages (เช่น ?action=getInitialData)
  if (e && e.parameter && e.parameter.action) {
    try {
      var result = handleApiRequest(e.parameter.action, e.parameter);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบจัดสอนแทน โรงเรียนบ้านสันดาบ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var contents;
    if (e && e.postData && e.postData.contents) {
      contents = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.data) {
      contents = JSON.parse(e.parameter.data);
    } else {
      contents = e ? (e.parameter || {}) : {};
    }

    // 1. ถ้าเป็น LINE Webhook Event
    if (contents.events) {
      return handleLineWebhook(contents);
    }

    // 2. ถ้าเป็น API Request จาก GitHub Pages (bansandab.github.io)
    var action = contents.action;
    var data = contents.data || contents;
    var result = handleApiRequest(action, data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) { 
    Logger.log("doPost Error: " + err.message);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleApiRequest(action, data) {
  var result;
  try {
    if (action === 'getInitialData') {
      var raw = getInitialData();
      result = typeof raw === 'string' ? JSON.parse(raw) : raw;
    } else if (action === 'saveBulkSubstitutions') {
      result = saveBulkSubstitutions(data.rows, data.defaultReason || data.reason);
    } else if (action === 'deleteSubstitution') {
      result = deleteSubstitution(Number(data.rowId));
    } else if (action === 'updateSubstitution') {
      result = updateSubstitution(Number(data.rowId), data.updatedData);
    } else if (action === 'savePeriodSwap') {
      result = savePeriodSwap(data.record || data);
    } else if (action === 'deletePeriodSwap') {
      result = deletePeriodSwap(data.id);
    } else if (action === 'saveRoomBooking') {
      result = saveRoomBooking(data.booking || data);
    } else if (action === 'deleteRoomBooking') {
      result = deleteRoomBooking(data.id);
    } else if (action === 'testLineNotification') {
      result = testLineNotification();
    } else if (action === 'getTeacherSubSummary') {
      result = getTeacherSubSummary(data.teacherName);
    } else if (action === 'setupDatabase') {
      setupDatabase();
      result = { success: true, message: "ติดตั้งฐานข้อมูลเรียบร้อยแล้ว" };
    } else {
      result = { success: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  return result;
}

function handleLineWebhook(data) {
  try {
    if (!data.events || data.events.length === 0) {
      return ContentService.createTextOutput("OK");
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Teacher_List");

    for (var i = 0; i < data.events.length; i++) {
      var event = data.events[i];
      var source = event.source || {};
      var groupId = source.groupId || source.roomId;

      if (groupId) {
        if (sheet) {
          sheet.getRange("G1").setValue("Group ID ของกลุ่มนี้คือ:");
          sheet.getRange("H1").setValue(groupId);
        }

        var replyToken = event.replyToken;
        if (replyToken) {
          if (event.type === 'join') {
            replyLineMessage(replyToken, "👋 สวัสดีครับ! บอทระบบจัดสอนแทน โรงเรียนบ้านสันดาบ เชื่อมต่อเข้ากลุ่มเรียบร้อยแล้ว\n📌 Group ID: " + groupId + "\n(ระบบบันทึก Group ID ลง Google Sheets ให้อัตโนมัติแล้วครับ)");
          } else if (event.type === 'message' && event.message && event.message.type === 'text') {
            var txt = event.message.text.trim();
            if (txt === 'เช็คกลุ่ม' || txt === 'groupid' || txt === 'id' || txt === 'ID' || txt === 'ทดสอบ') {
              replyLineMessage(replyToken, "✅ สถานะ: บอทเชื่อมต่อกลุ่มนี้อยู่ครับ\n📌 Group ID: " + groupId + "\n📊 บันทึกลงในฐานข้อมูลระบบจัดสอนแทนเรียบร้อยแล้ว");
            }
          }
        }
      }
    }
  } catch(err) { 
    Logger.log("Webhook Error: " + err.message); 
  }
  return ContentService.createTextOutput("OK");
}

function getInitialData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tSheet = ss.getSheetByName("Teacher_List");
    var mSheet = ss.getSheetByName("Master_Timetable");
    var sSheet = ss.getSheetByName("Substitution_Record");
    var swapSheet = ss.getSheetByName("Swap_Record");
    var roomSheet = ss.getSheetByName("Room_Booking_Record");

    // ตรวจสอบและสร้างชีต Substitution_Record หากยังไม่มี
    if (!sSheet) {
      sSheet = ss.insertSheet("Substitution_Record");
      sSheet.getRange(1, 1, 1, 8).setValues([["วันที่", "คาบที่", "ครูที่ไม่อยู่", "ครูสอนแทน", "วิชา", "ห้องเรียน", "สาเหตุการลา", "วันเวลาบันทึก"]]);
    }

    if (!swapSheet) {
      swapSheet = ss.insertSheet("Swap_Record");
      swapSheet.getRange(1, 1, 1, 10).setValues([["ID", "วันที่ของคาบ", "คาบที่", "วิชา", "ระดับชั้น", "ครูผู้ขอแลก", "ครูที่แลกด้วย", "วันที่ชดเชย", "คาบที่ชดเชย", "วันเวลาบันทึก"]]);
    }

    if (!roomSheet) {
      roomSheet = ss.insertSheet("Room_Booking_Record");
      roomSheet.getRange(1, 1, 1, 8).setValues([["ID", "วันที่จอง", "วัน", "คาบที่", "ครูผู้จอง", "ระดับชั้น", "วิชา/กิจกรรม", "วันเวลาบันทึก"]]);
    }

    if (!tSheet || !mSheet) {
      return JSON.stringify({ 
        success: false, 
        error: "ยังไม่ได้ติดตั้งฐานข้อมูล กรุณาเปิด Apps Script แล้วรันฟังก์ชัน setupDatabase() ก่อนใช้งานครับ" 
      });
    }

    // ดึงข้อมูลครูพร้อมตำแหน่งและครูประจำชั้น
    var teacherRows = tSheet.getLastRow() > 1 
      ? tSheet.getRange(2, 1, tSheet.getLastRow() - 1, 4).getValues()
      : [];
      
    var teachers = teacherRows.map(function(r) {
      var hr = String(r[2] || "-").trim().replace(/\s*\(ร่วมกับ.*?\)/g, "").trim();
      return {
        name: String(r[0]).trim(),
        role: String(r[1] || "ครูผู้สอน").trim(),
        homeroom: hr,
        grade: String(r[3] || "-").trim()
      };
    }).filter(function(t) { return t.name.length > 0; });

    // ดึงตารางสอนหลัก
    var timetable = mSheet.getLastRow() > 0 
      ? mSheet.getDataRange().getValues().filter(function(r) { return r[0] && r[0] !== "ชื่อครู"; })
      : [];

    // ดึงประวัติการสอนแทน
    var subs = sSheet.getLastRow() > 1 
      ? sSheet.getDataRange().getValues().filter(function(r) { return r[0] && r[0] !== "วันที่"; }) 
      : [];

    var formattedSubs = subs.map(function(row, idx) {
      var d = row[0];
      if (d instanceof Date) {
        d = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return {
        rowId: idx + 2, // ตรงกับ Row ใน Google Sheets (เริ่มจากแถว 2)
        date: d,
        period: row[1],
        absentTeacher: row[2],
        subTeacher: row[3],
        subject: row[4],
        room: row[5],
        reason: row[6] || "-",
        timestamp: row[7] instanceof Date 
          ? Utilities.formatDate(row[7], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
          : (row[7] || "-")
      };
    });

    // ดึงข้อมูลการแลกคาบ
    var swapRows = swapSheet.getLastRow() > 1
      ? swapSheet.getDataRange().getValues().filter(function(r) { return r[0] && r[0] !== "ID"; })
      : [];
    var formattedSwaps = swapRows.map(function(r) {
      var d = r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "yyyy-MM-dd") : r[1];
      var rd = r[7] instanceof Date ? Utilities.formatDate(r[7], Session.getScriptTimeZone(), "yyyy-MM-dd") : r[7];
      return {
        id: r[0],
        date: d,
        period: r[2],
        subject: r[3],
        room: r[4],
        fromTeacher: r[5],
        toTeacher: r[6],
        returnDate: rd,
        returnPeriod: r[8],
        timestamp: r[9] instanceof Date ? Utilities.formatDate(r[9], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : r[9]
      };
    });

    // ดึงข้อมูลการจองห้อง
    var roomRows = roomSheet.getLastRow() > 1
      ? roomSheet.getDataRange().getValues().filter(function(r) { return r[0] && r[0] !== "ID"; })
      : [];
    var formattedBookings = roomRows.map(function(r) {
      var d = r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "yyyy-MM-dd") : r[1];
      return {
        id: r[0],
        date: d,
        day: r[2],
        period: r[3],
        teacher: r[4],
        grade: r[5],
        subject: r[6],
        timestamp: r[7] instanceof Date ? Utilities.formatDate(r[7], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss") : r[7]
      };
    });

    var schoolInfo = {
      schoolName: "โรงเรียนบ้านสันดาบ",
      term: "ภาคเรียนที่ 2 ปีการศึกษา 2569",
      director: "นายกฤษฎา ผลพล",
      academicHead: "นางสาวสุจิตตรา ปางสุข"
    };

    var lineGroupId = "Cce2bd2588dc0251c77376428f6a85f57";
    if (tSheet) {
      var savedId = tSheet.getRange("H1").getValue();
      if (savedId && String(savedId).trim().length > 0) {
        lineGroupId = String(savedId).trim();
      }
    }

    return JSON.stringify({ 
      success: true, 
      teachers: teachers, 
      timetable: timetable, 
      subs: formattedSubs,
      swaps: formattedSwaps,
      roomBookings: formattedBookings,
      schoolInfo: schoolInfo,
      lineGroupId: lineGroupId
    });
  } catch (err) { 
    return JSON.stringify({ success: false, error: err.message }); 
  }
}

function saveBulkSubstitutions(rows, defaultReason) {
  if (!rows || rows.length === 0) {
    return { success: false, message: "ไม่มีข้อมูลรายการสอนแทนที่ต้องบันทึก" };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Substitution_Record");
  if (!sheet) {
    sheet = ss.insertSheet("Substitution_Record");
    sheet.getRange(1, 1, 1, 8).setValues([["วันที่", "คาบที่", "ครูที่ไม่อยู่", "ครูสอนแทน", "วิชา", "ห้องเรียน", "สาเหตุการลา", "วันเวลาบันทึก"]]);
  }

  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var reasonText = defaultReason || "-";
  
  var dataToInsert = rows.map(function(r) {
    return [
      r.date, 
      r.period, 
      r.absentTeacher, 
      r.subTeacher, 
      r.subject, 
      r.room, 
      r.reason || reasonText, 
      timestamp
    ];
  });
  
  sheet.getRange(sheet.getLastRow() + 1, 1, dataToInsert.length, dataToInsert[0].length).setValues(dataToInsert);
  
  // สร้างข้อความแจ้งเตือนเข้า LINE กลุ่มแบบจัดรูปแบบสวยงาม
  rows.sort(function(a, b) { return a.period - b.period; });
  
  var periodTimes = {
    1: "08:15 - 08:30 น.",
    2: "08:30 - 09:30 น.",
    3: "09:30 - 10:30 น.",
    4: "10:30 - 11:30 น.",
    5: "12:30 - 13:30 น.",
    6: "13:30 - 14:30 น.",
    7: "14:30 - 15:30 น.",
    8: "15:30 - 16:30 น."
  };

  var lineMsg = "📣 แจ้งตารางจัดสอนแทน โรงเรียนบ้านสันดาบ\n";
  lineMsg += "📅 วันที่: " + rows[0].date + "\n";
  lineMsg += "❌ ครูที่ไม่อยู่: ครู" + rows[0].absentTeacher + "\n";
  if (reasonText && reasonText !== "-") {
    lineMsg += "📝 สาเหตุ: " + reasonText + "\n";
  }
  lineMsg += "━━━━━━━━━━━━━━━━\n";
  
  var kgPeriodTimes = {
    1: "08:30 - 09:00 น.",
    2: "09:00 - 10:00 น.",
    3: "10:00 - 10:30 น. (หรือ 09:30 - 10:30 น.)",
    4: "10:30 - 11:00 น.",
    6: "12:00 - 14:00 น.",
    7: "14:00 - 14:30 น.",
    8: "14:30 - 15:00 น.",
    9: "15:00 - 15:30 น."
  };

  rows.forEach(function(row) {
    var isKg = row.absentTeacher === "สุวนันท์ ศรสิทธิ์" || row.absentTeacher === "ธารณี สินสุข" || (row.room && row.room.indexOf("อ.") === 0);
    var pTime = isKg ? (kgPeriodTimes[row.period] ? " (" + kgPeriodTimes[row.period] + ")" : "") : (periodTimes[row.period] ? " (" + periodTimes[row.period] + ")" : "");
    var slotName = isKg ? "อนุบาล ช่วงที่ " : "คาบที่ ";
    var actName = isKg ? "กิจกรรม" : "วิชา";
    lineMsg += "⏰ " + slotName + row.period + pTime + "\n";
    lineMsg += "   📚 " + actName + ": " + row.subject + " (" + row.room + ")\n";
    lineMsg += "   ✅ ครูสอนแทน: ครู" + row.subTeacher + "\n\n";
  });
  
  lineMsg += "━━━━━━━━━━━━━━━━\n";
  lineMsg += "🙏 ขอขอบพระคุณคุณครูทุกท่านที่ร่วมมือปฏิบัติหน้าที่สอนแทนครับ";
  
  var lineResult = sendLineBroadcast(lineMsg);
  if (lineResult && !lineResult.success) {
    return { 
      success: true, 
      lineSuccess: false,
      message: "⚠️ บันทึกข้อมูลลงฐานข้อมูลเรียบร้อยแล้ว แต่ส่งแจ้งเตือนเข้ากลุ่ม LINE ไม่สำเร็จ:\n" + lineResult.error,
      lineError: lineResult.error,
      groupId: lineResult.groupId
    };
  }
  return { 
    success: true, 
    lineSuccess: true,
    message: "บันทึกข้อมูลและส่งแจ้งเตือนเข้ากลุ่ม LINE เรียบร้อยแล้ว!" 
  };
}

function sendLineBroadcast(message) {
  var token = "67uuUWZz/PgKmDJd7+Fdn2YcTCy3HoaWgYtD9/SqvU5D6iZcQkV2CwDsegJNytVLWaEQ0Lq37JFeCi5QJRcEypMRYdXx0D5eqVn9nKpMKWyfhwzrgx5S665vzqKr2FGR5hMHuJe+V3mK3Tms4Q4jmgdB04t89/1O/w1cDnyilFU="; 
  var defaultGroupId = "Cce2bd2588dc0251c77376428f6a85f57"; 
  var groupId = defaultGroupId;

  // ถ้ามีการบันทึก Group ID จาก Webhook ไว้ที่ชีต Teacher_List เซลล์ H1 ให้นำมาใช้
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tSheet = ss ? ss.getSheetByName("Teacher_List") : null;
    if (tSheet) {
      var savedId = tSheet.getRange("H1").getValue();
      if (savedId && String(savedId).trim().length > 0) {
        groupId = String(savedId).trim();
      }
    }
  } catch(e) {
    Logger.log("Error reading dynamic group ID: " + e.message);
  }

  if (!token) {
    return { success: false, error: "ไม่พบคีย์ LINE Channel Access Token", groupId: groupId };
  }
  if (!groupId) {
    return { success: false, error: "ไม่พบ Group ID ของกลุ่ม LINE (ตรวจสอบช่อง H1 ในชีต Teacher_List)", groupId: groupId };
  }

  var url = "https://api.line.me/v2/bot/message/push";
  var options = {
    "method": "post",
    "headers": { 
      "Content-Type": "application/json", 
      "Authorization": "Bearer " + token 
    },
    "payload": JSON.stringify({ 
      "to": groupId, 
      "messages": [{"type": "text", "text": message}] 
    }),
    "muteHttpExceptions": true
  };
  
  try {
    var res = UrlFetchApp.fetch(url, options);
    var code = res.getResponseCode();
    var content = res.getContentText();
    Logger.log("LINE Response (" + code + "): " + content);

    if (code === 200) {
      return { success: true, groupId: groupId };
    } else {
      var errObj = {};
      try { errObj = JSON.parse(content); } catch(ex) {}
      var errMsg = errObj.message || content;
      if (code === 429) {
        errMsg = "โควตาข้อความของ LINE Bot เดือนนี้เต็มแล้ว (Free Plan 300 ข้อความ/เดือน โดยการส่งเข้ากลุ่ม 1 ครั้งจะตัดโควตาเท่ากับจำนวนสมาชิกทุกคนในกลุ่ม)";
      } else if (code === 400) {
        errMsg = "LINE ปฏิเสธคำขอ (Group ID ไม่ถูกต้อง หรือบอทยังไม่ได้อยู่ในกลุ่มนี้): " + errMsg;
      } else if (code === 401) {
        errMsg = "Channel Access Token ไม่ถูกต้องหรือหมดอายุ";
      } else if (code === 403) {
        errMsg = "บอทไม่ได้รับอนุญาตให้ส่งข้อความเข้ากลุ่ม (ตรวจสอบสิทธิ์ใน LINE Official Account Manager)";
      }
      return { success: false, code: code, error: errMsg, groupId: groupId };
    }
  } catch (err) {
    Logger.log("LINE fetch failed: " + err.message);
    return { success: false, error: "เกิดข้อผิดพลาดในการเชื่อมต่อ LINE API: " + err.message, groupId: groupId };
  }
}

function replyLineMessage(replyToken, messageText) {
  var token = "67uuUWZz/PgKmDJd7+Fdn2YcTCy3HoaWgYtD9/SqvU5D6iZcQkV2CwDsegJNytVLWaEQ0Lq37JFeCi5QJRcEypMRYdXx0D5eqVn9nKpMKWyfhwzrgx5S665vzqKr2FGR5hMHuJe+V3mK3Tms4Q4jmgdB04t89/1O/w1cDnyilFU=";
  var url = "https://api.line.me/v2/bot/message/reply";
  var options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + token
    },
    "payload": JSON.stringify({
      "replyToken": replyToken,
      "messages": [{"type": "text", "text": messageText}]
    }),
    "muteHttpExceptions": true
  };
  try {
    UrlFetchApp.fetch(url, options);
  } catch(e) {
    Logger.log("Reply error: " + e.message);
  }
}

function testLineNotification() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tSheet = ss ? ss.getSheetByName("Teacher_List") : null;
    var groupId = "Cce2bd2588dc0251c77376428f6a85f57";
    if (tSheet) {
      var savedId = tSheet.getRange("H1").getValue();
      if (savedId && String(savedId).trim().length > 0) {
        groupId = String(savedId).trim();
      }
    }

    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var testMsg = "🔔 [ทดสอบระบบจัดสอนแทน]\nโรงเรียนบ้านสันดาบ เชื่อมต่อกลุ่ม LINE สำเร็จแล้วครับ\n⏰ เวลาทดสอบ: " + nowStr + "\n📌 Group ID: " + groupId;
    
    var result = sendLineBroadcast(testMsg);
    var quotaInfo = getLineQuotaInfo();
    
    return {
      success: result.success,
      message: result.success ? "ส่งข้อความทดสอบเข้ากลุ่ม LINE สำเร็จแล้ว!" : result.error,
      groupId: groupId,
      quota: quotaInfo
    };
  } catch(e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.message };
  }
}

function getLineQuotaInfo() {
  var token = "67uuUWZz/PgKmDJd7+Fdn2YcTCy3HoaWgYtD9/SqvU5D6iZcQkV2CwDsegJNytVLWaEQ0Lq37JFeCi5QJRcEypMRYdXx0D5eqVn9nKpMKWyfhwzrgx5S665vzqKr2FGR5hMHuJe+V3mK3Tms4Q4jmgdB04t89/1O/w1cDnyilFU=";
  var quota = { total: 0, used: 0, remaining: 0 };
  try {
    var resTotal = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/quota", {
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });
    if (resTotal.getResponseCode() === 200) {
      var qData = JSON.parse(resTotal.getContentText());
      quota.total = qData.value || 0;
    }
    
    var resUsage = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/quota/consumption", {
      headers: { "Authorization": "Bearer " + token },
      muteHttpExceptions: true
    });
    if (resUsage.getResponseCode() === 200) {
      var uData = JSON.parse(resUsage.getContentText());
      quota.used = uData.totalUsage || 0;
    }
    quota.remaining = Math.max(0, quota.total - quota.used);
  } catch(e) {
    Logger.log("Error getting quota: " + e.message);
  }
  return quota;
}

function getTeacherSubSummary(teacherName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Substitution_Record");
  if (!sheet) {
    return { total: 0, records: [] };
  }

  var data = sheet.getDataRange().getValues();
  var history = [];
  for (var i = data.length - 1; i > 0; i--) { 
    if (String(data[i][3]).trim() == String(teacherName).trim()) { 
      var dateStr = data[i][0] instanceof Date 
        ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "yyyy-MM-dd") 
        : data[i][0];
      history.push({ 
        rowId: i + 1, 
        date: dateStr, 
        period: data[i][1], 
        absentTeacher: data[i][2], 
        subject: data[i][4], 
        room: data[i][5],
        reason: data[i][6] || "-"
      });
    }
  }
  return { total: history.length, records: history };
}

function deleteSubstitution(rowId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Substitution_Record");
  if (sheet && rowId > 1 && rowId <= sheet.getLastRow()) {
    sheet.deleteRow(rowId);
    return { success: true, message: "ลบรายการเรียบร้อยแล้ว" };
  }
  return { success: false, message: "ไม่พบรายการที่ต้องการลบ" };
}

function updateSubstitution(rowId, updatedData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Substitution_Record");
  if (sheet && rowId > 1 && rowId <= sheet.getLastRow()) {
    if (updatedData.date) sheet.getRange(rowId, 1).setValue(updatedData.date);
    if (updatedData.period) sheet.getRange(rowId, 2).setValue(updatedData.period);
    if (updatedData.absentTeacher) sheet.getRange(rowId, 3).setValue(updatedData.absentTeacher);
    if (updatedData.subTeacher) sheet.getRange(rowId, 4).setValue(updatedData.subTeacher);
    if (updatedData.subject) sheet.getRange(rowId, 5).setValue(updatedData.subject);
    if (updatedData.room) sheet.getRange(rowId, 6).setValue(updatedData.room);
    if (updatedData.reason !== undefined) sheet.getRange(rowId, 7).setValue(updatedData.reason);
    return { success: true, message: "แก้ไขข้อมูลเรียบร้อยแล้ว" };
  }
  return { success: false, message: "ไม่พบรายการที่ต้องการแก้ไข" };
}

function savePeriodSwap(record) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Swap_Record") || ss.insertSheet("Swap_Record");
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 10).setValues([["ID", "วันที่ของคาบ", "คาบที่", "วิชา", "ระดับชั้น", "ครูผู้ขอแลก", "ครูที่แลกด้วย", "วันที่ชดเชย", "คาบที่ชดเชย", "วันเวลาบันทึก"]]);
  }
  var id = record.id || Date.now();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([
    id,
    record.date,
    record.period,
    record.subject,
    record.room,
    record.fromTeacher,
    record.toTeacher,
    record.returnDate || "-",
    record.returnPeriod || "-",
    timestamp
  ]);
  return { success: true, message: "บันทึกการแลกคาบเรียบร้อยแล้ว", id: id };
}

function deletePeriodSwap(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Swap_Record");
  if (!sheet) return { success: false, message: "ไม่พบชีต" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "ลบรายการแลกคาบแล้ว" };
    }
  }
  return { success: false, message: "ไม่พบรายการ" };
}

function saveRoomBooking(booking) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Room_Booking_Record") || ss.insertSheet("Room_Booking_Record");
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([["ID", "วันที่จอง", "วัน", "คาบที่", "ครูผู้จอง", "ระดับชั้น", "วิชา/กิจกรรม", "วันเวลาบันทึก"]]);
  }
  var id = booking.id || Date.now();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  sheet.appendRow([
    id,
    booking.date,
    booking.day,
    booking.period,
    booking.teacher,
    booking.grade,
    booking.subject,
    timestamp
  ]);
  return { success: true, message: "บันทึกการจองห้องเรียบร้อยแล้ว", id: id };
}

function deleteRoomBooking(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Room_Booking_Record");
  if (!sheet) return { success: false, message: "ไม่พบชีต" };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "ยกเลิกการจองห้องแล้ว" };
    }
  }
  return { success: false, message: "ไม่พบรายการ" };
}

/** 
 * ฟังก์ชันสำหรับติดตั้งฐานข้อมูลใหม่ (ปีการศึกษา 2569 เทอม 2 ถูกต้อง 100% ตรวจสอบเทียบ Word/PDF ครบถ้วน)
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. จัดการรายชื่อครูและข้อมูลประจำชั้น
  var tSheet = ss.getSheetByName("Teacher_List") || ss.insertSheet("Teacher_List");
  var existingGroupId = "";
  try {
    existingGroupId = tSheet.getRange("H1").getValue();
  } catch(e) {}

  tSheet.clear();
  var teachers = [
    ["ชื่อ-สกุล", "ตำแหน่ง/หน้าที่", "ครูประจำชั้น", "สายชั้น"],
    ["สุวรรณี นิลประดับ", "ครูผู้สอน", "ประถมศึกษาปีที่ 1", "ป.1"],
    ["คเณศ ชุนสิทธิ์", "ครูผู้สอน", "ประถมศึกษาปีที่ 1", "ป.1"],
    ["นิศารัตน์ ฉิมขันธ์", "ครูผู้สอน", "ประถมศึกษาปีที่ 2", "ป.2"],
    ["ชลลดา ซ่อนกลิ่น", "ครูผู้สอน", "ประถมศึกษาปีที่ 3", "ป.3"],
    ["พิศณุ เพียสีนุ้ย", "ครูผู้สอน", "ประถมศึกษาปีที่ 3", "ป.3"],
    ["วรัญญา พิศแสนสุวรรณ", "ครูผู้สอน", "ประถมศึกษาปีที่ 4", "ป.4"],
    ["ธนโชค ขุนพิลึก", "ครูผู้สอน", "ประถมศึกษาปีที่ 4", "ป.4"],
    ["สุจิตตรา ปางสุข", "หัวหน้าฝ่ายบริหารวิชาการ", "ประถมศึกษาปีที่ 5", "ป.5"],
    ["สุโลธร ธนาวิชิตฤกษ์", "ครูผู้สอน", "ประถมศึกษาปีที่ 5", "ป.5"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "ครูผู้สอน", "ประถมศึกษาปีที่ 6", "ป.6"],
    ["สุวนันท์ ศรสิทธิ์", "ครูประจำชั้น", "อนุบาลปีที่ 1-2", "อ.1-2"],
    ["ธารณี สินสุข", "ครูประจำชั้น", "อนุบาลปีที่ 3", "อ.3"]
  ];
  tSheet.getRange(1, 1, teachers.length, 4).setValues(teachers);
  tSheet.getRange("G1").setValue("Group ID ของกลุ่มนี้คือ:");
  var finalGroupId = (existingGroupId && String(existingGroupId).trim().length > 0 && String(existingGroupId).trim().startsWith("C"))
    ? String(existingGroupId).trim()
    : "Cce2bd2588dc0251c77376428f6a85f57";
  tSheet.getRange("H1").setValue(finalGroupId);

  // 2. จัดการตารางสอนหลัก (ปีการศึกษา 2569 เทอม 2 ตรวจสอบเทียบเอกสารครบถ้วน ไร้ข้อขัดแย้ง 100%)
  var mSheet = ss.getSheetByName("Master_Timetable") || ss.insertSheet("Master_Timetable");
  mSheet.clear();
  
  var timetableData = [
    ["ชื่อครู", "วัน", "คาบที่", "วิชา", "ห้องเรียน"],
// 1. ครูสุวรรณี นิลประดับ
    ["สุวรรณี นิลประดับ", "จันทร์", 2, "ภาษาไทย", "ป.1"],
    ["สุวรรณี นิลประดับ", "จันทร์", 4, "ภาษาไทย", "ป.3"],
    ["สุวรรณี นิลประดับ", "จันทร์", 5, "ภาษาไทยเพิ่มเติม", "ป.2"],
    ["สุวรรณี นิลประดับ", "จันทร์", 6, "ภาษาไทยเพิ่มเติม", "ป.3"],
    ["สุวรรณี นิลประดับ", "อังคาร", 2, "ภาษาไทย", "ป.4"],
    ["สุวรรณี นิลประดับ", "อังคาร", 3, "ภาษาไทย", "ป.3"],
    ["สุวรรณี นิลประดับ", "อังคาร", 4, "ภาษาไทย", "ป.1"],
    ["สุวรรณี นิลประดับ", "อังคาร", 7, "ชุมนุม", "-"],
    ["สุวรรณี นิลประดับ", "พุธ", 2, "ภาษาไทย", "ป.1"],
    ["สุวรรณี นิลประดับ", "พุธ", 4, "ภาษาไทย", "ป.4"],
    ["สุวรรณี นิลประดับ", "พุธ", 6, "ภาษาไทย", "ป.3"],
    ["สุวรรณี นิลประดับ", "พุธ", 7, "ลูกเสือ", "-"],
    ["สุวรรณี นิลประดับ", "พฤหัสบดี", 2, "ภาษาไทย", "ป.3"],
    ["สุวรรณี นิลประดับ", "พฤหัสบดี", 3, "ภาษาไทย", "ป.1"],
    ["สุวรรณี นิลประดับ", "พฤหัสบดี", 5, "ภาษาไทย", "ป.4"],
    ["สุวรรณี นิลประดับ", "พฤหัสบดี", 7, "ภาษาไทยเพิ่มเติม", "ป.1"],
    ["สุวรรณี นิลประดับ", "พฤหัสบดี", 8, "PLC", "-"],
    ["สุวรรณี นิลประดับ", "ศุกร์", 2, "ภาษาไทย", "ป.4"],
    ["สุวรรณี นิลประดับ", "ศุกร์", 3, "ภาษาไทย", "ป.3"],
    ["สุวรรณี นิลประดับ", "ศุกร์", 4, "ภาษาไทย", "ป.1"],
    ["สุวรรณี นิลประดับ", "ศุกร์", 7, "แนะแนว", "-"],
    // 2. ครูคเณศ ชุนสิทธิ์
    ["คเณศ ชุนสิทธิ์", "จันทร์", 3, "สุขศึกษา/พลศึกษา", "ป.4"],
    ["คเณศ ชุนสิทธิ์", "จันทร์", 4, "สุขศึกษา/พลศึกษา", "ป.5"],
    ["คเณศ ชุนสิทธิ์", "จันทร์", 7, "เศรษฐกิจพอเพียงฯ", "ป.1"],
    ["คเณศ ชุนสิทธิ์", "อังคาร", 4, "ศิลปะ", "ป.3"],
    ["คเณศ ชุนสิทธิ์", "อังคาร", 6, "สุขศึกษา/พลศึกษา", "ป.6"],
    ["คเณศ ชุนสิทธิ์", "อังคาร", 7, "ชุมนุม", "-"],
    ["คเณศ ชุนสิทธิ์", "พุธ", 3, "พลศึกษา", "อ.1-2"],
    ["คเณศ ชุนสิทธิ์", "พุธ", 5, "ศิลปะ", "ป.5"],
    ["คเณศ ชุนสิทธิ์", "พุธ", 6, "ศิลปะ", "ป.4"],
    ["คเณศ ชุนสิทธิ์", "พุธ", 7, "ลูกเสือ", "-"],
    ["คเณศ ชุนสิทธิ์", "พฤหัสบดี", 4, "สุขศึกษา/พลศึกษา", "ป.1"],
    ["คเณศ ชุนสิทธิ์", "พฤหัสบดี", 5, "ศิลปะ", "ป.6"],
    ["คเณศ ชุนสิทธิ์", "พฤหัสบดี", 6, "สุขศึกษา/พลศึกษา", "ป.3"],
    ["คเณศ ชุนสิทธิ์", "พฤหัสบดี", 7, "ศิลปะ", "ป.2"],
    ["คเณศ ชุนสิทธิ์", "พฤหัสบดี", 8, "PLC", "-"],
    ["คเณศ ชุนสิทธิ์", "ศุกร์", 3, "พลศึกษา", "อ.3"],
    ["คเณศ ชุนสิทธิ์", "ศุกร์", 5, "ศิลปะ", "ป.1"],
    ["คเณศ ชุนสิทธิ์", "ศุกร์", 6, "สุขศึกษา/พลศึกษา", "ป.2"],
    ["คเณศ ชุนสิทธิ์", "ศุกร์", 7, "แนะแนว", "-"],
    // 3. ครูนิศารัตน์ ฉิมขันธ์
    ["นิศารัตน์ ฉิมขันธ์", "จันทร์", 2, "คณิตศาสตร์", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "จันทร์", 3, "คณิตศาสตร์", "ป.3"],
    ["นิศารัตน์ ฉิมขันธ์", "จันทร์", 4, "ภาษาไทย", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "จันทร์", 7, "การงานอาชีพ", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "อังคาร", 2, "คณิตศาสตร์", "ป.3"],
    ["นิศารัตน์ ฉิมขันธ์", "อังคาร", 3, "คณิตศาสตร์", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "อังคาร", 4, "ภาษาไทย", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "อังคาร", 7, "ชุมนุม", "-"],
    ["นิศารัตน์ ฉิมขันธ์", "พุธ", 2, "คณิตศาสตร์", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "พุธ", 3, "ภาษาไทย", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "พุธ", 4, "คณิตศาสตร์", "ป.3"],
    ["นิศารัตน์ ฉิมขันธ์", "พุธ", 6, "การงานอาชีพ", "ป.1"],
    ["นิศารัตน์ ฉิมขันธ์", "พุธ", 7, "ลูกเสือ", "-"],
    ["นิศารัตน์ ฉิมขันธ์", "พฤหัสบดี", 2, "คณิตศาสตร์", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "พฤหัสบดี", 3, "คณิตศาสตร์", "ป.3"],
    ["นิศารัตน์ ฉิมขันธ์", "พฤหัสบดี", 5, "ภาษาไทย", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "พฤหัสบดี", 8, "PLC", "-"],
    ["นิศารัตน์ ฉิมขันธ์", "ศุกร์", 2, "คณิตศาสตร์", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "ศุกร์", 3, "ภาษาไทย", "ป.2"],
    ["นิศารัตน์ ฉิมขันธ์", "ศุกร์", 5, "คณิตศาสตร์", "ป.3"],
    ["นิศารัตน์ ฉิมขันธ์", "ศุกร์", 7, "แนะแนว", "-"],
    // 4. ครูชลลดา ซ่อนกลิ่น
    ["ชลลดา ซ่อนกลิ่น", "จันทร์", 3, "ภาษาอังกฤษ", "ป.2"],
    ["ชลลดา ซ่อนกลิ่น", "จันทร์", 5, "ภาษาอังกฤษ", "ป.3"],
    ["ชลลดา ซ่อนกลิ่น", "จันทร์", 6, "ภาษาอังกฤษ", "ป.1"],
    ["ชลลดา ซ่อนกลิ่น", "อังคาร", 2, "ภาษาอังกฤษ", "ป.1"],
    ["ชลลดา ซ่อนกลิ่น", "อังคาร", 5, "ภาษาอังกฤษ", "ป.2"],
    ["ชลลดา ซ่อนกลิ่น", "อังคาร", 6, "ภาษาอังกฤษ", "ป.3"],
    ["ชลลดา ซ่อนกลิ่น", "อังคาร", 7, "ชุมนุม", "-"],
    ["ชลลดา ซ่อนกลิ่น", "พุธ", 2, "ภาษาอังกฤษเพื่อการสื่อสาร", "ป.3"],
    ["ชลลดา ซ่อนกลิ่น", "พุธ", 3, "ภาษาอังกฤษเพื่อการสื่อสาร", "ป.1"],
    ["ชลลดา ซ่อนกลิ่น", "พุธ", 4, "ภาษาอังกฤษเพื่อการสื่อสาร", "ป.2"],
    ["ชลลดา ซ่อนกลิ่น", "พุธ", 5, "การงานอาชีพ", "ป.3"],
    ["ชลลดา ซ่อนกลิ่น", "พุธ", 7, "ลูกเสือ", "-"],
    ["ชลลดา ซ่อนกลิ่น", "พฤหัสบดี", 4, "ภาษาอังกฤษ", "ป.2"],
    ["ชลลดา ซ่อนกลิ่น", "พฤหัสบดี", 5, "ภาษาอังกฤษ", "ป.3"],
    ["ชลลดา ซ่อนกลิ่น", "พฤหัสบดี", 6, "ภาษาอังกฤษ", "ป.1"],
    ["ชลลดา ซ่อนกลิ่น", "พฤหัสบดี", 7, "การงานอาชีพ", "ป.4"],
    ["ชลลดา ซ่อนกลิ่น", "พฤหัสบดี", 8, "PLC", "-"],
    ["ชลลดา ซ่อนกลิ่น", "ศุกร์", 4, "ภาษาอังกฤษ", "ป.3"],
    ["ชลลดา ซ่อนกลิ่น", "ศุกร์", 5, "ภาษาอังกฤษ", "ป.2"],
    ["ชลลดา ซ่อนกลิ่น", "ศุกร์", 6, "ภาษาอังกฤษ", "ป.1"],
    ["ชลลดา ซ่อนกลิ่น", "ศุกร์", 7, "แนะแนว", "-"],
    // 5. ครูพิศณุ เพียสีนุ้ย
    ["พิศณุ เพียสีนุ้ย", "จันทร์", 3, "ต้านทุจริตศึกษา", "อ.3"],
    ["พิศณุ เพียสีนุ้ย", "จันทร์", 4, "หน้าที่พลเมือง", "ป.1"],
    ["พิศณุ เพียสีนุ้ย", "จันทร์", 5, "ต้านทุจริตศึกษา", "ป.5"],
    ["พิศณุ เพียสีนุ้ย", "จันทร์", 6, "ประวัติศาสตร์", "ป.2"],
    ["พิศณุ เพียสีนุ้ย", "จันทร์", 7, "เศรษฐกิจพอเพียงฯ", "ป.3"],
    ["พิศณุ เพียสีนุ้ย", "อังคาร", 3, "ต้านทุจริตศึกษา", "ป.4"],
    ["พิศณุ เพียสีนุ้ย", "อังคาร", 5, "ประวัติศาสตร์", "ป.1"],
    ["พิศณุ เพียสีนุ้ย", "อังคาร", 6, "สังคมศึกษาฯ", "ป.2"],
    ["พิศณุ เพียสีนุ้ย", "อังคาร", 7, "ชุมนุม", "-"],
    ["พิศณุ เพียสีนุ้ย", "พุธ", 3, "สังคมศึกษาฯ", "ป.3"],
    ["พิศณุ เพียสีนุ้ย", "พุธ", 5, "สังคมศึกษาฯ", "ป.1"],
    ["พิศณุ เพียสีนุ้ย", "พุธ", 6, "หน้าที่พลเมือง", "ป.2"],
    ["พิศณุ เพียสีนุ้ย", "พุธ", 7, "ลูกเสือ", "-"],
    ["พิศณุ เพียสีนุ้ย", "พฤหัสบดี", 4, "ประวัติศาสตร์", "ป.3"],
    ["พิศณุ เพียสีนุ้ย", "พฤหัสบดี", 5, "ต้านทุจริตศึกษา", "ป.1"],
    ["พิศณุ เพียสีนุ้ย", "พฤหัสบดี", 6, "ต้านทุจริตศึกษา", "ป.2"],
    ["พิศณุ เพียสีนุ้ย", "พฤหัสบดี", 7, "ต้านทุจริตศึกษา", "ป.3"],
    ["พิศณุ เพียสีนุ้ย", "พฤหัสบดี", 8, "PLC", "-"],
    ["พิศณุ เพียสีนุ้ย", "ศุกร์", 5, "ต้านทุจริตศึกษา", "ป.6"],
    ["พิศณุ เพียสีนุ้ย", "ศุกร์", 6, "หน้าที่พลเมือง", "ป.3"],
    ["พิศณุ เพียสีนุ้ย", "ศุกร์", 7, "แนะแนว", "-"],
    // 6. ครูวรัญญา พิศแสนสุวรรณ
    ["วรัญญา พิศแสนสุวรรณ", "จันทร์", 2, "ภาษาอังกฤษ", "ป.4"],
    ["วรัญญา พิศแสนสุวรรณ", "จันทร์", 3, "ภาษาอังกฤษ", "ป.5"],
    ["วรัญญา พิศแสนสุวรรณ", "จันทร์", 5, "ภาษาอังกฤษ", "ป.6"],
    ["วรัญญา พิศแสนสุวรรณ", "จันทร์", 7, "การงานอาชีพ", "ป.5"],
    ["วรัญญา พิศแสนสุวรรณ", "อังคาร", 3, "ภาษาอังกฤษ", "ป.6"],
    ["วรัญญา พิศแสนสุวรรณ", "อังคาร", 4, "ภาษาอังกฤษ", "ป.4"],
    ["วรัญญา พิศแสนสุวรรณ", "อังคาร", 5, "ภาษาอังกฤษ", "ป.5"],
    ["วรัญญา พิศแสนสุวรรณ", "อังคาร", 7, "ชุมนุม", "-"],
    ["วรัญญา พิศแสนสุวรรณ", "พุธ", 2, "ภาษาอังกฤษเพื่อการสื่อสาร", "ป.5"],
    ["วรัญญา พิศแสนสุวรรณ", "พุธ", 4, "ภาษาอังกฤษเพื่อการสื่อสาร", "ป.6"],
    ["วรัญญา พิศแสนสุวรรณ", "พุธ", 5, "ภาษาอังกฤษเพื่อการสื่อสาร", "ป.4"],
    ["วรัญญา พิศแสนสุวรรณ", "พุธ", 7, "ลูกเสือ", "-"],
    ["วรัญญา พิศแสนสุวรรณ", "พฤหัสบดี", 2, "ภาษาอังกฤษ", "ป.4"],
    ["วรัญญา พิศแสนสุวรรณ", "พฤหัสบดี", 3, "ภาษาอังกฤษ", "ป.5"],
    ["วรัญญา พิศแสนสุวรรณ", "พฤหัสบดี", 6, "ภาษาอังกฤษ", "ป.6"],
    ["วรัญญา พิศแสนสุวรรณ", "พฤหัสบดี", 8, "PLC", "-"],
    ["วรัญญา พิศแสนสุวรรณ", "ศุกร์", 2, "ภาษาอังกฤษ", "ป.6"],
    ["วรัญญา พิศแสนสุวรรณ", "ศุกร์", 3, "ภาษาอังกฤษ", "ป.4"],
    ["วรัญญา พิศแสนสุวรรณ", "ศุกร์", 4, "ภาษาอังกฤษ", "ป.5"],
    ["วรัญญา พิศแสนสุวรรณ", "ศุกร์", 6, "การงานอาชีพ", "ป.6"],
    ["วรัญญา พิศแสนสุวรรณ", "ศุกร์", 7, "แนะแนว", "-"],
    // 7. ครูธนโชค ขุนพิลึก
    ["ธนโชค ขุนพิลึก", "จันทร์", 4, "ประวัติศาสตร์", "ป.6"],
    ["ธนโชค ขุนพิลึก", "จันทร์", 5, "สังคมศึกษา", "ป.4"],
    ["ธนโชค ขุนพิลึก", "จันทร์", 6, "เศรษฐกิจพอเพียงฯ", "ป.6"],
    ["ธนโชค ขุนพิลึก", "จันทร์", 7, "ประวัติศาสตร์", "ป.4"],
    ["ธนโชค ขุนพิลึก", "อังคาร", 3, "ต้านทุจริต", "อ.1-2"],
    ["ธนโชค ขุนพิลึก", "อังคาร", 4, "สังคมศึกษา", "ป.5"],
    ["ธนโชค ขุนพิลึก", "อังคาร", 6, "หน้าที่พลเมือง", "ป.5"],
    ["ธนโชค ขุนพิลึก", "อังคาร", 7, "ชุมนุม", "-"],
    ["ธนโชค ขุนพิลึก", "พุธ", 3, "สังคมศึกษา", "ป.6"],
    ["ธนโชค ขุนพิลึก", "พุธ", 6, "ประวัติศาสตร์", "ป.5"],
    ["ธนโชค ขุนพิลึก", "พุธ", 7, "ลูกเสือ", "-"],
    ["ธนโชค ขุนพิลึก", "พฤหัสบดี", 2, "สังคมศึกษา", "ป.6"],
    ["ธนโชค ขุนพิลึก", "พฤหัสบดี", 3, "สังคมศึกษา", "ป.4"],
    ["ธนโชค ขุนพิลึก", "พฤหัสบดี", 4, "หน้าที่พลเมือง", "ป.6"],
    ["ธนโชค ขุนพิลึก", "พฤหัสบดี", 6, "หน้าที่พลเมือง", "ป.4"],
    ["ธนโชค ขุนพิลึก", "พฤหัสบดี", 7, "เศรษฐกิจพอเพียงฯ", "ป.5"],
    ["ธนโชค ขุนพิลึก", "พฤหัสบดี", 8, "PLC", "-"],
    ["ธนโชค ขุนพิลึก", "ศุกร์", 4, "เศรษฐกิจพอเพียงฯ", "ป.2"],
    ["ธนโชค ขุนพิลึก", "ศุกร์", 5, "เศรษฐกิจพอเพียงฯ", "ป.4"],
    ["ธนโชค ขุนพิลึก", "ศุกร์", 6, "สังคมศึกษา", "ป.5"],
    ["ธนโชค ขุนพิลึก", "ศุกร์", 7, "แนะแนว", "-"],
    // 8. ครูสุจิตตรา ปางสุข
    ["สุจิตตรา ปางสุข", "จันทร์", 2, "ภาษาไทย", "ป.5"],
    ["สุจิตตรา ปางสุข", "จันทร์", 3, "ภาษาไทย", "ป.6"],
    ["สุจิตตรา ปางสุข", "จันทร์", 6, "ภาษาไทยเพิ่มเติม", "ป.4"],
    ["สุจิตตรา ปางสุข", "อังคาร", 2, "ภาษาไทย", "ป.6"],
    ["สุจิตตรา ปางสุข", "อังคาร", 3, "ภาษาไทย", "ป.5"],
    ["สุจิตตรา ปางสุข", "อังคาร", 4, "วิทยาศาสตร์", "ป.6"],
    ["สุจิตตรา ปางสุข", "อังคาร", 5, "วิทยาศาสตร์", "ป.4"],
    ["สุจิตตรา ปางสุข", "อังคาร", 7, "ชุมนุม", "-"],
    ["สุจิตตรา ปางสุข", "พุธ", 2, "ภาษาไทย", "ป.6"],
    ["สุจิตตรา ปางสุข", "พุธ", 3, "วิทยาศาสตร์", "ป.4"],
    ["สุจิตตรา ปางสุข", "พุธ", 4, "ภาษาไทย", "ป.5"],
    ["สุจิตตรา ปางสุข", "พุธ", 5, "วิทยาศาสตร์", "ป.6"],
    ["สุจิตตรา ปางสุข", "พุธ", 7, "ลูกเสือ", "-"],
    ["สุจิตตรา ปางสุข", "พฤหัสบดี", 2, "ภาษาไทย", "ป.5"],
    ["สุจิตตรา ปางสุข", "พฤหัสบดี", 4, "วิทยาศาสตร์", "ป.5"],
    ["สุจิตตรา ปางสุข", "พฤหัสบดี", 7, "ภาษาไทยเพิ่มเติม", "ป.6"],
    ["สุจิตตรา ปางสุข", "พฤหัสบดี", 8, "PLC", "-"],
    ["สุจิตตรา ปางสุข", "ศุกร์", 2, "วิทยาศาสตร์", "ป.5"],
    ["สุจิตตรา ปางสุข", "ศุกร์", 3, "ภาษาไทย", "ป.6"],
    ["สุจิตตรา ปางสุข", "ศุกร์", 5, "ภาษาไทยเพิ่มเติม", "ป.5"],
    ["สุจิตตรา ปางสุข", "ศุกร์", 7, "แนะแนว", "-"],
    // 9. ครูสุโลธร ธนาวิชิตฤกษ์
    ["สุโลธร ธนาวิชิตฤกษ์", "จันทร์", 2, "วิทยาศาสตร์ฯ", "ป.3"],
    ["สุโลธร ธนาวิชิตฤกษ์", "จันทร์", 3, "คอมพิวเตอร์", "อ.1-2"],
    ["สุโลธร ธนาวิชิตฤกษ์", "จันทร์", 5, "วิทยาศาสตร์ฯ", "ป.1"],
    ["สุโลธร ธนาวิชิตฤกษ์", "จันทร์", 6, "เทคโนโลยีเพื่อเรียนรู้", "ป.5"],
    ["สุโลธร ธนาวิชิตฤกษ์", "จันทร์", 7, "วิทยาการคำนวณ", "ป.6"],
    ["สุโลธร ธนาวิชิตฤกษ์", "อังคาร", 2, "เทคโนโลยีเพื่อเรียนรู้", "ป.2"],
    ["สุโลธร ธนาวิชิตฤกษ์", "อังคาร", 3, "คอมพิวเตอร์", "อ.3"],
    ["สุโลธร ธนาวิชิตฤกษ์", "อังคาร", 5, "วิทยาศาสตร์ฯ", "ป.3"],
    ["สุโลธร ธนาวิชิตฤกษ์", "อังคาร", 6, "เทคโนโลยีเพื่อเรียนรู้", "ป.1"],
    ["สุโลธร ธนาวิชิตฤกษ์", "อังคาร", 7, "ชุมนุม", "-"],
    ["สุโลธร ธนาวิชิตฤกษ์", "พุธ", 5, "วิทยาศาสตร์ฯ", "ป.2"],
    ["สุโลธร ธนาวิชิตฤกษ์", "พุธ", 6, "เทคโนโลยีเพื่อเรียนรู้", "ป.6"],
    ["สุโลธร ธนาวิชิตฤกษ์", "พุธ", 7, "ลูกเสือ", "-"],
    ["สุโลธร ธนาวิชิตฤกษ์", "พฤหัสบดี", 3, "วิทยาศาสตร์ฯ", "ป.2"],
    ["สุโลธร ธนาวิชิตฤกษ์", "พฤหัสบดี", 6, "วิทยาการคำนวณ", "ป.5"],
    ["สุโลธร ธนาวิชิตฤกษ์", "พฤหัสบดี", 8, "PLC", "-"],
    ["สุโลธร ธนาวิชิตฤกษ์", "ศุกร์", 2, "เทคโนโลยีเพื่อเรียนรู้", "ป.3"],
    ["สุโลธร ธนาวิชิตฤกษ์", "ศุกร์", 3, "วิทยาศาสตร์ฯ", "ป.1"],
    ["สุโลธร ธนาวิชิตฤกษ์", "ศุกร์", 4, "วิทยาการคำนวณ", "ป.4"],
    ["สุโลธร ธนาวิชิตฤกษ์", "ศุกร์", 6, "เทคโนโลยีเพื่อเรียนรู้", "ป.4"],
    ["สุโลธร ธนาวิชิตฤกษ์", "ศุกร์", 7, "แนะแนว", "-"],
    // 10. ครูณัฏฐ์ ไตรพุทธมาศ
    ["ณัฏฐ์ ไตรพุทธมาศ", "จันทร์", 2, "คณิตศาสตร์", "ป.6"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "จันทร์", 3, "คณิตศาสตร์", "ป.1"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "จันทร์", 4, "คณิตศาสตร์", "ป.4"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "อังคาร", 2, "คณิตศาสตร์", "ป.5"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "อังคาร", 3, "คณิตศาสตร์", "ป.1"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "อังคาร", 5, "คณิตศาสตร์", "ป.6"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "อังคาร", 6, "คณิตศาสตร์", "ป.4"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "อังคาร", 7, "ชุมนุม", "-"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พุธ", 2, "คณิตศาสตร์", "ป.4"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พุธ", 3, "คณิตศาสตร์", "ป.5"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พุธ", 4, "คณิตศาสตร์", "ป.1"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พุธ", 7, "ลูกเสือ", "-"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พฤหัสบดี", 2, "คณิตศาสตร์", "ป.1"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พฤหัสบดี", 3, "คณิตศาสตร์", "ป.6"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พฤหัสบดี", 4, "คณิตศาสตร์", "ป.4"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พฤหัสบดี", 5, "คณิตศาสตร์", "ป.5"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "พฤหัสบดี", 8, "PLC", "-"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "ศุกร์", 2, "คณิตศาสตร์", "ป.1"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "ศุกร์", 3, "คณิตศาสตร์", "ป.5"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "ศุกร์", 4, "คณิตศาสตร์", "ป.6"],
    ["ณัฏฐ์ ไตรพุทธมาศ", "ศุกร์", 7, "แนะแนว", "-"],
  ];
  
  mSheet.getRange(1, 1, timetableData.length, 5).setValues(timetableData);

  // 3. จัดการชีตประวัติการสอนแทน (Substitution_Record)
  var sSheet = ss.getSheetByName("Substitution_Record") || ss.insertSheet("Substitution_Record");
  if (sSheet.getLastRow() === 0) {
    sSheet.getRange(1, 1, 1, 8).setValues([["วันที่", "คาบที่", "ครูที่ไม่อยู่", "ครูสอนแทน", "วิชา", "ห้องเรียน", "สาเหตุการลา", "วันเวลาบันทึก"]]);
  }

  // 4. จัดการชีตบันทึกการแลกคาบ (Swap_Record)
  var swapSheet = ss.getSheetByName("Swap_Record") || ss.insertSheet("Swap_Record");
  if (swapSheet.getLastRow() === 0) {
    swapSheet.getRange(1, 1, 1, 10).setValues([["ID", "วันที่ของคาบ", "คาบที่", "วิชา", "ระดับชั้น", "ครูผู้ขอแลก", "ครูที่แลกด้วย", "วันที่ชดเชย", "คาบที่ชดเชย", "วันเวลาบันทึก"]]);
  }

  // 5. จัดการชีตบันทึกการจองห้องปฏิบัติการ (Room_Booking_Record)
  var roomSheet = ss.getSheetByName("Room_Booking_Record") || ss.insertSheet("Room_Booking_Record");
  if (roomSheet.getLastRow() === 0) {
    roomSheet.getRange(1, 1, 1, 8).setValues([["ID", "วันที่จอง", "วัน", "คาบที่", "ครูผู้จอง", "ระดับชั้น", "วิชา/กิจกรรม", "วันเวลาบันทึก"]]);
  }

  Logger.log("อัปเดตฐานข้อมูลปีการศึกษา 2569 เทอม 2 ครบถ้วนทุกชีตสำเร็จเรียบร้อยแล้วครับ");
}
