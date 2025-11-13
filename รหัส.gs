// --- ค่าคงที่และการตั้งค่าพื้นฐาน ---
const LOGIN_SHEET_ID = "18DUccMrxwubYVUeW9D8z16Jrd4NruUIodymoFd8dRdc";
const HOUSE_PHOTO_FOLDER_ID = "1HR3l9l5kAzYAjCOzE7PvnasNQYI-K6jO";
const STUDENT_PHOTO_FOLDER_ID = "17qoTEqubr-qTM43afWBH7rrZD682GNpX";
const TEACHER_ADDRESS = "172 หมู่ 7 ตำบล สันกำแพง อำเภอสันกำแพง เชียงใหม่ 50130";
const DISTANCE_THRESHOLD_KM = 0.5;

// =================================================================================
// --- ส่วนจัดการ Web App และ Session ---
// =================================================================================

function doGet(e) {
  if (e.parameter.page === 'teacher') {
    const token = e.parameter.token;
    const user = getUserFromToken(token);

    if (user) {
      const template = HtmlService.createTemplateFromFile('Teacher');
      template.user = user;
      template.token = token; // ส่ง token ไปให้หน้า Teacher.html
      return template.evaluate()
        .setTitle('ระบบจัดการข้อมูล')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } else {
      // ถ้าไม่มี token หรือ token ไม่ถูกต้อง ให้กลับไปหน้า Login
      return HtmlService.createTemplateFromFile('Login').evaluate()
        .setTitle('ลงชื่อเข้าใช้')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } else {
    const html = HtmlService.createTemplateFromFile('Index');
    html.url = ScriptApp.getService().getUrl();
    return html.evaluate()
      .setTitle('ระบบบันทึกข้อมูลเยี่ยมบ้าน')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

//สร้าง Session Token ที่ไม่ซ้ำกัน
function generateSessionToken() {
  return Utilities.getUuid();
}


//ดึงข้อมูลผู้ใช้จาก Token ใน Cache
function getUserFromToken(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const sessionData = cache.get(token);
  if (sessionData) {
    try {
      return JSON.parse(sessionData);
    } catch (e) {
      return null;
    }
  }
  return null;
}

//ตรวจสอบการ Login และสร้าง Token
function login(username, password) {
  try {
    const loginSpreadsheet = SpreadsheetApp.openById(LOGIN_SHEET_ID);
    const loginSheet = loginSpreadsheet.getSheetByName("Login");
    if (!loginSheet) return null;

    const data = loginSheet.getDataRange().getValues();
    const headers = data[0].map(h => h.trim());

    const usernameCol = headers.indexOf("Username");
    const passwordCol = headers.indexOf("password");
    const roleCol = headers.indexOf("role");
    const classroomCol = headers.indexOf("classroom");
    const fullnameCol = headers.indexOf("fullname");
    const picCol = headers.indexOf("pic");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[usernameCol] == username && row[passwordCol] == password) {
        const userSession = {
          username: row[usernameCol],
          fullname: row[fullnameCol],
          role: row[roleCol],
          classroom: row[classroomCol],
          pic: row[picCol],
        };

        const token = generateSessionToken();
        const cache = CacheService.getScriptCache();
        // เก็บข้อมูลใน cache เป็นเวลา 30 นาที (1800 วินาที)
        cache.put(token, JSON.stringify(userSession), 1800);

        return token; // ส่ง token กลับไป
      }
    }
    return null;
  } catch (e) {
    Logger.log("Login Function Error: " + e.toString());
    return null;
  }
}

//ออกจากระบบ (ลบ Token ออกจาก Cache)
function logout(token) {
  if (token) {
    const cache = CacheService.getScriptCache();
    cache.remove(token);
  }
  return "Logged out successfully";
}

//ใช้สำหรับ include ไฟล์อื่นๆ เข้าไปใน HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =================================================================================
// --- ฟังก์ชันจัดการข้อมูลนักเรียน (CRUD) ---
// =================================================================================

//ดึงข้อมูลแดชบอร์ดตามสิทธิ์ของผู้ใช้
function getTeacherDashboardData(token) {
  const session = getUserFromToken(token);
  if (!session) {
    throw new Error("เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่");
  }

  let studentObjects = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // สำหรับ Admin: เห็นข้อมูลทุกห้อง
  if (session.role === 'admin') {
    const allSheets = ss.getSheets();
    for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      if (sheetName.startsWith("ข้อมูลนักเรียน")) {
        const records = getFullRecordsForSheet(sheetName);
        if (records && records.length > 1) {
          const headers = records[0];
          const classroomName = sheetName;

          const newObjects = records.slice(1).map(row => {
            let obj = {};
            headers.forEach((header, i) => { obj[header] = row[i]; });
            obj.classroom = classroomName;
            return obj;
          });
          studentObjects.push(...newObjects);
        }
      }
    }
  }
  // สำหรับครู (User): เห็นเฉพาะห้องของตัวเอง
  else if (session.role === 'User') {
    const sheetName = session.classroom;
    if (sheetName) {
      const records = getFullRecordsForSheet(sheetName);
      if (records && records.length > 1) {
        const headers = records[0];
        const newObjects = records.slice(1).map(row => {
          let obj = {};
          headers.forEach((header, i) => { obj[header] = row[i]; });
          obj.classroom = sheetName;
          return obj;
        });
        studentObjects.push(...newObjects);
      }
    }
  }

  if (studentObjects.length === 0) {
    return { allStudents: [], bySubdistrict: {} };
  }

  const allStudentsSorted = [...studentObjects].sort((a, b) => {
    const distA = parseFloat(a['ระยะทาง(กม.)(คำนวณ)']);
    const distB = parseFloat(b['ระยะทาง(กม.)(คำนวณ)']);
    return distA - distB;
  });

  const studentsBySubdistrict = studentObjects.reduce((acc, student) => {
    const subdistrict = student['ตำบล/แขวง'] || 'ไม่ระบุตำบล';
    if (!acc[subdistrict]) acc[subdistrict] = [];
    acc[subdistrict].push(student);
    return acc;
  }, {});

  for (const subdistrict in studentsBySubdistrict) {
    studentsBySubdistrict[subdistrict].sort((a, b) => {
      const distA = parseFloat(a['ระยะทาง(กม.)(คำนวณ)']);
      const distB = parseFloat(b['ระยะทาง(กม.)(คำนวณ)']);
      return distA - distB;
    });
  }

  return { allStudents: allStudentsSorted, bySubdistrict: studentsBySubdistrict };
}

//ฟังก์ชันสำหรับอัปโหลดรูปภาพหน้าบ้าน
function uploadFileToGoogleDrive(fileObject) {
  try {
    const folder = DriveApp.getFolderById(HOUSE_PHOTO_FOLDER_ID);
    const decoded = Utilities.base64Decode(fileObject.base64);
    const blob = Utilities.newBlob(decoded, fileObject.mimeType, fileObject.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileID = file.getId();
    // REVERTED: Changed back to lh3.googleusercontent.com URL for direct image embedding.
    const fileURL = "https://lh3.googleusercontent.com/d/" + fileID;
    return fileURL;
  } catch (e) {
    return 'Error: ' + e.toString();
  }
}

//ฟังก์ชันใหม่สำหรับอัปโหลดรูปโปรไฟล์นักเรียน
function uploadProfilePhoto(fileObject) {
  try {
    const folder = DriveApp.getFolderById(STUDENT_PHOTO_FOLDER_ID);
    const decoded = Utilities.base64Decode(fileObject.base64);
    const blob = Utilities.newBlob(decoded, fileObject.mimeType, fileObject.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileID = file.getId();
    // REVERTED: Changed back to lh3.googleusercontent.com URL for direct image embedding.
    const fileURL = "https://lh3.googleusercontent.com/d/" + fileID;
    return fileURL;
  } catch (e) {
    return 'Error: ' + e.toString();
  }
}

function addStudentRecord(data) {
  try {
    const id = "ID" + new Date().getTime();
    const classSheetName = "ข้อมูลนักเรียน" + data.classroom;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(classSheetName);

    if (!sheet) {
      sheet = ss.insertSheet(classSheetName);
      const headers = [
        'ID', 'คำนำหน้า', 'ชื่อ-นามสกุล', 'เลขที่', 'เลขประจำตัวนักเรียน', 'เบอร์โทรศัพท์',
        'บ้านเลขที่', 'หมู่บ้าน,อาคาร', 'หมู่ที่', 'ซอย', 'ตำบล/แขวง', 'อำเภอ,เขต',
        'จังหวัด', 'รหัสไปรษณีย์', 'Latitude (ที่ปักหมุด)', 'Longitude (ที่ปักหมุด)', 'จุดสังเกตใกล้บ้าน',
        'URL รูปถ่ายนักเรียน', 'URL รูปถ่ายหน้าบ้าน', 'คำนำหน้าผู้ปกครอง', 'ชื่อ-สกุล ผู้ปกครอง', 'ความเกี่ยวข้อง',
        'เบอร์โทรศัพท์ผู้ปกครอง', 'ที่อยู่เต็ม(คำนวณ)', 'ระยะทาง(กม.)(คำนวณ)', 'สถานะการตรวจสอบ'
      ];
      sheet.appendRow(headers);
    }

    const nextRow = sheet.getLastRow() + 1;
    const fullAddressForDisplay = `บ้านเลขที่ : ${data.houseNumber} ${data.village || ''} ${data.moo ? 'หมู่ ' + data.moo : ''} ${data.alley ? 'ซอย ' + data.alley : ''} ต.${data.subdistrict} อ.${data.district} จ.${data.province} ${data.postalCode}`.replace(/\s+/g, ' ').trim();

    // MODIFIED: Added single quote prefix (') to force text format in Google Sheets for specified fields.
    const recordToAppend = [
      String(id),
      data.title_student,
      data.fullname,
      "'" + data.number,
      "'" + data.studentId,
      "'" + data.phone,
      "'" + data.houseNumber,
      "'" + data.village,
      "'" + data.moo,
      "'" + data.alley,
      data.subdistrict,
      data.district,
      data.province,
      "'" + data.postalCode,
      "'" + data.lat,
      "'" + data.lng,
      data.landmark,
      data.student_photo_url,
      data.house_photo_url,
      data.title_parent,
      data.fullname_parent,
      data.relevance,
      "'" + data.phone_parent,
      `=MANAGE_DATA(G${nextRow}:N${nextRow},O${nextRow},P${nextRow},true,$AA$2)`,
      '',
      'รอตรวจสอบ'
    ];
    sheet.appendRow(recordToAppend);

    const recordToReturn = [...recordToAppend];
    recordToReturn[23] = fullAddressForDisplay;

    const allHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    return { headers: allHeaders, record: recordToReturn };

  } catch (e) {
    return { error: e.message };
  }
}

function updateTeacherAddress(token, newAddress) {
  const session = getUserFromToken(token);
  if (!session) throw new Error("เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่");

  const ss = SpreadsheetApp.openById(LOGIN_SHEET_ID);
  const sheet = ss.getSheetByName("Login");
  if (!sheet) throw new Error("ไม่พบชีต 'Login'");

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const fullnameColIndex = headers.indexOf("fullname");
  const addressColIndex = headers.indexOf("address");

  if (fullnameColIndex === -1 || addressColIndex === -1) throw new Error("ไม่พบคอลัมน์ 'fullname' หรือ 'address'");

  const teacherId = session.fullname;
  for (let i = 1; i < data.length; i++) {
    if (data[i][fullnameColIndex] === teacherId) {
      sheet.getRange(i + 1, addressColIndex + 1).setValue(newAddress);
      return "อัปเดตที่อยู่เรียบร้อย";
    }
  }
  throw new Error("ไม่พบข้อมูลครู " + teacherId);
}

function deleteStudentRecord(token, id, sheetName) {
  const session = getUserFromToken(token);
  if (!session) throw new Error("เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return "ไม่พบชีทข้อมูล";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return "ลบข้อมูลเรียบร้อย";
    }
  }
  return "ไม่พบข้อมูล";
}

function verifyStudentData(token, id, status, sheetName) {
  const session = getUserFromToken(token);
  if (!session) throw new Error("เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return "ไม่พบชีทข้อมูล";
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowIndex = data.findIndex(row => row[0] === id);
  const statusColIndex = headers.indexOf('สถานะการตรวจสอบ');

  if (rowIndex !== -1 && statusColIndex !== -1) {
    sheet.getRange(rowIndex + 1, statusColIndex + 1).setValue(status);
    return "อัปเดตสถานะแล้ว";
  }
  return "ไม่พบข้อมูล หรือ ไม่พบคอลัมน์สถานะ";
}

// =================================================================================
// --- ฟังก์ชันคำนวณและอื่นๆ ---
// =================================================================================

/**
 * [ปรับปรุงใหม่] คำนวณเส้นทางที่สั้นที่สุด (Optimized Route) สำหรับการเยี่ยมบ้านนักเรียน
 * @param {string} token - Token สำหรับการยืนยันตัวตน
 * @param {string} origin - พิกัดจุดเริ่มต้น "lat,lng"
 * @param {Array<object>} studentData - อาร์เรย์ของอ็อบเจกต์นักเรียน [{id, name, coords:"lat,lng"}, ...]
 * @returns {object} ผลลัพธ์การคำนวณเส้นทาง
 */
function calculateRoute(token, origin, studentData) {
    const session = getUserFromToken(token);
    if (!session) {
        return { error: "เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่" };
    }

    if (!origin || !studentData || studentData.length === 0) {
        return { error: "กรุณาระบุจุดเริ่มต้นและเลือกนักเรียนอย่างน้อย 1 คน" };
    }
    
    const waypoints = studentData.map(s => s.coords);

    try {
        const directionsFinder = Maps.newDirectionFinder()
            .setOrigin(origin)
            .setDestination(origin) 
            .setMode(Maps.DirectionFinder.Mode.DRIVING);

        waypoints.forEach(wp => directionsFinder.addWaypoint(wp));

        directionsFinder.setOptimizeWaypoints(true);

        const directions = directionsFinder.getDirections();

        if (!directions || !directions.routes || directions.routes.length === 0) {
            return { error: "ไม่สามารถคำนวณเส้นทางได้ กรุณาตรวจสอบพิกัดของนักเรียน" };
        }

        const route = directions.routes[0];
        const waypointOrder = route.waypoint_order;

        const orderedStudents = waypointOrder.map(index => studentData[index]);

        let tripDistanceMeters = 0;
        let tripDurationSeconds = 0;

        const legsToVisitCount = route.legs.length - 1;

        for (let i = 0; i < legsToVisitCount; i++) {
            tripDistanceMeters += route.legs[i].distance.value;
            tripDurationSeconds += route.legs[i].duration.value;
        }

        if (waypoints.length === 1 && route.legs.length > 0) {
            tripDistanceMeters = route.legs[0].distance.value;
            tripDurationSeconds = route.legs[0].duration.value;
        }

        const hours = Math.floor(tripDurationSeconds / 3600);
        const minutes = Math.floor((tripDurationSeconds % 3600) / 60);
        let durationText = "";
        if (hours > 0) durationText += `${hours} ชั่วโมง `;
        if (minutes > 0) durationText += `${minutes} นาที`;
        if (durationText.trim() === "") durationText = `${Math.round(tripDurationSeconds)} วินาที`;
        
        return {
            success: true,
            orderedStudents: orderedStudents,
            totalDistanceKm: (tripDistanceMeters / 1000).toFixed(2),
            totalDurationText: durationText.trim()
        };

    } catch (e) {
        Logger.log(`Error in calculateRoute: ${e.toString()}`);
        return { error: e.message };
    }
}


function getFullRecordsForSheet(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 1) {
      return [];
    }
    return sheet.getDataRange().getValues();
  } catch (e) {
    Logger.log(`Could not get full records for sheet: ${sheetName}. Error: ${e.message}`);
    return [];
  }
}

function setTeacherAddressToSheet(token, address, sheetName) {
  const session = getUserFromToken(token);
  if (!session) throw new Error("เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่");

  if (sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if(sheet) {
      sheet.getRange("AA2").setValue(address);
    }
  }
}

// =================================================================================
// --- CUSTOM FUNCTION และฟังก์ชันเสริม ---
// =================================================================================

/**
 * @customfunction
 */
function MANAGE_DATA(addressRange, pinnedLat, pinnedLng, forceRecalculate = false, teacher_address_0) {
  // FIX: Check if the teacher's address (from cell AA2) is provided to prevent errors.
  if (!teacher_address_0 || String(teacher_address_0).trim() === '') {
    return [["โปรดระบุที่อยู่ครูในหน้าแดชบอร์ด", ""]];
  }

  if (!addressRange || !pinnedLat || !pinnedLng) {
    return [["กรุณาระบุข้อมูลให้ครบ", ""]];
  }

  const fullAddress = formatFullAddress(addressRange[0]);
  const cacheKey = `manage_data_${fullAddress}_${pinnedLat}_${pinnedLng}_${teacher_address_0}`;
  if (!forceRecalculate) {
    const cachedValue = getCache(cacheKey);
    if (cachedValue) {
      return [JSON.parse(cachedValue)];
    }
  }

  try {
    const geocodedCoords = geocodeAddress(fullAddress);
    if (!geocodedCoords) {
      const distanceFromPin = getDirectionsDistance(`${pinnedLat},${pinnedLng}`, teacher_address_0);
      const result = [fullAddress, distanceFromPin];
      setCache(cacheKey, JSON.stringify(result));
      return [result];
    }

    const distanceDiff = haversineDistance(
      pinnedLat, pinnedLng,
      geocodedCoords.lat, geocodedCoords.lng
    );

    let finalCoords;
    if (distanceDiff > DISTANCE_THRESHOLD_KM) {
      finalCoords = `${pinnedLat},${pinnedLng}`;
    } else {
      finalCoords = `${geocodedCoords.lat},${geocodedCoords.lng}`;
    }

    const finalDistance = getDirectionsDistance(finalCoords, teacher_address_0);
    const result = [fullAddress, finalDistance];
    setCache(cacheKey, JSON.stringify(result));
    return [result];

  } catch (e) {
    Logger.log(`MANAGE_DATA Error for address "${fullAddress}": ${e.toString()}`);
    return [[`เกิดข้อผิดพลาด: ${e.message}`, ""]];
  }
}

function formatFullAddress(addressParts) {
  const components = [`บ้านเลขที่ ${addressParts[0]}`, `${addressParts[1]}`, `หมู่ ${addressParts[2]}`, `ซอย ${addressParts[3]}`, `ต.${addressParts[4]}`, `อ.${addressParts[5]}`, `จ.${addressParts[6]}`, `${addressParts[7]}`];
  const validParts = addressParts.map(p => (p || "").toString().trim()).map((p, i) => (p && p !== "-") ? components[i] : null).filter(Boolean);
  return validParts.join(" ");
}

function haversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371; const dLat = (lat2 - lat1) * Math.PI / 180; const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) * Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function geocodeAddress(address) {
  // FIX: Validate the address before making an API call to prevent errors.
  if (!address || typeof address !== 'string' || address.trim() === '') {
    return null;
  }
  
  const key = `geocode_th_${address}`; const cached = getCache(key); if (cached) return JSON.parse(cached);
  const geocoder = Maps.newGeocoder().setLanguage('th').setRegion('th'); const response = geocoder.geocode(address);
  if (!response || !response.results || response.results.length === 0) { return null; }
  const { geometry: { location: { lat, lng } } } = response.results[0]; const result = { lat, lng };
  setCache(key, JSON.stringify(result)); return result;
}

function getDirectionsDistance(origin, destination) {
  // FIX: Validate both origin and destination before making an API call to prevent errors.
  if (!origin || !destination || String(origin).trim() === '' || String(destination).trim() === '') {
    return 'ข้อมูลพิกัดไม่ครบถ้วน';
  }

  const key = `distance_${origin}_${destination}`; const cached = getCache(key); if (cached) return cached;
  try {
    const { routes: [data] = [] } = Maps.newDirectionFinder().setOrigin(origin).setDestination(destination).setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();
    if (!data) return 'ไม่พบเส้นทาง';
    const { legs: [{ distance: { value: distanceValue } } = {}] = [] } = data;
    const distanceInKm = distanceValue / 1000; const formattedDistance = distanceInKm.toFixed(2);
    setCache(key, formattedDistance); return formattedDistance;
  } catch(e) { 
    Logger.log(`Directions API Error: ${e.message} | Origin: ${origin} | Destination: ${destination}`);
    return 'คำนวณไม่ได้'; 
  }
}

// --- Utility Functions for Caching ---
const md5 = (key = "") => Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key.toLowerCase().replace(/\s/g, "")).map((char) => (char + 256).toString(16).slice(-2)).join("");
const getCache = (key) => CacheService.getDocumentCache().get(md5(key));
const setCache = (key, value) => { CacheService.getDocumentCache().put(md5(key), value, 21600); };

