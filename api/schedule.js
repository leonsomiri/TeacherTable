const { OAuth2Client } = require('google-auth-library');
const { google } = require('googleapis');

const SPREADSHEET_ID = '1Fl6N0krFeQ-tB1OGOVMKG73FumWsqXCfdOwILFsbF1Y';

const oauthClient = new OAuth2Client(process.env.GOOGLE_CLIENT_ID);

async function verifyEmail(idToken) {
  const ticket = await oauthClient.verifyIdToken({
    idToken,
    audience: process.env.GOOGLE_CLIENT_ID,
  });
  const payload = ticket.getPayload();
  if (!payload || !payload.email || !payload.email_verified) {
    throw new Error('Email not verified');
  }
  return payload.email.toLowerCase().trim();
}

function sheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: (process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY || '').replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  return google.sheets({ version: 'v4', auth });
}

function rowsToObjects(rows) {
  if (!rows || rows.length === 0) return [];
  const [header, ...data] = rows;
  return data.map(row => {
    const obj = {};
    header.forEach((h, i) => { obj[h] = row[i] ?? ''; });
    return obj;
  });
}

const WEEKDAYS = ['SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'];

function getTzOffsetMinutes(utcMillis, timeZone) {
  const dtf = new Intl.DateTimeFormat('en-US', {
    timeZone, hourCycle: 'h23',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  });
  const parts = dtf.formatToParts(new Date(utcMillis)).reduce((acc, p) => {
    acc[p.type] = p.value;
    return acc;
  }, {});
  const asUTC = Date.UTC(+parts.year, +parts.month - 1, +parts.day, +parts.hour, +parts.minute, +parts.second);
  return (asUTC - utcMillis) / 60000;
}

function zonedWallTimeToUtc(year, month, day, hour, minute, timeZone) {
  const guess = Date.UTC(year, month - 1, day, hour, minute);
  const offset = getTzOffsetMinutes(guess, timeZone);
  let utcMillis = guess - offset * 60000;
  const offset2 = getTzOffsetMinutes(utcMillis, timeZone);
  if (offset2 !== offset) utcMillis = guess - offset2 * 60000;
  return utcMillis;
}

function nextCalendarDateForWeekday(weekdayIndex) {
  const now = new Date();
  const diff = (weekdayIndex - now.getUTCDay() + 7) % 7;
  return { year: now.getUTCFullYear(), month: now.getUTCMonth() + 1, day: now.getUTCDate() + diff };
}

function formatTimeInZone(utcMillis, timeZone) {
  const dtf = new Intl.DateTimeFormat('en-US', {
    timeZone, weekday: 'long', hour: 'numeric', minute: '2-digit', hour12: true,
  });
  const parts = dtf.formatToParts(new Date(utcMillis)).reduce((acc, p) => {
    acc[p.type] = p.value;
    return acc;
  }, {});
  return { weekday: parts.weekday, time: `${parts.hour}:${parts.minute} ${parts.dayPeriod}` };
}

function convertFixedSlotToTeacherTime(dayName, startTimeStr, durationHours, studentTimeZone, teacherTimeZone) {
  const weekdayIndex = WEEKDAYS.indexOf((dayName || '').trim().toUpperCase());
  const timeMatch = (startTimeStr || '').match(/(\d{1,2}):(\d{2})/);
  if (weekdayIndex === -1 || !timeMatch || !studentTimeZone || !teacherTimeZone) return null;

  const { year, month, day } = nextCalendarDateForWeekday(weekdayIndex);
  const startUtc = zonedWallTimeToUtc(year, month, day, +timeMatch[1], +timeMatch[2], studentTimeZone);
  const endUtc = startUtc + durationHours * 3600000;

  const start = formatTimeInZone(startUtc, teacherTimeZone);
  const end = formatTimeInZone(endUtc, teacherTimeZone);
  const admin = formatTimeInZone(startUtc, 'Africa/Cairo');
  return {
    day: start.weekday, startTime: start.time, endTime: end.time,
    adminDay: admin.weekday, adminTime: admin.time,
  };
}

function lessonDurationHours(studentRow, subject) {
  if (!studentRow || !subject) return 1;
  const target = subject.trim().toLowerCase();
  for (let i = 1; i <= 5; i++) {
    const subj = (studentRow[`Subject${i}`] || '').trim().toLowerCase();
    if (subj && subj === target) {
      const dur = parseFloat(studentRow[`Lesson_Duration${i}`]);
      return Number.isFinite(dur) && dur > 0 ? dur : 1;
    }
  }
  return 1;
}

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');

  const authHeader = req.headers.authorization || '';
  const idToken = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
  if (!idToken) {
    res.status(401).json({ error: 'Missing token' });
    return;
  }

  let email;
  try {
    email = await verifyEmail(idToken);
  } catch (err) {
    res.status(401).json({ error: 'Invalid token' });
    return;
  }

  try {
    const sheets = sheetsClient();
    const [adminsRes, teachersRes, lessonsRes, availabilityRes, studentsRes, completedRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'ADMINS!A:C' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Teachers!A:L' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Lessons_Schedule!A:AE' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Student_Availability!A:I' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Students!A:AC' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'completed!A:AI' }),
    ]);

    const admins = rowsToObjects(adminsRes.data.values);
    const teachers = rowsToObjects(teachersRes.data.values);
    const lessons = rowsToObjects(lessonsRes.data.values);
    const availability = rowsToObjects(availabilityRes.data.values);
    const students = rowsToObjects(studentsRes.data.values);
    const completedLessons = rowsToObjects(completedRes.data.values);

    const admin = admins.find(a => (a.Admin_Eamil || '').toLowerCase().trim() === email);
    const isAdmin = Boolean(admin);

    let teacherId = null;
    let viewerName = admin ? admin.Admin_Name : null;
    if (!isAdmin) {
      const teacher = teachers.find(t => (t.Email || '').toLowerCase().trim() === email);
      if (!teacher) {
        res.status(403).json({ error: 'No schedule found for this Google account' });
        return;
      }
      teacherId = teacher.Teacher_ID;
      viewerName = teacher.Teacher_Name;
    }

    const studentsById = new Map(students.map(s => [s.Student_ID, s]));
    const teachersById = new Map(teachers.map(t => [t.Teacher_ID, t]));

    const mapLessonRow = l => {
      const studentRow = studentsById.get(l.Student_ID);
      return {
        Lesson_ID: l.Lesson_ID,
        Teacher_ID: l.Teacher_ID,
        Teacher_Name: l.Teacher_Name,
        Student_Name: l.Student_Name,
        Student_Country: (studentRow && studentRow.Country) || '',
        Student_Grade: (studentRow && studentRow.Stage) || '',
        Subject: l.Subject,
        Teacher_Date: l.Teacher_Date,
        Start_Time_Teacher: l.Start_Time_Teacher,
        End_Time_Teacher: l.End_Time_Teacher,
        Teacher_TimeZone: l.Teacher_TimeZone,
        CAIRO_TIME: l.CAIRO_TIME,
        MEET_LINK: l.MEET_LINK,
        Status: l.Status,
      };
    };

    const filtered = [...lessons, ...completedLessons]
      .filter(l => l.Lesson_ID)
      .filter(l => isAdmin || l.Teacher_ID === teacherId)
      .map(mapLessonRow);

    const fixedSchedule = availability
      .filter(a => a.Student_ID && a.Pref_Teacher)
      .filter(a => isAdmin || a.Pref_Teacher === teacherId)
      .filter(a => {
        const studentRow = studentsById.get(a.Student_ID);
        return studentRow && (studentRow.Status || '').trim().toLowerCase() === 'active';
      })
      .map(a => {
        const studentRow = studentsById.get(a.Student_ID);
        const teacherRow = teachersById.get(a.Pref_Teacher);
        const studentTimeZone = studentRow && studentRow.Timezone;
        const teacherTimeZone = teacherRow && teacherRow.Timezone;
        const duration = lessonDurationHours(studentRow, a.Subject);
        const converted = convertFixedSlotToTeacherTime(a.Day, a.Preferred_Hours, duration, studentTimeZone, teacherTimeZone);
        if (!converted) return null;
        return {
          Student_Name: (studentRow && studentRow.Student_Name) || a.Student_Name,
          Student_Country: (studentRow && studentRow.Country) || '',
          Student_Grade: (studentRow && studentRow.Stage) || '',
          Subject: a.Subject,
          Teacher_ID: a.Pref_Teacher,
          Teacher_Name: (teacherRow && teacherRow.Teacher_Name) || a['Teacher Name'] || '',
          Day_Teacher: converted.day,
          Start_Time_Teacher: converted.startTime,
          End_Time_Teacher: converted.endTime,
          Teacher_TimeZone: teacherTimeZone || '',
          Admin_Day: converted.adminDay,
          Admin_Time: converted.adminTime,
        };
      })
      .filter(Boolean);

    res.status(200).json({ isAdmin, viewerName, lessons: filtered, fixedSchedule });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error reading schedule' });
  }
};
