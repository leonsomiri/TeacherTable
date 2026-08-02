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
    const [adminsRes, teachersRes, lessonsRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'ADMINS!A:C' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Teachers!A:L' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Lessons_Schedule!A:AE' }),
    ]);

    const admins = rowsToObjects(adminsRes.data.values);
    const teachers = rowsToObjects(teachersRes.data.values);
    const lessons = rowsToObjects(lessonsRes.data.values);

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

    const filtered = lessons
      .filter(l => l.Lesson_ID)
      .filter(l => isAdmin || l.Teacher_ID === teacherId)
      .map(l => ({
        Lesson_ID: l.Lesson_ID,
        Teacher_ID: l.Teacher_ID,
        Teacher_Name: l.Teacher_Name,
        Student_Name: l.Student_Name,
        Subject: l.Subject,
        Teacher_Date: l.Teacher_Date,
        Start_Time_Teacher: l.Start_Time_Teacher,
        End_Time_Teacher: l.End_Time_Teacher,
        Teacher_TimeZone: l.Teacher_TimeZone,
        CAIRO_TIME: l.CAIRO_TIME,
        MEET_LINK: l.MEET_LINK,
        Status: l.Status,
      }));

    res.status(200).json({ isAdmin, viewerName, lessons: filtered });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Server error reading schedule' });
  }
};
