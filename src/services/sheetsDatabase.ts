import {
  ClassItem,
  DatabasePayload,
  DEFAULT_CLASSES,
  DEFAULT_SCHOOL,
  DEFAULT_STUDENT_GRADES,
  DEFAULT_STUDENTS,
  DEFAULT_SUBJECT_TEMPLATES,
  DEFAULT_USERS,
  RoleType,
  SchoolData,
  SPREADSHEET_ID,
  StudentGradesMap,
  StudentItem,
  UserAccount,
} from '../types';

export interface SheetMetadataInfo {
  spreadsheetTitle: string;
  gid0Title: string;
  tabs: string[];
  gid0PreviewRows: string[][];
  hasStructuredTabs: boolean;
  lastSyncedAt: string;
}

const REQUIRED_TABS = ['Sekolah', 'Kelas', 'Siswa', 'Users', 'Nilai'];

function getAllSubjectNames(): string[] {
  const set = new Set<string>();
  Object.values(DEFAULT_SUBJECT_TEMPLATES).forEach((tpl) => {
    tpl.umum.forEach((s) => set.add(s.name));
    tpl.kejuruan.forEach((s) => set.add(s.name));
  });
  return Array.from(set);
}

export async function fetchFromGoogleSheets(accessToken: string): Promise<{
  payload: DatabasePayload;
  meta: SheetMetadataInfo;
  isEmptySheet: boolean;
}> {
  // 1. Fetch spreadsheet metadata
  const metaRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!metaRes.ok) {
    const errBody = await metaRes.json().catch(() => ({}));
    const msg = errBody?.error?.message || `HTTP ${metaRes.status}`;
    throw new Error(`Gagal mengakses Google Sheet: ${msg}`);
  }

  const metaJson = await metaRes.json();
  const spreadsheetTitle: string = metaJson.properties?.title || 'Database E-Raport STS';
  const sheetsList: Array<{ properties: { title: string; sheetId: number } }> =
    metaJson.sheets || [];

  const tabs = sheetsList.map((s) => s.properties.title);
  const gid0Sheet =
    sheetsList.find((s) => s.properties.sheetId === 0) || sheetsList[0];
  const gid0Title = gid0Sheet ? gid0Sheet.properties.title : 'Sheet1';

  const hasStructuredTabs = REQUIRED_TABS.every((t) => tabs.includes(t));

  // Build ranges to batchGet
  const rangesToFetch: string[] = [`'${gid0Title}'!A1:Z200`];
  REQUIRED_TABS.forEach((tab) => {
    if (tabs.includes(tab)) {
      rangesToFetch.push(`'${tab}'!A1:Z500`);
    }
  });

  const params = rangesToFetch
    .map((r) => `ranges=${encodeURIComponent(r)}`)
    .join('&');
  const valRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values:batchGet?${params}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!valRes.ok) {
    const errBody = await valRes.json().catch(() => ({}));
    throw new Error(
      errBody?.error?.message || 'Gagal membaca isi tabel dari Google Sheet.'
    );
  }

  const valJson = await valRes.json();
  const valueRanges: Array<{ range: string; values?: string[][] }> =
    valJson.valueRanges || [];

  const gid0Rows: string[][] = valueRanges[0]?.values || [];

  // Map tab name -> values
  const tabValuesMap: Record<string, string[][]> = {};
  let rangeIdx = 1;
  REQUIRED_TABS.forEach((tab) => {
    if (tabs.includes(tab)) {
      tabValuesMap[tab] = valueRanges[rangeIdx]?.values || [];
      rangeIdx++;
    }
  });

  // Check if the spreadsheet is completely empty
  const hasAnyStructuredData = Object.values(tabValuesMap).some(
    (rows) => rows && rows.length > 1
  );
  const hasGid0Data = gid0Rows.length > 1;

  if (!hasAnyStructuredData && !hasGid0Data) {
    return {
      payload: {
        school: { ...DEFAULT_SCHOOL },
        classes: [...DEFAULT_CLASSES],
        students: JSON.parse(JSON.stringify(DEFAULT_STUDENTS)),
        users: JSON.parse(JSON.stringify(DEFAULT_USERS)),
        studentGrades: JSON.parse(JSON.stringify(DEFAULT_STUDENT_GRADES)),
      },
      meta: {
        spreadsheetTitle,
        gid0Title,
        tabs,
        gid0PreviewRows: gid0Rows.slice(0, 15),
        hasStructuredTabs,
        lastSyncedAt: new Date().toLocaleTimeString('id-ID'),
      },
      isEmptySheet: true,
    };
  }

  // Parse School
  let school: SchoolData = { ...DEFAULT_SCHOOL };
  const sekolahRows = tabValuesMap['Sekolah'] || [];
  if (sekolahRows.length > 1) {
    const kv: Record<string, string> = {};
    sekolahRows.slice(1).forEach((r) => {
      if (r[0]) kv[r[0].trim()] = r[1] ?? '';
    });
    school = {
      name: kv['name'] || DEFAULT_SCHOOL.name,
      npsn: kv['npsn'] || DEFAULT_SCHOOL.npsn,
      logoUrl: kv['logoUrl'] || DEFAULT_SCHOOL.logoUrl,
      kopUrl: kv['kopUrl'] || DEFAULT_SCHOOL.kopUrl,
      principal: kv['principal'] || DEFAULT_SCHOOL.principal,
      principalNip: kv['principalNip'] || DEFAULT_SCHOOL.principalNip,
      tp: kv['tp'] || DEFAULT_SCHOOL.tp,
      datePrint: kv['datePrint'] || DEFAULT_SCHOOL.datePrint,
    };
  }

  // Parse Classes
  let classes: ClassItem[] = [...DEFAULT_CLASSES];
  const kelasRows = tabValuesMap['Kelas'] || [];
  if (kelasRows.length > 1) {
    classes = kelasRows
      .slice(1)
      .filter((r) => r.length >= 2 && r[1])
      .map((r, idx) => ({
        id: Number(r[0]) || idx + 1,
        name: r[1] || '',
        grade: r[2] || '10',
        major: r[3] || 'Teknik Kendaraan Ringan',
        waliKelasName: r[4] || '-',
        waliKelasNip: r[5] || '-',
      }));
  }

  // Parse Users
  let users: UserAccount[] = JSON.parse(JSON.stringify(DEFAULT_USERS));
  const userRows = tabValuesMap['Users'] || [];
  if (userRows.length > 1) {
    users = userRows
      .slice(1)
      .filter((r) => r.length >= 3 && r[1])
      .map((r, idx) => {
        const rawRoles = (r[5] || 'gurumapel')
          .split(',')
          .map((x) => x.trim().toLowerCase())
          .filter(
            (x): x is RoleType =>
              x === 'admin' || x === 'walikelas' || x === 'gurumapel'
          );
        return {
          id: Number(r[0]) || idx + 1,
          username: r[1] || '',
          password: r[2] || '123',
          name: r[3] || r[1] || '',
          nip: r[4] || '',
          roles: rawRoles.length > 0 ? rawRoles : ['gurumapel'],
          waliKelasFor: r[6] || '',
        };
      });
  }

  // Parse Students & Grades
  let students: StudentItem[] = [];
  const studentGrades: StudentGradesMap = {};

  const siswaRows = tabValuesMap['Siswa'] || [];
  const nilaiRows = tabValuesMap['Nilai'] || [];

  if (siswaRows.length > 1) {
    students = siswaRows
      .slice(1)
      .filter((r) => r.length >= 4 && r[3])
      .map((r, idx) => {
        const id = Number(r[0]) || 100 + idx + 1;
        studentGrades[id] = {};
        return {
          id,
          nis: r[1] || '',
          nisn: r[2] || '',
          name: r[3] || '',
          className: r[4] || '10 TKR 1',
          grade: r[5] || '10',
          major: r[6] || 'Teknik Kendaraan Ringan',
          attendance: {
            sakit: Number(r[7]) || 0,
            izin: Number(r[8]) || 0,
            alpa: Number(r[9]) || 0,
          },
          note: r[10] || 'Tingkatkan terus motivasi dan kedisiplinan belajar.',
        };
      });

    if (nilaiRows.length > 1) {
      nilaiRows.slice(1).forEach((r) => {
        const stId = Number(r[0]);
        const subjectName = r[3];
        const scoreVal = r[4];
        if (stId && subjectName && scoreVal !== undefined && scoreVal !== '') {
          if (!studentGrades[stId]) studentGrades[stId] = {};
          const num = Number(scoreVal);
          if (!Number.isNaN(num)) {
            studentGrades[stId][subjectName] = num;
          }
        }
      });
    }
  } else if (hasGid0Data) {
    // Parse directly from gid=0 if Siswa tab isn't populated yet
    const headers = gid0Rows[0].map((h) => (h || '').trim());
    const lowerHeaders = headers.map((h) => h.toLowerCase());

    const idxId = lowerHeaders.findIndex((h) => h === 'id' || h === 'no');
    const idxNis = lowerHeaders.findIndex((h) => h === 'nis');
    const idxNisn = lowerHeaders.findIndex((h) => h === 'nisn');
    const idxName = lowerHeaders.findIndex(
      (h) => h.includes('nama') || h.includes('name') || h.includes('siswa')
    );
    const idxClass = lowerHeaders.findIndex(
      (h) => h.includes('kelas') || h.includes('rombel') || h === 'classname'
    );
    const idxGrade = lowerHeaders.findIndex(
      (h) => h.includes('tingkat') || h === 'grade'
    );
    const idxMajor = lowerHeaders.findIndex(
      (h) => h.includes('jurusan') || h.includes('keahlian') || h === 'major'
    );
    const idxSakit = lowerHeaders.findIndex((h) => h === 'sakit' || h === 's');
    const idxIzin = lowerHeaders.findIndex(
      (h) => h === 'izin' || h === 'ijin' || h === 'i'
    );
    const idxAlpa = lowerHeaders.findIndex(
      (h) => h === 'alpa' || h.includes('tanpa') || h === 'a'
    );
    const idxNote = lowerHeaders.findIndex(
      (h) => h.includes('catatan') || h.includes('note')
    );

    const allSubjects = getAllSubjectNames();

    if (idxName !== -1) {
      students = gid0Rows
        .slice(1)
        .filter((r) => r[idxName] && r[idxName].trim() !== '')
        .map((r, i) => {
          const stId =
            (idxId !== -1 ? Number(r[idxId]) : 0) || 100 + i + 1;
          const clsName =
            (idxClass !== -1 ? r[idxClass] : '') || '10 TKR 1';
          const inferredGrade =
            (idxGrade !== -1 ? r[idxGrade] : '') ||
            (clsName.startsWith('12')
              ? '12'
              : clsName.startsWith('11')
              ? '11'
              : '10');
          const major =
            (idxMajor !== -1 ? r[idxMajor] : '') ||
            'Teknik Kendaraan Ringan';

          studentGrades[stId] = {};
          // Extract any subject columns present in gid=0
          headers.forEach((hName, colIndex) => {
            const matchedSub = allSubjects.find(
              (sub) => sub.toLowerCase() === hName.toLowerCase()
            );
            if (matchedSub && r[colIndex] !== undefined && r[colIndex] !== '') {
              const val = Number(r[colIndex]);
              if (!Number.isNaN(val)) {
                studentGrades[stId][matchedSub] = val;
              }
            }
          });

          return {
            id: stId,
            nis: (idxNis !== -1 ? r[idxNis] : '') || `2026100${i + 1}`,
            nisn: (idxNisn !== -1 ? r[idxNisn] : '') || `008123456${i + 1}`,
            name: r[idxName].trim().toUpperCase(),
            className: clsName,
            grade: inferredGrade,
            major,
            attendance: {
              sakit: idxSakit !== -1 ? Number(r[idxSakit]) || 0 : 0,
              izin: idxIzin !== -1 ? Number(r[idxIzin]) || 0 : 0,
              alpa: idxAlpa !== -1 ? Number(r[idxAlpa]) || 0 : 0,
            },
            note:
              (idxNote !== -1 ? r[idxNote] : '') ||
              'Tingkatkan terus motivasi dan kedisiplinan belajar.',
          };
        });

      // Ensure classes from gid=0 exist in classes list
      const existingClassNames = new Set(classes.map((c) => c.name));
      students.forEach((st) => {
        if (st.className && !existingClassNames.has(st.className)) {
          existingClassNames.add(st.className);
          classes.push({
            id: Date.now() + Math.floor(Math.random() * 1000),
            name: st.className,
            grade: st.grade,
            major: st.major,
            waliKelasName: '-',
            waliKelasNip: '-',
          });
        }
      });
    } else {
      // Fallback to default students if gid=0 has unstructured rows
      students = JSON.parse(JSON.stringify(DEFAULT_STUDENTS));
      Object.assign(
        studentGrades,
        JSON.parse(JSON.stringify(DEFAULT_STUDENT_GRADES))
      );
    }
  }

  // Cross-link Wali Kelas from Users into Classes if missing
  classes = classes.map((cls) => {
    const isMissingWali =
      !cls.waliKelasName ||
      cls.waliKelasName.trim() === '-' ||
      cls.waliKelasName.trim() === 'Nama Wali Kelas';
    const isMissingNip = !cls.waliKelasNip || cls.waliKelasNip.trim() === '-';

    const matchedWaliUser = users.find(
      (u) =>
        u.roles.includes('walikelas') &&
        u.waliKelasFor &&
        u.waliKelasFor.trim().toLowerCase() === cls.name.trim().toLowerCase()
    );

    if (matchedWaliUser) {
      return {
        ...cls,
        waliKelasName: isMissingWali ? matchedWaliUser.name : cls.waliKelasName,
        waliKelasNip: isMissingNip
          ? matchedWaliUser.nip || '-'
          : cls.waliKelasNip,
      };
    }
    return cls;
  });

  return {
    payload: {
      school,
      classes,
      students,
      users,
      studentGrades,
    },
    meta: {
      spreadsheetTitle,
      gid0Title,
      tabs,
      gid0PreviewRows: gid0Rows.slice(0, 15),
      hasStructuredTabs,
      lastSyncedAt: new Date().toLocaleTimeString('id-ID'),
    },
    isEmptySheet: false,
  };
}

export async function syncAllToGoogleSheets(
  accessToken: string,
  payload: DatabasePayload
): Promise<SheetMetadataInfo> {
  // 1. Fetch current metadata to check existing sheet tabs
  const metaRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!metaRes.ok) {
    const errBody = await metaRes.json().catch(() => ({}));
    throw new Error(
      errBody?.error?.message || 'Gagal membaca metadata Google Sheet.'
    );
  }

  const metaJson = await metaRes.json();
  const spreadsheetTitle: string =
    metaJson.properties?.title || 'Database E-Raport STS';
  const sheetsList: Array<{ properties: { title: string; sheetId: number } }> =
    metaJson.sheets || [];
  const existingTabs = sheetsList.map((s) => s.properties.title);
  const gid0Sheet =
    sheetsList.find((s) => s.properties.sheetId === 0) || sheetsList[0];
  const gid0Title = gid0Sheet ? gid0Sheet.properties.title : 'Sheet1';

  // 2. Create any missing required tabs
  const missingTabs = REQUIRED_TABS.filter((t) => !existingTabs.includes(t));
  if (missingTabs.length > 0) {
    const requests = missingTabs.map((title) => ({
      addSheet: { properties: { title } },
    }));
    const addTabRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}:batchUpdate`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ requests }),
      }
    );
    if (!addTabRes.ok) {
      const errBody = await addTabRes.json().catch(() => ({}));
      throw new Error(
        errBody?.error?.message || 'Gagal membuat tab baru di Google Sheet.'
      );
    }
  }

  // 3. Clear old values in target sheets before writing fresh rows
  const rangesToClear = [
    `'${gid0Title}'!A1:ZZ1000`,
    ...REQUIRED_TABS.map((t) => `'${t}'!A1:Z1000`),
  ];
  await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values:batchClear`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ ranges: rangesToClear }),
    }
  );

  // 4. Prepare tabular data for all tabs + gid=0 Master Leger
  const allSubjects = getAllSubjectNames();

  const gid0Headers = [
    'ID',
    'NIS',
    'NISN',
    'Nama Siswa',
    'Kelas',
    'Tingkat',
    'Konsentrasi Keahlian',
    'Sakit',
    'Izin',
    'Alpa',
    'Catatan Wali Kelas',
    'Total Nilai',
    'Rata-Rata',
    ...allSubjects,
  ];

  const gid0Values: (string | number)[][] = [gid0Headers];
  payload.students.forEach((st) => {
    const gMap = payload.studentGrades[st.id] || {};
    const numericScores = Object.values(gMap).filter(
      (v): v is number => typeof v === 'number' && !Number.isNaN(v)
    );
    const total = numericScores.reduce((a, b) => a + b, 0);
    const avg =
      numericScores.length > 0
        ? Number((total / numericScores.length).toFixed(2))
        : 0;

    gid0Values.push([
      st.id,
      st.nis,
      st.nisn,
      st.name,
      st.className,
      st.grade,
      st.major,
      st.attendance?.sakit || 0,
      st.attendance?.izin || 0,
      st.attendance?.alpa || 0,
      st.note || '',
      total,
      avg,
      ...allSubjects.map((sub) =>
        gMap[sub] !== undefined && gMap[sub] !== null ? gMap[sub] : ''
      ),
    ]);
  });

  const sekolahValues: (string | number)[][] = [
    ['Key', 'Value'],
    ['name', payload.school.name],
    ['npsn', payload.school.npsn],
    ['logoUrl', payload.school.logoUrl || DEFAULT_SCHOOL.logoUrl],
    ['kopUrl', payload.school.kopUrl],
    ['principal', payload.school.principal],
    ['principalNip', payload.school.principalNip],
    ['tp', payload.school.tp],
    ['datePrint', payload.school.datePrint],
  ];

  const kelasValues: (string | number)[][] = [
    ['ID', 'Nama Rombel', 'Tingkat', 'Konsentrasi Keahlian', 'Wali Kelas', 'NIP Wali Kelas'],
    ...payload.classes.map((c) => [
      c.id,
      c.name,
      c.grade,
      c.major,
      c.waliKelasName,
      c.waliKelasNip,
    ]),
  ];

  const siswaValues: (string | number)[][] = [
    [
      'ID',
      'NIS',
      'NISN',
      'Nama Lengkap',
      'Kelas',
      'Tingkat',
      'Konsentrasi Keahlian',
      'Sakit',
      'Izin',
      'Alpa',
      'Catatan Wali Kelas',
    ],
    ...payload.students.map((s) => [
      s.id,
      s.nis,
      s.nisn,
      s.name,
      s.className,
      s.grade,
      s.major,
      s.attendance?.sakit || 0,
      s.attendance?.izin || 0,
      s.attendance?.alpa || 0,
      s.note || '',
    ]),
  ];

  const usersValues: (string | number)[][] = [
    ['ID', 'Username', 'Password', 'Nama Lengkap', 'NIP', 'Roles', 'Wali Kelas Untuk'],
    ...payload.users.map((u) => [
      u.id,
      u.username,
      u.password,
      u.name,
      u.nip || '',
      u.roles.join(','),
      u.waliKelasFor || '',
    ]),
  ];

  const nilaiValues: (string | number)[][] = [
    ['Student ID', 'Nama Siswa', 'Kelas', 'Mata Pelajaran', 'Nilai STS'],
  ];
  payload.students.forEach((s) => {
    const gradesForStudent = payload.studentGrades[s.id] || {};
    Object.entries(gradesForStudent).forEach(([subjectName, score]) => {
      if (score !== undefined && score !== null && String(score) !== '') {
        nilaiValues.push([s.id, s.name, s.className, subjectName, score]);
      }
    });
  });

  // 5. Batch update values across all sheets
  const batchWriteRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: `'${gid0Title}'!A1`, values: gid0Values },
          { range: `'Sekolah'!A1`, values: sekolahValues },
          { range: `'Kelas'!A1`, values: kelasValues },
          { range: `'Siswa'!A1`, values: siswaValues },
          { range: `'Users'!A1`, values: usersValues },
          { range: `'Nilai'!A1`, values: nilaiValues },
        ],
      }),
    }
  );

  if (!batchWriteRes.ok) {
    const errBody = await batchWriteRes.json().catch(() => ({}));
    throw new Error(
      errBody?.error?.message || 'Gagal menyimpan data ke Google Sheet.'
    );
  }

  const finalTabs = Array.from(new Set([...existingTabs, ...REQUIRED_TABS]));

  return {
    spreadsheetTitle,
    gid0Title,
    tabs: finalTabs,
    gid0PreviewRows: gid0Values.slice(0, 15).map((r) => r.map(String)),
    hasStructuredTabs: true,
    lastSyncedAt: new Date().toLocaleTimeString('id-ID'),
  };
}
