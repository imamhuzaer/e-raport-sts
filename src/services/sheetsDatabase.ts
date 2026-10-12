import * as XLSX from 'xlsx';
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
  SignatureDisplayType,
  SPREADSHEET_ID,
  SPREADSHEET_URL,
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
  connectionMode?: 'oauth' | 'public_direct' | 'gas_bridge';
}

const REQUIRED_TABS = ['Sekolah', 'Kelas', 'Siswa', 'Users', 'Nilai'];
const LOCAL_DB_CACHE_KEY = 'eraport_smkn1_db_cache_v2';
const GAS_BRIDGE_URL_KEY = 'eraport_smkn1_gas_bridge_url';
const CUSTOM_SPREADSHEET_ID_KEY = 'eraport_smkn1_custom_spreadsheet_id';
const CUSTOM_SPREADSHEET_URL_KEY = 'eraport_smkn1_custom_spreadsheet_url';

export function extractSpreadsheetId(input: string): string | null {
  const trimmed = (input || '').trim();
  if (!trimmed) return null;
  // Match standard Google Spreadsheet URL: /spreadsheets/d/{ID}
  const urlMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch && urlMatch[1]) {
    return urlMatch[1];
  }
  // Or direct Spreadsheet ID (alphanumeric, hyphens, underscores, length >= 20)
  if (/^[a-zA-Z0-9-_]{20,}$/.test(trimmed)) {
    return trimmed;
  }
  return null;
}

export function getActiveSpreadsheetId(): string {
  if (typeof window === 'undefined') return SPREADSHEET_ID;
  const savedId = localStorage.getItem(CUSTOM_SPREADSHEET_ID_KEY);
  return savedId && savedId.trim() ? savedId.trim() : SPREADSHEET_ID;
}

export function getActiveSpreadsheetUrl(): string {
  if (typeof window === 'undefined') return SPREADSHEET_URL;
  const savedUrl = localStorage.getItem(CUSTOM_SPREADSHEET_URL_KEY);
  if (savedUrl && savedUrl.trim()) return savedUrl.trim();
  const id = getActiveSpreadsheetId();
  if (id !== SPREADSHEET_ID) {
    return `https://docs.google.com/spreadsheets/d/${id}/edit?gid=0#gid=0`;
  }
  return SPREADSHEET_URL;
}

export function saveActiveSpreadsheetLink(
  urlOrId: string
): { id: string; url: string } | null {
  const extractedId = extractSpreadsheetId(urlOrId);
  if (!extractedId) return null;
  const normalizedUrl = urlOrId.trim().startsWith('http')
    ? urlOrId.trim()
    : `https://docs.google.com/spreadsheets/d/${extractedId}/edit?gid=0#gid=0`;
  if (typeof window !== 'undefined') {
    localStorage.setItem(CUSTOM_SPREADSHEET_ID_KEY, extractedId);
    localStorage.setItem(CUSTOM_SPREADSHEET_URL_KEY, normalizedUrl);
  }
  return { id: extractedId, url: normalizedUrl };
}

export function resetActiveSpreadsheetLink(): { id: string; url: string } {
  if (typeof window !== 'undefined') {
    localStorage.removeItem(CUSTOM_SPREADSHEET_ID_KEY);
    localStorage.removeItem(CUSTOM_SPREADSHEET_URL_KEY);
  }
  return { id: SPREADSHEET_ID, url: SPREADSHEET_URL };
}

export function getSavedGasBridgeUrl(): string {
  if (typeof window === 'undefined') return '';
  return localStorage.getItem(GAS_BRIDGE_URL_KEY) || '';
}

export function saveGasBridgeUrl(url: string): void {
  if (typeof window === 'undefined') return;
  if (url.trim()) {
    localStorage.setItem(GAS_BRIDGE_URL_KEY, url.trim());
  } else {
    localStorage.removeItem(GAS_BRIDGE_URL_KEY);
  }
}

export function sanitizeDatabasePayload(
  payload: DatabasePayload
): DatabasePayload {
  const seenClassIds = new Set<number>();
  let nextClassId = 1;
  const getNextClassId = () => {
    while (seenClassIds.has(nextClassId)) nextClassId++;
    const id = nextClassId;
    seenClassIds.add(id);
    return id;
  };

  const rawClasses = Array.isArray(payload.classes) ? payload.classes : [];
  const classIdCounts = new Map<number, number>();
  rawClasses.forEach((c) => {
    const numId = Number(c?.id);
    if (Number.isFinite(numId) && numId > 0) {
      classIdCounts.set(numId, (classIdCounts.get(numId) || 0) + 1);
    }
  });

  const classes: ClassItem[] = rawClasses.map((c) => {
    const numId = Number(c?.id);
    const isTruncatedSci = numId >= 1_000_000_000_000 && numId % 100000 === 0;
    const isValidUnique =
      Number.isFinite(numId) &&
      numId > 0 &&
      !isTruncatedSci &&
      (classIdCounts.get(numId) || 0) === 1 &&
      !seenClassIds.has(numId);

    if (isValidUnique) {
      seenClassIds.add(numId);
      return { ...c, id: numId };
    }
    return { ...c, id: getNextClassId() };
  });

  const seenUserIds = new Set<number>();
  let nextUserId = 1;
  const getNextUserId = () => {
    while (seenUserIds.has(nextUserId)) nextUserId++;
    const id = nextUserId;
    seenUserIds.add(id);
    return id;
  };

  const rawUsers = Array.isArray(payload.users) ? payload.users : [];
  const userIdCounts = new Map<number, number>();
  rawUsers.forEach((u) => {
    const numId = Number(u?.id);
    if (Number.isFinite(numId) && numId > 0) {
      userIdCounts.set(numId, (userIdCounts.get(numId) || 0) + 1);
    }
  });

  const users: UserAccount[] = rawUsers.map((u) => {
    const numId = Number(u?.id);
    const isTruncatedSci = numId >= 1_000_000_000_000 && numId % 100000 === 0;
    const isValidUnique =
      Number.isFinite(numId) &&
      numId > 0 &&
      !isTruncatedSci &&
      (userIdCounts.get(numId) || 0) === 1 &&
      !seenUserIds.has(numId);

    if (isValidUnique) {
      seenUserIds.add(numId);
      return { ...u, id: numId };
    }
    return { ...u, id: getNextUserId() };
  });

  const seenStudentIds = new Set<number>();
  let nextStudentId = 101;
  const getNextStudentId = () => {
    while (seenStudentIds.has(nextStudentId)) nextStudentId++;
    const id = nextStudentId;
    seenStudentIds.add(id);
    return id;
  };

  const rawStudents = Array.isArray(payload.students) ? payload.students : [];
  const studentIdCounts = new Map<number, number>();
  rawStudents.forEach((s) => {
    const numId = Number(s?.id);
    if (Number.isFinite(numId) && numId > 0) {
      studentIdCounts.set(numId, (studentIdCounts.get(numId) || 0) + 1);
    }
  });

  const rawGrades = payload.studentGrades || {};
  const studentGrades: StudentGradesMap = {};

  const students: StudentItem[] = rawStudents.map((s) => {
    const oldId = Number(s?.id);
    const isTruncatedSci = oldId >= 1_000_000_000_000 && oldId % 100000 === 0;
    const isValidUnique =
      Number.isFinite(oldId) &&
      oldId > 0 &&
      !isTruncatedSci &&
      (studentIdCounts.get(oldId) || 0) === 1 &&
      !seenStudentIds.has(oldId);

    const finalId = isValidUnique ? oldId : getNextStudentId();
    seenStudentIds.add(finalId);

    const existingGradeMap = rawGrades[finalId] || rawGrades[oldId] || {};
    studentGrades[finalId] = { ...existingGradeMap };

    return {
      ...s,
      id: finalId,
    };
  });

  return {
    school: payload.school || { ...DEFAULT_SCHOOL },
    classes,
    students,
    users,
    studentGrades,
  };
}

function sheetToCleanStringRows(sheet: XLSX.WorkSheet): string[][] {
  const rawRows = XLSX.utils.sheet_to_json<(string | number | boolean)[]>(
    sheet,
    {
      header: 1,
      defval: '',
      raw: true,
    }
  );
  const fmtRows = XLSX.utils.sheet_to_json<(string | number)[]>(sheet, {
    header: 1,
    defval: '',
    raw: false,
  });

  const rowCount = Math.max(rawRows.length, fmtRows.length);
  const result: string[][] = [];

  for (let r = 0; r < rowCount; r++) {
    const rawRow = rawRows[r] || [];
    const fmtRow = fmtRows[r] || [];
    const colCount = Math.max(rawRow.length, fmtRow.length);
    const rowOut: string[] = [];
    for (let c = 0; c < colCount; c++) {
      const rawCell = rawRow[c];
      const fmtCell = String(fmtRow[c] ?? '').trim();
      if (
        typeof rawCell === 'number' &&
        Number.isFinite(rawCell) &&
        (/^\d+(\.\d+)?e[+-]?\d+$/i.test(fmtCell) || Math.abs(rawCell) >= 1e11)
      ) {
        rowOut.push(
          Math.round(rawCell).toLocaleString('fullwide', {
            useGrouping: false,
          })
        );
      } else {
        rowOut.push(fmtCell || String(rawCell ?? '').trim());
      }
    }
    result.push(rowOut);
  }
  return result;
}

export function saveLocalDatabaseCache(payload: DatabasePayload): void {
  if (typeof window === 'undefined') return;
  try {
    const clean = sanitizeDatabasePayload(payload);
    localStorage.setItem(LOCAL_DB_CACHE_KEY, JSON.stringify(clean));
  } catch {
    // Ignore storage quota errors
  }
}

export function loadLocalDatabaseCache(): DatabasePayload | null {
  if (typeof window === 'undefined') return null;
  try {
    const raw = localStorage.getItem(LOCAL_DB_CACHE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as DatabasePayload;
    if (
      parsed &&
      parsed.school &&
      Array.isArray(parsed.classes) &&
      Array.isArray(parsed.students)
    ) {
      const sanitized = sanitizeDatabasePayload(parsed);
      localStorage.setItem(LOCAL_DB_CACHE_KEY, JSON.stringify(sanitized));
      return sanitized;
    }
  } catch {
    // Ignore parse error
  }
  return null;
}

function getAllSubjectNames(): string[] {
  const set = new Set<string>();
  Object.values(DEFAULT_SUBJECT_TEMPLATES).forEach((tpl) => {
    tpl.umum.forEach((s) => set.add(s.name));
    tpl.kejuruan.forEach((s) => set.add(s.name));
  });
  return Array.from(set);
}

function parseSheetTabsToDatabase(
  tabValuesMap: Record<string, string[][]>,
  gid0Rows: string[][],
  spreadsheetTitle: string,
  gid0Title: string,
  tabs: string[],
  connectionMode: 'oauth' | 'public_direct' | 'gas_bridge' = 'oauth'
): {
  payload: DatabasePayload;
  meta: SheetMetadataInfo;
  isEmptySheet: boolean;
} {
  const hasStructuredTabs = REQUIRED_TABS.every((t) => tabs.includes(t));

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
        connectionMode,
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
      if (r[0]) kv[String(r[0]).trim()] = String(r[1] ?? '');
    });
    if (
      sekolahRows[0] &&
      sekolahRows[0][0] &&
      String(sekolahRows[0][0]).toLowerCase() !== 'key'
    ) {
      kv[String(sekolahRows[0][0]).trim()] = String(sekolahRows[0][1] ?? '');
    }
    const rawKsSigType = (kv['principalSignatureType'] || 'none')
      .trim()
      .toLowerCase();
    const principalSignatureType: SignatureDisplayType =
      rawKsSigType === 'signature' || rawKsSigType === 'qrcode'
        ? rawKsSigType
        : 'none';

    school = {
      name: kv['name'] || DEFAULT_SCHOOL.name,
      npsn: kv['npsn'] || DEFAULT_SCHOOL.npsn,
      logoUrl: kv['logoUrl'] || DEFAULT_SCHOOL.logoUrl,
      kopUrl: kv['kopUrl'] || DEFAULT_SCHOOL.kopUrl,
      principal: kv['principal'] || DEFAULT_SCHOOL.principal,
      principalNip: kv['principalNip'] || DEFAULT_SCHOOL.principalNip,
      tp: kv['tp'] || DEFAULT_SCHOOL.tp,
      datePrint: kv['datePrint'] || DEFAULT_SCHOOL.datePrint,
      principalSignatureType,
      principalSignatureUrl: kv['principalSignatureUrl'] || '',
    };
    if (kv['gasBridgeUrl'] && !getSavedGasBridgeUrl()) {
      saveGasBridgeUrl(kv['gasBridgeUrl']);
    }
  }

  // Parse Classes with guaranteed unique IDs
  let classes: ClassItem[] = [...DEFAULT_CLASSES];
  const kelasRows = tabValuesMap['Kelas'] || [];
  if (kelasRows.length > 1) {
    const usedClassIds = new Set<number>();
    let nextClassId = 1;
    classes = kelasRows
      .slice(1)
      .filter((r) => r.length >= 2 && r[1])
      .map((r) => {
        const rawWaliSigType = String(r[6] || 'none')
          .trim()
          .toLowerCase();
        const waliSignatureType: SignatureDisplayType =
          rawWaliSigType === 'signature' || rawWaliSigType === 'qrcode'
            ? rawWaliSigType
            : 'none';
        let id = Number(r[0]);
        const isTruncatedSci = id >= 1_000_000_000_000 && id % 100000 === 0;
        if (
          !Number.isFinite(id) ||
          id <= 0 ||
          isTruncatedSci ||
          usedClassIds.has(id)
        ) {
          while (usedClassIds.has(nextClassId)) nextClassId++;
          id = nextClassId;
        }
        usedClassIds.add(id);
        return {
          id,
          name: String(r[1] || '').trim(),
          grade: String(r[2] || '10').trim(),
          major: String(r[3] || 'Teknik Kendaraan Ringan').trim(),
          waliKelasName: String(r[4] || '-').trim(),
          waliKelasNip: String(r[5] || '-').trim(),
          waliSignatureType,
          waliSignatureUrl: String(r[7] || '').trim(),
        };
      });
  }

  // Parse Users with guaranteed unique IDs
  let users: UserAccount[] = JSON.parse(JSON.stringify(DEFAULT_USERS));
  const userRows = tabValuesMap['Users'] || [];
  if (userRows.length > 1) {
    const usedUserIds = new Set<number>();
    let nextUserId = 1;
    users = userRows
      .slice(1)
      .filter((r) => r.length >= 3 && r[1])
      .map((r) => {
        const rawRoles = String(r[5] || 'gurumapel')
          .split(',')
          .map((x) => x.trim().toLowerCase())
          .filter(
            (x): x is RoleType =>
              x === 'admin' || x === 'walikelas' || x === 'gurumapel'
          );
        const rawUserSigType = String(r[7] || 'none')
          .trim()
          .toLowerCase();
        const waliSignatureType: SignatureDisplayType =
          rawUserSigType === 'signature' || rawUserSigType === 'qrcode'
            ? rawUserSigType
            : 'none';
        let id = Number(r[0]);
        const isTruncatedSci = id >= 1_000_000_000_000 && id % 100000 === 0;
        if (
          !Number.isFinite(id) ||
          id <= 0 ||
          isTruncatedSci ||
          usedUserIds.has(id)
        ) {
          while (usedUserIds.has(nextUserId)) nextUserId++;
          id = nextUserId;
        }
        usedUserIds.add(id);
        return {
          id,
          username: String(r[1] || '').trim(),
          password: String(r[2] || '123').trim(),
          name: String(r[3] || r[1] || '').trim(),
          nip: String(r[4] || '').trim(),
          roles: rawRoles.length > 0 ? rawRoles : ['gurumapel'],
          waliKelasFor: String(r[6] || '').trim(),
          waliSignatureType,
          waliSignatureUrl: String(r[8] || '').trim(),
        };
      });
  }

  // Parse Students & Grades with guaranteed unique IDs
  let students: StudentItem[] = [];
  const studentGrades: StudentGradesMap = {};

  const siswaRows = tabValuesMap['Siswa'] || [];
  const nilaiRows = tabValuesMap['Nilai'] || [];

  if (siswaRows.length > 1) {
    const usedStudentIds = new Set<number>();
    let nextStudentId = 101;
    const rawIdToNewIdsMap = new Map<number, number[]>();

    students = siswaRows
      .slice(1)
      .filter((r) => r.length >= 4 && r[3])
      .map((r) => {
        const rawId = Number(r[0]);
        let id = rawId;
        const isTruncatedSci = id >= 1_000_000_000_000 && id % 100000 === 0;
        if (
          !Number.isFinite(id) ||
          id <= 0 ||
          isTruncatedSci ||
          usedStudentIds.has(id)
        ) {
          while (usedStudentIds.has(nextStudentId)) nextStudentId++;
          id = nextStudentId;
        }
        usedStudentIds.add(id);
        if (Number.isFinite(rawId) && rawId > 0) {
          const list = rawIdToNewIdsMap.get(rawId) || [];
          list.push(id);
          rawIdToNewIdsMap.set(rawId, list);
        }
        studentGrades[id] = {};
        return {
          id,
          nis: String(r[1] || '').trim(),
          nisn: String(r[2] || '').trim(),
          name: String(r[3] || '').trim(),
          className: String(r[4] || '10 TKR 1').trim(),
          grade: String(r[5] || '10').trim(),
          major: String(r[6] || 'Teknik Kendaraan Ringan').trim(),
          attendance: {
            sakit: Number(r[7]) || 0,
            izin: Number(r[8]) || 0,
            alpa: Number(r[9]) || 0,
          },
          note:
            String(r[10] || '').trim() ||
            'Tingkatkan terus motivasi dan kedisiplinan belajar.',
        };
      });

    // Build student lookup maps for accurate grade resolution even if raw Student IDs collided
    const byNameAndClass = new Map<string, number>();
    const byNameOnly = new Map<string, number>();
    const byNisn = new Map<string, number>();
    students.forEach((st) => {
      const nameKey = st.name.trim().toLowerCase();
      const clsKey = st.className.trim().toLowerCase();
      byNameAndClass.set(`${nameKey}|${clsKey}`, st.id);
      if (!byNameOnly.has(nameKey)) {
        byNameOnly.set(nameKey, st.id);
      }
      if (st.nisn) {
        byNisn.set(st.nisn.trim().toLowerCase(), st.id);
      }
    });

    if (nilaiRows.length > 1) {
      nilaiRows.slice(1).forEach((r) => {
        const rawStId = Number(r[0]);
        const stName = String(r[1] || '').trim().toLowerCase();
        const stClass = String(r[2] || '').trim().toLowerCase();
        const subjectName = String(r[3] || '').trim();
        const scoreVal = r[4];

        if (
          subjectName &&
          scoreVal !== undefined &&
          String(scoreVal).trim() !== ''
        ) {
          const num = Number(scoreVal);
          if (!Number.isNaN(num)) {
            let targetStId: number | undefined;
            if (stName && stClass && byNameAndClass.has(`${stName}|${stClass}`)) {
              targetStId = byNameAndClass.get(`${stName}|${stClass}`);
            } else if (stName && byNameOnly.has(stName)) {
              targetStId = byNameOnly.get(stName);
            } else if (rawStId) {
              const mappedIds = rawIdToNewIdsMap.get(rawStId);
              if (mappedIds && mappedIds.length === 1) {
                targetStId = mappedIds[0];
              } else if (studentGrades[rawStId]) {
                targetStId = rawStId;
              }
            }

            if (targetStId !== undefined) {
              if (!studentGrades[targetStId]) studentGrades[targetStId] = {};
              studentGrades[targetStId][subjectName] = num;
            }
          }
        }
      });
    }

    // Also recover any grades present in gid0Rows (Master_Leger) if missing from Nilai
    if (hasGid0Data) {
      const headers = gid0Rows[0].map((h) => String(h || '').trim());
      const lowerHeaders = headers.map((h) => h.toLowerCase());
      const idxNisn = lowerHeaders.findIndex((h) => h === 'nisn');
      const idxName = lowerHeaders.findIndex(
        (h) => h.includes('nama') || h.includes('name') || h.includes('siswa')
      );
      const idxClass = lowerHeaders.findIndex(
        (h) => h.includes('kelas') || h.includes('rombel') || h === 'classname'
      );
      const allSubjects = getAllSubjectNames();

      gid0Rows.slice(1).forEach((r) => {
        const rowNisn =
          idxNisn !== -1 ? String(r[idxNisn] || '').trim().toLowerCase() : '';
        const rowName =
          idxName !== -1 ? String(r[idxName] || '').trim().toLowerCase() : '';
        const rowClass =
          idxClass !== -1 ? String(r[idxClass] || '').trim().toLowerCase() : '';

        const targetStId =
          (rowName && rowClass && byNameAndClass.get(`${rowName}|${rowClass}`)) ||
          (rowNisn && byNisn.get(rowNisn)) ||
          (rowName && byNameOnly.get(rowName));

        if (targetStId !== undefined) {
          if (!studentGrades[targetStId]) studentGrades[targetStId] = {};
          headers.forEach((hName, colIndex) => {
            const matchedSub = allSubjects.find(
              (sub) => sub.toLowerCase() === hName.toLowerCase()
            );
            if (
              matchedSub &&
              studentGrades[targetStId][matchedSub] === undefined &&
              r[colIndex] !== undefined &&
              String(r[colIndex]).trim() !== ''
            ) {
              const val = Number(r[colIndex]);
              if (!Number.isNaN(val)) {
                studentGrades[targetStId][matchedSub] = val;
              }
            }
          });
        }
      });
    }
  } else if (hasGid0Data) {
    const headers = gid0Rows[0].map((h) => String(h || '').trim());
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
      const usedStudentIds = new Set<number>();
      let nextStudentId = 101;
      students = gid0Rows
        .slice(1)
        .filter((r) => r[idxName] && String(r[idxName]).trim() !== '')
        .map((r, i) => {
          let stId = idxId !== -1 ? Number(r[idxId]) : 0;
          const isTruncatedSci =
            stId >= 1_000_000_000_000 && stId % 100000 === 0;
          if (
            !Number.isFinite(stId) ||
            stId <= 0 ||
            isTruncatedSci ||
            usedStudentIds.has(stId)
          ) {
            while (usedStudentIds.has(nextStudentId)) nextStudentId++;
            stId = nextStudentId;
          }
          usedStudentIds.add(stId);
          const clsName =
            String(idxClass !== -1 ? r[idxClass] : '').trim() || '10 TKR 1';
          const inferredGrade =
            String(idxGrade !== -1 ? r[idxGrade] : '').trim() ||
            (clsName.startsWith('12')
              ? '12'
              : clsName.startsWith('11')
              ? '11'
              : '10');
          const major =
            String(idxMajor !== -1 ? r[idxMajor] : '').trim() ||
            'Teknik Kendaraan Ringan';

          studentGrades[stId] = {};
          headers.forEach((hName, colIndex) => {
            const matchedSub = allSubjects.find(
              (sub) => sub.toLowerCase() === hName.toLowerCase()
            );
            if (
              matchedSub &&
              r[colIndex] !== undefined &&
              String(r[colIndex]).trim() !== ''
            ) {
              const val = Number(r[colIndex]);
              if (!Number.isNaN(val)) {
                studentGrades[stId][matchedSub] = val;
              }
            }
          });

          return {
            id: stId,
            nis:
              String(idxNis !== -1 ? r[idxNis] : '').trim() || `2026100${i + 1}`,
            nisn:
              String(idxNisn !== -1 ? r[idxNisn] : '').trim() ||
              `008123456${i + 1}`,
            name: String(r[idxName]).trim().toUpperCase(),
            className: clsName,
            grade: inferredGrade,
            major,
            attendance: {
              sakit: idxSakit !== -1 ? Number(r[idxSakit]) || 0 : 0,
              izin: idxIzin !== -1 ? Number(r[idxIzin]) || 0 : 0,
              alpa: idxAlpa !== -1 ? Number(r[idxAlpa]) || 0 : 0,
            },
            note:
              String(idxNote !== -1 ? r[idxNote] : '').trim() ||
              'Tingkatkan terus motivasi dan kedisiplinan belajar.',
          };
        });

      const existingClassNames = new Set(classes.map((c) => c.name));
      let maxClsId = classes.reduce((acc, c) => Math.max(acc, c.id || 0), 0);
      students.forEach((st) => {
        if (st.className && !existingClassNames.has(st.className)) {
          existingClassNames.add(st.className);
          maxClsId += 1;
          classes.push({
            id: maxClsId,
            name: st.className,
            grade: st.grade,
            major: st.major,
            waliKelasName: '-',
            waliKelasNip: '-',
          });
        }
      });
    } else {
      students = JSON.parse(JSON.stringify(DEFAULT_STUDENTS));
      Object.assign(
        studentGrades,
        JSON.parse(JSON.stringify(DEFAULT_STUDENT_GRADES))
      );
    }
  }

  // Cross-link Wali Kelas from Users into Classes if missing
  const normKey = (s?: string) => (s || '').replace(/\s+/g, '').toLowerCase();
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
        normKey(u.waliKelasFor) === normKey(cls.name)
    );

    if (matchedWaliUser) {
      return {
        ...cls,
        waliKelasName: isMissingWali ? matchedWaliUser.name : cls.waliKelasName,
        waliKelasNip: isMissingNip
          ? matchedWaliUser.nip || '-'
          : cls.waliKelasNip,
        waliSignatureType:
          cls.waliSignatureType && cls.waliSignatureType !== 'none'
            ? cls.waliSignatureType
            : matchedWaliUser.waliSignatureType || 'none',
        waliSignatureUrl:
          cls.waliSignatureUrl || matchedWaliUser.waliSignatureUrl || '',
      };
    }
    return cls;
  });

  const payload = sanitizeDatabasePayload({
    school,
    classes,
    students,
    users,
    studentGrades,
  });
  saveLocalDatabaseCache(payload);

  return {
    payload,
    meta: {
      spreadsheetTitle,
      gid0Title,
      tabs,
      gid0PreviewRows: gid0Rows.slice(0, 15).map((r) => r.map(String)),
      hasStructuredTabs,
      lastSyncedAt: new Date().toLocaleTimeString('id-ID'),
      connectionMode,
    },
    isEmptySheet: false,
  };
}

/**
 * Direct Zero-Auth Public Reader for Google Sheets.
 * Works on ANY domain (including *.vercel.app, localhost, custom domains) WITHOUT requiring Google OAuth login!
 */
export async function fetchFromGoogleSheetsPublic(): Promise<{
  payload: DatabasePayload;
  meta: SheetMetadataInfo;
  isEmptySheet: boolean;
}> {
  const activeSheetId = getActiveSpreadsheetId();

  // Strategy 0: If Google Apps Script Web App Bridge URL is configured, try reading via doGet
  const gasUrl = getSavedGasBridgeUrl();
  if (gasUrl) {
    try {
      const gasRes = await fetch(
        `${gasUrl}${
          gasUrl.includes('?') ? '&' : '?'
        }action=readAll&spreadsheetId=${encodeURIComponent(
          activeSheetId
        )}&t=${Date.now()}`
      );
      if (gasRes.ok) {
        const gasJson = await gasRes.json();
        if (gasJson && gasJson.status === 'ok' && gasJson.sheets) {
          const tabValuesMap: Record<string, string[][]> = {};
          const tabs = Object.keys(gasJson.sheets);
          tabs.forEach((tabName) => {
            const rawRows = gasJson.sheets[tabName] || [];
            tabValuesMap[tabName] = rawRows.map((r: any[]) =>
              r.map((cell) => String(cell ?? '').trim())
            );
          });
          const gid0Title = gasJson.gid0Title || tabs[0] || 'Sheet1';
          const gid0Rows = tabValuesMap[gid0Title] || [];
          return parseSheetTabsToDatabase(
            tabValuesMap,
            gid0Rows,
            gasJson.spreadsheetTitle ||
              'Database E-Raport STS SMKN 1 Mesuji Timur',
            gid0Title,
            tabs,
            'gas_bridge'
          );
        }
      }
    } catch (err) {
      console.warn('GAS Bridge read fallback to public XLSX export:', err);
    }
  }

  // Strategy 1: Fetch full workbook via /export?format=xlsx (single request for all tabs!)
  try {
    const exportUrl = `https://docs.google.com/spreadsheets/d/${activeSheetId}/export?format=xlsx&t=${Date.now()}`;
    const res = await fetch(exportUrl);
    if (res.ok) {
      const arrayBuf = await res.arrayBuffer();
      const wb = XLSX.read(arrayBuf, { type: 'array' });
      const tabs = wb.SheetNames || [];
      const gid0Title = tabs[0] || 'Sheet1';
      const tabValuesMap: Record<string, string[][]> = {};

      tabs.forEach((tabName) => {
        const sheet = wb.Sheets[tabName];
        if (sheet) {
          tabValuesMap[tabName] = sheetToCleanStringRows(sheet);
        }
      });

      const gid0Rows = tabValuesMap[gid0Title] || [];
      return parseSheetTabsToDatabase(
        tabValuesMap,
        gid0Rows,
        'Database E-Raport STS SMKN 1 Mesuji Timur',
        gid0Title,
        tabs,
        'public_direct'
      );
    }
  } catch (err) {
    console.warn('Public XLSX export fallback to gviz CSV:', err);
  }

  // Strategy 2: Fetch each tab via /gviz/tq?tqx=out:csv
  const tabValuesMap: Record<string, string[][]> = {};
  const detectedTabs: string[] = [];

  for (const tabName of ['Sheet1', ...REQUIRED_TABS]) {
    try {
      const csvUrl = `https://docs.google.com/spreadsheets/d/${activeSheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(
        tabName
      )}&t=${Date.now()}`;
      const res = await fetch(csvUrl);
      if (res.ok) {
        const csvText = await res.text();
        const wb = XLSX.read(csvText, { type: 'string' });
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        if (firstSheet) {
          tabValuesMap[tabName] = sheetToCleanStringRows(firstSheet);
          detectedTabs.push(tabName);
        }
      }
    } catch {
      // Ignore individual tab fetch error
    }
  }

  if (detectedTabs.length === 0) {
    throw new Error(
      'Tidak dapat menghubungi Google Spreadsheet. Pastikan link Spreadsheet benar dan akses disetel "Siapa saja yang memiliki link".'
    );
  }

  const gid0Rows = tabValuesMap['Sheet1'] || [];
  return parseSheetTabsToDatabase(
    tabValuesMap,
    gid0Rows,
    'Database E-Raport STS SMKN 1 Mesuji Timur',
    'Sheet1',
    detectedTabs,
    'public_direct'
  );
}

export async function fetchFromGoogleSheets(
  accessToken?: string | null
): Promise<{
  payload: DatabasePayload;
  meta: SheetMetadataInfo;
  isEmptySheet: boolean;
}> {
  if (!accessToken) {
    return fetchFromGoogleSheetsPublic();
  }

  const activeSheetId = getActiveSpreadsheetId();

  try {
    // 1. Fetch spreadsheet metadata via OAuth
    const metaRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!metaRes.ok) {
      return await fetchFromGoogleSheetsPublic();
    }

    const metaJson = await metaRes.json();
    const spreadsheetTitle: string =
      metaJson.properties?.title || 'Database E-Raport STS';
    const sheetsList: Array<{ properties: { title: string; sheetId: number } }> =
      metaJson.sheets || [];

    const tabs = sheetsList.map((s) => s.properties.title);
    const gid0Sheet =
      sheetsList.find((s) => s.properties.sheetId === 0) || sheetsList[0];
    const gid0Title = gid0Sheet ? gid0Sheet.properties.title : 'Sheet1';

    // Build ranges to batchGet
    const rangesToFetch: string[] = [`'${gid0Title}'!A1:Z300`];
    REQUIRED_TABS.forEach((tab) => {
      if (tabs.includes(tab)) {
        rangesToFetch.push(`'${tab}'!A1:Z600`);
      }
    });

    const params = rangesToFetch
      .map((r) => `ranges=${encodeURIComponent(r)}`)
      .join('&');
    const valRes = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values:batchGet?${params}`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (!valRes.ok) {
      return await fetchFromGoogleSheetsPublic();
    }

    const valJson = await valRes.json();
    const valueRanges: Array<{ range: string; values?: string[][] }> =
      valJson.valueRanges || [];

    const gid0Rows: string[][] = valueRanges[0]?.values || [];

    const tabValuesMap: Record<string, string[][]> = {};
    let rangeIdx = 1;
    REQUIRED_TABS.forEach((tab) => {
      if (tabs.includes(tab)) {
        tabValuesMap[tab] = valueRanges[rangeIdx]?.values || [];
        rangeIdx++;
      }
    });

    return parseSheetTabsToDatabase(
      tabValuesMap,
      gid0Rows,
      spreadsheetTitle,
      gid0Title,
      tabs,
      'oauth'
    );
  } catch {
    return fetchFromGoogleSheetsPublic();
  }
}

function buildSheetTabularValues(payload: DatabasePayload) {
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

  const savedGasUrl = getSavedGasBridgeUrl();
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
    ['principalSignatureType', payload.school.principalSignatureType || 'none'],
    ['principalSignatureUrl', payload.school.principalSignatureUrl || ''],
    ...(savedGasUrl ? [['gasBridgeUrl', savedGasUrl]] : []),
  ];

  const kelasValues: (string | number)[][] = [
    [
      'ID',
      'Nama Rombel',
      'Tingkat',
      'Konsentrasi Keahlian',
      'Wali Kelas',
      'NIP Wali Kelas',
      'Tipe TTD/QR Wali Kelas',
      'Data/URL TTD/QR Wali Kelas',
    ],
    ...payload.classes.map((c) => [
      c.id,
      c.name,
      c.grade,
      c.major,
      c.waliKelasName,
      c.waliKelasNip,
      c.waliSignatureType || 'none',
      c.waliSignatureUrl || '',
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
    [
      'ID',
      'Username',
      'Password',
      'Nama Lengkap',
      'NIP',
      'Roles',
      'Wali Kelas Untuk',
      'Tipe TTD/QR Wali Kelas',
      'Data/URL TTD/QR Wali Kelas',
    ],
    ...payload.users.map((u) => [
      u.id,
      u.username,
      u.password,
      u.name,
      u.nip || '',
      u.roles.join(','),
      u.waliKelasFor || '',
      u.waliSignatureType || 'none',
      u.waliSignatureUrl || '',
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

  return {
    gid0Values,
    sekolahValues,
    kelasValues,
    siswaValues,
    usersValues,
    nilaiValues,
  };
}

/**
 * Writes all sheets via Google Apps Script Web App Bridge (Zero OAuth / Works on any Vercel domain for all roles)
 */
export async function syncViaGasWebApp(
  gasUrl: string,
  payload: DatabasePayload
): Promise<SheetMetadataInfo> {
  saveLocalDatabaseCache(payload);
  const tables = buildSheetTabularValues(payload);
  const activeSheetId = getActiveSpreadsheetId();

  const res = await fetch(gasUrl.trim(), {
    method: 'POST',
    headers: {
      // text/plain avoids CORS preflight OPTIONS failure on Google Apps Script Web Apps
      'Content-Type': 'text/plain;charset=utf-8',
    },
    body: JSON.stringify({
      action: 'syncAll',
      spreadsheetId: activeSheetId,
      sheets: {
        Sheet1: tables.gid0Values,
        Sekolah: tables.sekolahValues,
        Kelas: tables.kelasValues,
        Siswa: tables.siswaValues,
        Users: tables.usersValues,
        Nilai: tables.nilaiValues,
      },
    }),
  });

  if (!res.ok) {
    throw new Error(`Gagal menghubungi Web App GAS (HTTP ${res.status})`);
  }

  return {
    spreadsheetTitle: 'Database E-Raport STS SMKN 1 Mesuji Timur',
    gid0Title: 'Sheet1',
    tabs: ['Sheet1', ...REQUIRED_TABS],
    gid0PreviewRows: tables.gid0Values.slice(0, 15).map((r) => r.map(String)),
    hasStructuredTabs: true,
    lastSyncedAt: new Date().toLocaleTimeString('id-ID'),
    connectionMode: 'gas_bridge',
  };
}

export async function syncAllToGoogleSheets(
  accessToken: string | null | undefined,
  payload: DatabasePayload
): Promise<SheetMetadataInfo> {
  saveLocalDatabaseCache(payload);
  const gasUrl = getSavedGasBridgeUrl();
  const activeSheetId = getActiveSpreadsheetId();

  // If no OAuth token is present, use GAS Web App Bridge if configured
  if (!accessToken && gasUrl) {
    return syncViaGasWebApp(gasUrl, payload);
  }

  if (!accessToken) {
    throw new Error(
      'Token otorisasi Google belum tersedia. Silakan hubungkan Google atau pasang URL Web App Google Apps Script di tab Database.'
    );
  }

  // 1. Fetch current metadata to check existing sheet tabs
  const metaRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!metaRes.ok) {
    // If OAuth token expired and GAS bridge is available, fall back to GAS bridge
    if (gasUrl) {
      return syncViaGasWebApp(gasUrl, payload);
    }
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
      `https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}:batchUpdate`,
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
    `https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values:batchClear`,
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
  const tables = buildSheetTabularValues(payload);

  // 5. Batch update values across all sheets
  const batchWriteRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: `'${gid0Title}'!A1`, values: tables.gid0Values },
          { range: `'Sekolah'!A1`, values: tables.sekolahValues },
          { range: `'Kelas'!A1`, values: tables.kelasValues },
          { range: `'Siswa'!A1`, values: tables.siswaValues },
          { range: `'Users'!A1`, values: tables.usersValues },
          { range: `'Nilai'!A1`, values: tables.nilaiValues },
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
    gid0PreviewRows: tables.gid0Values.slice(0, 15).map((r) => r.map(String)),
    hasStructuredTabs: true,
    lastSyncedAt: new Date().toLocaleTimeString('id-ID'),
    connectionMode: 'oauth',
  };
}

export function getGasBridgeScriptCode(): string {
  const activeSheetId = getActiveSpreadsheetId();
  return `// ============================================================================
// GOOGLE APPS SCRIPT BRIDGE - E-RAPORT STS SMKN 1 MESUJI TIMUR
// Pasang di Spreadsheet (Ekstensi -> Apps Script -> Terapkan -> Deployment Baru -> Aplikasi Web -> Akses: Siapa saja / Anyone)
// Solusi agar aplikasi di Vercel dapat Membaca & Menyimpan ke Spreadsheet tanpa terkendala Login Google OAuth
// ============================================================================
const SPREADSHEET_ID = '${activeSheetId}';

function doGet(e) {
  try {
    const targetId = (e && e.parameter && e.parameter.spreadsheetId) ? e.parameter.spreadsheetId : SPREADSHEET_ID;
    const ss = SpreadsheetApp.openById(targetId);
    const sheets = ss.getSheets();
    const resultSheets = {};
    sheets.forEach(function(sh) {
      resultSheets[sh.getName()] = sh.getDataRange().getDisplayValues();
    });
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'ok',
        spreadsheetTitle: ss.getName(),
        gid0Title: sheets[0] ? sheets[0].getName() : 'Sheet1',
        sheets: resultSheets
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(body.spreadsheetId || SPREADSHEET_ID);
    const sheetsData = body.sheets || {};

    Object.keys(sheetsData).forEach(function(sheetName) {
      const rows = sheetsData[sheetName];
      if (!Array.isArray(rows) || rows.length === 0) return;
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
      }
      sheet.clearContents();
      const numRows = rows.length;
      const numCols = rows[0].length;
      sheet.getRange(1, 1, numRows, numCols).setValues(rows);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', updatedAt: new Date().toISOString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;
}


