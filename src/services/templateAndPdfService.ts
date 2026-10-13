import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import {
  ClassItem,
  RankedStudentItem,
  RoleType,
  SchoolData,
  StudentGradesMap,
  StudentItem,
  SubjectItem,
  UserAccount,
} from '../types';

/**
 * Formats teacher/principal names so the main name is UPPERCASE,
 * while academic titles (S.Pd., M.Pd., S.T., S.Kom., Drs., Dra., etc.)
 * use proper Indonesian EYD capitalization (never "S.PD.").
 */
export function formatTeacherNameWithTitle(rawName?: string): string {
  if (!rawName || !rawName.trim()) {
    return '-';
  }

  // Strip any stray placeholder dots (e.g. "....") from raw name
  const cleanedRaw = rawName.replace(/\.{2,}/g, '').trim();
  if (!cleanedRaw || cleanedRaw === '-') {
    return '-';
  }

  const titleReplacements: Array<[RegExp, string]> = [
    [/\bdrs\.\s*/gi, 'Drs. '],
    [/\bdra\.\s*/gi, 'Dra. '],
    [/\bdr\.\s*/gi, 'Dr. '],
    [/\bir\.\s*/gi, 'Ir. '],
    [/\bprof\.\s*/gi, 'Prof. '],
    [/\bs\.pd\.i\.?/gi, 'S.Pd.I.'],
    [/\bm\.pd\.i\.?/gi, 'M.Pd.I.'],
    [/\bs\.pd\.sd\.?/gi, 'S.Pd.SD.'],
    [/\bs\.pd\.?/gi, 'S.Pd.'],
    [/\bm\.pd\.?/gi, 'M.Pd.'],
    [/\bs\.kom\.?/gi, 'S.Kom.'],
    [/\bm\.kom\.?/gi, 'M.Kom.'],
    [/\bs\.t\.?/gi, 'S.T.'],
    [/\bm\.t\.?/gi, 'M.T.'],
    [/\bs\.si\.?/gi, 'S.Si.'],
    [/\bm\.si\.?/gi, 'M.Si.'],
    [/\bs\.ag\.?/gi, 'S.Ag.'],
    [/\bm\.ag\.?/gi, 'M.Ag.'],
    [/\bs\.hum\.?/gi, 'S.Hum.'],
    [/\bm\.hum\.?/gi, 'M.Hum.'],
    [/\bs\.sos\.?/gi, 'S.Sos.'],
    [/\bm\.sos\.?/gi, 'M.Sos.'],
    [/\bs\.e\.?/gi, 'S.E.'],
    [/\bm\.m\.?/gi, 'M.M.'],
    [/\bs\.psi\.?/gi, 'S.Psi.'],
    [/\bm\.psi\.?/gi, 'M.Psi.'],
    [/\bs\.or\.?/gi, 'S.Or.'],
    [/\bs\.sn\.?/gi, 'S.Sn.'],
    [/\bs\.tr\.kom\.?/gi, 'S.Tr.Kom.'],
    [/\bs\.tr\.t\.?/gi, 'S.Tr.T.'],
    [/\bgr\.?/gi, 'Gr.'],
  ];

  // Split by comma to separate main name from rear titles
  const parts = cleanedRaw.split(',');
  let mainName = parts[0].trim().toUpperCase();

  // Fix front titles on mainName (e.g., DRS., DRA., IR., DR.)
  mainName = mainName
    .replace(/^DRS\.\s*/i, 'Drs. ')
    .replace(/^DRA\.\s*/i, 'Dra. ')
    .replace(/^DR\.\s*/i, 'Dr. ')
    .replace(/^IR\.\s*/i, 'Ir. ')
    .replace(/^PROF\.\s*/i, 'Prof. ');

  if (parts.length === 1) {
    let result = mainName;
    titleReplacements.forEach(([regex, replacement]) => {
      result = result.replace(regex, replacement);
    });
    return result;
  }

  const rearTitlesRaw = parts.slice(1).join(', ').trim();
  let formattedRear = rearTitlesRaw;
  titleReplacements.forEach(([regex, replacement]) => {
    formattedRear = formattedRear.replace(regex, replacement);
  });

  return `${mainName}, ${formattedRear}`;
}

/**
 * Automatically detects whether an identifier is NIP or NUPTK:
 * - Empty, '-', 'belum ada', or no digits -> returns empty string (only teacher's name is shown without dots)
 * - Explicit prefix "NUPTK" or 16 digits -> "NUPTK. <nomor>"
 * - Explicit prefix "NIP" or 18 digits (or other valid digits) -> "NIP. <nomor>"
 */
export function formatNipOrNuptk(rawValue?: string): {
  type: 'NIP' | 'NUPTK' | 'NONE';
  numberOnly: string;
  formatted: string;
} {
  if (!rawValue || !rawValue.trim()) {
    return {
      type: 'NONE',
      numberOnly: '',
      formatted: '',
    };
  }

  const trimmed = rawValue.replace(/\.{2,}/g, '').trim();
  const digitsOnly = trimmed.replace(/\D/g, '');

  // If there are no actual digits (e.g., '-', '...', 'belum ada', 'tidak ada') or only zeros
  if (digitsOnly.length === 0 || /^0+$/.test(digitsOnly)) {
    return {
      type: 'NONE',
      numberOnly: '',
      formatted: '',
    };
  }

  const upper = trimmed.toUpperCase();

  if (upper.startsWith('NUPTK')) {
    const cleaned = trimmed.replace(/^NUPTK[\s.:/-]*/i, '').trim();
    if (!cleaned || !/\d/.test(cleaned)) {
      return { type: 'NONE', numberOnly: '', formatted: '' };
    }
    return {
      type: 'NUPTK',
      numberOnly: cleaned,
      formatted: `NUPTK. ${cleaned}`,
    };
  }

  if (upper.startsWith('NIP')) {
    const cleaned = trimmed
      .replace(/^(NIP\s*\/?\s*NUPTK|NIP)[\s.:/-]*/i, '')
      .trim();
    if (!cleaned || !/\d/.test(cleaned)) {
      return { type: 'NONE', numberOnly: '', formatted: '' };
    }
    if (digitsOnly.length === 16) {
      return {
        type: 'NUPTK',
        numberOnly: cleaned,
        formatted: `NUPTK. ${cleaned}`,
      };
    }
    return {
      type: 'NIP',
      numberOnly: cleaned,
      formatted: `NIP. ${cleaned}`,
    };
  }

  // Standard NUPTK is 16 digits
  if (digitsOnly.length === 16) {
    return {
      type: 'NUPTK',
      numberOnly: trimmed,
      formatted: `NUPTK. ${trimmed}`,
    };
  }

  // Default / 18-digit PNS/PPPK is NIP
  return {
    type: 'NIP',
    numberOnly: trimmed,
    formatted: `NIP. ${trimmed}`,
  };
}

// ============================================================================
// 1. DOWNLOAD TEMPLATES (.XLSX)
// ============================================================================

export function downloadClassTemplate() {
  const wsData = [
    ['Tingkat', 'Nama Rombel', 'Konsentrasi Keahlian', 'Nama Wali Kelas', 'NIP / NUPTK Wali Kelas'],
    ['10', '10 TKR 1', 'Teknik Kendaraan Ringan', 'Imam Huzaer, S.Pd.', '199205112020121007'],
    ['11', '11 TKR 1', 'Teknik Kendaraan Ringan', 'Siti Aminah, M.Pd.', '198807202012022005'],
    ['12', '12 TKR 1', 'Teknik Kendaraan Ringan', 'Ahmad Dahlan, S.T.', '4533762663200012'],
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 10 }, { wch: 16 }, { wch: 28 }, { wch: 28 }, { wch: 24 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Template_Kelas');
  XLSX.writeFile(wb, 'Template_Data_Kelas_SMKN1.xlsx');
}

export function downloadStudentTemplate() {
  const wsData = [
    ['Nama Lengkap Siswa', 'NISN', 'NIS', 'Kelas', 'Konsentrasi Keahlian'],
    ['ADITYA PRATAMA', '0081234561', '20261001', '10 TKR 1', 'Teknik Kendaraan Ringan'],
    ['ANNISA RAHMAWATI', '0081234562', '20261002', '10 TKR 1', 'Teknik Kendaraan Ringan'],
    ['BAGAS SAPUTRA', '0081234563', '20261003', '10 TKR 1', 'Teknik Kendaraan Ringan'],
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 28 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 28 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Template_Siswa');
  XLSX.writeFile(wb, 'Template_Data_Siswa_SMKN1.xlsx');
}

export function downloadWaliKelasUserTemplate() {
  const wsData = [
    [
      'Nama Lengkap Wali Kelas',
      'NIP / NUPTK (18 digit NIP / 16 digit NUPTK)',
      'Username',
      'Password',
      'Wali Kelas Untuk (Nama Rombel)',
      'Merangkap Guru Mapel? (Ya/Tidak)',
    ],
    ['Imam Huzaer, S.Pd.', '199205112020121007', 'imam', '123456', '10 TKR 1', 'Ya'],
    ['Siti Aminah, M.Pd.', '198807202012022005', 'siti', '123456', '11 TKR 1', 'Tidak'],
    ['Ahmad Dahlan, S.T.', '4533762663200012', 'ahmad', '123456', '12 TKR 1', 'Ya'],
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [
    { wch: 28 },
    { wch: 36 },
    { wch: 16 },
    { wch: 14 },
    { wch: 26 },
    { wch: 28 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Template_Wali_Kelas');
  XLSX.writeFile(wb, 'Template_User_Wali_Kelas_SMKN1.xlsx');
}

export function downloadGuruMapelUserTemplate() {
  const wsData = [
    [
      'Nama Lengkap Guru Mapel',
      'NIP / NUPTK (18 digit NIP / 16 digit NUPTK)',
      'Username',
      'Password',
    ],
    ['Dra. Ratna Juwita', '197904152005012003', 'ratna', '123456'],
    ['Hendra Wijaya, S.Kom.', '4851764665110023', 'hendra', '123456'],
    ['Rina Marlina, S.Pd.', '199402102022212009', 'rina', '123456'],
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 28 }, { wch: 36 }, { wch: 18 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Template_Guru_Mapel');
  XLSX.writeFile(wb, 'Template_User_Guru_Mapel_SMKN1.xlsx');
}

export function downloadGradesTemplate(
  className: string,
  studentsInClass: StudentItem[],
  subjectName: string,
  allSubjectsForClass: string[],
  currentGrades: StudentGradesMap,
  mode: 'single_subject' | 'all_subjects'
) {
  const wb = XLSX.utils.book_new();

  if (mode === 'single_subject') {
    const wsData: (string | number)[][] = [
      ['No', 'NISN', 'NIS', 'Nama Siswa', 'Kelas', subjectName],
      ...studentsInClass.map((st, idx) => [
        idx + 1,
        st.nisn,
        st.nis,
        st.name,
        st.className,
        currentGrades[st.id]?.[subjectName] ?? 85,
      ]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 15 },
      { wch: 14 },
      { wch: 28 },
      { wch: 14 },
      { wch: 28 },
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Input_Nilai_Mapel');
    const safeSub = subjectName.replace(/[^a-zA-Z0-9]/g, '_').slice(0, 25);
    XLSX.writeFile(
      wb,
      `Template_Nilai_${className.replace(/\s+/g, '_')}_${safeSub}.xlsx`
    );
  } else {
    const headers = ['No', 'NISN', 'NIS', 'Nama Siswa', 'Kelas', ...allSubjectsForClass];
    const wsData: (string | number)[][] = [
      headers,
      ...studentsInClass.map((st, idx) => [
        idx + 1,
        st.nisn,
        st.nis,
        st.name,
        st.className,
        ...allSubjectsForClass.map((sub) => currentGrades[st.id]?.[sub] ?? ''),
      ]),
    ];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [
      { wch: 6 },
      { wch: 15 },
      { wch: 14 },
      { wch: 28 },
      { wch: 14 },
      ...allSubjectsForClass.map(() => ({ wch: 20 })),
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'Leger_Nilai_Kelas');
    XLSX.writeFile(
      wb,
      `Template_Leger_Nilai_${className.replace(/\s+/g, '_')}.xlsx`
    );
  }
}

/**
 * Common abbreviations for standard subjects to keep Leger table headers clean and readable.
 */
export function getSubjectAbbreviation(name: string): string {
  if (!name) return '-';
  const clean = name.trim();
  const lower = clean.toLowerCase();
  if (lower.includes('agama') || lower.includes('budi pekerti')) return 'PABP';
  if (lower.includes('pancasila') || lower.includes('kewarganegaraan')) return 'PPKn';
  if (lower.includes('bahasa indonesia')) return 'B. Indo';
  if (lower.includes('jasmani') || lower.includes('pjok') || lower.includes('olahraga')) return 'PJOK';
  if (lower.includes('sejarah')) return 'Sejarah';
  if (lower.includes('seni')) return 'Seni';
  if (lower.includes('lampung')) return 'B. Lamp';
  if (lower.includes('anti korupsi')) return 'PAK';
  if (lower.includes('matematika')) return 'MTK';
  if (lower.includes('bahasa inggris')) return 'B. Ing';
  if (lower.includes('informatika')) return 'Informatika';
  if (lower.includes('ipas') || lower.includes('ilmu pengetahuan alam')) return 'IPAS';
  if (lower.includes('dasar program keahlian') || lower.includes('dpk')) return 'DPK';
  if (lower.includes('kreativitas') || lower.includes('kik') || lower.includes('kewirausahaan')) return 'KIK';
  if (lower.includes('konsentrasi keahlian')) return 'Konsentrasi';
  if (lower.includes('pilihan')) return 'Pilihan';
  if (clean.length > 14) {
    const words = clean.split(/\s+/);
    if (words.length >= 2) {
      return words.map((w) => w[0]?.toUpperCase()).join('');
    }
    return clean.slice(0, 12);
  }
  return clean;
}

export function exportClassLegerToExcel(params: {
  school: SchoolData;
  classInfo: Partial<ClassItem>;
  students: RankedStudentItem[];
  allSubjects: Array<{ name: string; category: string }>;
  grades: StudentGradesMap;
}) {
  const { school, classInfo, students, allSubjects, grades } = params;
  const wb = XLSX.utils.book_new();

  const titleRows: (string | number)[][] = [
    ['PEMERINTAH PROVINSI LAMPUNG - DINAS PENDIDIKAN DAN KEBUDAYAAN'],
    [`${school.name.toUpperCase()} - TAHUN PELAJARAN ${school.tp}`],
    ['LEGER NILAI HASIL BELAJAR PESERTA DIDIK (SUMATIF TENGAH SEMESTER)'],
    [
      `Kelas: ${classInfo.name || '-'}`,
      `Tingkat: Kelas ${classInfo.grade || '-'}`,
      `Konsentrasi Keahlian: ${classInfo.major || '-'}`,
      `Wali Kelas: ${formatTeacherNameWithTitle(classInfo.waliKelasName)}`,
      `Semester: Ganjil`,
    ],
    [],
  ];

  const subjectNames = allSubjects.map((s) => s.name);
  const subjectHeaders = allSubjects.map(
    (s) => `${getSubjectAbbreviation(s.name)} [${s.category}]`
  );

  const tableHeaders = [
    'No',
    'NISN',
    'NIS',
    'Nama Peserta Didik',
    ...subjectHeaders,
    'Total Nilai',
    'Rata-Rata',
    'Peringkat',
    'Sakit (S)',
    'Izin (I)',
    'Alpa (A)',
    'Total Absen',
    'Catatan Wali Kelas',
  ];

  const studentDataRows: (string | number)[][] = students.map((st, idx) => {
    const stGrades = grades[st.id] || {};
    const subScores = subjectNames.map((sub) => {
      const v = stGrades[sub];
      return v !== undefined && v !== null && !isNaN(Number(v)) ? Number(v) : '';
    });
    const s = st.attendance?.sakit || 0;
    const i = st.attendance?.izin || 0;
    const a = st.attendance?.alpa || 0;
    const totalAbsen = s + i + a;

    return [
      idx + 1,
      st.nisn,
      st.nis,
      st.name,
      ...subScores,
      st.totalScore,
      st.avgScore,
      st.rank,
      s,
      i,
      a,
      totalAbsen,
      st.note || '',
    ];
  });

  // Calculate subject statistics
  const subjectAverages = subjectNames.map((sub) => {
    const validScores: number[] = [];
    students.forEach((st) => {
      const v = grades[st.id]?.[sub];
      if (v !== undefined && v !== null && !isNaN(Number(v)) && Number(v) > 0) {
        validScores.push(Number(v));
      }
    });
    if (validScores.length === 0) return 0;
    return Number((validScores.reduce((a, b) => a + b, 0) / validScores.length).toFixed(1));
  });

  const subjectMaxs = subjectNames.map((sub) => {
    let max = 0;
    students.forEach((st) => {
      const v = Number(grades[st.id]?.[sub] || 0);
      if (v > max) max = v;
    });
    return max;
  });

  const subjectMins = subjectNames.map((sub) => {
    const validScores: number[] = [];
    students.forEach((st) => {
      const v = grades[st.id]?.[sub];
      if (v !== undefined && v !== null && !isNaN(Number(v)) && Number(v) > 0) {
        validScores.push(Number(v));
      }
    });
    return validScores.length > 0 ? Math.min(...validScores) : 0;
  });

  const avgTotal =
    students.length > 0
      ? Number((students.reduce((a, b) => a + b.totalScore, 0) / students.length).toFixed(1))
      : 0;
  const avgClass =
    students.length > 0
      ? Number((students.reduce((a, b) => a + b.avgScore, 0) / students.length).toFixed(1))
      : 0;

  const statAvgRow: (string | number)[] = [
    '',
    '',
    '',
    'RATA-RATA KELAS',
    ...subjectAverages,
    avgTotal,
    avgClass,
    '',
    '',
    '',
    '',
    '',
    '',
  ];

  const statMaxRow: (string | number)[] = [
    '',
    '',
    '',
    'NILAI TERTINGGI',
    ...subjectMaxs,
    Math.max(...students.map((s) => s.totalScore), 0),
    Math.max(...students.map((s) => s.avgScore), 0),
    '',
    '',
    '',
    '',
    '',
    '',
  ];

  const statMinRow: (string | number)[] = [
    '',
    '',
    '',
    'NILAI TERENDAH',
    ...subjectMins,
    Math.min(...students.map((s) => s.totalScore), 0),
    Math.min(...students.map((s) => s.avgScore), 0),
    '',
    '',
    '',
    '',
    '',
    '',
  ];

  const signatureRows: (string | number)[][] = [
    [],
    [],
    ['', '', '', '', '', '', '', '', '', '', `${school.datePrint}`],
    ['', '', '', 'Mengetahui,', '', '', '', '', '', '', 'Wali Kelas,'],
    ['', '', '', 'Kepala Sekolah', '', '', '', '', '', '', ''],
    [],
    [],
    [],
    [
      '',
      '',
      '',
      formatTeacherNameWithTitle(school.principal),
      '',
      '',
      '',
      '',
      '',
      '',
      formatTeacherNameWithTitle(classInfo.waliKelasName),
    ],
    [
      '',
      '',
      '',
      formatNipOrNuptk(school.principalNip).formatted || '-',
      '',
      '',
      '',
      '',
      '',
      '',
      formatNipOrNuptk(classInfo.waliKelasNip).formatted || '-',
    ],
    [],
    ['© 2026 E-Raport STS SMK Negeri 1 Mesuji Timur - Desain & Sistem Dibuat oleh Imam Huzaer, S.Pd.'],
  ];

  const fullSheetData: (string | number)[][] = [
    ...titleRows,
    tableHeaders,
    ...studentDataRows,
    statAvgRow,
    statMaxRow,
    statMinRow,
    ...signatureRows,
  ];

  const ws = XLSX.utils.aoa_to_sheet(fullSheetData);
  ws['!cols'] = [
    { wch: 5 },
    { wch: 14 },
    { wch: 12 },
    { wch: 28 },
    ...subjectNames.map(() => ({ wch: 14 })),
    { wch: 12 },
    { wch: 12 },
    { wch: 10 },
    { wch: 8 },
    { wch: 8 },
    { wch: 8 },
    { wch: 11 },
    { wch: 40 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, `Leger_${(classInfo.name || 'Kelas').replace(/\s+/g, '_')}`);
  const cleanClassName = (classInfo.name || 'Kelas').replace(/[^a-zA-Z0-9]/g, '_');
  XLSX.writeFile(wb, `Leger_Nilai_${cleanClassName}_STS_${school.tp.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`);
}

export function exportClassLegerToCsv(params: {
  school: SchoolData;
  classInfo: Partial<ClassItem>;
  students: RankedStudentItem[];
  allSubjects: Array<{ name: string; category: string }>;
  grades: StudentGradesMap;
}) {
  const { school, classInfo, students, allSubjects, grades } = params;
  const subjectNames = allSubjects.map((s) => s.name);
  const subjectHeaders = allSubjects.map(
    (s) => `"${getSubjectAbbreviation(s.name)}"`
  );

  const lines: string[] = [
    `"LEGER NILAI HASIL BELAJAR SISWA - ${school.name}"`,
    `"Kelas: ${classInfo.name || '-'}","Tingkat: ${classInfo.grade || '-'}","Konsentrasi: ${classInfo.major || '-'}","Wali Kelas: ${classInfo.waliKelasName || '-'}","TP: ${school.tp}"`,
    '',
    [
      '"No"',
      '"NISN"',
      '"NIS"',
      '"Nama Siswa"',
      ...subjectHeaders,
      '"Total Nilai"',
      '"Rata-Rata"',
      '"Peringkat"',
      '"Sakit"',
      '"Izin"',
      '"Alpa"',
      '"Catatan"',
    ].join(','),
  ];

  students.forEach((st, idx) => {
    const stGrades = grades[st.id] || {};
    const subScores = subjectNames.map((sub) => {
      const v = stGrades[sub];
      return v !== undefined && v !== null && !isNaN(Number(v)) ? Number(v) : '';
    });
    lines.push(
      [
        idx + 1,
        `"${st.nisn}"`,
        `"${st.nis}"`,
        `"${st.name.replace(/"/g, '""')}"`,
        ...subScores,
        st.totalScore,
        st.avgScore,
        st.rank,
        st.attendance?.sakit || 0,
        st.attendance?.izin || 0,
        st.attendance?.alpa || 0,
        `"${(st.note || '').replace(/"/g, '""')}"`,
      ].join(',')
    );
  });

  const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Leger_Nilai_${(classInfo.name || 'Kelas').replace(/\s+/g, '_')}_STS.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export function triggerPrintClassLeger(): void {
  const styleEl = document.createElement('style');
  styleEl.id = 'leger-landscape-print-style';
  styleEl.innerHTML = `@page { size: A4 landscape !important; margin: 6mm !important; }`;
  document.head.appendChild(styleEl);
  document.body.classList.add('printing-leger');

  const cleanup = () => {
    document.body.classList.remove('printing-leger');
    const existing = document.getElementById('leger-landscape-print-style');
    if (existing) existing.remove();
    window.removeEventListener('afterprint', cleanup);
  };

  window.addEventListener('afterprint', cleanup);
  setTimeout(() => {
    window.print();
    setTimeout(cleanup, 2000);
  }, 120);
}

// ============================================================================
// 2. PARSE EXCEL / CSV FILE OR CLIPBOARD TEXT
// ============================================================================

export async function readSpreadsheetFileToRows(file: File): Promise<string[][]> {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array' });
  const firstSheetName = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheetName];
  const rawRows = XLSX.utils.sheet_to_json<any[]>(ws, {
    header: 1,
    defval: '',
    raw: true,
  });
  const fmtRows = XLSX.utils.sheet_to_json<any[]>(ws, {
    header: 1,
    defval: '',
    raw: false,
  });
  const rowCount = Math.max(rawRows.length, fmtRows.length);
  const out: string[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const rawRow = Array.isArray(rawRows[r]) ? rawRows[r] : [];
    const fmtRow = Array.isArray(fmtRows[r]) ? fmtRows[r] : [];
    const colCount = Math.max(rawRow.length, fmtRow.length);
    const rowCells: string[] = [];
    for (let c = 0; c < colCount; c++) {
      const rawCell = rawRow[c];
      const fmtCell = String(fmtRow[c] ?? '').trim();
      if (
        typeof rawCell === 'number' &&
        Number.isFinite(rawCell) &&
        (/^\d+(\.\d+)?e[+-]?\d+$/i.test(fmtCell) || Math.abs(rawCell) >= 1e11)
      ) {
        rowCells.push(
          Math.round(rawCell).toLocaleString('fullwide', {
            useGrouping: false,
          })
        );
      } else {
        rowCells.push(fmtCell || String(rawCell ?? '').trim());
      }
    }
    if (rowCells.some((cell) => cell !== '')) {
      out.push(rowCells);
    }
  }
  return out;
}

export function parsePastedTextToRows(text: string): string[][] {
  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  return lines.map((line) => {
    if (line.includes('\t')) {
      return line.split('\t').map((c) => c.trim());
    }
    if (line.includes(';')) {
      return line.split(';').map((c) => c.trim());
    }
    return line.split(',').map((c) => c.trim());
  });
}

function isHeaderRow(firstRow: string[], keywords: string[]): boolean {
  const joined = firstRow.join(' ').toLowerCase();
  return keywords.some((k) => joined.includes(k));
}

function createCompactBaseId(): number {
  return (Date.now() % 800000) + 100000;
}

export function parseBulkClasses(rows: string[][]): ClassItem[] {
  if (rows.length === 0) return [];
  const dataRows = isHeaderRow(rows[0], ['tingkat', 'rombel', 'wali', 'konsentrasi'])
    ? rows.slice(1)
    : rows;

  const baseId = createCompactBaseId();
  return dataRows
    .filter((r) => r.length >= 2 && (r[0] || r[1]))
    .map((r, idx) => {
      let grade = r[0] || '10';
      let name = r[1] || '';
      if (grade.length > 2 && ['10', '11', '12'].includes(r[1])) {
        name = r[0];
        grade = r[1];
      } else if (!['10', '11', '12'].includes(grade)) {
        if (name.startsWith('12') || grade.startsWith('12')) grade = '12';
        else if (name.startsWith('11') || grade.startsWith('11')) grade = '11';
        else grade = '10';
      }
      return {
        id: baseId + idx,
        grade,
        name: name || r[0],
        major: r[2] || 'Teknik Kendaraan Ringan',
        waliKelasName: r[3] || '-',
        waliKelasNip: r[4] || '-',
      };
    });
}

export function parseBulkStudents(
  rows: string[][],
  existingClasses: ClassItem[]
): StudentItem[] {
  if (rows.length === 0) return [];
  const dataRows = isHeaderRow(rows[0], ['nama', 'nisn', 'nis', 'kelas'])
    ? rows.slice(1)
    : rows;

  const baseId = createCompactBaseId();
  return dataRows
    .filter((r) => r.length >= 1 && r[0])
    .map((r, idx) => {
      const name = (r[0] || '').toUpperCase();
      const nisn = r[1] || `008${Math.floor(1000000 + Math.random() * 9000000)}`;
      const nis = r[2] || `2026${Math.floor(1000 + Math.random() * 9000)}`;
      const className = r[3] || existingClasses[0]?.name || '10 TKR 1';
      const matchedClass = existingClasses.find(
        (c) => c.name.trim().toLowerCase() === className.trim().toLowerCase()
      );
      const grade =
        matchedClass?.grade ||
        (className.startsWith('12')
          ? '12'
          : className.startsWith('11')
          ? '11'
          : '10');
      const major = r[4] || matchedClass?.major || 'Teknik Kendaraan Ringan';

      return {
        id: baseId + idx,
        name,
        nisn,
        nis,
        className: matchedClass ? matchedClass.name : className,
        grade,
        major,
        attendance: { sakit: 0, izin: 0, alpa: 0 },
        note: 'Tingkatkan terus motivasi dan kedisiplinan belajar.',
      };
    });
}

export function parseBulkWaliKelasUsers(rows: string[][]): UserAccount[] {
  if (rows.length === 0) return [];
  const dataRows = isHeaderRow(rows[0], ['nama', 'username', 'password', 'nip', 'nuptk', 'wali'])
    ? rows.slice(1)
    : rows;

  const baseId = createCompactBaseId();
  return dataRows
    .filter((r) => r.length >= 2 && r[0])
    .map((r, idx) => {
      const name = r[0];
      const nip = r[1] || '-';
      const username = (
        r[2] ||
        name
          .split(',')[0]
          .trim()
          .split(/\s+/)[0]
          .toLowerCase() +
          (idx + 1)
      )
        .toLowerCase()
        .replace(/\s+/g, '');
      const password = r[3] || '123456';
      const waliKelasFor = r[4] || '10 TKR 1';
      const merangkapStr = (r[5] || '').toLowerCase();
      const roles: RoleType[] = ['walikelas'];
      if (
        merangkapStr.includes('ya') ||
        merangkapStr.includes('guru') ||
        merangkapStr.includes('true') ||
        merangkapStr === '1'
      ) {
        roles.push('gurumapel');
      }

      return {
        id: baseId + idx,
        name,
        nip,
        username,
        password,
        roles,
        waliKelasFor,
      };
    });
}

export function parseBulkGuruMapelUsers(rows: string[][]): UserAccount[] {
  if (rows.length === 0) return [];
  const dataRows = isHeaderRow(rows[0], ['nama', 'username', 'password', 'nip', 'nuptk'])
    ? rows.slice(1)
    : rows;

  const baseId = createCompactBaseId();
  return dataRows
    .filter((r) => r.length >= 1 && r[0])
    .map((r, idx) => {
      const name = r[0];
      const nip = r[1] || '-';
      const username = (
        r[2] ||
        name
          .split(',')[0]
          .trim()
          .split(/\s+/)[0]
          .toLowerCase() +
          (idx + 1)
      )
        .toLowerCase()
        .replace(/\s+/g, '');
      const password = r[3] || '123456';

      return {
        id: baseId + idx,
        name,
        nip,
        username,
        password,
        roles: ['gurumapel'],
        waliKelasFor: '',
      };
    });
}

export function parseBulkGrades(
  rows: string[][],
  studentsInClass: StudentItem[],
  selectedSubject: string,
  allSubjects: string[],
  mode: 'single_subject' | 'all_subjects' = 'single_subject'
): { updatedCount: number; gradeUpdates: Record<number, Record<string, number>> } {
  if (rows.length === 0) return { updatedCount: 0, gradeUpdates: {} };

  const firstRow = rows[0];
  const hasHeader = isHeaderRow(firstRow, [
    'nama',
    'nisn',
    'nis',
    'nilai',
    ...allSubjects.map((s) => s.toLowerCase()),
  ]);

  const gradeUpdates: Record<number, Record<string, number>> = {};
  let updatedCount = 0;

  if (hasHeader) {
    const headers = firstRow.map((h) => h.trim());
    const lowerHeaders = headers.map((h) => h.toLowerCase());

    const idxNisn = lowerHeaders.findIndex((h) => h === 'nisn');
    const idxNis = lowerHeaders.findIndex((h) => h === 'nis');
    const idxName = lowerHeaders.findIndex(
      (h) => h.includes('nama') || h.includes('siswa')
    );

    const subjectColumns: Array<{ colIdx: number; subjectName: string }> = [];

    if (mode === 'single_subject') {
      // Strictly map only to selectedSubject for Guru Mapel
      const exactSubCol = headers.findIndex(
        (h) => h.toLowerCase() === selectedSubject.toLowerCase()
      );
      const idxNilai = lowerHeaders.findIndex((h) => h.includes('nilai'));
      const targetCol =
        exactSubCol !== -1
          ? exactSubCol
          : idxNilai !== -1
          ? idxNilai
          : firstRow.length - 1;
      subjectColumns.push({ colIdx: targetCol, subjectName: selectedSubject });
    } else {
      // Admin 'all_subjects' mode: map all matching subject columns
      headers.forEach((h, cIdx) => {
        const matchedSub = allSubjects.find(
          (sub) => sub.toLowerCase() === h.toLowerCase()
        );
        if (matchedSub) {
          subjectColumns.push({ colIdx: cIdx, subjectName: matchedSub });
        }
      });

      // Fallback if headers didn't match exact subject names: map columns after Name/NIS sequentially to allSubjects
      if (subjectColumns.length === 0) {
        const startCol = Math.max(idxName, idxNis, idxNisn, 0) + 1;
        allSubjects.forEach((subName, idx) => {
          if (startCol + idx < firstRow.length) {
            subjectColumns.push({
              colIdx: startCol + idx,
              subjectName: subName,
            });
          }
        });
      }
    }

    rows.slice(1).forEach((r) => {
      const matchedStudent = studentsInClass.find((st) => {
        if (idxNisn !== -1 && r[idxNisn] && st.nisn === r[idxNisn].trim()) {
          return true;
        }
        if (idxNis !== -1 && r[idxNis] && st.nis === r[idxNis].trim()) {
          return true;
        }
        if (
          idxName !== -1 &&
          r[idxName] &&
          st.name.trim().toLowerCase() === r[idxName].trim().toLowerCase()
        ) {
          return true;
        }
        return false;
      });

      if (matchedStudent) {
        if (!gradeUpdates[matchedStudent.id]) {
          gradeUpdates[matchedStudent.id] = {};
        }
        let studentChanged = false;
        subjectColumns.forEach(({ colIdx, subjectName }) => {
          const rawScore = r[colIdx];
          if (rawScore !== undefined && rawScore.trim() !== '') {
            const num = Math.max(0, Math.min(100, Number(rawScore)));
            if (!Number.isNaN(num)) {
              gradeUpdates[matchedStudent.id][subjectName] = num;
              studentChanged = true;
            }
          }
        });
        if (studentChanged) updatedCount++;
      }
    });
  } else {
    rows.forEach((r, rowIdx) => {
      let matchedStudent = studentsInClass.find(
        (st) =>
          r.some((cell) => cell.trim() === st.nisn) ||
          r.some(
            (cell) => cell.trim().toLowerCase() === st.name.trim().toLowerCase()
          )
      );
      if (!matchedStudent && rowIdx < studentsInClass.length) {
        matchedStudent = studentsInClass[rowIdx];
      }

      if (matchedStudent) {
        if (mode === 'single_subject') {
          const lastCell = r[r.length - 1];
          const num = Math.max(0, Math.min(100, Number(lastCell)));
          if (!Number.isNaN(num)) {
            if (!gradeUpdates[matchedStudent.id]) {
              gradeUpdates[matchedStudent.id] = {};
            }
            gradeUpdates[matchedStudent.id][selectedSubject] = num;
            updatedCount++;
          }
        } else {
          // Admin 'all_subjects' mode without header: find student name/nisn column or start from col 0
          const nameOrIdCol = r.findIndex(
            (cell) =>
              cell.trim() === matchedStudent!.nisn ||
              cell.trim().toLowerCase() ===
                matchedStudent!.name.trim().toLowerCase()
          );
          const startScoreCol = nameOrIdCol !== -1 ? nameOrIdCol + 1 : 0;
          let studentChanged = false;
          if (!gradeUpdates[matchedStudent.id]) {
            gradeUpdates[matchedStudent.id] = {};
          }
          allSubjects.forEach((subName, sIdx) => {
            const rawVal = r[startScoreCol + sIdx];
            if (rawVal !== undefined && rawVal.trim() !== '') {
              const num = Math.max(0, Math.min(100, Number(rawVal)));
              if (!Number.isNaN(num)) {
                gradeUpdates[matchedStudent!.id][subName] = num;
                studentChanged = true;
              }
            }
          });
          if (studentChanged) updatedCount++;
        }
      }
    });
  }

  return { updatedCount, gradeUpdates };
}

// ============================================================================
// 3. KOP IMAGE TRIMMER & DIRECT A4 PDF GENERATOR
// ============================================================================

const trimmedKopCache = new Map<string, string>();

/**
 * Loads an image Blob into an HTMLImageElement, automatically crops any left/right/top/bottom
 * white or transparent padding so the Kop Surat aligns 100% edge-to-edge with the raport table,
 * and returns a CORS-safe base64 Data URL for both preview and PDF export.
 */
async function cropKopBlobToDataUrl(blob: Blob): Promise<string | null> {
  return new Promise((resolve) => {
    const objectUrl = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      try {
        const w = img.naturalWidth || img.width;
        const h = img.naturalHeight || img.height;
        if (!w || !h) {
          URL.revokeObjectURL(objectUrl);
          resolve(null);
          return;
        }

        const srcCanvas = document.createElement('canvas');
        srcCanvas.width = w;
        srcCanvas.height = h;
        const ctx = srcCanvas.getContext('2d');
        if (!ctx) {
          URL.revokeObjectURL(objectUrl);
          resolve(null);
          return;
        }

        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);

        const imgData = ctx.getImageData(0, 0, w, h).data;
        let minX = w;
        let maxX = 0;
        let minY = h;
        let maxY = 0;

        for (let y = 0; y < h; y++) {
          for (let x = 0; x < w; x++) {
            const i = (y * w + x) * 4;
            const r = imgData[i];
            const g = imgData[i + 1];
            const b = imgData[i + 2];
            const a = imgData[i + 3];
            // Non-white pixel threshold
            if (a > 25 && (r < 244 || g < 244 || b < 244)) {
              if (x < minX) minX = x;
              if (x > maxX) maxX = x;
              if (y < minY) minY = y;
              if (y > maxY) maxY = y;
            }
          }
        }

        URL.revokeObjectURL(objectUrl);

        if (maxX > minX && maxY > minY) {
          // Keep exact horizontal bounds (minX..maxX) so Kop is 100% flush with table left & right
          const cropX = Math.max(0, minX);
          const cropY = Math.max(0, minY - 2);
          const cropW = Math.min(w - cropX, maxX - cropX + 1);
          const cropH = Math.min(h - cropY, maxY - cropY + 3);

          const outCanvas = document.createElement('canvas');
          outCanvas.width = cropW;
          outCanvas.height = cropH;
          const outCtx = outCanvas.getContext('2d');
          if (outCtx) {
            outCtx.fillStyle = '#ffffff';
            outCtx.fillRect(0, 0, cropW, cropH);
            outCtx.drawImage(
              srcCanvas,
              cropX,
              cropY,
              cropW,
              cropH,
              0,
              0,
              cropW,
              cropH
            );
            resolve(outCanvas.toDataURL('image/png'));
            return;
          }
        }

        resolve(srcCanvas.toDataURL('image/png'));
      } catch {
        URL.revokeObjectURL(objectUrl);
        resolve(null);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(null);
    };
    img.src = objectUrl;
  });
}

export async function resolveAndTrimKopImage(
  url?: string
): Promise<string | null> {
  if (!url || !url.trim()) return null;
  const cleanUrl = url.trim();
  if (cleanUrl.startsWith('data:')) return cleanUrl;
  if (trimmedKopCache.has(cleanUrl)) {
    return trimmedKopCache.get(cleanUrl)!;
  }

  const candidateUrls = [
    cleanUrl,
    `https://wsrv.nl/?url=${encodeURIComponent(cleanUrl)}&output=png`,
    `https://images.weserv.nl/?url=${encodeURIComponent(cleanUrl)}&output=png`,
    `https://corsproxy.io/?${encodeURIComponent(cleanUrl)}`,
    `https://api.allorigins.win/raw?url=${encodeURIComponent(cleanUrl)}`,
  ];

  for (const candidate of candidateUrls) {
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 5500);
      const res = await fetch(candidate, {
        mode: 'cors',
        signal: controller.signal,
      });
      clearTimeout(timeoutId);
      if (!res.ok) continue;
      const blob = await res.blob();
      if (blob.size < 200) continue;

      const dataUrl = await cropKopBlobToDataUrl(blob);
      if (dataUrl) {
        trimmedKopCache.set(cleanUrl, dataUrl);
        return dataUrl;
      }
    } catch {
      // Try next candidate
    }
  }

  return null;
}

export async function processUploadedSignatureOrQrImage(
  file: File,
  mode: 'signature' | 'qrcode',
  removeWhiteBg: boolean = mode === 'signature'
): Promise<string> {
  return new Promise((resolve, reject) => {
    const objectUrl = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      try {
        const rawW = img.naturalWidth || img.width;
        const rawH = img.naturalHeight || img.height;
        if (!rawW || !rawH) {
          URL.revokeObjectURL(objectUrl);
          reject(new Error('Ukuran gambar tidak valid.'));
          return;
        }

        const maxW = mode === 'qrcode' ? 200 : 340;
        const maxH = mode === 'qrcode' ? 200 : 150;
        const scale = Math.min(1, maxW / rawW, maxH / rawH);
        const targetW = Math.max(1, Math.round(rawW * scale));
        const targetH = Math.max(1, Math.round(rawH * scale));

        const canvas = document.createElement('canvas');
        canvas.width = targetW;
        canvas.height = targetH;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          URL.revokeObjectURL(objectUrl);
          reject(new Error('Gagal memproses canvas gambar.'));
          return;
        }

        ctx.clearRect(0, 0, targetW, targetH);
        ctx.drawImage(img, 0, 0, targetW, targetH);
        URL.revokeObjectURL(objectUrl);

        if (removeWhiteBg && mode === 'signature') {
          const imageData = ctx.getImageData(0, 0, targetW, targetH);
          const data = imageData.data;
          for (let i = 0; i < data.length; i += 4) {
            const r = data[i];
            const g = data[i + 1];
            const b = data[i + 2];
            const a = data[i + 3];
            if (a > 0) {
              const brightness = (r + g + b) / 3;
              if (r > 230 && g > 230 && b > 230) {
                data[i + 3] = 0;
              } else if (brightness > 210) {
                // Smooth anti-aliased edge transition
                data[i + 3] = Math.round(((240 - brightness) / 30) * a);
              }
            }
          }
          ctx.putImageData(imageData, 0, 0);
        }

        let dataUrl = canvas.toDataURL('image/png');
        // Ensure compact size (< 42,000 chars) so it always fits inside a single Google Sheets cell
        if (dataUrl.length > 42000) {
          const smallerScale = 0.72;
          const smallCanvas = document.createElement('canvas');
          smallCanvas.width = Math.max(1, Math.round(targetW * smallerScale));
          smallCanvas.height = Math.max(1, Math.round(targetH * smallerScale));
          const sCtx = smallCanvas.getContext('2d');
          if (sCtx) {
            sCtx.drawImage(
              canvas,
              0,
              0,
              smallCanvas.width,
              smallCanvas.height
            );
            dataUrl = smallCanvas.toDataURL('image/png');
          }
        }

        resolve(dataUrl);
      } catch (err) {
        URL.revokeObjectURL(objectUrl);
        reject(err);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error('Gagal membaca file gambar.'));
    };
    img.src = objectUrl;
  });
}

export async function removeWhiteBackgroundFromDataUrl(
  srcDataUrl: string
): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      try {
        const w = img.naturalWidth || img.width;
        const h = img.naturalHeight || img.height;
        const canvas = document.createElement('canvas');
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          resolve(srcDataUrl);
          return;
        }
        ctx.clearRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        const imageData = ctx.getImageData(0, 0, w, h);
        const data = imageData.data;
        for (let i = 0; i < data.length; i += 4) {
          const r = data[i];
          const g = data[i + 1];
          const b = data[i + 2];
          const a = data[i + 3];
          if (a > 0) {
            const brightness = (r + g + b) / 3;
            if (r > 228 && g > 228 && b > 228) {
              data[i + 3] = 0;
            } else if (brightness > 208) {
              data[i + 3] = Math.round(((238 - brightness) / 30) * a);
            }
          }
        }
        ctx.putImageData(imageData, 0, 0);
        resolve(canvas.toDataURL('image/png'));
      } catch (e) {
        reject(e);
      }
    };
    img.onerror = () => reject(new Error('Gagal memuat gambar untuk diproses.'));
    img.src = srcDataUrl;
  });
}

export async function generateAndDownloadPdf(
  containerElement: HTMLElement,
  filename: string
): Promise<void> {
  const pageElements = Array.from(
    containerElement.querySelectorAll<HTMLElement>('.raport-preview')
  );
  if (pageElements.length === 0) {
    throw new Error('Tidak ditemukan halaman raport untuk diunduh.');
  }

  // Inject temporary style to guarantee html2canvas font baseline accuracy
  const fixStyle = document.createElement('style');
  fixStyle.id = 'h2c-baseline-fix-style';
  fixStyle.textContent = `
    img:not(.kop-img):not(.sig-qr-img) {
      display: inline !important;
      vertical-align: baseline !important;
    }
    .kop-img {
      display: block !important;
      width: 100% !important;
      max-width: 100% !important;
      height: auto !important;
      object-fit: fill !important;
    }
    .sig-qr-img {
      display: block !important;
      margin: 0 auto !important;
      object-fit: contain !important;
    }
  `;
  document.head.appendChild(fixStyle);

  // Create an unclipped offscreen host directly on document.body (avoids modal scroll clipping)
  const offscreenHost = document.createElement('div');
  offscreenHost.style.position = 'fixed';
  offscreenHost.style.left = '-10000px';
  offscreenHost.style.top = '0';
  offscreenHost.style.width = '794px'; // 210mm at 96dpi
  offscreenHost.style.background = '#ffffff';
  offscreenHost.style.zIndex = '-9999';
  offscreenHost.style.overflow = 'visible';
  document.body.appendChild(offscreenHost);

  try {
    const pdf = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4',
      compress: true,
    });

    for (let i = 0; i < pageElements.length; i++) {
      const sourceEl = pageElements[i];
      offscreenHost.innerHTML = '';

      const clonedEl = sourceEl.cloneNode(true) as HTMLElement;
      clonedEl.style.width = '794px';
      clonedEl.style.minHeight = '1123px'; // 297mm at 96dpi
      clonedEl.style.height = 'auto';
      clonedEl.style.margin = '0';
      clonedEl.style.boxShadow = 'none';
      clonedEl.style.overflow = 'visible';
      clonedEl.style.background = '#ffffff';
      offscreenHost.appendChild(clonedEl);

      // Ensure any .kop-img inside clonedEl uses a CORS-safe data URL
      const kopImgs = Array.from(
        clonedEl.querySelectorAll<HTMLImageElement>('img.kop-img')
      );
      for (const kopImg of kopImgs) {
        const rawSrc = kopImg.getAttribute('src') || '';
        if (rawSrc && !rawSrc.startsWith('data:')) {
          const resolvedDataUrl = await resolveAndTrimKopImage(rawSrc);
          if (resolvedDataUrl) {
            kopImg.src = resolvedDataUrl;
          }
        }
        // Wait for image decode if data URL
        if (kopImg.src.startsWith('data:') && !kopImg.complete) {
          await new Promise<void>((res) => {
            kopImg.onload = () => res();
            kopImg.onerror = () => res();
          });
        }
      }

      // Ensure any signature / QR code images (.sig-qr-img) are loaded
      const sigImgs = Array.from(
        clonedEl.querySelectorAll<HTMLImageElement>('img.sig-qr-img')
      );
      for (const sigImg of sigImgs) {
        if (!sigImg.complete) {
          await new Promise<void>((res) => {
            sigImg.onload = () => res();
            sigImg.onerror = () => res();
          });
        }
      }

      const fullWidth = 794;
      const fullHeight = Math.max(1123, clonedEl.scrollHeight);

      let canvas: HTMLCanvasElement;
      try {
        canvas = await html2canvas(clonedEl, {
          scale: 2,
          useCORS: true,
          allowTaint: false,
          backgroundColor: '#ffffff',
          logging: false,
          width: fullWidth,
          height: fullHeight,
          windowWidth: fullWidth,
          windowHeight: fullHeight,
          x: 0,
          y: 0,
          scrollX: 0,
          scrollY: 0,
        });
      } catch {
        // Fallback if external image could not be proxied: show official text kop instead of blank space
        kopImgs.forEach((kopImg) => {
          const fallbackHeader = clonedEl.querySelector<HTMLElement>(
            '.kop-fallback-header'
          );
          if (fallbackHeader) {
            fallbackHeader.style.display = 'block';
          }
          kopImg.remove();
        });
        const retryHeight = Math.max(1123, clonedEl.scrollHeight);
        canvas = await html2canvas(clonedEl, {
          scale: 2,
          useCORS: false,
          allowTaint: false,
          backgroundColor: '#ffffff',
          logging: false,
          width: fullWidth,
          height: retryHeight,
          windowWidth: fullWidth,
          windowHeight: retryHeight,
          x: 0,
          y: 0,
          scrollX: 0,
          scrollY: 0,
        });
      }

      const imgData = canvas.toDataURL('image/jpeg', 0.96);
      const pageWidthMm = 210;
      const pageHeightMm = 297;

      // Preserve exact aspect ratio and ensure bottom signature is never clipped
      const naturalHeightMm = (canvas.height * pageWidthMm) / canvas.width;
      let renderWidthMm = pageWidthMm;
      let renderHeightMm = naturalHeightMm;
      let offsetXMm = 0;
      let offsetYMm = 0;

      if (naturalHeightMm > pageHeightMm) {
        renderHeightMm = pageHeightMm - 2; // 295mm
        renderWidthMm = (canvas.width * renderHeightMm) / canvas.height;
        offsetXMm = (pageWidthMm - renderWidthMm) / 2;
        offsetYMm = 1;
      }

      if (i > 0) {
        pdf.addPage('a4', 'portrait');
      }
      pdf.addImage(
        imgData,
        'JPEG',
        offsetXMm,
        offsetYMm,
        renderWidthMm,
        renderHeightMm,
        undefined,
        'FAST'
      );
    }

    pdf.save(filename.endsWith('.pdf') ? filename : `${filename}.pdf`);
  } finally {
    offscreenHost.remove();
    fixStyle.remove();
  }
}


