export type RoleType = 'admin' | 'walikelas' | 'gurumapel';

export type SignatureDisplayType = 'none' | 'signature' | 'qrcode';

export interface UserAccount {
  id: number;
  username: string;
  password: string;
  name: string;
  nip: string;
  roles: RoleType[];
  waliKelasFor: string;
  waliSignatureType?: SignatureDisplayType;
  waliSignatureUrl?: string;
}

export interface SchoolData {
  name: string;
  npsn: string;
  logoUrl: string;
  kopUrl: string;
  principal: string;
  principalNip: string;
  tp: string;
  datePrint: string;
  principalSignatureType?: SignatureDisplayType;
  principalSignatureUrl?: string;
  gradeInputOpen?: boolean;
}

export interface ClassItem {
  id: number;
  name: string;
  grade: string;
  major: string;
  waliKelasName: string;
  waliKelasNip: string;
  waliSignatureType?: SignatureDisplayType;
  waliSignatureUrl?: string;
}

export interface SubjectItem {
  id: number;
  name: string;
}

export interface GradeSubjectTemplate {
  umum: SubjectItem[];
  kejuruan: SubjectItem[];
}

export type SubjectTemplatesMap = Record<string, GradeSubjectTemplate>;

export interface AttendanceRecord {
  sakit: number;
  izin: number;
  alpa: number;
}

export interface StudentItem {
  id: number;
  nis: string;
  nisn: string;
  name: string;
  className: string;
  grade: string;
  major: string;
  attendance: AttendanceRecord;
  note: string;
}

export interface RankedStudentItem extends StudentItem {
  totalScore: number;
  avgScore: number;
  rank: number;
}

export type StudentGradesMap = Record<number, Record<string, number>>;

export interface DatabasePayload {
  school: SchoolData;
  classes: ClassItem[];
  students: StudentItem[];
  users: UserAccount[];
  studentGrades: StudentGradesMap;
}

export const SPREADSHEET_ID = '13V9knEZFWKpGUWGD9EBClAdollTV7YQHT0AW6Oq3ahU';
export const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/13V9knEZFWKpGUWGD9EBClAdollTV7YQHT0AW6Oq3ahU/edit?gid=0#gid=0';

export const DEFAULT_SCHOOL: SchoolData = {
  name: 'SMK NEGERI 1 MESUJI TIMUR',
  npsn: '10810982',
  logoUrl: 'https://ihuzda.site.je/urlgambar/uploads/img2.png',
  kopUrl: 'https://ihuzda.site.je/urlgambar/uploads/img1.png',
  principal: 'TRI YULINAWATI, S.T.',
  principalNip: '19740728 200902 2 003',
  tp: '2026/2027',
  datePrint: 'Mesuji Timur, Oktober 2026',
  principalSignatureType: 'none',
  principalSignatureUrl: '',
  gradeInputOpen: true,
};

export const DEFAULT_CLASSES: ClassItem[] = [
  { id: 1, name: '10 TKR 1', grade: '10', major: 'Teknik Kendaraan Ringan', waliKelasName: 'Budi Santoso, S.Pd.', waliKelasNip: '198503122010011002', waliSignatureType: 'none', waliSignatureUrl: '' },
  { id: 2, name: '11 TKR 1', grade: '11', major: 'Teknik Kendaraan Ringan', waliKelasName: 'Siti Aminah, M.Pd.', waliKelasNip: '198807202012022005', waliSignatureType: 'none', waliSignatureUrl: '' },
  { id: 3, name: '12 TKR 1', grade: '12', major: 'Teknik Kendaraan Ringan', waliKelasName: 'Ahmad Dahlan, S.T.', waliKelasNip: '198211052008011004', waliSignatureType: 'none', waliSignatureUrl: '' },
];

export const DEFAULT_USERS: UserAccount[] = [
  { id: 1, username: 'admin', password: 'admin123', name: 'Administrator Raport', nip: '197407282009022003', roles: ['admin'], waliKelasFor: '' },
  { id: 2, username: 'budi', password: '123', name: 'Budi Santoso, S.Pd.', nip: '198503122010011002', roles: ['walikelas', 'gurumapel'], waliKelasFor: '10 TKR 1' },
  { id: 3, username: 'siti', password: '123', name: 'Siti Aminah, M.Pd.', nip: '198807202012022005', roles: ['walikelas'], waliKelasFor: '11 TKR 1' },
  { id: 4, username: 'ratna', password: '123', name: 'Dra. Ratna Juwita', nip: '197904152005012003', roles: ['gurumapel'], waliKelasFor: '' },
];

export const DEFAULT_SUBJECT_TEMPLATES: SubjectTemplatesMap = {
  '10': {
    umum: [
      { id: 1, name: 'Pendidikan Agama dan Budi Pekerti' },
      { id: 2, name: 'Pendidikan Pancasila dan Kewarganegaraan' },
      { id: 3, name: 'Bahasa Indonesia' },
      { id: 4, name: 'Pendidikan Jasmani Olahraga dan Kesehatan' },
      { id: 5, name: 'Sejarah Indonesia' },
      { id: 6, name: 'Seni Budaya' },
      { id: 7, name: 'Bahasa Lampung' },
    ],
    kejuruan: [
      { id: 8, name: 'Matematika' },
      { id: 9, name: 'Bahasa Inggris' },
      { id: 10, name: 'Informatika' },
      { id: 11, name: 'Project Ilmu Pengetahuan Alam dan Sosial (IPAS)' },
      { id: 12, name: 'Dasar Program Keahlian' },
    ],
  },
  '11': {
    umum: [
      { id: 1, name: 'Pendidikan Agama dan Budi Pekerti' },
      { id: 2, name: 'Pendidikan Pancasila dan Kewarganegaraan' },
      { id: 3, name: 'Bahasa Indonesia' },
      { id: 4, name: 'Pendidikan Jasmani Olahraga dan Kesehatan' },
      { id: 5, name: 'Sejarah Indonesia' },
      { id: 6, name: 'Pendidikan Anti Korupsi' },
    ],
    kejuruan: [
      { id: 7, name: 'Matematika' },
      { id: 8, name: 'Bahasa Inggris' },
      { id: 9, name: 'Kreativitas, Inovasi, dan Kewirausahaan (KIK)' },
      { id: 10, name: 'Konsentrasi Keahlian' },
      { id: 11, name: 'Mata Pelajaran Pilihan Kejuruan' },
    ],
  },
  '12': {
    umum: [
      { id: 1, name: 'Pendidikan Agama dan Budi Pekerti' },
      { id: 2, name: 'Pendidikan Pancasila dan Kewarganegaraan' },
      { id: 3, name: 'Bahasa Indonesia' },
      { id: 4, name: 'Pendidikan Anti Korupsi' },
    ],
    kejuruan: [
      { id: 5, name: 'Matematika' },
      { id: 6, name: 'Bahasa Inggris' },
      { id: 7, name: 'Kreativitas, Inovasi, dan Kewirausahaan (KIK)' },
      { id: 8, name: 'Konsentrasi Keahlian' },
      { id: 9, name: 'Mata Pelajaran Pilihan Kejuruan' },
    ],
  },
};

export const DEFAULT_STUDENTS: StudentItem[] = [
  { id: 101, nis: '20261001', nisn: '0081234561', name: 'ADITYA PRATAMA', className: '10 TKR 1', grade: '10', major: 'Teknik Kendaraan Ringan', attendance: { sakit: 1, izin: 0, alpa: 0 }, note: 'Sangat baik dalam praktikum kejuruan, pertahankan.' },
  { id: 102, nis: '20261002', nisn: '0081234562', name: 'ANNISA RAHMAWATI', className: '10 TKR 1', grade: '10', major: 'Teknik Kendaraan Ringan', attendance: { sakit: 0, izin: 1, alpa: 0 }, note: 'Prestasi akademik sangat memuaskan, tingkatkan keaktifan.' },
  { id: 103, nis: '20261003', nisn: '0081234563', name: 'BAGAS SAPUTRA', className: '10 TKR 1', grade: '10', major: 'Teknik Kendaraan Ringan', attendance: { sakit: 0, izin: 0, alpa: 0 }, note: 'Fokus pada peningkatan pemahaman materi teori kejuruan.' },
  { id: 201, nis: '20251101', nisn: '0071234588', name: 'DEWI LESTARI', className: '11 TKR 1', grade: '11', major: 'Teknik Kendaraan Ringan', attendance: { sakit: 0, izin: 0, alpa: 0 }, note: 'Sangat disiplin dan aktif pada kegiatan pembelajaran.' },
  { id: 301, nis: '20241201', nisn: '0061234599', name: 'FAJAR NUGROHO', className: '12 TKR 1', grade: '12', major: 'Teknik Kendaraan Ringan', attendance: { sakit: 0, izin: 0, alpa: 0 }, note: 'Tingkatkan persiapan menjelang Uji Kompetensi Keahlian.' },
];

export const DEFAULT_STUDENT_GRADES: StudentGradesMap = {
  101: { 'Pendidikan Agama dan Budi Pekerti': 85, 'Pendidikan Pancasila dan Kewarganegaraan': 88, 'Bahasa Indonesia': 82, 'Pendidikan Jasmani Olahraga dan Kesehatan': 90, 'Sejarah Indonesia': 80, 'Seni Budaya': 84, 'Bahasa Lampung': 86, 'Matematika': 78, 'Bahasa Inggris': 80, 'Informatika': 88, 'Project Ilmu Pengetahuan Alam dan Sosial (IPAS)': 83, 'Dasar Program Keahlian': 89 },
  102: { 'Pendidikan Agama dan Budi Pekerti': 92, 'Pendidikan Pancasila dan Kewarganegaraan': 90, 'Bahasa Indonesia': 89, 'Pendidikan Jasmani Olahraga dan Kesehatan': 88, 'Sejarah Indonesia': 87, 'Seni Budaya': 91, 'Bahasa Lampung': 90, 'Matematika': 85, 'Bahasa Inggris': 88, 'Informatika': 92, 'Project Ilmu Pengetahuan Alam dan Sosial (IPAS)': 89, 'Dasar Program Keahlian': 94 },
  103: { 'Pendidikan Agama dan Budi Pekerti': 78, 'Pendidikan Pancasila dan Kewarganegaraan': 75, 'Bahasa Indonesia': 80, 'Pendidikan Jasmani Olahraga dan Kesehatan': 85, 'Sejarah Indonesia': 76, 'Seni Budaya': 80, 'Bahasa Lampung': 78, 'Matematika': 72, 'Bahasa Inggris': 74, 'Informatika': 81, 'Project Ilmu Pengetahuan Alam dan Sosial (IPAS)': 79, 'Dasar Program Keahlian': 83 },
  201: { 'Pendidikan Agama dan Budi Pekerti': 88, 'Pendidikan Pancasila dan Kewarganegaraan': 86, 'Bahasa Indonesia': 85, 'Pendidikan Jasmani Olahraga dan Kesehatan': 89, 'Sejarah Indonesia': 84, 'Pendidikan Anti Korupsi': 90, 'Matematika': 82, 'Bahasa Inggris': 80, 'Kreativitas, Inovasi, dan Kewirausahaan (KIK)': 87, 'Konsentrasi Keahlian': 90, 'Mata Pelajaran Pilihan Kejuruan': 86 },
  301: { 'Pendidikan Agama dan Budi Pekerti': 90, 'Pendidikan Pancasila dan Kewarganegaraan': 88, 'Bahasa Indonesia': 87, 'Pendidikan Anti Korupsi': 92, 'Matematika': 84, 'Bahasa Inggris': 86, 'Kreativitas, Inovasi, dan Kewirausahaan (KIK)': 89, 'Konsentrasi Keahlian': 93, 'Mata Pelajaran Pilihan Kejuruan': 88 },
};
