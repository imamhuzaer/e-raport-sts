import React, { useEffect, useMemo, useRef, useState } from 'react';
import { User } from 'firebase/auth';
import {
  ClassItem,
  DatabasePayload,
  DEFAULT_CLASSES,
  DEFAULT_SCHOOL,
  DEFAULT_STUDENT_GRADES,
  DEFAULT_STUDENTS,
  DEFAULT_SUBJECT_TEMPLATES,
  DEFAULT_USERS,
  RankedStudentItem,
  RoleType,
  SchoolData,
  SignatureDisplayType,
  SPREADSHEET_ID,
  SPREADSHEET_URL,
  StudentGradesMap,
  StudentItem,
  UserAccount,
} from './types';
import {
  FIREBASE_PROJECT_ID,
  getAccessToken,
  GOOGLE_OAUTH_CLIENT_ID,
  googleSignIn,
  initAuth,
  logout as googleLogout,
} from './services/firebaseAuth';
import {
  fetchFromGoogleSheets,
  fetchFromGoogleSheetsPublic,
  getActiveSpreadsheetId,
  getActiveSpreadsheetUrl,
  getGasBridgeScriptCode,
  getSavedGasBridgeUrl,
  loadLocalDatabaseCache,
  resetActiveSpreadsheetLink,
  saveActiveSpreadsheetLink,
  saveGasBridgeUrl,
  saveLocalDatabaseCache,
  SheetMetadataInfo,
  syncAllToGoogleSheets,
} from './services/sheetsDatabase';
import {
  downloadClassTemplate,
  downloadGradesTemplate,
  downloadGuruMapelUserTemplate,
  downloadStudentTemplate,
  downloadWaliKelasUserTemplate,
  formatNipOrNuptk,
  formatTeacherNameWithTitle,
  generateAndDownloadPdf,
  parseBulkClasses,
  parseBulkGrades,
  parseBulkGuruMapelUsers,
  parseBulkStudents,
  parseBulkWaliKelasUsers,
  parsePastedTextToRows,
  processUploadedSignatureOrQrImage,
  readSpreadsheetFileToRows,
  removeWhiteBackgroundFromDataUrl,
} from './services/templateAndPdfService';
import { RaportTemplate } from './components/RaportTemplate';
import { GoogleSignInButton } from './components/GoogleSignInButton';

interface PendingSheetMutation {
  title: string;
  description: string;
  affectedSummary: string;
  isDestructive?: boolean;
  executeLocalAndReturnPayload: () => DatabasePayload;
}

type BulkImportType =
  | 'siswa'
  | 'kelas'
  | 'walikelas_user'
  | 'gurumapel_user'
  | 'nilai'
  | 'nilai_admin_semua_mapel';

type AdminTabType =
  | 'statistik'
  | 'nilai'
  | 'sekolah'
  | 'kelas'
  | 'siswa'
  | 'guru'
  | 'users'
  | 'database';

interface SavedAuthSession {
  isLoggedIn: boolean;
  showRoleModal: boolean;
  currentUser: UserAccount | null;
  activeRole: RoleType;
  adminTab?: AdminTabType;
  waliSelectedClass?: string;
  guruSelectedClass?: string;
  guruSelectedSubject?: string;
}

const LOCAL_AUTH_SESSION_KEY = 'eraport_saved_login_session_v1';

function loadSavedAuthSession(): SavedAuthSession | null {
  if (typeof window === 'undefined') return null;
  try {
    const raw = localStorage.getItem(LOCAL_AUTH_SESSION_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as SavedAuthSession;
    if (parsed && parsed.currentUser && parsed.currentUser.username) {
      return parsed;
    }
    return null;
  } catch {
    return null;
  }
}

function saveAuthSessionToLocal(session: SavedAuthSession | null): void {
  if (typeof window === 'undefined') return;
  try {
    if (!session || !session.currentUser) {
      localStorage.removeItem(LOCAL_AUTH_SESSION_KEY);
    } else {
      localStorage.setItem(LOCAL_AUTH_SESSION_KEY, JSON.stringify(session));
    }
  } catch {
    // Ignore storage quota errors
  }
}

export default function App() {
  const initialSession = useMemo(() => loadSavedAuthSession(), []);

  // Authentication & Role State (Hydrated from Local Browser Storage so Refresh Doesn't Logout)
  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(
    () => initialSession?.isLoggedIn ?? false
  );
  const [showRoleModal, setShowRoleModal] = useState<boolean>(
    () => initialSession?.showRoleModal ?? false
  );
  const [currentUser, setCurrentUser] = useState<UserAccount | null>(
    () => initialSession?.currentUser ?? null
  );
  const [activeRole, setActiveRole] = useState<RoleType>(
    () => initialSession?.activeRole ?? 'admin'
  );
  const [loginForm, setLoginForm] = useState({ username: '', password: '' });
  const [rememberLogin, setRememberLogin] = useState<boolean>(true);
  const [showLoginPassword, setShowLoginPassword] = useState<boolean>(false);
  const [showNewUserPassword, setShowNewUserPassword] =
    useState<boolean>(false);
  const [visibleUserPasswords, setVisibleUserPasswords] = useState<
    Record<number, boolean>
  >({});

  // Google Workspace OAuth & Vercel Bridge State
  const [needsGoogleAuth, setNeedsGoogleAuth] = useState(true);
  const [googleUser, setGoogleUser] = useState<User | null>(null);
  const [googleToken, setGoogleToken] = useState<string | null>(null);
  const [isSyncingSheets, setIsSyncingSheets] = useState(false);
  const [sheetMeta, setSheetMeta] = useState<SheetMetadataInfo | null>(null);
  const [gasBridgeUrlInput, setGasBridgeUrlInput] = useState<string>(() =>
    getSavedGasBridgeUrl()
  );
  const [customOAuthClientId, setCustomOAuthClientId] = useState<string>(() =>
    typeof window !== 'undefined'
      ? localStorage.getItem('eraport_custom_oauth_client_id') || ''
      : ''
  );
  const [showGasScriptModal, setShowGasScriptModal] = useState(false);
  const [vercelDomainError, setVercelDomainError] = useState<string | null>(
    null
  );
  const [activeSheetId, setActiveSheetId] = useState<string>(() =>
    getActiveSpreadsheetId()
  );
  const [activeSheetUrl, setActiveSheetUrl] = useState<string>(() =>
    getActiveSpreadsheetUrl()
  );
  const [spreadsheetLinkInput, setSpreadsheetLinkInput] = useState<string>(() =>
    getActiveSpreadsheetUrl()
  );

  // Core Application Database State
  const [school, setSchool] = useState<SchoolData>({ ...DEFAULT_SCHOOL });
  const [classes, setClasses] = useState<ClassItem[]>([...DEFAULT_CLASSES]);
  const [users, setUsers] = useState<UserAccount[]>(
    JSON.parse(JSON.stringify(DEFAULT_USERS))
  );
  const [students, setStudents] = useState<StudentItem[]>(
    JSON.parse(JSON.stringify(DEFAULT_STUDENTS))
  );
  const [studentGrades, setStudentGrades] = useState<StudentGradesMap>(
    JSON.parse(JSON.stringify(DEFAULT_STUDENT_GRADES))
  );
  const [subjectTemplates] = useState(DEFAULT_SUBJECT_TEMPLATES);

  // UI & Modal States
  const [adminTab, setAdminTab] = useState<AdminTabType>(
    () => initialSession?.adminTab ?? 'statistik'
  );
  const [adminStatFilterClass, setAdminStatFilterClass] = useState('');
  const [adminTop10FilterClass, setAdminTop10FilterClass] = useState('');
  const [adminGradeSelectedClass, setAdminGradeSelectedClass] =
    useState('10 TKR 1');
  const [logoImgError, setLogoImgError] = useState(false);
  const [toast, setToast] = useState<{
    show: boolean;
    message: string;
    type: 'success' | 'danger';
  }>({ show: false, message: '', type: 'success' });

  const [showAddUserModal, setShowAddUserModal] = useState(false);
  const [newUser, setNewUser] = useState<{
    name: string;
    nip: string;
    username: string;
    password: string;
    roles: RoleType[];
    waliKelasFor: string;
  }>({
    name: '',
    nip: '',
    username: '',
    password: '',
    roles: ['gurumapel'],
    waliKelasFor: '10 TKR 1',
  });

  const [showAddStudentModal, setShowAddStudentModal] = useState(false);
  const [newStudent, setNewStudent] = useState({
    name: '',
    nisn: '',
    nis: '',
    className: '10 TKR 1',
  });

  const [showAddClassModal, setShowAddClassModal] = useState(false);
  const [newClass, setNewClass] = useState<{
    grade: string;
    name: string;
    major: string;
    waliKelasName: string;
    waliKelasNip: string;
    waliSignatureType: SignatureDisplayType;
    waliSignatureUrl: string;
  }>({
    grade: '10',
    name: '',
    major: 'Teknik Kendaraan Ringan',
    waliKelasName: '',
    waliKelasNip: '',
    waliSignatureType: 'none',
    waliSignatureUrl: '',
  });

  // Edit Class Wali Kelas Modal
  const [editingClass, setEditingClass] = useState<ClassItem | null>(null);

  // Bulk Import Modal State
  const [bulkModalType, setBulkModalType] = useState<BulkImportType | null>(
    null
  );
  const [bulkPasteText, setBulkPasteText] = useState('');
  const [bulkReplaceExisting, setBulkReplaceExisting] = useState(false);

  const [adminStudentFilterClass, setAdminStudentFilterClass] = useState('');
  const [guruSelectedClass, setGuruSelectedClass] = useState(
    () => initialSession?.guruSelectedClass || '10 TKR 1'
  );
  const [guruSelectedSubject, setGuruSelectedSubject] = useState(
    () =>
      initialSession?.guruSelectedSubject ||
      'Pendidikan Agama dan Budi Pekerti'
  );
  const [waliSelectedClass, setWaliSelectedClass] = useState(
    () => initialSession?.waliSelectedClass || '10 TKR 1'
  );

  // Print & PDF Download Modal State
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [printMode, setPrintMode] = useState<'single' | 'class'>('single');
  const [selectedPrintStudent, setSelectedPrintStudent] =
    useState<RankedStudentItem | null>(null);
  const [isGeneratingPdf, setIsGeneratingPdf] = useState(false);
  const previewContainerRef = useRef<HTMLDivElement | null>(null);

  // Mandatory Confirmation Dialog State for Mutating/Destructive Operations
  const [pendingMutation, setPendingMutation] =
    useState<PendingSheetMutation | null>(null);

  const showToast = (msg: string, type: 'success' | 'danger' = 'success') => {
    setToast({ show: true, message: msg, type });
    setTimeout(() => {
      setToast((prev) => ({ ...prev, show: false }));
    }, 3500);
  };

  // Initialize Firebase Auth Listener & Auto-Load Spreadsheet Database on Mount (Works on Vercel without OAuth!)
  useEffect(() => {
    // 1. Hydrate from local cache first for instant startup
    const cached = loadLocalDatabaseCache();
    if (cached) {
      setSchool(cached.school);
      setClasses(cached.classes);
      setStudents(cached.students);
      setUsers(cached.users);
      setStudentGrades(cached.studentGrades);
    }

    // 2. Automatically fetch live database from Google Spreadsheet (Zero-Auth Public / GAS Bridge)
    fetchFromGoogleSheetsPublic()
      .then((result) => {
        if (!result.isEmptySheet) {
          setSchool(result.payload.school);
          setClasses(result.payload.classes);
          setStudents(result.payload.students);
          setUsers(result.payload.users);
          setStudentGrades(result.payload.studentGrades);
        }
        setSheetMeta(result.meta);
        const updatedGas = getSavedGasBridgeUrl();
        if (updatedGas) setGasBridgeUrlInput(updatedGas);
      })
      .catch((err) => {
        console.warn('Initial background sheet load notice:', err);
      });

    const unsubscribe = initAuth(
      (user, token) => {
        setGoogleUser(user);
        setGoogleToken(token);
        setNeedsGoogleAuth(false);
        setVercelDomainError(null);
      },
      () => {
        setGoogleUser(null);
        setGoogleToken(null);
        setNeedsGoogleAuth(true);
      }
    );
    return () => unsubscribe();
  }, []);

  // Load data from Google Sheets (Admin Only - Works WITH or WITHOUT Google OAuth Login!)
  const loadFromGoogleSheets = async (tokenOverride?: string) => {
    if (!currentUser?.roles.includes('admin') || activeRole !== 'admin') {
      showToast(
        'Akses ditolak: Hanya Administrator yang dapat mengakses database.',
        'danger'
      );
      return;
    }
    const activeToken =
      tokenOverride || (await getAccessToken()) || googleToken || null;

    setIsSyncingSheets(true);
    try {
      const result = await fetchFromGoogleSheets(activeToken);
      setSchool(result.payload.school);
      setClasses(result.payload.classes);
      setStudents(result.payload.students);
      setUsers(result.payload.users);
      setStudentGrades(result.payload.studentGrades);
      setSheetMeta(result.meta);
      const updatedGas = getSavedGasBridgeUrl();
      if (updatedGas) setGasBridgeUrlInput(updatedGas);

      const modeLabel =
        result.meta.connectionMode === 'oauth'
          ? 'Google OAuth'
          : result.meta.connectionMode === 'gas_bridge'
          ? 'Bridge Apps Script'
          : 'Koneksi Langsung Tanpa OAuth';

      if (result.isEmptySheet) {
        showToast(
          `Spreadsheet terhubung via ${modeLabel} (masih kosong). Klik "Tulis / Inisialisasi ke Spreadsheet" untuk mengisi tabel.`,
          'success'
        );
      } else {
        showToast(
          `Data berhasil dimuat dari Google Sheet (${result.meta.spreadsheetTitle}) via ${modeLabel}!`,
          'success'
        );
      }
    } catch (error: any) {
      console.error('Error loading Google Sheet:', error);
      showToast(
        error?.message || 'Gagal memuat data dari Google Sheet.',
        'danger'
      );
    } finally {
      setIsSyncingSheets(false);
    }
  };

  // Handle Google OAuth Sign-In (Admin Only)
  const handleGoogleConnect = async () => {
    if (!currentUser?.roles.includes('admin') || activeRole !== 'admin') {
      showToast(
        'Akses ditolak: Hanya Administrator yang dapat mengakses database.',
        'danger'
      );
      return;
    }
    setIsSyncingSheets(true);
    setVercelDomainError(null);
    try {
      const res = await googleSignIn(customOAuthClientId);
      if (res) {
        setGoogleUser(res.user);
        setGoogleToken(res.accessToken);
        setNeedsGoogleAuth(false);
        setVercelDomainError(null);
        await loadFromGoogleSheets(res.accessToken);
      }
    } catch (error: any) {
      console.error('Google Sign-In failed:', error);
      const host =
        typeof window !== 'undefined' ? window.location.hostname : 'vercel.app';
      const isDomainErr =
        error?.code === 'auth/unauthorized-domain' ||
        String(error?.message || '').includes('unauthorized-domain') ||
        String(error?.message || '').includes('Authorized Domains');

      if (isDomainErr) {
        setVercelDomainError(host);
        setAdminTab('database');
        // Still load data from Spreadsheet automatically via Direct/Public mode!
        await loadFromGoogleSheets(undefined);
        showToast(
          `Data Spreadsheet tetap berhasil dimuat secara langsung! Untuk login OAuth di domain "${host}", lihat panduan di tab Database.`,
          'danger'
        );
      } else {
        showToast(
          error?.message ||
            'Gagal otorisasi Google Sheets. Anda tetap dapat membaca data secara langsung di tab Database.',
          'danger'
        );
      }
    } finally {
      setIsSyncingSheets(false);
    }
  };

  // Execute confirmed mutation and sync to Google Sheets (via OAuth OR GAS Web App Bridge)
  const confirmAndExecuteMutation = async () => {
    if (!pendingMutation) return;
    const mutation = pendingMutation;
    setPendingMutation(null);

    const updatedPayload = mutation.executeLocalAndReturnPayload();
    saveLocalDatabaseCache(updatedPayload);

    const activeToken = (await getAccessToken()) || googleToken;
    const savedGasUrl = getSavedGasBridgeUrl();

    if (!activeToken && !savedGasUrl) {
      showToast(
        'Perubahan tersimpan di perangkat ini! (Hubungkan Google OAuth atau pasang URL Bridge Apps Script di tab Database untuk menulis otomatis ke Spreadsheet).',
        'success'
      );
      return;
    }

    setIsSyncingSheets(true);
    try {
      const updatedMeta = await syncAllToGoogleSheets(
        activeToken,
        updatedPayload
      );
      setSheetMeta(updatedMeta);
      showToast(
        `Perubahan berhasil disimpan dan disinkronkan ke Google Sheets (${
          updatedMeta.connectionMode === 'gas_bridge'
            ? 'via Apps Script Bridge'
            : 'via Google OAuth'
        })!`,
        'success'
      );
    } catch (error: any) {
      console.error('Failed syncing to Google Sheets:', error);
      showToast(
        `Tersimpan lokal, namun gagal menulis ke Google Sheets: ${
          error?.message || 'Error'
        }`,
        'danger'
      );
    } finally {
      setIsSyncingSheets(false);
    }
  };

  // Computed Data
  const filteredAdminStudents = useMemo(() => {
    if (!adminStudentFilterClass) return students;
    return students.filter((s) => s.className === adminStudentFilterClass);
  }, [students, adminStudentFilterClass]);

  const currentClassObj = useMemo(() => {
    return (
      classes.find((c) => c.name === guruSelectedClass) ||
      classes[0] || {
        id: 1,
        name: '10 TKR 1',
        grade: '10',
        major: 'Teknik Kendaraan Ringan',
        waliKelasName: '-',
        waliKelasNip: '-',
      }
    );
  }, [classes, guruSelectedClass]);

  const availableSubjectsForSelectedClass = useMemo(() => {
    const grade = currentClassObj ? currentClassObj.grade : '10';
    const template = subjectTemplates[grade] || subjectTemplates['10'];
    const list: Array<{ name: string; category: string }> = [];
    template.umum.forEach((s) =>
      list.push({ name: s.name, category: 'Umum' })
    );
    template.kejuruan.forEach((s) =>
      list.push({ name: s.name, category: 'Kejuruan' })
    );
    return list;
  }, [currentClassObj, subjectTemplates]);

  useEffect(() => {
    setLogoImgError(false);
  }, [school.logoUrl]);

  useEffect(() => {
    if (
      availableSubjectsForSelectedClass.length > 0 &&
      !availableSubjectsForSelectedClass.some(
        (s) => s.name === guruSelectedSubject
      )
    ) {
      setGuruSelectedSubject(availableSubjectsForSelectedClass[0].name);
    }
  }, [availableSubjectsForSelectedClass, guruSelectedSubject]);

  // Admin All-Subjects Grade Input Computed Data
  const adminGradeClassObj = useMemo(() => {
    return (
      classes.find(
        (c) =>
          c.name.trim().toLowerCase() ===
          adminGradeSelectedClass.trim().toLowerCase()
      ) ||
      classes[0] || {
        id: 1,
        name: '10 TKR 1',
        grade: '10',
        major: 'Teknik Kendaraan Ringan',
        waliKelasName: '-',
        waliKelasNip: '-',
      }
    );
  }, [classes, adminGradeSelectedClass]);

  const adminGradeClassStudents = useMemo(() => {
    return students.filter(
      (s) =>
        s.className.trim().toLowerCase() ===
        adminGradeSelectedClass.trim().toLowerCase()
    );
  }, [students, adminGradeSelectedClass]);

  const adminGradeClassSubjects = useMemo(() => {
    const grade = adminGradeClassObj ? adminGradeClassObj.grade : '10';
    const template = subjectTemplates[grade] || subjectTemplates['10'];
    const list: Array<{ name: string; category: 'Umum' | 'Kejuruan' }> = [];
    template.umum.forEach((s) =>
      list.push({ name: s.name, category: 'Umum' })
    );
    template.kejuruan.forEach((s) =>
      list.push({ name: s.name, category: 'Kejuruan' })
    );
    return list;
  }, [adminGradeClassObj, subjectTemplates]);

  const normalizeClassKey = (name?: string) =>
    (name || '').replace(/\s+/g, '').toLowerCase();

  const resolveLockedWaliClassForUser = (
    user: UserAccount | null,
    classesList: ClassItem[]
  ): string => {
    if (!user) return classesList[0]?.name || '10 TKR 1';

    if (
      user.waliKelasFor &&
      user.waliKelasFor.trim() !== '' &&
      user.waliKelasFor.trim() !== '-'
    ) {
      const matchedCls = classesList.find(
        (c) =>
          normalizeClassKey(c.name) === normalizeClassKey(user.waliKelasFor)
      );
      return matchedCls ? matchedCls.name : user.waliKelasFor.trim();
    }

    const cleanUserNip = (user.nip || '').replace(/\D/g, '');
    if (cleanUserNip.length >= 8) {
      const matchedByNip = classesList.find(
        (c) => (c.waliKelasNip || '').replace(/\D/g, '') === cleanUserNip
      );
      if (matchedByNip) return matchedByNip.name;
    }

    if (user.name && user.name.trim() !== '' && user.name.trim() !== '-') {
      const normUserName = formatTeacherNameWithTitle(user.name).toLowerCase();
      const matchedByName = classesList.find(
        (c) =>
          formatTeacherNameWithTitle(c.waliKelasName).toLowerCase() ===
            normUserName ||
          (c.waliKelasName || '').trim().toLowerCase() ===
            user.name.trim().toLowerCase()
      );
      if (matchedByName) return matchedByName.name;
    }

    return classesList[0]?.name || '10 TKR 1';
  };

  useEffect(() => {
    if (currentUser && currentUser.roles.includes('walikelas')) {
      const lockedCls = resolveLockedWaliClassForUser(currentUser, classes);
      if (lockedCls && lockedCls !== waliSelectedClass) {
        setWaliSelectedClass(lockedCls);
      }
    }
  }, [currentUser, classes, activeRole, waliSelectedClass]);

  const currentClassStudents = useMemo(() => {
    const clsName =
      activeRole === 'gurumapel' ? guruSelectedClass : waliSelectedClass;
    const normTarget = normalizeClassKey(clsName);
    return students.filter(
      (s) => normalizeClassKey(s.className) === normTarget
    );
  }, [students, activeRole, guruSelectedClass, waliSelectedClass]);

  const rankedClassStudents = useMemo<RankedStudentItem[]>(() => {
    const studentsInClass = currentClassStudents.map((s) => {
      const gradesObj = studentGrades[s.id] || {};
      const scores = Object.values(gradesObj).filter(
        (val): val is number => typeof val === 'number' && !Number.isNaN(val)
      );
      const sum = scores.reduce((a, b) => a + b, 0);
      const avg =
        scores.length > 0 ? Number((sum / scores.length).toFixed(2)) : 0;
      return {
        ...s,
        totalScore: sum,
        avgScore: avg,
        rank: 0,
      };
    });

    const sorted = [...studentsInClass].sort((a, b) => b.avgScore - a.avgScore);
    return sorted.map((s, idx) => ({
      ...s,
      rank: idx + 1,
    }));
  }, [currentClassStudents, studentGrades]);

  // Sync currentUser with latest users list when database loads, and persist session in localStorage
  useEffect(() => {
    if (!currentUser) return;
    const updatedMatch = users.find(
      (u) =>
        u.username.trim().toLowerCase() ===
          currentUser.username.trim().toLowerCase() || u.id === currentUser.id
    );
    if (
      updatedMatch &&
      JSON.stringify(updatedMatch) !== JSON.stringify(currentUser)
    ) {
      setCurrentUser(updatedMatch);
    }
  }, [users, currentUser]);

  useEffect(() => {
    if ((isLoggedIn || showRoleModal) && currentUser) {
      if (rememberLogin) {
        saveAuthSessionToLocal({
          isLoggedIn,
          showRoleModal,
          currentUser,
          activeRole,
          adminTab,
          waliSelectedClass,
          guruSelectedClass,
          guruSelectedSubject,
        });
      }
    }
  }, [
    isLoggedIn,
    showRoleModal,
    currentUser,
    activeRole,
    adminTab,
    waliSelectedClass,
    guruSelectedClass,
    guruSelectedSubject,
    rememberLogin,
  ]);

  // Authentication Handlers
  const performLoginWithUser = (user: UserAccount) => {
    setCurrentUser(user);
    setLoginForm({ username: '', password: '' });
    setShowLoginPassword(false);

    const lockedCls = resolveLockedWaliClassForUser(user, classes);
    if (lockedCls) {
      setWaliSelectedClass(lockedCls);
    }

    if (user.roles.length > 1) {
      setShowRoleModal(true);
      if (rememberLogin) {
        saveAuthSessionToLocal({
          isLoggedIn: false,
          showRoleModal: true,
          currentUser: user,
          activeRole: user.roles[0],
          adminTab,
          waliSelectedClass: lockedCls || waliSelectedClass,
          guruSelectedClass,
          guruSelectedSubject,
        });
      } else {
        saveAuthSessionToLocal(null);
      }
    } else {
      const nextRole = user.roles[0];
      setActiveRole(nextRole);
      setIsLoggedIn(true);
      if (rememberLogin) {
        saveAuthSessionToLocal({
          isLoggedIn: true,
          showRoleModal: false,
          currentUser: user,
          activeRole: nextRole,
          adminTab,
          waliSelectedClass: lockedCls || waliSelectedClass,
          guruSelectedClass,
          guruSelectedSubject,
        });
      } else {
        saveAuthSessionToLocal(null);
      }
      showToast(
        `Selamat datang, ${formatTeacherNameWithTitle(user.name)}!`,
        'success'
      );
    }
  };

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    const user = users.find(
      (u) =>
        u.username.toLowerCase() === loginForm.username.trim().toLowerCase() &&
        u.password === loginForm.password
    );
    if (!user) {
      showToast('Username atau kata sandi tidak cocok!', 'danger');
      return;
    }
    performLoginWithUser(user);
  };

  const selectRole = (role: RoleType) => {
    let nextWaliCls = waliSelectedClass;
    if (role === 'walikelas' && currentUser) {
      const lockedCls = resolveLockedWaliClassForUser(currentUser, classes);
      if (lockedCls) {
        nextWaliCls = lockedCls;
        setWaliSelectedClass(lockedCls);
      }
    }
    setActiveRole(role);
    setShowRoleModal(false);
    setIsLoggedIn(true);
    if (rememberLogin && currentUser) {
      saveAuthSessionToLocal({
        isLoggedIn: true,
        showRoleModal: false,
        currentUser,
        activeRole: role,
        adminTab,
        waliSelectedClass: nextWaliCls,
        guruSelectedClass,
        guruSelectedSubject,
      });
    }
    const roleLabel =
      role === 'walikelas'
        ? 'Wali Kelas'
        : role === 'gurumapel'
        ? 'Guru Mapel'
        : 'Administrator';
    showToast(`Berhasil masuk sebagai ${roleLabel}`, 'success');
  };

  const handleLogout = () => {
    saveAuthSessionToLocal(null);
    setIsLoggedIn(false);
    setShowRoleModal(false);
    setCurrentUser(null);
    setShowLoginPassword(false);
    showToast('Anda telah keluar dari aplikasi.', 'success');
  };

  /**
   * Resolves Class Info (including Wali Kelas Name & NIP/NUPTK) from:
   * 1. classes list
   * 2. users list (where role includes 'walikelas' and waliKelasFor matches className)
   * 3. currentUser if logged in as Wali Kelas for that class
   */
  const getClassInfo = (className: string): Partial<ClassItem> => {
    const normTarget = normalizeClassKey(className);
    const cls = classes.find((c) => normalizeClassKey(c.name) === normTarget);
    const waliUser = users.find(
      (u) =>
        u.roles.includes('walikelas') &&
        normalizeClassKey(u.waliKelasFor) === normTarget
    );

    let resolvedWaliName = cls?.waliKelasName || '';
    if (
      !resolvedWaliName ||
      resolvedWaliName.trim() === '-' ||
      resolvedWaliName.trim() === 'Nama Wali Kelas'
    ) {
      if (waliUser?.name) {
        resolvedWaliName = waliUser.name;
      } else if (
        currentUser?.roles.includes('walikelas') &&
        (!currentUser.waliKelasFor ||
          normalizeClassKey(currentUser.waliKelasFor) === normTarget)
      ) {
        resolvedWaliName = currentUser.name;
      }
    }

    let resolvedWaliNip = cls?.waliKelasNip || '';
    if (!resolvedWaliNip || resolvedWaliNip.trim() === '-') {
      if (waliUser?.nip && waliUser.nip.trim() !== '-') {
        resolvedWaliNip = waliUser.nip;
      } else if (
        currentUser?.roles.includes('walikelas') &&
        currentUser?.nip &&
        currentUser.nip.trim() !== '-'
      ) {
        resolvedWaliNip = currentUser.nip;
      }
    }

    let resolvedWaliSigType: SignatureDisplayType =
      cls?.waliSignatureType && cls.waliSignatureType !== 'none'
        ? cls.waliSignatureType
        : waliUser?.waliSignatureType && waliUser.waliSignatureType !== 'none'
        ? waliUser.waliSignatureType
        : 'none';

    let resolvedWaliSigUrl =
      cls?.waliSignatureUrl || waliUser?.waliSignatureUrl || '';

    if (
      resolvedWaliSigUrl &&
      (!resolvedWaliSigType || resolvedWaliSigType === 'none')
    ) {
      resolvedWaliSigType = 'signature';
    }

    return {
      ...cls,
      name: cls?.name || className,
      waliKelasName: resolvedWaliName,
      waliKelasNip: resolvedWaliNip,
      waliSignatureType: resolvedWaliSigType,
      waliSignatureUrl: resolvedWaliSigUrl,
    };
  };

  const getSubjectsForGrade = (grade: string) => {
    return subjectTemplates[grade] || subjectTemplates['10'];
  };

  // Image Upload Handlers for Scan TTD / QR Code (Kepala Sekolah & Wali Kelas)
  const handleUploadPrincipalSigOrQr = async (
    e: React.ChangeEvent<HTMLInputElement>,
    targetMode?: 'signature' | 'qrcode'
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const mode: 'signature' | 'qrcode' =
      targetMode ||
      (school.principalSignatureType === 'qrcode' ? 'qrcode' : 'signature');
    try {
      const dataUrl = await processUploadedSignatureOrQrImage(
        file,
        mode,
        mode === 'signature'
      );
      setSchool((prev) => ({
        ...prev,
        principalSignatureType: mode,
        principalSignatureUrl: dataUrl,
      }));
      showToast(
        `${
          mode === 'qrcode' ? 'QR Code' : 'Scan TTD'
        } Kepala Sekolah berhasil diunggah! Klik "Simpan Data Sekolah" untuk menyimpan.`,
        'success'
      );
    } catch (err: any) {
      showToast(
        err?.message || 'Gagal memproses file gambar TTD/QR Kepala Sekolah.',
        'danger'
      );
    } finally {
      e.target.value = '';
    }
  };

  const handleUploadEditingClassWaliSigOrQr = async (
    e: React.ChangeEvent<HTMLInputElement>,
    targetMode?: 'signature' | 'qrcode'
  ) => {
    const file = e.target.files?.[0];
    if (!file || !editingClass) return;
    const mode: 'signature' | 'qrcode' =
      targetMode ||
      (editingClass.waliSignatureType === 'qrcode' ? 'qrcode' : 'signature');
    try {
      const dataUrl = await processUploadedSignatureOrQrImage(
        file,
        mode,
        mode === 'signature'
      );
      setEditingClass({
        ...editingClass,
        waliSignatureType: mode,
        waliSignatureUrl: dataUrl,
      });
      showToast(
        `${
          mode === 'qrcode' ? 'QR Code' : 'Scan TTD'
        } Wali Kelas berhasil dimuat! Klik "Simpan Perubahan" untuk menyimpan.`,
        'success'
      );
    } catch (err: any) {
      showToast(
        err?.message || 'Gagal memproses file gambar TTD/QR Wali Kelas.',
        'danger'
      );
    } finally {
      e.target.value = '';
    }
  };

  const handleUploadNewClassWaliSigOrQr = async (
    e: React.ChangeEvent<HTMLInputElement>,
    targetMode?: 'signature' | 'qrcode'
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const mode: 'signature' | 'qrcode' =
      targetMode ||
      (newClass.waliSignatureType === 'qrcode' ? 'qrcode' : 'signature');
    try {
      const dataUrl = await processUploadedSignatureOrQrImage(
        file,
        mode,
        mode === 'signature'
      );
      setNewClass((prev) => ({
        ...prev,
        waliSignatureType: mode,
        waliSignatureUrl: dataUrl,
      }));
      showToast(
        `${
          mode === 'qrcode' ? 'QR Code' : 'Scan TTD'
        } Wali Kelas berhasil dimuat!`,
        'success'
      );
    } catch (err: any) {
      showToast(
        err?.message || 'Gagal memproses file gambar TTD/QR Wali Kelas.',
        'danger'
      );
    } finally {
      e.target.value = '';
    }
  };

  // Change / Reset Active Spreadsheet Link Handlers (Admin Only)
  const handleSaveCustomSpreadsheetLink = async () => {
    const saved = saveActiveSpreadsheetLink(spreadsheetLinkInput);
    if (!saved) {
      showToast(
        'Format Link atau ID Google Spreadsheet tidak valid. Pastikan menyalin URL lengkap dari Google Sheets.',
        'danger'
      );
      return;
    }
    setActiveSheetId(saved.id);
    setActiveSheetUrl(saved.url);
    setSpreadsheetLinkInput(saved.url);
    showToast(
      `Link Database Spreadsheet berhasil diubah (ID: ${saved.id.slice(
        0,
        14
      )}...)! Memuat data dari Spreadsheet baru...`,
      'success'
    );
    await loadFromGoogleSheets();
  };

  const handleResetSpreadsheetLink = async () => {
    const reset = resetActiveSpreadsheetLink();
    setActiveSheetId(reset.id);
    setActiveSheetUrl(reset.url);
    setSpreadsheetLinkInput(reset.url);
    showToast(
      'Link Database Spreadsheet dikembalikan ke Spreadsheet Default!',
      'success'
    );
    await loadFromGoogleSheets();
  };

  // Admin & CRUD Operations
  const handleSaveSchoolData = () => {
    setPendingMutation({
      title: 'Simpan Identitas Sekolah & TTD/QR Kepala Sekolah',
      description:
        'Anda akan memperbarui data Identitas Sekolah, Logo, serta TTD/QR Code Kepala Sekolah SMKN 1 Mesuji Timur.',
      affectedSummary: `Sekolah: ${school.name} (TP ${
        school.tp
      }, Kepsek: ${formatTeacherNameWithTitle(school.principal)}, Mode TTD/QR: ${
        school.principalSignatureType === 'qrcode'
          ? 'QR Code KS'
          : school.principalSignatureType === 'signature'
          ? 'Scan TTD KS'
          : 'Manual'
      })`,
      executeLocalAndReturnPayload: () => ({
        school,
        classes,
        students,
        users,
        studentGrades,
      }),
    });
  };

  const handleAddNewClass = () => {
    if (!newClass.name.trim()) {
      showToast('Nama kelas tidak boleh kosong.', 'danger');
      return;
    }
    const nextClsId =
      classes.reduce((max, c) => Math.max(max, c.id || 0), 0) + 1;
    const createdClass: ClassItem = {
      id: nextClsId,
      grade: newClass.grade,
      name: newClass.name.trim(),
      major: newClass.major.trim() || 'Teknik Kendaraan Ringan',
      waliKelasName: newClass.waliKelasName.trim() || '-',
      waliKelasNip: newClass.waliKelasNip.trim() || '-',
      waliSignatureType: newClass.waliSignatureUrl
        ? newClass.waliSignatureType === 'none'
          ? 'signature'
          : newClass.waliSignatureType
        : 'none',
      waliSignatureUrl: newClass.waliSignatureUrl || '',
    };

    setPendingMutation({
      title: 'Tambah Rombongan Belajar Baru',
      description: 'Anda akan menambahkan kelas/rombel baru ke dalam sistem.',
      affectedSummary: `Menambahkan rombel "${createdClass.name}" (Tingkat ${
        createdClass.grade
      }, Wali: ${formatTeacherNameWithTitle(createdClass.waliKelasName)})`,
      executeLocalAndReturnPayload: () => {
        const nextClasses = [...classes, createdClass];
        setClasses(nextClasses);
        setNewClass({
          grade: '10',
          name: '',
          major: 'Teknik Kendaraan Ringan',
          waliKelasName: '',
          waliKelasNip: '',
          waliSignatureType: 'none',
          waliSignatureUrl: '',
        });
        setShowAddClassModal(false);
        return {
          school,
          classes: nextClasses,
          students,
          users,
          studentGrades,
        };
      },
    });
  };

  const handleSaveEditedClass = () => {
    if (!editingClass) return;
    const normalizedSigType: SignatureDisplayType =
      editingClass.waliSignatureUrl &&
      (!editingClass.waliSignatureType ||
        editingClass.waliSignatureType === 'none')
        ? 'signature'
        : editingClass.waliSignatureUrl
        ? editingClass.waliSignatureType || 'signature'
        : 'none';

    const updatedCls: ClassItem = {
      ...editingClass,
      name: editingClass.name.trim(),
      major: editingClass.major.trim() || 'Teknik Kendaraan Ringan',
      waliKelasName: editingClass.waliKelasName.trim() || '-',
      waliKelasNip: editingClass.waliKelasNip.trim() || '-',
      waliSignatureType: normalizedSigType,
      waliSignatureUrl: editingClass.waliSignatureUrl || '',
    };

    setPendingMutation({
      title: 'Simpan Perubahan Data Kelas, Wali Kelas & TTD/QR',
      description: `Anda akan memperbarui data Wali Kelas, NIP/NUPTK, serta TTD/QR Code untuk rombel ${updatedCls.name}.`,
      affectedSummary: `Rombel ${updatedCls.name}: Wali Kelas "${formatTeacherNameWithTitle(
        updatedCls.waliKelasName
      )}"${
        formatNipOrNuptk(updatedCls.waliKelasNip).formatted
          ? ` (${formatNipOrNuptk(updatedCls.waliKelasNip).formatted})`
          : ''
      } • Mode TTD: ${
        updatedCls.waliSignatureType === 'qrcode'
          ? 'QR Code Wali Kelas'
          : updatedCls.waliSignatureType === 'signature'
          ? 'Scan TTD Wali Kelas'
          : 'Manual'
      }`,
      executeLocalAndReturnPayload: () => {
        const nextClasses = classes.map((c) =>
          c.id === updatedCls.id ? updatedCls : c
        );
        // Also sync back to any user assigned as walikelas for this class
        const nextUsers = users.map((u) => {
          if (
            u.roles.includes('walikelas') &&
            normalizeClassKey(u.waliKelasFor) ===
              normalizeClassKey(updatedCls.name)
          ) {
            return {
              ...u,
              name:
                updatedCls.waliKelasName !== '-'
                  ? updatedCls.waliKelasName
                  : u.name,
              nip:
                updatedCls.waliKelasNip !== '-'
                  ? updatedCls.waliKelasNip
                  : u.nip,
              waliSignatureType: updatedCls.waliSignatureType,
              waliSignatureUrl: updatedCls.waliSignatureUrl,
            };
          }
          return u;
        });
        setClasses(nextClasses);
        setUsers(nextUsers);
        setEditingClass(null);
        return {
          school,
          classes: nextClasses,
          students,
          users: nextUsers,
          studentGrades,
        };
      },
    });
  };

  const handleRemoveClass = (cls: ClassItem) => {
    setPendingMutation({
      title: 'Hapus Data Kelas',
      description:
        'Apakah Anda yakin ingin menghapus rombongan belajar ini dari database?',
      affectedSummary: `Menghapus rombel "${cls.name}" (Wali Kelas: ${cls.waliKelasName})`,
      isDestructive: true,
      executeLocalAndReturnPayload: () => {
        const nextClasses = classes.filter((c) => c.id !== cls.id);
        setClasses(nextClasses);
        return {
          school,
          classes: nextClasses,
          students,
          users,
          studentGrades,
        };
      },
    });
  };

  const handleAddNewStudent = () => {
    if (!newStudent.name.trim()) {
      showToast('Nama siswa tidak boleh kosong.', 'danger');
      return;
    }
    const clsObj = classes.find((c) => c.name === newStudent.className) || {
      grade: '10',
      major: 'Teknik Kendaraan Ringan',
    };
    const newId =
      students.reduce((max, s) => Math.max(max, s.id || 100), 100) + 1;
    const createdStudent: StudentItem = {
      id: newId,
      nis:
        newStudent.nis.trim() ||
        '2026' + Math.floor(1000 + Math.random() * 9000),
      nisn:
        newStudent.nisn.trim() ||
        '008' + Math.floor(1000000 + Math.random() * 9000000),
      name: newStudent.name.trim().toUpperCase(),
      className: newStudent.className,
      grade: clsObj.grade,
      major: clsObj.major,
      attendance: { sakit: 0, izin: 0, alpa: 0 },
      note: 'Tingkatkan motivasi belajar.',
    };

    setPendingMutation({
      title: 'Tambah Siswa Baru',
      description: 'Anda akan menambahkan data peserta didik baru.',
      affectedSummary: `Siswa: "${createdStudent.name}" (NISN: ${createdStudent.nisn}, Kelas: ${createdStudent.className})`,
      executeLocalAndReturnPayload: () => {
        const nextStudents = [...students, createdStudent];
        const nextGrades = { ...studentGrades, [newId]: {} };
        setStudents(nextStudents);
        setStudentGrades(nextGrades);
        setNewStudent({
          name: '',
          nisn: '',
          nis: '',
          className: classes[0]?.name || '10 TKR 1',
        });
        setShowAddStudentModal(false);
        return {
          school,
          classes,
          students: nextStudents,
          users,
          studentGrades: nextGrades,
        };
      },
    });
  };

  const handleDeleteStudent = (st: StudentItem) => {
    setPendingMutation({
      title: 'Hapus Peserta Didik',
      description:
        'Tindakan ini akan menghapus data siswa beserta nilai terkait dari sistem.',
      affectedSummary: `Menghapus siswa "${st.name}" (NISN: ${st.nisn}, Kelas: ${st.className})`,
      isDestructive: true,
      executeLocalAndReturnPayload: () => {
        const nextStudents = students.filter((s) => s.id !== st.id);
        const nextGrades = { ...studentGrades };
        delete nextGrades[st.id];
        setStudents(nextStudents);
        setStudentGrades(nextGrades);
        return {
          school,
          classes,
          students: nextStudents,
          users,
          studentGrades: nextGrades,
        };
      },
    });
  };

  const handleAddNewUser = () => {
    if (
      !newUser.name.trim() ||
      !newUser.username.trim() ||
      !newUser.password.trim()
    ) {
      showToast('Mohon lengkapi semua data pengguna.', 'danger');
      return;
    }
    if (newUser.roles.length === 0) {
      showToast('Pilih setidaknya satu role untuk pengguna ini.', 'danger');
      return;
    }

    const nextUsrId =
      users.reduce((max, u) => Math.max(max, u.id || 0), 0) + 1;
    const createdUser: UserAccount = {
      id: nextUsrId,
      name: newUser.name.trim(),
      nip: newUser.nip.trim() || '-',
      username: newUser.username.trim(),
      password: newUser.password,
      roles: [...newUser.roles],
      waliKelasFor: newUser.roles.includes('walikelas')
        ? newUser.waliKelasFor
        : '',
    };

    setPendingMutation({
      title: 'Tambah Pengguna & Hak Akses',
      description:
        'Anda akan menyimpan akun pengguna baru beserta hak akses rolenya.',
      affectedSummary: `Pengguna: "${formatTeacherNameWithTitle(
        createdUser.name
      )}" (${
        formatNipOrNuptk(createdUser.nip).formatted
          ? `${formatNipOrNuptk(createdUser.nip).formatted}, `
          : ''
      }Username: ${createdUser.username})`,
      executeLocalAndReturnPayload: () => {
        const nextUsers = [...users, createdUser];
        let nextClasses = [...classes];
        // Automatically update Wali Kelas Name & NIP on the target class
        if (
          createdUser.roles.includes('walikelas') &&
          createdUser.waliKelasFor
        ) {
          nextClasses = nextClasses.map((c) =>
            c.name.trim().toLowerCase() ===
            createdUser.waliKelasFor.trim().toLowerCase()
              ? {
                  ...c,
                  waliKelasName: createdUser.name,
                  waliKelasNip: createdUser.nip || c.waliKelasNip,
                }
              : c
          );
          setClasses(nextClasses);
        }
        setUsers(nextUsers);
        setNewUser({
          name: '',
          nip: '',
          username: '',
          password: '',
          roles: ['gurumapel'],
          waliKelasFor: classes[0]?.name || '10 TKR 1',
        });
        setShowAddUserModal(false);
        return {
          school,
          classes: nextClasses,
          students,
          users: nextUsers,
          studentGrades,
        };
      },
    });
  };

  const handleDeleteUser = (u: UserAccount) => {
    if (u.username === 'admin') return;
    setPendingMutation({
      title: 'Hapus Akun Pengguna',
      description: 'Apakah Anda yakin ingin menghapus akun pengguna ini?',
      affectedSummary: `Menghapus pengguna "${u.name}" (@${u.username})`,
      isDestructive: true,
      executeLocalAndReturnPayload: () => {
        const nextUsers = users.filter((item) => item.id !== u.id);
        setUsers(nextUsers);
        return {
          school,
          classes,
          students,
          users: nextUsers,
          studentGrades,
        };
      },
    });
  };

  // ============================================================================
  // BULK IMPORT HANDLERS (FILE UPLOAD OR COPY-PASTE)
  // ============================================================================

  const processBulkRows = (rows: string[][]) => {
    if (!bulkModalType || rows.length === 0) {
      showToast('Data kosong atau format tidak terbaca.', 'danger');
      return;
    }

    if (bulkModalType === 'kelas') {
      const parsedClasses = parseBulkClasses(rows);
      if (parsedClasses.length === 0) {
        showToast('Tidak ditemukan baris data kelas yang valid.', 'danger');
        return;
      }
      setPendingMutation({
        title: `Impor Massal ${parsedClasses.length} Kelas`,
        description: bulkReplaceExisting
          ? 'Mengganti daftar kelas lama dengan data kelas baru dari hasil impor.'
          : 'Menambahkan/memperbarui data kelas dan wali kelas dari hasil impor.',
        affectedSummary: `${parsedClasses.length} Rombel (${parsedClasses
          .slice(0, 4)
          .map((c) => c.name)
          .join(', ')}${parsedClasses.length > 4 ? '...' : ''})`,
        executeLocalAndReturnPayload: () => {
          let nextClasses: ClassItem[];
          if (bulkReplaceExisting) {
            nextClasses = parsedClasses;
          } else {
            const mapByName = new Map<string, ClassItem>();
            classes.forEach((c) =>
              mapByName.set(c.name.trim().toLowerCase(), c)
            );
            parsedClasses.forEach((pc) => {
              const existing = mapByName.get(pc.name.trim().toLowerCase());
              if (existing) {
                mapByName.set(pc.name.trim().toLowerCase(), {
                  ...existing,
                  ...pc,
                  id: existing.id,
                });
              } else {
                mapByName.set(pc.name.trim().toLowerCase(), pc);
              }
            });
            nextClasses = Array.from(mapByName.values());
          }
          setClasses(nextClasses);
          setBulkModalType(null);
          setBulkPasteText('');
          return {
            school,
            classes: nextClasses,
            students,
            users,
            studentGrades,
          };
        },
      });
    } else if (bulkModalType === 'siswa') {
      const parsedStudents = parseBulkStudents(rows, classes);
      if (parsedStudents.length === 0) {
        showToast('Tidak ditemukan baris data siswa yang valid.', 'danger');
        return;
      }
      setPendingMutation({
        title: `Impor Massal ${parsedStudents.length} Siswa`,
        description: bulkReplaceExisting
          ? 'Mengganti seluruh data siswa lama dengan daftar siswa hasil impor.'
          : 'Menambahkan daftar siswa baru dari template ke dalam database.',
        affectedSummary: `${parsedStudents.length} Peserta Didik (${parsedStudents
          .slice(0, 3)
          .map((s) => s.name)
          .join(', ')}${parsedStudents.length > 3 ? '...' : ''})`,
        executeLocalAndReturnPayload: () => {
          const nextStudents = bulkReplaceExisting
            ? parsedStudents
            : [...students, ...parsedStudents];
          const nextGrades = { ...studentGrades };
          parsedStudents.forEach((st) => {
            if (!nextGrades[st.id]) nextGrades[st.id] = {};
          });
          // Ensure any new class mentioned in student import is also registered
          const nextClasses = [...classes];
          const clsNames = new Set(
            nextClasses.map((c) => c.name.trim().toLowerCase())
          );
          let maxClassId = nextClasses.reduce(
            (max, c) => Math.max(max, c.id || 0),
            0
          );
          parsedStudents.forEach((st) => {
            const key = st.className.trim().toLowerCase();
            if (key && !clsNames.has(key)) {
              clsNames.add(key);
              maxClassId += 1;
              nextClasses.push({
                id: maxClassId,
                name: st.className,
                grade: st.grade,
                major: st.major,
                waliKelasName: '-',
                waliKelasNip: '-',
              });
            }
          });
          setClasses(nextClasses);
          setStudents(nextStudents);
          setStudentGrades(nextGrades);
          setBulkModalType(null);
          setBulkPasteText('');
          return {
            school,
            classes: nextClasses,
            students: nextStudents,
            users,
            studentGrades: nextGrades,
          };
        },
      });
    } else if (bulkModalType === 'walikelas_user') {
      const parsedWaliUsers = parseBulkWaliKelasUsers(rows);
      if (parsedWaliUsers.length === 0) {
        showToast(
          'Tidak ditemukan baris data user Wali Kelas yang valid.',
          'danger'
        );
        return;
      }
      setPendingMutation({
        title: `Impor Massal ${parsedWaliUsers.length} User Wali Kelas`,
        description:
          'Menambahkan akun Wali Kelas sekaligus memperbarui Nama & NIP/NUPTK Wali Kelas pada masing-masing Rombel secara otomatis.',
        affectedSummary: `${parsedWaliUsers.length} Wali Kelas (${parsedWaliUsers
          .slice(0, 3)
          .map(
            (u) =>
              `${formatTeacherNameWithTitle(u.name)} -> ${u.waliKelasFor}`
          )
          .join(', ')})`,
        executeLocalAndReturnPayload: () => {
          const byUsername = new Map<string, UserAccount>();
          users.forEach((u) => byUsername.set(u.username.toLowerCase(), u));
          parsedWaliUsers.forEach((pu) => {
            const existing = byUsername.get(pu.username.toLowerCase());
            byUsername.set(
              pu.username.toLowerCase(),
              existing ? { ...existing, ...pu, id: existing.id } : pu
            );
          });
          const nextUsers = Array.from(byUsername.values());

          // Sync Wali Kelas Name & NIP into classes
          const nextClasses = [...classes];
          parsedWaliUsers.forEach((pu) => {
            if (!pu.waliKelasFor) return;
            const clsIdx = nextClasses.findIndex(
              (c) =>
                c.name.trim().toLowerCase() ===
                pu.waliKelasFor.trim().toLowerCase()
            );
            if (clsIdx !== -1) {
              nextClasses[clsIdx] = {
                ...nextClasses[clsIdx],
                waliKelasName: pu.name,
                waliKelasNip: pu.nip,
              };
            } else {
              const grade = pu.waliKelasFor.startsWith('12')
                ? '12'
                : pu.waliKelasFor.startsWith('11')
                ? '11'
                : '10';
              const nextClsId =
                nextClasses.reduce((max, c) => Math.max(max, c.id || 0), 0) + 1;
              nextClasses.push({
                id: nextClsId,
                name: pu.waliKelasFor,
                grade,
                major: 'Teknik Kendaraan Ringan',
                waliKelasName: pu.name,
                waliKelasNip: pu.nip,
              });
            }
          });

          setUsers(nextUsers);
          setClasses(nextClasses);
          setBulkModalType(null);
          setBulkPasteText('');
          return {
            school,
            classes: nextClasses,
            students,
            users: nextUsers,
            studentGrades,
          };
        },
      });
    } else if (bulkModalType === 'gurumapel_user') {
      const parsedGuruUsers = parseBulkGuruMapelUsers(rows);
      if (parsedGuruUsers.length === 0) {
        showToast(
          'Tidak ditemukan baris data user Guru Mapel yang valid.',
          'danger'
        );
        return;
      }
      setPendingMutation({
        title: `Impor Massal ${parsedGuruUsers.length} User Guru Mapel`,
        description:
          'Menambahkan/memperbarui akun Guru Mata Pelajaran secara massal.',
        affectedSummary: `${parsedGuruUsers.length} Guru Mapel (${parsedGuruUsers
          .slice(0, 3)
          .map((u) => formatTeacherNameWithTitle(u.name))
          .join(', ')})`,
        executeLocalAndReturnPayload: () => {
          const byUsername = new Map<string, UserAccount>();
          users.forEach((u) => byUsername.set(u.username.toLowerCase(), u));
          parsedGuruUsers.forEach((pu) => {
            const existing = byUsername.get(pu.username.toLowerCase());
            byUsername.set(
              pu.username.toLowerCase(),
              existing ? { ...existing, ...pu, id: existing.id } : pu
            );
          });
          const nextUsers = Array.from(byUsername.values());
          setUsers(nextUsers);
          setBulkModalType(null);
          setBulkPasteText('');
          return {
            school,
            classes,
            students,
            users: nextUsers,
            studentGrades,
          };
        },
      });
    } else if (bulkModalType === 'nilai') {
      // Guru Mapel: Strictly Single Subject Import Only
      const allSubNames = availableSubjectsForSelectedClass.map((s) => s.name);
      const { updatedCount, gradeUpdates } = parseBulkGrades(
        rows,
        currentClassStudents,
        guruSelectedSubject,
        allSubNames,
        'single_subject'
      );
      if (updatedCount === 0) {
        showToast(
          `Tidak ditemukan kecocokan NISN/Nama siswa untuk kelas ${guruSelectedClass} pada mapel ${guruSelectedSubject}.`,
          'danger'
        );
        return;
      }
      setPendingMutation({
        title: `Impor Nilai Mapel ${guruSelectedSubject} (${updatedCount} Siswa - ${guruSelectedClass})`,
        description: `Memperbarui nilai mata pelajaran "${guruSelectedSubject}" hasil impor untuk ${updatedCount} siswa di kelas ${guruSelectedClass}.`,
        affectedSummary: `Kelas ${guruSelectedClass} • Mapel: ${guruSelectedSubject} • ${updatedCount} Siswa berhasil dicocokkan`,
        executeLocalAndReturnPayload: () => {
          const nextGrades: StudentGradesMap = { ...studentGrades };
          Object.entries(gradeUpdates).forEach(([stIdStr, subMap]) => {
            const stId = Number(stIdStr);
            nextGrades[stId] = {
              ...(nextGrades[stId] || {}),
              ...subMap,
            };
          });
          setStudentGrades(nextGrades);
          setBulkModalType(null);
          setBulkPasteText('');
          return {
            school,
            classes,
            students,
            users,
            studentGrades: nextGrades,
          };
        },
      });
    } else if (bulkModalType === 'nilai_admin_semua_mapel') {
      // Admin: Bulk Import All Subjects at Once for Selected Class
      const allSubNames = adminGradeClassSubjects.map((s) => s.name);
      const { updatedCount, gradeUpdates } = parseBulkGrades(
        rows,
        adminGradeClassStudents,
        allSubNames[0] || '',
        allSubNames,
        'all_subjects'
      );
      if (updatedCount === 0) {
        showToast(
          `Tidak ditemukan kecocokan NISN/Nama siswa untuk kelas ${adminGradeSelectedClass}. Pastikan memilih kelas yang sesuai.`,
          'danger'
        );
        return;
      }
      setPendingMutation({
        title: `Impor Massal Nilai Semua Mapel (${updatedCount} Siswa - Kelas ${adminGradeSelectedClass})`,
        description: `Memperbarui nilai seluruh mata pelajaran sekaligus untuk ${updatedCount} siswa di kelas ${adminGradeSelectedClass}.`,
        affectedSummary: `Kelas ${adminGradeSelectedClass} • Semua Mapel (${allSubNames.length} Mapel) • ${updatedCount} Siswa berhasil dicocokkan`,
        executeLocalAndReturnPayload: () => {
          const nextGrades: StudentGradesMap = { ...studentGrades };
          Object.entries(gradeUpdates).forEach(([stIdStr, subMap]) => {
            const stId = Number(stIdStr);
            nextGrades[stId] = {
              ...(nextGrades[stId] || {}),
              ...subMap,
            };
          });
          setStudentGrades(nextGrades);
          setBulkModalType(null);
          setBulkPasteText('');
          return {
            school,
            classes,
            students,
            users,
            studentGrades: nextGrades,
          };
        },
      });
    }
  };

  const handleBulkFileUpload = async (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const rows = await readSpreadsheetFileToRows(file);
      processBulkRows(rows);
    } catch (err: any) {
      showToast(
        `Gagal membaca file: ${err?.message || 'Format tidak didukung'}`,
        'danger'
      );
    } finally {
      e.target.value = '';
    }
  };

  const handleBulkPasteSubmit = () => {
    if (!bulkPasteText.trim()) {
      showToast('Tempelkan (Paste) data terlebih dahulu.', 'danger');
      return;
    }
    const rows = parsePastedTextToRows(bulkPasteText);
    processBulkRows(rows);
  };

  // Guru Mapel Grade Handlers
  const handleGradeChange = (
    studentId: number,
    subjectName: string,
    rawVal: string
  ) => {
    setStudentGrades((prev) => {
      const nextForStudent = { ...(prev[studentId] || {}) };
      if (rawVal === '') {
        delete nextForStudent[subjectName];
      } else {
        const num = Math.max(0, Math.min(100, Number(rawVal)));
        if (!Number.isNaN(num)) {
          nextForStudent[subjectName] = num;
        }
      }
      return { ...prev, [studentId]: nextForStudent };
    });
  };

  const fillRandomGrades = () => {
    setStudentGrades((prev) => {
      const next = { ...prev };
      currentClassStudents.forEach((st) => {
        const stGrades = { ...(next[st.id] || {}) };
        stGrades[guruSelectedSubject] = Math.floor(Math.random() * 21) + 75;
        next[st.id] = stGrades;
      });
      return next;
    });
    showToast(
      `Nilai sample otomatis diisi untuk mapel ${guruSelectedSubject}. Klik "Simpan Semua Nilai" untuk menyimpan.`,
      'success'
    );
  };

  const handleSaveGrades = () => {
    setPendingMutation({
      title: 'Simpan Nilai STS',
      description: `Anda akan memperbarui nilai mata pelajaran "${guruSelectedSubject}" untuk kelas ${guruSelectedClass}.`,
      affectedSummary: `${currentClassStudents.length} siswa di kelas ${guruSelectedClass} • Mapel: ${guruSelectedSubject}`,
      executeLocalAndReturnPayload: () => ({
        school,
        classes,
        students,
        users,
        studentGrades,
      }),
    });
  };

  // Admin All-Subjects Bulk Grade Handlers
  const fillAdminAllSubjectsRandomGrades = () => {
    setStudentGrades((prev) => {
      const next = { ...prev };
      adminGradeClassStudents.forEach((st) => {
        const stGrades = { ...(next[st.id] || {}) };
        adminGradeClassSubjects.forEach((sub) => {
          stGrades[sub.name] = Math.floor(Math.random() * 21) + 75;
        });
        next[st.id] = stGrades;
      });
      return next;
    });
    showToast(
      `Nilai sample otomatis diisi untuk seluruh mata pelajaran (${adminGradeClassSubjects.length} mapel) di kelas ${adminGradeSelectedClass}. Klik "Simpan Nilai Semua Mapel" untuk menyimpan.`,
      'success'
    );
  };

  const handleSaveAdminAllSubjectsGrades = () => {
    setPendingMutation({
      title: `Simpan Nilai Semua Mapel - Kelas ${adminGradeSelectedClass}`,
      description: `Anda akan menyimpan perubahan nilai seluruh mata pelajaran (${adminGradeClassSubjects.length} mapel) untuk kelas ${adminGradeSelectedClass}.`,
      affectedSummary: `${adminGradeClassStudents.length} Siswa di kelas ${adminGradeSelectedClass} • ${adminGradeClassSubjects.length} Mata Pelajaran`,
      executeLocalAndReturnPayload: () => ({
        school,
        classes,
        students,
        users,
        studentGrades,
      }),
    });
  };

  // Wali Kelas Attendance & Note Handlers
  const handleAttendanceChange = (
    studentId: number,
    field: 'sakit' | 'izin' | 'alpa',
    rawVal: string
  ) => {
    const val = Math.max(0, Number(rawVal) || 0);
    setStudents((prev) =>
      prev.map((s) =>
        s.id === studentId
          ? { ...s, attendance: { ...s.attendance, [field]: val } }
          : s
      )
    );
  };

  const handleNoteChange = (studentId: number, note: string) => {
    setStudents((prev) =>
      prev.map((s) => (s.id === studentId ? { ...s, note } : s))
    );
  };

  const handleSaveWaliData = () => {
    setPendingMutation({
      title: 'Simpan Kehadiran & Catatan Wali Kelas',
      description: `Anda akan menyimpan perubahan data absensi (Sakit, Izin, Alpa) dan catatan wali kelas untuk rombel ${waliSelectedClass}.`,
      affectedSummary: `${currentClassStudents.length} siswa pada kelas ${waliSelectedClass}`,
      executeLocalAndReturnPayload: () => ({
        school,
        classes,
        students,
        users,
        studentGrades,
      }),
    });
  };

  // Full Database Sync (Admin Only)
  const handleManualSyncToSheets = () => {
    if (!currentUser?.roles.includes('admin') || activeRole !== 'admin') {
      showToast(
        'Akses ditolak: Hanya Administrator yang dapat mengakses database.',
        'danger'
      );
      return;
    }
    setPendingMutation({
      title: 'Sinkronkan Seluruh Data ke Google Sheets',
      description:
        'Tindakan ini akan menulis ulang dan menyinkronkan seluruh data aplikasi (Sekolah, Kelas, Siswa, Pengguna, Nilai, dan Master Leger gid=0) ke Google Sheets.',
      affectedSummary: `Spreadsheet ID: ${activeSheetId} • ${classes.length} Kelas, ${students.length} Siswa, ${users.length} Pengguna`,
      executeLocalAndReturnPayload: () => ({
        school,
        classes,
        students,
        users,
        studentGrades,
      }),
    });
  };

  const toggleNewUserRole = (role: RoleType) => {
    setNewUser((prev) => {
      const exists = prev.roles.includes(role);
      const nextRoles = exists
        ? prev.roles.filter((r) => r !== role)
        : [...prev.roles, role];
      return { ...prev, roles: nextRoles };
    });
  };

  const openSingleStudentPreview = (student: RankedStudentItem) => {
    setPrintMode('single');
    setSelectedPrintStudent(student);
    setShowPrintModal(true);
  };

  const openClassPrintPreview = () => {
    setPrintMode('class');
    setShowPrintModal(true);
  };

  const triggerBrowserPrint = () => {
    window.print();
  };

  const handleDownloadPdf = async () => {
    if (!previewContainerRef.current) return;
    setIsGeneratingPdf(true);
    try {
      const filename =
        printMode === 'single' && selectedPrintStudent
          ? `Raport_STS_${selectedPrintStudent.name.replace(/\s+/g, '_')}.pdf`
          : `Raport_STS_Kelas_${waliSelectedClass.replace(/\s+/g, '_')}.pdf`;
      await generateAndDownloadPdf(previewContainerRef.current, filename);
      showToast(`File PDF (${filename}) berhasil diunduh!`, 'success');
    } catch (err: any) {
      console.error('PDF Download error:', err);
      showToast(
        `Gagal mengunduh PDF: ${err?.message || 'Silakan coba lagi'}`,
        'danger'
      );
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  const activeWaliClassInfo = getClassInfo(waliSelectedClass);

  // Comprehensive Visual Statistics for Admin Dashboard
  const adminStats = useMemo(() => {
    // 1. Teacher / Guru Statistics
    const teacherUsers = users.filter(
      (u) => u.roles.includes('gurumapel') || u.roles.includes('walikelas')
    );
    const guruMapelCount = teacherUsers.filter((u) =>
      u.roles.includes('gurumapel')
    ).length;
    const waliKelasUserCount = teacherUsers.filter((u) =>
      u.roles.includes('walikelas')
    ).length;
    const dualRoleGuruCount = teacherUsers.filter(
      (u) => u.roles.includes('gurumapel') && u.roles.includes('walikelas')
    ).length;

    // Combine unique teachers from users + any Wali Kelas assigned on classes not in users
    const uniqueTeacherMap = new Map<
      string,
      {
        name: string;
        formattedName: string;
        idInfo: ReturnType<typeof formatNipOrNuptk>;
        roles: RoleType[];
        waliKelasFor: string;
      }
    >();

    teacherUsers.forEach((u) => {
      const formattedName = formatTeacherNameWithTitle(u.name);
      const key = formattedName.toLowerCase();
      uniqueTeacherMap.set(key, {
        name: u.name,
        formattedName,
        idInfo: formatNipOrNuptk(u.nip),
        roles: u.roles,
        waliKelasFor: u.waliKelasFor || '',
      });
    });

    classes.forEach((c) => {
      const resolved = getClassInfo(c.name);
      const formattedName = formatTeacherNameWithTitle(resolved.waliKelasName);
      if (formattedName && formattedName !== '-') {
        const key = formattedName.toLowerCase();
        if (!uniqueTeacherMap.has(key)) {
          uniqueTeacherMap.set(key, {
            name: resolved.waliKelasName || '',
            formattedName,
            idInfo: formatNipOrNuptk(resolved.waliKelasNip),
            roles: ['walikelas'],
            waliKelasFor: c.name,
          });
        }
      }
    });

    const allTeachersList = Array.from(uniqueTeacherMap.values());
    const totalGuruCount = allTeachersList.length;
    const guruWithNipCount = allTeachersList.filter(
      (t) => t.idInfo.type === 'NIP'
    ).length;
    const guruWithNuptkCount = allTeachersList.filter(
      (t) => t.idInfo.type === 'NUPTK'
    ).length;
    const guruWithoutIdCount = allTeachersList.filter(
      (t) => t.idInfo.type === 'NONE'
    ).length;

    // 2. Class & Per-Class Student + Grade Statistics
    let totalExpectedGradesGlobal = 0;
    let totalFilledGradesGlobal = 0;
    let sumAllScoresGlobal = 0;
    let studentsCompleteGlobal = 0;
    let studentsPartialGlobal = 0;
    let studentsEmptyGlobal = 0;

    const classBreakdown = classes.map((c) => {
      const resolved = getClassInfo(c.name);
      const formattedWali = formatTeacherNameWithTitle(resolved.waliKelasName);
      const hasWali = formattedWali !== '-';
      const waliIdInfo = formatNipOrNuptk(resolved.waliKelasNip);

      const studentsInCls = students.filter(
        (s) => s.className.trim().toLowerCase() === c.name.trim().toLowerCase()
      );
      const stCount = studentsInCls.length;
      const subTemplate = getSubjectsForGrade(c.grade);
      const allSubjectsForCls = [
        ...subTemplate.umum.map((s) => ({ name: s.name, category: 'Umum' })),
        ...subTemplate.kejuruan.map((s) => ({
          name: s.name,
          category: 'Kejuruan',
        })),
      ];
      const subjectsCount = allSubjectsForCls.length;
      const expectedGrades = stCount * subjectsCount;

      let filledGrades = 0;
      let sumClsScores = 0;
      let completeStudentsInCls = 0;

      const rankedInCls: Array<
        RankedStudentItem & {
          filledSubjectsCount: number;
          subjectsCount: number;
          bestSubjectName: string;
          bestSubjectScore: number;
          predicateLabel: string;
          predicateColor: string;
        }
      > = [];

      studentsInCls.forEach((st) => {
        const stGradeObj = studentGrades[st.id] || {};
        let stFilled = 0;
        let stSum = 0;
        let bestSubName = '-';
        let bestSubScore = -1;

        allSubjectsForCls.forEach((sub) => {
          const val = stGradeObj[sub.name];
          if (
            val !== undefined &&
            String(val) !== '' &&
            !Number.isNaN(Number(val))
          ) {
            const numVal = Number(val);
            filledGrades++;
            stFilled++;
            sumClsScores += numVal;
            stSum += numVal;
            if (numVal > bestSubScore) {
              bestSubScore = numVal;
              bestSubName = sub.name;
            }
          }
        });
        if (subjectsCount > 0 && stFilled === subjectsCount) {
          completeStudentsInCls++;
          studentsCompleteGlobal++;
        } else if (stFilled > 0) {
          studentsPartialGlobal++;
        } else {
          studentsEmptyGlobal++;
        }

        const stAvg =
          stFilled > 0 ? Number((stSum / stFilled).toFixed(2)) : 0;
        let predicateLabel = 'Belum Ada Nilai';
        let predicateColor = 'bg-slate-100 text-slate-600';
        if (stFilled > 0) {
          if (stAvg >= 85) {
            predicateLabel = 'Sangat Baik (A)';
            predicateColor = 'bg-emerald-100 text-emerald-800';
          } else if (stAvg >= 75) {
            predicateLabel = 'Baik (B)';
            predicateColor = 'bg-blue-100 text-blue-800';
          } else if (stAvg >= 65) {
            predicateLabel = 'Cukup (C)';
            predicateColor = 'bg-amber-100 text-amber-800';
          } else {
            predicateLabel = 'Perlu Bimbingan';
            predicateColor = 'bg-red-100 text-red-800';
          }
        }

        rankedInCls.push({
          ...st,
          totalScore: stSum,
          avgScore: stAvg,
          rank: 0,
          filledSubjectsCount: stFilled,
          subjectsCount,
          bestSubjectName: bestSubName,
          bestSubjectScore: bestSubScore >= 0 ? bestSubScore : 0,
          predicateLabel,
          predicateColor,
        });
      });

      rankedInCls.sort((a, b) => {
        if (b.avgScore !== a.avgScore) return b.avgScore - a.avgScore;
        if (b.totalScore !== a.totalScore) return b.totalScore - a.totalScore;
        if (b.filledSubjectsCount !== a.filledSubjectsCount) {
          return b.filledSubjectsCount - a.filledSubjectsCount;
        }
        return a.name.localeCompare(b.name);
      });

      const top10Students = rankedInCls.slice(0, 10).map((item, idx) => ({
        ...item,
        rank: idx + 1,
      }));

      totalExpectedGradesGlobal += expectedGrades;
      totalFilledGradesGlobal += filledGrades;
      sumAllScoresGlobal += sumClsScores;

      const progressPercent =
        expectedGrades > 0
          ? Math.round((filledGrades / expectedGrades) * 100)
          : 0;
      const avgScore =
        filledGrades > 0 ? Number((sumClsScores / filledGrades).toFixed(1)) : 0;
      const studentSharePercent =
        students.length > 0 ? Math.round((stCount / students.length) * 100) : 0;

      return {
        id: c.id,
        name: c.name,
        grade: c.grade,
        major: c.major,
        hasWali,
        waliName: hasWali ? formattedWali : 'Belum Ditentukan',
        waliIdFormatted: waliIdInfo.formatted,
        waliIdType: waliIdInfo.type,
        studentCount: stCount,
        studentSharePercent,
        subjectsCount,
        expectedGrades,
        filledGrades,
        emptyGrades: Math.max(0, expectedGrades - filledGrades),
        progressPercent,
        avgScore,
        completeStudentsInCls,
        top10Students,
      };
    });

    const classesWithWaliCount = classBreakdown.filter((c) => c.hasWali).length;

    // 3. Students by Grade Level (10, 11, 12)
    const studentsByGrade = {
      '10': students.filter((s) => s.grade === '10').length,
      '11': students.filter((s) => s.grade === '11').length,
      '12': students.filter((s) => s.grade === '12').length,
    };

    // 4. Subject-level Grade Input Statistics (Filtered by adminStatFilterClass or Global)
    const targetClassesForSubStat = adminStatFilterClass
      ? classes.filter(
          (c) =>
            c.name.trim().toLowerCase() ===
            adminStatFilterClass.trim().toLowerCase()
        )
      : classes;

    const subjectStatMap = new Map<
      string,
      {
        name: string;
        category: 'Umum' | 'Kejuruan';
        targetStudents: number;
        filledStudents: number;
        sumScores: number;
      }
    >();

    targetClassesForSubStat.forEach((cls) => {
      const stsInCls = students.filter(
        (s) =>
          s.className.trim().toLowerCase() === cls.name.trim().toLowerCase()
      );
      const subTemplate = getSubjectsForGrade(cls.grade);
      const clsSubjects: Array<{ name: string; category: 'Umum' | 'Kejuruan' }> =
        [
          ...subTemplate.umum.map((s) => ({
            name: s.name,
            category: 'Umum' as const,
          })),
          ...subTemplate.kejuruan.map((s) => ({
            name: s.name,
            category: 'Kejuruan' as const,
          })),
        ];

      clsSubjects.forEach((sub) => {
        const existing = subjectStatMap.get(sub.name) || {
          name: sub.name,
          category: sub.category,
          targetStudents: 0,
          filledStudents: 0,
          sumScores: 0,
        };
        existing.targetStudents += stsInCls.length;
        stsInCls.forEach((st) => {
          const val = studentGrades[st.id]?.[sub.name];
          if (
            val !== undefined &&
            String(val) !== '' &&
            !Number.isNaN(Number(val))
          ) {
            existing.filledStudents += 1;
            existing.sumScores += Number(val);
          }
        });
        subjectStatMap.set(sub.name, existing);
      });
    });

    const subjectStatsList = Array.from(subjectStatMap.values()).map((item) => {
      const progressPercent =
        item.targetStudents > 0
          ? Math.round((item.filledStudents / item.targetStudents) * 100)
          : 0;
      const avgScore =
        item.filledStudents > 0
          ? Number((item.sumScores / item.filledStudents).toFixed(1))
          : 0;
      return {
        ...item,
        emptyStudents: Math.max(0, item.targetStudents - item.filledStudents),
        progressPercent,
        avgScore,
      };
    });

    const globalProgressPercent =
      totalExpectedGradesGlobal > 0
        ? Math.round((totalFilledGradesGlobal / totalExpectedGradesGlobal) * 100)
        : 0;
    const globalAvgScore =
      totalFilledGradesGlobal > 0
        ? Number((sumAllScoresGlobal / totalFilledGradesGlobal).toFixed(1))
        : 0;

    return {
      totalGuruCount,
      guruMapelCount,
      waliKelasUserCount,
      dualRoleGuruCount,
      guruWithNipCount,
      guruWithNuptkCount,
      guruWithoutIdCount,
      allTeachersList,
      totalClasses: classes.length,
      classesWithWaliCount,
      classBreakdown,
      totalStudentsGlobal: students.length,
      studentsByGrade,
      totalExpectedGradesGlobal,
      totalFilledGradesGlobal,
      totalEmptyGradesGlobal: Math.max(
        0,
        totalExpectedGradesGlobal - totalFilledGradesGlobal
      ),
      globalProgressPercent,
      globalAvgScore,
      studentsCompleteGlobal,
      studentsPartialGlobal,
      studentsEmptyGlobal,
      subjectStatsList,
    };
  }, [
    users,
    classes,
    students,
    studentGrades,
    subjectTemplates,
    adminStatFilterClass,
  ]);

  return (
    <div className="min-h-screen flex flex-col bg-[#f4f6f9] text-slate-800">
      {/* TOAST NOTIFICATION */}
      {toast.show && (
        <div className="fixed bottom-4 right-4 z-50 no-print max-w-md">
          <div
            className={`flex items-center justify-between gap-3 px-4 py-3 rounded-lg shadow-lg text-white text-sm font-medium ${
              toast.type === 'success' ? 'bg-emerald-600' : 'bg-red-600'
            }`}
          >
            <div className="flex items-center gap-2">
              <i
                className={
                  toast.type === 'success'
                    ? 'fas fa-check-circle'
                    : 'fas fa-exclamation-circle'
                }
              ></i>
              <span>{toast.message}</span>
            </div>
            <button
              type="button"
              onClick={() => setToast((prev) => ({ ...prev, show: false }))}
              className="text-white/80 hover:text-white cursor-pointer"
            >
              <i className="fas fa-times"></i>
            </button>
          </div>
        </div>
      )}

      {/* MANDATORY CONFIRMATION MODAL FOR MUTATIONS */}
      {pendingMutation && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/70 p-4 no-print">
          <div className="bg-white rounded-xl shadow-xl border border-slate-200 max-w-lg w-full overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between">
              <h5 className="font-bold text-slate-900 text-base flex items-center gap-2">
                <i
                  className={
                    pendingMutation.isDestructive
                      ? 'fas fa-exclamation-triangle text-red-600'
                      : 'fas fa-check-circle text-emerald-600'
                  }
                ></i>
                {pendingMutation.title}
              </h5>
              <button
                type="button"
                onClick={() => setPendingMutation(null)}
                className="text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <i className="fas fa-times"></i>
              </button>
            </div>
            <div className="p-6 space-y-3">
              <p className="text-sm text-slate-600 leading-relaxed">
                {pendingMutation.description}
              </p>
              <div className="p-3 rounded-lg bg-slate-50 border border-slate-200 text-xs font-mono text-slate-700">
                <div className="font-semibold text-slate-900 mb-1">
                  Detail Perubahan Data:
                </div>
                <div>{pendingMutation.affectedSummary}</div>
              </div>
              {needsGoogleAuth &&
                activeRole === 'admin' &&
                currentUser?.roles.includes('admin') && (
                  <div className="p-3 rounded-lg bg-amber-50 border border-amber-200 text-xs text-amber-800 flex items-center justify-between gap-2">
                    <span>
                      Akun Google belum terhubung. Hubungkan agar perubahan
                      langsung tertulis ke Google Sheets.
                    </span>
                    <button
                      type="button"
                      onClick={handleGoogleConnect}
                      className="px-2.5 py-1 bg-amber-600 text-white font-semibold rounded hover:bg-amber-700 shrink-0 cursor-pointer"
                    >
                      Hubungkan
                    </button>
                  </div>
                )}
            </div>
            <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setPendingMutation(null)}
                className="px-4 py-2 rounded-lg text-sm font-medium text-slate-600 hover:bg-slate-200 transition-colors cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={confirmAndExecuteMutation}
                className={`px-4 py-2 rounded-lg text-sm font-semibold text-white transition-colors cursor-pointer ${
                  pendingMutation.isDestructive
                    ? 'bg-red-600 hover:bg-red-700'
                    : 'bg-blue-600 hover:bg-blue-700'
                }`}
              >
                {pendingMutation.isDestructive
                  ? 'Konfirmasi & Hapus'
                  : 'Konfirmasi & Simpan'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* LOGIN SCREEN (CLEAN, NO DEMO ACCOUNTS, NO DATABASE LINK) */}
      {!isLoggedIn && !showRoleModal && (
        <div className="min-h-screen flex flex-col items-center justify-center bg-slate-100 py-8 px-4 no-print">
          <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden max-w-4xl w-full grid grid-cols-1 md:grid-cols-12">
            {/* Left Hero Section */}
            <div
              className="md:col-span-5 text-white p-6 md:p-8 flex flex-col justify-between"
              style={{
                background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
              }}
            >
              <div>
                <div className="flex items-center gap-3 mb-6">
                  <div className="bg-white/10 border border-white/20 rounded-xl p-1.5 flex items-center justify-center w-11 h-11 shrink-0 overflow-hidden">
                    {!logoImgError && school.logoUrl ? (
                      <img
                        src={school.logoUrl}
                        alt="Logo SMKN 1 Mesuji Timur"
                        className="w-full h-full object-contain"
                        referrerPolicy="no-referrer"
                        onError={() => setLogoImgError(true)}
                      />
                    ) : (
                      <i className="fas fa-graduation-cap fa-lg text-blue-400"></i>
                    )}
                  </div>
                  <h5 className="font-bold text-lg mb-0">E-Raport STS</h5>
                </div>
                <h3 className="font-bold text-2xl mb-2">SMKN 1 Mesuji Timur</h3>
                <p className="text-slate-300 text-sm leading-relaxed">
                  Sistem Laporan Hasil Belajar Sumatif Tengah Semester Kurikulum
                  Merdeka.
                </p>
              </div>

              <div className="mt-8 pt-6 border-t border-slate-700 space-y-2">
                <div>
                  <div className="text-xs text-slate-400">
                    <i className="fas fa-info-circle mr-1.5"></i>
                    Informasi Akses
                  </div>
                  <div className="text-xs text-slate-200 pt-1">
                    Tahun Pelajaran ·{' '}
                    <span className="font-semibold text-white">
                      {school.tp}
                    </span>
                  </div>
                </div>
                <div className="pt-3 border-t border-slate-800 text-[11px] text-slate-400">
                  © 2026 Desain & Sistem Dibuat oleh{' '}
                  <span className="font-bold text-amber-400">
                    Imam Huzaer, S.Pd.
                  </span>
                </div>
              </div>
            </div>

            {/* Right Form Section */}
            <div className="md:col-span-7 p-6 md:p-8 bg-white flex flex-col justify-between">
              <div className="my-auto">
                <div className="mb-6">
                  <h4 className="font-bold text-xl text-slate-900 mb-1">
                    Masuk ke Sistem
                  </h4>
                  <p className="text-slate-500 text-sm">
                    Gunakan username dan kata sandi akun Anda untuk mengakses
                    e-raport.
                  </p>
                </div>

                {/* Login Form */}
                <form onSubmit={handleLogin} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-600 mb-1">
                      Username / NIP
                    </label>
                    <div className="flex items-center border border-slate-300 rounded-lg bg-slate-50 focus-within:bg-white focus-within:border-blue-600 overflow-hidden">
                      <span className="px-3 text-slate-400">
                        <i className="fas fa-user"></i>
                      </span>
                      <input
                        type="text"
                        value={loginForm.username}
                        onChange={(e) =>
                          setLoginForm((prev) => ({
                            ...prev,
                            username: e.target.value,
                          }))
                        }
                        className="w-full py-2.5 pr-3 bg-transparent text-sm outline-none"
                        placeholder="Masukkan username"
                        required
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-600 mb-1">
                      Kata Sandi
                    </label>
                    <div className="flex items-center border border-slate-300 rounded-lg bg-slate-50 focus-within:bg-white focus-within:border-blue-600 overflow-hidden">
                      <span className="px-3 text-slate-400">
                        <i className="fas fa-lock"></i>
                      </span>
                      <input
                        type={showLoginPassword ? 'text' : 'password'}
                        value={loginForm.password}
                        onChange={(e) =>
                          setLoginForm((prev) => ({
                            ...prev,
                            password: e.target.value,
                          }))
                        }
                        className="w-full py-2.5 pr-2 bg-transparent text-sm outline-none"
                        placeholder="Masukkan password"
                        required
                      />
                      <button
                        type="button"
                        onClick={() =>
                          setShowLoginPassword((prev) => !prev)
                        }
                        title={
                          showLoginPassword
                            ? 'Sembunyikan Kata Sandi'
                            : 'Lihat Kata Sandi'
                        }
                        aria-label={
                          showLoginPassword
                            ? 'Sembunyikan Kata Sandi'
                            : 'Lihat Kata Sandi'
                        }
                        className="px-3 py-2 text-slate-400 hover:text-slate-700 transition-colors cursor-pointer"
                      >
                        <i
                          className={
                            showLoginPassword
                              ? 'fas fa-eye-slash'
                              : 'fas fa-eye'
                          }
                        ></i>
                      </button>
                    </div>
                  </div>

                  <div className="flex items-center justify-between pt-0.5">
                    <label className="flex items-center gap-2 text-xs text-slate-600 font-medium cursor-pointer select-none">
                      <input
                        type="checkbox"
                        checked={rememberLogin}
                        onChange={(e) => setRememberLogin(e.target.checked)}
                        className="rounded border-slate-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                      />
                      <span>
                        Simpan login di browser ini (Tidak login ulang saat
                        refresh)
                      </span>
                    </label>
                  </div>

                  <button
                    type="submit"
                    className="w-full py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm rounded-lg transition-colors cursor-pointer"
                  >
                    <i className="fas fa-sign-in-alt mr-1.5"></i> Masuk Sekarang
                  </button>
                </form>
              </div>

              <div className="mt-6 pt-4 border-t border-slate-100 text-center text-xs text-slate-500">
                Mandatory Copyright © 2026 · Desain & Aplikasi Dibuat oleh{' '}
                <b className="text-slate-800">Imam Huzaer, S.Pd.</b>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* DUAL ROLE SELECTION MODAL */}
      {showRoleModal && (
        <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/75 p-4 no-print">
          <div className="bg-white rounded-2xl shadow-xl max-w-md w-full p-6">
            <div className="mb-4">
              <div className="text-xs font-bold text-amber-600 uppercase tracking-wider mb-1">
                Role Ganda Terdeteksi
              </div>
              <h5 className="font-bold text-lg text-slate-900">
                Pilih Akses Masuk
              </h5>
              <p className="text-slate-500 text-xs mt-1">
                Selamat datang,{' '}
                <b>{formatTeacherNameWithTitle(currentUser?.name)}</b>! Anda
                memiliki lebih dari satu hak akses. Silakan pilih role yang
                ingin digunakan saat ini:
              </p>
            </div>

            <div className="space-y-2.5">
              {currentUser?.roles.includes('walikelas') && (
                <button
                  type="button"
                  onClick={() => selectRole('walikelas')}
                  className="w-full p-3.5 text-left flex items-center justify-between rounded-xl border-2 border-emerald-600/30 hover:border-emerald-600 hover:bg-emerald-50/50 transition-colors cursor-pointer"
                >
                  <div>
                    <div className="font-bold text-sm text-emerald-800">
                      <i className="fas fa-user-check mr-2"></i> Wali Kelas (
                      {currentUser?.waliKelasFor || 'Kelas'})
                    </div>
                    <div className="text-xs text-slate-500 mt-0.5">
                      Input absensi, catatan wali kelas, dan cetak raport A4.
                    </div>
                  </div>
                  <i className="fas fa-chevron-right text-emerald-600"></i>
                </button>
              )}

              {currentUser?.roles.includes('gurumapel') && (
                <button
                  type="button"
                  onClick={() => selectRole('gurumapel')}
                  className="w-full p-3.5 text-left flex items-center justify-between rounded-xl border-2 border-blue-600/30 hover:border-blue-600 hover:bg-blue-50/50 transition-colors cursor-pointer"
                >
                  <div>
                    <div className="font-bold text-sm text-blue-800">
                      <i className="fas fa-chalkboard-teacher mr-2"></i> Guru
                      Mata Pelajaran
                    </div>
                    <div className="text-xs text-slate-500 mt-0.5">
                      Input nilai Sumatif Tengah Semester (STS) siswa.
                    </div>
                  </div>
                  <i className="fas fa-chevron-right text-blue-600"></i>
                </button>
              )}

              {currentUser?.roles.includes('admin') && (
                <button
                  type="button"
                  onClick={() => selectRole('admin')}
                  className="w-full p-3.5 text-left flex items-center justify-between rounded-xl border-2 border-red-600/30 hover:border-red-600 hover:bg-red-50/50 transition-colors cursor-pointer"
                >
                  <div>
                    <div className="font-bold text-sm text-red-800">
                      <i className="fas fa-user-shield mr-2"></i> Administrator
                    </div>
                    <div className="text-xs text-slate-500 mt-0.5">
                      Kelola master data sekolah, siswa, kelas, guru, dan
                      pengguna.
                    </div>
                  </div>
                  <i className="fas fa-chevron-right text-red-600"></i>
                </button>
              )}
            </div>

            <div className="mt-5 pt-3 border-t border-slate-100">
              <button
                type="button"
                onClick={handleLogout}
                className="w-full py-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold text-slate-600 cursor-pointer"
              >
                Batal & Keluar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MAIN APPLICATION CONTAINER (VISIBLE ONLY WHEN LOGGED IN) */}
      {isLoggedIn && !showRoleModal && (
        <>
          {/* NAVBAR & HEADER (NO PRINT, WITHOUT PANDUAN GAS MENU) */}
          <header className="no-print bg-slate-900 text-white shadow-sm sticky top-0 z-30">
            <div className="max-w-7xl mx-auto px-4 py-2.5 flex flex-wrap items-center justify-between gap-4">
              <div className="flex items-center gap-3">
                <div className="bg-white/10 border border-white/20 rounded-xl p-1 flex items-center justify-center w-10 h-10 shrink-0 overflow-hidden">
                  {!logoImgError && school.logoUrl ? (
                    <img
                      src={school.logoUrl}
                      alt="Logo SMKN 1 Mesuji Timur"
                      className="w-full h-full object-contain"
                      referrerPolicy="no-referrer"
                      onError={() => setLogoImgError(true)}
                    />
                  ) : (
                    <i className="fas fa-graduation-cap fa-lg text-blue-400"></i>
                  )}
                </div>
                <div>
                  <h5 className="font-bold text-base mb-0 text-white whitespace-nowrap">
                    E-Raport STS
                  </h5>
                  <div className="text-xs text-slate-400">
                    SMK Negeri 1 Mesuji Timur · TP {school.tp}
                  </div>
                </div>
              </div>

              {/* Right Controls: Google Sheets Sync (ADMIN ONLY) + Active User & Role Switcher */}
              <div className="flex flex-wrap items-center gap-2.5">
                {activeRole === 'admin' &&
                  currentUser?.roles.includes('admin') && (
                    <div className="flex flex-wrap items-center gap-2">
                      {needsGoogleAuth && (
                        <GoogleSignInButton
                          onClick={handleGoogleConnect}
                          disabled={isSyncingSheets}
                        />
                      )}
                      <div className="flex items-center gap-1.5 bg-slate-800 px-2.5 py-1.5 rounded-lg border border-slate-700">
                        <button
                          type="button"
                          onClick={() => loadFromGoogleSheets()}
                          disabled={isSyncingSheets}
                          className="text-xs font-semibold text-emerald-400 hover:text-emerald-300 flex items-center gap-1.5 px-1.5 cursor-pointer whitespace-nowrap"
                          title="Baca data terbaru langsung dari Google Spreadsheet (Tanpa / Dengan OAuth)"
                        >
                          <i
                            className={`fas fa-sync-alt ${
                              isSyncingSheets ? 'fa-spin' : ''
                            }`}
                          ></i>
                          <span>Muat Sheet</span>
                        </button>
                        <span className="text-slate-600">·</span>
                        <button
                          type="button"
                          onClick={handleManualSyncToSheets}
                          disabled={isSyncingSheets}
                          className="text-xs font-semibold text-blue-400 hover:text-blue-300 flex items-center gap-1 px-1.5 cursor-pointer whitespace-nowrap"
                          title="Simpan seluruh perubahan ke Google Spreadsheet"
                        >
                          <i className="fas fa-cloud-upload-alt"></i>
                          <span>Simpan ke Sheet</span>
                        </button>
                      </div>
                    </div>
                  )}

                <div className="bg-slate-800 px-3 py-1.5 rounded-lg border border-slate-700 flex items-center gap-2.5">
                  <div className="text-right hidden sm:block">
                    <div className="font-bold text-xs text-white">
                      {formatTeacherNameWithTitle(currentUser?.name)}
                    </div>
                    <div className="text-[11px] text-slate-400">
                      Role Aktif:{' '}
                      <span className="text-amber-400 font-semibold uppercase">
                        {activeRole}
                      </span>
                    </div>
                  </div>

                  {currentUser && currentUser.roles.length > 1 && (
                    <button
                      type="button"
                      onClick={() => setShowRoleModal(true)}
                      className="px-2.5 py-1 rounded border border-amber-400/60 text-amber-300 hover:bg-amber-400 hover:text-slate-900 text-xs font-bold transition-colors cursor-pointer whitespace-nowrap"
                      title="Ganti Mode Role"
                    >
                      <i className="fas fa-exchange-alt mr-1"></i> Switch
                    </button>
                  )}
                </div>

                <button
                  type="button"
                  onClick={handleLogout}
                  className="px-3 py-1.5 rounded-lg border border-red-500/60 text-red-400 hover:bg-red-600 hover:text-white text-xs font-bold transition-colors cursor-pointer whitespace-nowrap"
                  title="Keluar"
                >
                  <i className="fas fa-sign-out-alt mr-1"></i> Logout
                </button>
              </div>
            </div>
          </header>

          {/* MAIN CONTENT */}
          <main className="flex-grow max-w-7xl w-full mx-auto px-4 py-6 no-print">
            {/* ================= ROLE: ADMIN ================= */}
            {activeRole === 'admin' && (
              <div className="space-y-6">
                {/* Admin Header Card */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 flex flex-wrap justify-between items-center gap-4">
                  <div>
                    <div className="text-xs font-semibold text-blue-600 mb-1">
                      Panel Administrator · Terhubung Google Sheets
                    </div>
                    <h4 className="font-bold text-xl text-slate-900 mb-1">
                      Pengaturan Master Data & Kurikulum
                    </h4>
                    <p className="text-slate-500 text-xs">
                      Kelola identitas sekolah, rombel, guru, pengguna & role,
                      serta impor template cepat dan sinkronisasi database.
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={handleManualSyncToSheets}
                      disabled={isSyncingSheets}
                      className="px-3.5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition-colors cursor-pointer whitespace-nowrap"
                    >
                      <i className="fas fa-cloud-upload-alt mr-1.5"></i>
                      Sinkronkan ke Google Sheets
                    </button>
                  </div>
                </div>

                {/* ADMIN TOP VISUAL KPI SUMMARY RIBBON */}
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  {/* KPI 1: Statistik Nilai yang Sudah Diinput */}
                  <div
                    onClick={() => setAdminTab('statistik')}
                    className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 hover:border-blue-400 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                        Nilai Sudah Diinput
                      </span>
                      <span className="w-8 h-8 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center">
                        <i className="fas fa-chart-pie"></i>
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-2">
                      <div className="text-2xl font-extrabold text-slate-900 font-mono">
                        {adminStats.globalProgressPercent}%
                      </div>
                      <div className="text-xs font-mono text-slate-600">
                        <b>{adminStats.totalFilledGradesGlobal}</b> /{' '}
                        {adminStats.totalExpectedGradesGlobal} Nilai
                      </div>
                    </div>
                    <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden mt-2.5">
                      <div
                        className={`h-full rounded-full transition-all duration-500 ${
                          adminStats.globalProgressPercent === 100
                            ? 'bg-emerald-500'
                            : adminStats.globalProgressPercent >= 50
                            ? 'bg-blue-600'
                            : 'bg-amber-500'
                        }`}
                        style={{
                          width: `${Math.min(
                            100,
                            adminStats.globalProgressPercent
                          )}%`,
                        }}
                      ></div>
                    </div>
                    <div className="flex items-center justify-between text-[11px] text-slate-500 mt-2">
                      <span>
                        Rata-rata Global: <b>{adminStats.globalAvgScore}</b>
                      </span>
                      <span>
                        Belum: <b>{adminStats.totalEmptyGradesGlobal}</b>
                      </span>
                    </div>
                  </div>

                  {/* KPI 2: Jumlah Guru */}
                  <div
                    onClick={() => setAdminTab('guru')}
                    className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 hover:border-indigo-400 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                        Jumlah Guru
                      </span>
                      <span className="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center">
                        <i className="fas fa-chalkboard-teacher"></i>
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-2">
                      <div className="text-2xl font-extrabold text-slate-900 font-mono">
                        {adminStats.totalGuruCount}{' '}
                        <span className="text-xs font-sans font-semibold text-slate-500">
                          Guru
                        </span>
                      </div>
                      <div className="text-xs text-indigo-700 font-semibold">
                        {adminStats.guruMapelCount} Mapel ·{' '}
                        {adminStats.waliKelasUserCount} Wali
                      </div>
                    </div>
                    <div className="mt-2.5 pt-2 border-t border-slate-100 flex flex-wrap items-center justify-between gap-1 text-[11px] text-slate-600">
                      <span>
                        NIP: <b>{adminStats.guruWithNipCount}</b>
                      </span>
                      <span>
                        NUPTK: <b>{adminStats.guruWithNuptkCount}</b>
                      </span>
                      <span>
                        Hanya Nama: <b>{adminStats.guruWithoutIdCount}</b>
                      </span>
                    </div>
                  </div>

                  {/* KPI 3: Jumlah Kelas & Wali Kelas */}
                  <div
                    onClick={() => setAdminTab('kelas')}
                    className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 hover:border-emerald-400 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                        Kelas & Wali Kelas
                      </span>
                      <span className="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center">
                        <i className="fas fa-door-open"></i>
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-2">
                      <div className="text-2xl font-extrabold text-slate-900 font-mono">
                        {adminStats.totalClasses}{' '}
                        <span className="text-xs font-sans font-semibold text-slate-500">
                          Rombel
                        </span>
                      </div>
                      <div className="text-xs text-emerald-700 font-semibold">
                        {adminStats.classesWithWaliCount}/
                        {adminStats.totalClasses} Ada Wali
                      </div>
                    </div>
                    <div className="mt-2.5 pt-2 border-t border-slate-100 text-[11px] text-slate-600 truncate">
                      {adminStats.classBreakdown
                        .slice(0, 2)
                        .map((c) => `${c.name}: ${c.waliName}`)
                        .join(' · ')}
                    </div>
                  </div>

                  {/* KPI 4: Jumlah Siswa Global & Per Kelas */}
                  <div
                    onClick={() => setAdminTab('siswa')}
                    className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 hover:border-amber-400 transition-colors cursor-pointer"
                  >
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                        Jumlah Siswa Global
                      </span>
                      <span className="w-8 h-8 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center">
                        <i className="fas fa-user-graduate"></i>
                      </span>
                    </div>
                    <div className="flex items-baseline justify-between gap-2">
                      <div className="text-2xl font-extrabold text-slate-900 font-mono">
                        {adminStats.totalStudentsGlobal}{' '}
                        <span className="text-xs font-sans font-semibold text-slate-500">
                          Siswa
                        </span>
                      </div>
                      <div className="text-xs text-amber-700 font-semibold font-mono">
                        X:{adminStats.studentsByGrade['10']} · XI:
                        {adminStats.studentsByGrade['11']} · XII:
                        {adminStats.studentsByGrade['12']}
                      </div>
                    </div>
                    <div className="mt-2.5 pt-2 border-t border-slate-100 flex flex-wrap items-center gap-2 text-[11px] text-slate-600">
                      {adminStats.classBreakdown.slice(0, 3).map((c) => (
                        <span key={c.id} className="font-mono">
                          {c.name}: <b>{c.studentCount}</b>
                        </span>
                      ))}
                    </div>
                  </div>
                </div>

                {/* Admin Navigation Tabs */}
                <div className="flex flex-wrap gap-1.5 border-b border-slate-200 pb-2">
                  <button
                    type="button"
                    onClick={() => setAdminTab('statistik')}
                    className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                      adminTab === 'statistik'
                        ? 'bg-blue-600 text-white'
                        : 'text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    <i className="fas fa-chart-bar mr-2"></i> Statistik &
                    Visualisasi Data
                  </button>
                  <button
                    type="button"
                    onClick={() => setAdminTab('nilai')}
                    className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                      adminTab === 'nilai'
                        ? 'bg-blue-600 text-white'
                        : 'text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    <i className="fas fa-table mr-2"></i> Input Nilai Massal
                    (Semua Mapel)
                  </button>
                  <button
                    type="button"
                    onClick={() => setAdminTab('sekolah')}
                    className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                      adminTab === 'sekolah'
                        ? 'bg-blue-600 text-white'
                        : 'text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    <i className="fas fa-school mr-2"></i> Data Sekolah & Logo
                  </button>
                  <button
                    type="button"
                    onClick={() => setAdminTab('kelas')}
                    className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                      adminTab === 'kelas'
                        ? 'bg-blue-600 text-white'
                        : 'text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    <i className="fas fa-door-open mr-2"></i> Data Kelas & Wali (
                    {classes.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setAdminTab('siswa')}
                    className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                      adminTab === 'siswa'
                        ? 'bg-blue-600 text-white'
                        : 'text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    <i className="fas fa-user-graduate mr-2"></i> Data Siswa (
                    {students.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setAdminTab('guru')}
                    className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                      adminTab === 'guru'
                        ? 'bg-blue-600 text-white'
                        : 'text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    <i className="fas fa-chalkboard-teacher mr-2"></i> Data Guru
                    ({adminStats.totalGuruCount})
                  </button>
                  <button
                    type="button"
                    onClick={() => setAdminTab('users')}
                    className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                      adminTab === 'users'
                        ? 'bg-blue-600 text-white'
                        : 'text-slate-600 hover:bg-slate-200/70'
                    }`}
                  >
                    <i className="fas fa-users-cog mr-2"></i> Pengguna & Role (
                    {users.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setAdminTab('database')}
                    className={`px-4 py-2 rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap ${
                      adminTab === 'database'
                        ? 'bg-emerald-600 text-white'
                        : 'text-emerald-700 hover:bg-emerald-50'
                    }`}
                  >
                    <i className="fas fa-database mr-2"></i> Database
                    Spreadsheet
                  </button>
                </div>

                {/* TAB: STATISTIK & VISUALISASI DATA */}
                {adminTab === 'statistik' && (
                  <div className="space-y-6">
                    {/* SECTION 1: DAFTAR KELAS LENGKAP DENGAN WALI KELAS, JUMLAH SISWA PER KELAS & PROGRES NILAI */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                      <div className="flex flex-wrap items-center justify-between border-b border-slate-200 pb-3 mb-4 gap-3">
                        <div>
                          <h5 className="font-bold text-base text-slate-900 mb-0.5">
                            <i className="fas fa-school text-emerald-600 mr-2"></i>
                            Visualisasi Kelas, Wali Kelas, Jumlah Siswa Per Kelas
                            & Progres Input Nilai
                          </h5>
                          <p className="text-xs text-slate-500">
                            Menampilkan total{' '}
                            <b>{adminStats.totalClasses} Kelas</b> lengkap
                            dengan Wali Kelasnya, distribusi dari{' '}
                            <b>{adminStats.totalStudentsGlobal} Siswa Global</b>
                            , serta persentase nilai yang sudah masuk di tiap
                            kelas.
                          </p>
                        </div>
                        <div className="flex flex-wrap items-center gap-2 text-xs">
                          <span className="px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 font-semibold">
                            Total Siswa Global:{' '}
                            <b className="font-mono text-slate-900">
                              {adminStats.totalStudentsGlobal} Siswa
                            </b>
                          </span>
                          <span className="px-3 py-1.5 rounded-lg bg-emerald-50 text-emerald-800 font-semibold">
                            Total Kelas:{' '}
                            <b className="font-mono">
                              {adminStats.totalClasses} Rombel
                            </b>
                          </span>
                        </div>
                      </div>

                      {/* Visual Cards Grid per Class */}
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-5">
                        {adminStats.classBreakdown.map((cls) => (
                          <div
                            key={cls.id}
                            className="p-4 rounded-xl border border-slate-200 bg-slate-50/60 flex flex-col justify-between space-y-3"
                          >
                            <div>
                              <div className="flex items-center justify-between gap-2 mb-1">
                                <span className="text-base font-extrabold text-slate-900">
                                  {cls.name}
                                </span>
                                <span className="text-xs font-bold text-blue-700 bg-blue-50 border border-blue-200 px-2 py-0.5 rounded">
                                  Tingkat {cls.grade}
                                </span>
                              </div>
                              <div className="text-xs text-slate-500 mb-3">
                                {cls.major}
                              </div>

                              {/* Wali Kelas Box (Only Name if no NIP/NUPTK, no dots!) */}
                              <div className="p-2.5 rounded-lg bg-white border border-slate-200/80 mb-3">
                                <div className="text-[11px] font-semibold text-slate-400 uppercase">
                                  Wali Kelas
                                </div>
                                <div className="font-bold text-xs text-slate-900 mt-0.5">
                                  {cls.waliName}
                                </div>
                                {cls.waliIdFormatted && (
                                  <div className="text-[11px] font-mono text-slate-500 mt-0.5">
                                    {cls.waliIdFormatted}
                                  </div>
                                )}
                              </div>

                              {/* Jumlah Siswa Per Kelas Visual Bar */}
                              <div className="mb-3">
                                <div className="flex items-center justify-between text-xs mb-1">
                                  <span className="font-semibold text-slate-600">
                                    Jumlah Siswa Kelas Ini
                                  </span>
                                  <span className="font-bold font-mono text-slate-900">
                                    {cls.studentCount} Siswa (
                                    {cls.studentSharePercent}%)
                                  </span>
                                </div>
                                <div className="w-full h-2 bg-slate-200 rounded-full overflow-hidden">
                                  <div
                                    className="h-full bg-amber-500 rounded-full"
                                    style={{
                                      width: `${Math.max(
                                        cls.studentCount > 0 ? 8 : 0,
                                        cls.studentSharePercent
                                      )}%`,
                                    }}
                                  ></div>
                                </div>
                              </div>

                              {/* Progres Nilai Masuk Kelas Ini */}
                              <div>
                                <div className="flex items-center justify-between text-xs mb-1">
                                  <span className="font-semibold text-slate-600">
                                    Nilai Sudah Diinput
                                  </span>
                                  <span className="font-bold font-mono text-slate-900">
                                    {cls.filledGrades}/{cls.expectedGrades} (
                                    {cls.progressPercent}%)
                                  </span>
                                </div>
                                <div className="w-full h-2.5 bg-slate-200 rounded-full overflow-hidden">
                                  <div
                                    className={`h-full rounded-full ${
                                      cls.progressPercent === 100
                                        ? 'bg-emerald-500'
                                        : cls.progressPercent >= 50
                                        ? 'bg-blue-600'
                                        : cls.progressPercent > 0
                                        ? 'bg-amber-500'
                                        : 'bg-slate-300'
                                    }`}
                                    style={{
                                      width: `${cls.progressPercent}%`,
                                    }}
                                  ></div>
                                </div>
                              </div>
                            </div>

                            <div className="pt-2 border-t border-slate-200/80 flex items-center justify-between text-[11px] text-slate-500">
                              <span>
                                Siswa Lengkap:{' '}
                                <b className="text-slate-800">
                                  {cls.completeStudentsInCls}/{cls.studentCount}
                                </b>
                              </span>
                              <span>
                                Rata-rata Nilai:{' '}
                                <b className="font-mono text-blue-700">
                                  {cls.avgScore}
                                </b>
                              </span>
                            </div>
                          </div>
                        ))}
                      </div>

                      {/* Tabel Rekapitulasi Kelas, Wali Kelas, Jumlah Siswa & Nilai */}
                      <div className="overflow-x-auto border border-slate-200 rounded-xl">
                        <table className="w-full text-sm text-left border-collapse">
                          <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                            <tr>
                              <th className="py-2.5 px-4">Rombel</th>
                              <th className="py-2.5 px-4">Wali Kelas</th>
                              <th className="py-2.5 px-4 text-center">
                                Jumlah Siswa
                              </th>
                              <th className="py-2.5 px-4 text-center">
                                Jumlah Mapel
                              </th>
                              <th className="py-2.5 px-4">
                                Progres Input Nilai
                              </th>
                              <th className="py-2.5 px-4 text-center">
                                Rata-Rata Kelas
                              </th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200">
                            {adminStats.classBreakdown.map((cls) => (
                              <tr key={cls.id} className="hover:bg-slate-50">
                                <td className="py-2.5 px-4">
                                  <div className="font-bold text-slate-900">
                                    {cls.name}
                                  </div>
                                  <div className="text-xs text-slate-500">
                                    Kelas {cls.grade} · {cls.major}
                                  </div>
                                </td>
                                <td className="py-2.5 px-4">
                                  <div className="font-bold text-slate-900">
                                    {cls.waliName}
                                  </div>
                                  {cls.waliIdFormatted && (
                                    <div className="text-xs font-mono text-slate-500">
                                      {cls.waliIdFormatted}
                                    </div>
                                  )}
                                </td>
                                <td className="py-2.5 px-4 text-center font-mono font-bold text-slate-900">
                                  {cls.studentCount} Siswa
                                </td>
                                <td className="py-2.5 px-4 text-center font-mono text-slate-700">
                                  {cls.subjectsCount} Mapel
                                </td>
                                <td className="py-2.5 px-4 min-w-[200px]">
                                  <div className="flex items-center justify-between text-xs mb-1">
                                    <span className="font-mono font-semibold text-slate-700">
                                      {cls.filledGrades} / {cls.expectedGrades}{' '}
                                      nilai
                                    </span>
                                    <span
                                      className={`font-mono font-bold ${
                                        cls.progressPercent === 100
                                          ? 'text-emerald-600'
                                          : cls.progressPercent > 0
                                          ? 'text-blue-600'
                                          : 'text-slate-400'
                                      }`}
                                    >
                                      {cls.progressPercent}%
                                    </span>
                                  </div>
                                  <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                                    <div
                                      className={`h-full rounded-full ${
                                        cls.progressPercent === 100
                                          ? 'bg-emerald-500'
                                          : cls.progressPercent >= 50
                                          ? 'bg-blue-600'
                                          : 'bg-amber-500'
                                      }`}
                                      style={{
                                        width: `${cls.progressPercent}%`,
                                      }}
                                    ></div>
                                  </div>
                                </td>
                                <td className="py-2.5 px-4 text-center font-mono font-bold text-blue-700">
                                  {cls.avgScore}
                                </td>
                              </tr>
                            ))}
                            <tr className="bg-slate-50 font-bold text-slate-900 border-t-2 border-slate-200">
                              <td className="py-3 px-4">
                                TOTAL GLOBAL ({adminStats.totalClasses} KELAS)
                              </td>
                              <td className="py-3 px-4 text-xs text-slate-600">
                                {adminStats.classesWithWaliCount} Wali Kelas
                                Terdaftar
                              </td>
                              <td className="py-3 px-4 text-center font-mono text-base text-slate-900">
                                {adminStats.totalStudentsGlobal} Siswa
                              </td>
                              <td className="py-3 px-4 text-center font-mono text-xs text-slate-600">
                                Semua Tingkat
                              </td>
                              <td className="py-3 px-4">
                                <div className="flex items-center justify-between text-xs mb-1">
                                  <span className="font-mono">
                                    {adminStats.totalFilledGradesGlobal} /{' '}
                                    {adminStats.totalExpectedGradesGlobal} nilai
                                  </span>
                                  <span className="font-mono text-emerald-700">
                                    {adminStats.globalProgressPercent}%
                                  </span>
                                </div>
                                <div className="w-full h-2 bg-slate-200 rounded-full overflow-hidden">
                                  <div
                                    className="h-full bg-emerald-600 rounded-full"
                                    style={{
                                      width: `${adminStats.globalProgressPercent}%`,
                                    }}
                                  ></div>
                                </div>
                              </td>
                              <td className="py-3 px-4 text-center font-mono text-blue-700">
                                {adminStats.globalAvgScore}
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* SECTION 2: VISUALISASI STATISTIK NILAI YANG SUDAH DIINPUT PER MATA PELAJARAN */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                      <div className="flex flex-wrap items-center justify-between border-b border-slate-200 pb-3 mb-4 gap-3">
                        <div>
                          <h5 className="font-bold text-base text-slate-900 mb-0.5">
                            <i className="fas fa-tasks text-blue-600 mr-2"></i>
                            Statistik Visual Nilai yang Sudah Diinput per Mata
                            Pelajaran
                          </h5>
                          <p className="text-xs text-slate-500">
                            Pantau progres pemasukan nilai STS untuk setiap mata
                            pelajaran secara global maupun per kelas.
                          </p>
                        </div>
                        <div className="flex items-center gap-2">
                          <label className="text-xs font-bold text-slate-600">
                            Filter Rombel:
                          </label>
                          <select
                            value={adminStatFilterClass}
                            onChange={(e) =>
                              setAdminStatFilterClass(e.target.value)
                            }
                            className="px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-bold outline-none focus:border-blue-600"
                          >
                            <option value="">
                              Semua Kelas / Global (
                              {adminStats.totalStudentsGlobal} Siswa)
                            </option>
                            {classes.map((c) => (
                              <option key={c.id} value={c.name}>
                                {c.name} (Kelas {c.grade})
                              </option>
                            ))}
                          </select>
                        </div>
                      </div>

                      {/* Global Status Summary Strip */}
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-5">
                        <div className="p-3.5 rounded-xl bg-emerald-50/70 border border-emerald-200 flex items-center justify-between">
                          <div>
                            <div className="text-xs font-semibold text-emerald-800">
                              Siswa Nilai Lengkap (100%)
                            </div>
                            <div className="text-xs text-emerald-600 mt-0.5">
                              Seluruh mapel telah terisi
                            </div>
                          </div>
                          <div className="text-xl font-extrabold font-mono text-emerald-700">
                            {adminStats.studentsCompleteGlobal} Siswa
                          </div>
                        </div>
                        <div className="p-3.5 rounded-xl bg-amber-50/70 border border-amber-200 flex items-center justify-between">
                          <div>
                            <div className="text-xs font-semibold text-amber-800">
                              Siswa Nilai Sebagian
                            </div>
                            <div className="text-xs text-amber-600 mt-0.5">
                              Masih ada mapel belum diisi
                            </div>
                          </div>
                          <div className="text-xl font-extrabold font-mono text-amber-700">
                            {adminStats.studentsPartialGlobal} Siswa
                          </div>
                        </div>
                        <div className="p-3.5 rounded-xl bg-slate-100 border border-slate-200 flex items-center justify-between">
                          <div>
                            <div className="text-xs font-semibold text-slate-700">
                              Siswa Belum Ada Nilai (0%)
                            </div>
                            <div className="text-xs text-slate-500 mt-0.5">
                              Belum ada mapel yang diinput
                            </div>
                          </div>
                          <div className="text-xl font-extrabold font-mono text-slate-700">
                            {adminStats.studentsEmptyGlobal} Siswa
                          </div>
                        </div>
                      </div>

                      {/* Visual Bar Chart List of Subjects */}
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3.5">
                        {adminStats.subjectStatsList.map((sub) => (
                          <div
                            key={sub.name}
                            className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/40 hover:bg-slate-50 transition-colors"
                          >
                            <div className="flex items-start justify-between gap-2 mb-1.5">
                              <div>
                                <span
                                  className={`inline-block text-[10px] font-bold uppercase px-1.5 py-0.5 rounded mr-1.5 ${
                                    sub.category === 'Umum'
                                      ? 'bg-blue-100 text-blue-800'
                                      : 'bg-purple-100 text-purple-800'
                                  }`}
                                >
                                  {sub.category}
                                </span>
                                <span className="font-bold text-xs text-slate-900">
                                  {sub.name}
                                </span>
                              </div>
                              <span
                                className={`text-xs font-mono font-bold shrink-0 ${
                                  sub.progressPercent === 100
                                    ? 'text-emerald-600'
                                    : sub.progressPercent > 0
                                    ? 'text-blue-600'
                                    : 'text-slate-400'
                                }`}
                              >
                                {sub.progressPercent}%
                              </span>
                            </div>

                            <div className="w-full h-2.5 bg-slate-200 rounded-full overflow-hidden mb-1.5">
                              <div
                                className={`h-full rounded-full transition-all duration-500 ${
                                  sub.progressPercent === 100
                                    ? 'bg-emerald-500'
                                    : sub.progressPercent >= 50
                                    ? 'bg-blue-600'
                                    : sub.progressPercent > 0
                                    ? 'bg-amber-500'
                                    : 'bg-slate-300'
                                }`}
                                style={{ width: `${sub.progressPercent}%` }}
                              ></div>
                            </div>

                            <div className="flex items-center justify-between text-[11px] text-slate-500">
                              <span>
                                Terisi:{' '}
                                <b className="font-mono text-slate-800">
                                  {sub.filledStudents}/{sub.targetStudents}
                                </b>{' '}
                                siswa
                                {sub.emptyStudents > 0 && (
                                  <span className="text-amber-600 ml-1">
                                    (Belum: {sub.emptyStudents})
                                  </span>
                                )}
                              </span>
                              <span>
                                Rata-rata:{' '}
                                <b className="font-mono text-slate-800">
                                  {sub.avgScore}
                                </b>
                              </span>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* SECTION 3: STATISTIK JUMLAH GURU & DAFTAR GURU LENGKAP */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                      <div className="flex flex-wrap items-center justify-between border-b border-slate-200 pb-3 mb-4 gap-3">
                        <div>
                          <h5 className="font-bold text-base text-slate-900 mb-0.5">
                            <i className="fas fa-chalkboard-teacher text-indigo-600 mr-2"></i>
                            Statistik Jumlah Guru (Wali Kelas & Guru Mapel)
                          </h5>
                          <p className="text-xs text-slate-500">
                            Guru yang belum memiliki NIP atau NUPTK ditampilkan{' '}
                            <b>hanya Nama Gurunya saja</b> tanpa titik-titik.
                          </p>
                        </div>
                        <div className="flex flex-wrap gap-2 text-xs">
                          <span className="px-3 py-1.5 rounded-lg bg-indigo-50 text-indigo-800 font-bold">
                            Total Guru: {adminStats.totalGuruCount}
                          </span>
                          <span className="px-3 py-1.5 rounded-lg bg-blue-50 text-blue-800 font-semibold">
                            Guru Mapel: {adminStats.guruMapelCount}
                          </span>
                          <span className="px-3 py-1.5 rounded-lg bg-emerald-50 text-emerald-800 font-semibold">
                            Wali Kelas: {adminStats.waliKelasUserCount}
                          </span>
                          {adminStats.dualRoleGuruCount > 0 && (
                            <span className="px-3 py-1.5 rounded-lg bg-amber-50 text-amber-800 font-semibold">
                              Rangkap Role: {adminStats.dualRoleGuruCount}
                            </span>
                          )}
                        </div>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                        {adminStats.allTeachersList.map((t, idx) => (
                          <div
                            key={idx}
                            className="p-3.5 rounded-xl border border-slate-200 bg-slate-50/50 flex flex-col justify-between"
                          >
                            <div>
                              <div className="font-bold text-sm text-slate-900">
                                {t.formattedName}
                              </div>
                              {t.idInfo.formatted && (
                                <div className="text-xs font-mono text-slate-600 mt-0.5">
                                  {t.idInfo.formatted}
                                </div>
                              )}
                            </div>
                            <div className="mt-2.5 pt-2 border-t border-slate-200/70 flex items-center justify-between text-[11px]">
                              <span className="font-semibold text-blue-700">
                                {t.roles
                                  .map((r) =>
                                    r === 'walikelas'
                                      ? `Wali Kelas (${t.waliKelasFor || '-'})`
                                      : r === 'gurumapel'
                                      ? 'Guru Mapel'
                                      : 'Admin'
                                  )
                                  .join(' · ')}
                              </span>
                              {t.idInfo.type !== 'NONE' ? (
                                <span className="font-mono text-[10px] px-1.5 py-0.5 rounded bg-slate-200 text-slate-700 font-bold">
                                  {t.idInfo.type}
                                </span>
                              ) : (
                                <span className="text-[10px] text-slate-400">
                                  Nama Saja
                                </span>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* SECTION 4: STATISTIK PERINGKAT 10 BESAR SISWA TERBAIK DI SETIAP KELAS */}
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                      <div className="flex flex-wrap items-center justify-between border-b border-slate-200 pb-4 mb-5 gap-4">
                        <div>
                          <div className="text-xs font-bold text-amber-600 uppercase tracking-wider mb-0.5">
                            Evaluasi Prestasi Akademik Siswa
                          </div>
                          <h5 className="font-bold text-base text-slate-900 mb-0.5">
                            <i className="fas fa-trophy text-amber-500 mr-2"></i>
                            Statistik Peringkat 10 Besar Siswa Terbaik di Setiap
                            Kelas
                          </h5>
                          <p className="text-xs text-slate-500">
                            Rekapitulasi 10 siswa dengan capaian rata-rata dan
                            total nilai tertinggi di masing-masing rombongan
                            belajar untuk evaluasi prestasi akademik.
                          </p>
                        </div>

                        <div className="flex flex-wrap items-center gap-2">
                          <label className="text-xs font-bold text-slate-600">
                            Pilih Kelas Evaluasi:
                          </label>
                          <select
                            value={adminTop10FilterClass}
                            onChange={(e) =>
                              setAdminTop10FilterClass(e.target.value)
                            }
                            className="px-3.5 py-2 border border-slate-300 bg-white rounded-lg text-xs font-bold text-slate-900 outline-none focus:border-blue-600"
                          >
                            <option value="">
                              Tampilkan Semua Kelas ({adminStats.totalClasses}{' '}
                              Rombel)
                            </option>
                            {classes.map((c) => (
                              <option key={c.id} value={c.name}>
                                Kelas {c.name} ({c.major})
                              </option>
                            ))}
                          </select>
                        </div>
                      </div>

                      {/* Quick Class Filter Buttons */}
                      <div className="flex flex-wrap gap-1.5 mb-5">
                        <button
                          type="button"
                          onClick={() => setAdminTop10FilterClass('')}
                          className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-colors cursor-pointer ${
                            adminTop10FilterClass === ''
                              ? 'bg-slate-900 text-white'
                              : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                          }`}
                        >
                          Semua Kelas ({adminStats.totalClasses})
                        </button>
                        {adminStats.classBreakdown.map((cls) => (
                          <button
                            key={cls.id}
                            type="button"
                            onClick={() => setAdminTop10FilterClass(cls.name)}
                            className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-colors cursor-pointer ${
                              adminTop10FilterClass === cls.name
                                ? 'bg-amber-500 text-slate-950 shadow-xs'
                                : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                            }`}
                          >
                            {cls.name}
                          </button>
                        ))}
                      </div>

                      <div className="space-y-6">
                        {adminStats.classBreakdown
                          .filter(
                            (cls) =>
                              !adminTop10FilterClass ||
                              cls.name.trim().toLowerCase() ===
                                adminTop10FilterClass.trim().toLowerCase()
                          )
                          .map((cls) => (
                            <div
                              key={cls.id}
                              className="border border-slate-200 rounded-xl overflow-hidden bg-white shadow-2xs"
                            >
                              {/* Class Banner Header */}
                              <div className="bg-slate-900 text-white px-5 py-3.5 flex flex-wrap items-center justify-between gap-3">
                                <div>
                                  <div className="flex items-center gap-2">
                                    <span className="px-2.5 py-0.5 rounded bg-amber-400 text-slate-950 font-extrabold text-xs">
                                      TOP 10 PRESTASI
                                    </span>
                                    <h6 className="font-bold text-base text-white mb-0">
                                      Rombel {cls.name} · Kelas {cls.grade} (
                                      {cls.major})
                                    </h6>
                                  </div>
                                  <div className="text-xs text-slate-300 mt-1">
                                    Wali Kelas:{' '}
                                    <b className="text-white">{cls.waliName}</b>
                                    {cls.waliIdFormatted
                                      ? ` (${cls.waliIdFormatted})`
                                      : ''}{' '}
                                    · Total Siswa:{' '}
                                    <b className="text-white">
                                      {cls.studentCount} Siswa
                                    </b>
                                  </div>
                                </div>
                                <div className="flex flex-wrap items-center gap-2 text-xs">
                                  <div className="bg-slate-800 border border-slate-700 px-3 py-1.5 rounded-lg">
                                    <span className="text-slate-400 mr-1.5">
                                      Rata-Rata Kelas:
                                    </span>
                                    <b className="font-mono text-amber-400 text-sm">
                                      {cls.avgScore}
                                    </b>
                                  </div>
                                  <div className="bg-slate-800 border border-slate-700 px-3 py-1.5 rounded-lg">
                                    <span className="text-slate-400 mr-1.5">
                                      Progres Nilai:
                                    </span>
                                    <b className="font-mono text-emerald-400 text-sm">
                                      {cls.progressPercent}%
                                    </b>
                                  </div>
                                </div>
                              </div>

                              {cls.top10Students.length === 0 ? (
                                <div className="p-8 text-center text-slate-400 text-xs">
                                  Belum ada data siswa di kelas {cls.name}.
                                </div>
                              ) : (
                                <div className="p-5 space-y-4">
                                  {/* Top 3 Podium Cards */}
                                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3.5">
                                    {cls.top10Students
                                      .slice(0, 3)
                                      .map((st, idx) => {
                                        const podiumStyles = [
                                          {
                                            border:
                                              'border-amber-300 bg-amber-50/60',
                                            badge:
                                              'bg-amber-400 text-slate-950',
                                            label: 'PERINGKAT 1 · TERBAIK I',
                                          },
                                          {
                                            border:
                                              'border-slate-300 bg-slate-50/80',
                                            badge: 'bg-slate-600 text-white',
                                            label: 'PERINGKAT 2 · TERBAIK II',
                                          },
                                          {
                                            border:
                                              'border-orange-300 bg-orange-50/50',
                                            badge: 'bg-orange-600 text-white',
                                            label: 'PERINGKAT 3 · TERBAIK III',
                                          },
                                        ][idx];
                                        return (
                                          <div
                                            key={st.id}
                                            className={`p-4 rounded-xl border ${podiumStyles.border} flex flex-col justify-between`}
                                          >
                                            <div>
                                              <div className="flex items-center justify-between gap-2 mb-2">
                                                <span
                                                  className={`px-2 py-0.5 rounded text-[10px] font-extrabold tracking-wide ${podiumStyles.badge}`}
                                                >
                                                  <i className="fas fa-medal mr-1"></i>
                                                  {podiumStyles.label}
                                                </span>
                                                <span
                                                  className={`text-[11px] font-bold px-2 py-0.5 rounded ${st.predicateColor}`}
                                                >
                                                  {st.predicateLabel}
                                                </span>
                                              </div>
                                              <div className="font-bold text-sm text-slate-900">
                                                {st.name}
                                              </div>
                                              <div className="text-[11px] font-mono text-slate-500 mt-0.5">
                                                NISN: {st.nisn} / NIS: {st.nis}
                                              </div>
                                            </div>

                                            <div className="mt-3 pt-2.5 border-t border-slate-200/80 flex items-center justify-between">
                                              <div>
                                                <div className="text-[10px] text-slate-500 uppercase font-semibold">
                                                  Total Nilai
                                                </div>
                                                <div className="font-mono font-bold text-sm text-slate-900">
                                                  {st.totalScore}{' '}
                                                  <span className="text-[11px] font-normal text-slate-500">
                                                    ({st.filledSubjectsCount}/
                                                    {st.subjectsCount} mapel)
                                                  </span>
                                                </div>
                                              </div>
                                              <div className="text-right">
                                                <div className="text-[10px] text-slate-500 uppercase font-semibold">
                                                  Rata-Rata
                                                </div>
                                                <div className="font-mono font-extrabold text-lg text-blue-700">
                                                  {st.avgScore}
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        );
                                      })}
                                  </div>

                                  {/* Full Top 10 Table */}
                                  <div className="overflow-x-auto border border-slate-200 rounded-xl">
                                    <table className="w-full text-sm text-left border-collapse">
                                      <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                                        <tr>
                                          <th className="py-2.5 px-3 text-center w-24">
                                            Peringkat
                                          </th>
                                          <th className="py-2.5 px-3">
                                            Nama Siswa & NISN/NIS
                                          </th>
                                          <th className="py-2.5 px-3 text-center">
                                            Mapel Terisi
                                          </th>
                                          <th className="py-2.5 px-3 text-center">
                                            Jumlah Nilai
                                          </th>
                                          <th className="py-2.5 px-3 min-w-[180px]">
                                            Rata-Rata & Visual Capaian
                                          </th>
                                          <th className="py-2.5 px-3">
                                            Mapel Tertinggi
                                          </th>
                                          <th className="py-2.5 px-3 text-center">
                                            Predikat
                                          </th>
                                          <th className="py-2.5 px-3 text-center">
                                            Raport
                                          </th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-200">
                                        {cls.top10Students.map((st) => (
                                          <tr
                                            key={st.id}
                                            className={
                                              st.rank <= 3
                                                ? 'bg-amber-50/25 hover:bg-amber-50/50'
                                                : 'hover:bg-slate-50'
                                            }
                                          >
                                            <td className="py-2.5 px-3 text-center">
                                              {st.rank === 1 ? (
                                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-amber-400 text-slate-950 font-mono font-extrabold text-xs">
                                                  <i className="fas fa-crown text-[10px]"></i>{' '}
                                                  #1
                                                </span>
                                              ) : st.rank === 2 ? (
                                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-slate-600 text-white font-mono font-extrabold text-xs">
                                                  <i className="fas fa-medal text-[10px]"></i>{' '}
                                                  #2
                                                </span>
                                              ) : st.rank === 3 ? (
                                                <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-orange-600 text-white font-mono font-extrabold text-xs">
                                                  <i className="fas fa-medal text-[10px]"></i>{' '}
                                                  #3
                                                </span>
                                              ) : (
                                                <span className="inline-block px-2.5 py-0.5 rounded bg-slate-100 text-slate-700 font-mono font-bold text-xs">
                                                  #{st.rank}
                                                </span>
                                              )}
                                            </td>
                                            <td className="py-2.5 px-3">
                                              <div className="font-bold text-slate-900">
                                                {st.name}
                                              </div>
                                              <div className="text-xs font-mono text-slate-500">
                                                {st.nisn} / {st.nis}
                                              </div>
                                            </td>
                                            <td className="py-2.5 px-3 text-center font-mono text-xs font-semibold text-slate-700">
                                              {st.filledSubjectsCount} /{' '}
                                              {st.subjectsCount} Mapel
                                            </td>
                                            <td className="py-2.5 px-3 text-center font-mono font-bold text-slate-900">
                                              {st.totalScore}
                                            </td>
                                            <td className="py-2.5 px-3">
                                              <div className="flex items-center justify-between text-xs mb-1">
                                                <span className="font-mono font-extrabold text-blue-700">
                                                  {st.avgScore}
                                                </span>
                                                <span className="text-[10px] text-slate-400 font-mono">
                                                  / 100
                                                </span>
                                              </div>
                                              <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                                                <div
                                                  className={`h-full rounded-full ${
                                                    st.avgScore >= 85
                                                      ? 'bg-emerald-500'
                                                      : st.avgScore >= 75
                                                      ? 'bg-blue-600'
                                                      : st.avgScore > 0
                                                      ? 'bg-amber-500'
                                                      : 'bg-slate-300'
                                                  }`}
                                                  style={{
                                                    width: `${Math.min(
                                                      100,
                                                      Math.max(0, st.avgScore)
                                                    )}%`,
                                                  }}
                                                ></div>
                                              </div>
                                            </td>
                                            <td className="py-2.5 px-3 text-xs">
                                              {st.bestSubjectName !== '-' ? (
                                                <div>
                                                  <span className="font-semibold text-slate-800">
                                                    {st.bestSubjectName}
                                                  </span>{' '}
                                                  <span className="font-mono font-bold text-emerald-700">
                                                    ({st.bestSubjectScore})
                                                  </span>
                                                </div>
                                              ) : (
                                                <span className="text-slate-400">
                                                  -
                                                </span>
                                              )}
                                            </td>
                                            <td className="py-2.5 px-3 text-center">
                                              <span
                                                className={`inline-block text-[11px] font-bold px-2 py-0.5 rounded whitespace-nowrap ${st.predicateColor}`}
                                              >
                                                {st.predicateLabel}
                                              </span>
                                            </td>
                                            <td className="py-2.5 px-3 text-center">
                                              <button
                                                type="button"
                                                onClick={() =>
                                                  openSingleStudentPreview(st)
                                                }
                                                className="px-2.5 py-1 border border-blue-600 text-blue-600 hover:bg-blue-600 hover:text-white rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap"
                                              >
                                                <i className="fas fa-file-pdf mr-1"></i>
                                                Raport
                                              </button>
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </div>
                              )}
                            </div>
                          ))}
                      </div>
                    </div>
                  </div>
                )}

                {/* TAB: INPUT NILAI MASSAL SEMUA MAPEL PER KELAS (KHUSUS ADMIN) */}
                {adminTab === 'nilai' && (
                  <div className="space-y-5">
                    <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                      <div className="flex flex-wrap items-center justify-between border-b border-slate-200 pb-4 mb-4 gap-4">
                        <div>
                          <div className="text-xs font-bold text-blue-600 uppercase tracking-wider mb-0.5">
                            Fitur Khusus Administrator
                          </div>
                          <h5 className="font-bold text-lg text-slate-900 mb-0.5">
                            <i className="fas fa-table text-blue-600 mr-2"></i>
                            Input & Impor Nilai Massal Semua Mata Pelajaran Per
                            Kelas
                          </h5>
                          <p className="text-xs text-slate-500">
                            Isi atau impor nilai seluruh mata pelajaran (Umum &
                            Kejuruan) sekaligus untuk satu kelas penuh dalam
                            satu tabel leger.
                          </p>
                        </div>

                        <div className="flex flex-wrap items-center gap-2">
                          <button
                            type="button"
                            onClick={() =>
                              downloadGradesTemplate(
                                adminGradeSelectedClass,
                                adminGradeClassStudents,
                                adminGradeClassSubjects[0]?.name || '',
                                adminGradeClassSubjects.map((s) => s.name),
                                studentGrades,
                                'all_subjects'
                              )
                            }
                            className="px-3.5 py-2 border border-emerald-600 text-emerald-700 hover:bg-emerald-50 text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                          >
                            <i className="fas fa-file-excel mr-1.5"></i> Unduh
                            Template Semua Mapel ({adminGradeSelectedClass})
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setBulkModalType('nilai_admin_semua_mapel');
                              setBulkPasteText('');
                            }}
                            className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                          >
                            <i className="fas fa-file-import mr-1.5"></i> Impor
                            Massal Semua Mapel
                          </button>
                        </div>
                      </div>

                      {/* Class Selector & Quick Action Bar */}
                      <div className="flex flex-wrap items-end justify-between gap-4 bg-slate-50 p-4 rounded-xl border border-slate-200 mb-5">
                        <div className="flex flex-wrap items-center gap-4">
                          <div>
                            <label className="block text-xs font-bold text-slate-600 mb-1">
                              Pilih Kelas / Rombel
                            </label>
                            <select
                              value={adminGradeSelectedClass}
                              onChange={(e) =>
                                setAdminGradeSelectedClass(e.target.value)
                              }
                              className="px-3.5 py-2 border border-slate-300 bg-white rounded-lg text-sm font-bold text-slate-900 outline-none focus:border-blue-600 min-w-[200px]"
                            >
                              {classes.map((c) => (
                                <option key={c.id} value={c.name}>
                                  {c.name} (Tingkat {c.grade})
                                </option>
                              ))}
                            </select>
                          </div>
                          <div className="text-xs text-slate-600 space-y-0.5 pt-2">
                            <div>
                              Jumlah Siswa:{' '}
                              <b className="font-mono text-slate-900">
                                {adminGradeClassStudents.length} Siswa
                              </b>{' '}
                              · Jumlah Mapel:{' '}
                              <b className="font-mono text-slate-900">
                                {adminGradeClassSubjects.length} Mata Pelajaran
                              </b>
                            </div>
                            <div>
                              Wali Kelas:{' '}
                              <b className="text-slate-900">
                                {formatTeacherNameWithTitle(
                                  getClassInfo(adminGradeSelectedClass)
                                    .waliKelasName
                                )}
                              </b>
                            </div>
                          </div>
                        </div>

                        <div className="flex flex-wrap items-center gap-2">
                          <button
                            type="button"
                            onClick={fillAdminAllSubjectsRandomGrades}
                            className="px-3.5 py-2 border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
                          >
                            <i className="fas fa-magic mr-1.5"></i> Isi Sample
                            Semua Mapel
                          </button>
                          <button
                            type="button"
                            onClick={handleSaveAdminAllSubjectsGrades}
                            className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer"
                          >
                            <i className="fas fa-save mr-1.5"></i> Simpan Nilai
                            Semua Mapel ({adminGradeSelectedClass})
                          </button>
                        </div>
                      </div>

                      {/* Matrix Leger Table for All Subjects in Selected Class */}
                      {adminGradeClassStudents.length === 0 ? (
                        <div className="text-center py-10 text-slate-400">
                          <i className="fas fa-user-slash fa-3x mb-3"></i>
                          <p className="text-sm">
                            Belum ada siswa di kelas {adminGradeSelectedClass}.
                          </p>
                        </div>
                      ) : (
                        <div className="overflow-x-auto border border-slate-200 rounded-xl">
                          <table className="w-full text-xs text-left border-collapse">
                            <thead className="bg-slate-100 text-slate-700 border-b border-slate-200">
                              <tr>
                                <th className="py-3 px-3 text-center w-10 sticky left-0 bg-slate-100 z-10 border-r border-slate-200">
                                  No
                                </th>
                                <th className="py-3 px-3 min-w-[200px] sticky left-10 bg-slate-100 z-10 border-r border-slate-200">
                                  Nama Siswa & NISN
                                </th>
                                {adminGradeClassSubjects.map((sub, sIdx) => (
                                  <th
                                    key={sub.name}
                                    className="py-2.5 px-2 text-center min-w-[120px] border-r border-slate-200"
                                    title={`[${sub.category}] ${sub.name}`}
                                  >
                                    <div
                                      className={`text-[10px] font-bold uppercase mb-0.5 ${
                                        sub.category === 'Umum'
                                          ? 'text-blue-700'
                                          : 'text-purple-700'
                                      }`}
                                    >
                                      {sub.category} #{sIdx + 1}
                                    </div>
                                    <div className="font-bold text-slate-900 line-clamp-2 leading-tight">
                                      {sub.name}
                                    </div>
                                  </th>
                                ))}
                                <th className="py-3 px-3 text-center w-20 bg-slate-50 font-bold">
                                  Total
                                </th>
                                <th className="py-3 px-3 text-center w-20 bg-slate-50 font-bold">
                                  Rata-Rata
                                </th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-200">
                              {adminGradeClassStudents.map((st, idx) => {
                                const stGrades = studentGrades[st.id] || {};
                                let total = 0;
                                let count = 0;
                                adminGradeClassSubjects.forEach((sub) => {
                                  const v = stGrades[sub.name];
                                  if (
                                    v !== undefined &&
                                    String(v) !== '' &&
                                    !Number.isNaN(Number(v))
                                  ) {
                                    total += Number(v);
                                    count++;
                                  }
                                });
                                const avg =
                                  count > 0
                                    ? Number((total / count).toFixed(1))
                                    : 0;

                                return (
                                  <tr
                                    key={st.id}
                                    className="hover:bg-blue-50/30"
                                  >
                                    <td className="py-2.5 px-3 text-center font-mono font-bold text-slate-500 sticky left-0 bg-white border-r border-slate-200">
                                      {idx + 1}
                                    </td>
                                    <td className="py-2.5 px-3 sticky left-10 bg-white border-r border-slate-200">
                                      <div className="font-bold text-slate-900">
                                        {st.name}
                                      </div>
                                      <div className="text-[11px] font-mono text-slate-500">
                                        {st.nisn}
                                      </div>
                                    </td>
                                    {adminGradeClassSubjects.map((sub) => {
                                      const val = stGrades[sub.name];
                                      const isFilled =
                                        val !== undefined && String(val) !== '';
                                      return (
                                        <td
                                          key={sub.name}
                                          className="p-1.5 border-r border-slate-100 text-center"
                                        >
                                          <input
                                            type="number"
                                            min={0}
                                            max={100}
                                            value={isFilled ? val : ''}
                                            onChange={(e) =>
                                              handleGradeChange(
                                                st.id,
                                                sub.name,
                                                e.target.value
                                              )
                                            }
                                            placeholder="-"
                                            className={`w-full px-2 py-1.5 border rounded-lg text-center font-mono font-bold text-xs outline-none focus:border-blue-600 ${
                                              isFilled
                                                ? 'border-slate-300 bg-white text-slate-900'
                                                : 'border-amber-200 bg-amber-50/40 text-slate-700'
                                            }`}
                                          />
                                        </td>
                                      );
                                    })}
                                    <td className="py-2.5 px-3 text-center font-mono font-bold text-slate-900 bg-slate-50/60">
                                      {total}
                                    </td>
                                    <td className="py-2.5 px-3 text-center font-mono font-bold text-blue-700 bg-slate-50/60">
                                      {avg}
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {/* TAB: SEKOLAH */}
                {adminTab === 'sekolah' && (
                  <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                    <h5 className="font-bold text-base text-slate-900 border-b border-slate-200 pb-3 mb-4">
                      <i className="fas fa-building text-blue-600 mr-2"></i>
                      Identitas Sekolah, Logo Aplikasi & Titimangsa
                    </h5>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-1">
                          Nama Sekolah
                        </label>
                        <input
                          type="text"
                          value={school.name}
                          onChange={(e) =>
                            setSchool({ ...school, name: e.target.value })
                          }
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-blue-600"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-1">
                          NPSN
                        </label>
                        <input
                          type="text"
                          value={school.npsn}
                          onChange={(e) =>
                            setSchool({ ...school, npsn: e.target.value })
                          }
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-blue-600 font-mono"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-1">
                          Kepala Sekolah (Gelar otomatis sesuai EYD, contoh:
                          S.T., M.Pd.)
                        </label>
                        <input
                          type="text"
                          value={school.principal}
                          onChange={(e) =>
                            setSchool({ ...school, principal: e.target.value })
                          }
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-blue-600"
                        />
                        <span className="text-[11px] text-slate-500 mt-0.5 block">
                          Tampilan Cetak:{' '}
                          <b>{formatTeacherNameWithTitle(school.principal)}</b>
                        </span>
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-1">
                          NIP / NUPTK Kepala Sekolah
                        </label>
                        <input
                          type="text"
                          value={school.principalNip}
                          onChange={(e) =>
                            setSchool({
                              ...school,
                              principalNip: e.target.value,
                            })
                          }
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-blue-600 font-mono"
                        />
                        <span className="text-[11px] text-slate-500 mt-0.5 block">
                          Terdeteksi Otomatis:{' '}
                          <b>
                            {formatNipOrNuptk(school.principalNip).formatted ||
                              'Tanpa NIP/NUPTK (Ditampilkan Hanya Nama)'}
                          </b>
                        </span>
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-1">
                          Tahun Pelajaran
                        </label>
                        <input
                          type="text"
                          value={school.tp}
                          onChange={(e) =>
                            setSchool({ ...school, tp: e.target.value })
                          }
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-blue-600"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-600 mb-1">
                          Tanggal & Tempat Cetak Raport
                        </label>
                        <input
                          type="text"
                          value={school.datePrint}
                          onChange={(e) =>
                            setSchool({ ...school, datePrint: e.target.value })
                          }
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-blue-600"
                        />
                      </div>

                      {/* URL LOGO APLIKASI / SEKOLAH */}
                      <div className="md:col-span-2 p-4 rounded-xl bg-slate-50 border border-slate-200">
                        <div className="flex flex-wrap items-center justify-between gap-4">
                          <div className="flex-grow">
                            <label className="block text-xs font-bold text-slate-700 mb-1">
                              URL Gambar Logo Aplikasi / Sekolah (Header &
                              Halaman Login)
                            </label>
                            <div className="flex flex-wrap gap-2">
                              <input
                                type="text"
                                value={school.logoUrl}
                                onChange={(e) =>
                                  setSchool({
                                    ...school,
                                    logoUrl: e.target.value,
                                  })
                                }
                                placeholder="https://ihuzda.site.je/urlgambar/uploads/img2.png"
                                className="flex-grow px-3 py-2 border border-slate-300 bg-white rounded-lg text-sm outline-none focus:border-blue-600 font-mono"
                              />
                              <button
                                type="button"
                                onClick={() =>
                                  setSchool({
                                    ...school,
                                    logoUrl:
                                      'https://ihuzda.site.je/urlgambar/uploads/img2.png',
                                  })
                                }
                                className="px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 text-xs font-bold rounded-lg cursor-pointer whitespace-nowrap"
                              >
                                Gunakan Logo Default
                              </button>
                            </div>
                            <span className="text-xs text-slate-500 mt-1 block">
                              Logo ini ditampilkan pada pojok kiri atas aplikasi
                              (E-Raport STS) dan halaman Login.
                            </span>
                          </div>
                          <div className="flex items-center gap-3 bg-slate-900 text-white px-4 py-2.5 rounded-xl shrink-0">
                            <div className="w-10 h-10 rounded-lg bg-white/10 border border-white/20 p-1 flex items-center justify-center overflow-hidden">
                              {!logoImgError && school.logoUrl ? (
                                <img
                                  src={school.logoUrl}
                                  alt="Preview Logo"
                                  className="w-full h-full object-contain"
                                  referrerPolicy="no-referrer"
                                  onError={() => setLogoImgError(true)}
                                />
                              ) : (
                                <i className="fas fa-graduation-cap text-blue-400"></i>
                              )}
                            </div>
                            <div>
                              <div className="text-[10px] text-slate-400 uppercase">
                                Preview Logo
                              </div>
                              <div className="font-bold text-xs">
                                E-Raport STS
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="md:col-span-2">
                        <label className="block text-xs font-bold text-slate-600 mb-1">
                          URL Image KOP Surat Raport
                        </label>
                        <input
                          type="text"
                          value={school.kopUrl}
                          onChange={(e) =>
                            setSchool({ ...school, kopUrl: e.target.value })
                          }
                          placeholder="https://..."
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm outline-none focus:border-blue-600 font-mono"
                        />
                        <span className="text-xs text-slate-400 mt-1 block">
                          Direct Link gambar KOP yang akan dipasang di bagian
                          paling atas raport.
                        </span>
                      </div>

                      {/* FITUR 1: UPLOAD SCAN TTD KEPALA SEKOLAH ATAU QR CODE KEPALA SEKOLAH */}
                      <div className="md:col-span-2 p-5 rounded-xl bg-blue-50/60 border border-blue-200 space-y-4">
                        <div className="flex flex-wrap items-center justify-between gap-2 border-b border-blue-200/80 pb-3">
                          <div>
                            <span className="inline-block px-2 py-0.5 rounded bg-blue-600 text-white text-[10px] font-bold uppercase tracking-wide mb-1">
                              Tanda Tangan Elektronik / Digital Kepala Sekolah
                            </span>
                            <h6 className="font-bold text-sm text-slate-900">
                              <i className="fas fa-signature text-blue-600 mr-1.5"></i>
                              Upload Scan TTD Kepala Sekolah atau Upload QR Code Kepala Sekolah
                            </h6>
                            <p className="text-xs text-slate-500 mt-0.5">
                              Pilih jenis tampilan tanda tangan Kepala Sekolah pada bagian bawah cetak Raport & PDF.
                            </p>
                          </div>
                          {school.principalSignatureUrl && (
                            <button
                              type="button"
                              onClick={() =>
                                setSchool({
                                  ...school,
                                  principalSignatureType: 'none',
                                  principalSignatureUrl: '',
                                })
                              }
                              className="px-3 py-1.5 bg-red-50 hover:bg-red-100 text-red-600 border border-red-200 rounded-lg text-xs font-bold cursor-pointer"
                            >
                              <i className="fas fa-trash-alt mr-1.5"></i>
                              Hapus Gambar TTD/QR KS
                            </button>
                          )}
                        </div>

                        <div className="grid grid-cols-1 lg:grid-cols-12 gap-5 items-start">
                          <div className="lg:col-span-7 space-y-3.5">
                            {/* Mode Selector */}
                            <div>
                              <label className="block text-xs font-bold text-slate-700 mb-1.5">
                                1. Pilih Jenis Tanda Tangan Kepala Sekolah di Raport:
                              </label>
                              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                                <button
                                  type="button"
                                  onClick={() =>
                                    setSchool({
                                      ...school,
                                      principalSignatureType: 'none',
                                    })
                                  }
                                  className={`px-3 py-2 rounded-lg text-xs font-bold border text-left transition-all cursor-pointer ${
                                    (!school.principalSignatureType ||
                                      school.principalSignatureType ===
                                        'none')
                                      ? 'bg-slate-800 text-white border-slate-800 shadow-sm'
                                      : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                                  }`}
                                >
                                  <i className="fas fa-pen-nib mr-1.5"></i>
                                  TTD Manual (Kosong)
                                </button>
                                <button
                                  type="button"
                                  onClick={() =>
                                    setSchool({
                                      ...school,
                                      principalSignatureType: 'signature',
                                    })
                                  }
                                  className={`px-3 py-2 rounded-lg text-xs font-bold border text-left transition-all cursor-pointer ${
                                    school.principalSignatureType ===
                                    'signature'
                                      ? 'bg-blue-600 text-white border-blue-600 shadow-sm'
                                      : 'bg-white text-slate-700 border-slate-300 hover:bg-blue-50'
                                  }`}
                                >
                                  <i className="fas fa-signature mr-1.5"></i>
                                  Scan TTD KS
                                </button>
                                <button
                                  type="button"
                                  onClick={() =>
                                    setSchool({
                                      ...school,
                                      principalSignatureType: 'qrcode',
                                    })
                                  }
                                  className={`px-3 py-2 rounded-lg text-xs font-bold border text-left transition-all cursor-pointer ${
                                    school.principalSignatureType === 'qrcode'
                                      ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                                      : 'bg-white text-slate-700 border-slate-300 hover:bg-emerald-50'
                                  }`}
                                >
                                  <i className="fas fa-qrcode mr-1.5"></i>
                                  QR Code KS
                                </button>
                              </div>
                            </div>

                            {/* Upload Buttons */}
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                              <label className="flex items-center justify-center gap-2 px-3.5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg cursor-pointer shadow-sm transition-colors">
                                <i className="fas fa-file-signature"></i>
                                <span>Upload File Scan TTD KS</span>
                                <input
                                  type="file"
                                  accept="image/*"
                                  onChange={(e) =>
                                    handleUploadPrincipalSigOrQr(
                                      e,
                                      'signature'
                                    )
                                  }
                                  className="hidden"
                                />
                              </label>
                              <label className="flex items-center justify-center gap-2 px-3.5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer shadow-sm transition-colors">
                                <i className="fas fa-qrcode"></i>
                                <span>Upload File QR Code KS</span>
                                <input
                                  type="file"
                                  accept="image/*"
                                  onChange={(e) =>
                                    handleUploadPrincipalSigOrQr(e, 'qrcode')
                                  }
                                  className="hidden"
                                />
                              </label>
                            </div>

                            {/* Optional URL Input & White Background Cleaner */}
                            <div>
                              <label className="block text-[11px] font-bold text-slate-600 mb-1">
                                Atau Tempelkan Link URL Gambar TTD / QR Code Kepala Sekolah:
                              </label>
                              <div className="flex flex-wrap sm:flex-nowrap gap-2">
                                <input
                                  type="text"
                                  value={school.principalSignatureUrl || ''}
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    setSchool({
                                      ...school,
                                      principalSignatureUrl: val,
                                      principalSignatureType:
                                        val &&
                                        (!school.principalSignatureType ||
                                          school.principalSignatureType ===
                                            'none')
                                          ? 'signature'
                                          : school.principalSignatureType,
                                    });
                                  }}
                                  placeholder="https://... atau upload file gambar di atas"
                                  className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs font-mono outline-none focus:border-blue-600"
                                />
                                {school.principalSignatureUrl &&
                                  school.principalSignatureType ===
                                    'signature' && (
                                    <button
                                      type="button"
                                      onClick={async () => {
                                        try {
                                          const cleaned =
                                            await removeWhiteBackgroundFromDataUrl(
                                              school.principalSignatureUrl!
                                            );
                                          setSchool({
                                            ...school,
                                            principalSignatureUrl: cleaned,
                                          });
                                          showToast(
                                            'Background putih pada Scan TTD Kepala Sekolah berhasil dibuat transparan!',
                                            'success'
                                          );
                                        } catch {
                                          showToast(
                                            'Gagal memproses background gambar eksternal.',
                                            'danger'
                                          );
                                        }
                                      }}
                                      className="px-3 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-lg text-xs font-bold cursor-pointer whitespace-nowrap"
                                      title="Hapus background kertas putih agar tanda tangan menyatu dengan raport"
                                    >
                                      <i className="fas fa-magic mr-1 text-blue-600"></i>
                                      Transparankan BG
                                    </button>
                                  )}
                              </div>
                            </div>
                          </div>

                          {/* Live Preview Box */}
                          <div className="lg:col-span-5 bg-white rounded-xl border border-slate-200 p-4 text-center shadow-xs">
                            <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-2 border-b border-slate-100 pb-1.5">
                              Preview Tampilan di Raport (Kepala Sekolah)
                            </div>
                            <div
                              style={{
                                fontFamily: "'Times New Roman', Times, serif",
                                fontSize: '10pt',
                                lineHeight: 1.25,
                                color: '#000',
                              }}
                            >
                              <p className="mb-0">Mengetahui,</p>
                              <p className="font-bold mb-0">
                                Kepala SMKN 1 Mesuji Timur
                              </p>
                              {school.principalSignatureType &&
                              school.principalSignatureType !== 'none' &&
                              school.principalSignatureUrl ? (
                                <div className="h-14 flex items-center justify-center my-1">
                                  <img
                                    src={school.principalSignatureUrl}
                                    alt="Preview TTD/QR KS"
                                    className={
                                      school.principalSignatureType === 'qrcode'
                                        ? 'w-12 h-12 object-contain mx-auto'
                                        : 'max-h-12 max-w-[170px] object-contain mx-auto'
                                    }
                                  />
                                </div>
                              ) : (
                                <div className="h-12 flex items-center justify-center text-[11px] font-sans text-slate-400 italic">
                                  (Ruang Tanda Tangan Basah / Manual)
                                </div>
                              )}
                              <p className="font-bold underline mb-0">
                                {formatTeacherNameWithTitle(school.principal)}
                              </p>
                              {formatNipOrNuptk(school.principalNip)
                                .formatted && (
                                <p
                                  className="mb-0"
                                  style={{ fontSize: '8.5pt' }}
                                >
                                  {
                                    formatNipOrNuptk(school.principalNip)
                                      .formatted
                                  }
                                </p>
                              )}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="mt-5 text-right">
                      <button
                        type="button"
                        onClick={handleSaveSchoolData}
                        className="px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer"
                      >
                        <i className="fas fa-save mr-1.5"></i> Simpan Data
                        Sekolah, Logo & TTD/QR KS
                      </button>
                    </div>
                  </div>
                )}

                {/* TAB: KELAS */}
                {adminTab === 'kelas' && (
                  <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                    <div className="flex flex-wrap justify-between items-center border-b border-slate-200 pb-3 mb-4 gap-3">
                      <div>
                        <h5 className="font-bold text-base text-slate-900 mb-0">
                          <i className="fas fa-list text-blue-600 mr-2"></i>
                          Daftar Rombongan Belajar, Wali Kelas & TTD/QR Wali Kelas
                        </h5>
                        <span className="text-xs text-slate-500">
                          Klik <b>Edit & Upload TTD/QR</b> pada kelas untuk
                          mengatur Nama, NIP/NUPTK, serta Scan TTD atau QR Code
                          Wali Kelas.
                        </span>
                      </div>
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={downloadClassTemplate}
                          className="px-3 py-2 border border-emerald-600 text-emerald-700 hover:bg-emerald-50 text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-file-excel mr-1.5"></i> Unduh
                          Template Kelas (.xlsx)
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setBulkModalType('kelas');
                            setBulkPasteText('');
                            setBulkReplaceExisting(false);
                          }}
                          className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-file-import mr-1.5"></i> Impor
                          Kelas Cepat
                        </button>
                        <button
                          type="button"
                          onClick={() => setShowAddClassModal(true)}
                          className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-plus mr-1.5"></i> Tambah Kelas
                        </button>
                      </div>
                    </div>

                    {classes.length === 0 ? (
                      <div className="text-center py-10 text-slate-400">
                        <i className="fas fa-folder-open fa-3x mb-3"></i>
                        <p className="text-sm">
                          Belum ada data kelas. Silakan tambah atau impor kelas
                          baru.
                        </p>
                      </div>
                    ) : (
                      <div className="overflow-x-auto">
                        <table className="w-full text-sm text-left border-collapse">
                          <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                            <tr>
                              <th className="py-3 px-4">Tingkat</th>
                              <th className="py-3 px-4">Nama Rombel</th>
                              <th className="py-3 px-4">
                                Konsentrasi Keahlian
                              </th>
                              <th className="py-3 px-4">
                                Wali Kelas (Format Raport)
                              </th>
                              <th className="py-3 px-4">
                                Identitas (Otomatis NIP / NUPTK)
                              </th>
                              <th className="py-3 px-4 text-center">
                                TTD / QR Wali Kelas
                              </th>
                              <th className="py-3 px-4 text-right">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200">
                            {classes.map((c) => {
                              const resolved = getClassInfo(c.name);
                              const idInfo = formatNipOrNuptk(
                                resolved.waliKelasNip
                              );
                              return (
                                <tr key={c.id} className="hover:bg-slate-50">
                                  <td className="py-3 px-4 font-semibold text-blue-600">
                                    Kelas {c.grade}
                                  </td>
                                  <td className="py-3 px-4 font-bold text-slate-900">
                                    {c.name}
                                  </td>
                                  <td className="py-3 px-4 text-slate-600">
                                    {c.major}
                                  </td>
                                  <td className="py-3 px-4 font-semibold text-slate-800">
                                    {formatTeacherNameWithTitle(
                                      resolved.waliKelasName
                                    )}
                                  </td>
                                  <td className="py-3 px-4 font-mono text-xs text-slate-700">
                                    {idInfo.formatted || (
                                      <span className="text-slate-400 font-sans">
                                        -
                                      </span>
                                    )}
                                  </td>
                                  <td className="py-3 px-4 text-center">
                                    {resolved.waliSignatureType &&
                                    resolved.waliSignatureType !== 'none' &&
                                    resolved.waliSignatureUrl ? (
                                      <div className="inline-flex items-center gap-2 px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-200">
                                        <img
                                          src={resolved.waliSignatureUrl}
                                          alt="TTD/QR Wali"
                                          className="h-6 w-auto max-w-[56px] object-contain bg-white rounded border border-slate-200 p-0.5"
                                        />
                                        <span className="text-[11px] font-bold text-emerald-700">
                                          {resolved.waliSignatureType ===
                                          'qrcode'
                                            ? 'QR Code'
                                            : 'Scan TTD'}
                                        </span>
                                      </div>
                                    ) : (
                                      <span className="text-xs text-slate-400">
                                        Manual / Belum Diupload
                                      </span>
                                    )}
                                  </td>
                                  <td className="py-3 px-4 text-right space-x-1.5 whitespace-nowrap">
                                    <button
                                      type="button"
                                      onClick={() =>
                                        setEditingClass({
                                          ...c,
                                          waliKelasName:
                                            resolved.waliKelasName || '',
                                          waliKelasNip:
                                            resolved.waliKelasNip || '',
                                          waliSignatureType:
                                            resolved.waliSignatureType ||
                                            'none',
                                          waliSignatureUrl:
                                            resolved.waliSignatureUrl || '',
                                        })
                                      }
                                      className="px-2.5 py-1 border border-blue-300 text-blue-600 hover:bg-blue-50 rounded text-xs font-semibold cursor-pointer"
                                      title="Edit Wali Kelas & Upload Scan TTD / QR Code"
                                    >
                                      <i className="fas fa-signature mr-1"></i>{' '}
                                      Edit & TTD/QR
                                    </button>
                                    <button
                                      type="button"
                                      onClick={() => handleRemoveClass(c)}
                                      className="px-2.5 py-1 border border-red-300 text-red-600 hover:bg-red-50 rounded text-xs cursor-pointer"
                                      title="Hapus Kelas"
                                    >
                                      <i className="fas fa-trash"></i>
                                    </button>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

                {/* TAB: SISWA */}
                {adminTab === 'siswa' && (
                  <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                    <div className="flex flex-wrap justify-between items-center border-b border-slate-200 pb-3 mb-4 gap-3">
                      <div className="flex items-center gap-3">
                        <h5 className="font-bold text-base text-slate-900 mb-0">
                          <i className="fas fa-users text-blue-600 mr-2"></i>
                          Data Peserta Didik
                        </h5>
                        <select
                          value={adminStudentFilterClass}
                          onChange={(e) =>
                            setAdminStudentFilterClass(e.target.value)
                          }
                          className="px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-semibold outline-none"
                        >
                          <option value="">Semua Rombel</option>
                          {classes.map((c) => (
                            <option key={c.id} value={c.name}>
                              {c.name}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={downloadStudentTemplate}
                          className="px-3 py-2 border border-emerald-600 text-emerald-700 hover:bg-emerald-50 text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-file-excel mr-1.5"></i> Unduh
                          Template Siswa (.xlsx)
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setBulkModalType('siswa');
                            setBulkPasteText('');
                            setBulkReplaceExisting(false);
                          }}
                          className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-file-import mr-1.5"></i> Impor
                          Data Siswa Cepat
                        </button>
                        <button
                          type="button"
                          onClick={() => setShowAddStudentModal(true)}
                          className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-user-plus mr-1.5"></i> Tambah
                          Siswa
                        </button>
                      </div>
                    </div>

                    {filteredAdminStudents.length === 0 ? (
                      <div className="text-center py-10 text-slate-400">
                        <i className="fas fa-user-slash fa-3x mb-3"></i>
                        <p className="text-sm">Tidak ditemukan data siswa.</p>
                      </div>
                    ) : (
                      <div className="overflow-x-auto">
                        <table className="w-full text-sm text-left border-collapse">
                          <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                            <tr>
                              <th className="py-3 px-4">No</th>
                              <th className="py-3 px-4">Nama Lengkap</th>
                              <th className="py-3 px-4">NISN / NIS</th>
                              <th className="py-3 px-4">Kelas</th>
                              <th className="py-3 px-4">
                                Konsentrasi Keahlian
                              </th>
                              <th className="py-3 px-4 text-right">Opsi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-200">
                            {filteredAdminStudents.map((s, index) => (
                              <tr key={s.id} className="hover:bg-slate-50">
                                <td className="py-3 px-4 text-slate-500 font-mono">
                                  {index + 1}
                                </td>
                                <td className="py-3 px-4 font-bold text-slate-900">
                                  {s.name}
                                </td>
                                <td className="py-3 px-4 font-mono text-xs text-slate-600">
                                  {s.nisn} / {s.nis}
                                </td>
                                <td className="py-3 px-4 font-semibold text-slate-700">
                                  {s.className}
                                </td>
                                <td className="py-3 px-4 text-xs text-slate-500">
                                  {s.major}
                                </td>
                                <td className="py-3 px-4 text-right">
                                  <button
                                    type="button"
                                    onClick={() => handleDeleteStudent(s)}
                                    className="px-2.5 py-1 border border-red-300 text-red-600 hover:bg-red-50 rounded text-xs cursor-pointer"
                                  >
                                    <i className="fas fa-trash"></i>
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

                {/* TAB: GURU */}
                {adminTab === 'guru' && (
                  <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                    <div className="flex flex-wrap justify-between items-center border-b border-slate-200 pb-3 mb-4 gap-3">
                      <div>
                        <h5 className="font-bold text-base text-slate-900 mb-0">
                          <i className="fas fa-chalkboard-teacher text-blue-600 mr-2"></i>
                          Data Guru (Wali Kelas & Guru Mapel)
                        </h5>
                        <p className="text-slate-500 text-xs mt-0.5">
                          Gelar akademik otomatis diformat sesuai EYD (misal:{' '}
                          <b>S.Pd.</b>) dan nomor identitas otomatis dibedakan
                          antara <b>NIP.</b> (18 digit) dan <b>NUPTK.</b> (16
                          digit).
                        </p>
                      </div>
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            setBulkModalType('walikelas_user');
                            setBulkPasteText('');
                          }}
                          className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-file-import mr-1.5"></i> Impor
                          Template Wali Kelas
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setBulkModalType('gurumapel_user');
                            setBulkPasteText('');
                          }}
                          className="px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-file-import mr-1.5"></i> Impor
                          Template Guru Mapel
                        </button>
                      </div>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full text-sm text-left border-collapse">
                        <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                          <tr>
                            <th className="py-3 px-4">
                              Nama Guru (Format EYD)
                            </th>
                            <th className="py-3 px-4">
                              Identitas (Otomatis NIP / NUPTK)
                            </th>
                            <th className="py-3 px-4">
                              Role / Hak Akses Terdaftar
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {users
                            .filter(
                              (x) =>
                                x.roles.includes('gurumapel') ||
                                x.roles.includes('walikelas')
                            )
                            .map((u) => {
                              const idFormatted = formatNipOrNuptk(u.nip);
                              return (
                                <tr key={u.id} className="hover:bg-slate-50">
                                  <td className="py-3 px-4 font-bold text-slate-900">
                                    {formatTeacherNameWithTitle(u.name)}
                                  </td>
                                  <td className="py-3 px-4 font-mono text-xs text-slate-700">
                                    {idFormatted.formatted || (
                                      <span className="text-slate-400 font-sans">
                                        -
                                      </span>
                                    )}
                                  </td>
                                  <td className="py-3 px-4 text-xs font-semibold text-slate-700">
                                    {u.roles
                                      .map((r) =>
                                        r === 'walikelas'
                                          ? `Wali Kelas (${
                                              u.waliKelasFor || '-'
                                            })`
                                          : r === 'admin'
                                          ? 'Admin'
                                          : 'Guru Mapel'
                                      )
                                      .join(' · ')}
                                  </td>
                                </tr>
                              );
                            })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* TAB: USERS & ROLES MANAGEMENT */}
                {adminTab === 'users' && (
                  <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                    <div className="flex flex-wrap justify-between items-center border-b border-slate-200 pb-3 mb-4 gap-3">
                      <div>
                        <h5 className="font-bold text-base text-slate-900 mb-0">
                          <i className="fas fa-user-shield text-blue-600 mr-2"></i>
                          Manajemen Pengguna & Hak Akses Role
                        </h5>
                        <span className="text-xs text-slate-500">
                          Impor cepat data akun Wali Kelas maupun Guru Mapel
                          menggunakan template Excel atau Copy-Paste.
                        </span>
                      </div>
                      <div className="flex flex-wrap gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            setBulkModalType('walikelas_user');
                            setBulkPasteText('');
                          }}
                          className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-file-import mr-1.5"></i> Impor
                          User Wali Kelas
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setBulkModalType('gurumapel_user');
                            setBulkPasteText('');
                          }}
                          className="px-3 py-2 bg-sky-600 hover:bg-sky-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-file-import mr-1.5"></i> Impor
                          User Guru Mapel
                        </button>
                        <button
                          type="button"
                          onClick={() => setShowAddUserModal(true)}
                          className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                        >
                          <i className="fas fa-user-plus mr-1.5"></i> Tambah
                          Pengguna
                        </button>
                      </div>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full text-sm text-left border-collapse">
                        <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                          <tr>
                            <th className="py-3 px-4">
                              Nama Lengkap & NIP/NUPTK
                            </th>
                            <th className="py-3 px-4">Username</th>
                            <th className="py-3 px-4">Kata Sandi</th>
                            <th className="py-3 px-4">Role Akses</th>
                            <th className="py-3 px-4">Wali Kelas Untuk</th>
                            <th className="py-3 px-4 text-right">Aksi</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {users.map((u) => (
                            <tr key={u.id} className="hover:bg-slate-50">
                              <td className="py-3 px-4">
                                <div className="font-bold text-slate-900">
                                  {formatTeacherNameWithTitle(u.name)}
                                </div>
                                {formatNipOrNuptk(u.nip).formatted && (
                                  <div className="text-xs font-mono text-slate-500">
                                    {formatNipOrNuptk(u.nip).formatted}
                                  </div>
                                )}
                              </td>
                              <td className="py-3 px-4 font-mono text-xs font-semibold text-slate-800">
                                {u.username}
                              </td>
                              <td className="py-3 px-4 font-mono text-xs text-slate-700">
                                <div className="inline-flex items-center gap-1.5 bg-slate-100 px-2 py-1 rounded border border-slate-200">
                                  <span>
                                    {visibleUserPasswords[u.id]
                                      ? u.password
                                      : '••••••'}
                                  </span>
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setVisibleUserPasswords((prev) => ({
                                        ...prev,
                                        [u.id]: !prev[u.id],
                                      }))
                                    }
                                    title={
                                      visibleUserPasswords[u.id]
                                        ? 'Sembunyikan Password'
                                        : 'Lihat Password'
                                    }
                                    className="text-slate-400 hover:text-slate-700 cursor-pointer"
                                  >
                                    <i
                                      className={
                                        visibleUserPasswords[u.id]
                                          ? 'fas fa-eye-slash'
                                          : 'fas fa-eye'
                                      }
                                    ></i>
                                  </button>
                                </div>
                              </td>
                              <td className="py-3 px-4 text-xs font-semibold text-slate-700">
                                {u.roles
                                  .map((r) =>
                                    r === 'admin'
                                      ? 'Admin'
                                      : r === 'walikelas'
                                      ? 'Wali Kelas'
                                      : 'Guru Mapel'
                                  )
                                  .join(' · ')}
                              </td>
                              <td className="py-3 px-4">
                                {u.roles.includes('walikelas') ? (
                                  <span className="font-bold text-emerald-700 text-xs">
                                    {u.waliKelasFor || '-'}
                                  </span>
                                ) : (
                                  <span className="text-slate-400 text-xs">
                                    -
                                  </span>
                                )}
                              </td>
                              <td className="py-3 px-4 text-right">
                                <button
                                  type="button"
                                  disabled={u.username === 'admin'}
                                  onClick={() => handleDeleteUser(u)}
                                  className="px-2.5 py-1 border border-red-300 text-red-600 hover:bg-red-50 disabled:opacity-40 rounded text-xs cursor-pointer"
                                >
                                  <i className="fas fa-trash"></i>
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* TAB: GOOGLE SHEETS DATABASE STATUS & GID=0 INSPECTOR */}
                {adminTab === 'database' && (
                  <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6 space-y-5">
                    <div className="flex flex-wrap items-center justify-between border-b border-slate-200 pb-4 gap-4">
                      <div>
                        <h5 className="font-bold text-base text-slate-900 mb-1">
                          <i className="fas fa-database text-emerald-600 mr-2"></i>
                          Integrasi Database Google Spreadsheet (Khusus Admin)
                        </h5>
                        <p className="text-xs text-slate-500">
                          Spreadsheet Aktif:{' '}
                          <a
                            href={activeSheetUrl}
                            target="_blank"
                            rel="noreferrer"
                            className="font-mono text-blue-600 underline break-all"
                          >
                            {activeSheetUrl}
                          </a>
                        </p>
                      </div>
                      <div className="flex flex-wrap items-center gap-2">
                        {needsGoogleAuth ? (
                          <GoogleSignInButton
                            onClick={handleGoogleConnect}
                            disabled={isSyncingSheets}
                          />
                        ) : (
                          <button
                            type="button"
                            onClick={async () => {
                              await googleLogout();
                              setGoogleUser(null);
                              setGoogleToken(null);
                              setNeedsGoogleAuth(true);
                              showToast(
                                'Sesi Google OAuth telah diputuskan.',
                                'success'
                              );
                            }}
                            className="px-3 py-2 border border-slate-300 hover:bg-slate-100 text-slate-700 text-xs font-bold rounded-lg cursor-pointer"
                          >
                            <i className="fas fa-unlink mr-1.5"></i>
                            Putuskan OAuth
                          </button>
                        )}
                        <button
                          type="button"
                          onClick={() => loadFromGoogleSheets()}
                          disabled={isSyncingSheets}
                          className="px-3.5 py-2 bg-slate-800 hover:bg-slate-900 text-white text-xs font-bold rounded-lg cursor-pointer"
                        >
                          <i
                            className={`fas fa-sync-alt mr-1.5 ${
                              isSyncingSheets ? 'fa-spin' : ''
                            }`}
                          ></i>
                          Baca Data dari Spreadsheet
                        </button>
                        <button
                          type="button"
                          onClick={handleManualSyncToSheets}
                          disabled={isSyncingSheets}
                          className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer"
                        >
                          <i className="fas fa-cloud-upload-alt mr-1.5"></i>
                          Tulis / Inisialisasi ke Spreadsheet
                        </button>
                      </div>
                    </div>

                    {/* FITUR 2: UBAH LINK DATABASE SPREADSHEET */}
                    <div className="p-5 rounded-xl bg-emerald-50/70 border border-emerald-200 space-y-3">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div>
                          <span className="inline-block px-2 py-0.5 rounded bg-emerald-600 text-white text-[10px] font-bold uppercase tracking-wide mb-1">
                            Pengaturan Link Spreadsheet Target
                          </span>
                          <h6 className="font-bold text-sm text-slate-900">
                            <i className="fas fa-link text-emerald-600 mr-1.5"></i>
                            Ubah Link / URL Database Google Spreadsheet
                          </h6>
                        </div>
                        <div className="flex items-center gap-2">
                          <span className="text-[11px] font-mono bg-white px-2.5 py-1 rounded border border-emerald-200 text-slate-700">
                            ID Aktif: <b>{activeSheetId}</b>
                          </span>
                          {activeSheetId !== SPREADSHEET_ID && (
                            <button
                              type="button"
                              onClick={handleResetSpreadsheetLink}
                              disabled={isSyncingSheets}
                              className="px-3 py-1.5 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-lg text-xs font-bold cursor-pointer"
                            >
                              <i className="fas fa-undo mr-1"></i> Reset ke
                              Spreadsheet Default
                            </button>
                          )}
                        </div>
                      </div>

                      <p className="text-xs text-slate-600">
                        Tempelkan link URL Google Spreadsheet baru Anda (contoh:{' '}
                        <span className="font-mono text-[11px]">
                          https://docs.google.com/spreadsheets/d/.../edit
                        </span>
                        ) atau ID Spreadsheet untuk mengganti database yang
                        digunakan aplikasi.
                      </p>

                      <div className="flex flex-wrap sm:flex-nowrap gap-2">
                        <input
                          type="text"
                          value={spreadsheetLinkInput}
                          onChange={(e) =>
                            setSpreadsheetLinkInput(e.target.value)
                          }
                          placeholder="https://docs.google.com/spreadsheets/d/13V9knEZFWKpGUWGD9EBClAdollTV7YQHT0AW6Oq3ahU/edit"
                          className="w-full px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-xs font-mono outline-none focus:border-emerald-600"
                        />
                        <button
                          type="button"
                          onClick={handleSaveCustomSpreadsheetLink}
                          disabled={isSyncingSheets}
                          className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer whitespace-nowrap shadow-xs"
                        >
                          <i className="fas fa-check-circle mr-1.5"></i> Simpan
                          & Hubungkan Spreadsheet
                        </button>
                      </div>
                    </div>

                    {/* VERCEL DOMAIN & CONNECTION MODE STATUS CARDS */}
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                      <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
                        <div className="text-xs text-slate-500 mb-1">
                          Mode Koneksi Database Aktif
                        </div>
                        <div className="font-bold text-sm text-slate-900 flex items-center gap-1.5">
                          {!needsGoogleAuth ? (
                            <>
                              <span className="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>
                              <span>
                                Google OAuth ({googleUser?.email || 'Aktif'})
                              </span>
                            </>
                          ) : getSavedGasBridgeUrl() ? (
                            <>
                              <span className="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>
                              <span>
                                Bridge Apps Script (Baca & Tulis Tanpa OAuth)
                              </span>
                            </>
                          ) : (
                            <>
                              <span className="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>
                              <span>
                                Koneksi Langsung (Baca Otomatis Tanpa Login)
                              </span>
                            </>
                          )}
                        </div>
                        <div className="text-xs text-slate-500 mt-1 font-mono">
                          Domain:{' '}
                          {typeof window !== 'undefined'
                            ? window.location.hostname
                            : '-'}
                        </div>
                      </div>

                      <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
                        <div className="text-xs text-slate-500 mb-1">
                          Nama File & Tab Terdeteksi
                        </div>
                        <div className="font-bold text-sm text-slate-900">
                          {sheetMeta?.spreadsheetTitle ||
                            'Klik "Baca Data dari Spreadsheet"'}
                        </div>
                        <div className="text-xs text-slate-500 mt-1">
                          Tab:{' '}
                          {sheetMeta?.tabs.length
                            ? sheetMeta.tabs.join(', ')
                            : 'Sekolah, Kelas, Siswa, Users, Nilai'}
                        </div>
                      </div>

                      <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
                        <div className="text-xs text-slate-500 mb-1">
                          Sinkronisasi Terakhir
                        </div>
                        <div className="font-bold text-sm text-slate-900 font-mono">
                          {sheetMeta?.lastSyncedAt || '-'}
                        </div>
                        <div className="text-xs text-slate-500 mt-1">
                          Sheet Utama (gid=0): {sheetMeta?.gid0Title || 'Sheet1'}
                        </div>
                      </div>
                    </div>

                    {/* VERCEL DEPLOYMENT & GOOGLE LOGIN FIX PANEL */}
                    <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 pt-2">
                      {/* SOLUTION 1: GOOGLE APPS SCRIPT WEB APP BRIDGE (ZERO OAUTH ON VERCEL) */}
                      <div className="p-5 rounded-xl bg-blue-50/70 border border-blue-200 space-y-3">
                        <div className="flex items-start justify-between gap-2">
                          <div>
                            <span className="inline-block px-2 py-0.5 rounded bg-blue-600 text-white text-[10px] font-bold uppercase tracking-wide mb-1">
                              Solusi 1 · Paling Stabil untuk Vercel (Tanpa Login Google)
                            </span>
                            <h6 className="font-bold text-sm text-slate-900">
                              <i className="fas fa-bolt text-blue-600 mr-1.5"></i>
                              Bridge Web App Google Apps Script (Baca & Simpan Otomatis)
                            </h6>
                          </div>
                          <button
                            type="button"
                            onClick={() =>
                              setShowGasScriptModal(!showGasScriptModal)
                            }
                            className="px-2.5 py-1.5 bg-white hover:bg-blue-100 text-blue-700 border border-blue-300 rounded-lg text-xs font-bold cursor-pointer whitespace-nowrap"
                          >
                            <i className="fas fa-code mr-1"></i>
                            {showGasScriptModal
                              ? 'Sembunyikan Kode'
                              : 'Lihat & Salin Kode Script'}
                          </button>
                        </div>

                        <p className="text-xs text-slate-600 leading-relaxed">
                          Saat di-deploy ke <b>Vercel</b>, Anda dapat menyimpan &
                          membaca database Spreadsheet secara langsung{' '}
                          <b>tanpa perlu tombol Sign in with Google</b> (bebas
                          kendala blokir domain OAuth) dengan menempelkan URL
                          Web App dari menu{' '}
                          <b>Ekstensi &rarr; Apps Script</b> di Spreadsheet Anda.
                        </p>

                        {showGasScriptModal && (
                          <div className="p-3 rounded-lg bg-slate-900 text-slate-100 text-[11px] font-mono space-y-2">
                            <div className="flex items-center justify-between border-b border-slate-700 pb-2">
                              <span className="text-amber-300 font-sans font-bold text-xs">
                                Langkah Pasang di Spreadsheet (Hanya 1 Kali):
                              </span>
                              <button
                                type="button"
                                onClick={() => {
                                  navigator.clipboard.writeText(
                                    getGasBridgeScriptCode()
                                  );
                                  showToast(
                                    'Kode Google Apps Script berhasil disalin!',
                                    'success'
                                  );
                                }}
                                className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded text-xs font-sans font-bold cursor-pointer"
                              >
                                <i className="fas fa-copy mr-1"></i> Salin Kode
                              </button>
                            </div>
                            <ol className="list-decimal list-inside font-sans text-xs text-slate-300 space-y-1">
                              <li>
                                Buka Spreadsheet Anda &rarr; klik menu{' '}
                                <b>Ekstensi &rarr; Apps Script</b>.
                              </li>
                              <li>
                                Hapus kode yang ada, tempelkan kode di bawah ini,
                                lalu klik ikon <b>Simpan</b>.
                              </li>
                              <li>
                                Klik tombol{' '}
                                <b>Terapkan (Deploy) &rarr; Deployment baru</b>{' '}
                                &rarr; Pilih jenis <b>Aplikasi Web (Web App)</b>.
                              </li>
                              <li>
                                Pada bagian <i>Siapa yang memiliki akses</i>,
                                pilih <b>Siapa saja (Anyone)</b> &rarr; klik{' '}
                                <b>Terapkan</b> &rarr; Salin <b>URL Aplikasi Web</b>.
                              </li>
                            </ol>
                            <pre className="overflow-x-auto max-h-48 p-2.5 bg-slate-950 rounded border border-slate-800 text-[10px] leading-relaxed">
                              {getGasBridgeScriptCode()}
                            </pre>
                          </div>
                        )}

                        <div className="space-y-2">
                          <label className="block text-xs font-bold text-slate-700">
                            URL Web App Google Apps Script (
                            <span className="font-mono text-[11px] text-slate-500">
                              https://script.google.com/macros/s/.../exec
                            </span>
                            )
                          </label>
                          <div className="flex flex-wrap sm:flex-nowrap gap-2">
                            <input
                              type="url"
                              value={gasBridgeUrlInput}
                              onChange={(e) =>
                                setGasBridgeUrlInput(e.target.value)
                              }
                              placeholder="Tempelkan URL Web App Google Apps Script di sini..."
                              className="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg text-xs font-mono outline-none focus:border-blue-600"
                            />
                            <button
                              type="button"
                              onClick={async () => {
                                saveGasBridgeUrl(gasBridgeUrlInput);
                                if (gasBridgeUrlInput.trim()) {
                                  showToast(
                                    'URL Bridge Apps Script disimpan! Menguji koneksi ke Spreadsheet...',
                                    'success'
                                  );
                                  await loadFromGoogleSheets();
                                } else {
                                  showToast(
                                    'URL Bridge Apps Script dikosongkan.',
                                    'success'
                                  );
                                }
                              }}
                              disabled={isSyncingSheets}
                              className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg cursor-pointer whitespace-nowrap"
                            >
                              <i className="fas fa-save mr-1.5"></i> Simpan &
                              Uji Bridge
                            </button>
                          </div>
                        </div>
                      </div>

                      {/* SOLUTION 2: AUTHORIZE VERCEL DOMAIN IN FIREBASE / GOOGLE CLOUD CONSOLE */}
                      <div
                        className={`p-5 rounded-xl border space-y-3 ${
                          vercelDomainError
                            ? 'bg-amber-50 border-amber-400 ring-2 ring-amber-300'
                            : 'bg-slate-50 border-slate-200'
                        }`}
                      >
                        <div>
                          <span className="inline-block px-2 py-0.5 rounded bg-amber-600 text-white text-[10px] font-bold uppercase tracking-wide mb-1">
                            Solusi 2 · Mengaktifkan Popup Login Google di Domain Vercel
                          </span>
                          <h6 className="font-bold text-sm text-slate-900">
                            <i className="fas fa-shield-alt text-amber-600 mr-1.5"></i>
                            Otorisasi Domain Vercel pada Firebase & Google Cloud
                          </h6>
                        </div>

                        <p className="text-xs text-slate-600 leading-relaxed">
                          Mengapa tombol <b>Sign in with Google</b> gagal saat
                          di-deploy ke Vercel? Karena Google / Firebase memblokir
                          popup login dari domain baru yang belum didaftarkan di{' '}
                          <b>Authorized Domains</b>. Salin domain Vercel Anda di
                          bawah ini lalu tambahkan ke Firebase Console:
                        </p>

                        <div className="p-3 bg-white rounded-lg border border-slate-200 flex items-center justify-between gap-2">
                          <div>
                            <div className="text-[10px] uppercase font-bold text-slate-400">
                              Domain Aplikasi Saat Ini (Hostname):
                            </div>
                            <div className="font-mono font-bold text-xs text-slate-900">
                              {typeof window !== 'undefined'
                                ? window.location.hostname
                                : 'nama-aplikasi.vercel.app'}
                            </div>
                          </div>
                          <button
                            type="button"
                            onClick={() => {
                              const host =
                                typeof window !== 'undefined'
                                  ? window.location.hostname
                                  : '';
                              navigator.clipboard.writeText(host);
                              showToast(
                                `Domain "${host}" berhasil disalin! Tempelkan di Authorized Domains Firebase Console.`,
                                'success'
                              );
                            }}
                            className="px-3 py-1.5 bg-slate-800 hover:bg-slate-900 text-white text-xs font-bold rounded-lg cursor-pointer whitespace-nowrap"
                          >
                            <i className="fas fa-copy mr-1.5"></i> Salin Domain
                          </button>
                        </div>

                        <div className="flex flex-wrap gap-2">
                          <a
                            href={`https://console.firebase.google.com/project/${FIREBASE_PROJECT_ID}/authentication/settings`}
                            target="_blank"
                            rel="noreferrer"
                            className="px-3 py-2 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold rounded-lg inline-flex items-center gap-1.5"
                          >
                            <i className="fas fa-external-link-alt"></i>
                            1. Buka Firebase Authorized Domains
                          </a>
                          <a
                            href={`https://console.cloud.google.com/apis/credentials?project=${FIREBASE_PROJECT_ID}`}
                            target="_blank"
                            rel="noreferrer"
                            className="px-3 py-2 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 text-xs font-bold rounded-lg inline-flex items-center gap-1.5"
                          >
                            <i className="fas fa-key"></i>
                            2. Buka GCP OAuth Credentials
                          </a>
                        </div>

                        <div className="pt-1 border-t border-slate-200/80 space-y-1.5">
                          <label className="block text-[11px] font-bold text-slate-600">
                            Opsional: Gunakan Google OAuth Client ID Milik Sendiri
                            (Jika membuat project GCP sendiri untuk Vercel)
                          </label>
                          <div className="flex gap-2">
                            <input
                              type="text"
                              value={customOAuthClientId}
                              onChange={(e) => {
                                const val = e.target.value;
                                setCustomOAuthClientId(val);
                                if (typeof window !== 'undefined') {
                                  if (val.trim()) {
                                    localStorage.setItem(
                                      'eraport_custom_oauth_client_id',
                                      val.trim()
                                    );
                                  } else {
                                    localStorage.removeItem(
                                      'eraport_custom_oauth_client_id'
                                    );
                                  }
                                }
                              }}
                              placeholder={`Default: ${GOOGLE_OAUTH_CLIENT_ID.slice(
                                0,
                                28
                              )}...`}
                              className="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-mono outline-none focus:border-blue-600"
                            />
                          </div>
                        </div>
                      </div>
                    </div>

                    {sheetMeta && sheetMeta.gid0PreviewRows.length > 0 && (
                      <div>
                        <h6 className="font-bold text-xs text-slate-700 uppercase mb-2">
                          Preview Data Lembar Pertama (gid=0:{' '}
                          {sheetMeta.gid0Title})
                        </h6>
                        <div className="overflow-x-auto border border-slate-200 rounded-lg">
                          <table className="w-full text-xs text-left border-collapse font-mono">
                            <tbody>
                              {sheetMeta.gid0PreviewRows.map((row, rIdx) => (
                                <tr
                                  key={rIdx}
                                  className={
                                    rIdx === 0
                                      ? 'bg-slate-100 font-bold text-slate-800 border-b border-slate-200'
                                      : 'border-b border-slate-100'
                                  }
                                >
                                  {row.slice(0, 10).map((cell, cIdx) => (
                                    <td
                                      key={cIdx}
                                      className="py-1.5 px-2.5 truncate max-w-[180px]"
                                    >
                                      {cell}
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* ================= ROLE: GURU MAPEL ================= */}
            {activeRole === 'gurumapel' && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 flex flex-wrap justify-between items-center gap-4">
                  <div>
                    <div className="text-xs font-semibold text-blue-600 mb-1">
                      Portal Guru Mata Pelajaran
                    </div>
                    <h4 className="font-bold text-xl text-slate-900 mb-1">
                      Input Nilai Sumatif Tengah Semester (STS) Per Mata
                      Pelajaran
                    </h4>
                    <p className="text-slate-500 text-xs">
                      Pilih Rombel dan Mata Pelajaran yang Anda ampu untuk
                      mengisi nilai atau mengimpor template khusus mata
                      pelajaran tersebut.
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={() =>
                        downloadGradesTemplate(
                          guruSelectedClass,
                          currentClassStudents,
                          guruSelectedSubject,
                          availableSubjectsForSelectedClass.map((s) => s.name),
                          studentGrades,
                          'single_subject'
                        )
                      }
                      className="px-3 py-2 border border-emerald-600 text-emerald-700 hover:bg-emerald-50 text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                    >
                      <i className="fas fa-file-excel mr-1.5"></i> Unduh
                      Template Mapel Ini (.xlsx)
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setBulkModalType('nilai');
                        setBulkPasteText('');
                      }}
                      className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer whitespace-nowrap"
                    >
                      <i className="fas fa-file-import mr-1.5"></i> Impor Nilai
                      Mapel Ini
                    </button>
                  </div>
                </div>

                {/* Selector Card */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
                  <div className="grid grid-cols-1 md:grid-cols-12 gap-4 items-end">
                    <div className="md:col-span-4">
                      <label className="block text-xs font-bold text-slate-600 mb-1">
                        Pilih Kelas / Rombel
                      </label>
                      <select
                        value={guruSelectedClass}
                        onChange={(e) => setGuruSelectedClass(e.target.value)}
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-semibold outline-none focus:border-blue-600"
                      >
                        {classes.map((c) => (
                          <option key={c.id} value={c.name}>
                            {c.name} (Kelas {c.grade})
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="md:col-span-5">
                      <label className="block text-xs font-bold text-slate-600 mb-1">
                        Pilih Mata Pelajaran
                      </label>
                      <select
                        value={guruSelectedSubject}
                        onChange={(e) => setGuruSelectedSubject(e.target.value)}
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-semibold outline-none focus:border-blue-600"
                      >
                        {availableSubjectsForSelectedClass.map((sub) => (
                          <option key={sub.name} value={sub.name}>
                            [{sub.category}] {sub.name}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="md:col-span-3">
                      <button
                        type="button"
                        onClick={fillRandomGrades}
                        className="w-full py-2 px-3 border border-slate-300 hover:bg-slate-100 text-slate-700 font-bold text-xs rounded-lg transition-colors cursor-pointer"
                      >
                        <i className="fas fa-magic mr-1.5"></i> Isian Otomatis
                        Sample
                      </button>
                    </div>
                  </div>
                </div>

                {/* Grade Input Table */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                  <div className="flex flex-wrap justify-between items-center border-b border-slate-200 pb-4 mb-4 gap-3">
                    <div>
                      <h5 className="font-bold text-base text-slate-900 mb-0">
                        Input Nilai - {guruSelectedClass}
                      </h5>
                      <div className="text-blue-600 font-semibold text-xs mt-0.5">
                        Mata Pelajaran: {guruSelectedSubject}
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={handleSaveGrades}
                      className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer"
                    >
                      <i className="fas fa-save mr-1.5"></i> Simpan Semua Nilai
                    </button>
                  </div>

                  {currentClassStudents.length === 0 ? (
                    <div className="text-center py-10 text-slate-400">
                      <i className="fas fa-user-slash fa-3x mb-3"></i>
                      <p className="text-sm">
                        Tidak ada siswa di kelas {guruSelectedClass}.
                      </p>
                    </div>
                  ) : (
                    <div className="overflow-x-auto">
                      <table className="w-full text-sm text-left border-collapse">
                        <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                          <tr>
                            <th className="py-3 px-4 w-12">No</th>
                            <th className="py-3 px-4">Nama Siswa</th>
                            <th className="py-3 px-4">NISN / NIS</th>
                            <th className="py-3 px-4 w-44 text-center">
                              Nilai STS (0 - 100)
                            </th>
                            <th className="py-3 px-4 text-center">Status</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {currentClassStudents.map((student, idx) => {
                            const val =
                              studentGrades[student.id]?.[guruSelectedSubject];
                            const isFilled =
                              val !== undefined && String(val) !== '';
                            return (
                              <tr
                                key={student.id}
                                className="hover:bg-slate-50"
                              >
                                <td className="py-3 px-4 text-slate-500 font-bold font-mono">
                                  {idx + 1}
                                </td>
                                <td className="py-3 px-4 font-bold text-slate-900">
                                  {student.name}
                                </td>
                                <td className="py-3 px-4 font-mono text-xs text-slate-500">
                                  {student.nisn} / {student.nis}
                                </td>
                                <td className="py-3 px-4">
                                  <input
                                    type="number"
                                    min={0}
                                    max={100}
                                    value={isFilled ? val : ''}
                                    onChange={(e) =>
                                      handleGradeChange(
                                        student.id,
                                        guruSelectedSubject,
                                        e.target.value
                                      )
                                    }
                                    placeholder="0"
                                    className="w-full px-3 py-1.5 border border-slate-300 rounded-lg text-center font-bold font-mono text-slate-900 outline-none focus:border-blue-600"
                                  />
                                </td>
                                <td className="py-3 px-4 text-center">
                                  {isFilled ? (
                                    <span className="text-xs font-semibold text-emerald-700">
                                      <i className="fas fa-check-circle mr-1"></i>
                                      Terisi
                                    </span>
                                  ) : (
                                    <span className="text-xs font-semibold text-amber-600">
                                      Kosong
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* ================= ROLE: WALI KELAS ================= */}
            {activeRole === 'walikelas' && (
              <div className="space-y-6">
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-5 flex flex-wrap justify-between items-center gap-4">
                  <div>
                    <div className="text-xs font-semibold text-emerald-600 mb-1">
                      Portal Wali Kelas
                    </div>
                    <h4 className="font-bold text-xl text-slate-900 mb-1">
                      Rekapitulasi, Kehadiran, Catatan & Cetak Raport
                    </h4>
                    <div className="text-slate-600 text-xs flex flex-wrap items-center gap-2 mt-1">
                      <span>
                        Wali Kelas Rombel <b>{waliSelectedClass}</b>:{' '}
                        <b className="text-slate-900">
                          {formatTeacherNameWithTitle(
                            activeWaliClassInfo.waliKelasName
                          )}
                        </b>
                        {formatNipOrNuptk(activeWaliClassInfo.waliKelasNip)
                          .formatted && (
                          <>
                            {' '}
                            (
                            <span className="font-mono">
                              {
                                formatNipOrNuptk(
                                  activeWaliClassInfo.waliKelasNip
                                ).formatted
                              }
                            </span>
                            )
                          </>
                        )}
                      </span>
                      <button
                        type="button"
                        onClick={() => {
                          const foundCls = classes.find(
                            (c) =>
                              normalizeClassKey(c.name) ===
                              normalizeClassKey(waliSelectedClass)
                          );
                          if (foundCls) {
                            setEditingClass({
                              ...foundCls,
                              waliKelasName:
                                activeWaliClassInfo.waliKelasName || '',
                              waliKelasNip:
                                activeWaliClassInfo.waliKelasNip || '',
                              waliSignatureType:
                                activeWaliClassInfo.waliSignatureType || 'none',
                              waliSignatureUrl:
                                activeWaliClassInfo.waliSignatureUrl || '',
                            });
                          }
                        }}
                        className="px-2.5 py-1 bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 font-bold rounded-lg text-xs cursor-pointer inline-flex items-center gap-1.5"
                      >
                        <i className="fas fa-signature"></i> Atur Nama, NIP &
                        Upload Scan TTD / QR Wali Kelas
                      </button>
                      {activeWaliClassInfo.waliSignatureType &&
                        activeWaliClassInfo.waliSignatureType !== 'none' &&
                        activeWaliClassInfo.waliSignatureUrl && (
                          <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-700 font-bold text-[11px]">
                            <i className="fas fa-check-circle"></i>
                            {activeWaliClassInfo.waliSignatureType === 'qrcode'
                              ? 'QR Code Wali Kelas Aktif'
                              : 'Scan TTD Wali Kelas Aktif'}
                          </span>
                        )}
                    </div>
                  </div>
                  <div className="flex items-center gap-2 bg-emerald-50 border border-emerald-200 px-3.5 py-2 rounded-lg">
                    <i className="fas fa-lock text-emerald-600 text-xs"></i>
                    <span className="font-bold text-xs text-slate-600">
                      Kelas Terkunci:
                    </span>
                    <span className="px-2.5 py-0.5 bg-emerald-600 text-white rounded text-xs font-bold">
                      {waliSelectedClass}
                    </span>
                  </div>
                </div>

                {/* Print Banner */}
                <div className="bg-slate-900 text-white rounded-xl shadow-sm p-5 flex flex-wrap justify-between items-center gap-4">
                  <div>
                    <div className="text-xs font-bold text-amber-400 mb-1">
                      Rombel {waliSelectedClass}
                    </div>
                    <h5 className="font-bold text-lg mb-1">
                      Cetak & Unduh PDF Raport Kelas {waliSelectedClass}
                    </h5>
                    <p className="text-slate-300 text-xs">
                      Unduh PDF raport langsung per siswa atau seluruh siswa
                      sekaligus dalam format kertas A4 siap cetak.
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={handleSaveWaliData}
                      className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer"
                    >
                      <i className="fas fa-save mr-1.5"></i> Simpan Kehadiran &
                      Catatan
                    </button>
                    <button
                      type="button"
                      onClick={openClassPrintPreview}
                      className="px-5 py-2.5 bg-amber-400 hover:bg-amber-500 text-slate-900 font-bold text-xs rounded-lg transition-colors cursor-pointer"
                    >
                      <i className="fas fa-file-pdf mr-1.5"></i> Preview &
                      Download PDF Sekelas ({currentClassStudents.length} Siswa)
                    </button>
                  </div>
                </div>

                {/* Student Grades & Attendance Table */}
                <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                  <div className="flex flex-wrap justify-between items-center border-b border-slate-200 pb-3 mb-4 gap-2">
                    <h5 className="font-bold text-base text-slate-900 mb-0">
                      <i className="fas fa-table text-blue-600 mr-2"></i> Leger
                      Nilai & Kehadiran
                    </h5>
                    <span className="text-xs text-slate-500">
                      Klik tombol <b>Preview / PDF</b> pada siswa untuk melihat
                      dan mengunduh file PDF raport A4.
                    </span>
                  </div>

                  {rankedClassStudents.length === 0 ? (
                    <div className="text-center py-10 text-slate-400">
                      <i className="fas fa-folder-open fa-3x mb-3"></i>
                      <p className="text-sm">
                        Tidak ada data siswa untuk kelas {waliSelectedClass}.
                      </p>
                    </div>
                  ) : (
                    <div className="overflow-x-auto">
                      <table className="w-full text-sm text-left border-collapse">
                        <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                          <tr>
                            <th className="py-3 px-2.5 text-center w-10">No</th>
                            <th className="py-3 px-3 min-w-[180px]">
                              Nama Siswa
                            </th>
                            <th className="py-3 px-2 text-center w-20 whitespace-nowrap">
                              Total
                            </th>
                            <th className="py-3 px-2 text-center w-20 whitespace-nowrap">
                              Rata-Rata
                            </th>
                            <th className="py-3 px-2 text-center w-20 whitespace-nowrap">
                              Peringkat
                            </th>
                            <th className="py-3 px-2.5 text-center w-44 whitespace-nowrap">
                              Absensi (S / I / A)
                            </th>
                            <th className="py-3 px-3 min-w-[340px] md:min-w-[440px] w-[42%]">
                              Catatan Wali Kelas
                            </th>
                            <th className="py-3 px-3 text-center w-32 whitespace-nowrap">
                              Aksi Raport
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {rankedClassStudents.map((student, idx) => (
                            <tr key={student.id} className="hover:bg-slate-50">
                              <td className="py-3 px-2.5 text-center text-slate-500 font-bold font-mono">
                                {idx + 1}
                              </td>
                              <td className="py-3 px-3">
                                <div className="font-bold text-slate-900">
                                  {student.name}
                                </div>
                                <div className="text-xs text-slate-500 font-mono">
                                  NISN: {student.nisn}
                                </div>
                              </td>
                              <td className="py-3 px-2 text-center font-bold font-mono text-slate-900">
                                {student.totalScore}
                              </td>
                              <td className="py-3 px-2 text-center font-bold font-mono text-blue-600">
                                {student.avgScore}
                              </td>
                              <td className="py-3 px-2 text-center font-mono font-bold text-amber-700 whitespace-nowrap">
                                Rank #{student.rank}
                              </td>
                              <td className="py-3 px-2.5">
                                <div className="flex gap-1.5 justify-center">
                                  <input
                                    title="Sakit"
                                    type="number"
                                    min={0}
                                    value={student.attendance.sakit}
                                    onChange={(e) =>
                                      handleAttendanceChange(
                                        student.id,
                                        'sakit',
                                        e.target.value
                                      )
                                    }
                                    className="w-12 px-1.5 py-1.5 border border-slate-300 rounded-lg text-center text-xs font-mono font-semibold focus:border-blue-600 outline-none"
                                    placeholder="S"
                                  />
                                  <input
                                    title="Izin"
                                    type="number"
                                    min={0}
                                    value={student.attendance.izin}
                                    onChange={(e) =>
                                      handleAttendanceChange(
                                        student.id,
                                        'izin',
                                        e.target.value
                                      )
                                    }
                                    className="w-12 px-1.5 py-1.5 border border-slate-300 rounded-lg text-center text-xs font-mono font-semibold focus:border-blue-600 outline-none"
                                    placeholder="I"
                                  />
                                  <input
                                    title="Alpa"
                                    type="number"
                                    min={0}
                                    value={student.attendance.alpa}
                                    onChange={(e) =>
                                      handleAttendanceChange(
                                        student.id,
                                        'alpa',
                                        e.target.value
                                      )
                                    }
                                    className="w-12 px-1.5 py-1.5 border border-slate-300 rounded-lg text-center text-xs font-mono font-semibold focus:border-blue-600 outline-none"
                                    placeholder="A"
                                  />
                                </div>
                              </td>
                              <td className="py-3 px-3 min-w-[340px] md:min-w-[440px]">
                                <textarea
                                  rows={2}
                                  value={student.note}
                                  onChange={(e) =>
                                    handleNoteChange(student.id, e.target.value)
                                  }
                                  className="w-full min-w-[320px] px-3.5 py-2 border border-slate-300 rounded-lg text-sm text-slate-800 leading-relaxed focus:border-blue-600 focus:ring-2 focus:ring-blue-100 outline-none resize-y"
                                  placeholder="Tulis catatan perkembangan belajar siswa..."
                                />
                              </td>
                              <td className="py-3 px-3 text-center">
                                <button
                                  type="button"
                                  onClick={() =>
                                    openSingleStudentPreview(student)
                                  }
                                  className="px-3 py-1.5 border border-blue-600 text-blue-600 hover:bg-blue-600 hover:text-white rounded-lg text-xs font-bold transition-colors cursor-pointer whitespace-nowrap"
                                >
                                  <i className="fas fa-file-pdf mr-1"></i>{' '}
                                  Preview / PDF
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </div>
            )}
          </main>

          {/* MANDATORY & COPYRIGHT FOOTER */}
          <footer className="no-print bg-white border-t border-slate-200 py-4 px-4 mt-auto">
            <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3 text-xs text-slate-600">
              <div className="flex items-center gap-2.5">
                {!logoImgError && school.logoUrl && (
                  <img
                    src={school.logoUrl}
                    alt="Logo"
                    className="w-5 h-5 object-contain"
                    referrerPolicy="no-referrer"
                  />
                )}
                <span>
                  <b>E-Raport STS Kurikulum Merdeka</b> · {school.name} (TP{' '}
                  {school.tp})
                </span>
              </div>
              <div className="text-slate-700">
                Mandatory Copyright © 2026 · Desain & Sistem Dibuat oleh{' '}
                <span className="font-bold text-slate-900">
                  Imam Huzaer, S.Pd.
                </span>
              </div>
            </div>
          </footer>

          {/* MODAL: BULK IMPORT VIA EXCEL TEMPLATE OR COPY-PASTE */}
          {bulkModalType && (
            <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/65 p-4 no-print">
              <div className="bg-white rounded-xl shadow-xl max-w-2xl w-full overflow-hidden">
                <div className="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50">
                  <div>
                    <h5 className="font-bold text-base text-slate-900">
                      <i className="fas fa-file-import text-emerald-600 mr-2"></i>
                      {bulkModalType === 'siswa' &&
                        'Impor Cepat Data Siswa (Template Excel / Copy-Paste)'}
                      {bulkModalType === 'kelas' &&
                        'Impor Cepat Data Kelas & Wali Kelas'}
                      {bulkModalType === 'walikelas_user' &&
                        'Impor Cepat Akun User Wali Kelas'}
                      {bulkModalType === 'gurumapel_user' &&
                        'Impor Cepat Akun User Guru Mata Pelajaran'}
                      {bulkModalType === 'nilai' &&
                        `Impor Nilai Mapel "${guruSelectedSubject}" - Kelas ${guruSelectedClass}`}
                      {bulkModalType === 'nilai_admin_semua_mapel' &&
                        `Impor Massal Nilai Semua Mapel - Kelas ${adminGradeSelectedClass}`}
                    </h5>
                  </div>
                  <button
                    type="button"
                    onClick={() => setBulkModalType(null)}
                    className="text-slate-400 hover:text-slate-600 cursor-pointer"
                  >
                    <i className="fas fa-times"></i>
                  </button>
                </div>

                <div className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
                  {/* Step 1: Download Template */}
                  <div className="p-4 rounded-xl bg-emerald-50/70 border border-emerald-200 flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <div className="text-xs font-bold text-emerald-900">
                        1. Unduh Template Excel (.xlsx)
                      </div>
                      <div className="text-xs text-emerald-700 mt-0.5">
                        Gunakan template resmi agar kolom otomatis sesuai, atau
                        langsung salin (Copy-Paste) baris dari Excel/Spreadsheet.
                      </div>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {bulkModalType === 'siswa' && (
                        <button
                          type="button"
                          onClick={downloadStudentTemplate}
                          className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer"
                        >
                          <i className="fas fa-download mr-1.5"></i> Unduh
                          Template Siswa
                        </button>
                      )}
                      {bulkModalType === 'kelas' && (
                        <button
                          type="button"
                          onClick={downloadClassTemplate}
                          className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer"
                        >
                          <i className="fas fa-download mr-1.5"></i> Unduh
                          Template Kelas
                        </button>
                      )}
                      {bulkModalType === 'walikelas_user' && (
                        <button
                          type="button"
                          onClick={downloadWaliKelasUserTemplate}
                          className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer"
                        >
                          <i className="fas fa-download mr-1.5"></i> Unduh
                          Template Wali Kelas
                        </button>
                      )}
                      {bulkModalType === 'gurumapel_user' && (
                        <button
                          type="button"
                          onClick={downloadGuruMapelUserTemplate}
                          className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer"
                        >
                          <i className="fas fa-download mr-1.5"></i> Unduh
                          Template Guru Mapel
                        </button>
                      )}
                      {bulkModalType === 'nilai' && (
                        <button
                          type="button"
                          onClick={() =>
                            downloadGradesTemplate(
                              guruSelectedClass,
                              currentClassStudents,
                              guruSelectedSubject,
                              availableSubjectsForSelectedClass.map(
                                (s) => s.name
                              ),
                              studentGrades,
                              'single_subject'
                            )
                          }
                          className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer"
                        >
                          <i className="fas fa-download mr-1.5"></i> Unduh
                          Template Mapel Ini
                        </button>
                      )}
                      {bulkModalType === 'nilai_admin_semua_mapel' && (
                        <button
                          type="button"
                          onClick={() =>
                            downloadGradesTemplate(
                              adminGradeSelectedClass,
                              adminGradeClassStudents,
                              adminGradeClassSubjects[0]?.name || '',
                              adminGradeClassSubjects.map((s) => s.name),
                              studentGrades,
                              'all_subjects'
                            )
                          }
                          className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg cursor-pointer"
                        >
                          <i className="fas fa-download mr-1.5"></i> Unduh
                          Template Semua Mapel ({adminGradeSelectedClass})
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Step 2: Upload Excel/CSV File */}
                  <div className="p-4 rounded-xl bg-slate-50 border border-slate-200">
                    <label className="block text-xs font-bold text-slate-800 mb-2">
                      2. Cara A: Pilih / Upload File Template Excel (.xlsx,
                      .xls, .csv)
                    </label>
                    <input
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={handleBulkFileUpload}
                      className="block w-full text-xs text-slate-600 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-bold file:bg-blue-600 file:text-white hover:file:bg-blue-700 cursor-pointer"
                    />
                  </div>

                  {/* Step 3: Or Copy-Paste Directly from Excel/Sheets */}
                  <div>
                    <label className="block text-xs font-bold text-slate-800 mb-1">
                      3. Cara B: Atau Copy-Paste Langsung Baris dari Excel /
                      Google Sheets
                    </label>
                    <p className="text-[11px] text-slate-500 mb-2">
                      {bulkModalType === 'siswa' &&
                        'Format Kolom: Nama Lengkap | NISN | NIS | Kelas | Konsentrasi Keahlian'}
                      {bulkModalType === 'kelas' &&
                        'Format Kolom: Tingkat (10/11/12) | Nama Rombel | Konsentrasi Keahlian | Nama Wali Kelas | NIP/NUPTK'}
                      {bulkModalType === 'walikelas_user' &&
                        'Format Kolom: Nama Lengkap Wali Kelas | NIP/NUPTK | Username | Password | Wali Kelas Untuk (Rombel) | Merangkap Guru Mapel? (Ya/Tidak)'}
                      {bulkModalType === 'gurumapel_user' &&
                        'Format Kolom: Nama Lengkap Guru Mapel | NIP/NUPTK | Username | Password'}
                      {bulkModalType === 'nilai' &&
                        `Format Kolom: NISN atau Nama Siswa | Nilai STS (Khusus mapel ${guruSelectedSubject})`}
                      {bulkModalType === 'nilai_admin_semua_mapel' &&
                        `Format Kolom: No | NISN | NIS | Nama Siswa | diikuti kolom seluruh mata pelajaran kelas ${adminGradeSelectedClass}`}
                    </p>
                    <textarea
                      rows={5}
                      value={bulkPasteText}
                      onChange={(e) => setBulkPasteText(e.target.value)}
                      placeholder="Blok tabel di Excel / Google Sheets lalu tekan Ctrl+V di sini..."
                      className="w-full p-3 border border-slate-300 rounded-lg text-xs font-mono outline-none focus:border-blue-600"
                    />
                  </div>

                  {(bulkModalType === 'siswa' || bulkModalType === 'kelas') && (
                    <label className="flex items-center gap-2 text-xs font-semibold text-slate-700 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={bulkReplaceExisting}
                        onChange={(e) =>
                          setBulkReplaceExisting(e.target.checked)
                        }
                      />
                      <span>
                        Ganti (Timpa) seluruh data lama dengan data yang baru
                        diimpor ini
                      </span>
                    </label>
                  )}
                </div>

                <div className="px-6 py-3 bg-slate-50 border-t border-slate-200 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setBulkModalType(null)}
                    className="px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:bg-slate-200 cursor-pointer"
                  >
                    Tutup
                  </button>
                  <button
                    type="button"
                    onClick={handleBulkPasteSubmit}
                    className="px-4 py-2 rounded-lg text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white cursor-pointer"
                  >
                    <i className="fas fa-check mr-1.5"></i> Proses Data
                    Copy-Paste
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* MODAL: EDIT CLASS, WALI KELAS & TTD/QR WALI KELAS */}
          {editingClass && (
            <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/60 p-4 no-print overflow-y-auto">
              <div className="bg-white rounded-xl shadow-xl max-w-lg w-full overflow-hidden my-8">
                <div className="px-5 py-4 border-b border-slate-200 flex justify-between items-center">
                  <h5 className="font-bold text-base text-slate-900">
                    Edit Wali Kelas & Upload TTD / QR Code ({editingClass.name})
                  </h5>
                  <button
                    type="button"
                    onClick={() => setEditingClass(null)}
                    className="text-slate-400 hover:text-slate-600 cursor-pointer"
                  >
                    <i className="fas fa-times"></i>
                  </button>
                </div>
                <div className="p-5 space-y-3.5 max-h-[80vh] overflow-y-auto">
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Tingkat Kelas
                      </label>
                      <select
                        value={editingClass.grade}
                        disabled={activeRole !== 'admin'}
                        onChange={(e) =>
                          setEditingClass({
                            ...editingClass,
                            grade: e.target.value,
                          })
                        }
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm disabled:bg-slate-100"
                      >
                        <option value="10">Kelas 10</option>
                        <option value="11">Kelas 11</option>
                        <option value="12">Kelas 12</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Nama Rombel
                      </label>
                      <input
                        type="text"
                        value={editingClass.name}
                        disabled={activeRole !== 'admin'}
                        onChange={(e) =>
                          setEditingClass({
                            ...editingClass,
                            name: e.target.value,
                          })
                        }
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm disabled:bg-slate-100"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Konsentrasi Keahlian
                    </label>
                    <input
                      type="text"
                      value={editingClass.major}
                      disabled={activeRole !== 'admin'}
                      onChange={(e) =>
                        setEditingClass({
                          ...editingClass,
                          major: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm disabled:bg-slate-100"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Nama Wali Kelas (Gelar otomatis sesuai EYD: S.Pd., M.Pd.)
                    </label>
                    <input
                      type="text"
                      value={editingClass.waliKelasName}
                      onChange={(e) =>
                        setEditingClass({
                          ...editingClass,
                          waliKelasName: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                      placeholder="Contoh: Imam Huzaer, S.Pd."
                    />
                    <span className="text-[11px] text-slate-500 mt-0.5 block">
                      Tampilan Raport:{' '}
                      <b>
                        {formatTeacherNameWithTitle(editingClass.waliKelasName)}
                      </b>
                    </span>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      NIP / NUPTK Wali Kelas (18 digit NIP / 16 digit NUPTK)
                    </label>
                    <input
                      type="text"
                      value={editingClass.waliKelasNip}
                      onChange={(e) =>
                        setEditingClass({
                          ...editingClass,
                          waliKelasNip: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-mono"
                      placeholder="Contoh: 199205112020121007"
                    />
                    <span className="text-[11px] text-slate-500 mt-0.5 block">
                      Terdeteksi Otomatis:{' '}
                      <b>
                        {formatNipOrNuptk(editingClass.waliKelasNip)
                          .formatted ||
                          'Tanpa NIP/NUPTK (Ditampilkan Hanya Nama Guru)'}
                      </b>
                    </span>
                  </div>

                  {/* FITUR 3: UPLOAD SCAN TTD WALI KELAS ATAU QR CODE WALI KELAS */}
                  <div className="p-4 rounded-xl bg-emerald-50/70 border border-emerald-200 space-y-3">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-slate-800">
                        <i className="fas fa-signature text-emerald-600 mr-1.5"></i>
                        Scan TTD / QR Code Wali Kelas (Tampil di Raport)
                      </span>
                      {editingClass.waliSignatureUrl && (
                        <button
                          type="button"
                          onClick={() =>
                            setEditingClass({
                              ...editingClass,
                              waliSignatureType: 'none',
                              waliSignatureUrl: '',
                            })
                          }
                          className="text-[11px] font-bold text-red-600 hover:underline cursor-pointer"
                        >
                          <i className="fas fa-trash-alt mr-1"></i> Hapus
                        </button>
                      )}
                    </div>

                    <div className="grid grid-cols-3 gap-1.5">
                      <button
                        type="button"
                        onClick={() =>
                          setEditingClass({
                            ...editingClass,
                            waliSignatureType: 'none',
                          })
                        }
                        className={`px-2.5 py-1.5 rounded-lg text-[11px] font-bold border cursor-pointer ${
                          !editingClass.waliSignatureType ||
                          editingClass.waliSignatureType === 'none'
                            ? 'bg-slate-800 text-white border-slate-800'
                            : 'bg-white text-slate-700 border-slate-300'
                        }`}
                      >
                        TTD Manual
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          setEditingClass({
                            ...editingClass,
                            waliSignatureType: 'signature',
                          })
                        }
                        className={`px-2.5 py-1.5 rounded-lg text-[11px] font-bold border cursor-pointer ${
                          editingClass.waliSignatureType === 'signature'
                            ? 'bg-blue-600 text-white border-blue-600'
                            : 'bg-white text-slate-700 border-slate-300'
                        }`}
                      >
                        Scan TTD Wali
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          setEditingClass({
                            ...editingClass,
                            waliSignatureType: 'qrcode',
                          })
                        }
                        className={`px-2.5 py-1.5 rounded-lg text-[11px] font-bold border cursor-pointer ${
                          editingClass.waliSignatureType === 'qrcode'
                            ? 'bg-emerald-600 text-white border-emerald-600'
                            : 'bg-white text-slate-700 border-slate-300'
                        }`}
                      >
                        QR Code Wali
                      </button>
                    </div>

                    <div className="grid grid-cols-2 gap-2">
                      <label className="flex items-center justify-center gap-1.5 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg cursor-pointer">
                        <i className="fas fa-file-signature"></i>
                        <span>Upload Scan TTD</span>
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) =>
                            handleUploadEditingClassWaliSigOrQr(e, 'signature')
                          }
                          className="hidden"
                        />
                      </label>
                      <label className="flex items-center justify-center gap-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer">
                        <i className="fas fa-qrcode"></i>
                        <span>Upload QR Code</span>
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) =>
                            handleUploadEditingClassWaliSigOrQr(e, 'qrcode')
                          }
                          className="hidden"
                        />
                      </label>
                    </div>

                    <div className="flex gap-2">
                      <input
                        type="text"
                        value={editingClass.waliSignatureUrl || ''}
                        onChange={(e) => {
                          const val = e.target.value;
                          setEditingClass({
                            ...editingClass,
                            waliSignatureUrl: val,
                            waliSignatureType:
                              val &&
                              (!editingClass.waliSignatureType ||
                                editingClass.waliSignatureType === 'none')
                                ? 'signature'
                                : editingClass.waliSignatureType,
                          });
                        }}
                        placeholder="Atau tempel URL gambar TTD / QR Code..."
                        className="w-full px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-mono outline-none focus:border-blue-600"
                      />
                      {editingClass.waliSignatureUrl &&
                        editingClass.waliSignatureType === 'signature' && (
                          <button
                            type="button"
                            onClick={async () => {
                              try {
                                const cleaned =
                                  await removeWhiteBackgroundFromDataUrl(
                                    editingClass.waliSignatureUrl!
                                  );
                                setEditingClass({
                                  ...editingClass,
                                  waliSignatureUrl: cleaned,
                                });
                                showToast(
                                  'Background putih Scan TTD Wali Kelas berhasil dibuat transparan!',
                                  'success'
                                );
                              } catch {
                                showToast(
                                  'Gagal memproses background gambar.',
                                  'danger'
                                );
                              }
                            }}
                            className="px-2.5 py-1.5 bg-white hover:bg-slate-100 text-slate-700 border border-slate-300 rounded-lg text-[11px] font-bold cursor-pointer whitespace-nowrap"
                          >
                            <i className="fas fa-magic mr-1 text-blue-600"></i>
                            Transparankan BG
                          </button>
                        )}
                    </div>

                    {/* Preview Wali Kelas Signature */}
                    <div className="bg-white rounded-lg border border-slate-200 p-3 text-center">
                      <div className="text-[10px] font-bold uppercase text-slate-400 mb-1">
                        Preview Tanda Tangan Wali Kelas di Raport
                      </div>
                      <div
                        style={{
                          fontFamily: "'Times New Roman', Times, serif",
                          fontSize: '9.5pt',
                          color: '#000',
                        }}
                      >
                        <p className="mb-0">{school.datePrint}</p>
                        <p className="mb-0">Wali Kelas,</p>
                        {editingClass.waliSignatureType &&
                        editingClass.waliSignatureType !== 'none' &&
                        editingClass.waliSignatureUrl ? (
                          <div className="h-12 flex items-center justify-center my-1">
                            <img
                              src={editingClass.waliSignatureUrl}
                              alt="Preview TTD/QR Wali"
                              className={
                                editingClass.waliSignatureType === 'qrcode'
                                  ? 'w-11 h-11 object-contain mx-auto'
                                  : 'max-h-11 max-w-[150px] object-contain mx-auto'
                              }
                            />
                          </div>
                        ) : (
                          <div className="h-10 flex items-center justify-center text-[11px] font-sans text-slate-400 italic">
                            (Ruang Tanda Tangan Manual)
                          </div>
                        )}
                        <p className="font-bold underline mb-0">
                          {formatTeacherNameWithTitle(
                            editingClass.waliKelasName
                          )}
                        </p>
                        {formatNipOrNuptk(editingClass.waliKelasNip)
                          .formatted && (
                          <p className="mb-0" style={{ fontSize: '8.5pt' }}>
                            {
                              formatNipOrNuptk(editingClass.waliKelasNip)
                                .formatted
                            }
                          </p>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
                <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setEditingClass(null)}
                    className="px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:bg-slate-200 cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="button"
                    onClick={handleSaveEditedClass}
                    className="px-4 py-2 rounded-lg text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white cursor-pointer"
                  >
                    Simpan Perubahan
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* MODAL: ADD USER (ADMIN ONLY) */}
          {showAddUserModal && (
            <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/60 p-4 no-print">
              <div className="bg-white rounded-xl shadow-xl max-w-md w-full overflow-hidden">
                <div className="px-5 py-4 border-b border-slate-200 flex justify-between items-center">
                  <h5 className="font-bold text-base text-slate-900">
                    Tambah Pengguna Baru
                  </h5>
                  <button
                    type="button"
                    onClick={() => setShowAddUserModal(false)}
                    className="text-slate-400 hover:text-slate-600 cursor-pointer"
                  >
                    <i className="fas fa-times"></i>
                  </button>
                </div>
                <div className="p-5 space-y-3.5">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Nama Lengkap Guru / Staf
                    </label>
                    <input
                      type="text"
                      value={newUser.name}
                      onChange={(e) =>
                        setNewUser({ ...newUser, name: e.target.value })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                      placeholder="Contoh: Imam Huzaer, S.Pd."
                    />
                    <span className="text-[11px] text-slate-500 mt-0.5 block">
                      Format Raport:{' '}
                      <b>{formatTeacherNameWithTitle(newUser.name)}</b>
                    </span>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      NIP (18 digit) atau NUPTK (16 digit)
                    </label>
                    <input
                      type="text"
                      value={newUser.nip}
                      onChange={(e) =>
                        setNewUser({ ...newUser, nip: e.target.value })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-mono"
                      placeholder="199205112020121007"
                    />
                    <span className="text-[11px] text-slate-500 mt-0.5 block">
                      Terdeteksi Otomatis:{' '}
                      <b>
                        {formatNipOrNuptk(newUser.nip).formatted ||
                          'Tanpa NIP/NUPTK (Ditampilkan Hanya Nama Guru)'}
                      </b>
                    </span>
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Username Login
                      </label>
                      <input
                        type="text"
                        value={newUser.username}
                        onChange={(e) =>
                          setNewUser({ ...newUser, username: e.target.value })
                        }
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-mono"
                        placeholder="imam"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Kata Sandi
                      </label>
                      <div className="flex items-center border border-slate-300 rounded-lg overflow-hidden focus-within:border-blue-600">
                        <input
                          type={showNewUserPassword ? 'text' : 'password'}
                          value={newUser.password}
                          onChange={(e) =>
                            setNewUser({ ...newUser, password: e.target.value })
                          }
                          className="w-full px-3 py-2 text-sm font-mono outline-none"
                          placeholder="123456"
                        />
                        <button
                          type="button"
                          onClick={() =>
                            setShowNewUserPassword((prev) => !prev)
                          }
                          title={
                            showNewUserPassword
                              ? 'Sembunyikan Kata Sandi'
                              : 'Lihat Kata Sandi'
                          }
                          className="px-2.5 py-2 text-slate-400 hover:text-slate-700 cursor-pointer"
                        >
                          <i
                            className={
                              showNewUserPassword
                                ? 'fas fa-eye-slash'
                                : 'fas fa-eye'
                            }
                          ></i>
                        </button>
                      </div>
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Pilih Role Hak Akses (Bisa Pilih Banyak / Role Ganda)
                    </label>
                    <div className="p-3 border border-slate-200 rounded-lg bg-slate-50 space-y-2">
                      <label className="flex items-center gap-2 text-xs font-semibold cursor-pointer">
                        <input
                          type="checkbox"
                          checked={newUser.roles.includes('gurumapel')}
                          onChange={() => toggleNewUserRole('gurumapel')}
                        />
                        <span>Guru Mata Pelajaran</span>
                      </label>
                      <label className="flex items-center gap-2 text-xs font-semibold cursor-pointer">
                        <input
                          type="checkbox"
                          checked={newUser.roles.includes('walikelas')}
                          onChange={() => toggleNewUserRole('walikelas')}
                        />
                        <span>Wali Kelas</span>
                      </label>
                      <label className="flex items-center gap-2 text-xs font-semibold cursor-pointer">
                        <input
                          type="checkbox"
                          checked={newUser.roles.includes('admin')}
                          onChange={() => toggleNewUserRole('admin')}
                        />
                        <span>Administrator</span>
                      </label>
                    </div>
                  </div>
                  {newUser.roles.includes('walikelas') && (
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Tugaskan Sebagai Wali Kelas
                      </label>
                      <select
                        value={newUser.waliKelasFor}
                        onChange={(e) =>
                          setNewUser({
                            ...newUser,
                            waliKelasFor: e.target.value,
                          })
                        }
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                      >
                        {classes.map((c) => (
                          <option key={c.id} value={c.name}>
                            {c.name}
                          </option>
                        ))}
                      </select>
                    </div>
                  )}
                </div>
                <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setShowAddUserModal(false)}
                    className="px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:bg-slate-200 cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="button"
                    onClick={handleAddNewUser}
                    className="px-4 py-2 rounded-lg text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white cursor-pointer"
                  >
                    Simpan Pengguna
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* MODAL: ADD STUDENT (ADMIN ONLY) */}
          {showAddStudentModal && (
            <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/60 p-4 no-print">
              <div className="bg-white rounded-xl shadow-xl max-w-md w-full overflow-hidden">
                <div className="px-5 py-4 border-b border-slate-200 flex justify-between items-center">
                  <h5 className="font-bold text-base text-slate-900">
                    Tambah Siswa Baru
                  </h5>
                  <button
                    type="button"
                    onClick={() => setShowAddStudentModal(false)}
                    className="text-slate-400 hover:text-slate-600 cursor-pointer"
                  >
                    <i className="fas fa-times"></i>
                  </button>
                </div>
                <div className="p-5 space-y-3.5">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Nama Lengkap Siswa
                    </label>
                    <input
                      type="text"
                      value={newStudent.name}
                      onChange={(e) =>
                        setNewStudent({ ...newStudent, name: e.target.value })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                      placeholder="Contoh: AHMAD FAUZI"
                    />
                  </div>
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        NISN
                      </label>
                      <input
                        type="text"
                        value={newStudent.nisn}
                        onChange={(e) =>
                          setNewStudent({ ...newStudent, nisn: e.target.value })
                        }
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-mono"
                        placeholder="008..."
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        NIS
                      </label>
                      <input
                        type="text"
                        value={newStudent.nis}
                        onChange={(e) =>
                          setNewStudent({ ...newStudent, nis: e.target.value })
                        }
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-mono"
                        placeholder="2026..."
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Kelas / Rombel
                    </label>
                    <select
                      value={newStudent.className}
                      onChange={(e) =>
                        setNewStudent({
                          ...newStudent,
                          className: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                    >
                      {classes.map((c) => (
                        <option key={c.id} value={c.name}>
                          {c.name}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setShowAddStudentModal(false)}
                    className="px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:bg-slate-200 cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="button"
                    onClick={handleAddNewStudent}
                    className="px-4 py-2 rounded-lg text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white cursor-pointer"
                  >
                    Simpan Siswa
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* MODAL: ADD CLASS (ADMIN ONLY) */}
          {showAddClassModal && (
            <div className="fixed inset-0 z-40 flex items-center justify-center bg-slate-900/60 p-4 no-print">
              <div className="bg-white rounded-xl shadow-xl max-w-md w-full overflow-hidden">
                <div className="px-5 py-4 border-b border-slate-200 flex justify-between items-center">
                  <h5 className="font-bold text-base text-slate-900">
                    Tambah Rombongan Belajar
                  </h5>
                  <button
                    type="button"
                    onClick={() => setShowAddClassModal(false)}
                    className="text-slate-400 hover:text-slate-600 cursor-pointer"
                  >
                    <i className="fas fa-times"></i>
                  </button>
                </div>
                <div className="p-5 space-y-3.5">
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Tingkat Kelas
                      </label>
                      <select
                        value={newClass.grade}
                        onChange={(e) =>
                          setNewClass({ ...newClass, grade: e.target.value })
                        }
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                      >
                        <option value="10">Kelas 10</option>
                        <option value="11">Kelas 11</option>
                        <option value="12">Kelas 12</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">
                        Nama Rombel
                      </label>
                      <input
                        type="text"
                        value={newClass.name}
                        onChange={(e) =>
                          setNewClass({ ...newClass, name: e.target.value })
                        }
                        className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                        placeholder="Contoh: 10 TKR 1"
                      />
                    </div>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Konsentrasi Keahlian
                    </label>
                    <input
                      type="text"
                      value={newClass.major}
                      onChange={(e) =>
                        setNewClass({ ...newClass, major: e.target.value })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                      placeholder="Teknik Kendaraan Ringan"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      Nama Wali Kelas (Gelar sesuai EYD, contoh: S.Pd.)
                    </label>
                    <input
                      type="text"
                      value={newClass.waliKelasName}
                      onChange={(e) =>
                        setNewClass({
                          ...newClass,
                          waliKelasName: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
                      placeholder="Imam Huzaer, S.Pd."
                    />
                    <span className="text-[11px] text-slate-500 mt-0.5 block">
                      Format Raport:{' '}
                      <b>{formatTeacherNameWithTitle(newClass.waliKelasName)}</b>
                    </span>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">
                      NIP (18 digit) atau NUPTK (16 digit) Wali Kelas
                    </label>
                    <input
                      type="text"
                      value={newClass.waliKelasNip}
                      onChange={(e) =>
                        setNewClass({
                          ...newClass,
                          waliKelasNip: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm font-mono"
                      placeholder="199205112020121007"
                    />
                    <span className="text-[11px] text-slate-500 mt-0.5 block">
                      Terdeteksi Otomatis:{' '}
                      <b>
                        {formatNipOrNuptk(newClass.waliKelasNip).formatted ||
                          'Tanpa NIP/NUPTK (Ditampilkan Hanya Nama Guru)'}
                      </b>
                    </span>
                  </div>

                  {/* Upload Scan TTD / QR Code Wali Kelas saat Tambah Kelas */}
                  <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2.5">
                    <label className="block text-xs font-bold text-slate-700">
                      Opsional: Upload Scan TTD / QR Code Wali Kelas
                    </label>
                    <div className="grid grid-cols-2 gap-2">
                      <label className="flex items-center justify-center gap-1.5 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg cursor-pointer">
                        <i className="fas fa-file-signature"></i>
                        <span>Upload Scan TTD</span>
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) =>
                            handleUploadNewClassWaliSigOrQr(e, 'signature')
                          }
                          className="hidden"
                        />
                      </label>
                      <label className="flex items-center justify-center gap-1.5 px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold rounded-lg cursor-pointer">
                        <i className="fas fa-qrcode"></i>
                        <span>Upload QR Code</span>
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) =>
                            handleUploadNewClassWaliSigOrQr(e, 'qrcode')
                          }
                          className="hidden"
                        />
                      </label>
                    </div>
                    {newClass.waliSignatureUrl && (
                      <div className="flex items-center justify-between bg-white p-2 rounded border border-slate-200">
                        <div className="flex items-center gap-2">
                          <img
                            src={newClass.waliSignatureUrl}
                            alt="Preview"
                            className="h-8 w-auto object-contain"
                          />
                          <span className="text-xs font-bold text-emerald-700">
                            {newClass.waliSignatureType === 'qrcode'
                              ? 'QR Code Siap'
                              : 'Scan TTD Siap'}
                          </span>
                        </div>
                        <button
                          type="button"
                          onClick={() =>
                            setNewClass({
                              ...newClass,
                              waliSignatureType: 'none',
                              waliSignatureUrl: '',
                            })
                          }
                          className="text-xs text-red-600 font-bold cursor-pointer"
                        >
                          Hapus
                        </button>
                      </div>
                    )}
                  </div>
                </div>
                <div className="px-5 py-3 bg-slate-50 border-t border-slate-200 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setShowAddClassModal(false)}
                    className="px-4 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:bg-slate-200 cursor-pointer"
                  >
                    Batal
                  </button>
                  <button
                    type="button"
                    onClick={handleAddNewClass}
                    className="px-4 py-2 rounded-lg text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white cursor-pointer"
                  >
                    Simpan Kelas
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* MODAL: PRINT PREVIEW & DIRECT PDF DOWNLOAD */}
          {showPrintModal && (
            <div className="fixed inset-0 z-40 flex flex-col bg-slate-900/80 no-print">
              <div className="bg-white px-6 py-3 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
                <h5 className="font-bold text-base text-slate-900 flex items-center gap-2">
                  <i className="fas fa-file-pdf text-red-600"></i> Preview &
                  Download PDF Raport STS A4
                </h5>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={handleDownloadPdf}
                    disabled={isGeneratingPdf}
                    className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white font-bold text-xs rounded-lg cursor-pointer flex items-center gap-1.5"
                  >
                    <i
                      className={
                        isGeneratingPdf
                          ? 'fas fa-spinner fa-spin'
                          : 'fas fa-download'
                      }
                    ></i>
                    <span>
                      {isGeneratingPdf
                        ? 'Menyiapkan File PDF...'
                        : 'Download PDF A4'}
                    </span>
                  </button>
                  <button
                    type="button"
                    onClick={triggerBrowserPrint}
                    className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-lg cursor-pointer"
                  >
                    <i className="fas fa-print mr-1.5"></i> Cetak Printer
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowPrintModal(false)}
                    className="px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold text-xs rounded-lg cursor-pointer"
                  >
                    Tutup
                  </button>
                </div>
              </div>

              <div
                ref={previewContainerRef}
                className="flex-grow overflow-y-auto p-6 flex flex-col items-center"
              >
                <div className="bg-emerald-50 border border-emerald-200 text-emerald-900 px-4 py-3 rounded-lg text-xs max-w-3xl w-full mb-6 shadow-sm flex items-center justify-between gap-3">
                  <div>
                    <i className="fas fa-check-circle mr-1.5 text-emerald-600"></i>
                    Klik tombol hijau <b>"Download PDF A4"</b> di pojok kanan
                    atas untuk mengunduh langsung file PDF ke perangkat Anda.
                  </div>
                </div>

                {printMode === 'single' && selectedPrintStudent && (
                  <div className="raport-preview">
                    <RaportTemplate
                      student={selectedPrintStudent}
                      school={school}
                      classInfo={getClassInfo(selectedPrintStudent.className)}
                      subjects={getSubjectsForGrade(selectedPrintStudent.grade)}
                      grades={studentGrades[selectedPrintStudent.id] || {}}
                    />
                  </div>
                )}

                {printMode === 'class' &&
                  rankedClassStudents.map((student) => (
                    <div key={student.id} className="raport-preview">
                      <RaportTemplate
                        student={student}
                        school={school}
                        classInfo={getClassInfo(student.className)}
                        subjects={getSubjectsForGrade(student.grade)}
                        grades={studentGrades[student.id] || {}}
                      />
                    </div>
                  ))}
              </div>
            </div>
          )}

          {/* DEDICATED PRINT-ONLY CONTAINER FOR EXACT A4 PAGINATION */}
          <div className="print-only-container">
            {printMode === 'single' && selectedPrintStudent && (
              <div className="raport-page">
                <RaportTemplate
                  student={selectedPrintStudent}
                  school={school}
                  classInfo={getClassInfo(selectedPrintStudent.className)}
                  subjects={getSubjectsForGrade(selectedPrintStudent.grade)}
                  grades={studentGrades[selectedPrintStudent.id] || {}}
                />
              </div>
            )}

            {printMode === 'class' &&
              rankedClassStudents.map((student) => (
                <div key={student.id} className="raport-page">
                  <RaportTemplate
                    student={student}
                    school={school}
                    classInfo={getClassInfo(student.className)}
                    subjects={getSubjectsForGrade(student.grade)}
                    grades={studentGrades[student.id] || {}}
                  />
                </div>
              ))}
          </div>
        </>
      )}
    </div>
  );
}
