import React, { useMemo, useState } from 'react';
import {
  ClassItem,
  StudentGradesMap,
  StudentItem,
  SubjectTemplatesMap,
  UserAccount,
} from '../types';
import {
  formatNipOrNuptk,
  formatTeacherNameWithTitle,
} from '../services/templateAndPdfService';

interface AdminStatisticsDashboardProps {
  classes: ClassItem[];
  students: StudentItem[];
  users: UserAccount[];
  studentGrades: StudentGradesMap;
  subjectTemplates: SubjectTemplatesMap;
  getClassInfo: (className: string) => Partial<ClassItem>;
}

export const AdminStatisticsDashboard: React.FC<
  AdminStatisticsDashboardProps
> = ({
  classes,
  students,
  users,
  studentGrades,
  subjectTemplates,
  getClassInfo,
}) => {
  const [selectedStatClass, setSelectedStatClass] = useState<string>(
    classes[0]?.name || ''
  );

  // Ensure selectedStatClass stays valid when classes change
  const activeClassForSubjectStats = useMemo(() => {
    if (
      selectedStatClass &&
      classes.some(
        (c) =>
          c.name.trim().toLowerCase() === selectedStatClass.trim().toLowerCase()
      )
    ) {
      return selectedStatClass;
    }
    return classes[0]?.name || '';
  }, [classes, selectedStatClass]);

  // 1. Teacher Statistics
  const teacherStats = useMemo(() => {
    const teachers = users.filter(
      (u) => u.roles.includes('gurumapel') || u.roles.includes('walikelas')
    );
    const waliCount = teachers.filter((u) =>
      u.roles.includes('walikelas')
    ).length;
    const guruMapelCount = teachers.filter((u) =>
      u.roles.includes('gurumapel')
    ).length;
    const dualRoleCount = teachers.filter(
      (u) => u.roles.includes('walikelas') && u.roles.includes('gurumapel')
    ).length;

    let withNip = 0;
    let withNuptk = 0;
    let withoutId = 0;
    teachers.forEach((t) => {
      const parsed = formatNipOrNuptk(t.nip);
      if (parsed.type === 'NIP') withNip++;
      else if (parsed.type === 'NUPTK') withNuptk++;
      else withoutId++;
    });

    return {
      totalTeachers: teachers.length,
      waliCount,
      guruMapelCount,
      dualRoleCount,
      withNip,
      withNuptk,
      withoutId,
    };
  }, [users]);

  // 2. Class & Per-Class Student + Grade Statistics
  const classDetailsStats = useMemo(() => {
    let globalExpectedCells = 0;
    let globalFilledCells = 0;
    let globalScoreSum = 0;

    const perClass = classes.map((cls) => {
      const resolvedInfo = getClassInfo(cls.name);
      const waliName =
        resolvedInfo.waliKelasName &&
        resolvedInfo.waliKelasName.trim() !== '-' &&
        resolvedInfo.waliKelasName.trim() !== 'Nama Wali Kelas'
          ? formatTeacherNameWithTitle(resolvedInfo.waliKelasName)
          : 'Belum Ditentukan';
      const waliIdInfo = formatNipOrNuptk(resolvedInfo.waliKelasNip);

      const classStudents = students.filter(
        (s) => s.className.trim().toLowerCase() === cls.name.trim().toLowerCase()
      );
      const tpl = subjectTemplates[cls.grade] || subjectTemplates['10'];
      const subjectsList = [
        ...tpl.umum.map((s) => ({ name: s.name, category: 'Umum' })),
        ...tpl.kejuruan.map((s) => ({ name: s.name, category: 'Kejuruan' })),
      ];

      const expectedCells = classStudents.length * subjectsList.length;
      let filledCells = 0;
      let scoreSum = 0;
      let completedSubjectsCount = 0;

      subjectsList.forEach((sub) => {
        let subFilled = 0;
        classStudents.forEach((st) => {
          const val = studentGrades[st.id]?.[sub.name];
          if (typeof val === 'number' && !Number.isNaN(val)) {
            filledCells++;
            subFilled++;
            scoreSum += val;
          }
        });
        if (classStudents.length > 0 && subFilled === classStudents.length) {
          completedSubjectsCount++;
        }
      });

      globalExpectedCells += expectedCells;
      globalFilledCells += filledCells;
      globalScoreSum += scoreSum;

      const progressPct =
        expectedCells > 0
          ? Math.round((filledCells / expectedCells) * 100)
          : 0;
      const avgScore =
        filledCells > 0 ? Number((scoreSum / filledCells).toFixed(2)) : 0;

      return {
        ...cls,
        waliDisplayName: waliName,
        waliIdFormatted: waliIdInfo.formatted,
        waliIdType: waliIdInfo.type,
        studentCount: classStudents.length,
        subjectCount: subjectsList.length,
        completedSubjectsCount,
        expectedCells,
        filledCells,
        progressPct,
        avgScore,
      };
    });

    const globalProgressPct =
      globalExpectedCells > 0
        ? Math.round((globalFilledCells / globalExpectedCells) * 100)
        : 0;
    const globalAvgScore =
      globalFilledCells > 0
        ? Number((globalScoreSum / globalFilledCells).toFixed(2))
        : 0;

    return {
      perClass,
      globalExpectedCells,
      globalFilledCells,
      globalProgressPct,
      globalAvgScore,
    };
  }, [classes, students, studentGrades, subjectTemplates, getClassInfo]);

  // 3. Student Distribution by Grade Level (Kelas 10, 11, 12)
  const studentsByGrade = useMemo(() => {
    const counts: Record<string, number> = { '10': 0, '11': 0, '12': 0 };
    students.forEach((s) => {
      const g = s.grade || '10';
      counts[g] = (counts[g] || 0) + 1;
    });
    return counts;
  }, [students]);

  // 4. Subject-Level Grade Input Statistics for Selected Class
  const subjectStatsForSelectedClass = useMemo(() => {
    const targetCls =
      classes.find(
        (c) =>
          c.name.trim().toLowerCase() ===
          activeClassForSubjectStats.trim().toLowerCase()
      ) || classes[0];

    if (!targetCls) return [];

    const classStudents = students.filter(
      (s) =>
        s.className.trim().toLowerCase() ===
        targetCls.name.trim().toLowerCase()
    );
    const tpl = subjectTemplates[targetCls.grade] || subjectTemplates['10'];
    const subjectsList = [
      ...tpl.umum.map((s) => ({ name: s.name, category: 'Umum' })),
      ...tpl.kejuruan.map((s) => ({ name: s.name, category: 'Kejuruan' })),
    ];

    return subjectsList.map((sub) => {
      const scores: number[] = [];
      classStudents.forEach((st) => {
        const val = studentGrades[st.id]?.[sub.name];
        if (typeof val === 'number' && !Number.isNaN(val)) {
          scores.push(val);
        }
      });

      const filled = scores.length;
      const total = classStudents.length;
      const pct = total > 0 ? Math.round((filled / total) * 100) : 0;
      const avg =
        filled > 0
          ? Number((scores.reduce((a, b) => a + b, 0) / filled).toFixed(1))
          : 0;
      const max = filled > 0 ? Math.max(...scores) : 0;
      const min = filled > 0 ? Math.min(...scores) : 0;

      return {
        name: sub.name,
        category: sub.category,
        filled,
        total,
        pct,
        avg,
        max,
        min,
      };
    });
  }, [
    classes,
    activeClassForSubjectStats,
    students,
    studentGrades,
    subjectTemplates,
  ]);

  const maxStudentsInSingleClass = useMemo(() => {
    return Math.max(
      1,
      ...classDetailsStats.perClass.map((c) => c.studentCount)
    );
  }, [classDetailsStats.perClass]);

  return (
    <div className="space-y-6">
      {/* TOP KPI SUMMARY CARDS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Card 1: Jumlah Siswa Global */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
              Jumlah Siswa Global
            </span>
            <div className="w-9 h-9 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center">
              <i className="fas fa-user-graduate"></i>
            </div>
          </div>
          <div className="text-3xl font-bold font-mono text-slate-900">
            {students.length}
            <span className="text-sm font-normal text-slate-500 ml-1.5">
              Siswa
            </span>
          </div>
          <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-xs text-slate-600 font-mono">
            <span>Kls 10: {studentsByGrade['10'] || 0}</span>
            <span>·</span>
            <span>Kls 11: {studentsByGrade['11'] || 0}</span>
            <span>·</span>
            <span>Kls 12: {studentsByGrade['12'] || 0}</span>
          </div>
        </div>

        {/* Card 2: Jumlah Kelas & Wali Kelas */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
              Jumlah Rombel / Kelas
            </span>
            <div className="w-9 h-9 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center">
              <i className="fas fa-door-open"></i>
            </div>
          </div>
          <div className="text-3xl font-bold font-mono text-slate-900">
            {classes.length}
            <span className="text-sm font-normal text-slate-500 ml-1.5">
              Rombel
            </span>
          </div>
          <div className="mt-3 pt-2.5 border-t border-slate-100 text-xs text-slate-600">
            Wali Kelas Terisi:{' '}
            <b className="font-mono text-emerald-700">
              {
                classDetailsStats.perClass.filter(
                  (c) => c.waliDisplayName !== 'Belum Ditentukan'
                ).length
              }{' '}
              / {classes.length}
            </b>{' '}
            Kelas
          </div>
        </div>

        {/* Card 3: Jumlah Guru */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
              Jumlah Guru Terdaftar
            </span>
            <div className="w-9 h-9 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center">
              <i className="fas fa-chalkboard-teacher"></i>
            </div>
          </div>
          <div className="text-3xl font-bold font-mono text-slate-900">
            {teacherStats.totalTeachers}
            <span className="text-sm font-normal text-slate-500 ml-1.5">
              Guru
            </span>
          </div>
          <div className="mt-3 pt-2.5 border-t border-slate-100 flex items-center justify-between text-xs text-slate-600">
            <span>
              Wali: <b className="font-mono">{teacherStats.waliCount}</b>
            </span>
            <span>·</span>
            <span>
              Guru Mapel:{' '}
              <b className="font-mono">{teacherStats.guruMapelCount}</b>
            </span>
            <span>·</span>
            <span>
              NIP/NUPTK:{' '}
              <b className="font-mono">
                {teacherStats.withNip + teacherStats.withNuptk}
              </b>
            </span>
          </div>
        </div>

        {/* Card 4: Statistik Input Nilai STS */}
        <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
              Progres Input Nilai STS
            </span>
            <div className="w-9 h-9 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center">
              <i className="fas fa-chart-pie"></i>
            </div>
          </div>
          <div className="flex items-baseline justify-between">
            <div className="text-3xl font-bold font-mono text-indigo-600">
              {classDetailsStats.globalProgressPct}%
            </div>
            <div className="text-xs text-slate-500 font-mono">
              Rata-rata: <b>{classDetailsStats.globalAvgScore}</b>
            </div>
          </div>
          <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden mt-2">
            <div
              className="h-full bg-indigo-600 transition-all duration-300"
              style={{ width: `${classDetailsStats.globalProgressPct}%` }}
            />
          </div>
          <div className="mt-2 text-[11px] text-slate-500 font-mono">
            {classDetailsStats.globalFilledCells} dari{' '}
            {classDetailsStats.globalExpectedCells} data nilai terisi
          </div>
        </div>
      </div>

      {/* SECTION 2: REKAPITULASI KELAS, WALI KELAS, JUMLAH SISWA PER KELAS & PROGRES NILAI */}
      <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
        <div className="flex flex-wrap items-center justify-between border-b border-slate-200 pb-3 mb-4 gap-2">
          <div>
            <h5 className="font-bold text-base text-slate-900">
              <i className="fas fa-school text-blue-600 mr-2"></i>
              Statistik Kelas, Wali Kelas, Jumlah Siswa Per Kelas & Ketuntasan
              Nilai
            </h5>
            <p className="text-xs text-slate-500 mt-0.5">
              Daftar lengkap seluruh kelas beserta Wali Kelas, jumlah siswa di
              setiap kelas, dan persentase nilai yang sudah diinput.
            </p>
          </div>
          <div className="text-xs font-mono text-slate-600">
            Total Global: <b>{classes.length} Kelas</b> ·{' '}
            <b>{students.length} Siswa</b>
          </div>
        </div>

        {classDetailsStats.perClass.length === 0 ? (
          <div className="text-center py-8 text-slate-400 text-sm">
            Belum ada data kelas terdaftar.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm text-left border-collapse">
              <thead className="bg-slate-50 text-slate-500 text-xs uppercase border-b border-slate-200">
                <tr>
                  <th className="py-3 px-4">Rombel / Kelas</th>
                  <th className="py-3 px-4">Wali Kelas & NIP/NUPTK</th>
                  <th className="py-3 px-4">Jumlah Siswa Per Kelas</th>
                  <th className="py-3 px-4">Mapel Tuntas</th>
                  <th className="py-3 px-4 w-56">Progres Input Nilai</th>
                  <th className="py-3 px-4 text-center">Rata-Rata Kelas</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {classDetailsStats.perClass.map((c) => {
                  const studentBarPct = Math.round(
                    (c.studentCount / maxStudentsInSingleClass) * 100
                  );
                  return (
                    <tr key={c.id} className="hover:bg-slate-50">
                      <td className="py-3.5 px-4">
                        <div className="font-bold text-slate-900">{c.name}</div>
                        <div className="text-xs text-slate-500">
                          Kelas {c.grade} · {c.major}
                        </div>
                      </td>
                      <td className="py-3.5 px-4">
                        <div className="font-bold text-slate-800">
                          {c.waliDisplayName}
                        </div>
                        {c.waliIdFormatted ? (
                          <div className="text-xs font-mono text-slate-500">
                            {c.waliIdFormatted}
                          </div>
                        ) : (
                          <div className="text-[11px] text-slate-400 italic">
                            Tanpa NIP/NUPTK (Hanya Nama)
                          </div>
                        )}
                      </td>
                      <td className="py-3.5 px-4">
                        <div className="flex items-center justify-between text-xs mb-1">
                          <span className="font-bold font-mono text-slate-900">
                            {c.studentCount} Siswa
                          </span>
                          <span className="text-slate-400 font-mono">
                            {students.length > 0
                              ? Math.round(
                                  (c.studentCount / students.length) * 100
                                )
                              : 0}
                            % dari total
                          </span>
                        </div>
                        <div className="w-36 h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-blue-600 rounded-full"
                            style={{ width: `${studentBarPct}%` }}
                          />
                        </div>
                      </td>
                      <td className="py-3.5 px-4 font-mono text-xs text-slate-700">
                        <b>{c.completedSubjectsCount}</b> / {c.subjectCount}{' '}
                        Mapel
                      </td>
                      <td className="py-3.5 px-4">
                        <div className="flex items-center justify-between text-xs mb-1 font-mono">
                          <span
                            className={`font-bold ${
                              c.progressPct === 100
                                ? 'text-emerald-700'
                                : c.progressPct > 0
                                ? 'text-amber-700'
                                : 'text-slate-500'
                            }`}
                          >
                            {c.progressPct}%
                          </span>
                          <span className="text-slate-400">
                            ({c.filledCells}/{c.expectedCells})
                          </span>
                        </div>
                        <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className={`h-full transition-all duration-300 ${
                              c.progressPct === 100
                                ? 'bg-emerald-600'
                                : c.progressPct > 0
                                ? 'bg-amber-500'
                                : 'bg-slate-300'
                            }`}
                            style={{ width: `${c.progressPct}%` }}
                          />
                        </div>
                      </td>
                      <td className="py-3.5 px-4 text-center font-mono font-bold text-blue-600">
                        {c.avgScore}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* SECTION 3: VISUAL STATISTIK NILAI PER MATA PELAJARAN */}
      <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm">
        <div className="flex flex-wrap items-center justify-between border-b border-slate-200 pb-4 mb-5 gap-3">
          <div>
            <h5 className="font-bold text-base text-slate-900">
              <i className="fas fa-chart-bar text-indigo-600 mr-2"></i>
              Visual Statistik Nilai yang Sudah Diinput per Mata Pelajaran
            </h5>
            <p className="text-xs text-slate-500 mt-0.5">
              Pantau progres keterisian nilai STS, nilai rata-rata, nilai
              tertinggi, dan terendah setiap mata pelajaran.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <label className="text-xs font-bold text-slate-600">
              Pilih Rombel:
            </label>
            <select
              value={activeClassForSubjectStats}
              onChange={(e) => setSelectedStatClass(e.target.value)}
              className="px-3 py-1.5 border border-slate-300 rounded-lg text-xs font-bold outline-none focus:border-blue-600"
            >
              {classes.map((c) => (
                <option key={c.id} value={c.name}>
                  {c.name} ({students.filter((s) => s.className === c.name).length}{' '}
                  Siswa)
                </option>
              ))}
            </select>
          </div>
        </div>

        {subjectStatsForSelectedClass.length === 0 ? (
          <div className="text-center py-8 text-slate-400 text-sm">
            Pilih kelas untuk melihat statistik mata pelajaran.
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {subjectStatsForSelectedClass.map((sub, idx) => (
              <div
                key={idx}
                className="p-4 rounded-xl border border-slate-200 bg-slate-50/60 flex flex-col justify-between gap-2.5"
              >
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <span className="text-[11px] font-semibold text-slate-500">
                      [{sub.category}]
                    </span>
                    <div className="font-bold text-xs text-slate-900 leading-snug">
                      {sub.name}
                    </div>
                  </div>
                  <span
                    className={`text-xs font-mono font-bold shrink-0 ${
                      sub.pct === 100
                        ? 'text-emerald-700'
                        : sub.pct > 0
                        ? 'text-amber-700'
                        : 'text-red-600'
                    }`}
                  >
                    {sub.pct}% ({sub.filled}/{sub.total} Siswa)
                  </span>
                </div>

                <div className="w-full h-2 bg-slate-200 rounded-full overflow-hidden">
                  <div
                    className={`h-full transition-all duration-300 ${
                      sub.pct === 100
                        ? 'bg-emerald-600'
                        : sub.pct > 0
                        ? 'bg-amber-500'
                        : 'bg-red-400'
                    }`}
                    style={{ width: `${sub.pct}%` }}
                  />
                </div>

                <div className="flex items-center justify-between text-[11px] text-slate-600 font-mono pt-1 border-t border-slate-200/80">
                  <span>
                    Rata-rata: <b className="text-slate-900">{sub.avg}</b>
                  </span>
                  <span>·</span>
                  <span>
                    Tertinggi:{' '}
                    <b className="text-emerald-700">{sub.max || '-'}</b>
                  </span>
                  <span>·</span>
                  <span>
                    Terendah: <b className="text-amber-700">{sub.min || '-'}</b>
                  </span>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
