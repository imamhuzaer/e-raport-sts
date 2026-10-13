import React, { useMemo, useState } from 'react';
import {
  ClassItem,
  RankedStudentItem,
  SchoolData,
  StudentGradesMap,
} from '../types';
import {
  exportClassLegerToCsv,
  exportClassLegerToExcel,
  formatNipOrNuptk,
  formatTeacherNameWithTitle,
  getSubjectAbbreviation,
  triggerPrintClassLeger,
} from '../services/templateAndPdfService';

interface LegerTableModalProps {
  isOpen: boolean;
  onClose: () => void;
  school: SchoolData;
  classInfo: ClassItem;
  students: RankedStudentItem[];
  subjects: Array<{ name: string; category: string }>;
  grades: StudentGradesMap;
}

export const LegerTableModal: React.FC<LegerTableModalProps> = ({
  isOpen,
  onClose,
  school,
  classInfo,
  students,
  subjects,
  grades,
}) => {
  const [searchTerm, setSearchTerm] = useState('');

  const filteredStudents = useMemo(() => {
    if (!searchTerm.trim()) return students;
    const q = searchTerm.toLowerCase();
    return students.filter(
      (s) =>
        s.name.toLowerCase().includes(q) ||
        s.nisn.toLowerCase().includes(q) ||
        s.nis.toLowerCase().includes(q)
    );
  }, [students, searchTerm]);

  // Calculations for subject statistics across all students in class
  const stats = useMemo(() => {
    const subjectStats: Record<
      string,
      { avg: number; max: number; min: number; count: number }
    > = {};

    subjects.forEach((sub) => {
      let sum = 0;
      let count = 0;
      let max = 0;
      let min = 100;

      students.forEach((st) => {
        const val = grades[st.id]?.[sub.name];
        if (val !== undefined && val !== null && !isNaN(Number(val)) && Number(val) > 0) {
          const num = Number(val);
          sum += num;
          count++;
          if (num > max) max = num;
          if (num < min) min = num;
        }
      });

      subjectStats[sub.name] = {
        avg: count > 0 ? Number((sum / count).toFixed(1)) : 0,
        max,
        min: count > 0 ? min : 0,
        count,
      };
    });

    const totalStudents = students.length;
    const grandTotalAvg =
      totalStudents > 0
        ? Number(
            (
              students.reduce((acc, s) => acc + s.totalScore, 0) /
              totalStudents
            ).toFixed(1)
          )
        : 0;

    const grandClassAvg =
      totalStudents > 0
        ? Number(
            (
              students.reduce((acc, s) => acc + s.avgScore, 0) /
              totalStudents
            ).toFixed(1)
          )
        : 0;

    return {
      subjectStats,
      grandTotalAvg,
      grandClassAvg,
    };
  }, [students, subjects, grades]);

  const handleExportExcel = () => {
    exportClassLegerToExcel({
      school,
      classInfo,
      students,
      allSubjects: subjects,
      grades,
    });
  };

  const handleExportCsv = () => {
    exportClassLegerToCsv({
      school,
      classInfo,
      students,
      allSubjects: subjects,
      grades,
    });
  };

  const handlePrint = () => {
    triggerPrintClassLeger();
  };

  const formattedWaliName = formatTeacherNameWithTitle(classInfo.waliKelasName);
  const formattedWaliNip = formatNipOrNuptk(classInfo.waliKelasNip);
  const formattedPrincipalName = formatTeacherNameWithTitle(school.principal);
  const formattedPrincipalNip = formatNipOrNuptk(school.principalNip);

  // Render the core Leger Table element reusable for Modal preview and Print layout
  const renderLegerContent = (isForPrint = false) => (
    <div
      className={
        isForPrint
          ? 'w-full text-black'
          : 'bg-white p-6 rounded-xl border border-slate-200'
      }
      style={{ fontFamily: "'Times New Roman', Times, serif" }}
    >
      {/* Official Leger Header */}
      <div className="text-center border-b-2 border-black pb-3 mb-4">
        <div className="text-xs uppercase tracking-wider font-bold">
          Pemerintah Provinsi Lampung · Dinas Pendidikan dan Kebudayaan
        </div>
        <div className="text-base font-bold uppercase tracking-wide">
          {school.name}
        </div>
        <div className="text-lg font-extrabold uppercase mt-1">
          Buku Leger Nilai Hasil Belajar Peserta Didik
        </div>
        <div className="text-xs italic text-slate-700">
          Penilaian Sumatif Tengah Semester (STS) · Kurikulum Merdeka
        </div>

        {/* Metadata Grid */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2 text-xs text-left mt-3 pt-2 border-t border-slate-300 font-sans">
          <div>
            <span className="text-slate-500">Kelas / Rombel:</span>{' '}
            <b className="text-slate-900">{classInfo.name}</b>
          </div>
          <div>
            <span className="text-slate-500">Tingkat / Fase:</span>{' '}
            <b className="text-slate-900">Kelas {classInfo.grade}</b>
          </div>
          <div>
            <span className="text-slate-500">Konsentrasi Keahlian:</span>{' '}
            <b className="text-slate-900">{classInfo.major}</b>
          </div>
          <div>
            <span className="text-slate-500">Tahun Pelajaran:</span>{' '}
            <b className="text-slate-900">{school.tp} (Ganjil)</b>
          </div>
          <div className="col-span-2">
            <span className="text-slate-500">Wali Kelas:</span>{' '}
            <b className="text-slate-900">{formattedWaliName}</b>
            {formattedWaliNip.formatted && (
              <span className="text-slate-600 font-mono ml-1">
                ({formattedWaliNip.formatted})
              </span>
            )}
          </div>
          <div className="col-span-2 text-right">
            <span className="text-slate-500">Jumlah Siswa:</span>{' '}
            <b className="text-slate-900">{students.length} Siswa</b>
            <span className="text-slate-400 mx-1">|</span>
            <span className="text-slate-500">Mata Pelajaran:</span>{' '}
            <b className="text-slate-900">{subjects.length} Mapel</b>
          </div>
        </div>
      </div>

      {/* Matrix Table */}
      <div className={isForPrint ? 'w-full' : 'overflow-x-auto'}>
        <table
          className="w-full text-left border-collapse"
          style={{
            fontSize: isForPrint ? '7.5pt' : '8.5pt',
            lineHeight: 1.15,
          }}
        >
          <thead>
            <tr className="bg-slate-100 text-slate-900 border border-black">
              <th
                rowSpan={2}
                className="border border-black p-1 text-center font-bold w-7"
              >
                No
              </th>
              <th
                rowSpan={2}
                className="border border-black p-1 text-center font-bold w-20"
              >
                NISN
              </th>
              <th
                rowSpan={2}
                className="border border-black p-1 text-center font-bold w-16"
              >
                NIS
              </th>
              <th
                rowSpan={2}
                className="border border-black p-1.5 font-bold min-w-[150px]"
              >
                Nama Peserta Didik
              </th>
              <th
                colSpan={subjects.length}
                className="border border-black p-1 text-center font-bold bg-slate-200"
              >
                Nilai Sumatif Mata Pelajaran (Skala 0 - 100)
              </th>
              <th
                rowSpan={2}
                className="border border-black p-1 text-center font-bold w-12 bg-amber-50"
              >
                Jml
              </th>
              <th
                rowSpan={2}
                className="border border-black p-1 text-center font-bold w-12 bg-blue-50"
              >
                Rata
              </th>
              <th
                rowSpan={2}
                className="border border-black p-1 text-center font-bold w-9 bg-emerald-50"
              >
                Rank
              </th>
              <th
                colSpan={3}
                className="border border-black p-1 text-center font-bold bg-slate-100"
              >
                Absensi
              </th>
              <th
                rowSpan={2}
                className="border border-black p-1 font-bold min-w-[140px]"
              >
                Catatan Wali Kelas
              </th>
            </tr>
            <tr className="bg-slate-50 text-slate-900 border border-black">
              {subjects.map((sub) => {
                const abbr = getSubjectAbbreviation(sub.name);
                return (
                  <th
                    key={sub.name}
                    title={`${sub.name} (${sub.category})`}
                    className="border border-black p-1 text-center font-bold whitespace-nowrap min-w-[34px]"
                  >
                    <div className="font-mono">{abbr}</div>
                    <div className="text-[6.5pt] font-sans text-slate-500 uppercase">
                      {sub.category === 'Kejuruan' ? 'K' : 'U'}
                    </div>
                  </th>
                );
              })}
              <th className="border border-black p-0.5 text-center font-bold w-6 text-[7pt]">
                S
              </th>
              <th className="border border-black p-0.5 text-center font-bold w-6 text-[7pt]">
                I
              </th>
              <th className="border border-black p-0.5 text-center font-bold w-6 text-[7pt]">
                A
              </th>
            </tr>
          </thead>
          <tbody>
            {filteredStudents.length === 0 ? (
              <tr>
                <td
                  colSpan={subjects.length + 9}
                  className="border border-black p-4 text-center text-slate-500 italic"
                >
                  Tidak ada data siswa ditemukan.
                </td>
              </tr>
            ) : (
              filteredStudents.map((st, idx) => {
                const sGrades = grades[st.id] || {};
                const s = st.attendance?.sakit || 0;
                const i = st.attendance?.izin || 0;
                const a = st.attendance?.alpa || 0;

                return (
                  <tr
                    key={st.id}
                    className={
                      st.rank <= 3
                        ? 'bg-amber-50/40 hover:bg-amber-50'
                        : 'hover:bg-slate-50'
                    }
                  >
                    <td className="border border-black p-1 text-center font-mono font-bold">
                      {idx + 1}
                    </td>
                    <td className="border border-black p-1 text-center font-mono text-[8pt]">
                      {st.nisn}
                    </td>
                    <td className="border border-black p-1 text-center font-mono text-[8pt]">
                      {st.nis}
                    </td>
                    <td className="border border-black p-1 font-semibold whitespace-nowrap">
                      {st.name}
                    </td>

                    {/* Subject Scores */}
                    {subjects.map((sub) => {
                      const val = sGrades[sub.name];
                      const hasVal =
                        val !== undefined &&
                        val !== null &&
                        !isNaN(Number(val));
                      const numVal = hasVal ? Number(val) : null;
                      return (
                        <td
                          key={sub.name}
                          className={`border border-black p-1 text-center font-mono ${
                            numVal !== null && numVal >= 88
                              ? 'font-bold text-emerald-800'
                              : numVal !== null && numVal < 75
                              ? 'text-rose-700 font-bold'
                              : ''
                          }`}
                        >
                          {numVal !== null ? numVal : '-'}
                        </td>
                      );
                    })}

                    {/* Total & Average */}
                    <td className="border border-black p-1 text-center font-mono font-bold bg-amber-50/30">
                      {st.totalScore}
                    </td>
                    <td className="border border-black p-1 text-center font-mono font-bold bg-blue-50/30 text-blue-900">
                      {st.avgScore}
                    </td>

                    {/* Rank */}
                    <td className="border border-black p-1 text-center font-mono font-bold bg-emerald-50/30">
                      {st.rank === 1 ? (
                        <span className="font-extrabold text-amber-700">
                          #1
                        </span>
                      ) : st.rank === 2 ? (
                        <span className="font-extrabold text-slate-700">
                          #2
                        </span>
                      ) : st.rank === 3 ? (
                        <span className="font-extrabold text-orange-700">
                          #3
                        </span>
                      ) : (
                        st.rank
                      )}
                    </td>

                    {/* Attendance */}
                    <td className="border border-black p-0.5 text-center font-mono text-[7.5pt]">
                      {s > 0 ? s : '-'}
                    </td>
                    <td className="border border-black p-0.5 text-center font-mono text-[7.5pt]">
                      {i > 0 ? i : '-'}
                    </td>
                    <td className="border border-black p-0.5 text-center font-mono text-[7.5pt]">
                      {a > 0 ? (
                        <span className="font-bold text-rose-700">{a}</span>
                      ) : (
                        '-'
                      )}
                    </td>

                    {/* Catatan */}
                    <td className="border border-black p-1 text-[7.5pt] leading-tight italic text-slate-700 max-w-[200px] truncate">
                      {st.note || '-'}
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>

          {/* Footer Statistics */}
          <tfoot>
            <tr className="bg-slate-100 font-bold border border-black">
              <td
                colSpan={4}
                className="border border-black p-1 text-right uppercase tracking-wider text-[8pt]"
              >
                Rata-Rata Kelas
              </td>
              {subjects.map((sub) => (
                <td
                  key={sub.name}
                  className="border border-black p-1 text-center font-mono text-[8pt] text-blue-900"
                >
                  {stats.subjectStats[sub.name]?.avg || '-'}
                </td>
              ))}
              <td className="border border-black p-1 text-center font-mono text-[8pt]">
                {stats.grandTotalAvg}
              </td>
              <td className="border border-black p-1 text-center font-mono text-[8pt] text-blue-900 font-extrabold">
                {stats.grandClassAvg}
              </td>
              <td colSpan={5} className="border border-black p-1"></td>
            </tr>

            <tr className="bg-slate-50 font-bold border border-black">
              <td
                colSpan={4}
                className="border border-black p-1 text-right uppercase tracking-wider text-[7.5pt] text-emerald-800"
              >
                Nilai Tertinggi
              </td>
              {subjects.map((sub) => (
                <td
                  key={sub.name}
                  className="border border-black p-1 text-center font-mono text-[7.5pt] text-emerald-800"
                >
                  {stats.subjectStats[sub.name]?.max || '-'}
                </td>
              ))}
              <td
                colSpan={7}
                className="border border-black p-1 text-right text-[7pt] text-slate-500"
              ></td>
            </tr>

            <tr className="bg-slate-50 font-bold border border-black">
              <td
                colSpan={4}
                className="border border-black p-1 text-right uppercase tracking-wider text-[7.5pt] text-rose-800"
              >
                Nilai Terendah
              </td>
              {subjects.map((sub) => (
                <td
                  key={sub.name}
                  className="border border-black p-1 text-center font-mono text-[7.5pt] text-rose-800"
                >
                  {stats.subjectStats[sub.name]?.min || '-'}
                </td>
              ))}
              <td
                colSpan={7}
                className="border border-black p-1 text-right text-[7pt] text-slate-500"
              ></td>
            </tr>
          </tfoot>
        </table>
      </div>

      {/* Legend & Abbreviation footnote */}
      <div className="mt-3 pt-2 border-t border-slate-300 font-sans text-[7.5pt] text-slate-600">
        <div className="font-bold text-slate-800 uppercase mb-1">
          Keterangan Singkatan Mata Pelajaran:
        </div>
        <div className="flex flex-wrap gap-x-3 gap-y-1">
          {subjects.map((s) => (
            <span key={s.name} className="inline-flex items-center gap-1">
              <b className="font-mono text-slate-900">
                [{getSubjectAbbreviation(s.name)}]
              </b>
              <span>{s.name}</span>
            </span>
          ))}
          <span className="inline-flex items-center gap-1 ml-2 text-slate-500">
            • <b>U</b> = Muatan Umum, <b>K</b> = Muatan Kejuruan
          </span>
        </div>
      </div>

      {/* Official Signatures for Leger Document */}
      <div className="mt-6 pt-4 border-t border-slate-300 flex justify-between items-start text-xs">
        {/* Kepala Sekolah */}
        <div className="text-center w-72">
          <p className="mb-0">Mengetahui,</p>
          <p className="mb-1">Kepala {school.name}</p>
          {school.principalSignatureType &&
          school.principalSignatureType !== 'none' &&
          school.principalSignatureUrl ? (
            <div className="h-14 flex items-center justify-center my-1">
              <img
                src={school.principalSignatureUrl}
                alt="TTD KS"
                className={
                  school.principalSignatureType === 'qrcode'
                    ? 'w-12 h-12 object-contain mx-auto'
                    : 'max-h-12 max-w-[160px] object-contain mx-auto'
                }
              />
            </div>
          ) : (
            <div className="h-12 flex items-center justify-center text-[10pt] text-slate-400 italic">
              (Ruang Tanda Tangan)
            </div>
          )}
          <p className="font-bold underline mb-0">{formattedPrincipalName}</p>
          {formattedPrincipalNip.formatted && (
            <p className="mb-0 text-[8.5pt]">
              {formattedPrincipalNip.formatted}
            </p>
          )}
        </div>

        {/* Wali Kelas */}
        <div className="text-center w-72">
          <p className="mb-0">{school.datePrint}</p>
          <p className="mb-1">Wali Kelas {classInfo.name}</p>
          {classInfo.waliSignatureType &&
          classInfo.waliSignatureType !== 'none' &&
          classInfo.waliSignatureUrl ? (
            <div className="h-14 flex items-center justify-center my-1">
              <img
                src={classInfo.waliSignatureUrl}
                alt="TTD Wali Kelas"
                className={
                  classInfo.waliSignatureType === 'qrcode'
                    ? 'w-12 h-12 object-contain mx-auto'
                    : 'max-h-12 max-w-[160px] object-contain mx-auto'
                }
              />
            </div>
          ) : (
            <div className="h-12 flex items-center justify-center text-[10pt] text-slate-400 italic">
              (Ruang Tanda Tangan)
            </div>
          )}
          <p className="font-bold underline mb-0">{formattedWaliName}</p>
          {formattedWaliNip.formatted && (
            <p className="mb-0 text-[8.5pt]">{formattedWaliNip.formatted}</p>
          )}
        </div>
      </div>

      {/* Mandatory Copyright Footer */}
      <div className="mt-4 pt-2 border-t border-slate-200 text-center font-sans text-[10px] text-slate-500">
        © 2026 E-Raport STS SMK Negeri 1 Mesuji Timur · Desain & Sistem Dibuat
        oleh{' '}
        <span className="font-bold text-slate-700">Imam Huzaer, S.Pd.</span>
      </div>
    </div>
  );

  return (
    <>
      {/* 1. INTERACTIVE MODAL (VIEW / FILTER / ACTIONS) */}
      {isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/80 p-3 md:p-6 no-print overflow-hidden">
          <div className="bg-white rounded-2xl shadow-2xl max-w-7xl w-full h-[94vh] flex flex-col overflow-hidden border border-slate-200">
            {/* Modal Top Bar */}
            <div className="bg-slate-900 text-white px-5 py-3.5 flex flex-wrap items-center justify-between gap-3 shrink-0">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center text-white text-lg">
                  <i className="fas fa-table"></i>
                </div>
                <div>
                  <h5 className="font-bold text-base text-white mb-0">
                    Buku Leger Nilai Siswa - Kelas {classInfo.name}
                  </h5>
                  <div className="text-xs text-slate-400">
                    {classInfo.major} · TP {school.tp} · Wali Kelas:{' '}
                    <span className="text-slate-200 font-semibold">
                      {formattedWaliName}
                    </span>
                  </div>
                </div>
              </div>

              {/* Action Buttons Toolbar */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={handlePrint}
                  className="px-3.5 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer flex items-center gap-1.5 shadow-sm"
                  title="Cetak format cetak landscape A4 siap print / save as PDF"
                >
                  <i className="fas fa-print"></i>
                  <span>Cetak Leger (Landscape / PDF)</span>
                </button>

                <button
                  type="button"
                  onClick={handleExportExcel}
                  className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer flex items-center gap-1.5 shadow-sm"
                  title="Unduh file Excel Leger lengkap dengan rumus statistik"
                >
                  <i className="fas fa-file-excel"></i>
                  <span>Export Excel (.xlsx)</span>
                </button>

                <button
                  type="button"
                  onClick={handleExportCsv}
                  className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 font-bold text-xs rounded-lg transition-colors cursor-pointer flex items-center gap-1.5 border border-slate-700"
                  title="Unduh file CSV"
                >
                  <i className="fas fa-file-csv"></i>
                  <span>CSV</span>
                </button>

                <button
                  type="button"
                  onClick={onClose}
                  className="px-3 py-2 bg-slate-800 hover:bg-red-600 text-white font-bold text-xs rounded-lg transition-colors cursor-pointer"
                  title="Tutup Modal"
                >
                  <i className="fas fa-times"></i>
                </button>
              </div>
            </div>

            {/* Filter & Quick Info Sub-Header */}
            <div className="bg-slate-100 px-5 py-2.5 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 shrink-0">
              <div className="flex items-center gap-2 w-full sm:w-auto">
                <span className="text-xs font-bold text-slate-600">
                  <i className="fas fa-search mr-1 text-slate-400"></i> Cari
                  Siswa:
                </span>
                <input
                  type="text"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  placeholder="Ketik nama atau NISN..."
                  className="px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-semibold outline-none focus:border-indigo-600 w-64"
                />
                {searchTerm && (
                  <button
                    type="button"
                    onClick={() => setSearchTerm('')}
                    className="text-xs text-slate-400 hover:text-slate-600 cursor-pointer font-bold"
                  >
                    Reset
                  </button>
                )}
              </div>

              <div className="flex items-center gap-3 text-xs text-slate-600">
                <span className="inline-flex items-center gap-1.5 bg-white px-2.5 py-1 rounded-md border border-slate-200 font-semibold">
                  <i className="fas fa-users text-indigo-600"></i> Total Siswa:{' '}
                  <b>{students.length}</b>
                </span>
                <span className="inline-flex items-center gap-1.5 bg-white px-2.5 py-1 rounded-md border border-slate-200 font-semibold">
                  <i className="fas fa-book text-emerald-600"></i> Mapel:{' '}
                  <b>{subjects.length}</b>
                </span>
                <span className="inline-flex items-center gap-1.5 bg-white px-2.5 py-1 rounded-md border border-slate-200 font-semibold">
                  <i className="fas fa-chart-line text-blue-600"></i> Rata-Rata
                  Kelas: <b>{stats.grandClassAvg}</b>
                </span>
              </div>
            </div>

            {/* Scrollable Preview Area */}
            <div className="flex-grow overflow-auto p-4 md:p-6 bg-slate-50">
              {renderLegerContent(false)}
            </div>
          </div>
        </div>
      )}

      {/* 2. DEDICATED PRINT VIEW CONTAINER (RENDERED DIRECTLY FOR WINDOW.PRINT()) */}
      <div className="leger-print-page">{renderLegerContent(true)}</div>
    </>
  );
};
