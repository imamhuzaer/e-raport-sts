import React, { useEffect, useState } from 'react';
import {
  ClassItem,
  GradeSubjectTemplate,
  RankedStudentItem,
  SchoolData,
} from '../types';
import {
  formatNipOrNuptk,
  formatTeacherNameWithTitle,
  resolveAndTrimKopImage,
} from '../services/templateAndPdfService';

interface RaportTemplateProps {
  student: RankedStudentItem;
  school: SchoolData;
  classInfo: Partial<ClassItem>;
  subjects: GradeSubjectTemplate;
  grades: Record<string, number>;
}

export const RaportTemplate: React.FC<RaportTemplateProps> = ({
  student,
  school,
  classInfo,
  subjects,
  grades,
}) => {
  const [imgError, setImgError] = useState(false);
  const [trimmedKopSrc, setTrimmedKopSrc] = useState<string | null>(null);

  useEffect(() => {
    let active = true;
    setImgError(false);
    if (school.kopUrl) {
      resolveAndTrimKopImage(school.kopUrl).then((dataUrl) => {
        if (active && dataUrl) {
          setTrimmedKopSrc(dataUrl);
        }
      });
    } else {
      setTrimmedKopSrc(null);
    }
    return () => {
      active = false;
    };
  }, [school.kopUrl]);

  const formattedWaliName = formatTeacherNameWithTitle(classInfo.waliKelasName);
  const formattedWaliId = formatNipOrNuptk(classInfo.waliKelasNip);
  const formattedPrincipalName = formatTeacherNameWithTitle(school.principal);
  const formattedPrincipalId = formatNipOrNuptk(school.principalNip);

  return (
    <div style={{ fontSize: '10pt', lineHeight: 1.25, color: '#000000' }}>
      {/* KOP SURAT IMAGE (100% FLUSH WITH CONTENT WIDTH) */}
      {!imgError && (trimmedKopSrc || school.kopUrl) && (
        <img
          src={trimmedKopSrc || school.kopUrl}
          alt="KOP Surat SMKN 1 Mesuji Timur"
          className="kop-img"
          style={{
            width: '100%',
            maxWidth: '100%',
            height: 'auto',
            display: 'block',
            marginBottom: '6px',
          }}
          referrerPolicy="no-referrer"
          onError={() => setImgError(true)}
        />
      )}

      {/* FALLBACK KOP SURAT (SHOWN IF IMAGE FAILS OR IN STRICT OFFLINE PDF EXPORT) */}
      <div
        className="kop-fallback-header"
        style={{
          display: !imgError && (trimmedKopSrc || school.kopUrl) ? 'none' : 'block',
          borderBottom: '3px solid #000',
          paddingBottom: '5px',
          marginBottom: '6px',
          textAlign: 'center',
          fontFamily: "'Times New Roman', Times, serif",
          width: '100%',
        }}
      >
        <div style={{ fontSize: '11.5pt', fontWeight: 'bold', lineHeight: 1.2 }}>
          PEMERINTAH PROVINSI LAMPUNG
        </div>
        <div style={{ fontSize: '11.5pt', fontWeight: 'bold', lineHeight: 1.2 }}>
          DINAS PENDIDIKAN DAN KEBUDAYAAN
        </div>
        <div
          style={{
            fontSize: '13.5pt',
            fontWeight: 'bold',
            letterSpacing: '0.4px',
            lineHeight: 1.2,
          }}
        >
          {school.name || 'SMK NEGERI 1 MESUJI TIMUR'}
        </div>
        <div style={{ fontSize: '8.5pt', fontWeight: 'bold', lineHeight: 1.2 }}>
          Konsentrasi Kejuruan Teknik Kendaraan Ringan, Manajemen Perkantoran & Teknik Sepeda Motor
        </div>
        <div style={{ fontSize: '8pt', lineHeight: 1.2 }}>
          Jl Soekarno Hatta No. 305, Tanjung Mas Rejo, Mesuji Timur, Mesuji, Lampung. NPSN : {school.npsn || '69945049'}
        </div>
        <div style={{ fontSize: '8pt', lineHeight: 1.2 }}>
          Email: <i>SMKNegeri1mesujitimur@gmail.com</i> Web: <i>SMKNegeri1mesujitimur.sch.id</i> Kode Pos 34699
        </div>
      </div>

      {/* REPORT TITLE */}
      <div
        className="text-center my-1.5"
        style={{ fontFamily: "'Times New Roman', Times, serif" }}
      >
        <h6
          className="font-bold uppercase underline mb-0"
          style={{ fontSize: '10.5pt', letterSpacing: '0.5px', lineHeight: 1.25 }}
        >
          LAPORAN HASIL BELAJAR PESERTA DIDIK
        </h6>
        <h6
          className="font-bold uppercase mb-0"
          style={{ fontSize: '10.5pt', lineHeight: 1.25 }}
        >
          SUMATIF TENGAH SEMESTER GANJIL
        </h6>
        <h6
          className="font-bold uppercase mb-1.5"
          style={{ fontSize: '10.5pt', lineHeight: 1.25 }}
        >
          TAHUN PELAJARAN {school.tp}
        </h6>
      </div>

      {/* STUDENT META TABLE */}
      <table
        className="w-full mb-1.5"
        style={{
          fontSize: '9.5pt',
          fontFamily: "'Times New Roman', Times, serif",
          borderCollapse: 'collapse',
        }}
      >
        <tbody>
          <tr>
            <td style={{ width: '150px', padding: '1px 0' }}>Nama</td>
            <td style={{ width: '10px' }}>:</td>
            <td className="font-bold uppercase" style={{ padding: '1px 0' }}>
              {student.name}
            </td>
            <td style={{ width: '80px', padding: '1px 0' }}>Kelas</td>
            <td style={{ width: '10px' }}>:</td>
            <td className="font-bold" style={{ padding: '1px 0' }}>
              {student.className}
            </td>
          </tr>
          <tr>
            <td style={{ padding: '1px 0' }}>NISN/NIS</td>
            <td>:</td>
            <td className="font-mono" style={{ padding: '1px 0' }}>
              {student.nisn} / {student.nis}
            </td>
            <td style={{ padding: '1px 0' }}>Kurikulum</td>
            <td>:</td>
            <td style={{ padding: '1px 0' }}>Kurikulum Merdeka</td>
          </tr>
          <tr>
            <td style={{ padding: '1px 0' }}>Konsentrasi Keahlian</td>
            <td>:</td>
            <td
              colSpan={4}
              className="font-semibold"
              style={{ padding: '1px 0' }}
            >
              {student.major}
            </td>
          </tr>
        </tbody>
      </table>

      {/* GRADES TABLE */}
      <table
        className="table-raport w-full mb-1.5"
        style={{
          fontSize: '9.5pt',
          fontFamily: "'Times New Roman', Times, serif",
        }}
      >
        <thead>
          <tr
            className="text-center font-bold"
            style={{ backgroundColor: '#f8f9fa' }}
          >
            <th
              style={{ width: '35px', border: '1px solid #000', padding: '2px' }}
            >
              NO
            </th>
            <th
              style={{
                border: '1px solid #000',
                padding: '2px 4px',
                textAlign: 'left',
              }}
            >
              MATA PELAJARAN
            </th>
            <th
              style={{ width: '80px', border: '1px solid #000', padding: '2px' }}
            >
              NILAI
            </th>
          </tr>
        </thead>
        <tbody>
          {/* Category A: Mata Pelajaran Umum */}
          <tr className="font-bold" style={{ backgroundColor: '#f8f9fa' }}>
            <td
              className="text-center"
              style={{ border: '1px solid #000', padding: '2px' }}
            >
              A
            </td>
            <td
              colSpan={2}
              style={{ border: '1px solid #000', padding: '2px 4px' }}
            >
              MATA PELAJARAN UMUM
            </td>
          </tr>
          {subjects.umum.map((m, idx) => (
            <tr key={`u-${idx}`}>
              <td
                className="text-center"
                style={{ border: '1px solid #000', padding: '1px 2px' }}
              >
                {idx + 1}
              </td>
              <td style={{ border: '1px solid #000', padding: '1px 4px' }}>
                {m.name}
              </td>
              <td
                className="text-center font-bold"
                style={{ border: '1px solid #000', padding: '1px 2px' }}
              >
                {grades[m.name] !== undefined && String(grades[m.name]) !== ''
                  ? grades[m.name]
                  : '-'}
              </td>
            </tr>
          ))}

          {/* Category B: Mata Pelajaran Kejuruan */}
          <tr className="font-bold" style={{ backgroundColor: '#f8f9fa' }}>
            <td
              className="text-center"
              style={{ border: '1px solid #000', padding: '2px' }}
            >
              B
            </td>
            <td
              colSpan={2}
              style={{ border: '1px solid #000', padding: '2px 4px' }}
            >
              MATA PELAJARAN KEJURUAN
            </td>
          </tr>
          {subjects.kejuruan.map((m, idx) => (
            <tr key={`k-${idx}`}>
              <td
                className="text-center"
                style={{ border: '1px solid #000', padding: '1px 2px' }}
              >
                {subjects.umum.length + idx + 1}
              </td>
              <td style={{ border: '1px solid #000', padding: '1px 4px' }}>
                {m.name}
              </td>
              <td
                className="text-center font-bold"
                style={{ border: '1px solid #000', padding: '1px 2px' }}
              >
                {grades[m.name] !== undefined && String(grades[m.name]) !== ''
                  ? grades[m.name]
                  : '-'}
              </td>
            </tr>
          ))}

          {/* Summary Rows */}
          <tr className="font-bold">
            <td
              className="text-center"
              colSpan={2}
              style={{ border: '1px solid #000', padding: '1px 4px' }}
            >
              Jumlah
            </td>
            <td
              className="text-center font-bold"
              style={{ border: '1px solid #000', padding: '1px 2px' }}
            >
              {student.totalScore || 0}
            </td>
          </tr>
          <tr className="font-bold">
            <td
              className="text-center"
              colSpan={2}
              style={{ border: '1px solid #000', padding: '1px 4px' }}
            >
              Rata-rata
            </td>
            <td
              className="text-center font-bold"
              style={{ border: '1px solid #000', padding: '1px 2px' }}
            >
              {student.avgScore || 0}
            </td>
          </tr>
          <tr className="font-bold">
            <td
              className="text-center"
              colSpan={2}
              style={{ border: '1px solid #000', padding: '1px 4px' }}
            >
              Peringkat
            </td>
            <td
              className="text-center font-bold"
              style={{ border: '1px solid #000', padding: '1px 2px' }}
            >
              {student.rank || '-'}
            </td>
          </tr>
        </tbody>
      </table>

      {/* ATTENDANCE TABLE */}
      <div className="mb-1.5">
        <table
          className="table-raport"
          style={{
            width: '200px',
            fontSize: '9pt',
            fontFamily: "'Times New Roman', Times, serif",
            marginBottom: 0,
          }}
        >
          <thead>
            <tr style={{ backgroundColor: '#f8f9fa' }}>
              <th
                colSpan={2}
                style={{
                  border: '1px solid #000',
                  padding: '1px 4px',
                  textAlign: 'left',
                }}
              >
                Kehadiran :
              </th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td
                style={{
                  border: '1px solid #000',
                  width: '115px',
                  padding: '1px 4px',
                }}
              >
                Sakit
              </td>
              <td
                className="text-center"
                style={{ border: '1px solid #000', padding: '1px 4px' }}
              >
                : {student.attendance?.sakit || 0} hari
              </td>
            </tr>
            <tr>
              <td style={{ border: '1px solid #000', padding: '1px 4px' }}>
                Ijin
              </td>
              <td
                className="text-center"
                style={{ border: '1px solid #000', padding: '1px 4px' }}
              >
                : {student.attendance?.izin || 0} hari
              </td>
            </tr>
            <tr>
              <td style={{ border: '1px solid #000', padding: '1px 4px' }}>
                Tanpa Keterangan
              </td>
              <td
                className="text-center"
                style={{ border: '1px solid #000', padding: '1px 4px' }}
              >
                : {student.attendance?.alpa || 0} hari
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* CATATAN WALI KELAS */}
      <div
        className="mb-1.5"
        style={{
          border: '1px solid #000',
          padding: '4px 6px',
          fontFamily: "'Times New Roman', Times, serif",
          fontSize: '9.5pt',
          minHeight: '34px',
        }}
      >
        <div className="font-bold underline mb-0.5">Catatan Wali Kelas :</div>
        <div className="italic">
          {student.note || 'Tingkatkan terus motivasi dan kedisiplinan belajar.'}
        </div>
      </div>

      {/* SIGNATURE AREA USING 100% RIGID TABLES (PREVENTS OVERLAPS) */}
      <table
        className="w-full"
        style={{
          fontFamily: "'Times New Roman', Times, serif",
          fontSize: '9.5pt',
          borderCollapse: 'collapse',
          marginTop: '4px',
        }}
      >
        <tbody>
          <tr>
            <td
              style={{
                width: '50%',
                textAlign: 'center',
                verticalAlign: 'top',
                padding: 0,
              }}
            >
              <p className="mb-0">Orangtua/Walimurid,</p>
              <div style={{ height: '52px' }}></div>
              <p className="font-bold mb-0">
                ( ............................................ )
              </p>
            </td>
            <td
              style={{
                width: '50%',
                textAlign: 'center',
                verticalAlign: 'top',
                padding: 0,
              }}
            >
              <p className="mb-0">{school.datePrint}</p>
              <p className="mb-0">Wali Kelas,</p>
              <div style={{ height: '40px' }}></div>
              {formattedWaliName !== '-' && (
                <p
                  className={`font-bold mb-0 ${
                    formattedWaliId.formatted ? 'underline' : ''
                  }`}
                >
                  {formattedWaliName}
                </p>
              )}
              {formattedWaliId.formatted && (
                <p className="mb-0" style={{ fontSize: '8.5pt' }}>
                  {formattedWaliId.formatted}
                </p>
              )}
            </td>
          </tr>
          <tr>
            <td
              colSpan={2}
              style={{
                textAlign: 'center',
                verticalAlign: 'top',
                paddingTop: '6px',
              }}
            >
              <p className="mb-0">Mengetahui,</p>
              <p className="font-bold mb-0">Kepala SMKN 1 Mesuji Timur</p>
              <div style={{ height: '44px' }}></div>
              {formattedPrincipalName !== '-' && (
                <p
                  className={`font-bold mb-0 ${
                    formattedPrincipalId.formatted ? 'underline' : ''
                  }`}
                >
                  {formattedPrincipalName}
                </p>
              )}
              {formattedPrincipalId.formatted && (
                <p className="mb-0" style={{ fontSize: '8.5pt' }}>
                  {formattedPrincipalId.formatted}
                </p>
              )}
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  );
};
