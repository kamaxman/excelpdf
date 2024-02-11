const fs = require('fs');
const PdfPrinter = require('pdfmake');
const ExcelJS = require('exceljs');

const printer = new PdfPrinter({
  Roboto: {
    normal: 'fonts/Roboto-Regular.ttf',
    bold: 'fonts/Roboto-Bold.ttf',
    italics: 'fonts/Roboto-Italic.ttf',
    bolditalics: 'fonts/Roboto-BoldItalic.ttf',
  }
});

async function getDataFromExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet(1);
  
  const data = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber !== 1) {
      const nama = row.getCell(1).value;
      const nis = row.getCell(2).value;
      const nisn = row.getCell(3).value;
      const alamat = row.getCell(4).value;
      const kelas = row.getCell(5).value;
      const semester = row.getCell(6).value;
      const tahun_pelajaran = row.getCell(7).value;
      const iman = row.getCell(8).value;
      const nalar = row.getCell(9).value;
      const mandiri = row.getCell(10).value;
      const global  = row.getCell(11).value;
      const kreatif = row.getCell(12).value;
      const gotongr = row.getCell(13).value;
      const pai = row.getCell(14).value;
      const pai_cp = row.getCell(15).value;
      const pkn = row.getCell(16).value;
      const pkn_cp = row.getCell(17).value;
      const bi = row.getCell(18).value;
      const bi_cp = row.getCell(19).value;
      const big = row.getCell(20).value;
      const big_cp = row.getCell(21).value;
      const mtk = row.getCell(22).value;
      const mtk_cp = row.getCell(23).value;
      const pkk = row.getCell(24).value;
      const pkk_cp = row.getCell(25).value;
      const aij = row.getCell(26).value;
      const aij_cp = row.getCell(27).value;
      const asj = row.getCell(28).value;
      const asj_cp = row.getCell(29).value;
      const tlj = row.getCell(30).value;
      const tlj_cp = row.getCell(31).value;
      const ekstra1 = row.getCell(32).value;
      const ekstra1_ket = row.getCell(33).value;
      const ekstra2 = row.getCell(34).value;
      const ekstra2_ket = row.getCell(35).value;
      const sakit = row.getCell(36).value;
      const izin  = row.getCell(37).value;
      const alpa = row.getCell(38).value;

      if (
        typeof nama === 'string' &&
        !isNaN(nis) &&
        !isNaN(nisn) &&
        typeof alamat === 'string' &&
        typeof kelas === 'string' &&
        typeof semester === 'string' &&
        typeof tahun_pelajaran === 'string' &&
        typeof iman === 'string' &&
        typeof nalar === 'string' &&
        typeof mandiri === 'string' &&
        typeof global  === 'string' &&
        typeof kreatif === 'string' &&
        typeof gotongr === 'string' &&
        !isNaN(pai) &&
        typeof pai_cp === 'string' &&
        !isNaN(pkn) &&
        typeof pkn_cp === 'string' &&
        !isNaN(bi) &&
        typeof bi_cp === 'string' &&
        !isNaN(big) &&
        typeof big_cp === 'string' &&
        !isNaN(mtk) &&
        typeof mtk_cp === 'string' &&
        !isNaN(pkk) &&
        typeof pkk_cp === 'string' &&
        !isNaN(aij) &&
        typeof aij_cp === 'string' &&
        !isNaN(asj) &&
        typeof asj_cp === 'string' &&
        !isNaN(tlj) &&
        typeof tlj_cp === 'string' &&
        typeof ekstra1 === 'string' &&
        typeof ekstra1_ket === 'string' &&
        typeof ekstra2 === 'string' &&
        typeof ekstra2_ket === 'string' &&
        !isNaN(sakit) &&
        !isNaN(izin ) &&
        !isNaN(alpa)
        
      ) {
        data.push({
          nama: nama,
          nis: nis,
          nisn: nisn,
          alamat: alamat,
          kelas: kelas,
          semester: semester,
          tahun_pelajaran: tahun_pelajaran,
          iman: iman,
          nalar: nalar,
          mandiri: mandiri,
          global : global ,
          kreatif: kreatif,
          gotongr: gotongr,
          pai: pai,
          pai_cp: pai_cp,
          pkn: pkn,
          pkn_cp: pkn_cp,
          bi: bi,
          bi_cp: bi_cp,
          big: big,
          big_cp: big_cp,
          mtk: mtk,
          mtk_cp: mtk_cp,
          pkk: pkk,
          pkk_cp: pkk_cp,
          aij: aij,
          aij_cp: aij_cp,
          asj: asj,
          asj_cp: asj_cp,
          tlj: tlj,
          tlj_cp: tlj_cp,
          ekstra1: ekstra1,
          ekstra1_ket: ekstra1_ket,
          ekstra2: ekstra2,
          ekstra2_ket: ekstra2_ket,
          sakit: sakit,
          izin : izin ,
          alpa: alpa
        });
      } else {
        console.error(`Data pada baris ${rowNumber} tidak valid. Melewatkan baris ini.`);
      }
    }
  });
  return data;
}


async function generatePDF() {
    const data = await getDataFromExcel('nilai_siswa3.xlsx');
    
    const content = [
    
    ];
  
    data.forEach(student => {
      const subjects = [
        { text: 'Matematika', score: student.mtk },
        { text: 'Fisika', score: student.pai },
        { text: 'Lain-lain', score: student.pkn }
      ];

      const nasional = [
        { text: 'Matematika', score: student.mtk },
        { text: 'Fisika', score: student.pai },
        { text: 'Lain-lain', score: student.pkn }
      ];

      const identitiTable = {
        layout: 'noBorders',
        table: {
          headerRows: 0,
          widths: ['auto', 'auto', '*', 'auto', 'auto', '*'],
          body: [
           
            ['Nama Peserta Didik', ':', student.nama, 'Kelas', ':', student.kelas],
            ['Nomor Induk/NISN', ':', student.nis +'/'+student.nisn, 'Semester', ':', student.semester],
            ['Sekolah', ':', 'SMKN ...', 'Tahun Pelajaran', ':', student.tahun_pelajaran],
            ['Alamat', ':', student.alamat, '', '', '']
          ],
          style: 'tableStyle'
          
        }
      };

      const attitudeTable = {
        table: {
          headerRows: 1,
          widths: ['auto', '*'],
          body: [
           
            ['Judul', 'Isi'],
            ['Sikap Spiritual:', student.iman],
            ['Sikap Sosial:', student.nalar]
          ],
          style: 'tableStyle'
        }
      };
  
      

      const subjectTable = {
        table: {
          headerRows: 1,
          widths: ['auto', '*', '*'], // Kolom pertama untuk nomor, kolom kedua dan ketiga untuk konten
          body: [
           
            ['No ', 'Mata Pelajaran', 'Nilai'], // Kolom pertama kosong untuk baris header kedua
            [
             
              {
                text: 'A. Muatan Nasional',
                colSpan: 3, // Menggabungkan dua kolom
                alignment: 'left'
              },
              {},
              {}             
            ],
            ...subjects.map((subject, index) => [index + 1, subject.text, subject.score]), // Menambahkan nomor urutan
            [
             
              {
                text: 'C. Peminatan Keahlian',
                colSpan: 3, // Menggabungkan dua kolom
                alignment: 'left'
              },
              {},
              {}
            ],
            [
             
              {
                text: 'C3. Kompetensi Kejuruan',
                colSpan: 3, // Menggabungkan dua kolom
                alignment: 'left'
              },
              {},
              {}
            ],
            ...nasional.map((nasional, index) => [index + 1, nasional.text, nasional.score]) // Menambahkan nomor urutan
          ],
          style: 'tableStyle'
        }
      };
      
      
      

      

      const ekstraTable = {
        table: {
          headerRows: 1,
          widths: ['auto', 'auto','*'],
          body: [
            ['1.', student.ekstra1, student.ekstra1_ket],
            ['2.', student.ekstra2, student.ekstra2_ket]
          ],
          style: 'tableStyle'
        }
      };
  
      
  
      const attendanceTable = {
        table: {
          headerRows: 1,
          widths: ['auto','*'],
          body: [
            ['Sakit:', student.sakit],
            ['Izin:', student.izin],
            ['Alpa:', student.alpa]
          ],
          style: 'tableStyle'
        }
      };
  
      const studentContent = [
        identitiTable,
      
        { text: '\n' },
        { text: 'A. Sikap', style: 'subheader' },
        attitudeTable,
        { text: '\n' },
        { text: 'B. Akademik', style: 'subheader' },
        subjectTable,
        { text: '\n' },
        { text: 'C. Ekstrakurikuler', style: 'subheader' },
        ekstraTable,
        { text: '\n' },
        { text: 'D. Kehadiran', style: 'subheader' },
        attendanceTable,
        { text: '\n\n' ,pageBreak: 'after' }
       
      ];
  
      content.push(...studentContent);
    });
  
    const docDefinition = {
      content: content,
      styles: {
        header: {
          fontSize: 18,
          bold: true,
          alignment: 'center',
          margin: [0, 0, 0, 10]
        },
        subheader: {
          fontSize: 16,
          bold: true,
          margin: [0, 10, 0, 5]
        },
        tableHeader: {
          bold: true,
          fontSize: 13,
          color: 'black'
        },
        tableStyle: {
          margin: [0, 5, 0, 15],
          alignment: 'center',
          border: 'solid'
        }
       
        
      }
    };
  
    const pdfDoc = printer.createPdfKitDocument(docDefinition);
  
    pdfDoc.pipe(fs.createWriteStream(`daftar_nilai_siswa.pdf`));
    pdfDoc.end(() => {
      console.log('Dokumen PDF telah dibuat.');
    });
  }
  
  
  generatePDF().catch(error => {
    console.error('Terjadi kesalahan:', error);
  });