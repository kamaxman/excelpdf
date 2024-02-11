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
      data.push({
        name: row.getCell(1).value,
        mathScore: row.getCell(2).value,
        physicsScore: row.getCell(3).value,
        otherScore: row.getCell(4).value,
        spiritualAttitude: row.getCell(5).value,
        socialAttitude: row.getCell(6).value,
        attendance: row.getCell(7).value // Menggunakan kolom attendance untuk kehadiran
      });
    }
  });
  return data;
}

async function generatePDF() {
  const data = await getDataFromExcel('nilai_siswa.xlsx');
  
  const content = [
    { text: 'Daftar Nilai Siswa', style: 'header' },
    { text: '\n' }
  ];

  // Tabel untuk nilai pelajaran
  const subjectsTable = {
    table: {
      headerRows: 1,
      widths: ['auto', 'auto', 'auto', 'auto'],
      body: [
        ['Nama Siswa', 'Matematika', 'Fisika', 'Lain-lain'],
        ...data.map(student => [student.name, student.mathScore, student.physicsScore, student.otherScore])
      ]
    }
  };

  // Tabel untuk nilai sikap
  const attitudeTable = {
    table: {
      headerRows: 1,
      widths: ['auto', 'auto', 'auto'],
      body: [
        ['Nama Siswa', 'Sikap Spiritual', 'Sikap Sosial'],
        ...data.map(student => [student.name, student.spiritualAttitude, student.socialAttitude])
      ]
    }
  };

  // Tabel untuk kehadiran
  const attendanceTable = {
    table: {
      headerRows: 1,
      widths: ['auto', 'auto'],
      body: [
        ['Nama Siswa', 'Kehadiran'],
        ...data.map(student => [student.name, student.attendance])
      ]
    }
  };

  content.push(
    { text: 'Nilai Pelajaran', style: 'subheader' },
    subjectsTable,
    { text: '\n\n' },
    { text: 'Sikap', style: 'subheader' },
    attitudeTable,
    { text: '\n\n' },
    { text: 'Kehadiran', style: 'subheader' },
    attendanceTable,
    { text: '\nTerima kasih!', style: 'footer' }
  );

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
      footer: {
        fontSize: 14,
        alignment: 'center',
        margin: [0, 10, 0, 0]
      }
    }
  };

  const pdfDoc = printer.createPdfKitDocument(docDefinition);

  pdfDoc.pipe(fs.createWriteStream('daftar_nilai_siswa.pdf'));
  pdfDoc.end(() => {
    console.log('Dokumen PDF telah dibuat.');
  });
}

generatePDF().catch(error => {
  console.error('Terjadi kesalahan:', error);
});
