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
      const name = row.getCell(1).value;
      const mathScore = row.getCell(2).value;
      const physicsScore = row.getCell(3).value;
      const otherScore = row.getCell(4).value;
      const spiritualAttitude = row.getCell(5).value;
      const socialAttitude = row.getCell(6).value;
      const attendance = row.getCell(7).value;
      const izin = row.getCell(8).value;
      const alpa = row.getCell(9).value;

      if (
        typeof name === 'string' &&
        !isNaN(mathScore) &&
        !isNaN(physicsScore) &&
        !isNaN(otherScore) &&
        typeof spiritualAttitude === 'string' &&
        typeof socialAttitude === 'string' &&
        !isNaN(attendance) &&
        !isNaN(izin) &&
        !isNaN(alpa)
        
      ) {
        data.push({
          name: name,
          mathScore: mathScore,
          physicsScore: physicsScore,
          otherScore: otherScore,
          spiritualAttitude: spiritualAttitude,
          socialAttitude: socialAttitude,
          attendance: attendance,
          izin: izin,
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
    const data = await getDataFromExcel('nilai_siswa2.xlsx');
    
    const content = [
      { text: 'Daftar Nilai Siswa', style: 'header' },
      { text: '\n' }
    ];
  
    data.forEach(student => {
      const subjects = [
        { text: 'Matematika', score: student.mathScore },
        { text: 'Fisika', score: student.physicsScore },
        { text: 'Lain-lain', score: student.otherScore }
      ];
  
      const subjectTable = {
        table: {
          headerRows: 1,
          widths: ['auto', '*'],
          body: [
            ['Mata Pelajaran', 'Nilai'],
            ...subjects.map(subject => [subject.text, subject.score])
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
            ['Sikap Spiritual:', student.spiritualAttitude],
            ['Sikap Sosial:', student.socialAttitude]
          ],
          style: 'tableStyle'
        }
      };
  
      const attendanceTable = {
        table: {
          headerRows: 1,
          widths: ['auto','*'],
          body: [
            ['Kehadiran', 'Jumlah'],
            ['Hadir:', student.attendance],
            ['Izin:', student.izin],
            ['Alpa:', student.alpa]
          ],
          style: 'tableStyle'
        }
      };
  
      const studentContent = [
        { text: `Nama Siswa: ${student.name}`, style: 'subheader' },
        { text: '\n' },
        { text: 'Nilai Pelajaran', style: 'subheader' },
        subjectTable,
        { text: '\n' },
        { text: 'Sikap', style: 'subheader' },
        attitudeTable,
        { text: '\n' },
        { text: 'Kehadiran', style: 'subheader' },
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