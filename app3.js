async function generatePDF() {
    const data = await getDataFromExcel('nilai_siswa.xlsx');
    
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
          widths: ['*'],
          body: [
            [{ text: 'Sikap', colSpan: 2, style: 'subheader' }],
            [{ text: `Sikap Spiritual: ${student.spiritualAttitude}` }],
            [{ text: `Sikap Sosial: ${student.socialAttitude}` }]
          ],
          style: 'tableStyle'
        }
      };
  
      const attendanceTable = {
        table: {
          headerRows: 1,
          widths: ['*'],
          body: [
            ['Kehadiran'],
            [student.attendance]
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
        attitudeTable,
        { text: '\n' },
        { text: 'Kehadiran', style: 'subheader' },
        attendanceTable,
        { text: '\n\n' }
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
  