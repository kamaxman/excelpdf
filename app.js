const fs = require('fs');
const PdfPrinter = require('pdfmake');

// Inisialisasi PdfPrinter dengan font yang diperlukan
const printer = new PdfPrinter({
  Roboto: {
    normal: 'fonts/Roboto-Regular.ttf',
    bold: 'fonts/Roboto-Bold.ttf',
    italics: 'fonts/Roboto-Italic.ttf',
    bolditalics: 'fonts/Roboto-BoldItalic.ttf',
  }
});

// Definisikan konten dokumen PDF
const docDefinition = {
  content: [
    { text: 'Contoh Dokumen PDF', style: 'header' },
    { text: 'Ini adalah contoh dokumen PDF sederhana yang dibuat menggunakan pdfmake.' },
    { text: 'Terima kasih!', style: 'footer' }
  ],
  styles: {
    header: {
      fontSize: 18,
      bold: true,
      alignment: 'center',
      margin: [0, 0, 0, 10] // margin: [left, top, right, bottom]
    },
    footer: {
      fontSize: 14,
      alignment: 'center',
      margin: [0, 10, 0, 0]
    }
  }
};

// Buat instance PdfKitDocument menggunakan printer
var pdfDoc = printer.createPdfKitDocument(docDefinition);

// Salurkan dokumen PDF ke file
pdfDoc.pipe(fs.createWriteStream('contoh.pdf'));
pdfDoc.end(() => {
  console.log('Dokumen PDF telah dibuat.');
});
