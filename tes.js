async function generatePDF() {
    const data = await getDataFromExcel('nilai_siswa3.xlsx');

    // Iterasi setiap siswa untuk membuat PDF
    data.forEach(async student => {
        const content = []; // Konten untuk setiap siswa

        // Konten untuk setiap siswa (seperti yang Anda buat sebelumnya)
        // Saya akan meninggalkan konten ini kosong agar Anda dapat menyesuaikannya dengan kebutuhan Anda

        const docDefinition = {
            content: content,
            footer: function(currentPage, pageCount) {
                return { text: currentPage.toString() + ' / ' + pageCount, alignment: 'center' };
            },
            pageMargins: [40, 60, 40, 60],
            pageSize: 'A4',
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

        // Menggunakan nama siswa untuk menamai file PDF
        const pdfDoc = printer.createPdfKitDocument(docDefinition);
        pdfDoc.pipe(fs.createWriteStream(`daftar_nilai_siswa_${student.nama.replace(/\s+/g, '_')}.pdf`)); // Penamaan file PDF
        pdfDoc.end(() => {
            console.log(`Dokumen PDF untuk ${student.nama} telah dibuat.`);
        });
    });
}

generatePDF().catch(error => {
    console.error('Terjadi kesalahan:', error);
});
