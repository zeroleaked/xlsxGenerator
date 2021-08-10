class GefExcel extends ExcelJS.Workbook {
    constructor() {
        super();
        this.creator = "who";
        this.created = new Date();
        this.modified = new Date();
        this.promises = [];
    }

    saveAs(fileName) {
        Promise.allSettled(this.promises)
            .then( p => {
                // console.log(p)
                return this.xlsx.writeBuffer();
            })
            .then( buffer => {
                const blob = new Blob([buffer], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
                saveAs(blob, fileName);
            })
    }

    addSheetPeralatan(data) {
        const sheetName = "Daftar dan Performa Peralatan";
        const ws = this.addWorksheet(sheetName, {
            pageSetup: {paperSize: 9, orientation:'landscape', fitToPage:true, fitToWidth:1, fitToHeight: 0},
            views: [{ style: "pageBreakPreview" }, {state: 'frozen', xSplit: 4, ySplit: 9}]
        });

        let rowCount = 0;
        let dataCount = 0;
        const numberOfDays = new Date(data.tahunInt, data.bulanInt, 0).getDate();
        const maxCol = numberOfDays + 8;
        const totalHours = 24 * numberOfDays;

        // set column options
        ws.columns = [
            { width: 8.09},
            { width: 30.18, horizontal: "left" },
            { width: 33.18},
            { width: 20.09},
            ...Array(numberOfDays).fill({ width: 8.00 }),
            { width: 16.64 },
            { width: 8.09},
            { width: 15.82},
            { width: 8.09},
        ].map( elm => ({
            width: elm.width,
            style: {
                font: {size: 11},
                alignment: {
                    horizontal: elm.horizontal ? elm.horizontal : "center",
                    vertical: "middle",
                    wrapText: true
                }
            },
        }));

        addHeader(data);
        const percentages = data.peralatan.map(p => addData(p));
        addFooter(data);

        return { sheetName, percentages };

        function addHeader({ fasilitas, bulanTahun, lembaran2, lembaran3 }) {
            addRow([]);
    
            addRow(["LAPORAN BULANAN UNJUK HASIL (PERFORMANCE)"]);
            ws.mergeCells(2, 1, 2, maxCol);
            currentRow().getCell('A').font = { bold: true, size: 11 };
            // currentRow().getCell('A').alignment = { wrapText: false };
    
            addRow([]);
    
            addRow([null, "Cabang Bandara", ": BANDARA SOEKARNO HATTA - TANGERANG",
                ...nullArr(numberOfDays+1), "LEMBARAN 1", ": DITJEN HUBUD" ]);
            currentRow().getCell('C').alignment = { horizontal: "left" };
            currentRow().getCell(maxCol-2).alignment = { horizontal: "left" };
    
            addRow([null, "Fasilitas", `: ${fasilitas}`, ...nullArr(numberOfDays+1),
                "LEMBARAN 2", `: ${lembaran2}` ]);
            currentRow().getCell('C').alignment = { horizontal: "left" };
            currentRow().getCell(maxCol-2).alignment = { horizontal: "left" };
    
            addRow([null, "Bulan / Tahun", `: ${bulanTahun}`, ...nullArr(numberOfDays+1),
                "LEMBARAN 3", `: ${lembaran3}` ]);
            currentRow().getCell('C').alignment = { horizontal: "left" };
            currentRow().getCell(maxCol-2).alignment = { horizontal: "left" };
    
            addRow([]);
    
            addRow(["No", "NAMA PERALATAN", "LOKASI", "MEREK", "TANGGAL",
                ...nullArr(numberOfDays-1), "JML JAM TERPUTUS", "JML JAM OPS / BL",
                "UNJUK HASIL", "KET" ]);
            ws.mergeCells(rowCount, 5, rowCount, numberOfDays+4);
            currentRow().getCell('B').alignment = { horizontal: "center", vertical: "middle", wrapText: true };
            for (let i = 0; i < maxCol; i++)
                currentRow().getCell(i+1).border = {
                    top: {style: "thin"},
                    bottom: {style: "thin"},
                    left: {style: "thin"},
                    right: {style: "thin"},
                };
    
            // baris tanggal
            addRow([
                ...nullArr(4),
                ...Array.from(Array(numberOfDays).keys(), i => i+ 1)
            ]);
            currentRow().font = { bold: true, size: 11 };
            for (let i = 0; i < maxCol; i++)
                currentRow().getCell(i+1).border = {
                    top: {style: "thin"},
                    bottom: {style: "thin"},
                    left: {style: "thin"},
                    right: {style: "thin"},
                };
            ws.mergeCells(rowCount-1, 1, rowCount, 1);
            ws.mergeCells(rowCount-1, 2, rowCount, 2);
            ws.mergeCells(rowCount-1, 3, rowCount, 3);
            ws.mergeCells(rowCount-1, 4, rowCount, 4);
            ws.mergeCells(rowCount-1, maxCol-3, rowCount, maxCol-3);
            ws.mergeCells(rowCount-1, maxCol-2, rowCount, maxCol-2);
            ws.mergeCells(rowCount-1, maxCol-1, rowCount, maxCol-1);
            ws.mergeCells(rowCount-1, maxCol, rowCount, maxCol);
    
            addRow(["I", "Non Terminal"]);
            currentRow().font = { bold: true, size: 11 };
            for (let i = 0; i < maxCol; i++)
                currentRow().getCell(i+1).border = {
                    top: {style: "double"},
                    bottom: {style: "dotted"},
                    left: {style: "thin"},
                    right: {style: "thin"},
                };
        }
    
        function addData({ fasilitas, alat }) {
            dataCount++;
    
            const setTableBorder = () => {
                for (let i = 0; i < maxCol; i++)
                    currentRow().getCell(i+1).border = {
                        top: {style: "dotted"},
                        bottom: {style: "dotted"},
                        left: {style: "thin"},
                        right: {style: "thin"},
                    };
            }
    
            const numberToLetter = (n) => {
                const ordA = 'A'.charCodeAt(0);
                const ordZ = 'Z'.charCodeAt(0);
                const len = ordZ - ordA + 1;
              
                let s = "";
                n--;
                while(n >= 0) {
                    s = String.fromCharCode(n % len + ordA) + s;
                    n = Math.floor(n / len) - 1;
                }
                return s;
            }
            addRow([`${numberToLetter(dataCount)}.`, fasilitas]);
            currentRow().fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: "FFFFFF00"} // yellow
            };
            currentRow().font = { bold: true, size: 11 };
            setTableBorder();
            const name = currentRow().getCell('B')._address;
    
            // baris tiap alat
            alat.forEach( (a, i) => {
                const perHari = Array(numberOfDays).fill(0);
                a.jamPutus.forEach( elm => perHari[elm.tanggal-1] = elm.jam );
                addRow([
                    i+1,
                    a.nama,
                    a.lokasi,
                    a.merek ? a.merek : '',
                    ...perHari,
                    null,
                    totalHours
                ]);
    
                const totalPutusCell = currentRow().getCell(maxCol-3);
                const firstDayAddress = currentRow().getCell('E')._address;
                const lastDayAddress = currentRow().getCell(numberOfDays+4)._address;
                totalPutusCell.value = {
                    formula: `SUM(${firstDayAddress}:${lastDayAddress})`
                }
                totalPutusCell.numFmt = "0;-0;;@";
    
                const percentageCell = currentRow().getCell(maxCol-1);
                const totalCellAddress = currentRow().getCell(maxCol-2)._address;
                const putusCellAddress = currentRow().getCell(maxCol-3)._address;
                percentageCell.value = {
                    formula: `(${totalCellAddress}-${putusCellAddress})/${totalCellAddress}*100`
                };
                percentageCell.numFmt = "0.00"
    
                setTableBorder();
            })
    
            addRow([]);
            setTableBorder();
    
            // average persentase fasilitas
            const averageCell = currentRow().getCell(maxCol-1);
            const firstPercentageAddress = ws.getRow(rowCount-alat.length).getCell(maxCol-1)._address;
            const lastPercentageAddress = ws.getRow(rowCount-1).getCell(maxCol-1)._address;
            averageCell.value = {
                formula: `AVERAGE(${firstPercentageAddress}:${lastPercentageAddress})`
            };
            averageCell.numFmt = "0.00";
            const percentage = averageCell._address;
    
            // format jam rusak perhari
            const firstCondAddress = ws.getRow(rowCount-alat.length).getCell('E')._address;
            const lastCondAddress = ws.getRow(rowCount-1).getCell(numberOfDays + 4)._address;
            ws.addConditionalFormatting({
                ref: `${firstCondAddress}:${lastCondAddress}`,
                rules: [
                    {
                    type: 'cellIs',
                    operator: 'equal',
                    formulae: [0],
                    style: {font: {color: {argb: 'FF008000'}}, size: 11},
                    },
                    {
                    type: 'cellIs',
                    operator: 'greaterThan',
                    formulae: [0],
                    style: {font: {color: {argb: 'FFFF0000'}}, size: 11},
                    }
                ]
            })
    
            return { name, percentage }; //addresses
        }
    
        function addFooter({ jabatanManajer, namaManajer, jabatanAsisten, namaAsisten, bulanTahun }) {
            addRow([]);
            const formula = percentages.reduce(
                    (reducer, facility, i) => reducer + (i ? ',' : '') + facility.percentage,
                    "AVERAGE("
                ) + ')';
            const averageCell = currentRow().getCell(maxCol-1);
            averageCell.value = { formula };
            averageCell.numFmt = "0.00";
            currentRow().fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: "FFFFC000"}
            };
            for (let i = 0; i < maxCol; i++)
                currentRow().getCell(i+1).border = {
                    top: {style: 'thin'},
                    bottom: {style: 'thin'},
                    left: {style: 'thin'},
                    right: {style: 'thin'},
                };
                
            addRow([]); addRow([]);
    
            addRow([ ...nullArr(numberOfDays), `Tangerang,          ${bulanTahun}` ]);
            currentRow().alignment = { textWrap: false, horizontal: 'center'}
    
            addRow([]);
    
            addRow([ ...nullArr(numberOfDays), "PT. ANGKASA PURA II (Persero)" ]);
            currentRow().alignment = { textWrap: false, horizontal: 'center'}
    
            addRow([ ...nullArr(numberOfDays), "Cabang Utama Bandara Soekarno - Hatta, Tangerang" ]);
            currentRow().alignment = { textWrap: false, horizontal: 'center'}
    
            addRow([ ...nullArr(2), "Mengetahui", ]);
            currentRow().alignment = { textWrap: false, horizontal: 'center'}
    
            addRow([ ...nullArr(2), jabatanManajer, ...nullArr(numberOfDays-3), jabatanAsisten ]);
            currentRow().alignment = { textWrap: false, horizontal: 'center'}
    
            for (let i = 0; i < 6; i++) addRow([]);
    
            addRow([ ...nullArr(2), namaManajer, ...nullArr(numberOfDays-3), namaAsisten ]);
            currentRow().alignment = { textWrap: false, horizontal: 'center'}
            currentRow().font = { bold: true, underline: true, size: 11 }
    
            addRow([" "]);
        }
    
        function addRow(row) {
            ws.addRow(row);
            rowCount ++;
        }
    
        function currentRow() {
            return ws.getRow(rowCount);
        }
    
        function nullArr(length) {
            return Array(length).fill(null);
        }
    }

    addSheetPersonil(data) {
        const ws = this.addWorksheet("Daftar Personil", {
            pageSetup:{paperSize: 9, orientation:'landscape', fitToPage:true, fitToWidth:1, fitToHeight: 0},
            views: [{ style: "pageBreakPreview" }]
        });
        let rowCount = 0;
        let dataCount = 0;

        // // set column width
        ws.columns = [
            { width: 3.82},
            { width: 10.91, horizontal: "left" },
            { width: 18.64},
            { width: 14.55},
            { width: 4.09},
            { width: 18.09 },
            { width: 5.64},
            { width: 8.18},
            { width: 20.36},
            { width: 24.09},
            { width: 13.82},
            { width: 7.91},
            { width: 10.91},
            { width: 10.36},
            { width: 13.82},
        ].map( elm => ({
            width: elm.width,
            style: {
                font: { size: 11 },
                alignment: {
                    horizontal: elm.horizontal ? elm.horizontal : "center",
                    vertical: "middle",
                    wrapText: true
                }
            },
        }));

        addHeader(data);
        data.personil.forEach(p => addData(p));
        addFooter(data);

        function addHeader({ fasilitas, bulanTahun, lembaran2, lembaran3 }) {
            addRow([]);
    
            addRow(["DAFTAR PERSONIL (BULANAN)"]);
            ws.mergeCells(rowCount, 1, rowCount, 15);
            currentRow().font = { bold: true, size: 11 };
    
            addRow([]);
    
            addRow(["Cabang Bandara", null, ": BANDARA SOEKARNO HATTA - TANGERANG", ...nullArr(7),
                "LEMBARAN I", ": DITJEN HUBUD"]);
            currentRow().alignment = { horizontal: 'left' };
    
            addRow([ "Fasilitas", null, `: ${fasilitas}`, ...nullArr(7), "LEMBARAN II",
                `: ${lembaran2}`]);
            currentRow().alignment = { horizontal: 'left' };
    
            addRow(["Bulan / Tahun", null, `: ${bulanTahun}`, ...nullArr(7), "LEMBARAN III",
                 `: ${lembaran3}`]);
            currentRow().alignment = { horizontal: 'left' };
    
            addRow([]); addRow([]);
    
            addRow(["NO", "NAME", null, "NIK", "KLS JBT", "JABATAN", "CABANG", null, "UNIT KERJA",
                "LICENSE", null, "RATING", null, null, "STATUS"])
            ws.mergeCells(rowCount, 7, rowCount, 8);
            ws.mergeCells(rowCount, 10, rowCount, 11);
            ws.mergeCells(rowCount, 12, rowCount, 14);
            currentRow().alignment = {horizontal: "center", vertical: "middle", wrapText: true};
            for (let i = 0; i < 15; i++)
                currentRow().getCell(i+1).border = {
                    top: {style:'thin'},
                    left: {style:'thin'},
                    bottom: {style:'thin'},
                    right: {style:'thin'}
                };
    
            addRow([...nullArr(6), "KODE", "BANDARA", null, "NOMOR", "TINGKAT", "NAMA", "TMT", "S/D"]);
            ws.mergeCells(rowCount-1, 1, rowCount, 1);
            ws.mergeCells(rowCount-1, 2, rowCount, 3);
            ws.mergeCells(rowCount-1, 4, rowCount, 4);
            ws.mergeCells(rowCount-1, 5, rowCount, 5);
            ws.mergeCells(rowCount-1, 6, rowCount, 6);
            ws.mergeCells(rowCount-1, 9, rowCount, 9);
            ws.mergeCells(rowCount-1, 15, rowCount, 15);
            currentRow().alignment = {horizontal: "center", vertical: "middle", wrapText: true};
            for (let i = 0; i < 15; i++)
                currentRow().getCell(i+1).border = {
                    top: {style:'thin'},
                    left: {style:'thin'},
                    bottom: {style:'thin'},
                    right: {style:'thin'}
                };
    
            addRow([]);
            ws.mergeCells(rowCount, 2, rowCount, 3);
            for (let i = 0; i < 15; i++)
                currentRow().getCell(i+1).border = {
                    top: {style:'double'},
                    left: {style:'thin'},
                    right: {style:'thin'}
                };

        }
        
        function addData({ name, nik, kelas, jabatan, unit, nomorLicense, tingkatLicense, namaRating,
                tmtRating, sdRating }) {
            dataCount++;
            
            const isValidDate = d => d instanceof Date && !isNaN(d);
            
            let status;
            const sdRatingDate = new Date(sdRating);
            if (isValidDate(sdRatingDate))
                if (sdRatingDate > new Date()) status = "OK"
                else status = "Expired"
            else status = "-"
    
            addRow([dataCount, name, null, nik, kelas, jabatan, "1", "BSH", unit, nomorLicense,
                tingkatLicense, namaRating, tmtRating, sdRating, status]);
            ws.mergeCells(rowCount, 2, rowCount, 3);
            for (let i = 0; i < 15; i++)
                currentRow().getCell(i+1).border = {
                    left: {style:'thin'},
                    right: {style:'thin'}
                };
            currentRow().getCell('M').numFmt = "dd-mmm-yy"
            currentRow().getCell('N').numFmt = "dd-mmm-yy"
            currentRow().getCell('O').font = { color: { argb: 'FFFF0000' }, size: 11 } // red default
            ws.addConditionalFormatting({
                ref: `O${rowCount}`,
                rules: [{
                    type: 'containsText',
                    operator: 'containsText',
                    text: "OK",
                    style: {font: {color: {argb: 'FF0000FF'}, bold: true, size: 11 } }, // blue if OK
                }]
            })
    
            addRow([]);
            ws.mergeCells(rowCount, 2, rowCount, 3);
            for (let i = 0; i < 15; i++)
                currentRow().getCell(i+1).border = {
                    left: {style:'thin'},
                    right: {style:'thin'}
                };
        }
    
        function addFooter({ bulanTahun, jabatanManajer, namaManajer, jabatanAsisten, namaAsisten }) {
            addRow([]);
            for (let i = 0; i < 15; i++)
                currentRow().getCell(i+1).border = { top: { style: "thin" } };
    
            addRow([]);
    
            addRow([...nullArr(12), `Tangerang,     ${bulanTahun}`]);
            currentRow().alignment = { wrapText: false, horizontal: 'center' };
    
            addRow([])
    
            addRow([...nullArr(12), "PT. ANGKASA PURA II (PERSERO)"])
            currentRow().alignment = { wrapText: false, horizontal: 'center' };
    
            addRow([...nullArr(12), "Cabang Utama Bandara Soekarno-Hatta, Tangerang"]);
            currentRow().alignment = { wrapText: false, horizontal: 'center' };
    
            addRow([null, "Mengetahui"]);
    
            addRow([null, jabatanManajer, ...nullArr(10), jabatanAsisten]);
            currentRow().alignment = { wrapText: false, horizontal: 'center' };
    
            for (let i = 0; i < 4; i++)
                addRow([]);
    
            addRow([null, namaManajer, ...nullArr(10), namaAsisten]);
            currentRow().font = { bold: true, underline: true, size: 11 }
            currentRow().alignment = { wrapText: false, horizontal: 'center' };
    
            addRow([" "])
        }

        function addRow (row) {
            ws.addRow(row);
            rowCount ++;
        }

        function currentRow() {
            return ws.getRow(rowCount);
        }

        function nullArr(length) {
            return Array(length).fill(null);
        }

    }
    
    addSheetKegiatan(data) {
        const ws = this.addWorksheet("Laporan Kegiatan Perbaikan", {
            pageSetup:{paperSize: 9, orientation:'landscape', fitToPage:true, fitToWidth:1, fitToHeight: 0},
            views: [{ style: "pageBreakPreview" }]
        });

        let rowCount = 0;
        let dataCount = 0;

        // set column width
        ws.columns = [
            { width: 23.36},
            { width: 19.5, horizontal: "left" },
            { width: 15.55},
            { width: 23.91},
            { width: 46.82, horizontal: "left" },
            { width: 15.91},
            { width: 17.36},
            { width: 15.09},
            { width: 18.55},
            { width: 47.82},
        ].map( elm => ({
            width: elm.width,
            style: {
                font: { size: 11 },
                alignment: {
                    horizontal: elm.horizontal ? elm.horizontal : "center",
                    vertical: "middle",
                    wrapText: true
                }
            },
        }));

        addHeader(data);

        const promises = data.kegiatan
            .map((p) => addData(this, p) )
        this.promises.push(Promise.allSettled(promises));

        addFooter(data)
    
        function addHeader({ fasilitas, bulanTahun, lembaran2, lembaran3 }) {
            addRow([]);
    
            addRow(["LAPORAN KEGIATAN PERBAIKAN (BULANAN)"]);
            ws.mergeCells(rowCount, 1, rowCount, 10);
            currentRow().getCell('A').font = { bold: true, size: 11 };
    
            addRow([]);
    
            addRow(["Cabang Bandara", null, ": Bandara Soekarno - Hatta, Tanggerang", null, null, "LEMBARAN I", ": DITJEN HUBUD"]);
            currentRow().getCell('C').alignment = { horizontal: 'left', wrapText: false }
            currentRow().getCell('G').alignment = { horizontal: 'left', wrapText: false }
    
            addRow(["Fasilitas", null, `: ${fasilitas}`, null, null,  "LEMBARAN II", `: ${lembaran2}`]);
            currentRow().getCell('C').alignment = { horizontal: 'left', wrapText: false }
            currentRow().getCell('G').alignment = { horizontal: 'left', wrapText: false }
    
            addRow(["Bulan / Tahun", null, `: ${bulanTahun}`, null, null, "LEMBARAN III",  `: ${lembaran3}`]);
            currentRow().getCell('C').alignment = { horizontal: 'left', wrapText: false }
            currentRow().getCell('G').alignment = { horizontal: 'left', wrapText: false }
            
            addRow([]);
    
            addRow(["NO", "NAMA PERALATAN", "KERUSAKAN KATEGORI", "BAGIAN", "TINDAKAN", "TGL/JAM KERUSAKAN", "TGL/JAM SELESAI", "JAM TERPUTUS", "KETERANGAN", "DOKUMENTASI"]);
            currentRow().font = { bold: true, size: 11 };
            currentRow().alignment = { horizontal: "center", vertical: "middle", wrapText: true };
            for (let i = 0; i < 10; i++) {
                currentRow().getCell(i+1).border = {
                    top: {style:'thin'},
                    left: {style:'thin'},
                    bottom: {style:'thin'},
                    right: {style:'thin'}
                }
            }
        }
        
        function addData(wb, { nama, lokasi, kategori, bagian, tindakan, tanggalRusak, jamRusak,
            tanggalSelesai, jamSelesai, jamPutus, keterangan, gambar }) {
            dataCount++;
    
            addRow([dataCount, `${nama}\nLokasi: ${lokasi}`, kategori, bagian, tindakan,
                    `${tanggalRusak}/${jamRusak}`, `${tanggalSelesai}/${jamSelesai}`, jamPutus, keterangan]);
            currentRow().height = 200;
            for (let i = 0; i < 10; i++)
                currentRow().getCell(i+1).border = {
                    top: {style:'thin'},
                    left: {style:'thin'},
                    bottom: {style:'thin'},
                    right: {style:'thin'}
                };
    
            const imageDownload = new Promise( (resolve, reject) => {
                const row = rowCount;
                const img = new Image();
                img.setAttribute('crossorigin', 'anonymous');
                img.src = gambar;
                img.onload = () => {
                    // convert to base64 jpeg
                    const canvas = document.createElement("canvas");
                    canvas.width = img.width;
                    canvas.height = img.height;
                    canvas.getContext("2d").drawImage(img, 0, 0);
                    const base64 = canvas.toDataURL("image/jpeg");
                    const extension = base64.split(";", 2)[0].split("/")[1];
    
                    // add to worksheet
                    const id = wb.addImage({ base64, extension })
                    const ext = {};
                    const maxWidth = 340; const maxHeight = 255;
                    const ratio = img.width/img.height;
                    if ( ratio > maxWidth/maxHeight ) {
                        ext.width = maxWidth;
                        ext.height = Math.round(maxWidth/ratio);
                    } else {
                        ext.height = maxHeight;
                        ext.width = Math.round(maxHeight*ratio);
                    }
                    const tl = { col: 9.1, row: row-0.6 };
                    ws.addImage( id, { tl, ext });
    
                    resolve({ id, row });
                }
                img.onerror = () => reject(`Failed to load ${gambar} on row ${rowCount}`);
            })
    
            return imageDownload;
        }
    
        function addFooter({ bulanTahun, jabatanManajer, namaManajer, jabatanAsisten, namaAsisten }) {
            addRow([]);
    
            addRow([...nullArr(5), `Tangerang,       ${bulanTahun}`]);
            ws.mergeCells(rowCount, 6, rowCount, 9);
    
            addRow([...nullArr(5), "PT. ANGKASA PURA II (PERSERO)"]);
            ws.mergeCells(rowCount, 6, rowCount, 9);
    
            addRow(["Mengetahui", ...nullArr(4), "Cabang Utama Bandara Soekarno-Hatta, Tangerang"]);
            ws.mergeCells(rowCount, 1, rowCount, 3);
            ws.mergeCells(rowCount, 6, rowCount, 9);
    
            addRow([jabatanManajer, ...nullArr(4), jabatanAsisten])
            ws.mergeCells(rowCount, 1, rowCount, 3);
            ws.mergeCells(rowCount, 6, rowCount, 9);
    
            for (let i = 0; i < 4; i++) addRow([]);
    
            addRow([namaManajer, ...nullArr(4), namaAsisten]);
            ws.mergeCells(rowCount, 1, rowCount, 3);
            ws.mergeCells(rowCount, 6, rowCount, 9);
            currentRow().getCell('A').font = { bold: true, underline: true, size: 11 };
            currentRow().getCell('F').font = { bold: true, underline: true, size: 11 };
    
            addRow([" "])
        }
    
        function addRow(row) {
            ws.addRow(row);
            rowCount ++;
        }
    
        function currentRow() {
            return ws.getRow(rowCount);
        }
    
        function nullArr(length) {
            return Array(length).fill(null);
        }
    }

    addSheetGrafik(data, sheetPeralatan) {
        const ws = this.addWorksheet("Grafik Performance Peralatan", {
            pageSetup:{paperSize: 9, orientation:'portrait', fitToPage:true, fitToWidth:1, fitToHeight: 0},
            views: [{ style: "pageBreakPreview" }]
        });
        let rowCount = 0;
        let dataCount = 0;

        // // set column width
        ws.columns = [
            { width: 6.09},
            { width: 24.91 },
            { width: 16.09},
        ].map( elm => ({
            width: elm.width ? elm.width : 8.00,
            style: {
                font: {size: 11},
                alignment: {
                    horizontal: elm.horizontal ? elm.horizontal : "center",
                    vertical: "middle",
                }
            },
        }));

        addHeader();
        sheetPeralatan.percentages.forEach(p => addData(p, sheetPeralatan.sheetName));
        addFooter(data);

        function addHeader() {
            addRow([]);
    
            addRow([...nullArr(6), "DATA PERFORMANCE PERALATAN BULAN MEI 2021"]);
            currentRow().font = { bold: true, size: 11 };
    
            addRow([]); addRow([]);
    
            addRow(["No", "Peralatan", "Performance"]);
            currentRow().font = { bold: true, size: 11 };
            for (let i = 0; i < 3; i++)
                currentRow().getCell(i+1).border = {
                    top: {style: "thin"},
                    bottom: {style: "thin"},
                    left: {style: "thin"},
                    right: {style: "thin"},
                };
        }
        
        function addData({ name, percentage }, sheetRef) {
            dataCount++;
    
            addRow([dataCount]);
    
            const peralatan = currentRow().getCell('B');
            peralatan.value = {
                formula: `='${sheetRef}'!${name}`
            };
    
            const performance = currentRow().getCell('C');
            performance.value = {
                formula: `='${sheetRef}'!${percentage}`
            };
            performance.numFmt = "0.00";
    
            // style
            for (let i = 0; i < 3; i++)
                currentRow().getCell(i+1).border = {
                    top: {style: "thin"},
                    bottom: {style: "thin"},
                    left: {style: "thin"},
                    right: {style: "thin"},
                };
        }
    
        function addFooter({ jabatanManajer, namaManajer, jabatanAsisten, namaAsisten, bulanTahun }) {
            for (let i = 0; i < 15; i++)
                addRow([]);
    
            addRow([ ...nullArr(12), `Tangerang,          ${bulanTahun}` ]);
    
            addRow([]);
            
            addRow([ ...nullArr(12), "PT. ANGKASA PURA II (Persero)" ]);
            
            addRow([ null, "Mengetahui", ...nullArr(10), "Cabang Utama Bandara Soekarno - Hatta, Tangerang" ]);
            
            addRow([ null, jabatanManajer, ...nullArr(10), jabatanAsisten ]);
    
            for (let i = 0; i < 5; i++) addRow([]);
            
            addRow([ null, namaManajer, ...nullArr(10), namaAsisten ]);
            currentRow().font = { bold: true, underline: true, size: 11 };
    
            addRow([" "])
        }
    
        function addRow(row) {
            ws.addRow(row);
            rowCount ++;
        }
    
        function nullArr(length) {
            return Array(length).fill(null);
        }
    
        function currentRow() {
            return ws.getRow(rowCount);
        }
    }
}

function gefXLDownload(data) {
    const wb = new GefExcel();
    const sheetPeralatanAddress = wb.addSheetPeralatan(data);
    wb.addSheetPersonil(data);
    wb.addSheetKegiatan(data);
    wb.addSheetGrafik(data, sheetPeralatanAddress);
    wb.saveAs(`xLap. Bul. ${data.bulanTahun} GEFNT.xlsx`);
}

gefXLDownload(mockdata);
