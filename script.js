const exportBtn = document.querySelector(".export");
const excelInput = document.getElementById("excelInput");
const pdfOutput = document.getElementById("pdfOutput");
const title = document.querySelector("h1");

title.addEventListener("click", () => {
    // Reload the page if user clicks on the title
    window.location.reload();
    localStorage.setItem("messageShowed", "true");
});

exportBtn.addEventListener("click", async () => {
    exportBtn.classList.add("button-press");

    if (!window.pdfTemplate || !window.excelData) {
        alert('Please load both an Excel file and a PDF first');
        return;
    }

    const checkmark = 'X';

    try {
        const originalPdfDoc = window.pdfTemplate;

        for (let i = 0; i < window.excelData.length; i++) {
            const row = window.excelData[i];
            const pdfCopy = await PDFLib.PDFDocument.load(await originalPdfDoc.save());
            const pages = pdfCopy.getPages();
            const page1 = pages[0];
            const page2 = pages.length > 1 ? pages[1] : null;
            const page4 = pages.length > 3 ? pages[3] : null;
            const page5 = pages.length > 4 ? pages[4] : null;
            const font = await pdfCopy.embedFont(PDFLib.StandardFonts.HelveticaBold);
            let fontSize = 10;

            let baseYear = null;

            for (let j = 0; j < 20; j++) {
                let value = row[j] ? String(row[j]) : '';
                let x = 50 + (j % 3) * 150;
                let y = 700 - Math.floor(j / 3) * 30;
                let page = j < 11 ? page1 : page2;

                // Handling specific values for each index j
                if (j === 0) { // SIN
                    x = 145;
                    y = 695;
                } else if (j === 1) { // First Name
                    x = 135;
                    y = 660;
                } else if (j === 2) { // Last Name
                    x = 135;
                    y = 645;
                } else if (j === 3) { // DOB
                    x = 135;
                    y = 630;
                } else if (j === 4) { // Empty
                    x = 120;
                    y = 580;
                } else if (j === 5) { // Phone
                    x = 130;
                    y = 590;
                } else if (j === 6) { // Apartement
                    x = 200;
                    y = 530;
                } else if (j === 7) { // Address
                    x = 210;
                    y = 530;
                } else if (j === 8) { // City
                    x = 210;
                    y = 510;
                } else if (j === 9) { // Province
                    x = 210;
                    y = 497;
                } else if (j === 10) { // Postal Code
                    x = 205;
                    y = 482;
                } else if (j === 11 && page2) { // Married or Single
                    if (value.toLowerCase() === "single") {
                        x = 88;
                        y = 488;
                        value = checkmark;
                    } else if (value.toLowerCase() === "married") {
                        x = 95;
                        y = 730;
                        value = checkmark;
                    }
                } else if (j === 12) { // Spouse SIN
                    x = 130;
                    y = 393;
                } else if (j === 13) { // First Name
                    x = 50;
                    y = 362;
                } else if (j === 14) { // Last Name
                    x = 50;
                    y = 345;
                } else if (j === 15) { // Spouse DOB
                    x = 65;
                    y = 330;
                } else if (j === 16) { // Date
                    fontSize = 10;
                    x = 330;
                    y = 212;

                    if (value.includes("-")) {
                        const yearPart = value.split("-")[0];
                        baseYear = parseInt(yearPart);
                    }
                } else if (j === 17 && baseYear && page4) {
                    value = String(baseYear);
                    x = 227;
                    y = 270;
                    page = page4;
                } else if (j === 18 && baseYear && page5) {
                    value = String(baseYear - 1);
                    x = 405;
                    y = 735;
                    page = page5;
                } else if (j === 19 && baseYear && page5) {
                    value = String(baseYear - 2);
                    x = 392;
                    y = 576;
                    page = page5;
                }

                // Draw text at the specific position
                if (page && value) {
                    page.drawText(value, {
                        x,
                        y,
                        size: fontSize,
                        font,
                        color: PDFLib.rgb(0, 0, 0),
                        
                    });
                }
            }

            // Save the filled PDF after all entries are added
            const pdfBytes = await pdfCopy.save();
            const blob = new Blob([pdfBytes], { type: 'application/pdf' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `${row[1]}${row[2]}.pdf`;
            document.body.appendChild(link); // ensures it's added to DOM
            link.click();
            document.body.removeChild(link); // cleanup

            await new Promise(resolve => setTimeout(resolve, 250)); // wait a bit for reliable download
        }
    } catch (error) {
        console.error('Error exporting PDF:', error);
    }

    setTimeout(() => {
        exportBtn.classList.remove("button-press");
    }, 5);
});




excelInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (file && file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
        try {
            const fileReader = new FileReader();
            fileReader.onerror = function(event) {
                console.error("FileReader error:", event.target.error);
            };

            fileReader.onload = async function () {
                try {
                    const data = new Uint8Array(this.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[sheetName];
                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    const parsedData = [];
                    for (let i = 1; i < rows.length; i++) {
                        const row = rows[i];
                    
                        // Skip completely empty rows
                        if (!row || row.every(cell => cell == null || cell === '')) continue;
                    
                        const limitedRow = [];
                        for (let j = 0; j < 17; j++) {
                            let cell = row[j];
                    
                            // If column index 3 and it's a number (likely an Excel date serial), convert it
                            if ((j === 3 || j === 15 || j == 16) && typeof cell === 'number') {
                                // Excel's date system starts from 1900-01-01, but it incorrectly assumes 1900 is a leap year.
                                const excelEpoch = new Date(1899, 11, 30); // Corrected for Excel's bug
                                const date = new Date(excelEpoch.getTime() + cell * 86400000);
                    
                                const yyyy = date.getFullYear();
                                const mm = String(date.getMonth() + 1).padStart(2, '0');
                                const dd = String(date.getDate()).padStart(2, '0');
                                cell = `${yyyy}-${mm}-${dd}`;
                            }
                    
                            limitedRow.push((cell == null || cell === '') ? '' : String(cell));
                        }
                    
                        parsedData.push(limitedRow);
                    }                    


                    window.excelData = parsedData;
                    alert('Excel file loaded successfully!');
                } catch (error) {
                    console.error('Error processing Excel file:', error);
                    alert('Error processing Excel file');
                }
            };

            fileReader.readAsArrayBuffer(file);
        } catch (error) {
            console.error('Error reading file:', error);
            alert('Error reading Excel file');
        }
    } else {
        alert('Please select a valid Excel file (.xlsx)');
    }
});



pdfOutput.addEventListener('change', async (e) => {
    const pdfFile = e.target.files[0];
    if (pdfFile && pdfFile.type === 'application/pdf') {
        try {
            const pdfData = await pdfFile.arrayBuffer();
            // Loading the PDF document
            const pdfDoc = await PDFLib.PDFDocument.load(pdfData, {
                ignoreEncryption: true
            });

            // Checking the page count to see if the pdf is valid in terms of pages
            const pageCount = pdfDoc.getPageCount();
            if (pageCount === 0) {
                throw new Error('The PDF does not contain any pages.');
            }
                        
            // Store the pdf document for later use
            window.pdfTemplate = pdfDoc;

            alert('PDF loaded successfully!');
        } catch (error) {
            console.error('Error loading PDF:', error);
            alert("Error loading PDF");
        }
    } else {
        alert('Please select a valid PDF file');
    }
});
