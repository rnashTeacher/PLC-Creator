// Global variable to hold the parsed sheet data
let sheetData = [];
// Global variable to hold the logo image data
let logoData = null;
// Global variable to hold the entire workbook
let loadedWorkbook = null;
// Global variables for interactive cell/column selection
let selectionMode = null; // Can be 'topicHeaderRow', 'firstStudent', 'guidanceRow', 'maxScoreRow', 'gradeCol', 'firstTopicCol', 'lastTopicCol'
let firstStudentRowIndex = null;
let gradeColIndex = null;
let maxScoreRowIndex = null;
let guidanceRowIndex = null;
let topicHeaderRowIndex = null;
let firstTopicColIndex = null;
let lastTopicColIndex = null;

document.getElementById('fileInput').addEventListener('change', handleFileSelect);
document.getElementById('generatePdfButton').addEventListener('click', generatePdfs);
document.getElementById('logoInput').addEventListener('change', handleLogoSelect);
document.getElementById('sheetSelector').addEventListener('change', displaySelectedSheet);
document.getElementById('setTopicHeaderBtn').addEventListener('click', () => enterSelectionMode('topicHeaderRow'));
document.getElementById('setGuidanceRowBtn').addEventListener('click', () => enterSelectionMode('guidanceRow'));
document.getElementById('setMaxScoreRowBtn').addEventListener('click', () => enterSelectionMode('maxScoreRow'));
document.getElementById('setFirstStudentBtn').addEventListener('click', () => enterSelectionMode('firstStudent'));
document.getElementById('setGradeColBtn').addEventListener('click', () => enterSelectionMode('gradeCol'));
document.getElementById('setFirstTopicBtn').addEventListener('click', () => enterSelectionMode('firstTopicCol'));
document.getElementById('setLastTopicBtn').addEventListener('click', () => enterSelectionMode('lastTopicCol'));

const dataContainer = document.getElementById('data-container');
dataContainer.addEventListener('click', handleTableClick);

function handleFileSelect(event) {
    const file = event.target.files[0];
    const errorContainer = document.getElementById('error-message');
    const loader = document.getElementById('loader');
    const outputControls = document.getElementById('output-controls');
    const sheetSelectorContainer = document.getElementById('sheet-selector-container');
    const cellConfigContainer = document.getElementById('cell-config-container');
    const sheetSelector = document.getElementById('sheetSelector');

    // Clear previous results
    dataContainer.innerHTML = '';
    errorContainer.innerHTML = '';
    outputControls.style.display = 'none';
    sheetData = []; // Clear old data
    cellConfigContainer.style.display = 'none';
    firstStudentRowIndex = null;
    gradeColIndex = null;
    maxScoreRowIndex = null;
    guidanceRowIndex = null;
    topicHeaderRowIndex = null;
    firstTopicColIndex = null;
    lastTopicColIndex = null;
    updateSelectionDisplay();


    sheetSelectorContainer.style.display = 'none';
    sheetSelector.innerHTML = '';

    if (!file) {
        return;
    }

    loader.style.display = 'block';

    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            loadedWorkbook = XLSX.read(data, { type: 'array' });

            // Populate the sheet selector dropdown
            loadedWorkbook.SheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelector.appendChild(option);
            });
            sheetSelectorContainer.style.display = 'block';
            displaySelectedSheet(); // Automatically display the first sheet
        } catch (error) {
            console.error('Error processing file:', error);
            errorContainer.textContent = `Failed to process file: ${error.message}. The file might be corrupted or in an unsupported format.`;
        } finally {
            loader.style.display = 'none';
        }
    };

    reader.readAsArrayBuffer(file);
}

function displaySelectedSheet() {
    const sheetSelector = document.getElementById('sheetSelector');
    const selectedSheetName = sheetSelector.value;
    const errorContainer = document.getElementById('error-message');
    const outputControls = document.getElementById('output-controls');
    const cellConfigContainer = document.getElementById('cell-config-container');

    // Clear previous table and errors
    dataContainer.innerHTML = '';
    errorContainer.innerHTML = '';
    outputControls.style.display = 'none';
    cellConfigContainer.style.display = 'none';
    firstStudentRowIndex = null;
    gradeColIndex = null;
    maxScoreRowIndex = null;
    guidanceRowIndex = null;
    topicHeaderRowIndex = null;
    firstTopicColIndex = null;
    lastTopicColIndex = null;
    updateSelectionDisplay();

    if (!loadedWorkbook || !selectedSheetName) {
        return;
    }

    const worksheet = loadedWorkbook.Sheets[selectedSheetName];
    const json_data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (json_data && json_data.length > 0) {
        sheetData = json_data; // Store data for PDF generation
        renderTable(json_data);
        cellConfigContainer.style.display = 'block';
        outputControls.style.display = 'block'; // Show the generate button
    } else {
        errorContainer.textContent = `No data found in the sheet "${selectedSheetName}".`;
    }
}

function handleLogoSelect(event) {
    const file = event.target.files[0];
    const errorContainer = document.getElementById('error-message');
    errorContainer.innerHTML = ''; // Clear previous errors

    if (!file) {
        logoData = null;
        return;
    }
    const reader = new FileReader();
    reader.onload = function(e) {
        logoData = e.target.result; // This is the Base64 data URL
    };
    reader.readAsDataURL(file);
}

function enterSelectionMode(mode) {
    selectionMode = mode;
    dataContainer.style.cursor = 'crosshair';
    let message = "Selection mode activated. ";
    switch (mode) {
        case 'topicHeaderRow':
            message += "Click on any cell in the row containing the topic headers (e.g., 'Q1 - 1.1...').";
            break;
        case 'firstStudent':
            message += "Click on any cell in the first student's row.";
            break;
        case 'guidanceRow':
            message += "Click on any cell in the row containing the guidance text/links for QR codes.";
            break;
        case 'gradeCol':
            message += "Click on the grade for the first student to identify the column.";
            break;
        case 'maxScoreRow':
            message += "Click on any cell in the row containing the maximum scores for each topic.";
            break;
        case 'firstTopicCol':
            message += "Click on the score for the first topic to identify the starting column.";
            break;
        case 'lastTopicCol':
            message += "Click on the score for the last topic to identify the ending column.";
            break;
    }
    alert(message);
}

function handleTableClick(event) {
    if (!selectionMode) return;

    const target = event.target.closest('td');
    if (!target) return;

    const rowIndex = target.parentElement.dataset.rowIndex;
    const colIndex = target.dataset.colIndex;

    if (rowIndex === undefined || colIndex === undefined) return;

    if (selectionMode === 'maxScoreRow') {
        maxScoreRowIndex = parseInt(rowIndex);
    } else if (selectionMode === 'topicHeaderRow') {
        topicHeaderRowIndex = parseInt(rowIndex);
    } else if (selectionMode === 'guidanceRow') {
        guidanceRowIndex = parseInt(rowIndex);
    } else if (selectionMode === 'firstStudent') {
        firstStudentRowIndex = parseInt(rowIndex);
    } else if (selectionMode === 'gradeCol') {
        gradeColIndex = parseInt(colIndex);
    } else if (selectionMode === 'firstTopicCol') {
        firstTopicColIndex = parseInt(colIndex);
    } else if (selectionMode === 'lastTopicCol') {
        lastTopicColIndex = parseInt(colIndex);
    }

    updateSelectionDisplay();
    selectionMode = null;
    dataContainer.style.cursor = 'default';
}

function updateSelectionDisplay() {
    const topicHeaderRowDisplay = document.getElementById('topicHeaderRowDisplay');
    if (topicHeaderRowIndex !== null) {
        topicHeaderRowDisplay.textContent = `Row ${topicHeaderRowIndex + 1}`;
    } else {
        topicHeaderRowDisplay.textContent = 'Not Set';
    }

    const firstStudentDisplay = document.getElementById('firstStudentRowDisplay');
    if (firstStudentRowIndex !== null) {
        firstStudentDisplay.textContent = `Row ${firstStudentRowIndex + 1}`;
    } else {
        firstStudentDisplay.textContent = 'Not Set';
    }

    const maxScoreRowDisplay = document.getElementById('maxScoreRowDisplay');
    if (maxScoreRowIndex !== null) {
        maxScoreRowDisplay.textContent = `Row ${maxScoreRowIndex + 1}`;
    } else {
        maxScoreRowDisplay.textContent = 'Not Set';
    }

    const guidanceRowDisplay = document.getElementById('guidanceRowDisplay');
    if (guidanceRowIndex !== null) {
        guidanceRowDisplay.textContent = `Row ${guidanceRowIndex + 1}`;
    } else {
        guidanceRowDisplay.textContent = 'Not Set';
    }

    const gradeColDisplay = document.getElementById('gradeColDisplay');
    if (gradeColIndex !== null) {
        gradeColDisplay.textContent = `Column ${getCellAddress(0, gradeColIndex).replace(/[0-9]/g, '')}`;
    } else {
        gradeColDisplay.textContent = 'Not Set';
    }

    const firstTopicColDisplay = document.getElementById('firstTopicColDisplay');
    if (firstTopicColIndex !== null) {
        firstTopicColDisplay.textContent = `Column ${getCellAddress(0, firstTopicColIndex).replace(/[0-9]/g, '')}`;
    } else {
        firstTopicColDisplay.textContent = 'Not Set';
    }

    const lastTopicColDisplay = document.getElementById('lastTopicColDisplay');
    if (lastTopicColIndex !== null) {
        lastTopicColDisplay.textContent = `Column ${getCellAddress(0, lastTopicColIndex).replace(/[0-9]/g, '')}`;
    } else {
        lastTopicColDisplay.textContent = 'Not Set';
    }
}

function getCellAddress(rowIndex, colIndex) {
    let colName = '';
    let num = colIndex;
    while (num >= 0) {
        colName = String.fromCharCode(65 + (num % 26)) + colName;
        num = Math.floor(num / 26) - 1;
    }
    return `${colName}${rowIndex + 1}`;
}

function generatePdfs() {
    if (topicHeaderRowIndex === null) {
        alert("PDF Generation Failed: Please select the 'Topic Header Row' before generating PDFs.");
        return;
    }

    if (firstStudentRowIndex === null) {
        alert("PDF Generation Failed: Please select the 'First Student Row' before generating PDFs.");
        return;
    }

    if (guidanceRowIndex === null) {
        alert("PDF Generation Failed: Please select the 'Guidance Row' before generating PDFs.");
        return;
    }

    if (maxScoreRowIndex === null) {
        alert("PDF Generation Failed: Please select the 'Max Score Row' before generating PDFs.");
        return;
    }

    if (gradeColIndex === null) {
        alert("PDF Generation Failed: Please select the 'Grade Column' before generating PDFs.");
        return;
    }

    if (firstTopicColIndex === null || lastTopicColIndex === null) {
        alert("PDF Generation Failed: Please select both the 'First Topic Column' and 'Last Topic Column'.");
        return;
    }

    if (lastTopicColIndex < firstTopicColIndex) {
        alert("PDF Generation Failed: The 'Last Topic Column' must be to the right of the 'First Topic Column'.");
        return;
    }

    // Get the jsPDF constructor from the window object
    const { jsPDF } = window.jspdf;

    // --- PDF HEADER CONFIGURATION ---
    const PDF_HEADER_TEXT = ""; // <-- CHANGE THIS TEXT

    // Get student names from cell A3 downwards (index 2 of the array)
    const students = sheetData.slice(firstStudentRowIndex);

    students.forEach(rowData => {
        const studentName = rowData[0]; // Name is in the first column (A) 
        const grade = rowData[gradeColIndex]; // Grade from the selected column
        if (studentName && studentName.trim() !== '') {
            // Create a new PDF document
            const doc = new jsPDF();

            const headerTextX = logoData ? 40 : 20; // Move text right if logo exists

            // 1. Add logo if it has been selected
            if (logoData) {
                // addImage(imageData, format, x, y, width, height)
                const img = new Image();
                img.src = logoData;
                const imgProps = doc.getImageProperties(logoData);
                
                // Define a max width for the logo and calculate height to maintain aspect ratio
                const maxWidth = 190;
                const newHeight = (imgProps.height * maxWidth) / imgProps.width;
                const x_pos = (doc.internal.pageSize.getWidth() - maxWidth) / 2; // Center the logo

                doc.addImage(logoData, 'PNG', x_pos, 15, maxWidth, newHeight);
            }

            // 2. Add the header text
            doc.setFontSize(12);
            doc.text(PDF_HEADER_TEXT, headerTextX, 20);

            // 3. Add the student's name below the header
            doc.setFontSize(16);
            doc.text(studentName, 20, 60);

            doc.setFontSize(16);
            doc.text("Mock Exam grade = " + grade, 20, 68);

            // 4. Process and categorize topic data for the current student
            const topicHeaders = sheetData[topicHeaderRowIndex];
            const maxScores = sheetData[maxScoreRowIndex];
            const veryInsecure = [];
            const insecure = [];
            const secure = [];

            for (let i = firstTopicColIndex; i <= lastTopicColIndex; i++) {
                const maxScore = parseFloat(maxScores[i]);
                const topicName = topicHeaders[i] || '';
                const studentScore = parseFloat(rowData[i]);

                // Ensure we have a valid topic, a valid score for that topic, and a valid max score for that topic
                if (topicName && !isNaN(studentScore) && !isNaN(maxScore) && maxScore > 0) {
                    const secureThreshold = 0.66 * maxScore;
                    const insecureThreshold = 0.33 * maxScore;

                    if (studentScore > secureThreshold) {
                        secure.push(topicName);
                    } else if (studentScore > insecureThreshold) {
                        insecure.push(topicName);
                    } else {
                        veryInsecure.push(topicName);
                    }
                }
            }

            // 5. Build the table body from the categorized topics
            const tableBody = [];
            const maxRows = Math.max(veryInsecure.length, insecure.length, secure.length);
            for (let i = 0; i < maxRows; i++) {
                tableBody.push([
                    veryInsecure[i] || '', // Use topic or empty string if undefined
                    insecure[i] || '',
                    secure[i] || '',
                ]);
            }

            // 6. Add the table using the AutoTable plugin
            const firstTable = doc.autoTable({
                startY: 75, // Start the table below the grade text
                head: [['Very Insecure Topics', 'Insecure Topics', 'Secure Topics']],
                body: tableBody,
                theme: 'grid',
                headStyles: { fillColor: [41, 128, 185] }, // A professional blue color
            });

            // 7. Prepare and add a second table for very insecure topics' guidance from Row 2
            if (veryInsecure.length > 0) {
                const guidanceTableBody = [];
                const guidanceData = sheetData[guidanceRowIndex];

                if (guidanceData) {
                    // Loop through insecure topics two at a time to build the new layout
                    for (let i = 0; i < veryInsecure.length; i += 2) {
                        const newRow = [];

                        // --- First Topic in the Pair ---
                        const topic1Name = veryInsecure[i];
                        const topic1Index = topicHeaders.indexOf(topic1Name);
                        const guidance1Text = (topic1Index !== -1) ? (guidanceData[topic1Index] || 'N/A') : 'N/A';
                        newRow.push(topic1Name, guidance1Text);

                        // --- Second Topic in the Pair (if it exists) ---
                        if (i + 1 < veryInsecure.length) {
                            const topic2Name = veryInsecure[i + 1];
                            const topic2Index = topicHeaders.indexOf(topic2Name);
                            const guidance2Text = (topic2Index !== -1) ? (guidanceData[topic2Index] || 'N/A') : 'N/A';
                            newRow.push(topic2Name, guidance2Text);
                        } else {
                            // If there's an odd number of topics, fill the rest of the row with empty strings
                            newRow.push('', '');
                        }
                        guidanceTableBody.push(newRow);
                    }
                }

                // --- Page Break Logic for Second Table ---
                let secondTableStartY = firstTable.lastAutoTable.finalY + 10;
                const pageHeight = doc.internal.pageSize.getHeight();
                const bottomMargin = 20; // A safe assumption for the page's bottom margin

                // Estimate the height of the second table
                const rowHeight = 35; // Based on minCellHeight
                const headerHeight = 15; // A reasonable estimate for the header
                const estimatedTableHeight = (guidanceTableBody.length * rowHeight) + headerHeight;

                // If the estimated height exceeds the remaining space, add a new page
                if (secondTableStartY + estimatedTableHeight > pageHeight - bottomMargin) {
                    doc.addPage();
                    secondTableStartY = 20; // Start near the top of the new page
                }

                // Calculate column widths manually for better compatibility
                const pageMargins = 20; // A standard margin of 20 points on each side
                const availableWidth = doc.internal.pageSize.getWidth() - (pageMargins * 2);
                const qrColWidth = 35; // Fixed width for QR code columns (QR size is 30)
                const remainingWidth = availableWidth - (qrColWidth * 2);
                const topicColWidth = remainingWidth / 2; // Distribute remaining width to topic columns

                doc.autoTable({
                    startY: secondTableStartY,
                    head: [['Topic', 'Guidance', 'Topic', 'Guidance']],
                    body: guidanceTableBody,
                    theme: 'striped',
                    headStyles: {
                        fillColor: [231, 76, 60], // A red color for 'very insecure'
                        minCellHeight: 12 // Make header shorter than body rows
                    },
                    columnStyles: {
                        0: { cellWidth: topicColWidth },
                        1: { cellWidth: qrColWidth },
                        2: { cellWidth: topicColWidth },
                        3: { cellWidth: qrColWidth }
                    },
                    styles: { minCellHeight: 35 }, // Ensure rows are tall enough for a larger QR code
                    willDrawCell: function(data) {
                        // Check if we are in a guidance column (2nd or 4th) and the cell contains a URL
                        if (data.section === 'body' && (data.column.index === 1 || data.column.index === 3)) {
                            const cellText = data.cell.raw;
                            if (typeof cellText === 'string' && (cellText.startsWith('http://') || cellText.startsWith('https://'))) {
                                data.cell.text = ''; // Clear the text so it's not drawn
                            }
                        }
                    },
                    didDrawCell: function(data) {
                        // Check if we are in a guidance column (2nd or 4th) of the body
                        if (data.section === 'body' && (data.column.index === 1 || data.column.index === 3)) {
                            const cellText = data.cell.raw;
                            // Check if the cell text is a URL
                            if (typeof cellText === 'string' && (cellText.startsWith('http://') || cellText.startsWith('https://'))) {
                                // Create a QR code
                                const qr = qrcode(0, 'L');
                                qr.addData(cellText);
                                qr.make();
                                const qrImgData = qr.createDataURL(4); // Create a data URL for the QR image

                                // Use a fixed size for all QR codes and center it in the cell
                                const qrSize = 30;
                                const qrX = data.cell.x + (data.cell.width - qrSize) / 2;
                                const qrY = data.cell.y + (data.cell.height - qrSize) / 2;

                                doc.addImage(qrImgData, 'PNG', qrX, qrY, qrSize, qrSize);
                            }
                        }
                    }
                });
            }
            
            // 8. Save the PDF
            doc.save(`${studentName}.pdf`);
        }
    });
}

function renderTable(data) {
    const dataContainer = document.getElementById('data-container');
    const table = document.createElement('table');

    // Create table header
    const headers = data[0];
    if (!headers || headers.length === 0) {
        document.getElementById('error-message').textContent = 'Sheet appears to be empty or has no header row.';
        return;
    }

    const headerRow = document.createElement('tr');
    headers.forEach(headerText => {
        const th = document.createElement('th');
        th.textContent = headerText || '';
        headerRow.appendChild(th);
    });
    const thead = document.createElement('thead');
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Create table body
    const tbody = document.createElement('tbody');
    const bodyRows = data.slice(0); // Get all rows to include header row index
    bodyRows.forEach((rowData, rowIndex) => {
        const tr = document.createElement('tr');
        tr.dataset.rowIndex = rowIndex;
        rowData.forEach((cellData, colIndex) => {
            const td = document.createElement('td');
            td.dataset.colIndex = colIndex;
            td.textContent = cellData || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    dataContainer.appendChild(table);
}