document.getElementById('fetch-sheets').addEventListener('click', async () => {
    const excelUrl = document.getElementById('excel-url').value;
    if (!excelUrl) {
        alert("Please enter a valid Excel file URL.");
        return;
    }

    try {
        const response = await fetch(excelUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

        const sheetListDiv = document.getElementById('sheet-list');
        sheetListDiv.innerHTML = '';

        workbook.SheetNames.forEach(sheetName => {
            const sheetLink = document.createElement('a');
            sheetLink.textContent = sheetName;
            sheetLink.href = `sheet.html?sheetName=${encodeURIComponent(sheetName)}&fileUrl=${encodeURIComponent(excelUrl)}`;
            sheetLink.classList.add('sheet-link');
            sheetListDiv.appendChild(sheetLink);
            sheetListDiv.appendChild(document.createElement('br'));
        });
    } catch (error) {
        console.error("Error loading Excel file:", error);
        alert("Failed to load the Excel file. Please check the URL and try again.");
    }
});
