
document.getElementById('fileInput').addEventListener('change', function (event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        // Read the first sheet
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        // Convert sheet data to JSON
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        // Display JSON data in the console or on the page
        console.log(jsonData);
        document.getElementById('output').innerHTML = '<pre>' + JSON.stringify(jsonData, null, 2) + '</pre>';
    };

    reader.readAsArrayBuffer(file);
});

