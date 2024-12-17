 // Show file input and enable convert button when the user selects a file
 document.getElementById('fileLabel').addEventListener('click', function() {
    document.getElementById('fileInput').click();
    document.getElementById('statusText').textContent = "Select a file to import.";
  });
  
  document.getElementById('convertButton').addEventListener('click', function() {
    document.getElementById('fileInput').click();
  });
  