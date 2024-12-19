// // Helper function to convert time from 12-hour to 24-hour format
// function convertTo24Hour(timeStr) {
//     timeStr = timeStr.replace("a.m.", "AM").replace("p.m.", "PM");
//     let date = new Date("01/01/2000 " + timeStr); // Arbitrary date, only time matters
//     return date.toTimeString().slice(0, 5); // Return HH:MM format
//   }
  
//   // Extract time range (start and end times) from a string like "10:30 a.m. - 12:00 p.m."
//   function extractTime(text) {
//     const timeRange = text.match(/(\d{1,2}:\d{2} [ap\.m]+) - (\d{1,2}:\d{2} [ap\.m]+)/);
//     if (timeRange) {
//       const startTime = convertTo24Hour(timeRange[1]);
//       const endTime = convertTo24Hour(timeRange[2]);
//       const [startHour, startMinute] = startTime.split(":").map(Number);
//       const [endHour, endMinute] = endTime.split(":").map(Number);
//       return { startHour, startMinute, endHour, endMinute };
//     } else {
//       return null;
//     }
//   }
  
//   // Extract location from the info text
//   function extractLocation(text) {
//     const parts = text.replace("\n", " ").split(" | ");
//     return parts.length >= 4 ? parts[3].trim() : null;
//   }
  
//   // Handle the file input and process the Excel file
//   document.getElementById('fileInput').addEventListener('change', function(event) {
//     const file = event.target.files[0];
//     if (!file) {
//       console.log("No file selected");
//       document.getElementById('statusText').textContent = "No file selected.";
//       return;
//     }
  
//     console.log("File selected:", file.name);
//     const reader = new FileReader();
    
//     reader.onload = function(e) {
//       try {
//         const data = e.target.result;
//         const wb = XLSX.read(data, { type: 'binary' });
//         const sheet = wb.Sheets[wb.SheetNames[0]];
        
//         // Verify if the sheet exists and log details for debugging
//         if (!sheet) {
//           throw new Error("No sheet found in the Excel file.");
//         }
  
//         console.log("Sheet data:", sheet);
//         const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
//         let icsData = "BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\n";
  
//         // Process each row in the Excel sheet (starting from row 4 in the original code)
//         for (let i = 3; i < rows.length; i++) {
//           const row = rows[i];
//           const section = row[4];
//           const info = row[7];
//           const start = row[10];
//           const end = row[11];
  
//           if (section && info && start && end) {
//             let { startHour, startMinute, endHour, endMinute } = extractTime(info);
//             if (startHour !== undefined && endHour !== undefined) {
//               let startDate = new Date(start.year, start.month - 1, start.day, startHour, startMinute);
//               let endDate = new Date(end.year, end.month - 1, end.day, endHour, endMinute);
  
//               icsData += `BEGIN:VEVENT\nSUMMARY:${section}\n`;
//               icsData += `DTSTART;TZID=America/Vancouver:${startDate.toISOString().replace(/[-:]/g, "").split(".")[0]}\n`;
//               icsData += `DTEND;TZID=America/Vancouver:${endDate.toISOString().replace(/[-:]/g, "").split(".")[0]}\n`;
//               icsData += `DESCRIPTION:${extractLocation(info)}\n`;
  
//               // Example recurrence rule, adjust as needed
//               let daysOfWeek = info.split('|')[1]?.trim().split(' ').map(day => day.slice(0, 2).toUpperCase()).join(',');
//               icsData += `RRULE:FREQ=WEEKLY;INTERVAL=1;BYDAY=${daysOfWeek}\n`;
//               icsData += "END:VEVENT\n";
//             }
//           }
//         }
  
//         icsData += "END:VCALENDAR";
  
//         // Trigger file download
//         const blob = new Blob([icsData], { type: 'text/calendar' });
//         const link = document.createElement('a');
//         link.href = URL.createObjectURL(blob);
//         link.download = 'schedule.ics';
//         link.click();
  
//         document.getElementById('statusText').textContent = "Done - File downloaded as 'schedule.ics'";
  
//       } catch (error) {
//         // Detailed error logging
//         console.error("Error processing the file:", error);
//         document.getElementById('statusText').textContent = "An error occurred while processing the file. Check console for details.";
//       }
//     };
  
//     // Check if the file type is valid (XLSX)
//     if (!file.name.endsWith('.xlsx')) {
//       console.error("Invalid file type. Please select an XLSX file.");
//       document.getElementById('statusText').textContent = "Invalid file type. Please select an XLSX file.";
//       return;
//     }
  
//     // Read the file as binary string
//     reader.readAsBinaryString(file);
//   });
  
//   // Show file input and enable convert button when the user selects a file
//   document.getElementById('fileLabel').addEventListener('click', function() {
//     document.getElementById('fileInput').click();
//     document.getElementById('statusText').textContent = "Select a file to import.";
//   });
  
//   document.getElementById('convertButton').addEventListener('click', function() {
//     document.getElementById('fileInput').click();
//   });
  