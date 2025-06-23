let coilData = [];
let loadedFileName = '';

// Use cache-busting version of GitHub Excel file
const GITHUB_EXCEL_URL = 'https://raw.githubusercontent.com/gmanand23/Coil_Info/main/coil-data.xlsx?' + new Date().getTime();

function clearLocalStorage() {
  try {
    localStorage.removeItem('coilData');
    localStorage.removeItem('loadedFileName');
    alert('Local storage cleared successfully!');
    console.log('Local storage cleared.');
    fetchAndLoadExcelFromUrl(GITHUB_EXCEL_URL);
  } catch (e) {
    console.error('Error clearing local storage:', e);
    alert('Failed to clear local storage. Check console for details.');
  }
}

// Updated resetApp function to clear all local storage and force a full reload
function resetApp() {
  try {
    localStorage.clear(); // Clear all local storage for a complete reset
    alert('App data cleared. Reloading from GitHub for a fresh install...');
    // Force a complete reload by adding a unique query parameter to bypass cache
    window.location.href = window.location.origin + window.location.pathname + '?cachebust=' + new Date().getTime();
  } catch (e) {
    console.error('Reset failed:', e);
    alert('Failed to reset. Check console.');
  }
}

async function fetchAndLoadExcelFromUrl(url) {
  const downloadButton = document.querySelector('button[onclick="downloadExcel()"]');
  if (downloadButton) {
    downloadButton.disabled = true;
    downloadButton.textContent = 'Loading Data...';
  }

  try {
    document.getElementById('fileName').textContent = `Loading data from GitHub...`;
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

    const arrayBuffer = await response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    coilData = XLSX.utils.sheet_to_json(sheet);

    const urlParts = url.split('/');
    loadedFileName = urlParts[urlParts.length - 1].split('?')[0] || 'coil-data.xlsx';

    // Save to local storage
    localStorage.setItem('coilData', JSON.stringify(coilData));
    localStorage.setItem('loadedFileName', loadedFileName);

    document.getElementById('fileName').textContent = `Loaded File (from GitHub): ${loadedFileName}`;
    alert('Excel loaded successfully from GitHub!');

    if (downloadButton) {
      downloadButton.disabled = false;
      downloadButton.textContent = 'Download Loaded Excel';
    }

  } catch (error) {
    console.error('Error loading Excel from URL:', error);
    document.getElementById('fileName').textContent = `Failed to load from GitHub.`;
    alert('Failed to load Excel from GitHub. Check console for details.');
    coilData = [];
    if (downloadButton) {
      downloadButton.disabled = true;
      downloadButton.textContent = 'Failed to Load Data';
    }
  }
}

document.getElementById('excelFile').addEventListener('change', (e) => {
  coilData = [];
  document.getElementById('result').innerHTML = '';
  document.getElementById('coilInput').value = '';
  document.getElementById('suggestions').style.display = 'none';

  const downloadButton = document.querySelector('button[onclick="downloadExcel()"]');
  if (downloadButton) {
    downloadButton.disabled = true;
    downloadButton.textContent = 'Loading Local File...';
  }

  const file = e.target.files[0];
  if (file) {
    loadedFileName = file.name;
    document.getElementById('fileName').textContent = `Loaded File (from local upload): ${loadedFileName}`;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      coilData = XLSX.utils.sheet_to_json(sheet);
      alert('Excel loaded successfully from local file!');

      localStorage.setItem('coilData', JSON.stringify(coilData));
      localStorage.setItem('loadedFileName', loadedFileName);

      if (downloadButton) {
        downloadButton.disabled = false;
        downloadButton.textContent = 'Download Loaded Excel';
      }
    };
    reader.onerror = (evt) => {
      console.error('Error reading local file:', evt);
      alert('Error reading local Excel file.');
      if (downloadButton) {
        downloadButton.disabled = true;
        downloadButton.textContent = 'Failed to Load Data';
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    if (downloadButton) {
      downloadButton.disabled = true;
      downloadButton.textContent = 'Download Loaded Excel';
    }
  }
});

function downloadExcel() {
  if (coilData.length === 0) {
    alert('No Excel data to download.');
    return;
  }

  if (!loadedFileName) {
    loadedFileName = 'downloaded_coil_data.xlsx';
  }

  const worksheet = XLSX.utils.json_to_sheet(coilData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Coil Data");
  XLSX.writeFile(workbook, loadedFileName);
}

document.addEventListener('DOMContentLoaded', async () => {
  // Check if there's cached data, otherwise load from GitHub
  const cachedData = localStorage.getItem('coilData');
  const cachedFileName = localStorage.getItem('loadedFileName');

  if (cachedData && cachedFileName) {
    coilData = JSON.parse(cachedData);
    loadedFileName = cachedFileName;
    document.getElementById('fileName').textContent = `Loaded from Local Storage: ${loadedFileName}`;
  } else {
    await fetchAndLoadExcelFromUrl(GITHUB_EXCEL_URL);
  }
});

function searchCoil() {
  const coilNumber = document.getElementById('coilInput').value.trim().toUpperCase();
  const result = coilData.find(row => {
    const keys = Object.keys(row);
    // Find the key that matches 'MILL COIL NO' case-insensitively and trim-safely
    const matchingKey = keys.find(k => k.trim().toUpperCase() === 'MILL COIL NO');
    if (!matchingKey) return false;
    const sheetCoil = String(row[matchingKey]).trim().toUpperCase();
    return sheetCoil === coilNumber;
  });
  displayResult(result);
}

function displayResult(data) {
  const resultDiv = document.getElementById('result');
  if (data) {
    let tableHTML = '<table style="font-family: Comic Sans MS; width:100%; border-collapse: collapse;">';
    tableHTML += '<thead><tr><th style="border: 1px solid #fff; padding: 8px; color: white;">Field</th><th style="border: 1px solid #fff; padding: 8px; color: white;">Value</th></tr></thead><tbody>';
    for (const [key, val] of Object.entries(data)) {
      tableHTML += `<tr><td style="border: 1px solid #fff; padding: 8px; color: white;">${key.trim()}</td><td style="border: 1px solid #fff; padding: 8px; color: white;">${val}</td></tr>`;
    }
    tableHTML += '</tbody></table>';
    resultDiv.innerHTML = tableHTML;
  } else {
    resultDiv.innerHTML = '<p>Coil number not found.</p>';
  }
}

let qrReader = null; // Declare globally for start/close control

function startScanner() {
  if (qrReader) closeScanner(); // Close any existing scanner before starting a new one
  const readerDiv = document.getElementById("reader");
  readerDiv.style.display = 'block'; // Make sure the reader element is visible
  readerDiv.innerHTML = ''; // Clear previous content

  qrReader = new Html5Qrcode("reader", { verbose: false });
  qrReader.start(
    { facingMode: "environment" }, // Prefer rear camera
    { fps: 10, qrbox: { width: 250, height: 250 } }, // Configuration for QR scanning
    (decodedText) => { // onScanSuccess
      document.getElementById('coilInput').value = decodedText;
      toggleClearButton(); // Ensure clear button shows if text is scanned
      searchCoil();
      closeScanner();
      scrollToResult(); // Scroll to the result after scan and search
    },
    (errorMessage) => { // onScanFailure
      console.warn(`QR error: ${errorMessage}`);
    }
  );
}

function closeScanner() {
  if (qrReader) {
    qrReader.stop().then(() => {
      qrReader.clear(); // Clear the UI of the QR code scanner
      document.getElementById("reader").innerHTML = ''; // Clear the div content
      document.getElementById("reader").style.display = 'none'; // Hide the reader div
      qrReader = null;
    }).catch(err => {
      console.error("Error stopping scanner", err);
      document.getElementById("reader").style.display = 'none'; // Ensure hidden even on error
      qrReader = null;
    });
  } else {
    document.getElementById("reader").style.display = 'none'; // Ensure hidden if qrReader is null
  }
}

function showSuggestions() {
  const input = document.getElementById('coilInput').value.trim().toUpperCase();
  const suggestionsDiv = document.getElementById('suggestions');
  suggestionsDiv.innerHTML = '';

  if (!input || coilData.length === 0) {
    suggestionsDiv.style.display = 'none';
    return;
  }

  // Find the 'MILL COIL NO' key dynamically, case-insensitively and trim-safely
  const keys = Object.keys(coilData[0] || {}); // Handle empty coilData gracefully
  const matchingKey = keys.find(k => k.trim().toUpperCase() === 'MILL COIL NO');
  if (!matchingKey) {
    suggestionsDiv.style.display = 'none';
    return;
  }

  const suggestions = coilData
    .map(row => String(row[matchingKey]).trim().toUpperCase())
    .filter(coil => coil.includes(input))
    .slice(0, 10); // Limit to 10 suggestions

  if (suggestions.length > 0) {
    suggestions.forEach(s => {
      const div = document.createElement('div');
      div.textContent = s;
      div.onclick = () => {
        document.getElementById('coilInput').value = s;
        suggestionsDiv.style.display = 'none';
        searchCoil();
        toggleClearButton(); // Ensure clear button shows when a suggestion is clicked
      };
      suggestionsDiv.appendChild(div);
    });
    suggestionsDiv.style.display = 'block';
  } else {
    suggestionsDiv.style.display = 'none';
  }
}

function toggleClearButton() {
  const input = document.getElementById('coilInput');
  const clearBtn = document.getElementById('clearInputBtn');
  if (clearBtn) {
    clearBtn.style.display = input.value.trim() !== '' ? 'block' : 'none';
  }
}

function clearInput() {
  document.getElementById('coilInput').value = '';
  document.getElementById('result').innerHTML = '';
  document.getElementById('suggestions').innerHTML = ''; // Clear suggestions content too
  document.getElementById('suggestions').style.display = 'none';
  const clearBtn = document.getElementById('clearInputBtn');
  if (clearBtn) clearBtn.style.display = 'none';
}

function scrollToResult() {
  const resultDiv = document.getElementById('result');
  if (resultDiv) {
    resultDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

// Event listener for input field to trigger search on Enter key
document.getElementById('coilInput').addEventListener('keypress', function(event) {
    if (event.key === 'Enter') {
        searchCoil();
    }
});

// Close suggestions when clicking outside
document.addEventListener('click', function(event) {
    const suggestionsDiv = document.getElementById('suggestions');
    const coilInput = document.getElementById('coilInput');
    // Check if the click is outside the suggestions div AND not on the input field itself
    if (!suggestionsDiv.contains(event.target) && event.target !== coilInput) {
        suggestionsDiv.style.display = 'none';
    }
});