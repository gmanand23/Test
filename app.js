let coilData = [];
let loadedFileName = '';

const GITHUB_EXCEL_URL = 'https://raw.githubusercontent.com/gmanand23/Coil_Info/main/coil-data.xlsx?' + new Date().getTime();

function resetApp() {
  try {
    localStorage.removeItem('coilData');
    localStorage.removeItem('loadedFileName');
    localStorage.setItem('forceReload', 'true'); // ✅ Trigger reload
    alert('App data cleared. Reloading from GitHub...');
    window.location.href = window.location.href; // ✅ Reload that works in APK
  } catch (e) {
    console.error('Reset failed:', e);
    alert('Failed to reset. Check console.');
  }
}

function clearLocalStorage() {
  try {
    localStorage.removeItem('coilData');
    localStorage.removeItem('loadedFileName');
    alert('Local Excel data cleared.');
    fetchAndLoadExcelFromUrl(GITHUB_EXCEL_URL);
  } catch (e) {
    console.error('Error clearing local storage:', e);
    alert('Failed to clear local storage.');
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
    alert('Failed to load Excel from GitHub.');
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

      localStorage.setItem('coilData', JSON.stringify(coilData));
      localStorage.setItem('loadedFileName', loadedFileName);

      alert('Excel loaded successfully from local file!');
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
  const forceReload = localStorage.getItem('forceReload');
  if (forceReload === 'true') {
    localStorage.removeItem('forceReload'); // Clear the flag
    await fetchAndLoadExcelFromUrl(GITHUB_EXCEL_URL);
    return;
  }

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

let html5QrcodeScanner; // Declare globally or in a scope accessible by startScanner/closeScanner

function searchCoil() {
  const input = document.getElementById('coilInput').value.trim().toUpperCase();
  const resultDiv = document.getElementById('result');
  resultDiv.innerHTML = '';

  if (!input) {
    resultDiv.innerHTML = '<p>Please enter a coil number.</p>';
    return;
  }

  const foundCoil = coilData.find(
    (coil) => String(coil['Coil No.']).trim().toUpperCase() === input
  );

  if (foundCoil) {
    displayResult(foundCoil);
  } else {
    resultDiv.innerHTML = `<p>Coil number "${input}" not found.</p>`;
  }
}

function displayResult(coil) {
  const resultDiv = document.getElementById('result');
  let html = '<table>';
  for (const key in coil) {
    // Exclude the Coil No. itself from the detailed display if it's already the search key
    if (key !== 'Coil No.') {
      html += `<tr><td><strong>${key}:</strong></td><td>${coil[key]}</td></tr>`;
    }
  }
  html += '</table>';
  resultDiv.innerHTML = html;
}

function showSuggestions() {
  const input = document.getElementById('coilInput').value.trim().toUpperCase();
  const suggestionsDiv = document.getElementById('suggestions');
  suggestionsDiv.innerHTML = '';

  if (input.length < 2) { // Only show suggestions if at least 2 characters are typed
    suggestionsDiv.style.display = 'none';
    return;
  }

  const filteredCoils = coilData.filter(coil =>
    String(coil['Coil No.']).trim().toUpperCase().startsWith(input)
  );

  if (filteredCoils.length > 0) {
    filteredCoils.forEach(coil => {
      const div = document.createElement('div');
      div.textContent = coil['Coil No.'];
      div.onclick = () => {
        document.getElementById('coilInput').value = coil['Coil No.'];
        suggestionsDiv.style.display = 'none';
        searchCoil(); // Trigger search immediately when a suggestion is clicked
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
  const clearButton = document.getElementById('clearInputBtn');
  if (input.value.length > 0) {
    clearButton.style.display = 'block';
  } else {
    clearButton.style.display = 'none';
  }
}

function clearInput() {
  document.getElementById('coilInput').value = '';
  document.getElementById('result').innerHTML = '';
  document.getElementById('suggestions').innerHTML = '';
  document.getElementById('suggestions').style.display = 'none';
  toggleClearButton(); // Hide the clear button
}


function onScanSuccess(decodedText, decodedResult) {
  console.log(`Code matched = ${decodedText}`, decodedResult);
  document.getElementById('coilInput').value = decodedText;
  searchCoil();
  closeScanner(); // Automatically close scanner after a successful scan
}

function onScanFailure(error) {
  // console.warn(`Code scan error = ${error}`);
}

function startScanner() {
  const readerDiv = document.getElementById('reader');
  readerDiv.style.display = 'block'; // Ensure the reader div is visible

  if (!html5QrcodeScanner) { // Initialize scanner only if it doesn't exist
    html5QrcodeScanner = new Html5QrcodeScanner(
      "reader", { fps: 10, qrbox: { width: 250, height: 250 } }, /* verbose= */ false);
    html5QrcodeScanner.render(onScanSuccess, onScanFailure);
  } else {
    // If scanner already exists, just resume it if it was paused or stopped
    // This part might need more robust handling depending on Html5QrcodeScanner's internal states.
    // For simplicity, re-rendering might be the easiest if it's designed to handle it.
    // However, calling render again on an already rendered scanner might cause issues.
    // A better approach would be to have html5QrcodeScanner.start() and html5QrcodeScanner.stop() methods.
    // Based on library's typical usage, if render was called once, it should manage its state.
    // So, this else block might not be strictly necessary or might need specific library methods.
  }
}

function closeScanner() {
  if (html5QrcodeScanner) {
    html5QrcodeScanner.clear().then(() => {
      console.log("QR Code scanner stopped.");
      document.getElementById('reader').innerHTML = ''; // Clear the reader content
      document.getElementById('reader').style.display = 'none'; // Hide the reader div
    }).catch((err) => {
      console.error("Failed to clear html5QrcodeScanner: ", err);
    });
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
    if (!suggestionsDiv.contains(event.target) && event.target !== coilInput) {
        suggestionsDiv.style.display = 'none';
    }
});