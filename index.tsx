import { GoogleGenAI } from "@google/genai";

// Since pdf.js is loaded from a CDN, we need to declare it to TypeScript
declare const pdfjsLib: any;

// Set the worker source for pdf.js
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.3.136/pdf.worker.min.mjs`;

const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

const App = () => {
  // Backing variables for state
  let file: File | null = null;
  let isLoading = false;
  let csvData: string | null = null;
  let error: string | null = null;

  // Forward-declare the main render function
  let render: () => void;
  let renderQueued = false;

  // This function batches multiple synchronous state changes into a single render call
  // using a microtask (Promise.resolve), preventing both unnecessary re-renders and call stack errors.
  const queueRender = () => {
    if (renderQueued) return;
    renderQueued = true;
    Promise.resolve().then(() => {
      render();
      renderQueued = false;
    });
  };

  // A state proxy object. Assigning to its properties updates the UI.
  const state = {
    get file() { return file; },
    set file(value: File | null) {
      if (file === value) return;
      file = value;
      queueRender();
    },
    get isLoading() { return isLoading; },
    set isLoading(value: boolean) {
      if (isLoading === value) return;
      isLoading = value;
      queueRender();
    },
    get csvData() { return csvData; },
    set csvData(value: string | null) {
      if (csvData === value) return;
      csvData = value;
      queueRender();
    },
    get error() { return error; },
    set error(value: string | null) {
      if (error === value) return;
      error = value;
      queueRender();
    }
  };
  
  const handleFileSelect = (selectedFile: File | null) => {
    if (!selectedFile) return;

    if (selectedFile.type !== "application/pdf") {
      state.error = "Invalid file type. Please upload a PDF.";
      state.file = null;
      state.csvData = null;
      return;
    }
    
    state.file = selectedFile;
    state.csvData = null;
    state.error = null;
  };

  const handleConvert = async () => {
    if (!state.file) return;

    state.isLoading = true;
    state.csvData = null;
    state.error = null;

    let fullText = "";

    // Step 1: Parse the PDF and extract text
    try {
      const arrayBuffer = await state.file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
      
      for (let i = 1; i <= pdf.numPages; i++) {
        try {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            if (textContent.items.length > 0) {
                 fullText += textContent.items.map((item: any) => item.str).join(" ") + "\n";
            }
        } catch(pageError) {
            console.warn(`Could not process page ${i}. It might be malformed.`, pageError);
            // Skip problematic page and continue
        }
      }
    } catch (e) {
      console.error("Error parsing PDF:", e);
      state.error = "Could not read the PDF file. It may be corrupted or protected.";
      state.isLoading = false;
      return; 
    }

    // Step 2: Check if any text was extracted
    if (fullText.trim() === "") {
        state.error = "No text could be extracted from this PDF. It might be image-based or empty.";
        state.isLoading = false;
        return;
    }

    // Step 3: Call the AI to convert the text to CSV
    try {
      const prompt = `You are an AI data extraction specialist. Your sole mission is to find and convert any and all tabular data from the provided text into a clean, single CSV output.

**Core Directives:**

1.  **Find All Tables:** Scrutinize the entire text for any data organized in rows and columns.
2.  **Merge Broken Tables:** The text comes from multiple pages. If a table is split across a page break, you MUST intelligently merge it back into one contiguous table.
3.  **Handle Complex Cells:** Cells may contain multiple lines of text. Enclose the entire content of such cells in double quotes ("") to preserve all lines within a single CSV field.
4.  **Separate Distinct Tables:** If you identify multiple, completely separate tables, place a single blank line between them in the CSV output. This is the unique separator.
5.  **Identify Headers:** The first row of a detected table should be considered the header row.
6.  **Ignore Non-Table Text:** You MUST ignore all non-tabular text, such as paragraphs, titles, headers, footers, or any other prose. Your output should only be the tables.
7.  **Strict CSV Formatting:** Adhere strictly to standard CSV rules (comma delimiter, quoting fields with commas/newlines/quotes, escaping double quotes with another double quote).

**Special Cases:**

-   **No Tables Found:** If, after thorough analysis, you find no data that can be structured as a table, your *only* response MUST be the exact string: \`NO_TABLES_FOUND\`. Do not explain or apologize.

**Final Output:**
Your entire response must be *only* the raw CSV data or the 'NO_TABLES_FOUND' string. Do not include introductory text, summaries, or markdown fences like \`\`\`csv.

**Text for Analysis:**
${fullText}`;
      
      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: prompt,
      });

      const rawResult = response.text.trim();
      
      if (rawResult === "NO_TABLES_FOUND") {
        state.error = "The AI could not find any structured tables in the document.";
        state.isLoading = false;
        return;
      }
      
      // The prompt now forbids markdown, but this is a good safeguard.
      const csvResult = rawResult.replace(/```csv\n|```/g, "").trim();
      
      if (!csvResult) {
          throw new Error("AI returned an empty or invalid response.");
      }
      
      state.csvData = csvResult;

    } catch (e) {
      console.error("Error generating CSV from AI:", e);
      state.error = "The AI failed to convert the data. Please check the PDF content and try again.";
    } finally {
      state.isLoading = false;
    }
  };
  
  const handleDownload = () => {
    if (!state.csvData) return;
    const blob = new Blob([state.csvData], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    const originalFileName = state.file?.name.replace(/\.pdf$/i, '') || 'converted';
    link.setAttribute('download', `${originalFileName}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // Assign the main render function
  render = () => {
    const root = document.getElementById('root');
    if (!root) return;

    root.innerHTML = `
      <div class="app-container">
        <header>
          <h1>AI PDF to Excel Converter</h1>
          <p>Upload a PDF, and our AI will convert its tables into a downloadable Excel-compatible CSV file.</p>
        </header>
        
        <div id="drop-zone" role="button" aria-label="File upload zone">
          <input type="file" id="file-input" accept=".pdf" />
          <p>Drag & drop your PDF here, or click to select a file</p>
        </div>

        <div id="file-name" aria-live="polite">${state.file ? state.file.name : ''}</div>

        <div id="status">
          ${state.isLoading ? `
            <div class="loader" role="status" aria-label="Converting...">
              <div class="spinner"></div>
              <span>Analyzing PDF and converting...</span>
            </div>
          ` : ''}
        </div>

        <div id="error-message" role="alert">${state.error || ''}</div>

        <div class="actions">
          ${!state.isLoading ? `
            <button id="convert-btn" class="btn btn-primary" ${!state.file ? 'disabled' : ''}>
              Convert to Excel
            </button>
            ${state.csvData ? `
              <button id="download-btn" class="btn btn-success">
                Download .csv File
              </button>
            ` : ''}
          ` : ''}
        </div>
      </div>
    `;
    
    // Add event listeners after render
    const dropZone = document.getElementById('drop-zone')!;
    const fileInput = document.getElementById('file-input') as HTMLInputElement;

    dropZone.onclick = () => fileInput.click();
    
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
      dropZone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
      });
    });
    
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'));
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'));
    });

    dropZone.addEventListener('drop', (e: DragEvent) => {
        const dt = e.dataTransfer;
        if (dt && dt.files.length) {
            handleFileSelect(dt.files[0]);
        }
    });

    fileInput.onchange = (e) => {
        const target = e.target as HTMLInputElement;
        if (target.files && target.files.length) {
            handleFileSelect(target.files[0]);
        }
    };
    
    const convertBtn = document.getElementById('convert-btn');
    if(convertBtn) convertBtn.onclick = handleConvert;

    const downloadBtn = document.getElementById('download-btn');
    if(downloadBtn) downloadBtn.onclick = handleDownload;
  };
  
  // Initial render
  render();
};

App();