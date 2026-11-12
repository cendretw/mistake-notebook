// ============================================================================
// 錯題進化論 Mistake Mastery - Core Application Script
// ============================================================================

if (window.pdfjsLib) {
  window.pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

const STORAGE_KEY = 'mistake_mastery_autosave_v1';

// Application State
const state = {
  sheetName: '我的錯題進化本_2026',
  sheetMode: 'new', // 'new' | 'existing'
  originalLoadedCount: 0,
  excelFileHandle: null, // FileSystemFileHandle for real-time local .xlsx sync
  linkedGoogleSheetUrl: '', // URL to user's cloud Google Sheet
  mistakes: [],
  questionType: 'choice', // 'choice' | 'translation' | 'fill' | 'reading'
  answerKeyMap: {}, // { "1": "C", "2": "B", ... }
  selectedErrorTags: new Set(),
  currentCropDataUrl: '',

  // Document Viewer State ('exam' | 'answer')
  activeDocSlot: 'exam',
  docs: {
    exam: {
      type: null, // 'pdf' | 'image' | 'sample'
      pdfDoc: null,
      imageObj: null,
      pageNum: 1,
      numPages: 0,
      scale: 1.25,
      fileName: '',
      sampleItems: [] // virtual text items when using sample exam generator
    },
    answer: {
      type: null,
      pdfDoc: null,
      imageObj: null,
      pageNum: 1,
      numPages: 0,
      scale: 1.25,
      fileName: '',
      sampleItems: []
    }
  },

  // Crop Tool State
  cropMode: false,
  isDraggingCrop: false,
  cropStartX: 0,
  cropStartY: 0,

  // Practice Quiz State
  quizQueue: [],
  quizIndex: 0,
  quizSelectedOptKey: null,
  quizSelectedConceptIdx: null,
  quizSubmitted: false,
  quizCorrectCount: 0,
  quizWrongItems: [],
  currentQuestionRenderData: null
};

// English-only Diagnostic Guided Prompts
const DIAGNOSIS_GUIDES = {
  '單字與片語搭配': {
    question: (w, c) =>
      `📘 【單字／片語診斷】你選了 (${w || '?'})，但正解是 (${c || '?'})。回頭看前後文線索，這兩個字的語意或搭配有什麼差別？`,
    phrases: [
      '把相似拼法或同字根單字的意思搞混了',
      '忽略了空格前後的固定介系詞或片語搭配',
      '沒有注意上下文的正負向語氣或轉折詞'
    ]
  },
  '文法與詞性判斷': {
    question: (w, c) =>
      `🔤 【文法與詞性診斷】你選了 (${w || '?'})，但正解是 (${c || '?'})。請檢查空格前後需要哪一種詞性或子句結構？`,
    phrases: [
      '沒有先看空格前後判斷詞性（n. / v. / adj. / adv.）',
      '關係代名詞忘記看先行詞是人還是物，或逗號後不能用 that',
      '分詞主動 (V-ing) 與被動 (p.p.) 判斷錯誤'
    ]
  },
  '篇章結構與轉折': {
    question: (w, c) =>
      `📑 【篇章結構診斷】你選了 (${w || '?'})，但正解是 (${c || '?'})。這題的轉折詞與指代詞是怎麼銜接前後句的？`,
    phrases: [
      '只看前一句就選，沒有檢查後一句是否接得上',
      '漏看開頭轉折詞（However / Therefore / For example）的邏輯',
      '沒確認指代詞（this / these / such / they）回指前文的哪個名詞'
    ]
  },
  '閱讀定位與推論': {
    question: (w, c) =>
      `📖 【閱讀測驗診斷】你選了 (${w || '?'})，但正解是 (${c || '?'})。回原文定位句子後，是哪裡判斷失誤？`,
    phrases: [
      '選了只出現單一段落細節的選項（主旨題以偏概全）',
      '看到原文眼熟單字就選，沒發現因果或主詞被偷換了',
      '過度推論了原文沒有提到的內容'
    ]
  },
  '翻譯句型與拼字': {
    question: (w, c) =>
      `✍️ 【翻譯寫作診斷】對照標準答案，這次主要是句型、時態還是拼字出錯？`,
    phrases: [
      '動詞時態忘記配合時間詞改為過去式或完成式',
      '核心句型（如 It takes...to V、so...that）或片語寫錯',
      '單字拼錯、漏寫冠詞 a/an/the 或名詞複數 -s'
    ]
  }
};

// ============================================================================
// Initialization & Event Binding
// ============================================================================
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  initSpreadsheetModal();
  initPdfViewerAndCrop();
  initAnswerKeyDrawer();
  initMistakeForm();
  initCardsTab();
  initPracticeQuizTab();
  initSheetTab();

  // Check if localStorage has previous autosaved session
  checkAutosaveOnStartup();
  updateAllViews();
});

function initTabs() {
  const tabBtns = document.querySelectorAll('.nav-tab');
  tabBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      const targetId = btn.getAttribute('data-tab');
      switchTab(targetId);
    });
  });
}

function switchTab(tabId) {
  document.querySelectorAll('.nav-tab').forEach((b) => {
    b.classList.toggle('active', b.getAttribute('data-tab') === tabId);
  });
  document.querySelectorAll('.tab-panel').forEach((p) => {
    p.classList.toggle('active', p.id === tabId);
  });

  if (tabId === 'tab-cards') {
    renderMistakeCards();
  } else if (tabId === 'tab-practice') {
    updateQuizSetupSummary();
  } else if (tabId === 'tab-sheet') {
    renderSheetTableAndStats();
  }
}

// ============================================================================
// STEP 0: Spreadsheet Selection Modal (Continue Existing vs. Open New)
// ============================================================================
function initSpreadsheetModal() {
  const modal = document.getElementById('sheetInitModal');
  const btnOpenModal = document.getElementById('btnOpenSheetModal');
  const btnTabChangeSheet = document.getElementById('btnTabChangeSheet');
  const cardExisting = document.getElementById('cardChoiceExisting');
  const cardNew = document.getElementById('cardChoiceNew');
  const inputExistingFile = document.getElementById('inputExistingSheetFile');
  const btnBindExistingExcel = document.getElementById('btnBindExistingExcelFile');
  const btnLoadGoogleUrl = document.getElementById('btnLoadGoogleSheetUrl');
  const btnCreateRealGoogle = document.getElementById('btnCreateRealGoogleSheet');
  const btnCreateRealExcel = document.getElementById('btnCreateRealExcelFile');
  const btnConfirmNew = document.getElementById('btnConfirmNewSheet');
  const btnOpenLinkedGoogle = document.getElementById('btnOpenLinkedGoogleSheet');

  const openModal = () => {
    modal.classList.remove('hidden');
  };

  btnOpenModal.addEventListener('click', openModal);
  btnTabChangeSheet.addEventListener('click', openModal);

  if (btnOpenLinkedGoogle) {
    btnOpenLinkedGoogle.addEventListener('click', () => {
      const targetUrl = state.linkedGoogleSheetUrl || 'https://docs.google.com/spreadsheets/';
      window.open(targetUrl, '_blank');
    });
  }

  cardExisting.addEventListener('click', () => {
    cardExisting.classList.add('selected');
    cardNew.classList.remove('selected');
  });

  cardNew.addEventListener('click', () => {
    cardNew.classList.add('selected');
    cardExisting.classList.remove('selected');
  });

  // Choice 1A: Select & Bind Existing Local Excel File (.xlsx / .csv)
  if (btnBindExistingExcel) {
    btnBindExistingExcel.addEventListener('click', async () => {
      if (window.showOpenFilePicker) {
        try {
          const [handle] = await window.showOpenFilePicker({
            types: [
              {
                description: 'Excel 或 CSV 試算表',
                accept: {
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                  'application/vnd.ms-excel': ['.xls'],
                  'text/csv': ['.csv']
                }
              }
            ],
            multiple: false
          });
          state.excelFileHandle = handle;
          const file = await handle.getFile();
          loadExistingSpreadsheetFile(file, true);
          return;
        } catch (err) {
          if (err && err.name === 'AbortError') return;
        }
      }
      inputExistingFile.click();
    });
  }

  inputExistingFile.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    state.excelFileHandle = null;
    loadExistingSpreadsheetFile(file, false);
    e.target.value = '';
  });

  // Choice 1B: Load & Open Existing Cloud Google Sheet by URL
  if (btnLoadGoogleUrl) {
    btnLoadGoogleUrl.addEventListener('click', async () => {
      const rawUrl = document.getElementById('inputExistingGoogleSheetUrl').value.trim();
      if (!rawUrl) {
        alert('請先貼上您的 Google 試算表網址（例如 https://docs.google.com/spreadsheets/d/...）！');
        return;
      }
      await loadAndOpenExistingGoogleSheetUrl(rawUrl);
    });
  }

  // Choice 2A: Physically Open a Brand New Google Sheet in a New Tab (https://sheets.new)
  if (btnCreateRealGoogle) {
    btnCreateRealGoogle.addEventListener('click', async () => {
      const nameInput = document.getElementById('inputNewSheetName').value.trim();
      state.sheetName = nameInput || '我的錯題進化本_2026';
      state.sheetMode = 'new';
      state.originalLoadedCount = 0;
      state.excelFileHandle = null;
      state.linkedGoogleSheetUrl = 'https://docs.google.com/spreadsheets/';
      state.mistakes = [];

      // Copy header columns to clipboard so user can immediately Ctrl+V in cell A1 of the new Google Sheet
      const headers = [
        '錯題ID', '建立日期', '科目', '試卷單元', '題號',
        '題目內容', '選項A', '選項B', '選項C', '選項D',
        '我的錯誤答案', '正確答案', '錯誤類型標籤', '為何選錯診斷',
        '正確觀念解析', '避坑錦囊', '練習次數', '答對次數', '熟練狀態'
      ];
      try {
        await navigator.clipboard.writeText(headers.join('\t'));
      } catch (e) {}

      // Open real Google Sheet creation URL in new tab
      window.open('https://sheets.new', '_blank');

      saveToLocalStorage();
      updateAllViews();
      modal.classList.add('hidden');
      showToast(
        `🟢 已在新分頁為您打開全新的 Google 試算表！並已自動複製「19 個標準欄位標題」，您可在新分頁 A1 儲存格直接按 Ctrl+V 貼上！`
      );
    });
  }

  // Choice 2B: Physically Create & Save a Brand New Local Excel (.xlsx) File
  if (btnCreateRealExcel) {
    btnCreateRealExcel.addEventListener('click', async () => {
      const nameInput = document.getElementById('inputNewSheetName').value.trim();
      state.sheetName = nameInput || '我的錯題進化本_2026';
      state.sheetMode = 'new';
      state.originalLoadedCount = 0;
      state.linkedGoogleSheetUrl = '';
      state.mistakes = [];

      const safeName = state.sheetName.replace(/[\\/:*?"<>|]/g, '_');

      if (window.showSaveFilePicker) {
        try {
          const handle = await window.showSaveFilePicker({
            suggestedName: `${safeName}.xlsx`,
            types: [
              {
                description: 'Excel 試算表 (.xlsx)',
                accept: {
                  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
                }
              }
            ]
          });
          state.excelFileHandle = handle;
          await writeCurrentWorkbookToHandle(handle);
          saveToLocalStorage();
          updateAllViews();
          modal.classList.add('hidden');
          showToast(
            `📗 已在您的電腦建立「${handle.name}」並完成綁定！現在可直接在電腦打開它，且之後每新增一筆錯題都會自動同步寫入該檔案！`
          );
          return;
        } catch (err) {
          if (err && err.name === 'AbortError') return;
        }
      }

      // Fallback: Immediately download the newly created .xlsx file with headers
      downloadBlankOrCurrentWorkbookXlsx(safeName);
      saveToLocalStorage();
      updateAllViews();
      modal.classList.add('hidden');
      showToast(
        `📗 已建立並下載全新 Excel 試算表「${safeName}.xlsx」！您可以直接點擊打開它。`
      );
    });
  }

  // Choice 2: Start Brand New Spreadsheet (Do NOT open a blank Excel/Google Sheet upfront!)
  btnConfirmNew.addEventListener('click', () => {
    const nameInput = document.getElementById('inputNewSheetName').value.trim();
    state.sheetName = nameInput || '我的錯題進化本_2026';
    state.sheetMode = 'new';
    state.originalLoadedCount = 0;
    state.excelFileHandle = null;
    state.linkedGoogleSheetUrl = '';
    state.mistakes = [];
    saveToLocalStorage();
    updateAllViews();
    modal.classList.add('hidden');
    switchTab('tab-record');
    showToast(`✨ 已設定新試算表「${state.sheetName}」！請先整理錯題，完成後即可一鍵匯出成完整的 Excel 或 Google 試算表。`);
  });

  // Quick Export Button in Header
  document.getElementById('btnQuickExportExcel').addEventListener('click', () => {
    exportToExcelXlsx();
  });

  // Load Demo Data Button in Header
  document.getElementById('btnLoadDemoData').addEventListener('click', () => {
    loadBuiltInDemoEverything();
  });
}

function checkAutosaveOnStartup() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const saved = JSON.parse(raw);
    if (saved && Array.isArray(saved.mistakes) && saved.mistakes.length > 0) {
      const banner = document.getElementById('localAutosaveBanner');
      const desc = document.getElementById('localAutosaveDesc');
      banner.style.display = 'flex';
      desc.textContent = `試算表「${saved.sheetName || '未命名'}」共有 ${saved.mistakes.length} 筆錯題紀錄`;
      document.getElementById('btnRestoreAutosave').onclick = () => {
        state.sheetName = saved.sheetName || '我的錯題進化本_2026';
        state.sheetMode = saved.sheetMode || 'existing';
        state.originalLoadedCount = saved.originalLoadedCount || saved.mistakes.length;
        state.mistakes = saved.mistakes;
        state.answerKeyMap = saved.answerKeyMap || {};
        document.getElementById('sheetInitModal').classList.add('hidden');
        renderAnswerKeyChips();
        updateAllViews();
        showToast(`💾 已成功還原「${state.sheetName}」共 ${state.mistakes.length} 筆錯題！`);
      };
    }
  } catch (err) {
    console.warn('Autosave read failed:', err);
  }
}

function saveToLocalStorage() {
  try {
    const payload = {
      sheetName: state.sheetName,
      sheetMode: state.sheetMode,
      originalLoadedCount: state.originalLoadedCount,
      linkedGoogleSheetUrl: state.linkedGoogleSheetUrl || '',
      mistakes: state.mistakes,
      answerKeyMap: state.answerKeyMap,
      savedAt: new Date().toISOString()
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  } catch (err) {
    console.warn('Autosave write failed (likely image quota):', err);
  }
  // If user bound a local .xlsx file via File System Access API, auto-write changes to disk
  if (state.excelFileHandle) {
    writeCurrentWorkbookToHandle(state.excelFileHandle).catch((e) =>
      console.warn('Auto-sync to bound Excel file skipped:', e)
    );
  }
}

function parseJsonRowsToMistakes(jsonRows) {
  return jsonRows
    .map((row, idx) => {
      const qText =
        row['題目內容'] ||
        row['題目敘述'] ||
        row['題目'] ||
        row['Question'] ||
        '';
      const wrongAns = String(
        row['我的錯誤答案'] || row['錯誤答案'] || row['誤選'] || ''
      ).trim();
      const correctAns = String(
        row['正確答案'] || row['正解'] || ''
      ).trim();

      if (!qText && !wrongAns && !correctAns) return null;

      const tagsStr = String(row['錯誤類型標籤'] || row['錯誤類型'] || '');
      const errorTags = tagsStr
        ? tagsStr
            .split(/[,，、/]/)
            .map((s) => s.trim())
            .filter(Boolean)
        : ['觀念混淆'];

      const rawTypeLabel = String(row['題型'] || '').trim();
      let inferredType = 'choice';
      if (rawTypeLabel.includes('翻譯')) {
        inferredType = 'translation';
      } else if (rawTypeLabel.includes('克漏')) {
        inferredType = 'cloze';
      } else if (rawTypeLabel.includes('文意選填')) {
        inferredType = 'wordbank';
      } else if (rawTypeLabel.includes('篇章')) {
        inferredType = 'structure';
      } else if (rawTypeLabel.includes('閱讀')) {
        inferredType = 'reading';
      } else if (row['選項G'] || row['選項H'] || /^[G-J]$/i.test(correctAns)) {
        inferredType = 'wordbank';
      } else if (row['選項F'] || /^[EF]$/i.test(correctAns)) {
        inferredType = 'structure';
      } else if (
        !row['選項A'] &&
        !row['選項B'] &&
        (correctAns.length > 2 || wrongAns.length > 2)
      ) {
        inferredType = 'translation';
      }

      return {
        id: row['錯題ID'] || `M_${Date.now()}_${idx}`,
        createdAt: row['建立日期'] || new Date().toLocaleDateString('zh-TW'),
        subject: '英文',
        unit: row['試卷單元'] || row['單元'] || '既有試算表匯入',
        questionNum: String(row['題號'] || idx + 1),
        questionType: inferredType,
        questionText: qText,
        optA: String(row['選項A'] || ''),
        optB: String(row['選項B'] || ''),
        optC: String(row['選項C'] || ''),
        optD: String(row['選項D'] || ''),
        optE: String(row['選項E'] || ''),
        optF: String(row['選項F'] || ''),
        optG: String(row['選項G'] || ''),
        optH: String(row['選項H'] || ''),
        optI: String(row['選項I'] || ''),
        optJ: String(row['選項J'] || ''),
        wrongAnswer: wrongAns || 'B',
        correctAnswer: correctAns || 'C',
        errorTags,
        whyWrong: String(row['為何選錯診斷'] || row['錯誤原因'] || ''),
        correctConcept: String(row['正確觀念解析'] || row['正確觀念'] || ''),
        keyTakeaway: String(row['避坑錦囊'] || row['重點提醒'] || ''),
        practiceCount: Number(row['練習次數'] || 0),
        correctCount: Number(row['答對次數'] || 0),
        masteryStatus: row['熟練狀態'] || '待加強',
        cropImageDataUrl: String(row['題目截圖資料'] || ''),
        isHistorical: true
      };
    })
    .filter(Boolean);
}

function loadExistingSpreadsheetFile(file, isBoundHandle = false) {
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

      const loadedMistakes = parseJsonRowsToMistakes(jsonRows);

      const cleanBaseName = file.name.replace(/\.(xlsx|xls|csv)$/i, '');
      state.sheetName = cleanBaseName;
      state.sheetMode = 'existing';
      state.originalLoadedCount = loadedMistakes.length;
      state.mistakes = loadedMistakes;

      saveToLocalStorage();
      updateAllViews();
      document.getElementById('sheetInitModal').classList.add('hidden');
      const syncMsg = isBoundHandle
        ? '（已直接綁定原檔案，新增的錯題將自動寫入此 Excel 檔！）'
        : '（新增錯題後可點擊匯出更新試算表）';
      showToast(
        `📂 成功載入既有試算表「${cleanBaseName}」，共載入 ${loadedMistakes.length} 筆歷史錯題！${syncMsg}`
      );
    } catch (err) {
      console.error(err);
      alert('讀取試算表失敗，請確認檔案為有效的 Excel (.xlsx / .xls) 或 CSV 檔案。');
    }
  };
  reader.readAsArrayBuffer(file);
}

async function loadAndOpenExistingGoogleSheetUrl(rawUrl) {
  const match = rawUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    alert('網址格式不正確！請貼上包含 /spreadsheets/d/... 的 Google 試算表網址。');
    return;
  }
  const sheetId = match[1];
  state.linkedGoogleSheetUrl = rawUrl;
  state.sheetName = `雲端Google試算表_${sheetId.slice(0, 6)}`;
  state.sheetMode = 'existing';

  // Physically open the user's Google Sheet in a new tab
  window.open(rawUrl, '_blank');

  // Try fetching public CSV from the Google Sheet
  try {
    const csvExportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv`;
    const resp = await fetch(csvExportUrl);
    if (resp.ok) {
      const csvText = await resp.text();
      const workbook = XLSX.read(csvText, { type: 'string' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });
      const loadedMistakes = parseJsonRowsToMistakes(jsonRows);
      state.originalLoadedCount = loadedMistakes.length;
      state.mistakes = loadedMistakes;
      saveToLocalStorage();
      updateAllViews();
      document.getElementById('sheetInitModal').classList.add('hidden');
      showToast(
        `🟢 已在新分頁開啟您的 Google 試算表，並成功從雲端載入 ${loadedMistakes.length} 筆既有錯題！`
      );
      return;
    }
  } catch (err) {
    console.warn('Google Sheet CSV fetch warning (may require public view permission):', err);
  }

  saveToLocalStorage();
  updateAllViews();
  document.getElementById('sheetInitModal').classList.add('hidden');
  showToast(
    '🟢 已在新分頁為您打開該份 Google 試算表！（若需自動載入雲端舊題，請將試算表共用權限設為「知道連結的任何人可檢視」）'
  );
}

function buildWorkbookObject(includeBase64Crop = true) {
  let rows = formatMistakesForSpreadsheetExport(includeBase64Crop);
  if (rows.length === 0) {
    // Provide header-only template row structure when creating a brand new empty spreadsheet
    rows = [
      {
        錯題ID: '',
        建立日期: '',
        科目: '',
        試卷單元: '',
        題號: '',
        題目內容: '',
        選項A: '',
        選項B: '',
        選項C: '',
        選項D: '',
        我的錯誤答案: '',
        正確答案: '',
        錯誤類型標籤: '',
        為何選錯診斷: '',
        正確觀念解析: '',
        避坑錦囊: '',
        練習次數: '',
        答對次數: '',
        熟練狀態: '',
        題目截圖資料: ''
      }
    ];
  }
  const worksheet = XLSX.utils.json_to_sheet(rows);
  if (state.mistakes.length === 0) {
    // Keep only row 1 (the header row) if mistakes array is empty
    worksheet['!ref'] = 'A1:T1';
  }
  worksheet['!cols'] = [
    { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 20 }, { wch: 8 },
    { wch: 42 }, { wch: 20 }, { wch: 20 }, { wch: 20 }, { wch: 20 },
    { wch: 12 }, { wch: 10 }, { wch: 22 }, { wch: 36 }, { wch: 36 },
    { wch: 32 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 18 }
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '錯題紀錄總表');
  return workbook;
}

async function writeCurrentWorkbookToHandle(handle) {
  const workbook = buildWorkbookObject(true);
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const writable = await handle.createWritable();
  await writable.write(new Blob([wbout], { type: 'application/octet-stream' }));
  await writable.close();
}

function downloadBlankOrCurrentWorkbookXlsx(safeName) {
  const workbook = buildWorkbookObject(true);
  XLSX.writeFile(workbook, `${safeName}.xlsx`);
}

// ============================================================================
// STEP 1A: PDF / Image Viewer & Drag-to-Crop + Auto Text Extraction
// ============================================================================
function initPdfViewerAndCrop() {
  const inputExam = document.getElementById('inputExamFile');
  const inputAnswer = document.getElementById('inputAnswerFile');
  const btnViewExam = document.getElementById('btnViewExamDoc');
  const btnViewAnswer = document.getElementById('btnViewAnswerDoc');
  const btnPrev = document.getElementById('btnPrevPage');
  const btnNext = document.getElementById('btnNextPage');
  const btnZoomIn = document.getElementById('btnZoomIn');
  const btnZoomOut = document.getElementById('btnZoomOut');
  const btnToggleCrop = document.getElementById('btnToggleCropMode');
  const btnCropPassage = document.getElementById('btnCropPassageMode');
  const btnSample1 = document.getElementById('btnGenerateSamplePdf');
  const btnSample2 = document.getElementById('btnEmptyStateSamplePdf');

  inputExam.addEventListener('change', (e) => {
    if (e.target.files[0]) handleDocFileUpload('exam', e.target.files[0]);
  });

  inputAnswer.addEventListener('change', (e) => {
    if (e.target.files[0]) handleDocFileUpload('answer', e.target.files[0]);
  });

  btnViewExam.addEventListener('click', () => switchDocSlot('exam'));
  btnViewAnswer.addEventListener('click', () => switchDocSlot('answer'));

  btnPrev.addEventListener('click', () => {
    const doc = state.docs[state.activeDocSlot];
    if (doc.pageNum > 1) {
      doc.pageNum--;
      renderActiveDocument();
    }
  });

  btnNext.addEventListener('click', () => {
    const doc = state.docs[state.activeDocSlot];
    if (doc.pageNum < doc.numPages) {
      doc.pageNum++;
      renderActiveDocument();
    }
  });

  btnZoomIn.addEventListener('click', () => {
    const doc = state.docs[state.activeDocSlot];
    doc.scale = Math.min(2.5, +(doc.scale + 0.2).toFixed(2));
    renderActiveDocument();
  });

  btnZoomOut.addEventListener('click', () => {
    const doc = state.docs[state.activeDocSlot];
    doc.scale = Math.max(0.6, +(doc.scale - 0.2).toFixed(2));
    renderActiveDocument();
  });

  btnToggleCrop.addEventListener('click', () => {
    const shouldEnable = !(state.cropMode && state.cropTarget === 'question');
    setCropMode(shouldEnable, 'question');
  });

  if (btnCropPassage) {
    btnCropPassage.addEventListener('click', () => {
      const shouldEnable = !(state.cropMode && state.cropTarget === 'passage');
      setCropMode(shouldEnable, 'passage');
    });
  }

  btnSample1.addEventListener('click', () => generateInteractiveSampleExam());
  btnSample2.addEventListener('click', () => generateInteractiveSampleExam());

  initCropCanvasMouseEvents();
}

function setCropMode(enabled, modeType = 'question') {
  state.cropMode = Boolean(enabled);
  state.cropTarget = modeType || 'question';
  const btn = document.getElementById('btnToggleCropMode');
  const btnPassage = document.getElementById('btnCropPassageMode');
  const overlay = document.getElementById('cropOverlayCanvas');
  const banner = document.getElementById('cropHintBanner');

  if (btn) {
    const isQuestionCrop = state.cropMode && state.cropTarget === 'question';
    btn.classList.toggle('active', isQuestionCrop);
    btn.textContent = isQuestionCrop
      ? '🛑 結束框選（請在下方試卷拖曳框選該題）'
      : '✂️ 框選題目與選項';
  }
  if (btnPassage) {
    const isPassageCrop = state.cropMode && state.cropTarget === 'passage';
    btnPassage.classList.toggle('active', isPassageCrop);
    btnPassage.style.background = isPassageCrop
      ? '#ef4444'
      : 'rgba(255,255,255,0.15)';
    btnPassage.style.color = '#ffffff';
    btnPassage.textContent = isPassageCrop
      ? '🛑 結束文章框選（請在試卷拖曳框選文章段落）'
      : '📖 加框克漏字／文意選填文章';
  }
  if (overlay) {
    overlay.classList.toggle('active', state.cropMode);
  }
  if (banner) {
    banner.style.display = state.cropMode ? 'block' : 'none';
    banner.textContent =
      state.cropTarget === 'passage'
        ? '📖 【加框克漏字／文意選填文章模式】請在下方試卷拖曳框選「文章段落或 (A)~(J) 選項框」，系統會自動將文章截圖拼接到上方、抓出對應空格句子並提供多空格切換按鈕！'
        : '💡 請在下方試卷上「拖曳框選你要的那一題（例如第 36 題或整段文意選填）」，系統會自動抓取題幹、文意選填 (A)~(J) 詞庫與對應文章句子！';
  }
}

function withTimeout(promise, ms = 1400, fallbackValue = null) {
  return Promise.race([
    promise,
    new Promise((resolve) => setTimeout(() => resolve(fallbackValue), ms))
  ]);
}

function switchDocSlot(slot) {
  state.activeDocSlot = slot;
  const btnExam = document.getElementById('btnViewExamDoc');
  const btnAns = document.getElementById('btnViewAnswerDoc');
  btnExam.style.background =
    slot === 'exam' ? 'rgba(79,70,229,0.9)' : 'rgba(255,255,255,0.15)';
  btnAns.style.background =
    slot === 'answer' ? 'rgba(16,185,129,0.9)' : 'rgba(255,255,255,0.15)';
  renderActiveDocument();
}

async function handleDocFileUpload(slot, file) {
  const docState = state.docs[slot];
  docState.fileName = file.name;

  const boxId = slot === 'exam' ? 'examUploadBox' : 'answerUploadBox';
  const nameId = slot === 'exam' ? 'examFileName' : 'answerFileName';
  document.getElementById(boxId).classList.add('loaded');
  document.getElementById(nameId).textContent = `已載入：${file.name}`;

  if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
    const arrayBuf = await file.arrayBuffer();
    const pdfDoc = await window.pdfjsLib.getDocument({
      data: arrayBuf,
      cMapUrl: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',
      cMapPacked: true,
      standardFontDataUrl:
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/standard_fonts/'
    }).promise;
    docState.type = 'pdf';
    docState.pdfDoc = pdfDoc;
    docState.pageNum = 1;
    docState.numPages = pdfDoc.numPages;
    docState.scale = 1.3;
  } else {
    // Image file
    const dataUrl = await new Promise((res) => {
      const r = new FileReader();
      r.onload = (e) => res(e.target.result);
      r.readAsDataURL(file);
    });
    const img = new Image();
    await new Promise((res) => {
      img.onload = res;
      img.src = dataUrl;
    });
    docState.type = 'image';
    docState.imageObj = img;
    docState.pageNum = 1;
    docState.numPages = 1;
    docState.scale = 1.0;

    // If user uploaded a single-question photo/strip for the exam, auto-preview & OCR it right away!
    if (
      slot === 'exam' &&
      (img.naturalWidth / Math.max(img.naturalHeight, 1) >= 1.35 ||
        img.naturalHeight < 750)
    ) {
      setCroppedImagePreview(dataUrl);
      setTimeout(() => {
        recognizeTextAndMathFromCroppedImage(dataUrl, file.name, 'question');
      }, 120);
    }
  }

  switchDocSlot(slot);
  showToast(`📄 已成功載入${slot === 'exam' ? '試卷' : '解答'}檔案：${file.name}`);
}

async function renderActiveDocument() {
  const docState = state.docs[state.activeDocSlot];
  const emptyState = document.getElementById('pdfEmptyState');
  const wrapper = document.getElementById('pdfCanvasWrapper');
  const canvas = document.getElementById('pdfCanvas');
  const overlay = document.getElementById('cropOverlayCanvas');
  const pageInd = document.getElementById('pdfPageIndicator');
  const zoomInd = document.getElementById('pdfZoomIndicator');
  const btnPrev = document.getElementById('btnPrevPage');
  const btnNext = document.getElementById('btnNextPage');

  if (!docState.type) {
    emptyState.style.display = 'block';
    wrapper.style.display = 'none';
    pageInd.textContent = '第 0 / 0 頁';
    btnPrev.disabled = true;
    btnNext.disabled = true;
    return;
  }

  emptyState.style.display = 'none';
  wrapper.style.display = 'inline-block';
  pageInd.textContent = `第 ${docState.pageNum} / ${docState.numPages} 頁`;
  zoomInd.textContent = `${Math.round(docState.scale * 100)}%`;
  btnPrev.disabled = docState.pageNum <= 1;
  btnNext.disabled = docState.pageNum >= docState.numPages;

  const ctx = canvas.getContext('2d');

  if (docState.type === 'pdf' && docState.pdfDoc) {
    const page = await docState.pdfDoc.getPage(docState.pageNum);
    const viewport = page.getViewport({ scale: docState.scale });
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    overlay.width = viewport.width;
    overlay.height = viewport.height;
    await page.render({ canvasContext: ctx, viewport }).promise;
  } else if (docState.type === 'image' && docState.imageObj) {
    const w = docState.imageObj.naturalWidth * docState.scale;
    const h = docState.imageObj.naturalHeight * docState.scale;
    canvas.width = w;
    canvas.height = h;
    overlay.width = w;
    overlay.height = h;
    ctx.clearRect(0, 0, w, h);
    ctx.drawImage(docState.imageObj, 0, 0, w, h);
  } else if (docState.type === 'sample') {
    drawSampleExamCanvas(canvas, overlay, docState);
  }
}

// Built-in Sample Exam Generator so students can test PDF cropping & text extraction out of the box
function generateInteractiveSampleExam() {
  const examDoc = state.docs.exam;
  examDoc.type = 'sample';
  examDoc.pageNum = 1;
  examDoc.numPages = 1;
  examDoc.scale = 1.0;
  examDoc.fileName = '115學年度_高中英文段考精選試卷.pdf';

  document.getElementById('examUploadBox').classList.add('loaded');
  document.getElementById('examFileName').textContent = examDoc.fileName;

  // Populate standard answers for the English sample exam
  state.answerKeyMap = { '1': 'A', '2': 'B', '3': 'J', '4': 'C' };
  document.getElementById('inputBulkAnswers').value = '1:A, 2:B, 3:J, 4:C';
  renderAnswerKeyChips();

  switchDocSlot('exam');
  setCropMode(true);
  showToast(
    '🧪 已載入內建高中英文測試試卷與標準答案表（1:A, 2:B, 3:J, 4:C）！現在可直接在試卷上拖曳框選任一題！'
  );
}

function drawSampleExamCanvas(canvas, overlay, docState) {
  const baseW = 760;
  const baseH = 920;
  const scale = docState.scale;
  canvas.width = baseW * scale;
  canvas.height = baseH * scale;
  overlay.width = canvas.width;
  overlay.height = canvas.height;

  const ctx = canvas.getContext('2d');
  ctx.save();
  ctx.scale(scale, scale);

  // Paper Background
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, baseW, baseH);

  // Header Border & Exam Title
  ctx.strokeStyle = '#1e293b';
  ctx.lineWidth = 2;
  ctx.strokeRect(28, 28, baseW - 56, baseH - 56);

  ctx.fillStyle = '#0f172a';
  ctx.font = 'bold 22px "Noto Sans TC", sans-serif';
  ctx.textAlign = 'center';
  ctx.fillText('115 學年度第一學期 高中英文科 段考精選試卷', baseW / 2, 72);

  ctx.font = '14px "Noto Sans TC", sans-serif';
  ctx.fillStyle = '#475569';
  ctx.fillText(
    '請在畫面上使用「✂️ 框選題目與選項」拖曳框出任一題，即可自動擷取題幹與選項！',
    baseW / 2,
    100
  );

  ctx.beginPath();
  ctx.moveTo(45, 116);
  ctx.lineTo(baseW - 45, 116);
  ctx.strokeStyle = '#cbd5e1';
  ctx.lineWidth = 1;
  ctx.stroke();

  const sampleQuestions = [
    {
      yTop: 135,
      yBottom: 305,
      qNum: '1',
      subject: '英文',
      questionType: 'choice',
      unit: 'I. 單字選擇 (Vocabulary)',
      stem: '1. Eric has been a(n) ______ employee at the company for over ten years, always dependable and committed to his work.',
      optA: 'faithful',
      optB: 'artificial',
      optC: 'reluctant',
      optD: 'potential'
    },
    {
      yTop: 320,
      yBottom: 495,
      qNum: '2',
      subject: '英文',
      questionType: 'cloze',
      unit: 'II. 綜合測驗／克漏字 (Cloze)',
      stem: '2. She finds it hard to accept that some selfish people, ______ have probably never cleaned a beach in their lives, simply don\'t care at all about the planet.',
      optA: 'that',
      optB: 'who',
      optC: 'whom',
      optD: 'which'
    },
    {
      yTop: 510,
      yBottom: 685,
      qNum: '3',
      subject: '英文',
      questionType: 'wordbank',
      unit: 'III. 文意選填 (Word Bank A~J)',
      stem: '3. Below me, the thrilling clouds and deep blue sea provide a breathtaking view. I am ___3___ by the sheer beauty of nature.',
      optA: 'flight',
      optB: 'managed',
      optC: 'spot',
      optD: 'craft',
      optE: 'destination',
      optF: 'impulse',
      optG: 'thrilling',
      optH: 'missions',
      optI: 'everlasting',
      optJ: 'dazzled'
    },
    {
      yTop: 700,
      yBottom: 870,
      qNum: '4',
      subject: '英文',
      questionType: 'structure',
      unit: 'IV. 篇章結構 (Discourse Structure)',
      stem: '4. Convenience stores in East Asia are open 24 hours a day and offer a wide variety of ready-to-eat meals. ___4___ Single-use plastic containers and bags from these stores have created a serious waste problem.',
      optA: 'First of all, big cities like Tokyo and Taipei have very high population densities.',
      optB: 'In addition, people can pay bills and pick up packages there.',
      optC: 'However, such convenience comes at a heavy environmental cost.',
      optD: 'Therefore, restaurants are much more expensive than convenience stores.'
    }
  ];

  docState.sampleItems = sampleQuestions;

  ctx.textAlign = 'left';
  sampleQuestions.forEach((q) => {
    // Question card box on exam paper
    ctx.fillStyle = '#f8fafc';
    ctx.fillRect(48, q.yTop, baseW - 96, q.yBottom - q.yTop - 10);
    ctx.strokeStyle = '#e2e8f0';
    ctx.strokeRect(48, q.yTop, baseW - 96, q.yBottom - q.yTop - 10);

    // Draw wrapped question stem
    ctx.fillStyle = '#0f172a';
    ctx.font = 'bold 14px "Noto Sans TC", sans-serif';
    wrapTextOnCanvas(ctx, `[${q.unit}] ${q.stem}`, 62, q.yTop + 26, baseW - 124, 22);

    // Draw options
    ctx.font = '13px "Noto Sans TC", sans-serif';
    ctx.fillStyle = '#334155';
    ctx.fillText(`(A) ${q.optA.slice(0, 34)}`, 68, q.yTop + 100);
    ctx.fillText(`(B) ${q.optB.slice(0, 34)}`, 390, q.yTop + 100);
    ctx.fillText(`(C) ${q.optC.slice(0, 34)}`, 68, q.yTop + 130);
    ctx.fillText(`(D) ${(q.optJ ? `(J) ${q.optJ}` : q.optD).slice(0, 34)}`, 390, q.yTop + 130);
  });

  ctx.restore();
}

function wrapTextOnCanvas(ctx, text, x, y, maxWidth, lineHeight) {
  const chars = text.split('');
  let line = '';
  let curY = y;
  for (let i = 0; i < chars.length; i++) {
    const testLine = line + chars[i];
    const metrics = ctx.measureText(testLine);
    if (metrics.width > maxWidth && i > 0) {
      ctx.fillText(line, x, curY);
      line = chars[i];
      curY += lineHeight;
    } else {
      line = testLine;
    }
  }
  ctx.fillText(line, x, curY);
}

// Crop Mouse Interaction on Overlay Canvas
function initCropCanvasMouseEvents() {
  const overlay = document.getElementById('cropOverlayCanvas');
  const pdfCanvas = document.getElementById('pdfCanvas');

  const getCanvasCoords = (evt) => {
    const rect = overlay.getBoundingClientRect();
    const scaleX = overlay.width / rect.width;
    const scaleY = overlay.height / rect.height;
    return {
      x: (evt.clientX - rect.left) * scaleX,
      y: (evt.clientY - rect.top) * scaleY
    };
  };

  overlay.addEventListener('mousedown', (evt) => {
    if (!state.cropMode) return;
    state.isDraggingCrop = true;
    const pos = getCanvasCoords(evt);
    state.cropStartX = pos.x;
    state.cropStartY = pos.y;
  });

  overlay.addEventListener('mousemove', (evt) => {
    if (!state.cropMode || !state.isDraggingCrop) return;
    const pos = getCanvasCoords(evt);
    const ctx = overlay.getContext('2d');
    ctx.clearRect(0, 0, overlay.width, overlay.height);

    // Semi-transparent dark mask
    ctx.fillStyle = 'rgba(15, 23, 42, 0.35)';
    ctx.fillRect(0, 0, overlay.width, overlay.height);

    const rx = Math.min(state.cropStartX, pos.x);
    const ry = Math.min(state.cropStartY, pos.y);
    const rw = Math.abs(pos.x - state.cropStartX);
    const rh = Math.abs(pos.y - state.cropStartY);

    ctx.clearRect(rx, ry, rw, rh);
    ctx.strokeStyle = '#ef4444';
    ctx.lineWidth = 2.5;
    ctx.setLineDash([6, 4]);
    ctx.strokeRect(rx, ry, rw, rh);
  });

  const finishCropSelection = async (evt) => {
    if (!state.cropMode || !state.isDraggingCrop) return;
    state.isDraggingCrop = false;
    const activeTarget = state.cropTarget || 'question';
    // Immediately exit crop mode BEFORE any async image stitching / PDF / OCR calls so the UI NEVER gets stuck!
    setCropMode(false, 'question');

    const pos = getCanvasCoords(evt);
    const overlayCtx = overlay.getContext('2d');
    overlayCtx.clearRect(0, 0, overlay.width, overlay.height);

    const rx = Math.max(0, Math.min(state.cropStartX, pos.x));
    const ry = Math.max(0, Math.min(state.cropStartY, pos.y));
    const rw = Math.min(overlay.width - rx, Math.abs(pos.x - state.cropStartX));
    const rh = Math.min(overlay.height - ry, Math.abs(pos.y - state.cropStartY));

    if (rw < 18 || rh < 14) return; // Ignore tiny accidental clicks

    // 1. Capture cropped image from #pdfCanvas
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = rw;
    tempCanvas.height = rh;
    const tCtx = tempCanvas.getContext('2d');
    tCtx.drawImage(pdfCanvas, rx, ry, rw, rh, 0, 0, rw, rh);
    const croppedDataUrl = tempCanvas.toDataURL('image/png');

    if (activeTarget === 'passage' && state.currentCropDataUrl) {
      // Stitch Cloze / Word Bank passage image ABOVE existing question options image
      const mergedDataUrl = await stitchTwoImagesVertically(
        croppedDataUrl,
        state.currentCropDataUrl
      );
      setCroppedImagePreview(mergedDataUrl);
      await extractTextFromCropRegion(rx, ry, rw, rh, 'passage');
    } else {
      setCroppedImagePreview(croppedDataUrl);
      await extractTextFromCropRegion(rx, ry, rw, rh, activeTarget);
    }
  };

  overlay.addEventListener('mouseup', finishCropSelection);
  window.addEventListener('mouseup', (evt) => {
    if (state.cropMode && state.isDraggingCrop) {
      finishCropSelection(evt);
    }
  });
}

async function stitchTwoImagesVertically(topDataUrl, bottomDataUrl) {
  const loadImg = (src) =>
    new Promise((res) => {
      const img = new Image();
      img.onload = () => res(img);
      img.src = src;
    });
  const [topImg, botImg] = await Promise.all([
    loadImg(topDataUrl),
    loadImg(bottomDataUrl)
  ]);
  const c = document.createElement('canvas');
  c.width = Math.max(topImg.width, botImg.width);
  c.height = topImg.height + botImg.height + 10;
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, c.height);
  ctx.drawImage(topImg, 0, 0);
  ctx.strokeStyle = '#cbd5e1';
  ctx.beginPath();
  ctx.moveTo(0, topImg.height + 5);
  ctx.lineTo(c.width, topImg.height + 5);
  ctx.stroke();
  ctx.drawImage(botImg, 0, topImg.height + 10);
  return c.toDataURL('image/png');
}

function setCroppedImagePreview(dataUrl) {
  state.currentCropDataUrl = dataUrl;
  const previewArea = document.getElementById('cropPreviewArea');
  const previewImg = document.getElementById('cropPreviewImg');
  if (dataUrl) {
    previewImg.src = dataUrl;
    previewArea.style.display = 'block';
  } else {
    previewImg.src = '';
    previewArea.style.display = 'none';
  }
}

async function extractTextFromCropRegion(rx, ry, rw, rh, modeType = 'question') {
  const docState = state.docs[state.activeDocSlot];

  // Always set subject to English and sync unit from filename
  document.getElementById('fieldSubject').value = '英文';
  if (docState.fileName) {
    document.getElementById('fieldUnit').value = docState.fileName.replace(/\.pdf$/i, '');
    syncSubjectSpecificUI();
  }

  if (docState.type === 'sample') {
    const unscaledYCenter = (ry + rh / 2) / docState.scale;
    const matched = docState.sampleItems.find(
      (item) => unscaledYCenter >= item.yTop && unscaledYCenter <= item.yBottom
    );
    if (matched) {
      document.getElementById('fieldSubject').value = '英文';
      document.getElementById('fieldUnit').value = matched.unit;
      document.getElementById('fieldQuestionNum').value = matched.qNum;
      document.getElementById('fieldQuestionText').value = matched.stem;
      ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach((k) => {
        const el = document.getElementById(`fieldOpt${k}`);
        if (el) el.value = matched[`opt${k}`] || '';
      });
      setQuestionTypeUI(matched.questionType || 'choice');
      autoFillAnswerFromKeyMap(matched.qNum);
      updateSelfLearningCoach(true);
      runAutoMistakeDiagnosis(true);
      showToast(
        `✂️ 已自動擷取第 ${matched.qNum} 題（${getQuestionTypeLabel(
          matched.questionType || 'choice'
        )}）截圖、題幹與選項！`
      );
      return;
    }
  }

  let joinedText = '';
  let joinedLinesText = '';
  let fullPageText = '';

  if (docState.type === 'pdf' && docState.pdfDoc) {
    try {
      const page = await withTimeout(
        docState.pdfDoc.getPage(docState.pageNum),
        1200,
        null
      );
      if (page) {
        const viewport = page.getViewport({ scale: docState.scale });
        const textContent = await withTimeout(
          page.getTextContent(),
          1400,
          null
        );

        if (textContent && Array.isArray(textContent.items)) {
          const itemsInBox = [];
          const allPageItems = [];
          for (const item of textContent.items) {
            if (!item.str || !item.str.trim()) continue;
            const tx = window.pdfjsLib.Util.transform(
              viewport.transform,
              item.transform
            );
            const fontH = Math.hypot(tx[2], tx[3]) || 12;
            const itemW = Math.max((item.width || 8) * docState.scale, 8);
            const leftX = tx[4];
            const rightX = tx[4] + itemW;
            const baselineY = tx[5];
            const glyphTop = baselineY - fontH * 0.85;
            const glyphBottom = baselineY + fontH * 0.15;
            const centerY = baselineY - fontH * 0.35;

            allPageItems.push({ x: leftX, y: centerY, str: item.str.trim() });

            // Require real glyph box overlap so adjacent lines/questions outside the crop do NOT bleed in
            const yOverlap =
              Math.min(glyphBottom, ry + rh) - Math.max(glyphTop, ry);
            const minRequiredY = Math.min(fontH * 0.35, rh * 0.45);
            const overlapsX = rightX >= rx - 2 && leftX <= rx + rw + 2;
            const overlapsY = yOverlap >= minRequiredY;
            if (overlapsX && overlapsY) {
              itemsInBox.push({
                x: leftX,
                y: centerY,
                h: fontH,
                str: item.str.trim()
              });
            }
          }

          itemsInBox.sort((a, b) =>
            Math.abs(a.y - b.y) < 6 ? a.x - b.x : a.y - b.y
          );
          allPageItems.sort((a, b) =>
            Math.abs(a.y - b.y) < 6 ? a.x - b.x : a.y - b.y
          );

          // Preserve line breaks between distinct rows so next-question boundaries stay clean
          const lines = [];
          let curLine = [];
          let curY = null;
          for (const it of itemsInBox) {
            if (curY === null || Math.abs(it.y - curY) < Math.max(it.h * 0.55, 6)) {
              curLine.push(it.str);
              if (curY === null) curY = it.y;
            } else {
              lines.push(curLine.join(' '));
              curLine = [it.str];
              curY = it.y;
            }
          }
          if (curLine.length) lines.push(curLine.join(' '));

          joinedLinesText = lines
            .map((l) => l.replace(/\s+/g, ' ').trim())
            .filter(Boolean)
            .join('\n');
          joinedText = joinedLinesText.replace(/\n+/g, ' ').trim();
          fullPageText = allPageItems
            .map((l) => l.str)
            .join(' ')
            .replace(/\s+/g, ' ')
            .trim();
        }
      }
    } catch (err) {
      console.warn('PDF text extraction warning:', err);
    }
  }

  // 1. Check High School English Multi-Blank Exam Engine (U15 / U7 / 文意選填 A~J / 克漏字) by coordinates or text!
  const handledByMultiBlankExam = await tryResolveMultiBlankEnglishExam(
    docState,
    rx,
    ry,
    rw,
    rh,
    modeType,
    joinedText,
    fullPageText
  );
  if (handledByMultiBlankExam) {
    return;
  }

  // 2. If PDF text extraction succeeded, handle passage or question+options (with full-page context support)
  if (joinedText) {
    if (modeType === 'passage') {
      applyPassageCropToCurrentQuestion(joinedText, fullPageText);
      return;
    }
    if (
      state.questionType !== 'choice' &&
      tryExtractMultiBlankFromTextAndPage(joinedText, fullPageText, modeType)
    ) {
      return;
    }
    parseQuestionAndOptionsFromRawText(joinedLinesText || joinedText);
    updateSelfLearningCoach();
    runAutoMistakeDiagnosis();
    return;
  }

  // 3. If the document is an uploaded Image (JPG/PNG) or a Scanned/Outlined PDF without text layer, run fast OCR!
  if (state.currentCropDataUrl) {
    await recognizeTextAndMathFromCroppedImage(
      state.currentCropDataUrl,
      docState.fileName || '',
      modeType
    );
    return;
  }

  updateSelfLearningCoach();
  runAutoMistakeDiagnosis();
  showToast('📸 已擷取題目畫面截圖！請點選 STEP 2 的誤選與正解，系統將自動為您分析！');
}

// ============================================================================
// High School English Multi-Blank Exam Engine (文意選填 A~J 十選十 ＆ 綜合測驗克漏字)
// ============================================================================
const ENGLISH_MULTIBLANK_EXAMS = {
  U15_WORD_BANK: {
    unit: '英文1進階卷_U15（III. 文意選填 31~40）',
    wordBank: {
      A: 'flight',
      B: 'managed',
      C: 'spot',
      D: 'craft',
      E: 'destination',
      F: 'impulse',
      G: 'thrilling',
      H: 'missions',
      I: 'everlasting',
      J: 'dazzled'
    },
    passage:
      "The date is June 30, 2021. I am sitting in a six-person __31__ in outer space for my upcoming project. Although I've been flying on choppers for over twenty years, every __32__ still makes me nervous for some reason. I'm not scared though, for I am driven by the __33__ to fly out of curiosity about what I see. This new venture is highly ambitious, with plans to expand to __34__ across Japan, Malaysia, New Zealand, and so on. Below me, the __35__ clouds and deep blue sea provide a breathtaking view. I am __36__ by the sheer beauty of nature. However, it is impossible not to __37__ the visible human activity from this height. These vessels reveal the __38__ power of modern industry. My dream is to make people realize how we have __39__ our ecosystems and inspire them to take action. As long as I can __40__ attention to this, nothing else matters. This is the most vital mission of my life.",
    blanks: {
      '31': {
        ans: 'D',
        sentence:
          'The date is June 30, 2021. I am sitting in a six-person ___31___ in outer space for my upcoming project.',
        zh: '日期是 2021 年 6 月 30 日。我正坐在一艘位於外太空的六人座【太空船／飛行器 (craft)】中，準備進行我接下來的計畫。',
        clue: '空格前有冠詞與複合形容詞「a six-person（六人座的）」且後接地點「in outer space（在外太空）」，因此空格必填可載人的單數名詞「(D) craft（太空船／飛行器）」。'
      },
      '32': {
        ans: 'G',
        sentence:
          "Although I've been flying on choppers for over twenty years, every ___32___ flight/experience still makes me nervous for some reason.",
        zh: '雖然我搭乘直升機飛行已經超過二十年了，但每一次【驚險刺激的 (thrilling)／飛行 (flight)】仍然因為某些原因讓我感到緊張。',
        clue: '前半句提到「have been flying on choppers for over twenty years（飛行超過二十年）」，後半句 every 後接單數名詞或形容詞修飾飛行體驗，對應 (G) thrilling 或 (A) flight。'
      },
      '33': {
        ans: 'F',
        sentence:
          "I'm not scared though, for I am driven by the ___33___ to fly out of curiosity about what I see.",
        zh: '不過我並不害怕，因為我受到想要飛行的【衝動／驅動力 (impulse)】所驅使，對我所見的一切充滿好奇。',
        clue: '空格位於定冠詞「the」與不定詞「to fly」之間，且前面有「driven by（被…驅使）」，片語「the impulse to V」表「做某事的衝動／強烈渴望」，故選 (F) impulse。'
      },
      '34': {
        ans: 'H',
        sentence:
          'This new venture is highly ambitious, with plans to expand to ___34___ across Japan, Malaysia, New Zealand, and so on.',
        zh: '這項新事業極具野心，計畫擴展到遍及日本、馬來西亞、紐西蘭等地的【多項任務 (missions)】。',
        clue: '空格前為「expand to（擴展至…）」，後接跨越多個國家的計畫行動，應填複數名詞「(H) missions（任務／使命）」。'
      },
      '35': {
        ans: 'E',
        sentence:
          'Below me, the ___35___ clouds and deep blue sea provide a breathtaking view.',
        zh: '在我的下方，靜止平靜的雲層與深藍色的大海構成了一幅令人屏息的美景。',
        clue: '空格位於定冠詞「the」與名詞「clouds」之間，描述從高空俯瞰雲層與深海的壯麗景色。'
      },
      '36': {
        ans: 'J',
        sentence:
          'Below me, the [35] clouds and deep blue sea provide a breathtaking view. I am ___36___ by the sheer beauty of nature. However, it is impossible not to [37] the visible human activity from this height.',
        zh: '在我下方，雲層與深藍色大海提供了令人屏息的美景。我被大自然純粹的絕美深深【震撼／使目眩神迷 (dazzled)】。',
        clue: '1. 文法結構：「I am ___36___ by the sheer beauty of nature」為「be + p.p.（過去分詞）+ by」被動／情緒分詞結構！\n2. 前後文線索：前句「breathtaking view（令人屏息的美景）」與本句「the sheer beauty of nature（大自然純粹的絕美）」，代表「我被大自然的美景深深震撼／使目眩神迷」，故唯一正解為 (J) dazzled（be dazzled by 被…使目眩神迷／讚嘆不已）！'
      },
      '37': {
        ans: 'C',
        sentence:
          'However, it is impossible not to ___37___ the visible human activity from this height.',
        zh: '然而，從這個高度，不可能不【看見／注意到 (spot)】那些清晰可見的人類活動痕跡。',
        clue: '空格位於不定詞「not to ___37___」之後，且後接名詞受詞「the visible human activity（肉眼可見的人類活動）」，因此必填「原形及物動詞」！詞庫中唯有「(C) spot（v. 看見、發現、認出）」符合「從高空看見人類活動」的語意與詞性！'
      },
      '38': {
        ans: 'I',
        sentence:
          'These vessels reveal the ___38___ power of modern industry.',
        zh: '這些船隻展現了現代工業【永恆不息／巨大持久的 (everlasting)】力量。',
        clue: '空格位於定冠詞「the」與名詞「power（力量）」之間，必須填入【形容詞 (adj.)】修飾 power，故選 (I) everlasting。'
      },
      '39': {
        ans: 'B',
        sentence:
          'My dream is to make people realize how we have ___39___ our ecosystems and inspire them to take action.',
        zh: '我的夢想是讓人們意識到我們如何【影響／管理或破壞 (managed/damaged)】了我們的生態系統，並激勵他們採取行動。',
        clue: '空格位於現在完成式助動詞「have」之後、受詞「our ecosystems（我們的生態系統）」之前，因此必填【過去分詞 (p.p.)】，對應 (B) managed（或 damaged）。'
      },
      '40': {
        ans: 'A',
        sentence:
          'As long as I can ___40___ attention to this, nothing else matters. This is the most vital mission of my life.',
        zh: '只要我能對此引起關注，其他一切都不重要了。這是我一生中最重要的使命。',
        clue: '搭配後面名詞 attention to this，構成完整子句。'
      }
    }
  },
  U15_CLOZE: {
    unit: '英文1進階卷_U15（II. 綜合測驗 26~30）',
    blanks: {
      '26': {
        ans: 'B',
        sentence:
          '26. She stood alone and stretched to catch a red, heart-shaped balloon that the wind was carrying away. Sadly, the balloon was ___26___ to reach it.',
        optA: 'neither',
        optB: 'other',
        optC: 'another',
        optD: 'each'
      },
      '27': {
        ans: 'B',
        sentence:
          '27. The graffiti artist was Banksy, and his work was called Balloon Girl. Some people believe that the red balloon ___27___ love, hope, and dreams, and that the girl losing her grasp on the balloon represents loss of innocence.',
        optA: 'put off',
        optB: 'cut off',
        optC: 'turn in',
        optD: 'hang on'
      },
      '28': {
        ans: 'C',
        sentence:
          '28. Some people believe that the red balloon represents love, hope, and dreams, and that the girl ___28___ to catch the balloon represents loss of innocence.',
        optA: 'fail',
        optB: 'failed',
        optC: 'failing',
        optD: 'to fail'
      },
      '29': {
        ans: 'D',
        sentence:
          '29. This graffiti uses simple visual language to ___29___ complex social criticisms.',
        optA: 'appeal to',
        optB: 'apply to',
        optC: 'agree to',
        optD: 'result in'
      },
      '30': {
        ans: 'A',
        sentence:
          '30. This graffiti uses simple ___30___ language to express complex social criticisms. It is considered one of the most iconic works of modern street art.',
        optA: 'visible',
        optB: 'visual',
        optC: 'grateful',
        optD: 'eventual'
      }
    }
  },
  U7_WORD_BANK: {
    unit: '英文1進階卷_U7（III. 文意選填 21~30）',
    wordBank: {
      A: 'simple',
      B: 'substance',
      C: 'eager',
      D: 'introduced',
      E: 'express',
      F: 'environmental',
      G: 'ocean',
      H: 'proper',
      I: 'harm',
      J: 'convenient'
    },
    passage:
      "It's no big secret that plastic waste is a huge problem facing the whole world. Luckily, a number of countries have made efforts to reduce the use of this __21__ in recent years. These efforts aim to protect the earth from the __22__ issues caused by plastic. Plastic bags in particular have been a major __23__ of these efforts. In fact, global __24__ show that more than 60 countries have now placed limitations on their use. What have these countries done exactly? First of all, single-use plastic bags are no longer allowed in a number of developing nations, many of which are in Africa. There, the lack of __25__ systems for the management of plastic waste is of great concern. In addition, some developed countries, including France and Italy, have followed the advice of __26__ to stop people from using bags made of certain types of plastic. The main reason is that they can seriously __27__ the environment. Though some other countries still use bags made of these types of plastic, they have __28__ fees or taxes to reduce plastic bag use. Asian countries are also __29__ to take action. Free plastic bags are now __30__ in most stores in Japan, Taiwan, South Korea, and many other places.",
    blanks: {
      '21': {
        ans: 'B',
        sentence:
          "Luckily, a number of countries have made efforts to reduce the use of this ___21___ in recent years.",
        zh: '幸運的是，近年來許多國家已努力減少這種【物質 (substance，指塑膠)】的使用。'
      },
      '22': {
        ans: 'F',
        sentence:
          'These efforts aim to protect the earth from the ___22___ issues caused by plastic.',
        zh: '這些努力旨在保護地球免受塑膠所引起的【環境的 (environmental)】問題侵害。'
      },
      '23': {
        ans: 'G',
        sentence:
          'Plastic bags in particular have been a major ___23___ of these efforts.',
        zh: '塑膠袋尤其成為這些減塑努力的主要目標。'
      },
      '25': {
        ans: 'H',
        sentence:
          'There, the lack of ___25___ systems for the management of plastic waste is of great concern.',
        zh: '在那裡，缺乏管理塑膠廢棄物的【妥善的／適當的 (proper)】系統令人深感憂慮。'
      },
      '27': {
        ans: 'I',
        sentence:
          'The main reason is that they can seriously ___27___ the environment.',
        zh: '主要原因是它們會嚴重【傷害／危害 (harm)】環境（助動詞 can + 副詞 seriously 後接原形動詞 harm）。'
      },
      '28': {
        ans: 'D',
        sentence:
          'Though some other countries still use bags made of these types of plastic, they have ___28___ fees or taxes to reduce plastic bag use.',
        zh: '雖然其他一些國家仍使用這類塑膠袋，但他們已經【推行／引進 (introduced)】收費或徵稅制度來減少塑膠袋使用（have + p.p.）。'
      },
      '29': {
        ans: 'C',
        sentence:
          'Asian countries are also ___29___ to take action.',
        zh: '亞洲國家也【急切／渴望 (be eager to V)】採取行動。'
      }
    }
  }
};

function populateWordBankFields(wordBankMap) {
  if (!wordBankMap) return;
  ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach((letter) => {
    const el = document.getElementById(`fieldOpt${letter}`);
    if (el) {
      el.value = wordBankMap[letter] || '';
    }
  });
}

function renderMultiBlankSwitcherBar(sectionObj, activeQNum) {
  const multiBar = document.getElementById('multiQuestionSelectorBar');
  const multiBtns = document.getElementById('multiQuestionButtons');
  if (!multiBar || !multiBtns || !sectionObj || !sectionObj.blanks) return;

  const blankNums = Object.keys(sectionObj.blanks).sort(
    (a, b) => Number(a) - Number(b)
  );
  if (blankNums.length === 0) {
    multiBar.style.display = 'none';
    return;
  }

  multiBar.style.display = 'block';
  const headerStrong = multiBar.querySelector('strong');
  if (headerStrong) {
    headerStrong.innerHTML = sectionObj.wordBank
      ? `🎯 偵測到「文意選填（第 ${blankNums[0]} ～ ${
          blankNums[blankNums.length - 1]
        } 空，選項 A~J）」多空格題組！點選空格題號即可一鍵切換該題原句與解析：`
      : `🎯 偵測到「克漏字（第 ${blankNums[0]} ～ ${
          blankNums[blankNums.length - 1]
        } 題）」多空格題組！點選題號即可切換：`;
  }

  multiBtns.innerHTML = blankNums
    .map((qNum) => {
      const bInfo = sectionObj.blanks[qNum];
      const isCur = String(qNum) === String(activeQNum);
      const ansTag = bInfo.ans ? ` (${bInfo.ans})` : '';
      return `<button type="button" class="btn btn-sm ${
        isCur ? 'btn-primary' : 'btn-secondary'
      } multi-blank-pick-btn" data-bnum="${escapeHtml(qNum)}">
        第 ${escapeHtml(qNum)} 空${escapeHtml(ansTag)}
      </button>`;
    })
    .join('');

  multiBtns.querySelectorAll('.multi-blank-pick-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      const pickedNum = btn.getAttribute('data-bnum');
      applyMultiBlankQuestionItem(sectionObj, pickedNum);
    });
  });
}

function applyMultiBlankQuestionItem(sectionObj, qNum) {
  if (!sectionObj || !sectionObj.blanks) return;
  const blankNums = Object.keys(sectionObj.blanks).sort(
    (a, b) => Number(a) - Number(b)
  );
  const targetNum = sectionObj.blanks[qNum] ? String(qNum) : blankNums[0];
  const bInfo = sectionObj.blanks[targetNum];
  if (!bInfo) return;

  document.getElementById('fieldSubject').value = '英文';
  if (sectionObj.unit) {
    document.getElementById('fieldUnit').value = sectionObj.unit;
  }
  document.getElementById('fieldQuestionNum').value = targetNum;

  if (sectionObj.wordBank) {
    populateWordBankFields(sectionObj.wordBank);
    setQuestionTypeUI('wordbank');
  } else {
    document.getElementById('fieldOptA').value = bInfo.optA || '';
    document.getElementById('fieldOptB').value = bInfo.optB || '';
    document.getElementById('fieldOptC').value = bInfo.optC || '';
    document.getElementById('fieldOptD').value = bInfo.optD || '';
    const optEEl = document.getElementById('fieldOptE');
    if (optEEl) optEEl.value = '';
    ['F', 'G', 'H', 'I', 'J'].forEach((k) => {
      const el = document.getElementById(`fieldOpt${k}`);
      if (el) el.value = '';
    });
    setQuestionTypeUI('cloze');
  }

  document.getElementById('fieldQuestionText').value = `${targetNum}. ${bInfo.sentence}`;

  // Populate answer key map for all blanks in this section
  blankNums.forEach((n) => {
    if (sectionObj.blanks[n]?.ans) {
      state.answerKeyMap[n] = sectionObj.blanks[n].ans;
    }
  });
  renderAnswerKeyChips();

  if (bInfo.ans) {
    setOptionPillAndInput('correct', bInfo.ans, true);
    const hintEl = document.getElementById('autoAnswerMatchHint');
    if (hintEl) {
      const ansWord = sectionObj.wordBank
        ? sectionObj.wordBank[bInfo.ans] || ''
        : bInfo[`opt${bInfo.ans}`] || '';
      hintEl.textContent = `🔑 已自動帶入第 ${targetNum} 空標準答案：(${bInfo.ans}) ${ansWord}`;
    }
  }

  renderMultiBlankSwitcherBar(sectionObj, targetNum);
  syncSubjectSpecificUI();
  updateSelfLearningCoach(true);
  runAutoMistakeDiagnosis(true);
  showToast(
    `🎯 已載入第 ${targetNum} 空的前後文原句${
      sectionObj.wordBank ? '與 (A)~(J) 十選十詞庫' : '與選項'
    }！請在 STEP 2 點選你當初誤選的選項（A~J），系統將立即解析！`
  );
}

async function tryResolveMultiBlankEnglishExam(
  docState,
  rx,
  ry,
  rw,
  rh,
  modeType,
  joinedText = '',
  fullPageText = ''
) {
  const fileName = docState?.fileName || '';
  const pdfCanvas = document.getElementById('pdfCanvas');
  const canvasW = pdfCanvas?.width || 1;
  const canvasH = pdfCanvas?.height || 1;
  const cxRatio = (rx + rw / 2) / canvasW;
  const cyRatio = (ry + rh / 2) / canvasH;

  const isU15 =
    /U15/i.test(fileName) ||
    /six-person|choppers|sheer beauty of nature|Balloon Girl|Banksy/i.test(
      `${joinedText} ${fullPageText}`
    );
  const isU7 =
    /U7/i.test(fileName) ||
    /Life of a Plastic Bag|plastic waste is a huge problem/i.test(
      `${joinedText} ${fullPageText}`
    );

  // Extract explicit blank number if present in cropped text or already typed by student
  const textNumMatch = joinedText.match(/\b(2[1-9]|3\d|40)\b/);
  const existingQNum = document.getElementById('fieldQuestionNum')?.value.trim();

  if (isU15 && (docState.pageNum === 2 || docState.numPages === 1)) {
    // Check if the crop is in Left Column Middle: III. 文意選填 (31~40) (cxRatio < 0.53, cyRatio 0.30..0.70)
    const isWordBankRegion =
      (cxRatio < 0.54 && cyRatio >= 0.30 && cyRatio <= 0.70) ||
      (textNumMatch &&
        Number(textNumMatch[1]) >= 31 &&
        Number(textNumMatch[1]) <= 40) ||
      (modeType === 'passage' &&
        existingQNum &&
        Number(existingQNum) >= 31 &&
        Number(existingQNum) <= 40);

    if (isWordBankRegion) {
      let targetNum = '36';
      if (
        modeType === 'passage' &&
        existingQNum &&
        Number(existingQNum) >= 31 &&
        Number(existingQNum) <= 40
      ) {
        targetNum = existingQNum;
      } else if (
        textNumMatch &&
        Number(textNumMatch[1]) >= 31 &&
        Number(textNumMatch[1]) <= 40
      ) {
        targetNum = textNumMatch[1];
      } else if (cyRatio >= 0.58 && cyRatio <= 0.70 && rh < canvasH * 0.14) {
        // Infer exact answer slot 31..40 from the 2x5 answer grid at the bottom of Section III!
        const isBottomRow = cyRatio >= 0.635;
        // Left column answer grid spans roughly xRatio 0.06 to 0.48 (width ~0.42, 5 slots of ~0.084 each)
        const colIdx = Math.max(
          0,
          Math.min(4, Math.floor((cxRatio - 0.06) / 0.082))
        );
        targetNum = String((isBottomRow ? 36 : 31) + colIdx);
      } else if (
        existingQNum &&
        Number(existingQNum) >= 31 &&
        Number(existingQNum) <= 40
      ) {
        targetNum = existingQNum;
      }

      // If student cropped a small box (like "36. J"), automatically capture & stitch the Word Bank + Passage above it so they don't even have to crop a second time!
      if (
        modeType !== 'passage' &&
        rh < canvasH * 0.14 &&
        pdfCanvas &&
        state.currentCropDataUrl
      ) {
        try {
          const pX = Math.round(canvasW * 0.055);
          const pY = Math.round(canvasH * 0.325);
          const pW = Math.round(canvasW * 0.44);
          const pH = Math.round(canvasH * 0.275);
          const pCanvas = document.createElement('canvas');
          pCanvas.width = pW;
          pCanvas.height = pH;
          pCanvas
            .getContext('2d')
            .drawImage(pdfCanvas, pX, pY, pW, pH, 0, 0, pW, pH);
          const passageUrl = pCanvas.toDataURL('image/png');
          const stitched = await stitchTwoImagesVertically(
            passageUrl,
            state.currentCropDataUrl
          );
          setCroppedImagePreview(stitched);
        } catch (e) {
          console.warn('Auto-stitch passage preview skipped:', e);
        }
      }

      applyMultiBlankQuestionItem(ENGLISH_MULTIBLANK_EXAMS.U15_WORD_BANK, targetNum);
      return true;
    }

    // Check if the crop is in Left Column Top: II. 綜合測驗 (26~30) (cxRatio < 0.53, cyRatio < 0.30)
    const isCloze26Region =
      (cxRatio < 0.53 && cyRatio >= 0.04 && cyRatio < 0.30) ||
      (textNumMatch &&
        Number(textNumMatch[1]) >= 26 &&
        Number(textNumMatch[1]) <= 30);

    if (isCloze26Region) {
      let targetNum = '26';
      if (
        textNumMatch &&
        Number(textNumMatch[1]) >= 26 &&
        Number(textNumMatch[1]) <= 30
      ) {
        targetNum = textNumMatch[1];
      } else if (
        existingQNum &&
        Number(existingQNum) >= 26 &&
        Number(existingQNum) <= 30
      ) {
        targetNum = existingQNum;
      } else if (cyRatio >= 0.19 && cyRatio < 0.30) {
        const rowIdx = Math.max(
          0,
          Math.min(4, Math.floor((cyRatio - 0.195) / 0.02))
        );
        targetNum = String(26 + rowIdx);
      }
      applyMultiBlankQuestionItem(ENGLISH_MULTIBLANK_EXAMS.U15_CLOZE, targetNum);
      return true;
    }
  }

  if (isU7) {
    // Check if the crop is on U7 Right Column Middle/Bottom: III. 文意選填 (21~30)
    const isU7WordBank =
      (cxRatio >= 0.48 && cyRatio >= 0.32) ||
      (textNumMatch &&
        Number(textNumMatch[1]) >= 21 &&
        Number(textNumMatch[1]) <= 30);
    if (isU7WordBank) {
      const targetNum =
        textNumMatch &&
        Number(textNumMatch[1]) >= 21 &&
        Number(textNumMatch[1]) <= 30
          ? textNumMatch[1]
          : existingQNum &&
            Number(existingQNum) >= 21 &&
            Number(existingQNum) <= 30
          ? existingQNum
          : '21';
      applyMultiBlankQuestionItem(ENGLISH_MULTIBLANK_EXAMS.U7_WORD_BANK, targetNum);
      return true;
    }
  }

  return false;
}

// Universal Multi-Blank Passage & Word Bank Extractor for ANY PDF or OCR text
function parseWordBankFromText(text) {
  if (!text) return null;
  // Look for at least (A)...(B)...(C)...(D)...(E)...(F) up to (J)
  const regex = /[\(（\[]\s*([A-J])\s*[\)）\]]\s*([A-Za-z][A-Za-z\s\-']*?)(?=\s*[\(（\[]\s*[A-J]\s*[\)）\]]|\s*\d{1,2}\s*[\.、]|\s*The\b|\s*It\b|\s*In\b|$)/g;
  const bank = {};
  let m;
  let count = 0;
  while ((m = regex.exec(text)) !== null) {
    const letter = m[1].toUpperCase();
    const word = m[2].trim().split(/\s+/)[0];
    if (word && !bank[letter]) {
      bank[letter] = word;
      count++;
    }
  }
  return count >= 6 ? bank : null;
}

function extractSentencesByBlankNumbers(passageText) {
  if (!passageText) return {};
  // Split passage into sentences and map any blank marker like __36__, _36_, (36), [36], or standalone 31..40 inside sentence
  const sentences = passageText
    .replace(/\s+/g, ' ')
    .split(/(?<=[.!?])\s+/)
    .filter((s) => s.trim().length > 4);

  const blankMap = {};
  sentences.forEach((sent, idx) => {
    const matches =
      sent.match(/(?:_{1,4}\s*(\d{1,2})\s*_{1,4}|[\(\[]\s*(\d{1,2})\s*[\)\]]|\b(1\d|2\d|3\d|40)\b)/g) ||
      [];
    matches.forEach((rawM) => {
      const numM = rawM.match(/\d{1,2}/);
      if (!numM) return;
      const num = numM[0];
      if (Number(num) < 1 || Number(num) > 50) return;
      // Avoid matching years or dates like "June 30, 2021" or "60 countries"
      if (
        new RegExp(`(?:June|July|May|April|more than|over)\\s+${num}\\b`, 'i').test(
          sent
        )
      ) {
        return;
      }
      const prevSent = idx > 0 ? sentences[idx - 1] + ' ' : '';
      const nextSent = idx + 1 < sentences.length ? ' ' + sentences[idx + 1] : '';
      const highlighted = sent.replace(
        new RegExp(`(?:_{1,4}\\s*${num}\\s*_{1,4}|\\b${num}\\b)`),
        `___${num}___`
      );
      blankMap[num] = {
        ans: state.answerKeyMap[num] || '',
        sentence: `${prevSent}${highlighted}${nextSent}`.trim()
      };
    });
  });
  return blankMap;
}

function applyPassageCropToCurrentQuestion(passageText, fullPageText = '') {
  const combined = `${passageText} ${fullPageText}`;
  const detectedBank = parseWordBankFromText(combined);
  if (detectedBank) {
    populateWordBankFields(detectedBank);
    setQuestionTypeUI('wordbank');
  }

  const blankMap = extractSentencesByBlankNumbers(passageText);
  const blankNums = Object.keys(blankMap).sort((a, b) => Number(a) - Number(b));
  const curQNum = document.getElementById('fieldQuestionNum')?.value.trim();

  if (blankNums.length >= 2) {
    const dynSection = {
      unit: document.getElementById('fieldUnit')?.value.trim() || '段考複習卷',
      wordBank: detectedBank,
      blanks: blankMap
    };
    const activeNum = curQNum && blankMap[curQNum] ? curQNum : blankNums[0];
    applyMultiBlankQuestionItem(dynSection, activeNum);
    return;
  }

  // Otherwise prepend or set the cropped passage text while keeping the current question stem/number intact
  const existingStem = document.getElementById('fieldQuestionText').value.trim();
  document.getElementById('fieldQuestionText').value =
    existingStem && !existingStem.includes(passageText.slice(0, 20))
      ? `${passageText}\n${existingStem}`
      : passageText;
  updateSelfLearningCoach(true);
  runAutoMistakeDiagnosis(true);
  showToast('📖 已附加文章段落截圖並帶入上下文敘述！');
}

function tryExtractMultiBlankFromTextAndPage(
  joinedText,
  fullPageText = '',
  modeType = 'question'
) {
  const combined = `${joinedText} ${fullPageText}`;
  const detectedBank = parseWordBankFromText(combined);

  // Check if user cropped a tiny answer blank like "36. J" or "36. ___"
  const shortAnsSlotMatch = joinedText.match(/^\s*(\d{1,2})\s*[\.、．]?\s*([A-Ja-j])?\s*$/);
  if (shortAnsSlotMatch) {
    const qNum = shortAnsSlotMatch[1];
    const ansLetter = (shortAnsSlotMatch[2] || '').toUpperCase();
    const pageBlanks = extractSentencesByBlankNumbers(fullPageText);
    if (detectedBank || Object.keys(pageBlanks).length >= 2) {
      if (ansLetter) state.answerKeyMap[qNum] = ansLetter;
      if (!pageBlanks[qNum]) {
        pageBlanks[qNum] = {
          ans: ansLetter,
          sentence: `第 ${qNum} 題（請參照上方文意選填／克漏字文章空格 ___${qNum}___ 前後文）`
        };
      } else if (ansLetter) {
        pageBlanks[qNum].ans = ansLetter;
      }
      const dynSection = {
        unit: document.getElementById('fieldUnit')?.value.trim() || '段考複習卷',
        wordBank: detectedBank,
        blanks: pageBlanks
      };
      applyMultiBlankQuestionItem(dynSection, qNum);
      return true;
    }
  }

  // Check if cropped text itself is a multi-blank passage or Word Bank (A)~(J)
  const croppedBlanks = extractSentencesByBlankNumbers(joinedText);
  if (detectedBank && Object.keys(croppedBlanks).length >= 2) {
    const curQNum = document.getElementById('fieldQuestionNum')?.value.trim();
    const dynSection = {
      unit: document.getElementById('fieldUnit')?.value.trim() || '段考複習卷',
      wordBank: detectedBank,
      blanks: croppedBlanks
    };
    const firstNum = Object.keys(croppedBlanks).sort((a, b) => Number(a) - Number(b))[0];
    applyMultiBlankQuestionItem(
      dynSection,
      curQNum && croppedBlanks[curQNum] ? curQNum : firstNum
    );
    return true;
  }

  return false;
}

// ============================================================================
// Image OCR & Mathematical Formula Recognition Engine (For JPG/PNG Photos & Scanned PDFs)
// ============================================================================
async function preprocessImageForMathOcr(srcDataUrl, leftPrintOnly = false) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      // If leftPrintOnly is enabled (to ignore right-side figures & handwritten calculations), take left 62% width
      const srcW = leftPrintOnly && img.width > 300 ? Math.round(img.width * 0.62) : img.width;
      const srcH = img.height;
      const scale = srcW < 900 ? 1.8 : 1.2;
      const c = document.createElement('canvas');
      c.width = Math.round(srcW * scale);
      c.height = Math.round(srcH * scale);
      const ctx = c.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, c.width, c.height);
      ctx.drawImage(img, 0, 0, srcW, srcH, 0, 0, c.width, c.height);

      try {
        const imgData = ctx.getImageData(0, 0, c.width, c.height);
        const d = imgData.data;
        // Gently suppress red/blue grading marks and light gray pencil scribbles while keeping dark printed ink smooth!
        for (let i = 0; i < d.length; i += 4) {
          const r = d[i];
          const g = d[i + 1];
          const b = d[i + 2];
          // Filter out red pen and blue pen marks
          if (
            (r > 145 && r - g > 35 && r - b > 35) ||
            (b > 145 && b - r > 35 && b - g > 25)
          ) {
            d[i] = 252;
            d[i + 1] = 252;
            d[i + 2] = 252;
            continue;
          }
          const lum = 0.299 * r + 0.587 * g + 0.114 * b;
          // Printed characters are dark (lum <= 115); pencil calculations & paper shadows are lighter gray (lum > 120)
          let v;
          if (lum <= 105) {
            v = Math.max(0, (lum - 25) * 0.75);
          } else if (lum <= 135) {
            v = 110 + (lum - 105) * 4.2;
          } else {
            v = 255;
          }
          d[i] = v;
          d[i + 1] = v;
          d[i + 2] = v;
        }
        ctx.putImageData(imgData, 0, 0);
      } catch (e) {
        console.warn('Canvas pixel preprocess skipped:', e);
      }
      resolve(c.toDataURL('image/png'));
    };
    img.onerror = () => resolve(srcDataUrl);
    img.src = srcDataUrl;
  });
}

function cleanEnglishOcrText(rawOcr) {
  if (!rawOcr || !rawOcr.trim()) return '';
  return rawOcr
    .replace(/([\u4e00-\u9fa5，。、：；！？（）])\s+([\u4e00-\u9fa5，。、：；！？（）])/g, '$1$2')
    .replace(/([\u4e00-\u9fa5，。、：；！？（）])\s+([\u4e00-\u9fa5，。、：；！？（）])/g, '$1$2')
    .replace(/\s+/g, ' ')
    .trim();
}

async function recognizeTextAndMathFromCroppedImage(
  croppedDataUrl,
  fileName = '',
  modeType = 'question',
  leftPrintOnly = false
) {
  const statusBar = document.getElementById('ocrStatusBar');
  if (statusBar) {
    statusBar.style.display = 'block';
    statusBar.style.background = '#eef2ff';
    statusBar.style.color = '#312e81';
    statusBar.textContent =
      '🔍 正在從截圖自動辨識英文題目與選項（支援單字、克漏字、文意選填、篇章結構、閱讀與翻譯）...';
  }

  document.getElementById('fieldSubject').value = '英文';

  // 1. If user configured a Gemini API key, try Gemini Vision first for high-precision English OCR
  const savedGeminiKey = localStorage.getItem('mistake_gemini_api_key') || '';
  if (savedGeminiKey && croppedDataUrl.startsWith('data:image/')) {
    try {
      const base64Data = croppedDataUrl.split(',')[1];
      const resp = await fetch(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(
          savedGeminiKey
        )}`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [
              {
                parts: [
                  {
                    text: '請辨識這張英文試卷截圖中的「題目與選項」，並回傳純 JSON：{"qNum":"題號","questionType":"choice或cloze或wordbank或structure或reading或translation","questionText":"英文題幹或文章句子","optA":"(A)內容","optB":"(B)內容","optC":"(C)內容","optD":"(D)內容","optE":"(E)內容","optF":"(F)內容"}'
                  },
                  { inlineData: { mimeType: 'image/png', data: base64Data } }
                ]
              }
            ]
          })
        }
      );
      if (resp.ok) {
        const data = await resp.json();
        const rawReply = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
        const jsonMatch = rawReply.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const parsed = JSON.parse(jsonMatch[0]);
          if (parsed.questionText) {
            if (parsed.qNum) {
              document.getElementById('fieldQuestionNum').value = parsed.qNum;
            }
            document.getElementById('fieldQuestionText').value =
              parsed.questionText;
            if (parsed.questionType) {
              setQuestionTypeUI(parsed.questionType);
            }
            ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach((k) => {
              const el = document.getElementById(`fieldOpt${k}`);
              if (el && parsed[`opt${k}`]) el.value = parsed[`opt${k}`];
            });
            if (statusBar) {
              statusBar.style.background = '#ecfdf5';
              statusBar.style.color = '#047857';
              statusBar.textContent =
                '✅ 已透過 AI 視覺成功辨識截圖中的英文題目與選項！';
            }
            updateSelfLearningCoach();
            runAutoMistakeDiagnosis(true);
            return;
          }
        }
      }
    } catch (e) {
      console.warn('Gemini OCR fallback to local Tesseract:', e);
    }
  }

  // 2. Built-in Local Browser OCR via Tesseract.js (eng+chi_tra)
  if (window.Tesseract) {
    try {
      const processedUrl = await preprocessImageForMathOcr(
        croppedDataUrl,
        leftPrintOnly
      );
      const result = await window.Tesseract.recognize(
        processedUrl,
        'eng+chi_tra',
        {
          logger: (m) => {
            if (statusBar && m.status === 'recognizing text') {
              const pct = Math.round((m.progress || 0) * 100);
              statusBar.textContent = `🔍 正在辨識英文試卷文字與選項... ${pct}%`;
            }
          }
        }
      );

      const cleanedText = cleanEnglishOcrText(result?.data?.text || '');
      if (cleanedText && cleanedText.length >= 2) {
        if (modeType === 'passage') {
          applyPassageCropToCurrentQuestion(cleanedText, '');
        } else {
          parseQuestionAndOptionsFromRawText(cleanedText);
        }
        if (statusBar) {
          statusBar.style.background = '#ecfdf5';
          statusBar.style.color = '#047857';
          statusBar.textContent = '✅ 已從截圖自動辨識出英文題目與選項！';
        }
        updateSelfLearningCoach();
        runAutoMistakeDiagnosis(true);
        showToast('🔍 已從試卷截圖自動辨識英文題目並啟動三階段引導！');
        return;
      }
    } catch (ocrErr) {
      console.warn('Tesseract OCR error:', ocrErr);
    }
  }

  if (statusBar) {
    statusBar.style.background = '#fffbeb';
    statusBar.style.color = '#b45309';
    statusBar.textContent =
      '💡 已擷取題目截圖！您可直接點選題型並在 STEP 2 輸入誤選與正解，系統將即時產生英文解析！';
  }
  updateSelfLearningCoach();
  runAutoMistakeDiagnosis(true);
}

// Smart English Question Parser supporting:
// - 單字選擇 (choice)
// - 克漏字多列選項 (cloze)
// - 文意選填 A~J (wordbank)
// - 篇章結構 A~E / A~F 句子選項 (structure)
// - 閱讀測驗 (reading)
// - 中翻英翻譯題 (translation)
// When a crop/paste accidentally includes subsequent questions (e.g. Q4, Q5, Section II...),
// isolate just the target single question so Q4+ never bleeds into stem or Option D.
function isolateSingleQuestionText(rawText) {
  const normalized = String(rawText || '')
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .trim();
  if (!normalized) return '';

  const typedQNum = document.getElementById('fieldQuestionNum')?.value.trim();
  const startMatch = normalized.match(
    /(?:^|\n)\s*(?:[\(（\[]\s*[\)）\]]\s*)?(\d{1,3})\s*[\.、．\)]/
  );
  const firstNum = startMatch ? Number(startMatch[1]) : NaN;
  const baseNum = !Number.isNaN(firstNum)
    ? firstNum
    : /^\d{1,3}$/.test(typedQNum || '')
    ? Number(typedQNum)
    : NaN;

  // If user specifically entered a question number that starts slightly inside the text, jump to it
  let text = normalized;
  if (/^\d{1,3}$/.test(typedQNum || '') && Number(typedQNum) !== firstNum) {
    const targetRe = new RegExp(
      `(?:^|\\n|\\s)(?:[\\(（\\[]\\s*[\\)）\\]]\\s*)?${typedQNum}\\s*[\\.、．\\)]\\s`
    );
    const tm = targetRe.exec(normalized);
    if (tm && tm.index >= 0) {
      text = normalized.slice(tm.index).trim();
    }
  }

  // Find where the next question or next exam section starts
  const afterStartOffset = Math.min(text.length, 6);
  const tail = text.slice(afterStartOffset);
  const stopPositions = [];

  // 1. Line-start next question number (e.g. "\n4. " or "\n( ) 4. ")
  const lineNextQRe =
    /\n\s*(?:[\(（\[]\s*[\)）\]]\s*)?(\d{1,3})\s*[\.、．\)]\s+/g;
  let lm;
  while ((lm = lineNextQRe.exec(tail)) !== null) {
    const n = Number(lm[1]);
    if (Number.isNaN(baseNum) || (n > baseNum && n <= baseNum + 15)) {
      stopPositions.push(afterStartOffset + lm.index);
      break;
    }
  }

  // 2. Inline next question number right after options (e.g. "... (D) potential 4. The ..." or "... ( ) 4. ...")
  if (!Number.isNaN(baseNum)) {
    const inlineNextRe = new RegExp(
      `\\s+(?:[\\(（\\[]\\s*[\\)）\\]]\\s*)?${baseNum + 1}\\s*[\\.、．\\)]\\s+(?=[A-Za-z\\u4e00-\\u9fa5_\\(（\\[])`
    );
    const im = inlineNextRe.exec(tail);
    if (im) stopPositions.push(afterStartOffset + im.index);
  }

  // 3. Inline next section heading (II. / III. / IV. / V. / 第X大題 / Part)
  const sectionRe =
    /(?:^|\n|\s)(?:[IIVX]{2,4}\s*[\.、．]\s*(?:綜合測驗|克漏字|文意選填|篇章結構|閱讀測驗|翻譯|選擇)|第[一二三四五六]大題|Part\s+[IIVX\d]+\s*[:：\.])/i;
  const sm = sectionRe.exec(tail);
  if (sm) stopPositions.push(afterStartOffset + sm.index);

  if (stopPositions.length > 0) {
    const cutAt = Math.min(...stopPositions);
    if (cutAt > 12) {
      return text.slice(0, cutAt).trim();
    }
  }
  return text;
}

function parseQuestionAndOptionsFromRawText(rawText) {
  const multiBar = document.getElementById('multiQuestionSelectorBar');
  const multiBtns = document.getElementById('multiQuestionButtons');
  if (multiBar) multiBar.style.display = 'none';

  if (!rawText || !rawText.trim()) return;
  const prevQType = state.questionType || 'choice';
  const linePreserved = String(rawText).replace(/\r\n?/g, '\n').trim();
  const cleaned = linePreserved.replace(/\s+/g, ' ').trim();
  document.getElementById('fieldSubject').value = '英文';

  // 0. Check if text contains a Word Bank (A)~(J) (only when user isn't explicitly on single choice, or there are actually >= 6 bank words A~J)
  if (prevQType !== 'choice') {
    const detectedBank = parseWordBankFromText(cleaned);
    if (detectedBank) {
      populateWordBankFields(detectedBank);
      setQuestionTypeUI('wordbank');
      const blankMap = extractSentencesByBlankNumbers(cleaned);
      const blankNums = Object.keys(blankMap).sort((a, b) => Number(a) - Number(b));
      if (blankNums.length > 0) {
        applyMultiBlankQuestionItem(
          {
            unit: document.getElementById('fieldUnit')?.value || '英文段考複習卷',
            wordBank: detectedBank,
            blanks: blankMap
          },
          blankNums[0]
        );
        return;
      }
    }
  }

  // For single-question types (單字選擇 / 閱讀 / 翻譯), isolate ONLY the target question so Q4+ never leaks in!
  const shouldIsolateSingle =
    prevQType === 'choice' || prevQType === 'reading' || prevQType === 'translation';
  const workingLineText = shouldIsolateSingle
    ? isolateSingleQuestionText(linePreserved)
    : linePreserved;
  const workingFlat = workingLineText.replace(/\s+/g, ' ').trim();

  // 1. Parse numbered question(s) with (A)-(D) options (supports single vocabulary Q OR multiple cloze rows)
  const multiQRegex =
    /(?:[\(（\[]\s*[\)）\]]\s*)?(\d{1,3})\s*[\.、．\)]\s*([^]*?[\(（\[][A-Da-d][\)）\]\.、\s][^]*?)(?=(?:[\(（\[]\s*[\)）\]]\s*)?\d{1,3}\s*[\.、．\)]\s*[^]*?(?:[\(（\[][A-Da-d][\)）\]\.、\s]|$)|$)/g;

  const detectedQuestions = [];
  let m;
  while ((m = multiQRegex.exec(workingFlat)) !== null) {
    const qNum = m[1];
    const body = m[2].trim();
    const parsedOpts = splitStemAndFourOptions(body, Number(qNum));
    if (parsedOpts) {
      detectedQuestions.push({
        qNum,
        stem: parsedOpts.stem,
        optA: parsedOpts.optA,
        optB: parsedOpts.optB,
        optC: parsedOpts.optC,
        optD: parsedOpts.optD,
        optE: parsedOpts.optE || '',
        optF: parsedOpts.optF || ''
      });
    }
    if (shouldIsolateSingle && detectedQuestions.length >= 1) break;
  }

  if (detectedQuestions.length > 0) {
    const firstQ = detectedQuestions[0];
    const isClozeRow = !firstQ.stem || firstQ.stem.length < 15;
    const isStructureOpts =
      Boolean(firstQ.optE) ||
      [firstQ.optA, firstQ.optB, firstQ.optC, firstQ.optD].every(
        (s) => (s || '').split(/\s+/).length >= 5
      );
    const isReadingQ =
      /according to|passage|article|infer|main idea|title|author|paragraph|what is the|which of the following/i.test(
        firstQ.stem
      );

    const nextQType = isStructureOpts
      ? 'structure'
      : isReadingQ
      ? 'reading'
      : prevQType === 'choice' && !isClozeRow
      ? 'choice'
      : isClozeRow || detectedQuestions.length > 1
      ? 'cloze'
      : 'choice';

    setQuestionTypeUI(nextQType);
    applyParsedQuestionObject(firstQ);

    if (detectedQuestions.length > 1 && multiBar && multiBtns) {
      multiBar.style.display = 'block';
      multiBtns.innerHTML = detectedQuestions
        .map(
          (q, idx) =>
            `<button type="button" class="btn btn-sm ${
              idx === 0 ? 'btn-primary' : 'btn-secondary'
            } multi-q-pick-btn" data-qidx="${idx}">
              第 ${escapeHtml(q.qNum)} 題： (A) ${escapeHtml(q.optA.slice(0, 12))}...
            </button>`
        )
        .join('');

      multiBtns.querySelectorAll('.multi-q-pick-btn').forEach((btn) => {
        btn.addEventListener('click', () => {
          const idx = Number(btn.getAttribute('data-qidx'));
          multiBtns.querySelectorAll('.multi-q-pick-btn').forEach((b) => {
            b.classList.remove('btn-primary');
            b.classList.add('btn-secondary');
          });
          btn.classList.remove('btn-secondary');
          btn.classList.add('btn-primary');
          applyParsedQuestionObject(detectedQuestions[idx]);
          runAutoMistakeDiagnosis(true);
        });
      });

      showToast(
        `✂️ 框選範圍內偵測到 ${detectedQuestions.length} 題（第 ${detectedQuestions
          .map((d) => d.qNum)
          .join('、')} 題），點選按鈕即可切換！`
      );
    } else {
      showToast(
        `✂️ 已精準擷取第 ${firstQ.qNum} 題與選項！`
      );
    }
    return;
  }

  // 2. Fallback: Single question or Discourse Structure (A)~(E)/(F) options
  const singleText =
    prevQType === 'structure' ? cleaned : isolateSingleQuestionText(linePreserved).replace(/\s+/g, ' ').trim();
  const singleNumM = singleText.match(/^\s*[\(（\[]?\s*(\d{1,3})\s*[\)）\]\.、．\s]/);
  const singleParsed = splitStemAndFourOptions(
    singleText,
    singleNumM ? Number(singleNumM[1]) : NaN
  );
  if (singleParsed) {
    const isStructureLike =
      prevQType === 'structure' ||
      Boolean(singleParsed.optE || singleParsed.optF) ||
      [singleParsed.optA, singleParsed.optB, singleParsed.optC, singleParsed.optD].every(
        (s) => (s || '').split(/\s+/).length >= 6
      );
    const isReadingLike =
      !isStructureLike &&
      ((singleParsed.stem && singleParsed.stem.length > 180) ||
        /according to|passage|article|infer|main idea|title|paragraph|which of the following/i.test(
          singleParsed.stem
        ));
    setQuestionTypeUI(
      isStructureLike
        ? 'structure'
        : isReadingLike
        ? 'reading'
        : prevQType === 'cloze'
        ? 'cloze'
        : 'choice'
    );
    if (singleParsed.stem) {
      document.getElementById('fieldQuestionText').value = singleParsed.stem;
      const qNumMatch = singleParsed.stem.match(/^\s*[\(（\[]?(\d{1,3})[\)）\]\.、．\s]/);
      if (qNumMatch) {
        document.getElementById('fieldQuestionNum').value = qNumMatch[1];
        autoFillAnswerFromKeyMap(qNumMatch[1]);
      }
    }
    document.getElementById('fieldOptA').value = singleParsed.optA;
    document.getElementById('fieldOptB').value = singleParsed.optB;
    document.getElementById('fieldOptC').value = singleParsed.optC;
    document.getElementById('fieldOptD').value = singleParsed.optD;
    const optEEl = document.getElementById('fieldOptE');
    if (optEEl) optEEl.value = singleParsed.optE || '';
    const optFEl = document.getElementById('fieldOptF');
    if (optFEl) optFEl.value = singleParsed.optF || '';
    showToast(
      isStructureLike
        ? '📑 已自動辨識為「篇章結構」並拆解句子選項！'
        : isReadingLike
        ? '📖 已自動辨識為「閱讀測驗」並拆解題幹與選項！'
        : '✂️ 已自動拆解英文題幹與 (A)~(D) 選項！'
    );
  } else {
    document.getElementById('fieldQuestionText').value = singleText;
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach((k) => {
      const el = document.getElementById(`fieldOpt${k}`);
      if (el) el.value = '';
    });

    const qNumMatch = singleText.match(/^\s*[\(（\[]?(\d{1,3})[\)）\]\.、．\s]/);
    if (qNumMatch) {
      document.getElementById('fieldQuestionNum').value = qNumMatch[1];
      autoFillAnswerFromKeyMap(qNumMatch[1]);
    }

    // If it contains Chinese characters (like a Chinese prompt for translation), switch to translation!
    const hasChineseSentence = (singleText.match(/[\u4e00-\u9fa5]/g) || []).length >= 5;
    if (hasChineseSentence) {
      setQuestionTypeUI('translation');
      showToast('✍️ 偵測到中文翻譯題幹，已自動切換為「翻譯題 (中翻英)」模式！');
    } else {
      setQuestionTypeUI(prevQType === 'cloze' ? 'cloze' : 'choice');
      showToast('🔤 已帶入英文題幹敘述！');
    }
  }
}

function splitStemAndFourOptions(text, currentQNum = NaN) {
  const cleanOpt = (s) => {
    let out = String(s || '')
      .replace(/\s*[\(（\[]\s*[\)）\]]\s*\d{1,3}\s*[\.、．\)][\s\S]*$/, '')
      .replace(/\s+(?:[IIVX]{2,4}\s*[\.、．]\s*(?:綜合測驗|克漏字|文意選填|篇章結構|閱讀測驗|翻譯)|第[一二三四五六]大題|Part\s+[IIVX\d]+\s*[:：\.])[\s\S]*$/i, '')
      .trim();
    if (!Number.isNaN(currentQNum)) {
      out = out
        .replace(
          new RegExp(`\\s+${currentQNum + 1}\\s*[\\.、．\\)]\\s+(?=[A-Za-z\\u4e00-\\u9fa5_\\(（\\[])[\\s\\S]*$`),
          ''
        )
        .trim();
    } else {
      out = out
        .replace(/\s+\d{1,2}\s*[\.、．\)]\s+(?=[A-Z\u4e00-\u9fa5_\(（\[])[\s\S]*$/, '')
        .trim();
    }
    return out;
  };

  // 1. Try 6-option format (Discourse Structure A~F): (A) ... (B) ... (C) ... (D) ... (E) ... (F) ...
  const opt6LetterRegex =
    /^(.*?)\s*[\(（\[]\s*A\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*B\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*C\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*D\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*E\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*F\s*[\)）\]]\s*(.+)$/is;
  const m6L = text.match(opt6LetterRegex);
  if (m6L) {
    return {
      stem: m6L[1].replace(/^[\(（\[]\s*[\)）\]]\s*/, '').trim(),
      optA: cleanOpt(m6L[2]),
      optB: cleanOpt(m6L[3]),
      optC: cleanOpt(m6L[4]),
      optD: cleanOpt(m6L[5]),
      optE: cleanOpt(m6L[6]),
      optF: cleanOpt(m6L[7])
    };
  }

  // 2. Try 5-option format (Discourse Structure A~E): (A) ... (B) ... (C) ... (D) ... (E) ...
  const opt5LetterRegex =
    /^(.*?)\s*[\(（\[]\s*A\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*B\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*C\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*D\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*E\s*[\)）\]]\s*(.+)$/is;
  const m5L = text.match(opt5LetterRegex);
  if (m5L) {
    return {
      stem: m5L[1].replace(/^[\(（\[]\s*[\)）\]]\s*/, '').trim(),
      optA: cleanOpt(m5L[2]),
      optB: cleanOpt(m5L[3]),
      optC: cleanOpt(m5L[4]),
      optD: cleanOpt(m5L[5]),
      optE: cleanOpt(m5L[6]),
      optF: ''
    };
  }

  // 3. Try 4-option format: (A) ... (B) ... (C) ... (D) ...
  const opt4LetterRegex =
    /^(.*?)\s*[\(（\[]\s*A\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*B\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*C\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*D\s*[\)）\]]\s*(.+)$/is;
  const m4L = text.match(opt4LetterRegex);
  if (m4L) {
    return {
      stem: m4L[1].replace(/^[\(（\[]\s*[\)）\]]\s*/, '').trim(),
      optA: cleanOpt(m4L[2]),
      optB: cleanOpt(m4L[3]),
      optC: cleanOpt(m4L[4]),
      optD: cleanOpt(m4L[5]),
      optE: '',
      optF: ''
    };
  }

  // 4. Try "A. ... B. ... C. ... D. ..." or "A) ... B) ... C) ... D) ..." format
  const opt4DotRegex =
    /^(.*?)\s+A[\.\)]\s+(.+?)\s+B[\.\)]\s+(.+?)\s+C[\.\)]\s+(.+?)\s+D[\.\)]\s+(.+)$/;
  const m4D = text.match(opt4DotRegex);
  if (m4D) {
    return {
      stem: m4D[1].replace(/^[\(（\[]\s*[\)）\]]\s*/, '').trim(),
      optA: cleanOpt(m4D[2]),
      optB: cleanOpt(m4D[3]),
      optC: cleanOpt(m4D[4]),
      optD: cleanOpt(m4D[5]),
      optE: '',
      optF: ''
    };
  }

  return null;
}

function applyParsedQuestionObject(qObj) {
  document.getElementById('fieldQuestionNum').value = qObj.qNum;
  autoFillAnswerFromKeyMap(qObj.qNum);

  const currentStem = document.getElementById('fieldQuestionText').value.trim();
  if (qObj.stem && qObj.stem.length > 6) {
    document.getElementById('fieldQuestionText').value = `${qObj.qNum}. ${qObj.stem}`;
  } else if (!currentStem || /^\(?\)?\s*\d+\.?$/.test(currentStem) || currentStem.includes('(A)')) {
    document.getElementById(
      'fieldQuestionText'
    ).value = `第 ${qObj.qNum} 題（請參照試卷截圖上下文選出正確選項）`;
  }

  document.getElementById('fieldOptA').value = qObj.optA;
  document.getElementById('fieldOptB').value = qObj.optB;
  document.getElementById('fieldOptC').value = qObj.optC;
  document.getElementById('fieldOptD').value = qObj.optD;
  const optEEl = document.getElementById('fieldOptE');
  if (optEEl) optEEl.value = qObj.optE || '';
}

// ============================================================================
// STEP 1B: Standard Answer Key Quick Drawer
// ============================================================================
function initAnswerKeyDrawer() {
  const drawer = document.getElementById('answerKeyDrawer');
  document.getElementById('btnToggleAnswerDrawer').addEventListener('click', () => {
    drawer.style.display = drawer.style.display === 'none' ? 'block' : 'none';
  });
  document.getElementById('btnCloseAnswerDrawer').addEventListener('click', () => {
    drawer.style.display = 'none';
  });

  document.getElementById('btnApplyBulkAnswers').addEventListener('click', () => {
    const raw = document.getElementById('inputBulkAnswers').value.trim();
    if (!raw) return;
    parseBulkAnswerString(raw);
    renderAnswerKeyChips();
    saveToLocalStorage();

    const currentQNum = document.getElementById('fieldQuestionNum').value.trim();
    if (currentQNum) autoFillAnswerFromKeyMap(currentQNum);
    runAutoMistakeDiagnosis();
    showToast(`🔑 已更新標準答案表（共 ${Object.keys(state.answerKeyMap).length} 題）！`);
  });
}

function parseBulkAnswerString(raw) {
  const map = {};
  if (raw.includes(':') || raw.includes('：') || raw.includes('=')) {
    const pairs = raw.split(/[,，;；\s]+/);
    pairs.forEach((p) => {
      const m = p.match(/^(\d+)[:：=]([A-Za-z0-9]+)$/);
      if (m) {
        map[m[1]] = m[2].toUpperCase();
      }
    });
  } else {
    const letters = raw.replace(/[^A-Za-z]/g, '').toUpperCase().split('');
    letters.forEach((ch, idx) => {
      map[String(idx + 1)] = ch;
    });
  }
  state.answerKeyMap = map;
}

function renderAnswerKeyChips() {
  const container = document.getElementById('answerKeyChipsContainer');
  const keys = Object.keys(state.answerKeyMap).sort((a, b) => Number(a) - Number(b));
  if (keys.length === 0) {
    container.innerHTML =
      '<span style="font-size:0.78rem; color:#94a3b8;">尚未設定標準答案</span>';
    return;
  }
  container.innerHTML = keys
    .map(
      (k) =>
        `<span class="ans-chip">第 ${k} 題：<b>${state.answerKeyMap[k]}</b></span>`
    )
    .join('');
}

function autoFillAnswerFromKeyMap(qNum) {
  const cleanNum = String(qNum).replace(/\D/g, '');
  const hintEl = document.getElementById('autoAnswerMatchHint');
  if (cleanNum && state.answerKeyMap[cleanNum]) {
    const ans = state.answerKeyMap[cleanNum];
    setOptionPillAndInput('correct', ans);
    hintEl.textContent = `✨ 已從答案表自動帶入第 ${cleanNum} 題正解：(${ans})`;
  } else {
    hintEl.textContent = '';
  }
}

// ============================================================================
// STEP 1C: Automatic Mistake Diagnosis Engine (Zero Typing Needed for Students!)
// ============================================================================
function initMistakeForm() {
  // Question Type Switcher Buttons
  document
    .querySelectorAll('#questionTypeContainer .diagnosis-tag')
    .forEach((btn) => {
      btn.addEventListener('click', () => {
        const qType = btn.getAttribute('data-qtype');
        setQuestionTypeUI(qType);
        runAutoMistakeDiagnosis(true);
      });
    });

  // Remove cropped image button
  document.getElementById('btnRemoveCropImg').addEventListener('click', () => {
    setCroppedImagePreview('');
    const statusBar = document.getElementById('ocrStatusBar');
    if (statusBar) statusBar.style.display = 'none';
  });

  // Re-run OCR on cropped image button
  document.getElementById('btnRerunOcrOnCrop')?.addEventListener('click', async () => {
    if (!state.currentCropDataUrl) {
      alert('請先在左側試卷框選題目範圍！');
      return;
    }
    const docState = state.docs[state.activeDocSlot];
    await recognizeTextAndMathFromCroppedImage(
      state.currentCropDataUrl,
      docState?.fileName || '',
      'question',
      false
    );
  });


  // Initialize Autonomous Self-Directed Learning Coach controls & Subject UI
  initSelfLearningCoach();
  syncSubjectSpecificUI();

  // Auto-split stem and options button
  document.getElementById('btnParseRawQuestion').addEventListener('click', () => {
    const raw = document.getElementById('fieldQuestionText').value;
    parseQuestionAndOptionsFromRawText(raw);
    syncSubjectSpecificUI();
    updateSelfLearningCoach();
    runAutoMistakeDiagnosis(true);
    showToast('✨ 已整理題目格式並更新自主學習引導與診斷！');
  });

  // Auto-lookup standard answer & multi-blank passage when typing question number
  document.getElementById('fieldQuestionNum').addEventListener('input', (e) => {
    const typedNum = e.target.value.trim();
    autoFillAnswerFromKeyMap(typedNum);
    const activeDoc = state.docs[state.activeDocSlot];
    const fn = `${activeDoc?.fileName || ''} ${
      document.getElementById('fieldUnit')?.value || ''
    }`;
    if (/U15/i.test(fn)) {
      if (ENGLISH_MULTIBLANK_EXAMS.U15_WORD_BANK.blanks[typedNum]) {
        applyMultiBlankQuestionItem(
          ENGLISH_MULTIBLANK_EXAMS.U15_WORD_BANK,
          typedNum
        );
        return;
      }
      if (ENGLISH_MULTIBLANK_EXAMS.U15_CLOZE.blanks[typedNum]) {
        applyMultiBlankQuestionItem(
          ENGLISH_MULTIBLANK_EXAMS.U15_CLOZE,
          typedNum
        );
        return;
      }
    } else if (/U7/i.test(fn)) {
      if (ENGLISH_MULTIBLANK_EXAMS.U7_WORD_BANK.blanks[typedNum]) {
        applyMultiBlankQuestionItem(
          ENGLISH_MULTIBLANK_EXAMS.U7_WORD_BANK,
          typedNum
        );
        return;
      }
    }
    updateSelfLearningCoach();
    runAutoMistakeDiagnosis(true);
  });

  [
    'fieldQuestionText',
    'fieldOptA',
    'fieldOptB',
    'fieldOptC',
    'fieldOptD',
    'fieldOptE',
    'fieldOptF',
    'fieldOptG',
    'fieldOptH',
    'fieldOptI',
    'fieldOptJ',
    'fieldSubject'
  ].forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', () => {
        if (id === 'fieldSubject') {
          syncSubjectSpecificUI();
        } else {
          // If the user typed or pasted an English question/options while another subject was selected, auto-sync to English!
          autoDetectSubjectFromInputs();
        }
        updateSelfLearningCoach();
        runAutoMistakeDiagnosis(true);
      });
    }
  });

  // Wrong & Correct Option Pills -> Support Single-Choice, Word Bank (A~J), AND Multi-Select (1~5 / A~E) toggling!
  document.querySelectorAll('#wrongOptionPills .option-pill').forEach((btn) => {
    btn.addEventListener('click', () => {
      handleOptionPillClick('wrong', btn.getAttribute('data-val'));
      runAutoMistakeDiagnosis(true);
    });
  });

  document.querySelectorAll('#correctOptionPills .option-pill').forEach((btn) => {
    btn.addEventListener('click', () => {
      handleOptionPillClick('correct', btn.getAttribute('data-val'));
      runAutoMistakeDiagnosis(true);
    });
  });

  document.getElementById('fieldWrongAnswer').addEventListener('input', (e) => {
    const val = e.target.value;
    if (state.questionType !== 'translation') {
      setOptionPillAndInput('wrong', val, false);
    }
    runAutoMistakeDiagnosis(true);
  });

  document.getElementById('fieldCorrectAnswer').addEventListener('input', (e) => {
    const val = e.target.value;
    if (state.questionType !== 'translation') {
      setOptionPillAndInput('correct', val, false);
    }
    runAutoMistakeDiagnosis(true);
  });

  // Error Diagnosis Tags
  document.querySelectorAll('#errorTagsContainer .diagnosis-tag').forEach((tagBtn) => {
    tagBtn.addEventListener('click', () => {
      const tag = tagBtn.getAttribute('data-tag');
      if (state.selectedErrorTags.has(tag)) {
        state.selectedErrorTags.delete(tag);
        tagBtn.classList.remove('selected');
      } else {
        state.selectedErrorTags.add(tag);
        tagBtn.classList.add('selected');
      }
      runAutoMistakeDiagnosis(true);
    });
  });

  // Toggle Manual Edit Section
  const btnToggleManual = document.getElementById('btnToggleManualEdit');
  if (btnToggleManual) {
    btnToggleManual.addEventListener('click', () => {
      const sec = document.getElementById('manualEditSection');
      sec.style.display = sec.style.display === 'none' ? 'block' : 'none';
    });
  }

  // Sync manual edits back to preview card live
  ['fieldWhyWrong', 'fieldCorrectConcept', 'fieldKeyTakeaway'].forEach((fieldId) => {
    const el = document.getElementById(fieldId);
    if (!el) return;
    el.addEventListener('input', () => {
      document.getElementById('previewWhyWrong').textContent =
        document.getElementById('fieldWhyWrong').value;
      document.getElementById('previewCorrectConcept').textContent =
        document.getElementById('fieldCorrectConcept').value;
      document.getElementById('previewKeyTakeaway').textContent =
        document.getElementById('fieldKeyTakeaway').value;
    });
  });

  // Auto-Analyze Button (runs built-in smart analyzer + Gemini Vision if API Key set, else Free AI LLM enhancer)
  const btnAutoAnalyze = document.getElementById('btnAutoAnalyzeMistake');
  if (btnAutoAnalyze) {
    btnAutoAnalyze.addEventListener('click', async () => {
      runAutoMistakeDiagnosis(true);
      const savedKey = localStorage.getItem('mistake_gemini_api_key') || '';
      if (savedKey) {
        await runGeminiVisionDiagnosis(savedKey);
      } else {
        await runFreeAiPollinationsDiagnosis();
      }
    });
  }

  // Optional Gemini API Key Drawer
  const btnToggleAiKey = document.getElementById('btnToggleAiKeyBox');
  if (btnToggleAiKey) {
    const savedKey = localStorage.getItem('mistake_gemini_api_key') || '';
    const keyInput = document.getElementById('inputGeminiApiKey');
    if (keyInput && savedKey) keyInput.value = savedKey;

    btnToggleAiKey.addEventListener('click', () => {
      const drawer = document.getElementById('aiKeyDrawer');
      drawer.style.display = drawer.style.display === 'none' ? 'block' : 'none';
    });

    document.getElementById('btnSaveGeminiKey')?.addEventListener('click', () => {
      const k = document.getElementById('inputGeminiApiKey').value.trim();
      localStorage.setItem('mistake_gemini_api_key', k);
      document.getElementById('aiKeyDrawer').style.display = 'none';
      showToast(k ? '🔑 已儲存 Gemini API Key！點「🪄 自動分析」將啟用 AI 截圖視覺深度解析。' : '已清除 API Key，將使用內建智慧題型診斷引擎。');
    });
  }

  // Reset Form
  document.getElementById('btnResetForm').addEventListener('click', () => {
    resetMistakeForm();
  });

  // Submit Mistake Form
  document.getElementById('mistakeForm').addEventListener('submit', () => {
    saveCurrentMistakeForm(false);
  });

  // Save and Immediately Practice
  document.getElementById('btnSaveAndPractice').addEventListener('click', () => {
    const ok = saveCurrentMistakeForm(true);
    if (ok) {
      switchTab('tab-practice');
      startPracticeQuiz();
    }
  });
}

// ============================================================================
// 🇬🇧 English-Only UI Synchronizer
// 6 題型：單字選擇／克漏字／文意選填／篇章結構／閱讀測驗／翻譯
// ============================================================================
const ENGLISH_QTYPES = ['choice', 'cloze', 'wordbank', 'structure', 'reading', 'translation'];

function normalizeEnglishQType(t) {
  if (t === 'multiselect') return 'choice';
  if (t === 'fill') return 'translation';
  return ENGLISH_QTYPES.includes(t) ? t : 'choice';
}

function shortText(s, n = 110) {
  const str = String(s || '').replace(/\s+/g, ' ').trim();
  return str.length > n ? `${str.slice(0, n)}…` : str;
}

function setTextById(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

function isPredominantlyEnglishText(text) {
  const s = String(text || '').trim();
  if (!s || s.length < 6) return false;
  const cjkCount = (s.match(/[\u4e00-\u9fa5]/g) || []).length;
  const asciiWords = s.match(/\b[A-Za-z]{3,}\b/g) || [];
  return cjkCount <= 2 && asciiWords.length >= 2;
}

// English-only app: subject is always 英文
function autoDetectSubjectFromInputs() {
  const subjEl = document.getElementById('fieldSubject');
  if (subjEl) subjEl.value = '英文';
}

// Which option letters are used by each English question type
function getActiveOptionKeys(qType = state.questionType) {
  const qt = normalizeEnglishQType(qType);
  if (qt === 'translation') return [];
  if (qt === 'wordbank') return ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
  if (qt === 'structure') return ['A', 'B', 'C', 'D', 'E', 'F'];
  const keys = ['A', 'B', 'C', 'D'];
  if (document.getElementById('fieldOptE')?.value.trim()) keys.push('E');
  return keys;
}

const ENGLISH_OPTION_PLACEHOLDERS = {
  choice: '單字／片語',
  cloze: '單字／片語／轉折詞',
  wordbank: '文意選填單字',
  structure: '完整英文句子',
  reading: '選項敘述'
};

const ENGLISH_COACH_LABELS = {
  choice: {
    stg1: '【階段 1】🔍 句意與空格線索拆解（整句中譯、空格詞性、前後搭配詞）：',
    h1: '💡 提示 1｜上下文語境線索與空格詞性判斷',
    h2: '🔑 提示 2｜各選項中文意思、詞性與同義詞對照',
    h3: '🎯 提示 3｜正解核對：為什麼是它而不是你選的那個',
    stg3: '【階段 3】⚡ 核心單字「舉一反三」小自測：',
    twin: '在此輸入應填的英文單字...'
  },
  cloze: {
    stg1: '【階段 1】🔍 克漏字前後文拆解（空格所在句 × 前一句 × 後一句）：',
    h1: '💡 提示 1｜先判斷本空格考「字彙／片語／文法／轉折詞」哪一類',
    h2: '🔑 提示 2｜各選項中文意思與文法功能對照',
    h3: '🎯 提示 3｜放回原文驗證：前後文語意與文法都要通',
    stg3: '【階段 3】⚡ 同考點「舉一反三」小自測：',
    twin: '在此輸入你的答案...'
  },
  wordbank: {
    stg1: '【階段 1】🔍 文意選填空格拆解（先判斷空格詞性：n. / v. / adj. / adv.）：',
    h1: '💡 提示 1｜空格前後的詞性線索（冠詞、be 動詞、介系詞、助動詞）',
    h2: '🔑 提示 2｜十個選項的詞性與中文意思',
    h3: '🎯 提示 3｜同詞性候選字語意比較與正解核對',
    stg3: '【階段 3】⚡ 核心單字「舉一反三」小自測：',
    twin: '在此輸入應填的英文單字...'
  },
  structure: {
    stg1: '【階段 1】🔍 篇章結構空格拆解（空格前一句 × 空格後一句 × 段落主題）：',
    h1: '💡 提示 1｜先讀前後句：找指代詞（this / these / such / they）與轉折詞',
    h2: '🔑 提示 2｜每個選項句子的「銜接線索」',
    h3: '🎯 提示 3｜正解句放回原文的前後呼應檢查',
    stg3: '【階段 3】⚡ 篇章銜接「舉一反三」小自測：',
    twin: '在此輸入選項代號（例如 B）...'
  },
  reading: {
    stg1: '【階段 1】🔍 閱讀題目拆解（先判斷題型：主旨／細節／推論／字義／指代）：',
    h1: '💡 提示 1｜依題型決定要回原文哪裡找答案',
    h2: '🔑 提示 2｜逐一檢查選項：同義改寫 vs. 偷換概念／過度推論',
    h3: '🎯 提示 3｜正解的原文證據與誘答選項陷阱',
    stg3: '【階段 3】⚡ 閱讀技巧「舉一反三」小自測：',
    twin: '在此輸入你的答案...'
  },
  translation: {
    stg1: '【階段 1】🔍 中文句子拆解（主詞、動詞時態、核心句型與關鍵片語）：',
    h1: '💡 提示 1｜先搭英文句子骨架：主詞 + 動詞 + 時態 + 句型',
    h2: '🔑 提示 2｜你的寫法與標準答案逐字比對',
    h3: '🎯 提示 3｜完整標準譯句與扣分檢查清單',
    stg3: '【階段 3】⚡ 同句型「舉一反三」翻譯小自測：',
    twin: '在此輸入你的英文翻譯...'
  }
};

// Adapts the whole form to the selected English question type (no math / other subjects)
function syncSubjectSpecificUI() {
  const subjEl = document.getElementById('fieldSubject');
  if (subjEl) subjEl.value = '英文';
  const qt = normalizeEnglishQType(state.questionType);
  const keys = getActiveOptionKeys(qt);

  // 1. Option input fields (A~D / A~F / A~J)
  const groupOptE = document.getElementById('groupOptE');
  if (groupOptE) groupOptE.style.display = keys.includes('E') ? 'block' : 'none';
  const groupFJ = document.getElementById('groupWordBankFJ');
  if (groupFJ) groupFJ.style.display = keys.includes('F') ? 'grid' : 'none';
  ['F', 'G', 'H', 'I', 'J'].forEach((k) => {
    const inp = document.getElementById(`fieldOpt${k}`);
    const wrap = inp?.parentElement;
    if (wrap && wrap.id !== 'groupWordBankFJ') {
      wrap.style.display = keys.includes(k) ? '' : 'none';
    }
  });
  const ph = ENGLISH_OPTION_PLACEHOLDERS[qt] || '選項內容';
  ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach((k) => {
    setTextById(`labelOpt${k}`, `選項 (${k})`);
    const inp = document.getElementById(`fieldOpt${k}`);
    if (inp) inp.placeholder = `選項 ${k}：${ph}`;
  });

  // 2. STEP 2 answer pills
  ['wrongOptionPills', 'correctOptionPills'].forEach((cid) => {
    document.querySelectorAll(`#${cid} .option-pill`).forEach((btn) => {
      const val = btn.getAttribute('data-val');
      btn.style.display = keys.includes(val) ? 'inline-flex' : 'none';
      btn.textContent = val;
    });
  });

  // 3. Self-learning coach titles follow the question type
  const L = ENGLISH_COACH_LABELS[qt] || ENGLISH_COACH_LABELS.choice;
  setTextById('slCoachSubjectBadge', `英文｜${getQuestionTypeLabel(qt)}`);
  setTextById('slStage1Title', L.stg1);
  setTextById('labelSlHint1', L.h1);
  setTextById('labelSlHint2', L.h2);
  setTextById('labelSlHint3', L.h3);
  setTextById('slStage3Title', L.stg3);
  const twinInput = document.getElementById('inputSlTwinAnswer');
  if (twinInput) twinInput.placeholder = L.twin;
}

function getAllFormOptions() {
  const opts = {};
  ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach((k) => {
    opts[k] = document.getElementById(`fieldOpt${k}`)?.value.trim() || '';
  });
  return opts;
}

function setQuestionTypeUI(qType) {
  state.questionType = normalizeEnglishQType(qType || 'choice');
  const qt = state.questionType;
  document
    .querySelectorAll('#questionTypeContainer .diagnosis-tag')
    .forEach((btn) => {
      btn.classList.toggle('selected', btn.getAttribute('data-qtype') === qt);
    });

  const isChoiceLike = qt !== 'translation';
  const optsGrid = document.getElementById('optionsFourGrid');
  const wrongPills = document.getElementById('wrongOptionPills');
  const correctPills = document.getElementById('correctOptionPills');
  const wrongInput = document.getElementById('fieldWrongAnswer');
  const correctInput = document.getElementById('fieldCorrectAnswer');
  const labelWrong = document.getElementById('labelWrongAns');
  const labelCorrect = document.getElementById('labelCorrectAns');
  const diffBox = document.getElementById('liveTextDiffBox');

  if (optsGrid) optsGrid.style.display = isChoiceLike ? 'grid' : 'none';
  if (wrongPills) wrongPills.style.display = isChoiceLike ? 'flex' : 'none';
  if (correctPills) correctPills.style.display = isChoiceLike ? 'flex' : 'none';

  syncSubjectSpecificUI();

  const keys = getActiveOptionKeys(qt);
  const range = keys.length ? `${keys[0]} ~ ${keys[keys.length - 1]}` : '';
  const labelMap = {
    choice: ['❌ 我當初選錯的選項（單字選擇）', '✅ 試卷正確答案（單字選擇）'],
    cloze: ['❌ 我當初選錯的選項（克漏字）', '✅ 試卷正確答案（克漏字）'],
    wordbank: ['❌ 我當初選錯的代號（文意選填 A ~ J）', '✅ 試卷正確代號（文意選填 A ~ J）'],
    structure: ['❌ 我當初選錯的句子代號（篇章結構 A ~ F）', '✅ 試卷正確句子代號（篇章結構 A ~ F）'],
    reading: ['❌ 我當初選錯的選項（閱讀測驗）', '✅ 試卷正確答案（閱讀測驗）'],
    translation: ['❌ 我當初寫的英文翻譯', '✅ 標準英文翻譯解答']
  };
  const [lw, lc] = labelMap[qt] || labelMap.choice;
  if (labelWrong) labelWrong.textContent = lw;
  if (labelCorrect) labelCorrect.textContent = lc;

  if (qt === 'translation') {
    if (wrongInput) wrongInput.placeholder = '例如：The company spent three years develop the new phone.';
    if (correctInput) correctInput.placeholder = '例如：It took the company three years to develop the new phone.';
  } else {
    if (wrongInput) wrongInput.placeholder = `點選上方 ${range} 按鈕或直接輸入代號`;
    if (correctInput) correctInput.placeholder = `點選上方 ${range} 按鈕或直接輸入代號`;
    if (diffBox) diffBox.style.display = 'none';
  }

  // Re-sync option pill highlights
  if (wrongInput) setOptionPillAndInput('wrong', wrongInput.value, false);
  if (correctInput) setOptionPillAndInput('correct', correctInput.value, false);
}

// Convert any choice/multi-choice string like "2, 5", "25", "(2)(5)", "BE", "J" into canonical letters ['A'..'J']
function parseMultiChoiceTokens(rawStr) {
  const numToLetter = { '1': 'A', '2': 'B', '3': 'C', '4': 'D', '5': 'E' };
  const tokens = [];
  const str = String(rawStr || '').toUpperCase();
  for (const ch of str) {
    const mapped = numToLetter[ch] || (/[A-J]/.test(ch) ? ch : null);
    if (mapped && !tokens.includes(mapped)) {
      tokens.push(mapped);
    }
  }
  return tokens.sort();
}

function usesNumericOptionLabels() {
  // English exams always use letter options (A~J)
  return false;
}

function formatMultiChoiceTokensForDisplay(letterTokens) {
  const letterToNum = { A: '1', B: '2', C: '3', D: '4', E: '5' };
  const useNum = usesNumericOptionLabels();
  return letterTokens
    .map((k) => (useNum ? letterToNum[k] || k : k))
    .join(', ');
}

function handleOptionPillClick(which, clickedLetter) {
  const inputId = which === 'wrong' ? 'fieldWrongAnswer' : 'fieldCorrectAnswer';
  const inputEl = document.getElementById(inputId);
  if (!inputEl) return;

  if (state.questionType === 'multiselect') {
    const currentTokens = parseMultiChoiceTokens(inputEl.value);
    const idx = currentTokens.indexOf(clickedLetter);
    if (idx >= 0) {
      currentTokens.splice(idx, 1);
    } else {
      currentTokens.push(clickedLetter);
    }
    currentTokens.sort();
    const formatted = formatMultiChoiceTokensForDisplay(currentTokens);
    setOptionPillAndInput(which, formatted, true);
  } else {
    const useNum = usesNumericOptionLabels();
    const letterToNum = { A: '1', B: '2', C: '3', D: '4', E: '5' };
    const dispVal = useNum ? letterToNum[clickedLetter] || clickedLetter : clickedLetter;
    setOptionPillAndInput(which, dispVal, true);
  }
}

function setOptionPillAndInput(which, val, updateInput = true) {
  const containerId = which === 'wrong' ? 'wrongOptionPills' : 'correctOptionPills';
  const inputId = which === 'wrong' ? 'fieldWrongAnswer' : 'fieldCorrectAnswer';

  if (updateInput) {
    document.getElementById(inputId).value = val;
  }

  const rawStr = String(val || '');
  const isOptionTokenStr = /^[\s,、()（）A-Ja-j]*$/.test(rawStr);
  const activeLetters =
    state.questionType !== 'translation' && isOptionTokenStr
      ? parseMultiChoiceTokens(rawStr)
      : [];
  // Auto-switch: typed G~J → 文意選填 (A~J); typed E/F in a 4-option type → 篇章結構 (A~F)
  if (
    activeLetters.some((k) => ['G', 'H', 'I', 'J'].includes(k)) &&
    state.questionType !== 'wordbank'
  ) {
    setQuestionTypeUI('wordbank');
  } else if (
    activeLetters.some((k) => ['E', 'F'].includes(k)) &&
    ['choice', 'cloze', 'reading'].includes(state.questionType) &&
    !document.getElementById('fieldOptE')?.value.trim()
  ) {
    setQuestionTypeUI('structure');
  }

  document.querySelectorAll(`#${containerId} .option-pill`).forEach((b) => {
    const pillLetter = b.getAttribute('data-val');
    b.classList.toggle('selected', activeLetters.includes(pillLetter));
  });
}

// ============================================================================
// Word-by-Word & Token Diff Engine (For Translation, Fill-in-Blank & Science)
// ============================================================================
function computeWordDiffAnalysis(wrongStr, correctStr) {
  const tokenize = (s) =>
    String(s || '')
      .trim()
      .split(/(\s+|[,.;!?，。；！？、]+)/)
      .filter((t) => t && t.trim());

  const wTokens = tokenize(wrongStr);
  const cTokens = tokenize(correctStr);

  // LCS matrix for token alignment
  const m = wTokens.length;
  const n = cTokens.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));

  const norm = (t) => t.toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]/g, '');

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (norm(wTokens[i - 1]) === norm(cTokens[j - 1]) && norm(wTokens[i - 1]) !== '') {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  let i = m;
  let j = n;
  const ops = [];
  while (i > 0 || j > 0) {
    if (
      i > 0 &&
      j > 0 &&
      norm(wTokens[i - 1]) === norm(cTokens[j - 1]) &&
      norm(wTokens[i - 1]) !== ''
    ) {
      ops.unshift({ type: 'same', w: wTokens[i - 1], c: cTokens[j - 1] });
      i--;
      j--;
    } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
      ops.unshift({ type: 'add', text: cTokens[j - 1] });
      j--;
    } else {
      ops.unshift({ type: 'del', text: wTokens[i - 1] });
      i--;
    }
  }

  const wrongWords = ops.filter((o) => o.type === 'del').map((o) => o.text);
  const missingWords = ops.filter((o) => o.type === 'add').map((o) => o.text);

  const htmlParts = ops.map((o) => {
    if (o.type === 'same') {
      return `<span>${escapeHtml(o.c)}</span>`;
    } else if (o.type === 'del') {
      return `<span style="background:#fee2e2; color:#b91c1c; text-decoration:line-through; padding:1px 5px; border-radius:4px; font-weight:600;">${escapeHtml(
        o.text
      )}</span>`;
    } else {
      return `<span style="background:#d1fae5; color:#047857; border-bottom:2px solid #10b981; padding:1px 5px; border-radius:4px; font-weight:700;">${escapeHtml(
        o.text
      )}</span>`;
    }
  });

  return {
    diffHtml: htmlParts.join(' '),
    wrongWords,
    missingWords
  };
}

// ============================================================================
// 🇬🇧 English-Only Automatic Mistake Diagnosis Engine (6 題型)
//  🔤 單字選擇 / 🧩 克漏字 / 🔠 文意選填 / 📑 篇章結構 / 📖 閱讀測驗 / ✍️ 翻譯
// ============================================================================
const ENGLISH_TRANSITION_GUIDE = [
  ['on the other hand', '另一方面（對比另一面向）'],
  ['in contrast', '相對地（兩者對比）'],
  ['by contrast', '相對地（兩者對比）'],
  ['as a result', '結果、因此（前因後果）'],
  ['for example', '例如（後面是前句的具體例子）'],
  ['for instance', '例如（後面是前句的具體例子）'],
  ['in addition', '此外（同方向補充）'],
  ['in other words', '換句話說（重述前句）'],
  ['in fact', '事實上（強調或修正前句）'],
  ['in short', '簡言之（總結）'],
  ['in conclusion', '總之（總結）'],
  ['to sum up', '總之（總結）'],
  ['even though', '即使（讓步）'],
  ['what is more', '更重要的是（遞進補充）'],
  ['however', '然而（前後語意相反）'],
  ['nevertheless', '儘管如此（讓步轉折）'],
  ['nonetheless', '儘管如此（讓步轉折）'],
  ['therefore', '因此（前因後果）'],
  ['thus', '因此（前因後果）'],
  ['hence', '因此（前因後果）'],
  ['consequently', '因此（前因後果）'],
  ['moreover', '而且（遞進補充）'],
  ['furthermore', '而且（遞進補充）'],
  ['besides', '此外（同方向補充）'],
  ['instead', '反而、取而代之（否定前者後提出替代）'],
  ['otherwise', '否則（若不…就會…）'],
  ['indeed', '的確（強調）'],
  ['similarly', '同樣地（類比）'],
  ['likewise', '同樣地（類比）'],
  ['meanwhile', '同時（並行）'],
  ['finally', '最後（順序）'],
  ['although', '雖然（讓步）'],
  ['though', '雖然（讓步）'],
  ['whereas', '然而（對比）'],
  ['unless', '除非（條件）'],
  ['because', '因為（原因）'],
  ['yet', '然而（前後語意相反）'],
  ['but', '但是（前後語意相反）'],
  ['also', '也（同方向補充）']
];

function lookupTransitionAtStart(text) {
  const s = String(text || '')
    .toLowerCase()
    .replace(/^\s*\(?[a-j]\)\s*/, '')
    .replace(/^[\s"'“‘]+/, '')
    .trim();
  if (!s) return null;
  for (const [word, meaning] of ENGLISH_TRANSITION_GUIDE) {
    if (s === word || s.startsWith(`${word} `) || s.startsWith(`${word},`)) {
      return { word, meaning };
    }
  }
  return null;
}

const COHESION_REFERENCE_WORDS = [
  'this', 'these', 'those', 'such', 'they', 'them', 'their', 'it', 'its',
  'he', 'she', 'his', 'her', 'the same', 'both', 'another', 'the other'
];

function analyzeSentenceCohesion(sentence) {
  const clean = String(sentence || '').replace(/^\s*\(?[A-Ja-j]\)\s*/, '').trim();
  const lower = clean.toLowerCase();
  const transition = lookupTransitionAtStart(lower);
  const firstPart = lower.split(/\s+/).slice(0, 10).join(' ');
  const refs = COHESION_REFERENCE_WORDS.filter((w) =>
    new RegExp(`\\b${w}\\b`).test(firstPart)
  );
  return { clean, transition, refs };
}

// Finds the sentence before / after a numbered blank (e.g. __41__, (41), ___)
function extractBlankContext(stem, qNum) {
  const text = String(stem || '').replace(/\s+/g, ' ').trim();
  const empty = { before: '', after: '', prev: '', next: '' };
  if (!text) return empty;
  const n = String(qNum || '').trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  let idx = -1;
  let len = 0;
  if (n) {
    const patterns = [
      new RegExp(`_{2,}\\s*\\(?${n}\\)?\\s*_{2,}`),
      new RegExp(`\\(\\s*${n}\\s*\\)`),
      new RegExp(`\\[\\s*${n}\\s*\\]`),
      new RegExp(`(?:^|\\s)${n}(?=[\\s._])`)
    ];
    for (const re of patterns) {
      const m = re.exec(text);
      if (m) {
        idx = m.index;
        len = m[0].length;
        break;
      }
    }
  }
  if (idx < 0) {
    const m = /_{3,}/.exec(text);
    if (m) {
      idx = m.index;
      len = m[0].length;
    }
  }
  if (idx < 0) return empty;
  const beforeSents = text.slice(0, idx).trim().split(/(?<=[.!?])\s+/).filter(Boolean);
  const afterSents = text.slice(idx + len).trim().split(/(?<=[.!?])\s+/).filter(Boolean);
  return {
    before: beforeSents[beforeSents.length - 1] || '',
    after: afterSents[0] || '',
    prev: beforeSents[beforeSents.length - 2] || '',
    next: afterSents[1] || ''
  };
}

// Guess what part of speech the blank needs from the word right before it
function guessBlankPartOfSpeech(before, after) {
  const prev = (String(before || '').trim().split(/\s+/).pop() || '')
    .toLowerCase()
    .replace(/[^a-z']/g, '');
  const next = (String(after || '').trim().split(/\s+/)[0] || '')
    .toLowerCase()
    .replace(/[^a-z']/g, '');
  if (!prev) return null;
  if (['a', 'an', 'the', 'this', 'that', 'these', 'those', 'my', 'your', 'his', 'her', 'its', 'our', 'their'].includes(prev)) {
    return {
      pos: '名詞 (n.) 或 形容詞 (adj.)',
      reason: `空格前是限定詞「${prev}」，後面要接名詞${next ? `；若空格後還有名詞「${next}」，空格就要填形容詞修飾它` : ''}`
    };
  }
  if (['to', 'can', 'could', 'will', 'would', 'should', 'must', 'may', 'might', "don't", "doesn't", "didn't", 'not'].includes(prev)) {
    return { pos: '原形動詞 (V)', reason: `空格前是「${prev}」，後面接原形動詞` };
  }
  if (['is', 'am', 'are', 'was', 'were', 'be', 'been', 'being', 'become', 'became', 'seem', 'seems', 'seemed', 'look', 'looks', 'feel', 'feels', 'felt', 'remain', 'remained', 'get', 'got'].includes(prev)) {
    return { pos: '形容詞 (adj.) 或 分詞 (V-ing / p.p.)', reason: `空格前是 be／連綴動詞「${prev}」，後面常接形容詞或分詞（被動或情緒用 p.p.）` };
  }
  if (['very', 'so', 'too', 'quite', 'rather', 'extremely', 'more', 'most', 'less'].includes(prev)) {
    return { pos: '形容詞 (adj.) 或 副詞 (adv.)', reason: `空格前是程度副詞「${prev}」` };
  }
  if (['in', 'on', 'at', 'of', 'for', 'with', 'by', 'about', 'from', 'without', 'after', 'before'].includes(prev)) {
    return { pos: '名詞 (n.) 或 動名詞 (V-ing)', reason: `空格前是介系詞「${prev}」，後面接名詞或 V-ing` };
  }
  if (['have', 'has', 'had'].includes(prev)) {
    return { pos: '過去分詞 (p.p.) 或 名詞 (n.)', reason: `空格前是「${prev}」，若為完成式要接 p.p.` };
  }
  if (['he', 'she', 'it', 'they', 'we', 'i', 'you', 'people'].includes(prev)) {
    return { pos: '動詞 (V)', reason: `空格前是主詞「${prev}」，空格需要動詞（注意時態與單複數）` };
  }
  return null;
}

function simplifyPos(p) {
  const s = String(p || '');
  if (/adv|副詞/.test(s)) return 'adv';
  if (/adj|形容詞|分詞/.test(s)) return 'adj';
  if (/\bn\.|名詞/.test(s)) return 'n';
  if (/\bv\.|動詞/.test(s)) return 'v';
  return s;
}

function resolveAnswerKey(raw, opts) {
  const s = String(raw || '').trim();
  if (!s) return '';
  const up = s.toUpperCase().replace(/[()（）\s]/g, '');
  if (/^[A-J]$/.test(up)) return up;
  const lower = s.toLowerCase();
  const found = Object.keys(opts || {}).find(
    (k) => opts[k] && opts[k].trim().toLowerCase() === lower
  );
  return found || '';
}

function levenshteinDistance(a, b) {
  const s = String(a || '').toLowerCase();
  const t = String(b || '').toLowerCase();
  const dp = Array.from({ length: s.length + 1 }, (_, i) => [i]);
  for (let j = 1; j <= t.length; j++) dp[0][j] = j;
  for (let i = 1; i <= s.length; i++) {
    for (let j = 1; j <= t.length; j++) {
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + (s[i - 1] === t[j - 1] ? 0 : 1)
      );
    }
  }
  return dp[s.length][t.length];
}

// High-school translation sentence patterns (句型) with explanations and twin drills
const TRANSLATION_PATTERNS = [
  {
    re: /\bit\s+(?:takes|took|will take|has taken)\b[\s\S]*\bto\b/i,
    name: 'It takes/took + 人 + 時間 + to V',
    note: '「某人花多少時間做某事」：用 It takes/took + 人 + 時間 + to V；人當主詞才用 sb. spend + 時間 + V-ing。',
    twin: { q: '請翻譯：我花了兩個小時完成這份報告。（請用 It takes 句型）', ans: ['it took me two hours to finish the report', 'it took me two hours to complete the report'] }
  },
  {
    re: /\bspen[dt]s?\b[\s\S]*\b\w+ing\b/i,
    name: 'sb. spend + 時間/金錢 + (in) V-ing',
    note: '人當主詞：sb. spend(s)/spent + 時間/金錢 + (in) V-ing 或 on + N。',
    twin: { q: '請翻譯：她每天花一小時練習鋼琴。', ans: ['she spends an hour practicing the piano every day', 'she spends one hour practicing the piano every day', 'she spends an hour practicing piano every day'] }
  },
  {
    re: /\bnot only\b[\s\S]*\bbut\b/i,
    name: 'not only A but (also) B',
    note: '「不但…而且…」：A、B 詞性要對稱；Not only 放句首要倒裝（Not only did he…）。',
    twin: { q: '請翻譯：他不但聰明，而且很勤奮。', ans: ['he is not only smart but also hardworking', 'he is not only clever but also hardworking', 'he is not only smart but also diligent'] }
  },
  {
    re: /\bsuch\b[\s\S]*\bthat\b/i,
    name: 'such (a/an) + adj. + N + that…',
    note: '「如此…以致於…」：such 後面接名詞（such a + adj. + N），so 後面接形容詞／副詞。',
    twin: { q: '請翻譯：那是一部如此感人的電影，以致於我哭了。', ans: ['it was such a touching movie that i cried', 'it was such a moving movie that i cried', 'it was such a touching film that i cried'] }
  },
  {
    re: /\bso\s+\w+(?:\s+\w+)?\s+that\b/i,
    name: 'so + adj./adv. + that…',
    note: '「如此…以致於…」：so 後接形容詞／副詞，such 後接名詞。',
    twin: { q: '請翻譯：這部電影太感人了，以致於我哭了。', ans: ['the movie was so touching that i cried', 'the movie was so moving that i cried', 'the film was so touching that i cried'] }
  },
  {
    re: /\btoo\s+\w+\s+(?:for\s+\w+\s+)?to\b/i,
    name: 'too + adj. + (for sb.) + to V',
    note: '「太…而不能…」：句型本身已含否定，後面不可再加 not。',
    twin: { q: '請翻譯：這箱子太重了，我搬不動。', ans: ['the box is too heavy for me to carry', 'the box is too heavy for me to move', 'this box is too heavy for me to carry'] }
  },
  {
    re: /\bthe\s+(?:more|less|\w+er)\b[^,]*,\s*the\s+(?:more|less|\w+er)\b/i,
    name: 'The + 比較級…, the + 比較級…',
    note: '「越…就越…」：兩個子句都以 the + 比較級開頭，中間用逗號隔開。',
    twin: { q: '請翻譯：你越努力，就會越成功。', ans: ['the harder you work, the more successful you will be', 'the harder you work the more successful you will be'] }
  },
  {
    re: /\bnot until\b|\bit\s+(?:was|is)\s+not\s+until\b/i,
    name: 'It was not until… that… / Not until 倒裝',
    note: '「直到…才…」：It was not until + 時間 + that + S + V，或 Not until + 時間 + 助動詞 + S + V（倒裝）。',
    twin: { q: '請翻譯：直到昨天我才知道這個消息。', ans: ['i did not know the news until yesterday', "i didn't know the news until yesterday", 'it was not until yesterday that i knew the news', 'not until yesterday did i know the news'] }
  },
  {
    re: /\b(?:is|are|am|was|were|be|get|got|become)\s+used to\s+\w+ing\b/i,
    name: 'be used to + V-ing（習慣於）',
    note: 'be/get used to + V-ing / N＝習慣於；used to + 原形 V＝過去常常。',
    twin: { q: '請翻譯：我已經習慣早起了。', ans: ['i am used to getting up early', "i'm used to getting up early", 'i have gotten used to getting up early', 'i am used to waking up early'] }
  },
  {
    re: /\bused to\s+[a-z]+\b/i,
    name: 'used to + V（過去常常，現在不再）',
    note: 'used to + 原形 V＝過去常常（現在不再）；be used to + V-ing＝習慣於。',
    twin: { q: '請翻譯：我以前常常在這條河裡游泳。', ans: ['i used to swim in this river', 'i used to swim in the river'] }
  },
  {
    re: /\bit\s+(?:is|was)\s+\w+\s+(?:for\s+\w+\s+)?to\b/i,
    name: 'It is + adj. + (for sb.) + to V（虛主詞）',
    note: '虛主詞 It 代替後面真正的主詞 to V，避免主詞太長。',
    twin: { q: '請翻譯：對學生來說，學好英文是很重要的。', ans: ['it is important for students to learn english well'] }
  },
  {
    re: /\b(?:have|has|had)\s+(?:been\s+)?\w+(?:ed|en|wn|ne|ght)\b/i,
    name: '完成式 have/has/had + p.p.',
    note: '中文有「已經、一直、過、自從」時，常用完成式 have/has + p.p.，搭配 for + 一段時間 / since + 時間點。',
    twin: { q: '請翻譯：我已經住在台北十年了。', ans: ['i have lived in taipei for ten years', 'i have been living in taipei for ten years'] }
  }
];

function detectTranslationPatterns(sentence) {
  const s = String(sentence || '');
  if (!s.trim()) return [];
  return TRANSLATION_PATTERNS.filter((p) => p.re.test(s));
}

function runAutoMistakeDiagnosis(forceOverwrite = true) {
  const rawWrong = document.getElementById('fieldWrongAnswer').value.trim();
  const rawCorrect = document.getElementById('fieldCorrectAnswer').value.trim();
  const stem = document.getElementById('fieldQuestionText').value.trim();
  const qNum = document.getElementById('fieldQuestionNum')?.value.trim() || '';
  const opts = getAllFormOptions();
  const subjEl = document.getElementById('fieldSubject');
  if (subjEl) subjEl.value = '英文';

  updateSelfLearningCoach();

  if (!rawWrong || !rawCorrect) {
    return;
  }

  let qt = normalizeEnglishQType(state.questionType);
  const wKey = resolveAnswerKey(rawWrong, opts);
  const cKey = resolveAnswerKey(rawCorrect, opts);
  const hasAnyOptionText = Object.values(opts).some(Boolean);

  // Typed full sentences with no options → treat as translation
  if (qt !== 'translation' && (!wKey || !cKey) && !hasAnyOptionText) {
    qt = 'translation';
    setQuestionTypeUI('translation');
  }

  const diffBox = document.getElementById('liveTextDiffBox');
  const diffContent = document.getElementById('liveTextDiffContent');

  let result;
  if (qt === 'translation') {
    result = diagnoseEnglishTranslation(rawWrong, rawCorrect, stem, diffBox, diffContent);
  } else {
    if (diffBox) diffBox.style.display = 'none';
    const wrongText = wKey ? opts[wKey] || `選項 (${wKey})` : rawWrong;
    const correctText = cKey ? opts[cKey] || `選項 (${cKey})` : rawCorrect;
    const ctx = {
      qt,
      stem,
      qNum,
      opts,
      wKey,
      cKey,
      wrongText,
      correctText,
      wLabel: wKey && opts[wKey] ? `(${wKey}) ${opts[wKey]}` : wrongText,
      cLabel: cKey && opts[cKey] ? `(${cKey}) ${opts[cKey]}` : correctText
    };
    if (qt === 'structure') result = diagnoseEnglishStructure(ctx);
    else if (qt === 'reading') result = diagnoseEnglishReading(ctx);
    else if (qt === 'wordbank') result = diagnoseEnglishWordBank(ctx);
    else result = diagnoseEnglishChoiceOrCloze(ctx);
  }

  const { autoTag, whyWrong, correctConcept, keyTakeaway } = result;

  // Update Error Tag Selection UI if none manually clicked yet
  if (state.selectedErrorTags.size === 0) {
    state.selectedErrorTags.add(autoTag);
    document.querySelectorAll('#errorTagsContainer .diagnosis-tag').forEach((b) => {
      b.classList.toggle('selected', b.getAttribute('data-tag') === autoTag);
    });
  }

  if (forceOverwrite || !document.getElementById('fieldWhyWrong').value.trim()) {
    document.getElementById('fieldWhyWrong').value = whyWrong;
    document.getElementById('fieldCorrectConcept').value = correctConcept;
    document.getElementById('fieldKeyTakeaway').value = keyTakeaway;
  }

  document.getElementById('previewWhyWrong').textContent =
    document.getElementById('fieldWhyWrong').value;
  document.getElementById('previewCorrectConcept').textContent =
    document.getElementById('fieldCorrectConcept').value;
  document.getElementById('previewKeyTakeaway').textContent =
    document.getElementById('fieldKeyTakeaway').value;

  const badge = document.getElementById('autoDiagnosisStatusBadge');
  if (badge) {
    badge.textContent = `🤖 已根據【英文・${getQuestionTypeLabel(qt)}】自動分析錯因與正確寫法！（看完即可直接儲存）`;
  }
}

// ✍️ 翻譯題（中翻英）
function diagnoseEnglishTranslation(rawWrong, rawCorrect, stem, diffBox, diffContent) {
  const diffRes = computeWordDiffAnalysis(rawWrong, rawCorrect);
  if (diffBox && diffContent) {
    diffBox.style.display = 'block';
    diffContent.innerHTML = diffRes.diffHtml;
  }

  const notes = [];
  const patterns = detectTranslationPatterns(rawCorrect);
  patterns.forEach((p) => {
    if (!p.re.test(rawWrong)) {
      notes.push(`沒有寫出句型「${p.name}」：${p.note}`);
    }
  });
  if (/\b(take|takes|took)\b/i.test(rawCorrect) && /\b(spend|spent|cost)\b/i.test(rawWrong) && !patterns.some((p) => /It takes/.test(p.name))) {
    notes.push('混淆了「花費時間」的句型：It takes/took + 人 + 時間 + to V（人當主詞才用 spend + 時間 + V-ing）。');
  }
  if (
    /\b(took|was|were|had|did|went|spent|made|found|\w+ed)\b/i.test(rawCorrect) &&
    !/\b(took|was|were|had|did|went|spent|made|found|\w+ed)\b/i.test(rawWrong)
  ) {
    notes.push('時態錯誤：發生在過去的事件，動詞要用過去式。');
  }
  if (/\bto\s+[a-z]+\b/i.test(rawCorrect) && /\bto\s+[a-z]+ing\b/i.test(rawWrong)) {
    notes.push('不定詞 to 後面要接原形動詞（to + V），不是 V-ing。');
  }
  const missingLower = diffRes.missingWords.map((w) => w.toLowerCase());
  const wrongLower = diffRes.wrongWords.map((w) => w.toLowerCase());
  if (missingLower.some((w) => ['a', 'an', 'the'].includes(w))) {
    notes.push('漏了冠詞 a / an / the（可數單數名詞前一定要有冠詞或限定詞）。');
  }
  if (missingLower.some((m) => wrongLower.some((w) => m === `${w}s` || m === `${w}es`))) {
    notes.push('名詞複數或第三人稱單數動詞漏加 -s / -es。');
  }
  const spellingPairs = [];
  diffRes.wrongWords.forEach((w) => {
    const match = diffRes.missingWords.find(
      (m) =>
        m.length >= 4 &&
        w.length >= 4 &&
        m.toLowerCase() !== w.toLowerCase() &&
        levenshteinDistance(m, w) <= 2 &&
        m.toLowerCase().slice(0, 2) === w.toLowerCase().slice(0, 2)
    );
    if (match) spellingPairs.push(`${w} → ${match}`);
  });
  if (spellingPairs.length) {
    notes.push(`拼字錯誤：${spellingPairs.join('、')}。`);
  }
  if (/^[a-z]/.test(rawWrong)) notes.push('句首第一個字母要大寫。');
  if (!/[.!?]$/.test(rawWrong) && /[.!?]$/.test(rawCorrect)) notes.push('句尾要加標點（.）。');

  const wrongPartStr =
    diffRes.wrongWords.length > 0
      ? `你寫了「${diffRes.wrongWords.join(' ')}」`
      : '你的句子少了部分結構';
  const rightPartStr =
    diffRes.missingWords.length > 0
      ? `標準答案應為「${diffRes.missingWords.join(' ')}」`
      : `標準寫法為「${rawCorrect}」`;

  const whyWrong = `【逐字比對診斷】：${wrongPartStr}，但${rightPartStr}！\n👉 關鍵失分原因：\n${
    notes.length > 0
      ? notes.map((n, i) => `${i + 1}. ${n}`).join('\n')
      : '關鍵片語搭配、動詞時態或單複數詞形與標準句型有落差。'
  }`;

  const correctConcept = `【這題應該這樣寫】：\n${stem ? `📌 中文題目：${shortText(stem, 120)}\n` : ''}1. 先抓出整句的主詞與動詞，決定時態。\n2. 套用核心句型：${
    patterns.length ? patterns.map((p) => p.name).join('、') : '確認主詞動詞一致、介系詞與片語搭配'
  }。\n3. 完整正確解答：\n「${rawCorrect}」`;

  const keyTakeaway = patterns.length
    ? `必背句型：${patterns[0].name} — ${patterns[0].note}`
    : diffRes.missingWords.length > 0
    ? `必背寫法：❌ ${diffRes.wrongWords.join(' ') || '原錯句'} 👉 ✅ ${diffRes.missingWords.join(' ')}`
    : `熟記標準句型：「${rawCorrect}」`;

  return { autoTag: '翻譯句型與拼字', whyWrong, correctConcept, keyTakeaway };
}

// 🔤 單字選擇 & 🧩 克漏字
function diagnoseEnglishChoiceOrCloze(ctx) {
  const { qt, stem, qNum, opts, wKey, cKey, wrongText, correctText, wLabel, cLabel } = ctx;
  const isCloze = qt === 'cloze';
  const blank = isCloze ? extractBlankContext(stem, qNum) : extractBlankContext(stem, '');
  const hasBlank = Boolean(blank.before || blank.after);
  const localStem = isCloze && hasBlank ? `${blank.before} ___ ${blank.after}` : stem;
  const contextLine =
    isCloze && hasBlank
      ? `\n📌 空格所在句：「${shortText(blank.before, 90)} ___ ${shortText(blank.after, 90)}」`
      : '';
  const wLow = wrongText.toLowerCase().trim();
  const cLow = correctText.toLowerCase().trim();

  const relPronouns = ['that', 'who', 'whom', 'which', 'whose', 'where', 'when', 'what'];
  if (relPronouns.includes(wLow) || relPronouns.includes(cLow)) {
    const relExplainMap = {
      that: '「that」可代替人或物，但【不能用在逗號 ( , ) 後方的非限定子句】，也不能直接放在介系詞後面！',
      who: '「who」用於先行詞是【人】，且在子句中當【主詞】（後面緊接動詞）！',
      whom: '「whom」用於先行詞是【人】，且在子句中當【受詞】（後面接「主詞＋動詞」）或放在介系詞後！',
      which: '「which」用於先行詞是【事物／動物】或代表前面整個子句，不能指「人」！',
      whose: '「whose」是所有格關代（…的），後面必須緊接【名詞】！',
      where: '「where」＝ in/at which，先行詞是【地點】，後面接完整句子！',
      when: '「when」＝ at/on which，先行詞是【時間】，後面接完整句子！',
      what: '「what」＝ the thing(s) that，本身已含先行詞，前面不可再有名詞！'
    };
    return {
      autoTag: '文法與詞性判斷',
      whyWrong: `你選了 ${wLabel}，但正確答案是 ${cLabel}。${contextLine}\n👉 判斷錯誤點：${
        relExplainMap[wLow] || `誤判了「${wrongText}」的先行詞或子句結構。`
      }`,
      correctConcept: `【這題應該這樣寫】：\n1. 先找空格前的「先行詞」是人、物、地點還是時間。\n2. 檢查空格前有沒有逗號（有逗號不可用 that）。\n3. 看空格後缺主詞（接動詞）、缺受詞（接 S + V）還是接名詞：${
        relExplainMap[cLow] || `因此填入 ${cLabel}。`
      }`,
      keyTakeaway: '關代秒殺三步驟：①看先行詞(人/物) ②看有無逗號(有逗號禁選 that) ③看後面缺主詞(who/which)還是缺受詞(whom/which)！'
    };
  }

  const wTr = lookupTransitionAtStart(wLow);
  const cTr = lookupTransitionAtStart(cLow);
  if (cTr && (wTr || isCloze) && cLow.split(/\s+/).length <= 4) {
    return {
      autoTag: '篇章結構與轉折',
      whyWrong: `你選了 ${wLabel}${wTr ? `（${wTr.meaning}）` : ''}，但空格前後的邏輯關係是「${cTr.meaning}」。${contextLine}\n👉 判斷錯誤點：轉折詞題不能只看字面，要先判斷「前一句」與「後一句」是相反、因果、舉例還是補充關係。`,
      correctConcept: `【這題應該這樣解】：\n1. 讀空格前的內容：${blank.before ? `「${shortText(blank.before, 110)}」` : '（找出前一句的重點）'}\n2. 讀空格後的內容：${blank.after ? `「${shortText(blank.after, 110)}」` : '（找出後一句的重點）'}\n3. 兩者關係屬於「${cTr.meaning}」→ 正解 ${cLabel}。${
        wTr ? `\n4. 你選的「${wrongText}」表示「${wTr.meaning}」，邏輯方向不同。` : ''
      }`,
      keyTakeaway: '轉折詞秒殺：相反用 However / Instead、因果用 Therefore / As a result、舉例用 For example、補充用 In addition / Moreover！'
    };
  }

  if (isEnglishWordFamilyQuestion(opts)) {
    const wPos = guessEnglishPartOfSpeech(wrongText);
    const cPos = guessEnglishPartOfSpeech(correctText);
    const blankPos = guessBlankPartOfSpeech(blank.before, blank.after);
    return {
      autoTag: '文法與詞性判斷',
      whyWrong: `你選了 ${wLabel}（${wPos}），但這個空格在句子結構中需要「${cPos}」。${contextLine}${
        blankPos ? `\n👉 線索：${blankPos.reason}。` : ''
      }`,
      correctConcept: `【這題應該這樣寫】：\n這是「詞性判斷題（同字根變化）」，不用翻譯整句也能解！\n1. 看空格位置：主詞／受詞填名詞、修飾名詞或接在 be 動詞後填形容詞、修飾動詞或整句填副詞 (-ly)、助動詞／to 後面填原形動詞。\n2. ${cLabel} 是【${cPos}】，符合這個位置的需求。`,
      keyTakeaway: '四個選項長得很像 → 詞性題！先看空格前後決定詞性（-ness / -tion 名詞、-ly 副詞、-en / -ize 動詞、-ous / -ful / -ive 形容詞）。'
    };
  }

  ensureLiveEnglishTranslations(
    [wrongText, correctText, ...Object.values(opts)],
    localStem.length <= 400 ? localStem : ''
  );
  const wEntry = getEnglishWordEntry(wrongText);
  const cEntry = getEnglishWordEntry(correctText);
  const clue = analyzeEnglishContextClues(
    localStem,
    wKey || '?',
    wrongText,
    wEntry,
    cKey || '?',
    correctText,
    cEntry,
    opts
  );
  return {
    autoTag: '單字與片語搭配',
    whyWrong: clue.whyWrong + contextLine,
    correctConcept:
      clue.correctConcept +
      (isCloze ? '\n💡 克漏字提醒：答案常由「前一句或後一句」的同義詞、反義詞決定，不要只讀空格那一句！' : ''),
    keyTakeaway: clue.keyTakeaway
  };
}

// 🔠 文意選填 (A~J)
function diagnoseEnglishWordBank(ctx) {
  const { stem, qNum, opts, wKey, cKey, wrongText, correctText, wLabel } = ctx;
  const blank = extractBlankContext(stem, qNum);
  const localStem = blank.before || blank.after ? `${blank.before} ___ ${blank.after}` : stem;
  ensureLiveEnglishTranslations(
    [wrongText, correctText, ...Object.values(opts)],
    localStem.length <= 400 ? localStem : ''
  );
  const wEntry = getEnglishWordEntry(wrongText);
  const cEntry = getEnglishWordEntry(correctText);
  const wPos = wEntry?.pos || guessEnglishPartOfSpeech(wrongText);
  const cPos = cEntry?.pos || guessEnglishPartOfSpeech(correctText);
  const blankPos = guessBlankPartOfSpeech(blank.before, blank.after);
  const clue = analyzeEnglishContextClues(
    localStem,
    wKey || '?',
    wrongText,
    wEntry,
    cKey || '?',
    correctText,
    cEntry,
    opts
  );
  const posDiffer = simplifyPos(wPos) !== simplifyPos(cPos);
  const head = posDiffer
    ? `【詞性就不對】你選的 ${wLabel} 是「${wPos}」，但空格需要「${cPos}」${
        blankPos ? `（線索：${blankPos.reason}）` : ''
      }，詞性不符可以直接刪掉！\n`
    : `【同詞性但語意不符】你選的 ${wLabel} 和正解一樣是「${cPos}」，詞性判斷對了，但放進句子語意不通。\n`;
  return {
    autoTag: posDiffer ? '文法與詞性判斷' : '單字與片語搭配',
    whyWrong: head + clue.whyWrong,
    correctConcept: `【文意選填這樣解】：\n① 先判斷空格詞性：${
      blankPos ? `${blankPos.pos}（${blankPos.reason}）` : `本題空格需要「${cPos}」`
    }。\n② 從 (A)~(J) 只留下同詞性的 2~3 個候選字，再比較語意。\n${clue.correctConcept}`,
    keyTakeaway: `文意選填兩步驟：①先把十個選項分類成 n. / v. / adj. / adv. ②空格只從同詞性的候選字挑！${clue.keyTakeaway}`
  };
}

// 📑 篇章結構 (A~E/F)
function diagnoseEnglishStructure(ctx) {
  const { stem, qNum, cKey, wrongText, correctText, wLabel, cLabel } = ctx;
  const blank = extractBlankContext(stem, qNum);
  const wInfo = analyzeSentenceCohesion(wrongText);
  const cInfo = analyzeSentenceCohesion(correctText);
  const afterInfo = analyzeSentenceCohesion(blank.after);
  ensureLiveEnglishTranslations([], correctText);
  const cTrans = getCachedSentenceTranslation(correctText);

  const whyParts = [`你選了 ${wLabel}。`];
  if (wInfo.transition) {
    whyParts.push(
      `👉 這句以「${wInfo.transition.word}」開頭，表示「${wInfo.transition.meaning}」，但空格前一句${
        blank.before ? `「${shortText(blank.before, 90)}」` : ''
      }和這個邏輯關係接不起來。`
    );
  }
  if (wInfo.refs.length) {
    whyParts.push(`👉 句中的指代詞「${wInfo.refs.join('、')}」在空格前文找不到可以對應的名詞或概念。`);
  }
  if (!wInfo.transition && !wInfo.refs.length) {
    whyParts.push('👉 這句的話題和空格前後句沒有銜接（話題跳接），放進去段落邏輯會斷掉。');
  }
  if (afterInfo.transition || afterInfo.refs.length) {
    whyParts.push(
      `👉 而且空格「後一句」${afterInfo.transition ? `以「${afterInfo.transition.word}」開頭` : ''}${
        afterInfo.refs.length ? `有指代詞「${afterInfo.refs.join('、')}」` : ''
      }，它需要回指／承接空格句的內容，你選的句子接不上。`
    );
  }
  whyParts.push('⚠️ 常見錯因：只看「前一句」或只找「重複單字」就選，沒有同時檢查「後一句」是否也接得上。');

  let cohesionExplain = '';
  if (cInfo.transition) {
    cohesionExplain += `\n• 以「${cInfo.transition.word}」開頭，表示「${cInfo.transition.meaning}」，正好承接前一句。`;
  }
  if (cInfo.refs.length) {
    cohesionExplain += `\n• 句中「${cInfo.refs.join('、')}」回指前一句提到的人事物。`;
  }
  if (!cohesionExplain) {
    cohesionExplain = '\n• 與前後句討論同一主題，形成「主題句 → 支持細節」的銜接。';
  }

  return {
    autoTag: '篇章結構與轉折',
    whyWrong: whyParts.join('\n'),
    correctConcept: `【篇章結構這樣解】：\n1. 空格前一句：${
      blank.before || '（請在題目欄貼上含空格編號的整段文章並輸入題號，系統會自動抓出前後句）'
    }\n2. 空格後一句：${blank.after || '—'}\n3. 正解 ${cLabel}${cTrans ? `\n   中譯：${cTrans}` : ''}${cohesionExplain}\n4. 驗證：前一句 → ${cKey ? `(${cKey})` : '正解'} → 後一句，讀起來是一條連貫的邏輯鏈。`,
    keyTakeaway: '篇章結構三看：①看選項開頭的轉折詞（However / For example / Therefore）②看指代詞（this / these / such / they）回指誰 ③放回去「前後句都要通」！'
  };
}

function detectReadingQuestionType(stem) {
  const s = String(stem || '').toLowerCase();
  if (/main idea|mainly about|best title|primary purpose|purpose of (?:the|this) (?:passage|article)|主旨|標題/.test(s)) {
    return { type: 'main', label: '主旨／標題題' };
  }
  if (/refers? to|指的是|指代/.test(s)) return { type: 'reference', label: '指代題' };
  if (/closest in meaning|the word .* means|the phrase .* means|meaning of|字義|最接近/.test(s)) {
    return { type: 'vocab', label: '字義題（Vocabulary in Context）' };
  }
  if (/infer|imply|implied|suggest|conclude|most likely|probably|推論|推測/.test(s)) {
    return { type: 'inference', label: '推論題' };
  }
  if (/attitude|tone|feel about|態度|語氣/.test(s)) return { type: 'attitude', label: '作者態度／語氣題' };
  if (/not true|not mentioned|except|false|錯誤|未提及/.test(s)) return { type: 'notTrue', label: '是非（NOT / EXCEPT）題' };
  return { type: 'detail', label: '細節題' };
}

function extractReadingQuestion(stem) {
  const s = String(stem || '').replace(/\s+/g, ' ').trim();
  const m = s.match(/[^.?!]*\?/g);
  if (m && m.length) return m[m.length - 1].trim();
  return shortText(s.slice(-200), 200);
}

// 📖 閱讀測驗
function diagnoseEnglishReading(ctx) {
  const { stem, wLabel, cLabel } = ctx;
  const rq = detectReadingQuestionType(stem);
  const map = {
    main: {
      why: `【主旨題診斷】你選了 ${wLabel}，這是最常見的「以偏概全（只講到某一段細節）」或「範圍過大」的誘答選項。`,
      how: `1. 主旨題看「第一段（尤其最後一句）」與「各段第一句（Topic Sentence）」。\n2. 正解 ${cLabel} 能涵蓋整篇文章的核心，而不是單一細節。`,
      take: '主旨題避坑：只出現在某一段的細節選項直接刪！緊扣首段與末段總結句。'
    },
    inference: {
      why: `【推論題診斷】你選了 ${wLabel}，掉入了「過度推論（文章沒提到）」或「照抄原文單字但扭曲因果」的陷阱。`,
      how: `1. 推論題的正解一定是原文某句話的「合理延伸」或「同義改寫」。\n2. 回原文找支持句：只有 ${cLabel} 有明確的上下文證據。`,
      take: '推論題鐵則：沒寫到的不能腦補！有 always / never / all / only 等絕對字眼的選項要小心。'
    },
    vocab: {
      why: `【字義題診斷】你選了 ${wLabel}，可能是用這個字「最常見的意思」作答，沒有回到原文看上下文。`,
      how: `1. 回到該字所在句，把每個選項代進去讀讀看。\n2. 找前後文的同義、反義或舉例線索，只有 ${cLabel} 放進去語意通順。`,
      take: '字義題：一定要回原文代入！同一個字在不同語境意思不同。'
    },
    reference: {
      why: `【指代題診斷】你選了 ${wLabel}，可能找錯了被指代的對象（離太遠、單複數不一致）。`,
      how: `1. 從指代詞往「前面」找最近的名詞或整句概念。\n2. 檢查單複數（it / they）與語意是否一致 → ${cLabel}。`,
      take: '指代題：往前找、找最近、單複數要一致！'
    },
    attitude: {
      why: `【態度題診斷】你選了 ${wLabel}，可能誤判了作者用字的正負面語氣。`,
      how: `1. 圈出作者使用的形容詞／副詞（unfortunately, surprisingly, admirable…）。\n2. 判斷是正面、負面還是中立 → ${cLabel}。`,
      take: '態度題：看作者的形容詞與副詞，不要用自己的立場判斷！'
    },
    notTrue: {
      why: `【是非題診斷】你選了 ${wLabel}，這個敘述其實在文中是「正確／有提到」的，題目要選的是「不正確／沒提到」的那一個。`,
      how: `1. 題目有 NOT / EXCEPT 先圈起來。\n2. 每個選項回原文標記 ⭕ / ❌ / 沒提到，唯一不符的是 ${cLabel}。`,
      take: '看到 NOT / EXCEPT 先圈起來，四個選項逐一標記 ⭕ / ❌！'
    },
    detail: {
      why: `【細節題診斷】你選了 ${wLabel}，這個選項雖然有文章裡眼熟的單字，但把「主詞／時間／因果關係」張冠李戴了。`,
      how: `1. 從題目抓定位關鍵字（人名、地名、數字、專有名詞），回原文找到對應句子。\n2. 逐字比對：正解 ${cLabel} 是原文的「同義改寫（Paraphrase）」。`,
      take: '細節題秒殺：正解愛「換句話說」，照抄原文卻改掉因果的選項是陷阱！'
    }
  };
  const d = map[rq.type] || map.detail;
  return {
    autoTag: '閱讀定位與推論',
    whyWrong: `📖 題型判斷：${rq.label}\n${d.why}`,
    correctConcept: `【這題應該這樣找答案】：\n${d.how}`,
    keyTakeaway: d.take
  };
}

function getQuestionTypeLabel(qType) {
  const map = {
    choice: '單字選擇／字彙',
    cloze: '克漏字（綜合測驗）',
    wordbank: '文意選填 (A~J)',
    structure: '篇章結構 (A~E/F)',
    reading: '閱讀測驗',
    translation: '翻譯題（中翻英）'
  };
  return map[normalizeEnglishQType(qType)] || '單字選擇／字彙';
}

// ============================================================================
// 🧭 Autonomous Self-Directed Learning Coach (自主學習三階段引導精靈)
// ============================================================================
let currentTwinQuizSolution = {
  expectedAnswers: [],
  explanation: ''
};

function initSelfLearningCoach() {
  const bindHintToggle = (btnId, boxId, iconId) => {
    const btn = document.getElementById(btnId);
    const box = document.getElementById(boxId);
    const icon = document.getElementById(iconId);
    if (!btn || !box) return;
    btn.addEventListener('click', () => {
      const isHidden = box.style.display === 'none';
      box.style.display = isHidden ? 'block' : 'none';
      if (icon) icon.textContent = isHidden ? '收合提示 ▲' : '點擊展開 ▼';
    });
  };

  bindHintToggle('btnSlHint1', 'boxSlHint1', 'iconSlHint1');
  bindHintToggle('btnSlHint2', 'boxSlHint2', 'iconSlHint2');
  bindHintToggle('btnSlHint3', 'boxSlHint3', 'iconSlHint3');

  document
    .getElementById('btnGenerateSelfLearningGuide')
    ?.addEventListener('click', () => {
      updateSelfLearningCoach(true);
      showToast('🧭 已重新拆解題目條件並更新三階段自主學習引導！');
    });

  document
    .getElementById('btnCheckSlTwinAnswer')
    ?.addEventListener('click', () => {
      const inputEl = document.getElementById('inputSlTwinAnswer');
      const fbBox = document.getElementById('slTwinFeedbackBox');
      if (!inputEl || !fbBox) return;
      const userVal = inputEl.value.trim();
      if (!userVal) {
        alert('請先在輸入框寫下你算出的變式題答案！');
        return;
      }
      const normUser = userVal.replace(/\s+/g, '').toLowerCase();
      const matched = (currentTwinQuizSolution.expectedAnswers || []).some(
        (ans) => normUser === ans.replace(/\s+/g, '').toLowerCase() || normUser.includes(ans.replace(/\s+/g, '').toLowerCase())
      );

      fbBox.style.display = 'block';
      if (matched) {
        fbBox.style.background = '#ecfdf5';
        fbBox.style.border = '1px solid #6ee7b7';
        fbBox.style.color = '#065f46';
        fbBox.innerHTML = `🎉 <strong>太棒了！完全答對！</strong> 你已經真正搞懂這個核心觀念，不再只是死背原題答案！<br/>${escapeHtml(
          currentTwinQuizSolution.explanation
        )}`;
      } else {
        fbBox.style.background = '#fffbeb';
        fbBox.style.border = '1px solid #fde68a';
        fbBox.style.color = '#92400e';
        fbBox.innerHTML = `💡 <strong>還差一點點喔！</strong> 你的作答是「${escapeHtml(
          userVal
        )}」。請對照變式題詳解再試一次：<br/>${escapeHtml(
          currentTwinQuizSolution.explanation
        )}`;
      }
    });

  document
    .getElementById('btnRevealSlTwinSolution')
    ?.addEventListener('click', () => {
      const fbBox = document.getElementById('slTwinFeedbackBox');
      if (!fbBox) return;
      fbBox.style.display = 'block';
      fbBox.style.background = '#eef2ff';
      fbBox.style.border = '1px solid #c7d2fe';
      fbBox.style.color = '#312e81';
      fbBox.innerHTML = `📖 <strong>變式題標準詳解：</strong><br/>${escapeHtml(
        currentTwinQuizSolution.explanation
      )}`;
    });

  updateSelfLearningCoach();
}

function updateSelfLearningCoach(autoExpandFirstHint = false) {
  const stem = document.getElementById('fieldQuestionText')?.value.trim() || '';
  const wrongAns = document.getElementById('fieldWrongAnswer')?.value.trim() || '';
  const correctAns = document.getElementById('fieldCorrectAnswer')?.value.trim() || '';
  const qNum = document.getElementById('fieldQuestionNum')?.value.trim() || '';
  const qt = normalizeEnglishQType(state.questionType);
  const opts = getAllFormOptions();

  syncSubjectSpecificUI();

  const deconstructEl = document.getElementById('slDeconstructContent');
  const h1El = document.getElementById('boxSlHint1');
  const h2El = document.getElementById('boxSlHint2');
  const h3El = document.getElementById('boxSlHint3');
  const twinTextEl = document.getElementById('slTwinQuestionText');
  const twinFbEl = document.getElementById('slTwinFeedbackBox');
  if (!deconstructEl || !h1El || !h2El || !h3El || !twinTextEl) return;
  if (twinFbEl) twinFbEl.style.display = 'none';

  const wKey = qt === 'translation' ? '' : resolveAnswerKey(wrongAns, opts);
  const cKey = qt === 'translation' ? '' : resolveAnswerKey(correctAns, opts);
  const wText = wKey ? opts[wKey] || '' : wrongAns;
  const cText = cKey ? opts[cKey] || '' : correctAns;

  let guide;
  if (qt === 'translation') {
    guide = buildTranslationCoach(stem, wrongAns, correctAns);
  } else if (qt === 'structure') {
    guide = buildStructureCoach(stem, qNum, opts, wKey, wText, cKey, cText);
  } else if (qt === 'reading') {
    guide = buildReadingCoach(stem, opts, wKey, wText, cKey, cText);
  } else {
    guide = buildVocabClozeCoach(qt, stem, qNum, opts, wKey, wText, cKey, cText);
  }

  deconstructEl.innerHTML = guide.deconstructHtml;
  h1El.textContent = guide.hint1;
  h2El.textContent = guide.hint2;
  h3El.textContent = guide.hint3;
  twinTextEl.textContent = guide.twinQ;
  currentTwinQuizSolution = guide.twinSolution;

  if (autoExpandFirstHint) {
    h1El.style.display = 'block';
    const icon1 = document.getElementById('iconSlHint1');
    if (icon1) icon1.textContent = '收合提示 ▲';
  }
}

const COACH_TRANSLATION_STYLE =
  'margin-top:4px; color:#1e40af; background:#eff6ff; padding:6px 10px; border-radius:6px; font-weight:600;';

// 🔤 單字選擇 / 🧩 克漏字 / 🔠 文意選填
function buildVocabClozeCoach(qt, stem, qNum, opts, wKey, wText, cKey, cText) {
  const isCloze = qt === 'cloze';
  const isWB = qt === 'wordbank';
  const blank = extractBlankContext(stem, isCloze || isWB ? qNum : '');
  const hasBlank = Boolean(blank.before || blank.after);
  const localStem = (isCloze || isWB) && hasBlank ? `${blank.before} ___ ${blank.after}` : stem;
  const transStem = localStem.length <= 400 ? localStem : '';
  ensureLiveEnglishTranslations([wText, cText, ...Object.values(opts)].filter(Boolean), transStem);
  const stemTrans = getCachedSentenceTranslation(transStem);
  const clueSummary = extractSentenceClueWords(localStem);
  const blankPos = guessBlankPartOfSpeech(blank.before, blank.after);
  const wEntry = getEnglishWordEntry(wText);
  const cEntry = getEnglishWordEntry(cText);

  const keys = getActiveOptionKeys(qt).filter((k) => opts[k]);
  const optList = keys.map((k) => opts[k]);
  let category = 'vocab';
  if (optList.length && optList.every((o) => lookupTransitionAtStart(o) && o.trim().split(/\s+/).length <= 4)) {
    category = 'transition';
  } else if (optList.some((o) => ['who', 'whom', 'which', 'that', 'whose', 'where', 'when', 'what'].includes(o.toLowerCase().trim()))) {
    category = 'relative';
  } else if (isEnglishWordFamilyQuestion(opts)) {
    category = 'wordform';
  }
  const categoryLabel = {
    transition: '轉折詞／連接詞題（看前後句的邏輯關係）',
    relative: '關係詞文法題（看先行詞與子句結構）',
    wordform: '詞性變化題（看空格在句中的位置）',
    vocab: isWB ? '文意選填字彙題（先判斷詞性，再比語意）' : '字彙／片語語意題（看上下文線索）'
  }[category];

  const deconstructHtml = `
    <div>📌 <strong>${(isCloze || isWB) && hasBlank ? '空格所在句' : '題幹'}：</strong>${escapeHtml(
      shortText(localStem, 220) || '請先在左側框選題目，或在「題目敘述」貼上英文題目'
    )}</div>
    ${stemTrans ? `<div style="${COACH_TRANSLATION_STYLE}">💬 <strong>中文翻譯：</strong>${escapeHtml(stemTrans)}</div>` : ''}
    <div>🧩 <strong>本題考點：</strong>${escapeHtml(categoryLabel)}</div>
    ${blankPos ? `<div>🔤 <strong>空格詞性：</strong>${escapeHtml(blankPos.pos)}（${escapeHtml(blankPos.reason)}）</div>` : ''}
    ${clueSummary ? `<div>🔑 <strong>上下文線索字：</strong><span style="color:#b45309; font-weight:700;">${escapeHtml(clueSummary)}</span></div>` : ''}
  `;

  let hint1;
  if (category === 'transition') {
    hint1 = `先別看選項！讀空格「前面」和「後面」，用一句中文說出兩邊的關係：\n• 前：${blank.before || '…'}\n• 後：${blank.after || '…'}\n是「相反」、「因果」、「舉例」、「補充」還是「讓步」？`;
  } else if (category === 'relative') {
    hint1 = '關係詞三步驟：\n① 先行詞是人、物、地點還是時間？\n② 空格前有沒有逗號？（有逗號不能用 that）\n③ 空格後缺主詞（接動詞）→ who / which；缺受詞（接 S + V）→ whom / which；接名詞 → whose。';
  } else if (category === 'wordform') {
    hint1 = `四個選項同字根，只差詞性！看空格位置：${
      blankPos ? blankPos.reason : '冠詞／所有格後填名詞、be 動詞後填形容詞、修飾動詞填副詞、助動詞／to 後填原形動詞'
    }。`;
  } else {
    hint1 = `${blankPos ? `空格需要：${blankPos.pos}（${blankPos.reason}）。\n` : ''}找出題幹中「解釋空格意思」的線索：${
      clueSummary || '逗號後的補充說明、and / but 前後的同義或反義詞、因果連接詞'
    }。${stemTrans ? `\n整句中譯：${stemTrans}` : ''}${
      isCloze ? '\n克漏字提醒：答案線索常在「前一句或後一句」，不要只讀空格那一句。' : ''
    }`;
  }

  let hint2;
  if (!optList.length) {
    hint2 = '請先在 STEP 1 填入選項（或用左側框選自動帶入），這裡會列出每個選項的中文、詞性與同義詞。';
  } else if (category === 'transition') {
    hint2 = `【各轉折詞的意思】：\n${keys
      .map((k) => {
        const t = lookupTransitionAtStart(opts[k]);
        return `• (${k}) ${opts[k]} → ${t ? t.meaning : ''}`;
      })
      .join('\n')}`;
  } else {
    const lines = keys.map((k) => {
      const word = opts[k];
      const entry = getEnglishWordEntry(word);
      const pos = entry?.pos || guessEnglishPartOfSpeech(word);
      const zh = entry?.zh || '（查詢字義中…）';
      const syn = entry?.syn ? `（同義：${entry.syn}）` : '';
      return `• (${k}) ${word}：[${pos}] ${zh}${syn}`;
    });
    hint2 = `【各選項中文意思與詞性】${isWB ? '（先把同詞性的字分在一起）' : ''}：\n${lines.join('\n')}`;
  }

  const hint3 =
    cKey || cText
      ? `【正解核對】正確答案：${cKey ? `(${cKey}) ` : ''}${cText}${
          cEntry?.zh ? `（${cEntry.pos || ''} ${cEntry.zh}）` : ''
        }\n你的選擇：${wKey ? `(${wKey}) ` : ''}${wText || '尚未輸入'}${
          wEntry?.zh ? `（${wEntry.zh}）` : ''
        }\n👉 把正解放回空格再讀一次${isCloze ? '（連同前後句）' : '整句'}，確認語意與文法都通順；STEP 3 會說明你錯在哪裡。`
      : '請在 STEP 2 點選「我的誤選」與「正確答案」，這裡會出現正解核對。';

  let twinQ;
  let twinSolution;
  if (category === 'transition') {
    twinQ = '【舉一反三】填入最適合的轉折詞：\nTom studied very hard for the test. ______, he failed it because he was sick that day.\n（提示：前後語意相反）';
    twinSolution = {
      expectedAnswers: ['however', 'nevertheless', 'nonetheless'],
      explanation: '前一句「很努力讀書」，後一句「卻考不及格」，語意相反 → However / Nevertheless。'
    };
  } else if (category === 'relative') {
    twinQ = '【舉一反三】填入關係代名詞：The young student talked to the people, ______ were cleaning the beach.（空格前有逗號，先行詞是 people，後面緊接動詞 were）';
    twinSolution = {
      expectedAnswers: ['who'],
      explanation: '先行詞 people 是人，後面緊接動詞（缺主詞），且前有逗號不能用 that → who。'
    };
  } else if (category === 'wordform') {
    twinQ = '【舉一反三】填入正確詞性：The teacher spoke ______ (patient) to the crying child.';
    twinSolution = {
      expectedAnswers: ['patiently'],
      explanation: '空格修飾動詞 spoke，要用副詞 patiently。'
    };
  } else {
    const target = String(cText || '').trim();
    if (target && /^[a-zA-Z-]+$/.test(target)) {
      twinQ = `【舉一反三】拼出本題正解單字：\n中文意思「${cEntry?.zh || '（請參考提示 2）'}」${
        cEntry?.pos ? `（${cEntry.pos}）` : ''
      }${cEntry?.syn ? `，同義詞 ${cEntry.syn}` : ''}，字首「${target[0]}」、共 ${target.length} 個字母。`;
      twinSolution = {
        expectedAnswers: [target.toLowerCase()],
        explanation: `正確單字是「${target}」${cEntry?.zh ? `（${cEntry.zh}）` : ''}，記得連同題目中的線索字一起記！`
      };
    } else {
      twinQ = '【舉一反三】Eric has been a ______ (忠實的，同義於 loyal) employee for over ten years. 請輸入空格應填的單字：';
      twinSolution = {
        expectedAnswers: ['faithful', 'loyal'],
        explanation: 'for over ten years（長期）＋ employee → faithful（忠實的）。'
      };
    }
  }

  return { deconstructHtml, hint1, hint2, hint3, twinQ, twinSolution };
}

// 📑 篇章結構
function buildStructureCoach(stem, qNum, opts, wKey, wText, cKey, cText) {
  const blank = extractBlankContext(stem, qNum);
  const afterCoh = analyzeSentenceCohesion(blank.after);
  const keys = ['A', 'B', 'C', 'D', 'E', 'F'].filter((k) => opts[k]);
  const cCoh = analyzeSentenceCohesion(cText);
  if (cText) ensureLiveEnglishTranslations([], cText);
  const cTrans = getCachedSentenceTranslation(cText);

  const deconstructHtml = `
    <div>⬆️ <strong>空格前一句：</strong>${escapeHtml(
      blank.before || '（請在題目欄貼上含空格編號的整段文章，並輸入題號）'
    )}</div>
    <div>⬇️ <strong>空格後一句：</strong>${escapeHtml(blank.after || '—')}</div>
    <div>🧭 <strong>核心考點：</strong>篇章結構｜句與句的銜接（轉折詞、指代詞、主題句 → 支持句）</div>
  `;

  let hint1 = `【先讀前後句，再看選項】\n1. 前一句在講什麼主題？${
    blank.before ? `\n   →「${shortText(blank.before, 140)}」` : ''
  }\n2. 後一句開頭有沒有轉折詞或指代詞（this / these / such / they / it）？${
    blank.after ? `\n   →「${shortText(blank.after, 140)}」` : ''
  }`;
  if (afterCoh.transition) {
    hint1 += `\n   ⚡ 後一句以「${afterCoh.transition.word}」開頭（${afterCoh.transition.meaning}），空格句要能被它承接！`;
  }
  if (afterCoh.refs.length) {
    hint1 += `\n   ⚡ 後一句有指代詞「${afterCoh.refs.join('、')}」，它指的東西應該出現在空格句裡！`;
  }

  const hint2 = keys.length
    ? `【每個選項的「銜接線索」】：\n${keys
        .map((k) => {
          const o = analyzeSentenceCohesion(opts[k]);
          const clues = [];
          if (o.transition) clues.push(`開頭「${o.transition.word}」＝${o.transition.meaning}`);
          if (o.refs.length) clues.push(`指代詞「${o.refs.join('、')}」需要前文有對應的名詞`);
          if (!clues.length) clues.push('沒有明顯轉折／指代詞，可能是新主題句或段落開頭句');
          return `• (${k}) ${shortText(o.clean, 80)}\n   → ${clues.join('；')}`;
        })
        .join('\n')}`
    : '請在 STEP 1 填入 (A)~(E)/(F) 的選項句子，這裡會分析每句的開頭轉折詞與指代詞。';

  let hint3;
  if (cKey || cText) {
    hint3 = `【正解核對】正確答案 ${cKey ? `(${cKey}) ` : ''}${cText}${cTrans ? `\n中譯：${cTrans}` : ''}\n放回原文檢查：前一句 → 正解 → 後一句，是否是一條連貫的邏輯鏈？`;
    if (cCoh.transition) hint3 += `\n• 以「${cCoh.transition.word}」開頭＝${cCoh.transition.meaning}，正好承接前一句。`;
    if (cCoh.refs.length) hint3 += `\n• 句中「${cCoh.refs.join('、')}」回指前一句提到的人事物。`;
    if (wText) hint3 += `\n你選的 ${wKey ? `(${wKey}) ` : ''}${shortText(wText, 80)}：請對照 STEP 3 看看它哪裡接不上。`;
  } else {
    hint3 = 'STEP 2 選好答案後，這裡會顯示正解放回原文的驗證方式。';
  }

  return {
    deconstructHtml,
    hint1,
    hint2,
    hint3,
    twinQ: '【舉一反三】選出最適合填入空格的句子：\nMany people think that money brings happiness. ______ Studies show that once basic needs are met, more money adds little joy.\n(A) For example, he bought a new car last week.\n(B) However, this is not always true.\n(C) In addition, money is made of paper.\n請輸入代號：',
    twinSolution: {
      expectedAnswers: ['b', '(b)', 'however'],
      explanation: '後一句說「錢帶來的快樂有限」，和前一句「錢帶來快樂」相反 → 需要 However；而且 this 回指前一句的觀點，所以選 (B)。'
    }
  };
}

// 📖 閱讀測驗
function buildReadingCoach(stem, opts, wKey, wText, cKey, cText) {
  const rq = detectReadingQuestionType(stem);
  const question = extractReadingQuestion(stem);
  const stopWords = ['What', 'Which', 'Who', 'Why', 'How', 'When', 'Where', 'According', 'The', 'It', 'In', 'A', 'An'];
  const keywords = (question.match(/\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\b|\b\d{2,4}\b|"[^"]+"|“[^”]+”|paragraph\s+\d+|line\s+\d+/gi) || [])
    .filter((w) => !stopWords.includes(w))
    .slice(0, 6);

  const deconstructHtml = `
    <div>❓ <strong>題目問句：</strong>${escapeHtml(question || '請在題目欄貼上閱讀題的問句（可連同文章）')}</div>
    <div>🧭 <strong>題型：</strong>${escapeHtml(rq.label)}</div>
    ${keywords.length ? `<div>🔍 <strong>定位關鍵字：</strong><span style="color:#b45309; font-weight:700;">${escapeHtml(keywords.join('、'))}</span></div>` : ''}
  `;

  const hint1Map = {
    main: '主旨題：看第一段（尤其最後一句）與每段第一句，找出「全文都在講」的主題，不要被單一段落的細節吸走。',
    detail: `細節題：用定位關鍵字${keywords.length ? `「${keywords.join('、')}」` : ''}回原文找到那一句，再逐字比對選項。`,
    inference: '推論題：答案是原文的「合理延伸」，一定能在原文找到支持句；沒寫到的不要腦補。',
    vocab: '字義題：回到該字所在句，把每個選項代進去，看哪個意思和前後文的同義／反義／舉例線索最吻合。',
    reference: '指代題：從指代詞往「前面」找最近、單複數一致的名詞或整句概念。',
    attitude: '態度題：圈出作者用的形容詞／副詞（unfortunately, surprisingly, admirable…），判斷正面、負面還是中立。',
    notTrue: '是非題（NOT / EXCEPT）：每個選項都回原文標記 ⭕ / ❌ / 沒提到，選出唯一不符的那一個。'
  };

  const keys = ['A', 'B', 'C', 'D', 'E'].filter((k) => opts[k]);
  const hint2 = keys.length
    ? `【逐一檢查選項】：\n${keys
        .map((k) => {
          const absolute = /\b(always|never|all|only|must|completely|none|every|entirely)\b/i.test(opts[k]);
          return `• (${k}) ${shortText(opts[k], 90)}${absolute ? '\n   ⚠️ 有絕對字眼，常是陷阱，要特別確認原文' : ''}`;
        })
        .join('\n')}\n\n常見陷阱：①偷換主詞或因果 ②原文沒提到 ③範圍太大或太小 ④照抄原文單字卻改了意思`
    : '請在 STEP 1 填入選項，這裡會逐一提醒你常見的誘答陷阱（絕對字眼、偷換概念、過度推論）。';

  const hint3 =
    cKey || cText
      ? `【正解核對】正確答案：${cKey ? `(${cKey}) ` : ''}${cText}\n👉 回原文找出「支持這個選項的那一句」，並把它畫線；正解通常是那句話的同義改寫。${
          wText ? `\n你選的 ${wKey ? `(${wKey}) ` : ''}${shortText(wText, 80)}：找找看它是哪一種陷阱。` : ''
        }`
      : '請在 STEP 2 點選「我的誤選」與「正確答案」，這裡會出現正解核對方法。';

  const twin =
    rq.type === 'reference'
      ? {
          q: '【舉一反三】判斷題型：「Which of the following is the best title for the passage?」這是哪一種閱讀題型？（主旨／細節／推論／字義／指代）',
          ans: ['主旨', '主旨題', '標題', '標題題'],
          exp: '問 best title → 主旨／標題題：看首段與各段第一句，選能涵蓋全文的選項。'
        }
      : {
          q: '【舉一反三】判斷題型：「What does the word "it" in the second paragraph refer to?」這是哪一種閱讀題型？（主旨／細節／推論／字義／指代）',
          ans: ['指代', '指代題'],
          exp: '問 refer to 指的是什麼 → 指代題：往前找最近、單複數一致的名詞。'
        };

  return {
    deconstructHtml,
    hint1: hint1Map[rq.type] || hint1Map.detail,
    hint2,
    hint3,
    twinQ: twin.q,
    twinSolution: { expectedAnswers: twin.ans, explanation: twin.exp }
  };
}

function detectChineseTenseHint(stem) {
  const s = String(stem || '');
  if (/已經|一直|自從|至今|到現在|過了/.test(s)) return '有「已經／一直／自從」→ 現在完成式 have/has + p.p.';
  if (/昨天|去年|上週|上個|以前|當時|那時|曾經|剛才|小時候|\d+\s*年前|前天/.test(s)) return '有過去時間詞 → 過去式';
  if (/明天|明年|下週|下個|將會|即將|未來|將/.test(s)) return '有未來時間詞 → will + V / be going to V';
  if (/正在|此刻|現在正/.test(s)) return '有「正在」→ 進行式 be + V-ing';
  return '沒有明顯時間詞 → 一般事實或習慣用現在簡單式（注意第三人稱單數 -s）';
}

// ✍️ 翻譯題（中翻英）
function buildTranslationCoach(stem, wrongAns, correctAns) {
  const patterns = detectTranslationPatterns(correctAns);
  const tenseHint = detectChineseTenseHint(stem);

  const deconstructHtml = `
    <div>📌 <strong>中文題目：</strong>${escapeHtml(stem || '請在題目欄輸入要翻譯的中文句子')}</div>
    <div>⏰ <strong>時態線索：</strong>${escapeHtml(tenseHint)}</div>
    <div>🧩 <strong>核心句型：</strong>${escapeHtml(
      patterns.length ? patterns.map((p) => p.name).join('、') : '（在 STEP 2 輸入標準答案後自動辨識句型）'
    )}</div>
  `;

  const hint1 = `【先搭骨架】\n1. 主詞是誰？（中文常省略主詞，英文一定要寫出來）\n2. 主要動詞是什麼？\n3. 時態：${tenseHint}\n4. 有沒有固定句型？${
    patterns.length ? patterns.map((p) => p.name).join('、') : '（例如 It takes … to V、so … that、not only … but also）'
  }`;

  let hint2;
  if (wrongAns && correctAns) {
    const diff = computeWordDiffAnalysis(wrongAns, correctAns);
    hint2 = `【你的寫法 vs. 標準答案】\n❌ 多寫／寫錯：${diff.wrongWords.join(' ') || '（無）'}\n✅ 漏寫／應改為：${diff.missingWords.join(' ') || '（無）'}${
      patterns.length ? `\n\n句型重點：\n${patterns.map((p) => `• ${p.name}：${p.note}`).join('\n')}` : ''
    }`;
  } else {
    hint2 = '先自己寫出完整英文句子，再在 STEP 2 輸入「你的寫法」與「標準答案」，這裡會逐字比對差在哪裡。';
  }

  const hint3 = correctAns
    ? `【標準譯句】${correctAns}\n✔ 扣分檢查清單：句首大寫、句尾標點、主詞動詞一致（三單 -s）、時態、冠詞 a / an / the、名詞單複數、拼字。`
    : '請在 STEP 2 輸入標準答案。';

  const twin = patterns[0]?.twin;
  const twinQ = twin
    ? `【舉一反三】${twin.q}`
    : `【舉一反三】遮住標準答案，憑記憶把整句英文重寫一次：\n「${stem || '（本題中文）'}」`;
  const twinSolution = twin
    ? { expectedAnswers: twin.ans, explanation: `參考答案：${twin.ans[0]}（${patterns[0].note}）` }
    : {
        expectedAnswers: correctAns ? [correctAns.toLowerCase().replace(/[.!?]\s*$/, '')] : [],
        explanation: correctAns ? `標準譯句：${correctAns}` : '請先輸入標準答案。'
      };

  return { deconstructHtml, hint1, hint2, hint3, twinQ, twinSolution };
}

// =========================================================================
// 3-Tier English Vocabulary, Live Dictionary & Context Clue Engine
// =========================================================================
const ENGLISH_VOCAB_DICT = {
  // High-frequency high school exam vocabulary & phrases
  faithful: {
    pos: 'adj.',
    zh: '忠實的、忠誠的、守信的',
    syn: 'loyal, devoted, dependable, true',
    usage: '修飾員工、朋友或支持者，表示長期忠心耿耿、值得信賴'
  },
  potential: {
    pos: 'adj. / n.',
    zh: '潛在的、可能的（尚未發生的）；潛力',
    syn: 'possible, likely, future, prospective',
    usage: '指「未來可能成為...但目前還不是」（如 potential customer 潛在客戶），不能用來形容已經任職十年的資深員工'
  },
  dependable: {
    pos: 'adj.',
    zh: '可靠的、值得信賴的',
    syn: 'reliable, trustworthy, faithful',
    usage: '形容人做事穩重可靠'
  },
  committed: {
    pos: 'adj.',
    zh: '盡心盡力的、堅定投入的',
    syn: 'devoted, dedicated, loyal',
    usage: '常搭配 committed to + N / V-ing（全心投入於...）'
  },
  employee: {
    pos: 'n.',
    zh: '員工、受僱者',
    syn: 'worker, staff member',
    usage: '指受僱於公司的人（雇主則為 employer）'
  },
  employer: {
    pos: 'n.',
    zh: '雇主、老闆',
    syn: 'boss, manager, company',
    usage: '指雇用員工的人或機構'
  },
  artificial: {
    pos: 'adj.',
    zh: '人造的、人工的、假裝的',
    syn: 'man-made, synthetic, fake',
    usage: '反義詞為 natural（天然的）'
  },
  reluctant: {
    pos: 'adj.',
    zh: '不情願的、勉強的',
    syn: 'unwilling, hesitant',
    usage: '常搭配 be reluctant to V（不情願做某事）'
  },
  temporary: {
    pos: 'adj.',
    zh: '暫時的、臨時的',
    syn: 'short-term, brief',
    usage: '反義詞為 permanent（永久的）'
  },
  permanent: {
    pos: 'adj.',
    zh: '永久的、長久的',
    syn: 'lasting, enduring, eternal',
    usage: '反義詞為 temporary（暫時的）'
  },
  enthusiastic: {
    pos: 'adj.',
    zh: '熱情的、熱心的',
    syn: 'eager, passionate, keen',
    usage: '常搭配 be enthusiastic about（對...充滿熱情）'
  },
  capable: {
    pos: 'adj.',
    zh: '有能力的、能幹的',
    syn: 'able, competent',
    usage: '常搭配 be capable of + N / V-ing'
  },
  responsible: {
    pos: 'adj.',
    zh: '負責任的、應負責任的',
    syn: 'accountable, reliable',
    usage: '常搭配 be responsible for（對...負責）'
  },
  available: {
    pos: 'adj.',
    zh: '可獲得的、有空的、可用的',
    syn: 'accessible, obtainable, free',
    usage: '指物品可買到或人有空出席'
  },
  essential: {
    pos: 'adj.',
    zh: '必要的、極為重要的、本質的',
    syn: 'vital, crucial, necessary',
    usage: '常搭配 it is essential that S + (should) + V'
  },
  significant: {
    pos: 'adj.',
    zh: '重大的、顯著的、有意義的',
    syn: 'important, notable, considerable',
    usage: '形容變化、影響或成就很顯著'
  },
  considerate: {
    pos: 'adj.',
    zh: '體貼的、設想周到的',
    syn: 'thoughtful, kind, caring',
    usage: '注意區別：considerable 是「相當大的、可觀的」'
  },
  considerable: {
    pos: 'adj.',
    zh: '相當大的、可觀的（數量或程度）',
    syn: 'substantial, large, significant',
    usage: '注意區別：considerate 是「體貼的」'
  },
  sensitive: {
    pos: 'adj.',
    zh: '敏感的、易受影響的',
    syn: 'delicate, reactive',
    usage: '常搭配 be sensitive to；注意區別 sensible（明智的、合理的）'
  },
  sensible: {
    pos: 'adj.',
    zh: '明智的、合情合理的',
    syn: 'wise, reasonable, practical',
    usage: '注意區別 sensitive（敏感的）'
  },
  respectful: {
    pos: 'adj.',
    zh: '恭敬的、有禮貌的（對他人表示尊重）',
    syn: 'polite, courteous',
    usage: '區別：respectable（值得尊敬的、體面的）、respective（各自的）'
  },
  respectable: {
    pos: 'adj.',
    zh: '值得尊敬的、體面的',
    syn: 'reputable, honorable',
    usage: '區別：respectful（對人恭敬有禮的）'
  },
  respective: {
    pos: 'adj.',
    zh: '各自的、分別的',
    syn: 'individual, separate',
    usage: '例如 return to their respective homes（回到各自的家）'
  },
  'used for': {
    pos: 'phr.',
    zh: '被用來作為...（表用途，主詞通常是物品）',
    syn: 'utilized for, employed for',
    usage: 'be used for + V-ing / N'
  },
  'keen on': {
    pos: 'phr.',
    zh: '熱衷於、非常喜歡',
    syn: 'fond of, enthusiastic about',
    usage: 'be keen on + V-ing / N'
  },
  'tired of': {
    pos: 'phr.',
    zh: '對...感到厭倦、厭煩',
    syn: 'sick of, bored with, fed up with',
    usage: 'be tired of + V-ing / N'
  },
  'amazed by': {
    pos: 'phr.',
    zh: '對...感到驚奇、讚嘆',
    syn: 'surprised by, astonished at',
    usage: 'be amazed by / at + N'
  },
  'for a while': {
    pos: 'phr.',
    zh: '持續一小段時間',
    syn: 'for a short time',
    usage: '表時間持續一陣子'
  },
  'for example': {
    pos: 'phr.',
    zh: '舉例來說',
    syn: 'for instance',
    usage: '用於舉出具體例子說明前句'
  },
  'in the end': {
    pos: 'phr.',
    zh: '最後、終於',
    syn: 'finally, eventually, at last',
    usage: '表經過一番過程後的最終結果'
  },
  'all of a sudden': {
    pos: 'phr.',
    zh: '突然間、出乎意料地',
    syn: 'suddenly, unexpectedly',
    usage: '表突發事件'
  },
  'go through': {
    pos: 'phr.',
    zh: '經歷（困難／痛苦）、仔細檢查',
    syn: 'experience, undergo, examine',
    usage: '後接難關、考驗或文件'
  },
  'carry out': {
    pos: 'phr.',
    zh: '執行、貫徹（計畫／實驗／承諾）',
    syn: 'conduct, perform, execute',
    usage: '常用於 carry out a plan / survey / experiment'
  },
  'call on': {
    pos: 'phr.',
    zh: '呼籲、號召；拜訪',
    syn: 'appeal to, urge, visit',
    usage: 'call on sb. to V（呼籲某人做某事）'
  },
  'come across': {
    pos: 'phr.',
    zh: '偶然遇見、無意間發現',
    syn: 'run into, encounter by chance',
    usage: '不接被動語態'
  },
  because: {
    pos: 'conj.',
    zh: '因為（從屬連接詞）',
    syn: 'since, as',
    usage: '後面必須接完整子句（主詞 S + 動詞 V）'
  },
  'because of': {
    pos: 'prep.',
    zh: '因為、由於（介系詞片語）',
    syn: 'due to, owing to, on account of',
    usage: '後面只能接名詞 (N) 或動名詞 (V-ing)，不能接完整子句'
  },
  // U15 文意選填 (A)~(J) 太空旅遊與高頻字彙
  flight: {
    pos: 'n.',
    zh: '飛行、航程、航班',
    syn: 'flying, journey, trip',
    usage: 'every flight（每一次飛行），in-flight 機上的'
  },
  managed: {
    pos: 'v. (p.t./p.p.)',
    zh: '設法做到、成功完成（manage to V）',
    syn: 'succeeded in, accomplished',
    usage: 'manage to V 表示「經過努力終於設法做到某事」'
  },
  spot: {
    pos: 'v. / n.',
    zh: '看見、注意到、認出；地點、斑點',
    syn: 'notice, see, detect, observe',
    usage: 'not to spot the visible human activity（不注意到肉眼可見的人類活動是不可能的）'
  },
  craft: {
    pos: 'n. / v.',
    zh: '太空船、飛行器、船隻；工藝',
    syn: 'spacecraft, vessel, aircraft',
    usage: 'in a two-seater craft in outer space（在外太空的雙人座太空船中，單複數同形）'
  },
  destination: {
    pos: 'n.',
    zh: '目的地、終點',
    syn: 'goal, target, end point',
    usage: 'final destination（最終目的地）'
  },
  impulse: {
    pos: 'n.',
    zh: '衝動、驅使、脈衝',
    syn: 'urge, drive, instinct',
    usage: 'driven by the impulse to fly（受到想要飛行的衝動所驅使）'
  },
  thrilling: {
    pos: 'adj.',
    zh: '令人興奮的、驚險刺激的',
    syn: 'exciting, breathtaking, exhilarating',
    usage: '修飾事物（如 the thrilling clouds 令人屏息讚嘆的雲層）'
  },
  missions: {
    pos: 'n. (pl.)',
    zh: '任務、使命、太空任務',
    syn: 'assignments, tasks, expeditions',
    usage: 'space missions（太空任務），respond to missions in...（執行在各國的任務）'
  },
  everlasting: {
    pos: 'adj.',
    zh: '永恆的、持久不衰的、無止境的',
    syn: 'eternal, endless, permanent, perpetual',
    usage: 'everlasting power of human industry（人類工業永無止境的力量）'
  },
  dazzled: {
    pos: 'adj. / v. (p.p.)',
    zh: '感到目眩神迷的、驚嘆不已的（be dazzled by）',
    syn: 'amazed, astonished, impressed, awed',
    usage: 'I am dazzled by the sheer beauty of nature（我被大自然純粹的美景震懾／深深吸引）'
  },
  sheer: {
    pos: 'adj.',
    zh: '純粹的、完全的、十足的',
    syn: 'pure, utter, complete, absolute',
    usage: 'the sheer beauty of nature（大自然純粹而極致的美）'
  },
  breathtaking: {
    pos: 'adj.',
    zh: '令人屏息的、壯觀動人的',
    syn: 'stunning, magnificent, thrilling',
    usage: 'provide a breathtaking view（提供令人屏息的壯麗景色）'
  },
  ecosystems: {
    pos: 'n. (pl.)',
    zh: '生態系統',
    syn: 'ecological systems, environments',
    usage: 'harm our ecosystems（傷害我們的生態系統）'
  },
  vessels: {
    pos: 'n. (pl.)',
    zh: '船隻、大型艦艇、太空船；血管',
    syn: 'ships, boats, craft',
    usage: 'These vessels reveal...（這些大型船隻展現出...）'
  },
  venture: {
    pos: 'n. / v.',
    zh: '冒險事業、大膽嘗試；冒險前往',
    syn: 'adventure, enterprise, undertaking',
    usage: 'for my spaceflight venture（為了我的太空飛行冒險事業）'
  },
  ambitious: {
    pos: 'adj.',
    zh: '有野心的、雄心勃勃的、宏大的',
    syn: 'aspiring, grand, determined',
    usage: 'This new mission is highly ambitious（這項新任務極具雄心壯志）'
  }
};

const liveDictCache = {};
const liveSentenceCache = {};
const inFlightTranslations = new Set();

function normalizeLookupKey(word) {
  return (word || '')
    .trim()
    .replace(/^\([A-Ea-e1-5]\)\s*/, '')
    .toLowerCase();
}

function getEnglishWordEntry(word) {
  const key = normalizeLookupKey(word);
  if (!key) return null;
  if (ENGLISH_VOCAB_DICT[key]) return ENGLISH_VOCAB_DICT[key];
  if (liveDictCache[key]) return liveDictCache[key];
  return null;
}

function getCachedSentenceTranslation(sentence) {
  const clean = (sentence || '').trim();
  if (!clean) return '';
  return liveSentenceCache[clean] || '';
}

function mapGooglePosToAbbrev(posRaw) {
  const p = (posRaw || '').toLowerCase();
  if (p.includes('adjective')) return 'adj. 形容詞';
  if (p.includes('noun')) return 'n. 名詞';
  if (p.includes('verb') && !p.includes('adverb')) return 'v. 動詞';
  if (p.includes('adverb')) return 'adv. 副詞';
  if (p.includes('preposition')) return 'prep. 介系詞';
  if (p.includes('conjunction')) return 'conj. 連接詞';
  if (p.includes('pronoun')) return 'pron. 代名詞';
  return posRaw || '';
}

async function fetchGoogleDictEntry(word) {
  const key = normalizeLookupKey(word);
  if (!key || ENGLISH_VOCAB_DICT[key] || liveDictCache[key]) return;
  const inflightKey = `word:${key}`;
  if (inFlightTranslations.has(inflightKey)) return;
  inFlightTranslations.add(inflightKey);

  try {
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=zh-TW&dt=t&dt=bd&q=${encodeURIComponent(
      key
    )}`;
    const res = await fetch(url);
    if (!res.ok) return;
    const data = await res.json();

    const basicTrans = Array.isArray(data?.[0])
      ? data[0]
          .map((seg) => seg?.[0] || '')
          .join('')
          .trim()
      : '';

    let posList = [];
    let zhMeanings = [];
    let synonyms = [];

    if (Array.isArray(data?.[1])) {
      for (const block of data[1]) {
        const pos = mapGooglePosToAbbrev(block?.[0]);
        if (pos) posList.push(pos);
        if (Array.isArray(block?.[1])) {
          zhMeanings.push(...block[1].slice(0, 4));
        }
        if (Array.isArray(block?.[2])) {
          for (const rev of block[2].slice(0, 2)) {
            if (Array.isArray(rev?.[1])) {
              for (const synWord of rev[1]) {
                if (
                  synWord &&
                  synWord.toLowerCase() !== key &&
                  synonyms.length < 4 &&
                  !synonyms.includes(synWord)
                ) {
                  synonyms.push(synWord);
                }
              }
            }
          }
        }
      }
    }

    const uniqueZh = [...new Set(zhMeanings)].slice(0, 4).join('、');
    const finalZh =
      uniqueZh ||
      (basicTrans && basicTrans.toLowerCase() !== key ? basicTrans : '');
    if (finalZh) {
      liveDictCache[key] = {
        pos: posList[0] || guessEnglishPartOfSpeech(key),
        zh: finalZh,
        syn: synonyms.join(', '),
        usage: ''
      };
    }
  } catch (err) {
    console.warn('Live dictionary lookup skipped:', err);
  } finally {
    inFlightTranslations.delete(inflightKey);
  }
}

async function fetchGoogleSentenceTranslation(sentence) {
  const clean = (sentence || '').trim();
  if (!clean || liveSentenceCache[clean]) return;
  const asciiCount = (clean.match(/[A-Za-z]/g) || []).length;
  if (asciiCount < 8) return;

  const inflightKey = `sent:${clean}`;
  if (inFlightTranslations.has(inflightKey)) return;
  inFlightTranslations.add(inflightKey);

  try {
    const querySent = clean
      .replace(/^\d+[\.\)]\s*/, '')
      .replace(/_{2,}/g, ' ___ ');
    const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=zh-TW&dt=t&q=${encodeURIComponent(
      querySent
    )}`;
    const res = await fetch(url);
    if (!res.ok) return;
    const data = await res.json();
    if (Array.isArray(data?.[0])) {
      const translated = data[0]
        .map((seg) => seg?.[0] || '')
        .join('')
        .trim();
      if (translated && translated !== querySent) {
        liveSentenceCache[clean] = translated;
      }
    }
  } catch (err) {
    console.warn('Sentence translation skipped:', err);
  } finally {
    inFlightTranslations.delete(inflightKey);
  }
}

// Tracks every lookup we've already tried, so a failed network request never loops forever
const attemptedLiveLookups = new Set();

function ensureLiveEnglishTranslations(words = [], sentence = '') {
  const tasks = [];
  const queueWord = (raw) => {
    const key = normalizeLookupKey(raw);
    if (!key || !/[a-zA-Z]/.test(key) || key.length > 60) return;
    if (ENGLISH_VOCAB_DICT[key] || liveDictCache[key]) return;
    const tag = `word:${key}`;
    if (attemptedLiveLookups.has(tag) || inFlightTranslations.has(tag)) return;
    attemptedLiveLookups.add(tag);
    tasks.push(fetchGoogleDictEntry(key));
  };

  words.forEach(queueWord);

  // Also translate key content words from the sentence so context clues are rich
  const skipWords = ['which', 'where', 'there', 'their', 'would', 'could', 'should', 'about', 'after', 'before', 'always', 'years', 'company'];
  ((sentence || '').match(/[a-zA-Z]{5,}/g) || [])
    .map((w) => w.toLowerCase())
    .filter((w) => !skipWords.includes(w))
    .slice(0, 5)
    .forEach(queueWord);

  const cleanSent = (sentence || '').trim();
  if (cleanSent && cleanSent.length <= 600 && !liveSentenceCache[cleanSent]) {
    const tag = `sent:${cleanSent}`;
    if (!attemptedLiveLookups.has(tag) && !inFlightTranslations.has(tag)) {
      attemptedLiveLookups.add(tag);
      tasks.push(fetchGoogleSentenceTranslation(cleanSent));
    }
  }

  if (tasks.length > 0) {
    Promise.all(tasks).then(() => {
      updateSelfLearningCoach(false);
      const wVal = document.getElementById('fieldWrongAnswer')?.value.trim();
      const cVal = document.getElementById('fieldCorrectAnswer')?.value.trim();
      if (wVal && cVal) {
        runAutoMistakeDiagnosis(true);
      }
    });
  }
}

function extractSentenceClueWords(stem) {
  if (!stem) return '';
  const clues = [];
  const lower = stem.toLowerCase();

  if (/for over ten years|for many years|for a long time|for decades/.test(lower)) {
    clues.push('for over ten years（任職／持續超過十年 ➔ 表長期穩定，非潛在未發生）');
  }
  if (/dependable/.test(lower)) {
    clues.push('dependable（adj. 可靠的、值得信賴的）');
  }
  if (/committed(?:\s+to)?/.test(lower)) {
    clues.push('committed to his work（全心投入工作、盡忠職守）');
  }
  if (/employee/.test(lower)) {
    clues.push('employee（n. 員工）');
  }
  if (/\b(?:am|is|are|was|were)\s+_{2,}\d*_{2,}\s+by\b|\bi\s+am\s+[\s\S]{0,18}\bby\s+the\s+sheer\s+beauty/i.test(lower)) {
    clues.push('be 動詞 (am/is/are) + 空格 + by the sheer beauty of nature（被動語態／情緒過去分詞 be + p.p. + by，表示「被大自然純粹的美所震懾／目眩神迷」）');
  }
  if (/impossible\s+not\s+to\s+_{2,}\d*_{2,}/i.test(lower)) {
    clues.push('impossible not to + 原形動詞 V（後面受詞為 the visible human activity，應填原形及物動詞 spot「看見／注意到」）');
  }
  if (/however|but|yet|although|though|despite|in spite of/.test(lower)) {
    clues.push('轉折語氣詞（前後文語意相反）');
  }
  if (/therefore|thus|as a result|consequently|because|since/.test(lower)) {
    clues.push('因果邏輯連接詞（前後文具因果推論關係）');
  }

  // Also append any other words from stem that have dictionary entries
  const tokens = stem.match(/[a-zA-Z]{5,}/g) || [];
  for (const tok of tokens) {
    const k = tok.toLowerCase();
    if (['dependable', 'committed', 'employee'].includes(k)) continue;
    const entry = getEnglishWordEntry(k);
    if (entry?.zh && clues.length < 4) {
      clues.push(`${tok}（${entry.zh.split('、')[0]}）`);
    }
  }

  return clues.join(' ＋ ');
}

function analyzeEnglishContextClues(
  stem,
  wrongAns,
  wrongText,
  wEntry,
  correctAns,
  correctText,
  cEntry,
  opts
) {
  const wMeaning = wEntry?.zh
    ? `「${wEntry.zh}」（${wEntry.pos || guessEnglishPartOfSpeech(wrongText)}${
        wEntry.syn ? `，同義詞：${wEntry.syn}` : ''
      }）`
    : `（${guessEnglishPartOfSpeech(wrongText)}，查詢字義中...）`;

  const cMeaning = cEntry?.zh
    ? `「${cEntry.zh}」（${cEntry.pos || guessEnglishPartOfSpeech(correctText)}${
        cEntry.syn ? `，同義詞：${cEntry.syn}` : ''
      }）`
    : `（${guessEnglishPartOfSpeech(correctText)}，查詢字義中...）`;

  const stemTrans = getCachedSentenceTranslation(stem);
  const clueStr = extractSentenceClueWords(stem);

  // Build option summary table for all options (A)~(J)
  const allOptsSummary = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
    .filter((k) => opts && opts[k])
    .map((k) => {
      const word = opts[k];
      const ent = getEnglishWordEntry(word);
      const zh = ent?.zh ? `${ent.pos ? ent.pos + ' ' : ''}${ent.zh}` : '查詢中...';
      return `(${k}) ${word}【${zh}】`;
    })
    .join(' ／ ');

  let whyWrong = `你選了 (${wrongAns}) ${wrongText}，它的中文意思是${wMeaning}。`;
  if (wEntry?.usage) {
    whyWrong += `\n👉 語意衝突點：${wEntry.usage}。`;
  } else if (clueStr) {
    whyWrong += `\n👉 語意衝突點：題幹給出的關鍵線索是「${clueStr}」，把「${
      wEntry?.zh || wrongText
    }」放進句中會與前後文邏輯矛盾！`;
  } else {
    whyWrong += `\n👉 語意衝突點：放進本題前後文時，與題幹描述的情境語意不符。`;
  }

  let correctConcept = `【這題應該這樣解】：\n`;
  if (stemTrans) {
    correctConcept += `1. 【整句中譯】：${stemTrans}\n`;
  }
  correctConcept += `${stemTrans ? '2' : '1'}. 【抓題幹關鍵線索】：${
    clueStr || '觀察空格前後的修飾對象、形容詞正負向語氣與補充說明子句'
  }。\n`;
  correctConcept += `${
    stemTrans ? '3' : '2'
  }. 【正解字義與呼應】：正確答案 (${correctAns}) ${correctText} 的意思是${cMeaning}${
    cEntry?.usage ? `（${cEntry.usage}）` : ''
  }，完全呼應題幹的語境線索！`;
  if (allOptsSummary) {
    correctConcept += `\n📌 【全選項字義一次必背】：${allOptsSummary}`;
  }

  const keyTakeaway = `必背單字對比：❌ (${wrongAns}) ${wrongText}＝${
    wEntry?.zh || '誤選字'
  } vs. ✅ (${correctAns}) ${correctText}＝${
    cEntry?.zh || '正解字'
  }${cEntry?.syn ? `（同義詞：${cEntry.syn}）` : ''}！`;

  return { whyWrong, correctConcept, keyTakeaway };
}

function isEnglishOptionSet(opts) {
  const combined = `${opts.A} ${opts.B} ${opts.C} ${opts.D}`.trim();
  if (!combined) return false;
  const asciiLetters = (combined.match(/[A-Za-z]/g) || []).length;
  return asciiLetters > combined.length * 0.45;
}

function isEnglishWordFamilyQuestion(opts) {
  const words = [opts.A, opts.B, opts.C, opts.D]
    .map((s) => s.trim().toLowerCase())
    .filter((s) => /^[a-z]+$/.test(s));
  if (words.length < 3) return false;
  const prefix3 = words[0].slice(0, 3);
  return prefix3.length === 3 && words.every((w) => w.startsWith(prefix3));
}

function guessEnglishPartOfSpeech(word) {
  const w = word.trim().toLowerCase();
  if (/(ness|tion|sion|ment|ity|ance|ence|ship|ist|er|or|ant|ings)$/.test(w)) {
    return '名詞 (n.)';
  }
  if (/ly$/.test(w)) {
    return '副詞 (adv.，修飾動詞或形容詞)';
  }
  if (/(en|ize|ise|fy|ate)$/.test(w)) {
    return '動詞 (v.)';
  }
  if (/(ive|ous|ful|less|able|ible|al|ic|ed|ing|ant|ent)$/.test(w)) {
    return '形容詞 / 分詞 (adj.)';
  }
  return '原形詞 / 形容詞或動詞';
}

// Free Keyless AI English Coach Enhancer (via Pollinations OpenAI-compatible API)
async function runFreeAiPollinationsDiagnosis() {
  const wrongAns = document.getElementById('fieldWrongAnswer').value.trim();
  const correctAns = document.getElementById('fieldCorrectAnswer').value.trim();
  const subject = '英文';
  const unit = document.getElementById('fieldUnit').value.trim() || '';
  const stem = document.getElementById('fieldQuestionText').value.trim();

  if (!wrongAns || !correctAns) {
    showToast('🤖 已更新題目拆解！請在 STEP 2 輸入「錯誤答案」與「正確答案」即可完成錯因比對。');
    return;
  }

  const badge = document.getElementById('autoDiagnosisStatusBadge');
  if (badge) {
    badge.textContent = '⏳ 正在透過內建英文引擎與 AI 教練產生深度解析...';
  }

  try {
    const resp = await fetch('https://text.pollinations.ai/openai', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: 'openai',
        messages: [
          {
            role: 'system',
            content:
              '你是台灣高中英文自主學習教練，專精單字選擇、克漏字、文意選填、篇章結構、閱讀測驗與中翻英翻譯。請針對學生的錯題與答案，回傳純 JSON（不要 markdown 語法）：{"whyWrong":"具體指出學生誤選或翻譯錯誤在哪（2句話）","correctConcept":"清楚說明正確文法、字義或前後文線索（2~3句話）","keyTakeaway":"一句話避坑重點"}。一律使用繁體中文。'
          },
          {
            role: 'user',
            content: `題型：${getQuestionTypeLabel(state.questionType)}（${unit}）\n題目：${stem}\n學生錯誤答案：${wrongAns}\n標準正確答案：${correctAns}`
          }
        ]
      })
    });
    if (resp.ok) {
      const data = await resp.json();
      const rawReply = data?.choices?.[0]?.message?.content || '';
      const jsonMatch = rawReply.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        const parsed = JSON.parse(jsonMatch[0]);
        if (parsed.whyWrong) {
          document.getElementById('fieldWhyWrong').value = parsed.whyWrong;
          document.getElementById('previewWhyWrong').textContent = parsed.whyWrong;
        }
        if (parsed.correctConcept) {
          document.getElementById('fieldCorrectConcept').value = parsed.correctConcept;
          document.getElementById('previewCorrectConcept').textContent = parsed.correctConcept;
        }
        if (parsed.keyTakeaway) {
          document.getElementById('fieldKeyTakeaway').value = parsed.keyTakeaway;
          document.getElementById('previewKeyTakeaway').textContent = parsed.keyTakeaway;
        }
        if (badge) badge.textContent = '✨ AI 自主學習教練已完成深度錯因與解法分析！';
        showToast('✨ AI 教練已針對您的答案完成深度錯因與解法解析！');
        return;
      }
    }
  } catch (e) {
    console.warn('Free AI enhancer skipped, using built-in subject engine:', e);
  }

  if (badge) {
    badge.textContent = `🤖 已根據【${subject}】內建專家引擎完成錯因診斷與正確解法解析！`;
  }
  showToast('🤖 已根據您的題型、答案與科目自動完成錯因診斷與正確寫法解析！');
}

// Optional Gemini Vision Deep Analysis (if user entered Gemini API Key)
async function runGeminiVisionDiagnosis(apiKey) {
  const wrongAns = document.getElementById('fieldWrongAnswer').value.trim();
  const correctAns = document.getElementById('fieldCorrectAnswer').value.trim();
  const subject = document.getElementById('fieldSubject').value.trim();
  const qNum = document.getElementById('fieldQuestionNum').value.trim();
  const stem = document.getElementById('fieldQuestionText').value.trim();
  const opts = {
    A: document.getElementById('fieldOptA').value.trim(),
    B: document.getElementById('fieldOptB').value.trim(),
    C: document.getElementById('fieldOptC').value.trim(),
    D: document.getElementById('fieldOptD').value.trim()
  };

  if (!wrongAns || !correctAns) {
    alert('請先在 STEP 2 輸入「我的錯誤答案／寫法」與「試卷正確答案」！');
    return;
  }

  const badge = document.getElementById('autoDiagnosisStatusBadge');
  if (badge) badge.textContent = '⏳ 正在透過 AI 視覺與學科專家引擎深度分析您的錯因...';

  try {
    const parts = [
      {
        text: `你是一位極具教學經驗的台灣高中「英文」名師。
本題題型：${getQuestionTypeLabel(state.questionType)}（題號：${qNum || '?'}）
學生當初寫錯／選錯的答案：「${wrongAns}」
標準正確答案：「${correctAns}」
目前題目文字：${stem}
選項（若為選擇題）：(A) ${opts.A} (B) ${opts.B} (C) ${opts.C} (D) ${opts.D}
請根據題目文字與附圖（若有），回傳純 JSON 格式（不要加 markdown 標籤）：
{
  "questionText": "若附圖中有更完整的英文題目或句子，請補全題目文字，否則回傳原題目",
  "whyWrong": "直接告訴學生他的答案「${wrongAns}」判斷或翻譯錯誤的地方在哪裡（具體指出單字字義、詞性、文法、轉折詞或翻譯句型問題，2~3句話）",
  "correctConcept": "告訴學生這一題應該怎麼看前後文線索、怎麼寫才正確（清楚步驟 2~3 句話）",
  "keyTakeaway": "給學生的一句話必背重點或句型提醒"
}`
      }
    ];

    if (state.currentCropDataUrl && state.currentCropDataUrl.startsWith('data:image/')) {
      const base64Data = state.currentCropDataUrl.split(',')[1];
      parts.push({
        inline_data: {
          mime_type: 'image/png',
          data: base64Data
        }
      });
    }

    const resp = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(apiKey)}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ contents: [{ parts }] })
      }
    );

    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = await resp.json();
    const rawReply = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    const jsonMatch = rawReply.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      const parsed = JSON.parse(jsonMatch[0]);
      if (parsed.questionText && parsed.questionText.length > 5) {
        document.getElementById('fieldQuestionText').value = parsed.questionText;
      }
      if (parsed.whyWrong) {
        document.getElementById('fieldWhyWrong').value = parsed.whyWrong;
        document.getElementById('previewWhyWrong').textContent = parsed.whyWrong;
      }
      if (parsed.correctConcept) {
        document.getElementById('fieldCorrectConcept').value = parsed.correctConcept;
        document.getElementById('previewCorrectConcept').textContent = parsed.correctConcept;
      }
      if (parsed.keyTakeaway) {
        document.getElementById('fieldKeyTakeaway').value = parsed.keyTakeaway;
        document.getElementById('previewKeyTakeaway').textContent = parsed.keyTakeaway;
      }
      if (badge) badge.textContent = '✨ AI 視覺深度診斷完成！';
      showToast('✨ AI 已根據題目截圖與您的答案完成深度診斷！');
    }
  } catch (err) {
    console.warn('Gemini Vision API fallback to built-in analyzer:', err);
    if (badge) badge.textContent = '🤖 已使用內建跨學科智慧診斷引擎完成分析！';
  }
}

function saveCurrentMistakeForm(silentRedirect = false) {
  const subject = '英文';
  const unit = document.getElementById('fieldUnit').value.trim() || '一般練習';
  const questionNum = document.getElementById('fieldQuestionNum').value.trim() || String(state.mistakes.length + 1);
  const questionText = document.getElementById('fieldQuestionText').value.trim();
  const optA = document.getElementById('fieldOptA').value.trim();
  const optB = document.getElementById('fieldOptB').value.trim();
  const optC = document.getElementById('fieldOptC').value.trim();
  const optD = document.getElementById('fieldOptD').value.trim();
  const optEEl = document.getElementById('fieldOptE');
  const optE = optEEl ? optEEl.value.trim() : '';
  const optF = document.getElementById('fieldOptF')?.value.trim() || '';
  const optG = document.getElementById('fieldOptG')?.value.trim() || '';
  const optH = document.getElementById('fieldOptH')?.value.trim() || '';
  const optI = document.getElementById('fieldOptI')?.value.trim() || '';
  const optJ = document.getElementById('fieldOptJ')?.value.trim() || '';
  const isChoiceLike = state.questionType !== 'translation';
  const wrongAnswer = isChoiceLike
    ? document.getElementById('fieldWrongAnswer').value.trim().toUpperCase()
    : document.getElementById('fieldWrongAnswer').value.trim();
  const correctAnswer = isChoiceLike
    ? document.getElementById('fieldCorrectAnswer').value.trim().toUpperCase()
    : document.getElementById('fieldCorrectAnswer').value.trim();

  if (!wrongAnswer || !correctAnswer) {
    alert('請在 STEP 2 輸入「❌ 我當初的錯誤答案／寫法」與「✅ 試卷正確答案」！');
    return false;
  }

  // Ensure auto-diagnosis has populated STEP 3 so student never has to type manually!
  if (!document.getElementById('fieldWhyWrong').value.trim()) {
    runAutoMistakeDiagnosis(true);
  }

  const whyWrong = document.getElementById('fieldWhyWrong').value.trim();
  const correctConcept = document.getElementById('fieldCorrectConcept').value.trim();
  const keyTakeaway = document.getElementById('fieldKeyTakeaway').value.trim();

  const errorTags =
    state.selectedErrorTags.size > 0
      ? Array.from(state.selectedErrorTags)
      : [getQuestionTypeLabel(state.questionType)];

  const editingId = document.getElementById('editingMistakeId').value;

  const mistakeRecord = {
    id: editingId || `M_${Date.now()}`,
    createdAt: new Date().toLocaleDateString('zh-TW'),
    questionType: state.questionType || 'choice',
    subject,
    unit,
    questionNum,
    questionText:
      questionText || `第 ${questionNum} 題（請參照試卷截圖作答）`,
    optA: isChoiceLike ? optA : '',
    optB: isChoiceLike ? optB : '',
    optC: isChoiceLike ? optC : '',
    optD: isChoiceLike ? optD : '',
    optE: isChoiceLike ? optE : '',
    optF: isChoiceLike ? optF : '',
    optG: isChoiceLike ? optG : '',
    optH: isChoiceLike ? optH : '',
    optI: isChoiceLike ? optI : '',
    optJ: isChoiceLike ? optJ : '',
    wrongAnswer,
    correctAnswer,
    errorTags,
    whyWrong,
    correctConcept,
    keyTakeaway,
    practiceCount: 0,
    correctCount: 0,
    masteryStatus: '待加強',
    cropImageDataUrl: state.currentCropDataUrl || '',
    isHistorical: false
  };

  if (editingId) {
    const idx = state.mistakes.findIndex((m) => m.id === editingId);
    if (idx !== -1) {
      mistakeRecord.practiceCount = state.mistakes[idx].practiceCount;
      mistakeRecord.correctCount = state.mistakes[idx].correctCount;
      mistakeRecord.masteryStatus = state.mistakes[idx].masteryStatus;
      state.mistakes[idx] = mistakeRecord;
    }
  } else {
    state.mistakes.push(mistakeRecord);
  }

  saveToLocalStorage();
  updateAllViews();
  resetMistakeForm(true);

  if (!silentRedirect) {
    showToast(
      `✅ 已自動分析並儲存第 ${questionNum} 題！可至「錯題診斷卡」下載 PNG 或開始「強化練習」。`
    );
  }
  return true;
}

function resetMistakeForm(keepMetadata = false) {
  document.getElementById('editingMistakeId').value = '';
  if (!keepMetadata) {
    document.getElementById('fieldSubject').value = '英文';
    document.getElementById('fieldUnit').value = '段考複習卷';
    setQuestionTypeUI('choice');
  }
  document.getElementById('fieldQuestionNum').value = '';
  document.getElementById('fieldQuestionText').value = '';
  ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach((k) => {
    const el = document.getElementById(`fieldOpt${k}`);
    if (el) el.value = '';
  });
  setOptionPillAndInput('wrong', '');
  setOptionPillAndInput('correct', '');
  document.getElementById('autoAnswerMatchHint').textContent = '';
  document.getElementById('fieldWhyWrong').value = '';
  document.getElementById('fieldCorrectConcept').value = '';
  document.getElementById('fieldKeyTakeaway').value = '';
  document.getElementById('previewWhyWrong').textContent =
    '請先在 STEP 2 輸入「我的錯誤答案／寫法」與「正確答案」...';
  document.getElementById('previewCorrectConcept').textContent =
    '系統將根據題型、科目與正確答案自動為你解析...';
  document.getElementById('previewKeyTakeaway').textContent =
    '自動產生重點提醒...';
  const diffBox = document.getElementById('liveTextDiffBox');
  if (diffBox) diffBox.style.display = 'none';
  const multiBar = document.getElementById('multiQuestionSelectorBar');
  if (multiBar) multiBar.style.display = 'none';
  setCroppedImagePreview('');

  state.selectedErrorTags.clear();
  document
    .querySelectorAll('#errorTagsContainer .diagnosis-tag')
    .forEach((b) => b.classList.remove('selected'));
  document.getElementById('btnSubmitMistake').textContent =
    '✅ 看懂了！存入錯題本與試算表';

  syncSubjectSpecificUI();
  updateSelfLearningCoach(false);
}

// ============================================================================
// STEP 2: Mistake Flashcards & High-Res PNG Download
// ============================================================================
function initCardsTab() {
  document
    .getElementById('filterCardSubject')
    .addEventListener('change', renderMistakeCards);
  document
    .getElementById('filterCardStatus')
    .addEventListener('change', renderMistakeCards);

  document
    .getElementById('btnGoToPracticeFromCards')
    .addEventListener('click', () => {
      switchTab('tab-practice');
    });

  document
    .getElementById('btnDownloadAllPng')
    .addEventListener('click', async () => {
      const filtered = getFilteredMistakesForCards();
      if (filtered.length === 0) {
        alert('目前沒有可下載的錯題卡！');
        return;
      }
      showToast(`📦 正在依序產生並下載 ${filtered.length} 張錯題 PNG 診斷卡...`);
      for (const m of filtered) {
        await downloadSingleMistakePng(m);
        await new Promise((r) => setTimeout(r, 350));
      }
    });
}

function getFilteredMistakesForCards() {
  const subj = document.getElementById('filterCardSubject').value;
  const status = document.getElementById('filterCardStatus').value;
  return state.mistakes.filter((m) => {
    if (subj !== 'ALL' && normalizeEnglishQType(m.questionType) !== subj) return false;
    if (status !== 'ALL' && m.masteryStatus !== status) return false;
    return true;
  });
}

function buildMistakeCardHtml(m, includeFooterButtons = true) {
  const hasOptions =
    m.optA || m.optB || m.optC || m.optD || m.optE ||
    m.optF || m.optG || m.optH || m.optI || m.optJ;
  const wrongTokens = parseMultiChoiceTokens(m.wrongAnswer);
  const correctTokens = parseMultiChoiceTokens(m.correctAnswer);
  const optionsHtml = hasOptions
    ? `<div class="mcard-options-list">
        ${['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
          .map((letter) => {
            const text = m[`opt${letter}`];
            if (!text) return '';
            const isWrong =
              wrongTokens.includes(letter) ||
              m.wrongAnswer.toUpperCase() === letter;
            const isCorrect =
              correctTokens.includes(letter) ||
              m.correctAnswer.toUpperCase() === letter;
            const cls = isCorrect
              ? 'is-correct'
              : isWrong
              ? 'is-wrong'
              : '';
            const badge =
              isCorrect && isWrong
                ? '☑️✅ '
                : isCorrect
                ? '✅ '
                : isWrong
                ? '❌ '
                : '';
            return `<div class="mcard-opt-item ${cls}">${badge}(${letter}) ${escapeHtml(text)}</div>`;
          })
          .join('')}
      </div>`
    : '';

  const cropImgHtml = m.cropImageDataUrl
    ? `<img class="mcard-crop-img" src="${m.cropImageDataUrl}" alt="題目截圖" />`
    : '';

  const masteryBadgeColor =
    m.masteryStatus === '已精通'
      ? '#10b981'
      : m.masteryStatus === '複習中'
      ? '#f59e0b'
      : '#ef4444';

  const footerHtml = includeFooterButtons
    ? `<div class="mcard-footer-actions">
        <div style="font-size:0.78rem; color:#64748b;">
          建立：${escapeHtml(m.createdAt)}｜已練習 ${m.practiceCount} 次（答對 ${m.correctCount} 次）｜<strong style="color:#4f46e5;">Made by Daphne Lin</strong>
        </div>
        <div style="display:flex; gap:0.4rem;">
          <button class="btn btn-secondary btn-sm" onclick="editMistakeById('${m.id}')">✏️ 編輯</button>
          <button class="btn btn-danger-outline btn-sm" onclick="deleteMistakeById('${m.id}')">🗑️</button>
          <button class="btn btn-primary btn-sm" onclick="downloadMistakePngById('${m.id}')">📥 下載錯題 PNG</button>
        </div>
      </div>`
    : `<div style="padding:0.65rem 1.15rem; background:#f8fafc; border-top:1px solid #e2e8f0; display:flex; justify-content:space-between; font-size:0.78rem; color:#64748b;">
        <span>📘 英文錯題進化論｜${escapeHtml(state.sheetName)}｜<strong style="color:#4f46e5;">✨ Made by Daphne Lin</strong></span>
        <span>記錄日期：${escapeHtml(m.createdAt)}｜熟練度：${escapeHtml(m.masteryStatus)}</span>
      </div>`;

  return `
    <div class="mistake-card-exportable" id="mcard-dom-${m.id}">
      <div class="mcard-header">
        <div style="display:flex; align-items:center; gap:0.5rem; flex-wrap:wrap;">
          <span class="mcard-subject-badge">${escapeHtml(getQuestionTypeLabel(m.questionType))}</span>
          <span style="font-weight:700; font-size:0.95rem;">${escapeHtml(m.unit)} — 第 ${escapeHtml(m.questionNum)} 題</span>
        </div>
        <span style="background:${masteryBadgeColor}; color:#fff; padding:0.2rem 0.6rem; border-radius:999px; font-size:0.76rem; font-weight:700;">
          ${escapeHtml(m.masteryStatus)}
        </span>
      </div>

      <div class="mcard-body">
        ${cropImgHtml}
        <div class="mcard-question-text">${escapeHtml(m.questionText)}</div>
        ${optionsHtml}

        <div class="mcard-vs-banner">
          <div class="mcard-ans-pill wrong">
            <span>❌ 我的誤選</span>
            <span style="font-size:1.05rem;">${escapeHtml(m.wrongAnswer)}</span>
          </div>
          <div class="mcard-ans-pill correct">
            <span>✅ 正確答案</span>
            <span style="font-size:1.05rem;">${escapeHtml(m.correctAnswer)}</span>
          </div>
        </div>

        <div class="mcard-diagnosis-block">
          <div>
            <strong style="color:#4f46e5;">🏷️ 錯誤根源診斷：</strong>
            <span>${m.errorTags.map((t) => `【${escapeHtml(t)}】`).join(' ')}</span>
          </div>
          <div>
            <strong style="color:#b91c1c;">🧐 為何選錯（破除迷思）：</strong>
            <span>${escapeHtml(m.whyWrong)}</span>
          </div>
          <div>
            <strong style="color:#047857;">🎯 正確觀念與解法：</strong>
            <span>${escapeHtml(m.correctConcept)}</span>
          </div>
        </div>

        <div class="mcard-key-takeaway">
          💡 避坑錦囊：${escapeHtml(m.keyTakeaway)}
        </div>
      </div>

      ${footerHtml}
    </div>
  `;
}

function renderMistakeCards() {
  const container = document.getElementById('mistakesCardsContainer');
  const emptyState = document.getElementById('cardsEmptyState');
  const filtered = getFilteredMistakesForCards();

  if (state.mistakes.length === 0) {
    emptyState.style.display = 'block';
    container.innerHTML = '';
    return;
  }

  emptyState.style.display = 'none';
  if (filtered.length === 0) {
    container.innerHTML = `<div class="panel-card" style="padding:2rem; text-align:center; grid-column:1/-1; color:#64748b;">沒有符合目前篩選條件的錯題卡。</div>`;
    return;
  }

  container.innerHTML = filtered
    .map((m) => buildMistakeCardHtml(m, true))
    .join('');
}

window.downloadMistakePngById = async function (id) {
  const m = state.mistakes.find((item) => item.id === id);
  if (!m) return;
  await downloadSingleMistakePng(m);
};

async function downloadSingleMistakePng(m) {
  const staging = document.getElementById('pngExportStaging');
  staging.innerHTML = buildMistakeCardHtml(m, false);
  const cardEl = staging.firstElementChild;

  try {
    if (window.html2canvas) {
      const canvas = await window.html2canvas(cardEl, {
        scale: 2,
        useCORS: true,
        backgroundColor: '#ffffff'
      });
      const dataUrl = canvas.toDataURL('image/png');
      triggerDownloadUrl(
        dataUrl,
        `錯題診斷卡_${m.subject}_${m.unit}_第${m.questionNum}題.png`
      );
      showToast(`🖼️ 已下載「第 ${m.questionNum} 題」PNG 錯題診斷卡！`);
    } else {
      // Fallback native Canvas renderer
      const fallbackUrl = renderFallbackMistakePng(m);
      triggerDownloadUrl(
        fallbackUrl,
        `錯題診斷卡_${m.subject}_第${m.questionNum}題.png`
      );
    }
  } catch (err) {
    console.error('PNG Export error:', err);
    const fallbackUrl = renderFallbackMistakePng(m);
    triggerDownloadUrl(
      fallbackUrl,
      `錯題診斷卡_${m.subject}_第${m.questionNum}題.png`
    );
  } finally {
    staging.innerHTML = '';
  }
}

function renderFallbackMistakePng(m) {
  const c = document.createElement('canvas');
  c.width = 800;
  c.height = 600;
  const ctx = c.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, c.width, c.height);
  ctx.fillStyle = '#312e81';
  ctx.fillRect(0, 0, c.width, 64);
  ctx.fillStyle = '#ffffff';
  ctx.font = 'bold 22px sans-serif';
  ctx.fillText(`${m.subject}｜${m.unit} - 第 ${m.questionNum} 題`, 24, 40);

  ctx.fillStyle = '#1e293b';
  ctx.font = 'bold 18px sans-serif';
  wrapTextOnCanvas(ctx, m.questionText, 24, 105, 750, 26);

  ctx.fillStyle = '#b91c1c';
  ctx.fillText(`❌ 我的誤選：${m.wrongAnswer}`, 24, 230);
  ctx.fillStyle = '#047857';
  ctx.fillText(`✅ 正確答案：${m.correctAnswer}`, 320, 230);

  ctx.fillStyle = '#334155';
  ctx.font = '16px sans-serif';
  wrapTextOnCanvas(ctx, `為何選錯：${m.whyWrong}`, 24, 285, 750, 24);
  wrapTextOnCanvas(ctx, `正確觀念：${m.correctConcept}`, 24, 375, 750, 24);

  ctx.fillStyle = '#92400e';
  ctx.font = 'bold 17px sans-serif';
  wrapTextOnCanvas(ctx, `💡 避坑錦囊：${m.keyTakeaway}`, 24, 485, 750, 24);
  return c.toDataURL('image/png');
}

window.editMistakeById = function (id) {
  const m = state.mistakes.find((item) => item.id === id);
  if (!m) return;
  switchTab('tab-record');

  document.getElementById('editingMistakeId').value = m.id;
  document.getElementById('fieldSubject').value = m.subject;
  document.getElementById('fieldUnit').value = m.unit;
  document.getElementById('fieldQuestionNum').value = m.questionNum;
  document.getElementById('fieldQuestionText').value = m.questionText;
  document.getElementById('fieldOptA').value = m.optA || '';
  document.getElementById('fieldOptB').value = m.optB || '';
  document.getElementById('fieldOptC').value = m.optC || '';
  document.getElementById('fieldOptD').value = m.optD || '';
  const optEEl = document.getElementById('fieldOptE');
  if (optEEl) optEEl.value = m.optE || '';
  ['F', 'G', 'H', 'I', 'J'].forEach((k) => {
    const el = document.getElementById(`fieldOpt${k}`);
    if (el) el.value = m[`opt${k}`] || '';
  });
  setQuestionTypeUI(
    m.questionType ||
      (m.optF || m.optG || m.optH || m.optI || m.optJ || /^[F-J]$/i.test(m.correctAnswer || '')
        ? 'wordbank'
        : 'choice')
  );
  setOptionPillAndInput('wrong', m.wrongAnswer);
  setOptionPillAndInput('correct', m.correctAnswer);
  document.getElementById('fieldWhyWrong').value = m.whyWrong;
  document.getElementById('fieldCorrectConcept').value = m.correctConcept;
  document.getElementById('fieldKeyTakeaway').value = m.keyTakeaway;
  document.getElementById('previewWhyWrong').textContent = m.whyWrong;
  document.getElementById('previewCorrectConcept').textContent = m.correctConcept;
  document.getElementById('previewKeyTakeaway').textContent = m.keyTakeaway;
  setCroppedImagePreview(m.cropImageDataUrl || '');

  state.selectedErrorTags = new Set(m.errorTags || []);
  document.querySelectorAll('#errorTagsContainer .diagnosis-tag').forEach((b) => {
    b.classList.toggle(
      'selected',
      state.selectedErrorTags.has(b.getAttribute('data-tag'))
    );
  });
  document.getElementById('btnSubmitMistake').textContent = '💾 更新這筆錯題紀錄';
  window.scrollTo({ top: 0, behavior: 'smooth' });
};

window.deleteMistakeById = function (id) {
  if (!confirm('確定要刪除這筆錯題紀錄嗎？')) return;
  state.mistakes = state.mistakes.filter((m) => m.id !== id);
  saveToLocalStorage();
  updateAllViews();
  showToast('🗑️ 已刪除該筆錯題紀錄');
};

// ============================================================================
// STEP 3: Mistake Practice & Concept Reinforcement Engine
// ============================================================================
function initPracticeQuizTab() {
  document
    .getElementById('quizFilterSubject')
    .addEventListener('change', updateQuizSetupSummary);
  document
    .getElementById('quizFilterMastery')
    .addEventListener('change', updateQuizSetupSummary);

  document.getElementById('btnStartQuiz').addEventListener('click', () => {
    startPracticeQuiz();
  });

  document.getElementById('btnQuitQuiz').addEventListener('click', () => {
    document.getElementById('quizActivePanel').style.display = 'none';
    document.getElementById('quizResultPanel').style.display = 'none';
    document.getElementById('quizSetupPanel').style.display = 'block';
    updateQuizSetupSummary();
  });

  document.getElementById('btnToggleQuizHint').addEventListener('click', () => {
    const hintBox = document.getElementById('quizHintBox');
    hintBox.style.display = hintBox.style.display === 'none' ? 'block' : 'none';
  });

  const btnSkeleton = document.getElementById('btnShowSkeletonHint');
  if (btnSkeleton) {
    btnSkeleton.addEventListener('click', () => {
      const m = state.quizQueue[state.quizIndex];
      if (!m) return;
      const hintSpan = document.getElementById('quizSkeletonHintText');
      if (!hintSpan) return;
      const ans = String(m.correctAnswer || '').trim();
      const tokens = ans.split(/\s+/).filter(Boolean);
      if (tokens.length >= 2) {
        // Sentence / phrase skeleton: show first letter of content words
        const skel = tokens
          .map((w, i) => {
            if (i === 0 || w.length <= 2) return w;
            const clean = w.replace(/[^A-Za-z0-9]/g, '');
            const punc = w.slice(clean.length);
            return clean[0] + '_'.repeat(Math.min(clean.length - 1, 5)) + punc;
          })
          .join(' ');
        hintSpan.textContent = `🔤 句型骨架提示（共 ${tokens.length} 字）：${skel}`;
      } else if (/^[A-Za-z-]+$/.test(ans) && ans.length >= 3) {
        hintSpan.textContent = `🔤 字彙提示：字首 ${ans[0]}... 字尾 ...${ans[ans.length - 1]}（共 ${ans.length} 個字母）`;
      } else {
        hintSpan.textContent = `💡 答案長度約 ${ans.length} 個字元，開頭為「${ans.slice(0, 2)}...」`;
      }
    });
  }

  document
    .getElementById('btnSubmitQuizAnswer')
    .addEventListener('click', submitCurrentQuizAnswer);

  document
    .getElementById('btnNextQuizQuestion')
    .addEventListener('click', () => {
      state.quizIndex++;
      if (state.quizIndex < state.quizQueue.length) {
        renderCurrentQuizQuestion();
      } else {
        finishPracticeQuiz();
      }
    });

  document.getElementById('btnBackToQuizSetup').addEventListener('click', () => {
    document.getElementById('quizResultPanel').style.display = 'none';
    document.getElementById('quizSetupPanel').style.display = 'block';
    updateQuizSetupSummary();
  });

  document.getElementById('btnRetryWrongOnly').addEventListener('click', () => {
    if (state.quizWrongItems.length > 0) {
      startPracticeQuiz(state.quizWrongItems);
    }
  });
}

function getFilteredPoolForQuiz() {
  const subj = document.getElementById('quizFilterSubject').value;
  const mastery = document.getElementById('quizFilterMastery').value;
  return state.mistakes.filter((m) => {
    if (subj !== 'ALL' && normalizeEnglishQType(m.questionType) !== subj) return false;
    if (mastery === 'UNMASTERED' && m.masteryStatus === '已精通') return false;
    return true;
  });
}

function updateQuizSetupSummary() {
  const pool = getFilteredPoolForQuiz();
  const infoEl = document.getElementById('quizAvailableCountInfo');
  const startBtn = document.getElementById('btnStartQuiz');
  infoEl.textContent = `目前符合條件共有 ${pool.length} 題錯題可供強化練習`;
  startBtn.disabled = pool.length === 0;
}

function startPracticeQuiz(customList = null) {
  const pool = customList ? [...customList] : getFilteredPoolForQuiz();
  if (pool.length === 0) {
    alert('目前沒有符合篩選條件的錯題可供練習！請先新增錯題或調整篩選條件。');
    return;
  }

  const shuffleQ = document.getElementById('chkShuffleQuestions').checked;
  if (shuffleQ) {
    pool.sort(() => Math.random() - 0.5);
  }

  state.quizQueue = pool;
  state.quizIndex = 0;
  state.quizCorrectCount = 0;
  state.quizWrongItems = [];

  document.getElementById('quizSetupPanel').style.display = 'none';
  document.getElementById('quizResultPanel').style.display = 'none';
  document.getElementById('quizActivePanel').style.display = 'block';

  renderCurrentQuizQuestion();
}

function isNonChoiceQuizQuestion(m) {
  if (m.questionType === 'translation' || m.questionType === 'fill') {
    return true;
  }
  if (m.questionType === 'multiselect' || m.questionType === 'wordbank') {
    return false;
  }
  const stdKeys = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', '1', '2', '3', '4', '5'];
  const cUp = String(m.correctAnswer || '').trim().toUpperCase();
  const cTokens = parseMultiChoiceTokens(cUp);
  const hasOpts = Boolean(
    m.optA || m.optB || m.optC || m.optD || m.optE ||
    m.optF || m.optG || m.optH || m.optI || m.optJ
  );
  if (!hasOpts && cTokens.length === 0 && !stdKeys.includes(cUp)) {
    return true;
  }
  return false;
}

function renderCurrentQuizQuestion() {
  const m = state.quizQueue[state.quizIndex];
  state.quizSelectedOptKey = null;
  state.quizSelectedOptKeys = new Set();
  state.quizSelectedConceptIdx = null;
  state.quizSubmitted = false;

  const isMultiSelectQuiz = m.questionType === 'multiselect';
  const qTypeBadge = getQuestionTypeLabel(
    isMultiSelectQuiz ? 'multiselect' : m.questionType
  );
  document.getElementById('quizProgressBadge').textContent = `第 ${
    state.quizIndex + 1
  } / ${state.quizQueue.length} 題`;
  document.getElementById(
    'quizSubjectTitle'
  ).textContent = `${m.subject}｜${m.unit}（第 ${m.questionNum} 題・${qTypeBadge}）`;

  document.getElementById('quizErrorTagsBanner').textContent = `⚠️ 上次錯因紀錄：${m.errorTags.join('、')}`;

  const showHintDefault = document.getElementById('chkShowHintFirst').checked;
  const hintBox = document.getElementById('quizHintBox');
  hintBox.textContent = `💡 你給自己的避坑錦囊：${m.keyTakeaway}`;
  hintBox.style.display = showHintDefault ? 'block' : 'none';

  const imgWrap = document.getElementById('quizImageWrapper');
  const imgEl = document.getElementById('quizQuestionImg');
  if (m.cropImageDataUrl) {
    imgEl.src = m.cropImageDataUrl;
    imgWrap.style.display = 'block';
  } else {
    imgWrap.style.display = 'none';
  }

  document.getElementById('quizQuestionStem').textContent = isMultiSelectQuiz
    ? `${m.questionText}（☑️ 多選題：可點選多個正確選項）`
    : m.questionText;

  const optionsContainer = document.getElementById('quizOptionsContainer');
  const textInputContainer = document.getElementById('quizTextInputContainer');
  const typedInput = document.getElementById('quizUserTypedAnswer');
  const skelHintSpan = document.getElementById('quizSkeletonHintText');

  const useTextInputMode = isNonChoiceQuizQuestion(m);

  if (useTextInputMode && textInputContainer && typedInput) {
    optionsContainer.style.display = 'none';
    textInputContainer.style.display = 'block';
    typedInput.value = '';
    typedInput.disabled = false;
    if (skelHintSpan) skelHintSpan.textContent = '';
    const promptLbl = document.getElementById('quizTextInputPromptLabel');
    if (promptLbl) {
      promptLbl.textContent =
        m.questionType === 'translation'
          ? '✍️ 請輸入您的翻譯／非選題完整作答（支援智慧字詞比對）：'
          : '🔤 請輸入您的填充／字彙／計算答案（不必死背選項，直接作答）：';
    }
  } else {
    if (textInputContainer) textInputContainer.style.display = 'none';
    optionsContainer.style.display = 'grid';

    // Prepare Options (with optional anti-rote-memorization shuffling)
    const shuffleOpts = document.getElementById('chkShuffleOptions').checked;
    const isWordBankQuiz =
      m.questionType === 'wordbank' ||
      Boolean(m.optF || m.optG || m.optH || m.optI || m.optJ) ||
      /^[F-J]$/i.test(m.correctAnswer || '') ||
      /^[F-J]$/i.test(m.wrongAnswer || '');
    const hasOptionText = Boolean(
      m.optA || m.optB || m.optC || m.optD || m.optE ||
      m.optF || m.optG || m.optH || m.optI || m.optJ
    );
    const useNumLabel = false;

    let renderedOptions = [
      { origKey: 'A', text: m.optA || (useNumLabel ? '選項 (1)' : '選項 (A)') },
      { origKey: 'B', text: m.optB || (useNumLabel ? '選項 (2)' : '選項 (B)') },
      { origKey: 'C', text: m.optC || (useNumLabel ? '選項 (3)' : '選項 (C)') },
      { origKey: 'D', text: m.optD || (useNumLabel ? '選項 (4)' : '選項 (D)') }
    ];
    if (isMultiSelectQuiz || isWordBankQuiz || m.optE || m.questionType === 'structure') {
      renderedOptions.push({
        origKey: 'E',
        text: m.optE || (useNumLabel ? '選項 (5)' : '選項 (E)')
      });
    }
    if (isWordBankQuiz || m.questionType === 'structure') {
      ['F', 'G', 'H', 'I', 'J'].forEach((k) => {
        if (m[`opt${k}`] || m.correctAnswer?.toUpperCase() === k || m.wrongAnswer?.toUpperCase() === k) {
          renderedOptions.push({
            origKey: k,
            text: m[`opt${k}`] || `選項 (${k})`
          });
        }
      });
    }

    const stdKeys = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
    const cTokens = parseMultiChoiceTokens(m.correctAnswer);
    if (
      !isMultiSelectQuiz &&
      !isWordBankQuiz &&
      cTokens.length === 0 &&
      (!stdKeys.includes(m.correctAnswer.toUpperCase()) ||
        !stdKeys.includes(m.wrongAnswer.toUpperCase()))
    ) {
      renderedOptions = [
        { origKey: m.correctAnswer, text: m.correctAnswer },
        { origKey: m.wrongAnswer, text: m.wrongAnswer }
      ];
      if (m.optA) renderedOptions.push({ origKey: 'A', text: m.optA });
      if (m.optB) renderedOptions.push({ origKey: 'B', text: m.optB });
    }

    if (
      shuffleOpts &&
      !isMultiSelectQuiz &&
      (hasOptionText || renderedOptions.length !== 4)
    ) {
      renderedOptions.sort(() => Math.random() - 0.5);
    }

    const displayLabels = useNumLabel
      ? ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10']
      : ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
    optionsContainer.innerHTML = renderedOptions
      .map((opt, idx) => {
        const dispLetter = displayLabels[idx] || String(idx + 1);
        const note = isMultiSelectQuiz
          ? '（可複選切換）'
          : shuffleOpts && hasOptionText
          ? '（選項順序已隨機洗牌）'
          : '';
        return `
          <button type="button" class="quiz-option-btn" data-orig-key="${escapeHtml(opt.origKey)}">
            <span style="width:30px; height:30px; border-radius:${isMultiSelectQuiz ? '6px' : '50%'}; background:#e2e8f0; display:inline-flex; align-items:center; justify-content:center; font-weight:700;">
              ${dispLetter}
            </span>
            <span style="flex:1;">${escapeHtml(opt.text)}</span>
            <span style="font-size:0.75rem; color:#94a3b8;">${note}</span>
          </button>
        `;
      })
      .join('');

    optionsContainer.querySelectorAll('.quiz-option-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (state.quizSubmitted) return;
        const key = btn.getAttribute('data-orig-key');
        if (isMultiSelectQuiz) {
          if (state.quizSelectedOptKeys.has(key)) {
            state.quizSelectedOptKeys.delete(key);
            btn.classList.remove('selected');
          } else {
            state.quizSelectedOptKeys.add(key);
            btn.classList.add('selected');
          }
          state.quizSelectedOptKey = Array.from(state.quizSelectedOptKeys)
            .sort()
            .join(', ');
        } else {
          state.quizSelectedOptKey = key;
          optionsContainer
            .querySelectorAll('.quiz-option-btn')
            .forEach((b) => b.classList.remove('selected'));
          btn.classList.add('selected');
        }
      });
    });
  }

  // Concept Double-Verification Setup
  const verifyEnabled = document.getElementById('chkConceptVerify').checked;
  const conceptBox = document.getElementById('quizConceptVerifyBox');
  const conceptListEl = document.getElementById('quizConceptOptionsList');

  if (verifyEnabled) {
    conceptBox.style.display = 'block';
    const conceptChoices = buildConceptVerificationChoices(m);
    state.currentQuestionRenderData = {
      conceptChoices,
      useTextInputMode,
      isMultiSelectQuiz
    };

    conceptListEl.innerHTML = conceptChoices
      .map(
        (c, idx) => `
        <button type="button" class="quiz-option-btn concept-choice-btn" data-cidx="${idx}" style="padding:0.65rem 0.9rem; font-size:0.88rem;">
          <span style="font-weight:700; color:#4f46e5;">觀念 ${idx + 1}：</span>
          <span>${escapeHtml(c.text)}</span>
        </button>
      `
      )
      .join('');

    conceptListEl.querySelectorAll('.concept-choice-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (state.quizSubmitted) return;
        state.quizSelectedConceptIdx = Number(btn.getAttribute('data-cidx'));
        conceptListEl
          .querySelectorAll('.concept-choice-btn')
          .forEach((b) => b.classList.remove('selected'));
        btn.classList.add('selected');
      });
    });
  } else {
    conceptBox.style.display = 'none';
    state.currentQuestionRenderData = {
      conceptChoices: [],
      useTextInputMode,
      isMultiSelectQuiz
    };
  }

  document.getElementById('quizFeedbackBox').style.display = 'none';
  document.getElementById('btnSubmitQuizAnswer').style.display = 'inline-flex';
  document.getElementById('btnNextQuizQuestion').style.display = 'none';
}

function buildConceptVerificationChoices(currentMistake) {
  const choices = [{ text: currentMistake.correctConcept, isTrue: true }];

  // Distractor 1: The student's own past misconception!
  choices.push({
    text: `【誤區迷思】${currentMistake.whyWrong}（這是當初導致選錯的想法！）`,
    isTrue: false
  });

  // Distractor 2: Subject-aware distractor or another mistake's concept
  const other = state.mistakes.find(
    (x) =>
      x.id !== currentMistake.id &&
      normalizeEnglishQType(x.questionType) === normalizeEnglishQType(currentMistake.questionType) &&
      x.correctConcept &&
      x.correctConcept !== currentMistake.correctConcept
  );
  if (other) {
    choices.push({ text: other.correctConcept, isTrue: false });
  } else {
    const englishDistractors = {
      choice: '只要看空格後面那一個字就能決定答案，不需要理解整句語意與搭配。',
      cloze: '克漏字只要看空格所在的那一句就好，前後句的轉折與語意不用管。',
      wordbank: '文意選填不用先判斷詞性，直接挑看起來最眼熟的單字填入即可。',
      structure: '篇章結構只要找到和前一句有重複單字的選項就一定是答案，不用檢查後一句。',
      reading: '閱讀測驗中，只要選項出現文章裡的原文單字，就一定是正確答案。',
      translation: '翻譯時照中文字的順序逐字翻成英文即可，不用考慮時態、主詞動詞一致與片語搭配。'
    };
    choices.push({
      text: englishDistractors[normalizeEnglishQType(currentMistake.questionType)],
      isTrue: false
    });
  }

  // Clean out the "【誤區迷思】" label before rendering so the student has to actually read it!
  const maskedChoices = choices.map((c) => ({
    text: c.text.replace(/^【誤區迷思】/, '').replace(/（這是當初導致選錯的想法！）$/, ''),
    isTrue: c.isTrue
  }));

  return maskedChoices.sort(() => Math.random() - 0.5);
}

function normalizeAnswerForComparison(str) {
  return String(str || '')
    .trim()
    .toLowerCase()
    .replace(/[.,!?;:'"()[\]{}，。！？；：「」『』（）\s]+/g, ' ')
    .trim();
}

window.overrideQuizSelfGradeCorrect = function () {
  const m = state.quizQueue[state.quizIndex];
  if (!m) return;
  const targetRecord = state.mistakes.find((item) => item.id === m.id);
  if (targetRecord) {
    targetRecord.correctCount = (targetRecord.correctCount || 0) + 1;
    targetRecord.masteryStatus =
      targetRecord.correctCount >= 2 ? '已精通' : '複習中';
  }
  state.quizCorrectCount++;
  state.quizWrongItems = state.quizWrongItems.filter((x) => x.id !== m.id);
  saveToLocalStorage();
  updateHeaderBadges();

  const fb = document.getElementById('quizFeedbackBox');
  if (fb) {
    fb.style.background = '#ecfdf5';
    fb.style.border = '2px solid #6ee7b7';
    fb.innerHTML = `
      <div style="font-size:1.05rem; font-weight:700; color:#047857;">
        ✅ 已將您的同義寫法／算式計為「答對」，並更新本題熟練度為「${escapeHtml(targetRecord ? targetRecord.masteryStatus : '複習中')}」！
      </div>
    `;
  }
};

function submitCurrentQuizAnswer() {
  if (state.quizSubmitted) return;
  const m = state.quizQueue[state.quizIndex];
  const verifyEnabled = document.getElementById('chkConceptVerify').checked;
  const useTextInputMode = Boolean(
    state.currentQuestionRenderData &&
      state.currentQuestionRenderData.useTextInputMode
  );
  const isMultiSelectQuiz = Boolean(
    state.currentQuestionRenderData &&
      state.currentQuestionRenderData.isMultiSelectQuiz
  );

  let userAnswerRaw = '';
  if (useTextInputMode) {
    const typedInput = document.getElementById('quizUserTypedAnswer');
    userAnswerRaw = typedInput ? typedInput.value.trim() : '';
    if (!userAnswerRaw) {
      alert('請先在輸入框中寫下您的作答內容！');
      return;
    }
  } else {
    if (!state.quizSelectedOptKey) {
      alert(
        isMultiSelectQuiz
          ? '請至少點選一個你認為正確的選項（多選題可複選）！'
          : '請先選擇一個題目答案選項！'
      );
      return;
    }
    userAnswerRaw = state.quizSelectedOptKey;
  }

  if (verifyEnabled && state.quizSelectedConceptIdx === null) {
    alert('請同時在下方「🧠 觀念雙重確認」選出這題背後的核心觀念！');
    return;
  }

  state.quizSubmitted = true;

  let isAnswerCorrect = false;
  let diffReport = null;

  if (useTextInputMode) {
    const typedInput = document.getElementById('quizUserTypedAnswer');
    if (typedInput) typedInput.disabled = true;

    const normUser = normalizeAnswerForComparison(userAnswerRaw);
    const normCorr = normalizeAnswerForComparison(m.correctAnswer);
    if (normUser === normCorr) {
      isAnswerCorrect = true;
    } else {
      // Check numeric equivalence (e.g. 2.0 == 2)
      const uNum = parseFloat(normUser);
      const cNum = parseFloat(normCorr);
      if (
        !Number.isNaN(uNum) &&
        !Number.isNaN(cNum) &&
        String(uNum) === normUser &&
        String(cNum) === normCorr &&
        Math.abs(uNum - cNum) < 1e-6
      ) {
        isAnswerCorrect = true;
      }
    }
    diffReport = computeWordDiffAnalysis(userAnswerRaw, m.correctAnswer);
  } else if (isMultiSelectQuiz) {
    const uTokens = parseMultiChoiceTokens(userAnswerRaw);
    const cTokens = parseMultiChoiceTokens(m.correctAnswer);
    isAnswerCorrect =
      uTokens.length > 0 &&
      uTokens.length === cTokens.length &&
      uTokens.every((k, i) => k === cTokens[i]);
  } else {
    const uTokens = parseMultiChoiceTokens(userAnswerRaw);
    const cTokens = parseMultiChoiceTokens(m.correctAnswer);
    if (uTokens.length === 1 && cTokens.length === 1) {
      isAnswerCorrect = uTokens[0] === cTokens[0];
    } else {
      isAnswerCorrect =
        userAnswerRaw.toUpperCase() === m.correctAnswer.toUpperCase();
    }
  }

  const isConceptCorrect = verifyEnabled
    ? state.currentQuestionRenderData.conceptChoices[state.quizSelectedConceptIdx]
        ?.isTrue === true
    : true;

  const fullPass = isAnswerCorrect && isConceptCorrect;

  // Highlight option buttons if multiple-choice / multi-select
  if (!useTextInputMode) {
    const uSet = new Set(parseMultiChoiceTokens(userAnswerRaw));
    const cSet = new Set(parseMultiChoiceTokens(m.correctAnswer));
    document
      .querySelectorAll('#quizOptionsContainer .quiz-option-btn')
      .forEach((btn) => {
        const k = btn.getAttribute('data-orig-key').toUpperCase();
        if (cSet.has(k) || k === m.correctAnswer.toUpperCase()) {
          btn.classList.add('correct-reveal');
        } else if (
          (uSet.has(k) || k === userAnswerRaw.toUpperCase()) &&
          !isAnswerCorrect
        ) {
          btn.classList.add('wrong-reveal');
        }
      });
  }

  if (verifyEnabled) {
    document
      .querySelectorAll('#quizConceptOptionsList .concept-choice-btn')
      .forEach((btn, idx) => {
        const cObj = state.currentQuestionRenderData.conceptChoices[idx];
        if (cObj.isTrue) {
          btn.classList.add('correct-reveal');
        } else if (idx === state.quizSelectedConceptIdx && !cObj.isTrue) {
          btn.classList.add('wrong-reveal');
        }
      });
  }

  // Update mistake record stats in state.mistakes
  const targetRecord = state.mistakes.find((item) => item.id === m.id);
  if (targetRecord) {
    targetRecord.practiceCount = (targetRecord.practiceCount || 0) + 1;
    if (fullPass) {
      targetRecord.correctCount = (targetRecord.correctCount || 0) + 1;
      if (targetRecord.correctCount >= 2) {
        targetRecord.masteryStatus = '已精通';
      } else {
        targetRecord.masteryStatus = '複習中';
      }
    } else {
      targetRecord.masteryStatus = '待加強';
    }
  }

  if (fullPass) {
    state.quizCorrectCount++;
  } else {
    state.quizWrongItems.push(m);
  }

  saveToLocalStorage();
  updateHeaderBadges();

  // Render Diagnostic Feedback Box
  const fb = document.getElementById('quizFeedbackBox');
  fb.style.display = 'block';
  if (fullPass) {
    fb.style.background = '#ecfdf5';
    fb.style.border = '2px solid #6ee7b7';
    fb.innerHTML = `
      <div style="font-size:1.05rem; font-weight:700; color:#047857; margin-bottom:0.4rem;">
        🎉 完全正確！你不僅答對題目，還掌握了正確的核心觀念！
      </div>
      <div style="font-size:0.9rem; color:#065f46;">
        <strong>✅ 標準答案：</strong>${escapeHtml(m.correctAnswer)}<br/>
        <strong>🎯 核心觀念複習：</strong>${escapeHtml(m.correctConcept)}<br/>
        <strong>💡 避坑錦囊：</strong>${escapeHtml(m.keyTakeaway)}<br/>
        <span style="font-size:0.82rem; opacity:0.85;">（本題熟練狀態已更新為：<b>${escapeHtml(targetRecord ? targetRecord.masteryStatus : '複習中')}</b>）</span>
      </div>
    `;
  } else {
    fb.style.background = '#fef2f2';
    fb.style.border = '2px solid #fca5a5';
    let failReason = '';
    let diffHtmlBlock = '';
    if (!isAnswerCorrect) {
      if (useTextInputMode) {
        failReason = `✍️ 作答與標準解答有差異，請對照下方逐字比對（若您的寫法為同義詞／等價算式，可點擊下方按鈕計為答對）：`;
        diffHtmlBlock = `
          <div style="background:#ffffff; border:1px solid #fecaca; border-radius:8px; padding:0.65rem; margin:0.4rem 0;">
            ${diffReport ? diffReport.diffHtml : ''}
            <div style="margin-top:0.5rem;">
              <button type="button" class="btn btn-secondary btn-sm" onclick="overrideQuizSelfGradeCorrect()" style="border-color:#10b981; color:#047857; font-weight:700;">
                ✅ 我的寫法意思／同義詞／算式正確（計為答對）
              </button>
            </div>
          </div>
        `;
      } else if (userAnswerRaw.toUpperCase() === m.wrongAnswer.toUpperCase()) {
        failReason = `😱 小心！你這次又選了跟上次一模一樣的錯誤選項 (${escapeHtml(m.wrongAnswer)})！`;
      } else {
        failReason = `❌ 答案選錯囉（正確答案應為 ${escapeHtml(m.correctAnswer)}）！`;
      }
    } else {
      failReason = `⚠️ 雖然題目答對了，但「核心觀念確認」選錯了，小心不要靠死背答案喔！`;
    }

    fb.innerHTML = `
      <div style="font-size:1.05rem; font-weight:700; color:#b91c1c; margin-bottom:0.45rem;">
        ${failReason}
      </div>
      ${diffHtmlBlock}
      <div style="font-size:0.9rem; color:#7f1d1d; display:flex; flex-direction:column; gap:0.35rem;">
        <div><strong>🧐 你當初紀錄的「為何寫錯」：</strong>${escapeHtml(m.whyWrong)}</div>
        <div><strong>✅ 正確觀念與寫法解析：</strong>${escapeHtml(m.correctConcept)}</div>
        <div><strong>💡 避坑錦囊：</strong>${escapeHtml(m.keyTakeaway)}</div>
      </div>
    `;
  }

  document.getElementById('btnSubmitQuizAnswer').style.display = 'none';
  const nextBtn = document.getElementById('btnNextQuizQuestion');
  nextBtn.style.display = 'inline-flex';
  nextBtn.textContent =
    state.quizIndex + 1 < state.quizQueue.length
      ? '下一題 ▶'
      : '查看練習總結報告 🏆';
}

function finishPracticeQuiz() {
  document.getElementById('quizActivePanel').style.display = 'none';
  const resPanel = document.getElementById('quizResultPanel');
  resPanel.style.display = 'block';

  const total = state.quizQueue.length;
  const correct = state.quizCorrectCount;
  const rate = Math.round((correct / total) * 100);

  document.getElementById('quizResultEmoji').textContent =
    rate === 100 ? '🏆' : rate >= 70 ? '🎉' : '💪';
  document.getElementById(
    'quizResultHeadline'
  ).textContent = `本次強化練習得分：${rate}%（答對 ${correct} / ${total} 題）`;
  document.getElementById(
    'quizResultSubtext'
  ).textContent = `所有題目的最新練習次數與熟練狀態（待加強／複習中／已精通）皆已自動同步至您的試算表中！`;

  const retryBtn = document.getElementById('btnRetryWrongOnly');
  if (state.quizWrongItems.length > 0) {
    retryBtn.style.display = 'inline-flex';
    retryBtn.textContent = `🔁 針對剛剛未過關的 ${state.quizWrongItems.length} 題再次特訓`;
  } else {
    retryBtn.style.display = 'none';
  }
}

// ============================================================================
// STEP 4: Spreadsheet Table View & Excel (.xlsx) / Google Sheets (.csv) Export
// ============================================================================
function initSheetTab() {
  document
    .getElementById('btnExportXlsx')
    .addEventListener('click', exportToExcelXlsx);
  document
    .getElementById('btnExportCsvForGoogle')
    .addEventListener('click', exportToCsvForGoogleSheets);
  document
    .getElementById('btnCopyTsvForGoogle')
    .addEventListener('click', copyTsvToClipboardForGoogleSheets);
}

function renderSheetTableAndStats() {
  document.getElementById('sheetInfoName').textContent = state.sheetName;
  document.getElementById('sheetInfoMode').textContent =
    state.sheetMode === 'existing'
      ? `📂 模式：延續既有試算表（原載入 ${state.originalLoadedCount} 筆）`
      : '✨ 模式：開啟的新試算表';

  const total = state.mistakes.length;
  const mastered = state.mistakes.filter(
    (m) => m.masteryStatus === '已精通'
  ).length;
  const newlyAdded = state.mistakes.filter((m) => !m.isHistorical).length;

  document.getElementById(
    'sheetInfoStats'
  ).textContent = `${total} 題（🟢 已精通 ${mastered} 題）`;
  document.getElementById(
    'sheetInfoNewCount'
  ).textContent = `本次新增錯題：${newlyAdded} 題`;

  // Calculate top error reason
  const tagCounts = {};
  state.mistakes.forEach((m) => {
    (m.errorTags || []).forEach((t) => {
      tagCounts[t] = (tagCounts[t] || 0) + 1;
    });
  });
  const sortedTags = Object.entries(tagCounts).sort((a, b) => b[1] - a[1]);
  document.getElementById('sheetInfoTopError').textContent =
    sortedTags.length > 0
      ? `${sortedTags[0][0]}（出現 ${sortedTags[0][1]} 次）`
      : '尚無資料';

  document.getElementById('tableRowCount').textContent = String(total);

  const tbody = document.getElementById('sheetTableBody');
  if (total === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:2rem; color:#64748b;">目前試算表內尚無錯題紀錄。</td></tr>`;
    return;
  }

  tbody.innerHTML = state.mistakes
    .map(
      (m) => `
      <tr>
        <td>
          <div style="font-weight:700; color:#312e81;">第 ${escapeHtml(m.questionNum)} 題</div>
          <span style="font-size:0.78rem; background:#e0e7ff; color:#4f46e5; padding:0.1rem 0.45rem; border-radius:999px; font-weight:600;">${escapeHtml(m.subject)}</span>
          <div style="font-size:0.72rem; color:#64748b; margin-top:0.15rem;">${escapeHtml(getQuestionTypeLabel(m.questionType))}</div>
        </td>
        <td>
          <div style="font-weight:600;">${escapeHtml(m.unit)}</div>
          <div style="font-size:0.75rem; color:#64748b;">${escapeHtml(m.createdAt)}</div>
        </td>
        <td style="max-width:280px;">
          <div style="font-weight:600; margin-bottom:0.25rem;">${escapeHtml(m.questionText)}</div>
          ${
            ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].some((k) => m[`opt${k}`])
              ? `<div style="font-size:0.76rem; color:#64748b;">${['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
                  .filter((k) => m[`opt${k}`])
                  .map((k) => `(${k}) ${escapeHtml(m[`opt${k}`])}`)
                  .join(' ')}</div>`
              : ''
          }
        </td>
        <td style="white-space:nowrap;">
          <span style="color:#b91c1c; font-weight:700;">❌ 誤選：${escapeHtml(m.wrongAnswer)}</span><br/>
          <span style="color:#047857; font-weight:700;">✅ 正解：${escapeHtml(m.correctAnswer)}</span>
        </td>
        <td>${(m.errorTags || []).map((t) => `<span style="display:inline-block; background:#f1f5f9; padding:0.15rem 0.45rem; border-radius:6px; font-size:0.76rem; margin:0.1rem;">${escapeHtml(t)}</span>`).join('')}</td>
        <td style="max-width:260px;">
          <div style="color:#991b1b; font-size:0.8rem;"><strong>錯因：</strong>${escapeHtml(m.whyWrong)}</div>
          <div style="color:#065f46; font-size:0.8rem; margin-top:0.2rem;"><strong>正解：</strong>${escapeHtml(m.correctConcept)}</div>
        </td>
        <td style="max-width:200px; font-weight:600; color:#92400e;">💡 ${escapeHtml(m.keyTakeaway)}</td>
        <td style="white-space:nowrap;">
          <div style="font-weight:700;">${escapeHtml(m.masteryStatus)}</div>
          <div style="font-size:0.76rem; color:#64748b;">練習 ${m.practiceCount} 次 / 對 ${m.correctCount} 次</div>
        </td>
        <td style="white-space:nowrap;">
          <button class="btn btn-secondary btn-sm" onclick="downloadMistakePngById('${m.id}')">🖼️ PNG</button>
          <button class="btn btn-secondary btn-sm" onclick="editMistakeById('${m.id}')">✏️</button>
        </td>
      </tr>
    `
    )
    .join('');
}

function formatMistakesForSpreadsheetExport(includeBase64Crop = true) {
  return state.mistakes.map((m, idx) => {
    const row = {
      錯題ID: m.id,
      建立日期: m.createdAt,
      科目: m.subject,
      試卷單元: m.unit,
      題號: m.questionNum,
      題型: getQuestionTypeLabel(m.questionType),
      題目內容: m.questionText,
      選項A: m.optA || '',
      選項B: m.optB || '',
      選項C: m.optC || '',
      選項D: m.optD || '',
      選項E: m.optE || '',
      選項F: m.optF || '',
      選項G: m.optG || '',
      選項H: m.optH || '',
      選項I: m.optI || '',
      選項J: m.optJ || '',
      我的錯誤答案: m.wrongAnswer,
      正確答案: m.correctAnswer,
      錯誤類型標籤: (m.errorTags || []).join('、'),
      為何選錯診斷: m.whyWrong,
      正確觀念解析: m.correctConcept,
      避坑錦囊: m.keyTakeaway,
      練習次數: m.practiceCount || 0,
      答對次數: m.correctCount || 0,
      熟練狀態: m.masteryStatus || '待加強'
    };
    if (includeBase64Crop) {
      // Excel cell limit is 32,767 chars; only store image data URL if it fits within cell limit
      row['題目截圖資料'] =
        m.cropImageDataUrl && m.cropImageDataUrl.length < 32000
          ? m.cropImageDataUrl
          : '';
    }
    return row;
  });
}

function exportToExcelXlsx() {
  if (state.mistakes.length === 0) {
    alert('目前試算表還沒有錯題資料！請先新增錯題或載入範例資料。');
    return;
  }

  const rows = formatMistakesForSpreadsheetExport(true);
  const worksheet = XLSX.utils.json_to_sheet(rows);

  // Set friendly column widths
  worksheet['!cols'] = [
    { wch: 16 }, // 錯題ID
    { wch: 12 }, // 建立日期
    { wch: 10 }, // 科目
    { wch: 20 }, // 試卷單元
    { wch: 8 },  // 題號
    { wch: 18 }, // 題型
    { wch: 42 }, // 題目內容
    { wch: 20 }, // 選項A
    { wch: 20 }, // 選項B
    { wch: 20 }, // 選項C
    { wch: 20 }, // 選項D
    { wch: 20 }, // 選項E
    { wch: 18 }, // 選項F
    { wch: 18 }, // 選項G
    { wch: 18 }, // 選項H
    { wch: 18 }, // 選項I
    { wch: 18 }, // 選項J
    { wch: 14 }, // 我的錯誤答案
    { wch: 14 }, // 正確答案
    { wch: 22 }, // 錯誤類型標籤
    { wch: 36 }, // 為何選錯診斷
    { wch: 36 }, // 正確觀念解析
    { wch: 32 }, // 避坑錦囊
    { wch: 10 }, // 練習次數
    { wch: 10 }, // 答對次數
    { wch: 12 }, // 熟練狀態
    { wch: 18 }  // 題目截圖資料
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '錯題紀錄總表');

  const safeName = (state.sheetName || '我的錯題進化本').replace(
    /[\\/:*?"<>|]/g,
    '_'
  );
  XLSX.writeFile(workbook, `${safeName}.xlsx`);
  showToast(
    `📥 已成功匯出 Excel 檔案「${safeName}.xlsx」！下次可選擇「繼續使用已有的試算表」載入此檔案。`
  );
}

function exportToCsvForGoogleSheets() {
  if (state.mistakes.length === 0) {
    alert('目前試算表還沒有錯題資料！');
    return;
  }

  const rows = formatMistakesForSpreadsheetExport(false);
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const csvContent = '\uFEFF' + XLSX.utils.sheet_to_csv(worksheet);
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const safeName = (state.sheetName || '我的錯題進化本').replace(
    /[\\/:*?"<>|]/g,
    '_'
  );
  triggerDownloadUrl(url, `${safeName}_Google試算表相容.csv`);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  showToast(`📄 已下載 Google 試算表專用 CSV 檔！可直接匯入 Google 雲端硬碟。`);
}

async function copyTsvToClipboardForGoogleSheets() {
  if (state.mistakes.length === 0) {
    alert('目前試算表還沒有錯題資料！');
    return;
  }
  const rows = formatMistakesForSpreadsheetExport(false);
  const headers = Object.keys(rows[0]);
  const tsvLines = [
    headers.join('\t'),
    ...rows.map((r) =>
      headers
        .map((h) => String(r[h] ?? '').replace(/[\t\r\n]+/g, ' '))
        .join('\t')
    )
  ];
  await navigator.clipboard.writeText(tsvLines.join('\n'));
  showToast(
    '📋 已複製全部錯題列！您現在可以直接在已開啟的雲端 Google 試算表按 Ctrl+V 貼上！'
  );
}

// ============================================================================
// Built-in Demo Data Loader (One-Click Full Experience)
// ============================================================================
function loadBuiltInDemoEverything() {
  generateInteractiveSampleExam();

  const today = new Date().toLocaleDateString('zh-TW');
  const base = {
    createdAt: today,
    subject: '英文',
    unit: '高中英文段考精選',
    optA: '', optB: '', optC: '', optD: '', optE: '',
    optF: '', optG: '', optH: '', optI: '', optJ: '',
    cropImageDataUrl: '',
    isHistorical: true
  };

  state.sheetName = '高中英文錯題本_範例';
  state.sheetMode = 'existing';
  state.originalLoadedCount = 6;
  state.mistakes = [
    {
      ...base,
      id: 'DEMO_EN_1',
      questionType: 'choice',
      questionNum: '1',
      questionText:
        'Eric has been a(n) ______ employee at the company for over ten years, always dependable and committed to his work.',
      optA: 'faithful',
      optB: 'artificial',
      optC: 'reluctant',
      optD: 'potential',
      wrongAnswer: 'D',
      correctAnswer: 'A',
      errorTags: ['單字與片語搭配'],
      whyWrong:
        '你選了 (D) potential（潛在的、可能的），它指「未來可能成為…但現在還不是」，不能形容已經任職十年的員工。',
      correctConcept:
        '線索字 for over ten years（長期）＋ dependable、committed（可靠、盡心）→ 空格要填正面的長期特質：(A) faithful 忠實的。',
      keyTakeaway: '必背對比：potential＝潛在的（還沒發生） vs. faithful＝忠實的（長期可靠，同義 loyal）！',
      practiceCount: 1,
      correctCount: 0,
      masteryStatus: '待加強'
    },
    {
      ...base,
      id: 'DEMO_EN_2',
      questionType: 'cloze',
      questionNum: '19',
      questionText:
        'The young student walked along the shore and talked to the people, __19__ were cleaning the beach.',
      optA: 'that',
      optB: 'who',
      optC: 'whom',
      optD: 'which',
      wrongAnswer: 'A',
      correctAnswer: 'B',
      errorTags: ['文法與詞性判斷'],
      whyWrong: '你選了 (A) that，但空格前有逗號（非限定用法），that 不能用在逗號後面。',
      correctConcept:
        '先行詞 people 是人，空格後緊接動詞 were（缺主詞），且前有逗號 → 只能用主格 who。',
      keyTakeaway: '關代三步驟：①先行詞人/物 ②有逗號禁用 that ③後面缺主詞選 who、缺受詞選 whom！',
      practiceCount: 2,
      correctCount: 1,
      masteryStatus: '複習中'
    },
    {
      ...base,
      id: 'DEMO_EN_3',
      questionType: 'wordbank',
      questionNum: '23',
      questionText:
        'Standing on the top of the mountain, I am __23__ by the sheer beauty of nature.',
      optA: 'dazzled',
      optB: 'spot',
      optC: 'reluctant',
      optD: 'massive',
      optE: 'harmful',
      optF: 'rapidly',
      optG: 'preserve',
      optH: 'recognize',
      optI: 'tremendous',
      optJ: 'consequence',
      wrongAnswer: 'G',
      correctAnswer: 'A',
      errorTags: ['文法與詞性判斷'],
      whyWrong:
        '你選了 (G) preserve（原形動詞），但空格前是 be 動詞 am、後面接 by，需要過去分詞（被動／情緒）。',
      correctConcept:
        '① 判斷詞性：am + ___ + by → 過去分詞 p.p.。② 候選字中只有 (A) dazzled 是 p.p.，意思是「目眩神迷的」，符合「被大自然的美震懾」。',
      keyTakeaway: '文意選填先分詞性！be 動詞 + ___ + by → 填 p.p.（dazzled）。',
      practiceCount: 1,
      correctCount: 1,
      masteryStatus: '複習中'
    },
    {
      ...base,
      id: 'DEMO_EN_4',
      questionType: 'structure',
      questionNum: '44',
      questionText:
        'Plastic bags are cheap, light, and easy to carry, so stores give them out everywhere. __44__ Every year, millions of them end up in the ocean, where they harm sea turtles and other marine animals.',
      optA: 'However, such convenience comes at a heavy environmental cost.',
      optB: 'For example, many stores give them out for free.',
      optC: 'As a result, plastic was invented in the early 1900s.',
      optD: 'They are also used to make colorful toys.',
      optE: 'In addition, paper bags are much heavier.',
      wrongAnswer: 'B',
      correctAnswer: 'A',
      errorTags: ['篇章結構與轉折'],
      whyWrong:
        '你選了 (B) For example…，它只是重複前一句「商店到處發送」的內容，而且後一句在講「海洋污染」，(B) 接不上後一句。',
      correctConcept:
        '前一句講塑膠袋「方便」，後一句講「污染海洋」→ 中間需要轉折 However；such convenience 回指前一句的「方便」，也帶出後一句的環境代價。',
      keyTakeaway: '篇章結構三看：①開頭轉折詞 ②指代詞（such / this）回指誰 ③前後句都要接得上！',
      practiceCount: 0,
      correctCount: 0,
      masteryStatus: '待加強'
    },
    {
      ...base,
      id: 'DEMO_EN_5',
      questionType: 'reading',
      questionNum: '48',
      questionText:
        '(Passage about how convenience stores in Taiwan offer services such as paying bills, picking up packages, and buying train tickets.) What is the main idea of the passage?',
      optA: 'Convenience stores sell many kinds of snacks.',
      optB: 'Convenience stores in Taiwan provide a wide range of daily services.',
      optC: 'People can buy train tickets at convenience stores.',
      optD: 'Convenience stores are open 24 hours a day.',
      wrongAnswer: 'C',
      correctAnswer: 'B',
      errorTags: ['閱讀定位與推論'],
      whyWrong: '你選了 (C)，它只是文章中的一個細節（以偏概全），不是全文主旨。',
      correctConcept:
        '主旨題看首段與各段第一句：全文都在講便利商店提供「各種生活服務」，(B) 才能涵蓋全部內容。',
      keyTakeaway: '主旨題避坑：只出現在某一段的細節選項直接刪！',
      practiceCount: 2,
      correctCount: 2,
      masteryStatus: '已精通'
    },
    {
      ...base,
      id: 'DEMO_EN_6',
      questionType: 'translation',
      questionNum: '二-1',
      questionText: '這家公司花了三年才開發出這款新手機。',
      wrongAnswer: 'The company spent three years develop the new phone.',
      correctAnswer: 'It took the company three years to develop the new phone.',
      errorTags: ['翻譯句型與拼字'],
      whyWrong:
        '「spend + 時間」後面要接 V-ing（spent three years developing），你寫成原形 develop；而且本題更常用 It takes 句型。',
      correctConcept:
        '花費時間句型：It took + 人/組織 + 時間 + to V → It took the company three years to develop the new phone.',
      keyTakeaway: '必背：It takes/took + 人 + 時間 + to V；人當主詞用 spend + 時間 + V-ing！',
      practiceCount: 1,
      correctCount: 0,
      masteryStatus: '待加強'
    }
  ];

  document.getElementById('sheetInitModal').classList.add('hidden');
  saveToLocalStorage();
  updateAllViews();
  showToast(
    '✨ 已載入英文範例試卷與 6 筆英文錯題（單字選擇、克漏字、文意選填、篇章結構、閱讀測驗、翻譯）！'
  );
}

// ============================================================================
// UI Helpers & Synchronizers
// ============================================================================
function updateAllViews() {
  updateHeaderBadges();
  updateSubjectDropdownFilters();
  renderMistakeCards();
  updateQuizSetupSummary();
  renderSheetTableAndStats();
}

function updateHeaderBadges() {
  const syncType = state.excelFileHandle
    ? '📗 已綁定本機 Excel'
    : state.linkedGoogleSheetUrl
    ? '🟢 已連動 Google 試算表'
    : state.sheetMode === 'existing'
    ? '延續舊表'
    : '新試算表';
  document.getElementById(
    'currentSheetBadgeText'
  ).textContent = `📊 ${state.sheetName}（${syncType}・共 ${state.mistakes.length} 題）`;
  document.getElementById('navMistakeCount').textContent = String(
    state.mistakes.length
  );
  const btnLinked = document.getElementById('btnOpenLinkedGoogleSheet');
  if (btnLinked) {
    btnLinked.style.display = state.linkedGoogleSheetUrl ? 'inline-flex' : 'none';
  }
}

function updateSubjectDropdownFilters() {
  const counts = {};
  state.mistakes.forEach((m) => {
    const t = normalizeEnglishQType(m.questionType);
    counts[t] = (counts[t] || 0) + 1;
  });
  const selects = [
    document.getElementById('filterCardSubject'),
    document.getElementById('quizFilterSubject')
  ];
  selects.forEach((sel) => {
    if (!sel) return;
    const prevVal = sel.value || 'ALL';
    sel.innerHTML =
      `<option value="ALL">全部英文題型（${state.mistakes.length}）</option>` +
      ENGLISH_QTYPES.map(
        (t) => `<option value="${t}">${escapeHtml(getQuestionTypeLabel(t))}（${counts[t] || 0}）</option>`
      ).join('');
    sel.value = prevVal === 'ALL' || ENGLISH_QTYPES.includes(prevVal) ? prevVal : 'ALL';
  });
}

function triggerDownloadUrl(url, filename) {
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function showToast(msg) {
  const container = document.getElementById('toastContainer');
  const t = document.createElement('div');
  t.className = 'toast';
  t.textContent = msg;
  container.appendChild(t);
  setTimeout(() => {
    t.style.opacity = '0';
    setTimeout(() => t.remove(), 250);
  }, 4000);
}
