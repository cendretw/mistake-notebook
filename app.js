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

// Diagnostic Guided Prompts & Quick Phrases Dictionary
const DIAGNOSIS_GUIDES = {
  '審題失誤': {
    question: (w, c) =>
      `🔍 【審題失誤診斷】你選了 (${w || '?'})，但正解是 (${c || '?'})。回頭看試卷題幹，你是漏看了「選出錯誤/正確」、「單位換算」，還是忽略了某個限定條件？`,
    phrases: [
      '題目問「何者錯誤」，我太急看成「何者正確」而誤選',
      '漏看題目關鍵限定詞（如：光滑平面、定溫下、至少、不包含）',
      '沒有注意到題目給的單位（如 cm 與 m、分鐘與秒）未換算',
      '只看前半句敘述就作答，沒看完後半句的轉折條件'
    ]
  },
  '觀念混淆': {
    question: (w, c) =>
      `🧠 【觀念混淆診斷】你把 (${w || '?'}) 跟正確答案 (${c || '?'}) 的概念搞混了！請問 (${w || '?'}) 其實是在什麼條件下才成立？它跟 (${c || '?'}) 的最大差異在哪裡？`,
    phrases: [
      '把兩個相似的定義/名詞記顛倒了（正比與反比、串联與並聯等）',
      '誤以為 A 變大時 B 一定跟著變大，忽略了公式中的負號或分母關係',
      '記錯公式適用條件，把甲單元的公式套用到乙單元',
      '因果關係倒置，把「結果」當成「原因」'
    ]
  },
  '掉入陷阱': {
    question: (w, c) =>
      `🪤 【破解出題陷阱】選項 (${w || '?'}) 是出題老師設計的經典誘答！它是用什麼「看起來很合理」的敘述騙過你的？`,
    phrases: [
      '選項敘述前半段是對的，但後半段偷換概念，我沒檢查完就選了',
      '直覺以為是這樣，但實際上這是一個常見的科學/數學迷思（特例不成立）',
      '被題目多給的無關數據干擾，掉入出題老師設計的誘答選項',
      '選項文字用「絕對、必定、所有」等極端詞，忽略了例外狀況'
    ]
  },
  '計算粗心': {
    question: (w, c) =>
      `🔢 【計算流程檢查】既然觀念會，為什麼算出 (${w || '?'}) 而不是 (${c || '?'})？是移項變號、分數通分、指數運算，還是抄錯數字？`,
    phrases: [
      '方程式移項時忘記變號（正負號寫反）',
      '算式列對了，但在心算或直式加減乘除時算錯數字',
      '求出中間未知數 X 後就直接選答案，忘了題目最後問的是 2X+1',
      '計算紙寫太亂，把上一行的數字抄錯到下一行'
    ]
  },
  '圖表判讀錯誤': {
    question: (w, c) =>
      `📊 【圖表判讀診斷】這題有附圖或表格！你選 (${w || '?'}) 是因為沒看清楚橫軸/縱軸代表的物理量，還是看錯斜率/面積的意義？`,
    phrases: [
      '沒看清楚座標軸的縱軸(Y)與橫軸(X)代表的變數名稱',
      '把 v-t 圖當成 x-t 圖來看，誤判了斜率或曲線下面積的意義',
      '圖表起點不是從 0 開始（有截距），我直接拿座標相除而算錯',
      '看錯圖例（實線與虛線代表的對象看反了）'
    ]
  },
  '完全不會': {
    question: (w, c) =>
      `❓ 【核心觀念補救】這題看到時不知如何下手而猜了 (${w || '?'})。對照解答 (${c || '?'}) 後，這題其實是在考課本哪一個核心定律或公式？`,
    phrases: [
      '完全忘記這個單元的核心公式，需要重新背誦並推導一次',
      '看不懂題目的文字敘述如何轉化為數學算式/物理模型',
      '這是跨單元綜合題，沒想到要把兩個公式聯立求解',
      '單字/專有名詞看不懂，導致無法理解題目與選項意思'
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
      if (rawTypeLabel.includes('翻譯') || rawTypeLabel.includes('非選')) {
        inferredType = 'translation';
      } else if (rawTypeLabel.includes('填充') || rawTypeLabel.includes('計算') || rawTypeLabel.includes('字彙')) {
        inferredType = 'fill';
      } else if (rawTypeLabel.includes('閱讀') || rawTypeLabel.includes('圖表')) {
        inferredType = 'reading';
      } else if (
        !row['選項A'] &&
        !row['選項B'] &&
        (correctAns.length > 2 || wrongAns.length > 2)
      ) {
        inferredType = /\s/.test(correctAns) && correctAns.length > 15 ? 'translation' : 'fill';
      }

      return {
        id: row['錯題ID'] || `M_${Date.now()}_${idx}`,
        createdAt: row['建立日期'] || new Date().toLocaleDateString('zh-TW'),
        subject: row['科目'] || '綜合科目',
        unit: row['試卷單元'] || row['單元'] || '既有試算表匯入',
        questionNum: String(row['題號'] || idx + 1),
        questionType: inferredType,
        questionText: qText,
        optA: String(row['選項A'] || ''),
        optB: String(row['選項B'] || ''),
        optC: String(row['選項C'] || ''),
        optD: String(row['選項D'] || ''),
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
    state.cropTarget = 'question';
    setCropMode(!state.cropMode, 'question');
  });

  if (btnCropPassage) {
    btnCropPassage.addEventListener('click', () => {
      state.cropTarget = 'passage';
      setCropMode(true, 'passage');
    });
  }

  btnSample1.addEventListener('click', () => generateInteractiveSampleExam());
  btnSample2.addEventListener('click', () => generateInteractiveSampleExam());

  initCropCanvasMouseEvents();
}

function setCropMode(enabled, modeType = 'question') {
  state.cropMode = enabled;
  state.cropTarget = modeType;
  const btn = document.getElementById('btnToggleCropMode');
  const overlay = document.getElementById('cropOverlayCanvas');
  const banner = document.getElementById('cropHintBanner');

  btn.classList.toggle('active', enabled && modeType === 'question');
  btn.textContent =
    enabled && modeType === 'question'
      ? '🛑 結束框選（請在下方試卷拖曳框選該題）'
      : '✂️ 框選題目與選項';
  overlay.classList.toggle('active', enabled);
  if (banner) {
    banner.style.display = enabled ? 'block' : 'none';
    banner.textContent =
      modeType === 'passage'
        ? '📖 【加框克漏字文章模式】請在下方試卷拖曳框選「克漏字文章段落（例如含 __19__ 的那段文字）」，系統會自動將文章截圖拼接到選項上方並帶入上下文！'
        : '💡 請在下方試卷上「拖曳框選你要的那一題（例如第 19 題）」，系統會精準抓取該題選項並過濾上下題干擾！';
  }
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
    const pdfDoc = await window.pdfjsLib.getDocument({ data: arrayBuf }).promise;
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
  examDoc.fileName = '115學年度_自然與數學段考模擬試卷.pdf';

  document.getElementById('examUploadBox').classList.add('loaded');
  document.getElementById('examFileName').textContent = examDoc.fileName;

  // Populate standard answers for the sample exam
  state.answerKeyMap = { '1': 'C', '2': 'B', '3': 'D', '4': 'A' };
  document.getElementById('inputBulkAnswers').value = '1:C, 2:B, 3:D, 4:A';
  renderAnswerKeyChips();

  switchDocSlot('exam');
  setCropMode(true);
  showToast(
    '🧪 已載入內建測試試卷與標準答案表（1:C, 2:B, 3:D, 4:A）！現在可直接在試卷上拖曳框選任一題！'
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
  ctx.fillText('115 學年度第一學期 自然與數學科 段考精選試卷', baseW / 2, 72);

  ctx.font = '14px "Noto Sans TC", sans-serif';
  ctx.fillStyle = '#475569';
  ctx.fillText(
    '請在畫面上使用「✂️ 框選題目截圖＋擷取文字」拖曳框出任一題，即可自動擷取題目與選項！',
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
      subject: '自然科學',
      unit: '直線運動與加速度',
      stem: '1. 一輛汽車在筆直公路上向東行駛，司機看到紅燈後踩下煞車做等減速度運動直到停止。請問在煞車期間，關於汽車的「速度方向」與「加速度方向」，下列敘述何者錯誤？',
      optA: '汽車的速度方向依然向東',
      optB: '汽車的位移方向向東',
      optC: '汽車的速度越來越慢，因此加速度方向也向東',
      optD: '汽車的加速度方向向西（與速度方向相反）'
    },
    {
      yTop: 320,
      yBottom: 495,
      qNum: '2',
      subject: '數學',
      unit: '一元二次方程式與乘法公式',
      stem: '2. 計算並化簡代數式 (2x - 3)² - (x - 5)(x + 5) 的結果，下列哪一個選項是正確的？',
      optA: '3x² - 12x - 16',
      optB: '3x² - 12x + 34',
      optC: '3x² - 6x + 34',
      optD: '5x² - 12x - 16'
    },
    {
      yTop: 510,
      yBottom: 685,
      qNum: '3',
      subject: '生物',
      unit: '細胞的構造與功能',
      stem: '3. 小明使用複式顯微鏡觀察水蘊草葉片細胞，若他想找尋被稱為「細胞能量發電廠」（負責進行呼吸作用產生能量）的胞器，應該觀察下列哪一個構造？',
      optA: '葉綠體',
      optB: '液泡',
      optC: '核糖體',
      optD: '粒線體'
    },
    {
      yTop: 700,
      yBottom: 870,
      qNum: '4',
      subject: '英文',
      unit: '假設語氣與完成式',
      stem: '4. Choose the correct answer: "Look at the dark clouds! Take an umbrella with you ______ it rains heavily this afternoon."',
      optA: 'in case',
      optB: 'even though',
      optC: 'as soon as',
      optD: 'unless'
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
    ctx.font = 'bold 15px "Noto Sans TC", sans-serif';
    wrapTextOnCanvas(ctx, q.stem, 62, q.yTop + 28, baseW - 124, 24);

    // Draw options
    ctx.font = '14px "Noto Sans TC", sans-serif';
    ctx.fillStyle = '#334155';
    ctx.fillText(`(A) ${q.optA}`, 68, q.yTop + 98);
    ctx.fillText(`(B) ${q.optB}`, 390, q.yTop + 98);
    ctx.fillText(`(C) ${q.optC}`, 68, q.yTop + 130);
    ctx.fillText(`(D) ${q.optD}`, 390, q.yTop + 130);
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

  overlay.addEventListener('mouseup', async (evt) => {
    if (!state.cropMode || !state.isDraggingCrop) return;
    state.isDraggingCrop = false;
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

    if (state.cropTarget === 'passage' && state.currentCropDataUrl) {
      // Stitch Cloze passage image ABOVE existing question options image
      const mergedDataUrl = await stitchTwoImagesVertically(
        croppedDataUrl,
        state.currentCropDataUrl
      );
      setCroppedImagePreview(mergedDataUrl);
      await extractTextFromCropRegion(rx, ry, rw, rh, 'passage');
    } else {
      setCroppedImagePreview(croppedDataUrl);
      await extractTextFromCropRegion(rx, ry, rw, rh, state.cropTarget || 'question');
    }
    setCropMode(false, 'question');
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

  if (docState.type === 'sample') {
    const unscaledYCenter = (ry + rh / 2) / docState.scale;
    const matched = docState.sampleItems.find(
      (item) => unscaledYCenter >= item.yTop && unscaledYCenter <= item.yBottom
    );
    if (matched) {
      document.getElementById('fieldSubject').value = matched.subject;
      document.getElementById('fieldUnit').value = matched.unit;
      document.getElementById('fieldQuestionNum').value = matched.qNum;
      document.getElementById('fieldQuestionText').value = matched.stem;
      document.getElementById('fieldOptA').value = matched.optA;
      document.getElementById('fieldOptB').value = matched.optB;
      document.getElementById('fieldOptC').value = matched.optC;
      document.getElementById('fieldOptD').value = matched.optD;
      autoFillAnswerFromKeyMap(matched.qNum);
      runAutoMistakeDiagnosis();
      showToast(
        `✂️ 已自動擷取第 ${matched.qNum} 題截圖、題幹與 (A)~(D) 選項！`
      );
      return;
    }
  }

  if (docState.type === 'pdf' && docState.pdfDoc) {
    try {
      const page = await docState.pdfDoc.getPage(docState.pageNum);
      const viewport = page.getViewport({ scale: docState.scale });
      const textContent = await page.getTextContent();

      // Auto-infer subject from PDF filename if obvious
      if (docState.fileName) {
        if (/英文|English/i.test(docState.fileName)) {
          document.getElementById('fieldSubject').value = '英文';
        } else if (/數學|Math/i.test(docState.fileName)) {
          document.getElementById('fieldSubject').value = '數學';
        } else if (/自然|物理|化學|生物/i.test(docState.fileName)) {
          document.getElementById('fieldSubject').value = '自然科學';
        }
        document.getElementById('fieldUnit').value = docState.fileName.replace(/\.pdf$/i, '');
      }

      // Strict center-based bounding box filter so adjacent lines (like Q18 above or Q20 below) never leak in!
      const itemsInBox = [];
      for (const item of textContent.items) {
        if (!item.str || !item.str.trim()) continue;
        const tx = window.pdfjsLib.Util.transform(
          viewport.transform,
          item.transform
        );
        const fontH = Math.hypot(tx[2], tx[3]) || 12;
        const itemW = (item.width || 8) * docState.scale;
        // tx[4] is left X, tx[5] is baseline Y (character body sits above baseline)
        const centerX = tx[4] + Math.min(itemW * 0.25, 12);
        const centerY = tx[5] - fontH * 0.35;

        if (
          centerX >= rx - 3 &&
          centerX <= rx + rw + 3 &&
          centerY >= ry - 2 &&
          centerY <= ry + rh + 2
        ) {
          itemsInBox.push({ x: tx[4], y: centerY, str: item.str.trim() });
        }
      }

      // Group into horizontal lines (tolerance 6px) and sort left-to-right within each line
      itemsInBox.sort((a, b) => (Math.abs(a.y - b.y) < 6 ? a.x - b.x : a.y - b.y));
      const joinedText = itemsInBox.map((l) => l.str).join(' ').replace(/\s+/g, ' ').trim();

      if (joinedText) {
        if (modeType === 'passage') {
          document.getElementById('fieldQuestionText').value = joinedText;
          runAutoMistakeDiagnosis();
          showToast('📖 已附加克漏字文章截圖並帶入上下文敘述！');
          return;
        }
        parseQuestionAndOptionsFromRawText(joinedText);
        runAutoMistakeDiagnosis();
        return;
      }
    } catch (err) {
      console.warn('PDF text extraction error:', err);
    }
  }

  runAutoMistakeDiagnosis();
  showToast('📸 已擷取題目畫面截圖！請點選 STEP 2 的誤選與正解，系統將自動為您分析！');
}

// Smart Parser that handles Cloze Option Rows, Multi-Question Crops, and Standard Questions without jumbling!
function parseQuestionAndOptionsFromRawText(rawText) {
  const multiBar = document.getElementById('multiQuestionSelectorBar');
  const multiBtns = document.getElementById('multiQuestionButtons');
  if (multiBar) multiBar.style.display = 'none';

  if (!rawText || !rawText.trim()) return;
  const cleaned = rawText.replace(/\s+/g, ' ').trim();

  // 1. Check if the cropped region contains one or more numbered option rows like:
  // "( ) 19. (A) that (B) who (C) whom (D) which ( ) 20. (A) used for..."
  // Notice we also strip any orphan tail words before the first question number (e.g. "madden ( ) 19...")
  const multiQRegex =
    /(?:[\(（\[]\s*[\)）\]]\s*)?(\d{1,3})\s*[\.、．\)]\s*([^]*?[\(（\[][A-Da-d][\)）\]\.、\s][^]*?)(?=(?:[\(（\[]\s*[\)）\]]\s*)?\d{1,3}\s*[\.、．\)]\s*[^]*?[\(（\[][A-Da-d][\)）\]\.、\s]|$)/g;

  const detectedQuestions = [];
  let m;
  while ((m = multiQRegex.exec(cleaned)) !== null) {
    const qNum = m[1];
    const body = m[2].trim();
    const parsedOpts = splitStemAndFourOptions(body);
    if (parsedOpts) {
      detectedQuestions.push({
        qNum,
        stem: parsedOpts.stem,
        optA: parsedOpts.optA,
        optB: parsedOpts.optB,
        optC: parsedOpts.optC,
        optD: parsedOpts.optD
      });
    }
  }

  if (detectedQuestions.length > 0) {
    setQuestionTypeUI('choice');
    applyParsedQuestionObject(detectedQuestions[0]);

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
          runAutoMistakeDiagnosis();
        });
      });

      showToast(
        `✂️ 框選範圍內偵測到 ${detectedQuestions.length} 題（第 ${detectedQuestions
          .map((d) => d.qNum)
          .join('、')} 題），已自動幫您拆開！可點擊按鈕切換題號。`
      );
    } else {
      showToast(
        `✂️ 已精準擷取第 ${detectedQuestions[0].qNum} 題的 (A)(B)(C)(D) 選項！`
      );
    }
    return;
  }

  // 2. Fallback: Single question with (A)(B)(C)(D) options
  const singleParsed = splitStemAndFourOptions(cleaned);
  if (singleParsed) {
    const isReadingLike =
      (singleParsed.stem && singleParsed.stem.length > 140) ||
      /according to|passage|article|infer|main idea|title|上文|根據本文|圖表|右圖|下圖/i.test(
        singleParsed.stem
      );
    setQuestionTypeUI(isReadingLike ? 'reading' : 'choice');
    if (singleParsed.stem) {
      document.getElementById('fieldQuestionText').value = singleParsed.stem;
    }
    document.getElementById('fieldOptA').value = singleParsed.optA;
    document.getElementById('fieldOptB').value = singleParsed.optB;
    document.getElementById('fieldOptC').value = singleParsed.optC;
    document.getElementById('fieldOptD').value = singleParsed.optD;
    showToast('✂️ 已自動拆解 (A)(B)(C)(D) 四個選項！');
  } else {
    // 3. No (A)(B)(C)(D) options found -> Automatically switch to Non-Choice mode (Fill / Calculation vs. Translation)!
    document.getElementById('fieldQuestionText').value = cleaned;
    document.getElementById('fieldOptA').value = '';
    document.getElementById('fieldOptB').value = '';
    document.getElementById('fieldOptC').value = '';
    document.getElementById('fieldOptD').value = '';

    const qNumMatch = cleaned.match(/^\s*[\(（\[]?(\d{1,3})[\)）\]\.、．\s]/);
    if (qNumMatch) {
      document.getElementById('fieldQuestionNum').value = qNumMatch[1];
      autoFillAnswerFromKeyMap(qNumMatch[1]);
    }

    const isFillOrCalc =
      /_{2,}|[a-zA-Z]\s*_{2,}\s*[a-zA-Z]|\([a-zA-Z]+\)|計算|求|等於多少|濃度|速度|加速度|面積|體積|=\s*\?/.test(
        cleaned
      );
    if (isFillOrCalc) {
      setQuestionTypeUI('fill');
      showToast(
        '🔤 偵測為「文意字彙／填充／計算題（非選擇題）」，已自動切換為填空作答模式！'
      );
    } else {
      setQuestionTypeUI('translation');
      showToast(
        '✍️ 偵測為「翻譯／簡答／非選題」，已自動隱藏 A/B/C/D 選項，請在 STEP 2 輸入您的寫法與標準解答！'
      );
    }
  }
}

function splitStemAndFourOptions(text) {
  // Match (A) ... (B) ... (C) ... (D) ...
  const optRegex =
    /^(.*?)\s*[\(（\[]\s*A\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*B\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*C\s*[\)）\]]\s*(.+?)\s*[\(（\[]\s*D\s*[\)）\]]\s*(.+)$/is;
  const match = text.match(optRegex);
  if (!match) return null;

  const cleanOpt = (s) =>
    s
      .replace(/[\(（\[]\s*[\)）\]]\s*\d{1,3}\s*[\.、．].*$/, '') // strip any accidental trailing "( ) 20."
      .trim();

  return {
    stem: match[1].replace(/^[\(（\[]\s*[\)）\]]\s*/, '').trim(),
    optA: cleanOpt(match[2]),
    optB: cleanOpt(match[3]),
    optC: cleanOpt(match[4]),
    optD: cleanOpt(match[5])
  };
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
  });

  // Auto-split stem and options button
  document.getElementById('btnParseRawQuestion').addEventListener('click', () => {
    const raw = document.getElementById('fieldQuestionText').value;
    parseQuestionAndOptionsFromRawText(raw);
    runAutoMistakeDiagnosis(true);
    showToast('✨ 已整理題目格式並更新自動診斷！');
  });

  // Auto-lookup standard answer when typing question number
  document.getElementById('fieldQuestionNum').addEventListener('input', (e) => {
    autoFillAnswerFromKeyMap(e.target.value);
    runAutoMistakeDiagnosis(true);
  });

  ['fieldQuestionText', 'fieldOptA', 'fieldOptB', 'fieldOptC', 'fieldOptD', 'fieldSubject'].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', () => runAutoMistakeDiagnosis(true));
  });

  // Wrong & Correct Option Pills -> Immediately trigger automatic diagnosis!
  document.querySelectorAll('#wrongOptionPills .option-pill').forEach((btn) => {
    btn.addEventListener('click', () => {
      setOptionPillAndInput('wrong', btn.getAttribute('data-val'));
      runAutoMistakeDiagnosis(true);
    });
  });

  document.querySelectorAll('#correctOptionPills .option-pill').forEach((btn) => {
    btn.addEventListener('click', () => {
      setOptionPillAndInput('correct', btn.getAttribute('data-val'));
      runAutoMistakeDiagnosis(true);
    });
  });

  document.getElementById('fieldWrongAnswer').addEventListener('input', (e) => {
    const val = e.target.value;
    if (state.questionType === 'choice' || state.questionType === 'reading') {
      setOptionPillAndInput('wrong', val.trim().toUpperCase(), false);
    }
    runAutoMistakeDiagnosis(true);
  });

  document.getElementById('fieldCorrectAnswer').addEventListener('input', (e) => {
    const val = e.target.value;
    if (state.questionType === 'choice' || state.questionType === 'reading') {
      setOptionPillAndInput('correct', val.trim().toUpperCase(), false);
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

  // Auto-Analyze Button (runs built-in smart analyzer + Gemini Vision if API Key set)
  const btnAutoAnalyze = document.getElementById('btnAutoAnalyzeMistake');
  if (btnAutoAnalyze) {
    btnAutoAnalyze.addEventListener('click', async () => {
      runAutoMistakeDiagnosis(true);
      const savedKey = localStorage.getItem('mistake_gemini_api_key') || '';
      if (savedKey) {
        await runGeminiVisionDiagnosis(savedKey);
      } else {
        showToast('🤖 已根據您的題型、答案與科目自動完成錯因診斷與正確寫法解析！');
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

function setQuestionTypeUI(qType) {
  state.questionType = qType || 'choice';
  document
    .querySelectorAll('#questionTypeContainer .diagnosis-tag')
    .forEach((btn) => {
      btn.classList.toggle(
        'selected',
        btn.getAttribute('data-qtype') === state.questionType
      );
    });

  const isChoiceLike =
    state.questionType === 'choice' || state.questionType === 'reading';
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

  if (state.questionType === 'translation') {
    if (labelWrong) labelWrong.textContent = '❌ 我當初寫錯的翻譯／簡答內容';
    if (labelCorrect) labelCorrect.textContent = '✅ 標準參考翻譯／正確解答';
    if (wrongInput)
      wrongInput.placeholder =
        '請輸入你當時寫的翻譯句子或簡答（例如：It spend the company 3 years...）';
    if (correctInput)
      correctInput.placeholder =
        '請輸入正確的標準翻譯或解答（例如：It took the company three years to...）';
  } else if (state.questionType === 'fill') {
    if (labelWrong) labelWrong.textContent = '❌ 我填錯的單字／算錯的答案';
    if (labelCorrect) labelCorrect.textContent = '✅ 正確單字／標準計算結果';
    if (wrongInput)
      wrongInput.placeholder = '例如：拼錯的單字 poison 或算錯的數值 20';
    if (correctInput)
      correctInput.placeholder = '例如：正確單字 poisonous 或正確數值 10 m/s';
  } else if (state.questionType === 'reading') {
    if (labelWrong) labelWrong.textContent = '❌ 我誤選的閱讀／圖表題選項';
    if (labelCorrect) labelCorrect.textContent = '✅ 正確答案選項';
    if (wrongInput) wrongInput.placeholder = '點上方 A/B/C/D 或輸入誤選答案';
    if (correctInput) correctInput.placeholder = '點上方 A/B/C/D 或輸入正確答案';
    if (diffBox) diffBox.style.display = 'none';
  } else {
    if (labelWrong) labelWrong.textContent = '❌ 我當初選錯的答案';
    if (labelCorrect) labelCorrect.textContent = '✅ 試卷正確答案';
    if (wrongInput) wrongInput.placeholder = '點上方 A/B/C/D 或輸入誤選答案';
    if (correctInput) correctInput.placeholder = '點上方 A/B/C/D 或輸入正確答案';
    if (diffBox) diffBox.style.display = 'none';
  }
}

function setOptionPillAndInput(which, val, updateInput = true) {
  const containerId = which === 'wrong' ? 'wrongOptionPills' : 'correctOptionPills';
  const inputId = which === 'wrong' ? 'fieldWrongAnswer' : 'fieldCorrectAnswer';

  if (updateInput) {
    document.getElementById(inputId).value = val;
  }

  document.querySelectorAll(`#${containerId} .option-pill`).forEach((b) => {
    b.classList.toggle('selected', b.getAttribute('data-val') === val);
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
// Multi-Subject & Multi-Question-Type Automatic Diagnostic Engine
// Supports:
//  1. ✍️ 翻譯 / 簡答 / 手寫非選題 (Translation & Short Answer)
//  2. 🔤 文意字彙 / 填充 / 物理・化學・數學・地科計算題 (Fill & Science Calc)
//  3. 📖 閱讀測驗 / 地理・歷史・公民圖表題組 (Reading & Chart Analysis)
//  4. 🔘 選擇題 / 克漏字 (All Subjects: 英文、數學、物理、化學、生物、地理、歷史、公民)
// ============================================================================
function runAutoMistakeDiagnosis(forceOverwrite = true) {
  const rawWrong = document.getElementById('fieldWrongAnswer').value.trim();
  const rawCorrect = document.getElementById('fieldCorrectAnswer').value.trim();
  const subject = document.getElementById('fieldSubject').value.trim() || '綜合科目';
  const stem = document.getElementById('fieldQuestionText').value.trim();
  const opts = {
    A: document.getElementById('fieldOptA').value.trim(),
    B: document.getElementById('fieldOptB').value.trim(),
    C: document.getElementById('fieldOptC').value.trim(),
    D: document.getElementById('fieldOptD').value.trim()
  };

  if (!rawWrong || !rawCorrect) {
    return;
  }

  // Auto-detect if user typed a full word/sentence/number instead of A/B/C/D while still in 'choice' mode
  const isSingleChoiceLetter = (s) => /^[A-Ea-e]$/.test(s.trim());
  const hasAnyOptionText = Boolean(opts.A || opts.B || opts.C || opts.D);
  let effectiveQType = state.questionType;

  if (
    effectiveQType === 'choice' &&
    (!isSingleChoiceLetter(rawWrong) || !isSingleChoiceLetter(rawCorrect)) &&
    !hasAnyOptionText
  ) {
    const looksLikeSentence =
      rawCorrect.split(/\s+/).length >= 3 || rawCorrect.length >= 15;
    effectiveQType = looksLikeSentence ? 'translation' : 'fill';
    setQuestionTypeUI(effectiveQType);
  }

  const diffBox = document.getElementById('liveTextDiffBox');
  const diffContent = document.getElementById('liveTextDiffContent');

  let autoTag = '觀念混淆';
  let whyWrong = '';
  let correctConcept = '';
  let keyTakeaway = '';

  // =========================================================================
  // MODE A: ✍️ 翻譯／造句／簡答非選題 (Translation & Open-Ended Short Answer)
  // =========================================================================
  if (effectiveQType === 'translation') {
    const diffRes = computeWordDiffAnalysis(rawWrong, rawCorrect);
    if (diffBox && diffContent) {
      diffBox.style.display = 'block';
      diffContent.innerHTML = diffRes.diffHtml;
    }

    autoTag = '片語與字彙搭配';
    const wrongPartStr =
      diffRes.wrongWords.length > 0
        ? `你寫了「${diffRes.wrongWords.join(' ')}」`
        : '你的句型結構未完整表達';
    const rightPartStr =
      diffRes.missingWords.length > 0
        ? `標準答案應為「${diffRes.missingWords.join(' ')}」`
        : `標準寫法為「${rawCorrect}」`;

    // Detect specific English translation pitfalls (tense, spend/take, to V vs V-ing, etc.)
    const specificNotes = [];
    if (/\b(take|takes|took)\b/i.test(rawCorrect) && /\bspend|spent|cost\b/i.test(rawWrong)) {
      specificNotes.push(
        '混淆了「花費時間」的句型：事物/虛主詞 It 作主詞時要用「It takes/took + 人/組織 + 時間 + to V」（人作主詞才用 spend + 時間 + V-ing）'
      );
    }
    if (/\b(took|was|were|had|did|went|spent|made|found)\b/i.test(rawCorrect) && !/\b(took|was|were|had|did|went|spent|made|found|ed)\b/i.test(rawWrong)) {
      specificNotes.push('忽略了「過去式時態」：發生在過去的事件，動詞必須改為過去式');
    }
    if (/\bto\s+[a-z]+\b/i.test(rawCorrect) && /\b[a-z]+ing\b/i.test(rawWrong)) {
      specificNotes.push('不定詞 (to + V) 與動名詞 (V-ing) 搭配混淆');
    }

    whyWrong = `【逐字比對診斷】：${wrongPartStr}，但${rightPartStr}！${
      specificNotes.length > 0 ? '\n👉 關鍵失分原因：' + specificNotes.join('；') + '。' : '\n👉 關鍵失分原因：關鍵片語搭配、動詞時態或單複數詞形與標準句型有落差。'
    }`;

    correctConcept = `【這題應該這樣寫】：\n1. 先抓出整句的主詞與動詞時態骨架。\n2. 套用本題核心片語與修飾語：${
      diffRes.missingWords.length > 0 ? `特別注意「${diffRes.missingWords.join(' ')}」的正確拼法與位置` : '確認介系詞與連接詞位置'
    }。\n3. 完整正確解答：\n「${rawCorrect}」`;

    keyTakeaway =
      diffRes.missingWords.length > 0
        ? `必背關鍵寫法：❌ 勿寫 ${diffRes.wrongWords.join(' ') || '原錯句'} 👉 ✅ 正確應寫「${diffRes.missingWords.join(' ')}」！`
        : `熟記標準句型：「${rawCorrect}」`;
  }
  // =========================================================================
  // MODE B: 🔤 文意字彙／填充／物理・化學・數學・地理計算與填空題 (Fill & Calc)
  // =========================================================================
  else if (effectiveQType === 'fill') {
    const diffRes = computeWordDiffAnalysis(rawWrong, rawCorrect);
    if (diffBox && diffContent) {
      diffBox.style.display = 'block';
      diffContent.innerHTML = diffRes.diffHtml;
    }

    // Check if this is a Numeric / Science / Math Calculation question
    const wNumMatch = rawWrong.match(/-?\d+(\.\d+)?/);
    const cNumMatch = rawCorrect.match(/-?\d+(\.\d+)?/);
    const isNumericCalc =
      Boolean(wNumMatch && cNumMatch) &&
      (/數學|物理|化學|自然|地科|地球科學|地理/.test(subject) ||
        /^[\d\.\-\+\/\*\sA-Za-z%°]+$/.test(rawCorrect));

    if (isNumericCalc) {
      autoTag = '計算粗心';
      const wVal = parseFloat(wNumMatch[0]);
      const cVal = parseFloat(cNumMatch[0]);
      let mathDiag = '';

      if (cVal !== 0 && Math.abs(wVal + cVal) < 1e-6) {
        mathDiag = `你算出的數值「${rawWrong}」與正確答案「${rawCorrect}」【正負號剛好相反】！請檢查：是否忽略了向量方向（如減速運動加速度為負、位移方向）、吸熱/放熱正負號，或解方程式移項時忘了變號。`;
      } else if (cVal !== 0 && (Math.abs(wVal / cVal - 2) < 0.01 || Math.abs(cVal / wVal - 2) < 0.01)) {
        mathDiag = `你算出的「${rawWrong}」與正解「${rawCorrect}」【剛好相差 2 倍】！常見原因：公式漏乘 1/2（例如動能 Ek = 1/2 mv²、等加速度位移 s = 1/2 at²、三角形面積），或是化學反應式莫耳係數比（1:2）忘了乘除！`;
      } else if (cVal !== 0 && (Math.abs(wVal / cVal - 3.6) < 0.05 || Math.abs(cVal / wVal - 3.6) < 0.05)) {
        mathDiag = `你算出的「${rawWrong}」與正解「${rawCorrect}」相差 3.6 倍！這是因為【速率單位 km/h 與 m/s 忘記換算】（1 m/s = 3.6 km/h）！`;
      } else if (cVal !== 0 && [10, 100, 1000, 0.1, 0.01, 0.001].some((r) => Math.abs(wVal / cVal - r) < 1e-5)) {
        mathDiag = `你算出的「${rawWrong}」與正解「${rawCorrect}」相差 10 的次方倍！請檢查【單位換算】（例如：cm 與 m、g 與 kg、mL 與 L 換算成莫耳濃度 M）。`;
      } else {
        mathDiag = `你算出的結果為「${rawWrong}」，但正確答案應為「${rawCorrect}」。問題出在代入公式時的條件判讀或中間四則運算誤差。`;
      }

      const sciFormulaGuide = getSubjectFormulaGuidance(subject, stem, rawCorrect);
      whyWrong = mathDiag;
      correctConcept = `【這題應該這樣算】：\n${sciFormulaGuide}\n👉 最終正確答案為：${rawCorrect}`;
      keyTakeaway = `列式後必做 3 檢查：①單位有無換算(m/s、L、kg) ②公式係數(如 1/2)與正負號 ③題目問的最終物理量！`;
    } else {
      // English Vocabulary Fill-in-the-blank or Word-Form Transformation (e.g. poison -> poisonous, c_____l -> chemical)
      autoTag = '文法與詞性判斷';
      const wClean = rawWrong.toLowerCase().trim();
      const cClean = rawCorrect.toLowerCase().trim();

      // Check if wClean and cClean share the same root (Word family transformation error!)
      if (wClean.length >= 3 && cClean.length >= 3 && cClean.slice(0, 3) === wClean.slice(0, 3)) {
        const wPos = guessEnglishPartOfSpeech(wClean);
        const cPos = guessEnglishPartOfSpeech(cClean);
        whyWrong = `【詞性與字尾變化錯誤】：你填了「${rawWrong}」（${wPos}），已經抓對字根了！但沒注意到空格在句子結構中需要的是【${cPos}】或是複數形，因此必須變化字尾為「${rawCorrect}」！`;
        correctConcept = `【這題應該這樣寫】：\n1. 先看空格前後的詞性線索（例如：系動詞 are/is + 副詞 usually 後方要填「形容詞」；所有格 their/its 後方要填「名詞」；助動詞/to 後方填「原形動詞」）。\n2. 將字根變化為正確詞性與單複數：填入「${rawCorrect}」（${cPos}）。`;
        keyTakeaway = `文意字彙填完後必檢查：①詞性字尾(-ous形容詞 / -tion,-ings名詞 / -ly副詞) ②名詞單複數(-s) ③動詞時態(-ed)！`;
      } else {
        whyWrong = `你填寫了「${rawWrong}」，但根據本題前後文語意與首尾字母線索，正確答案應為「${rawCorrect}」（共 ${cClean.length} 個字母，字首 ${cClean[0] || ''}、字尾 ${cClean[cClean.length - 1] || ''}）。`;
        correctConcept = `【這題應該這樣寫】：\n1. 圈出題幹中的上下文關鍵搭配字（Context Clues）。\n2. 判斷空格所需的詞性（${guessEnglishPartOfSpeech(cClean)}）與正負向語氣，正確單字/答案為：「${rawCorrect}」。`;
        keyTakeaway = `必背核心字彙／填空：「${rawCorrect}」— 注意其拼字（${cClean}）與前後文搭配！`;
      }
    }
  }
  // =========================================================================
  // MODE C: 📖 閱讀測驗／地理・歷史・公民・自然圖表題組 (Reading & Charts)
  // =========================================================================
  else if (effectiveQType === 'reading') {
    if (diffBox) diffBox.style.display = 'none';
    const wrongAns = rawWrong.toUpperCase();
    const correctAns = rawCorrect.toUpperCase();
    const wrongText = opts[wrongAns] || `選項 (${wrongAns})`;
    const correctText = opts[correctAns] || `選項 (${correctAns})`;

    autoTag = '審題失誤';
    const isMainIdea = /main idea|title|purpose|best describes|主旨|標題|目的|核心/.test(stem);
    const isInference = /infer|imply|suggest|conclude|推論|推測|暗示|得知/.test(stem);
    const isGeoOrChart = /地理|歷史|公民|地球科學|自然|物理|化學/.test(subject) || /圖|表|曲線|座標|氣候|等高線|洋流|人口/.test(stem);

    if (isGeoOrChart) {
      const geoGuide = getSubjectChartAndConceptGuidance(subject, stem, wrongText, correctText);
      whyWrong = `你選了 (${wrongAns})「${wrongText}」，在判讀題幹圖表／資料時，誤將局部特徵或相反條件當成結論（${geoGuide.trapReason}）。`;
      correctConcept = `【這題應該這樣判讀】：\n${geoGuide.howToSolve}\n因此正確答案應選 (${correctAns})「${correctText}」。`;
      keyTakeaway = geoGuide.takeaway;
    } else if (isMainIdea) {
      whyWrong = `【閱讀主旨題診斷】你選了 (${wrongAns})「${wrongText}」，這是閱讀測驗最常見的「以偏概全（太狹隘，只講到某一段的細節）」或「範圍過大」誘答選項！`;
      correctConcept = `【這題應該這樣找答案】：\n1. 主旨題（Main Idea / Best Title）要看「第一段結尾（Thesis Statement）」與「各段第一句（Topic Sentence）」。\n2. 正確答案 (${correctAns})「${correctText}」才能涵蓋整篇文章的核心主題，而非單一細節！`;
      keyTakeaway = `閱讀主旨題避坑：只出現在單一段落的細節選項必刪！緊扣首段轉折後與末段總結句！`;
    } else if (isInference) {
      whyWrong = `【閱讀推論題診斷】你選了 (${wrongAns})「${wrongText}」，掉入了「過度推論（文章根本沒提到）」或「直接照抄原文單字但扭曲因果關係」的陷阱！`;
      correctConcept = `【這題應該這樣找答案】：\n1. 推論題（infer / imply）的正確答案，一定是原文某句話的「正向邏輯延伸或同義改寫（Paraphrase）」。\n2. 對照原文線索，只有 (${correctAns})「${correctText}」有明確的上下文證據支持！`;
      keyTakeaway = `閱讀推論題鐵則：沒講到的不能自己腦補！選項有絕對詞（always/never/all/only）通常有詐！`;
    } else {
      whyWrong = `【閱讀細節定位診斷】你選了 (${wrongAns})「${wrongText}」，這個選項雖然出現了文章裡的眼熟單字，但把「主詞／時間／因果關係」張冠李戴了！`;
      correctConcept = `【這題應該這樣找答案】：\n1. 先從題幹抓出定位關鍵字（人名、地名、專有名詞、轉折詞 However / But / In fact 後面的句子）。\n2. 回到文章對應段落逐字比對：正確答案 (${correctAns})「${correctText}」使用了「同義字替換（Paraphrase）」重述原文重點！`;
      keyTakeaway = `閱讀測驗秒殺技巧：正解最愛「換句話說（同義字替換）」，完全照抄原文詞彙卻改掉因果的選項是陷阱！`;
    }
  }
  // =========================================================================
  // MODE D: 🔘 選擇題／克漏字（全科支援：英文、物理、化學、生物、數學、地理、歷史、公民）
  // =========================================================================
  else {
    if (diffBox) diffBox.style.display = 'none';
    const wrongAns = rawWrong.toUpperCase();
    const correctAns = rawCorrect.toUpperCase();
    const wrongText = opts[wrongAns] || `選項 (${wrongAns})`;
    const correctText = opts[correctAns] || `選項 (${correctAns})`;

    const relPronouns = ['that', 'who', 'whom', 'which', 'whose', 'where', 'when', 'what'];
    const isRelPronounQ =
      relPronouns.includes(wrongText.toLowerCase()) ||
      relPronouns.includes(correctText.toLowerCase());

    if (isRelPronounQ) {
      autoTag = '文法與詞性判斷';
      const wLow = wrongText.toLowerCase();
      const cLow = correctText.toLowerCase();
      const relExplainMap = {
        that: '「that」雖可代替人或物，但【絕對不能用在逗號 ( , ) 後方的非限定關係子句】，也不能直接放在介系詞後面！',
        who: '「who」用於先行詞是【人】，且在關係子句中擔任【主詞】（後面緊接動詞，如 have / is / V）！',
        whom: '「whom」用於先行詞是【人】，且在關係子句中擔任【受詞】（後面會接「主詞＋動詞」）或放在介系詞後！',
        which: '「which」用於先行詞是【事物／動物】或代表前面整個子句，不能用來指代「人（如 people / students）」！',
        whose: '「whose」是所有格關代（翻譯為「...的」），後面必須緊接一個【名詞】！'
      };

      whyWrong = `你選了 (${wrongAns}) ${wrongText}，但正確答案是 (${correctAns}) ${correctText}。\n判斷錯誤點：${
        relExplainMap[wLow] || `誤判了關係代名詞「${wrongText}」的先行詞或子句結構。`
      }`;
      correctConcept = `【這題應該這樣寫】：\n1. 先找空格前的「先行詞」是人還是物（若為 people / person 等人，先排除 which）。\n2. 檢查空格前後有沒有「逗號 ,」（若有逗號屬於補述用法，不可選 that）。\n3. 看空格後面緊接的是「動詞」還是「主詞＋動詞」：${
        relExplainMap[cLow] || `因此此處必須填入 (${correctAns}) ${correctText}。`
      }`;
      keyTakeaway = `關代秒殺三步驟：①看先行詞(人/物) ②看有無逗號(有逗號禁選 that) ③看後接動詞(選主格 who)還是主詞(選受格 whom)！`;
    } else if (isEnglishWordFamilyQuestion(opts)) {
      autoTag = '文法與詞性判斷';
      const wPos = guessEnglishPartOfSpeech(wrongText);
      const cPos = guessEnglishPartOfSpeech(correctText);
      whyWrong = `你選了 (${wrongAns}) ${wrongText}（${wPos}），但此空格在句子結構中需要的詞性是「${cPos}」，所以不能選 (${wrongAns})！`;
      correctConcept = `【這題應該這樣寫】：\n這題是經典的「詞性判斷題（同字根變化）」！\n1. 觀察空格前後的字詞結構，判斷空格缺的是主詞/受詞（填名詞）、修飾名詞或作受詞補語（填形容詞）、修飾一般動詞/整句（填副詞 -ly），還是助動詞/to 後面（填原形動詞）。\n2. 選項中 (${correctAns}) ${correctText} 為【${cPos}】，完全符合此句型位置需求，故選 (${correctAns})。`;
      keyTakeaway = `遇到四個選項長得很像的「詞性題」，先看空格前後決定詞性（-ness/-tion 名詞、-ly 副詞、-en/-ize 動詞），不用硬翻整段文章就能秒殺！`;
    } else if (isEnglishOptionSet(opts)) {
      const phraseDict = {
        'used for': 'be used for + V-ing（被用來做...，主詞通常是物品）',
        'keen on': 'be keen on + V-ing（熱衷於、喜愛做...）',
        'tired of': 'be tired of + V-ing（對...感到厭倦、受夠了...）',
        'amazed by': 'be amazed by（對...感到驚奇讚嘆）',
        'for a while': 'for a while（持續一小段時間）',
        'for example': 'for example（舉例來說，用於舉出具體實例）',
        'in the end': 'in the end（最後、終於，表結果）',
        'all of a sudden': 'all of a sudden（突然間＝suddenly，表突發狀況）',
        'go through': 'go through（經歷困難、仔細檢查）',
        'carry out': 'carry out（執行、實行計畫或命令）',
        'call on': 'call on（拜訪；呼籲、號召某人做某事）',
        'come across': 'come across（偶然遇見、無意間發現）',
        'because': 'because + 完整子句 (S + V)（因為...）',
        'because of': 'because of + 名詞 / V-ing（因為...，後面絕對不能接完整子句）',
        'since': 'since（自從...；既然/因為，後接子句）',
        'once': 'once + 子句（一旦...就...）',
        'in case': 'in case + 子句（以防萬一...，表事先準備）',
        'even though': 'even though + 子句（儘管、雖然...，表讓步轉折）',
        'as soon as': 'as soon as + 子句（一...就...）',
        'unless': 'unless + 子句（除非...＝if...not）',
        'infected': 'infected（adj. 受感染的，常搭配 be infected with）',
        'sensitive': 'sensitive（adj. 敏感的、易受傷害的，常搭配 be sensitive to 對...敏感）',
        'creative': 'creative（adj. 有創意的）',
        'plastic': 'plastic（n./adj. 塑膠的）',
        'distance': 'distance（n. 距離）',
        'convenience': 'convenience（n. 便利、方便）',
        'survival': 'survival（n. 生存、存活）',
        'substance': 'substance（n. 物質、化學物質）'
      };
      const wMeaning = phraseDict[wrongText.toLowerCase()] || `「${wrongText}」`;
      const cMeaning = phraseDict[correctText.toLowerCase()] || `「${correctText}」`;
      autoTag = '片語與字彙搭配';
      whyWrong = `你選了 (${wrongAns}) ${wrongText}，它的意思是 ${wMeaning}。放進本題前後文時，語意邏輯或後面接的詞性搭配不符，因此判斷錯誤。`;
      correctConcept = `【這題應該這樣寫】：\n1. 先看空格前後的關鍵字與介系詞／子句結構。\n2. 正確答案 (${correctAns}) ${correctText} 的用法為 ${cMeaning}，帶入句中不論是上下文語意還是文法搭配都完全吻合！`;
      keyTakeaway = `必背對比：誤選的 ${wrongText} vs. 正解 ${cMeaning}，記住後面的介系詞與上下文語氣！`;
    } else {
      // Multi-Subject Analyzer for Physics, Chemistry, Biology, Earth Science, Math, Geography, History, Civics!
      const subjGuide = getSubjectChartAndConceptGuidance(
        subject,
        `${stem} ${wrongText} ${correctText}`,
        wrongText,
        correctText
      );
      const isAskWrong = /何者錯誤|何者有誤|不正確|不包括|何者為非/.test(stem);
      if (isAskWrong) {
        autoTag = '審題失誤';
        whyWrong = `這題題幹問的是「何者錯誤／不正確」！你選的 (${wrongAns})「${wrongText}」本身敘述其實是正確的（或你沒注意到題目要選錯的）。\n${subjGuide.trapReason}`;
        correctConcept = `【這題應該這樣解】：\n1. 看到「何者錯誤」先畫雙底線，逐一在 (A)(B)(C)(D) 旁標註 ⭕ 或 ❌。\n2. ${subjGuide.howToSolve}\n3. 因此敘述有誤、應選的答案為 (${correctAns})「${correctText}」！`;
        keyTakeaway = `問「何者錯誤」先標 ⭕/❌！${subjGuide.takeaway}`;
      } else {
        autoTag = '觀念混淆';
        whyWrong = `你誤選了 (${wrongAns})「${wrongText}」。\n👉 判斷錯誤點：${subjGuide.trapReason}`;
        correctConcept = `【這題應該這樣解】：\n${subjGuide.howToSolve}\n👉 因此正確答案應選 (${correctAns})「${correctText}」。`;
        keyTakeaway = subjGuide.takeaway;
      }
    }
  }

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
    badge.textContent = `🤖 已根據【${subject}・${getQuestionTypeLabel(
      effectiveQType
    )}】自動分析錯因與正確寫法！（看完即可直接儲存）`;
  }
}

function getQuestionTypeLabel(qType) {
  const map = {
    choice: '選擇／克漏字',
    translation: '翻譯／簡答非選題',
    fill: '填充／文意字彙／計算題',
    reading: '閱讀測驗／圖表題組'
  };
  return map[qType] || '選擇題';
}

// Subject-Specific Formula & Concept Knowledge Base (Physics, Chemistry, Math, Geography, History, Biology...)
function getSubjectFormulaGuidance(subject, stem, correctAns) {
  const text = `${subject} ${stem} ${correctAns}`;
  if (/加速度|速度|位移|煞車|自由落體|等加速|v-t/.test(text)) {
    return '1. 先確認初速度 (v₀)、末速度 (v)、加速度 (a) 與時間 (t) 的方向（減速時 a 與 v 反向）。\n2. 套用等加速度三大公式：v = v₀ + at、s = v₀t + 1/2 at²、v² = v₀² + 2as（或利用 v-t 圖斜率求加速度、面積求位移）。';
  }
  if (/力|牛頓|摩擦|質量|F\s*=|重量/.test(text)) {
    return '1. 先畫出物體的「力圖（Free-body diagram）」，標出重力、正向力、外力與摩擦力。\n2. 套用牛頓第二運動定律 ΣF = ma（注意必須是「合力」而非單一外力；等速運動時合力為 0）。';
  }
  if (/功|動能|位能|功率|焦耳|能量/.test(text)) {
    return '1. 確認外力方向與位移方向夾角（W = F·s·cosθ，垂直時作功為 0）。\n2. 套用功能定理（合力作功＝動能變化 ΔEk = 1/2 mv² - 1/2 mv₀²）或力學能守恆（Ek + U = 定值）。';
  }
  if (/電流|電壓|電阻|歐姆|串聯|並聯|電功率/.test(text)) {
    return '1. 先判別電路是「串聯（電流 I 相同，電阻相加）」還是「並聯（電壓 V 相同，電阻倒數相加）」。\n2. 套用歐姆定律 V = I × R 與電功率公式 P = I·V = I²R = V²/R。';
  }
  if (/莫耳|濃度|化學|分子量|反應式|係數|pH|酸|鹼|氧化|還原/.test(text)) {
    return '1. 先將質量換成莫耳數（n = W / M，或由體積乘以莫耳濃度 n = C_M × V(公升)）。\n2. 列出平衡後的化學反應式，利用「係數比＝反應莫耳數比」（注意誰是限量試劑）求出產物或目標濃度。';
  }
  if (/氣候|雨量|氣溫|風帶|洋流|等高線|地形|經緯|時區|人口|產業/.test(text)) {
    return '1. 【地理判讀三步驟】：先看座標軸／圖例單位（如最冷月均溫決定熱帶>18°C、溫帶-3~18°C；最高溫在1月為南半球、7月為北半球）。\n2. 結合緯度、海陸位置、行星風帶（信風/西風/副熱帶高壓）或地形迎風/背風坡進行綜合推導。';
  }
  return '1. 先圈出題幹給定的已知數據與單位（若有非標準單位先換算）。\n2. 寫出該單元核心公式，將已知數代入求解，並檢查正負號與倍數關係。';
}

function getSubjectChartAndConceptGuidance(subject, text, wrongText, correctText) {
  if (/物理|理化/.test(subject) || /加速度|速度|力|功|電|波|折射|透鏡|浮力|壓力/.test(text)) {
    return {
      trapReason: `在物理觀念上，誤把「${wrongText}」成立的特定條件當成通則（例如混淆了速度與加速度方向、靜摩擦與動摩擦、或未考慮合力是否為零）。`,
      howToSolve: `1. 先確認物理量的定義（向量需同時看大小與方向，如減速時 a 與 v 反向；轉彎時必有向心加速度）。\n2. 列出對應定律（如 F=ma、P=F/A、浮力 B=V_排×d_液），即可驗證「${correctText}」才符合題意。`,
      takeaway: `物理避坑：①向量必看方向(正負號) ②等速＝合力為零 ③看清圖表縱軸是 x 還是 v！`
    };
  }
  if (/化學/.test(subject) || /莫耳|原子|離子|酸鹼|氧化|還原|電解|有機|酯|反應速率|平衡/.test(text)) {
    return {
      trapReason: `在化學概念上，把「${wrongText}」與「${correctText}」的反應條件、粒子數／莫耳數比或酸鹼／氧化還原定義記顛倒了。`,
      howToSolve: `1. 檢查化學反應式係數與粒子狀態（分子 vs. 原子 vs. 離子，是否完全解離）。\n2. 依據勒沙特列原理／莫耳計量／氧化數變化，推導出「${correctText}」。`,
      takeaway: `化學避坑：①係數比＝莫耳數比＝氣體體積比（非質量比！） ②看清問的是分子數還是原子總數！`
    };
  }
  if (/地理/.test(subject) || /氣候|季風|洋流|等高線|地形|冰河|石灰岩|都市|農業|工業|人口|GIS/.test(text)) {
    return {
      trapReason: `在地理空間與圖表判讀上，誤判了「${wrongText}」所屬的氣候帶／地形特徵／區位條件（例如沒注意南北半球月份相反、迎風坡與背風坡、或等高線彎曲方向）。`,
      howToSolve: `1. 先從題幹或圖表抓出「空間定位線索（緯度、海陸位置、最冷月均溫、雨季分布或等高線數值）」。\n2. 對照地理成因（行星風系／洋流性質／地形作用力／產業區位要素），即可鎖定「${correctText}」。`,
      takeaway: `地理秒殺口訣：先定半球（7月熱北、1月熱南）➔ 二看溫雨定氣候 ➔ 等高線凸向高處為山谷(發源河流)！`
    };
  }
  if (/生物|地科|地球科學/.test(subject) || /細胞|酵素|遺傳|染色體|生態|板塊|斷層|氣團|鋒面|星等|潮汐/.test(text)) {
    return {
      trapReason: `將「${wrongText}」與「${correctText}」的构造功能、作用階段或地質／天文週期特徵搞混了。`,
      howToSolve: `1. 先釐清題目問的對象層次（如胞器功能、減數分裂 vs. 細胞分裂、冷鋒 vs. 暖鋒、朔望月位置）。\n2. 逐一比對定義特徵，唯有「${correctText}」完全吻合。`,
      takeaway: `生科／地科重點：善用對比表格記憶相似名詞（發電廠＝粒線體、冷鋒＝鋒後降雨、逆斷層＝受擠壓）！`
    };
  }
  if (/歷史|公民/.test(subject)) {
    return {
      trapReason: `受「${wrongText}」字面敘述吸引，但忽略了題幹的時代背景時間軸（歷史）或法定構成要件／救濟程序（公民）。`,
      howToSolve: `1. 先圈出題幹中的時代關鍵字／制度名稱或法律行為主體。\n2. 排除時代不符或權責機關錯誤的「${wrongText}」，選出符合史實／法理的「${correctText}」。`,
      takeaway: `社會科解題關鍵：先抓題幹「時間／空間／主體關鍵字」，切勿只看選項本身通順就選！`
    };
  }
  return {
    trapReason: `將「${wrongText}」的適用條件與正確答案「${correctText}」混淆，或忽略了題幹中的關鍵限定詞。`,
    howToSolve: `1. 先圈出題幹的核心條件與問題焦點。\n2. 對照「${wrongText}」與「${correctText}」的差異，確認只有「${correctText}」完全符合定義與題意。`,
    takeaway: `熟記核心差異：❌ 誤區「${wrongText.slice(0, 18)}」 vs. ✅ 正解「${correctText.slice(0, 18)}」！`
  };
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
        text: `你是一位極具教學經驗的「${subject}」名師。
本題題型：${getQuestionTypeLabel(state.questionType)}（題號：${qNum || '?'}）
學生當初寫錯／選錯的答案：「${wrongAns}」
標準正確答案：「${correctAns}」
目前題目文字：${stem}
選項（若為選擇題）：(A) ${opts.A} (B) ${opts.B} (C) ${opts.C} (D) ${opts.D}
請根據題目文字與附圖（若有），回傳純 JSON 格式（不要加 markdown 標籤）：
{
  "questionText": "若附圖中有更完整的題目或句子，請補全題目文字，否則回傳原題目",
  "whyWrong": "直接告訴學生他的答案「${wrongAns}」判斷或書寫錯誤的地方在哪裡（具體指出錯字、文法、公式、圖表判讀或觀念誤區，2~3句話）",
  "correctConcept": "告訴學生這一題應該怎麼看、怎麼寫／怎麼算才正確（清楚步驟 2~3 句話）",
  "keyTakeaway": "給學生的一句話秒殺記憶口訣或公式提醒"
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
  const subject = document.getElementById('fieldSubject').value.trim() || '綜合科目';
  const unit = document.getElementById('fieldUnit').value.trim() || '一般練習';
  const questionNum = document.getElementById('fieldQuestionNum').value.trim() || String(state.mistakes.length + 1);
  const questionText = document.getElementById('fieldQuestionText').value.trim();
  const optA = document.getElementById('fieldOptA').value.trim();
  const optB = document.getElementById('fieldOptB').value.trim();
  const optC = document.getElementById('fieldOptC').value.trim();
  const optD = document.getElementById('fieldOptD').value.trim();
  const isChoiceLike = state.questionType === 'choice' || state.questionType === 'reading';
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
  document.getElementById('fieldOptA').value = '';
  document.getElementById('fieldOptB').value = '';
  document.getElementById('fieldOptC').value = '';
  document.getElementById('fieldOptD').value = '';
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
    if (subj !== 'ALL' && m.subject !== subj) return false;
    if (status !== 'ALL' && m.masteryStatus !== status) return false;
    return true;
  });
}

function buildMistakeCardHtml(m, includeFooterButtons = true) {
  const hasOptions = m.optA || m.optB || m.optC || m.optD;
  const optionsHtml = hasOptions
    ? `<div class="mcard-options-list">
        ${['A', 'B', 'C', 'D']
          .map((letter) => {
            const text = m[`opt${letter}`];
            if (!text) return '';
            const isWrong = m.wrongAnswer.toUpperCase() === letter;
            const isCorrect = m.correctAnswer.toUpperCase() === letter;
            const cls = isCorrect
              ? 'is-correct'
              : isWrong
              ? 'is-wrong'
              : '';
            const badge = isCorrect ? '✅ ' : isWrong ? '❌ ' : '';
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
          建立：${escapeHtml(m.createdAt)}｜已練習 ${m.practiceCount} 次（答對 ${m.correctCount} 次）
        </div>
        <div style="display:flex; gap:0.4rem;">
          <button class="btn btn-secondary btn-sm" onclick="editMistakeById('${m.id}')">✏️ 編輯</button>
          <button class="btn btn-danger-outline btn-sm" onclick="deleteMistakeById('${m.id}')">🗑️</button>
          <button class="btn btn-primary btn-sm" onclick="downloadMistakePngById('${m.id}')">📥 下載錯題 PNG</button>
        </div>
      </div>`
    : `<div style="padding:0.65rem 1.15rem; background:#f8fafc; border-top:1px solid #e2e8f0; display:flex; justify-content:space-between; font-size:0.78rem; color:#64748b;">
        <span>📘 錯題進化論｜${escapeHtml(state.sheetName)}</span>
        <span>記錄日期：${escapeHtml(m.createdAt)}｜熟練度：${escapeHtml(m.masteryStatus)}</span>
      </div>`;

  return `
    <div class="mistake-card-exportable" id="mcard-dom-${m.id}">
      <div class="mcard-header">
        <div style="display:flex; align-items:center; gap:0.5rem; flex-wrap:wrap;">
          <span class="mcard-subject-badge">${escapeHtml(m.subject)}</span>
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
  setQuestionTypeUI(m.questionType || 'choice');
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
    if (subj !== 'ALL' && m.subject !== subj) return false;
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
  const stdKeys = ['A', 'B', 'C', 'D', 'E'];
  const cUp = String(m.correctAnswer || '').trim().toUpperCase();
  const hasOpts = Boolean(m.optA || m.optB || m.optC || m.optD);
  if (!hasOpts && !stdKeys.includes(cUp)) {
    return true;
  }
  return false;
}

function renderCurrentQuizQuestion() {
  const m = state.quizQueue[state.quizIndex];
  state.quizSelectedOptKey = null;
  state.quizSelectedConceptIdx = null;
  state.quizSubmitted = false;

  const qTypeBadge = getQuestionTypeLabel(m.questionType);
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

  document.getElementById('quizQuestionStem').textContent = m.questionText;

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
    const hasOptionText = Boolean(m.optA || m.optB || m.optC || m.optD);

    let renderedOptions = [
      { origKey: 'A', text: m.optA || '選項 (A)' },
      { origKey: 'B', text: m.optB || '選項 (B)' },
      { origKey: 'C', text: m.optC || '選項 (C)' },
      { origKey: 'D', text: m.optD || '選項 (D)' }
    ];

    const stdKeys = ['A', 'B', 'C', 'D'];
    if (
      !stdKeys.includes(m.correctAnswer.toUpperCase()) ||
      !stdKeys.includes(m.wrongAnswer.toUpperCase())
    ) {
      renderedOptions = [
        { origKey: m.correctAnswer, text: m.correctAnswer },
        { origKey: m.wrongAnswer, text: m.wrongAnswer }
      ];
      if (m.optA) renderedOptions.push({ origKey: 'A', text: m.optA });
      if (m.optB) renderedOptions.push({ origKey: 'B', text: m.optB });
    }

    if (shuffleOpts && (hasOptionText || renderedOptions.length !== 4)) {
      renderedOptions.sort(() => Math.random() - 0.5);
    }

    const displayLabels = ['A', 'B', 'C', 'D'];
    optionsContainer.innerHTML = renderedOptions
      .map((opt, idx) => {
        const dispLetter = displayLabels[idx] || String(idx + 1);
        const note =
          shuffleOpts && hasOptionText ? '（選項順序已隨機洗牌）' : '';
        return `
          <button type="button" class="quiz-option-btn" data-orig-key="${escapeHtml(opt.origKey)}">
            <span style="width:30px; height:30px; border-radius:50%; background:#e2e8f0; display:inline-flex; align-items:center; justify-content:center; font-weight:700;">
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
        state.quizSelectedOptKey = btn.getAttribute('data-orig-key');
        optionsContainer
          .querySelectorAll('.quiz-option-btn')
          .forEach((b) => b.classList.remove('selected'));
        btn.classList.add('selected');
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
    state.currentQuestionRenderData = { conceptChoices, useTextInputMode };

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
    state.currentQuestionRenderData = { conceptChoices: [], useTextInputMode };
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
      x.subject === currentMistake.subject &&
      x.correctConcept
  );
  if (other) {
    choices.push({ text: other.correctConcept, isTrue: false });
  } else if (
    currentMistake.questionType === 'translation' ||
    /英|English/i.test(currentMistake.subject || '')
  ) {
    choices.push({
      text: '直接照中文逐字順序硬翻即可，不需要考慮英文主詞動詞一致、時態或介系詞固定搭配。',
      isTrue: false
    });
  } else if (/地理|歷史|公民|社會/.test(currentMistake.subject || '')) {
    choices.push({
      text: '看圖表時不需要看座標軸單位與圖例，直接憑直覺挑選數值最大的一格即可。',
      isTrue: false
    });
  } else {
    choices.push({
      text: '直接將題目出現的所有數字相加除以二即可得出答案，無須考慮方向、正負號或公式定義限制。',
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
      alert('請先選擇一個題目答案選項！');
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
  } else {
    isAnswerCorrect =
      userAnswerRaw.toUpperCase() === m.correctAnswer.toUpperCase();
  }

  const isConceptCorrect = verifyEnabled
    ? state.currentQuestionRenderData.conceptChoices[state.quizSelectedConceptIdx]
        ?.isTrue === true
    : true;

  const fullPass = isAnswerCorrect && isConceptCorrect;

  // Highlight option buttons if multiple-choice
  if (!useTextInputMode) {
    document
      .querySelectorAll('#quizOptionsContainer .quiz-option-btn')
      .forEach((btn) => {
        const k = btn.getAttribute('data-orig-key').toUpperCase();
        if (k === m.correctAnswer.toUpperCase()) {
          btn.classList.add('correct-reveal');
        } else if (k === userAnswerRaw.toUpperCase() && !isAnswerCorrect) {
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
            ${diffReport ? diffReport.html : ''}
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
            m.optA || m.optB
              ? `<div style="font-size:0.76rem; color:#64748b;">(A)${escapeHtml(m.optA)} (B)${escapeHtml(m.optB)} (C)${escapeHtml(m.optC)} (D)${escapeHtml(m.optD)}</div>`
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
    { wch: 12 }, // 科目
    { wch: 20 }, // 試卷單元
    { wch: 8 },  // 題號
    { wch: 42 }, // 題目內容
    { wch: 20 }, // 選項A
    { wch: 20 }, // 選項B
    { wch: 20 }, // 選項C
    { wch: 20 }, // 選項D
    { wch: 12 }, // 我的錯誤答案
    { wch: 10 }, // 正確答案
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

  state.sheetName = '國九段考衝刺錯題本_範例';
  state.sheetMode = 'existing';
  state.originalLoadedCount = 3;
  state.mistakes = [
    {
      id: 'DEMO_1',
      createdAt: new Date().toLocaleDateString('zh-TW'),
      subject: '自然科學',
      unit: '直線運動與加速度',
      questionNum: '1',
      questionText:
        '一輛汽車在筆直公路上向東行駛，司機看到紅燈後踩下煞車做等減速度運動直到停止。請問在煞車期間，關於汽車的「速度方向」與「加速度方向」，下列敘述何者錯誤？',
      optA: '汽車的速度方向依然向東',
      optB: '汽車的位移方向向東',
      optC: '汽車的速度越來越慢，因此加速度方向也向東',
      optD: '汽車的加速度方向向西（與速度方向相反）',
      wrongAnswer: 'D',
      correctAnswer: 'C',
      errorTags: ['審題失誤', '觀念混淆'],
      whyWrong:
        '題目最後問的是「何者錯誤」，我太急看成「何者正確」，看到 (D) 敘述是對的就直接選下去！',
      correctConcept:
        '減速運動時，加速度方向與速度方向「相反」（車向東開但減速，故加速度向西）。因此選項 (C) 說加速度向東是錯誤的，本題選錯誤選項應選 (C)。',
      keyTakeaway:
        '看到題目問「何者錯誤」先畫雙底線！減速運動＝加速度與速度反向！',
      practiceCount: 1,
      correctCount: 0,
      masteryStatus: '待加強',
      cropImageDataUrl: '',
      isHistorical: true
    },
    {
      id: 'DEMO_2',
      createdAt: new Date().toLocaleDateString('zh-TW'),
      subject: '數學',
      unit: '一元二次方程式與乘法公式',
      questionNum: '2',
      questionText:
        '計算並化簡代數式 (2x - 3)² - (x - 5)(x + 5) 的結果，下列哪一個選項是正確的？',
      optA: '3x² - 12x - 16',
      optB: '3x² - 12x + 34',
      optC: '3x² - 6x + 34',
      optD: '5x² - 12x - 16',
      wrongAnswer: 'A',
      correctAnswer: 'B',
      errorTags: ['計算粗心', '掉入陷阱'],
      whyWrong:
        '展開 -(x - 5)(x + 5) 時，寫成 -x² - 25，忘記括號前面有負號，負負得正應該變成 +25，結果常數項算成 9 - 25 = -16 而誤選 (A)。',
      correctConcept:
        '(2x - 3)² - (x² - 25) = 4x² - 12x + 9 - x² + 25 = 3x² - 12x + 34，正確答案為 (B)。',
      keyTakeaway: '減號後面有乘法公式展開時，一定要先加中括號再變號！',
      practiceCount: 2,
      correctCount: 1,
      masteryStatus: '複習中',
      cropImageDataUrl: '',
      isHistorical: true
    },
    {
      id: 'DEMO_3',
      createdAt: new Date().toLocaleDateString('zh-TW'),
      subject: '生物',
      unit: '細胞的構造與功能',
      questionNum: '3',
      questionText:
        '小明使用複式顯微鏡觀察水蘊草葉片細胞，若他想找尋被稱為「細胞能量發電廠」（負責進行呼吸作用產生能量）的胞器，應該觀察下列哪一個構造？',
      optA: '葉綠體',
      optB: '液泡',
      optC: '核糖體',
      optD: '粒線體',
      wrongAnswer: 'A',
      correctAnswer: 'D',
      errorTags: ['觀念混淆'],
      whyWrong:
        '看到題目寫「水蘊草葉片細胞」是植物，就直覺選了 (A) 葉綠體，把光合作用（製造養分）跟呼吸作用（產生能量發電廠）搞混了。',
      correctConcept:
        '無論動物或植物細胞，負責進行呼吸作用釋放能量的「細胞發電廠」都是 (D) 粒線體；葉綠體則是進行光合作用製造葡萄糖的場所。',
      keyTakeaway: '細胞發電廠＝粒線體（呼吸作用）；養分製造機＝葉綠體（光合作用）！',
      practiceCount: 2,
      correctCount: 2,
      masteryStatus: '已精通',
      cropImageDataUrl: '',
      isHistorical: true
    }
  ];

  document.getElementById('sheetInitModal').classList.add('hidden');
  saveToLocalStorage();
  updateAllViews();
  showToast(
    '✨ 已載入範例試卷與 3 筆跨科目錯題紀錄！可體驗「框選題目」、「下載 PNG 錯題卡」、「強化再練習」與「匯出試算表」！'
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
  const subjects = Array.from(
    new Set(state.mistakes.map((m) => m.subject).filter(Boolean))
  );
  const selects = [
    document.getElementById('filterCardSubject'),
    document.getElementById('quizFilterSubject')
  ];
  selects.forEach((sel) => {
    const prevVal = sel.value || 'ALL';
    sel.innerHTML =
      `<option value="ALL">全部科目（${state.mistakes.length}）</option>` +
      subjects
        .map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`)
        .join('');
    if (subjects.includes(prevVal)) sel.value = prevVal;
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
