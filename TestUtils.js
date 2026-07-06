/**
 * =========================================================
 * TEST UTILITIES & MOCK STUBS (TestUtils.gs)
 * =========================================================
 * Provides lightweight in-memory mocks for all Google Apps
 * Script services so that Tests.gs can execute without
 * a live Google Sheets / Drive / Gmail connection.
 *
 * HOW TO USE:
 *  1. Call MockFactory.reset() at the start of every test.
 *  2. Replace service globals with Mock objects before
 *     invoking the function under test.
 *  3. Read MockFactory state to assert outcomes.
 * =========================================================
 */

// ─────────────────────────────────────────────
// ASSERTION HELPERS
// ─────────────────────────────────────────────

const Assert = {
  /**
   * Throws an error if actual !== expected.
   */
  equals: (actual, expected, message) => {
    if (actual !== expected) {
      throw new Error(
        `❌ ASSERT FAILED: ${message || ''}\n   Expected: ${JSON.stringify(expected)}\n   Actual:   ${JSON.stringify(actual)}`
      );
    }
  },

  /**
   * Throws if value is falsy.
   */
  isTrue: (value, message) => {
    if (!value) {
      throw new Error(`❌ ASSERT FAILED (isTrue): ${message || ''} — Got: ${JSON.stringify(value)}`);
    }
  },

  /**
   * Throws if value is truthy.
   */
  isFalse: (value, message) => {
    if (value) {
      throw new Error(`❌ ASSERT FAILED (isFalse): ${message || ''} — Got: ${JSON.stringify(value)}`);
    }
  },

  /**
   * Throws if value is null or undefined.
   */
  notNull: (value, message) => {
    if (value == null) {
      throw new Error(`❌ ASSERT FAILED (notNull): ${message || ''}`);
    }
  },

  /**
   * Throws if value is NOT null/undefined.
   */
  isNull: (value, message) => {
    if (value != null) {
      throw new Error(`❌ ASSERT FAILED (isNull): ${message || ''} — Got: ${JSON.stringify(value)}`);
    }
  },

  /**
   * Asserts that fn() throws an error.
   */
  throws: (fn, message) => {
    try {
      fn();
      throw new Error(`❌ ASSERT FAILED (throws): ${message || ''} — No error was thrown.`);
    } catch (e) {
      if (e.message && e.message.startsWith('❌ ASSERT FAILED')) throw e;
      // Any other error is acceptable — test passes
    }
  }
};


// ─────────────────────────────────────────────
// SIMPLE TEST RUNNER
// ─────────────────────────────────────────────

const TestRunner = {
  _results: [],

  /**
   * Registers and runs a single test case.
   * @param {string} name  - Human-readable test name
   * @param {function} fn  - Test body; throws on failure
   */
  run: (name, fn) => {
    try {
      fn();
      TestRunner._results.push({ name, status: 'PASS' });
      Logger.log(`✅ PASS — ${name}`);
    } catch (e) {
      TestRunner._results.push({ name, status: 'FAIL', error: e.message });
      Logger.log(`❌ FAIL — ${name}\n   ${e.message}`);
    }
  },

  /**
   * Prints a summary table to the Apps Script log.
   */
  summary: () => {
    const total  = TestRunner._results.length;
    const passed = TestRunner._results.filter(r => r.status === 'PASS').length;
    const failed = total - passed;

    Logger.log('═══════════════════════════════════════════════');
    Logger.log(`  TEST SUMMARY:  ${passed} / ${total} passed   |   ${failed} failed`);
    Logger.log('═══════════════════════════════════════════════');

    TestRunner._results
      .filter(r => r.status === 'FAIL')
      .forEach(r => Logger.log(`  ⚠️  ${r.name}\n     ${r.error}`));

    TestRunner._results = []; // Reset for next run
  }
};


// ─────────────────────────────────────────────
// IN-MEMORY SPREADSHEET MOCK
// ─────────────────────────────────────────────

/**
 * MockFactory builds fresh in-memory data stores per test run.
 */
const MockFactory = {
  /** In-memory "sheets" keyed by sheet name. */
  _sheets: {},


  /** Drive folders/files created. */
  _driveFolders: {},

  /**
   * Resets all mock state. Call at the start of each test.
   */
  reset: () => {
    MockFactory._sheets      = {};
    MockFactory._driveFolders = {};
  },

  /**
   * Seeds a sheet with headers + optional row data.
   * @param {string}   sheetName
   * @param {string[]} headers
   * @param {Array[]}  rows      - Each element is an array of values
   */
  seedSheet: (sheetName, headers, rows = []) => {
    MockFactory._sheets[sheetName] = {
      _data: [headers, ...rows],

      getDataRange: () => ({
        getValues: () => MockFactory._sheets[sheetName]._data.map(r => [...r])
      }),

      appendRow: (rowArray) => {
        MockFactory._sheets[sheetName]._data.push([...rowArray]);
      },

      getRange: (rowIndex, colIndex) => ({
        setValue: (value) => {
          MockFactory._sheets[sheetName]._data[rowIndex - 1][colIndex - 1] = value;
        }
      })
    };
  },

  /**
   * Returns the mock SpreadsheetApp object.
   * Replaces the global SpreadsheetApp for testing.
   */
  getSpreadsheetApp: () => ({
    openById: (_id) => ({
      getSheetByName: (name) => {
        if (!MockFactory._sheets[name]) {
          return null; // Simulates missing sheet
        }
        return MockFactory._sheets[name];
      }
    })
  }),

  /**
   * Returns mock DriveApp.
   */
  getDriveApp: () => {
    const root = {
      _folders: {},
      getFoldersByName: (name) => {
        const match = root._folders[name];
        return {
          hasNext: () => !!match,
          next:    () => match
        };
      },
      createFolder: (name) => {
        const folder = MockFactory._makeMockFolder(name);
        root._folders[name] = folder;
        return folder;
      }
    };

    return {
      getFoldersByName: root.getFoldersByName,
      createFolder:     root.createFolder,
      getFolderById:    (id) => {
        // Find folder across all roots by folderId
        const all = Object.values(MockFactory._driveFolders);
        return all.find(f => f.getId() === id) || MockFactory._makeMockFolder('retrieved-' + id);
      },
      Access: { PRIVATE: 'PRIVATE' },
      Permission: { NONE: 'NONE' }
    };
  },

  /** Builds a mock Drive folder object. */
  _makeMockFolder: (name) => {
    const folderId = 'mock-folder-' + name.replace(/\s/g, '_');
    const folder = {
      _files:   {},
      _subFolders: {},
      getId:    () => folderId,
      getUrl:   () => 'https://drive.google.com/drive/folders/' + folderId,
      getName:  () => name,
      setSharing: (_a, _p) => {},
      createFile: (blob) => {
        const file = {
          _name: blob.getName ? blob.getName() : 'unknown',
          getName:   function() { return this._name; },
          setName:   function(n) { this._name = n; },
          getUrl:    () => 'https://drive.google.com/file/mock-' + folderId
        };
        folder._files[file._name] = file;
        return file;
      },
      getFilesByName: (fname) => {
        const match = folder._files[fname];
        return {
          hasNext: () => !!match,
          next:    () => match
        };
      },
      getFoldersByName: (fname) => {
        const match = folder._subFolders[fname];
        return {
          hasNext: () => !!match,
          next:    () => match
        };
      },
      createFolder: (fname) => {
        const sub = MockFactory._makeMockFolder(fname);
        folder._subFolders[fname] = sub;
        MockFactory._driveFolders[fname] = sub;
        return sub;
      }
    };
    MockFactory._driveFolders[name] = folder;
    return folder;
  },

  /**
   * Returns mock Session.
   */
  getSession: () => ({
    getActiveUser: () => ({ getEmail: () => 'test.user@yourcompany.com' })
  }),

  /**
   * Returns mock Utilities.
   */
  getUtilities: () => ({
    getUuid:      () => 'mock-uuid-' + Math.random().toString(36).slice(2, 9),
    base64Decode: (data) => data, // passthrough
    newBlob:      (data, mime, name) => ({ getData: () => data, getContentType: () => mime, getName: () => name })
  }),

  /**
   * Returns mock ScriptApp.
   */
  getScriptApp: () => ({
    getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/mock/exec' })
  })
};
