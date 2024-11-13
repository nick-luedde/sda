/**
 * Class for handling data access to app data google sheet
 * Expects each collection to have an id property that is treated as the unique identifier for that record
 * @type {ISheetDataAccess}
 */
class SheetDataAccess {

  /**
   * Constructor function
   * @param {object} source - Spreadsheet source options
   * @param {string} [source.id] - Spreadsheet id
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [source.ss] - Spreadsheet object
   * @param {object} [options] - options object
   * @param {{ [key: string]: AsvContext }} [options.schemas] - optional Schemas to apply to the datasource objects
   */
  constructor({ id, ss }, { schemas } = {}) {

    this.spreadsheet = ss || SpreadsheetApp.openById(id);
    this.collections = {};
    this.hasSchema = !!schemas;

    const sheets = this.spreadsheet.getSheets();
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      //allow for additional sheets to be left out of the datasource if prefixed with _;
      if (sheetName[0] !== '_') {

        /** @type {AsvContext | undefined} */
        let schema;
        if (this.hasSchema) {
          schema = schemas[sheetName];
          if (!schema)
            throw new Error(`${sheetName} has no schema model provided!`);
        }

        this.collections[sheetName] = new SheetDataCollection(sheet, { schema });
      }
    });
  }

  /**
   * Static helper prop for the offset from the Sheet row and the eventual data array index
   */
  static get ROW_INDEX_OFFSET() {
    return 2;
  }

  /**
   * Cell usage cap
   * //https://support.google.com/drive/answer/37603
   */
  static get CELL_CAP() {
    return 10000000;
  }

  /**
   * maps an array of data to an object with headers of the row as property keys
   * @param {object[]} row - row of data to map to object
   * @param {number} index - index of the object within the data array
   * @param {string[]} headers - array of header names in the order of appearance in sheet
   */
  static getRowAsObject(row, index, headers) {
    const obj = {
      _key: index + SheetDataAccess.ROW_INDEX_OFFSET
    };

    headers.forEach((header, index) => obj[header] = row[index]);
    return obj;
  };

  /**
   * Clears all empty rows from all collections
   */
  defrag() {
    Object.values(this.collections).forEach(coll => coll.defrag());
    return this;
  }

  /**
   * Archives entire spreadsheet content
   */
  wipe() {
    Object.values(this.collections).forEach(coll => coll.wipe());
    return this;
  }

  /**
   * Archives entire spreadsheet content
   * @param {string} folderId - id of sheet to archive to
   */
  archive(folderId) {
    const folder = DriveApp.getFolderById(folderId);
    const copy = this.spreadsheet.copy(`${this.spreadsheet.getName()}_${new Date().toJSON()}`);
    const file = DriveApp.getFileById(copy.getId());
    file.moveTo(folder);

    return this.wipe();
  }

  /**
   * Returns a usage percent report
   */
  inspect() {

    const breakdowns = Object.values(this.collections)
      .map(coll => coll.inspect());

    const totalRows = breakdowns.map(rep => rep.totalRows).reduce((sum, count) => sum += count, 0);
    const totalColumns = breakdowns.map(rep => rep.totalColumns).reduce((sum, count) => sum += count, 0);
    const totalCells = breakdowns.map(rep => rep.totalCells).reduce((sum, count) => sum += count, 0);
    const usagePercent = breakdowns.map(rep => rep.usagePercent).reduce((ttl, pct) => ttl += pct, 0) / (breakdowns.length || 1);

    const report = {
      summary: {
        totalRows,
        totalColumns,
        totalCells,
        usagePercent
      },
      breakdowns
    };

    return report;
  }

}

/**
 * Class that manages read writes to a specific collection of data based on a sheet
 * @type {ISheetDataCollection}
 */
class SheetDataCollection {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - sheet for the collection of data
   * @param {object} [options] - collections options
   * @param {AsvContext} [options.schema] - schema to apply to the collection 
   */
  constructor(sheet, { schema } = {}) {
    this.sheet = sheet;
    this.schema = schema;
    this.hasModel = !!schema;
    this._data = null;
    this._index = {};
    this._related = {};
    this._pkColumnIndex = 0;
  }

  /**
   * Sets the index of the pk column (only necessary if not 0)
   * @param {number} index - col index of pk field
   */
  pk(index) {
    this._pkColumnIndex = index;
    return this;
  }

  /**
   * setup any props needed for data writing methods
   */
  _init() {
    this.COLUMN_COUNT = this.sheet.getLastColumn();

    //check for cache
    if (!this._data) {
      this.headerRow = this.sheet.getRange(1, 1, 1, this.COLUMN_COUNT).getValues()[0];
    }
  }

  _values() {
    return this.sheet.getSheetValues(1, 1, this.rowCount(), this.sheet.getLastColumn());
  }

  /**
   * @param {any[]} row - data row array
   * @param {number} index - index of the data in the dataset
   */
  _getObject(row, index) {
    const obj = SheetDataAccess.getRowAsObject(row, index, this.headerRow);
    return this.hasModel ? this.schema.parse(obj) : obj;
  }

  /**
   * Gets shallow copies of records to save, applies schema if exists
   * @param {object[]} records - records to get saveable array
   */
  _getRecordsToSave(records, { ignoreErrors } = {}) {
    if (!this.hasModel)
      return records.map(rec => ({ ...rec }));

    return records.map((rec) => this.schema.exec(rec, { isNew: !rec._key, throwError: !ignoreErrors }));
  }

  /**
   * Gets records from schema or shallow copies if none
   * @param {object[]} records - records to get from schema 
   */
  _getFromSchemaRecords(records) {
    if (!this.hasModel)
      return records.map(rec => ({ ...rec }));

    return records.map(this.schema.parse);
  }

  /**
   * Gets current row count
   */
  rowCount() {
    return this.sheet.getLastRow();
  }

  /**
   * Helper to clear cached data to force refreshes
   */
  clearCached() {
    this._data = null;
    this._index = {};
    this._related = {};

    return this;
  }

  /**
   * Helper function that will replace the top row of a sheet with headers from the provided obj
   * @param {object} obj - object with headers to write
   */
  writeHeadersFromObject(obj) {
    const sheet = this.sheet;
    const firstRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
    firstRow.clear();

    const headers = Object.keys(obj);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);

    return this;
  }

  /**
   * Caches and returns a unique key map
   * @param {string} [key] - optional id
   */
  index(key = 'id') {
    if (!this._index[key]) {

      const data = this.data();

      const index = data.reduce((obj, record) => {
        obj[record[key]] = record;
        return obj;
      }, {});

      this._index[key] = index;
    }

    return this._index[key];
  }

  /**
   * Creates and returns a map of related items
   * @param {string} key - property key of related set to get
   */
  related(key) {
    if (!this._related[key]) {

      const data = this.data();

      const related = data.reduce((obj, record) => {
        if (!obj[record[key]])
          obj[record[key]] = [record]
        else
          obj[record[key]].push(record);

        return obj;
      }, {});

      this._related[key] = related;
    }

    return this._related[key];
  }

  /**
   * Enforces uniqueness of a given prop within the collection (throws an error if not unique)
   * @param {object} rec - record to enforce unique
   * @param {string} prop - prop to enforce being unique
   */
  enforceUnique(rec, prop) {
    if (!rec._key) {
      const index = this.index(prop);
      if (index[rec[prop]] !== undefined)
        throw new Error(`${this.sheet.getName()} ${prop} prop value ${rec[prop]} already exists!`);
    } else {
      // is the best to just filter it out??
      const others = this.data().filter(oth => oth._key !== rec._key);
      const set = new Set(others.map(oth => oth[prop]));
      if (set.has(rec[prop]))
        throw new Error(`${this.sheet.getName()} ${prop} prop value ${rec[prop]} already exists!`);
    }

    return this;
  }

  /**
   * Handles retrieving and caching item data from sheet
   */
  data() {
    if (this._data === null) {
      const values = this._values();

      this.headerRow = values.shift();
      this._data = [];

      values.forEach((row, index) => {
        if (row[this._pkColumnIndex] !== '')
          this._data.push(this._getObject(row, index));
      });
    }

    return this._data;
  }

  get advanced() {
    const context = this;
    return {
      /**
       * Gets data through advanced sheets service (more performance at larger scale)
       */
      data() {
        if (context._data === null) {
          const results = Sheets.Spreadsheets.Values.get(context.sheet.getParent().getId(), context.sheet.getName());
          context.headerRow = results.values.shift();

          context._data = [];

          results.values.forEach((row, index) => {
            if (row[context._pkColumnIndex] !== '')
              context._data.push(context._getObject(row, index));
          });
        }

        return context._data;
      }
    }
  }

  /**
   * Streams data in chunks...
   */
  *stream(size) {
    const CHUNK_SIZE = size || 5000;
    let i = 0;

    const sheet = this.sheet;
    this._init();

    const rows = this.rowCount()
    const columns = this.COLUMN_COUNT;

    const chunks = Math.ceil(rows / CHUNK_SIZE);

    while (i < chunks) {
      const startRow = i * CHUNK_SIZE + SheetDataAccess.ROW_INDEX_OFFSET;
      const rowsToGet = Math.min(rows - startRow, CHUNK_SIZE);
      const values = sheet.getSheetValues(startRow, 1, rowsToGet, columns);

      const data = [];
      values.forEach((row, index) => {
        if (row[this._pkColumnIndex] !== '')
          data.push(this._getObject(row, startRow - SheetDataAccess.ROW_INDEX_OFFSET + index));
      });

      i++;
      yield data;
    }
  }

  /**
   * Finds a record by a given key
   * @param {string} key - key of record to get
   * @param {string} [index] - optional index to use, defaults to '_key'
   */
  find(key, index = '_key') {
    const idx = this.index(index);
    return idx[key];
  }

  /**
   * Gets a row by key (row number)
   * @param {number} key 
   */
  get(key) {
    this._init();

    const keynum = Number(key);

    const [row] = this.sheet.getRange(keynum, 1, 1, this.COLUMN_COUNT).getValues();
    if (row[this._pkColumnIndex] !== '')
      return this._getObject(row, keynum - SheetDataAccess.ROW_INDEX_OFFSET);
    else
      return null;
  }

  /**
   * Performs an efficient lookup for a single record by value
   * @param {any} val 
   * @param {string} key 
   */
  lookup(val, key = 'id') {
    if (!this._data === null) {
      return this.find(val, key);
    } else {
      return this.fts({ q: val, matchCell: true }).find(r => r[key] === val);
    }
  }

  /**
   * Performs an efficient lookup for all records by value
   * @param {any} val 
   * @param {string} key 
   */
  many(val, key = 'id') {
    if (!this._data === null) {
      return this.find(val, key);
    } else {
      return this.fts({ q: val, matchCell: true }).filter(r => r[key] === val);
    }
  }

  /**
   * Saves record objects to sheet datasource
   * @param {object[]} records - record objects to save
   */
  upsert(records, { bypassSchema = false } = {}) {
    if (records.length === 0)
      return records;
    //saves record data objects to the spreadsheet
    //check for cache
    this._init();

    const schemaModels = !bypassSchema ? this._getRecordsToSave(records) : records;

    const updates = schemaModels.filter(rec => rec._key !== undefined && rec._key !== null);
    const adds = schemaModels.filter(rec => rec._key === undefined || rec._key === null);

    this.update(updates, { bypassSchema: true });
    this.add(adds, { bypassSchema: true })

    //clear cached data to force rebuild to account for changed/added records
    this.clearCached();

    return this._getFromSchemaRecords(schemaModels);
  }

  /**
   * Upserts one record (more concurrent safe)
   * @param {object} record - record to upsert
   * @param {{ bypassSchema: boolean }} [options] - options
   */
  upsertOne(record, { bypassSchema = false } = {}) {
    if (!record)
      return null;

    const isNew = record._key === undefined || record._key === null;
    const [saved] = isNew ? [this.addOne(record, { bypassSchema })] : this.update([record], { bypassSchema });
    return saved;
  }

  /**
   * Adds one record (more concurrent safe)
   * @param {object} record - record to add
   */
  addOne(record, { bypassSchema = false } = {}) {
    //saves record data objects to the spreadsheet
    if (!record)
      return null;

    const sheet = this.sheet;
    this._init();

    const recordToSave = !bypassSchema
      ? this._getRecordsToSave([record])[0]
      : record;

    const startingRowCount = this.rowCount();

    const row = this.headerRow.map(header => recordToSave[header]);
    sheet.appendRow(row);

    const rng = sheet.getRange(startingRowCount, this._pkColumnIndex + 1, this.rowCount());
    const finder = rng.createTextFinder(row[this._pkColumnIndex]);
    const found = finder.findNext();
    if (!found)
      throw new Error('Something went wrong with the add operation!');
    recordToSave._key = found.getRow();

    //clear cached data to force rebuild to account for changed/added records
    this.clearCached();

    //return objects to their from schema state
    //this seems like i should be doing this a different way....
    return this._getFromSchemaRecords([recordToSave])[0];
  }

  /**
   * adds the record models (NOT concurrent safe, use locking if necessary)
   * @param {object[]} records - records to add to the sheet datasource
   */
  add(records, { bypassSchema = false } = {}) {
    //saves record data objects to the spreadsheet
    if (records.length === 0)
      return [];

    const sheet = this.sheet;
    this._init();

    const recordsToSave = !bypassSchema
      ? this._getRecordsToSave(records)
      : records;

    const rowCount = this.rowCount();
    const recordArrays = recordsToSave
      .map((record, index) => {
        record._key = rowCount + index + 1;
        return this.headerRow.map(header => record[header]);
      });

    const range = sheet.getRange(rowCount + 1, 1, recordsToSave.length, this.COLUMN_COUNT);
    if (!!range.getValues()[0][this._pkColumnIndex])
      throw new Error('Add transaction error, try again!');
    range.setValues(recordArrays);

    //clear cached data to force rebuild to account for changed/added records
    this.clearCached();

    //return objects to their from schema state
    //this seems like i should be doing this a different way....
    return this._getFromSchemaRecords(recordsToSave);
  }

  /**
   * updates the record models
   * @param {object[]} records - records to update in the sheet datasource
   */
  update(records, { bypassSchema = false } = {}) {
    if (records.length === 0)
      return [];
    //saves record data objects to the spreadsheet
    this._init();

    const recordsToSave = !bypassSchema
      ? this._getRecordsToSave(records)
      : records;

    recordsToSave.forEach(record => {
      const recordValues = this.headerRow.map(header => record[header]);
      this._updateRow(record, recordValues, this.COLUMN_COUNT);
    });

    //clear cached data to force rebuild to account for changed/added records
    this.clearCached();

    //return objects to their from schema state
    //this seems like i should be doing this a different way....
    return this._getFromSchemaRecords(recordsToSave);;
  }

  /**
   * Patches the provided patch props onto existing models 
   * (allows for targeted updates, which could help with multi-users so that entire records arent saved, just the individual changes are applied)
   * @param {object[]} patches - list of patches to apply
   * @param {{ bypassSchema: boolean }} [options] 
   */
  patch(patches, { bypassSchema = false } = {}) {
    if (patches.length === 0)
      return [];

    this._init();

    const patchedRecords = patches.map(patch => {
      const existing = this.find(patch._key);
      if (!existing)
        throw new Error(`Could not patch record with key ${patch._key}. Key not found!`);

      return {
        ...existing,
        ...patch
      };
    });

    return this.update(patchedRecords, { bypassSchema });
  }

  /**
   * deletes record objects to sheet datasource
   */
  delete(records) {
    const sheet = this.sheet;
    //check for cache
    this._init();

    // find each record to remove...
    records.forEach(record => {
      const recordValues = this.headerRow.map(header => record[header]);
      const range = sheet.getRange(record._key, 1, 1, this.COLUMN_COUNT);

      //last second check to make sure 2d array and sheet are still in sync for this object
      if (String(range.getValues()[0][this._pkColumnIndex]) !== String(recordValues[this._pkColumnIndex]))
        throw new Error(`Id at row ${record._key} does not match id of object for ${recordValues[this._pkColumnIndex]}`);

      range.setValues([new Array(this.COLUMN_COUNT)]);
    });

    //clear cached data to force rebuild to account for deleted records
    this.clearCached();
  }

  /**
   * Performs batch update on the entire dataset for (meant for faster but more expensive updates)
   * (NOT concurrent safe, use locking if necessary)
   * @param {object[]} records - batch data to apply
   * @param {object} [options] - options
   */
  batch(records, { bypassSchema = false } = {}) {
    if (records.length === 0)
      return records;
    //saves record data objects to the spreadsheet
    //check for cache
    this._init();

    const schemaModels = !bypassSchema ? this._getRecordsToSave(records) : records;

    const updates = schemaModels.filter(rec => rec._key !== undefined && rec._key !== null);
    const adds = schemaModels.filter(rec => rec._key === undefined || rec._key === null);

    // get data content without the header row;
    const data = this._values().slice(1);

    updates.forEach(rec =>
      data.splice(rec._key - SheetDataAccess.ROW_INDEX_OFFSET, 1, this.headerRow.map(hdr => rec[hdr]))
    );
    data.push(...adds.map(rec => this.headerRow.map(hdr => rec[hdr])));

    const lock = LockService.getScriptLock();
    lock.tryLock(30 * 1000);

    if (!lock.hasLock())
      throw new Error('Could not perform batch operation, please try again!');

    this.wipe();

    this.sheet.getRange(2, 1, data.length, this.headerRow.length)
      .setValues(data);

    lock.releaseLock();

    this.clearCached();

    return this._getFromSchemaRecords(schemaModels);
  }

  /**
   * Performs a preflight validation of all records prior to saving
   *  This can be especially usefull with combined data actions that you want to be more transactional (all pass or all fail)
   * returns update/add methods prepared with the preflight records
   * Only allows a single call of a transaction method (will error if one is called again)
   * @param {object | object[]} records - records to prelight validate
   */
  preflight(records) {
    const arrayOfRecords = Array.isArray(records) ? records : [records];
    const recordsToSave = this._getRecordsToSave(arrayOfRecords);
    const bypassSchema = true;

    let transacted = false;
    const transact = (fn) => {
      if (transacted) throw new Error('Preflight transaction already complete!');
      const result = fn();
      transacted = true;

      return result;
    };

    return {
      addOne: () => transact(() => this.addOne(recordsToSave[0], { bypassSchema })),
      add: () => transact(() => this.add(recordsToSave, { bypassSchema })),
      update: () => transact(() => this.update(recordsToSave, { bypassSchema })),
      batch: () => transact(() => this.batch(recordsToSave, { bypassSchema })),
      upsert: () => transact(() => this.upsert(recordsToSave, { bypassSchema })),
      upsertOne: () => transact(() => this.upsertOne(recordsToSave[0], { bypassSchema })),
      record: () => this._getFromSchemaRecords(recordsToSave)[0],
      records: () => this._getFromSchemaRecords(recordsToSave),
    };
  }

  /**
   * Clears all records from the sheet
   * @param {number} [rows] - optional number of rows to delete (all rows if left out)
   */
  wipe(rows) {
    const s = this.sheet;
    const maxRows = s.getMaxRows();
    const rowsToWipe = rows !== undefined && rows >= 0 ? Math.min(rows, maxRows - 1) : maxRows - 1;
    if (rowsToWipe === 0)
      return;

    //Avoid 'Cant delete all non-frozen rows' error by leaving on empty row
    const secondRow = s.getRange(2, 1, 1, s.getMaxColumns());
    secondRow.clear();

    const remaining = rowsToWipe - 1;
    if (remaining > 0)
      s.deleteRows(3, remaining);

    return this;
  }

  /**
   * Removes non-data rows
   */
  defrag() {
    const sheet = this.sheet;
    const data = this._values().filter(row => row[this._pkColumnIndex] !== '').slice(1);
    if (data.length === 0)
      return;

    const maxRow = sheet.getMaxRows();
    if (data.length + 1 === maxRow)
      return;

    this._init();

    const contentRange = sheet.getRange(2, 1, maxRow, sheet.getMaxColumns());
    contentRange.clear();

    sheet.getRange(2, 1, data.length, this.headerRow.length)
      .setValues(data);

    sheet.deleteRows(data.length + 2, maxRow - data.length);
    return this;
  }

  /**
   * Archives sheet to the given spreadsheet
   * @param {string} id - sheet it to archive to
   */
  archive(id) {
    const ss = SpreadsheetApp.openById(id);
    this.sheet.copyTo(ss);
    this.wipe();
    return this;
  }

  /**
   * Performs full text search
   * @param {object} find - options
   */
  fts({ q, regex, matchCell, matchCase }) {
    this._init();

    const finder = this.sheet.createTextFinder(q);
    finder.useRegularExpression(!!regex);
    finder.matchEntireCell(!!matchCell);
    finder.matchCase(!!matchCase);

    const ranges = finder.findAll();
    const rows = ranges.map(rng => {
      const rowNum = rng.getRow();
      const [row] = this.sheet.getRange(rowNum, 1, 1, this.COLUMN_COUNT).getValues();
      return this._getObject(row, rowNum - SheetDataAccess.ROW_INDEX_OFFSET);
    });

    return rows;
  }

  /**
   * Sorts the source sheet data by column
   * @param {string} column - column name to sort
   * @param {boolean} [asc] - ascending order
   */
  sort(column, asc) {
    this._init();
    const headers = this.headerRow();
    const index = headers.findIndex(column);

    if (index !== -1) {
      this.sheet.sort(index + 1, !!asc);
      this.clearCached();
    }

    return this;
  }

  /**
   * Updates a range in the sheet datasource with the record data
   * @param {object} record - record object
   * @param {any[]} recordValues - record array values
   * @param {number} columnCount - number of columns in range
   */
  _updateRow(record, recordValues, columnCount) {
    const range = this.sheet.getRange(record._key, 1, 1, columnCount);

    //last second check to make sure 2d array and sheet are still in sync for this object
    if (String(range.getValues()[0][this._pkColumnIndex]) !== String(recordValues[this._pkColumnIndex]))
      throw new Error(`Id at row ${record._key} does not match id of object for ${recordValues[this._pkColumnIndex]}`);

    range.setValues([recordValues]);
  }

  /**
   * Returns a usage report
   */
  inspect() {
    const sheet = this.sheet;
    const totalColumns = sheet.getMaxColumns();
    const totalRows = sheet.getMaxRows();
    const totalCells = totalColumns * totalRows;

    const report = {
      name: sheet.getName(),
      totalColumns,
      totalRows,
      totalCells,
      usagePercent: parseFloat((totalCells / SheetDataAccess.CELL_CAP).toFixed(2))
    };

    return report;
  }

}