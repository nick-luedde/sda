/**
 * Class for handling data access to app data google sheet
 * Expects each collection to have an id property that is treated as the unique identifier for that record
 */
class SheetDataAccess implements ISheetDataAccess {

  /**
   * Static helper prop for the offset from the Sheet row and the eventual data array index
   */
  static get ROW_INDEX_OFFSET() {
    return 2;
  }

  /**
   * Creator function
   * @param {object} source - Spreadsheet source options
   * @param {string} [source.id] - Spreadsheet id
   * @param {Object} [source.ss] - Spreadsheet object
   * @param {Object} [options] - options object
   * @param {Schema} [options.schema] - optional Schema to apply to the datasource objects
   * @param {Object} [options.models] - optional schema models to apply to the datasource objects
   */
  static create({ id, ss }, { schema, models } = {}) {
    const spreadsheet = ss || SpreadsheetApp.openById(id);

    if (!schema !== !models)
      throw new Error('Missing schema or data models!');

    const collections = {};

    const sheets = spreadsheet.getSheets();
    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (name[0] !== '_') {
        const model = !!models ? models[name] : null;
        if (!model && !!models)
          throw new Error(`${name} has no schema model provided!`);

        collections[name] = SheetDataCollection.create(sheet, { schema, model });
      }
    });

    const defrag = () => Object.keys(collections).forEach(key => collections[key].defrag());

    const api = {
      collections,
      defrag
    };

    return api;
  }


  /**
   * maps an array of data to an object with headers of the row as property keys
   * @param {Array<Any>} row - row of data to map to object
   * @param {number} index - index of the object within the data array
   * @param {Array<string>} headers - array of header names in the order of appearance in sheet
   * @returns {Object} mapped object
   */
  static getRowAsObject(row, index, headers) {
    const obj = {
      _key: index + SheetDataAccess.ROW_INDEX_OFFSET
    };

    headers.forEach((header, index) => obj[header] = row[index]);
    return obj;
  };

}

/**
 * Class that manages read writes to a specific collection of data based on a sheet
 */
class SheetDataCollection {

  /**
   * @param {Sheet} sheet - sheet for the collection of data
   * @param {Object} [options] - collections options
   * @param {Schema} [options.schema] - schema to apply to the collection 
   * @param {Object} [options.model] - schema model to apply to the collection 
   */
  static create(sheet, { schema, model } = {}) {
    const hasSchema = !!schema;
    const hasModel = !!model;

    const context = {
      COLUMN_COUNT: null,
      ROW_COUNT: null,
    };

    const cache = {
      data: null,
      index: {},
      related: {},
      headerRow: null
    };

    /**
     * setup any props needed for data writing methods
     */
    const init = () => {
      context.COLUMN_COUNT = sheet.getLastColumn();
      context.ROW_COUNT = sheet.getLastRow();

      if (!cache.data)
        [cache.headerRow] = sheet.getRange(1, 1, 1, context.COLUMN_COUNT).getValues();
    };

    /**
     * Helper to clear cached data to force refreshes
     * @returns {SheetDataCollection} - this for chaining
     */
    const clearCached = () => {
      cache.data = null;
      cache.index = {};
      cache.related = {};

      return api;
    };

    /**
     * Gets a row as an object (from schema if defined)
     * @param {Array} row - data row array
     * @param {Number} index - index of the data in the dataset
     * @returns {Object} row data mapped to an object
     */
    const getObject = (row, index) => {
      const obj = SheetDataAccess.getRowAsObject(row, index, cache.headerRow);
      return hasModel ? schema.from(obj, model) : obj;
    };

    /**
     * Gets shallow copies of records to save, applies schema if exists
     * @param {Object[]} records - records to get saveable array
     * @returns {Object[]} new array of shallow copied/schema applied records
     */
    const getRecordsToSave = (records, { ignoreErrors } = {}) =>
      records.map(record =>
        hasModel ? schema.apply(record, model, { isNew: !record._key, ignoreErrors }) : { ...record }
      );

    /**
     * Gets records from schema or shallow copies if none
     * @param {Object[]} records - records to get from schema 
     */
    const getFromSchemaRecords = (records) => records.map(record =>
      hasModel ? schema.from(record, model) : { ...record }
    );

    /**
     * Helper function that will replace the top row of a sheet with headers from the provided obj
     * @param {Object} obj - object with headers to write
     * @returns {SheetDataCollection} - this for chaining
     */
    const writeHeadersFromObject = (obj) => {
      const firstRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
      firstRow.clear();

      const headers = Object.keys(obj);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);

      return api;
    };

    /**
     * Caches and returns a unique key map
     * @param {string} [key] - optional id
     * @returns {Object} index map
     */
    const index = (key = 'id') => {
      if (!cache.index[key]) {

        const data = data();

        const index = data.reduce((obj, record) => {
          obj[record[key]] = record;
          return obj;
        }, {});

        cache.index[key] = index;
      }

      return cache.index[key];
    };

    /**
     * Creates and returns a map of related items
     * @param {string} key - property key of related set to get
     * @returns {Object} related map
     */
    const related = (key) => {
      if (!cache.related[key]) {

        const data = data();

        const related = data.reduce((obj, record) => {
          if (!obj[record[key]])
            obj[record[key]] = [record]
          else
            obj[record[key]].push(record);

          return obj;
        }, {});

        cache.related[key] = related;
      }

      return cache.related[key];
    };

    /**
     * Throws error if prop value already exists in the dataset
     * @param {object} rec - record to check
     * @param {string} prop - property to enforce uniqueness
     * @returns {SheetDataAccess} this
     */
    const enforceUnique = (rec, prop) => {
      if (!rec._key) {
        const idx = index(prop);
        if (idx[rec[prop]] !== undefined)
          throw new Error(`${sheet.getName()} ${prop} prop value ${rec[prop]} already exists!`);
      } else {
        // is the best to just filter it out??
        const others = data().filter(oth => oth._key !== rec._key);
        const set = new Set(others.map(oth => oth[prop]));
        if (set.has(rec[prop]))
          throw new Error(`${sheet.getName()} ${prop} prop value ${rec[prop]} already exists!`);
      }
  
      return api;
    };

    /**
     * Handles retrieving and caching item data from sheet
     * @returns {Object[]} array of all item data
     */
    const data = () => {
      if (!cache.data) {
        init();
        const values = sheet.getDataRange().getValues();

        values.shift();
        cache.data = [];

        values.forEach((row, index) => {
          if (row[0] !== '')
            cache.data.push(getObject(row, index));
        });
      }

      return cache.data;
    };

    /**
     * Finds a record by a given key
     * @param {string} key - key of record to get
     * @param {string} [idx] - optional index to use, defaults to '_key'
     */
    const find = (key, idx = '_key') => index(idx)[key];

    /**
     * Updates a range in the sheet datasource with the record data
     * @param {Object} record - record object
     * @param {Array} recordValues - record array values
     * @param {Number} columnCount - number of columns in range
     */
    const updateRow = (record, recordValues, columnCount) => {
      const range = sheet.getRange(record._key, 1, 1, columnCount);

      //last second check to make sure 2d array and sheet are still in sync for this object
      if (String(range.getValues()[0][0]) !== String(recordValues[0]))
        throw new Error(`Id at row ${record._key} does not match id of object for ${recordValues[0]}`);

      range.setValues([recordValues]);
    };

    /**
     * Upserts one record (more concurrent safe)
     * @param {object} record - record to upsert
     * @param {{ bypassSchema: boolean }} [options] - options
     */
    const upsertOne = (record, { bypassSchema = false } = {}) => {
      if (!record)
        return null;
  
      const isNew = record._key === undefined || record._key === null;
      const [saved] = isNew ? [addOne(record, { bypassSchema })] : update([record], { bypassSchema });
      return saved;
    };

    /**
     * Saves record objects to sheet datasource
     * @param {Object[]} records - record objects to save
     * @returns {Object[]} records in their saved state
     */
    const upsert = (records, { bypassSchema = false } = {}) => {
      if (records.length === 0)
        return records;

      init();

      const schemaModels = !bypassSchema ? getRecordsToSave(records) : records;

      const updates = schemaModels.filter(rec => rec._key !== undefined && rec._key !== null);
      const adds = schemaModels.filter(rec => rec._key === undefined || rec._key === null);

      update(updates, { bypassSchema: true });
      add(adds, { bypassSchema: true })

      //clear cached data to force rebuild to account for changed/added records
      clearCached();

      return getFromSchemaRecords(schemaModels);
    };

    /**
     * Adds one record (safer than adding accross a range for collision)
     * @param {Object} record - record to add
     */
    const addOne = (record, { bypassSchema = false } = {}) => {
      //saves record data objects to the spreadsheet
      if (!record)
        return null;

      init();

      const recordToSave = !bypassSchema
        ? getRecordsToSave([record])[0]
        : record;

      const row = cache.headerRow.map(header => recordToSave[header]);
      sheet.appendRow(row);

      //clear cached data to force rebuild to account for changed/added records
      clearCached();

      //return objects to their from schema state
      //this seems like i should be doing this a different way....
      return getFromSchemaRecords([recordToSave])[0];
    };

    /**
     * adds the record models
     * @param {Object[]} records - records to add to the sheet datasource
     */
    const add = (records, { bypassSchema = false } = {}) => {
      //saves record data objects to the spreadsheet
      if (records.length === 0)
        return [];

      init();

      const recordsToSave = !bypassSchema
        ? getRecordsToSave(records)
        : records;

      const recordArrays = recordsToSave
        .map((record, index) => {
          record._key = context.ROW_COUNT + index + 1;
          return cache.headerRow.map(header => record[header]);
        });

      const range = sheet.getRange(context.ROW_COUNT + 1, 1, recordsToSave.length, context.COLUMN_COUNT);
      if (!!range.getValues()[0][0])
        throw new Error('Add transaction error, try again!');

      range.setValues(recordArrays);

      //clear cached data to force rebuild to account for changed/added records
      clearCached();

      //return objects to their from schema state
      //this seems like i should be doing this a different way....
      return getFromSchemaRecords(recordsToSave);
    };

    /**
     * updates the record models
     * @param {Object[]} records - records to update in the sheet datasource
     */
    const update = (records, { bypassSchema = false } = {}) => {
      if (records.length === 0)
        return [];

      init();

      const recordsToSave = !bypassSchema
        ? getRecordsToSave(records)
        : records;

      recordsToSave.forEach(record => {
        const recordValues = cache.headerRow.map(header => record[header]);
        updateRow(record, recordValues, context.COLUMN_COUNT);
      });

      //clear cached data to force rebuild to account for changed/added records
      clearCached();

      //return objects to their from schema state
      //this seems like i should be doing this a different way....
      return getFromSchemaRecords(recordsToSave);;
    };

    /**
     * Patches the provided patch props onto existing models 
     * (allows for targeted updates, which could help with multi-users so that entire records arent saved, just the individual changes are applied)
     * @param {Object[]} patches - list of patches to apply
     * @param {Object} options 
     */
    const patch = (patches, { bypassSchema = false } = {}) => {
      if (patches.length === 0)
        return [];

      init();

      const patchedRecords = patches.map(patch => {
        const existing = find(patch._key);
        if (!existing)
          throw new Error(`Could not patch record with key ${patch._key}. Key not found!`);

        return {
          ...existing,
          ...patch
        };
      });

      return update(patchedRecords, { bypassSchema });
    };


    /**
     * deletes record objects to sheet datasource
     */
    const del = (records) => {
      init();

      // find each record to remove...
      records.forEach(record => {
        const recordValues = cache.headerRow.map(header => record[header]);
        const range = sheet.getRange(record._key, 1, 1, context.COLUMN_COUNT);

        //last second check to make sure 2d array and sheet are still in sync for this object
        if (String(range.getValues()[0][0]) !== String(recordValues[0]))
          throw new Error(`Id at row ${record._key} does not match id of object for ${recordValues[0]}`);

        range.setValues([new Array(context.COLUMN_COUNT)]);
      });

      //clear cached data to force rebuild to account for deleted records
      clearCached();
    };

    /**
     * Performs batch update on the entire dataset for (meant for faster but more expensive updates)
     * @param {Object[]} records - batch data to apply
     * @param {Object} [options] - options
     */
    const batch = (records, { bypassSchema = false } = {}) => {
      if (records.length === 0)
        return records;

      init();

      const schemaModels = !bypassSchema ? getRecordsToSave(records) : records;

      const updates = schemaModels.filter(rec => rec._key !== undefined && rec._key !== null);
      const adds = schemaModels.filter(rec => rec._key === undefined || rec._key === null);

      // get data content without the header row;
      const data = sheet.getDataRange().getValues().slice(1);

      updates.forEach(rec =>
        data.splice(rec._key - SheetDataAccess.ROW_INDEX_OFFSET, 1, cache.headerRow.map(hdr => rec[hdr]))
      );
      data.push(...adds.map(rec => cache.headerRow.map(hdr => rec[hdr])));

      const lock = LockService.getScriptLock();
      lock.tryLock(1000 * 10);

      if (!lock.hasLock())
        throw new Error('Could not perform batch operation, please try again!');

      wipe();

      sheet.getRange(2, 1, data.length, cache.headerRow.length)
        .setValues(data);

      lock.releaseLock();

      clearCached();

      return getFromSchemaRecords(schemaModels);
    };

    /**
     * Clears all records from the sheet
     */
    const wipe = () => {
      const maxRow = sheet.getMaxRows();
      if (maxRow === 1)
        return;

      sheet.deleteRows(2, maxRow - 1);
    };

    /**
     * Removes non-data rows
     */
    const defrag = () => {
      const data = sheet.getDataRange().getValues().filter(row => row[0] !== '').slice(1);
      if (data.length === 0)
        return;

      const maxRow = sheet.getMaxRows();
      if (data.length + 1 === maxRow)
        return;

      init();

      const contentRange = sheet.getRange(2, 1, maxRow, sheet.getMaxColumns());
      contentRange.clear();

      sheet.getRange(2, 1, data.length, cache.headerRow.length)
        .setValues(data);

      sheet.deleteRows(data.length + 2, maxRow - data.length);

      return api;
    };


    //TODO: pick up here

    const api = {
      clearCached,
      writeHeadersFromObject,
      index,
      related,
      enforceUnique,
      data,
      find,
      upsertOne,
      upsert,
      addOne,
      add,
      update,
      patch,
      delete: del,
      batch,
      wipe,
      defrag
    };

    return api;
  }

}