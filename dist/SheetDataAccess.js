"use strict";
/**
 * Class for handling data access to app data google sheet
 * Expects each collection to have an id property that is treated as the unique identifier for that record
 */
class SheetDataAccess {
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
     * Creator function
     * @param {object} source - Spreadsheet source options
     * @param {string} [source.id] - Spreadsheet id
     * @param {any} [source.ss] - Spreadsheet object
     * @param {object} [options] - options object
     * @param {any} [options.schema] - optional Schema to apply to the datasource objects
     * @param {any} [options.models] - optional schema models to apply to the datasource objects
     */
    static create({ id, ss }, { schemas } = {}) {
        const spreadsheet = ss || SpreadsheetApp.openById(id);
        const collections = {};
        const hasSchema = !!schemas;
        const sheets = spreadsheet.getSheets();
        sheets.forEach(sheet => {
            const sheetName = sheet.getName();
            if (String(sheetName)[0] !== '_') {
                let schema;
                if (hasSchema) {
                    schema = schemas[sheetName];
                    if (!schema)
                        throw new Error(`${String(sheetName)} has no schema model provided!`);
                }
                collections[sheetName] = SheetDataCollection.create(sheet, { schema });
            }
        });
        /**
         * Clears all empty rows from all collections
         */
        const defrag = () => {
            Object.keys(collections).forEach(key => collections[key].defrag());
            return api;
        };
        /**
         * Archives entire spreadsheet content
         */
        const wipe = () => {
            Object.values(collections).forEach(coll => coll.wipe());
            return api;
        };
        /**
         * Archives entire spreadsheet content
         * @param {string} folderId - id of sheet to archive to
         */
        const archive = (folderId) => {
            //@ts-ignore
            const folder = DriveApp.getFolderById(folderId);
            const copy = spreadsheet.copy(`${spreadsheet.getName()}_${new Date().toJSON()}`);
            //@ts-ignore
            const file = DriveApp.getFileById(copy.getId());
            file.moveTo(folder);
            return wipe();
        };
        /**
         * Returns a usage report
         */
        const inspect = () => {
            const breakdowns = Object.values(collections)
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
        };
        const api = {
            spreadsheet,
            collections,
            defrag,
            wipe,
            archive,
            inspect,
        };
        return api;
    }
    /**
     * maps an array of data to an object with headers of the row as property keys
     * @param {any[]} row - row of data to map to object
     * @param {number} index - index of the object within the data array
     * @param {string[]} headers - array of header names in the order of appearance in sheet
     */
    static getRowAsObject(row, index, headers) {
        const obj = {
            _key: index + SheetDataAccess.ROW_INDEX_OFFSET
        };
        headers.forEach((header, index) => obj[header] = row[index]);
        return obj;
    }
    ;
}
/**
 * Class that manages read writes to a specific collection of data based on a sheet
 */
class SheetDataCollection {
    /**
     * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - sheet for the collection of data
     * @param {object} [options] - collections options
     * @param {any} [options.schema] - schema to apply to the collection
     * @param {object} [options.model] - schema model to apply to the collection
     */
    static create(sheet, { schema, model } = {}) {
        const hasModel = !!model;
        const context = {
            COLUMN_COUNT: 0,
            ROW_COUNT: 0,
            pkColumnIndex: 0,
        };
        const cache = {
            data: null,
            index: {},
            related: {},
            headerRow: []
        };
        const headerRow = () => {
            const [headers] = sheet.getRange(1, 1, 1, context.COLUMN_COUNT).getValues();
            return headers;
        };
        /**
         * setup any props needed for data writing methods
         */
        const init = () => {
            context.COLUMN_COUNT = sheet.getLastColumn();
            context.ROW_COUNT = rowCount();
            if (!cache.data)
                [cache.headerRow] = headerRow();
        };
        const rowCount = () => sheet.getLastRow();
        /**
         * Helper to clear cached data to force refreshes
         */
        const clearCached = () => {
            cache.data = null;
            cache.index = {};
            cache.related = {};
            return api;
        };
        /**
         * Gets a row as an object (from schema if defined)
         * @param {T[]} row - data row array
         * @param {number} index - index of the data in the dataset
         * @returns {T} row data mapped to an object
         */
        const getObject = (row, index) => {
            const obj = SheetDataAccess.getRowAsObject(row, index, cache.headerRow);
            return hasModel ? schema.from(obj, model) : obj;
        };
        /**
         * Gets shallow copies of records to save, applies schema if exists
         * @param {T[]} records - records to get saveable array
         * @returns {T[]} new array of shallow copied/schema applied records
         */
        const getRecordsToSave = (records, { ignoreErrors } = {}) => !hasModel
            ? records.map(rec => ({ ...rec }))
            : records.map((rec) => schema.exec(rec, { isNew: !rec._key, throwError: !ignoreErrors }));
        /**
         * Gets records from schema or shallow copies if none
         * @param {T[]} records - records to get from schema
         */
        const getFromSchemaRecords = (records) => !hasModel
            ? records.map(rec => ({ ...rec }))
            : records.map(schema.parse);
        /**
         * Helper function that will replace the top row of a sheet with headers from the provided obj
         * @param {T} obj - object with headers to write
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
         * Sets the index of the pk column (only necessary if not 0)
         * @param {number} i - col index of pk field
         */
        const pk = (i) => {
            context.pkColumnIndex = i;
            return api;
        };
        /**
         * Caches and returns a unique key map
         * @param {keyof T} [key] - optional id
         */
        const index = (key) => {
            const idx = cache.index[key];
            if (idx)
                return idx;
            const vals = data();
            const createdIdx = vals.reduce((obj, record) => {
                obj[String(record[key])] = record;
                return obj;
            }, {});
            cache.index[key] = createdIdx;
            return createdIdx;
        };
        /**
         * Creates and returns a map of related items
         * @param {keyof T} key - property key of related set to get
         */
        const related = (key) => {
            const idx = cache.related[key];
            if (idx)
                return idx;
            const vals = data();
            const createdIdx = vals.reduce((obj, record) => {
                const idxKey = String(record[key]);
                if (!obj[idxKey])
                    obj[idxKey] = [record];
                else
                    obj[idxKey].push(record);
                return obj;
            }, {});
            cache.related[key] = createdIdx;
            return createdIdx;
        };
        /**
         * Throws error if prop value already exists in the dataset
         * @param {T} rec - record to check
         * @param {keyof T} prop - property to enforce uniqueness
         */
        const enforceUnique = (rec, prop) => {
            if (!rec._key) {
                const idx = index(prop);
                if (idx[String(rec[prop])] !== undefined)
                    throw new Error(`${sheet.getName()} ${String(prop)} prop value ${rec[prop]} already exists!`);
            }
            else {
                // is the best to just filter it out??
                const others = data().filter(oth => oth._key !== rec._key);
                const set = new Set(others.map(oth => oth[prop]));
                if (set.has(rec[prop]))
                    throw new Error(`${sheet.getName()} ${String(prop)} prop value ${rec[prop]} already exists!`);
            }
            return api;
        };
        /**
         * Handles retrieving and caching item data from sheet
         */
        const data = () => {
            if (!cache.data) {
                init();
                const values = sheet.getDataRange().getValues();
                values.shift();
                cache.data = [];
                values.forEach((row, index) => {
                    if (row[0] !== '' && cache.data)
                        cache.data.push(getObject(row, index));
                });
            }
            return cache.data;
        };
        /**
         * Streams data in chunks...
         */
        const stream = function* (size) {
            const CHUNK_SIZE = size || 5000;
            let i = 0;
            init();
            const rows = rowCount();
            const columns = context.COLUMN_COUNT;
            const chunks = Math.ceil(rows / CHUNK_SIZE);
            while (i < chunks) {
                const startRow = i * CHUNK_SIZE + SheetDataAccess.ROW_INDEX_OFFSET;
                const rowsToGet = Math.min(rows - startRow, CHUNK_SIZE);
                const values = sheet.getSheetValues(startRow, 1, rowsToGet, columns);
                const data = [];
                values.forEach((row, index) => {
                    if (row[context.pkColumnIndex] !== '')
                        data.push(getObject(row, startRow - SheetDataAccess.ROW_INDEX_OFFSET + index));
                });
                i++;
                yield data;
            }
        };
        /**
         * Finds a record by a given key
         * @param {string} key - key of record to get
         * @param {keyof T} [idx] - optional index to use, defaults to '_key'
         */
        const find = (key, idx = '_key') => index(idx)[key];
        /**
         * Gets a row by key (row number)
         * @param {string | number} key
         */
        const get = (key) => {
            init();
            const keynum = Number(key);
            const [row] = sheet.getRange(keynum, 1, 1, context.COLUMN_COUNT).getValues();
            if (row[context.pkColumnIndex] !== '')
                return getObject(row, keynum - SheetDataAccess.ROW_INDEX_OFFSET);
            else
                return null;
        };
        /**
         * Performs an efficient lookup for a single record by value
         * @param {any} val
         * @param {keyof T} key
         */
        const lookup = (val, key = 'id') => {
            if (data !== null) {
                return find(val, key);
            }
            else {
                return fts({ q: val, matchCell: true }).find(r => r[key] === val);
            }
        };
        /**
         * Performs full text search
         * @param {object} find - options
         */
        const fts = ({ q, regex, matchCell, matchCase }) => {
            init();
            const finder = sheet.createTextFinder(q);
            finder.useRegularExpression(!!regex);
            finder.matchEntireCell(!!matchCell);
            finder.matchCase(!!matchCase);
            const ranges = finder.findAll();
            const rows = ranges.map(rng => {
                const rowNum = rng.getRow();
                const [row] = sheet.getRange(rowNum, 1, 1, context.COLUMN_COUNT).getValues();
                return getObject(row, rowNum - SheetDataAccess.ROW_INDEX_OFFSET);
            });
            return rows;
        };
        /**
         * Updates a range in the sheet datasource with the record data
         * @param {T} record - record object
         * @param {Array} recordValues - record array values
         * @param {Number} columnCount - number of columns in range
         */
        const updateRow = (record, recordValues, columnCount) => {
            const range = sheet.getRange(Number(record._key), 1, 1, columnCount);
            //last second check to make sure 2d array and sheet are still in sync for this object
            if (String(range.getValues()[0][0]) !== String(recordValues[0]))
                throw new Error(`Id at row ${record._key} does not match id of object for ${recordValues[0]}`);
            range.setValues([recordValues]);
        };
        /**
         * Upserts one record (more concurrent safe)
         * @param {T} record - record to upsert
         * @param {{ bypassSchema: boolean }} [options] - options
         */
        const upsertOne = (record, { bypassSchema = false } = {}) => {
            const isNew = record._key === undefined || record._key === null;
            const [saved] = isNew ? [addOne(record, { bypassSchema })] : update([record], { bypassSchema });
            return saved;
        };
        /**
         * Saves record objects to sheet datasource
         * @param {T[]} records - record objects to save
         * @returns {T[]} records in their saved state
         */
        const upsert = (records, { bypassSchema = false } = {}) => {
            if (records.length === 0)
                return records;
            init();
            const schemaModels = !bypassSchema ? getRecordsToSave(records) : records;
            const updates = schemaModels.filter(rec => rec._key !== undefined && rec._key !== null);
            const adds = schemaModels.filter(rec => rec._key === undefined || rec._key === null);
            update(updates, { bypassSchema: true });
            add(adds, { bypassSchema: true });
            //clear cached data to force rebuild to account for changed/added records
            clearCached();
            return getFromSchemaRecords(schemaModels);
        };
        /**
         * Adds one record (safer than adding accross a range for collision)
         * @param {T} record - record to add
         */
        const addOne = (record, { bypassSchema = false } = {}) => {
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
         * @param {T[]} records - records to add to the sheet datasource
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
         * @param {T[]} records - records to update in the sheet datasource
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
            return getFromSchemaRecords(recordsToSave);
            ;
        };
        /**
         * Patches the provided patch props onto existing models
         * (allows for targeted updates, which could help with multi-users so that entire records arent saved, just the individual changes are applied)
         * @param {T[]} patches - list of patches to apply
         * @param {SheetDataAccessBypassOption} options
         */
        const patch = (patches, { bypassSchema = false } = {}) => {
            if (patches.length === 0)
                return [];
            init();
            const patchedRecords = patches.map(patch => {
                const existing = find(String(patch._key || -1));
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
                const range = sheet.getRange(Number(record._key), 1, 1, context.COLUMN_COUNT);
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
         * @param {T[]} records - batch data to apply
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
            updates.forEach(rec => data.splice(rec._key - SheetDataAccess.ROW_INDEX_OFFSET, 1, cache.headerRow.map(hdr => rec[hdr])));
            data.push(...adds.map(rec => cache.headerRow.map(hdr => rec[hdr])));
            //@ts-ignore
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
         * Performs a preflight validation of all records prior to saving
         *  This can be especially usefull with combined data actions that you want to be more transactional (all pass or all fail)
         * returns update/add methods prepared with the preflight records
         * Only allows a single call of a transaction method (will error if one is called again)
         * @param {T | T[]} records - records to prelight validate
         */
        const preflight = (records) => {
            const arrayOfRecords = Array.isArray(records) ? records : [records];
            const recordsToSave = getRecordsToSave(arrayOfRecords);
            const bypassSchema = true;
            let transacted = false;
            const transact = (fn) => {
                if (transacted)
                    throw new Error('Preflight transaction already complete!');
                const result = fn();
                transacted = true;
                return result;
            };
            return {
                addOne: () => transact(() => addOne(recordsToSave[0], { bypassSchema })),
                add: () => transact(() => add(recordsToSave, { bypassSchema })),
                update: () => transact(() => update(recordsToSave, { bypassSchema })),
                batch: () => transact(() => batch(recordsToSave, { bypassSchema })),
                upsert: () => transact(() => upsert(recordsToSave, { bypassSchema })),
                upsertOne: () => transact(() => upsertOne(recordsToSave[0], { bypassSchema })),
                record: () => getFromSchemaRecords(recordsToSave)[0],
                records: () => getFromSchemaRecords(recordsToSave),
            };
        };
        /**
         * Sorts the source sheet data by column
         * @param {keyof T} column - column name to sort
         * @param {boolean} [asc] - ascending order
         */
        const sort = (column, asc) => {
            init();
            const headers = headerRow();
            const index = headers.indexOf(column);
            if (index !== -1) {
                sheet.sort(index + 1, !!asc);
                clearCached();
            }
            return api;
        };
        /**
         * Returns a usage report
         */
        const inspect = () => {
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
        };
        /**
         * Clears all records from the sheet
         */
        const wipe = () => {
            const maxRow = sheet.getMaxRows();
            if (maxRow === 1)
                return api;
            ;
            sheet.deleteRows(2, maxRow - 1);
            return api;
        };
        /**
         * Removes non-data rows
         */
        const defrag = () => {
            const data = sheet.getDataRange().getValues().filter(row => row[0] !== '').slice(1);
            if (data.length === 0)
                return api;
            const maxRow = sheet.getMaxRows();
            if (data.length + 1 === maxRow)
                return api;
            init();
            const contentRange = sheet.getRange(2, 1, maxRow, sheet.getMaxColumns());
            contentRange.clear();
            sheet.getRange(2, 1, data.length, cache.headerRow.length)
                .setValues(data);
            sheet.deleteRows(data.length + 2, maxRow - data.length);
            return api;
        };
        /**
         * Archives sheet to the given spreadsheet
         * @param {string} id - sheet it to archive to
         */
        const archive = (id) => {
            //@ts-ignore
            const ss = SpreadsheetApp.openById(id);
            sheet.copyTo(ss);
            wipe();
            return api;
        };
        const api = {
            sheet,
            rowCount,
            clearCached,
            writeHeadersFromObject,
            pk,
            index,
            related,
            enforceUnique,
            data,
            stream,
            find,
            get,
            lookup,
            fts,
            upsertOne,
            upsert,
            addOne,
            add,
            update,
            patch,
            delete: del,
            batch,
            preflight,
            sort,
            inspect,
            wipe,
            defrag,
            archive,
        };
        return api;
    }
}
