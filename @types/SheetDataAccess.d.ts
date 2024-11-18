type ISheetDataAccessObject = { _key: string | number, id?: string };
type SheetDataAccessIndexCache<T> = { [key in keyof T]?: { [key: string]: T } };
type SheetDataAccessRelatedCache<T> = { [key in keyof T]?: { [key: string]: T[] } };
type SheetDataAccessSpreadsheetArg = {
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet;
  id?: undefined;
};
type SheetDataAccessIdArg = {
  ss?: undefined;
  id: string;
};
type SheetDataAccessCreateArgs = SheetDataAccessSpreadsheetArg | SheetDataAccessIdArg;
type SheetDataAccessCreateOptions<M> = {
  schemas?: { [key in keyof M]: any };
};
type SheetDataAccessBypassOption = {
  bypassSchema?: boolean;
};

type SheetDataAccessRecordObjectKey = {
  _key?: number | string | null;
};

interface ISheetDataAccess<M> {
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  /** Each Sheet in the Spreadsheet as a Data Collection interface */
  collections: { [key in keyof M]: ISheetDataCollection<M[key]> };
  /** Removes empty rows from all Sheets */
  defrag: () => ISheetDataAccess<M>;
  /** Clears all Sheets of data */
  wipe: () => ISheetDataAccess<M>;
  /** Copies Sheet data to another Sheet */
  archive: (folderId: string) => ISheetDataAccess<M>;
  /** Inspects the usage of all Sheets */
  inspect: () => {
    summary: {
      totalColumns: number;
      totalRows: number;
      totalCells: number;
      usagePercent: number;
    };
    breakdowns: {
      name: string;
      totalColumns: number;
      totalRows: number;
      totalCells: number;
      usagePercent: number;
    }[]
  }
}

interface ISheetDataCollection<Type extends SheetDataAccessRecordObjectKey> {
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  /** Sets the primary key column (defaults to 0) */
  pk: (index: number) => ISheetDataCollection<Type>;
  /** Counts the number of rows used for data */
  rowCount: () => number;
  /** Clears and cached data loaded into this instance */
  clearCached: () => ISheetDataCollection<Type>;
  /** Easy way to write headers to your Sheet from a given object structure (mostly a dev/config tool) */
  writeHeadersFromObject: (obj: Type) => ISheetDataCollection<Type>;
  /** Computes/retrieves an object map index for the given key value (useful for a lot of lookups by id or another unique property) */
  index: (key: keyof Type) => { [key: string]: Type };
  /** Computes/retrieves an object map index for all rows that share given key value (useful for a lot of lookups by id or another foreign key type property) */
  related: (key: keyof Type) => { [key: string]: Type[] };
  /** Verifies the value is unique to the column, throws if not */
  enforceUnique: (rec: Type, prop: keyof Type) => ISheetDataCollection<Type>;
  /** Get all rows of data */
  data: () => Type[];
  /** Stream chunks of data, useful for larger data-sets for memory, but does not necessarily perform any faster than data() */
  stream: (size: number) => Generator<Type[]>;
  /** Finds row mathing the given index value */
  find: (val: string, index: keyof Type) => Type | undefined;
  /** Gets a row by its row index */
  get: (key: number) => Type | null;
  /** Higher performance lookup in larger data sets for a single row matching a value */
  lookup: (val: string, key: keyof Omit<Type, "_key">) => Type | undefined;
  /** Adds or updates all records based on whether they are indicated as already existing (flagged as existing by having a _key property) */
  upsert: (records: Type[], options?: { bypassSchema: boolean }) => Type[];
  /** Adds or updates a single record based on whether they are indicated as already existing (flagged as existing by having a _key property) */
  upsertOne: (record: Type, options?: { bypassSchema: boolean }) => Type;
  /** Adds one record */
  addOne: (record: Type, options?: { bypassSchema: boolean }) => Type;
  /** Adds all records (WARNING: can conflict with other parallel add() calls to this Sheet and result in data loss, use LockService if needed) */
  add: (records: Type[], options?: { bypassSchema: boolean }) => Type[];
  /** Updates one record */
  updateOne: (record: Type, options?: { bypassSchema: boolean }) => Type;
  /** Updates all records */
  update: (records: Type[], options?: { bypassSchema: boolean }) => Type[];
  /** Patches changes to the given records */
  patch: (patches: Partial<Type>[], options?: { bypassSchema: boolean }) => Type[];
  /** Deletes all records */
  delete: (records: Partial<Type>[]) => void;
  /** Batch update which will perform much better for many updates (WARNING: can conflict with other parallel batch() calls to this Sheet and result in data loss, use LockService if needed) */
  batch: (patches: Type[], options?: { bypassSchema: boolean }) => Type[];
  /** Preflight check against any provided schema, usefull if performing multiple data changes that should either all succeed or fail */
  preflight: (records: Type[]) => {
    addOne: () => Type;
    add: () => Type[];
    update: () => Type[];
    batch: () => Type[];
    upsert: () => Type[];
    upsertOne: () => Type;
    record: () => Type;
    records: () => Type[];
  };
  /** Wipe the given number of rows from the end of the data */
  wipe: (rows?: number) => ISheetDataCollection<Type>;
  /** Clears all blank rows */
  defrag: () => ISheetDataCollection<Type>
  /** Copies the Sheet data to another Sheet */
  archive: (id: string) => ISheetDataCollection<Type>;
  /** Full text search on the Sheet data */
  fts: (find: { q: string, regex?: boolean, matchCell?: boolean, matchCase?: boolean }) => Type[];
  /** Sorts the Sheet data in place by the given column */
  sort: (column: keyof Type, asc?: boolean) => ISheetDataCollection<Type>;
  /** Inspects the usage of the Sheet */
  inspect: () => {
    name: string;
    totalColumns: number;
    totalRows: number;
    totalCells: number;
    usagePercent: number;
  }
}