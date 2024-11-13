interface ISheetDataAccess {
  new(source: { id: string, ss: GoogleAppsScript.Spreadsheet.Spreadsheet }, options: { schemas?: { [key: string]: AsvContext } });
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  // collection: { [key: string]<Type>: SheetDataCollection<Type> };
  static ROW_INDEX_OFFSET: 2;
  static CELL_CAP: 10000000;
  static getRowAsObject: <Type>(row: any[], index: number, headers: [keyof Type]) => Type;
  defrag: () => ISheetDataAccess;
  wipe: () => ISheetDataAccess;
  archive: (folderId: string) => ISheetDataAccess;
  inspect: () => {
    summary: {
      totalColumns: number;
      totalRows: number;
      totalCells: number;
      usagePercent: string
    };
    breakdowns: {
      [key: string]: {
        name: string;
        totalColumns: number;
        totalRows: number;
        totalCells: number;
        usagePercent: string
      }
    }
  }
}

interface ISheetDataCollection<Type> {
  new(sheet: GoogleAppsScript.Spreadsheet.Sheet, options: { schema?: AsvContext });
  sheet: GoogleAppsScript.Spreadsheet.Sheet;
  schema?: AsvContext;
  hasModel: boolean;
  private _data: Type[];
  private _index: { [key: string]: Type[] };
  private _related: { [key: string]: Array<Type[]> };
  private _pkColumnIndex: number;

  private _init: () => void;
  private _values: () => Type[];
  private _getObject: (row: any[], index: number) => Type;
  private _getRecordsToSave: (records: Type[], oprions: { ignoreErrors?: boolean }) => any[];
  private _getFromSchemaRecords: (records: any[]) => Type[];
  private _updateRow: (record: Type, recordValues: any[], columCount: number) => void;
  pk: (index: number) => ISheetDataCollection<Type>;
  rowCount: () => number;
  clearCached: () => ISheetDataCollection<Type>;
  writeHeadersFromObject: (obj: Type) => ISheetDataCollection<Type>;
  index: (key: keyof Type = 'id') => { [key: string]: Type };
  related: (key: keyof Type) => { [key: string]: Type[] } | undefined;
  enforceUnique: (rec: Type, prop: keyof Type) => ISheetDataCollection<Type>;
  data: () => Type[];
  advanced: {
    data: () => Type[];
  };
  stream: (size: number) => Generator<Type[]>;
  find: (val: string, index: keyof Type) => Type | undefined;
  get: (key: number) => Type | null;
  lookup: (val: string, key: keyof Omit<Type, "_key">) => Type | undefined;
  upsert: (records: Type[], options: { bypassSchema: boolean = false }) => Type[];
  upsertOne: (record: Type, options: { bypassSchema: boolean = false }) => Type;
  add: (records: Type[], options: { bypassSchema: boolean = false }) => Type[];
  addOne: (record: Type, options: { bypassSchema: boolean = false }) => Type;
  update: (records: Type[], options: { bypassSchema: boolean = false }) => Type[];
  patch: (patches: Partial<Type>, options: { bypassSchema: boolean = false }) => Type[];
  delete: (records: Partial<Type>) => void;
  batch: (patches: Partial<Type>, options: { bypassSchema: boolean = false }) => Type[];
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
  wipe: (rows?: number) => ISheetDataCollection<Type>;
  defrag: () => ISheetDataCollection<Type>
  archive: (id: string) => ISheetDataCollection<Type>;
  fts: (find: { q?: string, regex?: string, matchCell?: boolean, matchCase?: boolean }) => Type[];
  sort: (column: keyof Type, asc?: boolean) => ISheetDataCollection<Type>;
  inspect: () => {
    name: string;
    totalColumns: number;
    totalRows: number;
    totalCells: number;
    usagePercent: string
  }
}