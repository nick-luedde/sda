type ISheetDataAccessObject = { _key: string | number, id?: string };
type SheetDataAccessIndexCache<T> = { [key in keyof T]?: { [key: string]: T } };
type SheetDataAccessRelatedCache<T> = { [key in keyof T]?: { [key: string]: T[] } };

interface ISheetDataAccess {
  spreadsheet: any;
  // collection: { [key: string]<Type>: SheetDataCollection<Type> };
  defrag: () => ISheetDataAccess;
  wipe: () => ISheetDataAccess;
  archive: (folderId: string) => ISheetDataAccess;
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

interface ISheetDataCollection<Type> {
  sheet: any;
  pk: (index: number) => ISheetDataCollection<Type>;
  rowCount: () => number;
  clearCached: () => ISheetDataCollection<Type>;
  writeHeadersFromObject: (obj: Type) => ISheetDataCollection<Type>;
  index: (key: keyof Type) => { [key: string]: Type };
  related: (key: keyof Type) => { [key: string]: Type[] } | undefined;
  enforceUnique: (rec: Type, prop: keyof Type) => ISheetDataCollection<Type>;
  data: () => Type[];
  stream: (size: number) => Generator<Type[]>;
  find: (val: string, index: keyof Type) => Type | undefined;
  get: (key: number) => Type | null;
  lookup: (val: string, key: keyof Omit<Type, "_key">) => Type | undefined;
  upsert: (records: Type[], options: { bypassSchema: boolean }) => Type[];
  upsertOne: (record: Type, options: { bypassSchema: boolean }) => Type;
  add: (records: Type[], options: { bypassSchema: boolean }) => Type[];
  addOne: (record: Type, options: { bypassSchema: boolean }) => Type;
  update: (records: Type[], options: { bypassSchema: boolean }) => Type[];
  patch: (patches: Partial<Type>, options: { bypassSchema: boolean }) => Type[];
  delete: (records: Partial<Type>) => void;
  batch: (patches: Partial<Type>, options: { bypassSchema: boolean }) => Type[];
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
  fts: (find: { q?: string, regex?: boolean, matchCell?: boolean, matchCase?: boolean }) => Type[];
  sort: (column: keyof Type, asc?: boolean) => ISheetDataCollection<Type>;
  inspect: () => {
    name: string;
    totalColumns: number;
    totalRows: number;
    totalCells: number;
    usagePercent: number;
  }
}