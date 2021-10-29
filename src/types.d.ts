export type Row<T> = Array<T>;
export type Table<T> = Array<Row<T>>;

export interface Config {
  filename?: string;
  sheets?: Array<string>;
  isGermanFormat?: boolean;
  colNames?: Array<string>;
  out?: string;
  configOut?: string;
  configFilename?: string;
}

export interface ExtendedConfig extends Config {
  colNamesPerSheet?: { [sheetName: string]: Array<string> };
  outPerSheet?: { [sheetName: string]: string };
}
