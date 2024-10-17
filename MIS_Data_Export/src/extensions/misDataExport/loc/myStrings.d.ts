declare interface IMisDataExportCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'MisDataExportCommandSetStrings' {
  const strings: IMisDataExportCommandSetStrings;
  export = strings;
}
