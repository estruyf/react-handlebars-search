export interface ISearchResults {
    PrimaryQueryResult: IPrimaryQueryResult;
}

export interface IPrimaryQueryResult {
    RelevantResults: IRelevantResults;
}

export interface IRelevantResults {
    Table: ITable;
    TotalRows: number;
    TotalRowsIncludingDuplicates: number;
}

export interface ITable {
    Rows: Array<ICells>;
}

export interface ICells {
    Cells: Array<ICellValue>;
}

export interface ICellValue {
    Key: string;
    Value: string;
    ValueType: string;
}

export interface ISearchResponse {
    results: any[];
    totalResults: number;
    totalResultsIncludingDuplicates: number;
    searchUrl: string;
}
