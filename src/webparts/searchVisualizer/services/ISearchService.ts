export interface ISearchResults {
	PrimaryQueryResult: IPrimaryQueryResult;
}

export interface IPrimaryQueryResult {
	RelevantResults: IRelevantResults;
}

export interface IRelevantResults {
    Table: ITable;
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
    searchUrl: string;
}
