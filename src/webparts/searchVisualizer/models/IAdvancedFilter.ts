export enum SearchFilter {
  contains = ":",
  containsStartsWith = ":...*",
  equals = "=",
  startsWith = "=...*",
  lessThan = "<",
  greaterThan = ">",
  notContains = "-VALUE",
  notEquals = "<>",
  notStartsWith = "-VALUE=...*"
}

export interface IAdvancedFilter {
  name: string;
  filter: SearchFilter;
  value: string;
  operator: string;
}
