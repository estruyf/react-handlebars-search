
import { ISortingOption } from "./models/ISortingOption";
export interface ISearchVisualizerWebPartProps {
    title: string;
    query: string;
    maxResults: number;
    sorting: string;
    mpSorting: ISortingOption[];
    debug: boolean;
    external: string;
    scriptloading: boolean;
    duplicates: boolean;
    privateGroups: boolean;
    enableAudienceTargeting: boolean;
    audienceColumnMapping: string;
    audienceColumnAllValue: string;
    audienceBooleanOperator: string;
}
