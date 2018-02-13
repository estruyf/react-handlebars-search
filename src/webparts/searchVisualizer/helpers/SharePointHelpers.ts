import { SPUser } from './../components/ISearchVisualizerProps';
// SharePoint helper to split SPUserField (?multiple) into a string

// The template provide the property which will be returned
export const splitSPUser = (userFieldValue, propertyRequested) => {
    if (userFieldValue == null)
        return null;

    const retValue: string[] = [];
    let userFieldValueArray = userFieldValue.split(';').forEach(user => {
        let userValues = user.split(' | ');
        let spuser: SPUser = {
            displayName: userValues[1],
            email: userValues[0]
        };
        retValue.push(spuser[propertyRequested]);
    });

    return retValue.join(', ');
};

// SharePoint helper to split the displaynames of for example the Author field (user1;user2...)
export const splitDisplayNames = (displayNames) => {
    if (displayNames == null && displayNames.indexOf(';') == -1) {
        return null;
    }

    return displayNames.split(';').join(", ");
};

// SharePoint helper to split the taxonomy name
export const splitSPTaxonomy = (taxonomyFieldValue) => {
    if (taxonomyFieldValue == null)
        return null;

    const retValue: string[] = [];

    let taxonomyFieldValueArray = taxonomyFieldValue.split(';GP0').forEach(taxonomy => {
        let taxonomyValues = taxonomy.split('|')[3];
        let termsValues = taxonomyValues.slice(0, taxonomyValues.lastIndexOf(';GTSet'));
        retValue.push(termsValues);
    });
    return retValue.join(', ');
};
