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
