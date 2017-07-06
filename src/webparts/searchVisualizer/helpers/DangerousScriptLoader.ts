import { SPComponentLoader } from '@microsoft/sp-loader';

// Finds and executes scripts in a newly added element's body.
// Needed since innerHTML does not run scripts.
//
// Argument element is an element in the dom.
export default function executeScript(element: HTMLElement) {
    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    (<any>window).ScriptGlobal = {};

    function nodeName(elem, name) {
        return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
    }

    function evalScript(elem) {
        var data = (elem.text || elem.textContent || elem.innerHTML || ""),
            head = document.getElementsByTagName("head")[0] ||
                document.documentElement,
            script = document.createElement("script");

        script.type = "text/javascript";
        if (elem.src && elem.src.length > 0) {
            return;
        }
        if (elem.onload && elem.onload.length > 0) {
            script.onload = elem.onload;
        }

        try {
            // doesn't work on ie...
            script.appendChild(document.createTextNode(data));
        } catch (e) {
            // IE has funky script nodes
            script.text = data;
        }

        head.insertBefore(script, head.firstChild);
        head.removeChild(script);
    }

    // main section of function
    var scripts = [],
        script,
        children_nodes = element.childNodes,
        child,
        i;

    for (i = 0; children_nodes[i]; i++) {
        child = children_nodes[i];
        if (nodeName(child, "script") && (!child.type || child.type.toLowerCase() === "text/javascript")) {
            scripts.push(child);
        }
    }

    const urls = [];
    const onLoads = [];
    for (i = 0; scripts[i]; i++) {
        script = scripts[i];
        if (script.src && script.src.length > 0) {
            urls.push(script.src);
        }
        if (script.onload && script.onload.length > 0) {
            onLoads.push(script.onload);
        }
    }

    // Execute promises in sequentially - https://hackernoon.com/functional-javascript-resolving-promises-sequentially-7aac18c4431e
    // Use "ScriptGlobal" as the global namein case script is AMD/UMD
    const allFuncs = urls.map(url => () => SPComponentLoader.loadScript(url, { globalExportsName: "ScriptGlobal" }));

    const promiseSerial = funcs =>
        funcs.reduce((promise, func) =>
            promise.then(result => func().then(Array.prototype.concat.bind(result))),
            Promise.resolve([]));

    // execute Promises in serial
    promiseSerial(allFuncs)
        .then(() => {
            // execute any onload people have added
            for (i = 0; onLoads[i]; i++) {
                onLoads[i]();
            }
            // execute script blocks
            for (i = 0; scripts[i]; i++) {
                script = scripts[i];
                if (script.parentNode) { script.parentNode.removeChild(script); }
                evalScript(scripts[i]);
            }
        }).catch(console.error);
}
