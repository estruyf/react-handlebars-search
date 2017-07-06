export default function TypeofHelper(context, options) {
    if (typeof context === "object") {
        return Object.prototype.toString.call(context);
    } else {
        return typeof context;
    }
}
