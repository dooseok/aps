/* Republic of Korea locals for flatpickr */
var flatpickr = flatpickr || { l10ns: {} };
flatpickr.l10ns.ko = {};

flatpickr.l10ns.ko.weekdays = {
	shorthand: ["?", "?", "?", "?", "?", "?", "?"],
	longhand: ["???", "???", "???", "???", "???", "???", "???"]
};

flatpickr.l10ns.ko.months = {
	shorthand: ["1?", "2?", "3?", "4?", "5?", "6?", "7?", "8?", "9?", "10?", "11?", "12?"],
	longhand: ["1?", "2?", "3?", "4?", "5?", "6?", "7?", "8?", "9?", "10?", "11?", "12?"]
};

flatpickr.l10ns.ko.ordinal = function () {
	return "?";
};
if (typeof module !== "undefined") module.exports = flatpickr.l10ns;