flat = require('flatjson');
xlsx = require('node-xlsx');

module.exports = function(obj, opts) {

    var sheet = opts.sheetname || "Sheet 1";
    var delimiter = opts.delimiter || ".";
    var filter = opts.filter || function() { return true; }
    var headings = opts.headings;
    var pivot = opts.pivot;
    var sort = opts.sort ? opts.sort : function(a, b) {
        if (a.value < b.value) return -1;
        if (a.value > b.value) return 1;
        return 0;
    }

    var input = [];
    if (obj instanceof Array) {
        input = obj;
    } else {
        input.push(obj);
    }

    //build headers
    var headers = []
    for (var i = 0; i < input.length; i++) {

        var fo = flat(input[i], delimiter, filter);
        var keys = Object.keys(fo);
        for (var j = 0; j < keys.length; j++) {
            if (headers.map(function(h) { return h.value }).indexOf(keys[j]) < 0) {
                var heading = {
                    value: keys[j]
                }
                if (headings && headings[keys[j]]) {
                    heading.label = headings[keys[j]];

                } else {
                    heading.label = keys[j];
                }
                headers.push(heading)
            }
        }
    }

    var col_count = headers.length;
    var row_count = 1;
    var data = [];
    headers.sort(sort);
    data.push(headers.map(function(h) { return h.label }));

    for (var i = 0; i < input.length; i++) {
        var actual_data = []
        var fo = flat(input[i], delimiter, filter);
        for (key in headers) {
            actual_data.push(fo[headers[key].value]);
        }
        data.push(actual_data);
        row_count++;
    }
    if (pivot) {
        // we have "row_count" arrays, each of "col_count" items.
        // we need to pivot this into "col_count" rows, each with "row_count" columns...
        var pivoted = [];
        for (var i = 0; i < col_count; i++) {
            var row = [];
            for (var j = 0; j < row_count; j++) {
                row.push(data[j][i]);
            }
            pivoted.push(row);
        }
        return xlsx.build([{ name: sheet, data: pivoted }]);
    } else {
        return xlsx.build([{ name: sheet, data: data }]);
    }

}