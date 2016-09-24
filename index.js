flat = require('flatjson');
xlsx = require('node-xlsx');

module.exports = function (obj, opts) {

    var sheet = opts.sheetname || "Sheet 1";
    var delimiter = opts.delimiter || ".";
    var filter = opts.filter || function() { return true;}
    var headings = opts.headings;
    var pivot = opts.pivot;

    var header_labels = [];
    var input = [];
    if (obj instanceof Array) {
        input = obj;
    }
    else {
        input.push(obj);
    }

    //build headers
    var headers = []
    for ( var i = 0; i < input.length; i++ ) {
        var fo = flat(input[i], delimiter, filter);
        var keys = Object.keys(fo);
        for (var j = 0; j < keys.length; j++ ) {
            if ( headers.indexOf(keys[j]) < 0) {
                if ( headings && headings[keys[j]]) {
                    header_labels.push(headings[keys[j]]);
                    
                }
                else {
                    header_labels.push(keys[j])
                }
                headers.push(keys[j]);
            }
        }
    }

    var col_count = header_labels.length;
    var row_count = 1;
    var data = [];
    data.push(header_labels);
    for ( var i = 0; i < input.length; i++ ) {
        var actual_data = []
        var fo = flat(input[i], delimiter, filter);
        for (key in headers) {
            actual_data.push(fo[headers[key]]);
        }
        data.push(actual_data);
        row_count++;
    }
    if ( pivot ) {
        console.log("pivoting");
        console.log(data);
        // we have "row_count" arrays, each of "col_count" items.
        // we need to pivot this into "col_count" rows, each with "row_count" columns...
        var pivoted = [];
        for (var i = 0; i < col_count; i++ ) {
            var row = [];
            for (var j = 0; j < row_count; j++ ) {
                row.push(data[j][i]);
            }
            pivoted.push(row);
        }
        console.log(pivoted);
        return xlsx.build([{name: sheet, data: pivoted}]);
    }
    else {
        return xlsx.build([{name: sheet, data: data}]);
    }
   
}
