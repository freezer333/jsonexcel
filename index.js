flat = require('flatjson');
xlsx = require('node-xlsx');

module.exports = function (obj, sheetname, delimiter, filter, headings) {
    var sheet = sheetname || "Sheet 1";
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

    var data = [];
    data.push(header_labels);
    for ( var i = 0; i < input.length; i++ ) {
        var actual_data = []
        var fo = flat(input[i], delimiter, filter);
        for (key in headers) {
            actual_data.push(fo[headers[key]]);
        }
        data.push(actual_data);
    }
    var buffer = xlsx.build([{name: sheet, data: data}]);
    return buffer;
}
