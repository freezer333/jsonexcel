flat = require('flatjson');
xlsx = require('node-xlsx');

module.exports = function (obj, sheetname, delimiter, filter) {
    var sheet = sheetname || "Sheet 1";
    
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
            if ( headers.indexOf(keys[j] < 0)) {
                headers.push(keys[j]);
            }
        }
    }

    var data = [];
    data.push(headers);
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