'use strict';

var gutil = require('gulp-util'),
    through = require('through2'),
    XLSX = require('xlsx');

var debug = false;

var log = function(msg) {
    if (debug) {
        console.log(msg);
    }
};

var getHeaders = function(worksheet, headRow, startColumn) {
    var headers = [];
    for (var key in worksheet) {

        var cell = worksheet[key],
            match = /([A-Z]+)(\d+)/.exec(key);

        if (!match) {
            continue;
        }

        var row = match[2]; // 1234
        var col = match[1]; // ABCD

        if (col >= startColumn) {
            if (row == headRow) {
                headers.push(cell.v);
                // console.log(JSON.stringify(headers));
            }
        }
    }
    return headers;
};

var createItem = function(headers) {
    var obj = {};
    headers.forEach(function(item) {
        obj[item] = null;
    });
    return obj;
};


/**
 * excel filename or workbook to json
 * @param fileName
 * @param headRow
 * @param valueRow
 * @returns {{}} json
 */
var toJson = function(fileName, headRow, valueRow, startColumn) {
    log('---------- toJSON() -------');
    var workbook;

    if (typeof fileName === 'string') {
        workbook = XLSX.readFile(fileName);
    } else {
        workbook = fileName;
    }

    var worksheet = workbook.Sheets[workbook.SheetNames[0]],
        json = [],
        namemap = {},
        curRow = 0,
        value,
        row,
        col,
        cell,
        match,
        headers = getHeaders(worksheet, headRow, startColumn);

    for (var key in worksheet) {
        if (worksheet.hasOwnProperty(key)) {
            cell = worksheet[key];
            match = /([A-Z]+)(\d+)/.exec(key);

            if (!match) {
                continue;
            }

            row = match[2]; // 1234
            col = match[1]; // ABCD

            // check if we have a new line
            if (row > curRow) {
                curRow = row;
                namemap = createItem(headers);
            }

            value = cell.v || null;

            if (col >= startColumn) {
                if (row < valueRow) {
                    //continue;
                } else {
                    log('-- key[' + col + ':' + row + '] => ' + value);

                    var colIndex = col.charCodeAt() - startColumn.charCodeAt();
                    namemap[headers[colIndex]] = value;

                    json[row - headRow - 1] = JSON.parse(JSON.stringify(namemap));
                }
            }
        }
    }
    return json;
};

module.exports = function(options) {
    options = options || {};
    debug = options.trace || false;

    return through.obj(function(file, enc, cb) {
        if (file.isNull()) {
            this.push(file);
            return cb();
        }

        if (file.isStream()) {
            this.emit('error', new gutil.PluginError(PLUGIN_NAME, 'Streaming not supported'));
            return cb();
        }

        var arr = [];
        for (var i = 0; i < file.contents.length; ++i) {
            arr[i] = String.fromCharCode(file.contents[i]);
        }

        var bString = arr.join("");

        /* Call XLSX */
        var workbook = XLSX.read(bString, {type: "binary", sheetStubs: true, cellHTML: true});
        file.contents = new Buffer(JSON.stringify(toJson(workbook, options.headRow || 1, options.valueRowStart || 2, options.startColumn || 'A')));

        log("Convert file: " + file.path);


        this.push(file);
        cb();
    });
};
