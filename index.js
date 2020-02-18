'use strict';

const readXlsxFile = require('read-excel-file/node');

var gutil = require('gulp-util'),
    through = require('through2'),
    fs = require('fs-extra'),
    filename = require('file-name')


function readXls(filePath) {
    let moduleObj = {}

    return readXlsxFile(filePath).then((rows) => {
        // `rows` is an array of rows
        // each row being an array of cells.
        let titleRow = rows[0];

        let titles = [];
        for (let i = 0; i < titleRow.length; i++) {
            titles.push(titleRow[i]);
            i > 0 && (moduleObj[titleRow[i]] = {});
        }
        for (let i = 1; i < rows.length; i++) {
            let row = rows[i];
            for (let j = 1; j < row.length; j++) {
                if (row[0]) {
                    moduleObj[titles[j]][row[0]] = row[j];
                }
            }
        }
        return moduleObj;
    });
}


module.exports = function(options) {
    options = options || {};

    return through.obj(function(file, enc, cb) {

        readXls(file.path).then(data => {
            let outputFileName = file.path.replace(/\.[^.]+$/, '') + '.json';
            fs.writeFileSync(outputFileName, JSON.stringify(data), 'utf-8');

            gutil.log('Output file: ' + gutil.colors.magenta(outputFileName));
            cb();
        })
        .catch((err) => {
            this.emit('error', new gutil.PluginError('parseXLS', err));
        });



        // let converted = readXls(file.path)
        // console.log(converted)
        // var isArr = converted instanceof Array;
        // console.log(isArr)
        // // file.contents = new Buffer(readXls(file.path))
        
        // console.log("Convert file: " + file.path);
        // this.push(file);
        // cb();
    });
};