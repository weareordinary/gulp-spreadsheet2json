# gulp-spreadsheet2json
> Excel (XLSX/XLS/ODS) to json.


## Usage
First, install `gulp-spreadsheet2json` as a development dependency:

```shell
npm install --save-dev gulp-spreadsheet2json
```

Then, add it to your `gulpfile.js`:

```javascript
var xls2json = require('gulp-spreadsheet2json'),
    spreadsheets= [
        'config/**.xlsx',
        'config/**.xls',
        'config/**.ods',
    ];

gulp.task('copy', function() {
    gulp.src(spreadsheets)
        .pipe(xls2json({
            headRow: 1,
            valueRowStart: 3,
            startColumn: 'C',
            trace: true
        }))
        .pipe(gulp.dest('build'))
});
```


## API

### excel2json([options])

#### options.headRow
Type: `number`

Default: `1`

The row number of head. (Start from 1).

#### options.valueRowStart
Type: `number`

Default: `3`

The start row number of values. (Start from 1)

#### options.startColumn
Type: `number`

Default: `A`

The start column Char of values. (Start from A)

#### options.trace
Type: `Boolean`

Default: `false`

Whether to log each file path while convert success.

## License
MIT &copy; Chris(https://github.com/chrisbing)
