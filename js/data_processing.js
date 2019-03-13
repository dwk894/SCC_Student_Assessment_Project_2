function processExcel(data) {
    var workbook = XLSX.read(data, { type : 'binary' });
    const targetSheetNames = ['core42list', 'pathways'];  // Use name to locate target sheets.
    var sheets = [null, null];  // sheets[0] is the 'Core42List' and sheets[1] is the 'Pathways'.
    for (var i = 0; i < workbook.SheetNames.length; i++) {
        if (workbook.SheetNames[i].toLowerCase() === targetSheetNames[0]) { sheets[0] = workbook.Sheets[workbook.SheetNames[i]]; }
        if (workbook.SheetNames[i].toLowerCase() === targetSheetNames[1]) { sheets[1] = workbook.Sheets[workbook.SheetNames[i]]; }
    }
    if (sheets[0] === null || sheets[1] === null) {
        alert('The spreadsheet does not have the required format.');
        return;
    }
    getOutput1(sheets);
}

function getOutput1(sheets) {
    /** @param sheets : [Excel spreadsheet, Excel spreadsheet]
        @return : void
    */
    var core42 = [], col = 'A', row = 1;
    while (sheets[0][col + row.toString()]) {
        if (sheets[0][col + row.toString()].v.toLowerCase().indexOf('core42') == -1) {
            core42.push(new Course(sheets[0][col + row.toString()].v.toUpperCase()));
        }
        row++;
    }
    var targetCols = new Set();
    for (var i = 1; i <= 24; i++) { targetCols.add('course' + i.toString()); }
    var colIndex = 0, headerRow = 1;
    while (sheets[1][toLetter(colIndex) + headerRow.toString()]) {
        if (targetCols.has(sheets[1][toLetter(colIndex) + headerRow.toString()].v.toLowerCase().replace(' ', ''))) {
            var j = 2;
            while (sheets[1]['A' + j.toString()]) {
                if (sheets[1][toLetter(colIndex) + j.toString()]) {
                    for (var k = 0; k < core42.length; k++) {
                        if (sheets[1][toLetter(colIndex) + j.toString()].v.toLowerCase().indexOf(core42[k].name.toLowerCase()) != -1) {
                            core42[k].freq[parseInt(sheets[1][toLetter(colIndex) + headerRow.toString()].v.substring(6)) - 1]++;
                        }
                    }
                }
                j++;
            }
        }
        colIndex++;
    }
    outputCourseSummary(core42);
}

function toLetter(x) {
    /** Convert the column in number to letter(s).
        @param x : Integer (zero-based column index)
        @return : String
    */
    var result = ''
    while (x != 0) {
        result = String.fromCharCode(x % 26 + 65) + result;
        x = Math.floor(x / 26);
    }
    if (result.length > 1) { result = String.fromCharCode(result.charCodeAt(0) - 1) + result.substring(1); }
    return result === '' ? 'A' : result;
}

function outputCourseSummary(courses) {
    /** @param courses : [Course]
        @return : void
    */
    var dataTable = [['Course Number', 'Overall', 'Semester 1', 'Semester 2', 'Semester 3', 'Semester 4']];
    var table = '<table style = "color: black; border-collapse: collapse; border: none; width: 100%;><tbody>"';
    table += '<tr><td style = "border:none; border: 2px solid blue; text-align: center; font-weight: bold; padding: 2pt 4pt;">Course Number</td>';
    table += '<td style = "border: none; border: 2px solid blue; text-align: center; font-weight: bold; padding: 2pt 4pt;">Appearances in Semester 1</td>';
    table += '<td style = "border: none; border: 2px solid blue; text-align: center; font-weight: bold; padding: 2pt 4pt;">Appearances in Semester 2</td>';
    table += '<td style = "border: none; border: 2px solid blue; text-align: center; font-weight: bold; padding: 2pt 4pt;">Appearances in Semester 3</td>';
    table += '<td style = "border: none; border: 2px solid blue; text-align: center; font-weight: bold; padding: 2pt 4pt;">Appearances in Semester 4</td>';
    table += '<td style = "border: none; border: 2px solid blue; text-align: center; font-weight: bold; padding: 2pt 4pt;">Overall Appearances</td></tr>';
    for (var i = 0; i < courses.length; i++) {
        var semester = courses[i].semesterSum();
        var overall = courses[i].overallSum();
        var currentRecord = [courses[i].name];
        currentRecord.push(overall);
        currentRecord = currentRecord.concat(semester);
        dataTable.push(currentRecord);
        table += '<tr><td style = "border: none; border: 2px solid blue; text-align: left; padding: 2pt 4pt;">' + courses[i].name + '</td>';
        table += '<td style = "border: none; border: 2px solid blue; text-align: center; padding: 2pt 4pt;">' + semester[0].toString() + '</td>';
        table += '<td style = "border: none; border: 2px solid blue; text-align: center; padding: 2pt 4pt;">' + semester[1].toString() + '</td>';
        table += '<td style = "border: none; border: 2px solid blue; text-align: center; padding: 2pt 4pt;">' + semester[2].toString() + '</td>';
        table += '<td style = "border: none; border: 2px solid blue; text-align: center; padding: 2pt 4pt;">' + semester[3].toString() + '</td>';
        table += '<td style = "border: none; border: 2px solid blue; text-align: center; padding: 2pt 4pt;">' + overall.toString() + '</td></tr>';
    }
    table += '</tbody></table>';
    drawColumnGraph(dataTable);
    $('#output').html(table);
}

function drawColumnGraph(dataArr) {
    google.charts.load('current', {'packages':['bar']});
    google.charts.setOnLoadCallback(drawChart);
    function drawChart() {
        var data = google.visualization.arrayToDataTable(dataArr);
        var options = {
            bars: 'horizontal',
            backgroundColor: 'transparent',
            chart: { title: 'Core42 Summary' },
            fontName: 'Tahoma',
            fontSize: '15',
            height: 9600
        };
        var chart = new google.charts.Bar(document.getElementById('chart'));
        chart.draw(data, google.charts.Bar.convertOptions(options));
    }
}