$('#processData').on('click', function () {
    var excelFile = $('#fileUpload')[0],
        regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx|.xlsm)$/;  // Supported file extensions: .xls, .xlsx and .xlsm.
    if (regex.test(excelFile.value.toLowerCase())) {
        if (typeof(FileReader) != 'undefined') {
            var reader = new FileReader();
            if (reader.readAsBinaryString) {
                reader.onload = function (e) { processExcel(e.target.result); };
                reader.readAsBinaryString(excelFile.files[0]);
            } else { alert('You are using IE.  Please use Google Chrome.'); }
        } else { alert('The browser does not support HTML5.  Please use Google Chrome.'); }
    } else { alert('Please provide a valid MS Excel file.'); }
});