        var fileInput = document.getElementById('fileInput');
        var c1 = document.getElementById('c1');

        fileInput.addEventListener('change', function(e) {
            var file = e.target.files[0];
            var reader = new FileReader();

            reader.onload = function(event) {
                var data = new Uint8Array(event.target.result);
                var workbook;

                if (file.name.endsWith('.xls')) {
                    workbook = XLS.read(data, { type: 'array' });
                } else {
                    workbook = XLSX.read(data, { type: 'array' });
                }

                var sheetName = workbook.SheetNames[0];
                var sheet = workbook.Sheets[sheetName];

                var range = {
                    s: { r: 15, c: 23 },
                    e: { r: 60, c: 23 }
                };

                var excelData;
                if (file.name.endsWith('.xls')) {
                    excelData = XLS.utils.sheet_to_json(sheet, { range });
                } else {
                    excelData = XLSX.utils.sheet_to_json(sheet, { range });
                }

                let htmlTable = '<table>';
                for (let i = 0; i < excelData.length; i++) {
                    htmlTable += '<tr>';
                    for (var key in excelData[i]) {
                        htmlTable += '<td style="border-right: hidden;border-left: hidden;text-align: center;background: #D9D9D9;">' + excelData[i][key] + '</td>';
                    }
                    htmlTable += '</tr>';
                }
                htmlTable += '</table>';
                c1.innerHTML = htmlTable;
            };

            reader.readAsArrayBuffer(file);
        });
