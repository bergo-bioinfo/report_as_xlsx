<?php
/**
 * PLUGIN NAME: Report as xlsx
 * DESCRIPTION: Convert a report as xlsx file.
 * VERSION: 1.0
 * AUTHOR: Yec'han Laizet
 */

// Call the REDCap Connect file in the main "redcap" directory
require_once "../redcap_connect.php";

// OPTIONAL: Your custom PHP code goes here. You may use any constants/variables listed in redcap_info().

/*
 * Page Logic
 */
// OPTIONAL: Display the project header
    require_once APP_PATH_DOCROOT . 'ProjectGeneral/header.php';
?>

    <script type="text/javascript" src="http://oss.sheetjs.com/js-xlsx/xlsx.full.min.js"></script>
    <script type="text/javascript" src="https://fastcdn.org/FileSaver.js/1.1.20151003/FileSaver.min.js"></script>
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/blob-polyfill/2.0.20171115/Blob.min.js"></script>
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/4.3.6/papaparse.js"></script>

    <h1>Convertir un fichier report CSV en fichier XLSX</h1>
    <br>
    <br>

    <input type="file" id="csvFile" name="csvFile" />
    <br>
    <input id="lineReturn" type="checkbox" name="lineReturn" value="1" checked=true>Remettre les retours à la ligne dans les cellules<br>

    <br>
    <button onClick="downloadXlsx('Report', data);">XLSX</button>

    <script>
        var data;

        function handleFileSelect(evt) {
            var f = evt.target.files[0]; // FileList object
            //console.log(f);

            // Only process csv files.
            if (!f.type.match('text/csv') && !f.type.match('application/vnd.ms-excel') && !f.type.match('application/csv')) {
                alert("The file is not a csv file!");
                return false ;
            }
        
            var reader = new FileReader();

            // Closure to capture the file information.
            reader.onload = (function(theFile) {
                return function(event) {
                    var  lineReturn = document.getElementById("lineReturn");
                    if (lineReturn.checked === true) {
                        var csv_result = Papa.parse(event.target.result.replace(/  /g, "\n"));
                    } else {
                        var csv_result = Papa.parse(event.target.result);
                    }
                    data = csv_result.data;
                    downloadXlsx("Report", data, f.name.replace(/.csv$/, ".xlsx"));
                };
            })(f);

            // Read in the image file as a data URL.
            reader.readAsText(f);
        }

        document.getElementById('csvFile').addEventListener('change', handleFileSelect, false);


        function Workbook() {
            if(!(this instanceof Workbook)) return new Workbook();
            this.SheetNames = [];
            this.Sheets = {};
        }

        function s2ab(s) {
            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;
        }

        function downloadXlsx(sheetName, data, filename) {
            var wb = new Workbook(), ws = XLSX.utils.aoa_to_sheet(data);

            /* add worksheet to workbook */
            wb.SheetNames.push(sheetName);
            wb.Sheets[sheetName] = ws;
            var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
            saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), filename)
        }

    </script>


<?php
// OPTIONAL: Display the project footer
require_once APP_PATH_DOCROOT . 'ProjectGeneral/footer.php';
