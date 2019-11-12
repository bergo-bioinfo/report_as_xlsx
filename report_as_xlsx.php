<?php

namespace Bergonie\ReportAsXlsx;
/**
 * PLUGIN NAME: Report as xlsx
 * DESCRIPTION: Convert a report as xlsx file.
 * VERSION: 1.0
 * AUTHOR: Yec'han Laizet
 */

// Call the REDCap Connect file in the main "redcap" directory

// OPTIONAL: Your custom PHP code goes here. You may use any constants/variables listed in redcap_info().

/*
 * Page Logic
 */
// OPTIONAL: Display the project header
    require_once APP_PATH_DOCROOT . 'ProjectGeneral/header.php';
?>

    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.15.1/xlsx.full.min.js"></script>

    <h1>Convertir un fichier report CSV en fichier XLSX</h1>
    <br>

    <input type="file" id="csvFile" name="csvFile" />
    <br>
    <button onClick="handleFile()">XLSX</button>
    <br>

    <script>
        function handleFile() {
            var e = $('#csvFile')[0];
            var files = e.files, f = files[0];
            var reader = new FileReader();
            reader.onload = function(e) {
                var data = new Uint8Array(e.target.result);
                var workbook = XLSX.read(data, {type: 'array'});
                XLSX.writeFile(workbook, f.name.substring(0,f.name.length - 4)+'.xlsx');

            };
            reader.readAsArrayBuffer(f);
    }
    </script>


<?php
// OPTIONAL: Display the project footer
require_once APP_PATH_DOCROOT . 'ProjectGeneral/footer.php';
