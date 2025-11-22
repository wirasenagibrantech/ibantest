<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;

// --- Format Tanggal Indonesia (short month) ---
function formatTanggalIndonesia($excelDate)
{
    if (!is_numeric($excelDate)) return $excelDate;

    $unix = ((int)$excelDate - 25569) * 86400;
    $date = new DateTime("@$unix");
    $date->setTimezone(new DateTimeZone("Asia/Jakarta"));

    $bulan = [
        1=>'Jan','Feb','Mar','Apr','Mei','Jun',
        'Jul','Agu','Sep','Okt','Nov','Des'
    ];

    return $date->format('j') . " " . $bulan[(int)$date->format('n')] . " " . $date->format('Y');
}

// --- LOAD EXCEL ---
$spreadsheet = IOFactory::load("data.xlsx");
$sheet = $spreadsheet->getActiveSheet();

$data = [];

// Ambil semua baris mulai Row 2
foreach ($sheet->getRowIterator(2) as $row) {
    $cells = [];
    foreach ($row->getCellIterator() as $cell) {
        $cells[] = $cell->getValue();
    }

    if ($cells[0] == null) continue;

    $excelDate = $cells[0];
    $penjualan = intval($cells[1]);

    // Convert Excel date → DateTime
    $unix = ((int)$excelDate - 25569) * 86400;
    $date = new DateTime("@$unix");

    $data[] = [
        "raw"       => $excelDate,
        "month"     => (int)$date->format("n"),
        "year"      => (int)$date->format("Y"),
        "tanggal"   => formatTanggalIndonesia($excelDate),
        "penjualan" => $penjualan
    ];
}

// --- SORT DATA BY YEAR THEN MONTH ---
usort($data, function($a, $b) {
    if ($a["year"] == $b["year"]) {
        return $a["month"] - $b["month"];
    }
    return $a["year"] - $b["year"];
});

// ChartJS dataset
$labels = array_column($data, "tanggal");
$values = array_column($data, "penjualan");
?>
<!DOCTYPE html>
<html>
<head>
    <title>Grafik & DataTable dari Excel</title>

    <!-- BOOTSTRAP 5.3.3 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">

    <!-- DataTables Bootstrap 5 CSS -->
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/dataTables.bootstrap5.min.css">
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.2/css/buttons.bootstrap5.min.css">

    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>

    <!-- Bootstrap JS -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>

    <!-- DataTables + Bootstrap 5 -->
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/dataTables.bootstrap5.min.js"></script>

    <!-- Buttons Export + Bootstrap -->
    <script src="https://cdn.datatables.net/buttons/2.4.2/js/dataTables.buttons.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.2/js/buttons.bootstrap5.min.js"></script>

    <!-- Excel export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>

    <!-- PDF export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js"></script>

    <!-- Export buttons -->
    <script src="https://cdn.datatables.net/buttons/2.4.2/js/buttons.html5.min.js"></script>
    <script src="https://cdn.datatables.net/buttons/2.4.2/js/buttons.print.min.js"></script>

    <!-- CHARTJS -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

</head>
<body class="bg-light p-4">

    <div class="container">

        <!-- CARD TABLE -->
        <div class="card shadow mb-4">
            <div class="card-header bg-primary text-white">
                <h4 class="mb-0">Data Penjualan</h4>
            </div>
            <div class="card-body">

                <table id="tabelData" class="table table-striped" style="width:100%">
                    <thead class="table-warning text-center">
                        <tr>
                            <th>Tanggal</th>
                            <th>Penjualan</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php foreach ($data as $row): ?>
                            <tr>
                                <td><?= $row["tanggal"] ?></td>
                                <td><?= $row["penjualan"] ?></td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>

            </div>
        </div>

        <!-- CARD GRAFIK -->
        <div class="card shadow">
            <div class="card-header bg-success text-white">
                <h4 class="mb-0">Grafik Penjualan</h4>
            </div>
            <div class="card-body">
                <canvas id="chartExcel" height="100"></canvas>
            </div>
        </div>

    </div>

    <script>
        $(document).ready(function() {
            $('#tabelData').DataTable({
                dom: 'Bfrtip',
                buttons: [
                    { extend: 'excelHtml5', className: 'btn btn-success', title: 'Data Penjualan' },
                    { extend: 'pdfHtml5', className: 'btn btn-danger', title: 'Data Penjualan', orientation: 'landscape', pageSize: 'A4' },
                    'copy', 'csv', 'print'
                    ]
            });
        });

// CHART
        new Chart(document.getElementById('chartExcel'), {
            type: 'line',
            data: {
                labels: <?= json_encode($labels) ?>,
                datasets: [{
                    label: 'Penjualan',
                    data: <?= json_encode($values) ?>,
                    borderWidth: 2,
                    tension: 0.3
                }]
            }
        });
    </script>

</body>
</html>
