<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// Load Excel
$spreadsheet = IOFactory::load("data.xlsx");
$sheet = $spreadsheet->getActiveSheet();

// Ambil data
$bulan = [];
$penjualan = [];

foreach ($sheet->getRowIterator(2) as $row) { // Mulai dari baris 2
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(false);

    $cells = [];
    foreach ($cellIterator as $cell) {
        $cells[] = $cell->getValue();
    }

    if ($cells[0] == null) continue;

    $bulan[] = $cells[0];
    $penjualan[] = intval($cells[1]);
}
?>

<!DOCTYPE html>
<html>
<head>
    <title>Grafik dari Excel</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
</head>
<body>

    <h2>Grafik Penjualan</h2>

    <canvas id="chartExcel"></canvas>

    <script>
        const ctx = document.getElementById('chartExcel');

        new Chart(ctx, {
            type: 'bar',
            data: {
                labels: <?= json_encode($bulan) ?>,
                datasets: [{
                    label: 'Penjualan',
                    data: <?= json_encode($penjualan) ?>,
                    borderWidth: 1
                }]
            },
            options: {
                scales: {
                    y: { beginAtZero: true }
                }
            }
        });
    </script>

</body>
</html>
