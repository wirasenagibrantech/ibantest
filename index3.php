<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// --- Fungsi Format Tanggal Indonesia ---
function formatTanggalIndonesia($excelDate)
{
    // Jika bukan angka (misalnya sudah string), langsung kembalikan
    if (!is_numeric($excelDate)) return $excelDate;

    // Konversi serial number Excel ke timestamp
    $unixTime = ((int)$excelDate - 25569) * 86400;
    $date = new DateTime("@$unixTime");

    // Sesuaikan zona waktu
    $date->setTimezone(new DateTimeZone('Asia/Jakarta'));

    // Nama bulan Indonesia
    $bulanIndonesia = [
        1 => 'Januari', 'Februari', 'Maret', 'April',
        'Mei', 'Juni', 'Juli', 'Agustus',
        'September', 'Oktober', 'November', 'Desember'
    ];

    $d = $date->format('j');
    $m = $bulanIndonesia[(int)$date->format('n')];
    $y = $date->format('Y');

    return "$m $y";
}

// --- Load Excel ---
$spreadsheet = IOFactory::load("data.xlsx");
$sheet = $spreadsheet->getActiveSheet();

// Ambil data ---
$bulan = [];
$penjualan = [];

foreach ($sheet->getRowIterator(2) as $row) { // Mulai dari baris 2 (abaikan header)
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(false);

    $cells = [];
    foreach ($cellIterator as $cell) {
        $cells[] = $cell->getValue();
    }

    if ($cells[0] == null) continue;

    // Format tanggal
    $tanggalIndo = formatTanggalIndonesia($cells[0]);

    $bulan[] = $tanggalIndo;
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
