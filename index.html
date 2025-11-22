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

    return "$d $m $y";
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
    <!-- BOOTSTRAP 5.3.3 CSS -->
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
<body>

    <h2>Grafik Penjualan</h2>

    <section>
        <div class="col-12 col-lg-4 mb-4">
          <div class="card shadow-sm">
            <div class="card">
                <canvas id="chartExcel"></canvas>

            </div>
        </div>
    </div>
</section>

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
