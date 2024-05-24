<?php
require_once 'vendor/autoload.php';
require_once 'connect2.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Borders;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

$spreadsheet = new Spreadsheet();
$spreadsheet->setActiveSheetIndex(0);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Dữ liệu');

$sheet->getColumnDimension('A')->setWidth(15);
$sheet->getColumnDimension('B')->setWidth(30);
$sheet->getColumnDimension('C')->setWidth(30);
$sheet->getColumnDimension('D')->setWidth(30);
$sheet->getColumnDimension('E')->setWidth(30);
$sheet->getColumnDimension('F')->setWidth(30);
$rowCount = 1;
$sheet->setCellValue('A' . $rowCount, 'STT');
$sheet->setCellValue('B' . $rowCount, 'Họ và tên');
$sheet->setCellValue('C' . $rowCount, 'Lớp');
$sheet->setCellValue('D' . $rowCount, 'Điểm Toán');
$sheet->setCellValue('E' . $rowCount, 'Điểm Lý');
$sheet->setCellValue('F' . $rowCount, 'Điểm hóa');

$sql = "SELECT * FROM students";
$stmt = $conn->prepare($sql);
$stmt->execute();
// Đặt chế độ trả về mảng kết hợp
$stmt->setFetchMode(PDO::FETCH_ASSOC);
$result = $stmt->fetchAll();
foreach ($result as $row) {
    $rowCount++;
    $sheet->setCellValue('A' . $rowCount, $row['id']);
    $sheet->setCellValue('B' . $rowCount, $row['sbd']);
    $sheet->setCellValue('C' . $rowCount, $row['gt']);
    $sheet->setCellValue('D' . $rowCount, $row['toan']);
    $sheet->setCellValue('E' . $rowCount, $row['anh']);
    $sheet->setCellValue('F' . $rowCount, $row['van']);
}
$sheet->mergeCells('A1:B1');
$style = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN, // kiểu đường viền
            'color' => ['argb' => 'FF000000'], // màu sắc của đường viền
        ],
    ],
];

// Áp dụng style cho cell A1
$sheet->getStyle('A1:F' . $rowCount)->applyFromArray($style);


// Tạo một đối tượng Writer cho định dạng Xlsx (Excel)
$writer = new Xlsx($spreadsheet);

// Đặt header cho tệp Excel để trình duyệt hiểu
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="example.xlsx"');
header('Cache-Control: max-age=0');

// Ghi tệp Excel vào đầu ra (output) 
$writer->save('php://output');
