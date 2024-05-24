<?php
require_once 'vendor/autoload.php';
require_once 'connect2.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Borders;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

$spreadsheet = new Spreadsheet();
$spreadsheet->setActiveSheetIndex(0);
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Bao_cao_hoc_sinh_o_lai');

$sheet->getColumnDimension('A')->setWidth(10);
$sheet->getColumnDimension('B')->setWidth(10);
$sheet->getColumnDimension('C')->setWidth(10);
$sheet->getColumnDimension('D')->setWidth(10);
$sheet->getColumnDimension('E')->setWidth(10);
$sheet->getColumnDimension('F')->setWidth(10);
$sheet->getColumnDimension('G')->setWidth(10);
$sheet->getColumnDimension('H')->setWidth(10);
$sheet->getColumnDimension('I')->setWidth(10);
$sheet->getColumnDimension('J')->setWidth(10);
$sheet->getColumnDimension('K')->setWidth(10);

$sheet->setCellValue('F2' , 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM');
$sheet->setCellValue('G3' , 'Độc Lập - Tự Do - Hạnh Phúc');
$sheet->getStyle('F2')->getFont()->setBold(true);
$sheet->getStyle('G3')->getFont()->setBold(true);
$sheet->setCellValue('B3' , 'TH&&TCS TRUNG TÚ');
$sheet->getStyle('B3')->getFont()->setBold(true);
$sheet->setCellValue('C5' , 'THÔNG KÊ KẾT QUẢ CUỐI NĂM');
$sheet->setCellValue('D6' , 'Năm Học 2023 - 2024');
$sheet->getStyle('C5')->getFont()->setBold(true);
$sheet->getStyle('D6')->getFont()->setBold(true);
$sheet->getStyle('D6')->getFont()->setItalic(true);
$sheet->getStyle('C5')->getFont()->getColor()->setARGB('#000000');
$sheet->getStyle('D6')->getFont()->getColor()->setARGB('#000000');
// căn ngang
$sheet->getStyle('D6')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
// căn dọc
//$sheet->getStyle('D6')->getAlignment()->setHorizontal(Alignment::VERTICAL_CENTER);

$sheet->setCellValue('A8','Khối' );
$sheet->setCellValue('B8','Lớp' );
$sheet->setCellValue('C8','SS' );
$sheet->setCellValue('D8','SS Thực Tế' );
$sheet->setCellValue('E8','LÊN LỚP' );
$sheet->setCellValue('G8','Ở LẠI' );
$sheet->mergeCells('E8:F8');
$sheet->mergeCells('G8:H8');
$sheet->getStyle('E8')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->getStyle('G8')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$sheet->getStyle('A8')->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
$sheet->getStyle('A8')->getBorders()->getAllBorders()->getColor()->setARGB(Color::COLOR_BLACK);

//$sheet->mergeCells('A1:B1');
$style = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN, // kiểu đường viền
            'color' => ['argb' => 'FF000000'], // màu sắc của đường viền
        ],
    ],
];

// Áp dụng style cho cell A1
//$sheet->getStyle('A1:F' . $rowCount)->applyFromArray($style);


// Tạo một đối tượng Writer cho định dạng Xlsx (Excel)
$writer = new Xlsx($spreadsheet);
// Đặt header cho tệp Excel để trình duyệt hiểu
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="example.xlsx"');
header('Cache-Control: max-age=0');

// Ghi tệp Excel vào đầu ra (output)
$writer->save('php://output');
