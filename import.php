<?php
require_once 'vendor/autoload.php';
require_once 'connect2.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

if (isset($_POST)) {
    $inputFileName = $_FILES['file']['tmp_name'];
    if (!empty($inputFileName)) {
        // đọc file
        $spreadsheet = IOFactory::load($inputFileName);
        // đọc dữ liệu ở sheet chỉ định
        $sheet = $spreadsheet->getActiveSheet('DLDiemthi');
        // lấy số hàng và số cột
        $highestRow = $sheet->getHighestRow();
        // lấy số cột dạng chữ
        $highestColumn = $sheet->getHighestColumn();
        // chuyển dạng chữ của cột về số
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
        // chuyển data từ file về dạng array với index là cột tương ứng
        // (dạng chữ , ko truyền tham số thì index là số)
        $data = $sheet->toArray('null', true, true, true);
        // vòng for bắt đầu = 2 là vì phần tử đầu tiên trong $data có index = 1 nhưng phần này chính là tiêu đề nên bỏ qua
        for ($i = 2; $i <= $highestColumnIndex; $i++) {
            $sbd = $data[$i]['A'];
            $gt = $data[$i]['B'];
            $toan = $data[$i]['C'];
            $ly = $data[$i]['D'];
            $hoa = $data[$i]['E'];
            $sql = "INSERT INTO students (sbd,gt,toan,anh,van)
        VALUES ('$sbd', '$gt', '$toan' , '$ly' , '$hoa')";
            if($conn->query($sql)){
                $check = 1;
            }
        }
        if($check){
            echo 'Thành công';
        }
    }
}