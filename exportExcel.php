<?php
    $servername = "localhost";
    $username = "chopchop";
    $password = "Oxe9q7#5";
    $dbname = "chopchop";

    // Create connection
    $conn = mysqli_connect($servername, $username, $password, $dbname);
    // Check connection
    if (!$conn) {
        die("Connection failed: " . mysqli_connect_error());
    }

    /** Include PHPExcel */
    require_once 'classes/PHPExcel.php';

    $date_start = $_GET["start"];
    $date_end = $_GET["end"];

    $start_dateformat = DateTime::createFromFormat('Y-m-d', $date_start);
    $end_dateformat = DateTime::createFromFormat('Y-m-d', $date_end);

    $timestamp_start = $start_dateformat->getTimestamp();
    $timestamp_end = $end_dateformat->getTimestamp();

    $sql = "SELECT name, phone, email, sum FROM orders_n  WHERE `time` >= $timestamp_start AND `time` <= $timestamp_end";
    mysqli_set_charset($conn,"utf8");

    $result = mysqli_query($conn, $sql);

    
    // Create new PHPExcel object
    $objPHPExcel = new PHPExcel();
    $counter_row = 2;

    $objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A1', 'שם')
    ->setCellValue('B1', 'טלפון')
    ->setCellValue('C1', 'אימייל')
    ->setCellValue('D1', 'סכום הזמנה');

    while($row = $result->fetch_assoc()) {
        $a = "A".$counter_row;
        $b = "B".$counter_row;
        $c = "C".$counter_row;
        $d = "D".$counter_row;

        $objPHPExcel->setActiveSheetIndex(0)
        ->setCellValue($a, $row["name"])
        ->setCellValue($b, $row["phone"])
        ->setCellValue($c, $row["email"])
        ->setCellValue($d, $row["sum"]);

        $counter_row++;
    }

    // Rename worksheet
    $objPHPExcel->getActiveSheet()->setTitle('statics');

    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
    $objPHPExcel->setActiveSheetIndex(0);

    // Redirect output to a client’s web browser (Excel5)
    header('Content-Type: Document/vnd.ms-excel');
    header('Content-Disposition: attachment; filename="פרטי לקוחות.xls"');
    header('Cache-Control: max-age=0');
    // If you're serving to IE 9, then the following may be needed
    header('Cache-Control: max-age=1');

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save('php://output');
    exit;


    // Save Excel 2007 file
    $callStartTime = microtime(true);

    // $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    // $objWriter->save(str_replace('.php', '.xlsx', __FILE__));
    // $objWriter->save(str_replace('.php', '.xlsx', 'C:\users\user\downloads'));
    // $callEndTime = microtime(true);
    // $callTime = $callEndTime - $callStartTime;
?>