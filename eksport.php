<?php
include('koneksi2.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'nama');
$sheet->setCellValue('B1', 'Jenis_kelamin');
$sheet->setCellValue('C1', 'Nisn');
$sheet->setCellValue('D1', 'Nik');
$sheet->setCellValue('E1', 'Tempat_lahir');
$sheet->setCellValue('F1', 'Tanggal_lahir');
$sheet->setCellValue('G1', 'NO_AKTA');
$sheet->setCellValue('H1', 'Agama');
$sheet->setCellValue('I1', 'Kewarganegaraan');
$sheet->setCellValue('J1', 'Berkebutuhan');
$sheet->setCellValue('K1', 'Alamat_rumah');
$sheet->setCellValue('L1', 'RT');
$sheet->setCellValue('M1', 'RW');
$sheet->setCellValue('N1', 'Nama_dusun');
$sheet->setCellValue('O1', 'Kelurahan');
$sheet->setCellValue('P1', 'Kecamatan');
$sheet->setCellValue('Q1', 'Kode_pos');
$sheet->setCellValue('R1', 'Lintang');
$sheet->setCellValue('S1', 'Bujur');
$sheet->setCellValue('T1', 'Tinggal');
$sheet->setCellValue('U1', 'Transportasi');
$sheet->setCellValue('V1', 'No_kks');
$sheet->setCellValue('W1', 'Anak_Ke');
$sheet->setCellValue('X1', 'PenerimaKps');
$sheet->setCellValue('Y1', 'No_kps');

$query = mysqli_query($conn,"select * from validation");
$i = 2;
while($row = mysqli_fetch_array($query))
{
	$sheet->setCellValue('A'.$i, $row['nama']);
	$sheet->setCellValue('B'.$i, $row['jkel']);
	$sheet->setCellValue('C'.$i, $row['nisn']);
	$sheet->setCellValue('D'.$i, $row['nik']);	
	$sheet->setCellValue('E'.$i, $row['temlahir']);
	$sheet->setCellValue('F'.$i, $row['tgllahir']);
	$sheet->setCellValue('G'.$i, $row['noregakte']);
	$sheet->setCellValue('H'.$i, $row['agama']);	
	$sheet->setCellValue('I'.$i, $row['kwn']);
	$sheet->setCellValue('J'.$i, $row['kebutuhan']);
	$sheet->setCellValue('K'.$i, $row['alamat']);
	$sheet->setCellValue('L'.$i, $row['rt']);	
	$sheet->setCellValue('M'.$i, $row['rw']);
	$sheet->setCellValue('N'.$i, $row['dusun']);
	$sheet->setCellValue('O'.$i, $row['kelurahan']);
	$sheet->setCellValue('P'.$i, $row['kecamatan']);	
	$sheet->setCellValue('Q'.$i, $row['kdpos']);
	$sheet->setCellValue('R'.$i, $row['lintang']);
	$sheet->setCellValue('S'.$i, $row['bujur']);
	$sheet->setCellValue('T'.$i, $row['temtinggal']);	
	$sheet->setCellValue('U'.$i, $row['transport']);
	$sheet->setCellValue('V'.$i, $row['nokks']);
	$sheet->setCellValue('W'.$i, $row['anakke']);
	$sheet->setCellValue('X'.$i, $row['kps']);
	$sheet->setCellValue('Y'.$i, $row['nokps']);	
	$i++;
}

$styleArray = [
			'borders' => [
				'allBorders' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
		];
$i = $i - 1;
$sheet->getStyle('A1:Y'.$i)->applyFromArray($styleArray);


$writer = new Xlsx($spreadsheet);
$writer->save('Report Data Formulir.xlsx');
echo "berhasil";
?>