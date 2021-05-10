<?php
include('koneksi.php'); //memasukan file koneksi.php
require 'vendor/autoload.php'; //merequire file autoload.php
use PhpOffice\PhpSpreadsheet\Spreadsheet;  
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;  
 
$spreadsheet = new Spreadsheet(); //membuat objek dengan nama $spreadsheet menggunakan class spreadsheet
$sheet = $spreadsheet->getActiveSheet(); //membuat variabel $sheet untuk activesheet di file excel
$sheet->setCellValue('A1', 'No');  //isi cell A1 pada excel dengan teks No
$sheet->setCellValue('B1', 'Nama'); //isi cell B1 pada excel dengan teks Nama
$sheet->setCellValue('C1', 'Kelas'); //isi cell C1 pada excel dengan teks Kelas
$sheet->setCellValue('D1', 'Alamat'); //isi cell D1 pada excel dengan teks Alamat
 
$query = mysqli_query($koneksi,"select * from tb_siswa"); //membuat variabel query untuk menyimpan eksekusi query ke database untuk menampilkan isi tabel tb_siswa
$i = 2; //membuat variabel $i yang memiliki nilai 2
$no = 1; //membuat variabel $no yang memiliki nilai 1
while($row = mysqli_fetch_array($query)) //membuat perulangan dengan variabel $row
{
	$sheet->setCellValue('A'.$i, $no++); //Mengisi cell  excel pada kolom A untuk nomor
	$sheet->setCellValue('B'.$i, $row['nama']);  //isi cell  excel pada kolom B untuk nama
	$sheet->setCellValue('C'.$i, $row['kelas']);  //isi cell  excel pada kolom C untuk kelas
	$sheet->setCellValue('D'.$i, $row['alamat']);  //isi cell  excel pada kolom D untuk alamat
	$i++; //Variabel $i tambah satu
}
 
$styleArray = [//untuk mengatur border
			'borders' => [
				'allBorders' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
		];
$i = $i - 1; // membuat variabel $i yang menimpan nilai $i-1
$sheet->getStyle('A1:D'.$i)->applyFromArray($styleArray); //untuk mengimplementasikan tampilan border yang sudah diatur
 
 
$writer = new Xlsx($spreadsheet); //membuat variabel $ writer untuk membuat file xlsx
$writer->save('Report Data Siswa.xlsx'); //meyimpan report file excel dengan nama file Report Data Siswa.xlsx