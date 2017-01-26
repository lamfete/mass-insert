<?php

include 'PHPExcel/Classes/PHPExcel/IOFactory.php';

$inputFileName = 'stok_awal.xls';
error_reporting(E_ERROR | E_WARNING);
//  Read your Excel workbook
try {
    $inputFileType = PHPExcel_IOFactory::identify($inputFileName);
    $objReader = PHPExcel_IOFactory::createReader($inputFileType);
    $objPHPExcel = $objReader->load($inputFileName);
} catch(Exception $e) {
    die('Error loading file "'.pathinfo($inputFileName,PATHINFO_BASENAME).'": '.$e->getMessage());
}

//  Get worksheet dimensions
$sheet = $objPHPExcel->getSheet(0);
$highestRow = $sheet->getHighestRow();
$highestColumn = $sheet->getHighestColumn();

//  Loop through each row of the worksheet in turn
$i = 0;
for ($row = 1; $row <= $highestRow; $row++){
    //  Read a row of data into an array
    $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
	
	/*
	var_dump($rowData);
	exit;
	*/
	
	
	//INSERT DATA STOK AWAL
	echo "INSERT INTO MUTASI_GUDANG_DETAIL(KODE_MUTASI_GUDANG, KODE_BARANG, JUMLAH)
		VALUES('MG00000001', '".$rowData[0][0]."', ".$rowData[0][2].");
		<br />";
	
	//UNTUK INSERT DATA STOK AWAL, TIDAK PERLU INSERT KE GUDANG_BARANG
	/*
	echo "INSERT INTO GUDANG_BARANG(KODE_GUDANG, KODE_BARANG, KODE_TRANSAKSI, SALDO_AWAL, DEBET, KREDIT, SALDO_AKHIR, TANGGAL, STATUS)
		VALUES
		('WH01', '".$rowData[0][0]."', 'MG00000001', 0, ".$rowData[0][2].", 0, ".$rowData[0][2].", '2017-01-01 00:00:00.000', 1);
		<br />";
	*/
	/*
	//INSERT DATA CUSTOMER
	echo "INSERT INTO CUSTOMER(
		KODE_CUSTOMER, NAMA, ALAMAT, KODE_POS, TELEPON, 
		FAX, KODE_LIMIT_ORDER, SYARAT_BAYAR, KODE_KOTA, KODE_PROPINSI, 
		SHIP_TO, BILL_TO, NAMA_SINGKAT, KODE_SALES, AKTIF, 
		KODE_WILAYAH, TIPE_CUSTOMER, KODE_STATUS_USAHA, KODE_DISKON, KODE_KECAMATAN, 
		KODE_KELURAHAN, EMAIL, NAMA_PEMILIK)
		VALUES
		('".$rowData[0][1]."', '".$rowData[0][2]."', '".$rowData[0][3]."', '".$rowData[0][4]."', '".$rowData[0][7]."',
		'".$rowData[0][8]."', '".$rowData[0][9]."', ".$rowData[0][10].", '".$rowData[0][11]."', '".$rowData[0][13]."',
		'".$rowData[0][15]."', '".$rowData[0][17]."', '".$rowData[0][20]."', '".$rowData[0][21]."', '".$rowData[0][22]."',
		'".$rowData[0][23]."', '".$rowData[0][24]."', '".$rowData[0][26]."', '".$rowData[0][27]."', '".$rowData[0][28]."', 
		'".$rowData[0][30]."', '".$rowData[0][32]."', '".$rowData[0][34]."');
		<br />";
	*/
	
	//INSERT DATA BARANG
	/*
	echo "INSERT INTO BARANG (
		KODE_BARANG, NAMA_BARANG, TIPE, KODE_MEREK_PRODUK,
		KODE_JENIS_PRODUK, HARGA_JUAL, WARNA, UKURAN,
		PANJANG, SATUAN, JUMLAH_ISI, TONASE,
		KODE_SUB_JENIS, KODE_UKURAN, KODE_WARNA, KODE_KWALITAS,
		STATUS_PINDAH_KODE
		)
		VALUES 
		('".$rowData[0][0]."', '".$rowData[0][1]."', '".$rowData[0][2]."', '".$rowData[0][3]."', 
		'".$rowData[0][4]."', ".$rowData[0][5].", '".$rowData[0][6]."', '".$rowData[0][7]."', 
		".$rowData[0][8].", '".$rowData[0][9]."', ".$rowData[0][10].", ".$rowData[0][11].", 
		'".$rowData[0][12]."', '".$rowData[0][13]."', '".$rowData[0][14]."' , '".$rowData[0][15]."',
		'".$rowData[0][16]."');<br />";
	*/
	$i++;
	/*
	echo "
		INSERT INTO SUB_JENIS(KODE_SUB_JENIS, NAMA_SUB_JENIS)
		VALUES (".$i.", '".$rowData[0][0]."');<br />
	";*/
	
	/*echo "
		UPDATE SUB_JENIS
		SET KODE_JENIS_PRODUK = '".$rowData[0][1]."',
		KODE_MEREK_PRODUK = '".$rowData[0][2]."'
		WHERE KODE_SUB_JENIS = '".$i."';<br />
	";*/
	
	/*
	//INSERT CUSTOMER_DISKON_BARU
	echo "
		INSERT INTO CUSTOMER_DISKON_BARU
		VALUES ('".$rowData[0][0]."', 'C', 'RUCIKA', 'WAVIN');<br />
	";
	
	echo "
		INSERT INTO CUSTOMER_DISKON_BARU
		VALUES ('".$rowData[0][0]."', 'C', 'WAVINSTD', 'WAVIN');<br />
	";
	*/
	
	/*
	//update pricelist minus ppn
	echo "
		UPDATE BARANG
		SET HARGA_BELI = '".$rowData[0][1]."'
		WHERE KODE_BARANG = '".$rowData[0][0]."';<br />
	"; */
	
	/*
	//INSERT CUSTOMER_LIMIT
	echo "
		INSERT INTO CUSTOMER_LIMIT
		VALUES ('".$rowData[0][0]."', '".$rowData[0][1]."');<br />
	";*/
}

?>
