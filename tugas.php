<!DOCTYPE html>
<html>
<head>
	<title>FORM</title>
	<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO" crossorigin="anonymous">
	<style>
	.warning {color: #ff9900;}
</style>
</head>
<body>
<?php
$error_nama = $error_jkel = $error_nisn = $error_nik = $error_temlahir = $error_tgllahir = $error_noregakte = $error_agama = $error_kwn = $error_kebutuhan = $error_alamat = $error_rt = $error_rw = $error_dusun = $error_kelurahan = $error_kecamatan = $error_kdpos = $error_lintang = $error_bujur = $error_temtinggal = $error_transport = $error_nokks = $error_anakke = $error_kps = $error_nokps = "";

$nama = "";
$jkel = "";
$nisn = "";
$nik = "";
$temlahir = "";
$tgllahir = "";
$noregakte = "";
$agama = "";
$kwn = "";
$kebutuhan = "";
$alamat = "";
$rt = "";
$rw = "";
$dusun = "";
$kelurahan = "";
$kecamatan = "";
$kdpos = "";
$lintang = "";
$bujur = "";
$temtinggal = "";
$transport = "";
$nokks = "";
$anakke = "";
$kps = "";
$nokps = "";

//pengecekan
if ($_SERVER["REQUEST_METHOD"]=="POST") {
	if (empty($_POST["nama"])) {
		$error_nama = "Nama tidak boleh kosong";
	}else{
		if (!preg_match("/^[a-z A-z]*$/", $_POST["nama"])) {
			$error_nama = "Inputan hanya boleh huruf dan spasi";
		}else{
			$nama=cek_input($_POST["nama"]);
		}
	}
	if (empty($_POST["jkel"])) {
		$error_jkel = "Tidak boleh kosong";
	}else{
		$jkel = cek_input($_POST["jkel"]);
	}
	if (empty($_POST["nisn"])) {
		$error_nisn = "Tidak boleh kosong";
	}else{
		if (!is_numeric($_POST["nisn"])) {
			$error_nisn = "NISN hanya boleh angka";
		}else{
			$nisn=cek_input($_POST["nisn"]);
		}
	}
	if (empty($_POST["nik"])) {
		$error_nik = "Tidak boleh kosong";
	}else{
		if (!is_numeric($_POST["nik"])) {
			$error_nik = "NIK hanya boleh angka";
		}else{
			$nik=cek_input($_POST["nik"]);
		}
	}
	if (empty($_POST["temlahir"])) {
		$error_temlahir = "Tidak boleh Kosong";
	}else{
		if (!preg_match("/^[a-z A-z]*$/", $_POST["temlahir"])) {
			$error_temlahir = "Inputan hanya boleh huruf dan spasi";
		}else{
			$temlahir=cek_input($_POST["temlahir"]);
		}
	}
	if (empty($_POST["tgllahir"])) {
		$error_tgllahir = "Tidak boleh kosong";
	}else{
		$tgllahir=cek_input($_POST["tgllahir"]);
	}
	if (empty($_POST["noregakte"])) {
		$error_noregakte = "Tidak boleh kosong";
	}else{
		if (!is_numeric($_POST["noregakte"])) {
			$error_noregakte = "Noreg akte hanya boleh angka";
		}else{
			$noregakte=cek_input($_POST["noregakte"]);
		}
	}
	if (empty($_POST["agama"])) {
		$error_agama = "Tidak boleh kosong";
	}else{
		$agama=cek_input($_POST["agama"]);
	}
	if (empty($_POST["kwn"])) {
		$error_kwn = "Tidak boleh kosong";
	}else{
		if (!preg_match("/^[a-z A-z]*$/", $_POST["kwn"])) {
			$error_kwn = "Kewarganegaraan hanya boleh huruf";
		}else{
			$kwn=cek_input($_POST["kwn"]);
		}
	}
	if (empty($_POST["kebutuhan"])) {
		$error_kebutuhan = "---Tidak boleh kosong";
	}else{
		$kebutuhan = cek_input($_POST["kebutuhan"]);
	}
	if (empty($_POST["alamat"])) {
		$error_alamat = "Tidak boleh kosong";
	}else{
		if (!preg_match("/^[a-z A-z0-9+&@#\/%?=~_|!:,.;]*$/", $_POST["alamat"])) {
			$error_alamat = "Alamat hanya boleh huruf dan angka";
		}else{
			$alamat=cek_input($_POST["alamat"]);
		}
	}
	if (empty($_POST["rt"])) {
		$error_rt = "Tidak boleh kosong";
	}else{
		if (!is_numeric($_POST["rt"])) {
			$error_rt = "RT hanya boleh angka";
		}else{
			$rt=cek_input($_POST["rt"]);
		}
	}
	if (empty($_POST["rw"])) {
		$error_rw = "Tidak boleh kosong";
	}else{
		if (!is_numeric($_POST["rw"])) {
			$error_rw = "RW hanya boleh angka";
		}else{
			$rw=cek_input($_POST["rw"]);
		}
	}
	if (empty($_POST["dusun"])) {
		$error_dusun = "Tidak boleh kosong";
	}else{
		if (!preg_match("/^[a-z A-z]*$/", $_POST["dusun"])) {
			$error_dusun = "Dusun hanya boleh huruf dan spasi";
		}else{
			$dusun=cek_input($_POST["dusun"]);
		}
	}
	if (empty($_POST["kelurahan"])) {
		$error_kelurahan = "Tidak boleh kosong";
	}else{
		if (!preg_match("/^[a-z A-z]*$/", $_POST["kelurahan"])) {
			$error_kelurahan = "Kelurahan hanya boleh huruf dan spasi";
		}else{
			$kelurahan=cek_input($_POST["kelurahan"]);
		}
	}
	if (empty($_POST["kecamatan"])) {
		$error_kecamatan = "Tidak boleh kosong";
	}else{
		if (!preg_match("/^[a-z A-z]*$/", $_POST["kecamatan"])) {
			$error_kecamatan = "Kecamatan hanya boleh huruf dan spasi";
		}else{
			$kecamatan=cek_input($_POST["kecamatan"]);
		}
	}
	if (empty($_POST["kdpos"])) {
		$error_kdpos="Tidak boleh Kosong";
	}else{
		if (!is_numeric($_POST["kdpos"])) {
			$error_kdpos = "hanya boleh angka";
		}else{
			$kdpos=cek_input($_POST["kdpos"]);
		}
	}
	if ($_POST["lintang"]) {
		$lintang = cek_input($_POST["lintang"]);
	}
	if ($_POST["bujur"]) {
	$bujur = cek_input($_POST["bujur"]);
	}
	if (empty($_POST["temtinggal"])) {
		$error_temtinggal = "Tidak boleh kosong";
	}else{
		$temtinggal=cek_input($_POST["temtinggal"]);
	}
	if (empty($_POST["transport"])) {
		$error_transport = "Tidak boleh kosong";
	}else{
		$transport=cek_input($_POST["transport"]);
	}
	if (empty($_POST["anakke"])) {
		$error_anakke = "Tidak boleh kosong";
	}else{
		if (!is_numeric($_POST["anakke"])) {
			$error_anakke = "Hanya boleh angka";
		}else{
			$anakke=cek_input($_POST["anakke"]);
		}
	}
	if (empty($_POST["kps"])) {
		$error_kps = "Tidak boleh kosong";
	}else{
		$kps = cek_input($_POST["kps"]);
	}
	
}
function cek_input($data){
	$data = trim($data);
	$data = stripcslashes($data);
	$data = htmlspecialchars($data);
	return $data;
}
?>


<form method="POST" action="<?php echo htmlspecialchars($_SERVER["PHP_SELF"]);?>">
<table >
	<tr>
		<td>Nama Lengkap</td>
		<td> <input type="text" name="nama" class="form-control<?php echo($error_nama !="" ? "is-invalid" : "");?>" placeholder = "Nama Lengkap" value = "<?php echo ($nama); ?>"><span class="warning"><?php echo $error_nama;?></span></td>
	</tr>
	<tr>
		<td>Jenis Kelamin</td>
		<td> <input type="radio" name="jkel" class="<?php if(isset($jkel) && $jkel=="Laki") echo("checked");?>" value="Laki-Laki">Laki - Laki
		<input type="radio" name="jkel"  class="<?php if(isset($jkel) && $jkel=="Perempuan") echo("checked");?>" value="Perempuan">Perempuan<span class="warning"><?php echo $error_jkel;?></span></td>
	</tr>
	<tr>
		<td>NISN</td>
		<td> <input type="text" name="nisn" class="form-control<?php echo($error_nisn !="" ? "is-invalid" : "");?>" placeholder = "MAKS 10 digit" value = "<?php echo ($nisn); ?>"><span class="warning"><?php echo $error_nisn;?></span></td>
	</tr>
	<tr>
		<td>NIK/No. KITAS(untuk WNA)</td>
		<td> <input type="text" name="nik" class="form-control<?php echo($error_nik !="" ? "is-invalid" : "");?>" placeholder = "MAKS 16 digit" value="<?php echo($nik); ?>"><span class="warning"><?php echo $error_nik;?></span></td>
	</tr>
	<tr>
		<td>Tempat Lahir</td>
		<td> <input type="text" name="temlahir" placeholder="Tempat Lahir" class="form-control<?php echo($error_nik !="" ? "is-invalid" : "");?>" value="<?php echo($temlahir); ?>"><span class="warning"><?php echo $error_temlahir;?></span></td>
	</tr>
	<tr>
		<td>Tanggal Lahir</td>
		<td> <input type="date" name="tgllahir" class="form-control<?php echo($error_tgllahir !="" ? "is-invalid" : "");?>" value="<?php echo($tgllahir); ?>"><span class="warning"><?php echo $error_tgllahir;?></span></td>
	</tr>
	<tr>
		<td>No. Registrasi Akta Lahir</td>
		<td> <input type="text" name="noregakte" class="form-control<?php echo($error_noregakte !="" ? "is-invalid" : "");?>" placeholder = "No.Akta" value="<?php echo($noregakte); ?>"><span class="warning"><?php echo $error_noregakte;?></span></td>
	</tr>
	<tr>
		<td>Agama</td>
		<td> <select name="agama" class="form-control">
			<option value="Islam" class="<?= ($_POST["agama"]=="1")?"selected":"" ?>">Islam</option>
			<option value="Kristen" class="<?= ($_POST["agama"]=="2")?"selected":"" ?>">Kristen</option>
			<option value="Katolik" class="<?= ($_POST["agama"]=="3")?"selected":"" ?>">Katolik</option>
			<option value="Budha" class="<?= ($_POST["agama"]=="4")?"selected":"" ?>">Budha</option>
			<option value="Hindu" class="<?= ($_POST["agama"]=="5")?"selected":"" ?>">Hindu</option>
			<option value="Kong Hu Cu" class="<?= ($_POST["agama"]=="6")?"selected":"" ?>">Kong Hu CU</option>
		</select><span class="warning"><?php echo $error_agama;?></span></td>
	</tr>
	<tr>
		<td>Kewarganegaraan</td>
		<td> <input type="text" name="kwn" class="form-control<?php echo($error_kwn !="" ? "is-invalid" : "");?>" placeholder = "Kewarganegaraan" value="<?php echo($kwn); ?>"><span class="warning"><?php echo $error_kwn;?></span></td>
	</tr>
	<tr>
		<td>Bekebutuhan Khusus</td>
		<td> <input type="radio" name="kebutuhan" class="<?php if(isset($kebutuhan) && $kebutuhan=="Iya") echo("checked");?>" value="Iya">IYA
			<input type="radio" name="kebutuhan" class="<?php if(isset($kebutuhan) && $kebutuhan=="Tidak") echo("checked");?>" value="Tidak">TIDAK
		<span class="warning"><?php echo $error_kebutuhan;?></span>
		</td>
	</tr>
	<tr>
		<td>Alamat Rumah</td>
		<td> <input type="text" name="alamat" class="form-control<?php echo($error_alamat !="" ? "is-invalid" : "");?>" placeholder = "Alamat Rumah" value="<?php echo($alamat); ?>"><span class="warning"><?php echo $error_alamat;?></span></td>
	</tr>
	<tr>
		<td>RT</td>
		<td> <input type="text" name="rt" class="form-control<?php echo($error_rt !="" ? "is-invalid" : "");?>" placeholder = "RT" value="<?php echo($rt); ?>"><span class="warning"><?php echo $error_rt;?></span></td>
	</tr>
	<tr>
		<td>RW</td>
		<td> <input type="text" name="rw" class="form-control<?php echo($error_rw !="" ? "is-invalid" : "");?>" placeholder = "RW" value="<?php echo($rw); ?>"><span class="warning"><?php echo $error_rw;?></span></td>
	</tr>
	<tr>
		<td>Nama dusun</td>
		<td> <input type="text" name="dusun" class="form-control<?php echo($error_dusun !="" ? "is-invalid" : "");?>" placeholder = "Nama Dusun" value="<?php echo($dusun); ?>"><span class="warning"><?php echo $error_dusun;?></span></td>
	</tr>
	<tr>
		<td>Nama Kelurahan/Desa</td>
		<td> <input type="text" name="kelurahan" class="form-control<?php echo($error_kelurahan !="" ? "is-invalid" : "");?>" placeholder = "Nama Kelurahan" value="<?php echo($kelurahan); ?>"><span class="warning"><?php echo $error_kelurahan;?></span></td>
	</tr>
	<tr>
		<td>Kecamatan</td>
		<td> <input type="text" name="kecamatan" class="form-control<?php echo($error_kecamatan !="" ? "is-invalid" : "");?>" placeholder = "Nama Kecamatan" value="<?php echo($kecamatan); ?>"><span class="warning"><?php echo $error_kecamatan;?></span></td>
	</tr>
	<tr>
		<td>Kode Pos</td>
		<td> <input type="text" name="kdpos" class="form-control<?php echo($error_kdpos !="" ? "is-invalid" : "");?>" placeholder = "Kode Pos" value="<?php echo($kdpos) ?>"><span class="warning"><?php echo $error_kdpos;?></span></td>
	</tr>
	<tr>
		<td>Lintang</td>
		<td> <input type="text" name="lintang" class="form-control<?php echo($error_lintang !="" ? "is-invalid" : "");?>"  placeholder = "Lintang" value="<?php echo($lintang); ?>"><span class="warning"><?php echo $error_lintang;?></span></td>
	</tr>
	<tr>
		<td>Bujur</td>
		<td> <input type="text" name="bujur" class="form-control<?php echo($error_bujur !="" ? "is-invalid" : "");?>" placeholder = "Bujur" value="<?php echo($bujur); ?>"><span class="warning"><?php echo $error_bujur;?></span></td>
	</tr>
	<tr>
		<td>Tempat Tinggal</td>
		<td> <select name="temtinggal" class="form-control">
			<option value="Bersama Orang tua" class="<?= ($_POST["temtinggal"]=="1")?"selected":"" ?>">Bersama Orang tua</option>
			<option value="Kos" class="<?= ($_POST["temtinggal"]=="2")?"selected":"" ?>">Kos</option>
			<option value="Kontrakan" class="<?= ($_POST["temtinggal"]=="3")?"selected":"" ?>">Kontrakan</option>
			<option value="Bersama Wali" class="<?= ($_POST["temtinggal"]=="4")?"selected":"" ?>">Bersama Wali</option>
			</select><span class="warning"><?php echo $error_temtinggal;?></span></td>
	</tr>
	<tr>
		<td>Moda Transportasi</td>
		<td> <select name="transport" class="form-control">
			<option value="Jalan kaki" class="<?= ($_POST["transport"]=="1")?"selected":"" ?>">Jalan kaki</option>
			<option value="kendaraan pribadi" class="<?= ($_POST["transport"]=="2")?"selected":"" ?>">Kendaraan Pribadi</option>
			<option value="Angkutan umum" class="<?= ($_POST["transport"]=="3")?"selected":"" ?>">Angkutan Umum</option>
		</select><span class="warning"><?php echo $error_transport;?></span></td>
	</tr>
	<tr>
		<td>Nomor KKS(Kartu Keluarga Sejahtera)</td>
		<td> <input type="text" name="nokks" class="form-control<?php echo($error_nokks !="" ? "is-invalid" : "");?>" placeholder = "jika ada silahkan isi" value="<?php echo($nokks); ?>"><span class="warning"><?php echo $error_nokks;?></span></td>
	</tr>
	<tr>
		<td>Anak ke-</td>
		<td><input type="number" name="anakke" class="form-control<?php echo($error_anakke !="" ? "is-invalid" : "");?>" value="<?php echo($anakke); ?>"><span class="warning"><?php echo $error_anakke;?></span></td>
	</tr>
	<tr>
		<td>Penerima KPS/PKH</td>
		<td> <input type="radio" name="kps" class="<?php if(isset($kps) && $kps=="Iya") echo("checked");?>" value="Iya">IYA
			<input type="radio" name="kps" class="<?php if(isset($kps) && $kps=="Tidak") echo("checked");?>" value="Tidak">TIDAK
		<span class="warning"><?php echo $error_kps;?></span>
		</td>
	</tr>
	<tr>
		<td>Nomor KPS/PKH</td>
		<td><input type="text" name="nokps" class="form-control<?php echo($error_nokps !="" ? "is-invalid" : "");?>" placeholder = "jika ounya boleh diisi" value="<?php echo($nokps); ?>"><span class="warning"><?php echo $error_nokps;?></span></td>
	</tr>
	<tr>
		<td colspan="2" ><input type="submit" name="kirim" value="Kirim"  class="btn btn-success btn-lg btn-block"></td>
	</tr>
</table>
</form>
</body>
</html>
<?php
$server = "localhost";
$username = "root";
$password = "";
$namadb = "web";

$conn = mysqli_connect($server,$username,$password)or die("Proses simpan ke database gagal");
$con = mysqli_select_db($conn,$namadb);

if (!empty($nama) && !empty($jkel) && !empty($nisn) && !empty($nik) && !empty($temlahir) && !empty($tgllahir) && !empty($noregakte) && !empty($agama) && !empty($kwn) && !empty($kebutuhan) && !empty($alamat) && !empty($rt) && !empty($rw) && !empty($dusun) && !empty($kelurahan) && !empty($kecamatan) && !empty($kdpos) && !empty($temtinggal) && !empty($transport) && !empty($anakke) && !empty($kps)) {
$sql = mysqli_query($conn,"INSERT INTO validation (nama,jkel,nisn,nik,temlahir,tgllahir,noregakte,agama,kwn,kebutuhan,alamat,rt,rw,dusun,kelurahan,kecamatan,kdpos,lintang,bujur,temtinggal,transport,nokks,anakke,kps,nokps) VALUES ('$nama','$jkel','$nisn','$nik','$temlahir','$tgllahir','$noregakte','$agama','$kwn','$kebutuhan','$alamat','$rt','$rw','$dusun','$kelurahan','$kecamatan','$kdpos','$lintang','$bujur','$temtinggal','$transport','$nokks','$anakke','$kps','$nokps')");
if ($sql) {
	echo "<script>alert('Berhasil simpan')</script>";
	echo "<script>location.href='tampil.php'</script>";
}else{
	echo "<script>alert('Ada yang kurang mohon periksa kembali')</script>";
}
}
?>