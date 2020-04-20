<?php 
include 'koneksi2.php';
?>
<h1>DATA FORMULIR PESERTA DIDIK</h1>
    <label>Export :</label>
    <a href="eksport.php"><button>Export to Excel</button></a><br />
    <hr />
<table border="1" cellpadding="5" cellspacing="0">
    <tr>
        <th>Nama</th>
        <th>Jenis Kelamin</th>
        <th>Nisn</th>
        <th>NIK</th>
        <th>Tempat Lahir</th>
        <th>Tanggal Lahir</th>
        <th>No. Registrasi Akta Lahir</th>
        <th>Agama & Kepercayaan</th>
        <th>Kewarganegaraan</th>
        <th>Berkebutuhan Khusus</th>
        <th>Alamat Jalan</th>
        <th>RT</th>
        <th>RW</th>
        <th>Nama Dusun</th>
        <th>Nama Kelurahan/Desa</th>
        <th>Kecamatan</th>
        <th>Kode Pos</th>
        <th>Lintang</th>
        <th>Bujur</th>
        <th>Tempat Tinggal</th>
        <th>Moda Transportasi</th>
        <th>Nomor KKS</th>
        <th>Anak ke-berapa</th>
        <th>Penerima KPS/KPH</th>
        <th>No. KPS/KPH</th>
    </tr>
    <?php 
    $sql = "SELECT * FROM validation";
    $qry=mysqli_query($conn, $sql) or die("Gagal");
    while($d = mysqli_fetch_array($qry)){
    ?>
    <tr>
        <td><?php echo $d['nama']; ?></td>
        <td><?php echo $d['jkel']; ?></td>
        <td><?php echo $d['nisn']; ?></td>
        <td><?php echo $d['nik']; ?></td>
        <td><?php echo $d['temlahir']; ?></td>
        <td><?php echo $d['tgllahir']; ?></td>
        <td><?php echo $d['noregakte']; ?></td>
        <td><?php echo $d['agama']; ?></td>
        <td><?php echo $d['kwn']; ?></td>
        <td><?php echo $d['kebutuhan']; ?></td>
        <td><?php echo $d['alamat']; ?></td>
        <td><?php echo $d['rt']; ?></td>
        <td><?php echo $d['rw']; ?></td>
        <td><?php echo $d['dusun']; ?></td>
        <td><?php echo $d['kelurahan']; ?></td>
        <td><?php echo $d['kecamatan']; ?></td>
        <td><?php echo $d['kdpos']; ?></td>
        <td><?php echo $d['lintang']; ?></td>
        <td><?php echo $d['bujur']; ?></td>
        <td><?php echo $d['temtinggal']; ?></td>
        <td><?php echo $d['transport']; ?></td>
        <td><?php echo $d['nokks']; ?></td>
        <td><?php echo $d['anakke']; ?></td>
        <td><?php echo $d['kps']; ?></td>
        <td><?php echo $d['nokps']; ?></td>
    </tr>
    <?php
    } 
    ?>
</table>
