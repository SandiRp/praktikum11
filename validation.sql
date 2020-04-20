-- phpMyAdmin SQL Dump
-- version 4.7.4
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: 13 Apr 2020 pada 23.35
-- Versi Server: 10.1.29-MariaDB
-- PHP Version: 7.2.0

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `sandi`
--

-- --------------------------------------------------------

--
-- Struktur dari tabel `validation`
--

CREATE TABLE `validation` (
  `nama` varchar(50) NOT NULL,
  `jkel` varchar(12) NOT NULL,
  `nisn` varchar(10) NOT NULL,
  `nik` varchar(16) NOT NULL,
  `temlahir` varchar(25) NOT NULL,
  `tgllahir` date NOT NULL,
  `noregakte` varchar(32) NOT NULL,
  `agama` varchar(15) NOT NULL,
  `kwn` varchar(20) NOT NULL,
  `kebutuhan` varchar(6) NOT NULL,
  `alamat` varchar(50) NOT NULL,
  `rt` varchar(4) NOT NULL,
  `rw` varchar(4) NOT NULL,
  `dusun` varchar(25) NOT NULL,
  `kelurahan` varchar(25) NOT NULL,
  `kecamatan` varchar(25) NOT NULL,
  `kdpos` varchar(5) NOT NULL,
  `lintang` varchar(25) DEFAULT NULL,
  `bujur` varchar(25) DEFAULT NULL,
  `temtinggal` varchar(20) NOT NULL,
  `transport` varchar(20) NOT NULL,
  `nokks` varchar(6) DEFAULT NULL,
  `anakke` varchar(2) NOT NULL,
  `kps` varchar(6) NOT NULL,
  `nokps` varchar(15) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data untuk tabel `validation`
--

INSERT INTO `validation` (`nama`, `jkel`, `nisn`, `nik`, `temlahir`, `tgllahir`, `noregakte`, `agama`, `kwn`, `kebutuhan`, `alamat`, `rt`, `rw`, `dusun`, `kelurahan`, `kecamatan`, `kdpos`, `lintang`, `bujur`, `temtinggal`, `transport`, `nokks`, `anakke`, `kps`, `nokps`) VALUES
('ragil', 'Laki-Laki', '0897656621', '098555627828392', 'surabaya', '2020-04-22', '98765456698', 'Islam', 'Indonesia', 'Tidak', 'kupang', '02', '05', 'kupang', 'kupang', 'sawahan', '60253', '', 'tau', 'Bersama Orang tua', 'kendaraan pribadi', '', '4', 'Iya', ''),
('sandi', 'Laki-Laki', '0987263536', '0985556278283927', 'surabaya', '2020-04-14', '98765456698', 'Islam', 'Indonesia', 'Tidak', 'kupang', '02', '05', 'kupang', 'kupang', 'sawahan', '60253', 'tau', 'tau', 'Bersama Orang tua', 'kendaraan pribadi', '097789', '-4', 'Iya', '9878888');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `validation`
--
ALTER TABLE `validation`
  ADD PRIMARY KEY (`nisn`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
