-- phpMyAdmin SQL Dump
-- version 4.8.5
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Jan 04, 2020 at 08:17 AM
-- Server version: 10.1.38-MariaDB
-- PHP Version: 5.6.40

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `db_pw`
--

-- --------------------------------------------------------

--
-- Stand-in structure for view `tampil_mhs`
-- (See below for the actual view)
--
CREATE TABLE `tampil_mhs` (
`id` int(11)
,`nim` varchar(12)
,`nama` varchar(32)
,`jk` varchar(1)
,`alamat` varchar(32)
,`telp` varchar(13)
);

-- --------------------------------------------------------

--
-- Stand-in structure for view `tampil_semua`
-- (See below for the actual view)
--
CREATE TABLE `tampil_semua` (
`nim` varchar(12)
,`nama` varchar(32)
,`jk` varchar(1)
,`alamat` varchar(32)
,`telp_lama` varchar(13)
,`telp_baru` varchar(13)
,`tgl_diubah` datetime
);

-- --------------------------------------------------------

--
-- Table structure for table `tb_log`
--

CREATE TABLE `tb_log` (
  `id` int(11) NOT NULL,
  `nim` varchar(12) NOT NULL,
  `telp_lama` varchar(13) NOT NULL,
  `telp_baru` varchar(13) NOT NULL,
  `tgl_diubah` datetime NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_log`
--

INSERT INTO `tb_log` (`id`, `nim`, `telp_lama`, `telp_baru`, `tgl_diubah`) VALUES
(1, '161240000551', '08987654321', '08987654320', '2020-01-03 19:51:21'),
(2, '161240000552', '08987654322', '08987654323', '2020-01-04 13:20:03'),
(3, '161240000552', '089501800531', '089501800531', '2020-01-04 13:20:08');

-- --------------------------------------------------------

--
-- Table structure for table `tb_mhs`
--

CREATE TABLE `tb_mhs` (
  `id` int(11) NOT NULL,
  `nim` varchar(12) NOT NULL,
  `nama` varchar(32) NOT NULL,
  `jk` varchar(1) NOT NULL,
  `alamat` varchar(32) NOT NULL,
  `telp` varchar(13) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tb_mhs`
--

INSERT INTO `tb_mhs` (`id`, `nim`, `nama`, `jk`, `alamat`, `telp`) VALUES
(1, '161240000551', 'Arky Ikfarikza', 'L', 'Wonosari', '08987654320'),
(2, '161240000552', 'Arky Ikfarikza', 'L', 'Wonosari', '08987654323'),
(3, '161240000553', 'Arky Ikfarikza', 'L', 'Wonosari', '08987654323'),
(4, '161240000554', 'Arky Ikfarikza', 'L', 'Wonosari', '08987654324'),
(5, '161240000555', 'Arky Ikfarikza', 'L', 'Wonosari', '08987654325'),
(6, '161240000556', 'Arky Ikfarikza', 'L', 'Wonosari', '08987654326'),
(7, '161240000557', 'Arky Ikfarikza', 'L', 'Wonosari', '08987654327'),
(8, '161240000558', 'Arky Ikfarikza', 'L', 'Wonosari', '08987654328'),
(10, '161240000552', 'Asli', 'P', 'Tahunan', '089501800531');

--
-- Triggers `tb_mhs`
--
DELIMITER $$
CREATE TRIGGER `update_telp` AFTER UPDATE ON `tb_mhs` FOR EACH ROW begin
insert into tb_log set
nim = old.nim,
telp_baru = new.telp,
telp_lama = old.telp,
tgl_diubah = now();
end
$$
DELIMITER ;

-- --------------------------------------------------------

--
-- Structure for view `tampil_mhs`
--
DROP TABLE IF EXISTS `tampil_mhs`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `tampil_mhs`  AS  select `tb_mhs`.`id` AS `id`,`tb_mhs`.`nim` AS `nim`,`tb_mhs`.`nama` AS `nama`,`tb_mhs`.`jk` AS `jk`,`tb_mhs`.`alamat` AS `alamat`,`tb_mhs`.`telp` AS `telp` from `tb_mhs` ;

-- --------------------------------------------------------

--
-- Structure for view `tampil_semua`
--
DROP TABLE IF EXISTS `tampil_semua`;

CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `tampil_semua`  AS  select `tb_mhs`.`nim` AS `nim`,`tb_mhs`.`nama` AS `nama`,`tb_mhs`.`jk` AS `jk`,`tb_mhs`.`alamat` AS `alamat`,`tb_log`.`telp_lama` AS `telp_lama`,`tb_log`.`telp_baru` AS `telp_baru`,`tb_log`.`tgl_diubah` AS `tgl_diubah` from (`tb_mhs` join `tb_log` on((`tb_mhs`.`nim` = `tb_log`.`nim`))) ;

--
-- Indexes for dumped tables
--

--
-- Indexes for table `tb_log`
--
ALTER TABLE `tb_log`
  ADD PRIMARY KEY (`id`);

--
-- Indexes for table `tb_mhs`
--
ALTER TABLE `tb_mhs`
  ADD PRIMARY KEY (`id`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `tb_log`
--
ALTER TABLE `tb_log`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4;

--
-- AUTO_INCREMENT for table `tb_mhs`
--
ALTER TABLE `tb_mhs`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=11;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
