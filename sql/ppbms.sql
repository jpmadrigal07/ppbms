-- phpMyAdmin SQL Dump
-- version 4.8.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Aug 10, 2018 at 01:04 PM
-- Server version: 10.1.33-MariaDB
-- PHP Version: 7.2.6

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `ppbms`
--

-- --------------------------------------------------------

--
-- Table structure for table `ppbms_area_price`
--

CREATE TABLE `ppbms_area_price` (
  `id` int(11) NOT NULL,
  `area_price_name` varchar(100) NOT NULL,
  `area_price_price` int(11) NOT NULL,
  `area_price_status` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `ppbms_area_price`
--

INSERT INTO `ppbms_area_price` (`id`, `area_price_name`, `area_price_price`, `area_price_status`) VALUES
(1, 'Paranaque', 9, 1),
(2, 'Las Pinas', 9, 1),
(3, 'Muntinlupa', 9, 1),
(4, 'Marikina', 9, 1),
(5, 'Rizal', 10, 1),
(6, 'Laguna', 10, 1);

-- --------------------------------------------------------

--
-- Table structure for table `ppbms_barcode_middle_text`
--

CREATE TABLE `ppbms_barcode_middle_text` (
  `id` int(11) NOT NULL,
  `barcode_middle_text_code` varchar(255) NOT NULL,
  `barcode_middle_text_status` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `ppbms_dispatch_control_data`
--

CREATE TABLE `ppbms_dispatch_control_data` (
  `id` int(11) NOT NULL,
  `dispatch_control_messenger_id` int(11) NOT NULL,
  `dispatch_control_data_cycle_code` varchar(255) NOT NULL,
  `dispatch_control_data_pickup_date` datetime NOT NULL,
  `dispatch_control_data_sender` varchar(255) NOT NULL,
  `dispatch_control_data_del_type` varchar(255) NOT NULL,
  `dispatch_control_data_status` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `ppbms_dispatch_control_messenger`
--

CREATE TABLE `ppbms_dispatch_control_messenger` (
  `id` int(11) NOT NULL,
  `dispatch_control_messenger_name` varchar(255) NOT NULL,
  `dispatch_control_messenger_address` varchar(255) NOT NULL,
  `dispatch_control_messenger_prepared` varchar(255) NOT NULL,
  `dispatch_control_messenger_date` datetime NOT NULL,
  `dispatch_control_messenger_status` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `ppbms_encode_list`
--

CREATE TABLE `ppbms_encode_list` (
  `id` int(11) NOT NULL,
  `encode_file_name` varchar(255) NOT NULL,
  `encode_lists_date` datetime NOT NULL,
  `encode_lists_status` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `ppbms_record`
--

CREATE TABLE `ppbms_record` (
  `id` int(11) NOT NULL,
  `messenger_id` int(11) NOT NULL,
  `encode_lists_id` int(11) NOT NULL,
  `record_sender` varchar(255) NOT NULL,
  `record_deltype` varchar(255) NOT NULL,
  `record_pud` varchar(255) NOT NULL,
  `record_month` varchar(255) NOT NULL,
  `record_year` varchar(100) NOT NULL,
  `record_job_num` varchar(255) NOT NULL,
  `record_check_list` varchar(255) NOT NULL,
  `record_file_name` varchar(255) NOT NULL,
  `record_seq_num` varchar(255) NOT NULL,
  `record_cycle_code` varchar(255) NOT NULL,
  `record_qty` varchar(255) NOT NULL,
  `record_address` varchar(255) NOT NULL,
  `record_area` varchar(255) NOT NULL,
  `record_subs_name` varchar(255) NOT NULL,
  `record_bar_code` varchar(255) NOT NULL,
  `record_acct_num` varchar(255) NOT NULL,
  `record_date_receive` datetime DEFAULT NULL,
  `record_receive_by` varchar(255) DEFAULT NULL,
  `record_relation` varchar(255) DEFAULT NULL,
  `record_messenger` varchar(255) DEFAULT NULL,
  `record_status` varchar(255) DEFAULT NULL,
  `record_reason_rts` varchar(255) DEFAULT NULL,
  `record_remarks` varchar(255) DEFAULT NULL,
  `record_date_reported` datetime DEFAULT NULL,
  `record_status_status` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- --------------------------------------------------------

--
-- Table structure for table `ppbms_user`
--

CREATE TABLE `ppbms_user` (
  `id` int(11) NOT NULL,
  `user_email` varchar(100) NOT NULL,
  `user_pass` varchar(255) NOT NULL,
  `user_ip` varchar(50) DEFAULT NULL,
  `user_date_created` datetime NOT NULL,
  `user_last_login` datetime NOT NULL,
  `user_level` int(11) NOT NULL,
  `user_status` int(11) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 ROW_FORMAT=COMPACT;

--
-- Dumping data for table `ppbms_user`
--

INSERT INTO `ppbms_user` (`id`, `user_email`, `user_pass`, `user_ip`, `user_date_created`, `user_last_login`, `user_level`, `user_status`) VALUES
(1, 'admin', 'admin123', '1', '2017-07-20 22:33:20', '2018-08-10 08:56:20', 1, 1);

--
-- Indexes for dumped tables
--

--
-- Indexes for table `ppbms_area_price`
--
ALTER TABLE `ppbms_area_price`
  ADD PRIMARY KEY (`id`);

--
-- Indexes for table `ppbms_barcode_middle_text`
--
ALTER TABLE `ppbms_barcode_middle_text`
  ADD PRIMARY KEY (`id`);

--
-- Indexes for table `ppbms_dispatch_control_data`
--
ALTER TABLE `ppbms_dispatch_control_data`
  ADD PRIMARY KEY (`id`),
  ADD KEY `dispatch_data_idx` (`dispatch_control_messenger_id`,`dispatch_control_data_status`);

--
-- Indexes for table `ppbms_dispatch_control_messenger`
--
ALTER TABLE `ppbms_dispatch_control_messenger`
  ADD PRIMARY KEY (`id`),
  ADD KEY `messenger_idx` (`dispatch_control_messenger_status`) USING BTREE;

--
-- Indexes for table `ppbms_encode_list`
--
ALTER TABLE `ppbms_encode_list`
  ADD PRIMARY KEY (`id`),
  ADD KEY `view_list_idx` (`encode_lists_status`);

--
-- Indexes for table `ppbms_record`
--
ALTER TABLE `ppbms_record`
  ADD PRIMARY KEY (`id`),
  ADD KEY `export_idx` (`record_sender`,`record_deltype`,`record_pud`,`record_month`,`record_year`,`record_cycle_code`,`record_status_status`) USING BTREE,
  ADD KEY `receipt_idx` (`encode_lists_id`,`record_status_status`),
  ADD KEY `encode_mlist_idx` (`record_bar_code`),
  ADD KEY `assign_messenger_idx` (`record_cycle_code`,`record_bar_code`) USING BTREE;

--
-- Indexes for table `ppbms_user`
--
ALTER TABLE `ppbms_user`
  ADD PRIMARY KEY (`id`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `ppbms_area_price`
--
ALTER TABLE `ppbms_area_price`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=7;

--
-- AUTO_INCREMENT for table `ppbms_barcode_middle_text`
--
ALTER TABLE `ppbms_barcode_middle_text`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `ppbms_dispatch_control_data`
--
ALTER TABLE `ppbms_dispatch_control_data`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `ppbms_dispatch_control_messenger`
--
ALTER TABLE `ppbms_dispatch_control_messenger`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `ppbms_encode_list`
--
ALTER TABLE `ppbms_encode_list`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `ppbms_record`
--
ALTER TABLE `ppbms_record`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT;

--
-- AUTO_INCREMENT for table `ppbms_user`
--
ALTER TABLE `ppbms_user`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=2;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
