<!DOCTYPE html>
<html>

<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <!-- judul tab -->
    <title><?= $judul; ?></title>
    <meta content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no" name="viewport">
    <!-- plugin dari folder vendor mulai -->
    <link rel="stylesheet" href="<?= base_url() ?>vendor/bower_components/bootstrap/dist/css/bootstrap.min.css">
    <link rel="stylesheet" href="<?= base_url() ?>vendor/bower_components/font-awesome/css/font-awesome.min.css">
    <link rel="stylesheet" href="<?= base_url() ?>vendor/bower_components/Ionicons/css/ionicons.min.css">
    <link rel="stylesheet" href="<?= base_url() ?>vendor/dist/css/AdminLTE.min.css">
    <link rel="stylesheet" href="<?= base_url() ?>vendor/dist/css/skins/skin-blue.min.css">
    <link rel="stylesheet" href="<?= base_url() ?>vendor/dist/css/skins/_all-skins.min.css">
    <link rel="stylesheet" href="<?= base_url() ?>vendor/bower_components/datatables.net-bs/css/dataTables.bootstrap.min.css">
    <!-- plugin dari folder vendor selesai -->
    <!-- plugin online -->
    <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Source+Sans+Pro:300,400,600,700,300italic,400italic,600italic">
</head>

<body class="hold-transition skin-blue sidebar-mini fixed">
    <div class="wrapper">
        <header class="main-header">
            <a href="<?= base_url() ?>" class="logo">
                <span class="logo-mini">pwfl</span>
                <span class="logo-lg">UAS</span>
            </a>
            <nav class="navbar navbar-static-top" role="navigation">
                <a href="#" class="sidebar-toggle" data-toggle="push-menu" role="button">
                    <span class="sr-only">Toggle navigation</span>
                </a>
            </nav>
        </header>