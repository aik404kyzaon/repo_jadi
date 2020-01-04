<div class="content-wrapper">
    <section class="content-header">
        <h1>
            Riwayat
        </h1>
    </section>
    <section class="content container-fluid">
        <div class="row">
            <div class="col-xs-12">
                <div class="box">
                    <a href="<?= base_url(); ?>mahasiswa/export_tampil_semua" class="btn bg-olive btn-flat margin">Export Excel</a>
                    <div class="box-body">
                        <table id="example1" class="table table-bordered table-striped">
                            <thead>
                                <tr>
                                    <th>No</th>
                                    <th>NIM</th>
                                    <th>Nama</th>
                                    <th>Jenik Kelamin</th>
                                    <th>Alamat</th>
                                    <th>Telepon Lama</th>
                                    <th>Telepon Baru</th>
                                    <th>Tanggal Diubah</th>
                                </tr>
                            </thead>
                            <tbody>
                                <?php
                                $i = 1;
                                foreach ($mahasiswa as $mhs) : ?>
                                    <tr>
                                        <td><?= $i++; ?></td>
                                        <td><?= $mhs['nim']; ?></td>
                                        <td><?= $mhs['nama']; ?></td>
                                        <td><?= $mhs['jk']; ?></td>
                                        <td><?= $mhs['alamat']; ?></td>
                                        <td><?= $mhs['telp_lama']; ?></td>
                                        <td><?= $mhs['telp_baru']; ?></td>
                                        <td><?= $mhs['tgl_diubah']; ?></td>
                                    </tr>
                                <?php endforeach; ?>
                            </tbody>
                            <tfoot>
                                <tr>
                                    <th>No</th>
                                    <th>NIM</th>
                                    <th>Nama</th>
                                    <th>Jenik Kelamin</th>
                                    <th>Alamat</th>
                                    <th>Telepon Lama</th>
                                    <th>Telepon Baru</th>
                                    <th>Tanggal Diubah</th>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </section>
</div>