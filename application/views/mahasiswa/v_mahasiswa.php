<div class="content-wrapper">
    <section class="content-header">
        <h1>
            Data Mahasiswa
        </h1>
    </section>
    <section class="content">
        <div class="row">
            <div class="col-xs-12">
                <div class="box">
                    <div class="box-header">
                        <!-- flash data mulai. -->
                        <!-- jika ada session flash data maka. -->
                        <?php if ($this->session->flashdata()) : ?>
                            <div class="alert alert-success alert-dismissible">
                                <button type="button" class="close" data-dismiss="alert" aria-hidden="true">×</button>
                                <?= $this->session->flashdata('flash'); ?>
                            </div>
                        <?php endif; ?>
                        <!-- flash data selesai. -->
                        <a href="<?= base_url(); ?>mahasiswa/add" class="btn bg-olive btn-flat margin">Tambah Data</a>
                        <a href="<?= base_url(); ?>mahasiswa/export_tampil_mhs" class="btn bg-olive btn-flat margin">Export Excel</a>
                    </div>
                    <div class="box-body">
                        <table id="example1" class="table table-bordered table-striped">
                            <thead>
                                <tr>
                                    <th>No</th>
                                    <th>NIM</th>
                                    <th>Nama</th>
                                    <th>Jenis Kelamin</th>
                                    <th>Alamat</th>
                                    <th>Telepon</th>
                                    <th>Aksi</th>
                                </tr>
                            </thead>
                            <tbody>
                                <!-- perualangan untuk menampilkan data mulai -->
                                <?php
                                $i = 1;
                                foreach ($mahasiswa as $mhs) : ?>
                                    <tr>
                                        <td><?= $i++; ?></td>
                                        <td><?= $mhs['nim']; ?></td>
                                        <td><?= $mhs['nama']; ?></td>
                                        <td><?= $mhs['jk']; ?></td>
                                        <td><?= $mhs['alamat']; ?></td>
                                        <td><?= $mhs['telp']; ?></td>
                                        <td>
                                            <a href="<?= base_url(); ?>mahasiswa/update/<?= $mhs['id']; ?>" class="btn btn-default btn-flat" title="Ubah"><i class="fa fa-edit" style="color:blue"></i></a>
                                            <a href="<?= base_url(); ?>mahasiswa/delete/<?= $mhs['id']; ?>" class="btn btn-default btn-flat" title="Hapus" onclick="return confirm('Anda yakin menghapus data ini?');"><i class="fa fa-trash" style="color:red"></i></a>
                                        </td>
                                    </tr>
                                <?php endforeach; ?>
                                <!-- perualangan untuk menampilkan data selesai -->
                            </tbody>
                            <tfoot>
                                <tr>
                                    <th>No</th>
                                    <th>NIM</th>
                                    <th>Nama</th>
                                    <th>Jenis Kelamin</th>
                                    <th>Alamat</th>
                                    <th>Telepon</th>
                                    <th>Aksi</th>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </section>
</div>