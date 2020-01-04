<div class="content-wrapper">
    <section class="content-header">
        <h1>
            Tambah Data Mahasiswa
        </h1>
    </section>
    <section class="content">
        <div class="row">
            <div class="col-md-12">
                <div class="box box-primary">
                    <?= form_open(); ?>
                    <div class="box-body">
                        <div class="form-group">
                            <label for="nim">NIM</label>
                            <input type="text" name="nim" class="form-control" id="nim" placeholder="Masukkan NIM">
                            <!-- jika validasi error tampilkan validasi error di controller nim -->
                            <small class="text-form text-danger"><?= form_error('nim'); ?></small>
                        </div>
                        <div class="form-group">
                            <label for="nama">Nama</label>
                            <input type="text" name="nama" class="form-control" id="nama" placeholder="Masukkan Nama">
                            <!-- jika validasi error tampilkan validasi error di controller nama -->
                            <small class="text-form text-danger"><?= form_error('nama'); ?></small>
                        </div>
                        <div class="form-group">
                            <label for="jk">Jenis Kelamin</label>
                            <div>
                                <?= form_radio('jk', 'L', TRUE) ?>Laki- Laki
                                <?= form_radio('jk', 'P', FALSE) ?> Perempuan
                            </div>
                        </div>
                        <div class="form-group">
                            <label for="alamat">Alamat</label>
                            <input type="text" name="alamat" class="form-control" id="alamat" placeholder="Masukkan Alamat">
                            <!-- jika validasi error tampilkan validasi error di controller alamat -->
                            <small class="text-form text-danger"><?= form_error('alamat'); ?></small>
                        </div>
                        <div class="form-group">
                            <label for="telp">Telepon</label>
                            <input type="text" name="telp" class="form-control" id="telp" placeholder="Masukan Nomer Telepon">
                            <!-- jika validasi error tampilkan validasi error di controller telp -->
                            <small class="text-form text-danger"><?= form_error('telp'); ?></small>
                        </div>
                        <div class="form-group">
                            <button type="submit" name="simpan" class="btn btn-primary btn-flat">Simpan</button>
                        </div>
                    </div>
                    <?= form_close(); ?>
                </div>
            </div>
        </div>
    </section>
</div>