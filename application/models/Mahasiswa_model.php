<?php

class Mahasiswa_model extends CI_Model

{
    // fungsi ambil data.
    public function getAllMahasiswa()
    {
        // mendapatkan dan mengembalikan nilai view tampil_mhs berbentuk array.
        return $this->db->get('tampil_mhs')->result_array();
    }

    // fungsi tambah data.
    public function tambahDataMahasiswa()
    {
        // array data mulai.
        $data = [
            // mengambil data dari inputan post "name" mulai.
            "nim" => $this->input->post('nim', true),
            "nama" => $this->input->post('nama', true),
            "jk" => $this->input->post('jk', true),
            "alamat" => $this->input->post('alamat', true),
            "telp" => $this->input->post('telp', true)
            // mengambil data dari inputan post "name" selesai.
        ];
        // array data selesai.
        // memasukkan $data array ke dalam database tabel tb_mhs.
        $this->db->insert('tb_mhs', $data);
    }

    // fungsi hapus data berdasarkan parameter $id.
    public function hapusDataMahasiswa($id)
    {
        // mencari data yang sama dengan parameter.
        $this->db->where('id', $id);
        // menghapus data yang dicari.
        $this->db->delete('tb_mhs');
    }

    // fungsi ambil data berdasarkan parameter $id.
    public function getMahasiswaById($id)
    {
        // mencari dan mengembalikan data yang didapat berbentuk baris array.
        return $this->db->get_where('tb_mhs', ['id' => $id])->row_array();
    }

    // fungsi ubah data.
    public function ubahDataMahasiswa()
    {
        // array data mulai.
        $data = [
            // mengambil data dari inputan post "name" mulai.
            "nim" => $this->input->post('nim', true),
            "nama" => $this->input->post('nama', true),
            "jk" => $this->input->post('jk', true),
            "alamat" => $this->input->post('alamat', true),
            "telp" => $this->input->post('telp', true)
            // mengambil data dari inputan post "name" selesai.
        ];
        // array data selesai.
        // mencari data id dari inputan post yang dikirim. 
        $this->db->where('id', $this->input->post('id'));
        // mengupdate data yang dicari dengan $data array ke dalam database tabel tb_mhs.
        $this->db->update('tb_mhs', $data);
    }

    // fungsi ambil data.
    public function getLog()
    {
        // mengembalikan nilai view tampil_semua berbentuk array.
        return $this->db->get('tampil_semua')->result_array();
    }
}

/* End of file Mahasiswa_model.php */
