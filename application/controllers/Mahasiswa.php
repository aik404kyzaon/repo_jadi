<?php

defined('BASEPATH') or exit('No direct script access allowed');

class Mahasiswa extends CI_Controller
{
    // fungsi selalu dijalankan.
    public function __construct()
    {
        parent::__construct();
        // load models.
        $this->load->model('Mahasiswa_model');
    }

    // fungsi pertama dijalankan.
    public function index()
    {
        // $array data mulai.
        $data = [
            // judul tab.
            "judul" => "Data Mahasiswa",
            // load fungsi dari model.
            "mahasiswa" => $this->Mahasiswa_model->getAllMahasiswa()
        ];
        // array data selesai.
        // load view mulai
        // mengandung array $data
        $this->load->view('template/header', $data);
        $this->load->view('template/sidebar');
        // mengandung array $data
        $this->load->view('mahasiswa/v_mahasiswa', $data);
        $this->load->view('template/footer');
        // load view selesai.
    }

    // fungsi tambah
    public function add()
    {
        // data array mulai.
        $data = [
            // judul tab.
            "judul" => "Tambah Data Mahasiswa"
        ];
        // data array selesai.
        // mengeatur validasi dari inputan "name" mulai.
        $this->form_validation->set_rules('nim', 'NIM', 'required|numeric|max_length[12]');
        $this->form_validation->set_rules('nama', 'Nama', 'required|max_length[32]');
        $this->form_validation->set_rules('alamat', 'Alamat', 'required|max_length[32]');
        $this->form_validation->set_rules('telp', 'Telepon', 'required|numeric|max_length[13]');
        // mengatur validasi dari inputan "name" selesai.
        // pengkondisian jika validasi berjalan mulai.
        // jika validasi salah maka.
        if ($this->form_validation->run() == FALSE) {
            // load view mulai.
            // mengandung array $data
            $this->load->view('template/header', $data);
            $this->load->view('template/sidebar');
            // mengandung array $data
            $this->load->view('mahasiswa/v_tambah', $data);
            $this->load->view('template/footer');
            // load view selesai.
        }
        // jika benar maka.
        else {
            // load fungsi dari model.
            $this->Mahasiswa_model->tambahDataMahasiswa();
            // membuat session flash.
            $this->session->set_flashdata('flash', 'Data berhasil ditambahkan!');
            // kembali ke controller mahasiswa.
            redirect('mahasiswa'); // kembali ke controller mahasiswa.
        }
        // pengkodisian jika validasi berjalan selesai.
    }

    // fungsi ubah berdasarkan paramater $id.
    public function update($id)
    {
        // data array mulai.
        $data = [
            // judul tab.
            'judul' => 'Ubah Data Mahasiswa',
            // load fungsi model dengan mengirimkan $id.
            'mahasiswa' => $this->Mahasiswa_model->getMahasiswaById($id)
        ];
        // data array selesai.
        // mengatur validasi dari inputan "name" mulai.
        $this->form_validation->set_rules('nim', 'NIM', 'required|numeric|max_length[12]');
        $this->form_validation->set_rules('nama', 'Nama', 'required|max_length[32]');
        $this->form_validation->set_rules('alamat', 'Alamat', 'required|max_length[32]');
        $this->form_validation->set_rules('telp', 'Telepon', 'required|numeric|max_length[13]');
        // mengatur validasi dari inputan "name" selesai.
        // pengkodisian jika validasi berjalan mulai.
        // jika validasi salah maka.
        if ($this->form_validation->run() == FALSE) {
            // jika validasi salah maka.
            // load view.
            // mengandung array $data.
            $this->load->view('template/header', $data);
            $this->load->view('template/sidebar');
            // mengandung array $data
            $this->load->view('mahasiswa/v_ubah', $data);
            $this->load->view('template/footer');
        }
        // jika benar maka.
        else {
            // load fungsi model.
            $this->Mahasiswa_model->ubahDataMahasiswa($id);
            // membuat session flash.
            $this->session->set_flashdata('flash', 'Data berhasil diubah!');
            // kembali ke controller mahasiswa.
            redirect('mahasiswa');
        }
        // pengkodisian jika validasi berjalan selesai.
    }

    // fungsi hapus berdasarkan paramater $id.
    public function delete($id)
    {
        // load fungsi model dengan mengirimkan $id.
        $this->Mahasiswa_model->hapusDataMahasiswa($id);
        // membuat session flash.
        $this->session->set_flashdata('flash', 'Data berhasil dihapus!');
        // kembali ke controller mahasiswa.
        redirect('mahasiswa');
    }

    // fungsi riwayat.
    public function log()
    {
        // data array mulai.
        $data = [
            // judul tab.
            'judul' => "Riwayat No. HP Mahasiswa",
            // load fungsi model.
            'mahasiswa' => $this->Mahasiswa_model->getLog()
        ];
        // data array selesai.
        // load view mulai.
        // mengandung data array.
        $this->load->view('template/header', $data);
        $this->load->view('template/sidebar');
        // mengandung data array.
        $this->load->view('mahasiswa/v_log', $data);
        $this->load->view('template/footer');
        // load view selesai.
    }

    public function export_tampil_semua()
    {
        // load fungsi model.
        $data['mahasiswa'] = $this->Mahasiswa_model->getLog();

        // load plugin PHPExcel mulai.
        require(APPPATH . 'PHPExcel-1.8/Classes/PHPExcel.php');
        require(APPPATH . 'PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php');
        // load plugin PHPExcel selesai.

        // Panggil class PHPExcel
        $objPHPExcel = new PHPExcel();

        // membuat properties file mulai.
        $objPHPExcel->getProperties()->setCreator("Arky");
        $objPHPExcel->getProperties()->setLastModifiedBy("Arky");
        $objPHPExcel->getProperties()->setTitle("Data Riwayat Ganti Nomer Mahasiswa");
        $objPHPExcel->getProperties()->setSubject("");
        $objPHPExcel->getProperties()->setDescription("");
        // membuat properties file selesai.

        // Buat sebuah variabel untuk menampung pengaturan style dari header tabel.
        $style_col = array(
            // Set font nya jadi bold.
            'font' => array('bold' => true),
            'alignment' => array(
                // Set text jadi ditengah secara horizontal (center)
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                // Set text jadi di tengah secara vertical (middle)
                'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
            ),
            'borders' => array(
                // set border atas, bawah, kanan, dan kiri dengan garis tipis mulai.
                'top' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'right' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'bottom' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'left' => array('style'  => PHPExcel_Style_Border::BORDER_THIN)
                // set border atas, bawah, kanan, dan kiri dengan garis tipis selesai.
            )
        );

        // Buat sebuah variabel untuk menampung pengaturan style dari isi tabel.
        $style_row = array(
            'alignment' => array(
                // Set text jadi ditengah secara horizontal (center)
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                // Set text jadi di tengah secara vertical (middle)
                'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
            ),
            'borders' => array(
                // set border atas, bawah, kanan, dan kiri dengan garis tipis mulai.
                'top' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'right' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'bottom' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'left' => array('style'  => PHPExcel_Style_Border::BORDER_THIN)
                // set border atas, bawah, kanan, dan kiri dengan garis tipis selesai.
            )
        );

        // mengatur sheet yang aktif.
        $objPHPExcel->setActiveSheetIndex(0); // sheet yang pertama = (0).
        // merge cells A1 Sampai H1
        $objPHPExcel->getActiveSheet()->mergeCells('A1:H1');
        // membuat font A1 bold
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(TRUE);
        // ukuran huruf A1
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(14);
        // Set text jadi ditengah secara horizontal (center)
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);



        // mengatur nama-nama field mulai.
        $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Riwayat Nomer Telepon Mahasiswa');
        $objPHPExcel->getActiveSheet()->setCellValue('A3', 'No');
        $objPHPExcel->getActiveSheet()->setCellValue('B3', 'NIM');
        $objPHPExcel->getActiveSheet()->setCellValue('C3', 'Nama');
        $objPHPExcel->getActiveSheet()->setCellValue('D3', 'Alamat');
        $objPHPExcel->getActiveSheet()->setCellValue('E3', 'Jenis Kelamin');
        $objPHPExcel->getActiveSheet()->setCellValue('F3', 'Nomer Hape Lama');
        $objPHPExcel->getActiveSheet()->setCellValue('G3', 'Nomer Hape Baru');
        $objPHPExcel->getActiveSheet()->setCellValue('H3', 'Tanggal Diubah');
        // mengatur nama-nama field selesai.
        // style kolom mulai.
        $objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('B3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('D3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('E3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('F3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('G3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('H3')->applyFromArray($style_col);
        // style kolom selesai.
        // memasukkan data mulai.
        // variabel untuk baris.
        $baris = 4;
        // variabel untuk nomer.
        $x = 1;

        //looping data (mulai)
        foreach ($data['mahasiswa'] as $data) {
            // memanggil data tabel.
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $baris, $x);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $baris, $data['nim']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $baris, $data['nama']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $baris, $data['alamat']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $baris, $data['jk']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $baris, $data['telp_lama']);
            $objPHPExcel->getActiveSheet()->setCellValue('G' . $baris, $data['telp_baru']);
            $objPHPExcel->getActiveSheet()->setCellValue('H' . $baris, $data['tgl_diubah']);
            // style row.
            $objPHPExcel->getActiveSheet()->getStyle('A' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('B' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('C' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('D' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('E' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('F' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('G' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('H' . $baris)->applyFromArray($style_row);
            // perulangan
            $x++;
            $baris++;
        }
        //looping data (selesai)
        // memasukkan data (selesai)
        // set lebar kolom otomatis mulai.
        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
        // set lebar kolom otomatis selesai.
        // Set tinggi semua kolom menjadi auto (mengikuti height isi dari kolommnya, jadi otomatis)
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);
        // membuat nama file excel
        $filename = "Riwayat Nomer Telepon Mahasiswa " . date('Y-m-d-H-i-s') . ".xlsx";
        // set judul aktif sheet
        $objPHPExcel->getActiveSheet()->setTitle("Riwayat Nomer Telepon Mahasiswa");

        // proses file excel mulai.
        header("Content-Type: apllication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $writer->save('php://output');
        exit;
        // proses file excel selesai.
    }

    public function export_tampil_mhs()
    {
        // load fungsi model.
        $data['mahasiswa'] = $this->Mahasiswa_model->getAllMahasiswa();

        // load plugin PHPExcel mulai.
        require(APPPATH . 'PHPExcel-1.8/Classes/PHPExcel.php');
        require(APPPATH . 'PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php');
        // load plugin PHPExcel selesai.

        // Panggil class PHPExcel
        $objPHPExcel = new PHPExcel();

        // membuat properties file mulai.
        $objPHPExcel->getProperties()->setCreator("Arky");
        $objPHPExcel->getProperties()->setLastModifiedBy("Arky");
        $objPHPExcel->getProperties()->setTitle("Data Mahasiswa");
        $objPHPExcel->getProperties()->setSubject("");
        $objPHPExcel->getProperties()->setDescription("");
        // membuat properties file selesai.

        // Buat sebuah variabel untuk menampung pengaturan style dari header tabel.
        $style_col = array(
            // Set font nya jadi bold.
            'font' => array('bold' => true),
            'alignment' => array(
                // Set text jadi ditengah secara horizontal (center)
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                // Set text jadi di tengah secara vertical (middle)
                'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
            ),
            'borders' => array(
                // set border atas, bawah, kanan, dan kiri dengan garis tipis mulai.
                'top' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'right' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'bottom' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'left' => array('style'  => PHPExcel_Style_Border::BORDER_THIN)
                // set border atas, bawah, kanan, dan kiri dengan garis tipis selesai.
            )
        );

        // Buat sebuah variabel untuk menampung pengaturan style dari isi tabel.
        $style_row = array(
            'alignment' => array(
                // Set text jadi ditengah secara horizontal (center)
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                // Set text jadi di tengah secara vertical (middle)
                'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER
            ),
            'borders' => array(
                // set border atas, bawah, kanan, dan kiri dengan garis tipis mulai.
                'top' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'right' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'bottom' => array('style'  => PHPExcel_Style_Border::BORDER_THIN),
                'left' => array('style'  => PHPExcel_Style_Border::BORDER_THIN)
                // set border atas, bawah, kanan, dan kiri dengan garis tipis selesai.
            )
        );

        // mengatur sheet yang aktif.
        $objPHPExcel->setActiveSheetIndex(0); // sheet yang pertama = (0).
        // merge cells A1 Sampai H1
        $objPHPExcel->getActiveSheet()->mergeCells('A1:F1');
        // membuat font A1 bold
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(TRUE);
        // ukuran huruf A1
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(14);
        // Set text jadi ditengah secara horizontal (center)
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        // mengatur nama-nama field mulai.
        $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Data Mahasiswa');
        $objPHPExcel->getActiveSheet()->setCellValue('A3', 'No');
        $objPHPExcel->getActiveSheet()->setCellValue('B3', 'NIM');
        $objPHPExcel->getActiveSheet()->setCellValue('C3', 'Nama');
        $objPHPExcel->getActiveSheet()->setCellValue('D3', 'Alamat');
        $objPHPExcel->getActiveSheet()->setCellValue('E3', 'Jenis Kelamin');
        $objPHPExcel->getActiveSheet()->setCellValue('F3', 'Nomer Telepon');
        // mengatur nama-nama field selesai.

        // style kolom mulai.
        $objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('B3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('C3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('D3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('E3')->applyFromArray($style_col);
        $objPHPExcel->getActiveSheet()->getStyle('F3')->applyFromArray($style_col);
        // style kolom selesai.

        // memasukkan data mulai.
        // variabel untuk baris.
        $baris = 4;
        // variabel untuk nomer.
        $x = 1;

        //looping data (mulai)
        foreach ($data['mahasiswa'] as $data) {
            // memanggil data tabel.
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $baris, $x);
            $objPHPExcel->getActiveSheet()->setCellValue('B' . $baris, $data['nim']);
            $objPHPExcel->getActiveSheet()->setCellValue('C' . $baris, $data['nama']);
            $objPHPExcel->getActiveSheet()->setCellValue('D' . $baris, $data['alamat']);
            $objPHPExcel->getActiveSheet()->setCellValue('E' . $baris, $data['jk']);
            $objPHPExcel->getActiveSheet()->setCellValue('F' . $baris, $data['telp']);
            // style row.
            $objPHPExcel->getActiveSheet()->getStyle('A' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('B' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('C' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('D' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('E' . $baris)->applyFromArray($style_row);
            $objPHPExcel->getActiveSheet()->getStyle('F' . $baris)->applyFromArray($style_row);
            // perulangan
            $x++;
            $baris++;
        }
        //looping data (selesai)
        // memasukkan data (selesai)
        // set lebar kolom otomatis mulai.
        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
        // set lebar kolom otomatis selesai.
        // Set tinggi semua kolom menjadi auto (mengikuti height isi dari kolommnya, jadi otomatis)
        $objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(-1);
        // membuat nama file excel
        $filename = "Data Mahasiswa " . date('Y-m-d-H-i-s') . ".xlsx";
        // set judul aktif sheet
        $objPHPExcel->getActiveSheet()->setTitle("Data Mahasiswa");

        // proses file excel mulai.
        header("Content-Type: apllication/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $writer->save('php://output');
        exit;
        // proses file excel selesai.
    }
}

/* End of file Mahasiswa.php */
