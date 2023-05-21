<?php
require 'vendor/autoload.php'; // Memuat pustaka PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Koneksi ke database
$host = 'localhost'; // Ganti dengan host basis data Anda
$dbname = 'pendaftaran_siswa'; // Ganti dengan nama database Anda
$username = 'root'; // Ganti dengan nama pengguna basis data Anda
$password = ''; // Ganti dengan kata sandi basis data Anda

try {
    $db = new PDO("mysql:host=$host;dbname=$dbname", $username, $password);
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
} catch (PDOException $e) {
    die("Koneksi database gagal: " . $e->getMessage());
}

// Fungsi untuk mengekspor data ke file Excel
function exportToExcel($data)
{
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Judul kolom
    $sheet->setCellValue('A1', 'Nomor');
    $sheet->setCellValue('B1', 'Nama Lengkap');
    $sheet->setCellValue('C1', 'Jenis Kelamin');
    $sheet->setCellValue('D1', 'NISN');
    $sheet->setCellValue('E1', 'NIK');
    $sheet->setCellValue('F1', 'Tempat Lahir');
    $sheet->setCellValue('G1', 'Tanggal Lahir');
    $sheet->setCellValue('H1', 'Agama');
    $sheet->setCellValue('I1', 'Berkebutuhan Khusus');
    $sheet->setCellValue('J1', 'Alamat');
    $sheet->setCellValue('K1', 'Tempat Tinggal');
    $sheet->setCellValue('L1', 'Transportasi');
    $sheet->setCellValue('M1', 'Nomor Handphone');
    $sheet->setCellValue('N1', 'Nomor Telepon');
    $sheet->setCellValue('O1', 'Email');

    // Mengisi data siswa
    $row = 2;
    $no = 1;
    foreach ($data as $siswa) {
        $sheet->setCellValue('A' . $row, $no++);
        $sheet->setCellValue('B' . $row, $siswa['nama_lengkap']);
        $sheet->setCellValue('C' . $row, $siswa['jenis_kelamin']);
        $sheet->setCellValue('D' . $row, $siswa['nisn']);
        $sheet->setCellValue('E' . $row, $siswa['nik']);
        $sheet->setCellValue('F' . $row, $siswa['tempat_lahir']);
        $sheet->setCellValue('G' . $row, $siswa['tanggal_lahir']);
        $sheet->setCellValue('H' . $row, $siswa['agama']);
        $sheet->setCellValue('I' . $row, $siswa['berkebutuhan_khusus']);
        $sheet->setCellValue('J' . $row, $siswa['alamat']);
        $sheet->setCellValue('K' . $row, $siswa['tempat_tinggal']);
        $sheet->setCellValue('L' . $row, $siswa['transportasi']);
        $sheet->setCellValue('M' . $row, $siswa['nomor_handphone']);
        $sheet->setCellValue('N' . $row, $siswa['nomor_telepon']);
        $sheet->setCellValue('O' . $row, $siswa['email']);
        $row++;
    }

    // Menyimpan file Excel
    $writer = new Xlsx($spreadsheet);
    $filename = 'data_siswa.xlsx';
    $writer->save($filename);

    return $filename;
}

$namaLengkap = '';
$jenisKelamin = '';
$nisn = '';
$nik = '';
$tempatLahir = '';
$tanggalLahir = '';
$agama = '';
$berkebutuhanKhusus = '';
$alamat = '';
$tempatTinggal = '';
$transportasi = '';
$nomorHandphone = '';
$nomorTelepon = '';
$email = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Mengambil data dari form
    $namaLengkap = $_POST['nama_lengkap'];
    $jenisKelamin = $_POST['jenis_kelamin'];
    $nisn = $_POST['nisn'];
    $nik = $_POST['nik'];
    $tempatLahir = $_POST['tempat_lahir'];
    $tanggalLahir = $_POST['tanggal_lahir'];
    $agama = $_POST['agama'];
    $berkebutuhanKhusus = $_POST['berkebutuhan_khusus'];
    $alamat = $_POST['alamat'];
    $tempatTinggal = $_POST['tempat_tinggal'];
    $transportasi = $_POST['transportasi'];
    $nomorHandphone = $_POST['nomor_handphone'];
    $nomorTelepon = $_POST['nomor_telepon'];
    $email = $_POST['email'];

    // Memasukkan data ke database
    try {
        $stmt = $db->prepare('INSERT INTO siswa (nama_lengkap, jenis_kelamin, nisn, nik, tempat_lahir, tanggal_lahir, agama, berkebutuhan_khusus, alamat, tempat_tinggal, transportasi, nomor_handphone, nomor_telepon, email) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)');
        $stmt->execute([$namaLengkap, $jenisKelamin, $nisn, $nik, $tempatLahir, $tanggalLahir, $agama, $berkebutuhanKhusus, $alamat, $tempatTinggal, $transportasi, $nomorHandphone, $nomorTelepon, $email]);

        // Mengambil semua data siswa dari database
        $stmt = $db->query('SELECT * FROM siswa');
        $dataSiswa = $stmt->fetchAll(PDO::FETCH_ASSOC);

        // Mengekspor data ke file Excel
        $excelFile = exportToExcel($dataSiswa);

        // Menampilkan pesan ke pengguna
        $_SESSION['message'] = "Pendaftaran berhasil! Data telah diekspor ke file Excel dengan nama: $excelFile";

        // Mengalihkan pengguna kembali ke halaman form
        header("Location: formpendaftaran.php");
        exit();
    } catch (PDOException $e) {
        die("Pendaftaran gagal: " . $e->getMessage());
    }
}

// Menghapus pesan setelah refresh
if (isset($_SESSION['message'])) {
    unset($_SESSION['message']);
}
?>

<!DOCTYPE html>
<html>

<head>
    <title>Form Pendaftaran Siswa</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        label {
            display: inline-block;
            width: 120px;
            text-align: right;
        }

        button[type="submit"] {
            margin-left: 120px;
        }

        .tengah {
            width: 150%;
            margin-left: 180px;
            height: auto;
        }
    </style>
</head>

<body>
    <div class="tengah">
        <div class="row">
            <div class="col-md-6">
                <div class="card">
                    <div class="card-header">
                        <h1 style="text-align: center;">Formulir Peserta Baru</h1><br>
                        <form method="POST" action="">
                            <div class="form-group row">
                                <label for="nama_lengkap" class="col-sm-2 col-form-label">Nama Lengkap:</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control" id="nama_lengkap" name="nama_lengkap"
                                        required>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="jenis_kelamin" class="col-sm-2 col-form-label">Jenis Kelamin:</label>
                                <div class="col-sm-10">
                                    <select class="form-control" id="jenis_kelamin" name="jenis_kelamin" required>
                                        <option value="">Pilih</option>
                                        <option value="Laki-laki">Laki-laki</option>
                                        <option value="Perempuan">Perempuan</option>
                                    </select>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="nisn" class="col-sm-2 col-form-label">NISN:</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control" id="nisn" name="nisn" required>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="nik" class="col-sm-2 col-form-label">NIK:</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control" id="nik" name="nik" required>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="tempat_lahir" class="col-sm-2 col-form-label">Tempat Lahir:</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control" id="tempat_lahir" name="tempat_lahir"
                                        required>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="tanggal_lahir" class="col-sm-2 col-form-label">Tanggal Lahir:</label>
                                <div class="col-sm-10">
                                    <input type="date" class="form-control" id="tanggal_lahir" name="tanggal_lahir"
                                        required>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="agama" class="col-sm-2 col-form-label">Agama:</label>
                                <div class="col-sm-10">
                                    <select class="form-control" id="agama" name="agama" required>
                                        <option selected disabled value="">Pilih</option>
                                        <option value="Islam">Islam</option>
										<option value="Kristen">Kristen</option>
										<option value="Katolik">Katolik</option>
										<option value="Hindu">Hindu</option>
										<option value="Buddha">Buddha</option>
										<option value="Konghucu">Konghucu</option>
                                    </select>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="berkebutuhan_khusus" class="col-sm-2 col-form-label">Berkebutuhan Khusus:</label>
                                <div class="col-sm-10">
                                    <select class="form-control" id="berkebutuhan_khusus" name="berkebutuhan_khusus" required>
                                    <option value="Tidak">Tidak</option>
										<option value="Netra">Netra</option>
										<option value="Rungu">Rungu</option>
										<option value="Grahita Sedang">Grahita Sedang</option>
										<option value="Daksa Ringan">Daksa Ringan</option>
										<option value="Daksa Sedang">Daksa Sedang</option>
										<option value="Laras">Laras</option>
										<option value="Wicara">Wicara</option>
										<option value="Tuna Ganda">Tuna Ganda</option>
										<option value="Hiper Aktif">Hiper Aktif</option>
										<option value="Cerdas Istimewa">Cerdas Istimewa</option>
										<option value="Bakat Istimewa">Bakat Istimewa</option>
										<option value="Kesulitan Belajar">Kesulitan Belajar</option>
										<option value="Narkoba">Narkoba</option>
										<option value="Indigo">Indigo</option>
										<option value="Down Sindrome">Down Sindrome</option>
										<option value="Autis">Autis</option>
										<option value="Lainnya">Lainnya</option>
                                    </select>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="alamat" class="col-sm-2 col-form-label">Alamat:</label>
                                <div class="col-sm-10">
                                    <textarea class="form-control" id="alamat" name="alamat" required></textarea>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="tempat_tinggal" class="col-sm-2 col-form-label">Tempat Tinggal:</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control" id="tempat_tinggal" name="tempat_tinggal"
                                        required>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="transportasi" class="col-sm-2 col-form-label">Transportasi:</label>
                                <div class="col-sm-10">
                                    <select class="form-control" id="transportasi" name="transportasi" required>
                                    <option selected disabled value="">-- Pilih --</option>
										<option value="Jalan kaki">Jalan kaki</option>
										<option value="Kendaraan Pribadi">Kendaraan Pribadi</option>
										<option value="Kendaraan Umum">Kendaraan Umum</option>
										<option value="Jemputan Sekolah">Jemputan Sekolah</option>
										<option value="Kereta Api">Kereta Api</option>
										<option value="Ojek">Ojek</option>
										<option value="Dokar/Becak">Dokar/Becak</option>
										<option value="Perahu Penyebrangan">Perahu Penyebrangan</option>
										<option value="Lainnya">Lainnya</option>
                                    </select>
                                </div>
                            </div>


                            <div class="form-group row">
                                <label for="nomor_handphone" class="col-sm-2 col-form-label">Nomor Handphone:</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control" id="nomor_handphone" name="nomor_handphone"
                                        required>
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="nomor_telepon" class="col-sm-2 col-form-label">Nomor Telepon:</label>
                                <div class="col-sm-10">
                                    <input type="text" class="form-control" id="nomor_telepon" name="nomor_telepon">
                                </div>
                            </div>

                            <div class="form-group row">
                                <label for="email" class="col-sm-2 col-form-label">Email:</label>
                                <div class="col-sm-10">
                                    <input type="email" class="form-control" id="email" name="email" required>
                                </div>
                            </div>

                            <div class="form-group row">
                                <div class="col-sm-10 offset-sm-2">
                                    <button type="submit" class="btn btn-primary">Daftar</button>
                                </div>
                            </div>
                        </form>
                    </div>

                    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
</body>

</html>