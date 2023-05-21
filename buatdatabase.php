<?php
$servername = "localhost";
$username = "root";
$password = "";

// Membuat koneksi
$conn = new mysqli($servername, $username, $password);

// Memeriksa koneksi
if ($conn->connect_error) {
    die("Koneksi gagal: " . $conn->connect_error);
}

// Membuat database
$sql = "CREATE DATABASE IF NOT EXISTS db_siswa DEFAULT CHARACTER SET latin1 COLLATE latin1_swedish_ci";
if ($conn->query($sql) === true) {
    echo "Database 'db_siswa' berhasil dibuat atau sudah ada";
} else {
    echo "Error saat membuat database: " . $conn->error;
}

// Menggunakan database
$conn->select_db("db_siswa");

// Membuat tabel
$sql = "CREATE TABLE IF NOT EXISTS tb_siswa (
    id_siswa int(11) NOT NULL AUTO_INCREMENT,
    nama varchar(255) NOT NULL,
    kelas varchar(100) NOT NULL,
    alamat varchar(255) NOT NULL,
    PRIMARY KEY (id_siswa)
) ENGINE=InnoDB DEFAULT CHARSET=latin1";
if ($conn->query($sql) === true) {
    echo "Tabel 'tb_siswa' berhasil dibuat atau sudah ada";
} else {
    echo "Error saat membuat tabel: " . $conn->error;
}

// Memasukkan data ke dalam tabel
$sql = "INSERT INTO tb_siswa (id_siswa, nama, kelas, alamat) VALUES
    (1, 'Budi Susanto', '1MM3', 'Sedati Gede'),
    (2, 'Dita Anggraini', '1MM2', 'Rungkut'),
    (3, 'Riska Nur Aini', '3MM1', 'Wonocolo')";
if ($conn->query($sql) === true) {
    echo "Data berhasil dimasukkan ke dalam tabel 'tb_siswa'";
} else {
    echo "Error saat memasukkan data ke dalam tabel: " . $conn
    ->error;
}

// Menambahkan indeks pada tabel
$sql = "ALTER TABLE tb_siswa ADD PRIMARY KEY (id_siswa)";
if ($conn->query($sql) === true) {
    echo "Indeks berhasil ditambahkan pada tabel 'tb_siswa'";
} else {
    echo "Error saat menambahkan indeks: " . $conn->error;
}

// Mengatur AUTO_INCREMENT untuk tabel
$sql = "ALTER TABLE tb_siswa MODIFY id_siswa int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=4";
if ($conn->query($sql) === true) {
    echo "AUTO_INCREMENT berhasil diatur pada tabel 'tb_siswa'";
} else {
    echo "Error saat mengatur AUTO_INCREMENT: " . $conn->error;
}

// Commit perubahan
$conn->commit();

// Menutup koneksi database
$conn->close();
?>