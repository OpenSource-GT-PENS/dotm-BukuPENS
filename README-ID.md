# Instalasi Template Word (.dotm)

README ini menyediakan instruksi untuk menginstal file template Word (.dotm) ke folder Custom Templates Microsoft Office Anda.

## Prasyarat

- Microsoft Word (versi terbaru apa saja)
- File template .dotm dari repositori ini

## Instruksi Instalasi

### Windows

1. Unduh file .dotm dari repositori ini
2. Temukan folder Custom Templates Microsoft Word Anda. Lokasi default adalah:
   `C:\Users\NamaUser\AppData\Roaming\Microsoft\Templates`
3. Salin file .dotm ke folder ini
4. Restart Microsoft Word jika sedang berjalan

### MacOS

1. Unduh file .dotm dari repositori ini
2. Temukan folder Custom Templates Microsoft Word Anda. Lokasi default adalah:
   `~/Library/Group Containers/UBF8T346G9.Office/User Content/Templates`
3. Jika folder Templates tidak ada, Anda mungkin perlu membuatnya
4. Salin file .dotm ke folder ini
5. Restart Microsoft Word jika sedang berjalan

### Metode Instalasi Alternatif

1. Buka Microsoft Word
2. Klik **File** > **Options** (Windows) atau **Word** > **Preferences** (Mac)
3. Di Windows: Pilih **Advanced** dari sidebar kiri, scroll ke bawah dan klik **File Locations**
   Di Mac: Klik **File Locations**
4. Cari "User templates" dan klik **Modify** (Windows) atau **Change** (Mac)
5. Ini akan membuka folder tempat Anda harus menempatkan file .dotm
6. Salin file .dotm ke folder ini
7. Restart Microsoft Word

## Menggunakan Template

1. Buka Microsoft Word
2. Klik **File** > **New** (Windows) atau **File** > **New from Template** (Mac)
3. Klik tab **Personal** atau **Custom** (tergantung versi Word Anda)
4. Pilih template dari opsi yang tersedia

## Mengimpor Template ke Dokumen yang Sudah Ada

### Metode 1: Menggunakan Organizer

1. Buka dokumen Word yang sudah ada di mana Anda ingin mengimpor template
2. Tekan **Alt+F11** (Windows) atau **Option+F11** (Mac) untuk membuka Visual Basic Editor
3. Di VBE, klik **Tools** > **Organizer**
4. Di dialog Organizer, klik tab **Styles**
5. Di sisi kanan, klik **Close File** kemudian **Open File**
6. Telusuri dan pilih file template .dotm
7. Pilih gaya yang ingin Anda salin dari template (sisi kanan) ke dokumen Anda (sisi kiri)
8. Klik **Copy** untuk mengimpor gaya yang dipilih
9. Tutup Organizer dan VBE

### Metode 2: Melampirkan Template

1. Buka dokumen Word yang sudah ada
2. Klik **File** > **Options** (Windows) atau **Word** > **Preferences** (Mac)
3. Pilih **Add-ins** (Windows) atau **Templates and Add-ins** (Mac)
4. Klik tab **Templates** (Windows) atau pastikan Anda berada di tab Templates (Mac)
5. Klik tombol **Attach**
6. Telusuri dan pilih file template .dotm
7. Centang kotak "Automatically update document styles" jika Anda ingin gaya diperbarui otomatis
8. Klik **OK**

### Metode 3: Salin dan Tempel Gaya

1. Buat dokumen baru berdasarkan template .dotm
2. Tekan **Ctrl+Alt+Shift+S** (Windows) atau **Command+Option+Shift+S** (Mac) untuk membuka panel Styles
3. Untuk setiap gaya yang ingin Anda impor, klik kanan dan pilih **Select All X Instance(s)** di mana X adalah jumlah instance
4. Salin konten yang dipilih (**Ctrl+C** / **Command+C**)
5. Pergi ke dokumen yang sudah ada dan tempel (**Ctrl+V** / **Command+V**)
6. Gaya akan diimpor bersama dengan konten
7. Anda kemudian dapat memodifikasi atau menghapus konten yang ditempel sambil mempertahankan gaya yang diimpor

## Pemecahan Masalah

Jika Anda tidak melihat template setelah instalasi:
- Pastikan Anda telah menempatkan file .dotm di folder yang benar
- Periksa apakah Anda memiliki izin yang cukup untuk mengakses folder templates
- Restart Microsoft Word sepenuhnya
- Di Mac, Anda mungkin perlu memeriksa apakah folder Library tersembunyi (tekan Cmd+Shift+. untuk menampilkan file tersembunyi)

## Dukungan

Jika Anda mengalami masalah dengan instalasi atau penggunaan template, silakan buka issue di repositori ini.

## Tentang Template BukuPENS

Template ini dirancang khusus untuk membantu mahasiswa dan dosen Politeknik Elektronika Negeri Surabaya (PENS) dalam membuat dokumen akademik yang sesuai dengan standar institusi. Template ini mencakup:

- Format dokumen yang sesuai standar PENS
- Gaya penulisan yang konsisten
- Template untuk berbagai bagian dokumen akademik
- Panduan penulisan yang terintegrasi

Untuk informasi lebih lanjut tentang penggunaan template ini, silakan merujuk ke dokumentasi yang disertakan dalam template atau hubungi tim pengembang melalui repositori ini.
