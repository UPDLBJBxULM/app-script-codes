# Create PDF dan Google Sheets Cloud Function Integration

## Deskripsi

Fungsi-fungsi ini digunakan untuk membuat dropdown, mengisi cell kosong dengan data yang diinginkan, melakukan perekapan, generate PDF dan mengirimkan data dari Google Sheets ke Cloud Function setiap kali sheet yang ditentukan diedit. Data yang dikirimkan mencakup seluruh isi sheet (termasuk header) dan informasi terkait seperti nama sheet, timestamp, dan URL spreadsheet untuk disimpan dalam Database BigQuery.

## Fungsi Utama pada Create PDF dan Google Sheets Cloud Function Integration

- **telahDiedit(e)**: Fungsi utama yang dipicu setiap kali ada perubahan pada sheet. Fungsi ini akan memeriksa apakah sheet yang sedang aktif ada dalam daftar sheet yang dipantau dan mengirimkan data ke Cloud Function yang sesuai jika kondisi terpenuhi dan setelahnya akan dimasukkan ke Database BigQuery.
  
  **Cara Kerja**:  
  Fungsi ini memeriksa apakah sheet yang sedang aktif terdaftar dalam konfigurasi `sheetConfigs`. Jika iya, fungsi akan mengambil data dari sheet, memeriksa apakah sel A1 valid, kemudian mengirimkan data ke Cloud Function yang sesuai. Jika pengiriman berhasil, data juga akan dimasukkan ke BigQuery untuk perekaman lebih lanjut.

- **isCellError(value)**: Fungsi pembantu untuk memeriksa apakah nilai sel kosong atau error, jika ada maka Cloud Functions tidak akan dipanggil untuk melakukan proses pengolahan data.
  
  **Cara Kerja**:  
  Fungsi ini digunakan untuk memvalidasi nilai dalam sel, terutama di A1. Jika nilai sel tersebut kosong atau error, fungsi ini akan mengembalikan nilai `true` dan mencegah pengiriman data ke Cloud Function. Hal ini memastikan bahwa data yang dikirimkan selalu lengkap dan valid.

- **updateDropdownList()**: Fungsi ini untuk membuat dropdown berdasarkan data yang diambil dari spreadsheet yang menjadi sumber datanya.
  
  **Cara Kerja**:  
  Fungsi ini akan mencari sumber data (biasanya sebuah sheet atau kolom tertentu) dan mengambil nilai-nilai yang ada. Kemudian, fungsi ini akan memperbarui daftar pilihan dropdown pada sheet target agar mencerminkan data terbaru dari sumber.

- **updateData(selectedValue)**: Fungsi ini akan mengisi cell kosong sesuai kebutuhan yang telah ditentukan dengan berdasarkan data dropdown yang telah dipilih.
  
  **Cara Kerja**:  
  Ketika pengguna memilih nilai dari dropdown yang telah diperbarui, fungsi ini akan mengisi cell kosong di sheet dengan nilai terkait yang sudah ditentukan sebelumnya berdasarkan pilihan dropdown. Fungsi ini digunakan untuk memastikan bahwa data terisi dengan benar sesuai dengan opsi yang tersedia.

- **manualUpdate()**: Fungsi untuk merefresh data dropdown yang belum terupdate.
  
  **Cara Kerja**:  
  Jika ada perubahan pada sumber data atau jika dropdown perlu diperbarui secara manual, fungsi ini akan dipanggil untuk memperbarui kembali pilihan yang tersedia pada dropdown, memastikan bahwa data yang ada di dropdown selalu terbaru dan relevan.

- **clearTargetCells()**: Fungsi membersihkan cell yang ditargetkan dari data-data sebelumnya, sehingga hanya ada template kosong yang siap diisi data kembali.
  
  **Cara Kerja**:  
  Fungsi ini digunakan untuk menghapus data yang sudah ada pada cell target, sehingga memberikan ruang kosong untuk pengisian data baru. Biasanya dipanggil sebelum proses pengisian data baru dilakukan agar tidak ada data lama yang mengganggu.

- **fillZonasi(e)**: Fungsi ini akan memberikan warna terhadap cell target secara otomatis berdasarkan proses algoritma yang telah ditentukan.
  
  **Cara Kerja**:  
  Fungsi ini akan menganalisis nilai di cell tertentu dan kemudian memberikan warna atau format tertentu pada cell tersebut berdasarkan algoritma yang sudah diset sebelumnya. Ini bisa digunakan, misalnya, untuk memberi warna pada cell berdasarkan kriteria tertentu seperti status atau kategori data.

- **generatePDFFinal()**: Fungsi untuk mengenerate PDF berdasarkan template spreadsheet yang telah berisi data-data.
  
  **Cara Kerja**:  
  Fungsi ini akan mengambil data dari sheet yang telah diisi dan memformatnya sesuai dengan template yang telah ditentukan. Setelah itu, fungsi akan mengonversi sheet tersebut menjadi PDF dan menyimpannya, atau mengirimkan hasilnya ke tempat yang diinginkan (misalnya, email atau folder penyimpanan).

- **rekapData()**: Fungsi yang merekap data-data yang terisi pada template untuk direkap ke worksheet tujuan, fungsi ini berjalan setelah fungsi `generatePDFFinal` selesai berjalan.
  
  **Cara Kerja**:  
  Setelah PDF dihasilkan, fungsi ini akan merekap data yang telah diisi pada sheet dan menyalinnya ke worksheet lain (biasanya worksheet ringkasan atau laporan). Fungsi ini digunakan untuk mencatat dan menyimpan hasil pengolahan data ke tempat yang tepat.

## Cara Kerja Umum Google Sheets Cloud Function Integration

1. **Pemantauan Sheet**: Daftar sheet yang dipantau dan URL Cloud Function yang sesuai disimpan dalam objek `sheetConfigs`.
2. **Validasi Sheet dan Sel**: 
   - Sheet yang aktif harus terdaftar di dalam `sheetConfigs`.
   - Sel A1 harus memiliki nilai yang valid (tidak kosong atau error).
3. **Pengambilan Data**: 
   - Data diambil dari sheet yang aktif, mulai dari baris kedua (baris pertama dianggap sebagai header).
   - Data yang dikirimkan akan mencakup header dan hanya baris yang memiliki nilai (baris kosong akan disaring).
4. **Pengiriman Data**: 
   - Data dikirim ke URL Cloud Function yang terkait dengan sheet yang aktif melalui request HTTP POST.
   - Payload yang dikirimkan mencakup nama sheet, header, data, timestamp, dan informasi spreadsheet.

## Struktur Data yang Dikirimkan Ke Cloud Functions

Payload yang dikirimkan ke Cloud Function adalah objek JSON dengan struktur sebagai berikut:

```json
{
  "sheetName": "nama_sheet_aktif",
  "headers": ["header1", "header2", ...],
  "data": [
    ["data1", "data2", ...],
    ["data3", "data4", ...]
  ],
  "timestamp": "2024-11-11T12:34:56.789Z",
  "spreadsheetId": "id_spreadsheet",
  "spreadsheetUrl": "https://docs.google.com/spreadsheets/d/id_spreadsheet/edit"
}
