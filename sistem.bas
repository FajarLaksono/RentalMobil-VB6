Attribute VB_Name = "sistem"
Public userNIK As String
Public getName As String
Public getJabatan As String
Public getPhoto As String
'di atas tempat menyimpan informasi petugas yang sudah melakukan login, semua informasi disimpan sementara di variabel di atas.
'dapat mempersingkat penulisan kedepanya. setiap form yang membutuhkan informasi untuk ditampilkan tidak harus selalu konek ke database atau menggunakan adodc hanya untuk mengambil informasi
Public isEditing As Boolean '1 = edit mode | 0 = peninjauan
Public currRecord As String 'berfungsi untuk menyimpan kunci. data apa yang akan di tampilkan pada form berikutnya

Public Function msgTitle() As String
    msgTitle = "Rental Mobil Purwokerto 0.5"
End Function

Public Function connectToDatabeseRentalMobil() As String
    'semua ConnectionString pada semua form akan mengambil alamat dari function ini
    connectToDatabeseRentalMobil = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\databases\databaseRentalMobil2003.mdb;Persist Security Info=False"
    'Digunakan untuk mengatur database dengan koding bukan GUI yang ada.
    'pengaturan yang kita atur secara manual (gui) kemungkinan hanya akan mencatat alamat lengkap dari drive computer yang digunakan untuk menyimpan hingga lokasi database tersebut
    'dan itu sebabnya kami memindahkan segala file seperti foto dan database ke satu forder yang dapat di jangkau(didalam lokasi program)
    
    'Pengaturan adodc pada gui hanya akan merekam alamat lengkap lokasi database, tidak berdasarkan lokasi program (mendeteksi keberadaan program secara otomatis dimanapun berada, program ini)
    'intinya jika kita memiliki program pada alamat c:/program/rental_mobil dan pada suatu saat kita akan memindah ke drive d: atau dalam forder lainya atau komputer lain.
    'kemungkinan banyak adodc atau loadPicture yang akan kehilangan file mereka. karena mempunyai alamat yang baru
    
    'app.data = untuk mendeteksi lokasi program
End Function
