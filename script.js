// =========================================================================
// PENTING: GANTI DENGAN API KEY DAN SPREADSHEET ID ANDA YANG SESUNGGUHNYA
// =========================================================================
const API_KEY = 'AIzaSyDMWF1iMJoPPhjQPfFe-DFuiUZ8B1VrhbU'; // API Key Anda
const SPREADSHEET_ID = '1v-IzKh1FHfxmAs0yCRbaN2utzQpfL8g9bZLTBGKBcw0'; // Spreadsheet ID Anda
// =========================================================================

// Data Karyawan Branch Lampung Selatan Kalianda
// Jam masuk dan pulang disesuaikan per jabatan
const karyawanData = [
    { nama: "ANANDA ILHAM HAKIKI", jabatan: "AREA COORDINATOR", jam_masuk: "08:30:00", jam_pulang: "17:30:00" },
    { nama: "ANGGA PRATAMA", jabatan: "SECURITY", jam_masuk: "07:00:00", jam_pulang: "19:00:00" }, // Security shift
    { nama: "ARIL FACHRI NURDIANSAH", jabatan: "NOB", jam_masuk: "08:00:00", jam_pulang: "17:00:00" },
    { nama: "BIMA PRATAMA SAPUTRO", jabatan: "BRANCH LEADER", jam_masuk: "08:30:00", jam_pulang: "17:30:00" },
    { nama: "BIMA RAFLI RAMADANI", jabatan: "NOA", jam_masuk: "08:00:00", jam_pulang: "17:00:00" },
    { nama: "DEFLI YANTI", jabatan: "ADMIN STORE", jam_masuk: "08:30:00", jam_pulang: "17:30:00" },
    { nama: "DEKI KURNIAWAN", jabatan: "SECURITY", jam_masuk: "07:00:00", jam_pulang: "19:00:00" }, // Security shift
    { nama: "FEBRI SAPUTRA", jabatan: "NOA", jam_masuk: "08:00:00", jam_pulang: "17:00:00" },
    { nama: "HABIBI", jabatan: "AREA COORDINATOR", jam_masuk: "08:30:00", jam_pulang: "17:30:00" },
    { nama: "MUHYIDIN", jabatan: "OB", jam_masuk: "07:00:00", jam_pulang: "17:00:00" }, // OB jam 7 pagi
    { nama: "REKA ANDRIANSYAH", jabatan: "SECURITY", jam_masuk: "07:00:00", jam_pulang: "19:00:00" }, // Security shift
    { nama: "RISKI APRIANTO", jabatan: "AREA COORDINATOR", jam_masuk: "08:30:00", jam_pulang: "17:30:00" },
    { nama: "TOPAN AJI PRATAMA", jabatan: "NOA", jam_masuk: "08:00:00", jam_pulang: "17:00:00" },
    { nama: "YANDI YAHYA", jabatan: "NOB", jam_masuk: "08:00:00", jam_pulang: "17:00:00" },
    { nama: "KEMAL AL FARUQ", jabatan: "NOA", jam_masuk: "08:00:00", jam_pulang: "17:00:00" }
];

// Data Plat Nomor Kendaraan
const kendaraanData = [
    { jenis: "Mobil", plat: "BE 1820 AAR" },
    { jenis: "Mobil", plat: "BE 1226 AAY" },
    { jenis: "Motor", plat: "BE 2668 ADO" }
];

// Variabel global untuk menyimpan status yang akan dicatat setelah catatan diisi
let pendingAbsensiStatus = '';

// --- Inisialisasi Google API Client ---
function initClient() {
    gapi.client.init({
        apiKey: API_KEY,
        discoveryDocs: ["https://sheets.googleapis.com/$discovery/rest?version=v4"],
    }).then(function () {
        console.log("Google API client loaded.");
        checkLoginStatus(); // Cek status login saat API client dimuat
    }, function(error) {
        console.error("Error loading GAPI client:", error);
        document.getElementById('loginStatus').innerText = 'Gagal memuat API Google. Pastikan API key dan koneksi internet Anda benar.';
    });
}

// Memuat script GAPI saat DOM siap
document.addEventListener('DOMContentLoaded', () => {
    const script = document.createElement('script');
    script.src = 'https://apis.google.com/js/api.js';
    script.onload = () => gapi.load('client', initClient);
    document.head.appendChild(script);

    // Populate dropdowns saat DOM siap (meskipun akan tersembunyi jika belum login)
    populateKaryawanDropdown(); // Untuk form absensi
    populateKendaraanDropdown('berangkat');
    populateKendaraanDropdown('pulang');
    populatePengemudiDropdown(); // Panggil fungsi baru untuk dropdown pengemudi

    // Populate dropdowns untuk rekap absensi bulanan
    populateRekapAbsensiDropdowns();
    populateRekapKaryawanDropdown(); // Untuk dropdown pegawai di rekap bulanan

    // Populate dropdowns untuk manajemen jadwal
    populateManajemenJadwalDropdowns();
    populateManajemenJadwalKaryawanDropdown();

    // Event listeners untuk form login dan tombol logout
    document.getElementById('loginForm').addEventListener('submit', handleLogin);
    document.getElementById('logoutButton').addEventListener('click', handleLogout);
    document.getElementById('saveScheduleBtn').addEventListener('click', saveEmployeeSchedules);

    // Event Listener untuk Dropdown Karyawan (untuk menampilkan tombol aksi Clock In/Out)
    document.getElementById('namaKaryawanSelect').addEventListener('change', function() {
        const selectedName = this.value;
        const karyawanAksiDiv = document.getElementById('karyawanAksi');
        const selectedKaryawanNameSpan = document.getElementById('selectedKaryawanName');
        const selectedKaryawanJabatanSpan = document.getElementById('selectedKaryawanJabatan');
        const catatanInputDiv = document.getElementById('catatanInputDiv');

        if (selectedName) {
            const selectedKaryawan = karyawanData.find(k => k.nama === selectedName);
            selectedKaryawanNameSpan.innerText = selectedName;
            selectedKaryawanJabatanSpan.innerText = selectedKaryawan ? selectedKaryawan.jabatan : '-';
            karyawanAksiDiv.style.display = 'block';
            catatanInputDiv.style.display = 'none'; // Sembunyikan catatan input saat karyawan berubah
            document.getElementById('absensiCatatan').value = ''; // Bersihkan catatan
        } else {
            karyawanAksiDiv.style.display = 'none';
            selectedKaryawanNameSpan.innerText = '';
            selectedKaryawanJabatanSpan.innerText = '';
            catatanInputDiv.style.display = 'none';
            document.getElementById('absensiCatatan').value = '';
        }
        document.getElementById('absensiStatus').innerText = ''; // Bersihkan status sebelumnya
    });

    // Event listeners untuk tombol submit dan batal catatan
    document.getElementById('submitCatatanBtn').addEventListener('click', () => {
        const catatan = document.getElementById('absensiCatatan').value;
        markAbsensiStatus(pendingAbsensiStatus, catatan);
    });

    document.getElementById('cancelCatatanBtn').addEventListener('click', () => {
        document.getElementById('catatanInputDiv').style.display = 'none';
        document.getElementById('absensiCatatan').value = '';
        pendingAbsensiStatus = '';
    });
});

// --- Fungsi Login/Logout ---
async function handleLogin(e) {
    e.preventDefault();
    const username = document.getElementById('loginUsername').value;
    const password = document.getElementById('loginPassword').value;
    const loginStatusDiv = document.getElementById('loginStatus');
    loginStatusDiv.innerText = 'Memverifikasi...';

    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Users!A:B', // Ambil kolom Username dan Password dari tab Users
        });

        const users = response.result.values;
        let isAuthenticated = false;

        if (users && users.length > 0) {
            // Lewati baris header (indeks 0) jika ada
            for (let i = 1; i < users.length; i++) {
                const userRow = users[i];
                const storedUsername = userRow[0];
                const storedPassword = userRow[1];

                if (username === storedUsername && password === storedPassword) {
                    isAuthenticated = true;
                    // Simpan status login di sessionStorage agar hilang saat browser ditutup
                    sessionStorage.setItem('loggedIn', 'true');
                    sessionStorage.setItem('username', username);
                    break;
                }
            }
        }

        if (isAuthenticated) {
            loginStatusDiv.innerText = 'Login berhasil!';
            showAppContent();
        } else {
            loginStatusDiv.innerText = 'Username atau password salah.';
        }

    } catch (error) {
        console.error('Error saat login:', error);
        loginStatusDiv.innerText = 'Terjadi kesalahan saat memverifikasi login. Cek konsol browser.';
    }
}

function handleLogout() {
    sessionStorage.removeItem('loggedIn');
    sessionStorage.removeItem('username');
    hideAppContent();
    document.getElementById('loginUsername').value = '';
    document.getElementById('loginPassword').value = '';
    document.getElementById('loginStatus').innerText = '';
}

function showAppContent() {
    document.getElementById('login-container').style.display = 'none';
    document.getElementById('app-container').style.display = 'block';
    document.body.classList.remove('no-scroll'); // Allow scrolling on app content
    showSection('dashboard-summary'); // Tampilkan dashboard utama sebagai halaman pertama setelah login
}

function hideAppContent() {
    document.getElementById('login-container').style.display = 'flex'; // Gunakan flex agar posisi card login tetap center
    document.getElementById('app-container').style.display = 'none';
    document.body.classList.add('no-scroll'); // Prevent scrolling on login screen
}

function checkLoginStatus() {
    if (sessionStorage.getItem('loggedIn') === 'true') {
        showAppContent();
    } else {
        hideAppContent();
    }
}

// --- Fungsi-fungsi Utama Aplikasi ---

// Mengatur tampilan section yang aktif
function showSection(sectionId) {
    document.querySelectorAll('section').forEach(section => {
        section.style.display = 'none';
        section.classList.remove('active');
    });
    document.getElementById(sectionId).style.display = 'block';
    document.getElementById(sectionId).classList.add('active');

    // Gulirkan ke bagian atas section yang aktif
    document.getElementById(sectionId).scrollIntoView({ behavior: 'smooth', block: 'start' });


    // Sembunyikan tombol simpan jadwal secara default
    document.getElementById('saveScheduleBtn').style.display = 'none';
    // Sembunyikan input catatan absensi
    document.getElementById('catatanInputDiv').style.display = 'none';


    // Muat data yang relevan saat section dibuka
    if (sectionId === 'dashboard-summary') loadDailyDashboardSummary();
    else if (sectionId === 'rekapAbsensi') generateAbsensiRecap(); // Memanggil rekap bulanan
    else if (sectionId === 'manajemenJadwal') {
        // Tidak langsung memuat jadwal, karena perlu input bulan/tahun/karyawan
        document.getElementById('manajemenJadwalTableContainer').innerHTML = '<p class="text-muted text-center">Pilih bulan, tahun, dan pegawai lalu klik \'Muat Jadwal\'.</p>';
        document.getElementById('jadwalStatus').innerText = '';
    }
    else if (sectionId === 'absensi') loadAbsensiHistory();
    else if (sectionId === 'kendaraan') {
        loadKendaraanHistory();
        populateKendaraanDropdown('pulang'); // Pastikan dropdown kepulangan terupdate
    }
    else if (sectionId === 'tamu') loadTamuHistory();
}

// Mendapatkan tanggal dan waktu saat ini (untuk format DD/MM/YYYY HH:MM:SS)
function getCurrentDateTime() {
    const now = new Date();
    const date = now.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const time = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    return { date, time, rawDate: now }; // Tambahkan rawDate untuk perbandingan lebih mudah
}

// Mendapatkan objek Date dari string tanggal (DD/MM/YYYY) dan waktu (HH:MM:SS)
function parseDateTime(dateStr, timeStr) {
    // Pastikan format dateStr adalah DD/MM/YYYY
    const [day, month, year] = dateStr.split('/').map(Number);
    const [hours, minutes, seconds] = timeStr.split(':').map(Number);
    // Bulan di JavaScript adalah 0-indexed
    return new Date(year, month - 1, day, hours, minutes, seconds);
}

// --- Dashboard Summary (Ringkasan Harian) ---
async function loadDailyDashboardSummary() {
    document.getElementById('currentDateDisplay').innerText = getCurrentDateTime().date;

    // Reset counts and details
    document.getElementById('totalTepatWaktu').innerText = '0';
    document.getElementById('totalTerlambat').innerText = '0';
    document.getElementById('totalCutiIzinSakit').innerText = '0';
    document.getElementById('totalTidakHadir').innerText = '0';
    document.getElementById('absensiDailyDetail').innerHTML = '<p class="text-muted">Memuat data absensi...</p>';
    document.getElementById('kendaraanKeluarCount').innerText = '0';
    document.getElementById('kendaraanDailyDetail').innerHTML = '<p class="text-muted">Memuat data kendaraan...</p>';
    document.getElementById('tamuTodayCount').innerText = '0';
    document.getElementById('tamuDailyDetail').innerHTML = '<p class="text-muted">Memuat data tamu...</p>';


    const todayDate = getCurrentDateTime().date;
    const allKaryawanNames = karyawanData.map(k => k.nama);

    let absensiHarian = {}; // {nama_karyawan: {status: 'V/TL/C/I/S/A/Libur', waktu_masuk: 'HH:MM:SS', waktu_keluar: 'HH:MM:SS'}}
    let totalTepatWaktu = 0;
    let totalTerlambat = 0;
    let totalCutiIzinSakit = 0;
    let totalTidakHadir = 0;
    let totalLibur = 0; // New counter for Libur

    // --- Absensi Harian ---
    try {
        const responseAbsensi = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Absensi!A:F', // Tanggal, Waktu Masuk, Waktu Keluar, Nama Karyawan, Status, Catatan
        });
        const absensiRows = responseAbsensi.result.values || [];
        const dailyAbsensiEntries = absensiRows.slice(1).filter(row => row[0] === todayDate);

        // Inisialisasi semua karyawan sebagai "Tidak Hadir" (A)
        allKaryawanNames.forEach(name => {
            absensiHarian[name] = { status: 'A', waktu_masuk: '', waktu_keluar: '', catatan: '' };
        });

        dailyAbsensiEntries.forEach(entry => {
            const nama = entry[3];
            const waktuMasukStr = entry[1];
            const waktuKeluarStr = entry[2];
            const statusSheet = entry[4]; // Kolom Status
            const catatanSheet = entry[5]; // Kolom Catatan

            // Temukan data karyawan untuk mendapatkan jam standar mereka
            const karyawan = karyawanData.find(k => k.nama === nama);
            if (!karyawan) return; // Lewati jika karyawan tidak ditemukan (seharusnya tidak terjadi)

            // Prioritas: Status dari kolom Status (Libur, Cuti, Izin, Sakit)
            if (statusSheet) {
                const lowerStatus = statusSheet.toLowerCase();
                if (lowerStatus.includes('libur')) {
                    absensiHarian[nama].status = 'Libur';
                } else if (lowerStatus.includes('cuti')) {
                    absensiHarian[nama].status = 'C';
                } else if (lowerStatus.includes('izin')) {
                    absensiHarian[nama].status = 'I';
                } else if (lowerStatus.includes('sakit')) {
                    absensiHarian[nama].status = 'S';
                }
                absensiHarian[nama].catatan = catatanSheet || '';
            }
            // Jika tidak ada status khusus, cek waktu masuk/keluar
            if (absensiHarian[nama].status === 'A') { // Hanya jika belum di-set oleh status khusus
                if (waktuMasukStr) {
                    absensiHarian[nama].waktu_masuk = waktuMasukStr;
                    absensiHarian[nama].waktu_keluar = waktuKeluarStr || '';

                    const waktuMasuk = parseDateTime(todayDate, waktuMasukStr);
                    const jamMasukStandar = parseDateTime(todayDate, karyawan.jam_masuk);
                    
                    if (waktuMasuk.getTime() > jamMasukStandar.getTime() + (5 * 60 * 1000)) {
                        absensiHarian[nama].status = 'TL';
                    } else {
                        absensiHarian[nama].status = 'V';
                    }

                    // Cek PSW jika ada waktu keluar
                    if (waktuKeluarStr) {
                        const waktuKeluar = parseDateTime(todayDate, waktuKeluarStr);
                        const jamPulangStandar = parseDateTime(todayDate, karyawan.jam_pulang);
                        if (waktuKeluar.getTime() < jamPulangStandar.getTime() - (5 * 60 * 1000)) {
                            absensiHarian[nama].status = (absensiHarian[nama].status === 'TL' ? 'TL+PSW' : 'PSW');
                        }
                    }
                } else {
                    absensiHarian[nama].status = 'A'; // Tidak ada catatan dan tidak ada waktu masuk
                }
            }
        });

        // Hitung total dan buat tabel detail absensi
        let absensiDetailHTML = '<table class="table table-striped table-hover"><thead><tr><th>Nama</th><th>Masuk</th><th>Keluar</th><th>Status</th></tr></thead><tbody>';
        karyawanData.forEach(karyawan => {
            const statusData = absensiHarian[karyawan.nama];
            let displayStatus = statusData.status;

            switch(statusData.status) {
                case 'V':
                    totalTepatWaktu++;
                    break;
                case 'TL':
                case 'TL+PSW':
                    totalTerlambat++;
                    break;
                case 'PSW':
                    totalTidakHadir++; // Dianggap tidak hadir penuh jika hanya PSW
                    break;
                case 'C':
                case 'I':
                case 'S':
                    totalCutiIzinSakit++;
                    // displayStatus += (statusData.catatan ? ` (${statusData.catatan})` : '');
                    break;
                case 'Libur':
                    totalLibur++; // Hitung libur secara terpisah
                    break;
                case 'A':
                    totalTidakHadir++;
                    break;
            }

            absensiDetailHTML += `<tr>
                <td>${karyawan.nama}</td>
                <td>${statusData.waktu_masuk}</td>
                <td>${statusData.waktu_keluar}</td>
                <td><span class="badge ${getStatusBadgeClass(statusData.status)}">${displayStatus}</span></td>
            </tr>`;
        });
        absensiDetailHTML += '</tbody></table>';

        document.getElementById('totalTepatWaktu').innerText = totalTepatWaktu;
        document.getElementById('totalTerlambat').innerText = totalTerlambat;
        document.getElementById('totalCutiIzinSakit').innerText = totalCutiIzinSakit + totalLibur; // Tambahkan Libur ke Cuti/Izin/Sakit
        document.getElementById('totalTidakHadir').innerText = totalTidakHadir;
        document.getElementById('absensiDailyDetail').innerHTML = absensiDetailHTML;

    } catch (error) {
        console.error('Error loading daily absensi summary:', error);
        document.getElementById('absensiDailyDetail').innerHTML = '<p class="text-danger">Gagal memuat ringkasan absensi harian.</p>';
    }

    // --- Log Kendaraan Aktif Hari Ini ---
    try {
        const responseKendaraan = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Kendaraan!A:J', // Tanggal, Waktu Berangkat, Waktu Pulang, Plat, Jenis, Pengemudi, Tujuan, KM Awal, KM Akhir, Catatan
        });
        const kendaraanRows = responseKendaraan.result.values || [];
        const activeKendaraanEntries = kendaraanRows.slice(1).filter(row =>
            row[0] === todayDate && (row[2] === '' || row[2] === undefined) // Tanggal hari ini & Waktu Pulang kosong
        );

        let kendaraanDetailHTML = '<table class="table table-striped table-hover"><thead><tr><th>Plat</th><th>Pengemudi</th><th>Tujuan</th><th>Waktu Berangkat</th></tr></thead><tbody>';
        if (activeKendaraanEntries.length > 0) {
            activeKendaraanEntries.forEach(entry => {
                kendaraanDetailHTML += `<tr>
                    <td>${entry[3] || ''}</td>
                    <td>${entry[5] || ''}</td>
                    <td>${entry[6] || ''}</td>
                    <td>${entry[1] || ''}</td>
                </tr>`;
            });
            document.getElementById('kendaraanKeluarCount').innerText = activeKendaraanEntries.length;
        } else {
            kendaraanDetailHTML += '<tr><td colspan="4" class="text-center text-muted">Tidak ada kendaraan sedang keluar.</td></tr>';
            document.getElementById('kendaraanKeluarCount').innerText = '0';
        }
        kendaraanDetailHTML += '</tbody></table>';
        document.getElementById('kendaraanDailyDetail').innerHTML = kendaraanDetailHTML;

    } catch (error) {
        console.error('Error loading daily kendaraan summary:', error);
        document.getElementById('kendaraanDailyDetail').innerHTML = '<p class="text-danger">Gagal memuat ringkasan kendaraan harian.</p>';
    }

    // --- Tamu Terdaftar Hari Ini ---
    try {
        const responseTamu = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Tamu!A:I', // Tanggal Masuk, Waktu Masuk, Tanggal Keluar, Waktu Keluar, Nama Tamu, Nomor Telepon, Tujuan Kunjungan, PIC, Catatan
        });
        const tamuRows = responseTamu.result.values || [];
        const dailyTamuEntries = tamuRows.slice(1).filter(row =>
            row[0] === todayDate && (row[2] === '' || row[2] === undefined) // Tanggal Masuk hari ini & Tanggal Keluar kosong
        );

        let tamuDetailHTML = '<table class="table table-striped table-hover"><thead><tr><th>Nama Tamu</th><th>Tujuan</th><th>PIC</th><th>Waktu Masuk</th></tr></thead><tbody>';
        if (dailyTamuEntries.length > 0) {
            dailyTamuEntries.forEach(entry => {
                tamuDetailHTML += `<tr>
                    <td>${entry[4] || ''}</td>
                    <td>${entry[6] || ''}</td>
                    <td>${entry[7] || ''}</td>
                    <td>${entry[1] || ''}</td>
                </tr>`;
            });
            document.getElementById('tamuTodayCount').innerText = dailyTamuEntries.length;
        } else {
            tamuDetailHTML += '<tr><td colspan="4" class="text-center text-muted">Tidak ada tamu terdaftar hari ini.</td></tr>';
            document.getElementById('tamuTodayCount').innerText = '0';
        }
        tamuDetailHTML += '</tbody></table>';
        document.getElementById('tamuDailyDetail').innerHTML = tamuDetailHTML;

    } catch (error) {
        console.error('Error loading daily tamu summary:', error);
        document.getElementById('tamuDailyDetail').innerHTML = '<p class="text-danger">Gagal memuat ringkasan tamu harian.</p>';
    }
}

// Helper untuk mendapatkan kelas badge CSS berdasarkan status
function getStatusBadgeClass(status) {
    switch (status) {
        case 'V': return 'bg-success';
        case 'TL': return 'bg-warning text-dark';
        case 'PSW': return 'bg-danger';
        case 'C': return 'bg-info';
        case 'I': return 'bg-primary'; // Changed to primary for Izin
        case 'S': return 'bg-purple'; // Custom class for purple
        case 'A': return 'bg-dark';
        case 'Libur': return 'bg-light text-dark'; // New class for Libur
        case 'TL+PSW': return 'bg-warning text-dark'; // Kombinasi
        default: return 'bg-light text-muted';
    }
}


// --- Rekap Absensi Bulanan Functions ---
function populateRekapAbsensiDropdowns() {
    const rekapBulanSelect = document.getElementById('rekapBulan');
    const rekapTahunSelect = document.getElementById('rekapTahun');

    const months = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ];
    const currentMonth = new Date().getMonth(); // 0-indexed
    const currentYear = new Date().getFullYear();

    // Populate Bulan
    rekapBulanSelect.innerHTML = '';
    months.forEach((month, index) => {
        const option = document.createElement('option');
        option.value = index + 1; // 1-indexed month value
        option.textContent = month;
        if (index === currentMonth) {
            option.selected = true;
        }
        rekapBulanSelect.appendChild(option);
    });

    // Populate Tahun (misalnya 5 tahun ke belakang dan 1 tahun ke depan)
    rekapTahunSelect.innerHTML = '';
    for (let year = currentYear - 2; year <= currentYear + 1; year++) {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        if (year === currentYear) {
            option.selected = true;
        }
        rekapTahunSelect.appendChild(option);
    }
}

function populateRekapKaryawanDropdown() {
    const rekapKaryawanSelect = document.getElementById('rekapKaryawan');
    karyawanData.forEach(karyawan => {
        const option = document.createElement('option');
        option.value = karyawan.nama;
        option.textContent = karyawan.nama;
        rekapKaryawanSelect.appendChild(option);
    });
}

async function generateAbsensiRecap() {
    const selectedMonth = parseInt(document.getElementById('rekapBulan').value);
    const selectedYear = parseInt(document.getElementById('rekapTahun').value);
    const selectedKaryawan = document.getElementById('rekapKaryawan').value;
    const tableContainer = document.getElementById('rekapAbsensiTableContainer');
    tableContainer.innerHTML = '<p class="text-muted text-center">Memuat rekap absensi...</p>';

    try {
        const responseAbsensi = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Absensi!A:F', // Tanggal, Waktu Masuk, Waktu Keluar, Nama Karyawan, Status, Catatan
        });
        const allAbsensiData = responseAbsensi.result.values || [];
        // Hapus header
        const absensiEntries = allAbsensiData.slice(1);

        // Ambil data jadwal dari Google Sheet 'Jadwal'
        const responseJadwal = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Jadwal!A:C', // Nama Karyawan, Tanggal, Status
        });
        const allJadwalData = responseJadwal.result.values || [];
        const monthlySchedules = allJadwalData.slice(1).filter(entry => {
            const [empName, dateStr] = entry;
            const [day, month, year] = dateStr.split('/').map(Number);
            return month === selectedMonth && year === selectedYear;
        });


        const daysInMonth = new Date(selectedYear, selectedMonth, 0).getDate();
        const datesOfMonth = Array.from({ length: daysInMonth }, (_, i) => i + 1);

        let filteredKaryawan = karyawanData;
        if (selectedKaryawan !== 'Semua') {
            filteredKaryawan = karyawanData.filter(k => k.nama === selectedKaryawan);
        }

        if (filteredKaryawan.length === 0) {
            tableContainer.innerHTML = '<p class="text-muted text-center">Tidak ada data karyawan yang dipilih.</p>';
            return;
        }

        let tableHTML = '<table class="table table-bordered table-sm" id="rekapAbsensiTable"><thead><tr>';
        tableHTML += '<th>No</th><th>Pegawai</th>';
        datesOfMonth.forEach(day => {
            tableHTML += `<th class="day-header">${day}</th>`;
        });
        tableHTML += '</tr></thead><tbody>';

        filteredKaryawan.forEach((karyawan, index) => {
            tableHTML += `<tr><td>${index + 1}</td><td>${karyawan.nama}</td>`;
            datesOfMonth.forEach(day => {
                const targetDate = `${String(day).padStart(2, '0')}/${String(selectedMonth).padStart(2, '0')}/${selectedYear}`;
                const dailyEntries = absensiEntries.filter(entry =>
                    entry[3] === karyawan.nama && entry[0] === targetDate
                );

                let status = '';
                let statusClass = 'status-empty'; // Default: tidak ada data

                // Cek jadwal dari sheet 'Jadwal' terlebih dahulu
                const scheduledEntry = monthlySchedules.find(entry =>
                    entry[0] === karyawan.nama && entry[1] === targetDate
                );

                if (scheduledEntry) {
                    const scheduledStatus = scheduledEntry[2];
                    status = scheduledStatus;
                    switch(scheduledStatus) {
                        case 'Libur': statusClass = 'status-off'; break; // Menggunakan status-off untuk Libur
                        case 'Tukar Libur': statusClass = 'bg-info text-white'; break;
                        case 'Cuti': statusClass = 'status-C'; break;
                        case 'Izin': statusClass = 'status-I'; break;
                        case 'Sakit': statusClass = 'status-S'; break;
                        case 'Kerja': statusClass = 'status-V'; break; // Default untuk 'Kerja' di jadwal
                        default: statusClass = 'status-empty'; break;
                    }
                }

                // Sekarang, proses absensi aktual jika status masih kosong atau 'Kerja' dari jadwal
                if (dailyEntries.length > 0) {
                    const latestEntry = dailyEntries[dailyEntries.length - 1]; // Ambil entri terbaru
                    const waktuMasukStr = latestEntry[1];
                    const waktuKeluarStr = latestEntry[2];
                    const statusSheet = latestEntry[4]; // Kolom Status dari sheet Absensi
                    // const catatanSheet = latestEntry[5]; // Kolom Catatan dari sheet Absensi

                    // Temukan data karyawan untuk mendapatkan jam standar mereka
                    const currentKaryawan = karyawanData.find(k => k.nama === karyawan.nama);
                    const jamMasukStandar = parseDateTime(targetDate, currentKaryawan.jam_masuk);
                    const jamPulangStandar = parseDateTime(targetDate, currentKaryawan.jam_pulang);

                    // Prioritas: Status dari sheet Absensi (Cuti, Izin, Sakit, Libur dari input manual)
                    if (statusSheet) {
                        const lowerStatus = statusSheet.toLowerCase();
                        if (lowerStatus.includes('libur')) { status = 'Libur'; statusClass = 'status-off'; }
                        else if (lowerStatus.includes('cuti')) { status = 'C'; statusClass = 'status-C'; }
                        else if (lowerStatus.includes('izin')) { status = 'I'; statusClass = 'status-I'; }
                        else if (lowerStatus.includes('sakit')) { status = 'S'; statusClass = 'status-S'; }
                        // Jika status dari sheet adalah salah satu dari ini, itu akan menimpa status jadwal
                        if (['Libur', 'C', 'I', 'S'].includes(status)) {
                            // Status sudah diatur, tidak perlu cek waktu
                        } else if (waktuMasukStr) { // Jika ada clock-in
                            const waktuMasuk = parseDateTime(targetDate, waktuMasukStr);
                            if (waktuMasuk.getTime() > jamMasukStandar.getTime() + (5 * 60 * 1000)) {
                                status = 'TL'; statusClass = 'status-TL';
                            } else {
                                status = 'V'; statusClass = 'status-V';
                            }
                            if (waktuKeluarStr && parseDateTime(targetDate, waktuKeluarStr).getTime() < jamPulangStandar.getTime() - (5 * 60 * 1000)) {
                                status = (status === 'TL' ? 'TL+PSW' : 'PSW');
                                statusClass = (status === 'TL+PSW' ? 'status-TL' : 'status-PSW');
                            }
                        } else { // Tidak ada clock-in, tidak ada status khusus di sheet absensi
                            status = 'A'; statusClass = 'status-A';
                        }
                    } else if (status === '' || status === 'Kerja') { // Jika tidak ada status khusus dari sheet absensi dan belum diatur oleh jadwal
                        if (waktuMasukStr) {
                            const waktuMasuk = parseDateTime(targetDate, waktuMasukStr);
                            if (waktuMasuk.getTime() > jamMasukStandar.getTime() + (5 * 60 * 1000)) {
                                status = 'TL'; statusClass = 'status-TL';
                            } else {
                                status = 'V'; statusClass = 'status-V';
                            }
                            if (waktuKeluarStr && parseDateTime(targetDate, waktuKeluarStr).getTime() < jamPulangStandar.getTime() - (5 * 60 * 1000)) {
                                status = (status === 'TL' ? 'TL+PSW' : 'PSW');
                                statusClass = (status === 'TL+PSW' ? 'status-TL' : 'status-PSW');
                            }
                        } else {
                            status = 'A'; statusClass = 'status-A';
                        }
                    }
                } else if (status === '' || status === 'Kerja') { // Tidak ada entri harian, dan tidak dijadwalkan sebagai Libur/Cuti/Izin/Sakit
                    status = 'A'; statusClass = 'status-A';
                }
                tableHTML += `<td class="${statusClass}">${status}</td>`;
            });
            tableHTML += '</tr>';
        });

        tableHTML += '</tbody></table>';
        tableContainer.innerHTML = tableHTML;

    } catch (error) {
        console.error('Error generating absensi recap:', error);
        tableContainer.innerHTML = '<p class="text-danger text-center">Gagal memuat rekap absensi. Cek konsol browser untuk detail error.</p>';
    }
}


// --- Absensi Karyawan ---
async function clockIn() {
    const nama = document.getElementById('namaKaryawanSelect').value;
    if (!nama) {
        alert('Nama karyawan harus dipilih!');
        return;
    }
    const { date, time } = getCurrentDateTime();
    // Struktur data sesuai header: Tanggal, Waktu Masuk, Waktu Keluar, Nama Karyawan, Status, Catatan
    const values = [[date, time, '', nama, 'Hadir', '']];

    try {
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Absensi!A:F', // Sesuaikan dengan jumlah kolom header
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: values },
        });
        document.getElementById('absensiStatus').className = 'mt-3 alert alert-success';
        document.getElementById('absensiStatus').innerText = `Clock In berhasil untuk ${nama} pada ${time}!`;
        // Reset form untuk input selanjutnya
        document.getElementById('namaKaryawanSelect').value = '';
        document.getElementById('karyawanAksi').style.display = 'none';
        document.getElementById('selectedKaryawanName').innerText = '';
        document.getElementById('selectedKaryawanJabatan').innerText = '';
        loadAbsensiHistory(); // Muat ulang riwayat
        loadDailyDashboardSummary(); // Perbarui dashboard harian
    } catch (error) {
        console.error('Error clocking in:', error);
        document.getElementById('absensiStatus').className = 'mt-3 alert alert-danger';
        document.getElementById('absensiStatus').innerText = 'Gagal Clock In. Cek konsol browser untuk detail error.';
    }
}

async function clockOut() {
    const nama = document.getElementById('namaKaryawanSelect').value;
    if (!nama) {
        alert('Nama karyawan harus dipilih!');
        return;
    }
    const { date: todayDate, time } = getCurrentDateTime();

    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Absensi!A:F',
        });
        const rows = response.result.values;
        if (rows) {
            let rowIndexToUpdate = -1;
            // Cari baris terakhir karyawan pada hari ini yang belum Clock Out
            for (let i = rows.length - 1; i >= 0; i--) { // Loop mundur dari baris terakhir
                // Kolom Tanggal (indeks 0), Kolom Nama Karyawan (indeks 3), Kolom Waktu Keluar (indeks 2)
                if (rows[i][0] === todayDate && rows[i][3] === nama && (rows[i][2] === '' || rows[i][2] === undefined)) {
                    rowIndexToUpdate = i;
                    break;
                }
            }

            if (rowIndexToUpdate !== -1) {
                // Update kolom Waktu Keluar (kolom C)
                await gapi.client.sheets.spreadsheets.values.update({
                    spreadsheetId: SPREADSHEET_ID,
                    range: `Absensi!C${rowIndexToUpdate + 1}`, // +1 karena sheet 1-indexed
                    valueInputOption: 'USER_ENTERED',
                    resource: { values: [[time]] },
                });
                document.getElementById('absensiStatus').className = 'mt-3 alert alert-success';
                document.getElementById('absensiStatus').innerText = `Clock Out berhasil untuk ${nama} pada ${time}!`;
                // Reset form untuk input selanjutnya
                document.getElementById('namaKaryawanSelect').value = '';
                document.getElementById('karyawanAksi').style.display = 'none';
                document.getElementById('selectedKaryawanName').innerText = '';
                document.getElementById('selectedKaryawanJabatan').innerText = '';
                loadAbsensiHistory(); // Muat ulang riwayat
                loadDailyDashboardSummary(); // Perbarui dashboard harian
            } else {
                document.getElementById('absensiStatus').className = 'mt-3 alert alert-warning';
                document.getElementById('absensiStatus').innerText = `Tidak ditemukan Clock In yang aktif untuk ${nama} pada hari ini.`;
            }
        } else {
            document.getElementById('absensiStatus').className = 'mt-3 alert alert-warning';
            document.getElementById('absensiStatus').innerText = 'Tidak ada data absensi untuk diproses.';
        }
    } catch (error) {
        console.error('Error clocking out:', error);
        document.getElementById('absensiStatus').className = 'mt-3 alert alert-danger';
        document.getElementById('absensiStatus').innerText = 'Gagal Clock Out. Cek konsol browser untuk detail error.';
    }
}

// Fungsi baru untuk menampilkan input catatan dan mengatur status yang akan dicatat
function promptForCatatan(status) {
    const nama = document.getElementById('namaKaryawanSelect').value;
    if (!nama) {
        alert('Nama karyawan harus dipilih sebelum mencatat status!');
        return;
    }
    pendingAbsensiStatus = status; // Simpan status yang akan dicatat
    document.getElementById('catatanInputDiv').style.display = 'block';
    document.getElementById('absensiCatatan').focus();
}

// Fungsi untuk mencatat status absensi (Libur, Cuti, Izin, Sakit)
async function markAbsensiStatus(statusType, catatan = '') {
    const nama = document.getElementById('namaKaryawanSelect').value;
    if (!nama) {
        alert('Nama karyawan harus dipilih!');
        return;
    }
    const { date } = getCurrentDateTime();

    // Untuk status non-hadir/libur, waktu masuk dan keluar dikosongkan
    // Struktur data: Tanggal, Waktu Masuk, Waktu Keluar, Nama Karyawan, Status, Catatan
    const values = [[date, '', '', nama, statusType, catatan]];

    try {
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Absensi!A:F',
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: values },
        });
        document.getElementById('absensiStatus').className = 'mt-3 alert alert-success';
        document.getElementById('absensiStatus').innerText = `Status ${statusType} berhasil dicatat untuk ${nama}!`;
        
        // Reset UI
        document.getElementById('namaKaryawanSelect').value = '';
        document.getElementById('karyawanAksi').style.display = 'none';
        document.getElementById('selectedKaryawanName').innerText = '';
        document.getElementById('selectedKaryawanJabatan').innerText = '';
        document.getElementById('catatanInputDiv').style.display = 'none';
        document.getElementById('absensiCatatan').value = '';
        pendingAbsensiStatus = '';

        loadAbsensiHistory(); // Muat ulang riwayat
        loadDailyDashboardSummary(); // Perbarui dashboard harian
    } catch (error) {
        console.error(`Error marking ${statusType}:`, error);
        document.getElementById('absensiStatus').className = 'mt-3 alert alert-danger';
        document.getElementById('absensiStatus').innerText = `Gagal mencatat status ${statusType}. Cek konsol browser untuk detail error.`;
    }
}


async function loadAbsensiHistory() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Absensi!A:F',
        });
        const rows = response.result.values;
        const historyDiv = document.getElementById('absensiHistory');
        historyDiv.innerHTML = ''; // Kosongkan dulu

        if (rows && rows.length > 1) { // Periksa jika ada data selain header
            const table = document.createElement('table');
            table.className = 'table table-striped table-bordered';
            let tableHTML = '<thead><tr><th>Tanggal</th><th>Masuk</th><th>Keluar</th><th>Nama</th><th>Status</th><th>Catatan</th></tr></thead><tbody>';
            // Loop mundur untuk menampilkan data terbaru di atas
            for (let i = rows.length - 1; i >= 1; i--) {
                const row = rows[i];
                tableHTML += `<tr>
                                <td>${row[0] || ''}</td>
                                <td>${row[1] || ''}</td>
                                <td>${row[2] || ''}</td>
                                <td>${row[3] || ''}</td>
                                <td>${row[4] || ''}</td>
                                <td>${row[5] || ''}</td>
                              </tr>`;
            }
            tableHTML += '</tbody>';
            table.innerHTML = tableHTML;
            historyDiv.appendChild(table);
        } else {
            historyDiv.innerHTML = '<p class="text-muted">Belum ada riwayat absensi.</p>';
        }
    } catch (error) {
        console.error('Error loading absensi history:', error);
        document.getElementById('absensiHistory').innerHTML = '<p class="text-danger">Gagal memuat riwayat absensi.</p>';
    }
}


// --- Log Kendaraan ---
// Fungsi untuk mengisi dropdown plat nomor (digunakan untuk keberangkatan dan kepulangan)
function populateKendaraanDropdown(type) { // 'berangkat' atau 'pulang'
    const selectElementId = type === 'berangkat' ? 'platNomorSelectBerangkat' : 'platNomorSelectPulang';
    const selectElement = document.getElementById(selectElementId);

    // Kosongkan dulu opsi lama dan tambahkan opsi default
    selectElement.innerHTML = `<option value="">-- Pilih Plat Nomor --</option>`;

    kendaraanData.forEach(kendaraan => {
        const option = document.createElement('option');
        option.value = kendaraan.plat;
        option.textContent = `${kendaraan.plat} (${kendaraan.jenis})`;
        selectElement.appendChild(option);
    });

    if (type === 'berangkat') {
        selectElement.addEventListener('change', (event) => {
            const selectedPlat = event.target.value;
            const selectedKendaraan = kendaraanData.find(k => k.plat === selectedPlat);
            document.getElementById('jenisKendaraanBerangkat').value = selectedKendaraan ? selectedKendaraan.jenis : '';
        });
    } else if (type === 'pulang') {
        selectElement.addEventListener('change', async (event) => {
            const selectedPlat = event.target.value;
            const pengemudiPulangInfo = document.getElementById('pengemudiPulangInfo');
            const tujuanPulangInfo = document.getElementById('tujuanPulangInfo');
            const kmAwalPulangInfo = document.getElementById('kmAwalPulangInfo');
            const statusKepulanganDiv = document.getElementById('statusKepulangan');

            // Reset info
            pengemudiPulangInfo.innerText = '';
            tujuanPulangInfo.innerText = '';
            kmAwalPulangInfo.innerText = '';
            statusKepulanganDiv.innerText = '';
            document.getElementById('formKepulangan').dataset.rowIndex = ''; // Hapus data rowIndex

            if (selectedPlat) {
                try {
                    const response = await gapi.client.sheets.spreadsheets.values.get({
                        spreadsheetId: SPREADSHEET_ID,
                        range: 'Kendaraan!A:J', // Sesuaikan dengan rentang header baru (A-J)
                    });
                    const rows = response.result.values;
                    if (rows) {
                        let foundEntry = null;
                        // Loop mundur untuk mencari keberangkatan terbaru tanpa waktu pulang
                        for (let i = rows.length - 1; i >= 1; i--) { // Lewati header (indeks 0)
                            const row = rows[i];
                            // Cek Plat Nomor (indeks 3) dan Waktu Pulang (indeks 2) harus kosong
                            if (row[3] === selectedPlat && (row[2] === '' || row[2] === undefined)) {
                                foundEntry = row;
                                // Simpan rowIndex agar bisa diupdate nanti
                                document.getElementById('formKepulangan').dataset.rowIndex = i + 1; // +1 karena sheet 1-indexed
                                break;
                            }
                        }

                        if (foundEntry) {
                            pengemudiPulangInfo.innerText = foundEntry[5] || '-'; // Nama Pengemudi (indeks 5)
                            tujuanPulangInfo.innerText = foundEntry[6] || '-';   // Tujuan (indeks 6)
                            kmAwalPulangInfo.innerText = foundEntry[7] || '-';   // KM Awal (indeks 7)
                        } else {
                            statusKepulanganDiv.className = 'mt-3 alert alert-warning';
                            statusKepulanganDiv.innerText = 'Tidak ada log keberangkatan aktif untuk kendaraan ini.';
                        }
                    }
                } catch (error) {
                    console.error('Error mencari log keberangkatan:', error);
                    statusKepulanganDiv.className = 'mt-3 alert alert-danger';
                    statusKepulanganDiv.innerText = 'Gagal mencari log keberangkatan. Cek konsol browser.';
                }
            }
        });
    }
}

// Fungsi baru untuk mengisi dropdown nama pengemudi dari data karyawan
function populatePengemudiDropdown() {
    const selectElement = document.getElementById('namaPengemudiBerangkat');
    // Kosongkan dulu opsi lama dan tambahkan opsi default
    selectElement.innerHTML = `<option value="">-- Pilih Pengemudi --</option>`;
    karyawanData.forEach(karyawan => {
        const option = document.createElement('option');
        option.value = karyawan.nama;
        option.textContent = karyawan.nama;
        selectElement.appendChild(option);
    });
}


// Handler untuk form Keberangkatan Kendaraan
document.getElementById('formKeberangkatan').addEventListener('submit', async function(e) {
    e.preventDefault();
    const platNomor = document.getElementById('platNomorSelectBerangkat').value;
    const jenisKendaraan = document.getElementById('jenisKendaraanBerangkat').value;
    const namaPengemudi = document.getElementById('namaPengemudiBerangkat').value;
    const tujuan = document.getElementById('tujuanKendaraanBerangkat').value;
    const kmAwal = document.getElementById('kmAwal').value;
    const { date, time } = getCurrentDateTime();

    if (!platNomor || !namaPengemudi || !tujuan || !kmAwal) {
        alert('Mohon lengkapi semua data Keberangkatan!');
        return;
    }

    // Struktur data sesuai header Kendaraan (A:J):
    // Tanggal, Waktu Berangkat, Waktu Pulang (kosong), Nomor Polisi, Jenis, Nama Pengemudi, Tujuan, KM Awal, KM Akhir (kosong), Catatan (kosong)
    const values = [[date, time, '', platNomor, jenisKendaraan, namaPengemudi, tujuan, kmAwal, '', '']];

    try {
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Kendaraan!A:J', // Sesuaikan dengan rentang kolom header baru (A-J)
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: values },
        });
        document.getElementById('statusKeberangkatan').className = 'mt-3 alert alert-success';
        document.getElementById('statusKeberangkatan').innerText = `Keberangkatan ${platNomor} berhasil dicatat!`;
        this.reset(); // Reset form
        document.getElementById('jenisKendaraanBerangkat').value = ''; // Reset juga jenis kendaraan
        loadKendaraanHistory(); // Muat ulang riwayat
        populateKendaraanDropdown('pulang'); // Muat ulang dropdown kepulangan (agar log baru muncul di pilihan)
        loadDailyDashboardSummary(); // Perbarui dashboard harian
    } catch (error) {
        console.error('Error mencatat keberangkatan:', error);
        document.getElementById('statusKeberangkatan').className = 'mt-3 alert alert-danger';
        document.getElementById('statusKeberangkatan').innerText = 'Gagal mencatat keberangkatan. Cek konsol browser.';
    }
});

// Handler untuk form Kepulungan Kendaraan
document.getElementById('formKepulangan').addEventListener('submit', async function(e) {
    e.preventDefault();
    const platNomor = document.getElementById('platNomorSelectPulang').value;
    const kmAkhir = document.getElementById('kmAkhir').value;
    const { time } = getCurrentDateTime();
    const rowIndexToUpdate = this.dataset.rowIndex; // Dapatkan rowIndex yang disimpan

    if (!platNomor || !kmAkhir || !rowIndexToUpdate) {
        alert('Mohon pilih plat nomor dan masukkan KM Akhir. Pastikan ada log keberangkatan yang belum selesai.');
        return;
    }

    try {
        // Update kolom 'Waktu Pulang' (Kolom C) dan 'KM Akhir' (Kolom I)
        await gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `Kendaraan!C${rowIndexToUpdate}`, // Kolom C untuk Waktu Pulang
            valueInputOption: 'USER_ENTERED',
            resource: { values: [[time]] },
        });

        await gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `Kendaraan!I${rowIndexToUpdate}`, // Kolom I untuk KM Akhir
            valueInputOption: 'USER_ENTERED',
            resource: { values: [[kmAkhir]] },
        });

        document.getElementById('statusKepulangan').className = 'mt-3 alert alert-success';
        document.getElementById('statusKepulangan').innerText = `Kepulangan ${platNomor} berhasil dicatat!`;
        this.reset(); // Reset form
        delete this.dataset.rowIndex; // Hapus data rowIndex setelah digunakan
        // Bersihkan info yang ditampilkan
        document.getElementById('pengemudiPulangInfo').innerText = '';
        document.getElementById('tujuanPulangInfo').innerText = '';
        document.getElementById('kmAwalPulangInfo').innerText = '';
        loadKendaraanHistory(); // Muat ulang riwayat
        populateKendaraanDropdown('pulang'); // Muat ulang dropdown kepulangan (ini akan menghilangkan plat nomor yang baru saja pulang)
        loadDailyDashboardSummary(); // Perbarui dashboard harian
    } catch (error) {
        console.error('Error mencatat kepulangan:', error);
        document.getElementById('statusKepulangan').className = 'mt-3 alert alert-danger';
        document.getElementById('statusKepulangan').innerText = 'Gagal mencatat kepulangan. Cek konsol browser.';
    }
});

async function loadKendaraanHistory() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Kendaraan!A:J', // Sesuaikan dengan rentang kolom header baru (A-J)
        });
        const rows = response.result.values;
        const historyDiv = document.getElementById('kendaraanHistory');
        historyDiv.innerHTML = '';

        if (rows && rows.length > 1) {
            const table = document.createElement('table');
            table.className = 'table table-striped table-bordered';
            // Update header tabel sesuai struktur baru
            let tableHTML = '<thead><tr><th>Tanggal</th><th>Berangkat</th><th>Pulang</th><th>Plat</th><th>Jenis</th><th>Pengemudi</th><th>Tujuan</th><th>KM Awal</th><th>KM Akhir</th><th>Catatan</th></tr></thead><tbody>';
            for (let i = rows.length - 1; i >= 1; i--) { // Loop mundur untuk menampilkan data terbaru di atas
                const row = rows[i];
                tableHTML += `<tr>
                                <td>${row[0] || ''}</td>
                                <td>${row[1] || ''}</td>
                                <td>${row[2] || ''}</td>
                                <td>${row[3] || ''}</td>
                                <td>${row[4] || ''}</td>
                                <td>${row[5] || ''}</td>
                                <td>${row[6] || ''}</td>
                                <td>${row[7] || ''}</td>
                                <td>${row[8] || ''}</td>
                                <td>${row[9] || ''}</td>
                              </tr>`;
            }
            tableHTML += '</tbody>';
            table.innerHTML = tableHTML;
            historyDiv.appendChild(table);
        } else {
            historyDiv.innerHTML = '<p class="text-muted">Belum ada riwayat log kendaraan.</p>';
        }
    } catch (error) {
        console.error('Error loading kendaraan history:', error);
        document.getElementById('kendaraanHistory').innerHTML = '<p class="text-danger">Gagal memuat riwayat log kendaraan.</p>';
    }
}


// --- Daftar Tamu ---
document.getElementById('formTamu').addEventListener('submit', async function(e) {
    e.preventDefault();
    const namaTamu = document.getElementById('namaTamu').value;
    const noTelpTamu = document.getElementById('noTelpTamu').value;
    const tujuanTamu = document.getElementById('tujuanTamu').value;
    const picTamu = document.getElementById('picTamu').value;
    const { date, time } = getCurrentDateTime();
    // Struktur data sesuai header: Tanggal Masuk, Waktu Masuk, Tanggal Keluar, Waktu Keluar, Nama Tamu, Nomor Telepon, Tujuan Kunjungan, PIC yang Dikunjungi, Catatan
    const values = [[date, time, '', '', namaTamu, noTelpTamu, tujuanTamu, picTamu, '']];

    try {
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Tamu!A:I', // Sesuaikan dengan jumlah kolom header
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: values },
        });
        document.getElementById('tamuStatus').className = 'mt-3 alert alert-success';
        document.getElementById('tamuStatus').innerText = `Tamu ${namaTamu} berhasil didaftarkan!`;
        this.reset();
        loadTamuHistory();
        loadDailyDashboardSummary(); // Perbarui dashboard harian
    } catch (error) {
        console.error('Error registering guest:', error);
        document.getElementById('tamuStatus').className = 'mt-3 alert alert-danger';
        document.getElementById('tamuStatus').innerText = 'Gagal mendaftarkan tamu. Cek konsol browser.';
    }
});

async function loadTamuHistory() {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Tamu!A:I',
        });
        const rows = response.result.values;
        const historyDiv = document.getElementById('tamuActiveList');
        historyDiv.innerHTML = '';

        if (rows && rows.length > 1) {
            const table = document.createElement('table');
            table.className = 'table table-striped table-bordered';
            // Header tabel untuk tampilan riwayat tamu
            let tableHTML = '<thead><tr><th>Tgl Masuk</th><th>Wkt Masuk</th><th>Nama Tamu</th><th>Tujuan</th><th>PIC</th><th>Catatan</th></tr></thead><tbody>';
            for (let i = rows.length - 1; i >= 1; i--) { // Loop mundur untuk menampilkan data terbaru di atas
                const row = rows[i];
                // Menampilkan kolom yang relevan
                tableHTML += `<tr>
                                <td>${row[0] || ''}</td>
                                <td>${row[1] || ''}</td>
                                <td>${row[4] || ''}</td>
                                <td>${row[6] || ''}</td>
                                <td>${row[7] || ''}</td>
                                <td>${row[8] || ''}</td>
                              </tr>`;
            }
            tableHTML += '</tbody>';
            table.innerHTML = tableHTML;
            historyDiv.appendChild(table);
        } else {
            historyDiv.innerHTML = '<p class="text-muted">Belum ada tamu terdaftar.</p>';
        }
    } catch (error) {
        console.error('Error loading guest history:', error);
        document.getElementById('tamuActiveList').innerHTML = '<p class="text-danger">Gagal memuat daftar tamu.</p>';
    }
}


// --- Fungsi untuk mengisi dropdown karyawan ---
function populateKaryawanDropdown() {
    const selectElement = document.getElementById('namaKaryawanSelect');
    karyawanData.forEach(karyawan => {
        const option = document.createElement('option');
        option.value = karyawan.nama;
        option.textContent = `${karyawan.nama} (${karyawan.jabatan})`;
        selectElement.appendChild(option);
    });
}

// =========================================================
// FUNGSI BARU UNTUK MANAJEMEN JADWAL KARYAWAN
// =========================================================

function populateManajemenJadwalDropdowns() {
    const jadwalBulanSelect = document.getElementById('jadwalBulan');
    const jadwalTahunSelect = document.getElementById('jadwalTahun');
    const jadwalKaryawanSelect = document.getElementById('jadwalKaryawan');

    const months = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ];
    const currentMonth = new Date().getMonth(); // 0-indexed
    const currentYear = new Date().getFullYear();

    // Populate Bulan
    jadwalBulanSelect.innerHTML = '';
    months.forEach((month, index) => {
        const option = document.createElement('option');
        option.value = index + 1; // 1-indexed month value
        option.textContent = month;
        if (index === currentMonth) {
            option.selected = true;
        }
        jadwalBulanSelect.appendChild(option);
    });

    // Populate Tahun (misalnya 2 tahun ke belakang dan 1 tahun ke depan)
    jadwalTahunSelect.innerHTML = '';
    for (let year = currentYear - 2; year <= currentYear + 1; year++) {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        if (year === currentYear) {
            option.selected = true;
        }
        jadwalTahunSelect.appendChild(option);
    }
}

function populateManajemenJadwalKaryawanDropdown() {
    const jadwalKaryawanSelect = document.getElementById('jadwalKaryawan');
    karyawanData.forEach(karyawan => {
        const option = document.createElement('option');
        option.value = karyawan.nama;
        option.textContent = karyawan.nama;
        jadwalKaryawanSelect.appendChild(option);
    });
}

async function loadEmployeeSchedules() {
    const selectedMonth = parseInt(document.getElementById('jadwalBulan').value);
    const selectedYear = parseInt(document.getElementById('jadwalTahun').value);
    const selectedKaryawan = document.getElementById('jadwalKaryawan').value;
    const tableContainer = document.getElementById('manajemenJadwalTableContainer');
    const saveButton = document.getElementById('saveScheduleBtn');
    const jadwalStatusDiv = document.getElementById('jadwalStatus');

    tableContainer.innerHTML = '<p class="text-muted text-center">Memuat jadwal...</p>';
    jadwalStatusDiv.innerText = '';
    saveButton.style.display = 'none'; // Sembunyikan tombol simpan saat memuat

    try {
        // Ambil data jadwal yang sudah ada dari Google Sheet 'Jadwal'
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Jadwal!A:C', // Kolom: Nama Karyawan, Tanggal (DD/MM/YYYY), Status (Libur/Kerja/Tukar Libur)
        });
        const allJadwalData = response.result.values || [];
        const existingSchedules = allJadwalData.slice(1); // Lewati header

        const daysInMonth = new Date(selectedYear, selectedMonth, 0).getDate();
        const datesOfMonth = Array.from({ length: daysInMonth }, (_, i) => i + 1);

        let filteredKaryawan = karyawanData;
        if (selectedKaryawan !== 'Semua') {
            filteredKaryawan = karyawanData.filter(k => k.nama === selectedKaryawan);
        }

        if (filteredKaryawan.length === 0) {
            tableContainer.innerHTML = '<p class="text-muted text-center">Tidak ada data karyawan yang dipilih.</p>';
            return;
        }

        let tableHTML = '<table class="table table-bordered table-sm" id="employeeScheduleTable"><thead><tr>';
        tableHTML += '<th>No</th><th>Pegawai</th>';
        datesOfMonth.forEach(day => {
            tableHTML += `<th class="day-header">${day}</th>`;
        });
        tableHTML += '</tr></thead><tbody>';

        filteredKaryawan.forEach((karyawan, index) => {
            tableHTML += `<tr><td>${index + 1}</td><td>${karyawan.nama}</td>`;
            datesOfMonth.forEach(day => {
                const targetDate = `${String(day).padStart(2, '0')}/${String(selectedMonth).padStart(2, '0')}/${selectedYear}`;
                
                let defaultStatus = 'Kerja'; // Default status
                let statusClass = 'status-V'; // Default class untuk Kerja

                // Cek apakah ada jadwal yang sudah tersimpan di sheet 'Jadwal'
                const existingEntry = existingSchedules.find(entry =>
                    entry[0] === karyawan.nama && entry[1] === targetDate
                );

                let currentStatus = defaultStatus;
                if (existingEntry) {
                    currentStatus = existingEntry[2] || defaultStatus; // Ambil status dari sheet jika ada
                    // Sesuaikan statusClass berdasarkan currentStatus dari sheet
                    switch(currentStatus) {
                        case 'Libur': statusClass = 'status-off'; break;
                        case 'Tukar Libur': statusClass = 'bg-info text-white'; break; // Contoh warna untuk Tukar Libur
                        case 'Cuti': statusClass = 'status-C'; break;
                        case 'Izin': statusClass = 'status-I'; break;
                        case 'Sakit': statusClass = 'status-S'; break;
                        default: statusClass = 'status-V'; break; // Kembali ke Kerja jika tidak ada status khusus
                    }
                }

                // Buat dropdown untuk setiap hari
                let selectHTML = `<select class="form-select form-select-sm ${statusClass}" 
                                        data-employee="${karyawan.nama}" 
                                        data-date="${targetDate}">`;
                const options = ['Kerja', 'Libur', 'Tukar Libur', 'Cuti', 'Izin', 'Sakit']; // Tambahkan opsi yang relevan
                options.forEach(opt => {
                    selectHTML += `<option value="${opt}" ${opt === currentStatus ? 'selected' : ''}>${opt}</option>`;
                });
                selectHTML += '</select>';

                tableHTML += `<td>${selectHTML}</td>`;
            });
            tableHTML += '</tr>';
        });

        tableHTML += '</tbody></table>';
        tableContainer.innerHTML = tableHTML;
        saveButton.style.display = 'block'; // Tampilkan tombol simpan setelah tabel dimuat

    } catch (error) {
        console.error('Error loading employee schedules:', error);
        tableContainer.innerHTML = '<p class="text-danger text-center">Gagal memuat jadwal karyawan. Pastikan tab "Jadwal" ada di Google Sheet Anda.</p>';
        jadwalStatusDiv.className = 'mt-3 alert alert-danger';
        jadwalStatusDiv.innerText = 'Gagal memuat jadwal. Cek konsol browser.';
    }
}

async function saveEmployeeSchedules() {
    const tableContainer = document.getElementById('manajemenJadwalTableContainer');
    const jadwalStatusDiv = document.getElementById('jadwalStatus');
    jadwalStatusDiv.innerText = 'Menyimpan jadwal...';
    jadwalStatusDiv.className = 'mt-3 alert alert-info';

    const selectedMonth = parseInt(document.getElementById('jadwalBulan').value);
    const selectedYear = parseInt(document.getElementById('jadwalTahun').value);
    const selectedKaryawanFilter = document.getElementById('jadwalKaryawan').value;

    const schedulesToSave = [];
    const rowsToClear = []; // Untuk menyimpan range baris yang akan dihapus jika ada jadwal lama

    // Kumpulkan data dari tabel yang diedit
    document.querySelectorAll('#employeeScheduleTable tbody tr').forEach(row => {
        row.querySelectorAll('select').forEach(select => {
            const employeeName = select.dataset.employee;
            const date = select.dataset.date;
            const status = select.value;
            schedulesToSave.push([employeeName, date, status]);
        });
    });

    if (schedulesToSave.length === 0) {
        jadwalStatusDiv.className = 'mt-3 alert alert-warning';
        jadwalStatusDiv.innerText = 'Tidak ada jadwal untuk disimpan.';
        return;
    }

    try {
        // Langkah 1: Hapus jadwal lama untuk bulan dan karyawan yang dipilih
        const responseExisting = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Jadwal!A:C',
        });
        const existingRows = responseExisting.result.values || [];

        // Identifikasi baris yang perlu dihapus
        for (let i = 1; i < existingRows.length; i++) { // Mulai dari indeks 1 untuk melewati header
            const row = existingRows[i];
            const [empName, dateStr, status] = row;
            const [day, month, year] = dateStr.split('/').map(Number);

            // Cek apakah baris ini termasuk dalam filter bulan/tahun/karyawan yang sedang diedit
            const isMatchingMonthYear = (month === selectedMonth && year === selectedYear);
            const isMatchingEmployee = (selectedKaryawanFilter === 'Semua' || empName === selectedKaryawanFilter);

            if (isMatchingMonthYear && isMatchingEmployee) {
                rowsToClear.push(i + 1); // Simpan nomor baris (1-indexed)
            }
        }

        // Hapus baris dari bawah ke atas untuk menghindari pergeseran indeks
        if (rowsToClear.length > 0) {
            // Urutkan dari terbesar ke terkecil
            rowsToClear.sort((a, b) => b - a);
            for (const rowIndex of rowsToClear) {
                await gapi.client.sheets.spreadsheets.batchUpdate({
                    spreadsheetId: SPREADSHEET_ID,
                    resource: {
                        requests: [{
                            deleteDimension: {
                                range: {
                                    sheetId: 0, // ID sheet 'Jadwal', biasanya 0 untuk sheet pertama
                                    dimension: 'ROWS',
                                    startIndex: rowIndex - 1, // 0-indexed
                                    endIndex: rowIndex // 0-indexed, jadi endIndex adalah rowIndex itu sendiri
                                }
                            }
                        }]
                    }
                });
            }
        }

        // Langkah 2: Tambahkan jadwal baru
        await gapi.client.sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Jadwal!A:C',
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: { values: schedulesToSave },
        });

        jadwalStatusDiv.className = 'mt-3 alert alert-success';
        jadwalStatusDiv.innerText = 'Jadwal berhasil disimpan!';
        // Setelah menyimpan, muat ulang jadwal untuk memastikan tampilan terbaru
        loadEmployeeSchedules();

    } catch (error) {
        console.error('Error saving employee schedules:', error);
        jadwalStatusDiv.className = 'mt-3 alert alert-danger';
        jadwalStatusDiv.innerText = 'Gagal menyimpan jadwal. Pastikan tab "Jadwal" ada di Google Sheet Anda dan API Key memiliki izin tulis.';
    }
}
