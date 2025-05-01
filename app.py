from flask import Flask, render_template, request, redirect, url_for, send_from_directory, jsonify, send_file
import pandas as pd
from datetime import datetime
import os
import json
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
import io
from dotenv import load_dotenv

app = Flask(__name__)
secret_key = os.urandom(24)  # Generate a random secret key for session management
app.secret_key = secret_key

# Load .env jika dijalankan secara lokal
if os.getenv("VERCEL") is None:  # Deteksi bukan di Vercel
    from dotenv import load_dotenv
    load_dotenv()

# Buat file credentials.json dari ENV
credentials_json = os.getenv("CREDENTIALS_JSON")

if credentials_json:
    # Tentukan lokasi file yang bisa ditulis (di Vercel /tmp)
    CREDENTIALS_FILE = '/tmp/credentials.json' if os.getenv('VERCEL') else 'credentials.json'
    
    # Menulis file credentials.json (di /tmp jika di Vercel)
    with open(CREDENTIALS_FILE, 'w') as f:
        f.write(credentials_json)
else:
    raise EnvironmentError("CREDENTIALS_JSON tidak ditemukan dalam environment variables")

# SCOPES dan folder_id diambil dari ENV
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
FOLDER_ID_SISWA = os.getenv('FOLDER_ID_SISWA') 
FOLDER_ID_SURAT = os.getenv('FOLDER_ID_SURAT')
FILE_NAME_SISWA = 'data_siswa.xlsx'

# Tentukan cache direktori, pastikan di /tmp di Vercel
CACHE_DIR = '/tmp/cache' if os.getenv('VERCEL') else './cache'
os.makedirs(CACHE_DIR, exist_ok=True)


# Fungsi autentikasi ke Google Drive
def authenticate_google_drive():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)


# Fungsi mendapatkan file ID dari nama file & folder
def get_file_id(service, folder_id, file_name):
    query = f"name = '{file_name}' and '{folder_id}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    return items[0]['id'] if items else None

# Fungsi untuk memuat data siswa dan menyimpan ke cache
def load_student_data_from_drive():
    try:
        # Nama file yang digunakan untuk cache
        cached_file_path = os.path.join(CACHE_DIR, FILE_NAME_SISWA)
        
        # Jika file sudah ada di cache, langsung baca dari cache
        if os.path.exists(cached_file_path):
            print(f"File {FILE_NAME_SISWA} sudah di-cache. Menggunakan data cache.")
            try:
                with open(cached_file_path, 'rb') as f:
                    file_data = f.read()
                    df = pd.read_excel(io.BytesIO(file_data))
            except Exception as e:
                print(f"Error reading cached file: {e}")
                print("Cache kemungkinan korup. Menghapus dan mencoba mengunduh ulang dari Google Drive...")
                os.remove(cached_file_path)  # Hapus file cache rusak
                return load_student_data_from_drive()  # Coba ulang
        else:
            # Jika file belum ada di cache, ambil dari Google Drive
            print(f"File {FILE_NAME_SISWA} belum ada di cache. Mengunduh dari Drive...")
            service = authenticate_google_drive()
            file_id = get_file_id(service, FOLDER_ID_SISWA, FILE_NAME_SISWA)
            if not file_id:
                print("File tidak ditemukan di folder siswa.")
                return pd.DataFrame()

            request = service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)

            done = False
            while not done:
                status, done = downloader.next_chunk()

            fh.seek(0)
            try:
                # Baca langsung dari stream untuk file yang diunduh
                df = pd.read_excel(fh)
            except Exception as e:
                print(f"Error reading downloaded file: {e}")
                return pd.DataFrame()  # Mengembalikan DataFrame kosong jika gagal membaca file

            # Simpan file yang diunduh ke cache
            with open(cached_file_path, 'wb') as f:
                f.write(fh.getvalue())  # Pastikan menggunakan getvalue() untuk menulis ke file
            print(f"File {FILE_NAME_SISWA} berhasil diunduh dan disimpan di cache.")

        # Cek kolom yang ada di DataFrame
        print(f"Kolom yang ditemukan di file: {df.columns.tolist()}")

        # Normalisasi kolom jika ada
        if 'nisn' in df.columns:
            df['nisn'] = df['nisn'].astype(str).str.strip()
        if 'tanggal_lahir' in df.columns:
            df['tanggal_lahir'] = pd.to_datetime(df['tanggal_lahir'], errors='coerce').dt.strftime('%Y-%m-%d')
        if 'status_kelulusan' in df.columns:
            df['status_kelulusan'] = df['status_kelulusan'].str.upper().str.strip()

        return df

    except Exception as e:
        print(f"Error loading student data: {e}")
        return pd.DataFrame()

# Fungsi untuk mengunduh dan menyimpan file ke cache lokal
def download_file_from_drive(file_name, folder_id):
    # Cek apakah file sudah ada di cache
    cached_file = download_file_from_cache(file_name)
    if cached_file:
        return cached_file

    try:
        service = authenticate_google_drive()
        file_id = get_file_id(service, folder_id, file_name)
        if not file_id:
            return None

        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)

        # Tentukan ekstensi file dan cache
        file_extension = file_name.split('.')[-1].lower()
        local_cache_path = f'./cache/{file_name}'

        # Caching berdasarkan ekstensi file
        if file_extension in ['pdf', 'xlsx']:
            with open(local_cache_path, 'wb') as f:
                f.write(fh.getvalue())
        
        return fh
    except Exception as e:
        print(f"Error downloading file: {e}")
        return None

# Fungsi untuk memeriksa file yang sudah ter-cache
def download_file_from_cache(file_name):
    cache_file_path = f'./cache/{file_name}'
    if os.path.exists(cache_file_path):
        with open(cache_file_path, 'rb') as f:
            return io.BytesIO(f.read())
    return None

@app.route('/download/<filename>')
def download(filename):
    # Cek apakah file sudah ada di cache
    cache_file_path = f'./cache/{filename}'
    if os.path.exists(cache_file_path):
        return send_file(cache_file_path, as_attachment=True, download_name=filename)
    
    # Jika file belum ada di cache, unduh dari Google Drive
    file_stream = download_file_from_drive(filename, FOLDER_ID_SURAT)
    if file_stream:
        return send_file(file_stream, as_attachment=True, download_name=filename)
    else:
        return jsonify({'success': False, 'message': 'File tidak ditemukan'}), 404
    
@app.route('/')
def index():
    form_aktif, next_schedule = get_schedule_status()
    server_current_timestamp = int(datetime.now().timestamp() * 1000)  # Convert to milliseconds
    server_target_timestamp = 0

    if next_schedule:
        # Konversi tanggal target ke timestamp
        target_date = next_schedule['mulai_obj']
        server_target_timestamp = int(target_date.timestamp() * 1000)  # Target timestamp dalam ms

    return render_template('index.html', 
                           hasil=None, 
                           data=None, 
                           error=None, 
                           form_aktif=form_aktif, 
                           next_schedule=next_schedule, 
                           server_current_timestamp=server_current_timestamp,
                           server_target_timestamp=server_target_timestamp)
    return jsonify({'success': False, 'message': 'File tidak ditemukan'}), 404

@app.route('/cek-kelulusan', methods=['POST', "GET"])
def cek_kelulusan():
    form_aktif, next_schedule = get_schedule_status()
    server_current_timestamp = int(datetime.now().timestamp() * 1000)  # Timestamp saat ini dalam ms
    server_target_timestamp = 0
    
    if next_schedule:
        # Konversi tanggal target ke timestamp
        target_date = next_schedule['mulai_obj']
        server_target_timestamp = int(target_date.timestamp() * 1000)  # Target timestamp dalam ms

    if request.method == 'POST':
        if not form_aktif:
            return render_template('index.html', 
                                   hasil=None, 
                                   data=None, 
                                   error='Form tidak tersedia saat ini.', 
                                   form_aktif=form_aktif,
                                   next_schedule=next_schedule,
                                   server_current_timestamp=server_current_timestamp,
                                   server_target_timestamp=server_target_timestamp)
        
        nisn = request.form['nisn']
        tanggal_lahir_str = request.form['tanggal_lahir']
        
        try:
            # Konversi input tanggal lahir dari format YYYY-MM-DD (HTML date input)
            # ke format yang cocok untuk perbandingan dengan data
            tanggal_lahir_obj = datetime.strptime(tanggal_lahir_str, '%Y-%m-%d')
            tanggal_lahir_formatted = tanggal_lahir_obj.strftime('%Y-%m-%d')
            
            # Load data siswa dari Excel
            df = load_student_data_from_drive()
            
            # Pastikan format tanggal lahir di database konsisten
            if not df.empty and 'tanggal_lahir' in df.columns:
                df['tanggal_lahir'] = pd.to_datetime(df['tanggal_lahir']).dt.strftime('%Y-%m-%d')
            
            # Cari siswa berdasarkan NISN
            siswa = df[df['nisn'] == nisn]
            
            if not siswa.empty:
                # Ambil data siswa pertama yang cocok dengan NISN
                data_siswa = siswa.iloc[0].to_dict()
                
                # Periksa apakah tanggal lahir cocok
                tanggal_siswa = data_siswa['tanggal_lahir']
                tanggal_siswa_obj = datetime.strptime(tanggal_siswa, '%Y-%m-%d')
                
                # Format tanggal untuk ditampilkan dengan format Indonesia
                data_siswa['tanggal_lahir_format'] = tanggal_siswa_obj.strftime('%d %B %Y')
                
                # Cek kecocokan tanggal lahir (hanya compare tahun, bulan, dan hari)
                if tanggal_siswa_obj.year == tanggal_lahir_obj.year and \
                   tanggal_siswa_obj.month == tanggal_lahir_obj.month and \
                   tanggal_siswa_obj.day == tanggal_lahir_obj.day:
                    
                    # Render halaman hasil sesuai status kelulusan
                    if data_siswa['status_kelulusan'].upper() == 'LULUS':
                        return render_template('index.html', 
                                            hasil='lulus', 
                                            data=data_siswa,
                                            error=None,
                                            form_aktif=form_aktif,
                                            next_schedule=next_schedule,
                                            server_current_timestamp=server_current_timestamp,
                                            server_target_timestamp=server_target_timestamp)
                    else:
                        return render_template('index.html', 
                                            hasil='tidak_lulus', 
                                            data=data_siswa,
                                            error=None,
                                            form_aktif=form_aktif,
                                            next_schedule=next_schedule,
                                            server_current_timestamp=server_current_timestamp,
                                            server_target_timestamp=server_target_timestamp)
                else:
                    # Tanggal lahir tidak cocok
                    return render_template('index.html', 
                                        hasil=None, 
                                        data=None,
                                        error='Tanggal lahir tidak sesuai dengan data NISN. Mohon periksa kembali.',
                                        form_aktif=form_aktif,
                                        next_schedule=next_schedule,
                                        server_current_timestamp=server_current_timestamp,
                                        server_target_timestamp=server_target_timestamp)
            else:
                # NISN tidak ditemukan
                return render_template('index.html', 
                                    hasil=None, 
                                    data=None,
                                    error='NISN tidak ditemukan. Mohon periksa kembali nomor NISN Anda.',
                                    form_aktif=form_aktif,
                                    next_schedule=next_schedule,
                                    server_current_timestamp=server_current_timestamp,
                                    server_target_timestamp=server_target_timestamp)
                                    
        except Exception as e:
            print(f"Error processing request: {e}")
            return render_template('index.html', 
                                hasil=None, 
                                data=None, 
                                error='Terjadi kesalahan dalam memproses data. Mohon coba lagi.',
                                form_aktif=form_aktif,
                                next_schedule=next_schedule,
                                server_current_timestamp=server_current_timestamp,
                                server_target_timestamp=server_target_timestamp)
    else:
        return render_template('index.html', 
                              hasil=None, 
                              data=None, 
                              error=None, 
                              form_aktif=form_aktif,
                              next_schedule=next_schedule,
                              server_current_timestamp=server_current_timestamp,
                              server_target_timestamp=server_target_timestamp)

SCHEDULE_FILE = 'schedule.json'

def load_schedule():
    if not os.path.exists(SCHEDULE_FILE):
        return []
    with open(SCHEDULE_FILE, 'r') as f:
        return json.load(f)

def save_schedule(schedule_baru):
    schedule = load_schedule()
    schedule.append(schedule_baru)
    with open(SCHEDULE_FILE, 'w') as f:
        json.dump(schedule, f, indent=4)

@app.route("/admin/schedule", methods=["GET", "POST"])
def atur_schedule():
    if request.method == "POST":
        mulai = request.form.get("mulai")
        berakhir = request.form.get("berakhir")
        keterangan = request.form.get("keterangan")

        save_schedule({
            "mulai": mulai,
            "berakhir": berakhir,
            "keterangan": keterangan,
            "waktu_input": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

        return redirect(url_for("atur_schedule"))

    schedule = load_schedule()
    return render_template("schedule.html", schedule=schedule)

@app.route("/admin/schedule/delete/<int:index>", methods=["POST"])
def hapus_schedule(index):
    schedule = load_schedule()
    if 0 <= index < len(schedule):
        del schedule[index]
        with open(SCHEDULE_FILE, 'w') as f:
            json.dump(schedule, f, indent=4)
    return redirect(url_for("atur_schedule"))

def get_schedule_status():
    with open('schedule.json') as f:
        data = json.load(f)
    now = datetime.now()
    form_aktif = None
    next_schedule = None
    for schedule in data:
        mulai = datetime.strptime(schedule['mulai'], '%Y-%m-%dT%H:%M')
        berakhir = datetime.strptime(schedule['berakhir'], '%Y-%m-%dT%H:%M')
        if mulai <= now <= berakhir:
            form_aktif = schedule
            # Convert datetime objects to strings for template
            form_aktif['mulai_obj'] = mulai
            form_aktif['berakhir_obj'] = berakhir
            break
        elif now < mulai:
            if not next_schedule or mulai < datetime.strptime(next_schedule['mulai'], '%Y-%m-%dT%H:%M'):
                next_schedule = schedule
                # Convert datetime objects to strings for template
                next_schedule['mulai_obj'] = mulai
                next_schedule['berakhir_obj'] = berakhir
    return form_aktif, next_schedule

@app.template_filter('format_datetime')
def format_datetime(value):
    bulan = {
        "01": "Januari", "02": "Februari", "03": "Maret",
        "04": "April", "05": "Mei", "06": "Juni",
        "07": "Juli", "08": "Agustus", "09": "September",
        "10": "Oktober", "11": "November", "12": "Desember"
    }
    
    # Handle both datetime object and string inputs
    if isinstance(value, str):
        dt = datetime.strptime(value, '%Y-%m-%dT%H:%M')
    else:
        dt = value
        
    return f"{dt.day} {bulan[dt.strftime('%m')]} {dt.year} {dt.strftime('%H:%M')} WIB"



# Fungsi untuk memulai cache
def get_all_files(service, folder_id):
    try:
        query = f"'{folder_id}' in parents and trashed = false"
        results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
        items = results.get('files', [])
        if not items:
            print("No files found.")
            return []
        return items
    except Exception as e:
        print("Error getting files:", e)
        return []

def warm_up_cache_for_files(folder_id):
    service = authenticate_google_drive()
    files = get_all_files(service, folder_id)

    for file in files:
        file_name = file['name']
        print(f"Memulai pre-caching file: {file_name}")
        file_stream = download_file_from_drive(file_name, folder_id)
        if file_stream:
            print(f"File {file_name} berhasil di-cache.")
        else:
            print(f"File {file_name} gagal diunduh dan tidak bisa di-cache.")

# Memanggil fungsi pre-caching langsung sebelum aplikasi dimulai
def pre_cache_student_data():
    load_student_data_from_drive()
    
def pre_cache_files():
    folder_id = FOLDER_ID_SURAT  # Gunakan folder ID yang sesuai
    pre_cache_student_data()
    warm_up_cache_for_files(folder_id)

if __name__ == '__main__':
    pre_cache_files() # Pre-cache files saat aplikasi dimulai
    app.run(debug=True)