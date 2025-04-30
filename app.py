from flask import Flask, render_template, request, redirect, url_for, send_from_directory, jsonify
import pandas as pd
from datetime import datetime
import os
import json

app = Flask(__name__)

# Function to load data from Excel file
def load_student_data():
    try:
        df = pd.read_excel('data_siswa.xlsx')

        if 'nisn' in df.columns:
            df['nisn'] = df['nisn'].astype(str).str.strip()
        
        if 'tanggal_lahir' in df.columns:
            df['tanggal_lahir'] = pd.to_datetime(df['tanggal_lahir'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d')
        
        if 'status_kelulusan' in df.columns:
            df['status_kelulusan'] = df['status_kelulusan'].str.upper().str.strip()

        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return pd.DataFrame()

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

@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join('static/surat_kelulusan', filename)
    if os.path.exists(file_path):
        return send_from_directory('static/surat_kelulusan', filename, as_attachment=True)
    else:
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
            df = load_student_data()
            
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


if __name__ == '__main__':
    app.run(debug=True)