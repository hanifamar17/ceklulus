from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from datetime import datetime

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
    return render_template('index.html')

@app.route('/cek-kelulusan', methods=['POST', "GET"])
def cek_kelulusan():
    if request.method == 'POST':
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
                                            error=None)
                    else:
                        return render_template('index.html', 
                                            hasil='tidak_lulus', 
                                            data=data_siswa,
                                            error=None)
                else:
                    # Tanggal lahir tidak cocok
                    return render_template('index.html', 
                                        hasil=None, 
                                        data=None,
                                        error='Tanggal lahir tidak sesuai dengan data NISN. Mohon periksa kembali.')
            else:
                # NISN tidak ditemukan
                return render_template('index.html', 
                                    hasil=None, 
                                    data=None,
                                    error='NISN tidak ditemukan. Mohon periksa kembali nomor NISN Anda.')
                                    
        except Exception as e:
            print(f"Error processing request: {e}")
            return render_template('index.html', 
                                hasil=None, 
                                data=None, 
                                error='Terjadi kesalahan dalam memproses data. Mohon coba lagi.')
    else:
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)