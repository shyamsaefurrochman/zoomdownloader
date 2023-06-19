import sys
import tkinter as tk
from tkinter import ttk, filedialog
from tkcalendar import DateEntry
from urllib.request import Request, urlopen
import threading
import requests
from datetime import date
from pathlib import Path
from datetime import datetime
from tkinter import messagebox
import json

# URL dan token Zoom
ZOOM_BASE_URL = "https://api.zoom.us/v2/"
zoom_user_id = ""  # User Id atau alamat email

zoom_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJhdWQiOm51bGwsImlzcyI6IjdOQ3RCLWZyU3cySGlNWHFzNl9zM2ciLCJleHAiOjE2ODgxNDQzNDAsImlhdCI6MTY4NjkwOTgzM30.zK8acp19ZWAms5N3_voU1FGm3sszFIaQkdxBT0IeR1E"
if not zoom_token:
    print("ZOOM_TOKEN belum diatur.")
    sys.exit(1)

# Ekstensi file rekaman Zoom
recording_file_extensions = {
    "MP4": "mp4",
    "M4A": "m4a",
    "TIMELINE": "json",
    "TRANSCRIPT": "vtt",
    "CHAT": "txt",
    "CC": "vtt",
    "CSV": "csv",
}

# Daftar pertemuan
meetings = []
stop_download_flag = False  # Flag untuk menghentikan proses pengunduhan

def get_recordings():
    global meetings, zoom_user_id
    zoom_user_id = zoom_user_id_entry.get()
    meetings.clear()
    start_date = start_date_cal.get_date()
    end_date = end_date_cal.get_date()

    while end_date >= start_date:
        recording_start = start_date.strftime("%Y-%m-%d") + "T00:00:00Z"
        next_month = start_date.month + 1
        next_year = start_date.year
        if next_month == 13:
            next_month = 1
            next_year = start_date.year + 1
        recording_end = date(next_year, next_month, 1).strftime("%Y-%m-%d") + "T00:00:00Z"

        url = f"{ZOOM_BASE_URL}users/{zoom_user_id}/recordings?page_size=300&from={recording_start}&to={recording_end}"

        log_console(f"Mengambil rekaman dari {recording_start} hingga {recording_end}...")

        req = Request(url, headers={
            "content-type": "application/json",
            "authorization": "Bearer " + zoom_token
        })
        with urlopen(req) as response:
            meeting_list = response.read()
            meetings_data = json.loads(meeting_list)
            for meeting_info in meetings_data.get("meetings", []):
                meeting_recordings = []
                for recording_info in meeting_info.get("recording_files", []):
                    if recording_info.get("file_type") == "M4A":
                        continue
                    meeting_recordings.append(recording_info)
                meeting_info["recordings"] = meeting_recordings
                del meeting_info["recording_files"]
                meetings.append(meeting_info)

        start_date = date(next_year, next_month, 1)

        log_console(f"Ditemukan {len(meetings)} rekaman...")

def choose_download_folder():
    folder_selected = filedialog.askdirectory()
    download_folder_var.set(folder_selected)

def download_recordings():
    global stop_download_flag
    stop_download_flag = False
    download_path = Path(download_folder_var.get())
    download_path.mkdir(exist_ok=True)

    def download_thread():
        total_files = sum(len(meeting["recordings"]) for meeting in meetings)
        current_file = 0
        successful_downloads = 0
        failed_downloads = 0

        for index, meeting in enumerate(meetings):
            if stop_download_flag:
                break

            if meeting["topic"].startswith("Peer Interview"):
                log_console(f"Mengabaikan {meeting['topic']}")
                continue

            log_console(f"Mengolah pertemuan {meeting['topic']}")
            meeting_date = datetime.strptime(meeting["start_time"], "%Y-%m-%dT%H:%M:%SZ").strftime("%Y-%m-%d")
            path_safe_uuid = "".join(x for x in meeting['uuid'] if x.isalnum())
            meeting_path = download_path / f"{meeting_date} - {meeting['topic']} - {path_safe_uuid}"
            meeting_path.mkdir(exist_ok=True)
            for recording in meeting["recordings"]:
                if stop_download_flag:
                    break

                file_name = f"{meeting['topic']} - {recording['recording_type']} - {meeting_date}.{recording_file_extensions[recording['file_type']]}"
                recording_path = meeting_path / file_name
                if recording_path.exists() and recording_path.stat().st_size == recording["file_size"]:
                    log_console(f"Rekaman {file_name} sudah ada. Melewati...")
                    current_file += 1
                    update_progress(total_files, current_file)
                    continue

                download_url = recording["download_url"] + "?access_token=" + zoom_token
                log_console(f"Mengunduh {file_name} dari {download_url} ({recording['file_size']} byte)...")
                try:
                    download_file(download_url, recording_path)
                    log_console(f"{file_name} berhasil diunduh.")
                    successful_downloads += 1
                except Exception as e:
                    log_console(f"Terjadi kesalahan saat mengunduh {file_name}: {str(e)}")
                    failed_downloads += 1
                current_file += 1
                update_progress(total_files, current_file)

        log_console(f"Unduhan selesai. Jumlah rekaman berhasil: {successful_downloads}, gagal: {failed_downloads}.")
        messagebox.showinfo("Unduhan Selesai", f"Unduhan selesai.\nJumlah rekaman berhasil: {successful_downloads}\nJumlah rekaman gagal: {failed_downloads}")


    # Membuat dan menjalankan thread pengunduhan
    thread = threading.Thread(target=download_thread)
    thread.start()

def stop_download_func():
    global stop_download_flag
    stop_download_flag = True

def download_file(url, path):
    response = requests.get(url, stream=True)
    response.raise_for_status()
    with open(path, "wb") as file:
        total_size = int(response.headers.get('content-length', 0))
        chunk_size = 65536
        bytes_downloaded = 0
        for chunk in response.iter_content(chunk_size=chunk_size):
            if stop_download_flag:
                break

            if chunk:
                file.write(chunk)
                bytes_downloaded += len(chunk)
                update_progress(total_size, bytes_downloaded)

def update_progress(total_size, bytes_downloaded):
    percentage = (bytes_downloaded / total_size) * 100
    progress_var.set(percentage)
    root.update()

def log_console(message):
    console_text.config(state=tk.NORMAL)
    console_text.insert(tk.END, message + "\n")
    console_text.config(state=tk.DISABLED)
    console_text.see(tk.END)

# Membuat jendela utama
root = tk.Tk()
root.title("Zoom Recording Downloader")


# Entry untuk Zoom User ID
zoom_user_id_label = ttk.Label(root, text="Email Pengguna Zoom:")
zoom_user_id_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
zoom_user_id_entry = ttk.Entry(root)
zoom_user_id_entry.grid(row=0, column=1, padx=5, pady=5, sticky="we")

# Widget pemilihan tanggal
start_date_label = ttk.Label(root, text="Tanggal Mulai:")
start_date_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
start_date_cal = DateEntry(root, width=12, background='darkblue', foreground='white', date_pattern='yyyy-mm-dd')
start_date_cal.grid(row=1, column=1, padx=5, pady=5)

end_date_label = ttk.Label(root, text="Tanggal Akhir:")
end_date_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
end_date_cal = DateEntry(root, width=12, background='darkblue', foreground='white', date_pattern='yyyy-mm-dd')
end_date_cal.grid(row=2, column=1, padx=5, pady=5)

# Pemilihan folder unduhan
download_folder_var = tk.StringVar()
download_folder_label = ttk.Label(root, text="Folder Unduhan:")
download_folder_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
download_folder_entry = ttk.Entry(root, textvariable=download_folder_var, state='readonly')
download_folder_entry.grid(row=3, column=1, padx=5, pady=5, sticky="we")
download_folder_button = ttk.Button(root, text="Pilih", command=choose_download_folder)
download_folder_button.grid(row=3, column=2, padx=5, pady=5)

# Tombol untuk mengambil rekaman
get_recordings_button = ttk.Button(root, text="Ambil Rekaman", command=get_recordings)
get_recordings_button.grid(row=4, column=0, padx=5, pady=5, sticky="w")

# Tombol untuk mengunduh rekaman
download_frame = ttk.Frame(root)
download_frame.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky="we")

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(download_frame, variable=progress_var, maximum=100)
progress_bar.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

download_button = ttk.Button(download_frame, text="Unduh Rekaman", command=download_recordings)
download_button.pack(side=tk.LEFT, padx=5, pady=5)

stop_download_button = ttk.Button(download_frame, text="Hentikan Unduhan", command=stop_download_func)
stop_download_button.pack(side=tk.LEFT, padx=5, pady=5)

# Console log
console_frame = ttk.Frame(root)
console_frame.grid(row=6, column=0, columnspan=3, padx=5, pady=5, sticky="we")

console_label = ttk.Label(console_frame, text="Console Log:")
console_label.pack()

console_text = tk.Text(console_frame, state=tk.DISABLED)
console_text.pack(fill=tk.BOTH, expand=True)

# Hide/Show Console Button
console_hidden = False
def toggle_console():
    global console_hidden
    if console_hidden:
        console_frame.grid()
        console_hidden = False
    else:
        console_frame.grid_remove()
        console_hidden = True

toggle_console_button = ttk.Button(root, text="Sembunyikan/Kembalikan Console", command=toggle_console)
toggle_console_button.grid(row=7, column=0, columnspan=3, padx=5, pady=5, sticky="we")


# Menampilkan jendela
root.mainloop()
