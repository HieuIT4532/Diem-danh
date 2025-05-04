import os
import cv2
import time
import unicodedata
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import sys
import pandas as pd
import retrying
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import google.auth.transport.requests
import requests
from tkinter import Tk, Label, Button, messagebox, Frame, Entry
from PIL import ImageTk, Image
import base64
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
import logging
from datetime import datetime

# Cấu hình logging
logging.basicConfig(filename='app.log', level=logging.DEBUG, encoding='utf-8')

# Cấu hình
SCOPES_OAUTH = ['openid', 'https://www.googleapis.com/auth/userinfo.profile', 'https://www.googleapis.com/auth/userinfo.email']
SCOPES_API = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
FOLDER_ACCOUNT_FILE = "dang_ky.json"
SERVICE_ACCOUNT_FILE = "diem_danh.json"
FOLDER_ID = "1wQl3N4jUmoJRks_JsM6JyTOUAFMIHtZP"
SHEET_ID = "1J0g-vhbQ8TiCBqfM-yi5V4z2-tKF-oGz-lsV35rrnMs"
SENDGRID_API_KEY = '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'

creds_global = None

# Khởi tạo dịch vụ Google API
credss = service_account.Credentials.from_service_account_file(FOLDER_ACCOUNT_FILE, scopes=SCOPES_API)
drive_service = build("drive", "v3", credentials=credss)

creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES_API)
gspread_client = gspread.authorize(creds)
sheet = gspread_client.open_by_key(SHEET_ID).sheet1

name_mapping = {}

# Đăng nhập Google
def get_user_info(creds):
    session = requests.Session()
    auth_req = google.auth.transport.requests.Request(session=session)
    creds.refresh(auth_req)
    resp = session.get('https://www.googleapis.com/oauth2/v1/userinfo', params={'alt': 'json'}, headers={'Authorization': f'Bearer {creds.token}'})
    return resp.json()

def login_with_google(window_to_close=None):
    global creds_global
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds_global = pickle.load(token)

    if not creds_global or not creds_global.valid:
        if creds_global and creds_global.expired and creds_global.refresh_token:
            creds_global.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('DangNhapBangTk.json', SCOPES_OAUTH)
            creds_global = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds_global, token)

    user_info = get_user_info(creds_global)
    email = user_info.get('email', 'Không tìm thấy email')

    if email == 'Không tìm thấy email':
        messagebox.showerror("Lỗi", "Không thể lấy email người dùng.")
        return False, None
    else:
        with open('email.txt', 'w') as f:
            f.write(email)
        messagebox.showinfo("Thành công", f"Đăng nhập thành công với Gmail:\n{email}")
        
        if window_to_close:
            window_to_close.destroy()
        return True, email

def show_login_window():
    root = Tk()
    root.title("Đăng nhập")
    
    # Kích thước cửa sổ
    window_width = 400
    window_height = 200
    
    # Lấy kích thước màn hình
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # Tính toán tọa độ để căn giữa
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    
    # Đặt kích thước và vị trí cửa sổ
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.configure(bg="#f5f5f5")

    # Tải và resize logo Google
    try:
        logo_google = Image.open("logo_google.png")
        logo_google = logo_google.resize((24, 24))
        logo_google = ImageTk.PhotoImage(logo_google)
    except FileNotFoundError:
        logo_google = None

    # Tiêu đề
    title = Label(root, text="điểm danh bằng Google", font=("Arial", 16, "bold"), bg="#f5f5f5")
    title.pack(pady=20)

    # Nút đăng nhập
    login_button = Button(
        root,
        text="  Đăng nhập với Google",
        image=logo_google,
        compound="left",
        font=("Arial", 13),
        bg="white",
        fg="black",
        activebackground="#e0e0e0",
        padx=10,
        pady=8,
        borderwidth=1,
        relief="ridge",
        command=lambda: login_and_open_main(root)
    )
    if logo_google:
        login_button.image = logo_google
    login_button.pack(pady=10)

    # Hiệu ứng hover cho nút đăng nhập
    def on_enter_login(e):
        login_button.config(bg="#f0f0f0")
    def on_leave_login(e):
        login_button.config(bg="white")
    login_button.bind("<Enter>", on_enter_login)
    login_button.bind("<Leave>", on_leave_login)

    def login_and_open_main(window):
        success, email = login_with_google(window)
        if success:
            show_main_window(email=email)

    if os.path.exists('token.pickle'):
        login_and_open_main(root)

    root.mainloop()

# Hàm viết hoa chữ cái đầu
def capitalize_name(name):
    return ' '.join(word.capitalize() for word in name.split())

# Chuyển từ có dấu thành không dấu
def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    only_ascii = nfkd_form.encode('ASCII', 'ignore').decode('utf-8')
    return only_ascii

# Kiểm tra ảnh hợp lệ
def is_valid_image(file_path):
    try:
        with Image.open(file_path) as img:
            img.verify()
        return True
    except Exception:
        return False

# Tải ảnh từ Drive
def download_images_from_drive():
    temp_folder = "temp_faces"
    os.makedirs(temp_folder, exist_ok=True)

    query = f"'{FOLDER_ID}' in parents"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()

    for file in results.get("files", []):
        try:
            request = drive_service.files().get_media(fileId=file["id"])
            original_name = os.path.splitext(file["name"])[0]
            filename = remove_accents(original_name).lower() + '.jpg'
            name_mapping[os.path.splitext(filename)[0]] = original_name

            file_path = os.path.join(temp_folder, filename)
            with open(file_path, "wb") as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()

            if not is_valid_image(file_path):
                logging.warning("Ảnh không hợp lệ: %s", file_path)
                os.remove(file_path)
        except Exception as e:
            logging.error("Error processing file %s: %s", file['name'], str(e))
            continue
    return temp_folder

# Upload ảnh lên Drive
def upload_to_drive(file_path, file_name):
    try:
        file_metadata = {"name": file_name, "parents": [FOLDER_ID]}
        media = MediaFileUpload(file_path, mimetype="image/jpeg")
        file = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()
        return file.get("id")
    except Exception as e:
        logging.error("Error uploading to Drive: %s", str(e))
        return None

# Đăng ký khuôn mặt
def register_face():
    name_window = Tk()
    name_window.title("Nhập tên")
    
    # Kích thước cửa sổ
    window_width = 300
    window_height = 150
    
    # Lấy kích thước màn hình
    screen_width = name_window.winfo_screenwidth()
    screen_height = name_window.winfo_screenheight()
    
    # Tính toán tọa độ để căn giữa
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    
    # Đặt kích thước và vị trí cửa sổ
    name_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
    name_window.configure(bg="#f5f5f5")

    Label(name_window, text="Nhập tên của bạn:", font=("Arial", 12), bg="#f5f5f5").pack(pady=10)
    name_entry = Entry(name_window, font=("Arial", 12), width=20)
    name_entry.pack(pady=10)

    def submit_name():
        original_name = name_entry.get().strip()
        if not original_name:
            messagebox.showerror("Lỗi", "Vui lòng nhập tên.")
            return
        capitalized_name = capitalize_name(original_name)
        name_window.destroy()

        update_status("Đang chụp ảnh khuôn mặt...")
        cap = cv2.VideoCapture(0)
        ret, frame = cap.read()
        cap.release()
        if not ret:
            messagebox.showerror("Lỗi", "Không thể mở camera.")
            update_status("Lỗi: Không thể mở camera.")
            return

        path = "temp.jpg"
        cv2.imwrite(path, frame)

        file_id = upload_to_drive(path, f"{capitalized_name}.jpg")
        if file_id:
            os.remove(path)
            #sheet.append_row([capitalized_name, capitalized_name])
            messagebox.showinfo("Thành công", f"Đã đăng ký khuôn mặt cho {capitalized_name}.")
            update_status(f"Đã đăng ký: {capitalized_name}")
        else:
            messagebox.showerror("Lỗi", "Lỗi khi đăng ký khuôn mặt.")
            update_status("Lỗi: Không thể đăng ký khuôn mặt.")

    confirm_button = Button(
        name_window,
        text="Xác nhận",
        font=("Arial", 12),
        bg="#4CAF50",
        fg="white",
        activebackground="#45A049",
        command=submit_name
    )
    confirm_button.pack(pady=10)

    # Hiệu ứng hover cho nút xác nhận
    def on_enter_confirm(e):
        confirm_button.config(bg="#66BB6A")
    def on_leave_confirm(e):
        confirm_button.config(bg="#4CAF50")
    confirm_button.bind("<Enter>", on_enter_confirm)
    confirm_button.bind("<Leave>", on_leave_confirm)

    name_window.mainloop()

# Điểm danh bằng khuôn mặt
def mark_attendance():
    update_status("Đang chụp ảnh để điểm danh...")
    cap = cv2.VideoCapture(0)
    ret, frame = cap.read()
    cap.release()

    if not ret:
        messagebox.showerror("Lỗi", "Lỗi khi chụp ảnh từ camera.")
        update_status("Lỗi: Không thể chụp ảnh.")
        return

    temp_image_path = "temp.jpg"
    cv2.imwrite(temp_image_path, frame)
    temp_folder = download_images_from_drive()

    match_found = False
    person_safe_name = ""
    person_original_name = ""
    best_match_count = 0

    for filename in os.listdir(temp_folder):
        known_image_path = os.path.join(temp_folder, filename)
        if not is_valid_image(known_image_path):
            logging.warning("Ảnh không hợp lệ: %s", known_image_path)
            continue

        img1 = cv2.imread(known_image_path, cv2.IMREAD_GRAYSCALE)
        img2 = cv2.imread(temp_image_path, cv2.IMREAD_GRAYSCALE)

        if img1 is None or img2 is None:
            logging.warning("Không thể đọc ảnh: %s", known_image_path)
            continue

        orb = cv2.ORB_create()
        kp1, des1 = orb.detectAndCompute(img1, None)
        kp2, des2 = orb.detectAndCompute(img2, None)

        if des1 is None or des2 is None:
            continue

        bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
        matches = bf.match(des1, des2)

        if len(matches) > best_match_count:
            best_match_count = len(matches)
            person_safe_name = os.path.splitext(filename)[0]
            person_original_name = name_mapping.get(person_safe_name, person_safe_name)
            match_found = True

    os.remove(temp_image_path)
    for file in os.listdir(temp_folder):
        os.remove(os.path.join(temp_folder, file))
    os.rmdir(temp_folder)

    if match_found:
        sheet_data = sheet.get_all_values()
        current_time = datetime.now().strftime("%H:%M:%S")
        existing_names = [row[0] for row in sheet_data[1:]]

        if person_original_name in existing_names:
            messagebox.showinfo("Thông báo", f"{person_original_name} đã điểm danh trước đó.")
            update_status(f"{person_original_name} đã điểm danh trước đó.")
        else:
            sheet.append_row([person_original_name, current_time])
            messagebox.showinfo("Thành công", f"{person_original_name} đã điểm danh lúc {current_time}.")
            update_status(f"Điểm danh: {person_original_name} lúc {current_time}")
    else:
        messagebox.showerror("Lỗi", "Không tìm thấy khuôn mặt trong danh sách.")
        update_status("Lỗi: Không tìm thấy khuôn mặt.")

# Xuất file học sinh chưa điểm danh và điểm danh muộn, sau đó gửi email
@retrying.retry(stop_max_attempt_number=3, wait_fixed=2000)
def export_absent_list():
    update_status("Đang tạo danh sách điểm danh...")
    try:
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        SERVICE_ACCOUNT_FILE = "diem_danh.json"

        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)

        sheet_lop = client.open_by_key("1pzXBIopl1kkpPYrV__DmLyVXYNW4BtyRMIRLkOx5Ock").sheet1
        sheet_diemdanh = client.open_by_key("1J0g-vhbQ8TiCBqfM-yi5V4z2-tKF-oGz-lsV35rrnMs").sheet1

        lop_data = sheet_lop.get_all_values()[1:]
        diemdanh_data = sheet_diemdanh.get_all_values()[1:]

        df_lop = pd.DataFrame(lop_data, columns=["Họ tên"])
        df_lop["STT"] = df_lop.index + 1
        df_diemdanh = pd.DataFrame(diemdanh_data, columns=["Họ tên", "Thời gian"])

        df_lop["Tên chuẩn"] = df_lop["Họ tên"].str.strip().str.lower()
        df_diemdanh["Tên chuẩn"] = df_diemdanh["Họ tên"].str.strip().str.lower()

        chua_diemdanh = df_lop[~df_lop["Tên chuẩn"].isin(df_diemdanh["Tên chuẩn"])][["STT", "Họ tên"]]
        df_diemdanh["Giờ"] = pd.to_datetime(df_diemdanh["Thời gian"], format="%H:%M:%S", errors='coerce').dt.time
        qua_gio = df_diemdanh[df_diemdanh["Giờ"] > datetime.strptime("07:00", "%H:%M").time()][["Họ tên", "Thời gian"]]

        output_file = "bao_cao_diem_danh.xlsx"
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            chua_diemdanh.to_excel(writer, sheet_name="Chưa điểm danh", index=False)
            qua_gio.to_excel(writer, sheet_name="Điểm danh muộn", index=False)

        logging.info("Đã tạo file: %s", output_file)

        try:
            with open('email.txt', 'r') as f:
                email_list = [line.strip() for line in f if line.strip()]
        except FileNotFoundError:
            messagebox.showerror("Lỗi", "Không tìm thấy file email.txt")
            update_status("Lỗi: Không tìm thấy file email.txt")
            return

        if email_list and os.path.isfile(output_file):
            try:
                with open(output_file, 'rb') as f:
                    encoded_file = base64.b64encode(f.read()).decode()

                attachment = Attachment()
                attachment.file_content = FileContent(encoded_file)
                attachment.file_type = FileType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                attachment.file_name = FileName(output_file)
                attachment.disposition = Disposition('attachment')

                sg = SendGridAPIClient(SENDGRID_API_KEY)
                for email in email_list:
                    message = Mail(
                        from_email='hohoanganh20092021@gmail.com',
                        to_emails=email,
                        subject='Báo cáo điểm danh',
                        html_content='<strong>Vui lòng xem file đính kèm để biết danh sách học sinh chưa điểm danh và điểm danh muộn.</strong>'
                    )
                    message.attachment = attachment
                    response = sg.send(message)
                    logging.info("Đã gửi email đến %s, mã trạng thái: %s", email, response.status_code)
                messagebox.showinfo("Thành công", "Đã gửi báo cáo điểm danh qua email.")
                update_status("Đã gửi báo cáo qua email.")
            except Exception as e:
                messagebox.showerror("Lỗi", "Lỗi khi gửi email báo cáo.")
                update_status("Lỗi: Không thể gửi email.")
        else:
            messagebox.showerror("Lỗi", "Không có email để gửi hoặc file báo cáo không tồn tại.")
            update_status("Lỗi: Không có email hoặc file báo cáo.")

    except Exception as e:
        messagebox.showerror("Lỗi", "Đã xảy ra lỗi khi xuất danh sách.")
        update_status("Lỗi: Không thể xuất danh sách.")
        logging.error("Lỗi khi xuất danh sách: %s", str(e))

# Hàm cập nhật trạng thái
def update_status(message):
    global status_label
    if status_label:
        status_label.config(text=message)

# Giao diện chính
def show_main_window(email=None):
    global register_button, attendance_button, export_button, exit_button, status_label
    root = Tk()
    root.title("Hệ thống điểm danh")
    
    # Kích thước cửa sổ
    window_width = 500
    window_height = 550
    
    # Lấy kích thước màn hình
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # Tính toán tọa độ để căn giữa
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    
    # Đặt kích thước và vị trí cửa sổ
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    root.configure(bg="#f5f5f5")

    # Tiêu đề
    title = Label(root, text="Hệ thống điểm danh", font=("Arial", 20, "bold"), bg="#f5f5f5")
    title.pack(pady=20)

    # Frame chứa các nút
    button_frame = Frame(root, bg="#f5f5f5")
    button_frame.pack(pady=20)

    # Nút Đăng ký khuôn mặt
    register_button = Button(
        button_frame,
        text="Đăng ký khuôn mặt",
        font=("Arial", 14),
        bg="#4CAF50",
        fg="white",
        activebackground="#45A049",
        padx=20,
        pady=10,
        command=register_face
    )
    register_button.pack(pady=10, fill="x")
    def on_enter_register(e):
        register_button.config(bg="#66BB6A")
    def on_leave_register(e):
        register_button.config(bg="#4CAF50")
    register_button.bind("<Enter>", on_enter_register)
    register_button.bind("<Leave>", on_leave_register)

    # Nút Điểm danh
    attendance_button = Button(
        button_frame,
        text="Điểm danh",
        font=("Arial", 14),
        bg="#FFC107",
        fg="black",
        activebackground="#FFB300",
        padx=20,
        pady=10,
        command=mark_attendance
    )
    attendance_button.pack(pady=10, fill="x")
    def on_enter_attendance(e):
        attendance_button.config(bg="#FFCA28")
    def on_leave_attendance(e):
        attendance_button.config(bg="#FFC107")
    attendance_button.bind("<Enter>", on_enter_attendance)
    attendance_button.bind("<Leave>", on_leave_attendance)

    # Nút Xuất danh sách
    export_button = Button(
        button_frame,
        text="Xuất danh sách",
        font=("Arial", 14),
        bg="#F44336",
        fg="white",
        activebackground="#DA190B",
        padx=20,
        pady=10,
        command=export_absent_list
    )
    export_button.pack(pady=10, fill="x")
    def on_enter_export(e):
        export_button.config(bg="#EF5350")
    def on_leave_export(e):
        export_button.config(bg="#F44336")
    export_button.bind("<Enter>", on_enter_export)
    export_button.bind("<Leave>", on_leave_export)

    # Nút Thoát
    exit_button = Button(
        button_frame,
        text="Thoát",
        font=("Arial", 14),
        bg="#757575",
        fg="white",
        activebackground="#616161",
        padx=20,
        pady=10,
        command=root.destroy
    )
    exit_button.pack(pady=10, fill="x")
    def on_enter_exit(e):
        exit_button.config(bg="#8D8D8D")
    def on_leave_exit(e):
        exit_button.config(bg="#757575")
    exit_button.bind("<Enter>", on_enter_exit)
    exit_button.bind("<Leave>", on_leave_exit)

    #trạng thái
    status_label = Label(root, text=f"Đã đăng nhập: {email}" if email else "Sẵn sàng", font=("Arial", 12), bg="#f5f5f5", wraplength=450)
    status_label.pack(pady=20)

    root.mainloop()

# Chạy chương trình
if __name__ == "__main__":
    if sys.stdout is not None:
        sys.stdout.reconfigure(encoding='utf-8')
    show_login_window()