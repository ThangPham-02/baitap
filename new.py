import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import csv
import os
import pandas as pd

# File lưu trữ dữ liệu nhân viên
DATA_FILE = "employee_data.csv"

# Khởi tạo dữ liệu nếu chưa có file
if not os.path.exists(DATA_FILE):
    with open(DATA_FILE, mode="w", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Ngày cấp", "Nơi cấp"])

def save_data():
    """Lưu thông tin nhân viên vào file CSV"""
    employee_id = entry_id.get().strip()
    name = entry_name.get().strip()
    unit = combo_unit.get().strip()
    position = entry_position.get().strip()
    birth_date = entry_birth_date.get().strip()
    gender = gender_var.get()
    id_number = entry_id_number.get().strip()
    issue_date = entry_issue_date.get().strip()
    issue_place = entry_issue_place.get().strip()

    if not all([employee_id, name, unit, birth_date, gender]):
        messagebox.showwarning("Cảnh báo", "Vui lòng điền đầy đủ các trường bắt buộc.")
        return

    try:
        # Kiểm tra ngày hợp lệ
        datetime.strptime(birth_date, "%d/%m/%Y")
        if issue_date:
            datetime.strptime(issue_date, "%d/%m/%Y")
    except ValueError:
        messagebox.showerror("Lỗi", "Ngày tháng không hợp lệ. Vui lòng nhập theo định dạng DD/MM/YYYY.")
        return

    # Lưu dữ liệu vào file CSV
    with open(DATA_FILE, mode="a", newline="") as file:
        writer = csv.writer(file)
        writer.writerow([employee_id, name, unit, position, birth_date, gender, id_number, issue_date, issue_place])

    messagebox.showinfo("Thành công", "Lưu thông tin thành công.")
    clear_form()

def clear_form():
    """Xóa toàn bộ thông tin trên form"""
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    combo_unit.set("")
    entry_position.delete(0, tk.END)
    entry_birth_date.delete(0, tk.END)
    gender_var.set("Nam")
    entry_id_number.delete(0, tk.END)
    entry_issue_date.delete(0, tk.END)
    entry_issue_place.delete(0, tk.END)

def show_today_birthdays():
    """Hiển thị danh sách nhân viên có sinh nhật hôm nay"""
    today = datetime.now().strftime("%d/%m")
    birthdays = []

    with open(DATA_FILE, mode="r") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if today in row["Ngày sinh"]:
                birthdays.append(row)

    if birthdays:
        msg = "\n".join([f"{b['Tên']} ({b['Ngày sinh']})" for b in birthdays])
        messagebox.showinfo("Sinh nhật hôm nay", f"Những nhân viên có sinh nhật hôm nay:\n{msg}")
    else:
        messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")

def export_to_excel():
    """Xuất toàn bộ danh sách nhân viên ra file Excel, sắp xếp theo tuổi giảm dần"""
    try:
        df = pd.read_csv(DATA_FILE)
        df["Ngày sinh"] = pd.to_datetime(df["Ngày sinh"], format="%d/%m/%Y")
        df = df.sort_values(by="Ngày sinh")
        output_file = "employee_data.xlsx"
        df.to_excel(output_file, index=False)
        messagebox.showinfo("Thành công", f"Xuất danh sách ra file {output_file} thành công.")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể xuất file Excel: {e}")

# Tạo giao diện chính
root = tk.Tk()
root.title("Thông tin nhân viên")

# Khung thông tin nhân viên
frame = ttk.LabelFrame(root, text="Thông tin nhân viên")
frame.pack(padx=10, pady=10, fill="x")

# Các trường nhập liệu
labels = ["Mã *", "Tên *", "Đơn vị *", "Chức danh", "Ngày sinh * (DD/MM/YYYY)", "Giới tính *", "Số CMND", "Ngày cấp (DD/MM/YYYY)", "Nơi cấp"]
entries = []

entry_id = ttk.Entry(frame)
entry_name = ttk.Entry(frame)
combo_unit = ttk.Combobox(frame, values=["Phân xưởng cơ khí", "Phân xưởng que hàn", "Phòng kỹ thuật"])
entry_position = ttk.Entry(frame)
entry_birth_date = ttk.Entry(frame)
entry_id_number = ttk.Entry(frame)
entry_issue_date = ttk.Entry(frame)
entry_issue_place = ttk.Entry(frame)

# Giới tính radio button
gender_var = tk.StringVar(value="Nam")
gender_frame = ttk.Frame(frame)
radio_male = ttk.Radiobutton(gender_frame, text="Nam", variable=gender_var, value="Nam")
radio_female = ttk.Radiobutton(gender_frame, text="Nữ", variable=gender_var, value="Nữ")
radio_male.pack(side="left")
radio_female.pack(side="left")

entries = [entry_id, entry_name, combo_unit, entry_position, entry_birth_date, gender_frame, entry_id_number, entry_issue_date, entry_issue_place]

# Hiển thị các trường
for i, (label, entry) in enumerate(zip(labels, entries)):
    ttk.Label(frame, text=label).grid(row=i, column=0, sticky="w", padx=5, pady=5)
    entry.grid(row=i, column=1, sticky="ew", padx=5, pady=5)

# Nút chức năng
btn_frame = ttk.Frame(root)
btn_frame.pack(pady=10)

btn_save = ttk.Button(btn_frame, text="Lưu thông tin", command=save_data)
btn_save.pack(side="left", padx=5)

btn_birthday = ttk.Button(btn_frame, text="Sinh nhật hôm nay", command=show_today_birthdays)
btn_birthday.pack(side="left", padx=5)

btn_export = ttk.Button(btn_frame, text="Xuất toàn bộ danh sách", command=export_to_excel)
btn_export.pack(side="left", padx=5)

# Chạy chương trình
root.mainloop()
