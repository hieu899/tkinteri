import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import pandas as pd
from datetime import datetime

# Tạo giao diện chính
root = tk.Tk()
root.title("Quản lý nhân viên")
root.geometry("800x600")

# Các trường thông tin
fields = ["Mã", "Tên", "Ngày sinh (DD/MM/YYYY)", "Giới tính", "Đơn vị", "Chức danh", "Số CMND", "Nơi cấp"]

# Danh sách nhân viên
employee_data = []

# Hàm lưu dữ liệu vào CSV
def save_to_csv():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if file_path:
        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(fields)
            writer.writerows(employee_data)
        messagebox.showinfo("Thông báo", "Dữ liệu đã được lưu thành công!")

# Hàm thêm nhân viên
def add_employee():
    employee = [entry[field].get() for field in fields]
    if not all(employee):
        messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
        return
    try:
        datetime.strptime(employee[2], "%d/%m/%Y")  # Kiểm tra định dạng ngày
    except ValueError:
        messagebox.showerror("Lỗi", "Ngày sinh không hợp lệ!")
        return

    employee_data.append(employee)
    update_table()
    for field in fields:
        entry[field].delete(0, tk.END)

# Hàm cập nhật bảng
def update_table():
    for row in tree.get_children():
        tree.delete(row)
    for emp in employee_data:
        tree.insert("", tk.END, values=emp)

# Hàm tìm nhân viên có sinh nhật hôm nay
def find_birthdays_today():
    today = datetime.today().strftime("%d/%m")
    birthday_list = [emp for emp in employee_data if emp[2][:5] == today]
    if birthday_list:
        messagebox.showinfo("Danh sách sinh nhật hôm nay", "\n".join([emp[1] for emp in birthday_list]))
    else:
        messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")

# Hàm xuất danh sách theo tuổi
def export_sorted_by_age():
    sorted_data = sorted(employee_data, key=lambda x: datetime.strptime(x[2], "%d/%m/%Y"), reverse=True)
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = pd.DataFrame(sorted_data, columns=fields)
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Thông báo", "Danh sách đã được xuất thành công!")

# Giao diện nhập liệu
frame_form = tk.Frame(root)
frame_form.pack(pady=10)

entry = {}
for field in fields:
    tk.Label(frame_form, text=field).grid(row=fields.index(field), column=0, padx=5, pady=5, sticky="w")
    entry[field] = tk.Entry(frame_form, width=30)
    entry[field].grid(row=fields.index(field), column=1, padx=5, pady=5)

# Các nút chức năng
frame_buttons = tk.Frame(root)
frame_buttons.pack(pady=10)

tk.Button(frame_buttons, text="Thêm nhân viên", command=add_employee).grid(row=0, column=0, padx=5, pady=5)
tk.Button(frame_buttons, text="Sinh nhật hôm nay", command=find_birthdays_today).grid(row=0, column=1, padx=5, pady=5)
tk.Button(frame_buttons, text="Lưu CSV", command=save_to_csv).grid(row=0, column=2, padx=5, pady=5)
tk.Button(frame_buttons, text="Xuất danh sách theo tuổi", command=export_sorted_by_age).grid(row=0, column=3, padx=5, pady=5)

# Bảng hiển thị danh sách nhân viên
frame_table = tk.Frame(root)
frame_table.pack(pady=10)

columns = fields
tree = ttk.Treeview(frame_table, columns=columns, show="headings", height=10)
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)

tree.pack()

root.mainloop()
