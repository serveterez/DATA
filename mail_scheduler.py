import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkcalendar import Calendar
from tkhtmlview import HTMLLabel
import pandas as pd
import json

CONFIG_FILE = "config.json"


def select_file(entry, filetypes):
    path = filedialog.askopenfilename(filetypes=filetypes)
    if path:
        entry.delete(0, ctk.END)
        entry.insert(0, path)


def fill_message_from_files():
    excel_path = entry_excel.get()
    template_path = entry_html.get()
    if not excel_path or not template_path:
        messagebox.showwarning("Uyarı", "Lütfen önce Excel ve şablon dosyasını seçin.")
        return
    try:
        df = pd.read_excel(excel_path)
        with open(template_path, "r", encoding="utf-8") as f:
            template = f.read()
        row = df.iloc[0]
        full_name = str(row["İsim"]).strip()
        first_name = full_name.split(" ")[0]
        title = str(row["Cinsiyet"]).strip()
        custom_name = f"{first_name} {title}"
        mail_body = template.replace("{isim}", custom_name)
        message_html.delete("1.0", ctk.END)
        message_html.insert(ctk.END, mail_body)
        update_html_preview()
    except Exception as e:
        messagebox.showerror("Hata", f"Otomatik mesaj doldurulamadı:\n{e}")


def add_signature():
    sign_path = entry_signature.get()
    if not sign_path:
        messagebox.showwarning("Uyarı", "Lütfen önce imza dosyası seçin.")
        return
    try:
        with open(sign_path, "r", encoding="utf-8") as f:
            signature = f.read()
        message_html.insert(ctk.END, "\n" + signature)
        update_html_preview()
    except Exception as e:
        messagebox.showerror("Hata", f"İmza eklenemedi:\n{e}")


def update_html_preview(event=None):
    html = message_html.get("1.0", ctk.END).strip()
    html_preview.set_html(html)


def send_mail_from_box():
    import win32com.client as win32
    subject = entry_subject.get()
    body = message_html.get("1.0", ctk.END).strip()
    if not body or not subject:
        messagebox.showerror("Hata", "Konu ve mesaj alanı boş olamaz.")
        return
    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = ""  # İstediğin test adresini yaz
        mail.Subject = subject
        mail.HTMLBody = body
        mail.Display()
        messagebox.showinfo("Başarılı", "Mail penceresi açıldı. Kontrol edip gönderebilirsiniz.")
    except Exception as e:
        messagebox.showerror("Hata", f"Mail gönderilemedi:\n{e}")


def add_time():
    hour = hour_entry.get()
    minute = minute_entry.get()
    if not hour.isdigit() or not minute.isdigit() or not (0 <= int(hour) <= 23) or not (0 <= int(minute) <= 59):
        messagebox.showerror("Hata", "Saat ve dakika 0-23 ve 0-59 aralığında olmalı.")
        return
    time_str = f"{int(hour):02d}:{int(minute):02d}"
    if time_str not in list_times.get(0, tk.END):
        list_times.insert(tk.END, time_str)


def del_time():
    selected = list(list_times.curselection())
    for index in reversed(selected):
        list_times.delete(index)


def add_schedules():
    # Çoklu gün ve saat seçimi ile kombinasyonları ekle
    days = sorted(list(selected_days_set))
    times = list(list_times.get(0, tk.END))
    if not days or not times:
        messagebox.showerror("Hata", "En az bir gün ve bir saat seçmelisiniz.")
        return
    for day in days:
        for t in times:
            dt_str = f"{day} {t}"
            if dt_str not in list_schedules.get(0, tk.END):
                list_schedules.insert(tk.END, dt_str)


def del_schedules():
    selected = list(list_schedules.curselection())
    for index in reversed(selected):
        list_schedules.delete(index)


def save_settings():
    excel = entry_excel.get()
    html = entry_html.get()
    sign = entry_signature.get()
    subject = entry_subject.get()
    batch = entry_batch.get()
    schedule_times = list(list_schedules.get(0, tk.END))
    mail_message = message_html.get("1.0", ctk.END).strip()
    if not excel or not html or not subject or not batch or not schedule_times or not mail_message:
        messagebox.showerror("Hata", "Tüm alanları doldurun!")
        return
    try:
        batch_int = int(batch)
        if batch_int < 1:
            raise ValueError
    except Exception:
        messagebox.showerror("Hata", "Batch boyutu sayı olmalı!")
        return
    config = {
        "excel_path": excel,
        "template_path": html,
        "signature_path": sign,
        "subject": subject,
        "batch_size": batch_int,
        "schedule_datetimes": schedule_times,
        "mail_message": mail_message,
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    messagebox.showinfo("Başarılı", "Ayarlar kaydedildi.")
    root.destroy()


# --- Çoklu Gün Seçimi için yardımcı fonksiyonlar

def on_cal_click(event):
    day = calendar.get_date()
    # Date string olarak gelsin
    if day in selected_days_set:
        selected_days_set.remove(day)
        calendar.calevent_remove('sel', pd.to_datetime(day))
    else:
        selected_days_set.add(day)
        calendar.calevent_create(pd.to_datetime(day), 'Seçili', 'sel')
    calendar.tag_config('sel', background='orange', foreground='white')


# --- Arayüz ---

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Takvimli Mail Scheduler")
root.geometry("800x800")
root.minsize(800, 800)

tabControl = ttk.Notebook(root)
tab1 = ctk.CTkFrame(tabControl)
tab2 = ctk.CTkFrame(tabControl)
tabControl.add(tab1, text="Dosya & İçerik")
tabControl.add(tab2, text="Zamanlama")
tabControl.pack(expand=True, fill="both")

# --- TAB1: Dosya & İçerik ---
ctk.CTkLabel(tab1, text="Excel Dosyası:").pack(anchor="w", padx=10)
entry_excel = ctk.CTkEntry(tab1, width=280)
entry_excel.pack(anchor="w", padx=20, pady=2)
ctk.CTkButton(tab1, text="Seç", width=80, command=lambda: select_file(entry_excel, [("Excel Dosyası", "*.xlsx *.xls")])).pack(anchor="w", padx=20, pady=2)

ctk.CTkLabel(tab1, text="HTML Şablon Dosyası:").pack(anchor="w", padx=10)
entry_html = ctk.CTkEntry(tab1, width=280)
entry_html.pack(anchor="w", padx=20, pady=2)
ctk.CTkButton(tab1, text="Seç", width=80, command=lambda: select_file(entry_html, [("HTML Dosyası", "*.html")])).pack(anchor="w", padx=20, pady=2)

ctk.CTkLabel(tab1, text="İmza HTML Dosyası (Opsiyonel):").pack(anchor="w", padx=10)
entry_signature = ctk.CTkEntry(tab1, width=280)
entry_signature.pack(anchor="w", padx=20, pady=2)
ctk.CTkButton(tab1, text="Seç", width=80, command=lambda: select_file(entry_signature, [("HTML Dosyası", "*.html")])).pack(anchor="w", padx=20, pady=2)

ctk.CTkLabel(tab1, text="Konu:").pack(anchor="w", padx=10)
entry_subject = ctk.CTkEntry(tab1, width=350)
entry_subject.pack(anchor="w", padx=20, pady=2)

ctk.CTkLabel(tab1, text="Batch (Bir Seferde Kaç Mail Gönderilsin?):").pack(anchor="w", padx=10)
entry_batch = ctk.CTkEntry(tab1, width=80)
entry_batch.pack(anchor="w", padx=20, pady=2)

ctk.CTkLabel(tab1, text="Mail Mesajı (HTML kod veya elle yaz):").pack(anchor="w", padx=10, pady=6)
msg_outer_frame = tk.Frame(tab1)
msg_outer_frame.pack(anchor="w", padx=10, pady=2, fill="x", expand=False)
message_html = tk.Text(msg_outer_frame, wrap="word", width=38, height=10, font=("Arial", 11))
message_html.pack(side="left", padx=8, pady=2)
message_html.bind("<KeyRelease>", update_html_preview)
preview_frame = tk.Frame(msg_outer_frame)
preview_frame.pack(side="left", padx=10, pady=2)
ctk.CTkLabel(preview_frame, text="Canlı Önizleme:").pack()
html_preview = HTMLLabel(preview_frame, width=30, height=11, background="white")
html_preview.pack(pady=8)
html_preview.set_html("")

btn_frame = ctk.CTkFrame(tab1, fg_color="transparent")
btn_frame.pack(anchor="w", padx=20, pady=8)
ctk.CTkButton(btn_frame, text="Dosyalardan Doldur", width=110, command=fill_message_from_files).pack(side="left", padx=4)
ctk.CTkButton(btn_frame, text="İmza Ekle", width=80, command=add_signature).pack(side="left", padx=4)
ctk.CTkButton(btn_frame, text="Maili Gönder", width=110, command=send_mail_from_box, fg_color="orange").pack(side="left", padx=4)

# --- TAB2: Zamanlama (Çoklu Gün & Saat) ---
ctk.CTkLabel(tab2, text="Çoklu Gün Seç (Takvimden tıklayarak seç/çıkart):").pack(anchor="w", padx=10, pady=8)
calendar = Calendar(tab2, selectmode="day", date_pattern='yyyy-mm-dd')
calendar.pack(anchor="w", padx=20, pady=4)
selected_days_set = set()
calendar.bind("<<CalendarSelected>>", on_cal_click)

ctk.CTkLabel(tab2, text="Saat/Dakika Ekle:").pack(anchor="w", padx=10, pady=8)
time_frame = ctk.CTkFrame(tab2)
time_frame.pack(anchor="w", padx=20, pady=2)
hour_entry = ctk.CTkEntry(time_frame, width=40)
hour_entry.insert(0, "10")
hour_entry.pack(side="left", padx=3)
ctk.CTkLabel(time_frame, text=":").pack(side="left")
minute_entry = ctk.CTkEntry(time_frame, width=40)
minute_entry.insert(0, "00")
minute_entry.pack(side="left", padx=3)
ctk.CTkButton(time_frame, text="Saat Ekle", width=80, command=add_time).pack(side="left", padx=6)
ctk.CTkButton(time_frame, text="Saat Sil", width=80, command=del_time).pack(side="left", padx=3)

list_times = tk.Listbox(tab2, width=8, height=5, font=("Arial", 11))
list_times.pack(anchor="w", padx=20, pady=4)

ctk.CTkButton(tab2, text="Seçili Günler ve Saatleri Listeye Ekle", width=200, command=add_schedules, fg_color="blue").pack(anchor="w", padx=20, pady=3)
ctk.CTkButton(tab2, text="Listedekileri Sil", width=120, command=del_schedules, fg_color="red").pack(anchor="w", padx=20, pady=3)

list_schedules = tk.Listbox(tab2, width=28, height=7, font=("Arial", 11))
list_schedules.pack(anchor="w", padx=20, pady=4)

ctk.CTkButton(root, text="Ayarları Kaydet ve Çık", width=220, command=save_settings, fg_color="green").pack(pady=10)

root.mainloop()
