import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
from PIL import Image
from docx import Document
import pandas as pd
from moviepy.editor import VideoFileClip
from PyPDF2 import PdfReader
from fpdf import FPDF
import zipfile
import chardet

window = tk.Tk()
window.title("格式轉換神器 FormatX Pro")
window.geometry("600x500")

selected_file = None
conversion_type = tk.StringVar()
conversion_type.set("選擇轉換類型")

# ------------------------------
def select_file():
    global selected_file
    selected_file = filedialog.askopenfilename()
    file_label.config(text=f"已選：{os.path.basename(selected_file)}")

def convert():
    if not selected_file:
        messagebox.showerror("錯誤", "請先選擇檔案")
        return

    try:
        ext = os.path.splitext(selected_file)[1].lower()
        output_path = ""

        # Word ➜ PDF
        if conversion_type.get() == "Word ➜ PDF" and ext == ".docx":
            doc = Document(selected_file)
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            for para in doc.paragraphs:
                pdf.multi_cell(0, 10, para.text)
            output_path = selected_file.replace(".docx", ".pdf")
            pdf.output(output_path)

        # PDF ➜ Text
        elif conversion_type.get() == "PDF ➜ Text" and ext == ".pdf":
            reader = PdfReader(selected_file)
            text = ""
            for page in reader.pages:
                text += page.extract_text()
            output_path = selected_file.replace(".pdf", ".txt")
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text)

        # Excel ➜ CSV
        elif conversion_type.get() == "Excel ➜ CSV" and ext in [".xlsx", ".xls"]:
            df = pd.read_excel(selected_file)
            output_path = selected_file.rsplit(".", 1)[0] + ".csv"
            df.to_csv(output_path, index=False)

        # CSV ➜ JSON
        elif conversion_type.get() == "CSV ➜ JSON" and ext == ".csv":
            df = pd.read_csv(selected_file)
            output_path = selected_file.replace(".csv", ".json")
            df.to_json(output_path, orient='records', force_ascii=False)

        # JPG ➜ PNG
        elif conversion_type.get() == "JPG ➜ PNG" and ext == ".jpg":
            img = Image.open(selected_file)
            output_path = selected_file.replace(".jpg", ".png")
            img.save(output_path)

        # PNG ➜ BMP
        elif conversion_type.get() == "PNG ➜ BMP" and ext == ".png":
            img = Image.open(selected_file)
            output_path = selected_file.replace(".png", ".bmp")
            img.save(output_path)

        # MP4 ➜ MP3
        elif conversion_type.get() == "MP4 ➜ MP3" and ext == ".mp4":
            video = VideoFileClip(selected_file)
            output_path = selected_file.replace(".mp4", ".mp3")
            video.audio.write_audiofile(output_path)

        # MOV ➜ MP4
        elif conversion_type.get() == "MOV ➜ MP4" and ext == ".mov":
            clip = VideoFileClip(selected_file)
            output_path = selected_file.replace(".mov", ".mp4")
            clip.write_videofile(output_path)

        # ZIP ➜ 解壓
        elif conversion_type.get() == "ZIP ➜ 解壓縮" and ext == ".zip":
            output_path = selected_file.replace(".zip", "_unzipped/")
            with zipfile.ZipFile(selected_file, 'r') as zip_ref:
                zip_ref.extractall(output_path)

        # Big5 ➜ UTF-8
        elif conversion_type.get() == "Big5 ➜ UTF-8" and ext == ".txt":
            with open(selected_file, 'rb') as f:
                raw = f.read()
                encoding = chardet.detect(raw)['encoding']
            if encoding.lower() != 'big5':
                raise Exception("檔案不是 Big5 編碼")
            content = raw.decode('big5')
            output_path = selected_file.replace(".txt", "_utf8.txt")
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)

        else:
            messagebox.showwarning("格式錯誤", f"此轉換格式目前不支援或副檔名不符合")
            return

        messagebox.showinfo("完成", f"轉換成功：\n{output_path}")

    except Exception as e:
        messagebox.showerror("錯誤", str(e))

# ------------------------------
tk.Button(window, text="選擇檔案", command=select_file).pack(pady=10)
file_label = tk.Label(window, text="尚未選擇檔案")
file_label.pack()

options = [
    "Word ➜ PDF", "PDF ➜ Text", "Excel ➜ CSV", "CSV ➜ JSON",
    "JPG ➜ PNG", "PNG ➜ BMP", "MP4 ➜ MP3", "MOV ➜ MP4",
    "ZIP ➜ 解壓縮", "Big5 ➜ UTF-8"
]

tk.OptionMenu(window, conversion_type, *options).pack(pady=10)
tk.Button(window, text="開始轉換", command=convert, bg="lightblue").pack(pady=20)
tk.Label(window, text="📌 支援常見辦公、影像、音訊、壓縮、編碼轉換").pack(pady=10)

window.mainloop()