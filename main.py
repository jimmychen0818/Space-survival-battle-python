import tkinter as tk
from tkinter import filedialog, messagebox
import os
from PIL import Image
from docx import Document
import pandas as pd
from moviepy.editor import VideoFileClip
from PyPDF2 import PdfReader
from fpdf import FPDF

window = tk.Tk()
window.title("格式轉換神器 FormatX Pro - 批次版")
window.geometry("600x500")

selected_files = []
conversion_type = tk.StringVar()
conversion_type.set("選擇轉換類型")

# 選檔
def select_files():
    global selected_files
    selected_files = filedialog.askopenfilenames()
    filenames = [os.path.basename(f) for f in selected_files]
    file_label.config(text=f"已選 {len(selected_files)} 個檔案\n" + "\n".join(filenames[:5]) + ("..." if len(filenames) > 5 else ""))

# 批量轉換
def convert_files():
    if not selected_files:
        messagebox.showerror("錯誤", "請先選擇檔案")
        return

    success_count = 0
    failed_count = 0

    for file in selected_files:
        try:
            ext = os.path.splitext(file)[1].lower()
            output_path = ""

            if conversion_type.get() == "Word ➜ PDF" and ext == ".docx":
                doc = Document(file)
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                for para in doc.paragraphs:
                    pdf.multi_cell(0, 10, para.text)
                output_path = file.replace(".docx", ".pdf")
                pdf.output(output_path)

            elif conversion_type.get() == "Excel ➜ CSV" and ext in [".xlsx", ".xls"]:
                df = pd.read_excel(file)
                output_path = file.rsplit(".", 1)[0] + ".csv"
                df.to_csv(output_path, index=False)

            elif conversion_type.get() == "JPG ➜ PNG" and ext == ".jpg":
                img = Image.open(file)
                output_path = file.replace(".jpg", ".png")
                img.save(output_path)

            elif conversion_type.get() == "MP4 ➜ MP3" and ext == ".mp4":
                video = VideoFileClip(file)
                output_path = file.replace(".mp4", ".mp3")
                video.audio.write_audiofile(output_path)

            elif conversion_type.get() == "PDF ➜ TXT" and ext == ".pdf":
                reader = PdfReader(file)
                text = ""
                for page in reader.pages:
                    text += page.extract_text()
                output_path = file.replace(".pdf", ".txt")
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(text)

            else:
                raise Exception("格式不符")

            success_count += 1
        except Exception as e:
            print(f"失敗：{file}：{str(e)}")
            failed_count += 1

    messagebox.showinfo("完成",
        f"成功轉換 {success_count} 筆檔案\n失敗 {failed_count} 筆")

# 介面
tk.Button(window, text="選擇多個檔案", command=select_files).pack(pady=10)
file_label = tk.Label(window, text="尚未選擇檔案")
file_label.pack()

options = [
    "Word ➜ PDF", "Excel ➜ CSV",
    "JPG ➜ PNG", "MP4 ➜ MP3",
    "PDF ➜ TXT"
]
tk.OptionMenu(window, conversion_type, *options).pack(pady=10)

tk.Button(window, text="開始批次轉換", command=convert_files, bg="lightblue").pack(pady=20)
tk.Label(window, text="📁 可選多個檔案，批量轉換").pack(pady=10)

window.mainloop()