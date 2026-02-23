import customtkinter as ctk
from tkinter import filedialog, messagebox
import pyttsx3
from docx import Document
import threading
import os
import comtypes.client

# Cấu hình giao diện chung
ctk.set_appearance_mode("System")  # Chế độ sáng/tối theo Windows
ctk.set_default_color_theme("blue")

class TextToMp3App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AI Voice Converter - Word to MP3")
        self.geometry("700x620")  # tăng chiều cao một chút

        # Biến lưu trữ
        self.file_path = ""
        self.voices = []
        self.current_voice_id = None

        # --- UI ELEMENTS ---
        self.label_title = ctk.CTkLabel(self, text="CHUYỂN VĂN BẢN THÀNH ÂM THANH", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=20)

        # 1. Chọn File Word
        self.frame_file = ctk.CTkFrame(self)
        self.frame_file.pack(pady=10, padx=20, fill="x")
        
        self.btn_browse = ctk.CTkButton(self.frame_file, text="Chọn File Word (.docx)", command=self.browse_file)
        self.btn_browse.pack(side="left", padx=10, pady=10)
        
        self.lbl_filename = ctk.CTkLabel(self.frame_file, text="Chưa chọn file...", text_color="gray")
        self.lbl_filename.pack(side="left", padx=10)

        # 2. Hoặc Nhập văn bản trực tiếp
        self.label_or = ctk.CTkLabel(self, text="-- HOẶC NHẬP VĂN BẢN TRỰC TIẾP VÀO ĐÂY --")
        self.label_or.pack(pady=5)
        
        self.textbox = ctk.CTkTextbox(self, height=150)
        self.textbox.pack(pady=10, padx=20, fill="both", expand=True)

        # 3. Cài đặt giọng đọc (ngôn ngữ + giọng + speed + volume)
        self.frame_settings = ctk.CTkFrame(self)
        self.frame_settings.pack(pady=10, padx=20, fill="x")

        # Ngôn ngữ
        self.lbl_lang = ctk.CTkLabel(self.frame_settings, text="Ngôn ngữ:")
        self.lbl_lang.grid(row=0, column=0, padx=20, pady=8, sticky="w")
        
        self.lang_combobox = ctk.CTkComboBox(self.frame_settings, width=220, command=self.on_language_change)
        self.lang_combobox.grid(row=0, column=1, padx=10, pady=8)

        # Giọng đọc
        self.lbl_voice = ctk.CTkLabel(self.frame_settings, text="Giọng đọc:")
        self.lbl_voice.grid(row=1, column=0, padx=20, pady=8, sticky="w")
        
        self.voice_combobox = ctk.CTkComboBox(self.frame_settings, width=300)
        self.voice_combobox.grid(row=1, column=1, padx=10, pady=8)

        # Tốc độ
        self.lbl_speed = ctk.CTkLabel(self.frame_settings, text="Tốc độ đọc:")
        self.lbl_speed.grid(row=2, column=0, padx=20, pady=8, sticky="w")
        self.slider_speed = ctk.CTkSlider(self.frame_settings, from_=80, to=300, number_of_steps=22)
        self.slider_speed.set(170)
        self.slider_speed.grid(row=2, column=1, padx=10, pady=8, sticky="ew")

        # Âm lượng
        self.lbl_volume = ctk.CTkLabel(self.frame_settings, text="Âm lượng:")
        self.lbl_volume.grid(row=3, column=0, padx=20, pady=8, sticky="w")
        self.slider_volume = ctk.CTkSlider(self.frame_settings, from_=0, to=100, number_of_steps=20)
        self.slider_volume.set(100)
        self.slider_volume.grid(row=3, column=1, padx=10, pady=8, sticky="ew")

        # 4. Nút Chuyển đổi
        self.btn_convert = ctk.CTkButton(self, text="BẮT ĐẦU CHUYỂN ĐỔI (MP3)", 
                                         fg_color="green", hover_color="darkgreen",
                                         command=self.start_conversion_thread)
        self.btn_convert.pack(pady=20)

        self.lbl_status = ctk.CTkLabel(self, text="Sẵn sàng", text_color="gray")
        self.lbl_status.pack(pady=5)

        # Tải danh sách giọng ngay khi mở app
        self.load_voices()
    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if filename:
            self.file_path = filename
            self.lbl_filename.configure(text=os.path.basename(filename), text_color="white")
            self.textbox.delete("1.0", "end") # Xóa text nếu đã chọn file
    def load_voices(self):
        try:
            engine = pyttsx3.init()
            self.voices = engine.getProperty('voices')
            engine.stop()

            # Lấy danh sách ngôn ngữ duy nhất
            languages = sorted(set(v.languages[0] if v.languages else "Unknown" for v in self.voices))
            self.lang_combobox.configure(values=languages)
            if languages:
                self.lang_combobox.set(languages[0])  # mặc định ngôn ngữ đầu tiên
                self.on_language_change(languages[0])
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không lấy được danh sách giọng: {e}")

    def on_language_change(self, lang):
        # Lọc giọng theo ngôn ngữ
        filtered_voices = [v for v in self.voices if (v.languages and v.languages[0] == lang)]
        voice_display = [f"{v.name} ({v.gender if hasattr(v, 'gender') else ''})" for v in filtered_voices]
        
        self.voice_combobox.configure(values=voice_display)
        if voice_display:
            self.voice_combobox.set(voice_display[0])
            # lưu id của giọng đang chọn
            self.current_voice_id = filtered_voices[0].id
    def convert_process(self):
        try:
            # Lấy nội dung
            content = ""
            if self.file_path:
                doc = Document(self.file_path)
                content = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            else:
                content = self.textbox.get("1.0", "end-1c")

            if not content.strip():
                messagebox.showwarning("Lỗi", "Vui lòng chọn file hoặc nhập văn bản!")
                return

            # Chọn nơi lưu file MP3
            save_path = filedialog.asksaveasfilename(defaultextension=".mp3", filetypes=[("MP3 files", "*.mp3")])
            if not save_path:
                return

            self.lbl_status.configure(text="Đang xử lý... Vui lòng đợi (không tắt app)", text_color="yellow")
            self.btn_convert.configure(state="disabled")

            # Khởi tạo engine
            engine = pyttsx3.init()
            engine.setProperty('rate', int(self.slider_speed.get()))
            engine.setProperty('volume', self.slider_volume.get() / 100)

            # Áp dụng giọng đang chọn
            if self.current_voice_id:
                engine.setProperty('voice', self.current_voice_id)
            else:
                # fallback lấy giọng đầu tiên
                voices = engine.getProperty('voices')
                engine.setProperty('voice', voices[0].id)

            # Lưu file
            engine.save_to_file(content, save_path)
            engine.runAndWait()

            self.lbl_status.configure(text="Hoàn thành!", text_color="green")
            messagebox.showinfo("Thành công", f"File đã được lưu tại:\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("Lỗi hệ thống", str(e))
        finally:
            self.btn_convert.configure(state="normal")

    def start_conversion_thread(self):
        # Chạy trong thread riêng để UI không bị "treo" khi xử lý file 70.000 từ
        thread = threading.Thread(target=self.convert_process)
        thread.start()

if __name__ == "__main__":
    app = TextToMp3App()
    app.mainloop()