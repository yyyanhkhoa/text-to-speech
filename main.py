import customtkinter as ctk
from tkinter import filedialog, messagebox
from docx import Document
import edge_tts
import asyncio
import threading
import os

# Cấu hình giao diện
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class TextToMp3App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AI Voice Converter - Edge TTS (Bản Pro)")
        self.geometry("700x650")

        self.file_path = ""
        
        # --- UI ELEMENTS ---
        self.label_title = ctk.CTkLabel(self, text="CHUYỂN VĂN BẢN THÀNH SÁCH NÓI", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=20)

        # 1. Chọn File Word
        self.frame_file = ctk.CTkFrame(self)
        self.frame_file.pack(pady=5, padx=20, fill="x")
        
        self.btn_browse = ctk.CTkButton(self.frame_file, text="Chọn File Word (.docx)", command=self.browse_file)
        self.btn_browse.pack(side="left", padx=10, pady=10)
        
        self.lbl_filename = ctk.CTkLabel(self.frame_file, text="Chưa chọn file...", text_color="gray")
        self.lbl_filename.pack(side="left", padx=10)

        # 2. Hoặc Nhập văn bản trực tiếp
        self.textbox = ctk.CTkTextbox(self, height=120)
        self.textbox.pack(pady=10, padx=20, fill="both", expand=True)

        # 3. Bảng điều khiển (Giọng đọc & Tốc độ)
        self.frame_controls = ctk.CTkFrame(self)
        self.frame_controls.pack(pady=10, padx=20, fill="x")

        # Chọn giọng
        self.lbl_voice = ctk.CTkLabel(self.frame_controls, text="Chọn giọng đọc:")
        self.lbl_voice.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        self.voice_var = ctk.StringVar(value="Giọng Nữ (Hoài My)")
        self.combo_voice = ctk.CTkComboBox(self.frame_controls, values=["Giọng Nữ (Hoài My)", "Giọng Nam (Nam Minh)"], variable=self.voice_var)
        self.combo_voice.grid(row=0, column=1, padx=10, pady=10)

        # Chỉnh tốc độ
        self.lbl_speed = ctk.CTkLabel(self.frame_controls, text="Tốc độ:")
        self.lbl_speed.grid(row=0, column=2, padx=10, pady=10, sticky="w")
        
        # Slider từ -50% đến +50%
        self.slider_speed = ctk.CTkSlider(self.frame_controls, from_=-50, to=50, number_of_steps=100)
        self.slider_speed.set(0) # Mặc định là 0 (Tốc độ chuẩn)
        self.slider_speed.grid(row=0, column=3, padx=10, pady=10)

        # 4. Thanh tiến trình (Progress Bar)
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.pack(pady=10, padx=20, fill="x")
        self.progress_bar.set(0)

        self.lbl_status = ctk.CTkLabel(self, text="Sẵn sàng", text_color="gray")
        self.lbl_status.pack(pady=5)

        # 5. Nút Chuyển đổi
        self.btn_convert = ctk.CTkButton(self, text="BẮT ĐẦU CHUYỂN ĐỔI (MP3)", 
                                         fg_color="green", hover_color="darkgreen",
                                         command=self.start_conversion_thread)
        self.btn_convert.pack(pady=15)

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if filename:
            self.file_path = filename
            self.lbl_filename.configure(text=os.path.basename(filename), text_color="white")
            self.textbox.delete("1.0", "end")

    def update_ui_progress(self, value, text):
        self.progress_bar.set(value)
        self.lbl_status.configure(text=text)

    # Hàm xử lý bất đồng bộ (Async) cho edge-tts
    async def async_tts_process(self, chunks, voice_id, rate_str, save_path):
        total_chunks = len(chunks)
        with open(save_path, 'wb') as audio_file:
            for index, chunk in enumerate(chunks):
                progress = index / total_chunks
                percent = int(progress * 100)
                self.after(0, self.update_ui_progress, progress, f"Đang tổng hợp giọng nói... {percent}%")
                
                # Gọi edge-tts và lấy luồng dữ liệu (stream)
                communicate = edge_tts.Communicate(chunk, voice_id, rate=rate_str)
                async for audio_chunk in communicate.stream():
                    if audio_chunk["type"] == "audio":
                        audio_file.write(audio_chunk["data"])

        self.after(0, self.update_ui_progress, 1.0, "Hoàn thành!")
        messagebox.showinfo("Thành công", f"Sách nói đã được xuất tại:\n{save_path}")

    def convert_process(self):
        try:
            self.after(0, self.update_ui_progress, 0, "Đang chuẩn bị dữ liệu...")
            self.btn_convert.configure(state="disabled")

            # 1. Trích xuất văn bản
            raw_text_lines = []
            if self.file_path:
                doc = Document(self.file_path)
                raw_text_lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            else:
                text = self.textbox.get("1.0", "end-1c").strip()
                raw_text_lines = text.split('\n')

            if not raw_text_lines or all(not line for line in raw_text_lines):
                messagebox.showwarning("Cảnh báo", "Không có nội dung để chuyển đổi!")
                self.btn_convert.configure(state="normal")
                return

            save_path = filedialog.asksaveasfilename(defaultextension=".mp3", filetypes=[("MP3 files", "*.mp3")])
            if not save_path:
                self.btn_convert.configure(state="normal")
                return

            # 2. Xử lý thiết lập giọng và tốc độ
            selected_voice = self.voice_var.get()
            voice_id = "vi-VN-HoaiMyNeural" if "Nữ" in selected_voice else "vi-VN-NamMinhNeural"
            
            speed_val = int(self.slider_speed.get())
            rate_str = f"+{speed_val}%" if speed_val >= 0 else f"{speed_val}%"

            # 3. Cắt nhỏ văn bản (Chunking) ~ 2000 ký tự mỗi cụm
            chunks = []
            current_chunk = ""
            for line in raw_text_lines:
                if len(current_chunk) + len(line) < 2000:
                    current_chunk += line + "\n"
                else:
                    chunks.append(current_chunk.strip())
                    current_chunk = line + "\n"
            if current_chunk:
                chunks.append(current_chunk.strip())

            # 4. Chạy vòng lặp sự kiện Asyncio trong Thread phụ
            asyncio.run(self.async_tts_process(chunks, voice_id, rate_str, save_path))
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")
            self.after(0, self.update_ui_progress, 0, "Tiến trình bị lỗi!")
        finally:
            self.btn_convert.configure(state="normal")

    def start_conversion_thread(self):
        thread = threading.Thread(target=self.convert_process)
        thread.start()

if __name__ == "__main__":
    app = TextToMp3App()
    app.mainloop()