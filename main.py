import customtkinter as ctk
from tkinter import filedialog, messagebox
from docx import Document
import edge_tts
import asyncio
import threading
import os
import pygame
import tempfile
import shutil
import subprocess  # Thêm thư viện để gọi FFmpeg

# Cấu hình giao diện
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class TextToMp3App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AI Voice Converter - Turbo Speed (FFmpeg Edition)")
        self.geometry("750x650")

        self.file_path = ""
        pygame.mixer.init()
        
        # --- UI ELEMENTS ---
        self.label_title = ctk.CTkLabel(self, text="CHUYỂN VĂN BẢN THÀNH SÁCH NÓI (TURBO)", font=ctk.CTkFont(size=20, weight="bold"))
        self.label_title.pack(pady=20)

        # 1. Chọn File Word
        self.frame_file = ctk.CTkFrame(self)
        self.frame_file.pack(pady=5, padx=20, fill="x")
        
        self.btn_browse = ctk.CTkButton(self.frame_file, text="Chọn File Word (.docx)", command=self.browse_file)
        self.btn_browse.pack(side="left", padx=10, pady=10)
        
        self.lbl_filename = ctk.CTkLabel(self.frame_file, text="Chưa chọn file...", text_color="gray")
        self.lbl_filename.pack(side="left", padx=10)

        # 2. Nhập văn bản trực tiếp
        self.textbox = ctk.CTkTextbox(self, height=120)
        self.textbox.pack(pady=10, padx=20, fill="both", expand=True)

        # 3. Bảng điều khiển
        self.frame_controls = ctk.CTkFrame(self)
        self.frame_controls.pack(pady=10, padx=20, fill="x")

        self.lbl_voice = ctk.CTkLabel(self.frame_controls, text="Chọn giọng:")
        self.lbl_voice.grid(row=0, column=0, padx=5, pady=10, sticky="w")
        
        self.voice_var = ctk.StringVar(value="Giọng Nữ (Hoài My)")
        self.combo_voice = ctk.CTkComboBox(self.frame_controls, values=["Giọng Nữ (Hoài My)", "Giọng Nam (Nam Minh)"], variable=self.voice_var, width=150)
        self.combo_voice.grid(row=0, column=1, padx=5, pady=10)

        self.lbl_speed = ctk.CTkLabel(self.frame_controls, text="Tốc độ:")
        self.lbl_speed.grid(row=0, column=2, padx=5, pady=10, sticky="w")
        
        self.slider_speed = ctk.CTkSlider(self.frame_controls, from_=-50, to=50, number_of_steps=100, width=120)
        self.slider_speed.set(0) 
        self.slider_speed.grid(row=0, column=3, padx=5, pady=10)

        self.btn_preview = ctk.CTkButton(self.frame_controls, text="🔊 NGHE THỬ", fg_color="#E67E22", hover_color="#D35400", width=100, command=self.start_preview_thread)
        self.btn_preview.grid(row=0, column=4, padx=15, pady=10)

        # 4. Thanh tiến trình
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.pack(pady=10, padx=20, fill="x")
        self.progress_bar.set(0)

        self.lbl_status = ctk.CTkLabel(self, text="Sẵn sàng", text_color="gray")
        self.lbl_status.pack(pady=5)

        # 5. Nút Chuyển đổi chính
        self.btn_convert = ctk.CTkButton(self, text="⚡ BẮT ĐẦU CHUYỂN ĐỔI SIÊU TỐC", 
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

    # --- HÀM XỬ LÝ NGHE THỬ ---
    async def async_preview_process(self, preview_text, voice_id, rate_str):
        try:
            self.after(0, self.update_ui_progress, 0, "Đang tải bản nghe thử...")
            temp_file = "temp_preview.mp3"
            pygame.mixer.music.unload() if hasattr(pygame.mixer.music, 'unload') else pygame.mixer.music.stop()

            communicate = edge_tts.Communicate(preview_text, voice_id, rate=rate_str)
            await communicate.save(temp_file)

            pygame.mixer.music.load(temp_file)
            pygame.mixer.music.play()
            self.after(0, self.update_ui_progress, 1.0, "Đang phát bản nghe thử...")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi nghe thử: {str(e)}")

    def start_preview_thread(self):
        selected_voice = self.voice_var.get()
        voice_id = "vi-VN-HoaiMyNeural" if "Nữ" in selected_voice else "vi-VN-NamMinhNeural"
        speed_val = int(self.slider_speed.get())
        rate_str = f"+{speed_val}%" if speed_val >= 0 else f"{speed_val}%"

        preview_text = self.textbox.get("1.0", "end-1c").strip()
        if not preview_text:
            preview_text = "Chào bạn, hệ thống xử lý song song FFmpeg đã sẵn sàng."
        else:
            preview_text = preview_text[:100] + "..."

        threading.Thread(target=lambda: asyncio.run(self.async_preview_process(preview_text, voice_id, rate_str))).start()

    # --- Tải 1 khối độc lập ---
    async def download_single_chunk(self, sem, chunk_text, voice_id, rate_str, chunk_index, temp_dir):
        async with sem:
            temp_path = os.path.join(temp_dir, f"chunk_{chunk_index:04d}.mp3")
            communicate = edge_tts.Communicate(chunk_text, voice_id, rate=rate_str)
            await communicate.save(temp_path)
            return temp_path

    # --- HÀM MỚI: Quản lý xử lý song song & Ghép nối siêu tốc ---
    async def async_tts_parallel_process(self, chunks, voice_id, rate_str, save_path):
        temp_dir = tempfile.mkdtemp()
        sem = asyncio.Semaphore(5)    # 5 Luồng tải cùng lúc
        total_chunks = len(chunks)
        completed = 0

        # Tải song song các file âm thanh
        tasks = [
            self.download_single_chunk(sem, chunk, voice_id, rate_str, i, temp_dir)
            for i, chunk in enumerate(chunks)
        ]

        for future in asyncio.as_completed(tasks):
            await future
            completed += 1
            progress = completed / total_chunks
            percent = int(progress * 100)
            self.after(0, self.update_ui_progress, progress, f"⚡ Đang tải song song (Luồng {completed}/{total_chunks}) ... {percent}%")

        # === Giai đoạn nối file bằng FFmpeg ===
        self.after(0, self.update_ui_progress, 1.0, "Đang ghép nối siêu tốc (FFmpeg)...")
        
        # 1. Tìm đường dẫn tuyệt đối của ffmpeg.exe (Sửa lỗi WinError 2)
        current_dir = os.path.dirname(os.path.abspath(__file__))
        ffmpeg_path = os.path.join(current_dir, "ffmpeg.exe")

        if not os.path.exists(ffmpeg_path):
            messagebox.showerror("Thiếu file", "Không tìm thấy ffmpeg.exe!\nVui lòng copy ffmpeg.exe vào cùng thư mục với file code.")
            shutil.rmtree(temp_dir, ignore_errors=True)
            self.after(0, self.update_ui_progress, 0, "Lỗi thiếu FFmpeg!")
            return

        # 2. Tạo file cấu hình list.txt cho FFmpeg
        list_file_path = os.path.join(temp_dir, "list.txt")
        with open(list_file_path, "w", encoding="utf-8") as f:
            for i in range(total_chunks):
                chunk_name = f"chunk_{i:04d}.mp3"
                if os.path.exists(os.path.join(temp_dir, chunk_name)):
                    f.write(f"file '{chunk_name}'\n")

        # 3. Lệnh FFmpeg để dán file trực tiếp không cần giải mã
        command = [
            ffmpeg_path,
            "-y",                 # Ghi đè nếu có
            "-f", "concat",       # Dùng bộ nối
            "-safe", "0",         # Bỏ qua kiểm tra bảo mật file text
            "-i", list_file_path, # File đầu vào
            "-c", "copy",         # Chế độ copy siêu tốc
            save_path             # File đầu ra
        ]

        try:
            # Chạy ẩn (không bung cửa sổ cmd đen)
            subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)
            
            # Dọn dẹp thư mục tạm
            shutil.rmtree(temp_dir, ignore_errors=True)

            self.after(0, self.update_ui_progress, 1.0, "Hoàn thành xuất sắc!")
            messagebox.showinfo("Thành công", f"Sách nói siêu tốc đã được xuất tại:\n{save_path}")
            
            # --- Tính năng mới: Tự động mở thư mục chứa file MP3 ---
            os.startfile(os.path.dirname(save_path))

        except subprocess.CalledProcessError as e:
            messagebox.showerror("Lỗi FFmpeg", f"Có lỗi khi ghép file: {e}")
            self.after(0, self.update_ui_progress, 0, "Lỗi ghép file!")
            shutil.rmtree(temp_dir, ignore_errors=True)

    def convert_process(self):
        try:
            self.after(0, self.update_ui_progress, 0, "Đang chuẩn bị dữ liệu...")
            self.btn_convert.configure(state="disabled")

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

            selected_voice = self.voice_var.get()
            voice_id = "vi-VN-HoaiMyNeural" if "Nữ" in selected_voice else "vi-VN-NamMinhNeural"
            speed_val = int(self.slider_speed.get())
            rate_str = f"+{speed_val}%" if speed_val >= 0 else f"{speed_val}%"

            chunks = []
            current_chunk = ""
            for line in raw_text_lines:
                if len(current_chunk) + len(line) < 4000:
                    current_chunk += line + "\n"
                else:
                    chunks.append(current_chunk.strip())
                    current_chunk = line + "\n"
            if current_chunk:
                chunks.append(current_chunk.strip())

            pygame.mixer.music.stop()
            
            asyncio.run(self.async_tts_parallel_process(chunks, voice_id, rate_str, save_path))
            
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