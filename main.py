# =====================================================================
# Phiên bản GỘP: Giữ ưu điểm hiển thị tốc độ % + xử lý lỗi an toàn hơn
# =====================================================================

import customtkinter as ctk
from tkinter import filedialog, messagebox ,BooleanVar, Radiobutton
from docx import Document
import edge_tts
import asyncio
import threading
import os
import pygame
import tempfile
import shutil
import subprocess

# Cấu hình giao diện
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class TextToMp3App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AI Voice Converter - Turbo Speed (FFmpeg Edition)")
        self.geometry("800x650")

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

        # 2. Textbox hiển thị và chỉnh sửa nội dung
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

        self.lbl_speed_title = ctk.CTkLabel(self.frame_controls, text="Tốc độ:")
        self.lbl_speed_title.grid(row=0, column=2, padx=(15, 5), pady=10, sticky="w")
        
        self.slider_speed = ctk.CTkSlider(self.frame_controls, from_=-50, to=50, number_of_steps=100, width=120, command=self.update_speed_label)
        self.slider_speed.set(0) 
        self.slider_speed.grid(row=0, column=3, padx=5, pady=10)

        self.lbl_speed_val = ctk.CTkLabel(self.frame_controls, text="0%", font=ctk.CTkFont(weight="bold"), text_color="#3498DB")
        self.lbl_speed_val.grid(row=0, column=4, padx=5, pady=10)

        self.btn_preview = ctk.CTkButton(self.frame_controls, text="🔊 NGHE THỬ", fg_color="#E67E22", hover_color="#D35400", width=100, command=self.start_preview_thread)
        self.btn_preview.grid(row=0, column=5, padx=15, pady=10)

        # 4. Thanh tiến trình
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.pack(pady=10, padx=20, fill="x")
        self.progress_bar.set(0)

        self.lbl_status = ctk.CTkLabel(self, text="Sẵn sàng", text_color="gray")
        self.lbl_status.pack(pady=5)

        # 5. Nút chuyển đổi chính
        self.btn_convert = ctk.CTkButton(self, text="⚡ BẮT ĐẦU CHUYỂN ĐỔI SIÊU TỐC", 
                                         fg_color="green", hover_color="darkgreen",
                                         command=self.start_conversion_thread)
        self.btn_convert.pack(pady=15)
        # 6. Chọn ảnh nền (cho video)
        self.frame_image = ctk.CTkFrame(self)
        self.frame_image.pack(pady=5, padx=20, fill="x")
        
        self.btn_browse_image = ctk.CTkButton(self.frame_image, text="Chọn ảnh nền (cho video)", command=self.browse_image)
        self.btn_browse_image.pack(side="left", padx=10, pady=10)
        
        self.lbl_image_name = ctk.CTkLabel(self.frame_image, text="Chưa chọn ảnh...", text_color="gray")
        self.lbl_image_name.pack(side="left", padx=10)
        
        self.image_path = ""  # biến lưu đường dẫn ảnh

        # 7. Chọn định dạng đầu ra
        self.frame_format = ctk.CTkFrame(self)
        self.frame_format.pack(pady=5, padx=20, fill="x")
        
        self.output_var = ctk.StringVar(value="mp3")
        
        ctk.CTkLabel(self.frame_format, text="Định dạng đầu ra:").pack(side="left", padx=10)
        
        Radiobutton(self.frame_format, text="MP3 (chỉ âm thanh)", 
                    variable=self.output_var, value="mp3",
                    fg="#ffffff", selectcolor="#2b2b2b").pack(side="left", padx=20)
        
        Radiobutton(self.frame_format, text="MP4 (video + 1 ảnh tĩnh)", 
                    variable=self.output_var, value="mp4",
                    fg="#ffffff", selectcolor="#2b2b2b").pack(side="left", padx=20)
        
    def update_speed_label(self, value):
        val = int(value)
        sign = "+" if val > 0 else ""
        self.lbl_speed_val.configure(text=f"{sign}{val}%")

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if filename:
            self.file_path = filename
            self.lbl_filename.configure(text=os.path.basename(filename), text_color="white")
            try:
                self.lbl_status.configure(text="Đang tải dữ liệu từ file Word...", text_color="yellow")
                self.update_idletasks()
                self.textbox.delete("1.0", "end")
                doc = Document(filename)
                full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                self.textbox.insert("1.0", full_text)
                self.lbl_status.configure(text="Đã tải xong! Bạn có thể chỉnh sửa trước khi chuyển đổi.", text_color="green")
            except Exception as e:
                messagebox.showerror("Lỗi đọc file", str(e))
                self.lbl_status.configure(text="Lỗi khi đọc file", text_color="red")
    def browse_image(self):
        filename = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
        if filename:
            self.image_path = filename
            self.lbl_image_name.configure(text=os.path.basename(filename), text_color="white")
    def update_ui_progress(self, value, text):
        self.progress_bar.set(value)
        self.lbl_status.configure(text=text)

    async def async_preview_process(self, preview_text, voice_id, rate_str):
        try:
            self.after(0, self.update_ui_progress, 0, "Đang tạo bản nghe thử...")
            temp_file = "temp_preview.mp3"
            pygame.mixer.music.unload() if hasattr(pygame.mixer.music, 'unload') else pygame.mixer.music.stop()
            communicate = edge_tts.Communicate(preview_text, voice_id, rate=rate_str)
            await communicate.save(temp_file)
            pygame.mixer.music.load(temp_file)
            pygame.mixer.music.play()
            self.after(0, self.update_ui_progress, 1.0, "Đang phát bản nghe thử...")
        except Exception as e:
            messagebox.showerror("Lỗi nghe thử", str(e))

    def start_preview_thread(self):
        selected_voice = self.voice_var.get()
        voice_id = "vi-VN-HoaiMyNeural" if "Nữ" in selected_voice else "vi-VN-NamMinhNeural"
        speed_val = int(self.slider_speed.get())
        rate_str = f"+{speed_val}%" if speed_val >= 0 else f"{speed_val}%"
        preview_text = self.textbox.get("1.0", "end-1c").strip()
        preview_text = (preview_text[:150] + "...") if preview_text else "Hệ thống đã sẵn sàng."
        threading.Thread(target=lambda: asyncio.run(self.async_preview_process(preview_text, voice_id, rate_str))).start()

    async def download_single_chunk(self, sem, chunk_text, voice_id, rate_str, chunk_index, temp_dir):
        async with sem:
            temp_path = os.path.join(temp_dir, f"chunk_{chunk_index:04d}.mp3")
            communicate = edge_tts.Communicate(chunk_text, voice_id, rate=rate_str)
            await communicate.save(temp_path)
            return temp_path

    async def async_tts_parallel_process(self, chunks, voice_id, rate_str, save_path):
        temp_dir = tempfile.mkdtemp()
        sem = asyncio.Semaphore(5)  # Giới hạn 5 luồng đồng thời
        total_chunks = len(chunks)
        completed = 0

        tasks = [self.download_single_chunk(sem, chunk, voice_id, rate_str, i, temp_dir) for i, chunk in enumerate(chunks)]

        for future in asyncio.as_completed(tasks):
            await future
            completed += 1
            progress = completed / total_chunks
            self.after(0, self.update_ui_progress, progress, f"Đang tải song song ({completed}/{total_chunks}) ... {int(progress*100)}%")

        # Ghép file bằng FFmpeg
        current_dir = os.path.dirname(os.path.abspath(__file__))
        ffmpeg_path = os.path.join(current_dir, "ffmpeg.exe")

        if not os.path.exists(ffmpeg_path):
            raise FileNotFoundError("Không tìm thấy ffmpeg.exe trong thư mục chương trình!")

        list_file_path = os.path.join(temp_dir, "list.txt")
        with open(list_file_path, "w", encoding="utf-8") as f:
            for i in range(total_chunks):
                f.write(f"file 'chunk_{i:04d}.mp3'\n")

        command = [ffmpeg_path, "-y", "-f", "concat", "-safe", "0", "-i", list_file_path, "-c", "copy", save_path]

        # Không hiển thị messagebox ở đây nữa, chuyển ra hàm convert_process
        subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)

    def convert_process(self):
        self.btn_convert.configure(state="disabled")
        try:
            text = self.textbox.get("1.0", "end-1c").strip()
            if not text:
                messagebox.showwarning("Cảnh báo", "Không có nội dung để chuyển đổi!")
                return

            output_format = self.output_var.get()
            if output_format == "mp4" and not self.image_path:
                messagebox.showwarning("Cảnh báo", "Để xuất MP4 bạn phải chọn một ảnh nền!")
                return

            defaultext = ".mp4" if output_format == "mp4" else ".mp3"
            filetypes = [("MP4 files", "*.mp4")] if output_format == "mp4" else [("MP3 files", "*.mp3")]
            
            save_path = filedialog.asksaveasfilename(defaultextension=defaultext, filetypes=filetypes)
            if not save_path:
                return

            selected_voice = self.voice_var.get()
            voice_id = "vi-VN-HoaiMyNeural" if "Nữ" in selected_voice else "vi-VN-NamMinhNeural"
            speed_val = int(self.slider_speed.get())
            rate_str = f"+{speed_val}%" if speed_val >= 0 else f"{speed_val}%"

            # Chia chunk
            raw_text_lines = text.split('\n')
            chunks = []
            current_chunk = ""
            for line in raw_text_lines:
                if len(current_chunk) + len(line) < 4000:
                    current_chunk += line + "\n"
                else:
                    if current_chunk.strip():
                        chunks.append(current_chunk.strip())
                    current_chunk = line + "\n"
            if current_chunk.strip():
                chunks.append(current_chunk.strip())

            if not chunks:
                messagebox.showwarning("Cảnh báo", "Không có nội dung hợp lệ!")
                return

            pygame.mixer.music.stop()

            # Tạo thư mục tạm
            temp_dir = tempfile.mkdtemp()
            mp3_final = os.path.join(temp_dir, "final_audio.mp3")

            # Tạo audio trước
            asyncio.run(self.async_tts_parallel_process(chunks, voice_id, rate_str, mp3_final))

            self.after(0, self.update_ui_progress, 0.92, "Đang xử lý file cuối...")

            current_dir = os.path.dirname(os.path.abspath(__file__))
            ffmpeg_path = os.path.join(current_dir, "ffmpeg.exe")
            if not os.path.exists(ffmpeg_path):
                raise FileNotFoundError("Không tìm thấy ffmpeg.exe!")

            if output_format == "mp3":
                shutil.copy(mp3_final, save_path)
                self.after(0, self.update_ui_progress, 1.0, "Hoàn thành MP3!")
            else:
                # Debug: Kiểm tra file tồn tại trước khi chạy FFmpeg
                if not os.path.exists(self.image_path):
                    raise FileNotFoundError(f"Ảnh không tồn tại: {self.image_path}")
                if not os.path.exists(mp3_final):
                    raise FileNotFoundError(f"Audio tạm không tồn tại: {mp3_final}")

                command = [
                    ffmpeg_path, "-y",
                    "-loop", "1",
                    "-i", self.image_path,
                    "-i", mp3_final,
                    "-c:v", "libx264",
                    "-tune", "stillimage",
                    "-pix_fmt", "yuv420p",
                    "-c:a", "copy",  # copy audio để nhanh, tránh encode lại
                    "-shortest",
                    "-movflags", "+faststart",
                    save_path
                ]

                try:
                    # Chạy và capture output để debug
                    result = subprocess.run(command, check=True, capture_output=True, text=True)
                    print("FFmpeg output:", result.stdout)  # in ra console để bạn xem
                    print("FFmpeg error (nếu có):", result.stderr)
                except subprocess.CalledProcessError as ffmpeg_err:
                    error_msg = f"Lỗi FFmpeg khi tạo video:\n{ffmpeg_err.stderr}\nCommand: {' '.join(command)}"
                    raise RuntimeError(error_msg)

                self.after(0, self.update_ui_progress, 1.0, "Hoàn thành MP4!")

            shutil.rmtree(temp_dir, ignore_errors=True)
            messagebox.showinfo("Thành công", f"File đã lưu tại:\n{save_path}")
            os.startfile(os.path.dirname(save_path))

        except Exception as e:
            messagebox.showerror("Lỗi chuyển đổi", str(e))
            self.after(0, self.update_ui_progress, 0, "Xảy ra lỗi!")
        finally:
            self.btn_convert.configure(state="normal")
    def start_conversion_thread(self):
        threading.Thread(target=self.convert_process, daemon=True).start()

if __name__ == "__main__":
    app = TextToMp3App()
    app.mainloop()