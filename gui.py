import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import pdf_ops
import os
import threading

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class PDFEditorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Doc-Tor")
        self.geometry("850x600")

        # Layout configuration
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create Navigation Frame (Sidebar)
        self.sidebar_frame = ctk.CTkFrame(self, width=160, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Doc-Tor", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Navigation Buttons
        self.home_btn = self.create_nav_btn("Home / Info", self.show_home, 1)
        self.word_btn = self.create_nav_btn("Convert Word", self.show_word, 2)
        self.edit_btn = self.create_nav_btn("Edit / Redact", self.show_edit, 3)
        self.pages_btn = self.create_nav_btn("Page Tools", self.show_pages, 4)
        self.security_btn = self.create_nav_btn("Security", self.show_security, 5)
        self.merge_btn = self.create_nav_btn("Merge Files", self.show_merge, 6)
        self.extract_btn = self.create_nav_btn("Extract Text", self.show_extract, 7)

        # Main Content Frames
        self.home_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.word_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.edit_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.pages_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.security_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.merge_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.extract_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")

        self.setup_home_frame()
        self.setup_word_frame()
        self.setup_edit_frame()
        self.setup_pages_frame()
        self.setup_security_frame()
        self.setup_merge_frame()
        self.setup_extract_frame()

        # Default View
        self.select_frame_by_name("Home / Info")

    def create_nav_btn(self, text, command, row):
        btn = ctk.CTkButton(self.sidebar_frame, text=text, command=command, fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"), anchor="w")
        btn.grid(row=row, column=0, padx=20, pady=10, sticky="ew")
        return btn

    def select_frame_by_name(self, name):
        # Reset buttons
        for btn in [self.home_btn, self.word_btn, self.edit_btn, self.pages_btn, self.security_btn, self.merge_btn, self.extract_btn]:
            btn.configure(fg_color="transparent")

        # Hide all frames
        for frame in [self.home_frame, self.word_frame, self.edit_frame, self.pages_frame, self.security_frame, self.merge_frame, self.extract_frame]:
            frame.grid_forget()

        # Show selected
        frames = {
            "Home / Info": (self.home_frame, self.home_btn),
            "Convert Word": (self.word_frame, self.word_btn),
            "Edit / Redact": (self.edit_frame, self.edit_btn),
            "Page Tools": (self.pages_frame, self.pages_btn),
            "Security": (self.security_frame, self.security_btn),
            "Merge Files": (self.merge_frame, self.merge_btn),
            "Extract Text": (self.extract_frame, self.extract_btn)
        }
        
        frame, btn = frames[name]
        frame.grid(row=0, column=1, sticky="nsew")
        btn.configure(fg_color=("gray75", "gray25"))

    def show_home(self): self.select_frame_by_name("Home / Info")
    def show_word(self): self.select_frame_by_name("Convert Word")
    def show_edit(self): self.select_frame_by_name("Edit / Redact")
    def show_pages(self): self.select_frame_by_name("Page Tools")
    def show_security(self): self.select_frame_by_name("Security")
    def show_merge(self): self.select_frame_by_name("Merge Files")
    def show_extract(self): self.select_frame_by_name("Extract Text")

    # --- HOME FRAME ---
    def setup_home_frame(self):
        self.home_label = ctk.CTkLabel(self.home_frame, text="PDF Metadata Viewer", font=ctk.CTkFont(size=18, weight="bold"))
        self.home_label.pack(pady=20)
        ctk.CTkButton(self.home_frame, text="Select PDF File", command=self.load_pdf_info).pack(pady=10)
        self.info_textbox = ctk.CTkTextbox(self.home_frame, width=400, height=250)
        self.info_textbox.pack(pady=10)
        self.info_textbox.insert("0.0", "No file selected.")
        self.info_textbox.configure(state="disabled")

    def load_pdf_info(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename:
            try:
                info = pdf_ops.get_pdf_info(filename)
                display_text = f"File: {os.path.basename(filename)}\n\n"
                for k, v in info.items():
                    display_text += f"{k.replace('_', ' ').title()}: {v}\n"
                self.info_textbox.configure(state="normal")
                self.info_textbox.delete("0.0", "end")
                self.info_textbox.insert("0.0", display_text)
                self.info_textbox.configure(state="disabled")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    # --- WORD CONVERSION FRAME ---
    def setup_word_frame(self):
        ctk.CTkLabel(self.word_frame, text="Word <-> PDF Conversion", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        ctk.CTkButton(self.word_frame, text="Convert PDF to Word", command=self.run_word_conversion).pack(pady=10)
        ctk.CTkButton(self.word_frame, text="Convert Word back to PDF", command=self.run_pdf_conversion, fg_color="green", hover_color="darkgreen").pack(pady=10)
        self.word_status = ctk.CTkLabel(self.word_frame, text="Select a tool above.")
        self.word_status.pack(pady=20)
        self.word_progress = ctk.CTkProgressBar(self.word_frame, mode="indeterminate", width=300)

    def run_word_conversion(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if not filename: return
        self.run_threaded_task(lambda: pdf_ops.convert_pdf_to_word(filename), "Converting to Word...", self.word_status, self.word_progress, open_file=True)

    def run_pdf_conversion(self):
        filename = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if not filename: return
        self.run_threaded_task(lambda: pdf_ops.convert_word_to_pdf(filename), "Converting to PDF...", self.word_status, self.word_progress)

    def run_threaded_task(self, func, status_msg, status_label, progress_bar, open_file=False):
        def task():
            try:
                progress_bar.pack(pady=10)
                progress_bar.start()
                status_label.configure(text=status_msg, text_color="orange")
                path = func()
                status_label.configure(text=f"Success: {os.path.basename(path)}", text_color="green")
                if open_file: os.startfile(path)
                else: messagebox.showinfo("Success", f"Saved to:\n{os.path.basename(path)}")
            except Exception as e:
                status_label.configure(text=f"Error: {str(e)}", text_color="red")
            finally:
                progress_bar.stop()
                progress_bar.pack_forget()
        threading.Thread(target=task).start()

    # --- PAGE TOOLS FRAME ---
    def setup_pages_frame(self):
        ctk.CTkLabel(self.pages_frame, text="Page Manipulation", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        
        # File Selection
        self.pages_file_btn = ctk.CTkButton(self.pages_frame, text="Select PDF", command=self.select_pages_file)
        self.pages_file_btn.pack(pady=5)
        self.pages_file_label = ctk.CTkLabel(self.pages_frame, text="No file selected")
        self.pages_file_label.pack(pady=5)

        # Rotate
        ctk.CTkLabel(self.pages_frame, text="Rotate Pages:").pack(pady=(15, 5))
        self.rotate_var = ctk.StringVar(value="90")
        ctk.CTkComboBox(self.pages_frame, values=["90", "180", "270"], variable=self.rotate_var).pack(pady=5)
        ctk.CTkButton(self.pages_frame, text="Rotate", command=self.run_rotate).pack(pady=5)

        # Delete
        ctk.CTkLabel(self.pages_frame, text="Delete Pages (e.g. 1, 3-5):").pack(pady=(15, 5))
        self.delete_entry = ctk.CTkEntry(self.pages_frame, placeholder_text="1, 3, 5")
        self.delete_entry.pack(pady=5)
        ctk.CTkButton(self.pages_frame, text="Delete Pages", command=self.run_delete, fg_color="red", hover_color="darkred").pack(pady=5)

        # Split
        ctk.CTkLabel(self.pages_frame, text="Extract Range (Split):").pack(pady=(15, 5))
        split_frame = ctk.CTkFrame(self.pages_frame, fg_color="transparent")
        split_frame.pack(pady=5)
        self.split_start = ctk.CTkEntry(split_frame, width=60, placeholder_text="Start")
        self.split_start.pack(side="left", padx=5)
        self.split_end = ctk.CTkEntry(split_frame, width=60, placeholder_text="End")
        self.split_end.pack(side="left", padx=5)
        ctk.CTkButton(self.pages_frame, text="Extract Range", command=self.run_split).pack(pady=5)

    def select_pages_file(self):
        self.selected_pages_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_pages_file:
            self.pages_file_label.configure(text=os.path.basename(self.selected_pages_file))

    def run_rotate(self):
        if not hasattr(self, 'selected_pages_file') or not self.selected_pages_file: return
        try:
            path = pdf_ops.rotate_pages(self.selected_pages_file, int(self.rotate_var.get()))
            messagebox.showinfo("Success", f"Rotated file saved to:\n{os.path.basename(path)}")
        except Exception as e: messagebox.showerror("Error", str(e))

    def run_delete(self):
        if not hasattr(self, 'selected_pages_file') or not self.selected_pages_file: return
        pages_str = self.delete_entry.get()
        if not pages_str: return
        try:
            # Parse "1, 3, 5" -> [0, 2, 4]
            page_list = [int(p.strip()) - 1 for p in pages_str.split(",")]
            path = pdf_ops.delete_pages(self.selected_pages_file, page_list)
            messagebox.showinfo("Success", f"Pages deleted. Saved to:\n{os.path.basename(path)}")
        except Exception as e: messagebox.showerror("Error", str(e))

    def run_split(self):
        if not hasattr(self, 'selected_pages_file') or not self.selected_pages_file: return
        try:
            start = int(self.split_start.get())
            end = int(self.split_end.get())
            path = pdf_ops.extract_page_range(self.selected_pages_file, start, end)
            messagebox.showinfo("Success", f"Range extracted. Saved to:\n{os.path.basename(path)}")
        except Exception as e: messagebox.showerror("Error", str(e))


    # --- SECURITY FRAME ---
    def setup_security_frame(self):
        ctk.CTkLabel(self.security_frame, text="PDF Security", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        
        self.security_file_btn = ctk.CTkButton(self.security_frame, text="Select PDF", command=self.select_security_file)
        self.security_file_btn.pack(pady=5)
        self.security_file_label = ctk.CTkLabel(self.security_frame, text="No file selected")
        self.security_file_label.pack(pady=5)

        self.password_entry = ctk.CTkEntry(self.security_frame, placeholder_text="Enter Password", show="*")
        self.password_entry.pack(pady=20, fill="x", padx=60)

        ctk.CTkButton(self.security_frame, text="Encrypt (Protect)", command=self.run_encrypt).pack(pady=10)
        ctk.CTkButton(self.security_frame, text="Decrypt (Unlock)", command=self.run_decrypt, fg_color="green", hover_color="darkgreen").pack(pady=10)

    def select_security_file(self):
        self.selected_security_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_security_file:
            self.security_file_label.configure(text=os.path.basename(self.selected_security_file))

    def run_encrypt(self):
        if not hasattr(self, 'selected_security_file') or not self.selected_security_file: return
        pwd = self.password_entry.get()
        if not pwd: 
            messagebox.showwarning("Warning", "Enter a password.")
            return
        try:
            path = pdf_ops.encrypt_pdf(self.selected_security_file, pwd)
            messagebox.showinfo("Success", f"Encrypted file saved to:\n{os.path.basename(path)}")
        except Exception as e: messagebox.showerror("Error", str(e))

    def run_decrypt(self):
        if not hasattr(self, 'selected_security_file') or not self.selected_security_file: return
        pwd = self.password_entry.get()
        if not pwd: 
            messagebox.showwarning("Warning", "Enter a password.")
            return
        try:
            path = pdf_ops.decrypt_pdf(self.selected_security_file, pwd)
            messagebox.showinfo("Success", f"Decrypted file saved to:\n{os.path.basename(path)}")
        except Exception as e: messagebox.showerror("Error", str(e))

    # --- EDIT / REDACT FRAME ---
    def setup_edit_frame(self):
        ctk.CTkLabel(self.edit_frame, text="Edit or Redact Text", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        self.edit_file_btn = ctk.CTkButton(self.edit_frame, text="Select PDF", command=self.select_edit_file)
        self.edit_file_btn.pack(pady=5)
        self.edit_file_label = ctk.CTkLabel(self.edit_frame, text="No file selected")
        self.edit_file_label.pack(pady=5)
        self.search_entry = ctk.CTkEntry(self.edit_frame, placeholder_text="Text to search...")
        self.search_entry.pack(pady=10, fill="x", padx=40)
        self.replace_entry = ctk.CTkEntry(self.edit_frame, placeholder_text="New text (Leave empty for Redaction)")
        self.replace_entry.pack(pady=10, fill="x", padx=40)
        ctk.CTkButton(self.edit_frame, text="Execute", command=self.run_edit_action).pack(pady=20)

    def select_edit_file(self):
        self.selected_edit_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_edit_file:
            self.edit_file_label.configure(text=os.path.basename(self.selected_edit_file))

    def run_edit_action(self):
        if not hasattr(self, 'selected_edit_file') or not self.selected_edit_file:
            messagebox.showwarning("Warning", "Please select a file first.")
            return
        search_text = self.search_entry.get()
        replace_text = self.replace_entry.get()
        if not search_text:
            messagebox.showwarning("Warning", "Please enter text to search.")
            return
        try:
            if replace_text:
                count, path = pdf_ops.edit_pdf_text(self.selected_edit_file, search_text, replace_text)
                msg = f"Replaced {count} instances.\nSaved to: {os.path.basename(path)}"
            else:
                count, path = pdf_ops.redact_pdf(self.selected_edit_file, search_text)
                msg = f"Redacted {count} instances.\nSaved to: {os.path.basename(path)}"
            messagebox.showinfo("Success", msg)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # --- MERGE FRAME ---
    def setup_merge_frame(self):
        ctk.CTkLabel(self.merge_frame, text="Merge PDFs", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        self.merge_listbox = ctk.CTkTextbox(self.merge_frame, width=400, height=200)
        self.merge_listbox.pack(pady=10)
        self.merge_files = []
        btn_frame = ctk.CTkFrame(self.merge_frame, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="Add File", command=self.add_merge_file).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Clear", command=self.clear_merge_list, fg_color="red", hover_color="darkred").pack(side="left", padx=10)
        ctk.CTkButton(self.merge_frame, text="Merge Files", command=self.run_merge).pack(pady=20)

    def add_merge_file(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename:
            self.merge_files.append(filename)
            self.merge_listbox.insert("end", f"{os.path.basename(filename)}\n")

    def clear_merge_list(self):
        self.merge_files = []
        self.merge_listbox.delete("0.0", "end")

    def run_merge(self):
        if len(self.merge_files) < 2:
            messagebox.showwarning("Warning", "Add at least 2 files.")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if save_path:
            try:
                pdf_ops.merge_pdfs(self.merge_files, save_path)
                messagebox.showinfo("Success", "Files merged successfully!")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    # --- EXTRACT FRAME ---
    def setup_extract_frame(self):
        ctk.CTkLabel(self.extract_frame, text="Extract Text", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=20)
        self.extract_btn_select = ctk.CTkButton(self.extract_frame, text="Select PDF & Extract", command=self.run_extract)
        self.extract_btn_select.pack(pady=20)
        self.extract_status = ctk.CTkLabel(self.extract_frame, text="")
        self.extract_status.pack(pady=20)

    def run_extract(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename:
            try:
                path = pdf_ops.extract_text_from_pdf(filename)
                self.extract_status.configure(text=f"Extracted to: {os.path.basename(path)}", text_color="green")
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = PDFEditorApp()
    app.mainloop()
