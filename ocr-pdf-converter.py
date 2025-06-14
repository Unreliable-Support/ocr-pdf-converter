import tkinter
import tkinter.messagebox
import customtkinter
from tkinter import filedialog
import os
import threading
import subprocess # For Pandoc
import pytesseract
import fitz # PyMuPDF
from PIL import Image
import io
import multiprocessing # For parallel processing
import re # For Regex operations

# Specify the path to Tesseract if it's not in the system's PATH (may be needed for Windows)
try:
    subprocess.run(["tesseract", "--version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    print("Tesseract OCR found in PATH.")
    TESSERACT_AVAILABLE = True
except (subprocess.CalledProcessError, FileNotFoundError):
    print("Tesseract OCR not found in PATH. Attempting to set manually or add to PATH.")
    TESSERACT_AVAILABLE = False
    if os.name == 'nt': # Only for Windows
        tesseract_path_windows = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        if os.path.exists(tesseract_path_windows):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path_windows
            print(f"Tesseract OCR path set to: {tesseract_path_windows}")
            TESSERACT_AVAILABLE = True
        else:
            print(f"Default Tesseract OCR path ({tesseract_path_windows}) not found.")

customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

# OCR function to be executed by worker processes
def ocr_page_worker_function(args_tuple):
    pdf_path, page_num, ocr_lang, dpi, tesseract_cmd_path_from_main = args_tuple

    if tesseract_cmd_path_from_main:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd_path_from_main
    
    try:
        doc = fitz.open(pdf_path) 
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        doc.close() 

        img_bytes = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_bytes))

        custom_config = r'--psm 3' # Page segmentation mode
        page_text = pytesseract.image_to_string(img, lang=ocr_lang, config=custom_config)
        return (page_num, page_text)
    except pytesseract.TesseractNotFoundError:
        print(f"TESSERACT NOT FOUND IN WORKER PROCESS: Page {page_num+1}, File: {os.path.basename(pdf_path)}. Command Path: {getattr(pytesseract.pytesseract, 'tesseract_cmd', 'Not Set')}")
        return (page_num, f"[Tesseract Not Found Error for Page {page_num+1}]")
    except Exception as e:
        print(f"OCR Error (Worker): Page {page_num+1}, File: {os.path.basename(pdf_path)}: {e}")
        return (page_num, f"[OCR Error for Page {page_num+1}: {str(e)}]")

def format_text_with_heuristics(text):
    lines = text.splitlines()
    formatted_lines = []
    
    numbered_heading_pattern = re.compile(r"^\s*(\d+(\.\d+)*\.?)\s+([A-Z][\w\s:,()-]+)$")
    chapter_heading_pattern = re.compile(r"^(?:\d+\s*[-\u2013\u2014]?\s*)?CHAPTER\s*\d*[:\-\s]*([A-Z0-9].*)$", re.IGNORECASE)
    all_caps_short_heading_pattern = re.compile(r"^\s*([A-Z0-9][A-Z0-9\s'-]{5,80})\s*$")
    common_section_keywords = ["introduction", "conclusion", "summary", "abstract", "references", "appendix", "acknowledgements", "contents", "figure", "table"]
    common_section_pattern = re.compile(r"^\s*(" + "|".join(common_section_keywords) + r")[:\.]?\s*$", re.IGNORECASE)

    for i, line in enumerate(lines):
        stripped_line = line.strip()
        original_line_to_append = line

        if not stripped_line:
            formatted_lines.append(line)
            continue

        match_numbered = numbered_heading_pattern.match(stripped_line)
        if match_numbered:
            text_part = match_numbered.group(3).strip()
            if text_part and len(text_part.split()) < 10 and text_part[0].isupper():
                formatted_lines.append(f"## {stripped_line}")
                continue

        match_chapter = chapter_heading_pattern.match(stripped_line)
        if match_chapter:
            text_part = match_chapter.group(1).strip()
            if text_part:
                formatted_lines.append(f"# {text_part.upper()}")
                if i + 1 < len(lines):
                    next_line_stripped = lines[i+1].strip()
                    if next_line_stripped and next_line_stripped[0].islower() and len(next_line_stripped.split()) < 7:
                        pass
                continue

        if stripped_line.isupper() and \
           len(stripped_line.split()) > 0 and \
           len(stripped_line.split()) < 8 and \
           not re.search(r'[.,;:!?]$', stripped_line) and \
           sum(1 for char in stripped_line if char.isalpha()) > len(stripped_line) * 0.6:
            is_likely_standalone_heading = True
            if i > 0 and lines[i-1].strip():
                if not lines[i-1].strip().endswith(('.', '!', '?', ':')):
                    is_likely_standalone_heading = False
            if i + 1 < len(lines) and lines[i+1].strip():
                if not lines[i+1].strip()[0].isupper():
                     is_likely_standalone_heading = False
            if is_likely_standalone_heading:
                formatted_lines.append(f"### {stripped_line}")
                continue
        
        match_common_section = common_section_pattern.match(stripped_line)
        if match_common_section:
            prev_line_empty_or_heading = (i == 0) or \
                                         (not lines[i-1].strip()) or \
                                         (lines[i-1].strip().startswith("#"))
            if prev_line_empty_or_heading:
                formatted_lines.append(f"## {stripped_line.capitalize()}")
                continue

        match_text_then_num = re.match(r"^\s*([A-Za-z][\w\s'-]+?)\s*[-\u2013\u2014]\s*\d+\s*$", stripped_line)
        if match_text_then_num:
            text_part = match_text_then_num.group(1).strip()
            if len(text_part.split()) < 8 and len(text_part) > 5:
                formatted_lines.append(f"### {text_part}")
                continue

        formatted_lines.append(original_line_to_append)

    final_text = "\n".join(formatted_lines)
    final_text = re.sub(r'\n{3,}', '\n\n', final_text)
    return final_text

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Intelligent PDF Converter (OCR + Pandoc)")
        self.geometry(f"{800}x{680}")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.current_font_size = customtkinter.StringVar(value="11pt")
        self.current_margin = customtkinter.StringVar(value="0.7in")
        self.current_main_font = customtkinter.StringVar(value="Liberation Serif")
        self.current_pdf_engine = customtkinter.StringVar(value="xelatex")
        self.current_line_spacing = customtkinter.StringVar(value="1.0")

        self.tabview = customtkinter.CTkTabview(self, width=250)
        self.tabview.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.tabview.add("Files")
        self.tabview.add("Settings")

        # --- Files Tab ---
        self.tabview.tab("Files").grid_columnconfigure(0, weight=1)
        self.file_list_frame = customtkinter.CTkFrame(self.tabview.tab("Files"))
        self.file_list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.file_list_frame.grid_columnconfigure(0, weight=1)
        self.file_list_frame.grid_rowconfigure(0, weight=1)
        self.file_listbox = tkinter.Listbox(self.file_list_frame, selectmode=tkinter.EXTENDED, borderwidth=0, highlightthickness=0)
        self.file_listbox.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        self.listbox_scrollbar = customtkinter.CTkScrollbar(self.file_list_frame, command=self.file_listbox.yview)
        self.listbox_scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=self.listbox_scrollbar.set)
        
        self.button_frame = customtkinter.CTkFrame(self.tabview.tab("Files"))
        self.button_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        self.add_files_button = customtkinter.CTkButton(self.button_frame, text="Add Files", command=self.add_files)
        self.add_files_button.pack(side="left", padx=5, pady=5)
        self.remove_selected_button = customtkinter.CTkButton(self.button_frame, text="Remove Selected", command=self.remove_selected_files)
        self.remove_selected_button.pack(side="left", padx=5, pady=5)
        self.clear_list_button = customtkinter.CTkButton(self.button_frame, text="Clear List", command=self.clear_file_list)
        self.clear_list_button.pack(side="left", padx=5, pady=5)
        
        self.target_dir_label = customtkinter.CTkLabel(self.tabview.tab("Files"), text="Output Folder: Not selected")
        self.target_dir_label.grid(row=2, column=0, padx=10, pady=(10,0), sticky="w")
        self.select_target_button = customtkinter.CTkButton(self.tabview.tab("Files"), text="Select Output Folder", command=self.select_target_directory)
        self.select_target_button.grid(row=3, column=0, padx=10, pady=5, sticky="ew")
        
        self.language_label = customtkinter.CTkLabel(self.tabview.tab("Files"), text="OCR Language:")
        self.language_label.grid(row=4, column=0, padx=10, pady=(10,0), sticky="w")
        self.language_var = customtkinter.StringVar(value="eng+tur")
        self.language_options = ["eng", "tur", "eng+tur", "deu", "fra", "ara", "rus", "spa", "jpn", "chi_sim"]
        self.language_menu = customtkinter.CTkOptionMenu(self.tabview.tab("Files"), variable=self.language_var, values=self.language_options)
        self.language_menu.grid(row=5, column=0, padx=10, pady=5, sticky="ew")
        
        self.convert_button = customtkinter.CTkButton(self.tabview.tab("Files"), text="Convert", command=self.start_conversion_thread, font=customtkinter.CTkFont(size=16, weight="bold"))
        self.convert_button.grid(row=6, column=0, padx=10, pady=20, sticky="ew")
        
        self.status_label = customtkinter.CTkLabel(self.tabview.tab("Files"), text="Status: Idle")
        self.status_label.grid(row=7, column=0, padx=10, pady=5, sticky="ew")
        self.progressbar = customtkinter.CTkProgressBar(self.tabview.tab("Files"))
        self.progressbar.grid(row=8, column=0, padx=10, pady=10, sticky="ew")
        self.progressbar.set(0)

        # --- Settings Tab ---
        self.tabview.tab("Settings").grid_columnconfigure(0, weight=1)
        self.appearance_mode_label = customtkinter.CTkLabel(self.tabview.tab("Settings"), text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=0, column=0, padx=20, pady=(10, 0), sticky="w")
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.tabview.tab("Settings"), values=["Light", "Dark", "System"], command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.appearance_mode_optionemenu.set("System")
        
        self.pandoc_settings_label = customtkinter.CTkLabel(self.tabview.tab("Settings"), text="Pandoc PDF Settings", font=customtkinter.CTkFont(size=14, weight="bold"))
        self.pandoc_settings_label.grid(row=2, column=0, padx=20, pady=(10,5), sticky="w")
        
        self.font_size_label = customtkinter.CTkLabel(self.tabview.tab("Settings"), text="Font Size:")
        self.font_size_label.grid(row=3, column=0, padx=20, pady=(5,0), sticky="w")
        self.font_size_options = ["8pt", "9pt", "10pt", "11pt", "12pt", "13pt", "14pt", "16pt", "18pt", "20pt", "22pt", "24pt"]
        self.font_size_dropdown = customtkinter.CTkOptionMenu(self.tabview.tab("Settings"), variable=self.current_font_size, values=self.font_size_options)
        self.font_size_dropdown.grid(row=4, column=0, padx=20, pady=(0,10), sticky="ew")
        
        self.main_font_label = customtkinter.CTkLabel(self.tabview.tab("Settings"), text="Main Font (for Xe/LuaLaTeX):")
        self.main_font_label.grid(row=5, column=0, padx=20, pady=(5,0), sticky="w")
        self.main_font_options = ["Liberation Serif", "Linux Libertine O", "Times New Roman", "Arial", "Calibri", "DejaVu Serif", "DejaVu Sans", "Georgia", "Verdana", "EB Garamond", "Noto Serif", "Noto Sans", "CMU Serif", "CMU Sans Serif"]
        self.main_font_dropdown = customtkinter.CTkOptionMenu(self.tabview.tab("Settings"), variable=self.current_main_font, values=self.main_font_options)
        self.main_font_dropdown.grid(row=6, column=0, padx=20, pady=(0,10), sticky="ew")
        
        self.margin_label = customtkinter.CTkLabel(self.tabview.tab("Settings"), text="Margin (e.g., 0.7in, 2cm):")
        self.margin_label.grid(row=7, column=0, padx=20, pady=(5,0), sticky="w")
        self.margin_entry = customtkinter.CTkEntry(self.tabview.tab("Settings"), textvariable=self.current_margin)
        self.margin_entry.grid(row=8, column=0, padx=20, pady=(0,10), sticky="ew")
        
        self.pdf_engine_label = customtkinter.CTkLabel(self.tabview.tab("Settings"), text="PDF Engine:")
        self.pdf_engine_label.grid(row=9, column=0, padx=20, pady=(5,0), sticky="w")
        self.pdf_engine_options = ["xelatex", "lualatex", "pdflatex"]
        self.pdf_engine_dropdown = customtkinter.CTkOptionMenu(self.tabview.tab("Settings"), variable=self.current_pdf_engine, values=self.pdf_engine_options)
        self.pdf_engine_dropdown.grid(row=10, column=0, padx=20, pady=(0,10), sticky="ew")
        
        self.line_spacing_label = customtkinter.CTkLabel(self.tabview.tab("Settings"), text="Line Spacing (e.g., 1.0, 1.5):")
        self.line_spacing_label.grid(row=11, column=0, padx=20, pady=(5,0), sticky="w")
        self.line_spacing_entry = customtkinter.CTkEntry(self.tabview.tab("Settings"), textvariable=self.current_line_spacing)
        self.line_spacing_entry.grid(row=12, column=0, padx=20, pady=(0,10), sticky="ew")
        
        self.save_settings_button = customtkinter.CTkButton(self.tabview.tab("Settings"), text="Apply Settings (Informational)", command=self.apply_pandoc_settings)
        self.save_settings_button.grid(row=13, column=0, padx=20, pady=20, sticky="ew")

        self.selected_files = []
        self.target_directory = ""
        
        if not TESSERACT_AVAILABLE:
            if os.name == 'nt' and not os.path.exists(r'C:\Program Files\Tesseract-OCR\tesseract.exe'):
                 self.after(100, lambda: tkinter.messagebox.showwarning("Warning", "Tesseract OCR not found or not in PATH.\nThe default path C:\\Program Files\\Tesseract-OCR\\tesseract.exe is also invalid.\nPlease install it and add to PATH, or specify its path in the code."))
            elif os.name != 'nt':
                 self.after(100, lambda: tkinter.messagebox.showwarning("Warning", "Tesseract OCR not found or not in PATH.\nPlease install it and add to your system's PATH."))

        self.change_appearance_mode_event(customtkinter.get_appearance_mode())

    def apply_pandoc_settings(self):
        font_size = self.current_font_size.get()
        main_font = self.current_main_font.get()
        margin = self.current_margin.get()
        pdf_engine = self.current_pdf_engine.get()
        line_spacing = self.current_line_spacing.get()
        tkinter.messagebox.showinfo("Settings",
                                    f"Settings updated (will be used during conversion):\n"
                                    f"- Font Size: {font_size}\n"
                                    f"- Font: {main_font}\n"
                                    f"- Margin: {margin}\n"
                                    f"- PDF Engine: {pdf_engine}\n"
                                    f"- Line Spacing: {line_spacing}")
        print(f"Pandoc settings updated in UI: Size={font_size}, Font={main_font}, Margin={margin}, Engine={pdf_engine}, Line Spacing={line_spacing}")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)
        effective_mode = customtkinter.get_appearance_mode() 
        if effective_mode == "Dark":
            self.file_listbox.configure(bg="#2B2B2B", fg="white")
        else: # Light
            self.file_listbox.configure(bg="white", fg="black")

    def add_files(self):
        files = filedialog.askopenfilenames(title="Select PDF Files", filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*")))
        if files:
            for file_path in files:
                if file_path not in self.selected_files:
                    self.selected_files.append(file_path)
                    self.file_listbox.insert(tkinter.END, os.path.basename(file_path))
            self.update_status(f"{len(files)} file(s) added.")

    def remove_selected_files(self):
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            tkinter.messagebox.showwarning("Warning", "Select at least one file to remove.")
            return
        for index in reversed(selected_indices):
            del self.selected_files[index]
            self.file_listbox.delete(index)
        self.update_status(f"{len(selected_indices)} file(s) removed.")

    def clear_file_list(self):
        self.selected_files.clear()
        self.file_listbox.delete(0, tkinter.END)
        self.update_status("File list cleared.")

    def select_target_directory(self):
        directory = filedialog.askdirectory(title="Select Output Folder")
        if directory:
            self.target_directory = directory
            self.target_dir_label.configure(text=f"Output Folder: {self.target_directory}")
            self.update_status(f"Output folder set to: {self.target_directory}")

    def update_status(self, message):
        self.status_label.configure(text=f"Status: {message}")
        self.update_idletasks()

    def update_progressbar(self, value):
        self.progressbar.set(value)
        self.update_idletasks()

    def set_ui_elements_state(self, state):
        widgets_to_toggle = [
            self.convert_button, self.add_files_button, self.remove_selected_button,
            self.clear_list_button, self.select_target_button, self.language_menu,
            self.font_size_dropdown, self.main_font_dropdown, self.margin_entry,
            self.pdf_engine_dropdown, self.line_spacing_entry,
            self.save_settings_button, self.appearance_mode_optionemenu
        ]
        for widget in widgets_to_toggle:
            if widget: widget.configure(state=state)

    def start_conversion_thread(self):
        if not TESSERACT_AVAILABLE:
            tkinter.messagebox.showerror("Error", "Tesseract OCR not found or not configured correctly.")
            return
        if not self.selected_files:
            tkinter.messagebox.showerror("Error", "Please select at least one PDF file to convert.")
            return
        if not self.target_directory:
            tkinter.messagebox.showerror("Error", "Please select an output folder.")
            return

        self.set_ui_elements_state("disabled")
        self.update_progressbar(0)
        conversion_thread = threading.Thread(target=self.process_files, daemon=True)
        conversion_thread.start()

    def process_files(self):
        total_files = len(self.selected_files)
        ocr_lang = self.language_var.get()

        font_size = self.current_font_size.get()
        margin = self.current_margin.get().strip()
        main_font = self.current_main_font.get()
        pdf_engine = self.current_pdf_engine.get()
        line_spacing = self.current_line_spacing.get().strip()

        if not margin: margin = "0.7in"
        if not line_spacing: line_spacing = "1.0"
        else:
            try:
                float_val = float(line_spacing)
                if float_val <= 0:
                    print(f"Invalid line spacing: '{line_spacing}'. Using default '1.0'.")
                    line_spacing = "1.0"
            except ValueError:
                print(f"Invalid line spacing format: '{line_spacing}'. Using default '1.0'.")
                line_spacing = "1.0"

        print(f"Settings: Size={font_size}, Font={main_font}, Margin={margin}, Engine={pdf_engine}, Line Spacing={line_spacing}")
        
        tesseract_cmd_for_worker = getattr(pytesseract.pytesseract, 'tesseract_cmd', None)
        
        cpu_cores = os.cpu_count() or 1
        num_processes = max(1, min(cpu_cores, 8)) 
        maxtasksperchild = 10 
        print(f"Using {num_processes} parallel processes, Maxtasksperchild: {maxtasksperchild}")

        for i, pdf_path in enumerate(self.selected_files):
            current_file_basename = os.path.basename(pdf_path)
            base_name_no_ext = os.path.splitext(current_file_basename)[0]
            self.update_status(f"Processing: {current_file_basename} ({i+1}/{total_files})")
            
            raw_ocr_text = ""

            try:
                total_pages = 0
                try:
                    doc_meta = fitz.open(pdf_path)
                    total_pages = len(doc_meta)
                    doc_meta.close()
                except Exception as e_meta:
                    print(f"Could not open PDF/read page count ({current_file_basename}): {e_meta}")
                    self.update_status(f"Error (page count): {current_file_basename}")
                    raw_ocr_text = f"[Could not open document or read page count: {e_meta}]\n\n"
                    output_txt_path_error = os.path.join(self.target_directory, f"{base_name_no_ext}_ocr.txt")
                    try:
                        with open(output_txt_path_error, "w", encoding="utf-8") as f_txt_err:
                            f_txt_err.write(raw_ocr_text)
                        print(f"Text output for failed file: {output_txt_path_error}")
                    except Exception as e_write_txt_error:
                        print(f"Could not write TXT for failed file: {e_write_txt_error}")
                    continue 

                full_text_parts = [None] * total_pages
                
                if total_pages == 0:
                    print(f"PDF ({current_file_basename}) is empty or has no pages.")
                    raw_ocr_text = "[Document is empty or contains no pages]\n\n"
                else:
                    self.update_status(f"Starting OCR: {current_file_basename} - {total_pages} pages")
                    
                    page_task_args_list = [
                        (pdf_path, page_num, ocr_lang, 300, tesseract_cmd_for_worker) 
                        for page_num in range(total_pages)
                    ]
                    
                    pages_processed_count = 0
                    with multiprocessing.Pool(processes=num_processes, maxtasksperchild=maxtasksperchild) as pool:
                        results_iterator = pool.imap_unordered(ocr_page_worker_function, page_task_args_list)
                        
                        for page_num_result, page_text_result in results_iterator:
                            if 0 <= page_num_result < total_pages:
                                full_text_parts[page_num_result] = page_text_result
                            else:
                                print(f"Warning: Invalid page number ({page_num_result}) returned from OCR result.")
                            
                            pages_processed_count += 1
                            self.update_status(f"OCR: {current_file_basename} - Page {pages_processed_count}/{total_pages}")
                    
                    raw_ocr_text = "\n\n".join(text for text in full_text_parts if text is not None)

                output_txt_path = os.path.join(self.target_directory, f"{base_name_no_ext}_ocr.txt")
                try:
                    with open(output_txt_path, "w", encoding="utf-8") as f_txt:
                        f_txt.write(raw_ocr_text) 
                    print(f"Text output saved: {output_txt_path}")
                except Exception as e_write_txt:
                    print(f"Could not write TXT file ({output_txt_path}): {e_write_txt}")
                    self.update_status(f"TXT Write Error: {current_file_basename}")

                formatted_for_pandoc = format_text_with_heuristics(raw_ocr_text)
                pandoc_input_text = formatted_for_pandoc.replace('\\', r'\\') 

                output_pdf_path = os.path.join(self.target_directory, f"{base_name_no_ext}_ocr.pdf")
                pandoc_command = [
                    'pandoc', '-s', f'--pdf-engine={pdf_engine}',
                    '-V', f'fontsize={font_size}', '-V', f'geometry:margin={margin}',
                    '-V', 'papersize=a4', '-V', f'linestretch={line_spacing}',
                    '-f', 'markdown', '-o', output_pdf_path,
                    '-V', 'documentclass=scrartcl'
                ]
                if pdf_engine in ["xelatex", "lualatex"]:
                    pandoc_command.extend(['-V', f'mainfont={main_font}', '-V', 'lang=en-US']) # Changed lang to en-US
                elif pdf_engine == "pdflatex":
                    pandoc_command.extend(['-V', 'fontenc=T1', '-V', 'inputenc=utf8'])
                
                print(f"Executing Pandoc command: {' '.join(pandoc_command)}")
                process = subprocess.Popen(pandoc_command, stdin=subprocess.PIPE, text=True, encoding='utf-8',
                                           stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                stdout, stderr = process.communicate(input=pandoc_input_text)

                if process.returncode != 0:
                    error_message = f"Pandoc error ({current_file_basename}):\n{stderr.strip()}\n\nStdout:\n{stdout.strip()}"
                    print(error_message)
                    self.update_status(f"Pandoc error: {current_file_basename}")
                    self.after(0, lambda msg=error_message: tkinter.messagebox.showerror("Pandoc Error", msg))
                else:
                    print(f"PDF successfully converted: {output_pdf_path}")
                    if stderr.strip(): print(f"Pandoc warnings:\n{stderr.strip()}")

            except Exception as e_file:
                error_message_general = f"General error while processing file ({current_file_basename}): {e_file}"
                print(error_message_general)
                import traceback
                traceback.print_exc()
                self.update_status(f"Error: {current_file_basename}")
                self.after(0, lambda msg=error_message_general: tkinter.messagebox.showerror("Processing Error", msg))
            
            self.update_progressbar((i + 1) / total_files)

        self.update_status("All files processed!")
        self.after(0, lambda: tkinter.messagebox.showinfo("Complete", "Conversion process has finished."))
        self.set_ui_elements_state("normal")

if __name__ == "__main__":
    multiprocessing.freeze_support()
    app = App()
    app.mainloop()