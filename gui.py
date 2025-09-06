import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import threading
from pathlib import Path
import json
from datetime import datetime

# Import ydata-profiling with fallback
try:
    from ydata_profiling import ProfileReport
    YDATA_AVAILABLE = True
    PROFILING_LIB = "ydata-profiling"
except ImportError:
    try:
        from pandas_profiling import ProfileReport
        YDATA_AVAILABLE = True
        PROFILING_LIB = "pandas-profiling"
    except ImportError:
        YDATA_AVAILABLE = False
        PROFILING_LIB = None

class YDataProfilingGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("YData Profiling Report Generator")
        self.root.geometry("800x700")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.title_var = tk.StringVar(value="Data Analysis Report")
        self.minimal_var = tk.BooleanVar(value=False)
        self.explorative_var = tk.BooleanVar(value=True)
        self.dark_mode_var = tk.BooleanVar(value=False)
        self.orange_mode_var = tk.BooleanVar(value=False)
        self.progress_var = tk.DoubleVar()
        
        # Advanced settings
        self.correlations_var = tk.BooleanVar(value=True)
        self.missing_diagrams_var = tk.BooleanVar(value=True)
        self.duplicates_var = tk.BooleanVar(value=True)
        self.interactions_var = tk.BooleanVar(value=False)
        self.samples_var = tk.BooleanVar(value=True)
        
        self.create_widgets()
        self.check_dependencies()
        
    def check_dependencies(self):
        if not YDATA_AVAILABLE:
            messagebox.showerror(
                "Thiếu thư viện",
                "Vui lòng cài đặt một trong hai:\n" +
                "pip install ydata-profiling\n" + 
                "hoặc\n" +
                "pip install pandas-profiling"
            )
        else:
            self.log(f"Sử dụng thư viện: {PROFILING_LIB}")
            
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="YData Profiling Report Generator", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Chọn File Dữ Liệu", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="File Input:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        ttk.Entry(file_frame, textvariable=self.input_file, width=60).grid(row=0, column=1, 
                                                                          sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(file_frame, text="Chọn File", 
                  command=self.browse_input_file).grid(row=0, column=2, padx=(5, 0))
        
        ttk.Label(file_frame, text="File Output:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        ttk.Entry(file_frame, textvariable=self.output_file, width=60).grid(row=1, column=1, 
                                                                           sticky=(tk.W, tk.E), padx=(5, 5))
        ttk.Button(file_frame, text="Chọn Nơi Lưu", 
                  command=self.browse_output_file).grid(row=1, column=2, padx=(5, 0))
        
        # Report settings section
        settings_frame = ttk.LabelFrame(main_frame, text="Cài Đặt Báo Cáo", padding="10")
        settings_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        settings_frame.columnconfigure(1, weight=1)
        
        ttk.Label(settings_frame, text="Tiêu đề báo cáo:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(settings_frame, textvariable=self.title_var, width=50).grid(row=0, column=1, 
                                                                             sticky=(tk.W, tk.E), padx=(5, 0))
        
        # Report type
        type_frame = ttk.Frame(settings_frame)
        type_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(type_frame, text="Loại báo cáo:").grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(type_frame, text="Minimal (nhanh)", variable=self.minimal_var, 
                       value=True, command=self.on_minimal_selected).grid(row=0, column=1, padx=(10, 0))
        ttk.Radiobutton(type_frame, text="Explorative (đầy đủ)", variable=self.explorative_var, 
                       value=True, command=self.on_explorative_selected).grid(row=0, column=2, padx=(10, 0))
        
        # Theme settings
        theme_frame = ttk.Frame(settings_frame)
        theme_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Label(theme_frame, text="Giao diện:").grid(row=0, column=0, sticky=tk.W)
        ttk.Checkbutton(theme_frame, text="Dark Mode", 
                       variable=self.dark_mode_var).grid(row=0, column=1, padx=(10, 0))
        ttk.Checkbutton(theme_frame, text="Orange Theme", 
                       variable=self.orange_mode_var).grid(row=0, column=2, padx=(10, 0))
        
        # Advanced settings section
        advanced_frame = ttk.LabelFrame(main_frame, text="Cài Đặt Nâng Cao", padding="10")
        advanced_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Create checkboxes for advanced features
        advanced_options = [
            ("Phân tích tương quan", self.correlations_var),
            ("Biểu đồ dữ liệu thiếu", self.missing_diagrams_var),
            ("Phân tích trùng lặp", self.duplicates_var),
            ("Phân tích tương tác", self.interactions_var),
            ("Hiển thị mẫu dữ liệu", self.samples_var)
        ]
        
        for i, (text, var) in enumerate(advanced_options):
            row = i // 2
            col = i % 2
            ttk.Checkbutton(advanced_frame, text=text, variable=var).grid(
                row=row, column=col, sticky=tk.W, padx=(0, 20), pady=2
            )
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Tiến Độ", padding="10")
        progress_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.status_label = ttk.Label(progress_frame, text="Sẵn sàng tạo báo cáo")
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        # Log section
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=(10, 0))
        
        ttk.Button(button_frame, text="Tạo Báo Cáo", 
                  command=self.generate_report).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Xóa Log", 
                  command=self.clear_log).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Lưu Cài Đặt", 
                  command=self.save_settings).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Tải Cài Đặt", 
                  command=self.load_settings).pack(side=tk.LEFT)
        
        self.log("Ứng dụng đã khởi động thành công!")
        
    def on_minimal_selected(self):
        """Handle minimal radio button selection"""
        if self.minimal_var.get():
            self.explorative_var.set(False)
    
    def on_explorative_selected(self):
        """Handle explorative radio button selection"""
        if self.explorative_var.get():
            self.minimal_var.set(False)
        
    def log(self, message):
        """Add message to log with timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def clear_log(self):
        """Clear the log text"""
        self.log_text.delete(1.0, tk.END)
        
    def browse_input_file(self):
        """Browse for input file"""
        filetypes = [
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx *.xls"),
            ("JSON files", "*.json"),
            ("Parquet files", "*.parquet"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Chọn file dữ liệu",
            filetypes=filetypes
        )
        
        if filename:
            self.input_file.set(filename)
            # Auto-generate output filename
            base_name = Path(filename).stem
            output_path = Path(filename).parent / f"{base_name}_report.html"
            self.output_file.set(str(output_path))
            self.log(f"Đã chọn file input: {filename}")
            
    def browse_output_file(self):
        """Browse for output file"""
        filename = filedialog.asksaveasfilename(
            title="Chọn nơi lưu báo cáo",
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")]
        )
        
        if filename:
            self.output_file.set(filename)
            self.log(f"Đã chọn file output: {filename}")
            
    def load_data(self, file_path):
        """Load data from various file formats"""
        file_ext = Path(file_path).suffix.lower()
        
        try:
            if file_ext == '.csv':
                # Try different encodings
                for encoding in ['utf-8', 'latin-1', 'cp1252', 'utf-8-sig']:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding)
                        self.log(f"Đã đọc CSV với encoding {encoding}")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise ValueError("Không thể đọc file CSV với encoding hỗ trợ")
                    
            elif file_ext in ['.xlsx', '.xls']:
                df = pd.read_excel(file_path)
                self.log("Đã đọc file Excel")
                
            elif file_ext == '.json':
                df = pd.read_json(file_path)
                self.log("Đã đọc file JSON")
                
            elif file_ext == '.parquet':
                df = pd.read_parquet(file_path)
                self.log("Đã đọc file Parquet")
                
            else:
                raise ValueError(f"Định dạng file không được hỗ trợ: {file_ext}")
                
            self.log(f"Kích thước dữ liệu: {df.shape[0]} dòng, {df.shape[1]} cột")
            return df
            
        except Exception as e:
            raise Exception(f"Lỗi khi đọc file: {str(e)}")
            
    def create_profile_config(self):
        """Create configuration for ProfileReport"""
        # Base configuration
        config = {
            "title": self.title_var.get(),
        }
        
        # Report type - only set one of these
        if self.minimal_var.get():
            config["minimal"] = True
        else:
            config["explorative"] = True
        
        # Theme settings
        if self.dark_mode_var.get():
            config["dark_mode"] = True
        
        if self.orange_mode_var.get():
            config["orange_mode"] = True
        
        # Advanced settings - only add if not minimal
        if not self.minimal_var.get():
            try:
                config.update({
                    "correlations": {
                        "calculate": self.correlations_var.get()
                    },
                    "missing_diagrams": {
                        "calculate": self.missing_diagrams_var.get()
                    },
                    "duplicates": {
                        "calculate": self.duplicates_var.get()
                    },
                    "interactions": {
                        "calculate": self.interactions_var.get()
                    },
                    "samples": {
                        "calculate": self.samples_var.get()
                    }
                })
            except Exception as e:
                self.log(f"Cảnh báo: Một số cài đặt nâng cao có thể không được hỗ trợ: {e}")
        
        return config
        
    def generate_report_thread(self):
        """Generate report in separate thread"""
        try:
            if not YDATA_AVAILABLE:
                raise Exception(f"{PROFILING_LIB or 'YData/Pandas Profiling'} không khả dụng")
                
            input_path = self.input_file.get()
            output_path = self.output_file.get()
            
            if not input_path:
                raise Exception("Vui lòng chọn file input")
                
            if not output_path:
                raise Exception("Vui lòng chọn nơi lưu báo cáo")
                
            if not os.path.exists(input_path):
                raise Exception("File input không tồn tại")
                
            self.log("Bắt đầu tạo báo cáo...")
            self.log(f"Sử dụng thư viện: {PROFILING_LIB}")
            self.status_label.config(text="Đang đọc dữ liệu...")
            
            # Load data
            df = self.load_data(input_path)
            
            self.status_label.config(text="Đang tạo profile...")
            
            # Show progress for large datasets
            if df.shape[0] > 10000:
                self.log("Dataset lớn - có thể mất vài phút để xử lý...")
            
            # Create profile configuration
            config = self.create_profile_config()
            self.log(f"Cấu hình báo cáo: {'Minimal' if config.get('minimal') else 'Explorative'}")
            
            # Generate profile with error handling
            try:
                profile = ProfileReport(df, **config)
            except Exception as profile_error:
                self.log(f"Lỗi khi tạo profile với cấu hình đầy đủ: {str(profile_error)}")
                # Try with minimal config as fallback
                self.log("Thử lại với cấu hình tối thiểu...")
                minimal_config = {"title": self.title_var.get(), "minimal": True}
                profile = ProfileReport(df, **minimal_config)
            
            self.status_label.config(text="Đang lưu báo cáo...")
            
            # Save report
            profile.to_file(output_path)
            
            self.log(f"Báo cáo đã được tạo thành công: {output_path}")
            self.status_label.config(text="Hoàn thành!")
            
            # Show file size
            file_size = os.path.getsize(output_path) / 1024 / 1024
            self.log(f"Kích thước file báo cáo: {file_size:.2f} MB")
            
            # Ask if user wants to open the report
            if messagebox.askyesno("Thành công", 
                                 "Báo cáo đã được tạo thành công!\nBạn có muốn mở báo cáo không?"):
                try:
                    os.startfile(output_path)
                except:
                    # Alternative for non-Windows systems
                    import webbrowser
                    webbrowser.open(f'file://{os.path.abspath(output_path)}')
                
        except Exception as e:
            error_msg = f"Lỗi: {str(e)}"
            self.log(error_msg)
            self.status_label.config(text="Có lỗi xảy ra")
            messagebox.showerror("Lỗi", error_msg)
            
        finally:
            self.progress_bar.stop()
            
    def generate_report(self):
        """Start report generation in thread"""
        if not YDATA_AVAILABLE:
            messagebox.showerror("Lỗi", "Thư viện profiling không khả dụng!")
            return
            
        self.progress_bar.start(10)
        self.status_label.config(text="Đang xử lý...")
        
        # Run in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self.generate_report_thread)
        thread.daemon = True
        thread.start()
        
    def save_settings(self):
        """Save current settings to file"""
        settings = {
            "title": self.title_var.get(),
            "minimal": self.minimal_var.get(),
            "explorative": self.explorative_var.get(),
            "dark_mode": self.dark_mode_var.get(),
            "orange_mode": self.orange_mode_var.get(),
            "correlations": self.correlations_var.get(),
            "missing_diagrams": self.missing_diagrams_var.get(),
            "duplicates": self.duplicates_var.get(),
            "interactions": self.interactions_var.get(),
            "samples": self.samples_var.get()
        }
        
        filename = filedialog.asksaveasfilename(
            title="Lưu cài đặt",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(settings, f, indent=4, ensure_ascii=False)
                self.log(f"Đã lưu cài đặt: {filename}")
                messagebox.showinfo("Thành công", "Cài đặt đã được lưu!")
            except Exception as e:
                self.log(f"Lỗi khi lưu cài đặt: {str(e)}")
                messagebox.showerror("Lỗi", f"Không thể lưu cài đặt: {str(e)}")
                
    def load_settings(self):
        """Load settings from file"""
        filename = filedialog.askopenfilename(
            title="Tải cài đặt",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                # Apply settings
                self.title_var.set(settings.get("title", "Data Analysis Report"))
                self.minimal_var.set(settings.get("minimal", False))
                self.explorative_var.set(settings.get("explorative", True))
                self.dark_mode_var.set(settings.get("dark_mode", False))
                self.orange_mode_var.set(settings.get("orange_mode", False))
                self.correlations_var.set(settings.get("correlations", True))
                self.missing_diagrams_var.set(settings.get("missing_diagrams", True))
                self.duplicates_var.set(settings.get("duplicates", True))
                self.interactions_var.set(settings.get("interactions", False))
                self.samples_var.set(settings.get("samples", True))
                
                self.log(f"Đã tải cài đặt: {filename}")
                messagebox.showinfo("Thành công", "Cài đặt đã được tải!")
                
            except Exception as e:
                self.log(f"Lỗi khi tải cài đặt: {str(e)}")
                messagebox.showerror("Lỗi", f"Không thể tải cài đặt: {str(e)}")

def main():
    root = tk.Tk()
    app = YDataProfilingGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()