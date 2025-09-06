import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import os
import threading
from pathlib import Path
import json
from datetime import datetime
import sys
import webbrowser
from typing import Dict, Any, Optional
import warnings

# Suppress warnings
warnings.filterwarnings('ignore')

# Import profiling libraries with fallback
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

class ModernStyle:
    """Modern styling for the GUI"""
    
    # Color scheme
    COLORS = {
        'primary': '#2563eb',
        'secondary': '#64748b',
        'success': '#10b981',
        'warning': '#f59e0b',
        'error': '#ef4444',
        'background': '#f8fafc',
        'surface': '#ffffff',
        'text_primary': '#1e293b',
        'text_secondary': '#64748b',
        'border': '#e2e8f0',
        'accent': '#8b5cf6'
    }
    
    @staticmethod
    def configure_styles():
        """Configure modern styles for ttk widgets"""
        style = ttk.Style()
        
        # Configure main styles
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 18, 'bold'),
                       foreground=ModernStyle.COLORS['text_primary'])
        
        style.configure('Subtitle.TLabel', 
                       font=('Segoe UI', 10, 'bold'),
                       foreground=ModernStyle.COLORS['text_secondary'])
        
        style.configure('Modern.TButton',
                       font=('Segoe UI', 9),
                       padding=(12, 8))
        
        style.configure('Primary.TButton',
                       font=('Segoe UI', 9, 'bold'),
                       padding=(12, 8))
        
        style.configure('Success.TButton',
                       font=('Segoe UI', 9, 'bold'),
                       padding=(12, 8))
        
        style.configure('Modern.TLabelFrame',
                       padding=15,
                       font=('Segoe UI', 9, 'bold'))
        
        style.configure('Modern.TLabelFrame.Label',
                       font=('Segoe UI', 9, 'bold'),
                       foreground=ModernStyle.COLORS['primary'])

class DataProfilerGUI:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.initialize_variables()
        self.create_main_interface()
        self.check_dependencies()
        self.load_last_session()
        
    def setup_window(self):
        """Setup main window with modern design"""
        self.root.title("🚀 Advanced Data Profiling Studio v2.0")
        self.root.geometry("1200x900")
        self.root.minsize(1000, 700)
        self.root.configure(bg=ModernStyle.COLORS['background'])
        
        # Configure styles
        ModernStyle.configure_styles()
        
        # Center window on screen
        self.center_window()
        
        # Configure icon (if available)
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
            
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def initialize_variables(self):
        """Initialize all GUI variables"""
        # File paths
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.output_format = tk.StringVar(value="html")
        
        # Report settings
        self.title_var = tk.StringVar(value="Advanced Data Analysis Report")
        self.minimal_var = tk.BooleanVar(value=False)
        self.explorative_var = tk.BooleanVar(value=True)
        self.dataset_description = tk.StringVar()
        
        # Theme and appearance
        self.theme_var = tk.StringVar(value="default")
        self.color_scheme = tk.StringVar(value="blue")
        
        # Advanced analysis options
        self.correlations_var = tk.BooleanVar(value=True)
        self.missing_diagrams_var = tk.BooleanVar(value=True)
        self.duplicates_var = tk.BooleanVar(value=True)
        self.interactions_var = tk.BooleanVar(value=False)
        self.samples_var = tk.BooleanVar(value=True)
        
        # Performance settings
        self.sample_size = tk.IntVar(value=0)  # 0 = use all data
        self.timeout_var = tk.IntVar(value=600)  # 10 minutes
        self.memory_limit = tk.IntVar(value=2048)  # 2GB in MB
        
        # Data preprocessing
        self.auto_clean = tk.BooleanVar(value=False)
        self.handle_missing = tk.StringVar(value="skip")
        self.encoding_detect = tk.BooleanVar(value=True)
        
        # Export options
        self.include_config = tk.BooleanVar(value=True)
        self.compress_output = tk.BooleanVar(value=False)
        
        # Progress tracking
        self.progress_var = tk.DoubleVar()
        self.current_step = tk.StringVar(value="Ready")
        
    def create_main_interface(self):
        """Create the main interface with modern design"""
        # Create main container with notebook
        self.main_container = ttk.Frame(self.root, padding="20")
        self.main_container.pack(fill=tk.BOTH, expand=True)
        
        # Create header
        self.create_header()
        
        # Create notebook for organized tabs
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        
        # Create tabs
        self.create_basic_tab()
        self.create_advanced_tab()
        self.create_performance_tab()
        self.create_export_tab()
        self.create_log_tab()
        
        # Create bottom panel
        self.create_bottom_panel()
        
    def create_header(self):
        """Create modern header with title and info"""
        header_frame = ttk.Frame(self.main_container)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Title and subtitle
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT)
        
        title_label = ttk.Label(title_frame, text="🚀 Advanced Data Profiling Studio", 
                               style='Title.TLabel')
        title_label.pack(anchor=tk.W)
        
        subtitle_label = ttk.Label(title_frame, 
                                  text="Professional data analysis and reporting tool",
                                  style='Subtitle.TLabel')
        subtitle_label.pack(anchor=tk.W)
        
        # Status indicator
        status_frame = ttk.Frame(header_frame)
        status_frame.pack(side=tk.RIGHT)
        
        self.status_indicator = ttk.Label(status_frame, text="●", 
                                         foreground=ModernStyle.COLORS['success'],
                                         font=('Arial', 16))
        self.status_indicator.pack(side=tk.RIGHT, padx=(0, 10))
        
        self.lib_status = ttk.Label(status_frame, 
                                   text=f"Library: {PROFILING_LIB or 'Not Available'}",
                                   style='Subtitle.TLabel')
        self.lib_status.pack(side=tk.RIGHT)
        
    def create_basic_tab(self):
        """Create basic settings tab"""
        basic_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(basic_frame, text="📁 Basic Settings")
        
        # File selection section
        self.create_file_section(basic_frame)
        
        # Report configuration section
        self.create_report_config_section(basic_frame)
        
        # Quick analysis section
        self.create_quick_analysis_section(basic_frame)
        
    def create_file_section(self, parent):
        """Create file selection section with drag & drop"""
        file_frame = ttk.LabelFrame(parent, text="📂 Data Source", style='Modern.TLabelFrame')
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Input file selection
        input_frame = ttk.Frame(file_frame)
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(input_frame, text="Input File:", font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W)
        
        input_select_frame = ttk.Frame(input_frame)
        input_select_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.input_entry = ttk.Entry(input_select_frame, textvariable=self.input_file, 
                                    font=('Segoe UI', 9))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        ttk.Button(input_select_frame, text="📁 Browse", 
                  command=self.browse_input_file, style='Modern.TButton').pack(side=tk.RIGHT, padx=(0, 5))
        
        ttk.Button(input_select_frame, text="🔍 Preview", 
                  command=self.preview_data, style='Modern.TButton').pack(side=tk.RIGHT)
        
        # File info display
        self.file_info_frame = ttk.Frame(file_frame)
        self.file_info_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.file_info_label = ttk.Label(self.file_info_frame, text="", 
                                        style='Subtitle.TLabel')
        self.file_info_label.pack(anchor=tk.W)
        
        # Output file selection
        output_frame = ttk.Frame(file_frame)
        output_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(output_frame, text="Output File:", font=('Segoe UI', 9, 'bold')).pack(anchor=tk.W)
        
        output_select_frame = ttk.Frame(output_frame)
        output_select_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Entry(output_select_frame, textvariable=self.output_file, 
                 font=('Segoe UI', 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        ttk.Button(output_select_frame, text="📁 Save As", 
                  command=self.browse_output_file, style='Modern.TButton').pack(side=tk.RIGHT)
        
        # Format selection
        format_frame = ttk.Frame(file_frame)
        format_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(format_frame, text="Output Format:", font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT)
        
        formats = [("HTML Report", "html"), ("JSON Data", "json")]
        for i, (text, value) in enumerate(formats):
            ttk.Radiobutton(format_frame, text=text, variable=self.output_format, 
                           value=value).pack(side=tk.LEFT, padx=(20 if i == 0 else 10, 0))
        
    def create_report_config_section(self, parent):
        """Create report configuration section"""
        config_frame = ttk.LabelFrame(parent, text="⚙️ Report Configuration", 
                                     style='Modern.TLabelFrame')
        config_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Title and description
        title_frame = ttk.Frame(config_frame)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(title_frame, text="Report Title:", font=('Segoe UI', 9, 'bold')).grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(title_frame, textvariable=self.title_var, font=('Segoe UI', 9), 
                 width=40).grid(row=0, column=1, padx=(10, 0), sticky=tk.W)
        
        ttk.Label(title_frame, text="Description:", font=('Segoe UI', 9, 'bold')).grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        ttk.Entry(title_frame, textvariable=self.dataset_description, font=('Segoe UI', 9), 
                 width=40).grid(row=1, column=1, padx=(10, 0), sticky=tk.W, pady=(10, 0))
        
        # Report type
        type_frame = ttk.Frame(config_frame)
        type_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(type_frame, text="Analysis Type:", font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT)
        
        ttk.Radiobutton(type_frame, text="⚡ Quick (Fast)", variable=self.minimal_var, 
                       value=True, command=self.on_type_changed).pack(side=tk.LEFT, padx=(20, 0))
        ttk.Radiobutton(type_frame, text="🔬 Comprehensive (Detailed)", variable=self.explorative_var, 
                       value=True, command=self.on_type_changed).pack(side=tk.LEFT, padx=(10, 0))
        
        # Theme selection
        theme_frame = ttk.Frame(config_frame)
        theme_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(theme_frame, text="Theme:", font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT)
        
        theme_combo = ttk.Combobox(theme_frame, textvariable=self.theme_var, 
                                  values=["default", "dark", "orange", "united"], 
                                  state="readonly", width=15)
        theme_combo.pack(side=tk.LEFT, padx=(20, 20))
        
        ttk.Label(theme_frame, text="Color:", font=('Segoe UI', 9, 'bold')).pack(side=tk.LEFT)
        
        color_combo = ttk.Combobox(theme_frame, textvariable=self.color_scheme,
                                  values=["blue", "red", "green", "purple", "orange"],
                                  state="readonly", width=15)
        color_combo.pack(side=tk.LEFT, padx=(10, 0))
        
    def create_quick_analysis_section(self, parent):
        """Create quick analysis section"""
        quick_frame = ttk.LabelFrame(parent, text="🚀 Quick Actions", 
                                    style='Modern.TLabelFrame')
        quick_frame.pack(fill=tk.X)
        
        buttons_frame = ttk.Frame(quick_frame)
        buttons_frame.pack(pady=10)
        
        ttk.Button(buttons_frame, text="🔍 Data Overview", 
                  command=self.show_data_overview, 
                  style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(buttons_frame, text="📊 Column Summary", 
                  command=self.show_column_summary, 
                  style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(buttons_frame, text="🎯 Data Quality", 
                  command=self.check_data_quality, 
                  style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(buttons_frame, text="📈 Generate Report", 
                  command=self.generate_report, 
                  style='Primary.TButton').pack(side=tk.LEFT)
        
    def create_advanced_tab(self):
        """Create advanced settings tab"""
        advanced_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(advanced_frame, text="🔧 Advanced")
        
        # Analysis components
        components_frame = ttk.LabelFrame(advanced_frame, text="📊 Analysis Components", 
                                         style='Modern.TLabelFrame')
        components_frame.pack(fill=tk.X, pady=(0, 15))
        
        components_grid = ttk.Frame(components_frame)
        components_grid.pack(fill=tk.X)
        
        # Create checkboxes in a grid
        components = [
            ("Correlation Analysis", self.correlations_var, "Analyze relationships between variables"),
            ("Missing Data Diagrams", self.missing_diagrams_var, "Visualize missing data patterns"),
            ("Duplicate Analysis", self.duplicates_var, "Identify and analyze duplicate records"),
            ("Variable Interactions", self.interactions_var, "Analyze interactions between variables"),
            ("Sample Data Display", self.samples_var, "Show sample data in the report")
        ]
        
        for i, (text, var, tooltip) in enumerate(components):
            row, col = divmod(i, 2)
            cb_frame = ttk.Frame(components_grid)
            cb_frame.grid(row=row, column=col, sticky=tk.W, padx=(0, 30), pady=5)
            
            cb = ttk.Checkbutton(cb_frame, text=text, variable=var)
            cb.pack(side=tk.LEFT)
            
            # Add tooltip (simplified)
            ttk.Label(cb_frame, text="ⓘ", foreground=ModernStyle.COLORS['secondary']).pack(side=tk.LEFT, padx=(5, 0))
        
        # Data preprocessing
        preprocess_frame = ttk.LabelFrame(advanced_frame, text="🧹 Data Preprocessing", 
                                         style='Modern.TLabelFrame')
        preprocess_frame.pack(fill=tk.X, pady=(0, 15))
        
        preprocess_options = ttk.Frame(preprocess_frame)
        preprocess_options.pack(fill=tk.X)
        
        ttk.Checkbutton(preprocess_options, text="Auto-detect encoding", 
                       variable=self.encoding_detect).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Checkbutton(preprocess_options, text="Auto-clean data", 
                       variable=self.auto_clean).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Label(preprocess_options, text="Missing values:").pack(side=tk.LEFT, padx=(0, 10))
        
        missing_combo = ttk.Combobox(preprocess_options, textvariable=self.handle_missing,
                                    values=["skip", "show", "interpolate"], 
                                    state="readonly", width=12)
        missing_combo.pack(side=tk.LEFT)
        
    def create_performance_tab(self):
        """Create performance settings tab"""
        perf_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(perf_frame, text="⚡ Performance")
        
        # Sampling settings
        sampling_frame = ttk.LabelFrame(perf_frame, text="📊 Data Sampling", 
                                       style='Modern.TLabelFrame')
        sampling_frame.pack(fill=tk.X, pady=(0, 15))
        
        sampling_grid = ttk.Frame(sampling_frame)
        sampling_grid.pack(fill=tk.X)
        
        ttk.Label(sampling_grid, text="Sample size (0 = all data):").grid(row=0, column=0, sticky=tk.W)
        sample_spinbox = ttk.Spinbox(sampling_grid, from_=0, to=1000000, increment=1000,
                                    textvariable=self.sample_size, width=15)
        sample_spinbox.grid(row=0, column=1, padx=(10, 0), sticky=tk.W)
        
        ttk.Label(sampling_grid, text="Timeout (seconds):").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        timeout_spinbox = ttk.Spinbox(sampling_grid, from_=60, to=3600, increment=60,
                                     textvariable=self.timeout_var, width=15)
        timeout_spinbox.grid(row=1, column=1, padx=(10, 0), sticky=tk.W, pady=(10, 0))
        
        ttk.Label(sampling_grid, text="Memory limit (MB):").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        memory_spinbox = ttk.Spinbox(sampling_grid, from_=512, to=8192, increment=512,
                                    textvariable=self.memory_limit, width=15)
        memory_spinbox.grid(row=2, column=1, padx=(10, 0), sticky=tk.W, pady=(10, 0))
        
        # Performance tips
        tips_frame = ttk.LabelFrame(perf_frame, text="💡 Performance Tips", 
                                   style='Modern.TLabelFrame')
        tips_frame.pack(fill=tk.BOTH, expand=True)
        
        tips_text = tk.Text(tips_frame, height=10, wrap=tk.WORD, 
                           font=('Segoe UI', 9), bg=ModernStyle.COLORS['surface'])
        tips_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tips_content = """
💡 Performance Optimization Tips:

• For large datasets (>100k rows), consider using sampling to speed up analysis
• Quick analysis mode is 5-10x faster than comprehensive mode
• Disable unused analysis components to improve performance
• Use CSV files instead of Excel when possible for better performance
• Close other memory-intensive applications during analysis
• For very large files, consider using Parquet format for better I/O performance
• Set appropriate timeout values based on your dataset size
        """.strip()
        
        tips_text.insert(tk.END, tips_content)
        tips_text.config(state=tk.DISABLED)
        
    def create_export_tab(self):
        """Create export settings tab"""
        export_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(export_frame, text="💾 Export")
        
        # Export options
        options_frame = ttk.LabelFrame(export_frame, text="📁 Export Options", 
                                      style='Modern.TLabelFrame')
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        export_options = ttk.Frame(options_frame)
        export_options.pack(fill=tk.X)
        
        ttk.Checkbutton(export_options, text="Include configuration file", 
                       variable=self.include_config).pack(anchor=tk.W, pady=5)
        
        ttk.Checkbutton(export_options, text="Compress output (ZIP)", 
                       variable=self.compress_output).pack(anchor=tk.W, pady=5)
        
        # Export actions
        actions_frame = ttk.LabelFrame(export_frame, text="🚀 Export Actions", 
                                      style='Modern.TLabelFrame')
        actions_frame.pack(fill=tk.X, pady=(0, 15))
        
        actions_buttons = ttk.Frame(actions_frame)
        actions_buttons.pack(pady=10)
        
        ttk.Button(actions_buttons, text="💾 Save Configuration", 
                  command=self.save_configuration, 
                  style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(actions_buttons, text="📂 Load Configuration", 
                  command=self.load_configuration, 
                  style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(actions_buttons, text="🔄 Reset to Defaults", 
                  command=self.reset_to_defaults, 
                  style='Modern.TButton').pack(side=tk.LEFT)
        
        # Recent files
        recent_frame = ttk.LabelFrame(export_frame, text="📋 Recent Files", 
                                     style='Modern.TLabelFrame')
        recent_frame.pack(fill=tk.BOTH, expand=True)
        
        # Recent files listbox with scrollbar
        recent_list_frame = ttk.Frame(recent_frame)
        recent_list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.recent_listbox = tk.Listbox(recent_list_frame, font=('Segoe UI', 9))
        recent_scrollbar = ttk.Scrollbar(recent_list_frame, orient=tk.VERTICAL, 
                                        command=self.recent_listbox.yview)
        
        self.recent_listbox.configure(yscrollcommand=recent_scrollbar.set)
        self.recent_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        recent_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.recent_listbox.bind('<Double-Button-1>', self.load_recent_file)
        
    def create_log_tab(self):
        """Create log and monitoring tab"""
        log_frame = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(log_frame, text="📋 Logs")
        
        # Progress section
        progress_frame = ttk.LabelFrame(log_frame, text="⏳ Progress", 
                                       style='Modern.TLabelFrame')
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', 
                                           variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # Current step
        self.step_label = ttk.Label(progress_frame, textvariable=self.current_step, 
                                   font=('Segoe UI', 9))
        self.step_label.pack(anchor=tk.W)
        
        # Log section
        log_display_frame = ttk.LabelFrame(log_frame, text="📝 Activity Log", 
                                          style='Modern.TLabelFrame')
        log_display_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create log text with scrollbar
        log_text_frame = ttk.Frame(log_display_frame)
        log_text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_text_frame, 
                                                 font=('Consolas', 9),
                                                 bg=ModernStyle.COLORS['surface'],
                                                 fg=ModernStyle.COLORS['text_primary'])
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Log controls
        log_controls = ttk.Frame(log_display_frame)
        log_controls.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(log_controls, text="🗑️ Clear Log", 
                  command=self.clear_log, 
                  style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(log_controls, text="💾 Save Log", 
                  command=self.save_log, 
                  style='Modern.TButton').pack(side=tk.LEFT)
        
    def create_bottom_panel(self):
        """Create bottom panel with main controls"""
        bottom_frame = ttk.Frame(self.main_container)
        bottom_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Left side - status
        status_frame = ttk.Frame(bottom_frame)
        status_frame.pack(side=tk.LEFT)
        
        self.main_status_label = ttk.Label(status_frame, text="Ready to analyze data", 
                                          style='Subtitle.TLabel')
        self.main_status_label.pack()
        
        # Right side - main actions
        action_frame = ttk.Frame(bottom_frame)
        action_frame.pack(side=tk.RIGHT)
        
        ttk.Button(action_frame, text="❌ Cancel", 
                  command=self.cancel_operation, 
                  style='Modern.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(action_frame, text="📊 Generate Report", 
                  command=self.generate_report, 
                  style='Success.TButton').pack(side=tk.LEFT)
        
    # Event handlers and utility methods
    def on_type_changed(self):
        """Handle report type change"""
        if self.minimal_var.get():
            self.explorative_var.set(False)
            self.interactions_var.set(False)  # Disable heavy analysis
        else:
            self.minimal_var.set(False)
            
    def browse_input_file(self):
        """Browse for input file with enhanced file type support"""
        filetypes = [
            ("All Data Files", "*.csv;*.xlsx;*.xls;*.json;*.parquet;*.tsv;*.txt"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx;*.xls"),
            ("JSON files", "*.json"),
            ("Parquet files", "*.parquet"),
            ("Tab-separated", "*.tsv;*.txt"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=filetypes
        )
        
        if filename:
            self.input_file.set(filename)
            self.update_file_info(filename)
            self.auto_generate_output_path(filename)
            self.add_to_recent_files(filename)
            self.log(f"📁 Selected input file: {Path(filename).name}")
            
    def update_file_info(self, filename):
        """Update file information display"""
        try:
            file_path = Path(filename)
            file_size = file_path.stat().st_size
            size_mb = file_size / (1024 * 1024)
            
            info_text = f"📄 {file_path.name} ({size_mb:.1f} MB)"
            self.file_info_label.config(text=info_text)
            
            # Try to get basic data info
            try:
                df = self.load_data_sample(filename, nrows=1000)
                info_text += f" • {df.shape[0]}+ rows × {df.shape[1]} columns"
                self.file_info_label.config(text=info_text)
            except:
                pass
                
        except Exception as e:
            self.file_info_label.config(text="❌ Error reading file info")
            
    def auto_generate_output_path(self, input_path):
        """Auto-generate output file path"""
        input_file = Path(input_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if self.output_format.get() == "html":
            output_name = f"{input_file.stem}_report_{timestamp}.html"
        else:
            output_name = f"{input_file.stem}_profile_{timestamp}.json"
            
        output_path = input_file.parent / output_name
        self.output_file.set(str(output_path))
        
    def browse_output_file(self):
        """Browse for output file location"""
        if self.output_format.get() == "html":
            filetypes = [("HTML files", "*.html"), ("All files", "*.*")]
            default_ext = ".html"
        else:
            filetypes = [("JSON files", "*.json"), ("All files", "*.*")]
            default_ext = ".json"
            
        filename = filedialog.asksaveasfilename(
            title="Save Report As",
            defaultextension=default_ext,
            filetypes=filetypes
        )
        
        if filename:
            self.output_file.set(filename)
            self.log(f"💾 Output location set: {Path(filename).name}")
            
    def preview_data(self):
        """Show data preview window"""
        if not self.input_file.get():
            messagebox.showwarning("Warning", "Please select an input file first")
            return
            
        try:
            df = self.load_data_sample(self.input_file.get(), nrows=100)
            self.show_data_preview_window(df)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot preview data: {str(e)}")
            
    def show_data_preview_window(self, df):
        """Display data preview in new window"""
        preview_window = tk.Toplevel(self.root)
        preview_window.title("🔍 Data Preview")
        preview_window.geometry("800x600")
        
        # Create treeview for data display
        tree_frame = ttk.Frame(preview_window, padding="10")
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Info label
        info_label = ttk.Label(tree_frame, 
                              text=f"📊 Showing first 100 rows • Total: {df.shape[0]} rows × {df.shape[1]} columns",
                              font=('Segoe UI', 9, 'bold'))
        info_label.pack(pady=(0, 10))
        
        # Treeview with scrollbars
        tree_container = ttk.Frame(tree_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(tree_container, show='headings')
        
        # Configure columns
        columns = list(df.columns)
        tree['columns'] = columns
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, minwidth=50)
            
        # Add scrollbars
        v_scrollbar = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=tree.xview)
        
        tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack widgets
        tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        # Insert data
        for _, row in df.iterrows():
            values = [str(val)[:50] + "..." if len(str(val)) > 50 else str(val) for val in row]
            tree.insert('', tk.END, values=values)
            
    def show_data_overview(self):
        """Show quick data overview"""
        if not self.input_file.get():
            messagebox.showwarning("Warning", "Please select an input file first")
            return
            
        try:
            df = self.load_data_sample(self.input_file.get())
            overview = self.get_data_overview(df)
            self.show_info_dialog("📊 Data Overview", overview)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot analyze data: {str(e)}")
            
    def get_data_overview(self, df):
        """Generate data overview text"""
        overview = f"""
📊 Dataset Overview:

📏 Dimensions: {df.shape[0]:,} rows × {df.shape[1]} columns

📈 Column Types:
{df.dtypes.value_counts().to_string()}

🔍 Missing Values: {df.isnull().sum().sum():,} total
({(df.isnull().sum().sum() / df.size * 100):.1f}% of all values)

💾 Memory Usage: {df.memory_usage(deep=True).sum() / 1024**2:.1f} MB

📊 Numeric Columns: {len(df.select_dtypes(include=[np.number]).columns)}
📝 Text Columns: {len(df.select_dtypes(include=['object']).columns)}
📅 Date Columns: {len(df.select_dtypes(include=['datetime']).columns)}
        """.strip()
        
        return overview
        
    def show_column_summary(self):
        """Show column summary information"""
        if not self.input_file.get():
            messagebox.showwarning("Warning", "Please select an input file first")
            return
            
        try:
            df = self.load_data_sample(self.input_file.get())
            summary = self.get_column_summary(df)
            self.show_info_dialog("📋 Column Summary", summary)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot analyze columns: {str(e)}")
            
    def get_column_summary(self, df):
        """Generate column summary text"""
        summary = "📋 Column Analysis:\n\n"
        
        for col in df.columns:
            dtype = str(df[col].dtype)
            null_count = df[col].isnull().sum()
            null_pct = (null_count / len(df)) * 100
            unique_count = df[col].nunique()
            
            summary += f"📊 {col}:\n"
            summary += f"   Type: {dtype}\n"
            summary += f"   Missing: {null_count:,} ({null_pct:.1f}%)\n"
            summary += f"   Unique: {unique_count:,}\n"
            
            if df[col].dtype in ['int64', 'float64']:
                summary += f"   Range: {df[col].min():.2f} to {df[col].max():.2f}\n"
                
            summary += "\n"
            
        return summary
        
    def check_data_quality(self):
        """Perform data quality check"""
        if not self.input_file.get():
            messagebox.showwarning("Warning", "Please select an input file first")
            return
            
        try:
            df = self.load_data_sample(self.input_file.get())
            quality_report = self.generate_quality_report(df)
            self.show_info_dialog("🎯 Data Quality Report", quality_report)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot check data quality: {str(e)}")
            
    def generate_quality_report(self, df):
        """Generate data quality report"""
        total_cells = df.size
        missing_cells = df.isnull().sum().sum()
        duplicate_rows = df.duplicated().sum()
        
        # Calculate quality score
        missing_penalty = (missing_cells / total_cells) * 50
        duplicate_penalty = (duplicate_rows / len(df)) * 30
        quality_score = max(0, 100 - missing_penalty - duplicate_penalty)
        
        # Determine quality level
        if quality_score >= 90:
            quality_level = "🟢 Excellent"
        elif quality_score >= 70:
            quality_level = "🟡 Good"
        elif quality_score >= 50:
            quality_level = "🟠 Fair"
        else:
            quality_level = "🔴 Poor"
            
        report = f"""
🎯 Data Quality Assessment:

Overall Quality: {quality_level} ({quality_score:.1f}/100)

📊 Issues Found:
• Missing Values: {missing_cells:,} ({(missing_cells/total_cells)*100:.1f}%)
• Duplicate Rows: {duplicate_rows:,} ({(duplicate_rows/len(df))*100:.1f}%)

🔍 Column-level Issues:
"""
        
        for col in df.columns:
            issues = []
            if df[col].isnull().sum() > 0:
                issues.append(f"{df[col].isnull().sum():,} missing")
            if len(df[col].unique()) == 1:
                issues.append("constant values")
            if df[col].dtype == 'object' and df[col].str.len().var() == 0:
                issues.append("uniform length")
                
            if issues:
                report += f"  • {col}: {', '.join(issues)}\n"
                
        if quality_score < 70:
            report += """
💡 Recommendations:
• Consider data cleaning before analysis
• Handle missing values appropriately
• Remove or investigate duplicate records
• Validate data consistency
"""
        
        return report.strip()
        
    def show_info_dialog(self, title, message):
        """Show information dialog with scrollable text"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("600x500")
        dialog.grab_set()
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        text_widget = scrolledtext.ScrolledText(text_frame, wrap=tk.WORD, 
                                               font=('Consolas', 9))
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, message)
        text_widget.config(state=tk.DISABLED)
        
        # Close button
        ttk.Button(main_frame, text="Close", 
                  command=dialog.destroy, 
                  style='Modern.TButton').pack()
                  
    def load_data_sample(self, file_path, nrows=None):
        """Load a sample of data for preview/analysis"""
        return self.load_data(file_path, nrows=nrows)
        
    def load_data(self, file_path, nrows=None):
        """Load data from various file formats"""
        file_ext = Path(file_path).suffix.lower()
        
        try:
            if file_ext == '.csv':
                # Auto-detect encoding if enabled
                if self.encoding_detect.get():
                    encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
                else:
                    encodings = ['utf-8']
                    
                for encoding in encodings:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding, nrows=nrows)
                        self.log(f"✅ CSV loaded with {encoding} encoding")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise ValueError("Cannot read CSV with supported encodings")
                    
            elif file_ext in ['.xlsx', '.xls']:
                df = pd.read_excel(file_path, nrows=nrows)
                self.log("✅ Excel file loaded")
                
            elif file_ext == '.json':
                df = pd.read_json(file_path)
                if nrows:
                    df = df.head(nrows)
                self.log("✅ JSON file loaded")
                
            elif file_ext == '.parquet':
                df = pd.read_parquet(file_path)
                if nrows:
                    df = df.head(nrows)
                self.log("✅ Parquet file loaded")
                
            elif file_ext in ['.tsv', '.txt']:
                for encoding in ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']:
                    try:
                        df = pd.read_csv(file_path, sep='\t', encoding=encoding, nrows=nrows)
                        self.log(f"✅ TSV loaded with {encoding} encoding")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    raise ValueError("Cannot read TSV with supported encodings")
                    
            else:
                raise ValueError(f"Unsupported file format: {file_ext}")
                
            # Auto-clean data if enabled
            if self.auto_clean.get() and nrows is None:  # Only for full data load
                df = self.clean_data(df)
                
            if nrows is None:  # Only log for full data load
                self.log(f"📊 Data shape: {df.shape[0]:,} rows × {df.shape[1]} columns")
                
            return df
            
        except Exception as e:
            raise Exception(f"Error loading data: {str(e)}")
            
    def clean_data(self, df):
        """Perform basic data cleaning"""
        original_shape = df.shape
        
        # Handle missing values based on setting
        if self.handle_missing.get() == "skip":
            # Remove columns with all missing values
            df = df.dropna(axis=1, how='all')
        elif self.handle_missing.get() == "interpolate":
            # Simple interpolation for numeric columns
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            df[numeric_cols] = df[numeric_cols].interpolate()
            
        # Remove duplicate rows if enabled
        if self.duplicates_var.get():
            df = df.drop_duplicates()
            
        self.log(f"🧹 Data cleaned: {original_shape} → {df.shape}")
        return df
        
    def create_profile_config(self):
        """Create comprehensive ProfileReport configuration"""
        config = {
            "title": self.title_var.get(),
            "dataset": {
                "description": self.dataset_description.get(),
                "copyright_holder": "Data Profiling Studio",
                "copyright_year": datetime.now().year,
            }
        }
        
        # Report type
        if self.minimal_var.get():
            config["minimal"] = True
        else:
            config["explorative"] = True
            
        # Theme configuration
        theme = self.theme_var.get()
        if theme != "default":
            if theme == "dark":
                config["dark_mode"] = True
            elif theme == "orange":
                config["orange_mode"] = True
            elif theme == "united":
                config["theme"] = "united"
                
        # Advanced analysis settings
        if not self.minimal_var.get():
            config.update({
                "correlations": {
                    "calculate": self.correlations_var.get(),
                    "threshold": 0.9
                },
                "missing_diagrams": {
                    "calculate": self.missing_diagrams_var.get()
                },
                "duplicates": {
                    "calculate": self.duplicates_var.get()
                },
                "interactions": {
                    "calculate": self.interactions_var.get(),
                    "threshold": 0.1
                },
                "samples": {
                    "calculate": self.samples_var.get(),
                    "head": 10,
                    "tail": 10
                }
            })
            
        # Performance settings
        if self.sample_size.get() > 0:
            config["sample"] = {"n": self.sample_size.get()}
            
        return config
        
    def generate_report_thread(self):
        """Generate profiling report in separate thread"""
        try:
            self.update_progress(0, "Initializing...")
            
            if not YDATA_AVAILABLE:
                raise Exception("Profiling library not available")
                
            input_path = self.input_file.get()
            output_path = self.output_file.get()
            
            if not input_path or not output_path:
                raise Exception("Please select both input and output files")
                
            if not os.path.exists(input_path):
                raise Exception("Input file does not exist")
                
            self.update_progress(10, "Loading data...")
            self.log("🚀 Starting report generation...")
            
            # Load data with sampling if specified
            if self.sample_size.get() > 0:
                df = self.load_data(input_path, nrows=self.sample_size.get())
                self.log(f"📊 Using sample of {self.sample_size.get():,} rows")
            else:
                df = self.load_data(input_path)
                
            self.update_progress(30, "Configuring analysis...")
            
            # Show dataset info
            if df.shape[0] > 100000:
                self.log("⚠️ Large dataset detected - this may take several minutes")
                
            # Create configuration
            config = self.create_profile_config()
            analysis_type = "Quick" if config.get("minimal") else "Comprehensive"
            self.log(f"⚙️ Configuration: {analysis_type} analysis with {PROFILING_LIB}")
            
            self.update_progress(40, "Generating profile...")
            
            # Generate profile with timeout handling
            try:
                import signal
                
                def timeout_handler(signum, frame):
                    raise TimeoutError("Profile generation timed out")
                    
                # Set timeout (only on Unix systems)
                if hasattr(signal, 'SIGALRM'):
                    signal.signal(signal.SIGALRM, timeout_handler)
                    signal.alarm(self.timeout_var.get())
                
                profile = ProfileReport(df, **config)
                
                if hasattr(signal, 'SIGALRM'):
                    signal.alarm(0)  # Cancel timeout
                    
            except TimeoutError:
                raise Exception(f"Analysis timed out after {self.timeout_var.get()} seconds")
            except Exception as profile_error:
                self.log(f"⚠️ Error with full configuration: {str(profile_error)}")
                self.log("🔄 Retrying with minimal configuration...")
                
                minimal_config = {
                    "title": self.title_var.get(),
                    "minimal": True
                }
                profile = ProfileReport(df, **minimal_config)
                
            self.update_progress(80, "Saving report...")
            
            # Save report
            if self.output_format.get() == "html":
                profile.to_file(output_path)
            else:
                # Save as JSON
                json_data = profile.to_json()
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(json_data)
                    
            self.update_progress(90, "Finalizing...")
            
            # Save configuration if requested
            if self.include_config.get():
                config_path = Path(output_path).with_suffix('.config.json')
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=2, default=str)
                self.log(f"💾 Configuration saved: {config_path.name}")
                
            # Compress if requested
            if self.compress_output.get():
                self.compress_output_file(output_path)
                
            self.update_progress(100, "Complete!")
            
            # Show results
            file_size = os.path.getsize(output_path) / 1024 / 1024
            self.log(f"✅ Report generated successfully!")
            self.log(f"📁 Output: {Path(output_path).name}")
            self.log(f"💾 Size: {file_size:.1f} MB")
            
            # Update recent files
            self.add_to_recent_files(input_path)
            
            # Ask to open report
            if messagebox.askyesno("Success! 🎉", 
                                 f"Report generated successfully!\n\nFile: {Path(output_path).name}\nSize: {file_size:.1f} MB\n\nOpen the report now?"):
                self.open_file(output_path)
                
        except Exception as e:
            error_msg = f"❌ Error: {str(e)}"
            self.log(error_msg)
            self.update_progress(0, "Error occurred")
            messagebox.showerror("Error", str(e))
            
        finally:
            self.reset_progress()
            
    def compress_output_file(self, file_path):
        """Compress output file to ZIP"""
        import zipfile
        
        zip_path = Path(file_path).with_suffix(Path(file_path).suffix + '.zip')
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(file_path, Path(file_path).name)
            
            # Add config file if exists
            config_path = Path(file_path).with_suffix('.config.json')
            if config_path.exists():
                zipf.write(config_path, config_path.name)
                
        self.log(f"🗜️ Compressed to: {zip_path.name}")
        
    def update_progress(self, value, message):
        """Update progress bar and status"""
        self.progress_var.set(value)
        self.current_step.set(message)
        self.main_status_label.config(text=message)
        self.root.update_idletasks()
        
    def reset_progress(self):
        """Reset progress indicators"""
        self.progress_var.set(0)
        self.current_step.set("Ready")
        self.main_status_label.config(text="Ready to analyze data")
        
    def generate_report(self):
        """Start report generation"""
        if not YDATA_AVAILABLE:
            messagebox.showerror("Error", "Profiling library not available!\n\nPlease install:\npip install ydata-profiling")
            return
            
        if not self.input_file.get():
            messagebox.showwarning("Warning", "Please select an input file")
            return
            
        if not self.output_file.get():
            messagebox.showwarning("Warning", "Please specify output location")
            return
            
        # Start generation in thread
        self.generation_thread = threading.Thread(target=self.generate_report_thread)
        self.generation_thread.daemon = True
        self.generation_thread.start()
        
    def cancel_operation(self):
        """Cancel current operation"""
        # Note: This is a simplified cancel - in production you'd want more robust cancellation
        self.reset_progress()
        self.log("❌ Operation cancelled by user")
        
    def save_configuration(self):
        """Save current configuration to file"""
        config = {
            "title": self.title_var.get(),
            "description": self.dataset_description.get(),
            "analysis_type": "minimal" if self.minimal_var.get() else "comprehensive",
            "theme": self.theme_var.get(),
            "color_scheme": self.color_scheme.get(),
            "output_format": self.output_format.get(),
            
            # Analysis options
            "correlations": self.correlations_var.get(),
            "missing_diagrams": self.missing_diagrams_var.get(),
            "duplicates": self.duplicates_var.get(),
            "interactions": self.interactions_var.get(),
            "samples": self.samples_var.get(),
            
            # Performance settings
            "sample_size": self.sample_size.get(),
            "timeout": self.timeout_var.get(),
            "memory_limit": self.memory_limit.get(),
            
            # Preprocessing
            "auto_clean": self.auto_clean.get(),
            "handle_missing": self.handle_missing.get(),
            "encoding_detect": self.encoding_detect.get(),
            
            # Export options
            "include_config": self.include_config.get(),
            "compress_output": self.compress_output.get()
        }
        
        filename = filedialog.asksaveasfilename(
            title="Save Configuration",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=2, ensure_ascii=False)
                self.log(f"💾 Configuration saved: {Path(filename).name}")
                messagebox.showinfo("Success", "Configuration saved successfully!")
            except Exception as e:
                self.log(f"❌ Error saving configuration: {str(e)}")
                messagebox.showerror("Error", f"Cannot save configuration: {str(e)}")
                
    def load_configuration(self):
        """Load configuration from file"""
        filename = filedialog.askopenfilename(
            title="Load Configuration",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                self.apply_configuration(config)
                self.log(f"📂 Configuration loaded: {Path(filename).name}")
                messagebox.showinfo("Success", "Configuration loaded successfully!")
                
            except Exception as e:
                self.log(f"❌ Error loading configuration: {str(e)}")
                messagebox.showerror("Error", f"Cannot load configuration: {str(e)}")
                
    def apply_configuration(self, config):
        """Apply loaded configuration"""
        # Basic settings
        self.title_var.set(config.get("title", "Data Analysis Report"))
        self.dataset_description.set(config.get("description", ""))
        self.theme_var.set(config.get("theme", "default"))
        self.color_scheme.set(config.get("color_scheme", "blue"))
        self.output_format.set(config.get("output_format", "html"))
        
        # Analysis type
        analysis_type = config.get("analysis_type", "comprehensive")
        if analysis_type == "minimal":
            self.minimal_var.set(True)
            self.explorative_var.set(False)
        else:
            self.minimal_var.set(False)
            self.explorative_var.set(True)
            
        # Analysis options
        self.correlations_var.set(config.get("correlations", True))
        self.missing_diagrams_var.set(config.get("missing_diagrams", True))
        self.duplicates_var.set(config.get("duplicates", True))
        self.interactions_var.set(config.get("interactions", False))
        self.samples_var.set(config.get("samples", True))
        
        # Performance settings
        self.sample_size.set(config.get("sample_size", 0))
        self.timeout_var.set(config.get("timeout", 600))
        self.memory_limit.set(config.get("memory_limit", 2048))
        
        # Preprocessing
        self.auto_clean.set(config.get("auto_clean", False))
        self.handle_missing.set(config.get("handle_missing", "skip"))
        self.encoding_detect.set(config.get("encoding_detect", True))
        
        # Export options
        self.include_config.set(config.get("include_config", True))
        self.compress_output.set(config.get("compress_output", False))
        
    def reset_to_defaults(self):
        """Reset all settings to defaults"""
        if messagebox.askyesno("Reset Settings", "Reset all settings to defaults?"):
            self.initialize_variables()
            self.log("🔄 Settings reset to defaults")
            
    def add_to_recent_files(self, file_path):
        """Add file to recent files list"""
        # Simple recent files management
        current_items = list(self.recent_listbox.get(0, tk.END))
        
        # Remove if already exists
        display_text = f"{Path(file_path).name} - {file_path}"
        if display_text in current_items:
            index = current_items.index(display_text)
            self.recent_listbox.delete(index)
            
        # Add to top
        self.recent_listbox.insert(0, display_text)
        
        # Keep only last 10 items
        if self.recent_listbox.size() > 10:
            self.recent_listbox.delete(10, tk.END)
            
    def load_recent_file(self, event):
        """Load selected recent file"""
        selection = self.recent_listbox.curselection()
        if selection:
            item = self.recent_listbox.get(selection[0])
            # Extract file path from display text
            file_path = item.split(" - ", 1)[1] if " - " in item else item
            
            if os.path.exists(file_path):
                self.input_file.set(file_path)
                self.update_file_info(file_path)
                self.auto_generate_output_path(file_path)
                self.log(f"📂 Loaded recent file: {Path(file_path).name}")
            else:
                messagebox.showwarning("File Not Found", f"File no longer exists:\n{file_path}")
                self.recent_listbox.delete(selection[0])
                
    def save_last_session(self):
        """Save last session data"""
        session_data = {
            "input_file": self.input_file.get(),
            "output_file": self.output_file.get(),
            "recent_files": list(self.recent_listbox.get(0, tk.END))
        }
        
        try:
            session_file = Path.home() / ".profiling_studio_session.json"
            with open(session_file, 'w') as f:
                json.dump(session_data, f)
        except:
            pass  # Silent fail for session save
            
    def load_last_session(self):
        """Load last session data"""
        try:
            session_file = Path.home() / ".profiling_studio_session.json"
            if session_file.exists():
                with open(session_file, 'r') as f:
                    session_data = json.load(f)
                    
                # Restore recent files
                for item in session_data.get("recent_files", []):
                    self.recent_listbox.insert(tk.END, item)
                    
        except:
            pass  # Silent fail for session load
            
    def open_file(self, file_path):
        """Open file with default application"""
        try:
            if sys.platform == "win32":
                os.startfile(file_path)
            elif sys.platform == "darwin":
                os.system(f"open '{file_path}'")
            else:
                os.system(f"xdg-open '{file_path}'")
        except Exception as e:
            self.log(f"❌ Cannot open file: {str(e)}")
            
    def save_log(self):
        """Save log to file"""
        filename = filedialog.asksaveasfilename(
            title="Save Log",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                log_content = self.log_text.get(1.0, tk.END)
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                self.log(f"💾 Log saved: {Path(filename).name}")
                messagebox.showinfo("Success", "Log saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Cannot save log: {str(e)}")
                
    def clear_log(self):
        """Clear the log"""
        if messagebox.askyesno("Clear Log", "Clear all log entries?"):
            self.log_text.delete(1.0, tk.END)
            self.log("🗑️ Log cleared")
            
    def log(self, message):
        """Add timestamped message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
        # Color coding for different message types
        if "❌" in message or "Error" in message:
            # Red for errors
            line_start = self.log_text.index("end-2c linestart")
            line_end = self.log_text.index("end-1c")
            self.log_text.tag_add("error", line_start, line_end)
            self.log_text.tag_config("error", foreground="#ef4444")
        elif "✅" in message or "Success" in message:
            # Green for success
            line_start = self.log_text.index("end-2c linestart")
            line_end = self.log_text.index("end-1c")
            self.log_text.tag_add("success", line_start, line_end)
            self.log_text.tag_config("success", foreground="#10b981")
        elif "⚠️" in message or "Warning" in message:
            # Orange for warnings
            line_start = self.log_text.index("end-2c linestart")
            line_end = self.log_text.index("end-1c")
            self.log_text.tag_add("warning", line_start, line_end)
            self.log_text.tag_config("warning", foreground="#f59e0b")
            
    def check_dependencies(self):
        """Check and report dependency status"""
        if not YDATA_AVAILABLE:
            self.status_indicator.config(text="●", foreground="#ef4444")
            self.lib_status.config(text="Library: Not Available")
            
            error_msg = """
🚨 Profiling Library Not Found!

Please install one of the following:

📦 Option 1 (Recommended):
pip install ydata-profiling

📦 Option 2 (Legacy):
pip install pandas-profiling

💡 Additional dependencies:
pip install pandas matplotlib seaborn

🔧 For full functionality:
pip install ydata-profiling[notebook]
            """.strip()
            
            messagebox.showerror("Dependencies Missing", error_msg)
            self.log("❌ Profiling library not available")
        else:
            self.status_indicator.config(text="●", foreground="#10b981")
            self.lib_status.config(text=f"Library: {PROFILING_LIB}")
            self.log(f"✅ Using {PROFILING_LIB} library")
            
        # Check additional optional dependencies
        optional_deps = {
            "matplotlib": "Required for visualizations",
            "seaborn": "Enhanced statistical plots",
            "plotly": "Interactive visualizations",
            "phik": "Advanced correlation analysis"
        }
        
        missing_optional = []
        for dep, description in optional_deps.items():
            try:
                __import__(dep)
            except ImportError:
                missing_optional.append(f"• {dep}: {description}")
                
        if missing_optional:
            self.log("⚠️ Optional dependencies missing:")
            for dep in missing_optional:
                self.log(f"  {dep}")
                
    def on_closing(self):
        """Handle application closing"""
        try:
            self.save_last_session()
        except:
            pass
            
        self.root.destroy()

def main():
    """Main application entry point"""
    # Create and configure root window
    root = tk.Tk()
    
    # Set window icon if available
    try:
        # You can add an icon file here
        # root.iconbitmap('icon.ico')
        pass
    except:
        pass
    
    # Create application
    app = DataProfilerGUI(root)
    
    # Handle window closing
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Add welcome message
    app.log("🚀 Welcome to Advanced Data Profiling Studio v2.0!")
    app.log("💡 Select a data file and configure your analysis settings")
    
    if YDATA_AVAILABLE:
        app.log("✨ Ready to generate professional data reports!")
    else:
        app.log("⚠️ Install profiling library to enable report generation")
    
    # Start the application
    try:
        root.mainloop()
    except KeyboardInterrupt:
        app.log("👋 Application closed by user")
        root.destroy()

if __name__ == "__main__":
    main()