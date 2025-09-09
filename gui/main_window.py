#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Main Window GUI - Enhanced desktop interface for CDR analysis
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import logging
from pathlib import Path
from PIL import Image, ImageTk  # Add this for image support

from core.cdr_processor import CDRProcessor
from core.excel_generator import ExcelGenerator
from gui.dialogs import ProgressDialog, PreviewDialog
from gui.components import FileListFrame, ControlFrame, StatusFrame
from utils.theme_manager import ThemeManager

class MainWindow:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.is_processing_flag = False
        self.current_thread = None
        self.processor = None
        self.generator = None
        
        # Initialize theme manager
        self.theme_manager = ThemeManager(root, config)
        
        self.setup_ui()
        self.setup_menu()
        
    def setup_ui(self):
        """Setup the main user interface"""
        # --- Add background image ---
        try:
            bg_image = Image.open("assets/space_bg.png")
            bg_photo = ImageTk.PhotoImage(bg_image)
            self.bg_label = tk.Label(self.root, image=bg_photo)
            self.bg_label.image = bg_photo  # Keep reference # type: ignore
            self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        except Exception as e:
            print("Background image not loaded:", e)
        # --- End background image ---

        # Create main frame with padding
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S)) # type: ignore
        
        # Configure grid weights for responsive layout
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        
        # Header frame with title and theme button
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky=(tk.W, tk.E))# type: ignore
        header_frame.columnconfigure(0, weight=1)

        # --- Add logo image ---
        try:
            logo_img = Image.open("assets/planet_logo.png").resize((48, 48))
            logo_photo = ImageTk.PhotoImage(logo_img)
            self.logo_label = tk.Label(header_frame, image=logo_photo, bg="#00000000")
            self.logo_label.image = logo_photo  # Keep reference # type: ignore
            self.logo_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            title_col = 1
        except Exception as e:
            title_col = 0
        # --- End logo image ---

        # Title
        title_label = ttk.Label(
            header_frame, 
            text="CDR Desktop Analyzer", 
            font=('Arial', 16, 'bold')
        )
        title_label.grid(row=0, column=title_col, sticky=tk.W)
        
        # Theme toggle button
        self.theme_button = ttk.Button(
            header_frame,
            text="🌙 Dark",
            command=self.toggle_theme,
            width=8
        )
        self.theme_button.grid(row=0, column=1, sticky=tk.E)
        
        # Update theme button text based on current theme
        self.update_theme_button_text()
        
        # Left panel - Controls
        left_frame = ttk.LabelFrame(self.main_frame, text="Controls", padding="10")
        left_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10)) # type: ignore
        
        self.control_frame = ControlFrame(left_frame, self)
        self.control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))# type: ignore
        
        # Right panel - File list and preview
        right_frame = ttk.LabelFrame(self.main_frame, text="Files & Preview", padding="10")
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))# type: ignore
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        
        self.file_list_frame = FileListFrame(right_frame, self)
        self.file_list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))# type: ignore
        # type: ignore
        # Bottom panel - Status and logs
        bottom_frame = ttk.LabelFrame(self.main_frame, text="Status & Logs", padding="10")
        bottom_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))# type: ignore
        bottom_frame.columnconfigure(0, weight=1)
        
        self.status_frame = StatusFrame(bottom_frame)
        self.status_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))# type: ignore
        
    def setup_menu(self):
        """Setup application menu"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Add CSV Files", command=self.add_files)
        file_menu.add_separator()
        file_menu.add_command(label="Clear All Files", command=self.clear_files)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Preview Data", command=self.preview_data)
        tools_menu.add_command(label="Validate Files", command=self.validate_files)
        tools_menu.add_separator()
        tools_menu.add_command(label="Toggle Theme", command=self.toggle_theme)
        
        # View menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Light Theme", command=lambda: self.set_theme('light'))
        view_menu.add_command(label="Dark Theme", command=lambda: self.set_theme('dark'))
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="User Guide", command=self.show_help)
        
    def add_files(self):
        """Add CSV files to the processing list"""
        try:
            filetypes = [
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
            
            files = filedialog.askopenfilenames(
                title="Select CSV files",
                filetypes=filetypes,
                initialdir=self.config.get('paths', 'last_input_dir', fallback=str(Path.home()))
            )
            
            if files:
                # Update last directory
                self.config.set('paths', 'last_input_dir', str(Path(files[0]).parent))
                
                # Add files to list
                added_count = self.file_list_frame.add_files(files)
                self.status_frame.set_status(f"Added {added_count} files")
                
                # Enable process button if files are available
                if self.file_list_frame.get_file_count() > 0:
                    self.control_frame.enable_process_button()
                    
        except Exception as e:
            logging.error(f"Error adding files: {e}")
            messagebox.showerror("Error", f"Error adding files:\n{str(e)}")
    
    def clear_files(self):
        """Clear all files from the list"""
        if self.is_processing_flag:
            messagebox.showwarning("Warning", "Cannot clear files while processing is in progress")
            return
            
        self.file_list_frame.clear_files()
        self.control_frame.disable_process_button()
        self.status_frame.set_status("All files cleared")
    
    def validate_files(self):
        """Validate selected files"""
        try:
            files = self.file_list_frame.get_selected_files()
            if not files:
                messagebox.showwarning("Warning", "No files selected")
                return
                
            self.status_frame.set_status("Validating files...")
            
            # Basic validation
            invalid_files = []
            for file_path in files:
                if not os.path.exists(file_path):
                    invalid_files.append(f"{file_path} - File not found")
                elif os.path.getsize(file_path) == 0:
                    invalid_files.append(f"{file_path} - Empty file")
                    
            if invalid_files:
                error_msg = "Invalid files found:\n" + "\n".join(invalid_files)
                messagebox.showerror("Validation Error", error_msg)
                self.status_frame.set_status("Validation failed")
            else:
                messagebox.showinfo("Validation", "All selected files are valid")
                self.status_frame.set_status("Validation successful")
                
        except Exception as e:
            logging.error(f"Error validating files: {e}")
            messagebox.showerror("Error", f"Error during validation:\n{str(e)}")
    
    def preview_data(self):
        """Show data preview dialog"""
        try:
            files = self.file_list_frame.get_selected_files()
            if not files:
                messagebox.showwarning("Warning", "No files selected")
                return
            
            # Show preview dialog
            dialog = PreviewDialog(self.root, files[0])  # Preview first file
            
        except Exception as e:
            logging.error(f"Error showing preview: {e}")
            messagebox.showerror("Error", f"Error showing preview:\n{str(e)}")
    
    def process_files(self):
        """Start file processing in background thread"""
        try:
            if self.is_processing_flag:
                messagebox.showwarning("Warning", "Processing is already in progress")
                return
                
            files = self.file_list_frame.get_selected_files()
            if not files:
                messagebox.showwarning("Warning", "No files selected for processing")
                return
            
            # Get output path
            output_path = filedialog.asksaveasfilename(
                title="Save Excel file as",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                defaultextension=".xlsx",
                initialdir=self.config.get('paths', 'last_output_dir', fallback=str(Path.home()))
            )
            
            if not output_path:
                return
                
            # Update last directory
            self.config.set('paths', 'last_output_dir', str(Path(output_path).parent))
            
            # Start processing
            self.start_processing(files, output_path)
            
        except Exception as e:
            logging.error(f"Error starting processing: {e}")
            messagebox.showerror("Error", f"Error starting processing:\n{str(e)}")
    
    def start_processing(self, files, output_path):
        """Start processing in background thread"""
        self.is_processing_flag = True
        self.control_frame.set_processing_state(True)
        self.status_frame.set_status("Processing started...")
        
        # Create progress dialog
        self.progress_dialog = ProgressDialog(self.root, self.cancel_processing)
        
        # Start processing thread
        self.current_thread = threading.Thread(
            target=self.processing_worker,
            args=(files, output_path),
            daemon=True
        )
        self.current_thread.start()
    
    def processing_worker(self, files, output_path):
        """Worker function for file processing"""
        try:
            # Initialize processor and generator
            self.processor = CDRProcessor(progress_callback=self.update_progress)
            self.generator = ExcelGenerator(progress_callback=self.update_progress)
            
            # Process files
            self.update_progress(0, "Starting data processing...")
            combined_df = self.processor.process_files(files)
            
            if self.processor.cancel_flag:
                self.processing_completed(False, "Processing cancelled by user")
                return
            
            # Generate Excel file
            self.update_progress(0, "Generating Excel file...")
            self.generator.generate_excel_file(combined_df, output_path)
            
            if self.generator.cancel_flag:
                self.processing_completed(False, "Excel generation cancelled by user")
                return
            
            # Success
            self.processing_completed(True, f"Processing completed successfully!\nOutput saved to: {output_path}")
            
        except Exception as e:
            error_msg = f"Error during processing: {str(e)}"
            logging.error(error_msg)
            self.processing_completed(False, error_msg)
    
    def update_progress(self, percent, message=""):
        """Update progress dialog from worker thread"""
        if hasattr(self, 'progress_dialog') and self.progress_dialog:
            self.root.after(0, lambda: self.progress_dialog.update_progress(percent, message))
            
        # Also update status frame
        self.root.after(0, lambda: self.status_frame.set_status(message if message else f"Processing... {percent}%"))
    
    def cancel_processing(self):
        """Cancel current processing"""
        try:
            if self.processor:
                self.processor.set_cancel_flag()
            if self.generator:
                self.generator.set_cancel_flag()
                
            self.status_frame.set_status("Cancelling processing...")
            
        except Exception as e:
            logging.error(f"Error cancelling processing: {e}")
    
    def processing_completed(self, success, message):
        """Handle processing completion"""
        def complete():
            self.is_processing_flag = False
            self.control_frame.set_processing_state(False)
            
            if hasattr(self, 'progress_dialog') and self.progress_dialog:
                self.progress_dialog.close()
                
            if success:
                messagebox.showinfo("Success", message)
                self.status_frame.set_status("Processing completed successfully")
            else:
                messagebox.showerror("Error", message)
                self.status_frame.set_status("Processing failed")
            
            # Reset processors
            self.processor = None
            self.generator = None
            self.current_thread = None
        
        self.root.after(0, complete)
    
    def show_about(self):
        """Show about dialog"""
        about_text = """CDR Desktop Analyzer v1.0

A robust desktop application for analyzing Call Detail Records (CDR) from CSV files.

Features:
• Multi-file CDR import and processing
• 16 different analysis sheets
• Robust error handling and validation
• Progress tracking and cancellation
• Data preview capabilities

Developed with Python and tkinter
        """
        messagebox.showinfo("About", about_text)
    
    def show_help(self):
        """Show help/user guide"""
        help_text = """CDR Desktop Analyzer - User Guide

1. Adding Files:
   - Click 'Add Files' or use File > Add CSV Files
   - Select one or more CSV files containing CDR data
   - Files will be validated automatically

2. Processing:
   - Select files from the list (or leave all selected)
   - Click 'Process Files'
   - Choose output location for Excel file
   - Monitor progress in the dialog

3. Analysis Sheets:
   The generated Excel file contains 16 sheets:
   - Mapping: Basic call/SMS records
   - Summary: Aggregated statistics
   - MaxCalls/MaxDuration: Top contacts by activity
   - Night/Day analysis for different time periods
   - Location and roaming analysis
   - Device (IMEI/IMSI) analysis
   - International calls analysis

4. Tips:
   - Use 'Preview Data' to check file format
   - 'Validate Files' checks for common issues
   - Processing can be cancelled anytime
   - Large files may take several minutes

For support, check the application logs in the Status panel.
        """
        
        # Create help window
        help_window = tk.Toplevel(self.root)
        help_window.title("User Guide")
        help_window.geometry("600x500")
        
        text_widget = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)
    
    def is_processing(self):
        """Check if processing is currently active"""
        return self.is_processing_flag
    
    def toggle_theme(self):
        """Toggle between light and dark themes"""
        try:
            new_theme = self.theme_manager.toggle_theme()
            self.update_theme_button_text()
            logging.info(f"Switched to {new_theme} theme")
            
        except Exception as e:
            logging.error(f"Error toggling theme: {e}")
    
    def set_theme(self, theme_name):
        """Set specific theme"""
        try:
            self.theme_manager.apply_theme(theme_name)
            self.update_theme_button_text()
            logging.info(f"Applied {theme_name} theme")
            
        except Exception as e:
            logging.error(f"Error setting theme: {e}")
    
    def update_theme_button_text(self):
        """Update theme button text based on current theme"""
        try:
            current_theme = self.theme_manager.get_current_theme()
            if current_theme == 'light':
                self.theme_button.config(text="🌙 Dark")
            else:
                self.theme_button.config(text="☀️ Light")
        except Exception as e:
            logging.error(f"Error updating theme button: {e}")
