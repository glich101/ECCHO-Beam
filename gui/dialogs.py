#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dialog windows for the CDR Analyzer application
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
import pandas as pd
import logging

class ProgressDialog:
    def __init__(self, parent, cancel_callback):
        self.parent = parent
        self.cancel_callback = cancel_callback
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Processing...")
        self.dialog.geometry("400x150")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.geometry("+%d+%d" % (
            parent.winfo_rootx() + 50,
            parent.winfo_rooty() + 50
        ))
        
        # Prevent closing with X button
        self.dialog.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the progress dialog UI"""
        frame = ttk.Frame(self.dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Status label
        self.status_label = ttk.Label(frame, text="Starting...")
        self.status_label.pack(pady=(0, 10))
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            frame, 
            variable=self.progress_var, 
            maximum=100,
            length=300
        )
        self.progress_bar.pack(pady=(0, 10))
        
        # Percentage label
        self.percent_label = ttk.Label(frame, text="0%")
        self.percent_label.pack(pady=(0, 10))
        
        # Cancel button
        self.cancel_button = ttk.Button(frame, text="Cancel", command=self.on_cancel)
        self.cancel_button.pack()
        
    def update_progress(self, percent, message=""):
        """Update progress bar and message"""
        try:
            self.progress_var.set(percent)
            self.percent_label.config(text=f"{percent:.1f}%")
            
            if message:
                self.status_label.config(text=message)
                
            self.dialog.update_idletasks()
            
        except tk.TclError:
            # Dialog might be destroyed
            pass
    
    def on_cancel(self):
        """Handle cancel button click"""
        if self.cancel_callback:
            self.cancel_callback()
        self.status_label.config(text="Cancelling...")
        self.cancel_button.config(state='disabled')
    
    def close(self):
        """Close the dialog"""
        try:
            if self.dialog.winfo_exists():
                self.dialog.destroy()
        except tk.TclError:
            pass

class PreviewDialog:
    def __init__(self, parent, file_path):
        self.parent = parent
        self.file_path = file_path
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Data Preview - {file_path}")
        self.dialog.geometry("800x600")
        self.dialog.transient(parent)
        
        # Center the dialog
        self.dialog.geometry("+%d+%d" % (
            parent.winfo_rootx() + 50,
            parent.winfo_rooty() + 50
        ))
        
        self.setup_ui()
        self.load_preview()
        
    def setup_ui(self):
        """Setup the preview dialog UI"""
        # Main frame
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="File Information", padding="5")
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.info_text = ttk.Label(info_frame, text="Loading...")
        self.info_text.pack()
        
        # Preview frame
        preview_frame = ttk.LabelFrame(main_frame, text="Data Preview (First 100 rows)", padding="5")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Text widget with scrollbars
        self.text_widget = scrolledtext.ScrolledText(
            preview_frame,
            wrap=tk.NONE,
            width=80,
            height=20,
            font=('Courier', 9)
        )
        self.text_widget.pack(fill=tk.BOTH, expand=True)
        
        # Horizontal scrollbar
        h_scroll = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.text_widget.xview)
        self.text_widget.configure(xscrollcommand=h_scroll.set)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Close", command=self.dialog.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Refresh", command=self.load_preview).pack(side=tk.RIGHT, padx=(0, 10))
        
    def load_preview(self):
        """Load and display file preview"""
        try:
            self.info_text.config(text="Loading file...")
            self.text_widget.delete('1.0', tk.END)
            
            # Try to detect header start (reuse from processor)
            from core.cdr_processor import CDRProcessor
            processor = CDRProcessor()
            header_start = processor.detect_header_start(self.file_path)
            
            # Load first 1000 rows
            df = pd.read_csv(
                self.file_path,
                engine="python",
                sep=",",
                header=0,
                skiprows=header_start,
                nrows=1000,
                on_bad_lines="skip",
                dtype=str
            )
            
            # Update info
            file_size = pd.io.common.get_filepath_or_buffer(self.file_path)[0] # type: ignore
            try:
                import os
                size_mb = os.path.getsize(self.file_path) / (1024 * 1024)
                info_text = f"File: {self.file_path}\nSize: {size_mb:.2f} MB\nColumns: {len(df.columns)}\nRows shown: {len(df)}\nHeader starts at line: {header_start + 1}"
            except:
                info_text = f"File: {self.file_path}\nColumns: {len(df.columns)}\nRows shown: {len(df)}\nHeader starts at line: {header_start + 1}"
            
            self.info_text.config(text=info_text)
            
            # Display data
            preview_text = df.to_string(max_rows=100, max_cols=20, show_dimensions=True)
            self.text_widget.insert('1.0', preview_text)
            
            # Show column names separately
            columns_text = f"\n\nColumn Names ({len(df.columns)}):\n" + "\n".join([f"{i+1:2d}. {col}" for i, col in enumerate(df.columns)])
            self.text_widget.insert(tk.END, columns_text)
            
            logging.info(f"Preview loaded for {self.file_path}: {len(df)} rows, {len(df.columns)} columns")
            
        except Exception as e:
            error_msg = f"Error loading preview: {str(e)}"
            logging.error(error_msg)
            self.info_text.config(text="Error loading file")
            self.text_widget.insert('1.0', error_msg)

class ErrorDialog:
    def __init__(self, parent, title, message, details=""):
        self.parent = parent
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.geometry("+%d+%d" % (
            parent.winfo_rootx() + 50,
            parent.winfo_rooty() + 50
        ))
        
        self.setup_ui(message, details)
        
    def setup_ui(self, message, details):
        """Setup the error dialog UI"""
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Message
        message_label = ttk.Label(main_frame, text=message, wraplength=400)
        message_label.pack(pady=(0, 10))
        
        if details:
            # Details frame
            details_frame = ttk.LabelFrame(main_frame, text="Details", padding="5")
            details_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
            
            details_text = scrolledtext.ScrolledText(
                details_frame,
                wrap=tk.WORD,
                height=15,
                font=('Courier', 9)
            )
            details_text.pack(fill=tk.BOTH, expand=True)
            details_text.insert('1.0', details)
            details_text.config(state=tk.DISABLED)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="OK", command=self.dialog.destroy).pack(side=tk.RIGHT)
