#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI components for the CDR Analyzer application
"""

import tkinter as tk
from tkinter import ttk, messagebox
import os
import logging

class FileListFrame(ttk.Frame):
    def __init__(self, parent, main_window):
        super().__init__(parent)
        self.main_window = main_window
        self.files = []
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the file list interface"""
        # File list with scrollbar
        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Listbox
        self.file_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            selectmode=tk.EXTENDED,
            height=15
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Buttons frame
        buttons_frame = ttk.Frame(self)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(
            buttons_frame, 
            text="Add Files", 
            command=self.main_window.add_files
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            buttons_frame, 
            text="Remove Selected", 
            command=self.remove_selected
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            buttons_frame, 
            text="Clear All", 
            command=self.clear_files
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            buttons_frame, 
            text="Preview", 
            command=self.preview_selected
        ).pack(side=tk.RIGHT)
        
        # Context menu
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Preview", command=self.preview_selected)
        self.context_menu.add_command(label="Remove", command=self.remove_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Select All", command=self.select_all)
        
        self.file_listbox.bind("<Button-3>", self.show_context_menu)
        self.file_listbox.bind("<Double-1>", self.preview_selected)
        
    def add_files(self, file_paths):
        """Add files to the list"""
        added_count = 0
        for file_path in file_paths:
            if file_path not in self.files:
                self.files.append(file_path)
                filename = os.path.basename(file_path)
                self.file_listbox.insert(tk.END, f"{filename} ({file_path})")
                added_count += 1
                
        # Select all files by default
        self.select_all()
        return added_count
    
    def remove_selected(self):
        """Remove selected files from the list"""
        try:
            selected_indices = self.file_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("Warning", "No files selected")
                return
                
            # Remove in reverse order to maintain indices
            for index in reversed(selected_indices):
                self.file_listbox.delete(index)
                del self.files[index]
                
            logging.info(f"Removed {len(selected_indices)} files from list")
            
            # Update process button state
            if len(self.files) == 0:
                self.main_window.control_frame.disable_process_button()
                
        except Exception as e:
            logging.error(f"Error removing files: {e}")
            messagebox.showerror("Error", f"Error removing files:\n{str(e)}")
    
    def clear_files(self):
        """Clear all files from the list"""
        self.files.clear()
        self.file_listbox.delete(0, tk.END)
        logging.info("Cleared all files from list")
    
    def select_all(self):
        """Select all files in the list"""
        self.file_listbox.select_set(0, tk.END)
    
    def preview_selected(self, event=None):
        """Preview selected file"""
        try:
            selected_indices = self.file_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("Warning", "No file selected")
                return
                
            # Preview first selected file
            file_path = self.files[selected_indices[0]]
            self.main_window.preview_data_file(file_path)
            
        except Exception as e:
            logging.error(f"Error previewing file: {e}")
            messagebox.showerror("Error", f"Error previewing file:\n{str(e)}")
    
    def show_context_menu(self, event):
        """Show context menu"""
        try:
            self.context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            logging.error(f"Error showing context menu: {e}")
    
    def get_selected_files(self):
        """Get list of selected file paths"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            # If none selected, return all files
            return self.files.copy()
        return [self.files[i] for i in selected_indices]
    
    def get_file_count(self):
        """Get total number of files"""
        return len(self.files)

class ControlFrame(ttk.Frame):
    def __init__(self, parent, main_window):
        super().__init__(parent)
        self.main_window = main_window
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the control panel"""
        # Process button
        self.process_button = ttk.Button(
            self,
            text="Process Files",
            command=self.main_window.process_files,
            state='disabled'
        )
        self.process_button.pack(fill=tk.X, pady=(0, 10))
        
        # Separator
        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, pady=(0, 10))
        
        # Analysis options frame
        options_frame = ttk.LabelFrame(self, text="Analysis Options", padding="5")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Checkboxes for analysis options (for future use)
        self.include_night_analysis = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Include Night/Day Analysis",
            variable=self.include_night_analysis
        ).pack(anchor=tk.W)
        
        self.include_roaming_analysis = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Include Roaming Analysis",
            variable=self.include_roaming_analysis
        ).pack(anchor=tk.W)
        
        self.include_device_analysis = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Include Device Analysis",
            variable=self.include_device_analysis
        ).pack(anchor=tk.W)
        
        # Separator
        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, pady=(10, 10))
        
        # Quick actions frame
        actions_frame = ttk.LabelFrame(self, text="Quick Actions", padding="5")
        actions_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(
            actions_frame,
            text="Validate Files",
            command=self.main_window.validate_files
        ).pack(fill=tk.X, pady=(0, 5))
        
        ttk.Button(
            actions_frame,
            text="Preview Data",
            command=self.main_window.preview_data
        ).pack(fill=tk.X, pady=(0, 5))
        
        # Separator
        ttk.Separator(self, orient='horizontal').pack(fill=tk.X, pady=(10, 10))
        
        # Settings frame
        settings_frame = ttk.LabelFrame(self, text="Settings", padding="5")
        settings_frame.pack(fill=tk.X)
        
        # Night time settings
        night_frame = ttk.Frame(settings_frame)
        night_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(night_frame, text="Night Hours:").pack(side=tk.LEFT)
        
        self.night_start = tk.StringVar(value="18")
        night_start_spin = ttk.Spinbox(
            night_frame,
            from_=0, to=23,
            textvariable=self.night_start,
            width=3
        )
        night_start_spin.pack(side=tk.LEFT, padx=(5, 2))
        
        ttk.Label(night_frame, text="to").pack(side=tk.LEFT, padx=(0, 2))
        
        self.night_end = tk.StringVar(value="6")
        night_end_spin = ttk.Spinbox(
            night_frame,
            from_=0, to=23,
            textvariable=self.night_end,
            width=3
        )
        night_end_spin.pack(side=tk.LEFT, padx=(2, 5))
        
    def enable_process_button(self):
        """Enable the process button"""
        self.process_button.config(state='normal')
        
    def disable_process_button(self):
        """Disable the process button"""
        self.process_button.config(state='disabled')
        
    def set_processing_state(self, is_processing):
        """Set UI state based on processing status"""
        state = 'disabled' if is_processing else 'normal'
        self.process_button.config(
            state=state,
            text="Processing..." if is_processing else "Process Files"
        )

class StatusFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the status panel"""
        # Status bar
        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT)
        self.status_label = ttk.Label(status_frame, text="Ready")
        self.status_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # Log viewer
        log_frame = ttk.LabelFrame(self, text="Application Log", padding="5")
        log_frame.pack(fill=tk.X)
        
        self.log_text = tk.Text(
            log_frame,
            height=6,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=('Courier', 9)
        )
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Log scrollbar
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scroll.set)
        
        # Clear log button
        ttk.Button(
            log_frame,
            text="Clear",
            command=self.clear_log
        ).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Setup log handler
        self.setup_log_handler()
        
    def set_status(self, status):
        """Set status message"""
        self.status_label.config(text=status)
        logging.info(f"Status: {status}")
        
    def setup_log_handler(self):
        """Setup custom log handler to display in GUI"""
        class GUILogHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
                
            def emit(self, record):
                try:
                    msg = self.format(record)
                    self.text_widget.config(state=tk.NORMAL)
                    self.text_widget.insert(tk.END, msg + '\n')
                    self.text_widget.see(tk.END)
                    self.text_widget.config(state=tk.DISABLED)
                except Exception:
                    pass
        
        # Add GUI handler to root logger
        gui_handler = GUILogHandler(self.log_text)
        gui_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                                                 datefmt='%H:%M:%S'))
        logging.getLogger().addHandler(gui_handler)
        
    def clear_log(self):
        """Clear the log display"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state=tk.DISABLED)
