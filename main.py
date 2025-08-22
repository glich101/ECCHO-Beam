#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CDR Desktop Analyzer - Main Entry Point
Enhanced desktop application for CDR analysis with robust GUI and error handling
"""

import sys
import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import logging

# Add current directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.main_window import MainWindow
from utils.logger import setup_logger
from utils.config import Config

class CDRAnalyzerApp:
    def __init__(self):
        self.root = None
        self.main_window = None
        self.config = Config()
        
    def setup_logging(self):
        """Setup application logging"""
        log_level = self.config.get('logging', 'level', fallback='INFO')
        setup_logger(log_level)
        
    def handle_exception(self, exc_type, exc_value, exc_traceback):
        """Global exception handler"""
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
            
        logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
        
        if self.root:
            error_msg = f"An unexpected error occurred:\n{exc_type.__name__}: {str(exc_value)}"
            messagebox.showerror("Application Error", error_msg)
    
    def run(self):
        """Initialize and run the application"""
        try:
            # Setup logging
            self.setup_logging()
            logging.info("Starting CDR Desktop Analyzer")
            
            # Set global exception handler
            sys.excepthook = self.handle_exception
            
            # Create main window
            self.root = tk.Tk()
            self.root.title("CDR Desktop Analyzer")
            self.root.geometry("1200x800")
            self.root.minsize(800, 600)
            
            # Set application icon (if available)
            try:
                # Note: SVG not directly supported by tkinter, would need conversion
                # For now, using default system icon
                pass
            except Exception:
                pass
            
            # Initialize main window
            self.main_window = MainWindow(self.root, self.config)
            
            # Configure window closing
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
            # Start main loop
            self.root.mainloop()
            
        except Exception as e:
            logging.error(f"Failed to start application: {e}")
            if hasattr(self, 'root') and self.root:
                messagebox.showerror("Startup Error", f"Failed to start application:\n{str(e)}")
            sys.exit(1)
    
    def on_closing(self):
        """Handle application closing"""
        try:
            # Save configuration
            self.config.save()
            
            # Check for running operations
            if self.main_window and self.main_window.is_processing():
                if messagebox.askyesno("Exit", "Processing is in progress. Do you want to exit anyway?"):
                    self.main_window.cancel_processing()
                    self.root.destroy()
            else:
                self.root.destroy()
                
        except Exception as e:
            logging.error(f"Error during application shutdown: {e}")
            self.root.destroy()

def main():
    """Main entry point"""
    app = CDRAnalyzerApp()
    app.run()

if __name__ == "__main__":
    main()
