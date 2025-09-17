 # Echo Beam — GUI Rewrites

 
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Theme Manager for Echo Beam - Light and Dark theme support (modernized)
"""

import tkinter as tk
from tkinter import ttk
import logging

class ThemeManager:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.current_theme = self.config.get('app', 'theme', fallback='light')

        # Accent color chosen to harmonize with your logo/background
        self.accent = '#7C5CFF'  # soft purple-blue

        # Define theme colors
        self.themes = {
            'light': {
                'bg': '#f7f8fb',
                'fg': '#111827',
                'accent': self.accent,
                'frame_bg': '#ffffff',
                'muted': '#6b7280',
                'button_bg': '#f3f4f6',
                'button_fg': '#111827',
                'entry_bg': '#ffffff',
                'listbox_bg': '#ffffff',
                'status_bg': '#eef2ff'
            },
            'dark': {
                'bg': '#0f1724',
                'fg': '#e6eef8',
                'accent': self.accent,
                'frame_bg': '#0b1220',
                'muted': '#94a3b8',
                'button_bg': '#111827',
                'button_fg': '#e6eef8',
                'entry_bg': '#071022',
                'listbox_bg': '#071022',
                'status_bg': '#061226'
            }
        }

        self.setup_ttk_styles()
        self.apply_theme(self.current_theme)

    def setup_ttk_styles(self):
        """Setup TTK styles for theming"""
        self.style = ttk.Style()

        # Base font
        default_font = ('Segoe UI', 10)
        header_font = ('Segoe UI', 12, 'bold')

        # Create a modern button style
        self.style.configure('EB.TButton',
                             font=default_font,
                             relief='flat',
                             padding=(8, 6))
        self.style.map('EB.TButton',
                       background=[('active', '!disabled', '#dedffb')],
                       foreground=[('disabled', '#888888')])

        # Primary (accent) button
        self.style.configure('EB.Primary.TButton',
                             font=default_font,
                             foreground='#ffffff',
                             padding=(8, 6))

        # Listbox-like frame
        self.style.configure('EB.TFrame', background='#ffffff')
        self.style.configure('EB.TLabelFrame', background='#ffffff')

        # Progressbar style
        try:
            self.style.configure('EB.Horizontal.TProgressbar',
                                 thickness=16)
        except Exception:
            pass

        # Header label style
        self.style.configure('EB.Header.TLabel', font=header_font)

    def apply_theme(self, theme_name):
        """Apply a theme to the application"""
        try:
            if theme_name not in self.themes:
                logging.warning(f"Unknown theme: {theme_name}, using light theme")
                theme_name = 'light'

            self.current_theme = theme_name
            theme_colors = self.themes[theme_name]

            # Apply to root window
            self.root.configure(bg=theme_colors['bg'])

            # Apply to ttk where possible
            # Configure primary button background dynamically using style map
            accent = theme_colors['accent']
            btn_bg = theme_colors['button_bg']

            # Primary button background uses accent color
            self.style.configure('EB.Primary.TButton', background=accent, foreground='#ffffff')
            self.style.map('EB.Primary.TButton',
                           background=[('active', accent), ('!disabled', accent)],
                           foreground=[('!disabled', "#ffffff8d")])
            
            self.style.configure('EB.TButton', background=btn_bg, foreground=theme_colors['button_fg'])

            # Recursively apply colors to widgets (best-effort)
            self._apply_theme_to_widgets(self.root, theme_colors)

            # Save theme preference
            self.config.set('app', 'theme', theme_name)

            logging.info(f"Applied {theme_name} theme")

        except Exception as e:
            logging.error(f"Error applying theme: {e}")

    def _apply_theme_to_widgets(self, widget, colors):
        """Recursively apply theme to widgets (best-effort)
        We avoid changing widget classes that may break behavior.
        """
        try:
            cls = widget.winfo_class()

            # Apply common widget attributes
            if cls in ('TFrame', 'Frame'):
                try:
                    widget.configure(bg=colors['frame_bg'])
                except Exception:
                    pass
            elif cls in ('TLabel', 'Label'):
                try:
                    widget.configure(bg=colors['frame_bg'], fg=colors['fg'])
                except Exception:
                    pass
            elif cls in ('Button', 'TButton'):
                try:
                    widget.configure(bg=colors['button_bg'], fg=colors['button_fg'], activebackground=colors['accent'])
                except Exception:
                    pass
            elif cls in ('Entry', 'TEntry'):
                try:
                    widget.configure(bg=colors['entry_bg'], fg=colors['fg'], insertbackground=colors['fg'])
                except Exception:
                    pass
            elif cls in ('Text',):
                try:
                    widget.configure(bg=colors['entry_bg'], fg=colors['fg'], insertbackground=colors['fg'])
                except Exception:
                    pass
            elif cls in ('Listbox',):
                try:
                    widget.configure(bg=colors['listbox_bg'], fg=colors['fg'], selectbackground=colors['accent'], selectforeground='#ffffff')
                except Exception:
                    pass

            for child in widget.winfo_children():
                self._apply_theme_to_widgets(child, colors)

        except Exception:
            # Some widgets will not accept configuration; skip silently
            pass

    def toggle_theme(self):
        """Toggle between light and dark themes"""
        new_theme = 'dark' if self.current_theme == 'light' else 'light'
        self.apply_theme(new_theme)
        return new_theme

    def get_current_theme(self):
        """Get the current theme name"""
        return self.current_theme

    def get_theme_colors(self, theme_name=None):
        """Get color scheme for a theme"""
        if theme_name is None:
            theme_name = self.current_theme
        return self.themes.get(theme_name, self.themes['light'])
