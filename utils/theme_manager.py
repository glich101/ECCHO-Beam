#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Theme Manager for CDR Analyzer - Light and Dark theme support
"""

import tkinter as tk
from tkinter import ttk
import logging

class ThemeManager:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.current_theme = self.config.get('app', 'theme', fallback='light')
        
        # Define theme colors
        self.themes = {
            'light': {
                'bg': '#ffffff',
                'fg': '#000000',
                'select_bg': '#0078d4',
                'select_fg': '#ffffff',
                'frame_bg': '#f0f0f0',
                'button_bg': '#e1e1e1',
                'button_fg': '#000000',
                'entry_bg': '#ffffff',
                'entry_fg': '#000000',
                'listbox_bg': '#ffffff',
                'listbox_fg': '#000000',
                'menu_bg': '#f0f0f0',
                'menu_fg': '#000000',
                'status_bg': '#e8e8e8',
                'status_fg': '#000000'
            },
            'dark': {
                'bg': '#2b2b2b',
                'fg': '#ffffff',
                'select_bg': '#404040',
                'select_fg': '#ffffff',
                'frame_bg': '#3c3c3c',
                'button_bg': '#404040',
                'button_fg': '#ffffff',
                'entry_bg': '#404040',
                'entry_fg': '#ffffff',
                'listbox_bg': '#333333',
                'listbox_fg': '#ffffff',
                'menu_bg': '#2b2b2b',
                'menu_fg': '#ffffff',
                'status_bg': '#1e1e1e',
                'status_fg': '#ffffff'
            }
        }
        
        self.setup_ttk_styles()
        self.apply_theme(self.current_theme)
    
    def setup_ttk_styles(self):
        """Setup TTK styles for theming"""
        self.style = ttk.Style()
        
        # Configure styles for light theme
        self.style.theme_create('light_theme', parent='alt', settings={
            'TLabel': {
                'configure': {'background': '#ffffff', 'foreground': '#000000'}
            },
            'TButton': {
                'configure': {'background': '#e1e1e1', 'foreground': '#000000'},
                'map': {
                    'background': [('active', '#d1d1d1')],
                    'foreground': [('active', '#000000')]
                }
            },
            'TFrame': {
                'configure': {'background': '#f0f0f0'}
            },
            'TLabelFrame': {
                'configure': {'background': '#f0f0f0', 'foreground': '#000000'}
            },
            'TEntry': {
                'configure': {'fieldbackground': '#ffffff', 'foreground': '#000000'}
            },
            'TScrollbar': {
                'configure': {'background': '#e1e1e1', 'troughcolor': '#f0f0f0'}
            }
        })
        
        # Configure styles for dark theme
        self.style.theme_create('dark_theme', parent='alt', settings={
            'TLabel': {
                'configure': {'background': '#2b2b2b', 'foreground': '#ffffff'}
            },
            'TButton': {
                'configure': {'background': '#404040', 'foreground': '#ffffff'},
                'map': {
                    'background': [('active', '#505050')],
                    'foreground': [('active', '#ffffff')]
                }
            },
            'TFrame': {
                'configure': {'background': '#3c3c3c'}
            },
            'TLabelFrame': {
                'configure': {'background': '#3c3c3c', 'foreground': '#ffffff'}
            },
            'TEntry': {
                'configure': {'fieldbackground': '#404040', 'foreground': '#ffffff'}
            },
            'TScrollbar': {
                'configure': {'background': '#404040', 'troughcolor': '#2b2b2b'}
            }
        })
    
    def apply_theme(self, theme_name):
        """Apply a theme to the application"""
        try:
            if theme_name not in self.themes:
                logging.warning(f"Unknown theme: {theme_name}, using light theme")
                theme_name = 'light'
            
            self.current_theme = theme_name
            theme_colors = self.themes[theme_name]
            
            # Apply TTK theme
            if theme_name == 'dark':
                self.style.theme_use('dark_theme')
            else:
                self.style.theme_use('light_theme')
            
            # Apply to root window
            self.root.configure(bg=theme_colors['bg'])
            
            # Apply to all existing widgets
            self._apply_theme_to_widgets(self.root, theme_colors)
            
            # Save theme preference
            self.config.set('app', 'theme', theme_name)
            
            logging.info(f"Applied {theme_name} theme")
            
        except Exception as e:
            logging.error(f"Error applying theme: {e}")
    
    def _apply_theme_to_widgets(self, widget, colors):
        """Recursively apply theme to all widgets"""
        try:
            widget_class = widget.winfo_class()
            
            # Apply theme based on widget type
            if widget_class == 'Toplevel':
                widget.configure(bg=colors['bg'])
            elif widget_class == 'Frame':
                widget.configure(bg=colors['frame_bg'])
            elif widget_class == 'Label':
                widget.configure(bg=colors['bg'], fg=colors['fg'])
            elif widget_class == 'Button':
                widget.configure(
                    bg=colors['button_bg'], 
                    fg=colors['button_fg'],
                    activebackground=colors['select_bg'],
                    activeforeground=colors['select_fg']
                )
            elif widget_class == 'Entry':
                widget.configure(
                    bg=colors['entry_bg'], 
                    fg=colors['entry_fg'],
                    insertbackground=colors['fg']
                )
            elif widget_class == 'Text':
                widget.configure(
                    bg=colors['entry_bg'], 
                    fg=colors['entry_fg'],
                    insertbackground=colors['fg']
                )
            elif widget_class == 'Listbox':
                widget.configure(
                    bg=colors['listbox_bg'], 
                    fg=colors['listbox_fg'],
                    selectbackground=colors['select_bg'],
                    selectforeground=colors['select_fg']
                )
            elif widget_class == 'Menu':
                widget.configure(
                    bg=colors['menu_bg'], 
                    fg=colors['menu_fg'],
                    activebackground=colors['select_bg'],
                    activeforeground=colors['select_fg']
                )
            elif widget_class == 'Scrollbar':
                widget.configure(
                    bg=colors['button_bg'],
                    troughcolor=colors['frame_bg'],
                    activebackground=colors['select_bg']
                )
            
            # Recursively apply to children
            for child in widget.winfo_children():
                self._apply_theme_to_widgets(child, colors)
                
        except Exception as e:
            # Some widgets might not support certain configurations
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