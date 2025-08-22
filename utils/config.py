#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configuration management for the CDR Analyzer
"""

import configparser
import os
import logging
from pathlib import Path

class Config:
    def __init__(self, config_file='config/settings.ini'):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        
        # Ensure config directory exists
        config_dir = Path(config_file).parent
        config_dir.mkdir(exist_ok=True)
        
        # Load configuration
        self.load()
        
    def load(self):
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                self.config.read(self.config_file, encoding='utf-8')
                logging.info(f"Configuration loaded from {self.config_file}")
            else:
                # Create default configuration
                self.create_default_config()
                logging.info("Created default configuration")
                
        except Exception as e:
            logging.error(f"Error loading configuration: {e}")
            self.create_default_config()
    
    def create_default_config(self):
        """Create default configuration"""
        self.config['DEFAULT'] = {
            'version': '1.0.0',
            'created': '2025-01-01'
        }
        
        self.config['app'] = {
            'window_width': '1200',
            'window_height': '800',
            'theme': 'default',
            'auto_save_config': 'true'
        }
        
        self.config['paths'] = {
            'last_input_dir': str(Path.home()),
            'last_output_dir': str(Path.home()),
            'log_dir': 'logs'
        }
        
        self.config['processing'] = {
            'max_preview_rows': '100',
            'night_start_hour': '18',
            'night_end_hour': '6',
            'enable_backup': 'true',
            'validate_files': 'true'
        }
        
        self.config['logging'] = {
            'level': 'INFO',
            'max_log_size': '10MB',
            'backup_count': '5'
        }
        
        self.config['analysis'] = {
            'include_night_analysis': 'true',
            'include_roaming_analysis': 'true',
            'include_device_analysis': 'true',
            'max_results_per_sheet': '1000'
        }
        
        # Save default config
        self.save()
    
    def get(self, section, option, fallback=None):
        """Get configuration value"""
        try:
            return self.config.get(section, option, fallback=fallback)
        except Exception:
            return fallback
    
    def getint(self, section, option, fallback=None):
        """Get integer configuration value"""
        try:
            return self.config.getint(section, option, fallback=fallback)
        except Exception:
            return fallback
    
    def getboolean(self, section, option, fallback=None):
        """Get boolean configuration value"""
        try:
            return self.config.getboolean(section, option, fallback=fallback)
        except Exception:
            return fallback
    
    def set(self, section, option, value):
        """Set configuration value"""
        try:
            if not self.config.has_section(section):
                self.config.add_section(section)
            self.config.set(section, option, str(value))
            
            # Auto-save if enabled
            if self.getboolean('app', 'auto_save_config', fallback=True):
                self.save()
                
        except Exception as e:
            logging.error(f"Error setting config value {section}.{option}: {e}")
    
    def save(self):
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
            logging.debug(f"Configuration saved to {self.config_file}")
            
        except Exception as e:
            logging.error(f"Error saving configuration: {e}")
    
    def get_all_sections(self):
        """Get all configuration sections"""
        return self.config.sections()
    
    def get_section_items(self, section):
        """Get all items in a section"""
        try:
            return dict(self.config.items(section))
        except Exception:
            return {}
    
    def reset_to_defaults(self):
        """Reset configuration to defaults"""
        try:
            # Clear current config
            self.config.clear()
            
            # Create default config
            self.create_default_config()
            
            logging.info("Configuration reset to defaults")
            
        except Exception as e:
            logging.error(f"Error resetting configuration: {e}")
    
    def export_config(self, export_path):
        """Export configuration to file"""
        try:
            with open(export_path, 'w', encoding='utf-8') as f:
                self.config.write(f)
            logging.info(f"Configuration exported to {export_path}")
            return True
            
        except Exception as e:
            logging.error(f"Error exporting configuration: {e}")
            return False
    
    def import_config(self, import_path):
        """Import configuration from file"""
        try:
            if os.path.exists(import_path):
                self.config.read(import_path, encoding='utf-8')
                self.save()  # Save imported config
                logging.info(f"Configuration imported from {import_path}")
                return True
            else:
                logging.error(f"Import file not found: {import_path}")
                return False
                
        except Exception as e:
            logging.error(f"Error importing configuration: {e}")
            return False
