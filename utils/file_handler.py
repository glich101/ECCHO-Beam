#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
File handling utilities for the CDR Analyzer
"""

import os
import csv
import logging
from pathlib import Path
import pandas as pd

class FileHandler:
    @staticmethod
    def validate_csv_file(file_path):
        """Validate CSV file format and structure"""
        errors = []
        warnings = []
        
        try:
            # Check if file exists
            if not os.path.exists(file_path):
                errors.append("File does not exist")
                return errors, warnings
            
            # Check file size
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                errors.append("File is empty")
                return errors, warnings
            
            if file_size > 500 * 1024 * 1024:  # 500MB
                warnings.append("Large file size may cause slow processing")
            
            # Check file extension
            if not file_path.lower().endswith('.csv'):
                warnings.append("File does not have .csv extension")
            
            # Try to read first few lines
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    first_lines = [f.readline() for _ in range(10)]
                
                # Check for common CSV patterns
                has_comma = any(',' in line for line in first_lines)
                if not has_comma:
                    errors.append("File does not appear to be comma-separated")
                
                # Check for potential CDR column headers
                content = ''.join(first_lines).lower()
                cdr_indicators = [
                    'calling party', 'called party', 'a party', 'b party',
                    'call date', 'call time', 'duration', 'imei', 'cell'
                ]
                
                if not any(indicator in content for indicator in cdr_indicators):
                    warnings.append("File may not contain CDR data")
                
            except Exception as e:
                errors.append(f"Error reading file: {str(e)}")
            
        except Exception as e:
            errors.append(f"Validation error: {str(e)}")
        
        return errors, warnings
    
    @staticmethod
    def get_file_info(file_path):
        """Get basic information about a file"""
        try:
            stat = os.stat(file_path)
            return {
                'path': file_path,
                'name': os.path.basename(file_path),
                'size': stat.st_size,
                'size_mb': stat.st_size / (1024 * 1024),
                'modified': stat.st_mtime,
                'readable': os.access(file_path, os.R_OK)
            }
        except Exception as e:
            logging.error(f"Error getting file info for {file_path}: {e}")
            return None
    
    @staticmethod
    def detect_encoding(file_path, sample_size=1024):
        """Detect file encoding"""
        try:
            import chardet
            
            with open(file_path, 'rb') as f:
                sample = f.read(sample_size)
                result = chardet.detect(sample)
                return result.get('encoding', 'utf-8')
                
        except ImportError:
            # chardet not available, assume utf-8
            return 'utf-8'
        except Exception as e:
            logging.warning(f"Error detecting encoding for {file_path}: {e}")
            return 'utf-8'
    
    @staticmethod
    def safe_create_directory(dir_path):
        """Safely create directory if it doesn't exist"""
        try:
            Path(dir_path).mkdir(parents=True, exist_ok=True)
            return True
        except Exception as e:
            logging.error(f"Error creating directory {dir_path}: {e}")
            return False
    
    @staticmethod
    def get_safe_filename(filename):
        """Get a safe filename by removing/replacing invalid characters"""
        import re
        
        # Remove or replace invalid characters
        safe_name = re.sub(r'[<>:"/\\|?*]', '_', filename)
        
        # Remove multiple underscores
        safe_name = re.sub(r'_{2,}', '_', safe_name)
        
        # Trim and ensure not empty
        safe_name = safe_name.strip('_').strip()
        if not safe_name:
            safe_name = 'output'
        
        return safe_name
    
    @staticmethod
    def backup_file(file_path):
        """Create a backup of a file"""
        try:
            if not os.path.exists(file_path):
                return False
                
            backup_path = file_path + '.backup'
            
            # If backup already exists, create numbered backup
            counter = 1
            while os.path.exists(backup_path):
                backup_path = f"{file_path}.backup{counter}"
                counter += 1
            
            import shutil
            shutil.copy2(file_path, backup_path)
            logging.info(f"Created backup: {backup_path}")
            return backup_path
            
        except Exception as e:
            logging.error(f"Error creating backup for {file_path}: {e}")
            return False
