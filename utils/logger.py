#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Logging configuration for the CDR Analyzer
"""

import logging
import os
import sys
from datetime import datetime
from pathlib import Path

def setup_logger(level='INFO'):
    """Setup application logging"""
    
    # Create logs directory if it doesn't exist
    log_dir = Path('logs')
    log_dir.mkdir(exist_ok=True)
    
    # Create log filename with timestamp
    log_filename = f"cdr_analyzer_{datetime.now().strftime('%Y%m%d')}.log"
    log_path = log_dir / log_filename
    
    # Configure logging format
    log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'
    
    # Convert string level to logging constant
    numeric_level = getattr(logging, level.upper(), logging.INFO)
    
    # Configure root logger
    logging.basicConfig(
        level=numeric_level,
        format=log_format,
        datefmt=date_format,
        handlers=[
            # File handler for persistent logging
            logging.FileHandler(log_path, encoding='utf-8'),
            # Console handler for immediate feedback
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    # Get logger instance
    logger = logging.getLogger('CDRAnalyzer')
    
    # Log startup information
    logger.info("="*50)
    logger.info("CDR Desktop Analyzer - Application Started")
    logger.info(f"Log Level: {level}")
    logger.info(f"Log File: {log_path}")
    logger.info(f"Python Version: {sys.version}")
    logger.info("="*50)
    
    return logger

def get_logger(name=None):
    """Get a logger instance"""
    return logging.getLogger(name or 'CDRAnalyzer')

def log_exception(logger, message="An error occurred"):
    """Log exception with full traceback"""
    logger.exception(message)

def log_performance(logger, operation, start_time, end_time):
    """Log performance metrics"""
    duration = end_time - start_time
    logger.info(f"Performance - {operation}: {duration:.2f} seconds")

class PerformanceLogger:
    """Context manager for performance logging"""
    
    def __init__(self, operation_name, logger=None):
        self.operation_name = operation_name
        self.logger = logger or get_logger()
        self.start_time = None
        
    def __enter__(self):
        self.start_time = datetime.now()
        self.logger.info(f"Starting operation: {self.operation_name}")
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        end_time = datetime.now()
        duration = (end_time - self.start_time).total_seconds()
        
        if exc_type is None:
            self.logger.info(f"Completed operation: {self.operation_name} in {duration:.2f}s")
        else:
            self.logger.error(f"Failed operation: {self.operation_name} after {duration:.2f}s - {exc_val}")
        
        return False  # Don't suppress exceptions

class MemoryLogger:
    """Logger that keeps messages in memory"""
    
    def __init__(self, max_messages=1000):
        self.messages = []
        self.max_messages = max_messages
        
    def add_message(self, level, message):
        """Add a message to the memory log"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.messages.append({
            'timestamp': timestamp,
            'level': level,
            'message': message
        })
        
        # Keep only recent messages
        if len(self.messages) > self.max_messages:
            self.messages = self.messages[-self.max_messages:]
    
    def get_messages(self, level=None):
        """Get messages, optionally filtered by level"""
        if level is None:
            return self.messages
        return [msg for msg in self.messages if msg['level'] == level]
    
    def clear(self):
        """Clear all messages"""
        self.messages.clear()
    
    def to_string(self):
        """Convert messages to string format"""
        return '\n'.join([
            f"{msg['timestamp']} - {msg['level']} - {msg['message']}"
            for msg in self.messages
        ])

# Global memory logger instance
memory_logger = MemoryLogger()

def log_to_memory(level, message):
    """Log message to memory logger"""
    memory_logger.add_message(level, message)
