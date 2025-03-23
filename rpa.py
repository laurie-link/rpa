import sys
import os
import json
import time
import random
import signal
import re
import urllib.parse
import shutil
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QLineEdit, 
                            QTextEdit, QFileDialog, QProgressBar, QMessageBox,
                            QCheckBox, QGroupBox, QTabWidget, QSplitter)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSettings
from PyQt5.QtGui import QIcon, QTextCursor
from playwright.sync_api import sync_playwright



class LogRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""

    def write(self, text):
        self.buffer += text
        if '\n' in self.buffer:
            lines = self.buffer.split('\n')
            self.buffer = lines[-1]
            for line in lines[:-1]:
                self.text_widget.append(line)
                self.text_widget.moveCursor(QTextCursor.End)
        
    def flush(self):
        if self.buffer:
            self.text_widget.append(self.buffer)
            self.text_widget.moveCursor(QTextCursor.End)
            self.buffer = ""


class RPAWorker(QThread):
    progress_updated = pyqtSignal(int, int)
    log_message = pyqtSignal(str)
    task_completed = pyqtSignal(str, bool)
    
    def __init__(self, urls, md_template_path, settings):
        super().__init__()
        self.urls = urls
        self.md_template_path = md_template_path
        self.settings = settings
        self.browser_instance = None
        self.abort_flag = False
        
    def run(self):
        total_urls = len(self.urls)
        for i, url in enumerate(self.urls):
            if self.abort_flag:
                self.log_message.emit("任务已中止")
                break
                
            self.log_message.emit(f"处理 URL {i+1}/{total_urls}: {url}")
            self.progress_updated.emit(i, total_urls)
            
            try:
                self.process_url(url)
                self.task_completed.emit(url, True)
            except Exception as e:
                self.log_message.emit(f"处理 {url} 时出错: {str(e)}")
                self.task_completed.emit(url, False)
                
        self.log_message.emit("所有任务完成!")
        
    def abort(self):
        self.abort_flag = True
        self.log_message.emit("正在中止任务...")
        
    def process_url(self, page_url):
        """处理单个URL的RPA任务"""
        # 获取页面名称
        page_name = self.extract_page_name(page_url)
        domain = self.extract_domain(page_url)
        
        self.log_message.emit(f"提取的页面名称: {page_name}")
        self.log_message.emit(f"提取的域名: {domain}")
        
        