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
import semrush_module


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
    
    def __init__(self, urls, settings):
        super().__init__()
        self.urls = urls
        self.settings = settings
        self.browser_instance = None
        self.abort_flag = False
        self.is_original_mode = self.settings.value("original_article_mode", "false") == "true"
        
    def run(self):
        total_urls = len(self.urls)
        for i, url in enumerate(self.urls):
            if self.abort_flag:
                self.log_message.emit("任务已中止")
                break
                
            if self.is_original_mode:
                self.log_message.emit(f"处理关键词 {i+1}/{total_urls}: {url}")
            else:
                self.log_message.emit(f"处理URL {i+1}/{total_urls}: {url}")
            self.progress_updated.emit(i, total_urls)
            
            try:
                self.process_url(url)
                self.task_completed.emit(url, True)
            except Exception as e:
                if self.is_original_mode:
                    self.log_message.emit(f"处理关键词 {url} 时出错: {str(e)}")
                else:
                    self.log_message.emit(f"处理 {url} 时出错: {str(e)}")
                self.task_completed.emit(url, False)
                
        self.log_message.emit("所有任务完成!")
        
    def abort(self):
        self.abort_flag = True
        self.log_message.emit("正在中止任务...")
        
    def process_url(self, page_url):
        """处理单个URL或关键词的RPA任务"""
        # 判断是否为原创文章模式
        if self.is_original_mode:
            # 在原创文章模式下，输入的是关键词而不是URL
            keyword = page_url
            page_name = keyword.strip().replace(" ", "-").lower()
            
            self.log_message.emit(f"原创文章模式: 处理关键词 \"{keyword}\"")
            self.log_message.emit(f"使用的文件名: {page_name}")
            
            # 确保截图目录存在
            screenshot_dir = self.settings.value("screenshot_dir", "screenshots")
            if not os.path.exists(screenshot_dir):
                os.makedirs(screenshot_dir)
                
            with sync_playwright() as p:
                try:
                    # 处理Google搜索（原创文章模式下，只处理SERP和SEMrush）
                    if self.settings.value("scrape_serp", "true") == "true":
                        self.log_message.emit(f"开始处理Google搜索数据，搜索查询: {keyword}")
                        self.process_google_search_incognito(p, keyword, page_name, screenshot_dir)
                    else:
                        self.log_message.emit("已跳过SERP数据抓取（根据设置）")
                    
                    # 处理SEMrush
                    if self.settings.value("scrape_semrush", "true") == "true":
                        self.log_message.emit(f"开始处理SEMrush关键词数据")
                        # 创建一个新的浏览器和页面
                        browser = self.launch_browser(p)
                        page = browser.new_page()
                        self.setup_page(page)
                        semrush_module.process_semrush(self.log_message.emit, page, page_name, screenshot_dir)
                        # 关闭浏览器
                        browser.close()
                    else:
                        self.log_message.emit("已跳过SEMrush数据抓取（根据设置）")
                    
                    # 处理GA
                    if self.settings.value("scrape_ga", "true") == "true":
                        self.process_ga(page, ga_url, page_name, ga_screenshot_path, screenshot_dir)
                    else:
                        self.log_message.emit("已跳过GA数据抓取（根据设置）")
                    
                    # 关闭带配置的浏览器
                    browser.close()
                    
                    # 处理Google搜索（无痕模式）
                    if self.settings.value("scrape_serp", "true") == "true":
                        search_query = page_name.replace("-", " ")
                        self.log_message.emit(f"开始处理Google搜索数据，搜索查询: {search_query}")
                        self.process_google_search_incognito(p, search_query, page_name, screenshot_dir)
                    else:
                        self.log_message.emit("已跳过SERP数据抓取（根据设置）")
                except Exception as e:
                    self.log_message.emit(f"执行RPA时出错: {str(e)}")
                    try:
                        if self.browser_instance:
                            self.browser_instance.close()
                    except:
                        pass
                    raise e
        else:
            # 原来的URL处理逻辑
            # 获取页面名称
            page_name = self.extract_page_name(page_url)
            domain = self.extract_domain(page_url)
            
            self.log_message.emit(f"提取的页面名称: {page_name}")
            self.log_message.emit(f"提取的域名: {domain}")
            
            # 确保截图目录存在
            screenshot_dir = self.settings.value("screenshot_dir", "screenshots")
            if not os.path.exists(screenshot_dir):
                os.makedirs(screenshot_dir)
                
            # 构建截图路径
            first_screenshot_path = os.path.join(screenshot_dir, f"gsc-{page_name}-chart1.png")
            second_screenshot_path = os.path.join(screenshot_dir, f"gsc-{page_name}-chart2.png")
            ga_screenshot_path = os.path.join(screenshot_dir, f"ga-{page_name}.png")
            
            # 构建GSC和GA URL
            gsc_url, ga_url = self.build_urls(page_url, domain, page_name)
            
            with sync_playwright() as p:
                try:
                    # 启动浏览器（使用用户配置文件，用于GSC和GA访问，因为需要登录状态）
                    browser = self.launch_browser(p)
                    
                    # 创建新页面
                    page = browser.new_page()
                    self.setup_page(page)
                    
                    # 处理GSC
                    if self.settings.value("scrape_gsc", "true") == "true":
                        self.process_gsc(page, gsc_url, page_name, first_screenshot_path, second_screenshot_path, screenshot_dir)
                    else:
                        self.log_message.emit("已跳过GSC数据抓取（根据设置）")
                    
                    # 处理GA
                    if self.settings.value("scrape_ga", "true") == "true":
                        self.process_ga(page, ga_url, page_name, ga_screenshot_path, screenshot_dir)
                    else:
                        self.log_message.emit("已跳过GA数据抓取（根据设置）")
                    
                    # 关闭带配置的浏览器
                    browser.close()
                    
                    # 处理Google搜索（无痕模式）
                    if self.settings.value("scrape_serp", "true") == "true":
                        search_query = page_name.replace("-", " ")
                        self.log_message.emit(f"开始处理Google搜索数据，搜索查询: {search_query}")
                        self.process_google_search_incognito(p, search_query, page_name, screenshot_dir)
                    else:
                        self.log_message.emit("已跳过SERP数据抓取（根据设置）")
                    
                    # 处理SEMrush
                    if self.settings.value("scrape_semrush", "true") == "true":
                        self.log_message.emit(f"开始处理SEMrush关键词数据")
                        # 创建一个新的浏览器和页面
                        browser = self.launch_browser(p)
                        page = browser.new_page()
                        self.setup_page(page)
                        # 处理SEMrush
                        semrush_module.process_semrush(self.log_message.emit, page, page_name, screenshot_dir)
                        # 关闭浏览器
                        browser.close()
                    else:
                        self.log_message.emit("已跳过SEMrush数据抓取（根据设置）")
                except Exception as e:
                    self.log_message.emit(f"执行RPA时出错: {str(e)}")
                    # 确保浏览器关闭
                    try:
                        if self.browser_instance:
                            self.browser_instance.close()
                    except:
                        pass
                    raise e
    
    def extract_page_name(self, url):
        """从URL中提取页面名称"""
        try:
            parsed_url = urllib.parse.urlparse(url)
            path = parsed_url.path
            
            if path.startswith('/'):
                path = path[1:]
            if path.endswith('.html'):
                path = path[:-5]
                
            page_name = path.split('/')[-1]
            return page_name
        except Exception as e:
            self.log_message.emit(f"URL解析错误: {str(e)}")
            return "unknown-page"
            
    def extract_domain(self, url):
        """从URL中提取域名"""
        try:
            parsed_url = urllib.parse.urlparse(url)
            return parsed_url.netloc
        except Exception as e:
            self.log_message.emit(f"域名解析错误: {str(e)}")
            return "example.com"
            
    def build_urls(self, page_url, domain, page_name):
        """构建GSC和GA的URL"""
        # 解析URL获取完整路径
        parsed_url = urllib.parse.urlparse(page_url)
        path = parsed_url.path
        
        # 构建GSC URL
        encoded_domain = urllib.parse.quote(f"https://{domain}/")
        encoded_page = urllib.parse.quote(page_url)
        gsc_url = f"https://search.google.com/u/0/search-console/performance/search-analytics?resource_id={encoded_domain}&metrics=CLICKS%2CIMPRESSIONS%2CPOSITION&breakdown=query&pli=1&page=*{encoded_page}&num_of_months=3"
        
        # 构建GA URL
        ga_url = f"https://analytics.google.com/analytics/web/?authuser=0#/p309178187/reports/explorer?params=_u..nav%3Dmaui%26_r.explorerCard..startRow%3D0%26_r.explorerCard..filterTerm%3D{page_name}%26_u.dateOption%3Dlast90Days%26_u.comparisonOption%3Ddisabled%26_r.explorerCard..columnFilters%3D%7B%22conversionEvent%22:%22wclick_download%22%7D&r=5958195737&ruid=landing-page,life-cycle,engagement&collectionId=5958209258"
        
        self.log_message.emit(f"构建的GSC URL: {gsc_url}")
        self.log_message.emit(f"构建的GA URL: {ga_url}")
        
        return gsc_url, ga_url
    
    def launch_browser(self, playwright):
        """启动浏览器"""
        self.log_message.emit("启动浏览器...")
        
        # 获取用户数据目录
        user_data_dir = self.settings.value("chrome_profile", r"C:\Users\34897\AppData\Local\Google\Chrome\User Data")
        
        # 随机选择用户代理
        user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0"
        ]
        user_agent = random.choice(user_agents)
        
        # 获取模式设置
        headless = self.settings.value("headless_mode", "false") == "true"
        invisible_browser = self.settings.value("invisible_browser", "true") == "true"
        
        # 增强反检测浏览器参数
        browser_args = [
            '--profile-directory=Default',
            '--disable-blink-features=AutomationControlled',
            '--no-sandbox',
            '--disable-extensions',
            '--disable-default-apps',
            '--disable-popup-blocking',
            '--start-maximized',
            f'--user-agent={user_agent}',
            # 增加以下参数来绕过检测
            '--disable-web-security',
            '--disable-features=IsolateOrigins,site-per-process',
            '--disable-site-isolation-trials',
            '--disable-blink-features',
            '--disable-device-orientation',
            '--disable-features=Translate',
            '--disable-infobars',
            '--ignore-certifcate-errors',
            '--ignore-certifcate-errors-spki-list',
            '--allow-running-insecure-content',
            '--disable-gpu'
        ]
        
        # 根据设置选择启动模式
        if headless:
            # 完全无头模式 - 可能被检测
            self.log_message.emit("使用完全无头模式 (可能被检测为机器人)")
            browser = playwright.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=True,
                viewport={'width': 1920, 'height': 1080},
                args=browser_args
            )
        elif invisible_browser:
            # 隐形浏览器模式 - 使用兼容方法实现
            self.log_message.emit("使用隐形浏览器模式 (有浏览器但不可见，降低被检测概率)")
            # 确保窗口被正确放置在屏幕外，通过JavaScript而不仅仅是启动参数
            invisible_args = browser_args + [
                '--window-size=1920,1080'
                # 移除 '--window-position' 启动参数，改为用JavaScript控制
            ]
            
            try:
                # 尝试使用带is_visible参数的方法（新版本Playwright）
                browser = playwright.chromium.launch_persistent_context(
                    user_data_dir=user_data_dir,
                    headless=False,
                    is_visible=False,  # 可能不被支持
                    reduce_motion="reduce",
                    viewport={'width': 1920, 'height': 1080},
                    args=invisible_args
                )
            except TypeError:
                # 如果is_visible不被支持，使用不带该参数的方法
                self.log_message.emit("当前Playwright版本不支持is_visible参数，使用备选方法")
                browser = playwright.chromium.launch_persistent_context(
                    user_data_dir=user_data_dir,
                    headless=False,
                    viewport={'width': 1920, 'height': 1080},
                    args=invisible_args
                )
            
            # 确保窗口在屏幕外（即使没有is_visible参数也能工作）
            try:
                # 等待一个页面加载
                time.sleep(1)
                # 安全处理页面访问
                pages = browser.pages
                if callable(pages):
                    pages = pages()
                
                if pages and len(pages) > 0:
                    # 使用JavaScript将窗口移动到屏幕外，但确保位置值有效
                    first_page = pages[0]
                    if hasattr(first_page, 'evaluate') and callable(first_page.evaluate):
                        first_page.evaluate("""
                            try {
                                // 尝试将窗口移到屏幕外
                                window.moveTo(-10000, -10000);
                                // 如果不成功，尝试另一种方法
                                if (window.screenX > -5000) {
                                    window.moveTo(-2000, -2000);
                                }
                                window.resizeTo(1920, 1080);
                            } catch (e) {
                                console.error("无法移动窗口", e);
                            }
                        """)
            except Exception as e:
                self.log_message.emit(f"移动窗口时出错: {str(e)}")
        else:
            # 常规有头模式
            self.log_message.emit("使用正常有头模式")
            browser = playwright.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                viewport={'width': 1920, 'height': 1080},
                args=browser_args
            )
        
        self.browser_instance = browser
        return browser
    
    def setup_page(self, page):
        """设置页面参数和反检测措施"""
        # 执行窗口最大化
        page.evaluate("""
            window.moveTo(0, 0);
            window.resizeTo(screen.width, screen.height);
        """)
        
        # 添加CDP会话进一步修改浏览器指纹
        client = page.context.new_cdp_session(page)
        
        # 通过CDP会话执行更多反检测
        self.log_message.emit("应用反检测措施...")
        
        # 修改WebRTC行为
        client.send("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                // 阻止WebRTC泄露真实IP
                const originalGetUserMedia = navigator.mediaDevices?.getUserMedia;
                if (originalGetUserMedia) {
                    navigator.mediaDevices.getUserMedia = function() {
                        return new Promise((resolve, reject) => {
                            reject(new DOMException('Permission denied', 'NotAllowedError'));
                        });
                    };
                }
                
                // 阻止WebRTC API
                if (RTCPeerConnection) {
                    RTCPeerConnection = function() {
                        throw new Error("WebRTC is disabled");
                    };
                    RTCPeerConnection.prototype = {};
                }
            """
        })
        
        # 模拟正常的Canvas指纹
        client.send("Page.addScriptToEvaluateOnNewDocument", {
            "source": """
                // 修改Canvas指纹
                const originalToDataURL = HTMLCanvasElement.prototype.toDataURL;
                const originalGetImageData = CanvasRenderingContext2D.prototype.getImageData;
                
                HTMLCanvasElement.prototype.toDataURL = function(type) {
                    if (this.width > 1 && this.height > 1) {
                        // 轻微修改Canvas数据来改变指纹
                        const context = this.getContext("2d");
                        const imageData = context.getImageData(0, 0, 1, 1);
                        // 随机改变一个像素
                        imageData.data[0] = imageData.data[0] < 255 ? imageData.data[0] + 1 : imageData.data[0] - 1;
                        context.putImageData(imageData, 0, 0);
                    }
                    return originalToDataURL.apply(this, arguments);
                };
                
                CanvasRenderingContext2D.prototype.getImageData = function() {
                    const imageData = originalGetImageData.apply(this, arguments);
                    // 略微修改ImageData
                    if (imageData && imageData.data && imageData.data.length > 10) {
                        const offset = Math.floor(Math.random() * (imageData.data.length - 10));
                        imageData.data[offset] = (imageData.data[offset] + 1) % 256;
                    }
                    return imageData;
                };
            """
        })
        
        # 应用一般反检测措施
        client.send('Page.addScriptToEvaluateOnNewDocument', {
            'source': '''
                // 覆盖navigator.webdriver
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => false,
                });
                
                // 覆盖window.navigator.chrome
                window.navigator.chrome = {
                    runtime: {},
                    app: {
                        InstallState: {
                            DISABLED: 'disabled',
                            INSTALLED: 'installed',
                            NOT_INSTALLED: 'not_installed'
                        },
                        RunningState: {
                            CANNOT_RUN: 'cannot_run',
                            READY_TO_RUN: 'ready_to_run',
                            RUNNING: 'running'
                        },
                        getDetails: function() {},
                        getIsInstalled: function() {},
                        installState: function() { 
                            return 'installed';
                        },
                        isInstalled: true,
                        runningState: function() {
                            return 'running';
                        }
                    }
                };
                
                // 覆盖window.chrome
                window.chrome = {
                    runtime: {
                        OnInstalledReason: {
                            CHROME_UPDATE: 'chrome_update',
                            INSTALL: 'install',
                            SHARED_MODULE_UPDATE: 'shared_module_update',
                            UPDATE: 'update'
                        },
                        OnRestartRequiredReason: {
                            APP_UPDATE: 'app_update',
                            OS_UPDATE: 'os_update',
                            PERIODIC: 'periodic'
                        },
                        PlatformArch: {
                            ARM: 'arm',
                            ARM64: 'arm64',
                            MIPS: 'mips',
                            MIPS64: 'mips64',
                            X86_32: 'x86-32',
                            X86_64: 'x86-64'
                        },
                        PlatformNaclArch: {
                            ARM: 'arm',
                            MIPS: 'mips',
                            MIPS64: 'mips64',
                            X86_32: 'x86-32',
                            X86_64: 'x86-64'
                        },
                        PlatformOs: {
                            ANDROID: 'android',
                            CROS: 'cros',
                            LINUX: 'linux',
                            MAC: 'mac',
                            OPENBSD: 'openbsd',
                            WIN: 'win'
                        },
                        RequestUpdateCheckStatus: {
                            NO_UPDATE: 'no_update',
                            THROTTLED: 'throttled',
                            UPDATE_AVAILABLE: 'update_available'
                        }
                    },
                    app: {
                        isInstalled: true
                    }
                };
                
                // 修改navigator.plugins
                const makePluginArray = () => {
                    const plugins = [
                        { name: 'Chrome PDF Plugin', filename: 'internal-pdf-viewer', description: 'Portable Document Format' },
                        { name: 'Chrome PDF Viewer', filename: 'mhjfbmdgcfjbbpaeojofohoefgiehjai', description: 'Portable Document Format' },
                        { name: 'Native Client', filename: 'internal-nacl-plugin', description: '' }
                    ];
                    
                    const pluginArray = plugins.map(plugin => {
                        const pluginObj = {};
                        Object.defineProperty(pluginObj, 'name', { value: plugin.name });
                        Object.defineProperty(pluginObj, 'filename', { value: plugin.filename });
                        Object.defineProperty(pluginObj, 'description', { value: plugin.description });
                        return pluginObj;
                    });
                    
                    return Object.create(PluginArray.prototype, {
                        length: { value: plugins.length },
                        item: { value: index => pluginArray[index] },
                        namedItem: { value: name => pluginArray.find(plugin => plugin.name === name) },
                        ...pluginArray.reduce((acc, plugin, index) => {
                            acc[index] = { value: plugin };
                            return acc;
                        }, {})
                    });
                };
                
                // 应用插件覆盖
                Object.defineProperty(navigator, 'plugins', {
                    get: () => makePluginArray(),
                });
                
                // 覆盖语言设置
                Object.defineProperty(navigator, 'languages', {
                    get: () => ['zh-CN', 'zh', 'en-US', 'en'],
                });
                
                // 模拟正常的硬件并发层级
                Object.defineProperty(navigator, 'hardwareConcurrency', {
                    get: () => 8,
                });
                
                // 模拟正常的设备内存
                Object.defineProperty(navigator, 'deviceMemory', {
                    get: () => 8,
                });
                
                // 修改连接信息
                Object.defineProperty(navigator, 'connection', {
                    get: () => ({
                        effectiveType: '4g',
                        rtt: 50,
                        downlink: 10.0,
                        saveData: false
                    }),
                });
                
                // 模拟Notification API
                Object.defineProperty(window, 'Notification', {
                    get: () => function(title, options) {
                        this.title = title;
                        this.options = options;
                        this.permission = 'granted';
                    }
                });
                
                // 修改屏幕尺寸信息
                Object.defineProperty(window, 'screen', {
                    get: () => ({
                        availHeight: 1040,
                        availLeft: 0,
                        availTop: 0,
                        availWidth: 1920,
                        colorDepth: 24,
                        height: 1080,
                        width: 1920,
                        pixelDepth: 24
                    })
                });
                
                // WebGL指纹修改
                const getParameter = WebGLRenderingContext.prototype.getParameter;
                WebGLRenderingContext.prototype.getParameter = function(parameter) {
                    // UNMASKED_VENDOR_WEBGL
                    if (parameter === 37445) {
                        return 'Google Inc. (NVIDIA)';
                    }
                    // UNMASKED_RENDERER_WEBGL
                    if (parameter === 37446) {
                        return 'ANGLE (NVIDIA, NVIDIA GeForce GTX 1070 Direct3D11 vs_5_0 ps_5_0, D3D11)';
                    }
                    return getParameter.apply(this, arguments);
                };
                
                // 封锁Automation检测
                const newProto = navigator.__proto__;
                delete newProto.webdriver;
                navigator.__proto__ = newProto;
                
                // 阻止特征检测的特定属性
                Object.defineProperty(navigator, 'permissions', {
                    get: () => {
                        return {
                            query: function() { 
                                return Promise.resolve({state: 'prompt'});
                            }
                        }
                    }
                });
            '''
        })
        
        # 修复CSS和布局问题
        page.evaluate("""
            // 强制重新计算布局
            document.body.style.width = '100vw';
            document.body.style.height = '100vh';
            document.body.style.overflow = 'auto';
            
            // 触发窗口大小调整事件
            window.dispatchEvent(new Event('resize'));
        """)
    
    def process_gsc(self, page, gsc_url, page_name, first_screenshot_path, second_screenshot_path, screenshot_dir):
        """处理GSC相关的任务"""
        self.log_message.emit("导航到Google Search Console...")
        page.goto(gsc_url, timeout=60000)
        
        # 检查是否需要登录
        if page.url.startswith("https://accounts.google.com/"):
            self.log_message.emit("检测到需要登录，开始自动登录流程...")
            self.handle_google_login(page, "crawlyjoe98@gmail.com", "Cccy1314ss")
            self.log_message.emit("登录完成，继续执行...")
        
        # 添加随机滚动
        for _ in range(random.randint(2, 4)):
            page.mouse.wheel(0, random.randint(100, 300))
            time.sleep(random.uniform(0.5, 1.5))
        
        # 截取第一个图表
        try:

            time.sleep(10)
            self.log_message.emit("等待10秒...")
            selector = "#yDmH0d > c-wiz.zQTmif.SSPGKf.eejsDc > c-wiz > div > div.OoO4Vb > div > div > div.VfPpkd-WsjYwc.VfPpkd-WsjYwc-OWXEXe-INsAgc.KC1dQ.Usd1Ac.AaN0Dd.YJ1SEc.pTyMIf > c-wiz"
            
            self.log_message.emit("定位第一个目标元素...")
            element = page.wait_for_selector(selector, timeout=90000)
            
            if element:
                self.log_message.emit("找到元素，正在截图...")
                element.screenshot(path=first_screenshot_path)
                self.log_message.emit(f"第一个截图已保存为: {first_screenshot_path}")
            else:
                raise Exception("未找到目标元素")
        except Exception as e:
            self.log_message.emit(f"定位元素时出错: {str(e)}")
            self.log_message.emit("尝试全页截图作为备选...")
            full_page_path = os.path.join(screenshot_dir, f"gsc-{page_name}-chart1-full.png")
            page.screenshot(path=full_page_path, full_page=True)
            self.log_message.emit(f"整页截图已保存为: {full_page_path}")
        
        # 点击并截取第二个图表
        try:
            self.log_message.emit("进行额外操作：点击指定元素...")
            click_selector = "#\\31  > div > c-wiz > div > div > div:nth-child(2) > div:nth-child(2) > div > table > thead > tr > th:nth-child(3) > span > button > span > svg"
            
            page.wait_for_selector(click_selector, state="visible", timeout=45000)
            self.log_message.emit(f"点击元素: {click_selector}")
            page.click(click_selector)
            
            time.sleep(random.uniform(0.5, 1.0))
            
            second_selector = "#yDmH0d > c-wiz.zQTmif.SSPGKf.eejsDc > c-wiz > div > div.OoO4Vb > div > div > div:nth-child(2) > div"
            
            self.log_message.emit(f"定位第二个目标元素: {second_selector}")
            second_element = page.wait_for_selector(second_selector, timeout=45000)
            
            if second_element:
                self.log_message.emit("找到第二个元素，正在截图...")
                second_element.screenshot(path=second_screenshot_path)
                self.log_message.emit(f"第二个截图已保存为: {second_screenshot_path}")
            else:
                self.log_message.emit("未找到第二个元素")
        except Exception as e:
            self.log_message.emit(f"执行额外操作时出错: {str(e)}")
            self.log_message.emit("尝试全页截图作为备选...")
            second_full_page_path = os.path.join(screenshot_dir, f"gsc-{page_name}-chart2-full.png")
            page.screenshot(path=second_full_page_path, full_page=True)
            self.log_message.emit(f"第二个全页截图已保存为: {second_full_page_path}")
        
        # 提取GSC前10个结果并更新MD文件
        self.extract_and_update_md(page, page_name)
    
    def extract_and_update_md(self, page, page_name):
        """提取GSC前10个查询并更新MD文件"""
        try:
            self.log_message.emit("提取GSC前10个结果...")
            
            # 使用一个更通用的选择器来获取表体
            tbody_selector = "#\\31  > div > c-wiz > div > div > div:nth-child(2) > div:nth-child(2) > div > table > tbody"
            
            # 直接等待表体加载
            page.wait_for_selector(tbody_selector, timeout=45000)
            
            # 获取前10个查询文本
            gsc_queries = []
            
            # 获取所有行，然后提取前10个
            rows = page.query_selector_all(f"{tbody_selector} > tr")
            self.log_message.emit(f"找到 {len(rows)} 行数据")
            
            # 确保我们最多只处理10行
            rows = rows[:10] if len(rows) > 10 else rows
            
            for i, row in enumerate(rows):
                try:
                    # 获取每行的第一个单元格（查询名称）
                    query_cell = row.query_selector("td.XgRaPc[data-label='QUERIES'] span span")
                    if query_cell:
                        query_text = query_cell.inner_text().strip()
                        gsc_queries.append(query_text)
                        self.log_message.emit(f"提取到查询 {i+1}: {query_text}")
                    else:
                        # 备选方法：尝试其他选择器模式
                        query_cell = row.query_selector("td:first-child")
                        if query_cell:
                            query_text = query_cell.inner_text().strip()
                            gsc_queries.append(query_text)
                            self.log_message.emit(f"使用备选选择器提取到查询 {i+1}: {query_text}")
                        else:
                            self.log_message.emit(f"无法提取第 {i+1} 行的查询文本")
                except Exception as row_error:
                    self.log_message.emit(f"处理第 {i+1} 行时出错: {str(row_error)}")
            
            # 如果上述方法失败，尝试使用JavaScript评估来获取文本
            if not gsc_queries:
                self.log_message.emit("尝试使用JavaScript评估提取查询...")
                gsc_queries = page.evaluate("""
                    () => {
                        const rows = document.querySelectorAll("table tbody tr");
                        const queries = [];
                        for (let i = 0; i < Math.min(10, rows.length); i++) {
                            const cell = rows[i].querySelector("td:first-child");
                            if (cell) {
                                queries.push(cell.textContent.trim());
                            }
                        }
                        return queries;
                    }
                """)
                self.log_message.emit(f"使用JavaScript评估提取到 {len(gsc_queries)} 个查询")
            
            # 更新markdown文件
            self.update_markdown_file(page_name, gsc_queries, "GSC热门查询")
            
        except Exception as extract_error:
            self.log_message.emit(f"提取查询时出错: {str(extract_error)}")
    
    def update_markdown_file(self, page_name, items, section_name):
        """更新或创建MD文件并填入提取的内容"""
        if not items:
            self.log_message.emit(f"没有{section_name}结果可以更新到MD文件")
            return
            
        # 创建目标MD文件名
        md_file_path = f"{page_name}.md"
        
        # 检查文件是否存在，如果不存在则创建一个基本结构
        if not os.path.exists(md_file_path):
            # 直接创建一个基本结构，不再使用模板
            self.log_message.emit(f"创建新的MD文件: {md_file_path}")
            # 获取显示名称（替换连字符为空格，首字母大写）
            display_name = page_name.replace("-", " ").title()
            with open(md_file_path, "w", encoding="utf-8") as file:
                file.write(f"""
# {display_name}

## 关键词来源

### Google 搜索下拉框

### 相关搜索

### GSC热门查询

### 相关问题
""")
        
        # 读取现有内容
        try:
            with open(md_file_path, "r", encoding="utf-8") as file:
                md_content = file.read()
        except Exception as e:
            self.log_message.emit(f"读取MD文件时出错: {str(e)}")
            return
        
        # 查找对应部分
        section_header = f"### {section_name}"
        section_index = md_content.find(section_header)
        next_section_index = md_content.find("###", section_index + 1)
        
        if section_index != -1:
            # 构建新内容 - 使用标准的Markdown有序列表格式
            new_content = ""
            for i, item in enumerate(items):
                new_content += f"{i+1}. {item}\n"
            
            # 插入新内容
            if next_section_index != -1:
                updated_content = md_content[:section_index + len(section_header)] + "\n\n" + new_content + "\n" + md_content[next_section_index:]
            else:
                updated_content = md_content[:section_index + len(section_header)] + "\n\n" + new_content
            
            # 保存更新后的内容
            with open(md_file_path, "w", encoding="utf-8") as file:
                file.write(updated_content)
            
            self.log_message.emit(f"成功将 {len(items)} 个{section_name}结果保存到 {md_file_path}")
        else:
            self.log_message.emit(f"在文件 {md_file_path} 中未找到'{section_header}'部分")
    
    def process_ga(self, page, ga_url, page_name, ga_screenshot_path, screenshot_dir):
        """处理GA相关的任务"""
        try:
            self.log_message.emit(f"导航到GA4分析页面: {ga_url}")
            page.goto(ga_url, timeout=90000)
            
            # 添加更长的等待时间让页面完全加载
            self.log_message.emit("等待10秒让GA4页面完全加载...")
            time.sleep(10)
            
            # 执行额外的页面交互，帮助确保内容加载
            try:
                # 尝试滚动页面以确保触发懒加载内容
                page.evaluate("""
                    window.scrollTo(0, 100);
                    setTimeout(() => { window.scrollTo(0, 0); }, 500);
                """)
                self.log_message.emit("执行页面滚动以触发内容加载")
            except Exception as scroll_error:
                self.log_message.emit(f"页面滚动时出错: {str(scroll_error)}")
            
            # 直接尝试定位GA4报表元素，使用新的选择器组合
            ga_selectors = [
                # 原始选择器
                "body > ga-hybrid-app-root > ui-view-wrapper > div > app-root > div > div > ui-view-wrapper > div > ga-report-container > div > div > div > report-view > ui-view-wrapper > div > ui-view > ga-explorer-report > div > div > div > ga-card-list.explorer-card-list.ga-card-list.ng-star-inserted > div",
                # 新版GA4的可能选择器
                "report-view ga-explorer-report .explorer-cards-wrap",
                "ga-report-container .grid-layout-wrap",
                "ga-card-list.explorer-card-list",
                ".ga-card-list",
                "ga-report-container report-view",
                # 更通用的选择器
                "report-view .visualize-item-wrap",
                ".explorer-card-content"
            ]
            
            # 获取GA4页面结构以便找到正确的选择器
            self.log_message.emit("分析GA4页面结构...")
            try:
                # 使用JavaScript来分析页面结构并找到可能的报表元素
                selectors_info = page.evaluate("""
                    () => {
                        // 尝试查找GA4报表的各种可能元素
                        const possibleElements = [
                            // 卡片容器
                            document.querySelectorAll('ga-card-list'),
                            document.querySelectorAll('.ga-card-list'),
                            document.querySelectorAll('.grid-layout-wrap'),
                            document.querySelectorAll('.explorer-cards-wrap'),
                            // 报表容器
                            document.querySelectorAll('report-view'),
                            document.querySelectorAll('ga-explorer-report'),
                            document.querySelectorAll('.visualize-item-wrap'),
                            // 图表元素
                            document.querySelectorAll('.explorer-card'),
                            document.querySelectorAll('.explorer-card-content')
                        ];
                        
                        // 找到元素数量最多的集合
                        let maxElements = null;
                        let maxCount = 0;
                        let description = '';
                        
                        possibleElements.forEach((collection, index) => {
                            if (collection && collection.length > maxCount) {
                                maxElements = collection;
                                maxCount = collection.length;
                                if (collection.length > 0 && collection[0]) {
                                    description += `找到 ${collection.length} 个元素，类型: ${collection[0].tagName || 'unknown'}, `;
                                    description += `类名: ${collection[0].className || 'no-class'}\n`;
                                }
                            }
                        });
                        
                        // 尝试获取最大的报表容器元素
                        let mainReportContainer = null;
                        try {
                            const reportContainers = document.querySelectorAll('ga-report-container');
                            if (reportContainers && reportContainers.length > 0) {
                                // 找到最大的报表容器
                                let maxArea = 0;
                                for (const container of reportContainers) {
                                    const rect = container.getBoundingClientRect();
                                    const area = rect.width * rect.height;
                                    if (area > maxArea) {
                                        maxArea = area;
                                        mainReportContainer = container;
                                    }
                                }
                            }
                        } catch (e) {
                            description += `查找报表容器错误: ${e.message}\n`;
                        }
                        
                        // 获取页面结构
                        const pageStructure = [];
                        try {
                            // 查找主要内容区域
                            const reportView = document.querySelector('report-view');
                            if (reportView) {
                                // 生成选择器
                                const getPath = (el) => {
                                    if (!el) return '';
                                    if (el === document.body) return 'body';
                                    
                                    let path = '';
                                    
                                    // 特殊处理组件标签
                                    if (el.tagName && el.tagName.toLowerCase().includes('-')) {
                                        path = el.tagName.toLowerCase();
                                    } else {
                                        path = el.tagName.toLowerCase();
                                        if (el.className) {
                                            const classes = el.className.split(' ')
                                                .filter(c => c && !c.includes('ng-'))
                                                .map(c => '.' + c)
                                                .join('');
                                            if (classes) path += classes;
                                        }
                                    }
                                    
                                    return getPath(el.parentElement) + ' > ' + path;
                                };
                                
                                // 获取主要内容元素及其父元素链的选择器
                                pageStructure.push({
                                    element: 'reportView',
                                    selector: getPath(reportView)
                                });
                                
                                // 查找报表卡片
                                const cards = reportView.querySelectorAll('.explorer-card, .explorer-card-content');
                                if (cards && cards.length > 0) {
                                    pageStructure.push({
                                        element: 'cards',
                                        count: cards.length,
                                        selector: getPath(cards[0])
                                    });
                                }
                            }
                        } catch (e) {
                            description += `生成选择器错误: ${e.message}\n`;
                        }
                        
                        return {
                            description: description,
                            elementCount: maxCount,
                            pageStructure: pageStructure,
                            // 提供最可能的新选择器
                            recommendedSelector: pageStructure.length > 0 ? 
                                pageStructure[pageStructure.length - 1].selector : null
                        };
                    }
                """)
                
                # 记录找到的选择器信息
                if selectors_info:
                    self.log_message.emit(f"GA4页面分析结果:\n{selectors_info.get('description', '')}")
                    if selectors_info.get('recommendedSelector'):
                        recommended_selector = selectors_info.get('recommendedSelector')
                        self.log_message.emit(f"推荐的GA4选择器: {recommended_selector}")
                        # 如果找到了推荐选择器，将其添加到尝试列表的开头
                        if recommended_selector not in ga_selectors:
                            ga_selectors.insert(0, recommended_selector)
                            self.log_message.emit(f"已将推荐选择器添加到尝试列表")
                    else:
                        self.log_message.emit("未能生成推荐选择器")
                
                # 在分析完页面后再等待一会，确保元素完全渲染
                time.sleep(2)
                
            except Exception as analyze_error:
                self.log_message.emit(f"分析GA4页面结构时出错: {str(analyze_error)}")
            
            self.log_message.emit("定位GA4报表元素...")
            
            # 尝试所有可能的选择器
            found_element = False
            for ga_selector in ga_selectors:
                try:
                    # 等待元素出现
                    self.log_message.emit(f"尝试选择器: {ga_selector}")
                    ga_element = page.wait_for_selector(ga_selector, timeout=15000)  # 减少每个选择器的超时时间
                    
                    if ga_element:
                        self.log_message.emit(f"找到GA4元素，使用选择器: {ga_selector}")
                        self.log_message.emit("正在截图...")
                        ga_element.screenshot(path=ga_screenshot_path)
                        self.log_message.emit(f"GA4截图已保存为: {ga_screenshot_path}")
                        found_element = True
                        break
                except Exception as wait_error:
                    self.log_message.emit(f"使用选择器 {ga_selector} 查找元素失败: {str(wait_error)}")
                    continue
            
            # 如果所有选择器都失败了，尝试直接查询
            if not found_element:
                self.log_message.emit("尝试直接查询...")
                for ga_selector in ga_selectors:
                    try:
                        ga_element = page.query_selector(ga_selector)
                        if ga_element:
                            self.log_message.emit(f"使用直接查询找到GA4元素，选择器: {ga_selector}")
                            ga_element.screenshot(path=ga_screenshot_path)
                            self.log_message.emit(f"GA4截图已保存为: {ga_screenshot_path}")
                            found_element = True
                            break
                    except Exception as query_error:
                        continue
            
            # 如果上述所有方法都失败，截取整个页面
            if not found_element:
                self.log_message.emit("未找到特定GA4元素，截取整个页面...")
                ga_full_path = os.path.join(screenshot_dir, f"ga-{page_name}-full.png")
                page.screenshot(path=ga_full_path, full_page=True)
                self.log_message.emit(f"GA4整页截图已保存为: {ga_full_path}")
        except Exception as ga_error:
            self.log_message.emit(f"GA4截图过程中发生错误: {str(ga_error)}")
            try:
                # 截取当前页面作为错误记录
                ga_error_path = os.path.join(screenshot_dir, f"ga-{page_name}-error.png")
                page.screenshot(path=ga_error_path)
                self.log_message.emit(f"错误状态截图已保存为: {ga_error_path}")
            except:
                self.log_message.emit("无法保存GA4错误截图")
                
    def process_google_search_incognito(self, playwright, search_query, page_name, screenshot_dir):
        """在无痕模式下处理Google搜索下拉框、PAA和相关搜索"""
        # 使用无痕模式启动新的浏览器实例
        self.log_message.emit("以无痕模式启动浏览器进行Google搜索...")
        
        # 随机选择用户代理
        user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0"
        ]
        user_agent = random.choice(user_agents)
        
        # 获取模式设置
        headless = self.settings.value("headless_mode", "false") == "true"
        invisible_browser = self.settings.value("invisible_browser", "true") == "true"
        
        # 增强无痕模式下的反检测浏览器参数
        incognito_args = [
            '--disable-blink-features=AutomationControlled',
            '--no-sandbox',
            '--disable-extensions',
            '--disable-default-apps',
            '--disable-popup-blocking',
            '--start-maximized',
            f'--user-agent={user_agent}',
            # 增加以下参数提高匿名性和绕过检测
            '--disable-web-security',
            '--disable-features=IsolateOrigins,site-per-process',
            '--disable-site-isolation-trials',
            '--disable-blink-features',
            '--disable-device-orientation',
            '--disable-features=Translate',
            '--disable-infobars',
            '--ignore-certifcate-errors',
            '--ignore-certifcate-errors-spki-list',
            '--allow-running-insecure-content',
            '--disable-gpu',
            # 增强隐私保护
            '--incognito',
            '--disable-plugins-discovery',
            '--disable-notifications',
            '--disable-permissions-api'
        ]
        
        # 设置隐形浏览器的特定参数
        if invisible_browser:
            incognito_args += [
                '--window-size=1920,1080'
                # 移除 '--window-position=-32000,-32000' 参数
            ]
        
        # 根据设置选择启动模式
        if headless:
            # 完全无头模式
            self.log_message.emit("以完全无头模式进行搜索 (可能被检测)")
            browser_context = playwright.chromium.launch(
                headless=True,
                args=incognito_args
            )
        elif invisible_browser:
            # 隐形浏览器模式 - 兼容性方法
            self.log_message.emit("以隐形浏览器模式进行搜索 (降低被检测风险)")
            try:
                # 尝试使用带is_visible参数的方法（较新版本Playwright）
                browser_context = playwright.chromium.launch(
                    headless=False,
                    is_visible=False,
                    args=incognito_args
                )
            except TypeError:
                # 如果is_visible参数不被支持，使用标准方法
                self.log_message.emit("当前Playwright版本不支持is_visible参数，使用备选方法")
                browser_context = playwright.chromium.launch(
                    headless=False,
                    args=incognito_args
                )
        else:
            # 标准有头模式
            self.log_message.emit("以有头模式进行搜索")
            browser_context = playwright.chromium.launch(
                headless=False,
                args=incognito_args
            )
        
        try:
            # 创建一个新的上下文（相当于一个新的无痕窗口）
            context = browser_context.new_context(
                viewport={'width': 1920, 'height': 1080},
                user_agent=user_agent,
                java_script_enabled=True,
                ignore_https_errors=True,
                # 设置地理位置模拟中国
                geolocation={"latitude": 39.9042, "longitude": 116.4074},
                locale='zh-CN',
                timezone_id='Asia/Shanghai',
                reduced_motion='reduce'  # 减少动画，可能降低CPU使用率
            )
            
            # 创建新页面
            page = context.new_page()
            
            # 如果使用隐形模式，确保窗口在屏幕外
            if invisible_browser:
                try:
                    # 直接在页面上执行移动窗口的脚本，使用改进的方法
                    page.evaluate("""
                        try {
                            // 尝试将窗口移到屏幕外，但使用更可靠的值
                            window.moveTo(-10000, -10000);
                            // 如果不成功，尝试另一种方法
                            if (window.screenX > -5000) {
                                window.moveTo(-2000, -2000);
                            }
                            window.resizeTo(1920, 1080);
                        } catch (e) {
                            console.error("无法移动窗口", e);
                        }
                    """)
                except Exception as e:
                    self.log_message.emit(f"设置隐形窗口时出错: {str(e)}")
            
            # 应用额外的反检测措施
            context.add_init_script("""
                // 覆盖navigator.webdriver
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => false,
                });
                
                // 覆盖window.navigator.chrome
                window.navigator.chrome = {
                    runtime: {},
                    app: {
                        InstallState: {
                            DISABLED: 'disabled',
                            INSTALLED: 'installed',
                            NOT_INSTALLED: 'not_installed'
                        },
                        RunningState: {
                            CANNOT_RUN: 'cannot_run',
                            READY_TO_RUN: 'ready_to_run',
                            RUNNING: 'running'
                        },
                        isInstalled: true
                    }
                };
                
                // 覆盖window.chrome
                window.chrome = {
                    runtime: {},
                    app: {
                        isInstalled: true
                    }
                };
                
                // 覆盖语言设置
                Object.defineProperty(navigator, 'languages', {
                    get: () => ['zh-CN', 'zh', 'en-US', 'en'],
                });
                
                // 模拟正常的硬件并发层级
                Object.defineProperty(navigator, 'hardwareConcurrency', {
                    get: () => 8,
                });
                
                // 模拟正常的设备内存
                Object.defineProperty(navigator, 'deviceMemory', {
                    get: () => 8,
                });
                
                // 修改屏幕尺寸信息
                Object.defineProperty(window, 'screen', {
                    get: () => ({
                        availHeight: 1040,
                        availLeft: 0,
                        availTop: 0,
                        availWidth: 1920,
                        colorDepth: 24,
                        height: 1080,
                        width: 1920,
                        pixelDepth: 24
                    })
                });
                
                // 修改连接信息为非慢速
                Object.defineProperty(navigator, 'connection', {
                    get: () => ({
                        effectiveType: '4g',
                        rtt: 50,
                        downlink: 10.0,
                        saveData: false
                    })
                });
                
                // Canvas指纹修改
                const originalToDataURL = HTMLCanvasElement.prototype.toDataURL;
                HTMLCanvasElement.prototype.toDataURL = function(type) {
                    if (this.width > 1 && this.height > 1) {
                        // 轻微修改Canvas数据来改变指纹
                        const context = this.getContext("2d");
                        const imageData = context.getImageData(0, 0, 1, 1);
                        // 随机改变一个像素
                        imageData.data[0] = imageData.data[0] < 255 ? imageData.data[0] + 1 : imageData.data[0] - 1;
                        context.putImageData(imageData, 0, 0);
                    }
                    return originalToDataURL.apply(this, arguments);
                };
            """)
            
            # 创建新页面
            page = context.new_page()
            
            try:
                # 导航到Google搜索
                self.log_message.emit(f"以无痕模式导航到Google搜索页面...")
                page.goto("https://www.google.com/", timeout=30000)
                
                # 检查并处理同意条款页面
                self.handle_consent_page(page)
                
                # 等待搜索框加载
                search_selector = "textarea[name='q']"
                self.log_message.emit("等待搜索框加载...")
                page.wait_for_selector(search_selector, state="visible", timeout=30000)
                
                # 输入搜索词
                self.log_message.emit(f"输入搜索词: {search_query}")
                page.fill(search_selector, search_query)
                
                # 等待搜索下拉框加载
                self.log_message.emit("等待搜索下拉框加载...")
                time.sleep(3)  # 给下拉框更多时间加载
                
                # 获取搜索下拉框内容
                dropdown_suggestions = self.extract_dropdown_suggestions(page)
                if dropdown_suggestions:
                    self.log_message.emit(f"提取到 {len(dropdown_suggestions)} 个搜索下拉框建议")
                    for i, suggestion in enumerate(dropdown_suggestions):
                        self.log_message.emit(f"建议 {i+1}: {suggestion}")
                    self.update_markdown_file(page_name, dropdown_suggestions, "Google 搜索下拉框")
                else:
                    self.log_message.emit("未能提取到搜索下拉框建议")
                    
                # 提交搜索
                self.log_message.emit("提交搜索...")
                page.press(search_selector, "Enter")
                page.wait_for_load_state("networkidle", timeout=30000)
                self.log_message.emit("搜索结果页面已加载")
                
                # 提取PAA问题
                self.log_message.emit("开始提取PAA问题...")
                paa_questions = self.extract_paa_questions(page)
                if paa_questions:
                    self.log_message.emit(f"提取到 {len(paa_questions)} 个PAA问题")
                    for i, question in enumerate(paa_questions):
                        self.log_message.emit(f"问题 {i+1}: {question}")
                    self.update_markdown_file(page_name, paa_questions, "相关问题")
                else:
                    self.log_message.emit("未能提取到PAA问题")
                
                # 提取相关搜索前更充分地滚动页面
                self.log_message.emit("滚动到页面底部以加载相关搜索...")
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(2)  # 给页面更多时间加载底部内容
                
                # 提取相关搜索
                self.log_message.emit("开始提取相关搜索...")
                related_searches = self.extract_related_searches(page)
                if related_searches:
                    self.log_message.emit(f"提取到 {len(related_searches)} 个相关搜索")
                    for i, search in enumerate(related_searches):
                        self.log_message.emit(f"相关搜索 {i+1}: {search}")
                    self.update_markdown_file(page_name, related_searches, "相关搜索")
                else:
                    self.log_message.emit("未能提取到相关搜索")
            
            except Exception as google_error:
                self.log_message.emit(f"无痕模式Google搜索过程中发生错误: {str(google_error)}")
                self.log_message.emit(f"错误详情: {google_error}")
            
            finally:
                # 关闭页面和上下文
                page.close()
                context.close()
        
        finally:
            # 关闭浏览器
            browser_context.close()
            
    def extract_dropdown_suggestions(self, page):
        """提取Google搜索下拉框建议"""
        try:
            # 等待下拉框出现 - 使用一个通用的选择器确保下拉框已加载
            dropdown_container_selector = "div[jsname='aajZCb']"
            page.wait_for_selector(dropdown_container_selector, state="visible", timeout=5000)
            
            # 使用更直接的方法提取搜索建议
            suggestions = page.evaluate("""
                () => {
                    // 尝试确定当前Google界面下的下拉框结构
                    function getAllSuggestions() {
                        // 不同的可能选择器组合
                        const possibleSelectors = [
                            // 针对当前截图所示结构
                            ".wM6W7d",
                            ".OBMEnb .wM6W7d",
                            "ul[role='listbox'] li",
                            "div[jsname='aajZCb'] .wM6W7d",
                            // 针对老结构
                            ".sbct",
                            ".sbsb_a li",
                            ".sbpqs_a li",
                            ".G43f7e li"
                        ];
                        
                        let elements = [];
                        // 尝试所有可能的选择器
                        for (const selector of possibleSelectors) {
                            const found = document.querySelectorAll(selector);
                            if (found && found.length > 0) {
                                elements = Array.from(found);
                                console.log(`找到选择器 ${selector} 匹配的元素: ${found.length} 个`);
                                break;
                            }
                        }
                        
                        // 如果没有找到任何元素，返回空数组
                        if (elements.length === 0) {
                            console.log("未找到任何匹配的下拉框元素");
                            return [];
                        }
                        
                        // 处理找到的元素，提取文本
                        return elements.map(el => {
                            // 获取纯文本内容
                            return el.textContent.trim();
                        }).filter(text => text.length > 0); // 过滤掉空文本
                    }
                    
                    // 调用方法获取所有建议
                    const results = getAllSuggestions();
                    console.log(`找到 ${results.length} 个下拉框建议`);
                    console.log("建议内容:", results);
                    
                    return results;
                }
            """)
            
            self.log_message.emit(f"找到 {len(suggestions)} 个搜索下拉框建议")
            
            if len(suggestions) == 0:
                # 如果无法提取到建议，尝试使用最后的备选方法
                self.log_message.emit("尝试使用备选方法从页面源码提取下拉建议...")
                
                # 将页面源码保存到文件以便分析
                page_content = page.content()
                debug_dir = os.path.join(self.settings.value("screenshot_dir", "screenshots"), "debug")
                if not os.path.exists(debug_dir):
                    os.makedirs(debug_dir)
                with open(os.path.join(debug_dir, "page_source.html"), "w", encoding="utf-8") as f:
                    f.write(page_content)
                
                # 使用正则表达式从页面源码中提取可能的搜索建议
                import re
                search_query = page.input_value("textarea[name='q']")
                self.log_message.emit(f"当前搜索词: {search_query}")
                
                # 尝试匹配基于当前搜索词的建议
                pattern = re.compile(f'({re.escape(search_query)}[^<>"]*)', re.IGNORECASE)
                matches = pattern.findall(page_content)
                
                if matches:
                    # 仅选择有意义的匹配项（长度合适且不包含HTML标签）
                    valid_matches = [m for m in matches if 
                                    len(m) > len(search_query) and 
                                    len(m) < 100 and 
                                    '<' not in m and 
                                    '>' not in m]
                    
                    # 去重
                    unique_matches = list(set(valid_matches))
                    
                    self.log_message.emit(f"通过页面源码找到 {len(unique_matches)} 个可能的搜索建议")
                    return unique_matches
            
            return suggestions
            
        except Exception as dropdown_error:
            self.log_message.emit(f"提取搜索下拉框建议时出错: {str(dropdown_error)}")
            # 记录更多调试信息
            self.log_message.emit(f"错误详情: {dropdown_error}")
            return []
    
    def extract_paa_questions(self, page):
        """提取PAA（People Also Ask）问题"""
        try:
            # 确保页面有足够时间加载PAA部分
            time.sleep(1)
            
            # 使用更全面的JavaScript方法提取PAA问题
            self.log_message.emit("使用JavaScript方法提取PAA问题...")
            questions = page.evaluate("""
                () => {
                    // 辅助函数：获取元素的可见文本，忽略隐藏元素
                    function getVisibleText(element) {
                        if (!element) return '';
                        
                        const style = window.getComputedStyle(element);
                        if (style.display === 'none' || style.visibility === 'hidden') return '';
                        
                        let text = '';
                        for (const child of element.childNodes) {
                            if (child.nodeType === Node.TEXT_NODE) {
                                text += child.textContent.trim() + ' ';
                            } else if (child.nodeType === Node.ELEMENT_NODE) {
                                text += getVisibleText(child) + ' ';
                            }
                        }
                        return text.trim();
                    }
                    
                    // 尝试不同的选择器来找到PAA问题
                    const questions = new Set(); // 使用Set去重
                    
                    // 最新的选择器列表，按可能性排序
                    const selectors = [
                        // 常见的PAA容器选择器
                        "div.related-question-pair",
                        ".g .related-question-pair",
                        ".related-questions-pair",
                        "div[jsname='N760b']",
                        
                        // 直接选择问题文本元素
                        ".related-question-pair .JlqpRe",
                        ".related-question-pair .wQiwMc .JlqpRe",
                        "div[jsname='Cpkphb'] .JlqpRe",
                        "div[jsname='N760b'] .wQiwMc .JlqpRe",
                        "div[data-ved] .CSkcDe",
                        
                        // 额外尝试其他可能的问题选择器
                        ".e24Kjd",
                        ".iDjcJe",
                        "[role='heading']"
                    ];
                    
                    // 对于每个选择器，尝试提取问题
                    for (const selector of selectors) {
                        const elements = document.querySelectorAll(selector);
                        if (elements && elements.length > 0) {
                            for (const element of elements) {
                                // 依赖于结构的选择器
                                let questionText = '';
                                
                                // 尝试寻找专门的问题容器
                                const questionContainer = element.querySelector('.JlqpRe, .CSkcDe, [role="heading"], .wWOJcd, .e24Kjd, .iDjcJe');
                                
                                if (questionContainer) {
                                    questionText = getVisibleText(questionContainer);
                                } else {
                                    // 如果没有找到专门的容器，尝试获取元素本身的文本
                                    questionText = getVisibleText(element);
                                }
                                
                                // 验证并添加问题
                                if (questionText && questionText.length > 10 && questionText.length < 200) {
                                    // 仅添加符合问题长度的合理文本（避免过短或过长）
                                    // 并过滤掉明显不是问题的文本
                                    questions.add(questionText);
                                }
                            }
                        }
                        
                        // 如果我们找到了问题，就不需要继续尝试
                        if (questions.size > 0) {
                            break;
                        }
                    }
                    
                    // 如果上面的方法都失败了，尝试最后的备选方法：
                    // 搜索页面中看起来像问题的标题元素
                    if (questions.size === 0) {
                        const headings = document.querySelectorAll('h3, h4, [role="heading"]');
                        for (const heading of headings) {
                            const text = getVisibleText(heading);
                            // 检查文本是否看起来像问题（含有问号或问题词）
                            if (text && (
                                text.includes('?') || 
                                text.includes('how') || 
                                text.includes('what') || 
                                text.includes('why') || 
                                text.includes('when') || 
                                text.includes('where') ||
                                text.includes('which') ||
                                text.includes('who') ||
                                text.includes('can') ||
                                text.includes('do')
                            ) && text.length > 15 && text.length < 200) {
                                questions.add(text);
                            }
                        }
                    }
                    
                    return Array.from(questions);
                }
            """)
            
            self.log_message.emit(f"通过JavaScript评估找到 {len(questions)} 个PAA问题")
            return questions
            
        except Exception as paa_error:
            self.log_message.emit(f"提取PAA问题时出错: {str(paa_error)}")
            self.log_message.emit(f"详细错误: {paa_error}")
            return []
    
    def extract_related_searches(self, page):
        """提取相关搜索"""
        try:
            # 相关搜索部分通常位于页面底部
            # 先滚动到页面底部
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(2)  # 增加等待时间确保内容完全加载
            
            # 使用您提供的最新选择器
            new_selector = "#bres > div.ULSxyf > div > div > div > div.y6Uyqe > div > div:nth-child(1) > div:nth-child(4) > div > div > a > div > div.wyccme > div > div > div > span"
            
            # 尝试获取相关搜索
            self.log_message.emit(f"尝试使用新选择器获取相关搜索...")
            
            # 定位相关搜索部分的所有链接
            # 使用更通用的选择器组合
            searches = []
            
            # 使用JavaScript评估提取相关搜索，处理各种可能的HTML结构
            searches = page.evaluate("""
                () => {
                    const searches = [];
                    
                    // 尝试获取相关搜索
                    function getSearchesFromElements(elements) {
                        const results = [];
                        for (const element of elements) {
                            const textContent = element.textContent.trim();
                            if (textContent && !textContent.match(/^(全部|视频|短视频|图片|购物|新闻|网页|图书|地图|航班)$/)) {
                                results.push(textContent);
                            }
                        }
                        return results;
                    }
                    
                    // 尝试不同的选择器组合
                    const selectorCombinations = [
                        // 您提供的新选择器
                        "#bres > div.ULSxyf > div > div > div > div.y6Uyqe > div > div > div > div > div > a > div > div.wyccme > div > div > div > span",
                        
                        // 更简化的选择器，捕获相关搜索文本区域
                        "div.y6Uyqe a div.wyccme span",
                        "div.y6Uyqe a div.dXS2h span",
                        
                        // 使用文本容器类
                        "div.y6Uyqe a div.mtv5bd span.dg6jd",
                        
                        // 使用父容器定位相关搜索部分
                        "#botstuff div.card-section a",
                        "div.brs_col a span",
                        
                        // 尝试完全不同的方法 - 寻找页面底部带有标题的部分
                        "div[data-hveid] a" // 通用相关内容选择器
                    ];
                    
                    for (const selector of selectorCombinations) {
                        const elements = document.querySelectorAll(selector);
                        if (elements.length > 0) {
                            const results = getSearchesFromElements(elements);
                            if (results.length > 0) {
                                searches.push(...results);
                                break; // 如果我们找到了相关搜索，就停止尝试其他选择器
                            }
                        }
                    }
                    
                    // 如果上述方法失败，尝试最后的通用方法：寻找页面底部的所有链接
                    if (searches.length === 0) {
                        // 获取页面下半部分的所有链接
                        const allLinks = Array.from(document.querySelectorAll('a'));
                        const pageHeight = document.body.scrollHeight;
                        const bottomLinks = allLinks.filter(link => {
                            const rect = link.getBoundingClientRect();
                            const linkTop = rect.top + window.pageYOffset;
                            return linkTop > pageHeight * 0.7; // 只考虑页面底部70%区域的链接
                        });
                        
                        // 从这些链接中过滤出可能的相关搜索
                        for (const link of bottomLinks) {
                            const text = link.textContent.trim();
                            // 排除导航链接和无意义的短文本
                            if (text && text.length > 3 && !text.match(/^(全部|视频|短视频|图片|购物|新闻|网页|图书|地图|航班)$/)) {
                                searches.push(text);
                            }
                        }
                    }
                    
                    return [...new Set(searches)]; // 返回去重后的结果
                }
            """)
            
            self.log_message.emit(f"找到 {len(searches)} 个相关搜索")
            
            # 排除Google导航分类
            excluded_terms = ["全部", "视频", "短视频", "图片", "购物", "新闻", "网页", "图书", "地图", "航班"]
            filtered_searches = [s for s in searches if s not in excluded_terms]
            
            if len(filtered_searches) < len(searches):
                self.log_message.emit(f"过滤掉 {len(searches) - len(filtered_searches)} 个无关项")
            
            return filtered_searches
            
        except Exception as related_error:
            self.log_message.emit(f"提取相关搜索时出错: {str(related_error)}")
            self.log_message.emit(f"错误详情: {related_error}")
            return []
            
    def handle_consent_page(self, page):
        """处理Google同意条款页面"""
        try:
            # 检查是否在同意条款页面
            if "consent.google.com" in page.url:
                self.log_message.emit("检测到Google同意条款页面，尝试点击同意按钮...")
                
                # 尝试点击"我同意"按钮（不同地区可能有不同ID）
                consent_buttons = [
                    "button#L2AGLb",  # 常见的"我同意"按钮ID
                    "button[aria-label='同意使用 Cookie']",
                    "button[jsname='higCR']",  # 另一种可能的ID
                    "form:nth-child(2) button"  # 基于位置的选择器
                ]
                
                for button_selector in consent_buttons:
                    try:
                        if page.query_selector(button_selector):
                            page.click(button_selector)
                            self.log_message.emit(f"已点击同意按钮: {button_selector}")
                            # 等待页面导航完成
                            page.wait_for_navigation(timeout=10000)
                            break
                    except Exception as click_error:
                        self.log_message.emit(f"点击按钮 {button_selector} 时出错: {str(click_error)}")
                
                # 确认是否已离开同意页面
                if "consent.google.com" not in page.url:
                    self.log_message.emit("已成功处理同意条款页面")
                else:
                    self.log_message.emit("未能自动处理同意条款页面，请在浏览器中手动操作...")
                    # 等待用户手动操作
                    page.wait_for_url(lambda url: "consent.google.com" not in url, timeout=60000)
                    self.log_message.emit("检测到已离开同意条款页面")
        except Exception as consent_error:
            self.log_message.emit(f"处理同意条款页面时出错: {str(consent_error)}")

    def handle_google_login(self, page, username, password):
        """处理Google账号登录流程"""
        try:
            self.log_message.emit("开始处理Google登录...")
            
            # 第一步：输入邮箱
            self.log_message.emit("查找并填写邮箱输入框...")
            # 等待邮箱输入框出现
            email_selector = "input[type='email']"
            page.wait_for_selector(email_selector, state="visible", timeout=30000)
            
            # 随机延迟模拟人工输入
            time.sleep(random.uniform(0.5, 1.5))
            
            # 填写邮箱
            page.fill(email_selector, username)
            self.log_message.emit("邮箱已输入")
            
            # 点击"下一步"按钮
            next_button_selector = "button:has-text('下一步'), button:has-text('Next')"
            self.log_message.emit("点击下一步按钮...")
            page.click(next_button_selector)
            
            # 第二步：输入密码
            self.log_message.emit("等待密码输入框出现...")
            password_selector = "input[type='password']"
            page.wait_for_selector(password_selector, state="visible", timeout=30000)
            
            # 随机延迟
            time.sleep(random.uniform(1.0, 2.0))
            
            # 填写密码
            page.fill(password_selector, password)
            self.log_message.emit("密码已输入")
            
            # 点击"下一步"按钮登录
            self.log_message.emit("点击登录按钮...")
            page.click(next_button_selector)
            
            # 等待登录完成，页面跳转
            self.log_message.emit("等待登录完成并跳转...")
            page.wait_for_url(lambda url: "search-console" in url or "search.google.com" in url, timeout=60000)
            
            # 额外检查是否存在二次验证或其他安全检查
            if "accounts.google.com" in page.url or "signin" in page.url:
                self.log_message.emit("检测到需要额外验证，可能需要手动操作...")
                page.wait_for_url(lambda url: "search-console" in url or "search.google.com" in url, timeout=120000)
                
            self.log_message.emit("登录成功完成")
            return True
            
        except Exception as e:
            self.log_message.emit(f"自动登录过程中出错: {str(e)}")
            self.log_message.emit("尝试等待手动登录...")
            # 仍然等待用户可能的手动登录
            page.wait_for_url(lambda url: "search-console" in url or "search.google.com" in url, timeout=120000)
            return False


class SeoRpaMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.worker = None
        
        # 加载设置
        self.settings = QSettings("SeoRpaTool", "Settings")
        self.load_settings()
        
    def init_ui(self):
        self.setWindowTitle("SEO关键词调研RPA工具")
        self.setGeometry(100, 100, 1200, 800)
        
        # 创建主窗口布局
        main_layout = QVBoxLayout()
        
        # 创建选项卡窗口
        self.tabs = QTabWidget()
        self.task_tab = QWidget()
        self.settings_tab = QWidget()
        
        self.tabs.addTab(self.task_tab, "任务")
        self.tabs.addTab(self.settings_tab, "设置")
        
        # 设置任务选项卡
        self.setup_task_tab()
        
        # 设置设置选项卡
        self.setup_settings_tab()
        
        main_layout.addWidget(self.tabs)
        
        # 设置主窗口部件
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        
    def setup_task_tab(self):
        # 创建任务选项卡布局
        task_layout = QVBoxLayout()
        
        # URL输入区域
        self.url_group = QGroupBox("URL列表")
        url_layout = QVBoxLayout()
        
        # URL输入多行文本框
        self.url_input = QTextEdit()
        self.url_input.setPlaceholderText("输入要处理的URL，每行一个")
        url_layout.addWidget(self.url_input)
        
        # URL相关按钮
        url_buttons_layout = QHBoxLayout()
        self.load_urls_button = QPushButton("从文件加载URL")
        self.clear_urls_button = QPushButton("清空URL")
        url_buttons_layout.addWidget(self.load_urls_button)
        url_buttons_layout.addWidget(self.clear_urls_button)
        url_layout.addLayout(url_buttons_layout)
        
        self.url_group.setLayout(url_layout)
        task_layout.addWidget(self.url_group)
        
        # 任务执行控制区域
        control_group = QGroupBox("任务控制")
        control_layout = QVBoxLayout()
        
        # 进度条
        progress_layout = QHBoxLayout()
        self.progress_label = QLabel("准备就绪")
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_label)
        progress_layout.addWidget(self.progress_bar)
        control_layout.addLayout(progress_layout)
        
        # 控制按钮
        buttons_layout = QHBoxLayout()
        self.start_button = QPushButton("开始任务")
        self.stop_button = QPushButton("停止任务")
        self.stop_button.setEnabled(False)
        buttons_layout.addWidget(self.start_button)
        buttons_layout.addWidget(self.stop_button)
        control_layout.addLayout(buttons_layout)
        
        control_group.setLayout(control_layout)
        task_layout.addWidget(control_group)
        
        # 日志区域
        log_group = QGroupBox("任务日志")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        # 日志按钮
        log_buttons_layout = QHBoxLayout()
        self.clear_log_button = QPushButton("清空日志")
        self.save_log_button = QPushButton("保存日志")
        log_buttons_layout.addWidget(self.clear_log_button)
        log_buttons_layout.addWidget(self.save_log_button)
        log_layout.addLayout(log_buttons_layout)
        
        log_group.setLayout(log_layout)
        task_layout.addWidget(log_group)
        
        # 设置任务选项卡布局
        self.task_tab.setLayout(task_layout)
        
        # 连接信号
        self.load_urls_button.clicked.connect(self.load_urls_from_file)
        self.clear_urls_button.clicked.connect(self.clear_urls)
        self.start_button.clicked.connect(self.start_task)
        self.stop_button.clicked.connect(self.stop_task)
        self.clear_log_button.clicked.connect(self.clear_log)
        self.save_log_button.clicked.connect(self.save_log)
        
    def setup_settings_tab(self):
        # 创建设置选项卡布局
        settings_layout = QVBoxLayout()
        
        # Chrome配置文件路径
        chrome_group = QGroupBox("Chrome配置")
        chrome_layout = QHBoxLayout()
        
        self.chrome_profile_input = QLineEdit()
        self.chrome_profile_input.setPlaceholderText("Chrome用户数据目录路径")
        
        self.select_chrome_profile_button = QPushButton("选择目录")
        
        chrome_layout.addWidget(self.chrome_profile_input, 7)
        chrome_layout.addWidget(self.select_chrome_profile_button, 3)
        
        chrome_group.setLayout(chrome_layout)
        settings_layout.addWidget(chrome_group)
        
        # 截图目录设置
        screenshot_group = QGroupBox("截图设置")
        screenshot_layout = QHBoxLayout()
        
        self.screenshot_dir_input = QLineEdit()
        self.screenshot_dir_input.setPlaceholderText("截图保存目录")
        
        self.select_screenshot_dir_button = QPushButton("选择目录")
        
        screenshot_layout.addWidget(self.screenshot_dir_input, 7)
        screenshot_layout.addWidget(self.select_screenshot_dir_button, 3)
        
        screenshot_group.setLayout(screenshot_layout)
        settings_layout.addWidget(screenshot_group)
        
        # 其他设置
        options_group = QGroupBox("其他设置")
        options_layout = QVBoxLayout()
        
        self.headless_checkbox = QCheckBox("无头模式 (不显示浏览器界面)")
        self.invisible_browser_checkbox = QCheckBox("隐形浏览器 (有浏览器但不可见，推荐用于绕过安全检测)")
        self.invisible_browser_checkbox.setChecked(True)
        self.scrape_ga_checkbox = QCheckBox("抓取GA数据")
        self.scrape_ga_checkbox.setChecked(True)
        self.scrape_gsc_checkbox = QCheckBox("抓取GSC数据")
        self.scrape_gsc_checkbox.setChecked(True)
        self.scrape_serp_checkbox = QCheckBox("抓取SERP元素")
        self.scrape_serp_checkbox.setChecked(True)
        self.scrape_semrush_checkbox = QCheckBox("抓取SEMrush关键词")
        self.scrape_semrush_checkbox.setChecked(True)
        self.original_article_checkbox = QCheckBox("原创文章模式 (输入关键词而非URL)")
        self.original_article_checkbox.setToolTip("启用后将只收集搜索相关数据，不抓取GSC和GA数据，可通过上方复选框控制具体抓取内容")
        
        # 添加工具提示
        self.headless_checkbox.setToolTip("完全无头模式，效率更高但可能被检测为机器人")
        self.invisible_browser_checkbox.setToolTip("在后台运行有头浏览器但不显示界面，可以更好地避免安全检测")
        self.scrape_ga_checkbox.setToolTip("是否抓取Google Analytics数据")
        self.scrape_gsc_checkbox.setToolTip("是否抓取Google Search Console数据")
        self.scrape_serp_checkbox.setToolTip("是否抓取Google搜索结果页面(SERP)数据")
        self.scrape_semrush_checkbox.setToolTip("是否抓取SEMrush关键词数据")
        
        # 设置无头模式和隐形浏览器复选框互斥
        def update_checkboxes():
            if self.headless_checkbox.isChecked():
                self.invisible_browser_checkbox.setChecked(False)
        
        def update_invisible_checkboxes():
            if self.invisible_browser_checkbox.isChecked():
                self.headless_checkbox.setChecked(False)
        
        def update_article_mode_checkboxes():
            if self.original_article_checkbox.isChecked():
                # 原创文章模式下禁用GSC和GA选项
                self.scrape_gsc_checkbox.setEnabled(False)
                self.scrape_ga_checkbox.setEnabled(False)
            else:
                # 非原创文章模式下启用所有选项
                self.scrape_gsc_checkbox.setEnabled(True)
                self.scrape_ga_checkbox.setEnabled(True)
            self.update_input_labels()
        
        self.headless_checkbox.stateChanged.connect(update_checkboxes)
        self.invisible_browser_checkbox.stateChanged.connect(update_invisible_checkboxes)
        self.original_article_checkbox.stateChanged.connect(update_article_mode_checkboxes)
        
        options_layout.addWidget(self.headless_checkbox)
        options_layout.addWidget(self.invisible_browser_checkbox)
        options_layout.addWidget(self.scrape_ga_checkbox)
        options_layout.addWidget(self.scrape_gsc_checkbox)
        options_layout.addWidget(self.scrape_serp_checkbox)
        options_layout.addWidget(self.scrape_semrush_checkbox)
        options_layout.addWidget(self.original_article_checkbox)
        
        options_group.setLayout(options_layout)
        settings_layout.addWidget(options_group)
        
        # 保存设置按钮
        save_settings_layout = QHBoxLayout()
        self.save_settings_button = QPushButton("保存设置")
        save_settings_layout.addStretch()
        save_settings_layout.addWidget(self.save_settings_button)
        settings_layout.addLayout(save_settings_layout)
        
        # 添加弹性空间
        settings_layout.addStretch()
        
        # 设置设置选项卡布局
        self.settings_tab.setLayout(settings_layout)
        
        # 连接信号
        self.select_chrome_profile_button.clicked.connect(self.select_chrome_profile)
        self.select_screenshot_dir_button.clicked.connect(self.select_screenshot_dir)
        self.save_settings_button.clicked.connect(self.save_settings)
        
    def update_input_labels(self):
        """根据原创文章模式切换更新输入框标签和提示"""
        if self.original_article_checkbox.isChecked():
            self.url_group.setTitle("关键词列表")
            self.url_input.setPlaceholderText("输入要研究的关键词，每行一个")
            self.load_urls_button.setText("从文件加载关键词")
            self.clear_urls_button.setText("清空关键词")
        else:
            self.url_group.setTitle("URL列表")
            self.url_input.setPlaceholderText("输入要处理的URL，每行一个")
            self.load_urls_button.setText("从文件加载URL")
            self.clear_urls_button.setText("清空URL")
    
    def load_urls_from_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择URL文件", "", "文本文件 (*.txt);;所有文件 (*)")
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    urls = file.read()
                self.url_input.setText(urls)
                self.log_message(f"从文件 {file_path} 加载了URL")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"加载URL文件时出错: {str(e)}")
                
    def clear_urls(self):
        self.url_input.clear()
        self.log_message("已清空URL列表")
        
    def select_chrome_profile(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择Chrome用户数据目录")
        if dir_path:
            self.chrome_profile_input.setText(dir_path)
            
    def select_screenshot_dir(self):
        dir_path = QFileDialog.getExistingDirectory(self, "选择截图保存目录")
        if dir_path:
            self.screenshot_dir_input.setText(dir_path)
            
    def save_settings(self):
        # 保存设置
        self.settings.setValue("chrome_profile", self.chrome_profile_input.text())
        self.settings.setValue("screenshot_dir", self.screenshot_dir_input.text())
        self.settings.setValue("headless_mode", "true" if self.headless_checkbox.isChecked() else "false")
        self.settings.setValue("invisible_browser", "true" if self.invisible_browser_checkbox.isChecked() else "false")
        self.settings.setValue("scrape_ga", "true" if self.scrape_ga_checkbox.isChecked() else "false")
        self.settings.setValue("scrape_gsc", "true" if self.scrape_gsc_checkbox.isChecked() else "false")
        self.settings.setValue("scrape_serp", "true" if self.scrape_serp_checkbox.isChecked() else "false")
        self.settings.setValue("scrape_semrush", "true" if self.scrape_semrush_checkbox.isChecked() else "false")
        self.settings.setValue("original_article_mode", "true" if self.original_article_checkbox.isChecked() else "false")
        
        QMessageBox.information(self, "设置", "设置已保存")
        self.log_message("设置已更新")
        
    def load_settings(self):
        # 加载设置
        self.chrome_profile_input.setText(self.settings.value("chrome_profile", ""))
        self.screenshot_dir_input.setText(self.settings.value("screenshot_dir", "screenshots"))
        self.headless_checkbox.setChecked(self.settings.value("headless_mode", "false") == "true")
        self.invisible_browser_checkbox.setChecked(self.settings.value("invisible_browser", "true") == "true")
        self.scrape_ga_checkbox.setChecked(self.settings.value("scrape_ga", "true") == "true")
        self.scrape_gsc_checkbox.setChecked(self.settings.value("scrape_gsc", "true") == "true")
        self.scrape_serp_checkbox.setChecked(self.settings.value("scrape_serp", "true") == "true")
        self.scrape_semrush_checkbox.setChecked(self.settings.value("scrape_semrush", "true") == "true")
        self.original_article_checkbox.setChecked(self.settings.value("original_article_mode", "false") == "true")
        
        # 确保无头模式和隐形浏览器模式不会同时被选中
        if self.headless_checkbox.isChecked() and self.invisible_browser_checkbox.isChecked():
            self.invisible_browser_checkbox.setChecked(True)
            self.headless_checkbox.setChecked(False)
        
        # 更新输入标签
        self.update_input_labels()
        
        # 更新原创文章模式下的复选框状态
        if self.original_article_checkbox.isChecked():
            self.scrape_gsc_checkbox.setEnabled(False)
            self.scrape_ga_checkbox.setEnabled(False)
        else:
            self.scrape_gsc_checkbox.setEnabled(True)
            self.scrape_ga_checkbox.setEnabled(True)
    
    def start_task(self):
        # 获取URL列表或关键词列表
        input_text = self.url_input.toPlainText().strip()
        if not input_text:
            if self.original_article_checkbox.isChecked():
                QMessageBox.warning(self, "错误", "请输入至少一个关键词")
            else:
                QMessageBox.warning(self, "错误", "请输入至少一个URL")
            return
            
        # 分割输入内容
        items = [item.strip() for item in input_text.split("\n") if item.strip()]
        
        # 更新UI状态
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.progress_bar.setValue(0)
        
        # 根据模式设置进度标签
        if self.original_article_checkbox.isChecked():
            self.progress_label.setText(f"处理中... (0/{len(items)}) - 原创文章模式")
        else:
            self.progress_label.setText(f"处理中... (0/{len(items)})")
        
        # 重定向标准输出到日志窗口
        self.stdout_redirect = LogRedirector(self.log_text)
        sys.stdout = self.stdout_redirect
        
        # 创建并启动工作线程
        self.worker = RPAWorker(items, self.settings)
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.log_message.connect(self.log_message)
        self.worker.task_completed.connect(self.on_task_completed)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.start()
        
    def stop_task(self):
        if self.worker and self.worker.isRunning():
            self.log_message("正在停止任务...")
            self.worker.abort()
            self.stop_button.setEnabled(False)
            
    def update_progress(self, current, total):
        percentage = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(percentage)
        self.progress_label.setText(f"处理中... ({current}/{total})")
        
    def on_task_completed(self, url, success):
        is_original_mode = self.original_article_checkbox.isChecked()
        
        if success:
            if is_original_mode:
                self.log_message(f"成功完成关键词: {url}")
            else:
                self.log_message(f"成功完成URL: {url}")
        else:
            if is_original_mode:
                self.log_message(f"处理关键词时出错: {url}")
            else:
                self.log_message(f"处理URL时出错: {url}")
            
    def on_worker_finished(self):
        # 恢复标准输出
        sys.stdout = sys.__stdout__
        
        # 更新UI状态
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.progress_label.setText("任务完成")
        
        QMessageBox.information(self, "任务完成", "所有URL处理完成")
        
    def log_message(self, message):
        self.log_text.append(message)
        self.log_text.moveCursor(QTextCursor.End)
        
    def clear_log(self):
        self.log_text.clear()
        
    def save_log(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "保存日志", "", "文本文件 (*.txt);;所有文件 (*)")
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(self.log_text.toPlainText())
                QMessageBox.information(self, "保存日志", f"日志已保存至 {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "错误", f"保存日志时出错: {str(e)}")
                
    def closeEvent(self, event):
        # 确保在关闭窗口时终止工作线程
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(self, '确认', '任务正在运行中，确定要退出吗？',
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.worker.abort()
                self.worker.wait()  # 等待线程结束
            else:
                event.ignore()
                return
        event.accept()

# 主程序入口
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion风格获得更现代的外观
    window = SeoRpaMainWindow()
    window.show()
    sys.exit(app.exec_())
