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
                # 启动浏览器
                browser = self.launch_browser(p)
                
                # 创建新页面
                page = browser.new_page()
                self.setup_page(page)
                
                # 处理GSC
                self.process_gsc(page, gsc_url, page_name, first_screenshot_path, second_screenshot_path, screenshot_dir)
                
                # 处理GA
                if self.settings.value("scrape_ga", "true") == "true":
                    self.process_ga(page, ga_url, page_name, ga_screenshot_path, screenshot_dir)
                
                # 关闭浏览器
                browser.close()
                
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
        user_data_dir = self.settings.value("chrome_profile", r"C:\Users\U191115\AppData\Local\Google\Chrome\User Data")
        
        # 随机选择用户代理
        user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0"
        ]
        user_agent = random.choice(user_agents)
        
        # 确定是否使用有头模式
        headless = self.settings.value("headless_mode", "false") == "true"
        
        # 启动浏览器
        browser = playwright.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=headless,
            viewport={'width': 1920, 'height': 1080},
            args=[
                '--profile-directory=Default',
                '--disable-blink-features=AutomationControlled',
                '--no-sandbox',
                '--disable-extensions',
                '--disable-default-apps',
                '--disable-popup-blocking',
                '--start-maximized',
                f'--user-agent={user_agent}'
            ]
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
        client.send('Page.addScriptToEvaluateOnNewDocument', {
            'source': '''
                // 覆盖navigator.webdriver
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => false,
                });
                
                // 覆盖window.navigator.chrome
                window.navigator.chrome = {
                    runtime: {},
                };
                
                // 覆盖window.chrome
                window.chrome = {
                    runtime: {},
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
            self.log_message.emit("需要登录，请手动完成登录...")
            # 等待导航到原始URL或包含search-console的URL
            page.wait_for_url(lambda url: "search-console" in url or "search.google.com" in url, timeout=120000)
            self.log_message.emit("检测到登录成功，继续执行...")
        
        # 添加随机滚动
        for _ in range(random.randint(2, 4)):
            page.mouse.wheel(0, random.randint(100, 300))
            time.sleep(random.uniform(0.5, 1.5))
        
        # 截取第一个图表
        try:
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
            self.update_markdown_file(page_name, gsc_queries)
            
        except Exception as extract_error:
            self.log_message.emit(f"提取查询时出错: {str(extract_error)}")
    
    def update_markdown_file(self, page_name, gsc_queries):
        """更新或创建MD文件并填入GSC查询结果"""
        if not gsc_queries:
            self.log_message.emit("没有查询结果可以更新到MD文件")
            return
            
        # 创建目标MD文件名
        md_file_path = f"{page_name}.md"
        
        # 检查文件是否存在，如果不存在则复制模板
        if not os.path.exists(md_file_path):
            if self.md_template_path and os.path.exists(self.md_template_path):
                self.log_message.emit(f"复制模板文件 {self.md_template_path} 到 {md_file_path}")
                shutil.copy2(self.md_template_path, md_file_path)
            else:
                # 如果没有模板，创建一个基本结构
                self.log_message.emit(f"创建新的MD文件: {md_file_path}")
                with open(md_file_path, "w", encoding="utf-8") as file:
                    file.write(f"""
# {page_name}

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
        
        # 查找"### GSC热门查询"部分
        gsc_section_index = md_content.find("### GSC热门查询")
        next_section_index = md_content.find("###", gsc_section_index + 1)
        
        if gsc_section_index != -1:
            # 构建新内容
            new_content = ""
            for i, query in enumerate(gsc_queries):
                if i == 0:
                    new_content += f"- {i+1}.{query}\n"
                else:
                    new_content += f"{i+1}.{query}\n"
            
            # 插入新内容
            if next_section_index != -1:
                updated_content = md_content[:gsc_section_index + len("### GSC热门查询")] + "\n\n" + new_content + "\n" + md_content[next_section_index:]
            else:
                updated_content = md_content[:gsc_section_index + len("### GSC热门查询")] + "\n\n" + new_content
            
            # 保存更新后的内容
            with open(md_file_path, "w", encoding="utf-8") as file:
                file.write(updated_content)
            
            self.log_message.emit(f"成功将前 {len(gsc_queries)} 个GSC查询保存到 {md_file_path}")
        else:
            self.log_message.emit(f"在文件 {md_file_path} 中未找到'### GSC热门查询'部分")
    
    def process_ga(self, page, ga_url, page_name, ga_screenshot_path, screenshot_dir):
        """处理GA相关的任务"""
        try:
            self.log_message.emit(f"导航到GA4分析页面: {ga_url}")
            page.goto(ga_url, timeout=90000)
            
            # 添加5秒等待时间让页面完全加载
            self.log_message.emit("等待5秒让页面完全加载...")
            time.sleep(5)
            
            # 直接尝试定位GA4报表元素
            ga_selector = "body > ga-hybrid-app-root > ui-view-wrapper > div > app-root > div > div > ui-view-wrapper > div > ga-report-container > div > div > div > report-view > ui-view-wrapper > div > ui-view > ga-explorer-report > div > div > div > ga-card-list.explorer-card-list.ga-card-list.ng-star-inserted > div"
            
            self.log_message.emit("定位GA4报表元素...")
            
            try:
                # 等待元素出现
                ga_element = page.wait_for_selector(ga_selector, timeout=90000)
                
                if ga_element:
                    self.log_message.emit("找到GA4元素，正在截图...")
                    ga_element.screenshot(path=ga_screenshot_path)
                    self.log_message.emit(f"GA4截图已保存为: {ga_screenshot_path}")
            except Exception as wait_error:
                self.log_message.emit(f"等待GA4元素时出错: {str(wait_error)}")
                
                # 如果等待失败，尝试直接查询
                ga_element = page.query_selector(ga_selector)
                
                if ga_element:
                    self.log_message.emit("使用直接查询找到GA4元素，正在截图...")
                    ga_element.screenshot(path=ga_screenshot_path)
                    self.log_message.emit(f"GA4截图已保存为: {ga_screenshot_path}")
                else:
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
        url_group = QGroupBox("URL列表")
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
        
        url_group.setLayout(url_layout)
        task_layout.addWidget(url_group)
        
        # MD模板选择区域
        template_group = QGroupBox("Markdown模板")
        template_layout = QHBoxLayout()
        
        self.template_path_input = QLineEdit()
        self.template_path_input.setPlaceholderText("选择一个.md文件作为模板")
        self.template_path_input.setReadOnly(True)
        
        self.select_template_button = QPushButton("选择模板")
        
        template_layout.addWidget(self.template_path_input, 7)
        template_layout.addWidget(self.select_template_button, 3)
        
        template_group.setLayout(template_layout)
        task_layout.addWidget(template_group)
        
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
        self.select_template_button.clicked.connect(self.select_template)
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
        self.scrape_ga_checkbox = QCheckBox("抓取GA数据")
        self.scrape_ga_checkbox.setChecked(True)
        
        options_layout.addWidget(self.headless_checkbox)
        options_layout.addWidget(self.scrape_ga_checkbox)
        
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
        
    def select_template(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择Markdown模板", "", "Markdown文件 (*.md);;所有文件 (*)")
        if file_path:
            self.template_path_input.setText(file_path)
            self.settings.setValue("md_template", file_path)
            self.log_message(f"已选择模板文件: {file_path}")
            
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
        self.settings.setValue("scrape_ga", "true" if self.scrape_ga_checkbox.isChecked() else "false")
        
        QMessageBox.information(self, "设置", "设置已保存")
        self.log_message("设置已更新")
        
    def load_settings(self):
        # 加载设置
        self.chrome_profile_input.setText(self.settings.value("chrome_profile", ""))
        self.screenshot_dir_input.setText(self.settings.value("screenshot_dir", "screenshots"))
        self.template_path_input.setText(self.settings.value("md_template", ""))
        self.headless_checkbox.setChecked(self.settings.value("headless_mode", "false") == "true")
        self.scrape_ga_checkbox.setChecked(self.settings.value("scrape_ga", "true") == "true")
        
    def start_task(self):
        # 获取URL列表
        urls_text = self.url_input.toPlainText().strip()
        if not urls_text:
            QMessageBox.warning(self, "错误", "请输入至少一个URL")
            return
            
        # 分割URL
        urls = [url.strip() for url in urls_text.split("\n") if url.strip()]
        
        # 获取模板路径
        template_path = self.template_path_input.text()
        
        # 更新UI状态
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)
        self.progress_bar.setValue(0)
        self.progress_label.setText(f"处理中... (0/{len(urls)})")
        
        # 重定向标准输出到日志窗口
        self.stdout_redirect = LogRedirector(self.log_text)
        sys.stdout = self.stdout_redirect
        
        # 创建并启动工作线程
        self.worker = RPAWorker(urls, template_path, self.settings)
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
        if success:
            self.log_message(f"成功完成URL: {url}")
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