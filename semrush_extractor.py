"""
SEMrush数据提取模块 - 用于从SEMrush获取关键词数据并保存到Markdown文件

此模块提供以下功能:
1. 登录SEMrush并导航到Keywords Magic Tool
2. 提取边栏数据和主要关键词数据
3. 将数据保存为Markdown表格格式
"""

import os
import time

def process_semrush(self, playwright, page_name, screenshot_dir):
    """处理SEMrush数据提取任务"""
    try:
        # 启动无痕浏览器实例
        self.log_message.emit("以无痕模式启动浏览器进行SEMrush数据提取...")
        browser_context = playwright.chromium.launch(
            headless=self.settings.value("headless_mode", "false") == "true"
        )
        
        # 创建上下文和页面
        context = browser_context.new_context()
        page = context.new_page()
        
        try:
            # 导航到SEMrush登录页面
            self.log_message.emit("导航到SEMrush登录页面...")
            page.goto("https://tool.seotools8.com/#/login", timeout=60000)
            
            # 填写登录信息
            self.log_message.emit("填写登录信息...")
            page.fill("input[type='text']", "khy7788")
            page.fill("input[type='password']", "123123")
            
            # 点击登录按钮
            login_button_selector = "#q-app > div > div.content > div > form > div:nth-child(4) > button > span.q-btn__wrapper.col.row.q-anchor--skip > span"
            self.log_message.emit("点击登录按钮...")
            page.click(login_button_selector)
            
            # 等待页面加载
            time.sleep(2)
            
            # 点击选择账号登录按钮
            account_selector = "#q-app > div > div.content > div > div.navigation-col1 > div:nth-child(1) > div.q-bar.row.no-wrap.items-center.q-bar--standard.q-bar--light > button > span.q-btn__wrapper.col.row.q-anchor--skip > span"
            self.log_message.emit("点击选择账号登录按钮...")
            
            try:
                # 检查选择账号按钮是否存在
                if page.is_visible(account_selector):
                    page.click(account_selector)
                    self.log_message.emit("已点击选择账号按钮")
                else:
                    self.log_message.emit("选择账号按钮不可见，可能已经登录")
            except Exception as e:
                self.log_message.emit(f"点击选择账号按钮时出错: {str(e)}")
                # 截图记录当前页面状态
                error_path = os.path.join(screenshot_dir, f"semrush-login-error-{page_name}.png")
                page.screenshot(path=error_path)
                self.log_message.emit(f"登录错误状态截图已保存: {error_path}")
            
            # 等待页面加载
            time.sleep(3)
            
            # 构建SEMrush Keywords Magic Tool URL
            search_term = page_name.replace("-", "+")
            semrush_url = f"https://tool-sem.seotools8.com/analytics/keywordmagic/?q={search_term}&db=us"
            
            self.log_message.emit(f"导航到SEMrush关键词工具页面: {semrush_url}")
            page.goto(semrush_url, timeout=90000)
            
            # 等待页面加载
            self.log_message.emit("等待页面加载...")
            time.sleep(5)
            
            # 保存页面截图用于调试
            semrush_screenshot_path = os.path.join(screenshot_dir, f"semrush-{page_name}.png")
            page.screenshot(path=semrush_screenshot_path)
            self.log_message.emit(f"SEMrush页面截图已保存为: {semrush_screenshot_path}")
            
            # 提取边栏数据
            self.log_message.emit("开始提取SEMrush边栏数据...")
            sidebar_data = self.extract_semrush_sidebar_data(page)
            
            # 提取主要关键词数据
            self.log_message.emit("开始提取SEMrush主要关键词数据...")
            main_data = self.extract_semrush_main_data(page)
            
            # 将数据组合成表格并保存到markdown文件
            self.log_message.emit("组合数据并保存到markdown文件...")
            self.save_semrush_data_to_markdown(page_name, sidebar_data, main_data)
            
        except Exception as e:
            self.log_message.emit(f"处理SEMrush数据时出错: {str(e)}")
            # 截图记录错误状态
            error_path = os.path.join(screenshot_dir, f"semrush-error-{page_name}.png")
            try:
                page.screenshot(path=error_path)
                self.log_message.emit(f"错误状态截