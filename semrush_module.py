import time
import os
import re

def process_semrush(log_message_callback, page, page_name, screenshot_dir):
    """处理SEMrush关键词数据提取
    
    Args:
        log_message_callback: 用于记录日志的回调函数
        page: Playwright页面对象
        page_name: 页面名称
        screenshot_dir: 截图保存目录
    """
    try:
        # 导航到SEMrush登录页面
        log_message_callback("导航到SEMrush登录页面...")
        login_url = "https://tool.seotools8.com/#/login"
        page.goto(login_url, timeout=30000)
        
        # 检查是否已经登录
        if "login" in page.url:
            log_message_callback("需要登录SEMrush...")
            
            # 输入用户名和密码
            username_selector = "input[type='text']"
            password_selector = "input[type='password']"
            
            log_message_callback("填写登录信息...")
            page.fill(username_selector, "khy7788")
            page.fill(password_selector, "123123")
            
            # 点击登录按钮 - 使用更可靠的选择器
            login_button_selector = "button[type='submit']"
            log_message_callback("点击登录按钮...")
            page.click(login_button_selector)
            
            # 等待登录完成
            page.wait_for_load_state("networkidle", timeout=30000)
            
            # 点击选择账号登录按钮 - 如果需要
            if page.query_selector("button.q-btn:has-text('登录')"):
                log_message_callback("点击选择账号登录按钮...")
                page.click("button.q-btn:has-text('登录')")
                page.wait_for_load_state("networkidle", timeout=30000)
        
        # 构建Keywords Magic Tool URL
        search_keyword = page_name.replace("-", "+")
        semrush_url = f"https://tool-sem.seotools8.com/analytics/keywordmagic/?q={search_keyword}&db=us"
        
        log_message_callback(f"导航到SEMrush Keywords Magic Tool页面: {semrush_url}")
        page.goto(semrush_url, timeout=60000)
        
        # 等待页面加载
        log_message_callback("等待SEMrush页面加载...")
        page.wait_for_load_state("networkidle", timeout=30000)
        
        # 增加额外等待时间，确保页面和JavaScript完全加载
        log_message_callback("额外等待8秒确保页面完全加载...")
        time.sleep(8)
        
        # 截取SEMrush页面
        semrush_screenshot_path = os.path.join(screenshot_dir, f"semrush-{page_name}.png")
        page.screenshot(path=semrush_screenshot_path, full_page=True)
        log_message_callback(f"SEMrush页面完整截图已保存为: {semrush_screenshot_path}")
        
        # 简单检查页面是否有数据
        log_message_callback("检查页面是否有关键词数据...")
        
        # 提取边栏数据
        log_message_callback("提取SEMrush边栏数据...")
        sidebar_data = extract_semrush_sidebar_data(log_message_callback, page)
        
        # 提取主要关键词数据
        log_message_callback("提取SEMrush主要关键词数据...")
        keyword_data = extract_semrush_keyword_data(log_message_callback, page)
        
        # 检验提取的数据
        if keyword_data and len(keyword_data) > 0:
            # 整合数据并更新markdown文件
            log_message_callback("整合SEMrush数据并更新markdown文件...")
            update_semrush_markdown(log_message_callback, page_name, sidebar_data, keyword_data)
            return True
        else:
            log_message_callback("未找到有效的SEMrush关键词数据")
            return False
            
    except Exception as semrush_error:
        log_message_callback(f"处理SEMrush数据时出错: {str(semrush_error)}")
        # 截取当前页面状态作为错误记录
        try:
            error_screenshot_path = os.path.join(screenshot_dir, f"semrush-error-{page_name}.png")
            page.screenshot(path=error_screenshot_path)
            log_message_callback(f"错误状态截图已保存为: {error_screenshot_path}")
        except:
            pass
        return False

def extract_semrush_sidebar_data(log_message_callback, page):
    """提取SEMrush边栏数据"""
    try:
        # 使用JavaScript评估提取边栏数据
        sidebar_data = page.evaluate("""
            () => {
                const result = [];
                
                // 获取所有关键词组标题和数量
                const groups = document.querySelectorAll(".sm-group-content");
                
                for (const group of groups) {
                    const textElement = group.querySelector(".sm-group-content__text");
                    const valueElement = group.querySelector(".sm-group-content__value");
                    
                    if (textElement && valueElement) {
                        const text = textElement.textContent.trim();
                        const value = valueElement.textContent.trim();
                        
                        // 只有当两者都存在值时添加，并且跳过"All keywords"
                        if (text && value && text !== "All keywords") {
                            result.push({
                                text: text,
                                value: value
                            });
                        }
                    }
                }
                
                return result;
            }
        """)
        
        log_message_callback(f"提取到 {len(sidebar_data)} 个SEMrush边栏数据项")
        for i, item in enumerate(sidebar_data):
            log_message_callback(f"边栏数据 {i+1}: {item['text']} - {item['value']}")
            
        return sidebar_data
        
    except Exception as e:
        log_message_callback(f"提取SEMrush边栏数据时出错: {str(e)}")
        return []

def extract_semrush_keyword_data(log_message_callback, page):
    """提取SEMrush主要关键词数据 - 基于页面观察优化"""
    try:
        # 创建调试目录
        debug_dir = "debug_screenshots"
        os.makedirs(debug_dir, exist_ok=True)
        
        # 直接提取关键词数据 - 使用简化但有效的方法
        log_message_callback("使用简化方法直接提取关键词数据...")
        keyword_data_direct = []
        
        # 直接从页面提取文本内容
        page_text = page.content()
        
        # 直接保存页面源码进行调试
        with open(os.path.join(debug_dir, "semrush_page_source.html"), "w", encoding="utf-8") as f:
            f.write(page_text)
        
        # 使用更简单的Javascript来提取数据
        raw_data = page.evaluate("""
            () => {
                // 找到所有关键词链接
                const keywordLinks = Array.from(document.querySelectorAll('a span'));
                const volumeCells = Array.from(document.querySelectorAll('[role="cell"]'));
                
                // 提取关键词文本
                const keywordTexts = keywordLinks.map(link => link.textContent.trim())
                    .filter(text => text.length > 3 && !text.includes('dashboard') && !text.includes('profile'));
                
                // 提取可能的搜索量数据 (通常是数字格式)
                const volumeData = volumeCells.map(cell => cell.textContent.trim())
                    .filter(text => /^[0-9,.]+$/.test(text) || /^[0-9,.]+[KMB]$/.test(text));
                
                // 提取可能的KD数据 (通常带有%)
                const kdData = volumeCells.map(cell => cell.textContent.trim())
                    .filter(text => text.endsWith('%') || text === 'n/a');
                
                return {
                    keywords: keywordTexts,
                    volumes: volumeData,
                    kds: kdData
                };
            }
        """)
        
        log_message_callback(f"找到 {len(raw_data['keywords'])} 个关键词")
        log_message_callback(f"找到 {len(raw_data['volumes'])} 个搜索量数据")
        log_message_callback(f"找到 {len(raw_data['kds'])} 个KD数据")
        
        # 构建关键词数据
        max_items = min(len(raw_data['keywords']), 200)  # 限制处理项数
        for i in range(max_items):
            keyword = raw_data['keywords'][i] if i < len(raw_data['keywords']) else ""
            volume = raw_data['volumes'][i] if i < len(raw_data['volumes']) else "0"
            kd = raw_data['kds'][i] if i < len(raw_data['kds']) else "n/a"
            
            # 确保这是有效的关键词
            if keyword and not any(ui_term in keyword.lower() for ui_term in ["profile", "dashboard", "projects", "analytics"]):
                keyword_data_direct.append({
                    "keyword": keyword,
                    "volume": volume,
                    "kd": kd
                })
        
        # 使用一个额外的截图确保我们已经处理了页面
        page.screenshot(path=os.path.join(debug_dir, "after_extraction.png"))
        
        # 最后的过滤 - 确保关键词与页面主题相关
        filtered_data = []
        for item in keyword_data_direct:
            keyword = item['keyword'].lower()
            # 检查相关性 - 包含主题相关词汇
            if "spotify" in keyword or "premium" in keyword or "crack" in keyword or "pc" in keyword:
                filtered_data.append(item)
                
        log_message_callback(f"提取并过滤得到 {len(filtered_data)} 个有效的关键词数据项")
        
        # 记录前10个关键词数据
        for i, item in enumerate(filtered_data[:10]):
            log_message_callback(f"关键词数据 {i+1}: {item['keyword']} - Volume:{item['volume']} - KD:{item['kd']}")
            
        return filtered_data
        
    except Exception as e:
        log_message_callback(f"提取SEMrush关键词数据时出错: {str(e)}")
        log_message_callback(f"错误详情: {e}")
        
        # 保存错误截图
        try:
            page.screenshot(path=os.path.join("debug_screenshots", "extraction_error.png"))
        except:
            pass
            
        return []

def update_semrush_markdown(log_message_callback, page_name, sidebar_data, keyword_data):
    """更新markdown文件中的SEMrush数据"""
    if not sidebar_data and not keyword_data:
        log_message_callback("没有SEMrush数据可以更新到MD文件")
        return
        
    # 创建目标MD文件名
    md_file_path = f"{page_name}.md"
    
    # 检查文件是否存在
    if not os.path.exists(md_file_path):
        log_message_callback(f"MD文件 {md_file_path} 不存在，创建新文件")
        with open(md_file_path, "w", encoding="utf-8") as file:
            file.write(f"""
# {page_name}

## 关键词来源

### Google 搜索下拉框

### 相关搜索

### GSC热门查询

### 相关问题

### SEMrush
""")
    
    # 读取现有内容
    try:
        with open(md_file_path, "r", encoding="utf-8") as file:
            md_content = file.read()
    except Exception as e:
        log_message_callback(f"读取MD文件时出错: {str(e)}")
        return
    
    # 查找SEMrush部分
    semrush_header = "### SEMrush"
    semrush_index = md_content.find(semrush_header)
    
    # 如果未找到SEMrush部分，添加它
    if semrush_index == -1:
        log_message_callback(f"在文件 {md_file_path} 中未找到'{semrush_header}'部分，添加该部分")
        md_content += f"\n\n{semrush_header}\n"
        semrush_index = md_content.find(semrush_header)
    
    # 构建Markdown表格
    table_content = "\n\n| main words | main word count | Key words | Volume | Keyword Difficulty |\n"
    table_content += "| --- | --- | --- | --- | --- |\n"
    
    # 确定实际有多少行数据
    max_rows = max(len(sidebar_data), len(keyword_data))
    
    # 如果没有数据，至少添加一行空行
    if max_rows == 0:
        table_content += "| | | | | |\n"
    else:
        # 填充表格数据
        for i in range(max_rows):
            main_word = sidebar_data[i]['text'] if i < len(sidebar_data) else ""
            main_word_count = sidebar_data[i]['value'] if i < len(sidebar_data) else ""
            keyword = keyword_data[i]['keyword'] if i < len(keyword_data) else ""
            volume = keyword_data[i]['volume'] if i < len(keyword_data) else ""
            kd = keyword_data[i]['kd'] if i < len(keyword_data) else ""
            
            table_content += f"| {main_word} | {main_word_count} | {keyword} | {volume} | {kd} |\n"
    
    # 查找下一部分的开始
    next_section_index = md_content.find("###", semrush_index + len(semrush_header))
    
    # 插入表格内容
    if next_section_index != -1:
        updated_content = md_content[:semrush_index + len(semrush_header)] + table_content + "\n" + md_content[next_section_index:]
    else:
        updated_content = md_content[:semrush_index + len(semrush_header)] + table_content
    
    # 保存更新后的内容
    try:
        with open(md_file_path, "w", encoding="utf-8") as file:
            file.write(updated_content)
        
        log_message_callback(f"成功将SEMrush数据保存到 {md_file_path}")
    except Exception as e:
        log_message_callback(f"保存MD文件时出错: {str(e)}")