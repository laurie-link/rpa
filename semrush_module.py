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
        
        # 提取边栏数据
        log_message_callback("提取SEMrush边栏数据(最多20条)...")
        sidebar_data = extract_semrush_sidebar_data(log_message_callback, page)
        
        # 提取主要关键词数据
        log_message_callback("提取SEMrush主要关键词数据(最多20条)...")
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
        return False

def extract_semrush_sidebar_data(log_message_callback, page):
    """提取SEMrush边栏数据，最多返回20条"""
    try:
        # 使用JavaScript评估提取边栏数据
        sidebar_data = page.evaluate("""
            () => {
                const result = [];
                
                // 获取所有关键词组标题和数量
                const groups = document.querySelectorAll(".sm-group-content");
                let count = 0;
                
                for (const group of groups) {
                    // 只提取前20条数据
                    if (count >= 20) break;
                    
                    const textElement = group.querySelector(".sm-group-content__text");
                    const valueElement = group.querySelector(".sm-group-content__value");
                    
                    if (textElement && valueElement) {
                        const text = textElement.textContent.trim();
                        const value = valueElement.textContent.trim();
                        
                        // 只有当两者都存在值时添加，并且跳过"All keywords"和"PPC Keyword Tool"
                        if (text && value && 
                            text !== "All keywords" && 
                            !text.includes("PPC")) {
                            result.push({
                                text: text,
                                value: value
                            });
                            count++;
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
    """提取SEMrush主要关键词数据，最多返回20条，并改进过滤逻辑"""
    try:
        # 使用更精确的JavaScript提取每一行的完整数据
        keyword_rows = page.evaluate("""
            () => {
                // 用于存储所有关键词行的数据
                const rows = [];
                
                // 找到表格或关键词容器
                const tableContainer = document.querySelector('.sm-table-layout') || 
                                      document.querySelector('table') || 
                                      document.body;
                
                // 用于计数已提取的有效关键词数量
                let keywordCount = 0;
                const maxKeywords = 100; // 先提取更多，后面再过滤
                
                // 常见UI元素和导航菜单项列表
                const uiTerms = [
                    'Features', 'Pricing', 'Help Center', 'What\\'s New', 'Webinars', 
                    'Insights', 'Hire', 'Academy', 'Top Websites', 'Content Marketing', 
                    'Local Marketing', 'About Us', 'Login', 'Sign Up', 'Contact', 
                    'Support', 'Documentation', 'Blog', 'API', 'Tools'
                ];
                
                // 关键词有效性检查函数
                const isValidKeyword = (text) => {
                    if (!text || text.length < 3) return false;
                    
                    // 跳过工具名称和UI元素
                    if (text === 'PPC Keyword Tool' ||
                        text.includes('dashboard') ||
                        text.includes('profile') ||
                        text.includes('Domain') ||
                        text.includes('Projects') ||
                        text.includes('Analytics')) {
                        return false;
                    }
                    
                    // 跳过过长的文本（可能是描述性文本）
                    if (text.length > 80) return false;
                    
                    // 跳过包含HTML标签的文本
                    if (text.includes('<') || text.includes('>')) return false;
                    
                    // 检查是否是UI元素
                    for (const term of uiTerms) {
                        if (text === term || text.startsWith(term) || 
                            text.toLowerCase() === term.toLowerCase() || 
                            text.toLowerCase().startsWith(term.toLowerCase())) {
                            return false;
                        }
                    }
                    
                    // 跳过可能是URL或路径的文本
                    if (text.includes('/') || text.includes('http')) return false;
                    
                    // 跳过首字母大写的单词（可能是导航项）
                    if (/^[A-Z][a-z]+$/.test(text)) return false;
                    
                    // 跳过包含过多符号的文本 - 修复转义问题
                    const symbolCount = (text.match(/[!@#$%^&*()_+=\\[\\]{};':"\\|,.<>\\/?-]/g) || []).length;
                    if (symbolCount > 2) return false;
                    
                    // 跳过只有单个字符的文本
                    if (text.trim().split(/\\s+/).length < 2 && text.length < 10) return false;
                    
                    return true;
                };
                
                // 尝试方法1: 直接提取关键词行
                try {
                    // 获取所有表格行
                    const tableRows = Array.from(tableContainer.querySelectorAll('[role="row"], tr, .sm-table-layout__row'));
                    
                    // 分析每一行
                    for (let i = 0; i < tableRows.length && keywordCount < maxKeywords; i++) {
                        const row = tableRows[i];
                        
                        // 跳过表头行
                        const isHeader = row.getAttribute('role') === 'rowheader' || 
                                        row.querySelector('th') !== null || 
                                        row.classList.contains('sm-table-layout__header-row');
                        
                        if (isHeader) {
                            continue;
                        }
                        
                        // 在当前行中获取关键词、搜索量和KD
                        const keywordElement = row.querySelector('a span') || row.querySelector('a');
                        
                        if (!keywordElement) {
                            continue; // 没有关键词元素，跳过
                        }
                        
                        const keyword = keywordElement.textContent.trim();
                        
                        // 使用改进的关键词过滤逻辑
                        if (!isValidKeyword(keyword)) {
                            continue;
                        }
                        
                        // 获取单元格
                        const cells = Array.from(row.querySelectorAll('[role="cell"], td, .sm-table-layout__cell'));
                        
                        // 提取搜索量和KD (通常在第5个和第9个单元格)
                        let volume = "0";
                        let kd = "n/a";
                        
                        // 尝试从第5个单元格获取搜索量
                        if (cells.length >= 5) {
                            const volumeText = cells[4].textContent.trim();
                            if (/^[0-9,.]+$/.test(volumeText) || /^[0-9,.]+[KMB]$/.test(volumeText)) {
                                volume = volumeText;
                            }
                        }
                        
                        // 尝试从第9个单元格获取KD
                        if (cells.length >= 9) {
                            const kdText = cells[8].textContent.trim();
                            if (kdText.endsWith('%') || kdText === 'n/a') {
                                kd = kdText;
                            } else if (/^[0-9]+$/.test(kdText)) {
                                kd = kdText + '%';  // 添加百分号
                            }
                        }
                        
                        // 如果上述方法未能提取搜索量和KD，尝试遍历所有单元格
                        if (volume === "0" || kd === "n/a") {
                            for (let j = 0; j < cells.length; j++) {
                                const text = cells[j].textContent.trim();
                                
                                // 识别搜索量
                                if (volume === "0" && 
                                    (/^[0-9,.]+$/.test(text) || /^[0-9,.]+[KMB]$/.test(text))) {
                                    volume = text;
                                }
                                
                                // 识别KD
                                if (kd === "n/a" && 
                                    (text.endsWith('%') || 
                                     (/^[0-9]+$/.test(text) && parseInt(text) >= 0 && parseInt(text) <= 100))) {
                                    kd = text.endsWith('%') ? text : text + '%';
                                }
                            }
                        }
                        
                        // 添加到结果并递增计数器
                        rows.push({
                            keyword: keyword,
                            volume: volume,
                            kd: kd
                        });
                        keywordCount++;
                    }
                } catch (e) {
                    // 出错时继续尝试下一种方法
                    console.error("方法1出错:", e);
                }
                
                // 过滤结果，确保没有UI元素
                const filteredRows = rows.filter(row => 
                    isValidKeyword(row.keyword) &&
                    row.keyword !== 'PPC Keyword Tool' &&
                    !row.keyword.includes('PPC')
                );
                
                // 返回最多20条记录
                return filteredRows.slice(0, 20);
            }
        """)
        
        log_message_callback(f"提取到 {len(keyword_rows)} 个关键词数据行")
        
        # 记录所有提取的关键词数据
        for i, row in enumerate(keyword_rows):
            log_message_callback(f"关键词数据 {i+1}: {row['keyword']} - Volume:{row['volume']} - KD:{row['kd']}")
        
        return keyword_rows
        
    except Exception as e:
        log_message_callback(f"提取SEMrush关键词数据时出错: {str(e)}")
        log_message_callback(f"错误详情: {e}")
        return []

def update_semrush_markdown(log_message_callback, page_name, sidebar_data, keyword_data):
    """更新markdown文件中的SEMrush数据 - 保持原来的逻辑但修复PPC问题"""
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
    
    # 构建Markdown表格 - 保持原来的逻辑
    table_content = "\n\n| main words | main word count | Key words | Volume | Keyword Difficulty |\n"
    table_content += "| --- | --- | --- | --- | --- |\n"
    
    # 确定实际有多少行数据
    max_rows = max(len(sidebar_data), len(keyword_data))
    
    # 如果没有数据，至少添加一行空行
    if max_rows == 0:
        table_content += "| | | | | |\n"
    else:
        # 填充表格数据 - 确保过滤掉PPC内容
        filtered_sidebar = [item for item in sidebar_data if not item['text'].startswith('PPC')]
        filtered_keywords = [item for item in keyword_data if not item['keyword'].startswith('PPC')]
        
        max_rows = max(len(filtered_sidebar), len(filtered_keywords))
        
        for i in range(max_rows):
            main_word = filtered_sidebar[i]['text'] if i < len(filtered_sidebar) else ""
            main_word_count = filtered_sidebar[i]['value'] if i < len(filtered_sidebar) else ""
            keyword = filtered_keywords[i]['keyword'] if i < len(filtered_keywords) else ""
            volume = filtered_keywords[i]['volume'] if i < len(filtered_keywords) else ""
            kd = filtered_keywords[i]['kd'] if i < len(filtered_keywords) else ""
            
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