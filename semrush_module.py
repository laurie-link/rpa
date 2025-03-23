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
        
        # 取消截图保存，因为用户不需要
        # 直接进入提取数据步骤
        
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
        # 不再保存错误截图
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
                        
                        // 只有当两者都存在值时添加，并且跳过"All keywords"和"PPC Keyword Tool"
                        if (text && value && 
                            text !== "All keywords" && 
                            !text.includes("PPC")) {
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
    """提取SEMrush主要关键词数据 - 修复Volume和KD提取问题"""
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
                
                // 尝试方法1: 直接提取关键词行
                try {
                    // 获取所有表格行
                    const tableRows = Array.from(tableContainer.querySelectorAll('[role="row"], tr, .sm-table-layout__row'));
                    
                    // 分析每一行
                    for (let i = 0; i < tableRows.length; i++) {
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
                        
                        // 跳过工具名称和UI元素
                        if (keyword === 'PPC Keyword Tool' || 
                            keyword.includes('dashboard') || 
                            keyword.includes('profile') ||
                            keyword.includes('Domain') ||
                            keyword.includes('Projects') ||
                            keyword.includes('Analytics')) {
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
                        
                        // 添加到结果
                        rows.push({
                            keyword: keyword,
                            volume: volume,
                            kd: kd
                        });
                    }
                } catch (e) {
                    // 出错时继续尝试下一种方法
                }
                
                // 如果方法1没有找到数据，尝试方法2: 直接查找表格单元格
                if (rows.length === 0) {
                    try {
                        // 查找所有关键词单元格
                        const keywordCells = Array.from(document.querySelectorAll('a span, a'))
                            .filter(el => el.textContent.trim().length > 3);
                        
                        for (let i = 0; i < keywordCells.length; i++) {
                            const keywordCell = keywordCells[i];
                            const keyword = keywordCell.textContent.trim();
                            
                            // 跳过工具名称和UI元素
                            if (keyword === 'PPC Keyword Tool' || 
                                keyword.includes('dashboard') || 
                                keyword.includes('profile')) {
                                continue;
                            }
                            
                            // 尝试找到包含该关键词的行
                            let rowElement = keywordCell;
                            while (rowElement && 
                                  !rowElement.matches('[role="row"], tr, .sm-table-layout__row')) {
                                rowElement = rowElement.parentElement;
                                if (!rowElement) break;
                            }
                            
                            let volume = "0";
                            let kd = "n/a";
                            
                            // 如果找到了行，尝试获取相关数据
                            if (rowElement) {
                                // 获取所有单元格
                                const cells = Array.from(rowElement.querySelectorAll('[role="cell"], td, .sm-table-layout__cell, div'));
                                
                                // 遍历单元格寻找搜索量和KD
                                for (const cell of cells) {
                                    const text = cell.textContent.trim();
                                    
                                    // 识别搜索量
                                    if (/^[0-9,.]+$/.test(text) || /^[0-9,.]+[KMB]$/.test(text)) {
                                        volume = text;
                                    }
                                    
                                    // 识别KD
                                    if (text.endsWith('%') || 
                                        (/^[0-9]+$/.test(text) && parseInt(text) >= 0 && parseInt(text) <= 100)) {
                                        kd = text.endsWith('%') ? text : text + '%';
                                    }
                                }
                            }
                            
                            // 添加到结果
                            rows.push({
                                keyword: keyword,
                                volume: volume,
                                kd: kd
                            });
                        }
                    } catch (e) {
                        // 出错时继续尝试
                    }
                }
                
                // 返回过滤掉PPC相关内容的行数据
                return rows.filter(row => 
                    row.keyword && 
                    row.keyword !== 'PPC Keyword Tool' &&
                    !row.keyword.includes('PPC') &&
                    !row.keyword.includes('dashboard') &&
                    !row.keyword.includes('profile')
                );
            }
        """)
        
        log_message_callback(f"提取到 {len(keyword_rows)} 个关键词数据行")
        
        # 记录前10个关键词数据
        for i, row in enumerate(keyword_rows[:10]):
            log_message_callback(f"关键词数据 {i+1}: {row['keyword']} - Volume:{row['volume']} - KD:{row['kd']}")
        
        return keyword_rows
        
    except Exception as e:
        log_message_callback(f"提取SEMrush关键词数据时出错: {str(e)}")
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