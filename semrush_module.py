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
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            # 每次重试都重新导航到登录页面
            log_message_callback(f"导航到SEMrush登录页面...(尝试 {retry_count + 1}/{max_retries})")
            login_url = "https://tool.seotools8.com/#/login"
            page.goto(login_url, timeout=30000)
            
            # 进行登录
            login_semrush(log_message_callback, page)
            
            # 构建Keywords Magic Tool URL
            search_keyword = page_name.replace("-", "+")
            semrush_url = f"https://tool-sem.seotools8.com/analytics/keywordmagic/?q={search_keyword}&db=us&gsort=volume_desc"
            
            log_message_callback(f"导航到SEMrush Keywords Magic Tool页面: {semrush_url}")
            page.goto(semrush_url, timeout=60000)
            
            # 等待页面加载
            log_message_callback("等待SEMrush页面加载...")
            page.wait_for_load_state("networkidle", timeout=30000)
            
            # 检查是否出现错误页面
            if check_semrush_error_page(log_message_callback, page):
                log_message_callback("检测到SEMrush错误页面，将进行重试...")
                retry_count += 1
                continue
                
            # 等待关键词元素出现
            log_message_callback("等待SEMrush关键词元素出现...")
            try:
                # 尝试等待关键词表格行或关键词组元素出现
                page.wait_for_selector(".sm-table-layout__row, [role='row'], tr, .sm-group-content", 
                                     state="visible", timeout=60000)
                log_message_callback("SEMrush关键词元素已出现，继续处理...")
            except Exception as wait_error:
                log_message_callback(f"等待元素超时，将检查页面状态: {str(wait_error)}")
                
                # 检查是否是无数据页面
                if check_no_data_page(log_message_callback, page):
                    log_message_callback("检测到此关键词没有搜索量数据，创建空记录...")
                    # 创建空数据并返回
                    update_semrush_markdown(log_message_callback, page_name, [], [])
                    return True
                    
                # 检查是否是错误页面
                if check_semrush_error_page(log_message_callback, page):
                    log_message_callback("检测到SEMrush错误页面，将进行重试...")
                    retry_count += 1
                    continue
            
            # 提取边栏数据
            log_message_callback("提取SEMrush边栏数据(最多20条)...")
            sidebar_data = extract_semrush_sidebar_data(log_message_callback, page)
            
            # 提取主要关键词数据
            log_message_callback("提取SEMrush主要关键词数据(最多20条)...")
            keyword_data = extract_semrush_keyword_data(log_message_callback, page)
            
            # 检验提取的数据
            if (keyword_data and len(keyword_data) > 0) or (sidebar_data and len(sidebar_data) > 0):
                # 整合数据并更新markdown文件
                log_message_callback("整合SEMrush数据并更新markdown文件...")
                update_semrush_markdown(log_message_callback, page_name, sidebar_data, keyword_data)
                return True
            else:
                log_message_callback("未找到有效的SEMrush关键词数据，将尝试重试...")
                # 截取当前页面快照以便调试
                screenshot_path = os.path.join(screenshot_dir, f"semrush-error-{page_name}.png")
                page.screenshot(path=screenshot_path)
                log_message_callback(f"已保存错误页面截图: {screenshot_path}")
                
                retry_count += 1
                
        except Exception as semrush_error:
            log_message_callback(f"处理SEMrush数据时出错: {str(semrush_error)}")
            # 截取当前页面快照以便调试
            try:
                screenshot_path = os.path.join(screenshot_dir, f"semrush-exception-{page_name}.png")
                page.screenshot(path=screenshot_path)
                log_message_callback(f"已保存异常页面截图: {screenshot_path}")
            except:
                pass
                
            retry_count += 1
            log_message_callback(f"将进行第 {retry_count}/{max_retries} 次重试...")
    
    # 所有重试都失败
    log_message_callback(f"在 {max_retries} 次尝试后仍未能成功获取SEMrush数据")
    # 创建空数据以避免完全失败
    update_semrush_markdown(log_message_callback, page_name, [], [])
    return False

def login_semrush(log_message_callback, page):
    """登录SEMrush账号"""
    # 检查是否已经登录
    if "login" not in page.url and "#/login" not in page.url:
        log_message_callback("似乎已经登录SEMrush，检查会话状态...")
        # 尝试访问一个需要登录的页面来验证会话
        try:
            page.goto("https://tool.seotools8.com/#/dashboard", timeout=30000)
            page.wait_for_load_state("networkidle", timeout=30000)
            
            # 如果没有重定向到登录页面，说明已经登录
            if "login" not in page.url and "#/login" not in page.url:
                log_message_callback("SEMrush会话有效，无需重新登录")
                return True
        except:
            log_message_callback("会话检查失败，将尝试重新登录")
    
    log_message_callback("需要登录SEMrush...")
    
    # 确保在登录页面
    if "login" not in page.url and "#/login" not in page.url:
        log_message_callback("重定向到登录页面...")
        page.goto("https://tool.seotools8.com/#/login", timeout=30000)
        page.wait_for_load_state("networkidle", timeout=30000)
    
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
    
    # 验证登录状态
    if "login" in page.url or "#/login" in page.url:
        log_message_callback("登录失败，可能是账号密码错误或服务器问题")
        return False
    
    log_message_callback("SEMrush登录成功")
    return True

def check_semrush_error_page(log_message_callback, page):
    """检查是否是SEMrush错误页面"""
    try:
        # 检查常见错误页面内容
        error_content = page.evaluate("""
            () => {
                // 检查"Something went wrong"错误
                if (document.body.innerText.includes('Something went wrong') || 
                    document.body.innerText.includes('went wrong')) {
                    return 'something_went_wrong';
                }
                
                // 检查400错误页面
                if (document.body.innerText.includes('400') && 
                    (document.body.innerText.includes('登录已失效') || 
                     document.body.innerText.includes('失效'))) {
                    return 'login_expired';
                }
                
                // 检查401/403错误
                if (document.body.innerText.includes('401') || 
                    document.body.innerText.includes('403') || 
                    document.body.innerText.includes('Unauthorized') ||
                    document.body.innerText.includes('Forbidden')) {
                    return 'unauthorized';
                }
                
                // 检查其他一般错误消息
                if (document.body.innerText.includes('error') || 
                    document.body.innerText.includes('Error') ||
                    document.body.innerText.includes('失败') ||
                    document.body.innerText.includes('错误')) {
                    return 'general_error';
                }
                
                return false;
            }
        """)
        
        if error_content:
            log_message_callback(f"检测到SEMrush错误页面: {error_content}")
            return True
        
        return False
    except Exception as e:
        log_message_callback(f"检查错误页面时发生异常: {str(e)}")
        return False

def check_no_data_page(log_message_callback, page):
    """检查是否是无数据页面"""
    try:
        # 检查常见的无数据指示
        no_data_indicators = page.evaluate("""
            () => {
                // 检查是否显示了空结果或更新指示器
                const hasUpdateMetrics = document.body.innerText.includes('Update metrics');
                const hasVolume = document.body.innerText.includes('Volume') && 
                                  document.body.innerText.includes('Keyword Difficulty');
                
                // 检查实际数据
                const volumeElement = document.querySelector('.KWO__metrics-volume, .KWO__volume');
                const hasActualVolume = volumeElement && volumeElement.innerText.match(/[0-9,.]+[KMB]?/);
                
                // 如果有更新指示器和搜索量字段，但没有实际数据
                if (hasUpdateMetrics && hasVolume && !hasActualVolume) {
                    return 'no_volume';
                }
                
                // 检查是否有明确的"没有数据"消息
                if (document.body.innerText.includes('No data') || 
                    document.body.innerText.includes('没有数据') || 
                    document.body.innerText.includes('0 results')) {
                    return 'no_data_message';
                }
                
                return false;
            }
        """)
        
        if no_data_indicators:
            log_message_callback(f"检测到没有数据的页面: {no_data_indicators}")
            return True
        
        return False
    except Exception as e:
        log_message_callback(f"检查无数据页面时发生异常: {str(e)}")
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
        # 使用更精确的JavaScript提取每一行的完整数据，确保包含第一行
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
                    
                    // 跳过首字母大写的单词（可能是导航项）- 注意：这条规则可能会误排除正常关键词
                    // 仅当不是搜索结果中的第一个关键词时才应用此规则
                    // if (/^[A-Z][a-z]+$/.test(text) && keywordCount > 0) return false;
                    
                    // 放宽这个规则，允许单个词的关键词存在
                    const symbolCount = (text.match(/[!@#$%^&*()_+=\\[\\]{};':"\\|,.<>\\/?-]/g) || []).length;
                    if (symbolCount > 2) return false;
                    
                    // 放宽这个规则，允许单个词的关键词存在
                    // if (text.trim().split(/\\s+/).length < 2 && text.length < 10) return false;
                    
                    return true;
                };
                
                // 使用新方法尝试提取表头和数据
                try {
                    // 首先检查是否存在关键词总计信息，这可以帮助我们识别有效数据区域
                    const headerInfo = document.querySelector('.sm-keywords-header-layout__header, .sm-keywords-table-header');
                    if (headerInfo) {
                        console.log("找到关键词头部信息:", headerInfo.innerText);
                    }
                    
                    // 尝试直接获取所有关键词行元素
                    // 注意：我们使用多种选择器组合来确保能找到表格行
                    const allRows = Array.from(document.querySelectorAll(
                        '.sm-table-layout__row, [role="row"], tr, .sm-table-layout tbody tr, .sm-table tr, [data-type="keyword-row"]'
                    ));
                    
                    console.log(`找到 ${allRows.length} 个可能的行元素`);
                    
                    // 尝试识别第一个关键词行 - 它通常有特殊的样式或属性
                    let firstKeywordRow = null;
                    
                    // 获取排除表头后的所有行
                    const dataRows = allRows.filter(row => {
                        // 排除明确的表头行
                        const isHeader = 
                            row.getAttribute('role') === 'rowheader' || 
                            row.querySelector('th') !== null || 
                            row.classList.contains('sm-table-layout__header-row') ||
                            row.getAttribute('aria-rowindex') === '1';  // 第一行经常是表头
                        
                        return !isHeader;
                    });
                    
                    console.log(`找到 ${dataRows.length} 个数据行`);
                    
                    // 尝试解析表头以确定每列的作用
                    let volumeColumnIndex = -1;
                    let kdColumnIndex = -1;
                    
                    // 查找表头行来识别列
                    const headerRows = allRows.filter(row => 
                        row.getAttribute('role') === 'rowheader' || 
                        row.querySelector('th') !== null || 
                        row.classList.contains('sm-table-layout__header-row') ||
                        row.getAttribute('aria-rowindex') === '1'
                    );
                    
                    if (headerRows.length > 0) {
                        const headerCells = Array.from(headerRows[0].querySelectorAll('th, td, [role="columnheader"]'));
                        console.log(`找到 ${headerCells.length} 个表头单元格`);
                        
                        headerCells.forEach((cell, index) => {
                            const cellText = cell.textContent.trim().toLowerCase();
                            console.log(`表头单元格 ${index}: ${cellText}`);
                            
                            // 查找搜索量列
                            if (cellText.includes('volume') || cellText.includes('vol') || 
                                cellText.includes('搜索量') || cellText.includes('流量')) {
                                volumeColumnIndex = index;
                                console.log(`搜索量列索引: ${volumeColumnIndex}`);
                            }
                            
                            // 查找KD列
                            if (cellText.includes('kd') || cellText.includes('difficulty') || 
                                cellText.includes('难度') || cellText.includes('竞争') || 
                                cellText.includes('kdi')) {
                                kdColumnIndex = index;
                                console.log(`KD列索引: ${kdColumnIndex}`);
                            }
                        });
                    }
                    
                    // 处理每一行数据
                    for (let i = 0; i < dataRows.length && keywordCount < maxKeywords; i++) {
                        const row = dataRows[i];
                        
                        // 获取关键词元素 - 尝试多种选择器
                        const keywordElement = 
                            row.querySelector('.sm-table-layout__cell:first-child a') || 
                            row.querySelector('a span') || 
                            row.querySelector('a') || 
                            row.querySelector('[data-type="keyword"]') ||
                            row.querySelector('[role="cell"]:first-child') ||
                            row.querySelector('td:first-child');
                        
                        if (!keywordElement) {
                            console.log("找不到关键词元素，跳过行:", row.innerText.substring(0, 50));
                            continue;
                        }
                        
                        const keyword = keywordElement.textContent.trim();
                        console.log(`发现潜在关键词: "${keyword}"`);
                        
                        // 特殊处理第一行 - 如果是搜索词本身，确保不被过滤
                        const isFirstRow = i === 0;
                        
                        // 检查关键词有效性，但对第一行做特殊处理
                        if (!isFirstRow && !isValidKeyword(keyword)) {
                            console.log(`关键词 "${keyword}" 被过滤规则排除`);
                            continue;
                        }
                        
                        // 获取单元格
                        const cells = Array.from(row.querySelectorAll('[role="cell"], td, .sm-table-layout__cell'));
                        
                        if (cells.length === 0) {
                            console.log("找不到单元格，尝试获取行中的所有文本节点");
                            continue;
                        }
                        
                        // 提取搜索量和KD
                        let volume = "0";
                        let kd = "n/a";
                        
                        // 调试输出所有单元格内容
                        if (isFirstRow) {
                            console.log("第一行单元格内容:");
                            cells.forEach((cell, idx) => {
                                console.log(`单元格 ${idx}: ${cell.textContent.trim()}`);
                            });
                        }
                        
                        // 改进的搜索量和KD提取逻辑
                        // 首先使用通过表头识别的列索引（如果可用）
                        if (volumeColumnIndex >= 0 && volumeColumnIndex < cells.length) {
                            const volumeText = cells[volumeColumnIndex].textContent.trim();
                            if (/^[0-9,.]+[KMB]?$/.test(volumeText) || /^[0-9,.]+$/.test(volumeText)) {
                                volume = volumeText;
                                console.log(`通过列索引找到搜索量: ${volume}`);
                            }
                        }
                        
                        if (kdColumnIndex >= 0 && kdColumnIndex < cells.length) {
                            const kdText = cells[kdColumnIndex].textContent.trim();
                            if (kdText.endsWith('%') || /^[0-9]+$/.test(kdText)) {
                                kd = kdText.endsWith('%') ? kdText : kdText + '%';
                                console.log(`通过列索引找到KD: ${kd}`);
                            }
                        }
                        
                        // 如果通过列索引没有找到，使用表格结构的基本规律
                        if (volume === "0" && cells.length >= 2) {
                            // 搜索量通常是第2列，它是一个数值，可能带有K、M、B等单位
                            const volumeText = cells[1].textContent.trim();
                            if (/^[0-9,.]+[KMB]?$/.test(volumeText) || /^[0-9,.]+$/.test(volumeText)) {
                                volume = volumeText;
                                console.log(`找到搜索量: ${volume}`);
                            }
                        }
                        
                        if (cells.length >= 3) {
                            // KD通常是第3列，它是一个带百分号的数值
                            const kdText = cells[2].textContent.trim();
                            if (kdText.endsWith('%') || /^[0-9]+$/.test(kdText)) {
                                kd = kdText.endsWith('%') ? kdText : kdText + '%';
                                console.log(`找到KD: ${kd}`);
                            }
                        }
                        
                        // 如果上面的方法没有找到搜索量和KD，尝试遍历所有单元格
                        if (volume === "0" || kd === "n/a") {
                            console.log("使用备选方法查找搜索量和KD");
                            // 遍历所有单元格，查找可能的搜索量和KD值
                            for (let j = 0; j < cells.length; j++) {
                                const text = cells[j].textContent.trim();
                                
                                // 识别搜索量 - 通常是带K、M、B的数字
                                if (volume === "0" && 
                                    (/^[0-9,.]+[KMB]$/.test(text) || /^[0-9,.]+$/.test(text))) {
                                    volume = text;
                                    console.log(`备选方法找到搜索量: ${volume}`);
                                }
                                
                                // 识别KD - 通常是百分比或0-100之间的数字
                                if (kd === "n/a" && 
                                    (text.endsWith('%') || 
                                     (/^[0-9]+$/.test(text) && parseInt(text) >= 0 && parseInt(text) <= 100))) {
                                    kd = text.endsWith('%') ? text : text + '%';
                                    console.log(`备选方法找到KD: ${kd}`);
                                }
                            }
                        }
                        
                        // 最后的备选方法：直接从HTML元素属性中提取数据
                        if (volume === "0" || kd === "n/a") {
                            // 尝试从data-testid或其他属性中提取
                            console.log("尝试从属性中提取数据");
                            for (let j = 0; j < cells.length; j++) {
                                // 检查是否有data-属性存储值
                                const dataVolume = cells[j].getAttribute('data-testid')?.includes('volume') ? 
                                    cells[j].textContent.trim() : null;
                                const dataKd = cells[j].getAttribute('data-testid')?.includes('kd') ? 
                                    cells[j].textContent.trim() : null;
                                
                                if (dataVolume && volume === "0") {
                                    volume = dataVolume;
                                    console.log(`从属性中找到搜索量: ${volume}`);
                                }
                                
                                if (dataKd && kd === "n/a") {
                                    kd = dataKd.endsWith('%') ? dataKd : dataKd + '%';
                                    console.log(`从属性中找到KD: ${kd}`);
                                }
                            }
                        }
                        
                        // 添加到结果
                        console.log(`添加关键词: ${keyword}, 搜索量: ${volume}, KD: ${kd}`);
                        rows.push({
                            keyword: keyword,
                            volume: volume,
                            kd: kd
                        });
                        keywordCount++;
                    }
                    
                    // 作为备用，尝试使用旧方法
                    if (rows.length === 0) {
                        console.log("新方法没有找到关键词，尝试备用方法");
                        // 这里可以使用旧的方法作为备用
                    }
                    
                } catch (e) {
                    console.error("提取关键词时出错:", e);
                }
                
                // 最终的过滤和返回
                const filteredRows = rows.filter(row => 
                    row.keyword !== 'PPC Keyword Tool' &&
                    !row.keyword.includes('PPC')
                );
                
                console.log(`最终提取了 ${filteredRows.length} 个关键词`);
                
                // 如果没有找到关键词或者数据不完整，尝试使用专门的选择器直接提取
                if (filteredRows.length === 0 || filteredRows.some(row => row.volume === "0" || row.kd === "n/a")) {
                    console.log("尝试使用直接选择器方法提取数据");
                    try {
                        // 针对截图中看到的SEMrush表格结构
                        const directRows = [];
                        const keywordRows = document.querySelectorAll('tr[data-id], .sm-table-layout__row, tr.sm-kw-row, tr.sm-table-row, tr.sm-mt-row');
                        
                        // 尝试确定每列的角色
                        let keywordColumnIndex = 0;
                        let volumeColumnIndex = 1;
                        let kdColumnIndex = 2;
                        
                        // 先查找表头确定列
                        const headers = document.querySelectorAll('th, .sm-table-layout__cell--header, .sm-table__th');
                        headers.forEach((header, index) => {
                            const headerText = header.textContent.toLowerCase();
                            if (headerText.includes('keyword') || headerText.includes('关键词')) {
                                keywordColumnIndex = index;
                            } else if (headerText.includes('volume') || headerText.includes('vol') || headerText.includes('搜索量')) {
                                volumeColumnIndex = index;
                            } else if (headerText.includes('kd') || headerText.includes('difficulty') || headerText.includes('难度')) {
                                kdColumnIndex = index;
                            }
                        });
                        
                        for (let i = 0; i < Math.min(20, keywordRows.length); i++) {
                            const row = keywordRows[i];
                            const cells = row.querySelectorAll('td, .sm-table-layout__cell');
                            
                            if (cells.length <= Math.max(keywordColumnIndex, volumeColumnIndex, kdColumnIndex)) {
                                continue;
                            }
                            
                            let keyword = cells[keywordColumnIndex].textContent.trim();
                            let volume = cells[volumeColumnIndex].textContent.trim();
                            let kd = cells[kdColumnIndex].textContent.trim();
                            
                            // 清理搜索量
                            if (!/^[0-9,.]+[KMB]?$/.test(volume)) {
                                // 尝试使用数字提取正则
                                const volumeMatch = volume.match(/([0-9,.]+[KMB]?)/);
                                if (volumeMatch) {
                                    volume = volumeMatch[1];
                                }
                            }
                            
                            // 清理KD
                            if (!kd.endsWith('%')) {
                                const kdMatch = kd.match(/([0-9,.]+)%?/);
                                if (kdMatch) {
                                    kd = kdMatch[1] + '%';
                                }
                            }
                            
                            // 如果有有效的关键词，添加到结果
                            if (keyword) {
                                directRows.push({
                                    keyword: keyword,
                                    volume: volume || "0",
                                    kd: kd || "n/a"
                                });
                            }
                        }
                        
                        console.log(`通过直接选择器找到了 ${directRows.length} 个关键词`);
                        
                        // 如果找到了关键词，并且比之前的结果更好，就使用它
                        if (directRows.length > 0 && (
                            filteredRows.length === 0 || 
                            directRows.length > filteredRows.length ||
                            directRows.some(r => r.volume !== "0" && filteredRows.every(fr => fr.volume === "0"))
                        )) {
                            return directRows.slice(0, 20);
                        }
                    } catch (e) {
                        console.error("使用直接选择器时出错:", e);
                    }
                }
                
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
    table_content = "\n\n| main words | main word vloume | Key words | Volume | Keyword Difficulty |\n"
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