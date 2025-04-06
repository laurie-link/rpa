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
            
            # 立即检查是否出现任何错误页面
            error_type = check_semrush_error_page(log_message_callback, page)
            if error_type:
                if error_type == 'login_expired' or error_type == 'redirected_to_login':
                    log_message_callback(f"检测到SEMrush账号在其他地方登录或会话失效，立即重试...")
                    # 截取400错误页面截图以便调试
                    error_screenshot_path = os.path.join(screenshot_dir, f"semrush-400error-{page_name}-{retry_count}.png")
                    try:
                        page.screenshot(path=error_screenshot_path)
                        log_message_callback(f"已保存400错误页面截图: {error_screenshot_path}")
                    except Exception as ss_error:
                        log_message_callback(f"保存截图时出错: {str(ss_error)}")
                elif error_type == 'data_unavailable':
                    log_message_callback(f"检测到SEMrush数据不可用错误，这也表示没有相关数据...")
                    # 截取数据错误页面截图以便调试
                    error_screenshot_path = os.path.join(screenshot_dir, f"semrush-data-error-{page_name}-{retry_count}.png")
                    try:
                        page.screenshot(path=error_screenshot_path)
                        log_message_callback(f"已保存数据错误页面截图: {error_screenshot_path}")
                    except Exception as ss_error:
                        log_message_callback(f"保存截图时出错: {str(ss_error)}")
                    
                    # 与no_data_found类似，直接创建空记录并继续
                    log_message_callback("由于SEMrush报告没有此关键词的数据（数据不可用错误），创建空记录并继续...")
                    update_semrush_markdown(log_message_callback, page_name, [], [], {
                        'allKeywords': '0',
                        'totalVolume': '0',
                        'averageKD': 'N/A',
                        'note': 'SEMrush报告没有此关键词的相关数据（数据不可用错误）'
                    })
                    return True
                elif error_type == 'no_data_found':
                    log_message_callback(f"检测到SEMrush无数据错误页面，无法找到相关关键词数据...")
                    # 截取无数据错误页面截图
                    error_screenshot_path = os.path.join(screenshot_dir, f"semrush-no-data-{page_name}-{retry_count}.png")
                    try:
                        page.screenshot(path=error_screenshot_path)
                        log_message_callback(f"已保存无数据错误页面截图: {error_screenshot_path}")
                    except Exception as ss_error:
                        log_message_callback(f"保存截图时出错: {str(ss_error)}")
                    
                    # 在这种情况下，我们可以选择创建一个空的数据记录并返回，而不是重试
                    # 因为这表示该关键词确实没有数据，重试也不会有结果
                    log_message_callback("由于SEMrush报告没有此关键词的数据，创建空记录并继续...")
                    update_semrush_markdown(log_message_callback, page_name, [], [], {
                        'allKeywords': '0',
                        'totalVolume': '0',
                        'averageKD': 'N/A',
                        'note': 'SEMrush报告没有此关键词的相关数据'
                    })
                    return True
                
                # 只有非特殊错误类型才立即重试
                if error_type and error_type != 'data_unavailable' and error_type != 'no_data_found':
                    retry_count += 1
                    # 延迟短暂时间后重试
                    time.sleep(2)
                    continue
            else:
                log_message_callback("等待关键词元素出现...")
                try:
                    # 尝试等待关键词表格行或关键词组元素出现
                    page.wait_for_selector(".sm-table-layout__row, [role='row'], tr, .sm-group-content", 
                                         state="visible", timeout=60000)
                    log_message_callback("SEMrush关键词元素已出现，继续处理...")
                except Exception as wait_error:
                    log_message_callback(f"等待元素超时，将检查页面状态: {str(wait_error)}")
                    
                    # 再次检查是否是错误页面
                    error_type = check_semrush_error_page(log_message_callback, page)
                    if error_type:
                        if error_type == 'login_expired' or error_type == 'redirected_to_login':
                            log_message_callback(f"在等待元素超时后检测到SEMrush账号在其他地方登录或会话失效，立即重试...")
                            # 截取400错误页面截图以便调试
                            error_screenshot_path = os.path.join(screenshot_dir, f"semrush-400error-timeout-{page_name}-{retry_count}.png")
                            try:
                                page.screenshot(path=error_screenshot_path)
                                log_message_callback(f"已保存400错误页面截图: {error_screenshot_path}")
                            except Exception as ss_error:
                                log_message_callback(f"保存截图时出错: {str(ss_error)}")
                        elif error_type == 'data_unavailable':
                            log_message_callback(f"在等待元素超时后检测到SEMrush数据不可用错误，这也表示没有相关数据...")
                            # 截取数据错误页面截图以便调试
                            error_screenshot_path = os.path.join(screenshot_dir, f"semrush-data-error-{page_name}-{retry_count}.png")
                            try:
                                page.screenshot(path=error_screenshot_path)
                                log_message_callback(f"已保存数据错误页面截图: {error_screenshot_path}")
                            except Exception as ss_error:
                                log_message_callback(f"保存截图时出错: {str(ss_error)}")
                            # 与no_data_found类似，直接创建空记录并继续
                            log_message_callback("由于SEMrush报告没有此关键词的数据（数据不可用错误），创建空记录并继续...")
                            update_semrush_markdown(log_message_callback, page_name, [], [], {
                                'allKeywords': '0',
                                'totalVolume': '0',
                                'averageKD': 'N/A',
                                'note': 'SEMrush报告没有此关键词的相关数据（数据不可用错误）'
                            })
                            return True
                        elif error_type == 'no_data_found':
                            log_message_callback(f"在等待元素超时后检测到SEMrush无数据错误页面，无法找到相关关键词数据...")
                            # 截取无数据错误页面截图
                            error_screenshot_path = os.path.join(screenshot_dir, f"semrush-no-data-timeout-{page_name}-{retry_count}.png")
                            try:
                                page.screenshot(path=error_screenshot_path)
                                log_message_callback(f"已保存无数据错误页面截图: {error_screenshot_path}")
                            except Exception as ss_error:
                                log_message_callback(f"保存截图时出错: {str(ss_error)}")
                            
                            # 在这种情况下，我们可以选择创建一个空的数据记录并返回，而不是重试
                            # 因为这表示该关键词确实没有数据，重试也不会有结果
                            log_message_callback("由于SEMrush报告没有此关键词的数据，创建空记录并继续...")
                            update_semrush_markdown(log_message_callback, page_name, [], [], {
                                'allKeywords': '0',
                                'totalVolume': '0',
                                'averageKD': 'N/A',
                                'note': 'SEMrush报告没有此关键词的相关数据'
                            })
                            return True
                        else:
                            log_message_callback(f"在等待元素超时后检测到SEMrush错误: {error_type}，将进行重试...")
                            # 截取通用错误页面截图
                            error_screenshot_path = os.path.join(screenshot_dir, f"semrush-general-error-timeout-{page_name}-{retry_count}.png")
                            try:
                                page.screenshot(path=error_screenshot_path)
                                log_message_callback(f"已保存错误页面截图: {error_screenshot_path}")
                            except Exception as ss_error:
                                log_message_callback(f"保存截图时出错: {str(ss_error)}")
                        
                        # 只有非特殊错误类型才进行重试
                        if not error_type or (error_type != 'data_unavailable' and error_type != 'no_data_found'):
                            retry_count += 1
                            # 延迟短暂时间后重试
                            time.sleep(3)  # 超时后多等待一秒
                            continue
            
            # 提取统计信息
            log_message_callback("提取SEMrush页面统计信息...")
            stats_data = extract_semrush_stats(log_message_callback, page)
            
            # 提取边栏数据
            log_message_callback("提取SEMrush边栏数据(最多20条)...")
            sidebar_data = extract_semrush_sidebar_data(log_message_callback, page)
            
            # 提取主要关键词数据
            log_message_callback("提取SEMrush主要关键词数据(最多20条)...")
            keyword_data = extract_semrush_keyword_data(log_message_callback, page)
            
            # 检验提取的数据
            if (keyword_data and len(keyword_data) > 0) or (sidebar_data and len(sidebar_data) > 0) or stats_data:
                # 整合数据并更新markdown文件
                log_message_callback("整合SEMrush数据并更新markdown文件...")
                update_semrush_markdown(log_message_callback, page_name, sidebar_data, keyword_data, stats_data)
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
    update_semrush_markdown(log_message_callback, page_name, [], [], {})
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
    """检查是否是SEMrush错误页面，加强对400错误和其他错误页面的检测"""
    try:
        # 首先尝试使用提供的选择器快速检测400错误页面
        try:
            # 检查是否存在指定的400错误页面元素
            error_element = page.query_selector("body > div.main > div > div:nth-child(1)")
            if error_element:
                error_text = error_element.inner_text()
                if '400' in error_text:
                    log_message_callback(f"使用选择器检测到400错误页面: {error_text[:50]}...")
                    return 'login_expired'  # 返回错误类型而不是布尔值，以提供更多信息
            
            # 新增: 检测第三种错误类型 - 没有找到相关数据的错误
            no_data_found_section = page.query_selector("section.sm-global-na, [data-testid='nothing-found-card']")
            if no_data_found_section:
                try:
                    error_text = no_data_found_section.inner_text()
                    if "couldn't find any data" in error_text or "no data" in error_text.lower():
                        log_message_callback(f"检测到SEMrush无数据错误页面: {error_text[:50]}...")
                        return 'no_data_found'
                except Exception as text_error:
                    log_message_callback(f"获取无数据错误文本时出错: {str(text_error)}")
            
            # 备用检测第三种错误类型的方法
            try:
                title_element = page.query_selector(".sm-global-na__title, [data-testid='nothing-found-title']")
                if title_element:
                    title_text = title_element.inner_text()
                    if "couldn't find any data" in title_text or "no data" in title_text.lower():
                        log_message_callback(f"通过标题检测到SEMrush无数据错误: {title_text}")
                        return 'no_data_found'
            except Exception as title_error:
                log_message_callback(f"获取错误标题时出错: {str(title_error)}")
            
            # 新增: 检测图片错误元素
            error_img = page.query_selector("img.kwo-global-na__img")
            if error_img:
                log_message_callback("检测到SEMrush错误图片 (kwo-global-na__img)，可能是服务不可用或数据加载错误")
                return 'data_unavailable'
            
            # 尝试检测其他可能的错误图片或容器
            additional_error_selectors = [
                ".kwo-global-na", # 图片的父容器
                ".kwo-global-na__title", # 可能存在的错误标题
                ".kwo-global-na__text", # 可能存在的错误文本
                ".sm-kw-error", # 其他可能的错误类
                ".sm-error-container",
                "[class*='error']", # 任何包含error的类名
                "[class*='na__img']" # 任何包含na__img的类名
            ]
            
            for selector in additional_error_selectors:
                error_el = page.query_selector(selector)
                if error_el:
                    try:
                        error_text = error_el.inner_text() or "无文本内容"
                    except:
                        error_text = "无法获取文本"
                    
                    log_message_callback(f"检测到SEMrush错误元素 ({selector}): {error_text[:50]}...")
                    return 'data_unavailable'
            
            # 备用选择器检测400错误
            backup_selectors = [
                "div.error-container", 
                ".error-code", 
                ".error-message",
                "div.main > div > div",  # 更通用的选择器
                "h1.error-title"
            ]
            
            for selector in backup_selectors:
                el = page.query_selector(selector)
                if el:
                    text = el.inner_text()
                    if '400' in text or '错误' in text or 'Error' in text or '登录已失效' in text:
                        log_message_callback(f"使用备用选择器 '{selector}' 检测到错误页面: {text[:50]}...")
                        return 'login_expired'
        except Exception as selector_error:
            log_message_callback(f"使用选择器检测错误时出现异常: {str(selector_error)}")
        
        # 使用更全面的JavaScript评估来检查页面内容
        error_content = page.evaluate("""
            () => {
                // 检查页面URL，某些错误会反映在URL中
                if (window.location.href.includes('error') || 
                    window.location.href.includes('400') || 
                    window.location.href.includes('401') || 
                    window.location.href.includes('403')) {
                    return 'error_in_url';
                }
                
                // 检查页面标题
                if (document.title.includes('Error') || 
                    document.title.includes('错误') || 
                    document.title.includes('400')) {
                    return 'error_in_title';
                }
                
                // 检查整个页面文本
                const fullText = document.body.innerText;
                
                // 检查400错误页面 - 最优先检查
                if (fullText.includes('400') && 
                    (fullText.includes('登录已失效') || 
                     fullText.includes('失效') ||
                     fullText.includes('已在其他地方登录') ||
                     fullText.includes('请重新登录'))) {
                    return 'login_expired';
                }
                
                // 检查"Something went wrong"错误
                if (fullText.includes('Something went wrong') || 
                    fullText.includes('went wrong') ||
                    fullText.includes('出错了')) {
                    return 'something_went_wrong';
                }
                
                // 检查401/403错误
                if (fullText.includes('401') || 
                    fullText.includes('403') || 
                    fullText.includes('Unauthorized') ||
                    fullText.includes('Forbidden') ||
                    fullText.includes('未授权')) {
                    return 'unauthorized';
                }
                
                // 检查登录页面（可能是被重定向了）
                if (fullText.includes('登录') && 
                    fullText.includes('密码') && 
                    (fullText.includes('Sign in') || fullText.includes('Log in'))) {
                    return 'redirected_to_login';
                }
                
                // 检查其他一般错误消息
                if (fullText.includes('error') || 
                    fullText.includes('Error') ||
                    fullText.includes('失败') ||
                    fullText.includes('错误')) {
                    return 'general_error';
                }
                
                // 检查是否有特定的错误元素
                const errorElements = document.querySelectorAll('.error, .error-message, .error-container, [class*=error]');
                if (errorElements.length > 0) {
                    for (const el of errorElements) {
                        if (el.offsetWidth > 0 && el.offsetHeight > 0) { // 确保元素可见
                            return 'error_element_found';
                        }
                    }
                }
                
                return false;
            }
        """)
        
        if error_content:
            log_message_callback(f"检测到SEMrush错误页面: {error_content}")
            return error_content
        
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
    """提取SEMrush主要关键词数据和页面顶部统计信息"""
    try:
        # 首先尝试提取页面顶部的统计信息
        stats_info = extract_semrush_stats(log_message_callback, page)
        
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
                        
                        // 如果通过列索引没有找到搜索量和KD，使用表格结构的基本规律
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

def extract_semrush_stats(log_message_callback, page):
    """提取SEMrush页面顶部的统计信息"""
    try:
        # 使用JavaScript评估提取统计信息
        stats = page.evaluate("""
            () => {
                // 查找可能包含统计信息的元素
                const statsElements = [
                    // 尝试多种选择器定位统计信息
                    document.querySelector('.sm-keywords-table-header-animation'),
                    document.querySelector('.sm-keywords-table-header'),
                    document.querySelector('.sm-kw-table-header'),
                    document.querySelector('.sm-mt-table-header'),
                    document.querySelector('[class*="keywords-table-header"]'),
                    // 如图片所示的元素位置
                    document.querySelector('div[class*="table-header-animation"]')
                ].filter(el => el);
                
                // 如果找到了元素
                if (statsElements.length > 0) {
                    const statsContainer = statsElements[0];
                    const statsText = statsContainer.innerText;
                    
                    // 尝试从文本中提取统计数据
                    const allKeywordsMatch = statsText.match(/All keywords[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i) || 
                                            statsText.match(/(\\d[\\d,\\.]*[KMB]?)\\s*keywords/i);
                    const totalVolumeMatch = statsText.match(/Total Volume[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i) || 
                                            statsText.match(/Volume[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i);
                    const avgKDMatch = statsText.match(/Average KD[:\\s]*(\\d+%)/i) || 
                                      statsText.match(/Avg[\\s.]*KD[:\\s]*(\\d+%)/i) ||
                                      statsText.match(/KD[:\\s]*(\\d+%)/i);
                    
                    return {
                        allKeywords: allKeywordsMatch ? allKeywordsMatch[1] : null,
                        totalVolume: totalVolumeMatch ? totalVolumeMatch[1] : null,
                        averageKD: avgKDMatch ? avgKDMatch[1] : null,
                        rawText: statsText
                    };
                }
                
                // 备选方法：尝试查找具有特定内容的元素
                const allTexts = [];
                document.querySelectorAll('div, span, p').forEach(el => {
                    const text = el.innerText.trim();
                    if (text && (
                        text.includes('keywords') || 
                        text.includes('volume') || 
                        text.includes('KD')
                    )) {
                        allTexts.push({
                            element: el.tagName,
                            text: text
                        });
                    }
                });
                
                // 从收集的文本中提取统计信息
                let allKeywords = null, totalVolume = null, averageKD = null;
                
                allTexts.forEach(item => {
                    if (!allKeywords && 
                        (item.text.match(/All keywords[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i) || 
                         item.text.match(/(\\d[\\d,\\.]*[KMB]?)\\s*keywords/i))) {
                        const match = item.text.match(/All keywords[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i) || 
                                     item.text.match(/(\\d[\\d,\\.]*[KMB]?)\\s*keywords/i);
                        allKeywords = match ? match[1] : null;
                    }
                    
                    if (!totalVolume && 
                        (item.text.match(/Total Volume[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i) || 
                         item.text.match(/Volume[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i))) {
                        const match = item.text.match(/Total Volume[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i) || 
                                     item.text.match(/Volume[:\\s]*(\\d[\\d,\\.]*[KMB]?)/i);
                        totalVolume = match ? match[1] : null;
                    }
                    
                    if (!averageKD && 
                        (item.text.match(/Average KD[:\\s]*(\\d+%)/i) || 
                         item.text.match(/Avg[\\s.]*KD[:\\s]*(\\d+%)/i) ||
                         item.text.match(/KD[:\\s]*(\\d+%)/i))) {
                        const match = item.text.match(/Average KD[:\\s]*(\\d+%)/i) || 
                                     item.text.match(/Avg[\\s.]*KD[:\\s]*(\\d+%)/i) ||
                                     item.text.match(/KD[:\\s]*(\\d+%)/i);
                        averageKD = match ? match[1] : null;
                    }
                });
                
                if (allKeywords || totalVolume || averageKD) {
                    return {
                        allKeywords: allKeywords,
                        totalVolume: totalVolume,
                        averageKD: averageKD,
                        rawText: allTexts.map(item => item.text).join(' | ')
                    };
                }
                
                return {
                    allKeywords: null,
                    totalVolume: null,
                    averageKD: null,
                    rawText: "未找到统计信息"
                };
            }
        """)
        
        # 记录找到的统计信息
        if stats:
            log_message_callback("提取的SEMrush统计信息:")
            if stats.get('allKeywords'):
                log_message_callback(f"关键词总数: {stats.get('allKeywords')}")
            if stats.get('totalVolume'):
                log_message_callback(f"总搜索量: {stats.get('totalVolume')}")
            if stats.get('averageKD'):
                log_message_callback(f"平均关键词难度: {stats.get('averageKD')}")
            
            if not (stats.get('allKeywords') or stats.get('totalVolume') or stats.get('averageKD')):
                log_message_callback(f"未能找到统计信息，原始文本: {stats.get('rawText', '无文本')}")
        else:
            log_message_callback("未能提取SEMrush统计信息")
        
        return stats
        
    except Exception as e:
        log_message_callback(f"提取SEMrush统计信息时出错: {str(e)}")
        return {
            'allKeywords': None,
            'totalVolume': None,
            'averageKD': None
        }

def update_semrush_markdown(log_message_callback, page_name, sidebar_data, keyword_data, stats_data=None):
    """更新markdown文件中的SEMrush数据 - 使用对齐表格格式，包含统计信息"""
    if not sidebar_data and not keyword_data and not stats_data:
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
    
    # 添加统计信息（如果存在）
    stats_content = ""
    if stats_data:
        stats_content = "\n"
        if stats_data.get('allKeywords'):
            stats_content += f"- ALL Keywords: **{stats_data.get('allKeywords')}**\n"
        if stats_data.get('totalVolume'):
            stats_content += f"- Total Volume: **{stats_data.get('totalVolume')}**\n"
        if stats_data.get('averageKD'):
            stats_content += f"- Average KD: **{stats_data.get('averageKD')}**\n"
        if stats_data.get('note'):
            stats_content += f"- 说明: *{stats_data.get('note')}*\n"
        stats_content += "\n"
    
    # 确定实际有多少行数据
    max_rows = max(len(sidebar_data), len(keyword_data))
    
    # 定义表头
    table_headers = ["main words", "main word vloume", "Key words", "Volume", "Keyword Difficulty"]
    
    # 如果没有数据，至少添加一行空行
    if max_rows == 0:
        # 无数据时使用默认表头宽度
        main_words_width = len(table_headers[0])
        main_word_volume_width = len(table_headers[1])
        key_words_width = len(table_headers[2])
        volume_width = len(table_headers[3])
        kd_width = len(table_headers[4])
        
        # 构建表头行
        header_row = f"| {table_headers[0].ljust(main_words_width)} | {table_headers[1].ljust(main_word_volume_width)} | {table_headers[2].ljust(key_words_width)} | {table_headers[3].ljust(volume_width)} | {table_headers[4].ljust(kd_width)} |"
        
        # 构建分隔行
        separator_row = f"| {':'.ljust(main_words_width, '-')} | {':'.ljust(main_word_volume_width, '-')} | {':'.ljust(key_words_width, '-')} | {':'.ljust(volume_width, '-')} | {':'.ljust(kd_width, '-')} |"
        
        # 添加表头、分隔行和空行
        table_content = f"{header_row}\n{separator_row}\n| {''.ljust(main_words_width)} | {''.ljust(main_word_volume_width)} | {''.ljust(key_words_width)} | {''.ljust(volume_width)} | {''.ljust(kd_width)} |\n"
    else:
        # 填充表格数据 - 确保过滤掉PPC内容
        filtered_sidebar = [item for item in sidebar_data if not item['text'].startswith('PPC')]
        filtered_keywords = [item for item in keyword_data if not item['keyword'].startswith('PPC')]
        
        max_rows = max(len(filtered_sidebar), len(filtered_keywords))
        
        # 判断是否有数据，根据最大数据长度调整表头宽度
        if filtered_sidebar or filtered_keywords:
            # 计算每列数据的最大宽度
            main_words_width = max(len(table_headers[0]), *[len(item['text']) for item in filtered_sidebar]) if filtered_sidebar else len(table_headers[0])
            main_word_volume_width = max(len(table_headers[1]), *[len(item['value']) for item in filtered_sidebar]) if filtered_sidebar else len(table_headers[1])
            key_words_width = max(len(table_headers[2]), *[len(item['keyword']) for item in filtered_keywords]) if filtered_keywords else len(table_headers[2])
            volume_width = max(len(table_headers[3]), *[len(item['volume']) for item in filtered_keywords]) if filtered_keywords else len(table_headers[3])
            kd_width = max(len(table_headers[4]), *[len(item['kd']) for item in filtered_keywords]) if filtered_keywords else len(table_headers[4])
        else:
            # 无数据时使用默认表头宽度
            main_words_width = len(table_headers[0])
            main_word_volume_width = len(table_headers[1])
            key_words_width = len(table_headers[2])
            volume_width = len(table_headers[3])
            kd_width = len(table_headers[4])
        
        # 构建表头行，确保每个标题宽度与数据一致
        header_row = f"| {table_headers[0].ljust(main_words_width)} | {table_headers[1].ljust(main_word_volume_width)} | {table_headers[2].ljust(key_words_width)} | {table_headers[3].ljust(volume_width)} | {table_headers[4].ljust(kd_width)} |"
        
        # 构建分隔行，确保与表头宽度一致
        separator_row = f"| {':'.ljust(main_words_width, '-')} | {':'.ljust(main_word_volume_width, '-')} | {':'.ljust(key_words_width, '-')} | {':'.ljust(volume_width, '-')} | {':'.ljust(kd_width, '-')} |"
        
        # 添加表头和分隔行
        table_content = f"{header_row}\n{separator_row}\n"
        
        # 处理表格行数据
        for i in range(max_rows):
            main_word = filtered_sidebar[i]['text'] if i < len(filtered_sidebar) else ""
            main_word_count = filtered_sidebar[i]['value'] if i < len(filtered_sidebar) else ""
            keyword = filtered_keywords[i]['keyword'] if i < len(filtered_keywords) else ""
            volume = filtered_keywords[i]['volume'] if i < len(filtered_keywords) else ""
            kd = filtered_keywords[i]['kd'] if i < len(filtered_keywords) else ""
            
            # 填充表格行
            table_content += f"| {main_word.ljust(main_words_width)} | {main_word_count.ljust(main_word_volume_width)} | {keyword.ljust(key_words_width)} | {volume.ljust(volume_width)} | {kd.ljust(kd_width)} |\n"
    
    # 查找下一部分的开始
    next_section_index = md_content.find("###", semrush_index + len(semrush_header))
    
    # 将统计信息和表格内容组合
    combined_content = stats_content + table_content
    
    # 插入统计信息和表格内容
    if next_section_index != -1:
        updated_content = md_content[:semrush_index + len(semrush_header)] + combined_content + "\n" + md_content[next_section_index:]
    else:
        updated_content = md_content[:semrush_index + len(semrush_header)] + combined_content
    
    # 保存更新后的内容
    try:
        with open(md_file_path, "w", encoding="utf-8") as file:
            file.write(updated_content)
        
        log_message_callback(f"成功将SEMrush数据保存到 {md_file_path}")
    except Exception as e:
        log_message_callback(f"保存MD文件时出错: {str(e)}")