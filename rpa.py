from playwright.sync_api import sync_playwright
import time
import os
import json
import random
import signal
import sys
import re

def run_rpa(keep_open=True):
    # 已登录的Chrome用户配置文件路径
    user_data_dir = r"C:\Users\U191115\AppData\Local\Google\Chrome\User Data"
    
    # 确保目标目录存在
    screenshot_dir = "screenshots"
    if not os.path.exists(screenshot_dir):
        os.makedirs(screenshot_dir)
    
    # URL和页面名称提取
    url = "https://search.google.com/u/0/search-console/performance/search-analytics?resource_id=https%3A%2F%2Fwww.drmare.com%2F&metrics=CLICKS%2CIMPRESSIONS%2CPOSITION&breakdown=query&pli=1&page=*https%3A%2F%2Fwww.drmare.com%2Fspotify-music%2Fget-spotify-unblocked.html&num_of_months=3"
    
    # 提取页面名称
    page_url_match = re.search(r'page=\*https%3A%2F%2Fwww\.drmare\.com%2F.*%2F([^%]+)\.html', url)
    if page_url_match:
        page_name = page_url_match.group(1)
    else:
        # 备选匹配方法
        page_url_match = re.search(r'www\.drmare\.com/([^/]+)/([^/]+)\.html', url)
        if page_url_match:
            page_name = page_url_match.group(2)
        else:
            page_name = "unknown-page"
    
    # URL中的%2F会被解码为/，所以我们需要检查page_name是否包含多个部分
    if '/' in page_name:
        page_name = page_name.split('/')[-1]
    
    print(f"提取的页面名称: {page_name}")
    
    # 自定义截图命名
    first_screenshot_path = os.path.join(screenshot_dir, f"gsc-{page_name}-chart1.png")
    second_screenshot_path = os.path.join(screenshot_dir, f"gsc-{page_name}-chart2.png")
    
    # 全局变量用于保持浏览器引用
    global browser_instance
    browser_instance = None
    
    with sync_playwright() as p:
        try:
            print("启动浏览器...")
            
            # 常见的现代用户代理字符串
            user_agents = [
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36 Edg/121.0.0.0"
            ]
            
            # 随机选择一个用户代理
            user_agent = random.choice(user_agents)
            
            # 启动带有持久化配置文件的浏览器
            browser = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,  # 保持为可见状态，方便调试
                viewport={'width': 1920, 'height': 1080},  # 设置更大的视窗尺寸
                args=[
                    '--profile-directory=Default',  # 使用默认配置文件或指定其他配置文件
                    '--disable-blink-features=AutomationControlled',  # 关键：禁用自动化控制标记
                    '--no-sandbox',
                    '--disable-extensions',
                    '--disable-default-apps',
                    '--disable-popup-blocking',
                    '--start-maximized',  # 添加此参数使窗口最大化
                    f'--user-agent={user_agent}'
                ]
            )
            
            # 保存浏览器引用到全局变量
            browser_instance = browser
            
            # 创建新页面
            page = browser.new_page()
            
            # 执行窗口最大化
            page.evaluate("""
                window.moveTo(0, 0);
                window.resizeTo(screen.width, screen.height);
            """)
            
            # 添加CDP会话进一步修改浏览器指纹
            client = page.context.new_cdp_session(page)
            
            # 通过CDP会话执行更多反检测
            print("应用反检测措施...")
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
            
            # 随机延迟，模拟真实用户行为
            time.sleep(random.uniform(1, 3))
            
            # 访问目标URL
            print("导航到Google Search Console...")
            page.goto(url)
            
            # 等待页面加载完成
            print("等待页面加载完成...")
            page.wait_for_load_state("networkidle")
            
            # 检查是否需要登录
            if page.url.startswith("https://accounts.google.com/"):
                print("需要登录，尝试检查登录状态...")
                # 如果重定向到登录页面，可以等待用户手动登录
                print("请手动完成登录，然后脚本将继续...")
                
                # 等待导航到原始URL或包含search-console的URL
                page.wait_for_url(lambda url: "search-console" in url or "search.google.com" in url, timeout=120000)
                print("检测到登录成功，继续执行...")
            
            # 添加随机滚动，更像人类行为
            for _ in range(random.randint(2, 4)):
                page.mouse.wheel(0, random.randint(100, 300))
                time.sleep(random.uniform(0.5, 1.5))
            
            # 等待内容加载
            print("等待内容完全加载...")
            time.sleep(random.uniform(5, 8))
            
            # 尝试定位目标元素
            selector = "#data-container > div > div.VfPpkd-WsjYwc.VfPpkd-WsjYwc-OWXEXe-INsAgc.KC1dQ.Usd1Ac.AaN0Dd.YJ1SEc.pTyMIf > c-wiz"
            
            try:
                print("定位第一个目标元素...")
                # 使用更宽松的超时时间
                element = page.wait_for_selector(selector, timeout=60000)
                
                if element:
                    print(f"找到元素，正在截图...")
                    element.screenshot(path=first_screenshot_path)
                    print(f"第一个截图已保存为: {first_screenshot_path}")
                else:
                    raise Exception("未找到目标元素")
                    
            except Exception as e:
                print(f"定位元素时出错: {str(e)}")
                print("尝试查找其他可能的选择器...")
                
                # 尝试其他可能的选择器
                alternative_selectors = [
                    "c-wiz[data-node-index]",
                    ".KC1dQ",
                    "#data-container div[jscontroller]",
                    "#analytics-table"
                ]
                
                element_found = False
                for alt_selector in alternative_selectors:
                    try:
                        print(f"尝试选择器: {alt_selector}")
                        element = page.wait_for_selector(alt_selector, timeout=10000)
                        if element:
                            print(f"使用备选选择器 {alt_selector} 找到元素")
                            element.screenshot(path=first_screenshot_path)
                            print(f"第一个截图已保存为: {first_screenshot_path}")
                            element_found = True
                            break
                    except:
                        continue
                
                if not element_found:
                    # 如果所有选择器都失败，则截取整个页面
                    print("未能找到任何目标元素，截取整个页面...")
                    full_page_path = os.path.join(screenshot_dir, f"gsc-{page_name}-chart1-full.png")
                    page.screenshot(path=full_page_path, full_page=True)
                    print(f"整页截图已保存为: {full_page_path}")
            
            # 新增操作：点击指定元素并截取第二个元素
            try:
                print("进行额外操作：点击指定元素...")
                # 等待并点击指定元素
                click_selector = "#\\31  > div > c-wiz > div > div > div:nth-child(2) > div:nth-child(2) > div > table > thead > tr > th:nth-child(3) > span > button > span > svg"
                
                # 确保元素可见
                page.wait_for_selector(click_selector, state="visible", timeout=30000)
                
                # 随机短暂延迟
                time.sleep(random.uniform(0.8, 1.5))
                
                # 点击元素
                print(f"点击元素: {click_selector}")
                page.click(click_selector)
                
                # 等待点击后的页面变化
                time.sleep(random.uniform(1, 2))
                
                # 截取第二个指定元素
                second_selector = "#data-container > div > div:nth-child(2) > div"
                
                print(f"定位第二个目标元素: {second_selector}")
                second_element = page.wait_for_selector(second_selector, timeout=30000)
                
                if second_element:
                    print("找到第二个元素，正在截图...")
                    second_element.screenshot(path=second_screenshot_path)
                    print(f"第二个截图已保存为: {second_screenshot_path}")
                else:
                    print("未找到第二个元素")
                    
            except Exception as e:
                print(f"执行额外操作时出错: {str(e)}")
                print("尝试全页截图作为备选...")
                second_full_page_path = os.path.join(screenshot_dir, f"gsc-{page_name}-chart2-full.png")
                page.screenshot(path=second_full_page_path, full_page=True)
                print(f"第二个全页截图已保存为: {second_full_page_path}")
            
            # 在截取第二张图之后提取GSC前10个结果文本
            try:
                print("提取GSC前10个结果...")
                
                # 使用一个更通用的选择器来获取表体
                tbody_selector = "#\\31  > div > c-wiz > div > div > div:nth-child(2) > div:nth-child(2) > div > table > tbody"
                
                # 等待表体加载
                page.wait_for_selector(tbody_selector, timeout=30000)
                
                # 获取前10个查询文本
                gsc_queries = []
                
                # 方法1：直接获取所有行，然后提取前10个
                rows = page.query_selector_all(f"{tbody_selector} > tr")
                print(f"找到 {len(rows)} 行数据")
                
                # 确保我们最多只处理10行
                rows = rows[:10] if len(rows) > 10 else rows
                
                for i, row in enumerate(rows):
                    try:
                        # 获取每行的第一个单元格（查询名称）
                        query_cell = row.query_selector("td.XgRaPc[data-label='QUERIES'] span span")
                        if query_cell:
                            query_text = query_cell.inner_text().strip()
                            gsc_queries.append(query_text)
                            print(f"提取到查询 {i+1}: {query_text}")
                        else:
                            # 备选方法：尝试其他选择器模式
                            query_cell = row.query_selector("td:first-child")
                            if query_cell:
                                query_text = query_cell.inner_text().strip()
                                gsc_queries.append(query_text)
                                print(f"使用备选选择器提取到查询 {i+1}: {query_text}")
                            else:
                                print(f"无法提取第 {i+1} 行的查询文本")
                    except Exception as row_error:
                        print(f"处理第 {i+1} 行时出错: {str(row_error)}")
                
                # 如果上述方法失败，尝试使用JavaScript评估来获取文本
                if not gsc_queries:
                    print("尝试使用JavaScript评估提取查询...")
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
                    print(f"使用JavaScript评估提取到 {len(gsc_queries)} 个查询")
                
                # 更新markdown文件
                md_file_path = "add spotify to video.md"
                
                # 读取现有内容
                try:
                    with open(md_file_path, "r", encoding="utf-8") as file:
                        md_content = file.read()
                except FileNotFoundError:
                    # 如果文件不存在，创建基础内容
                    md_content = """
# add spotify to video

## 关键词来源

### Google 搜索下拉框

### 相关搜索

### GSC热门查询

### 相关问题
"""
                
                # 查找"### GSC热门查询"部分
                gsc_section_index = md_content.find("### GSC热门查询")
                next_section_index = md_content.find("###", gsc_section_index + 1)
                
                if gsc_section_index != -1:
                    # 构建新内容，加上"- "前缀和序号
                    new_content = ""
                    for i, query in enumerate(gsc_queries):
                        new_content += f"- {i+1}. {query}\n"
                    
                    # 插入新内容
                    if next_section_index != -1:
                        updated_content = md_content[:gsc_section_index + len("### GSC热门查询")] + "\n\n" + new_content + "\n" + md_content[next_section_index:]
                    else:
                        updated_content = md_content[:gsc_section_index + len("### GSC热门查询")] + "\n\n" + new_content
                    
                    # 保存更新后的内容
                    with open(md_file_path, "w", encoding="utf-8") as file:
                        file.write(updated_content)
                    
                    print(f"成功将前 {len(gsc_queries)} 个GSC查询保存到 {md_file_path}")
                else:
                    print(f"在文件 {md_file_path} 中未找到'### GSC热门查询'部分")
                
            except Exception as extract_error:
                print(f"提取查询时出错: {str(extract_error)}")
            
            # 保存cookies以供将来使用（如果需要）
            cookies = browser.cookies()
            with open('google_cookies.json', 'w') as f:
                json.dump(cookies, f)
            
            # 随机等待
            time.sleep(random.uniform(1, 3))
            
            # 任务完成提示
            print("任务完成!")
            print("浏览器将保持打开状态，您可以继续使用浏览器。")
            print("按Ctrl+C终止脚本（浏览器将保持打开状态）。")
            
            if keep_open:
                # 注册信号处理器以便干净地退出
                def signal_handler(sig, frame):
                    print("\n脚本已终止，但浏览器仍然保持打开状态。")
                    sys.exit(0)
                
                signal.signal(signal.SIGINT, signal_handler)
                
                # 使用无限循环保持脚本运行，但允许浏览器窗口被用户操作
                while True:
                    time.sleep(1)  # 减少CPU使用率
            else:
                # 如果不需要保持打开，则关闭浏览器
                browser.close()
            
            return first_screenshot_path, second_screenshot_path
            
        except Exception as e:
            print(f"发生错误: {str(e)}")
            try:
                # 尝试截取当前页面作为错误记录
                error_screenshot = os.path.join(screenshot_dir, f"error_{page_name}_{time.strftime('%Y%m%d_%H%M%S')}.png")
                page.screenshot(path=error_screenshot)
                print(f"错误截图已保存为: {error_screenshot}")
            except:
                print("无法保存错误截图")
            
            # 如果出错但仍然需要保持浏览器打开
            if keep_open and browser_instance:
                print("出现错误，但浏览器将保持打开状态。")
                print("按Ctrl+C终止脚本。")
                
                try:
                    # 注册信号处理器以便干净地退出
                    def signal_handler(sig, frame):
                        print("\n脚本已终止，但浏览器仍然保持打开状态。")
                        sys.exit(0)
                    
                    signal.signal(signal.SIGINT, signal_handler)
                    
                    # 使用无限循环保持脚本运行
                    while True:
                        time.sleep(1)
                except KeyboardInterrupt:
                    print("\n脚本已终止，但浏览器仍然保持打开状态。")
                    sys.exit(0)
            else:
                # 确保浏览器关闭（如果不需要保持打开）
                try:
                    if browser_instance:
                        browser_instance.close()
                except:
                    pass
                
            raise e

if __name__ == "__main__":
    try:
        # 传入keep_open=True保持浏览器窗口打开
        screenshots = run_rpa(keep_open=True)
        print(f"RPA任务成功完成，截图位置: {screenshots}")
    except Exception as e:
        print(f"RPA任务失败: {str(e)}")