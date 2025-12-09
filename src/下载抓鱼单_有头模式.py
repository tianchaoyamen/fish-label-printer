import subprocess
import sys
import time
import os

def get_chromium_path():
    """获取 Chromium 浏览器路径"""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        return p.chromium.executable_path

def ensure_playwright_installed():
    """确保 Playwright 已安装"""
    try:
        from playwright.sync_api import sync_playwright
        get_chromium_path()
        print("✓ Playwright 已安装")
        return True
    except ImportError:
        print("Playwright 库未安装，正在安装...")
        install_playwright_library()
        install_chromium_browser()
        return True
    except Exception as e:
        print(f"Playwright 库已安装，但浏览器未安装: {e}")
        install_chromium_browser()
        return True

def install_playwright_library():
    """安装 Playwright 库"""
    print("正在安装 Playwright 库...")
    subprocess.check_call([
        sys.executable, 
        "-m", 
        "pip", 
        "install", 
        "playwright"
    ])
    print("✓ Playwright 库安装完成")

def install_chromium_browser():
    """安装 Chromium 浏览器"""
    print("正在安装 Chromium 浏览器...")
    print("这可能需要几分钟，请耐心等待...\n")
    subprocess.check_call([
        sys.executable,
        "-m",
        "playwright",
        "install",
        "chromium"
    ])
    print("✓ Chromium 浏览器安装完成")

def main():
    """主函数"""
    from playwright.sync_api import sync_playwright
    
    print("\n开始运行主程序...")
    print("=" * 50)
    
    with sync_playwright() as p:
        # 使用持久化用户数据目录（首次运行扫码登录，后续会重用登录状态）
        from pathlib import Path
        profile_dir = Path(__file__).parent / "playwright_profile"  # 保存浏览器配置的目录

        context = p.chromium.launch_persistent_context(
             user_data_dir=str(profile_dir),
             headless=False,
             slow_mo=0,  # 取消每步强制慢速，避免不必要的等待
             accept_downloads=True,
             args=['--start-maximized'],
             viewport={'width': 1920, 'height': 1080},
             user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
         )
        # 从持久化上下文中获取页面（首次可能为空）
        page = context.new_page()
        
        # 钉钉应用URL
        target_url = "https://app82759.eapps.dingtalkcloud.com/dsp_base_app/index.html?sys=9befbf6d068e4096bb7283edc4bec916#/dashboard/7ad53c390ed94c34ac8354213afa6697?sys=9befbf6d068e4096bb7283edc4bec916&id=7ad53c390ed94c34ac8354213afa6697"
        
        print(f"正在访问: {target_url}")
        
        try:
            # 访问目标网站，增加超时时间
            # 不再等待 networkidle（可能较慢），页面 DOM 就绪后即可进行查找/点击
            page.goto(target_url, wait_until="domcontentloaded", timeout=30000)
            print(f"✓ 页面加载完成")
            
            # 页面一旦 DOM 就绪，立即开始查找导出按钮（不再阻塞太久）
            try:
                # DOM 就绪后尽快开始查找（短超时）
                page.wait_for_load_state("domcontentloaded", timeout=3000)
            except Exception:
                # 兜底短等待，确保能开始查找
                time.sleep(0.5)
            
            # 检查当前URL
            current_url = page.url
            print(f"当前URL: {current_url}")
            
            # 检查是否被重定向到登录页
            if "login" in current_url.lower() or current_url != target_url:
                print("\n" + "=" * 50)
                print("⚠️  检测到需要登录")
                print("=" * 50)
                print("\n请在浏览器窗口中完成以下操作:")
                print("1. 使用钉钉扫码登录")
                print("2. 登录成功后，确保页面正确加载")
                print("3. 登录完成后，回到此终端按 Enter 继续...")
                print("\n")
                input()
                
                # 登录后重新访问目标页面
                print("正在重新访问目标页面...")
                page.goto(target_url, wait_until="networkidle", timeout=30000)
                time.sleep(2)
            
            print(f"\n当前页面标题: {page.title()}")
            print(f"当前URL: {page.url}")
            
        except Exception as e:
            print(f"⚠️  页面加载出现问题: {e}")
            print("但浏览器窗口已打开，你可以手动操作")
        
        print("\n" + "=" * 50)
        print("🔍 浏览器窗口已打开，请查看需要点击的元素")
        print("=" * 50)
        print("\n💡 查找导出按钮的方法:")
        print("1. 在浏览器中右键点击导出按钮 -> 检查")
        print("2. 查看元素的属性（id, class, text等）")
        print("3. 记录下来，稍后添加到代码中")
        print("\n常见选择器示例:")
        print("   - 按文本: page.get_by_text('导出')")
        print("   - 按角色: page.get_by_role('button', name='导出')")
        print("   - 按ID: page.locator('#export-btn')")
        print("   - 按Class: page.locator('.export-button')")
        
        # 自动跳过手动测试，先做“快速路径”短超时尝试，能马上点就不用走冗长重试
        print("\n自动模式：直接使用选择器 i.el-tooltip.b-icon-import 点击导出触发器（跳过手动输入）")
        # 只用你指定的 selector（你可以在这里改成其它 selector）
        auto_candidates = ["i.el-tooltip.b-icon-import"]
        # 快速路径：短超时尝试直接定位并点击，失败后再走稳健重试
        clicked = False
        fast_sel = "i.el-tooltip.b-icon-import"
        try:
            page.locator(fast_sel).first.wait_for(state="visible", timeout=3000)
            page.locator(fast_sel).first.click(timeout=5000)
            clicked = True
            print("✓ 快速路径：直接通过 i.el-tooltip.b-icon-import 点击成功")
        except Exception:
            clicked = try_click_selectors(page, auto_candidates)

        if clicked:
            print("✓ 自动点击成功（使用 i.el-tooltip.b-icon-import）")
        else:
            print("⚠ 自动点击失败：未找到或点击被阻挡（可检查页面或改用更具体的 selector ）")
        
        # ===== 如果导出触发成功，自动处理“立即下载”按钮并捕获下载 =====
        # 标识是否已成功下载并处理（用于自动退出）
        download_done = False

        if clicked:
            print("等待通知并点击“立即下载”（最多 60s）...")
            from pathlib import Path
            import mimetypes
            import requests  # 作为最后兜底用法（可选）
            # 优先使用 Windows 用户下载目录，回退到脚本目录下的 downloads
            downloads_dir = None
            try:
                up = os.environ.get("USERPROFILE")
                if up:
                    downloads_dir = Path(up) / "Downloads"
            except Exception:
                downloads_dir = None
            if not downloads_dir or not downloads_dir.exists():
                downloads_dir = Path.home() / "Downloads"
            if not downloads_dir.exists():
                downloads_dir = Path(__file__).parent / "downloads"
                downloads_dir.mkdir(exist_ok=True)
            download_selector = "text=立即下载"

            # 等待导出流程结束：若出现“导出文件准备中”先等待其消失，再等待“立即下载”
            try:
                try:
                    # 如果出现“准备中”，最多等待 10 秒让它消失；超过10秒则继续尝试点击下载
                    page.wait_for_selector("text=导出文件准备中", state="visible", timeout=8000)
                    print("检测到“导出文件准备中”，最多等待 10 秒...")
                    try:
                        page.wait_for_selector("text=导出文件准备中", state="hidden", timeout=10000)
                        print("“导出文件准备中”已消失，继续等待“立即下载”。")
                    except Exception:
                        # 超时 10 秒，放弃长等待，继续尝试后续下载流程
                        print("⚠ “导出文件准备中”超过 10 秒仍未完成，继续尝试下载（不再阻塞等待）。")
                except Exception:
                    # 未出现“准备中”，直接继续等待“立即下载”
                    pass
                # 尝试等待“立即下载”出现（短超时），若未出现则后续有兜底逻辑
                try:
                    page.wait_for_selector(download_selector, timeout=15000)
                except Exception:
                    print("未检测到“立即下载”元素，后续逻辑会尝试其他方法（expect_download / requests）")
            except Exception:
                print("⚠ 等待导出完成或“立即下载”出现超时，将继续尝试（后续有兜底逻辑）")

            # 1) 等待“导出文件准备完毕”通知出现
            try:
                page.wait_for_selector("text=导出文件准备完毕", timeout=6000)
            except Exception:
                # 如果找不到上面的完整文本，后面仍会尝试查找“立即下载”
                pass

            # 2) 首选使用 expect_download 捕获下载
            try:
                with page.expect_download(timeout=60000) as dl_info:
                    page.wait_for_selector(download_selector, timeout=30000)
                    page.click(download_selector, timeout=10000)
                download = dl_info.value  # Playwright 下载对象
                # 1) 优先用 Playwright 提供的建议文件名
                filename = download.suggested_filename or "downloaded_file"
                # 2) 如果没有扩展名，尝试从响应头或 content-type 推断
                try:
                    resp = download.response()
                    # 从 content-disposition 优先解析真实文件名与扩展名
                    if resp:
                        ct = resp.headers.get("content-type", "")
                        ext = mimetypes.guess_extension(ct.split(";")[0].strip() or "")
                        if ext and "." not in filename:
                            filename += ext
                except Exception:
                    pass

                # 强制确保文件以 .xlsx 结尾
                if not filename.lower().endswith(".xlsx"):
                    base, _ = os.path.splitext(filename)
                    filename = base + ".xlsx"
                target = downloads_dir / filename
                # 保存到目标路径（覆盖同名）
                download.save_as(str(target))
                # 如果 Playwright 没提供临时路径或保存后仍无扩展名，二次校验强制改为 .xlsx
                if not target.exists():
                    print("⚠ 保存失败，目标文件不存在")
                elif not target.name.lower().endswith(".xlsx"):
                    new_target = target.with_suffix(".xlsx")
                    try:
                        os.replace(str(target), str(new_target))
                        target = new_target
                    except Exception:
                        pass

                print(f"✓ 下载完成并保存: {target}")
                download_done = True
                # 立即关闭并退出（自动化流程完成）
                try:
                    print("自动关闭浏览器并退出（下载完成）。")
                    context.close()
                except Exception:
                    pass
                return
            
            except Exception as e:
                print(f"⚠ 未通过 expect_download 成功捕获下载: {e}")

                # 3) 兜底：尝试读取“立即下载”元素的 href / data-url 并用 requests 下载（带 cookie）
                try:
                    # 先找元素并获取下载链接
                    link = page.eval_on_selector(
                        download_selector, 
                        "el => el.href || el.getAttribute('data-url') || el.dataset?.url || ''"
                    )
                    if link:
                        # 蒋当前 context 的 cookie 注入 requests
                        cookies = context.cookies()
                        jar = requests.cookies.RequestsCookieJar()
                        for c in cookies:
                            jar.set(c.get("name"), c.get("value"), domain=c.get("domain"), path=c.get("path"))
                        r = requests.get(link, cookies=jar, stream=True, timeout=60)
                        r.raise_for_status()
                        # 尝试从响应头或 URL 得到文件名
                        fn = None
                        cd = r.headers.get("content-disposition", "")
                        if "filename=" in cd:
                            fn = cd.split("filename=")[-1].strip(' "\'')
                        if not fn:
                            fn = Path(link).name or "downloaded_file"
                        # 兜底将文件名强制设为 .xlsx（你的导出是 xlsx）
                        if not fn.lower().endswith(".xlsx"):
                            base, _ = os.path.splitext(fn)
                            fn = base + ".xlsx"
                        out = downloads_dir / fn
                        with open(out, "wb") as fh:
                            for chunk in r.iter_content(1024*32):
                                fh.write(chunk)
                        print(f"✓ 通过 requests 成功下载并保存: {out}")
                        download_done = True
                        # 自动打开文件夹后关闭浏览器并退出
                        try:
                            os.startfile(str(out.parent))
                        except Exception:
                            try:
                                subprocess.run(["explorer", str(out.parent)])
                            except Exception:
                                pass
                        try:
                            print("自动关闭浏览器并退出（下载完成）。")
                            context.close()
                        except Exception:
                            pass
                        return
                except Exception as e2:
                    print(f"❌ 通过页面元素下载/requests 兜底失败: {e2}")
        
        # TODO: 在这里添加你的导出逻辑
        # 若没有触发下载，保留手动关闭浏览器（避免意外关闭）
        if not download_done:
            print("\n按 Enter 关闭浏览器...")
            input()
            try:
                context.close()
                print("✓ 浏览器已关闭")
            except Exception:
                print("⚠ 关闭浏览器时出错（可能已关闭）")

def try_click_selectors(page, candidates, max_retries: int = 3, parent_levels: int = 5) -> bool:
            """更稳健的点击尝试：等待、重试、element_handle.click、尝试父节点并搜索 frames"""
            for sel in candidates:
                try:
                    # 先在主 frame 找
                    el = page.locator(sel).first
                    try:
                        el.wait_for(state="attached", timeout=3000)
                    except Exception:
                        # 未命中主 frame，则搜遍所有 frame
                        found_in_frame = False
                        for fr in page.frames:
                            try:
                                f_el = fr.locator(sel).first
                                if f_el.count() > 0:
                                    el = f_el
                                    found_in_frame = True
                                    break
                            except Exception:
                                continue
                        if not found_in_frame:
                            print(f"未找到选择器: {sel}")
                            continue

                    cnt = 0
                    try:
                        cnt = el.count()
                    except Exception:
                        cnt = 0
                    print(f"检查 {sel} -> 匹配数量: {cnt}")
                    if cnt == 0:
                        continue

                    try:
                        el.scroll_into_view_if_needed()
                    except Exception:
                        pass

                    # 重试点击
                    for attempt in range(max_retries):
                        try:
                            el.click(timeout=5000)
                            print(f"✓ 使用 {sel} 点击成功")
                            return True
                        except Exception as click_err:
                            print(f"直接 click 失败（{sel}）尝试 {attempt+1}/{max_retries}: {click_err}")
                            # 尝试 element_handle 的 evaluate 点击
                            try:
                                handle = el.element_handle()
                                if handle:
                                    page.evaluate("(e) => e.click()", handle)
                                    print(f"✓ 使用 element_handle 点击成功（{sel}）")
                                    return True
                            except Exception:
                                pass

                            # 尝试向上查找可点击的父节点
                            try:
                                ok = page.evaluate(
                                    f"""
                                    (el) => {{
                                        let node = el;
                                        for (let i = 0; i < {parent_levels}; i++) {{
                                            node = node.parentElement;
                                            if (!node) break;
                                            const tag = node.tagName ? node.tagName.toLowerCase() : '';
                                            if (['button','a'].includes(tag) || (node.getAttribute && node.getAttribute('role') === 'button') || node.onclick) {{
                                                node.click();
                                                return true;
                                            }}
                                        }}
                                        return false;
                                    }}
                                """, el)
                                if ok:
                                    print(f"✓ 使用父节点点击成功（{sel}）")
                                    return True
                            except Exception as e:
                                print(f"尝试父节点点击时出错（{sel}）: {e}")

                        # 小等待后重试
                        time.sleep(0.5)

                except Exception as e:
                    print(f"检查选择器 {sel} 时出错: {e}")
            return False

if __name__ == "__main__":
    try:
        ensure_playwright_installed()
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  程序被用户中断")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)