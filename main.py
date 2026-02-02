import flet as ft
import openpyxl
import os
import datetime
import tempfile

def main(page: ft.Page):
    # --- UI 样式完全对齐原版 ---
    page.title = "🦅 小鹰提货明细生成器"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 400
    page.padding = 20
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.AUTO

    # --- 核心辅助函数 (与原版逻辑像素级同步) ---

    def get_asset_path(filename):
        # 手机端无法手动选路径，统一规定放在 assets 文件夹中
        return os.path.join("assets", filename)

    def search_customer(keyword):
        # 完全复制原版 search_customer 的匹配逻辑
        data_path = get_asset_path("data.xlsx")
        if not os.path.exists(data_path):
            show_toast("错误：请将 data.xlsx 放入 assets 文件夹")
            return None
        try:
            wb = openpyxl.load_workbook(data_path, data_only=True)
            ws = wb["Sheet2"]
            matches = {}
            for row in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row, column=2).value
                if cell_value and keyword in str(cell_value):
                    # 关键逻辑：直接存 4 个原始格子的值，不加标签
                    info = [
                        cell_value,
                        ws.cell(row=row + 1, column=2).value,
                        ws.cell(row=row + 2, column=2).value,
                        ws.cell(row=row + 3, column=2).value
                    ]
                    matches[str(cell_value)] = info
            wb.close()
            return matches
        except Exception as e:
            show_toast(f"读取出错: {e}")
            return None

    def generate_and_share(final_info):
        """生成并直接调起微信/系统分享"""
        try:
            tpl_path = get_asset_path("template.xlsx")
            if not os.path.exists(tpl_path):
                show_toast("错误：请将 template.xlsx 放入 assets 文件夹")
                return

            wb = openpyxl.load_workbook(tpl_path)
            ws = wb["1"]

            # 1. 填写日期 (同步 C2)
            today = datetime.datetime.now()
            ws["C2"] = today.strftime("%Y年%m月%d日")

            # 2. 填写客户数据 (严格对照原版单元格位置)
            ws["B6"] = final_info[0]
            ws["E6"] = final_info[1]
            ws["C6"] = final_info[2]
            ws["D6"] = final_info[3]

            # 3. 填写 UI 输入内容
            ws["G6"] = product_input.value
            ws["J6"] = count_input.value
            ws["M6"] = temp_radio.value

            # 4. 生成临时文件 (同步原版命名方式)
            date_str = today.strftime("%m%d")
            keyword = search_entry.value
            save_name = f"小鹰提明细{keyword}{date_str}.xlsx"
            temp_path = os.path.join(tempfile.gettempdir(), save_name)
            wb.save(temp_path)
            wb.close()

            # --- 关键：拉起手机分享 (支持 Flet 最新版 API) ---
            # 方案 1: 最新的 share API
            try:
                if hasattr(page, "share") and page.share:
                    page.share.files([ft.ShareFile(temp_path)])
                    return
            except:
                pass
            
            # 方案 2: 旧版 API (保底)
            try:
                page.share_files([temp_path])
            except AttributeError:
                show_toast("当前环境不支持分享，请检查 Flet 版本")

        except Exception as e:
            show_toast(f"生成失败: {e}")

    def show_toast(text):
        sb = ft.SnackBar(ft.Text(text))
        page.overlay.append(sb)
        sb.open = True
        page.update()

    # --- UI 界面渲染 (对照桌面版布局) ---

    search_entry = ft.TextField(
        label="🔍 客户关键字",
        on_submit=lambda _: handle_gen_click(None),
        border_radius=10
    )
    product_input = ft.TextField(label="📦 产品名称", border_radius=10)
    count_input = ft.TextField(label="📊 件数", value="1", border_radius=10)
    
    temp_radio = ft.RadioGroup(
        content=ft.Row([
            ft.Radio(value="常温", label="常温"),
            ft.Radio(value="冷链", label="冷链")
        ], alignment=ft.MainAxisAlignment.CENTER),
        value="常温"
    )

    bottom_sheet = ft.BottomSheet(ft.Container(padding=15))
    page.overlay.append(bottom_sheet)

    def handle_gen_click(e):
        keyword = search_entry.value.strip()
        if not keyword:
            show_toast("请输入关键字")
            return
        
        matches = search_customer(keyword)
        if not matches:
            show_toast("未找到客户")
            return
        
        if len(matches) == 1:
            generate_and_share(list(matches.values())[0])
        else:
            # 多个匹配项：弹出列表选择 (替代原版的弹出窗口)
            options = []
            for name, info in matches.items():
                def make_select(v):
                    return lambda _: [setattr(bottom_sheet, "open", False), page.update(), generate_and_share(v)]
                options.append(ft.ListTile(
                    title=ft.Text(f"👤 {name}"),
                    on_click=make_select(info)
                ))
            bottom_sheet.content = ft.Column(options, tight=True, scroll=ft.ScrollMode.AUTO)
            bottom_sheet.open = True
            page.update()
    page.add(
        ft.Text("🦅 提货明细生成器", size=24, weight="bold", color=ft.Colors.BLUE_900),
        ft.Text("快速生成并一键分享微信", size=12, color=ft.Colors.GREY_600),
        ft.Container(height=10),
        search_entry,
        product_input,
        count_input,
        ft.Row([ft.Text("🌡️ 温度:"), temp_radio], alignment=ft.MainAxisAlignment.CENTER),
        ft.Container(height=20),
        ft.ElevatedButton(
            "🚀 生成并发送给微信好友",
            on_click=handle_gen_click,
            width=300,
            height=50,
            style=ft.ButtonStyle(
                bgcolor=ft.Colors.BLUE_600,
                color=ft.Colors.WHITE,
                shape=ft.RoundedRectangleBorder(radius=10)
            )
        )
    )

if __name__ == "__main__":
    # assets_dir 必须指定，用于存放 data.xlsx 和 template.xlsx
    ft.app(target=main, assets_dir="assets")
