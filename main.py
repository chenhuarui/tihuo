import flet as ft
import openpyxl
import os
import datetime
import tempfile
import webbrowser # 导入浏览器模块作为最终保底

def main(page: ft.Page):
    # --- 基础配置 (完全保留) ---
    page.title = "提货明细生成器"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20

    def get_asset_path(filename):
        return os.path.join("assets", filename)

    def search_customer(keyword):
        data_path = get_asset_path("data.xlsx")
        if not os.path.exists(data_path):
            show_toast(f"找不到数据源: assets/data.xlsx")
            return []
        try:
            wb = openpyxl.load_workbook(data_path, data_only=True)
            ws = wb["Sheet2"]
            matches = []
            for row in range(1, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=2).value
                if cell_val and keyword in str(cell_val):
                    matches.append({
                        "name": cell_val,
                        "phone": ws.cell(row=row + 1, column=2).value,
                        "addr": ws.cell(row=row + 2, column=2).value,
                        "extra": ws.cell(row=row + 3, column=2).value
                    })
            wb.close()
            return matches
        except Exception as e:
            show_toast(f"读取 Excel 出错: {e}")
            return []

    def generate_and_share(customer_info):
        try:
            tpl_path = get_asset_path("template.xlsx")
            if not os.path.exists(tpl_path):
                show_toast("找不到模板文件 assets/template.xlsx")
                return

            wb = openpyxl.load_workbook(tpl_path)
            ws = wb["1"]

            ws["C2"] = datetime.datetime.now().strftime("%Y年%m月%d日")
            ws["B6"], ws["E6"] = customer_info["name"], customer_info["phone"]
            ws["C6"], ws["D6"] = customer_info["addr"], customer_info["extra"]
            ws["G6"], ws["J6"], ws["M6"] = product_input.value, count_input.value, temp_radio.value

            temp_dir = tempfile.gettempdir()
            save_name = f"提货明细_{customer_info['name']}.xlsx"
            save_path = os.path.join(temp_dir, save_name)
            wb.save(save_path)
            wb.close()

            # ---大师级暴力兼容分享逻辑---
            print(f"正在尝试处理文件: {save_path}")
            
            # 方案 A: 尝试调用手机系统的 share_files (如果版本支持)
            try:
                # 不再用 hasattr 检查，直接尝试运行
                page.share_files([save_path])
                show_toast("已调起系统分享")
            except (AttributeError, Exception) as e:
                # 方案 B: 如果方案 A 报错 AttributeError，说明 page 真的没有这个属性
                print(f"系统分享不可用，改用保底方案: {e}")
                
                # 如果在安卓上，尝试用 file 协议唤起
                try:
                    # 将路径转换为 URI 格式
                    file_url = f"file://{save_path}"
                    page.launch_url(file_url)
                    show_toast("尝试通过系统打开文件...")
                except:
                    # 方案 C: 最终保底，弹窗告诉用户路径，并尝试在 PC 上打开文件夹
                    show_toast(f"分享接口失效。文件已存至: {save_path}")
                    if os.name == 'nt': # 如果是电脑端调试，直接打开目录
                        os.startfile(temp_dir)

        except Exception as e:
            show_toast(f"生成失败: {e}")

    def show_toast(text):
        sb = ft.SnackBar(ft.Text(text))
        page.overlay.append(sb)
        sb.open = True
        page.update()

    def handle_gen_click(e):
        if not search_input.value:
            show_toast("请输入搜索关键字")
            return
        results = search_customer(search_input.value)
        if not results:
            show_toast("未找到匹配客户")
            return
        if len(results) > 1:
            options = []
            for item in results:
                def make_handler(info):
                    return lambda _: [setattr(bottom_sheet, "open", False), page.update(), generate_and_share(info)]
                options.append(ft.ListTile(title=ft.Text(item["name"]), on_click=make_handler(item)))
            bottom_sheet.content = ft.Container(content=ft.Column(options, tight=True, scroll=ft.ScrollMode.AUTO), padding=10, height=400)
            bottom_sheet.open = True
            page.update()
        else:
            generate_and_share(results[0])

    # --- UI 组件保持不变 ---
    search_input = ft.TextField(label="🔍 搜索客户", border_radius=12)
    product_input = ft.TextField(label="📦 产品名称", border_radius=12)
    count_input = ft.TextField(label="📊 件数", keyboard_type=ft.KeyboardType.NUMBER, border_radius=12)
    temp_radio = ft.RadioGroup(content=ft.Row([ft.Radio(value="常温", label="常温"), ft.Radio(value="冷链", label="冷链")], alignment=ft.MainAxisAlignment.CENTER), value="常温")
    bottom_sheet = ft.BottomSheet(ft.Container(padding=10))
    page.overlay.append(bottom_sheet)

    page.add(
        ft.Container(height=10),
        ft.Text("🦅 提货明细生成器", size=26, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_700),
        ft.Divider(height=20),
        search_input,
        product_input,
        count_input,
        ft.Row([ft.Text("🌡️ 温度选择:"), temp_radio], alignment=ft.MainAxisAlignment.CENTER),
        ft.Container(height=20),
        ft.ElevatedButton(
            content=ft.Row([ft.Icon(ft.Icons.SEND), ft.Text("生成并分享文件", size=16)], alignment=ft.MainAxisAlignment.CENTER),
            width=300, height=50, on_click=handle_gen_click, bgcolor=ft.Colors.BLUE_600, color=ft.Colors.WHITE
        )
    )

if __name__ == "__main__":
    ft.app(target=main, assets_dir="assets")
