import flet as ft
import openpyxl
import os
import datetime
import tempfile


def main(page: ft.Page):
    # 基本配置
    page.title = "提货明细生成器"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.AUTO
    page.padding = 20

    # --- 1. 逻辑函数 ---

    def get_asset_path(filename):
        # 兼容性路径：直接访问 assets 目录
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

            # 填充数据
            ws["C2"] = datetime.datetime.now().strftime("%Y年%m月%d日")
            ws["B6"] = customer_info["name"]
            ws["E6"] = customer_info["phone"]
            ws["C6"] = customer_info["addr"]
            ws["D6"] = customer_info["extra"]
            ws["G6"] = product_input.value
            ws["J6"] = count_input.value
            ws["M6"] = temp_radio.value

            # 生成临时文件
            temp_dir = tempfile.gettempdir()
            save_name = f"提货明细_{customer_info['name']}.xlsx"
            save_path = os.path.join(temp_dir, save_name)
            wb.save(save_path)
            wb.close()

            # 调用分享
            page.share_files([save_path])
            show_toast("生成成功！请选择分享应用")

        except Exception as e:
            show_toast(f"生成失败: {e}")

    def show_toast(text):
        # 使用最稳的方法弹出提示
        sb = ft.SnackBar(ft.Text(text))
        page.overlay.append(sb)  # 确保SnackBar在覆盖层里
        sb.open = True
        page.update()

    # --- 2. UI 事件 ---

    def handle_gen_click(e):
        if not search_input.value:
            show_toast("请输入搜索关键字")
            return

        results = search_customer(search_input.value)
        if not results:
            show_toast("未找到匹配客户")
            return

        if len(results) > 1:
            # 多选逻辑
            options = []
            for item in results:
                def make_click_handler(info):
                    return lambda _: [setattr(bottom_sheet, "open", False), page.update(), generate_and_share(info)]

                options.append(ft.ListTile(
                    title=ft.Text(item["name"]),
                    on_click=make_click_handler(item)
                ))

            bottom_sheet.content = ft.Column(options, tight=True)
            bottom_sheet.open = True  # 传统写法，最稳
            page.update()
        else:
            generate_and_share(results[0])

    # --- 3. UI 组件 ---

    search_input = ft.TextField(label="🔍 搜索客户", border_radius=12)
    product_input = ft.TextField(label="📦 产品名称", border_radius=12)
    count_input = ft.TextField(label="📊 件数", keyboard_type=ft.KeyboardType.NUMBER, border_radius=12)

    temp_radio = ft.RadioGroup(
        content=ft.Row([
            ft.Radio(value="常温", label="常温"),
            ft.Radio(value="冷链", label="冷链")
        ], alignment=ft.MainAxisAlignment.CENTER),
        value="常温"
    )

    # 底部选择面板，必须先加入 overlay
    bottom_sheet = ft.BottomSheet(ft.Container(padding=10))
    page.overlay.append(bottom_sheet)

    # 主界面布局
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
            content=ft.Row(
                [ft.Icon(ft.Icons.SEND), ft.Text("生成并分享文件", size=16)],
                alignment=ft.MainAxisAlignment.CENTER
            ),
            width=300,
            height=50,
            on_click=handle_gen_click,
            bgcolor=ft.Colors.BLUE_600,
            color=ft.Colors.WHITE
        )
    )


# 使用 ft.app 虽有警告但在移动端构建中更成熟
if __name__ == "__main__":
    ft.app(target=main, assets_dir="assets")
