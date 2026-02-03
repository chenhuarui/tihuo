import flet as ft
import openpyxl
import os
import datetime
import tempfile


def main(page: ft.Page):
    # --- 基础配置 ---
    page.title = "🦅 小鹰提货明细生成器"
    page.padding = 20
    page.theme_mode = ft.ThemeMode.LIGHT
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.scroll = ft.ScrollMode.AUTO

    # --- 逻辑处理 ---

    def get_asset_path(filename):
        """手机端 assets 路径路径方案"""
        return os.path.join("assets", filename)

    def search_customer(keyword):
        """核心搜索逻辑：完全搬运桌面版，不做任何改动"""
        data_path = get_asset_path("data.xlsx")
        if not os.path.exists(data_path):
            show_toast("错误：assets 文件夹内缺少 data.xlsx")
            return None
        try:
            wb = openpyxl.load_workbook(data_path, data_only=True)
            ws = wb["Sheet2"]
            matches = {}
            for row in range(1, ws.max_row + 1):
                cell_value = ws.cell(row=row, column=2).value
                if cell_value and keyword in str(cell_value):
                    # 抓取逻辑：当前行及后三行 (B列)
                    info = [
                        cell_value,  # final_info[0]
                        ws.cell(row=row + 1, column=2).value,  # final_info[1]
                        ws.cell(row=row + 2, column=2).value,  # final_info[2]
                        ws.cell(row=row + 3, column=2).value  # final_info[3]
                    ]
                    matches[str(cell_value)] = info
            wb.close()
            return matches
        except Exception as e:
            show_toast(f"读取数据源出错: {e}")
            return None

    def generate_and_share(final_info):
        """
        核心生成与分享逻辑
        基于你提供的官网 ShareFile 指南进行严谨实现
        """
        try:
            tpl_path = get_asset_path("template.xlsx")
            if not os.path.exists(tpl_path):
                show_toast("错误：assets 文件夹内缺少 template.xlsx")
                return

            wb = openpyxl.load_workbook(tpl_path)
            ws = wb["1"]

            # 1. 填写日期 (C2)
            today = datetime.datetime.now()
            ws["C2"] = today.strftime("%Y年%m月%d日")

            # 2. 填写客户信息 (严格对应桌面版单元格)
            ws["B6"] = final_info[0]
            ws["E6"] = final_info[1]
            ws["C6"] = final_info[2]
            ws["D6"] = final_info[3]

            # 3. 填写 UI 表单输入
            ws["G6"] = product_input.value
            ws["J6"] = count_input.value
            ws["M6"] = temp_radio.value

            # 4. 保存到临时目录
            date_str = today.strftime("%m%d")
            # 这里的 keyword 采用当前搜索框的值，模仿原版 base_filename 逻辑
            keyword = search_entry.value
            save_name = f"小鹰提明细{keyword}{date_str}.xlsx"
            save_path = os.path.join(tempfile.gettempdir(), save_name)
            wb.save(save_path)
            wb.close()

            # --- 遵照官网指南的严谨分享段落 ---

            # 检查 page.share 是否存在 (Flet 0.22+ 规范)
            if hasattr(page, "share") and page.share is not None:
                # 使用你查出的 ft.ShareFile.from_path 方法
                # 这会将本地文件包装成 Flet 能够理解的分享对象
                share_file = ft.ShareFile.from_path(save_path)

                # 调用分享接口，传入列表 []
                page.share.files([share_file])
            else:
                # 最后的保底措施，如果由于某种原因 page.share 依然没找到
                show_toast("当前环境不支持 page.share 功能")

        except Exception as e:
            show_toast(f"处理失败: {str(e)}")

    def show_toast(text):
        sb = ft.SnackBar(ft.Text(text))
        page.overlay.append(sb)
        sb.open = True
        page.update()

    # --- UI 界面渲染 (1:1 还原桌面版的功能字段) ---

    search_entry = ft.TextField(
        label="🔍 客户关键字 ",
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
            show_toast("请输入搜索关键字")
            return

        matches = search_customer(keyword)
        if not matches:
            show_toast("未找到匹配客户")
            return

        if len(matches) == 1:
            # 唯一匹配，直接生成
            generate_and_share(list(matches.values())[0])
        else:
            # 多个匹配，弹出列表供用户点击
            options = []
            for name, info in matches.items():
                def create_click_handler(v):
                    return lambda _: [setattr(bottom_sheet, "open", False), page.update(), generate_and_share(v)]

                options.append(ft.ListTile(
                    leading=ft.Icon(ft.Icons.PERSON),
                    title=ft.Text(name),
                    on_click=create_click_handler(info)
                ))
            bottom_sheet.content = ft.Column(options, tight=True, scroll=ft.ScrollMode.AUTO)
            bottom_sheet.open = True
            page.update()

    # 构建主界面绘制
    page.add(
        ft.Column([
            ft.Text("🦅 小鹰提货明细生成器", size=26, weight="bold", color=ft.Colors.BLUE_800),
            ft.Text("版本：手机适配版", size=12, color=ft.Colors.GREY_500),
            ft.Divider(height=20, color="transparent"),
            search_entry,
            product_input,
            count_input,
            ft.Container(
                content=ft.Column([
                    ft.Text("温度设置", size=14, weight="bold"),
                    temp_radio,
                ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
                padding=10,
                bgcolor=ft.Colors.BLUE_GREY_50,
                border_radius=10
            ),
            ft.Divider(height=20, color="transparent"),
            ft.ElevatedButton(
                "生成表格并分享",
                on_click=handle_gen_click,
                width=320,
                height=55,
                style=ft.ButtonStyle(
                    bgcolor=ft.Colors.BLUE_600,
                    color=ft.Colors.WHITE,
                    shape=ft.RoundedRectangleBorder(radius=12)
                )
            )
        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER)
    )


if __name__ == "__main__":
    # assets_dir 目录必须存在，且放入 data.xlsx 和 template.xlsx
    ft.app(target=main, assets_dir="assets")
