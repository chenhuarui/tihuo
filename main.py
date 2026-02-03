import flet as ft
import openpyxl
from openpyxl.styles import Alignment, Border, Side, Font
import os
import tempfile
from datetime import datetime

def main(page: ft.Page):
    page.title = "提货明细生成器"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.window_width = 400
    page.window_height = 800
    # 手机适配设置
    page.padding = 10

    # --- 1. 初始化分享控件 (关键步骤) ---
    share_controller = ft.Share()
    page.overlay.append(share_controller)

    # 结果提示函数
    def show_toast(text):
        page.snack_bar = ft.SnackBar(ft.Text(text))
        page.snack_bar.open = True
        page.update()

    # --- 2. 生成 Excel 并分享的逻辑 ---
    def generate_and_share(e):
        # 验证输入
        if not customer_name.value:
            show_toast("请输入客户姓名")
            return

        try:
            # 创建 Excel (保持你原始的 openpyxl 逻辑)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "提货明细"

            # 设置列宽
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 20

            # 写入基础信息
            data = [
                ["客户姓名", customer_name.value],
                ["提货日期", datetime.now().strftime("%Y-%m-%d")],
                ["", ""],
                ["品名", "数量"]
            ]

            # 遍历表格中的数据
            for row_view in data_table.rows:
                product = row_view.cells[0].content.value
                amount = row_view.cells[1].content.value
                if product and amount:
                    data.append([product, amount])

            for row in data:
                ws.append(row)

            # --- 保存文件 ---
            temp_dir = tempfile.gettempdir()
            file_name = f"提货单_{customer_name.value}_{datetime.now().strftime('%H%M%S')}.xlsx"
            save_path = os.path.join(temp_dir, file_name)
            wb.save(save_path)
            wb.close()

            # --- 3. 根据你查到的文档执行分享 ---
            # 使用 ShareFile.from_path
            excel_file = ft.ShareFile.from_path(
                path=save_path,
                name=file_name, # 分享给别人时显示的名字
                mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # 调用 share_files 方法
            share_controller.share_files([excel_file])
            show_toast("正在调起系统分享...")

        except Exception as ex:
            show_toast(f"发生错误: {str(ex)}")

    # --- UI 界面部分 (保持或优化你的原始设计) ---
    customer_name = ft.TextField(label="客户姓名", hint_text="输入姓名")
    
    # 手动输入区域
    input_product = ft.TextField(label="品名", expand=2)
    input_amount = ft.TextField(label="数量", expand=1, keyboard_type=ft.KeyboardType.NUMBER)

    def add_row(e):
        if input_product.value and input_amount.value:
            data_table.rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.TextField(value=input_product.value)),
                        ft.DataCell(ft.TextField(value=input_amount.value)),
                    ]
                )
            )
            input_product.value = ""
            input_amount.value = ""
            page.update()

    data_table = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("品名")),
            ft.DataColumn(ft.Text("数量")),
        ],
        rows=[]
    )

    # 页面布局
    page.add(
        ft.Column([
            ft.Text("提货信息录入", size=20, weight=ft.FontWeight.BOLD),
            customer_name,
            ft.Row([input_product, input_amount, ft.IconButton(ft.icons.ADD, on_click=add_row)]),
            ft.Divider(),
            ft.Text("已添加列表:"),
            data_table,
            ft.ElevatedButton(
                "生成 Excel 并一键分享",
                icon=ft.icons.SHARE,
                on_click=generate_and_share,
                style=ft.ButtonStyle(
                    color=ft.colors.WHITE,
                    bgcolor=ft.colors.BLUE,
                ),
                width=400
            )
        ], scroll=ft.ScrollMode.AUTO)
    )

if __name__ == "__main__":
    # 如果是在本地开发环境
    # ft.app(target=main) 
    
    # 如果是打包成手机 App 或在移动端运行
    ft.app(target=main)
