import json
import re
from openpyxl import load_workbook

def get_merged_value(ws, cell_ref):
    """读取合并单元格值（支持跨表格合并单元格）"""
    cell = ws[cell_ref]
    if cell.value is not None:
        return str(cell.value).replace("\n", "")
    for merged in ws.merged_cells.ranges:
        if cell.coordinate in merged:
            val = ws.cell(merged.min_row, merged.min_col).value
            return str(val).replace("\n", "") if val else ""
    return ""

def find_all_block_starts(ws):
    """自动扫描所有子订单起始行"""
    block_starts = []
    for row in range(1, ws.max_row + 1):
        cell_val = ws[f"C{row}"].value
        if isinstance(cell_val, str) and "物料号" in cell_val:  # 子订单表头标志
            block_starts.append(row)
    return block_starts

def parse_order_excel(file_path):
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    all_orders = []
    block_starts = find_all_block_starts(ws)

    for start_row in block_starts:
        # 取相对行的值（每个子订单结构相同）
        name = get_merged_value(ws, f"B{start_row + 3}")
        receiver = get_merged_value(ws, f"J{start_row + 2}")
        is_install = get_merged_value(ws, f"C{start_row + 10}")
        remark1 = get_merged_value(ws, f"D{start_row + 9}")
        remark2 = get_merged_value(ws, f"D{start_row + 10}")

        order_data = {
            "销售组织": "3900",
            "分销渠道": "10",
            "产品组": "10",
            "销售组": "330",
            "名称": name,
            "收货人信息": receiver,
            "是否安装调试": is_install,
            "备注1": remark1,
            "备注2": remark2,
            "items": []
        }

        # 表格中数据从表头下一行开始
        row = start_row + 1
        while row <= ws.max_row:
            material_no = ws[f"C{row}"].value
            pump_model = ws[f"D{row}"].value

            # 如果读到空行或下一个“物料号”，说明子订单结束
            if not material_no:
                break
            if not re.fullmatch(r"\d+", str(material_no).strip()):
                break
            item = {
                "物料号": get_merged_value(ws, f"C{row}"),
                "水泵型号": get_merged_value(ws, f"D{row}"),
                "工厂": get_merged_value(ws, f"E{row}"),
                "数量": get_merged_value(ws, f"F{row}"),
                "单价": get_merged_value(ws, f"G{row}"),
                "金额": get_merged_value(ws, f"H{row}"),
                "交期": get_merged_value(ws, f"I{row}")
            }
            order_data["items"].append(item)
            row += 1

        all_orders.append(order_data)

    return all_orders


if __name__ == "__main__":
    file_path = r"D:\青臣云起\项目\南方流体模板解析\发货单.xlsx"
    result = parse_order_excel(file_path)
    print(json.dumps(result, ensure_ascii=False, indent=4))
