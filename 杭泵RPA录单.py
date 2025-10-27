import json
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

def parse_order_excel(file_path):
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    all_orders = []  # 保存多个子订单
    block_start_rows = [3]  # 第一个子订单起始行（即物料号表头所在行）
    step = 14  # 每个子订单相隔 14 行

    # 自动判断有多少个子订单
    max_row = ws.max_row
    while block_start_rows[-1] + step <= max_row:
        next_row = block_start_rows[-1] + step
        # 如果下一个区域的C列单元格有表头（"物料号"），说明存在新子订单
        if ws[f"C{next_row}"].value and "物料" in str(ws[f"C{next_row}"].value):
            block_start_rows.append(next_row)
        else:
            break

    # 遍历每个子订单区域
    for start_row in block_start_rows:
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

        # 解析物料行，从表头下一行开始（start_row+1）
        row = start_row + 1
        while True:
            material_no = ws[f"C{row}"].value
            pump_model = ws[f"D{row}"].value
            if not (material_no or pump_model):
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
