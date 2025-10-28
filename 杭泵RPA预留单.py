import json
import re
import os
from openpyxl import load_workbook

def get_merged_value(ws, cell_ref):
    """读取合并单元格值"""
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
        val = ws[f"C{row}"].value
        if isinstance(val, str) and re.search(r"物料 *号|物料编 *号", val):
            block_starts.append(row)
    return block_starts

def parse_order_excel(file_path):
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    all_orders = []
    block_starts = find_all_block_starts(ws)

    for start_row in block_starts:
        name = get_merged_value(ws, f"B{start_row + 3}")
        receiver = get_merged_value(ws, f"H{start_row + 3}")
        warehouse_number = get_merged_value(ws, f"I{start_row + 3}")
        match = re.match(r"(\d+)", warehouse_number)
        number = match.group(1) if match else ""

        order_data = {
            "移动类型": "311",
            "工厂": "3900",
            "营销仓库位号": warehouse_number,
            "收货存储地点": number,
            "单位名称": name,
            "收货人信息": receiver,
            "items": []
        }

        # 表格中数据从表头下一行开始
        row = start_row + 1
        while row <= ws.max_row:
            text = str(ws[f"C{row}"].value or "").strip()

            # 读取到下一物料号本子订单结束
            if re.search(r"物料 *号|物料编 *号", text):
                break

            # 跳过空行或非物料号
            if not text:
                row += 1
                continue

            item = {
                "物料号": get_merged_value(ws, f"C{row}"),
                "水泵型号": get_merged_value(ws, f"D{row}"),
                "数量": get_merged_value(ws, f"E{row}"),
                "单价": get_merged_value(ws, f"F{row}"),
                "金额": get_merged_value(ws, f"G{row}"),
            }
            order_data["items"].append(item)
            row += 1

        all_orders.append(order_data)


    return all_orders


if __name__ == "__main__":
    file_path = r"D:\青臣云起\项目\南方流体模板解析\预留单.xlsx"
    result = parse_order_excel(file_path)

    # 保存为 JSON 文件
    json_path = os.path.splitext(file_path)[0] + "_result.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=4)

    print(f"解析完成，结果已保存到：{json_path}")
