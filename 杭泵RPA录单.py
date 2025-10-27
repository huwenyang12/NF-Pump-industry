import json
import re
import os
from openpyxl import load_workbook
from itertools import groupby

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
        receiver = get_merged_value(ws, f"J{start_row + 2}")

        order_data = {
            "销售组织": "3900",
            "分销渠道": "10",
            "产品组": "10",
            "销售组": "330",
            "名称": name,
            "收货人信息": receiver,
            "是否安装调试": "",
            "备注1": "",
            "备注2": "",
            "items": []
        }

        # 表格中数据从表头下一行开始
        row = start_row + 1
        while row <= ws.max_row:
            text = str(ws[f"C{row}"].value or "").strip()

            # 遇到 "是否安装调试验收" 说明子订单结束
            if "是否安装调试验收" in text:
                next_val = get_merged_value(ws, f"C{row + 1}")
                order_data["是否安装调试"] = next_val
                # 开始往下找 D 列备注1 / 备注2
                for i in range(row + 1, row + 6):  # 往下查几行范围
                    d_val = str(ws[f"D{i}"].value or "").strip()
                    if "备注1" in d_val:
                        # 提取冒号后的文字
                        order_data["备注1"] = d_val.split("：")[-1].strip() if "：" in d_val else ""
                    if "备注2" in d_val:
                        order_data["备注2"] = d_val.split("：")[-1].strip() if "：" in d_val else ""
                break

            # 跳过空行或非物料号
            if not text:
                row += 1
                continue

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

    # 第二阶段：根据“工厂”分组再拆分子订单
    grouped_orders = []
    for order in all_orders:
        # 使用 groupby 前要确保 items 是按工厂排序的
        items_sorted = sorted(order["items"], key=lambda x: x["工厂"])
        for factory, group in groupby(items_sorted, key=lambda x: x["工厂"]):
            new_order = order.copy()
            new_order["items"] = list(group)
            grouped_orders.append(new_order)

    return grouped_orders


if __name__ == "__main__":
    file_path = r"D:\青臣云起\项目\南方流体模板解析\发货单.xlsx"
    result = parse_order_excel(file_path)

    # 保存为 JSON 文件
    json_path = os.path.splitext(file_path)[0] + "_result.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=4)

    print(f"解析完成，结果已保存到：{json_path}")
