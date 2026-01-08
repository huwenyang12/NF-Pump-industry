import json
import re
import os
from openpyxl import load_workbook


def get_merged_value(ws, cell_ref):
    """读取合并单元格值（取合并区域左上角）"""
    cell = ws[cell_ref]
    if cell.value is not None:
        return str(cell.value).replace("\n", "").strip()
    for merged in ws.merged_cells.ranges:
        if cell.coordinate in merged:
            val = ws.cell(merged.min_row, merged.min_col).value
            return str(val).replace("\n", "").strip() if val is not None else ""
    return ""


def find_all_block_starts(ws):
    """扫描所有订单块表头行：H列包含 物料号"""
    starts = []
    for row in range(1, ws.max_row + 1):
        val = ws[f"H{row}"].value
        if isinstance(val, str) and re.search(r"物料\s*号|物料\s*编\s*号", val):
            starts.append(row)
    return starts


def parse_order_excel(file_path):
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active

    all_orders = []
    block_starts = find_all_block_starts(ws)

    for idx, start_row in enumerate(block_starts):
        # 关键：限定本订单块的解析范围，防止串到下一单
        end_row = (block_starts[idx + 1] - 1) if idx < len(block_starts) - 1 else ws.max_row

        # 订单头信息就在表头下一行（同时也是第一条明细行）
        info_row = start_row + 1
        quote = get_merged_value(ws, f"A{info_row}")       # 是否引用
        no = get_merged_value(ws, f"B{info_row}")          # 订单号
        sales_reach = get_merged_value(ws, f"C{info_row}")  # 库位
        storage_number = get_merged_value(ws, f"D{info_row}")  # 客户名称编码
        kehucankao = get_merged_value(ws, f"E{info_row}")  # 客户名称编码
        kehubianma = get_merged_value(ws, f"F{info_row}") 

        order_data = {
            "订单类型": "Z008",
            "销售组织": "1072",
            "分销渠道": "10",
            "产品组": "10",
            "销售办事处": "",
            "销售组": "270",
            "是否引用": quote,
            "订单号": no,
            "售达方": sales_reach,
            "库位": storage_number,
            "客户参考": kehucankao,
            "客户名称编码": kehubianma,
            "是否安装调试验收": "",
            "备注1": "",
            "备注2": "",
            "退货行号列表": [],
            "items": []
        }

        row = start_row + 1
        while row <= end_row:
            f_text = get_merged_value(ws, f"H{row}")  # 物料号列（核心列）

            # 结束标记：是否安装调试验收（模板里文字在F列区域）
            if "是否安装调试验收" in f_text:
                # 取下一行 F 列的值（你截图里就是 N）
                order_data["是否安装调试验收"] = get_merged_value(ws, f"H{row + 1}")

                # 备注1 / 备注2：G列为“备注1/备注2”，H列为值
                for r in range(row, min(row + 8, end_row + 1)):
                    g_val = get_merged_value(ws, f"G{r}")
                    if "备注1" in g_val:
                        order_data["备注1"] = get_merged_value(ws, f"I{r}")
                    if "备注2" in g_val:
                        order_data["备注2"] = get_merged_value(ws, f"I{r}")
                break

            # 跳过空行
            if not f_text:
                row += 1
                continue

            # 避免把表头“物料号”当物料行
            if re.search(r"物料\s*号|物料\s*编\s*号", f_text):
                row += 1
                continue

            # 只认“纯数字”的物料号，过滤“合计/客户组”等杂行
            if not re.fullmatch(r"\d{6,30}", f_text):
                row += 1
                continue

            item = {
                "项目行号": get_merged_value(ws, f"G{row}"),
                "物料号": get_merged_value(ws, f"H{row}"),
                "配件型号": get_merged_value(ws, f"I{row}"),
                "序列号": get_merged_value(ws, f"J{row}"),
                "订单数量": get_merged_value(ws, f"K{row}"),
                "金额": get_merged_value(ws, f"L{row}"),
                "工厂": get_merged_value(ws, f"M{row}"),
                "采购订单编号": get_merged_value(ws, f"N{row}"),
                "拒绝原因": get_merged_value(ws, f"O{row}"),
                "显示抬头详细信息": get_merged_value(ws, f"P{row}"),
                "物料文本": get_merged_value(ws, f"Q{row}"),
            }
            order_data["items"].append(item)
            line_no = str(item.get("项目行号", "")).strip()
            if line_no: order_data["退货行号列表"].append(line_no)
            row += 1

        # 没有明细就不加（防止误扫到空块）
        if order_data["items"]:
            all_orders.append(order_data)

    return all_orders


if __name__ == "__main__":
    file_path = r"D:\青臣云起\项目\南方流体模板解析\文件\引用不引用退货单模板.xlsx"
    data = parse_order_excel(file_path)

    json_path = os.path.join(os.path.dirname(file_path), "json数据解析.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

    print(f"解析文件 {file_path} 完成，订单数：", len(data))
