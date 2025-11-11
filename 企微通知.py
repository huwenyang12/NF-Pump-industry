from datetime import datetime
import os, time
import dtParser
import sap
import shutil, re
import log, util
from shutil import copytree, rmtree
import json
from openpyxl import Workbook,load_workbook
import traceback
import pandas as pd
from exceptions import *
from clicknium import clicknium as cc, locator
from myFtp import myFtp
import dingtalk_utils

def checkProcessedFile(orderNums, processed, newPath):
    """
    核对已创建订单是否在 SAP 中存在并导出表格，成功后删除导出的临时文件
    :param orderNums: 成功订单号列表
    :param processed: 已处理文件夹路径（可用于查找历史文件）
    :param newPath: 当前处理的文件路径（可用于下载导出表格）
    :return: True 成功，False 失败（若出现异常会抛出）
    """
    log.logger.info(f"开始核对文件: {newPath}")

    for idx, orderNum in enumerate(orderNums):
        if not orderNum:
            continue  # 防止空值

        log.logger.info(f"核对订单号: {orderNum}")

        # 使用 SAP 客户端统一查询导出路径
        try:
            mypath = sap.searchFile(orderNum, newPath, idx)
            download_file = os.path.join(mypath, "EXPORT.XLSX")

            if not os.path.exists(download_file):
                raise FileNotFoundError(f"SAP导出表格未找到: {download_file}")

            # 核对逻辑可以扩展：比如比对 Excel 内容是否与订单号匹配
            # ...

            # 核对完成后删除临时导出文件
            os.remove(download_file)
            log.logger.info(f"订单 {orderNum} 核对完成，已删除临时文件")
        except Exception as e:
            log.logger.error(f"订单 {orderNum} 核对失败: {str(e)}")
            raise e  # 出现任何异常，整个核对流程失败

    return True


def processFile(file):
    log.logger.info("处理文件：" + file)

    prefix = os.path.splitext(os.path.basename(file))[0]
    orders = dtParser.GetOrderItems(file)
    if not orders:
        raise ExcelDataException('1000', '表格数据解析失败，请核查后重新投递')

    sap.logon()

    success_orders = []
    failed_info = []
    skipped_info = []

    # 进度日志目录
    log_dir = os.path.join(os.path.dirname(file), "进度日志")
    os.makedirs(log_dir, exist_ok=True)
    progress_log = os.path.join(log_dir, f"{prefix}_进度.txt")

    # 读取已处理记录（编号+客户名称）
    processed_combinations = set()
    if os.path.exists(progress_log):
        with open(progress_log, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split(",")
                if len(parts) >= 4:
                    processed_combinations.add(f"{parts[2]}|{parts[3]}|{parts[1]}")

    for o in orders:
        combo_key = f"{o['编号']}|{o['名称']}|{o['订单类型']}"

        # 跳过已处理
        if combo_key in processed_combinations:
            log.logger.info(f"{o['订单类型']} - {o['编号']} - {o['名称']} 已处理过，跳过。")
            skipped_info.append({"订单类型": o["订单类型"],"编号": o['编号'], "名称": o['名称']})
            with open(progress_log, "a", encoding="utf-8") as fwriter:
                fwriter.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')},{o['订单类型']},{o['编号']},{o['名称']},跳过\n")
            continue

        try:
            log.logger.info(f"创建订单：{o['编号']}  {o['名称']}  {o['订单类型']} ")
            res = sap.createOrder(o)

            if not res:
                raise Exception("sap.createOrder 返回为空")

            if "已保存" in res or "已创建" in res:
                # 提取任意连续数字串作为订单号
                order_id_match = re.search(r'(\d{6,})', res)  # 一般订单号6位以上
                if not order_id_match:
                    raise Exception(f"SAP返回成功但未识别订单号: {res}")
                order_id = order_id_match.group(1)

                # 成功记录
                success_orders.append({
                    "订单类型": o["订单类型"],
                    "编号": o["编号"],
                    "订单号": order_id
                })

                mytime = dtParser.get_network_time().replace('.', '-')
                dtParser.toexcel_id(order_id, mytime)

                # 写日志
                with open(progress_log, "a", encoding="utf-8") as fwriter:
                    fwriter.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')},{o['订单类型']},{o['编号']},{o['名称']},成功,{order_id}\n")
                log.logger.info(f"成功创建：{o['订单类型']}  - {o['编号']} - {o['名称']} 订单号: {order_id}")


                processed_combinations.add(combo_key)
                log.logger.info(f"订单 {o['订单类型']} {o['编号']} 创建成功, 订单号：{order_id}\n")

            else:
                # 其他异常
                res_clean = res.replace('/', '').replace(' ', '')
                raise Exception(f"SAP返回异常: {res_clean}")

        except Exception as e:
            log.logger.error(f"{o['订单类型']} {o['编号']} 创建失败: {str(e)}")

            match = re.search(r'订单\s*(\d+)', str(e))
            failed_order_no = match.group(1) if match else ""
            err_msg = str(e).split(":", 1)[-1].strip() if "SAP返回异常:" in str(e) else str(e)

            failed_info.append({
                "订单类型": o["订单类型"],
                "编号": o["编号"],
                "名称": o["名称"],
                "订单号": failed_order_no,
                "错误信息": err_msg
            })

            with open(progress_log, "a", encoding="utf-8") as fwriter:
                fwriter.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')},{o['订单类型']},{o['编号']},{o['名称']},失败,{failed_order_no},{err_msg}\n")

            processed_combinations.add(combo_key)

    # 统计信息
    total_count = len(orders)
    success_count = len(success_orders)
    failed_count = len(failed_info)
    skipped_count = len(skipped_info)

    success_list_str = "\n" + "\n".join([f"  - {o['订单类型']} {x['编号']} ({x['订单号']})" for x in success_orders]) if success_orders else " 无"
    failed_list_str = "\n" + "\n".join([f"  - {o['订单类型']} {x['编号']} ({x['名称']}): {x['错误信息']}" for x in failed_info]) if failed_info else " 无"
    skipped_list_str = "\n" + "\n".join([f"  - {o['订单类型']} {x['编号']} ({x['名称']})" for x in skipped_info]) if skipped_info else " 无"

    overview = (
        f"\n\n=== 总览统计 ===\n"
        f"{prefix} 总录单数为 {total_count} 条，\n"
        f"成功 {success_count} 条:{success_list_str}\n"
        f"失败 {failed_count} 条:{failed_list_str}\n"
        f"跳过 {skipped_count} 条:{skipped_list_str}\n"
    )

    errorMsg_test = (
        f"{prefix} 总录单数为 {total_count} 条，成功 {success_count} 条, 失败 {failed_count} 条, 跳过 {skipped_count} 条\n"
        f"失败 {failed_count} 条:{failed_list_str}\n"
    )

    with open(progress_log, "a", encoding="utf-8") as fwriter:
        fwriter.write(overview + "\n")

    fName = f"{prefix}_{time.strftime('%Y%m%d%H%M%S')}"
    errorMsg = "" if failed_count == 0 else errorMsg_test
    return orders, fName, errorMsg, overview, success_orders, failed_info, skipped_info

def processFile_yuliu(file):
    log.logger.info("处理文件：" + file)

    prefix = os.path.splitext(os.path.basename(file))[0]
    orders = dtParser.GetOrderItems_yuliu(file)
    if not orders:
        raise ExcelDataException('1000', '表格数据解析失败，请核查后重新投递')

    sap.logon()

    success_orders = []
    failed_info = []
    skipped_info = []

    # 进度日志目录
    log_dir = os.path.join(os.path.dirname(file), "进度日志")
    os.makedirs(log_dir, exist_ok=True)
    progress_log = os.path.join(log_dir, f"{prefix}_进度.txt")

    # 读取已处理记录（编号+客户名称）
    processed_combinations = set()
    if os.path.exists(progress_log):
        with open(progress_log, "r", encoding="utf-8") as f:
            for line in f:
                parts = line.strip().split(",")
                if len(parts) >= 4:
                    processed_combinations.add(f"{parts[2]}|{parts[3]}|{parts[1]}")

    for o in orders:
        combo_key = f"{o['编号']}|{o['单位名称']}"

        # 跳过已处理
        if combo_key in processed_combinations:
            log.logger.info(f"{o['编号']} - {o['单位名称']} 已处理过，跳过。")
            skipped_info.append({"编号": o['编号'], "单位名称": o['单位名称']})
            with open(progress_log, "a", encoding="utf-8") as fwriter:
                fwriter.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')},{o['编号']},{o['单位名称']},跳过\n")
            continue

        try:
            log.logger.info(f"创建订单：{o['编号']}  {o['单位名称']}")
            res = sap.MB21(o)

            if not res:
                raise Exception("sap.createOrder 返回为空")

            if "已记账" in res:
                # 提取任意连续数字串作为订单号
                order_id_match = re.search(r'(\d{6,})', res)  # 一般订单号6位以上
                if not order_id_match:
                    raise Exception(f"SAP返回成功但未识别记账号: {res}")
                order_id = order_id_match.group(1)

                # 成功记录
                success_orders.append({
                    "编号": o["编号"],
                    "订单号": order_id
                })

                mytime = dtParser.get_network_time().replace('.', '-')
                dtParser.toexcel_id(order_id, mytime)

                # 写日志
                with open(progress_log, "a", encoding="utf-8") as fwriter:
                    fwriter.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')},{o['编号']},{o['单位名称']},成功,{order_id}\n")
                log.logger.info(f"成功创建：{o['编号']} - {o['单位名称']} 订单号: {order_id}")


                processed_combinations.add(combo_key)
                log.logger.info(f"订单 {o['编号']} 创建成功, 订单号：{order_id}\n")

            else:
                # 其他异常
                res_clean = res.replace('/', '').replace(' ', '')
                raise Exception(f"SAP返回异常: {res_clean}")

        except Exception as e:
            log.logger.error(f"{o['编号']} 创建失败: {str(e)}")

            match = re.search(r'订单\s*(\d+)', str(e))
            failed_order_no = match.group(1) if match else ""
            err_msg = str(e).split(":", 1)[-1].strip() if "SAP返回异常:" in str(e) else str(e)

            failed_info.append({
                "编号": o["编号"],
                "单位名称": o["单位名称"],
                "订单号": failed_order_no,
                "错误信息": err_msg
            })

            with open(progress_log, "a", encoding="utf-8") as fwriter:
                fwriter.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')},{o['编号']},{o['单位名称']},失败,{failed_order_no},{err_msg}\n")
            processed_combinations.add(combo_key)

    # 统计信息
    total_count = len(orders)
    success_count = len(success_orders)
    failed_count = len(failed_info)
    skipped_count = len(skipped_info)

    success_list_str = "\n" + "\n".join([f"  - {x['编号']}({x['订单号']})" for x in success_orders]) if success_orders else " 无"
    failed_list_str = "\n" + "\n".join([f"  - {x['编号']}({x['单位名称']}): {x['错误信息']}" for x in failed_info]) if failed_info else " 无"
    skipped_list_str = "\n" + "\n".join([f"  - {x['编号']}({x['单位名称']})" for x in skipped_info]) if skipped_info else " 无"

    overview = (
        f"\n\n=== 总览统计 ===\n"
        f"{prefix} 总录单数为 {total_count} 条，\n"
        f"成功 {success_count} 条:{success_list_str}\n"
        f"失败 {failed_count} 条:{failed_list_str}\n"
        f"跳过 {skipped_count} 条:{skipped_list_str}\n"
    )

    errorMsg_test = (
        f"{prefix} 总录单数为 {total_count} 条，成功 {success_count} 条, 失败 {failed_count} 条, 跳过 {skipped_count} 条\n"
        f"失败 {failed_count} 条:{failed_list_str}\n"
    )

    with open(progress_log, "a", encoding="utf-8") as fwriter:
        fwriter.write(overview + "\n")

    fName = f"{prefix}_{time.strftime('%Y%m%d%H%M%S')}"
    errorMsg = "" if failed_count == 0 else errorMsg_test

    return orders, fName, errorMsg, overview, success_orders, failed_info, skipped_info

#获取需要录单的文件
def get_handle_files(root_folder):
    pendingFolder = os.path.join(root_folder, "待处理")
    files = os.listdir(pendingFolder)
    if len(files) > 0:
        files=dtParser.getEarlyFiles(root_folder)
    target_files = []
    for f in files:
        base_name = os.path.basename(f)
        if base_name.startswith('~') or  (not f.endswith('xlsx') and not f.endswith('xls')):
            continue
        target_files.append(f)
        print(target_files)
    return target_files

def send_order_summary_notify(site_config, file_name, orders, success_order, failed_info, skipped_info=None):
    try:
        total_count = len(orders)
        success_count = len(success_order)
        failed_count = len(failed_info)
        skipped_count = len(skipped_info) if skipped_info else 0
        summary_line = f"子订单总数：{total_count}，成功：{success_count}，失败：{failed_count}，跳过：{skipped_count}"
        notify_text = (
            f"{site_config['name']}: 录单结束\n"
            f"文件名：{os.path.basename(file_name)}\n"
            f"{summary_line}\n"
        )
        if success_count > 0:
            success_lines = [f"{x['订单类型']} {x['编号']}{x['订单号']}）" for x in success_order]
            notify_text += "\n成功订单：\n" + "\n".join(success_lines[:15])
            if len(success_lines) > 15:
                notify_text += f"\n... 共 {len(success_lines)} 条\n"
        if failed_count > 0:
            failed_lines = [f"{x['订单类型']} {x['编号']}（{x['名称']}）: {x['错误信息'][:80]}" for x in failed_info]
            notify_text += "\n失败订单：\n" + "\n".join(failed_lines[:10])
            if len(failed_lines) > 10:
                notify_text += f"\n... 共 {len(failed_lines)} 条\n"
        if skipped_count > 0:
            skipped_lines = [f"{x['订单类型']} {x['编号']}（{x['名称']}）" for x in skipped_info]
            notify_text += "\n重复订单：\n" + "\n".join(skipped_lines[:10])
            if len(skipped_lines) > 10:
                notify_text += f"\n... 共 {len(skipped_lines)} 条\n"
        if len(notify_text) > 1900:
            notify_text = notify_text[:1900] + "\n...（内容过长已截断）"
        dingtalk_utils.send_message(notify_text)
    except Exception as e:
        log.logger.warning(f"[通知] 发送录单结果失败: {e}")

def handle(site_config, f):
    root_folder = site_config["录单"]["root路径"]
    pendingFolder = os.path.join(root_folder, "待处理")
    processed = os.path.join(root_folder, "已处理")
    failed = os.path.join(root_folder, "失败")
    isChecked = os.path.join(root_folder, "已核对")
    backup_path = os.path.join(root_folder, "备份")

    myName = os.path.splitext(f)[0]   # 默认值，避免异常时未定义
    filePath = os.path.join(pendingFolder, f)
    backup_file = os.path.join(backup_path, f)  # 保留完整文件名

    if myName.startswith('~') or (not f.endswith('xlsx') and not f.endswith('xls')):
        log.logger.info(f"文件{f}不是excel文件,忽略")
        return

    if not os.path.exists(backup_file):
        shutil.copy(filePath, backup_file)
        time.sleep(1)

    mystr = os.path.splitext(f)[0].split('_')[0]
    disable_db = site_config.get("录单").get("disable_db")

    # 全局可用
    base = os.path.splitext(os.path.basename(f))[0]
    timestamp = time.strftime("%Y%m%d%H%M%S")

    try:
        tmp_export = r"C:\\TEMP\\EXPORT.XLSX"
        if os.path.exists(tmp_export):
            os.remove(tmp_export)

        if any(k in f for k in ["预留单", "预留通知单"]):
            log.logger.info(f"识别为【预留单】类型文件: {f}")
            orders, newName, errMsg, overview, success_order, failed_info, skipped_info = processFile_yuliu(filePath)
        elif any(k in f for k in ["发货单", "发货通知单", "销售单"]):
            log.logger.info(f"识别为【销售单】类型文件: {f}")
            orders, newName, errMsg, overview, success_order, failed_info, skipped_info = processFile(filePath)
        else:
            raise Exception("无法识别的录单类型")

        myName = newName  # 用 processFile 返回的标准名

        ext = os.path.splitext(f)[1].lstrip(".")
        base = os.path.splitext(f)[0]
        
        # 有部分订单失败
        if errMsg:  
            # 先插入数据库记录，再移动文件【先注释】
            util.insert_record_controlled(filePath, orders, -1, 2, errMsg, disable_db)
            
            errFile = os.path.join(failed, f"{base}_{timestamp}.txt")
            with open(errFile, "a", encoding="utf-8") as fwriter:
                fwriter.write(f"文件: {myName}\n")
                fwriter.write(errMsg + "\n")

            newPath = os.path.join(failed, myName + "." + ext)
            log.logger.info(f"move {filePath} -> {newPath}")
            shutil.move(filePath, newPath)
            log.logger.info(overview)


        else:  # 全部成功
            newPath = os.path.join(processed, myName + "." + ext)
            checkPath = os.path.join(isChecked, myName + "." + ext)
            log.logger.info(f"move {filePath} -> {newPath}")
            shutil.move(filePath, newPath)
            log.logger.info(overview)
            time.sleep(1)

    except Exception as e:  # 文件级别异常
        log.logger.error("处理订单失败：" + traceback.format_exc())
        error = "处理订单失败：" + str(e)
        errFile = os.path.join(failed, f"{base}_{timestamp}.txt")
        with open(errFile, "a", encoding="utf-8") as fwriter:
            fwriter.write(myName + "\n")  # 这里 myName 一定有值（默认 + 覆盖）
            fwriter.write(error + "\n")
        failedPath = os.path.join(failed, mystr + ".xlsx")
        if os.path.exists(filePath):
            shutil.move(filePath, failedPath)
    
    # 录单完成后通知
    send_order_summary_notify(
        site_config=site_config,
        file_name=f,
        orders=orders,
        success_order=success_order,
        failed_info=failed_info,
        skipped_info=skipped_info
    )

    time.sleep(3)

if __name__ == "__main__":
    file = r"C:\Users\robot\Desktop\流体无锡办2025-10-31-03无锡市良泰机械设备有限公司.xlsx"
    processFile(file)