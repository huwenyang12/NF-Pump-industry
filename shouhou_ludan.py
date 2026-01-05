from datetime import datetime
import os, time
import shouhou_dtParser as dtParser
import shouhou_sap as sap
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

def checkProcessedFile(orderNums,processed,newPath):
    log.logger.info("核对文件：" + newPath)
    mypath=None
    for i in range(orderNums.__len__()):
        if orderNums[i].startswith('200') or orderNums[i].startswith('220'):
            mypath=sap.searchFile(orderNums[i],newPath,i)
            download_file=os.path.join(mypath,"EXPORT.XLSX")
            os.remove(download_file)#删除sap导出表格
        elif orderNums[i].startswith('450'):
            mypath=sap.searchFile_me2n(orderNums[i], newPath,i)
            download_file=os.path.join(mypath,"EXPORT.XLSX")
            os.remove(download_file)#删除sap导出表格
    return True

def processFile(file, 订单记录路径):
    log.logger.info("处理文件：" + file)
    orders = dtParser.GetOrderItems(file)
    if orders==[]:
        raise ExcelDataException('1000','表格数据解析失败，请核查后重新投递')
    sap.logon()
    errorMsg = ""
    fName =""
    ids = []
    #从解析的表格内容中根据客户名称分流
    departments=[]#依据的客户分流表
    departNums=[]
    res=None
    try:
        filename = os.path.join(os.getcwd(),'办事处库存地点.xlsx')
        mydf=pd.read_excel(filename,sheet_name=0)
        for idx,row in mydf.iterrows():
            departNums.append(str(mydf.loc[idx, row.index[1]]))
            departments.append(str(mydf.loc[idx, row.index[2]]))
        jiaoqi = None
        for o in orders:
            #判断交期格式以及是否比今天晚
            '''
            try:
                jiaoqi = dtParser.parseDate(dtParser.getDate(o["交期"]))
            except:
                try:
                    jiaoqi = dtParser.parseDate(o["交期"])
                except:
                    raise ExcelDataException("1007","交期格式不合法")            
            jiaoqi = datetime.strptime(jiaoqi, '%Y.%m.%d')
            offset = datetime.now(tz=jiaoqi.tzinfo).replace(hour=0, minute=0, second=0, microsecond=0) - jiaoqi
            if offset.days > 0:
                raise Exception("交期已过，跳过录单")
            '''
            log.logger.info("创建订单：" + o["编号"] + "  "+o["客户"])
            #ME21N流程
            if len(o['items']) > 20:
                raise Exception("ME21N 物料个数超过20个,烦请人工录入")
            o['地点']=departNums[departments.index(o["客户"])]

            my_o = []
            my_keys = []  # 存“工厂_分组类型”的组合key，不再只存工厂
            for item in o['items']:  # 直接遍历items，比用range更简洁
                factory = item['工厂']
                material_no = item['料号']
                part_model = item.get('配件型号', '')  # 拿配件型号，没有就为空
                
                # 按你的规则判断分组类型
                if material_no.startswith('12'):
                    group_type = '12'
                elif material_no.startswith('15'):
                    group_type = '15'
                elif material_no.startswith('14') and '内芯' in part_model:
                    group_type = '14内芯'
                else:
                    group_type = 'other'
                
                # 生成组合key（工厂+分组类型，确保同工厂内不同规则的物料分开）
                key = f"{factory}_{group_type}"
                
                # 按组合key分组
                if key not in my_keys:
                    my_o.append({
                        "编号": o["编号"],
                        "客户": o["客户"],
                        "收货人": o["收货人"],
                        "订单类型": o["订单类型"],
                        "地点": o["地点"],
                        "items": [item]
                    })
                    my_keys.append(key)
                else:
                    index = my_keys.index(key)
                    my_o[index]["items"].append(item)

            id=None
            res_list = []
            for new_o in my_o:
                if "退货单" in file:
                    new_o['单据类型'] = "退货单"
                else:
                    new_o['单据类型'] = "发货单"

                # 重新设置订单类型
                if not new_o["items"]:
                    raise ExcelDataException("1009", "分单后的 items 为空，无法判断订单类型")

                item_factory = str(new_o["items"][0]["工厂"]).strip()

                ZUB_FACTORIES = {"1072", "1073", "1079"}
                ZNB_FACTORIES = {"3510", "3520", "1100"}

                if item_factory in ZUB_FACTORIES:
                    new_o["订单类型"] = "ZUB"
                elif item_factory in ZNB_FACTORIES:
                    new_o["订单类型"] = "ZNB"
                else:
                    raise ExcelDataException("1008", f"未知工厂类型：{item_factory}（无法判断订单类型）")

                # 工厂分类后解析的数据        
                output_path = "工厂分类后解析.json"
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(my_o, f, ensure_ascii=False, indent=2)
                print(f"数据提取完成，已保存为：{output_path}")

                res=sap.createME21N(new_o)
                if not res or str(res).strip() == "":
                    log.logger.warning(f"录单结果为空，可能界面卡死，准备重试一次：订单 {new_o['编号']}")
                    for retry_i in range(2):  # 最多再试两次
                        time.sleep(2)
                        res = sap.createME21N(new_o)
                        if res and str(res).strip() != "":
                            log.logger.info(f"重试录单成功，结果：{res}")
                            break
                        else:
                            log.logger.warning(f"第 {retry_i+1} 次重试仍返回空，继续重试...")
                    if not res or str(res).strip() == "":
                        log.logger.error(f"连续3次录单结果为空，放弃此订单：{new_o['编号']}")
                        res = "录单后结果为空，SAP未响应"
                print(res)
                if res.__contains__("已保存") or res.__contains__("已创建"):
                    id = re.sub("\D", "", str(res))
                    ids.append(id)
                    # 1) 写台账：采购订单号一行
                    try:
                        dtParser.toexcel_id(订单记录路径, id)
                        log.logger.info(f"写入台账成功：{订单记录路径} -> {id}")
                    except Exception as e:
                        log.logger.warning(f"写台账失败：{e}")
                        try:
                            backup_dir = os.path.dirname(订单记录路径)
                            os.makedirs(backup_dir, exist_ok=True)
                            backup_txt = os.path.join(backup_dir, "录单结果备份.txt")

                            with open(backup_txt, "a", encoding="utf-8") as f:
                                f.write(f"{datetime.now()} {id}\n")
                            log.logger.info(f"已写入备份：{backup_txt} -> {id}")
                        except Exception as e2:
                            log.logger.warning(f"写备份文件失败：{e2}")
                    # 2) 写快照：snapshots/{PO}.json（保存该子订单 items）
                    try:
                        snap_path = dtParser.save_snapshot(订单记录路径, id, new_o, source_file=file)
                        log.logger.info(f"保存快照成功：{snap_path}")
                    except Exception as e:
                        log.logger.warning(f"保存快照失败：{e}")
                elif any(k in res for k in ["未被维护", "物料", "工厂", "价格没有维护"]):
                    log.logger.error(f"订单 {new_o['编号']} 录单失败（业务异常）：{res}")
                    id = None
                else:
                    res = str(res)
                    raise Exception(res.replace(' ', '').replace('/', '').replace('\n', ''))
                res_list.append(res)
             
            errors = ""  
            wait_retry = 0
            while ("价格没有维护" in res) or ("A版价格" in res) or (res == ""):
                cc.send_hotkey('{ENTER}')
                time.sleep(0.5)
                res = cc.sap.find_element(locator.sap.items.result_panel).get_text() or ""
                res = res.strip()
                if not (("价格没有维护" in res) or ("A版价格" in res) or (res == "")):
                    cc.find_element(locator.sap.order.exit_1).click()
                    time.sleep(0.5)
                    cc.find_element(locator.sap.order.exit_1).click()
                    break
                wait_retry += 1
                if wait_retry > 100:
                    cc.find_element(locator.sap.order.exit_1).click()
                    time.sleep(0.5)
                    cc.find_element(locator.sap.order.exit_1).click()
                    break
            for r0 in res_list:
                if not (("已保存" in r0) or ("已创建" in r0)):
                    errorMsg = r0
                    errors = f"内层异常捕获：{str(r0).replace('/','').replace(' ','')}"
                    log.logger.warning(errors)
                    break  # 找到一个异常就够了（可按你需求改成不 break）
            if errors:
                log.logger.warning(f"发现部分异常: {errors}, 但继续执行。")

    except Exception as e:
        log.logger.error(traceback.format_exc())
        fName = "-".join(ids)
        return orders, fName, "errors:" + str(e)

    fName = "-".join(ids)
    return orders, fName, errorMsg

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

    return target_files


def short_msg(msg: str, n: int = 30) -> str:
    s = (msg or "").strip()
    # 压缩空白，避免 Excel 里显示难看
    s = re.sub(r"\s+", " ", s)
    return s[:n]


def parse_migo(text: str):
    if not text:
        return None
    s = str(text)
    if "已过账" not in s:
        return None
    m = re.search(r"物料凭证\s*(\d+)", s)
    return m.group(1) if m else None     


def is_migo_done_row(row: dict) -> bool:
    return str(row.get("过账状态", "")).strip() == "成功"
    

def update_migo_into_file(订单记录_path: str, 销售凭证: str, 采购订单号: str, 物料凭证: str, 过账状态: str):
    date_str = dtParser.get_network_time().split()[0].replace('.', '-')
    mypath = os.path.join(订单记录_path, f"{date_str}.xlsx")
    if not os.path.exists(mypath):
        return
    df = pd.read_excel(mypath, dtype=str)
    df.fillna("", inplace=True)

    for col in ["物料凭证", "过账状态"]:
        if col not in df.columns:
            df[col] = ""

    so_str = (销售凭证 or "").strip()
    po_str = (采购订单号 or "").strip()
    mask = (df["销售凭证"].astype(str).str.strip() == so_str) & \
           (df["采购订单号"].astype(str).str.strip() == po_str)

    状态 = (过账状态 or "").strip()
    if 状态 == "成功":
        df.loc[mask, "物料凭证"] = str(物料凭证 or "").strip()
        df.loc[mask, "过账状态"] = "成功"
    else:
        df.loc[mask, "物料凭证"] = ""   # 明确置空，避免旧值残留
        df.loc[mask, "过账状态"] = 状态

    df.to_excel(mypath, index=False)



def handle_migo(site_config):
    record_dir = dtParser._normalize_record_folder(site_config["录单"]["订单记录路径"])
    os.makedirs(record_dir, exist_ok=True)
    rows = dtParser.load_订单记录文件(record_dir)
    if not rows:
        log.logger.info("[MIGO] 台账为空，无需处理")
        return True

    # 逐行处理：每行(PO) 对应一个 SO
    for r in rows:
        so = (r.get("销售凭证", "") or "").strip()
        po = (r.get("采购订单号", "") or "").strip()
        # 基本过滤
        if not so or so == "跳过":
            continue
        if not po:
            continue
        # 去重：过账状态=成功则跳过
        if is_migo_done_row(r):
            continue
        # 取快照 items（一个 PO 对应一个 snapshot）
        snap = dtParser.load_snapshot(record_dir, po)
        items = snap.get("items", []) if snap else []
        if not items:
            log.logger.warning(f"[MIGO] 快照无 items：PO={po}, SO={so}")
            continue
        log.logger.info(f"[MIGO] 准备执行：SO={so}, PO={po}, item数={len(items)}")
        # 登录SAP
        sap.logon()
        # 执行 MIGO
        migo_msg = ""
        try:
            migo_msg = sap.migo_收货(so, items) or ""
        except Exception as e:
            migo_msg = f"MIGO异常：{e}"
            log.logger.error(f"[MIGO] 异常：SO={so}, PO={po}, err={e}")
        # 解析结果
        material_doc = parse_migo(migo_msg)
        # 回写台账
        if material_doc:
            update_migo_into_file(record_dir, so, po, material_doc, "成功")
            log.logger.info(f"[MIGO] 成功：SO={so}, PO={po}, 物料凭证={material_doc}")
        else:
            reason = short_msg(migo_msg, 30) or "失败"
            update_migo_into_file(record_dir, so, po, "", reason)
            log.logger.warning(f"[MIGO] 失败：SO={so}, PO={po}, 返回={migo_msg}")
    return True



def handle_jiaohuodan(site_config):
    root_folder = site_config["录单"]["订单记录路径"]
    rows = dtParser.load_订单记录文件(root_folder)
    for r in rows:
        if r['销售凭证'] != "": continue
        if r['采购订单号'] == "": continue
        销售凭证 = sap.vl10b_销售凭证(r['采购订单号'])
        if 销售凭证 is None or 销售凭证 == "":
            销售凭证 = "跳过"
        dtParser.update_凭证_into_file(root_folder, r['采购订单号'], 销售凭证)

#录单单个文件
def handle(site_config,f):
    订单记录路径 = site_config["录单"]["订单记录路径"]

    root_folder = site_config["录单"]["root路径"]
    pendingFolder = os.path.join(root_folder, "待处理")
    processed  = os.path.join(root_folder, "已处理")
    failed = os.path.join(root_folder, "失败")
    isChecked = os.path.join(root_folder, "已核对")
    backup_path = os.path.join(root_folder, "备份")

    myname=os.path.basename(f)
    if myname.startswith('~') or  (not f.endswith('xlsx') and not f.endswith('xls')):
        log.logger.info("文件{}不是excel文件,忽略".format(f))
        return
    filePath = os.path.join(pendingFolder, f)
    backup_file = os.path.join(backup_path, myname)
    if not os.path.exists(backup_file):
        shutil.copy(filePath, backup_file)
        time.sleep(1)
    mystr=myname.split('_')
    mystr=mystr[0]
    orders = []
    try:
        filename=os.path.basename(f)
        filePath = os.path.join(pendingFolder, f)
            
        #print("处理文件：" + filePath)
        if os.path.exists("C:\TEMP\EXPORT.XLSX"):
            os.remove("C:\TEMP\EXPORT.XLSX")

        orders, newName, errMsg = processFile(filePath, 订单记录路径)
        log.logger.info(f'errMsg: {errMsg}')
        myName=mystr+newName
        if errMsg != "":
            errFile = os.path.join(failed, f+".txt")
            fwriter = open(errFile, "a")
            fwriter.write(myName)
            fwriter.write(errMsg)
            fwriter.close()
            ext = f.split(".")[-1]
            newPath = os.path.join(failed, myName+"."+ext)
            log.logger.info(f"move {filePath} to {newPath}")
            shutil.move(filePath, newPath)

            #处理失败，记录到数据库
            util.insert_record(filePath, orders, -1, 2, errMsg)
        else:#处理无误，继续后续操作
            ext = f.split(".")[-1]
            #复制一份源文件到核对文件夹
            newPath = os.path.join(processed, myName+"."+ext)
            checkPath=os.path.join(isChecked, myName+"."+ext)
            #shutil.copy(filePath,checkPath)
            failPath=os.path.join(failed, myName+"."+ext)
            log.logger.info(f"move {filePath} to {newPath}")
            shutil.move(filePath, newPath)
            time.sleep(3)
            #核对
            log.logger.info('开始核对文件')
            orderNums=newName.split('-')
            flag=None
            errMsg = ''
            try:
                #flag=checkProcessedFile(orderNums,processed,newPath)
                flag = True
            except Exception as e:
                log.logger.error("核对失败："+traceback.format_exc())
                #log.logger.info(f"move {newPath} to {failPath}")
                #shutil.move(newPath,failPath)
                errMsg="核对出现错误:" + str(e)
                errFile = os.path.join(failed, f+".txt")
                fwriter = open(errFile, "a")
                fwriter.write(myname)
                fwriter.write(errMsg)
                fwriter.close()
            if flag==True:
                #log.logger.info(f"move {newPath} to {checkPath}")
                #shutil.move(newPath, checkPath)
                util.insert_record(filePath, orders, 0, 0, '')
            else:
                #log.logger.info(f"move {newPath} to {failPath}")
                #shutil.move(newPath,failPath)
                util.insert_record(filePath, orders, -1, 1, errMsg)          
    except Exception as e:
        log.logger.error("处理订单失败："+traceback.format_exc())
        error="处理订单失败：" + str(e).replace(' ','').replace('/','').replace('?','').replace('<','').replace('>','').replace('|','').replace(':','')
        
        util.insert_record(filePath, orders, -1, 3, error)
        errFile = os.path.join(failed, f+".txt")
        if os.path.exists(errFile):
            os.remove(errFile)
        fwriter = open(errFile, "a")
        fwriter.write(myname)
        fwriter.write(error)
        fwriter.close()
        failedPath=os.path.join(failed, mystr+".xlsx")
        if os.path.exists(failedPath):
            os.remove(failedPath)
            time.sleep(2)
        if os.path.exists(filePath):
            log.logger.info(f"move {filePath} to {failedPath}")
            shutil.move(filePath, failedPath)
        elif os.path.exists(newPath):
            log.logger.info(f"move {newPath} to {failedPath}")
            shutil.move(newPath, failedPath)
    time.sleep(3)

if __name__ == "__main__":
    file = r"C:\Users\robot\Desktop\测试分组.xlsx"
    processFile(file)