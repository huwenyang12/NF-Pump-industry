from clicknium import clicknium as cc, locator,ui
from clicknium.common.models.mouselocation import MouseLocation
import subprocess
import time
import pyperclip as pc
import log,util
import dtParser
import os
import traceback

def logon():
    os.system("taskkill /f /im saplogon.exe")
    time.sleep(2)
    info = util.get_ludan_sap()
    connection = 'ECQ-500' if info["is_test"] else 'cnpsap'
    client = '500' if info["is_test"] else '800'
    cc.sap.login(info["sap_path"], connection, client, info["sap_user"], info["sap_password"])
    time.sleep(3)
    try:
        cc.sap.find_element(locator.sap.logon.continue_logon).click()
        cc.sap.find_element(locator.sap.logon.login_cfm).click()
        cc.find_element(locator.sap.logon.start_close).click()
    except:
        print("...No prompt for login")

def input_text(loc, text):
    cc.wait_appear(loc,wait_timeout=1) 
    cc.sap.find_element(loc).click()
    cc.send_hotkey("^a") 
    cc.send_hotkey("{DEL}") 
    time.sleep(0.1)
    pc.copy(text)
    cc.send_hotkey("^v")
    # cc.send_text(text)
    time.sleep(0.5)
    #cc.send_hotkey("{ENTER}")

#----------------查询流程------------------
def searchFile_me2n(orderNum,newPath,i):
    time.sleep(3)
    log.logger.info("录入ME2N")
    cc.sap.find_element(locator.sap.order.va01).call_transaction("ME2N")
    '''
    input_text(locator.sap.order.va01, "ME2N")
    cc.send_hotkey("{ENTER}")
    '''
    if os.path.exists("C:\TEMP\EXPORT.XLSX"):
        os.remove("C:\TEMP\EXPORT.XLSX")
    time.sleep(0.5)
    log.logger.info("填写查询数据")
    checkedPath=None
    input_text(locator.sap.me2n.date_start, dtParser.get_network_time())
    src_orders = dtParser.GetOrderItems(newPath)#源excel数据
    src_len=src_orders.__len__()
    log.logger.info(f"填写{orderNum}查询并检查")
    input_text(locator.sap.me2n.order_num, orderNum)
    cc.sap.find_element(locator.sap.me2n.search_sure).click()
    time.sleep(0.5)
    if(cc.sap.find_element(locator.sap.me2n.orderNum)):           
        cc.sap.find_element(locator.sap.me2n.orderNum).click(mouse_button="right")
        cc.wait_appear(locator.sap.me2n.to_excel,wait_timeout=10)
        #cc.find_element(locator.sap.me2n.to_excel).highlight(duration=1)
        cc.find_element(locator.sap.me2n.to_excel).click(by='mouse-emulation')
        time.sleep(1)
        cc.send_hotkey('{ENTER}')
        time.sleep(0.5)
    else:
        raise Exception("未能唤起右键二级菜单")
    address_btn=cc.wait_appear(locator.sap.me2n.address_box)
    #address_btn.highlight(duration=1)
    if address_btn!=None:
        #cc.sap.find_element(locator.sap.check.address_sure_button).click()
        log.logger.info("地址按钮")
        checkedPath="C:\TEMP"
    else:
        log.logger.info("保存按钮")
        checkedPath="E:\order"
    cc.send_hotkey("{ENTER}")
    replace_btn=cc.wait_appear(locator.sap.me2n.replace_sure1,wait_timeout=1)
    if replace_btn!=None:
        log.logger.info("第1种")
        #replace_btn.highlight(duration=1)
        replace_btn.click(by="mouse-emulation")
    cc.send_hotkey("{ENTER}")
    log.logger.info("安全性enter")
    time.sleep(0.5)
    
    cc.send_hotkey("{ENTER}")
    log.logger.info("日志enter")
    cc.send_hotkey("{ENTER}")
    log.logger.info("日志enter")
    cc.send_hotkey("{ENTER}")
    log.logger.info("日志enter")
    cc.send_hotkey("{ENTER}")
    log.logger.info("日志enter")
    cc.send_hotkey("{ENTER}")
    log.logger.info("日志enter")
    try:
        time.sleep(1)
        cc.sap.find_element(locator.sap.me2n.me2n_close).click()
        cc.sap.find_element(locator.sap.me2n.search_close).click()
    except:
        log.logger.error("点击关闭失败, ignore:"+traceback.format_exc())

    dwl_path=os.path.join(checkedPath,"EXPORT.XLSX")
    dwl=dtParser.parseExport_me2n(dwl_path)
    try:
        for j in dwl['items']:
            for k in range(src_orders[i]['items'].__len__()):
                if j['wuliao_num']==src_orders[i]['items'][k]['料号'] and j['order_count']==src_orders[i]['items'][k]['数量']:
                    del src_orders[i]['items'][k]
                    break
    finally:
        return r"C:\TEMP"
    # if src_orders[i]['items'].__len__()!=0:
    #     raise Exception("未能完成验证匹配")




def searchFile(orderNum,checkPath,i):
    log.logger.info("录入zsd002")
    input_text(locator.sap.order.va01, "zsd002")
    cc.send_hotkey("{ENTER}")
    time.sleep(0.5)
    log.logger.info("填写查询数据")
    checkedPath=None
    input_text(locator.sap.search.factory_2083, "1073")
    cc.sap.find_element(locator.sap.search.factory_add).click()
    input_text(locator.sap.search.factory_add2, "1074")
    input_text(locator.sap.search.factory_add3, "1079")
    input_text(locator.sap.search.factory_add4, "3510")
    input_text(locator.sap.search.factory_add5, "3520")
    input_text(locator.sap.search.factory_add6, "1078")
    cc.sap.find_element(locator.sap.search.add_sure).click()
    time.sleep(0.5)
    input_text(locator.sap.search.date_start, dtParser.get_network_time())
    #input_text(locator.sap.search.date_end, dtParser.get_network_time())
    src_orders = dtParser.GetOrderItems(checkPath)#源excel数据
    src_len=src_orders.__len__()
    log.logger.info(f"填写{orderNum}查询并检查")
    cc.wait_appear(locator.sap.search.orderNum,wait_timeout=1)
    input_text(locator.sap.search.orderNum, orderNum)
    cc.sap.find_element(locator.sap.search.search_sure).click()
    time.sleep(0.5)
    flag=cc.wait_appear(locator.sap.check.order_num,wait_timeout=10)
    if flag==None:
        raise Exception(f"查询{orderNum}无数据")
    if(cc.sap.find_element(locator.sap.check.order_num)):           
        cc.sap.find_element(locator.sap.check.order_num).click()
        cc.sap.find_element(locator.sap.check.order_num).click(mouse_button="right")
        cc.wait_appear(locator.sap.check.to_excel,wait_timeout=30)
        cc.find_element(locator.sap.check.to_excel).highlight(duration=2)
        cc.find_element(locator.sap.check.to_excel).click(by='mouse-emulation')
        
        time.sleep(0.5)
        #cc.sap.find_element(locator.sap.check.to_excel).click(by='mouse-emulation')
        #cc.sap.find_element(locator.sap.check.excel_sure).highlight(duration=1)
        cc.sap.find_element(locator.sap.check.excel_sure).click(by='mouse-emulation')
        time.sleep(0.5)
    else:
        raise Exception("未能唤起右键二级菜单")

    address_btn=cc.wait_appear(locator.sap.check.address_box,wait_timeout=1)
    if address_btn!=None:
        #cc.sap.find_element(locator.sap.check.address_sure_button).click()
        log.logger.info("地址按钮")
        checkedPath=r"C:\TEMP"
    else:
        log.logger.info("保存按钮")
        checkedPath=r"C:\TEMP"
        #input_text(locator.sap.check.address_input,isChecked)
        
        #cc.sap.find_element(locator.sap.check.address_sure).click()
    time.sleep(0.5)
    cc.send_hotkey("{ENTER}")
    replace_btn=cc.wait_appear(locator.sap.check.replace_sure,wait_timeout=1)
    if replace_btn!=None:
        log.logger.info("第0种")
        #replace_btn.highlight(duration=1)
        replace_btn.click(by="mouse-emulation")
    replace_btn=cc.wait_appear(locator.sap.check.button_允许,wait_timeout=30)
    if replace_btn!=None:
        log.logger.info("第1种")
        #replace_btn.highlight(duration=1)
        replace_btn.click(by="mouse-emulation")
    else:
        log.logger.info("无法点击允许按钮")
    time.sleep(0.5)
    cc.send_hotkey("{ENTER}")
    log.logger.info("安全性enter")
    time.sleep(0.5)
    
    cc.send_hotkey("{ENTER}")
    log.logger.info("日志enter")
    time.sleep(0.5)
    cc.send_hotkey("{ENTER}")
    log.logger.info("日志enter")
    time.sleep(0.5)
    cc.send_hotkey("{ENTER}")
    log.logger.info("日志enter")
    time.sleep(1)
    cc.sap.find_element(locator.sap.check.toback).click()
    time.sleep(1)
    cc.sap.find_element(locator.sap.check.toback).click()

    log.logger.info("正在对比单号")
    dwl_path=os.path.join(checkedPath,"EXPORT.XLSX")
    dwl=dtParser.parseExport(dwl_path)
    print(dwl,src_orders)
    
    for i in range(src_len):
        if dwl["hetong_num"]!=src_orders[i]["合同号"] or dwl['caigou_no'] != src_orders[i]["编号"]:
            continue
        src_orders[i]["客户"]=src_orders[i]["客户"].replace('（','(')
        src_orders[i]["客户"]=src_orders[i]["客户"].replace('）',')')
        if dwl["customer_name"]!=src_orders[i]["客户"]:
            continue
        src_orders[i]['收货人']=src_orders[i]['收货人'].replace(' ','').replace('（','(').replace('）',')').replace('\n','').replace('\r','')
        dwl['deal_address']=dwl['deal_address'].replace(' ','').replace('（','(').replace('）',')').replace('\n','').replace('\r','')
        if src_orders[i]['收货人'] not in dwl['deal_address']:
            continue
        # if dwl['deal_address']!=src_orders[i]['收货人']:
        #     continue
        mytime=dtParser.getDate(src_orders[i]["交期"])
        if mytime==None:
            mytime=dtParser.parseDate(src_orders[i]["交期"])
        else:
            mytime=dtParser.parseDate(dtParser.getDate(src_orders[i]["交期"]))
        if dtParser.parseDate(dwl["deal_date"])!= mytime:
            continue
        src_orders_len=src_orders[i]["items"].__len__()
        for dwl_item in dwl["items"]:
            
            for j in range(src_orders_len):
                print(dwl_item,src_orders[i]["items"])
                if dwl_item["wuliao_num"]==src_orders[i]["items"][j]["料号"]:
                    num=dwl_item["wuliao_num"]
                    if int(float(dwl_item["order_count"]))!=int(float(src_orders[i]["items"][j]["数量"])):
                        raise Exception(f"{num}数量不匹配")
                    #TODO 验证收货地址
                    del src_orders[i]["items"][j]
                    break
        if src_orders[i]["items"].__len__()!=0 and src_orders_len!=src_orders[i]["items"].__len__():
            raise Exception("存在料号未匹配")
        if src_orders[i]["items"].__len__()==0:
            del src_orders[i]
            break

    return r"C:\TEMP"

def checkFile(orderNum):
    log.logger.info(f"填写{orderNum}查询并检查")
    input_text(locator.sap.search.orderNum, orderNum)
    cc.sap.find_element(locator.sap.search.search_sure).click()
    time.sleep(0.5)

def anzhuang_yanshou_set(order):
    if order["订单类型"] == "ZKB": return
    if not order.__contains__('是否安装调试验收'): return
    if order.__contains__('是否安装调试验收') and order['是否安装调试验收'] not in ['Y','N']:
        return
    ui(locator.saplogon.img_state).click()
    cc.sap.find_element(locator.saplogon.对象状态).click()
    if str(order['是否安装调试验收']).upper() == 'N':
        cc.sap.find_element(locator.saplogon.radio_签收).click()
    else:
        cc.sap.find_element(locator.saplogon.radio_安装调试验收).click()
    cc.sap.find_element(locator.saplogon.button_激活状态).click()
    time.sleep(1)
    cc.sap.find_element(locator.saplogon.button_后退).click()

def createOrder(order):
    retry = 0
    while True:
        try:
            flag=cc.wait_appear(locator.sap.logon.start_close,wait_timeout=1)
            if flag!=None:
                flag.click()
                time.sleep(0.5)
            log.logger.info("录入va01")
            va01=cc.wait_appear(locator.sap.order.va01,wait_timeout=1)
            if va01==None:
                cc.find_element(locator.sap.order.exit_1).click()
                time.sleep(0.5)
                cc.find_element(locator.sap.order.exit_1).click()
            input_text(locator.sap.order.va01, "va01")
            cc.send_hotkey("{ENTER}")
            log.logger.info("录入订单类型等数据")
            otype = order["订单类型"]
            #if(str(order["工厂"]).startswith("3")):
            #    otype=  "Z007"
            if cc.sap.find_element(locator.sap.order.order_type).get_text()!='':
                clear_text(locator.sap.order.order_type, {})
                input_text(locator.sap.order.order_type, otype)
            else:
                if order.__contains__('新办事处') and str(order['新办事处']) not in ['', 'nan']:
                    clear_text(locator.sap.order.办事处, {})
                    input_text(locator.sap.order.办事处, str(int(float(order["新办事处"]))))

                clear_text(locator.sap.order.prod, {})
                input_text(locator.sap.order.prod, "10")

                clear_text(locator.sap.order.channel, {})
                input_text(locator.sap.order.channel, "10")

                clear_text(locator.sap.order.sale_org, {})
                input_text(locator.sap.order.sale_org, order["销售组织"])

                clear_text(locator.sap.order.order_type, {})
                input_text(locator.sap.order.order_type, otype)

                clear_text(locator.sap.order.sale_group, {})
                input_text(locator.sap.order.sale_group, "330")
                
                
            cc.send_hotkey("{ENTER}")

            log.logger.info("录入订单详细数据")
            #cc.sap.find_element(locator.sap.logon.create_pane).highlight(duration=1)

            cc.sap.find_element(locator.sap.detail.buyer).click()
            cc.send_hotkey("{F4}") 
            address_in=cc.wait_appear(locator.sap.detail.address_in,wait_timeout=2)

            if address_in !=None:
                address_in.highlight(duration=2)
                address_in.double_click(by="mouse-emulation") 
            time.sleep(0.5)
            #cc.find_element(locator.sap.detail.address_pane).highlight(duration=2)
            cc.find_element(locator.sap.detail.address_pane).click(by="mouse-emulation")
            log.logger.info("查找客户")
            pc.copy(order["客户"])
            time.sleep(1)
            #cc.sap.find_element(locator.sap.detail.buyer_name).click()
            cc.send_hotkey("^v")
            #input_text(locator.sap.detail.buyer_name, order["客户"])
            #cc.sap.find_element(locator.sap.detail.buyer_name_cfm).click()
            cc.send_hotkey('{ENTER}')
            itemBtn = cc.wait_appear(locator.sap.detail.buyer_item_cfm,wait_timeout=3)
            if itemBtn == None:
                cc.find_element(locator.sap.detail.name_in).click()
                mycustom=order["客户"]
                mycustom=mycustom.replace('（','(')
                mycustom=mycustom.replace('）',')')
                pc.copy(mycustom)
                clear_text(locator.sap.detail.name_in,{})
                cc.send_hotkey('^v')
                cc.send_hotkey('{ENTER}')
                itemBtn = cc.wait_appear(locator.sap.detail.buyer_item_cfm,wait_timeout=3)
            if itemBtn == None:
                log.logger.info("客户没找到： "+order["客户"])
                cc.send_hotkey("{ESC}")
                raise RuntimeError("客户没找到")

            #cc.sap.find_element(locator.sap.detail.buyer_item_cfm).click()
            cc.send_hotkey('{ENTER}')
            time.sleep(1)
            cc.sap.find_element(locator.sap.detail.buyer_title).click()
            time.sleep(1)
            text=cc.sap.find_element(locator.sap.items.result_panel).get_text()
            if text.__contains__("信贷限额"):
                cc.send_hotkey("{ENTER}")
                cc.sap.find_element(locator.sap.detail.buyer_title).click()
            #cc.sap.find_element(locator.sap.detail.partner).click(mouse_location=MouseLocation('center',-50, 0,0,0))
            #2024-06-25新需求
            anzhuang_yanshou_set(order)
            
            ui(locator.saplogon.img_partner).click()
            cc.sap.find_element(locator.sap.detail.send_target_id).double_click()
            
            log.logger.info("输入客户地址")
            fill_retry = 0
            while True:
                try:
                    FillConsiAddress(order["收货人"])
                    break
                except Exception as e:
                    fill_retry += 1
                    if fill_retry > 3:
                        raise e
                    time.sleep(1)

            log.logger.info("确认客户地址")
            cc.sap.find_element(locator.sap.detail.sender_cfm).click()
            time.sleep(0.5)
            sender_cfm=cc.wait_appear(locator.sap.detail.sender_cfm,wait_timeout=1)
            if sender_cfm!=None:
                cc.find_element(locator.sap.detail.sender_address,{'idx':1}).double_click(by="mouse-emulation")
                cc.send_hotkey('{ENTER}')
            time.sleep(0.5)
            #cc.sap.find_element(locator.sap.detail.back_btn).click()
            log.logger.info("输入参考号和贵方,编号为" + order["编号"])
            #cc.sap.find_element(locator.sap.detail.order_data).click()
            ui(locator.saplogon.img_order).click()
            pc.copy(order["编号"])
            cc.sap.find_element(locator.sap.detail.order_cust_ref).click()
            time.sleep(0.5)
            cc.send_hotkey("^v")
            time.sleep(0.5)
            
            input_text(locator.sap.detail.the_other_ref, order["合同号"])
            log.logger.info("输入备注2")
            #cc.sap.find_element(locator.sap.detail.to_text).click()
            ui(locator.saplogon.img_text).click()
            time.sleep(0.5)
            cc.sap.find_element(locator.sap.detail.in_text).click()
            pc.copy(order["备注2"])
            cc.send_hotkey("^v")
            time.sleep(0.5)
            log.logger.info("返回详细页输入料号等信息")
            cc.sap.find_element(locator.sap.detail.back_btn).click()

            log.logger.info("输入交期")
            cc.wait_appear(locator.sap.items.shipDate,wait_timeout=3)
            txt = cc.sap.find_element(locator.sap.items.shipDate).get_text()
            cc.sap.find_element(locator.sap.items.shipDate).click()
            while txt != "":
                txt = cc.sap.find_element(locator.sap.items.shipDate).get_text()
                time.sleep(0.1)
                cc.send_hotkey("{DEL}")
                cc.send_hotkey("{BKSP}")
            time.sleep(0.5)
            input_text(locator.sap.items.shipDate, order["交期"])
            cc.send_hotkey("{ENTER}")
            save_date=cc.wait_appear(locator.sap.items.save_date,wait_timeout=4)
            if save_date!=None:
                cc.sap.find_element(locator.sap.items.save_date).click()

            log.logger.info("输入料号")
            FillItems(order, order["items"])
            break
        except Exception as e:
            log.logger.error(traceback.format_exc())
            retry += 1
            if retry > 5:
                raise e
            log.logger.info("retry")
            time.sleep(3)
            logon()

    time.sleep(2)
    cc.sap.find_element(locator.sap.items.save).click()
    time.sleep(1)
    result = cc.sap.find_element(locator.sap.items.result_panel).get_text()
    while result.__contains__("价格没有维护"):
        print('there')
        cc.send_hotkey('{ENTER}')
        time.sleep(0.5)
        result = cc.sap.find_element(locator.sap.items.result_panel).get_text()

    if result.__contains__("已保存") == False:
        time.sleep(0.5)
        cc.send_hotkey("{ENTER}")
        result = cc.sap.find_element(locator.sap.items.result_panel).get_text()

    print(result)
    if result.__contains__("已保存"):
        tryTimes = 5
        while tryTimes > 0:
            try:
                cc.sap.find_element(locator.sap.logon.main_pane).highlight(duration=1, timeout=5)
                break
            except:
                cc.sap.find_element(locator.sap.detail.back_btn).click()
                tryTimes = tryTimes - 1
        if tryTimes < 0:
            cc.send_hotkey("%{F4}")
            cc.sap.find_element(locator.sap.logon.exit_btn).click()
  
    log.logger.info("保存结果："+result)
    return result

def FillConsiAddress(addrStr:str):

    #order["收货人"]
    lenLimit = 35
    lines = []
    while len(addrStr)>lenLimit:
        subStr = addrStr[0:lenLimit-len(addrStr)]
        lines.append(subStr)
        addrStr = addrStr[lenLimit:len(addrStr)]
    
    lines.append(addrStr)
    cc.find_element(locator.sap.detail.factory).highlight(duration=1)
    cc.find_element(locator.sap.detail.factory).click()
    cc.send_hotkey('{ENTER}')
    time.sleep(1)
    if len(lines) > 2:
        try:
            cc.sap.find_element(locator.sap.detail.more_addr_btn1).highlight(duration=1)
            cc.sap.find_element(locator.sap.detail.more_addr_btn1).double_click(by="mouse-emulation")
            size = cc.sap.find_element(locator.sap.detail.sender_address, {"idx":3}).get_size()
            if size.Height == 0 or size.Width == 0:
                raise Exception('error')
        except:
            cc.sap.find_element(locator.sap.detail.more_addr_btn).highlight(duration=1)
            cc.sap.find_element(locator.sap.detail.more_addr_btn).double_click(by="mouse-emulation")
        time.sleep(1)
    idx = 1
    for l in lines:
        laddr = l
        pc.copy(laddr)    
        cc.sap.find_element(locator.sap.detail.sender_address, {"idx":idx}).double_click(by="mouse-emulation")
        txt = cc.sap.find_element(locator.sap.detail.sender_address, {"idx":idx}).get_text()
        while txt != "":
            txt = cc.sap.find_element(locator.sap.detail.sender_address, {"idx":idx}).get_text()
            time.sleep(0.1)
            cc.send_hotkey("{DEL}")
            cc.send_hotkey("{BKSP}")
        time.sleep(0.5)
        cc.send_hotkey("^v")
        time.sleep(0.5)
        cc.sap.find_element(locator.sap.detail.company_post).click()
        idx = idx + 1


def clear_text(loc, var):
    txt = cc.sap.find_element(loc, var).get_text()
    cc.sap.find_element(loc, var).click()
    retry = 0
    while txt != "": 
        txt = cc.sap.find_element(loc, var).get_text()
        time.sleep(0.1)
        cc.send_hotkey("{DEL}")
        cc.send_hotkey("{BKSP}")
        retry += 1
        if retry > 200:
            raise Exception("操作遇到错误,请看录屏")

def input_text_enter(loc,var, text):
    cc.sap.find_element(loc, var).click()
    cc.send_text(text)
    cc.send_hotkey("{ENTER}")

def input_text_simple(loc,var, text):
    cc.sap.find_element(loc, var).click()
    cc.send_text(text)
    time.sleep(0.5)

def FillItems(order, items):
    if len(items) <= 15:
        OldFillItems(order, items)
    else:
        NewFillItems(order, items)

def NewFillItems(order, items):
    
    cur_num=0
    for i in range(len(items)):
        time.sleep(0.5)        
        log.logger.info("输入料号: " + items[i]["料号"])
        cc.wait_appear(locator.sap.items.partNo_input, {"idx":0},wait_timeout=3)
        input_text_simple(locator.sap.items.partNo_input, {"idx":0}, items[i]["料号"])
        input_text_simple(locator.sap.items.qty, {"idx":0}, items[i]["数量"])
        if order['订单类型'] != 'ZKB':
            input_text_simple(locator.sap.items.price, {"idx":0}, items[i]["单价"])
        time.sleep(0.5)
        input_text_simple(locator.sap.items.factory, {"idx":0}, items[i]["工厂"])

        cc.send_hotkey("{ENTER}")
        time.sleep(0.5)
        wuliao_status=cc.wait_appear(locator.sap.items.result_panel,wait_timeout=3)
        wuliao_status=cc.find_element(locator.sap.items.result_panel).get_text()
        log.logger.info(wuliao_status)
        if wuliao_status=="":
            cc.send_hotkey("{ENTER}")#应对物料号被删除的情况
        elif wuliao_status.__contains__("未") or wuliao_status.__contains__("不") or wuliao_status.__contains__("错误"):          
            raise Exception(wuliao_status.replace(' ','').replace(':','').replace('/','').replace('?','').replace('<','').replace('>','').replace('|',''))
        time.sleep(1)
        ui(locator.saplogon.img_翻页).click(mouse_location=MouseLocation(xrate=0.25, yrate=0.25))
        time.sleep(3)
        retry = 0
        while True:
            text = ui(locator.sap.items.partNo_input, {"idx":0}).get_text()
            if text == '':
                break
            else:
                retry += 1
                log.logger.info("wait for first line empty")
                time.sleep(1)
                if retry > 15:
                    raise Exception("录入明细失败")
    
    # update to klv
    time.sleep(1)
    cc.sap.find_element(locator.sap.items.procure).click()
    cc.wait_appear(locator.sap.items.req_type, {"idx":0},wait_timeout=15)
    '''
    注意大条件卡死工厂只有（1078 3510  3520  1073   1074  1075， 1079），其他直接报错
    1、工厂1078 都是KELV（无需判断物料号）  
    2、3510  3520  1073   1074  1075， 1079 工厂 
        2.1 物料号不是1100开头的 无需判断任何都是KSV; 
        2.2 物料号是以1100开头的，备注1有备注就都是KELV，备注1无备注要求KSV
    '''
    for i in range(len(items)):
        if items[i]["备注1"]=="":
            cc.wait_appear(locator.sap.items.req_type, {"idx":0},wait_timeout=3)
            clear_text(locator.sap.items.req_type, {"idx":0})
            cc.sap.find_element(locator.sap.items.req_type, {"idx":0}).click()
            if items[i]["工厂"] == '1078':
                cc.send_text("KELV")
            else:
                cc.send_text("KSV")
            time.sleep(0.1)
            cc.send_hotkey("{ENTER}")
            time.sleep(0.5)
        else:
            cc.find_element(locator.sap.items.partNo_choose, {"idx":0}).highlight(duration=2)
            cc.find_element(locator.sap.items.partNo_choose, {"idx":0}).double_click(by='mouse-emulation')
            time.sleep(0.5)
            pc.copy(items[i]["备注1"])
            to_text=cc.wait_appear(locator.sap.items.to_text,wait_timeout=1)
            if to_text!=None:
                to_text.click()
            else:
                cc.find_element(locator.sap.items.to_text_1).click()
            time.sleep(0.5)
            cc.find_element(locator.sap.items.text_in).click() 
            cc.send_hotkey('^v')
            time.sleep(0.5)
            cc.find_element(locator.sap.items.toBack).click()
            time.sleep(0.5)
            clear_text(locator.sap.items.req_type, {"idx":0})
            cc.sap.find_element(locator.sap.items.req_type, {"idx":0}).click()
            if items[i]["工厂"] == '1078':
                cc.send_text("KELV")
            else:
                if items[i]['料号'].startswith('1100'):
                    cc.send_text("KELV")
                else:
                    cc.send_text("KSV")
            time.sleep(0.1)
            cc.send_hotkey("{ENTER}")
            time.sleep(0.5)
        ui(locator.saplogon.img_翻页1).click(mouse_location=MouseLocation(xrate=0.25, yrate=0.25))
        time.sleep(5)

def OldFillItems(order, items):
    cur_num=0
    for i in range(len(items)):
        time.sleep(0.5)        
        log.logger.info("输入料号: " + items[i]["料号"])
        cc.wait_appear(locator.sap.items.partNo_input, {"idx":i},wait_timeout=3)
        input_text_simple(locator.sap.items.partNo_input, {"idx":i}, items[i]["料号"])
        input_text_simple(locator.sap.items.qty, {"idx":i}, items[i]["数量"])
        if order['订单类型'] != 'ZKB':
            input_text_simple(locator.sap.items.price, {"idx":i}, items[i]["单价"])
        time.sleep(0.5)
        input_text_simple(locator.sap.items.factory, {"idx":i}, items[i]["工厂"])
        if i <8:
            cc.send_hotkey("{ENTER}")
            time.sleep(0.5)
            wuliao_status=cc.wait_appear(locator.sap.items.result_panel,wait_timeout=3)
            wuliao_status=cc.find_element(locator.sap.items.result_panel).get_text()
            log.logger.info(wuliao_status)
            if wuliao_status=="":
                cc.send_hotkey("{ENTER}")#应对物料号被删除的情况
            elif wuliao_status.__contains__("未") or wuliao_status.__contains__("不") or wuliao_status.__contains__("错误"):          
                raise Exception(wuliao_status.replace(' ','').replace(':','').replace('/','').replace('?','').replace('<','').replace('>','').replace('|',''))
        time.sleep(0.5)

    # update to klv
    time.sleep(1)
    cc.sap.find_element(locator.sap.items.procure).click()
    for i in range(len(items)):
        #if items[i]["备注1"]=="" or (not items[i]["料号"].startswith("1100")):
        if items[i]["备注1"]=="":
            cc.wait_appear(locator.sap.items.req_type, {"idx":i},wait_timeout=3)
            clear_text(locator.sap.items.req_type, {"idx":i})
            cc.sap.find_element(locator.sap.items.req_type, {"idx":i}).click()
            if items[i]["工厂"] == '1078':
                cc.send_text("KELV")
            else:
                cc.send_text("KSV")
            time.sleep(0.1)
            cc.send_hotkey("{ENTER}")
            time.sleep(0.5)
        else:

            cc.find_element(locator.sap.items.partNo_choose, {"idx":i}).highlight(duration=2)
            cc.find_element(locator.sap.items.partNo_choose, {"idx":i}).double_click(by='mouse-emulation')
            time.sleep(0.5)
            pc.copy(items[i]["备注1"])
            to_text=cc.wait_appear(locator.sap.items.to_text,wait_timeout=1)
            if to_text!=None:
                to_text.click()
            else:
                cc.find_element(locator.sap.items.to_text_1).click()
            time.sleep(0.5)
            cc.find_element(locator.sap.items.text_in).click() 
            cc.send_hotkey('^v')
            time.sleep(0.5)
            cc.find_element(locator.sap.items.toBack).click()
            time.sleep(0.5)
            clear_text(locator.sap.items.req_type, {"idx":i})
            cc.sap.find_element(locator.sap.items.req_type, {"idx":i}).click()
            if items[i]["工厂"] == '1078':
                cc.send_text("KELV")
            else:
                if items[i]['料号'].startswith('1100'):
                    cc.send_text("KELV")
                else:
                    cc.send_text("KSV")
            time.sleep(0.1)
            cc.send_hotkey("{ENTER}")
            time.sleep(0.5)

def vl10b_销售凭证(order_no):
    log.logger.info("VL10B 根据采购编号抓取销售凭证: " + order_no)
    logon()
    retry = 0
    while True:
        try:
            flag=cc.wait_appear(locator.sap.logon.start_close,wait_timeout=1)
            if flag!=None:
                flag.click()

            cc.sap.find_element(locator.sap.order.va01).call_transaction("VL10B")
            time.sleep(5)
            ui(locator.sap.vl10b1.purchase_order).click()
            time.sleep(1)
            clear_text(locator.sap.vl10b1.purchase_voucher,{})
            input_text(locator.sap.vl10b1.purchase_voucher, order_no)
            
            cc.sap.find_element(locator.saplogon.purchase_search).click()
            time.sleep(1)
            cc.sap.find_element(locator.sap.vl10b1.purchase_selectall).click()
            time.sleep(1)
            cc.sap.find_element(locator.sap.vl10b1.btn_backgroud).click()
            time.sleep(1)
            cc.sap.find_element(locator.sap.vl10b1.btn_display).click()
            time.sleep(3)
            #抓取销售凭证号
            voucher = cc.sap.find_element(locator.sap.vl10b1.purchase_voucher_text).get_text()
            log.logger.info(f"{order_no}销售凭证为: {voucher}")

            return voucher
        except Exception as e:
            log.logger.error(traceback.format_exc())
            retry += 1
            if retry > 2:
                raise e
            log.logger.info("retry")
            time.sleep(3)
            logon()

def createME21N(item):
    log.logger.info("录入ME21N: " + item["编号"])
    result = ""
    retry = 0
    while True:
        try:
            flag=cc.wait_appear(locator.sap.logon.start_close,wait_timeout=1)
            if flag!=None:
                flag.click()

            cc.sap.find_element(locator.sap.order.va01).call_transaction("ME21N")
            time.sleep(0.5)
            flag=cc.wait_appear(locator.sap.me21n.sure_open,wait_timeout=1)
            if flag!=None:
                flag.click()
            try:
                time.sleep(0.5)
                cc.sap.find_element(locator.sap.me21n.order_kind_choose).click()
            except:
                cc.sap.find_element(locator.sap.me21n.fold_out).click()
                time.sleep(0.5)
                cc.sap.find_element(locator.sap.me21n.order_kind_choose).click()
            time.sleep(1)
            # 办事处转储采购 ：1072、1073、1079
            if str(item['订单类型'])=="ZUB":
                log.logger.info(f"工厂：{item['items'][0]['工厂']}，订单类型为：ZUB-办事处转储采购")
                cc.send_hotkey('{PGUP}')
                cc.send_hotkey('{PGUP}')
                cc.send_hotkey('{PGUP}')
            # 跨公司转储：3510、3520、1100
            elif str(item['订单类型'])=="ZNB":
                log.logger.info(f"工厂：{item['items'][0]['工厂']}，订单类型为：ZNB-跨公司转储")

                for _ in range(4):
                    cc.send_hotkey('{DOWN}')
                    time.sleep(0.5)
                '''
                cc.find_element(locator.sap.me21n.factory_3).highlight(duration=2)
                cc.find_element(locator.sap.me21n.factory_3).click()
                '''
            time.sleep(1)
            cc.send_hotkey('{ENTER}')
            clear_text(locator.sap.me21n.factory_input,{})
            input_text(locator.sap.me21n.factory_input, item['items'][0]['工厂'])   #供货工厂
            cc.send_hotkey("{ENTER}")
            log.logger.info("根据表格填写采购信息")
            fillCaigou(item["items"][0]['工厂'])
            cc.find_element(locator.sap.me21n.tongxin).click()
            log.logger.info(item['items'][0]["委托编号"])
            编号=item['items'][0]["委托编号"]
            pc.copy(编号)
            cc.sap.find_element(locator.sap.me21n.other_refer).click()
            cc.send_hotkey("^V")

            log.logger.info("文本页")
            cc.sap.find_element(locator.sap.me21n.text_page).click()
            if str(item["收货人"]) not in ['', 'nan']:
                pc.copy(item["收货人"])
                cc.sap.find_element(locator.sap.me21n.address_in).click()
                cc.send_hotkey("^V")
            time.sleep(1)

            cc.sap.find_element(locator.sap.me21n.detail.wuliao).click()

            result = fillDetail(item['items'],item['地点'],str(item['订单类型']), item['单据类型'])
            
            cc.sap.find_element(locator.sap.items.save).click()
            time.sleep(1)
            try:
                error_existing = cc.is_existing(locator.saplogon.button_error_cancel,timeout=5)
            except:
                error_existing = False
            if error_existing:
                cc.sap.find_element(locator.saplogon.button_error_cancel).click()
                try:
                    cc.sap.find_element(locator.saplogon.button_back).click()
                    cc.sap.find_element(locator.saplogon.button_not_save).click()
                except:
                    log.logger.info("退出失败，ignore")
                time.sleep(1)
                raise Exception(f"保存凭证出错,取消保存: {result}")
            break
        except Exception as e:
            err_str = str(e)

            if "请输入" in err_str:
                log.logger.warning(f"文本输入异常，尝试重试...")

            # 判断是否为业务异常，不重试
            if "未被维护" in err_str or "物料" in err_str or "工厂"in err_str:
                log.logger.error(f"业务异常：{err_str}")
                return err_str  # 直接返回，不 raise

            retry += 1
            log.logger.error(traceback.format_exc())
            if retry > 2:
                log.logger.error("超过最大重试次数，终止录单")
                raise e
            log.logger.info("retry")
            time.sleep(3)
            logon()
    

    #result = cc.sap.find_element(locator.sap.items.result_panel).get_text()
    time.sleep(1)
    status = cc.sap.find_element(locator.sap.items.result_panel).get_statusbar()
    result = status.Text[0]
    log.logger.info("录单结果" + result)

    while result.__contains__("A版价格"):
        cc.send_hotkey('{ENTER}')
        time.sleep(0.5)
        result = cc.sap.find_element(locator.sap.items.result_panel).get_text()

    if result.__contains__("已创建") == False:
        time.sleep(0.5)
        cc.send_hotkey("{ENTER}")
        result = cc.sap.find_element(locator.sap.items.result_panel).get_text()
        return result

    cc.find_element(locator.sap.me21n.me21n_close).click()
    log.logger.info("保存结果："+result)
    return result


def click_if_exist(elem_locator, timeout=5):
    """
    若控件存在则点击，不存在不报错
    """
    elem = cc.wait_appear(elem_locator, wait_timeout=timeout)
    if elem:
        try:
            elem.click()
            time.sleep(0.8)
            return True
        except:
            return None


def check_11_hang(all_items):
    """
    根据 all_items，依次点击物料号以 11 开头的行
    - 前7行：locator.sap.MIGO.行选择区
    - 7行以后：先点折叠详细信息，再用 locator.sap.MIGO.行选择区大于7
    """
    for i, item in enumerate(all_items):
        try:
            mat = str(item.get("料号") or "").strip()
            if not mat.startswith("11"):
                log.logger.info(f"第 {i+1} 行物料: {mat} 非11开头, 跳过")
                continue
            # 分组选择正确的 locator
            if i < 7:
                row_locator = locator.sap.MIGO.行选择区
                row_params = {"idx": i}
            else:
                click_if_exist(locator.sap.MIGO.折叠详细信息, timeout=3)
                time.sleep(1.5)
                row_locator = locator.sap.MIGO.行选择区大于7
                row_params = {"idx": i}
            row = cc.sap.find_element(row_locator, row_params)
            row.click(by="mouse-emulation")
            log.logger.info(f"第 {i+1} 行物料: {mat} 点击成功)")
            time.sleep(1.2)
            # ======== 查找序列号处理 ========
            ui(locator.sap.MIGO.img_序列号).click()
            time.sleep(1.2)

            result = click_if_exist(locator.sap.MIGO.查找序列号)
            if result is None:
                log.logger.info(f"未发现'查找序列号'控件，跳过当前物料行：{mat}")
                continue

            # ======== 序列号填写 ========
            serial_list = item.get("序列号") or []
            if not serial_list:
                log.logger.warning(f"物料 {mat} 未提供序列号列表")
                continue
            for i, sn in enumerate(serial_list):
                input_text_simple(locator.sap.MIGO.序列号填写行,{'idx':i},sn)
                time.sleep(0.3)
                log.logger.info(f"第 {i+1} 行物料 {mat}：填写序列号 {i+1}/{len(serial_list)} = {sn}")

        except Exception as e:
            log.logger.error(f"第 {i+1} 行物料11开头点击失败: {e}")
            raise


def migo_收货(so, all_items):
    retry = 0
    while True:
        try:
            # =========== 初始化 ==========
            flag = cc.wait_appear(locator.sap.logon.start_close, wait_timeout=1)
            if flag:
                flag.click()
                time.sleep(0.5)

            log.logger.info("录入va01")
            va01 = cc.wait_appear(locator.sap.order.va01, wait_timeout=5)
            if va01 is None:
                raise RuntimeError("未找到事务码输入框控件")
            cc.find_element(locator.sap.order.va01).click()
            safe_input(locator.sap.order.va01, "MIGO")
            cc.send_hotkey("{ENTER}")
            time.sleep(0.8)

            # =========== 填写外向交货单号 ===========
            click_if_exist(locator.sap.MIGO.折叠详细信息)
            safe_input(locator.sap.MIGO.销售凭证单号,so)
            safe_click(locator.sap.MIGO.执行)
            time.sleep(0.8)

            # =========== 勾选全选 ===========
            item_count = len(all_items)
            log.logger.info(f"[MIGO] 本次预计勾选行数={item_count}, so={so}")
            check_all_items(item_count)

            # =========== 依次点击item_count中11开头的物料号 ===========
            cc.wait_appear(locator.sap.MIGO.详细数据,wait_timeout=5).click()
            time.sleep(0.8)
            check_11_hang(all_items)

            # =========== 点击销售过账 ===========
            safe_click(locator.sap.MIGO.过账)
            time.sleep(1.2)
            result = cc.wait_appear(locator.sap.MIGO.过账文本框, wait_timeout=10).get_text()

            if "已过账" not in result:
                cc.send_hotkey("{ENTER}")
                time.sleep(0.5)
                result = cc.sap.find_element(locator.sap.items.result_panel).get_text()
                
            log.logger.info("保存结果：" + result)
            return result

        except Exception as e:
            log.logger.error(traceback.format_exc())
            retry += 1
            if retry > 5:
                log.logger.error(f"录单失败超过重试次数 ({retry})，销售凭证：{so}")
                raise e
            log.logger.warning(f"录单异常，第 {retry} 次重试中...")
            time.sleep(3)
            logon()



def vl02n_收货(so, no, all_items):
    retry = 0
    while True:
        try:
            # =========== 初始化 ==========
            flag = cc.wait_appear(locator.sap.logon.start_close, wait_timeout=1)
            if flag:
                flag.click()
                time.sleep(0.5)

            log.logger.info("录入vl02n")
            va01 = cc.wait_appear(locator.sap.order.va01, wait_timeout=5)
            if va01 is None:
                raise RuntimeError("未找到事务码输入框控件")
            cc.find_element(locator.sap.order.va01).click()
            safe_input(locator.sap.order.va01, "VL02N")
            cc.send_hotkey("{ENTER}")
            time.sleep(0.8)

            safe_input(locator.sap.VL02N.外向交货,so)
            cc.send_hotkey("{ENTER}")
            ui(locator.sap.VL02N.img_拣配).click()
            time.sleep(0.8)

            # 填写存储5101,以及序列号
            for i in range(all_items):
                input_text_simple(locator.sap.VL02N.存储,{'idx':i},"5101")
                time.sleep(0.3)
                cc.sap.find_element(locator.sap.VL02N.存储,{'idx':i}).click()
                cc.send_hotkey("%a") 
                time.sleep(0.3)
                cc.send_hotkey("^s") 
                safe_click(locator.sap.VL02N.查找序列号)
                # 填写物料凭证号
                safe_input(locator.sap.VL02N.物料凭证,no)
                safe_click(locator.sap.VL02N.确定2)
                safe_click(locator.sap.VL02N.确定3)
                safe_click(locator.sap.VL02N.确定4)

            # =========== 点击销售过账 ===========
            safe_click(locator.sap.VL02N.过账发货)

            time.sleep(1.2)
            result = cc.wait_appear(locator.sap.VL02N.过账发货文本框, wait_timeout=10).get_text()

            if "已过账" not in result:
                cc.send_hotkey("{ENTER}")
                time.sleep(0.5)
                result = cc.sap.find_element(locator.sap.items.result_panel).get_text()
                
            log.logger.info("保存结果：" + result)
            return result

        except Exception as e:
            log.logger.error(traceback.format_exc())
            retry += 1
            if retry > 5:
                log.logger.error(f"录单失败超过重试次数 ({retry})，销售凭证：{so}")
                raise e
            log.logger.warning(f"录单异常，第 {retry} 次重试中...")
            time.sleep(3)
            logon()



def fillDetail(items,addr_num,kind, 单据类型):
    table_len=items.__len__()#同订单的物料数决定了表要填几行
    factory_index=11
    address_index=12
    if kind=="ZNB":
        factory_index=15
        address_index=16
    flag=0
    result = ""
    for i in range(table_len):
        try:
            if cc.sap.find_element(locator.sap.me21n.detail.detail_true):#如果细节展开则收回
                time.sleep(0.5)
                cc.sap.find_element(locator.sap.me21n.detail.detail_true).click()           
        finally:
            #填物料
            time.sleep(2)
            cc.sap.find_element(locator.sap.me21n.detail.detail_shorttext).click(by="mouse-emulation")
            time.sleep(1)
            cc.send_hotkey("+{TAB}")
            cc.sap.find_element(locator.sap.me21n.detail.wuliao).highlight(duration=1)
            cc.sap.find_element(locator.sap.me21n.detail.wuliao).click(by="mouse-emulation")
            liaohao = items[i]['料号']
            pc.copy(items[i]['料号'])
            log.logger.info(f'填物料:{liaohao}')

            time.sleep(1)
            fillPGDN(i)

            cc.send_hotkey('^V')

            cc.find_element(locator.sap.detail.number_change,{'idx':i}).click()
            fold_close=cc.wait_appear(locator.sap.me21n.fold_close,wait_timeout=3)
            if fold_close!=None:
                fold_close.click()
            #clear_text(locator.sap.detail.number_change,{'idx':i})
            input_text_simple(locator.sap.detail.number_change,{'idx':i},items[i]['数量'])  #采购订单数量
            #cc.send_hotkey('{ENTER}')
            time.sleep(0.5)
            log.logger.info('填工厂')
            cc.find_element(locator.sap.detail.factory_change,{'idy':factory_index,'idx':i}).highlight(duration=1)
            cc.find_element(locator.sap.detail.factory_change,{'idy':factory_index,'idx':i}).click(by="mouse-emulation")
            time.sleep(1)
            input_text_simple(locator.sap.detail.factory_change,{'idy':factory_index,'idx':i},'1072')
            time.sleep(1)

            '''
            if kind=="ZNB" and flag==0:
                cc.send_hotkey("{ENTER}")
                flag=1
            '''

            cc.find_element(locator.sap.detail.factory_change,{'idy':factory_index,'idx':i}).click(by="mouse-emulation")
            time.sleep(1)
            cc.send_hotkey("{TAB}")
            addr_locator = locator.sap.detail.address_change
            try:
                size = cc.find_element(locator.sap.detail.address_change,{'idy':address_index,'idx':i}).get_size(timeout=5)
                if size.Height == 0 or size.Width == 0:
                    addr_locator = locator.sap.detail.address_change1
            except:
                addr_locator = locator.sap.detail.address_change1
           
            cc.find_element(addr_locator,{'idy':address_index,'idx':i}).highlight(duration=1,timeout=5)
            cc.find_element(addr_locator,{'idy':address_index,'idx':i}).click(by="mouse-emulation")
            time.sleep(1)
            input_text_simple(addr_locator,{'idy':address_index,'idx':i},addr_num)   #库存地点

            if 单据类型 == "退货单":
                for _ in range(7):
                    cc.send_hotkey("{TAB}")
                    time.sleep(1)
                cc.find_element(locator.sap.detail.return_check,{'idx':i}).click(by="mouse-emulation")  #勾选退货
                time.sleep(1)
                '''
                for _ in range(13):
                    cc.send_hotkey("+{TAB}")
                    time.sleep(1)
                '''

            cc.send_hotkey('{ENTER}')
            time.sleep(8)
            try:
                status = cc.sap.find_element(locator.sap.items.result_panel).get_statusbar()
                level = status.Type[0]
                result1 = status.Text[0]
                if result1 != "" and level == "Error":
                    result = result + ", " + result1
            except:
                pass
            wuliao_status=cc.find_element(locator.sap.items.result_panel).get_text()
            if wuliao_status.__contains__("未被维护"):
                raise Exception(wuliao_status)
            time.sleep(0.5)
            if cc.is_existing(locator.saplogon.project_detail):
                ui(locator.saplogon.project_detail).click()
                time.sleep(1)
            pos = cc.find_element(locator.saplogon.img_左右).get_position()
            x = int(pos.Left + (pos.Right - pos.Left)/4)
            y = int(pos.Top + (pos.Bottom - pos.Top)/2)
            for j in range(3):
                cc.mouse.click(x, y)
                time.sleep(1)
            if items[i]['备注1'] not in ['', 'nan']:
                if i > 0:
                    cc.sap.find_element(locator.saplogon.wuliao_wenben_next).click()
                    time.sleep(1)
                pc.copy(items[i]['备注1'])
                ui(locator.saplogon.wuliao_wenben).click()
                time.sleep(5)
                cc.sap.find_element(locator.saplogon.wuliao_wenben_text).click()
                time.sleep(1)
                cc.send_hotkey("^V")
                time.sleep(1)
            try:
                status = cc.sap.find_element(locator.sap.items.result_panel).get_statusbar()
                level = status.Type[0]
                result1 = status.Text[0]
                if result1 != "" and level == "Error":
                    result = result + ", " + result1
            except:
                pass
            cc.sap.find_element(locator.saplogon.wuliao_expand_btn).click()
            time.sleep(1)

    return result

def fillPGDN(cnt):
    for i in range(cnt):
        time.sleep(0.5)
        cc.send_hotkey('{DOWN}')
        time.sleep(0.5)

def fillCaigou(number):
    cai_1="1002"  #采购组织
    cai_2="W01"   #采购组
    cai_3='1000'  #公司代码
    input_text(locator.sap.me21n.cai_1, cai_1)
    input_text(locator.sap.me21n.cai_2, cai_2)
    input_text(locator.sap.me21n.cai_3, cai_3)
    time.sleep(0.5)


def safe_input(elem_locator, text, timeout=3, retry=3):
    """
    安全输入文本
    """
    for i in range(retry):
        elem = cc.wait_appear(elem_locator, wait_timeout=timeout)
        if not elem:
            log.logger.info(f"[safe_input] 查找控件超时：{elem_locator}, 正在重试 {i+1}/{retry}")
            time.sleep(1)
            continue
        try:
            elem.click()
            elem.send_hotkey("^a")
            elem.send_hotkey("{DEL}")
            pc.copy(text)
            elem.send_hotkey("^v")
            time.sleep(0.2)
            # 校验
            if (elem.get_text() or "").strip() == str(text).strip():
                return True
            log.logger.info(f"[safe_input] 校验失败 重试 {i+1}/{retry}")
            elem.click()
        except Exception as e:
            log.logger.debug(f"[safe_input] 内部异常：{e}")
        time.sleep(1)
    raise Exception(f"safe_input 输入失败：{elem_locator}")

def safe_click(elem_locator, timeout=5, retry=3, sleep=2):
    """
    通用等待点击：
    """
    for attempt in range(1, retry + 1):
        elem = cc.wait_appear(elem_locator, wait_timeout=timeout)
        if elem:
            elem.click()
            time.sleep(1)
            return True
        log.logger.warning(f"[safe_click] 第 {attempt}/{retry} 次等待失败：未找到元素 {elem_locator}")
        time.sleep(sleep)
    raise Exception(f"[safe_click] 点击失败：无法找到元素 {elem_locator}，累计尝试 {retry} 次")


def check_all_items(item_count):
    """
    根据物料条数，依次勾选 SAP 出库单中的 checkbox
    """
    for i in range(item_count):
        try:
            cb = cc.sap.find_element(
                locator.sap.MIGO.勾选,  
                {"idx": i}                      
            )
            cb.click(by="mouse-emulation")
            time.sleep(0.2)
        except Exception as e:
            log.logger.error(f"第 {i+1} 行勾选失败: {e}")
            raise


if __name__ == "__main__":
    order_no = ""
    vorchure = vl10b_销售凭证(order_no)
    print(vorchure)

    