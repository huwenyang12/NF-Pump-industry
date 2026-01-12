import os
import subprocess
import traceback
import util
import log
import app_ludan,shouhou_ludan
import record, time
from datetime import datetime, timedelta
import dingtalk_utils
from clicknium import clicknium as cc
import sys

def ludan(site_config):
    if not site_config.__contains__("录单"):
        return False
    root_folder = site_config["录单"]["root路径"]
    video_folder = site_config["录单"]["录屏保存路径"]
    files = app_ludan.get_handle_files(root_folder)
    if len(files) == 0:
        return False

    try:
        # 企业微信通知
        dingtalk_utils.send_message(f"{site_config['name']}: 开始录单，文件为{files[0]}")
        r = record.start_recorder(video_folder, files[0])
        if "出库单" in files[0]:
            app_ludan.handle(site_config, files[0])
        elif "发货单" in files[0] or "退货单" in files[0]:
            shouhou_ludan.handle(site_config, files[0])

    except Exception as e:
        log.logger.error("异常 {}".format(traceback.format_exc()))
        raise e
    finally:
        record.stop_record(r)
    return True


def jiaohuodan(site_config):
    if not site_config.__contains__("录单"):
        return False
    if "订单记录路径" not in site_config["录单"]:
        return False
    try:
        return shouhou_ludan.handle_jiaohuodan(site_config)
    except Exception as e:
        log.logger.error("异常 {}".format(traceback.format_exc()))
        raise e


def migo(site_config):
    if not site_config.__contains__("录单"):
        return False
    if "订单记录路径" not in site_config["录单"]:
        return False
    try:
        return shouhou_ludan.handle_migo(site_config)
    except Exception:
        log.logger.error("异常 {}".format(traceback.format_exc()))
        raise

def vl02n(site_config):
    if not site_config.__contains__("录单"):
        return False
    if "订单记录路径" not in site_config["录单"]:
        return False
    try:
        return shouhou_ludan.handle_vl02n(site_config)
    except Exception:
        log.logger.error("异常 {}".format(traceback.format_exc()))
        raise



def start_mail_monitor():
    """启动邮件监听子进程，优化管道处理"""
    mail_script = os.path.join(os.path.dirname(__file__), 'mail_monitor.py')

    
    # 重定向输出到文件，避免管道阻塞
    log_file = os.path.join(os.getcwd(), 'logs', 'mail_monitor.log')
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    
    try:
        # 使用 DEVNULL 或重定向到日志文件
        with open(log_file, 'a', encoding='utf-8') as f:
            process = subprocess.Popen(
                [sys.executable, mail_script],
                stdout=f,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL,
                cwd=os.path.dirname(__file__)
            )
        log.logger.info(f"邮件监听子进程已启动，PID: {process.pid}")
        return process
    except Exception as e:
        log.logger.error(f"启动邮件监听失败: {e}")
        return None

def check_mail_process(process):
    """检查邮件监听进程状态"""
    if process is None:
        return False
    
    poll_result = process.poll()
    if poll_result is not None:
        log.logger.warning(f"邮件监听进程已退出，返回码: {poll_result}")
        return False
    return True

if __name__ == "__main__":
    
    cc.config.disable_telemetry()
    util.change_language()
    site_config = util.get_site_config()

    # 启动邮件监听子进程
    mail_process = start_mail_monitor()
    
    time.sleep(2)
    
    try:
        while True:
            try:
                ret = ludan(site_config)
                jiaohuodan(site_config)
                migo(site_config)
                vl02n(site_config)
                if not ret:
                    if datetime.now().hour > 22:
                        util.remove_archive_video_files(site_config)
                        log.logger.info("晚上10点退出")
                        break
                    log.logger.info("等待30s")
                    time.sleep(30)
            except Exception as e:
                log.logger.error(f"主循环异常: {str(e)}")
                time.sleep(30)  # 异常后等待一段时间再继续
                
    except KeyboardInterrupt:
        log.logger.info("收到中断信号，正在退出...")
    finally:
        # 清理子进程
        pass
