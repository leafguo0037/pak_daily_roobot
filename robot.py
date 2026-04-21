import getdata as gd
import requests
import json
import os
import traceback
import pandas as pd
os.chdir('/opt/workspace/pak_risk_group_drive/robot')
import logging  # 导入logging模块，用于记录日志

# 配置日志记录
logging.basicConfig(
    filename='script.log',  # 日志文件名
    level=logging.INFO,  # 日志级别
    format='%(asctime)s - %(levelname)s - %(message)s'  # 日志格式
)


# 将数字转换为百分比或保留原样
def to_persent(x):
    if pd.isna(x):
        return 'null'
    return str(round(x*100,2))+'%'


def build_markdown(js, date):
    res = '# 巴基斯坦业务每日情况\n日期：%s  单位：USD' % date
    parts = [
        '\n',
        '\n\n## 大盘',
        '\n总放款金额(排除冷静期)：%.1f，占月度目标：%s，本月累计占月度目标：%s'% (js.get('日总放款金额(排除冷静期)'),to_persent(js.get('日总放款金额占月度目标')),to_persent(js.get('日总放款金额本月累计占月度目标'))),
        '\n新客戳额人数：%d' % js.get('自然日新客戳额人数'),
        '\n新客放款订单数：%d，占月度目标：%s，本月累计占月度目标：%s'% (js.get('新客放款订单数'),to_persent(js.get('新客放款订单占月度目标')),to_persent(js.get('新客放款订单累计占月度目标'))),
        '\n新客放款金额：%.1f' % js.get('新客放款金额'),
        '\n新客结清订单数：%.1f' % js.get('新客结清订单数'),
        '\n老客戳额人数：%d' % js.get('自然日老客戳额人数'),
        '\n老客放款订单数：%d' % js.get('老客放款订单数'),
        '\n老客放款金额：%.1f' % js.get('老客放款金额'),
        '\n老客结清订单数：%.1f' % js.get('老客结清订单数'),
        '\n整体静态注册成交率：%s，占月度目标%s'% (to_persent(js.get('新客静态注册成交率')),to_persent(js.get('新客静态注册成交率和目标比值'))),
        '\n',

        '\n\n## 产品运营',
        '\n端内T3新客注册申请率：%s，占月度目标%s'% (to_persent(js.get('端内T3新客注册申请率')),to_persent(js.get('T3新客注册申请率和目标比值'))),
        '\n复贷申请率T0：%s，占月度目标%s' % (to_persent(js.get('复贷申请率T0')),to_persent(js.get('复贷申请率T0和目标比值'))),
        '\n新客有额提现率t0：%s，占月度目标%s' % (to_persent(js.get('新客有额提现率t0')),to_persent(js.get('新客有额提现率t0和目标比值'))),
        '\n老客有额提现率t0：%s，占月度目标%s' % (to_persent(js.get('老客有额提现率t0')),to_persent(js.get('老客有额提现率t0和目标比值'))),
        # '\n新客有额提现率t7：%s，占月度目标%s'% (to_persent(js.get('新客有额提现率t7')),to_persent(js.get('新客有额提现率t7和目标比值'))),
        # '\n老客有额提现率t7：%s，占月度目标%s'% (to_persent(js.get('老客有额提现率t7')),to_persent(js.get('老客有额提现率t7和目标比值'))),
        '\n注册戳额率t0：%s，占月度目标：%s'% (to_persent(js.get('注册戳额率t0')),to_persent(js.get('注册戳额率t0和目标比值'))),
        # '\n注册戳额率t7：%s，占月度目标：%s'% (to_persent(js.get('注册戳额率t7')),to_persent(js.get('注册戳额率t7和目标比值'))),
        '\n',

        '\n\n## 投放(新客)',
        '\n首次戳额有额率：%s' % to_persent(js.get('新客通过率(首戳)')),
        '\n新客静态CPS(排除冷静期)：%.2f，占月度目标%s'% (js.get('新客静态CPS(排除冷静期)'),to_persent(js.get('新客静态CPS和目标比值'))),
        '\nt0cps：%.2f'% (js.get('t0cps')),
        '\nt3cps：%.2f，占月度目标%s'% (js.get('t3cps'),to_persent(js.get('t3cps和目标比值'))),
        '\n',

        '\n\n## 风控',
        '\n自然日新客戳额人数：%d, 自然日新客戳额通过率:%s'% (js.get('自然日新客戳额人数'),to_persent(js.get('自然日新客戳额通过率'))),
        '\n自然日老客戳额人数：%d, 自然日老客戳额通过率:%s, 新转老通过率：%s'% (js.get('自然日老客戳额人数'),to_persent(js.get('自然日老客戳额通过率')),to_persent(js.get('自然日新转老戳额通过率'))),
        '\n老客循环贷发标通过率：%s, 放款订单数:%d, 放款金额%.1f'% (to_persent(js.get('老客循环贷发标通过率')),js.get('老客循环贷放款订单数'),js.get('老客循环贷放款金额')),
        '\n新客平均期限：%.1f，占月度目标%s，老客平均期限%.1f，占月度目标%s'% (js.get('新客平均期限'),to_persent(js.get('新客平均期限和目标比值')),js.get('老客平均期限'),to_persent(js.get('老客平均期限和目标比值'))),
        '\n新客件均：%.1f，占月度目标%s'% (js.get('新客件均'),to_persent(js.get('新客件均和目标比值'))),
        '\n老客件均（笔均）：%.1f，占月度目标%s'% (js.get('老客件均'),to_persent(js.get('老客件均和目标比值'))),
        '\n老客循环贷件均：%.1f，老客非循环贷件均（人均）：%.1f'% (js.get('老客循环贷件均'),js.get('老客非循环贷件均')),
        '\n新客$fpd1/fpd7 ：%s(%s)'% (to_persent(js.get('新客$fpd1')),to_persent(js.get('新客$fpd7'))),
        '\n老客$fpd1/fpd7 ：%s(%s)'% (to_persent(js.get('老客$fpd1')),to_persent(js.get('老客$fpd7'))),
        '\n',

        '\n\n## BNPL',
        '\n日申请人数：%d, 日申请通过率: %s, 本月申请总人数：%d, '%(js.get('bnpl申请人数'),to_persent(js.get('bnpl申请通过率')),js.get('本月bnpl申请人数')),
        '\n成交笔数：%d, 成交金额: %.1f, 本月成交总笔数: %d'%(js.get('bnpl成交笔数'),js.get('bnpl成交金额'),js.get('本月bnpl成交笔数')),
        '\n平均期限：%.1f, 平均首付比例：%s'% (js.get('bnpl平均期限'),to_persent(js.get('bnpl平均首付比例'))),
        '\nbnpl fpd1到期数量：%s $fpd1 ：%s'% (js.get('bnplfpd1到期数量'),to_persent(js.get('bnpl$fpd1'))),
        '\nbnpl fpd7到期数量：%s $fpd7 ：%s'% (js.get('bnplfpd7到期数量'),to_persent(js.get('bnpl$fpd7'))),
    ]

    res += ''.join(parts)

    return res



# 用于发送Webhook消息
def send_webhook(url,input_json=None):
    if input_json is None:
        input_json = {"msgtype": "text",
            "text": {"content": "no input"}}
    response = requests.post(url, data=json.dumps(input_json))
    logging.info(f'Webhook消息发送结果: {response.text}')  # 记录日志
    return response



if __name__ == "__main__":
    webhook = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=191d1cdf-833a-4410-8b65-138439fd066e'
    # webhook = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=471d9a53-ceca-4aa4-84c2-a75523e5309f'


    try:
        info_o = gd.Get_info_json()
        big_panel_json = info_o.fill_big_panel_json()
        print(big_panel_json)
        final_txt = build_markdown(big_panel_json,info_o.yesterday_date_str)

        input_json = {"msgtype": "markdown",
                    "markdown": {"content": final_txt}}
        response = send_webhook(webhook,input_json)
        logging.info(f'主程序执行完毕，响应状态码: {response.status_code}')  # 记录日志
    except Exception as e:
        # 获取完整的堆栈跟踪信息
        tb = traceback.format_exc()
        logging.error(f"Error occurred in main: {str(e)}\n完整堆栈跟踪:\n{tb}")



