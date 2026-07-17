
import os
import odps 
import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta
import json
from datetime import date,timedelta
from odps import ODPS
import math
import itertools
import operator
import openpyxl
import datetime
import logging


o = ODPS(
    access_id=os.getenv("ODPS_ACCESS_KEY"),
    secret_access_key=os.getenv("ODPS_SECRET_KEY"),
    project=os.getenv("ODPS_PROJECT"),
    endpoint=os.getenv("ODPS_ENDPOINT")
)

class Get_info_json():
    def __init__(self,config_file='config.txt'):
        with open(config_file,'r') as f:
            self.config = eval(f.read())
        self.yesterday_date = None
        self.month_first_date = None
        self.yesterday_date_str = None
        self.month_first_date_str = None
        self.today_date_str()
        self.final_json = {}
        
    def today_date_str(self):
        self.yesterday_date = datetime.datetime.today().date()-datetime.timedelta(days=1)
        self.month_first_date = datetime.date(self.yesterday_date.year,self.yesterday_date.month,1)
        self.yesterday_date_str = str(self.yesterday_date)
        self.month_first_date_str = str(self.month_first_date)
        
    def get_info1(self):
        big_panel_json = {}
        
        sql1 = """
        with datatable as (SELECT  biz_date "业务日期"
                ,fst_level_channel "渠道"
                ,campaign_id "活动id"
                ,CASE   WHEN product_type = 30 THEN '不分期'
                        WHEN product_type = 31 THEN '分期'
                END AS "产品类型"
                ,product_user_type "新老客"
                ,reloan_type "老客类型"
                ,CASE   WHEN is_calm_period_repay = 1 THEN '是'
                        ELSE '否'
                END AS "是否退款"
                ,withdraw_succ_listing_cnt "成交笔数"
                ,withdraw_succ_user_cnt "成交人数"
                ,withdraw_succ_amount "成交金额"
        FROM    pk_data.dws_asset_natdate_deal_cnt
        where product_type != 32)

        SELECT 业务日期,新老客,是否退款,SUM(成交笔数) as 成交笔数,SUM(成交人数) as 成交人数,SUM(成交金额) as 成交金额
        from datatable 
        WHERE 业务日期>='{month_first_day}'
        and 业务日期<='{yesterday}'
        GROUP BY 业务日期,新老客,是否退款
        """.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)

        month_eff = o.execute_sql(sql1).open_reader(tunnel=True).to_pandas()
        month_eff['成交金额']=month_eff['成交金额'].astype(float)
        
        big_panel_json['日总放款金额(排除冷静期)'] = round(month_eff[(month_eff['业务日期']==self.yesterday_date_str)&(month_eff['是否退款']=='否')]['成交金额'].sum()/self.config['汇率'],1)
        big_panel_json['日总放款金额占月度目标'] = round(big_panel_json['日总放款金额(排除冷静期)']/self.config['月放款金额目标'],4)
        big_panel_json['日总放款金额本月累计占月度目标'] = round(month_eff[(month_eff['是否退款']=='否')]['成交金额'].sum()/self.config['汇率']/self.config['月放款金额目标'],4)
        big_panel_json['新客放款订单数'] = month_eff[(month_eff['业务日期']==self.yesterday_date_str)&(month_eff['是否退款']=='否')&(month_eff['新老客']=='新客')]['成交笔数'].sum()
        big_panel_json['老客放款订单数'] = month_eff[(month_eff['业务日期']==self.yesterday_date_str)&(month_eff['是否退款']=='否')&(month_eff['新老客']=='老客')]['成交笔数'].sum()
        big_panel_json['新客放款金额'] = round(month_eff[(month_eff['业务日期']==self.yesterday_date_str)&(month_eff['是否退款']=='否')&(month_eff['新老客']=='新客')]['成交金额'].sum()/self.config['汇率'],1)
        big_panel_json['老客放款金额'] = round(month_eff[(month_eff['业务日期']==self.yesterday_date_str)&(month_eff['是否退款']=='否')&(month_eff['新老客']=='老客')]['成交金额'].sum()/self.config['汇率'],1)
        big_panel_json['新客放款订单占月度目标'] = round(big_panel_json['新客放款订单数']/self.config['月新客放款订单数目标'],4)
        big_panel_json['老客放款订单占月度目标'] = round(big_panel_json['老客放款订单数']/self.config['月老客放款订单数目标'],4)
        big_panel_json['新客放款订单累计占月度目标'] = round(month_eff[(month_eff['业务日期']<=self.yesterday_date_str)&(month_eff['是否退款']=='否')&(month_eff['新老客']=='新客')]['成交笔数'].sum()/self.config['月新客放款订单数目标'],4)
        big_panel_json['老客放款订单累计占月度目标'] = round(month_eff[(month_eff['业务日期']<=self.yesterday_date_str)&(month_eff['是否退款']=='否')&(month_eff['新老客']=='老客')]['成交笔数'].sum()/self.config['月老客放款订单数目标'],4)
        return big_panel_json

    def get_info2(self):
        big_panel_json = {}
        sql='''with listing as (SELECT user_id, p_column["processFlag"] as processflag,substr(inserttime,1,10) as listing_date
            FROM pk_data.dwd_risk_user_pata_result_dicts_dly
            WHERE b_column["bizId"] IN ('22050')
            AND dt >= '2026-01-01'
            --AND i_column["isWhiteListUser"] = 0
            and i_column["isReloanCustomer"] = 1
            AND substr(inserttime,1,10) between '{month_first_day}' and '{yesterday}')

            select listing_date as dt,avg(cast(processflag as int)) as pass_rate
            from listing
            group by listing_date
        ;
        '''.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)
        data_all = o.execute_sql(sql).open_reader(tunnel=True).to_pandas().fillna(0)
        data = data_all
        big_panel_json['老客循环贷发标通过率'] = data[data['dt']==self.yesterday_date_str]['pass_rate'].max()
        
        return big_panel_json


    def get_info3(self):
        big_panel_json = {}
        before3_date = self.yesterday_date-datetime.timedelta(days=3)
        sql3='''select 
        register_date as dt
        ,sum(register_user_cnt) as register_user_cnt  -- 注册用户数
        ,sum(apply_limit_user_cnt) as apply_limit_user_cnt  -- 申请戳额用户数
        ,sum(fst_apply_has_limit_user_cnt) as fst_apply_has_limit_user_cnt  -- 首次戳额有额用户数
        ,sum(fst_apply_total_limit) as fst_apply_total_limit  -- 首次戳额给额额度总值
        ,sum(fst_has_limit_user_cnt) as fst_has_limit_user_cnt  -- 首次有额用户数(不一定是首戳)
        ,sum(fst_has_total_limit) as fst_has_total_limit  -- 首次有额额度总值(不一定是首戳)
        ,sum(apply_loan_user_cnt) as apply_loan_user_cnt  -- 申请借款用户数
        ,sum(loan_pass_user_cnt) as loan_pass_user_cnt  -- 借款通过用户数
        ,sum(full_bid_user_cnt) as full_bid_user_cnt  -- 发标满标人数
        ,sum(withdraw_succ_user_cnt) as withdraw_succ_user_cnt -- 提现(打款)成功人数
        ,sum(withdraw_succ_amount) as withdraw_succ_amount -- 提现(打款)成功金额
        ,sum(withdraw_succ_limit) as withdraw_succ_limit  -- 提现成功对应戳额额度总值(提现成功用户对应的额度)
        from pk_anls.tmp_prod_xk_link_conv_cnt_by_register
        where register_date between '{before3_date}' and '{yesterday}'
        and asset_product='cashloan'
        group by register_date
        '''.format(before3_date=str(before3_date),yesterday=self.yesterday_date_str)
        data = o.execute_sql(sql3).open_reader(tunnel=True).to_pandas().fillna(0)
        data['dt'] = data['dt'].astype(str)
        data['端内T3新客注册申请率'] = (data['apply_limit_user_cnt']/data['register_user_cnt']).astype(float)
        data['新客有额提现率'] = (data['withdraw_succ_user_cnt']/data['fst_has_limit_user_cnt']).astype(float)

        big_panel_json['新客有额提现率'] = data[data['dt']==self.yesterday_date_str]['新客有额提现率'].sum()
        # big_panel_json['新客有额提现率和目标比值'] = big_panel_json['新客有额提现率']/self.config['新客有额提现率目标']
        big_panel_json['端内T3新客注册申请率'] = data[data['dt']==str(before3_date)]['端内T3新客注册申请率'].sum()
        big_panel_json['T3新客注册申请率和目标比值'] = big_panel_json['端内T3新客注册申请率']/self.config['端内T3新客注册申请率目标']
        return big_panel_json



    def get_info4(self):
        big_panel_json = {}
        sql4='''select 
            a.biz_date as dt
            ,sum(a.register_user_cnt) as register_user_cnt     -- 注册人数
            ,sum(a.fst_apply_limit_user_cnt) as fst_apply_limit_user_cnt  -- 首戳人数(新客戳额人数)
            ,sum(a.repay_user_cnt) as repay_user_cnt   -- 还款人数
            ,sum(a.apply_limit_user_cnt_t0) as apply_limit_user_cnt_t0  -- 还款后戳额人数t0
            ,sum(a.fst_apply_has_limit_user_cnt) as fst_apply_has_limit_user_cnt    -- 首戳有额人数(新客戳额有额人数)
            ,sum(a.lk_loan_pass_amount) as lk_loan_pass_amount    -- 老客借款通过金额
            ,sum(a.lk_loan_pass_cnt) as lk_loan_pass_cnt          -- 老客借款通过次数
        from pk_anls.tmp_prod_natdate_oper_conv_cnt as a
        where biz_date between '{month_first_day}' and '{yesterday}'
        and asset_product='cashloan'
        group by a.biz_date
        ;'''.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)
        data = o.execute_sql(sql4).open_reader(tunnel=True).to_pandas()
        data['dt']=data['dt'].astype(str)
        data['复贷申请率T0'] = (data['apply_limit_user_cnt_t0']/data['repay_user_cnt']).astype(float)
        
        # data['注册戳额率(自然日)'] = (data['fst_apply_limit_user_cnt']/data['register_user_cnt']).astype(float)
        data['首次戳额有额率'] = (data['fst_apply_has_limit_user_cnt']/data['fst_apply_limit_user_cnt']).astype(float)

        big_panel_json['复贷申请率T0'] = data[data['dt']==self.yesterday_date_str]['复贷申请率T0'].sum()
        big_panel_json['复贷申请率T0和目标比值'] = big_panel_json['复贷申请率T0']/self.config['复贷申请率T0目标']
        # big_panel_json['注册戳额率(自然日)'] = data[data['dt']==self.yesterday_date_str]['注册戳额率(自然日)'].sum()
        # big_panel_json['注册戳额率(自然日)和目标比值']=big_panel_json['注册戳额率(自然日)']/self.config['自然日注册戳额率目标']
        big_panel_json['新客通过率(首戳)'] = data[data['dt']==self.yesterday_date_str]['首次戳额有额率'].sum()
        
        return big_panel_json


    def get_info5(self):
        big_panel_json = {}
        sql5 = '''select 
            repay_date as dt
            ,sum(has_limit_user_cnt) as has_limit_user_cnt  -- 戳额有额人数
            ,sum(reloan_withdraw_user_cnt) as reloan_withdraw_user_cnt  -- 复贷提现人数
            from pk_anls.tmp_prod_settle_reloan_conv_by_settle_date
            where repay_date between '{month_first_day}' and '{yesterday}'
            group by repay_date
            '''.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)
        data = o.execute_sql(sql5).open_reader(tunnel=True).to_pandas()
        data['老客有额提现率'] = (data['reloan_withdraw_user_cnt']/data['has_limit_user_cnt']).astype(float)

        big_panel_json['老客有额提现率'] = data[data['dt']==self.yesterday_date_str]['老客有额提现率'].sum()
        return big_panel_json

    
    def get_info6(self):
        big_panel_json={}
        sql6='''SELECT  t1.biz_date as dt----日期
        ,t1.cost ---消耗
        ,t2.fst_withdraw_succ_cnt
        ,t1.cost / t2.fst_withdraw_succ_cnt AS s_cps
        FROM    (
                    SELECT  a.biz_date
                            ,SUM(a.cost) AS cost
                    FROM    pk_data.dws_mkt_media_adset_placement_report_data a
                    where   a.fst_level_channel in ('tiktok','google','facebook')
                    and a.biz_date between '{month_first_day}' and '{yesterday}'
                    GROUP BY a.biz_date
                ) t1
        LEFT JOIN   (
                        SELECT  b.biz_date
                                ,SUM(b.fst_withdraw_succ_cnt) AS fst_withdraw_succ_cnt
                        FROM    pk_data.dws_asset_natdate_conv_cnt b
                        where asset_product='cashloan'
                        GROUP BY b.biz_date
                    ) t2
        ON      t1.biz_date = t2.biz_date;'''.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)
        data = o.execute_sql(sql6).open_reader(tunnel=True).to_pandas()

        # print(data[data['dt']==self.yesterday_date_str])
        big_panel_json['新客静态CPS(排除冷静期)'] = float(data[data['dt']==self.yesterday_date_str]['s_cps'].sum())
        big_panel_json['新客静态CPS和目标比值'] = float(big_panel_json['新客静态CPS(排除冷静期)']/self.config['新客静态CPS目标'])
        return big_panel_json


    def get_info7(self):
        big_panel_json={}
        sql7='''with term1 as (SELECT  listing_id,principal,list_amount,due_date
        ,CASE WHEN (repay_time IS NOT NULL and debt_status=2) THEN DATEDIFF(TO_DATE(repay_time),TO_DATE(due_date)) 
        else DATEDIFF(TO_DATE(NOW()),TO_DATE(due_date)) END as term1_overdue_days
        ,DATEDIFF(TO_DATE(NOW()),TO_DATE(due_date)) as term1_due_days
        FROM    pk_data.dwd_asset_loan_debt
        WHERE   asset_product in ('cashloan','bnpl_chuanyin')
        AND     is_calm_period_repay = 0
        and is_deal=1
        AND isactive = 1
        and period_seq=1),

        listing as (SELECT listing_id,term_quantity,product_user_type,asset_product
        from pk_data.dwb_asset_cmn_listing  
        where asset_subject is not NULL
        and asset_product in ('cashloan','bnpl_chuanyin') )

        SELECT term1.*,term_quantity,product_user_type,asset_product
        ,case when term1_overdue_days>=1 then 1 else 0 end as term1_dpd1
        ,case when term1_overdue_days>=7 then 1 else 0 end as term1_dpd7
        from term1 left join listing on term1.listing_id=listing.listing_id
        WHERE term1_due_days in (1,7)
        '''.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)
        data = o.execute_sql(sql7).open_reader(tunnel=True).to_pandas()
        data[['principal','list_amount']] = data[['principal','list_amount']].astype(float)

        temp = data[(data['product_user_type']=='新客')&(data['term1_due_days']==1)&(data['asset_product']=='cashloan')]
        big_panel_json['最新到fpd1放款日期'] = str(pd.to_datetime(temp['due_date']).dt.date.max()-pd.Timedelta(days=14))
        big_panel_json['新客$fpd1'] = (temp['term1_dpd1']*temp['principal']).sum()/temp['principal'].sum()
        temp = data[(data['product_user_type']=='老客')&(data['term1_due_days']==1)&(data['asset_product']=='cashloan')]
        big_panel_json['老客$fpd1'] = (temp['term1_dpd1']*temp['principal']).sum()/temp['principal'].sum()
        temp = data[(data['product_user_type']=='新客')&(data['term1_due_days']==7)&(data['asset_product']=='cashloan')]
        big_panel_json['最新到fpd7放款日期'] = str(pd.to_datetime(temp['due_date']).dt.date.max()-pd.Timedelta(days=14))
        big_panel_json['新客$fpd7'] = (temp['term1_dpd7']*temp['principal']).sum()/temp['principal'].sum()
        big_panel_json['新客fpd7金订比'] = big_panel_json['新客$fpd7']/temp['term1_dpd7'].mean()
        temp = data[(data['product_user_type']=='老客')&(data['term1_due_days']==7)&(data['asset_product']=='cashloan')]
        big_panel_json['老客$fpd7'] = (temp['term1_dpd7']*temp['principal']).sum()/temp['principal'].sum()
        big_panel_json['老客fpd7金订比'] = big_panel_json['老客$fpd7']/temp['term1_dpd7'].mean()

        temp = data[(data['term1_due_days']==1)&(data['asset_product']=='bnpl_chuanyin')]
        big_panel_json['bnpl$fpd1'] = (temp['term1_dpd1']*temp['principal']).sum()/temp['principal'].sum()
        big_panel_json['bnplfpd1到期数量'] = temp.shape[0]
        temp = data[(data['term1_due_days']==7)&(data['asset_product']=='bnpl_chuanyin')]
        big_panel_json['bnpl$fpd7'] = (temp['term1_dpd7']*temp['principal']).sum()/temp['principal'].sum()
        big_panel_json['bnplfpd7到期数量'] = temp.shape[0]
        
        return big_panel_json


    def get_info8(self):
        big_panel_json={}
        sql='''with withdraw as (SELECT  user_id
        ,listing_id
        ,apply_amount
        ,product_type
        ,CASE   WHEN total_bid_num['deal_non_calm_repay_asc'] = 1 THEN '新客' ELSE '老客' END AS user_type
        ,substr(withdraw_time,1,10) as withdraw_date 
        ,total_period
        ,row_number() over(partition by user_id order by withdraw_time) as loan_n
        FROM    pk_data.dwb_asset_cmn_listing
        WHERE   asset_product='cashloan'
        AND     withdraw_time IS NOT NULL),

        termdata AS 
        (
            SELECT  listing_id,due_date,substr(repay_time,1,10) as repay_date
            FROM    pk_data.dwd_asset_loan_debt
            WHERE   asset_product='cashloan'
            and period_no = period_seq
        )


        SELECT withdraw.*,repay_date as dt
        from withdraw left join termdata on withdraw.listing_id = termdata.listing_id
        where repay_date between '{month_first_day}' and '{yesterday}';
        '''.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)
        data = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()

        big_panel_json['新客结清订单数'] = data[(data['dt']==self.yesterday_date_str)&(data['user_type']=='新客')].shape[0]
        big_panel_json['老客结清订单数'] = data[(data['dt']==self.yesterday_date_str)&(data['user_type']=='老客')].shape[0]
        return big_panel_json

    def get_info9(self):
        big_panel_json={}
        sql='''with lim as (select user_id,inserttime as limit_time,SUBSTR(inserttime,1,10) AS limit_date,flow_id,p_column["processFlag"] AS limit_processFlag
        ,i_column["isReloanCustomer"] as is_reloan,u_column['userPayedLoansCnt'] as repay_cnt
        ,case when b_column["bizId"] in ("10000","10001") then 'cashloan' else 'bnpl' end as biz_type
        FROM    pk_data.dwd_risk_user_pata_result_dicts_dly
        WHERE   b_column["bizId"] in ("10000","10001","11000") -- 戳额
        AND     dt >= "2024-12-30"
        AND     (i_column["isWhiteListUser"] = 0 or i_column["isWhiteListUser"] is null)
        and substr(inserttime,1,10) between '{month_first_day}' and '{yesterday}')

        select limit_date as dt,is_reloan,biz_type,count(flow_id) as num,sum(limit_processFlag) as is_pass,avg(limit_processFlag) as pass_rate
        ,count(distinct user_id) as limit_num
        ,avg(case when repay_cnt='1' then limit_processFlag else null end) as pass_rate_newtoold
        from lim
        group by limit_date,is_reloan,biz_type
        ;
        '''.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)
        data_all = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()
        data_all['dt']=data_all['dt'].astype(str)
        data = data_all[data_all['biz_type']=='cashloan']
        data_bnpl = data_all[data_all['biz_type']=='bnpl']
        
        big_panel_json['bnpl申请人数'] = data_bnpl[(data_bnpl['dt']==self.yesterday_date_str)]['limit_num'].sum()
        big_panel_json['本月bnpl申请人数'] = data_bnpl['num'].sum()
        big_panel_json['bnpl申请通过率'] = data_bnpl[(data_bnpl['dt']==self.yesterday_date_str)]['is_pass'].sum()/data_bnpl[(data_bnpl['dt']==self.yesterday_date_str)]['num'].sum()

        big_panel_json['自然日新客戳额人数'] = data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']=='0')]['limit_num'].sum()
        big_panel_json['自然日新客戳额通过率'] = data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']=='0')]['pass_rate'].mean()
        big_panel_json['自然日老客戳额人数'] = data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']=='1')]['limit_num'].sum()
        big_panel_json['自然日老客戳额通过率'] = data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']=='1')]['pass_rate'].mean()
        big_panel_json['自然日新转老戳额通过率'] = data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']=='1')]['pass_rate_newtoold'].mean()
        return big_panel_json

    def get_info10(self):
        # asset_product='bnpl_chuanyin'
        big_panel_json={}
        sql='''with loan as (select withdraw_time,real_amount,SUBSTR(withdraw_time,1,10) AS withdraw_date,listing_id,user_id
        ,CASE when product_user_type='老客' then 1 else 0 end as is_reloan,period_no,asset_product
        from pk_data.dwb_asset_cmn_listing
        WHERE   is_deal = 1
        AND     asset_product in ('cashloan','bnpl_chuanyin')
        and bid_stage >= "80" 
        and substr(withdraw_time,1,10) between '{month_first_day}' and '{yesterday}'),

        lim as (SELECT distinct user_id
        FROM    pk_data.dwd_risk_user_pata_result_dicts_dly
        WHERE   dt >= '2026-02-03'
        AND     b_column["bizId"] = "12001"
        AND     inserttime >= "2026-02-03 18:26:00"
        AND     i_column["isCyclic"] = "true")

        select withdraw_date as dt,is_reloan,asset_product,case when lim.user_id is null then 'notre' else 're' end as re_type,avg(real_amount) as real_amount,avg(cast(period_no as int)) as period_no
        ,sum(real_amount*cast(period_no as int)) as period_multi_amount,sum(real_amount) as amount_sum,count(distinct listing_id) as deal_num,count(listing_id) as num
        from loan left join lim on loan.user_id = lim.user_id
        group by withdraw_date,is_reloan,asset_product,case when lim.user_id is null then 'notre' else 're' end
        ;
        '''.format(month_first_day=self.month_first_date_str,yesterday=self.yesterday_date_str)
        data_all = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()
        data = data_all[(data_all['re_type'].isin(['notre','re']))&(data_all['asset_product']=='cashloan')]

        big_panel_json['新客件均'] = float(data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==0)]['amount_sum'].sum())/self.config['汇率']/data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==0)]['num'].sum()
        big_panel_json['新客件均和目标比值'] = big_panel_json['新客件均']/self.config['新客件均目标']
        big_panel_json['老客件均'] = float(data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['amount_sum'].sum())/self.config['汇率']/data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['num'].sum()
        big_panel_json['老客件均和目标比值'] = big_panel_json['老客件均']/self.config['老客件均目标']
        # big_panel_json['新客平均期限'] = float(data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']=='0')]['period_no'].astype(float).sum()*15)
        big_panel_json['新客平均期限'] = data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==0)]['period_multi_amount'].astype(float).sum()/data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==0)]['amount_sum'].astype(float).sum()*15
        big_panel_json['新客平均期限和目标比值'] = big_panel_json['新客平均期限']/self.config['新客平均期限目标']
        # big_panel_json['老客平均期限'] = float(data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']=='1')]['period_no'].astype(float).sum()*15)
        big_panel_json['老客平均期限'] = data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['period_multi_amount'].astype(float).sum()/data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['amount_sum'].astype(float).sum()*15
        big_panel_json['老客平均期限和目标比值'] = big_panel_json['老客平均期限']/self.config['老客平均期限目标']

        data = data_all[(data_all['re_type'].isin(['re']))&(data_all['asset_product']=='cashloan')]
        big_panel_json['老客循环贷件均'] = float(data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['amount_sum'].sum())/self.config['汇率']/data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['num'].sum()
        big_panel_json['老客循环贷放款订单数'] = float(data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['num'].sum())
        big_panel_json['老客循环贷放款金额'] = float(data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['amount_sum'].sum())/self.config['汇率']

        data = data_all[(data_all['re_type'].isin(['notre']))&(data_all['asset_product']=='cashloan')]
        big_panel_json['老客非循环贷件均'] = float(data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['amount_sum'].sum())/self.config['汇率']/data[(data['dt']==self.yesterday_date_str)&(data['is_reloan']==1)]['num'].sum()

        data = data_all[data_all['asset_product']=='bnpl_chuanyin']
        big_panel_json['bnpl平均期限'] = data[(data['dt']==self.yesterday_date_str)]['period_multi_amount'].astype(float).sum()/data[(data['dt']==self.yesterday_date_str)]['amount_sum'].astype(float).sum()*30
        return big_panel_json

    def get_info11(self):
        big_panel_json = {}

        sql = """WITH lim AS (
            SELECT
                user_id, inserttime AS limit_time, substr(inserttime,1,10) AS limit_date, flow_id, i_column["isReloanCustomer"] AS is_reloan,
                CASE WHEN to_date(inserttime) = to_date(date_add('{yesterday}', -7)) THEN 't7'
                    WHEN to_date(inserttime) = to_date('{yesterday}') THEN 't0' ELSE NULL END AS limit_date_tn
            FROM pk_data.dwd_risk_user_pata_result_dicts_dly
            WHERE b_column["bizId"] IN ('10000','10001')
            AND dt >= '2024-12-30'
            AND i_column["isWhiteListUser"] = 0
            AND p_column["processFlag"] = '1'
            AND to_date(inserttime) >= to_date(date_add('{yesterday}', -7))
        ),

        mid AS ( SELECT limit_flow_no,listing_id,substr(withdraw_time,1,10) AS withdraw_date
            FROM pk_data.ddm_asset_limit_loan_dtl
            WHERE asset_product = 'cashloan' ),

        loan AS ( SELECT listing_id
            FROM pk_data.dwd_asset_loan_list
            WHERE isactive = 1
            AND asset_product IS NOT NULL
            AND bid_stage >= '80' ),

        base AS (SELECT l.is_reloan,l.limit_date_tn,
                CASE WHEN datediff(to_date(m.withdraw_date), to_date(l.limit_date)) <= 0 THEN 1 ELSE 0 END AS is_in_0,
                CASE WHEN datediff(to_date(m.withdraw_date), to_date(l.limit_date)) <= 7 THEN 1 ELSE 0 END AS is_in_7
            FROM lim l
            LEFT JOIN mid m ON l.flow_id = m.limit_flow_no
            LEFT JOIN loan lo ON m.listing_id = lo.listing_id
            WHERE l.limit_date_tn IN ('t0','t7'))

        SELECT is_reloan, limit_date_tn,avg(is_in_0) AS withdraw_rate_t0,avg(is_in_7) AS withdraw_rate_t7
        FROM base
        GROUP BY is_reloan,limit_date_tn
        """.format(yesterday=self.yesterday_date_str)

        df = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()

        def get_rate(is_reloan, tn, col):
            row = df[(df['is_reloan'] == is_reloan) & (df['limit_date_tn'] == tn)]
            return row[col].iloc[0] if not row.empty else 0

        big_panel_json['新客有额提现率t0'] = get_rate('0', 't0', 'withdraw_rate_t0')
        big_panel_json['老客有额提现率t0'] = get_rate('1', 't0', 'withdraw_rate_t0')

        big_panel_json['新客有额提现率t7'] = get_rate('0', 't7', 'withdraw_rate_t7')
        big_panel_json['老客有额提现率t7'] = get_rate('1', 't7', 'withdraw_rate_t7')

        big_panel_json['新客有额提现率t0和目标比值'] = (big_panel_json['新客有额提现率t0'] / self.config['新客有额提现率t0目标'])
        big_panel_json['老客有额提现率t0和目标比值'] = (big_panel_json['老客有额提现率t0'] / self.config['老客有额提现率t0目标'])
        big_panel_json['新客有额提现率t7和目标比值'] = (big_panel_json['新客有额提现率t7'] / self.config['新客有额提现率t7目标'])
        big_panel_json['老客有额提现率t7和目标比值'] = (big_panel_json['老客有额提现率t7'] / self.config['老客有额提现率t7目标'])

        return big_panel_json



    def get_info12(self):
        big_panel_json={}
        sql1 = """
        SELECT count(distinct user_id) as reg_cnt
        from pk_data.s_dim_user_basic_info_snp 
        WHERE dt = MAX_PT('pk_data.s_dim_user_basic_info_snp')
        and to_date(register_time)=to_date('{yesterday}')
                """.format(yesterday=self.yesterday_date_str)
        data0 = o.execute_sql(sql1).open_reader(tunnel=True).to_pandas()

        sql1 = """
        select count(DISTINCT listing_id) as deal_cnt
        from pk_data.dwb_asset_cmn_listing
        where asset_subject is not null
        and asset_product='cashloan'
        and to_date(withdraw_time)='{yesterday}'
        and is_deal=1
        and product_user_type='新客'
        limit 100""".format(yesterday=self.yesterday_date_str)
        data1 = o.execute_sql(sql1).open_reader(tunnel=True).to_pandas()

        big_panel_json['新客静态注册成交率'] = float(data1.loc[0,'deal_cnt']/data0.loc[0,'reg_cnt'])
        big_panel_json['新客静态注册成交率和目标比值'] = big_panel_json['新客静态注册成交率']/self.config['静态注册成交率目标']
        return big_panel_json

    def get_info13(self):
        big_panel_json={}
        sql1 = """
        with cus as (
        select user_id,register_time,case when to_date(register_time)=to_date(date_add('{yesterday}',-7)) then 't7'
                        when to_date(register_time)=to_date(date_add('{yesterday}',-0)) then 't0' else null end as reg_date_tn
        from  pk_data.s_dim_user_basic_info_snp a_user
        where dt = MAX_PT('pk_data.s_dim_user_basic_info_snp')
        and to_date(register_time)>=to_date(date_add('{yesterday}',-8))),

        source as (SELECT user_id
        from pk_data.dwb_mkt_user_register_channel_dtl
        where fst_level_channel not in ('cy_bnpl','daraz')),

        lim as (select flow_id,inserttime as limit_time,user_id
        FROM    pk_data.dwd_risk_user_pata_result_dicts_dly
        WHERE   b_column["bizId"] in ("10000","10001")
        AND     dt >= "2025-03-25"
        AND     i_column["isWhiteListUser"] = 0
        and i_column["isReloanCustomer"]=0
        and to_date(inserttime)>=to_date(date_add('{yesterday}',-8)))

        SELECT cus.*, limit_time,to_date(register_time) as register_date,to_date(limit_time) as limit_date
        from cus left join lim on cus.user_id=lim.user_id
        inner join source on cus.user_id=source.user_id
                """.format(yesterday=self.yesterday_date_str)
        data0 = o.execute_sql(sql1).open_reader(tunnel=True).to_pandas()

        data0['days_diff'] = (pd.to_datetime(data0['limit_date'])-pd.to_datetime(data0['register_date'])).dt.days
        big_panel_json['注册戳额率t0'] = (data0[data0['reg_date_tn']=='t0']['days_diff']<=0).mean()
        big_panel_json['注册戳额率t7'] = (data0[data0['reg_date_tn']=='t7']['days_diff']<=7).mean()
        big_panel_json['注册戳额率t0和目标比值'] = big_panel_json['注册戳额率t0']/self.config['注册戳额率t0目标']
        big_panel_json['注册戳额率t7和目标比值'] = big_panel_json['注册戳额率t7']/self.config['注册戳额率t7目标']
        return big_panel_json

    def get_info14(self):
        big_panel_json={}
        sql1 = """
        SELECT  t1.biz_date
                ,t2.cost
                ,case when t1.biz_date = '{yesterday}' then 't0' when t1.biz_date = to_date(date_add('{yesterday}',-2)) then 't3' else null end as t
                ,COUNT(DISTINCT t1.user_id) AS register_cnt ----注册用户数
                ,COUNT(CASE    WHEN t1.biz_date = t1.withdraw_date and t1.is_deal=1 and t1.is_calm_period_repay=0 THEN t1.user_id END) AS fst_withdraw_succ_cnt ---T0首次提现用户数
                ,COUNT(DISTINCT CASE    WHEN DATEDIFF(t1.withdraw_date,t1.biz_date) <= 3  and t1.is_deal=1 and t1.is_calm_period_repay=0 THEN t1.user_id END) AS t3_withdraw_user_num --注册3天内新客提现人数
        FROM    (
                    SELECT  b.user_id
                            ,b.fst_level_channel
                            ,b.campaign_name
                            ,SUBSTR(b.register_time,1,10) AS biz_date ---注册日期
                            ,SUBSTR(a.withdraw_time,1,10) AS withdraw_date ---提现日期
                            ,a.is_deal
                            ,a.is_calm_period_repay
                    FROM    pk_data.dwb_mkt_user_register_channel_dtl b
                    LEFT JOIN pk_data.dwb_asset_limit_loan_dtl a
                    ON      a.user_id = b.user_id
                    WHERE  a.asset_product = 'cashloan'
                ) t1
        LEFT JOIN   (
                        SELECT  SUBSTR(a.biz_date,1,10) AS biz_date
                                ,SUM(a.cost) AS cost
                        FROM    pk_data.dws_mkt_media_adset_placement_report_data a
                        WHERE   a.fst_level_channel IN ('tiktok','google','facebook')
                        GROUP BY SUBSTR(a.biz_date,1,10)
                    ) t2
        ON      t1.biz_date = t2.biz_date
        where t1.biz_date in ('{yesterday}',to_date(date_add('{yesterday}',-2)))
        GROUP BY t1.biz_date,t2.cost """.format(yesterday=self.yesterday_date_str)
        data0 = o.execute_sql(sql1).open_reader(tunnel=True).to_pandas()

        big_panel_json['t0cps'] = (data0[data0['t']=='t0']['cost'].astype(float)/data0[data0['t']=='t0']['fst_withdraw_succ_cnt'].astype(float)).max()
        big_panel_json['t3cps'] = (data0[data0['t']=='t3']['cost'].astype(float)/data0[data0['t']=='t3']['t3_withdraw_user_num'].astype(float)).max()
        big_panel_json['t0cps和目标比值'] = big_panel_json['t0cps']/self.config['t0cps目标']
        big_panel_json['t3cps和目标比值'] = big_panel_json['t3cps']/self.config['t3cps目标']
        return big_panel_json


    def get_info15(self):
        big_panel_json={}
        sql1 = """
        with asset as (select withdraw_time,real_amount,listing_id,period_no,substr(withdraw_time,1,10) as dt 
        from pk_data.dwd_asset_loan_list
        WHERE   isactive = 1
        AND     asset_product='bnpl_chuanyin'
        and bid_stage >= "80"
        and TO_DATE(withdraw_time)>='{month_first_date_str}'),

        mid as (select loan_apply_no,listing_id
        from pk_data.ddm_asset_limit_loan_dtl
        where asset_product='bnpl_chuanyin'),

        down as (select loan_apply_no,down_payment_ratio
        from pk_data.dwd_trade_merchant_order_record),

        res as (SELECT asset.*,down_payment_ratio
        from asset left join mid on asset.listing_id = mid.listing_id
        left join down on mid.loan_apply_no = down.loan_apply_no)

        SELECT *
        from res;""".format(month_first_date_str=self.month_first_date_str)
        data0 = o.execute_sql(sql1).open_reader(tunnel=True).to_pandas()
        big_panel_json['bnpl成交笔数'] = data0[(data0['dt']==self.yesterday_date_str)].shape[0]
        big_panel_json['本月bnpl成交笔数'] = data0.shape[0]
        big_panel_json['bnpl成交金额'] = float(data0[(data0['dt']==self.yesterday_date_str)]['real_amount'].sum())/self.config['汇率']
        big_panel_json['bnpl平均期限'] = float(data0[(data0['dt']==self.yesterday_date_str)]['period_no'].astype(float).mean()*30)
        big_panel_json['bnpl平均首付比例'] = float(data0[(data0['dt']==self.yesterday_date_str)]['down_payment_ratio'].astype(float).mean())/100
        return big_panel_json

    
    def get_info16(self):
        big_panel_json = {}
        yesterday = self.yesterday_date_str
        # 计算前一天日期
        day_before = (datetime.datetime.strptime(yesterday, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')

        sql = """
        WITH t0_lim AS (
            SELECT user_id, flow_id,
                i_column["isReloanCustomer"] AS is_reloan,
                u_column['userPayedLoansCnt'] AS repay_cnt,
                to_date(inserttime) AS limit_date
            FROM pk_data.dwd_risk_user_pata_result_dicts_dly
            WHERE dt >= '2026-01-01'
            AND b_column["bizId"] IN ('10000','10001')
            AND i_column["isWhiteListUser"] = 0
            AND p_column["processFlag"] = '1'
            AND i_column["isReloanCustomer"] = 1
            AND to_date(inserttime) IN ('{yesterday}', '{day_before}')
        ),
        withdraw_data AS (
            SELECT l.user_id, l.flow_id, l.limit_date,
                CASE WHEN CAST(l.repay_cnt AS INT)=1 THEN '新转老' ELSE '非新转老' END AS loan_type,
                CASE WHEN m.listing_id IS NOT NULL AND to_date(m.withdraw_time)=l.limit_date THEN 1 ELSE 0 END AS is_withdraw
            FROM t0_lim l
            LEFT JOIN pk_data.ddm_asset_limit_loan_dtl m 
                ON l.flow_id=m.limit_flow_no AND m.asset_product='cashloan'
            LEFT JOIN pk_data.dwd_asset_loan_list lo 
                ON m.listing_id=lo.listing_id AND lo.bid_stage>='80'
        )
        SELECT loan_type, cast(limit_date as string) as limit_date,
            COUNT(DISTINCT user_id) AS user_cnt,
            SUM(is_withdraw) AS withdraw_cnt,
            AVG(is_withdraw) AS withdraw_rate
        FROM withdraw_data
        GROUP BY loan_type, limit_date
        """.format(yesterday=yesterday, day_before=day_before)

        data = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()

        # 提取昨日数据
        ndata_y = data[(data['loan_type']=='新转老') & (data['limit_date']==yesterday)]
        odata_y = data[(data['loan_type']=='非新转老') & (data['limit_date']==yesterday)]
        # 提取前日数据
        ndata_b = data[(data['loan_type']=='新转老') & (data['limit_date']==day_before)]
        odata_b = data[(data['loan_type']=='非新转老') & (data['limit_date']==day_before)]

        rate_n_y = float(ndata_y['withdraw_rate'].iloc[0]) if not ndata_y.empty else 0
        rate_o_y = float(odata_y['withdraw_rate'].iloc[0]) if not odata_y.empty else 0
        rate_n_b = float(ndata_b['withdraw_rate'].iloc[0]) if not ndata_b.empty else None
        rate_o_b = float(odata_b['withdraw_rate'].iloc[0]) if not odata_b.empty else None

        big_panel_json['单笔单批新转老有额提现率T0'] = rate_n_y
        big_panel_json['单笔单批新转老有额提现率T0和目标比值'] = rate_n_y / self.config.get('单笔单批新转老有额提现率T0目标', 1)
        big_panel_json['单笔单批新转老有额提现率T0环比'] = (rate_n_y - rate_n_b) if rate_n_b is not None else None

        big_panel_json['单笔单批非新转老有额提现率T0'] = rate_o_y
        big_panel_json['单笔单批非新转老有额提现率T0和目标比值'] = rate_o_y / self.config.get('单笔单批非新转老有额提现率T0目标', 1)
        big_panel_json['单笔单批非新转老有额提现率T0环比'] = (rate_o_y - rate_o_b) if rate_o_b is not None else None

        return big_panel_json

    def get_info17(self):
        big_panel_json = {}
        yesterday = self.yesterday_date_str
        day_before = (datetime.datetime.strptime(yesterday, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')

        sql = """
        WITH t0_lim AS (
            SELECT user_id, flow_id,
                i_column["isReloanCustomer"] AS is_reloan,
                c_column['creditLimit'] AS credit_limit,
                to_date(inserttime) AS limit_date
            FROM pk_data.dwd_risk_user_pata_result_dicts_dly
            WHERE dt >= '2026-01-01'
            AND b_column["bizId"] = '12001'
            AND i_column["isReloanCustomer"] = 1
            AND i_column["isCyclic"] = 'true'
            AND to_date(inserttime) IN ('{yesterday}', '{day_before}')
        ),
        user_withdraw AS (
            SELECT user_id, withdraw_time,
                ROW_NUMBER() OVER (PARTITION BY user_id ORDER BY withdraw_time ASC) AS rn
            FROM pk_data.dwd_asset_loan_list
            WHERE asset_product = 'cashloan'
            AND bid_stage >= '80'
            AND to_date(withdraw_time) IN ('{yesterday}', '{day_before}')
        )
        SELECT cast(l.limit_date as string) as limit_date,
            COUNT(DISTINCT l.user_id) AS total_user_cnt,
            SUM(CASE WHEN CAST(l.credit_limit AS FLOAT) > 0 THEN 1 ELSE 0 END) AS has_limit_cnt,
            SUM(CASE WHEN CAST(l.credit_limit AS FLOAT) > 0 AND w.user_id IS NOT NULL THEN 1 ELSE 0 END) AS withdraw_cnt,
            CAST(SUM(CASE WHEN CAST(l.credit_limit AS FLOAT) > 0 AND w.user_id IS NOT NULL THEN 1 ELSE 0 END) AS DOUBLE)
                / NULLIF(SUM(CASE WHEN CAST(l.credit_limit AS FLOAT) > 0 THEN 1 ELSE 0 END), 0) AS withdraw_rate
        FROM t0_lim l
        LEFT JOIN user_withdraw w ON l.user_id = w.user_id AND w.rn = 1 AND l.limit_date = to_date(w.withdraw_time)
        GROUP BY l.limit_date
        """.format(yesterday=yesterday, day_before=day_before)

        data = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()

        rate_y = float(data[data['limit_date']==yesterday]['withdraw_rate'].iloc[0]) \
                if not data[data['limit_date']==yesterday].empty and data[data['limit_date']==yesterday]['withdraw_rate'].iloc[0] is not None else 0
        rate_b = float(data[data['limit_date']==day_before]['withdraw_rate'].iloc[0]) \
                if not data[data['limit_date']==day_before].empty and data[data['limit_date']==day_before]['withdraw_rate'].iloc[0] is not None else None

        big_panel_json['循环贷老客有额提现率T0'] = rate_y
        big_panel_json['循环贷老客有额提现率T0和目标比值'] = rate_y / self.config.get('循环贷老客有额提现率T0目标', 1)
        big_panel_json['循环贷老客有额提现率T0环比'] = (rate_y - rate_b) if rate_b is not None else None

        return big_panel_json

    def get_info18(self):
        big_panel_json = {}
        yesterday = self.yesterday_date_str
        day_before = (datetime.datetime.strptime(yesterday, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')

        sql = """
        WITH t0_lim AS (
            SELECT user_id, flow_id,
                i_column["isReloanCustomer"] AS is_reloan,
                c_column['creditLimit'] AS credit_limit,
                a_column['auditNewRiskGradeV1'] AS risk_grade,
                to_date(inserttime) AS limit_date
            FROM pk_data.dwd_risk_user_pata_result_dicts_dly
            WHERE dt >= '2026-01-01'
            AND b_column["bizId"] IN ('10000','10001')
            AND i_column["isWhiteListUser"] = 0
            AND p_column["processFlag"] = '1'
            AND i_column["isReloanCustomer"] = 0
            AND to_date(inserttime) IN ('{yesterday}', '{day_before}')
        ),
        loan_data AS (
            SELECT limit_flow_no, listing_id,
                to_date(loan_apply_time) AS loan_date
            FROM pk_data.ddm_asset_limit_loan_dtl
            WHERE asset_product = 'cashloan'
        )
        SELECT cast(l.limit_date as string) as limit_date,
            SUM(CASE WHEN CAST(credit_limit AS FLOAT) > 0 THEN 1 ELSE 0 END) AS has_limit_cnt,
            SUM(CASE WHEN CAST(credit_limit AS FLOAT) > 0 AND datediff(ld.loan_date, l.limit_date) = 0 THEN 1 ELSE 0 END) AS bid_cnt,
            CAST(SUM(CASE WHEN CAST(credit_limit AS FLOAT) > 0 AND datediff(ld.loan_date, l.limit_date) = 0 THEN 1 ELSE 0 END) AS DOUBLE) 
                / NULLIF(SUM(CASE WHEN CAST(credit_limit AS FLOAT) > 0 THEN 1 ELSE 0 END), 0) AS bid_rate,
            SUM(CASE WHEN risk_grade IN ('A','B','C') AND CAST(credit_limit AS FLOAT) > 0 THEN 1 ELSE 0 END) AS high_quality_has_limit_cnt,
            SUM(CASE WHEN risk_grade IN ('A','B','C') AND CAST(credit_limit AS FLOAT) > 0 AND datediff(ld.loan_date, l.limit_date) = 0 THEN 1 ELSE 0 END) AS high_quality_bid_cnt,
            CAST(SUM(CASE WHEN risk_grade IN ('A','B','C') AND CAST(credit_limit AS FLOAT) > 0 AND datediff(ld.loan_date, l.limit_date) = 0 THEN 1 ELSE 0 END) AS DOUBLE)
                / NULLIF(SUM(CASE WHEN risk_grade IN ('A','B','C') AND CAST(credit_limit AS FLOAT) > 0 THEN 1 ELSE 0 END), 0) AS high_quality_bid_rate
        FROM t0_lim l
        LEFT JOIN loan_data ld ON l.flow_id = ld.limit_flow_no
        GROUP BY l.limit_date
        """.format(yesterday=yesterday, day_before=day_before)

        data = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()

        def get_rate(df, date, col):
            row = df[df['limit_date']==date]
            return float(row[col].iloc[0]) if not row.empty and row[col].iloc[0] is not None else (0 if date==yesterday else None)

        rate_y = get_rate(data, yesterday, 'bid_rate')
        rate_b = get_rate(data, day_before, 'bid_rate')
        high_rate_y = get_rate(data, yesterday, 'high_quality_bid_rate')
        high_rate_b = get_rate(data, day_before, 'high_quality_bid_rate')

        big_panel_json['新客有额发标率T0'] = rate_y
        big_panel_json['新客有额发标率T0和目标比值'] = rate_y / self.config.get('新客有额发标率T0目标', 1)
        big_panel_json['新客有额发标率T0环比'] = (rate_y - rate_b) if rate_b is not None else None

        big_panel_json['新客优质客群有额发标率T0'] = high_rate_y
        big_panel_json['新客优质客群有额发标率T0和目标比值'] = high_rate_y / self.config.get('新客优质客群有额发标率T0目标', 1)
        big_panel_json['新客优质客群有额发标率T0环比'] = (high_rate_y - high_rate_b) if high_rate_b is not None else None

        return big_panel_json


    def get_info19(self):
        big_panel_json = {}
        yesterday = self.yesterday_date_str
        day_before = (datetime.datetime.strptime(yesterday, '%Y-%m-%d') - timedelta(days=1)).strftime('%Y-%m-%d')

        sql = """
        WITH reg AS (
            SELECT user_id, to_date(register_time) AS reg_date
            FROM pk_data.s_dim_user_basic_info_snp
            WHERE dt = MAX_PT('pk_data.s_dim_user_basic_info_snp')
            AND to_date(register_time) IN ('{yesterday}', '{day_before}')
        ),
        lim AS (
            SELECT user_id, to_date(inserttime) AS lim_date
            FROM pk_data.dwd_risk_user_pata_result_dicts_dly
            WHERE dt >= '2026-01-01'
            AND b_column["bizId"] IN ('10000','10001')
            AND i_column["isWhiteListUser"] = 0
            AND to_date(inserttime) IN ('{yesterday}', '{day_before}')
        )
        SELECT cast(r.reg_date as string) as reg_date,
            COUNT(DISTINCT r.user_id) AS reg_cnt,
            COUNT(DISTINCT l.user_id) AS lim_cnt,
            COUNT(DISTINCT l.user_id) * 1.0 / COUNT(DISTINCT r.user_id) AS reg_to_lim_rate_t0
        FROM reg r
        LEFT JOIN lim l ON r.user_id = l.user_id AND r.reg_date = l.lim_date
        GROUP BY r.reg_date
        """.format(yesterday=yesterday, day_before=day_before)

        data = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()

        rate_y = float(data[data['reg_date']==yesterday]['reg_to_lim_rate_t0'].iloc[0]) \
                if not data[data['reg_date']==yesterday].empty and data[data['reg_date']==yesterday]['reg_to_lim_rate_t0'].iloc[0] is not None else 0
        rate_b = float(data[data['reg_date']==day_before]['reg_to_lim_rate_t0'].iloc[0]) \
                if not data[data['reg_date']==day_before].empty and data[data['reg_date']==day_before]['reg_to_lim_rate_t0'].iloc[0] is not None else None

        big_panel_json['注册戳额率T0'] = rate_y
        big_panel_json['注册戳额率T0和目标比值'] = rate_y / self.config.get('注册戳额率T0目标', 1)
        big_panel_json['注册戳额率T0环比'] = (rate_y - rate_b) if rate_b is not None else None

        return big_panel_json

    def get_info20(self):
        big_panel_json = {}
        yesterday = self.yesterday_date_str
        sql = """
        with listing as (select user_id,inserttime,flow_id as loan_flow_no,listing_id,b_column["bizId"] as biz_id,cast(p_column['processFlag'] as int) as listing_process_flag
        ,cast(i_column["isReloanCustomer"] as int) as is_reloan,case when cast(u_column['userPayedLoansCnt'] as int)=1 then 1 else 0 end AS is_new_to_old
        FROM    pk_data.dwd_risk_user_pata_result_dicts_dly
        WHERE   b_column["bizId"] in ("20050",'22050') -- 戳额
        AND     dt = "{yesterday}")

        SELECT is_reloan,is_new_to_old,COUNT(listing_process_flag) as listing_num,sum(listing_process_flag) as pass_num,avg(listing_process_flag) as pass_rate
        from listing
        GROUP BY is_reloan,is_new_to_old;
        """.format(yesterday=yesterday)

        data = o.execute_sql(sql).open_reader(tunnel=True).to_pandas()
        big_panel_json['新客发标人数'] = data[data['is_reloan']==0]['listing_num'].sum()
        big_panel_json['老客发标人数'] = data[data['is_reloan']==1]['listing_num'].sum()
        big_panel_json['新客发标通过率'] = float(data[(data['is_reloan']==0)]['pass_num'].sum()/data[(data['is_reloan']==0)]['listing_num'].sum())
        big_panel_json['老客发标通过率'] = float(data[(data['is_reloan']==1)]['pass_num'].sum()/data[(data['is_reloan']==1)]['listing_num'].sum())
        big_panel_json['新转老发标通过率'] = float(data[(data['is_reloan']==1)&(data['is_new_to_old']==1)]['pass_num'].sum()/data[(data['is_reloan']==1)&(data['is_new_to_old']==1)]['listing_num'].sum())
        return big_panel_json

    def fill_big_panel_json(self):
        big_panel_json = {}
        for i in range(1, 21):
            method_name = f"get_info{i}"
            try:
                method = getattr(self, method_name)
                result = method()
                big_panel_json.update(result)
                logging.info(f"Successfully updated with {method_name}")
                print(i)
            except Exception as e:
                logging.error(f"Error occurred in {method_name}: {str(e)}")
        return big_panel_json


if __name__ == "__main__":
    import os
    os.chdir('/opt/workspace/pak_risk_group_drive/robot')
    info_o = Get_info_json()
    print(info_o.get_info13())
