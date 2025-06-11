import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
from sqlalchemy import create_engine
from scipy import optimize
import os
import time
import utility as ut
import warnings

warnings.filterwarnings('ignore')

# Export path
out_path_fig = "D:\\folder1\\folder2"
out_path = "D:\\folder3\\folder4"


db_info = {'user': 'username',
           'password': 'password',
           'host': '00.0.00.000',
           'database': 'database name'  # the company's database
           }
engine = create_engine('mysql+pymysql://%(user)s:%(password)s@%(host)s/%(database)s?charset=utf8' % db_info,
                       encoding='utf-8')


sql = '''SELECT a.s_info_windcode, b.b_info_issuer, a.trade_dt, DATEDIFF(b.b_info_paymentdate,a.trade_dt)/365 as ttm, d.b_anal_ytm, b.s_info_exchmarket, a.b_dq_volume FROM
(
SELECT s_info_windcode, b_dq_originclose, trade_dt, b_dq_volume from lh_bond_price
where b_dq_amount>=100 and trade_dt in 
(
SELECT * FROM
(
SELECT DISTINCT trade_dt FROM lh_bond_analysis_cnbd
ORDER BY trade_dt DESC
LIMIT 2
) x
)
) a
INNER JOIN
(
select s_info_windcode, b_info_issuer, b_info_paymentdate, b_info_couponrate, s_info_exchmarket from lh_bond_description
where (
industry_name1 = '短期融资券' or
industry_name1 = '中期票据' or
industry_name1 = '公司债' or
industry_name1 = '企业债'
)
and b_info_interesttype != '浮动利率' and b_info_issuetype != '私募' and b_info_guarantor is null
) b
ON a.s_info_windcode=b.s_info_windcode
INNER JOIN
(
SELECT s_info_windcode, trade_dt, b_anal_ytm FROM lh_bond_valuation
WHERE trade_dt in 
(
SELECT * FROM
(
SELECT DISTINCT trade_dt FROM lh_bond_analysis_cnbd
ORDER BY trade_dt DESC
LIMIT 2
) x
)
) d
ON a.s_info_windcode=d.s_info_windcode and a.trade_dt=d.trade_dt
LEFT JOIN
(
SELECT s_info_windcode, GROUP_CONCAT(b_info_provisiontype) as provisiontype FROM lh_bond_special_conditions
GROUP BY s_info_windcode
) c
ON a.s_info_windcode=c.s_info_windcode
WHERE FIND_IN_SET('延期条款',c.provisiontype)=0 or ISNULL(c.provisiontype)
ORDER BY a.trade_dt, a.s_info_windcode
'''  # Retrieve market trading data from the past 2 trading days that meet the specified conditions


print("Searching...\n")
trade_data = pd.read_sql_query(sql=sql, con=engine)
min_trade_dt = min(list(trade_data['trade_dt']))
trade_data = trade_data[(trade_data['trade_dt'] != min_trade_dt) | (trade_data['ttm'] >= 3)]

eval_path = "D:\\folder1\\folder2"
for eval_file_list in os.walk(eval_path):
    pass
eval_file_list = eval_file_list[2]
eval_file = max(eval_file_list)
issuer_info = pd.read_excel(os.path.join(eval_path, eval_file), sheet_name='主体隐含评级') 
issuer_info.drop(['rating_num', 'spread'], inplace=True, axis=1)
print("Lookup finished. Generating yield curve...\n")

#  Merge two sheets
trade_data = pd.merge(trade_data, issuer_info, how="inner", left_on="b_info_issuer", right_on="issuer", sort=False)
trade_data.drop(columns=['issuer'], inplace=True, axis=1)
trade_data = trade_data[(trade_data['ttm'] < 5) & (trade_data['b_anal_ytm'] > 0)]


# Latest day's yield curve
sql2 = '''SELECT work_time, term, yield FROM lh_bond_yield_data_5
WHERE work_time IN
(SELECT max(work_time) FROM lh_bond_yield_data_5
) AND term<=10 AND yield_type=1
'''
rfcurve = pd.read_sql_query(sql=sql2, con=engine)
xi = np.round(np.arange(0, 5.01, 0.01), 2)
yi = ut.HermiteInter(rfcurve['term'], rfcurve['yield'], xi)  # yield curve with 0.01 interval steps
rfcurve = pd.DataFrame(data=xi, columns=['xi'])
rfcurve['yi'] = pd.DataFrame(data=yi)
trade_data['roundttm'] = round(trade_data['ttm'], 2)
trade_data = pd.merge(trade_data, rfcurve, how='left', left_on='roundttm', right_on='xi')  # Merge to the main sheet
trade_data.drop(columns=['xi', 'roundttm'], axis=1, inplace=True)

#  Process AA+ rated bond transaction samples first
group = trade_data[(trade_data['rating_num_final'] == 3) & (~trade_data['yi'].isnull())]
group['exch_flag'] = group['s_info_exchmarket'].apply(lambda r: 1 if '银行间' in r else 0)
group.sort_values(by=['b_dq_volume', 'exch_flag', 'b_anal_ytm'], ascending=[0, 0, 1], inplace=True)
group.reset_index(drop=True, inplace=True)
group['rankvol'] = group.index.values + 1  # Ranking of transaction point weights
group['weight'] = (max(group['rankvol']) + 1 - group['rankvol'])/(sum(group['rankvol']))  # Transaction point weights
# Determine the number and range of key tenors
interval_nums = math.floor(math.sqrt(len(group)))  # Number of key tenors
num_of_points = math.floor(len(group)/interval_nums)  # Number of data points in each interval
remains = len(group) % interval_nums  # remainder
# Determine the key tenor points
group.sort_values(by=['ttm'], ascending=1, inplace=True)
group.reset_index(drop=True, inplace=True)
key_time = np.zeros(interval_nums + 1)
division = np.ones(interval_nums) * num_of_points;
added = np.hstack((np.ones(remains), np.zeros(interval_nums - remains)))
division = np.cumsum(division + added)
for i in range(0, len(division)):
    if i != len(division)-1:
        key_time[i+1] = (group.loc[division[i], 'ttm'] + group.loc[division[i] + 1, 'ttm']) * 0.5
    else:
        key_time[i+1] = 5.0


#  # Interpolation for AA+ (benchmark yield curve)
key_rate0 = np.zeros(len(key_time))  # Initial yield values at key tenor points
bnds = [[0.0, 0.01]]
bnds.extend([[0.0, 0.002]] * (len(key_time)-2))
bnds.append([0.0, 0.0])  # 'bnds' – list of upper and lower bounds
xopt = optimize.minimize(fun=ut.myfunc, x0=key_rate0, args=[group, key_time], bounds=bnds)  # Unconstrained optimization to solve for yields
key_rate_spread = xopt.x  # Today's yields at key tenor points
key_rate = rfcurve.loc[rfcurve['xi'].isin(np.round(key_time, 2)), 'yi'].values/100 + np.cumsum(key_rate_spread)

# other yield curve 
#  AAA-
shift0 = np.zeros(len(key_time))
bnds = [[(-1) * key_rate_spread[0], -0.001]]
for i in range(1, len(key_rate_spread)):
    bnds.append([(-1) * key_rate_spread[i], 0.0])
key_rate_spread_aaa_neg = ut.make_curve(trade_data, 2, key_time, key_rate_spread, shift0, bnds)
key_rate_aaa_neg = rfcurve.loc[rfcurve['xi'].isin(np.round(key_time, 2)), 'yi'].values/100 + \
                   np.cumsum(key_rate_spread_aaa_neg)

# AAA+
shift0 = np.zeros(len(key_time))
bnds = [[(-1) * key_rate_spread_aaa_neg[0], -0.001]]
for i in range(1, len(key_rate_spread_aaa_neg)):
    bnds.append([(-1) * key_rate_spread_aaa_neg[i], 0.0])
key_rate_spread_aaa_pos = ut.make_curve(trade_data, 0, key_time, key_rate_spread_aaa_neg, shift0, bnds)
key_rate_aaa_pos = rfcurve.loc[rfcurve['xi'].isin(np.round(key_time, 2)), 'yi'].values/100 + \
                   np.cumsum(key_rate_spread_aaa_pos)

# AAA
key_rate_aaa = np.sqrt(key_rate_aaa_pos * key_rate_aaa_neg)

# AA
shift0 = np.zeros(len(key_time))
bnds = [[0.001, 0.003]]
bnds.extend([[0.0, 0.0005]] * (len(key_time)-2))
bnds.append([0.0, 0.0])  # bnds list of upper and lower bounds
key_rate_spread_aa = ut.make_curve(trade_data, 4, key_time, key_rate_spread, shift0, bnds)
key_rate_aa = rfcurve.loc[rfcurve['xi'].isin(np.round(key_time, 2)), 'yi'].values/100 + \
              np.cumsum(key_rate_spread_aa)

# AA-
shift0 = np.zeros(len(key_time))
bnds = [[0.014, 0.018]]
bnds.extend([[0.0, 0.0005]] * (len(key_time)-2))
bnds.append([0.0, 0.0])  # bnds list of upper and lower bounds
key_rate_spread_aa_neg = ut.make_curve(trade_data, 5, key_time, key_rate_spread_aa, shift0, bnds)
key_rate_aa_neg = rfcurve.loc[rfcurve['xi'].isin(np.round(key_time, 2)), 'yi'].values/100 + \
              np.cumsum(key_rate_spread_aa_neg)

# A+
shift0 = 0
bnds = [[0.019, 0.021]]
key_rate_a_pos = ut.make_curve(trade_data, 6, key_time, key_rate_aa_neg, shift0, bnds)

# A
shift0 = 0
bnds = [[0.019, 0.021]]
key_rate_a = ut.make_curve(trade_data, 7, key_time, key_rate_a_pos, shift0, bnds)

# A-
shift0 = 0
bnds = [[0.019, 0.021]]
key_rate_a_neg = ut.make_curve(trade_data, 8, key_time, key_rate_a, shift0, bnds)


xt = np.arange(0, 5.01, 0.01)
yt = ut.HermiteInter(key_time, key_rate, xt)  # aa+
yt1 = ut.HermiteInter(key_time, key_rate_aaa_pos, xt)  # aaa+
yt2 = ut.HermiteInter(key_time, key_rate_aaa, xt)  # aaa
yt3 = ut.HermiteInter(key_time, key_rate_aaa_neg, xt)  # aaa-
yt4 = ut.HermiteInter(key_time, key_rate_aa, xt)  # aa
yt5 = ut.HermiteInter(key_time, key_rate_aa_neg, xt)  # aa-
yt6 = ut.HermiteInter(key_time, key_rate_a_pos, xt)  # a+
yt7 = ut.HermiteInter(key_time, key_rate_a, xt)  # a
yt8 = ut.HermiteInter(key_time, key_rate_a_neg, xt)  # a-
colors = ['springgreen', 'lightpink', 'royalblue', 'orange', 'dimgray', 'lightblue', 'plum', 'mediumaquamarine', 'mediumpurple']
Y = [yt1, yt2, yt3, yt, yt4, yt5, yt6, yt7, yt8]
# plot
print("Generating today's yield curve...\n")
for i in range(0, len(Y)):
    plt.plot(xt, Y[i], color=colors[i], linewidth=1.3)
plt.plot(xi, yi/100, color='gray', linewidth=1.1, linestyle='--')
plt.grid(linestyle='--', color='#2F4F4F')

# save the figure
out_file_fig = '到期收益曲线' + max(trade_data['trade_dt']) + '.png'
plt.savefig(os.path.join(out_path_fig, out_file_fig))

# save the figure's data
print("Saving yield curve data...\n")
df = pd.DataFrame(columns=['work_time', 'yield_type', 'term', 'yield'])
temp_list = [key_rate_aaa_pos, key_rate_aaa, key_rate_aaa_neg, key_rate, key_rate_aa, key_rate_aa_neg, key_rate_a_pos,
             key_rate_a, key_rate_a_neg]
for yield_type in range(0, len(temp_list)):
    for i in range(0, len(key_time)):
        temp = [max(trade_data['trade_dt']), yield_type, round(key_time[i], 4), round(temp_list[yield_type][i]*100, 4)]
        df = df.append(pd.DataFrame([temp], columns=['work_time', 'yield_type', 'term', 'yield']), ignore_index=True)

curve_file = '到期收益曲线数据.xlsx'
history_curve_data = pd.read_excel(os.path.join(out_path, curve_file), index_col=False)
history_curve_data['work_time'] = history_curve_data['work_time'].astype(str)  # Convert to string type to maintain consistency with today's incremental data
old_history_curve_data = ut.df_dif(history_curve_data, df, keys=['work_time', 'yield_type', 'term'])
curve_data = pd.concat([old_history_curve_data, df])
curve_data.sort_values(by=['work_time', 'yield_type', 'term'], inplace=True)
curve_data.to_excel(os.path.join(out_path, curve_file), index=False)
print("Yield curve update complete...\n")


