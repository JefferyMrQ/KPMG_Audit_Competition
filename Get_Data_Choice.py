# -*- coding:utf-8 -*-
"""
 作者: QL
 日期: 2022年07月11日
"""
from EmQuantAPI import *
import pandas as pd
import re
import os
import logging

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(thread)d %(levelname)s %(module)s - %(message)s')
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)


indicator_dic = {
    "现金流量表": ["经营活动现金流入小计 经营活动现金流出小计 经营活动产生的现金流量净额 "
              "投资活动现金流入小计 投资活动现金流出小计 投资活动产生的现金流量净额 "
              "筹资活动现金流入小计 筹资活动现金流出小计 筹资活动产生的现金流量净额 "
              "期初现金及现金等价物余额 期末现金及现金等价物余额 现金及现金等价物净增加额",
              "CASHFLOWSTATEMENT_25,CASHFLOWSTATEMENT_37,CASHFLOWSTATEMENT_39,"
              "CASHFLOWSTATEMENT_48,CASHFLOWSTATEMENT_57,CASHFLOWSTATEMENT_59,"
              "CASHFLOWSTATEMENT_68,CASHFLOWSTATEMENT_75,CASHFLOWSTATEMENT_77,"
              "CASHFLOWSTATEMENT_83,CASHFLOWSTATEMENT_84,CASHFLOWSTATEMENT_82"],
    "资产负债表": ["货币资金 交易性金融资产 衍生金融资产 应收票据及应收账款 预付款项 存货 "
              "流动资产合计 长期股权投资 其他非流动金融资产 固定资产 在建工程 无形资产 商誉 "
              "非流动资产合计 资产总计 短期借款 应付票据及应付账款 合同负债 应付职工薪酬 流动负债合计 "
              "长期借款 应付债券 租赁负债 递延收益 非流动负债合计 负债合计 股东权益合计",
              "BALANCESTATEMENT_9,BALANCESTATEMENT_224,BALANCESTATEMENT_51,"
              "BALANCESTATEMENT_216,BALANCESTATEMENT_14,BALANCESTATEMENT_17,"
              "BALANCESTATEMENT_25,BALANCESTATEMENT_29,BALANCESTATEMENT_220,"
              "BALANCESTATEMENT_31,BALANCESTATEMENT_33,BALANCESTATEMENT_37,"
              "BALANCESTATEMENT_39,BALANCESTATEMENT_46,BALANCESTATEMENT_74,"
              "BALANCESTATEMENT_75,BALANCESTATEMENT_221,BALANCESTATEMENT_213,"
              "BALANCESTATEMENT_80,BALANCESTATEMENT_93,BALANCESTATEMENT_94,"
              "BALANCESTATEMENT_95,BALANCESTATEMENT_228,BALANCESTATEMENT_148,"
              "BALANCESTATEMENT_103,BALANCESTATEMENT_128,BALANCESTATEMENT_141"],
    "利润表": ["主营营业收入 主营营业支出 其他业务收入(附注) 营业成本 营业收入 净利润 营业利润 研发费用 销售费用 管理费用 财务费用",
            "MBREVENUE,MBCOST,OTHREVENUE,INCOMESTATEMENT_10,INCOMESTATEMENT_9,"
            "INCOMESTATEMENT_60,INCOMESTATEMENT_48,INCOMESTATEMENT_89,"
            "INCOMESTATEMENT_12,INCOMESTATEMENT_13,INCOMESTATEMENT_14"],
    "资本结构": ["资产负债率 流动资产/总资产 归属母公司股东的权益/全部投入资本",
             "LIBILITYTOASSET,CATOASSET,EQUITYTOTOTALCAPITAL"],
    "偿债能力": ["流动比率 速动比率 经营活动产生的现金流量净额/负债合计 资产负债率 产权比率 已获利息倍数",
             "CURRENTTATIO,QUICKTATIO,CFOTOLIBILITY,LIBILITYTOASSET,LIBILITYTOEQUITY,EBITTOINTEREST"],
    "偿债能力_行业": ["流动比率(算术平均) 速动比率(算术平均) 经营活动产生的现金流量净额/负债合计(算术平均) "
                "资产负债率(算术平均) 产权比率(算术平均) 已获利息倍数(算术平均)",
                "CRAVG,QRAVG,OCFTODEBTAVG,DEBTTOASTAVG,DEBTTOEQTAVG,EBITTOINAVG"],
    "成长能力": ["营业收入同比增长率 营业利润同比增长率 利润总额同比增长率 "
             "归属母公司股东的净利润同比增长率 净资产收益率同比增长率(摊薄) 总资产同比增长率",
             "YOYOR,YOYOP,YOYEBT,YOYPNI,YOYROELILUTED,YOYASSET"],
    "成长能力_行业": ["营业收入合计(同比增长率) 营业利润合计(同比增长率) 利润总额合计(同比增长率) "
                "归属母公司股东的净利润合计(同比增长率) 净资产收益率(整体法)(同比增长率) 资产总计(合计)",
                "YOYOR,YOYOP,YOYEBT,YOYPNI,YOYROEALL,TASSETALL"],
    "盈利能力": ["总资产净利率ROA 净资产收益率ROE(加权) 销售毛利率 净利润/营业总收入",
             "NROA,ROEWA,GPMARGIN,NITOGR"],
    "盈利能力_行业": ["总资产净利率(算术平均) 净资产收益率-加权(算术平均) 销售毛利率(算术平均) 净利润/营业总收入(算术平均)",
                "NROAAVG,ROEWAVG,GPMARGINAVG,NITOGRAVG"],
    "营运能力": ["应收账款周转率(含应收票据) 应付账款周转率(含应付票据) 存货周转率 总资产周转率 固定资产周转率 营业周期",
             "ARTURNRATIO,APTURNRATIO,INVTURNRATIO,ASSETTURNRATIO,FATURNRATIO,TURNDAYS"],
    "营运能力_行业": ["应收账款周转率(算术平均) 应付账款周转率（算术平均） 存货周转率(算术平均) "
                "总资产周转率(算术平均) 固定资产周转率(算术平均) 营业周期(算术平均)",
                "ARTURNRTOAVG,APTURNRTOAVG,INVTURNRTOAVG,ASTTURNRTOAVG,FATURNRTOAVG,TURNDAYSAVG"],
    "应收账款明细": ["应收账款—金额 应收账款—比例 应收账款—坏账准备 应收账款合计",
               "ARDETAILS_1,ARDETAILS_2,ARDETAILS_3,SUMARDETAILS"],
    "存货项目明细": ["存货明细-原材料 存货明细-在产品 存货明细-产成品 存货明细-库存商品 "
               "存货明细-周转材料 存货明细-委托加工材料 存货明细-合计",
               "INVENTORYDETAILS_1,INVENTORYDETAILS_2,INVENTORYDETAILS_3,"
               "INVENTORYDETAILS_9,INVENTORYDETAILS_11,INVENTORYDETAILS_6,INVENTORYDETAILS_8"],
    "每股指标": ["每股收益EPS(基本) 每股营业总收入",
             "EPSBASIC,GRPS"],
    "每股指标_行业": ["每股收益EPS-基本(算术平均) 每股营业总收入(算术平均)",
                "EPSBASICAVG,GRPSAVG"],
    "公司资料": ["公司中文名称 成立日期 注册资本 主营业务 省份 城市 "
             "审计机构 所属证监局 律师事务所 证券事务代表 股票代码 首发上市日期 所属中信行业(2020)",
             "COMPNAME,FOUNDDATE,REGCAPITAL,MAINBUSINESS,PROVINCE,CITY,"
             "AUDITOR,SFC,LAWFIRM,SECPRESENT,CODE,LISTDATE,CITIC2020"],
    "董监高指标": ["员工总数 员工人均薪酬 董事长薪酬 前三名董事报酬总额 前三名高管报酬总额",
              "EMPLOYEENUM,EMPLOYEEPMT,CHAIRMANPMT,FIRTHREEDIRECTORSPMT,FIRTHREESENIORPMT"],
    "本科人数": ["各学历员工人数", "LAYEREMPLOYEENUM"],
    "硕士人数": ["各学历员工人数", "LAYEREMPLOYEENUM"],
    "博士人数": ["各学历员工人数", "LAYEREMPLOYEENUM"],
    "舞弊分析数据": ["负债和股东权益合计 应收账款 应收票据 其他应收款合计 预付款项 存货 在建工程 "
               "长期待摊费用 固定资产 商誉 资产减值准备 短期借款 应付票据 应付账款 "
               "其他应付款合计 总资产周转率 净资产收益率 营业总收入 主营营业收入 "
               "营业总成本 销售毛利率 销售净利率 销售费用 管理费用 财务费用 应收账款周转率 "
               "存货周转率 固定资产周转率 研发费用 经营活动产生的现金流量净额 资产负债率 货币资金 带息负债 "
               "存货跌价准备合计 期末现金及现金等价物余额 坏账准备合计 应收账款合计 固定资产累计折旧 "
               "存货周转天数 应收账款周转天数 应付账款周转天数",
               "BALANCESTATEMENT_145,BALANCESTATEMENT_12,BALANCESTATEMENT_11,BALANCESTATEMENT_222,"
               "BALANCESTATEMENT_14,BALANCESTATEMENT_17,BALANCESTATEMENT_33,BALANCESTATEMENT_40,"
               "BALANCESTATEMENT_31,BALANCESTATEMENT_39,CASHFLOWSTATEMENT_86,BALANCESTATEMENT_75,"
               "BALANCESTATEMENT_77,BALANCESTATEMENT_78,BALANCESTATEMENT_227,ASSETTURNRATIO,ROEAVG,"
               "INCOMESTATEMENT_83,MBREVENUE,INCOMESTATEMENT_84,GPMARGIN,NPMARGIN,INCOMESTATEMENT_12,"
               "INCOMESTATEMENT_13,INCOMESTATEMENT_14,ARTURNRATIO,INVTURNRATIO,FATURNRATIO,"
               "INCOMESTATEMENT_89,CASHFLOWSTATEMENT_39,LIBILITYTOASSET,BALANCESTATEMENT_9,INTERESTLIBILITY,"
               "ASSETSIMPAIRMENT_7,CASHFLOWSTATEMENT_84,ARBADDEBTSUM_3,SUMARDETAILS,FIXEDASSETS_14,"
               "INVTURNDAYS,ARTURNDAYS,APTURNDAYS"],
    "舞弊分析数据_行业1": ["负债和股东权益合计 应收账款 预付款项 存货 固定资产 应付账款 总资产周转率 "  # 板块指标一次不能超过15个
                   "净资产收益率 营业总收入 营业总成本 销售毛利率 销售净利率 销售费用 管理费用 财务费用",
                   "TASSETAVG,ACCRECAVG,ADPAYAVG,INVENTORAVG,FIXEDASSAVG,ACCPAYAVG,ASTTURNRTOAVG,"
                   "ROEAVG,SECTOPREAVG,TOTALOPEXAVG,GPMARGINAVG,NPMARGINAVG,SALEEXPAVG,MANAEXPAVG,FINEXPAVG"],
    "舞弊分析数据_行业2": ["应收账款周转率 存货周转率 固定资产周转率 经营活动产生的现金流量净额 资产负债率 货币资金 存货周转天数 应收账款周转天数 应付账款周转天数",
                   "ARTURNRTOAVG,INVTURNRTOAVG,FATURNRTOAVG,NETOPCFAVG,DEBTTOASTAVG,CUCASHAVG,INVTURNDAYSAVG,ARTURNDAYSAVG,APTURNDAYSAVG"]
}


def connect_server():
    # 连接服务器
    try:
        loginresult = c.start()
        logger.info(loginresult)
        if re.search("(?<=ErrorMsg=).+?(?=,)", str(loginresult)).group() != 'success':
            raise Exception("登陆错误!")
        else:
            logger.info("登陆成功！")
    except Exception as e:
        logger.error(e)
        os._exit(0)


def get_data(ticker: str, start_year: int, period: int) -> dict:
    """
    :param ticker: 沪深股票代码， 例如：688981.SH
    :param start_year: 起始年份
    :param period:周期跨度（需要多少年的年报数据）
    :return:包含三张报表和财务分析的字典
    """
    try:
        connect_server()
        data_dic = {}  # 存放dataframe的字典

        # 生成时间列表
        start = str(start_year - 4) + '1231'  # 报表起始时间(-5是为了算excel 分析中部分指标的判断)
        date_range = pd.date_range(start=start, periods=period + 4, freq='Y')
        date_lst = [i.strftime('%Y-%m-%d') for i in date_range]

        for key, indicators in indicator_dic.items():
            if key.find('行业') > 0:
                ticker_revised = get_ind_code(ticker)  # 获得行业指标代码
            else:
                ticker_revised = ticker

            indicator = indicators[1]  # 指标参数
            data_lst = []  # 存放临时数据的列表

            for i in date_lst:

                # 参数选项设置
                if key == '现金流量表' or key == '资产负债表' or key == '利润表' or key == '应收账款明细':
                    options = "ispandas=1,rowindex=2,ReportDate=" + i + ",type=1"  # 对三大报表而言，有type=1参数
                elif key == '存货项目明细':
                    options = "ispandas=1,rowindex=2,ReportDate=" + i + ",type=1,DataType=1"
                elif key == '偿债能力':
                    options = "ispandas=1,rowindex=2,ReportDate=" + i + ",DataAdjustType=1"
                elif key.find('行业') > 0:
                    if key.find('成长能力') > 0:
                        options = "ispandas=1,rowindex=2,ReportDate=" + i + ",IsHistory=0,type=1"
                    elif key.find('舞弊分析数据'):
                        options = "ispandas=1,rowindex=2,ReportDate=" + i + ",IsHistory=0,type=1"
                    else:
                        options = "ispandas=1,rowindex=2,ReportDate=" + i + ",IsHistory=0"
                elif key == '公司资料':
                    options = "ispandas=1,rowindex=2,ReportDate=" + i + ",ClassiFication=4"
                elif key == '董监高指标':
                    options = "ispandas=1,rowindex=2" + ",Year=" + i[:4] + ",Payyear=" + i[:4]
                elif key == '本科人数':
                    options = "ispandas=1,rowindex=2,EndDate=" + i + ",AcademicLevel=3"
                elif key == '硕士人数':
                    options = "ispandas=1,rowindex=2,EndDate=" + i + ",AcademicLevel=2"
                elif key == '博士人数':
                    options = "ispandas=1,rowindex=2,EndDate=" + i + ",AcademicLevel=1"
                elif key == '舞弊分析数据':
                    options = "ispandas=1,rowindex=2,ReportDate=" + i + ",type=1,ItemsCode=11,DataType=1"
                else:
                    options = "ispandas=1,rowindex=2,ReportDate=" + i
                logger.debug(options)

                # 请求数据
                logger.info(f"开始请求{key}数据 {i}")
                if key.find('行业') > 0:
                    data = c.cses(ticker_revised, indicator, options)
                else:
                    data = c.css(ticker_revised, indicator, options)
                logger.info(f"请求{key}数据成功 {i}")
                logger.debug(f",数据为:\n{data}")
                data.index = [i]
                data_lst.append(data)

            # 对数据进行整理
            df = pd.concat([data_lst[i] for i in range(len(data_lst))])
            df.rename(columns=dict(zip(indicators[1].split(','), indicators[0].split(' '))),
                      inplace=True)
            df.index = pd.to_datetime(df.index)

            # # 把数据类型改为str，防止to_csv中数据缺失
            # for i in df.columns[1:]:
            #     df[i] = df[i].astype('str')

            # df.to_csv(".\\data\\" + ticker + "_CFS" + ".csv")
            data_dic[key] = [ticker, df]  # 将数据保存到字典
            logger.debug(f"成功将{key}存入字典!")

        save_data(data_dic)
        logger.info("成功获得数据！")
    except Exception as e:
        logger.error(e)
        logger.warning("获取数据失败！")
    finally:
        stop_server()

    return data_dic


def get_ind_code(ticker: str):
    """
    找到股票代码对应所属Choice里中信行业板块代码（2020）
    :param ticker:沪深股票代码， 例如：688981.SH（中芯国际）
    :return: Choice里中信行业板块代码（2020）， 例如：B_025002001001（动力煤）
    """
    #  数据整理（代码对应行业名称）
    level1_name = "石油石化 煤炭 有色金属 电力及公用事业 钢铁 基础化工 建筑 建材 轻工制造 机械 电力设备及新能源 国防军工 汽车 " \
                  "商贸零售 消费者服务 家电 纺织服装 医药 食品饮料 农林牧渔 银行 非银行金融 房地产 综合金融 交通运输 电子 通信 " \
                  "计算机 传媒 综合".split(' ')
    level1_code = "B_025001,B_025002,B_025003,B_025004,B_025005,B_025006,B_025007,B_025008,B_025009,B_025010," \
                  "B_025011,B_025012,B_025013,B_025014,B_025015,B_025016,B_025017,B_025018,B_025019,B_025020," \
                  "B_025021,B_025022,B_025023,B_025024,B_025025,B_025026,B_025027,B_025028,B_025029,B_025030".split(',')
    level1 = dict(zip(level1_name, level1_code))

    level2_name = "石油开采Ⅱ 石油化工 油服工程 煤炭开采洗选 煤炭化工 贵金属 工业金属 稀有金属 发电及电网 环保及公用事业 普钢 " \
                  "其他钢铁 特材 农用化工 化学纤维 化学原料 其他化学制品Ⅱ 塑料及制品 橡胶及制品 建筑施工 建筑装修Ⅱ 建筑设计及服务Ⅱ " \
                  "结构材料 装饰材料 专用材料Ⅱ 造纸Ⅱ 包装印刷 家居 文娱轻工Ⅱ 其他轻工Ⅱ 工程机械Ⅱ 专用机械 通用设备 运输设备 " \
                  "仪器仪表Ⅱ 金属制品Ⅱ 电气设备 电源设备 新能源动力系统 航空航天 兵器兵装Ⅱ 其他军工Ⅱ 乘用车Ⅱ 商用车 汽车零部件Ⅱ " \
                  "汽车销售及服务Ⅱ 摩托车及其他Ⅱ 一般零售 贸易Ⅱ 专营连锁 电商及服务Ⅱ 专业市场经营Ⅱ 旅游及休闲 酒店及餐饮 教育 " \
                  "综合服务 白色家电Ⅱ 黑色家电Ⅱ 小家电Ⅱ 照明电工及其他 厨房电器Ⅱ 纺织制造 品牌服饰 化学制药 中药生产 生物医药Ⅱ " \
                  "其他医药医疗 酒类 饮料 食品 种植业 畜牧业 林业 渔业 农产品加工Ⅱ 国有大型银行Ⅱ 全国性股份制银行Ⅱ 区域性银行 " \
                  "证券Ⅱ 保险Ⅱ 多元金融 房地产开发和运营 房地产服务 资产管理Ⅱ 多领域控股Ⅱ 新兴金融服务Ⅱ 其他综合金融Ⅱ 公路铁路 " \
                  "物流 航运港口 航空机场 半导体 元器件 光学光电 消费电子 其他电子零组件Ⅱ 电信运营Ⅱ 通信设备 增值服务Ⅱ 通讯工程服务 " \
                  "计算机设备 计算机软件 云服务 产业互联网 媒体 广告营销 文化娱乐 互联网媒体 综合Ⅱ".split(' ')
    level2_code = "B_025001001,B_025001002,B_025001003,B_025002001,B_025002002,B_025003001,B_025003002," \
                  "B_025003003,B_025004001,B_025004002,B_025005001,B_025005002,B_025005003,B_025006001,B_025006002," \
                  "B_025006003,B_025006004,B_025006005,B_025006006,B_025007001,B_025007002,B_025007003,B_025008001," \
                  "B_025008002,B_025008003,B_025009001,B_025009002,B_025009003,B_025009004,B_025009005,B_025010001," \
                  "B_025010002,B_025010003,B_025010004,B_025010005,B_025010006,B_025011001,B_025011002,B_025011003," \
                  "B_025012001,B_025012002,B_025012003,B_025013001,B_025013002,B_025013003,B_025013004,B_025013005," \
                  "B_025014001,B_025014002,B_025014003,B_025014004,B_025014005,B_025015001,B_025015002,B_025015003," \
                  "B_025015004,B_025016001,B_025016002,B_025016003,B_025016004,B_025016005,B_025017001,B_025017002," \
                  "B_025018001,B_025018002,B_025018003,B_025018004,B_025019001,B_025019002,B_025019003,B_025020001," \
                  "B_025020002,B_025020003,B_025020004,B_025020005,B_025021001,B_025021002,B_025021003,B_025022001," \
                  "B_025022002,B_025022003,B_025023001,B_025023002,B_025024001,B_025024002,B_025024003,B_025024004," \
                  "B_025025001,B_025025002,B_025025003,B_025025004,B_025026001,B_025026002,B_025026003,B_025026004," \
                  "B_025026005,B_025027001,B_025027002,B_025027003,B_025027004,B_025028001,B_025028002,B_025028003," \
                  "B_025028004,B_025029001,B_025029002,B_025029003,B_025029004,B_025030001".split(',')
    level2 = dict(zip(level2_name, level2_code))

    level3_name = "石油开采Ⅲ 炼油 油品销售及仓储 其他石化 油田服务 工程服务 动力煤 无烟煤 炼焦煤 焦炭 其他煤化工 黄金 铜 铅锌 " \
                  "铝 稀土及磁性材料 镍钴锡锑 钨 锂 其他稀有金属 火电 水电 其他发电 电网 燃气 供热或其他 环保及水务 长材 板材 铁矿石 " \
                  "贸易流通 钢铁耗材 特钢 氮肥 钾肥 复合肥 农药 磷肥及磷化工 涤纶 氨纶 粘胶 绵纶 碳纤维 纯碱 氯碱 无机盐 其他化学原料 " \
                  "钛白粉 日用化学品 民爆用品 涂料油墨颜料 印染化学品 其他化学制品Ⅲ 食品及饲料添加剂 电子化学品 锂电化学品 氟化工 有机硅 " \
                  "聚氨酯 橡胶助剂 改性塑料 合成树脂 膜材料 其他塑料制品 轮胎 橡胶制品 房建建设 基建建设 园林工程 专业工程及其他 建筑装修Ⅲ " \
                  "建筑设计及服务Ⅲ 水泥 玻璃 其他结构材料 陶瓷 其他装饰材料 玻璃纤维 其他专用材料 造纸Ⅲ 印刷 纸包装 金属包装 其他包装 " \
                  "家具 其他家居 文娱轻工Ⅲ 其他轻工Ⅲ 工程机械Ⅲ 叉车 电梯 高空作业车 矿山冶金机械 纺织服装机械 其他专用机械 油气装备 " \
                  "核电设备 光伏设备 3C设备 锂电设备 锅炉设备 机床设备 起重运输设备 基础件 其他通用机械 工业机器人及工控系统 服务机器人 " \
                  "塑料加工机械 激光加工设备 铁路交通设备 船舶制造 其他运输设备 仪器仪表Ⅲ 金属制品Ⅲ 输变电设备 配电设备 电力电子及自动化 " \
                  "电机 风电 核电 太阳能 储能 综合能源设备 锂电池 燃料电池 车用电机电控 电池综合服务 航空军工 航天军工 兵器兵装Ⅲ " \
                  "其他军工Ⅲ 乘用车Ⅲ 卡车 客车 专用汽车 汽车零部件Ⅲ 汽车销售及服务Ⅲ 摩托车及其他Ⅲ 百货 超市及便利店 综合业态 贸易Ⅲ " \
                  "家电3C 珠宝首饰及钟表 医疗美容 其他连锁 电商及服务Ⅲ 专业市场经营Ⅲ 景区 旅游服务 旅游零售 休闲综合 酒店 餐饮 早幼教 " \
                  "K12基础教育 K12培训 高等及职业教育 教育信息化及在线教育 教育综合 人力资源服务 服务综合 白色家电Ⅲ 黑色家电Ⅲ 小家电Ⅲ " \
                  "照明电工 其他家电 厨房电器Ⅲ 棉纺制品 非棉纺织品 印染 其他纺织 中高端成人品牌服饰 大众成人品牌服饰 体育及户外品牌 家纺 " \
                  "儿童品牌 功能性服饰 其他时尚品 化学原料药 化学制剂 中药饮片 中成药 生物医药Ⅲ 医药流通 医疗器械 医疗服务 白酒 啤酒 " \
                  "其他酒 非乳饮料 乳制品 肉制品 调味品 其他食品 休闲食品 速冻食品 种业 种植 饲料加工 动物疫苗及兽药 畜牧养殖 林木及加工 " \
                  "水产养殖 水产捕捞 水产品加工 农产品加工Ⅲ 宠物食品 国有大型银行Ⅲ 全国性股份制银行Ⅲ 城商行 农商行 证券Ⅲ 保险Ⅲ 信托 " \
                  "其他非银金融 租赁 金融交易所及数据 住宅物业开发 非住宅物业开发和运营 园区综合开发 物业经纪服务 物业管理服务 资产管理Ⅲ " \
                  "多领域控股Ⅲ 网络信贷 数字金融服务 其他新兴金融服务 其他综合金融Ⅲ 公路 铁路 公交 物流综合 快递 航运 港口 航空 机场 " \
                  "集成电路 分立器件 半导体材料 半导体设备 PCB 被动元件 安防 LED 面板 显示零组 消费电子组件 消费电子设备 其他电子零组件Ⅲ " \
                  "电信运营Ⅲ 动力设备 其他通信设备 通信终端及配件 网络接配及塔设 系统设备 线缆 增值服务Ⅲ 网络覆盖优化与运维 " \
                  "网络规划设计和工程施工 通用计算机设备 专用计算机设备 基础软件及管理办公软件 行业应用软件 新兴计算机软件 云基础设施服务 " \
                  "云平台服务 云软件服务 咨询实施及其他服务 产业互联网信息服务 产业互联网平台服务 产业互联网综合服务 出版 广播电视 " \
                  "互联网广告营销 其他广告营销 影视 动漫 游戏 其他文化娱乐 信息搜索与聚合 社交与互动媒体 互联网影视音频 综合Ⅲ".split(' ')
    level3_code = "B_025001001001,B_025001002001,B_025001002002,B_025001002003,B_025001003001,B_025001003002," \
                  "B_025002001001,B_025002001002,B_025002001003,B_025002002001,B_025002002002,B_025003001001," \
                  "B_025003002001,B_025003002002,B_025003002003,B_025003003001,B_025003003002,B_025003003003," \
                  "B_025003003004,B_025003003005,B_025004001001,B_025004001002,B_025004001003,B_025004001004," \
                  "B_025004002001,B_025004002002,B_025004002003,B_025005001001,B_025005001002,B_025005002001," \
                  "B_025005002002,B_025005002003,B_025005003001,B_025006001001,B_025006001002,B_025006001003," \
                  "B_025006001004,B_025006001005,B_025006002001,B_025006002002,B_025006002003,B_025006002004," \
                  "B_025006002005,B_025006003001,B_025006003002,B_025006003003,B_025006003004,B_025006003005," \
                  "B_025006004001,B_025006004002,B_025006004003,B_025006004004,B_025006004005,B_025006004006," \
                  "B_025006004007,B_025006004008,B_025006004009,B_025006004010,B_025006004011,B_025006004012," \
                  "B_025006005001,B_025006005002,B_025006005003,B_025006005004,B_025006006001,B_025006006002," \
                  "B_025007001001,B_025007001002,B_025007001003,B_025007001004,B_025007002001,B_025007003001," \
                  "B_025008001001,B_025008001002,B_025008001003,B_025008002001,B_025008002002,B_025008003001," \
                  "B_025008003002,B_025009001001,B_025009002001,B_025009002002,B_025009002003,B_025009002004," \
                  "B_025009003001,B_025009003002,B_025009004001,B_025009005001,B_025010001001,B_025010001002," \
                  "B_025010001003,B_025010001004,B_025010002001,B_025010002002,B_025010002003,B_025010002004," \
                  "B_025010002005,B_025010002006,B_025010002007,B_025010002008,B_025010003001,B_025010003002," \
                  "B_025010003003,B_025010003004,B_025010003005,B_025010003006,B_025010003007,B_025010003008," \
                  "B_025010003009,B_025010004001,B_025010004002,B_025010004003,B_025010005001,B_025010006001," \
                  "B_025011001001,B_025011001002,B_025011001003,B_025011001004,B_025011002001,B_025011002002," \
                  "B_025011002003,B_025011002004,B_025011002005,B_025011003001,B_025011003002,B_025011003003," \
                  "B_025011003004,B_025012001001,B_025012001002,B_025012002001,B_025012003001,B_025013001001," \
                  "B_025013002001,B_025013002002,B_025013002003,B_025013003001,B_025013004001,B_025013005001," \
                  "B_025014001001,B_025014001002,B_025014001003,B_025014002001,B_025014003001,B_025014003002," \
                  "B_025014003003,B_025014003004,B_025014004001,B_025014005001,B_025015001001,B_025015001002," \
                  "B_025015001003,B_025015001004,B_025015002001,B_025015002002,B_025015003001,B_025015003002," \
                  "B_025015003003,B_025015003004,B_025015003005,B_025015003006,B_025015004001,B_025015004002," \
                  "B_025016001001,B_025016002001,B_025016003001,B_025016004001,B_025016004002,B_025016005001," \
                  "B_025017001001,B_025017001002,B_025017001003,B_025017001004,B_025017002001,B_025017002002," \
                  "B_025017002003,B_025017002004,B_025017002005,B_025017002006,B_025017002007,B_025018001001," \
                  "B_025018001002,B_025018002001,B_025018002002,B_025018003001,B_025018004001,B_025018004002," \
                  "B_025018004003,B_025019001001,B_025019001002,B_025019001003,B_025019002001,B_025019002002," \
                  "B_025019003001,B_025019003002,B_025019003003,B_025019003004,B_025019003005,B_025020001001," \
                  "B_025020001002,B_025020002001,B_025020002002,B_025020002003,B_025020003001,B_025020004001," \
                  "B_025020004002,B_025020004003,B_025020005001,B_025020005002,B_025021001001,B_025021002001," \
                  "B_025021003001,B_025021003002,B_025022001001,B_025022002001,B_025022003001,B_025022003002," \
                  "B_025022003003,B_025022003004,B_025023001001,B_025023001002,B_025023001003,B_025023002001," \
                  "B_025023002002,B_025024001001,B_025024002001,B_025024003001,B_025024003002,B_025024003003," \
                  "B_025024004001,B_025025001001,B_025025001002,B_025025001003,B_025025002001,B_025025002002," \
                  "B_025025003001,B_025025003002,B_025025004001,B_025025004002,B_025026001001,B_025026001002," \
                  "B_025026001003,B_025026001004,B_025026002001,B_025026002002,B_025026003001,B_025026003002," \
                  "B_025026003003,B_025026003004,B_025026004001,B_025026004002,B_025026005001,B_025027001001," \
                  "B_025027002001,B_025027002002,B_025027002003,B_025027002004,B_025027002005,B_025027002006," \
                  "B_025027003001,B_025027004001,B_025027004002,B_025028001001,B_025028001002,B_025028002001," \
                  "B_025028002002,B_025028002003,B_025028003001,B_025028003002,B_025028003003,B_025028003004," \
                  "B_025028004001,B_025028004002,B_025028004003,B_025029001001,B_025029001002,B_025029002001," \
                  "B_025029002002,B_025029003001,B_025029003002,B_025029003003,B_025029003004,B_025029004001," \
                  "B_025029004002,B_025029004003,B_025030001001".split(',')
    level3 = dict(zip(level3_name, level3_code))

    #  整理成DataFrame
    df1 = pd.DataFrame.from_dict(data=level1, orient='index', columns=['CODE'])
    df1['LEVEL'] = 1
    df1.reset_index(inplace=True)

    df2 = pd.DataFrame.from_dict(data=level2, orient='index', columns=['CODE'])
    df2['LEVEL'] = 2
    df2.reset_index(inplace=True)

    df3 = pd.DataFrame.from_dict(data=level3, orient='index', columns=['CODE'])
    df3['LEVEL'] = 3
    df3.reset_index(inplace=True)

    df = pd.concat([df1, df2, df3], ignore_index=True)
    df.rename(columns={'index': 'NAME'}, inplace=True)

    lv3_name = c.css(ticker, "CITIC2020", "ClassiFication=3").Data[ticker][0]
    code = df[df['NAME'] == lv3_name]['CODE'].values[0]

    return code


def save_data(data_dic: dict):
    """
    :param data_dic: 含有[ticker, dataframe]列表数据的字典
    :return: 无
    """
    for key, data in data_dic.items():
        data[1].to_csv(".\\data\\" + data[0] + "_" + key + ".csv")
        logger.debug(f"成功保存{key}数据")


def stop_server():
    try:
        loginresult = c.stop()
        logger.info(loginresult)
        if re.search("(?<=ErrorMsg=).+?(?=,)", str(loginresult)).group() != 'success':
            raise Exception("登陆错误!")
        else:
            logger.info("登出成功！")
    except Exception as e:
        logger.error(e)
        os._exit(0)


if __name__ == '__main__':
    # ticker = str(input("请输入（上证或深证）股票代码, 例如 688981.SH: "))
    # start_year = int(input("请输入起始年份: "))
    # period = int(input("请输入年份跨度: "))

    # ticker, start_year, period = '000002.SZ', 2017, 5
    # data = get_data(ticker, start_year, period)
    # logger.info('成功获得数据!')

    # ticker_lst = ['600248.SZ', '601369.SH', '688596.SH', '000540.SZ', '002499.SZ',
    #               '300757.SZ', '300345.SZ', '300133.SZ', '002288.SZ', '300588.SZ',
    #               '000676.SZ', '300273.SZ', '300485.SZ', '002089.SZ', '002647.SZ',
    #               '002798.SZ', '003012.SZ', '601137.SH', '300336.SZ', '600545.SH',
    #               '002485.SZ', '300356.SZ', '002569.SZ', '000820.SZ', '600310.SH',
    #               '600666.SH', '600712.SH', '000687.SH', '300173.SZ', '000525.SZ',
    #               '000663.SZ', '002164.SZ', '603613.SH', '300518.SZ', '002127.SZ',
    #               '002024.SZ', '002280.SZ', '300792.SZ', '000002.SZ', '000006.SZ',
    #               '000007.SZ', '000009.SZ', '000010.SZ', '000011.SZ', '000012.SZ',
    #               '000014.SZ', '000016.SZ', '000017.SZ', '000019.SZ', '000020.SZ']

    # ticker_lst = [z + '.SH' if z[:1] == '6' else z + '.SZ' for z in s]
    # ticker_lst = [z.strip() for z in s]
    # ticker_lst = [z + '.SH' for z in ticker_lst if z[:1] == '6']

    import numpy as np

    # seed = 123

    # df = pd.read_excel(r'D:\生活\临时文件\公司文件110146312\TRD_Co.xlsx')
    # tikcer_lst_all = df.iloc[2:, 0].tolist()
    # ticker_set = set()
    # while len(ticker_set) < 1000:
    #     ticker_set.add(tikcer_lst_all[np.random.randint(0, len(tikcer_lst_all))])
    # ticker_lst = list(ticker_set)
    # pd.DataFrame(ticker_lst).to_excel('./data/trial_data/id.xlsx')

    df = pd.read_excel(r'D:\Programming\Python\Code\sophomore_year\KPMG\data\trial_data\id.xlsx', index_col=0)
    ticker_lst = df.iloc[:, 0]
    ticker_lst = [str(z).zfill(6) for z in ticker_lst]
    ticker_lst = [z + '.SH' if z[:1] == '6' else z + '.SZ' for z in ticker_lst]

    ticker_lst = ['000509.SZ']

    start_year = 2014
    period = 8

    for ticker in ticker_lst:
        get_data(ticker, start_year, period)
        logger.info(f'{ticker}: 成功获得数据!')
