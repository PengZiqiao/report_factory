# coding=utf-8

from time import sleep

import numpy as np
import pandas as pd
import seaborn as sns
from matplotlib import pyplot as plt
from matplotlib.ticker import MultipleLocator
from winsun.date import Week
from winsun.func import percent
from winsun.tools import GisSpider, QUANSHI_BUHAN_LIGAO


class PyGis():
    def __init__(self):
        self.gis = GisSpider()
        sleep(3)

    def trend_gxj(self, wuye):
        """周度供销走势"""
        # 选项
        option = {
            # 周度供销
            'by': 'week',
            # 选择物业类型
            'usg': wuye,
            # 板块为全市不含溧水高淳
            'plate': QUANSHI_BUHAN_LIGAO,
            # 查询10周的数据
            'period': 10
        }
        df = self.gis.trend_gxj(**option)

        # 日期
        w = Week()
        wN = w.N
        wList = list(w.history(i) for i in range(10))
        wList.reverse()

        # x轴标签为“第n周”
        # df.index = list(f'第{i}周' for i in range(wN - 9, wN + 1))
        # x轴标签为日期
        df.index = list(f'{each[0].strftime("%m/%d")}-\n{each[1].strftime("%m/%d")}' for each in wList)
        # 系列名称
        df.columns = ['上市面积(万㎡)', '成交面积(万㎡)', '成交均价(元/㎡)']

        return df

    def plate_gxj(self, wuye):
        """当周供销价"""
        option = {
            # 周度供销
            'by': 'week',
            # 选择物业类型
            'usg': wuye,
            # 板块为全市不含溧水高淳
            'plate': QUANSHI_BUHAN_LIGAO,
        }

        df = self.gis.current_gxj(**option)
        # 删除合计行
        df = df[:-1]
        # 填充空行，修改列名
        index = ['城中', '城东', '城南', '河西', '城北', '仙林', '江宁', '浦口', '六合']
        df = pd.concat([pd.DataFrame(index=index), df], axis=1)
        df.columns = ['上市面积(万㎡)', '成交面积(万㎡)', '成交均价(元/㎡)']
        # df = df.fillna(0)
        # df.iloc[:,-1] = df.iloc[:,-1].astype('int')
        return df.reindex(index)

    def rank(self, wuye):
        """排行榜
        :param wuye: 物业类型
        :return: df组成的列表，分别为面积、套数、金额、均价排行榜
        """
        return self.gis.rank(by='week', usg=wuye, plate=QUANSHI_BUHAN_LIGAO)

    def zhuzhai_rank(self):
        dfs = self.rank(['住宅'])
        # 遍历面积、套数榜
        for i in [0, 1]:
            df = dfs[i]
            # 替换仙西为仙林
            df = df.apply(lambda x: x.replace('仙西', '仙林'))
            # 0:案名,1:推广名,2:板块,3:面积,4:套数,5:金额,6:均价
            # i + 3， i=0时为3即面积，i=1时为4即套数
            df = df.iloc[:, [1, 2, i + 3]]
            # 在最左插入一列‘排名’列
            df.insert(0, '排名', df.index)
            dfs[i] = df
        return dfs[:2]

    def rank_shuoli(self, wuye):
        # 获得面积排行榜
        df = self.gis.rank(by='week', usg=wuye, plate=QUANSHI_BUHAN_LIGAO)[0]
        # 替换仙西为仙林
        df = df.apply(lambda x: x.replace('仙西', '仙林'))
        # 找出排名前3
        text = ''
        for i in range(3):
            # 0:案名,1:推广名,2:板块,3:面积,4:套数,5:金额,6:均价
            name, plate, area, sets, price = df.iloc[i, [1, 2, 3, 4, 6]]
            # 面积除以10000后保留两位小数
            text += f'{plate}{name}（{round(area/1e4,2)}万㎡，{sets}套，{price}元/㎡）、'
        # 最后一个顿号改成句号
        return text[:-1] + '。'


class ShuoLi():
    """
    df 传入一个供销走势的DataFrame
    ss, cj, jj 分别代表上市、成交、均价
    ss_, cj_, jj_ 加下划线代表对应环比
    """

    def __init__(self, df, degree=0):
        # 计算环比只需要近2周的数据
        df = df[-2:]
        # 计算出环比
        huanbi = df.pct_change()
        # 调整
        self.df = pd.concat([df.iloc[-1:], huanbi.iloc[-1:]])
        self.df.index = ['value', 'change']
        self.df.columns = ['ss', 'cj', 'jj']

    def shangshi(self):
        value = self.df.at['value', 'ss']
        change = self.df.at['change', 'ss']
        # 上市量为0或为空
        if value == 0 or np.isnan(value):
            return '无上市'
        # 环比值为空
        elif np.isnan(change):
            return f'上市{value:.2f}万㎡'
        else:
            return f'上市{value:.2f}万㎡，环比{percent(change,0)}'

    def chengjiao(self):
        value = self.df.at['value', 'cj']
        change = self.df.at['change', 'cj']
        # 成交量为0或为空
        if value == 0 or np.isnan(value):
            return '无成交'
        # 环比值为空
        elif np.isnan(change):
            return f'成交{value:.2f}万㎡'
        else:
            return f'成交{value:.2f}万㎡，环比{percent(change,0)}'

    def junjia(self):
        value = self.df.at['value', 'jj']
        change = self.df.at['change', 'jj']

        # 没有成交即没有成交均价
        v_cj = self.df.at['value', 'cj']
        if v_cj == 0 or np.isnan(v_cj):
            return None
        # 环比值为空
        elif np.isnan(change):
            return f'成交均价{value:.0f}元/㎡'
        else:
            return f'成交均价{value:.0f}元/㎡，环比{percent(change,0)}'

    def all(self):
        ss = self.shangshi()
        cj = self.chengjiao()
        jj = self.junjia()
        return f'{ss}；{cj}，{jj}。'


def gxj_chart(df, bar_width=0.4, size=(8, 5), color=('#465E77', '#D74B4B', '#A1A598'),
              output='chart.png', xticks_style={}):
    """生成chart图片文件"""

    def show_value(ax, left, data, bottom=False, fontsize=12, integer=False, rotation=0):
        """
        显示数据标签
        :param ax: matplot的ax对象
        :param left: 左边距离，列表
        :param data: 数据， 列表
        :param bottom: 固定在底部的距离
        :param fontsize: 字体大小
        :param integer: 不显示小数
        :param rotation: 旋转角度
        """
        for x, value in zip(left, data):
            # 固定底部的距离为0.1
            if bottom:
                y = bottom
            else:
                y = value + data.max() / 100 * 2

            # NaN与0不显示
            if not np.isnan(value) and not value == 0:
                # 是否显示小数
                if integer:
                    value = f'{value:.0f}'
                else:
                    value = f'{value:.2f}'
                # 数据标签。va：垂直对齐；ha：水平对齐
                ax.text(x, y, value, va='bottom', ha='center', fontsize=fontsize, rotation=rotation)

    # 系列名称、数据条数、plt画图的left参数
    col = df.columns
    n = len(df)
    left = [
        list(range(n)),
        # 柱2在原有位置上向右移bar_width的距离
        list(x + bar_width for x in range(n)),
        # 拆线和横坐标轴向右移bar_width/2的距离， 即柱1与柱2中间
        list(x + bar_width / 2 for x in range(n))
    ]

    # seaborn 样式
    sns.set_style('ticks', rc={
        'font.family': ['SimHei'],
        'axes.linewidth': size[0] / 10,
        'ytick.major.pad': 5
    })

    sns.set_context('paper', font_scale=size[1] / 3, rc={
        "figure.figsize": size
    })

    # chart对象
    fig, ax1 = plt.subplots()

    for i in range(2):
        # 柱状图
        ax1.bar(left[i], df[col[i]], bar_width, color=color[i], label=col[i])
        # bottom 为ax1 高度（即上市量成交量中最大值）的 2%。
        bottom = df.iloc[:, [0, 1]].max().max() / 100 * 2
        show_value(ax1, left[i], df[col[i]], bottom=bottom, rotation=90)
    # 调整y1刻度，
    ax1.set_ylim(0)
    # MultipleLocator(y_max // 5)计算出平分成5段需要的整数间隔， plt.LinearLocator(5)会出现小数
    # y_max = ax1.get_ylim()[1]
    # ax1.yaxis.set_major_locator(MultipleLocator(y_max // 5))


    # 折线，再添加一个y轴
    ax2 = ax1.twinx()
    ax2.plot(left[2], df[col[2]], color=color[2], marker='D')
    show_value(ax2, left[2], df[col[2]], integer=True)
    # 调整y2刻度
    y_max = ax2.get_ylim()[1]
    ax2.set_ylim(0, y_max + 5000)
    ax2.yaxis.set_major_locator(MultipleLocator(5000))

    # 图例
    ax1.plot(np.nan, label=col[2], color=color[2], marker='D')
    ax1.legend(loc='lower center', bbox_to_anchor=(0.5, -0.2), ncol=3, fontsize='10')

    # 调整x轴位置，显示的标签改为 df.index
    ax1.set_xticks(left[2])
    ax1.set_xticklabels(df.index, **xticks_style)
    ax2.set_xticks(left[2])

    # save
    plt.savefig(output, bbox_inches='tight', dpi=300)
