import numpy as np
import pandas as pd
from data_class import MonthlySalesAnalyzer
from chart_class import PlotStackedBar, PlotStackedBarPlus, PlotHeatGrid
import matplotlib as mpl
import matplotlib.pyplot as plt
import datetime
from matplotlib.gridspec import GridSpec

D_CLIENT_MAP = {
    "京东大药房泰州连锁有限公司": "京东",
    "阿里健康大药房医药连锁有限公司": "阿里",
    "北京德开医药科技有限公司": "北京德开",
    "广东健客医药有限公司": "广东健客",
    "广东亮健药业有限公司": "广东亮键",
    "北京好药师大药房连锁有限公司": "好药师大药房",
    "四川泉源堂药业有限公司": "四川泉源堂",
    "仁和药房网国华（北京）医药有限公司": "仁和集团",
    "广州方舟医药有限公司": "广东健客",
    "广东康爱多连锁药店有限公司": "广东康爱多",
    "广东瑞美药业有限公司": "广东瑞美",
    "阿里健康大药房连锁有限公司（浙江）": "阿里",
    "广州健客药业有限公司": "广东健客",
    "北京壹壹壹商业连锁有限公司": "北京壹壹壹",
    "叮当安健医药科技（北京）有限公司": "仁和集团",
    "阿里健康大药房医药连锁有限公司（浙江）": "阿里",
    "北京凯尔康大药房有限责任公司": "北京凯尔康",
    "广东康爱多数字健康科技有限公司": "广东康爱多",
    "广东圆心药业有限公司": "广东圆心",
    "叮当智慧药房(上海)有限公司": "叮当智慧药房",
    "北京佳康医药有限责任公司": "北京佳康",
    "京东医药（北京）有限公司": "京东",
    "广东美团大药房有限公司（美团自营）": "美团",
    "广东非常大药业有限公司（百度自营）": "百度",
    "江苏苏宁大药房有限公司（苏宁自营）": "苏宁",
    "苏州纳百特大药房有限公司（平安好医生）": "平安好医生",
    "阿里健康大药房医药连锁有限公司（信立泰旗舰店）": "阿里",
}

if __name__ == "__main__":
    df = pd.read_excel("data.xlsx", sheet_name="data", engine="openpyxl")
    df.rename(columns={"标准数量": "销售盒数", "指标数量": "指标盒数"}, inplace=True)
    df["年月"] = df["年月"].apply(lambda x: str(x)[:4] + "-" + str(x)[4:])
    # df["年月"] = df["年月"].apply(
    #     lambda x: datetime.datetime(year=int(str(x)[:4]), month=int(str(x)[4:]), day=1)
    # )  # 修改日期列格式
    df["销售金额（万元）"] = df["销售金额（元）"] / 10000
    df["销售盒数（万盒）"] = df["销售盒数"] / 10000
    df["指标金额（万元）"] = df["指标金额（元）"] / 10000
    df["指标盒数（万盒）"] = df["指标盒数"] / 10000
    
    df["主要产品"] = df["品名"].apply(lambda x: "其他" if x not in ["信立坦", "泰嘉"] else x)
    df["客户"] = df["客户名称"].map(D_CLIENT_MAP)
    print(df["客户"].unique())
    df["主要客户"] = df["客户"].apply(lambda x: "其他" if x not in ["京东", "阿里"] else x)

    a = MonthlySalesAnalyzer(data=df, name="电商销售", date_column="年月")
    print(a.get_kpi())

    # plot_data = [
    #     a.get_pivot(
    #         index="品名",
    #         columns="客户",
    #         values="销售金额（万元）",
    #         aggfunc=sum,
    #         filter={a.date_column: a.daterange_ytd},
    #     ),
    #     a.get_pivot(
    #         index="客户",
    #         columns="品名",
    #         values="销售金额（万元）",
    #         aggfunc=sum,
    #         filter={a.date_column: a.daterange_ytd},
    #     ),
    # ]
    # print(plot_data)
    # fmt = [".0f", ".0f"]
    # title = "电商 客户 versus 产品 YTD销售表现"
    # gs = GridSpec(1, 2, width_ratios=[2, 1], wspace=0)
    # gs_title = ["客户 versus 产品", "产品 versus 客户"]

    # f = plt.figure(
    #     FigureClass=PlotStackedBar,
    #     width=16,
    #     height=5,
    #     savepath=a.savepath,
    #     fmt=fmt,
    #     data=plot_data,
    #     fontsize=10,
    #     gs=gs,
    #     style={
    #         "title": title,
    #         "gs_title": gs_title,
    #         "xlabel_rotation": 90,
    #         "ylabel": "销售金额（万元）",
    #         "remove_yticks": True,
    #     },
    # )

    # f.plot(threshold=20, show_total_label=True)

    # plot_data = [
    #     a.get_pivot(
    #         index="品名",
    #         columns="客户",
    #         values="销售金额（万元）",
    #         aggfunc=sum,
    #         filter={a.date_column: a.daterange_ytd},
    #     ).apply(lambda x: x / x.sum()),
    #     a.get_pivot(
    #         index="客户",
    #         columns="品名",
    #         values="销售金额（万元）",
    #         aggfunc=sum,
    #         filter={a.date_column: a.daterange_ytd},
    #     ).apply(lambda x: x / x.sum()),
    # ]
    # print(plot_data)
    # fmt = [".0%", ".0%"]
    # title = "电商 客户 versus 产品 YTD销售贡献"
    # gs = GridSpec(1, 2, width_ratios=[2, 1], wspace=0)
    # gs_title = ["客户 versus 产品", "产品 versus 客户"]

    # f = plt.figure(
    #     FigureClass=PlotStackedBar,
    #     width=16,
    #     height=5,
    #     savepath=a.savepath,
    #     fmt=fmt,
    #     data=plot_data,
    #     fontsize=9,
    #     gs=gs,
    #     style={
    #         "title": title,
    #         "gs_title": gs_title,
    #         "xlabel_rotation": 90,
    #         "remove_yticks": True,
    #         "ylabel":"销售金额贡献",
    #     },
    # )

    # f.plot(threshold=0.02)

    query_str = "主要客户=='其他'"
    # query_str = "ilevel_0 in ilevel_0"

    column = "主要客户"
    monthly_sales = a.get_pivot(
        index="年月", values="销售金额（万元）", columns=column, aggfunc=sum, query_str=query_str
    )
    fmt = [".0f"]

    title = "电商渠道泰嘉%s贡献月度趋势 - 金额 - %s" % (column, a.date.strftime("%Y-%m"))
    f = plt.figure(
        FigureClass=PlotStackedBar,
        width=16,
        height=4,
        savepath=a.savepath,
        data=monthly_sales.transpose(),
        fmt=fmt,
        fontsize=11,
        style=dict(
            title=title,
            remove_yticks=True,
            xlabel_rotation=90,
        ),
    )

    f.plot(show_total_label=True, show_label=True, threshold=20)

    monthly_sales = a.get_pivot(
        index="年月", values="销售金额（万元）", columns=column, aggfunc=sum, query_str=query_str
    )
    monthly_sales = monthly_sales.apply(lambda x: x / x.sum(), axis=1)
    fmt = [".0%"]

    f = plt.figure(
        FigureClass=PlotStackedBar,
        width=16,
        height=4,
        savepath=a.savepath,
        data=monthly_sales.transpose(),
        fmt=fmt,
        fontsize=11,
        style=dict(
            title=None,
            remove_yticks=True,
            xlabel_rotation=90,
        ),
    )

    f.plot(show_total_label=False, show_label=True, threshold=0.03)

    monthly_sales = a.get_pivot(
        index="年月", values="销售金额（万元）", aggfunc=sum, query_str=query_str
    )
    monthly_target = a.get_pivot(
        index="年月", values="指标金额（万元）", aggfunc=sum, query_str=query_str
    )
    monthly_ach = monthly_sales["销售金额（万元）"] / monthly_target["指标金额（万元）"]
    monthly_ach = monthly_ach.to_frame(name="达成率")

    fmt = [".0f"]
    fmt_line = [".0%"]
    title = "电商渠道 其他平台 销售及达成月度趋势 - 金额 - %s" % a.date.strftime("%Y-%m")
    f = plt.figure(
        FigureClass=PlotStackedBar,
        width=16,
        height=3,
        savepath=a.savepath,
        data=monthly_sales.transpose(),
        fmt=fmt,
        data_line=monthly_ach.transpose(),
        fmt_line=fmt_line,
        fontsize=10,
        style={
            "title": title,
            "remove_yticks": True,
            "xlabel_rotation": 90,
        },
        table_data=a.get_kpi(query_str),
    )

    f.plot(show_total_label=True, show_label=False)

    monthly_sales = a.get_pivot(
        index="年月", values="销售盒数（万盒）", aggfunc=sum, query_str=query_str
    )
    monthly_target = a.get_pivot(
        index="年月", values="指标盒数（万盒）", aggfunc=sum, query_str=query_str
    )
    monthly_ach = monthly_sales["销售盒数（万盒）"] / monthly_target["指标盒数（万盒）"]
    monthly_ach = monthly_ach.to_frame(name="达成率")

    fmt = [".0f"]
    fmt_line = [".0%"]
    title = "电商渠道 其他平台 销售及达成月度趋势 - 标准盒数 - %s" % a.date.strftime("%Y-%m")
    f = plt.figure(
        FigureClass=PlotStackedBar,
        width=16,
        height=3,
        savepath=a.savepath,
        data=monthly_sales.transpose(),
        fmt=fmt,
        data_line=monthly_ach.transpose(),
        fmt_line=fmt_line,
        fontsize=10,
        style={
            "title": title,
            "remove_yticks": True,
            "xlabel_rotation": 90,
        },
        table_data=a.get_kpi(query_str, unit="盒数"),
    )

    f.plot(show_total_label=True, show_label=False)
    # plot_data = [
    #     a.get_pivot(index="品名", columns="YTD", values="销售金额（万元）", aggfunc=sum),
    #     a.get_pivot(index="主要客户", columns="YTD", values="销售金额（万元）", aggfunc=sum),
    # ]

    # title = "电商销售金额分布及同比变化"
    # gs = GridSpec(1, 2, wspace=0.2)
    # gs_title = ["产品分布", "客户分布"]
    # fmt = [".0f", ".0f"]

    # f = plt.figure(
    #     width=16,
    #     height=5,
    #     FigureClass=PlotStackedBarPlus,
    #     gs=gs,
    #     fmt=fmt,
    #     savepath=a.savepath,
    #     data=plot_data,
    #     fontsize=11,
    #     style={
    #         "title": title,
    #         "gs_title": gs_title,
    #         "ylabel": "销售金额（万元）",
    #         "hide_top_right_spines": True,
    #     },
    # )

    # f.plot(ylabel="销售金额（万元）")

    plot_data = []
    for brand in ["信立坦", "泰嘉"]:
        monthly_sales = a.get_pivot(
            index="年月", values="销售金额（万元）", aggfunc=sum, query_str=f"品名=='{brand}'"
        )
        monthly_volume = a.get_pivot(
            index="年月", values="销售盒数（万盒）", aggfunc=sum, query_str=f"品名=='{brand}'"
        )
        monthly_price = monthly_sales["销售金额（万元）"] / monthly_volume["销售盒数（万盒）"]
        monthly_price = monthly_price.to_frame(name="单价")
        print(monthly_price)
        plot_data = [
            monthly_sales.transpose(),
            monthly_volume.transpose(),
            monthly_price.transpose(),
        ]

        title = "电商渠道 %s 金额/盒数及单价月度趋势 - %s" % (brand, a.date.strftime("%Y-%m"))
        gs = GridSpec(3, 1, wspace=0.1)
        fmt = [".0f", ".1f", ".1f"]

        f = plt.figure(
            width=16,
            height=5,
            FigureClass=PlotStackedBar,
            gs=gs,
            fmt=fmt,
            savepath=a.savepath,
            data=plot_data,
            fontsize=10,
            style={
                "title": title,
                "hide_top_right_spines": True,
                "xlabel_rotation": 90,
                "last_xticks_only": True,
            },
        )

        f.plot(show_label=False, show_total_label=True)
