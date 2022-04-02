import numpy as np
import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta


FORMAT_ABS = "{:,.1f}"
FORMAT_DIFF = "{:+,.1f}"
FORMAT_SHARE = "{:.1%}"
FORMAT_GR = "{:+.1%}"
FORMAT_CURRENCY = "¥{:,.0f}"

D_SORTER = {
    "年月": pd.date_range("2020-01", "2022-02", freq="MS").strftime("%Y-%m").tolist(),
    "YTD": ["去年同期", "YTD销售"],
    "主要客户": ["京东", "阿里", "北京德开", "广东健客", "四川泉源堂", "仁和集团", "广东亮健", "其他"],
}


def merge_dicts(d1: dict, d2: dict) -> dict:
    if d1 == None:
        return d2
    elif d2 == None:
        return d1
    else:
        return {**d1, **d2}


class DfAnalyzer(pd.DataFrame):
    @property
    def _constructor(self):
        return DfAnalyzer._internal_constructor(self.__class__)

    class _internal_constructor(object):
        def __init__(self, cls):
            self.cls = cls

        def __call__(self, *args, **kwargs):
            kwargs["name"] = None
            kwargs["date_column"] = None
            return self.cls(*args, **kwargs)

        def _from_axes(self, *args, **kwargs):
            return self.cls._from_axes(*args, **kwargs)

    def __init__(
        self,
        data: pd.DataFrame,
        name: str,
        date_column: str,
        savepath: str = "./plots/",
        index=None,
        columns=None,
        dtype=None,
        copy=True,
    ):
        super(DfAnalyzer, self).__init__(
            data=data, index=index, columns=columns, dtype=dtype, copy=copy
        )
        self.data = data
        self.name = name
        self.date_column = date_column
        self.savepath = savepath

    # # 根据列名和列值做数据筛选
    # def filtered(self, filter: dict = None):
    #     if filter is not None:
    #         # https: // stackoverflow.com / questions / 38137821 / filter - dataframe - using - dictionary
    #         return self[self.isin(filter).sum(1) == len(filter.keys())]
    #     else:
    #         return self

    # 透视
    def get_pivot(
        self,
        index: str = None,
        columns: str = None,
        values: str = None,
        aggfunc: str = None,
        query_str: str = "ilevel_0 in ilevel_0",  # 默认query语句能返回df总体
        perc: bool = False,
        sort_values: bool = True,
        dropna: bool = True,
        fillna: bool = True,
        **kwargs,
    ) -> pd.DataFrame:

        pivoted = pd.pivot_table(
            self.query(query_str),
            values=values,
            index=index,
            columns=columns,
            aggfunc=aggfunc,
        )
        # pivot table对象转为默认df
        pivoted = pd.DataFrame(pivoted.to_records())
        try:
            pivoted.set_index(index, inplace=True)
        except KeyError:  # 当index=None时，捕捉错误并set_index("index"字符串)
            pivoted.set_index("index", inplace=True)

        if sort_values is True:
            s = pivoted.sum(axis=1).sort_values(ascending=False)
            pivoted = pivoted.loc[s.index, :]  # 行按照汇总总和大小排序
            s = pivoted.sum(axis=0).sort_values(ascending=False)
            pivoted = pivoted.loc[:, s.index]  # 列按照汇总总和大小排序

        if columns in D_SORTER:
            pivoted = pivoted.reindex(columns=D_SORTER[columns])

        if index in D_SORTER:
            pivoted = pivoted.reindex(D_SORTER[index])  # 对于部分变量有固定排序

        # 删除NA或替换NA为0
        if dropna is True:
            pivoted = pivoted.dropna(how="all")
            pivoted = pivoted.dropna(axis=1, how="all")
        else:
            if fillna is True:
                pivoted = pivoted.fillna(0)

        if perc is True:
            pivoted = pivoted.div(pivoted.sum(axis=1), axis=0)  # 计算行汇总的百分比

        if "head" in kwargs:  # 只取top n items
            pivoted = pivoted.head(kwargs["head"])

        if "tail" in kwargs:  # 只取bottom n items
            pivoted = pivoted.tail(kwargs["tail"])

        # if index == self.date_column:
        #     pivoted.index = pd.to_datetime(pivoted.index, format='%Y-%m')

        # if columns == self.date_column:
        #     pivoted.columns = pd.to_datetime(pivoted.columns, format='%Y-%m')

        return pivoted


class MonthlySalesAnalyzer(DfAnalyzer):
    @property
    def date(self):
        date_str = self[self.date_column].max()
        return datetime.datetime(
            year=int(date_str[:4]), month=int(date_str[5:]), day=1
        )  # 最新月份

    @property
    def date_ya(self):
        return self.date.replace(year=self.date.year - 1)  # 同比月份

    @property
    def date_year_begin(self):
        return self.date.replace(month=1)  # 本年度开头

    @property
    def date_ya_begin(self):
        return self.date_ya.replace(month=1)  # 去年开头

    @property
    def date_mat_begin(self):
        return self.date + relativedelta(months=-11)  # 滚动年开头

    @property
    def date_matya_begin(self):
        return self.date_ya + relativedelta(months=-11)  # 滚动年同比开头

    @property
    def date_mqt_begin(self):
        return self.date + relativedelta(months=-2)  # 滚动季开头

    @property
    def date_mqtya_begin(self):
        return self.date_ya + relativedelta(months=-2)  # 滚动季同比开头

    @property
    def date_mqtqa_begin(self):
        return self.date + relativedelta(months=-5)  # 滚动季环比开头

    @property
    def date_mqtqa_end(self):
        return self.date + relativedelta(months=-3)  # 滚动季环比结尾

    @property
    def daterange(self):
        return (
            pd.date_range(
                self[self.date_column].min(), self[self.date_column].max(), freq="MS"
            )
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_ytd(self):
        return (
            pd.date_range(self.date_year_begin, self.date, freq="MS")
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_ytdya(self):
        return (
            pd.date_range(self.date_ya_begin, self.date_ya, freq="MS")
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_mat(self):
        return (
            pd.date_range(self.date_mat_begin, self.date, freq="MS")
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_matya(self):
        return (
            pd.date_range(self.date_matya_begin, self.date_ya, freq="MS")
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_mqt(self):
        return (
            pd.date_range(self.date_mqt_begin, self.date, freq="MS")
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_mqtya(self):
        return (
            pd.date_range(self.date_mqtya_begin, self.date_ya, freq="MS")
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_mqtqa(self):
        return (
            pd.date_range(self.date_mqtqa_begin, self.date_mqtqa_end, freq="MS")
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_mon(self):
        return pd.date_range(self.date, self.date, freq="MS").strftime("%Y-%m").tolist()

    @property
    def daterange_monya(self):
        return (
            pd.date_range(self.date_ya, self.date_ya, freq="MS")
            .strftime("%Y-%m")
            .tolist()
        )

    @property
    def daterange_monqa(self):
        return (
            pd.date_range(
                self.date + relativedelta(months=-1),
                self.date + relativedelta(months=-1),
                freq="MS",
            )
            .strftime("%Y-%m")
            .tolist()
        )

    def filter_date(self, period: str, year_ago: bool = False):
        df = self.data
        df["Date"] = df[self.date_column].apply(
            lambda x: datetime.datetime(year=int(x[:4]), month=int(x[5:]), day=1)
        )
        if period == "ytd":
            mask = (df["Date"] >= self.date_year_begin) & (df["Date"] <= self.date)
            mask_ya = (df["Date"] >= self.date_ya_begin) & (df["Date"] <= self.date_ya)
        elif period == "mat":
            mask = (df["Date"] >= self.date + relativedelta(months=-11)) & (
                df["Date"] <= self.date
            )
            mask_ya = (df["Date"] >= self.date_ya + relativedelta(months=-11)) & (
                df["Date"] <= self.date_ya
            )
        elif period == "mqt":
            mask = (df["Date"] >= self.date + relativedelta(months=-2)) & (
                df["Date"] <= self.date
            )
            mask_ya = (df["Date"] >= self.date_ya + relativedelta(months=-2)) & (
                df["Date"] <= self.date_ya
            )
        elif period == "mon":
            mask = df["Date"] == self.date
            mask_ya = df["Date"] == self.date_ya
        elif period == "qtr":  # 返回当季和环比季度的mask，当季可能不是一个完整季，环比季度是一个完整季
            month = self.date.month
            first_month_in_qtr = (month - 1) // 3 * 3 + 1  # 找到本季度的第一个月
            date_first_month_in_qtr = self.date.replace(month=first_month_in_qtr)
            date_first_month_in_qtrqa = date_first_month_in_qtr + relativedelta(
                months=-3
            )
            date_last_month_in_qtrqa = date_first_month_in_qtr + relativedelta(
                months=-1
            )
            mask = (df["Date"] >= date_first_month_in_qtr) & (df["Date"] <= self.date)
            mask_ya = (df["Date"] >= date_first_month_in_qtrqa) & (
                df["Date"] <= date_last_month_in_qtrqa
            )

        if year_ago:
            return df.loc[mask_ya, :]
        else:
            return df.loc[mask, :]

    def get_kpi(
        self, query_str: str = "ilevel_0 in ilevel_0", unit: str = "金额"
    ):  # 默认query语句能返回df总体
        df_value = self.get_pivot(
            index=self.date_column,
            values="销售金额（元）" if unit == "金额" else "销售盒数",
            aggfunc=sum,
            query_str=query_str,
        )

        value_ytd = df_value.loc[self.daterange_ytd, :].sum().values[0]
        value_ytdya = df_value.loc[self.daterange_ytdya, :].sum().values[0]
        value_mat = df_value.loc[self.daterange_mat, :].sum().values[0]
        value_matya = df_value.loc[self.daterange_matya, :].sum().values[0]
        value_mqt = df_value.loc[self.daterange_mqt, :].sum().values[0]
        value_mqtya = df_value.loc[self.daterange_mqtya, :].sum().values[0]
        value_mqtqa = df_value.loc[self.daterange_mqtqa, :].sum().values[0]
        value_mon = df_value.loc[self.daterange_mon, :].sum().values[0]
        value_monya = df_value.loc[self.daterange_monya, :].sum().values[0]
        value_monqa = df_value.loc[self.daterange_monqa, :].sum().values[0]

        df_target = self.get_pivot(
            index=self.date_column,
            values="指标金额（元）" if unit == "金额" else "指标盒数",
            aggfunc=sum,
            query_str=query_str,
        )

        value_target_ytd = df_target.loc[self.daterange_ytd, :].sum().values[0]
        value_target_mat = df_target.loc[self.daterange_mat, :].sum().values[0]
        value_target_mqt = df_target.loc[self.daterange_mqt, :].sum().values[0]
        value_target_mon = df_target.loc[self.daterange_mon, :].sum().values[0]

        dict_kpi = {
            "YTD": [
                value_ytd / 10000,
                (value_ytd - value_ytdya) / 10000,
                value_ytd / value_ytdya - 1,
                None,
                value_ytd / value_target_ytd,
            ],
            "MAT": [
                value_mat / 10000,
                (value_mat - value_matya) / 10000,
                value_mat / value_matya - 1,
                None,
                value_mat / value_target_mat,
            ],
            "MQT": [
                value_mqt / 10000,
                (value_mqt - value_mqtya) / 10000,
                value_mqt / value_mqtya - 1,
                value_mqt / value_mqtqa - 1,
                value_mqt / value_target_mqt,
            ],
            "MON": [
                value_mon / 10000,
                (value_mon - value_monya) / 10000,
                value_mon / value_monya - 1,
                value_mon / value_monqa - 1,
                value_mon / value_target_mon,
            ],
        }

        if unit == "金额":
            list_index = [
                "销售金额（万元）",
                "金额同比净增长（万元）",
                "金额同比增长率（%）",
                "金额环比增长率（%）",
                "金额达成率（%）",
            ]
        elif unit in ["盒数", "标准盒数"]:
            list_index = [
                "销售盒数（万盒）",
                "盒数同比净增长（万盒）",
                "盒数同比增长率（%）",
                "盒数环比增长率（%）",
                "盒数达成率（%）",
            ]
        df_table = pd.DataFrame(
            dict_kpi,
            index=list_index,
        )

        for idx in df_table.index:
            if "份额" in idx or "贡献" in idx or "达成率" in idx:
                df_table.loc[idx, :] = df_table.loc[idx, :].map(FORMAT_SHARE.format)
            elif "净增长" in idx:
                df_table.loc[idx, :] = df_table.loc[idx, :].map(FORMAT_DIFF.format)
            elif "价格" in idx or "单价" in idx:
                df_table.loc[idx, :] = df_table.loc[idx, :].map(FORMAT_CURRENCY.format)
            elif (
                "同比增长" in idx
                or "增长率" in idx
                or "CAGR" in idx
                or "同比变化" in idx
                or "vs." in idx
                or "%" in idx
            ):
                df_table.loc[idx, :] = df_table.loc[idx, :].map(FORMAT_GR.format)
            else:
                df_table.loc[idx, :] = df_table.loc[idx, :].map(FORMAT_ABS.format)

        return df_table


if __name__ == "__main__":
    print(D_SORTER)
