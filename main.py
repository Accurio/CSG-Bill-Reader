# Copyright 2025 Accurio (https://github.com/Accurio)
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

from collections.abc import Sequence, Mapping, MutableMapping, Callable, Generator
from typing import Any

import sys
import os
import re
import datetime
import pypdf
import pandas

# 由于使用pypdf和正则表达式匹配文本和表格，若pypdf版本更新或电费单细节更新，可能会无法匹配。
if not tuple(map(int, pypdf.__version__.split('.'))) < (5, 1):
    raise RuntimeError("请使用 'pip install pypdf==5.0.1' 安装指定版本的pypdf！")


################################################################################
# 需要保存的数据

# 账单表格索引
BILLS_INDICES: list[str] = [
    "用户编号", "计量点编号", "用电开始时间", "用电结束时间",
]

# 账单表格文件名
BILLS_FILE_NAME: str = "账单.csv"

# 账单透视表格透视参数
BILLS_PIVOT_ARGUMENTS: dict[str, list[str]] = {
    "index": ["用电开始时间", "用电结束时间"],  # 首2列
    "columns": ["用户编号", "计量点编号"],  # 首2行
    "values": ["有功总合计电量"],
}

# 账单透视表文件名
BILLS_PIVOT_FILE_NAME: str = "账单透视表.csv"


################################################################################
# 函数

def search_to_dict(pattern: re.Pattern[str], string: str, info: str) -> dict[str, str]:
    """使用正则表达式模式查找并以字典类型返回匹配组"""
    match = pattern.search(string)
    if match is None:
        raise ValueError(info+"无匹配！")
    return match.groupdict()


def str_add_single_quotation_mark(string: str | Any) -> str | Any:
    """为数字文本添加单引号前缀（Excel）"""
    if not isinstance(string, str) or string.startswith("'"):
        return string
    return "'" + string

def date_from_iso_format(date: str | datetime.date | Any) -> datetime.date | Any:
    """将日期字符串转换为`datetime.date`"""
    if isinstance(date, datetime.date):
        return date
    elif sys.version_info < (3, 11) and len(date) != 10:
        return datetime.date(int(date[0:4]), int(date[4:6]), int(date[6:8]))
    return datetime.date.fromisoformat(date)

def convert_type(mapping: MutableMapping[str, Any],
    conversions: Mapping[str, Callable[[Any], Any]],
) -> None:
    """转换数据类型"""
    for key, func in conversions.items():
        if key not in mapping:
            continue
        value = mapping[key]
        if value is None:
            mapping.pop(key)
            continue
        mapping[key] = func(value)


################################################################################
# 检查是否为中国南方电网电费通知单

CHECK_PATTERN: re.Pattern[str] = re.compile(
    r"中国南方电网公司 ?\w+电网公司 ?电费通知单", flags=re.DOTALL)

def check_is_bill(text: str) -> bool:
    """检查是否为中国南方电网电费通知单"""
    return True if CHECK_PATTERN.search(text) is not None else False


################################################################################
# 替换空白字符

SUBSTITUTION_PATTERNS_REPLACEMENTS: tuple[tuple[re.Pattern[str], str], ...] = (
    (re.compile(r" {2,}"), ' '), # 压缩多个空格
    # ("表计资产编号 示数类型 上次表示数 本次表示数 倍率抄见电量\n (千瓦时)"
    #  "换表电量\n (千瓦时)退补电量\n (千瓦时)变/线损\n电量"       "公摊电量\n (千瓦时)"
    #  "免费电量\n (千瓦时)分表电量\n (千瓦时)合计电量\n (千瓦时)\n")
    # ("表计资产编号 示数类型 上次表示数 本次表示数 倍率抄见电量\n (千瓦时)"
    #  "换表电量\n (千瓦时)退补电量\n (千瓦时)变/线损电量\n (千瓦时)公摊电量\n (千瓦时)"
    #  "免费电量\n (千瓦时)分表电量\n (千瓦时)合计电量\n (千瓦时)\n")
    (re.compile(
            r" ?"
            r"(\(\d+\))?"        # "应收电费合计"序号"(1)"
            r"([\w/]+)"          # "变/线损电量\n (千瓦时)"的"变/线损电量"
            r"(\([\w/]+\))?"
            r"\n ?"
            r"(\([\w/]+\)|电量)" # "变/线损电量\n (千瓦时)"的"(千瓦时)", "变/线损\n电量"的"电量"
        ), r" \1\2\3\4"),
)

def substitute_white_space(text: str) -> str:
    """替换空白字符"""
    for pattern, replacement in SUBSTITUTION_PATTERNS_REPLACEMENTS:
        text = pattern.sub(replacement, text)
    return text


################################################################################
# 提取基本信息 Basic information

INFORMATION_PATTERN: re.Pattern[str] = re.compile((
    # 中国南方电网公司 电网公司 电费通知单
    r"尊敬的： ?"     r"(?P<用户>\w+).*"
    r"用户编号： ?"   r"(?P<用户编号>\w+).*"
    r"结算户号： ?"   r"(?P<结算户号>\w+).*"
    r"结算户名： ?"   r"(?P<结算户名>\w+).*"
    r"计量点编号： ?" r"(?P<计量点编号>\w+).*"
    # 基本信息 Basic information
    r"市场化属性分类： ?" r"(?P<市场化属性分类>\w+).*"
    r"用电类别： ?"       r"(?P<用电类别>\w+).*"
    r"用电开始时间： ?"   r"(?P<用电开始时间>\w+).*"
    r"用电结束时间： ?"   r"(?P<用电结束时间>\w+)"
), flags=re.DOTALL)

INFORMATION_CONVERSIONS: dict[str, Callable[[str], Any]] = {
      "用户编号": str_add_single_quotation_mark,
      "结算户号": str_add_single_quotation_mark,
    "计量点编号": str_add_single_quotation_mark,
    "用电开始时间": date_from_iso_format,
    "用电结束时间": date_from_iso_format,
}

def extract_information(text: str) -> dict[str, str | Any]:
    """提取基本信息 Basic information"""
    information = search_to_dict(INFORMATION_PATTERN, text, "基本信息")
    convert_type(information, INFORMATION_CONVERSIONS)
    return information


################################################################################
# 提取电量信息 Electricity Consumption Details

CONSUMPTION_非分时_示数类型: tuple[str, ...] = ("有功总", "无功总")
CONSUMPTION_分时_示数类型: tuple[str, ...] = ("尖", "峰", "平", "谷", "无功总")

CONSUMPTION_非分时_CHECK_PATTERN: re.Pattern[str] = re.compile(
    r".*".join(CONSUMPTION_非分时_示数类型), flags=re.DOTALL)
CONSUMPTION_分时_CHECK_PATTERN: re.Pattern[str] = re.compile(
    r".*".join(CONSUMPTION_分时_示数类型), flags=re.DOTALL)

CONSUMPTION_PATTERN_STR_ITEMS: tuple[str, ...] = (
    r"[\d\.-]+", r"\(千瓦时\)", "表计资产编号",         #  {0}-{2}
    "上次表示数", "本次表示数", "倍率", "抄见电量",     #  {3}-{6}
    "换表电量", "退补电量",   "变线损电量", "公摊电量", #  {7}-{10}
    "免费电量", "分表电量", "尖峰调整电量", "合计电量", # {11}-{14}
)

CONSUMPTION_非分时_PATTERN_STR = (
    # 电量信息 Electricity Consumption Details
    r"{2} ?示数类型 ?{3} ?{4} ?{5} ?{6}{1} ?"
    r"{7}{1} ?{8}{1} ?变/线损电量(?:{1})? ?{10}{1} ?"
    r"{11}{1} ?{12}{1} ?{14}{1}\n"
    # 有功总
    r"(?P<{15}{2}>(?:\w+\n\w+|\w+)) ?{15} "
    r"(?P<{15}{3}>{0}) (?P<{15}{4}>{0}) (?P<{15}{5}>{0}) (?P<{15}{6}>{0}) "
    r"(?P<{15}{7}>{0}) (?P<{15}{8}>{0}) (?P<{15}{9}>{0}) (?P<{15}{10}>{0}) "
    r"(?P<{15}{11}>{0}) (?P<{15}{12}>{0}) (?P<{15}{14}>{0})\n"
    # 无功总
    r"(?P<{16}{2}>(?:\w+\n\w+|\w+)) ?{16} "
    r"(?P<{16}{3}>{0}) (?P<{16}{4}>{0}) (?P<{16}{5}>{0}) (?P<{16}{6}>{0}) "
    r"(?P<{16}{7}>{0}) (?P<{16}{8}>{0}) (?P<{16}{9}>{0}) (?P<{16}{10}>{0}) "
    r"(?P<{16}{12}>{0}) (?P<{16}{14}>{0})\n"
).format(*(*CONSUMPTION_PATTERN_STR_ITEMS, *CONSUMPTION_非分时_示数类型))

CONSUMPTION_非分时_PATTERN: re.Pattern[str] = re.compile(CONSUMPTION_非分时_PATTERN_STR, flags=re.DOTALL)

CONSUMPTION_分时_PATTERN_STR = (
    # 电量信息 Electricity Consumption Details
    r"{2} ?示数类型 ?{3} ?{4} ?{5} ?{6}{1} ?"
    r"{7}{1} ?{8}{1} ?变/线损电量(?:{1})? ?{10}{1} ?"
    r"{11}{1} ?{12}{1}(?: ?{13}(?:{1})?)? ?{14}{1}\n"
    # 尖
    r"(?P<{15}{2}>(?:\w+\n\w+|\w+)) ?{15} "
    r"(?P<{15}{3}>{0}) (?P<{15}{4}>{0}) (?P<{15}{5}>{0}) (?P<{15}{6}>{0}) "
    r"(?P<{15}{7}>{0}) (?P<{15}{8}>{0}) (?P<{15}{9}>{0}) (?P<{15}{10}>{0}) "
    r"(?P<{15}{11}>{0}) (?P<{15}{12}>{0})(?: (?P<{15}{13}>{0}))? (?P<{15}{14}>{0})\n"
    # 峰
    r"(?P<{16}{2}>(?:\w+\n\w+|\w+)) ?{16} "
    r"(?P<{16}{3}>{0}) (?P<{16}{4}>{0}) (?P<{16}{5}>{0}) (?P<{16}{6}>{0}) "
    r"(?P<{16}{7}>{0}) (?P<{16}{8}>{0}) (?P<{16}{9}>{0}) (?P<{16}{10}>{0}) "
    r"(?P<{16}{11}>{0}) (?P<{16}{12}>{0})(?: (?P<{16}{13}>{0}))? (?P<{16}{14}>{0})\n"
    # 平
    r"(?P<{17}{2}>(?:\w+\n\w+|\w+)) ?{17} "
    r"(?P<{17}{3}>{0}) (?P<{17}{4}>{0}) (?P<{17}{5}>{0}) (?P<{17}{6}>{0}) "
    r"(?P<{17}{7}>{0}) (?P<{17}{8}>{0}) (?P<{17}{9}>{0}) (?P<{17}{10}>{0}) "
    r"(?P<{17}{11}>{0}) (?P<{17}{12}>{0}) (?P<{17}{14}>{0})\n"
    # 谷
    r"(?P<{18}{2}>(?:\w+\n\w+|\w+)) ?{18} "
    r"(?P<{18}{3}>{0}) (?P<{18}{4}>{0}) (?P<{18}{5}>{0}) (?P<{18}{6}>{0}) "
    r"(?P<{18}{7}>{0}) (?P<{18}{8}>{0}) (?P<{18}{9}>{0}) (?P<{18}{10}>{0}) "
    r"(?P<{18}{11}>{0}) (?P<{18}{12}>{0}) (?P<{18}{14}>{0})\n"
    # 无功总
    r"(?P<{19}{2}>(?:\w+\n\w+|\w+)) ?{19} "
    r"(?P<{19}{3}>{0}) (?P<{19}{4}>{0}) (?P<{19}{5}>{0}) (?P<{19}{6}>{0}) "
    r"(?P<{19}{7}>{0}) (?P<{19}{8}>{0}) (?P<{19}{9}>{0}) (?P<{19}{10}>{0}) "
    r"(?P<{19}{12}>{0}) (?P<{19}{14}>{0})\n"
).format(*(*CONSUMPTION_PATTERN_STR_ITEMS, *CONSUMPTION_分时_示数类型))

CONSUMPTION_分时_PATTERN: re.Pattern[str] = re.compile(CONSUMPTION_分时_PATTERN_STR, flags=re.DOTALL)

CONSUMPTION_CONVERSIONS: dict[str, Callable[[str], Any]] = {
    "表计资产编号": lambda s: s.replace('\n', ''),
    "上次表示数": float, "本次表示数": float, "倍率": int, "抄见电量": float,       #  {3}-{6}
    "换表电量": float, "退补电量": float,   "变线损电量": float, "公摊电量": float, #  {7}-{10}
    "免费电量": float, "分表电量": float, "尖峰调整电量": float, "合计电量": float, # {11}-{14}
}

CONSUMPTION_非分时_CONVERSIONS: dict[str, Callable[[str], Any]] = {
    (示数类型+key): func
    for key, func in CONSUMPTION_CONVERSIONS.items()
    for 示数类型 in CONSUMPTION_非分时_示数类型}

CONSUMPTION_分时_CONVERSIONS: dict[str, Callable[[str], Any]] = {
    (示数类型+key): func
    for key, func in CONSUMPTION_CONVERSIONS.items()
    for 示数类型 in CONSUMPTION_分时_示数类型}

def extract_consumption(text: str) -> dict[str, str | Any]:
    """提取电量信息 Electricity Consumption Details"""
    if CONSUMPTION_非分时_CHECK_PATTERN.search(text) is not None:
        consumption = search_to_dict(CONSUMPTION_非分时_PATTERN, text, "非分时电量信息")
        convert_type(consumption, CONSUMPTION_非分时_CONVERSIONS)
        consumption["表计资产编号"] = consumption["有功总表计资产编号"]

    elif CONSUMPTION_分时_CHECK_PATTERN.search(text) is not None:
        consumption = search_to_dict(CONSUMPTION_分时_PATTERN, text, "分时电量信息")
        convert_type(consumption, CONSUMPTION_分时_CONVERSIONS)
        consumption["表计资产编号"] = consumption["平表计资产编号"]
        consumption["有功总合计电量"] = str(sum(
            float(consumption[示数类型+"合计电量"])
            for 示数类型 in CONSUMPTION_分时_示数类型[:-1]))

    else:
        raise ValueError("电量信息示数类型无匹配！")
    return consumption


################################################################################
# 提取电费信息 Electricity Bill Information

BILL_PATTERN: re.Pattern[str] = re.compile((
    # 电费信息 Electricity Bill Information
    r"应收电费合计（大写）： ?(?P<应收电费合计大写>\w+) ?元?.*"
    r"应收电费合计（小写）： ?(?P<应收电费合计>[\d\.]+) ?元.*"
    r"平均电价： ?(?P<平均电价>[\d\.]+) ?\(元/千瓦时\)"
), re.DOTALL)

BILL_CONVERSIONS: dict[str, Callable[[Any], Any]] = {
    "应收电费合计": float, "平均电价": float}

def extract_bill(text: str) -> dict[str, str | Any]:
    """提取电费信息 Electricity Bill Information"""
    bill = search_to_dict(BILL_PATTERN, text, "电费信息")
    convert_type(bill, BILL_CONVERSIONS)
    return bill


################################################################################
# 提取和保存

def extract(text: str) -> dict[str, Any]:
    """提取电费通知单数据"""
    text = substitute_white_space(text)
    information = extract_information(text)
    consumption = extract_consumption(text)
    bill = extract_bill(text)
    return information | consumption | bill

def save(bills: Sequence[Mapping[str, Any]]) -> None:
    """保存为CSV"""
    frame = pandas.DataFrame.from_records(bills)
    frame.set_index(BILLS_INDICES).sort_index().to_csv(
        os.path.join(directory, BILLS_FILE_NAME), encoding='utf-8-sig')
    pandas.pivot(frame, **BILLS_PIVOT_ARGUMENTS).sort_index().to_csv(
        os.path.join(directory, BILLS_PIVOT_FILE_NAME), encoding='utf-8-sig')


################################################################################
# 主函数

directory = sys.argv[-1] if len(sys.argv) == 2 else os.path.dirname(__file__)
bills = []
for path, dirs_name, files_name in os.walk(directory):
    for file_name in files_name:
        if os.path.splitext(file_name)[-1] != '.pdf':
            continue
        for page in pypdf.PdfReader(os.path.join(path, file_name)).pages:
            text = page.extract_text()
            if not check_is_bill(text):
                continue
            try:
                bills.append(extract(text))
            except ValueError as e:
                print(file_name, *e.args)
if len(bills) == 0:
    print("没有可匹配的电费单！")
else:
    save(bills)
