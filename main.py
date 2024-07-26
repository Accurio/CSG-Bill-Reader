import sys
import os
import re
import datetime
import pypdf
import pandas

from typing import Any, Callable, Generator, Sequence


################################################################################
# Functions

def search(pattern: re.Pattern, string: str) -> dict[str, str]:
    match = pattern.search(string)
    if match is None:
        raise ValueError("No Match!")
    else:
        return match.groupdict()


def str_add_apostrophe(string: str | Any) -> str | Any:
    if isinstance(string, str) and not string.startswith('\''):
        return '\'' + string
    else:
        return string

def date_fromisoformat(date: str | datetime.date | Any) -> datetime.date | Any:
    if isinstance(date, datetime.date):
        return date
    elif not sys.version_info >= (3, 11) and len(date) == 8:
        return datetime.date(int(date[0:4]), int(date[4:6]), int(date[6:8]))
    else:
        return datetime.date.fromisoformat(date)


def convert_type(pairs: dict[str, Any],
    conversions: dict[Callable, Sequence[str]]
) -> None:
    for t, keys in conversions.items():
        for key in keys:
            value = pairs.get(key)
            if value is not None:
                pairs[key] = t(value)
            else:
                pairs.pop(key, None)


################################################################################
# Check and Substitute

CHECK_PATTERN: re.Pattern = re.compile(
    r"中国南方电网公司 ?\w+电网公司 ?电费通知单", flags=re.DOTALL)

def check_text(pattern: re.Pattern, text: str) -> bool:
    # For Excel
    return True if pattern.search(text) is not None else False


SUBSTITUTION_PATTERNS_REPLACEMENTS: tuple[tuple[re.Pattern, str]] = (
    (re.compile(r" {2,}"), ' '),
    (re.compile(r"(\d+SG\d+)\n(\d+) ?"), r"\1\2 "), # 表计资产编号
    (re.compile(r" ?(\(\d+\))?([\w/]+)(\([\w/]+\))?\n ?(\([\w/]+\)|电量)"), 
        r" \1\2\3\4"), # 电量\n\(千瓦时\)
)

def substitute_text(patterns_repls: re.Pattern, text: str) -> str:
    for pattern, replacement in patterns_repls:
        text = pattern.sub(replacement, text)
    return text


################################################################################
# Extract 基本信息 Basic information

INFORMATION_PATTERN: re.Pattern = re.compile((
    # 中国南方电网公司 电网公司 电费通知单
    r"尊敬的： ?"     r"(?P<用户>\w+).*"
    r"用户编号： ?"   r"(?P<用户编号>\d+).*"
    r"结算户号： ?"   r"(?P<结算户号>\d+).*"
    r"结算户名： ?"   r"(?P<结算户名>\w+).*"
    r"计量点编号： ?" r"(?P<计量点编号>\d+).*"
    # 基本信息 Basic information
    r"市场化属性分类： ?" r"(?P<市场化属性分类>\w+).*"
    r"用电类别： ?"       r"(?P<用电类别>\w+).*"
    r"用电开始时间： ?"   r"(?P<用电开始时间>\d+).*"
    r"用电结束时间： ?"   r"(?P<用电结束时间>\d+)"
), flags=re.DOTALL)

INFORMATION_CONVERSIONS: dict[Callable, tuple[str]] = {
    str_add_apostrophe: ("用户编号", "结算户号", "计量点编号"),
    date_fromisoformat: ("用电开始时间", "用电结束时间")}

def extract_information(text: str) -> dict[str, str | Any]:
    information = search(INFORMATION_PATTERN, text)
    convert_type(information, INFORMATION_CONVERSIONS)
    return information


################################################################################
# Extract 电量信息 Electricity Consumption Details

CONSUMPTION_非分时_ITEMS: tuple[str] = ("有功总", "无功总")
CONSUMPTION_分时_ITEMS: tuple[str] = ("尖", "峰", "平", "谷", "无功总")

CONSUMPTION_非分时_CHECK_PATTERN: re.Pattern = re.compile(
    r".*".join(CONSUMPTION_非分时_ITEMS), flags=re.DOTALL)
CONSUMPTION_分时_CHECK_PATTERN: re.Pattern = re.compile(
    r".*".join(CONSUMPTION_分时_ITEMS), flags=re.DOTALL)

CONSUMPTION_PATTERN_STR_ITEMS: tuple[str] = (
    r"[\d\.-]+", r"\(千瓦时\)", *CONSUMPTION_非分时_ITEMS, "表计资产编号",
    "上次表示数", "本次表示数", "倍率", "抄见电量", "换表电量", "退补电量",
    "变线损电量", "公摊电量", "免费电量", "分表电量", "尖峰调整电量", "合计电量",)

CONSUMPTION_非分时_PATTERN: re.Pattern = re.compile((
    # 电量信息 Electricity Consumption Details
    r"{4} ?示数类型 ?{5} ?{6} ?{7} ?{8}{1} ?{9}{1} ?{10}{1} ?"
    r"变/线损电量(?:{1})? ?{12}{1} ?{13}{1} ?{14}{1} ?{16}{1}\n"
    # 有功总
    r"(?P<{2}{4}>\d+SG\d+) {2} (?P<{2}{5}>{0}) (?P<{2}{6}>{0}) (?P<{2}{7}>{0}) "
    r"(?P<{2}{8}>{0}) (?P<{2}{9}>{0}) (?P<{2}{10}>{0}) (?P<{2}{11}>{0}) "
    r"(?P<{2}{12}>{0}) (?P<{2}{13}>{0}) (?P<{2}{14}>{0}) (?P<{2}{16}>{0})\n"
    # 无功总
    r"(?P<{3}{4}>\d+SG\d+) {3} (?P<{3}{5}>{0}) (?P<{3}{6}>{0}) (?P<{3}{7}>{0}) "
    r"(?P<{3}{8}>{0}) (?P<{3}{9}>{0}) (?P<{3}{10}>{0}) (?P<{3}{11}>{0}) "
    r"(?P<{3}{12}>{0}) (?P<{3}{14}>{0}) (?P<{3}{16}>{0})\n"
).format(*CONSUMPTION_PATTERN_STR_ITEMS), flags=re.DOTALL)

CONSUMPTION_分时_PATTERN: re.Pattern = re.compile((
    # 电量信息 Electricity Consumption Details
    r"{4} ?示数类型 ?{5} ?{6} ?{7} ?{8}{1} ?{9}{1} ?{10}{1} ?"
    r"变/线损电量(?:{1})? ?{12}{1} ?{13}{1} ?{14}{1}(?: ?{15}(?:{1})?)? ?{16}{1}\n"
    # 尖
    r"(?P<尖{4}>\d+SG\d+) 尖 (?P<尖{5}>{0}) (?P<尖{6}>{0}) (?P<尖{7}>{0}) "
    r"(?P<尖{8}>{0}) (?P<尖{9}>{0}) (?P<尖{10}>{0}) (?P<尖{11}>{0}) (?P<尖{12}>{0}) "
    r"(?P<尖{13}>{0}) (?P<尖{14}>{0})(?: (?P<尖{15}>{0}))? (?P<尖{16}>{0})\n"
    # 峰
    r"(?P<峰{4}>\d+SG\d+) 峰 (?P<峰{5}>{0}) (?P<峰{6}>{0}) (?P<峰{7}>{0}) "
    r"(?P<峰{8}>{0}) (?P<峰{9}>{0}) (?P<峰{10}>{0}) (?P<峰{11}>{0}) (?P<峰{12}>{0}) "
    r"(?P<峰{13}>{0}) (?P<峰{14}>{0})(?: (?P<峰{15}>{0}))? (?P<峰{16}>{0})\n"
    # 平
    r"(?P<平{4}>\d+SG\d+) 平 (?P<平{5}>{0}) (?P<平{6}>{0}) (?P<平{7}>{0}) "
    r"(?P<平{8}>{0}) (?P<平{9}>{0}) (?P<平{10}>{0}) (?P<平{11}>{0}) "
    r"(?P<平{12}>{0}) (?P<平{13}>{0}) (?P<平{14}>{0}) (?P<平{16}>{0})\n"
    # 谷
    r"(?P<谷{4}>\d+SG\d+) 谷 (?P<谷{5}>{0}) (?P<谷{6}>{0}) (?P<谷{7}>{0}) "
    r"(?P<谷{8}>{0}) (?P<谷{9}>{0}) (?P<谷{10}>{0}) (?P<谷{11}>{0}) "
    r"(?P<谷{12}>{0}) (?P<谷{13}>{0}) (?P<谷{14}>{0}) (?P<谷{16}>{0})\n"
    # 无功总
    r"(?P<{3}{4}>\d+SG\d+) {3} (?P<{3}{5}>{0}) (?P<{3}{6}>{0}) "
    r"(?P<{3}{7}>{0}) (?P<{3}{8}>{0}) (?P<{3}{9}>{0}) (?P<{3}{10}>{0}) "
    r"(?P<{3}{11}>{0}) (?P<{3}{12}>{0}) (?P<{3}{14}>{0}) (?P<{3}{16}>{0})\n"
).format(*CONSUMPTION_PATTERN_STR_ITEMS), flags=re.DOTALL)

CONSUMPTION_CONVERSIONS: dict[Callable, tuple[str]] = {
    int: CONSUMPTION_PATTERN_STR_ITEMS[7:8],
    float: (*CONSUMPTION_PATTERN_STR_ITEMS[5:7],
        *CONSUMPTION_PATTERN_STR_ITEMS[8:])}

CONSUMPTION_非分时_CONVERSIONS: dict[Callable, tuple[str]] = {
    t: tuple(item+key for item in CONSUMPTION_非分时_ITEMS for key in keys)
    for t, keys in CONSUMPTION_CONVERSIONS.items()}

CONSUMPTION_分时_CONVERSIONS: dict[Callable, tuple[str]] = {
    t: tuple(item+key for item in CONSUMPTION_分时_ITEMS for key in keys)
    for t, keys in CONSUMPTION_CONVERSIONS.items()}

def extract_consumption(text: str) -> dict[str, str | Any]:
    if CONSUMPTION_非分时_CHECK_PATTERN.search(text) is not None:
        consumption = search(CONSUMPTION_非分时_PATTERN, text)
        consumption["表计资产编号"] = consumption["有功总表计资产编号"]
        convert_type(consumption, CONSUMPTION_非分时_CONVERSIONS)

    elif CONSUMPTION_分时_CHECK_PATTERN.search(text) is not None:
        consumption = search(CONSUMPTION_分时_PATTERN, text)
        consumption["表计资产编号"] = consumption["平表计资产编号"]
        consumption["有功总合计电量"] = sum(
            float(consumption[示数类型+"合计电量"])
            for 示数类型 in CONSUMPTION_分时_ITEMS[:-1])
        convert_type(consumption, CONSUMPTION_分时_CONVERSIONS)

    else:
        raise ValueError("Unknown 用电类别!")
    return consumption


################################################################################
# Extract 电费信息 Electricity Bill Information

BILL_PATTERN: re.Pattern = re.compile((
    # 电费信息 Electricity Bill Information
    r"应收电费合计（大写）： ?(?P<应收电费合计大写>\w+) ?元?.*"
    r"应收电费合计（小写）： ?(?P<应收电费合计>[\d\.]+) ?元.*"
    r"平均电价： ?(?P<平均电价>[\d\.]+) ?\(元/千瓦时\)"
), re.DOTALL)

BILL_CONVERSIONS: dict[Callable, tuple[str]] = {
    float: ("应收电费合计", "平均电价")}

def extract_bill(text: str) -> dict[str, str | Any]:
    bill = search(BILL_PATTERN, text)
    convert_type(bill, BILL_CONVERSIONS)
    return bill


################################################################################
# Read and Extract

directory = sys.argv[-1] if len(sys.argv) == 2 else os.path.dirname(__file__)

def read(directory: str) -> Generator[dict[str, str], None, None]:
    for path, dirs_name, files_name in os.walk(directory):
        for file_name in files_name:
            if os.path.splitext(file_name)[-1] == '.pdf':
                for page in pypdf.PdfReader(os.path.join(path, file_name)).pages:
                    yield page.extract_text()

def extract(text: str) -> dict[str, str | Any] | None:
    if not check_text(CHECK_PATTERN, text):
        return None
    text = substitute_text(SUBSTITUTION_PATTERNS_REPLACEMENTS, text)
    return extract_information(text) \
        | extract_consumption(text) | extract_bill(text)

def load(directory: str) -> Generator[dict[str, str | Any], None, None]:
    for text in read(directory):
        bill = extract(text)
        if bill is not None:
            yield bill


################################################################################
# Rearrange and Save

BILLS_INDICES: list[str] = [
    "用户编号", "计量点编号", "用电开始时间", "用电结束时间"]
BILLS_FILE_NAME: str = "账单.csv"

BILLS_REARRANGED_KEYS: tuple = (
    ("用户编号", "计量点编号"), # Columns
    ("用电开始时间", "用电结束时间"), # Indices
    "有功总合计电量") # Value
BILLS_REARRANGED_FILE_NAME: str = BILLS_REARRANGED_KEYS[-1] + '.csv'

def rearrange_bills(bills: list[dict[Any, Any]],
    columns: tuple[Any], indices: tuple[Any], value: Any
):
    dicts = dict()
    for bill in bills:
        dicts.setdefault(tuple(bill[column] for column in columns), dict())\
        [tuple(bill[index] for index in indices)] = bill[value]
    return dicts

def sort_index_and_write_csv(dataframe: pandas.DataFrame, path: str) -> None:
    dataframe.sort_index(inplace=True)
    dataframe.to_csv(path, encoding='utf-8-sig')


################################################################################
# Main

bills = list(load(directory))
if bills:

    bills_dataframe = pandas.DataFrame(bills)
    bills_dataframe.set_index(BILLS_INDICES, inplace=True)
    sort_index_and_write_csv(bills_dataframe,
        os.path.join(directory, BILLS_FILE_NAME))

    bills_rearranged = rearrange_bills(bills, *BILLS_REARRANGED_KEYS)
    bills_rearranged_dataframe = pandas.DataFrame(bills_rearranged)
    sort_index_and_write_csv(bills_rearranged_dataframe,
        os.path.join(directory, BILLS_REARRANGED_FILE_NAME))
