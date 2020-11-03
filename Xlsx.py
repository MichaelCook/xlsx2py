# Copyright (C) Michael Cook <michael@waxrat.com>
import datetime
import types

def RANGE(sheet, start_col: int, start_row: int, end_col: int, end_row: int):
    """
    Column and row arguments are 1-based.  For example, start_col=1 is column A
    """
    for rowi in range(start_row - 1, min(end_row, len(sheet.cell_methods))):
        row = sheet.cell_methods[rowi]
        for coli in range(start_col - 1, min(end_col, len(row))):
            cell_method = row[coli]
            if cell_method is not None:
                yield cell_method()

def IF(cond, if_value, then_value):
    return if_value if cond else then_value

def DATE(year, month, day, hour, minute):
    """
    Convert to Excel time (floating point days since 1900)
    """
    posix_time = datetime.datetime(year, month, day, hour, minute,
                                   tzinfo=datetime.timezone.utc).timestamp()
    date_1970 = 25569           # =DATE(1970,1,1)
    return posix_time / (24 * 60 * 60) + date_1970

def TIME(hour, minute, second, microsecond):
    t = hour * 60 * 60 + minute * 60 + second + microsecond * 1e-6
    return t / (24 * 60 * 60)


def all_values(items):
    """
    `items` is zero or more cell values or generators of cell values (e.g., RANGE)
    Yield all the non-None values
    """
    for item in items:
        if item is None:
            continue
        if isinstance(item, types.GeneratorType):
            for cell_value in item:
                yield cell_value
        else:
            yield item

def all_numbers(items):
    """
    `items` is zero or more cell values or generators of cell values (e.g., RANGE)
    Yield all the numeric values
    """
    for item in all_values(items):
        if isinstance(item, (int, float)):
            yield item

def AVERAGE(*args):
    nums = list(all_numbers(args))
    if not nums:
        return None             # '#DIV/0!'
    return sum(nums) / len(nums)

def MAX(*args):
    nums = list(all_numbers(args))
    if not nums:
        return 0
    return max(nums)

def MIN(*args):
    nums = list(all_numbers(args))
    if not nums:
        return 0
    return min(nums)

def REPT(text, count):
    return str(text) * int(count)

def SUM(*args):
    return sum(all_numbers(args))
