# Copyright (C) 2021 Quan Zhou
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <http://www.gnu.org/licenses/>.
#
# Authors:
#     Quan Zhou <zhquan7@gmail.com>
#


import argparse

import pandas
from prettytable import PrettyTable


SHOW_ALL = "all"


def main():
    args = arg_parser()
    file = args.excel
    column_name = args.column_name
    show = args.show

    pd = pandas.read_excel(file, index_col=0, engine='openpyxl')
    words = count(column_name, pd)

    table = create_table(column_name, show, words)
    print(table)


def arg_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument("excel",
                        help="Excel file path")
    parser.add_argument("column_name",
                        help="The name of the column that you want to count")
    parser.add_argument("--show", default="5",
                        help="Show the number of the names sorted by the values."
                             "If you want to show all write 'all', by default shown the top 5")

    args = parser.parse_args()
    return args


def count(column_name, pd):
    names_count = {}
    for name in pd[column_name]:
        if name not in names_count:
            names_count[name] = 1
        else:
            names_count[name] += 1
    sorted_value = dict(reversed(sorted(names_count.items(), key=lambda item: item[1])))

    return sorted_value


def create_table(column_name, show, words):
    if SHOW_ALL in show:
        max = len(words)
    else:
        max = int(show)

    table = PrettyTable()
    table.field_names = [column_name, "Times"]
    for name in words:
        if max == 0:
            break
        table.add_row([name, words[name]])
        max -= 1

    return table


if __name__ == '__main__':
    main()
