This script counts the occurrences of words in a specific column in an EXCEL file.

# Requirements

Install the requirements file using `pip3`:
```
pip3 install -r requirements.txt
```

# How to use
Execute `-h` to know the arguments
```
usage: count_words.py [-h] [--show SHOW] excel column_name

positional arguments:
  excel        Excel file path
  column_name  The name of the column that you want to count

optional arguments:
  -h, --help   show this help message and exit
  --show SHOW  Show the number of the names sorted by the values.If you want
               to show all write 'all', by default shown the top 5

```

Example:
- The name of the excel file: `test.xlsx`
- The column that you want to count: `ProductName`
- Only show the top 3
```
pyhton3 count_words.py test.xlsx "ProductName" --show 3

+-------------+-------+
| ProductName | Times |
+-------------+-------+
|   product3  |   4   |
|   product1  |   3   |
|   product4  |   2   |
+-------------+-------+

```

# License

Licensed under GNU General Public License (GPL), version 3 or later.
