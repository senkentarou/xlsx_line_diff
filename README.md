# Usage
```
xlsx_line_diff.py [-h] [--page-name PAGE_NAME] xlsx_from_path xlsx_to_path
```

# Examples
```
$ python xlsx_line_diff.py ./sample_data/aiueo_1.xlsx ./sample_data/aiueo_2.xlsx  # select active sheet on both files.
$ pythonlxlsx_line_diff.py ./sample_data/aiueo_1.xlsx ./sample_data/aiueo_2.xlsx --page-name Sheet2  # select page name.
```

```
$ python xlsx_line_diff.py ./sample_data/aiueo_1.xlsx ./sample_data/aiueo_2.xlsx  # select active sheet on both files.

- Loading ./sample_data/aiueo_1.xlsx ...
+ Loading ./sample_data/aiueo_2.xlsx ...
---------- Diff lines! ----------
1: - か,き,く,け,こ
2: + あ,き,く,け,こ
か <=> あ
3: - さ,し,す,せ,そ
4: + さ,い,す,せ,そ
し <=> い
5: - な,に,ぬ,ね,の
6: + な,に,う,ね,の
ぬ <=> う
7: - や,None,ゆ,None,よ
8: + や,None,ゆ,え,よ
None <=> え
9: - ば,び,ぶ,べ,ぼ
10: + ば,び,ぶ,べ,お
ぼ <=> お
----------    Done.    ----------
```

# Environments
```
python==3.8.0

et-xmlfile==1.1.0                                                                                                                                                   │
openpyxl==3.0.7
```
