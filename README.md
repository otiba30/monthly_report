# monthly_report
PythonからExcelを操作できるOpenPyXLを使って月報作成を自動化

## 1. ライブラリ導入
まずはOpenPyXLを導入します。

```sh
$ pip3 install openpyxl
```

## 2. 事前準備

以下のディレクトリ構成を想定します。

```
monthly_report
├── origin.xlsx(自身で配置)
├── script.py
└── shift.txt
```

ソースを準備します。

```sh
git clone https://github.com/otiba30/monthly_report.git
cd monthly_report
```

### 2.1. Excel元本の準備(original.xlsx)
月報Excelファイルの元本はMacVer3で動作を確認しています。</br>
ファイルを用意し、origin.xlsxにリネームしてmonthly_reportフォルダ内に配置してください。

### 2.2. シフト表の入力(shift.txt)
2022年9月のシフト表で、2日が日勤で出社、3日が夜勤で出社、8日が日勤で在宅の場合は下記のように記述します。

``` shift.txt
2022 9

2 1
3 2
8 3
```

### 2.3. 勤務時間などの入力(script.py)
参画しているプロジェクト名やシフトの勤務時間など編集をします。</br>
時刻は[days, seconds]で入力になります。<br>
日勤が8:30-17:00、夜勤が16:30-33:30の場合は下記のように記述します。

- 17:00 -> 0日 61200秒 -> [0,61200]
- 33:30 -> 1日 34200秒 -> [1,34200]


```python script.py
pjt_name = 'プロジェクト名'
company = '株式会社オブジェクティブコード'
work_style = '≪シフト（夜勤有）≫'
name = 'ブレイクザダークネス・オブスキュア'

# 所定開始, 所定終了, 休憩開始_1, 休憩終了_1, 休憩開始_2, 休憩終了_2,
# 開始時間, 終了時間,
# 出社or在宅,
# 備考

# 日勤(day shift)
list_ds = [[0, 30600], [0, 61200], [0, 43200], [0, 46800], None, None]\
         + [[0, 30600], [0, 61200]]\
         + ['出社']\
         + ['日勤:お仕事内容']

# 宿直(night shift)
list_ns = [[0, 59400], [1, 34200], [0, 75600], [0, 79200], [1, 10800], [1, 14400]]\
         + [[0, 59400], [1, 34200]]\
         + ['出社']\
         + ['宿直:お仕事内容']

# 日勤在宅(work remotely)
list_wr = [[0, 30600], [0, 61200], [0, 43200], [0, 46800], None, None]\
         + [[0, 30600], [0, 61200]]\
         + ['在宅']\
         + ['日勤:お仕事内容']
         
# 行き先, 用件, 金額, 使用機関, 移動経路_1, 移動経路_2, 往復or片道
list_te = ['OBG本社', '業務', 336, '電車', '東京', '赤坂見附', '往復']
```

## 3. 実行
実行すると【**name**】**month**月作業分_月末書類_MacVer3.xlsx が作成されます。</br>
条件付きフォーマット(Conditional Formatting)についてWarningが出ますが問題ありません。</br>
備考欄などの修正は作成されたファイルを編集してください。

```sh
$ python3 script.py shift.txt
```
