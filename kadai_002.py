import pandas as pd

# データフレームの作成
df = pd.DataFrame({
    '日付':
      ['2023-05-17', '2023-05-18', '2023-05-19', '2023-05-20', '2023-05-21'],
    '社員名': ['山田', '佐藤', '鈴木', '田中', '高橋'],
    '売上': [100, 200, 150, 300, 250],
    '部門': ['メーカー', '代理店', 'メーカー', '商社', '代理店'],
})

df["平均売上"] = df["売上"].mean()

ave = df["売上"].mean()
# print(mean)

def performance(level):
  result = '';

  if level >= ave + 50:
    result = "A"
  elif level >= ave:
    result = "B"
  else:
    result = "C"
  return result


df["業績ランク"] = df["売上"].apply(performance)

writer = pd.ExcelWriter("業績.xlsx")


df.to_excel(writer, sheet_name="sheet1", index=False)


writer.close()
