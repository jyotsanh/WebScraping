import pandas as pd
df = pd.read_csv("dataa.txt", sep=",")
df.to_csv("eng_to_nep.csv",index=False)
