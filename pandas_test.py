import pandas as pd
import numpy as np
df1=pd.DataFrame(np.random.randint(10,20,size=(4,4)))
print("df1=\n{0}".format(df1))
print("df1=\n{0}".format(df1.apply(lambda x:x**2)))
print("df1=\n{0}".format(df1.sum()))
print("df1=\n{0}".format(df1.sum(1)))