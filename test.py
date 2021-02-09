
import pandas as pd

df_left = pd.DataFrame([['x',2,True],
                        ['y',3,False],
                        ['x',1,False]],
                       columns=['alpha','num','class'])

df_right = pd.DataFrame([['x',.99],
                         ['b',.88],
                         ['z',.66]],
                        columns=['alpha2','score'])
df_right.set_index(['alpha2'],inplace=True)


print(df_left)
print('\n')
print(df_right)
print('\n')
print(df_left.join(df_right,on=['alpha'],how='left'))