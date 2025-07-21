import pandas as pd
import numpy as np


file_path = 'data.xlsx'


matrix_df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)


conversion_df = pd.read_excel(file_path, sheet_name='Sheet2')


lower_conversion = conversion_df[['Data Chart for lower AHP', 'Unnamed: 1']].iloc[1:].dropna()
lower_conversion.columns = ['Present_Value', 'Final_Value']
lower_conversion['Present_Value'] = lower_conversion['Present_Value'].astype(float)
lower_conversion['Final_Value'] = lower_conversion['Final_Value'].astype(float)


upper_conversion = conversion_df[['Data Chart for upper AHP', 'Unnamed: 4']].iloc[1:].dropna()
upper_conversion.columns = ['Present_Value', 'Final_Value']
upper_conversion['Present_Value'] = upper_conversion['Present_Value'].astype(float)
upper_conversion['Final_Value'] = upper_conversion['Final_Value'].astype(float)


lower_map = dict(zip(lower_conversion['Present_Value'], lower_conversion['Final_Value']))
upper_map = dict(zip(upper_conversion['Present_Value'], upper_conversion['Final_Value']))

def map_values(matrix, conv_map):
    sorted_keys = sorted(conv_map.keys())
    mapped_matrix = matrix.copy()
    for i in range(matrix.shape[0]):
        for j in range(matrix.shape[1]):
            val = matrix.iat[i, j]
            
            closest_key = min(sorted_keys, key=lambda x: abs(x-val))
            mapped_matrix.iat[i, j] = conv_map[closest_key]
    return mapped_matrix


lower_ahp_matrix = map_values(matrix_df, lower_map)
upper_ahp_matrix = map_values(matrix_df, upper_map)


print('Lower AHP Matrix:')
print(lower_ahp_matrix)
print('\nUpper AHP Matrix:')
print(upper_ahp_matrix)


with pd.ExcelWriter('Mapped_AHP_Matrices.xlsx') as writer:
    lower_ahp_matrix.to_excel(writer, sheet_name='Lower AHP', index=False, header=False)
    upper_ahp_matrix.to_excel(writer, sheet_name='Upper AHP', index=False, header=False)
