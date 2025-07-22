import pandas as pd

file_path = 'data.xlsx'

# Load the original matrix
matrix_df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)

# Load the conversion scales
conversion_df = pd.read_excel(file_path, sheet_name='Sheet2')

# Prepare lower AHP conversion table
lower_conversion = conversion_df[['Data Chart for lower AHP', 'Unnamed: 1']].iloc[1:].dropna()
lower_conversion.columns = ['Present_Value', 'Final_Value']
lower_conversion['Present_Value'] = lower_conversion['Present_Value'].astype(float)
lower_conversion['Final_Value'] = lower_conversion['Final_Value'].astype(float)

# Prepare upper AHP conversion table
upper_conversion = conversion_df[['Data Chart for upper AHP', 'Unnamed: 4']].iloc[1:].dropna()
upper_conversion.columns = ['Present_Value', 'Final_Value']
upper_conversion['Present_Value'] = upper_conversion['Present_Value'].astype(float)
upper_conversion['Final_Value'] = upper_conversion['Final_Value'].astype(float)

# Build mapping dictionary for upper AHP
upper_map = dict(zip(upper_conversion['Present_Value'], upper_conversion['Final_Value']))

# Build mapping dictionary for lower AHP (forward mapping)
lower_map_forward = dict(zip(lower_conversion['Present_Value'], lower_conversion['Final_Value']))

# Function to map a value for lower AHP including reciprocal handling
def map_lower_value(val):
    if val == 1.0:
        return 1.0
    elif val > 1.0:
        keys = sorted(k for k in lower_map_forward.keys() if k >= 1.0)
        closest_key = min(keys, key=lambda x: abs(x - val))
        return lower_map_forward[closest_key]
    else:  # val < 1
        reciprocal_val = 1.0 / val
        keys = sorted(k for k in lower_map_forward.keys() if k >= 1.0)
        closest_key = min(keys, key=lambda x: abs(x - reciprocal_val))
        return 1.0 / lower_map_forward[closest_key]

# Function to map a value for upper AHP (simple nearest neighbor mapping)
def map_upper_value(val):
    keys = sorted(upper_map.keys())
    closest_key = min(keys, key=lambda x: abs(x - val))
    return upper_map[closest_key]

# Map entire matrix for lower AHP
lower_ahp_matrix = matrix_df.copy()
for i in range(matrix_df.shape[0]):
    for j in range(matrix_df.shape[1]):
        lower_ahp_matrix.iat[i, j] = map_lower_value(matrix_df.iat[i, j])

# Map entire matrix for upper AHP
upper_ahp_matrix = matrix_df.copy()
for i in range(matrix_df.shape[0]):
    for j in range(matrix_df.shape[1]):
        upper_ahp_matrix.iat[i, j] = map_upper_value(matrix_df.iat[i, j])

# Display the results
print('Lower AHP Matrix:')
print(lower_ahp_matrix)
print('\nUpper AHP Matrix:')
print(upper_ahp_matrix)

# Save to Excel
with pd.ExcelWriter('Mapped_AHP_Matrices.xlsx') as writer:
    lower_ahp_matrix.to_excel(writer, sheet_name='Lower AHP', index=False, header=False)
    upper_ahp_matrix.to_excel(writer, sheet_name='Upper AHP', index=False, header=False)
