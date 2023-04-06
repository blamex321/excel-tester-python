import pandas as pd

data = pd.read_excel("/Users/blamex321/programming/python/dad-project/SL Agile Maturity - Self Assessment (Copy).xls")

yes_or_no_rows = [10, 11, 12, 13, 14, 15, 16, 19, 20, 21]
question_rows = list(range(22, 57))

for group_name, group_data in data.groupby('Scrum Leader Name'):
    result_data = pd.DataFrame()

    for row_number in yes_or_no_rows:
        temp_data = pd.DataFrame()
        yes_count = sum(group_data.iloc[:, row_number] == 'Yes')
        no_count = sum(group_data.iloc[:, row_number] == 'No')
        temp_data = pd.concat([temp_data, pd.DataFrame({
            'Group Name': [group_name],
            'Column Name': [data.columns[row_number]],
            'Yes Count': [yes_count],
            'No Count': [no_count],
            'Always': [0],
            'Mostly': [0],
            'Sometimes': [0],
            'Rarely': [0],
            'Never': [0]
        })])
        result_data = pd.concat([result_data, temp_data], ignore_index=True)

    for row_number in question_rows:
        temp_data = pd.DataFrame()
        always = sum(group_data.iloc[:, row_number] == 'Always')
        mostly = sum(group_data.iloc[:, row_number] == 'Mostly')
        sometimes = sum(group_data.iloc[:, row_number] == 'Sometimes')
        rarely = sum(group_data.iloc[:, row_number] == 'Rarely')
        never = sum(group_data.iloc[:, row_number] == 'Never')
        temp_data = pd.concat([temp_data, pd.DataFrame({
            'Group Name': [group_name],
            'Column Name': [data.columns[row_number]],
            'Yes Count': [0],
            'No Count': [0],
            'Always': [always],
            'Mostly': [mostly],
            'Sometimes': [sometimes],
            'Rarely': [rarely],
            'Never': [never]
        })])
        result_data = pd.concat([result_data, temp_data], ignore_index=True)

    result_data.to_excel(f"{group_name}.xlsx", index=False)
