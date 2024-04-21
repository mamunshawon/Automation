import pandas as pd

# Load Excel files
excel_file1 = pd.ExcelFile("excel_file1.xlsx")
excel_file2 = pd.ExcelFile("excel_file2.xlsx")

# Get sheet names from both Excel files
sheet_names1 = excel_file1.sheet_names
sheet_names2 = excel_file2.sheet_names

# Define aliases for table names (single letter)
table_aliases = {}
# Assigning single-letter aliases to each table name
alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
alias_index = 0
for table_name in sheet_names1:  # Using sheet names as table names
    table_aliases[table_name] = alphabet[alias_index]
    alias_index += 1

# Iterate over corresponding sheets
for sheet_name1, sheet_name2 in zip(sheet_names1, sheet_names2):
    df1 = excel_file1.parse(sheet_name1)

    # Check if the corresponding sheet exists in Excel file 2
    if sheet_name2 in sheet_names2:
        df2 = excel_file2.parse(sheet_name2)

        # Check if there is any non-null data in df2
        if df2.notnull().any().any():
            # Assuming column names are the same in both files
            table_updates = {}  # Dictionary to accumulate updates for each table
            for column in df1.columns:
                # Check if column has missing values in df1 but not in df2
                missing_mask = df1[column].notnull() & df2[column].isnull()
                if missing_mask.any():
                    if sheet_name1 not in table_updates:
                        table_updates[sheet_name1] = []
                    table_updates[sheet_name1].append(column)

            # Check if TRANSACTION_DATE_TIME is empty in df1 but not in df2
            if 'TRANSACTION_DATE_TIME' in df1.columns and df1['TRANSACTION_DATE_TIME'].isnull().all() and not df2['TRANSACTION_DATE_TIME'].isnull().all():
                # Update e_money_refund table
                if 'e_money_refund' in table_updates:
                    if 'TRANSACTION_DATE_TIME' not in table_updates['e_money_refund']:
                        table_updates['e_money_refund'].append('TRANSACTION_DATE_TIME')
                    e_money_freeze_df = excel_file1.parse('e_money_freeze')
                    transaction_date_time = e_money_freeze_df['TRANSACTION_DATE_TIME'].iloc[0]

            # Print accumulated update queries for each table
            for table_name, columns in table_updates.items():
                alias = table_aliases.get(table_name, table_name)
                print(f"UPDATE {table_name} {alias}")
                print("SET")
                updates = [f"    {alias}.{col} = NULL" if col != 'TRANSACTION_DATE_TIME' else f"    {alias}.{col} = '{transaction_date_time}'" if col == 'TRANSACTION_DATE_TIME' and alias == 'E' else f"    {alias}.{col} = ''" for col in columns]
                updates.append(f"    {alias}.MODIFIED_BY = {alias}.CREATED_BY")
                print(",\n".join(updates))
                ids = "', '".join(df1.loc[df1[columns[0]].notnull(), 'ID'].astype(str))
                print(f"WHERE {alias}.ID IN ('{ids}');\n")
        else:
            print(f"DELETE FROM {sheet_name1} WHERE ID IN ('{df1['ID'].astype(str).str.cat(sep = ', ')}')")
            print("\n")