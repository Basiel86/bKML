import pandas as pd

if __name__ == '__main__':
    process_columns_df = pd.DataFrame(columns=['COL_INDEX', 'COL_NAME'])
    process_columns_df = process_columns_df.append({'COL_INDEX': 1, 'COL_NAME': 11}, ignore_index=True)
    process_columns_df = process_columns_df.append({'COL_INDEX': 2, 'COL_NAME': 22}, ignore_index=True)
    process_columns_df = process_columns_df.append({'COL_INDEX': 3, 'COL_NAME': 33}, ignore_index=True)

    process_columns_df.loc[-0.5] = [9, 99]  # adding a row
    process_columns_df.index = process_columns_df.index + 1  # shifting index
    process_columns_df = process_columns_df.sort_index()  # sorting by index

    print(process_columns_df)
