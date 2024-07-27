# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import pandas as pd


def match(df, match_sku):
    sku = df["sku"]
    sku = sku.split("*")[0]
    return int(match_sku[sku])


# Replace 'your_file.xls' with the path to your XLS file


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    match_sku = {'135156123156': '457852133', '156351781123': '567890123'}
    out_storage = 'out_storage.xlsx'
    # sheet_name = '订单模板'

    out_temu = 'out_temu.xlsx'

    # Read the Excel file
    df_out_storage = pd.read_excel(out_storage)  # 仓库
    df_out_temu = pd.read_excel(out_temu)  # temu

    # Select and rename the columns you want to combine
    df_temu_selected = df_out_temu[
        ["订单号", "子订单号", "货品SPU ID", "商品件数", "收货人姓名", "收货人联系方式", "详细地址1"]]
    df_storage_selected = df_out_storage[["跟踪号", "派送渠道", "收件人", "收件人电话", "收件地址1", "产品明细"]]

    # Rename columns to match the desired order
    df_temu_selected.rename(columns={"货品SPU ID": "sku", "收货人姓名": "收件人",
                                     "收货人联系方式": "收件人电话", "详细地址1": "收件人地址"}, inplace=True)
    df_storage_selected.rename(columns={"跟踪号": "跟踪单号", "派送方式": "物流承运商",
                                        "收件地址1": "收件人地址", "产品明细": "sku"}, inplace=True)

    df_storage_selected["sku"] = df_storage_selected.apply(lambda r: match(r, match_sku), axis=1)

    combined_df = pd.merge(df_temu_selected, df_storage_selected,
                           on=["sku", "收件人电话", "收件人", "收件人地址"],
                           how="inner")

    # Combine the selected columns into one dataframe

    # Write the combined dataframe to a new Excel file
    combined_df.to_excel('发货文件.xlsx', index=False)

    print("Combined DataFrame written to 'combined_output.xlsx'")
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
