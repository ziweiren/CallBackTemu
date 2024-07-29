# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import pandas as pd


# Replace 'your_file.xls' with the path to your XLS file
def change_shipment(df):
    if df["物流承运商"] == "USPS-SC":
        return "USPS"
    elif df["物流承运商"] == "UPS-F":
        return "UPS"
    elif df["物流承运商"] == "Fedex-Z":
        return "FedEx"
    elif df["物流承运商"] == "DHL eCommerce":
        return df["物流承运商"]


def combine_table(table_temu, table_storage):
    out_storage = table_storage

    out_temu = table_temu

    # Read the Excel file
    df_out_storage = pd.read_excel(out_storage)  # 仓库
    df_out_temu = pd.read_excel(out_temu)  # temu

    # Select and rename the columns you want to combine
    df_temu_selected = df_out_temu[["订单号", "子订单号", "商品件数", "货品SKU ID"]]
    df_storage_selected = df_out_storage[["跟踪号", "派送渠道", "客户订单号"]]

    # Rename columns to match the desired order
    df_temu_selected.rename(columns={"货品SKU ID": "商品SKUID"}, inplace=True)
    df_storage_selected.rename(columns={"跟踪号": "跟踪单号", "派送渠道": "物流承运商", "客户订单号": "订单号"},
                               inplace=True)

    combined_df = pd.merge(df_temu_selected, df_storage_selected,
                           on=["订单号"],
                           how="inner")
    combined_df["物流承运商"] = combined_df.apply(lambda r: change_shipment(r), axis=1)

    # Combine the selected columns into one dataframe

    # Write the combined dataframe to a new Excel file
    combined_df.to_excel('发货文件.xlsx', index=False)

    print("Combined DataFrame written to '发货文件.xlsx'")


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    out_storage = 'out_storage.xlsx'
    out_temu = 'out_temu.xlsx'
    combine_table(out_temu, out_storage)
