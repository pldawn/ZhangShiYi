import pandas as pd


def main(fin_path, fout_path):
    incomes = pd.read_excel(fin_path, sheet_name="收入")
    costs = pd.read_excel(fin_path, sheet_name="成本")

    incomes.drop(incomes[incomes["摘要"] == "核算账簿累计"].index, inplace=True)
    costs.drop(costs[costs["摘要"] == "核算账簿累计"].index, inplace=True)

    incomes.drop(incomes[incomes["摘要"] == "总计"].index, inplace=True)
    costs.drop(costs[costs["摘要"] == "总计"].index, inplace=True)

    incomes.fillna({"贷方累计": 0.0}, inplace=True)
    costs.fillna({"借方累计": 0.0}, inplace=True)

    incomes.reset_index(inplace=True)
    costs.reset_index(inplace=True)

    def str_to_float(value):
        if type(value) == str:
            value = float(value.replace(" ", "").replace(",", ""))

        return value

    incomes["贷方累计"] = incomes["贷方累计"].map(str_to_float)
    costs["借方累计"] = costs["借方累计"].map(str_to_float)

    if "核算账簿名称" in incomes.columns:
        incomes_group = incomes.groupby(["核算账簿名称", "物料基本分类名称", "销售类型名称", "客商名称"])["贷方累计"].sum()
        costs_group = costs.groupby(["核算账簿名称", "物料基本分类名称", "销售类型名称", "客商名称"])["借方累计"].sum()

    elif "业务单元" in incomes.columns:
        incomes_group = incomes.groupby(["业务单元", "物料基本分类名称", "销售类型名称", "客商名称"])["贷方累计"].sum()
        costs_group = costs.groupby(["业务单元", "物料基本分类名称", "销售类型名称", "客商名称"])["借方累计"].sum()
    else:
        incomes_group = incomes.groupby(["物料基本分类名称", "销售类型名称", "客商名称"])["贷方累计"].sum()
        costs_group = costs.groupby(["物料基本分类名称", "销售类型名称", "客商名称"])["借方累计"].sum()

    result = pd.concat([incomes_group, costs_group], axis=1).fillna(0.0)
    result.reset_index(inplace=True)
    result.reset_index(inplace=True)
    result["index"] = result["index"] + 1

    if "核算账簿名称" in incomes.columns:
        result.columns = ["序号", "核算账簿名称", "商品品种", "贸易类别", "下游客户", "销售收入（万元）", "销售成本（万元）"]
    elif "业务单元" in incomes.columns:
        result.columns = ["序号", "业务单元", "商品品种", "贸易类别", "下游客户", "销售收入（万元）", "销售成本（万元）"]
    else:
        result.columns = ["序号", "商品品种", "贸易类别", "下游客户", "销售收入（万元）", "销售成本（万元）"]

    result.to_excel(fout_path, index=False)

    return len(result)


if __name__ == "__main__":
    my_fin_path = r"C:\Users\zhangwentao\Desktop\贸易类业务排查【ERP原始数据】删除表头\22年收入&成本【原物流】.xls"
    my_fout_path = r"C:\Users\zhangwentao\Desktop\贸易类业务排查【ERP原始数据】删除表头\22年收入&成本【原物流】result.xls"

    main(my_fin_path, my_fout_path)

    # my_fin_path = "data.xlsx"
    # my_fout_path = "result.xlsx"
    #
    # main(my_fin_path, my_fout_path)
    #
    # my_fin_path = "华南.xls"
    # my_fout_path = "华南result.xlsx"
    #
    # main(my_fin_path, my_fout_path)
    #
    # my_fin_path = "航发.xls"
    # my_fout_path = "航发result.xlsx"
    #
    # main(my_fin_path, my_fout_path)
