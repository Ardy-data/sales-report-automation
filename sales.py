import pandas as pd

# 1. 读取数据
df = pd.read_csv("python_sales.csv")

# 2. 基本统计
total_sales = df["Sales"].sum()
average_sales = df["Sales"].mean()

sales_by_product = df.groupby("Product")["Sales"].sum().sort_values(ascending=False)
best_product = sales_by_product.idxmax()
best_product_sales = sales_by_product.max()

# 3. 英文总结（美国客户非常吃这一套）
summary_text = (
    f"Total sales reached ${total_sales:.2f}. "
    f"The average daily sales were ${average_sales:.2f}. "
    f"The best performing product was '{best_product}', "
    f"with total sales of ${best_product_sales:.2f}."
)

summary_df = pd.DataFrame({
    "Item": [
        "Executive Summary",
        "Total Sales",
        "Average Sales",
        "Best Product",
        "Best Product Sales"
    ],
    "Value": [
        summary_text,
        round(total_sales, 2),
        round(average_sales, 2),
        best_product,
        round(best_product_sales, 2)
    ]
})

# 4. Top Products 表
top_products_df = sales_by_product.reset_index()
top_products_df.columns = ["Product", "Total Sales"]

# 5. 写入 Excel
with pd.ExcelWriter("sales_report1.xlsx", engine="openpyxl") as writer:
    summary_df.to_excel(writer, sheet_name="Summary", index=False)
    top_products_df.to_excel(writer, sheet_name="Top Products", index=False)
    df.to_excel(writer, sheet_name="Raw Data", index=False)

print("✅ Professional sales report generated: sales_report.xlsx")
