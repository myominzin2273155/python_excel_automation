import polars as pl

df = pl.read_excel("calculated_sales.xlsx")

print("--- All Data ---")
print(df)

filtered_df = df.filter(pl.col("Total") > 10000)

filtered_df.write_excel("polars_filtered.sales.xlsx")

print("\nPolars with Filtering Save finish")