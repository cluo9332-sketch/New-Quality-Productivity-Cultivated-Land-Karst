# -*- coding: utf-8 -*-


# ============================================================
# 1. Import libraries
# ============================================================

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl



# ============================================================
# 2. Basic configuration
# ============================================================

# Excel file name
excel_name = "year.xlsx"


# Get the current directory

try:
    base_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:
    base_dir = os.getcwd()


input_path = os.path.join(base_dir, excel_name)


if not os.path.exists(input_path):
    raise FileNotFoundError(
        f"Excel file not found:\n{input_path}\n"
        f"Please check whether excel_name is correct."
    )


print("--------------------------------------------------")
print("Reading file:")
print(input_path)
print("--------------------------------------------------")



# ============================================================
# 3. Font and plotting settings
# ============================================================

mpl.rcParams["font.family"] = "Times New Roman"
mpl.rcParams["mathtext.fontset"] = "stix"
mpl.rcParams["axes.unicode_minus"] = False

mpl.rcParams["font.size"] = 12

mpl.rcParams["figure.dpi"] = 300
mpl.rcParams["savefig.dpi"] = 600



# ============================================================
# 4. Read Excel file
# ============================================================

df = pd.read_excel(input_path)


print("\nData column names:")
print(df.columns.tolist())



# ============================================================
# 5. Column settings
# ============================================================

id_col = "id"
year_col = "year"


# Two study variables
variables = [
    "NQP",
    "CLIU"
]


required_cols = [
    id_col,
    year_col
] + variables



for col in required_cols:
    if col not in df.columns:
        raise ValueError(
            f"Missing column in Excel file: {col}"
        )


print("\nIdentified variables:")
print(variables)



# ============================================================
# 6. Data preparation
# ============================================================

data = df[
    [id_col, year_col] + variables
].copy()



# Remove rows with missing id or year

data = data.dropna(
    subset=[
        id_col,
        year_col
    ]
)



# Convert year to integer

data[year_col] = (
    data[year_col]
    .astype(int)
)



# Check for duplicate id-year combinations

dup_mask = data.duplicated(
    subset=[
        id_col,
        year_col
    ],
    keep=False
)


dup_count = dup_mask.sum()



if dup_count > 0:

    print(
        "\nDuplicate id-year combinations found; "
        "values will be averaged by id-year."
    )

    print(
        "Number of duplicate records:",
        dup_count
    )


    data = (
        data
        .groupby(
            [
                id_col,
                year_col
            ],
            as_index=False
        )[variables]
        .mean()
    )


else:

    print(
        "\nNo duplicate id-year combinations found."
    )



# Sort years

years = sorted(
    data[year_col]
    .unique()
    .tolist()
)


print("\nStudy years:")
print(years)



# ============================================================
# 7. Export descriptive statistics
# ============================================================

stats_output = os.path.join(
    base_dir,
    "Two_variables_descriptive_stats.xlsx"
)



with pd.ExcelWriter(
        stats_output,
        engine="openpyxl"
) as writer:


    # Cleaned data

    data.to_excel(
        writer,
        sheet_name="cleaned_data",
        index=False
    )


    # Statistics for each variable

    for var in variables:

        desc = (
            data
            .groupby(year_col)[var]
            .agg(
                n="count",
                mean="mean",
                median="median",
                std="std",
                min="min",
                q1=lambda x:x.quantile(0.25),
                q3=lambda x:x.quantile(0.75),
                max="max"
            )
            .reset_index()
        )


        desc.to_excel(
            writer,
            sheet_name=var,
            index=False
        )



print("\nDescriptive statistics output:")
print(stats_output)



# ============================================================
# 8. Plotting parameters
# ============================================================


colors = {

    "NQP":
        "#4C78A8",

    "CLIU":
        "#E45756"

}



panel_labels = {

    "NQP":
        "(a)",

    "CLIU":
        "(b)"

}



# Set a fixed random seed

rng = np.random.default_rng(2026)



# ============================================================
# 9. Create 1×2 subplots
# ============================================================


fig, axes = plt.subplots(
    1,
    2,
    figsize=(14,5)
)


axes = axes.flatten()



for ax, var in zip(
        axes,
        variables
):


    grouped_values = []

    valid_years = []



    for y in years:


        vals = (
            data
            .loc[
                data[year_col] == y,
                var
            ]
            .dropna()
            .values
        )


        if len(vals)>0:

            grouped_values.append(vals)

            valid_years.append(y)



    positions = np.arange(
        1,
        len(valid_years)+1
    )



    # --------------------------------------------------------
    # Box plots
    # --------------------------------------------------------

    ax.boxplot(

        grouped_values,

        positions=positions,

        widths=0.55,

        patch_artist=True,

        showfliers=False,


        medianprops=dict(
            color="black",
            linewidth=1.5
        ),


        whiskerprops=dict(
            color="#555555",
            linewidth=1.1
        ),


        capprops=dict(
            color="#555555",
            linewidth=1.1
        ),


        boxprops=dict(
            facecolor=colors[var],
            edgecolor="#444444",
            linewidth=1.1,
            alpha=0.35
        )

    )



    # --------------------------------------------------------
    # Jittered scatter points
    # --------------------------------------------------------

    for i, vals in enumerate(
            grouped_values,
            start=1
    ):


        jitter_x = rng.normal(
            loc=i,
            scale=0.06,
            size=len(vals)
        )


        ax.scatter(

            jitter_x,

            vals,

            s=14,

            color=colors[var],

            alpha=0.35,

            edgecolors="none",

            zorder=3

        )



    # --------------------------------------------------------
    # Panel labels
    # --------------------------------------------------------

    ax.text(

        0.02,
        0.96,

        panel_labels[var],

        transform=ax.transAxes,

        ha="left",

        va="top",

        fontsize=14,

        fontweight="bold",

        fontname="Times New Roman"

    )



    # Title

    ax.set_title(

        var,

        fontsize=14,

        fontname="Times New Roman",

        pad=8

    )



    ax.set_xlabel(

        "Year",

        fontsize=12,

        fontname="Times New Roman"

    )


    ax.set_ylabel(

        "Value",

        fontsize=12,

        fontname="Times New Roman"

    )



    ax.set_xticks(
        positions
    )


    ax.set_xticklabels(

        valid_years,

        rotation=45

    )



    # Tick-label fonts

    for label in ax.get_xticklabels():

        label.set_fontname(
            "Times New Roman"
        )

        label.set_fontsize(11)



    for label in ax.get_yticklabels():

        label.set_fontname(
            "Times New Roman"
        )

        label.set_fontsize(11)



    # Grid

    ax.grid(

        axis="y",

        linestyle="--",

        linewidth=0.8,

        alpha=0.35

    )



    # Borders

    for spine in ax.spines.values():

        spine.set_linewidth(1.0)

        spine.set_color("#444444")



    # Adjust y-axis limits

    ymin = np.nanmin(
        np.concatenate(grouped_values)
    )


    ymax = np.nanmax(
        np.concatenate(grouped_values)
    )


    yrange = ymax-ymin


    if yrange == 0:

        yrange = (
            abs(ymax)*0.1
            if ymax != 0
            else 1
        )


    ax.set_ylim(

        ymin-0.08*yrange,

        ymax+0.08*yrange

    )



# Automatically adjust layout

plt.tight_layout()



# ============================================================
# 10. Save figures
# ============================================================


png_output = os.path.join(
    base_dir,
    "Two_variables_boxplot_jitter.png"
)


tiff_output = os.path.join(
    base_dir,
    "Two_variables_boxplot_jitter.tiff"
)



plt.savefig(
    png_output,
    dpi=600,
    bbox_inches="tight"
)


plt.savefig(
    tiff_output,
    dpi=600,
    bbox_inches="tight"
)



plt.close()



print("\nFigure outputs:")

print(png_output)

print(tiff_output)



print("\n========================")
print("All tasks completed")
print("========================")