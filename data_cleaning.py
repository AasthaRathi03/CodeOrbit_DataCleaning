"""
Data Cleaning Project - CodeOrbit Tech Internship
=====================================================
Purpose: Clean and preprocess retail sales data
Author: [Your Name]
Date: 2024
Description: This script handles missing values, duplicates,
             and data formatting issues in customer shopping dataset.
"""

import pandas as pd
import numpy as np
from datetime import datetime


class DataCleaner:
    """
    A professional data cleaning class for retail sales data.
    Handles missing values, duplicates, and data type conversions.
    """

    def __init__(self, file_path):
        """
        Initialize the DataCleaner with input file path.

        Args:
            file_path (str): Path to the Excel/CSV file
        """
        self.input_file = file_path
        self.df = None
        self.df_cleaned = None
        self.cleaning_log = []

    def load_data(self):
        """Load data from Excel file."""
        try:
            print("=" * 60)
            print("STEP 1: LOADING DATA")
            print("=" * 60)
            self.df = pd.read_csv(self.input_file)
            print("File loaded successfully")
            print("Total Rows: {}".format(len(self.df)))
            print("Total Columns: {}".format(len(self.df.columns)))
            self.cleaning_log.append("Step 1: Data loaded successfully")
            return True
        except Exception as e:
            print("Error loading file: {}".format(e))
            return False

    def display_data_info(self):
        """Display basic information about the dataset."""
        print("\n" + "=" * 60)
        print("DATASET OVERVIEW")
        print("=" * 60)
        print("\nColumn Names and Types:")
        for col in self.df.columns:
            print("  {} : {}".format(col, self.df[col].dtype))

        print("\nFirst 5 rows of data:")
        print(self.df.head())
        print("\nLast 5 rows of data:")
        print(self.df.tail())

    def check_missing_values(self):
        """Identify and report missing values."""
        print("\n" + "=" * 60)
        print("STEP 2: CHECKING FOR MISSING VALUES")
        print("=" * 60)

        missing_data = self.df.isnull().sum()
        missing_percent = (self.df.isnull().sum() / len(self.df)) * 100

        if missing_data.sum() == 0:
            print("No missing values found")
        else:
            print("\nMissing Values Detected:")
            for col in missing_data[missing_data > 0].index:
                print(
                    "  {} : {} missing ({}%)".format(
                        col, missing_data[col], round(missing_percent[col], 2)
                    )
                )

        self.cleaning_log.append(
            "Step 2: Checked missing values - Found {} total".format(missing_data.sum())
        )
        return missing_data

    def check_duplicates(self):
        """Identify and report duplicate rows."""
        print("\n" + "=" * 60)
        print("STEP 3: CHECKING FOR DUPLICATES")
        print("=" * 60)

        duplicates = self.df.duplicated().sum()
        print("Total Duplicate Rows: {}".format(duplicates))

        if duplicates > 0:
            print("\nFirst few duplicate rows:")
            print(self.df[self.df.duplicated(keep=False)].head())
        else:
            print("No duplicate rows found")

        self.cleaning_log.append(
            "Step 3: Checked duplicates - Found {} duplicates".format(duplicates)
        )
        return duplicates

    def clean_data(self):
        """Apply all cleaning transformations."""
        print("\n" + "=" * 60)
        print("STEP 4: CLEANING DATA")
        print("=" * 60)

        self.df_cleaned = self.df.copy()

        # Remove duplicates
        print("\n  Removing duplicate rows...")
        initial_rows = len(self.df_cleaned)
        self.df_cleaned = self.df_cleaned.drop_duplicates()
        removed = initial_rows - len(self.df_cleaned)
        print("  Removed {} duplicate rows".format(removed))
        self.cleaning_log.append("Removed {} duplicate rows".format(removed))

        # Handle missing values
        print("\n  Handling missing values...")
        for col in self.df_cleaned.columns:
            if self.df_cleaned[col].isnull().sum() > 0:
                if self.df_cleaned[col].dtype in ["int64", "float64"]:
                    median_val = self.df_cleaned[col].median()
                    self.df_cleaned[col].fillna(median_val, inplace=True)
                    print("  {} : Filled with median ({})".format(col, median_val))
                    self.cleaning_log.append("Filled {} with median value".format(col))
                else:
                    mode_val = self.df_cleaned[col].mode()[0]
                    self.df_cleaned[col].fillna(mode_val, inplace=True)
                    print("  {} : Filled with mode ({})".format(col, mode_val))
                    self.cleaning_log.append("Filled {} with mode value".format(col))

        # Reset index
        print("\n  Resetting index...")
        self.df_cleaned.reset_index(drop=True, inplace=True)
        print("  Index reset successfully")

        print("\nData cleaning completed")

    def display_cleaning_summary(self):
        """Display before/after comparison."""
        print("\n" + "=" * 60)
        print("CLEANING SUMMARY - BEFORE vs AFTER")
        print("=" * 60)

        print("\nOriginal Dataset:")
        print("  Total Rows: {}".format(len(self.df)))
        print("  Total Columns: {}".format(len(self.df.columns)))
        print("  Missing Values: {}".format(self.df.isnull().sum().sum()))
        print("  Duplicate Rows: {}".format(self.df.duplicated().sum()))

        print("\nCleaned Dataset:")
        print("  Total Rows: {}".format(len(self.df_cleaned)))
        print("  Total Columns: {}".format(len(self.df_cleaned.columns)))
        print("  Missing Values: {}".format(self.df_cleaned.isnull().sum().sum()))
        print("  Duplicate Rows: {}".format(self.df_cleaned.duplicated().sum()))

        print("\nRows Removed: {}".format(len(self.df) - len(self.df_cleaned)))

    def save_cleaned_data(self, output_file="cleaned_data.xlsx"):
        """Save cleaned data to Excel file."""
        print("\n" + "=" * 60)
        print("STEP 5: SAVING CLEANED DATA")
        print("=" * 60)
        try:
            self.df_cleaned.to_excel(output_file, index=False)
            print("Cleaned data saved to: {}".format(output_file))
            self.cleaning_log.append("Saved cleaned data to {}".format(output_file))
        except Exception as e:
            print("Error saving file: {}".format(e))

    def generate_report(self, report_file="cleaning_report.txt"):
        """Generate cleaning documentation report."""
        print("\n" + "=" * 60)
        print("STEP 6: GENERATING REPORT")
        print("=" * 60)

        report_content = """
CLEANING DOCUMENTATION REPORT
=====================================
Generated on: {}

PROJECT: CodeOrbit Tech Data Analyst Internship
TASK: Data Cleaning and Preprocessing

CLEANING STEPS PERFORMED:
=====================================
""".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        for i, log in enumerate(self.cleaning_log, 1):
            report_content += "\n{}. {}".format(i, log)

        report_content += """

BEFORE & AFTER STATISTICS:
=====================================
Original Dataset:
  Total Rows: {}
  Total Columns: {}
  Missing Values: {}
  Duplicate Rows: {}

Cleaned Dataset:
  Total Rows: {}
  Total Columns: {}
  Missing Values: {}
  Duplicate Rows: {}

COLUMNS IN DATASET:
=====================================
""".format(
            len(self.df),
            len(self.df.columns),
            self.df.isnull().sum().sum(),
            self.df.duplicated().sum(),
            len(self.df_cleaned),
            len(self.df_cleaned.columns),
            self.df_cleaned.isnull().sum().sum(),
            self.df_cleaned.duplicated().sum(),
        )

        for col in self.df_cleaned.columns:
            report_content += "\n {} ({})".format(col, self.df_cleaned[col].dtype)

        try:
            with open(report_file, "w") as f:
                f.write(report_content)
            print("Report saved to: {}".format(report_file))
        except Exception as e:
            print("Error saving report: {}".format(e))

    def run_complete_cleaning(self):
        """Run the complete cleaning pipeline."""
        if self.load_data():
            self.display_data_info()
            self.check_missing_values()
            self.check_duplicates()
            self.clean_data()
            self.display_cleaning_summary()
            self.save_cleaned_data()
            self.generate_report()

            print("\n" + "=" * 60)
            print("DATA CLEANING COMPLETED SUCCESSFULLY")
            print("=" * 60)
            print("\nOutput Files Generated:")
            print("  1. cleaned_data.xlsx (Cleaned dataset)")
            print("  2. cleaning_report.txt (Documentation)")


if __name__ == "__main__":
    cleaner = DataCleaner("customer_shopping_data.csv")
    cleaner.run_complete_cleaning()
