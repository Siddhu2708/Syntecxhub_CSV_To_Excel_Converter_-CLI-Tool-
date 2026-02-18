import argparse
import logging
import sys
from pathlib import Path
import pandas as pd


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:

    logger.info("Cleaning column names...")
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )

    logger.info("Handling missing values...")
    df = df.fillna("")

    logger.info("Attempting date parsing...")
    for col in df.columns:
        try:
            df[col] = pd.to_datetime(df[col])
        except (ValueError, TypeError):
            continue

    return df


def parse_rename_columns(rename_str: str) -> dict:

    rename_dict = {}
    if rename_str:
        pairs = rename_str.split(",")
        for pair in pairs:
            old, new = pair.split(":")
            rename_dict[old.strip().lower()] = new.strip()
    return rename_dict


def convert_csv_to_excel(input_file: str, output_file: str, rename_str: str = None):
    try:
        logger.info(f"Reading CSV file: {input_file}")
        df = pd.read_csv(input_file)

    except FileNotFoundError:
        logger.error("Input file not found.")
        sys.exit(1)
    except pd.errors.EmptyDataError:
        logger.error("CSV file is empty.")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error reading CSV: {e}")
        sys.exit(1)

    df = clean_dataframe(df)

    if rename_str:
        rename_dict = parse_rename_columns(rename_str)
        logger.info(f"Renaming columns: {rename_dict}")
        df = df.rename(columns=rename_dict)

    try:
        logger.info(f"Writing Excel file: {output_file}")
        df.to_excel(output_file, index=False, engine="openpyxl")
        logger.info("Conversion completed successfully!")

    except Exception as e:
        logger.error(f"Error writing Excel file: {e}")
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="CSV to Excel Converter with Cleaning & Normalization"
    )

    parser.add_argument(
        "-i", "--input",
        required=True,
        help="Path to input CSV file"
    )

    parser.add_argument(
        "-o", "--output",
        required=True,
        help="Path to output Excel file (.xlsx)"
    )

    parser.add_argument(
        "-r", "--rename",
        help="Column rename mapping: old1:new1,old2:new2"
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        logger.error("Input file does not exist.")
        sys.exit(1)

    if output_path.suffix != ".xlsx":
        logger.error("Output file must have .xlsx extension.")
        sys.exit(1)

    convert_csv_to_excel(
        input_file=str(input_path),
        output_file=str(output_path),
        rename_str=args.rename,
    )

if __name__ == "__main__":
    main()
