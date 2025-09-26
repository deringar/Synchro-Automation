import argparse

from core import select_reports_and_extract


def main():
    parser = argparse.ArgumentParser(
        description="Select Synchro text reports and export their data to CSV."
    )
    parser.add_argument(
        "--output-dir",
        help="Optional folder to store the generated CSV files. Defaults to the same folder as each source report.",
    )
    args = parser.parse_args()
    select_reports_and_extract(args.output_dir)


if __name__ == "__main__":
    main()