#!/usr/bin/env python3

# =============================================================================
# File name: dictparser.py
# Author: Aaron Wade
# Email: aaron.wade@yale.edu
# Date last modified: 01/13/2021
# Python version: 3.8.3
# =============================================================================

# =============================================================================

import re
from pathlib import Path
from string import punctuation, whitespace

import pandas as pd

# =============================================================================

# (Be sure to use FORWARD slashes when entering the paths below, e.g., "this/is/my/path/")

# 1. Insert path to the input file below (inside the quotation marks "")

# ================================#
# INSERT PATH TO INPUT FILE BELOW # (Code expects an Excel (.xlsx) file)
# ================================# (Use FORWARD slashes, e.g., "this/is/my/path/")
INPUT_FILE = r"Sample_Input.xlsx"

# 2. Insert sheet name below (inside the quotation marks "")

# ===========================#
# INSERT NAME OF SHEET BELOW #
# ===========================#

SHEET_NAME = r"sheet1"

# 3. Insert path to output folder below (inside the quotation marks "")

# ======================================#
# INSERT PATH TO OUTPUT DIRECTORY BELOW # (By default, the output folder will be created wherever this code is saved)
# ======================================# (Use FORWARD slashes, e.g., "this/is/my/path/")
OUTPUT_FOLDER = r"Parser_Output"

# Construct input path and store as a string
input_file_path = str(Path(INPUT_FILE).resolve())

# Construct output path and store as a string
Path(OUTPUT_FOLDER).mkdir(parents=True, exist_ok=True)
output_folder = Path(OUTPUT_FOLDER)
output_file_path = str((output_folder / "Parser_Output.xlsx").resolve())

# Construct dataframes
raw_df = pd.read_excel(input_file_path, sheet_name=SHEET_NAME)
raw_df.fillna("", inplace=True)
parsed_df = pd.DataFrame(columns=raw_df.columns.values.tolist())

# Regular expressions
term_plus_industry_pattern = (
    r"((?:(?:\"[A-Z]+\" )?(?:(?:\b[A-Z0-9]+\b) *-* *)+)(?<=[A-Z])[^A-Z]*)\(([^A-Z]+?)\)"
)
term_plus_industry_pattern_nogroups = (
    r"(?:(?:\"[A-Z]+\" )?(?:(?:\b[A-Z0-9]+\b) *-* *)+)(?<=[A-Z])[^A-Z]*\([^A-Z]+?\)"
)
number_pattern = r"(^[^a-zA-Z]+)\."

row_number = 1

# For initial testing, parse only the first 1000 rows of input
for index, row in raw_df.head(1000).iterrows():
    print("Parsing row " + str(row_number))

    # Collect row data
    first_term = str(row["term"]).strip()
    first_industry = str(row["industry"]).strip()
    first_number = str(row["number"]).strip()
    string_to_parse = str(row["definition"]).strip()

    # We parse the "definition" column of the input data, as this column might contain
    # data that should be split across multiple rows

    if string_to_parse:
        # If we get here, then there's text in the "definition" column of the current row (as expected)

        # ------------------- #
        # | TERM & INDUSTRY | #
        # ------------------- #

        # Parse out "chunks" containing a single term and its corresponding industry
        term_plus_industry_chunks = re.finditer(
            term_plus_industry_pattern, string_to_parse
        )

        terms = []
        term_examples = []
        industries = []

        for chunk in term_plus_industry_chunks:
            # Parse out 'term examples,' i.e., the text after the capitalized portion of the term (but preceding the industry)
            term_components = chunk.group(1).split(";", 1)

            if len(term_components) > 1:
                # If 'term examples' are present, grab them
                term_examples.append(
                    term_components[1].strip().lstrip(punctuation + whitespace)
                )
            else:
                term_examples.append("")

            # Grab term
            terms.append(term_components[0].strip())

            # Grab industry
            industries.append(chunk.group(2).strip())

        # ----------------------- #
        # | NUMBER & DEFINITION | #
        # ----------------------- #

        # Parse out "chunks" containing a number (if present) and a single definition
        number_plus_definition_chunks = re.split(
            term_plus_industry_pattern_nogroups, string_to_parse
        )
        # The first definition in the string being parsed belongs to the original row
        first_definition = (
            number_plus_definition_chunks.pop(0)
            .strip()
            .lstrip(punctuation + whitespace)
        )

        numbers = []
        definitions = []

        for chunk in number_plus_definition_chunks:
            # Parse out numbers
            number_match = re.match(number_pattern, chunk)

            definition = ""

            if number_match:
                # If number is present, grab it
                number = number_match.group(1).strip().lstrip(punctuation + whitespace)
                numbers.append(number)

                # Parse out definition
                definition = (
                    chunk.replace(number, "").strip().lstrip(punctuation + whitespace)
                )
            else:
                numbers.append("")

                # Parse out definition
                definition = chunk.strip().lstrip(punctuation + whitespace)

            # Grab definition
            definitions.append(definition)

        # Write the original row to the dataframe storing the parsed data
        parsed_df.loc[len(parsed_df)] = [
            first_number,
            first_term,
            first_industry,
            first_definition,
        ]

        # Write any newly-created rows to the dataframe storing the parsed data
        for i in range(len(terms)):
            # Append 'term examples' to definition, using a vertical bar (|) as a separator
            definition = (
                (term_examples[i] + " | " + definitions[i])
                if term_examples[i]
                else definitions[i]
            )
            parsed_df.loc[len(parsed_df)] = [
                numbers[i],
                terms[i],
                industries[i],
                definition,
            ]
    else:
        # If we end up here, then the definition column of the current row is empty
        # Just copy the row over to the dataframe storing the parsed data
        parsed_df.loc[len(parsed_df)] = [
            first_number,
            first_term,
            first_industry,
            string_to_parse,
        ]

    row_number += 1

# Write the parsed data to an excel file
parsed_df.to_excel(output_file_path, index=False)
