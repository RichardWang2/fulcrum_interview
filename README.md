# LLM Excel Analyzer

A Python-based tool that leverages Large Language Models (LLMs) to analyze and process Excel files intelligently, outputting a JSON file with aggregated data.

## Overview

1. Read all Excel files in a directory using openpyxl, and extract all tables found in the files.
2. Use GPT-4 to identify semantically similar columns in the tables, then map the recognized columns to a standardized name.
3. Aggregate the data with coverage tiers.
4. Output JSON file for each Excel file with the aggregated data.

## How to run

1. clone the repository
2. add openai api key to .env file (OPENAI_API_KEY=your_api_key)
3. add all the files you want to analyze to the files/ directory
4. pip install -r requirements.txt
5. run llm_excel_analyzer.py

## Requirements

- Python 3.11
- Required packages (to be listed in requirements.txt)

## Roadblocks

- openpyxl is not able to extract all tables properly.
- some columns names are headers, some are in the first column, how to handle these cases?

## Current result:

I tried to identify similar columns and map columns to a standardized name, but the result is not good. Mainly because the raw data extracted from openpyxl is not sufficient to handle the non-standardized data.

