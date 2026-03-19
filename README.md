# Reparations Payments Database (RPD)

This repository contains an **open-access dataset** and the scripts used to build and analyze global reparations programs from World War II to 2024. It is part of the **REPAIR Project** at the Department of Anthropology of the University of Amsterdam.

See: https://www.reparationsresearch.eu

The database focuses on **financial reparations** and includes information on:

- Payors and recipients
- Amount pledged and paid (where available)
- Year of agreement
- Year of payment(s)
- Type of reparation

The repository also provides **Python scripts** to clean, fill, convert, and analyze the data. Documentation and methodology notes are included to ensure **transparency, reproducibility, and usability**.

## Repository Structure and Usage

This section outlines the repository structure, its contents, and provides guidance on how to use it.

`01_data`

		This folder contains:
		•	The main Reparations Payments Dataset (RPD) in Excel format
		•	External datasets used for data conversion and analysis (e.g., exchange rates)
		Users can download these files to reproduce the full workflow or conduct independent analyses.

`02_scripts`

		This folder is organized into three subfolders:
		•	data_analysis/ : Scripts used to analyze the dataset and generate outputs (e.g., descriptive statistics, tables, or figures).
		•	data_auto_fill/ : Scripts designed to automatically populate missing or incomplete fields in the RPD using predefined rules or external data sources.
		•	data_conversion/ : Scripts used to standardize data, including currency conversions and value transformations.
		All scripts in this folder include introductory comments explaining:
				•	Their purpose
				•	Required inputs
				•	Main instructions for use

`03_docs`

		This folder contains supporting materials to help users understand and work with the dataset, including:
		•	A codebook describing variables, definitions, and coding decisions

⸻
