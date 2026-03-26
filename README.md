# Reparations Payments Database (RPD)

This repository contains the **Reparations Payments Database (RPD)**, an open-access dataset of reparations payments committed and disbursed between 1945 and 2024. The dataset has been compiled by the REPAIR team and is accompanied by scripts and documentation used to build, clean, and analyze data on global reparations programs, enabling replication and further research.

The RPD is part of the REPAIR Project at the Department of Anthropology, University of Amsterdam.
More information: https://www.reparationsresearch.eu

The database focuses on **financial reparations** and includes information on:

- Payors and recipients
- Amounts pledged and, where available, paid
- Year of agreement
- Year of payments
- Type of reparations



## Repository Structure and Usage

This section outlines the repository structure, its contents, and provides guidance on how to use it.


### `01_data`

This folder contains:
- external/: External datasets used for data conversion and analysis (e.g., exchange rates)
- raw/: The main Reparations Payments Dataset (RPD) in Excel format
Users can download these files to reproduce the workflow or conduct independent analyses.


### `02_scripts`

This folder is organized into three subfolders:
- data_analysis/ : Scripts used to analyze the dataset and generate outputs (e.g., descriptive statistics, tables, or figures).
- data_auto_fill/ : Scripts designed to automatically populate missing or incomplete fields in the RPD using predefined rules or external data sources.
- data_conversion/ : Scripts used to standardize data, including currency conversions and value transformations.

All scripts in this folder include introductory comments explaining:
- Their purpose
- Required inputs
- Main instructions for use


### `03_docs`

This folder contains supporting materials to help users understand and work with the dataset, including:
- A codebook detailing variables, definitions, and coding decisions used in the RPD
