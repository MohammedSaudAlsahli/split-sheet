import re

import pandas as pd


class SplitSheet:
    def __init__(
        self,
        input_file: str,
        output_file: str,
        column_name: str = None,
        number: int = None,
    ):
        self.column_name: str = column_name
        self.output_file: str = output_file
        self.input_file: str = input_file
        self.number: int = number

    def run(self):
        dataframes = self._split_data()
        self._write_to_excel(dataframes)

    def _check_file_extension(self):
        if re.search(r"\.csv$", self.input_file, re.IGNORECASE):
            return "csv"
        elif re.search(r"\.xlsx$", self.input_file, re.IGNORECASE):
            return "xlsx"
        elif re.search(r"\.xls$", self.input_file, re.IGNORECASE):
            return "xls"
        elif re.search(r"\.ods$", self.input_file, re.IGNORECASE):
            return "ods"
        else:
            raise ValueError(
                "Unsupported file format. Please provide a .csv, .xlsx, .xls, or .ods file."
            )

    def _read_file(self) -> pd.DataFrame:
        """
        Read the input file and return a pandas DataFrame.

        Returns:
            pd.DataFrame: The contents of the input file as a DataFrame.
        """
        file_type = self._check_file_extension()
        if file_type == "csv":
            return pd.read_csv(self.input_file)
        elif file_type == "xlsx":
            return pd.read_excel(self.input_file, engine="openpyxl")
        elif file_type == "xls":
            return pd.read_excel(self.input_file, engine="xlrd")
        elif file_type == "ods":
            return pd.read_excel(self.input_file, engine="odf")

    def _is_date_column(self, data: pd.DataFrame, column: str) -> bool:
        """
        Check if a column in a DataFrame is a date column.

        Args:
            data (pd.DataFrame): The DataFrame to check.
            column (str): The name of the column to check.

        Returns:
            bool: True if the column is a date column, False otherwise.
        """
        try:
            pd.to_datetime(data[column])
            return True
        except ValueError:
            return False

    def _split_data(self):
        """
        This function splits the data based on the provided parameters: either by a specified number of rows or a specific column value.
        If 'number' is provided, it splits the data into parts based on the number of rows.
        If 'column_name' is provided, it splits the data based on unique values in that column.
        Returns a dictionary with keys based on the split criteria and corresponding dataframes.
        Raises a ValueError if neither 'column_name' nor 'number' is provided.
        """
        data = self._read_file()

        if self.number is not None:
            return {
                f"Part_{i // self.number + 1}": data.iloc[i : i + self.number]
                for i in range(0, len(data), self.number)
            }

        if self.column_name is not None:
            if self._is_date_column(data, self.column_name):
                data["Year"] = pd.to_datetime(data[self.column_name]).dt.year
                unique_values = data["Year"].unique()
                return {
                    year: data[data["Year"] == year].drop(columns=["Year"])
                    for year in unique_values
                }
            else:
                unique_values = data[self.column_name].unique()
                return {
                    value: data[data[self.column_name] == value]
                    for value in unique_values
                }

        raise ValueError("Either column_name or number must be provided.")

    def _write_to_excel(self, dataframes):
        with pd.ExcelWriter(self.output_file, engine="openpyxl") as writer:
            for sheet_name, df in dataframes.items():
                # Ensure sheet names are strings and limit the length
                sheet_name_str = str(sheet_name)[:31]
                df.to_excel(writer, sheet_name=sheet_name_str, index=False)
        print(f"Data has been split and saved to {self.output_file}")
