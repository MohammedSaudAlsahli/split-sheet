import typer

from . import SplitSheet

app = typer.Typer()


@app.command()
def main(
    input_file: str,
    output_file: str,
    column_name: str = typer.Option(
        None, "-c", "--column-name", help="Column name to split by"
    ),
    number: int = typer.Option(
        None, "-n", "--number", help="Number of rows to split by"
    ),
):
    """Split a spreadsheet by a column value or number of rows"""
    split_sheet = SplitSheet(input_file, output_file, column_name, number)
    split_sheet.run()


if __name__ == "__main__":
    app()
