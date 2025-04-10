import os
import openpyxl
import argparse
from tqdm import tqdm
import statistics
from concurrent.futures import ThreadPoolExecutor


def get_parser():
    parser = argparse.ArgumentParser("金融科技創新應用")
    parser.add_argument("--input", default="Data/")
    parser.add_argument("--element", default="mom")
    parser.add_argument("--sheet-name", default="預測IC")
    return parser.parse_args()


def find_element(worksheet, find):
    row_pos, col_pos = None, None
    for row in worksheet.iter_rows(
        min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column
    ):
        for cell in row:
            if str(cell.value) == find:
                row_pos = cell.row
                col_pos = cell.column
                break
        if all([row_pos is not None, col_pos is not None]):
            break
    return row_pos, col_pos


class PortfolioCalculate:
    def __init__(self, inputs, sheet):
        self.input = inputs
        self.sheet = sheet

    def process(self, element):
        InputData = openpyxl.load_workbook(
            os.path.join(self.input, f"{element}.xlsx"), data_only=True
        )
        WorkSheet = InputData[f"{element}補值"]
        NextReturn = InputData["下個月月報酬補值"]
        StartDate = "2013/12"
        EndDate = "2024/12"
        _, StartColumn = find_element(WorkSheet, StartDate)
        _, EndColumn = find_element(WorkSheet, EndDate)

        ws_values = list(WorkSheet.values)
        nr_values = list(NextReturn.values)
        SplitNum = [96, 48, 19, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1]

        def _process_column(col_num):
            Data = []
            for row_idx in range(1, len(ws_values) - 1):
                key_val = ws_values[row_idx][col_num]
                nr_val = nr_values[row_idx][col_num]
                Data.append({key_val: nr_val})
            Data.sort(key=lambda d: list(d.keys())[0])
            SortedNR = [list(d.values())[0] for d in Data]
            BuyHigh, BuyLow = [], []
            for sp in SplitNum:
                buyhigh = statistics.mean(SortedNR[-sp:])
                buylow = statistics.mean(SortedNR[:sp])
                BuyHigh.append(buyhigh)
                BuyLow.append(buylow)
            Results = [low - high for high, low in zip(BuyHigh, BuyLow)]
            return Results

        columns_range = range(StartColumn - 1, EndColumn)
        DataStoration = []

        with ThreadPoolExecutor() as executor:
            results = list(
                tqdm(
                    executor.map(_process_column, columns_range),
                    total=len(list(columns_range)),
                    desc="Processing columns...",
                )
            )
            DataStoration.extend(results)

        self.store_data(DataStoration)

    def store_data(self, Results):
        Output = openpyxl.load_workbook(os.path.join(self.input, "IC.xlsx"))
        WorkSheet = Output[self.sheet]
        find = "投資組合"
        StandardRow, StandardCol = find_element(WorkSheet, find)
        StartRow = StandardRow + 2
        StartCol = StandardCol + 2

        for MonthData in tqdm(Results, desc="Saving Data..."):
            for idx, data in enumerate(MonthData):
                WorkSheet.cell(
                    row=StartRow+idx,
                    column=StartCol ,
                    value=round(float(data), 8),
                )
            StartCol += 1

        Output.save(os.path.join(self.input, "IC.xlsx"))
        print("Save Successfully!!!")


def main():
    args = get_parser()
    INPUTDIR = args.input
    SheetName = args.sheet_name
    PC = PortfolioCalculate(INPUTDIR, SheetName)
    PC.process(args.element)


if __name__ == "__main__":
    main()
