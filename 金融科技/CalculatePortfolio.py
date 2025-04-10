import os 
import openpyxl
import argparse
from tqdm import tqdm, trange


def get_parser():
    
    parser = argparse.ArgumentParser('金融科技創新應用')
    parser.add_argument('--input', default='data/')
    parser.add_argument('--element', default='mom')
    parser.add_argument('--sheet-name', default='預測IC')
    

    return parser.parse_args()

def find_element(worksheet, find):
    row_pos, col_pos = None, None
    
    for row in worksheet.iter_rows(min_row=1, 
                                   max_row=worksheet.max_row, 
                                   min_col=1, 
                                   max_col=worksheet.max_column):
        for cell in row:
            if str(cell.value) == find:
                    
                row_pos = cell.row
                col_pos = cell.column
                break
        if all([row_pos != None, col_pos != None]):
            break

    return row_pos, col_pos 
    

class PortfolioCalculate:
    def __init__(self, inputs, sheet):
        
        self.input = inputs
        self.sheet = sheet

    def process(self, element):
        
        # Call Data
        InputData = openpyxl.open(os.path.join(self.input, f'{element}.xlsx'))
        WorkSheet = InputData[f'{element}補值']
        NextReturn = InputData['下個月月報酬補值']

        # Set Time
        StartDate = '2014/01'
        EndDate  = '2025/01'

        # Find Position
        row, start_column = find_element(WorkSheet, StartDate)
        _, end_column = find_element(WorkSheet, EndDate)

        total_data = end_column - start_column + 1

        
        data_storation = []
        for col_num in trange(total_data, desc='Storing Data...'):
            data = []
            for row_num in range(1, row):
                breakpoint()
                data.append({WorkSheet.cell(row=row_num, column=col_num): NextReturn.cell(row=row_num, column=col_num) })
                


        





def main():
    args = get_parser()
    INPUTDIR = args.input
    SheetName = args.sheet_name
    PC = PortfolioCalculate(INPUTDIR, SheetName)
    PC.process(args.element)


if __name__ == '__main__':
    main()


