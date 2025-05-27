import openpyxl
import argparse
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any, Sequence
from utils.tools import load_worksheet, find_element, cumulative_reward
from tqdm import tqdm, trange

ELEMENT_ID = ["OLS", "RF", "NN1", "NN2", "NN3", "NN4", "NN5"]

def get_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser()
    p.add_argument("--input-dir", default="../data/")
    p.add_argument("--output", required=True)
    return p.parse_args()

def main() -> None:
    args = get_parser()
    input_dir = args.input_dir
    output = args.output

    data_dict = load_worksheet(inputs=input_dir, element_id=ELEMENT_ID)

    # 載入真實的投資組合    
    real = data_dict["OLS"]["真實IC"]

    real_portfolio = [[list(real.values)[11+c_idx][2+r_idx]for c_idx in range(13)] for r_idx in range(134) ]
    real_cumulative_reward = cumulative_reward(real_portfolio)

    # 讀取大盤
    mrow, mcol = find_element(real, "大盤")
    market = [
        list(real.values)[mrow - 1][mcol + 1 + col_idx]
        for col_idx in trange(134, desc="讀取大盤")
    ]

    rewards_dict = {}

    for element in ELEMENT_ID:
        values = list(data_dict[element][f"{element}累積報酬"].values)
        rewards_dict[element] = [[values[11+c_idx][2+r_idx]for r_idx in range(134)]for c_idx in range(13)]

    


    breakpoint()









    
    
    
    
    

if __name__ =="__main__":
    main()