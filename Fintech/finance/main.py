import os
import argparse
from tqdm import tqdm, trange
from utils.ic_calculate import IC
from utils.tools import *


'''
執行程式碼前須要注意是否有真實IC跟預測IC的表格樣式
'''

os.environ["CUDA_VISIBLE_DEVICE"] = "0"



def main() -> None:
    args = get_parser()
    input_dir = args.directory
    output = os.path.join(input_dir, args.save_file)
    model_name = args.model
    date_str = args.date
    n_jobs = args.n_jobs if args.n_jobs > 0 else os.cpu_count() * 10

    data_dict = load_worksheet(input_dir, output)  # 讀入所有 Excel
    print(f"呼叫 {model_name} model（n_jobs = {n_jobs}）")
    """
    times 是 RandomForest中的random_state參數
    為了驗證最佳報酬所給定的範圍(如果只想做一次，則times = 1
    然後RF model設定則為random_state = 999)
    """

    # 讀取大盤資料
    ic_ws = data_dict["IC"]["預測IC"]
    mrow, mcol = find_element(ic_ws, "大盤")
    market = [
        list(ic_ws.values)[mrow - 1][mcol + 1 + col_idx]
        for col_idx in trange(134, desc="讀取大盤")
    ]

    if not args.arr:

        start = int(args.range[0]) if model_name == "RF" else 1
        times = int(args.range[1]) if model_name == "RF" else 1
        last_data = None

        for random_state in range(start, times + 1):
            print(f"Times:{random_state}/{times}")

            random_state = 1042 if start == times else random_state
            # 計算所有因子的IC
            total_results = IC(input_dir, output, model_name, n_jobs).process(random_state)

            # 選擇基準因子
            data_storage, id_storage, select_values = select_factor(total_results)
            # 計算投資組合
            portfolio_data = []
            total_data = len(id_storage)
            for val, elem in tqdm(
                zip(select_values, id_storage),
                total=total_data,
                desc="計算投資組合",
            ):
                date_str, portfolio_data = portfolio_calculate(
                    data_dict, val, elem, portfolio_data, date_str
                )

            date_str = args.date
            # 計算勝率與累積報酬
            portfolio_value, total_count, win_rates = win_rate(portfolio_data, market)
            cumulative_rewards = cumulative_reward(portfolio_data)
            select_reward = max(
                [cumulative_rewards[i][-1] for i in range(len(cumulative_rewards))]
            )

            candidate = dict(
                IC=data_storage,
                ID=id_storage,
                select_values=select_values,
                portfolio_data=portfolio_data,
                portfolio_value=portfolio_value,
                total_count=total_count,
                win_rate=win_rates,
                cumulative_rewards=cumulative_rewards,
                select_reward=select_reward,
                best_reward=random_state,
            )
            if times == 1:
                last_data = candidate

            current_reward = last_data["select_reward"] if last_data is not None else None  
            current_state = last_data["best_reward"] if last_data is not None else None
            if model_name == "RF":
                print(f"目前設定參數為{random_state}，其累積報酬率{select_reward}，目前最佳參數{current_state}報酬率:{current_reward}")
            # 如果新的預測結果比目前的存取的最佳紀錄就更新，方便之後可以直接用最佳數據存取
            if last_data is None or select_reward > last_data["select_reward"]:
                if model_name == "RF":
                    print(f"在參數{random_state}有更好的累積報酬率")
                last_data = candidate

        best_state = last_data["best_reward"]
        best_reward = last_data["select_reward"]

        max_records, min_records, max_company, min_company = find_company(input_dir, id_storage)

        last_data['max_records'] = max_records
        last_data['max_company'] = max_company
        last_data['min_records'] = min_records
        last_data['min_company'] = min_company



        if model_name == "RF":
            print(f"最後選定參數為:{best_state}，其最佳累積報酬率:{best_reward}")

        # 計算夏普比率
        sharpe_data, sharpe_market = sharpe_ratio(portfolio_data, market)
        last_data["sharpe_ratio"] = sharpe_data
        last_data["sharpe_ratio_market"] = sharpe_market

        # 計算R-square
        real_ic = list(data_dict["IC"]["真實IC"].values)
        real_bm = real_ic[2][2:]
        real_sz = real_ic[3][2:]
        real_mm = real_ic[4][2:]

        real_ic = [[bm, sz, mm] for bm, sz, mm in zip(real_bm, real_sz, real_mm)]

        pr, rr, RS = r_square(real_ic, data_storage)

        last_data["PR_RS"] = pr
        last_data["RR_RS"] = rr
        last_data["RS"] = RS
        print(f"將數據儲存至{output}")

        data_store(output, last_data, model_name)


    else:
        '''
        由於Real.xlsx並沒有計算累積報酬、勝率等訊息
        因此在這邊會先計算完，如果存在還是會再重新算一次，不會影響到後續的數值
        '''

        real_data = openpyxl.load_workbook(os.path.join(input_dir, 'Real.xlsx'), data_only=True)
        


        real_portfolio_data =  [[list(real_data["預測IC"].values)[11+c_idx][2+r_idx]
                                 for c_idx in range(13)] for r_idx in range(134) ]
        


        real_portfolio_value, real_total_count, real_win_rates = win_rate(real_portfolio_data, market)
        real_cumulative_reward = cumulative_reward(real_portfolio_data)
        real_sharpe_data, sharpe_market = sharpe_ratio(real_portfolio_data, market)

        data = {
            "win_rate":real_win_rates,
            "portfolio_value":real_portfolio_value,
            "total_count": real_total_count,
            "cumulative_rewards":real_cumulative_reward,
            "sharpe_ratio":real_sharpe_data,
            "sharpe_ratio_market": sharpe_market
        }

        real_data_store(os.path.join(input_dir, 'Real.xlsx'), real_data, data)


        ''' 
        再投資累積報酬
        將(IC值+1)*(下一期IC值+1)...累乘後-1即為最後的再投資累積報酬率
        
        '''
        ELEMENT_ID = ["Real", "OLS", "RF", "NN1", "NN2", "NN3", "NN4", "NN5"]
        data_dict = load_worksheet(inputs=input_dir, element_id=ELEMENT_ID)

        for name, wb in tqdm(data_dict.items(), total=len(data_dict), desc='計算再投資累積報酬...'):
            portfolio_data =  [[list(wb["預測IC"].values)[11+c_idx][2+r_idx]
                                 for c_idx in range(13)] for r_idx in range(134) ]

            arr_cumulative(wb, name, portfolio_data, market, input_dir)

        print("計算完成!!!") 



def get_parser() -> argparse.Namespace:
    p = argparse.ArgumentParser("金融投資策略")
    p.add_argument("--directory", default="../data/", help="資料資料夾")
    p.add_argument("--save-file", default="NN1.xlsx", help="輸出檔名")
    p.add_argument("--date", default="2013/12", help="起始日期")
    p.add_argument("--range", nargs=2, default=[1, 1], help="設置要跑的參數範圍")
    p.add_argument(
        "--model", choices=["LR", "RF", "NN"], default="LR", help="選擇模型種類"
    )
    p.add_argument('--arr', action='store_true', help='計算所有的再投資累積報酬率')
    p.add_argument(
        "--n-jobs", type=int, default=1, help="平行執行緒數量；0=直接燃燒所有CPU"
    )
    return p.parse_args()


if __name__ == "__main__":
    main()
