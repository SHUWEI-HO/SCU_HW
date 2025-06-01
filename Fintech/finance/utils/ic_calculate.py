import os
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List
import pandas as pd
import random
from tqdm import tqdm
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
from utils.module import neural_network
from utils.tools import *

os.environ["TF_CPP_MIN_LOG_LEVEL"] = "3" 
warnings.filterwarnings("ignore")

import tensorflow as tf

tf.get_logger().setLevel("FATAL")  
try:
    tf.autograph.set_verbosity(0)  
except AttributeError:
    pass


SEED = 1042

os.environ["OMP_NUM_THREADS"] = str(os.cpu_count())
random.seed(SEED)
np.random.seed(SEED)
tf.random.set_seed(SEED)
os.environ["PYTHONHASEED"] = str(SEED)

MODEL_MAP = {
    "LR": LinearRegression,
    "RF": RandomForestRegressor,
    "NN": "NN",
}

TARGETS = ["bm", "size", "mom"]


class IC:
    def __init__(
        self,
        directory: str,
        outputs: str,
        model_name: str,
        n_jobs: int = os.cpu_count(),
    ) -> None:
        self.dir = directory
        self.output = outputs
        self.n_jobs = n_jobs
        self.model_cls = MODEL_MAP[model_name]
        # For NeauralNetwork
        if self.model_cls == "NN":
            tf.keras.utils.disable_interactive_logging()
            self.model, self.early_stop = self._make_model() 

    def _make_model(self, random_state: int = 0):
        if self.model_cls is RandomForestRegressor:
            return self.model_cls(
                max_depth=3,
                random_state=random_state,
                n_estimators=100,
                n_jobs=self.n_jobs,
            ), None

        if self.model_cls == "NN":
            model, early_stop = neural_network(input_dim=964)
            return model, early_stop

        return self.model_cls(), None 

    def process(self, random_state: int = 999) -> Dict[str, List[float]]:
        total_results: Dict[str, List[float]] = {}

        def _fit_predict(i: int, idx: int, neural_network: bool = False) -> float:

            model, _ = self._make_model(random_state)
            X = df_x.iloc[:idx].values
            y = df_y.iloc[1 : idx + 1].values.ravel()
            tx = df_x.iloc[idx : idx + 1].values
            model.fit(X, y)
            return i, model.predict(tx)[0]

        if self.model_cls == "NN":
            for file_name in tqdm(TARGETS, total=len(TARGETS), desc="計算因子"):
                input_path = os.path.join(self.dir, f"{file_name}.xlsx")
                df_x = pd.read_excel(input_path, sheet_name=f"{file_name}補值").T
                df_y = pd.read_excel(input_path, sheet_name=f"{file_name}IC").T
                df_x.columns, df_y.columns = df_x.iloc[0], df_y.iloc[1]
                df_x, df_y = df_x.iloc[2:], df_y.iloc[2:]

                total_run = len(df_x) - 1
                start_time = 166 if file_name == "mom" else 178
                n0 = 48
                results = []
                for idx in range(start_time, total_run):
                    X = df_x.iloc[:idx].values.astype(np.float64)
                    y = df_y.iloc[1 : idx + 1].values.ravel().astype(np.float64)
                    tx = df_x.iloc[idx : idx + 1].values.astype(np.float64)

                    X_tr, X_vl = X[:idx-n0], X[idx-n0:idx]
                    y_tr, y_vl = y[:idx-n0], y[idx-n0:idx]

                    self.model.fit(
                        X_tr,
                        y_tr, 
                        validation_data=(X_vl, y_vl),
                        epochs=100, 
                        batch_size=32, 
                        verbose=0,
                        callbacks=[self.early_stop])

                    pred = self.model.predict(tx)
                    results.append(pred.tolist()[0][0])
                    
                total_results[file_name] = results

        else:
            for file_name in TARGETS:
                input_path = os.path.join(self.dir, f"{file_name}.xlsx")
                df_x = pd.read_excel(input_path, sheet_name=f"{file_name}補值").T
                df_y = pd.read_excel(input_path, sheet_name=f"{file_name}IC").T
                df_x.columns, df_y.columns = df_x.iloc[0], df_y.iloc[1]
                df_x, df_y = df_x.iloc[2:], df_y.iloc[2:]

                total_run = len(df_x) - 1
                start_time = 166 if file_name == "mom" else 178
                idx_range = list(range(start_time, total_run))
                results = [None] * len(idx_range)
                with ThreadPoolExecutor(max_workers=self.n_jobs) as executor:
                    futures = [
                        executor.submit(_fit_predict, i, idx)
                        for i, idx in enumerate(idx_range)
                    ]
                    for f in tqdm(
                        as_completed(futures),
                        total=len(futures),
                        desc=f"計算{file_name} IC因子 ",
                    ):
                        i, pred = f.result()
                        results[i] = pred

                total_results[file_name] = results

        return total_results
