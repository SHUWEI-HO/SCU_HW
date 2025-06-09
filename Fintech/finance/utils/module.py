from keras.models import Sequential
import tensorflow as tf
from keras.layers import Dense
from tensorflow.keras import regularizers
from tensorflow.keras.optimizers import Adam
from tensorflow.keras.callbacks import EarlyStopping

def neural_network(input_dim):

    model = Sequential([
        Dense(32, activation="relu", input_dim=input_dim, 
              kernel_regularizer=regularizers.l1(1e-4)),
        # Dense(16, activation="relu"),
        # Dense(8, activation="relu"),
        # Dense(4, activation="relu"),
        # Dense(2, activation="relu"),
        Dense(1)
        ]
    )

    optim = Adam(learning_rate=1e-2)

    model.compile(optim, loss="mean_squared_error")
    
    early_stop = EarlyStopping(
        monitor="val_loss",
        patience = 5,
        mode = "min",
        restore_best_weights = True
    )

    return model, early_stop
