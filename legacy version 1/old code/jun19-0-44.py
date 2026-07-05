# %%
import random
import tensorflow as tf
from tensorflow import keras
from tensorflow.keras import layers
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import *

# %%
df = pd.read_csv('/Users/chengyizhou/pioneer/final project/Meshva_dataset.csv')
print(f"Columns: {df.columns.tolist()}")
TARGET = 'generated_power_kw'
DROP_COLS = [
    'wind_speed_80_m_above_gnd',
    'wind_direction_80_m_above_gnd',
    'wind_speed_900_mb',
    'wind_direction_900_mb',
    'wind_gust_10_m_above_gnd'
]
df = df.drop(columns=DROP_COLS)

y = df[TARGET]
X = df.drop(columns=[TARGET])

X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
scaler = StandardScaler()
X_train = scaler.fit_transform(X_train)
X_test = scaler.transform(X_test)

features_num = X_train.shape[1]
print(f"Number of features: {features_num}")


# %%
model = keras.Sequential([
    layers.Input(shape=(features_num,)),
    layers.Dense(32, activation="relu"),
    layers.Dense(16, activation="relu"),
    layers.Dense(1),
])

model.compile(
    optimizer=keras.optimizers.Adam(learning_rate=0.001),
    loss="mse",
    metrics=[keras.metrics.RootMeanSquaredError(name="rmse")],
)

history = model.fit(X_train, y_train, validation_split=0.2, epochs=200, batch_size=32)

# %%
y_true = y_test                          # already kW
y_pred = model.predict(X_test).flatten() # already kW

rmse = np.sqrt(mean_squared_error(y_true, y_pred))
mae  = mean_absolute_error(y_true, y_pred)
r2   = r2_score(y_true, y_pred)
mbe  = np.mean(y_pred - y_true)

print(f"RMSE: {rmse:.2f} kW")
print(f"MAE:  {mae:.2f} kW")
print(f"R²:   {r2:.4f}")
print(f"MBE:  {mbe:.2f} kW")


