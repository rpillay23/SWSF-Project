import pandas as pd
import matplotlib.pyplot as plt

# Sample data simulating the "edited" DataFrame
data = {
    "Liquidity (1-10)": [3, 7, 6, 5, 8],
    "Volatility (1-10)": [6, 5, 7, 4, 8],
    "Investment Name": ["Inv A", "Inv B", "Inv C", "Inv D", "Inv E"]
}
edited = pd.DataFrame(data)

# Plot Liquidity vs Volatility
fig, ax = plt.subplots(figsize=(4, 3))
ax.scatter(edited["Liquidity (1-10)"], edited["Volatility (1-10)"], color="red", alpha=0.6)
ax.set_xlabel("Liquidity (1-10)", fontsize=10)
ax.set_ylabel("Volatility (1-10)", fontsize=10)
ax.set_title("Liquidity vs Volatility")
ax.grid(True)
fig.tight_layout(
