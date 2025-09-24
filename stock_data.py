import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl.utils import get_column_letter

# 1. Fetch historical stock price data
def fetch_data(tickers, start, end):
    data = yf.download(
        tickers,
        start=start,
        end=end,
        auto_adjust=False
    )
    return data['Adj Close']

# 2. Compute daily returns
def calculate_daily_returns(price_data):
    return price_data.pct_change().dropna()

# 3. Compute rolling volatility (default: 30-day window)
def calculate_volatility(daily_returns, window=30):
    return daily_returns.rolling(window).std()

# 4. Plot adjusted closing prices
def plot_prices(prices, title):
    prices.plot(figsize=(10, 6))
    plt.title(title)
    plt.xlabel("Date")
    plt.ylabel("Adjusted Closing Price")
    plt.legend(title="Ticker")
    plt.show()

# 5. Plot rolling volatility
def plot_volatility(volatility_data, title):
    volatility_data.plot(figsize=(10, 6))
    plt.title(title)
    plt.xlabel("Date")
    plt.ylabel("Volatility")
    plt.legend(title="Ticker")
    plt.show()

# 6. Summary statistics (mean, median, std) + boxplot
def summarize_statistics(daily_returns):
    summary = pd.DataFrame({
        "Mean": daily_returns.mean(),
        "Median": daily_returns.median(),
        "Std Dev": daily_returns.std()
    })
    print(summary)

    # Boxplot of daily returns
    plt.figure(figsize=(8, 6))
    sns.boxplot(data=daily_returns)
    plt.title("Boxplot of Daily Returns")
    plt.ylabel("Daily Return")
    plt.grid(True, linestyle="--", alpha=0.6)
    plt.show()
    return summary

# 7. Correlation heatmap
def plot_correlation_heatmap(daily_returns):
    corr = daily_returns.corr()
    plt.figure(figsize=(8, 6))
    sns.heatmap(
        corr,
        annot=True,
        cmap="coolwarm",
        fmt=".2f",
        linewidths=0.5
    )
    plt.title("Correlation Heatmap of Daily Returns")
    plt.show()

# 8. Portfolio performance
def calculate_portfolio_performance(daily_returns, weights=None):
    if weights is None:
        weights = [1 / len(daily_returns.columns)] * len(daily_returns.columns)

    portfolio_returns = (daily_returns * weights).sum(axis=1)

    cumulative_portfolio = (1 + portfolio_returns).cumprod()
    cumulative_stocks = (1 + daily_returns).cumprod()

    return portfolio_returns, cumulative_portfolio, cumulative_stocks

# Plot portfolio vs. individual stock performance
def plot_portfolio_vs_stocks(cumulative_portfolio, cumulative_stocks):
    plt.figure(figsize=(10, 6))
    plt.plot(cumulative_portfolio, label="Portfolio", linewidth=2, color="black")
    for col in cumulative_stocks.columns:
        plt.plot(cumulative_stocks[col], label=col, linestyle="--")
    plt.title("Portfolio vs Individual Stocks (Cumulative Returns)")
    plt.xlabel("Date")
    plt.ylabel("Cumulative Return")
    plt.legend()
    plt.grid(True, linestyle="--", alpha=0.6)
    plt.show()

# 9. Sharpe ratio calculation
def calculate_sharpe_ratio(returns, risk_free_rate=0.02 / 252):
    mean_return = returns.mean()
    std_dev = returns.std()
    sharpe_ratios = (mean_return - risk_free_rate) / std_dev
    return sharpe_ratios

# Plot Sharpe ratios (stocks vs. portfolio)
def plot_sharpe_ratios(sharpe_ratios, portfolio_sharpe):
    all_ratios = sharpe_ratios.copy()
    all_ratios = pd.concat([all_ratios, pd.Series({"Portfolio": portfolio_sharpe})])

    plt.figure(figsize=(8, 6))
    ax = all_ratios.plot(kind="bar", color="skyblue", edgecolor="black")
    plt.title("Sharpe Ratios: Stocks vs Portfolio")
    plt.ylabel("Sharpe Ratio")
    plt.xticks(rotation=0)
    plt.grid(axis="y", linestyle="--", alpha=0.6)

    # Annotate bar values
    for p in ax.patches:
        value = p.get_height()
        ax.annotate(
            f"{value:.3f}",
            (p.get_x() + p.get_width() / 2., value),
            ha="center", va="bottom", fontsize=9, color="black",
            xytext=(0, 3),
            textcoords="offset points"
        )
    plt.show()

# 10. Export to Excel File
def export_to_excel(
    daily_returns, summary, sharpe_ratios,
    portfolio_returns, cumulative_portfolio,
    cumulative_stocks, filename="stock_analysis.xlsx"
):
    with pd.ExcelWriter(filename, engine="openpyxl", date_format="YYYY-MM-DD") as writer:
        daily_returns.to_excel(writer, sheet_name="Daily Returns")
        summary.to_excel(writer, sheet_name="Summary Stats")
        sharpe_ratios.to_frame("Sharpe Ratio").to_excel(writer, sheet_name="Sharpe Ratios")
        portfolio_returns.to_frame("Portfolio Returns").to_excel(writer, sheet_name="Portfolio Returns")
        cumulative_portfolio.to_frame("Cumulative Portfolio").to_excel(writer, sheet_name="Cumulative Portfolio")
        cumulative_stocks.to_excel(writer, sheet_name="Cumulative Stocks")

        # Auto-adjust column widths for all sheets
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for col in worksheet.columns:
                max_length = 0
                column = get_column_letter(col[0].column)
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = max_length + 2
                worksheet.column_dimensions[column].width = adjusted_width

    print(f"✅ Data successfully exported to {filename} with formatted dates and adjusted columns")

# Main execution
if __name__ == "__main__":
    tickers = ["AAPL", "MSFT", "GOOGL"]
    start = "2020-01-01"
    end = "2025-07-30"

    # Fetch and preprocess data
    prices = fetch_data(tickers, start, end)
    daily_returns = calculate_daily_returns(prices)
    volatility = calculate_volatility(daily_returns)

    # Summary statistics
    summary = summarize_statistics(daily_returns)

    # Visualization
    plot_prices(prices, "Stock Prices (Adjusted Close)")
    plot_volatility(volatility, "Stock Volatility (30-Day Rolling)")
    plot_correlation_heatmap(daily_returns)

    # Portfolio analysis
    portfolio_returns, cumulative_portfolio, cumulative_stocks = calculate_portfolio_performance(daily_returns)
    plot_portfolio_vs_stocks(cumulative_portfolio, cumulative_stocks)

    # Sharpe ratio analysis
    sharpe_ratios = calculate_sharpe_ratio(daily_returns)
    print("\nSharpe Ratios (Individual Stocks):")
    print(sharpe_ratios)

    portfolio_sharpe = calculate_sharpe_ratio(portfolio_returns)
    print("\nSharpe Ratio (Portfolio):")
    print(portfolio_sharpe)

    plot_sharpe_ratios(sharpe_ratios, portfolio_sharpe)
    
    # Export all results
    export_to_excel(daily_returns, summary, sharpe_ratios, portfolio_returns, cumulative_portfolio, cumulative_stocks)