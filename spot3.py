import requests
import pandas as pd
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import warnings
from urllib3.exceptions import InsecureRequestWarning
import plotly.graph_objects as go

# Suppressing warnings
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)
warnings.filterwarnings("ignore", category=InsecureRequestWarning)

# Function to download XLS file
def download_xls(url):
    response = requests.get(url, verify=False)
    with open("data.xls", "wb") as f:
        f.write(response.content)

# Function to process XLS file
def process_xls(filename):
    df = pd.read_excel(filename, skiprows=6, nrows=24, usecols="B", header=None, names=["Value"])
    return df

# Function to get current exchange rate
def get_exchange_rate():
    url = "https://www.cnb.cz/cs/financni-trhy/devizovy-trh/kurzy-devizoveho-trhu/kurzy-devizoveho-trhu/"
    response = requests.get(url, verify=False)
    soup = BeautifulSoup(response.text, "html.parser")
    exchange_rate_text = soup.find("td", string="EUR").find_next_sibling("td").text
    exchange_rate = float(exchange_rate_text.replace(",", "."))
    return exchange_rate

# Function to convert values from EUR to CZK
def convert_to_czk(df, exchange_rate):
    df["Value_CZK"] = df["Value"] * exchange_rate / 1000  # divide by 1000
    df["Value_CZK"] = df["Value_CZK"].round(2)  # round to 2 decimal places
    df.index = df.index.astype(str) + ":00"  # convert index to hour format
    return df

# Function to process data for a given day
def process_day_data(date):
    url = f"https://www.ote-cr.cz/pubweb/attachments/01/{date.year}/month{date.month:02d}/day{date.day:02d}/DT_{date.day:02d}_{date.month:02d}_{date.year}_CZ.xls"
    try:
        download_xls(url)
        dataframe = process_xls("data.xls")
        return dataframe
    except Exception as e:
        print(f"Chyba při zpracování dat pro {date}: {e}")
        return None

# Main function
def main():
    # Get current exchange rate
    exchange_rate = get_exchange_rate()

    # Get current date and hour
    now = datetime.now()
    current_hour = now.hour

    # Process data for current day
    current_day_data = process_day_data(now)
    if current_day_data is None:
        print("Data pro aktuální den nejsou k dispozici.")
        return

    # Convert values from EUR to CZK for current day
    current_day_data = convert_to_czk(current_day_data, exchange_rate)

    # Create figure
    fig = go.Figure()

    # Add traces for current day
    fig.add_trace(go.Bar(x=current_day_data.index, y=current_day_data["Value_CZK"], name="Aktuální den",
                         text=current_day_data["Value_CZK"], textposition="outside", texttemplate="%{text:.2f} Kč",
                         marker_color=['yellow' if val < 0.4 else 'mediumblue' for val in current_day_data["Value_CZK"]],
                         # Highlight current hour
                         marker_line_color=['red' if i == current_hour else 'rgba(0,0,0,0)' for i in range(24)],
                         marker_line_width=3))

    # Update layout
    fig.update_layout(title=f"Spotové ceny elektřiny v Česku k {now.strftime('%d.%m.%Y')}",
                      title_font=dict(size=24),
                      xaxis_title="Hodina", yaxis_title="Cena v Kč", barmode="group",
                      legend_title="Legenda",
                      legend=dict(x=0.01, y=0.95, orientation="h", xanchor="left", yanchor="middle", bgcolor='rgba(255, 255, 255, 0.5)'),
                      font=dict(size=18))

    # Show plot
    fig.show()

if __name__ == "__main__":
    main()
