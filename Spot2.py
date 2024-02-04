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
    download_xls(url)
    dataframe = process_xls("data.xls")
    return dataframe

# Main function
def main():
    # Get current exchange rate
    exchange_rate = get_exchange_rate()

    # Get current date and next day
    today = datetime.now()
    next_day = today + timedelta(days=1)

    # Process data for both days
    current_day_data = process_day_data(today)
    next_day_data = process_day_data(next_day)

    # Convert values from EUR to CZK for both days
    current_day_data = convert_to_czk(current_day_data, exchange_rate)
    next_day_data = convert_to_czk(next_day_data, exchange_rate)

    # Combine dataframes
    combined_data = pd.concat([current_day_data, next_day_data], axis=0)

    # Create figure
    fig = go.Figure()

    # Add traces for both days
    fig.add_trace(go.Bar(x=current_day_data.index, y=current_day_data["Value_CZK"], name="Aktuální den",
                         text=current_day_data["Value_CZK"], textposition="outside", texttemplate="%{text:.2f} Kč",
                         marker_color='mediumblue'))
    fig.add_trace(go.Bar(x=next_day_data.index, y=next_day_data["Value_CZK"], name="Následující den",
                         text=next_day_data["Value_CZK"], textposition="outside", texttemplate="%{text:.2f} Kč",
                         marker_color='firebrick'))

    # Change color for values below 0.4 Kč
    current_day_colors = ['yellow' if val < 0.4 else 'mediumblue' for val in current_day_data["Value_CZK"]]
    next_day_colors = ['yellow' if val < 0.4 else 'firebrick' for val in next_day_data["Value_CZK"]]

    fig.update_traces(marker_color=current_day_colors, selector=dict(type='bar', name="Aktuální den"))
    fig.update_traces(marker_color=next_day_colors, selector=dict(type='bar', name="Následující den"))

    # Update layout
    fig.update_layout(title="Spotové ceny elektřiny v Česku",
                      title_font=dict(size=24),
                      xaxis_title="Hodina", yaxis_title="Cena v Kč", barmode="group",
                      legend_title="Legenda",
                      legend=dict(x=0.01, y=0.95, orientation="h", xanchor="left", yanchor="middle"),
                      font=dict(size=18))

    # Add dates to legend
    fig.data[0].name = f"{fig.data[0].name} ({today.strftime('%e.%#m.%Y')})"
    fig.data[1].name = f"{fig.data[1].name} ({next_day.strftime('%e.%#m.%Y')})"

    # Show plot
    fig.show()

if __name__ == "__main__":
    main()
