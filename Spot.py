import requests
import pandas as pd
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import warnings
from urllib3.exceptions import InsecureRequestWarning
import plotly.express as px

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

# Main function
def main():
    # Determine URL for downloading XLS file
    today = datetime.now() + timedelta(days=1)  # Adding one day to the current date
    url = f"https://www.ote-cr.cz/pubweb/attachments/01/{today.year}/month{today.month:02d}/day{today.day:02d}/DT_{today.day:02d}_{today.month:02d}_{today.year}_CZ.xls"

    # Download XLS file
    download_xls(url)
    # print("XLS file was successfully downloaded.")

    # Process XLS file
    dataframe = process_xls("data.xls")
    # print("XLS file was successfully processed.")

    # Get current exchange rate
    exchange_rate = get_exchange_rate()

    # Convert values from EUR to CZK and round to 2 decimal places
    dataframe = convert_to_czk(dataframe, exchange_rate)

    print("cena spotu k " + str(today.strftime('%e.%#m.%Y')))
    # print(dataframe)

    # Plotting
    fig = px.bar(dataframe, y="Value_CZK", title=f"Spotové ceny v Česku k: {today.strftime('%e.%#m.%Y')} za jednu kWh.", labels={"Value_CZK": "Cena v Kč"}, color="Value_CZK", color_continuous_scale='RdYlGn')  # bar chart
    fig.update_xaxes(title_text="Hodina")  # Set the x-axis title
    fig.update_traces(text=[f"{price} Kč" for price in dataframe["Value_CZK"].values])  # Show prices directly on the bars with "Kč"
    fig.show()

if __name__ == "__main__":
    main()
