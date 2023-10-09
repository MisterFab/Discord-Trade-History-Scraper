# Discord Trade History Scraper

This repository houses a comprehensive tool designed for traders keen on extracting and analyzing option trade data from Discord users that provide trade signals. It not only identifies relevant trades from messages but also fetches historical trade data using the Polygon API. The consolidated results, inclusive of historical highs, lows, and other pertinent data, are saved into Excel spreadsheets for ease of analysis and auditing.

## Features

- Discord Data Extraction: Efficiently fetches all messages from designated Discord channels based on custom criteria
- Data Filtering: Incorporates regular expressions for precise filtering and processing of pertinent trading data from Discord messages.
- Historical Data Retrieval: Taps into the Polygon API to amass historical price data for recognized trades.
- Excel Integration: Meticulously processes the data and channels the results into well-structured Excel sheets.

## Prerequisites

- discord.py==1.7.3
- A Discord bot token and permissions to read messages from the desired channels.
- Polygon API token for fetching historical trade data.

## Usage

- Configure the data.json file with relevant parameters for filtering and extracting trade data from your Discord channels. I suggest https://regex101.com/ to test your regular expressions.

## Contributing

If you'd like to contribute to the project, feel free to open an issue or submit a pull request.