import os
import re
import json
import asyncio
import pandas as pd
from datetime import datetime
import discord

DISCORD_TOKEN = os.environ.get('DISCORD_TOKEN')
DATE_FORMAT = '%Y-%m-%d'
JSON_FILE = 'data.json'

class MyClient(discord.Client):
    def __init__(self, row, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.row = row
        self.data_fetched = asyncio.Event()
        self.fetched_data = None

    async def on_ready(self):
        try:
            channel = self.get_channel(self.row[1])
            messages = await channel.history(limit=None).flatten()
        except Exception as e:
            print(f"Error fetching channel: {e}")
            self.data_fetched.set()
            await self.close()
            return

        records = [self.get_data(message) for message in messages if self.is_relevant_message(message) and self.is_regex(message)]
        df = pd.DataFrame(records, columns=['Date', 'Time', 'Author', 'Symbol', 'Expiry', 'Type', 'Strike', 'Entry'])
        print(f"{self.row[0]} was saved.")
        self.fetched_data = {self.row[0]: df}
        self.data_fetched.set()
        await self.close()
    
    def is_regex(self, message):
        content = message.embeds[0].description if self.row[9] == True else message.content
        search = re.search(self.row[4], re.sub("[0-9]+\.[0-9]+(?![cpCP])|\.[0-9]+(?![cpCP])", '', content))
        return search

    def is_relevant_message(self, message):
        return (
            (message.author.id == self.row[2] or self.row[3] == True) and
            (len(message.embeds) > 0 if self.row[9] == True else True)
        )

    def get_data(self, message):
        search = self.is_regex(message)
        date_posted = datetime.strptime(str(message.created_at)[:10], DATE_FORMAT)
        date_expiry = self.get_date_expiry(search.group(0), date_posted)
        try:
            return [
                str(message.created_at)[:10],
                str(message.created_at)[11:],
                message.author,
                self.get_symbol(search.group(0)),
                datetime.strftime(date_expiry, '%d%b%y') if date_expiry else '',
                re.search(self.row[7], search.group(0)).group(0).upper(),
                "{:.2f}".format(float(re.search(self.row[8], search.group(0)).group(0))),
                '-'
            ]
        except Exception as e:
            return [
                str(message.created_at)[:10],
                str(message.created_at)[11:],
                message.author,
                self.get_symbol(search.group(0)),
                datetime.strftime(date_expiry, '%d%b%y') if date_expiry else '',
                None,
                None,
                '-'
            ]

    def get_date_expiry(self, search_group, date_posted):
        try:
            date_expiry = re.search(self.row[5], search_group).group(0)
            date_expiry = datetime.strptime(date_expiry, '%m/%d').replace(year=date_posted.year)
            return date_expiry
        except Exception as e:
            return None

    def get_symbol(self, search_group):
        try:
            symbol = re.search(self.row[6], search_group.replace('BTO', '')).group(0).upper()
            return 'SPXW' if symbol == 'SPX' else symbol
        except:
            return '-'

async def run_client(token, row):
    client = MyClient(row)
    await client.start(token, bot=False)
    await client.data_fetched.wait()
    return client.fetched_data

async def main():
    with open(JSON_FILE, 'r') as file:
        data = json.load(file)

    sheets = {}
    for row in data:
        await asyncio.sleep(5)
        row_as_tuple = tuple(row.values())
        sheet = await run_client(DISCORD_TOKEN, row_as_tuple)
        if sheet:
            sheets.update(sheet)
    
    writer = pd.ExcelWriter('trades.xlsx')
    for sheet_name, sheet_df in sheets.items():
        sheet_df.to_excel(writer, sheet_name, index=False)
    writer.save()

if __name__ == '__main__':
    asyncio.run(main())