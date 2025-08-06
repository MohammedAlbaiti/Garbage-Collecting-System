from telethon import TelegramClient
import pandas as pd

# Get these from my.telegram.org
API_ID = 21224562  # Replace with your API ID
API_HASH = "1980c320cbec56efc00b4bbbf75985f1"

# Initialize the client (but don't start it yet)
client = TelegramClient("session_name", API_ID, API_HASH)

async def store_message_in_excel(message):
    """
    Parse the message and store each bin address + value in an Excel file.
    Expected format:
        Bin A - Street Name: 74
        Bin B - Some Road: 22
    """
    lines = message.strip().split('\n')
    data = []
    readed_address = []

    for line in lines:
        if ':' in line:
            address, value = line.split(':', 1)
            if address in readed_address:
                continue
            else:
                readed_address.append(address)
            try:
                percent = int(value.strip())
                data.append({"Address": address.strip(), "Message": percent})
            except ValueError:
                print(f"Invalid number in line: {line}")

    df = pd.DataFrame(data)
    df.to_excel("latest_message.xlsx", index=False)


async def get_latest_messages():
    """Get the latest message from the Telegram chat."""
    chat_id = -1002046623038  # Replace with your chat/group ID
    messages = await client.get_messages(chat_id, limit=100)
    all_data = []  # Collect data from all valid messages
    readed_address = []
    for msg in messages:
        sender = await msg.get_sender()
        if sender is None or sender.first_name != "Smart garbage" or not msg.text or ":" not in msg.text:
            continue

        sender_name = sender.first_name if sender else "Unknown"
        is_bot = sender.bot if sender else False
        
        last_message = msg.text
        print(f"From: {sender_name} (Bot: {is_bot}) - Message: {last_message}")
        
        # Process lines and collect data
        lines = last_message.strip().split('\n')
        for line in lines:
            if ':' in line:
                address, value = line.split(':', 1)
                if address in readed_address:
                    continue
                else:
                    readed_address.append(address)
                try:
                    percent = int(value.strip())
                    all_data.append({"Address": address.strip(), "Message": percent})
                except ValueError:
                    print(f"Invalid number in line: {line}")

    # Write all collected data to Excel once
    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel("latest_message.xlsx", index=False)
    return last_message if messages else None
async def main():

    """Main function to start the client and get messages."""
    await client.start()  # Start the client session
    await get_latest_messages()  # Get the latest message

# Start the client in the background
with client:
    client.loop.run_until_complete(main())

