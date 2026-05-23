from slack_sdk import WebClient

# Replace with your Bot User OAuth Token
SLACK_TOKEN = "xoxb-your-slack-bot-token"

# Replace with your channel ID
CHANNEL_ID = "C08LEFNF0BS"

client = WebClient(token=SLACK_TOKEN)

response = client.conversations_history(
    channel=CHANNEL_ID,
    limit=10
)

for message in response["messages"]:
    print(message["text"])