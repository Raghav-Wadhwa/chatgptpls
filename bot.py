from botbuilder.core import ActivityHandler, TurnContext
from botbuilder.schema import ChannelAccount
from pandasai import SmartDataframe, SmartDatalake
from pandasai.llm.azure_openai import AzureOpenAI
import pandas as pd

# Initialize Azure OpenAI LLM
llm = AzureOpenAI(
    api_token="YOUR_API_TOKEN",
    api_base="YOUR_API_BASE",
    api_version="2024-08-01-preview",
    deployment_name="eus-gpt-4",
)

# Load dataset (Excel file)
csv_file = "path_to_your_dataset.xlsx"
df = pd.read_excel(csv_file)
for col in df.columns:
    if df[col].dtype == 'object':
        df[col] = df[col].astype(str).fillna("")
    else:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

sdf = SmartDataframe(df, config={"llm": llm, "enable_cache": False, "save_charts": True})
data_lake = SmartDatalake([sdf], config={"llm": llm, "enable_cache": False, "save_charts": True})

class MyBot(ActivityHandler):
    async def on_message_activity(self, turn_context: TurnContext):
        user_query = turn_context.activity.text
        response = data_lake.chat(user_query)

        if isinstance(response, str):
            await turn_context.send_activity(response)
        elif isinstance(response, pd.DataFrame):
            df_string = response.to_string(index=False)
            await turn_context.send_activity(f"Here are the results:\n```\n{df_string}\n```")
        elif isinstance(response, list) or isinstance(response, dict):
            await turn_context.send_activity(str(response))
        else:
            await turn_context.send_activity("Unexpected response format.")

    async def on_members_added_activity(
            self, members_added, turn_context: TurnContext):
        for member_added in members_added:
            if member_added.id != turn_context.activity.recipient.id:
                await turn_context.send_activity("Hello and welcome!")
