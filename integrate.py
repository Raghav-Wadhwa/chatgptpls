import os
import pandas as pd
import base64
from flask import Flask, request, jsonify
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings, TurnContext
from botbuilder.schema import Activity
from pandasai import SmartDataframe, SmartDatalake
from pandasai.llm.azure_openai import AzureOpenAI

# Initialize Flask app
app = Flask(__name__)

# Bot Framework Adapter (Azure Bot)
settings = BotFrameworkAdapterSettings(app_id="1967c66d-5394-4a33-b042-4f178dd521ae", app_password="YXq8Q~PABt92db6RufVOpE~sxgV_a4.VRihlVb6B")
bot_adapter = BotFrameworkAdapter(settings)

# Load dataset (Excel file)
csv_file = r"C:\Users\irt\Desktop\MS Teams Bot\data.xlsx"
df = pd.read_excel(csv_file)

# Preprocess dataset
for col in df.columns:
    if df[col].dtype == 'object':
        df[col] = df[col].astype(str).fillna("")
    else:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# Initialize Azure OpenAI LLM
llm = AzureOpenAI(
    api_token="9RaBd2y8Q5SWaCXK4yZeaR0cUwsQVYQduQgTDoy5SXFkHcf35QXEJQQJ99BBACYeBjFXJ3w3AAAAACOGBVjL",
    api_base="https://aiservices-eus-qlik-hzl-prod-spk-si.cognitiveservices.azure.com",
    api_version="2024-08-01-preview",
    deployment_name="eus-gpt-4",
)

# Create SmartDataframe and SmartDataLake
sdf = SmartDataframe(df, config={"llm": llm, "enable_cache": False, "save_charts": True})
data_lake = SmartDatalake([sdf], config={"llm": llm, "enable_cache": False, "save_charts": True})

# Function to process messages from Teams
async def process_message(turn_context: TurnContext):
    user_query = turn_context.activity.text
    response = data_lake.chat(user_query)

    # Handle different response types
    if isinstance(response, str):
        await turn_context.send_activity(response)
    elif isinstance(response, pd.DataFrame):
        df_string = response.to_string(index=False)
        await turn_context.send_activity(f"Here are the results:\n```\n{df_string}\n```")
    elif isinstance(response, list) or isinstance(response, dict):
        await turn_context.send_activity(str(response))
    else:
        await turn_context.send_activity("Unexpected response format.")

# Flask endpoint to receive messages
@app.route("/api/messages", methods=["POST"])
@app.route("/api/messages", methods=["POST"])
def messages():
    body = request.json
    activity = Activity().deserialize(body)

    # Ensure correct serviceUrl
    if not activity.service_url:
        activity.service_url = "https://6865-146-196-38-231.ngrok-free.app"

    async def turn_call(turn_context):
        await process_message(turn_context)

    task = bot_adapter.process_activity(activity, "", turn_call)
    return jsonify({"status": "ok"}), 200


if __name__ == "__main__":
    app.run(port=3978, debug=True)
