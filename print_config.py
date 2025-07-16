import json

with open("config_clients.json", "r", encoding="utf-8") as f:
    data = json.load(f)
print(data)
