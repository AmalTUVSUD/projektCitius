import json

def load_keywords(path="keywords.json"):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)