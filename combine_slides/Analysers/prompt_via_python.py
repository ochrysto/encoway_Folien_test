import ollama

# llama3 muss zunächst gestartet werden
# Eingabeaufforderung öffnen
# ubuntu eingeben
# ollama run llama3 eingeben

# Verbindung zur lokalen Instanz herstellen
client = ollama.Client(host='http://localhost:11434')

# Prompt senden und Antwort erhalten
prompt = """Kann
  """
response = client.generate(model='llama3', prompt=prompt)
print(response['response'])

# um llama3 zu beenden (Eingabeaufforderung)
# strg + d eingeben
# exit eingeben