import requests

url = "https://api.engineeringvillage.com/EvDataWebServices/records"
apiKey = "bbcd5fe7831eb12082993dcbaaa6d72c"
inst_token = "4f3d2a4d46c51cbb68e83cf0b7150f45"
h = {"Accept":"application/json","X-ELS-APIKey":apiKey,"X-ELS-Insttoken":inst_token}

r = requests.get(url, headers=h, params={"docId":"cpx_M34b4eba21581b42d616M7bc510178163171"})
results = r.json()

print(results)
print(r.status_code)