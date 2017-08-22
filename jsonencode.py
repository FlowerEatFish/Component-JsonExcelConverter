import json

class JsonEncoder():
    def __init__(self, data):
        if type(data) == dict:
            return json.dumps(data)
        print('Failure: The data you enter isn\'t dict type.')

class JsonDecoder():
    def __init__(self, data):
        if type(data) == str:
            try:
                return json.loads(data)
            except:
                print('Failure: Your data isn\'t json type.')
