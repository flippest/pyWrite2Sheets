#!/usr/local/bin/python3

from flask import Flask, request
import apiWrite2Sheets

app = Flask(__name__)

@app.route('/write', methods=['POST'])
def write():
    data = request.json
    event = data['event']
    badgeid = data['badgeid']
    inout = data['inout']
    
    apiWrite2Sheets.main(event, badgeid, inout)
    
    return 'Success'

if __name__ == '__main__':
    app.run()
