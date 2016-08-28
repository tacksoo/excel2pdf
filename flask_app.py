# A very simple Flask Hello World app for you to get started with...

from flask import Flask
from flask import request
from flask import jsonify

app = Flask(__name__)

@app.route('/')
def hello_world():
  #jsonify({'ip': request.remote_addr}), 200
  return "Hello World"

@app.route('/dude')
def dude():
  return "<h1>Dude!</h1>"

