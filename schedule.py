from flask import Blueprint
from flask import session

second = Blueprint('schedule', __name__)

@second.route('/')
def home():
    names = session['names']
    return str(names)