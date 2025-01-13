
from flask import Flask

app = Flask(__name__)

@app.route('/test-azure')
def hello_world():
    return "Hello, World! Your Flask app is running on Azure."

if __name__ == '__main__':
    app.run(debug=True)
